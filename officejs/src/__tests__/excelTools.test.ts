// Unit tests for the agent's Excel tool implementations. The handlers call
// Excel.run against the live workbook, so we install a fake `Excel` global that
// records every range operation. This exercises the dispatch table, A1/sheet
// address parsing, 2D-value normalization, resize + formula-matrix construction,
// and formatting — none of which needed live Excel to be correct, only reachable.

import { executeExcelTool, WRITE_TOOLS, EXCEL_TOOLS, resolveChartType } from "../excelTools";

interface Preset {
  selectionAddress?: string;
  selectionValues?: any[][];
  sheetNames?: string[];
  rangeValues?: Record<string, any[][]>;
  rangeDims?: Record<string, [number, number]>;
}

let captured: {
  ranges: any[];
  resized: Array<{ from: string; rows: number; cols: number }>;
  addedSheet: string | null;
  charts: any[];
};

function installExcel(preset: Preset = {}) {
  captured = { ranges: [], resized: [], addedSheet: null, charts: [] };

  function makeRange(address: string): any {
    const dims = preset.rangeDims?.[address] ?? [1, 1];
    const rng: any = {
      address,
      values: preset.rangeValues?.[address] ?? [["v"]],
      rowCount: dims[0],
      columnCount: dims[1],
      format: { fill: {}, font: {} },
      load() {
        /* no-op: values are preset */
      },
      getAbsoluteResizedRange(rows: number, cols: number) {
        captured.resized.push({ from: address, rows, cols });
        return makeRange(`${address}#${rows}x${cols}`);
      },
    };
    captured.ranges.push(rng);
    return rng;
  }

  const makeCharts = () => ({
    add(type: string, range: any, seriesBy: string) {
      const chart: any = { name: "Chart 1", type, range, seriesBy, title: {}, load() {} };
      captured.charts.push(chart);
      return chart;
    },
  });

  const worksheets = {
    items: [] as any[],
    load() {
      this.items = (preset.sheetNames ?? ["Sheet1"]).map((name) => ({ name }));
    },
    getItem(name: string) {
      return { name, getRange: (a: string) => makeRange(`${name}!${a}`), charts: makeCharts() };
    },
    getActiveWorksheet() {
      return { name: "Active", getRange: (a: string) => makeRange(a), charts: makeCharts() };
    },
    add(name: string) {
      captured.addedSheet = name;
      return { name, load() {} };
    },
  };

  const ctx = {
    workbook: {
      worksheets,
      getSelectedRange() {
        const r = makeRange(preset.selectionAddress ?? "B2");
        r.values = preset.selectionValues ?? [["sel"]];
        return r;
      },
    },
    sync: async () => {},
  };

  (global as any).Excel = {
    run: async (cb: any) => cb(ctx),
    ChartSeriesBy: { auto: "Auto" },
  };
}

function rangeAt(address: string) {
  return captured.ranges.find((r) => r.address === address);
}

describe("executeExcelTool", () => {
  test("read_range parses Sheet!A1 addresses and returns values", async () => {
    installExcel({ rangeValues: { "Sheet1!A1:B2": [[1, 2], [3, 4]] } });
    const out = await executeExcelTool("read_range", { address: "Sheet1!A1:B2" });
    expect(out).toBe("Sheet1!A1:B2 = [[1,2],[3,4]]");
  });

  test("read_range strips single quotes around a sheet name with spaces", async () => {
    installExcel({ rangeValues: { "My Sheet!A1": [[9]] } });
    const out = await executeExcelTool("read_range", { address: "'My Sheet'!A1" });
    expect(out).toBe("My Sheet!A1 = [[9]]");
  });

  test("get_selection reports address and values", async () => {
    installExcel({ selectionAddress: "B2:B3", selectionValues: [[10], [20]] });
    const out = await executeExcelTool("get_selection", {});
    expect(out).toBe("Selection B2:B3 = [[10],[20]]");
  });

  test("list_sheets joins worksheet names", async () => {
    installExcel({ sheetNames: ["Sheet1", "Data"] });
    expect(await executeExcelTool("list_sheets", {})).toBe("Sheets: Sheet1, Data");
  });

  test("write_range resizes to the value dimensions and writes them", async () => {
    installExcel();
    const out = await executeExcelTool("write_range", { address: "A1", values: [[1, 2], [3, 4]] });
    expect(captured.resized).toContainEqual({ from: "A1", rows: 2, cols: 2 });
    expect(rangeAt("A1#2x2").values).toEqual([[1, 2], [3, 4]]);
    expect(out).toBe("Wrote 2x2 to A1#2x2.");
  });

  test("write_range normalizes a flat array to a single row", async () => {
    installExcel();
    await executeExcelTool("write_range", { address: "A1", values: [1, 2, 3] });
    expect(captured.resized).toContainEqual({ from: "A1", rows: 1, cols: 3 });
    expect(rangeAt("A1#1x3").values).toEqual([[1, 2, 3]]);
  });

  test("write_range normalizes a scalar to a single cell", async () => {
    installExcel();
    await executeExcelTool("write_range", { address: "A1", values: 42 });
    expect(rangeAt("A1#1x1").values).toEqual([[42]]);
  });

  test("write_range refuses an empty value set instead of writing garbage", async () => {
    installExcel();
    const out = await executeExcelTool("write_range", { address: "A1", values: [] });
    expect(out).toBe("No values to write.");
    expect(captured.resized).toHaveLength(0);
  });

  test("write_formula fills a formula matrix over the target range", async () => {
    installExcel({ rangeDims: { "C1:C3": [3, 1] } });
    const out = await executeExcelTool("write_formula", { address: "C1:C3", formula: "=A1*B1" });
    expect(rangeAt("C1:C3").formulas).toEqual([["=A1*B1"], ["=A1*B1"], ["=A1*B1"]]);
    expect(out).toBe("Set formula =A1*B1 on C1:C3.");
  });

  test("set_format applies fill, font, and a number-format matrix", async () => {
    installExcel({ rangeDims: { A1: [1, 2] } });
    const out = await executeExcelTool("set_format", {
      address: "A1",
      fill: "#FFEB9C",
      bold: true,
      italic: false,
      numberFormat: "0.00",
    });
    const r = rangeAt("A1");
    expect(r.format.fill.color).toBe("#FFEB9C");
    expect(r.format.font.bold).toBe(true);
    expect(r.format.font.italic).toBe(false);
    expect(r.numberFormat).toEqual([["0.00", "0.00"]]);
    expect(out).toBe("Formatted A1.");
  });

  test("add_worksheet creates a sheet by name", async () => {
    installExcel();
    const out = await executeExcelTool("add_worksheet", { name: "Results" });
    expect(captured.addedSheet).toBe("Results");
    expect(out).toBe('Added worksheet "Results".');
  });

  test("create_chart adds a chart of the resolved type from the data range", async () => {
    installExcel();
    const out = await executeExcelTool("create_chart", {
      address: "A1:B10",
      chartType: "bar",
      title: "Sales",
    });
    expect(captured.charts).toHaveLength(1);
    expect(captured.charts[0].type).toBe("BarClustered");
    expect(captured.charts[0].title.text).toBe("Sales");
    expect(out).toBe('Created BarClustered chart "Sales" from A1:B10.');
  });

  test("create_chart falls back to a column chart for an unknown type", async () => {
    installExcel();
    await executeExcelTool("create_chart", { address: "A1:B10", chartType: "wobble" });
    expect(captured.charts[0].type).toBe("ColumnClustered");
  });

  test("unknown tool name returns a diagnostic instead of throwing", async () => {
    installExcel();
    expect(await executeExcelTool("bogus", {})).toBe("Unknown tool: bogus");
  });
});

describe("resolveChartType", () => {
  test("maps synonyms and normalizes punctuation/casing", () => {
    expect(resolveChartType("bar")).toBe("BarClustered");
    expect(resolveChartType("Column Clustered")).toBe("ColumnClustered");
    expect(resolveChartType("PIE")).toBe("Pie");
    expect(resolveChartType("scatter")).toBe("XYScatter");
    expect(resolveChartType("nonsense")).toBe("ColumnClustered"); // safe default
  });
});

describe("tool metadata", () => {
  test("WRITE_TOOLS covers exactly the mutating tools", () => {
    expect([...WRITE_TOOLS].sort()).toEqual([
      "add_worksheet",
      "create_chart",
      "set_format",
      "write_formula",
      "write_range",
    ]);
    expect(WRITE_TOOLS.has("read_range")).toBe(false);
  });

  test("every WRITE_TOOL has a matching schema in EXCEL_TOOLS", () => {
    const names = new Set(EXCEL_TOOLS.map((t) => t.name));
    for (const w of WRITE_TOOLS) expect(names.has(w)).toBe(true);
  });
});
