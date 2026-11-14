// Excel tool implementations for the agent (browser/Office.js). Each tool has a
// JSON-schema for the model and a handler that runs via Excel.run against the
// live workbook. Not unit-tested (needs Excel); the agent loop that drives them
// is tested in core/__tests__/agent.test.ts.

import { ToolSchema } from "./core/agent";
import {
  SteeringOp,
  SteeringTask,
  columnWidthToPoints,
  normalizeSteeringTask,
  rowHeightToPoints,
} from "./core/steering";

/* global Excel */

export const EXCEL_TOOLS: ToolSchema[] = [
  {
    name: "read_range",
    description: "Read the values of a range in A1 notation (e.g. B2:B10 or Sheet1!A1:C5).",
    parameters: {
      type: "object",
      properties: { address: { type: "string", description: "A1-style range address" } },
      required: ["address"],
    },
  },
  {
    name: "get_selection",
    description: "Get the address and values of the user's currently selected range.",
    parameters: { type: "object", properties: {} },
  },
  {
    name: "list_sheets",
    description: "List the worksheet names in the workbook.",
    parameters: { type: "object", properties: {} },
  },
  {
    name: "write_range",
    description: "Write a 2D array of values starting at a top-left cell/range (resizes to fit).",
    parameters: {
      type: "object",
      properties: {
        address: { type: "string" },
        values: { type: "array", description: "2D array of cell values", items: { type: "array" } },
      },
      required: ["address", "values"],
    },
  },
  {
    name: "write_formula",
    description: "Fill a range with a formula (relative references adjust per cell), e.g. =A2*B2.",
    parameters: {
      type: "object",
      properties: { address: { type: "string" }, formula: { type: "string" } },
      required: ["address", "formula"],
    },
  },
  {
    name: "set_format",
    description: "Format a range: fill color, font color, bold, italic, or number format.",
    parameters: {
      type: "object",
      properties: {
        address: { type: "string" },
        fill: { type: "string", description: "background hex, e.g. #FFEB9C" },
        fontColor: { type: "string" },
        bold: { type: "boolean" },
        italic: { type: "boolean" },
        numberFormat: { type: "string", description: "e.g. 0.00 or $#,##0" },
      },
      required: ["address"],
    },
  },
  {
    name: "add_worksheet",
    description: "Add a new worksheet with the given name.",
    parameters: {
      type: "object",
      properties: { name: { type: "string" } },
      required: ["name"],
    },
  },
  {
    name: "create_chart",
    description: "Create a chart from a data range (e.g. bar chart of B2:C20). Optional title.",
    parameters: {
      type: "object",
      properties: {
        address: { type: "string", description: "A1 range holding the chart data (incl. headers)" },
        chartType: {
          type: "string",
          description: "One of: ColumnClustered, BarClustered, Line, Pie, XYScatter, Area",
        },
        title: { type: "string" },
      },
      required: ["address", "chartType"],
    },
  },
  {
    name: "set_number_format",
    description: "Apply an Excel number format to a range, e.g. #,##0.00, 0.0%, DD.MM.YYYY.",
    parameters: {
      type: "object",
      properties: {
        address: { type: "string", description: "A1 cell/range, optionally Sheet1!A1:B2" },
        format: { type: "string", description: "Excel format code" },
      },
      required: ["address", "format"],
    },
  },
  {
    name: "merge_cells",
    description: "Merge a range into one cell; optionally centre the result (default true).",
    parameters: {
      type: "object",
      properties: {
        address: { type: "string" },
        center: { type: "boolean" },
      },
      required: ["address"],
    },
  },
  {
    name: "unmerge_cells",
    description: "Split a merged range back into individual cells.",
    parameters: {
      type: "object",
      properties: { address: { type: "string" } },
      required: ["address"],
    },
  },
  {
    name: "set_column_width",
    description: "Set a column width. Units: chars (default), points, or px.",
    parameters: {
      type: "object",
      properties: {
        column: { type: "string", description: "Column letter, e.g. B" },
        width: { type: "number" },
        unit: { type: "string", description: "chars | points | px" },
        sheet: { type: "string" },
      },
      required: ["column", "width"],
    },
  },
  {
    name: "set_row_height",
    description: "Set a row height in points (default) or px.",
    parameters: {
      type: "object",
      properties: {
        row: { type: "number", description: "1-based row number" },
        height: { type: "number" },
        unit: { type: "string", description: "points | px" },
        sheet: { type: "string" },
      },
      required: ["row", "height"],
    },
  },
  {
    name: "recalculate",
    description: "Force a full workbook recalculation.",
    parameters: { type: "object", properties: {} },
  },
  {
    name: "rename_worksheet",
    description: "Rename an existing worksheet.",
    parameters: {
      type: "object",
      properties: { from: { type: "string" }, to: { type: "string" } },
      required: ["from", "to"],
    },
  },
  {
    name: "run_steering",
    description:
      "Run a portable Excel steering task (the same JSON file excel-steer uses): " +
      "{ops:[{op:set_value|set_formula|format|merge|unmerge|column_width|row_height|add_sheet|rename_sheet|recalculate|assert, ...}]}. " +
      "Use it to replay a saved recipe exactly instead of issuing many separate edits.",
    parameters: {
      type: "object",
      properties: {
        task: {
          description: "The task object, or its JSON text",
          type: "object",
        },
      },
      required: ["task"],
    },
  },
];

/** Tools that change the workbook (vs. read-only). Used for approve-before-apply. */
export const WRITE_TOOLS = new Set([
  "write_range",
  "write_formula",
  "set_format",
  "add_worksheet",
  "create_chart",
  "set_number_format",
  "merge_cells",
  "unmerge_cells",
  "set_column_width",
  "set_row_height",
  "recalculate",
  "rename_worksheet",
  "run_steering",
]);

// Chart-type names the model may use → Excel.ChartType. Kept permissive so common
// synonyms (bar/column/scatter) still resolve.
const CHART_TYPES: Record<string, string> = {
  columnclustered: "ColumnClustered",
  column: "ColumnClustered",
  bar: "BarClustered",
  barclustered: "BarClustered",
  line: "Line",
  pie: "Pie",
  scatter: "XYScatter",
  xyscatter: "XYScatter",
  area: "Area",
};

export async function executeExcelTool(name: string, args: any): Promise<string> {
  switch (name) {
    case "read_range":
      return readRange(String(args.address));
    case "get_selection":
      return getSelection();
    case "list_sheets":
      return listSheets();
    case "write_range":
      return writeRange(String(args.address), to2D(args.values));
    case "write_formula":
      return writeFormula(String(args.address), String(args.formula));
    case "set_format":
      return setFormat(args);
    case "add_worksheet":
      return addWorksheet(String(args.name));
    case "create_chart":
      return createChart(String(args.address), String(args.chartType), args.title ? String(args.title) : "");
    case "set_number_format":
      return setNumberFormat(String(args.address), String(args.format));
    case "merge_cells":
      return mergeCells(String(args.address), args.center !== false);
    case "unmerge_cells":
      return unmergeCells(String(args.address));
    case "set_column_width":
      return setColumnWidth(String(args.column), Number(args.width), args.unit, args.sheet);
    case "set_row_height":
      return setRowHeight(Number(args.row), Number(args.height), args.unit, args.sheet);
    case "recalculate":
      return recalculateWorkbook();
    case "rename_worksheet":
      return renameWorksheet(String(args.from), String(args.to));
    case "run_steering":
      return runSteering(args.task);
    default:
      return `Unknown tool: ${name}`;
  }
}

function rangeFromAddress(ctx: Excel.RequestContext, address: string): Excel.Range {
  const bang = address.indexOf("!");
  if (bang >= 0) {
    const sheet = address.slice(0, bang).replace(/^'|'$/g, "");
    return ctx.workbook.worksheets.getItem(sheet).getRange(address.slice(bang + 1));
  }
  return ctx.workbook.worksheets.getActiveWorksheet().getRange(address);
}

function to2D(values: any): any[][] {
  if (!Array.isArray(values)) return [[values]];
  if (values.length === 0) return [[]];
  return Array.isArray(values[0]) ? values : [values];
}

async function readRange(address: string): Promise<string> {
  return Excel.run(async (ctx) => {
    const rng = rangeFromAddress(ctx, address);
    rng.load("values,address");
    await ctx.sync();
    return `${rng.address} = ${JSON.stringify(rng.values)}`;
  });
}

async function getSelection(): Promise<string> {
  return Excel.run(async (ctx) => {
    const rng = ctx.workbook.getSelectedRange();
    rng.load("address,values");
    await ctx.sync();
    return `Selection ${rng.address} = ${JSON.stringify(rng.values)}`;
  });
}

async function listSheets(): Promise<string> {
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();
    return `Sheets: ${sheets.items.map((s) => s.name).join(", ")}`;
  });
}

async function writeRange(address: string, values: any[][]): Promise<string> {
  return Excel.run(async (ctx) => {
    const rows = values.length;
    const cols = values[0]?.length ?? 0;
    if (rows === 0 || cols === 0) return "No values to write.";
    const target = rangeFromAddress(ctx, address).getAbsoluteResizedRange(rows, cols);
    target.values = values;
    target.load("address");
    await ctx.sync();
    return `Wrote ${rows}x${cols} to ${target.address}.`;
  });
}

async function writeFormula(address: string, formula: string): Promise<string> {
  return Excel.run(async (ctx) => {
    const rng = rangeFromAddress(ctx, address);
    rng.load("rowCount,columnCount");
    await ctx.sync();
    const f = Array.from({ length: rng.rowCount }, () => Array.from({ length: rng.columnCount }, () => formula));
    rng.formulas = f;
    rng.load("address");
    await ctx.sync();
    return `Set formula ${formula} on ${rng.address}.`;
  });
}

async function setFormat(args: any): Promise<string> {
  return Excel.run(async (ctx) => {
    const rng = rangeFromAddress(ctx, String(args.address));
    if (args.fill) rng.format.fill.color = String(args.fill);
    if (args.fontColor) rng.format.font.color = String(args.fontColor);
    if (typeof args.bold === "boolean") rng.format.font.bold = args.bold;
    if (typeof args.italic === "boolean") rng.format.font.italic = args.italic;
    if (args.numberFormat) {
      rng.load("rowCount,columnCount");
      await ctx.sync();
      rng.numberFormat = Array.from({ length: rng.rowCount }, () =>
        Array.from({ length: rng.columnCount }, () => String(args.numberFormat))
      );
    }
    rng.load("address");
    await ctx.sync();
    return `Formatted ${rng.address}.`;
  });
}

async function addWorksheet(name: string): Promise<string> {
  return Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.add(name);
    ws.load("name");
    await ctx.sync();
    return `Added worksheet "${ws.name}".`;
  });
}

/** Map a model-supplied chart type (loose synonyms) to an Excel.ChartType name. */
export function resolveChartType(t: string): string {
  const key = String(t).trim().toLowerCase().replace(/[^a-z]/g, "");
  return CHART_TYPES[key] ?? "ColumnClustered";
}

async function createChart(address: string, chartType: string, title: string): Promise<string> {
  return Excel.run(async (ctx) => {
    const bang = address.indexOf("!");
    const sheet =
      bang >= 0
        ? ctx.workbook.worksheets.getItem(address.slice(0, bang).replace(/^'|'$/g, ""))
        : ctx.workbook.worksheets.getActiveWorksheet();
    const dataRange = sheet.getRange(bang >= 0 ? address.slice(bang + 1) : address);
    const type = resolveChartType(chartType);
    const chart = sheet.charts.add(type as Excel.ChartType, dataRange, Excel.ChartSeriesBy.auto);
    if (title) chart.title.text = title;
    chart.load("name");
    await ctx.sync();
    return `Created ${type} chart${title ? ` "${title}"` : ""} from ${address}.`;
  });
}


// --------------------------------------------------------------------------
// Portable steering operations

function worksheetFromName(ctx: Excel.RequestContext, sheet?: string): Excel.Worksheet {
  return sheet
    ? ctx.workbook.worksheets.getItem(sheet)
    : ctx.workbook.worksheets.getActiveWorksheet();
}

async function setNumberFormat(address: string, format: string): Promise<string> {
  return Excel.run(async (ctx) => {
    const rng = rangeFromAddress(ctx, address);
    rng.load("rowCount,columnCount,address");
    await ctx.sync();
    rng.numberFormat = Array.from({ length: rng.rowCount }, () =>
      Array.from({ length: rng.columnCount }, () => format)
    );
    await ctx.sync();
    return `Applied ${format} to ${rng.address}.`;
  });
}

async function mergeCells(address: string, center: boolean): Promise<string> {
  return Excel.run(async (ctx) => {
    const rng = rangeFromAddress(ctx, address);
    rng.merge(false);
    if (center) rng.format.horizontalAlignment = Excel.HorizontalAlignment.center;
    rng.load("address");
    await ctx.sync();
    return `Merged ${rng.address}${center ? " (centred)" : ""}.`;
  });
}

async function unmergeCells(address: string): Promise<string> {
  return Excel.run(async (ctx) => {
    const rng = rangeFromAddress(ctx, address);
    rng.unmerge();
    rng.load("address");
    await ctx.sync();
    return `Unmerged ${rng.address}.`;
  });
}

async function setColumnWidth(
  column: string,
  width: number,
  unit: string | undefined,
  sheet: string | undefined
): Promise<string> {
  const points = columnWidthToPoints(width, (unit as "chars" | "points" | "px") ?? "chars");
  return Excel.run(async (ctx) => {
    const ws = worksheetFromName(ctx, sheet);
    ws.getRange(`${column}:${column}`).format.columnWidth = points;
    await ctx.sync();
    return `Set column ${column}${sheet ? ` on ${sheet}` : ""} to ${points.toFixed(2)} pt.`;
  });
}

async function setRowHeight(
  row: number,
  height: number,
  unit: string | undefined,
  sheet: string | undefined
): Promise<string> {
  const points = rowHeightToPoints(height, (unit as "points" | "px") ?? "points");
  return Excel.run(async (ctx) => {
    const ws = worksheetFromName(ctx, sheet);
    ws.getRange(`${row}:${row}`).format.rowHeight = points;
    await ctx.sync();
    return `Set row ${row}${sheet ? ` on ${sheet}` : ""} to ${points.toFixed(2)} pt.`;
  });
}

async function recalculateWorkbook(): Promise<string> {
  return Excel.run(async (ctx) => {
    ctx.workbook.application.calculate("Full");
    await ctx.sync();
    return "Recalculated the workbook.";
  });
}

async function renameWorksheet(from: string, to: string): Promise<string> {
  return Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(from);
    ws.name = to;
    await ctx.sync();
    return `Renamed "${from}" to "${to}".`;
  });
}

/** Execute one steering operation against an open request context. */
async function applySteeringOp(ctx: Excel.RequestContext, op: SteeringOp): Promise<string> {
  switch (op.op) {
    case "set_value": {
      const rng = rangeFromAddress(ctx, op.sheet ? `${op.sheet}!${op.ref}` : op.ref);
      rng.values = [[op.value]];
      return `set_value ${op.ref}`;
    }
    case "set_formula": {
      const rng = rangeFromAddress(ctx, op.sheet ? `${op.sheet}!${op.range}` : op.range);
      rng.load("rowCount,columnCount");
      await ctx.sync();
      const formula = op.formula.startsWith("=") ? op.formula : `=${op.formula}`;
      rng.formulas = Array.from({ length: rng.rowCount }, () =>
        Array.from({ length: rng.columnCount }, () => formula)
      );
      return `set_formula ${op.range}`;
    }
    case "format": {
      const rng = rangeFromAddress(ctx, op.sheet ? `${op.sheet}!${op.range}` : op.range);
      if (typeof op.bold === "boolean") rng.format.font.bold = op.bold;
      if (typeof op.italic === "boolean") rng.format.font.italic = op.italic;
      if (typeof op.font_size === "number") rng.format.font.size = op.font_size;
      if (op.font_color) rng.format.font.color = op.font_color;
      if (op.fill) rng.format.fill.color = op.fill;
      if (op.number_format) {
        rng.load("rowCount,columnCount");
        await ctx.sync();
        rng.numberFormat = Array.from({ length: rng.rowCount }, () =>
          Array.from({ length: rng.columnCount }, () => String(op.number_format))
        );
      }
      if (op.align) {
        const alignment = {
          left: Excel.HorizontalAlignment.left,
          center: Excel.HorizontalAlignment.center,
          right: Excel.HorizontalAlignment.right,
          general: Excel.HorizontalAlignment.general,
        }[op.align];
        rng.format.horizontalAlignment = alignment;
      }
      return `format ${op.range}`;
    }
    case "merge": {
      const rng = rangeFromAddress(ctx, op.sheet ? `${op.sheet}!${op.range}` : op.range);
      rng.merge(false);
      if (op.center !== false) rng.format.horizontalAlignment = Excel.HorizontalAlignment.center;
      return `merge ${op.range}`;
    }
    case "unmerge": {
      const rng = rangeFromAddress(ctx, op.sheet ? `${op.sheet}!${op.range}` : op.range);
      rng.unmerge();
      return `unmerge ${op.range}`;
    }
    case "column_width": {
      const ws = worksheetFromName(ctx, op.sheet);
      ws.getRange(`${op.col}:${op.col}`).format.columnWidth = columnWidthToPoints(op.width, op.unit);
      return `column_width ${op.col}=${op.width}`;
    }
    case "row_height": {
      const ws = worksheetFromName(ctx, op.sheet);
      ws.getRange(`${op.row}:${op.row}`).format.rowHeight = rowHeightToPoints(op.height, op.unit);
      return `row_height ${op.row}=${op.height}`;
    }
    case "add_sheet": {
      ctx.workbook.worksheets.add(op.name);
      return `add_sheet ${op.name}`;
    }
    case "rename_sheet": {
      ctx.workbook.worksheets.getItem(op.from).name = op.to;
      return `rename_sheet ${op.from}->${op.to}`;
    }
    case "recalculate": {
      ctx.workbook.application.calculate("Full");
      return "recalculate";
    }
    case "assert": {
      const rng = rangeFromAddress(ctx, op.sheet ? `${op.sheet}!${op.ref}` : op.ref);
      rng.load("values");
      await ctx.sync();
      const actual = rng.values[0]?.[0];
      const ok = String(actual ?? "") === String(op.equals ?? "");
      return `assert ${op.ref}: expected ${JSON.stringify(op.equals)}, got ${JSON.stringify(actual)} ${ok ? "OK" : "MISMATCH"}`;
    }
  }
}

async function runSteering(taskInput: unknown): Promise<string> {
  const task: SteeringTask = normalizeSteeringTask(taskInput);
  const results = await Excel.run(async (ctx) => {
    const done: string[] = [];
    for (const op of task.ops) done.push(await applySteeringOp(ctx, op));
    if (task.recalc) ctx.workbook.application.calculate("Full");
    await ctx.sync();
    return done;
  });
  return `${task.name ?? "steering task"}: ${results.length} ops applied — ${results.join("; ")}`;
}
