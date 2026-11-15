import {
  columnWidthToPoints,
  normalizeSteeringTask,
  rowHeightToPoints,
  steeringToToolCalls,
  summarizeTask,
} from "../steering";

const headerTask = {
  name: "01-format-header",
  recalc: false,
  ops: [
    {
      op: "format",
      range: "A1:D1",
      bold: true,
      font_size: 12,
      fill: "FFEB9C",
      font_color: "9C0006",
      number_format: "0.00",
      align: "center",
    },
    { op: "column_width", col: "B", width: 18 },
    { op: "row_height", row: 1, height: 24 },
  ],
};

describe("normalizeSteeringTask", () => {
  test("accepts the CLI task shape and preserves defaults", () => {
    const task = normalizeSteeringTask(headerTask);
    expect(task.name).toBe("01-format-header");
    expect(task.ops).toHaveLength(3);
    expect(task.ops[0]).toMatchObject({ op: "format", range: "A1:D1", bold: true });
    expect(task.ops[1]).toMatchObject({ op: "column_width", unit: "chars" });
    expect(task.ops[2]).toMatchObject({ op: "row_height", unit: "points" });
  });

  test("parses JSON text, which is how a model may pass a saved recipe", () => {
    expect(normalizeSteeringTask(JSON.stringify(headerTask)).ops).toHaveLength(3);
  });

  test("rejects an empty task and an unknown operation", () => {
    expect(() => normalizeSteeringTask({ ops: [] })).toThrow(/non-empty/);
    expect(() => normalizeSteeringTask({ ops: [{ op: "explode" }] })).toThrow(/unknown op/);
  });

  test("rejects malformed addresses and non-scalar values", () => {
    expect(() => normalizeSteeringTask({ ops: [{ op: "set_value", ref: "A0", value: 1 }] })).toThrow(
      /A1 cell/
    );
    expect(() => normalizeSteeringTask({ ops: [{ op: "set_value", ref: "A1", value: { x: 1 } }] })).toThrow(
      /scalar/
    );
    expect(() => normalizeSteeringTask({ ops: [{ op: "format", range: "A1:B", bold: true }] })).toThrow(
      /A1 range/
    );
  });

  test("rejects non-positive geometry and a bad row number", () => {
    expect(() => normalizeSteeringTask({ ops: [{ op: "column_width", col: "A", width: 0 }] })).toThrow(
      /positive/
    );
    expect(() => normalizeSteeringTask({ ops: [{ op: "row_height", row: 0, height: 10 }] })).toThrow(
      /1-based/
    );
  });

  test("caps a task at 500 operations", () => {
    const ops = Array.from({ length: 501 }, () => ({ op: "recalculate" }));
    expect(() => normalizeSteeringTask({ ops })).toThrow(/at most 500/);
  });
});

describe("steeringToToolCalls", () => {
  test("maps operations onto the agent tool vocabulary", () => {
    const calls = steeringToToolCalls(normalizeSteeringTask(headerTask));
    expect(calls.map((call) => call.tool)).toEqual(["set_format", "set_column_width", "set_row_height"]);
    expect(calls[0].args).toMatchObject({
      address: "A1:D1",
      bold: true,
      fill: "FFEB9C",
      numberFormat: "0.00",
      horizontalAlignment: "center",
    });
  });

  test("qualifies addresses with a sheet name", () => {
    const calls = steeringToToolCalls(
      normalizeSteeringTask({ ops: [{ op: "merge", sheet: "Data", range: "A1:B1" }] })
    );
    expect(calls[0]).toEqual({ tool: "merge_cells", args: { address: "Data!A1:B1", center: true } });
  });
});

describe("unit helpers", () => {
  test("character widths become points and px respects 96 dpi", () => {
    expect(columnWidthToPoints(18, "chars")).toBeCloseTo(94.5);
    expect(columnWidthToPoints(18, "points")).toBe(18);
    expect(columnWidthToPoints(96, "px")).toBe(72);
    expect(rowHeightToPoints(24, "points")).toBe(24);
    expect(rowHeightToPoints(32, "px")).toBe(24);
  });

  test("summary counts each operation", () => {
    expect(summarizeTask(normalizeSteeringTask(headerTask))).toContain("3 ops");
  });
});
