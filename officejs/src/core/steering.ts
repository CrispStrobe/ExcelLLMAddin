// Portable Excel steering tasks.
//
// The same JSON task file is consumed by `excel-steer` (Quds / Excel on the
// desktop) and by the agent's `run_steering` tool here, so a saved recipe can
// be replayed from either side. This module is pure: it validates and normalises
// operations and maps them onto the existing tool vocabulary. Excel execution
// lives in ../excelTools.

export type SteeringOp =
  | { op: "set_value"; sheet?: string; ref: string; value: string | number | boolean | null }
  | { op: "set_formula"; sheet?: string; range: string; formula: string }
  | {
      op: "format";
      sheet?: string;
      range: string;
      bold?: boolean;
      italic?: boolean;
      font_size?: number;
      font_color?: string;
      fill?: string;
      number_format?: string;
      align?: "left" | "center" | "right" | "general";
    }
  | { op: "merge"; sheet?: string; range: string; center?: boolean }
  | { op: "unmerge"; sheet?: string; range: string }
  | { op: "column_width"; sheet?: string; col: string; width: number; unit?: "chars" | "points" | "px" }
  | { op: "row_height"; sheet?: string; row: number; height: number; unit?: "points" | "px" }
  | { op: "add_sheet"; name: string }
  | { op: "rename_sheet"; from: string; to: string }
  | { op: "recalculate" }
  | { op: "assert"; sheet?: string; ref: string; equals: string | number | boolean };

export interface SteeringTask {
  name?: string;
  description?: string;
  recalc?: boolean;
  ops: SteeringOp[];
}

export interface SteeringPlanEntry {
  tool: string;
  args: Record<string, unknown>;
}

const CELL = /^\$?[A-Za-z]{1,3}\$?[1-9]\d*$/;
const RANGE = /^\$?[A-Za-z]{1,3}\$?[1-9]\d*(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?$/;

function requireString(value: unknown, field: string): string {
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error(`steering: "${field}" must be a non-empty string`);
  }
  return value.trim();
}

function requireCell(value: unknown, field: string): string {
  const text = requireString(value, field);
  if (!CELL.test(text)) throw new Error(`steering: "${field}" is not an A1 cell (${text})`);
  return text;
}

function requireRange(value: unknown, field: string): string {
  const text = requireString(value, field);
  if (!RANGE.test(text)) throw new Error(`steering: "${field}" is not an A1 range (${text})`);
  return text;
}

function optionalSheet(value: unknown): string | undefined {
  return typeof value === "string" && value.trim() !== "" ? value.trim() : undefined;
}

function qualified(sheet: string | undefined, address: string): string {
  return sheet ? `${sheet}!${address}` : address;
}

/**
 * Validate and normalise a task. Unknown operations are rejected rather than
 * silently ignored: a recipe that no longer matches the vocabulary must fail
 * loudly before it writes anything.
 */
export function normalizeSteeringTask(input: unknown): SteeringTask {
  const raw = typeof input === "string" ? JSON.parse(input) : input;
  if (!raw || typeof raw !== "object") throw new Error("steering: task must be an object");
  const record = raw as Record<string, unknown>;
  const opsRaw = record.ops;
  if (!Array.isArray(opsRaw) || opsRaw.length === 0) {
    throw new Error("steering: task needs a non-empty \"ops\" array");
  }
  if (opsRaw.length > 500) {
    throw new Error("steering: a task may contain at most 500 operations");
  }
  const ops = opsRaw.map((entry, index) => normalizeOp(entry, index));
  return {
    name: typeof record.name === "string" ? record.name : undefined,
    description: typeof record.description === "string" ? record.description : undefined,
    recalc: record.recalc === true,
    ops,
  };
}

function normalizeOp(entry: unknown, index: number): SteeringOp {
  if (!entry || typeof entry !== "object") {
    throw new Error(`steering: op ${index} must be an object`);
  }
  const raw = entry as Record<string, any>;
  const sheet = optionalSheet(raw.sheet);
  switch (raw.op) {
    case "set_value": {
      const value = raw.value;
      if (value !== null && !["string", "number", "boolean"].includes(typeof value)) {
        throw new Error(`steering: op ${index} set_value needs a scalar value`);
      }
      return { op: "set_value", sheet, ref: requireCell(raw.ref, "ref"), value };
    }
    case "set_formula":
      return { op: "set_formula", sheet, range: requireRange(raw.range, "range"), formula: requireString(raw.formula, "formula") };
    case "format": {
      const align = raw.align;
      if (align !== undefined && !["left", "center", "right", "general"].includes(align)) {
        throw new Error(`steering: op ${index} has an unsupported alignment`);
      }
      return {
        op: "format",
        sheet,
        range: requireRange(raw.range, "range"),
        bold: typeof raw.bold === "boolean" ? raw.bold : undefined,
        italic: typeof raw.italic === "boolean" ? raw.italic : undefined,
        font_size: typeof raw.font_size === "number" ? raw.font_size : undefined,
        font_color: typeof raw.font_color === "string" ? raw.font_color : undefined,
        fill: typeof raw.fill === "string" ? raw.fill : undefined,
        number_format: typeof raw.number_format === "string" ? raw.number_format : undefined,
        align,
      };
    }
    case "merge":
      return { op: "merge", sheet, range: requireRange(raw.range, "range"), center: raw.center !== false };
    case "unmerge":
      return { op: "unmerge", sheet, range: requireRange(raw.range, "range") };
    case "column_width": {
      const width = Number(raw.width);
      if (!Number.isFinite(width) || width <= 0) throw new Error(`steering: op ${index} column_width must be positive`);
      const unit = raw.unit === "points" || raw.unit === "px" ? raw.unit : "chars";
      return { op: "column_width", sheet, col: requireString(raw.col, "col"), width, unit };
    }
    case "row_height": {
      const height = Number(raw.height);
      if (!Number.isFinite(height) || height <= 0) throw new Error(`steering: op ${index} row_height must be positive`);
      const unit = raw.unit === "px" ? "px" : "points";
      const row = Number(raw.row);
      if (!Number.isInteger(row) || row < 1) throw new Error(`steering: op ${index} row_height needs a 1-based row`);
      return { op: "row_height", sheet, row, height, unit };
    }
    case "add_sheet":
      return { op: "add_sheet", name: requireString(raw.name, "name") };
    case "rename_sheet":
      return { op: "rename_sheet", from: requireString(raw.from, "from"), to: requireString(raw.to, "to") };
    case "recalculate":
      return { op: "recalculate" };
    case "assert":
      return { op: "assert", sheet, ref: requireCell(raw.ref, "ref"), equals: raw.equals };
    default:
      throw new Error(`steering: unknown op "${String(raw.op)}"`);
  }
}

/**
 * Express a task as calls to the agent's Excel tools. Useful for previews and
 * for hosts that only expose the granular tool executor.
 */
export function steeringToToolCalls(task: SteeringTask): SteeringPlanEntry[] {
  return task.ops.flatMap((op): SteeringPlanEntry[] => {
    switch (op.op) {
      case "set_value":
        return [{ tool: "write_range", args: { address: qualified(op.sheet, op.ref), values: [[op.value]] } }];
      case "set_formula":
        return [{ tool: "write_formula", args: { address: qualified(op.sheet, op.range), formula: op.formula } }];
      case "format":
        return [
          {
            tool: "set_format",
            args: {
              address: qualified(op.sheet, op.range),
              bold: op.bold,
              italic: op.italic,
              fill: op.fill,
              fontColor: op.font_color,
              numberFormat: op.number_format,
              fontSize: op.font_size,
              horizontalAlignment: op.align,
            },
          },
        ];
      case "merge":
        return [{ tool: "merge_cells", args: { address: qualified(op.sheet, op.range), center: op.center } }];
      case "unmerge":
        return [{ tool: "unmerge_cells", args: { address: qualified(op.sheet, op.range) } }];
      case "column_width":
        return [{ tool: "set_column_width", args: { column: op.col, width: op.width, unit: op.unit, sheet: op.sheet } }];
      case "row_height":
        return [{ tool: "set_row_height", args: { row: op.row, height: op.height, unit: op.unit, sheet: op.sheet } }];
      case "add_sheet":
        return [{ tool: "add_worksheet", args: { name: op.name } }];
      case "rename_sheet":
        return [{ tool: "rename_worksheet", args: { from: op.from, to: op.to } }];
      case "recalculate":
        return [{ tool: "recalculate", args: {} }];
      case "assert":
        return [{ tool: "read_range", args: { address: qualified(op.sheet, op.ref) } }];
    }
  });
}

/** Convert Excel's character-based width to points (Calibri 11, 96 dpi). */
export function columnWidthToPoints(width: number, unit: "chars" | "points" | "px" = "chars"): number {
  if (unit === "px") return (width * 72) / 96;
  if (unit === "points") return width;
  return (width * 7 * 72) / 96;
}

export function rowHeightToPoints(height: number, unit: "points" | "px" = "points"): number {
  return unit === "px" ? (height * 72) / 96 : height;
}

export function summarizeTask(task: SteeringTask): string {
  const counts = new Map<string, number>();
  for (const op of task.ops) counts.set(op.op, (counts.get(op.op) ?? 0) + 1);
  const parts = [...counts.entries()].map(([name, count]) => `${count}× ${name}`);
  return `${task.name ?? "steering task"}: ${task.ops.length} ops (${parts.join(", ")})`;
}
