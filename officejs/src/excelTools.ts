// Excel tool implementations for the agent (browser/Office.js). Each tool has a
// JSON-schema for the model and a handler that runs via Excel.run against the
// live workbook. Not unit-tested (needs Excel); the agent loop that drives them
// is tested in core/__tests__/agent.test.ts.

import { ToolSchema } from "./core/agent";

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
];

/** Tools that change the workbook (vs. read-only). Used for approve-before-apply. */
export const WRITE_TOOLS = new Set(["write_range", "write_formula", "set_format", "add_worksheet"]);

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
