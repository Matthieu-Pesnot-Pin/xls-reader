#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import dotenv from "dotenv";
import path from "path";
import { fileURLToPath } from "url";
import { z } from "zod";
import { setupLogging } from "./logger.js";
import { listSheets, readSheet, readWorkbook, type WorkbookSource } from "./excel.js";
import { deliverFile, RELAY_SHAPE_DESCRIPTIONS } from "./relay.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const rootDir = path.resolve(__dirname, "..");

dotenv.config({ path: path.join(rootDir, ".env"), override: true });
setupLogging("MASTER");

const server = new McpServer({
  name: "xls-reader",
  version: "1.0.0",
});

// ─── Workbook source (path on this server, or inline content) ─────────────────
const sourceShape = {
  file_path: z
    .string()
    .optional()
    .describe(
      "Path to the Excel file on the machine running this server. Use it only when the file is local to the server; otherwise send file_content"
    ),
  file_content: z
    .string()
    .optional()
    .describe(
      "Base64-encoded content of the Excel file (a 'data:' URL prefix is accepted). Filled in by the client-side relay, which reads the file named by file_path on the caller's machine — a caller that provides file_path has nothing to put here. Takes precedence over file_path"
    ),
  file_name: z
    .string()
    .optional()
    .describe(
      "Original file name, used only to label the output. Also filled in by the relay alongside file_content"
    ),
};

function toSource(args: {
  file_path?: string;
  file_content?: string;
  file_name?: string;
}): WorkbookSource {
  if (!args.file_path && !args.file_content) {
    throw new Error(
      "Provide either file_path (file local to the server) or file_content (base64 content of the file)."
    );
  }
  return { filePath: args.file_path, fileContent: args.file_content, fileName: args.file_name };
}

// ─── list_sheets ──────────────────────────────────────────────────────────────
server.tool(
  "list_sheets",
  [
    "List all sheet names (tabs) in an Excel file (.xlsx, .xls, .ods, .csv).",
    "Pass file_path when the file sits on the machine running this server, or file_content (base64) when it does not.",
  ].join(" "),
  sourceShape,
  async (args) => {
    try {
      const sheets = listSheets(toSource(args));
      return {
        content: [{ type: "text", text: sheets.join("\n") }],
      };
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : String(error);
      return {
        content: [{ type: "text", text: `Error: ${message}` }],
        isError: true,
      };
    }
  }
);

// ─── get_sheet ────────────────────────────────────────────────────────────────
server.tool(
  "get_sheet",
  [
    "Read an Excel sheet and return its content.",
    "Pass file_path when the file sits on the machine running this server, or file_content (base64) when it does not.",
    "Empty rows are treated as section separators — they close the current table and start a new one.",
    "format='markdown' (default) returns Markdown table(s); format='json' returns { sheet, truncated, groups }",
    "where groups is an array of sections. With header_row, each row is an object keyed by column name (empty cells omitted, types preserved); otherwise each row is an array of values.",
    "To export without any size limit, use export_workbook, which writes a JSON file instead of returning it.",
    "Data is limited by a cell budget (cols × rows, default 2000). Use max_cols / max_rows / cell_budget to override.",
  ].join(" "),
  {
    ...sourceShape,
    sheet_name: z
      .string()
      .describe("Name of the sheet to read"),
    format: z
      .enum(["markdown", "json"])
      .optional()
      .describe("Output format: 'markdown' (default) or 'json'"),
    cell_budget: z
      .number()
      .optional()
      .describe("Maximum number of cells to return (cols × rows). Default: 2000"),
    max_cols: z.number().optional().describe("Override maximum number of columns"),
    max_rows: z.number().optional().describe("Override maximum number of rows"),
    header_row: z
      .boolean()
      .optional()
      .describe("Treat the first row as a column header (default: true)"),
  },
  async ({ sheet_name, format, cell_budget, max_cols, max_rows, header_row, ...rest }) => {
    try {
      const source = toSource(rest);

      const output = readSheet(source, sheet_name, {
        format,
        cellBudget: cell_budget,
        maxCols: max_cols,
        maxRows: max_rows,
        headerRow: header_row,
      });
      return {
        content: [{ type: "text", text: output }],
      };
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : String(error);
      return {
        content: [{ type: "text", text: `Error: ${message}` }],
        isError: true,
      };
    }
  }
);

// ─── export_workbook ──────────────────────────────────────────────────────────
// L'export est un outil à part, et non un `format` de get_sheet : le relais rend
// obligatoire tout `<x>_path` dont il voit la paire, donc porter `destination_path` dans
// get_sheet forcerait une destination sur chaque simple lecture de feuille.
server.tool(
  "export_workbook",
  [
    "Export an Excel workbook as a JSON file written on the agent machine, with no size limit.",
    "Exports the whole workbook, or only sheet_name when provided.",
    "Pass file_path when the file sits on the machine running this server, or file_content (base64) when it does not.",
    "The exported file never passes through the conversation.",
  ].join(" "),
  {
    ...sourceShape,
    sheet_name: z
      .string()
      .optional()
      .describe("Omit to export the whole workbook, or provide it to export only that sheet"),
    header_row: z
      .boolean()
      .optional()
      .describe("Treat the first row as a column header (default: true)"),
    destination_path: z.string().describe(RELAY_SHAPE_DESCRIPTIONS.destination_path),
    filename: z.string().describe(RELAY_SHAPE_DESCRIPTIONS.filename),
    destination_overwrite: z
      .boolean()
      .optional()
      .describe(RELAY_SHAPE_DESCRIPTIONS.destination_overwrite),
    destination_relay: z
      .boolean()
      .optional()
      .describe(RELAY_SHAPE_DESCRIPTIONS.destination_relay),
  },
  async ({
    sheet_name,
    header_row,
    destination_path,
    filename,
    destination_overwrite,
    destination_relay,
    ...rest
  }) => {
    try {
      const json = readWorkbook(toSource(rest), { headerRow: header_row, sheetName: sheet_name });
      const blocks = deliverFile(Buffer.from(json, "utf-8"), "application/json", {
        destination_path,
        filename,
        destination_overwrite,
        destination_relay,
      });
      return { content: blocks };
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : String(error);
      return {
        content: [{ type: "text", text: `Error: ${message}` }],
        isError: true,
      };
    }
  }
);

// ─── main ─────────────────────────────────────────────────────────────────────
async function main(): Promise<void> {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("[xls-reader] MCP server started");
}

main().catch((err: unknown) => {
  console.error("[xls-reader] Fatal error:", err);
  process.exit(1);
});
