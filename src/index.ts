#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import dotenv from "dotenv";
import path from "path";
import { fileURLToPath } from "url";
import { z } from "zod";
import { setupLogging } from "./logger.js";
import { listSheets, readSheet } from "./excel.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const rootDir = path.resolve(__dirname, "..");

dotenv.config({ path: path.join(rootDir, ".env"), override: true });
setupLogging("MASTER");

const server = new McpServer({
  name: "xls-reader",
  version: "1.0.0",
});

// ─── list_sheets ──────────────────────────────────────────────────────────────
server.tool(
  "list_sheets",
  "List all sheet names (tabs) in an Excel file (.xlsx, .xls, .ods, .csv).",
  {
    file_path: z.string().describe("Absolute or relative path to the Excel file"),
  },
  async ({ file_path }) => {
    try {
      const sheets = listSheets(file_path);
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
    "Empty rows are treated as section separators — they close the current table and start a new one.",
    "format='markdown' (default) returns Markdown table(s); format='json' returns { sheet, truncated, groups }",
    "where groups is an array of sections. With header_row, each row is an object keyed by column name (empty cells omitted, types preserved); otherwise each row is an array of values.",
    "Data is limited by a cell budget (cols × rows, default 2000). Use max_cols / max_rows / cell_budget to override.",
  ].join(" "),
  {
    file_path: z.string().describe("Absolute or relative path to the Excel file"),
    sheet_name: z.string().describe("Name of the sheet to read"),
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
  async ({ file_path, sheet_name, format, cell_budget, max_cols, max_rows, header_row }) => {
    try {
      const output = readSheet(file_path, sheet_name, {
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
