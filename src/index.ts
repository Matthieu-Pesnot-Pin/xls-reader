#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import dotenv from "dotenv";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import { z } from "zod";
import { setupLogging } from "./logger.js";
import { listSheets, readSheet, readWorkbook } from "./excel.js";

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
    "format='json-file' writes JSON (no size limit) to output_path (required) and returns the path; it exports the whole workbook, or only sheet_name when provided. Size limits are ignored.",
    "Data is limited by a cell budget (cols × rows, default 2000) for 'markdown' and 'json'. Use max_cols / max_rows / cell_budget to override.",
  ].join(" "),
  {
    file_path: z.string().describe("Absolute or relative path to the Excel file"),
    sheet_name: z
      .string()
      .optional()
      .describe("Name of the sheet to read. Required for 'markdown'/'json'. For 'json-file': omit to export the whole workbook, or provide it to export only that sheet"),
    format: z
      .enum(["markdown", "json", "json-file"])
      .optional()
      .describe("Output format: 'markdown' (default), 'json', or 'json-file' (whole workbook to disk)"),
    output_path: z
      .string()
      .optional()
      .describe("Destination file path for the JSON export (required when format is 'json-file')"),
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
  async ({ file_path, sheet_name, format, output_path, cell_budget, max_cols, max_rows, header_row }) => {
    try {
      if (format === "json-file") {
        if (!output_path) {
          return {
            content: [{ type: "text", text: "Error: output_path is required when format is 'json-file'." }],
            isError: true,
          };
        }
        const json = readWorkbook(file_path, { headerRow: header_row, sheetName: sheet_name });
        const resolved = path.resolve(output_path);
        fs.mkdirSync(path.dirname(resolved), { recursive: true });
        fs.writeFileSync(resolved, json, "utf-8");
        return {
          content: [{ type: "text", text: `Workbook exported to ${resolved}` }],
        };
      }

      if (!sheet_name) {
        return {
          content: [{ type: "text", text: "Error: sheet_name is required." }],
          isError: true,
        };
      }

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
