#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import dotenv from "dotenv";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import { z } from "zod";
import { setupLogging } from "./logger.js";
import { listSheets, readSheet, readWorkbook, type WorkbookSource } from "./excel.js";

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
    "format='json-file' writes JSON (no size limit) to output_path (required) and returns the path; it exports the whole workbook, or only sheet_name when provided. Size limits are ignored.",
    "Data is limited by a cell budget (cols × rows, default 2000) for 'markdown' and 'json'. Use max_cols / max_rows / cell_budget to override.",
  ].join(" "),
  {
    ...sourceShape,
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
      .describe(
        "Destination path on the server's disk for the JSON export (required when format is 'json-file')"
      ),
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
  async ({ sheet_name, format, output_path, cell_budget, max_cols, max_rows, header_row, ...rest }) => {
    try {
      const source = toSource(rest);

      if (format === "json-file") {
        if (!output_path) {
          return {
            content: [{ type: "text", text: "Error: output_path is required when format is 'json-file'." }],
            isError: true,
          };
        }
        const json = readWorkbook(source, { headerRow: header_row, sheetName: sheet_name });
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
