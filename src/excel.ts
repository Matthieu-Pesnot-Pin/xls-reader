import XLSX from "xlsx";
import fs from "fs";
import path from "path";

export function listSheets(filePath: string): string[] {
  if (!fs.existsSync(filePath)) {
    throw new Error(`File not found: ${filePath}`);
  }
  const workbook = XLSX.readFile(filePath);
  return workbook.SheetNames;
}

export type OutputFormat = "markdown" | "json";

export interface ReadSheetOptions {
  cellBudget?: number;
  maxCols?: number;
  maxRows?: number;
  headerRow?: boolean;
  format?: OutputFormat;
}

// ─── Cell value helpers ─────────────────────────────────────────────────────

function isEmptyValue(value: unknown): boolean {
  return value === null || value === undefined || value === "";
}

/** Stringify a cell for Markdown (in-cell line breaks escaped to keep the table intact). */
function cellToString(value: unknown): string {
  if (isEmptyValue(value)) return "";
  if (value instanceof Date) return value.toLocaleDateString();
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value).trim().replace(/\r\n|\r|\n/g, "\\n");
}

/** Convert a cell to a native JSON value (types preserved, dates as ISO, real line breaks kept). */
function cellToJson(value: unknown): unknown {
  if (value instanceof Date) return value.toISOString();
  if (typeof value === "string") return value.trim();
  return value;
}

function toStringCells(raw: unknown[], numCols: number): string[] {
  return Array.from({ length: numCols }, (_, i) => cellToString(raw[i]));
}

// ─── Sections (an empty row delimits a new table) ───────────────────────────

function isEmptyRawRow(row: unknown[]): boolean {
  return row.every(isEmptyValue);
}

function groupIntoSections(rows: unknown[][]): unknown[][][] {
  const sections: unknown[][][] = [[]];
  for (const row of rows) {
    if (isEmptyRawRow(row)) {
      if (sections[sections.length - 1].length > 0) {
        sections.push([]);
      }
    } else {
      sections[sections.length - 1].push(row);
    }
  }
  return sections.filter((s) => s.length > 0);
}

// ─── Shared extraction core ─────────────────────────────────────────────────

interface SheetData {
  headerRaw: unknown[] | null;
  sections: unknown[][][];
  numCols: number;
  note: string | null;
}

interface ExtractOptions {
  cellBudget?: number;
  maxCols?: number;
  maxRows?: number;
  headerRow?: boolean;
  /** When true, ignore all limits and return the whole sheet. */
  unlimited?: boolean;
}

/** Read a worksheet into header + grouped sections, applying the size limit unless `unlimited`. */
function extractSheetData(
  ws: XLSX.WorkSheet,
  options: ExtractOptions
): SheetData | null {
  const rawRows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: "" });
  if (rawRows.length === 0) return null;

  const { cellBudget = 2000, maxCols, maxRows, headerRow = true, unlimited = false } = options;

  const totalCols = rawRows.reduce((max, row) => Math.max(max, (row as unknown[]).length), 0);

  let effectiveCols = totalCols;
  let effectiveRows = rawRows.length;
  let note: string | null = null;

  if (!unlimited) {
    effectiveCols = Math.min(totalCols, maxCols ?? totalCols);
    effectiveRows = Math.min(
      rawRows.length,
      maxRows ?? Math.floor(cellBudget / Math.max(effectiveCols, 1))
    );

    const truncatedCols = effectiveCols < totalCols;
    const truncatedRows = effectiveRows < rawRows.length;
    if (truncatedCols || truncatedRows) {
      const w: string[] = [];
      if (truncatedCols) w.push(`${effectiveCols} columns out of ${totalCols}`);
      if (truncatedRows) w.push(`${effectiveRows} rows out of ${rawRows.length}`);
      note = `Display limited to ${w.join(", ")}. Use max_cols / max_rows / cell_budget to adjust.`;
    }
  }

  const matrix = rawRows
    .slice(0, effectiveRows)
    .map((row) => (row as unknown[]).slice(0, effectiveCols));

  let headerRaw: unknown[] | null = null;
  let dataRows = matrix;
  if (headerRow && matrix.length > 0) {
    headerRaw = matrix[0];
    dataRows = matrix.slice(1);
  }

  return { headerRaw, sections: groupIntoSections(dataRows), numCols: effectiveCols, note };
}

// ─── JSON groups builder (shared by single-sheet and whole-workbook output) ──

function buildHeaderKeys(headerRaw: unknown[], numCols: number): string[] {
  return Array.from({ length: numCols }, (_, i) => {
    const key = cellToString(headerRaw[i]);
    return key === "" ? `column_${i + 1}` : key;
  });
}

function rowToObject(row: unknown[], keys: string[]): Record<string, unknown> {
  const obj: Record<string, unknown> = {};
  keys.forEach((key, i) => {
    if (!isEmptyValue(row[i])) obj[key] = cellToJson(row[i]);
  });
  return obj;
}

function rowToArray(row: unknown[], numCols: number): unknown[] {
  return Array.from({ length: numCols }, (_, i) =>
    isEmptyValue(row[i]) ? null : cellToJson(row[i])
  );
}

/** Turn a sheet's sections into JSON groups (objects keyed by header, or arrays of values). */
function buildGroups(data: SheetData): unknown[] {
  if (data.headerRaw) {
    const keys = buildHeaderKeys(data.headerRaw, data.numCols);
    return data.sections.map((section) => section.map((row) => rowToObject(row, keys)));
  }
  return data.sections.map((section) => section.map((row) => rowToArray(row, data.numCols)));
}

// ─── Markdown renderer ──────────────────────────────────────────────────────

function formatRow(cells: string[], numCols: number): string {
  const padded = Array.from({ length: numCols }, (_, i) => cells[i] ?? "");
  return "| " + padded.join(" | ") + " |";
}

function markdownHeader(header: string[], numCols: number): string[] {
  const sep = Array.from({ length: numCols }, (_, i) =>
    "-".repeat(Math.max((header[i] ?? "").length, 3))
  );
  return [formatRow(header, numCols), "| " + sep.join(" | ") + " |"];
}

function renderMarkdown(data: SheetData): string {
  const header = data.headerRaw ? toStringCells(data.headerRaw, data.numCols) : null;

  if (data.sections.length === 0) {
    if (!header) return "_Empty sheet._";
    return [...markdownHeader(header, data.numCols), "_No data rows._"].join("\n");
  }

  const parts = data.sections.map((section) => {
    const lines = header ? markdownHeader(header, data.numCols) : [];
    for (const row of section) lines.push(formatRow(toStringCells(row, data.numCols), data.numCols));
    return lines.join("\n");
  });

  let result = parts.join("\n\n");
  if (data.note) result += `\n\n> ⚠ ${data.note}`;
  return result;
}

// ─── Public API ─────────────────────────────────────────────────────────────

export function readSheet(
  filePath: string,
  sheetName: string,
  options: ReadSheetOptions = {}
): string {
  if (!fs.existsSync(filePath)) {
    throw new Error(`File not found: ${filePath}`);
  }

  const workbook = XLSX.readFile(filePath, { cellDates: true });
  if (!workbook.SheetNames.includes(sheetName)) {
    throw new Error(
      `Sheet '${sheetName}' not found. Available: ${workbook.SheetNames.join(", ")}`
    );
  }

  const { format = "markdown", ...rest } = options;
  const data = extractSheetData(workbook.Sheets[sheetName], rest);

  if (format === "json") {
    const out: Record<string, unknown> = { sheet: sheetName, truncated: data?.note != null };
    if (data?.note) out.note = data.note;
    out.groups = data ? buildGroups(data) : [];
    return JSON.stringify(out, null, 2);
  }

  return data ? renderMarkdown(data) : "_Empty sheet._";
}

/**
 * Read a workbook into a single JSON object, with no size limit.
 * Exports every sheet, or only `sheetName` when provided.
 */
export function readWorkbook(
  filePath: string,
  options: { headerRow?: boolean; sheetName?: string } = {}
): string {
  if (!fs.existsSync(filePath)) {
    throw new Error(`File not found: ${filePath}`);
  }

  const workbook = XLSX.readFile(filePath, { cellDates: true });

  let names = workbook.SheetNames;
  if (options.sheetName) {
    if (!names.includes(options.sheetName)) {
      throw new Error(
        `Sheet '${options.sheetName}' not found. Available: ${names.join(", ")}`
      );
    }
    names = [options.sheetName];
  }

  const sheets = names.map((name) => {
    const data = extractSheetData(workbook.Sheets[name], {
      headerRow: options.headerRow,
      unlimited: true,
    });
    return { sheet: name, groups: data ? buildGroups(data) : [] };
  });

  return JSON.stringify({ file: path.basename(filePath), sheets }, null, 2);
}
