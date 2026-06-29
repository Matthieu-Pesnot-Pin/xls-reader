import XLSX from "xlsx";
import fs from "fs";

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

// ─── Markdown renderer ──────────────────────────────────────────────────────

function formatRow(cells: string[], numCols: number): string {
  const padded = Array.from({ length: numCols }, (_, i) => cells[i] ?? "");
  return "| " + padded.join(" | ") + " |";
}

function sectionToMarkdown(
  section: string[][],
  header: string[] | null,
  numCols: number
): string {
  const lines: string[] = [];
  if (header) {
    lines.push(formatRow(header, numCols));
    const sep = Array.from({ length: numCols }, (_, i) =>
      "-".repeat(Math.max((header[i] ?? "").length, 3))
    );
    lines.push("| " + sep.join(" | ") + " |");
  }
  for (const row of section) {
    lines.push(formatRow(row, numCols));
  }
  return lines.join("\n");
}

function renderMarkdown(
  headerRaw: unknown[] | null,
  sections: unknown[][][],
  numCols: number,
  note: string | null
): string {
  const header = headerRaw ? toStringCells(headerRaw, numCols) : null;

  if (sections.length === 0) {
    if (!header) return "_Empty sheet._";
    const sep = Array.from({ length: numCols }, (_, i) =>
      "-".repeat(Math.max((header[i] ?? "").length, 3))
    );
    return [formatRow(header, numCols), "| " + sep.join(" | ") + " |", "_No data rows._"].join("\n");
  }

  const parts = sections.map((section) =>
    sectionToMarkdown(
      section.map((row) => toStringCells(row, numCols)),
      header,
      numCols
    )
  );

  let result = parts.join("\n\n");
  if (note) result += `\n\n> ⚠ ${note}`;
  return result;
}

// ─── JSON renderer ──────────────────────────────────────────────────────────

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

function renderJson(
  sheetName: string,
  headerRaw: unknown[] | null,
  sections: unknown[][][],
  numCols: number,
  note: string | null
): string {
  let groups: unknown[];
  if (headerRaw) {
    const keys = buildHeaderKeys(headerRaw, numCols);
    groups = sections.map((section) => section.map((row) => rowToObject(row, keys)));
  } else {
    groups = sections.map((section) => section.map((row) => rowToArray(row, numCols)));
  }

  const out: Record<string, unknown> = { sheet: sheetName, truncated: note !== null };
  if (note) out.note = note;
  out.groups = groups;
  return JSON.stringify(out, null, 2);
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

  const { cellBudget = 2000, maxCols, maxRows, headerRow = true, format = "markdown" } = options;

  const ws = workbook.Sheets[sheetName];
  const rawRows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: "" });

  if (rawRows.length === 0) {
    return format === "json"
      ? JSON.stringify({ sheet: sheetName, truncated: false, groups: [] }, null, 2)
      : "_Empty sheet._";
  }

  const totalCols = rawRows.reduce((max, row) => Math.max(max, (row as unknown[]).length), 0);
  const effectiveCols = Math.min(totalCols, maxCols ?? totalCols);
  const effectiveRows = Math.min(
    rawRows.length,
    maxRows ?? Math.floor(cellBudget / Math.max(effectiveCols, 1))
  );

  const truncatedCols = effectiveCols < totalCols;
  const truncatedRows = effectiveRows < rawRows.length;

  let note: string | null = null;
  if (truncatedCols || truncatedRows) {
    const w: string[] = [];
    if (truncatedCols) w.push(`${effectiveCols} columns out of ${totalCols}`);
    if (truncatedRows) w.push(`${effectiveRows} rows out of ${rawRows.length}`);
    note = `Display limited to ${w.join(", ")}. Use max_cols / max_rows / cell_budget to adjust.`;
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

  const sections = groupIntoSections(dataRows);

  return format === "json"
    ? renderJson(sheetName, headerRaw, sections, effectiveCols, note)
    : renderMarkdown(headerRaw, sections, effectiveCols, note);
}
