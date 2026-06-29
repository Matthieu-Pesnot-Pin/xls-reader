import XLSX from "xlsx";
import fs from "fs";

export function listSheets(filePath: string): string[] {
  if (!fs.existsSync(filePath)) {
    throw new Error(`File not found: ${filePath}`);
  }
  const workbook = XLSX.readFile(filePath);
  return workbook.SheetNames;
}

export interface ReadSheetOptions {
  cellBudget?: number;
  maxCols?: number;
  maxRows?: number;
  headerRow?: boolean;
}

function cellToString(value: unknown): string {
  if (value === null || value === undefined || value === "") return "";
  if (value instanceof Date) return value.toLocaleDateString();
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  // Replace in-cell line breaks with a literal \n so they don't break the Markdown table row.
  return String(value).trim().replace(/\r\n|\r|\n/g, "\\n");
}

function toStringRow(raw: unknown[], numCols: number): string[] {
  return Array.from({ length: numCols }, (_, i) => cellToString(raw[i]));
}

function isEmptyRow(cells: string[]): boolean {
  return cells.every((c) => c === "");
}

function groupIntoSections(rows: string[][]): string[][][] {
  const sections: string[][][] = [[]];
  for (const row of rows) {
    if (isEmptyRow(row)) {
      if (sections[sections.length - 1].length > 0) {
        sections.push([]);
      }
    } else {
      sections[sections.length - 1].push(row);
    }
  }
  return sections.filter((s) => s.length > 0);
}

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
    const sep = Array.from(
      { length: numCols },
      (_, i) => "-".repeat(Math.max((header[i] ?? "").length, 3))
    );
    lines.push("| " + sep.join(" | ") + " |");
  }

  for (const row of section) {
    lines.push(formatRow(row, numCols));
  }

  return lines.join("\n");
}

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

  const ws = workbook.Sheets[sheetName];
  const rawRows = XLSX.utils.sheet_to_json<unknown[]>(ws, {
    header: 1,
    defval: "",
  });

  if (rawRows.length === 0) {
    return "_Empty sheet._";
  }

  const { cellBudget = 2000, maxCols, maxRows, headerRow = true } = options;

  const totalCols = rawRows.reduce(
    (max, row) => Math.max(max, (row as unknown[]).length),
    0
  );
  const effectiveCols = Math.min(totalCols, maxCols ?? totalCols);
  const effectiveRows = Math.min(
    rawRows.length,
    maxRows ?? Math.floor(cellBudget / Math.max(effectiveCols, 1))
  );

  const truncatedCols = effectiveCols < totalCols;
  const truncatedRows = effectiveRows < rawRows.length;

  const stringRows = rawRows
    .slice(0, effectiveRows)
    .map((row) => toStringRow(row as unknown[], effectiveCols));

  let header: string[] | null = null;
  let dataRows = stringRows;

  if (headerRow && stringRows.length > 0) {
    header = stringRows[0];
    dataRows = stringRows.slice(1);
  }

  const sections = groupIntoSections(dataRows);

  if (sections.length === 0) {
    const lines: string[] = [];
    if (header) {
      lines.push(formatRow(header, effectiveCols));
      const sep = Array.from(
        { length: effectiveCols },
        (_, i) => "-".repeat(Math.max((header![i] ?? "").length, 3))
      );
      lines.push("| " + sep.join(" | ") + " |");
      lines.push("_No data rows._");
    } else {
      return "_Empty sheet._";
    }
    return lines.join("\n");
  }

  const parts = sections.map((section) =>
    sectionToMarkdown(section, header, effectiveCols)
  );

  let result = parts.join("\n\n");

  if (truncatedCols || truncatedRows) {
    const warnings: string[] = [];
    if (truncatedCols)
      warnings.push(`${effectiveCols} columns out of ${totalCols}`);
    if (truncatedRows)
      warnings.push(`${effectiveRows} rows out of ${rawRows.length}`);
    result += `\n\n> ⚠ Display limited to ${warnings.join(", ")}. Use \`max_cols\` / \`max_rows\` / \`cell_budget\` to adjust.`;
  }

  return result;
}
