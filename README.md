# @imenam/xls-reader

An MCP (Model Context Protocol) server that lets AI agents read Excel files. It can list the sheets of a workbook and render any sheet as a Markdown table, with smart handling of empty rows, empty cells, and in-cell line breaks, plus a configurable size limit to keep responses LLM-friendly.

Supports `.xlsx`, `.xls`, `.ods` and `.csv` (anything [SheetJS](https://sheetjs.com/) can read).

## Tools

### `list_sheets`

Lists all sheet names (tabs) in a workbook.

| Parameter | Type | Required | Description |
|---|---|---|---|
| `file_path` | `string` | ✅ | Absolute or relative path to the Excel file |

Returns the sheet names, one per line.

### `get_sheet`

Reads a sheet and returns its content as a Markdown table.

| Parameter | Type | Required | Description |
|---|---|---|---|
| `file_path` | `string` | ✅ | Absolute or relative path to the Excel file |
| `sheet_name` | `string` | ⚠️ | Name of the sheet to read. Required for `markdown`/`json`. For `json-file`: omit to export the whole workbook, or set it to export only that sheet |
| `format` | `"markdown" \| "json" \| "json-file"` | ❌ | Output format (default: `markdown`) |
| `output_path` | `string` | ⚠️ | Destination file for the JSON export (required when `format` is `json-file`) |
| `cell_budget` | `number` | ❌ | Maximum number of cells to return (cols × rows). Default: `2000` |
| `max_cols` | `number` | ❌ | Override the maximum number of columns |
| `max_rows` | `number` | ❌ | Override the maximum number of rows |
| `header_row` | `boolean` | ❌ | Treat the first row as a column header (default: `true`) |

#### Smart content handling

- **Empty rows are section separators.** A fully empty row closes the current table and starts a new one; when `header_row` is enabled, the header is repeated on each section.
- **Empty cells are kept as blanks** so columns stay aligned.
- **In-cell line breaks** are converted to a literal `\n` so they don't break the Markdown table row.
- **Dates and booleans** are rendered as readable text rather than Excel serial numbers.

#### Output formats

With `format: "markdown"` (default) you get one Markdown table per section.

With `format: "json"` you get a structured object — better when the agent needs to *process* the data rather than read it. Native types are preserved (numbers, booleans, dates as ISO strings) and real line breaks are kept:

```json
{
  "sheet": "Feuille1",
  "truncated": false,
  "groups": [
    [
      { "Nom": "Dupont", "Prénom": "Jean", "Age": 42 },
      { "Nom": "Martin", "Age": 35 }
    ]
  ]
}
```

- `groups` is an array of sections (one per block separated by an empty row).
- With `header_row: true`, each row is an **object** keyed by column name; **empty cells are omitted**.
- With `header_row: false`, each row is an **array** of values (empty cells become `null`).
- When data is truncated, `truncated` is `true` and a `note` field explains what was omitted.

With `format: "json-file"` the data (no size limit) is written to `output_path` and the tool returns the saved path. Use this when the data is too large to return inline. By default the **entire workbook** is exported; provide `sheet_name` to export only that one sheet. Size limits are ignored. The file content is:

```json
{
  "file": "data.xlsx",
  "sheets": [
    { "sheet": "Feuille1", "groups": [ /* …same shape as json… */ ] },
    { "sheet": "Feuille2", "groups": [ /* … */ ] }
  ]
}
```

#### Size limit

By default the response is capped at `cell_budget = 2000` cells (`cols × rows`). With 10 columns you get ~200 rows; with 20 columns, ~100 rows. When data is truncated, a warning is appended telling you how much was omitted. Override with `cell_budget`, `max_cols`, or `max_rows` to read more (or less).

## Installation

```bash
npm install -g @imenam/xls-reader
```

Or use it directly with `npx`:

```json
{
  "xls-reader": {
    "command": "npx",
    "args": ["-y", "@imenam/xls-reader"]
  }
}
```

Or point to a local build:

```json
{
  "xls-reader": {
    "command": "node",
    "args": ["/absolute/path/to/xls-reader/dist/index.js"]
  }
}
```

## Configuration

No environment variables are required. Optionally, set `MCP_LOG_DIR` to change where logs are written (default: `C:\var\log\xls-reader` on Windows, `/var/log/xls-reader` otherwise).

```env
# .env
MCP_LOG_DIR=/absolute/path/to/logs
```

## Development

```bash
npm install
npm run build      # compile TypeScript to dist/
npm run dev        # run from source with tsx
```

## License

MIT
