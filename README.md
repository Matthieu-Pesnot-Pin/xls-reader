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
| `sheet_name` | `string` | ✅ | Name of the sheet to read |
| `cell_budget` | `number` | ❌ | Maximum number of cells to return (cols × rows). Default: `2000` |
| `max_cols` | `number` | ❌ | Override the maximum number of columns |
| `max_rows` | `number` | ❌ | Override the maximum number of rows |
| `header_row` | `boolean` | ❌ | Treat the first row as a column header (default: `true`) |

#### Smart content handling

- **Empty rows are section separators.** A fully empty row closes the current table and starts a new one; when `header_row` is enabled, the header is repeated on each section.
- **Empty cells are kept as blanks** so columns stay aligned.
- **In-cell line breaks** are converted to a literal `\n` so they don't break the Markdown table row.
- **Dates and booleans** are rendered as readable text rather than Excel serial numbers.

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
