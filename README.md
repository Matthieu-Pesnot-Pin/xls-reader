# @imenam/xls-reader

An MCP (Model Context Protocol) server that lets AI agents read Excel files. It can list the sheets of a workbook and render any sheet as a Markdown table, with smart handling of empty rows, empty cells, and in-cell line breaks, plus a configurable size limit to keep responses LLM-friendly.

Supports `.xlsx`, `.xls`, `.ods` and `.csv` (anything [SheetJS](https://sheetjs.com/) can read).

## Where the file comes from

Both tools take the workbook either **by path** or **by content**:

- `file_path` — a path on the machine *running the server*. This is what a caller provides.
- `file_content` — the file itself, base64-encoded (a `data:…;base64,` prefix is accepted), with `file_name` to label the output.

At least one of the two is required. When both are given, `file_content` wins.

**When the server runs on another machine, the content is what must travel** — a path would point at a file the server cannot see. That translation is not the agent's job: the client-side relay does it. With [`@imenam/mcp-http-gateway`](https://www.npmjs.com/package/@imenam/mcp-http-gateway) in `--mcp` mode, the relay runs next to the agent, reads the file named by `file_path`, sends `file_content`, and hides `file_content` / `file_name` from the schema it publishes — so the agent only ever names a path. Enable it with `GATEWAY_INLINE_ROOTS`, which also confines *which* folders can be read; see that package's README.

Over stdio on a single machine, `file_path` is read directly and nothing else is needed.

## Tools

### `list_sheets`

Lists all sheet names (tabs) in a workbook.

| Parameter | Type | Required | Description |
|---|---|---|---|
| `file_path` | `string` | ⚠️ | Path to the Excel file. Resolved on the server, or read by the relay (see above) |
| `file_content` | `string` | ⚠️ | Base64 content of the file. Normally filled in by the relay, not by the agent |
| `file_name` | `string` | ❌ | Original file name, used only for labelling. Filled in with `file_content` |

Returns the sheet names, one per line.

### `get_sheet`

Reads a sheet and returns its content as a Markdown table.

| Parameter | Type | Required | Description |
|---|---|---|---|
| `file_path` | `string` | ⚠️ | Path to the Excel file. Resolved on the server, or read by the relay (see above) |
| `file_content` | `string` | ⚠️ | Base64 content of the file. Normally filled in by the relay, not by the agent |
| `file_name` | `string` | ❌ | Original file name, used only for labelling. Filled in with `file_content` |
| `sheet_name` | `string` | ✅ | Name of the sheet to read |
| `format` | `"markdown" \| "json"` | ❌ | Output format (default: `markdown`) |
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

To export without any size limit, use [`export_workbook`](#export_workbook), which writes a JSON file instead of returning it.

#### Size limit

By default the response is capped at `cell_budget = 2000` cells (`cols × rows`). With 10 columns you get ~200 rows; with 20 columns, ~100 rows. When data is truncated, a warning is appended telling you how much was omitted. Override with `cell_budget`, `max_cols`, or `max_rows` to read more (or less).

### `export_workbook`

Exports a workbook as a JSON file, with **no size limit**, and writes it **on the agent machine**. Use it when the data is too large to return inline. By default the entire workbook is exported; provide `sheet_name` to export only one sheet.

| Parameter | Type | Required | Description |
|---|---|---|---|
| `file_path` | `string` | ⚠️ | Path to the Excel file. Resolved on the server, or read by the relay (see above) |
| `file_content` | `string` | ⚠️ | Base64 content of the file. Normally filled in by the relay, not by the agent |
| `file_name` | `string` | ❌ | Original file name, used only for labelling. Filled in with `file_content` |
| `sheet_name` | `string` | ❌ | Omit to export the whole workbook, or set it to export only that sheet |
| `header_row` | `boolean` | ❌ | Treat the first row as a column header (default: `true`) |
| `destination_path` | `string` | ✅ | **Absolute directory** on the agent machine. Created if missing |
| `filename` | `string` | ✅ | Name the exported file gets there, e.g. `data.json`. No path separators |
| `destination_overwrite` | `boolean` | ❌ | Replace the destination file if it exists (default: `false`) |
| `destination_relay` | `boolean` | ❌ | Filled in by the relay, never by the agent |

The exported file never passes through the conversation. When this MCP runs next to the agent (stdio), it writes the file itself. Behind [`@imenam/mcp-http-gateway`](https://www.npmjs.com/package/@imenam/mcp-http-gateway) it runs on the gateway's machine and cannot reach the agent's disk: it then returns the bytes, and the relay — which *does* run on the agent machine — writes them at `destination_path`. That path requires `GATEWAY_INLINE_WRITE_ROOTS` in the relay's `env` block; without it the call fails explicitly rather than dropping the file on the wrong machine.

The file content is:

```json
{
  "file": "data.xlsx",
  "sheets": [
    { "sheet": "Feuille1", "groups": [ /* …same shape as json… */ ] },
    { "sheet": "Feuille2", "groups": [ /* … */ ] }
  ]
}
```

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
