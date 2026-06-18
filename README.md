# excel-live-mac

**Live Excel MCP server for macOS** — interact with running Excel instances in real time.

Unlike file-based tools (openpyxl) that work on saved files offline, this server talks to the live Excel app via the xlwings AppleScript bridge. Changes appear instantly in the spreadsheet you're looking at.

## Features

- **10 tools** for reading, writing, and interacting with open Excel workbooks
- **Real-time updates** — write data and watch it appear in Excel immediately
- **Formula support** — write Excel formulas that calculate live
- **VBA macro execution** — trigger existing macros from your AI agent
- **Selection awareness** — read what the user currently has highlighted
- **Zero config** — no API keys, no network, runs entirely local

## Requirements

- **macOS** (AppleScript bridge — not available on Windows/Linux)
- **Microsoft Excel** installed
- **Python 3.10+**

## Installation

### For Kiro IDE (as a Power)

Import from GitHub in the Kiro Powers panel:
```
https://github.com/dcol3/excel-live-mac
```

### For any MCP client (via uvx)

```bash
uvx --from git+https://github.com/dcol3/excel-live-mac excel-live-mac
```

Or add to your MCP config:
```json
{
  "mcpServers": {
    "excel-live-mac": {
      "command": "uvx",
      "args": ["--from", "git+https://github.com/dcol3/excel-live-mac", "excel-live-mac"],
      "disabled": false
    }
  }
}
```

### Manual install

```bash
pip install excel-live-mac
excel-live-mac
```

## macOS Permission

On first use, macOS will prompt you to allow your terminal to control Microsoft Excel.

Go to **System Settings → Privacy & Security → Automation** and ensure your terminal/IDE has permission to control Excel.

## Tools

| Tool | Description |
|------|-------------|
| `excel_list_workbooks` | List all open workbooks (name, path, sheets) |
| `excel_get_sheet_names` | Get sheet names from a workbook |
| `excel_read_range` | Read a cell range as a 2D array |
| `excel_write_range` | Write a 2D array (appears instantly) |
| `excel_write_cell` | Write a value or formula to one cell |
| `excel_read_cell` | Read one cell's value and formula |
| `excel_open_workbook` | Open a file in Excel |
| `excel_save_workbook` | Save the workbook |
| `excel_run_macro` | Execute a VBA macro |
| `excel_get_selection` | Get the current selection |

## Use Cases

- **Report generation** — AI builds formatted spreadsheets while you watch
- **Financial modeling** — populate budgets, forecasts, and scenario tables live
- **Data entry automation** — read from databases, write directly into open workbooks
- **Interactive analysis** — AI reads your current selection, computes, writes results back

## Kiro Power

This repo includes a `kiro_power/` directory with `POWER.md`, `mcp.json`, and steering files for use as a Kiro Power. When installed as a Power, Kiro's agent automatically knows when and how to use these tools based on keyword activation.

## License

MIT
