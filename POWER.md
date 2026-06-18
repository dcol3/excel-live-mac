---
name: "excel-live-mac"
displayName: "Live Excel Automation (macOS)"
description: "Interact with running Excel instances in real time on macOS. Read, write, and format cells in open workbooks, run VBA macros, and build spreadsheets interactively while users watch changes appear instantly."
keywords: ["excel", "spreadsheet", "xlwings", "workbook", "live", "automation", "macos", "finance", "reports", "mac"]
author: "Dillon Cole"
---

# Onboarding

## Prerequisites

1. **macOS** — This power uses the AppleScript bridge (xlwings) which is macOS-only
2. **Microsoft Excel** installed and licensed on your Mac
3. **Python 3.10+** with `xlwings` and `mcp` packages installed:
   ```bash
   pip install xlwings mcp
   ```
4. **macOS Automation permission** — On first use, macOS will prompt you to allow Terminal (or your IDE) to control Excel. Grant this in System Settings → Privacy & Security → Automation

## Quick Test

After installation, open any `.xlsx` file in Excel, then ask Kiro:
> "List all open workbooks"

If it returns the workbook name, you're connected.

---

# Overview

Live Excel Automation lets Kiro interact with **running Excel instances** on macOS. Unlike file-based tools (openpyxl) that work on saved files offline, this power talks to the live Excel app — changes appear instantly in the spreadsheet you're looking at.

**Key difference from file-based Excel tools:**

| Capability | File-Based (openpyxl) | Live Excel (this power) |
|---|---|---|
| See changes in real time | No — must reopen file | Yes — instant |
| Read current selection | No | Yes |
| Run VBA macros | No | Yes |
| Work with open workbooks | No — file must be closed | Yes — requires Excel open |
| Platform | Any OS | macOS only |

**Authentication**: None required. Communicates locally via AppleScript bridge.

## Available MCP Servers

### excel-live-mac
**Package:** `excel-live-mac` (via uvx/PyPI)
**Connection:** stdio (local process)
**Dependencies:** `xlwings`, `mcp` Python packages
**Platform:** macOS only (AppleScript bridge to Excel)

**Available Tools (10):**

| Tool | Description |
|---|---|
| `excel_list_workbooks` | List all currently open workbooks (name, path, sheet count) |
| `excel_get_sheet_names` | Get all sheet names from an open workbook |
| `excel_read_range` | Read a range from an open workbook as a 2D array |
| `excel_write_range` | Write a 2D array to a range (changes appear immediately) |
| `excel_write_cell` | Write a single value or formula to a cell |
| `excel_read_cell` | Read value and formula from a single cell |
| `excel_open_workbook` | Open a workbook file in Excel (launches Excel if needed) |
| `excel_save_workbook` | Save the active or specified workbook |
| `excel_run_macro` | Execute a VBA macro in the active workbook |
| `excel_get_selection` | Get the currently selected range, values, and sheet |

## Tool Usage Examples

```python
# List open workbooks
excel_list_workbooks()

# Read a range
excel_read_range(range="A1:D20", sheet="Summary")

# Write data (appears instantly in Excel)
excel_write_range(range="A1", data=[
    ["Product", "Revenue", "Units", "Margin"],
    ["Widget A", 125000, 5000, "18.5%"],
    ["Widget B", 98000, 3200, "22.1%"]
])

# Write a formula
excel_write_cell(cell="E2", value="=B2*C2")

# Read what the user has selected
excel_get_selection()

# Run a macro
excel_run_macro(macro="RecalcAll")

# Open a specific file
excel_open_workbook(file_path="/Users/me/Documents/report.xlsx")
```

## Common Workflows

### Build a Report from Data
1. Pull data from a database or API
2. Open or create a workbook with `excel_open_workbook`
3. Write headers and data with `excel_write_range`
4. Add formulas with `excel_write_cell`
5. Save with `excel_save_workbook`

### Interactive Data Entry
1. User has a workbook open in Excel
2. Kiro reads the current state with `excel_read_range`
3. Computes new values based on analysis
4. Writes results back — user sees updates live

### Build a Financial Model
1. Pull data from a database or API
2. Open or create a workbook with `excel_open_workbook`
3. Write assumptions, inputs, and formulas with `excel_write_range` and `excel_write_cell`
4. User watches the model populate in real time
5. Save with `excel_save_workbook`

## Troubleshooting

**"Excel is not running"**: Open Excel manually or use `excel_open_workbook` with a file path.

**"Cannot connect to Excel"**: Check that macOS Automation permission is granted. Go to System Settings → Privacy & Security → Automation → ensure your terminal/IDE can control Microsoft Excel.

**Permission denied on first use**: macOS will show a dialog asking if Terminal can control Excel. Click "OK". If you dismissed it, go to System Settings to re-enable.

**xlwings not found**: Run `pip install xlwings` in the same Python environment that Kiro uses.

**Windows/Linux**: This power is macOS-only. For cross-platform Excel work, use the file-based `excel-tools` MCP server instead (openpyxl-based, works on closed files).

## Configuration

```json
{
  "mcpServers": {
    "excel-live-mac": {
      "command": "uvx",
      "args": ["excel-live-mac@latest"],
      "env": {},
      "disabled": false
    }
  }
}
```

## License

MIT License. Created by Dillon Cole.
