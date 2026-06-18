---
inclusion: always
---

# Excel Live MCP — Steering Guide

This steering file provides guidance for using the Live Excel MCP server to interact with running Excel instances on macOS.

## When to Use Excel Live Tools

Use these tools when the user wants to:
- **Build or populate a spreadsheet interactively** — writing data, formulas, or formatting to a workbook that's open in Excel
- **Read data from an open workbook** — pulling values from a running Excel instance (not a closed file)
- **Run VBA macros** — executing existing macros in a workbook
- **Check the user's current selection** — reading what cells the user has highlighted
- **Create formatted reports live** — building workbooks while the user watches

## When NOT to Use Excel Live Tools

Do NOT use these tools when:
- **Excel is not open** and the user just wants to analyze a .xlsx file → use `excel-tools` (file-based, openpyxl) or pandas instead
- **Creating a brand new file from scratch without live interaction** → use `excel-tools` `write_excel_data` or `build_excel_shell`
- **User is on Windows or Linux** → this power is macOS-only
- **Reading structure/formulas from a closed file** → use `excel-tools` `read_excel_structure` or `analyze_excel_formulas`

## Tool Selection Guide

| Task | Tool | Notes |
|---|---|---|
| See what's open | `excel_list_workbooks` | Good first step to orient |
| Check data size (rows/cols) | `excel_get_used_range` | ALWAYS call this before reading large data. Avoids probing row counts manually. |
| Read data | `excel_read_range` | Use `expand=true` for tables. For large datasets (>500 rows), read only what you need. |
| Write a table of data | `excel_write_range` | 2D array, instant in Excel. Rows auto-padded to equal length. |
| Write a single value or formula | `excel_write_cell` | Formulas start with `=` |
| Create a new blank workbook | `excel_new_workbook` | Use when you need a new file (dashboards, reports). Save after with a path. |
| Open a file | `excel_open_workbook` | Opens an EXISTING file. Do NOT use to create new files. |
| Add a sheet/tab | `excel_add_sheet` | Only works on .xlsx/.xlsm. Does NOT work on CSV files. |
| Save work | `excel_save_workbook` | Optionally save-as with path (use to convert CSV to .xlsx) |
| Run a macro | `excel_run_macro` | Macro must exist in workbook |
| Check user selection | `excel_get_selection` | Contextual awareness |

## Critical Workflow Rules

### NEVER Merge Cells

**This is the #1 cause of formatting failures.** Merged cells cause Excel to throw blocking dialogs when xlwings tries to format them later. The AppleScript bridge freezes and the operation times out.

Rules:
- NEVER use `merge()` or pass merged-range data through `excel_write_range`
- Do NOT write a title that spans columns as a single merged row
- Instead, write the title to cell A1 and color the entire row (A1:D1) without merging
- If cells ARE already merged (from a prior write), unmerge them FIRST before any formatting

### Before Reading Large Data
ALWAYS call `excel_get_used_range` first to learn the dimensions. Do NOT probe row counts by reading A500, A1000, A5000, etc. One call tells you exactly how many rows and columns exist.

### CSV Files Are Single-Sheet
CSV files opened in Excel can only have one sheet. You CANNOT add sheets with `excel_add_sheet`. If you need multiple sheets:
1. Create a new workbook with `excel_new_workbook`
2. Write your dashboard/summary to the new workbook
3. Reference data from the CSV workbook using cross-workbook formulas

Or alternatively:
1. Save the CSV as .xlsx first: `excel_save_workbook(workbook="file.csv", path="/path/to/file.xlsx")`
2. Then add sheets to the .xlsx version

### Writing Data — Row Padding
`excel_write_range` automatically pads rows to equal length. You do NOT need to manually add empty strings to make all rows the same width. Just pass the data naturally.

### Large Datasets (>1000 rows)
For workbooks with thousands of rows, do NOT try to read the entire dataset into context. Instead:
- Use `excel_get_used_range` to understand size
- Read only headers (row 1) to understand columns
- Write formulas (COUNTIF, SUMIF, AVERAGE, etc.) that compute over the range natively in Excel
- This avoids token overflow and is faster than reading + computing in the agent

## Best Practices

### Writing Data
- Always write headers and data together in one `excel_write_range` call when building a new table
- Use `excel_write_cell` for individual formulas that reference other cells
- Write formulas as strings starting with `=` (e.g., `"=SUM(B2:B10)"`)

### Reading Data
- Use `expand=true` when reading a table whose exact range you don't know — it auto-detects the contiguous region
- For large datasets, specify an explicit range to avoid pulling the entire sheet

### Workbook Management
- Call `excel_list_workbooks` first to confirm which workbook is active
- Specify workbook name explicitly when multiple workbooks are open
- Always save after completing a write workflow

### Error Handling
- If "Excel is not running" → suggest user open Excel or use `excel_open_workbook`
- If "Workbook not found" → call `excel_list_workbooks` to show what's available
- If macOS permission error → direct user to System Settings → Privacy & Security → Automation

## Formatting via xlwings Scripts (REQUIRED APPROACH)

The MCP tools do NOT expose cell formatting (colors, fonts, borders, column widths). To format cells, you MUST use a standalone xlwings Python script executed via bash.

### Key Rules
1. **Always do ALL formatting in a SINGLE script** — do not split into multiple steps
2. **Always unmerge first** if any prior operation may have merged cells
3. **Write the script to a file** (`/tmp/format_excel.py`) — do NOT use inline `-c` strings (multiline breaks in conda)
4. **Use a 60-second timeout** — xlwings over AppleScript is slow for bulk operations
5. **Never use AppleScript heredocs** for formatting — they timeout on complex operations

### The Pattern (One Script, One Execution)

```python
# Write to /tmp/format_excel.py, then run:
# conda run -n base python /tmp/format_excel.py
# Timeout: 60000ms

import xlwings as xw

wb = xw.Book("workbook_name.xlsx")
ws = wb.sheets["Sheet Name"]

# STEP 1: Always unmerge first (prevents blocking dialogs)
ws.range("A1:Z100").api.unmerge()

# STEP 2: Column widths
ws.range("A:A").column_width = 34
ws.range("B:B").column_width = 18

# STEP 3: Colors and fonts (all in one pass)
NAVY = (23, 42, 69)
WHITE = (255, 255, 255)
ACCENT_BLUE = (41, 128, 185)
LIGHT_GRAY = (245, 247, 250)

# Title
ws.range("A1").font.size = 18
ws.range("A1").font.bold = True
ws.range("A1").font.color = WHITE
ws.range("A1:D1").color = NAVY
ws.range("A1").row_height = 38

# Table headers
ws.range("A5:D5").color = ACCENT_BLUE
ws.range("A5:D5").font.color = WHITE
ws.range("A5:D5").font.bold = True

# Alternating rows
for i in range(6, 20):
    ws.range(f"A{i}:D{i}").color = LIGHT_GRAY if i % 2 == 0 else WHITE

# STEP 4: Hide gridlines
try:
    wb.app.api.active_window.display_gridlines.set(False)
except:
    pass

# STEP 5: Borders (optional — may fail silently on some setups)
tables = ["A5:D19"]
for t in tables:
    rng = ws.range(t)
    for edge in range(7, 13):
        try:
            border = rng.api.borders[edge]
            border.line_style.set(1)
            border.weight.set(2)
            border.color.set(14277081)
        except:
            pass

# STEP 6: Save
wb.save()
print("Done")
```

### Color Palette (Standard Professional)

```python
NAVY = (23, 42, 69)         # Dark headers, title bars
WHITE = (255, 255, 255)     # White text on dark bg, clean rows
LIGHT_GRAY = (245, 247, 250)  # Alternating row shading
ACCENT_BLUE = (41, 128, 185)  # Table headers, links
RED_ALERT = (192, 57, 43)   # Warning/risk highlights
GREEN_GOOD = (39, 174, 96)  # Positive/healthy highlights
GOLD = (212, 175, 55)       # Premium accent
```

### What Works vs What Doesn't

| Operation | Works? | Notes |
|---|---|---|
| `.color = (R,G,B)` on range | Yes | Background fill |
| `.font.color = (R,G,B)` | Yes | Text color |
| `.font.bold = True` | Yes | |
| `.font.size = N` | Yes | |
| `.font.italic = True` | Yes | |
| `.font.name = "Calibri"` | Yes | |
| `.row_height = N` | Yes | |
| `.column_width = N` | Yes | |
| `.api.unmerge()` | Yes | Required before formatting merged cells |
| `.merge()` | AVOID | Causes blocking dialogs on subsequent operations |
| `.api.borders[N]` | Yes | May fail silently; wrap in try/except |
| `display_gridlines.set(False)` | Yes | Wrap in try/except |
| `.number_format` | Yes | Number formatting (e.g., "#,##0") |

### Execution Command

```bash
conda run -n base python /tmp/format_excel.py
```

Always use timeout of **60000ms** (60 seconds). xlwings formatting operations are slow because each property set is a separate AppleScript call to Excel.

## Financial Model Workflow

When building structured .xlsx/.xlsm files (budgets, forecasts, models):

1. **Open the template or existing file**: `excel_open_workbook(file_path="...")`
2. **Write all data using MCP tools**: headers, formulas, values via `excel_write_range` and `excel_write_cell`
3. **Format in one script**: Write a single Python file that applies ALL formatting (colors, widths, heights, borders, gridlines) and run it once
4. **Save**: `excel_save_workbook()`

### Complete Build + Format Workflow (Two Steps Only)

**Step 1 — Data (via MCP tools):**
- `excel_add_sheet` → create the sheet
- `excel_write_range` → write headers + data + formulas (NO merges)
- Verify with `excel_read_range`

**Step 2 — Format (via single Python script):**
- Write `/tmp/format_excel.py` with ALL formatting
- Run once with 60s timeout
- Done

This is a TWO-STEP workflow. Not three. Not five. Two.

## Interaction Pattern

The ideal workflow is conversational:
1. Kiro writes data/formulas to Excel via MCP tools (user sees values appear)
2. Kiro applies professional formatting in one script (user sees colors/styling appear)
3. User reviews and requests changes
4. Kiro updates specific cells or re-runs format script
5. Save final version

## Troubleshooting

### Script Hangs / Times Out
- **Cause**: Merged cells from a prior write. Excel shows a dialog that blocks the AppleScript bridge.
- **Fix**: Add `ws.range("A1:Z100").api.unmerge()` as the FIRST line of the formatting script.

### "Excel is not running"
- Open Excel manually or use `excel_open_workbook` with a file path.

### "Workbook not found"
- Call `excel_list_workbooks` to confirm what's open.

### Multiline Python in `-c` Flag Breaks
- **Cause**: `conda run -n base python -c "..."` flattens multiline strings.
- **Fix**: ALWAYS write to a file first (`/tmp/script.py`), then execute the file.

### AppleScript Heredocs Don't Work for Complex Operations
- **Cause**: `osascript << 'EOF'` works for simple single-property calls but times out on bulk formatting.
- **Fix**: Use xlwings Python scripts instead. They batch operations more efficiently.

### conda Entry Point Warning
- The message `Error while loading conda entry point: conda-libmamba-solver` is harmless. Ignore it. The script still executes correctly.
