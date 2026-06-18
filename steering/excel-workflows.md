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

### Before Reading Large Data
ALWAYS call `excel_get_used_range` first to learn the dimensions. Do NOT probe row counts by reading A500, A1000, A5000, etc. One call tells you exactly how many rows and columns exist.

### CSV Files Are Single-Sheet
CSV files opened in Excel can only have one sheet. You CANNOT add sheets with `excel_add_sheet`. If you need multiple sheets:
1. Create a new workbook with `excel_new_workbook`
2. Write your dashboard/summary to the new workbook
3. Reference data from the CSV workbook using cross-workbook formulas like `='[filename.csv]Sheet1'!A1:D100`

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

## Formatting Guidance

The xlwings bridge does not directly expose cell formatting (colors, borders, column widths) through these MCP tools. For formatting:
- Use VBA macros if the workbook has them (`excel_run_macro`)
- For new formatted workbooks, use the file-based `excel-tools` `build_excel_shell` tool to create the formatted structure first, then open with `excel_open_workbook` and populate with `excel_write_range`

## Financial Model Workflow

When building structured .xlsx/.xlsm files (budgets, forecasts, models):

1. **Open the template or existing file**: `excel_open_workbook(file_path="...")`
2. **Write Assumptions sheet**: key inputs, rates, date ranges, sources
3. **Write Calculations sheet**: formulas that reference assumption cells
4. **Write Summary sheet**: headline outputs, charts data, key metrics
5. **Write Data sheet**: raw source data for auditability
6. **Add formulas**: link cells between sheets using `excel_write_cell`
7. **Save**: `excel_save_workbook()`

This creates a live, interactive experience where the user watches the file build in real time.

## Interaction Pattern

The ideal workflow is conversational:
1. Kiro computes or pulls data
2. Kiro writes to Excel (user sees it instantly)
3. User reviews and requests changes
4. Kiro updates specific cells
5. Repeat until complete
6. Save final version
