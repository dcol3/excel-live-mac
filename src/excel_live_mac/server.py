"""
excel_live_mac — Live Excel MCP Server for macOS.

Interacts with RUNNING Excel instances via the xlwings AppleScript bridge.
Changes appear instantly in the open workbook.

Tools (13):
  excel_list_workbooks     - List all currently open workbooks
  excel_get_sheet_names    - Get sheet names from an open workbook
  excel_get_used_range     - Get dimensions of used range (rows, cols) without reading data
  excel_read_range         - Read a range as a 2D array
  excel_write_range        - Write a 2D array to a range (instant in Excel)
  excel_write_cell         - Write a single value or formula to a cell
  excel_read_cell          - Read a single cell value
  excel_new_workbook       - Create a new blank workbook
  excel_open_workbook      - Open a workbook file (launches Excel if needed)
  excel_save_workbook      - Save the active or specified workbook
  excel_add_sheet          - Add a new sheet to a workbook
  excel_run_macro          - Execute a VBA macro
  excel_get_selection      - Get the currently selected range info and values

Requirements:
  - macOS with Microsoft Excel installed
  - pip install mcp xlwings
"""

import asyncio
import json
import sys
from typing import Any

try:
    import xlwings as xw
except ImportError:
    print("ERROR: xlwings not installed. Run: pip install xlwings", file=sys.stderr)
    sys.exit(1)

try:
    from mcp.server import Server
    from mcp.server.stdio import stdio_server
    from mcp import types
except ImportError:
    print("ERROR: mcp package not installed. Run: pip install mcp", file=sys.stderr)
    sys.exit(1)

server = Server("excel-live-mac")


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _get_app() -> xw.App:
    """Get the running Excel app instance, or raise a clear error."""
    try:
        apps = xw.apps
        if not apps:
            raise RuntimeError("Excel is not running. Open Excel first or use excel_open_workbook.")
        return apps.active
    except Exception as e:
        raise RuntimeError(f"Cannot connect to Excel: {e}. Make sure Excel is running on your Mac.")


def _get_workbook(name: str = "") -> xw.Book:
    """Get a workbook by name, or the active one if name is empty."""
    app = _get_app()
    if not name:
        if not app.books:
            raise RuntimeError("No workbooks open in Excel.")
        return app.books.active
    for book in app.books:
        if book.name == name or book.name.lower() == name.lower():
            return book
    for book in app.books:
        if name.lower() in book.name.lower():
            return book
    raise RuntimeError(f"Workbook '{name}' not found. Open workbooks: {[b.name for b in app.books]}")


def _serialize_value(val):
    """Make cell values JSON-serializable."""
    if val is None:
        return None
    if isinstance(val, (int, float, str, bool)):
        return val
    return str(val)


def _serialize_range(data):
    """Serialize range data (could be single value, list, or list of lists)."""
    if data is None:
        return None
    if isinstance(data, (int, float, str, bool)):
        return data
    if isinstance(data, list):
        if data and isinstance(data[0], list):
            return [[_serialize_value(cell) for cell in row] for row in data]
        return [_serialize_value(cell) for cell in data]
    return str(data)


# ─────────────────────────────────────────────────────────────────────────────
# Tool list
# ─────────────────────────────────────────────────────────────────────────────

@server.list_tools()
async def list_tools() -> list:
    return [
        types.Tool(
            name="excel_list_workbooks",
            description="List all currently open workbooks in Excel. Shows name, path, and sheet count.",
            inputSchema={
                "type": "object",
                "properties": {},
                "required": []
            }
        ),
        types.Tool(
            name="excel_get_sheet_names",
            description="Get all sheet names from an open workbook.",
            inputSchema={
                "type": "object",
                "properties": {
                    "workbook": {"type": "string", "description": "Workbook name (or empty for active)"}
                },
                "required": []
            }
        ),
        types.Tool(
            name="excel_get_used_range",
            description="Get the dimensions of the used range in a sheet without reading cell data. Returns row count, column count, first cell, and last cell address. Use this to understand data size before reading.",
            inputSchema={
                "type": "object",
                "properties": {
                    "workbook": {"type": "string", "description": "Workbook name (or empty for active)"},
                    "sheet": {"type": "string", "description": "Sheet name (or empty for active sheet)"}
                },
                "required": []
            }
        ),
        types.Tool(
            name="excel_read_range",
            description="Read a range from an open workbook. Returns cell values as a 2D array.",
            inputSchema={
                "type": "object",
                "properties": {
                    "workbook": {"type": "string", "description": "Workbook name (or empty for active)"},
                    "sheet": {"type": "string", "description": "Sheet name (or empty for active sheet)"},
                    "range": {"type": "string", "description": "Cell range like 'A1:D20', 'B:B', or 'A1'"},
                    "expand": {"type": "boolean", "default": False, "description": "Auto-expand to contiguous data region from top-left cell"}
                },
                "required": ["range"]
            }
        ),
        types.Tool(
            name="excel_write_range",
            description="Write a 2D array of data to a range in an open workbook. Changes appear immediately in Excel.",
            inputSchema={
                "type": "object",
                "properties": {
                    "workbook": {"type": "string", "description": "Workbook name (or empty for active)"},
                    "sheet": {"type": "string", "description": "Sheet name (or empty for active sheet)"},
                    "range": {"type": "string", "description": "Top-left cell to start writing, e.g. 'A1' or 'B5'"},
                    "data": {"type": "array", "description": "2D array of values (rows of columns)"},
                },
                "required": ["range", "data"]
            }
        ),
        types.Tool(
            name="excel_write_cell",
            description="Write a single value to a specific cell. Change appears immediately.",
            inputSchema={
                "type": "object",
                "properties": {
                    "workbook": {"type": "string"},
                    "sheet": {"type": "string"},
                    "cell": {"type": "string", "description": "Cell address like 'B5'"},
                    "value": {"description": "Value to write (string, number, or formula starting with =)"},
                },
                "required": ["cell", "value"]
            }
        ),
        types.Tool(
            name="excel_read_cell",
            description="Read the value of a single cell.",
            inputSchema={
                "type": "object",
                "properties": {
                    "workbook": {"type": "string"},
                    "sheet": {"type": "string"},
                    "cell": {"type": "string", "description": "Cell address like 'B5'"},
                },
                "required": ["cell"]
            }
        ),
        types.Tool(
            name="excel_open_workbook",
            description="Open a workbook file in Excel. Launches Excel if not running.",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {"type": "string", "description": "Absolute path to .xlsx/.xlsm/.xls file"}
                },
                "required": ["file_path"]
            }
        ),
        types.Tool(
            name="excel_new_workbook",
            description="Create a new blank workbook in Excel. Returns the workbook name. Save it with excel_save_workbook to persist to disk.",
            inputSchema={
                "type": "object",
                "properties": {},
                "required": []
            }
        ),
        types.Tool(
            name="excel_add_sheet",
            description="Add a new sheet to a workbook. Use this when you need a separate tab for a dashboard or summary. Does NOT work on CSV files — save as .xlsx first.",
            inputSchema={
                "type": "object",
                "properties": {
                    "workbook": {"type": "string", "description": "Workbook name (or empty for active)"},
                    "name": {"type": "string", "description": "Name for the new sheet"},
                    "after": {"type": "string", "description": "Name of existing sheet to place after (or empty for end)"}
                },
                "required": ["name"]
            }
        ),
        types.Tool(
            name="excel_save_workbook",
            description="Save the specified (or active) workbook.",
            inputSchema={
                "type": "object",
                "properties": {
                    "workbook": {"type": "string", "description": "Workbook name (or empty for active)"},
                    "path": {"type": "string", "description": "Save-as path (empty = save in place)"},
                },
                "required": []
            }
        ),
        types.Tool(
            name="excel_run_macro",
            description="Run a VBA macro in the active workbook. Macro must already exist in the workbook.",
            inputSchema={
                "type": "object",
                "properties": {
                    "workbook": {"type": "string"},
                    "macro": {"type": "string", "description": "Macro name (e.g. 'Sheet1.MyMacro' or just 'MyMacro')"},
                    "args": {"type": "array", "description": "Arguments to pass to the macro", "default": []},
                },
                "required": ["macro"]
            }
        ),
        types.Tool(
            name="excel_get_selection",
            description="Get info about the currently selected range in Excel (address, values, sheet).",
            inputSchema={
                "type": "object",
                "properties": {},
                "required": []
            }
        ),
    ]


# ─────────────────────────────────────────────────────────────────────────────
# Tool implementations
# ─────────────────────────────────────────────────────────────────────────────

@server.call_tool()
async def call_tool(name: str, arguments: dict[str, Any]) -> list:

    try:
        if name == "excel_list_workbooks":
            app = _get_app()
            books = []
            for book in app.books:
                books.append({
                    "name": book.name,
                    "path": book.fullname,
                    "sheets": len(book.sheets),
                    "sheet_names": [s.name for s in book.sheets],
                })
            return [types.TextContent(type="text", text=json.dumps({"workbooks": books}, indent=2))]

        elif name == "excel_get_sheet_names":
            book = _get_workbook(arguments.get("workbook", ""))
            sheets = [{"name": s.name, "index": i + 1} for i, s in enumerate(book.sheets)]
            return [types.TextContent(type="text", text=json.dumps({"workbook": book.name, "sheets": sheets}, indent=2))]

        elif name == "excel_get_used_range":
            book = _get_workbook(arguments.get("workbook", ""))
            sheet_name = arguments.get("sheet", "")
            ws = book.sheets[sheet_name] if sheet_name else book.sheets.active
            used = ws.used_range
            return [types.TextContent(type="text", text=json.dumps({
                "workbook": book.name,
                "sheet": ws.name,
                "rows": used.shape[0],
                "cols": used.shape[1],
                "address": used.address,
                "last_cell": used.last_cell.address,
            }, indent=2))]

        elif name == "excel_read_range":
            book = _get_workbook(arguments.get("workbook", ""))
            sheet_name = arguments.get("sheet", "")
            ws = book.sheets[sheet_name] if sheet_name else book.sheets.active
            rng = ws.range(arguments["range"])
            if arguments.get("expand", False):
                rng = rng.expand()
            data = rng.value
            result = {
                "workbook": book.name,
                "sheet": ws.name,
                "range": rng.address,
                "rows": rng.shape[0] if hasattr(rng, 'shape') else 1,
                "cols": rng.shape[1] if hasattr(rng, 'shape') and len(rng.shape) > 1 else 1,
                "data": _serialize_range(data),
            }
            return [types.TextContent(type="text", text=json.dumps(result, indent=2))]

        elif name == "excel_write_range":
            book = _get_workbook(arguments.get("workbook", ""))
            sheet_name = arguments.get("sheet", "")
            ws = book.sheets[sheet_name] if sheet_name else book.sheets.active
            rng = ws.range(arguments["range"])
            data = arguments["data"]
            # Auto-pad rows to uniform length (fixes "must be same length" errors)
            if data and isinstance(data[0], list):
                max_cols = max(len(row) for row in data)
                data = [row + [None] * (max_cols - len(row)) for row in data]
            rng.value = data
            return [types.TextContent(type="text", text=json.dumps({
                "status": "written",
                "workbook": book.name,
                "sheet": ws.name,
                "range": arguments["range"],
                "rows_written": len(data) if isinstance(data, list) else 1,
            }, indent=2))]

        elif name == "excel_write_cell":
            book = _get_workbook(arguments.get("workbook", ""))
            sheet_name = arguments.get("sheet", "")
            ws = book.sheets[sheet_name] if sheet_name else book.sheets.active
            cell = ws.range(arguments["cell"])
            value = arguments["value"]
            if isinstance(value, str) and value.startswith("="):
                cell.formula = value
            else:
                cell.value = value
            return [types.TextContent(type="text", text=json.dumps({
                "status": "written",
                "cell": arguments["cell"],
                "value": str(value),
                "workbook": book.name,
                "sheet": ws.name,
            }, indent=2))]

        elif name == "excel_read_cell":
            book = _get_workbook(arguments.get("workbook", ""))
            sheet_name = arguments.get("sheet", "")
            ws = book.sheets[sheet_name] if sheet_name else book.sheets.active
            cell = ws.range(arguments["cell"])
            return [types.TextContent(type="text", text=json.dumps({
                "cell": arguments["cell"],
                "value": _serialize_value(cell.value),
                "formula": cell.formula if cell.formula else None,
                "workbook": book.name,
                "sheet": ws.name,
            }, indent=2))]

        elif name == "excel_open_workbook":
            file_path = arguments["file_path"]
            book = xw.Book(file_path)
            return [types.TextContent(type="text", text=json.dumps({
                "status": "opened",
                "name": book.name,
                "path": book.fullname,
                "sheets": [s.name for s in book.sheets],
            }, indent=2))]

        elif name == "excel_new_workbook":
            book = xw.Book()
            return [types.TextContent(type="text", text=json.dumps({
                "status": "created",
                "name": book.name,
                "sheets": [s.name for s in book.sheets],
                "note": "Use excel_save_workbook with a path to save to disk."
            }, indent=2))]

        elif name == "excel_add_sheet":
            book = _get_workbook(arguments.get("workbook", ""))
            sheet_name = arguments["name"]
            after = arguments.get("after", "")
            if after:
                new_sheet = book.sheets.add(sheet_name, after=book.sheets[after])
            else:
                new_sheet = book.sheets.add(sheet_name, after=book.sheets[-1])
            return [types.TextContent(type="text", text=json.dumps({
                "status": "created",
                "sheet": new_sheet.name,
                "workbook": book.name,
                "all_sheets": [s.name for s in book.sheets],
            }, indent=2))]

        elif name == "excel_save_workbook":
            book = _get_workbook(arguments.get("workbook", ""))
            save_path = arguments.get("path", "")
            if save_path:
                book.save(save_path)
            else:
                book.save()
            return [types.TextContent(type="text", text=json.dumps({
                "status": "saved",
                "name": book.name,
                "path": book.fullname,
            }, indent=2))]

        elif name == "excel_run_macro":
            book = _get_workbook(arguments.get("workbook", ""))
            macro_name = arguments["macro"]
            args = arguments.get("args", [])
            macro = book.macro(macro_name)
            result = macro(*args) if args else macro()
            return [types.TextContent(type="text", text=json.dumps({
                "status": "executed",
                "macro": macro_name,
                "result": _serialize_value(result),
                "workbook": book.name,
            }, indent=2))]

        elif name == "excel_get_selection":
            app = _get_app()
            sel = app.selection
            if sel is None:
                return [types.TextContent(type="text", text=json.dumps({"error": "No selection active"}))]
            data = sel.value
            return [types.TextContent(type="text", text=json.dumps({
                "address": sel.address,
                "sheet": sel.sheet.name,
                "workbook": sel.sheet.book.name,
                "shape": list(sel.shape) if hasattr(sel, 'shape') else [1, 1],
                "data": _serialize_range(data),
            }, indent=2))]

        else:
            return [types.TextContent(type="text", text=f"Unknown tool: {name}")]

    except Exception as e:
        return [types.TextContent(type="text", text=json.dumps({"error": str(e)}, indent=2))]


# ─────────────────────────────────────────────────────────────────────────────
# Entry points
# ─────────────────────────────────────────────────────────────────────────────

async def main():
    """Async entry point for the MCP server."""
    async with stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream, write_stream,
            server.create_initialization_options()
        )


def main_sync():
    """Synchronous entry point (called by `excel-live-mac` CLI command)."""
    asyncio.run(main())


if __name__ == "__main__":
    main_sync()
