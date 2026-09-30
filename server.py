r"""VERTEX SAP GUI MCP: low-level tools for an already open SAP GUI session.

The server uses stdio and is intended to be started by an MCP client.
It never logs in to SAP and intentionally provides no tool named save, post,
delete, or execute arbitrary code.
"""

import re
import time
import os
import threading
from contextlib import contextmanager
from typing import Any

import pythoncom
import win32com.client
from mcp.server.fastmcp import FastMCP


mcp = FastMCP("VERTEX SAP GUI")

# SAP GUI exposes its automation object through the Windows Running Object
# Table.  On a first connection SAP may display its scripting consent dialog;
# while the dialog is open the object is not always immediately usable.
CONNECT_TIMEOUT_SECONDS = float(os.getenv("VERTEX_SAP_CONNECT_TIMEOUT", "25"))
CONNECT_RETRY_SECONDS = 0.5
ACTIONS_ENABLED = os.getenv("VERTEX_SAP_ENABLE_ACTIONS", "0") == "1"
SESSION_READY_TIMEOUT_SECONDS = float(os.getenv("VERTEX_SAP_READY_TIMEOUT", "20"))
LAYOUT_CACHE_SECONDS = float(os.getenv("VERTEX_SAP_LAYOUT_CACHE_SECONDS", "300"))
_sap_lock = threading.RLock()
_com_state = threading.local()
# Cache only a screen's static field schema (names, labels and editability),
# never values, status messages, or employee data.
_layout_cache: dict[tuple[str, ...], tuple[float, list[dict[str, Any]]]] = {}


@contextmanager
def _sap_call():
    """Serialize SAP access using a persistent COM apartment per MCP worker.

    SAP GUI Scripting is attached once and then reused while the MCP process
    lives.  Reinitialising and uninitialising COM for every tool call can make
    SAP treat each call as a new script attachment and show a warning dialog.
    """
    if not getattr(_com_state, "initialized", False):
        pythoncom.CoInitialize()
        _com_state.initialized = True
    with _sap_lock:
        yield


def _value(component: Any, name: str, default: Any = "") -> Any:
    """Read a COM property without failing the whole screen inspection."""
    try:
        return getattr(component, name)
    except Exception:
        return default


def _walk(component: Any):
    yield component
    try:
        children = component.Children
        for index in range(children.Count):
            yield from _walk(children(index))
    except Exception:
        return


def _application() -> Any:
    cached = getattr(_com_state, "application", None)
    if cached is not None:
        try:
            # Touch the object so a closed SAP GUI does not leave a stale proxy.
            _ = cached.Children.Count
            return cached
        except Exception:
            _com_state.application = None

    deadline = time.monotonic() + CONNECT_TIMEOUT_SECONDS
    last_error: Exception | None = None
    while True:
        try:
            sap_gui = win32com.client.GetObject("SAPGUI")
            application = sap_gui.GetScriptingEngine
            _com_state.application = application
            return application
        except Exception as error:
            last_error = error
            if time.monotonic() >= deadline:
                break
            time.sleep(CONNECT_RETRY_SECONDS)
    raise RuntimeError(
        "SAP GUI Scripting did not become available within "
        f"{CONNECT_TIMEOUT_SECONDS:g} seconds. Open SAP GUI, sign in, "
        "and confirm the Scripting consent dialog. "
        f"Last COM error: {last_error}"
    ) from last_error


def _available_sessions(application: Any | None = None) -> list[Any]:
    application = application or _application()
    sessions: list[Any] = []
    for connection_index in range(application.Children.Count):
        connection = application.Children(connection_index)
        for session_index in range(connection.Children.Count):
            sessions.append(connection.Children(session_index))
    try:
        active = application.ActiveSession
        if active and all(_value(item, "Id") != _value(active, "Id") for item in sessions):
            sessions.append(active)
    except Exception:
        pass
    return sessions


def _session(session_id: str | None = None, require_explicit: bool = False) -> Any:
    sessions = _available_sessions()
    if session_id:
        for session in sessions:
            if _value(session, "Id", "") == session_id:
                return session
        raise ValueError("The specified SAP session was not found. Call sap_list_sessions.")
    if not sessions:
        raise RuntimeError("No working SAP session found.")
    if require_explicit and len(sessions) > 1:
        raise ValueError("Multiple SAP sessions found. Provide session_id from sap_list_sessions.")
    return sessions[0]


def _control(session: Any, control_id: str) -> Any:
    if not control_id.startswith("wnd["):
        raise ValueError("control_id must start with 'wnd['.")
    try:
        return session.findById(control_id)
    except Exception as error:
        raise RuntimeError(f"Control not found: {control_id}") from error


def _editable_text_control_by_name(session: Any, name: str) -> Any:
    """Find the visible editable PA30 field identified by its SAP technical name."""
    fallback = None
    for component in _walk(session):
        if _value(component, "Name", "") != name:
            continue
        if _value(component, "Type", "") not in {
            "GuiCTextField",
            "GuiTextField",
            "GuiComboBox",
        }:
            continue
        if not _value(component, "Changeable", False):
            continue
        if _value(component, "Visible", True):
            return component
        fallback = fallback or component
    if fallback is not None:
        return fallback
    raise RuntimeError(f"Field {name} was not found or is not editable.")


def _visible_control_by_name(session: Any, name: str) -> Any:
    """Return the first visible control with an exact SAP technical name."""
    for component in _walk(session):
        if _value(component, "Name", "") == name and _value(component, "Visible", True):
            return component
    raise RuntimeError(f"Visible control {name} not found.")


def _select_pa30_infotype_title(session: Any, title: str) -> bool:
    """Select a visible PA30 infotype-menu row by its displayed title.

    PA30's direct-selection field is not consistent across localizations and
    customer menus: it can keep the old selected row even after setting its
    text. Selecting the matching table row is deterministic and does not rely
    on a built-in title-to-number mapping.
    """
    expected = title.casefold()
    for component in _walk(session):
        if _value(component, "Text", "").strip().casefold() != expected:
            continue
        control_id = _value(component, "Id", "")
        match = re.search(r"/tbl[^/]+/[^/]+\[\d+,(\d+)\]$", control_id)
        if not match:
            continue
        table_id = control_id.rsplit("/", 1)[0]
        table = session.findById(table_id)
        absolute_row = int(match.group(1)) + int(_value(_value(table, "VerticalScrollbar", None), "Position", 0))
        table.getAbsoluteRow(absolute_row).Selected = True
        return True
    return False


def _wait_until_ready(session: Any) -> None:
    """Wait only while SAP is processing a server round trip."""
    deadline = time.monotonic() + SESSION_READY_TIMEOUT_SECONDS
    while _value(session, "Busy", False):
        if time.monotonic() >= deadline:
            raise TimeoutError("SAP did not finish processing within the allotted time.")
        time.sleep(0.05)


def _open_infotype(session: Any, infotype: str, mode: str) -> dict[str, Any]:
    """Open a PA30 infotype in an already selected session without saving it."""
    selection = infotype.strip()
    if not selection or len(selection) > 80:
        raise ValueError("infotype must contain 1 to 80 characters.")
    function_codes = {"display": "=DIS", "change": "=MOD", "create": "=INS"}
    normalized_mode = mode.strip().lower()
    if normalized_mode not in function_codes:
        raise ValueError("mode must be display, change, or create.")
    if _value(_value(session, "Info", None), "Transaction", "").upper() != "PA30":
        raise ValueError("sap_open_infotype is only available in transaction PA30.")
    if re.fullmatch(r"\d{4}", selection):
        _editable_text_control_by_name(session, "RP50G-CHOIC").text = selection
    elif not _select_pa30_infotype_title(session, selection):
        raise ValueError(
            "No infotype with that title was found among visible PA30 rows. "
            "Open the required tab or provide a four-digit number."
        )
    command_field = session.findById("wnd[0]/tbar[0]/okcd")
    command_field.text = function_codes[normalized_mode]
    session.findById("wnd[0]").sendVKey(0)
    _wait_until_ready(session)
    return {"infotype": selection, "mode": normalized_mode, "saved": False, **_screen_summary(session)}


def _visible_fields(session: Any, max_fields: int = 100) -> list[dict[str, Any]]:
    """Return fields and their labels in one focused pass over the current dynpro."""
    labels: dict[str, str] = {}
    fields: list[dict[str, Any]] = []
    field_types = {"GuiCTextField", "GuiTextField", "GuiComboBox", "GuiCheckBox", "GuiRadioButton"}
    for component in _walk(session):
        if not _value(component, "Visible", True):
            continue
        name = _value(component, "Name", "")
        component_type = _value(component, "Type", "")
        if component_type == "GuiLabel" and name:
            labels[name] = _value(component, "Text", "").strip()
        elif component_type in field_types and name and len(fields) < max_fields:
            fields.append({
                "name": name,
                "type": component_type,
                "label": "",
                "changeable": bool(_value(component, "Changeable", False)),
            })
    for field in fields:
        field["label"] = labels.get(field["name"], "")
    return fields


def _layout_cache_key(session: Any) -> tuple[str, ...]:
    info = _value(session, "Info", None)
    main_window = session.findById("wnd[0]")
    return (
        _value(info, "SystemName", ""),
        _value(info, "Client", ""),
        _value(info, "Transaction", ""),
        _value(info, "Program", ""),
        str(_value(info, "ScreenNumber", "")),
        _value(main_window, "Text", ""),
    )


def _cached_visible_fields(session: Any, max_fields: int) -> tuple[list[dict[str, Any]], bool]:
    """Read a dynpro schema once and reuse it while the screen remains known."""
    if LAYOUT_CACHE_SECONDS <= 0:
        return _visible_fields(session, max_fields), False
    key = _layout_cache_key(session)
    cached = _layout_cache.get(key)
    now = time.monotonic()
    if cached and now - cached[0] < LAYOUT_CACHE_SECONDS:
        return [dict(field) for field in cached[1][:max_fields]], True
    fields = _visible_fields(session, 200)
    _layout_cache[key] = (now, [dict(field) for field in fields])
    return fields[:max_fields], False


def _require_actions_enabled() -> None:
    if not ACTIONS_ENABLED:
        raise PermissionError(
            "Changing actions are disabled. To enable them, set "
            "VERTEX_SAP_ENABLE_ACTIONS=1 in the MCP configuration and restart the server."
        )


def _screen_summary(session: Any) -> dict[str, Any]:
    main_window = session.findById("wnd[0]")
    status_bar = session.findById("wnd[0]/sbar")
    return {
        "transaction": _value(_value(session, "Info", None), "Transaction", ""),
        "title": _value(main_window, "Text", ""),
        "status": {
            "text": _value(status_bar, "Text", ""),
            "type": _value(status_bar, "MessageType", ""),
        },
    }


@mcp.tool()
def sap_list_sessions() -> list[dict[str, Any]]:
    """List all available SAP sessions so a caller can select the correct system and client."""
    with _sap_call():
        application = _application()
        try:
            active_id = _value(application.ActiveSession, "Id", "")
        except Exception:
            active_id = ""
        result = []
        for session in _available_sessions(application):
            result.append({
                "session_id": _value(session, "Id", ""),
                "system": _value(_value(session, "Info", None), "SystemName", ""),
                "client": _value(_value(session, "Info", None), "Client", ""),
                "user": _value(_value(session, "Info", None), "User", ""),
                "transaction": _value(_value(session, "Info", None), "Transaction", ""),
                "active": _value(session, "Id", "") == active_id,
            })
        return result


@mcp.tool()
def sap_status(session_id: str | None = None) -> dict[str, Any]:
    """Return the currently selected SAP session and its status bar message."""
    with _sap_call():
        session = _session(session_id)
        return {
            "session_id": _value(session, "Id", ""),
            "system": _value(_value(session, "Info", None), "SystemName", ""),
            "client": _value(_value(session, "Info", None), "Client", ""),
            "user": _value(_value(session, "Info", None), "User", ""),
            **_screen_summary(session),
        }


@mcp.tool()
def sap_clear_layout_cache() -> dict[str, int]:
    """Clear cached static screen schemas; no SAP data is changed."""
    with _sap_call():
        count = len(_layout_cache)
        _layout_cache.clear()
        return {"cleared_entries": count}


@mcp.tool()
def sap_get_screen(session_id: str | None = None, max_controls: int = 120) -> dict[str, Any]:
    """Return basic screen metadata plus visible controls in the active SAP session."""
    if not 1 <= max_controls <= 500:
        raise ValueError("max_controls must be between 1 and 500.")

    with _sap_call():
        session = _session(session_id)
        controls: list[dict[str, Any]] = []
        for component in _walk(session):
            if len(controls) >= max_controls:
                break
            if not _value(component, "Visible", True):
                continue
            control_id = _value(component, "Id", "")
            if not control_id:
                continue
            controls.append(
                {
                    "id": control_id,
                    "name": _value(component, "Name", ""),
                    "type": _value(component, "Type", ""),
                    "text": _value(component, "Text", ""),
                    "tooltip": _value(component, "Tooltip", ""),
                    "changeable": bool(_value(component, "Changeable", False)),
                }
            )
        return {**_screen_summary(session), "controls": controls}


@mcp.tool()
def sap_find_controls(query: str, session_id: str | None = None, max_results: int = 30) -> list[dict[str, Any]]:
    """Find visible SAP controls by a case-insensitive match on ID, name, text, or tooltip."""
    if not query.strip():
        raise ValueError("query must not be empty.")
    if not 1 <= max_results <= 100:
        raise ValueError("max_results must be between 1 and 100.")

    with _sap_call():
        needle = query.casefold()
        results: list[dict[str, Any]] = []
        for component in _walk(_session(session_id)):
            if len(results) >= max_results or not _value(component, "Visible", True):
                continue
            item = {
                "id": _value(component, "Id", ""),
                "name": _value(component, "Name", ""),
                "type": _value(component, "Type", ""),
                "text": _value(component, "Text", ""),
                "tooltip": _value(component, "Tooltip", ""),
                "changeable": bool(_value(component, "Changeable", False)),
            }
            searchable = " ".join(str(value) for value in item.values())
            if needle in searchable.casefold():
                results.append(item)
        return results


@mcp.tool()
def sap_read_fields(control_ids: list[str], session_id: str | None = None) -> list[dict[str, Any]]:
    """Read explicitly identified SAP fields without traversing the screen tree."""
    if not 1 <= len(control_ids) <= 50:
        raise ValueError("control_ids must contain 1 to 50 identifiers.")
    with _sap_call():
        session = _session(session_id)
        fields: list[dict[str, Any]] = []
        for control_id in control_ids:
            control = _control(session, control_id)
            fields.append(
                {
                    "id": control_id,
                    "name": _value(control, "Name", ""),
                    "type": _value(control, "Type", ""),
                    "text": _value(control, "Text", ""),
                    "changeable": bool(_value(control, "Changeable", False)),
                }
            )
        return fields


def _grid_columns(grid: Any) -> list[str]:
    """Convert the SAP GUI grid's COM column collection to plain strings."""
    columns = _value(grid, "ColumnOrder", ())
    try:
        return [str(column) for column in columns]
    except TypeError as error:
        raise RuntimeError("Could not read the SAP Grid column list.") from error


def _activate_grid_row(session: Any, grid: Any, row: int, column: str | None = None) -> None:
    if row >= int(_value(grid, "RowCount", 0)):
        raise ValueError("The specified row does not exist in the SAP Grid.")
    columns = _grid_columns(grid)
    if column is not None and column not in columns:
        raise ValueError("The specified column does not exist in the SAP Grid.")
    current_column = column or (columns[0] if columns else "")
    if not current_column:
        raise RuntimeError("No columns found in the SAP Grid.")

    try:
        grid.FirstVisibleRow = row
    except Exception:
        pass
    try:
        grid.SelectedRows = str(row)
    except Exception:
        pass
    try:
        grid.CurrentCellRow = row
        grid.CurrentCellColumn = current_column
        grid.CurrentCellMoved()
    except Exception:
        pass

    errors: list[str] = []
    for method_name, args in (
        ("DoubleClick", (row, current_column)),
        ("doubleClick", (row, current_column)),
        ("DoubleClickCurrentCell", ()),
        ("doubleClickCurrentCell", ()),
    ):
        method = _value(grid, method_name, None)
        if method is None:
            continue
        try:
            method(*args)
            _wait_until_ready(session)
            return
        except Exception as error:
            errors.append(f"{method_name}: {error}")

    for key in (2, 0):
        try:
            session.findById("wnd[0]").sendVKey(key)
            _wait_until_ready(session)
            return
        except Exception as error:
            errors.append(f"sendVKey({key}): {error}")

    detail = "; ".join(errors) if errors else "row activation methods are unavailable"
    raise RuntimeError(f"Could not activate the SAP Grid row: {detail}")


@mcp.tool()
def sap_read_grid_rows(
    grid_id: str,
    columns: list[str] | None = None,
    max_rows: int = 100,
    session_id: str | None = None,
) -> dict[str, Any]:
    """Read rows directly from a visible SAP GUI grid without opening records.

    Call once without ``columns`` to discover technical column names, then pass
    only the needed names (for example ``PERNR`` and ``ENAME``) on later calls.
    """
    if not 1 <= max_rows <= 500:
        raise ValueError("max_rows must be between 1 and 500.")
    with _sap_call():
        session = _session(session_id)
        grid = _control(session, grid_id)
        available_columns = _grid_columns(grid)
        selected_columns = columns or available_columns
        if not selected_columns:
            raise RuntimeError("No columns found in the SAP Grid.")
        unknown = sorted(set(selected_columns) - set(available_columns))
        if unknown:
            raise ValueError(f"Columns not found in the SAP Grid: {', '.join(unknown)}.")
        row_count = min(int(_value(grid, "RowCount", 0)), max_rows)
        rows = [
            {
                column: _value(grid, "GetCellValue")(row, column)
                for column in selected_columns
            }
            for row in range(row_count)
        ]
        return {
            "grid_id": grid_id,
            "columns": available_columns,
            "row_count": int(_value(grid, "RowCount", 0)),
            "truncated": int(_value(grid, "RowCount", 0)) > max_rows,
            "rows": rows,
        }


@mcp.tool()
def sap_select_grid_row(
    grid_id: str,
    row: int,
    column: str | None = None,
    session_id: str | None = None,
) -> dict[str, Any]:
    """Select a visible SAP GUI grid row without saving data."""
    _require_actions_enabled()
    if row < 0:
        raise ValueError("row must not be negative.")
    with _sap_call():
        session = _session(session_id, require_explicit=True)
        grid = _control(session, grid_id)
        _activate_grid_row(session, grid, row, column)
        return _screen_summary(session)


@mcp.tool()
def sap_read_grid_infotype_fields(
    grid_id: str,
    field_ids: list[str],
    columns: list[str] | None = None,
    max_rows: int = 100,
    session_id: str | None = None,
) -> dict[str, Any]:
    """Activate each visible grid row and read known infotype field ids."""
    _require_actions_enabled()
    if not 1 <= max_rows <= 100:
        raise ValueError("max_rows must be between 1 and 100.")
    if not 1 <= len(field_ids) <= 20:
        raise ValueError("field_ids must contain 1 to 20 fields.")
    with _sap_call():
        session = _session(session_id, require_explicit=True)
        grid = _control(session, grid_id)
        available_columns = _grid_columns(grid)
        selected_columns = columns or available_columns
        unknown = sorted(set(selected_columns) - set(available_columns))
        if unknown:
            raise ValueError(f"Columns not found in the SAP Grid: {', '.join(unknown)}.")
        row_count = min(int(_value(grid, "RowCount", 0)), max_rows)
        total_start = time.perf_counter()
        results = []
        for row in range(row_count):
            row_start = time.perf_counter()
            grid_values = {
                column: _value(grid, "GetCellValue")(row, column)
                for column in selected_columns
            }
            _activate_grid_row(session, grid, row, selected_columns[0] if selected_columns else None)
            fields = {
                field_id: _value(_control(session, field_id), "Text", "")
                for field_id in field_ids
            }
            results.append(
                {
                    "row": row,
                    "grid": grid_values,
                    "fields": fields,
                    "duration_ms": round((time.perf_counter() - row_start) * 1000),
                }
            )
        return {
            "grid_id": grid_id,
            "row_count": int(_value(grid, "RowCount", 0)),
            "truncated": int(_value(grid, "RowCount", 0)) > max_rows,
            "total_ms": round((time.perf_counter() - total_start) * 1000),
            "results": results,
        }


@mcp.tool()
def sap_open_transaction(code: str, session_id: str | None = None) -> dict[str, Any]:
    """Open a SAP transaction when actions are explicitly enabled."""
    _require_actions_enabled()
    normalized = code.strip().upper()
    if not re.fullmatch(r"/?[A-Z0-9_]{2,30}(?:/[A-Z0-9_]{1,30})*", normalized):
        raise ValueError("Invalid transaction code.")

    with _sap_call():
        session = _session(session_id, require_explicit=True)
        session.StartTransaction(normalized)
        _wait_until_ready(session)
        return _screen_summary(session)


@mcp.tool()
def sap_open_infotype(
    infotype: str,
    mode: str = "display",
    session_id: str | None = None,
) -> dict[str, Any]:
    """Open a PA30 infotype by its number or visible name, without saving data.

    Examples: ``0004`` or ``Challenge``.  A title is matched to the visible
    PA30 menu row, so localized and customer-specific infotype titles are
    supported without a static title-to-number mapping.  ``mode`` may be
    ``display``, ``change``, or ``create``; opening an infotype never saves it.
    """
    _require_actions_enabled()
    with _sap_call():
        session = _session(session_id, require_explicit=True)
        return _open_infotype(session, infotype, mode)


@mcp.tool()
def sap_inspect_infotype(
    infotype: str,
    mode: str = "display",
    session_id: str | None = None,
    max_fields: int = 100,
) -> dict[str, Any]:
    """Open a PA30 infotype and return its visible field names in one MCP call.

    The screen is opened but never saved. Use ``mode=create`` to inspect a
    form when no existing record is available to display or change.
    """
    _require_actions_enabled()
    if not 1 <= max_fields <= 200:
        raise ValueError("max_fields must be between 1 and 200.")
    with _sap_call():
        session = _session(session_id, require_explicit=True)
        result = _open_infotype(session, infotype, mode)
        fields, cache_hit = _cached_visible_fields(session, max_fields)
        return {**result, "fields": fields, "layout_cache_hit": cache_hit}


@mcp.tool()
def sap_read_personnel_names(
    personnel_numbers: list[str],
    session_id: str | None = None,
) -> dict[str, Any]:
    """Read IT0002 names in one optimized PA20 batch.

    PA20 and IT0002 are opened once.  Each next employee is reached by Back,
    changing the personnel number, and Display.  The three name controls are
    read directly, without reselecting the infotype or traversing the screen.
    No fields are saved or changed.
    """
    _require_actions_enabled()
    if not 1 <= len(personnel_numbers) <= 100:
        raise ValueError("personnel_numbers must contain 1 to 100 numbers.")
    normalized = [number.strip() for number in personnel_numbers]
    if any(not re.fullmatch(r"\d{1,16}", number) for number in normalized):
        raise ValueError("Each personnel number must contain 1 to 16 digits.")

    with _sap_call():
        session = _session(session_id, require_explicit=True)
        session.StartTransaction("PA20")
        _wait_until_ready(session)
        infotype_field = session.findById(
            "wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/"
            "subSUBSCR_ITKEYS:SAPMP50A:0350/ctxtRP50G-CHOIC"
        )
        infotype_field.text = "0002"
        started_at = time.perf_counter()
        result: list[dict[str, Any]] = []
        for index, personnel_number in enumerate(normalized):
            item_started_at = time.perf_counter()
            try:
                if index:
                    session.findById("wnd[0]").sendVKey(3)
                    _wait_until_ready(session)
                personnel_field = session.findById("wnd[0]/usr/ctxtRP50G-PERNR")
                display_button = session.findById("wnd[0]/tbar[1]/btn[7]")
                personnel_field.text = personnel_number
                display_button.press()
                _wait_until_ready(session)
                name_parts = [
                    _value(session.findById("wnd[0]/usr/txtP0002-NACHN"), "Text", "").strip(),
                    _value(session.findById("wnd[0]/usr/txtP0002-VORNA"), "Text", "").strip(),
                    _value(session.findById("wnd[0]/usr/txtP0002-MIDNM"), "Text", "").strip(),
                ]
                result.append({
                    "personnel_number": personnel_number,
                    "full_name": " ".join(part for part in name_parts if part),
                    "duration_ms": round((time.perf_counter() - item_started_at) * 1000),
                })
            except Exception as error:
                result.append({
                    "personnel_number": personnel_number,
                    "error": str(error),
                    "duration_ms": round((time.perf_counter() - item_started_at) * 1000),
                })
        return {
            "transaction": "PA20",
            "infotype": "0002",
            "total_ms": round((time.perf_counter() - started_at) * 1000),
            "results": result,
        }


@mcp.tool()
def sap_set_text(control_id: str, value: str, session_id: str | None = None) -> dict[str, Any]:
    """Set text in an editable SAP field. This does not save or post data."""
    _require_actions_enabled()
    with _sap_call():
        session = _session(session_id, require_explicit=True)
        control = _control(session, control_id)
        if not _value(control, "Changeable", False):
            raise ValueError("Field is not editable.")
        if _value(control, "Type", "") not in {"GuiCTextField", "GuiTextField", "GuiComboBox", "GuiOkCodeField"}:
            raise ValueError("Control is not a text field.")
        control.text = value
        return {"control_id": control_id, "value_set": True}


@mcp.tool()
def sap_press(control_id: str, session_id: str | None = None) -> dict[str, Any]:
    """Press an explicitly identified SAP button when actions are explicitly enabled."""
    _require_actions_enabled()
    with _sap_call():
        session = _session(session_id, require_explicit=True)
        control = _control(session, control_id)
        if _value(control, "Type", "") not in {"GuiButton", "GuiMenu", "GuiTab"}:
            raise ValueError("Control is not a supported button, menu, or tab.")
        control.press()
        _wait_until_ready(session)
        return _screen_summary(session)


@mcp.tool()
def sap_send_vkey(key: int, session_id: str | None = None) -> dict[str, Any]:
    """Send an allowed SAP virtual key when actions are explicitly enabled."""
    _require_actions_enabled()
    allowed = {0, 3, 8, 12}
    if key not in allowed:
        raise ValueError(f"Only these keys are allowed: {sorted(allowed)}.")
    with _sap_call():
        session = _session(session_id, require_explicit=True)
        session.findById("wnd[0]").sendVKey(key)
        _wait_until_ready(session)
        return _screen_summary(session)


if __name__ == "__main__":
    mcp.run(transport="stdio")
