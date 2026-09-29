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
_sap_lock = threading.RLock()


@contextmanager
def _sap_call():
    """Serialize SAP access and initialise COM in the current MCP worker."""
    pythoncom.CoInitialize()
    try:
        with _sap_lock:
            yield
    finally:
        pythoncom.CoUninitialize()


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
    deadline = time.monotonic() + CONNECT_TIMEOUT_SECONDS
    last_error: Exception | None = None
    while True:
        try:
            sap_gui = win32com.client.GetObject("SAPGUI")
            return sap_gui.GetScriptingEngine
        except Exception as error:
            last_error = error
            if time.monotonic() >= deadline:
                break
            time.sleep(CONNECT_RETRY_SECONDS)
    raise RuntimeError(
        "SAP GUI Scripting не стал доступен за "
        f"{CONNECT_TIMEOUT_SECONDS:g} с. Откройте SAP GUI, войдите в систему "
        "и подтвердите окно разрешения Scripting. "
        f"Последняя COM-ошибка: {last_error}"
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
        raise ValueError("Указанная SAP-сессия не найдена. Вызовите sap_list_sessions.")
    if not sessions:
        raise RuntimeError("Рабочая SAP-сессия не найдена.")
    if require_explicit and len(sessions) > 1:
        raise ValueError("Найдено несколько SAP-сессий. Укажите session_id из sap_list_sessions.")
    return sessions[0]


def _control(session: Any, control_id: str) -> Any:
    if not control_id.startswith("wnd["):
        raise ValueError("control_id должен начинаться с 'wnd['.")
    try:
        return session.findById(control_id)
    except Exception as error:
        raise RuntimeError(f"Контрол не найден: {control_id}") from error


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
    raise RuntimeError(f"Поле {name} не найдено или недоступно для ввода.")


def _require_actions_enabled() -> None:
    if not ACTIONS_ENABLED:
        raise PermissionError(
            "Изменяющие действия отключены. Для их включения задайте "
            "VERTEX_SAP_ENABLE_ACTIONS=1 в конфигурации MCP и перезапустите сервер."
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
def sap_get_screen(session_id: str | None = None, max_controls: int = 120) -> dict[str, Any]:
    """Return basic screen metadata plus visible controls in the active SAP session."""
    if not 1 <= max_controls <= 500:
        raise ValueError("max_controls должен быть от 1 до 500.")

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
        raise ValueError("query не должен быть пустым.")
    if not 1 <= max_results <= 100:
        raise ValueError("max_results должен быть от 1 до 100.")

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
def sap_open_transaction(code: str, session_id: str | None = None) -> dict[str, Any]:
    """Open a SAP transaction when actions are explicitly enabled."""
    _require_actions_enabled()
    normalized = code.strip().upper()
    if not re.fullmatch(r"/?[A-Z0-9_]{2,30}(?:/[A-Z0-9_]{1,30})*", normalized):
        raise ValueError("Некорректный код транзакции.")

    with _sap_call():
        session = _session(session_id, require_explicit=True)
        session.StartTransaction(normalized)
        time.sleep(0.3)
        return _screen_summary(session)


@mcp.tool()
def sap_open_infotype(
    infotype: str,
    mode: str = "display",
    session_id: str | None = None,
) -> dict[str, Any]:
    """Open a PA30 infotype by its number or visible name, without saving data.

    Examples: ``0004`` or ``Challenge``.  The name is resolved by the current
    PA30 configuration, so localized and customer-specific infotype titles are
    supported.  ``mode`` may be ``display``, ``change``, or ``create``; opening
    an infotype never saves it.
    """
    _require_actions_enabled()
    selection = infotype.strip()
    if not selection or len(selection) > 80:
        raise ValueError("infotype должен содержать от 1 до 80 символов.")
    function_codes = {"display": "=DIS", "change": "=MOD", "create": "=INS"}
    normalized_mode = mode.strip().lower()
    if normalized_mode not in function_codes:
        raise ValueError("mode должен быть display, change или create.")

    with _sap_call():
        session = _session(session_id, require_explicit=True)
        if _value(_value(session, "Info", None), "Transaction", "").upper() != "PA30":
            raise ValueError("sap_open_infotype доступен только в транзакции PA30.")
        _editable_text_control_by_name(session, "RP50G-CHOIC").text = selection
        command_field = session.findById("wnd[0]/tbar[0]/okcd")
        command_field.text = function_codes[normalized_mode]
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(0.3)
        return {
            "infotype": selection,
            "mode": normalized_mode,
            "saved": False,
            **_screen_summary(session),
        }


@mcp.tool()
def sap_set_text(control_id: str, value: str, session_id: str | None = None) -> dict[str, Any]:
    """Set text in an editable SAP field. This does not save or post data."""
    _require_actions_enabled()
    with _sap_call():
        session = _session(session_id, require_explicit=True)
        control = _control(session, control_id)
        if not _value(control, "Changeable", False):
            raise ValueError("Поле недоступно для ввода.")
        if _value(control, "Type", "") not in {"GuiCTextField", "GuiTextField", "GuiComboBox", "GuiOkCodeField"}:
            raise ValueError("Контрол не является текстовым полем.")
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
            raise ValueError("Контрол не является поддерживаемой кнопкой, меню или вкладкой.")
        control.press()
        time.sleep(0.3)
        return _screen_summary(session)


@mcp.tool()
def sap_send_vkey(key: int, session_id: str | None = None) -> dict[str, Any]:
    """Send an allowed SAP virtual key when actions are explicitly enabled."""
    _require_actions_enabled()
    allowed = {0, 3, 8, 12}
    if key not in allowed:
        raise ValueError(f"Разрешены только клавиши: {sorted(allowed)}.")
    with _sap_call():
        session = _session(session_id, require_explicit=True)
        session.findById("wnd[0]").sendVKey(key)
        time.sleep(0.3)
        return _screen_summary(session)


if __name__ == "__main__":
    mcp.run(transport="stdio")
