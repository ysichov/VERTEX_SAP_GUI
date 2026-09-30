r"""Open PA30 for personnel number 91000005, infotype 0002, in change mode.

The script only opens the change screen. It does not change fields or press Save.
Run:
    py C:\soft\GitHub\VERTEX_SAP_GUI\open_pa30_it0002_edit.py
"""

import sys
import time

import win32com.client


PERNR = "91000005"
INFOTYPE = "0002"


def walk(component):
    """Yield a component and all nested SAP GUI controls."""
    yield component
    try:
        children = component.Children
        for index in range(children.Count):
            yield from walk(children(index))
    except Exception:
        return


def find_by_name(session, name):
    fallback = None
    for control in walk(session):
        try:
            if control.Name == name:
                # The same technical name can appear on a label, a hidden
                # field, and the actual input field. We only need editable
                # text or combo-box fields.
                if control.Type not in ("GuiCTextField", "GuiTextField", "GuiComboBox"):
                    continue
                if not control.Changeable:
                    continue
                # PA30 screens can contain hidden and visible copies of some
                # fields. Direct selection needs the visible copy.
                if fallback is None:
                    fallback = control
                if control.Visible:
                    return control
        except Exception:
            continue
    if fallback is not None:
        return fallback
    raise RuntimeError(f"Контрол {name} не найден")


def wait_for_by_name(session, name, timeout=10):
    deadline = time.monotonic() + timeout
    last_error = None
    while time.monotonic() < deadline:
        try:
            return find_by_name(session, name)
        except RuntimeError as error:
            last_error = error
            time.sleep(0.5)
    raise last_error or RuntimeError(f"Контрол {name} не найден")


def open_in_change_mode(session):
    """Send PA30's standard MOD function code (Change)."""
    command_field = session.findById("wnd[0]/tbar[0]/okcd")
    command_field.text = "=MOD"
    session.findById("wnd[0]").sendVKey(0)


def active_session(application):
    deadline = time.monotonic() + 30
    while time.monotonic() < deadline:
        try:
            session = application.ActiveSession
            if session:
                return session
        except Exception:
            pass
        # This fallback helps when PowerShell is the active window.
        try:
            for connection_index in range(application.Children.Count):
                connection = application.Children(connection_index)
                if connection.Children.Count:
                    return connection.Children(0)
        except Exception:
            pass
        time.sleep(1)
    raise RuntimeError("Рабочая SAP-сессия не найдена")


def main() -> None:
    try:
        sap_gui = win32com.client.GetObject("SAPGUI")
        application = sap_gui.GetScriptingEngine
        print("Если SAP спросит разрешение на доступ скрипта, нажмите OK.")
        session = active_session(application)

        session.StartTransaction("PA30")

        wait_for_by_name(session, "RP50G-PERNR").text = PERNR
        wait_for_by_name(session, "RP50G-CHOIC").text = INFOTYPE
        open_in_change_mode(session)

        print(f"Открыт табельный номер {PERNR}, инфотип {INFOTYPE}, режим изменения.")
        print("Данные не изменялись и не сохранялись.")
    except Exception as error:
        print("Не удалось открыть инфотип на изменение.")
        print(f"Техническая информация: {error}")
        sys.exit(1)


if __name__ == "__main__":
    main()
