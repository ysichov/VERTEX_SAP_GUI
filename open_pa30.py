r"""Open transaction PA30 in the first available SAP GUI session.

Before running, sign in to SAP and open the SAP Easy Access main window.
Run:
    py C:\soft\GitHub\VERTEX_SAP_GUI\open_pa30.py
"""

import sys
import time

import win32com.client


def main() -> None:
    try:
        sap_gui = win32com.client.GetObject("SAPGUI")
        application = sap_gui.GetScriptingEngine

        print("Если SAP спросит разрешение на доступ скрипта, нажмите OK.")

        # SAP GUI may show a consent dialog after the first attachment.
        # Wait for that interaction to finish, then look for a session.
        deadline = time.monotonic() + 30
        active_session_error = None
        while time.monotonic() < deadline:
            # When launched from SAP Logon, Children may describe only the
            # Logon window. ActiveSession points to the screen the user is
            # currently working with.
            try:
                session = application.ActiveSession
                if session:
                    session.StartTransaction("PA30")
                    print("Транзакция PA30 открыта.")
                    return
            except Exception as error:
                active_session_error = error

            for connection_index in range(application.Children.Count):
                connection = application.Children(connection_index)
                if connection.Children.Count:
                    session = connection.Children(0)
                    session.StartTransaction("PA30")
                    print("Транзакция PA30 открыта.")
                    return
            time.sleep(1)

        print("Рабочая SAP-сессия не найдена.")
        if active_session_error:
            print(f"Не удалось получить активную сессию: {active_session_error}")
        print("Откройте систему в SAP Logon и войдите в SAP, затем повторите запуск.")
        sys.exit(1)
    except Exception as error:
        print("Не удалось открыть PA30.")
        print(f"Техническая информация: {error}")
        sys.exit(1)


if __name__ == "__main__":
    main()
