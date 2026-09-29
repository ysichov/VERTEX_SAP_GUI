r"""Открывает транзакцию PA30 в первой доступной сессии SAP GUI.

Перед запуском войдите в SAP и откройте главное окно SAP Easy Access.
Запуск:
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

        # После первого подключения SAP GUI может показать пользователю окно
        # подтверждения. Ждём завершения этого действия, затем ищем сессию.
        deadline = time.monotonic() + 30
        active_session_error = None
        while time.monotonic() < deadline:
            # При запуске из SAP Logon коллекция Children может описывать
            # только само окно Logon. ActiveSession указывает на экран, с
            # которым сейчас работает пользователь.
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
