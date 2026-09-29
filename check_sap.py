r"""Проверка подключения к уже открытой сессии SAP GUI через Scripting API.

Перед запуском откройте SAP Logon, войдите в систему и оставьте SAP GUI запущенным.
Запуск из PowerShell:
    py C:\soft\GitHub\VERTEX_SAP_GUI\check_sap.py
"""

import sys

import win32com.client


def main() -> None:
    try:
        sap_gui = win32com.client.GetObject("SAPGUI")
        application = sap_gui.GetScriptingEngine
    except Exception as error:
        print("Не удалось подключиться к SAP GUI.")
        print("Откройте SAP GUI, войдите в систему и убедитесь, что Scripting включён.")
        print(f"Техническая информация: {error}")
        sys.exit(1)

    connection_count = application.Children.Count
    print(f"Подключений SAP: {connection_count}")

    if connection_count == 0:
        print("Открытых SAP-подключений не найдено.")
        return

    for connection_index in range(connection_count):
        connection = application.Children(connection_index)
        session_count = connection.Children.Count
        print(f"\nПодключение {connection_index}: сессий — {session_count}")

        for session_index in range(session_count):
            session = connection.Children(session_index)
            info = session.Info
            print(f"  Сессия {session_index}:")
            print(f"    Система: {info.SystemName}")
            print(f"    Клиент: {info.Client}")
            print(f"    Пользователь: {info.User}")
            print(f"    Транзакция: {info.Transaction or '(начальный экран)'}")


if __name__ == "__main__":
    main()
