r"""Check the connection to an already open SAP GUI session through Scripting API.

Before running, open SAP Logon, sign in, and leave SAP GUI running.
Run from PowerShell:
    py C:\soft\GitHub\VERTEX_SAP_GUI\check_sap.py
"""

import sys

import win32com.client


def main() -> None:
    try:
        sap_gui = win32com.client.GetObject("SAPGUI")
        application = sap_gui.GetScriptingEngine
    except Exception as error:
        print("Could not connect to SAP GUI.")
        print("Open SAP GUI, sign in, and make sure Scripting is enabled.")
        print(f"Technical details: {error}")
        sys.exit(1)

    connection_count = application.Children.Count
    print(f"SAP connections: {connection_count}")

    if connection_count == 0:
        print("No open SAP connections found.")
        return

    for connection_index in range(connection_count):
        connection = application.Children(connection_index)
        session_count = connection.Children.Count
        print(f"\nConnection {connection_index}: sessions - {session_count}")

        for session_index in range(session_count):
            session = connection.Children(session_index)
            info = session.Info
            print(f"  Session {session_index}:")
            print(f"    System: {info.SystemName}")
            print(f"    Client: {info.Client}")
            print(f"    User: {info.User}")
            print(f"    Transaction: {info.Transaction or '(initial screen)'}")


if __name__ == "__main__":
    main()
