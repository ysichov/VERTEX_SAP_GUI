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

        print("If SAP asks for scripting access, click OK.")

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
                    print("Transaction PA30 opened.")
                    return
            except Exception as error:
                active_session_error = error

            for connection_index in range(application.Children.Count):
                connection = application.Children(connection_index)
                if connection.Children.Count:
                    session = connection.Children(0)
                    session.StartTransaction("PA30")
                    print("Transaction PA30 opened.")
                    return
            time.sleep(1)

        print("No working SAP session found.")
        if active_session_error:
            print(f"Could not get the active session: {active_session_error}")
        print("Open the system in SAP Logon, sign in, then run the script again.")
        sys.exit(1)
    except Exception as error:
        print("Could not open PA30.")
        print(f"Technical details: {error}")
        sys.exit(1)


if __name__ == "__main__":
    main()
