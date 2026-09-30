# MCP Setup

This guide explains how to run the VERTEX SAP GUI MCP server from a local AI client.

## Prerequisites

- Windows with SAP GUI for Windows installed.
- SAP GUI Scripting enabled on the SAP server and in the SAP GUI client.
- An already open, signed-in SAP GUI session. The MCP server does not log in to SAP.
- Python 3.11 or newer.
- Project dependencies installed in the same Python environment that the MCP client will start.

Install dependencies:

```powershell
py -m pip install -r C:\soft\GitHub\VERTEX_SAP_GUI\requirements.txt
```

If you use a virtual environment, use that environment's full `python.exe` path in the MCP configuration.

## Verify SAP GUI Access

Open SAP Logon, sign in to the target SAP system, and leave SAP GUI running.

Then run:

```powershell
py C:\soft\GitHub\VERTEX_SAP_GUI\check_sap.py
```

Expected result:

- at least one SAP connection is listed;
- at least one session is listed;
- system, client, user, and transaction are printed.

If SAP shows a scripting consent dialog, confirm it manually. The MCP server never clicks security dialogs for you.

## MCP Configuration

Use `mcp.json.example` as the base configuration:

```json
{
  "mcpServers": {
    "vertex-sap-gui": {
      "command": "D:\\Soft\\python.exe",
      "args": [
        "C:\\soft\\GitHub\\VERTEX_SAP_GUI\\server.py"
      ],
      "env": {
        "VERTEX_SAP_CONNECT_TIMEOUT": "25"
      }
    }
  }
}
```

Change `command` to the absolute path of the Python executable that can run `check_sap.py`.

Change the path in `args` if this repository is located somewhere else.

## Enable Actions

By default, inspection tools are available, but changing actions are disabled.

Inspection tools include:

- `sap_list_sessions`
- `sap_status`
- `sap_get_screen`
- `sap_find_controls`
- `sap_read_fields`
- `sap_read_grid_rows`

To enable navigation, field input, button presses, and infotype opening, add:

```json
"VERTEX_SAP_ENABLE_ACTIONS": "1"
```

Example:

```json
{
  "mcpServers": {
    "vertex-sap-gui": {
      "command": "D:\\Soft\\python.exe",
      "args": [
        "C:\\soft\\GitHub\\VERTEX_SAP_GUI\\server.py"
      ],
      "env": {
        "VERTEX_SAP_CONNECT_TIMEOUT": "25",
        "VERTEX_SAP_ENABLE_ACTIONS": "1"
      }
    }
  }
}
```

Restart the MCP client after changing this configuration.

## Optional Environment Variables

| Variable | Default | Purpose |
| --- | --- | --- |
| `VERTEX_SAP_CONNECT_TIMEOUT` | `25` | Seconds to wait for the SAP GUI Scripting Engine. |
| `VERTEX_SAP_ENABLE_ACTIONS` | `0` | Enables changing actions when set to `1`. |
| `VERTEX_SAP_READY_TIMEOUT` | `20` | Seconds to wait for SAP GUI to finish a round trip. |
| `VERTEX_SAP_LAYOUT_CACHE_SECONDS` | `300` | Seconds to cache static screen field schemas. Set to `0` to disable the cache. |

## First Check From The MCP Client

After the MCP client starts the server, call:

```text
sap_list_sessions
```

You should see sessions like:

```json
{
  "session_id": "/app/con[0]/ses[0]",
  "system": "ALC",
  "client": "200",
  "user": "USER",
  "transaction": "PA20",
  "active": true
}
```

Use `session_id` explicitly when more than one SAP session is open.

## Common Problems

### The server cannot connect to SAP GUI

Check that:

- SAP GUI is open and signed in;
- SAP GUI Scripting is enabled;
- the same Windows user runs SAP GUI and the MCP server;
- the Python environment has `pywin32` installed;
- the scripting consent dialog was confirmed.

### Actions are rejected

Set `VERTEX_SAP_ENABLE_ACTIONS` to `1` and restart the MCP client.

### Multiple sessions are found

Pass the exact `session_id` returned by `sap_list_sessions`.

### SAP prompts for scripting access repeatedly

The server keeps one COM attachment per worker thread to reduce repeated prompts. If prompts still appear, check local SAP GUI security settings and SAP GUI Scripting notification settings.

## Safe Operating Notes

- Keep write actions disabled unless you explicitly need them.
- Prefer read-only tools for discovery and diagnostics.
- Read the current screen or target field before changing data.
- Treat `sap_press`, `sap_set_text`, `sap_send_vkey`, and `sap_open_infotype(mode="change"|"create")` as action tools.
- Saving data is only possible through explicit button/key actions and should be confirmed at the workflow level before use.
