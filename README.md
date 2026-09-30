# VERTEX SAP GUI MCP

## Goal

**VERTEX SAP GUI MCP** is an MCP server for safely controlling an already open SAP GUI for Windows session from AI clients such as Codex, Claude Code, or Copilot. It turns the SAP GUI Scripting API into a small, controlled toolset: an agent can navigate to transactions, find and fill fields, press buttons, read tables, and report the result.

The first goal is a reliable, transparent prototype for a limited set of business scenarios. A universal autonomous agent for every SAP screen is a later stage: SAP GUI contains custom controls, ALV grids, trees, modal windows, and error messages that require dedicated adapters and controls.

## Architecture

```text
Codex / Claude Code / Copilot
              |
              | MCP (local process)
              v
+-----------------------------------+
|          VERTEX SAP GUI MCP       |
|                                   |
|  session | screen | controls      |
|  input   | actions | table reader |
|  policy  | audit   | confirmations|
+----------------+------------------+
                 |
                 | Python API
                 v
       SAP GUI Scripting COM API
                 |
                 v
        SAP GUI for Windows
                 |
                 v
           SAP ECC / SAP S/4HANA
```

The server does not keep business logic in the prompt. It exposes predictable operations against the active SAP session, while the AI client chooses the call sequence and interprets the results. Each operation returns structured data: screen state, discovered controls, system message, error, and, when needed, a screenshot.

## Technology Stack

| Layer | Technology | Purpose |
| --- | --- | --- |
| AI client | Codex, Claude Code, Copilot | Plans actions and calls MCP tools |
| Integration | Model Context Protocol (MCP) | Standardized interface for local tools |
| Server | Python 3.11+ | Orchestration, policies, audit, data models |
| Windows COM | `pywin32` / `pythoncom` | Access to SAP GUI COM objects |
| Automation | SAP GUI Scripting API | Work with SAP GUI sessions and elements |
| Observability | JSON logs, snapshots, correlation ID | Audit and scenario debugging |

SAP GUI Scripting must be enabled on the SAP server side and in SAP GUI for Windows. The server connects only to an already authorized user session; it must not accept or store logins or passwords.

## MCP Tool Set

A minimal, safe contract can start with the following tools.

| Tool | Purpose |
| --- | --- |
| `sap_connect` | Finds available SAP GUI connections/sessions and selects a session by explicit ID. |
| `sap_get_screen` | Returns the transaction, window title, status bar, modal windows, and a compact element tree. |
| `sap_find_controls` | Finds controls by technical ID, type, label, or pattern. |
| `sap_read_fields` | Reads values from a list of known controls directly, without walking the screen tree. |
| `sap_read_grid_rows` | Reads SAP Grid rows directly, without opening records. |
| `sap_select_grid_row` | Selects a SAP Grid row without saving data; the column is optional. |
| `sap_read_grid_infotype_fields` | Switches visible SAP Grid rows and reads known infotype fields in one call. |
| `sap_set_text` | Writes a value to a field after checking that the field is editable. |
| `sap_open_infotype` | Opens a PA30 infotype by number or visible title, without saving data. |
| `sap_inspect_infotype` | Opens a PA30 infotype and returns visible fields in one call. |
| `sap_read_personnel_names` | Reads full names from IT0002 for an explicit list of personnel numbers in one fast PA20 batch. |
| `sap_press` | Presses a button, toolbar item, or menu item by an allowed ID. |
| `sap_select` | Selects a value in a combo box, radio button, tab, or list. |
| `sap_send_vkey` | Sends a limited set of SAP GUI virtual keys, such as Enter or Back. |
| `sap_transaction` | Navigates to an allowlisted transaction and returns the navigation result. |
| `sap_read_table` | Reads an ALV/table control with row-count limits. |
| `sap_read_tree` | Reads and expands allowed SAP tree-control nodes. |
| `sap_get_message` | Returns the latest system message: type, number, and text. |
| `sap_capture` | Creates a snapshot of the current window for diagnostics after explicit approval. |
| `sap_confirm` | Confirms a higher-risk action with a one-time user token. |

In early stages, it is useful to split tools into read-only and write/action groups. This lets AI start in observation mode and test screen recognition without changing data.

## Safety And Control

- Use only an already open interactive SAP session; do not pass credentials through MCP.
- Allow transactions, system commands, tables, and actions through an allowlist tied to the user's role.
- Run in **read-only** mode by default; enable write actions through separate configuration.
- Require explicit confirmation before saving, posting, deleting, sending messages, and mass changes.
- Limit ALV row reads, exports, and field contents so extra personal or commercial data is not sent to the AI client.
- Keep immutable audit data: user, session, tool, masked parameters, result, time, and correlation ID.
- Set timeouts, retry limits, and expected-screen checks after every changing action.
- Do not execute arbitrary SAP commands, VBScript, or Python code received from the model.

## Implementation Stages

1. **Connection and observation.** Implement SAP GUI session discovery, `sap_connect`, `sap_get_screen`, control search, and status reads. Verify on one test system.
2. **Basic actions.** Add input, pressing, value selection, and transaction navigation from an allowlist; verify the screen and status message after each action.
3. **Scenarios and adapters.** Implement ALV reading, tree controls, and pop-up handling. Describe two or three concrete business flows as tests.
4. **Policies and audit.** Add roles, read-only/write modes, confirmations, logging, data masking, and limits.
5. **Reliability.** Add common error classification, recovery from expected screens, diagnostic snapshots, and regression tests.
6. **Agent workflow.** Connect an MCP client, give it tool descriptions and scenario templates. Add a planner and vision fallback only after deterministic scenarios become measurably stable.

## Example Scenario

Task: check open orders for a customer in an allowed transaction and return a short summary without changing data.

```text
1. sap_connect(session_id="...")
2. sap_transaction(code="VA05")
3. sap_get_screen()
4. sap_find_controls(query="Sold-to party")
5. sap_set_text(control_id="...", value="10001234")
6. sap_press(control_id="...execute...")
7. sap_get_message()
8. sap_read_table(table_id="...", max_rows=100)
9. The AI client aggregates only the required fields and shows the summary to the user.
```

In a write scenario, before a step that saves or posts a document, the server should return a confirmation requirement. The AI client shows the user a clear description of the change, receives approval, and passes a one-time token to `sap_confirm` or to a protected action tool.

## Comparison With SAP Voyager

| Aspect | VERTEX SAP GUI MCP | SAP Voyager |
| --- | --- | --- |
| Core idea | Thin, controlled MCP layer over SAP GUI Scripting | Broader agent system with planning, self-correction, and vision fallback |
| Starting complexity | Low: Python, COM, and explicit tools | Higher: needs an agent loop and additional recovery mechanisms |
| Predictability | High for limited allowlisted scenarios | Depends on the planner, models, and visual recognition |
| Adoption approach | Deterministic workflows and audit first | Focus on more autonomous task execution |
| When to choose | Internal tool, PoC, controlled operations, Codex/Claude Code integration | When more universal navigation, auto-correction, and complex-screen handling are needed |

VERTEX SAP GUI MCP does not try to replace SAP Voyager immediately. It creates a simpler and more controlled foundation: reliable tool calls, access policy, and observability. Planner, auto-correction, and vision fallback can be added on top when basic SAP scenarios become measurably stable.

## First Run And MCP Setup

For a step-by-step local MCP configuration guide, see [MCP_SETUP.md](MCP_SETUP.md).

On first access, SAP GUI may show the user a Scripting consent dialog.
The server waits up to 25 seconds for the Scripting Engine and retries the
connection, so confirming the dialog does not require another tool call.
The timeout is controlled by the `VERTEX_SAP_CONNECT_TIMEOUT` environment
variable, in seconds. Inspection tools (`sap_list_sessions`, `sap_status`,
`sap_get_screen`, `sap_find_controls`) are always available. Navigation,
input, and pressing are disabled by default. Enable them only when needed by
adding `VERTEX_SAP_ENABLE_ACTIONS: "1"` to the MCP configuration `env`.

After actions are enabled, PA30 infotypes can be opened by number or by visible
title: `sap_open_infotype(infotype="0004")` or
`sap_open_infotype(infotype="Challenge")`. For a title, the tool finds and
selects the matching row in the current PA30 menu, without replacing it with a
static code; this supports localized and customer-specific titles. By default,
the infotype opens in display mode; `mode="change"` and `mode="create"` only
open the screen and do not save the record.

For fast scenarios, use `sap_inspect_infotype`: it combines opening an infotype
and reading fields in one call. After actions, the server waits for SAP session
readiness through the `Busy` property instead of using a fixed sleep.
Field schemas are cached for five minutes per system, client, transaction, and
screen; field values and messages are never cached. If screen settings change,
call `sap_clear_layout_cache`.

MCP also keeps the COM connection to SAP GUI in a live worker thread and reuses
the Scripting Engine between calls. This reduces overhead and repeated script
attachment prompts; if SAP GUI is closed, the stale connection is reset
automatically.

For small explicit lists, use `sap_read_personnel_names`. The tool opens PA20
and selects IT0002 once, then changes only the personnel number between
employees and returns total and per-row execution time.

Use the absolute path to the same `python.exe` that successfully runs
`check_sap.py`. See `mcp.json.example` for a sample configuration. After
changing the MCP configuration, restart the MCP server, usually by restarting
Codex, and call `sap_list_sessions`.
