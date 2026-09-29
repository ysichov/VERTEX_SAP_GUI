"""Unit tests for MCP session discovery; no SAP system is required."""

import unittest
from unittest.mock import patch

import server


class FakeInfo:
    def __init__(self, system: str, client: str, user: str, transaction: str) -> None:
        self.SystemName = system
        self.Client = client
        self.User = user
        self.Transaction = transaction


class FakeSession:
    def __init__(self, identifier: str, client: str) -> None:
        self.Id = identifier
        self.Info = FakeInfo("ALC", client, "SYCHOV", "SESSION_MANAGER")


class FakeChildren:
    def __init__(self, items) -> None:
        self._items = list(items)
        self.Count = len(self._items)

    def __call__(self, index: int):
        return self._items[index]


class FakeConnection:
    def __init__(self, sessions) -> None:
        self.Children = FakeChildren(sessions)


class FakeApplication:
    def __init__(self, connections, active) -> None:
        self.Children = FakeChildren(connections)
        self.ActiveSession = active


class FakeControl:
    def __init__(self, name: str, *, visible: bool = True) -> None:
        self.Name = name
        self.Type = "GuiCTextField"
        self.Changeable = True
        self.Visible = visible
        self.text = ""


class FakeWindow:
    def __init__(self) -> None:
        self.sent_keys: list[int] = []
        self.Text = "Maintain HR Master Data"

    def sendVKey(self, key: int) -> None:
        self.sent_keys.append(key)


class FakeRow:
    def __init__(self) -> None:
        self.Selected = False


class FakeTable:
    def __init__(self) -> None:
        self.VerticalScrollbar = type("Scrollbar", (), {"Position": 0})()
        self.row = FakeRow()

    def getAbsoluteRow(self, index: int) -> FakeRow:
        if index != 0:
            raise AssertionError(f"Unexpected row {index}")
        return self.row


class FakePa30Session:
    def __init__(self) -> None:
        self.Id = "/app/con[0]/ses[0]"
        self.Info = FakeInfo("ALC", "200", "TEST", "PA30")
        self.choice = FakeControl("RP50G-CHOIC")
        self.command = FakeControl("okcd")
        self.title = FakeControl("GV_ITEXT")
        self.title.Text = "Challenge"
        self.title.Id = "wnd[0]/usr/tblSAPMP50ATC_MENU/txtGV_ITEXT[0,0]"
        self.table = FakeTable()
        self.window = FakeWindow()
        self.Children = FakeChildren([self.choice, self.title])

    def findById(self, control_id: str):
        return {
            "wnd[0]/tbar[0]/okcd": self.command,
            "wnd[0]/usr/tblSAPMP50ATC_MENU": self.table,
            "wnd[0]": self.window,
            "wnd[0]/sbar": type("Status", (), {"Text": "", "MessageType": ""})(),
        }[control_id]


class ServerTests(unittest.TestCase):
    def test_actions_are_disabled_by_default(self) -> None:
        with patch.object(server, "ACTIONS_ENABLED", False):
            with self.assertRaises(PermissionError):
                server._require_actions_enabled()

    def test_list_sessions_returns_every_connection_and_marks_active(self) -> None:
        client_200 = FakeSession("/app/con[0]/ses[0]", "200")
        client_100 = FakeSession("/app/con[1]/ses[0]", "100")
        application = FakeApplication(
            [FakeConnection([client_200]), FakeConnection([client_100])], client_100
        )

        with patch.object(server, "_application", return_value=application):
            sessions = server.sap_list_sessions()

        self.assertEqual([item["client"] for item in sessions], ["200", "100"])
        self.assertEqual([item["active"] for item in sessions], [False, True])

    def test_application_retries_until_sap_scripting_is_available(self) -> None:
        engine = object()

        class FakeSapGui:
            GetScriptingEngine = engine

        with (
            patch.object(
                server.win32com.client,
                "GetObject",
                side_effect=[RuntimeError("consent dialog"), FakeSapGui()],
            ) as get_object,
            patch.object(server.time, "sleep") as sleep,
        ):
            self.assertIs(server._application(), engine)

        self.assertEqual(get_object.call_count, 2)
        sleep.assert_called_once_with(server.CONNECT_RETRY_SECONDS)

    def test_open_infotype_accepts_a_visible_pa30_title(self) -> None:
        session = FakePa30Session()
        with (
            patch.object(server, "ACTIONS_ENABLED", True),
            patch.object(server, "_session", return_value=session),
            patch.object(server.time, "sleep"),
        ):
            result = server.sap_open_infotype("Challenge", "display", session.Id)

        self.assertEqual(session.choice.text, "")
        self.assertTrue(session.table.row.Selected)
        self.assertEqual(session.command.text, "=DIS")
        self.assertEqual(session.window.sent_keys, [0])
        self.assertEqual(result["infotype"], "Challenge")
        self.assertFalse(result["saved"])


if __name__ == "__main__":
    unittest.main()
