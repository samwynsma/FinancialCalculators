import app_help
from display_info import GetExcelInformation


class FakeButton:
    def __init__(self):
        self.options = {}

    def config(self, **options):
        self.options.update(options)


def test_get_help_toggles_mode_and_updates_button(monkeypatch):
    info_window = GetExcelInformation()
    info_window.help_button = FakeButton()
    monkeypatch.setattr(app_help, "is_help_mode", False)

    assert info_window.get_help() is True
    assert app_help.is_help_mode is True
    assert info_window.help_button.options["text"] == "Help: On"

    assert info_window.get_help() is False
    assert app_help.is_help_mode is False
    assert info_window.help_button.options["text"] == "Help"


def test_run_or_show_help_suppresses_action_and_displays_info_page_help(monkeypatch):
    info_window = GetExcelInformation()
    calls = []
    monkeypatch.setattr(app_help, "is_help_mode", True)
    monkeypatch.setattr(
        app_help, "get_help", lambda category, index: calls.append((category, index))
    )
    action_calls = []

    info_window.run_or_show_help(4, lambda: action_calls.append("export"))

    assert calls == [("InfoPage", 4)]
    assert action_calls == []


def test_run_or_show_help_runs_action_when_help_is_off(monkeypatch):
    info_window = GetExcelInformation()
    monkeypatch.setattr(app_help, "is_help_mode", False)
    action_calls = []

    info_window.run_or_show_help(4, lambda: action_calls.append("export"))

    assert action_calls == ["export"]


def test_export_csv_passes_current_page_to_document(monkeypatch):
    info_window = GetExcelInformation()
    info_window.current_page = 2
    calls = []
    monkeypatch.setattr(info_window, "select_export_pages", lambda file_type: "current")
    monkeypatch.setattr(
        info_window.document,
        "print_document_csv",
        lambda page_index=None: calls.append(page_index),
    )

    info_window.export_csv()

    assert calls == [2]


def test_export_csv_does_not_export_when_selection_is_cancelled(monkeypatch):
    info_window = GetExcelInformation()
    calls = []
    monkeypatch.setattr(info_window, "select_export_pages", lambda file_type: "cancel")
    monkeypatch.setattr(
        info_window.document,
        "print_document_csv",
        lambda page_index=None: calls.append(page_index),
    )

    info_window.export_csv()

    assert calls == []
