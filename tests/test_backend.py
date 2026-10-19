"""Portable tests for native sheet replacement and transactional publication."""
from __future__ import annotations

from contextlib import ExitStack
from dataclasses import replace
import json
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch

from pops_update import backend
from pops_update.config import CountryUpdate, SheetUpdate, UpdateConfig
from pops_update.excel_session import ExcelSession


ALLOW_OPTIONS = (
    "AllowFormattingCells", "AllowFormattingColumns", "AllowFormattingRows",
    "AllowInsertingColumns", "AllowInsertingRows", "AllowInsertingHyperlinks",
    "AllowDeletingColumns", "AllowDeletingRows", "AllowSorting", "AllowFiltering",
    "AllowUsingPivotTables",
)


class FakeSheet:
    def __init__(self, workbook, name, payload, protection=None):
        self.workbook, self.Name, self.payload = workbook, name, payload
        self.ProtectDrawingObjects = self.ProtectContents = self.ProtectScenarios = False
        self.ProtectionMode = False
        self.Protection = SimpleNamespace(**{option: False for option in ALLOW_OPTIONS})
        self.EnableSelection, self._password = 0, ""
        if protection:
            self._set_protection(protection)
            self.EnableSelection = protection.get("EnableSelection", 0)

    @property
    def protected(self):
        return self.ProtectDrawingObjects or self.ProtectContents or self.ProtectScenarios

    def _set_protection(self, options):
        self.ProtectDrawingObjects = options.get("DrawingObjects", True)
        self.ProtectContents = options.get("Contents", True)
        self.ProtectScenarios = options.get("Scenarios", True)
        self.ProtectionMode = options.get("UserInterfaceOnly", False)
        self.Protection = SimpleNamespace(**{option: options.get(option, False) for option in ALLOW_OPTIONS})
        self._password = options.get("Password", "")

    def protection_state(self):
        return {
            "Password": self._password, "DrawingObjects": self.ProtectDrawingObjects,
            "Contents": self.ProtectContents, "Scenarios": self.ProtectScenarios,
            "UserInterfaceOnly": self.ProtectionMode, "EnableSelection": self.EnableSelection,
            **vars(self.Protection),
        }

    def Unprotect(self, *, Password=""):
        self.workbook.session.events.append(("unprotect_sheet", self.Name, Password))
        if self.protected and Password != self._password:
            raise RuntimeError("Incorrect sheet password")
        self.ProtectDrawingObjects = self.ProtectContents = self.ProtectScenarios = False
        self.ProtectionMode = False

    def Protect(self, **options):
        self.workbook.session.events.append(("protect_sheet", self.Name, options))
        self._set_protection(options)

    def Delete(self):
        self.workbook.sheets.remove(self)


class FakeGroup:
    def __init__(self, sheets):
        self.sheets = sheets


class FakeSheets:
    def __init__(self, workbook):
        self.workbook = workbook

    @property
    def Count(self):
        return len(self.workbook.sheets)

    def __call__(self, key):
        if isinstance(key, tuple):
            return FakeGroup([self(name) for name in key])
        if isinstance(key, int):
            return self.workbook.sheets[key - 1]
        return next(sheet for sheet in self.workbook.sheets if sheet.Name == key)


class FakeWorkbook:
    def __init__(self, session, path, names, payload, *, read_only=False, protected=False):
        self.session, self.path, self.read_only = session, Path(path), read_only
        self.ProtectStructure = protected
        self.ProtectWindows = protected
        self.sheets = [FakeSheet(self, name, payload) for name in names]
        self.Sheets = self.Worksheets = FakeSheets(self)
        self.links = []

    @property
    def Name(self):
        return self.path.name

    @property
    def FullName(self):
        return str(self.path)

    def Unprotect(self, **kwargs):
        self.ProtectStructure = False
        self.ProtectWindows = False

    def Protect(self, **kwargs):
        self.ProtectStructure = kwargs["Structure"]
        self.ProtectWindows = kwargs["Windows"]

    def SaveAs(self, *, Filename, FileFormat):
        self.path = Path(Filename)
        self.path.write_bytes(b"saved scratch workbook")

    def Save(self):
        if self.read_only:
            raise AssertionError("Read-only sources must never be saved")
        self.session.events.append(("save", self.path))
        self.path.write_text(json.dumps({
            "sheets": [(sheet.Name, sheet.payload) for sheet in self.sheets],
            "protected": self.ProtectStructure, "windows_protected": self.ProtectWindows, "links": self.links,
            "sheet_protection": {sheet.Name: sheet.protection_state() for sheet in self.sheets},
        }), encoding="utf-8")

    def LinkSources(self, link_type):
        return tuple(self.links)

    def ChangeLink(self, *, Name, NewName, Type):
        if any(sheet.protected for sheet in self.sheets):
            raise RuntimeError("Cannot change links on a protected worksheet")
        self.session.events.append(("change_link", Name, NewName))
        if self.session.link_error:
            raise self.session.link_error
        self.links.remove(Name)
        if NewName != self.FullName:
            self.links.append(NewName)


class FakeExcelSession:
    run = ExcelSession.run
    def __init__(self, source_path, destination_names, source_names):
        self.source_path, self.destination_names, self.source_names = source_path, destination_names, source_names
        self.events, self.open_workbooks = [], []
        self.calculation_error = None
        self.link_error = None
        self.destination_protection, self.source_protection = {}, {}
        self.app = SimpleNamespace(CalculateFullRebuild=self.calculate,
                                   Workbooks=SimpleNamespace(Add=self.new_workbook))

    def __enter__(self):
        return self

    def __exit__(self, *_):
        for workbook in list(self.open_workbooks):
            self.close_workbook(workbook)
        self.events.append(("exit",))
        return False

    def call(self, operation):
        return operation()

    def open_workbook(self, path, *, read_only):
        self.events.append(("open", path, read_only))
        if not path.is_file():
            raise FileNotFoundError(path)
        is_source = path == self.source_path
        workbook = FakeWorkbook(self, path, self.source_names if is_source else self.destination_names,
                                "new source content and formatting" if is_source else "original content",
                                read_only=read_only, protected=True)
        protection = self.source_protection if is_source else self.destination_protection
        for sheet in workbook.sheets:
            if sheet.Name in protection:
                sheet._set_protection(protection[sheet.Name])
                sheet.EnableSelection = protection[sheet.Name].get("EnableSelection", 0)
        if is_source:
            self.source = workbook
        else:
            self.destination = workbook
        return self.track_workbook(workbook)

    def new_workbook(self, _):
        return FakeWorkbook(self, "Book1", ("Sheet1",), "scratch")

    def track_workbook(self, workbook):
        self.open_workbooks.append(workbook)
        return workbook

    def close_workbook(self, workbook):
        self.events.append(("close", workbook.path, workbook.read_only))
        self.open_workbooks.remove(workbook)

    def add_sheet_after(self, workbook, anchor):
        sheet = FakeSheet(workbook, "Temporary anchor", "placeholder")
        workbook.sheets.insert(workbook.sheets.index(anchor) + 1, sheet)
        return sheet

    def move_after(self, selection, anchor):
        sheets = selection.sheets if isinstance(selection, FakeGroup) else [selection]
        original, target = sheets[0].workbook, anchor.workbook
        if original is not target:
            original.links.append(target.FullName)
        for sheet in sheets:
            original.sheets.remove(sheet)
        position = target.sheets.index(anchor) + 1
        for offset, sheet in enumerate(sheets):
            sheet.workbook = target
            target.sheets.insert(position + offset, sheet)

    def copy_after(self, selection, anchor):
        sheets = selection.sheets if isinstance(selection, FakeGroup) else [selection]
        self.events.append(("copy", tuple(sheet.Name for sheet in sheets)))
        target = anchor.workbook
        target.links.append(sheets[0].workbook.FullName)
        position = target.sheets.index(anchor) + 1
        for offset, sheet in enumerate(sheets):
            copied = FakeSheet(target, sheet.Name, sheet.payload,
                               sheet.protection_state() if sheet.protected else None)
            target.sheets.insert(position + offset, copied)

    def restore_calculation(self, workbook):
        self.events.append(("restore_calculation", workbook.path))

    def calculate(self):
        self.events.append(("calculate",))
        if self.calculation_error:
            raise self.calculation_error


class RefreshTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        final, source = self.root / "final.xlsx", self.root / "portugal.xlsx"
        # Deliberately not ZIP packages: no workbook parsing belongs in Python.
        final.write_bytes(b"original final bytes")
        source.write_bytes(b"original source bytes")
        sheets = (SheetUpdate("NPL SALES", "Portugal_NPL SALES"), SheetUpdate("ID Card", "Portugal_ID Card"))
        update = CountryUpdate("Portugal", "Portugal", source, sheets)
        self.config = UpdateConfig(final, self.root / "output" / "updated.xlsx", True, True, "a", (update,),
                                   (("ID Card", "ID Card"), ("NPL SALES", "NPL SALES"), ("FTE", "FTE")))
        self.plan = backend.build_plan(self.config)
        self.original_inputs = {path: path.read_bytes() for path in (final, source)}
        self.names = ("Summary", "Portugal_ID Card", "Portugal_FTE", "Portugal_NPL SALES")
        self.session = FakeExcelSession(source, self.names, ("ID Card", "NPL SALES", "FTE"))
        self.patches = ExitStack()
        self.addCleanup(self.patches.close)
        self.excel = self.patches.enter_context(patch.object(backend, "ExcelSession", return_value=self.session))
        self.patches.enter_context(patch("zipfile.ZipFile", side_effect=AssertionError("No XML inspection is allowed")))

    def assert_inputs_unchanged(self):
        for path, content in self.original_inputs.items():
            self.assertEqual(path.read_bytes(), content)

    def assert_no_candidate(self):
        self.assertEqual(list(self.config.output_workbook.parent.glob(".*.candidate-*.xlsx")), [])

    def existing_output(self):
        self.config.output_workbook.parent.mkdir()
        self.config.output_workbook.write_bytes(b"previous successful output")
        return self.config.output_workbook.read_bytes()

    def configure_sheet_protection(self):
        summary = {
            "Password": "a", "DrawingObjects": False, "Contents": True,
            "Scenarios": False, "UserInterfaceOnly": True, "EnableSelection": 1,
            **{option: index % 2 == 0 for index, option in enumerate(ALLOW_OPTIONS)},
        }
        card = {
            "Password": "", "DrawingObjects": True, "Contents": False,
            "Scenarios": True, "EnableSelection": -4142,
            **{option: index % 2 == 1 for index, option in enumerate(ALLOW_OPTIONS)},
        }
        sales = {"Password": "a", "AllowFiltering": True, "EnableSelection": 0}
        self.session.destination_protection = {"Summary": summary}
        self.session.source_protection = {"ID Card": card, "NPL SALES": sales}
        return {name: FakeSheet(None, name, "", settings).protection_state()
                for name, settings in (("Summary", summary), ("Portugal_ID Card", card),
                                       ("Portugal_NPL SALES", sales))}

    def assert_destination_protection(self, expected):
        for name, state in expected.items():
            self.assertEqual(self.session.destination.Worksheets(name).protection_state(), state, name)

    def test_plan_uses_selection_without_inspection_even_with_legacy_check_enabled(self):
        with patch.object(Path, "open", side_effect=AssertionError("Planning must not open workbook files")):
            self.assertIs(backend.build_plan(self.config), self.config.updates)

    def test_copies_selected_group_keeps_order_and_publishes_after_excel_closes(self):
        original_replace = backend.os.replace

        def publish(candidate, output):
            self.assertEqual(self.session.open_workbooks, [])
            self.assertEqual(self.session.events[-1], ("exit",))
            return original_replace(candidate, output)

        messages = []
        with patch.object(backend.os, "replace", side_effect=publish):
            result = backend.refresh(self.config, self.plan, progress=messages.append)
        saved = json.loads(self.config.output_workbook.read_text())
        self.assertEqual([name for name, _ in saved["sheets"]], list(self.names))
        self.assertEqual(dict(saved["sheets"]), {
            "Summary": "original content", "Portugal_FTE": "original content",
            "Portugal_ID Card": "new source content and formatting",
            "Portugal_NPL SALES": "new source content and formatting",
        })
        self.assertTrue(saved["protected"])
        self.assertTrue(saved["windows_protected"])
        self.assertEqual(saved["links"], [str(self.config.updates[0].file)])
        self.assertEqual([event for event in self.session.events if event[0] == "copy"],
                         [("copy", ("Portugal_NPL SALES", "Portugal_ID Card"))])
        self.assertEqual([sheet.Name for sheet in self.session.source.sheets],
                         ["Portugal_ID Card", "Portugal_NPL SALES", "FTE"])
        operations = [event[0] for event in self.session.events]
        self.assertLess(operations.index("calculate"), operations.index("save"))
        self.assertEqual(json.loads(Path(result["report_file"]).read_text()), result)
        self.assertTrue(any("Recalculating" in message for message in messages))
        self.assertFalse(any(event[0] in ("unprotect_sheet", "protect_sheet")
                             for event in self.session.events))
        self.assert_inputs_unchanged()
        self.assert_no_candidate()

    def test_reconnects_with_protected_dependents_and_imported_sheets_then_restores_each(self):
        expected = self.configure_sheet_protection()
        backend.refresh(self.config, self.plan)
        saved = json.loads(self.config.output_workbook.read_text())
        for name, state in expected.items():
            self.assertEqual(saved["sheet_protection"][name], state, name)
        self.assert_destination_protection(expected)
        # The passwordless imported sheet must remain passwordless despite a
        # configured password, and each permission/selection choice survives.
        self.assertIn(("unprotect_sheet", "Portugal_ID Card", ""), self.session.events)
        self.assertNotIn(("unprotect_sheet", "Portugal_ID Card", "a"), self.session.events)
        protected_names = {event[1] for event in self.session.events if event[0] == "protect_sheet"}
        self.assertEqual(protected_names, set(expected))
        self.assertFalse(self.session.destination.Worksheets("Portugal_FTE").protected)
        self.assertEqual(self.session.source.Worksheets("Portugal_ID Card").protection_state(),
                         expected["Portugal_ID Card"])
        self.assert_inputs_unchanged()
        self.assert_no_candidate()

    def test_wrong_sheet_password_names_sheet_and_preserves_inputs_and_existing_output(self):
        expected = self.configure_sheet_protection()
        self.session.source_protection["ID Card"]["Password"] = "different-password"
        expected["Portugal_ID Card"]["Password"] = "different-password"
        existing = self.existing_output()
        report = self.config.output_workbook.with_name("updated.update_report.json")
        report.write_bytes(b"previous report")
        with self.assertRaisesRegex(RuntimeError, "Portugal_ID Card"):
            backend.refresh(self.config, self.plan)
        self.assert_destination_protection(expected)
        self.assertEqual(self.config.output_workbook.read_bytes(), existing)
        self.assertEqual(report.read_bytes(), b"previous report")
        self.assertFalse(any(event[0] in ("change_link", "save") for event in self.session.events))
        self.assert_inputs_unchanged()
        self.assert_no_candidate()
        self.assertEqual(self.session.open_workbooks, [])

    def test_reconnect_failure_restores_protection_and_preserves_existing_output(self):
        expected = self.configure_sheet_protection()
        existing = self.existing_output()
        self.session.link_error = RuntimeError("Excel link update failed")
        with self.assertRaisesRegex(RuntimeError, "Excel link update failed"):
            backend.refresh(self.config, self.plan)
        self.assert_destination_protection(expected)
        self.assertEqual(self.config.output_workbook.read_bytes(), existing)
        self.assert_inputs_unchanged()
        self.assert_no_candidate()
        self.assertEqual(self.session.open_workbooks, [])

    def test_failed_reprotection_prevents_publication(self):
        self.configure_sheet_protection()
        existing = self.existing_output()
        original_protect = FakeSheet.Protect

        def fail_summary_restore(sheet, **options):
            if sheet.Name == "Summary":
                raise RuntimeError("Excel protection restore failed")
            return original_protect(sheet, **options)

        with patch.object(FakeSheet, "Protect", fail_summary_restore):
            with self.assertRaisesRegex(RuntimeError, "Summary"):
                backend.refresh(self.config, self.plan)
        self.assertEqual(self.config.output_workbook.read_bytes(), existing)
        self.assertFalse(any(event[0] == "save" for event in self.session.events))
        self.assert_inputs_unchanged()
        self.assert_no_candidate()

    def test_no_replacement_links_leaves_protected_sheets_untouched(self):
        destination = FakeWorkbook(self.session, self.root / "candidate.xlsx", ("Summary",), "original")
        previous = FakeWorkbook(self.session, self.root / "previous.xlsx", ("Sheet1",), "scratch")
        summary = destination.Worksheets("Summary")
        summary._set_protection({"Password": "different-password"})
        before = summary.protection_state()
        destination.links = [str(self.config.updates[0].file)]
        backend._reconnect(self.session, destination, previous, "a")
        self.assertEqual(summary.protection_state(), before)
        self.assertEqual(self.session.events, [])

    def test_can_replace_every_destination_tab(self):
        self.session.destination_names = ("Portugal_ID Card", "Portugal_NPL SALES")
        backend.refresh(self.config, self.plan)
        saved = json.loads(self.config.output_workbook.read_text())
        self.assertEqual([name for name, _ in saved["sheets"]], list(self.session.destination_names))
        self.assertTrue(all(payload == "new source content and formatting" for _, payload in saved["sheets"]))
        self.assert_inputs_unchanged()

    def test_excel_failure_preserves_inputs_existing_output_and_report(self):
        existing = self.existing_output()
        report = self.config.output_workbook.with_name("updated.update_report.json")
        report.write_bytes(b"previous report")
        self.session.calculation_error = RuntimeError("Excel calculation failed")
        with self.assertRaisesRegex(RuntimeError, "Excel calculation failed"):
            backend.refresh(self.config, self.plan)
        self.assertEqual(self.config.output_workbook.read_bytes(), existing)
        self.assertEqual(report.read_bytes(), b"previous report")
        self.assert_inputs_unchanged()
        self.assert_no_candidate()
        self.assertEqual(self.session.open_workbooks, [])

    def test_cleanup_lock_does_not_replace_primary_error(self):
        existing = self.existing_output()
        primary = RuntimeError("Excel calculation failed")
        self.session.calculation_error = primary
        original_unlink = Path.unlink

        def locked_candidate(path, *args, **kwargs):
            if ".candidate-" in path.name:
                raise PermissionError("candidate still locked")
            return original_unlink(path, *args, **kwargs)

        with patch.object(Path, "unlink", locked_candidate):
            with self.assertRaises(RuntimeError) as caught:
                backend.refresh(self.config, self.plan)
        self.assertIs(caught.exception.__cause__, primary)
        if hasattr(primary, "add_note"):
            self.assertIn("candidate still locked", caught.exception.__notes__[0])
        self.assertEqual(self.config.output_workbook.read_bytes(), existing)
        self.assert_inputs_unchanged()
        self.assertEqual(self.session.open_workbooks, [])

    def test_overwrite_disabled_preserves_existing_output_without_starting_excel(self):
        existing = self.existing_output()
        config = replace(self.config, overwrite_output=False)
        with self.assertRaises(FileExistsError):
            backend.refresh(config, backend.build_plan(config))
        self.excel.assert_not_called()
        self.assertEqual(self.config.output_workbook.read_bytes(), existing)
        self.assert_inputs_unchanged()
        self.assert_no_candidate()


if __name__ == "__main__":
    unittest.main()
