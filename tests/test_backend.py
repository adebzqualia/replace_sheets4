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


class FakeSheet:
    def __init__(self, workbook, name, payload):
        self.workbook, self.Name, self.payload = workbook, name, payload

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
        }), encoding="utf-8")

    def LinkSources(self, link_type):
        return tuple(self.links)

    def ChangeLink(self, *, Name, NewName, Type):
        self.session.events.append(("change_link", Name, NewName))
        self.links.remove(Name)
        if NewName != self.FullName:
            self.links.append(NewName)


class FakeExcelSession:
    def __init__(self, source_path, destination_names, source_names):
        self.source_path, self.destination_names, self.source_names = source_path, destination_names, source_names
        self.events, self.open_workbooks = [], []
        self.calculation_error = None
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
        if is_source:
            self.source = workbook
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
            target.sheets.insert(position + offset, FakeSheet(target, sheet.Name, sheet.payload))

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
        self.assert_inputs_unchanged()
        self.assert_no_candidate()

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
        self.assertIs(caught.exception, primary)
        if hasattr(primary, "add_note"):
            self.assertIn("candidate still locked", primary.__notes__[0])
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
