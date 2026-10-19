"""Optional end-to-end worksheet replacement against Microsoft Excel desktop.

Run on Windows with Excel installed:
    $env:POPS_RUN_EXCEL_TESTS = '1'
    python -m unittest discover -s tests -p test_excel_integration.py -v
"""
from __future__ import annotations

from hashlib import sha256
import os
from pathlib import Path
import tempfile
import unittest

from pops_update.backend import build_plan, refresh
from pops_update.config import CountryUpdate, SheetUpdate, UpdateConfig
from pops_update.excel_session import ExcelSession


ALLOW_OPTIONS = (
    "AllowFormattingCells", "AllowFormattingColumns", "AllowFormattingRows",
    "AllowInsertingColumns", "AllowInsertingRows", "AllowInsertingHyperlinks",
    "AllowDeletingColumns", "AllowDeletingRows", "AllowSorting", "AllowFiltering",
    "AllowUsingPivotTables",
)


@unittest.skipUnless(os.name == "nt" and os.environ.get("POPS_RUN_EXCEL_TESTS") == "1",
                     "Set POPS_RUN_EXCEL_TESTS=1 on Windows with Excel desktop to run")
class ExcelReplacementTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)

    def create_workbook(self, session, names):
        workbook = session.track_workbook(session.call(lambda: session.app.Workbooks.Add(-4167)))
        workbook.Worksheets(1).Name = names[-1]
        for name in reversed(names[:-1]):
            sheet = session.call(lambda: workbook.Worksheets.Add(
                Before=workbook.Worksheets(1)))
            sheet.Name = name
        return workbook

    def save_fixture(self, session, workbook, path):
        session.call(lambda: session.app.CalculateFullRebuild())
        session.call(lambda: workbook.SaveAs(Filename=str(path), FileFormat=51))
        session.close_workbook(workbook)

    def protect_fixture(self, sheet, password, *, objects=True, scenarios=True, selection=0, **permissions):
        sheet.Protect(Password=password, DrawingObjects=objects, Contents=True,
                      Scenarios=scenarios, UserInterfaceOnly=False,
                      **{option: permissions.get(option, False) for option in ALLOW_OPTIONS})
        sheet.EnableSelection = selection

    def protection_state(self, sheet):
        return {
            "DrawingObjects": bool(sheet.ProtectDrawingObjects), "Contents": bool(sheet.ProtectContents),
            "Scenarios": bool(sheet.ProtectScenarios), "UserInterfaceOnly": bool(sheet.ProtectionMode),
            "EnableSelection": int(sheet.EnableSelection),
            **{option: bool(getattr(sheet.Protection, option)) for option in ALLOW_OPTIONS},
        }

    def test_replaces_whole_sheets_and_recalculates_existing_dependents(self):
        final, source = self.root / "final.xlsx", self.root / "portugal.xlsx"
        original_names = ("Summary", "Portugal_ID Card", "Portugal_FTE", "Portugal_NPL SALES", "Static")
        with ExcelSession() as session:
            workbook = self.create_workbook(session, original_names)
            workbook.Worksheets("Portugal_ID Card").Range("A1:B1").Value2 = ((1, 2),)
            workbook.Worksheets("Portugal_NPL SALES").Range("A1").Value2 = 3
            workbook.Worksheets("Portugal_FTE").Range("A1").Value2 = 7
            summary = workbook.Worksheets("Summary")
            summary.Range("A1").Formula = "=SUM('Portugal_ID Card'!A1:B1)"
            summary.Range("A2").Formula = "='Portugal_NPL SALES'!A1"
            summary.Range("A3").Formula = "='Portugal_ID Card'!A1"
            summary.Range("D4").Formula = "=#REF!"
            workbook.Worksheets("Static").Range("A1").Value2 = "Keep this sheet"
            self.protect_fixture(summary, "a", objects=False, selection=1,
                                 AllowFormattingCells=True, AllowSorting=True, AllowFiltering=True)
            self.save_fixture(session, workbook, final)

            workbook = self.create_workbook(session, ("ID Card", "NPL SALES", "FTE"))
            card = workbook.Worksheets("ID Card")
            card.Range("A1:B1").Value2 = ((10, 20),)
            card.Range("A1").Font.Bold = True
            card.Range("A1").Interior.Color = 65535
            card.Range("A1").NumberFormat = "0.00"
            card.Columns("A").ColumnWidth = 23.5
            expected_width = float(card.Columns("A").ColumnWidth)
            card.Range("D4:E4").Merge()
            card.Range("D4").Value2 = "New source layout"
            sales = workbook.Worksheets("NPL SALES")
            sales.Range("A1").Formula = "=SUM('ID Card'!A1:B1)"
            sales.Range("B1").Formula = "='ID Card'!A1*2"
            sales.Range("C1").Formula = "=FTE!A1"
            workbook.Worksheets("FTE").Range("A1").Value2 = 99
            self.protect_fixture(card, "", scenarios=False, selection=-4142,
                                 AllowFormattingColumns=True, AllowInsertingRows=True)
            self.protect_fixture(sales, "a", AllowFormattingRows=True,
                                 AllowInsertingHyperlinks=True, AllowUsingPivotTables=True)
            self.save_fixture(session, workbook, source)

            # Capture what Excel actually persisted, including settings whose
            # runtime values may change across save/reopen on different versions.
            workbook = session.open_workbook(final, read_only=True)
            expected_protection = {"Summary": self.protection_state(workbook.Worksheets("Summary"))}
            session.close_workbook(workbook)
            workbook = session.open_workbook(source, read_only=True)
            for name in ("ID Card", "NPL SALES"):
                expected_protection["Portugal_" + name] = self.protection_state(workbook.Worksheets(name))
            session.close_workbook(workbook)

        before = {path: sha256(path.read_bytes()).digest() for path in (final, source)}
        sheets = (SheetUpdate("NPL SALES", "Portugal_NPL SALES"),
                  SheetUpdate("ID Card", "Portugal_ID Card"))
        update = CountryUpdate("Portugal", "Portugal", source, sheets)
        config = UpdateConfig(final_workbook=final, output_workbook=self.root / "output" / "updated.xlsx",
                              overwrite_output=False, check_text_layout=True, sheet_password="a",
                              updates=(update,), sheet_mapping=(("ID Card", "ID Card"),
                                                               ("NPL SALES", "NPL SALES"), ("FTE", "FTE")))
        result = refresh(config, build_plan(config))
        self.assertEqual(Path(result["output_workbook"]), config.output_workbook)
        for path, digest in before.items():
            self.assertEqual(sha256(path.read_bytes()).digest(), digest)

        with ExcelSession() as session:
            workbook = session.open_workbook(config.output_workbook, read_only=True)
            self.assertEqual(tuple(workbook.Worksheets(i).Name
                                   for i in range(1, workbook.Worksheets.Count + 1)), original_names)
            summary = workbook.Worksheets("Summary")
            self.assertEqual(summary.Range("A1:A3").Value2, ((30.0,), (30.0,), (10.0,)))
            self.assertIn("#REF!", summary.Range("D4").Formula)
            self.assertEqual(workbook.Worksheets("Static").Range("A1").Value2, "Keep this sheet")
            self.assertEqual(workbook.Worksheets("Portugal_FTE").Range("A1").Value2, 7)
            sales = workbook.Worksheets("Portugal_NPL SALES")
            # An unselected helper remains linked to its original source.
            self.assertEqual(sales.Range("A1:C1").Value2, ((30.0, 20.0, 99.0),))
            card = workbook.Worksheets("Portugal_ID Card")
            self.assertTrue(card.Range("A1").Font.Bold)
            self.assertEqual(card.Range("A1").Interior.Color, 65535)
            self.assertEqual(card.Range("A1").NumberFormat, "0.00")
            self.assertAlmostEqual(float(card.Columns("A").ColumnWidth), expected_width, places=2)
            self.assertTrue(card.Range("D4").MergeCells)
            self.assertEqual(card.Range("D4").MergeArea.Address, "$D$4:$E$4")
            self.assertEqual(card.Range("D4").Value2, "New source layout")
            links = workbook.LinkSources(1) or ()
            if isinstance(links, str):
                links = (links,)
            self.assertEqual(tuple(Path(link).name.casefold() for link in links), (source.name.casefold(),),
                             "Only the existing unselected-helper source link should remain")
            for address in ("A1", "A2", "A3"):
                self.assertNotIn("[", summary.Range(address).Formula)
            for address in ("A1", "B1"):
                self.assertNotIn("[", sales.Range(address).Formula)
            self.assertIn(source.name, sales.Range("C1").Formula)
            self.assertIn("FTE", sales.Range("C1").Formula)
            for name, state in expected_protection.items():
                self.assertEqual(self.protection_state(workbook.Worksheets(name)), state, name)
            self.assertFalse(workbook.Worksheets("Static").ProtectContents)
            self.assertFalse(workbook.Worksheets("Portugal_FTE").ProtectContents)
            # Passwordless imported protection stays passwordless; protected
            # sheets with passwords can still be unlocked using that password.
            card.Unprotect(Password="")
            self.assertFalse(card.ProtectContents)
            summary.Unprotect(Password="a")
            self.assertFalse(summary.ProtectContents)
            sales.Unprotect(Password="a")
            self.assertFalse(sales.ProtectContents)
        self.assertEqual(sorted(path.name for path in config.output_workbook.parent.glob("*.xlsx")),
                         [config.output_workbook.name])


if __name__ == "__main__":
    unittest.main()
