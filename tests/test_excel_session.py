"""Native-call arguments and error context without an Excel installation."""
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import Mock

from pops_update.excel_session import ExcelSession, ExcelSessionError


class ExcelSessionTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.path = Path(temporary.name) / "report.xlsx"
        self.path.write_bytes(b"workbook fixture")
        self.workbook = SimpleNamespace(UpdateLinks=3, UpdateRemoteReferences=True,
                                        Application=SimpleNamespace(Hwnd=1))
        self.session = ExcelSession()
        self.session._pythoncom = SimpleNamespace(Empty=object())
        self.session.app = SimpleNamespace(
            Workbooks=SimpleNamespace(Open=Mock(return_value=self.workbook)),
            Hwnd=1, Calculation=-4105, CalculateBeforeSave=True)

    def test_open_disables_external_refresh_and_remembers_calculation(self):
        workbook = self.session.open_workbook(self.path, read_only=True)
        arguments = self.session.app.Workbooks.Open.call_args.kwargs
        self.assertEqual(arguments["Filename"], str(self.path.resolve()))
        self.assertEqual(arguments["UpdateLinks"], 0)
        self.assertTrue(arguments["ReadOnly"])
        self.assertEqual(workbook.UpdateLinks, 2)
        self.assertFalse(workbook.UpdateRemoteReferences)
        self.assertEqual(self.session._calculation_settings[id(workbook)], (-4105, True))
        self.assertEqual(self.session.app.Calculation, -4135)

    def test_scratch_workbook_also_disables_linked_object_refresh(self):
        self.session.track_workbook(self.workbook)
        self.assertEqual(self.workbook.UpdateLinks, 2)
        self.assertFalse(self.workbook.UpdateRemoteReferences)

    def test_open_error_identifies_input_and_keeps_missing_cache_path(self):
        cache = r"C:\Users\example\AppData\Local\Microsoft\Windows\INetCache\Content.MSO\8614C64A.xlsx"
        original = RuntimeError(f"Cannot find {cache}")
        self.session.app.Workbooks.Open.side_effect = original
        with self.assertRaises(ExcelSessionError) as caught:
            self.session.open_workbook(self.path)
        self.assertIn(f"Opening workbook {self.path.resolve()}", str(caught.exception))
        self.assertIn(cache, str(caught.exception))
        self.assertIs(caught.exception.__cause__, original)

    def test_copy_and_move_keep_explicit_before_argument(self):
        selection, anchor = Mock(), object()
        self.session.copy_after(selection, anchor)
        self.session.move_after(selection, anchor)
        selection.Copy.assert_called_once_with(self.session._pythoncom.Empty, anchor)
        selection.Move.assert_called_once_with(self.session._pythoncom.Empty, anchor)

    def test_copy_error_names_the_operation(self):
        selection = Mock()
        selection.Copy.side_effect = RuntimeError("Excel missing cached workbook")
        with self.assertRaisesRegex(ExcelSessionError, "Copying selected worksheets.*missing cached workbook"):
            self.session.copy_after(selection, object())


if __name__ == "__main__":
    unittest.main()
