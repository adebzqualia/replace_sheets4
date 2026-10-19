"""Protection restoration and failure behavior without installed Excel."""
from copy import deepcopy
from types import SimpleNamespace
import unittest

from pops_update.excel_session import ExcelSession, ExcelSessionError
from pops_update.protection import ALLOW_OPTIONS, unprotected_worksheets


class Worksheet:
    def __init__(self, name, password="a", protected=True):
        self.Name, self.password = name, password
        self.ProtectContents = protected
        self.ProtectDrawingObjects = protected
        self.ProtectScenarios = False
        self.ProtectionMode = protected
        self.Protection = SimpleNamespace(**{key: index % 2 == 0 for index, key in enumerate(ALLOW_OPTIONS)})
        self.EnableSelection = 1
        self.events = []
        self.fail_restore = False

    def Unprotect(self, Password):
        self.events.append("unprotect")
        if self.password and Password != self.password:
            raise RuntimeError("Incorrect password")
        self.ProtectContents = self.ProtectDrawingObjects = self.ProtectScenarios = False
        self.ProtectionMode = False
        self.EnableSelection = 0

    def Protect(self, Password, **options):
        self.events.append("protect")
        if self.fail_restore:
            raise RuntimeError("Excel could not restore protection")
        self.password = Password
        self.ProtectContents = options["Contents"]
        self.ProtectDrawingObjects = options["DrawingObjects"]
        self.ProtectScenarios = options["Scenarios"]
        self.ProtectionMode = options["UserInterfaceOnly"]
        self.Protection = SimpleNamespace(**{key: options[key] for key in ALLOW_OPTIONS})
        self.EnableSelection = 0

    def state(self):
        return (self.password, self.ProtectContents, self.ProtectDrawingObjects,
                self.ProtectScenarios, self.ProtectionMode, vars(self.Protection), self.EnableSelection)


class Sheets:
    def __init__(self, sheets):
        self.sheets = sheets
        self.Count = len(sheets)

    def __call__(self, index):
        return self.sheets[index - 1]


class ProtectionTests(unittest.TestCase):
    def setUp(self):
        self.session = ExcelSession(retry_delay=0)

    def workbook(self, *sheets):
        return SimpleNamespace(Worksheets=Sheets(sheets))

    def test_restores_permissions_selection_and_passwordless_protection(self):
        sheets = (Worksheet("Summary"), Worksheet("Copied sheet", password=""),
                  Worksheet("Unprotected", password="", protected=False))
        before = [deepcopy(sheet.state()) for sheet in sheets]
        with unprotected_worksheets(self.session, self.workbook(*sheets), "a"):
            self.assertTrue(all(not sheet.ProtectContents for sheet in sheets))
        self.assertEqual([sheet.state() for sheet in sheets], before)
        self.assertEqual(sheets[-1].events, [])

    def test_object_only_protection_is_also_restored(self):
        sheet = Worksheet("Objects")
        sheet.ProtectContents = False
        before = deepcopy(sheet.state())
        with unprotected_worksheets(self.session, self.workbook(sheet), "a"):
            self.assertFalse(sheet.ProtectDrawingObjects)
        self.assertEqual(sheet.state(), before)

    def test_wrong_password_restores_already_unlocked_sheets_before_raising(self):
        first, locked = Worksheet("First"), Worksheet("Other password", password="different")
        before = deepcopy(first.state())
        with self.assertRaisesRegex(ExcelSessionError, "Other password.*check sheet_password"):
            with unprotected_worksheets(self.session, self.workbook(first, locked), "a"):
                self.fail("The link update must not run after an unlock failure")
        self.assertEqual(first.state(), before)
        self.assertTrue(locked.ProtectContents)
        self.assertNotIn("protect", locked.events)

    def test_failed_link_update_restores_all_sheets_and_keeps_original_error(self):
        sheet = Worksheet("Summary")
        before = deepcopy(sheet.state())
        primary = RuntimeError("ChangeLink failed")
        with self.assertRaises(RuntimeError) as caught:
            with unprotected_worksheets(self.session, self.workbook(sheet), "a"):
                raise primary
        self.assertIs(caught.exception, primary)
        self.assertEqual(sheet.state(), before)

    def test_restore_failure_is_fatal_and_other_sheets_are_still_restored(self):
        first, failing = Worksheet("First"), Worksheet("Restore failure")
        failing.fail_restore = True
        with self.assertRaisesRegex(ExcelSessionError, "Restore failure"):
            with unprotected_worksheets(self.session, self.workbook(first, failing), "a"):
                pass
        self.assertTrue(first.ProtectContents)

    def test_restore_failure_does_not_hide_a_link_update_failure(self):
        sheet = Worksheet("Summary")
        sheet.fail_restore = True
        primary = RuntimeError("Original ChangeLink error")
        with self.assertRaises(RuntimeError) as caught:
            with unprotected_worksheets(self.session, self.workbook(sheet), "a"):
                raise primary
        self.assertIs(caught.exception, primary)
        if hasattr(primary, "add_note"):
            self.assertIn("Restoring protection", primary.__notes__[0])


if __name__ == "__main__":
    unittest.main()
