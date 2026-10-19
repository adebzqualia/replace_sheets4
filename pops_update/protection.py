"""Temporarily unlock worksheets for Excel's workbook-wide ChangeLink call."""
from contextlib import contextmanager

from .excel_session import ExcelSessionError


ALLOW_OPTIONS = (
    "AllowFormattingCells", "AllowFormattingColumns", "AllowFormattingRows",
    "AllowInsertingColumns", "AllowInsertingRows", "AllowInsertingHyperlinks",
    "AllowDeletingColumns", "AllowDeletingRows", "AllowSorting",
    "AllowFiltering", "AllowUsingPivotTables",
)


def _settings(worksheet):
    options = {
        "DrawingObjects": bool(worksheet.ProtectDrawingObjects),
        "Contents": bool(worksheet.ProtectContents),
        "Scenarios": bool(worksheet.ProtectScenarios),
    }
    if not any(options.values()):
        return None
    options["UserInterfaceOnly"] = bool(worksheet.ProtectionMode)
    protection = worksheet.Protection
    options.update((name, bool(getattr(protection, name))) for name in ALLOW_OPTIONS)
    return options, worksheet.EnableSelection


@contextmanager
def unprotected_worksheets(session, workbook, password):
    """Restore existing protection, including permissions and passwordless sheets."""
    unlocked = []
    primary = None
    try:
        for index in range(1, int(workbook.Worksheets.Count) + 1):
            worksheet = workbook.Worksheets(index)
            label = str(worksheet.Name)
            settings = session.run(f"Reading protection of worksheet {label!r}",
                                   lambda ws=worksheet: _settings(ws))
            if settings is None:
                continue
            # Excel ignores a supplied password on a passwordless sheet. Try
            # empty first so restoring protection does not introduce a password.
            used_password = ""
            try:
                session.call(lambda ws=worksheet: ws.Unprotect(Password=""))
            except Exception:
                used_password = password
                session.run(f"Unprotecting worksheet {label!r}; check sheet_password",
                            lambda ws=worksheet: ws.Unprotect(Password=password))
            options, selection = settings
            unlocked.append((worksheet, label, used_password, options, selection))
        yield
    except BaseException as exc:
        primary = exc
        raise
    finally:
        errors = []
        for worksheet, label, used_password, options, selection in reversed(unlocked):
            try:
                session.run(f"Restoring protection of worksheet {label!r}",
                            lambda ws=worksheet, key=used_password, opts=options:
                            ws.Protect(Password=key, **opts))
                session.run(f"Restoring selection settings of worksheet {label!r}",
                            lambda ws=worksheet, value=selection: setattr(ws, "EnableSelection", value))
            except Exception as exc:
                errors.append(exc)
        if errors:
            details = "; ".join(str(error) for error in errors)
            if primary is None:
                raise ExcelSessionError(details) from errors[0]
            if hasattr(primary, "add_note"):
                primary.add_note("Worksheet protection cleanup: " + details)
