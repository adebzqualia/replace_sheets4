"""Replace selected worksheets using native Excel copy, move and calculation."""
from __future__ import annotations

import json
import os
from pathlib import Path
import shutil
import tempfile
from uuid import uuid4

from .config import UpdateConfig
from .excel_session import ExcelSession


def build_plan(config: UpdateConfig):
    """Display the configured selection; Excel handles the workbook contents."""
    return config.updates


def _selected(workbook, names):
    return workbook.Worksheets(names[0] if len(names) == 1 else tuple(names))


def _reconnect(session, destination, previous):
    """Ask Excel to point references to moved tabs at their replacements."""
    links = session.call(lambda: destination.LinkSources(1)) or ()
    if isinstance(links, str):
        links = (links,)
    previous_path = os.path.normcase(os.path.normpath(str(previous.FullName)))
    for link in links:
        if (os.path.normcase(os.path.normpath(str(link))) == previous_path
                or str(link).casefold() == str(previous.Name).casefold()):
            session.run("Reconnecting references to replacement sheets", lambda link=link: destination.ChangeLink(
                Name=link, NewName=str(destination.FullName), Type=1))


def _replace_country(session, destination, previous, update, password, progress):
    if progress:
        progress(f"Opening {update.country}: {update.file}")
    source = session.open_workbook(update.file, read_only=True)
    try:
        if source.ProtectStructure:
            session.call(lambda: source.Unprotect(Password=password))
        targets = [sheet.target_name for sheet in update.sheets]
        # Only selected tabs are renamed. References to unselected source tabs
        # retain their original names and Excel's normal external-link behavior.
        selected = [(source.Worksheets(sheet.source_name), sheet.target_name)
                    for sheet in update.sheets]
        for worksheet, _ in selected:
            temporary = "__POPS_" + uuid4().hex[:12]
            session.call(lambda ws=worksheet, name=temporary: setattr(ws, "Name", name))
        for worksheet, name in selected:
            session.call(lambda ws=worksheet, name=name: setattr(ws, "Name", name))

        # Moving instead of deleting lets Excel retain downstream references.
        if progress:
            progress(f"Replacing {update.country}: {', '.join(targets)}")
        session.move_after(_selected(destination, targets), previous.Sheets(1))
        session.copy_after(_selected(source, targets), destination.Sheets(destination.Sheets.Count))
        _reconnect(session, destination, previous)
        if progress:
            for sheet in update.sheets:
                progress(f"Replaced {update.country}: {sheet.source_name} -> {sheet.target_name}")
    finally:
        session.close_workbook(source)


def refresh(config: UpdateConfig, plan, progress=None):
    """Copy the final workbook, replace its selected tabs and let Excel calculate."""
    output = config.output_workbook
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists() and not config.overwrite_output:
        raise FileExistsError(f"Output already exists: {output}")
    candidate = output.with_name(f".{output.stem}.candidate-{uuid4().hex[:12]}.xlsx")
    try:
        shutil.copy2(config.final_workbook, candidate)
        with tempfile.TemporaryDirectory(prefix="pops-update-", ignore_cleanup_errors=True) as scratch:
            with ExcelSession() as session:
                if progress:
                    progress(f"Opening working copy of: {config.final_workbook}")
                destination = session.open_workbook(candidate, read_only=False)
                was_protected = bool(destination.ProtectStructure)
                windows_protected = bool(destination.ProtectWindows)
                if was_protected:
                    session.call(lambda: destination.Unprotect(Password=config.sheet_password))
                order = [str(destination.Sheets(i).Name)
                         for i in range(1, int(destination.Sheets.Count) + 1)]
                # This visible placeholder also permits replacing every tab.
                anchor = session.add_sheet_after(destination, destination.Sheets(destination.Sheets.Count))
                previous = session.track_workbook(session.call(lambda: session.app.Workbooks.Add(-4167)))
                session.call(lambda: previous.SaveAs(
                    Filename=str(Path(scratch) / "previous_sheets.xlsx"), FileFormat=51))
                for update in plan:
                    _replace_country(session, destination, previous, update, config.sheet_password, progress)
                last = anchor
                for name in order:
                    worksheet = destination.Sheets(name)
                    session.move_after(worksheet, last)
                    last = worksheet
                session.call(lambda: anchor.Delete())
                session.close_workbook(previous)
                if was_protected:
                    session.call(lambda: destination.Protect(
                        Password=config.sheet_password, Structure=True, Windows=windows_protected))
                session.restore_calculation(destination)
                if progress:
                    progress("Recalculating and saving with Excel...")
                session.run("Recalculating workbook", lambda: session.app.CalculateFullRebuild())
                session.run("Saving updated workbook", lambda: destination.Save())
                session.close_workbook(destination)
        if output.exists() and not config.overwrite_output:
            raise FileExistsError(f"Output already exists: {output}")
        os.replace(candidate, output)
        report = output.with_name(f"{output.stem}.update_report.json")
        result = {
            "output_workbook": str(output), "report_file": str(report),
            "countries": [update.country for update in plan],
            "warnings": ["FILLING_RATE, VALIDATION and PPT_VALIDATION remain generation snapshots."],
        }
        report.write_text(json.dumps(result, indent=2), encoding="utf-8")
        return result
    except BaseException as primary:
        try:
            candidate.unlink(missing_ok=True)
        except OSError as cleanup_error:
            if hasattr(primary, "add_note"):
                primary.add_note(f"Could not remove temporary workbook {candidate}: {cleanup_error}")
        raise
