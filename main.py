"""Replace configured country sheets in a copy of a final POPS workbook."""
from __future__ import annotations

import argparse
from pathlib import Path
import sys
import traceback

from pops_update.config import load_config
from pops_update.backend import build_plan, refresh
from pops_update.excel_session import ExcelSession


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--config', type=Path, default=Path(__file__).resolve().parent / 'config' / 'update.yaml')
    parser.add_argument('--debug', action='store_true', help='Show the full traceback if an operation fails')
    modes = parser.add_mutually_exclusive_group()
    modes.add_argument('--validate-config', action='store_true', help='Validate settings without requiring workbook files or Excel')
    modes.add_argument('--dry-run', action='store_true', help='Check file paths and display the selected sheets without starting Excel')
    modes.add_argument('--check-excel', action='store_true', help='Check isolated desktop Excel automation without opening reporting workbooks')
    args = parser.parse_args()
    try:
        if args.check_excel:
            with ExcelSession() as session:
                print(f'Microsoft Excel {session.app.Version} is available.')
            return 0
        config = load_config(args.config, check_files=not args.validate_config)
        if args.validate_config:
            print(f'Configuration is valid: {len(config.updates)} country updates.')
            return 0
        plan = build_plan(config)
        for country in plan:
            for sheet in country.sheets:
                print(f'{country.country}: {sheet.source_name} -> {sheet.target_name}')
        if args.dry_run:
            print('Update plan shown. No files changed. Excel will open and replace the selected sheets during the update.')
            return 0
        result = refresh(config, plan, progress=lambda text: print(text, flush=True))
        print(f'Created: {result["output_workbook"]}')
        print(f'Update report: {result["report_file"]}')
        for warning in result['warnings']:
            print(f'NOTE: {warning}')
        return 0
    except KeyboardInterrupt:
        print('Interrupted. Input workbooks were not saved.', file=sys.stderr)
        return 130
    except Exception as exc:
        print(f'Update failed: {exc}', file=sys.stderr)
        if args.debug:
            traceback.print_exc()
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
