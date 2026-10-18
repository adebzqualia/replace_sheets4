# POPS Update

Replace selected country worksheets in a copy of the final POPS workbook, rename them to their consolidated names, and let desktop Excel recalculate the result.

**Author:** Adnane El Bouhali  
**Team:** Value Chain team — BNP Paribas Personal Finance

The corrected source sheets replace the complete selected destination sheets, including their cell contents and formatting. The existing tab order is restored. The original final workbook and corrected source files are closed without saving changes; the result is saved to a separate output path.

## Install

An update requires Windows, Python 3.10 or later, and installed desktop Microsoft Excel. Inputs and output use `.xlsx` files.

From PowerShell in this project directory:

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe main.py --check-excel
```

`--check-excel` starts an isolated Excel instance without opening reporting workbooks.

## Configure

Edit `config/update.yaml`. When installing from `POPS_Update_simple.zip`, keep your existing file; for a new installation, copy `config/update.example.yaml` to `config/update.yaml` first.

Paths are relative to the YAML file's directory. The supplied country selections and workbook paths are examples. A selection fragment looks like this:

```yaml
final_workbook: ../data/Consolidated_with_PPT.xlsx
output_workbook: ../output/Consolidated_with_PPT_updated.xlsx
overwrite_output: false
sheet_password: "a"

updates:
  - country: Germany
    file: ../data/updates/POPS2027-Germany.xlsx
    sheets: [ID Card, FTE]
  - country: Belgium
    file: ../data/updates/POPS2027-Belgium.xlsx
    sheets: [ID Card, FTE]
```

Keep the full file's `version`, `country_prefixes` and `sheet_mapping` sections. For example, country `Germany` and source sheet `ID Card` select destination `Germany_ID Card`. The selected source and destination sheets must exist. Omitting `sheets` selects every entry in `sheet_mapping`.

| Setting | Meaning |
| --- | --- |
| `final_workbook` | Existing final workbook used as the base. |
| `output_workbook` | Separate result path, different from all input files. |
| `overwrite_output` | Set `true` to replace an existing output. Close that output in Excel first. |
| `sheet_password` | Password used to unprotect workbook structure for moving and renaming tabs. Previously protected destination structure is protected again. |
| `country_prefixes` | Country selection name to consolidated prefix. |
| `sheet_mapping` | Source-sheet name to consolidated suffix. |
| `updates` | Countries, corrected files and selected source-sheet names. |
| `check_text_layout` | Legacy boolean accepted for existing configurations. It is ignored; either value gives the same update behavior. |

`Findomestic` and `Italy` both map to `Italy`; `PFF` and `France` both map to `PFF`. Use one alias per destination country in a run.

## Run

```powershell
# Validate YAML settings without requiring the workbook files
.\.venv\Scripts\python.exe main.py --validate-config

# Check paths and display the selected source/destination sheets
.\.venv\Scripts\python.exe main.py --dry-run

# Replace sheets, recalculate and save using Excel
.\.venv\Scripts\python.exe main.py
```

Use `--config 'C:/POPS/config/update.yaml'` to select another configuration file. A dry run does not open workbooks, inspect their contents or confirm that the selected tabs exist.

The output is accompanied by `<output-stem>.update_report.json`, listing the output path, selected countries and snapshot warning.

## What Excel does

1. Open a temporary copy of the final workbook for editing and each corrected source read-only.
2. Rename only the selected source tabs in memory to their consolidated names.
3. Move the old destination tabs into a temporary scratch workbook, then copy the selected source tabs together into the destination. Moving the old tabs lets Excel retain downstream references during replacement.
4. Use Excel's native `ChangeLink` to point references to the moved tabs back at their replacements, then restore the original tab order.
5. Run Excel's `CalculateFullRebuild`, save and close the candidate, then publish it to the configured output.

Python orchestrates the Excel operations. It does not calculate business results, rewrite formula text, scan for formula errors, compare layouts or reject a workbook because it contains `#REF!`.

The selected source tabs are copied as a group so their mutual references can remain internal. References to unselected source tabs or helper sheets can remain normal external Excel links to the corrected source workbook; those links are preserved. The updater does not refresh external workbooks or data connections.

Workbooks open with `UpdateLinks=0`. Automatic embedded-object updates are also disabled on opened and temporary workbooks using `UpdateLinks=2` and `UpdateRemoteReferences=False`. The output retains these link-refresh settings; worksheet formulas still recalculate. Microsoft documents the [open-time external-link setting](https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open) and the [OLE/DDE recalculation setting](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.updateremotereferences).

## Excel errors and missing Office cache files

An error naming `INetCache\Content.MSO\....xlsx` means Excel attempted to access a missing cache workbook. Disabling automatic link refresh may avoid that lookup, but cannot recreate the file or repair a stale link. The error now identifies the operation, such as opening a specific input, copying sheets, reconnecting references or recalculating.

For the complete error location, run:

```powershell
python main.py --debug
```

If Excel still requires the missing workbook, use the reported operation and input file to identify the affected workbook. Links can occur in cells, names, objects and charts; see Microsoft's [workbook link management guide](https://support.microsoft.com/en-us/excel/manage-workbook-links). The updater does not automatically break links or replace their formulas with cached values.

## Result boundaries

Review the updated values and representative intermediary and synthesis results in Excel. Recalculation uses existing formulas and coordinates; it does not rebuild reporting mappings or repair pre-existing `#REF!` errors. A change in template structure can therefore require adjustments in the original POPS generators.

`FILLING_RATE`, `VALIDATION` and `PPT_VALIDATION` remain snapshots from their earlier generation. The updater does not regenerate them or create a PowerPoint file.

## Tests

Run the portable tests without Excel:

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests
```

To also run native Excel integration tests on a Windows machine with desktop Excel:

```powershell
$env:POPS_RUN_EXCEL_TESTS = '1'
.\.venv\Scripts\python.exe -m unittest discover -s tests
Remove-Item Env:POPS_RUN_EXCEL_TESTS
```

Native tests are opt-in. Desktop Excel was unavailable on the development machine, so the native workflow must be exercised on a machine with Excel; portable tests alone do not establish its behavior with production workbooks.
