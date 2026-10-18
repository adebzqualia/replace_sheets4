# Apply the simplified sheet-replacement update

Extract `POPS_Update_simple.zip` into the project folder containing `main.py`
(for example, `POPS_REPLACE_SHEETS2`) and replace the supplied code files.
The archive contains the project code, tests and an example configuration.
It excludes your `config/update.yaml` and workbook data, so your current paths
and country selections are preserved. For a new installation, copy
`config/update.example.yaml` to `config/update.yaml` and configure it.

From your activated Python environment, run:

```powershell
python -m pip install -r requirements.txt
python main.py --check-excel
python main.py --dry-run
python main.py
```

The dry run checks paths and displays the selection without opening workbooks.
The update replaces complete selected worksheets, renames them to the
consolidated names, restores tab order, and lets desktop Excel recalculate and
save a separate output. The input files are not saved.

Old destination tabs are moved into a scratch workbook before the selected
source tabs are copied as a group. Excel's native `ChangeLink` reconnects
references to the replacements. Only selected source tabs are renamed;
references to unselected source or helper tabs can remain external links.

There are no formula scans, layout comparisons or `#REF!` rejection checks.
Python performs no business calculations, and pre-existing `#REF!` errors are
not repaired. You may leave `check_text_layout` in your YAML: it is accepted
for compatibility and ignored, regardless of its value. `sheet_password` is
used for workbook structure protection.

This version also disables automatic external-link updates when opening files
and automatic OLE/DDE updates during recalculation. These settings can avoid
lookups of expired Office cache files such as `INetCache\Content.MSO\....xlsx`;
they do not repair a stale workbook link. If Excel still fails, run
`python main.py --debug` to see the operation, affected input and full traceback.
Do not combine this package with the earlier `POPS_Update_fix.zip` patch.

Run the portable tests with:

```powershell
python -m unittest discover -s tests
```

Native Excel integration tests are opt-in:

```powershell
$env:POPS_RUN_EXCEL_TESTS = '1'
python -m unittest discover -s tests
Remove-Item Env:POPS_RUN_EXCEL_TESTS
```

Desktop Excel was unavailable on the development machine. Run the native tests
and review the resulting workbook on a machine with Excel. `FILLING_RATE`,
`VALIDATION` and `PPT_VALIDATION` remain generation snapshots.
