# Protected worksheet fix

Extract `POPS_Update_protection_fix.zip` into the project folder containing
`main.py` and replace the included files. It updates `pops_update/backend.py`
and adds `pops_update/protection.py`. The archive also includes the related
tests. Your configuration and workbook data are not included or overwritten.

The update now temporarily unprotects protected destination worksheets before
Excel reconnects references. This includes copied country sheets and existing
summary sheets. It restores their protection flags, allowed actions and
selection setting after the link update, including when Excel raises an error.
Unprotected sheets stay unprotected; passwordless protection stays passwordless.

The existing YAML setting supplies the worksheet password as well as the
workbook structure password:

```yaml
sheet_password: "a"
```

Keep that value if it is the correct password for your protected sheets.
If a sheet uses a different password, the error identifies the sheet and
instructs you to check `sheet_password`. The updater does not bypass passwords.
If unlocking or restoring protection fails, it does not publish the candidate
or replace an existing output.

Run from your activated environment:

```powershell
python main.py --debug
```

Portable regression tests cover protected existing/copied sheets, permission
restoration, wrong passwords and failed updates. The optional native Excel
test also includes protected worksheets; it requires desktop Excel and was
not run on the development machine.
