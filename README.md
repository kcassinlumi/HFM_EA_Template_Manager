# HFM EA Template Manager

A Windows GUI for managing HFM **Extended Analytics (EA) extract templates** directly in the HFM database. You can browse, edit, create, rename, delete and copy templates between users or applications, and export or import them as files.

It is built with PowerShell 5.1 and WPF, so nothing needs to be installed. It runs on any Windows server with .NET Framework 4.x, typically the HFM server itself.

## Starting

**On another server:** copy the single file `dist\EATemplateManager.ps1` anywhere, then right-click it and choose **Run with PowerShell**, or run:

```powershell
powershell -ExecutionPolicy Bypass -File .\EATemplateManager.ps1
```

That one file contains everything: the data module, the GUI and the window layouts. It creates `Backups\` and the log next to itself.

**In this source folder:** double-click `Start-EATemplateManager.cmd`.

### Building the single file

The source is split into `EATemplate.psm1`, `EATemplateManager.ps1` and `Xaml\*.xaml`. After changing any of them, rebuild with:

```powershell
powershell -ExecutionPolicy Bypass -File .\Build-SingleFile.ps1
```

The Connect dialog opens first. Enter the SQL Server and the HFM database (e.g. `localhost` / `hfm`), and choose either **Windows authentication** or a **SQL login**.

After a successful connect, these are remembered for each Windows user in `%APPDATA%\EATemplateManager\preferences.json`:
- the server and database
- the authentication type and SQL login name
- the last application and user

Next time the dialog is pre-filled and the tool reopens where you left off. **The password is never stored.** Delete that file to reset.

### Database permissions needed

| Object | Permission | Used for |
|---|---|---|
| `<APP>_USERPARAMS` | SELECT, INSERT, DELETE | the templates themselves |
| `HSX_DATASOURCES` | SELECT | list of HFM applications (falls back to `sys.tables`) |
| `HSX_DSN_DETAILS` | SELECT (only `DsnName` is read) | EA destination DSN drop-down |
| `HSV_ACTIVITY_USERS`, `HSV_ACTIVITY_SESSIONS` | SELECT | list of HFM users, plus the "user is logged in" warning |

Only `<APP>_USERPARAMS` is required. If the other tables can't be read, the related feature is quietly skipped.

## Using it

1. Pick the **Application** and **User**. The user list shows users who already have EA templates. To work on any other user, type the name exactly as HFM stores it (e.g. `jsmith@Native Directory`) and press Enter.
2. Select a template on the left. The editor has three tabs:
   - **POV:** one line per dimension (Scenario, Year, Period, View, Entity, Value, Account, ICP, Custom1..n). Each expression starts with `#` for members (`#Jan;Feb;Mar`) or `{` for member lists (`{Top.[Hierarchy]}`). You can combine several with `;` (`{[Base]};{[Hierarchy]}`). Errors are shown as you type, and the full resulting POV is shown below the lines. Use **+ Custom / − Custom** to add or remove Custom dimensions.
   - **Options:** extract format, DSN, table prefix, file prefix, delimiter, export option, line item detail, and the data flags. Any option the tool doesn't know about is kept unchanged.
   - **Raw XML:** the exact XML that will be stored. Tick *Edit raw XML* to edit it directly (advanced).
3. **Save** writes the change (Ctrl+S). **Revert** reloads it from the database. **Save as new...** stores the editor contents under a new name.
4. List actions:
   - **New...:** creates a template.
   - **Rename...:** renames one template.
   - **Delete:** deletes the selected templates.
   - **Copy to...:** copies the selected templates to one or more users, in this application or another. If a template with that name already exists, you choose to skip it, overwrite it, or keep both with a `_copy` suffix.
   - **Export... / Import...:** saves the selected templates to an XML file, or loads them into chosen users. Use these to move templates between environments.
   - Ctrl/Shift-click selects several templates at once.
5. `SS_DEFAULT` (shown in italics) holds the user's last-used EA settings. It is read-only here, but **Save as new...** turns it into a normal template.

## How HFM stores EA templates

Each application has an `<APP>_USERPARAMS (Username, EntryIDX, ParameterKey, BLOBData varbinary(2000))` table. All three key columns form the primary key.

- Values are UTF-16LE text. Anything longer than 2000 bytes is split over several rows (EntryIDX 1, 2, 3...).
- **`SS_PREF`** is each user's template index. HFM only lists templates that appear here:
  `<?xml version = '1.0' encoding = 'UTF-16LE'?>` + `<SS_TEMPLATES><SS_TEMPLATE name="ACT" key="SS_PREF_ACT"/>...</SS_TEMPLATES>`
- **`SS_PREF_<name>`** holds one template:
  `<povTemplate version="11.2"><povEA><![CDATA[S#Actual.Y#2015.P#Jan...]]></povEA><options>...</options></povTemplate>`
- **`SS_DEFAULT`** holds the last-used settings, in the same format.

The tool keeps the index and the template rows in step and writes HFM's own compact format. Templates you don't change are written back byte-for-byte unchanged. If a template row is missing from the index (or the reverse), the list flags it; **Tools > Repair template index** fixes it.

## Safety

- Every change runs in a single SQL transaction. If any step fails, nothing is written.
- Before every change, the affected rows are saved to `Backups\<timestamp>_<app>_<action>.csv`. **File > Back up all EA rows** saves every EA row of an application.
- To restore a backup, use **File > Restore from backup...** and pick the CSV. It puts every key in the file back to its backed-up state, including removing keys that didn't exist at the time of the backup. The current rows are backed up first. From the command line (source folder only):

  ```powershell
  Import-Module .\EATemplate.psm1
  Connect-EaDatabase -Server localhost -Database hfm
  Restore-EaBackup -Path .\Backups\<file>.csv
  ```
- If the selected user has an active HFM session, a warning is shown. HFM may cache user preferences, so the user should log off and back on to see changes (and to avoid HFM overwriting them).
- Errors are written to `EATemplateManager.log`.

## Things to confirm on your system

- **Extract format labels:** the order 0 Standard, 1 Metadata Only, 2 Selected Metadata Only, 3 SQL Aggregation, 4 Essbase, 5 Data Warehouse matches HFM's own resource strings (`FinancialManagement\bin\xfm.en-us.csv`). 6 Flat File and 7 Flat File without Header match existing templates. They can be overridden with `ExtractFormats` in `EATemplateManager.config.json`.
- **Line item detail values:** only `EA_LINEITEM_EXCLUDE` has been seen in real data. The other values in `LineItemOptions` are suggestions; the box also accepts any typed value.
- **Picking up changes:** after changing a template in this tool, open HFM > Consolidation > Extract > Extended Analytics as that user and check that it loads and runs.

## Tests

```powershell
.\Tests\Test-EATemplateCodec.ps1        # read-only: every stored template/index must round-trip byte-for-byte
.\Tests\Test-EATemplateIntegration.ps1  # writes, but only to throwaway users zz_eatm_test1/2 in DLLSIM, then removes them
```

Both accept `-Server` and `-Database`; the integration test also accepts `-App`.

## Files

| File | Purpose |
|---|---|
| `dist\EATemplateManager.ps1` | **single-file build to copy to other servers** (generated; don't edit) |
| `Build-SingleFile.ps1` | regenerates the single file from the sources below |
| `Start-EATemplateManager.cmd` | launcher for the source version |
| `EATemplateManager.ps1` | WPF front end |
| `EATemplate.psm1` | data layer (SQL, chunking, index, XML/POV parsing, copy/export/import, backups); can also be used from the PowerShell command line |
| `Xaml\*.xaml` | window layouts |
| `EATemplateManager.config.json` | optional, shared by everyone using that copy: overrides for extract format labels and line item options, and default connection values. Built-in defaults are used when it's absent; personal preferences in `%APPDATA%` take priority over it |
