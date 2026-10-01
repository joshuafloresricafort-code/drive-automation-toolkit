# Google Drive Automation Toolkit

Google Apps Script tools for the bulk Drive and Sheets jobs that come up in reporting work: pulling data out of hundreds of Excel files, cleaning up file names, and keeping long-running jobs inside Apps Script's time limits.

| Tool | Folder | What it does |
|---|---|---|
| Excel Property Description ETL | [`excel-property-description-etl`](excel-property-description-etl) | Walks a Drive folder tree, converts Excel files to Google Sheets, pulls out every unique property description, and writes the results to a central sheet |
| Bulk "Copy of" Remover | [`bulk-rename-copy-of`](bulk-rename-copy-of) | Strips the `Copy of ` prefix from every file in a folder |

More tools will be added to this repo over time.

---

## Excel Property Description ETL

Built at Community Minerals to build a master list of property descriptions from hundreds of operator revenue statements, each one an Excel file buried somewhere in a nested Drive folder.

**Pipeline**

1. **Discover:** recursively collects every `.xlsx` / `.xls` file under a root folder
2. **Filter:** optionally processes only files named on a `Community Minerals_DATA` list (names are normalised by stripping `Copy of ` and file extensions)
3. **Convert:** turns each Excel file into a temporary Google Sheet through the Drive advanced service, retrying up to five times
4. **Extract:** finds the description column even when its header varies (`Property Description`, `Property Name`, `CC Description`, `Property name`) and removes duplicate values
5. **Load:** appends `[description, source file]` rows to the `Data` sheet, then trashes the temporary file

**Built to run inside Apps Script limits**
- **Timeout guard:** stops cleanly at about 4.5 minutes, well inside Apps Script's six-minute execution limit
- **Batching:** handles up to 10 files per run
- **Checkpointing:** saves the last processed folder in Script Properties, so the next run carries on from there
- **Self-scheduling:** an optional five-minute time-based trigger keeps processing batches until the folder is done
- **Idempotent:** files already in the `Data` sheet are skipped, so re-running never duplicates rows
- **Audit trail:** every run writes to its own timestamped `Logs_` sheet, and one menu item removes them all
- **Single-file mode:** paste a Drive link to process one Excel file on demand

**Setup:** in a Google Sheet, open **Extensions → Apps Script**, paste in `PropertyDescriptionETL.gs`, add the **Drive API** advanced service (**v2**, because the script uses `Drive.Files.insert`), reload, and use the **Data Processing** menu.

---

## Bulk "Copy of" Remover

Copying files in Drive leaves you with folders full of `Copy of …` names. This adds a **File Management** menu to a sheet: enter a folder ID and every file in that folder is renamed.

**Setup:** paste `RemoveCopyOf.gs` into a sheet's Apps Script editor, reload, and use the **File Management** menu.

---

## Tech

Google Apps Script · DriveApp · Drive advanced service · Script Properties · Time-based triggers · Regex
