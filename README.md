# Payroll Sheet Update Tool

This folder contains PowerShell scripts that help with payroll Excel workbooks: creating/updating numbered sheets, fixing formula refs, locking sheets, and checking whether branch folders have incentives files for the current payroll week.

You do **not** need to know programming. You only need to edit a small settings file and run a script.

---

## Scripts in this folder

| Script | What it does |
|--------|----------------|
| **generate-voucher.ps1** | Create or update numbered payroll sheets from a template |
| **fix-copied-payroll-sheet.ps1** | Fix formula sheet references after copying sheets |
| **shift-payroll-ref-column.ps1** | Shift column letters in sheet formula refs |
| **lock-sheets.ps1** / **unlock-sheets.ps1** | Protect or unprotect numbered sheets |
| **check-incentives.ps1** | Check which branch folders have an incentives Excel file for the payroll week |

---

## What generate-voucher does

Each payroll workbook has numbered sheet tabs (for example: **1**, **2**, **3**…). Sheet **1** is usually the template.

When you run **generate-voucher.ps1**, it will:

1. **Create missing sheets** — If a numbered sheet in your range does not exist yet, the tool copies the template sheet and names it with the correct number.
2. **Update existing sheets** — If a numbered sheet already exists, the tool adjusts certain formulas so row references stay correct.
3. **Save the workbook** — Changes are saved back to the same Excel file(s).

New sheets are added in order from left to right (lowest number first, highest number last).

---

## Check incentives (branch folders)

**check-incentives.ps1** resolves the payroll week (Wed–Tue by default), looks in each branch folder listed in **BRANCH_PATHS**, and checks whether any Excel file name contains that week’s day range (for example `8-14` in `JULY 8-14, 2026`). It then shows a Windows popup with ✓ / ✗ per branch.

### How to run

```powershell
cd "C:\Users\itope\OneDrive\Documents\dmpc-hr-scripts"
.\check-incentives.ps1
```

Optional:

```powershell
.\check-incentives.ps1 -NoPopup          # log/console only, no dialog
.\check-incentives.ps1 -EnvPath .\other.env
```

### Settings (check-incentives)

| Setting | What it means |
|--------|----------------|
| **BRANCH_PATHS** | Branch map: one `BRANCH_PATHS=Label=FolderPath` line per branch (preferred). You can also put several pairs on one line separated by `;`. |
| **PAYROLL_TARGET_PERIOD** | `next` (default, upcoming Wed–Tue week) or `current` (period ending on/before today) |
| **BRANCH_PAYROLL_START_DAY** | Period start weekday (default `Wednesday`) |
| **BRANCH_PAYROLL_END_DAY** | Period end weekday (default `Tuesday`) |
| **REF_RUN_DATE** | Optional fixed date `yyyy-MM-dd` used instead of today when calculating the period |
| **LOG_PATH** | Log file path (e.g. `logs\check-incentives.log`) |

Matching is by **day range only** in the filename (e.g. `8-14` or `08-14`). Month formatting can vary. Only top-level `.xlsx` / `.xls` files are checked (Excel lock files starting with `~$` are skipped).

**Example — multi-line BRANCH_PATHS (easier to edit):**

```env
BRANCH_PATHS=Pandesalan=\\server\path\to\pandesalan
BRANCH_PATHS=Mindoro=\\server\path\to\mindoro
```

### Running tests

```powershell
.\tests\Check-Incentives.Tests.ps1
```

---

## What you need

- A **Windows** computer
- **Microsoft Excel** installed (for the workbook update / lock scripts; not required for **check-incentives.ps1**)
- The files in this folder (the script and settings)

**Important:** Close any Excel files you plan to update before running workbook scripts. The file must not be open in Excel.

---

## First-time setup

1. In this folder, find the file named **`.env.example`**.
2. Make a copy of it and rename the copy to **`.env`** (same folder).
3. Open **`.env`** in Notepad and change the settings to match your files (see [Settings explained](#settings-explained) below).
4. Save and close **`.env`**.

You only need to do this once, unless your files or sheet numbers change.

---

## How to run generate-voucher

1. Close Excel if it is open.
2. Open **PowerShell** (search for “PowerShell” in the Start menu).
3. Go to this folder. For example:

   ```powershell
   cd "C:\Users\itope\OneDrive\Documents\dmpc-hr-scripts"
   ```

   (Use the actual path where this folder lives on your computer.)

4. Run:

   ```powershell
   .\generate-voucher.ps1
   ```

5. Wait until it finishes. You will see messages on screen.

If something goes wrong, check the log file (see [Checking the log](#checking-the-log)).

### Easier way to run (optional)

You can also right-click a `.ps1` script in File Explorer and choose **Run with PowerShell**. Make sure your **`.env`** file is set up first, and that Excel is closed when updating workbooks.

---

## Settings explained

All settings live in the **`.env`** file. Lines starting with `#` are comments — the tool ignores them.

### Which file(s) to update

| Setting | What it means |
|--------|----------------|
| **FOLDER_PATH** | Path to a **folder** of Excel files. If you fill this in, the tool updates **every `.xlsx` file** in that folder. |
| **FILE_PATH** | Path to **one** Excel file. Used when **FOLDER_PATH** is left empty. |
| **RECURSIVE** | Only used with **FOLDER_PATH**. Set to `true` to also include Excel files inside subfolders. Use `false` to only look in the main folder. |

**Rule:** If **FOLDER_PATH** is filled in, it is used instead of **FILE_PATH**.

**Example — one file:**

```env
FOLDER_PATH=
FILE_PATH=c:\Users\w\SSD\Code\dmpc-payroll-update\test\payroll-test-source.xlsx
RECURSIVE=false
```

**Example — every file in a folder:**

```env
FOLDER_PATH=c:\Users\w\SSD\Code\dmpc-payroll-update\samples
FILE_PATH=
RECURSIVE=false
```

### Sheet settings

| Setting | What it means |
|--------|----------------|
| **SHEET_NAME** | The template sheet tab name (usually `1`). New sheets are copied from this tab. |
| **SHEET_RANGE** | Which numbered sheets to create or update. Two numbers separated by a comma: start and end. Example: `6,10` means sheets 6, 7, 8, 9, and 10. |

### Other settings

| Setting | What it means |
|--------|----------------|
| **LOG_PATH** | Where the tool writes a record of what it did. Default: `logs\payroll-update.log` |
| **TARGET_CELLS** | Which cells have formulas that need updating. Usually leave this as provided unless your IT contact tells you to change it. |

---

## Checking the log

After each run, open the log file (path shown in **LOG_PATH** in your `.env` file).

The log shows:

- Which workbook(s) were processed
- Which sheets were created or updated
- Any warnings or errors

If a run fails partway through a folder, the tool tries to continue with the remaining files and lists any failures at the end.

---

## Tips and common issues

**“File not found”**  
Check that **FILE_PATH** or **FOLDER_PATH** in `.env` points to the correct location. Paths can be full paths (starting with `c:\`) or relative to this folder.

**Excel error or script hangs**  
Make sure the workbook is **closed in Excel** before running the tool.

**Nothing seems to change**  
Check **SHEET_RANGE** — the sheet numbers you want must fall within that range. Sheets that already exist get updated; missing ones get created.

**Processing many files**  
Use **FOLDER_PATH** and set **RECURSIVE** to `true` only if you need files in subfolders too.

**Backup your files**  
Before running on important payroll files, make a copy of the Excel file(s) or folder. The tool saves changes directly to the originals.

**Incentives check shows ✗ for every branch**  
Confirm **BRANCH_PATHS** folders exist and that an Excel filename in each folder contains the expected day range (for example `15-21`). Check **PAYROLL_TARGET_PERIOD** (`next` vs `current`) and optional **REF_RUN_DATE**.

---

## Quick checklist before each run

- [ ] `.env` is set up with the correct file or folder
- [ ] **SHEET_RANGE** matches the sheet numbers you need (for voucher/lock scripts)
- [ ] Excel is closed (when updating workbooks)
- [ ] You have a backup (recommended)

---

## Need help?

If the log shows errors you do not understand, share the log file and your `.env` settings (with paths) with whoever set up this tool for you.
