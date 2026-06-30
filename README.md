# Payroll Sheet Update Tool

This tool helps you update payroll Excel workbooks automatically. It can create new weekly sheets (numbered tabs like 3, 4, 5…) and adjust formulas so they point to the correct rows in other sheets.

You do **not** need to know programming. You only need to edit a small settings file and run one script.

---

## What this tool does

Each payroll workbook has numbered sheet tabs (for example: **1**, **2**, **3**…). Sheet **1** is usually the template.

When you run the tool, it will:

1. **Create missing sheets** — If a numbered sheet in your range does not exist yet, the tool copies the template sheet and names it with the correct number.
2. **Update existing sheets** — If a numbered sheet already exists, the tool adjusts certain formulas so row references stay correct.
3. **Save the workbook** — Changes are saved back to the same Excel file(s).

New sheets are added in order from left to right (lowest number first, highest number last).

---

## What you need

- A **Windows** computer
- **Microsoft Excel** installed
- The files in this folder (the script and settings)

**Important:** Close any Excel files you plan to update before running the tool. The file must not be open in Excel.

---

## First-time setup

1. In this folder, find the file named **`.env.example`**.
2. Make a copy of it and rename the copy to **`.env`** (same folder).
3. Open **`.env`** in Notepad and change the settings to match your files (see [Settings explained](#settings-explained) below).
4. Save and close **`.env`**.

You only need to do this once, unless your files or sheet numbers change.

---

## How to run the tool

1. Close Excel if it is open.
2. Open **PowerShell** (search for “PowerShell” in the Start menu).
3. Go to this folder. For example:

   ```powershell
   cd "c:\Users\w\SSD\Code\dmpc-payroll-update"
   ```

   (Use the actual path where this folder lives on your computer.)

4. Run:

   ```powershell
   .\Update-PayrollSheets.ps1
   ```

5. Wait until it finishes. You will see messages on screen.

If something goes wrong, check the log file (see [Checking the log](#checking-the-log)).

### Easier way to run (optional)

You can also right-click **`Update-PayrollSheets.ps1`** in File Explorer and choose **Run with PowerShell**. Make sure your **`.env`** file is set up first, and that Excel is closed.

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

---

## Quick checklist before each run

- [ ] `.env` is set up with the correct file or folder
- [ ] **SHEET_RANGE** matches the sheet numbers you need
- [ ] Excel is closed
- [ ] You have a backup (recommended)

---

## Need help?

If the log shows errors you do not understand, share the log file and your `.env` settings (with paths) with whoever set up this tool for you.
