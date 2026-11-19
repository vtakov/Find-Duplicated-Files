# Find Duplicated Files

A PowerShell script that finds identical files across two or more folders and lets you safely move the extra copies to the Recycle Bin.

> **Note:** This script was generated with the help of AI (Claude by Anthropic). Please test it on non-critical data first and use it at your own risk.

## Features

- Compares files by size **and** SHA-256 hash of their contents (not by name or date)
- Fast: same-size files ≥ 1 MB are pre-filtered by a quick hash of their first and last 64 KB
- Results are shown in sortable, filterable grid windows; you pick the files to recycle there
- Nothing is deleted without a preview and explicit confirmation
- You choose which folders to clean; all other folders are never touched
- At least one copy of every duplicate is always kept
- Files go to the Recycle Bin; Windows asks before anything would be deleted permanently
- Each file is re-checked (size and hash) right before recycling
- Skips empty files, system files, symlinks/junctions and network locations; hidden files are optional

## Requirements

- Windows, 64-bit
- Windows PowerShell 5.1 or PowerShell 7+

## Usage

```powershell
# Pick folders with a dialog
.\Find-duplicated-files.ps1

# Pass folders directly
.\Find-duplicated-files.ps1 -Folders 'D:\Photos\Master', 'D:\Photos\Backup'

# Also scan hidden files and folders
.\Find-duplicated-files.ps1 -IncludeHidden

# Also save the results as CSV files
.\Find-duplicated-files.ps1 -CsvReports
```

If script execution is blocked, run it once with:

```powershell
powershell -ExecutionPolicy Bypass -File .\Find-duplicated-files.ps1
```

## How it works

1. Select at least two non-overlapping folders.
2. The script scans them and opens a window listing all duplicates. Review it and click OK to continue.
3. In the folder window, select the folders you want to clean (Ctrl+click for several) and click OK.
4. In the preview window, select the files to recycle (Ctrl+A selects all) and click OK. Files you don't select are kept.
5. Confirm in the console, and the selected files are moved to the Recycle Bin. A results window shows what was recycled or skipped.

Clicking Cancel in the folder or preview window exits without deleting anything. The grid windows can be sorted by clicking a column header and filtered with the search box or **Add criteria**.

Folder order matters: if every copy of a file is in folders being cleaned, the copy in the lowest-numbered folder is kept.

With `-CsvReports`, the script also saves `All-Duplicates-`, `Proposed-Deletions-` and `Recycle-Results-<timestamp>.csv` next to the script. They use your Windows list separator, so they open directly in Excel.

## Disclaimer

This software is provided "as is", without warranty of any kind. Always keep a backup of important data.
