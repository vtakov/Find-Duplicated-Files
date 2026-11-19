# Find Duplicated Files

A PowerShell script that finds identical files across two or more folders and lets you safely move the extra copies to the Recycle Bin.

> **Note:** This script was generated with the help of AI (Claude by Anthropic). Please test it on non-critical data first and use it at your own risk.

## Features

- Compares files by size **and** SHA-256 hash of their contents (not by name or date)
- Fast: same-size files ≥ 1 MB are pre-filtered by a quick hash of their first and last 64 KB
- Nothing is deleted without a CSV preview and explicit confirmation
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
```

If script execution is blocked, run it once with:

```powershell
powershell -ExecutionPolicy Bypass -File .\Find-duplicated-files.ps1
```

## How it works

1. Select at least two non-overlapping folders.
2. The script scans them and writes `All-Duplicates-<timestamp>.csv`.
3. Enter the numbers of the folders you want to clean.
4. Review `Proposed-Deletions-<timestamp>.csv`. To keep a file, delete its row and save the CSV.
5. Confirm, and the listed files are moved to the Recycle Bin. Results go to `Recycle-Results-<timestamp>.csv`.

Folder order matters: if every copy of a file is in folders being cleaned, the copy in the lowest-numbered folder is kept.

CSV reports are saved next to the script and use your Windows list separator, so they open directly in Excel.

## Disclaimer

This software is provided "as is", without warranty of any kind. Always keep a backup of important data.
