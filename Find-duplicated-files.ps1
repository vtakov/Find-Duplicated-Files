# Find identical files across several folders, then selectively recycle copies.
# Works with Windows PowerShell 5.1 and PowerShell 7 on Windows (64-bit).
# Files are compared by size AND SHA-256 of their actual contents.
# No files are deleted without a preview report and an explicit confirmation.
#
# Usage:
#   .\Find-duplicated-files.ps1                           # pick folders with a dialog
#   .\Find-duplicated-files.ps1 -Folders 'D:\A','D:\B'    # skip the dialog
#   .\Find-duplicated-files.ps1 -IncludeHidden            # also scan hidden files/folders
#   pwsh .\Find-duplicated-files.ps1                      # run with PowerShell 7
#
# Always skipped: empty (0-byte) files, files/folders with the System attribute,
# symlinks/junctions, and network locations.
# CSV reports are written to the folder that contains this script. They use
# your Windows list separator (e.g. ';' on German systems) so Excel opens them
# directly. You can delete rows from Proposed-Deletions-*.csv to keep files.
#
# Folder order matters: when a copy must be kept from folders being cleaned,
# the copy in the lowest-numbered folder is kept.

param(
    # OPTIONAL: paths to scan. Leave empty to use the graphical folder picker.
    # Example: @('D:\Photos\Master', 'D:\Photos\Backup1', 'E:\OldPhotos')
    [string[]]$Folders = @(),

    # Include hidden files and folders. If not given, you are asked (default: No).
    [switch]$IncludeHidden
)

$ErrorActionPreference = 'Stop'

# Files at least this big are first compared by a quick hash of their first
# and last 64 KB; only files that still match are fully hashed.
$quickHashMinSize = 1MB
$quickHashChunk = 64KB

# UTF-8 with BOM so Excel shows accented characters correctly.
# ('UTF8' means "with BOM" in 5.1 but "without BOM" in PowerShell 7.)
$csvEncoding = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }

if (-not [Environment]::Is64BitProcess) {
    throw 'Please run this script in 64-bit PowerShell.'
}

# Recycle Bin helper (Windows shell API). Unlike a silent delete, it shows a
# warning before any file would be PERMANENTLY deleted instead of recycled
# (for example, the file is too big for the Recycle Bin, or the Recycle Bin is
# turned off for that drive). Answering "No" skips the file.
if (-not ('DupFinder.Recycler' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

namespace DupFinder {
    public static class Recycler {
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct SHFILEOPSTRUCT {
            public IntPtr hwnd;
            public uint wFunc;
            [MarshalAs(UnmanagedType.LPWStr)] public string pFrom;
            [MarshalAs(UnmanagedType.LPWStr)] public string pTo;
            public ushort fFlags;
            [MarshalAs(UnmanagedType.Bool)] public bool fAnyOperationsAborted;
            public IntPtr hNameMappings;
            [MarshalAs(UnmanagedType.LPWStr)] public string lpszProgressTitle;
        }

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        private static extern int SHFileOperation(ref SHFILEOPSTRUCT op);

        private const uint FO_DELETE = 3;
        private const ushort FOF_SILENT = 0x0004;
        private const ushort FOF_NOCONFIRMATION = 0x0010;
        private const ushort FOF_ALLOWUNDO = 0x0040;
        private const ushort FOF_WANTNUKEWARNING = 0x4000;

        // Returns 0 on success, -1 if the operation was cancelled,
        // otherwise the shell error code.
        public static int SendToRecycleBin(string path) {
            SHFILEOPSTRUCT op = new SHFILEOPSTRUCT();
            op.hwnd = IntPtr.Zero;
            op.wFunc = FO_DELETE;
            op.pFrom = path + "\0";   // list must end with a double null
            op.pTo = null;
            op.fFlags = (ushort)(FOF_SILENT | FOF_NOCONFIRMATION | FOF_ALLOWUNDO | FOF_WANTNUKEWARNING);
            int rc = SHFileOperation(ref op);
            if (rc == 0 && op.fAnyOperationsAborted) { return -1; }
            return rc;
        }
    }
}
'@
}

function Normalize-Folder([string]$path) {
    $item = Get-Item -LiteralPath $path -Force -ErrorAction Stop
    if (-not $item.PSIsContainer) { throw "Not a folder: $path" }
    $full = [System.IO.Path]::GetFullPath($item.FullName)
    $root = [System.IO.Path]::GetPathRoot($full)
    if ($full -ne $root) { $full = $full.TrimEnd([char[]]@('\', '/')) }
    return $full
}

function Is-Same-Or-Child([string]$path, [string]$folder) {
    if ($path.Equals($folder, [StringComparison]::OrdinalIgnoreCase)) { return $true }
    $prefix = if ($folder.EndsWith('\')) { $folder } else { "$folder\" }
    return $path.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)
}

function Test-NetworkPath([string]$path) {
    if ($path.StartsWith('\\')) { return $true }   # UNC path, e.g. \\server\share
    try {
        $drive = [IO.DriveInfo]::new([IO.Path]::GetPathRoot($path))
        return $drive.DriveType -eq [IO.DriveType]::Network
    }
    catch { return $false }
}

# Lists non-empty files under $root. Skips System items and reparse points
# (symlinks/junctions) always, and Hidden items unless $includeHidden is set.
# Skipped folders are not entered. Unreadable folders are added to $scanErrors.
function Get-ScanFiles([string]$root, [bool]$includeHidden) {
    $skip = [IO.FileAttributes]::ReparsePoint -bor [IO.FileAttributes]::System
    if (-not $includeHidden) { $skip = $skip -bor [IO.FileAttributes]::Hidden }
    $skipMask = [int]$skip

    $pending = [Collections.Generic.Stack[IO.DirectoryInfo]]::new()
    $pending.Push([IO.DirectoryInfo]::new($root))
    while ($pending.Count -gt 0) {
        $dir = $pending.Pop()
        try { $entries = $dir.GetFileSystemInfos() }
        catch {
            $scanErrors.Add("$($dir.FullName): $($_.Exception.Message)")
            continue
        }
        foreach ($entry in $entries) {
            if (([int]$entry.Attributes -band $skipMask) -ne 0) { continue }
            if ($entry -is [IO.DirectoryInfo]) { $pending.Push($entry) }
            elseif ($entry.Length -gt 0) { $entry }
        }
    }
}

# Throws if the file is gone, changed type/size, or (when $expectedHash is
# given) no longer has the content recorded during scanning.
function Assert-Unchanged([string]$path, [long]$size, [string]$expectedHash) {
    $item = Get-Item -LiteralPath $path -Force -ErrorAction Stop
    if ($item.PSIsContainer -or
        ([int]$item.Attributes -band [int][IO.FileAttributes]::ReparsePoint) -ne 0 -or
        $item.Length -ne $size) {
        throw "File type or size changed since scanning: $path"
    }
    if ($expectedHash) {
        $hash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        if ($hash -ne $expectedHash) { throw "Content changed since scanning: $path" }
    }
}

# SHA-256 of the first and last $chunk bytes. File must be at least $chunk long.
function Get-PartialHash([string]$path, [int]$chunk) {
    $stream = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $buffer = New-Object byte[] (2 * $chunk)
        $read = 0
        foreach ($start in @(0, ($stream.Length - $chunk))) {
            [void]$stream.Seek($start, [IO.SeekOrigin]::Begin)
            $got = 0
            while ($got -lt $chunk) {
                $n = $stream.Read($buffer, $read + $got, $chunk - $got)
                if ($n -le 0) { break }
                $got += $n
            }
            $read += $got
        }
        return [BitConverter]::ToString($sha.ComputeHash($buffer, 0, $read))
    }
    finally {
        $sha.Dispose()
        $stream.Dispose()
    }
}

function Show-Progress([string]$activity, [int]$done, [int]$total) {
    if ($progressTimer.ElapsedMilliseconds -lt 250) { return }
    Write-Progress -Activity $activity -Status "$($done + 1) of $total" `
        -PercentComplete ([int](100 * $done / $total))
    $progressTimer.Restart()
}

function Save-Csv($data, [string]$path) {
    $options = @{ LiteralPath = $path; Encoding = $csvEncoding; UseCulture = $true }
    # Needed in 5.1; obsolete (and the default) in PowerShell 7.
    if ($PSVersionTable.PSVersion.Major -lt 6) { $options['NoTypeInformation'] = $true }
    $data | Export-Csv @options
}

# ---------------------------------------------------------------------------
# 1. Choose folders
# ---------------------------------------------------------------------------
if ($Folders.Count -eq 0) {
    Add-Type -AssemblyName System.Windows.Forms
    Write-Host 'Select the folders to scan, one at a time.' -ForegroundColor Cyan
    $picked = [Collections.Generic.List[string]]::new()
    $lastPicked = $null
    do {
        $picker = New-Object System.Windows.Forms.FolderBrowserDialog
        $picker.Description = 'Select a folder to include in duplicate scanning'
        $picker.ShowNewFolderButton = $false
        # Pre-select the last folder: the dialog then opens in its parent with
        # that folder highlighted, so its sibling folders are visible.
        if ($lastPicked) { $picker.SelectedPath = $lastPicked }
        try {
            $result = $picker.ShowDialog()
            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                $full = Normalize-Folder $picker.SelectedPath
                $clash = $picked | Where-Object {
                    (Is-Same-Or-Child $full $_) -or (Is-Same-Or-Child $_ $full)
                } | Select-Object -First 1
                if ($clash) {
                    Write-Warning "Not added: '$full' is the same as, inside, or contains '$clash'."
                }
                else {
                    $picked.Add($full)
                    $Folders += $full
                    Write-Host "Added: $full"
                }
                $lastPicked = $full
            }
        }
        finally { $picker.Dispose() }
        $answer = Read-Host 'Add another folder? [Y/n]'
    } while ($answer -notmatch '^\s*[Nn]')   # Enter (empty) means Yes
}

$unique = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$normalized = @(
    foreach ($path in $Folders) {
        $full = Normalize-Folder $path
        if ($unique.Add($full)) { $full }
    }
)

if ($normalized.Count -lt 2) { throw 'Select at least two distinct folders.' }
foreach ($path in $normalized) {
    if (Test-NetworkPath $path) { throw "Network locations are not supported: $path" }
}
for ($i = 0; $i -lt $normalized.Count; $i++) {
    for ($j = $i + 1; $j -lt $normalized.Count; $j++) {
        if ((Is-Same-Or-Child $normalized[$i] $normalized[$j]) -or
            (Is-Same-Or-Child $normalized[$j] $normalized[$i])) {
            throw "Folders must not overlap: '$($normalized[$i])' and '$($normalized[$j])'"
        }
    }
}

$includeHiddenFiles = $IncludeHidden.IsPresent
if (-not $PSBoundParameters.ContainsKey('IncludeHidden')) {
    $answer = Read-Host 'Include hidden files and folders? [y/N]'
    $includeHiddenFiles = $answer -match '^[Yy]'
}

Write-Host "`nFolders to scan:" -ForegroundColor Cyan
for ($i = 0; $i -lt $normalized.Count; $i++) {
    Write-Host (" {0,2}. {1}" -f ($i + 1), $normalized[$i])
}
$hiddenText = if ($includeHiddenFiles) { 'included' } else { 'skipped' }
Write-Host "Hidden files/folders: $hiddenText. System files, empty files and symlinks: skipped."

# CSV reports go next to this script (or the current folder if run without a file).
$reportDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).ProviderPath }
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'

# ---------------------------------------------------------------------------
# 2. Inventory files
# ---------------------------------------------------------------------------
$files = [Collections.Generic.List[object]]::new()
$scanErrors = [Collections.Generic.List[string]]::new()
for ($i = 0; $i -lt $normalized.Count; $i++) {
    Write-Host "Scanning folder $($i + 1) of $($normalized.Count)..."
    foreach ($f in (Get-ScanFiles $normalized[$i] $includeHiddenFiles)) {
        $files.Add([pscustomobject]@{
            FolderNumber = $i + 1
            Path = $f.FullName
            Size = [long]$f.Length
        })
    }
}
if ($scanErrors.Count -gt 0) {
    Write-Warning "$($scanErrors.Count) folders could not be read (for example, access denied). Files in them were not scanned."
    $scanErrors | Select-Object -First 5 | ForEach-Object { Write-Warning "  $_" }
}
Write-Host "Found $($files.Count) non-empty files."

# ---------------------------------------------------------------------------
# 3. Find duplicates in three passes, from cheapest to most thorough:
#    a) same size, b) same quick hash of first+last 64 KB (files >= 1 MB),
#    c) same full SHA-256 hash.
# ---------------------------------------------------------------------------
$sizeCounts = @{}
foreach ($f in $files) {
    $key = [string]$f.Size
    if ($sizeCounts.ContainsKey($key)) { $sizeCounts[$key]++ }
    else { $sizeCounts[$key] = 1 }
}
$sameSize = @($files | Where-Object { $sizeCounts[[string]$_.Size] -ge 2 })

$progressTimer = [Diagnostics.Stopwatch]::StartNew()
$quick = [Collections.Generic.List[object]]::new()
$quickCounts = @{}
for ($n = 0; $n -lt $sameSize.Count; $n++) {
    $f = $sameSize[$n]
    Show-Progress 'Quick check of same-size files' $n $sameSize.Count
    $key = [string]$f.Size
    if ($f.Size -ge $quickHashMinSize) {
        try { $key += ':' + (Get-PartialHash $f.Path $quickHashChunk) }
        catch { Write-Warning "Could not read '$($f.Path)': $_"; continue }
    }
    $quick.Add([pscustomobject]@{ File = $f; Key = $key })
    if ($quickCounts.ContainsKey($key)) { $quickCounts[$key]++ }
    else { $quickCounts[$key] = 1 }
}
Write-Progress -Activity 'Quick check of same-size files' -Completed

$toHash = @($quick | Where-Object { $quickCounts[$_.Key] -ge 2 } | ForEach-Object { $_.File })
$hashed = [Collections.Generic.List[object]]::new()
for ($n = 0; $n -lt $toHash.Count; $n++) {
    $f = $toHash[$n]
    Show-Progress 'Hashing possible duplicates' $n $toHash.Count
    try {
        $digest = (Get-FileHash -LiteralPath $f.Path -Algorithm SHA256).Hash
        $hashed.Add([pscustomobject]@{
            FolderNumber = $f.FolderNumber
            Path = $f.Path
            Size = $f.Size
            SHA256 = $digest
            GroupKey = "$($f.Size):$digest"
        })
    }
    catch { Write-Warning "Could not hash '$($f.Path)': $_" }
}
Write-Progress -Activity 'Hashing possible duplicates' -Completed

$groups = @($hashed | Group-Object -Property GroupKey | Where-Object { $_.Count -ge 2 })
if ($groups.Count -eq 0) {
    Write-Host 'No duplicates found.' -ForegroundColor Green
    return
}

# ---------------------------------------------------------------------------
# 4. Full report: every file that has at least one identical copy
# ---------------------------------------------------------------------------
$allReport = Join-Path $reportDir "All-Duplicates-$stamp.csv"
$rows = @(
    for ($i = 0; $i -lt $groups.Count; $i++) {
        foreach ($file in $groups[$i].Group) {
            [pscustomobject]@{
                DuplicateGroup = $i + 1
                FolderNumber = $file.FolderNumber
                FilePath = $file.Path
                SizeBytes = $file.Size
                SHA256 = $file.SHA256
            }
        }
    }
)
Save-Csv $rows $allReport
Write-Host "`nFound $($groups.Count) duplicate groups ($($rows.Count) files)." -ForegroundColor Green
Write-Host "Full report: $allReport"

# ---------------------------------------------------------------------------
# 5. Choose which folders to clean
# ---------------------------------------------------------------------------
Write-Host "`nWhich folders should be CLEANED? (All other folders stay untouched.)" -ForegroundColor Cyan
Write-Host 'If every copy is in folders being cleaned, the copy in the lowest-numbered folder is kept.'
for ($i = 0; $i -lt $normalized.Count; $i++) {
    Write-Host (" {0,2}. {1}" -f ($i + 1), $normalized[$i])
}
$selection = Read-Host 'Enter folder numbers separated by commas (e.g. 2,4), or Enter to exit'
if ([string]::IsNullOrWhiteSpace($selection)) { return }
$tokens = @($selection -split '[,;\s]+' | Where-Object { $_ })
$selected = [Collections.Generic.HashSet[int]]::new()
foreach ($token in $tokens) {
    $number = 0
    if (-not [int]::TryParse($token, [ref]$number) -or
        $number -lt 1 -or $number -gt $normalized.Count) {
        throw "Invalid folder number: $token"
    }
    [void]$selected.Add($number)
}

# ---------------------------------------------------------------------------
# 6. Decide what to keep. All copies in unselected folders are kept. If a
#    group exists ONLY in selected folders, one copy is preserved so nothing
#    is lost completely: the one in the lowest-numbered folder (then by path).
# ---------------------------------------------------------------------------
$candidates = [Collections.Generic.List[object]]::new()
for ($i = 0; $i -lt $groups.Count; $i++) {
    $members = @($groups[$i].Group | Sort-Object -Property FolderNumber, Path)
    $outside = @($members | Where-Object { -not $selected.Contains($_.FolderNumber) })
    $keeper = if ($outside.Count -gt 0) { $outside[0] } else { $members[0] }
    foreach ($f in $members) {
        if ($selected.Contains($f.FolderNumber) -and $f.Path -ne $keeper.Path) {
            $candidates.Add([pscustomobject]@{
                DuplicateGroup = $i + 1
                FolderNumber = $f.FolderNumber
                FileToRecycle = $f.Path
                KeptMatchingCopy = $keeper.Path
                SizeBytes = $f.Size
                SHA256 = $f.SHA256
            })
        }
    }
}
if ($candidates.Count -eq 0) {
    Write-Host 'No files need deleting from the selected folders.' -ForegroundColor Green
    return
}

$deleteReport = Join-Path $reportDir "Proposed-Deletions-$stamp.csv"
Save-Csv $candidates $deleteReport
$sum = ($candidates | Measure-Object -Property SizeBytes -Sum).Sum
Write-Host "`nPREVIEW: $($candidates.Count) files would be recycled ($([math]::Round($sum / 1GB, 2)) GB)." -ForegroundColor Yellow
Write-Host "Review this CSV BEFORE continuing: $deleteReport"
Write-Host 'To KEEP a file, delete its row from the CSV and save it (same name, CSV format).'
Write-Host 'Rows can only be removed; anything not in the original preview is ignored.'
Write-Host 'At least one matching copy of each duplicate group will be retained.'
Write-Host 'Unique files are never selected. Unselected folders are never changed.'
Write-Host 'If a file cannot go to the Recycle Bin, Windows asks before deleting it permanently.'

# Re-read the (possibly edited) CSV. Only rows that match the original preview
# are used, and all details come from the scan, never from the edited file.
$candidateByPath = @{}   # PowerShell hashtables ignore case
foreach ($c in $candidates) { $candidateByPath[$c.FileToRecycle] = $c }

while ($true) {
    $confirmation = Read-Host 'Move the files listed in the CSV to the Recycle Bin? [y/N]'
    if ($confirmation -notmatch '^\s*[Yy]') {   # Enter (empty) means No
        Write-Host 'Cancelled. No files deleted.'
        return
    }
    try {
        $editedRows = @(Import-Csv -LiteralPath $deleteReport -UseCulture -ErrorAction Stop)
        break
    }
    catch {
        Write-Warning "Could not read the CSV: $($_.Exception.Message)"
        Write-Warning 'If it is open in Excel, close it and try again.'
    }
}

$approved = [Collections.Generic.List[object]]::new()
$approvedPaths = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($row in $editedRows) {
    $path = $row.FileToRecycle
    if ([string]::IsNullOrWhiteSpace($path) -or -not $candidateByPath.ContainsKey($path)) {
        Write-Warning "Ignored a CSV row that is not in the original preview: '$path'"
        continue
    }
    if ($approvedPaths.Add($path)) { $approved.Add($candidateByPath[$path]) }
}
$keptByUser = $candidates.Count - $approved.Count
if ($keptByUser -gt 0) {
    Write-Host "$keptByUser files were removed from the list and will be kept." -ForegroundColor Cyan
}
if ($approved.Count -eq 0) {
    Write-Host 'Nothing left to recycle. No files deleted.'
    return
}
Write-Host "Recycling $($approved.Count) files..."

# ---------------------------------------------------------------------------
# 7. Recycle. Before each file, re-check that the kept copy and the file to
#    remove still exist with the size and content recorded during scanning.
# ---------------------------------------------------------------------------
$recycled = 0
$skipped = 0
$keeperHashChecked = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$results = [Collections.Generic.List[object]]::new()
foreach ($item in $approved) {
    $status = 'Recycled'
    $reason = ''
    try {
        # Kept copy: existence and size every time; full hash only once.
        $keeperHash = if ($keeperHashChecked.Contains($item.KeptMatchingCopy)) { $null } else { $item.SHA256 }
        Assert-Unchanged $item.KeptMatchingCopy $item.SizeBytes $keeperHash
        [void]$keeperHashChecked.Add($item.KeptMatchingCopy)

        Assert-Unchanged $item.FileToRecycle $item.SizeBytes $item.SHA256

        $rc = [DupFinder.Recycler]::SendToRecycleBin($item.FileToRecycle)
        if ($rc -eq -1) { throw 'Cancelled (it would have been deleted permanently).' }
        if ($rc -ne 0) { throw "Recycle Bin operation failed (code $rc)." }
        if (Test-Path -LiteralPath $item.FileToRecycle) { throw 'File still exists after recycling.' }
        $recycled++
    }
    catch {
        $skipped++
        $status = 'Skipped'
        $reason = $_.Exception.Message
        Write-Warning "Skipped '$($item.FileToRecycle)': $reason"
    }
    $results.Add([pscustomobject]@{
        Status = $status
        Reason = $reason
        FileToRecycle = $item.FileToRecycle
        KeptMatchingCopy = $item.KeptMatchingCopy
        SizeBytes = $item.SizeBytes
    })
}

$resultReport = Join-Path $reportDir "Recycle-Results-$stamp.csv"
Save-Csv $results $resultReport
Write-Host "`nFinished: $recycled files recycled; $skipped skipped." -ForegroundColor Green
Write-Host "Results report: $resultReport"
