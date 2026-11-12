# RUNME.ps1
# First-run setup for FAIS (was Create-FAISShortcut.ps1):
#   1. Copy the bundled ImportExcel folder into the current-user module path
#      so FAIS.ps1 can run Import-Module ImportExcel.
#   2. Leave MasterData.xlsx in this folder (next to RUNME.ps1 / FAIS.ps1).
#      FAIS.ps1 reads it from here. A previous RUNME that moved it one folder
#      up will move it back.
#   3. Build Windows shortcuts (hidden PowerShell, no Notepad):
#        FAIS.lnk             -> powershell.exe -WindowStyle Hidden -File FAIS.ps1
#        MasterData.lnk       -> MasterData.xlsx in this folder
#        OutputSchedule.lnk   -> OutputSchedule.xlsx in this folder
#        RUNME.lnk / RUNME.cmd -> powershell.exe -File RUNME.ps1 (this folder)
#   4. Paste student names (Last, First). Writes them onto:
#        MasterData Students     (SPName / IsCAP / IsInternational, current columns)
#        String Tracker          (StudentName column — OutputSchedule.xlsx first,
#                                 then MasterData.xlsx if that is where the sheet is)
#      Asks if anyone is CAP or International, and creates a blank syllabus tab
#      per new student in OutputSchedule.xlsx (same layout as the tabs already
#      there). Existing String Tracker names are not duplicated.
#
# Put FAISlogo.ico in:
#   <this folder>\App Icons\FAISlogo.ico
# which is FAIS\App Icons when this folder is the FAIS application folder.
#
# Put the ImportExcel module folder in:
#   <this folder>\ImportExcel
# so this script can copy it to:
#   <Documents>\WindowsPowerShell\Modules\ImportExcel
# That is the CurrentUser module path for Windows PowerShell 5.1 — the same
# place Install-Module ImportExcel -Scope CurrentUser would put it.
#
# Documents (and therefore the module path) includes a DOD ID folder. On this
# machine that folder is 1595161645A. On another computer it is the profile
# folder two levels up from this script:
#   C:\Users\<DoDID>\Desktop\FAIS9.1  ->  C:\Users\<DoDID>
# That folder name is used as-is (do not append A; the trailing A is already
# there when the profile requires it). RUNME 1:1 swaps that one folder. The
# rest of the path is unchanged. No DOD ID popup.
#
# At first run RUNME opens a folder picker starting at C:\Users\<DoDID>\ and
# stores the chosen path in $global:teamsLocation. A global variable ends when
# this PowerShell window closes, so the path is also saved to
#   <this folder>\TeamsLocation.txt   (one line: the folder)
# FAIS.ps1 reads that file at startup and writes ATIS.xlsx into that folder.
# Cancel keeps the folder saved last time. No file (or the folder cannot be
# found when FAIS runs): ATIS.xlsx goes to the Desktop, as before.
#
# MasterData.xlsx stays next to this script. FAIS.ps1 looks there first:
#   Join-Path $PSScriptRoot "MasterData.xlsx"
# then falls back to one folder up if a copy was left there by an older RUNME.
#
# OutputSchedule.xlsx stays next to this script. A Desktop shortcut
# (OutputSchedule.lnk) opens that workbook — the file is not copied to the
# Desktop. Older RUNME copies on the Desktop are moved back here if this
# folder does not already have one (never overwrite a live local workbook).
#
# Run this script from the FAIS folder (the same folder as FAIS.ps1).
#
# Windows binds the Open verb for .ps1/.psm1 to notepad.exe. Double-clicking
# a script NEVER runs it — it opens Notepad. That is by design (and STIG).
# Do not change the file association.
#
# How to run this file:
#   Right-click RUNME.ps1 -> Run with PowerShell
#   or double-click RUNME.cmd / RUNME.lnk (this script writes both)
#   or:
#     powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File .\RUNME.ps1
#
# This script never Start-Process / Invoke-Item / ShellExecute a .ps1 or .psm1.
# It never launches FAIS.ps1 (only writes FAIS.lnk). It never imports the
# FAIS Modules\*.psm1 files (Add Students is fully inlined below).
#
# Optional:
#   .\RUNME.ps1 -StartMenu
#   .\RUNME.ps1 -Path 'D:\Shortcuts' -Name 'FAIS'
#   .\RUNME.ps1 -SkipAddStudents

#Requires -Version 5.1
[CmdletBinding()]
param (
    # Folder the .lnk file is written to. Default: the current user's Desktop.
    [string]$Path,

    # Shortcut file name without extension for the FAIS launcher. Default: FAIS
    [string]$Name = 'FAIS',

    # Also write a copy of the FAIS launcher to the current user's Start Menu Programs folder.
    [switch]$StartMenu,

    # Setup only: skip the paste-names / CAP / International dialogs.
    [switch]$SkipAddStudents
)

$ErrorActionPreference = 'Stop'

if ($env:OS -ne 'Windows_NT') {
    Write-Host 'This script creates a Windows .lnk shortcut and must be run on Windows.' -ForegroundColor Red
    return
}

# Hide the console so the terminal is not sitting on screen. WinForms dialogs
# still show. Never Start-Process a .ps1/.psm1 — that opens Notepad.
function Hide-RunmeConsole {
    try {
        if (-not ('FAIS.RunmeNative' -as [type])) {
            Add-Type -Namespace FAIS -Name RunmeNative -MemberDefinition @"
[System.Runtime.InteropServices.DllImport("kernel32.dll")]
public static extern System.IntPtr GetConsoleWindow();
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool ShowWindow(System.IntPtr hWnd, int nCmdShow);
"@
        }
        $hwnd = [FAIS.RunmeNative]::GetConsoleWindow()
        if ($hwnd -ne [IntPtr]::Zero) {
            # 0 = SW_HIDE, 6 = SW_MINIMIZE. Hide so it is not visible.
            [void][FAIS.RunmeNative]::ShowWindow($hwnd, 0)
        }
        try { [Console]::Title = 'FAIS' } catch { }
    }
    catch { }
}

function Get-RunmePowerShellExe {
    $exe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (Test-Path -LiteralPath $exe) { return $exe }
    return (Get-Process -Id $PID).Path
}

function Test-RunmeIsScriptPath {
    param ([string]$FilePath)
    if ([string]::IsNullOrWhiteSpace($FilePath)) { return $false }
    return ($FilePath -match '\.(ps1|psm1|psd1)$')
}

function Show-RunmeError {
    param ([string]$Message, [string]$Title = 'RUNME')
    Write-Host $Message -ForegroundColor Red
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
        [void][System.Windows.Forms.MessageBox]::Show(
            $Message, $Title,
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    catch { }
}

# DOD ID folder currently baked into the ImportExcel destination path.
# 1:1 swap this folder name only; everything else in the path stays the same.
$script:RunmeCurrentDodFolder = '1595161645A'

# Process-wide Teams folder the user picks at first run (C:\Users\<DoDID>\...).
$global:teamsLocation = $null

# Script runs from C:\Users\<DoDID>\Desktop\FAIS9.1. Two levels up is
# C:\Users\<DoDID>. Use that profile path as-is (do not append A).
function Get-RunmeUserProfileTwoLevelsUp {
    $here = $PSScriptRoot
    if ([string]::IsNullOrWhiteSpace($here)) {
        if (-not [string]::IsNullOrWhiteSpace($PSCommandPath)) {
            $here = Split-Path -Parent $PSCommandPath
        }
        else {
            $here = (Get-Location).ProviderPath
        }
    }

    $twoUp = $here
    for ($i = 0; $i -lt 2; $i++) {
        $parent = Split-Path -Parent $twoUp
        if ([string]::IsNullOrWhiteSpace($parent)) { break }
        $twoUp = $parent
    }

    if ([string]::IsNullOrWhiteSpace($twoUp) -or -not (Test-Path -LiteralPath $twoUp)) {
        $fallback = $env:USERPROFILE
        if (-not [string]::IsNullOrWhiteSpace($fallback) -and (Test-Path -LiteralPath $fallback)) {
            return $fallback
        }
    }
    return $twoUp
}

function Get-RunmeDodFolderNameFromProfile {
    param ([string]$UserProfilePath)

    if ([string]::IsNullOrWhiteSpace($UserProfilePath)) { return $null }
    return (Split-Path -Leaf $UserProfilePath.TrimEnd('\', '/'))
}

# Replace one path segment. Drive, separators, and every other folder stay put.
function Convert-RunmePathFolderName {
    param (
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [string]$OldFolder,
        [Parameter(Mandatory)] [string]$NewFolder
    )

    if ([string]::IsNullOrWhiteSpace($Path)) { return $Path }
    if ([string]::Equals($OldFolder, $NewFolder, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $Path
    }

    $root = [System.IO.Path]::GetPathRoot($Path)
    if ([string]::IsNullOrEmpty($root)) { $root = '' }
    $relative = $Path
    if ($root.Length -gt 0 -and $Path.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) {
        $relative = $Path.Substring($root.Length)
    }

    $parts = New-Object System.Collections.Generic.List[string]
    $exactHits = 0
    foreach ($p in ($relative -split '[\\/]')) {
        if ($p -eq '') { continue }
        if ([string]::Equals($p, $OldFolder, [System.StringComparison]::OrdinalIgnoreCase)) {
            $parts.Add($NewFolder)
            $exactHits++
        }
        else {
            $parts.Add($p)
        }
    }

    # Other computer: the DOD ID folder is already a different <digits> or
    # <digits>A. Still 1:1 swap that one folder; nothing else in the path changes.
    if ($exactHits -eq 0 -and $NewFolder -match '^\d{10}A?$') {
        for ($i = 0; $i -lt $parts.Count; $i++) {
            if ($parts[$i] -match '^\d{10}A?$') {
                $parts[$i] = $NewFolder
                break
            }
        }
    }

    $joined = ($parts -join '\')
    if ($root -eq '') { return $joined }
    if ($joined -eq '') { return ($root.TrimEnd('\') + '\') }
    return (Join-Path $root.TrimEnd('\') $joined)
}

# Folder picker starting at C:\Users\<DoDID>\. The chosen path is copied to
# $global:teamsLocation. Returns that path, or $null if the user cancels.
# -CurrentLocation (the folder saved last time) is shown in the picker so
# the user knows Cancel keeps it.
function Show-RunmeTeamsFolderDialog {
    param ([string]$InitialDirectory, [string]$CurrentLocation)

    $selected = $null
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
        [System.Windows.Forms.Application]::EnableVisualStyles()
    }
    catch {
        $global:teamsLocation = $null
        return $null
    }

    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = 'Select the Teams folder FAIS saves ATIS.xlsx to'
    if (-not [string]::IsNullOrWhiteSpace($CurrentLocation)) {
        $dialog.Description += "`nNow: $CurrentLocation (Cancel keeps it)"
    }
    $dialog.ShowNewFolderButton = $true
    if (-not [string]::IsNullOrWhiteSpace($InitialDirectory) -and (Test-Path -LiteralPath $InitialDirectory)) {
        $dialog.SelectedPath = $InitialDirectory
    }

    # Hidden owner so the picker is TopMost while the RUNME console is hidden.
    $owner = New-Object System.Windows.Forms.Form
    $owner.Text = 'FAIS'
    $owner.TopMost = $true
    $owner.ShowInTaskbar = $false
    $owner.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
    $owner.StartPosition = 'CenterScreen'
    $owner.Size = New-Object System.Drawing.Size(1, 1)
    $owner.Opacity = 0

    try {
        [void]$owner.Show()
        $result = $dialog.ShowDialog($owner)
        if ($result -eq [System.Windows.Forms.DialogResult]::OK -and -not [string]::IsNullOrWhiteSpace($dialog.SelectedPath)) {
            $selected = $dialog.SelectedPath
        }
    }
    finally {
        try { $owner.Close() } catch { }
        $owner.Dispose()
        $dialog.Dispose()
    }

    $global:teamsLocation = $selected
    return $selected
}

# TeamsLocation.txt (next to RUNME.ps1 / FAIS.ps1) carries $global:teamsLocation
# over to FAIS.ps1, which runs in its own PowerShell process later. One line:
# the folder. FAIS.ps1 reads it the same way (first non-blank line, quotes
# stripped, so a hand-edited "Copy as path" line works too).
function Read-RunmeTeamsLocation {
    param ([string]$FilePath)

    if ([string]::IsNullOrWhiteSpace($FilePath) -or -not (Test-Path -LiteralPath $FilePath -PathType Leaf)) { return $null }
    try {
        foreach ($line in [System.IO.File]::ReadAllLines($FilePath)) {
            $folder = $line.Trim().Trim('"').Trim()
            if ($folder -ne '') { return $folder }
        }
    }
    catch { }
    return $null
}

function Save-RunmeTeamsLocation {
    param (
        [Parameter(Mandatory)] [string]$FilePath,
        [Parameter(Mandatory)] [string]$Location
    )

    # UTF-8 without a BOM, so any folder name survives the round trip.
    [System.IO.File]::WriteAllText($FilePath, $Location.Trim(), (New-Object System.Text.UTF8Encoding($false)))
}

# CreateProcess (UseShellExecute = $false). File associations are never
# consulted, so a .ps1/.psm1 cannot open in Notepad. FileName must be an
# .exe — never a script. ArgumentList is ONE string (PS 5.1 array quoting
# is broken and can drop -File, after which the path is ShellExecute'd).
function Start-RunmePowerShellFile {
    param (
        [Parameter(Mandatory)] [string]$ScriptPath,
        [string]$ArgumentLine = '',
        [string]$WorkingDirectory
    )

    if (-not (Test-Path -LiteralPath $ScriptPath)) {
        throw "Script not found: $ScriptPath"
    }
    $exe = Get-RunmePowerShellExe
    if (Test-RunmeIsScriptPath $exe) {
        throw "Refusing to launch a script path as the host executable: $exe"
    }

    $argLine = "-NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`""
    if (-not [string]::IsNullOrWhiteSpace($ArgumentLine)) {
        $argLine = "$argLine $ArgumentLine"
    }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName         = $exe
    $psi.Arguments        = $argLine
    $psi.UseShellExecute  = $false
    $psi.CreateNoWindow   = $true
    $psi.WindowStyle      = [System.Diagnostics.ProcessWindowStyle]::Hidden
    if (-not [string]::IsNullOrWhiteSpace($WorkingDirectory) -and (Test-Path -LiteralPath $WorkingDirectory)) {
        $psi.WorkingDirectory = $WorkingDirectory
    }
    [void][System.Diagnostics.Process]::Start($psi)
}

function Get-RunmeStaRelaunchArgs {
    # Read script-scope param() values. $PSBoundParameters inside this
    # function is empty (it is this function's own bound set).
    $extra = ''
    if (-not [string]::IsNullOrWhiteSpace($Path)) {
        $extra += " -Path `"$Path`""
    }
    if (-not [string]::IsNullOrWhiteSpace($Name) -and $Name -ne 'FAIS') {
        $extra += " -Name `"$Name`""
    }
    if ($StartMenu)       { $extra += ' -StartMenu' }
    if ($SkipAddStudents) { $extra += ' -SkipAddStudents' }
    return $extra.Trim()
}

function Write-RunmeCmdLauncher {
    param ([Parameter(Mandatory)] [string]$Folder)
    $cmdPath = Join-Path $Folder 'RUNME.cmd'
    $lines = @(
        '@echo off'
        'setlocal'
        'REM Double-click THIS file. Double-clicking RUNME.ps1 opens Notepad.'
        'set "HERE=%~dp0"'
        'set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"'
        'if not exist "%PS%" set "PS=powershell.exe"'
        'start "" "%PS%" -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "%HERE%RUNME.ps1" %*'
    )
    Set-Content -LiteralPath $cmdPath -Value $lines -Encoding ASCII -Force
}

# WinForms (paste-names / CAP dialogs) and the FAIS shortcut both want STA.
# "Run with PowerShell" starts MTA. Restart via CreateProcess in a hidden
# window; do not Start-Process the .ps1 itself (that is Notepad).
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
    Hide-RunmeConsole
    $here = $PSScriptRoot
    if ([string]::IsNullOrWhiteSpace($here)) { $here = Split-Path -Parent $PSCommandPath }
    try {
        Start-RunmePowerShellFile -ScriptPath $PSCommandPath -ArgumentLine (Get-RunmeStaRelaunchArgs) -WorkingDirectory $here
    }
    catch {
        Show-RunmeError "Could not restart RUNME.ps1 in STA:`n$($_.Exception.Message)`n`nRight-click RUNME.ps1 -> Run with PowerShell.`nDo not double-click RUNME.ps1 (Windows opens it in Notepad)."
    }
    return
}

Hide-RunmeConsole

# --- Resolve FAIS.ps1 and the icon (FAIS\App Icons\FAISlogo.ico) ---
$faisRoot   = $PSScriptRoot
$faisScript = Join-Path $faisRoot 'FAIS.ps1'
$iconPath   = Join-Path $faisRoot 'App Icons\FAISlogo.ico'

if (-not (Test-Path -LiteralPath $faisScript)) {
    Show-RunmeError "FAIS.ps1 was not found at $faisScript.`nRun this script from the FAIS folder (the same folder as FAIS.ps1)."
    return
}

# MOTW on Modules\*.psm1 makes Import-Module throw on RemoteSigned. Unblock
# the tree we ship. This does not open files.
Get-ChildItem -LiteralPath $faisRoot -Recurse -Include *.ps1,*.psm1,*.psd1 -File -ErrorAction SilentlyContinue |
    Unblock-File -ErrorAction SilentlyContinue

# Double-clickable launcher next to this script. Explorer executes .cmd;
# the .cmd then starts powershell.exe -File (never the Open verb).
try { Write-RunmeCmdLauncher -Folder $faisRoot } catch { }

# --- Copy ImportExcel into the Windows PowerShell 5.1 CurrentUser module path ---
# Destination:
#   <My Documents>\WindowsPowerShell\Modules\ImportExcel
# The Documents path contains a DOD ID folder (currently 1595161645A). On
# another computer that folder is two levels up from this script
# (C:\Users\<DoDID>\Desktop\FAIS9.1 -> C:\Users\<DoDID>). Use that folder
# name as-is — do not prompt, do not append A. 1:1 swap that folder only;
# the rest of the path stays the same.
# PowerShell looks there automatically, so FAIS.ps1's
#   Import-Module ImportExcel
# succeeds without Install-Module (which many DoD machines block).
$importExcelSrc     = Join-Path $faisRoot 'ImportExcel'
$documentsPath      = [Environment]::GetFolderPath('MyDocuments')
$userProfilePath    = Get-RunmeUserProfileTwoLevelsUp
$detectedDodFolder  = Get-RunmeDodFolderNameFromProfile $userProfilePath
if (-not [string]::IsNullOrWhiteSpace($detectedDodFolder)) {
    $documentsPath = Convert-RunmePathFolderName -Path $documentsPath -OldFolder $script:RunmeCurrentDodFolder -NewFolder $detectedDodFolder
}

# Folder picker at C:\Users\<DoDID>\. Chosen path -> $global:teamsLocation,
# then saved to TeamsLocation.txt so FAIS.ps1 writes ATIS.xlsx there.
# Cancel keeps whatever folder was saved last time.
$teamsLocationFile  = Join-Path $faisRoot 'TeamsLocation.txt'
$savedTeamsLocation = Read-RunmeTeamsLocation -FilePath $teamsLocationFile
[void](Show-RunmeTeamsFolderDialog -InitialDirectory $userProfilePath -CurrentLocation $savedTeamsLocation)
if (-not [string]::IsNullOrWhiteSpace($global:teamsLocation)) {
    try {
        Save-RunmeTeamsLocation -FilePath $teamsLocationFile -Location $global:teamsLocation
    }
    catch {
        Show-RunmeError "The Teams folder could not be saved to $teamsLocationFile`n$($_.Exception.Message)`n`nFAIS will keep writing ATIS.xlsx where it did before."
    }
}
else {
    $global:teamsLocation = $savedTeamsLocation
}

$userModulesPath    = Join-Path $documentsPath 'WindowsPowerShell\Modules'
$importExcelDest    = Join-Path $userModulesPath 'ImportExcel'

if (Test-Path -LiteralPath $importExcelSrc) {
    if (-not (Test-Path -LiteralPath $userModulesPath)) {
        New-Item -ItemType Directory -Path $userModulesPath -Force | Out-Null
    }
    Copy-Item -LiteralPath $importExcelSrc -Destination $userModulesPath -Recurse -Force
    Get-ChildItem -LiteralPath $importExcelDest -Recurse -File -ErrorAction SilentlyContinue |
        Unblock-File -ErrorAction SilentlyContinue
}

# --- Keep MasterData.xlsx next to RUNME.ps1 (do not move it one folder up) ---
$localMasterData  = Join-Path $faisRoot 'MasterData.xlsx'
$parentFolder     = Split-Path $faisRoot -Parent
$parentMasterData = Join-Path $parentFolder 'MasterData.xlsx'

# Older RUNME moved the workbook to the parent. Put it back next to this script
# unless a copy is already here (never overwrite a live local workbook).
if (-not (Test-Path -LiteralPath $localMasterData) -and (Test-Path -LiteralPath $parentMasterData)) {
    Move-Item -LiteralPath $parentMasterData -Destination $localMasterData
}

if (-not (Test-Path -LiteralPath $localMasterData)) {
    Write-Host "MasterData.xlsx was not found next to RUNME.ps1: $localMasterData" -ForegroundColor Yellow
}

# --- Keep OutputSchedule.xlsx next to RUNME.ps1 (do not copy it to the Desktop) ---
$localOutput   = Join-Path $faisRoot 'OutputSchedule.xlsx'
$desktopPath   = [Environment]::GetFolderPath('Desktop')
$desktopOutput = Join-Path $desktopPath 'OutputSchedule.xlsx'

# Older RUNME copied the workbook to the Desktop. Put it back next to this
# script unless a copy is already here (never overwrite a live local workbook).
if (-not (Test-Path -LiteralPath $localOutput) -and (Test-Path -LiteralPath $desktopOutput)) {
    Move-Item -LiteralPath $desktopOutput -Destination $localOutput
}

if (-not (Test-Path -LiteralPath $localOutput)) {
    Write-Host "OutputSchedule.xlsx was not found next to RUNME.ps1: $localOutput" -ForegroundColor Yellow
}

$powershell = Get-RunmePowerShellExe
if (-not (Test-Path -LiteralPath $powershell)) {
    Show-RunmeError "Windows PowerShell 5.1 was not found at $powershell"
    return
}

# Launch FAIS.ps1 through powershell.exe (never the .ps1 itself — that opens
# Notepad). -WindowStyle Hidden keeps the terminal off screen; WPF/WinForms still show.
$arguments = "-NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$faisScript`""

if (-not $Path) {
    $Path = [Environment]::GetFolderPath('Desktop')
}
if (-not (Test-Path -LiteralPath $Path)) {
    New-Item -ItemType Directory -Path $Path -Force | Out-Null
}

$destinations = @(Join-Path $Path "$Name.lnk")
if ($StartMenu) {
    $programs = Join-Path ([Environment]::GetFolderPath('StartMenu')) 'Programs'
    if (-not (Test-Path -LiteralPath $programs)) {
        New-Item -ItemType Directory -Path $programs -Force | Out-Null
    }
    $destinations += Join-Path $programs "$Name.lnk"
}

$iconOk = Test-Path -LiteralPath $iconPath

$shell = New-Object -ComObject WScript.Shell
try {
    foreach ($lnkPath in $destinations) {
        $shortcut = $shell.CreateShortcut($lnkPath)
        $shortcut.TargetPath       = $powershell
        $shortcut.Arguments        = $arguments
        $shortcut.WorkingDirectory = $faisRoot
        $shortcut.WindowStyle      = 7   # 7 = Minimized (console not visible)
        $shortcut.Description      = 'FAIS - Flight scheduler (puck board)'
        if ($iconOk) {
            $shortcut.IconLocation = "$iconPath,0"
        }
        $shortcut.Save()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($shortcut)
    }

    # Same-folder RUNME.lnk so double-click does not hit notepad.exe %1.
    $runmeLnk  = Join-Path $faisRoot 'RUNME.lnk'
    $runmeArgs = "-NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PSCommandPath`""
    $runmeSc   = $shell.CreateShortcut($runmeLnk)
    $runmeSc.TargetPath       = $powershell
    $runmeSc.Arguments        = $runmeArgs
    $runmeSc.WorkingDirectory = $faisRoot
    $runmeSc.WindowStyle      = 7
    $runmeSc.Description      = 'FAIS first-run setup. Right-click RUNME.ps1 -> Run with PowerShell also works.'
    if ($iconOk) { $runmeSc.IconLocation = "$iconPath,0" }
    $runmeSc.Save()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($runmeSc)

    # Desktop (or -Path) shortcut that opens MasterData.xlsx in this folder.
    if (Test-Path -LiteralPath $localMasterData) {
        $mdLnkPath = Join-Path $Path 'MasterData.lnk'
        $mdShortcut = $shell.CreateShortcut($mdLnkPath)
        $mdShortcut.TargetPath       = $localMasterData
        $mdShortcut.WorkingDirectory = $faisRoot
        $mdShortcut.WindowStyle      = 1
        $mdShortcut.Description      = 'FAIS MasterData.xlsx'
        $mdShortcut.Save()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($mdShortcut)
        if ($StartMenu) {
            $mdStart = Join-Path $programs 'MasterData.lnk'
            $mdShortcut2 = $shell.CreateShortcut($mdStart)
            $mdShortcut2.TargetPath       = $localMasterData
            $mdShortcut2.WorkingDirectory = $faisRoot
            $mdShortcut2.WindowStyle      = 1
            $mdShortcut2.Description      = 'FAIS MasterData.xlsx'
            $mdShortcut2.Save()
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($mdShortcut2)
        }
    }

    # Desktop (or -Path) shortcut that opens OutputSchedule.xlsx in this folder.
    if (Test-Path -LiteralPath $localOutput) {
        $osLnkPath = Join-Path $Path 'OutputSchedule.lnk'
        $osShortcut = $shell.CreateShortcut($osLnkPath)
        $osShortcut.TargetPath       = $localOutput
        $osShortcut.WorkingDirectory = $faisRoot
        $osShortcut.WindowStyle      = 1
        $osShortcut.Description      = 'FAIS OutputSchedule.xlsx'
        $osShortcut.Save()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($osShortcut)
        if ($StartMenu) {
            $osStart = Join-Path $programs 'OutputSchedule.lnk'
            $osShortcut2 = $shell.CreateShortcut($osStart)
            $osShortcut2.TargetPath       = $localOutput
            $osShortcut2.WorkingDirectory = $faisRoot
            $osShortcut2.WindowStyle      = 1
            $osShortcut2.Description      = 'FAIS OutputSchedule.xlsx'
            $osShortcut2.Save()
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($osShortcut2)
        }
    }

    # Repair leftover .lnk files whose TargetPath is a .ps1/.psm1 (Open = Notepad).
    $lnkDirs = @($faisRoot, $Path)
    if ($StartMenu -and $programs) { $lnkDirs += $programs }
    foreach ($dir in ($lnkDirs | Select-Object -Unique)) {
        if (-not (Test-Path -LiteralPath $dir)) { continue }
        Get-ChildItem -LiteralPath $dir -Filter '*.lnk' -File -ErrorAction SilentlyContinue | ForEach-Object {
            $fix = $null
            try {
                $fix = $shell.CreateShortcut($_.FullName)
                if (-not (Test-RunmeIsScriptPath $fix.TargetPath)) { return }
                $scriptToRun = $fix.TargetPath
                $fix.TargetPath = $powershell
                if ([string]::IsNullOrWhiteSpace($fix.Arguments) -or $fix.Arguments -notmatch '-File') {
                    $fix.Arguments = "-NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptToRun`""
                }
                $fix.Save()
            }
            catch { }
            finally {
                if ($null -ne $fix) {
                    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($fix)
                }
            }
        }
    }
}
finally {
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell)
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# --- Add students: paste Last, First, mark CAP / International, write Excel ---
# Parser, Students-sheet write, String Tracker StudentName write, and blank
# syllabus tabs live in THIS script so a stale PowerShell 5.1 module export
# list cannot break Add Students. FAIS Modules\*.psm1 are NOT imported here.
# Loading DataManipulation.psm1, DataUpload.psm1, and PreRunMenu.psm1 was
# opening those three files in Notepad (Open verb / ShellExecute).
if ($SkipAddStudents) {
    return
}

function Initialize-RunmeWinForms {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()
}

function ConvertTo-RunmeBool {
    param ($Value, [bool]$Default = $false)

    if ($null -eq $Value) { return $Default }
    if ($Value -is [bool]) { return $Value }
    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) { return ($Value -ne 0) }

    $text = ([string]$Value).Trim().ToLowerInvariant()
    if ($text -eq '') { return $Default }
    if ($text -in @('true', 'yes', 'y', '1', 'x', '-1')) { return $true }
    if ($text -in @('false', 'no', 'n', '0')) { return $false }
    return $Default
}

# "Last, First" with one space after the comma, extra spaces squeezed out.
function ConvertTo-RunmeStudentDisplayName {
    param ([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $clean = ($Text.Trim() -replace '\xA0', ' ' -replace '\s+', ' ')
    $comma = $clean.IndexOf(',')
    if ($comma -lt 0) { return $clean }
    $last = $clean.Substring(0, $comma).Trim()
    $first = $clean.Substring($comma + 1).Trim()
    if ($last -eq '' -or $first -eq '') { return $clean }
    return "$last, $first"
}

function Test-RunmeStudentSheetName {
    param ([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
    if ($Name.Length -gt 31) { return $false }
    if ($Name -match '[\[\]\*\?/\\:]') { return $false }
    return $true
}

# Paste box -> unique "Last, First" names. Blank lines are skipped.
# Names with no comma, an empty last/first, or a duplicate are listed in Errors;
# they are not in Names.
function ConvertTo-RunmePastedStudentNames {
    param ([string]$Text)

    $names = New-Object System.Collections.Generic.List[string]
    $errors = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    if ([string]::IsNullOrWhiteSpace($Text)) {
        return [PSCustomObject]@{ Names = @(); Errors = @() }
    }

    $lines = @($Text -replace "`r", '' -split "`n")
    $lineNo = 0
    foreach ($raw in $lines) {
        $lineNo++
        $trimmed = if ($null -eq $raw) { '' } else { $raw.Trim() }
        if ($trimmed -eq '') { continue }

        $comma = $trimmed.IndexOf(',')
        if ($comma -lt 0) {
            $errors.Add("Line ${lineNo}: '$trimmed' is not 'Last, First' (missing comma).")
            continue
        }
        $last = $trimmed.Substring(0, $comma).Trim()
        $first = $trimmed.Substring($comma + 1).Trim()
        if ($last -eq '' -or $first -eq '') {
            $errors.Add("Line ${lineNo}: '$trimmed' is not 'Last, First' (blank last or first name).")
            continue
        }
        $display = ConvertTo-RunmeStudentDisplayName $trimmed
        $key = $display.ToLowerInvariant()
        if ($seen.ContainsKey($key)) {
            $errors.Add("Line ${lineNo}: '$display' is listed more than once; the first copy is kept.")
            continue
        }
        $seen[$key] = $true
        $names.Add($display)
    }

    return [PSCustomObject]@{
        Names  = $names.ToArray()
        Errors = $errors.ToArray()
    }
}

function Test-RunmeFileWritable {
    param ([string]$FilePath)

    if ([string]::IsNullOrWhiteSpace($FilePath)) { return $true }
    if (-not (Test-Path -LiteralPath $FilePath)) { return $true }
    try {
        $fs = [System.IO.File]::Open($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        $fs.Close()
        $fs.Dispose()
        return $true
    }
    catch {
        return $false
    }
}

function Show-RunmeAddStudentsNameDialog {
    Initialize-RunmeWinForms

    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "Add Students"
    $Form.Size = New-Object System.Drawing.Size(520, 520)
    $Form.StartPosition = "CenterScreen"
    $Form.Font = New-Object System.Drawing.Font("Segoe UI", 11)
    $Form.MinimizeBox = $false

    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = "Paste student names, one per line, as Last, First"
    $Label.Location = New-Object System.Drawing.Point(20, 16)
    $Label.AutoSize = $true
    $Form.Controls.Add($Label)

    $Hint = New-Object System.Windows.Forms.Label
    $Hint.Text = "Example:`nByrnes, Aidan`nGunn, Christopher"
    $Hint.Location = New-Object System.Drawing.Point(20, 44)
    $Hint.Size = New-Object System.Drawing.Size(460, 52)
    $Hint.ForeColor = [System.Drawing.Color]::DimGray
    $Form.Controls.Add($Hint)

    $Box = New-Object System.Windows.Forms.TextBox
    $Box.Multiline = $true
    $Box.ScrollBars = 'Vertical'
    $Box.AcceptsReturn = $true
    $Box.Location = New-Object System.Drawing.Point(20, 100)
    $Box.Size = New-Object System.Drawing.Size(460, 300)
    $Box.Font = New-Object System.Drawing.Font("Consolas", 11)
    $Form.Controls.Add($Box)

    $AddBtn = New-Object System.Windows.Forms.Button
    $AddBtn.Text = "Next"
    $AddBtn.Location = New-Object System.Drawing.Point(20, 420)
    $AddBtn.Size = New-Object System.Drawing.Size(160, 40)
    $AddBtn.BackColor = [System.Drawing.Color]::LightGreen
    $AddBtn.Add_Click({
        $Form.Tag = $Box.Text
        $Form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $Form.Close()
    })
    $Form.Controls.Add($AddBtn)

    $CancelBtn = New-Object System.Windows.Forms.Button
    $CancelBtn.Text = "Cancel"
    $CancelBtn.Location = New-Object System.Drawing.Point(320, 420)
    $CancelBtn.Size = New-Object System.Drawing.Size(160, 40)
    $CancelBtn.Add_Click({
        $Form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $Form.Close()
    })
    $Form.Controls.Add($CancelBtn)

    $result = $Form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) { return [string]$Form.Tag }
    return $null
}

function Show-RunmeCapInternationalQuestion {
    Initialize-RunmeWinForms
    $answer = [System.Windows.Forms.MessageBox]::Show(
        "Is anyone CAP or International?`n`nYes: pick who. Their IsCAP / IsInternational columns are marked TRUE.`nNo: every pasted student is FALSE in both columns.",
        "CAP / International",
        [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) { return $true }
    if ($answer -eq [System.Windows.Forms.DialogResult]::No) { return $false }
    return $null
}

function Show-RunmeCapInternationalPicker {
    param ([array]$Entries = @())

    Initialize-RunmeWinForms

    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "CAP / International"
    $Form.Size = New-Object System.Drawing.Size(560, 480)
    $Form.StartPosition = "CenterScreen"
    $Form.Font = New-Object System.Drawing.Font("Segoe UI", 11)
    $Form.MinimizeBox = $false

    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = "Check CAP and/or International for each student it applies to:"
    $Label.Location = New-Object System.Drawing.Point(20, 16)
    $Label.AutoSize = $true
    $Form.Controls.Add($Label)

    $Grid = New-Object System.Windows.Forms.DataGridView
    $Grid.Location = New-Object System.Drawing.Point(20, 50)
    $Grid.Size = New-Object System.Drawing.Size(500, 320)
    $Grid.AllowUserToAddRows = $false
    $Grid.AllowUserToDeleteRows = $false
    $Grid.RowHeadersVisible = $false
    $Grid.SelectionMode = 'FullRowSelect'
    $Grid.AutoSizeColumnsMode = 'Fill'
    $Grid.BackgroundColor = [System.Drawing.Color]::White

    $nameCol = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $nameCol.HeaderText = 'Student'
    $nameCol.ReadOnly = $true
    $nameCol.FillWeight = 70
    [void]$Grid.Columns.Add($nameCol)

    $capCol = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $capCol.HeaderText = 'CAP'
    $capCol.FillWeight = 15
    [void]$Grid.Columns.Add($capCol)

    $intlCol = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $intlCol.HeaderText = 'International'
    $intlCol.FillWeight = 20
    [void]$Grid.Columns.Add($intlCol)

    foreach ($e in @($Entries)) {
        $isCap = $false
        $isIntl = $false
        if ($null -ne $e.IsCap) { $isCap = [bool]$e.IsCap }
        if ($null -ne $e.IsInternational) { $isIntl = [bool]$e.IsInternational }
        [void]$Grid.Rows.Add($e.Name, $isCap, $isIntl)
    }
    $Form.Controls.Add($Grid)

    $SaveBtn = New-Object System.Windows.Forms.Button
    $SaveBtn.Text = "Save flags"
    $SaveBtn.Location = New-Object System.Drawing.Point(20, 385)
    $SaveBtn.Size = New-Object System.Drawing.Size(160, 40)
    $SaveBtn.BackColor = [System.Drawing.Color]::LightGreen
    $SaveBtn.Add_Click({
        $Grid.EndEdit()
        $picked = New-Object System.Collections.Generic.List[object]
        foreach ($row in $Grid.Rows) {
            if ($row.IsNewRow) { continue }
            $picked.Add([PSCustomObject]@{
                Name            = [string]$row.Cells[0].Value
                IsCap           = [bool]$row.Cells[1].Value
                IsInternational = [bool]$row.Cells[2].Value
            })
        }
        $Form.Tag = $picked.ToArray()
        $Form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $Form.Close()
    })
    $Form.Controls.Add($SaveBtn)

    $CancelBtn = New-Object System.Windows.Forms.Button
    $CancelBtn.Text = "Cancel"
    $CancelBtn.Location = New-Object System.Drawing.Point(360, 385)
    $CancelBtn.Size = New-Object System.Drawing.Size(160, 40)
    $CancelBtn.Add_Click({
        $Form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $Form.Close()
    })
    $Form.Controls.Add($CancelBtn)

    $result = $Form.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    return $Form.Tag
}

# Worksheet lookup that ignores spaces and case ("String Tracker" == "StringTracker").
function Get-RunmeWorksheetLoose {
    param ($Workbook, [string]$Name)

    if ($null -eq $Workbook -or [string]::IsNullOrWhiteSpace($Name)) { return $null }
    $want = (($Name -replace '\s+', '')).ToLowerInvariant()
    foreach ($ws in $Workbook.Worksheets) {
        $have = (($ws.Name -replace '\s+', '')).ToLowerInvariant()
        if ($have -eq $want) { return $ws }
    }
    return $null
}

# First cell in rows 1-15 whose heading matches (spaces/case ignored).
function Get-RunmeHeaderCell {
    param ($Worksheet, [string]$Heading, [int]$MaxHeaderRow = 15)

    if ($null -eq $Worksheet -or [string]::IsNullOrWhiteSpace($Heading)) { return $null }
    $want = (($Heading -replace '\s+', '')).ToLowerInvariant()
    $endRow = $MaxHeaderRow
    $endCol = 40
    if ($null -ne $Worksheet.Dimension) {
        $endRow = [Math]::Min($MaxHeaderRow, [int]$Worksheet.Dimension.End.Row)
        $endCol = [Math]::Max(1, [int]$Worksheet.Dimension.End.Column)
    }
    for ($r = 1; $r -le $endRow; $r++) {
        for ($c = 1; $c -le $endCol; $c++) {
            $text = ([string]$Worksheet.Cells[$r, $c].Value).Trim()
            if ($text -eq '') { continue }
            $have = (($text -replace '\s+', '')).ToLowerInvariant()
            if ($have -eq $want) {
                return [PSCustomObject]@{ Row = $r; Column = $c }
            }
        }
    }
    return $null
}

# Writes unique names into String Tracker!StudentName on an already-open package.
# Fills gaps above the last named row, then appends. Does not save or close.
function Add-RunmeStringTrackerNamesToPackage {
    param (
        $ExcelPackage,
        [array]$Names = @()
    )

    $out = [PSCustomObject]@{
        Found      = $false
        Added      = 0
        SheetName  = ''
        Error      = ''
    }
    if ($null -eq $ExcelPackage) { return $out }

    $ws = Get-RunmeWorksheetLoose -Workbook $ExcelPackage.Workbook -Name 'String Tracker'
    if ($null -eq $ws) { return $out }
    $out.Found = $true
    $out.SheetName = $ws.Name

    $header = Get-RunmeHeaderCell -Worksheet $ws -Heading 'StudentName'
    if ($null -eq $header) {
        $out.Error = "Sheet '$($ws.Name)' has no 'StudentName' heading in the first 15 rows."
        return $out
    }

    $nameCol = [int]$header.Column
    $startRow = [int]$header.Row + 1
    $endRow = $startRow
    if ($null -ne $ws.Dimension) {
        $endRow = [Math]::Max($startRow, [int]$ws.Dimension.End.Row)
    }

    $existing = @{}
    $gapRows = New-Object System.Collections.Generic.List[int]
    $lastNamed = [int]$header.Row
    for ($r = $startRow; $r -le $endRow; $r++) {
        $raw = $ws.Cells[$r, $nameCol].Value
        if ([string]::IsNullOrWhiteSpace([string]$raw)) { continue }
        $display = ConvertTo-RunmeStudentDisplayName ([string]$raw)
        if ($display -eq '') { continue }
        $existing[$display.ToLowerInvariant()] = $r
        $lastNamed = $r
    }
    for ($r = $startRow; $r -le $lastNamed; $r++) {
        if ([string]::IsNullOrWhiteSpace([string]$ws.Cells[$r, $nameCol].Value)) {
            $gapRows.Add($r)
        }
    }

    $gapIdx = 0
    $appendRow = $lastNamed + 1
    foreach ($n in @($Names)) {
        $display = ConvertTo-RunmeStudentDisplayName ([string]$n)
        if ($display -eq '') { continue }
        $key = $display.ToLowerInvariant()
        if ($existing.ContainsKey($key)) { continue }

        $r = 0
        if ($gapIdx -lt $gapRows.Count) {
            $r = [int]$gapRows[$gapIdx]
            $gapIdx++
        }
        else {
            $r = $appendRow
            $appendRow++
        }

        $ws.Cells[$r, $nameCol].Value = $display
        $ws.Cells[$r, $nameCol].Style.Numberformat.Format = '@'
        $existing[$key] = $r
        $out.Added++
    }
    return $out
}

# MasterData Syllabus sheet as blank tab rows: every event incomplete, no IP / ETD.
# Columns match existing OutputSchedule student tabs:
#   SyllabusID, Mission, Prereqs, Instructor, ETD, Type, IsCheck, IsComplete, OptingEvents, Level
function Get-RunmeBlankSyllabusRows {
    param ([Parameter(Mandatory)] [string]$MasterDataPath)

    $raw = @(Import-Excel -Path $MasterDataPath -WorksheetName 'Syllabus' -ErrorAction Stop)
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($mission in $raw) {
        if ($null -eq $mission -or [string]::IsNullOrWhiteSpace([string]$mission.Mission)) { continue }
        $opt = $mission.OptingEvents
        if ($null -ne $opt -and [string]$opt -like '=*') { $opt = 0 }
        $idRaw = $mission.SyllabusID
        $idNum = 0
        $id = $idRaw
        if ($null -ne $idRaw -and [int]::TryParse(([string]$idRaw).Trim(), [ref]$idNum)) { $id = $idNum }
        $type = if ($null -eq $mission.Type) { '' } else { ([string]$mission.Type).Trim() }
        $rows.Add([PSCustomObject][ordered]@{
            SyllabusID   = $id
            Mission      = ([string]$mission.Mission).Trim()
            Prereqs      = $mission.Prereqs
            Instructor   = ''
            ETD          = ''
            Type         = $type
            IsCheck      = [bool](ConvertTo-RunmeBool $mission.IsCheck)
            IsComplete   = $false
            OptingEvents = $opt
            Level        = $mission.Level
        })
    }
    if ($rows.Count -eq 0) {
        throw "The MasterData Syllabus sheet has no missions, so a blank student tab cannot be built."
    }
    return $rows.ToArray()
}

# Writes new names onto MasterData Students (same columns as the current sheet),
# onto String Tracker!StudentName (OutputSchedule.xlsx first, then MasterData),
# and gives each new student a blank syllabus tab in OutputSchedule. Existing
# names are not duplicated; their IsCAP / IsInternational cells are updated
# when -UpdateExistingFlags is set. Existing syllabus tabs are never cleared.
function Add-RunmeStudentRoster {
    param (
        [Parameter(Mandatory)] [string]$MasterDataPath,
        [string]$OutputPath,
        [array]$Entries = @(),
        [switch]$UpdateExistingFlags
    )

    $result = [PSCustomObject]@{
        Added                = @()
        TabsCreated          = 0
        MasterDataAdded      = 0
        MasterDataUpdated    = 0
        StringTrackerAdded   = 0
        StringTrackerSheets  = @()
        SkippedTabs          = @()
        Error                = ''
    }

    $valid = @(foreach ($e in @($Entries)) {
        if ($null -eq $e) { continue }
        $n = ConvertTo-RunmeStudentDisplayName ([string]$e.Name)
        if ($n -eq '') { continue }
        [PSCustomObject]@{
            Name            = $n
            IsCap           = [bool](ConvertTo-RunmeBool $e.IsCap)
            IsInternational = [bool](ConvertTo-RunmeBool $e.IsInternational)
        }
    })
    if ($valid.Count -eq 0) { return $result }
    if (-not (Test-Path -LiteralPath $MasterDataPath)) {
        throw "MasterData.xlsx was not found at $MasterDataPath"
    }

    $toWrite = @($MasterDataPath)
    if (-not [string]::IsNullOrWhiteSpace($OutputPath)) { $toWrite += $OutputPath }
    foreach ($file in $toWrite) {
        if (-not (Test-Path -LiteralPath $file)) { continue }
        if (-not (Test-RunmeFileWritable $file)) {
            throw "$file is open (Excel still has a lock on it). Close it, then run RUNME again."
        }
    }

    $blankRows = @(Get-RunmeBlankSyllabusRows -MasterDataPath $MasterDataPath)
    $addedNames = New-Object System.Collections.Generic.List[string]
    $tabNames = New-Object System.Collections.Generic.List[string]
    $skippedTabs = New-Object System.Collections.Generic.List[string]
    $trackerHits = New-Object System.Collections.Generic.List[object]
    $trackerNames = @()
    $mdHasTracker = $false
    $counts = @{ Added = 0; Updated = 0 }
    $doUpdateFlags = [bool]$UpdateExistingFlags

    $package = $null
    try {
        $package = Open-ExcelPackage -Path $MasterDataPath
        $ws = $package.Workbook.Worksheets['Students']
        if ($null -eq $ws) { throw "Worksheet 'Students' was not found in MasterData.xlsx" }

        $col = @{}
        $lastCol = if ($null -ne $ws.Dimension) { $ws.Dimension.End.Column } else { 0 }
        for ($c = 1; $c -le $lastCol; $c++) {
            $heading = ([string]$ws.Cells[1, $c].Value).Trim()
            if ($heading -ne '' -and -not $col.ContainsKey($heading)) { $col[$heading] = $c }
        }
        foreach ($need in @('SPName', 'IsCAP', 'IsInternational')) {
            if (-not $col.ContainsKey($need)) {
                throw "The MasterData Students sheet has no '$need' heading in row 1."
            }
        }
        $nameCol = $col['SPName']
        $keyCol = if ($col.ContainsKey('SPKey')) { $col['SPKey'] } else { 0 }
        $capCol = $col['IsCAP']
        $intlCol = $col['IsInternational']
        $priCol = if ($col.ContainsKey('Priority')) { $col['Priority'] } else { 0 }
        $unsatSimCol = if ($col.ContainsKey('UnsatSim')) { $col['UnsatSim'] } else { 0 }
        $unsatFltCol = if ($col.ContainsKey('UnsatFlight')) { $col['UnsatFlight'] } else { 0 }
        $holdCol = if ($col.ContainsKey('OnHold')) { $col['OnHold'] } else { 0 }
        $debtCol = if ($col.ContainsKey('SimDebt')) { $col['SimDebt'] } else { 0 }

        $endRow = if ($null -ne $ws.Dimension) { $ws.Dimension.End.Row } else { 1 }
        $existingRow = @{}
        $maxKey = 0
        $lastNamed = 1
        for ($r = 2; $r -le $endRow; $r++) {
            $nameVal = $ws.Cells[$r, $nameCol].Value
            if ([string]::IsNullOrWhiteSpace([string]$nameVal)) { continue }
            $existingRow[$nameVal.ToString().Trim().ToLowerInvariant()] = $r
            $lastNamed = $r
            if ($keyCol -gt 0) {
                $keyNum = 0
                $rawKey = $ws.Cells[$r, $keyCol].Value
                if ($null -ne $rawKey -and [int]::TryParse(([string]$rawKey).Trim(), [ref]$keyNum)) {
                    if ($keyNum -gt $maxKey) { $maxKey = $keyNum }
                }
            }
        }
        $appendRow = $lastNamed + 1

        foreach ($entry in $valid) {
            $key = $entry.Name.ToLowerInvariant()
            if ($existingRow.ContainsKey($key)) {
                if ($doUpdateFlags) {
                    $r = $existingRow[$key]
                    $ws.Cells[$r, $capCol].Value = [bool]$entry.IsCap
                    $ws.Cells[$r, $intlCol].Value = [bool]$entry.IsInternational
                    $counts.Updated++
                }
            }
            else {
                $maxKey++
                $idText = if ($maxKey -lt 100) { $maxKey.ToString('00') } else { [string]$maxKey }
                $r = $appendRow
                $appendRow++
                if ($keyCol -gt 0) {
                    $ws.Cells[$r, $keyCol].Value = $idText
                    $ws.Cells[$r, $keyCol].Style.Numberformat.Format = '@'
                }
                $ws.Cells[$r, $nameCol].Value = $entry.Name
                $ws.Cells[$r, $capCol].Value = [bool]$entry.IsCap
                $ws.Cells[$r, $intlCol].Value = [bool]$entry.IsInternational
                if ($priCol -gt 0) { $ws.Cells[$r, $priCol].Value = 1 }
                if ($unsatSimCol -gt 0) { $ws.Cells[$r, $unsatSimCol].Value = $false }
                if ($unsatFltCol -gt 0) { $ws.Cells[$r, $unsatFltCol].Value = $false }
                if ($holdCol -gt 0) { $ws.Cells[$r, $holdCol].Value = $false }
                if ($debtCol -gt 0) {
                    $ws.Cells[$r, $debtCol].Value = 0
                    $ws.Cells[$r, $debtCol].Style.Numberformat.Format = '0.0'
                }
                $existingRow[$key] = $r
                $counts.Added++
                $addedNames.Add($entry.Name)
                if (Test-RunmeStudentSheetName $entry.Name) { $tabNames.Add($entry.Name) }
                else { $skippedTabs.Add($entry.Name) }
            }
        }

        # Pasted names go in String Tracker!StudentName even if they were
        # already on the Students sheet (the board is a separate list).
        # Prefer OutputSchedule.xlsx; only write on MasterData when that is
        # the only workbook that has the sheet.
        $trackerNames = @($valid | ForEach-Object { $_.Name })
        $mdHasTracker = $null -ne (Get-RunmeWorksheetLoose -Workbook $package.Workbook -Name 'String Tracker')
        $osFileExists = -not [string]::IsNullOrWhiteSpace($OutputPath) -and (Test-Path -LiteralPath $OutputPath)
        if ($mdHasTracker -and -not $osFileExists) {
            $mdTracker = Add-RunmeStringTrackerNamesToPackage -ExcelPackage $package -Names $trackerNames
            if ($mdTracker.Error) { throw "MasterData.xlsx $($mdTracker.Error)" }
            $result.StringTrackerAdded += [int]$mdTracker.Added
            $trackerHits.Add("MasterData.xlsx '$($mdTracker.SheetName)'")
        }

        Close-ExcelPackage -ExcelPackage $package
        $package = $null
    }
    catch {
        if ($null -ne $package) { Close-ExcelPackage -ExcelPackage $package -NoSave }
        throw
    }

    $tabsCreated = 0
    if (-not [string]::IsNullOrWhiteSpace($OutputPath) -and (Test-Path -LiteralPath $OutputPath)) {
        $outPkg = $null
        try {
            $outPkg = Open-ExcelPackage -Path $OutputPath -Create
            foreach ($sheet in $tabNames) {
                $existing = $outPkg.Workbook.Worksheets[$sheet]
                if ($null -ne $existing -and $null -ne $existing.Dimension) { continue }
                $outPkg = $blankRows | Export-Excel -ExcelPackage $outPkg -WorksheetName $sheet -AutoSize -PassThru -ErrorAction Stop
                $tabsCreated++
            }

            $osTracker = Add-RunmeStringTrackerNamesToPackage -ExcelPackage $outPkg -Names $trackerNames
            if ($osTracker.Found) {
                if ($osTracker.Error) { throw "OutputSchedule.xlsx $($osTracker.Error)" }
                $result.StringTrackerAdded += [int]$osTracker.Added
                $trackerHits.Add("OutputSchedule.xlsx '$($osTracker.SheetName)'")
            }

            Close-ExcelPackage -ExcelPackage $outPkg
            $outPkg = $null
        }
        catch {
            if ($null -ne $outPkg) { Close-ExcelPackage -ExcelPackage $outPkg -NoSave }
            throw
        }
    }

    # OutputSchedule had no String Tracker — fall back to MasterData.
    if ($trackerHits.Count -eq 0 -and $mdHasTracker) {
        $package = $null
        try {
            $package = Open-ExcelPackage -Path $MasterDataPath
            $mdTracker = Add-RunmeStringTrackerNamesToPackage -ExcelPackage $package -Names $trackerNames
            if ($mdTracker.Error) { throw "MasterData.xlsx $($mdTracker.Error)" }
            $result.StringTrackerAdded += [int]$mdTracker.Added
            $trackerHits.Add("MasterData.xlsx '$($mdTracker.SheetName)'")
            Close-ExcelPackage -ExcelPackage $package
            $package = $null
        }
        catch {
            if ($null -ne $package) { Close-ExcelPackage -ExcelPackage $package -NoSave }
            throw
        }
    }

    if ($trackerHits.Count -eq 0) {
        $looked = @($MasterDataPath)
        if (-not [string]::IsNullOrWhiteSpace($OutputPath)) { $looked += $OutputPath }
        throw "No 'String Tracker' sheet was found (looked for a 'StudentName' column). Checked:`n$($looked -join "`n")"
    }

    $result.Added = $addedNames.ToArray()
    $result.TabsCreated = $tabsCreated
    $result.MasterDataAdded = [int]$counts.Added
    $result.MasterDataUpdated = [int]$counts.Updated
    $result.StringTrackerSheets = $trackerHits.ToArray()
    $result.SkippedTabs = $skippedTabs.ToArray()
    Write-Host ("Students sheet: {0} added, {1} flag(s) updated. String Tracker StudentName: {2} written ({3}). {4} blank syllabus tab(s)." -f $counts.Added, $counts.Updated, $result.StringTrackerAdded, ($trackerHits -join '; '), $tabsCreated) -ForegroundColor Green
    return $result
}

# Student tabs are written into the workbook next to this script (the same
# file OutputSchedule.lnk on the Desktop opens). Do not copy it to Desktop.
$outputPath = $localOutput

$masterDataPath = $localMasterData
if (-not (Test-Path -LiteralPath $masterDataPath)) {
    try {
        Initialize-RunmeWinForms
        [void][System.Windows.Forms.MessageBox]::Show(
            "MasterData.xlsx was not found next to RUNME.ps1:`n$masterDataPath`n`nPlace MasterData.xlsx in the same folder as RUNME.ps1, then run this script again.",
            "MasterData.xlsx missing",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    }
    catch { }
    return
}

try {
    Import-Module ImportExcel -ErrorAction Stop
}
catch {
    try {
        Initialize-RunmeWinForms
        [void][System.Windows.Forms.MessageBox]::Show(
            "ImportExcel could not be loaded, so students were not added:`n$($_.Exception.Message)`n`nCopy the ImportExcel folder next to FAIS.ps1, run this script again, then paste names.",
            "ImportExcel missing",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    }
    catch { }
    return
}

# Always use the inlined parser / roster writer. Do not call into
# Modules\*.psm1 even if a profile already imported them.
try {
    Initialize-RunmeWinForms
    $text = Show-RunmeAddStudentsNameDialog
    if ($null -eq $text) {
        return
    }

    $parsed = ConvertTo-RunmePastedStudentNames -Text $text
    if ($parsed.Errors.Count -gt 0) {
        $errText = $parsed.Errors -join "`n"
        if ($parsed.Names.Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show($errText, "No valid names", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $keep = [System.Windows.Forms.MessageBox]::Show(
            "$errText`n`nContinue with $($parsed.Names.Count) valid name(s)?",
            "Some lines were skipped",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($keep -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }
    if ($parsed.Names.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show("Paste at least one name as Last, First.", "No names", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $anyone = Show-RunmeCapInternationalQuestion
    if ($null -eq $anyone) {
        return
    }

    $entries = @(foreach ($n in $parsed.Names) {
        [PSCustomObject]@{ Name = $n; IsCap = $false; IsInternational = $false }
    })
    if ($anyone) {
        $picked = Show-RunmeCapInternationalPicker -Entries $entries
        if ($null -eq $picked) {
            return
        }
        $entries = @($picked)
    }

    $written = Add-RunmeStudentRoster -MasterDataPath $masterDataPath -OutputPath $outputPath -Entries $entries -UpdateExistingFlags:$anyone

    $skip = if ($written.SkippedTabs.Count -gt 0) { "`nTabs skipped (name cannot be an Excel sheet): $($written.SkippedTabs -join '; ')" } else { '' }
    $trackerWhere = if ($written.StringTrackerSheets.Count -gt 0) { $written.StringTrackerSheets -join '; ' } else { 'none' }
    [void][System.Windows.Forms.MessageBox]::Show(
        "MasterData Students: $($written.MasterDataAdded) added, $($written.MasterDataUpdated) CAP/International flag(s) updated.`nString Tracker StudentName: $($written.StringTrackerAdded) name(s) written ($trackerWhere).`nOutputSchedule: $($written.TabsCreated) blank syllabus tab(s) created (existing tabs were left as they were).$skip",
        "Students added",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
}
catch {
    try {
        Initialize-RunmeWinForms
        [void][System.Windows.Forms.MessageBox]::Show(
            "The roster could not be written:`n$($_.Exception.Message)",
            "Add Students failed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    catch { }
}
