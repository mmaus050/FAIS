#FAIS
# cd C:\Users\1595161645A\Desktop\FAIS7.1
#
# Run order:
#   1. Open MasterData / OutputSchedule workbooks are saved in Excel so the latest edits are read
#   2. Pre-run pop-ups: log hooked (unsat) / incomplete events from the last schedule, HOLD status,
#      Add Student Absence (all saved right away)
#   3. Puck board: lines come from MasterData's Schedule sheet, a tab for each day it has lines for
#      (re-read whenever MasterData is saved), or are typed; hardcode IPs and students, see the + / - and rule flags.
#      Add Student Absence blocks students off (saved to MasterData's StudentAbsences sheet right away)
#   4. "Fill this day" or "Fill all days": the scheduler fills the open lines, you review it
#   5. Export, all or nothing, for the filled days: OutputSchedule.xlsx (MasterSchedule, other exported
#      days kept, + each student's syllabus tab), SimDebt + cleaned-up unsats to MasterData, and
#      ATIS.xlsx (one tab per day: flights, then sims). Excel files are saved and closed first.
#      A window then says what was written (or why nothing was), with buttons to open the files.
#
# -Trace prints every scheduling decision to the console.
param (
    [switch]$Trace
)


# =================================================================================
#  ____ ___  _   _ _____ ___ ____ _   _ ____      _  _____ ___ ___  _   _
# / ___/ _ \| \ | |  ___|_ _/ ___| | | |  _ \    / \|_   _|_ _/ _ \| \ | |
#| |  | | | |  \| | |_   | | |  _| | | | |_) |  / _ \ | |  | | | | |  \| |
#| |__| |_| | |\  |  _|  | | |_| | |_| |  _ <  / ___ \| |  | | |_| | |\  |
# \____\___/|_| \_|_|   |___\____|\___/|_| \_\/_/   \_\_| |___\___/|_| \_|
# =================================================================================


# ==========================================
#       STARTUP CHECKS
# ==========================================
# The puck board is a WPF window: it needs full language mode and an STA thread.
if ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') {
    Write-Host "FAIS needs PowerShell in FullLanguage mode (the puck board uses WPF). Current mode: $($ExecutionContext.SessionState.LanguageMode)" -ForegroundColor Yellow
    return
}
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
    Write-Host "Restarting PowerShell in STA mode for the puck board..." -ForegroundColor Cyan
    $exe = (Get-Process -Id $PID).Path
    $argList = @('-NoProfile', '-STA')
    $policy = Get-ExecutionPolicy -Scope Process
    if ($policy -ne 'Undefined') { $argList += @('-ExecutionPolicy', [string]$policy) }
    $argList += @('-File', "`"$PSCommandPath`"")
    if ($Trace) { $argList += '-Trace' }
    Start-Process -FilePath $exe -ArgumentList $argList -Wait -NoNewWindow
    return
}


# ==========================================
#       CONFIGURATION
# ==========================================
$excelInputPath = Join-Path $PSScriptRoot "MasterData.xlsx"
$desktopPath = [Environment]::GetFolderPath('Desktop')
$excelOutputPath = Join-Path -Path $PSScriptRoot -ChildPath "OutputSchedule.xlsx"
$atisOutputPath = Join-Path -Path $desktopPath -ChildPath "ATIS.xlsx"

# Puck board
$BoardSettings = @{
    OBReturnOffsetMinutes = 180                                 # typing "1015 O&B" adds its return line this many minutes later
    ReturnSortieType      = 'O&B Return'                        # SortieType written on return lines
    SimInstructor         = 'SIM IP'                            # IP shown on OFT / IFT lines with no IP
    BoardDir              = Join-Path $PSScriptRoot 'Boards'    # autosaved boards, one file per day
    TemplateDir           = Join-Path $PSScriptRoot 'Templates' # saved wave templates
    Trace                 = [bool]$Trace                        # print every scheduling decision while filling
}

# Output (the student tabs and MasterData are always written on Export)
$OutputSettings = @{
    WriteAtis              = $true                      # ATIS.xlsx on the Desktop: one tab per day, flights then sims
    KeepPastDays           = 0                          # MasterSchedule keeps exported days from today on (plus this many days back)
    SaveAndCloseExcel      = $true                      # save and close a workbook in Excel before FAIS writes it (no "file is open" errors)
    ReopenExcel            = $true                      # open those workbooks again once FAIS has written them
    CloseAllExcelWorkbooks = $false                     # $true = save and close EVERY open workbook before writing, not just FAIS's
    PublishToTeams         = $false                     # also write the daily sheets to the Teams / OneDrive workbook
    TeamsFolder            = ''                         # e.g. C:\Users\...\OneDrive - United States Air Force\ATIS - E Flight - E Flight
    TeamsFileName          = 'TESTING ATIS NOT OFFICIAL.xlsx'
}


# ==========================================
#      IMPORT MODULES AND FUNCTIONS
# ==========================================
# Makes sure you have excel
try {
    Import-Module ImportExcel -ErrorAction Stop
}
catch {
    Write-Host "Please install or copy the ImportExcel module, then try again." -ForegroundColor Yellow
    return
}

# Order matters: later modules call earlier ones
$ModuleOrder = @('DataManipulation', 'ExcelFiles', 'Syllabus', 'DataDownload', 'SchedulingEngine', 'DataUpload', 'SyllabusImport', 'PreRunMenu', 'BoardModel', 'Assessment', 'PuckBoard')
foreach ($moduleName in $ModuleOrder) {
    $modulePath = Join-Path $PSScriptRoot "Modules\$moduleName.psm1"
    try {
        Import-Module $modulePath -Force -DisableNameChecking -Global -ErrorAction Stop
    }
    catch {
        Write-Host "Could not load $modulePath : $($_.Exception.Message)" -ForegroundColor Red
        return
    }
}


# ==========================================
#       INGEST DATA
# ==========================================
if (-not (Test-Path -LiteralPath $excelInputPath)) {
    Write-Host "MasterData.xlsx was not found at $excelInputPath" -ForegroundColor Red
    return
}
# Workbooks open in Excel are saved first, so edits not saved yet are read too
[void](Close-FaisExcelFiles -Path @($excelInputPath, $excelOutputPath) -SaveOnly)
$Ingested = IngestData -excelInputPath $excelInputPath -excelOutputPath $excelOutputPath
if ($Ingested.OutputError) {
    # Carrying on would start every student from the blank syllabus and overwrite their tabs on export
    Write-Host "OutputSchedule.xlsx could not be opened, so the students' progress can't be read: $($Ingested.OutputError)" -ForegroundColor Red
    Write-Host "Nothing was changed. Fix or close $excelOutputPath, then run FAIS again." -ForegroundColor Red
    return
}
$ScheduleSheet = $Ingested.ScheduleSheet
$Instructors = $Ingested.Instructors
$Students = $Ingested.Students
$Formations = $Ingested.Formations
$SoftWeights = $Ingested.SoftWeights

if ($null -eq $SoftWeights) {
    Write-Host "The SoftWeights sheet could not be read, so nothing can be scored. Fix MasterData.xlsx and run again." -ForegroundColor Red
    return
}
Write-Host ("Loaded {0} students, {1} IPs, {2} formation callsigns from {3}." -f @($Students).Count, @($Instructors).Count, @($Formations).Count, $excelInputPath) -ForegroundColor Cyan


# ==========================================
#      PRE RUN MENU
# ==========================================
# Brings in the previous schedule (with the complex objects attached) so lines can be
# marked hooked (unsat) or incomplete before anything is scheduled
$PrevSchedule = @(DownloadPrevSchedule -Students $Students -Instructors $Instructors -excelOutputPath $excelOutputPath)
# Add Student Absence (the same window as the board's button), on the board's first day
$AddAbsence = { param($prompt) Invoke-FaisAddAbsence -Students $Students -MasterDataPath $excelInputPath -LockedPrompt $prompt }
if (-not (Invoke-PreRunMenu -Schedule $PrevSchedule -Students $Students -SoftWeights $SoftWeights -MasterDataPath $excelInputPath -OutputPath $excelOutputPath -AddAbsence $AddAbsence)) {
    return
}


# ==========================================
#      PUCK BOARD
# ==========================================
# Build the lines, hardcode anyone, then "Fill this day" / "Fill all days" runs the scheduler
# (Invoke-FillSchedule) and puts its picks on the board to change; Export writes the board as it stands
$BoardData = @{
    Students       = $Students
    Instructors    = $Instructors
    SoftWeights    = $SoftWeights
    Formations     = $Formations
    PrevSchedule   = $PrevSchedule
    ScheduleSheet  = $ScheduleSheet
    MasterDataPath = $excelInputPath
    OutputPath     = $excelOutputPath
}
$FirstDay = (Get-NextFlyingDay).ToString('yyyy-MM-dd')
$BoardResult = Show-PuckBoard -Data $BoardData -Days @($FirstDay) -Settings $BoardSettings

if ($null -eq $BoardResult -or $BoardResult.Action -ne 'Export') {
    Write-Host "Puck board closed without exporting - nothing was written. Your lines are saved for next time." -ForegroundColor Yellow
    return
}
$FillResult = $BoardResult.FillResult
if ($Trace) {
    foreach ($entry in $FillResult.Log) { Write-Host "  $entry" -ForegroundColor DarkGray }
}


# =====================================================================================
#   ____ _     _____    _    _   _      _    _   _ ____     ___  _   _ _____ ____  _   _ _____
#  / ___| |   | ____|  / \  | \ | |    / \  | \ | |  _ \   / _ \| | | |_   _|  _ \| | | |_   _|
# | |   | |   |  _|   / _ \ |  \| |   / _ \ |  \| | | | | | | | | | | | | | | |_) | | | | | |
# | |___| |___| |___ / ___ \| |\  |  / ___ \| |\  | |_| | | |_| | |_| | | | |  __/| |_| | | |
#  \____|_____|_____/_/   \_\_| \_| /_/   \_\_| \_|____/   \___/ \___/  |_| |_|    \___/  |_|
# ====================================================================================
# Only shown when FAIS could not save and close the file itself (a cell still being edited,
# a dialog open in Excel, or someone else has it open)
$LockedPrompt = {
    param($path)
    $answer = [System.Windows.MessageBox]::Show("$path is still open and FAIS could not close it (a cell may be in edit mode, a dialog may be open in Excel, or someone else has it open).`n`nClose it, then press OK to try again.", 'File is open',
        [System.Windows.MessageBoxButton]::OKCancel, [System.Windows.MessageBoxImage]::Warning)
    return ($answer -eq [System.Windows.MessageBoxResult]::OK)
}

# Unsat flags the export will change and absences it adds (read before it writes them), and its
# warnings (why it failed, or names MasterData did not have) for the result window
$UnsatChanges = Get-UnsatFlagUpdates -Students $FillResult.Students
$PendingAbsences = @(Get-PendingStudentAbsences -Students $FillResult.Students)
$ExportWarnings = @()
$Exported = Export-FaisOutput -FillResult $FillResult -OutputPath $excelOutputPath -MasterDataPath $excelInputPath -AtisPath $atisOutputPath -Settings $OutputSettings -LockedPrompt $LockedPrompt -Days @($BoardResult.Days) -Instructors $Instructors -WarningVariable ExportWarnings
$AtisWritten = if ($OutputSettings.WriteAtis) { $atisOutputPath } else { '' }
$ExportSummary = Get-FaisExportSummary -Ok ([bool]$Exported) -FillResult $FillResult -Days @($BoardResult.Days) -OutputPath $excelOutputPath -MasterDataPath $excelInputPath `
    -AtisPath $AtisWritten -UnsatChanges $UnsatChanges -Absences $PendingAbsences -Warnings @($ExportWarnings | ForEach-Object { [string]$_.Message })
if (-not $Exported) {
    Write-Host "Nothing was written (OutputSchedule, MasterData and ATIS are as they were). Your board is saved; run FAIS again and Export once the files are closed." -ForegroundColor Red
    try { Show-FaisExportResult -Summary $ExportSummary } catch { Write-Host "Could not show the export result window: $($_.Exception.Message)" -ForegroundColor Yellow }
    return
}

$Stats = $FillResult.Stats
Write-Host ""
Write-Host ("Schedule written to {0} ({1})" -f $excelOutputPath, (@($BoardResult.Days | ForEach-Object { Format-BoardDayLabel $_ }) -join ', ')) -ForegroundColor Green
Write-Host ("  {0} lines: {1} hardcoded, {2} filled by the scheduler, {3} wingmen, {4} return legs, {5} direct support, {6} for sale, {7} NEED SP" -f `
        $Stats.Lines, $Stats.Pinned, $Stats.Filled, $Stats.Wing, $Stats.Return, $Stats.DS, $Stats.ForSale, $Stats.NeedSP) -ForegroundColor Green
foreach ($o in $FillResult.Overrides) {
    Write-Host ("  OVERRIDE {0} {1}: {2}" -f (Format-EtdShort $o.ETD), $o.Student, ($o.Violations -join '; ')) -ForegroundColor Yellow
}
# The same in a window, with buttons that open the files
try { Show-FaisExportResult -Summary $ExportSummary -OutputPath $excelOutputPath -AtisPath $AtisWritten }
catch { Write-Host "Could not show the export result window: $($_.Exception.Message)" -ForegroundColor Yellow }
