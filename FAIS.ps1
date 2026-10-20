#FAIS
# cd C:\Users\1595161645A\Desktop\FAIS7.1
#
# Run order:
#   1. Pre-run pop-ups: log hooked (unsat) / incomplete events from the last schedule
#   2. Puck board: build the day's lines, hardcode IPs and students, see the + / - and rule flags
#   3. "Fill the rest of the lines": the scheduler fills every open line, you review it
#   4. Export: OutputSchedule.xlsx (MasterSchedule + each student's syllabus tab), SimDebt to MasterData,
#      and ATIS.xlsx (one tab per day: flights, then sims)
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
$excelInputPath = Join-Path (Split-Path $PSScriptRoot -Parent) "MasterData.xlsx"
$desktopPath = [Environment]::GetFolderPath('Desktop')
$excelOutputPath = Join-Path -Path $desktopPath -ChildPath "OutputSchedule.xlsx"
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

# Output
$OutputSettings = @{
    WriteStudentTabs = $true                            # each student's syllabus tab in OutputSchedule + SimDebt in MasterData
    WriteAtis        = $true                            # ATIS.xlsx on the Desktop: one tab per day, flights then sims
    PublishToTeams   = $false                           # also write the daily sheets to the Teams / OneDrive workbook
    TeamsFolder      = ''                               # e.g. C:\Users\...\OneDrive - United States Air Force\ATIS - E Flight - E Flight
    TeamsFileName    = 'TESTING ATIS NOT OFFICIAL.xlsx'
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
$ModuleOrder = @('DataManipulation', 'Syllabus', 'DataDownload', 'SchedulingEngine', 'DataUpload', 'PreRunMenu', 'BoardModel', 'Assessment', 'PuckBoard')
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
$Ingested = IngestData -excelInputPath $excelInputPath -desktopPath $desktopPath -excelOutputPath $excelOutputPath
$Schedule = $Ingested.Schedule
$Instructors = $Ingested.Instructors
$Students = $Ingested.Students
$Formations = $Ingested.Formations
$SoftWeights = $Ingested.SoftWeights

if ($null -eq $SoftWeights) {
    Write-Host "The SoftWeights sheet could not be read, so nothing can be scored. Fix MasterData.xlsx and run again." -ForegroundColor Red
    return
}
Write-Host ("Loaded {0} students, {1} IPs, {2} formation callsigns." -f @($Students).Count, @($Instructors).Count, @($Formations).Count) -ForegroundColor Cyan


# ==========================================
#      PRE RUN MENU
# ==========================================
# Brings in the previous schedule (with the complex objects attached) so lines can be
# marked hooked (unsat) or incomplete before anything is scheduled
$PrevSchedule = @(DownloadPrevSchedule -Students $Students -Instructors $Instructors -desktopPath $desktopPath -excelOutputPath $excelOutputPath)
if (-not (Invoke-PreRunMenu -Schedule $PrevSchedule -Students $Students -SoftWeights $SoftWeights -MasterDataPath $excelInputPath)) {
    return
}


# ==========================================
#      PUCK BOARD
# ==========================================
# Build the lines, hardcode anyone, then "Fill the rest of the lines" runs the scheduler
# (Invoke-FillSchedule) and shows the result for review
$BoardData = @{
    Students       = $Students
    Instructors    = $Instructors
    SoftWeights    = $SoftWeights
    Formations     = $Formations
    PrevSchedule   = $PrevSchedule
    ImportSchedule = $Schedule
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
$LockedPrompt = {
    param($path)
    $answer = [System.Windows.MessageBox]::Show("$path is open (probably in Excel).`n`nClose it, then press OK to try again.", 'File is open',
        [System.Windows.MessageBoxButton]::OKCancel, [System.Windows.MessageBoxImage]::Warning)
    return ($answer -eq [System.Windows.MessageBoxResult]::OK)
}

$Exported = Export-FaisOutput -FillResult $FillResult -OutputPath $excelOutputPath -MasterDataPath $excelInputPath -AtisPath $atisOutputPath -Settings $OutputSettings -LockedPrompt $LockedPrompt
if (-not $Exported) {
    Write-Host "The schedule was NOT written. Your board is saved; run FAIS again and Export when the file is closed." -ForegroundColor Red
    return
}

$Stats = $FillResult.Stats
Write-Host ""
Write-Host ("Schedule written to {0}" -f $excelOutputPath) -ForegroundColor Green
Write-Host ("  {0} lines: {1} hardcoded, {2} filled by the scheduler, {3} wingmen, {4} return legs, {5} direct support, {6} for sale" -f `
        $Stats.Lines, $Stats.Pinned, $Stats.Filled, $Stats.Wing, $Stats.Return, $Stats.DS, $Stats.ForSale) -ForegroundColor Green
foreach ($o in $FillResult.Overrides) {
    Write-Host ("  OVERRIDE {0} {1}: {2}" -f (Format-EtdShort $o.ETD), $o.Student, ($o.Violations -join '; ')) -ForegroundColor Yellow
}
