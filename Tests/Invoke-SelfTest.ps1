# FAIS self-test for the Windows machine. Needs no Pester and touches no real files.
#
#   powershell.exe -STA -ExecutionPolicy Bypass -File .\Tests\Invoke-SelfTest.ps1
#
# Checks, under the real Windows PowerShell 5.1:
#   - the scheduling engine, board model and placement checks (synthetic students)
#   - Excel round trip through ImportExcel (temp files only), when ImportExcel is installed
#   - the puck board: the XAML loads, every template builds, and the real window is driven
#     off-screen through pin / move / undo / ETD / fill / export (a window may flash briefly)
# Exit code 0 = everything passed. -SkipWindow skips the window part.
param (
    [switch]$SkipWindow
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Fixtures\New-FaisTestData.ps1')
Import-FaisModules -IncludeUI

$script:Passed = 0
$script:Failed = New-Object System.Collections.Generic.List[string]

# Runs one check; the scriptblock's last output must be $true
function Assert-That {
    param ([string]$Name, [scriptblock]$Test)

    try {
        $out = @(& $Test)
        $ok = ($out.Count -gt 0 -and $out[$out.Count - 1] -eq $true)
    }
    catch {
        $ok = $false
        $Name = "$Name  ($($_.Exception.Message))"
    }
    if ($ok) {
        $script:Passed++
        Write-Host "  PASS  $Name" -ForegroundColor Green
    }
    else {
        $script:Failed.Add($Name)
        Write-Host "  FAIL  $Name" -ForegroundColor Red
    }
}

function Get-Row {
    param ($Result, [string]$Id)
    return ($Result.Schedule | Where-Object { $_.BoardLineId -eq $Id } | Select-Object -First 1)
}

$ips = New-TestInstructors
$sw = New-TestSoftWeights
$forms = New-TestFormations
$formDone = @('FI3101|2026-10-10 08:00', 'FE3102|2026-10-11 08:00', 'FT5001|2026-10-12 08:00|Ortiz, Sam', 'FT5102|2026-10-13 08:00|Ortiz, Sam',
    'FT5103|2026-10-14 08:00|Ortiz, Sam', 'FN3201|2026-10-15 08:00')
$work = Join-Path ([System.IO.Path]::GetTempPath()) ("FAIS_selftest_" + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $work -Force | Out-Null

Write-Host "`nFAIS self-test  (PowerShell $($PSVersionTable.PSVersion), $([System.Threading.Thread]::CurrentThread.GetApartmentState()))" -ForegroundColor Cyan


# ---------------------------------------------------------------------------------
Write-Host "`nTyping lines" -ForegroundColor Cyan
Assert-That 'ETD 730 / 7:30 / 0730L read as 07:30' {
    (ConvertTo-BoardTime '730').TotalMinutes -eq 450 -and (ConvertTo-BoardTime '7:30').TotalMinutes -eq 450 -and (ConvertTo-BoardTime '0730L').TotalMinutes -eq 450
}
Assert-That 'ETD 2500 is rejected' { $null -eq (ConvertTo-BoardTime '2500') }
Assert-That "quick entry '0730 DL, 1015 O&B, 1300 OFT x2' makes 5 lines with a linked return at 1315" {
    $q = ConvertFrom-QuickEntry -Text '0730 DL, 1015 O&B, 1300 OFT x2' -Day '2026-10-20' -Instructors $ips -Students @()
    $ob = $q.Lines | Where-Object { $_.SortieType -eq 'O&B' }
    $ret = $q.Lines | Where-Object { $_.Id -eq $ob.LinkedLineId }
    $q.Errors.Count -eq 0 -and @($q.Lines).Count -eq 5 -and $ret.SortieType -eq 'O&B Return' -and $ret.ETD.ToString('HHmm') -eq '1315'
}
Assert-That "'0900 DL Lee' picks Lee, Ann (not Leech, Tom)" {
    (ConvertFrom-QuickEntry -Text '0900 DL Lee' -Day '2026-10-20' -Instructors $ips -Students @()).Lines[0].Instructor -eq 'Lee, Ann'
}


# ---------------------------------------------------------------------------------
Write-Host "`nHard rules and scoring" -ForegroundColor Cyan
Assert-That 'turn time message' {
    $s = New-TestStudent -Name 'Adams, Al' -Done 'FI3101|2026-10-15 08:00', 'FT5001|2026-10-20 07:30|Jones, Bob'
    $row = New-TestRow -ETD '2026-10-20 09:00' -Instructor 'Jones, Bob'
    $r = Test-HardRules -Student $s -Sortie $row -Students @($s) -Instructors $ips -Explain
    @($r.Violations | ForEach-Object { $_.Message }) -contains 'Turn time 1.5 hr < 3.25 hr minimum (vs 0730 FT5001)'
}
Assert-That '-Explain lists every broken rule (hold + turn time)' {
    $s = New-TestStudent -Name 'Adams, Al' -OnHold -Done 'FI3101|2026-10-15 08:00', 'FT5001|2026-10-20 07:30|Jones, Bob'
    $row = New-TestRow -ETD '2026-10-20 09:00' -Instructor 'Jones, Bob'
    $codes = @((Test-HardRules -Student $s -Sortie $row -Students @($s) -Instructors $ips -Explain).Violations | ForEach-Object { $_.Rule })
    ($codes -contains 'ON_HOLD') -and ($codes -contains 'TURN_TIME')
}
Assert-That 'the rules do not change the student' {
    $s = New-TestStudent -Name 'Adams, Al' -Done 'FI3101|2026-10-15 08:00'
    $before = ConvertTo-Json -InputObject $s -Depth 6
    [void](Test-HardRules -Student $s -Sortie (New-TestRow -ETD '2026-10-20 09:00' -SortieType 'O&B' -Instructor 'Jones, Bob') -Students @($s) -Instructors $ips -Explain)
    (ConvertTo-Json -InputObject $s -Depth 6) -eq $before
}
Assert-That 'a non-counter IP scores Non-Counter' {
    $s = New-TestStudent -Name 'Adams, Al' -Done 'FI3101|2026-10-15 08:00', 'FT5001|2026-10-16 08:00|Jones, Bob'
    $row = New-TestRow -ETD '2026-10-20 09:00' -Instructor 'Leech, Tom'
    $row.Instructor = Resolve-SortieInstructor -Instructors $ips -Instructor $row.Instructor
    $row.Mission = Find-SyllabusMission -Student $s -Mission 'FT5102'
    (SoftRules -Student $s -Config $sw -Sortie $row -Instructors $ips -Detailed).Breakdown['Non-Counter'] -eq -15
}


# ---------------------------------------------------------------------------------
Write-Host "`nFilling the schedule" -ForegroundColor Cyan
$students = @(
    New-TestStudent -Name 'Adams, Al' -Done $formDone
    New-TestStudent -Name 'Brown, Bea' -Done $formDone
    New-TestStudent -Name 'Cruz, Cy' -Done 'FI3101|2026-10-15 08:00'
)
$rows = @(
    New-TestRow -ETD '2026-10-20 07:30' -Instructor 'Jones, Bob' -Id 'A' -Student 'Adams, Al' -Mission 'FF5201'
    New-TestRow -ETD '2026-10-20 07:40' -Instructor 'Lee, Ann' -Id 'B'
    New-TestRow -ETD '2026-10-20 07:45' -SortieType 'IFT' -Id 'S'
    New-TestRow -ETD '2026-10-20 10:15' -SortieType 'O&B' -Instructor 'Ortiz, Sam' -Id 'O' -LinkedLineId 'R'
    New-TestRow -ETD '2026-10-20 13:15' -SortieType 'O&B Return' -Instructor 'Ortiz, Sam' -Id 'R' -LinkedLineId 'O'
    New-TestRow -ETD '2026-10-20 15:00' -Instructor 'Jones, Bob' -Id 'Z'
)
$copyA = @(Copy-StudentList $students)
$copyB = @(Copy-StudentList $students)
$fill = Invoke-FillSchedule -Schedule @($rows | ForEach-Object { $_.PSObject.Copy() }) -Students $copyA -Instructors $ips -SoftWeights $sw -Formations $forms 6>$null
$fill2 = Invoke-FillSchedule -Schedule @($rows | ForEach-Object { $_.PSObject.Copy() }) -Students $copyB -Instructors $ips -SoftWeights $sw -Formations $forms 6>$null

Assert-That 'the hardcoded formation lead stays and gets HAZE 11' { (Get-NameText (Get-Row $fill 'A').Student) -eq 'Adams, Al' -and (Get-Row $fill 'A').Num -eq 11 }
Assert-That 'a Day Local line becomes the wingman (HAZE 12)' { (Get-Row $fill 'B').FilledBy -eq 'Wing' -and (Get-Row $fill 'B').Num -eq 12 }
Assert-That 'the IFT line is never the wingman' { (Get-Row $fill 'S').FilledBy -notin @('Wing', 'DS') }
Assert-That 'the O&B return leg follows its O&B' {
    $o = Get-Row $fill 'O'; $r = Get-Row $fill 'R'
    ($o.FilledBy -eq 'ForSale' -and $r.FilledBy -eq 'ForSale') -or ((Get-NameText $r.Student) -eq (Get-NameText $o.Student) -and $r.FilledBy -eq 'Return')
}
Assert-That 'picked students are marked complete (no double booking)' {
    $names = @($fill.Schedule | Where-Object { $_.FilledBy -in @('Pin', 'Engine', 'Wing') } | ForEach-Object { "$(Get-NameText $_.Student)|$(Get-NameText $_.Mission)" })
    @($names | Select-Object -Unique).Count -eq $names.Count
}
Assert-That 'two runs give the same schedule' {
    $a = @($fill.Schedule | ForEach-Object { "$($_.BoardLineId)|$(Get-NameText $_.Student)|$(Get-NameText $_.Mission)|$($_.Notes)" }) -join "`n"
    $b = @($fill2.Schedule | ForEach-Object { "$($_.BoardLineId)|$(Get-NameText $_.Student)|$(Get-NameText $_.Mission)|$($_.Notes)" }) -join "`n"
    $a -eq $b
}
Assert-That 'a pin that breaks a rule is kept and flagged OVERRIDE' {
    $s = @(New-TestStudent -Name 'Adams, Al' -OnHold -Done 'FI3101|2026-10-15 08:00')
    $r = Invoke-FillSchedule -Schedule @(New-TestRow -ETD '2026-10-20 08:00' -Instructor 'Jones, Bob' -Id 'P' -Student 'Adams, Al') -Students $s -Instructors $ips -SoftWeights $sw 6>$null
    (Get-Row $r 'P').Notes -like '*OVERRIDE: Student is on hold*'
}


# ---------------------------------------------------------------------------------
Write-Host "`nBoard and placement checks" -ForegroundColor Cyan
$board = New-Board -Days '2026-10-20'
Add-BoardLine -Board $board -Line (ConvertFrom-QuickEntry -Text '0730 DL Jones, 0900 DL Lee, 1300 IFT' -Day '2026-10-20' -Instructors $ips -Students @()).Lines
$bl = @(Get-BoardDayLines -Board $board -Day '2026-10-20')
[void](Set-BoardPin -Board $board -Id $bl[0].Id -Student 'Adams, Al')
$data = @{ Students = $students; Instructors = $ips; SoftWeights = $sw; Formations = $forms }
Assert-That 'a student 1.5 hr after their own pin is flagged red' {
    $a = Get-PlacementAssessment -Board $board -LineId $bl[1].Id -StudentName 'Adams, Al' -Data $data
    $a.Status -eq 'Hard' -and @($a.Violations | Where-Object { $_.Rule -eq 'TURN_TIME' }).Count -eq 1
}
Assert-That 'the pop-up text has + / - reasons' {
    $a = Get-PlacementAssessment -Board $board -LineId $bl[0].Id -StudentName 'Adams, Al' -Data $data
    @(Get-AssessmentLines $a | Where-Object { $_.Kind -eq 'Plus' -or $_.Kind -eq 'Minus' }).Count -gt 0
}
Assert-That 'IFT lines cannot be paired as a wingman' { -not (Set-FormationPair -Board $board -WingId $bl[2].Id -LeadId $bl[0].Id) }
Assert-That 'undo puts the board back' {
    $n = $board.Lines.Count
    Remove-BoardLine -Board $board -Id $bl[1].Id
    $removed = ($board.Lines.Count -eq $n - 1)
    [void](Undo-Board $board)
    $removed -and $board.Lines.Count -eq $n
}
Assert-That 'the board saves and reloads exactly' {
    Save-BoardState -Board $board -Dir $work
    $back = Open-BoardState -Dir $work -Days @('2026-10-20')
    $a = (ConvertTo-ScheduleRows -Board $board | ConvertTo-Json -Depth 4)
    $b = (ConvertTo-ScheduleRows -Board $back | ConvertTo-Json -Depth 4)
    $a -eq $b
}


# ---------------------------------------------------------------------------------
Write-Host "`nExcel round trip (temp files)" -ForegroundColor Cyan
if (Get-Module -ListAvailable -Name ImportExcel) {
    Import-Module ImportExcel
    $xlsx = Join-Path $work 'OutputSchedule.xlsx'
    Assert-That 'MasterSchedule + student tabs are written' {
        Export-FormattedMasterSchedule -ScheduleData (ConvertTo-ExportRows -Schedule $fill.Schedule) -FilePath $xlsx -WorksheetName 'MasterSchedule' 6>$null
        Export-StudentSyllabusTabs -Students $copyA -FilePath $xlsx 6>$null
        Set-MasterScheduleFirst -FilePath $xlsx 6>$null
        Test-Path $xlsx
    }
    Assert-That "a student's tab reads back with the same completions" {
        $tab = @(Import-Excel -Path $xlsx -WorksheetName 'Adams, Al')
        $done = @($tab | Where-Object { ConvertTo-FaisBool $_.IsComplete } | ForEach-Object { $_.Mission })
        $want = @($copyA[0].Syllabus | Where-Object { Test-MissionComplete $_ } | ForEach-Object { $_.Mission })
        ($done -join ',') -eq ($want -join ',') -and $null -ne $tab[0].PSObject.Properties['OptingEvents']
    }
    Assert-That 'MasterSchedule reads back for the hook / incomplete menu' {
        $prev = @(DownloadPrevSchedule -Students $copyA -Instructors $ips -excelOutputPath $xlsx)
        $prev.Count -eq @($fill.Schedule).Count -and @($prev | Where-Object { $_.Student -isnot [string] -and $null -ne $_.Student }).Count -gt 0
    }
    Assert-That 'ATIS has a tab for the day: flights by ETD, then the IFT under a thick border' {
        $atis = Join-Path $work 'ATIS.xlsx'
        Export-AtisWorkbook -ScheduleData (ConvertTo-ExportRows -Schedule $fill.Schedule) -FilePath $atis 6>$null
        $sheet = Get-DaySheetName -Date ([datetime]'2026-10-20')
        $types = @(Import-Excel -Path $atis -WorksheetName $sheet | ForEach-Object { $_.SortieType })
        $pkg = Open-ExcelPackage -Path $atis
        $simRow = [array]::IndexOf($types, 'IFT') + 2
        $thick = $pkg.Workbook.Worksheets[$sheet].Cells[$simRow, 1].Style.Border.Top.Style -eq [OfficeOpenXml.Style.ExcelBorderStyle]::Thick
        Close-ExcelPackage -ExcelPackage $pkg -NoSave
        $types.Count -eq @($fill.Schedule).Count -and $types[$types.Count - 1] -eq 'IFT' -and $thick
    }
}
else {
    Write-Host "  SKIP  ImportExcel is not installed" -ForegroundColor Yellow
}


# ---------------------------------------------------------------------------------
Write-Host "`nPuck board window" -ForegroundColor Cyan
if ($SkipWindow) {
    Write-Host "  SKIP  -SkipWindow" -ForegroundColor Yellow
}
elseif ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
    $script:Failed.Add('window checks need STA: run powershell.exe -STA -File .\Tests\Invoke-SelfTest.ps1')
    Write-Host "  FAIL  window checks need STA: run powershell.exe -STA -File .\Tests\Invoke-SelfTest.ps1" -ForegroundColor Red
}
else {
    Initialize-PuckBoardAssemblies
    Assert-That 'PuckBoard.xaml loads and every template builds' {
        $root = Split-Path $PSScriptRoot -Parent
        [xml]$x = Get-Content -LiteralPath (Join-Path $root 'Modules\PuckBoard.xaml') -Raw -Encoding UTF8
        $w = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $x))
        $built = foreach ($k in @('LineRowTemplate', 'IPPuckTemplate', 'StudentPuckTemplate', 'LineHeaderTemplate')) { $null -ne $w.FindResource($k).LoadContent() }
        @($built | Where-Object { -not $_ }).Count -eq 0
    }
    $windowData = @{
        Students       = @(Copy-StudentList $students)
        Instructors    = $ips
        SoftWeights    = $sw
        Formations     = $forms
        PrevSchedule   = @()
        ImportSchedule = @()
    }
    $problems = @(Test-PuckBoardWindow -Data $windowData -Day '2026-10-21' -WorkDir $work)
    foreach ($p in $problems) { $script:Failed.Add("window: $p"); Write-Host "  FAIL  window: $p" -ForegroundColor Red }
    if ($problems.Count -eq 0) { $script:Passed++; Write-Host "  PASS  the board window ran pin / move / undo / ETD / type / mission / notes / checks / fill / export with no binding errors" -ForegroundColor Green }
}


# ---------------------------------------------------------------------------------
try { Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue } catch { }
Write-Host ""
if ($script:Failed.Count -eq 0) {
    Write-Host "ALL $($script:Passed) CHECKS PASSED" -ForegroundColor Green
    exit 0
}
Write-Host "$($script:Failed.Count) FAILED, $($script:Passed) passed:" -ForegroundColor Red
foreach ($f in $script:Failed) { Write-Host "  - $f" -ForegroundColor Red }
Write-Host "Board errors are also logged to $env:TEMP\FAIS_board.log" -ForegroundColor Yellow
exit 1
