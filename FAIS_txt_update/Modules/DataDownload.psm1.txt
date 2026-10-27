# DATADOWNLOAD.PSM1 DIRECTORY:
# Contains all functions that download data to use in the program
# IngestData            (MasterData.xlsx + each student's tab in OutputSchedule.xlsx)
# Open-FaisWorkbookForRead / Close-FaisWorkbookForRead (read a workbook even while Excel has it open)
# Read-ScheduleSheet / Get-ScheduleSheetDays / Get-ScheduleDayFingerprint (MasterData Schedule sheet)
# DownloadPrevSchedule  (previous MasterSchedule with students/missions/IPs linked back up)
# Import-MasterSchedule (same sheet through Excel COM, strings only)


# ========================================================================================
#  ___                       _   ____        _
# |_ _|_ __   __ _  ___  ___| |_|  _ \  __ _| |_ __ _
#  | || '_ \ / _` |/ _ \/ __| __| | | |/ _` | __/ _` |
#  | || | | | (_| |  __/\__ \ |_| |_| | (_| | || (_| |
# |___|_| |_|\__, |\___||___/\__|____/ \__,_|\__\__,_|
#            |___/
# ========================================================================================
# Reads MasterData.xlsx and OutputSchedule.xlsx and returns one object with:
# Schedule, ScheduleSheet, Instructors, Students, Formations, SoftWeights, StudentAbsences
# The Schedule sheet is optional: the puck board loads it one day at a time.
function IngestData {
                param (
                                [string]$excelInputPath,
                                [string]$excelOutputPath,
                                [int]$MaxStudents = 100
                )

                # Default to OutputSchedule.xlsx in the FAIS folder when no path was passed in
                if ([string]::IsNullOrWhiteSpace($excelOutputPath)) {
                                $excelOutputPath = Join-Path (Split-Path $PSScriptRoot -Parent) "OutputSchedule.xlsx"
                }

                # ------------------------------------------------------------------
                # Instructors (blank IsCounter counts as a counter)
                # ------------------------------------------------------------------
                $excelData = Import-Excel -Path $excelInputPath -WorksheetName "Instructors"
                $Instructors = @(
                                foreach ($IP in $excelData) {
                                                if ([string]::IsNullOrWhiteSpace([string]$IP.IPName)) { continue }
                                                [PSCustomObject]@{
                                                                Name          = ([string]$IP.IPName).Trim()
                                                                IsCounter     = (ConvertTo-FaisBool $IP.IsCounter -Default $true)
                                                                IsExperienced = (ConvertTo-FaisBool $IP.IsExperienced)
                                                                IsCheck       = (ConvertTo-FaisBool $IP.IsCheck)
                                                                Callsign      = $IP.Callsign
                                                                Num           = $IP.Num
                                                                OSC           = $IP.OSC
                                                                OSNum         = $IP.OSNum
                                                }
                                }
                )

                # ------------------------------------------------------------------
                # Students (prefer each student's tab in OutputSchedule, else master Syllabus)
                # ------------------------------------------------------------------
                $rawStudents = Import-Excel -Path $excelInputPath -WorksheetName "Students"
                $named = @($rawStudents | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.SPName) })
                if ($named.Count -gt $MaxStudents) {
                                Write-Host "WARNING: Students sheet has $($named.Count) students; only the first $MaxStudents are used." -ForegroundColor Yellow
                }
                $validStudents = @($named | Select-Object -First $MaxStudents)

                # OutputSchedule is opened once for every student's tab (it is read even while open in Excel)
                $masterSyllabus = $null
                $outputBook = Open-FaisWorkbookForRead -Path $excelOutputPath
                $Students = @(
                                foreach ($Student in $validStudents) {
                                                $studentSyllabus = $null

                                                # Prefer the student's own tab in OutputSchedule when it has one
                                                $tabName = ([string]$Student.SPName).Trim()
                                                if ($null -ne $outputBook.Package -and $null -ne $outputBook.Package.Workbook.Worksheets[$tabName]) {
                                                                try {
                                                                                $studentSyllabus = @(Import-Excel -ExcelPackage $outputBook.Package -WorksheetName $tabName -ErrorAction Stop)
                                                                }
                                                                catch {
                                                                                Write-Host "WARNING: Could not read $tabName's tab in OutputSchedule ($($_.Exception.Message)). Using the master Syllabus." -ForegroundColor Yellow
                                                                }
                                                }
                                                elseif ($null -ne $outputBook.Package) {
                                                                Write-Host "No tab for $tabName in OutputSchedule yet: starting from the master Syllabus." -ForegroundColor DarkGray
                                                }

                                                # Fall back to the master Syllabus sheet when no student-specific tab was found
                                                if ($null -eq $studentSyllabus) {
                                                                if ($null -eq $masterSyllabus) { $masterSyllabus = @(Import-Excel -Path $excelInputPath -WorksheetName "Syllabus") }
                                                                $studentSyllabus = $masterSyllabus
                                                }

                                                # Normalize each mission into the standard syllabus object
                                                $cleanSyllabus = @(
                                                                foreach ($mission in $studentSyllabus) {
                                                                                if ([string]::IsNullOrWhiteSpace([string]$mission.Mission)) { continue }
                                                                                $etd = Parse-ETD $mission.ETD
                                                                                [PSCustomObject]@{
                                                                                                SyllabusID = $mission.SyllabusID
                                                                                                Mission    = ([string]$mission.Mission).Trim()
                                                                                                Prereqs    = $mission.Prereqs
                                                                                                Instructor = if ($null -eq $mission.Instructor) { '' } else { [string]$mission.Instructor }
                                                                                                ETD        = if ($null -ne $etd) { $etd.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
                                                                                                Type       = if ($null -eq $mission.Type) { '' } else { ([string]$mission.Type).Trim() }
                                                                                                IsCheck    = (ConvertTo-FaisBool $mission.IsCheck)
                                                                                                IsComplete = (ConvertTo-FaisBool $mission.IsComplete)
                                                                                                NumOpting  = $mission.OptingEvents
                                                                                                Level      = $mission.Level
                                                                                }
                                                                }
                                                )

                                                $debt = 0.0
                                                if ($null -ne $Student.SimDebt -and [string]$Student.SimDebt -ne '') { [void][double]::TryParse([string]$Student.SimDebt, [ref]$debt) }
                                                # Blank (or no DiscontinuityScore column): counted from the syllabus once SoftWeights is read
                                                $disc = $null
                                                $discNum = 0.0
                                                if ($null -ne $Student.DiscontinuityScore -and [string]$Student.DiscontinuityScore -ne '' -and
                                                                [double]::TryParse([string]$Student.DiscontinuityScore, [ref]$discNum)) { $disc = $discNum }
                                                $priority = 0.0
                                                if ($null -ne $Student.Priority -and [string]$Student.Priority -ne '') { [void][double]::TryParse([string]$Student.Priority, [ref]$priority) }

                                                # MasterData only says TRUE/FALSE, so the hooked sim is taken to be the latest one flown
                                                $unsatSim = ConvertTo-FaisBool $Student.UnsatSim
                                                $unsatFlight = ConvertTo-FaisBool $Student.UnsatFlight
                                                $onHold = ConvertTo-FaisBool $Student.OnHold
                                                $unsatSimType = if ($unsatSim) { Get-LastSimType -Student ([PSCustomObject]@{ Syllabus = $cleanSyllabus }) } else { '' }

                                                [PSCustomObject]@{
                                                                ID                      = $Student.SPKey
                                                                Name                    = ([string]$Student.SPName).Trim()
                                                                IsCap                   = (ConvertTo-FaisBool $Student.IsCap)
                                                                IsInternational         = (ConvertTo-FaisBool $Student.IsInternational)
                                                                Priority                = $priority
                                                                UnsatSim                = $unsatSim
                                                                UnsatSimType            = $unsatSimType
                                                                UnsatFlight             = $unsatFlight
                                                                # What MasterData holds now: Export writes UnsatSim / UnsatFlight / OnHold back only when they differ
                                                                UnsatSimInMasterData    = $unsatSim
                                                                UnsatFlightInMasterData = $unsatFlight
                                                                OnHold                  = $onHold
                                                                OnHoldInMasterData      = $onHold
                                                                SimDebt                 = $debt
                                                                DiscontinuityScore      = $disc
                                                                Syllabus                = $cleanSyllabus
                                                                Absences                = @()
                                                                Reservations            = @()
                                                                SyllabusRevision        = 0
                                                }
                                }
                )
                Close-FaisWorkbookForRead $outputBook

                # ------------------------------------------------------------------
                # Schedule sheet (optional: the puck board loads it one day at a time)
                # ------------------------------------------------------------------
                $ScheduleSheet = Read-ScheduleSheet -Path $excelInputPath
                if ($ScheduleSheet.Error) { Write-Host "Schedule sheet: $($ScheduleSheet.Error)" -ForegroundColor Yellow }
                elseif (@($ScheduleSheet.Rows).Count -gt 0) {
                                $dayText = @(Get-ScheduleSheetDays -Rows $ScheduleSheet.Rows | ForEach-Object { "$($_.Label) ($($_.Count))" }) -join ', '
                                Write-Host "Schedule sheet: $(@($ScheduleSheet.Rows).Count) line(s): $dayText" -ForegroundColor Cyan
                }
                elseif ($ScheduleSheet.Found) { Write-Host "Schedule sheet: no lines" -ForegroundColor Yellow }
                else { Write-Host "Schedule sheet: MasterData has no sheet named 'Schedule'" -ForegroundColor Yellow }
                foreach ($skip in @($ScheduleSheet.Skipped)) { Write-Host "  Schedule sheet $skip" -ForegroundColor Yellow }
                $Schedule = @($ScheduleSheet.Rows | Where-Object { $_.ETD } | Sort-Object -Property ETD)

                # ------------------------------------------------------------------
                # Formations
                # ------------------------------------------------------------------
                $Formations = @()
                try {
                                $formData = Import-Excel -Path $excelInputPath -WorksheetName "FormCallsigns" -ErrorAction Stop
                                $index = 1
                                $Formations = @(
                                                foreach ($row in $formData) {
                                                                if ([string]::IsNullOrWhiteSpace([string]$row.Callsign)) { continue }
                                                                [PSCustomObject]@{
                                                                                NumFormation = $index
                                                                                Callsign     = $row.Callsign
                                                                                Num          = $row.Num
                                                                }
                                                                $index++
                                                }
                                )
                }
                catch {
                                Write-Host "Could not read the 'FormCallsigns' sheet: formation lines will not get formation callsigns." -ForegroundColor Yellow
                }

                # ------------------------------------------------------------------
                # Soft Weights (first data row of the SoftWeights sheet)
                # ------------------------------------------------------------------
                $SoftWeights = $null
                try {
                                $RawWeights = @(Import-Excel -Path $excelInputPath -WorksheetName "SoftWeights" -ErrorAction Stop)
                                $weightRow = $RawWeights[0]
                                $SoftWeights = [PSCustomObject]@{
                                                ContString        = $weightRow.ContString
                                                String4           = $weightRow.String4
                                                String5           = $weightRow.String5
                                                Zipper            = $weightRow.Zipper
                                                NonCounter        = $weightRow.NonCounter
                                                BreakString       = $weightRow.BreakString
                                                ExhaustionMult    = $weightRow.ExhaustMult
                                                ExhaustionDays    = $weightRow.ExhaustDays
                                                BIT_X             = $weightRow.BITPerDay
                                                BIT_Y             = $weightRow.BITDays
                                                CAPBonus          = $weightRow.CAPBonus
                                                FormLookAhead     = $weightRow.FormLookAhead
                                                LevelDifference   = $weightRow.LevelDifference
                                                CounterContinuity = $weightRow.CounterContinuity
                                                # Optional columns: a blank or missing one uses the default in SoftRules
                                                Alternating       = $weightRow.Alternating
                                                NewCounter        = $weightRow.NewCounter
                                                OFTIFTPen         = $weightRow.OFTIFTPen
                                                DollarRideBonus   = $weightRow.DollarRideBonus
                                                TransSimRatio     = $weightRow.TransSimRatio
                                                FormSimRatio      = $weightRow.FormSimRatio
                                                SimStopMult       = $weightRow.SimStopMult
                                                Disc1String       = $weightRow.Disc1String
                                                DiscCounter       = $weightRow.DiscCounter
                                                Disc3PlusString   = $weightRow.Disc3PlusString
                                }
                                $SoftWeights.TransSimRatio = [double]$SoftWeights.TransSimRatio * [double]$SoftWeights.SimStopMult
                                $SoftWeights.FormSimRatio = [double]$SoftWeights.FormSimRatio * [double]$SoftWeights.SimStopMult
                }
                catch {
                                Write-Host "An error occurred while loading Soft Weights: $($_.Exception.Message)" -ForegroundColor Yellow
                }
                $counted = @(Initialize-StudentDiscontinuity -Students $Students -Instructors $Instructors -SoftWeights $SoftWeights)
                if ($counted.Count -gt 0) {
                                Write-Host "DiscontinuityScore is blank in MasterData for $($counted.Count) student(s): counted from their syllabus (Export writes it)." -ForegroundColor DarkGray
                }

                # ------------------------------------------------------------------
                # Student Absences (active rows only, dates stored as DateTime)
                # ------------------------------------------------------------------
                $StudentAbsences = @()
                try {
                                $AbsenceData = Import-Excel -Path $excelInputPath -WorksheetName "StudentAbsences" -ErrorAction Stop
                                $StudentAbsences = @(
                                                foreach ($row in $AbsenceData) {
                                                                $studentName = if ($row.Student) { ([string]$row.Student).Trim() } else { "" }
                                                                if ([string]::IsNullOrWhiteSpace($studentName)) { continue }
                                                                if (-not (ConvertTo-FaisBool $row.IsActive)) { continue }
                                                                $start = Parse-ETD $row.StartTime
                                                                $end = Parse-ETD $row.EndTime
                                                                if ($null -eq $start -or $null -eq $end) { continue }
                                                                [PSCustomObject]@{
                                                                                Student   = $studentName
                                                                                StartTime = $start
                                                                                EndTime   = $end
                                                                                Reason    = if ($row.Reason) { ([string]$row.Reason).Trim() } else { "" }
                                                                }
                                                }
                                )
                }
                catch {
                                Write-Host "Could not read the 'StudentAbsences' sheet." -ForegroundColor Yellow
                                $StudentAbsences = @()
                }

                # Attach each absence to the student it belongs to so the hard rules can read it off the student object.
                # Names must match exactly (case and outer spaces aside).
                foreach ($s in $Students) {
                                $studentName = if ($s.Name) { $s.Name.Trim() } else { "" }
                                if ([string]::IsNullOrWhiteSpace($studentName)) { continue }

                                $s.Absences = @(
                                                foreach ($absence in $StudentAbsences) {
                                                                if ($absence.Student.Trim() -ieq $studentName) { $absence }
                                                }
                                )
                }
                $knownNames = @($Students | ForEach-Object { ([string]$_.Name).Trim() })
                foreach ($name in @($StudentAbsences | ForEach-Object { $_.Student.Trim() } | Sort-Object -Unique)) {
                                if ($knownNames -notcontains $name) {
                                                Write-Host "WARNING: The absence for '$name' doesn't match a student name on the Students sheet exactly, so it was not applied." -ForegroundColor Yellow
                                }
                }

                return [PSCustomObject]@{
                                # Set when OutputSchedule exists but could not be opened: the students' progress is unknown
                                OutputError     = $outputBook.Error
                                Schedule        = $Schedule
                                ScheduleSheet   = $ScheduleSheet
                                Instructors     = $Instructors
                                Students        = $Students
                                Formations      = $Formations
                                SoftWeights     = $SoftWeights
                                StudentAbsences = $StudentAbsences
                }
}


# =================================================================================
#  READING A WORKBOOK THAT MAY BE OPEN IN EXCEL
# =================================================================================

# Opens a workbook for reading even while Excel has it open (it reads the last saved copy).
# Returns {Package; Stream; Error}. Package is $null when the file is missing or unreadable.
function Open-FaisWorkbookForRead {
                param ([string]$Path)

                $book = [PSCustomObject]@{ Package = $null; Stream = $null; Error = '' }
                if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { return $book }
                try {
                                $book.Stream = New-Object System.IO.FileStream -ArgumentList $Path, 'Open', 'Read', 'ReadWrite'
                                $pkg = New-Object OfficeOpenXml.ExcelPackage
                                $pkg.Load($book.Stream)
                                $book.Package = $pkg
                } catch {
                                $book.Error = $_.Exception.Message
                                Write-Host "WARNING: Could not open $Path : $($book.Error)" -ForegroundColor Yellow
                                Close-FaisWorkbookForRead $book
                }
                return $book
}

function Close-FaisWorkbookForRead {
                param ($Book)

                if ($null -eq $Book) { return }
                if ($null -ne $Book.Package) { try { $Book.Package.Dispose() } catch { } ; $Book.Package = $null }
                if ($null -ne $Book.Stream) { try { $Book.Stream.Close() } catch { } ; $Book.Stream = $null }
}


# =================================================================================
#  ____       _              _       _        ____  _               _
# / ___|  ___| |__   ___  __| |_   _| | ___  / ___|| |__   ___  ___| |_
# \___ \ / __| '_ \ / _ \/ _` | | | | |/ _ \ \___ \| '_ \ / _ \/ _ \ __|
#  ___) | (__| | | |  __/ (_| | |_| | |  __/  ___) | | | |  __/  __/ |_
# |____/ \___|_| |_|\___|\__,_|\__,_|_|\___| |____/|_| |_|\___|\___|\__|
#
# The MasterData Schedule sheet: the lines to fly, one row each. The puck board loads it
# one day at a time and reloads it whenever MasterData is saved.
#  - Columns are found by their headings (ETD / Time, Date, Instructor / IP, SortieType / Type,
#    Student / SP, Mission / Event, Callsign, Num / #, Notes / Remarks), in any order.
#    With no heading row, the legacy order is used: ETD, Instructor, SortieType, Student,
#    Mission, Callsign, Num, Notes.
#  - Blank rows and repeated heading rows are skipped (they no longer end the sheet).
#  - A row with only a date on it ("Tuesday, October 21, 2026") dates the rows under it.
#  - ETDs can be full dates or times only (0730, 730, 7:30, a time-formatted cell). A time
#    only row takes its day from a Date column or the date row above it; with neither it
#    goes onto whichever day it is loaded onto.
# =================================================================================

$script:ScheduleHeadings = @{
                # heading (letters and digits only, lower case) = column it fills
                Primary   = @{
                                etd = 'ETD'; starttime = 'ETD'; date = 'Date'; instructor = 'Instructor'; ip = 'Instructor'
                                sortietype = 'SortieType'; type = 'SortieType'; student = 'Student'; sp = 'Student'; spname = 'Student'
                                mission = 'Mission'; callsign = 'Callsign'; num = 'Num'; notes = 'Notes'; remarks = 'Notes'
                }
                # used only when no primary heading claimed that column
                Secondary = @{
                                time = 'ETD'; takeoff = 'ETD'; takeofftime = 'ETD'; etdlocal = 'ETD'; day = 'Date'; flydate = 'Date'
                                ipname = 'Instructor'; simulatorinstructor = 'Instructor'; simip = 'Instructor'; device = 'SortieType'
                                linetype = 'SortieType'; sortie = 'SortieType'; studentname = 'Student'; event = 'Mission'
                                eventname = 'Mission'; msn = 'Mission'; cs = 'Callsign'; number = 'Num'; note = 'Notes'
                                remark = 'Notes'; comments = 'Notes'; comment = 'Notes'
                }
}
$script:ScheduleLegacyColumns = @('ETD', 'Instructor', 'SortieType', 'Student', 'Mission', 'Callsign', 'Num', 'Notes')

# Which column each heading names: field -> column number (1-based), from one row's cells
function Get-ScheduleHeadingMap {
                param ([array]$Cells)

                $map = @{}
                foreach ($pass in @('Primary', 'Secondary')) {
                                $names = $script:ScheduleHeadings[$pass]
                                for ($c = 0; $c -lt $Cells.Count; $c++) {
                                                $raw = ([string]$Cells[$c]).Trim()
                                                if ($raw -eq '') { continue }
                                                $key = if ($raw -eq '#') { 'num' } else { ($raw -replace '[^a-zA-Z0-9]', '').ToLowerInvariant() }
                                                if (-not $names.ContainsKey($key)) { continue }
                                                $field = $names[$key]
                                                if (-not $map.ContainsKey($field) -and ($map.Values -notcontains ($c + 1))) { $map[$field] = $c + 1 }
                                }
                }
                return $map
}

# A heading row names the ETD column and at least one other, or any three columns
function Test-ScheduleHeadingRow {
                param ([array]$Cells)

                $map = Get-ScheduleHeadingMap $Cells
                return (($map.ContainsKey('ETD') -and $map.Count -ge 2) -or $map.Count -ge 3)
}

function ConvertTo-ScheduleCellText {
                param ($Value)

                if ($null -eq $Value) { return '' }
                if ($Value -is [double] -and $Value -eq [math]::Floor($Value) -and [math]::Abs($Value) -lt 1e9) { return ([long]$Value).ToString() }
                return ([string]$Value).Trim()
}

# Reads the Schedule sheet. Returns
#   {Path; Worksheet; Found; Error; Layout; HeaderRow; Rows; Skipped; ReadAt}
# Each row: {ETD ('yyyy-MM-dd HH:mm:ss', $null for a time with no day); Day ('yyyy-MM-dd' or '');
#            Time ('HH:mm'); TimeOnly; Instructor; SortieType; Student; Mission; Callsign; Num; Notes; SheetRow}
# Skipped holds one 'row N: why' line for each row that looked like a line but could not be read.
function Read-ScheduleSheet {
                param (
                                [string]$Path,
                                [string]$WorksheetName = 'Schedule',
                                [string]$ReturnType = 'O&B Return',
                                [string]$SimInstructor = 'SIM IP'
                )

                $result = [PSCustomObject]@{
                                Path      = $Path
                                Worksheet = $WorksheetName
                                Found     = $false
                                Error     = ''
                                Layout    = ''
                                HeaderRow = 0
                                Rows      = @()
                                Skipped   = @()
                                ReadAt    = (Get-Date)
                }
                if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
                                $result.Error = "$Path was not found."
                                return $result
                }

                $raw = @()
                $book = Open-FaisWorkbookForRead -Path $Path
                try {
                                if ($null -eq $book.Package) { $result.Error = "Could not open $Path : $($book.Error)"; return $result }
                                $ws = $book.Package.Workbook.Worksheets[$WorksheetName]
                                if ($null -eq $ws) { return $result }
                                $result.Found = $true
                                if ($null -eq $ws.Dimension) { return $result }
                                $raw = @(Import-Excel -ExcelPackage $book.Package -WorksheetName $WorksheetName -NoHeader -ErrorAction Stop)
                } catch {
                                $result.Error = "Could not read the '$WorksheetName' sheet: $($_.Exception.Message)"
                                return $result
                } finally {
                                Close-FaisWorkbookForRead $book
                }

                $cellsOf = {
                                param($row)
                                @(foreach ($p in $row.PSObject.Properties) { $p.Value })
                }

                # Heading row: the first row (of the first 15 with anything in them) that names the columns
                $map = $null
                $seen = 0
                for ($i = 0; $i -lt $raw.Count -and $seen -lt 15; $i++) {
                                $cells = & $cellsOf $raw[$i]
                                if (@($cells | Where-Object { ([string]$_).Trim() -ne '' }).Count -eq 0) { continue }
                                $seen++
                                if (Test-ScheduleHeadingRow $cells) {
                                                $map = Get-ScheduleHeadingMap $cells
                                                $result.HeaderRow = $i + 1
                                                break
                                }
                }
                if ($null -eq $map) {
                                $map = @{}
                                for ($c = 0; $c -lt $script:ScheduleLegacyColumns.Count; $c++) { $map[$script:ScheduleLegacyColumns[$c]] = $c + 1 }
                                $result.Layout = 'Fixed columns (ETD, Instructor, SortieType, Student, Mission, Callsign, Num, Notes)'
                }
                else {
                                $result.Layout = 'Headings on row ' + $result.HeaderRow
                                if (-not $map.ContainsKey('ETD') -or -not $map.ContainsKey('SortieType')) {
                                                $missing = @(@('ETD', 'SortieType') | Where-Object { -not $map.ContainsKey($_) }) -join ' and '
                                                $result.Error = "The '$WorksheetName' sheet has headings on row $($result.HeaderRow) but no $missing column."
                                                return $result
                                }
                }

                $rows = New-Object System.Collections.Generic.List[object]
                $skipped = New-Object System.Collections.Generic.List[string]
                $currentDay = $null
                for ($i = 0; $i -lt $raw.Count; $i++) {
                                $sheetRow = $i + 1
                                if ($sheetRow -eq $result.HeaderRow) { continue }
                                $cells = & $cellsOf $raw[$i]
                                $filled = @($cells | Where-Object { $null -ne $_ -and ([string]$_).Trim() -ne '' })
                                if ($filled.Count -eq 0) { continue }
                                if (Test-ScheduleHeadingRow $cells) { continue }

                                $get = @{}
                                foreach ($f in @('ETD', 'Date', 'Instructor', 'SortieType', 'Student', 'Mission', 'Callsign', 'Num', 'Notes')) {
                                                $get[$f] = if ($map.ContainsKey($f) -and $map[$f] -le $cells.Count) { $cells[$map[$f] - 1] } else { $null }
                                }
                                $typeText = ConvertTo-ScheduleCellText $get.SortieType

                                # A row with a date and no line on it (or a lone date in any column) dates the rows below it
                                $noLine = ($typeText -eq '' -and (ConvertTo-ScheduleCellText $get.Instructor) -eq '' -and (ConvertTo-ScheduleCellText $get.Student) -eq '')
                                if ($noLine -or $filled.Count -eq 1) {
                                                $dayMark = $null
                                                foreach ($v in $filled) {
                                                                $tv = ConvertTo-FaisTimeValue $v
                                                                if ($null -ne $tv -and $tv.HasDate) { $dayMark = $tv.Value.Date; break }
                                                }
                                                if ($null -ne $dayMark) { $currentDay = $dayMark; continue }
                                }
                                if ($noLine) {
                                                if ($null -ne $get.ETD -and ([string]$get.ETD).Trim() -ne '') { $skipped.Add("row ${sheetRow}: no sortie type") }
                                                continue
                                }

                                $type = ConvertTo-CanonicalSortieType -Text $typeText -ReturnType $ReturnType
                                $etdText = ConvertTo-ScheduleCellText $get.ETD
                                if ($etdText -eq '') { $skipped.Add("row ${sheetRow}: no ETD ($type line)"); continue }
                                $tv = ConvertTo-FaisTimeValue $get.ETD
                                if ($null -eq $tv) { $skipped.Add("row ${sheetRow}: ETD '$etdText' is not a time (use 0730, 7:30 or a date and time)"); continue }
                                if ($typeText -eq '') { $skipped.Add("row ${sheetRow}: no sortie type"); continue }

                                $day = $null
                                if ($tv.HasDate) { $day = $tv.Value.Date }
                                else {
                                                $dv = ConvertTo-FaisTimeValue $get.Date
                                                if ($null -ne $dv -and $dv.HasDate) { $day = $dv.Value.Date }
                                                elseif ($null -ne $currentDay) { $day = $currentDay }
                                }
                                $time = $tv.Value.TimeOfDay
                                $ip = ConvertTo-ScheduleCellText $get.Instructor
                                if ($ip -eq '' -and (Test-IsSimSortieType $type)) { $ip = $SimInstructor }
                                $student = ConvertTo-ScheduleCellText $get.Student
                                $mission = ConvertTo-ScheduleCellText $get.Mission
                                $notes = ConvertTo-ScheduleCellText $get.Notes
                                $callsign = ConvertTo-ScheduleCellText $get.Callsign
                                $num = ConvertTo-ScheduleCellText $get.Num

                                $rows.Add([PSCustomObject]@{
                                                ETD        = $(if ($null -ne $day) { $day.Add($time).ToString('yyyy-MM-dd HH:mm:ss') } else { $null })
                                                Instructor = $(if ($ip -eq '') { $null } else { $ip })
                                                SortieType = $type
                                                Student    = $(if ($student -eq '') { $null } else { $student })
                                                Mission    = $(if ($mission -eq '') { $null } else { $mission })
                                                Callsign   = $(if ($callsign -eq '') { $null } else { $callsign })
                                                Num        = $(if ($num -eq '') { $null } else { $num })
                                                Notes      = $(if ($notes -eq '') { $null } else { $notes })
                                                Day        = $(if ($null -ne $day) { $day.ToString('yyyy-MM-dd') } else { '' })
                                                Time       = ('{0:00}:{1:00}' -f $time.Hours, $time.Minutes)
                                                TimeOnly   = ($null -eq $day)
                                                SheetRow   = $sheetRow
                                })
                }
                $result.Rows = $rows.ToArray()
                $result.Skipped = $skipped.ToArray()
                return $result
}

# The days a schedule sheet read has lines for, in date order, plus one entry with Day = ''
# for rows that have a time but no date. Each: {Day; Count; Students; Label}
function Get-ScheduleSheetDays {
                param ([array]$Rows = @())

                $groups = @{}
                foreach ($r in $Rows) {
                                $d = [string]$r.Day
                                if (-not $groups.ContainsKey($d)) { $groups[$d] = New-Object System.Collections.Generic.List[object] }
                                $groups[$d].Add($r)
                }
                $out = foreach ($d in ($groups.Keys | Sort-Object)) {
                                $list = $groups[$d]
                                $label = if ($d -eq '') { 'No date' } else { ([datetime]::ParseExact($d, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)).ToString('ddd dd MMM yyyy') }
                                [PSCustomObject]@{
                                                Day      = $d
                                                Count    = $list.Count
                                                Students = @($list | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.Student) }).Count
                                                Label    = $label
                                }
                }
                # Dated days first, the no-date group last
                return @(@($out | Where-Object { $_.Day -ne '' }) + @($out | Where-Object { $_.Day -eq '' }))
}

# The sheet's rows for one day (time-only rows too with -IncludeTimeOnly), as one short
# text that changes whenever any of those lines change. Row order does not matter.
function Get-ScheduleDayFingerprint {
                param ([array]$Rows = @(), [string]$Day, [switch]$IncludeTimeOnly)

                $lines = foreach ($r in $Rows) {
                                if ([string]$r.Day -ne $Day -and -not ($IncludeTimeOnly -and $r.TimeOnly)) { continue }
                                (@($r.Time, $r.SortieType, $r.Instructor, $r.Student, $r.Mission, $r.Callsign, $r.Num, $r.Notes) | ForEach-Object { ([string]$_).Trim() }) -join '|'
                }
                return (Get-FaisFingerprint -Parts @($lines))
}


# ====================================================================================
# _                    _ ____                  ____       _              _       _
#| |    ___   __ _  __| |  _ \ _ __ _____   __/ ___|  ___| |__   ___  __| |_   _| | ___
#| |   / _ \ / _` |/ _` | |_) | '__/ _ \ \ / /\___ \ / __| '_ \ / _ \/ _` | | | | |/ _ \
#| |__| (_) | (_| | (_| |  __/| | |  __/\ V /  ___) | (__| | | |  __/ (_| | |_| | |  __/
#|_____\___/ \__,_|\__,_|_|   |_|  \___| \_/  |____/ \___|_| |_|\___|\__,_|\__,_|_|\___|
# ====================================================================================
# Brings in the previous MasterSchedule with the student, mission and IP objects linked
# back up, so the pre-run menu can mark hooked / incomplete events.
function DownloadPrevSchedule {
                param (
                                $Students,
                                $Instructors,
                                $excelOutputPath
                )

                $filePath = $excelOutputPath
                if ([string]::IsNullOrWhiteSpace($filePath)) {
                                $filePath = Join-Path (Split-Path $PSScriptRoot -Parent) "OutputSchedule.xlsx"
                }

                if (-not (Test-Path $filePath)) {
                                Write-Warning "Could not find OutputSchedule.xlsx. Please run the main scheduler first."
                                return @()
                }

                try {
                                $RawSchedule = @(Import-Excel -Path $filePath -WorksheetName "MasterSchedule" -ErrorAction Stop)
                }
                catch {
                                Write-Warning "Could not read the MasterSchedule sheet in $filePath. Is it open in Excel? $($_.Exception.Message)"
                                return @()
                }

                foreach ($row in $RawSchedule) {
                                $dt = Parse-ETD $row.ETD
                                if ($null -ne $dt) { $row.ETD = $dt.ToString("yyyy-MM-dd HH:mm:ss") }

                                $studentName = Get-NameText $row.Student
                                if ($studentName -ne '' -and -not (Test-IsSentinelStudent $studentName)) {
                                                $matchedStudent = $Students | Where-Object { $_.Name -eq $studentName } | Select-Object -First 1
                                                if ($matchedStudent) { $row.Student = $matchedStudent }
                                }

                                $missionName = Get-NameText $row.Mission
                                if ($missionName -ne '' -and $row.Student -isnot [string] -and $null -ne $row.Student) {
                                                $matchedMission = Find-SyllabusMission -Student $row.Student -Mission $missionName
                                                if ($matchedMission) { $row.Mission = $matchedMission }
                                }

                                $ipName = Get-NameText $row.Instructor
                                if ($ipName -ne '' -and $ipName -notmatch "SIM IP") {
                                                $matchedInst = $Instructors | Where-Object { $_.Name -eq $ipName } | Select-Object -First 1
                                                if ($matchedInst) { $row.Instructor = $matchedInst }
                                }
                }
                return $RawSchedule
}


function Import-MasterSchedule {
    [CmdletBinding()]
    param(
        [string]$WorkbookPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'OutputSchedule.xlsx'),

        [string]$WorksheetName = 'MasterSchedule',

        # The destination properties required by the scheduling/menu code.
        [string[]]$RequiredColumns = @(
            'ETD',
            'SortieType',
            'Student',
            'Mission',
            'Instructor',
            'Notes'
        )
    )

    # Always return this object on failure rather than terminating the script.
    $emptySchedule = @()

    $excel     = $null
    $workbook  = $null
    $worksheet = $null
    $usedRange = $null

    try {
        # Locate OutputSchedule if the supplied default path does not exist.
        if (-not (Test-Path -LiteralPath $WorkbookPath)) {
            $faisDir = Split-Path $PSScriptRoot -Parent

            $matches = @(
                Get-ChildItem -Path $faisDir -File -ErrorAction SilentlyContinue |
                    Where-Object {
                        $_.BaseName -eq 'OutputSchedule' -and
                        $_.Extension -in '.xlsx', '.xlsm', '.xls'
                    }
            )

            if ($matches.Count -eq 0) {
                Write-Warning "OutputSchedule workbook was not found in the FAIS folder. Continuing with an empty schedule."
                return $emptySchedule
            }

            if ($matches.Count -gt 1) {
                Write-Warning "Multiple OutputSchedule workbooks were found in the FAIS folder. Continuing with an empty schedule. Specify -WorkbookPath explicitly."
                return $emptySchedule
            }

            $WorkbookPath = $matches[0].FullName
        }

        # Start Excel and open the workbook read-only.
        try {
                    $excel = New-Object -ComObject Excel.Application -ErrorAction Stop -Verbose:$false
            $excel.Visible = $false
            $excel.DisplayAlerts = $false

            $workbook = $excel.Workbooks.Open($WorkbookPath, 0, $true)
        }
        catch {
            Write-Warning "Could not open '$WorkbookPath'. Continuing with an empty schedule. Details: $($_.Exception.Message)"
            return $emptySchedule
        }

        # Locate the requested worksheet.
        try {
            $worksheet = $workbook.Worksheets.Item($WorksheetName)
        }
        catch {
            $availableSheets = @($workbook.Worksheets | ForEach-Object { $_.Name }) -join ', '

            Write-Warning (
                "Worksheet '$WorksheetName' was not found in '$WorkbookPath'. " +
                "Available worksheets: $availableSheets. Continuing with an empty schedule."
            )

            return $emptySchedule
        }

        $usedRange = $worksheet.UsedRange
        $rowCount = [int]$usedRange.Rows.Count
        $columnCount = [int]$usedRange.Columns.Count

        # Require a header row and at least one data row.
        if ($rowCount -lt 2 -or $columnCount -lt 1) {
            Write-Warning "Worksheet '$WorksheetName' has no usable data rows. Continuing with an empty schedule."
            return $emptySchedule
        }

        # Normalize headings so variations such as "Sortie Type", "sortie_type",
        # and "SORTIETYPE" are treated as the same name.
        function Normalize-ColumnName {
            param([string]$Name)

            if ([string]::IsNullOrWhiteSpace($Name)) {
                return ''
            }

            return (($Name -replace '[^a-zA-Z0-9]', '').ToLowerInvariant())
        }

        # Create a lookup table: normalized heading -> Excel column number.
        $columnLookup = @{}

        for ($column = 1; $column -le $columnCount; $column++) {
            $headerText = ([string]$usedRange.Cells.Item(1, $column).Text).Trim()
            $normalizedHeader = Normalize-ColumnName $headerText

            if (-not [string]::IsNullOrWhiteSpace($normalizedHeader)) {
                if ($columnLookup.ContainsKey($normalizedHeader)) {
                    Write-Warning "Duplicate column heading '$headerText' found. Using the first occurrence."
                }
                else {
                    $columnLookup[$normalizedHeader] = $column
                }
            }
        }

        # Check for required properties but do not stop execution if one is absent.
        $missingColumns = @()

        foreach ($requiredColumn in $RequiredColumns) {
            $normalizedRequiredColumn = Normalize-ColumnName $requiredColumn

            if (-not $columnLookup.ContainsKey($normalizedRequiredColumn)) {
                $missingColumns += $requiredColumn
            }
        }

        if ($missingColumns.Count -gt 0) {
            Write-Warning (
                "Worksheet '$WorksheetName' is missing expected column(s): " +
                "$($missingColumns -join ', '). " +
                "The importer will continue; missing values will be set to `$null."
            )
        }

        # Convert each Excel row into a schedule object.
                $schedule = for ($row = 2; $row -le $rowCount; $row++) {
            $record = [ordered]@{}

            # Ensure the core properties always exist.
            foreach ($requiredColumn in $RequiredColumns) {
                $normalizedRequiredColumn = Normalize-ColumnName $requiredColumn

                if ($columnLookup.ContainsKey($normalizedRequiredColumn)) {
                    $excelColumn = $columnLookup[$normalizedRequiredColumn]
                    $record[$requiredColumn] = $usedRange.Cells.Item($row, $excelColumn).Value2
                }
                else {
                    $record[$requiredColumn] = $null
                }
            }

            # Skip completely blank rows.
            $hasData = $record.Values | Where-Object {
                $null -ne $_ -and -not [string]::IsNullOrWhiteSpace([string]$_)
            }

            if ($hasData) {
                [PSCustomObject]$record
            }
        }

        Write-Verbose "Imported $(@($schedule).Count) record(s) from '$WorksheetName' in '$WorkbookPath'."

        return @($schedule)
    }
    catch {
        # Final protection against unexpected errors.
        Write-Warning "Unexpected error while importing MasterSchedule. Continuing with an empty schedule. Details: $($_.Exception.Message)"
        return $emptySchedule
    }
    finally {
        # Release Excel resources without allowing cleanup failures to stop the script.
        if ($usedRange) {
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($usedRange) } catch {}
        }

        if ($worksheet) {
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) } catch {}
        }

        if ($workbook) {
            try { $workbook.Close($false) } catch {}
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) } catch {}
        }

        if ($excel) {
            try { $excel.Quit() } catch {}
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch {}
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}
