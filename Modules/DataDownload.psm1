# DATADOWNLOAD.PSM1 DIRECTORY:
# Contains all functions that download data to use in the program
# IngestData            (MasterData.xlsx + each student's tab in OutputSchedule.xlsx)
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
# Schedule, Instructors, Students, Formations, SoftWeights, StudentAbsences
# The Schedule sheet is optional now that the puck board builds the lines.
function IngestData {
                param (
                                [string]$excelInputPath,
                                [string]$desktopPath,
                                [string]$excelOutputPath,
                                [int]$MaxStudents = 20
                )

                # Build the output path from the Desktop folder when it was not passed in
                if ([string]::IsNullOrWhiteSpace($excelOutputPath) -and -not [string]::IsNullOrWhiteSpace($desktopPath)) {
                                $excelOutputPath = Join-Path $desktopPath "OutputSchedule.xlsx"
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

                $masterSyllabus = $null
                $Students = @(
                                foreach ($Student in $validStudents) {
                                                $studentSyllabus = $null

                                                # Prefer the student's own tab in OutputSchedule when that file exists
                                                if (Test-Path $excelOutputPath) {
                                                                try {
                                                                                $studentSyllabus = Import-Excel -Path $excelOutputPath -WorksheetName ([string]$Student.SPName).Trim() -ErrorAction Stop
                                                                }
                                                                catch {
                                                                                Write-Host "WARNING: Could not read $($Student.SPName) from OutputSchedule. Is Excel currently open? (Ignore if you're starting a class)" -ForegroundColor Yellow
                                                                }
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
                                                $priority = 0.0
                                                if ($null -ne $Student.Priority -and [string]$Student.Priority -ne '') { [void][double]::TryParse([string]$Student.Priority, [ref]$priority) }

                                                [PSCustomObject]@{
                                                                ID               = $Student.SPKey
                                                                Name             = ([string]$Student.SPName).Trim()
                                                                IsCap            = (ConvertTo-FaisBool $Student.IsCap)
                                                                IsInternational  = (ConvertTo-FaisBool $Student.IsInternational)
                                                                Priority         = $priority
                                                                UnsatSim         = (ConvertTo-FaisBool $Student.UnsatSim)
                                                                UnsatFlight      = (ConvertTo-FaisBool $Student.UnsatFlight)
                                                                OnHold           = (ConvertTo-FaisBool $Student.OnHold)
                                                                SimDebt          = $debt
                                                                Syllabus         = $cleanSyllabus
                                                                Absences         = @()
                                                                Reservations     = @()
                                                                SyllabusRevision = 0
                                                }
                                }
                )

                # ------------------------------------------------------------------
                # Schedule sheet (optional: the puck board replaces it)
                # Fixed column positions, skip pasted headers, stop on blank ETD
                # ------------------------------------------------------------------
                $rawSchedule = @()
                try {
                                $rawSchedule = @(Import-Excel -Path $excelInputPath -WorksheetName "Schedule" -NoHeader -ErrorAction Stop)
                }
                catch {
                                $rawSchedule = @()
                }
                $headerNames = @(
                                "ETD", "Date", "Callsign", "Num", "Instructor", "Student", "Mission",
                                "SortieType", "Notes", "Sortie Type", "Event Name", "#", "Remark",
                                "Device", "Simulator Instructor", "Start Time"
                )
                $CleanSchedule = New-Object System.Collections.Generic.List[object]

                foreach ($row in $rawSchedule) {
                                # Skip rows whose first column is a known header name
                                if ($row.P1 -in $headerNames) { continue }

                                $cleanRow = [PSCustomObject]@{
                                                ETD        = $row.P1
                                                Instructor = $row.P2
                                                SortieType = $row.P3
                                                Student    = $row.P4
                                                Mission    = $row.P5
                                                Callsign   = $row.P6
                                                Num        = $row.P7
                                                Notes      = $row.P8
                                }

                                # Stop at the first row with a blank ETD
                                if ([string]::IsNullOrWhiteSpace([string]$cleanRow.ETD)) { break }

                                # Convert Excel serial dates and date strings to a standard timestamp
                                $parsed = Parse-ETD $cleanRow.ETD
                                if ($null -ne $parsed) { $cleanRow.ETD = $parsed.ToString("yyyy-MM-dd HH:mm:ss") }

                                # OFT/IFT default Instructor to "SIM IP" only when that column was left blank
                                if ((Test-IsSimSortieType ([string]$cleanRow.SortieType)) -and [string]::IsNullOrWhiteSpace([string]$cleanRow.Instructor)) {
                                                $cleanRow.Instructor = "SIM IP"
                                }

                                $CleanSchedule.Add($cleanRow)
                }

                $Schedule = @($CleanSchedule | Sort-Object -Property ETD)

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
                                                OFTIFTPen         = $weightRow.OFTIFTPen
                                                DollarRideBonus   = $weightRow.DollarRideBonus
                                                TransSimRatio     = $weightRow.TransSimRatio
                                                FormSimRatio      = $weightRow.FormSimRatio
                                                SimStopMult       = $weightRow.SimStopMult
                                }
                                $SoftWeights.TransSimRatio = [double]$SoftWeights.TransSimRatio * [double]$SoftWeights.SimStopMult
                                $SoftWeights.FormSimRatio = [double]$SoftWeights.FormSimRatio * [double]$SoftWeights.SimStopMult
                }
                catch {
                                Write-Host "An error occurred while loading Soft Weights: $($_.Exception.Message)" -ForegroundColor Yellow
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

                # Attach each absence to the student it belongs to so the hard rules can read it off the student object
                foreach ($s in $Students) {
                                $studentName = if ($s.Name) { $s.Name.Trim().ToLower() } else { "" }
                                if ([string]::IsNullOrWhiteSpace($studentName)) { continue }

                                $s.Absences = @(
                                                foreach ($absence in $StudentAbsences) {
                                                                $absenceName = $absence.Student.Trim().ToLower()
                                                                if ($absenceName -eq $studentName -or $studentName -like "*$absenceName*" -or $absenceName -like "*$studentName*") { $absence }
                                                }
                                )
                }

                return [PSCustomObject]@{
                                Schedule        = $Schedule
                                Instructors     = $Instructors
                                Students        = $Students
                                Formations      = $Formations
                                SoftWeights     = $SoftWeights
                                StudentAbsences = $StudentAbsences
                }
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
                                $desktopPath,
                                $excelOutputPath
                )

                $filePath = $excelOutputPath
                if ([string]::IsNullOrWhiteSpace($filePath)) {
                                if ([string]::IsNullOrWhiteSpace($desktopPath)) { $desktopPath = [Environment]::GetFolderPath('Desktop') }
                                $filePath = Join-Path $desktopPath "OutputSchedule.xlsx"
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
        [string]$WorkbookPath = (Join-Path $env:USERPROFILE 'Desktop\OutputSchedule.xlsx'),

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
            $desktop = Join-Path $env:USERPROFILE 'Desktop'

            $matches = @(
                Get-ChildItem -Path $desktop -File -ErrorAction SilentlyContinue |
                    Where-Object {
                        $_.BaseName -eq 'OutputSchedule' -and
                        $_.Extension -in '.xlsx', '.xlsm', '.xls'
                    }
            )

            if ($matches.Count -eq 0) {
                Write-Warning "OutputSchedule workbook was not found on the Desktop. Continuing with an empty schedule."
                return $emptySchedule
            }

            if ($matches.Count -gt 1) {
                Write-Warning "Multiple OutputSchedule workbooks were found on the Desktop. Continuing with an empty schedule. Specify -WorkbookPath explicitly."
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
