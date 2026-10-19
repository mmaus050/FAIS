# DATA Download.PSM1 DIRECTORY:
# Contains all functions that downloads data to use in the program
# IngestData
# DownloadPrevSchedule (formally known as Import-MasterSchedule) - I have two down there rn, Download PrevSchedule links all the data back up whereas Import-MasterSchedule just returns strings TBD on what I wantInges
# The Functions from DatabaseUpdate or Update From Daily Flying
#
#
# ============================================================================================================


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
function IngestData {
                param (
                                [string]$excelInputPath,
                                [string]$desktopPath,
                                [string]$excelOutputPath
                )

                # Build the output path from the Desktop folder when it was not passed in
                if ([string]::IsNullOrWhiteSpace($excelOutputPath) -and -not [string]::IsNullOrWhiteSpace($desktopPath)) {
                                $excelOutputPath = Join-Path $desktopPath "OutputSchedule.xlsx"
                }

                # ------------------------------------------------------------------
                # Instructors
                # ------------------------------------------------------------------
                $excelData = Import-Excel -Path $excelInputPath -WorksheetName "Instructors"
                $Instructors = @(
                                foreach ($IP in $excelData) {
                                                [PSCustomObject]@{
                                                                Name          = $IP.IPName
                                                                IsCounter     = $IP.IsCounter
                                                                IsExperienced = $IP.IsExperienced
                                                                IsCheck       = $IP.IsCheck
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
                $validStudents = @(
                                $rawStudents |
                                                Where-Object { -not [string]::IsNullOrWhiteSpace($_.SPName) } |
                                                Select-Object -First 20
                )

                $Students = @(
                                foreach ($Student in $validStudents) {
                                                $studentSyllabus = $null

                                                # Prefer the student's own tab in OutputSchedule when that file exists
                                                if (Test-Path $excelOutputPath) {
                                                                try {
                                                                                $studentSyllabus = Import-Excel -Path $excelOutputPath -WorksheetName $Student.SPName -ErrorAction Stop
                                                                }
                                                                catch {
                                                                                Write-Host "WARNING: Could not read $($Student.SPName) from OutputSchedule. Is Excel currently open? (Ignore if you're starting a class)" -ForegroundColor Yellow
                                                                }
                                                }

                                                # Fall back to the master Syllabus sheet when no student-specific tab was found
                                                if ($null -eq $studentSyllabus) {
                                                                $studentSyllabus = Import-Excel -Path $excelInputPath -WorksheetName "Syllabus"
                                                }

                                                # Normalize each mission into the standard syllabus object
                                                $cleanSyllabus = @(
                                                                foreach ($mission in $studentSyllabus) {
                                                                                $isComp = $false
                                                                                if ($mission.IsComplete -eq $true -or $mission.IsComplete -match "True") {
                                                                                                $isComp = $true
                                                                                }

                                                                                [PSCustomObject]@{
                                                                                                SyllabusID = $mission.SyllabusID
                                                                                                Mission    = $mission.Mission
                                                                                                Prereqs    = $mission.Prereqs
                                                                                                Instructor = $mission.Instructor
                                                                                                ETD        = $mission.ETD
                                                                                                Type       = $mission.Type
                                                                                                IsCheck    = $mission.IsCheck
                                                                                                IsComplete = $isComp
                                                                                                NumOpting  = $mission.OptingEvents
                                                                                                Level      = $mission.Level
                                                                                }
                                                                }
                                                )

                                                [PSCustomObject]@{
                                                                ID              = $Student.SPKey
                                                                Name            = $Student.SPName
                                                                IsCap           = $Student.IsCap
                                                                IsInternational = $Student.IsInternational
                                                                Priority        = $Student.Priority
                                                                UnsatSim        = $Student.UnsatSim
                                                                UnsatFlight     = $Student.UnsatFlight
                                                                OnHold          = $Student.OnHold
                                                                SimDebt         = if ($null -eq $Student.SimDebt -or $Student.SimDebt -eq "") {
                                                                                0.0
                                                                }
                                                                else {
                                                                                [double]$Student.SimDebt
                                                                }
                                                                Syllabus        = $cleanSyllabus
                                                                Absences        = @()
                                                }
                                }
                )

                # ------------------------------------------------------------------
                # Schedule (fixed column positions, skip pasted headers, stop on blank ETD)
                # ------------------------------------------------------------------
                $rawSchedule = Import-Excel -Path $excelInputPath -WorksheetName "Schedule" -NoHeader
                $headerNames = @(
                                "ETD", "Date", "Callsign", "Num", "Instructor", "Student", "Mission",
                                "SortieType", "Notes", "Sortie Type", "Event Name", "#", "Remark",
                                "Device", "Simulator Instructor", "Start Time"
                )
                $CleanSchedule = @()

                foreach ($row in $rawSchedule) {
                                # Skip rows whose first column is a known header name
                                if ($row.P1 -in $headerNames) {
                                                continue
                                }

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
                                if ([string]::IsNullOrWhiteSpace($cleanRow.ETD)) {
                                                break
                                }

                                # Convert Excel serial dates and date strings to a standard timestamp
                                if ($cleanRow.ETD -match '^\d+(\.\d+)?$') {
                                                $cleanRow.ETD = [datetime]::FromOADate([double]$cleanRow.ETD).ToString("yyyy-MM-dd HH:mm:ss")
                                }
                                else {
                                                $parsed = $cleanRow.ETD -as [datetime]
                                                if ($parsed) {
                                                                $cleanRow.ETD = $parsed.ToString("yyyy-MM-dd HH:mm:ss")
                                                }
                                }

                                # OFT/IFT default Instructor to "SIM IP" only when that column was left blank
                                if (($cleanRow.SortieType -like "*OFT*" -or $cleanRow.SortieType -like "*IFT*") -and [string]::IsNullOrWhiteSpace($cleanRow.Instructor)) {
                                                $cleanRow.Instructor = "SIM IP"
                                }

                                $CleanSchedule += $cleanRow
                }

                $Schedule = @($CleanSchedule | Sort-Object -Property ETD)

                # ------------------------------------------------------------------
                # Formations
                # ------------------------------------------------------------------
                $formData = Import-Excel -Path $excelInputPath -WorksheetName "FormCallsigns"
                $index = 1
                $Formations = @(
                                foreach ($row in $formData) {
                                                [PSCustomObject]@{
                                                                NumFormation = $index
                                                                Callsign     = $row.Callsign
                                                                Num          = $row.Num
                                                }
                                                $index++
                                }
                )

                # ------------------------------------------------------------------
                # Soft Weights (first data row of the SoftWeights sheet)
                # ------------------------------------------------------------------
                $SoftWeights = $null
                try {
                                $RawWeights = Import-Excel -Path $excelInputPath -WorksheetName "SoftWeights"
                                $Row = $RawWeights[0]
                                $SoftWeights = [PSCustomObject]@{
                                                ContString        = $Row.ContString
                                                String4           = $Row.String4
                                                String5           = $Row.String5
                                                Zipper            = $Row.Zipper
                                                NonCounter        = $Row.NonCounter
                                                BreakString       = $Row.BreakString
                                                ExhaustionMult    = $Row.ExhaustMult
                                                ExhaustionDays    = $Row.ExhaustDays
                                                BIT_X             = $Row.BITPerDay
                                                BIT_Y             = $Row.BITDays
                                                CAPBonus          = $Row.CAPBonus
                                                FormLookAhead     = $Row.FormLookAhead
                                                LevelDifference   = $Row.LevelDifference
                                                CounterContinuity = $Row.CounterContinuity
                                                OFTIFTPen         = $Row.OFTIFTPen
                                                DollarRideBonus   = $Row.DollarRideBonus
                                                TransSimRatio     = $Row.TransSimRatio
                                                FormSimRatio      = $Row.FormSimRatio
                                                SimStopMult       = $Row.SimStopMult
                                }
                                $SoftWeights.TransSimRatio = $SoftWeights.TransSimRatio * $SoftWeights.SimStopMult
                                $SoftWeights.FormSimRatio = $SoftWeights.FormSimRatio * $SoftWeights.SimStopMult
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

                                foreach ($row in $AbsenceData) {
                                                $studentName = if ($row.Student) { ([string]$row.Student).Trim() } else { "" }
                                                $isActiveRaw = $row.IsActive
                                                $startRaw = $row.StartTime
                                                $endRaw = $row.EndTime

                                                # Convert StartTime and EndTime to DateTime
                                                $start = $null
                                                $end = $null

                                                if ($startRaw -is [datetime]) {
                                                                $start = $startRaw
                                                }
                                                elseif ([string]$startRaw -match '^\d+(\.\d+)?$') {
                                                                $start = [datetime]::FromOADate([double]$startRaw)
                                                }
                                                else {
                                                                $start = $startRaw -as [datetime]
                                                }

                                                if ($endRaw -is [datetime]) {
                                                                $end = $endRaw
                                                }
                                                elseif ([string]$endRaw -match '^\d+(\.\d+)?$') {
                                                                $end = [datetime]::FromOADate([double]$endRaw)
                                                }
                                                else {
                                                                $end = $endRaw -as [datetime]
                                                }

                                                if ([string]::IsNullOrWhiteSpace($studentName)) {
                                                                continue
                                                }

                                                $isActive = $false
                                                if ($isActiveRaw -eq $true -or
                                                                $isActiveRaw -eq 'True' -or
                                                                $isActiveRaw -eq 'TRUE' -or
                                                                $isActiveRaw -eq 'true' -or
                                                                $isActiveRaw -eq 1 -or
                                                                $isActiveRaw -eq '1' -or
                                                                $isActiveRaw -eq 'Yes' -or
                                                                $isActiveRaw -eq 'YES') {
                                                                $isActive = $true
                                                }

                                                if (-not $isActive) {
                                                                continue
                                                }

                                                if ($null -eq $start -or $null -eq $end) {
                                                                continue
                                                }

                                                $StudentAbsences += [PSCustomObject]@{
                                                                Student   = $studentName
                                                                StartTime = $start
                                                                EndTime   = $end
                                                                Reason    = if ($row.Reason) { ([string]$row.Reason).Trim() } else { "" }
                                                }
                                }
                }
                catch {
                                Write-Host "Could not read the 'StudentAbsences' sheet." -ForegroundColor Yellow
                                $StudentAbsences = @()
                }

                # Attach each absence to the student it belongs to so HardRules can read it off the student object
                foreach ($s in $Students) {
                                $studentName = if ($s.Name) { $s.Name.Trim().ToLower() } else { "" }
                                if ([string]::IsNullOrWhiteSpace($studentName)) {
                                                continue
                                }

                                $matched = @()
                                foreach ($absence in $StudentAbsences) {
                                                $absenceName = $absence.Student.Trim().ToLower()
                                                if ($absenceName -eq $studentName -or $studentName -like "*$absenceName*" -or $absenceName -like "*$studentName*") {
                                                                $matched += $absence
                                                }
                                }
                                $s.Absences = $matched
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
function DownloadPrevSchedule {
                param (
                                $Students,
                                $Instructors,
                                $desktopPath,
                                $excelOutputPath
                )
                $desktopPath = [Environment]::GetFolderPath('Desktop')
                $filePath = Join-Path $desktopPath "OutputSchedule.xlsx"

                if (-not (Test-Path $filePath)) {
                                Write-Warning "Could not find OutputSchedule.xlsx. Please run the main scheduler first."
                                return @()
                }

                $RawSchedule = Import-Excel -Path $filePath -WorksheetName "MasterSchedule"

                foreach ($row in $RawSchedule) {
                                if ($row.ETD -match '^\d+(\.\d+)?$') {
                                                $row.ETD = [datetime]::FromOADate([double]$row.ETD).ToString("yyyy-MM-dd HH:mm:ss")
                                } elseif ($row.ETD -as [datetime]) {
                                                $row.ETD = ([datetime]$row.ETD).ToString("yyyy-MM-dd HH:mm:ss")
                                }

                                if (-not [string]::IsNullOrWhiteSpace($row.Student)) {
                                                $matchedStudent = $Students | Where-Object { $_.Name -eq $row.Student } | Select-Object -First 1
                                                if ($matchedStudent) { $row.Student = $matchedStudent }
                                }

                                if (-not [string]::IsNullOrWhiteSpace($row.Mission) -and $row.Student -is [PSCustomObject]) {
                                                $matchedMission = $row.Student.Syllabus | Where-Object { $_.Mission -eq $row.Mission } | Select-Object -First 1
                                                if ($matchedMission) { $row.Mission = $matchedMission }
                                }

                                if (-not [string]::IsNullOrWhiteSpace($row.Instructor) -and $row.Instructor -notmatch "SIM IP") {
                                                $matchedInst = $Instructors | Where-Object { $_.Name -eq $row.Instructor } | Select-Object -First 1
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
