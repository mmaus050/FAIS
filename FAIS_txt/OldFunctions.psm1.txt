Connects SyllabusToStudent

Diagnostics

Formation Pairing1x8

GetOptedEvents1x2

HardAndSoftRules1x4

OutputToTeams

PreRunMenu

Scheduling Handling1x3

 

 

 

 

 

 

 

 

 

 

#Connecting the Syllabus to each student Import and Export

# cd "C:\Users\1595161645A\Desktop\ModularScheduler\Old Code"

# Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Unrestricted Run this if you have a permissions error

# NEED TO CHANGE THE NAME OF THE SCHEDULE FILE IF APPLICABLE

 

 

function Write-ConsoleStatus {

                [CmdletBinding()]

                param([Parameter(Mandatory=$true)][string]$Message)

                Write-Host -NoNewline "`r$Message"

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

 

 

 

 

 

function IngestStudents {

                # Setup paths to MasterData and the desktop OutputSchedule

                $masterPath = Join-Path (Split-Path $PSScriptRoot -Parent) "MasterData.xlsx"

                $desktopPath = [Environment]::GetFolderPath('Desktop')

                $outputPath = Join-Path $desktopPath "OutputSchedule.xlsx"

 

                # Read the Students sheet, keep only rows that have a real SPName, hard-cap at 20

                $rawStudents = Import-Excel -Path $masterPath -WorksheetName "Students"

                $Students = @(

                                $rawStudents |

                                                Where-Object { -not [string]::IsNullOrWhiteSpace($_.SPName) } |

                                                Select-Object -First 20

                )

 

                # Build one student object per valid row

                $StudentData = foreach ($Student in $Students) {

                                $studentSyllabus = $null

                                $DataSource = "MasterData"

 

                                # Prefer the student's own tab in OutputSchedule when it exists

                                if (Test-Path $outputPath) {

                                                try {

                                                                $rawSyllabus = Import-Excel -Path $outputPath -WorksheetName $Student.SPName -ErrorAction Stop

                                                                $studentSyllabus = $rawSyllabus

                                                                $DataSource = "OutputSchedule"

                                                }

                                                catch {

                                                                Write-Host "WARNING: Could not read $($Student.SPName) from OutputSchedule. Is Excel currently open? (Ignore if you're starting a class)" -ForegroundColor Red

                                                }

                                }

 

                                # Fall back to the master Syllabus sheet when no student-specific tab was found

                                if ($null -eq $studentSyllabus) {

                                                $studentSyllabus = Import-Excel -Path $masterPath -WorksheetName "Syllabus"

                                }

 

                                # Normalize each mission into the standard syllabus object

                                $cleanSyllabus = foreach ($mission in $studentSyllabus) {

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

 

                                # Emit the finished student object

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

                                }

                }

 

                return $StudentData

}

 

 

 

# Ingest Instructors

function IngestInstructors {

                #These two functions locate the file and change "Syllabus" when you want to change the worksheet

                #This also creates a custom object $syllabusPSCO

                $filePath = Join-Path (Split-Path $PSScriptRoot -Parent) "MasterData.xlsx"

               

                #This changes what worksheet we are reading now, we need to do this because we are using it in an array

                $excelData = Import-Excel -Path $filePath -WorksheetName "Instructors"

 

                #Trying to load all student data, including the syllabus into one large object

                $InstructorData = foreach($IP in $excelData){

                                [PSCustomObject]@{

                                Name = $IP.IPName

                                IsCounter = $IP.IsCounter

                                IsExperienced = $IP.IsExperienced

                                IsCheck = $IP.IsCheck

                                Callsign = $IP.Callsign

                                Num = $IP.Num

                                OSC = $IP.OSC

                                OSNum = $IP.OSNum

                                }

                }

                return $InstructorData

}

 

 

 

function IngestFormCallsigns {

                $filePath = Join-Path (Split-Path $PSScriptRoot -Parent) "MasterData.xlsx"

                $excelData = Import-Excel -Path $filePath -WorksheetName "FormCallsigns"

                $index = 1

 

                $Formations = foreach ($row in $excelData){

                                [PSCustomObject]@{

                                NumFormation = $index

                                Callsign = $row.Callsign

                                Num = $row.Num

                                }

                                $index++

                }

                return $Formations

}

 

 

 

 

Function LoadSoftWeights {

                try {

                                # Defines the path to your master data file

                                $filePath = Join-Path (Split-Path $PSScriptRoot -Parent) "MasterData.xlsx"

                               

                                # Imports the specific worksheet containing the soft weights

                                $RawWeights = Import-Excel -Path $filePath -WorksheetName "SoftWeights"

                               

                                # Grabs the first row of data beneath the headers

                                $Row = $RawWeights[0]

                               

                                # Maps the Excel columns to the PSCustomObject properties

                                $SoftWeights = [PSCustomObject]@{

                                                ContString                       = $Row.ContString

                                                String4                                                = $Row.String4

                                                String5                                                = $Row.String5

                                                Zipper                                  = $Row.Zipper

                                                NonCounter                    = $Row.NonCounter

                                                BreakString                     = $Row.BreakString

                                                ExhaustionMult                            = $Row.ExhaustMult

                                                ExhaustionDays                           = $Row.ExhaustDays

                                                BIT_X                                   = $Row.BITPerDay

                                                BIT_Y                                   = $Row.BITDays

                                                CAPBonus                       = $Row.CAPBonus

                                                FormLookAhead                          = $Row.FormLookAhead

                                                LevelDifference                             = $Row.LevelDifference

                                                CounterContinuity     = $Row.CounterContinuity

                                                OFTIFTPen               = $Row.OFTIFTPen

                                                DollarRideBonus                         = $Row.DollarRideBonus

                                                TransSimRatio                              = $Row.TransSimRatio

                                                FormSimRatio                               = $Row.FormSimRatio

                                                SimStopMult                  = $Row.SimStopMult

                                }

                               

                                $SoftWeights.TransSimRatio = $SoftWeights.TransSimRatio * $SoftWeights.SimStopMult

                                $SoftWeights.FormSimRatio = $SoftWeights.FormSimRatio * $SoftWeights.SimStopMult

 

 

 

                                return $SoftWeights

                }

                catch {

                                Write-Warning "An error occurred while loading Soft Weights: $($_.Exception.Message)"

                                return $null

                }

}

 

 

Function SimDebtUpdate {

 

                param (

                                $Sortie,

                                $SimStopMult,

                                $TransSimRatio,

                                $FormSimRatio

                )

                # Flights

                if ($Sortie.Mission.Mission -like "*FT5*" -or $Sortie.Mission.Mission -like "*FF5*" -or

            $Sortie.Mission.Mission -like "*AA5*") {

                                $Sortie.Student.SimDebt += $SimStopMult

                }

                # Trans Sims

                elseif ($Sortie.Mission.Mission -like "*FI3*" -or $Sortie.Mission.Mission -like "*FE3*" -or

                                $Sortie.Mission.Mission -like "*FI4*" -or $Sortie.Mission.Mission -like "*FE4*") {

                                $Sortie.Student.SimDebt -= $TransSimRatio

                }

 

                # Form Sims

                elseif ($Sortie.Mission.Mission -like "*FN3*" -or $Sortie.Mission.Mission -like "*FN4*" -or

                                $Sortie.Mission.Mission -like "*AA4*") {

                                $Sortie.Student.SimDebt -= $FormSimRatio

                }

                return $Sortie.Student.SimDebt

}

 

 

 

##

# SaveSimDebtToMasterData.psm1

# Robust helper module for saving Simulation Debt (SimDebt) information

# back to the MasterData.xlsx file.

# Designed to be highly modular, robust (with validation, error handling, logging),

# speed is secondary.

# NO REGEX USED ANYWHERE.

# All code uses functions, proper indentation with one tab per level.

# Follows the project's style for PowerShell scheduler.

 

function Save-SimDebtToMasterData {

<#

.SYNOPSIS

                Saves or updates Simulation Debt (SimDebt) data for students back into MasterData.xlsx.

                Primary goal: robustness and correctness. Uses structured error handling and validation.

 

.DESCRIPTION

                This function takes the processed $Students array (which should contain updated

                Syllabus or debt tracking info), extracts relevant SimDebt fields if present,

                and writes them back to the "Students" worksheet in MasterData.xlsx.

                It avoids destructive overwrites by using ClearSheet + re-export pattern safely.

                It supports optional backup creation before write.

                Does NOT use any regex matching (-match). Uses string methods and type checks only.

 

.PARAMETER StudentData

                The array of student objects (usually $Students from IngestStudents) that may

                contain updated SimDebt or syllabus completion counters.

 

.PARAMETER MasterDataPath

                Optional. Full path to MasterData.xlsx. If not provided, it auto-detects relative

                to the script root (standard project layout).

 

.PARAMETER CreateBackup

                Switch. If present, creates a timestamped backup of MasterData before modifying.

 

.PARAMETER SimDebtColumns

                Optional string array of column names that represent SimDebt fields to preserve/update.

                Default: @("SimDebt", "TotalSimHoursOwed", "SimHoursCompleted", "RemainingSimDebt")

 

.EXAMPLE

                Save-SimDebtToMasterData -StudentData $Students -CreateBackup

 

.NOTES

                Call this after your main scheduling loop has updated student objects.

                Always import this module with -Force during development.

#>

                [CmdletBinding()]

                param (

                                [Parameter(Mandatory = $true)]

                                [ValidateNotNullOrEmpty()]

                                [array]$StudentData,

 

                                [Parameter(Mandatory = $false)]

                                [string]$MasterDataPath,

 

                                [Parameter(Mandatory = $false)]

                                [switch]$CreateBackup,

 

                                [Parameter(Mandatory = $false)]

                                [string[]]$SimDebtColumns = @("SimDebt", "TotalSimHoursOwed", "SimHoursCompleted", "RemainingSimDebt")

                )

 

                # ==========================================

                #      ROBUST INITIALIZATION & VALIDATION

                # ==========================================

                Write-Verbose "Starting Save-SimDebtToMasterData function..."

 

                # Auto-detect MasterData path if not supplied (consistent with other modules)

                if ([string]::IsNullOrWhiteSpace($MasterDataPath)) {

                                $MasterDataPath = Join-Path (Split-Path $PSScriptRoot -Parent) "MasterData.xlsx"

                }

 

                # Validate that MasterData.xlsx actually exists

                if (-not (Test-Path -Path $MasterDataPath -PathType Leaf)) {

                                Write-Error "MasterData.xlsx not found at path: $MasterDataPath"

                                return

                }

 

                # Validate that StudentData contains at least one object

                if ($StudentData.Count -eq 0) {

                                Write-Warning "No student data provided to Save-SimDebtToMasterData. Nothing to save."

                                return

                }

 

                # ==========================================

                #      OPTIONAL BACKUP (ROBUST SAFETY)

                # ==========================================

                if ($CreateBackup) {

                                try {

                                                $backupDir = Join-Path (Split-Path $MasterDataPath -Parent) "MasterData Backups"

                                                if (-not (Test-Path $backupDir)) {

                                                                New-Item -ItemType Directory -Path $backupDir -Force | Out-Null

                                                }

                                                $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

                                                $backupFile = Join-Path $backupDir "MasterData_SimDebtBackup_$timestamp.xlsx"

                                                Copy-Item -Path $MasterDataPath -Destination $backupFile -Force

                                                Write-Verbose "Created backup of MasterData at: $backupFile"

                                }

                                catch {

                                                Write-Warning "Failed to create backup before saving SimDebt. Continuing anyway. Error: $_"

                                }

                }

 

                # ==========================================

                #      EXTRACT / PREPARE SIM DEBT DATA

                #      (Only keep relevant columns to avoid bloat)

                # ==========================================

                try {

                                # Build a clean table containing only the SimDebt-related columns + key identifier (Name)

                                $simDebtData = @()

 

                                foreach ($student in $StudentData) {

                                                # Create a new PSCustomObject with only the fields we care about for SimDebt

                                                $debtRow = [PSCustomObject]@{

                                                                Name = $student.Name

                                                }

 

                                                # Dynamically add only the columns the caller asked for (or that exist)

                                                foreach ($col in $SimDebtColumns) {

                                                                if ($student.PSObject.Properties.Name -contains $col) {

                                                                                $debtRow | Add-Member -MemberType NoteProperty -Name $col -Value $student.$col

                                                                }

                                                                else {

                                                                                # If the column doesn't exist on this student object yet, initialize as empty

                                                                                $debtRow | Add-Member -MemberType NoteProperty -Name $col -Value $null

                                                                }

                                                }

 

                                                $simDebtData += $debtRow

                                }

 

                                if ($simDebtData.Count -eq 0) {

                                                Write-Warning "After processing, no SimDebt rows were prepared. Nothing written to MasterData."

                                                return

                                }

 

                                # ==========================================

                                #      WRITE TO MASTERDATA (ROBUST PATTERN)

                                #      Use ClearSheet + Export to safely replace only the SimDebt section

                                #      Never use non-existent parameters like IncludeOnly

                                # ==========================================

 

                                # First, read the existing Students sheet so we can merge intelligently if needed

                                # (This keeps other student columns intact)

                                $existingStudents = Import-Excel -Path $MasterDataPath -WorksheetName "Students" -ErrorAction Stop

 

                                # Filter out any rows with blank/NULL Name (common with trailing empty rows in Excel)

                                $existingStudents = $existingStudents | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Name) }

 

                                # Create a lookup for quick debt updates

                                $debtLookup = @{}

                                foreach ($row in $simDebtData) {

                                                if (-not [string]::IsNullOrWhiteSpace($row.Name)) {

                                               $debtLookup[$row.Name] = $row

                                                }

                                }

 

                                # Merge: update only the SimDebt columns in the existing data

                                $updatedStudents = @()

                                foreach ($existingRow in $existingStudents) {

                                $studentName = $existingRow.Name

                                if (-not [string]::IsNullOrWhiteSpace($studentName) -and $debtLookup.ContainsKey($studentName)) {

                               $debtInfo = $debtLookup[$studentName]

                               foreach ($col in $SimDebtColumns) {

                                               if ($debtInfo.PSObject.Properties.Name -contains $col -and $debtInfo.$col -ne $null) {

                                                               $existingRow | Add-Member -MemberType NoteProperty -Name $col -Value $debtInfo.$col -Force

                                               }

                               }

                                }

                                $updatedStudents += $existingRow

                                }

 

                                # Now export the merged data back to the Students worksheet

                                # IMPORTANT: We do NOT use any -IncludeOnly parameter (it does not exist in Export-Excel)

                                # Instead we already filtered/merged the data above.

                                $updatedStudents | Export-Excel -Path $MasterDataPath -WorksheetName "Students" -AutoSize -ErrorAction Stop

 

                                Write-Verbose "Successfully saved SimDebt data for $($simDebtData.Count) students back to MasterData.xlsx"

                                Write-Host "SimDebt information successfully written to MasterData.xlsx (Students sheet)." -ForegroundColor Green

 

                }

                catch {

                                # Robust error handling - surface the real problem without crashing the whole scheduler

                                $errorMessage = $_.Exception.Message

                                Write-Error "Failed to save SimDebt: $errorMessage"

                                # Optionally re-throw if you want the caller to stop: throw

                                return

                }

 

                Write-Verbose "Save-SimDebtToMasterData completed successfully."

}

 

 

function Convert-SortieObjectsToText {

<#

                .SYNOPSIS

                                Converts complex Instructor, Mission, and Student objects to their .Name or .Mission string equivalents for clean export display.

                                This function mutates the input objects in place for simplicity and performance.

#>

                param(

                                [Parameter(Mandatory = $true)]

                                [AllowEmptyCollection()]

                                [array]$Schedule

                )

 

                foreach ($Sortie in $Schedule) {

                                if ($Sortie.Instructor -is [PSCustomObject]) {

                                                $Sortie.Instructor = $Sortie.Instructor.Name

                                }

                                if ($Sortie.Mission -is [PSCustomObject]) {

                                                $Sortie.Mission = $Sortie.Mission.Mission

                                }

                                if ($Sortie.Student -is [PSCustomObject]) {

                                                $Sortie.Student = $Sortie.Student.Name

                                }

                }

}

 

 

 

# ==========================================

#      UTILITY FUNCTIONS (MODULAR)

# ==========================================

 

function Close-AllOpenExcelWorkbooks {

<#

                .SYNOPSIS

                                Saves and closes every open Excel workbook via COM automation.

                                Prevents file locking conflicts when FAIS ingests or exports MasterData.xlsx.

                                Highly defensive: handles cases where Excel is not running or workbooks are protected.

#>

                param()

 

                try {

                                $excelApp = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")

                                if ($null -ne $excelApp) {

                                                Write-Host "FAIS: Open Excel application detected. Saving and closing all workbooks before run..." -ForegroundColor Yellow

                                                $workbookCount = $excelApp.Workbooks.Count

                                                for ($i = $workbookCount; $i -ge 1; $i--) {

                                                                $workbook = $excelApp.Workbooks.Item($i)

                                                                $workbookName = $workbook.Name

                                                                try {

                                                                                if (-not $workbook.Saved) {

                                                                                                $workbook.Save()

                                                                                                Write-Host "FAIS: Saved workbook -> $workbookName" -ForegroundColor Cyan

                                                                                }

                                                                                $workbook.Close($false)

                                                                                Write-Host "FAIS: Closed workbook -> $workbookName" -ForegroundColor Green

                                                                } catch {

                                                                                Write-Host "FAIS: Warning - could not save/close '$workbookName' (protected or in use)." -ForegroundColor Yellow

                                                                }

                                                }

                                                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelApp) | Out-Null

                                                [System.GC]::Collect()

                                                [System.GC]::WaitForPendingFinalizers()

                                                Write-Host "FAIS: All open Excel workbooks saved and closed. Ready to run." -ForegroundColor Green

                                } else {

                                                Write-Host "FAIS: No open Excel application found. Proceeding directly to scheduler run." -ForegroundColor Gray

                                }

                } catch [System.Runtime.InteropServices.COMException] {

                                Write-Host "FAIS: No running Excel instance found. Continuing with scheduler..." -ForegroundColor Gray

                } catch {

                                Write-Host "FAIS: Note - could not access Excel COM (safe to ignore if Excel not open)." -ForegroundColor Gray

                }

}

 

 

 

 

function Update-SimDebtColumnP {

    [CmdletBinding()]

    param(

        [Parameter(Mandatory)]

        [array]$StudentData,

 

        [Parameter(Mandatory)]

        [string]$MasterDataPath,

 

        [switch]$CreateBackup

    )

 

    if (-not (Test-Path $MasterDataPath)) {

        Write-Error "MasterData.xlsx not found at: $MasterDataPath"

        return

    }

 

    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

 

    # 1. Create full backup of MasterData.xlsx

    if ($CreateBackup) {

        try {

            $backupDir = Split-Path $MasterDataPath -Parent

            $fullBackup = Join-Path $backupDir "MasterData_FULL_Backup_$timestamp.xlsx"

            Copy-Item -Path $MasterDataPath -Destination $fullBackup -Force

            Write-Host "Full backup created: $fullBackup" -ForegroundColor Cyan

        }

        catch {

            Write-Warning "Could not create full backup. Continuing anyway."

        }

    }

 

    try {

        # 2. Read current Students sheet

        Write-Host "Reading current Students tab..." -ForegroundColor Yellow

        $students = Import-Excel -Path $MasterDataPath -WorksheetName "Students" -ErrorAction Stop

 

        # Remove any blank rows

        $students = $students | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Name) }

 

        # 3. Build lookup table from your incoming data (Name → SimDebt)

        $debtLookup = @{}

        foreach ($s in $StudentData) {

            if (-not [string]::IsNullOrWhiteSpace($s.Name) -and $s.PSObject.Properties.Name -contains 'SimDebt') {

                $debtLookup[$s.Name] = $s.SimDebt

            }

        }

 

        Write-Host "Found $($debtLookup.Count) students with SimDebt values to apply." -ForegroundColor Cyan

 

        # 4. Update ONLY the SimDebt column (in memory)

        $updatedCount = 0

        foreach ($row in $students) {

            if ($debtLookup.ContainsKey($row.Name)) {

                $row | Add-Member -MemberType NoteProperty -Name 'SimDebt' -Value $debtLookup[$row.Name] -Force

                $updatedCount++

            }

        }

 

        Write-Host "Updated SimDebt for $updatedCount students in memory." -ForegroundColor Green

 

        # 5. Write back WITHOUT clearing the sheet (this is the key to not wiping anything)

        $students | Export-Excel -Path $MasterDataPath -WorksheetName "Students" -AutoSize -ErrorAction Stop

 

        Write-Host "SUCCESS: SimDebt column (Column P) has been updated in MasterData.xlsx" -ForegroundColor Green

        Write-Host "Only the SimDebt column was modified. All other columns are untouched." -ForegroundColor Green

 

    }

    catch {

        Write-Error "Failed to update SimDebt column: $_"

    }

}

 

 

 

 

 

 

 

 

 

function Show-InstructionManual {

    [CmdletBinding()]

    param(

        [Parameter(Mandatory = $true)]

        [string[]]$Steps,

 

        [string]$Title = "Instruction Manual",

 

        [int]$FontSize = 22

    )

 

    Add-Type -AssemblyName System.Windows.Forms

    Add-Type -AssemblyName System.Drawing

 

    if ($Steps.Count -eq 0) {

        Write-Warning "No steps provided."

        return

    }

 

    # Shared state – hashtable is a reference type, so event handlers see the same object

    $state = @{

        Current = 0

    }

 

    # Main form

    $form = New-Object System.Windows.Forms.Form

    $form.Text            = $Title

    $form.Size            = New-Object System.Drawing.Size(700, 450)

    $form.StartPosition   = "CenterScreen"

    $form.FormBorderStyle = "FixedDialog"

    $form.MaximizeBox     = $false

    $form.MinimizeBox     = $false

    $form.TopMost         = $true

    $form.BackColor       = [System.Drawing.Color]::WhiteSmoke

 

    # Large instruction label

    $label = New-Object System.Windows.Forms.Label

    $label.Font           = New-Object System.Drawing.Font("Segoe UI", $FontSize, [System.Drawing.FontStyle]::Regular)

    $label.TextAlign      = "MiddleCenter"

    $label.Dock           = "Fill"

    $label.Padding        = New-Object System.Windows.Forms.Padding(30)

    $label.BackColor      = [System.Drawing.Color]::White

    $form.Controls.Add($label)

 

    # Button panel

    $buttonPanel = New-Object System.Windows.Forms.Panel

    $buttonPanel.Height     = 70

    $buttonPanel.Dock       = "Bottom"

    $buttonPanel.BackColor  = [System.Drawing.Color]::WhiteSmoke

    $form.Controls.Add($buttonPanel)

 

    # Previous button

    $btnPrev = New-Object System.Windows.Forms.Button

    $btnPrev.Text     = "<==  Previous"

    $btnPrev.Size     = New-Object System.Drawing.Size(140, 40)

    $btnPrev.Location = New-Object System.Drawing.Point(30, 15)

    $btnPrev.Font     = New-Object System.Drawing.Font("Segoe UI", 11)

    $buttonPanel.Controls.Add($btnPrev)

 

    # Step counter

    $lblCounter = New-Object System.Windows.Forms.Label

    $lblCounter.AutoSize  = $true

    $lblCounter.Font      = New-Object System.Drawing.Font("Segoe UI", 11)

    $lblCounter.Location  = New-Object System.Drawing.Point(300, 25)

    $buttonPanel.Controls.Add($lblCounter)

 

    # Next button

    $btnNext = New-Object System.Windows.Forms.Button

    $btnNext.Text     = "Next  ==>"

    $btnNext.Size     = New-Object System.Drawing.Size(140, 40)

    $btnNext.Location = New-Object System.Drawing.Point(520, 15)

    $btnNext.Font     = New-Object System.Drawing.Font("Segoe UI", 11)

    $buttonPanel.Controls.Add($btnNext)

 

    # Close button (shown only on last page)

    $btnClose = New-Object System.Windows.Forms.Button

    $btnClose.Text     = "Close"

    $btnClose.Size     = New-Object System.Drawing.Size(140, 40)

    $btnClose.Location = New-Object System.Drawing.Point(520, 15)

    $btnClose.Font     = New-Object System.Drawing.Font("Segoe UI", 11)

    $btnClose.Visible  = $false

    $buttonPanel.Controls.Add($btnClose)

 

    # Update UI from current state

    $updateDisplay = {

        $label.Text       = $Steps[$state.Current]

        $lblCounter.Text  = "Step $($state.Current + 1) of $($Steps.Count)"

        $btnPrev.Enabled  = ($state.Current -gt 0)

        $btnNext.Visible  = ($state.Current -lt $Steps.Count - 1)

        $btnClose.Visible = ($state.Current -eq $Steps.Count - 1)

    }

 

    # Event handlers – they all close over the same $state object

    $btnPrev.Add_Click({

        if ($state.Current -gt 0) {

            $state.Current--

            & $updateDisplay

        }

    }.GetNewClosure())

 

    $btnNext.Add_Click({

        if ($state.Current -lt $Steps.Count - 1) {

            $state.Current++

            & $updateDisplay

        }

    }.GetNewClosure())

 

    $btnClose.Add_Click({ $form.Close() }.GetNewClosure())

 

    # First page

    & $updateDisplay

 

    [void]$form.ShowDialog()

    $form.Dispose()

}

 

 

 

function Get-SimDebtDelta {

                <#

                .SYNOPSIS

                                Returns the SimDebt change contributed by a single completed mission.

                #>

                param (

                                [Parameter(Mandatory = $true)]

                                [string]$MissionName,

 

                                [Parameter(Mandatory = $true)]

                                [double]$SimStopMult,

 

                                [Parameter(Mandatory = $true)]

                                [double]$TransSimRatio,

 

                                [Parameter(Mandatory = $true)]

                                [double]$FormSimRatio

                )

 

                # Flights increase debt

                if ($MissionName -like "*FT5*" -or $MissionName -like "*FF5*" -or $MissionName -like "*AA5*") {

                                return $SimStopMult

                }

 

                # Transition sims decrease debt

                if ($MissionName -like "*FI3*" -or $MissionName -like "*FE3*" -or

                                $MissionName -like "*FI4*" -or $MissionName -like "*FE4*") {

                                return -$TransSimRatio

                }

 

                # Formation sims decrease debt

                if ($MissionName -like "*FN3*" -or $MissionName -like "*FN4*" -or $MissionName -like "*AA4*") {

                                return -$FormSimRatio

                }

 

                return 0.0

}

 

 

function Calculate-SimDebtFromSyllabus {

                <#

                .SYNOPSIS

                                Walks every event on a student worksheet and returns the total SimDebt.

                #>

                param (

                                [Parameter(Mandatory = $true)]

                                [array]$Syllabus,

 

                                [Parameter(Mandatory = $true)]

                                [double]$SimStopMult,

 

                                [Parameter(Mandatory = $true)]

                                [double]$TransSimRatio,

 

                                [Parameter(Mandatory = $true)]

                                [double]$FormSimRatio

                )

 

                $debt = 0.0

 

                foreach ($mission in $Syllabus) {

                                # Normalize IsComplete the same way IngestStudents does

                                $isComplete = $false

                                if ($mission.IsComplete -eq $true -or $mission.IsComplete -eq "True") {

                                                $isComplete = $true

                                }

 

                                if (-not $isComplete) { continue }

                                if ([string]::IsNullOrWhiteSpace($mission.Mission)) { continue }

 

                                $debt += Get-SimDebtDelta -MissionName $mission.Mission `

                                                -SimStopMult $SimStopMult `

                                                -TransSimRatio $TransSimRatio `

                                                -FormSimRatio $FormSimRatio

                }

 

                return $debt

}

 

 

 

 

 

 

 

 

 

 

 

 

 

function Update-MasterDataSimDebtColumnP {

                <#

                .SYNOPSIS

                                Writes calculated SimDebt values into Column P of the Students sheet.

                                Matches by SPName in column B so only the debt cell is changed.

                #>

                param (

                                [Parameter(Mandatory = $true)]

                                [hashtable]$SimDebtByName,

 

                                [Parameter(Mandatory = $true)]

                                [string]$MasterDataPath

                )

 

                if (-not (Test-Path $MasterDataPath)) {

                                Write-Error "MasterData.xlsx not found at: $MasterDataPath"

                                return

                }

 

                $package = $null

                try {

                                $package = Open-ExcelPackage -Path $MasterDataPath

                                $ws = $package.Workbook.Worksheets["Students"]

 

                                if (-not $ws) {

                                                Write-Error "Worksheet 'Students' not found in MasterData.xlsx"

                                                return

                                }

 

                                # Column B = SPName, Column P = SimDebt

                                $nameCol = 2

                                $debtCol = 16

                                $row = 2

                                $updatedCount = 0

 

                                while ($true) {

                                                $nameVal = $ws.Cells[$row, $nameCol].Value

                                                if ([string]::IsNullOrWhiteSpace($nameVal)) { break }

 

                                                $studentName = $nameVal.ToString().Trim()

                                                if ($SimDebtByName.ContainsKey($studentName)) {

                                                                $ws.Cells[$row, $debtCol].Value = [double]$SimDebtByName[$studentName]

                                                                $updatedCount++

                                                }

 

                                                $row++

                                }

 

                                Close-ExcelPackage -ExcelPackage $package -SaveAs $MasterDataPath

                                Write-Host "Successfully updated SimDebt (Column P) for $updatedCount students." -ForegroundColor Green

                }

                catch {

                                Write-Error "Failed to update SimDebt in MasterData: $_"

                                if ($package) {

                                                Close-ExcelPackage -ExcelPackage $package

                                }

                }

}

 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

<#

function Update-SimDebtWorksheet {

                <#

                .SYNOPSIS

                                Creates or overwrites the SimDebt worksheet in MasterData.xlsx.

                                Each row is StudentName + SimDebt.

               

                param (

                                [Parameter(Mandatory = $true)]

                                [hashtable]$SimDebtByName,

 

                                [Parameter(Mandatory = $true)]

                                [string]$MasterDataPath

                )

 

                if (-not (Test-Path $MasterDataPath)) {

                                Write-Error "MasterData.xlsx not found at: $MasterDataPath"

                                return

                }

 

                # Build the rows that will become the SimDebt sheet

                $exportRows = foreach ($name in ($SimDebtByName.Keys | Sort-Object)) {

                                [PSCustomObject]@{

                                                Name    = $name

                                                SimDebt = [double]$SimDebtByName[$name]

                                }

                }

 

                try {

                                $exportRows | Export-Excel -Path $MasterDataPath `

                                                -WorksheetName "SimDebt" `

                                                -ClearSheet `

                                                -AutoSize `

                                                -ErrorAction Stop

 

                                Write-Host "Successfully wrote $($exportRows.Count) rows to the SimDebt worksheet." -ForegroundColor Green

                }

                catch {

                                Write-Error "Failed to write SimDebt worksheet: $_"

                }

}

#>

 

 

# Diagnostics and Testing

 

function pathDiagnostics{

               

                Write-Host "--- PATH DIAGNOSTICS ---" -ForegroundColor Cyan

 

                $moduleFolder = $PSScriptRoot

                Write-Host "Looking for modules in: $moduleFolder"

 

                # 1. Check for the first module

                $module1 = Join-Path $moduleFolder "ConnectsSyllabusToStudent.psm1" # Update this to the full name

                if (Test-Path $module1) { Write-Host "[OK] Found ConnectSyllabusToStudent" -ForegroundColor Green }

                else { Write-Host "[ERROR] Missing ConnectSyllabusToStudent" -ForegroundColor Red }

 

                # 2. Check for the second module

                $module1 = Join-Path $moduleFolder "ScheduleHandling.psm1" # Update this to the full name

                if (Test-Path $module1) { Write-Host "[OK] Found ScheduleHandling" -ForegroundColor Green }

                else { Write-Host "[ERROR] Missing ScheduleHandling" -ForegroundColor Red }

 

                # 3. Check for the third module

                $module1 = Join-Path $moduleFolder "GetOptedEvents.psm1" # Update this to the full name

                if (Test-Path $module1) { Write-Host "[OK] Found GetOptedEvents" -ForegroundColor Green }

                else { Write-Host "[ERROR] Missing GetOptedEvents" -ForegroundColor Red }

 

                # 4. Check for the fourth module

                $module1 = Join-Path $moduleFolder "HardAndSoftRules.psm1" # Update this to the full name

                if (Test-Path $module1) { Write-Host "[OK] Found HardAndSoftRules" -ForegroundColor Green }

                else { Write-Host "[ERROR] Missing HardAndSoftRules" -ForegroundColor Red }

 

<#

                # 5. Check for MasterData.xlsx (Checking the main folder)

                $excelPath = Split-Path -Parent $PSScriptRoot "MasterData.xlsx"

                if (Test-Path $excelPath) { Write-Host "[OK] Found MasterData.xlsx in main folder" -ForegroundColor Green }

                else { Write-Host "[ERROR] Missing MasterData.xlsx in main folder" -ForegroundColor Red }

#>

 

                Write-Host "------------------------" -ForegroundColor Cyan

}

 

 

function Trace-DataRead {

    param (

        [array]$Students

    )

 

    Write-Host "`n=== DIAGNOSTIC: Data Read Check ===" -ForegroundColor Cyan

    Write-Host "Total Students Loaded: $($Students.Count)"

    foreach ($s in $Students) {

        if ($null -ne $s.Syllabus) {

            $completed = @($s.Syllabus | Where-Object { $_.IsComplete -eq $true -or $_.IsComplete -eq 'True' }).Count

            $incomplete = @($s.Syllabus | Where-Object { $_.IsComplete -eq $false -or $_.IsComplete -eq 'False' }).Count

            Write-Host "Student: $($s.Name) | Syllabus Items: $($s.Syllabus.Count) | Complete: $completed | Incomplete: $incomplete"

        } else {

            Write-Host "Student: $($s.Name) | Syllabus Loaded: FAILED (Null)" -ForegroundColor Red

        }

    }

    Write-Host "===================================`n" -ForegroundColor Cyan

}

 

 

 

function Trace-SchedulingDecision {

    param (

        $Sortie,

        $SPCandidates,

        $BestMatch,

        $SelectedMission

    )

    Write-Host "`n=== DIAGNOSTIC: Scheduling Decision for ETD $($Sortie.ETD) ===" -ForegroundColor Yellow

    Write-Host "Eligible Candidates: $($SPCandidates.Count)"

   

    foreach ($c in $SPCandidates) {

        $opted = $c.OptedEvents.Mission -join ', '

        Write-Host "  - Candidate: $($c.Name) | Score: $($c.Score) | Opted Events: $opted"

    }

   

    Write-Host "--> WINNER: $($BestMatch.Name) with Score $($BestMatch.Score)" -ForegroundColor Green

    Write-Host "--> ASSIGNED MISSION: $SelectedMission" -ForegroundColor Green

    Write-Host "===========================================================`n" -ForegroundColor Yellow

}

 

 

 

##

# Handles the logic for pairing a Formation flight. First attempts to find another student

# to act as a wingman. If none are available, it assigns the next available live-fly Day Local

# sortie inside the look-ahead window as Direct Support.

# FormCallsigns is stored as consecutive Lead/Wing pairs (e.g. HAZE 11, HAZE 12).

# Each successful or attempted pair consumes two entries and advances FormIndex by 2.

function Pair-FormationWingman {

                param (

                                $PrimarySortie,

                                $PrimaryStudent,

                                $PrimaryMission,

                                $Schedule,

                                $Students,

                                $Instructors,

                                $SoftWeights,

                                $FormLookAhead,

                                $LevelDifference,

                                $FormIndex,

                                $Formations

                )

                try {

                                if ($null -eq $Formations -or $Formations.Count -eq 0) {

                                                return $FormIndex

                                }

 

                                # Normalize FormIndex into range and resolve the Lead / Wing pair

                                if ($FormIndex -ge $Formations.Count) { $FormIndex = 0 }

                                $LeadEntry = $Formations[$FormIndex]

 

                                $WingIndex = $FormIndex + 1

                                if ($WingIndex -ge $Formations.Count) { $WingIndex = 0 }

                                $WingEntry = $Formations[$WingIndex]

 

                                # Lead always receives the Lead entry from the pair table

                                $PrimarySortie.Callsign = $LeadEntry.Callsign

                                $PrimarySortie.Num = $LeadEntry.Num

 

                                $CurrentETD = [datetime]$PrimarySortie.ETD

                                $TimeWindowEnd = $CurrentETD.AddMinutes($FormLookAhead)

                                $BestWingmanCandidates = @()

 

                                # Search only open live-fly Day Local lines inside the look-ahead window

                                foreach ($WingSortie in $Schedule) {

                                                if (-not [string]::IsNullOrWhiteSpace($WingSortie.Student)) { continue }

                                                if ($WingSortie.SortieType -like "*Return*") { continue }

                                                if ($WingSortie.SortieType -like "*SIM*" -or $WingSortie.SortieType -like "*OFT*") { continue }

                                                if ($WingSortie.SortieType -like "*O&B*") { continue }

                                                if ($WingSortie -eq $PrimarySortie) { continue }

 

                                                $FutureETD = [datetime]$WingSortie.ETD

                                                if ($FutureETD -lt $CurrentETD -or $FutureETD -gt $TimeWindowEnd) { continue }

 

                                                # Resolve instructor object only — never copy instructor Callsign/Num onto a formation line

                                                if ($WingSortie.Instructor -isnot [PSCustomObject]) {

                                                                $WingInst = Get-MatchedInstructor -Instructors $Instructors -TargetName $WingSortie.Instructor

                                                                if ($null -eq $WingInst) {

                                                                                $fallback = if ([string]::IsNullOrWhiteSpace($WingSortie.Instructor)) { "NEED IP" } else { $WingSortie.Instructor }

                                                                                $WingInst = [PSCustomObject]@{ Name = $fallback; IsCounter = $true; IsExperienced = $false; Callsign = ""; Num = "" }

                                                                }

                                                                $WingSortie.Instructor = $WingInst

                                                }

 

                                                $HighestScoringForLine = $null

 

                                                foreach ($FormStudent in $Students) {

                                                                if ($FormStudent.Name -eq $PrimaryStudent.Name) { continue }

 

                                                                $ValidEvents = Get-ValidOptedEvents -Student $FormStudent -Sortie $WingSortie

                                                                if ($null -eq $ValidEvents) { continue }

 

                                                                $TentativeWingMission = $null

                                                                foreach ($Event in $ValidEvents) {

                                                                                if ($Event.Type -like "*Form*") {

                                                                                                $TentativeWingMission = $Event

                                                                                                break

                                                                                }

                                                                }

                                                                if ($null -eq $TentativeWingMission) { continue }

 

                                                                $WingSortie.Mission = $TentativeWingMission

 

                                                                if (HardRules -Student $FormStudent -Sortie $WingSortie -Schedule $Schedule -FormLookAhead $FormLookAhead -Students $Students -IsWingCandidate) {

                                                                                $WingScore = SoftRules -Student $FormStudent -Config $SoftWeights -Sortie $WingSortie

 

                                                                                $primaryLvl = if ($null -ne $PrimaryMission.Level) { [int]$PrimaryMission.Level } else { 0 }

                                                                                $wingLvl = if ($null -ne $TentativeWingMission.Level) { [int]$TentativeWingMission.Level } else { 0 }

                                                                                $WingScore += ($wingLvl - $primaryLvl) * $LevelDifference

 

                                                                                $CandidateObj = [PSCustomObject]@{

                                                                                                Student = $FormStudent

                                                                                                Sortie  = $WingSortie

                                                                                                Score   = $WingScore

                                                                                                Mission = $TentativeWingMission

                                                                                }

 

                                                                                if ($null -eq $HighestScoringForLine -or $WingScore -gt $HighestScoringForLine.Score) {

                                                                                                $HighestScoringForLine = $CandidateObj

                                                                                }

                                                                }

 

                                                                $WingSortie.Mission = $null

                                                }

 

                                                if ($null -ne $HighestScoringForLine) {

                                                                $BestWingmanCandidates += $HighestScoringForLine

                                                }

                                }

 

                                # Finalize student wingman when one exists

                                if ($BestWingmanCandidates.Count -gt 0) {

                                                $AbsoluteBest = $BestWingmanCandidates | Sort-Object -Property Score -Descending | Select-Object -First 1

 

                                                $WinningSortie = $AbsoluteBest.Sortie

                                                $WinningStudent = $AbsoluteBest.Student

                                                $WinningMission = $AbsoluteBest.Mission

 

                                                $WinningSortie.Student = $WinningStudent.Name

                                                $WinningSortie.Mission = $WinningMission

 

                                                # Wing always receives the Wing entry from the same pair

                                                $WinningSortie.Callsign = $WingEntry.Callsign

                                                $WinningSortie.Num = $WingEntry.Num

 

                                                $WinningSortie.Notes = "Wingman to $($PrimaryStudent.Name)"

                                                $PrimarySortie.Notes = "Wingman: $($WinningStudent.Name)"

 

                                                foreach ($s in $Students) {

                                                                if ($s.Name -eq $WinningStudent.Name) {

                                                                                foreach ($m in $s.Syllabus) {

                                                                                                if ($m.Mission -eq $WinningMission.Mission) {

                                                                                                                $m.IsComplete = $true

                                                                                                                $m.Instructor = $WinningSortie.Instructor.Name

                                                                                                                $m.ETD = $WinningSortie.ETD

                                                                                                                break

                                                                                                }

                                                                                }

                                                                                break

                                                                }

                                                }

                                } else {

                                                # No student wingman — claim soonest open live-fly Day Local as Direct Support

                                                $DirectSupportSortie = $null

                                                foreach ($Candidate in $Schedule) {

                                                                if (-not [string]::IsNullOrWhiteSpace($Candidate.Student)) { continue }

                                                                if ($Candidate.SortieType -like "*Return*") { continue }

                                                                if ($Candidate.SortieType -like "*SIM*" -or $Candidate.SortieType -like "*OFT*") { continue }

                                                                if ($Candidate.SortieType -like "*O&B*") { continue }

                                                                if ($Candidate -eq $PrimarySortie) { continue }

 

                                                                $CandidateETD = [datetime]$Candidate.ETD

                                                                if ($CandidateETD -lt $CurrentETD -or $CandidateETD -gt $TimeWindowEnd) { continue }

 

                                                                $DirectSupportSortie = $Candidate

                                                                break

                                                }

 

                                                if ($null -ne $DirectSupportSortie) {

                                                                $DirectSupportSortie.Student = "Need Form SP"

                                                                $DirectSupportSortie.Callsign = $WingEntry.Callsign

                                                                $DirectSupportSortie.Num = $WingEntry.Num

                                                                $DirectSupportSortie.Notes = "Direct Support for $($PrimaryStudent.Name)"

                                                                $PrimarySortie.Notes = "Wingman: DS"

                                                }

                                }

 

                                # Always consume the full Lead/Wing pair

                                $FormIndex += 2

                                if ($FormIndex -ge $Formations.Count) { $FormIndex = 0 }

 

                                return $FormIndex

 

                } catch {

                                Write-Warning "Error resolving formation pairing: $($_.Exception.Message)"

                                return $FormIndex

                }

}

 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

# Reorders the schedule so each wingman sits immediately under its lead on the same day.

# Identifies wingmen by even Num (FormCallsigns pairs are Lead odd / Wing even).

Function Sort-Formations {

                param (

                                $Schedule

                )

               

                try {

                                # Iterates through the schedule using a 0-indexed for-loop

                                for ($movingIndex = 0; $movingIndex -lt $Schedule.Count; $movingIndex++) {

                                                $Sortie = $Schedule[$movingIndex]

                                               

                                                if ($Sortie.Mission.Type -like "*Form*") {

                                                               

                                                                # Checks if the Sortie Num is even, indicating it is a Wingman

                                                                if ($Sortie.Num % 2 -eq 0) {

                                                                                $targetNum = $Sortie.Num - 1

                                                                                $targetCallsign = $Sortie.Callsign

                                                                                $wingmanDate = ([datetime]$Sortie.ETD).Date

                                                                                $storedIndex = -1

                                                                               

                                                                                # Scans the schedule to find the corresponding Lead sortie index on the same day

                                                                                for ($i = 0; $i -lt $Schedule.Count; $i++) {

                                                                                                if ($Schedule[$i].Num -eq $targetNum -and

                                                                                                                $Schedule[$i].Callsign -eq $targetCallsign -and

                                                                                                                ([datetime]$Schedule[$i].ETD).Date -eq $wingmanDate) {

                                                                                                                $storedIndex = $i

                                                                                                                break

                                                                                                }

                                                                                }

                                                                               

                                                                                # Collects indices of non-O&B sorties between the Lead and the Wingman (inclusive of Wingman)

                                                                                $movableIndices = @()

                                                                                if ($storedIndex -ge 0) {

                                                                                                for ($i = $storedIndex + 1; $i -le $movingIndex; $i++) {

                                                                                                                if ($Schedule[$i].SortieType -notlike "*O&B*") {

                                                                                                                                $movableIndices += $i

                                                                                                                }

                                                                                                }

                                                                                }

                                                                               

                                                                                # Moves the Wingman into the earliest non-O&B slot after the Lead when it is not already there

                                                                                if ($movableIndices.Count -gt 0 -and $movableIndices[0] -ne $movingIndex) {

                                                                                               

                                                                                                # Stores the ETDs of the movable slots in chronological order

                                                                                                $blockETDs = @()

                                                                                                foreach ($idx in $movableIndices) {

                                                                                                                $blockETDs += $Schedule[$idx].ETD

                                                                                                }

                                                                                               

                                                                                                # Builds the new order of objects with the Wingman placed first

                                                                                                $movableObjects = @()

                                                                                                foreach ($idx in $movableIndices) {

                                                                                                                if ($idx -ne $movingIndex) {

                                                                                                                                $movableObjects += $Schedule[$idx]

                                                                                                                }

                                                                                                }

                                                                                                $movableObjects = @($Schedule[$movingIndex]) + $movableObjects

                                                                                               

                                                                                                # Writes the reordered objects back into the movable positions and reapplies the original ETDs

                                                                                                for ($k = 0; $k -lt $movableIndices.Count; $k++) {

                                                                                                                $Schedule[$movableIndices[$k]] = $movableObjects[$k]

                                                                                                                $Schedule[$movableIndices[$k]].ETD = $blockETDs[$k]

                                                                                                }

                                                                                }

                                                                }

                                                }

                                }

                                return $Schedule

                }

                catch {

                                Write-Warning "An error occurred while sorting formations: $($_.Exception.Message)"

                                return $Schedule

                }

}

 

 

 

 

 

 

#Finds the opted events for a student

 

function Get-OptedEvents {

    param (

        [Parameter(Mandatory=$true)]

        [array]$SyllabusData # Accepts an array of custom objects

    )

 

    # 1. Build a hash table for blazing-fast lookups

    $completionStatus = @{}

    foreach ($row in $SyllabusData) {

        # Handle both boolean and string representations safely

        $isComplete = [bool]$row.IsComplete -or $row.IsComplete -eq 'True'

        $completionStatus[$row.Mission] = $isComplete

    }

 

    $optedEvents = @()

 

    # 2. Find Opted Events

    foreach ($row in $SyllabusData) {

        $isComplete = [bool]$row.IsComplete -or $row.IsComplete -eq 'True'

 

        # Disqualify the event immediately if it is already complete

        if (-not $isComplete) {

 

            $prereqsMet = $true

 

            if ($row.Prereqs -ne '0' -and -not [string]::IsNullOrWhiteSpace($row.Prereqs)) {

                $prereqList = $row.Prereqs -split ',' | ForEach-Object { $_.Trim() }

 

                foreach ($prereq in $prereqList) {

                    if (-not $completionStatus[$prereq]) {

                        $prereqsMet = $false

                        break

                    }

                }

            }

 

            if ($prereqsMet) {

                $optedEvents += $row

            }

        }

    }

 

    return $optedEvents

}

 

 

 

 

 

 

 

 

 

 

 

 

function Get-ValidOptedEvents {

                param (

                                $Student,

                                $Sortie

                )

 

                $AllOptedEvents = Get-OptedEvents -SyllabusData $Student.Syllabus

                if (-not $AllOptedEvents) { return $null }

 

                $OptedEvents = @()

                foreach ($Event in $AllOptedEvents) {

                                # SOLO RULE: Exclude "SingleSolo" and "FormSolo" from O&Bs

                                if ($Sortie.SortieType -match "(?i)O&B|O\s*&\s*B") {

                                                if ($Event.Type -in @("Single", "Formation", "Both")) {

                                                                $OptedEvents += $Event

                                                }

                                }

                                elseif ($Sortie.SortieType -like "*Day Local*") {

                                                if ($Event.Type -in @("Single", "SingleSolo", "Formation", "FormSolo", "Both")) {

                                                                $OptedEvents += $Event

                                                }

                                }

                                elseif ($Sortie.SortieType -like "*IFT*") {

                                                if ($Event.Type -eq "IFT") {

                                                                $OptedEvents += $Event

                                                }

                                }

                                elseif ($Sortie.SortieType -like "*OFT*") {

                                                if ($Event.Type -in @("IFT", "OFT")) {

                                                                $OptedEvents += $Event

                                                }

                                }

                }

 

                if ($OptedEvents.Count -eq 0) { return $null }

 

                # O&B DOUBLE RIDE PREREQ (LOOK-AHEAD)

                if ($Sortie.SortieType -match "(?i)O&B$|O\s*&\s*B$") {

                                $SimulatedMission = $OptedEvents | Sort-Object -Property SyllabusID | Select-Object -First 1

                               

                                $SimulatedMission.IsComplete = $true

                               

                                $NextEvents = Get-OptedEvents -SyllabusData $Student.Syllabus

                                $NextFlightEvents = @()

 

                                if ($null -ne $NextEvents) {

                                                foreach ($e in $NextEvents) {

                                                                # Leg 2 also cannot be a Solo event

                                                                if ($e.Type -in @("Single", "Formation", "Both")) {

                                                                                $NextFlightEvents += $e

                                                                }

                                                }

                                }

                               

                                $SimulatedMission.IsComplete = $false

                               

                                if ($NextFlightEvents.Count -eq 0) {

                                                return $null

                                }

                }

 

                return $OptedEvents

}

 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

# --- How to use it ---

# $availableMissions = Get-OptedEvents -SyllabusData $importData

# $availableMissions | Select-Object Mission, Prereqs, Type, Level

 

 

 

 

 

 

##

# ===========================================================

# RESETTING THE VARIABLES

$TurnTimeCheck = $true

$CrewRestCheck = $false

$DutyDayCheck = $false

$OandBLock = $false

# ==========================================================

 

# Safely converts both Excel serial dates and normal date strings

function Parse-ETD {

                param ($rawDate)

                if ([string]::IsNullOrWhiteSpace($rawDate)) { return $null }

               

                if ($rawDate -match '^\d+(\.\d+)?$') {

                                return [datetime]::FromOADate([double]$rawDate)

                }

               

                return [datetime]$rawDate

}

 

 

 

function SoftRules {

                param (

                                $Student,

                                $Config,

                                $Sortie,

                                $FlightRecords,

                                [switch]$Detailed

                )

 

                $Score = 0.0

                $Breakdown = [ordered]@{}

                $NewNotes = @()

                $SortieDate = Parse-ETD $Sortie.ETD

 

                ####################################

                # RULE : DOLLAR RIDE BONUS

                # Gets dollar rides out the door

                ####################################

                if ($Sortie.Mission.Mission -like "*FT5001*") {

                                $pts = $Config.DollarRideBonus

                                $Score += $pts

                                $Breakdown["Dollar Ride Bonus"] = $pts

                }

 

                ###############################################################

                # RULE : Sim Stop Prevention

                # Every flight that gets scheduled creates a 'debt' that is paid off by completing sims

                ###############################################################

                if ($Sortie.Mission.Mission -like "*FT5*" -or $Sortie.Mission.Mission -like "*FF5*" -or

                                $Sortie.Mission.Mission -like "*AA5*") {

                                $pts = -$Student.SimDebt

                                $Score += $pts

                                $Breakdown["Sim Debt (flight)"] = $pts

                }

                elseif ($Sortie.Mission.Mission -like "*FI3*" -or $Sortie.Mission.Mission -like "*FE3*" -or

                                $Sortie.Mission.Mission -like "*FI4*" -or $Sortie.Mission.Mission -like "*FE4*" -or

                                $Sortie.Mission.Mission -like "*FN3*" -or $Sortie.Mission.Mission -like "*FN4*" -or

                                $Sortie.Mission.Mission -like "*AA4*") {

                                $pts = $Student.SimDebt

                                $Score += $pts

                                $Breakdown["Sim Debt (sim)"] = $pts

                }

 

                ####################################

                # RULE: CAP BONUS

                # CAPsters don't usually double turn, so give them a bonus to get on the schedule

                ####################################

                if ($Student.IsCap) {

                                $pts = $Config.CAPBonus

                                $Score += $pts

                                $Breakdown["CAP Bonus"] = $pts

                }

 

                #####################################

                # RULE: STRINGS AND COUNTERS

                # Rewards three strings, penalizes everything else to varying degrees

                #####################################

                $PastFlights = @(

                                $Student.Syllabus | Where-Object {

                                                ($_.IsComplete -eq $true -or $_.IsComplete -match "True") -and

                                                -not [string]::IsNullOrWhiteSpace($_.ETD) -and

                                                $_.Instructor -ne "SIM IP" -and

                                                -not [string]::IsNullOrWhiteSpace($_.Instructor) -and

                                                (Parse-ETD $_.ETD) -lt $SortieDate

                                } | Sort-Object -Property { Parse-ETD $_.ETD } -Descending

                )

 

                $IP1 = if ($PastFlights.Count -ge 1) { $inst = $PastFlights[0].Instructor; $name = if ($inst -is [PSCustomObject]) { $inst.Name } else { $inst }; ([string]$name).Trim() } else { "" }

                $IP2 = if ($PastFlights.Count -ge 2) { $inst = $PastFlights[1].Instructor; $name = if ($inst -is [PSCustomObject]) { $inst.Name } else { $inst }; ([string]$name).Trim() } else { "" }

                $IP3 = if ($PastFlights.Count -ge 3) { $inst = $PastFlights[2].Instructor; $name = if ($inst -is [PSCustomObject]) { $inst.Name } else { $inst }; ([string]$name).Trim() } else { "" }

                $IP4 = if ($PastFlights.Count -ge 4) { $inst = $PastFlights[3].Instructor; $name = if ($inst -is [PSCustomObject]) { $inst.Name } else { $inst }; ([string]$name).Trim() } else { "" }

 

                $CurrentIP = ""

                $IsEffectiveCounter = $true

 

                if ($null -ne $Sortie.InstructorDetails -and $null -ne $Sortie.InstructorDetails.Name) {

                                $CurrentIP = ([string]$Sortie.InstructorDetails.Name).Trim()

                                $IsEffectiveCounter = $Sortie.InstructorDetails.IsCounter

                } elseif ($null -ne $Sortie.Instructor) {

                                $inst = $Sortie.Instructor

                                $name = if ($inst -is [PSCustomObject] -or $inst.GetType().Name -eq "PSCustomObject") { $inst.Name } else { $inst }

                                $CurrentIP = ([string]$name).Trim()

                }

 

                if ($CurrentIP -ne "" -and $CurrentIP -notmatch "SIM IP") {

                                $IP1_Base = if ($IP1) { $IP1 -replace '\s*\[.*?\]', '' } else { "" }

                                $IP2_Base = if ($IP2) { $IP2 -replace '\s*\[.*?\]', '' } else { "" }

                                $IP3_Base = if ($IP3) { $IP3 -replace '\s*\[.*?\]', '' } else { "" }

                                $IP4_Base = if ($IP4) { $IP4 -replace '\s*\[.*?\]', '' } else { "" }

                                $CurrentIP_Base = $CurrentIP -replace '\s*\[.*?\]', ''

 

                                if (($IsEffectiveCounter -eq $false -or $IsEffectiveCounter -match "False") -and ($CurrentIP_Base -eq $IP1_Base -or ($CurrentIP_Base -eq $IP2_Base -and $CurrentIP_Base -ne ""))) {

                                                $IsEffectiveCounter = $true

                                }

 

                                if ($IsEffectiveCounter -eq $false -or $IsEffectiveCounter -match "False") {

                                                $pts = $Config.NonCounter

                                                $Score += $pts

                                                $Breakdown["Non-Counter"] = $pts

                                }

                                else {

                                                if ($IP1_Base -eq $CurrentIP_Base) {

                                                                $Str4 = if ($null -ne $Config.String4) { $Config.String4 } else { -10 }

                                                                $Str5 = if ($null -ne $Config.String5) { $Config.String5 } else { -100 }

 

                                                                if ($IP1_Base -eq $IP2_Base -and $IP2_Base -eq $IP3_Base -and $IP3_Base -eq $IP4_Base) {

                                                                                $pts = $Str5

                                                                                $Score += $pts

                                                                                $Breakdown["String 5"] = $pts

                                                                }

                                                                elseif ($IP1_Base -eq $IP2_Base -and $IP2_Base -eq $IP3_Base) {

                                                                                $pts = $Str4

                                                                                $Score += $pts

                                                                                $Breakdown["String 4"] = $pts

                                                                }

                                                                else {

                                                                                $pts = $Config.ContString

                                                                                $Score += $pts

                                                                                $Breakdown["Continue String"] = $pts

                                                                                if ($Sortie.Student.IsCap) {

                                                                                                $pts2 = $Config.CAPBonus

                                                                                                $Score += $pts2

                                                                                                $Breakdown["CAP Bonus (string)"] = $pts2

                                                                                }

                                                                }

                                                }

                                                elseif ($IP2_Base -eq $CurrentIP_Base -and $IP1_Base -ne $CurrentIP_Base -and $IP2_Base -ne "") {

                                                                $pts = $Config.Zipper

                                                                $Score += $pts

                                                                $Breakdown["Zipper"] = $pts

                                                }

                                                elseif ($IP1_Base -ne "" -and $IP1_Base -ne $CurrentIP_Base) {

                                                                $pts = $Config.BreakString

                                                                $Score += $pts

                                                                $Breakdown["Break String"] = $pts

                                                }

                                }

                }

 

                ##################################################

                # RULE : SIMULATOR TYPE MATCHING

                # Prioritizes putting OFT events in OFT slots and IFT events in IFT slots.

                ##################################################

                if ($Sortie.SortieType -match "OFT" -and $Sortie.Mission.Type -match "IFT") {

                                $pts = $OFTIFTPen

                                $Score += $pts

                                $Breakdown["OFT/IFT Mismatch"] = $pts

                                $NewNotes += "IFT in an OFT"

                }

 

                ############################################

                # RULE: EXHAUSTION POINTS

                # Starts subtracting points the busier and busier the SP is

                ############################################

                $ExhaustionDays = if ($null -ne $Config.ExhaustionDays) { $Config.ExhaustionDays } else { 5 }

                $ExhaustionMult = if ($null -ne $Config.ExhaustionMult) { $Config.ExhaustionMult } else { 1 }

 

                $RecentEvents = @(

                                $Student.Syllabus | Where-Object {

                                                ($_.IsComplete -eq $true -or $_.IsComplete -match "True") -and

                                                -not [string]::IsNullOrWhiteSpace($_.ETD)

                                }

                ) | Where-Object {

                                $EventDate = Parse-ETD $_.ETD

                                if ($null -ne $EventDate) {

                                                $DaysDiff = ($SortieDate.Date - $EventDate.Date).TotalDays

                                                $DaysDiff -ge 0 -and $DaysDiff -le $ExhaustionDays -and $EventDate -lt $SortieDate

                                } else {

                                                $false

                                }

                }

                $RecentSims = @($RecentEvents | Where-Object { $_.Instructor -match "SIM IP" }).Count

                $RecentFlights = @($RecentEvents | Where-Object { $_.Instructor -notmatch "SIM IP" }).Count

 

                $ExhaustionPoints = ($RecentFlights * 2) + ($RecentSims * 1)

                $pts = $ExhaustionPoints * $ExhaustionMult

                $Score += $pts

                $Breakdown["Exhaustion"] = $pts

 

                ##########################################################

                # RULE: BIT TRACKER

                # Only run the BIT Tracker if the event being scheduled is a flight (Not IFT/OFT)

                ##########################################################

                if ($Sortie.SortieType -notmatch "(?i)OFT|IFT") {

                                if ($PastFlights.Count -gt 0) {

                                                $BIT_X = if ($null -ne $Config.BIT_X) { $Config.BIT_X } else { 1.5 }

                                                $BIT_Y = if ($null -ne $Config.BIT_Y) { $Config.BIT_Y } else { 6 }

 

                                                $LastFlightDate = Parse-ETD $PastFlights[0].ETD

                                                $DaysSinceLastFlight = ($SortieDate.Date - $LastFlightDate.Date).TotalDays

                                                $EffectiveBreakDays = $DaysSinceLastFlight - 1

 

                                                $CAP_BIT_multiplyer = 1

                                                if ($Student.IsCap) {

                                                                $CAP_BIT_multiplyer = 2

                                                }

 

                                                if ($EffectiveBreakDays -gt 0) {

                                                                if ($EffectiveBreakDays -le $BIT_Y) {

                                                                                $pts = ($EffectiveBreakDays * $BIT_X * $CAP_BIT_multiplyer)

                                                                                $Score += $pts

                                                                                $Breakdown["BIT Tracker"] = $pts

                                                                } else {

                                                                                $pts = ($BIT_X * ($BIT_Y / 2.0))

                                                                                $Score += $pts

                                                                                $Breakdown["BIT Tracker (capped)"] = $pts

                                                                                $NewNotes += "BIT! Consider an 86"

                                                                }

                                                }

                                }

                }

 

                ##################################################

                # RULE : Counter Limit

                # Tries to keep the total number of counters low

                ##################################################

                $AllPastIPs = @(

                                $PastFlights | Where-Object {

                                                -not [string]::IsNullOrWhiteSpace($_.Instructor) -and $_.Instructor -notmatch "SIM IP"

                                } | ForEach-Object {

                                                $ipName = if ($_.Instructor -is [PSCustomObject]) { $_.Instructor.Name } else { [string]$_.Instructor }

                                                $ipName.Trim() -replace '\s*\[.*?\]', ''

                                }

                ) | Select-Object -Unique

 

                if (($IsEffectiveCounter -eq $true -or $IsEffectiveCounter -match "True") -and ($AllPastIPs -contains $CurrentIP_Base)) {

                                $pts = if ($null -ne $Config.CounterContinuity) { $Config.CounterContinuity } else { 5 }

                                $Score += $pts

                                $Breakdown["Counter Continuity"] = $pts

                }

 

                ################################################

                # RULE : PRIORITIZE CHECKRIDES

                # If we have a check IP, we should use them for checkrides

                ################################################

                if ($Sortie.Instructor.IsCheck -and $Sortie.Mission.IsCheck) {

                                $pts = 50

                                $Score += $pts

                                $Breakdown["Checkride Priority"] = $pts

                                $NewNotes += "Checkride! Requires Formal"

                }

 

                ##########################################################

                # RULE: STUDENT PRIORITY

                # This just spots them whatever their priority is

                ##########################################################

                $pts = $Student.Priority

                $Score += $pts

                $Breakdown["Student Priority"] = $pts

 

                if ($NewNotes.Count -gt 0) {

                                $joinedNotes = $NewNotes -join " | "

                                $Student | Add-Member -MemberType NoteProperty -Name "PendingNote" -Value $joinedNotes -Force

                } else {

                                $Student | Add-Member -MemberType NoteProperty -Name "PendingNote" -Value $null -Force

                }

 

                if ($Detailed) {

                                return [PSCustomObject]@{

                                                Score      = $Score

                                                Breakdown  = $Breakdown

                                                Notes      = $NewNotes

                                                Student    = $Student.Name

                                                Mission    = $Sortie.Mission.Mission

                                                ETD        = $Sortie.ETD

                                }

                }

 

                return $Score

}

 

 

 

 

 

 

 

 

 

 

 

 

############################################################

#

# HARD RULES

#

############################################################

function HardRules {

                param (

                                $Student,

                                $Sortie,

                                $Schedule,

                                $FormLookAhead = 22,

                                $Students,

                                [switch]$IsWingCandidate

                )

               

                $NewNotes = @()       

                $NextStart = Parse-ETD $Sortie.ETD

 

                $completedEvents = @(

                                $Student.Syllabus | Where-Object {

                                                ($_.IsComplete -eq $true -or $_.IsComplete -eq 'True') -and

                                                -not [string]::IsNullOrWhiteSpace($_.ETD)

                                } | Sort-Object -Property { Parse-ETD $_.ETD } -Descending

                )

 

                $todaysEvents = @(

                                $completedEvents | Where-Object {

                                                $evtDate = Parse-ETD $_.ETD

                                                $null -ne $evtDate -and $evtDate.Date -eq $NextStart.Date

                                }

                )

 

 

 

 

 

# ==========================================================

# HARD RULE: STUDENT UNAVAILABILITY / ABSENCES

#

# This blocks a student from being scheduled if the proposed sortie

# overlaps with any time they are marked unavailable in the StudentAbsences sheet.

#

# NEW BUFFER LOGIC (as requested):

#   - The sortie is considered to "really begin" 1.5 hours BEFORE the ETD

#   - The sortie is considered to "end"       2.5 hours AFTER  the ETD

#

# This creates a larger protected window around the sortie so the student

# is not double-booked with appointments, medical, leave, etc.

#

# Example: A 1000 sortie will now be treated as occupying ~0830 – 1230

#          when checking against student absences.

# ==========================================================

 

if ($null -ne $global:StudentAbsences -and $global:StudentAbsences.Count -gt 0) {

 

    $BufferBeforeHours = -1.5

    $BufferAfterHours  =  2.5

 

    $SortieEffectiveStart = $NextStart.AddHours($BufferBeforeHours)

    $SortieEffectiveEnd   = $NextStart.AddHours($BufferAfterHours)

 

    foreach ($absence in $global:StudentAbsences) {

 

        # Very defensive name matching

        $absenceName = $absence.Student.Trim().ToLower()

        $studentName = if ($Student.Name) { $Student.Name.Trim().ToLower() } else { "" }

 

        if ($absenceName -eq $studentName -or $studentName -like "*$absenceName*" -or $absenceName -like "*$studentName*") {

 

            #Write-Host ">>> ABSENCE CHECK TRIGGERED for $studentName against $($absence.StartTime.ToString('yyyy-MM-dd HH:mm')) - $($absence.EndTime.ToString('yyyy-MM-dd HH:mm'))" -ForegroundColor Yellow

 

            if ($SortieEffectiveStart -lt $absence.EndTime -and $SortieEffectiveEnd -gt $absence.StartTime) {

                #Write-Host "    >>> BLOCKING - Overlap detected!" -ForegroundColor Red

                $NewNotes += "Student unavailable ($($absence.Reason))"

                return $false

            }

            else {

                #Write-Host "    No overlap with this absence" -ForegroundColor DarkGray

            }

        }

    }

}

 

 

 

 

 

##################################################

# Block checkrides on O&Bs entirely

# - TESTED

##################################################

                if ($Sortie.SortieType -like "*O&B*") {

                                if ($Sortie.Mission.IsCheck) {

                                                return $false

                                }

                }

 

 

 

#################################################

# TEMPORARY Block Formations from O&B

# - TESTED

#################################################

                if ($Sortie.Mission.Type -like "*Form*" -and $Sortie.SortieType -like "*O&B*") {

                                return $false

                }

 

 

 

#################################################

# Prevent O&B Return Overwrite

# - TESTED

#################################################

                if ($Sortie.SortieType -like "*Return*") {

                                return $false

                }

 

 

 

 

 

 

 

 

 

 

################################################

                # FORMATION BLOCK

                # Only applied when the student is being considered as a lead.

                # A wing candidate already has a lead that claimed this line, so the

                # open-line count is skipped for wing evaluations.

################################################

                if ($Sortie.Mission.Type -like "*Form*" -and -not $IsWingCandidate) {

                                try {

                                                $UsableOpenLines = 0

                                                $IsThereASingleShipOption = $false

                                                $CurrentETD = [datetime]$Sortie.ETD

                                                $TimeWindowEnd = $CurrentETD.AddMinutes($FormLookAhead)

 

                                                foreach ($FutureSortie in $Schedule) {

                                                                if ($FutureSortie -eq $Sortie) { continue }

                                                                if (-not [string]::IsNullOrWhiteSpace($FutureSortie.Student)) { continue }

                                                                if ($FutureSortie.SortieType -like "*Return*") { continue }

                                                                if ($FutureSortie.SortieType -like "*SIM*" -or $FutureSortie.SortieType -like "*OFT*") { continue }

                                                                if ($FutureSortie.SortieType -like "*O&B*") { continue }

 

                                                                $FutureETD = [datetime]$FutureSortie.ETD

                                                                if ($FutureETD -ge $CurrentETD -and $FutureETD -le $TimeWindowEnd) {

                                                                                $UsableOpenLines++

                                                                }

                                                }

 

                                                # No usable wing / DS slot exists inside the window → reject Form lead

                                                if ($UsableOpenLines -lt 1) { return $false }

 

                                                # Check whether any other student still has a Form-capable event

                                                $hasFormStudent = $false

                                                foreach ($otherStudent in $Students) {

                                                                if ([string]::IsNullOrWhiteSpace($otherStudent.Name) -or $otherStudent.Name -eq $Student.Name) {

                                                                                continue

                                                                }

 

                                                                $ValidEvents = Get-ValidOptedEvents -Student $otherStudent -Sortie $Sortie

                                                                if ($null -ne $ValidEvents) {

                                                                                foreach ($Event in $ValidEvents) {

                                                                                                if ($Event.Type -like "*Form*") {

                                                                                                                $hasFormStudent = $true

                                                                                                                break

                                                                                                }

                                                                                                elseif ($Event.Type -eq "Single" -or $Event.Type -eq "SingleSolo") {

                                                                                                                $IsThereASingleShipOption = $true

                                                                                                }

                                                                                }

                                                                }

                                                                if ($hasFormStudent) { break }

                                                }

 

                                                # Prefer a single-ship option when no Form wingman exists

                                                if (-not $hasFormStudent) {

                                                                if ($IsThereASingleShipOption) {

                                                                                return $false

                                                                } else {

                                                                                $Sortie.Notes += " | Direct Support Required"

                                                                }

                                                }

                                } catch {

                                                return $false

                                }

                }

 

 

 

 

 

 

 

 

 

 

 

###############################################

# BANS CAPsters FROM O&Bs

# - TESTED

###############################################

                if ($Student.IsCap -and $Sortie.SortieType -like "*O&B*") {

                                return $false

                }

 

 

 

##############################################

# CAPsters single turn ONLY (Flights)

# - TESTED

##############################################

                if ($Student.IsCAP) {

                                if ($Sortie.SortieType -notlike "*OFT*" -and $Sortie.SortieType -notlike "*IFT*") {

                                                $flightsToday = 0

                                                foreach ($event in $todaysEvents) {

                                                                if ($event.Instructor -notlike "*SIM IP*") {

                                                                                $flightsToday += 1

                                                                }

                                                }

                                                if ($flightsToday -ge 1) { return $false }

                                }

                }

 

 

 

###############################################

# CHECKRIDE EVALUATOR INDEPENDENCE

#

###############################################

                if ($Sortie.Mission.IsCheck) {

                                $CurrentIP_Base = $Sortie.Instructor.Name.Split("[")[0].Trim()

                                $PastIPs = @(

                                                $completedEvents | Where-Object {

                                                                -not [string]::IsNullOrWhiteSpace($_.Instructor) -and $_.Instructor -notlike "*SIM IP*"

                                                } | ForEach-Object {

                                                                $ipName = if ($_.Instructor -is [PSCustomObject]) { $_.Instructor.Name } else { [string]$_.Instructor }

                                                                $ipName.Split("[")[0].Trim()

                                                }

                                )

                                if ($PastIPs -contains $CurrentIP_Base) {

                                                return $false

                                }

                               

                                # Checkrides with only Check IPs

                                if (-not $Sortie.Instructor.IsCheck) {

                                                return $false

                                }

                }

 

 

 

##############################################

# HOLD Status

# - TESTED

##############################################

                if ($Student.OnHold) { return $false }

 

 

 

##############################################

# TRIPLE TURN PREVENTION

# - TESTED

##############################################

                if ($todaysEvents.Count -ge 2) {

                                return $false

                }

                elseif ($Sortie.SortieType -eq "O&B" -and $todaysEvents.Count -ge 1) {

                                return $false

                }

 

 

 

##############################################

# O&B RETURN LEG HARD RULES LOOKAHEAD

#

##############################################

                $ReturnStart = $null

                $ReturnMission = $null

                if ($Sortie.SortieType -eq "O&B") {

                                $TargetName = if ($Sortie.Instructor -is [PSCustomObject]) { $Sortie.Instructor.Name } else { $Sortie.Instructor }

                                $ReturnSortie = $Schedule | Where-Object {

                                                $_.SortieType -like "*Return*" -and

                                                [string]::IsNullOrWhiteSpace($_.Student) -and

                                                ([datetime]$_.ETD).Date -eq $NextStart.Date

                                } | Where-Object {

                                                $rowInst = if ($_.Instructor -is [PSCustomObject]) { $_.Instructor.Name } else { $_.Instructor }

                                                [string]$rowInst.Trim() -eq [string]$TargetName.Trim()

                                } | Select-Object -First 1

 

                                if ($null -ne $ReturnSortie) {

                                                $ReturnStart = Parse-ETD $ReturnSortie.ETD

                                }

 

                                $SimulatedMission = $Sortie.Mission

                                $SimulatedMission.IsComplete = $true

                                $NextEvents = Get-OptedEvents -SyllabusData $Student.Syllabus

                                if ($null -ne $NextEvents) {

                                                foreach ($e in $NextEvents) {

                                                                if ($e.Type -eq "Single" -or $e.Type -eq "Formation" -or $e.Type -eq "Both") {

                                                                                $ReturnMission = $e

                                                                                break

                                                                }

                                                }

                                }

                                $SimulatedMission.IsComplete = $false

 

                                # Ensure return leg isn't a checkride

                                if ($null -ne $ReturnMission -and $ReturnMission.IsCheck) {

                                                return $false

                                }

                }

 

 

 

######################################################

# TURN TIME, DUTY DAY, CREW REST

# Pretty Thoroughly Tested, Complex Rule so lots of room for error

######################################################

                $CrewRestCheck = $true

                $TurnTimeCheck = $true

                $DutyDayCheck  = $true

 

                if ($Sortie.SortieType -notlike "*Return*") {

                                foreach ($evt in $todaysEvents) {

                                                $evtDate = Parse-ETD $evt.ETD

                                                if ($null -ne $evtDate) {

                                                                $TurnTime = [Math]::Abs(($NextStart - $evtDate).TotalHours)

                                                               

                                                                if ($TurnTime -lt 3.25) {

                                                                                $TurnTimeCheck = $false

                                                                                break

                                                                }

                                                                elseif ($TurnTime -ge 3.25 -and $TurnTime -lt 3.5) {

                                                                                $TurnTimeCheck = $true

                                                                                $NewNotes += "Less than 3.5 hr TT, Flt/CC Approval Required"

                                                                }

                                                                elseif ($TurnTime -ge 6 -and $TurnTime -le 9) {

                                                                                $TurnTimeCheck = $true

                                                                                $NewNotes += "If this is a Night Flight SP will not have Duty Day"

                                                                }

                                                                elseif ($TurnTime -gt 9) {

                                                                                $DutyDayCheck = $false

                                                                                break

                                                                }

 

                                                                if ($null -ne $ReturnStart) {

                                                                                $ReturnTurnTime = [Math]::Abs(($ReturnStart - $evtDate).TotalHours)

                                                                                if ($ReturnTurnTime -lt 3.25) {

                                                                                                $TurnTimeCheck = $false

                                                                                                break

                                                                                }

                                                                                elseif ($ReturnTurnTime -ge 3.25 -and $ReturnTurnTime -lt 3.5) {

                                                                                                $TurnTimeCheck = $true

                                                                                                $NewNotes += "Return Leg: Less than 3.5 hr TT, Flt/CC Approval Required"

                                                                                }

                                                                                elseif ($ReturnTurnTime -ge 6 -and $ReturnTurnTime -le 9) {

                                                                                                $TurnTimeCheck = $true

                                                                                                $NewNotes += "Return Leg: Night Flight Warning"

                                                                                }

                                                                                elseif ($ReturnTurnTime -gt 9) {

                                                                                                $DutyDayCheck = $false

                                                                                                break

                                                                                }

                                                                }

                                                }

                                }

                }

 

                $PastEvents = @($completedEvents | Where-Object { (Parse-ETD $_.ETD).Date -lt $NextStart.Date })

                if ($PastEvents.Count -gt 0) {

                                $LastStart = Parse-ETD $PastEvents[0].ETD

                                if ($null -ne $LastStart -and $LastStart.Date -eq $NextStart.Date.AddDays(-1)) {

                                                $CrewRest = ($NextStart - $LastStart).TotalHours

                                                if ($CrewRest -lt 15) {

                                                                $CrewRestCheck = $false

                                                }

                                }

                }

 

                if (-not $CrewRestCheck) { return $false }

                if (-not $TurnTimeCheck) { return $false }

                if (-not $DutyDayCheck) { return $false }

 

 

 

###########################################################

# CAP/INTERNATIONAL AND EXP IPs

# - TESTED

###########################################################

                if ($Student.IsCAP -or $Student.IsInternational) { 

                                if ($Sortie.SortieType -like "*Day Local*" -or $Sortie.SortieType -like "*O&B*") {

                                                if (-not $Sortie.Instructor.IsExperienced) { return $false }

                                }

                }

 

 

 

##########################################################

# TRAINING FOLLOWING UNSAT EVENT

#

##########################################################

                if ($Student.UnsatSim -and ($Sortie.SortieType -like "*Day Local*" -or $Sortie.SortieType -like "*O&B*")) { return $false }

                if ($Student.UnsatFlight -and ($Sortie.SortieType -like "*IFT*" -or $Sortie.SortieType -like "*OFT*")) { return $false }

 

 

 

###########################################################

# SOLO FLIGHT CONTINUITY

# - TESTED

###########################################################

                if ($Sortie.SortieType -notlike "*OFT*" -and $Sortie.SortieType -notlike "*IFT*" -and -not [string]::IsNullOrWhiteSpace($Sortie.Instructor)) {

                               

                                $completionStatus = @{}

                                foreach ($r in $Student.Syllabus) { $completionStatus[$r.Mission] = ([bool]$r.IsComplete -or $r.IsComplete -match 'True') }

                               

                                $IsNextFlightSolo = $false

                                foreach ($r in ($Student.Syllabus | Sort-Object SyllabusID)) {

                                                if (-not $completionStatus[$r.Mission]) {

                                                                $prereqsMet = $true

                                                                if ($r.Prereqs -ne '0' -and -not [string]::IsNullOrWhiteSpace($r.Prereqs)) {

                                                                                foreach ($p in ($r.Prereqs -split ',')) {

                                                                                                if (-not $completionStatus[$p.Trim()]) { $prereqsMet = $false; break }

                                                                                }

                                                                }

                                                               

                                                                if ($prereqsMet -and ($r.Type -eq "Single" -or $r.Type -eq "SingleSolo" -or $r.Type -eq "Formation" -or $r.Type -eq "FormSolo" -or $r.Type -eq "Both")) {

                                                                                if ($r.Type -like "*Solo*") { $IsNextFlightSolo = $true }

                                                                                break

                                                                }

                                                }

                                }

 

                                if ($IsNextFlightSolo) {

                                                $NewNotes += "Solo Flight"

                                               

                                                $CurrentIP_Base = $Sortie.Instructor.Name.Split("[")[0].Trim()

                                               

                                                if ($CurrentIP_Base -notlike "*SIM IP*") {

                                                                $Last5Flights = @(

                                                                                $completedEvents | Where-Object {

                                                                                                # Safely get the instructor name whether it's an object or a string

                                                                                                $instName = if ($_.Instructor -is [PSCustomObject]) { $_.Instructor.Name } else { [string]$_.Instructor }

                                                                                                $instName -notlike "*SIM IP*" -and -not [string]::IsNullOrWhiteSpace($instName)

                                                                                } | Select-Object -First 5

                                                                )

                                                               

                                                                $IPFound = $false

                                                                foreach ($flight in $Last5Flights) {

                                                                                # Safely extract the string before attempting to Split

                                                                                $flightInstName = if ($flight.Instructor -is [PSCustomObject]) { $flight.Instructor.Name } else { [string]$flight.Instructor }

                                                                                $flightIP_Base = $flightInstName.Split("[")[0].Trim()

                                                                               

                                                                                if ($flightIP_Base -eq $CurrentIP_Base) {

                                                                                                $IPFound = $true

                                                                                                break

                                                                                }

                                                                }

                                                               

                                                                if (-not $IPFound) {

                                                                                $NewNotes += "Need IP in Continuity to Brief"

                                                                }

                                                }

                                }

                }

 

 

                if ($NewNotes.Count -gt 0) {

                                $joinedNotes = $NewNotes -join " | "

                                $Student | Add-Member -MemberType NoteProperty -Name "PendingNote" -Value $joinedNotes -Force

                } else {

                                $Student | Add-Member -MemberType NoteProperty -Name "PendingNote" -Value $null -Force

                }

               

                return $true

}

 

Export-ModuleMember -Function SoftRules, HardRules, Parse-ETD

 

 

 

 

function Get-DaySheetName {

                param(

                                [Parameter(Mandatory = $true)]

                                [datetime]$Date

                )

                # Returns the worksheet title in the form Wednesday. September 23, 2026

                return $Date.ToString("dddd. MMMM d, yyyy")

}

 

function Get-SortieRowKey {

                param(

                                [Parameter(Mandatory = $true)]

                                $Row

                )

                # Builds a composite key from the primary identifying fields so duplicate rows can be detected across runs

                $etd = if ($Row.ETD) { $Row.ETD.ToString() } else { "" }

                $inst = if ($Row.Instructor) { $Row.Instructor.ToString() } else { "" }

                $stype = if ($Row.SortieType) { $Row.SortieType.ToString() } else { "" }

                $stud = if ($Row.Student) { $Row.Student.ToString() } else { "" }

                $miss = if ($Row.Mission) { $Row.Mission.ToString() } else { "" }

                return "$etd|$inst|$stype|$stud|$miss"

}

 

function Merge-DayRows {

                param(

                                [array]$ExistingRows,

                                [array]$NewRows

                )

                # Combines existing sheet rows with newly generated rows, removes duplicates by key, then sorts chronologically by ETD

                $combined = @()

                if ($ExistingRows) { $combined += $ExistingRows }

                if ($NewRows) { $combined += $NewRows }

 

                $seen = @{}

                $unique = @()

                foreach ($row in $combined) {

                                $key = Get-SortieRowKey -Row $row

                                if (-not $seen.ContainsKey($key)) {

                                                $seen[$key] = $true

                                                $unique += $row

                                }

                }

 

                return $unique | Sort-Object {

                                $val = $_.ETD

                                if ($val -is [datetime]) { $val }

                                elseif ($val -as [datetime]) { [datetime]$val }

                                else { [datetime]::MinValue }

                }

}

 

function Export-DailyScheduleSheets {

                [CmdletBinding()]

                param(

                                [Parameter(Mandatory = $true)]

                                [array]$ScheduleData,

 

                                [string]$TargetFolder = "C:\Users\1595161645A\OneDrive - United States Air Force\ATIS - E Flight - E Flight",

 

                                [string]$FileName = "TESTING ATIS NOT OFFICIAL.xlsx"

                )

                # Ensures the destination folder exists

                if (-not (Test-Path -Path $TargetFolder -PathType Container)) {

                                New-Item -Path $TargetFolder -ItemType Directory -Force | Out-Null

                }

 

                # Forces the file name to end with .xlsx so Excel and Teams recognize the file type

                if (-not $FileName.EndsWith(".xlsx", [System.StringComparison]::OrdinalIgnoreCase)) {

                                $FileName = "$FileName.xlsx"

                }

 

                $fullPath = Join-Path -Path $TargetFolder -ChildPath $FileName

                Write-Host ("[DailyExport] Writing day sheets to {0}" -f $fullPath) -ForegroundColor Cyan

 

                # Creates a temporary local workbook path so OneDrive cannot interrupt the write

                $tempPath = Join-Path -Path $env:TEMP -ChildPath ("ATIS_DailySchedule_" + [guid]::NewGuid().ToString() + ".xlsx")

 

                try {

                                # If a previous version already exists in Teams/OneDrive, copy it to temp first so other sheets are preserved

                                if (Test-Path -Path $fullPath -PathType Leaf) {

                                                Copy-Item -Path $fullPath -Destination $tempPath -Force -ErrorAction Stop

                                }

 

                                # Groups the incoming schedule rows by calendar day

                                $byDay = @{}

                                foreach ($sortie in $ScheduleData) {

                                                $etdVal = $sortie.ETD

                                                $day = $null

                                                if ($etdVal -is [datetime]) {

                                                                $day = $etdVal.Date

                                                } elseif ($etdVal) {

                                                                try { $day = ([datetime]::Parse($etdVal.ToString())).Date } catch { continue }

                                                } else {

                                                                continue

                                                }

 

                                                $key = $day.ToString("yyyy-MM-dd")

                                                if (-not $byDay.ContainsKey($key)) {

                                                                $byDay[$key] = @()

                                                }

 

                                                $flat = [PSCustomObject]@{

                                                                ETD        = $sortie.ETD

                                                                Instructor = $sortie.Instructor

                                                                SortieType = $sortie.SortieType

                                                                Student    = $sortie.Student

                                                                Mission    = $sortie.Mission

                                                                Callsign   = $sortie.Callsign

                                                                Num        = $sortie.Num

                                                                Notes      = $sortie.Notes

                                                }

                                                $byDay[$key] += $flat

                                }

 

                                # For each day, merges with any pre-existing worksheet content and writes only that sheet

                                foreach ($dayKey in ($byDay.Keys | Sort-Object)) {

                                                $dayDate = [datetime]::ParseExact($dayKey, "yyyy-MM-dd", $null)

                                                $sheetName = Get-DaySheetName -Date $dayDate

                                                $newRows = $byDay[$dayKey]

 

                                                $existingRows = @()

                                                if (Test-Path -Path $tempPath -PathType Leaf) {

                                                                try {

                                                                                $existingRows = @(Import-Excel -Path $tempPath -WorksheetName $sheetName -ErrorAction Stop)

                                                                } catch {

                                                                                $existingRows = @()

                                                                }

                                                }

 

                                                $merged = Merge-DayRows -ExistingRows $existingRows -NewRows $newRows

 

                                                # Writes the merged rows to the named sheet; all other worksheets in the workbook are left unchanged

                                                $merged | Export-Excel -Path $tempPath -WorksheetName $sheetName -AutoSize

                                                Write-Host ("[DailyExport] Updated sheet '{0}' with {1} rows" -f $sheetName, $merged.Count) -ForegroundColor Green

                                }

 

                                # Copies the completed workbook to the OneDrive target in one operation

                                Copy-Item -Path $tempPath -Destination $fullPath -Force -ErrorAction Stop

                                Write-Host ("[DailyExport] Successfully updated {0}" -f $fullPath) -ForegroundColor Green

                }

                catch {

                                Write-Host ("[DailyExport] Failed to write schedule: {0}" -f $_.Exception.Message) -ForegroundColor Magenta

                }

                finally {

                                # Removes the temporary file

                                if (Test-Path -Path $tempPath) {

                                                Remove-Item -Path $tempPath -Force -ErrorAction SilentlyContinue

                                }

                }

}

 

 

 

Add-Type -AssemblyName System.Windows.Forms

Add-Type -AssemblyName System.Drawing

 

# Displays the primary looping menu

function Show-PreRunMenu {

                $Form = New-Object System.Windows.Forms.Form

                $Form.Text = "Pre-Run Operations Menu"

                $Form.Size = New-Object System.Drawing.Size(400, 300)

                $Form.StartPosition = "CenterScreen"

                $Form.Font = New-Object System.Drawing.Font("Segoe UI", 11)

 

                $Label = New-Object System.Windows.Forms.Label

                $Label.Text = "Select an action before running the scheduler:"

                $Label.Location = New-Object System.Drawing.Point(20, 20)

                $Label.AutoSize = $true

                $Form.Controls.Add($Label)

 

                $btnHook = New-Object System.Windows.Forms.Button

                $btnHook.Text = "Log Hooked Event (Unsat) TESTING"

                $btnHook.Location = New-Object System.Drawing.Point(20, 60)

                $btnHook.Size = New-Object System.Drawing.Size(340, 40)

                $btnHook.Add_Click({ $Form.Tag = "Hook"; $Form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $Form.Close() })

                $Form.Controls.Add($btnHook)

 

                $btnIncomplete = New-Object System.Windows.Forms.Button

                $btnIncomplete.Text = "Log Incomplete Event"

                $btnIncomplete.Location = New-Object System.Drawing.Point(20, 110)

                $btnIncomplete.Size = New-Object System.Drawing.Size(340, 40)

                $btnIncomplete.Add_Click({ $Form.Tag = "Incomplete"; $Form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $Form.Close() })

                $Form.Controls.Add($btnIncomplete)

 

                $btnRun = New-Object System.Windows.Forms.Button

                $btnRun.Text = "Ready to Run Program ->"

                $btnRun.Location = New-Object System.Drawing.Point(20, 180)

                $btnRun.Size = New-Object System.Drawing.Size(340, 40)

                $btnRun.BackColor = [System.Drawing.Color]::LightGreen

                $btnRun.Add_Click({ $Form.Tag = "Run"; $Form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $Form.Close() })

                $Form.Controls.Add($btnRun)

 

                $result = $Form.ShowDialog()

                if ($result -eq [System.Windows.Forms.DialogResult]::OK) { return $Form.Tag } else { return "Exit" }

}

 

 

 

 

# Displays the list of events to select from

function Show-EventSelectionMenu {

    param (

        $Schedule,

        $Title

    )

 

    $Form = New-Object System.Windows.Forms.Form

    $Form.Text = $Title

    $Form.Size = New-Object System.Drawing.Size(850, 500)

    $Form.StartPosition = "CenterScreen"

    $Form.Font = New-Object System.Drawing.Font("Segoe UI", 11)

 

    $HeaderLabel = New-Object System.Windows.Forms.Label

    $HeaderLabel.Text = "Select the scheduled flights/sims:"

    $HeaderLabel.Location = New-Object System.Drawing.Point(20, 20)

    $HeaderLabel.AutoSize = $true

    $Form.Controls.Add($HeaderLabel)

 

    $CheckedListBox = New-Object System.Windows.Forms.CheckedListBox

    $CheckedListBox.Location = New-Object System.Drawing.Point(20, 60)

    $CheckedListBox.Size = New-Object System.Drawing.Size(790, 320)

    $CheckedListBox.CheckOnClick = $true

    $CheckedListBox.Font = New-Object System.Drawing.Font("Consolas", 10)

 

    # Excel imports Student values as strings. Do not require a PSCustomObject.

    $ScheduledSorties = @(

        $Schedule |

            Where-Object {

                $null -ne $_.Student -and

                -not [string]::IsNullOrWhiteSpace([string]$_.Student)

            } |

            Sort-Object -Property {

                try {

                    [datetime]$_.ETD

                }

                catch {

                    [datetime]::MinValue

                }

            } -Descending

    )

 

    foreach ($Sortie in $ScheduledSorties) {

 

        # Supports Student as either an object with .Name or an Excel text value.

        $studentName = if (

            $Sortie.Student -isnot [string] -and

            $null -ne $Sortie.Student.PSObject.Properties['Name']

        ) {

            [string]$Sortie.Student.Name

        }

        else {

            [string]$Sortie.Student

        }

 

        # Supports Mission as either an object with .Mission or an Excel text value.

        $missionName = if (

            $null -ne $Sortie.Mission -and

            $Sortie.Mission -isnot [string] -and

            $null -ne $Sortie.Mission.PSObject.Properties['Mission']

        ) {

            [string]$Sortie.Mission.Mission

        }

        elseif ($null -ne $Sortie.Mission) {

            [string]$Sortie.Mission

        }

        else {

            'NONE'

        }

 

        # Supports Instructor as either an object with .Name or an Excel text value.

        $ipName = if (

            $null -ne $Sortie.Instructor -and

            $Sortie.Instructor -isnot [string] -and

            $null -ne $Sortie.Instructor.PSObject.Properties['Name']

        ) {

            [string]$Sortie.Instructor.Name

        }

        elseif (-not [string]::IsNullOrWhiteSpace([string]$Sortie.Instructor)) {

            [string]$Sortie.Instructor

        }

        else {

            'NO IP'

        }

 

        $displayString = "{0,-20} | {1,-30} | SP: {2,-15} | Mis: {3,-10} | {4}" -f `

            $Sortie.ETD,

            $ipName,

            $studentName,

            $missionName,

            $Sortie.SortieType

 

        $listItem = [PSCustomObject]@{

            Display = $displayString

            Sortie  = $Sortie

        }

 

        [void]$CheckedListBox.Items.Add($listItem)

    }

 

    $CheckedListBox.DisplayMember = "Display"

    $Form.Controls.Add($CheckedListBox)

 

    $SubmitBtn = New-Object System.Windows.Forms.Button

    $SubmitBtn.Location = New-Object System.Drawing.Point(20, 400)

    $SubmitBtn.Size = New-Object System.Drawing.Size(200, 40)

    $SubmitBtn.Text = "Confirm Selection"

    $SubmitBtn.BackColor = [System.Drawing.Color]::LightCoral

    $SubmitBtn.Add_Click({

        if ($CheckedListBox.CheckedItems.Count -gt 0) {

            $Form.Tag = $CheckedListBox.CheckedItems

            $Form.DialogResult = [System.Windows.Forms.DialogResult]::OK

            $Form.Close()

        }

        else {

            [System.Windows.Forms.MessageBox]::Show(

                "Please select at least one event.",

                "None Selected",

                [System.Windows.Forms.MessageBoxButtons]::OK,

                [System.Windows.Forms.MessageBoxIcon]::Warning

            )

        }

    })

    $Form.Controls.Add($SubmitBtn)

 

    $ExitBtn = New-Object System.Windows.Forms.Button

    $ExitBtn.Location = New-Object System.Drawing.Point(610, 400)

    $ExitBtn.Size = New-Object System.Drawing.Size(200, 40)

    $ExitBtn.Text = "Cancel"

    $ExitBtn.BackColor = [System.Drawing.Color]::LightGray

    $ExitBtn.Add_Click({

        $Form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

        $Form.Close()

    })

    $Form.Controls.Add($ExitBtn)

 

    $result = $Form.ShowDialog()

 

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {

        return $Form.Tag

    }

 

    return $null

}

 

 

 

 

function Update-StudentSimDebtFromSyllabus {

    param(

        [Parameter(Mandatory)]

        $Student,

 

        [Parameter(Mandatory)]

        $SoftWeights

    )

 

    # Do not change debt if the required configuration is unavailable.

    if ($null -eq $Student -or $null -eq $Student.Syllabus) {

        Write-Warning "Cannot recalculate SimDebt: student or syllabus is missing."

        return $null

    }

 

    foreach ($propertyName in @('SimStopMult', 'TransSimRatio', 'FormSimRatio')) {

        if ($null -eq $SoftWeights.$propertyName -or $SoftWeights.$propertyName -eq '') {

            Write-Warning "Cannot recalculate SimDebt for '$($Student.Name)': SoftWeights.$propertyName is missing."

            return $null

        }

    }

 

    $simStopMult  = [double]$SoftWeights.SimStopMult

    $transSimRatio = [double]$SoftWeights.TransSimRatio

    $formSimRatio  = [double]$SoftWeights.FormSimRatio

 

    $debt = 0.0

 

    foreach ($mission in $Student.Syllabus) {

        # Only completed missions contribute to current SimDebt.

        $isComplete = ($mission.IsComplete -eq $true -or $mission.IsComplete -eq 'True')

 

        if (-not $isComplete) {

            continue

        }

 

        $missionName = [string]$mission.Mission

 

        if ([string]::IsNullOrWhiteSpace($missionName)) {

            continue

        }

 

        # Flights increase SimDebt.

        if (

            $missionName -like '*FT5*' -or

            $missionName -like '*FF5*' -or

            $missionName -like '*AA5*'

        ) {

            $debt += $simStopMult

        }

 

        # Transition sims decrease SimDebt.

        elseif (

            $missionName -like '*FI3*' -or

            $missionName -like '*FE3*' -or

            $missionName -like '*FI4*' -or

            $missionName -like '*FE4*'

        ) {

            $debt -= $transSimRatio

        }

 

        # Formation sims decrease SimDebt.

        elseif (

            $missionName -like '*FN3*' -or

            $missionName -like '*FN4*' -or

            $missionName -like '*AA4*'

        ) {

            $debt -= $formSimRatio

        }

    }

 

    # Ensure the student object has a SimDebt property.

    $Student |

        Add-Member -MemberType NoteProperty -Name 'SimDebt' -Value ([math]::Round($debt, 2)) -Force

 

    return $Student.SimDebt

}

 

 

 

 

# Processes Hooked logic

# Processes Hooked (Unsat) logic against the live student objects

function Process-HookedEvents {

                param (

                                [Parameter(Mandatory)]

                                $SelectedItems,

 

                                [Parameter(Mandatory)]

                                $Students

                )

 

                $updatedCount = 0

 

                foreach ($Item in $SelectedItems) {

                                $Sortie = $Item.Sortie

 

                                if ($null -eq $Sortie -or $null -eq $Sortie.Student) {

                                                Write-Host "Skipped a hooked event because it has no assigned student." -ForegroundColor Yellow

                                                continue

                                }

 

                                # Supports Student as either an object with .Name or an Excel text value.

                                $studentName = if (

                                                $Sortie.Student -isnot [string] -and

                                                $null -ne $Sortie.Student.PSObject.Properties['Name']

                                ) {

                                                [string]$Sortie.Student.Name

                                }

                                else {

                                                [string]$Sortie.Student

                                }

 

                                $matchedStudent = $Students |

                                                Where-Object { $_.Name -eq $studentName } |

                                                Select-Object -First 1

 

                                if ($null -eq $matchedStudent) {

                                                Write-Host "Could not find student '$studentName' in the student list. Unsat flag was not set." -ForegroundColor Yellow

                                                continue

                                }

 

                                # Sim unsat blocks later flights; flight unsat blocks later sims.

                                if ($Sortie.SortieType -like "*OFT*" -or $Sortie.SortieType -like "*IFT*") {

                                                $matchedStudent.UnsatSim = $true

                                                Write-Host "Logged UnsatSim for $($matchedStudent.Name) (sortie type: $($Sortie.SortieType))." -ForegroundColor Cyan

                                }

                                else {

                                                $matchedStudent.UnsatFlight = $true

                                                Write-Host "Logged UnsatFlight for $($matchedStudent.Name) (sortie type: $($Sortie.SortieType))." -ForegroundColor Cyan

                                }

 

                                # Annotate the schedule row so the unsat is visible on output.

                                if ([string]::IsNullOrWhiteSpace([string]$Sortie.Notes)) {

                                                $Sortie.Notes = "Event Hooked (Unsat)"

                                }

                                else {

                                                $Sortie.Notes = "$($Sortie.Notes) | Event Hooked (Unsat)"

                                }

 

                                $updatedCount++

                }

 

                return $updatedCount

}

 

 

 

 

# Processes Incomplete logic

function Process-IncompleteEvents {

    param(

        [Parameter(Mandatory)]

        $SelectedItems,

 

        [Parameter(Mandatory)]

        $Students,

 

        [Parameter(Mandatory)]

        $SoftWeights

    )

 

 

                if ($null -eq $SoftWeights) {

                                Write-Warning "SoftWeights was not provided. Incomplete events can be processed, but SimDebt will not be recalculated or saved."

                }

 

    # Returns the student objects whose SimDebt was recalculated.

    $updatedStudents = @()

 

    foreach ($Item in $SelectedItems) {

        $Sortie = $Item.Sortie

 

        if ($null -eq $Sortie -or $null -eq $Sortie.Student) {

            Write-Warning "Skipped an incomplete event because it has no assigned student."

            continue

        }

 

        # Supports either objects from the scheduler or text imported from Excel.

        $studentName = if (

            $Sortie.Student -isnot [string] -and

            $null -ne $Sortie.Student.PSObject.Properties['Name']

        ) {

            [string]$Sortie.Student.Name

        }

        else {

            [string]$Sortie.Student

        }

 

        $missionName = if (

            $null -ne $Sortie.Mission -and

            $Sortie.Mission -isnot [string] -and

            $null -ne $Sortie.Mission.PSObject.Properties['Mission']

        ) {

            [string]$Sortie.Mission.Mission

        }

        elseif ($null -ne $Sortie.Mission) {

            [string]$Sortie.Mission

        }

        else {

            $null

        }

 

        $matchedStudent = $Students |

            Where-Object { $_.Name -eq $studentName } |

            Select-Object -First 1

 

        $syllabusUpdated = $false

 

        if ($null -eq $matchedStudent) {

            Write-Warning "Could not find student '$studentName' in the student list."

        }

        elseif ([string]::IsNullOrWhiteSpace($missionName)) {

            Write-Warning "Could not identify a mission for student '$studentName'."

        }

        else {

            foreach ($syllabusMission in $matchedStudent.Syllabus) {

                if ($syllabusMission.Mission -eq $missionName) {

                    $syllabusMission.IsComplete = $false

                    $syllabusMission.Instructor = ''

                    $syllabusMission.ETD = ''

 

                    $syllabusUpdated = $true

                    break

                }

            }

 

            if ($syllabusUpdated) {

                                if ($null -ne $SoftWeights) {

                                                $newDebt = Update-StudentSimDebtFromSyllabus `

                                               -Student $matchedStudent `

                                               -SoftWeights $SoftWeights

 

                                                if ($null -ne $newDebt) {

                                               Write-Host "Updated SimDebt for $($matchedStudent.Name): $newDebt" -ForegroundColor Cyan

                                               $updatedStudents += $matchedStudent

                                                }

                                }

            }

            else {

                Write-Warning "Mission '$missionName' was not found in $studentName's syllabus. SimDebt was not changed."

            }

        }

 

        # Preserve the existing schedule update behavior.

        $Sortie.Student = $null

        $Sortie.Mission = $null

 

        if ([string]::IsNullOrWhiteSpace([string]$Sortie.Notes)) {

            $Sortie.Notes = 'Incomplete'

        }

        else {

            $Sortie.Notes = "$($Sortie.Notes) | Incomplete"

        }

    }

 

    return $updatedStudents

}

 

 

 

 

# Encapsulates the entire loop so the main code stays extremely clean

function Invoke-PreRunMenu {

    param (

        $Schedule,

        $Students,

        $SoftWeights,

        [string]$MasterDataPath

    )

                $KeepRunning = $true

                while ($KeepRunning) {

                                $UserChoice = Show-PreRunMenu

                                                                if ($UserChoice -eq "Hook") {

                                                $SelectedItems = Show-EventSelectionMenu -Schedule $Schedule -Title "Select Hooked Events"

                                                if ($null -ne $SelectedItems) {

                                                                $hookedCount = Process-HookedEvents -SelectedItems $SelectedItems -Students $Students

                                                                if ($hookedCount -gt 0) {

                                                                                [System.Windows.Forms.MessageBox]::Show(

                                                                                                "Hooked events recorded successfully ($hookedCount).",

                                                                                                "Success",

                                                                                                [System.Windows.Forms.MessageBoxButtons]::OK,

                                                                                                [System.Windows.Forms.MessageBoxIcon]::Information

                                                                                )

                                                                }

                                                                else {

                                                                                [System.Windows.Forms.MessageBox]::Show(

                                                                                                "No hooked events were applied. Check the console for details.",

                                                                                                "Nothing Updated",

                                                                                                [System.Windows.Forms.MessageBoxButtons]::OK,

                                                                                                [System.Windows.Forms.MessageBoxIcon]::Warning

                                                                                )

                                                                }

                                                }

                                } elseif ($UserChoice -eq "Incomplete") {

                                                $SelectedItems = Show-EventSelectionMenu -Schedule $Schedule -Title "Select Incomplete Events"

 

                                                if ($null -ne $SelectedItems) {

                                               $updatedStudents = @(

                                                               Process-IncompleteEvents `

                                                               -SelectedItems $SelectedItems `

                                                               -Students $Students `

                                                               -SoftWeights $SoftWeights

                                               )

 

                               # Persist SimDebt only for students changed during this selection.

                               if ($updatedStudents.Count -gt 0) {

                               $simDebtByName = @{}

 

                               foreach ($student in $updatedStudents) {

                                               if ( -not [string]::IsNullOrWhiteSpace([string]$student.Name) -and $null -ne $student.SimDebt) {

                                                              $simDebtByName[$student.Name] = [double]$student.SimDebt

                                               }

                               }

 

                               if ($simDebtByName.Count -gt 0) {

                                               if (-not [string]::IsNullOrWhiteSpace($MasterDataPath) -and (Test-Path -LiteralPath $MasterDataPath) ) {

                                                              try {

                                                              Update-MasterDataSimDebtColumnP `

                                                              -SimDebtByName $simDebtByName `

                                                              -MasterDataPath $MasterDataPath

                                                              }

                                                              catch {

                                                              Write-Warning "SimDebt changed in memory but could not be saved to MasterData.xlsx: $($_.Exception.Message)"

                                                              }

                                               }

                               else {

                                              Write-Warning "SimDebt changed in memory but was not saved because MasterDataPath is unavailable."

                               }

               }

        }

 

        [System.Windows.Forms.MessageBox]::Show(

            "Incomplete events processed and SimDebt recalculated.",

            "Success",

            [System.Windows.Forms.MessageBoxButtons]::OK,

            [System.Windows.Forms.MessageBoxIcon]::Information

        )

    }

                } elseif ($UserChoice -eq "Run") {

                                                $KeepRunning = $false

                                } else {

                                                Write-Host "Exiting Program..." -ForegroundColor Yellow

                                                exit

                                }

                }

}

 

Export-ModuleMember -Function Invoke-PreRunMenu

 

 

 

##

#Ingest Schedule and Output Schedule

##

#Ingest Schedule and Output Schedule

function IngestSchedule {

 

                $filePath = Join-Path (Split-Path $PSScriptRoot -Parent) "MasterData.xlsx"

 

                # -NoHeader + direct P1/P2... = completely ignores any pasted headers, duplicate headers, or formatting

                $rawSchedule = Import-Excel -Path $filePath -WorksheetName "Schedule" -NoHeader

 

                $CleanSchedule = @() # Create an empty list to hold the good data

 

                foreach ($row in $rawSchedule) {

 

                                # Ignore any row whose first column contains a header name (handles users pasting headers or duplicate headers)

                                if ($row.P1 -in @("ETD", "Date", "Callsign", "Num", "Instructor", "Student", "Mission", "SortieType", "Notes", "Sortie Type", "Event Name", "#", "Remark", "Device", "Event Name", "Simulator Instructor", "Start Time" )) { continue }

 

                                # Build a clean object using fixed column positions (P1 = ETD, P7 = Num, etc.)

                                $cleanRow = [PSCustomObject]@{

                                                ETD         = $row.P1

                                                Instructor  = $row.P2

                                                SortieType  = $row.P3

                                                Student     = $row.P4

                                                Mission     = $row.P5

                                                Callsign    = $row.P6

                                                Num         = $row.P7

                                                Notes       = $row.P8

                                }

 

                                # Replace 'Date' with a column name that should ALWAYS have text in a real row

                                if ([string]::IsNullOrWhiteSpace($cleanRow.ETD)) {

                                                # As soon as it hits a row where the Date is blank, it STOPS pulling data

                                                break

                                }

 

                                # --- FIX: FORMAT EXCEL DATE TO READABLE STRING ---

                                if ($cleanRow.ETD -match '^\d+(\.\d+)?$') {

                                                # Convert raw decimal to a formatted date string

                                                $cleanRow.ETD = [datetime]::FromOADate([double]$cleanRow.ETD).ToString("yyyy-MM-dd HH:mm:ss")

 

                                } else {

                                                # If it's a normal string, standardize its format just in case

                                                $parsed = $cleanRow.ETD -as [datetime]

                                                if ($parsed) {

                                                                $cleanRow.ETD = $parsed.ToString("yyyy-MM-dd HH:mm:ss")

                                                }

                                }

 

                                # For OFT/IFT, default Instructor to "SIM IP" only when the input left the Instructor column blank.

                                # If an instructor name was provided, pass it through unchanged so it appears on the output schedule.

                                if (($cleanRow.SortieType -like "*OFT*" -or $cleanRow.SortieType -like "*IFT*") -and [string]::IsNullOrWhiteSpace($cleanRow.Instructor)) {

                                                $cleanRow.Instructor = "SIM IP"

                                }

 

                                # If the row isn't blank, add it to our clean list

                                $CleanSchedule += $cleanRow

                }

 

                # Sort the clean array by the ETD property (earliest to latest) before returning

                return $CleanSchedule | Sort-Object -Property ETD

}

 

 

 

 

 

 

 

 

 

 

function Get-MatchedInstructor {

                param (

                                [array]$Instructors,

                                [string]$TargetName

                )

 

                if ([string]::IsNullOrWhiteSpace($TargetName)) {

                                return $null

                }

 

                try {

                                $BaseTarget = $TargetName -replace '\s*\[.*?\]', ''

                                $BaseTarget = $BaseTarget.Trim()

 

                                $Match = $Instructors | Where-Object {

                                                $BaseSource = $_.Name -replace '\s*\[.*?\]', ''

                                                $BaseSource = $BaseSource.Trim()

                                                $BaseSource -match [regex]::Escape($BaseTarget)

                                } | Select-Object -First 1

 

                                return $Match

                } catch {

                                Write-Warning "Error matching instructor: $($_.Exception.Message)"

                                return $null

                }

}

 

 

 

function OandB_Screener {

                param (

                                $Student,

                                $Sortie,

                                $NewNotes,

                                $bannedSorties

                )

                foreach ($sortie in $BannedFirstSorties){

                                if($Sortie.Mission -eq $sortie){return $false}

                }

                if($Sortie.Mission.IsCheck){return $false}

                return $true

 

                $completedEvents = @(

                                $Student.Syllabus | Where-Object {

                                                ($_.IsComplete -eq $true -or $_.IsComplete -eq 'True') -and

                                                -not [string]::IsNullOrWhiteSpace($_.ETD)

                                } | Sort-Object -Property { Parse-ETD $_.ETD } -Descending

                )

 

                $todaysEvents = @(

                                $completedEvents | Where-Object {

                                                $evtDate = Parse-ETD $_.ETD

                                                $null -ne $evtDate -and $evtDate.Date -eq $NextStart.Date

                                }

                )

               

                # prevents accidental triple turning

                if($todaysEvents.Count -gt 0) {return $false}

}

 

 

 

 

 

 

 

 

# Places the return leg logic into a distinct function

function Schedule-ReturnLeg {

                param (

                                $Sortie,

                                $BestMatch,

                                $Schedule,

                                $Students

                )

 

                # 1. Cleanly grab the Target Instructor and Date

                $TargetName = $Sortie.Instructor.Name

                $SortieDate = [datetime]$Sortie.ETD

 

                # 2. Use Where-Object to safely find the correct Return Leg

                $ReturnSortie = $Schedule | Where-Object {

                                # Must be a return leg, unassigned, and on the SAME DAY

                                $_.SortieType -like "*Return*" -and

                                [string]::IsNullOrWhiteSpace($_.Student) -and

                                ([datetime]$_.ETD).Date -eq $SortieDate.Date

                } | Where-Object {

                                # Safely extract the instructor name whether it's an object or a raw string from Excel

                                $rowInst = if ($_.Instructor -is [PSCustomObject]) { $_.Instructor.Name } else { $_.Instructor }

                                [string]$rowInst.Trim() -eq [string]$TargetName.Trim()

                } | Select-Object -First 1

               

                # Ensures a valid return sortie was found before attempting to map data

                if ($null -ne $ReturnSortie) {

                                # Select-Object returns a single item, no need for [0] indexing

                                $OBReturn = $ReturnSortie

                                $OBReturn.Student = $BestMatch

                               

                                # Adds the Callsign to the spreadsheet

                                $OBReturn.Callsign = $Sortie.Callsign

                                $OBReturn.Num = $Sortie.Num

 

                                $ValidEvents = Get-ValidOptedEvents -Student $BestMatch -Sortie $OBReturn

                               

                                if ($null -ne $ValidEvents) {

                                                $BestMatch | Add-Member -MemberType NoteProperty -Name "OptedEvents" -Value $ValidEvents -Force

 

                                                $SelectedMission = $BestMatch.OptedEvents | Sort-Object -Property SyllabusID | Select-Object -First 1

                                                $OBReturn.Mission = $SelectedMission

 

                                                foreach ($Student in $Students) {

                                                                if ($Student.Name -eq $BestMatch.Name) {

                                                                                foreach ($m in $Student.Syllabus) {

                                                                                                if ($m.Mission -eq $OBReturn.Mission.Mission) {

                                                                                                                $m.IsComplete = $true

                                                                                                                $m.Instructor = $OBReturn.Instructor

                                                                                                                $m.ETD = $OBReturn.ETD

                                                                                                                break

                                                                                                }

                                                                                }

                                                                                break

                                                                }

                                                }

                                } else {

                                                $OBReturn.Mission = $null

                                                Add-SortieNote -Sortie $OBReturn -Note "No valid mission for return leg"

                                }

 

 

                                # Select the mission in which the SP is going to fly, based on priority, aka Syllabus ID

                                $SelectedMission = $BestMatch.OptedEvents | Sort-Object -Property SyllabusID | Select-Object -First 1

                                $OBReturn.Mission = $SelectedMission

 

                                # Updates the computer's local copy of the Student's syllabus

                                foreach ($Student in $Students) {

                                                if ($Student.Name -eq $BestMatch.Name) {

                                                                foreach ($m in $Student.Syllabus) {

                                                                                if ($m.Mission -eq $OBReturn.Mission.Mission) {

                                                                                                $m.IsComplete = $true

                                                                                                $m.Instructor = $OBReturn.Instructor

                                                                                                $m.ETD = $OBReturn.ETD

                                                                                                break

                                                                                }

                                                                }

                                                                break

                                                }

                                }

                } else {

                                Write-Warning "Could not find an empty O&B Return row on $($SortieDate.ToString('yyyy-MM-dd')) for $TargetName."

                }

               

                $ValidEvents = $null

}

 

 

 

 

 

 

 

 

 

function Get-NormalizedLastName {

                param (

                                [string]$Name

                )

                # Returns only the last-name portion of a name string (everything before the first comma)

                if ([string]::IsNullOrWhiteSpace($Name)) {

                                return ""

                }

                $normalized = $Name.Trim()

                $commaIndex = $normalized.IndexOf(',')

                if ($commaIndex -ge 0) {

                                $normalized = $normalized.Substring(0, $commaIndex).Trim()

                }

                return $normalized

}

 

 

 

 

 

 

 

function Process-HardcodedSortie {

                param (

                                [Parameter(Mandatory=$true)] $Sortie,

                                [Parameter(Mandatory=$true)] $SortieInstructor,

                                [Parameter(Mandatory=$true)] [array]$Students,

                                [Parameter(Mandatory=$true)] [array]$Schedule,

                                [Parameter(Mandatory=$true)] [array]$Instructors,

                                $SoftWeights,

                                $Formations,

                                [int]$FormIndex = 0

                )

                # 1. Resolve Student Name and locate their object in $Students

                $rawStudentName = if ($Sortie.Student -is [PSCustomObject]) { $Sortie.Student.Name } else { [string]$Sortie.Student }

                $normalizedRawName = Get-NormalizedLastName -Name $rawStudentName

                $matchedStudent = $Students | Where-Object {

                                (Get-NormalizedLastName -Name $_.Name).ToLower() -eq $normalizedRawName.ToLower()

                } | Select-Object -First 1

                if ($null -ne $matchedStudent) {

                                # Link the full student object

                                $Sortie.Student = $matchedStudent

                                # 2. Resolve Mission (from row or auto-find top valid opted event if blank)

                                $rawMissionName = if ($Sortie.Mission -is [PSCustomObject]) { $Sortie.Mission.Mission } else { [string]$Sortie.Mission }

                                $matchedMission = $null

                                if (-not [string]::IsNullOrWhiteSpace($rawMissionName)) {

                                                $matchedMission = $matchedStudent.Syllabus | Where-Object { $_.Mission.Trim() -eq $rawMissionName.Trim() } | Select-Object -First 1

                                } else {

                                                $validEvents = Get-ValidOptedEvents -Student $matchedStudent -Sortie $Sortie

                                                if ($null -ne $validEvents) {

                                                                $matchedMission = $validEvents | Sort-Object -Property SyllabusID | Select-Object -First 1

                                                }

                                }

                                # 3. Update the student's syllabus event in memory

                                if ($null -ne $matchedMission) {

                                                $Sortie.Mission = $matchedMission

                                                foreach ($m in $matchedStudent.Syllabus) {

                                                                if ($m.Mission -eq $matchedMission.Mission) {

                                                                                $m.IsComplete = $true

                                                                                $m.Instructor = if ($SortieInstructor -is [PSCustomObject]) { $SortieInstructor.Name } else { $SortieInstructor }

                                                                                $m.ETD = $Sortie.ETD

                                                                                break

                                                                }

                                                }

                                                Write-Host " [BYPASS] Scheduled '$($matchedStudent.Name)' on '$($matchedMission.Mission)' (ETD: $($Sortie.ETD)) - Syllabus Updated." -ForegroundColor Cyan

                                } else {

                                                Write-Warning " [BYPASS] Student '$($matchedStudent.Name)' has no valid mission matching '$rawMissionName' on ETD $($Sortie.ETD)."

                                }

                                # 4. Ensure Instructor and Callsigns are populated

                                if (-not ($Sortie.Instructor -is [PSCustomObject])) {

                                                $Sortie.Instructor = $SortieInstructor

                                }

                                if ([string]::IsNullOrWhiteSpace($Sortie.Callsign)) {

                                                $Sortie.Callsign = $SortieInstructor.Callsign

                                                $Sortie.Num = $SortieInstructor.Num

                                }

 

 

                                # 5. Handle Formation Pairing only for true leads.

                                #    Skip if this line is already a wingman (Notes already contain "Wingman to").

                                $alreadyWingman = $false

                                if (-not [string]::IsNullOrWhiteSpace($Sortie.Notes) -and $Sortie.Notes -like "*Wingman to*") {

                                                $alreadyWingman = $true

                                }

 

                                if ($null -ne $matchedMission -and $matchedMission.Type -like "*Form*" -and -not $alreadyWingman) {

                                                $FormIndex = Pair-FormationWingman -PrimarySortie $Sortie `

                                                                                   -PrimaryStudent $matchedStudent `

                                                                                   -PrimaryMission $matchedMission `

                                                                                   -Schedule $Schedule `

                                                                                   -Students $Students `

                                                                                   -Instructors $Instructors `

                                                                                   -SoftWeights $SoftWeights `

                                                                                   -FormLookAhead $SoftWeights.FormLookAhead `

                                                                                   -LevelDifference $SoftWeights.LevelDifference `

                                                                                   -FormIndex $FormIndex `

                                                                                   -Formations $Formations

                                }

 

 

                                # 6. Handle O&B Return Leg if applicable

                                if ($Sortie.SortieType -eq "O&B" -and $null -ne $matchedMission) {

                                                Schedule-ReturnLeg -Sortie $Sortie -BestMatch $matchedStudent -Schedule $Schedule -Students $Students

                                }

                } else {

                                Write-Warning " [BYPASS] Student '$rawStudentName' on ETD $($Sortie.ETD) was not found in MasterData."

                }

                # Return updated FormIndex so formation callsigns remain in sequence

                return $FormIndex

}

 

 

 

 

 

 

 

 

 

 

# Safely appends notes to a Sortie without overwriting or leaving ugly formatting

# Add-SortieNote -Sortie $Sortie -Note "O&B Leg 1 Not Filled"

function Add-SortieNote {

                param (

                                $Sortie,

                                [string]$Note

                )

               

                if ([string]::IsNullOrWhiteSpace($Note)) { return }

               

                if ([string]::IsNullOrWhiteSpace($Sortie.Notes)) {

                                $Sortie.Notes = $Note

                } else {

                                $Sortie.Notes += " | $Note"

                }

}

 

 

 

 

 

 

 

function IngestStudentAbsences {

 

    $filePath = Join-Path (Split-Path $PSScriptRoot -Parent) "MasterData.xlsx"

 

    try {

        $AbsenceData = Import-Excel -Path $filePath -WorksheetName "StudentAbsences" -ErrorAction Stop

    } catch {

        Write-Warning "Could not read the 'StudentAbsences' sheet."

        return @()

    }

 

    $CleanAbsences = @()

 

    foreach ($row in $AbsenceData) {

 

        $studentName = if ($row.Student) { ([string]$row.Student).Trim() } else { "" }

        $isActiveRaw = $row.IsActive

        $startRaw    = $row.StartTime

        $endRaw      = $row.EndTime

 

        # === Convert to proper DateTime objects right away ===

        $start = $null

        $end   = $null

 

        if ($startRaw -is [datetime]) {

            $start = $startRaw

        } elseif ([string]$startRaw -match '^\d+(\.\d+)?$') {

            $start = [datetime]::FromOADate([double]$startRaw)

        } else {

            $start = $startRaw -as [datetime]

        }

 

        if ($endRaw -is [datetime]) {

            $end = $endRaw

        } elseif ([string]$endRaw -match '^\d+(\.\d+)?$') {

            $end = [datetime]::FromOADate([double]$endRaw)

        } else {

            $end = $endRaw -as [datetime]

        }

 

        if ([string]::IsNullOrWhiteSpace($studentName)) {

            continue

        }

 

        # IsActive check

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

            #Write-Host "  -> SKIPPED: IsActive is not active"

            continue

        }

 

        if ($null -eq $start -or $null -eq $end) {

            #Write-Host "  -> SKIPPED: Could not parse dates into real DateTime objects"

            continue

        }

 

        #Write-Host "  -> KEPT as active absence" -ForegroundColor Green

 

        $CleanAbsences += [PSCustomObject]@{

            Student   = $studentName

            StartTime = $start          # Now a real DateTime object

            EndTime   = $end            # Now a real DateTime object

            Reason    = if ($row.Reason) { ([string]$row.Reason).Trim() } else { "" }

        }

    }

 

    return $CleanAbsences

}

 

 

 

 

 

 

# ==========================================

#      HELPER: Classify sortie as Flight or Sim (no regex)

# ==========================================

function Get-SortieCategory {

                [CmdletBinding()]

                param(

                                [Parameter(Mandatory = $true, ValueFromPipeline = $true)]

                                [string]$SortieType

                )

                process {

                                if ([string]::IsNullOrWhiteSpace($SortieType)) {

                                                return 'Unknown'

                                }

                                # Simple -like checks only (consistent with existing code style)

                                if ($SortieType -like '*OFT*' -or $SortieType -like '*IFT*') {

                                                return 'Sim'

                                } else {

                                                return 'Flight'

                                }

                }

}

 

# ==========================================

#      HELPER: Draw thick black border ONLY around the outside of the whole table

# ==========================================

function Apply-ThickBorders {

                [CmdletBinding()]

                param(

                                [Parameter(Mandatory = $true)]

                                $Worksheet

                )

                if ($null -eq $Worksheet -or $null -eq $Worksheet.Dimension) {

                                return

                }

 

                $startRow = $Worksheet.Dimension.Start.Row

                $endRow   = $Worksheet.Dimension.End.Row

                $startCol = $Worksheet.Dimension.Start.Column

                $endCol   = $Worksheet.Dimension.End.Column

 

                $thick = [OfficeOpenXml.Style.ExcelBorderStyle]::Thick

                $black = [System.Drawing.Color]::Black

 

                # Top edge of header row

                for ($col = $startCol; $col -le $endCol; $col++) {

                                $cell = $Worksheet.Cells[$startRow, $col]

                                $cell.Style.Border.Top.Style = $thick

                                $cell.Style.Border.Top.Color.SetColor($black)

                }

 

                # Bottom edge of last data row

                for ($col = $startCol; $col -le $endCol; $col++) {

                                $cell = $Worksheet.Cells[$endRow, $col]

                                $cell.Style.Border.Bottom.Style = $thick

                                $cell.Style.Border.Bottom.Color.SetColor($black)

                }

 

                # Left edge of first column

                for ($row = $startRow; $row -le $endRow; $row++) {

                                $cell = $Worksheet.Cells[$row, $startCol]

                                $cell.Style.Border.Left.Style = $thick

                                $cell.Style.Border.Left.Color.SetColor($black)

                }

 

                # Right edge of last column

                for ($row = $startRow; $row -le $endRow; $row++) {

                                $cell = $Worksheet.Cells[$row, $endCol]

                                $cell.Style.Border.Right.Style = $thick

                                $cell.Style.Border.Right.Color.SetColor($black)

                }

}

 

# ==========================================

#      HELPER: Export MasterSchedule sorted Day → Flight/Sim → Time with group borders

# ==========================================

function Export-FormattedMasterSchedule {

                [CmdletBinding()]

                param(

                                [Parameter(Mandatory = $true)]

                                [array]$ScheduleData,

                                [Parameter(Mandatory = $true)]

                                [string]$FilePath,

                                [Parameter(Mandatory = $true)]

                                [string]$WorksheetName = 'MasterSchedule'

                )

 

                Write-Host ("[Export] Preparing {0} sorties for export..." -f $ScheduleData.Count) -ForegroundColor Cyan

 

                # Enrich with Day (used only for sorting + internal borders) and Category

                $enriched = $ScheduleData | Select-Object *,

                                @{

                                                Name = 'Day'

                                                Expression = {

                                                                $val = $_.ETD

                                                                if ($val -is [datetime]) {

                                                                                return $val.Date

                                                                } elseif ($val) {

                                                                                try { return ([datetime]::Parse($val)).Date } catch { return [datetime]::MinValue }

                                                                } else {

                                                                                return [datetime]::MinValue

                                                                }

                                                }

                                },

                                @{ Name = 'Category'; Expression = { Get-SortieCategory -SortieType $_.SortieType } }

 

                # Sort: Day, then Flight before Sim, then Time

                $sorted = $enriched | Sort-Object Day, Category, ETD

 

                # Final columns - Day is intentionally omitted (date already lives in ETD)

                $exportData = $sorted | Select-Object ETD, Instructor, SortieType, Student, Mission, Callsign, Num, Notes

 

                $excel = $exportData | Export-Excel -Path $FilePath -WorksheetName $WorksheetName -AutoSize -PassThru

 

                try {

                                $ws = $excel.Workbook.Worksheets[$WorksheetName]

 

                                # No outer border around the whole schedule (per your request)

                                # Only the internal group separators below are drawn

                                Apply-GroupBorders -Worksheet $ws -SortedData $sorted

 

                                $ws.Row(1).Style.Font.Bold = $true

 

                                $excel.Save()

                } finally {

                                if ($excel) { $excel.Dispose() }

                }

}

 

 

 

 

 

 

# ==========================================

#      HELPER: Apply medium border between Flight/Sim groups and thick black border between days

# ==========================================

function Apply-GroupBorders {

                [CmdletBinding()]

                param(

                                [Parameter(Mandatory = $true)]

                                $Worksheet,

                                [Parameter(Mandatory = $true)]

                                [array]$SortedData

                )

                if ($null -eq $Worksheet -or $null -eq $Worksheet.Dimension -or $SortedData.Count -eq 0) {

                                return

                }

 

                $lastRow = $Worksheet.Dimension.End.Row

                $lastCol = $Worksheet.Dimension.End.Column

 

                $prevDay = $null

                $prevCategory = $null

 

                for ($row = 2; $row -le $lastRow; $row++) {

                                $dataIndex = $row - 2

                                if ($dataIndex -ge $SortedData.Count) { break }

 

                                $current = $SortedData[$dataIndex]

                                $currDay = $current.Day

                                $currCategory = $current.Category

 

                                $applyMedium = $false

                                $applyThickBlack = $false

 

                                if ($prevDay -ne $null) {

                                                if ($currDay -ne $prevDay) {

                                                                $applyThickBlack = $true

                                                } elseif ($currCategory -ne $prevCategory -and $currCategory -eq 'Sim') {

                                                                # Medium separator only when Flight section ends and Sim section begins (same day)

                                                                $applyMedium = $true

                                                }

                                }

 

                                if ($applyThickBlack -or $applyMedium) {

                                                $borderStyle = if ($applyThickBlack) {

                                                                [OfficeOpenXml.Style.ExcelBorderStyle]::Thick

                                                } else {

                                                                [OfficeOpenXml.Style.ExcelBorderStyle]::Medium

                                                }

                                                $borderColor = [System.Drawing.Color]::Black

 

                                                for ($col = 1; $col -le $lastCol; $col++) {

                                                                $cell = $Worksheet.Cells[$row, $col]

                                                                $cell.Style.Border.Top.Style = $borderStyle

                                                                $cell.Style.Border.Top.Color.SetColor($borderColor)

                                                }

                                }

 

                                $prevDay = $currDay

                                $prevCategory = $currCategory

                }

}

 

 

# ==========================================

# HELPER: Move MasterSchedule worksheet to first position

# ==========================================

function Set-MasterScheduleFirst {

                [CmdletBinding()]

                param(

                                [Parameter(Mandatory = $true)]

                                [string]$FilePath

                )

 

                if (-not (Test-Path -Path $FilePath)) {

                                Write-Host "Set-MasterScheduleFirst: file not found - $FilePath" -ForegroundColor Magenta

                                return

                }

 

                $package = $null

                try {

                                $package = Open-ExcelPackage -Path $FilePath

                                $ws = $package.Workbook.Worksheets["MasterSchedule"]

 

                                if ($null -eq $ws) {

                                                Write-Host "Set-MasterScheduleFirst: MasterSchedule worksheet not found in $FilePath" -ForegroundColor Magenta

                                                return

                                }

 

                                # EPPlus uses 1-based index; move sheet to position 1 when it is not already first

                                if ($ws.Index -ne 1) {

                                                $package.Workbook.Worksheets.MoveToStart("MasterSchedule")

                                }

 

                                Close-ExcelPackage -ExcelPackage $package -Save

                                Write-Host "MasterSchedule is now the first worksheet." -ForegroundColor Green

                }

                catch {

                                Write-Host "Set-MasterScheduleFirst failed: $_" -ForegroundColor Magenta

                                if ($null -ne $package) {

                                                Close-ExcelPackage -ExcelPackage $package

                                }

                }

}

 

 

 

 

Export-ModuleMember -Function Get-MatchedInstructor, IngestSchedule, IngestStudentAbsences, OandB_Screener, Schedule-ReturnLeg, Add-SortieNote, Process-HardcodedSortie, Get-SortieCategory, Apply-ThickBorders, Apply-GroupBorders, Export-FormattedMasterSchedule, Set-MasterScheduleFirst