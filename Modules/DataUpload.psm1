# DATAUPLOAD.PSM1 DIRECTORY:
# Contains all functions that move data from the Program to an output destination
# ConvertTo-ExportRows
# Export-FormattedMasterSchedule / Apply-GroupBorders / Set-MasterScheduleFirst
# Export-StudentSyllabusTabs
# Update-MasterDataSimDebtColumnP
# Export-AtisWorkbook (ATIS.xlsx, one tab per day)
# Export-DailyScheduleSheets (Teams) / Get-DaySheetName / Get-SortieRowKey / Merge-DayRows
# Test-FileWritable / Export-FaisOutput


# =================================================================================
# Flat text copies of the filled schedule (the live rows keep their objects)
# =================================================================================
function ConvertTo-ExportRows {
                param (
                                [AllowEmptyCollection()]
                                [array]$Schedule
                )

                return @(foreach ($s in $Schedule) {
                                $dt = Parse-ETD $s.ETD
                                [PSCustomObject]@{
                                                ETD        = if ($null -ne $dt) { $dt.ToString('yyyy-MM-dd HH:mm:ss') } else { [string]$s.ETD }
                                                Instructor = Get-NameText $s.Instructor
                                                SortieType = [string]$s.SortieType
                                                Student    = Get-NameText $s.Student
                                                Mission    = Get-NameText $s.Mission
                                                Callsign   = $s.Callsign
                                                Num        = $s.Num
                                                Notes      = [string]$s.Notes
                                }
                })
}


# ==========================================
#      Export MasterSchedule sorted Day -> Flight/Sim -> Time with group borders
# ==========================================
function Export-FormattedMasterSchedule {
                [CmdletBinding()]
                param(
                                [Parameter(Mandatory = $true)]
                                [AllowEmptyCollection()]
                                [array]$ScheduleData,
                                [Parameter(Mandatory = $true)]
                                [string]$FilePath,
                                [string]$WorksheetName = 'MasterSchedule'
                )

                Write-Host ("[Export] Preparing {0} sorties for export..." -f $ScheduleData.Count) -ForegroundColor Cyan

                # Enrich with Day (used only for sorting + internal borders) and Category
                $enriched = $ScheduleData | Select-Object *,
                                @{
                                                Name = 'Day'
                                                Expression = {
                                                                $d = Parse-ETD $_.ETD
                                                                if ($null -ne $d) { $d.Date } else { [datetime]::MinValue }
                                                }
                                },
                                @{ Name = 'Category'; Expression = { Get-SortieCategory -SortieType $_.SortieType } }

                # Sort: Day, then Flight before Sim, then Time
                $sorted = @($enriched | Sort-Object Day, Category, ETD)

                # Final columns - Day is intentionally omitted (date already lives in ETD)
                $exportData = $sorted | Select-Object ETD, Instructor, SortieType, Student, Mission, Callsign, Num, Notes

                # -ClearSheet so a shorter schedule never leaves old rows behind
                $excel = $exportData | Export-Excel -Path $FilePath -WorksheetName $WorksheetName -AutoSize -ClearSheet -PassThru

                try {
                                $ws = $excel.Workbook.Worksheets[$WorksheetName]
                                Apply-GroupBorders -Worksheet $ws -SortedData $sorted
                                $ws.Row(1).Style.Font.Bold = $true
                                $excel.Save()
                } finally {
                                if ($excel) { $excel.Dispose() }
                }
}


# ==========================================
#      Medium border between Flight/Sim groups and thick black border between days
# ==========================================
function Apply-GroupBorders {
                [CmdletBinding()]
                param(
                                [Parameter(Mandatory = $true)]
                                $Worksheet,
                                [Parameter(Mandatory = $true)]
                                [AllowEmptyCollection()]
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
                                if ($null -ne $prevDay) {
                                                if ($currDay -ne $prevDay) {
                                                                $applyThickBlack = $true
                                                } elseif ($currCategory -ne $prevCategory -and $currCategory -eq 'Sim') {
                                                                # Medium separator only when Flight section ends and Sim section begins (same day)
                                                                $applyMedium = $true
                                                }
                                }

                                if ($applyThickBlack -or $applyMedium) {
                                                $borderStyle = if ($applyThickBlack) { [OfficeOpenXml.Style.ExcelBorderStyle]::Thick } else { [OfficeOpenXml.Style.ExcelBorderStyle]::Medium }
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
# Move MasterSchedule worksheet to first position
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
                                                Close-ExcelPackage -ExcelPackage $package -NoSave
                                                return
                                }
                                # EPPlus uses 1-based index; move sheet to position 1 when it is not already first
                                if ($ws.Index -ne 1) {
                                                $package.Workbook.Worksheets.MoveToStart("MasterSchedule")
                                }
                                Close-ExcelPackage -ExcelPackage $package
                                Write-Host "MasterSchedule is now the first worksheet." -ForegroundColor Green
                }
                catch {
                                Write-Host "Set-MasterScheduleFirst failed: $_" -ForegroundColor Magenta
                                if ($null -ne $package) { Close-ExcelPackage -ExcelPackage $package -NoSave }
                }
}


# ==========================================
# One tab per student in OutputSchedule.xlsx, in the exact columns IngestData reads
# back, so tomorrow's run starts from today's schedule
# ==========================================
function Export-StudentSyllabusTabs {
                param (
                                [array]$Students,
                                [Parameter(Mandatory)] [string]$FilePath
                )

                $package = $null
                try {
                                $package = Open-ExcelPackage -Path $FilePath -Create
                                foreach ($s in $Students) {
                                                $sheet = ([string]$s.Name).Trim()
                                                if ($sheet -eq '') { continue }
                                                if ($sheet.Length -gt 31 -or $sheet -match '[\[\]\*\?/\\:]') {
                                                                Write-Warning "Student name '$sheet' can't be an Excel sheet name (31 chars, no []*?/\:). Their tab was not written."
                                                                continue
                                                }
                                                $rows = @(foreach ($m in @($s.Syllabus)) {
                                                                [PSCustomObject]@{
                                                                                SyllabusID   = $m.SyllabusID
                                                                                Mission      = $m.Mission
                                                                                Prereqs      = $m.Prereqs
                                                                                Instructor   = Get-NameText $m.Instructor
                                                                                ETD          = [string]$m.ETD
                                                                                Type         = $m.Type
                                                                                IsCheck      = [bool](ConvertTo-FaisBool $m.IsCheck)
                                                                                IsComplete   = [bool](Test-MissionComplete $m)
                                                                                OptingEvents = $m.NumOpting
                                                                                Level        = $m.Level
                                                                }
                                                })
                                                if ($rows.Count -eq 0) { continue }
                                                $package = $rows | Export-Excel -ExcelPackage $package -WorksheetName $sheet -ClearSheet -AutoSize -PassThru
                                }
                                Close-ExcelPackage -ExcelPackage $package
                                $package = $null
                                Write-Host "Student syllabus tabs written to $FilePath" -ForegroundColor Green
                }
                catch {
                                if ($null -ne $package) { Close-ExcelPackage -ExcelPackage $package -NoSave }
                                throw
                }
}


# ==========================================
# Writes SimDebt into Column P of the MasterData Students sheet.
# Matches by SPName in column B so only the debt cell is changed.
# ==========================================
function Update-MasterDataSimDebtColumnP {
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
                                                Close-ExcelPackage -ExcelPackage $package -NoSave
                                                return
                                }

                                # Column B = SPName, Column P = SimDebt
                                $nameCol = 2
                                $debtCol = 16
                                $row = 2
                                $updatedCount = 0
                                while ($true) {
                                                $nameVal = $ws.Cells[$row, $nameCol].Value
                                                if ([string]::IsNullOrWhiteSpace([string]$nameVal)) { break }
                                                $studentName = $nameVal.ToString().Trim()
                                                if ($SimDebtByName.ContainsKey($studentName)) {
                                                                $ws.Cells[$row, $debtCol].Value = [double]$SimDebtByName[$studentName]
                                                                $updatedCount++
                                                }
                                                $row++
                                }

                                Close-ExcelPackage -ExcelPackage $package
                                Write-Host "Successfully updated SimDebt (Column P) for $updatedCount students." -ForegroundColor Green
                }
                catch {
                                Write-Error "Failed to update SimDebt in MasterData: $_"
                                if ($package) { Close-ExcelPackage -ExcelPackage $package -NoSave }
                }
}


# ==========================================
# ATIS.xlsx: one tab per day, flights in time order, then sims in time order below a
# thick border. Each line goes on its day's tab in MasterSchedule columns. A day's tab
# is rebuilt every time that day is exported; tabs for other days are kept.
# ==========================================
function Export-AtisWorkbook {
                [CmdletBinding()]
                param(
                                [Parameter(Mandatory = $true)]
                                [AllowEmptyCollection()]
                                [array]$ScheduleData,
                                [Parameter(Mandatory = $true)]
                                [string]$FilePath
                )

                # Group by day; Index keeps the MasterSchedule order for lines at the same time
                $byDay = @{}
                $index = 0
                $noEtd = 0
                foreach ($s in $ScheduleData) {
                                $dt = Parse-ETD $s.ETD
                                if ($null -eq $dt) { $noEtd++; continue }
                                $key = $dt.ToString('yyyy-MM-dd')
                                if (-not $byDay.ContainsKey($key)) { $byDay[$key] = New-Object System.Collections.Generic.List[object] }
                                $byDay[$key].Add([PSCustomObject]@{ Row = $s; Time = $dt; IsSim = [bool](Test-IsSimSortieType ([string]$s.SortieType)); Index = $index })
                                $index++
                }
                if ($noEtd -gt 0) { Write-Warning "$noEtd line(s) have no ETD and were left off the ATIS." }
                if ($byDay.Count -eq 0) { return }

                $package = $null
                try {
                                $package = Open-ExcelPackage -Path $FilePath -Create
                                $written = New-Object System.Collections.Generic.List[string]
                                foreach ($key in ($byDay.Keys | Sort-Object)) {
                                                $sheet = Get-DaySheetName -Date ([datetime]::ParseExact($key, 'yyyy-MM-dd', $null))
                                                # Flights first ($false sorts before $true), each block by ETD
                                                $lines = @($byDay[$key] | Sort-Object IsSim, Time, Index)
                                                $rows = @($lines | ForEach-Object { $_.Row | Select-Object ETD, Instructor, SortieType, Student, Mission, Callsign, Num, Notes })
                                                # -ClearSheet so a re-export of the day never leaves old lines or borders behind
                                                $package = $rows | Export-Excel -ExcelPackage $package -WorksheetName $sheet -ClearSheet -AutoSize -PassThru
                                                $ws = $package.Workbook.Worksheets[$sheet]
                                                $ws.Row(1).Style.Font.Bold = $true

                                                $firstSim = -1
                                                for ($i = 0; $i -lt $lines.Count; $i++) {
                                                                if ($lines[$i].IsSim) { $firstSim = $i; break }
                                                }
                                                # Thick border between the flights and the sims (row 1 is the header)
                                                if ($firstSim -gt 0) {
                                                                $border = $ws.Cells[($firstSim + 2), 1, ($firstSim + 2), $ws.Dimension.End.Column].Style.Border.Top
                                                                $border.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thick
                                                                $border.Color.SetColor([System.Drawing.Color]::Black)
                                                }
                                                $written.Add($sheet)
                                }

                                # Day tabs in date order (names come from Get-DaySheetName); any other tabs stay in front
                                $dayTabs = foreach ($w in $package.Workbook.Worksheets) {
                                                $date = [datetime]::MinValue
                                                if ([datetime]::TryParseExact($w.Name, 'dddd. MMMM d, yyyy', [cultureinfo]::CurrentCulture, [System.Globalization.DateTimeStyles]::None, [ref]$date)) {
                                                                [PSCustomObject]@{ Name = $w.Name; Date = $date }
                                                }
                                }
                                foreach ($t in @($dayTabs | Sort-Object Date)) { $package.Workbook.Worksheets.MoveToEnd($t.Name) }

                                # Open on the first day just exported. EPPlus 4.5's TabSelected = $true also
                                # unselects the sheet itself; TabSelectedMulti selects it without clearing.
                                foreach ($w in $package.Workbook.Worksheets) { $w.View.TabSelected = $false }
                                $view = $package.Workbook.Worksheets[$written[0]].View
                                if ($null -ne $view.PSObject.Properties['TabSelectedMulti']) { $view.TabSelectedMulti = $true } else { $view.TabSelected = $true }

                                Close-ExcelPackage -ExcelPackage $package
                                $package = $null
                                Write-Host ("ATIS written to {0}: {1}" -f $FilePath, ($written -join ', ')) -ForegroundColor Green
                }
                catch {
                                if ($null -ne $package) { Close-ExcelPackage -ExcelPackage $package -NoSave }
                                Write-Warning "ATIS was not written: $($_.Exception.Message)"
                }
}


# ==========================================
# TEAMS / ONEDRIVE DAILY SHEETS (optional, off by default)
# ==========================================
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
                # Composite key from the identifying fields so duplicate rows can be detected across runs
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
                # Combines existing sheet rows with new rows, removes duplicates by key, sorts by ETD
                $seen = @{}
                $unique = New-Object System.Collections.Generic.List[object]
                foreach ($row in (@($ExistingRows) + @($NewRows))) {
                                if ($null -eq $row) { continue }
                                $key = Get-SortieRowKey -Row $row
                                if (-not $seen.ContainsKey($key)) {
                                                $seen[$key] = $true
                                                $unique.Add($row)
                                }
                }
                return @($unique | Sort-Object { $d = Parse-ETD $_.ETD; if ($null -eq $d) { [datetime]::MinValue } else { $d } })
}

function Export-DailyScheduleSheets {
                [CmdletBinding()]
                param(
                                [Parameter(Mandatory = $true)]
                                [array]$ScheduleData,
                                [Parameter(Mandatory = $true)]
                                [string]$TargetFolder,
                                [string]$FileName = "TESTING ATIS NOT OFFICIAL.xlsx"
                )
                if (-not (Test-Path -Path $TargetFolder -PathType Container)) {
                                New-Item -Path $TargetFolder -ItemType Directory -Force | Out-Null
                }
                if (-not $FileName.EndsWith(".xlsx", [System.StringComparison]::OrdinalIgnoreCase)) {
                                $FileName = "$FileName.xlsx"
                }

                $fullPath = Join-Path -Path $TargetFolder -ChildPath $FileName
                Write-Host ("[DailyExport] Writing day sheets to {0}" -f $fullPath) -ForegroundColor Cyan

                # Temporary local workbook so OneDrive cannot interrupt the write
                $tempPath = Join-Path -Path $env:TEMP -ChildPath ("ATIS_DailySchedule_" + [guid]::NewGuid().ToString() + ".xlsx")
                try {
                                if (Test-Path -Path $fullPath -PathType Leaf) {
                                                Copy-Item -Path $fullPath -Destination $tempPath -Force -ErrorAction Stop
                                }

                                $byDay = @{}
                                foreach ($sortie in $ScheduleData) {
                                                $day = Parse-ETD $sortie.ETD
                                                if ($null -eq $day) { continue }
                                                $key = $day.ToString("yyyy-MM-dd")
                                                if (-not $byDay.ContainsKey($key)) { $byDay[$key] = New-Object System.Collections.Generic.List[object] }
                                                $byDay[$key].Add([PSCustomObject]@{
                                                                ETD        = $sortie.ETD
                                                                Instructor = $sortie.Instructor
                                                                SortieType = $sortie.SortieType
                                                                Student    = $sortie.Student
                                                                Mission    = $sortie.Mission
                                                                Callsign   = $sortie.Callsign
                                                                Num        = $sortie.Num
                                                                Notes      = $sortie.Notes
                                                })
                                }

                                foreach ($dayKey in ($byDay.Keys | Sort-Object)) {
                                                $dayDate = [datetime]::ParseExact($dayKey, "yyyy-MM-dd", $null)
                                                $sheetName = Get-DaySheetName -Date $dayDate
                                                $existingRows = @()
                                                if (Test-Path -Path $tempPath -PathType Leaf) {
                                                                try { $existingRows = @(Import-Excel -Path $tempPath -WorksheetName $sheetName -ErrorAction Stop) } catch { $existingRows = @() }
                                                }
                                                $merged = Merge-DayRows -ExistingRows $existingRows -NewRows $byDay[$dayKey].ToArray()
                                                $merged | Export-Excel -Path $tempPath -WorksheetName $sheetName -AutoSize -ClearSheet
                                                Write-Host ("[DailyExport] Updated sheet '{0}' with {1} rows" -f $sheetName, $merged.Count) -ForegroundColor Green
                                }

                                Copy-Item -Path $tempPath -Destination $fullPath -Force -ErrorAction Stop
                                Write-Host ("[DailyExport] Successfully updated {0}" -f $fullPath) -ForegroundColor Green
                }
                catch {
                                Write-Host ("[DailyExport] Failed to write schedule: {0}" -f $_.Exception.Message) -ForegroundColor Magenta
                }
                finally {
                                if (Test-Path -Path $tempPath) { Remove-Item -Path $tempPath -Force -ErrorAction SilentlyContinue }
                }
}


# ==========================================
# FINAL OUTPUT
# ==========================================

# True when the file does not exist yet or can be opened for writing (not open in Excel)
function Test-FileWritable {
                param ([string]$Path)

                if (-not (Test-Path -LiteralPath $Path)) { return $true }
                try {
                                $fs = [System.IO.File]::Open($Path, 'Open', 'ReadWrite', 'None')
                                $fs.Close()
                                return $true
                } catch {
                                return $false
                }
}

# Waits for a file to be closed in Excel. -Prompt is a scriptblock that gets the path
# and returns $true to retry. Returns $false if the user gave up.
function Wait-FileWritable {
                param (
                                [string]$Path,
                                [scriptblock]$Prompt
                )

                while (-not (Test-FileWritable $Path)) {
                                $retry = $false
                                if ($null -ne $Prompt) { $retry = [bool](& $Prompt $Path) }
                                else {
                                                $answer = Read-Host "$Path is open (probably in Excel). Close it and press Enter to retry, or type N to skip"
                                                $retry = ($answer -notmatch '^(?i)n')
                                }
                                if (-not $retry) { return $false }
                }
                return $true
}

# Writes everything a run produces:
# MasterSchedule, each student's syllabus tab, SimDebt to MasterData, ATIS, optional Teams copy.
function Export-FaisOutput {
                param (
                                [Parameter(Mandatory)] $FillResult,
                                [Parameter(Mandatory)] [string]$OutputPath,
                                [string]$MasterDataPath,
                                [string]$AtisPath,
                                $Settings = @{},
                                [scriptblock]$LockedPrompt
                )

                $rows = ConvertTo-ExportRows -Schedule $FillResult.Schedule
                $writeTabs = $true
                if ($null -ne $Settings -and $Settings.ContainsKey('WriteStudentTabs')) { $writeTabs = [bool]$Settings.WriteStudentTabs }

                if (-not (Wait-FileWritable -Path $OutputPath -Prompt $LockedPrompt)) {
                                Write-Warning "OutputSchedule was not written (file still open)."
                                return $false
                }
                Export-FormattedMasterSchedule -ScheduleData $rows -FilePath $OutputPath -WorksheetName 'MasterSchedule'
                if ($writeTabs) { Export-StudentSyllabusTabs -Students $FillResult.Students -FilePath $OutputPath }
                Set-MasterScheduleFirst -FilePath $OutputPath

                if ($writeTabs -and -not [string]::IsNullOrWhiteSpace($MasterDataPath)) {
                                if (Wait-FileWritable -Path $MasterDataPath -Prompt $LockedPrompt) {
                                                $debts = @{}
                                                foreach ($s in $FillResult.Students) {
                                                                if (-not [string]::IsNullOrWhiteSpace([string]$s.Name)) { $debts[[string]$s.Name] = [double]$s.SimDebt }
                                                }
                                                Update-MasterDataSimDebtColumnP -SimDebtByName $debts -MasterDataPath $MasterDataPath
                                }
                                else {
                                                Write-Warning "SimDebt was not saved to MasterData (file still open)."
                                }
                }

                $writeAtis = -not [string]::IsNullOrWhiteSpace($AtisPath)
                if ($writeAtis -and $null -ne $Settings -and $Settings.ContainsKey('WriteAtis')) { $writeAtis = [bool]$Settings.WriteAtis }
                if ($writeAtis) {
                                if (Wait-FileWritable -Path $AtisPath -Prompt $LockedPrompt) {
                                                Export-AtisWorkbook -ScheduleData $rows -FilePath $AtisPath
                                }
                                else {
                                                Write-Warning "ATIS was not written (file still open)."
                                }
                }

                if ($null -ne $Settings -and $Settings.ContainsKey('PublishToTeams') -and $Settings.PublishToTeams -and -not [string]::IsNullOrWhiteSpace([string]$Settings.TeamsFolder)) {
                                $teamsFile = if ($Settings.ContainsKey('TeamsFileName') -and $Settings.TeamsFileName) { $Settings.TeamsFileName } else { 'TESTING ATIS NOT OFFICIAL.xlsx' }
                                Export-DailyScheduleSheets -ScheduleData $rows -TargetFolder $Settings.TeamsFolder -FileName $teamsFile
                }
                return $true
}
