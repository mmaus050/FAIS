# DATAUPLOAD.PSM1 DIRECTORY:
# Contains all functions that move data from the Program to an output destination
# ConvertTo-ExportRows
# Read-MasterScheduleRows / Merge-MasterScheduleRows (one day at a time: other days are kept)
# Export-FormattedMasterSchedule / Apply-GroupBorders / Set-MasterScheduleFirst
# Export-StudentSyllabusTabs / ConvertTo-SyllabusTabRow / Update-SyllabusTabCells / Format-TabEtd
# Update-MasterDataStudents / Get-UnsatFlagUpdates
# Export-AtisWorkbook (ATIS.xlsx, one tab per day)
# Export-DailyScheduleSheets (Teams) / Get-DaySheetName
# Invoke-WithFileBackup / Export-FaisOutput
# (Test-FileWritable / Wait-FileWritable and saving / closing workbooks in Excel live in ExcelFiles)


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
#      ONE DAY AT A TIME
#      MasterSchedule holds every exported day from today on. Exporting replaces only
#      the days that were filled, so a day can be updated without touching the others.
# ==========================================

# The rows already in OutputSchedule's MasterSchedule, as text (empty when there is none)
function Read-MasterScheduleRows {
                param ([string]$Path, [string]$WorksheetName = 'MasterSchedule')

                $book = Open-FaisWorkbookForRead -Path $Path
                try {
                                if ($null -eq $book.Package) { return @() }
                                $ws = $book.Package.Workbook.Worksheets[$WorksheetName]
                                if ($null -eq $ws -or $null -eq $ws.Dimension -or $ws.Dimension.End.Row -lt 2) { return @() }
                                $raw = @(Import-Excel -ExcelPackage $book.Package -WorksheetName $WorksheetName -ErrorAction Stop)
                                return @(ConvertTo-ExportRows -Schedule @($raw | Where-Object { $null -ne (Parse-ETD $_.ETD) }))
                } catch {
                                Write-Warning "Could not read the existing $WorksheetName in $Path : $($_.Exception.Message)"
                                return @()
                } finally {
                                Close-FaisWorkbookForRead $book
                }
}

# Existing MasterSchedule rows with -Days replaced by -New. Rows from days before -KeepFromDay
# ('yyyy-MM-dd') are dropped so the sheet does not grow for ever.
function Merge-MasterScheduleRows {
                param (
                                [array]$Existing = @(),
                                [array]$New = @(),
                                [string[]]$Days = @(),
                                [string]$KeepFromDay = ''
                )

                $kept = @(foreach ($r in $Existing) {
                                $d = Parse-ETD $r.ETD
                                if ($null -eq $d) { continue }
                                $day = $d.ToString('yyyy-MM-dd')
                                if ($Days -contains $day) { continue }
                                if ($KeepFromDay -ne '' -and $day -lt $KeepFromDay) { continue }
                                $r
                })
                return @(@($kept) + @($New) | Where-Object { $null -ne $_ })
}

# 'yyyy-MM-dd' of an export row ('' when its ETD is not a date)
function Get-ExportRowDay {
                param ($Row)

                $d = Parse-ETD $Row.ETD
                if ($null -eq $d) { return '' }
                return $d.ToString('yyyy-MM-dd')
}

function Format-ExportDay {
                param ([string]$Day)
                return ([datetime]::ParseExact($Day, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)).ToString('ddd dd MMM')
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

                # -ClearSheet so a shorter schedule never leaves old rows behind (the only sheet FAIS clears)
                $excel = $exportData | Export-Excel -Path $FilePath -WorksheetName $WorksheetName -AutoSize -ClearSheet -PassThru -ErrorAction Stop

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
# One tab per student in OutputSchedule.xlsx, in the columns IngestData reads back,
# so tomorrow's run starts from today's schedule.
# A student with no tab gets a full one. An existing tab is never cleared: only the
# Instructor / ETD / IsComplete cells of events that changed are written, and every
# other cell, row, column and format on it stays as it was.
# ==========================================

# FAIS's columns on a student tab, in the order a new tab gets them
$script:SyllabusTabColumns = @('SyllabusID', 'Mission', 'Prereqs', 'Instructor', 'ETD', 'Type', 'IsCheck', 'IsComplete', 'OptingEvents', 'Level')

function ConvertTo-SyllabusTabRow {
                param ($Mission)

                return [PSCustomObject]@{
                                SyllabusID   = $Mission.SyllabusID
                                Mission      = $Mission.Mission
                                Prereqs      = $Mission.Prereqs
                                Instructor   = Get-NameText $Mission.Instructor
                                ETD          = [string]$Mission.ETD
                                Type         = $Mission.Type
                                IsCheck      = [bool](ConvertTo-FaisBool $Mission.IsCheck)
                                IsComplete   = [bool](Test-MissionComplete $Mission)
                                OptingEvents = $Mission.NumOpting
                                Level        = $Mission.Level
                }
}

function Export-StudentSyllabusTabs {
                param (
                                [array]$Students,
                                [Parameter(Mandatory)] [string]$FilePath
                )

                $package = $null
                try {
                                $package = Open-ExcelPackage -Path $FilePath -Create
                                $created = 0
                                $cells = 0
                                foreach ($s in $Students) {
                                                $sheet = ([string]$s.Name).Trim()
                                                if ($sheet -eq '') { continue }
                                                if ($sheet.Length -gt 31 -or $sheet -match '[\[\]\*\?/\\:]') {
                                                                Write-Warning "Student name '$sheet' can't be an Excel sheet name (31 chars, no []*?/\:). Their tab was not written."
                                                                continue
                                                }
                                                $rows = @(foreach ($m in @($s.Syllabus)) { if ($null -ne $m) { ConvertTo-SyllabusTabRow $m } })
                                                if ($rows.Count -eq 0) { continue }

                                                $ws = $package.Workbook.Worksheets[$sheet]
                                                if ($null -eq $ws -or $null -eq $ws.Dimension) {
                                                                $package = $rows | Export-Excel -ExcelPackage $package -WorksheetName $sheet -AutoSize -PassThru -ErrorAction Stop
                                                                $created++
                                                }
                                                else {
                                                                $cells += Update-SyllabusTabCells -Worksheet $ws -Rows $rows
                                                }
                                }
                                Close-ExcelPackage -ExcelPackage $package
                                $package = $null
                                Write-Host "Student syllabus tabs updated in $FilePath ($cells cell(s) changed, $created new tab(s))." -ForegroundColor Green
                }
                catch {
                                if ($null -ne $package) { Close-ExcelPackage -ExcelPackage $package -NoSave }
                                throw
                }
}

# Writes the Instructor / ETD / IsComplete cells that differ from -Rows, matched to the
# tab's rows by Mission (a name listed twice matches in order). A mission missing from
# the tab is added at the bottom. Returns how many cells were written.
function Update-SyllabusTabCells {
                param (
                                [Parameter(Mandatory)] $Worksheet,
                                [array]$Rows
                )

                $lastRow = $Worksheet.Dimension.End.Row
                $lastCol = $Worksheet.Dimension.End.Column
                $col = @{}
                for ($c = 1; $c -le $lastCol; $c++) {
                                $heading = ([string]$Worksheet.Cells[1, $c].Value).Trim()
                                if ($heading -ne '' -and -not $col.ContainsKey($heading)) { $col[$heading] = $c }
                }
                if (-not $col.ContainsKey('Mission')) {
                                throw "The '$($Worksheet.Name)' tab has no Mission heading in row 1, so its events can't be matched."
                }
                foreach ($heading in @('Instructor', 'ETD', 'IsComplete')) {
                                if (-not $col.ContainsKey($heading)) {
                                                $lastCol++
                                                $Worksheet.Cells[1, $lastCol].Value = $heading
                                                $col[$heading] = $lastCol
                                }
                }

                # Mission -> the tab's row numbers for it, top to bottom
                $rowsOf = @{}
                for ($r = 2; $r -le $lastRow; $r++) {
                                $name = ([string]$Worksheet.Cells[$r, $col['Mission']].Value).Trim()
                                if ($name -eq '') { continue }
                                if (-not $rowsOf.ContainsKey($name)) { $rowsOf[$name] = New-Object System.Collections.Generic.List[int] }
                                $rowsOf[$name].Add($r)
                }
                $used = @{}

                $count = 0
                foreach ($row in $Rows) {
                                $name = ([string]$row.Mission).Trim()
                                if ($name -eq '') { continue }
                                $n = if ($used.ContainsKey($name)) { $used[$name] } else { 0 }
                                $used[$name] = $n + 1

                                if (-not $rowsOf.ContainsKey($name) -or $n -ge $rowsOf[$name].Count) {
                                                $lastRow++
                                                foreach ($heading in $script:SyllabusTabColumns) {
                                                                if ($col.ContainsKey($heading) -and $null -ne $row.$heading -and [string]$row.$heading -ne '') {
                                                                                $Worksheet.Cells[$lastRow, $col[$heading]].Value = $row.$heading
                                                                                $count++
                                                                }
                                                }
                                                continue
                                }

                                $r = $rowsOf[$name][$n]
                                $cell = $Worksheet.Cells[$r, $col['IsComplete']]
                                if ((ConvertTo-FaisBool $cell.Value) -ne [bool]$row.IsComplete) {
                                                $cell.Value = [bool]$row.IsComplete
                                                $count++
                                }
                                $etd = Format-TabEtd $row.ETD
                                $cell = $Worksheet.Cells[$r, $col['ETD']]
                                if ((Format-TabEtd $cell.Value) -ne $etd) {
                                                $cell.Value = if ($etd -eq '') { $null } else { $etd }
                                                $count++
                                }
                                $ip = ([string]$row.Instructor).Trim()
                                $cell = $Worksheet.Cells[$r, $col['Instructor']]
                                if (([string]$cell.Value).Trim() -cne $ip) {
                                                $cell.Value = if ($ip -eq '') { $null } else { $ip }
                                                $count++
                                }
                }
                return $count
}

# An ETD cell as the text FAIS writes ('' when blank or not a date), for comparing
function Format-TabEtd {
                param ($Value)

                $dt = Parse-ETD $Value
                if ($null -eq $dt) { return '' }
                return $dt.ToString('yyyy-MM-dd HH:mm:ss')
}


# ==========================================
# Writes student values into the MasterData Students sheet, matched by SPName (the column
# headed SPName, else B): SimDebt into the column headed SimDebt (else P), UnsatSim /
# UnsatFlight into the columns with those headings.
# Only the cells passed in change. Throws when MasterData can't be written.
#   -SimDebtDeltaByName: name -> amount added to the SimDebt already there (Update Syllabi)
#   -FlagsByName: name -> @{ UnsatSim = $true/$false; UnsatFlight = $true/$false; OnHold = $true/$false } (any of the keys)
# ==========================================
function Update-MasterDataStudents {
                param (
                                [Parameter(Mandatory)] [string]$MasterDataPath,
                                [hashtable]$SimDebtByName = @{},
                                [hashtable]$SimDebtDeltaByName = @{},
                                [hashtable]$FlagsByName = @{}
                )

                if (-not (Test-Path -LiteralPath $MasterDataPath)) { throw "MasterData.xlsx not found at: $MasterDataPath" }
                if ($SimDebtByName.Count -eq 0 -and $SimDebtDeltaByName.Count -eq 0 -and $FlagsByName.Count -eq 0) { return }

                $package = $null
                try {
                                $package = Open-ExcelPackage -Path $MasterDataPath
                                $ws = $package.Workbook.Worksheets["Students"]
                                if ($null -eq $ws) { throw "Worksheet 'Students' not found in MasterData.xlsx" }

                                # Columns by their headings (the ones IngestData reads); B = SPName and P = SimDebt
                                # when those headings are missing
                                $nameCol = 2
                                $debtCol = 16
                                $flagCol = @{}
                                $lastCol = if ($null -ne $ws.Dimension) { $ws.Dimension.End.Column } else { 0 }
                                $seenName = $false; $seenDebt = $false
                                for ($c = 1; $c -le $lastCol; $c++) {
                                                $heading = ([string]$ws.Cells[1, $c].Value).Trim()
                                                if ($heading -in @('UnsatSim', 'UnsatFlight', 'OnHold') -and -not $flagCol.ContainsKey($heading)) { $flagCol[$heading] = $c }
                                                elseif ($heading -ieq 'SPName' -and -not $seenName) { $nameCol = $c; $seenName = $true }
                                                elseif ($heading -ieq 'SimDebt' -and -not $seenDebt) { $debtCol = $c; $seenDebt = $true }
                                }
                                foreach ($flags in $FlagsByName.Values) {
                                                foreach ($key in $flags.Keys) {
                                                                if (-not $flagCol.ContainsKey($key)) { throw "The MasterData Students sheet has no '$key' heading in row 1." }
                                                }
                                }

                                $row = 2
                                $found = @{}
                                while ($true) {
                                                $nameVal = $ws.Cells[$row, $nameCol].Value
                                                if ([string]::IsNullOrWhiteSpace([string]$nameVal)) { break }
                                                $studentName = $nameVal.ToString().Trim()
                                                if ($SimDebtByName.ContainsKey($studentName)) {
                                                                $ws.Cells[$row, $debtCol].Value = [double]$SimDebtByName[$studentName]
                                                                $found[$studentName] = $true
                                                }
                                                elseif ($SimDebtDeltaByName.ContainsKey($studentName)) {
                                                                $current = 0.0
                                                                $old = $ws.Cells[$row, $debtCol].Value
                                                                if ($null -ne $old -and ([string]$old).Trim() -ne '') {
                                                                                try { $current = [double]$old } catch { [void][double]::TryParse([string]$old, [ref]$current) }
                                                                }
                                                                $ws.Cells[$row, $debtCol].Value = [math]::Round($current + [double]$SimDebtDeltaByName[$studentName], 2)
                                                                $found[$studentName] = $true
                                                }
                                                if ($FlagsByName.ContainsKey($studentName)) {
                                                                $flags = $FlagsByName[$studentName]
                                                                foreach ($key in $flags.Keys) { $ws.Cells[$row, $flagCol[$key]].Value = [bool]$flags[$key] }
                                                                $found[$studentName] = $true
                                                }
                                                $row++
                                }

                                Close-ExcelPackage -ExcelPackage $package
                                $package = $null

                                $missing = @(@($SimDebtByName.Keys) + @($SimDebtDeltaByName.Keys) + @($FlagsByName.Keys) | Sort-Object -Unique | Where-Object { -not $found.ContainsKey($_) })
                                if ($missing.Count -gt 0) {
                                                Write-Warning "Not found in the SPName column of the MasterData Students sheet (above the first blank name), so nothing was written for: $($missing -join '; ')"
                                }
                                $debtNames = @(@($SimDebtByName.Keys) + @($SimDebtDeltaByName.Keys) | Sort-Object -Unique | Where-Object { $found.ContainsKey($_) })
                                Write-Host ("MasterData updated: SimDebt for {0} student(s), unsat / HOLD flags for {1}." -f $debtNames.Count, @($FlagsByName.Keys | Where-Object { $found.ContainsKey($_) }).Count) -ForegroundColor Green
                }
                catch {
                                if ($null -ne $package) { Close-ExcelPackage -ExcelPackage $package -NoSave }
                                throw
                }
}

# The UnsatSim / UnsatFlight / OnHold values that no longer match what was read from MasterData:
# $false where this run's schedule cleaned up the unsat, $true where a hook (or a re-planned
# clean-up) set one that MasterData doesn't have yet, and a HOLD change the pre-run menu
# could not save. Same shape as -FlagsByName above.
function Get-UnsatFlagUpdates {
                param ([array]$Students)

                $out = @{}
                foreach ($s in @($Students)) {
                                if ($null -eq $s) { continue }
                                $name = ([string]$s.Name).Trim()
                                if ($name -eq '') { continue }
                                $flags = @{}
                                foreach ($key in @('UnsatSim', 'UnsatFlight', 'OnHold')) {
                                                $saved = $s.PSObject.Properties["${key}InMasterData"]
                                                if ($null -eq $saved) { continue }
                                                $now = ConvertTo-FaisBool $s.$key
                                                if ($now -ne (ConvertTo-FaisBool $saved.Value)) { $flags[$key] = $now }
                                }
                                if ($flags.Count -gt 0) { $out[$name] = $flags }
                }
                return $out
}


# ==========================================
# ATIS.xlsx: one tab per day, flights in time order, then sims in time order below a
# thick border. Each line goes on its day's tab in MasterSchedule columns. A day's tab
# is rebuilt every time that day is exported; tabs for other days are kept.
# Throws when ATIS can't be written.
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
                                                $package = $rows | Export-Excel -ExcelPackage $package -WorksheetName $sheet -ClearSheet -AutoSize -PassThru -ErrorAction Stop
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
                                # Export-FaisOutput puts every file back when this fails
                                throw "ATIS was not written: $($_.Exception.Message)"
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
                $tempDir = if ($env:TEMP) { $env:TEMP } else { [System.IO.Path]::GetTempPath() }
                $tempPath = Join-Path -Path $tempDir -ChildPath ("ATIS_DailySchedule_" + [guid]::NewGuid().ToString() + ".xlsx")
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

                                # Each exported day replaces its own sheet (a re-exported day never shows its old lines too)
                                foreach ($dayKey in ($byDay.Keys | Sort-Object)) {
                                                $dayDate = [datetime]::ParseExact($dayKey, "yyyy-MM-dd", $null)
                                                $sheetName = Get-DaySheetName -Date $dayDate
                                                $dayRows = @($byDay[$dayKey].ToArray() | Sort-Object { Parse-ETD $_.ETD })
                                                $dayRows | Export-Excel -Path $tempPath -WorksheetName $sheetName -AutoSize -ClearSheet
                                                Write-Host ("[DailyExport] Wrote sheet '{0}' with {1} rows" -f $sheetName, $dayRows.Count) -ForegroundColor Green
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

# Runs -Action. If it throws, every file in -Files goes back to how it was before (a file
# that didn't exist is removed), so a write that fails part-way leaves nothing half-done.
# The error is rethrown.
function Invoke-WithFileBackup {
                param (
                                [string[]]$Files,
                                [Parameter(Mandatory)] [scriptblock]$Action
                )

                $backupOf = @{}
                try {
                                foreach ($target in $Files) {
                                                if ([string]::IsNullOrWhiteSpace($target)) { continue }
                                                if (Test-Path -LiteralPath $target) {
                                                                $copy = Join-Path ([System.IO.Path]::GetTempPath()) ('FAIS_' + [guid]::NewGuid().ToString('N') + [System.IO.Path]::GetExtension($target))
                                                                Copy-Item -LiteralPath $target -Destination $copy -Force -ErrorAction Stop
                                                                $backupOf[$target] = $copy
                                                }
                                                else {
                                                                $backupOf[$target] = $null
                                                }
                                }
                                & $Action
                }
                catch {
                                foreach ($target in @($backupOf.Keys)) {
                                                try {
                                                                if ($null -ne $backupOf[$target]) { Copy-Item -LiteralPath $backupOf[$target] -Destination $target -Force -ErrorAction Stop }
                                                                elseif (Test-Path -LiteralPath $target) { Remove-Item -LiteralPath $target -Force -ErrorAction Stop }
                                                }
                                                catch {
                                                                Write-Warning "Could not put $target back the way it was: $($_.Exception.Message)"
                                                }
                                }
                                throw
                }
                finally {
                                foreach ($copy in @($backupOf.Values)) {
                                                if ($null -ne $copy -and (Test-Path -LiteralPath $copy)) { Remove-Item -LiteralPath $copy -Force -ErrorAction SilentlyContinue }
                                }
                }
}

function Get-OutputSetting {
                param ($Settings, [string]$Name, $Default)

                if ($null -ne $Settings -and $Settings -is [System.Collections.IDictionary] -and $Settings.Contains($Name)) { return $Settings[$Name] }
                return $Default
}

# Writes everything a run produces, or nothing: MasterSchedule, each student's syllabus
# tab, SimDebt and the unsat flags that changed to MasterData, ATIS, then the optional Teams
# copy. If any write fails, all of them are put back as they were.
# Only the days that were filled (-Days; default: every day in the fill) are written:
# MasterSchedule keeps the other days already exported (from today on, plus KeepPastDays),
# and ATIS keeps their tabs. Each file open in Excel is saved and closed first
# (SaveAndCloseExcel) and opened again afterwards (ReopenExcel).
# Returns $true when everything was written.
function Export-FaisOutput {
                param (
                                [Parameter(Mandatory)] $FillResult,
                                [Parameter(Mandatory)] [string]$OutputPath,
                                [Parameter(Mandatory)] [string]$MasterDataPath,
                                [string]$AtisPath,
                                $Settings = @{},
                                [scriptblock]$LockedPrompt,
                                [string[]]$Days = @(),
                                [datetime]$Today = (Get-Date)
                )

                $allRows = @(ConvertTo-ExportRows -Schedule $FillResult.Schedule)
                if ($Days.Count -eq 0 -and $null -ne $FillResult.PSObject.Properties['Days']) { $Days = @($FillResult.Days) }
                if ($Days.Count -eq 0) { $Days = @($allRows | ForEach-Object { Get-ExportRowDay $_ } | Where-Object { $_ } | Sort-Object -Unique) }
                $rows = @($allRows | Where-Object { $Days -contains (Get-ExportRowDay $_) })

                $debts = @{}
                foreach ($s in $FillResult.Students) {
                                if (-not [string]::IsNullOrWhiteSpace([string]$s.Name)) { $debts[([string]$s.Name).Trim()] = [double]$s.SimDebt }
                }
                $flags = Get-UnsatFlagUpdates -Students $FillResult.Students

                $writeAtis = -not [string]::IsNullOrWhiteSpace($AtisPath)
                if ($writeAtis) { $writeAtis = [bool](Get-OutputSetting $Settings 'WriteAtis' $true) }
                $useExcel = [bool](Get-OutputSetting $Settings 'SaveAndCloseExcel' $true)
                $reopen = [bool](Get-OutputSetting $Settings 'ReopenExcel' $true)
                $keepPast = [int](Get-OutputSetting $Settings 'KeepPastDays' 0)
                $outFiles = @($OutputPath, $MasterDataPath)
                if ($writeAtis) { $outFiles += $AtisPath }

                if (-not (Test-Path -LiteralPath $MasterDataPath)) {
                                Write-Warning "MasterData.xlsx was not found at $MasterDataPath, so nothing was written."
                                return $false
                }

                $closed = New-Object System.Collections.Generic.List[string]
                try {
                                if ($useExcel -and [bool](Get-OutputSetting $Settings 'CloseAllExcelWorkbooks' $false)) {
                                                Write-Host "Saving and closing every open Excel workbook before writing..." -ForegroundColor Cyan
                                                foreach ($r in @(Close-FaisExcelFiles -All)) {
                                                                if ($r.Closed -and $r.Path -and (Test-Path -LiteralPath $r.Path -ErrorAction SilentlyContinue)) { $closed.Add($r.Path) }
                                                }
                                }
                                foreach ($file in $outFiles) {
                                                if (-not (Request-FaisFileWritable -Path $file -Prompt $LockedPrompt -Closed $closed -NoExcel:(-not $useExcel))) {
                                                                Write-Warning "$file is still open, so nothing was written."
                                                                return $false
                                                }
                                }

                                # The other days already in MasterSchedule stay; past days drop off
                                $existing = @(Read-MasterScheduleRows -Path $OutputPath)
                                $keepFrom = $Today.Date.AddDays(-[math]::Max(0, $keepPast)).ToString('yyyy-MM-dd')
                                $merged = @(Merge-MasterScheduleRows -Existing $existing -New $rows -Days $Days -KeepFromDay $keepFrom)

                                try {
                                                Invoke-WithFileBackup -Files $outFiles -Action {
                                                                Export-FormattedMasterSchedule -ScheduleData $merged -FilePath $OutputPath -WorksheetName 'MasterSchedule'
                                                                Export-StudentSyllabusTabs -Students $FillResult.Students -FilePath $OutputPath
                                                                Set-MasterScheduleFirst -FilePath $OutputPath
                                                                Update-MasterDataStudents -MasterDataPath $MasterDataPath -SimDebtByName $debts -FlagsByName $flags
                                                                if ($writeAtis) { Export-AtisWorkbook -ScheduleData $rows -FilePath $AtisPath }
                                                }
                                }
                                catch {
                                                Write-Warning "Export failed, so OutputSchedule, MasterData and ATIS were left as they were: $($_.Exception.Message)"
                                                return $false
                                }

                                $wrote = @($Days | ForEach-Object { $d = $_; "{0} ({1} lines)" -f (Format-ExportDay $d), @($rows | Where-Object { (Get-ExportRowDay $_) -eq $d }).Count })
                                Write-Host ("MasterSchedule: wrote {0}." -f ($wrote -join ', ')) -ForegroundColor Green
                                $otherDays = @($merged | ForEach-Object { Get-ExportRowDay $_ } | Where-Object { $_ -and $Days -notcontains $_ } | Sort-Object -Unique)
                                if ($otherDays.Count -gt 0) { Write-Host ("MasterSchedule: kept {0} as exported before." -f (@($otherDays | ForEach-Object { Format-ExportDay $_ }) -join ', ')) -ForegroundColor Green }
                                foreach ($name in @($flags.Keys | Sort-Object)) {
                                                foreach ($key in @($flags[$name].Keys | Sort-Object)) {
                                                                $what = if ($flags[$name][$key]) { 'set' } else { 'cleared (cleaned up on this schedule)' }
                                                                Write-Host "  $key $what for $name" -ForegroundColor Cyan
                                                }
                                }

                                $teamsFolder = [string](Get-OutputSetting $Settings 'TeamsFolder' '')
                                if ([bool](Get-OutputSetting $Settings 'PublishToTeams' $false) -and -not [string]::IsNullOrWhiteSpace($teamsFolder)) {
                                                $teamsFile = [string](Get-OutputSetting $Settings 'TeamsFileName' 'TESTING ATIS NOT OFFICIAL.xlsx')
                                                if ([string]::IsNullOrWhiteSpace($teamsFile)) { $teamsFile = 'TESTING ATIS NOT OFFICIAL.xlsx' }
                                                if (-not $teamsFile.EndsWith('.xlsx', [System.StringComparison]::OrdinalIgnoreCase)) { $teamsFile = "$teamsFile.xlsx" }
                                                [void](Request-FaisFileWritable -Path (Join-Path $teamsFolder $teamsFile) -Prompt $LockedPrompt -Closed $closed -NoExcel:(-not $useExcel))
                                                Export-DailyScheduleSheets -ScheduleData $rows -TargetFolder $teamsFolder -FileName $teamsFile
                                }
                                return $true
                }
                finally {
                                if ($reopen -and $closed.Count -gt 0) { Open-FaisWorkbooks -Path $closed.ToArray() }
                }
}
