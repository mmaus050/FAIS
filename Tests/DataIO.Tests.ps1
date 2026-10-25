# Pester 5 tests for the Excel in/out (Modules\DataDownload.psm1, Modules\DataUpload.psm1):
# IngestData, the student tabs (updated in place, never cleared), the MasterData Students
# writes and the all-or-nothing Export. Workbooks go to TestDrive. Needs ImportExcel.
# Dates (2026): Fri 20 Feb .. Fri 27 Feb history, Tue 03 Mar = the day being scheduled.

BeforeDiscovery {
                $script:HasExcel = $null -ne (Get-Module -ListAvailable -Name ImportExcel)
}

BeforeAll {
                . "$PSScriptRoot/Fixtures/New-FaisTestData.ps1"
                Import-FaisModules
                if (Get-Module -ListAvailable -Name ImportExcel) { Import-Module ImportExcel -ErrorAction Stop }

                $script:W = New-TestSoftWeights

                # Students sheet laid out like MasterData: SPName in B, SimDebt in P
                function New-TStudentRow {
                                param ([string]$Name, [int]$Key = 1, $UnsatSim = $false, $UnsatFlight = $false, [double]$SimDebt = 0, $OnHold = $false)
                                return [PSCustomObject][ordered]@{
                                                SPKey = $Key; SPName = $Name; IsCap = $false; IsInternational = $false; Priority = 0
                                                UnsatSim = $UnsatSim; UnsatFlight = $UnsatFlight; OnHold = $OnHold
                                                Col9 = ''; Col10 = ''; Col11 = ''; Col12 = ''; Col13 = ''; Col14 = 'keep'; Col15 = ''
                                                SimDebt = $SimDebt
                                }
                }

                function New-TMasterData {
                                param (
                                                [string]$Path,
                                                [array]$Students,
                                                [array]$Absences = @(),
                                                [switch]$NoStudentsSheet
                                )
                                if (Test-Path $Path) { Remove-Item $Path -Force }
                                $Students | Export-Excel -Path $Path -WorksheetName $(if ($NoStudentsSheet) { 'NotStudents' } else { 'Students' }) 3>$null
                                @([PSCustomObject]@{ IPName = 'Jones, Bob'; IsCounter = $true; IsExperienced = $true; IsCheck = $true; Callsign = 'TALON'; Num = 31; OSC = ''; OSNum = '' }) |
                                                Export-Excel -Path $Path -WorksheetName 'Instructors' 3>$null
                                @(foreach ($m in (New-TestSyllabus)) { ConvertTo-SyllabusTabRow $m }) | Export-Excel -Path $Path -WorksheetName 'Syllabus' 3>$null
                                @([PSCustomObject]@{ Callsign = 'HAZE'; Num = 11 }, [PSCustomObject]@{ Callsign = 'HAZE'; Num = 12 }) | Export-Excel -Path $Path -WorksheetName 'FormCallsigns' 3>$null
                                @([PSCustomObject]@{
                                                                ContString = 10; String4 = -10; String5 = -100; Zipper = 5; NonCounter = -15; BreakString = -5; ExhaustMult = -1; ExhaustDays = 5
                                                                BITPerDay = 1.5; BITDays = 6; CAPBonus = 8; FormLookAhead = 22; LevelDifference = 2; CounterContinuity = 5; OFTIFTPen = -7
                                                                DollarRideBonus = 20; TransSimRatio = 0.5; FormSimRatio = 0.5; SimStopMult = 1
                                                }) | Export-Excel -Path $Path -WorksheetName 'SoftWeights' 3>$null
                                $abs = if ($Absences.Count -gt 0) { $Absences } else { @([PSCustomObject]@{ Student = 'Nobody, No'; StartTime = '2026-03-03 08:00'; EndTime = '2026-03-03 09:00'; Reason = 'x'; IsActive = $false }) }
                                $abs | Export-Excel -Path $Path -WorksheetName 'StudentAbsences' 3>$null
                }

                function Get-TCell {
                                param ([string]$Path, [string]$Sheet, [int]$Row, [int]$Col)
                                $pkg = Open-ExcelPackage -Path $Path
                                try { return $pkg.Workbook.Worksheets[$Sheet].Cells[$Row, $Col].Value }
                                finally { Close-ExcelPackage -ExcelPackage $pkg -NoSave }
                }

                # Row number of a mission on a student tab (Mission is column B)
                function Get-TTabRow {
                                param ($Worksheet, [string]$Mission)
                                for ($r = 2; $r -le $Worksheet.Dimension.End.Row; $r++) { if ([string]$Worksheet.Cells[$r, 2].Value -eq $Mission) { return $r } }
                                return $null
                }
}

Describe 'Get-UnsatFlagUpdates' {
                It 'lists only the flags that differ from what MasterData holds' {
                                $cleared = New-TestStudent -Name 'A' -UnsatSim
                                $cleared | Add-Member -NotePropertyName 'UnsatSimInMasterData' -NotePropertyValue $true
                                $cleared | Add-Member -NotePropertyName 'UnsatFlightInMasterData' -NotePropertyValue $false
                                $cleared.UnsatSim = $false
                                $hooked = New-TestStudent -Name 'B' -UnsatFlight
                                $hooked | Add-Member -NotePropertyName 'UnsatSimInMasterData' -NotePropertyValue $false
                                $hooked | Add-Member -NotePropertyName 'UnsatFlightInMasterData' -NotePropertyValue $false
                                $same = New-TestStudent -Name 'C' -UnsatSim
                                $same | Add-Member -NotePropertyName 'UnsatSimInMasterData' -NotePropertyValue $true
                                $same | Add-Member -NotePropertyName 'UnsatFlightInMasterData' -NotePropertyValue $false
                                $unknown = New-TestStudent -Name 'D' -UnsatSim

                                $u = Get-UnsatFlagUpdates -Students @($cleared, $hooked, $same, $unknown)
                                @($u.Keys | Sort-Object) | Should -Be @('A', 'B')
                                $u['A'].Count | Should -Be 1
                                $u['A']['UnsatSim'] | Should -BeFalse
                                $u['B'].Count | Should -Be 1
                                $u['B']['UnsatFlight'] | Should -BeTrue
                }

                It 'lists a HOLD change MasterData does not have yet' {
                                $held = New-TestStudent -Name 'E' -OnHold
                                $held | Add-Member -NotePropertyName 'OnHoldInMasterData' -NotePropertyValue $false
                                $same = New-TestStudent -Name 'F' -OnHold
                                $same | Add-Member -NotePropertyName 'OnHoldInMasterData' -NotePropertyValue $true
                                $u = Get-UnsatFlagUpdates -Students @($held, $same)
                                @($u.Keys) | Should -Be @('E')
                                $u['E'].Count | Should -Be 1
                                $u['E']['OnHold'] | Should -BeTrue
                }
}

Describe 'Update-StudentHoldStatus' {
                It 'changes only the students whose box changed and returns them for MasterData' {
                                $a = New-TestStudent -Name 'Adams, Al'
                                $b = New-TestStudent -Name 'Brown, Bea' -OnHold
                                $c = New-TestStudent -Name 'Cole, Cy' -OnHold
                                $d = New-TestStudent -Name 'Dunn, Di'
                                $changes = Update-StudentHoldStatus -HoldByName @{ 'Adams, Al' = $true; 'Brown, Bea' = $false; 'Cole, Cy' = $true } -Students @($a, $b, $c, $d) 6>$null
                                @($changes.Keys | Sort-Object) | Should -Be @('Adams, Al', 'Brown, Bea')
                                $changes['Adams, Al']['OnHold'] | Should -BeTrue
                                $changes['Brown, Bea']['OnHold'] | Should -BeFalse
                                @($a, $b, $c, $d | ForEach-Object { $_.OnHold }) | Should -Be @($true, $false, $true, $false)
                }

                It 'returns no changes when no box changed' {
                                $a = New-TestStudent -Name 'Adams, Al' -OnHold
                                (Update-StudentHoldStatus -HoldByName @{ 'Adams, Al' = $true } -Students @($a) 6>$null).Count | Should -Be 0
                                $a.OnHold | Should -BeTrue
                }
}

Describe 'Export-StudentSyllabusTabs' -Skip:(-not $script:HasExcel) {
                It 'gives a student with no tab a full tab that IngestData can read back' {
                                $path = Join-Path $TestDrive 'new.xlsx'
                                $s = New-TestStudent -Name 'Adams, Al' -Done 'FI3101|2026-02-20 08:00'
                                Export-StudentSyllabusTabs -Students @($s) -FilePath $path 3>$null 6>$null
                                $tab = @(Import-Excel -Path $path -WorksheetName 'Adams, Al')
                                $tab.Count | Should -Be 10
                                @($tab[0].PSObject.Properties.Name) | Should -Be @('SyllabusID', 'Mission', 'Prereqs', 'Instructor', 'ETD', 'Type', 'IsCheck', 'IsComplete', 'OptingEvents', 'Level')
                                @($tab | Where-Object { ConvertTo-FaisBool $_.IsComplete } | ForEach-Object { $_.Mission }) | Should -Be @('FI3101')
                }

                It 'updates an existing tab in place: only changed events'' cells, everything else kept' {
                                $path = Join-Path $TestDrive 'inplace.xlsx'
                                $s = New-TestStudent -Name 'Adams, Al' -Done 'FI3101|2026-02-20 08:00', 'FE3102|2026-02-21 08:00'
                                Export-StudentSyllabusTabs -Students @($s) -FilePath $path 3>$null 6>$null

                                # Hand edits FAIS doesn't own: an extra column, a note row, a changed Prereqs cell, bold text
                                $pkg = Open-ExcelPackage -Path $path
                                $ws = $pkg.Workbook.Worksheets['Adams, Al']
                                $ws.Cells[1, 11].Value = 'Grade'
                                $ws.Cells[2, 11].Value = 'Good'
                                $ws.Cells[12, 1].Value = 'Phase 2 starts here'
                                $ws.Cells[(Get-TTabRow $ws 'FT5001'), 3].Value = 'FI3101, FE3102'
                                $ws.Cells[(Get-TTabRow $ws 'FE3102'), 2].Style.Font.Bold = $true
                                Close-ExcelPackage -ExcelPackage $pkg

                                # This run: FT5001 flown, FE3102 logged incomplete
                                $null = Complete-StudentMission -Student $s -Mission 'FT5001' -InstructorName 'Jones, Bob' -ETD '2026-03-03 07:30' -SoftWeights $script:W
                                $null = Undo-StudentMission -Student $s -Mission 'FE3102' -SoftWeights $script:W
                                Export-StudentSyllabusTabs -Students @($s) -FilePath $path 3>$null 6>$null

                                $pkg = Open-ExcelPackage -Path $path
                                try {
                                                $ws = $pkg.Workbook.Worksheets['Adams, Al']
                                                $ft = Get-TTabRow $ws 'FT5001'
                                                $fe = Get-TTabRow $ws 'FE3102'
                                                $fi = Get-TTabRow $ws 'FI3101'
                                                $ws.Cells[$ft, 8].Value | Should -BeTrue
                                                $ws.Cells[$ft, 4].Value | Should -Be 'Jones, Bob'
                                                $ws.Cells[$ft, 5].Value | Should -Be '2026-03-03 07:30:00'
                                                $ws.Cells[$fe, 8].Value | Should -BeFalse
                                                $ws.Cells[$fe, 4].Value | Should -BeNullOrEmpty
                                                $ws.Cells[$fe, 5].Value | Should -BeNullOrEmpty
                                                # Untouched event, hand edits and formatting are all still there
                                                $ws.Cells[$fi, 8].Value | Should -BeTrue
                                                $ws.Cells[$fi, 5].Value | Should -Be '2026-02-20 08:00:00'
                                                $ws.Cells[1, 11].Value | Should -Be 'Grade'
                                                $ws.Cells[2, 11].Value | Should -Be 'Good'
                                                $ws.Cells[12, 1].Value | Should -Be 'Phase 2 starts here'
                                                $ws.Cells[$ft, 3].Value | Should -Be 'FI3101, FE3102'
                                                $ws.Cells[$fe, 2].Style.Font.Bold | Should -BeTrue
                                                $ws.Dimension.End.Row | Should -Be 12
                                }
                                finally { Close-ExcelPackage -ExcelPackage $pkg -NoSave }
                }

                It 'writes nothing to a tab when no event changed' {
                                $path = Join-Path $TestDrive 'same.xlsx'
                                $s = New-TestStudent -Name 'Adams, Al' -Done 'FI3101|2026-02-20 08:00'
                                Export-StudentSyllabusTabs -Students @($s) -FilePath $path 3>$null 6>$null
                                $out = Export-StudentSyllabusTabs -Students @($s) -FilePath $path 3>$null 6>&1
                                ($out | Out-String) | Should -Match '\(0 cell\(s\) changed, 0 new tab\(s\)\)'
                }

                It 'adds an event missing from the tab at the bottom' {
                                $path = Join-Path $TestDrive 'missing.xlsx'
                                $s = New-TestStudent -Name 'Adams, Al'
                                Export-StudentSyllabusTabs -Students @($s) -FilePath $path 3>$null 6>$null
                                $s.Syllabus += [PSCustomObject]@{ SyllabusID = 11; Mission = 'FT5299'; Prereqs = 'FT5190'; Instructor = 'Jones, Bob'; ETD = '2026-03-03 07:30:00'; Type = 'Single'; IsCheck = $false; IsComplete = $true; NumOpting = 0; Level = 5 }
                                Export-StudentSyllabusTabs -Students @($s) -FilePath $path 3>$null 6>$null
                                $tab = @(Import-Excel -Path $path -WorksheetName 'Adams, Al')
                                $tab.Count | Should -Be 11
                                $tab[10].Mission | Should -Be 'FT5299'
                                ConvertTo-FaisBool $tab[10].IsComplete | Should -BeTrue
                }

                It 'matches a mission listed twice on the tab in order' {
                                $path = Join-Path $TestDrive 'twice.xlsx'
                                $s = New-TestStudent -Name 'Adams, Al'
                                $s.Syllabus += [PSCustomObject]@{ SyllabusID = 11; Mission = 'FT5001'; Prereqs = 'FT5190'; Instructor = ''; ETD = ''; Type = 'Single'; IsCheck = $false; IsComplete = $false; NumOpting = 0; Level = 5 }
                                Export-StudentSyllabusTabs -Students @($s) -FilePath $path 3>$null 6>$null
                                $second = $s.Syllabus[10]
                                $second.IsComplete = $true
                                $second.Instructor = 'Jones, Bob'
                                $second.ETD = '2026-03-03 07:30:00'
                                Export-StudentSyllabusTabs -Students @($s) -FilePath $path 3>$null 6>$null
                                $tab = @(Import-Excel -Path $path -WorksheetName 'Adams, Al')
                                $tab.Count | Should -Be 11
                                ConvertTo-FaisBool $tab[2].IsComplete | Should -BeFalse
                                ConvertTo-FaisBool $tab[10].IsComplete | Should -BeTrue
                                $tab[10].Instructor | Should -Be 'Jones, Bob'
                }

                It 'throws, and saves nothing, when a tab has no Mission heading' {
                                $path = Join-Path $TestDrive 'nomission.xlsx'
                                @([PSCustomObject]@{ Event = 'FI3101'; IsComplete = $false }) | Export-Excel -Path $path -WorksheetName 'Adams, Al' 3>$null
                                $before = [System.IO.File]::ReadAllBytes($path)
                                $s = New-TestStudent -Name 'Adams, Al' -Done 'FI3101|2026-02-20 08:00'
                                { Export-StudentSyllabusTabs -Students @($s) -FilePath $path 3>$null 6>$null } | Should -Throw '*no Mission heading*'
                                [System.IO.File]::ReadAllBytes($path) | Should -Be $before
                }
}

Describe 'Update-MasterDataStudents' -Skip:(-not $script:HasExcel) {
                BeforeEach {
                                $script:MD = Join-Path $TestDrive 'MasterData.xlsx'
                                New-TMasterData -Path $script:MD -Students @(
                                                (New-TStudentRow -Name 'Adams, Al' -Key 1 -SimDebt 1)
                                                (New-TStudentRow -Name 'Brown, Bea' -Key 2 -UnsatSim $true -SimDebt 2)
                                                (New-TStudentRow -Name 'Cole, Cy' -Key 3 -SimDebt 3)
                                )
                }

                It 'writes SimDebt to column P and unsat flags to their headed columns, nothing else' {
                                Update-MasterDataStudents -MasterDataPath $script:MD -SimDebtByName @{ 'Adams, Al' = 2.5; 'Brown, Bea' = 1 } -FlagsByName @{ 'Brown, Bea' = @{ UnsatSim = $false }; 'Cole, Cy' = @{ UnsatFlight = $true } } 6>$null
                                $rows = @(Import-Excel -Path $script:MD -WorksheetName 'Students')
                                $rows[0].SimDebt | Should -Be 2.5
                                $rows[1].SimDebt | Should -Be 1
                                $rows[2].SimDebt | Should -Be 3
                                ConvertTo-FaisBool $rows[1].UnsatSim | Should -BeFalse
                                ConvertTo-FaisBool $rows[2].UnsatFlight | Should -BeTrue
                                ConvertTo-FaisBool $rows[0].UnsatSim | Should -BeFalse
                                $rows[1].Col14 | Should -Be 'keep'
                                (Get-TCell -Path $script:MD -Sheet 'Students' -Row 1 -Col 16) | Should -Be 'SimDebt'
                }

                It 'writes OnHold to its headed column, on and off, nothing else' {
                                $path = Join-Path $TestDrive 'Hold.xlsx'
                                New-TMasterData -Path $path -Students @((New-TStudentRow -Name 'Adams, Al' -Key 1 -SimDebt 1), (New-TStudentRow -Name 'Brown, Bea' -Key 2 -OnHold $true -SimDebt 2))
                                Update-MasterDataStudents -MasterDataPath $path -FlagsByName @{ 'Adams, Al' = @{ OnHold = $true }; 'Brown, Bea' = @{ OnHold = $false } } 6>$null
                                $rows = @(Import-Excel -Path $path -WorksheetName 'Students')
                                ConvertTo-FaisBool $rows[0].OnHold | Should -BeTrue
                                ConvertTo-FaisBool $rows[1].OnHold | Should -BeFalse
                                ConvertTo-FaisBool $rows[0].UnsatSim | Should -BeFalse
                                $rows[0].SimDebt | Should -Be 1
                                $rows[1].SimDebt | Should -Be 2
                }

                It 'warns about names it can''t find and still writes the rest' {
                                $w = Update-MasterDataStudents -MasterDataPath $script:MD -SimDebtByName @{ 'Adams, Al' = 9; 'Nobody, No' = 1 } 6>$null 3>&1
                                ($w | Out-String) | Should -Match 'Nobody, No'
                                @(Import-Excel -Path $script:MD -WorksheetName 'Students')[0].SimDebt | Should -Be 9
                }

                It 'throws, and saves nothing, when an unsat heading is missing' {
                                $path = Join-Path $TestDrive 'NoFlags.xlsx'
                                @([PSCustomObject][ordered]@{ SPKey = 1; SPName = 'Adams, Al'; SimDebt = 1 }) | Export-Excel -Path $path -WorksheetName 'Students' 3>$null
                                $before = [System.IO.File]::ReadAllBytes($path)
                                { Update-MasterDataStudents -MasterDataPath $path -FlagsByName @{ 'Adams, Al' = @{ UnsatSim = $false } } 6>$null } | Should -Throw "*no 'UnsatSim' heading*"
                                [System.IO.File]::ReadAllBytes($path) | Should -Be $before
                }

                It 'throws when MasterData or its Students sheet is missing' {
                                { Update-MasterDataStudents -MasterDataPath (Join-Path $TestDrive 'nope.xlsx') -SimDebtByName @{ 'A' = 1 } } | Should -Throw '*not found*'
                                $path = Join-Path $TestDrive 'NoStudents.xlsx'
                                New-TMasterData -Path $path -Students @(New-TStudentRow -Name 'Adams, Al') -NoStudentsSheet
                                { Update-MasterDataStudents -MasterDataPath $path -SimDebtByName @{ 'Adams, Al' = 1 } } | Should -Throw "*'Students' not found*"
                }
}

Describe 'Export-FaisOutput' -Skip:(-not $script:HasExcel) {
                BeforeAll {
                                function New-TFillResult {
                                                $a = New-TestStudent -Name 'Adams, Al' -Done 'FI3101|2026-02-20 08:00', 'FT5001|2026-03-03 07:30|Jones, Bob' -SimDebt 4
                                                $a | Add-Member -NotePropertyName 'UnsatSimInMasterData' -NotePropertyValue $true
                                                $a | Add-Member -NotePropertyName 'UnsatFlightInMasterData' -NotePropertyValue $false
                                                $b = New-TestStudent -Name 'Brown, Bea' -Done 'FI3101|2026-02-20 08:00' -SimDebt 2
                                                return [PSCustomObject]@{
                                                                Schedule = @(
                                                                                (New-TestRow -ETD '2026-03-03 07:30' -Instructor 'Jones, Bob' -Student 'Adams, Al' -Mission 'FT5001' -Notes 'Clears Unsat Sim (OFT)')
                                                                                (New-TestRow -ETD '2026-03-03 09:00' -SortieType 'OFT' -Student 'For Sale')
                                                                )
                                                                Students = @($a, $b)
                                                }
                                }
                }

                It 'writes MasterSchedule, the student tabs, SimDebt and the cleared unsat' {
                                $md = Join-Path $TestDrive 'MD-ok.xlsx'
                                $out = Join-Path $TestDrive 'Out-ok.xlsx'
                                New-TMasterData -Path $md -Students @((New-TStudentRow -Name 'Adams, Al' -Key 1 -UnsatSim $true -SimDebt 3), (New-TStudentRow -Name 'Brown, Bea' -Key 2 -SimDebt 2))
                                Export-FaisOutput -FillResult (New-TFillResult) -OutputPath $out -MasterDataPath $md 3>$null 6>$null | Should -BeTrue

                                $sched = @(Import-Excel -Path $out -WorksheetName 'MasterSchedule')
                                $sched.Count | Should -Be 2
                                $sched[0].Student | Should -Be 'Adams, Al'
                                @(Import-Excel -Path $out -WorksheetName 'Adams, Al' | Where-Object { ConvertTo-FaisBool $_.IsComplete } | ForEach-Object { $_.Mission }) | Should -Be @('FI3101', 'FT5001')
                                @(Import-Excel -Path $out -WorksheetName 'Brown, Bea').Count | Should -Be 10
                                $students = @(Import-Excel -Path $md -WorksheetName 'Students')
                                $students[0].SimDebt | Should -Be 4
                                ConvertTo-FaisBool $students[0].UnsatSim | Should -BeFalse
                                $students[1].SimDebt | Should -Be 2
                }

                It 'writes nothing to either file when one of the writes fails' {
                                $md = Join-Path $TestDrive 'MD-bad.xlsx'
                                $out = Join-Path $TestDrive 'Out-bad.xlsx'
                                New-TMasterData -Path $md -Students @(New-TStudentRow -Name 'Adams, Al') -NoStudentsSheet
                                @([PSCustomObject]@{ ETD = 'yesterday'; Student = 'old' }) | Export-Excel -Path $out -WorksheetName 'MasterSchedule' 3>$null
                                $outBefore = [System.IO.File]::ReadAllBytes($out)
                                $mdBefore = [System.IO.File]::ReadAllBytes($md)

                                $warn = Export-FaisOutput -FillResult (New-TFillResult) -OutputPath $out -MasterDataPath $md 6>$null 3>&1
                                @($warn | Where-Object { $_ -is [bool] }) | Should -Be @($false)
                                ($warn | Out-String) | Should -Match 'left as they were'
                                [System.IO.File]::ReadAllBytes($out) | Should -Be $outBefore
                                [System.IO.File]::ReadAllBytes($md) | Should -Be $mdBefore
                }

                It 'removes an OutputSchedule it created when the export fails' {
                                $md = Join-Path $TestDrive 'MD-bad2.xlsx'
                                $out = Join-Path $TestDrive 'Out-new.xlsx'
                                New-TMasterData -Path $md -Students @(New-TStudentRow -Name 'Adams, Al') -NoStudentsSheet
                                Export-FaisOutput -FillResult (New-TFillResult) -OutputPath $out -MasterDataPath $md 3>$null 6>$null | Should -BeFalse
                                Test-Path $out | Should -BeFalse
                }

                It 'writes nothing when MasterData is missing' {
                                $out = Join-Path $TestDrive 'Out-nomd.xlsx'
                                Export-FaisOutput -FillResult (New-TFillResult) -OutputPath $out -MasterDataPath (Join-Path $TestDrive 'none.xlsx') 3>$null 6>$null | Should -BeFalse
                                Test-Path $out | Should -BeFalse
                }
}

Describe 'Pre-run saves (incomplete and hooked events)' -Skip:(-not $script:HasExcel) {
                BeforeAll {
                                $script:PMD = Join-Path $TestDrive 'Pre-MD.xlsx'
                                $script:POut = Join-Path $TestDrive 'Pre-Out.xlsx'
                                New-TMasterData -Path $script:PMD -Students @((New-TStudentRow -Name 'Adams, Al' -Key 1 -SimDebt 3), (New-TStudentRow -Name 'Brown, Bea' -Key 2 -SimDebt 1))
                                $a = New-TestStudent -Name 'Adams, Al' -Done 'FI3101|2026-02-20 08:00', 'FE3102|2026-03-02 07:30', 'FT5001|2026-03-02 12:00|Jones, Bob'
                                Export-StudentSyllabusTabs -Students @($a) -FilePath $script:POut 3>$null 6>$null
                                $prevRows = @(
                                                (New-TestRow -ETD '2026-03-02 07:30' -SortieType 'OFT' -Student 'Adams, Al' -Mission 'FE3102')
                                                (New-TestRow -ETD '2026-03-02 12:00' -Instructor 'Jones, Bob' -Student 'Adams, Al' -Mission 'FT5001')
                                )
                                Export-FormattedMasterSchedule -ScheduleData (ConvertTo-ExportRows -Schedule $prevRows) -FilePath $script:POut 3>$null 6>$null

                                $script:PIng = IngestData -excelInputPath $script:PMD -excelOutputPath $script:POut 3>$null 6>$null
                                $script:Prev = @(DownloadPrevSchedule -Students $script:PIng.Students -Instructors $script:PIng.Instructors -excelOutputPath $script:POut 3>$null)
                                function Get-PRow { param ([string]$Mission) return ($script:Prev | Where-Object { (Get-NameText $_.Mission) -eq $Mission } | Select-Object -First 1) }
                }

                It 'an incomplete event backs out its sim debt and saves MasterData and that event''s tab cells' {
                                $changed = @(Process-IncompleteEvents -SelectedItems @([PSCustomObject]@{ Sortie = (Get-PRow 'FT5001') }) -Students $script:PIng.Students -SoftWeights $script:PIng.SoftWeights 3>$null 6>$null)
                                @($changed | ForEach-Object { $_.Name }) | Should -Be @('Adams, Al')
                                $changed[0].SimDebt | Should -Be 2
                                Save-PreRunChanges -MasterDataPath $script:PMD -OutputPath $script:POut -SimDebtByName @{ 'Adams, Al' = 2 } -TabStudents $changed 3>$null 6>$null | Should -BeTrue

                                @(Import-Excel -Path $script:PMD -WorksheetName 'Students')[0].SimDebt | Should -Be 2
                                @(Import-Excel -Path $script:PMD -WorksheetName 'Students')[1].SimDebt | Should -Be 1
                                $tab = @(Import-Excel -Path $script:POut -WorksheetName 'Adams, Al')
                                @($tab | Where-Object { ConvertTo-FaisBool $_.IsComplete } | ForEach-Object { $_.Mission }) | Should -Be @('FI3101', 'FE3102')
                                $ft = $tab | Where-Object { $_.Mission -eq 'FT5001' }
                                $ft.Instructor | Should -BeNullOrEmpty
                                $ft.ETD | Should -BeNullOrEmpty
                                # The MasterSchedule tab is left alone until Export
                                @(Import-Excel -Path $script:POut -WorksheetName 'MasterSchedule').Count | Should -Be 2
                }

                It 'logging the same incomplete event again changes nothing' {
                                $row = New-TestRow -ETD '2026-03-02 12:00' -Instructor 'Jones, Bob' -Student 'Adams, Al' -Mission 'FT5001'
                                $changed = @(Process-IncompleteEvents -SelectedItems @([PSCustomObject]@{ Sortie = $row }) -Students $script:PIng.Students -SoftWeights $script:PIng.SoftWeights 3>$null 6>$null)
                                $changed.Count | Should -Be 0
                                ($script:PIng.Students | Where-Object { $_.Name -eq 'Adams, Al' }).SimDebt | Should -Be 2
                }

                It 'a hooked sim sets UnsatSim TRUE in MasterData with the sim type on the student' {
                                $hooked = @(Process-HookedEvents -SelectedItems @([PSCustomObject]@{ Sortie = (Get-PRow 'FE3102') }) -Students $script:PIng.Students 6>$null)
                                @($hooked | ForEach-Object { $_.Name }) | Should -Be @('Adams, Al')
                                $hooked[0].UnsatSim | Should -BeTrue
                                $hooked[0].UnsatSimType | Should -Be 'OFT'
                                $flags = Get-UnsatFlagUpdates -Students $hooked
                                $flags['Adams, Al']['UnsatSim'] | Should -BeTrue
                                Save-PreRunChanges -MasterDataPath $script:PMD -OutputPath $script:POut -FlagsByName $flags 3>$null 6>$null | Should -BeTrue
                                $rows = @(Import-Excel -Path $script:PMD -WorksheetName 'Students')
                                ConvertTo-FaisBool $rows[0].UnsatSim | Should -BeTrue
                                ConvertTo-FaisBool $rows[0].UnsatFlight | Should -BeFalse
                                ConvertTo-FaisBool $rows[1].UnsatSim | Should -BeFalse
                }

                It 'a HOLD change sets OnHold in MasterData right away, on and then off' {
                                $students = $script:PIng.Students
                                ($students | Where-Object { $_.Name -eq 'Adams, Al' }).OnHoldInMasterData | Should -BeFalse
                                $changes = Update-StudentHoldStatus -HoldByName @{ 'Adams, Al' = $true; 'Brown, Bea' = $false } -Students $students 6>$null
                                @($changes.Keys) | Should -Be @('Adams, Al')
                                # Not saved yet: Export would write it
                                (Get-UnsatFlagUpdates -Students $students)['Adams, Al']['OnHold'] | Should -BeTrue
                                Save-PreRunChanges -MasterDataPath $script:PMD -OutputPath $script:POut -FlagsByName $changes 3>$null 6>$null | Should -BeTrue
                                $rows = @(Import-Excel -Path $script:PMD -WorksheetName 'Students')
                                ConvertTo-FaisBool $rows[0].OnHold | Should -BeTrue
                                ConvertTo-FaisBool $rows[1].OnHold | Should -BeFalse

                                $changes = Update-StudentHoldStatus -HoldByName @{ 'Adams, Al' = $false } -Students $students 6>$null
                                Save-PreRunChanges -MasterDataPath $script:PMD -OutputPath $script:POut -FlagsByName $changes 3>$null 6>$null | Should -BeTrue
                                ConvertTo-FaisBool @(Import-Excel -Path $script:PMD -WorksheetName 'Students')[0].OnHold | Should -BeFalse
                                (IngestData -excelInputPath $script:PMD -excelOutputPath $script:POut 3>$null 6>$null).Students[0].OnHold | Should -BeFalse
                }

                It 'keeps the change in memory, and saves nothing, when MasterData is missing' {
                                Save-PreRunChanges -MasterDataPath (Join-Path $TestDrive 'gone.xlsx') -OutputPath $script:POut -SimDebtByName @{ 'Adams, Al' = 1 } 3>$null 6>$null | Should -BeFalse
                }
}

Describe 'IngestData' -Skip:(-not $script:HasExcel) {
                BeforeAll {
                                $script:IMD = Join-Path $TestDrive 'Ingest-MD.xlsx'
                                $script:IOut = Join-Path $TestDrive 'Ingest-Out.xlsx'
                                $rows = @(
                                                (New-TStudentRow -Name 'Adams, Al' -Key 1 -UnsatSim 'TRUE' -SimDebt 1.5)
                                                (New-TStudentRow -Name 'Lee, Ann' -Key 2 -UnsatFlight $true)
                                                (New-TStudentRow -Name 'Leech, Tom' -Key 3)
                                )
                                $rows += @(for ($i = 4; $i -le 105; $i++) { New-TStudentRow -Name ('Student{0:000}, Sam' -f $i) -Key $i })
                                $abs = @(
                                                [PSCustomObject]@{ Student = 'adams, al'; StartTime = '2026-03-03 08:00'; EndTime = '2026-03-03 10:00'; Reason = 'Dentist'; IsActive = $true }
                                                [PSCustomObject]@{ Student = 'Lee'; StartTime = '2026-03-03 12:00'; EndTime = '2026-03-03 14:00'; Reason = 'Typo'; IsActive = $true }
                                )
                                New-TMasterData -Path $script:IMD -Students $rows -Absences $abs

                                # Adams has a tab: the last sim flown is the OFT FE3102
                                $a = New-TestStudent -Name 'Adams, Al' -Done 'FI3101|2026-02-20 08:00', 'FE3102|2026-02-21 08:00', 'FT5001|2026-02-23 08:00|Jones, Bob'
                                Export-StudentSyllabusTabs -Students @($a) -FilePath $script:IOut 3>$null 6>$null

                                $script:Ing = IngestData -excelInputPath $script:IMD -excelOutputPath $script:IOut 3>$null 6>$null
                                function Get-IStudent { param ([string]$Name) return ($script:Ing.Students | Where-Object { $_.Name -eq $Name } | Select-Object -First 1) }
                }

                It 'loads the first 100 students' {
                                @($script:Ing.Students).Count | Should -Be 100
                                $script:Ing.Students[99].Name | Should -Be 'Student100, Sam'
                }

                It 'reads a student''s tab when there is one, else the blank Syllabus' {
                                @((Get-IStudent 'Adams, Al').Syllabus | Where-Object { $_.IsComplete } | ForEach-Object { $_.Mission }) | Should -Be @('FI3101', 'FE3102', 'FT5001')
                                @((Get-IStudent 'Lee, Ann').Syllabus | Where-Object { $_.IsComplete }).Count | Should -Be 0
                                @((Get-IStudent 'Lee, Ann').Syllabus).Count | Should -Be 10
                }

                It 'attaches absences by exact name only (case aside)' {
                                @((Get-IStudent 'Adams, Al').Absences).Count | Should -Be 1
                                (Get-IStudent 'Adams, Al').Absences[0].Reason | Should -Be 'Dentist'
                                @((Get-IStudent 'Lee, Ann').Absences).Count | Should -Be 0
                                @((Get-IStudent 'Leech, Tom').Absences).Count | Should -Be 0
                }

                It 'records the unsat flags as read, and the hooked sim as the latest one flown' {
                                $adams = Get-IStudent 'Adams, Al'
                                $adams.UnsatSim | Should -BeTrue
                                $adams.UnsatSimType | Should -Be 'OFT'
                                $adams.UnsatSimInMasterData | Should -BeTrue
                                $adams.UnsatFlightInMasterData | Should -BeFalse
                                $lee = Get-IStudent 'Lee, Ann'
                                $lee.UnsatFlight | Should -BeTrue
                                $lee.UnsatSimType | Should -Be ''
                                $lee.UnsatFlightInMasterData | Should -BeTrue
                                $adams.SimDebt | Should -Be 1.5
                }
}

Describe 'Get-PendingStudentAbsences' {
                It 'lists only the absences added this run that MasterData does not have yet, each once' {
                                $read = [PSCustomObject]@{ Student = 'A'; StartTime = [datetime]'2026-03-03 08:00'; EndTime = [datetime]'2026-03-03 09:00'; Reason = '' }
                                $saved = [PSCustomObject]@{ Student = 'A'; StartTime = [datetime]'2026-03-03 10:00'; EndTime = [datetime]'2026-03-03 11:00'; Reason = ''; InMasterData = $true }
                                $new = [PSCustomObject]@{ Student = 'A'; StartTime = [datetime]'2026-03-03 12:00'; EndTime = [datetime]'2026-03-03 13:00'; Reason = ''; InMasterData = $false }
                                $a = New-TestStudent -Name 'A' -Absences @($read, $saved, $new)
                                $pending = @(Get-PendingStudentAbsences -Students @($a, (Copy-StudentForEval $a), $null, (New-TestStudent -Name 'B')))
                                $pending.Count | Should -Be 1
                                [object]::ReferenceEquals($pending[0], $new) | Should -BeTrue
                }
}

Describe 'Add Student Absence: lines on the StudentAbsences sheet' -Skip:(-not $script:HasExcel) {
                BeforeAll {
                                function New-TAbsence {
                                                param ([string]$Student, [string]$From, [string]$To, [string]$Reason = '')
                                                return [PSCustomObject]@{ Student = $Student; StartTime = [datetime]$From; EndTime = [datetime]$To; Reason = $Reason; InMasterData = $false }
                                }
                                $script:AbsStudents = @((New-TStudentRow -Name 'Adams, Al' -Key 1), (New-TStudentRow -Name 'Brown, Bea' -Key 2))
                                $script:OldAbsence = [PSCustomObject]@{ Student = 'Adams, Al'; StartTime = '2026-03-02 08:00'; EndTime = '2026-03-02 09:00'; Reason = 'Old'; IsActive = $true }
                }

                It 'adds a line under the last one with IsActive TRUE, and IngestData gives it to the student' {
                                $md = Join-Path $TestDrive 'Abs-MD.xlsx'
                                New-TMasterData -Path $md -Students $script:AbsStudents -Absences @($script:OldAbsence)
                                Add-MasterDataAbsences -MasterDataPath $md -Absences @((New-TAbsence 'Brown, Bea' '2026-03-03 13:00' '2026-03-03 15:30' ' Dentist '), (New-TAbsence 'Adams, Al' '2026-03-03 22:00' '2026-03-04 06:00')) 6>$null
                                $rows = @(Import-Excel -Path $md -WorksheetName 'StudentAbsences')
                                $rows.Count | Should -Be 3
                                $rows[0].Reason | Should -Be 'Old'
                                $rows[1].Student | Should -Be 'Brown, Bea'
                                $rows[1].Reason | Should -Be 'Dentist'
                                ConvertTo-FaisBool $rows[1].IsActive | Should -BeTrue
                                (Parse-ETD $rows[2].EndTime) | Should -Be ([datetime]'2026-03-04 06:00')
                                [string]::IsNullOrWhiteSpace([string]$rows[2].Reason) | Should -BeTrue

                                $ing = IngestData -excelInputPath $md -excelOutputPath (Join-Path $TestDrive 'Abs-Out.xlsx') 3>$null 6>$null
                                $brown = $ing.Students | Where-Object { $_.Name -eq 'Brown, Bea' }
                                @($brown.Absences).Count | Should -Be 1
                                $brown.Absences[0].StartTime | Should -Be ([datetime]'2026-03-03 13:00')
                                $brown.Absences[0].EndTime | Should -Be ([datetime]'2026-03-03 15:30')
                                $brown.Absences[0].Reason | Should -Be 'Dentist'
                                @(($ing.Students | Where-Object { $_.Name -eq 'Adams, Al' }).Absences).Count | Should -Be 2
                }

                It 'fills in blank lines formatted ahead of time, keeping their date format' {
                                $md = Join-Path $TestDrive 'Abs-Formatted.xlsx'
                                New-TMasterData -Path $md -Students $script:AbsStudents -Absences @($script:OldAbsence)
                                $pkg = Open-ExcelPackage -Path $md
                                $ws = $pkg.Workbook.Worksheets['StudentAbsences']
                                foreach ($r in 3..6) { foreach ($c in 2..3) { $ws.Cells[$r, $c].Style.Numberformat.Format = 'd mmm yy hh:mm' } }
                                Close-ExcelPackage -ExcelPackage $pkg
                                Add-MasterDataAbsences -MasterDataPath $md -Absences @(New-TAbsence 'Brown, Bea' '2026-03-03 13:00' '2026-03-03 15:30') 6>$null
                                $pkg = Open-ExcelPackage -Path $md
                                try {
                                                $ws = $pkg.Workbook.Worksheets['StudentAbsences']
                                                $ws.Cells[3, 1].Value | Should -Be 'Brown, Bea'
                                                $ws.Cells[3, 2].Style.Numberformat.Format | Should -Be 'd mmm yy hh:mm'
                                                $ws.Cells[4, 1].Value | Should -BeNullOrEmpty
                                }
                                finally { Close-ExcelPackage -ExcelPackage $pkg -NoSave }
                }

                It 'adds the sheet, with its headings, when MasterData has none' {
                                $md = Join-Path $TestDrive 'Abs-NoSheet.xlsx'
                                New-TMasterData -Path $md -Students $script:AbsStudents
                                $pkg = Open-ExcelPackage -Path $md
                                $pkg.Workbook.Worksheets.Delete('StudentAbsences')
                                Close-ExcelPackage -ExcelPackage $pkg
                                Add-MasterDataAbsences -MasterDataPath $md -Absences @(New-TAbsence 'Adams, Al' '2026-03-03 08:00' '2026-03-03 09:00' 'Leave') 6>$null
                                $rows = @(Import-Excel -Path $md -WorksheetName 'StudentAbsences')
                                @($rows[0].PSObject.Properties.Name) | Should -Be @('Student', 'StartTime', 'EndTime', 'Reason', 'IsActive')
                                $rows.Count | Should -Be 1
                                $rows[0].Student | Should -Be 'Adams, Al'
                }

                It 'throws, and saves nothing, when a heading IngestData reads is missing' {
                                $path = Join-Path $TestDrive 'Abs-NoActive.xlsx'
                                @([PSCustomObject][ordered]@{ Student = 'Adams, Al'; StartTime = '2026-03-02 08:00'; EndTime = '2026-03-02 09:00' }) | Export-Excel -Path $path -WorksheetName 'StudentAbsences' 3>$null
                                $before = [System.IO.File]::ReadAllBytes($path)
                                { Add-MasterDataAbsences -MasterDataPath $path -Absences @(New-TAbsence 'Brown, Bea' '2026-03-03 13:00' '2026-03-03 14:00') 6>$null } | Should -Throw "*no 'IsActive' heading*"
                                [System.IO.File]::ReadAllBytes($path) | Should -Be $before
                }

                It 'Save-StudentAbsences saves what MasterData does not have yet and marks it saved; a missing file keeps it for Export' {
                                $md = Join-Path $TestDrive 'Abs-Save.xlsx'
                                New-TMasterData -Path $md -Students $script:AbsStudents -Absences @($script:OldAbsence)
                                $a = New-TestStudent -Name 'Adams, Al'
                                $b = New-TestStudent -Name 'Brown, Bea'
                                $made = New-StudentAbsences -Students @('Adams, Al', 'Brown, Bea') -Day '2026-03-03' -From '1300' -To '0230 PM' -Reason 'Brief'
                                [void](Add-StudentAbsences -Absences $made.Absences -Students @($a, $b))

                                $r = Save-StudentAbsences -MasterDataPath $md -Students @($a, $b) 6>$null
                                $r.Saved | Should -BeTrue
                                $r.Count | Should -Be 2
                                @(Get-PendingStudentAbsences -Students @($a, $b)).Count | Should -Be 0
                                @(Import-Excel -Path $md -WorksheetName 'StudentAbsences' | ForEach-Object { $_.Student }) | Should -Be @('Adams, Al', 'Adams, Al', 'Brown, Bea')
                                (Save-StudentAbsences -MasterDataPath $md -Students @($a, $b) 6>$null).Count | Should -Be 0 -Because 'nothing is left to save'

                                $later = New-StudentAbsences -Students @('Brown, Bea') -Day '2026-03-04' -From '0800' -To '0900'
                                [void](Add-StudentAbsences -Absences $later.Absences -Students @($b))
                                $gone = Save-StudentAbsences -MasterDataPath (Join-Path $TestDrive 'nope.xlsx') -Students @($a, $b) 6>$null
                                $gone.Saved | Should -BeFalse
                                $gone.Message | Should -Be 'MasterData.xlsx was not found'
                                @(Get-PendingStudentAbsences -Students @($a, $b)).Count | Should -Be 1
                }

                It 'Export writes an absence that could not be saved before, with everything else' {
                                $md = Join-Path $TestDrive 'Abs-Export-MD.xlsx'
                                $out = Join-Path $TestDrive 'Abs-Export-Out.xlsx'
                                New-TMasterData -Path $md -Students $script:AbsStudents -Absences @($script:OldAbsence)
                                $a = New-TestStudent -Name 'Adams, Al' -Done 'FI3101|2026-02-20 08:00'
                                $made = New-StudentAbsences -Students @('Adams, Al') -Day '2026-03-03' -From '1500' -To '1600' -Reason 'Medical'
                                [void](Add-StudentAbsences -Absences $made.Absences -Students @($a))
                                $fill = [PSCustomObject]@{ Schedule = @(New-TestRow -ETD '2026-03-03 07:30' -Instructor 'Jones, Bob' -Student 'Adams, Al' -Mission 'FT5001'); Students = @(Copy-StudentForEval $a) }
                                Export-FaisOutput -FillResult $fill -OutputPath $out -MasterDataPath $md 3>$null 6>$null | Should -BeTrue
                                $rows = @(Import-Excel -Path $md -WorksheetName 'StudentAbsences')
                                $rows.Count | Should -Be 2
                                $rows[1].Reason | Should -Be 'Medical'
                                $a.Absences[0].InMasterData | Should -BeTrue
                }
}
