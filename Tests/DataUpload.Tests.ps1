# Pester 5 tests for ATIS.xlsx: Export-AtisWorkbook and its call from Export-FaisOutput
# (Modules\DataUpload.psm1). Needs ImportExcel; skipped without it.
# Day 1 is Wed 2026-10-21, day 2 is Thu 2026-10-22.

BeforeDiscovery {
                $script:NoExcel = $null -eq (Get-Module -ListAvailable -Name ImportExcel)
}

BeforeAll {
                . "$PSScriptRoot/Fixtures/New-FaisTestData.ps1"
                Import-FaisModules
                Import-Module ImportExcel

                $script:Day1 = Get-DaySheetName -Date ([datetime]'2026-10-21')
                $script:Day2 = Get-DaySheetName -Date ([datetime]'2026-10-22')

                # A MasterSchedule-shaped row (what ConvertTo-ExportRows returns)
                function New-AtisRow {
                                param (
                                                [string]$ETD,
                                                [string]$SortieType = 'Day Local',
                                                [string]$Instructor = 'Jones, Bob',
                                                [string]$Student = ''
                                )
                                [PSCustomObject]@{
                                                ETD        = $ETD
                                                Instructor = $Instructor
                                                SortieType = $SortieType
                                                Student    = $Student
                                                Mission    = ''
                                                Callsign   = ''
                                                Num        = ''
                                                Notes      = ''
                                }
                }

                # Students column of a tab, top to bottom
                function Get-TabStudents {
                                param ([string]$Path, [string]$Sheet)
                                return @(Import-Excel -Path $Path -WorksheetName $Sheet | ForEach-Object { $_.Student })
                }

                # Data rows (row 1 is the header) whose top border is Thick on every column
                function Get-ThickRows {
                                param ([string]$Path, [string]$Sheet)
                                $pkg = Open-ExcelPackage -Path $Path
                                try {
                                                $ws = $pkg.Workbook.Worksheets[$Sheet]
                                                $lastCol = $ws.Dimension.End.Column
                                                return @(for ($r = 2; $r -le $ws.Dimension.End.Row; $r++) {
                                                                $thick = @(for ($c = 1; $c -le $lastCol; $c++) { $ws.Cells[$r, $c].Style.Border.Top.Style -eq [OfficeOpenXml.Style.ExcelBorderStyle]::Thick })
                                                                if ($thick -notcontains $false) { $r }
                                                })
                                }
                                finally { Close-ExcelPackage -ExcelPackage $pkg -NoSave }
                }

                function Get-TabNames {
                                param ([string]$Path)
                                $pkg = Open-ExcelPackage -Path $Path
                                try { return @($pkg.Workbook.Worksheets | ForEach-Object { $_.Name }) }
                                finally { Close-ExcelPackage -ExcelPackage $pkg -NoSave }
                }
}

Describe 'Export-AtisWorkbook' -Skip:$script:NoExcel {
                BeforeEach {
                                $script:Path = Join-Path ([System.IO.Path]::GetTempPath()) ("FAIS_atis_" + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.xlsx')
                }
                AfterEach {
                                Remove-Item -LiteralPath $script:Path -Force -ErrorAction SilentlyContinue
                }

                It 'lists the flights in ETD order, then the sims in ETD order' {
                                $rows = @(
                                                New-AtisRow -ETD '2026-10-21 13:00:00' -SortieType 'OFT' -Instructor 'SIM IP' -Student 'Sim 1300'
                                                New-AtisRow -ETD '2026-10-21 09:00:00' -Student 'Flt 0900'
                                                New-AtisRow -ETD '2026-10-21 07:45:00' -SortieType 'IFT' -Instructor 'SIM IP' -Student 'Sim 0745'
                                                New-AtisRow -ETD '2026-10-21 07:30:00' -Student 'Flt 0730'
                                                New-AtisRow -ETD '2026-10-21 10:15:00' -SortieType 'O&B' -Student 'Flt 1015'
                                )
                                Export-AtisWorkbook -ScheduleData $rows -FilePath $script:Path 6>$null

                                Get-TabStudents $script:Path $script:Day1 | Should -Be @('Flt 0730', 'Flt 0900', 'Flt 1015', 'Sim 0745', 'Sim 1300')
                }

                It 'keeps the MasterSchedule order for lines at the same time' {
                                $rows = @(
                                                New-AtisRow -ETD '2026-10-21 07:30:00' -Instructor 'Lee, Ann' -Student 'Lead'
                                                New-AtisRow -ETD '2026-10-21 07:30:00' -Instructor 'Ortiz, Sam' -Student 'Wing'
                                                New-AtisRow -ETD '2026-10-21 07:30:00' -Instructor 'Adams, Al' -Student 'Third'
                                )
                                Export-AtisWorkbook -ScheduleData $rows -FilePath $script:Path 6>$null

                                Get-TabStudents $script:Path $script:Day1 | Should -Be @('Lead', 'Wing', 'Third')
                }

                It 'puts a thick border across every column between the flights and the sims, and nowhere else' {
                                $rows = @(
                                                New-AtisRow -ETD '2026-10-21 07:30:00' -Student 'F1'
                                                New-AtisRow -ETD '2026-10-21 09:00:00' -Student 'F2'
                                                New-AtisRow -ETD '2026-10-21 08:00:00' -SortieType 'OFT' -Student 'S1'
                                                New-AtisRow -ETD '2026-10-21 10:00:00' -SortieType 'IFT' -Student 'S2'
                                )
                                Export-AtisWorkbook -ScheduleData $rows -FilePath $script:Path 6>$null

                                # Header, F1, F2, then S1 on row 4
                                Get-ThickRows $script:Path $script:Day1 | Should -Be @(4)
                }

                It 'draws no border on a day with only flights or only sims' {
                                $rows = @(
                                                New-AtisRow -ETD '2026-10-21 07:30:00' -Student 'F1'
                                                New-AtisRow -ETD '2026-10-21 09:00:00' -Student 'F2'
                                                New-AtisRow -ETD '2026-10-22 08:00:00' -SortieType 'OFT' -Student 'S1'
                                                New-AtisRow -ETD '2026-10-22 10:00:00' -SortieType 'IFT' -Student 'S2'
                                )
                                Export-AtisWorkbook -ScheduleData $rows -FilePath $script:Path 6>$null

                                Get-ThickRows $script:Path $script:Day1 | Should -BeNullOrEmpty
                                Get-ThickRows $script:Path $script:Day2 | Should -BeNullOrEmpty
                }

                It 'writes one tab per day with only that day''s lines, in the MasterSchedule columns' {
                                $rows = @(
                                                New-AtisRow -ETD '2026-10-22 07:30:00' -Student 'Thu'
                                                New-AtisRow -ETD '2026-10-21 07:30:00' -Student 'Wed'
                                )
                                Export-AtisWorkbook -ScheduleData $rows -FilePath $script:Path 6>$null

                                Get-TabNames $script:Path | Should -Be @($script:Day1, $script:Day2)
                                Get-TabStudents $script:Path $script:Day1 | Should -Be @('Wed')
                                Get-TabStudents $script:Path $script:Day2 | Should -Be @('Thu')
                                $cols = @((Import-Excel -Path $script:Path -WorksheetName $script:Day1)[0].PSObject.Properties | ForEach-Object { $_.Name })
                                $cols | Should -Be @('ETD', 'Instructor', 'SortieType', 'Student', 'Mission', 'Callsign', 'Num', 'Notes')
                }

                It 'rebuilds a re-exported day (no old lines or border) and keeps the other days' {
                                $first = @(
                                                New-AtisRow -ETD '2026-10-21 07:30:00' -Student 'Old F'
                                                New-AtisRow -ETD '2026-10-21 08:00:00' -SortieType 'OFT' -Student 'Old S'
                                                New-AtisRow -ETD '2026-10-22 07:30:00' -Student 'Thu'
                                )
                                Export-AtisWorkbook -ScheduleData $first -FilePath $script:Path 6>$null
                                Export-AtisWorkbook -ScheduleData @(New-AtisRow -ETD '2026-10-21 09:00:00' -Student 'New F') -FilePath $script:Path 6>$null

                                Get-TabStudents $script:Path $script:Day1 | Should -Be @('New F')
                                Get-ThickRows $script:Path $script:Day1 | Should -BeNullOrEmpty
                                Get-TabStudents $script:Path $script:Day2 | Should -Be @('Thu')
                }

                It 'keeps the day tabs in date order, with other tabs left in front' {
                                @([PSCustomObject]@{ Note = 'mine' }) | Export-Excel -Path $script:Path -WorksheetName 'Notes'
                                Export-AtisWorkbook -ScheduleData @(New-AtisRow -ETD '2026-10-22 07:30:00' -Student 'Thu') -FilePath $script:Path 6>$null
                                Export-AtisWorkbook -ScheduleData @(New-AtisRow -ETD '2026-10-21 07:30:00' -Student 'Wed') -FilePath $script:Path 6>$null

                                Get-TabNames $script:Path | Should -Be @('Notes', $script:Day1, $script:Day2)
                }

                It 'opens on the first day just exported (and only that tab is selected)' {
                                Export-AtisWorkbook -ScheduleData @(New-AtisRow -ETD '2026-10-21 07:30:00' -Student 'Wed') -FilePath $script:Path 6>$null
                                Export-AtisWorkbook -ScheduleData @(New-AtisRow -ETD '2026-10-22 07:30:00' -Student 'Thu') -FilePath $script:Path 6>$null

                                $pkg = Open-ExcelPackage -Path $script:Path
                                try {
                                                $selected = @($pkg.Workbook.Worksheets | Where-Object { $_.View.TabSelected } | ForEach-Object { $_.Name })
                                                $active = $pkg.Workbook.Worksheets | Where-Object { $_.Index - 1 -eq $pkg.Workbook.View.ActiveTab } | ForEach-Object { $_.Name }
                                }
                                finally { Close-ExcelPackage -ExcelPackage $pkg -NoSave }
                                $selected | Should -Be @($script:Day2)
                                $active | Should -Be $script:Day2
                }

                It 'leaves lines with no ETD off, with a warning' {
                                $rows = @(
                                                New-AtisRow -ETD '2026-10-21 07:30:00' -Student 'Wed'
                                                New-AtisRow -ETD '' -Student 'Nowhere'
                                )
                                Export-AtisWorkbook -ScheduleData $rows -FilePath $script:Path -WarningVariable warned -WarningAction SilentlyContinue 6>$null

                                Get-TabStudents $script:Path $script:Day1 | Should -Be @('Wed')
                                "$warned" | Should -Match '1 line\(s\) have no ETD'
                }

                It 'writes no file for an empty schedule' {
                                Export-AtisWorkbook -ScheduleData @() -FilePath $script:Path 6>$null
                                Test-Path -LiteralPath $script:Path | Should -BeFalse
                }
}

Describe 'Export-FaisOutput writes ATIS' -Skip:$script:NoExcel {
                BeforeEach {
                                $script:Dir = Join-Path ([System.IO.Path]::GetTempPath()) ("FAIS_out_" + [guid]::NewGuid().ToString('N').Substring(0, 8))
                                New-Item -ItemType Directory -Path $script:Dir -Force | Out-Null
                                $script:Fill = [PSCustomObject]@{
                                                Schedule = @(
                                                                New-TestRow -ETD '2026-10-21 07:45' -SortieType 'IFT'
                                                                New-TestRow -ETD '2026-10-21 07:30' -Instructor 'Jones, Bob'
                                                )
                                                Students = @()
                                }
                }
                AfterEach {
                                Remove-Item -LiteralPath $script:Dir -Recurse -Force -ErrorAction SilentlyContinue
                }

                It 'writes ATIS.xlsx next to OutputSchedule when -AtisPath is given' {
                                $atis = Join-Path $script:Dir 'ATIS.xlsx'
                                Export-FaisOutput -FillResult $script:Fill -OutputPath (Join-Path $script:Dir 'OutputSchedule.xlsx') -AtisPath $atis -Settings @{ WriteStudentTabs = $false } 6>$null | Should -BeTrue

                                @(Import-Excel -Path $atis -WorksheetName $script:Day1 | ForEach-Object { $_.SortieType }) | Should -Be @('Day Local', 'IFT')
                }

                It 'skips ATIS when WriteAtis is off' {
                                $atis = Join-Path $script:Dir 'ATIS.xlsx'
                                Export-FaisOutput -FillResult $script:Fill -OutputPath (Join-Path $script:Dir 'OutputSchedule.xlsx') -AtisPath $atis -Settings @{ WriteStudentTabs = $false; WriteAtis = $false } 6>$null | Should -BeTrue

                                Test-Path -LiteralPath $atis | Should -BeFalse
                }
}
