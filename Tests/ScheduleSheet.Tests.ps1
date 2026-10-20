# Pester 5 tests for reading MasterData: the Schedule sheet (Read-ScheduleSheet and friends),
# the time / date cell parser, and IngestData end to end. Workbooks are built in TestDrive.

BeforeAll {
    . "$PSScriptRoot/Fixtures/New-FaisTestData.ps1"
    Import-FaisModules
    Import-Module ImportExcel -ErrorAction Stop 3>$null

    # Writes a worksheet from an array of rows (each row an array of cell values).
    # A cell given as @{ V = value; F = 'number format' } gets that format.
    function New-TBook {
        param ([string]$Name, [hashtable]$Sheets)
        $path = Join-Path $TestDrive $Name
        if (Test-Path $path) { Remove-Item $path -Force }
        $pkg = Open-ExcelPackage -Path $path -Create
        foreach ($sheet in $Sheets.Keys) {
            $ws = Add-Worksheet -ExcelPackage $pkg -WorksheetName $sheet
            $r = 1
            foreach ($row in $Sheets[$sheet]) {
                $c = 1
                foreach ($cell in @($row)) {
                    if ($cell -is [hashtable]) {
                        $ws.Cells[$r, $c].Value = $cell.V
                        $ws.Cells[$r, $c].Style.Numberformat.Format = $cell.F
                    }
                    elseif ($null -ne $cell) { $ws.Cells[$r, $c].Value = $cell }
                    $c++
                }
                $r++
            }
        }
        Close-ExcelPackage $pkg
        return $path
    }

    function T { param ([string]$At) return @{ V = ([datetime]$At).ToOADate(); F = 'm/d/yyyy h:mm' } }
}

Describe 'ConvertTo-FaisTimeValue' {
    It '<In> reads as <Want> (has date: <HasDate>)' -ForEach @(
        @{ In = '0730'; Want = '07:30'; HasDate = $false }
        @{ In = '730'; Want = '07:30'; HasDate = $false }
        @{ In = '7:30'; Want = '07:30'; HasDate = $false }
        @{ In = '07:30:00'; Want = '07:30'; HasDate = $false }
        @{ In = '1:15 PM'; Want = '13:15'; HasDate = $false }
        @{ In = '12:05 am'; Want = '00:05'; HasDate = $false }
        @{ In = '0730L'; Want = '07:30'; HasDate = $false }
        @{ In = '10/21/2026 0730'; Want = '2026-10-21 07:30'; HasDate = $true }
        @{ In = '2026-10-21 07:30'; Want = '2026-10-21 07:30'; HasDate = $true }
        @{ In = 'Tuesday, October 21, 2026'; Want = '2026-10-21 00:00'; HasDate = $true }
        @{ In = 'Wednesday. October 21, 2026'; Want = '2026-10-21 00:00'; HasDate = $true }
        @{ In = 'Tue 21 Oct 2026 1315'; Want = '2026-10-21 13:15'; HasDate = $true }
        @{ In = '46316.3125'; Want = '2026-10-21 07:30'; HasDate = $true }
    ) {
        $r = ConvertTo-FaisTimeValue $In
        $null -ne $r | Should -BeTrue
        $r.HasDate | Should -Be $HasDate
        $fmt = if ($HasDate) { 'yyyy-MM-dd HH:mm' } else { 'HH:mm' }
        $r.Value.ToString($fmt) | Should -BeExactly $Want
    }

    It 'numbers: Excel serials, fractions of a day and HHMM' {
        $s = ConvertTo-FaisTimeValue ([double]46316.3125)
        $s.HasDate | Should -BeTrue
        $s.Value | Should -Be ([datetime]'2026-10-21 07:30')
        $f = ConvertTo-FaisTimeValue ([double]0.3125)
        $f.HasDate | Should -BeFalse
        $f.Value.ToString('HH:mm') | Should -BeExactly '07:30'
        $h = ConvertTo-FaisTimeValue 1315
        $h.HasDate | Should -BeFalse
        $h.Value.ToString('HH:mm') | Should -BeExactly '13:15'
    }

    It 'a time-only date cell (1899-12-30 07:30) is a time with no date' {
        $r = ConvertTo-FaisTimeValue ([datetime]'1899-12-30 07:30')
        $r.HasDate | Should -BeFalse
        $r.Value.ToString('HH:mm') | Should -BeExactly '07:30'
    }

    It 'a real date stays a date' {
        $r = ConvertTo-FaisTimeValue ([datetime]'2026-10-21 09:45')
        $r.HasDate | Should -BeTrue
        $r.Value | Should -Be ([datetime]'2026-10-21 09:45')
    }

    It 'rejects <In>' -ForEach @(
        @{ In = '2500' }, @{ In = '0775' }, @{ In = 'DL' }, @{ In = 'Smith' }, @{ In = '' }, @{ In = '13:15 PM' }
    ) {
        ConvertTo-FaisTimeValue $In | Should -BeNullOrEmpty
    }

    It 'rejects $null and numbers that are neither serials nor times' {
        ConvertTo-FaisTimeValue $null | Should -BeNullOrEmpty
        ConvertTo-FaisTimeValue ([double]5000) | Should -BeNullOrEmpty
        ConvertTo-FaisTimeValue 1290 | Should -BeNullOrEmpty
    }
}

Describe 'ConvertTo-CanonicalSortieType' {
    It '<In> -> <Want>' -ForEach @(
        @{ In = 'DL'; Want = 'Day Local' }, @{ In = 'day local'; Want = 'Day Local' }, @{ In = 'Local'; Want = 'Day Local' }
        @{ In = 'OB'; Want = 'O&B' }, @{ In = 'O&B'; Want = 'O&B' }, @{ In = 'o & b'; Want = 'O&B' }
        @{ In = 'RTN'; Want = 'O&B Return' }, @{ In = 'O&B Return'; Want = 'O&B Return' }
        @{ In = 'oft'; Want = 'OFT' }, @{ In = 'IFT'; Want = 'IFT' }
        @{ In = 'Night Local'; Want = 'Night Local' }, @{ In = ' OFT 2 '; Want = 'OFT 2' }
    ) {
        ConvertTo-CanonicalSortieType $In | Should -BeExactly $Want
    }
}

Describe 'Read-ScheduleSheet' {
    Context 'legacy layout, no headings' {
        BeforeAll {
            $script:P = New-TBook 'legacy.xlsx' @{ Schedule = @(
                    , @((T '2026-10-21 07:30'), 'Jones, Bob', 'Day Local', 'Adams, Al', 'FT5103', 'TALON', 31, 'early')
                    , @((T '2026-10-21 10:15'), 'Ortiz, Sam', 'O&B')
                    , @()
                    , @((T '2026-10-21 13:15'), 'Ortiz, Sam', 'O&B Return')
                    , @((T '2026-10-21 13:00'), $null, 'OFT')
                ) }
            $script:R = Read-ScheduleSheet -Path $script:P
        }

        It 'finds the sheet and uses the fixed columns' {
            $script:R.Found | Should -BeTrue
            $script:R.Error | Should -BeExactly ''
            $script:R.Layout | Should -BeLike 'Fixed columns*'
        }

        It 'keeps reading past a blank row (it used to stop there)' {
            @($script:R.Rows).Count | Should -Be 4
            @($script:R.Rows | ForEach-Object { $_.SheetRow }) -join ',' | Should -BeExactly '1,2,4,5'
        }

        It 'reads every column of a line' {
            $r = $script:R.Rows[0]
            $r.ETD | Should -BeExactly '2026-10-21 07:30:00'
            $r.Day | Should -BeExactly '2026-10-21'
            $r.Time | Should -BeExactly '07:30'
            $r.TimeOnly | Should -BeFalse
            $r.Instructor | Should -BeExactly 'Jones, Bob'
            $r.SortieType | Should -BeExactly 'Day Local'
            $r.Student | Should -BeExactly 'Adams, Al'
            $r.Mission | Should -BeExactly 'FT5103'
            $r.Callsign | Should -BeExactly 'TALON'
            $r.Num | Should -BeExactly '31'
            $r.Notes | Should -BeExactly 'early'
        }

        It 'blank cells are $null and a sim line with no IP gets SIM IP' {
            $r = $script:R.Rows[1]
            $r.Student | Should -BeNullOrEmpty
            $r.Mission | Should -BeNullOrEmpty
            $script:R.Rows[3].Instructor | Should -BeExactly 'SIM IP'
        }
    }

    Context 'headings in any order, a Date column and short sortie types' {
        BeforeAll {
            $script:P = New-TBook 'headings.xlsx' @{ Schedule = @(
                    , @('Schedule for the week')
                    , @('Date', 'Type', 'Time', 'IP', 'SP', 'Event', 'Remarks')
                    , @('10/21/2026', 'DL', '0730', 'Jones, Bob', 'Adams, Al', 'FT5103', 'brief 0630')
                    , @('10/21/2026', 'OB', @{ V = 0.4270833333; F = 'hh:mm' }, 'Lee, Ann')
                    , @('10/22/2026', 'RTN', '1:15 PM', 'Lee, Ann')
                    , @('Date', 'Type', 'Time', 'IP', 'SP', 'Event', 'Remarks')
                    , @('10/22/2026', 'IFT', '0800')
                ) }
            $script:R = Read-ScheduleSheet -Path $script:P
        }

        It 'finds the heading row below a title' {
            $script:R.HeaderRow | Should -Be 2
            $script:R.Layout | Should -BeExactly 'Headings on row 2'
        }

        It 'maps the columns by heading and skips the repeated heading row' {
            @($script:R.Rows).Count | Should -Be 4
            $r = $script:R.Rows[0]
            $r.ETD | Should -BeExactly '2026-10-21 07:30:00'
            $r.Instructor | Should -BeExactly 'Jones, Bob'
            $r.Student | Should -BeExactly 'Adams, Al'
            $r.Mission | Should -BeExactly 'FT5103'
            $r.Notes | Should -BeExactly 'brief 0630'
        }

        It 'turns DL / OB / RTN into the board''s sortie types' {
            (@($script:R.Rows | ForEach-Object { $_.SortieType }) -join ',') | Should -BeExactly 'Day Local,O&B,O&B Return,IFT'
        }

        It 'dates time-only cells from the Date column' {
            $script:R.Rows[1].ETD | Should -BeExactly '2026-10-21 10:15:00'
            $script:R.Rows[2].ETD | Should -BeExactly '2026-10-22 13:15:00'
            $script:R.Rows[3].Day | Should -BeExactly '2026-10-22'
        }
    }

    Context 'date title rows and times with no date' {
        BeforeAll {
            $script:P = New-TBook 'titles.xlsx' @{ Schedule = @(
                    , @('ETD', 'Instructor', 'SortieType', 'Student')
                    , @('0645', 'Jones, Bob', 'DL')
                    , @('Wednesday, October 21, 2026')
                    , @('0730', 'Jones, Bob', 'DL')
                    , @(@{ V = 0.5; F = 'h:mm AM/PM' }, 'Lee, Ann', 'Day Local')
                    , @('Thursday. October 22, 2026')
                    , @('0800', 'Ortiz, Sam', 'DL')
                ) }
            $script:R = Read-ScheduleSheet -Path $script:P
        }

        It 'a row above any date title has no day (it goes to the day it is loaded onto)' {
            $r = $script:R.Rows[0]
            $r.TimeOnly | Should -BeTrue
            $r.Day | Should -BeExactly ''
            $r.ETD | Should -BeNullOrEmpty
            $r.Time | Should -BeExactly '06:45'
        }

        It 'a date title row dates the rows under it' {
            $script:R.Rows[1].ETD | Should -BeExactly '2026-10-21 07:30:00'
            $script:R.Rows[2].ETD | Should -BeExactly '2026-10-21 12:00:00'
            $script:R.Rows[3].ETD | Should -BeExactly '2026-10-22 08:00:00'
            @($script:R.Rows).Count | Should -Be 4
        }
    }

    Context 'a date title in a column that is not the ETD column' {
        BeforeAll {
            $script:P = New-TBook 'titlecol.xlsx' @{ Schedule = @(
                    , @('Type', 'Time', 'IP')
                    , @('Wednesday, October 21, 2026')
                    , @('DL', '0730', 'Jones, Bob')
                    , @()
                    , @('OFT', '1300')
                ) }
            $script:R = Read-ScheduleSheet -Path $script:P
        }

        It 'still dates the rows under it' {
            @($script:R.Skipped).Count | Should -Be 0
            (@($script:R.Rows | ForEach-Object { $_.ETD }) -join ',') | Should -BeExactly '2026-10-21 07:30:00,2026-10-21 13:00:00'
        }
    }

    Context 'rows that cannot be read' {
        BeforeAll {
            $script:P = New-TBook 'bad.xlsx' @{ Schedule = @(
                    , @('ETD', 'Instructor', 'SortieType')
                    , @('2500', 'Jones, Bob', 'DL')
                    , @('0730', 'Jones, Bob')
                    , @($null, 'Jones, Bob', 'DL')
                    , @('0800', 'Lee, Ann', 'DL')
                ) }
            $script:R = Read-ScheduleSheet -Path $script:P
        }

        It 'reads the good row and says why each other row was skipped' {
            @($script:R.Rows).Count | Should -Be 1
            $script:R.Rows[0].Instructor | Should -BeExactly 'Lee, Ann'
            $s = @($script:R.Skipped)
            $s.Count | Should -Be 3
            $s[0] | Should -BeLike "row 2: ETD '2500' is not a time*"
            $s[1] | Should -BeExactly 'row 3: no sortie type'
            $s[2] | Should -BeExactly 'row 4: no ETD (Day Local line)'
        }
    }

    It 'a workbook with no Schedule sheet is fine (Found is false, no error)' {
        $p = New-TBook 'nosheet.xlsx' @{ Students = @(, @('SPName')) }
        $r = Read-ScheduleSheet -Path $p
        $r.Found | Should -BeFalse
        $r.Error | Should -BeExactly ''
        @($r.Rows).Count | Should -Be 0
    }

    It 'an empty Schedule sheet has no rows and no error' {
        $p = New-TBook 'empty.xlsx' @{ Schedule = @() }
        $r = Read-ScheduleSheet -Path $p
        $r.Found | Should -BeTrue
        $r.Error | Should -BeExactly ''
        @($r.Rows).Count | Should -Be 0
    }

    It 'a missing file is an error' {
        (Read-ScheduleSheet -Path (Join-Path $TestDrive 'none.xlsx')).Error | Should -BeLike '*was not found*'
    }

    It 'headings with no sortie type column are an error that says so' {
        $p = New-TBook 'notype.xlsx' @{ Schedule = @(@('ETD', 'Instructor', 'Student'), @('0730', 'Jones, Bob', 'Adams, Al')) }
        (Read-ScheduleSheet -Path $p).Error | Should -BeLike '*no SortieType column*'
    }

    It 'reads a workbook another program holds open' {
        $p = New-TBook 'held.xlsx' @{ Schedule = @(, @((T '2026-10-21 07:30'), 'Jones, Bob', 'DL')) }
        $fs = [System.IO.File]::Open($p, 'Open', 'ReadWrite', 'Read')
        try { @((Read-ScheduleSheet -Path $p).Rows).Count | Should -Be 1 }
        finally { $fs.Close() }
    }
}

Describe 'Get-ScheduleSheetDays / Get-ScheduleDayFingerprint' {
    BeforeAll {
        function New-SR {
            param ([string]$Day, [string]$Time, [string]$Type = 'Day Local', [string]$IP = 'Jones, Bob', [string]$Student = $null)
            [PSCustomObject]@{ ETD = $(if ($Day) { "$Day $($Time):00" } else { $null }); Day = $Day; Time = $Time; TimeOnly = (-not $Day)
                Instructor = $IP; SortieType = $Type; Student = $Student; Mission = $null; Callsign = $null; Num = $null; Notes = $null; SheetRow = 1 }
        }
        $script:Rows = @(
            New-SR '2026-10-22' '07:30'
            New-SR '2026-10-21' '07:30' -Student 'Adams, Al'
            New-SR '2026-10-21' '08:00'
            New-SR '' '09:00'
        )
    }

    It 'lists dated days in order with counts, and the no-date group last' {
        $d = @(Get-ScheduleSheetDays -Rows $script:Rows)
        (@($d | ForEach-Object { "$($_.Day)=$($_.Count)/$($_.Students)" }) -join ',') | Should -BeExactly '2026-10-21=2/1,2026-10-22=1/0,=1/0'
        $d[0].Label | Should -BeExactly 'Wed 21 Oct 2026'
        $d[2].Label | Should -BeExactly 'No date'
    }

    It 'the fingerprint ignores row order but sees any change' {
        $a = Get-ScheduleDayFingerprint -Rows $script:Rows -Day '2026-10-21'
        $b = Get-ScheduleDayFingerprint -Rows @($script:Rows[2], $script:Rows[1]) -Day '2026-10-21'
        $a | Should -Not -BeNullOrEmpty
        $a | Should -BeExactly $b
        $changed = @($script:Rows | ForEach-Object { $_.PSObject.Copy() })
        $changed[2].Instructor = 'Lee, Ann'
        Get-ScheduleDayFingerprint -Rows $changed -Day '2026-10-21' | Should -Not -BeExactly $a
    }

    It 'a day with no lines has an empty fingerprint; -IncludeTimeOnly adds the no-date rows' {
        Get-ScheduleDayFingerprint -Rows $script:Rows -Day '2026-10-30' | Should -BeExactly ''
        Get-ScheduleDayFingerprint -Rows $script:Rows -Day '2026-10-30' -IncludeTimeOnly | Should -Not -BeExactly ''
    }
}

Describe 'IngestData' {
    BeforeAll {
        $script:Master = New-TBook 'MasterData.xlsx' @{
            Instructors     = @(@('IPName', 'IsCounter', 'IsExperienced', 'IsCheck', 'Callsign', 'Num'), @('Jones, Bob', $null, 'TRUE', 'FALSE', 'TALON', 31))
            Students        = @(@('SPKey', 'SPName', 'IsCap', 'SimDebt', 'Priority'), @(1, 'Adams, Al', 'FALSE', 1.5, 2), @(2, 'Brown, Bea', 'TRUE', $null, $null))
            Syllabus        = @(@('SyllabusID', 'Mission', 'Prereqs', 'Type', 'IsCheck', 'IsComplete', 'Level'), @(1, 'FI3101', '0', 'IFT', 'FALSE', 'FALSE', 1), @(2, 'FT5001', 'FI3101', 'Single', 'FALSE', 'FALSE', 1))
            SoftWeights     = @(@('ContString', 'TransSimRatio', 'FormSimRatio', 'SimStopMult', 'FormLookAhead'), @(10, 0.5, 0.5, 2, 22))
            FormCallsigns   = @(@('Callsign', 'Num'), @('HAZE', 11), @('HAZE', 12))
            StudentAbsences = @(@('Student', 'StartTime', 'EndTime', 'Reason', 'IsActive'), @('Adams, Al', (T '2026-10-21 08:00'), (T '2026-10-21 12:00'), 'Dentist', 'TRUE'))
            Schedule        = @(@('ETD', 'Instructor', 'SortieType'), @((T '2026-10-21 07:30'), 'Jones, Bob', 'DL'), @(), @('0900', 'Jones, Bob', 'DL'))
        }
        $script:Output = New-TBook 'OutputSchedule.xlsx' @{
            MasterSchedule = @(@('ETD', 'Instructor'))
            'Adams, Al'    = @(@('SyllabusID', 'Mission', 'Prereqs', 'Instructor', 'ETD', 'Type', 'IsCheck', 'IsComplete', 'OptingEvents', 'Level'),
                @(1, 'FI3101', '0', 'SIM IP', '2026-10-15 08:00:00', 'IFT', 'FALSE', 'TRUE', 0, 1), @(2, 'FT5001', 'FI3101', '', '', 'Single', 'FALSE', 'FALSE', 0, 1))
        }
        $script:I = IngestData -excelInputPath $script:Master -excelOutputPath $script:Output 6>$null
    }

    It 'reads the Schedule sheet into both Schedule (dated rows) and ScheduleSheet (every row)' {
        @($script:I.Schedule).Count | Should -Be 1
        $script:I.Schedule[0].ETD | Should -BeExactly '2026-10-21 07:30:00'
        @($script:I.ScheduleSheet.Rows).Count | Should -Be 2
        $script:I.ScheduleSheet.Rows[1].TimeOnly | Should -BeTrue
    }

    It 'reads a student''s own tab from OutputSchedule and the master Syllabus for a student with no tab' {
        $a = $script:I.Students | Where-Object { $_.Name -eq 'Adams, Al' }
        $b = $script:I.Students | Where-Object { $_.Name -eq 'Brown, Bea' }
        (@($a.Syllabus | Where-Object { $_.IsComplete }).Mission -join ',') | Should -BeExactly 'FI3101'
        @($b.Syllabus | Where-Object { $_.IsComplete }).Count | Should -Be 0
        @($b.Syllabus).Count | Should -Be 2
    }

    It 'reports no OutputSchedule error when it opened fine' {
        $script:I.OutputError | Should -BeExactly ''
    }

    It 'reports an OutputSchedule that exists but cannot be opened (so FAIS stops instead of starting everyone over)' {
        $bad = Join-Path $TestDrive 'Broken.xlsx'
        [System.IO.File]::WriteAllText($bad, 'not a workbook')
        $i = IngestData -excelInputPath $script:Master -excelOutputPath $bad 6>$null
        $i.OutputError | Should -Not -BeNullOrEmpty
    }

    It 'reads instructors, flags, sim debt, weights, formations and absences' {
        $script:I.Instructors[0].IsCounter | Should -BeTrue
        ($script:I.Students | Where-Object { $_.Name -eq 'Adams, Al' }).SimDebt | Should -Be 1.5
        ($script:I.Students | Where-Object { $_.Name -eq 'Brown, Bea' }).IsCap | Should -BeTrue
        $script:I.SoftWeights.TransSimRatio | Should -Be 1
        @($script:I.Formations).Count | Should -Be 2
        @(($script:I.Students | Where-Object { $_.Name -eq 'Adams, Al' }).Absences).Count | Should -Be 1
    }
}
