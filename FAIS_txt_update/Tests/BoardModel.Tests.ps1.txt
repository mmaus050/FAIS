# Pester 5 tests for the puck board's data model (Modules\BoardModel.psm1).
# Pure PowerShell: no WPF, no Excel. Files go to TestDrive.
# Dates: Tue 2026-10-20 (D1) and Wed 2026-10-21 (D2) are the board days.
# Copy-previous-day uses Fri 2026-10-23 -> Mon 2026-10-26 (across the weekend).

BeforeAll {
    . "$PSScriptRoot/Fixtures/New-FaisTestData.ps1"
    Import-FaisModules

    $script:IPs = New-TestInstructors
    $script:Stu = @(
        New-TestStudent -Name 'Adams, Amy'
        New-TestStudent -Name 'Baker, Ben'
        New-TestStudent -Name 'Cruz, Cal'
    )
    $script:D1 = '2026-10-20'
    $script:D2 = '2026-10-21'
    $script:RowProps = 'ETD,Instructor,SortieType,Student,Mission,Callsign,Num,Notes,BoardLineId,LinkedLineId,WingOfLineId,FilledBy'

    function New-TB {
        param ([string[]]$Days = @('2026-10-20'), [hashtable]$Settings)
        if ($null -ne $Settings) { return (New-Board -Days $Days -Settings $Settings) }
        return (New-Board -Days $Days)
    }

    function New-TL {
        param (
            [string]$At,
            [string]$Type = 'Day Local',
            [string]$IP = '',
            [string]$Student = '',
            [string]$Mission = '',
            [string]$Notes = '',
            [string]$Callsign = '',
            [string]$Num = ''
        )
        return (New-BoardLine -ETD ([datetime]$At) -SortieType $Type -Instructor $IP -Student $Student -Mission $Mission -Notes $Notes -Callsign $Callsign -Num $Num)
    }

    # Puts lines on a board without an undo point
    function Add-TL {
        param ($Board, [array]$Lines)
        Add-BoardLine -Board $Board -Line $Lines -NoUndo
    }

    # O&B plus its linked return (default +3 hr); returns {OB; Ret}
    function Add-TOB {
        param ($Board, [string]$At, [string]$IP = '', [string]$ReturnAt = '')
        if ($ReturnAt -eq '') { $ReturnAt = ([datetime]$At).AddMinutes(180).ToString('yyyy-MM-dd HH:mm') }
        $ob = New-TL -At $At -Type 'O&B' -IP $IP
        $ret = New-TL -At $ReturnAt -Type 'O&B Return' -IP $IP
        $ob.LinkedLineId = $ret.Id
        $ret.LinkedLineId = $ob.Id
        Add-BoardLine -Board $Board -Line @($ob, $ret) -NoUndo
        return [PSCustomObject]@{ OB = $ob; Ret = $ret }
    }

    function Get-TSnap {
        param ($Board)
        return (ConvertTo-BoardSnapshot $Board)
    }

    function Get-TQ {
        param ([string]$Text, $Board, $Settings, $Day = '2026-10-20')
        $splat = @{ Text = $Text; Day = $Day; Instructors = $script:IPs; Students = $script:Stu }
        if ($null -ne $Board) { $splat['Board'] = $Board }
        if ($null -ne $Settings) { $splat['Settings'] = $Settings }
        return (ConvertFrom-QuickEntry @splat 3>$null)
    }
}


Describe 'Board basics' {
    It 'New-Board normalizes, sorts and de-duplicates days and starts at version 0' {
        $b = New-Board -Days @('2026-10-21', '2026-10-20 13:00', '2026-10-21')
        ($b.Days.ToArray() -join ',') | Should -BeExactly '2026-10-20,2026-10-21'
        $b.Version | Should -Be 0
        $b.Undo.Count | Should -Be 0
        $b.Lines.Count | Should -Be 0
        $b.SchemaVersion | Should -Be 1
    }

    It 'New-Board takes the default settings and lets -Settings override them' {
        $b = New-Board -Days @('2026-10-20')
        $b.Settings.OBReturnOffsetMinutes | Should -Be 180
        $b.Settings.ReturnSortieType | Should -BeExactly 'O&B Return'
        $b.Settings.SimInstructor | Should -BeExactly 'SIM IP'
        $b.Settings.UndoDepth | Should -Be 50
        $b2 = New-Board -Days @('2026-10-20') -Settings @{ OBReturnOffsetMinutes = 120 }
        $b2.Settings.OBReturnOffsetMinutes | Should -Be 120
        $b2.Settings.ReturnSortieType | Should -BeExactly 'O&B Return'
    }

    It 'New-BoardLine sets Day from the ETD and gives an 8-character id' {
        $l = New-TL -At '2026-10-20 07:30' -IP 'Jones, Bob'
        $l.Day | Should -BeExactly '2026-10-20'
        $l.ETD | Should -Be ([datetime]'2026-10-20 07:30')
        $l.Id | Should -Match '^[0-9a-f]{8}$'
        $l.Source | Should -BeExactly 'Manual'
        $l.LinkedLineId | Should -BeExactly ''
        $l.WingOfLineId | Should -BeExactly ''
    }

    It 'Get-BoardDayLines sorts by ETD and keeps the added order for ties' {
        $b = New-TB
        $a = New-TL -At '2026-10-20 13:00' -Type 'OFT'
        $c = New-TL -At '2026-10-20 07:30'
        $d = New-TL -At '2026-10-20 13:00' -Type 'IFT'
        $other = New-TL -At '2026-10-21 06:00'
        Add-TL $b @($a, $c, $d, $other)
        $ids = @(Get-BoardDayLines -Board $b -Day '2026-10-20' | ForEach-Object { $_.Id })
        ($ids -join ',') | Should -BeExactly (@($c.Id, $a.Id, $d.Id) -join ',')
        @(Get-BoardDayLines -Board $b -Day '').Count | Should -Be 4
    }

    It 'Add-BoardDay / Remove-BoardDay change the days with one undo step each' {
        $b = New-TB
        Add-TL $b @(New-TL -At '2026-10-21 08:00')
        (Add-BoardDay -Board $b -Day '2026-10-19') | Should -BeExactly '2026-10-19'
        ($b.Days.ToArray() -join ',') | Should -BeExactly '2026-10-19,2026-10-20,2026-10-21'
        Remove-BoardDay -Board $b -Day '2026-10-21'
        ($b.Days.ToArray() -join ',') | Should -BeExactly '2026-10-19,2026-10-20'
        $b.Lines.Count | Should -Be 0
        (Undo-Board $b) | Should -BeTrue
        $b.Lines.Count | Should -Be 1
        $b.Days.Contains('2026-10-21') | Should -BeTrue
    }

    It 'Set-BoardMission and Set-BoardLineField change one field and skip no-op edits' {
        $b = New-TB
        $l = New-TL -At '2026-10-20 08:00' -Student 'Adams, Amy'
        Add-TL $b @($l)
        $v = $b.Version
        Set-BoardMission -Board $b -Id $l.Id -Mission 'FT5103'
        $l.Mission | Should -BeExactly 'FT5103'
        Set-BoardMission -Board $b -Id $l.Id -Mission 'FT5103'
        Set-BoardLineField -Board $b -Id $l.Id -Field 'Callsign' -Value 'HAZE'
        Set-BoardLineField -Board $b -Id $l.Id -Field 'Num' -Value '11'
        Set-BoardLineField -Board $b -Id $l.Id -Field 'Notes' -Value 'Wx brief'
        Set-BoardLineField -Board $b -Id $l.Id -Field 'Notes' -Value 'Wx brief'
        $l.Callsign | Should -BeExactly 'HAZE'
        $l.Num | Should -BeExactly '11'
        $l.Notes | Should -BeExactly 'Wx brief'
        $b.Version | Should -Be ($v + 4)
        $b.Undo.Count | Should -Be 4
    }

    It 'Unlink-ReturnLine clears the link on both lines' {
        $b = New-TB
        $p = Add-TOB $b '2026-10-20 10:15' 'Jones, Bob'
        Unlink-ReturnLine -Board $b -Id $p.Ret.Id
        $p.OB.LinkedLineId | Should -BeExactly ''
        $p.Ret.LinkedLineId | Should -BeExactly ''
        $b.Lines.Count | Should -Be 2
    }
}


Describe 'ConvertTo-BoardTime / Format-BoardTime' {
    It 'reads <T> as <H>:<M>' -TestCases @(
        @{ T = '0730'; H = 7; M = 30 }
        @{ T = '730'; H = 7; M = 30 }
        @{ T = '7:30'; H = 7; M = 30 }
        @{ T = '07:30'; H = 7; M = 30 }
        @{ T = '07.30'; H = 7; M = 30 }
        @{ T = '0730L'; H = 7; M = 30 }
        @{ T = '0730l'; H = 7; M = 30 }
        @{ T = '1015Z'; H = 10; M = 15 }
        @{ T = '7'; H = 7; M = 0 }
        @{ T = '13'; H = 13; M = 0 }
        @{ T = ' 1300 '; H = 13; M = 0 }
        @{ T = '0000'; H = 0; M = 0 }
        @{ T = '2359'; H = 23; M = 59 }
    ) {
        $ts = ConvertTo-BoardTime $T
        $ts | Should -BeOfType [timespan]
        $ts.Hours | Should -Be $H
        $ts.Minutes | Should -Be $M
        $ts.Days | Should -Be 0
    }

    It 'rejects <T>' -TestCases @(
        @{ T = '2500' }
        @{ T = '0760' }
        @{ T = '24:00' }
        @{ T = '7:60' }
        @{ T = 'abc' }
        @{ T = '' }
        @{ T = '   ' }
        @{ T = '12345' }
        @{ T = '7:5' }
        @{ T = '-730' }
        @{ T = '07:30pm' }
    ) {
        $null -eq (ConvertTo-BoardTime $T) | Should -BeTrue
    }

    It 'rejects $null' {
        $null -eq (ConvertTo-BoardTime $null) | Should -BeTrue
    }

    It 'formats a timespan as HHmm' {
        Format-BoardTime (New-TimeSpan -Hours 7 -Minutes 30) | Should -BeExactly '0730'
        Format-BoardTime (New-TimeSpan -Hours 0) | Should -BeExactly '0000'
        Format-BoardTime (ConvertTo-BoardTime '7:05') | Should -BeExactly '0705'
    }

    It 'formats a datetime or a date string as HHmm' {
        Format-BoardTime ([datetime]'2026-10-20 13:05') | Should -BeExactly '1305'
        Format-BoardTime '2026-10-20 06:45:00' | Should -BeExactly '0645'
    }

    It 'returns an empty string for nothing or junk' {
        Format-BoardTime $null | Should -BeExactly ''
        Format-BoardTime 'not a time' | Should -BeExactly ''
    }
}


Describe 'ConvertTo-BoardSortieType' {
    It '<Token> -> <Want>' -TestCases @(
        @{ Token = 'DL'; Want = 'Day Local' }
        @{ Token = 'dl'; Want = 'Day Local' }
        @{ Token = 'Local'; Want = 'Day Local' }
        @{ Token = 'DayLocal'; Want = 'Day Local' }
        @{ Token = 'Day Local'; Want = 'Day Local' }
        @{ Token = 'day-local'; Want = 'Day Local' }
        @{ Token = 'D'; Want = 'Day Local' }
        @{ Token = 'L'; Want = 'Day Local' }
        @{ Token = 'OB'; Want = 'O&B' }
        @{ Token = 'ob'; Want = 'O&B' }
        @{ Token = 'O&B'; Want = 'O&B' }
        @{ Token = 'o&b'; Want = 'O&B' }
        @{ Token = 'ONB'; Want = 'O&B' }
        @{ Token = 'OANDB'; Want = 'O&B' }
        @{ Token = 'O/B'; Want = 'O&B' }
        @{ Token = 'RTN'; Want = 'O&B Return' }
        @{ Token = 'rtn'; Want = 'O&B Return' }
        @{ Token = 'RET'; Want = 'O&B Return' }
        @{ Token = 'Return'; Want = 'O&B Return' }
        @{ Token = 'R'; Want = 'O&B Return' }
        @{ Token = 'O&B Return'; Want = 'O&B Return' }
        @{ Token = 'OB RTN'; Want = 'O&B Return' }
        @{ Token = 'O&BRTN'; Want = 'O&B Return' }
        @{ Token = 'OFT'; Want = 'OFT' }
        @{ Token = 'oft'; Want = 'OFT' }
        @{ Token = 'IFT'; Want = 'IFT' }
        @{ Token = 'ift'; Want = 'IFT' }
        @{ Token = 'SIM'; Want = 'AMBIGUOUS' }
        @{ Token = 'sim'; Want = 'AMBIGUOUS' }
        @{ Token = 'S'; Want = 'AMBIGUOUS' }
    ) {
        ConvertTo-BoardSortieType $Token | Should -BeExactly $Want
    }

    It 'returns $null for <Token>' -TestCases @(
        @{ Token = 'Jones' }
        @{ Token = 'XC' }
        @{ Token = '' }
        @{ Token = '0730' }
    ) {
        $null -eq (ConvertTo-BoardSortieType $Token) | Should -BeTrue
    }

    It 'uses the settings return type name' {
        ConvertTo-BoardSortieType 'RTN' @{ ReturnSortieType = 'Return Leg' } | Should -BeExactly 'Return Leg'
        ConvertTo-BoardSortieType 'RTN' @{} | Should -BeExactly 'O&B Return'
    }
}


Describe 'ConvertFrom-QuickEntry' {
    Context 'Day Local and sim lines' {
        It '<Text> gives one Day Local line at 0730' -TestCases @(
            @{ Text = '0730 DL' }
            @{ Text = '730 dl' }
            @{ Text = '7:30 local' }
            @{ Text = '0730 Day Local' }
            @{ Text = '07.30 day local' }
            @{ Text = '0730L DL' }
        ) {
            $r = Get-TQ $Text
            $r.Errors.Count | Should -Be 0
            $r.Warnings.Count | Should -Be 0
            $r.Lines.Count | Should -Be 1
            $l = $r.Lines[0]
            $l.ETD | Should -Be ([datetime]'2026-10-20 07:30')
            $l.Day | Should -BeExactly '2026-10-20'
            $l.SortieType | Should -BeExactly 'Day Local'
            $l.Instructor | Should -BeExactly ''
            $l.Student | Should -BeExactly ''
            $l.Mission | Should -BeExactly ''
            $l.LinkedLineId | Should -BeExactly ''
            $l.Source | Should -BeExactly 'Manual'
        }

        It 'accepts the day as a [datetime]' {
            $r = Get-TQ '0730 DL' -Day ([datetime]'2026-10-21')
            $r.Lines[0].ETD | Should -Be ([datetime]'2026-10-21 07:30')
            $r.Lines[0].Day | Should -BeExactly '2026-10-21'
        }

        It '1300 OFT x2 gives two separate OFT lines that export with SIM IP' {
            $r = Get-TQ '1300 OFT x2'
            $r.Errors.Count | Should -Be 0
            $r.Lines.Count | Should -Be 2
            foreach ($l in $r.Lines) {
                $l.SortieType | Should -BeExactly 'OFT'
                $l.ETD | Should -Be ([datetime]'2026-10-20 13:00')
                $l.Instructor | Should -BeExactly ''
            }
            $r.Lines[0].Id | Should -Not -Be $r.Lines[1].Id
            $b = New-TB
            Add-BoardLine -Board $b -Line $r.Lines
            $rows = @(ConvertTo-ScheduleRows -Board $b)
            $rows.Count | Should -Be 2
            foreach ($row in $rows) { $row.Instructor | Should -BeExactly 'SIM IP' }
        }

        It '0900 ift x3 gives three IFT lines' {
            $r = Get-TQ '0900 ift x3'
            $r.Lines.Count | Should -Be 3
            @($r.Lines | Where-Object { $_.SortieType -ceq 'IFT' }).Count | Should -Be 3
        }

        It 'x2 with a student pins the student on the first line only' {
            $r = Get-TQ '1300 OFT x2 sp:Adams m:FE3102'
            $r.Lines.Count | Should -Be 2
            $r.Lines[0].Student | Should -BeExactly 'Adams, Amy'
            $r.Lines[0].Mission | Should -BeExactly 'FE3102'
            $r.Lines[1].Student | Should -BeExactly ''
            $r.Lines[1].Mission | Should -BeExactly ''
        }
    }

    Context 'O&B and its return' {
        It '1015 O&B adds the O&B and a linked O&B Return at 1315' {
            $r = Get-TQ '1015 O&B'
            $r.Errors.Count | Should -Be 0
            $r.Warnings.Count | Should -Be 0
            $r.Lines.Count | Should -Be 2
            $ob = $r.Lines[0]; $ret = $r.Lines[1]
            $ob.SortieType | Should -BeExactly 'O&B'
            $ob.ETD | Should -Be ([datetime]'2026-10-20 10:15')
            $ret.SortieType | Should -BeExactly 'O&B Return'
            $ret.ETD | Should -Be ([datetime]'2026-10-20 13:15')
            $ret.Day | Should -BeExactly '2026-10-20'
            $ob.LinkedLineId | Should -BeExactly $ret.Id
            $ret.LinkedLineId | Should -BeExactly $ob.Id
            $ret.Student | Should -BeExactly ''
        }

        It '1015 OB 1330 puts the return at 1330' {
            $r = Get-TQ '1015 OB 1330'
            $r.Errors.Count | Should -Be 0
            $r.Lines.Count | Should -Be 2
            $r.Lines[0].SortieType | Should -BeExactly 'O&B'
            $r.Lines[1].ETD | Should -Be ([datetime]'2026-10-20 13:30')
            $r.Lines[0].LinkedLineId | Should -BeExactly $r.Lines[1].Id
        }

        It 'the O&B IP flies the return too' {
            $r = Get-TQ '1015 o & b 1400 Ortiz'
            $r.Errors.Count | Should -Be 0
            $r.Lines[0].SortieType | Should -BeExactly 'O&B'
            $r.Lines[0].Instructor | Should -BeExactly 'Ortiz, Sam'
            $r.Lines[1].Instructor | Should -BeExactly 'Ortiz, Sam'
            $r.Lines[1].ETD | Should -Be ([datetime]'2026-10-20 14:00')
        }

        It '1015 O&B x2 gives two O&Bs, each with its own return' {
            $r = Get-TQ '1015 O&B x2'
            $r.Lines.Count | Should -Be 4
            $obs = @($r.Lines | Where-Object { $_.SortieType -eq 'O&B' })
            $obs.Count | Should -Be 2
            $obs[0].LinkedLineId | Should -Not -Be $obs[1].LinkedLineId
            foreach ($ob in $obs) {
                $ret = $r.Lines | Where-Object { $_.Id -eq $ob.LinkedLineId }
                $ret.SortieType | Should -BeExactly 'O&B Return'
                $ret.LinkedLineId | Should -BeExactly $ob.Id
                $ret.ETD | Should -Be ([datetime]'2026-10-20 13:15')
            }
        }

        It 'uses the offset from -Settings or from the board' {
            $r = Get-TQ '1015 O&B' -Settings @{ OBReturnOffsetMinutes = 120; ReturnSortieType = 'O&B Return' }
            $r.Lines[1].ETD | Should -Be ([datetime]'2026-10-20 12:15')
            $b = New-TB -Settings @{ OBReturnOffsetMinutes = 150 }
            $r2 = Get-TQ '1015 O&B' -Board $b
            $r2.Lines[1].ETD | Should -Be ([datetime]'2026-10-20 12:45')
        }

        It 'warns when the return time is not after the O&B but still adds both lines' {
            $r = Get-TQ '1015 OB 0900'
            $r.Errors.Count | Should -Be 0
            $r.Lines.Count | Should -Be 2
            $r.Warnings.Count | Should -Be 1
            $r.Warnings[0] | Should -BeExactly "Return time 0900 is not after the O&B in '1015 OB 0900'"
        }
    }

    Context 'RTN linking' {
        BeforeEach {
            $script:B = New-TB
            $script:Ob0800 = New-TL -At '2026-10-20 08:00' -Type 'O&B' -IP 'Lee, Ann'
            $script:Ob1000 = New-TL -At '2026-10-20 10:00' -Type 'O&B' -IP 'Jones, Bob'
            Add-TL $script:B @($script:Ob0800, $script:Ob1000)
            $script:Linked = Add-TOB $script:B '2026-10-20 11:00' 'Ortiz, Sam'
        }

        It '1330 RTN links to the latest unlinked O&B earlier that day on the board' {
            $r = Get-TQ '1330 RTN' -Board $script:B
            $r.Errors.Count | Should -Be 0
            $r.Warnings.Count | Should -Be 0
            $r.Lines.Count | Should -Be 1
            $ret = $r.Lines[0]
            $ret.SortieType | Should -BeExactly 'O&B Return'
            $ret.ETD | Should -Be ([datetime]'2026-10-20 13:30')
            $ret.LinkedLineId | Should -BeExactly $script:Ob1000.Id
            $ret.Instructor | Should -BeExactly 'Jones, Bob'
            # Parsing alone changes nothing; the link is applied when the lines are added
            $script:Ob1000.LinkedLineId | Should -BeExactly ''
            Add-BoardLine -Board $script:B -Line $r.Lines -Link $r.Links
            $script:Ob1000.LinkedLineId | Should -BeExactly $ret.Id
            $script:Ob0800.LinkedLineId | Should -BeExactly ''
            $script:Linked.OB.LinkedLineId | Should -BeExactly $script:Linked.Ret.Id
        }

        It 'a return before 1000 takes the 0800 O&B' {
            $r = Get-TQ '0930 RTN' -Board $script:B
            $r.Lines[0].LinkedLineId | Should -BeExactly $script:Ob0800.Id
            $r.Lines[0].Instructor | Should -BeExactly 'Lee, Ann'
        }

        It 'two returns in one entry take the two O&Bs latest first' {
            $r = Get-TQ '1330 RTN, 1400 RTN' -Board $script:B
            $r.Errors.Count | Should -Be 0
            $r.Lines.Count | Should -Be 2
            $r.Lines[0].LinkedLineId | Should -BeExactly $script:Ob1000.Id
            $r.Lines[1].LinkedLineId | Should -BeExactly $script:Ob0800.Id
        }

        It 'an explicit IP on the return is kept' {
            $r = Get-TQ '1330 RTN Leech' -Board $script:B
            $r.Lines[0].Instructor | Should -BeExactly 'Leech, Tom'
            $r.Lines[0].LinkedLineId | Should -BeExactly $script:Ob1000.Id
        }

        It 'O&B Return typed out in full is a return' {
            $r = Get-TQ '1330 O&B Return' -Board $script:B
            $r.Errors.Count | Should -Be 0
            $r.Lines[0].SortieType | Should -BeExactly 'O&B Return'
            $r.Lines[0].LinkedLineId | Should -BeExactly $script:Ob1000.Id
        }

        It 'never links to an O&B that is later, on another day, or already linked' {
            $b = New-TB -Days @('2026-10-20', '2026-10-21')
            $late = New-TL -At '2026-10-20 14:00' -Type 'O&B'
            $otherDay = New-TL -At '2026-10-21 08:00' -Type 'O&B'
            Add-TL $b @($late, $otherDay)
            [void](Add-TOB $b '2026-10-20 08:00')
            $r = Get-TQ '1330 RTN' -Board $b
            $r.Errors.Count | Should -Be 0
            $r.Lines.Count | Should -Be 1
            $r.Lines[0].LinkedLineId | Should -BeExactly ''
            $r.Warnings.Count | Should -Be 1
            $r.Warnings[0] | Should -BeExactly 'No O&B without a return before 1330; return line added unlinked'
        }

        It 'without a board the return is added unlinked with a warning' {
            $r = Get-TQ '1330 RTN'
            $r.Lines.Count | Should -Be 1
            $r.Lines[0].LinkedLineId | Should -BeExactly ''
            $r.Warnings[0] | Should -BeExactly 'No O&B without a return before 1330; return line added unlinked'
        }

        It 'a return later in the same text does not take an O&B that already got its auto return' {
            # Every quick-entry O&B gets its own return, so the 1330 RTN has nothing to link to
            $r = Get-TQ '1015 O&B, 1330 RTN'
            $r.Errors.Count | Should -Be 0
            $r.Lines.Count | Should -Be 3
            $r.Lines[0].LinkedLineId | Should -BeExactly $r.Lines[1].Id
            $r.Lines[1].ETD | Should -Be ([datetime]'2026-10-20 13:15')
            $r.Lines[2].LinkedLineId | Should -BeExactly ''
            $r.Warnings[0] | Should -BeExactly 'No O&B without a return before 1330; return line added unlinked'
        }

        It 'a student on a return line is ignored with a warning' {
            $r = Get-TQ '1330 RTN sp:Adams' -Board $script:B
            $r.Errors.Count | Should -Be 0
            $r.Lines.Count | Should -Be 1
            $r.Lines[0].Student | Should -BeExactly ''
            $r.Warnings[0] | Should -BeExactly "Return lines follow their O&B; student ignored on '1330 RTN sp:Adams'"
        }

        It 'does not change the board when an entry in the text fails' {
            # Parsing must be side-effect free: the box shows the error and adds nothing,
            # so the board's O&B must not be left pointing at a line that never got added.
            $before = Get-TSnap $script:B
            $r = Get-TQ '1330 RTN, 2500 DL' -Board $script:B
            $r.Errors.Count | Should -Be 1
            $script:Ob1000.LinkedLineId | Should -BeExactly ''
            (Get-TSnap $script:B) | Should -BeExactly $before
        }

        It 'undoing the added return leaves the O&B unlinked' {
            $r = Get-TQ '1330 RTN' -Board $script:B
            Add-BoardLine -Board $script:B -Line $r.Lines
            (Undo-Board $script:B) | Should -BeTrue
            $ob = Get-BoardLine $script:B $script:Ob1000.Id
            $null -eq (Get-BoardLine $script:B $r.Lines[0].Id) | Should -BeTrue
            $ob.LinkedLineId | Should -BeExactly ''
        }
    }

    Context 'IP, student and mission' {
        It '0800 DL Jones sp:Adams m:FT5103 sets the IP, the pinned student and the mission' {
            $r = Get-TQ '0800 DL Jones sp:Adams m:FT5103'
            $r.Errors.Count | Should -Be 0
            $r.Warnings.Count | Should -Be 0
            $r.Lines.Count | Should -Be 1
            $l = $r.Lines[0]
            $l.ETD | Should -Be ([datetime]'2026-10-20 08:00')
            $l.SortieType | Should -BeExactly 'Day Local'
            $l.Instructor | Should -BeExactly 'Jones, Bob'
            $l.Student | Should -BeExactly 'Adams, Amy'
            $l.Mission | Should -BeExactly 'FT5103'
        }

        It '<Typed> resolves to <Want>' -TestCases @(
            @{ Typed = 'Lee'; Want = 'Lee, Ann' }
            @{ Typed = 'lee'; Want = 'Lee, Ann' }
            @{ Typed = 'Leech'; Want = 'Leech, Tom' }
            @{ Typed = 'ip:Ortiz'; Want = 'Ortiz, Sam' }
            @{ Typed = 'ip=Lee'; Want = 'Lee, Ann' }
            @{ Typed = 'Jones'; Want = 'Jones, Bob' }
        ) {
            $r = Get-TQ "0800 DL $Typed"
            $r.Errors.Count | Should -Be 0
            $r.Lines[0].Instructor | Should -BeExactly $Want
        }

        It 'the IP may come before the type' {
            $r = Get-TQ '0800 Lee DL'
            $r.Errors.Count | Should -Be 0
            $r.Lines[0].SortieType | Should -BeExactly 'Day Local'
            $r.Lines[0].Instructor | Should -BeExactly 'Lee, Ann'
        }

        It 'takes stu:, student= and msn: spellings' {
            $r = Get-TQ '0800 DL Ortiz stu:Baker msn:FT5001, 0900 DL student=cruz mission=FT5102'
            $r.Errors.Count | Should -Be 0
            $r.Lines[0].Student | Should -BeExactly 'Baker, Ben'
            $r.Lines[0].Mission | Should -BeExactly 'FT5001'
            $r.Lines[0].Instructor | Should -BeExactly 'Ortiz, Sam'
            $r.Lines[1].Student | Should -BeExactly 'Cruz, Cal'
            $r.Lines[1].Mission | Should -BeExactly 'FT5102'
        }
    }

    Context 'lists' {
        It 'comma and semicolon separate entries' {
            $r = Get-TQ '0730 DL, 0800 DL Lee; 1300 OFT'
            $r.Errors.Count | Should -Be 0
            $r.Lines.Count | Should -Be 3
            (@($r.Lines | ForEach-Object { $_.ETD.ToString('HHmm') + ' ' + $_.SortieType }) -join '|') | Should -BeExactly '0730 Day Local|0800 Day Local|1300 OFT'
            $r.Lines[1].Instructor | Should -BeExactly 'Lee, Ann'
        }

        It 'a bad entry does not stop the others' {
            $r = Get-TQ '0730 DL, 2500 DL, 0800 DL'
            $r.Lines.Count | Should -Be 2
            $r.Errors.Count | Should -Be 1
        }

        It 'empty entries and empty text are ignored' {
            $r = Get-TQ '0730 DL,, ; '
            $r.Lines.Count | Should -Be 1
            $r.Errors.Count | Should -Be 0
            $e = Get-TQ ''
            $e.Lines.Count | Should -Be 0
            $e.Errors.Count | Should -Be 0
            $e.Warnings.Count | Should -Be 0
        }
    }

    Context 'errors' {
        It '<Text> -> <Msg>' -TestCases @(
            @{ Text = '2500 DL'; Msg = "Invalid time '2500' in '2500 DL' (use 0730, 730 or 7:30)" }
            @{ Text = '0760 DL'; Msg = "Invalid time '0760' in '0760 DL' (use 0730, 730 or 7:30)" }
            @{ Text = 'abc DL'; Msg = "Invalid time 'abc' in 'abc DL' (use 0730, 730 or 7:30)" }
            @{ Text = 'DL 0730'; Msg = "Invalid time 'DL' in 'DL 0730' (use 0730, 730 or 7:30)" }
            @{ Text = '0730'; Msg = "Missing sortie type in '0730' (DL, O&B, RTN, OFT, IFT)" }
            @{ Text = '0730 Jones'; Msg = "Missing sortie type in '0730 Jones' (DL, O&B, RTN, OFT, IFT)" }
            @{ Text = '1300 SIM'; Msg = "'SIM' is ambiguous - use OFT or IFT ('1300 SIM')" }
            @{ Text = '1300 sim x2'; Msg = "'sim' is ambiguous - use OFT or IFT ('1300 sim x2')" }
            @{ Text = '0800 DL Smith'; Msg = "IP 'Smith' not found ('0800 DL Smith')" }
            @{ Text = '0800 DL sp:Nobody'; Msg = "Student 'Nobody' not found ('0800 DL sp:Nobody')" }
        ) {
            $r = Get-TQ $Text
            $r.Lines.Count | Should -Be 0
            $r.Errors.Count | Should -Be 1
            $r.Errors[0] | Should -BeExactly $Msg
        }

        It 'a last name two students share is not found' {
            $students = @(New-TestStudent -Name 'Baker, Ben'; New-TestStudent -Name 'Baker, Bea')
            $r = ConvertFrom-QuickEntry -Text '0800 DL sp:Baker' -Day '2026-10-20' -Instructors $script:IPs -Students $students 3>$null
            $r.Lines.Count | Should -Be 0
            $r.Errors[0] | Should -BeExactly "Student 'Baker' not found ('0800 DL sp:Baker')"
        }

        It 'an IP that is not found adds nothing even with a good student' {
            $r = Get-TQ '0800 DL Smith sp:Adams'
            $r.Lines.Count | Should -Be 0
            $r.Errors[0] | Should -BeExactly "IP 'Smith' not found ('0800 DL Smith sp:Adams')"
        }

        It 'an IP is not found when there are no instructors' {
            $r = ConvertFrom-QuickEntry -Text '0800 DL Jones' -Day '2026-10-20'
            $r.Lines.Count | Should -Be 0
            $r.Errors[0] | Should -BeExactly "IP 'Jones' not found ('0800 DL Jones')"
        }
    }
}


Describe 'Add-BoardLine / Remove-BoardLine' {
    It 'adds lines in one undo step, adds new days and bumps the version' {
        $b = New-TB
        $a = New-TL -At '2026-10-20 07:30'
        $c = New-TL -At '2026-10-22 07:30'
        Add-BoardLine -Board $b -Line @($a, $c)
        $b.Lines.Count | Should -Be 2
        ($b.Days.ToArray() -join ',') | Should -BeExactly '2026-10-20,2026-10-22'
        $b.Version | Should -Be 1
        $b.Undo.Count | Should -Be 1
        (Undo-Board $b) | Should -BeTrue
        $b.Lines.Count | Should -Be 0
        ($b.Days.ToArray() -join ',') | Should -BeExactly '2026-10-20'
    }

    It '-NoUndo adds without an undo point' {
        $b = New-TB
        Add-BoardLine -Board $b -Line @(New-TL -At '2026-10-20 07:30') -NoUndo
        $b.Lines.Count | Should -Be 1
        $b.Undo.Count | Should -Be 0
        $b.Version | Should -Be 1
    }

    It 'removing an O&B takes its return with it' {
        $b = New-TB
        $keep = New-TL -At '2026-10-20 07:30'
        Add-TL $b @($keep)
        $p = Add-TOB $b '2026-10-20 10:15' 'Jones, Bob'
        Remove-BoardLine -Board $b -Id $p.OB.Id
        $b.Lines.Count | Should -Be 1
        $b.Lines[0].Id | Should -BeExactly $keep.Id
    }

    It 'removing a return unlinks its O&B' {
        $b = New-TB
        $p = Add-TOB $b '2026-10-20 10:15' 'Jones, Bob'
        Remove-BoardLine -Board $b -Id $p.Ret.Id
        $b.Lines.Count | Should -Be 1
        $p.OB.LinkedLineId | Should -BeExactly ''
    }

    It 'removing a lead clears WingOfLineId on its wing' {
        $b = New-TB
        $lead = New-TL -At '2026-10-20 08:00' -Student 'Adams, Amy'
        $wing = New-TL -At '2026-10-20 08:00' -Student 'Baker, Ben'
        Add-TL $b @($lead, $wing)
        $wing.WingOfLineId = $lead.Id
        Remove-BoardLine -Board $b -Id $lead.Id
        $wing.WingOfLineId | Should -BeExactly ''
    }

    It 'removing an unknown id changes nothing' {
        $b = New-TB
        Add-TL $b @(New-TL -At '2026-10-20 07:30')
        $v = $b.Version
        Remove-BoardLine -Board $b -Id 'nope'
        $b.Lines.Count | Should -Be 1
        $b.Version | Should -Be $v
        $b.Undo.Count | Should -Be 0
    }

    It 'one undo brings the O&B and its return back, still linked' {
        $b = New-TB
        $p = Add-TOB $b '2026-10-20 10:15' 'Jones, Bob'
        Remove-BoardLine -Board $b -Id $p.OB.Id
        (Undo-Board $b) | Should -BeTrue
        $b.Lines.Count | Should -Be 2
        $ob = Get-BoardLine $b $p.OB.Id
        $ret = Get-BoardLine $b $p.Ret.Id
        $ob.LinkedLineId | Should -BeExactly $p.Ret.Id
        $ret.LinkedLineId | Should -BeExactly $p.OB.Id
        $ret.ETD | Should -Be ([datetime]'2026-10-20 13:15')
    }
}


Describe 'Set-BoardLineEtd' {
    BeforeEach {
        $script:B = New-TB
        $script:P = Add-TOB $script:B '2026-10-20 10:15' 'Jones, Bob' '2026-10-20 13:30'
    }

    It 'moving the O&B later moves its return by the same amount' {
        Set-BoardLineEtd -Board $script:B -Id $script:P.OB.Id -ETD ([datetime]'2026-10-20 11:00')
        $script:P.OB.ETD | Should -Be ([datetime]'2026-10-20 11:00')
        $script:P.Ret.ETD | Should -Be ([datetime]'2026-10-20 14:15')
        $script:P.Ret.Day | Should -BeExactly '2026-10-20'
    }

    It 'moving the O&B earlier moves its return earlier' {
        Set-BoardLineEtd -Board $script:B -Id $script:P.OB.Id -ETD ([datetime]'2026-10-20 08:45')
        $script:P.Ret.ETD | Should -Be ([datetime]'2026-10-20 12:00')
    }

    It 'moving the return leaves the O&B alone' {
        Set-BoardLineEtd -Board $script:B -Id $script:P.Ret.Id -ETD ([datetime]'2026-10-20 15:00')
        $script:P.Ret.ETD | Should -Be ([datetime]'2026-10-20 15:00')
        $script:P.OB.ETD | Should -Be ([datetime]'2026-10-20 10:15')
    }

    It 'moving the O&B to another day takes the return along and adds the day' {
        Set-BoardLineEtd -Board $script:B -Id $script:P.OB.Id -ETD ([datetime]'2026-10-21 10:15')
        $script:P.OB.Day | Should -BeExactly '2026-10-21'
        $script:P.Ret.Day | Should -BeExactly '2026-10-21'
        $script:P.Ret.ETD | Should -Be ([datetime]'2026-10-21 13:30')
        $script:B.Days.Contains('2026-10-21') | Should -BeTrue
    }

    It 'a return pushed past midnight stays on the board, in the rows and in the saved files' {
        # Set-BoardLineEtd adds the O&B's new day but not the return's, so a 1900 O&B moved
        # to 2200 leaves its 0100 return on a day the board does not have: ConvertTo-ScheduleRows
        # and Save-BoardState only walk Board.Days, so the return is dropped and lost on reload.
        $b = New-TB
        $p = Add-TOB $b '2026-10-20 19:00' 'Jones, Bob'
        Set-BoardLineEtd -Board $b -Id $p.OB.Id -ETD ([datetime]'2026-10-20 22:00')
        $p.Ret.ETD | Should -Be ([datetime]'2026-10-21 01:00')
        $b.Days.Contains('2026-10-21') | Should -BeTrue
        @(ConvertTo-ScheduleRows -Board $b).Count | Should -Be 2
        $dir = Join-Path $TestDrive 'midnight'
        Save-BoardState -Board $b -Dir $dir
        (Open-BoardState -Dir $dir -Days @('2026-10-20', '2026-10-21')).Lines.Count | Should -Be 2
    }

    It 'one undo puts both legs back' {
        $v = $script:B.Version
        Set-BoardLineEtd -Board $script:B -Id $script:P.OB.Id -ETD ([datetime]'2026-10-20 11:00')
        $script:B.Version | Should -Be ($v + 1)
        (Undo-Board $script:B) | Should -BeTrue
        (Get-BoardLine $script:B $script:P.OB.Id).ETD | Should -Be ([datetime]'2026-10-20 10:15')
        (Get-BoardLine $script:B $script:P.Ret.Id).ETD | Should -Be ([datetime]'2026-10-20 13:30')
    }
}


Describe 'Set-BoardLineType' {
    It 'Day Local -> O&B adds a linked return at +180 min with the same IP' {
        $b = New-TB
        $l = New-TL -At '2026-10-20 09:00' -IP 'Jones, Bob'
        Add-TL $b @($l)
        Set-BoardLineType -Board $b -Id $l.Id -SortieType 'O&B'
        $b.Lines.Count | Should -Be 2
        $ret = Get-BoardLine $b $l.LinkedLineId
        $ret.SortieType | Should -BeExactly 'O&B Return'
        $ret.ETD | Should -Be ([datetime]'2026-10-20 12:00')
        $ret.Instructor | Should -BeExactly 'Jones, Bob'
        $ret.LinkedLineId | Should -BeExactly $l.Id
    }

    It 'uses the board offset and return type name' {
        $b = New-TB -Settings @{ OBReturnOffsetMinutes = 90; ReturnSortieType = 'Return Leg' }
        $l = New-TL -At '2026-10-20 09:00'
        Add-TL $b @($l)
        Set-BoardLineType -Board $b -Id $l.Id -SortieType 'O&B'
        $ret = Get-BoardLine $b $l.LinkedLineId
        $ret.ETD | Should -Be ([datetime]'2026-10-20 10:30')
        $ret.SortieType | Should -BeExactly 'Return Leg'
    }

    It 'O&B -> <To> removes the return' -TestCases @(
        @{ To = 'Day Local' }
        @{ To = 'OFT' }
        @{ To = 'IFT' }
    ) {
        $b = New-TB
        $p = Add-TOB $b '2026-10-20 10:15' 'Jones, Bob'
        Set-BoardLineType -Board $b -Id $p.OB.Id -SortieType $To
        $b.Lines.Count | Should -Be 1
        $p.OB.SortieType | Should -BeExactly $To
        $p.OB.LinkedLineId | Should -BeExactly ''
        $null -eq (Get-BoardLine $b $p.Ret.Id) | Should -BeTrue
    }

    It 'the same type is a no-op' {
        $b = New-TB
        $p = Add-TOB $b '2026-10-20 10:15'
        $v = $b.Version
        Set-BoardLineType -Board $b -Id $p.OB.Id -SortieType 'O&B'
        $b.Version | Should -Be $v
        $b.Lines.Count | Should -Be 2
        $b.Undo.Count | Should -Be 0
    }

    It 'becoming a return clears the student and mission' {
        $b = New-TB
        $l = New-TL -At '2026-10-20 13:00' -Student 'Adams, Amy' -Mission 'FT5103'
        Add-TL $b @($l)
        Set-BoardLineType -Board $b -Id $l.Id -SortieType 'O&B Return'
        $l.Student | Should -BeExactly ''
        $l.Mission | Should -BeExactly ''
    }

    It 'one undo removes the created return' {
        $b = New-TB
        $l = New-TL -At '2026-10-20 09:00'
        Add-TL $b @($l)
        Set-BoardLineType -Board $b -Id $l.Id -SortieType 'O&B'
        (Undo-Board $b) | Should -BeTrue
        $b.Lines.Count | Should -Be 1
        $b.Lines[0].SortieType | Should -BeExactly 'Day Local'
        $b.Lines[0].LinkedLineId | Should -BeExactly ''
    }

    It 'turning a paired wing into <To> drops the wing pairing' -TestCases @(
        @{ To = 'OFT' }
        @{ To = 'IFT' }
        @{ To = 'O&B' }
    ) {
        # User rule: OFT / IFT / O&B lines can never fly wing. Set-FormationPair refuses
        # them, but a type change leaves the old pairing behind ("wing to ..." on the board).
        $b = New-TB
        $lead = New-TL -At '2026-10-20 08:00' -Student 'Adams, Amy'
        $wing = New-TL -At '2026-10-20 08:00' -Student 'Baker, Ben'
        Add-TL $b @($lead, $wing)
        (Set-FormationPair -Board $b -WingId $wing.Id -LeadId $lead.Id) | Should -BeTrue
        Set-BoardLineType -Board $b -Id $wing.Id -SortieType $To
        $wing.WingOfLineId | Should -BeExactly ''
    }

    It 'turning a lead into <To> drops the wing that pointed at it' -TestCases @(
        @{ To = 'OFT' }
        @{ To = 'IFT' }
        @{ To = 'O&B' }
    ) {
        $b = New-TB
        $lead = New-TL -At '2026-10-20 08:00' -Student 'Adams, Amy'
        $wing = New-TL -At '2026-10-20 08:00' -Student 'Baker, Ben'
        Add-TL $b @($lead, $wing)
        (Set-FormationPair -Board $b -WingId $wing.Id -LeadId $lead.Id) | Should -BeTrue
        Set-BoardLineType -Board $b -Id $lead.Id -SortieType $To
        $wing.WingOfLineId | Should -BeExactly ''
    }

    It 'a linked return changed to Day Local is no longer the O&B return' {
        # The type box is enabled on return rows. After the change the O&B still points
        # at the Day Local line: the rows hand the engine that link (Schedule-ReturnLeg then
        # overwrites the pinned student with the O&B student), and deleting the O&B deletes it.
        $b = New-TB
        $p = Add-TOB $b '2026-10-20 10:15' 'Jones, Bob'
        Set-BoardLineType -Board $b -Id $p.Ret.Id -SortieType 'Day Local'
        Set-BoardInstructor -Board $b -Id $p.Ret.Id -Name 'Lee, Ann'
        [void](Set-BoardPin -Board $b -Id $p.Ret.Id -Student 'Baker, Ben')
        $p.OB.LinkedLineId | Should -BeExactly ''
        $p.Ret.LinkedLineId | Should -BeExactly ''
        $rows = @(ConvertTo-ScheduleRows -Board $b)
        foreach ($r in $rows) { $null -eq $r.LinkedLineId | Should -BeTrue }
        Remove-BoardLine -Board $b -Id $p.OB.Id
        $null -ne (Get-BoardLine $b $p.Ret.Id) | Should -BeTrue
    }
}


Describe 'Set-BoardInstructor' {
    It 'an O&B IP is mirrored onto its return' {
        $b = New-TB
        $p = Add-TOB $b '2026-10-20 10:15'
        Set-BoardInstructor -Board $b -Id $p.OB.Id -Name 'Jones, Bob'
        $p.OB.Instructor | Should -BeExactly 'Jones, Bob'
        $p.Ret.Instructor | Should -BeExactly 'Jones, Bob'
        Set-BoardInstructor -Board $b -Id $p.OB.Id -Name ''
        $p.Ret.Instructor | Should -BeExactly ''
        $b.Undo.Count | Should -Be 2
    }

    It 'a Day Local IP changes only that line' {
        $b = New-TB
        $a = New-TL -At '2026-10-20 07:30'
        $c = New-TL -At '2026-10-20 07:30'
        Add-TL $b @($a, $c)
        Set-BoardInstructor -Board $b -Id $a.Id -Name 'Lee, Ann'
        $a.Instructor | Should -BeExactly 'Lee, Ann'
        $c.Instructor | Should -BeExactly ''
    }

    It 'setting the return IP does not change the O&B' {
        $b = New-TB
        $p = Add-TOB $b '2026-10-20 10:15' 'Jones, Bob'
        Set-BoardInstructor -Board $b -Id $p.Ret.Id -Name 'Lee, Ann'
        $p.Ret.Instructor | Should -BeExactly 'Lee, Ann'
        $p.OB.Instructor | Should -BeExactly 'Jones, Bob'
    }
}


Describe 'Move-BoardInstructor' {
    BeforeEach {
        $script:B = New-TB
        $script:A = New-TL -At '2026-10-20 07:30' -IP 'Jones, Bob'
        $script:C = New-TL -At '2026-10-20 08:00'
        $script:E = New-TL -At '2026-10-20 09:00' -IP 'Lee, Ann'
        Add-TL $script:B @($script:A, $script:C, $script:E)
        $script:P = Add-TOB $script:B '2026-10-20 10:15' 'Ortiz, Sam'
    }

    It 'moves onto an empty line' {
        (Move-BoardInstructor -Board $script:B -FromId $script:A.Id -ToId $script:C.Id) | Should -BeTrue
        $script:C.Instructor | Should -BeExactly 'Jones, Bob'
        $script:A.Instructor | Should -BeExactly ''
    }

    It 'swaps when the target has an IP' {
        (Move-BoardInstructor -Board $script:B -FromId $script:A.Id -ToId $script:E.Id) | Should -BeTrue
        $script:A.Instructor | Should -BeExactly 'Lee, Ann'
        $script:E.Instructor | Should -BeExactly 'Jones, Bob'
    }

    It '-Copy keeps the original' {
        (Move-BoardInstructor -Board $script:B -FromId $script:A.Id -ToId $script:E.Id -Copy) | Should -BeTrue
        $script:A.Instructor | Should -BeExactly 'Jones, Bob'
        $script:E.Instructor | Should -BeExactly 'Jones, Bob'
    }

    It 'moving onto an O&B puts the IP on its return too' {
        [void](Move-BoardInstructor -Board $script:B -FromId $script:A.Id -ToId $script:P.OB.Id)
        $script:P.OB.Instructor | Should -BeExactly 'Jones, Bob'
        $script:P.Ret.Instructor | Should -BeExactly 'Jones, Bob'
        $script:A.Instructor | Should -BeExactly 'Ortiz, Sam'
    }

    It 'moving off an O&B moves its return IP too' {
        [void](Move-BoardInstructor -Board $script:B -FromId $script:P.OB.Id -ToId $script:C.Id)
        $script:C.Instructor | Should -BeExactly 'Ortiz, Sam'
        $script:P.OB.Instructor | Should -BeExactly ''
        $script:P.Ret.Instructor | Should -BeExactly ''
    }

    It 'copying off an O&B leaves the O&B and its return as they were' {
        [void](Move-BoardInstructor -Board $script:B -FromId $script:P.OB.Id -ToId $script:E.Id -Copy)
        $script:E.Instructor | Should -BeExactly 'Ortiz, Sam'
        $script:P.OB.Instructor | Should -BeExactly 'Ortiz, Sam'
        $script:P.Ret.Instructor | Should -BeExactly 'Ortiz, Sam'
    }

    It 'is a single undo step' {
        $u = $script:B.Undo.Count
        $v = $script:B.Version
        [void](Move-BoardInstructor -Board $script:B -FromId $script:A.Id -ToId $script:P.OB.Id)
        $script:B.Undo.Count | Should -Be ($u + 1)
        $script:B.Version | Should -Be ($v + 1)
        (Undo-Board $script:B) | Should -BeTrue
        (Get-BoardLine $script:B $script:A.Id).Instructor | Should -BeExactly 'Jones, Bob'
        (Get-BoardLine $script:B $script:P.OB.Id).Instructor | Should -BeExactly 'Ortiz, Sam'
        (Get-BoardLine $script:B $script:P.Ret.Id).Instructor | Should -BeExactly 'Ortiz, Sam'
    }

    It 'returns $false and changes nothing for the same line or an unknown line' {
        $v = $script:B.Version
        (Move-BoardInstructor -Board $script:B -FromId $script:A.Id -ToId $script:A.Id) | Should -BeFalse
        (Move-BoardInstructor -Board $script:B -FromId $script:A.Id -ToId 'nope') | Should -BeFalse
        (Move-BoardInstructor -Board $script:B -FromId 'nope' -ToId $script:A.Id) | Should -BeFalse
        $script:B.Version | Should -Be $v
        $script:B.Undo.Count | Should -Be 0
    }
}


Describe 'Set-BoardPin / Move-BoardPin / Clear-BoardPin' {
    BeforeEach {
        $script:B = New-TB
        $script:A = New-TL -At '2026-10-20 07:30' -IP 'Jones, Bob'
        $script:C = New-TL -At '2026-10-20 08:00' -IP 'Lee, Ann'
        $script:E = New-TL -At '2026-10-20 09:00' -IP 'Ortiz, Sam'
        Add-TL $script:B @($script:A, $script:C, $script:E)
        $script:P = Add-TOB $script:B '2026-10-20 10:15' 'Leech, Tom'
    }

    Context 'Set-BoardPin' {
        It 'pins the student and mission' {
            $v = $script:B.Version
            (Set-BoardPin -Board $script:B -Id $script:A.Id -Student 'Adams, Amy' -Mission 'FT5103') | Should -BeTrue
            $script:A.Student | Should -BeExactly 'Adams, Amy'
            $script:A.Mission | Should -BeExactly 'FT5103'
            $script:B.Version | Should -Be ($v + 1)
            $script:B.Undo.Count | Should -Be 1
        }

        It 'refuses a return line and changes nothing' {
            $before = Get-TSnap $script:B
            $v = $script:B.Version
            (Set-BoardPin -Board $script:B -Id $script:P.Ret.Id -Student 'Adams, Amy') | Should -BeFalse
            $script:P.Ret.Student | Should -BeExactly ''
            $script:B.Version | Should -Be $v
            $script:B.Undo.Count | Should -Be 0
            (Get-TSnap $script:B) | Should -BeExactly $before
        }

        It 'can pin an O&B (the return follows it)' {
            (Set-BoardPin -Board $script:B -Id $script:P.OB.Id -Student 'Adams, Amy') | Should -BeTrue
            $script:P.OB.Student | Should -BeExactly 'Adams, Amy'
            $script:P.Ret.Student | Should -BeExactly ''
        }

        It 'returns $false for an unknown line' {
            (Set-BoardPin -Board $script:B -Id 'nope' -Student 'Adams, Amy') | Should -BeFalse
        }

        It 'a different student clears the wing pairing, the same student keeps it' {
            [void](Set-BoardPin -Board $script:B -Id $script:A.Id -Student 'Adams, Amy')
            [void](Set-BoardPin -Board $script:B -Id $script:C.Id -Student 'Baker, Ben')
            [void](Set-FormationPair -Board $script:B -WingId $script:C.Id -LeadId $script:A.Id)
            [void](Set-BoardPin -Board $script:B -Id $script:C.Id -Student 'Baker, Ben' -Mission 'FF5201')
            $script:C.WingOfLineId | Should -BeExactly $script:A.Id
            [void](Set-BoardPin -Board $script:B -Id $script:C.Id -Student 'Cruz, Cal')
            $script:C.WingOfLineId | Should -BeExactly ''
        }
    }

    Context 'Move-BoardPin' {
        BeforeEach {
            [void](Set-BoardPin -Board $script:B -Id $script:A.Id -Student 'Adams, Amy' -Mission 'FT5103')
            [void](Set-BoardPin -Board $script:B -Id $script:E.Id -Student 'Baker, Ben' -Mission 'FT5001')
        }

        It 'moves onto an empty line' {
            (Move-BoardPin -Board $script:B -FromId $script:A.Id -ToId $script:C.Id) | Should -BeTrue
            $script:C.Student | Should -BeExactly 'Adams, Amy'
            $script:C.Mission | Should -BeExactly 'FT5103'
            $script:A.Student | Should -BeExactly ''
            $script:A.Mission | Should -BeExactly ''
        }

        It 'swaps when the target is pinned' {
            (Move-BoardPin -Board $script:B -FromId $script:A.Id -ToId $script:E.Id) | Should -BeTrue
            $script:A.Student | Should -BeExactly 'Baker, Ben'
            $script:A.Mission | Should -BeExactly 'FT5001'
            $script:E.Student | Should -BeExactly 'Adams, Amy'
            $script:E.Mission | Should -BeExactly 'FT5103'
        }

        It '-Copy keeps the original and leaves the copy on auto mission' {
            (Move-BoardPin -Board $script:B -FromId $script:A.Id -ToId $script:C.Id -Copy) | Should -BeTrue
            $script:A.Student | Should -BeExactly 'Adams, Amy'
            $script:A.Mission | Should -BeExactly 'FT5103'
            $script:C.Student | Should -BeExactly 'Adams, Amy'
            $script:C.Mission | Should -BeExactly ''
        }

        It 'refuses a return line as the target' {
            $before = Get-TSnap $script:B
            (Move-BoardPin -Board $script:B -FromId $script:A.Id -ToId $script:P.Ret.Id) | Should -BeFalse
            (Get-TSnap $script:B) | Should -BeExactly $before
        }

        It 'refuses the same line or an unknown line' {
            (Move-BoardPin -Board $script:B -FromId $script:A.Id -ToId $script:A.Id) | Should -BeFalse
            (Move-BoardPin -Board $script:B -FromId $script:A.Id -ToId 'nope') | Should -BeFalse
        }

        It 'clears wing pairings on both lines and on lines that pointed at them' {
            [void](Set-BoardPin -Board $script:B -Id $script:C.Id -Student 'Cruz, Cal')
            [void](Set-FormationPair -Board $script:B -WingId $script:C.Id -LeadId $script:A.Id)
            (Move-BoardPin -Board $script:B -FromId $script:A.Id -ToId $script:E.Id) | Should -BeTrue
            $script:C.WingOfLineId | Should -BeExactly ''
            [void](Set-FormationPair -Board $script:B -WingId $script:A.Id -LeadId $script:C.Id)
            (Move-BoardPin -Board $script:B -FromId $script:A.Id -ToId $script:E.Id) | Should -BeTrue
            $script:A.WingOfLineId | Should -BeExactly ''
        }

        It 'is a single undo step' {
            [void](Move-BoardPin -Board $script:B -FromId $script:A.Id -ToId $script:E.Id)
            (Undo-Board $script:B) | Should -BeTrue
            (Get-BoardLine $script:B $script:A.Id).Student | Should -BeExactly 'Adams, Amy'
            (Get-BoardLine $script:B $script:E.Id).Student | Should -BeExactly 'Baker, Ben'
        }
    }

    Context 'Clear-BoardPin' {
        It 'clears the student, the mission and both directions of wing pairing' {
            [void](Set-BoardPin -Board $script:B -Id $script:A.Id -Student 'Adams, Amy' -Mission 'FF5201')
            [void](Set-BoardPin -Board $script:B -Id $script:C.Id -Student 'Baker, Ben')
            [void](Set-BoardPin -Board $script:B -Id $script:E.Id -Student 'Cruz, Cal')
            [void](Set-FormationPair -Board $script:B -WingId $script:C.Id -LeadId $script:A.Id)
            $script:A.WingOfLineId = $script:E.Id
            Clear-BoardPin -Board $script:B -Id $script:A.Id
            $script:A.Student | Should -BeExactly ''
            $script:A.Mission | Should -BeExactly ''
            $script:A.WingOfLineId | Should -BeExactly ''
            $script:C.WingOfLineId | Should -BeExactly ''
            $script:C.Student | Should -BeExactly 'Baker, Ben'
        }
    }
}


Describe 'Set-FormationPair' {
    BeforeEach {
        $script:B = New-TB
        $script:Lead = New-TL -At '2026-10-20 08:00' -IP 'Jones, Bob' -Student 'Adams, Amy'
        $script:Wing = New-TL -At '2026-10-20 08:00' -IP 'Lee, Ann' -Student 'Baker, Ben'
        $script:Wing2 = New-TL -At '2026-10-20 08:00' -IP 'Ortiz, Sam' -Student 'Cruz, Cal'
        Add-TL $script:B @($script:Lead, $script:Wing, $script:Wing2)
    }

    It 'pairs two Day Local lines' {
        $v = $script:B.Version
        (Set-FormationPair -Board $script:B -WingId $script:Wing.Id -LeadId $script:Lead.Id) | Should -BeTrue
        $script:Wing.WingOfLineId | Should -BeExactly $script:Lead.Id
        $script:Lead.WingOfLineId | Should -BeExactly ''
        $script:B.Version | Should -Be ($v + 1)
        $script:B.Undo.Count | Should -Be 1
    }

    It 'a lead has one wing: a new wing replaces the old one' {
        [void](Set-FormationPair -Board $script:B -WingId $script:Wing.Id -LeadId $script:Lead.Id)
        (Set-FormationPair -Board $script:B -WingId $script:Wing2.Id -LeadId $script:Lead.Id) | Should -BeTrue
        $script:Wing2.WingOfLineId | Should -BeExactly $script:Lead.Id
        $script:Wing.WingOfLineId | Should -BeExactly ''
    }

    It 'returns $false and changes nothing when the wing is <Type>' -TestCases @(
        @{ Type = 'OFT' }
        @{ Type = 'IFT' }
        @{ Type = 'O&B' }
        @{ Type = 'O&B Return' }
    ) {
        $w = New-TL -At '2026-10-20 08:00' -Type $Type -Student 'Cruz, Cal'
        Add-TL $script:B @($w)
        $before = Get-TSnap $script:B
        $v = $script:B.Version
        (Set-FormationPair -Board $script:B -WingId $w.Id -LeadId $script:Lead.Id) | Should -BeFalse
        $w.WingOfLineId | Should -BeExactly ''
        $script:B.Version | Should -Be $v
        $script:B.Undo.Count | Should -Be 0
        (Get-TSnap $script:B) | Should -BeExactly $before
    }

    It 'returns $false and changes nothing when the lead is <Type>' -TestCases @(
        @{ Type = 'OFT' }
        @{ Type = 'IFT' }
        @{ Type = 'O&B' }
        @{ Type = 'O&B Return' }
    ) {
        $ld = New-TL -At '2026-10-20 08:00' -Type $Type -Student 'Cruz, Cal'
        Add-TL $script:B @($ld)
        $before = Get-TSnap $script:B
        $v = $script:B.Version
        (Set-FormationPair -Board $script:B -WingId $script:Wing.Id -LeadId $ld.Id) | Should -BeFalse
        $script:Wing.WingOfLineId | Should -BeExactly ''
        $script:B.Version | Should -Be $v
        (Get-TSnap $script:B) | Should -BeExactly $before
    }

    It 'a refused pairing keeps the existing pairing' {
        [void](Set-FormationPair -Board $script:B -WingId $script:Wing.Id -LeadId $script:Lead.Id)
        $sim = New-TL -At '2026-10-20 08:00' -Type 'IFT'
        Add-TL $script:B @($sim)
        (Set-FormationPair -Board $script:B -WingId $script:Wing.Id -LeadId $sim.Id) | Should -BeFalse
        (Set-FormationPair -Board $script:B -WingId $sim.Id -LeadId $script:Lead.Id) | Should -BeFalse
        $script:Wing.WingOfLineId | Should -BeExactly $script:Lead.Id
        $sim.WingOfLineId | Should -BeExactly ''
    }

    It 'LeadId empty clears the pairing' {
        [void](Set-FormationPair -Board $script:B -WingId $script:Wing.Id -LeadId $script:Lead.Id)
        (Set-FormationPair -Board $script:B -WingId $script:Wing.Id -LeadId '') | Should -BeTrue
        $script:Wing.WingOfLineId | Should -BeExactly ''
        (Undo-Board $script:B) | Should -BeTrue
        (Get-BoardLine $script:B $script:Wing.Id).WingOfLineId | Should -BeExactly $script:Lead.Id
    }

    It 'refuses itself, an unknown lead or an unknown wing' {
        (Set-FormationPair -Board $script:B -WingId $script:Wing.Id -LeadId $script:Wing.Id) | Should -BeFalse
        (Set-FormationPair -Board $script:B -WingId $script:Wing.Id -LeadId 'nope') | Should -BeFalse
        (Set-FormationPair -Board $script:B -WingId 'nope' -LeadId $script:Lead.Id) | Should -BeFalse
        $script:B.Undo.Count | Should -Be 0
    }
}


Describe 'Undo-Board' {
    It 'returns $false with nothing to undo and leaves the version alone' {
        $b = New-TB
        (Undo-Board $b) | Should -BeFalse
        $b.Version | Should -Be 0
    }

    It 'undoes edits newest first' {
        $b = New-TB
        $l = New-TL -At '2026-10-20 08:00'
        Add-BoardLine -Board $b -Line @($l)
        [void](Set-BoardPin -Board $b -Id $l.Id -Student 'Adams, Amy' -Mission 'FT5103')
        Set-BoardInstructor -Board $b -Id $l.Id -Name 'Jones, Bob'
        (Undo-Board $b) | Should -BeTrue
        $x = Get-BoardLine $b $l.Id
        $x.Instructor | Should -BeExactly ''
        $x.Student | Should -BeExactly 'Adams, Amy'
        $x.Mission | Should -BeExactly 'FT5103'
        (Undo-Board $b) | Should -BeTrue
        (Get-BoardLine $b $l.Id).Student | Should -BeExactly ''
        (Undo-Board $b) | Should -BeTrue
        $b.Lines.Count | Should -Be 0
        (Undo-Board $b) | Should -BeFalse
    }

    It 'restores every field, the links and the days exactly' {
        $b = New-TB -Days @('2026-10-20', '2026-10-21')
        $lead = New-TL -At '2026-10-20 08:00' -IP 'Jones, Bob' -Student 'Adams, Amy' -Mission 'FF5201' -Notes 'a | b "q"' -Callsign 'HAZE' -Num '11'
        $wing = New-TL -At '2026-10-20 08:00' -IP 'Lee, Ann' -Student 'Baker, Ben'
        $wing.WingOfLineId = $lead.Id
        Add-TL $b @($lead, $wing)
        $p = Add-TOB $b '2026-10-21 10:15' 'Ortiz, Sam' '2026-10-21 13:40'
        $before = Get-TSnap $b
        Remove-BoardDay -Board $b -Day '2026-10-21'
        Remove-BoardLine -Board $b -Id $lead.Id
        (Undo-Board $b) | Should -BeTrue
        (Undo-Board $b) | Should -BeTrue
        (Get-TSnap $b) | Should -BeExactly $before
        $r = Get-BoardLine $b $p.Ret.Id
        $r.ETD | Should -Be ([datetime]'2026-10-21 13:40')
        $r.ETD | Should -BeOfType [datetime]
        $r.Day | Should -BeExactly '2026-10-21'
        $r.LinkedLineId | Should -BeExactly $p.OB.Id
        (Get-BoardLine $b $wing.Id).WingOfLineId | Should -BeExactly $lead.Id
        (Get-BoardLine $b $lead.Id).Notes | Should -BeExactly 'a | b "q"'
        (Get-BoardLine $b $lead.Id).Num | Should -BeExactly '11'
    }

    It 'Version always goes up, through edits and undos' {
        $b = New-TB
        $seen = New-Object System.Collections.Generic.List[int]
        $seen.Add([int]$b.Version)
        $l = New-TL -At '2026-10-20 08:00'
        Add-BoardLine -Board $b -Line @($l); $seen.Add([int]$b.Version)
        Set-BoardInstructor -Board $b -Id $l.Id -Name 'Jones, Bob'; $seen.Add([int]$b.Version)
        [void](Undo-Board $b); $seen.Add([int]$b.Version)
        Set-BoardLineEtd -Board $b -Id $l.Id -ETD ([datetime]'2026-10-20 09:00'); $seen.Add([int]$b.Version)
        [void](Undo-Board $b); $seen.Add([int]$b.Version)
        [void](Undo-Board $b); $seen.Add([int]$b.Version)
        for ($i = 1; $i -lt $seen.Count; $i++) { $seen[$i] | Should -BeGreaterThan $seen[$i - 1] }
        $b.Lines.Count | Should -Be 0
    }

    It 'keeps only UndoDepth snapshots' {
        $b = New-TB -Settings @{ UndoDepth = 5 }
        $l = New-TL -At '2026-10-20 08:00'
        Add-TL $b @($l)
        for ($i = 1; $i -le 8; $i++) { Set-BoardLineField -Board $b -Id $l.Id -Field 'Notes' -Value "n$i" }
        $b.Undo.Count | Should -Be 5
        for ($i = 1; $i -le 5; $i++) { (Undo-Board $b) | Should -BeTrue }
        (Undo-Board $b) | Should -BeFalse
        (Get-BoardLine $b $l.Id).Notes | Should -BeExactly 'n3'
    }

    It 'the default depth is 50' {
        $b = New-TB
        $l = New-TL -At '2026-10-20 08:00'
        Add-TL $b @($l)
        for ($i = 1; $i -le 55; $i++) { Set-BoardLineField -Board $b -Id $l.Id -Field 'Notes' -Value "n$i" }
        $b.Undo.Count | Should -Be 50
    }
}


Describe 'ConvertTo-ScheduleRows' {
    BeforeAll {
        $script:SB = New-TB -Days @('2026-10-20', '2026-10-21')
        $script:SDl = New-TL -At '2026-10-20 07:30'
        $script:SPin = New-TL -At '2026-10-20 08:00' -IP '  Jones, Bob ' -Student ' Adams, Amy ' -Mission ' FT5103 ' -Callsign 'HAZE' -Num '11' -Notes 'Wx'
        $script:SWing = New-TL -At '2026-10-20 08:00' -IP 'Lee, Ann' -Student 'Baker, Ben'
        $script:SWing.WingOfLineId = $script:SPin.Id
        $script:SOft = New-TL -At '2026-10-20 13:00' -Type 'OFT'
        $script:SIft = New-TL -At '2026-10-20 06:00' -Type 'IFT' -IP 'Ortiz, Sam'
        $script:SDay2 = New-TL -At '2026-10-21 07:00' -IP 'Lee, Ann'
        Add-TL $script:SB @($script:SDay2, $script:SOft, $script:SPin, $script:SDl, $script:SWing, $script:SIft)
        $script:SP = Add-TOB $script:SB '2026-10-20 10:15' 'Ortiz, Sam'
        # A stray student on a return line never reaches the engine
        $script:SP.Ret.Student = 'Cruz, Cal'
        $script:SP.Ret.Mission = 'FT5102'
        $script:SRows = @(ConvertTo-ScheduleRows -Board $script:SB)
        $script:ById = @{}
        foreach ($r in $script:SRows) { $script:ById[$r.BoardLineId] = $r }
    }

    It 'has exactly the IngestData columns plus the board ids, in order' {
        foreach ($r in $script:SRows) { ($r.PSObject.Properties.Name -join ',') | Should -BeExactly $script:RowProps }
    }

    It 'has the same shape the fill tests build with New-TestRow' {
        ((New-TestRow -ETD '2026-10-20 07:30').PSObject.Properties.Name -join ',') | Should -BeExactly $script:RowProps
    }

    It 'emits one row per line, sorted by day then ETD' {
        $script:SRows.Count | Should -Be 8
        (@($script:SRows | ForEach-Object { $_.ETD.Substring(0, 16) + ' ' + $_.SortieType }) -join '|') | Should -BeExactly (
            '2026-10-20 06:00 IFT|2026-10-20 07:30 Day Local|2026-10-20 08:00 Day Local|2026-10-20 08:00 Day Local|' +
            '2026-10-20 10:15 O&B|2026-10-20 13:00 OFT|2026-10-20 13:15 O&B Return|2026-10-21 07:00 Day Local')
    }

    It 'writes ETD as yyyy-MM-dd HH:mm:ss text' {
        $r = $script:ById[$script:SDl.Id]
        $r.ETD | Should -BeOfType [string]
        $r.ETD | Should -BeExactly '2026-10-20 07:30:00'
    }

    It 'turns blanks into $null' {
        $r = $script:ById[$script:SDl.Id]
        foreach ($p in @('Instructor', 'Student', 'Mission', 'Callsign', 'Num', 'Notes', 'LinkedLineId', 'WingOfLineId', 'FilledBy')) {
            $null -eq $r.$p | Should -BeTrue -Because "$p is blank on the board"
        }
        $r.SortieType | Should -BeExactly 'Day Local'
        $r.BoardLineId | Should -BeExactly $script:SDl.Id
    }

    It 'trims the pinned values and keeps callsign, num and notes' {
        $r = $script:ById[$script:SPin.Id]
        $r.Instructor | Should -BeExactly 'Jones, Bob'
        $r.Student | Should -BeExactly 'Adams, Amy'
        $r.Mission | Should -BeExactly 'FT5103'
        $r.Callsign | Should -BeExactly 'HAZE'
        $r.Num | Should -BeExactly '11'
        $r.Notes | Should -BeExactly 'Wx'
    }

    It 'a sim line with no IP gets SIM IP; a named sim IP is kept' {
        $script:ById[$script:SOft.Id].Instructor | Should -BeExactly 'SIM IP'
        $script:ById[$script:SIft.Id].Instructor | Should -BeExactly 'Ortiz, Sam'
    }

    It 'uses the board SimInstructor setting' {
        $b = New-TB -Settings @{ SimInstructor = 'CONTRACT SIM' }
        Add-TL $b @(New-TL -At '2026-10-20 13:00' -Type 'IFT')
        @(ConvertTo-ScheduleRows -Board $b)[0].Instructor | Should -BeExactly 'CONTRACT SIM'
    }

    It 'a return row has no student or mission and keeps both links' {
        $ret = $script:ById[$script:SP.Ret.Id]
        $null -eq $ret.Student | Should -BeTrue
        $null -eq $ret.Mission | Should -BeTrue
        $ret.Instructor | Should -BeExactly 'Ortiz, Sam'
        $ret.LinkedLineId | Should -BeExactly $script:SP.OB.Id
        $script:ById[$script:SP.OB.Id].LinkedLineId | Should -BeExactly $script:SP.Ret.Id
    }

    It 'carries WingOfLineId' {
        $script:ById[$script:SWing.Id].WingOfLineId | Should -BeExactly $script:SPin.Id
    }

    It '-Days limits the rows' {
        $rows = @(ConvertTo-ScheduleRows -Board $script:SB -Days @('2026-10-21'))
        $rows.Count | Should -Be 1
        $rows[0].BoardLineId | Should -BeExactly $script:SDay2.Id
    }

    It 'an empty board gives no rows' {
        @(ConvertTo-ScheduleRows -Board (New-TB)).Count | Should -Be 0
    }
}


Describe 'ConvertFrom-ScheduleRows' {
    BeforeAll {
        function New-TR {
            param ([string]$At, [string]$Type = 'Day Local', $IP = '', $Student = $null, $Mission = $null, [string]$Notes = $null, $Callsign = $null, $Num = $null)
            $r = New-TestRow -ETD $At -SortieType $Type -Instructor '' -Student $Student -Mission $Mission -Notes $Notes
            $r.Instructor = $IP
            if ($Type -like '*OFT*' -or $Type -like '*IFT*') { if ([string]::IsNullOrWhiteSpace((Get-NameText $IP))) { $r.Instructor = 'SIM IP' } }
            $r.Callsign = $Callsign
            $r.Num = $Num
            return $r
        }
    }

    It 'turns rows into board lines' {
        $lines = @(ConvertFrom-ScheduleRows -Rows @(
            (New-TR '2026-10-20 07:30' 'Day Local' 'Jones, Bob'),
            (New-TR '2026-10-20 13:00' 'OFT')
        ))
        $lines.Count | Should -Be 2
        $lines[0].ETD | Should -Be ([datetime]'2026-10-20 07:30')
        $lines[0].ETD | Should -BeOfType [datetime]
        $lines[0].Day | Should -BeExactly '2026-10-20'
        $lines[0].SortieType | Should -BeExactly 'Day Local'
        $lines[0].Instructor | Should -BeExactly 'Jones, Bob'
        $lines[0].Source | Should -BeExactly 'Imported'
        $lines[0].Id | Should -Match '^[0-9a-f]{8}$'
        $lines[1].Instructor | Should -BeExactly ''
    }

    It 'uses -Source' {
        @(ConvertFrom-ScheduleRows -Rows @(New-TR '2026-10-20 07:30') -Source 'Schedule sheet')[0].Source | Should -BeExactly 'Schedule sheet'
    }

    It 'reads an Excel serial ETD' {
        $row = New-TR '2026-10-20 07:30'
        $row.ETD = ([datetime]'2026-10-20 07:30').ToOADate()
        $l = @(ConvertFrom-ScheduleRows -Rows @($row))[0]
        $l.ETD.ToString('yyyy-MM-dd HH:mm') | Should -BeExactly '2026-10-20 07:30'
    }

    It 'skips rows with no usable ETD or no sortie type' {
        $bad = New-TR '2026-10-20 07:30'
        $bad.ETD = 'not a date'
        $blankType = New-TR '2026-10-20 08:00'
        $blankType.SortieType = ''
        $lines = @(ConvertFrom-ScheduleRows -Rows @($bad, $blankType, (New-TR '2026-10-20 09:00')))
        $lines.Count | Should -Be 1
        $lines[0].ETD | Should -Be ([datetime]'2026-10-20 09:00')
    }

    It 'drops SIM IP and NEED IP and reads instructor objects' {
        $obj = $script:IPs | Where-Object { $_.Name -eq 'Lee, Ann' }
        $lines = @(ConvertFrom-ScheduleRows -Rows @(
            (New-TR '2026-10-20 13:00' 'IFT' 'SIM IP'),
            (New-TR '2026-10-20 07:30' 'Day Local' 'NEED IP'),
            (New-TR '2026-10-20 08:00' 'Day Local' $obj)
        ))
        $lines[0].Instructor | Should -BeExactly ''
        $lines[1].Instructor | Should -BeExactly ''
        $lines[2].Instructor | Should -BeExactly 'Lee, Ann'
    }

    It 'drops students, missions, notes, callsign and num by default' {
        $l = @(ConvertFrom-ScheduleRows -Rows @(New-TR '2026-10-20 08:00' 'Day Local' 'Jones, Bob' 'Adams, Amy' 'FT5103' 'note' 'HAZE' 11))[0]
        $l.Student | Should -BeExactly ''
        $l.Mission | Should -BeExactly ''
        $l.Notes | Should -BeExactly ''
        $l.Callsign | Should -BeExactly ''
        $l.Num | Should -BeExactly ''
        $l.Instructor | Should -BeExactly 'Jones, Bob'
    }

    It '-KeepStudents keeps students, missions, callsign and num (strings or objects)' {
        $stu = $script:Stu[1]
        $msn = $stu.Syllabus | Where-Object { $_.Mission -eq 'FT5001' }
        $lines = @(ConvertFrom-ScheduleRows -KeepStudents -Rows @(
            (New-TR '2026-10-20 08:00' 'Day Local' 'Jones, Bob' 'Adams, Amy' 'FT5103' 'note' 'HAZE' 11),
            (New-TR '2026-10-20 09:00' 'Day Local' 'Lee, Ann' $stu $msn)
        ))
        $lines[0].Student | Should -BeExactly 'Adams, Amy'
        $lines[0].Mission | Should -BeExactly 'FT5103'
        $lines[0].Callsign | Should -BeExactly 'HAZE'
        $lines[0].Num | Should -BeExactly '11'
        $lines[0].Notes | Should -BeExactly ''
        $lines[1].Student | Should -BeExactly 'Baker, Ben'
        $lines[1].Mission | Should -BeExactly 'FT5001'
    }

    It '-KeepStudents drops the engine placeholder <S>' -TestCases @(
        @{ S = 'For Sale' }
        @{ S = 'Need Form SP' }
        @{ S = 'NEED IP' }
        @{ S = 'SIM IP' }
    ) {
        $l = @(ConvertFrom-ScheduleRows -KeepStudents -Rows @(New-TR '2026-10-20 08:00' 'Day Local' 'Jones, Bob' $S 'FT5103'))[0]
        $l.Student | Should -BeExactly ''
        $l.Mission | Should -BeExactly ''
    }

    It '-KeepStudents never puts a student on a return line' {
        $lines = @(ConvertFrom-ScheduleRows -KeepStudents -Rows @(
            (New-TR '2026-10-20 10:15' 'O&B' 'Jones, Bob' 'Adams, Amy' 'FT5102'),
            (New-TR '2026-10-20 13:15' 'O&B Return' 'Jones, Bob' 'Adams, Amy' 'FT5102')
        ))
        $lines[0].Student | Should -BeExactly 'Adams, Amy'
        $lines[1].Student | Should -BeExactly ''
        $lines[1].Mission | Should -BeExactly ''
    }

    It '-KeepNotes keeps the notes' {
        $l = @(ConvertFrom-ScheduleRows -KeepNotes -Rows @(New-TR '2026-10-20 08:00' 'Day Local' 'Jones, Bob' $null $null 'Wx brief'))[0]
        $l.Notes | Should -BeExactly 'Wx brief'
    }

    It '-NoInstructors drops the IP, callsign and num' {
        $l = @(ConvertFrom-ScheduleRows -NoInstructors -KeepStudents -Rows @(New-TR '2026-10-20 08:00' 'Day Local' 'Jones, Bob' 'Adams, Amy' $null $null 'HAZE' 11))[0]
        $l.Instructor | Should -BeExactly ''
        $l.Callsign | Should -BeExactly ''
        $l.Num | Should -BeExactly ''
        $l.Student | Should -BeExactly 'Adams, Amy'
    }

    It '-TargetDay keeps the times and moves the lines to that day' {
        $lines = @(ConvertFrom-ScheduleRows -TargetDay '2026-10-26' -Rows @(
            (New-TR '2026-10-23 07:30'),
            (New-TR '2026-10-23 13:05' 'OFT')
        ))
        $lines[0].ETD | Should -Be ([datetime]'2026-10-26 07:30')
        $lines[0].Day | Should -BeExactly '2026-10-26'
        $lines[1].ETD | Should -Be ([datetime]'2026-10-26 13:05')
    }

    It 'links each O&B to the next return that day with the same IP' {
        $lines = @(ConvertFrom-ScheduleRows -Rows @(
            (New-TR '2026-10-20 08:00' 'O&B' 'Jones, Bob'),
            (New-TR '2026-10-20 09:00' 'O&B' 'Lee, Ann'),
            (New-TR '2026-10-20 11:00' 'O&B Return' 'Lee, Ann'),
            (New-TR '2026-10-20 12:00' 'O&B Return' 'Jones, Bob')
        ))
        $lines[0].LinkedLineId | Should -BeExactly $lines[3].Id
        $lines[3].LinkedLineId | Should -BeExactly $lines[0].Id
        $lines[1].LinkedLineId | Should -BeExactly $lines[2].Id
        $lines[2].LinkedLineId | Should -BeExactly $lines[1].Id
    }

    It 'a blank IP links to the next unlinked return; earlier or other-day returns are skipped' {
        $lines = @(ConvertFrom-ScheduleRows -Rows @(
            (New-TR '2026-10-20 07:00' 'O&B Return'),
            (New-TR '2026-10-20 08:00' 'O&B'),
            (New-TR '2026-10-21 09:00' 'O&B Return'),
            (New-TR '2026-10-21 06:00' 'O&B'),
            (New-TR '2026-10-21 05:00' 'O&B')
        ))
        $lines[0].LinkedLineId | Should -BeExactly ''
        $lines[1].LinkedLineId | Should -BeExactly ''
        $lines[2].LinkedLineId | Should -BeExactly $lines[4].Id
        $lines[4].LinkedLineId | Should -BeExactly $lines[2].Id
        $lines[3].LinkedLineId | Should -BeExactly ''
    }

    It 'rows from the board keep each O&B with its own return without the IPs' {
        # The 0800 O&B comes back at 1300 and the 0900 O&B at 1200. The rows carry
        # BoardLineId / LinkedLineId, but -NoInstructors (copy previous day's default)
        # pairs by order only and swaps the returns.
        $b = New-TB
        $p1 = Add-TOB $b '2026-10-20 08:00' 'Jones, Bob' '2026-10-20 13:00'
        $p2 = Add-TOB $b '2026-10-20 09:00' 'Lee, Ann' '2026-10-20 12:00'
        $rows = @(ConvertTo-ScheduleRows -Board $b)
        $lines = @(ConvertFrom-ScheduleRows -Rows $rows -NoInstructors)
        $ob0800 = $lines | Where-Object { $_.SortieType -eq 'O&B' -and $_.ETD.Hour -eq 8 }
        $ret = $lines | Where-Object { $_.Id -eq $ob0800.LinkedLineId }
        $ret.ETD | Should -Be ([datetime]'2026-10-20 13:00')
    }
}


Describe 'Schedule sheet notes' {
    # The sheet loads through ConvertFrom-SheetRows (Tests\OneDay.Tests.ps1 covers the per-day load)
    It 'drops the fill''s "Clears Unsat" notes from the sheet rows but keeps the others' {
        $row = [PSCustomObject]@{ ETD = '2026-10-20 07:30:00'; Day = '2026-10-20'; Time = '07:30'; TimeOnly = $false; Instructor = $null; SortieType = 'OFT'
            Student = 'Adams, Amy'; Mission = $null; Callsign = $null; Num = $null; Notes = 'Bring a jet | Clears Unsat Sim (OFT) | OVERRIDE: x'; SheetRow = 2 }
        $l = @(ConvertFrom-SheetRows -Rows @($row) -Day '2026-10-20' -KeepStudents -KeepNotes)
        $l[0].Notes | Should -BeExactly 'Bring a jet | OVERRIDE: x'
        $l[0].Student | Should -BeExactly 'Adams, Amy'
    }
}


Describe 'Save-BoardState / Open-BoardState' {
    BeforeAll {
        $script:Dir = Join-Path $TestDrive 'boards'
        $script:OB = New-TB -Days @('2026-10-20', '2026-10-21', '2026-10-22')
        $script:OLead = New-TL -At '2026-10-20 08:00' -IP 'Jones, Bob' -Student 'Adams, Amy' -Mission 'FF5201' -Notes 'Brief "early" | 0700' -Callsign 'HAZE' -Num '11'
        $script:OWing = New-TL -At '2026-10-20 08:00' -IP 'Lee, Ann' -Student 'Baker, Ben'
        $script:OWing.WingOfLineId = $script:OLead.Id
        $script:OOft = New-TL -At '2026-10-20 13:00' -Type 'OFT'
        $script:OOft.Source = 'Template'
        Add-TL $script:OB @($script:OLead, $script:OWing, $script:OOft)
        $script:OP = Add-TOB $script:OB '2026-10-20 10:15' 'Ortiz, Sam' '2026-10-20 13:40'
        $script:OSingle = New-TL -At '2026-10-21 07:45' -IP 'Leech, Tom'
        Add-TL $script:OB @($script:OSingle)
        Save-BoardState -Board $script:OB -Dir $script:Dir
        $script:Re = Open-BoardState -Dir $script:Dir -Days @('2026-10-20', '2026-10-21', '2026-10-22')
    }

    It 'writes one yyyy-MM-dd.json per day and no temp files' {
        $names = @(Get-ChildItem -LiteralPath $script:Dir | ForEach-Object { $_.Name } | Sort-Object)
        ($names -join ',') | Should -BeExactly '2026-10-20.json,2026-10-21.json,2026-10-22.json'
    }

    It 'stores dates as text and a single-line day as an array' {
        $raw = [System.IO.File]::ReadAllText((Join-Path $script:Dir '2026-10-21.json'))
        $raw | Should -Match '"Lines":\s*\['
        $raw | Should -Match '"ETD":\s*"2026-10-21T07:45"'
        $raw | Should -Match '"Day":\s*"2026-10-21"'
        $empty = [System.IO.File]::ReadAllText((Join-Path $script:Dir '2026-10-22.json'))
        $empty | Should -Match '"Lines":\s*\['
    }

    It 'opens the same days with the same lines' {
        ($script:Re.Days.ToArray() -join ',') | Should -BeExactly '2026-10-20,2026-10-21,2026-10-22'
        $script:Re.Lines.Count | Should -Be 6
        @(Get-BoardDayLines -Board $script:Re -Day '2026-10-21').Count | Should -Be 1
        @(Get-BoardDayLines -Board $script:Re -Day '2026-10-22').Count | Should -Be 0
        $script:Re.Version | Should -Be 0
        $script:Re.Undo.Count | Should -Be 0
    }

    It 'round-trips every field of every line exactly' {
        foreach ($orig in $script:OB.Lines) {
            $back = Get-BoardLine $script:Re $orig.Id
            $null -ne $back | Should -BeTrue
            $back.ETD | Should -BeOfType [datetime]
            $back.ETD | Should -Be $orig.ETD
            foreach ($p in @('Id', 'Day', 'SortieType', 'Instructor', 'Student', 'Mission', 'Callsign', 'Num', 'Notes', 'LinkedLineId', 'WingOfLineId', 'Source')) {
                [string]$back.$p | Should -BeExactly ([string]$orig.$p) -Because "$p of $($orig.Id)"
            }
        }
    }

    It 'keeps the O&B link and the wing pairing' {
        (Get-BoardLine $script:Re $script:OP.OB.Id).LinkedLineId | Should -BeExactly $script:OP.Ret.Id
        (Get-BoardLine $script:Re $script:OP.Ret.Id).LinkedLineId | Should -BeExactly $script:OP.OB.Id
        (Get-BoardLine $script:Re $script:OP.Ret.Id).ETD | Should -Be ([datetime]'2026-10-20 13:40')
        (Get-BoardLine $script:Re $script:OWing.Id).WingOfLineId | Should -BeExactly $script:OLead.Id
        (Get-BoardLine $script:Re $script:OLead.Id).Notes | Should -BeExactly 'Brief "early" | 0700'
    }

    It 'the reopened board still edits: the O&B drags its return' {
        Set-BoardLineEtd -Board $script:Re -Id $script:OP.OB.Id -ETD ([datetime]'2026-10-20 10:45')
        (Get-BoardLine $script:Re $script:OP.Ret.Id).ETD | Should -Be ([datetime]'2026-10-20 14:10')
    }

    It 'a day with no file opens empty' {
        $b = Open-BoardState -Dir $script:Dir -Days @('2026-10-28')
        ($b.Days.ToArray() -join ',') | Should -BeExactly '2026-10-28'
        $b.Lines.Count | Should -Be 0
    }

    It 'a damaged file opens empty with a warning instead of throwing' {
        $dir = Join-Path $TestDrive 'damaged'
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        [System.IO.File]::WriteAllText((Join-Path $dir '2026-10-20.json'), '{ not json')
        $out = @(Open-BoardState -Dir $dir -Days @('2026-10-20') 3>&1)
        $warn = @($out | Where-Object { $_ -is [System.Management.Automation.WarningRecord] })
        $b = @($out | Where-Object { $_ -isnot [System.Management.Automation.WarningRecord] })[0]
        $warn.Count | Should -Be 1
        $warn[0].Message | Should -BeLike "Could not read saved board $(Join-Path $dir '2026-10-20.json') : *"
        ($b.Days.ToArray() -join ',') | Should -BeExactly '2026-10-20'
        $b.Lines.Count | Should -Be 0
    }

    It 'saving again overwrites the day' {
        $dir = Join-Path $TestDrive 'resave'
        $b = New-TB
        $l1 = New-TL -At '2026-10-20 07:30'
        $l2 = New-TL -At '2026-10-20 08:30'
        Add-TL $b @($l1, $l2)
        Save-BoardState -Board $b -Dir $dir
        Remove-BoardLine -Board $b -Id $l1.Id
        Save-BoardState -Board $b -Dir $dir
        $re = Open-BoardState -Dir $dir -Days @('2026-10-20')
        $re.Lines.Count | Should -Be 1
        $re.Lines[0].Id | Should -BeExactly $l2.Id
    }

    It 'Get-SavedBoardDays gives nothing for a missing folder' {
        @(Get-SavedBoardDays -Dir (Join-Path $TestDrive 'missing')).Count | Should -Be 0
    }

    It 'Get-SavedBoardDays lists the saved days on or after -From' {
        # [datetime]::TryParseExact(..., [ref]$d) with $d = $null finds no overload on
        # PowerShell 7 (span overloads), so this throws as soon as a board file exists.
        # Show-PuckBoard calls it on every start.
        $dir = Join-Path $TestDrive 'saveddays'
        Save-BoardState -Board $script:OB -Dir $dir
        [System.IO.File]::WriteAllText((Join-Path $dir 'notes.json'), '{}')
        (@(Get-SavedBoardDays -Dir $dir) -join ',') | Should -BeExactly '2026-10-20,2026-10-21,2026-10-22'
        (@(Get-SavedBoardDays -Dir $dir -From ([datetime]'2026-10-21')) -join ',') | Should -BeExactly '2026-10-21,2026-10-22'
    }
}


Describe 'Board templates' {
    BeforeAll {
        $script:TDir = Join-Path $TestDrive 'templates'
        $script:TB = New-TB
        $script:TDl = New-TL -At '2026-10-20 07:30' -IP 'Jones, Bob' -Student 'Adams, Amy' -Mission 'FT5103' -Notes 'x'
        $script:TOft = New-TL -At '2026-10-20 13:00' -Type 'OFT'
        $script:TLoose = New-TL -At '2026-10-20 16:00' -Type 'O&B Return' -IP 'Leech, Tom'
        Add-TL $script:TB @($script:TDl, $script:TOft, $script:TLoose)
        $script:TP = Add-TOB $script:TB '2026-10-20 10:15' 'Lee, Ann' '2026-10-20 14:00'
        $script:TOpen = New-TL -At '2026-10-20 11:00' -Type 'O&B'
        Add-TL $script:TB @($script:TOpen)
        $script:TPath = Save-BoardTemplate -Board $script:TB -Day '2026-10-20' -Name 'AM wave!' -Dir $script:TDir
        $script:TPathIP = Save-BoardTemplate -Board $script:TB -Day '2026-10-20' -Name 'AM wave IPs' -Dir $script:TDir -IncludeInstructors
        $script:TJson = ConvertFrom-Json ([System.IO.File]::ReadAllText($script:TPath))
        $script:TJsonIP = ConvertFrom-Json ([System.IO.File]::ReadAllText($script:TPathIP))
    }

    It 'saves to the template dir under a file-safe name and keeps the typed name' {
        $script:TPath | Should -BeExactly (Join-Path $script:TDir 'AM wave.json')
        Test-Path -LiteralPath $script:TPath | Should -BeTrue
        $script:TJson.Name | Should -BeExactly 'AM wave!'
        $script:TJson.SchemaVersion | Should -Be 1
    }

    It 'stores times and types; the linked return lives inside its O&B' {
        $items = @($script:TJson.Lines)
        $items.Count | Should -Be 5
        (@($items | ForEach-Object { $_.Time + ' ' + $_.SortieType + ' ' + $_.ReturnTime }) -join '|') | Should -BeExactly '07:30 Day Local |10:15 O&B 14:00|11:00 O&B |13:00 OFT |16:00 O&B Return '
    }

    It 'leaves the IPs out unless -IncludeInstructors' {
        foreach ($t in @($script:TJson.Lines)) { [string]$t.Instructor | Should -BeExactly '' }
        $ip = @($script:TJsonIP.Lines)
        $ip[0].Instructor | Should -BeExactly 'Jones, Bob'
        $ip[1].Instructor | Should -BeExactly 'Lee, Ann'
        $ip[4].Instructor | Should -BeExactly 'Leech, Tom'
    }

    It 'never stores students, missions or notes' {
        $raw = [System.IO.File]::ReadAllText($script:TPathIP)
        $raw | Should -Not -Match 'Adams'
        $raw | Should -Not -Match 'FT5103'
    }

    It 'throws on a name with no usable characters' {
        { Save-BoardTemplate -Board $script:TB -Day '2026-10-20' -Name '!!!' -Dir $script:TDir } | Should -Throw -ExpectedMessage "Template name '!!!' has no usable characters"
    }

    It 'Get-BoardTemplateList lists name, path and line count, sorted, skipping bad files' {
        $dir = Join-Path $TestDrive 'tlist'
        [void](Save-BoardTemplate -Board $script:TB -Day '2026-10-20' -Name 'Zulu' -Dir $dir)
        $one = New-TB
        Add-TL $one @(New-TL -At '2026-10-20 07:30')
        [void](Save-BoardTemplate -Board $one -Day '2026-10-20' -Name 'Alpha' -Dir $dir)
        [System.IO.File]::WriteAllText((Join-Path $dir 'broken.json'), '{ nope')
        $list = @(Get-BoardTemplateList -Dir $dir)
        $list.Count | Should -Be 2
        $list[0].Name | Should -BeExactly 'Alpha'
        $list[0].Count | Should -Be 1
        $list[0].Path | Should -BeExactly (Join-Path $dir 'Alpha.json')
        $list[1].Name | Should -BeExactly 'Zulu'
        $list[1].Count | Should -Be 5
        @(Get-BoardTemplateList -Dir (Join-Path $TestDrive 'none')).Count | Should -Be 0
    }

    It 'Import-BoardTemplate adds the lines to the day with their returns, in one undo step' {
        $b = New-TB
        $v = $b.Version
        $added = Import-BoardTemplate -Board $b -Day '2026-10-21' -Path $script:TPathIP
        $added | Should -Be 7
        ($b.Days.ToArray() -join ',') | Should -BeExactly '2026-10-20,2026-10-21'
        $lines = @(Get-BoardDayLines -Board $b -Day '2026-10-21')
        $lines.Count | Should -Be 7
        foreach ($l in $lines) { $l.Source | Should -BeExactly 'Template'; $l.Student | Should -BeExactly '' }
        $ob = $lines | Where-Object { $_.SortieType -eq 'O&B' -and $_.ETD.Hour -eq 10 }
        $ob.ETD | Should -Be ([datetime]'2026-10-21 10:15')
        $ob.Instructor | Should -BeExactly 'Lee, Ann'
        $ret = Get-BoardLine $b $ob.LinkedLineId
        $ret.SortieType | Should -BeExactly 'O&B Return'
        $ret.ETD | Should -Be ([datetime]'2026-10-21 14:00')
        $ret.Instructor | Should -BeExactly 'Lee, Ann'
        $ret.LinkedLineId | Should -BeExactly $ob.Id
        $open = $lines | Where-Object { $_.SortieType -eq 'O&B' -and $_.ETD.Hour -eq 11 }
        (Get-BoardLine $b $open.LinkedLineId).ETD | Should -Be ([datetime]'2026-10-21 14:00')
        $loose = $lines | Where-Object { $_.SortieType -eq 'O&B Return' -and $_.ETD.Hour -eq 16 }
        $loose.LinkedLineId | Should -BeExactly ''
        $b.Version | Should -Be ($v + 1)
        $b.Undo.Count | Should -Be 1
        (Undo-Board $b) | Should -BeTrue
        $b.Lines.Count | Should -Be 0
    }

    It 'a template saved without IPs imports with blank IPs' {
        $b = New-TB
        [void](Import-BoardTemplate -Board $b -Day '2026-10-20' -Path $script:TPath)
        foreach ($l in $b.Lines) { $l.Instructor | Should -BeExactly '' }
    }

    It 'without -Replace the lines are added; -Replace clears only that day first' {
        $b = New-TB -Days @('2026-10-20', '2026-10-21')
        $keep = New-TL -At '2026-10-20 06:00'
        $old = New-TL -At '2026-10-21 06:00'
        Add-TL $b @($keep, $old)
        [void](Import-BoardTemplate -Board $b -Day '2026-10-21' -Path $script:TPath)
        @(Get-BoardDayLines -Board $b -Day '2026-10-21').Count | Should -Be 8
        $added = Import-BoardTemplate -Board $b -Day '2026-10-21' -Path $script:TPath -Replace
        $added | Should -Be 7
        @(Get-BoardDayLines -Board $b -Day '2026-10-21').Count | Should -Be 7
        $null -eq (Get-BoardLine $b $old.Id) | Should -BeTrue
        $null -ne (Get-BoardLine $b $keep.Id) | Should -BeTrue
        (Undo-Board $b) | Should -BeTrue
        @(Get-BoardDayLines -Board $b -Day '2026-10-21').Count | Should -Be 8
        $null -ne (Get-BoardLine $b $old.Id) | Should -BeTrue
    }
}


Describe 'Copy-BoardPreviousDay' {
    BeforeAll {
        $script:Fri = '2026-10-23'
        $script:Mon = '2026-10-26'

        function Add-TFriday {
            param ($Board)
            $dl = New-TL -At '2026-10-23 07:30' -IP 'Jones, Bob' -Student 'Adams, Amy' -Mission 'FT5103' -Notes 'Wx' -Callsign 'HAZE' -Num '11'
            $oft = New-TL -At '2026-10-23 13:00' -Type 'OFT'
            Add-TL $Board @($dl, $oft)
            [void](Add-TOB $Board '2026-10-23 10:15' 'Lee, Ann' '2026-10-23 14:00')
        }
    }

    It 'copies Friday from the board onto Monday: times, types, links, no students' {
        $b = New-TB -Days @($script:Fri, $script:Mon)
        Add-TFriday $b
        $fridayBefore = @(Get-BoardDayLines -Board $b -Day $script:Fri | ForEach-Object { ConvertTo-Json (ConvertTo-LineDto $_) -Compress }) -join '|'
        $res = Copy-BoardPreviousDay -Board $b -TargetDay $script:Mon
        $res.Source | Should -BeExactly 'board 2026-10-23'
        $res.Count | Should -Be 4
        $mon = @(Get-BoardDayLines -Board $b -Day $script:Mon)
        $mon.Count | Should -Be 4
        (@($mon | ForEach-Object { $_.ETD.ToString('yyyy-MM-dd HHmm') + ' ' + $_.SortieType }) -join '|') | Should -BeExactly '2026-10-26 0730 Day Local|2026-10-26 1015 O&B|2026-10-26 1300 OFT|2026-10-26 1400 O&B Return'
        foreach ($l in $mon) {
            $l.Student | Should -BeExactly ''
            $l.Mission | Should -BeExactly ''
            $l.Notes | Should -BeExactly ''
            $l.Callsign | Should -BeExactly ''
            $l.Instructor | Should -BeExactly ''
            $l.Source | Should -BeExactly 'Copied'
        }
        $ob = $mon[1]
        $ob.LinkedLineId | Should -BeExactly $mon[3].Id
        $mon[3].LinkedLineId | Should -BeExactly $ob.Id
        $fridayAfter = @(Get-BoardDayLines -Board $b -Day $script:Fri | ForEach-Object { ConvertTo-Json (ConvertTo-LineDto $_) -Compress }) -join '|'
        $fridayAfter | Should -BeExactly $fridayBefore
    }

    It '-IncludeInstructors copies the IPs (the sim stays blank)' {
        $b = New-TB -Days @($script:Fri, $script:Mon)
        Add-TFriday $b
        [void](Copy-BoardPreviousDay -Board $b -TargetDay $script:Mon -IncludeInstructors)
        $mon = @(Get-BoardDayLines -Board $b -Day $script:Mon)
        (@($mon | ForEach-Object { $_.Instructor }) -join '|') | Should -BeExactly 'Jones, Bob|Lee, Ann||Lee, Ann'
        $mon[0].Student | Should -BeExactly ''
        $mon[1].LinkedLineId | Should -BeExactly $mon[3].Id
    }

    It 'is one undo step and adds the target day' {
        $b = New-TB -Days @($script:Fri)
        Add-TFriday $b
        $u = $b.Undo.Count
        [void](Copy-BoardPreviousDay -Board $b -TargetDay $script:Mon)
        $b.Days.Contains($script:Mon) | Should -BeTrue
        $b.Undo.Count | Should -Be ($u + 1)
        (Undo-Board $b) | Should -BeTrue
        @(Get-BoardDayLines -Board $b -Day $script:Mon).Count | Should -Be 0
        $b.Lines.Count | Should -Be 4
    }

    It 'a saved day with a single line copies as one line' {
        $dir = Join-Path $TestDrive 'prevday-single'
        $fri = New-TB -Days @($script:Fri)
        Add-TL $fri @(New-TL -At '2026-10-23 09:30' -IP 'Ortiz, Sam')
        Save-BoardState -Board $fri -Dir $dir
        $b = New-TB -Days @($script:Mon)
        $res = Copy-BoardPreviousDay -Board $b -TargetDay $script:Mon -Dir $dir
        $res.Source | Should -BeExactly 'saved board 2026-10-23'
        $res.Count | Should -Be 1
        $b.Lines.Count | Should -Be 1
        $b.Lines[0].ETD | Should -Be ([datetime]'2026-10-26 09:30')
    }

    It 'the nearest earlier day wins' {
        $b = New-TB -Days @('2026-10-22', $script:Fri, $script:Mon)
        Add-TL $b @(New-TL -At '2026-10-22 06:00' -Type 'IFT')
        Add-TL $b @(New-TL -At '2026-10-23 09:00')
        $res = Copy-BoardPreviousDay -Board $b -TargetDay $script:Mon
        $res.Source | Should -BeExactly 'board 2026-10-23'
        $res.Count | Should -Be 1
        @(Get-BoardDayLines -Board $b -Day $script:Mon)[0].ETD | Should -Be ([datetime]'2026-10-26 09:00')
    }

    It 'copies Friday from a saved board file' {
        $dir = Join-Path $TestDrive 'prevday'
        $fri = New-TB -Days @($script:Fri)
        Add-TFriday $fri
        Save-BoardState -Board $fri -Dir $dir
        $b = New-TB -Days @($script:Mon)
        $res = Copy-BoardPreviousDay -Board $b -TargetDay $script:Mon -Dir $dir -IncludeInstructors
        $res.Source | Should -BeExactly 'saved board 2026-10-23'
        $res.Count | Should -Be 4
        $mon = @(Get-BoardDayLines -Board $b -Day $script:Mon)
        (@($mon | ForEach-Object { $_.ETD.ToString('HHmm') + ' ' + $_.SortieType + ' ' + $_.Instructor + ' ' + $_.Student }) -join '|') | Should -BeExactly '0730 Day Local Jones, Bob |1015 O&B Lee, Ann |1300 OFT  |1400 O&B Return Lee, Ann '
        $mon[1].LinkedLineId | Should -BeExactly $mon[3].Id
        @($b.Lines | Where-Object { $_.Day -eq $script:Fri }).Count | Should -Be 0
    }

    It 'copies the latest earlier day of the previous MasterSchedule' {
        $jones = $script:IPs | Where-Object { $_.Name -eq 'Jones, Bob' }
        $lee = $script:IPs | Where-Object { $_.Name -eq 'Lee, Ann' }
        $mk = {
            param ($At, $Type, $IP, $Student)
            $r = New-TestRow -ETD $At -SortieType $Type -Student $Student -Mission 'FT5103'
            if ($null -ne $IP) { $r.Instructor = $IP }
            $r
        }
        $prev = @(
            (& $mk '2026-10-22 07:00' 'Day Local' $jones $script:Stu[0]),
            (& $mk '2026-10-23 07:30' 'Day Local' $jones $script:Stu[0]),
            (& $mk '2026-10-23 10:15' 'O&B' $lee $script:Stu[1]),
            (& $mk '2026-10-23 13:15' 'O&B Return' $lee $script:Stu[1]),
            (& $mk '2026-10-23 13:00' 'OFT' $null 'For Sale'),
            (& $mk '2026-10-26 08:00' 'Day Local' $jones $script:Stu[2])
        )
        $b = New-TB -Days @($script:Mon)
        $res = Copy-BoardPreviousDay -Board $b -TargetDay $script:Mon -PrevSchedule $prev -IncludeInstructors
        $res.Source | Should -BeExactly 'MasterSchedule 2026-10-23'
        $res.Count | Should -Be 4
        $mon = @(Get-BoardDayLines -Board $b -Day $script:Mon)
        (@($mon | ForEach-Object { $_.ETD.ToString('HHmm') + ' ' + $_.SortieType + ' ' + $_.Instructor }) -join '|') | Should -BeExactly '0730 Day Local Jones, Bob|1015 O&B Lee, Ann|1300 OFT |1315 O&B Return Lee, Ann'
        foreach ($l in $mon) { $l.Student | Should -BeExactly '' }
        $mon[1].LinkedLineId | Should -BeExactly $mon[3].Id
    }

    It 'the board wins over the previous MasterSchedule' {
        $b = New-TB -Days @($script:Fri, $script:Mon)
        Add-TFriday $b
        $prev = @(New-TestRow -ETD '2026-10-23 06:00' -SortieType 'IFT')
        $res = Copy-BoardPreviousDay -Board $b -TargetDay $script:Mon -PrevSchedule $prev
        $res.Source | Should -BeExactly 'board 2026-10-23'
    }

    It 'returns $null and changes nothing when there is nothing within 7 days' {
        $b = New-TB -Days @('2026-10-18', $script:Mon)
        Add-TL $b @(New-TL -At '2026-10-18 08:00')
        $v = $b.Version
        $res = Copy-BoardPreviousDay -Board $b -TargetDay $script:Mon -Dir (Join-Path $TestDrive 'nothing-here')
        $null -eq $res | Should -BeTrue
        $b.Version | Should -Be $v
        $b.Lines.Count | Should -Be 1
    }

    It 'ignores MasterSchedule rows on or after the target day' {
        $b = New-TB -Days @($script:Mon)
        $prev = @(New-TestRow -ETD '2026-10-26 06:00' -SortieType 'IFT'; New-TestRow -ETD '2026-10-27 06:00')
        $null -eq (Copy-BoardPreviousDay -Board $b -TargetDay $script:Mon -PrevSchedule $prev) | Should -BeTrue
        $b.Lines.Count | Should -Be 0
    }
}


Describe 'Get-FillOverlay' {
    It 'maps each board line id to what the fill put there' {
        $jones = $script:IPs | Where-Object { $_.Name -eq 'Jones, Bob' }
        $stu = $script:Stu[0]
        $msn = $stu.Syllabus | Where-Object { $_.Mission -eq 'FT5103' }
        $r1 = New-TestRow -ETD '2026-10-20 08:00' -Id 'aaaa0001' -Student $stu -Mission $msn
        $r1.Instructor = $jones
        $r1.FilledBy = 'Engine'
        $r1.Notes = 'Continue String'
        $r1.Callsign = 'TALON'
        $r1.Num = 31
        $r2 = New-TestRow -ETD '2026-10-20 09:00' -Instructor 'Lee, Ann' -Id 'aaaa0002' -Student 'Baker, Ben' -Mission 'FT5001' -Notes 'OVERRIDE: Turn time 2.8 hr < 3.25 hr minimum (vs 0730 FT5103)'
        $r3 = New-TestRow -ETD '2026-10-20 10:00' -Id 'aaaa0003' -Student 'For Sale'
        $r3.FilledBy = 'ForSale'
        $r4 = New-TestRow -ETD '2026-10-20 11:00' -Student 'Cruz, Cal'
        $map = Get-FillOverlay -FillResult ([PSCustomObject]@{ Schedule = @($r1, $r2, $r3, $r4) })
        $map | Should -BeOfType [hashtable]
        $map.Count | Should -Be 3
        $a = $map['aaaa0001']
        $a.Student | Should -BeExactly 'Adams, Amy'
        $a.Mission | Should -BeExactly 'FT5103'
        $a.Instructor | Should -BeExactly 'Jones, Bob'
        $a.FilledBy | Should -BeExactly 'Engine'
        $a.Notes | Should -BeExactly 'Continue String'
        $a.Callsign | Should -BeExactly 'TALON'
        $a.Num | Should -BeExactly '31'
        $a.Override | Should -BeFalse
        $map['aaaa0002'].Override | Should -BeTrue
        $map['aaaa0002'].Student | Should -BeExactly 'Baker, Ben'
        $map['aaaa0002'].FilledBy | Should -BeExactly ''
        $map['aaaa0003'].FilledBy | Should -BeExactly 'ForSale'
        $map['aaaa0003'].Student | Should -BeExactly 'For Sale'
    }

    It 'an empty fill gives an empty map' {
        $map = Get-FillOverlay -FillResult ([PSCustomObject]@{ Schedule = @() })
        $map.Count | Should -Be 0
    }
}


Describe 'Set-BoardFill / Clear-BoardFill' {
    BeforeEach {
        $script:FB = New-TB -Days @($script:D1, $script:D2)
        $script:FPin = New-TL -At '2026-10-20 07:30' -IP 'Jones, Bob' -Student 'Cruz, Cal'
        $script:FLead = New-TL -At '2026-10-20 08:00' -IP 'Jones, Bob'
        $script:FWing = New-TL -At '2026-10-20 08:10' -IP 'Lee, Ann'
        $script:FOpen = New-TL -At '2026-10-20 13:00' -IP 'Lee, Ann'
        $script:FWed = New-TL -At '2026-10-21 08:00' -IP 'Lee, Ann'
        Add-TL $script:FB @($script:FPin, $script:FLead, $script:FWing, $script:FOpen, $script:FWed)
        $stu = $script:Stu[0]
        $msn = $stu.Syllabus | Where-Object { $_.Mission -eq 'FT5103' }
        $pin = New-TestRow -ETD '2026-10-20 07:30' -Id $script:FPin.Id -Student 'Cruz, Cal'; $pin.FilledBy = 'Pin'
        $lead = New-TestRow -ETD '2026-10-20 08:00' -Id $script:FLead.Id -Student $stu -Mission $msn -Notes 'Wingman: Baker, Ben'; $lead.FilledBy = 'Engine'
        $wing = New-TestRow -ETD '2026-10-20 08:10' -Id $script:FWing.Id -Student 'Baker, Ben' -Mission 'FF5201' -Notes 'Wingman to Adams, Amy'; $wing.FilledBy = 'Wing'
        $sale = New-TestRow -ETD '2026-10-20 13:00' -Id $script:FOpen.Id -Student 'For Sale'; $sale.FilledBy = 'ForSale'
        $script:FFill = [PSCustomObject]@{ Schedule = @($pin, $lead, $wing, $sale) }
    }

    It "puts the scheduler's picks on the board, marked, with the wingman paired to its lead" {
        Set-BoardFill -Board $script:FB -FillResult $script:FFill -Days @($script:D1) | Should -Be 2
        "$($script:FLead.Student)|$($script:FLead.Mission)|$($script:FLead.Filled)" | Should -BeExactly 'Adams, Amy|FT5103|Engine'
        "$($script:FWing.Student)|$($script:FWing.Mission)|$($script:FWing.Filled)" | Should -BeExactly 'Baker, Ben|FF5201|Wing'
        $script:FWing.WingOfLineId | Should -BeExactly $script:FLead.Id
        "$($script:FOpen.Student)|$($script:FOpen.Filled)" | Should -BeExactly '|ForSale'
        "$($script:FPin.Student)|$($script:FPin.Filled)" | Should -BeExactly 'Cruz, Cal|' -Because 'a line hardcoded by hand stays as it is'
        $script:FWed.Filled | Should -BeExactly ''
    }

    It 'is one undo step, and Clear-BoardFill takes only the picks off' {
        $v = $script:FB.Version
        [void](Set-BoardFill -Board $script:FB -FillResult $script:FFill -Days @($script:D1))
        $script:FB.Version | Should -BeGreaterThan $v
        Clear-BoardFill -Board $script:FB -Days @($script:D1) | Should -Be 2
        "$($script:FLead.Student)|$($script:FLead.Filled)|$($script:FWing.Student)|$($script:FWing.WingOfLineId)|$($script:FOpen.Filled)" | Should -BeExactly '||||'
        $script:FPin.Student | Should -BeExactly 'Cruz, Cal'
        Clear-BoardFill -Board $script:FB -Days @($script:D1) | Should -Be 0
        [void](Undo-Board $script:FB)
        (Get-BoardLine $script:FB $script:FLead.Id).Filled | Should -BeExactly 'Engine' -Because 'Ctrl+Z puts the picks back'
        [void](Undo-Board $script:FB)
        (Get-BoardLine $script:FB $script:FLead.Id).Student | Should -BeExactly '' -Because 'and once more takes the fill off'
    }

    It 'a pick changed by hand is the user''s: it loses its mark and Clear-BoardFill leaves it' {
        [void](Set-BoardFill -Board $script:FB -FillResult $script:FFill -Days @($script:D1))
        Set-BoardMission -Board $script:FB -Id $script:FLead.Id -Mission 'FT5001'
        $script:FLead.Filled | Should -BeExactly ''
        [void](Set-BoardPin -Board $script:FB -Id $script:FOpen.Id -Student 'Cruz, Cal')
        $script:FOpen.Filled | Should -BeExactly ''
        [void](Clear-BoardFill -Board $script:FB -Days @($script:D1))
        $script:FLead.Student | Should -BeExactly 'Adams, Amy'
        $script:FWing.Student | Should -BeExactly ''
    }

    It 'the mark is saved with the board and read back' {
        [void](Set-BoardFill -Board $script:FB -FillResult $script:FFill -Days @($script:D1))
        $dir = Join-Path $TestDrive 'fillmark'
        Save-BoardState -Board $script:FB -Dir $dir
        $back = Open-BoardState -Dir $dir -Days @($script:D1, $script:D2)
        (Get-BoardLine $back $script:FLead.Id).Filled | Should -BeExactly 'Engine'
        (Get-BoardLine $back $script:FPin.Id).Filled | Should -BeExactly ''
    }
}


Describe 'Test-BoardReadyToFill' {
    It 'an empty board says so' {
        $w = @(Test-BoardReadyToFill -Board (New-TB) -Students $script:Stu -Instructors $script:IPs)
        ($w -join '|') | Should -BeExactly 'There are no lines on the board.'
    }

    It 'a clean board has no warnings (sim lines need no IP)' {
        $b = New-TB
        Add-TL $b @(
            (New-TL -At '2026-10-20 07:30' -IP 'Jones, Bob' -Student 'Adams, Amy'),
            (New-TL -At '2026-10-20 13:00' -Type 'OFT' -Student 'Baker, Ben'),
            (New-TL -At '2026-10-20 06:00' -Type 'IFT')
        )
        [void](Add-TOB $b '2026-10-20 10:15' 'Lee, Ann')
        @(Test-BoardReadyToFill -Board $b -Students $script:Stu -Instructors $script:IPs).Count | Should -Be 0
    }

    It 'lists every problem with its exact message' {
        $b = New-TB -Days @('2026-10-20', '2026-10-21')
        Add-TL $b @(
            (New-TL -At '2026-10-20 07:30' -IP 'Jones, Bob' -Student 'Adams, Amy'),
            (New-TL -At '2026-10-20 07:30' -IP 'Lee, Ann' -Student 'Adams, Amy'),
            (New-TL -At '2026-10-20 08:00'),
            (New-TL -At '2026-10-20 09:00' -IP 'Smith, Joe'),
            (New-TL -At '2026-10-20 09:30' -IP 'Ortiz, Sam' -Student 'Nobody'),
            (New-TL -At '2026-10-20 10:15' -Type 'O&B' -IP 'Leech, Tom')
        )
        $w = @(Test-BoardReadyToFill -Board $b -Students $script:Stu -Instructors $script:IPs 3>$null)
        $w | Should -Contain 'Adams, Amy is pinned twice at Tue 0730'
        $w | Should -Contain 'Tue 0800 Day Local has no IP (it will be NEED IP)'
        $w | Should -Contain "Tue 0900 Day Local IP 'Smith, Joe' is not in MasterData"
        $w | Should -Contain "Tue 0930 Day Local student 'Nobody' is not in MasterData"
        $w | Should -Contain 'Tue 1015 O&B has no linked O&B Return line'
        $w | Should -Contain '2026-10-21 has no lines'
        $w.Count | Should -Be 6
    }
}


Describe 'Get-NextFlyingDay' {
    It '<From> -> <Want>' -TestCases @(
        @{ From = '2026-10-23 15:00'; Want = '2026-10-26' }
        @{ From = '2026-10-24 09:00'; Want = '2026-10-26' }
        @{ From = '2026-10-25 23:59'; Want = '2026-10-26' }
        @{ From = '2026-10-26 06:00'; Want = '2026-10-27' }
        @{ From = '2026-10-22 00:00'; Want = '2026-10-23' }
    ) {
        $d = Get-NextFlyingDay -From ([datetime]$From)
        $d | Should -BeOfType [datetime]
        $d | Should -Be ([datetime]$Want)
        $d.TimeOfDay.Ticks | Should -Be 0
    }

    It 'defaults to the next weekday after today' {
        $d = Get-NextFlyingDay
        $d | Should -BeGreaterThan (Get-Date).Date
        @('Saturday', 'Sunday') | Should -Not -Contain ([string]$d.DayOfWeek)
    }
}

Describe 'Set-BoardMission' {
    It 'saves a change of case (the dropdown shows the line''s own spelling)' {
        $b = New-TB
        $l = New-TL -At '2026-10-20 08:00' -Mission 'ft5001'
        Add-BoardLine -Board $b -Line @($l) -NoUndo
        $v = $b.Version
        Set-BoardMission -Board $b -Id $l.Id -Mission 'FT5001'
        (Get-BoardLine $b $l.Id).Mission | Should -BeExactly 'FT5001'
        $b.Version | Should -BeGreaterThan $v
    }
}

Describe 'New-BulkBoardLines (Add Bulk Lines)' {
    It 'makes the lines that many minutes apart, 3 by default' {
        $r = New-BulkBoardLines -Day $script:D1 -StartTime '0730' -Count '10'
        @($r.Errors).Count | Should -Be 0
        ($r.Times -join ',') | Should -BeExactly '0730,0733,0736,0739,0742,0745,0748,0751,0754,0757'
        @($r.Lines).Count | Should -Be 10
        $r.Lines[0].ETD | Should -Be ([datetime]'2026-10-20 07:30')
        $r.Lines[9].ETD | Should -Be ([datetime]'2026-10-20 07:57')
        @($r.Lines | Where-Object { $_.SortieType -ne 'Day Local' -or $_.Day -ne $script:D1 -or $_.Student -or $_.Instructor }).Count | Should -Be 0
        @($r.Lines | ForEach-Object { $_.Id } | Select-Object -Unique).Count | Should -Be 10
    }

    It 'reads the boxes the way people type them' {
        ((New-BulkBoardLines -Day $script:D1 -StartTime '7:30' -Count '3' -Every '15').Times -join ',') | Should -BeExactly '0730,0745,0800'
        ((New-BulkBoardLines -Day $script:D1 -StartTime ' 730 ' -Count ' 2 ' -Every ' 5 ').Times -join ',') | Should -BeExactly '0730,0735'
        ((New-BulkBoardLines -Day $script:D1 -StartTime '1300' -Count '2' -Every '0').Times -join ',') | Should -BeExactly '1300,1300'
        ((New-BulkBoardLines -Day $script:D1 -StartTime '0800' -Count '1' -SortieType 'OFT').Lines[0].SortieType) | Should -BeExactly 'OFT'
    }

    It 'gives each O&B its linked return leg' {
        $r = New-BulkBoardLines -Day $script:D1 -StartTime '1000' -Count '2' -Every '5' -SortieType 'O&B'
        ($r.Times -join ',') | Should -BeExactly '1000,1005'
        @($r.Lines).Count | Should -Be 4
        $ob = $r.Lines[0]; $ret = $r.Lines[1]
        $ret.SortieType | Should -BeExactly 'O&B Return'
        $ret.ETD | Should -Be ([datetime]'2026-10-20 13:00')
        $ob.LinkedLineId | Should -BeExactly $ret.Id
        $ret.LinkedLineId | Should -BeExactly $ob.Id
        $r.Lines[3].ETD | Should -Be ([datetime]'2026-10-20 13:05')
    }

    It "uses the board's return offset and return type" {
        $r = New-BulkBoardLines -Day $script:D1 -StartTime '1000' -Count '1' -SortieType 'O&B' -Settings @{ OBReturnOffsetMinutes = 120; ReturnSortieType = 'OB RTN' }
        $r.Lines[1].ETD | Should -Be ([datetime]'2026-10-20 12:00')
        $r.Lines[1].SortieType | Should -BeExactly 'OB RTN'
    }

    It 'goes on a board as one undo step' {
        $b = New-TB
        $r = New-BulkBoardLines -Day $script:D1 -StartTime '0800' -Count '4' -Settings $b.Settings
        Add-BoardLine -Board $b -Line $r.Lines
        @(Get-BoardDayLines -Board $b -Day $script:D1).Count | Should -Be 4
        Undo-Board $b | Should -BeTrue
        @(Get-BoardDayLines -Board $b -Day $script:D1).Count | Should -Be 0
    }

    It 'says what is wrong (<Why>) and makes nothing' -TestCases @(
        @{ Why = 'not a time'; Start = '25:00'; Count = '3'; Every = '3'; Type = 'Day Local'; Want = "Start time '25:00' isn't a time (type 0730, 730 or 7:30)." }
        @{ Why = 'no start'; Start = ''; Count = '3'; Every = '3'; Type = 'Day Local'; Want = "Start time '' isn't a time (type 0730, 730 or 7:30)." }
        @{ Why = 'no lines'; Start = '0730'; Count = '0'; Every = '3'; Type = 'Day Local'; Want = 'Number of lines must be a whole number from 1 to 60.' }
        @{ Why = 'count in words'; Start = '0730'; Count = 'ten'; Every = '3'; Type = 'Day Local'; Want = 'Number of lines must be a whole number from 1 to 60.' }
        @{ Why = 'too many'; Start = '0730'; Count = '61'; Every = '3'; Type = 'Day Local'; Want = 'Number of lines must be a whole number from 1 to 60.' }
        @{ Why = 'negative gap'; Start = '0730'; Count = '3'; Every = '-3'; Type = 'Day Local'; Want = 'Minutes between lines must be a whole number from 0 to 240.' }
        @{ Why = 'no type'; Start = '0730'; Count = '3'; Every = '3'; Type = ' '; Want = 'Pick a sortie type.' }
        @{ Why = 'past midnight'; Start = '2330'; Count = '20'; Every = '3'; Type = 'Day Local'; Want = 'The last line would be at 0027 the next day. Start earlier, add fewer lines or put them closer together.' }
        @{ Why = 'O&B return past midnight'; Start = '2200'; Count = '1'; Every = '3'; Type = 'O&B'; Want = "The last O&B's return leg would be at 0100 the next day. Start earlier or add fewer lines." }
    ) {
        $r = New-BulkBoardLines -Day $script:D1 -StartTime $Start -Count $Count -Every $Every -SortieType $Type
        ($r.Errors -join '|') | Should -BeExactly $Want
        @($r.Lines).Count | Should -Be 0
        @($r.Times).Count | Should -Be 0
    }
}

Describe 'Add Student Absence: times, pasted names and the absences' {
    It 'reads <T> as <Want>' -TestCases @(
        @{ T = '0730'; Want = '07:30' }
        @{ T = '730'; Want = '07:30' }
        @{ T = '7:30'; Want = '07:30' }
        @{ T = '1530'; Want = '15:30' }
        @{ T = '0730 AM'; Want = '07:30' }
        @{ T = '0330 PM'; Want = '15:30' }
        @{ T = '0330PM'; Want = '15:30' }
        @{ T = '330pm'; Want = '15:30' }
        @{ T = '3:30 p.m.'; Want = '15:30' }
        @{ T = '3 PM'; Want = '15:00' }
        @{ T = '12:00 AM'; Want = '00:00' }
        @{ T = '1230 am'; Want = '00:30' }
        @{ T = '1200 PM'; Want = '12:00' }
        @{ T = '1159 pm'; Want = '23:59' }
        @{ T = ' 0900 '; Want = '09:00' }
    ) {
        $ts = ConvertTo-ClockTime $T
        $ts | Should -BeOfType [timespan]
        $ts.ToString('hh\:mm') | Should -BeExactly $Want
    }

    It 'rejects <T>' -TestCases @(
        @{ T = '1530 PM' }
        @{ T = '0030 AM' }
        @{ T = '13 PM' }
        @{ T = '2500' }
        @{ T = '0760' }
        @{ T = 'noon' }
        @{ T = 'PM' }
        @{ T = '' }
    ) {
        $null -eq (ConvertTo-ClockTime $T) | Should -BeTrue
    }

    It 'pasted names: one per line, cell or ";", matched exactly or as First Last or with a rank, in the order pasted, each once' {
        $stu = @((New-TestStudent -Name 'Quill, Casey'), (New-TestStudent -Name 'Adams, Amy'), (New-TestStudent -Name 'Baker, Ben'), (New-TestStudent -Name 'SP1'))
        $r = Resolve-PastedStudents -Text "Casey B. Quill`r`nadams, amy`tSP1; Baker, Ben [2d Lt]`n`nQuill, Casey" -Students $stu
        @($r.Students | ForEach-Object { $_.Name }) | Should -Be @('Quill, Casey', 'Adams, Amy', 'SP1', 'Baker, Ben')
        @($r.Unknown) | Should -BeNullOrEmpty
    }

    It 'pasted names: "Adams, Baker" is two names; a name that is no one, or more than one, is listed' {
        $stu = @($script:Stu) + @((New-TestStudent -Name 'Lee, Ann'), (New-TestStudent -Name 'Lee, Bob'))
        $r = Resolve-PastedStudents -Text "Adams, Baker`nNobody, No`nLee`nCruz" -Students $stu
        @($r.Students | ForEach-Object { $_.Name }) | Should -Be @('Adams, Amy', 'Baker, Ben', 'Cruz, Cal')
        @($r.Unknown) | Should -Be @('Nobody, No', 'Lee')
        @((Resolve-PastedStudents -Text "  `r`n " -Students $stu).Students).Count | Should -Be 0
    }

    It 'one absence per student, on the day and at the times typed, the reason trimmed, not in MasterData yet' {
        $r = New-StudentAbsences -Students @('Adams, Amy', 'Baker, Ben', 'Adams, Amy') -Day ([datetime]'2026-10-20') -From '0730' -To '0330 PM' -Reason ' Dentist '
        @($r.Errors) | Should -BeNullOrEmpty
        @($r.Absences).Count | Should -Be 2
        $r.Absences[0].Student | Should -BeExactly 'Adams, Amy'
        $r.Absences[0].StartTime | Should -Be ([datetime]'2026-10-20 07:30')
        $r.Absences[0].EndTime | Should -Be ([datetime]'2026-10-20 15:30')
        $r.Absences[1].Student | Should -BeExactly 'Baker, Ben'
        $r.Absences[1].Reason | Should -BeExactly 'Dentist'
        $r.Absences[1].InMasterData | Should -BeFalse
        Format-AbsenceTime $r.Absences[0] | Should -BeExactly 'Tue 20 Oct 0730-1530'
    }

    It 'ends on a later day when one is picked (past midnight, or over several days)' {
        $r = New-StudentAbsences -Students @('Adams, Amy') -Day '2026-10-20' -From '2200' -To '0600' -EndDay '2026-10-21'
        $r.Absences[0].EndTime | Should -Be ([datetime]'2026-10-21 06:00')
        Format-AbsenceTime $r.Absences[0] | Should -BeExactly 'Tue 20 Oct 2200 - Wed 21 Oct 0600'
    }

    It 'says what is missing or wrong, and makes nothing: <Why>' -TestCases @(
        @{ Why = 'no student'; Students = @(); Day = '2026-10-20'; From = '0730'; To = '0900'; End = $null; Want = 'Add at least one student.' }
        @{ Why = 'no day'; Students = @('A'); Day = $null; From = '0730'; To = '0900'; End = $null; Want = 'Pick the day on the calendar.' }
        @{ Why = 'no start time'; Students = @('A'); Day = '2026-10-20'; From = ' '; To = '0900'; End = $null; Want = 'Type the start time. *' }
        @{ Why = 'a 24-hour time with PM'; Students = @('A'); Day = '2026-10-20'; From = '0730'; To = '1530 PM'; End = $null; Want = "End time '1530 PM' isn't a time. Type 24-hour time (0730, 1530) or 12-hour time with AM or PM (0730 AM, 0330 PM)." }
        @{ Why = 'ends before it starts'; Students = @('A'); Day = '2026-10-20'; From = '1300'; To = '0900'; End = $null; Want = 'It ends (Tue 20 Oct 0900) before it starts (Tue 20 Oct 1300). For an absence past midnight, pick the next day for the end.' }
        @{ Why = 'ends the day before'; Students = @('A'); Day = '2026-10-20'; From = '0730'; To = '0900'; End = '2026-10-19'; Want = 'It ends (Mon 19 Oct 0900) before it starts*' }
    ) {
        $r = New-StudentAbsences -Students $Students -Day $Day -From $From -To $To -EndDay $End
        @($r.Absences).Count | Should -Be 0
        @($r.Errors).Count | Should -Be 1
        $r.Errors[0] | Should -BeLike $Want
    }

    It 'gives each student their absences once, a student with none yet included; a name that is no one is left out' {
        $a = New-TestStudent -Name 'Adams, Amy'
        $b = [PSCustomObject]@{ Name = 'Baker, Ben' }
        $r = New-StudentAbsences -Students @('Adams, Amy', 'baker, ben', 'Nobody, No') -Day '2026-10-20' -From '0900' -To '1100'
        @(Add-StudentAbsences -Absences $r.Absences -Students @($a, $b)).Count | Should -Be 2
        @($a.Absences).Count | Should -Be 1
        @($b.Absences).Count | Should -Be 1
        @(Add-StudentAbsences -Absences $r.Absences -Students @($a, $b)).Count | Should -Be 0 -Because 'they already have them'
        @($a.Absences).Count | Should -Be 1
        $later = New-StudentAbsences -Students @('Adams, Amy') -Day '2026-10-20' -From '1300' -To '1400'
        @(Add-StudentAbsences -Absences $later.Absences -Students @($a)).Count | Should -Be 1
        @($a.Absences | ForEach-Object { Format-AbsenceTime $_ }) | Should -Be @('Tue 20 Oct 0900-1100', 'Tue 20 Oct 1300-1400')
    }

    It 'the ABSENCE hard rule sees an absence added this way' {
        $a = New-TestStudent -Name 'Adams, Amy' -Done 'FI3101|2026-10-15 08:00'
        $board = New-TB
        Add-BoardLine -Board $board -Line @(New-TL -At '2026-10-20 10:00' -IP 'Jones, Bob' -Student 'Adams, Amy') -NoUndo
        $data = @{ Students = @($a); Instructors = $script:IPs; SoftWeights = (New-TestSoftWeights); Formations = (New-TestFormations) }
        $r = New-StudentAbsences -Students @('Adams, Amy') -Day '2026-10-20' -From '0800' -To '0900 AM' -Reason 'Dentist'
        [void](Add-StudentAbsences -Absences $r.Absences -Students @($a))
        Clear-AssessmentCache
        $check = Get-PlacementAssessment -Board $board -LineId $board.Lines[0].Id -StudentName 'Adams, Amy' -Data $data
        @($check.Violations | Where-Object { $_.Rule -eq 'ABSENCE' } | ForEach-Object { $_.Message }) | Should -Be @('Unavailable 0800-0900 (Dentist); sortie window 0830-1230')
    }
}
