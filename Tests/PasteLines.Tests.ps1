# Pester 5 tests for Paste lines: lines copied from Excel read into board lines
# (ConvertFrom-ClipboardCells, Read-PastedSchedule, ConvertFrom-ScheduleCells -Pasted,
# ConvertFrom-PastedLines, Import-BoardLines). Pure PowerShell: no Excel, no window.
# The Paste lines window itself is in PuckBoardLogic.Tests.ps1.
# A pasted date must be within five years of today, so the dated tests use tomorrow (P1)
# and the day after (P2), written yyyy-MM-dd so every culture reads them the same.

BeforeAll {
    . "$PSScriptRoot/Fixtures/New-FaisTestData.ps1"
    Import-FaisModules

    $script:IPs = New-TestInstructors
    $script:SPs = @((New-TestStudent -Name 'Adams, Al'), (New-TestStudent -Name 'Brown, Bea'))
    $script:P1 = (Get-Date).Date.AddDays(1).ToString('yyyy-MM-dd')
    $script:P2 = (Get-Date).Date.AddDays(2).ToString('yyyy-MM-dd')

    # Cells as Excel copies them. Write a row as 'a ; b ; c': the cells get a tab between
    # them and every row ends in CR LF.
    function ConvertTo-Copied {
        param ([string[]]$Rows)
        return (@($Rows | ForEach-Object { ($_ -split '\s*;\s*') -join "`t" }) -join "`r`n") + "`r`n"
    }
    function Read-Paste {
        param ([string[]]$Rows)
        return (Read-PastedSchedule -Text (ConvertTo-Copied $Rows) -Instructors $script:IPs -Students $script:SPs)
    }
    # '07:30 Day Local Jones, Bob | 13:00 OFT SIM IP'
    function Get-RowText {
        param ($Read)
        return (@($Read.Rows | ForEach-Object { "$($_.Time) $($_.SortieType) $($_.Instructor)".Trim() }) -join ' | ')
    }
    # '0730 Day Local Jones, Bob | ...' for board lines, in ETD order
    function Get-LineText {
        param ([array]$Lines)
        return (@($Lines | Sort-Object -Property ETD | ForEach-Object { "$($_.ETD.ToString('HHmm')) $($_.SortieType) $($_.Instructor)".Trim() }) -join ' | ')
    }

    $script:Wing = @(
        'Line ; ETD ; Callsign ; # ; Student ; Mission ; Instructor ; Remarks ; Sortie Type'
        '1 ; 0645 ;  ;  ; Adams, Al [2d Lt] ; FT5103 ; JONES, BOB [Maj] ; MND ; Day Local'
        '2 ; 0648 ; Iron ; 78 ; Nobody, Nick [2d Lt] ; FT5006 ; Lee,  Ann ; TCL ; O&B'
        '3 ; 0700 ;  ;  ; For sale ;  ; Stoll, Joshua L. [Maj] ;  ; Day Local'
        '4 ; 0948 ;  ;  ;  ;  ; Lee, Ann ;  ; O&B Return'
        '5 ; 1000 ;  ;  ;  ;  ; Ortiz ;  ; O&B'
        '6 ; 1300 ;  ;  ; Brown ;  ; DILLINGER, JOE [CIV] ;  ; OFT 3'
    )
}

Describe 'ConvertFrom-ClipboardCells' {
    It 'splits rows on line breaks and cells on tabs' {
        $g = @(ConvertFrom-ClipboardCells -Text "a`tb`r`nc`t`td`r`n")
        ($g[0] -join '|') | Should -BeExactly 'a|b'
        ($g[1] -join '|') | Should -BeExactly 'c||d'
        @($g | Where-Object { @($_ | Where-Object { $_ -ne '' }).Count -gt 0 }).Count | Should -Be 2
    }

    It 'a lone LF or CR ends a row too, and cells are trimmed' {
        $g = @(ConvertFrom-ClipboardCells -Text "a `t b`nc`td`re`tf")
        (@($g | ForEach-Object { $_ -join '|' }) -join ' / ') | Should -BeExactly 'a|b / c|d / e|f'
    }

    It 'a quoted cell keeps its tabs and quotes, and its line break becomes a space' {
        $g = @(ConvertFrom-ClipboardCells -Text "ETD`tNotes`r`n0730`t`"brief 0630`nbring `"`"pubs`"`"`"`r`n0800`t`"a`tb`"`r`n")
        $g.Count | Should -Be 3
        $g[1][1] | Should -BeExactly 'brief 0630 bring "pubs"'
        $g[2][1] | Should -BeExactly "a`tb"
    }

    It 'a quote that is not at the start of a cell, or never closed, is just a quote' {
        $g = @(ConvertFrom-ClipboardCells -Text "5`" gap`tx`r`n`"open`ty")
        ($g[0] -join '|') | Should -BeExactly '5" gap|x'
        ($g[1] -join '|') | Should -BeExactly '"open|y'
    }

    It 'text with no tabs is split on | or on runs of two or more spaces' {
        $g = @(ConvertFrom-ClipboardCells -Text "| ETD | IP | Type |`n0730 | Jones, Bob | DL")
        (@($g | ForEach-Object { $_ -join '/' }) -join ' ; ') | Should -BeExactly 'ETD/IP/Type ; 0730/Jones, Bob/DL'
        $s = @(ConvertFrom-ClipboardCells -Text "0730  Day Local   Jones, Bob`n0800  OFT")
        ($s[0] -join '/') | Should -BeExactly '0730/Day Local/Jones, Bob'
        ($s[1] -join '/') | Should -BeExactly '0800/OFT'
    }

    It 'nothing gives no rows' {
        @(ConvertFrom-ClipboardCells -Text '').Count | Should -Be 0
        @(ConvertFrom-ClipboardCells -Text $null).Count | Should -Be 0
    }
}

Describe 'Read-PastedSchedule' {
    Context 'with the heading row copied too' {
        It 'finds Sortie Type, ETD and Instructor by their headings, in any order' {
            $r = Read-Paste @('Sortie Type ; ETD ; Instructor', 'Day Local ; 0730 ; Jones, Bob', 'OB ; 1015 ; Ortiz, Sam', 'O&B Return ; 1315 ; Ortiz, Sam', 'OFT ; 1300 ; ')
            $r.Found | Should -BeTrue
            $r.Error | Should -BeExactly ''
            $r.Layout | Should -BeExactly 'Headings on row 1'
            $r.HeaderRow | Should -Be 1
            Get-RowText $r | Should -BeExactly '07:30 Day Local Jones, Bob | 10:15 O&B Ortiz, Sam | 13:15 O&B Return Ortiz, Sam | 13:00 OFT SIM IP'
            $r.ColumnText | Should -BeExactly 'ETD = column 2, Sortie Type = column 1, IP = column 3 (from the headings on row 1)'
            @($r.Detected).Count | Should -Be 0
            (@($r.Rows | ForEach-Object { $_.SheetRow }) -join ',') | Should -BeExactly '2,3,4,5'
        }

        It 'a time with no date has no day yet (it goes on the day the lines are put on)' {
            $r = Read-Paste @('ETD ; Instructor ; SortieType', '0730 ; Jones, Bob ; DL')
            $r.Rows[0].TimeOnly | Should -BeTrue
            $r.Rows[0].Day | Should -BeExactly ''
            $r.Rows[0].ETD | Should -BeNullOrEmpty
        }

        It 'reads the flying schedule''s Flying tab as it is: other columns are left alone, a tail number is not a date' {
            $r = Read-Paste @(
                'Line ; Tail # ; ETD ; Callsign ; # ; Student ; Mission ; Instructor ; Remarks ; Sortie Type'
                '1 ; 03-3678 ; 0645 ;  ;  ; Adams, Al [2d Lt] ; FT5103 ; Jones, Bob [Maj] ; MND ; Day Local'
                '2 ; 02-3658 ; 0648 ; Iron ; 78 ;  ; FT5006 ; Lee,  Ann ; TCL ; O&B'
                '3 ; 03-3678 ;  ;  ;  ;  ;  ;  ;  ; '
                '4 ; 3855 ; 0948 ;  ;  ;  ;  ; Lee, Ann ;  ; O&B Return'
            )
            $r.Error | Should -BeExactly ''
            @($r.Skipped).Count | Should -Be 0 -Because 'a row with no line on it is not a line that failed'
            (@($r.Rows | ForEach-Object { "$($_.Time) $($_.SortieType)" }) -join ' | ') | Should -BeExactly '06:45 Day Local | 06:48 O&B | 09:48 O&B Return'
            @($r.Rows | Where-Object { $_.Day -ne '' }).Count | Should -Be 0
            $r.Rows[0].Student | Should -BeExactly 'Adams, Al [2d Lt]'
            $r.Rows[0].Mission | Should -BeExactly 'FT5103'
            $r.Rows[0].Notes | Should -BeExactly 'MND'
            $r.Rows[1].Callsign | Should -BeExactly 'Iron'
            $r.Rows[1].Num | Should -BeExactly '78'
            $r.ColumnText | Should -BeExactly 'ETD = column 3, Sortie Type = column 10, IP = column 8, Student = column 6, Mission = column 7, Callsign = column 4, # = column 5, Notes = column 9 (from the headings on row 1)'
        }

        It 'reads the Sims tab: Start Time, Device and Simulator Instructor' {
            $r = Read-Paste @('SQ ; Duration ; Device ; Simulator Instructor ; Student ; Event Name ; REMARK ; Start Time', 'ST37 ; 1.3 ; OFT 3 ; DILLINGER, JOSEPH A. [CIV] ;  ;  ;  ; 0930')
            Get-RowText $r | Should -BeExactly '09:30 OFT 3 DILLINGER, JOSEPH A. [CIV]'
        }

        It 'a heading FAIS does not know for the IP: the column is found from the names in it' {
            $r = Read-Paste @('ETD ; Type ; Crew', '0730 ; DL ; Jones, Bob', '0800 ; DL ; Lee, Ann')
            Get-RowText $r | Should -BeExactly '07:30 Day Local Jones, Bob | 08:00 Day Local Lee, Ann'
            (@($r.Detected) -join ',') | Should -BeExactly 'Instructor'
            $r.ColumnText | Should -BeExactly 'ETD = column 1, Sortie Type = column 2, IP = column 3 (from the headings on row 1; IP worked out from the cells)'
        }

        It 'a date and time in the ETD dates the line' {
            $r = Read-Paste @('ETD ; Instructor ; SortieType', "$($script:P1) 08:10 ; Jones, Bob ; Day Local", "$($script:P2) 0845 ; Lee, Ann ; Day Local")
            (@($r.Rows | ForEach-Object { $_.ETD }) -join ',') | Should -BeExactly "$($script:P1) 08:10:00,$($script:P2) 08:45:00"
            $r.Rows[0].TimeOnly | Should -BeFalse
        }

        It 'a repeated heading row and blank rows are skipped' {
            $r = Read-Paste @('ETD ; IP ; Type', '0730 ; Jones, Bob ; DL', ' ;  ; ', 'ETD ; IP ; Type', '0800 ; Lee, Ann ; DL')
            Get-RowText $r | Should -BeExactly '07:30 Day Local Jones, Bob | 08:00 Day Local Lee, Ann'
        }
    }

    Context 'with no heading row' {
        It 'works the columns out from the cells, in whatever order they are' {
            $r = Read-Paste @('Day Local ; 0730 ; Jones, Bob', 'DL ; 7:45 ; Lee, Ann [Capt]', 'OFT ; 1300 ; ')
            $r.Error | Should -BeExactly ''
            $r.Layout | Should -BeExactly 'No heading row'
            $r.HeaderRow | Should -Be 0
            Get-RowText $r | Should -BeExactly '07:30 Day Local Jones, Bob | 07:45 Day Local Lee, Ann [Capt] | 13:00 OFT SIM IP'
            $r.ColumnText | Should -BeExactly 'ETD = column 2, Sortie Type = column 1, IP = column 3 (no heading row: worked out from the cells)'
            (@($r.Detected) -join ',') | Should -BeExactly 'ETD,SortieType,Instructor'
        }

        It 'IPs MasterData does not have: the first column of Last, First names is the IP' {
            $r = Read-Paste @('0730 ; Day Local ; Stoll, Joshua L. [Maj]', '0745 ; Day Local ; Bates, Thomas R. [Capt]')
            $r.Columns.Instructor | Should -Be 3
            $r.Rows[1].Instructor | Should -BeExactly 'Bates, Thomas R. [Capt]'
        }

        It 'tells a column of students from one of IPs, and finds the missions' {
            $r = Read-Paste @('Adams, Al ; 0730 ; DL ; Jones, Bob ; FT5103', 'Brown, Bea ; 0900 ; DL ; Lee, Ann ; FT5001')
            $r.Columns.Student | Should -Be 1
            $r.Columns.Instructor | Should -Be 4
            $r.Columns.Mission | Should -Be 5
            $r.Rows[1].Student | Should -BeExactly 'Brown, Bea'
            $r.Rows[1].Mission | Should -BeExactly 'FT5001'
        }

        It 'a column of dates dates the lines' {
            $r = Read-Paste @("$($script:P1) ; 0730 ; DL ; Jones, Bob", "$($script:P2) ; 0800 ; DL ; Lee, Ann")
            $r.Columns.Date | Should -Be 1
            (@($r.Rows | ForEach-Object { $_.ETD }) -join ',') | Should -BeExactly "$($script:P1) 07:30:00,$($script:P2) 08:00:00"
        }

        It 'a date on a row of its own dates the rows under it' {
            $r = Read-Paste @($script:P1, '0730 ; DL ; Jones, Bob', '0800 ; DL ; Lee, Ann')
            (@($r.Rows | ForEach-Object { $_.ETD }) -join ',') | Should -BeExactly "$($script:P1) 07:30:00,$($script:P1) 08:00:00"
        }

        It 'tail numbers and line numbers are neither ETDs nor dates' {
            $r = Read-Paste @('1 ; 03-3678 ; 0645 ; Day Local ; Jones, Bob', '2 ; 02-3658 ; 0648 ; Day Local ; Lee, Ann')
            $r.Columns.ETD | Should -Be 3
            $r.Columns.ContainsKey('Date') | Should -BeFalse
            Get-RowText $r | Should -BeExactly '06:45 Day Local Jones, Bob | 06:48 Day Local Lee, Ann'
            @($r.Rows | Where-Object { $_.Day -ne '' }).Count | Should -Be 0
        }
    }

    Context 'what cannot be read' {
        It 'text that is not a schedule says what is missing' {
            $r = Read-PastedSchedule -Text 'hello there'
            $r.Found | Should -BeTrue
            $r.Error | Should -BeExactly 'What was pasted has no heading row, and FAIS could not find the ETD and Sortie Type columns from the cells. Copy the heading row too (ETD, Sortie Type, Instructor).'
            @($r.Rows).Count | Should -Be 0
        }

        It 'headings with no sortie type column, and none to be found in the cells' {
            $r = Read-Paste @('ETD ; Instructor', '0730 ; Jones, Bob')
            $r.Error | Should -BeExactly 'What was pasted has headings on row 1 but no Sortie Type heading, and FAIS could not find the column from the cells. A line needs an ETD and a Sortie Type.'
        }

        It 'nothing pasted is not an error' {
            foreach ($text in @('', "  `r`n`t`r`n", $null)) {
                $r = Read-PastedSchedule -Text $text
                $r.Found | Should -BeFalse
                $r.Error | Should -BeExactly ''
                @($r.Rows).Count | Should -Be 0
            }
        }

        It 'reads the good rows and says why each other row was skipped' {
            $r = Read-Paste @('ETD ; Instructor ; SortieType', '2500 ; Jones, Bob ; DL', ' ; Jones, Bob ; DL', '0800 ; Lee, Ann ; DL')
            Get-RowText $r | Should -BeExactly '08:00 Day Local Lee, Ann'
            $s = @($r.Skipped)
            $s.Count | Should -Be 2
            $s[0] | Should -BeLike "row 2: ETD '2500' is not a time*"
            $s[1] | Should -BeExactly 'row 3: no ETD (Day Local line)'
        }
    }
}

Describe 'ConvertFrom-ScheduleCells' {
    It 'the Schedule sheet with no headings still uses its fixed columns; pasted cells are worked out' {
        $grid = @(@('0730', 'Jones, Bob', 'DL'), @('1300', '', 'OFT'))
        $sheet = ConvertFrom-ScheduleCells -Grid $grid
        $sheet.Layout | Should -BeLike 'Fixed columns*'
        Get-RowText $sheet | Should -BeExactly '07:30 Day Local Jones, Bob | 13:00 OFT SIM IP'
        $pasted = ConvertFrom-ScheduleCells -Grid $grid -Pasted -Instructors $script:IPs
        $pasted.Layout | Should -BeExactly 'No heading row'
        Get-RowText $pasted | Should -BeExactly '07:30 Day Local Jones, Bob | 13:00 OFT SIM IP'
    }

    It 'the Schedule sheet with headings but no ETD column is an error that names the sheet' {
        $r = ConvertFrom-ScheduleCells -Grid @(@('Instructor', 'SortieType', 'Student'), @('Jones, Bob', 'DL', 'Adams, Al')) -What "The 'Lines' sheet"
        $r.Error | Should -BeExactly "The 'Lines' sheet has headings on row 1 but no ETD column."
    }

    It 'on the sheet a date among other cells on a row with no line dates the rows under it; pasted, the date must be on its own' {
        $grid = @(@('ETD', 'Instructor', 'SortieType', 'Notes', 'Week'), @('', '', '', $script:P1, 'week 43'), @('0730', 'Jones, Bob', 'DL', '', ''))
        (ConvertFrom-ScheduleCells -Grid $grid).Rows[0].Day | Should -BeExactly $script:P1
        $pasted = ConvertFrom-ScheduleCells -Grid $grid -Pasted
        @($pasted.Rows).Count | Should -Be 1
        $pasted.Rows[0].Day | Should -BeExactly ''
    }

    It 'pasted, a date more than five years off is not a date' {
        $far = (Get-Date).Date.AddYears(9).ToString('yyyy-MM-dd')
        $pasted = ConvertFrom-ScheduleCells -Grid @(@('ETD', 'Instructor', 'SortieType'), @($far, '', ''), @('0730', 'Jones, Bob', 'DL'), @("$far 08:00", 'Lee, Ann', 'DL')) -Pasted
        @($pasted.Rows).Count | Should -Be 1
        $pasted.Rows[0].Day | Should -BeExactly ''
        @($pasted.Skipped | Where-Object { $_ -like "row 4: ETD '$far 08:00' is not a time*" }).Count | Should -Be 1
    }
}

Describe 'Find-ScheduleColumns' {
    It 'takes a column for a field only when at least half of its cells read as one' {
        $grid = @(@('0730', 'note', 'DL'), @('0745', '0800', 'DL'), @('0800', 'x', 'OFT'), @('0815', 'y', 'DL'))
        $c = Find-ScheduleColumns -Grid $grid
        $c.ETD | Should -Be 1
        $c.SortieType | Should -Be 3
        $c.ContainsKey('Instructor') | Should -BeFalse
    }

    It 'leaves alone the columns a heading already spoke for, and looks only for the fields asked for' {
        $grid = @(@('0730', '0900', 'DL', 'Jones, Bob'), @('0745', '0915', 'DL', 'Lee, Ann'))
        $c = Find-ScheduleColumns -Grid $grid -Fields @('ETD', 'Instructor') -Taken @(1) -Instructors $script:IPs
        $c.ETD | Should -Be 2
        $c.Instructor | Should -Be 4
        $c.ContainsKey('SortieType') | Should -BeFalse
    }
}

Describe 'ConvertFrom-PastedLines' {
    Context 'the lines' {
        BeforeAll {
            $script:R = ConvertFrom-PastedLines -Text (ConvertTo-Copied $script:Wing) -Day '2026-10-21' -KeepStudents -Instructors $script:IPs -Students $script:SPs
            $script:L = @($script:R.Lines | Sort-Object -Property ETD)
        }

        It 'puts a line on the day for every row read' {
            $script:R.Error | Should -BeExactly ''
            $script:R.Found | Should -BeTrue
            $script:R.Rows | Should -Be 6
            (@($script:R.Days | ForEach-Object { "$($_.Day)=$($_.Count)" }) -join ',') | Should -BeExactly '2026-10-21=6'
            @($script:R.PastedDays).Count | Should -Be 0
            (@($script:L | ForEach-Object { $_.Source } | Select-Object -Unique) -join ',') | Should -BeExactly 'Pasted'
            $script:L[0].ETD | Should -Be ([datetime]'2026-10-21 06:45')
        }

        It 'writes the IPs the way MasterData spells them, and lists the ones it does not have' {
            Get-LineText $script:L | Should -BeExactly '0645 Day Local Jones, Bob | 0648 O&B Lee, Ann | 0700 Day Local Stoll, Joshua L. [Maj] | 0948 O&B Return Lee, Ann | 1000 O&B Ortiz, Sam | 1300 OFT 3 DILLINGER, JOE [CIV]'
            (@($script:R.UnknownInstructors) -join ';') | Should -BeExactly 'Stoll, Joshua L. [Maj]' -Because 'a sim line''s instructor is not on the Instructors sheet anyway'
        }

        It 'links an O&B to its return leg, and counts the O&Bs that have none' {
            $ob = $script:L[1]
            $ret = $script:L[3]
            $ob.LinkedLineId | Should -BeExactly $ret.Id
            $ret.LinkedLineId | Should -BeExactly $ob.Id
            $script:L[4].LinkedLineId | Should -BeExactly ''
            $script:R.OpenOBs | Should -Be 1
        }

        It 'hardcodes the students FAIS knows (with the mission), and leaves the others off their lines' {
            $script:L[0].Student | Should -BeExactly 'Adams, Al'
            $script:L[0].Mission | Should -BeExactly 'FT5103'
            $script:L[5].Student | Should -BeExactly 'Brown, Bea' -Because 'Brown is one student'
            $script:L[1].Student | Should -BeExactly ''
            $script:L[1].Mission | Should -BeExactly ''
            $script:L[2].Student | Should -BeExactly '' -Because 'For sale is a placeholder'
            (@($script:R.UnknownStudents) -join ';') | Should -BeExactly 'Nobody, Nick [2d Lt]'
            $script:R.PastedStudents | Should -Be 3
            $script:R.Students | Should -Be 2
        }

        It 'keeps the notes, and the callsign of a line whose students are kept' {
            $script:L[0].Notes | Should -BeExactly 'MND'
            $script:L[1].Callsign | Should -BeExactly 'Iron'
            $script:L[1].Num | Should -BeExactly '78'
        }
    }

    It 'without -KeepStudents no student, mission or callsign goes on' {
        $r = ConvertFrom-PastedLines -Text (ConvertTo-Copied $script:Wing) -Day '2026-10-21' -Instructors $script:IPs -Students $script:SPs
        @($r.Lines | Where-Object { $_.Student -or $_.Mission -or $_.Callsign }).Count | Should -Be 0
        $r.Students | Should -Be 0
        $r.PastedStudents | Should -Be 3
        @($r.UnknownStudents).Count | Should -Be 0
        @($r.Lines | Sort-Object -Property ETD)[1].Notes | Should -BeExactly 'TCL'
    }

    Context 'dates that were pasted' {
        BeforeAll {
            $script:Dated = ConvertTo-Copied @('ETD ; Instructor ; SortieType', "$($script:P1) 08:10 ; Jones, Bob ; Day Local", "$($script:P2) 08:45 ; Lee, Ann ; Day Local", '0900 ; Lee ; DL')
        }

        It 'all go on the day picked' {
            $r = ConvertFrom-PastedLines -Text $script:Dated -Day '2026-10-23' -Instructors $script:IPs
            (@($r.Days | ForEach-Object { "$($_.Day)=$($_.Count)" }) -join ',') | Should -BeExactly '2026-10-23=3'
            Get-LineText $r.Lines | Should -BeExactly '0810 Day Local Jones, Bob | 0845 Day Local Lee, Ann | 0900 Day Local Lee, Ann'
            (@($r.PastedDays | ForEach-Object { "$($_.Day)=$($_.Count)" }) -join ',') | Should -BeExactly "$($script:P1)=1,$($script:P2)=1"
        }

        It 'with -KeepDates each goes on its own date, and a line with none on the day picked' {
            $r = ConvertFrom-PastedLines -Text $script:Dated -Day ([datetime]'2026-10-23') -KeepDates -Instructors $script:IPs
            (@($r.Days | ForEach-Object { "$($_.Day)=$($_.Count)" } | Sort-Object) -join ',') | Should -BeExactly (@("$($script:P1)=1", "$($script:P2)=1", '2026-10-23=1' | Sort-Object) -join ',')
            @($r.Lines | Where-Object { $_.Day -eq $script:P2 })[0].ETD.ToString('HHmm') | Should -BeExactly '0845'
        }
    }

    It 'no day, nothing pasted and text that is not lines' {
        (ConvertFrom-PastedLines -Text (ConvertTo-Copied $script:Wing) -Day '').Error | Should -BeExactly 'Pick the day to put the lines on.'
        $none = ConvertFrom-PastedLines -Text '' -Day '2026-10-21'
        $none.Found | Should -BeFalse
        $none.Error | Should -BeExactly ''
        @($none.Lines).Count | Should -Be 0
        $junk = ConvertFrom-PastedLines -Text 'hello there' -Day '2026-10-21'
        $junk.Error | Should -BeLike 'What was pasted has no heading row*'
        @($junk.Lines).Count | Should -Be 0
    }

    It 'uses the board''s return type and sim instructor' {
        $r = ConvertFrom-PastedLines -Text (ConvertTo-Copied @('ETD ; Type ; IP', '1015 ; OB ; Ortiz, Sam', '1315 ; RTN ; Ortiz, Sam', '1300 ; OFT ; ')) -Day '2026-10-21' -Instructors $script:IPs -Settings @{ ReturnSortieType = 'OB Return'; SimInstructor = 'SIM' }
        Get-LineText $r.Lines | Should -BeExactly '1015 O&B Ortiz, Sam | 1300 OFT SIM | 1315 OB Return Ortiz, Sam'
        $r.OpenOBs | Should -Be 0
    }
}

Describe 'Import-BoardLines' {
    BeforeEach {
        $script:B = New-Board -Days '2026-10-20'
        Add-BoardLine -Board $script:B -Line (New-BoardLine -ETD ([datetime]'2026-10-20 06:00') -SortieType 'Day Local' -Id 'OLD') -NoUndo
        $script:New = @(
            New-BoardLine -ETD ([datetime]'2026-10-20 07:30') -SortieType 'Day Local' -Instructor 'Jones, Bob'
            New-BoardLine -ETD ([datetime]'2026-10-21 08:00') -SortieType 'Day Local' -Instructor 'Lee, Ann'
            New-BoardLine -ETD ([datetime]'2026-10-22 09:00') -SortieType 'OFT'
        )
    }

    It 'adds every line on its own day (new days too) as one undo step' {
        $res = Import-BoardLines -Board $script:B -Lines $script:New
        $res.Added | Should -Be 3
        $res.Removed | Should -Be 0
        ($res.Days -join ',') | Should -BeExactly '2026-10-20,2026-10-21,2026-10-22'
        ($script:B.Days.ToArray() -join ',') | Should -BeExactly '2026-10-20,2026-10-21,2026-10-22'
        $script:B.Lines.Count | Should -Be 4
        $script:B.Undo.Count | Should -Be 1
        Undo-Board $script:B | Should -BeTrue
        $script:B.Lines.Count | Should -Be 1
        ($script:B.Days.ToArray() -join ',') | Should -BeExactly '2026-10-20'
    }

    It '-Replace clears the days the lines go on, and no other day' {
        Add-BoardLine -Board $script:B -Line (New-BoardLine -ETD ([datetime]'2026-10-23 06:00') -SortieType 'Day Local' -Id 'KEEP') -NoUndo
        $res = Import-BoardLines -Board $script:B -Lines $script:New -Replace
        $res.Removed | Should -Be 1
        Get-BoardLine $script:B 'OLD' | Should -BeNullOrEmpty
        Get-BoardLine $script:B 'KEEP' | Should -Not -BeNullOrEmpty
        $script:B.Lines.Count | Should -Be 4
        Undo-Board $script:B | Should -BeTrue
        Get-BoardLine $script:B 'OLD' | Should -Not -BeNullOrEmpty
    }

    It 'no lines changes nothing and leaves no undo step' {
        $v = $script:B.Version
        $res = Import-BoardLines -Board $script:B -Lines @()
        $res.Added | Should -Be 0
        $script:B.Undo.Count | Should -Be 0
        $script:B.Version | Should -Be $v
    }
}
