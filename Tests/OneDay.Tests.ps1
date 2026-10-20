# Pester 5 tests for working one day at a time (Modules\BoardModel.psm1):
# loading a day from MasterData's Schedule sheet or the last exported schedule, tracking
# whether the sheet changed since, and filling one day with the other exported days as context.
# Dates: Tue 2026-10-20 (D1), Wed 2026-10-21 (D2), Thu 2026-10-22 (D3).

BeforeAll {
    . "$PSScriptRoot/Fixtures/New-FaisTestData.ps1"
    Import-FaisModules

    $script:D1 = '2026-10-20'
    $script:D2 = '2026-10-21'
    $script:D3 = '2026-10-22'
    $script:IPs = New-TestInstructors

    # A Read-ScheduleSheet row
    function New-SR {
        param ([string]$Day, [string]$Time, [string]$Type = 'Day Local', [string]$IP = 'Jones, Bob', [string]$Student = $null, [string]$Mission = $null, [string]$Notes = $null)
        [PSCustomObject]@{
            ETD = $(if ($Day) { "$Day $($Time):00" } else { $null }); Day = $Day; Time = $Time; TimeOnly = (-not $Day)
            Instructor = $IP; SortieType = $Type; Student = $(if ($Student) { $Student } else { $null }); Mission = $(if ($Mission) { $Mission } else { $null })
            Callsign = $null; Num = $null; Notes = $(if ($Notes) { $Notes } else { $null }); SheetRow = 1
        }
    }

    # A row as DownloadPrevSchedule gives it back from the exported MasterSchedule
    function New-PR {
        param ([string]$At, [string]$Type = 'Day Local', [string]$IP = 'Jones, Bob', [string]$Student = 'For Sale', [string]$Mission = '', [string]$Notes = '', [string]$Callsign = 'TALON', $Num = 31)
        [PSCustomObject]@{ ETD = ([datetime]$At).ToString('yyyy-MM-dd HH:mm:ss'); Instructor = $IP; SortieType = $Type; Student = $Student; Mission = $Mission; Callsign = $Callsign; Num = $Num; Notes = $Notes }
    }

    function Get-DayIds { param ($Board, [string]$Day) @($Board.Lines | Where-Object { $_.Day -eq $Day } | ForEach-Object { $_.Id }) }
}

Describe 'ConvertFrom-SheetRows' {
    BeforeAll {
        $script:Rows = @(
            New-SR $script:D1 '07:30' -Student 'Adams, Al' -Mission 'FT5103' -Notes 'brief 0630'
            New-SR $script:D1 '10:15' -Type 'O&B' -IP 'Ortiz, Sam'
            New-SR $script:D1 '13:15' -Type 'O&B Return' -IP 'Ortiz, Sam'
            New-SR $script:D2 '08:00'
            New-SR '' '15:00' -Type 'IFT' -IP 'SIM IP'
        )
    }

    It 'takes only the asked day' {
        $l = @(ConvertFrom-SheetRows -Rows $script:Rows -Day $script:D1)
        $l.Count | Should -Be 3
        @($l | Where-Object { $_.Day -ne $script:D1 }).Count | Should -Be 0
    }

    It 'adds the rows with no date onto that day with -IncludeTimeOnly' {
        $l = @(ConvertFrom-SheetRows -Rows $script:Rows -Day $script:D2 -IncludeTimeOnly)
        $l.Count | Should -Be 2
        ($l | Where-Object { $_.SortieType -eq 'IFT' }).ETD | Should -Be ([datetime]"$($script:D2) 15:00")
        ($l | Where-Object { $_.SortieType -eq 'IFT' }).Instructor | Should -BeExactly ''
        $script:Rows[4].ETD | Should -BeNullOrEmpty -Because 'the sheet rows are not changed'
    }

    It '-FromDay moves another day''s rows (or the no-date rows) onto -Day' {
        $l = @(ConvertFrom-SheetRows -Rows $script:Rows -Day $script:D3 -FromDay $script:D2)
        $l.Count | Should -Be 1
        $l[0].ETD | Should -Be ([datetime]"$($script:D3) 08:00")
        $n = @(ConvertFrom-SheetRows -Rows $script:Rows -Day $script:D3 -FromDay '')
        $n.Count | Should -Be 1
        $n[0].SortieType | Should -BeExactly 'IFT'
    }

    It 'Resolve-SheetRowDays dates the no-date rows on the day being built' {
        $r = @(Resolve-SheetRowDays -Rows $script:Rows -DefaultDay $script:D1)
        $r.Count | Should -Be 5
        $r[4].Day | Should -BeExactly $script:D1
        $r[4].ETD | Should -BeExactly "$($script:D1) 15:00:00"
        $r[4].TimeOnly | Should -BeTrue
        [object]::ReferenceEquals($r[0], $script:Rows[0]) | Should -BeTrue
        $script:Rows[4].Day | Should -BeExactly ''
        @(ConvertFrom-SheetRows -Rows $r -Day $script:D1).Count | Should -Be 4
    }

    It 'hardcodes students only with -KeepStudents, and links the O&B to its return' {
        $plain = @(ConvertFrom-SheetRows -Rows $script:Rows -Day $script:D1)
        $plain[0].Student | Should -BeExactly ''
        $kept = @(ConvertFrom-SheetRows -Rows $script:Rows -Day $script:D1 -KeepStudents -KeepNotes)
        $kept[0].Student | Should -BeExactly 'Adams, Al'
        $kept[0].Mission | Should -BeExactly 'FT5103'
        $kept[0].Notes | Should -BeExactly 'brief 0630'
        $ob = $kept | Where-Object { $_.SortieType -eq 'O&B' }
        $ret = $kept | Where-Object { $_.SortieType -eq 'O&B Return' }
        $ob.LinkedLineId | Should -BeExactly $ret.Id
        $ret.LinkedLineId | Should -BeExactly $ob.Id
    }
}

Describe 'Import-BoardDayLines' {
    It 'replaces one day only, as one undo step, and records the sheet sync' {
        $b = New-Board -Days @($script:D1, $script:D2)
        Add-BoardLine -Board $b -Line @(
            New-BoardLine -ETD ([datetime]"$($script:D1) 07:00") -SortieType 'Day Local'
            New-BoardLine -ETD ([datetime]"$($script:D2) 07:00") -SortieType 'Day Local'
        ) -NoUndo
        $keep = Get-DayIds $b $script:D2
        $before = ConvertTo-BoardSnapshot $b
        $new = @(ConvertFrom-SheetRows -Rows @((New-SR $script:D1 '08:00'), (New-SR $script:D1 '09:00')) -Day $script:D1)
        $r = Import-BoardDayLines -Board $b -Day $script:D1 -Lines $new -Replace -Sync (New-SheetSyncEntry -Fingerprint 'abc')
        $r.Added | Should -Be 2
        $r.Removed | Should -Be 1
        (@(Get-BoardDayLines -Board $b -Day $script:D1) | ForEach-Object { $_.ETD.ToString('HHmm') }) -join ',' | Should -BeExactly '0800,0900'
        (Get-DayIds $b $script:D2) -join ',' | Should -BeExactly ($keep -join ',')
        $b.SheetSync[$script:D1].Fingerprint | Should -BeExactly 'abc'
        [void](Undo-Board $b)
        (ConvertTo-BoardSnapshot $b) | Should -BeExactly $before
        @(Get-BoardSheetSyncDays $b).Count | Should -Be 0
    }

    It 'without -Replace adds to the day, and adds a new day tab' {
        $b = New-Board -Days @($script:D1)
        Add-BoardLine -Board $b -Line @(New-BoardLine -ETD ([datetime]"$($script:D1) 07:00") -SortieType 'Day Local') -NoUndo
        [void](Import-BoardDayLines -Board $b -Day $script:D1 -Lines @(ConvertFrom-SheetRows -Rows @(New-SR $script:D1 '08:00') -Day $script:D1))
        @(Get-BoardDayLines -Board $b -Day $script:D1).Count | Should -Be 2
        [void](Import-BoardDayLines -Board $b -Day $script:D3 -Lines @(ConvertFrom-SheetRows -Rows @(New-SR $script:D3 '08:00') -Day $script:D3))
        ($b.Days.ToArray() -join ',') | Should -BeExactly "$($script:D1),$($script:D3)"
    }

    It 'clears links that pointed at a replaced line' {
        $b = New-Board -Days @($script:D1)
        $lead = New-BoardLine -ETD ([datetime]"$($script:D1) 07:00") -SortieType 'Day Local' -Student 'Adams, Al'
        $b.Lines.Add($lead)
        $wingNextDay = New-BoardLine -ETD ([datetime]"$($script:D2) 07:00") -SortieType 'Day Local' -Student 'Brown, Bea' -WingOfLineId $lead.Id
        $b.Lines.Add($wingNextDay); $b.Days.Add($script:D2)
        [void](Import-BoardDayLines -Board $b -Day $script:D1 -Lines @() -Replace)
        $wingNextDay.WingOfLineId | Should -BeExactly ''
    }
}

Describe 'Sheet sync state and banner' {
    BeforeAll {
        $script:Sheet = @(
            New-SR $script:D1 '07:30'
            New-SR $script:D1 '08:00'
            New-SR $script:D2 '07:30'
            New-SR $script:D3 '07:30'
            New-SR '2026-10-19' '07:30'
        )
        function New-SyncBoard {
            $b = New-Board -Days @($script:D1, $script:D2)
            return $b
        }
    }

    It 'an empty board day with sheet lines is EmptyDay; a day the sheet lacks is NoSheetLines' {
        $b = New-SyncBoard
        (Get-BoardSheetState -Board $b -SheetRows $script:Sheet -Day $script:D1).State | Should -BeExactly 'EmptyDay'
        (Get-BoardSheetState -Board $b -SheetRows $script:Sheet -Day '2026-10-23').State | Should -BeExactly 'NoSheetLines'
    }

    It 'after loading it is InSync; a sheet change is Outdated while the day is untouched, Changed once it is edited' {
        $b = New-SyncBoard
        $st = Get-BoardSheetState -Board $b -SheetRows $script:Sheet -Day $script:D1
        [void](Import-BoardDayLines -Board $b -Day $script:D1 -Lines @(ConvertFrom-SheetRows -Rows $script:Sheet -Day $script:D1) -Sync (New-SheetSyncEntry $st.Fingerprint))
        (Get-BoardSheetState -Board $b -SheetRows $script:Sheet -Day $script:D1).State | Should -BeExactly 'InSync'
        $b.SheetSync[$script:D1].BoardFingerprint | Should -BeExactly (Get-BoardDayFingerprint -Board $b -Day $script:D1)
        $edited = @($script:Sheet | ForEach-Object { $_.PSObject.Copy() })
        $edited[1].Instructor = 'Lee, Ann'
        $s2 = Get-BoardSheetState -Board $b -SheetRows $edited -Day $script:D1
        $s2.State | Should -BeExactly 'Outdated'
        $s2.SheetCount | Should -Be 2
        $s2.BoardCount | Should -Be 2
        $line = @(Get-BoardDayLines -Board $b -Day $script:D1)[0]
        Set-BoardInstructor -Board $b -Id $line.Id -Name 'Ortiz, Sam'
        (Get-BoardSheetState -Board $b -SheetRows $edited -Day $script:D1).State | Should -BeExactly 'Changed'
    }

    It 'the board fingerprint sees edits but not new line ids' {
        $b = New-SyncBoard
        [void](Import-BoardDayLines -Board $b -Day $script:D1 -Lines @(ConvertFrom-SheetRows -Rows $script:Sheet -Day $script:D1))
        $fp = Get-BoardDayFingerprint -Board $b -Day $script:D1
        $b2 = New-SyncBoard
        [void](Import-BoardDayLines -Board $b2 -Day $script:D1 -Lines @(ConvertFrom-SheetRows -Rows $script:Sheet -Day $script:D1))
        Get-BoardDayFingerprint -Board $b2 -Day $script:D1 | Should -BeExactly $fp
        Set-BoardLineField -Board $b2 -Id @(Get-BoardDayLines -Board $b2 -Day $script:D1)[0].Id -Field 'Notes' -Value 'x'
        Get-BoardDayFingerprint -Board $b2 -Day $script:D1 | Should -Not -BeExactly $fp
        Get-BoardDayFingerprint -Board $b2 -Day '2026-10-30' | Should -BeExactly ''
    }

    It 'a day built by hand is New until the user loads or keeps it' {
        $b = New-SyncBoard
        Add-BoardLine -Board $b -Line @(New-BoardLine -ETD ([datetime]"$($script:D1) 06:00") -SortieType 'Day Local') -NoUndo
        $st = Get-BoardSheetState -Board $b -SheetRows $script:Sheet -Day $script:D1
        $st.State | Should -BeExactly 'New'
        Set-BoardSheetSync -Board $b -Day $script:D1 -Entry (New-SheetSyncEntry -Fingerprint $st.Fingerprint -Action 'Ignored')
        (Get-BoardSheetState -Board $b -SheetRows $script:Sheet -Day $script:D1).State | Should -BeExactly 'InSync'
    }

    It 'the banner names the active day first, then other board days, then later sheet days' {
        $b = New-SyncBoard
        Add-BoardLine -Board $b -Line @(
            New-BoardLine -ETD ([datetime]"$($script:D1) 06:00") -SortieType 'Day Local'
            New-BoardLine -ETD ([datetime]"$($script:D2) 06:00") -SortieType 'Day Local'
        ) -NoUndo
        $i = Get-SheetBannerInfo -Board $b -SheetRows $script:Sheet -ActiveDay $script:D2
        $i.Day | Should -BeExactly $script:D2
        $i.State | Should -BeExactly 'New'
        $i.Text | Should -BeLike "*1 line(s) for Wed 21 Oct that are not on the board*"
        foreach ($d in @($script:D1, $script:D2)) {
            $st = Get-BoardSheetState -Board $b -SheetRows $script:Sheet -Day $d
            Set-BoardSheetSync -Board $b -Day $d -Entry (New-SheetSyncEntry -Fingerprint $st.Fingerprint -Action 'Ignored')
        }
        $o = Get-SheetBannerInfo -Board $b -SheetRows $script:Sheet -ActiveDay $script:D2
        $o.State | Should -BeExactly 'OtherDay'
        $o.Day | Should -BeExactly $script:D3 -Because 'Mon 19 Oct is before the board and is not suggested'
        $o.ActionText | Should -BeExactly 'Add Thu 22 Oct to the board'
        Get-SheetBannerInfo -Board $b -SheetRows $script:Sheet -ActiveDay $script:D2 -Ignore @("$($o.Day)|$($o.Fingerprint)") | Should -BeNullOrEmpty
    }

    It 'no sheet rows, no banner' {
        Get-SheetBannerInfo -Board (New-SyncBoard) -SheetRows @() | Should -BeNullOrEmpty
    }

    It 'the sync record is saved with the day, reopened, and dropped with the day' {
        $dir = Join-Path $TestDrive 'sync'
        $b = New-SyncBoard
        Set-BoardSheetSync -Board $b -Day $script:D1 -Entry (New-SheetSyncEntry -Fingerprint 'f1')
        Save-BoardState -Board $b -Dir $dir
        $re = Open-BoardState -Dir $dir -Days @($script:D1, $script:D2)
        $re.SheetSync[$script:D1].Fingerprint | Should -BeExactly 'f1'
        $re.SheetSync[$script:D1].Action | Should -BeExactly 'Loaded'
        @(Get-BoardSheetSyncDays $re) -join ',' | Should -BeExactly $script:D1
        Remove-BoardDay -Board $re -Day $script:D1
        @(Get-BoardSheetSyncDays $re).Count | Should -Be 0
        [void](Undo-Board $re)
        $re.SheetSync[$script:D1].Fingerprint | Should -BeExactly 'f1'
    }

    It 'a board file saved before sync existed still opens' {
        $dir = Join-Path $TestDrive 'oldfile'
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        [System.IO.File]::WriteAllText((Join-Path $dir "$($script:D1).json"), '{"SchemaVersion":1,"Day":"2026-10-20","Lines":[{"Id":"a1","ETD":"2026-10-20T07:30","SortieType":"Day Local","Instructor":"","Student":"","Mission":"","Callsign":"","Num":"","Notes":"","LinkedLineId":"","WingOfLineId":"","Source":"Manual"}]}')
        $re = Open-BoardState -Dir $dir -Days @($script:D1)
        $re.Lines.Count | Should -Be 1
        @(Get-BoardSheetSyncDays $re).Count | Should -Be 0
    }
}

Describe 'Loading a day from the last exported schedule' {
    BeforeAll {
        $script:Pub = @(
            New-PR "$($script:D2) 07:30" -Student 'Adams, Al' -Mission 'FF5201' -Notes 'Wingman: Brown, Bea' -Callsign 'HAZE' -Num 11
            New-PR "$($script:D2) 07:40" -IP 'Lee, Ann' -Student 'Brown, Bea' -Mission 'FF5201' -Notes 'Wingman to Adams, Al | Less than 3.5 hr TT' -Callsign 'HAZE' -Num 12
            New-PR "$($script:D2) 09:00" -IP 'Leech, Tom' -Notes 'No Available SP'
            New-PR "$($script:D2) 10:15" -Type 'O&B' -IP 'Ortiz, Sam' -Student 'Cruz, Cy' -Mission 'FT5102'
            New-PR "$($script:D2) 13:15" -Type 'O&B Return' -IP 'Ortiz, Sam' -Student 'Cruz, Cy' -Mission 'FT5103'
            New-PR "$($script:D2) 13:00" -Type 'IFT' -IP 'SIM IP' -Student 'Dunn, Di' -Mission 'FN3201'
            New-PR "$($script:D3) 07:30" -Student 'Dunn, Di' -Mission 'FT5104'
        )
        $script:L = @(ConvertFrom-PublishedRows -Rows $script:Pub -Day $script:D2)
        function Get-L { param ([string]$Hhmm) $script:L | Where-Object { $_.ETD.ToString('HHmm') -eq $Hhmm } }
    }

    It 'brings back that day only' {
        $script:L.Count | Should -Be 6
        @(Get-PublishedDays -Rows $script:Pub) -join ',' | Should -BeExactly "$($script:D2),$($script:D3)"
    }

    It 'keeps every student and mission hardcoded where they were' {
        (Get-L '0730').Student | Should -BeExactly 'Adams, Al'
        (Get-L '0730').Mission | Should -BeExactly 'FF5201'
        (Get-L '1015').Student | Should -BeExactly 'Cruz, Cy'
        (Get-L '1300').Student | Should -BeExactly 'Dunn, Di'
        (Get-L '1300').Instructor | Should -BeExactly '' -Because 'SIM IP is the board''s blank sim IP'
    }

    It '-TargetDay puts the day''s lines on another day' {
        $moved = @(ConvertFrom-PublishedRows -Rows $script:Pub -Day $script:D2 -TargetDay $script:D3)
        $moved.Count | Should -Be 6
        @($moved | Where-Object { $_.Day -ne $script:D3 }).Count | Should -Be 0
    }

    It 'For Sale comes back as an open line' {
        (Get-L '0900').Student | Should -BeExactly ''
        (Get-L '0900').Instructor | Should -BeExactly 'Leech, Tom'
    }

    It 'pairs the wingman with the lead and links the O&B with its return' {
        (Get-L '0740').WingOfLineId | Should -BeExactly (Get-L '0730').Id
        (Get-L '1015').LinkedLineId | Should -BeExactly (Get-L '1315').Id
        (Get-L '1315').Student | Should -BeExactly '' -Because 'return legs follow their O&B'
    }

    It 'leaves notes and callsigns for the next fill to write' {
        @($script:L | Where-Object { $_.Notes -or $_.Callsign -or $_.Num }).Count | Should -Be 0
    }
}

Describe 'Get-PublishedContextRows' {
    BeforeAll {
        $script:Pub = @(
            New-PR "$($script:D1) 07:30" -Student 'Adams, Al' -Mission 'FT5001'
            New-PR "$($script:D2) 07:30" -Student 'Adams, Al' -Mission 'FF5201' -Notes 'Wingman: Brown, Bea'
            New-PR "$($script:D2) 07:40" -IP 'Lee, Ann' -Student 'Brown, Bea' -Mission 'FF5201' -Notes 'Wingman to Adams, Al | x'
            New-PR "$($script:D2) 09:00" -Student 'For Sale'
            New-PR "$($script:D2) 09:30" -Student 'Need Form SP'
            New-PR "$($script:D2) 10:15" -Type 'O&B' -IP 'Ortiz, Sam' -Student 'Cruz, Cy' -Mission 'FT5102'
            New-PR "$($script:D2) 13:15" -Type 'O&B Return' -IP 'Ortiz, Sam' -Student 'Cruz, Cy' -Mission 'FT5103'
            New-PR "$($script:D3) 07:30" -Student 'Dunn, Di' -Mission 'FT5104'
        )
    }

    It 'takes exported days on/after -FromDay that are not being filled' {
        $c = @(Get-PublishedContextRows -Rows $script:Pub -FromDay $script:D1 -ExcludeDays @($script:D2))
        (@($c | ForEach-Object { ([datetime]$_.ETD).ToString('MMdd HHmm') }) -join ',') | Should -BeExactly '1020 0730,1022 0730'
        @(Get-PublishedContextRows -Rows $script:Pub -FromDay $script:D2 -ExcludeDays @($script:D2)).Count | Should -Be 1
    }

    It 'leaves out For Sale / direct support, blanks return legs and keeps unique ids' {
        $c = @(Get-PublishedContextRows -Rows $script:Pub -FromDay $script:D2 -ExcludeDays @($script:D3))
        $c.Count | Should -Be 4
        ($c | Where-Object { $_.SortieType -eq 'O&B Return' }).Student | Should -BeNullOrEmpty
        ($c | Where-Object { $_.SortieType -eq 'O&B Return' }).Mission | Should -BeNullOrEmpty
        @($c | ForEach-Object { $_.BoardLineId } | Select-Object -Unique).Count | Should -Be 4
    }

    It 'keeps the wing pairing so the lead does not take a new wingman' {
        $c = @(Get-PublishedContextRows -Rows $script:Pub -FromDay $script:D2 -ExcludeDays @($script:D3))
        $lead = $c | Where-Object { (Get-NameText $_.Student) -eq 'Adams, Al' }
        $wing = $c | Where-Object { (Get-NameText $_.Student) -eq 'Brown, Bea' }
        $wing.WingOfLineId | Should -BeExactly $lead.BoardLineId
        $wing.Notes | Should -BeExactly 'Wingman to Adams, Al'
        $lead.Notes | Should -BeNullOrEmpty
    }
}

Describe 'Filling one day with the exported days as context' {
    BeforeAll {
        $script:Sw = New-TestSoftWeights
        $done = @('FI3101|2026-10-15 08:00')
        # Wednesday is already exported: Adams flies FT5001 then
        $script:Pub = @(New-PR "$($script:D2) 07:30" -Student 'Adams, Al' -Mission 'FT5001')
        $script:Board = New-Board -Days @($script:D1)
        $script:Open = New-BoardLine -ETD ([datetime]"$($script:D1) 09:00") -SortieType 'Day Local' -Instructor 'Jones, Bob'
        Add-BoardLine -Board $script:Board -Line @($script:Open) -NoUndo

        $students = @(New-TestStudent -Name 'Adams, Al' -Done $done)
        $rows = @(ConvertTo-ScheduleRows -Board $script:Board -Days @($script:D1)) + @(Get-PublishedContextRows -Rows $script:Pub -FromDay $script:D1 -ExcludeDays @($script:D1))
        $fill = Invoke-FillSchedule -Schedule $rows -Students $students -Instructors $script:IPs -SoftWeights $script:Sw 6>$null
        $script:R = Split-FillResult -FillResult $fill -Days @($script:D1)

        $alone = @(New-TestStudent -Name 'Adams, Al' -Done $done)
        $fillAlone = Invoke-FillSchedule -Schedule @(ConvertTo-ScheduleRows -Board $script:Board -Days @($script:D1)) -Students $alone -Instructors $script:IPs -SoftWeights $script:Sw 6>$null
        $script:Alone = Split-FillResult -FillResult $fillAlone -Days @($script:D1)
    }

    It 'without context the open Tuesday line would take Wednesday''s event' {
        $row = $script:Alone.Schedule[0]
        (Get-NameText $row.Student) | Should -BeExactly 'Adams, Al'
        (Get-NameText $row.Mission) | Should -BeExactly 'FT5001'
    }

    It 'with context Tuesday does not double-book Wednesday''s event' {
        $row = $script:R.Schedule[0]
        "$(Get-NameText $row.Student)|$(Get-NameText $row.Mission)" | Should -Not -BeExactly 'Adams, Al|FT5001'
    }

    It 'the exported Wednesday line stays as it was, with no override' {
        $ctx = $script:R.Context[0]
        (Get-NameText $ctx.Student) | Should -BeExactly 'Adams, Al'
        (Get-NameText $ctx.Mission) | Should -BeExactly 'FT5001'
        $ctx.FilledBy | Should -BeExactly 'Pin'
        @($script:R.ContextOverrides).Count | Should -Be 0
    }

    It 'the students carry Wednesday''s event, so their tabs match the exported schedule' {
        $m = $script:R.Students[0].Syllabus | Where-Object { $_.Mission -eq 'FT5001' }
        Test-MissionComplete $m | Should -BeTrue
        (Parse-ETD $m.ETD).ToString('yyyy-MM-dd') | Should -BeExactly $script:D2
    }

    It 'Split-FillResult counts the filled days only' {
        $script:R.Stats.Lines | Should -Be 1
        $script:R.Stats.Pinned | Should -Be 0
        @($script:R.Schedule).Count | Should -Be 1
        @($script:R.Context).Count | Should -Be 1
        ($script:R.Days -join ',') | Should -BeExactly $script:D1
    }
}

Describe 'Test-BoardReadyToFill -Days' {
    It 'only checks the days being filled' {
        $b = New-Board -Days @($script:D1, $script:D2)
        Add-BoardLine -Board $b -Line @(New-BoardLine -ETD ([datetime]"$($script:D1) 07:00") -SortieType 'Day Local' -Instructor 'Jones, Bob') -NoUndo
        @(Test-BoardReadyToFill -Board $b -Instructors $script:IPs) -join '|' | Should -BeLike "*$($script:D2) has no lines*"
        @(Test-BoardReadyToFill -Board $b -Instructors $script:IPs -Days @($script:D1)).Count | Should -Be 0
    }
}
