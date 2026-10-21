# Pester 5 tests for the puck board's one-day-at-a-time logic (Modules\PuckBoard.psm1) without
# a window: the WPF-only helpers are mocked and the controls are plain objects. The real window
# is driven by Tests\Invoke-SelfTest.ps1 on Windows.
# Dates: Tue 2026-10-20 (D1), Wed 2026-10-21 (D2).

BeforeAll {
    . "$PSScriptRoot/Fixtures/New-FaisTestData.ps1"
    Import-FaisModules -IncludeUI

    $script:D1 = '2026-10-20'
    $script:D2 = '2026-10-21'

    function New-SR {
        param ([string]$Day, [string]$Time, [string]$Type = 'Day Local', [string]$IP = 'Jones, Bob', [string]$Student = $null)
        [PSCustomObject]@{
            ETD = "$Day $($Time):00"; Day = $Day; Time = $Time; TimeOnly = $false; Instructor = $IP; SortieType = $Type
            Student = $(if ($Student) { $Student } else { $null }); Mission = $null; Callsign = $null; Num = $null; Notes = $null; SheetRow = 1
        }
    }

    function New-Sheet {
        param ([array]$Rows, [string[]]$Skipped = @())
        [PSCustomObject]@{ Path = ''; Worksheet = 'Schedule'; Found = $true; Error = ''; Layout = 'test'; HeaderRow = 1; Rows = $Rows; Skipped = $Skipped; ReadAt = (Get-Date) }
    }

    # A board state with plain objects standing in for the WPF controls
    function Initialize-TestBoard {
        param ([string[]]$Days, $Sheet, [array]$PrevSchedule = @(), [string]$Active)

        $students = @(
            New-TestStudent -Name 'Adams, Al' -Done 'FI3101|2026-10-15 08:00'
            New-TestStudent -Name 'Brown, Bea' -Done 'FI3101|2026-10-15 08:00'
        )
        $data = @{ Students = $students; Instructors = (New-TestInstructors); SoftWeights = (New-TestSoftWeights); Formations = (New-TestFormations)
            PrevSchedule = $PrevSchedule; ScheduleSheet = $Sheet; MasterDataPath = ''; OutputPath = '' }
        $board = New-Board -Days $Days
        InModuleScope PuckBoard -Parameters @{ Board = $board; Data = $data; Active = $(if ($Active) { $Active } else { $Days[0] }) } {
            param ($Board, $Data, $Active)
            $o = { param($h) [PSCustomObject]$h }
            $script:PB = @{
                Window = $null; Board = $Board; Data = $Data; Settings = (Get-DefaultBoardSettings)
                C = @{
                    SheetBanner = (& $o @{ Visibility = 'Collapsed' }); SheetBannerText = (& $o @{ Text = '' })
                    SheetBannerLoadButton = (& $o @{ Content = '' }); SheetBannerIgnoreButton = (& $o @{ Content = '' })
                    FillDayOnlyCheck = (& $o @{ IsChecked = $false; Content = ''; Visibility = 'Visible' })
                    FillButton = (& $o @{ Content = ''; Visibility = 'Visible' }); BackToEditButton = (& $o @{ Visibility = 'Collapsed' })
                    RefillButton = (& $o @{ Visibility = 'Collapsed' }); ExportButton = (& $o @{ Visibility = 'Collapsed' })
                    ModeBanner = (& $o @{ Visibility = 'Collapsed' }); ModeBannerText = (& $o @{ Text = '' })
                    StatusText = (& $o @{ Text = '' }); CountsText = (& $o @{ Text = '' })
                }
                DayColls = @{}; Assess = @{}; RowAssess = @{}; Queue = (New-Object System.Collections.Generic.List[string])
                Commands = (New-Object System.Collections.Generic.List[object]); Drag = $null; Heat = $null; Mode = 'Edit'; Overlay = $null
                FillResult = $null; Result = $null; Busy = $false; Draining = $false; Filling = $false; LastFocusDay = $Active
                FillDays = @(); SheetIgnore = (New-Object System.Collections.Generic.List[string]); SheetBannerInfo = $null
                SheetStamp = ''; SheetPending = ''; SheetAutoPending = $false; ActiveDay = $Active; StatusLog = (New-Object System.Collections.Generic.List[string])
                SaveTimer = (& $o @{ }); SheetTimer = (& $o @{ })
            }
            $script:PB.SheetTimer | Add-Member -MemberType ScriptMethod -Name Stop -Value { }
        }
    }

    function Get-PB { InModuleScope PuckBoard { $script:PB } }
}

Describe 'Puck board, one day at a time (no window)' {
    BeforeEach {
        # The WPF-only pieces
        Mock -ModuleName PuckBoard Update-PBTabs { }
        Mock -ModuleName PuckBoard Update-PBDay { }
        Mock -ModuleName PuckBoard Update-PBPools { }
        Mock -ModuleName PuckBoard Restore-PBFocus { }
        Mock -ModuleName PuckBoard Get-PBFocusState { $null }
        Mock -ModuleName PuckBoard Request-PBAutosave { }
        Mock -ModuleName PuckBoard Start-PBAssessQueue { }
        Mock -ModuleName PuckBoard Show-BusyOverlay { }
        Mock -ModuleName PuckBoard Hide-BusyOverlay { }
        Mock -ModuleName PuckBoard Invoke-PBDoEvents { }
        Mock -ModuleName PuckBoard Close-PBPopup { }
        # Mock bodies run in the test's scope; the board state lives in the module's
        Mock -ModuleName PuckBoard Get-PBActiveDay { & (Get-Module PuckBoard) { $script:PB.ActiveDay } }
        Mock -ModuleName PuckBoard Select-PBDay { & (Get-Module PuckBoard) { param($d) $script:PB.ActiveDay = $d } $Day }
        Mock -ModuleName PuckBoard Set-PBStatus { & (Get-Module PuckBoard) { param($t, $k) $script:PB.C.StatusText.Text = $t; $script:PB.StatusLog.Add("$k|$t") } $Text $Kind }
    }

    It 'opening the board loads the sheet onto an empty day and leaves the banner off' {
        Initialize-TestBoard -Days @($script:D1) -Sheet (New-Sheet @((New-SR $script:D1 '07:30' -Student 'Adams, Al'), (New-SR $script:D1 '09:00')))
        $msg = InModuleScope PuckBoard { Invoke-PBSheetStartup }
        $pb = Get-PB
        @(Get-BoardDayLines -Board $pb.Board -Day $script:D1).Count | Should -Be 2
        (@(Get-BoardDayLines -Board $pb.Board -Day $script:D1))[0].Student | Should -BeExactly 'Adams, Al'
        $msg | Should -BeLike 'Loaded Tue 20 Oct (2 line(s)) from MasterData*'
        $pb.C.SheetBanner.Visibility | Should -BeExactly 'Collapsed'
        $pb.SheetBannerInfo | Should -BeNullOrEmpty
    }

    It 'a saved sheet change updates an untouched day by itself, but only asks for an edited one' {
        $rows = @((New-SR $script:D1 '07:30'), (New-SR $script:D1 '09:00'))
        Initialize-TestBoard -Days @($script:D1) -Sheet (New-Sheet $rows)
        [void](InModuleScope PuckBoard { Invoke-PBSheetStartup })

        $pb = Get-PB
        $pb.Data.ScheduleSheet = New-Sheet @((New-SR $script:D1 '07:45'), (New-SR $script:D1 '09:00'))
        $done = @(InModuleScope PuckBoard { Invoke-PBSheetAutoUpdate })
        $done.Count | Should -Be 1
        (@(Get-BoardDayLines -Board $pb.Board -Day $script:D1) | ForEach-Object { $_.ETD.ToString('HHmm') }) -join ',' | Should -BeExactly '0745,0900'

        $l = @(Get-BoardDayLines -Board $pb.Board -Day $script:D1)[0]
        Set-BoardInstructor -Board $pb.Board -Id $l.Id -Name 'Lee, Ann'
        $pb.Data.ScheduleSheet = New-Sheet @((New-SR $script:D1 '08:00'), (New-SR $script:D1 '09:00'))
        @(InModuleScope PuckBoard { Invoke-PBSheetAutoUpdate }).Count | Should -Be 0
        InModuleScope PuckBoard { Update-PBSheetBanner }
        $pb.C.SheetBanner.Visibility | Should -BeExactly 'Visible'
        $pb.SheetBannerInfo.State | Should -BeExactly 'Changed'
        $pb.C.SheetBannerLoadButton.Content | Should -BeExactly 'Replace Tue 20 Oct with the sheet'
        (@(Get-BoardDayLines -Board $pb.Board -Day $script:D1))[0].Instructor | Should -BeExactly 'Lee, Ann' -Because 'the edited day was not replaced'

        InModuleScope PuckBoard { Invoke-PBSheetBannerLoad }
        (@(Get-BoardDayLines -Board $pb.Board -Day $script:D1) | ForEach-Object { $_.ETD.ToString('HHmm') + ' ' + $_.Instructor }) -join ',' | Should -BeExactly '0800 Jones, Bob,0900 Jones, Bob'
        [void](Undo-Board $pb.Board)
        (@(Get-BoardDayLines -Board $pb.Board -Day $script:D1))[0].Instructor | Should -BeExactly 'Lee, Ann' -Because 'Ctrl+Z takes the replacement back'
    }

    It 'the sheet tick reads, then auto-updates only in edit mode and when nobody is typing in a row' {
        Initialize-TestBoard -Days @($script:D1) -Sheet (New-Sheet @(New-SR $script:D1 '07:30'))
        [void](InModuleScope PuckBoard { Invoke-PBSheetStartup })
        $pb = Get-PB
        # A MasterData file that has not changed since it was read: the tick only applies what is pending
        $md = Join-Path $TestDrive 'MasterData.xlsx'
        [System.IO.File]::WriteAllText($md, 'x')
        $pb.Data.MasterDataPath = $md
        $pb.SheetStamp = InModuleScope PuckBoard -Parameters @{ P = $md } { param($P) Get-PBFileStamp $P }
        $pb.Data.ScheduleSheet = New-Sheet @(New-SR $script:D1 '07:40')
        $pb.SheetAutoPending = $true
        $pb.Mode = 'Review'
        InModuleScope PuckBoard { Invoke-PBSheetTick }
        (@(Get-BoardDayLines -Board $pb.Board -Day $script:D1))[0].ETD.ToString('HHmm') | Should -BeExactly '0730' -Because 'review mode waits'
        $pb.Mode = 'Edit'
        Mock -ModuleName PuckBoard Get-PBFocusState { @{ LineId = 'x'; Tag = 'EtdBox' } }
        InModuleScope PuckBoard { Invoke-PBSheetTick }
        (@(Get-BoardDayLines -Board $pb.Board -Day $script:D1))[0].ETD.ToString('HHmm') | Should -BeExactly '0730' -Because 'typing in a row waits'
        Mock -ModuleName PuckBoard Get-PBFocusState { $null }
        InModuleScope PuckBoard { Invoke-PBSheetTick }
        (@(Get-BoardDayLines -Board $pb.Board -Day $script:D1))[0].ETD.ToString('HHmm') | Should -BeExactly '0740'
        $pb.SheetAutoPending | Should -BeFalse
        $pb.C.StatusText.Text | Should -BeLike 'MasterData was saved: updated Tue 20 Oct*'
    }

    It 'the banner offers a later sheet day, and "Not now" hides it until that day changes' {
        Initialize-TestBoard -Days @($script:D1) -Sheet (New-Sheet @(New-SR $script:D2 '07:30'))
        [void](InModuleScope PuckBoard { Invoke-PBSheetStartup })
        $pb = Get-PB
        $pb.SheetBannerInfo.State | Should -BeExactly 'OtherDay'
        $pb.C.SheetBannerIgnoreButton.Content | Should -BeExactly 'Not now'
        InModuleScope PuckBoard { Invoke-PBSheetBannerIgnore }
        $pb.C.SheetBanner.Visibility | Should -BeExactly 'Collapsed'
        $pb.Data.ScheduleSheet = New-Sheet @((New-SR $script:D2 '07:30'), (New-SR $script:D2 '08:00'))
        InModuleScope PuckBoard { Update-PBSheetBanner }
        $pb.C.SheetBanner.Visibility | Should -BeExactly 'Visible'
        InModuleScope PuckBoard { Invoke-PBSheetBannerLoad }
        ($pb.Board.Days.ToArray() -join ',') | Should -BeExactly "$($script:D1),$($script:D2)"
        $pb.ActiveDay | Should -BeExactly $script:D2 -Because 'the loaded day is brought on screen'
    }

    It 'loading an exported day puts everyone back hardcoded and reports it' {
        $pub = @(
            [PSCustomObject]@{ ETD = "$($script:D1) 07:30:00"; Instructor = 'Jones, Bob'; SortieType = 'Day Local'; Student = 'Adams, Al'; Mission = 'FT5001'; Callsign = 'TALON'; Num = 31; Notes = '' }
            [PSCustomObject]@{ ETD = "$($script:D1) 09:00:00"; Instructor = 'Lee, Ann'; SortieType = 'Day Local'; Student = 'For Sale'; Mission = ''; Callsign = 'TALON'; Num = 32; Notes = 'No Available SP' }
        )
        Initialize-TestBoard -Days @($script:D1) -Sheet $null -PrevSchedule $pub
        $res = InModuleScope PuckBoard { Invoke-PBLoadDay -Source 'Exported' -SourceDay '2026-10-20' -TargetDay '2026-10-20' -Replace -KeepStudents }
        $res.Added | Should -Be 2
        $res.Students | Should -Be 1
        $pb = Get-PB
        $pb.C.StatusText.Text | Should -BeLike 'Loaded 2 line(s) onto Tue 20 Oct from the last exported schedule, 1 with students hardcoded*'
        @(Get-BoardSheetSyncDays $pb.Board).Count | Should -Be 0 -Because 'the exported schedule is not the sheet'
    }

    It 'a sheet day with unreadable rows says so when loaded' {
        Initialize-TestBoard -Days @($script:D1) -Sheet (New-Sheet @(New-SR $script:D1 '07:30') -Skipped @('row 9: no sortie type'))
        [void](InModuleScope PuckBoard { Invoke-PBLoadDay -Source 'Sheet' -SourceDay '2026-10-20' -TargetDay '2026-10-20' -Replace -KeepStudents })
        $pb = Get-PB
        $pb.StatusLog[-1] | Should -BeLike 'Warn|*1 sheet row(s) could not be read*'
    }

    It '"Only this day" fills and reviews just the day on screen, with exported days as context' {
        $pub = @([PSCustomObject]@{ ETD = "$($script:D2) 07:30:00"; Instructor = 'Jones, Bob'; SortieType = 'Day Local'; Student = 'Adams, Al'; Mission = 'FT5001'; Callsign = 'TALON'; Num = 31; Notes = '' })
        Initialize-TestBoard -Days @($script:D1, $script:D2) -Sheet $null -PrevSchedule $pub -Active $script:D1
        $pb = Get-PB
        Add-BoardLine -Board $pb.Board -Line @(
            New-BoardLine -ETD ([datetime]"$($script:D1) 09:00") -SortieType 'Day Local' -Instructor 'Jones, Bob'
            New-BoardLine -ETD ([datetime]"$($script:D2) 10:00") -SortieType 'Day Local' -Instructor 'Lee, Ann'
        ) -NoUndo
        $pb.C.FillDayOnlyCheck.IsChecked = $true
        InModuleScope PuckBoard { Update-PBFillScope }
        $pb.C.FillButton.Content | Should -BeExactly 'Fill this day'
        $pb.C.FillDayOnlyCheck.Content | Should -BeExactly 'Only Tue 20 Oct'

        InModuleScope PuckBoard { Invoke-PBFill } 6>$null
        $pb.Mode | Should -BeExactly 'Review'
        ($pb.FillDays -join ',') | Should -BeExactly $script:D1
        @($pb.FillResult.Schedule).Count | Should -Be 1
        @($pb.FillResult.Context).Count | Should -Be 1 -Because 'the exported Wednesday rode along'
        $row = $pb.FillResult.Schedule[0]
        "$(Get-NameText $row.Student)|$(Get-NameText $row.Mission)" | Should -Not -BeExactly 'Adams, Al|FT5001'
        $pb.C.ModeBannerText.Text | Should -BeLike 'REVIEW (Tue 20 Oct)*Export writes Tue 20 Oct*'
        $pb.C.FillDayOnlyCheck.Visibility | Should -BeExactly 'Collapsed'

        # Re-fill repeats the same day even after switching tabs
        $pb.ActiveDay = $script:D2
        InModuleScope PuckBoard { Invoke-PBFill -Refill } 6>$null
        ($pb.FillDays -join ',') | Should -BeExactly $script:D1

        InModuleScope PuckBoard { Exit-PBReview }
        $pb.C.FillDayOnlyCheck.IsChecked = $false
        InModuleScope PuckBoard { Invoke-PBFill } 6>$null
        ($pb.FillDays -join ',') | Should -BeExactly "$($script:D1),$($script:D2)"
        @($pb.FillResult.Context).Count | Should -Be 0 -Because 'Wednesday is filled now, not context'
    }

    It 'Format-PBDays names one day or a range' {
        InModuleScope PuckBoard { Format-PBDays @('2026-10-21') } | Should -BeExactly 'Wed 21 Oct'
        InModuleScope PuckBoard { Format-PBDays @('2026-10-22', '2026-10-20', '2026-10-21') } | Should -BeExactly 'Tue 20 Oct - Thu 22 Oct (3 days)'
        InModuleScope PuckBoard { Format-PBDays @() } | Should -BeExactly 'no days'
    }
}


Describe 'Puck board: missions picked by hand, Add Bulk Lines, Update Syllabi (no window)' {
    BeforeEach {
        Mock -ModuleName PuckBoard Update-PBTabs { }
        Mock -ModuleName PuckBoard Update-PBDay { }
        Mock -ModuleName PuckBoard Update-PBPools { }
        Mock -ModuleName PuckBoard Restore-PBFocus { }
        Mock -ModuleName PuckBoard Get-PBFocusState { $null }
        Mock -ModuleName PuckBoard Request-PBAutosave { }
        Mock -ModuleName PuckBoard Start-PBAssessQueue { }
        Mock -ModuleName PuckBoard Close-PBPopup { }
        Mock -ModuleName PuckBoard Show-PBPlacementPopup { }
        Mock -ModuleName PuckBoard Get-PBActiveDay { & (Get-Module PuckBoard) { $script:PB.ActiveDay } }
        Mock -ModuleName PuckBoard Set-PBStatus { & (Get-Module PuckBoard) { param($t, $k) $script:PB.C.StatusText.Text = $t; $script:PB.StatusLog.Add("$k|$t") } $Text $Kind }
        Clear-AssessmentCache

        # Adams has flown FI3101 only: opted for FT5001 (Day Local) and FE3102 (OFT), not FT5102
        Initialize-TestBoard -Days @($script:D1) -Sheet (New-Sheet @())
        $script:Board = (Get-PB).Board
        Add-BoardLine -Board $script:Board -Line @(New-BoardLine -ETD ([datetime]"$($script:D1) 07:30") -SortieType 'Day Local' -Instructor 'Jones, Bob' -Student 'Adams, Al' -Id 'L1') -NoUndo

        function Invoke-Choose {
            param ([string]$Choice)
            & (Get-Module PuckBoard) { param($c) Invoke-PBChooseMission -LineId 'L1' -Choice $c } $Choice
        }
        function Get-L1Mission { [string](Get-BoardLine $script:Board 'L1').Mission }
        function Get-LastStatus { @((Get-PB).StatusLog)[-1] }
    }

    It 'a mission the student is not opted for: Yes puts the recommended mission on instead' {
        Mock -ModuleName PuckBoard Show-PBMessage { 'Yes' }
        Invoke-Choose 'FT5102 (not opted)'
        Get-L1Mission | Should -BeExactly 'FT5001'
        Should -Invoke -ModuleName PuckBoard Show-PBMessage -Times 1 -Exactly -ParameterFilter {
            $Title -eq 'Check the mission for Adams, Al' -and $Buttons -eq 'YesNoCancel' -and $Text -like "Adams, Al isn't opted for FT5102.*Not opted for FT5102: FT5001 isn't complete yet*"
        }
        Get-LastStatus | Should -BeExactly '|Mission set to FT5001.'
    }

    It 'No keeps the mission the user picked, flagged' {
        Mock -ModuleName PuckBoard Show-PBMessage { 'No' }
        Invoke-Choose 'FT5102 (not opted)'
        Get-L1Mission | Should -BeExactly 'FT5102'
        Get-LastStatus | Should -BeExactly 'Warn|FT5102 kept - it is flagged OVERRIDE on the output.'
        $a = Get-PlacementAssessment -Board $script:Board -LineId 'L1' -StudentName 'Adams, Al' -MissionName 'FT5102' -Data (Get-PB).Data
        $a.Status | Should -BeExactly 'Hard'
        @($a.Violations | ForEach-Object { $_.Rule }) | Should -Contain 'NOT_OPTED'
    }

    It 'Cancel changes nothing' {
        Mock -ModuleName PuckBoard Show-PBMessage { 'Cancel' }
        $v = $script:Board.Version
        Invoke-Choose 'FT5102 (not opted)'
        Get-L1Mission | Should -BeExactly ''
        $script:Board.Version | Should -Be $v
        Get-LastStatus | Should -BeExactly '|Mission left as it was.'
        Should -Invoke -ModuleName PuckBoard Update-PBDay -Times 1 -Because 'the dropdown is put back'
    }

    It 'the right mission goes straight on, no questions' {
        Mock -ModuleName PuckBoard Show-PBMessage { throw 'should not ask' }
        Invoke-Choose 'FT5001'
        Get-L1Mission | Should -BeExactly 'FT5001'
        Invoke-Choose '(auto)'
        Get-L1Mission | Should -BeExactly ''
        Should -Invoke -ModuleName PuckBoard Show-PBMessage -Times 0 -Exactly
    }

    It 'Type a mission...: takes what is typed, in the syllabus spelling' {
        Mock -ModuleName PuckBoard Read-PBText { ' ft5001 ' }
        Mock -ModuleName PuckBoard Show-PBMessage { throw 'should not ask' }
        Invoke-Choose 'Type a mission...'
        Get-L1Mission | Should -BeExactly 'FT5001'
    }

    It 'Type a mission...: a mission not in the syllabus is explained, and kept on No' {
        Mock -ModuleName PuckBoard Read-PBText { 'XX1234' }
        Mock -ModuleName PuckBoard Show-PBMessage { 'No' }
        Invoke-Choose 'Type a mission...'
        Get-L1Mission | Should -BeExactly 'XX1234'
        Should -Invoke -ModuleName PuckBoard Show-PBMessage -Times 1 -Exactly -ParameterFilter { $Text -like "XX1234 isn't in Adams, Al's syllabus.*Yes = put FT5001 on the line*" }
    }

    It 'Type a mission...: cancelling the prompt changes nothing' {
        Mock -ModuleName PuckBoard Read-PBText { '' }
        Mock -ModuleName PuckBoard Show-PBMessage { throw 'should not ask' }
        $v = $script:Board.Version
        Invoke-Choose 'Type a mission...'
        Get-L1Mission | Should -BeExactly ''
        $script:Board.Version | Should -Be $v
    }

    It 'the dropdown lists the missions the student is not opted for' {
        $texts = @(& (Get-Module PuckBoard) { param($b) Get-PBMissionChoices (Get-BoardLine $b 'L1') } $script:Board)
        $texts[0] | Should -BeExactly '(auto)'
        $texts[1] | Should -BeExactly 'FT5001'
        $texts | Should -Contain 'FT5102 (not opted)'
        $texts[-1] | Should -BeExactly 'Type a mission...'
    }

    It 'Add Bulk Lines puts the lines on the day as one undo step' {
        $ok = & (Get-Module PuckBoard) { Add-PBBulkLines -Day '2026-10-20' -StartTime '0800' -Count '5' }
        $ok | Should -BeTrue
        (@(Get-BoardDayLines -Board $script:Board -Day $script:D1) | Where-Object { $_.Id -ne 'L1' } | ForEach-Object { $_.ETD.ToString('HHmm') }) -join ',' | Should -BeExactly '0800,0803,0806,0809,0812'
        Get-LastStatus | Should -BeExactly '|Added 5 Day Local lines on Tue 20 Oct: 0800 to 0812, every 3 min.'
        Undo-Board $script:Board | Should -BeTrue
        @(Get-BoardDayLines -Board $script:Board -Day $script:D1).Count | Should -Be 1
    }

    It 'Add Bulk Lines with a bad box says so and adds nothing' {
        $v = $script:Board.Version
        $ok = & (Get-Module PuckBoard) { Add-PBBulkLines -Day '2026-10-20' -StartTime '0800' -Count 'lots' }
        $ok | Should -BeFalse
        $script:Board.Version | Should -Be $v
        Get-LastStatus | Should -BeExactly 'Warn|Number of lines must be a whole number from 1 to 60.'
    }

    It 'the Add Bulk Lines button uses what the dialog asked for, and nothing when it is cancelled' {
        Mock -ModuleName PuckBoard Show-PBBulkLinesDialog { @{ StartTime = '1000'; Count = '2'; Every = '10'; SortieType = 'O&B' } }
        & (Get-Module PuckBoard) { Invoke-PBBulkLines }
        $new = @(Get-BoardDayLines -Board $script:Board -Day $script:D1 | Where-Object { $_.Id -ne 'L1' })
        ($new | ForEach-Object { $_.ETD.ToString('HHmm') + ' ' + $_.SortieType }) -join ',' | Should -BeExactly '1000 O&B,1010 O&B,1300 O&B Return,1310 O&B Return'
        Should -Invoke -ModuleName PuckBoard Show-PBBulkLinesDialog -Times 1 -Exactly -ParameterFilter { $Day -eq '2026-10-20' }

        Mock -ModuleName PuckBoard Show-PBBulkLinesDialog { $null }
        $v = $script:Board.Version
        & (Get-Module PuckBoard) { Invoke-PBBulkLines }
        $script:Board.Version | Should -Be $v
    }

    It "the dialog starts just after the day's last line (returns don't count), else at 0730" {
        $d = & (Get-Module PuckBoard) { Get-PBBulkDefaults -Day '2026-10-21' }
        $d.StartTime | Should -BeExactly '0730'
        $d.Every | Should -BeExactly '3'
        $d.Count | Should -BeExactly ''
        $qe = ConvertFrom-QuickEntry -Text '1015 O&B' -Day $script:D1 -Instructors (New-TestInstructors) -Students @() -Board $script:Board
        Add-BoardLine -Board $script:Board -Line $qe.Lines -NoUndo
        (& (Get-Module PuckBoard) { Get-PBBulkDefaults -Day '2026-10-20' }).StartTime | Should -BeExactly '1018'
    }

    It 'Format-PBBulkTimes describes the run' {
        & (Get-Module PuckBoard) { Format-PBBulkTimes -Times @('0730', '0733', '0736') -Every '3' -SortieType 'Day Local' } | Should -BeExactly '3 Day Local lines: 0730 to 0736, every 3 min'
        & (Get-Module PuckBoard) { Format-PBBulkTimes -Times @('0900') -Every '3' -SortieType 'OFT' -DayText 'Tue 20 Oct' } | Should -BeExactly '1 OFT line on Tue 20 Oct at 0900'
        & (Get-Module PuckBoard) { Format-PBBulkTimes -Times @('1000', '1005') -Every '5' -SortieType 'O&B' } | Should -BeExactly '2 O&B lines: 1000 to 1005, every 5 min (each with its return leg)'
    }

    Context 'Update Syllabi' {
        BeforeEach {
            Mock -ModuleName PuckBoard Show-BusyOverlay { }
            Mock -ModuleName PuckBoard Hide-BusyOverlay { }
            Mock -ModuleName PuckBoard Write-Host { }
            Mock -ModuleName PuckBoard Show-PBOpenFileDialog { Join-Path ([System.IO.Path]::GetTempPath()) 'Sched.xlsx' }
            Mock -ModuleName PuckBoard Save-SyllabusUpdate { [PSCustomObject]@{ Saved = $true; Changes = @(); Message = '' } }

            # The flying schedule: Adams flew FT5001 (CM), Brown's FI3101 line was cancelled (CX)
            function New-SchedFile {
                param ([array]$Facts)
                [PSCustomObject]@{ Path = 'Sched.xlsx'; Kind = 'DailySchedule'; Label = 'Flying schedule for Fri 16 Oct 2026'; Day = [datetime]'2026-10-16'
                    Sheets = @('Flying (16-Oct-2026)'); Events = @(); Facts = $Facts; People = @($Facts | ForEach-Object { $_.Student } | Sort-Object -Unique); Error = '' }
            }
            function New-SchedFact {
                param ([string]$Student, [string]$Mission, [string]$Status, [string]$IP = '', [string]$At = '')
                [PSCustomObject]@{ Student = $Student; Mission = $Mission; Complete = ($Status -in @('CM', 'EXEC')); Explicit = $true; Status = $Status; Instructor = $IP
                    ETD = $(if ($At) { [datetime]$At } else { $null }) }
            }
            $script:Sched = New-SchedFile @(
                New-SchedFact 'Adams, Al [2d Lt]' 'FT5001' 'CM' 'Jones, Bob [Maj]' '2026-10-16 07:30'
                New-SchedFact 'Brown, Bea [2d Lt]' 'FI3101' 'CX'
            )
            Mock -ModuleName PuckBoard Read-TrainingFile { $script:Sched }

            function Get-Event {
                param ([string]$Student, [string]$Mission)
                $s = (Get-PB).Data.Students | Where-Object { $_.Name -eq $Student }
                $s.Syllabus | Where-Object { $_.Mission -eq $Mission }
            }
        }

        It 'cancelling the file picker changes nothing' {
            Mock -ModuleName PuckBoard Show-PBOpenFileDialog { '' }
            Mock -ModuleName PuckBoard Show-PBMessage { throw 'should not ask' }
            & (Get-Module PuckBoard) { Invoke-PBUpdateSyllabi }
            Should -Invoke -ModuleName PuckBoard Read-TrainingFile -Times 0 -Exactly
            (Get-Event 'Adams, Al' 'FT5001').IsComplete | Should -BeFalse
        }

        It 'a file that is neither training file is explained, and nothing changes' {
            Mock -ModuleName PuckBoard Read-TrainingFile { [PSCustomObject]@{ Kind = ''; Error = 'Notes.xlsx is not a flying schedule (...) or a Big Board (...), so nothing was changed.' } }
            Mock -ModuleName PuckBoard Show-PBMessage { 'OK' }
            & (Get-Module PuckBoard) { Invoke-PBUpdateSyllabi }
            Should -Invoke -ModuleName PuckBoard Show-PBMessage -Times 1 -Exactly -ParameterFilter { $Icon -eq 'Warning' -and $Text -like 'Notes.xlsx is not a flying schedule*' }
            Should -Invoke -ModuleName PuckBoard Save-SyllabusUpdate -Times 0 -Exactly
            Get-LastStatus | Should -BeExactly 'Warn|Update Syllabi: Sched.xlsx was not used - nothing was changed.'
        }

        It 'shows what will change first, and No changes nothing' {
            Mock -ModuleName PuckBoard Show-PBMessage { 'No' }
            & (Get-Module PuckBoard) { Invoke-PBUpdateSyllabi }
            Should -Invoke -ModuleName PuckBoard Show-PBMessage -Times 1 -Exactly -ParameterFilter {
                $Title -eq 'Update Syllabi' -and $Buttons -eq 'YesNo' -and
                $Text -like 'Flying schedule for Fri 16 Oct 2026*Adams, Al: complete FT5001; SimDebt +1*Brown, Bea: NOT complete FI3101 (CX); SimDebt +0.5*Update these syllabi?*' -and
                $Text -notlike '*fill you are reviewing*'
            }
            (Get-Event 'Adams, Al' 'FT5001').IsComplete | Should -BeFalse
            (Get-Event 'Brown, Bea' 'FI3101').IsComplete | Should -BeTrue
            Should -Invoke -ModuleName PuckBoard Save-SyllabusUpdate -Times 0 -Exactly
            Get-LastStatus | Should -BeExactly '|Update Syllabi cancelled - nothing was changed.'
        }

        It 'Yes changes the syllabi on the board, saves them right away and checks the board again' {
            Mock -ModuleName PuckBoard Show-PBMessage { 'Yes' }
            & (Get-Module PuckBoard) { Invoke-PBUpdateSyllabi }
            $a = Get-Event 'Adams, Al' 'FT5001'
            $a.IsComplete | Should -BeTrue
            $a.Instructor | Should -BeExactly 'Jones, Bob'
            $a.ETD | Should -BeExactly '2026-10-16 07:30:00'
            (Get-Event 'Brown, Bea' 'FI3101').IsComplete | Should -BeFalse
            ((Get-PB).Data.Students | Where-Object { $_.Name -eq 'Adams, Al' }).SimDebt | Should -Be 1
            Should -Invoke -ModuleName PuckBoard Save-SyllabusUpdate -Times 1 -Exactly -ParameterFilter { @($Targets).Count -eq 2 -and @($MemoryChanges).Count -eq 2 }
            Should -Invoke -ModuleName PuckBoard Start-PBAssessQueue -Times 1
            Get-LastStatus | Should -BeExactly 'Good|Updated 2 event(s) for 2 student(s) from Sched.xlsx and saved them to OutputSchedule and MasterData.'
        }

        It 'when the files cannot be saved the change is kept for this run, and it says so' {
            Mock -ModuleName PuckBoard Show-PBMessage { if ($Buttons -eq 'YesNo') { 'Yes' } else { 'OK' } }
            Mock -ModuleName PuckBoard Save-SyllabusUpdate { [PSCustomObject]@{ Saved = $false; Changes = @(); Message = 'MasterData.xlsx is still open' } }
            & (Get-Module PuckBoard) { Invoke-PBUpdateSyllabi }
            (Get-Event 'Adams, Al' 'FT5001').IsComplete | Should -BeTrue
            Should -Invoke -ModuleName PuckBoard Show-PBMessage -Times 1 -Exactly -ParameterFilter { $Icon -eq 'Warning' -and $Text -like '*for this run only*(MasterData.xlsx is still open). Export writes them*' }
            Get-LastStatus | Should -BeLike 'Warn|Updated 2 event(s) for 2 student(s) from Sched.xlsx for this run only*'
        }

        It 'while reviewing a fill: warns that the fill goes, and Yes throws it away' {
            & (Get-Module PuckBoard) { $script:PB.Mode = 'Review'; $script:PB.FillResult = 'old fill' }
            Mock -ModuleName PuckBoard Show-PBMessage { 'Yes' }
            & (Get-Module PuckBoard) { Invoke-PBUpdateSyllabi }
            Should -Invoke -ModuleName PuckBoard Show-PBMessage -Times 1 -Exactly -ParameterFilter { $Text -like '*The fill you are reviewing is thrown away*' }
            (Get-PB).Mode | Should -BeExactly 'Edit'
            (Get-PB).FillResult | Should -BeNullOrEmpty
        }

        It 'when every syllabus already matches it says so and asks nothing' {
            $script:Sched = New-SchedFile @(New-SchedFact 'Adams, Al [2d Lt]' 'FI3101' 'CM' 'SIM IP' '2026-10-15 08:00')
            Mock -ModuleName PuckBoard Show-PBMessage { 'OK' }
            & (Get-Module PuckBoard) { Invoke-PBUpdateSyllabi }
            Should -Invoke -ModuleName PuckBoard Show-PBMessage -Times 1 -Exactly -ParameterFilter { $Buttons -eq 'OK' -and $Icon -eq 'Information' -and $Text -like '*nothing to change*' }
            Should -Invoke -ModuleName PuckBoard Save-SyllabusUpdate -Times 0 -Exactly
            Get-LastStatus | Should -BeExactly 'Good|Update Syllabi: every syllabus already matches Sched.xlsx - nothing was changed.'
        }
    }
}
