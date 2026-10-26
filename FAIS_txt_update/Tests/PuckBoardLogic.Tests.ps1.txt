# Pester 5 tests for the puck board's one-day-at-a-time logic (Modules\PuckBoard.psm1) without
# a window: the WPF-only helpers are mocked and the controls are plain objects. The real window
# is driven by Tests\Invoke-SelfTest.ps1 on Windows.
# Dates: Tue 2026-10-20 (D1), Wed 2026-10-21 (D2), Thu 2026-10-22 (D3).

BeforeAll {
    . "$PSScriptRoot/Fixtures/New-FaisTestData.ps1"
    Import-FaisModules -IncludeUI

    $script:D1 = '2026-10-20'
    $script:D2 = '2026-10-21'
    $script:D3 = '2026-10-22'

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
                    FillDayButton = (& $o @{ Content = ''; Visibility = 'Visible' }); FillAllButton = (& $o @{ Content = ''; Visibility = 'Visible' })
                    BackToEditButton = (& $o @{ Visibility = 'Collapsed' })
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

    It 'opening the board gives each sheet day from the first day on a tab of its own, with its lines' {
        $rows = @((New-SR $script:D1 '07:30' -Student 'Adams, Al'), (New-SR $script:D1 '09:00'), (New-SR $script:D2 '07:30'), (New-SR $script:D3 '08:00'), (New-SR $script:D3 '08:30'), (New-SR '2026-10-19' '07:30'))
        Initialize-TestBoard -Days @($script:D1) -Sheet (New-Sheet $rows)
        $msg = InModuleScope PuckBoard { Invoke-PBSheetStartup }
        $pb = Get-PB
        ($pb.Board.Days.ToArray() -join ',') | Should -BeExactly "$($script:D1),$($script:D2),$($script:D3)" -Because 'Mon 19 Oct is before the board and is left off'
        @(Get-BoardDayLines -Board $pb.Board -Day $script:D1).Count | Should -Be 2
        @(Get-BoardDayLines -Board $pb.Board -Day $script:D2).Count | Should -Be 1
        @(Get-BoardDayLines -Board $pb.Board -Day $script:D3).Count | Should -Be 2
        (@(Get-BoardDayLines -Board $pb.Board -Day $script:D1))[0].Student | Should -BeExactly 'Adams, Al'
        $msg | Should -BeLike 'Loaded Tue 20 Oct (2 line(s)), Wed 21 Oct (1 line(s)), Thu 22 Oct (2 line(s)) from MasterData*'
        $pb.ActiveDay | Should -BeExactly $script:D1 -Because 'the board stays on its first day'
        foreach ($d in @($script:D1, $script:D2, $script:D3)) { (Get-BoardSheetState -Board $pb.Board -SheetRows $rows -Day $d).State | Should -BeExactly 'InSync' }
        $pb.SheetBannerInfo | Should -BeNullOrEmpty
    }

    It 'a new day on the sheet gets its own tab when MasterData is saved' {
        Initialize-TestBoard -Days @($script:D1) -Sheet (New-Sheet @(New-SR $script:D1 '07:30'))
        [void](InModuleScope PuckBoard { Invoke-PBSheetStartup })
        $pb = Get-PB
        $pb.Data.ScheduleSheet = New-Sheet @((New-SR $script:D1 '07:30'), (New-SR $script:D2 '08:00'))
        $done = @(InModuleScope PuckBoard { Invoke-PBSheetAutoUpdate })
        $done -join ', ' | Should -BeExactly 'Wed 21 Oct (1 line(s))'
        ($pb.Board.Days.ToArray() -join ',') | Should -BeExactly "$($script:D1),$($script:D2)"
        $pb.ActiveDay | Should -BeExactly $script:D1 -Because 'a background update does not change the day on screen'
    }

    It 'a removed sheet day stays off when MasterData is saved, until its lines on the sheet change' {
        Initialize-TestBoard -Days @($script:D1) -Sheet (New-Sheet @((New-SR $script:D1 '07:30'), (New-SR $script:D2 '08:00')))
        [void](InModuleScope PuckBoard { Invoke-PBSheetStartup })
        $pb = Get-PB
        $pb.Settings.BoardDir = $TestDrive
        $saved = Join-Path $TestDrive "$($script:D2).json"
        [System.IO.File]::WriteAllText($saved, '{}')
        InModuleScope PuckBoard -Parameters @{ D = $script:D2 } { param($D) Remove-PBDay $D }
        ($pb.Board.Days.ToArray() -join ',') | Should -BeExactly $script:D1
        Test-Path -LiteralPath $saved | Should -BeFalse -Because 'its autosave goes with it'
        @(InModuleScope PuckBoard { Invoke-PBSheetAutoUpdate }).Count | Should -Be 0
        ($pb.Board.Days.ToArray() -join ',') | Should -BeExactly $script:D1
        $pb.Data.ScheduleSheet = New-Sheet @((New-SR $script:D1 '07:30'), (New-SR $script:D2 '08:15'))
        @(InModuleScope PuckBoard { Invoke-PBSheetAutoUpdate }).Count | Should -Be 1
        ($pb.Board.Days.ToArray() -join ',') | Should -BeExactly "$($script:D1),$($script:D2)" -Because 'the sheet changed for that day'
    }

    It 'a sheet with only days before the board says so, and can put them on the first day' {
        Initialize-TestBoard -Days @($script:D2) -Sheet (New-Sheet @((New-SR $script:D1 '07:30' -Student 'Adams, Al'), (New-SR $script:D1 '08:10')))
        $msg = InModuleScope PuckBoard { Invoke-PBSheetStartup }
        $pb = Get-PB
        $msg | Should -BeExactly ''
        $pb.Board.Lines.Count | Should -Be 0
        $pb.C.SheetBanner.Visibility | Should -BeExactly 'Visible'
        $pb.SheetBannerInfo.State | Should -BeExactly 'PastDay'
        $pb.C.SheetBannerText.Text | Should -BeLike "*only has lines for days before Wed 21 Oct*2 line(s) for Tue 20 Oct. They were not loaded."
        $pb.C.SheetBannerLoadButton.Content | Should -BeExactly 'Put them on Wed 21 Oct'
        $pb.C.SheetBannerIgnoreButton.Content | Should -BeExactly 'Not now'
        InModuleScope PuckBoard { Invoke-PBSheetBannerLoad }
        ($pb.Board.Days.ToArray() -join ',') | Should -BeExactly $script:D2 -Because 'the earlier day is not added'
        (@(Get-BoardDayLines -Board $pb.Board -Day $script:D2) | ForEach-Object { $_.ETD.ToString('HHmm') + ' ' + $_.Student }) -join ',' | Should -BeExactly '0730 Adams, Al,0810 '
        $pb.C.SheetBanner.Visibility | Should -BeExactly 'Collapsed'
    }

    It '"Not now" on the earlier-days banner hides it until those lines change' {
        Initialize-TestBoard -Days @($script:D2) -Sheet (New-Sheet @(New-SR $script:D1 '07:30'))
        [void](InModuleScope PuckBoard { Invoke-PBSheetStartup })
        $pb = Get-PB
        InModuleScope PuckBoard { Invoke-PBSheetBannerIgnore }
        $pb.C.SheetBanner.Visibility | Should -BeExactly 'Collapsed'
        $pb.Data.ScheduleSheet = New-Sheet @(New-SR $script:D1 '07:45')
        InModuleScope PuckBoard { Update-PBSheetBanner }
        $pb.SheetBannerInfo.State | Should -BeExactly 'PastDay'
    }

    It 'the banner offers a later sheet day read before the board could add it, and "Not now" hides it until that day changes' {
        Initialize-TestBoard -Days @($script:D1) -Sheet (New-Sheet @(New-SR $script:D1 '07:30'))
        [void](InModuleScope PuckBoard { Invoke-PBSheetStartup })
        $pb = Get-PB
        # Read (in review mode, say) but not applied yet
        $pb.Data.ScheduleSheet = New-Sheet @((New-SR $script:D1 '07:30'), (New-SR $script:D2 '07:30'))
        InModuleScope PuckBoard { Update-PBSheetBanner }
        $pb.SheetBannerInfo.State | Should -BeExactly 'OtherDay'
        $pb.C.SheetBannerIgnoreButton.Content | Should -BeExactly 'Not now'
        InModuleScope PuckBoard { Invoke-PBSheetBannerIgnore }
        $pb.C.SheetBanner.Visibility | Should -BeExactly 'Collapsed'
        @(InModuleScope PuckBoard { Invoke-PBSheetAutoUpdate }).Count | Should -Be 0 -Because 'a day waved off is not added by itself either'
        $pb.Data.ScheduleSheet = New-Sheet @((New-SR $script:D1 '07:30'), (New-SR $script:D2 '07:30'), (New-SR $script:D2 '08:00'))
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

    It '"Fill this day" fills and reviews just the day on screen, with exported days as context' {
        $pub = @([PSCustomObject]@{ ETD = "$($script:D2) 07:30:00"; Instructor = 'Jones, Bob'; SortieType = 'Day Local'; Student = 'Adams, Al'; Mission = 'FT5001'; Callsign = 'TALON'; Num = 31; Notes = '' })
        Initialize-TestBoard -Days @($script:D1, $script:D2) -Sheet $null -PrevSchedule $pub -Active $script:D1
        $pb = Get-PB
        Add-BoardLine -Board $pb.Board -Line @(
            New-BoardLine -ETD ([datetime]"$($script:D1) 09:00") -SortieType 'Day Local' -Instructor 'Jones, Bob'
            New-BoardLine -ETD ([datetime]"$($script:D2) 10:00") -SortieType 'Day Local' -Instructor 'Lee, Ann'
        ) -NoUndo
        InModuleScope PuckBoard { Update-PBFillScope }
        $pb.C.FillDayButton.Content | Should -BeExactly 'Fill Tue 20 Oct'
        $pb.C.FillAllButton.Content | Should -BeExactly 'Fill all 2 days'
        $pb.C.FillAllButton.Visibility | Should -BeExactly 'Visible'

        InModuleScope PuckBoard { Invoke-PBFill -Scope Day } 6>$null
        $pb.Mode | Should -BeExactly 'Review'
        ($pb.FillDays -join ',') | Should -BeExactly $script:D1
        @($pb.FillResult.Schedule).Count | Should -Be 1
        @($pb.FillResult.Context).Count | Should -Be 1 -Because 'the exported Wednesday rode along'
        $row = $pb.FillResult.Schedule[0]
        "$(Get-NameText $row.Student)|$(Get-NameText $row.Mission)" | Should -Not -BeExactly 'Adams, Al|FT5001'
        $pb.C.ModeBannerText.Text | Should -BeLike 'REVIEW (Tue 20 Oct)*Export writes Tue 20 Oct*'
        $pb.C.FillDayButton.Visibility | Should -BeExactly 'Collapsed' -Because 'both fill buttons hide while reviewing'
        $pb.C.FillAllButton.Visibility | Should -BeExactly 'Collapsed'

        # Re-fill repeats the same day even after switching tabs
        $pb.ActiveDay = $script:D2
        InModuleScope PuckBoard { Invoke-PBFill -Refill } 6>$null
        ($pb.FillDays -join ',') | Should -BeExactly $script:D1

        InModuleScope PuckBoard { Exit-PBReview }
        InModuleScope PuckBoard { Invoke-PBFill -Scope All } 6>$null
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

    It 'the dropdown lists the missions the student is not opted for, once the line is checked' {
        $texts = @(& (Get-Module PuckBoard) { param($b) Get-PBMissionChoices (Get-BoardLine $b 'L1') } $script:Board)
        ($texts -join '|') | Should -BeExactly '(auto)|FT5001|Type a mission...' -Because 'until the check is in, only MasterData''s opted events'
        & (Get-Module PuckBoard) { param($b) $script:PB.Assess['L1'] = Get-PlacementAssessment -Board $b -LineId 'L1' -StudentName 'Adams, Al' -Data $script:PB.Data } $script:Board
        $texts = @(& (Get-Module PuckBoard) { param($b) Get-PBMissionChoices (Get-BoardLine $b 'L1') } $script:Board)
        $texts[0] | Should -BeExactly '(auto)'
        $texts[1] | Should -BeExactly 'FT5001'
        $texts | Should -Contain 'FT5102 (not opted)'
        $texts[-1] | Should -BeExactly 'Type a mission...'
    }

    It 'a pick in the dropdown becomes a command only when it changes something: <Why>' -TestCases @(
        @{ Why = '(auto) on an (auto) line'; Value = '(auto)'; Raw = ''; Enabled = $true; Want = $false }
        @{ Why = 'the line''s own mission again'; Value = 'FT5102'; Raw = 'FT5102'; Enabled = $true; Want = $false }
        @{ Why = 'a different mission'; Value = 'FT5001'; Raw = ''; Enabled = $true; Want = $true }
        @{ Why = 'a not-opted mission'; Value = 'FT5102 (not opted)'; Raw = 'FT5001'; Enabled = $true; Want = $true }
        @{ Why = 'back to (auto)'; Value = '(auto)'; Raw = 'FT5001'; Enabled = $true; Want = $true }
        @{ Why = 'Type a mission... always asks'; Value = 'Type a mission...'; Raw = 'FT5001'; Enabled = $true; Want = $true }
        @{ Why = 'a case-only fix'; Value = 'FT5001'; Raw = 'ft5001'; Enabled = $true; Want = $true }
        @{ Why = 'a disabled row'; Value = 'FT5001'; Raw = ''; Enabled = $false; Want = $false }
        @{ Why = 'text not in the list'; Value = 'XX9999'; Raw = ''; Enabled = $true; Want = $false }
    ) {
        $vm = [PSCustomObject]@{ LineId = 'L1'; MissionRaw = $Raw; MissionEnabled = $Enabled
            MissionOptions = @('(auto)', 'FT5001', 'ft5001', 'FT5102', 'FT5102 (not opted)', 'Type a mission...') }
        $cmd = & (Get-Module PuckBoard) { param($v, $t) Get-PBMissionCommand -VM $v -Value $t } $vm $Value
        ($null -ne $cmd) | Should -Be $Want
        if ($Want) { $cmd.Cmd | Should -BeExactly 'ChooseMission'; $cmd.LineId | Should -BeExactly 'L1'; $cmd.Value | Should -BeExactly $Value }
    }

    It 'the ChooseMission command runs the pick' {
        Mock -ModuleName PuckBoard Show-PBMessage { throw 'should not ask' }
        & (Get-Module PuckBoard) { Invoke-PBCommand @{ Cmd = 'ChooseMission'; LineId = 'L1'; Value = 'FT5001' } }
        Get-L1Mission | Should -BeExactly 'FT5001'
    }

    It 'a kept rule break says so on the status bar, after the pop-up' {
        Mock -ModuleName PuckBoard Show-PBMessage { 'No' }
        Mock -ModuleName PuckBoard Show-PBPlacementPopup { & (Get-Module PuckBoard) { Set-PBStatus 'Adams, Al -> 0730 Day Local: FT5102, 1 rule flag' 'Hard' } }
        Invoke-Choose 'FT5102 (not opted)'
        Get-LastStatus | Should -BeExactly 'Warn|FT5102 kept - it is flagged OVERRIDE on the output.'
    }

    It 'a case-only change of the mission is saved' {
        Mock -ModuleName PuckBoard Show-PBMessage { throw 'should not ask' }
        Set-BoardMission -Board $script:Board -Id 'L1' -Mission 'ft5001'
        Invoke-Choose 'FT5001'
        Get-L1Mission | Should -BeExactly 'FT5001'
    }

    It 'a typed line with a student and a mission shows the rule check' {
        $typed = New-BoardLine -ETD ([datetime]"$($script:D1) 09:00") -SortieType 'Day Local' -Instructor 'Jones, Bob' -Student 'Adams, Al' -Mission 'FT5102' -Id 'Q1'
        & (Get-Module PuckBoard) { param($l) Invoke-PBCommand @{ Cmd = 'AddLines'; Lines = @($l); Status = 'Added 1 line(s).' } } $typed
        Should -Invoke -ModuleName PuckBoard Show-PBPlacementPopup -Times 1 -Exactly -ParameterFilter { $LineId -eq 'Q1' }
        $plain = New-BoardLine -ETD ([datetime]"$($script:D1) 10:00") -SortieType 'Day Local' -Instructor 'Jones, Bob' -Id 'Q2'
        & (Get-Module PuckBoard) { param($l) Invoke-PBCommand @{ Cmd = 'AddLines'; Lines = @($l); Status = 'Added 1 line(s).' } } $plain
        Should -Invoke -ModuleName PuckBoard Show-PBPlacementPopup -Times 1 -Exactly -Because 'a line with no student needs no check'
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


Describe 'Puck board: click and place, Line details, fill buttons, export result (no window)' {
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
        Mock -ModuleName PuckBoard Stop-PBHeat { }
        Mock -ModuleName PuckBoard Start-PBHeat { }
        Mock -ModuleName PuckBoard Update-PBHoldBanner { }
        Mock -ModuleName PuckBoard Update-PBHeldHeat { }
        Mock -ModuleName PuckBoard Set-PBRowHighlight { }
        Mock -ModuleName PuckBoard Get-PBActiveDay { & (Get-Module PuckBoard) { $script:PB.ActiveDay } }
        Mock -ModuleName PuckBoard Set-PBStatus { & (Get-Module PuckBoard) { param($t, $k) $script:PB.C.StatusText.Text = $t; $script:PB.StatusLog.Add("$k|$t") } $Text $Kind }
        # Commands run straight away (the window would run them after the click)
        Mock -ModuleName PuckBoard Request-PBCommand { & (Get-Module PuckBoard) { param($c) Invoke-PBCommand $c } $Command }
        # The panel's text is kept as "Kind|Text" instead of WPF text blocks
        Mock -ModuleName PuckBoard New-PBDetailsText { "$Kind|$Text" }
        Mock -ModuleName PuckBoard Set-PBDetailsChecks { & (Get-Module PuckBoard) { param($p) $script:PB.TestChecks = @($p) } $Parts }
        Clear-AssessmentCache

        Initialize-TestBoard -Days @($script:D1, $script:D2) -Sheet (New-Sheet @())
        $script:Board = (Get-PB).Board
        $qe = ConvertFrom-QuickEntry -Text '1015 O&B' -Day $script:D1 -Instructors (New-TestInstructors) -Students @() -Board $script:Board
        Add-BoardLine -Board $script:Board -Line @(New-BoardLine -ETD ([datetime]"$($script:D1) 07:30") -SortieType 'Day Local' -Instructor 'Jones, Bob' -Id 'L1') -NoUndo
        Add-BoardLine -Board $script:Board -Line $qe.Lines -NoUndo
        $script:ReturnId = @($qe.Lines | Where-Object { $_.SortieType -eq 'O&B Return' })[0].Id
        # Cruz is on hold, so he breaks a hard rule on every line
        (Get-PB).Data.Students = @((Get-PB).Data.Students) + @(New-TestStudent -Name 'Cruz, Cy' -OnHold -Done 'FI3101|2026-10-15 08:00')

        & (Get-Module PuckBoard) {
            $o = { param($h) [PSCustomObject]$h }
            foreach ($n in @('DetailsTitle', 'DetailsListInfo')) { $script:PB.C[$n] = & $o @{ Text = '' } }
            $script:PB.C.DetailsHideHardCheck = & $o @{ IsChecked = $false }
            $script:PB.C.DetailsList = & $o @{ SelectedItem = $null; ItemsSource = $null }
            $script:PB.C.DetailsPlaceButton = & $o @{ IsEnabled = $false }
            $timer = & $o @{ Running = $false }
            $timer | Add-Member -MemberType ScriptMethod -Name Start -Value { $this.Running = $true }
            $timer | Add-Member -MemberType ScriptMethod -Name Stop -Value { $this.Running = $false }
            $script:PB.DetailsTimer = $timer
            $script:PB.DetailsVisible = $true
            $script:PB.SelectedLineId = ''
            $script:PB.Held = $null
            $script:PB.TestChecks = @()
            Initialize-PBDetails
        }

        function Invoke-PB { param ([hashtable]$Cmd) & (Get-Module PuckBoard) { param($c) Invoke-PBCommand $c } $Cmd }
        function Get-Line { param ([string]$Id) Get-BoardLine $script:Board $Id }
        function Get-LastStatus { @((Get-PB).StatusLog)[-1] }
        function Invoke-DetailsDrain {
            $n = 0
            while ((Get-PB).DetailsTimer.Running -and $n -lt 200) { & (Get-Module PuckBoard) { Invoke-PBDetailsTick }; $n++ }
        }
    }

    It 'a click on a puck picks it up, and a second click lets go' {
        Invoke-PB @{ Cmd = 'Hold'; Kind = 'Student'; Name = 'Adams, Al' }
        (Get-PB).Held.Name | Should -BeExactly 'Adams, Al'
        Get-LastStatus | Should -BeLike '|Holding Adams, Al. Click a student slot*Esc*'
        Invoke-PB @{ Cmd = 'Hold'; Kind = 'Student'; Name = 'Adams, Al' }
        (Get-PB).Held | Should -BeNullOrEmpty
        Get-LastStatus | Should -BeExactly '|Let go of Adams, Al.'
    }

    It 'clicking a slot while holding a student hardcodes them there and lets go; Ctrl keeps holding' {
        Invoke-PB @{ Cmd = 'Hold'; Kind = 'Student'; Name = 'Adams, Al' }
        Invoke-PB @{ Cmd = 'PlaceHeld'; LineId = 'L1' }
        (Get-Line 'L1').Student | Should -BeExactly 'Adams, Al'
        (Get-PB).Held | Should -BeNullOrEmpty
        Invoke-PB @{ Cmd = 'Hold'; Kind = 'Student'; Name = 'Brown, Bea' }
        Invoke-PB @{ Cmd = 'PlaceHeld'; LineId = 'L1'; Keep = $true }
        (Get-Line 'L1').Student | Should -BeExactly 'Brown, Bea'
        (Get-PB).Held.Name | Should -BeExactly 'Brown, Bea' -Because 'Ctrl+click keeps holding for the next line'
        [void](Undo-Board $script:Board)
        (Get-Line 'L1').Student | Should -BeExactly 'Adams, Al' -Because 'each placement is one undo step'
    }

    It 'a held IP goes on the IP slot' {
        Invoke-PB @{ Cmd = 'Hold'; Kind = 'IP'; Name = 'Lee, Ann' }
        Invoke-PB @{ Cmd = 'PlaceHeld'; LineId = 'L1' }
        (Get-Line 'L1').Instructor | Should -BeExactly 'Lee, Ann'
        (Get-PB).Held | Should -BeNullOrEmpty
    }

    It 'a return line takes nobody directly, and the puck stays held' {
        Invoke-PB @{ Cmd = 'Hold'; Kind = 'Student'; Name = 'Adams, Al' }
        Invoke-PB @{ Cmd = 'PlaceHeld'; LineId = $script:ReturnId }
        (Get-Line $script:ReturnId).Student | Should -BeNullOrEmpty
        Get-LastStatus | Should -BeLike 'Warn|1315 O&B Return follows its O&B: put Adams, Al on the O&B line instead.'
        (Get-PB).Held.Name | Should -BeExactly 'Adams, Al'
    }

    It 'the wrong kind of slot is explained and nothing is placed' {
        Invoke-PB @{ Cmd = 'Hold'; Kind = 'Student'; Name = 'Adams, Al' }
        $slot = [PSCustomObject]@{ Tag = 'IPSlot'; DataContext = [PSCustomObject]@{ LineId = 'L1' } }
        & (Get-Module PuckBoard) { param($s) Request-PBPlaceOnSlot -Slot $s } $slot | Should -BeFalse
        Get-LastStatus | Should -BeExactly "Warn|Adams, Al is a student: click a line's student slot."
        (Get-Line 'L1').Student | Should -BeNullOrEmpty
        (Get-Line 'L1').Instructor | Should -BeExactly 'Jones, Bob'
        $slot.Tag = 'StudentSlot'
        & (Get-Module PuckBoard) { param($s) Request-PBPlaceOnSlot -Slot $s } $slot | Should -BeTrue
        (Get-Line 'L1').Student | Should -BeExactly 'Adams, Al'
    }

    It 'a double-click on a puck needs a selected line, then puts the puck on it' {
        Invoke-PB @{ Cmd = 'PlaceOnSelected'; Kind = 'Student'; Name = 'Brown, Bea' }
        Get-LastStatus | Should -BeLike 'Warn|Click a line first, then double-click Brown, Bea*'
        & (Get-Module PuckBoard) { Select-PBLine 'L1' }
        Invoke-PB @{ Cmd = 'PlaceOnSelected'; Kind = 'Student'; Name = 'Brown, Bea' }
        (Get-Line 'L1').Student | Should -BeExactly 'Brown, Bea'
        Invoke-PB @{ Cmd = 'PlaceOnSelected'; Kind = 'IP'; Name = 'Lee, Ann' }
        (Get-Line 'L1').Instructor | Should -BeExactly 'Lee, Ann'
    }

    It 'the selection is dropped when its line goes or another day is on screen' {
        & (Get-Module PuckBoard) { Select-PBLine 'L1'; Sync-PBSelection }
        (Get-PB).SelectedLineId | Should -BeExactly 'L1'
        (Get-PB).ActiveDay = $script:D2
        & (Get-Module PuckBoard) { Sync-PBSelection }
        (Get-PB).SelectedLineId | Should -BeExactly ''
        Should -Invoke -ModuleName PuckBoard Update-PBDay -ParameterFilter { $Day -eq '2026-10-20' } -Because 'the old day is redrawn without its outline'
        (Get-PB).ActiveDay = $script:D1
        & (Get-Module PuckBoard) { Select-PBLine 'L1' }
        Remove-BoardLine -Board $script:Board -Id 'L1'
        & (Get-Module PuckBoard) { Sync-PBSelection -NoRedraw }
        (Get-PB).SelectedLineId | Should -BeExactly ''
    }

    It 'Line details ranks every student on the selected line: fits first, rule breaks last' {
        & (Get-Module PuckBoard) { Select-PBLine 'L1' }
        (Get-PB).C.DetailsTitle.Text | Should -BeExactly '0730 Day Local   |   IP Jones, Bob   |   no student yet'
        (Get-PB).DetailsTimer.Running | Should -BeTrue
        Invoke-DetailsDrain
        $items = @((Get-PB).Details.Items)
        $items.Count | Should -Be 3
        $items[-1].StudentName | Should -BeExactly 'Cruz, Cy'
        $items[-1].Rank | Should -Be 2
        $items[-1].Icon | Should -BeExactly ([string][char]0x2716)
        $items[-1].ReasonBrush | Should -BeExactly '#B91C1C'
        $items[0].Rank | Should -BeLessThan 2
        for ($i = 1; $i -lt $items.Count; $i++) { $items[$i - 1].Rank | Should -BeLessOrEqual $items[$i].Rank }
        (Get-PB).C.DetailsListInfo.Text | Should -BeLike '3 students: *, 1 break a hard rule. Double-click a student*'
        @((Get-PB).TestChecks)[0] | Should -BeExactly 'Head|Nobody is on this line yet'
    }

    It "picking a student in the list shows their hard rules, warnings and soft rules, and Put on this line hardcodes them" {
        & (Get-Module PuckBoard) { Select-PBLine 'L1' }
        Invoke-DetailsDrain
        $brown = @((Get-PB).Details.Items | Where-Object { $_.StudentName -eq 'Brown, Bea' })[0]
        (Get-PB).C.DetailsList.SelectedItem = $brown
        & (Get-Module PuckBoard) { Invoke-PBDetailsPick }
        (Get-PB).Details.Focus | Should -BeExactly 'Brown, Bea'
        (Get-PB).C.DetailsPlaceButton.IsEnabled | Should -BeTrue
        $checks = @((Get-PB).TestChecks)
        $checks[0] | Should -BeExactly 'Head|Brown, Bea, if they flew this line'
        $checks | Should -Contain 'Section|HARD RULES'
        $checks | Should -Contain 'Section|WARNINGS'
        @($checks | Where-Object { $_ -like 'Section|SOFT RULES*' }).Count | Should -Be 1

        & (Get-Module PuckBoard) { Invoke-PBDetailsPlace }
        (Get-Line 'L1').Student | Should -BeExactly 'Brown, Bea'
        Invoke-DetailsDrain
        @((Get-PB).TestChecks)[0] | Should -BeExactly 'Head|Brown, Bea, hardcoded on this line'
        @((Get-PB).Details.Items | Where-Object { $_.NameText -eq 'Brown, Bea  (on this line)' }).Count | Should -Be 1
    }

    It 'Hide rule breaks leaves out the students who break a hard rule' {
        & (Get-Module PuckBoard) { Select-PBLine 'L1' }
        Invoke-DetailsDrain
        (Get-PB).C.DetailsHideHardCheck.IsChecked = $true
        & (Get-Module PuckBoard) { Update-PBDetailsFilter }
        @((Get-PB).Details.Items).Count | Should -Be 2
        (Get-PB).Details.All.Count | Should -Be 3
        (Get-PB).C.DetailsListInfo.Text | Should -BeLike '*(1 hidden)*'
    }

    It 'a return line shows what it follows and ranks nobody' {
        & (Get-Module PuckBoard) { param($id) Select-PBLine $id } $script:ReturnId
        (Get-PB).DetailsTimer.Running | Should -BeFalse
        (Get-PB).C.DetailsListInfo.Text | Should -BeLike "Return lines fly with their O&B's student*"
        # (a mock does not see -Kind's default, so plain text shows as "|...")
        @((Get-PB).TestChecks)[1] | Should -BeLike '|Flies with the 1015 O&B. Whoever is put on the O&B flies this return too.'
    }

    It 'the list order: rank, then the best score, then the name' {
        $before = { param($a, $b) & (Get-Module PuckBoard) { param($x, $y) Test-PBCandidateBefore $x $y } $a $b }
        $c = { param($n, $r, $s) [PSCustomObject]@{ StudentName = $n; Rank = $r; SortScore = $s } }
        & $before (& $c 'Zed' 0 1) (& $c 'Amy' 1 99) | Should -BeTrue
        & $before (& $c 'Amy' 1 5) (& $c 'Zed' 1 9) | Should -BeFalse
        & $before (& $c 'Amy' 1 5) (& $c 'Zed' 1 5) | Should -BeTrue
    }

    It 'Fill all days only shows when the board has more than one day' {
        & (Get-Module PuckBoard) { Update-PBFillScope }
        (Get-PB).C.FillAllButton.Visibility | Should -BeExactly 'Visible'
        (& (Get-Module PuckBoard) { Get-PBFillScope -Scope Day }) -join ',' | Should -BeExactly $script:D1
        (& (Get-Module PuckBoard) { Get-PBFillScope -Scope All }) -join ',' | Should -BeExactly "$($script:D1),$($script:D2)"
        Remove-BoardDay -Board $script:Board -Day $script:D2
        & (Get-Module PuckBoard) { Update-PBFillScope }
        (Get-PB).C.FillAllButton.Visibility | Should -BeExactly 'Collapsed'
        (Get-PB).C.FillDayButton.Content | Should -BeExactly 'Fill Tue 20 Oct'
    }

    It 'the export result says what was written, unsats and overrides included' {
        $fill = [PSCustomObject]@{
            Stats     = [PSCustomObject]@{ Lines = 12; Pinned = 3; Filled = 7; Wing = 1; Return = 1; DS = 0; ForSale = 1 }
            Overrides = @([PSCustomObject]@{ ETD = "$($script:D1) 07:30:00"; Student = 'Adams, Al'; Violations = @('Turn time 2.8 hr < 3.25 hr minimum') })
        }
        $s = Get-FaisExportSummary -Ok $true -FillResult $fill -Days @($script:D1) -OutputPath 'C:\o.xlsx' -MasterDataPath 'C:\m.xlsx' -AtisPath 'C:\a.xlsx' `
            -UnsatChanges @{ 'Brown, Bea' = @{ UnsatSim = $false }; 'Cruz, Cy' = @{ OnHold = $true } } -Warnings @('Not found in the SPName column: Cruz, Cy')
        $s.Ok | Should -BeTrue
        $s.Title | Should -BeExactly 'Schedule exported: Tue 20 Oct'
        $texts = @($s.Lines | ForEach-Object { "$($_.Kind)|$($_.Text)" })
        $texts | Should -Contain 'Text|12 lines: 3 hardcoded, 7 filled by the scheduler, 1 wingmen, 1 return legs, 0 direct support, 1 for sale'
        $texts | Should -Contain 'Warn|1 line(s) are For Sale: no student could fly them.'
        $texts | Should -Contain 'Path|C:\a.xlsx'
        $texts | Should -Contain 'Good|UnsatSim cleared for Brown, Bea (cleaned up on this schedule)'
        $texts | Should -Contain 'Text|Cruz, Cy put on hold'
        $texts | Should -Contain 'Hard|0730 Adams, Al: Turn time 2.8 hr < 3.25 hr minimum'
        $texts | Should -Contain 'Warn|Not found in the SPName column: Cruz, Cy'
    }

    It 'a failed export says nothing was written, and why' {
        $s = Get-FaisExportSummary -Ok $false -FillResult $null -Days @($script:D1) -Warnings @('C:\o.xlsx is still open, so nothing was written.')
        $s.Ok | Should -BeFalse
        $s.Title | Should -BeExactly 'Export failed: nothing was written'
        @($s.Lines | Where-Object { $_.Kind -eq 'Hard' } | ForEach-Object { $_.Text }) | Should -Be @('C:\o.xlsx is still open, so nothing was written.')
        (Get-FaisExportSummary -Ok $false -FillResult $null -Days @($script:D1)).Lines[1].Text | Should -BeLike 'The files could not be written*'
    }

    It 'the color key covers every section with a real color for each swatch' {
        $key = @(& (Get-Module PuckBoard) { Get-PBColorKey })
        $key.Count | Should -BeGreaterOrEqual 8
        foreach ($section in $key) {
            @($section.Items).Count | Should -BeGreaterThan 0
            foreach ($it in $section.Items) {
                if ($it.Swatch -ne 'None') { $it.Fill | Should -Match '^#[0-9A-Fa-f]{6}$' -Because "$($section.Title): $($it.Text)" }
            }
        }
    }

    It 'a non-counter IP puck has its own fill, in the pool and on a line' {
        $r = & (Get-Module PuckBoard) {
            @{
                K          = $script:PBColors
                Counter    = Get-PBIPPuckColors (Find-PBInstructor 'jones, bob')
                NonCounter = Get-PBIPPuckColors (Find-PBInstructor 'Leech, Tom')
                TextFalse  = Get-PBIPPuckColors ([PSCustomObject]@{ Name = 'Text, Ty'; IsCounter = 'FALSE' })
                Blank      = Get-PBIPPuckColors ([PSCustomObject]@{ Name = 'Blank, Bo'; IsCounter = '' })
                Unknown    = Get-PBIPPuckColors (Find-PBInstructor 'Nobody, Ned')
            }
        }
        $r.K.IPPuckNC | Should -Not -Be $r.K.IPPuck
        $r.Counter.Background | Should -Be $r.K.IPPuck
        $r.Counter.Border | Should -Be $r.K.IPPuckBorder
        $r.NonCounter.Background | Should -Be $r.K.IPPuckNC
        $r.NonCounter.Border | Should -Be $r.K.IPPuckNCBorder
        $r.TextFalse.Background | Should -Be $r.K.IPPuckNC
        $r.Blank.Background | Should -Be $r.K.IPPuck -Because 'a blank IsCounter counts as a counter'
        $r.Unknown.Background | Should -Be $r.K.IPPuck

        Invoke-PB @{ Cmd = 'SetIP'; LineId = 'L1'; Name = 'Leech, Tom' }
        $vm = & (Get-Module PuckBoard) { param($l) New-PBLineVM $l } (Get-Line 'L1')
        $vm.IPBackground | Should -Be $r.K.IPPuckNC
        $vm.IPBorder | Should -Be $r.K.IPPuckNCBorder
        Invoke-PB @{ Cmd = 'SetIP'; LineId = 'L1'; Name = 'Lee, Ann' }
        $vm = & (Get-Module PuckBoard) { param($l) New-PBLineVM $l } (Get-Line 'L1')
        $vm.IPBackground | Should -Be $r.K.IPPuck
    }
}

Describe 'Puck board: Add Student Absence (no window)' {
    BeforeEach {
        Mock -ModuleName PuckBoard Update-PBTabs { }
        Mock -ModuleName PuckBoard Update-PBDay { }
        Mock -ModuleName PuckBoard Update-PBPools { }
        Mock -ModuleName PuckBoard Restore-PBFocus { }
        Mock -ModuleName PuckBoard Get-PBFocusState { $null }
        Mock -ModuleName PuckBoard Request-PBAutosave { }
        Mock -ModuleName PuckBoard Start-PBAssessQueue { }
        Mock -ModuleName PuckBoard Close-PBPopup { }
        Mock -ModuleName PuckBoard Stop-PBHeat { }
        Mock -ModuleName PuckBoard Update-PBHoldBanner { }
        Mock -ModuleName PuckBoard Show-BusyOverlay { }
        Mock -ModuleName PuckBoard Hide-BusyOverlay { }
        Mock -ModuleName PuckBoard Get-PBActiveDay { & (Get-Module PuckBoard) { $script:PB.ActiveDay } }
        Mock -ModuleName PuckBoard Set-PBStatus { & (Get-Module PuckBoard) { param($t, $k) $script:PB.C.StatusText.Text = $t; $script:PB.StatusLog.Add("$k|$t") } $Text $Kind }
        Mock -ModuleName PuckBoard Save-StudentAbsences { [PSCustomObject]@{ Saved = $true; Count = 1; Message = '' } }
        Mock -ModuleName PuckBoard Show-PBMessage { throw 'should not ask' }
        Clear-AssessmentCache

        # Adams is hardcoded on Tue 0730 (his sortie window is 0600-1000)
        Initialize-TestBoard -Days @($script:D1) -Sheet (New-Sheet @())
        $script:Board = (Get-PB).Board
        Add-BoardLine -Board $script:Board -Line @(New-BoardLine -ETD ([datetime]"$($script:D1) 07:30") -SortieType 'Day Local' -Instructor 'Jones, Bob' -Student 'Adams, Al' -Id 'L1') -NoUndo
        & (Get-Module PuckBoard) { $script:PB.Held = $null }

        function New-AbsAsk {
            param ([string[]]$Students, [string]$From, [string]$To, [string]$Reason = '')
            @{ Students = $Students; Day = [datetime]'2026-10-20'; From = $From; To = $To; EndDay = [datetime]'2026-10-20'; Reason = $Reason }
        }
        function Get-Student { param ([string]$Name) (Get-PB).Data.Students | Where-Object { $_.Name -eq $Name } }
        function Get-LastStatus { @((Get-PB).StatusLog)[-1] }
    }

    It 'the button blocks the students off, checks the board again and saves to MasterData right away' {
        Mock -ModuleName PuckBoard Show-PBAbsenceDialog { New-AbsAsk -Students @('Brown, Bea') -From '1300' -To '0330 PM' -Reason 'Dentist' }
        & (Get-Module PuckBoard) { Invoke-PBAddAbsence }
        Should -Invoke -ModuleName PuckBoard Show-PBAbsenceDialog -Times 1 -Exactly -ParameterFilter { $Day -eq '2026-10-20' -and @($Students).Count -eq 0 }
        $abs = @((Get-Student 'Brown, Bea').Absences)
        $abs.Count | Should -Be 1
        $abs[0].StartTime | Should -Be ([datetime]'2026-10-20 13:00')
        $abs[0].EndTime | Should -Be ([datetime]'2026-10-20 15:30')
        $abs[0].Reason | Should -BeExactly 'Dentist'
        Should -Invoke -ModuleName PuckBoard Save-StudentAbsences -Times 1 -Exactly -ParameterFilter { @($Students).Count -eq 2 -and $null -ne $LockedPrompt }
        Should -Invoke -ModuleName PuckBoard Start-PBAssessQueue -Times 1
        Get-LastStatus | Should -BeExactly "Good|Brown, Bea blocked off Tue 20 Oct 1300-1530 (Dentist). Saved to MasterData's StudentAbsences sheet."
    }

    It 'a student already hardcoded inside the absence is named, and their line breaks the ABSENCE rule' {
        Mock -ModuleName PuckBoard Show-PBAbsenceDialog { New-AbsAsk -Students @('Adams, Al', 'Brown, Bea') -From '0700 AM' -To '0800' }
        & (Get-Module PuckBoard) { Invoke-PBAddAbsence }
        Get-LastStatus | Should -BeExactly "Warn|Adams, Al and Brown, Bea blocked off Tue 20 Oct 0700-0800. Already hardcoded inside it (flagged red): Tue 0730 Adams, Al. Saved to MasterData's StudentAbsences sheet."
        $check = Get-PlacementAssessment -Board $script:Board -LineId 'L1' -StudentName 'Adams, Al' -Data (Get-PB).Data
        @($check.Violations | ForEach-Object { $_.Rule }) | Should -Contain 'ABSENCE'
    }

    It 'when MasterData cannot be saved it holds for this run, and it says so' {
        Mock -ModuleName PuckBoard Show-PBAbsenceDialog { New-AbsAsk -Students @('Brown, Bea') -From '1300' -To '1400' }
        Mock -ModuleName PuckBoard Save-StudentAbsences { [PSCustomObject]@{ Saved = $false; Count = 0; Message = 'MasterData.xlsx is still open' } }
        Mock -ModuleName PuckBoard Show-PBMessage { 'OK' }
        & (Get-Module PuckBoard) { Invoke-PBAddAbsence }
        @((Get-Student 'Brown, Bea').Absences).Count | Should -Be 1
        Should -Invoke -ModuleName PuckBoard Show-PBMessage -Times 1 -Exactly -ParameterFilter { $Icon -eq 'Warning' -and $Text -like '*for this run only: MasterData could not be saved (MasterData.xlsx is still open). Export writes it*' }
        Get-LastStatus | Should -BeLike 'Warn|Brown, Bea blocked off Tue 20 Oct 1300-1400. It holds for this run only*'
    }

    It 'cancelling the window changes nothing, and so does an absence the student already has' {
        Mock -ModuleName PuckBoard Show-PBAbsenceDialog { $null }
        & (Get-Module PuckBoard) { Invoke-PBAddAbsence }
        @((Get-Student 'Brown, Bea').Absences).Count | Should -Be 0
        Should -Invoke -ModuleName PuckBoard Save-StudentAbsences -Times 0 -Exactly

        Mock -ModuleName PuckBoard Show-PBAbsenceDialog { New-AbsAsk -Students @('Brown, Bea') -From '1300' -To '1400' }
        & (Get-Module PuckBoard) { Invoke-PBAddAbsence }
        & (Get-Module PuckBoard) { Invoke-PBAddAbsence }
        @((Get-Student 'Brown, Bea').Absences).Count | Should -Be 1
        Should -Invoke -ModuleName PuckBoard Save-StudentAbsences -Times 1 -Exactly
        Get-LastStatus | Should -BeExactly 'Warn|Brown, Bea already blocked off Tue 20 Oct 1300-1400 - nothing was added.'
    }

    It 'a student picked up on the board starts in the box, and is let go of once blocked off' {
        & (Get-Module PuckBoard) { $script:PB.Held = @{ Kind = 'Student'; Name = 'Brown, Bea' } }
        Mock -ModuleName PuckBoard Show-PBAbsenceDialog { New-AbsAsk -Students $Students -From '1300' -To '1400' }
        & (Get-Module PuckBoard) { Invoke-PBAddAbsence }
        Should -Invoke -ModuleName PuckBoard Show-PBAbsenceDialog -Times 1 -Exactly -ParameterFilter { (@($Students) -join ',') -eq 'Brown, Bea' }
        @((Get-Student 'Brown, Bea').Absences).Count | Should -Be 1
        (Get-PB).Held | Should -BeNullOrEmpty
    }

    It 'while reviewing a fill: the fill is thrown away (it was made without the absence)' {
        & (Get-Module PuckBoard) { $script:PB.Mode = 'Review'; $script:PB.FillResult = 'old fill' }
        Mock -ModuleName PuckBoard Show-PBAbsenceDialog { New-AbsAsk -Students @('Brown, Bea') -From '1300' -To '1400' }
        & (Get-Module PuckBoard) { Invoke-PBAddAbsence }
        (Get-PB).Mode | Should -BeExactly 'Edit'
        (Get-PB).FillResult | Should -BeNullOrEmpty
    }

    It 'the export result lists the absences Export wrote because they could not be saved before' {
        $fill = [PSCustomObject]@{ Stats = [PSCustomObject]@{ Lines = 1; Pinned = 1; Filled = 0; Wing = 0; Return = 0; DS = 0; ForSale = 0 }; Overrides = @() }
        $abs = @([PSCustomObject]@{ Student = 'Adams, Al'; StartTime = [datetime]'2026-10-20 13:00'; EndTime = [datetime]'2026-10-20 15:30'; Reason = 'Dentist'; InMasterData = $false })
        $s = Get-FaisExportSummary -Ok $true -FillResult $fill -Days @($script:D1) -OutputPath 'C:\o.xlsx' -MasterDataPath 'C:\m.xlsx' -Absences $abs
        $texts = @($s.Lines | ForEach-Object { "$($_.Kind)|$($_.Text)" })
        $texts | Should -Contain 'Text|MasterData: SimDebt; StudentAbsences: 1 line(s) Add Student Absence could not save'
        $texts | Should -Contain 'Head|Student absences'
        $texts | Should -Contain 'Text|Adams, Al: Tue 20 Oct 1300-1530 (Dentist)'
        @((Get-FaisExportSummary -Ok $true -FillResult $fill -Days @($script:D1)).Lines | Where-Object { $_.Text -like '*absence*' }).Count | Should -Be 0
    }

    It 'from the pre-run menu: the same window over every student, the absence goes on them and is saved' {
        Mock -ModuleName PuckBoard Initialize-PuckBoardAssemblies { }
        Mock -ModuleName PuckBoard Show-PBAbsenceDialog { New-AbsAsk -Students @('Brown, Bea') -From '1300' -To '0200 PM' -Reason 'Dentist' }
        $stu = @((New-TestStudent -Name 'Adams, Al'), (New-TestStudent -Name 'Brown, Bea'))
        $text = Invoke-FaisAddAbsence -Students $stu -MasterDataPath 'C:\m.xlsx' -Day '2026-10-20' -LockedPrompt { $true }
        $text | Should -BeExactly "Brown, Bea blocked off Tue 20 Oct 1300-1400 (Dentist). Saved to MasterData's StudentAbsences sheet."
        Should -Invoke -ModuleName PuckBoard Show-PBAbsenceDialog -Times 1 -Exactly -ParameterFilter { $Day -eq '2026-10-20' -and @($Roster).Count -eq 2 }
        Should -Invoke -ModuleName PuckBoard Save-StudentAbsences -Times 1 -Exactly -ParameterFilter { $MasterDataPath -eq 'C:\m.xlsx' -and @($Students).Count -eq 2 }
        @($stu[1].Absences).Count | Should -Be 1
        @($stu[0].Absences).Count | Should -Be 0

        Invoke-FaisAddAbsence -Students $stu -MasterDataPath 'C:\m.xlsx' -Day '2026-10-20' | Should -BeExactly 'Brown, Bea already blocked off Tue 20 Oct 1300-1400 - nothing was added.'
        Mock -ModuleName PuckBoard Show-PBAbsenceDialog { New-AbsAsk -Students @('Adams, Al') -From '1300' -To '1400' }
        Mock -ModuleName PuckBoard Save-StudentAbsences { [PSCustomObject]@{ Saved = $false; Count = 0; Message = 'MasterData.xlsx is still open' } }
        Invoke-FaisAddAbsence -Students $stu -MasterDataPath 'C:\m.xlsx' -Day '2026-10-20' | Should -BeExactly 'Adams, Al blocked off Tue 20 Oct 1300-1400 for this run only: MasterData could not be saved (MasterData.xlsx is still open). Export writes it with the schedule.'
        Mock -ModuleName PuckBoard Show-PBAbsenceDialog { $null }
        Invoke-FaisAddAbsence -Students $stu -MasterDataPath 'C:\m.xlsx' -Day '2026-10-20' | Should -BeExactly ''
    }

    Context 'the window (plain objects stand in for its boxes)' {
        BeforeEach {
            Mock -ModuleName PuckBoard Get-PBBrush { $Color }
            Mock -ModuleName PuckBoard Get-PBClipboardText { "Brown, Bea [2d Lt]`r`nNobody, No" }
            & (Get-Module PuckBoard) {
                $o = { param($h) [PSCustomObject]$h }
                $script:PBAbs = @{ Window = (& $o @{ DialogResult = $null }); Result = $null; Chosen = (New-Object System.Collections.Generic.List[string])
                    Held = ''; Press = $null; Drag = $null; PendingDrop = $null; LastDay = [datetime]'2026-10-20'; Students = @($script:PB.Data.Students) }
                foreach ($box in @('SearchBox', 'FromBox', 'ToBox', 'ReasonBox', 'HoldText', 'PasteText', 'PreviewText')) { $script:PBAbs[$box] = & $o @{ Text = ''; Foreground = '' } }
                $script:PBAbs.StudentPool = & $o @{ ItemsSource = $null }
                $script:PBAbs.ChosenList = & $o @{ ItemsSource = $null }
                $script:PBAbs.DropHint = & $o @{ Visibility = 'Visible' }
                $script:PBAbs.DropBox = & $o @{ BorderBrush = '' }
                $script:PBAbs.AddButton = & $o @{ IsEnabled = $true }
                $script:PBAbs.DayPicker = & $o @{ SelectedDate = [datetime]'2026-10-20' }
                $script:PBAbs.EndDayPicker = & $o @{ SelectedDate = [datetime]'2026-10-20' }
            }
            function Get-Abs { & (Get-Module PuckBoard) { $script:PBAbs } }
            function Invoke-Abs { param ([scriptblock]$Do) & (Get-Module PuckBoard) $Do }
            function Get-PoolNames { @((Get-Abs).StudentPool.ItemsSource | ForEach-Object { $_.Name }) }
        }
        AfterEach { & (Get-Module PuckBoard) { $script:PBAbs = $null } }

        It 'Add absence waits for a student and both times, then says what it will add' {
            Invoke-Abs { Update-PBAbsenceView }
            (Get-Abs).AddButton.IsEnabled | Should -BeFalse
            (Get-Abs).PreviewText.Foreground | Should -BeExactly '#64748B' -Because 'boxes still empty are not an error'
            (Get-Abs).DropHint.Visibility | Should -BeExactly 'Visible'
            Get-PoolNames | Should -Be @('Adams, Al', 'Brown, Bea')

            Invoke-Abs { Add-PBAbsenceStudent -Name 'brown, bea' }
            @((Get-Abs).Chosen) | Should -Be @('Brown, Bea')
            Get-PoolNames | Should -Be @('Adams, Al') -Because 'a student in the box leaves the list'
            @((Get-Abs).ChosenList.ItemsSource | ForEach-Object { $_.Name }) | Should -Be @('Brown, Bea')
            (Get-Abs).DropHint.Visibility | Should -BeExactly 'Collapsed'

            (Get-Abs).FromBox.Text = '1300'
            (Get-Abs).ToBox.Text = '0330 pm'
            Invoke-Abs { Update-PBAbsencePreview } | Should -BeTrue
            (Get-Abs).AddButton.IsEnabled | Should -BeTrue
            (Get-Abs).PreviewText.Text | Should -BeLike "Brown, Bea: away Tue 20 Oct 1300-1530. A line for each goes on MasterData's StudentAbsences sheet."

            (Get-Abs).ToBox.Text = '1200'
            Invoke-Abs { Update-PBAbsencePreview } | Should -BeFalse
            (Get-Abs).PreviewText.Foreground | Should -BeExactly '#B91C1C'
            (Get-Abs).PreviewText.Text | Should -BeLike 'It ends (Tue 20 Oct 1200) before it starts*'
        }

        It 'a click picks a student up (thick outline) and a second click lets go; putting them in the box lets go too' {
            Invoke-Abs { Set-PBAbsenceHeld -Name 'Adams, Al' }
            (Get-Abs).Held | Should -BeExactly 'Adams, Al'
            (Get-Abs).HoldText.Text | Should -BeExactly 'Holding Adams, Al: click the box to add them (Esc lets go).'
            (Get-Abs).DropBox.BorderBrush | Should -BeExactly '#1D4ED8'
            @((Get-Abs).StudentPool.ItemsSource | Where-Object { $_.Name -eq 'Adams, Al' })[0].PuckBorderThickness | Should -BeExactly '3'
            Invoke-Abs { Set-PBAbsenceHeld -Name 'Adams, Al' }
            (Get-Abs).Held | Should -BeExactly ''
            (Get-Abs).HoldText.Text | Should -BeExactly ''

            Invoke-Abs { Set-PBAbsenceHeld -Name 'Adams, Al' }
            Invoke-Abs { Add-PBAbsenceStudent -Name $script:PBAbs.Held }
            @((Get-Abs).Chosen) | Should -Be @('Adams, Al')
            (Get-Abs).Held | Should -BeExactly ''
            (Get-Abs).DropBox.BorderBrush | Should -BeExactly '#94A3B8'
        }

        It 'pasting adds the students named in what was copied, and lists the names that are no one' {
            Invoke-Abs { Invoke-PBAbsencePaste }
            @((Get-Abs).Chosen) | Should -Be @('Brown, Bea')
            (Get-Abs).PasteText.Text | Should -BeExactly 'Not a student (or more than one): Nobody, No'
            Invoke-Abs { Invoke-PBAbsencePaste -Text 'Adams, Al' }
            @((Get-Abs).Chosen) | Should -Be @('Brown, Bea', 'Adams, Al')
            (Get-Abs).PasteText.Text | Should -BeExactly ''
            Mock -ModuleName PuckBoard Get-PBClipboardText { '' }
            Invoke-Abs { Invoke-PBAbsencePaste }
            (Get-Abs).PasteText.Text | Should -BeLike 'Nothing to paste*'

            Invoke-Abs { Remove-PBAbsenceStudent -Name 'Adams, Al' }
            @((Get-Abs).Chosen) | Should -Be @('Brown, Bea')
            Invoke-Abs { Clear-PBAbsenceStudents }
            (Get-Abs).Chosen.Count | Should -Be 0
        }

        It 'mouse actions queued while an event runs are done in order once it is over' {
            Invoke-Abs {
                $script:PBAbs.Actions = New-Object System.Collections.Generic.List[object]
                foreach ($a in @(@{ Action = 'Hold'; Name = 'Adams, Al' }, @{ Action = 'Add'; Name = 'Adams, Al' }, @{ Action = 'Paste'; Text = 'Brown, Bea' }, @{ Action = 'Remove'; Name = 'Adams, Al' })) {
                    $script:PBAbs.Actions.Add($a)
                }
                Invoke-PBAbsenceDrain
            }
            @((Get-Abs).Chosen) | Should -Be @('Brown, Bea')
            (Get-Abs).Held | Should -BeExactly ''
            (Get-Abs).Actions.Count | Should -Be 0
        }

        It 'the end day follows the day picked, unless it was set to another day' {
            (Get-Abs).DayPicker.SelectedDate = [datetime]'2026-10-21'
            Invoke-Abs { Invoke-PBAbsenceDayChanged }
            (Get-Abs).EndDayPicker.SelectedDate | Should -Be ([datetime]'2026-10-21')
            (Get-Abs).EndDayPicker.SelectedDate = [datetime]'2026-10-23'
            (Get-Abs).DayPicker.SelectedDate = [datetime]'2026-10-22'
            Invoke-Abs { Invoke-PBAbsenceDayChanged }
            (Get-Abs).EndDayPicker.SelectedDate | Should -Be ([datetime]'2026-10-23')
        }

        It "the list shows a student's absences on the day picked, an absence they already have can't be added again, and Add returns the boxes" {
            $r = New-StudentAbsences -Students @('Adams, Al') -Day '2026-10-20' -From '0900' -To '1100' -Reason 'Dentist'
            [void](Add-StudentAbsences -Absences $r.Absences -Students @((Get-PB).Data.Students))
            Invoke-Abs { Update-PBAbsencePool }
            $adams = @((Get-Abs).StudentPool.ItemsSource | Where-Object { $_.Name -eq 'Adams, Al' })[0]
            $adams.BadgeText | Should -BeExactly 'ABS 0900-1100'
            $adams.Background | Should -BeExactly '#FEF3C7'
            $adams.ToolTipText | Should -BeLike 'Absent 0900-1100 Dentist*'

            Invoke-Abs { Add-PBAbsenceStudent -Name 'Adams, Al' }
            (Get-Abs).FromBox.Text = '0900'
            (Get-Abs).ToBox.Text = '1100'
            Invoke-Abs { Update-PBAbsencePreview } | Should -BeFalse
            (Get-Abs).PreviewText.Text | Should -BeLike '*Already blocked off then (nothing is added for them): Adams, Al'

            Invoke-Abs { Add-PBAbsenceStudent -Name 'Brown, Bea' }
            (Get-Abs).ReasonBox.Text = 'Leave'
            Invoke-Abs { Submit-PBAbsenceDialog }
            (Get-Abs).Window.DialogResult | Should -BeTrue
            $ask = (Get-Abs).Result
            @($ask.Students) | Should -Be @('Adams, Al', 'Brown, Bea')
            $ask.From | Should -BeExactly '0900'
            $ask.Reason | Should -BeExactly 'Leave'
            $ask.Day | Should -Be ([datetime]'2026-10-20')
        }
    }
}
