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
