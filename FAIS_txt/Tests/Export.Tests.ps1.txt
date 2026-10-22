# Pester 5 tests for writing the schedule out one day at a time (Modules\DataUpload.psm1).
# Workbooks are built in TestDrive. Dates: Mon 2026-10-19 (past), Tue 2026-10-20 (today), Wed 2026-10-21.

BeforeAll {
    . "$PSScriptRoot/Fixtures/New-FaisTestData.ps1"
    Import-FaisModules
    Import-Module ImportExcel -ErrorAction Stop 3>$null

    $script:Today = [datetime]'2026-10-20 15:00'

    function New-XR {
        param ([string]$At, [string]$Student = 'For Sale', [string]$Mission = '', [string]$IP = 'Jones, Bob', [string]$Type = 'Day Local')
        [PSCustomObject]@{ ETD = ([datetime]$At).ToString('yyyy-MM-dd HH:mm:ss'); Instructor = $IP; SortieType = $Type; Student = $Student; Mission = $Mission; Callsign = 'TALON'; Num = 31; Notes = '' }
    }

    function New-Output {
        param ([string]$Name)
        $p = Join-Path $TestDrive $Name
        if (Test-Path $p) { Remove-Item $p -Force }
        @(
            New-XR '2026-10-19 07:30' -Student 'Old, Ollie' -Mission 'FT5001'
            New-XR '2026-10-20 07:30' -Student 'Adams, Al' -Mission 'FT5001'
            New-XR '2026-10-21 07:30' -Student 'Adams, Al' -Mission 'FT5102'
            New-XR '2026-10-21 08:00' -Student 'Brown, Bea' -Mission 'FT5001'
        ) | Export-Excel -Path $p -WorksheetName 'MasterSchedule' 3>$null
        return $p
    }

    function New-Master {
        param ([string]$Name)
        $p = Join-Path $TestDrive $Name
        if (Test-Path $p) { Remove-Item $p -Force }
        @(
            [PSCustomObject]@{ SPKey = 1; SPName = 'Adams, Al'; SimDebt = 0 }
            [PSCustomObject]@{ SPKey = 2; SPName = 'Brown, Bea'; SimDebt = 0 }
        ) | Export-Excel -Path $p -WorksheetName 'Students' 3>$null
        return $p
    }

    function New-Fill {
        param ([array]$Rows, [string[]]$Days)
        $a = New-TestStudent -Name 'Adams, Al' -SimDebt 2.5
        $b = New-TestStudent -Name 'Brown, Bea' -SimDebt 1
        return [PSCustomObject]@{ Schedule = $Rows; Students = @($a, $b); Days = $Days }
    }

    function Get-MSDays { param ([string]$Path) @(Import-Excel -Path $Path -WorksheetName 'MasterSchedule' | ForEach-Object { ([datetime]$_.ETD).ToString('MMdd HHmm') + ' ' + $_.Student }) -join ',' }
}

Describe 'Merge-MasterScheduleRows' {
    It 'replaces the exported days, keeps the others, and drops days before -KeepFromDay' {
        $existing = @((New-XR '2026-10-19 07:30' -Student 'Old'), (New-XR '2026-10-20 07:30' -Student 'A'), (New-XR '2026-10-21 07:30' -Student 'B'))
        $new = @(New-XR '2026-10-21 09:00' -Student 'C')
        $m = @(Merge-MasterScheduleRows -Existing $existing -New $new -Days @('2026-10-21') -KeepFromDay '2026-10-20')
        (@($m | ForEach-Object { $_.Student }) -join ',') | Should -BeExactly 'A,C'
    }

    It 'with no -KeepFromDay keeps every other day' {
        $existing = @((New-XR '2026-10-19 07:30' -Student 'Old'), (New-XR '2026-10-21 07:30' -Student 'B'))
        @(Merge-MasterScheduleRows -Existing $existing -New @() -Days @('2026-10-21')).Count | Should -Be 1
    }
}

Describe 'Export-FaisOutput one day at a time' {
    BeforeAll {
        $script:Out = New-Output 'OutputSchedule.xlsx'
        $script:MD = New-Master 'MasterData.xlsx'
        $rows = @((New-XR '2026-10-21 07:45' -Student 'Brown, Bea' -Mission 'FT5102'), (New-XR '2026-10-21 09:00'))
        $script:Ok = Export-FaisOutput -FillResult (New-Fill $rows @('2026-10-21')) -OutputPath $script:Out -MasterDataPath $script:MD -Today $script:Today -Settings @{ ReopenExcel = $false } 6>$null 3>$null
    }

    It 'reports success' {
        $script:Ok | Should -BeTrue
    }

    It 'replaces Wednesday, keeps today, and drops the past day' {
        Get-MSDays $script:Out | Should -BeExactly '1020 0730 Adams, Al,1021 0745 Brown, Bea,1021 0900 For Sale'
    }

    It 'keeps MasterSchedule first and writes the student tabs' {
        $pkg = Open-ExcelPackage -Path $script:Out
        try {
            $pkg.Workbook.Worksheets[1].Name | Should -BeExactly 'MasterSchedule'
            $null -ne $pkg.Workbook.Worksheets['Adams, Al'] | Should -BeTrue
        } finally { Close-ExcelPackage $pkg -NoSave }
    }

    It 'writes SimDebt into the column headed SimDebt' {
        $s = @(Import-Excel -Path $script:MD -WorksheetName 'Students')
        ($s | Where-Object { $_.SPName -eq 'Adams, Al' }).SimDebt | Should -Be 2.5
        ($s | Where-Object { $_.SPName -eq 'Brown, Bea' }).SimDebt | Should -Be 1
        @($s[0].PSObject.Properties.Name) -join ',' | Should -BeExactly 'SPKey,SPName,SimDebt'
    }

    It 'exporting the same day again replaces it (no duplicates)' {
        $rows = @(New-XR '2026-10-21 10:00' -Student 'Adams, Al' -Mission 'FT5102')
        [void](Export-FaisOutput -FillResult (New-Fill $rows @('2026-10-21')) -OutputPath $script:Out -Today $script:Today -MasterDataPath $script:MD -Settings @{ ReopenExcel = $false } 6>$null 3>$null)
        Get-MSDays $script:Out | Should -BeExactly '1020 0730 Adams, Al,1021 1000 Adams, Al'
    }

    It 'only writes rows of the exported days even if the fill held context rows' {
        $p = New-Output 'ctx.xlsx'
        $fill = New-Fill @((New-XR '2026-10-20 08:00' -Student 'Brown, Bea'), (New-XR '2026-10-21 07:30' -Student 'Context, Row')) @('2026-10-20')
        [void](Export-FaisOutput -FillResult $fill -OutputPath $p -Today $script:Today -MasterDataPath $script:MD -Settings @{ ReopenExcel = $false } 6>$null 3>$null)
        Get-MSDays $p | Should -BeExactly '1020 0800 Brown, Bea,1021 0730 Adams, Al,1021 0800 Brown, Bea'
    }

    It 'KeepPastDays keeps that many days before today' {
        $p = New-Output 'keep.xlsx'
        [void](Export-FaisOutput -FillResult (New-Fill @(New-XR '2026-10-21 09:00') @('2026-10-21')) -OutputPath $p -Today $script:Today -MasterDataPath $script:MD -Settings @{ ReopenExcel = $false; KeepPastDays = 1 } 6>$null 3>$null)
        Get-MSDays $p | Should -BeLike '1019 0730 Old, Ollie,*'
    }

    It 'writes a new OutputSchedule when there is none' {
        $p = Join-Path $TestDrive 'brandnew.xlsx'
        [void](Export-FaisOutput -FillResult (New-Fill @(New-XR '2026-10-21 09:00') @('2026-10-21')) -OutputPath $p -MasterDataPath $script:MD -Today $script:Today -Settings @{ ReopenExcel = $false } 6>$null 3>$null)
        Get-MSDays $p | Should -BeExactly '1021 0900 For Sale'
    }

    It 'writes ATIS for the exported days only' {
        $p = New-Output 'atisout.xlsx'
        $atis = Join-Path $TestDrive 'ATIS.xlsx'
        $fill = New-Fill @((New-XR '2026-10-20 08:00' -Student 'Brown, Bea'), (New-XR '2026-10-21 07:30' -Student 'Context, Row')) @('2026-10-20')
        Export-FaisOutput -FillResult $fill -OutputPath $p -MasterDataPath $script:MD -AtisPath $atis -Today $script:Today -Settings @{ ReopenExcel = $false } 6>$null 3>$null | Should -BeTrue
        $pkg = Open-ExcelPackage -Path $atis
        try { @($pkg.Workbook.Worksheets | ForEach-Object { $_.Name }) -join ',' | Should -BeExactly (Get-DaySheetName -Date ([datetime]'2026-10-20')) }
        finally { Close-ExcelPackage $pkg -NoSave }
    }

    It 'writes nothing and returns $false when the file stays locked' {
        $p = New-Output 'locked.xlsx'
        $before = Get-MSDays $p
        $fs = [System.IO.File]::Open($p, 'Open', 'ReadWrite', 'None')
        try {
            $ok = Export-FaisOutput -FillResult (New-Fill @(New-XR '2026-10-21 09:00') @('2026-10-21')) -OutputPath $p -MasterDataPath $script:MD -Today $script:Today -LockedPrompt { param($x) $false } -Settings @{ ReopenExcel = $false } 6>$null 3>$null 2>$null
        } finally { $fs.Close() }
        $ok | Should -BeFalse
        Get-MSDays $p | Should -BeExactly $before
    }
}

Describe 'Teams day sheets' {
    It 'a re-exported day replaces its sheet instead of adding to it' {
        $dir = Join-Path $TestDrive 'teams'
        Export-DailyScheduleSheets -ScheduleData @((New-XR '2026-10-21 07:30' -Student 'A'), (New-XR '2026-10-21 08:00' -Student 'B')) -TargetFolder $dir -FileName 'Teams' 6>$null 3>$null
        Export-DailyScheduleSheets -ScheduleData @(New-XR '2026-10-21 09:00' -Student 'C') -TargetFolder $dir -FileName 'Teams' 6>$null 3>$null
        $rows = @(Import-Excel -Path (Join-Path $dir 'Teams.xlsx') -WorksheetName (Get-DaySheetName -Date ([datetime]'2026-10-21')))
        (@($rows | ForEach-Object { $_.Student }) -join ',') | Should -BeExactly 'C'
    }
}
