# Pester 5 tests for Modules\ExcelFiles.psm1 (save and close workbooks before FAIS reads or writes).
# Real Excel is not needed: a fake Excel object stands in for the COM one.

BeforeAll {
    . "$PSScriptRoot/Fixtures/New-FaisTestData.ps1"
    Import-FaisModules

    # A fake open workbook. -Lock holds the file open the way Excel does.
    function New-FakeWorkbook {
        param ([string]$FullName, [bool]$Saved = $true, [bool]$ReadOnly = $false, [string]$Dir = 'x', [switch]$Lock, [int]$BusyCalls = 0)
        $wb = [PSCustomObject]@{
            Name      = @(($FullName -replace '/', '\') -split '\\')[-1]
            FullName  = $FullName
            Path      = $Dir
            Saved     = $Saved
            ReadOnly  = $ReadOnly
            SaveCount = 0
            Closed    = $false
            Busy      = $BusyCalls
            Stream    = $null
        }
        if ($Lock) { $wb.Stream = [System.IO.File]::Open($FullName, 'OpenOrCreate', 'ReadWrite', 'None') }
        $wb | Add-Member -MemberType ScriptMethod -Name Save -Value {
            if ($this.Busy -gt 0) {
                $this.Busy--
                $ex = New-Object System.Runtime.InteropServices.COMException('busy', -2147418111)
                throw $ex
            }
            $this.SaveCount++
            $this.Saved = $true
        }
        $wb | Add-Member -MemberType ScriptMethod -Name Close -Value {
            param($save)
            $this.Closed = $true
            if ($null -ne $this.Stream) { $this.Stream.Close(); $this.Stream = $null }
        }
        return $wb
    }

    function New-FakeExcel {
        param ([array]$Workbooks)
        return [PSCustomObject]@{ Workbooks = $Workbooks }
    }

    function New-TFile {
        param ([string]$Name)
        $p = Join-Path $TestDrive $Name
        [System.IO.File]::WriteAllText($p, 'x')
        return $p
    }
}

Describe 'Test-FileWritable / Wait-FileWritable' {
    It 'a missing file is writable' {
        Test-FileWritable (Join-Path $TestDrive 'nope.xlsx') | Should -BeTrue
    }

    It 'a file held open is not writable, and is again once released' {
        $p = New-TFile 'held.xlsx'
        $fs = [System.IO.File]::Open($p, 'Open', 'ReadWrite', 'None')
        try { Test-FileWritable $p | Should -BeFalse }
        finally { $fs.Close() }
        Test-FileWritable $p | Should -BeTrue
    }

    It 'Wait-FileWritable gives up when the prompt says no' {
        $p = New-TFile 'held2.xlsx'
        $fs = [System.IO.File]::Open($p, 'Open', 'ReadWrite', 'None')
        try { Wait-FileWritable -Path $p -Prompt { param($x) $false } | Should -BeFalse }
        finally { $fs.Close() }
    }
}

Describe 'Find-FaisOpenWorkbook' {
    It 'matches the full path, case-insensitively' {
        $p = New-TFile 'MasterData.xlsx'
        $wb = New-FakeWorkbook -FullName $p.ToUpperInvariant()
        $other = New-FakeWorkbook -FullName (Join-Path $TestDrive 'Other.xlsx')
        $found = Find-FaisOpenWorkbook -ExcelApp (New-FakeExcel @($other, $wb)) -Path $p
        [object]::ReferenceEquals($found, $wb) | Should -BeTrue
    }

    It 'matches a OneDrive workbook (web address) by file name' {
        $p = New-TFile 'OutputSchedule.xlsx'
        $wb = New-FakeWorkbook -FullName 'https://d.docs.live.net/abc/Desktop/OutputSchedule.xlsx'
        $wb.Name = 'OutputSchedule.xlsx'
        $found = Find-FaisOpenWorkbook -ExcelApp (New-FakeExcel @($wb)) -Path $p
        [object]::ReferenceEquals($found, $wb) | Should -BeTrue
    }

    It 'ignores a same-named workbook from another folder while this file is free' {
        $p = New-TFile 'Same.xlsx'
        $elsewhere = New-FakeWorkbook -FullName 'Z:\Backups\Same.xlsx'
        $elsewhere.Name = 'Same.xlsx'
        Find-FaisOpenWorkbook -ExcelApp (New-FakeExcel @($elsewhere)) -Path $p | Should -BeNullOrEmpty
    }

    It 'takes a same-named workbook when this file is locked (mapped drive vs UNC path)' {
        $p = New-TFile 'Mapped.xlsx'
        $wb = New-FakeWorkbook -FullName '\\server\share\Mapped.xlsx'
        $wb.Name = 'Mapped.xlsx'
        $fs = [System.IO.File]::Open($p, 'Open', 'ReadWrite', 'None')
        try {
            $found = Find-FaisOpenWorkbook -ExcelApp (New-FakeExcel @($wb)) -Path $p
            [object]::ReferenceEquals($found, $wb) | Should -BeTrue
        }
        finally { $fs.Close() }
    }

    It 'no Excel means nothing found' {
        Find-FaisOpenWorkbook -ExcelApp $null -Path (New-TFile 'n.xlsx') | Should -BeNullOrEmpty
    }
}

Describe 'Save-FaisWorkbook' {
    It 'saves a workbook with unsaved changes and keeps it open' {
        $p = New-TFile 'Dirty.xlsx'
        $wb = New-FakeWorkbook -FullName $p -Saved $false
        $r = Save-FaisWorkbook -Path $p -ExcelApp (New-FakeExcel @($wb))
        $r.WasOpen | Should -BeTrue
        $r.Saved | Should -BeTrue
        $r.Closed | Should -BeFalse
        $wb.SaveCount | Should -Be 1
        $wb.Closed | Should -BeFalse
        $r.Message | Should -BeExactly 'Dirty.xlsx saved in Excel.'
    }

    It 'with -Close saves, closes and waits until the file is free' {
        $p = New-TFile 'Locked.xlsx'
        $wb = New-FakeWorkbook -FullName $p -Saved $false -Lock
        Test-FileWritable $p | Should -BeFalse
        $r = Save-FaisWorkbook -Path $p -Close -ExcelApp (New-FakeExcel @($wb))
        $r.Saved | Should -BeTrue
        $r.Closed | Should -BeTrue
        $wb.Closed | Should -BeTrue
        Test-FileWritable $p | Should -BeTrue
        $r.Message | Should -BeExactly 'Locked.xlsx saved and closed in Excel.'
    }

    It 'does not save a workbook that has no changes' {
        $p = New-TFile 'Clean.xlsx'
        $wb = New-FakeWorkbook -FullName $p -Saved $true
        $r = Save-FaisWorkbook -Path $p -Close -ExcelApp (New-FakeExcel @($wb))
        $wb.SaveCount | Should -Be 0
        $r.Saved | Should -BeFalse
        $r.Closed | Should -BeTrue
    }

    It 'leaves a read-only workbook with changes open (closing would lose them)' {
        $p = New-TFile 'RO.xlsx'
        $wb = New-FakeWorkbook -FullName $p -Saved $false -ReadOnly $true
        $r = Save-FaisWorkbook -Path $p -Close -ExcelApp (New-FakeExcel @($wb))
        $wb.Closed | Should -BeFalse
        $wb.SaveCount | Should -Be 0
        $r.Message | Should -BeLike '*read-only*left it open*'
    }

    It 'retries while Excel is busy, then saves' {
        $p = New-TFile 'Busy.xlsx'
        $wb = New-FakeWorkbook -FullName $p -Saved $false -BusyCalls 2
        $r = Save-FaisWorkbook -Path $p -ExcelApp (New-FakeExcel @($wb))
        $r.Saved | Should -BeTrue
        $wb.SaveCount | Should -Be 1
    }

    It 'reports (does not throw) when Excel stays busy' {
        $p = New-TFile 'Stuck.xlsx'
        $wb = New-FakeWorkbook -FullName $p -Saved $false -BusyCalls 99
        $r = Save-FaisWorkbook -Path $p -Close -ExcelApp (New-FakeExcel @($wb))
        $r.Saved | Should -BeFalse
        $r.Closed | Should -BeFalse
        $r.Message | Should -BeLike 'Could not save Stuck.xlsx in Excel: Excel is busy*'
    }

    It 'a file Excel does not have open is reported as not open' {
        $p = New-TFile 'NotOpen.xlsx'
        $r = Save-FaisWorkbook -Path $p -ExcelApp (New-FakeExcel @())
        $r.WasOpen | Should -BeFalse
        $r.Message | Should -BeExactly ''
    }

    It 'without Excel (or off Windows) it does nothing' {
        $r = Save-FaisWorkbook -Path (New-TFile 'NoExcel.xlsx')
        $r.WasOpen | Should -BeFalse
    }
}

Describe 'Close-FaisExcelFiles' {
    It 'saves and closes only the listed files' {
        $a = New-TFile 'A.xlsx'; $b = New-TFile 'B.xlsx'; $c = New-TFile 'C.xlsx'
        $wa = New-FakeWorkbook -FullName $a -Saved $false
        $wb = New-FakeWorkbook -FullName $b
        $wc = New-FakeWorkbook -FullName $c -Saved $false
        $res = @(Close-FaisExcelFiles -Path @($a, $b, (Join-Path $TestDrive 'missing.xlsx')) -ExcelApp (New-FakeExcel @($wa, $wb, $wc)) -Quiet)
        $res.Count | Should -Be 2
        $wa.Closed | Should -BeTrue
        $wa.SaveCount | Should -Be 1
        $wb.Closed | Should -BeTrue
        $wc.Closed | Should -BeFalse
        $wc.SaveCount | Should -Be 0
    }

    It '-SaveOnly saves but keeps the workbooks open' {
        $a = New-TFile 'SA.xlsx'
        $wa = New-FakeWorkbook -FullName $a -Saved $false
        [void](Close-FaisExcelFiles -Path @($a) -SaveOnly -ExcelApp (New-FakeExcel @($wa)) -Quiet)
        $wa.SaveCount | Should -Be 1
        $wa.Closed | Should -BeFalse
    }

    It '-All also closes other saved workbooks, but never one that was never saved' {
        $a = New-TFile 'AllA.xlsx'
        $wa = New-FakeWorkbook -FullName $a
        $other = New-FakeWorkbook -FullName (Join-Path $TestDrive 'Budget.xlsx') -Saved $false
        $book1 = New-FakeWorkbook -FullName 'Book1' -Saved $false -Dir ''
        $book1.Name = 'Book1'
        $personal = New-FakeWorkbook -FullName 'C:\Users\me\AppData\Roaming\Microsoft\Excel\XLSTART\PERSONAL.XLSB'
        $res = @(Close-FaisExcelFiles -Path @($a) -All -ExcelApp (New-FakeExcel @($wa, $other, $book1, $personal)) -Quiet)
        $personal.Closed | Should -BeFalse
        $wa.Closed | Should -BeTrue
        $other.SaveCount | Should -Be 1
        $other.Closed | Should -BeTrue
        $book1.Closed | Should -BeFalse
        @($res | Where-Object { $_.Message -like '*never been saved*' }).Count | Should -Be 1
    }

    It 'without Excel returns nothing' {
        @(Close-FaisExcelFiles -Path @((New-TFile 'q.xlsx')) -Quiet).Count | Should -Be 0
    }
}

Describe 'Request-FaisFileWritable' {
    It 'a free file is writable straight away' {
        Request-FaisFileWritable -Path (New-TFile 'free.xlsx') -Prompt { param($x) $false } | Should -BeTrue
    }

    It 'a file open in Excel is saved, closed, recorded for reopening and then writable' {
        $p = New-TFile 'Out.xlsx'
        $wb = New-FakeWorkbook -FullName $p -Saved $false -Lock
        $closed = New-Object System.Collections.Generic.List[string]
        $ok = Request-FaisFileWritable -Path $p -Prompt { param($x) $false } -Closed $closed -ExcelApp (New-FakeExcel @($wb)) 6>$null
        $ok | Should -BeTrue
        $wb.SaveCount | Should -Be 1
        ($closed -join ',') | Should -BeExactly $p
    }

    It 'falls back to the prompt when something else holds the file' {
        $p = New-TFile 'Other.xlsx'
        $fs = [System.IO.File]::Open($p, 'Open', 'ReadWrite', 'None')
        $script:asked = 0
        try {
            $ok = Request-FaisFileWritable -Path $p -Prompt { param($x) $script:asked++; $false } -ExcelApp (New-FakeExcel @()) 6>$null
            $ok | Should -BeFalse
            $script:asked | Should -Be 1
        }
        finally { $fs.Close() }
    }
}
