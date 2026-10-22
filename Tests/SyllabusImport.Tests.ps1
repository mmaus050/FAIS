# Pester 5 tests for Update Syllabi (Modules\SyllabusImport.psm1): matching names, what each
# training file line does to an event, the plan, the update in memory, reading the flying
# schedule and the Big Board, and saving to OutputSchedule / MasterData. The workbook parts
# need ImportExcel and are skipped without it. Synthetic names and files only.
# Test syllabus sim debt: FT5 / FF5 +1 (SimStopMult), FI3 / FE3 / FN3 -0.5.

BeforeDiscovery {
    $script:NoExcel = $null -eq (Get-Module -ListAvailable -Name ImportExcel)
}

BeforeAll {
    . "$PSScriptRoot/Fixtures/New-FaisTestData.ps1"
    Import-FaisModules
    if (Get-Module -ListAvailable -Name ImportExcel) { Import-Module ImportExcel 3>$null }

    $script:SW = New-TestSoftWeights
    $script:Now = [datetime]'2026-09-23 12:00'

    # A Read-TrainingFile result built by hand
    function New-TFile {
        param ([string]$Kind = 'DailySchedule', [array]$Facts = @(), [string[]]$People = @(), [string[]]$Events = @())
        if ($People.Count -eq 0) { $People = @($Facts | ForEach-Object { $_.Student } | Sort-Object -Unique) }
        [PSCustomObject]@{ Path = 'x.xlsx'; Kind = $Kind; Label = $(if ($Kind -eq 'BigBoard') { 'Big Board (E FLT)' } else { 'Flying schedule for Mon 21 Sep 2026' })
            Day = [datetime]'2026-09-21'; Sheets = @('Flying (21-Sep-2026)', 'Sims (21-Sep-2026)'); Events = $Events; Facts = $Facts; People = $People; Error = '' }
    }

    # One line of a flying schedule (Status CM / EXEC / INC / CX), or a Big Board entry (-Board; no -At = no entry)
    function New-TFact {
        param ([string]$Student, [string]$Mission, [string]$Status = 'CM', [string]$IP = '', [string]$At = '', [switch]$Board)
        $etd = if ($At) { [datetime]$At } else { $null }
        if ($Board) {
            $has = ($At -ne '' -or $IP -ne '')
            return [PSCustomObject]@{ Student = $Student; Mission = $Mission; Complete = $has; Explicit = $has; Status = ''; Instructor = $IP; ETD = $etd }
        }
        [PSCustomObject]@{ Student = $Student; Mission = $Mission; Complete = ($Status -in @('CM', 'EXEC')); Explicit = $true; Status = $Status; Instructor = $IP; ETD = $etd }
    }

    function Get-Entry { param ($Student, [string]$Mission) $Student.Syllabus | Where-Object { $_.Mission -eq $Mission } | Select-Object -First 1 }

    # Writes a workbook from @{ sheet = rows } (each row an array of cells; $null = empty).
    # -Merge: @{ sheet = @('G5:J5', ...) }
    function New-TBook {
        param ([string]$Name, [System.Collections.IDictionary]$Sheets, [hashtable]$Merge = @{})
        $path = Join-Path $TestDrive $Name
        if (Test-Path $path) { Remove-Item $path -Force }
        $pkg = Open-ExcelPackage -Path $path -Create
        foreach ($sheet in $Sheets.Keys) {
            $ws = Add-Worksheet -ExcelPackage $pkg -WorksheetName $sheet
            $r = 1
            foreach ($row in $Sheets[$sheet]) {
                $c = 1
                foreach ($cell in @($row)) {
                    if ($null -ne $cell) { $ws.Cells[$r, $c].Value = $cell }
                    $c++
                }
                $r++
            }
            foreach ($m in @($Merge[$sheet])) { if ($m) { $ws.Cells[$m].Merge = $true } }
        }
        Close-ExcelPackage $pkg
        return $path
    }
}


Describe 'Matching names' {
    It 'reads a comma name one way: last name, then first names (no rank tag, case or periods)' {
        $p = ConvertTo-PersonName 'Quill, Cy  D. [2d Lt]'
        @($p.Readings).Count | Should -Be 1
        $p.Readings[0].Last | Should -BeExactly 'quill'
        $p.Readings[0].Given -join ' ' | Should -BeExactly 'cy d'
        (ConvertTo-PersonName 'Diaz-Ortega, Dee').Readings[0].Last | Should -BeExactly 'diaz ortega'
        (ConvertTo-PersonName 'Quill Jr., Cy').Readings[0].Last | Should -BeExactly 'quill'
    }

    It 'reads a name with no comma every way it can be: all last name, first names first, and last name first' {
        @((ConvertTo-PersonName 'Quill').Readings | ForEach-Object { "$($_.Last)/$($_.Given -join ' ')/$($_.Reversed)" }) | Should -Be @('quill//False')
        @((ConvertTo-PersonName '2d Lt Cy D. Quill').Readings | ForEach-Object { "$($_.Last)/$($_.Given -join ' ')/$($_.Reversed)" }) | Should -Be @(
            'cy d quill//False', 'cy/d quill/True', 'quill/cy d/False', 'cy d/quill/True')
    }

    It '<Known> vs <Other> is <Want>' -ForEach @(
        # Last, First on both sides
        @{ Known = 'Quill, Cy D.'; Other = 'QUILL, CY D. [2d Lt]'; Want = 'Exact' }
        @{ Known = 'Quill, Cy'; Other = 'Quill, Cy D. [2d Lt]'; Want = 'Given' }
        @{ Known = 'Quill, C.'; Other = 'Quill, Cy D. [2d Lt]'; Want = 'Given' }
        @{ Known = 'Quill, Cy D.'; Other = 'Quill, Cy E. [2d Lt]'; Want = '' }
        @{ Known = 'Quill, Cy'; Other = 'Quill, Al [2d Lt]'; Want = '' }
        @{ Known = 'Quill, Cy'; Other = 'Quill, Cal [2d Lt]'; Want = '' }
        @{ Known = 'Quill'; Other = 'Quill, Cy D. [2d Lt]'; Want = 'LastOnly' }
        @{ Known = 'Quill, Cy'; Other = 'Quillon, Cy [2d Lt]'; Want = '' }
        @{ Known = 'Diaz Ortega, Dee'; Other = 'DIAZ-ORTEGA, Dee M. [Capt]'; Want = 'Given' }
        # First Last (the Big Board), with loose ranks, suffixes and two-word last names
        @{ Known = 'Quill, Cy D.'; Other = 'Cy D. Quill'; Want = 'Exact' }
        @{ Known = 'Quill, Cy D.'; Other = 'CY QUILL'; Want = 'Given' }
        @{ Known = 'Quill, Cy D.'; Other = 'Cy Dale Quill'; Want = 'Given' }
        @{ Known = 'Quill, Cy D.'; Other = '2d Lt Cy D. Quill'; Want = 'Exact' }
        @{ Known = 'Quill, Cy'; Other = 'Capt Cy Quill Jr.'; Want = 'Exact' }
        @{ Known = 'Quill'; Other = 'Cy D. Quill'; Want = 'LastOnly' }
        @{ Known = 'Vega Ortiz, Ana M.'; Other = 'Ana M. Vega Ortiz'; Want = 'Exact' }
        @{ Known = 'Vega Ortiz'; Other = 'Ana Vega-Ortiz'; Want = 'LastOnly' }
        @{ Known = 'Diaz Ortega, Dee'; Other = 'Dee Diaz-Ortega'; Want = 'Exact' }
        @{ Known = 'Quill, Cy'; Other = 'Al Quill'; Want = '' }
        @{ Known = 'Casey'; Other = 'Casey Quill'; Want = '' }
        # A FAIS name typed first name first
        @{ Known = 'Cy Quill'; Other = 'Quill, Cy D. [2d Lt]'; Want = 'Given' }
        # Last First with no comma: only when the first names agree too
        @{ Known = 'Quill, Cy D.'; Other = 'Quill Cy D'; Want = 'ReversedExact' }
        @{ Known = 'Quill, Cy'; Other = 'QUILL CY D'; Want = 'ReversedGiven' }
        @{ Known = 'Quill'; Other = 'Quill Cy D'; Want = '' }
    ) {
        $m = Get-PersonNameMatch -Known (ConvertTo-PersonName $Known) -Other (ConvertTo-PersonName $Other)
        $(if ($m) { $m.Kind } else { '' }) | Should -BeExactly $Want
    }

    It 'matches the Big Board''s First Last names to FAIS''s Last, First and last-name-only students' {
        $students = @((New-TestStudent -Name 'Quill, Cy D.'), (New-TestStudent -Name 'Adams, Al'), (New-TestStudent -Name 'Vance'), (New-TestStudent -Name 'Vega Ortiz, Ana'), (New-TestStudent -Name 'Casey'))
        $m = Resolve-TrainingStudents -Names @('Cy D. Quill', 'AL B ADAMS', 'Max Vance', 'Ana M. Vega Ortiz', 'Casey Moss', 'Zoe Zed') -Students $students
        $m.Map['Cy D. Quill'].Name | Should -BeExactly 'Quill, Cy D.'
        $m.Map['AL B ADAMS'].Name | Should -BeExactly 'Adams, Al'
        $m.Map['Max Vance'].Name | Should -BeExactly 'Vance'
        $m.Map['Ana M. Vega Ortiz'].Name | Should -BeExactly 'Vega Ortiz, Ana'
        @($m.Unknown) | Should -Be @('Casey Moss', 'Zoe Zed') -Because 'Casey is Moss''s first name, not the FAIS student Casey'
        @($m.Ambiguous).Count | Should -Be 0
    }

    It 'a last-name-only student is left alone when two First Last people share the last name' {
        $m = Resolve-TrainingStudents -Names @('Max Vance', 'Lou K. Vance') -Students @(New-TestStudent -Name 'Vance')
        $m.Map.Count | Should -Be 0
        @($m.Ambiguous) | Should -Be @('Vance (more than one person on the file is called Vance)')
    }

    It 'a Last First name with no comma matches only when nothing better does' {
        $students = @((New-TestStudent -Name 'Quill, Cy'), (New-TestStudent -Name 'Cy, Quill'))
        $m = Resolve-TrainingStudents -Names @('Quill Cy') -Students $students
        $m.Map['Quill Cy'].Name | Should -BeExactly 'Cy, Quill' -Because 'read first name first, Quill Cy is Quill (first) Cy (last)'
        (Resolve-TrainingStudents -Names @('Quill Cy') -Students @($students[0])).Map['Quill Cy'].Name | Should -BeExactly 'Quill, Cy'
    }

    It 'matches the wing-wide schedule to FAIS students on first name too, never on last name alone' {
        $students = @((New-TestStudent -Name 'Smith, John'), (New-TestStudent -Name 'Adams, Al'))
        $m = Resolve-TrainingStudents -Names @('Smith, Rory M. [2d Lt]', 'Adams, Al B. [2d Lt]', 'Zed, Zoe [2d Lt]') -Students $students
        $m.Map['Adams, Al B. [2d Lt]'].Name | Should -BeExactly 'Adams, Al'
        $m.Map.ContainsKey('Smith, Rory M. [2d Lt]') | Should -BeFalse
        @($m.Unknown) | Should -Be @('Smith, Rory M. [2d Lt]', 'Zed, Zoe [2d Lt]')
        @($m.Ambiguous).Count | Should -Be 0
    }

    It 'a student FAIS knows by last name only matches when the file has just one person with it' {
        $students = @(New-TestStudent -Name 'Vance')
        (Resolve-TrainingStudents -Names @('Vance, Lou K. [2d Lt]') -Students $students).Map['Vance, Lou K. [2d Lt]'].Name | Should -BeExactly 'Vance'
        $two = Resolve-TrainingStudents -Names @('Vance, Lou K. [2d Lt]', 'Vance, Max [Capt]') -Students $students
        $two.Map.Count | Should -Be 0
        @($two.Ambiguous) | Should -Be @('Vance (more than one person on the file is called Vance)')
    }

    It 'a name that could be two students is left alone' {
        $students = @((New-TestStudent -Name 'Lee, Ann'), (New-TestStudent -Name 'Lee, Al'))
        $m = Resolve-TrainingStudents -Names @('Lee, A. [2d Lt]') -Students $students
        $m.Map.Count | Should -Be 0
        @($m.Ambiguous) | Should -Be @('Lee, A. [2d Lt] (could be Lee, Ann or Lee, Al)')
    }

    It 'a student who could be two people on the file is left alone' {
        $students = @(New-TestStudent -Name 'Lee, A.')
        $m = Resolve-TrainingStudents -Names @('Lee, Ann [2d Lt]', 'Lee, Al [2d Lt]') -Students $students
        $m.Map.Count | Should -Be 0
        @($m.Ambiguous) | Should -Be @('Lee, A. (the file has Lee, Al [2d Lt] and Lee, Ann [2d Lt])')
    }

    It 'the same student written two ways on the file is still one student' {
        $students = @(New-TestStudent -Name 'Quill, Cy D.')
        $m = Resolve-TrainingStudents -Names @('Quill, Cy D. [2d Lt]', 'Quill, Cy') -Students $students
        $m.Map.Count | Should -Be 2
        @($m.Ambiguous).Count | Should -Be 0
    }

    It 'an IP is written the way FAIS spells them, else as the file has them without the rank' {
        $ips = New-TestInstructors
        Resolve-TrainingInstructor -Name 'JONES, BOB K. [Maj]' -Instructors $ips | Should -BeExactly 'Jones, Bob'
        Resolve-TrainingInstructor -Name 'Ortiz, S. [Capt]' -Instructors $ips | Should -BeExactly 'Ortiz, Sam'
        Resolve-TrainingInstructor -Name 'PARR, PAT  Q. [CIV]' -Instructors $ips | Should -BeExactly 'PARR, PAT Q.'
        Resolve-TrainingInstructor -Name '' -Instructors $ips | Should -BeExactly ''
        Resolve-TrainingInstructor -Name 'Ann Lee' -Instructors $ips | Should -BeExactly 'Lee, Ann'
        Resolve-TrainingInstructor -Name 'Maj Tom Leech' -Instructors $ips | Should -BeExactly 'Leech, Tom'
        Resolve-TrainingInstructor -Name 'Pat Q. Parr' -Instructors $ips | Should -BeExactly 'Pat Q. Parr'
    }
}


Describe 'Get-SyllabusTargetResult (what one line does to one event)' {
    BeforeAll {
        function Cur { param ([bool]$Done, [string]$IP = '', [string]$At = '') [PSCustomObject]@{ IsComplete = $Done; Instructor = $IP; ETD = $At } }
        function Tgt { param ([bool]$Done, [string]$IP = '', [string]$At = '', [bool]$Explicit = $true)
            [PSCustomObject]@{ Complete = $Done; Explicit = $Explicit; Instructor = $IP; ETD = $(if ($At) { [datetime]$At } else { $null }) } }
    }

    It 'complete on an open event: complete, with the line''s IP and ETD' {
        $r = Get-SyllabusTargetResult -Current (Cur $false) -Target (Tgt $true 'Jones, Bob' '2026-09-21 07:30') -Now $script:Now
        $r.Kind | Should -BeExactly 'Complete'
        $r.Instructor | Should -BeExactly 'Jones, Bob'
        $r.ETD | Should -BeExactly '2026-09-21 07:30:00'
    }

    It 'complete on an event already complete the same way: nothing to do' {
        Get-SyllabusTargetResult -Current (Cur $true 'Jones, Bob' '2026-09-21 07:30:00') -Target (Tgt $true 'Jones, Bob' '2026-09-21 07:30') -Now $script:Now | Should -BeNullOrEmpty
    }

    It 'complete on an event complete with another IP or time: Updated' {
        $r = Get-SyllabusTargetResult -Current (Cur $true 'Lee, Ann' '2026-09-22 09:00:00') -Target (Tgt $true 'Jones, Bob' '2026-09-21 07:30') -Now $script:Now
        $r.Kind | Should -BeExactly 'Updated'
        $r.ETD | Should -BeExactly '2026-09-21 07:30:00'
    }

    It 'a line with no IP or ETD keeps the event''s own' {
        Get-SyllabusTargetResult -Current (Cur $true 'SIM IP' '2026-09-21 07:30:00') -Target (Tgt $true) -Now $script:Now | Should -BeNullOrEmpty
        $r = Get-SyllabusTargetResult -Current (Cur $false) -Target (Tgt $true) -Now $script:Now
        $r.Kind | Should -BeExactly 'Complete'
        $r.Instructor | Should -BeExactly ''
        $r.ETD | Should -BeExactly ''
    }

    It 'INC / CX on a complete event: not complete, IP and ETD cleared' {
        $r = Get-SyllabusTargetResult -Current (Cur $true 'Jones, Bob' '2026-09-25 07:30:00') -Target (Tgt $false) -Now $script:Now
        $r.Kind | Should -BeExactly 'NotComplete'
        $r.IsComplete | Should -BeFalse
        $r.Instructor + $r.ETD | Should -BeExactly ''
    }

    It 'not complete on an open event: nothing to do' {
        Get-SyllabusTargetResult -Current (Cur $false) -Target (Tgt $false) -Now $script:Now | Should -BeNullOrEmpty
    }

    It 'the Big Board''s empty columns leave an event FAIS planned after now, and take back one before' {
        (Get-SyllabusTargetResult -Current (Cur $true 'Jones, Bob' '2026-09-24 07:30:00') -Target (Tgt $false -Explicit $false) -Now $script:Now).Kind | Should -BeExactly 'Planned'
        (Get-SyllabusTargetResult -Current (Cur $true 'Jones, Bob' '2026-09-22 07:30:00') -Target (Tgt $false -Explicit $false) -Now $script:Now).Kind | Should -BeExactly 'NotComplete'
    }

    It 'reads tab cells as Excel has them (TRUE text, real dates)' {
        Get-SyllabusTargetResult -Current ([PSCustomObject]@{ IsComplete = 'TRUE'; Instructor = 'Jones, Bob'; ETD = [datetime]'2026-09-21 07:30' }) -Target (Tgt $true 'Jones, Bob' '2026-09-21 07:30') -Now $script:Now | Should -BeNullOrEmpty
    }
}


Describe 'The plan and the update in memory (flying schedule)' {
    BeforeEach {
        $script:Adams = New-TestStudent -Name 'Adams, Al' -Done @('FI3101|2026-09-18 08:00', 'FE3102|2026-09-19 08:00', 'FT5001|2026-09-21 13:00|Lee, Ann') -SimDebt 2
        $script:Brown = New-TestStudent -Name 'Brown, Bea' -Done @('FI3101|2026-09-18 08:00') -SimDebt 1
        $script:Students = @($script:Adams, $script:Brown)
        $script:File = New-TFile -Facts @(
            New-TFact 'Adams, Al [2d Lt]' 'FT5001' 'CX' 'LEE, ANN [Capt]' '2026-09-21 13:00'
            New-TFact 'Brown, Bea [2d Lt]' 'FE3102' 'CM' 'PARR, PAT [CIV]' '2026-09-21 09:30'
            New-TFact 'Brown, Bea [2d Lt]' 'FT5001' 'INC' 'JONES, BOB [Maj]' '2026-09-21 07:00'
            New-TFact 'Brown, Bea [2d Lt]' 'FT5001' 'EXEC' 'JONES, BOB [Maj]' '2026-09-21 11:00'
            New-TFact 'Brown, Bea [2d Lt]' 'T6 OFT' 'EXEC' 'PARR, PAT [CIV]' '2026-09-21 15:00'
            New-TFact 'Other, Ollie [2d Lt]' 'FT5001' 'CM' 'JONES, BOB [Maj]' '2026-09-21 07:30'
        )
        $script:Plan = Get-SyllabusUpdatePlan -File $script:File -Students $script:Students -Instructors (New-TestInstructors) -SoftWeights $script:SW -Now $script:Now
    }

    It 'CX takes back an event; CM / EXEC complete one; a later complete line wins over INC' {
        $c = @($script:Plan.Changes | Sort-Object Student, Mission | ForEach-Object { "$($_.Student) $($_.Mission) $($_.Kind) $($_.Instructor) $($_.ETD)".Trim() })
        $c | Should -Be @(
            'Adams, Al FT5001 NotComplete'
            'Brown, Bea FE3102 Complete PARR, PAT 2026-09-21 09:30:00'
            'Brown, Bea FT5001 Complete Jones, Bob 2026-09-21 11:00:00'
        )
    }

    It 'says what it left alone: events not in the syllabus and people who are not FAIS students' {
        @($script:Plan.NotInSyllabus) | Should -Be @('Brown, Bea T6 OFT')
        @($script:Plan.Unknown) | Should -Be @('Other, Ollie [2d Lt]')
        @($script:Plan.Matched) | Should -Be @('Adams, Al', 'Brown, Bea')
        @($script:Plan.NotOnFile).Count | Should -Be 0 -Because 'only the Big Board lists the students it does not have'
    }

    It 'the plan changes nothing by itself' {
        (Get-Entry $script:Adams 'FT5001').IsComplete | Should -BeTrue
        (Get-Entry $script:Brown 'FE3102').IsComplete | Should -BeFalse
        $script:Brown.SimDebt | Should -Be 1
    }

    It 'the sim debt each change brings' {
        ($script:Plan.Changes | Where-Object { $_.Student -eq 'Adams, Al' }).SimDebt | Should -Be -1
        @($script:Plan.Changes | Where-Object { $_.Student -eq 'Brown, Bea' } | ForEach-Object { $_.SimDebt }) -join ',' | Should -BeExactly '-0.5,1'
    }

    It 'Invoke-SyllabusUpdate makes the changes: syllabus, IP, ETD and SimDebt, and leaves unsat flags alone' {
        $script:Brown.UnsatFlight = $true
        $rev = [int](Get-OptionalProperty $script:Brown 'SyllabusRevision' 0)
        $done = @(Invoke-SyllabusUpdate -Targets $script:Plan.Targets -SoftWeights $script:SW -Now $script:Now)
        $done.Count | Should -Be 3
        $a = Get-Entry $script:Adams 'FT5001'
        $a.IsComplete | Should -BeFalse
        $a.Instructor + $a.ETD | Should -BeExactly ''
        $script:Adams.SimDebt | Should -Be 1
        $b = Get-Entry $script:Brown 'FT5001'
        $b.IsComplete | Should -BeTrue
        $b.Instructor | Should -BeExactly 'Jones, Bob'
        $b.ETD | Should -BeExactly '2026-09-21 11:00:00'
        (Get-Entry $script:Brown 'FE3102').Instructor | Should -BeExactly 'PARR, PAT'
        $script:Brown.SimDebt | Should -Be 1.5
        $script:Brown.UnsatFlight | Should -BeTrue
        [int]$script:Brown.SyllabusRevision | Should -BeGreaterThan $rev
    }

    It 'running it again changes nothing' {
        [void](Invoke-SyllabusUpdate -Targets $script:Plan.Targets -SoftWeights $script:SW -Now $script:Now)
        $again = Get-SyllabusUpdatePlan -File $script:File -Students $script:Students -Instructors (New-TestInstructors) -SoftWeights $script:SW -Now $script:Now
        @($again.Changes).Count | Should -Be 0
        $script:Brown.SimDebt | Should -Be 1.5
    }

    It 'the summary lists each student''s changes and what was left alone' {
        $text = Format-SyllabusUpdateSummary -File $script:File -Plan $script:Plan
        $text | Should -Match 'Flying schedule for Mon 21 Sep 2026'
        $text | Should -Match 'FAIS students on it: 2'
        $text | Should -Match ([regex]::Escape('Changes for 2 student(s): 2 event(s) marked complete, 1 marked NOT complete, 0 IP / ETD updated.'))
        $text | Should -Match ([regex]::Escape('Adams, Al: NOT complete FT5001 (CX); SimDebt -1'))
        $text | Should -Match ([regex]::Escape('Brown, Bea: complete FE3102 FT5001; SimDebt +0.5'))
        $text | Should -Match ([regex]::Escape("Not in the student's syllabus: Brown, Bea T6 OFT"))
        $text | Should -Match ([regex]::Escape('1 other people on the schedule are not FAIS students.'))
    }

    It 'the summary stays short enough for a message box: students and events per student are capped' {
        $many = @(foreach ($s in 1..20) { foreach ($m in 1..12) { [PSCustomObject]@{ Student = ('Student, S{0:00}' -f $s); Mission = ('FT51{0:00}' -f $m); Kind = 'Complete'; Instructor = ''; ETD = ''; Status = 'CM'; SimDebt = 0 } } })
        $plan = [PSCustomObject]@{ Changes = $many; Planned = @(); Matched = @(1..20); Unknown = @(); Ambiguous = @(); NotInSyllabus = @(); NotOnFile = @() }
        $text = Format-SyllabusUpdateSummary -File $script:File -Plan $plan
        $text | Should -Match ([regex]::Escape('Student, S01: complete FT5101 FT5102 FT5103 FT5104 FT5105 FT5106 FT5107 FT5108 +4 more'))
        $text | Should -Match ([regex]::Escape('...and 5 more student(s) (the console lists them all)'))
        @($text -split "`n").Count | Should -BeLessThan 25
        $all = Format-SyllabusUpdateSummary -File $script:File -Plan $plan -MaxLines 100000 -MaxMissions 100000
        $all | Should -Match '(?m)^  Student, S20: complete FT5101 .* FT5112$'
    }

    It 'the summary says so when every syllabus already matches' {
        $plan = Get-SyllabusUpdatePlan -File (New-TFile -Facts @(New-TFact 'Adams, Al' 'FI3101' 'CM')) -Students $script:Students -SoftWeights $script:SW -Now $script:Now
        $plan.Changes.Count | Should -Be 0
        Format-SyllabusUpdateSummary -File $script:File -Plan $plan | Should -Match 'nothing to change'
    }
}


Describe 'The plan (Big Board)' {
    BeforeEach {
        # FT5102 is planned for 24 Sep (after now), FT5001 was flown 21 Sep
        $script:Adams = New-TestStudent -Name 'Adams, Al B.' -Done @('FI3101|2026-09-18 08:00', 'FT5001|2026-09-21 08:00|Lee, Ann', 'FT5102|2026-09-24 08:00|Jones, Bob')
        $script:Brown = New-TestStudent -Name 'Brown, Bea' -Done @('FI3101|2026-09-18 08:00', 'FE3102|2026-09-20 08:00')
        $script:Quill = New-TestStudent -Name 'Quill, Cy'
        $script:Students = @($script:Adams, $script:Brown, $script:Quill)
        $events = @('BA101', 'FI3101', 'FE3102', 'FT5001', 'FT5102')
        $script:File = New-TFile -Kind 'BigBoard' -Events $events -People @('Al B. Adams', 'Bea Brown', 'Vic Visitor') -Facts @(
            New-TFact 'Al B. Adams' 'BA101' -At '2026-07-24 15:00' -IP 'Instructor, Ivy' -Board
            New-TFact 'Al B. Adams' 'FI3101' -At '2026-09-18 08:00' -IP 'SIM IP' -Board
            New-TFact 'Al B. Adams' 'FE3102' -At '2026-09-19 09:00' -IP 'PARR, PAT' -Board
            New-TFact 'Al B. Adams' 'FT5001' -Board
            New-TFact 'Al B. Adams' 'FT5102' -Board
            New-TFact 'Bea Brown' 'BA101' -Board
            New-TFact 'Bea Brown' 'FI3101' -At '2026-09-18 08:00' -Board
            New-TFact 'Bea Brown' 'FE3102' -Board
            New-TFact 'Bea Brown' 'FT5001' -Board
            New-TFact 'Bea Brown' 'FT5102' -Board
        )
        $script:Plan = Get-SyllabusUpdatePlan -File $script:File -Students $script:Students -Instructors (New-TestInstructors) -SoftWeights $script:SW -Now $script:Now
    }

    It 'an entry completes the event, an empty block takes it back; events not in the syllabus are ignored' {
        @($script:Plan.Changes | Sort-Object Student, Mission | ForEach-Object { "$($_.Student) $($_.Mission) $($_.Kind)" }) | Should -Be @(
            'Adams, Al B. FE3102 Complete'
            'Adams, Al B. FT5001 NotComplete'
            'Brown, Bea FE3102 NotComplete'
        )
        @($script:Plan.NotInSyllabus).Count | Should -Be 0
    }

    It 'an event FAIS planned after now is left as planned' {
        @($script:Plan.Planned | ForEach-Object { "$($_.Student) $($_.Mission)" }) | Should -Be @('Adams, Al B. FT5102')
        [void](Invoke-SyllabusUpdate -Targets $script:Plan.Targets -SoftWeights $script:SW -Now $script:Now)
        (Get-Entry $script:Adams 'FT5102').IsComplete | Should -BeTrue
        (Get-Entry $script:Adams 'FT5001').IsComplete | Should -BeFalse
    }

    It 'an entry with no IP keeps the IP an event already had' {
        $t = $script:Plan.Targets | Where-Object { $_.Student.Name -eq 'Brown, Bea' -and $_.Mission -eq 'FI3101' }
        $t.Instructor | Should -BeExactly ''
        @($script:Plan.Changes | Where-Object { $_.Student -eq 'Brown, Bea' -and $_.Mission -eq 'FI3101' }).Count | Should -Be 0
    }

    It 'lists the FAIS students the Big Board does not have, and its people who are not FAIS students' {
        @($script:Plan.NotOnFile) | Should -Be @('Quill, Cy')
        @($script:Plan.Unknown) | Should -Be @('Vic Visitor')
        $text = Format-SyllabusUpdateSummary -File $script:File -Plan $script:Plan
        $text | Should -Match ([regex]::Escape('FAIS students not on the Big Board: Quill, Cy'))
        $text | Should -Match ([regex]::Escape('Not FAIS students: Vic Visitor'))
        $text | Should -Match ([regex]::Escape("Left as planned: 1 event(s) FAIS has scheduled after now, which the Big Board can't show yet."))
    }
}


Describe 'Reading the training files' -Skip:$script:NoExcel {
    BeforeAll {
        $flying = @(
            , @('Line', 'Go', 'ETD', 'Student', 'Stu Stat', 'Mission', 'IP GNG', 'Instructor', 'IP Event', 'Sortie Type', 'Line Status', 'CX Reason', 'Event Status')
            , @('1', 1, ([datetime]'2026-09-21 06:45'), 'Adams, Al B. [2d Lt]', '', 'FT5001', $null, 'Jones, Bob [Maj]', '', 'Day Local', 'CM', '', 'CM')
            , @('2', 1, ([datetime]'2026-09-21 07:15'), 'Brown, Bea [2d Lt]', '', 'FT5001', $null, 'Lee, Ann [Capt]', '', 'Day Local', 'CX', 'WX', 'CNX')
            , @('3', 1, '1300', 'Quill, Cy [2d Lt]', '', 'FT5102', $null, 'Ortiz, Sam [Capt]', '', 'O&B', 'exec', '', '')
            , @('4', 1, ([datetime]'2026-09-21 14:00'), 'Diaz, Dee [2d Lt]', '', 'FT5103', $null, 'Ortiz, Sam [Capt]', '', 'Day Local', 'INC', '', 'INC')
            , @('5', 1, ([datetime]'2026-09-21 15:00'), 'Eng, Eve [2d Lt]', '', 'FT5104', $null, 'Jones, Bob [Maj]', '', 'Day Local', '', '', '')
            , @('6', 1, ([datetime]'2026-09-21 15:30'), 'For Sale', '', '', $null, 'Lee, Ann [Capt]', '', 'Day Local', 'CX', 'PN', '')
            , @('7', 1, ([datetime]'2026-09-21 16:00'), $null, '', 'FT5001', $null, 'Lee, Ann [Capt]', '', 'Day Local', 'CM', '', '')
        )
        $sims = @(
            , @('SQ', 'Duration', 'IP GNG', 'Device', 'Simulator Instructor', 'IP Event', 'Clr Tag', 'SP GNG', 'Student', 'Event Name', 'Scheduled Line Status', 'Scheduled Line Status Reason', 'REMARK', 'Start Time')
            , @('E FLT', 1.3, $null, 'OFT 2', 'PARR, PAT Q. [CIV]', '', '', $null, 'Brown, Bea [2d Lt]', 'FE3102', 'CM', '', $null, ([datetime]'2026-09-21 09:30'))
            , @('E FLT', 0, $null, 'IFT 5', 'ROE, RAY [CIV]', '', '', $null, 'For sale', '', 'CX', '', 'NO SP', ([datetime]'2026-09-21 11:30'))
        )
        $vr = @(, @('ORG', 'Flight', 'Start Time', 'Sim Type', 'Sim', 'Instructor', 'Student', 'Event', 'Syl Dur', 'Remarks'))
        $script:Daily = New-TBook 'Daily.xlsx' ([ordered]@{ 'Flying (21-Sep-2026)' = $flying; 'ValueList_Helper' = @(, @($null, 'x')); 'Sims (21-Sep-2026)' = $sims; 'VR SIM (21-Sep-2026)' = $vr })

        # Big Board: headings on row 5, each event over 4 columns (date, status, grade, IP); BA101 over 6
        $row5 = @($null, 'Student Name', 'Track', 'Org', $null, 'Class', 'BA101', $null, $null, $null, $null, $null, 'FI3101', $null, $null, $null, 'FT5001', $null, $null, $null, 'FT5102', $null, $null, $null)
        $adams = @($null, 'Al B. Adams', 'Default', 'E FLT', $null, '27-05', '07-24-2026-1500', $null, 'Complete', $null, $null, 'Instructor, Ivy', '09-18-2026-0800', 'Complete', $null, 'PARR, PAT Q.', '09-22-2026-0913', 'Complete', 'G', 'Jones, Bob', $null, $null, $null, $null)
        $brown = @($null, 'Bea Brown', 'Default', 'E FLT', $null, '27-05', $null, $null, $null, $null, $null, $null, '09-18-2026-0815', 'Complete', 'U', $null, '09-17-2026-0830', $null, $null, $null, $null, $null, $null, $null)
        $script:Board = New-TBook 'BigBoard.xlsx' ([ordered]@{ 'Big Board' = @(@($null), @($null, $null, 'Scheduling Aid for Organization E FLT'), @($null), @($null, $null, $null, $null, $null, $null, 'Default'), $row5, $adams, $brown, @($null, $null, $null, $null, 'For CUI Use Only')) }) `
            -Merge @{ 'Big Board' = @('G5:L5', 'M5:P5', 'Q5:T5', 'U5:X5', 'D5:E5', 'G6:H6', 'I6:J6') }

        $script:D = Read-TrainingFile -Path $script:Daily
        $script:B = Read-TrainingFile -Path $script:Board
    }

    It 'knows the flying schedule by its Line Status columns and reads its Flying and Sims tabs only' {
        $script:D.Error | Should -BeExactly ''
        $script:D.Kind | Should -BeExactly 'DailySchedule'
        @($script:D.Sheets) | Should -Be @('Flying (21-Sep-2026)', 'Sims (21-Sep-2026)')
        $script:D.Label | Should -BeExactly 'Flying schedule for Mon 21 Sep 2026'
    }

    It 'turns CM / EXEC into complete and INC / CX into not complete; other statuses, For Sale and blank students are skipped' {
        @($script:D.Facts | ForEach-Object { "$($_.Student)|$($_.Mission)|$($_.Status)|$($_.Complete)|$($_.Instructor)|$(if ($_.ETD) { $_.ETD.ToString('yyyy-MM-dd HHmm') })" }) | Should -Be @(
            'Adams, Al B. [2d Lt]|FT5001|CM|True|Jones, Bob [Maj]|2026-09-21 0645'
            'Brown, Bea [2d Lt]|FT5001|CX|False|Lee, Ann [Capt]|2026-09-21 0715'
            'Quill, Cy [2d Lt]|FT5102|EXEC|True|Ortiz, Sam [Capt]|2026-09-21 1300'
            'Diaz, Dee [2d Lt]|FT5103|INC|False|Ortiz, Sam [Capt]|2026-09-21 1400'
            'Brown, Bea [2d Lt]|FE3102|CM|True|PARR, PAT Q. [CIV]|2026-09-21 0930'
        )
        @($script:D.People) | Should -Contain 'Eng, Eve [2d Lt]'
        @($script:D.People) | Should -Not -Contain 'For Sale'
    }

    It 'knows the Big Board by its Student Name row and reads each event block' {
        $script:B.Error | Should -BeExactly ''
        $script:B.Kind | Should -BeExactly 'BigBoard'
        $script:B.Label | Should -BeExactly 'Big Board (Scheduling Aid for Organization E FLT)'
        @($script:B.Events) | Should -Be @('BA101', 'FI3101', 'FT5001', 'FT5102')
        @($script:B.People) | Should -Be @('Al B. Adams', 'Bea Brown')
    }

    It 'a Big Board entry is complete with its date and IP (not its status or grade); an empty block is not' {
        @($script:B.Facts | ForEach-Object { "$($_.Student)|$($_.Mission)|$($_.Complete)|$($_.Instructor)|$(if ($_.ETD) { $_.ETD.ToString('yyyy-MM-dd HHmm') })" }) | Should -Be @(
            'Al B. Adams|BA101|True|Instructor, Ivy|2026-07-24 1500'
            'Al B. Adams|FI3101|True|PARR, PAT Q.|2026-09-18 0800'
            'Al B. Adams|FT5001|True|Jones, Bob|2026-09-22 0913'
            'Al B. Adams|FT5102|False||'
            'Bea Brown|BA101|False||'
            'Bea Brown|FI3101|True||2026-09-18 0815'
            'Bea Brown|FT5001|True||2026-09-17 0830'
            'Bea Brown|FT5102|False||'
        )
    }

    It 'the Big Board''s First Last names find the FAIS students written Last, First or by last name' {
        $students = @((New-TestStudent -Name 'Adams, Al B.'), (New-TestStudent -Name 'Brown'))
        $plan = Get-SyllabusUpdatePlan -File $script:B -Students $students -SoftWeights $script:SW -Now $script:Now
        @($plan.Matched) | Should -Be @('Adams, Al B.', 'Brown')
        @($plan.Changes | Where-Object { $_.Kind -eq 'Complete' } | ForEach-Object { "$($_.Student) $($_.Mission)" } | Sort-Object) |
            Should -Be @('Adams, Al B. FI3101', 'Adams, Al B. FT5001', 'Brown FI3101', 'Brown FT5001')
    }

    It 'says so for a workbook that is neither, and for a file that is not a workbook' {
        $other = New-TBook 'OutputSchedule.xlsx' @{ MasterSchedule = @(@('ETD', 'Instructor', 'SortieType', 'Student', 'Mission', 'Callsign', 'Num', 'Notes'), @('2026-09-21 07:30:00', 'Jones, Bob', 'Day Local', 'Adams, Al', 'FT5001', '', '', '')) }
        $r = Read-TrainingFile -Path $other
        $r.Kind | Should -BeExactly ''
        $r.Error | Should -Match 'is not a flying schedule'
        $bad = Join-Path $TestDrive 'NotAWorkbook.xlsx'
        [System.IO.File]::WriteAllText($bad, 'nope')
        $r2 = Read-TrainingFile -Path $bad 3>$null
        $r2.Error | Should -Match '^Could not open NotAWorkbook.xlsx'
        (Read-TrainingFile -Path (Join-Path $TestDrive 'missing.xlsx')).Error | Should -Match 'was not found'
    }
}


Describe 'Saving an update' -Skip:$script:NoExcel {
    BeforeEach {
        $script:Out = Join-Path $TestDrive 'OutputSchedule.xlsx'
        $script:MD = Join-Path $TestDrive 'MasterData.xlsx'
        foreach ($p in @($script:Out, $script:MD)) { if (Test-Path $p) { Remove-Item $p -Force } }

        # Adams has a tab (with a Notes column of the user's own); Brown has none yet
        $script:Adams = New-TestStudent -Name 'Adams, Al' -Done @('FI3101|2026-09-18 08:00', 'FT5001|2026-09-21 07:30|Lee, Ann') -SimDebt 2
        $script:Brown = New-TestStudent -Name 'Brown, Bea' -Done @('FI3101|2026-09-18 08:00') -SimDebt 1
        $rows = @(foreach ($m in $script:Adams.Syllabus) { $r = ConvertTo-SyllabusTabRow $m; $r | Add-Member -NotePropertyName Notes -NotePropertyValue "note $($m.Mission)"; $r })
        $rows | Export-Excel -Path $script:Out -WorksheetName 'Adams, Al' 3>$null
        @([PSCustomObject]@{ SPKey = 1; SPName = 'Adams, Al'; SimDebt = 2 }, [PSCustomObject]@{ SPKey = 2; SPName = 'Brown, Bea'; SimDebt = 1 }) |
            Export-Excel -Path $script:MD -WorksheetName 'Students' 3>$null

        $file = New-TFile -Facts @(
            New-TFact 'Adams, Al [2d Lt]' 'FT5001' 'CX' 'Lee, Ann [Capt]' '2026-09-21 07:30'
            New-TFact 'Adams, Al [2d Lt]' 'FE3102' 'CM' 'PARR, PAT [CIV]' '2026-09-21 13:00'
            New-TFact 'Brown, Bea [2d Lt]' 'FT5001' 'CM' 'Jones, Bob [Maj]' '2026-09-21 09:00'
        )
        $script:Plan = Get-SyllabusUpdatePlan -File $file -Students @($script:Adams, $script:Brown) -Instructors (New-TestInstructors) -OutputPath $script:Out -SoftWeights $script:SW -Now $script:Now
    }

    It 'works the changes out against the student''s tab, which is what gets saved' {
        @($script:Plan.Changes | Sort-Object Student, Mission | ForEach-Object { "$($_.Student) $($_.Mission) $($_.Kind)" }) | Should -Be @(
            'Adams, Al FE3102 Complete', 'Adams, Al FT5001 NotComplete', 'Brown, Bea FT5001 Complete')
    }

    It 'writes only the changed events'' cells, makes a tab for a student without one, and moves SimDebt in MasterData' {
        $memory = @(Invoke-SyllabusUpdate -Targets $script:Plan.Targets -SoftWeights $script:SW -Now $script:Now)
        $saved = Save-SyllabusUpdate -Targets $script:Plan.Targets -MemoryChanges $memory -OutputPath $script:Out -MasterDataPath $script:MD -SoftWeights $script:SW -Now $script:Now 6>$null
        $saved.Message | Should -BeExactly ''
        $saved.Saved | Should -BeTrue

        $tab = @(Import-Excel -Path $script:Out -WorksheetName 'Adams, Al')
        $ft = $tab | Where-Object { $_.Mission -eq 'FT5001' }
        [string]$ft.IsComplete | Should -BeExactly 'False'
        [string]$ft.Instructor + [string]$ft.ETD | Should -BeExactly ''
        $fe = $tab | Where-Object { $_.Mission -eq 'FE3102' }
        [string]$fe.IsComplete | Should -BeExactly 'True'
        $fe.Instructor | Should -BeExactly 'PARR, PAT'
        $fe.ETD | Should -BeExactly '2026-09-21 13:00:00'
        @($tab | ForEach-Object { $_.Notes }) -join ',' | Should -BeExactly (@($script:Adams.Syllabus | ForEach-Object { "note $($_.Mission)" }) -join ',') -Because 'nothing else on the tab changes'
        ($tab | Where-Object { $_.Mission -eq 'FI3101' }).ETD | Should -BeExactly '2026-09-18 08:00:00'

        $brownTab = @(Import-Excel -Path $script:Out -WorksheetName 'Brown, Bea')
        ($brownTab | Where-Object { $_.Mission -eq 'FT5001' }).Instructor | Should -BeExactly 'Jones, Bob'

        # Adams: FT5001 taken back (-1), FE3102 flown (-0.5); Brown: FT5001 flown (+1)
        $md = @(Import-Excel -Path $script:MD -WorksheetName 'Students')
        ($md | Where-Object { $_.SPName -eq 'Adams, Al' }).SimDebt | Should -Be 0.5
        ($md | Where-Object { $_.SPName -eq 'Brown, Bea' }).SimDebt | Should -Be 2
        $script:Adams.SimDebt | Should -Be 0.5
        $script:Brown.SimDebt | Should -Be 2
    }

    It 'saves nothing when a file stays open, and says which' {
        $memory = @(Invoke-SyllabusUpdate -Targets $script:Plan.Targets -SoftWeights $script:SW -Now $script:Now)
        $before = [System.IO.File]::ReadAllBytes($script:Out)
        $lock = [System.IO.File]::Open($script:MD, 'Open', 'ReadWrite', 'None')
        try {
            $saved = Save-SyllabusUpdate -Targets $script:Plan.Targets -MemoryChanges $memory -OutputPath $script:Out -MasterDataPath $script:MD -SoftWeights $script:SW -LockedPrompt { $false } -Now $script:Now 6>$null
        }
        finally { $lock.Close() }
        $saved.Saved | Should -BeFalse
        $saved.Message | Should -BeExactly 'MasterData.xlsx is still open'
        [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($script:Out)) | Should -BeExactly ([System.Convert]::ToBase64String($before))
    }

    It 'puts both files back when a write fails part-way' {
        Mock -ModuleName SyllabusImport Update-MasterDataStudents { throw 'disk full' }
        $before = [System.IO.File]::ReadAllBytes($script:Out)
        $memory = @(Invoke-SyllabusUpdate -Targets $script:Plan.Targets -SoftWeights $script:SW -Now $script:Now)
        $saved = Save-SyllabusUpdate -Targets $script:Plan.Targets -MemoryChanges $memory -OutputPath $script:Out -MasterDataPath $script:MD -SoftWeights $script:SW -Now $script:Now 6>$null
        $saved.Saved | Should -BeFalse
        $saved.Message | Should -BeExactly 'disk full'
        [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($script:Out)) | Should -BeExactly ([System.Convert]::ToBase64String($before))
    }

    It 'without the files the update is kept for this run only' {
        $saved = Save-SyllabusUpdate -Targets $script:Plan.Targets -OutputPath '' -MasterDataPath '' -SoftWeights $script:SW
        $saved.Saved | Should -BeFalse
        $saved.Message | Should -Match 'was not found'
    }
}
