# Pester 5 tests for the fill engine: Invoke-FillSchedule, Process-HardcodedSortie,
# Pair-FormationWingman and Schedule-ReturnLeg (Modules\SchedulingEngine.psm1).
# Day 1 is Tue 2026-10-20, day 2 is Wed 2026-10-21.

BeforeAll {
                . "$PSScriptRoot/Fixtures/New-FaisTestData.ps1"
                Import-FaisModules

                # Everything flown except the formation block (only FF5201 / FF5202 remain)
                $script:FormDone = @(
                                'FI3101|2026-10-10 08:00'
                                'FE3102|2026-10-11 08:00'
                                'FT5001|2026-10-12 08:00|Ortiz, Sam'
                                'FT5102|2026-10-13 08:00|Ortiz, Sam'
                                'FT5103|2026-10-14 08:00|Ortiz, Sam'
                                'FN3201|2026-10-15 08:00'
                                'FT5104|2026-10-16 08:00|Ortiz, Sam'
                                'FT5190|2026-10-17 08:00|Leech, Tom'
                )
                # Ready for the checkride: FF5201 (ID 7) and FT5190 (ID 10, IsCheck) remain on a Day Local
                $script:CheckDone = @(
                                'FI3101|2026-10-10 08:00'
                                'FE3102|2026-10-11 08:00'
                                'FT5001|2026-10-12 08:00|Ortiz, Sam'
                                'FT5102|2026-10-13 08:00|Ortiz, Sam'
                                'FT5103|2026-10-14 08:00|Ortiz, Sam'
                                'FN3201|2026-10-15 08:00'
                                'FT5104|2026-10-16 08:00|Ortiz, Sam'
                )

                function Invoke-TFill {
                                param (
                                                [array]$Schedule,
                                                [array]$Students,
                                                $Formations,
                                                [scriptblock]$OnProgress
                                )
                                if ($null -eq $Formations) { $Formations = New-TestFormations }
                                $splat = @{
                                                Schedule    = $Schedule
                                                Students    = $Students
                                                Instructors = (New-TestInstructors)
                                                SoftWeights = (New-TestSoftWeights)
                                                Formations  = $Formations
                                }
                                if ($null -ne $OnProgress) { $splat['OnProgress'] = $OnProgress }
                                $result = Invoke-FillSchedule @splat 3>$null 6>$null
                                return $result
                }

                function Get-TRow {
                                param ($Result, [string]$Id)
                                foreach ($r in $Result.Schedule) { if ([string]$r.BoardLineId -eq $Id) { return $r } }
                                return $null
                }

                function Get-TNotes {
                                param ($Row)
                                if ([string]::IsNullOrWhiteSpace([string]$Row.Notes)) { return @() }
                                return @(([string]$Row.Notes) -split '\s*\|\s*' | Where-Object { $_ -ne '' })
                }

                function Get-TEntry {
                                param ($Student, [string]$Mission)
                                return (Find-SyllabusMission -Student $Student -Mission $Mission)
                }

                function Get-TStudent {
                                param ($Result, [string]$Name)
                                foreach ($s in $Result.Students) { if ([string]$s.Name -eq $Name) { return $s } }
                                return $null
                }

                # Completed events on one day, as 'HHmm Mission'
                function Get-TDayEvents {
                                param ($Student, [string]$Day)
                                return @($Student.Syllabus | Where-Object { $_.IsComplete -and ([string]$_.ETD).StartsWith($Day) } |
                                                Sort-Object -Property ETD | ForEach-Object { '{0} {1}' -f ([datetime]$_.ETD).ToString('HHmm'), $_.Mission })
                }

                # Everything a fill decides, as text, so two runs can be compared
                function Get-TSnapshot {
                                param ($Result)
                                $lines = foreach ($r in $Result.Schedule) {
                                                '{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}' -f $r.BoardLineId, $r.ETD, (Get-NameText $r.Instructor), (Get-NameText $r.Student), (Get-NameText $r.Mission), $r.FilledBy, $r.Callsign, $r.Num, $r.Notes
                                }
                                $studs = foreach ($s in ($Result.Students | Sort-Object -Property Name)) {
                                                '{0}|{1}|{2}' -f $s.Name, $s.SimDebt, ((@($s.Syllabus | Where-Object { $_.IsComplete } | ForEach-Object { "$($_.Mission)@$($_.ETD)/$($_.Instructor)" })) -join ',')
                                }
                                $ovr = foreach ($o in $Result.Overrides) { 'OVR {0}|{1}|{2}' -f $o.LineId, $o.Student, (@($o.Violations) -join ';') }
                                return ((@($lines) + @($studs) + @($ovr) + @(($Result.Stats | ConvertTo-Json -Compress))) -join "`n")
                }

                # Mixed day: pin, engine O&B + return, formation pair (twice), For Sale, unknown pin
                function New-TMixedStudents {
                                return @(
                                                New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00')
                                                New-TestStudent -Name 'Cole, Cy' -Done @('FI3101|2026-10-14 08:00')
                                                New-TestStudent -Name 'Adams, Al' -Done $script:FormDone
                                                New-TestStudent -Name 'Brown, Bea' -Done $script:FormDone
                                )
                }
                function New-TMixedRows {
                                return @(
                                                New-TestRow -ETD '2026-10-20 07:00' -Instructor 'Lee, Ann' -Id 'P1' -Student 'Able, Amy' -Mission 'FT5001'
                                                New-TestRow -ETD '2026-10-20 06:00' -SortieType 'O&B' -Instructor 'Ortiz, Sam' -Id 'OB1' -LinkedLineId 'RT1'
                                                New-TestRow -ETD '2026-10-20 09:00' -SortieType 'O&B Return' -Instructor 'Ortiz, Sam' -Id 'RT1'
                                                New-TestRow -ETD '2026-10-20 08:00' -Instructor 'Jones, Bob' -Id 'A1'
                                                New-TestRow -ETD '2026-10-20 08:10' -Instructor 'Leech, Tom' -Id 'A2'
                                                New-TestRow -ETD '2026-10-20 16:30' -SortieType 'OFT' -Id 'S1'
                                                New-TestRow -ETD '2026-10-20 15:00' -Instructor 'Jones, Bob' -Id 'U1' -Student 'Nobody, Ned'
                                                New-TestRow -ETD '2026-10-20 13:00' -Instructor 'Jones, Bob' -Id 'C1'
                                                New-TestRow -ETD '2026-10-20 13:10' -Instructor 'Lee, Ann' -Id 'C2'
                                )
                }
}

Describe 'Invoke-FillSchedule' {

                Context 'pinned lines' {

                                It 'keeps a pinned student and marks the mission complete with the IP name and ETD' {
                                                $students = @(New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00'))
                                                $rows = @(New-TestRow -ETD '2026-10-20 07:30' -Instructor 'Jones, Bob' -Id 'P1' -Student 'Able, Amy' -Mission 'FT5001')
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                $row = Get-TRow $r 'P1'
                                                [object]::ReferenceEquals($row.Student, $students[0]) | Should -BeTrue
                                                Get-NameText $row.Mission | Should -Be 'FT5001'
                                                Get-NameText $row.Instructor | Should -Be 'Jones, Bob'
                                                $row.FilledBy | Should -Be 'Pin'
                                                $row.Callsign | Should -Be 'TALON'
                                                $row.Num | Should -Be 31
                                                $row.Notes | Should -BeNullOrEmpty

                                                $e = Get-TEntry $students[0] 'FT5001'
                                                $e.IsComplete | Should -BeTrue
                                                $e.Instructor | Should -BeExactly 'Jones, Bob'
                                                $e.ETD | Should -Be '2026-10-20 07:30:00'
                                                @($r.Overrides) | Should -HaveCount 0
                                }

                                It 'resolves a pin typed as a unique last name' {
                                                $students = @(
                                                                New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00')
                                                                New-TestStudent -Name 'Baker, Bo' -Done @('FI3101|2026-10-15 08:00')
                                                )
                                                $rows = @(New-TestRow -ETD '2026-10-20 07:30' -Instructor 'Jones, Bob' -Id 'P1' -Student 'Able' -Mission 'FT5001')
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                Get-NameText (Get-TRow $r 'P1').Student | Should -Be 'Able, Amy'
                                                (Get-TEntry $students[0] 'FT5001').IsComplete | Should -BeTrue
                                                (Get-TEntry $students[1] 'FT5001').IsComplete | Should -BeFalse
                                }

                                It "gives an '(auto)' pin the best-scoring passing event, not the lowest syllabus ID" {
                                                $students = @(New-TestStudent -Name 'Chk, Carl' -Done $script:CheckDone)
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 08:00' -Instructor 'Jones, Bob' -Id 'P1' -Student 'Chk, Carl'
                                                                New-TestRow -ETD '2026-10-20 08:10' -Instructor 'Lee, Ann' -Id 'L1'
                                                )

                                                # Guard: both FF5201 (ID 7) and FT5190 (ID 10) pass on this line
                                                $probeStudent = Copy-StudentForEval $students[0]
                                                $hr = Test-HardRules -Student $probeStudent -Sortie $rows[0] -Schedule $rows -Students @($probeStudent) -Instructors (New-TestInstructors) -Explain
                                                $passed = @($hr.PassedEvents | ForEach-Object { $_.Mission })
                                                $passed | Should -Contain 'FF5201'
                                                $passed | Should -Contain 'FT5190'

                                                $r = Invoke-TFill -Schedule $rows -Students $students
                                                $row = Get-TRow $r 'P1'
                                                Get-NameText $row.Mission | Should -Be 'FT5190'
                                                $row.FilledBy | Should -Be 'Pin'
                                                @(Get-TNotes $row) | Should -Contain 'Checkride! Requires Formal'
                                                $row.Notes | Should -Not -BeLike '*OVERRIDE*'
                                                (Get-TEntry $students[0] 'FT5190').Instructor | Should -Be 'Jones, Bob'
                                                (Get-TEntry $students[0] 'FF5201').IsComplete | Should -BeFalse
                                                @($r.Overrides) | Should -HaveCount 0
                                }

                                It "gives an '(auto)' pin the same event the engine picks on an open line" {
                                                $pinStudents = @(New-TestStudent -Name 'Chk, Carl' -Done $script:CheckDone)
                                                $pinRows = @(
                                                                New-TestRow -ETD '2026-10-20 08:00' -Instructor 'Jones, Bob' -Id 'P1' -Student 'Chk, Carl'
                                                                New-TestRow -ETD '2026-10-20 08:10' -Instructor 'Lee, Ann' -Id 'L1'
                                                )
                                                $openStudents = @(New-TestStudent -Name 'Chk, Carl' -Done $script:CheckDone)
                                                $openRows = @(
                                                                New-TestRow -ETD '2026-10-20 08:00' -Instructor 'Jones, Bob' -Id 'P1'
                                                                New-TestRow -ETD '2026-10-20 08:10' -Instructor 'Lee, Ann' -Id 'L1'
                                                )
                                                $pin = Invoke-TFill -Schedule $pinRows -Students $pinStudents
                                                $open = Invoke-TFill -Schedule $openRows -Students $openStudents

                                                (Get-TRow $open 'P1').FilledBy | Should -Be 'Engine'
                                                Get-NameText (Get-TRow $open 'P1').Student | Should -Be 'Chk, Carl'
                                                Get-NameText (Get-TRow $pin 'P1').Mission | Should -Be (Get-NameText (Get-TRow $open 'P1').Mission)
                                }

                                It "treats an empty-string pin mission like '(auto)'" {
                                                $students = @(New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00'))
                                                $rows = @(New-TestRow -ETD '2026-10-20 07:30' -Instructor 'Lee, Ann' -Id 'P1' -Student 'Able, Amy' -Mission '')
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                Get-NameText (Get-TRow $r 'P1').Mission | Should -Be 'FT5001'
                                                (Get-TEntry $students[0] 'FT5001').IsComplete | Should -BeTrue
                                                @($r.Overrides) | Should -HaveCount 0
                                }

                                It "skips events that fail event-level rules for an '(auto)' pin (checkride with a non-Check IP)" {
                                                $students = @(New-TestStudent -Name 'Chk, Carl' -Done $script:CheckDone)
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 08:00' -Instructor 'Lee, Ann' -Id 'P1' -Student 'Chk, Carl'
                                                                New-TestRow -ETD '2026-10-20 08:10' -Instructor 'Jones, Bob' -Id 'L1'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                $p1 = Get-TRow $r 'P1'
                                                Get-NameText $p1.Mission | Should -Be 'FF5201'
                                                $p1.Callsign | Should -Be 'HAZE'
                                                $p1.Num | Should -Be 11
                                                @(Get-TNotes $p1) | Should -Contain 'Wingman: DS'
                                                @(Get-TNotes $p1) | Should -Contain 'Direct Support Required'
                                                $p1.Notes | Should -Not -BeLike '*OVERRIDE*'

                                                $l1 = Get-TRow $r 'L1'
                                                $l1.Student | Should -Be 'Need Form SP'
                                                $l1.FilledBy | Should -Be 'DS'
                                                @($r.Overrides) | Should -HaveCount 0
                                }

                                It "falls back to the first event and flags it when no event passes for an '(auto)' pin" {
                                                $students = @(New-TestStudent -Name 'Chk, Carl' -Done $script:CheckDone)
                                                $rows = @(New-TestRow -ETD '2026-10-20 08:00' -Instructor 'Lee, Ann' -Id 'P1' -Student 'Chk, Carl')
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                $p1 = Get-TRow $r 'P1'
                                                Get-NameText $p1.Mission | Should -Be 'FF5201'
                                                @(Get-TNotes $p1) | Should -Contain 'Wingman: none found'
                                                @(Get-TNotes $p1) | Should -Contain 'OVERRIDE: No open Day Local line within 22 min for a wingman/DS'
                                                @($r.Overrides) | Should -HaveCount 1
                                                @($r.Overrides[0].Violations) | Should -Be @('No open Day Local line within 22 min for a wingman/DS')
                                }

                                It 'keeps an explicit pinned mission that is not eligible, with an OVERRIDE note and an Overrides entry' {
                                                $students = @(New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00'))
                                                $rows = @(New-TestRow -ETD '2026-10-20 07:00' -Instructor 'Lee, Ann' -Id 'P1' -Student 'Able, Amy' -Mission 'FT5103')
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                $p1 = Get-TRow $r 'P1'
                                                Get-NameText $p1.Mission | Should -Be 'FT5103'
                                                $p1.FilledBy | Should -Be 'Pin'
                                                @(Get-TNotes $p1) | Should -Contain "OVERRIDE: FT5103 isn't eligible here (prereq FT5102 not complete)"
                                                @(Get-TNotes $p1) | Should -Contain 'Solo Flight'
                                                (Get-TNotes $p1)[-1] | Should -BeLike 'OVERRIDE: *'

                                                $e = Get-TEntry $students[0] 'FT5103'
                                                $e.IsComplete | Should -BeTrue
                                                $e.Instructor | Should -Be 'Lee, Ann'

                                                @($r.Overrides) | Should -HaveCount 1
                                                $o = $r.Overrides[0]
                                                $o.LineId | Should -Be 'P1'
                                                $o.ETD | Should -Be '2026-10-20 07:00:00'
                                                $o.Student | Should -Be 'Able, Amy'
                                                @($o.Violations) | Should -Be @("FT5103 isn't eligible here (prereq FT5102 not complete)")
                                }

                                It 'flags a pinned checkride with a non-Check IP' {
                                                $students = @(New-TestStudent -Name 'Chk, Carl' -Done $script:CheckDone)
                                                $rows = @(New-TestRow -ETD '2026-10-20 08:00' -Instructor 'Lee, Ann' -Id 'P1' -Student 'Chk, Carl' -Mission 'FT5190')
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                $p1 = Get-TRow $r 'P1'
                                                Get-NameText $p1.Mission | Should -Be 'FT5190'
                                                $p1.Notes | Should -Be "OVERRIDE: Checkride needs a Check IP (Lee, Ann isn't)"
                                                (Get-TEntry $students[0] 'FT5190').IsComplete | Should -BeTrue
                                }

                                It 'flags a pinned event on the wrong line type' {
                                                $students = @(New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00'))
                                                $rows = @(New-TestRow -ETD '2026-10-20 07:00' -Instructor 'Lee, Ann' -Id 'P1' -Student 'Able, Amy' -Mission 'FE3102')
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                (Get-TRow $r 'P1').Notes | Should -Be "OVERRIDE: FE3102 isn't eligible here (Type OFT not allowed on Day Local)"
                                                (Get-TEntry $students[0] 'FE3102').IsComplete | Should -BeTrue
                                }

                                It 'joins several hard-rule breaks into one OVERRIDE note, on both pins of a short turn' {
                                                $students = @(New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00') -OnHold)
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 07:00' -Instructor 'Lee, Ann' -Id 'P1' -Student 'Able, Amy'
                                                                New-TestRow -ETD '2026-10-20 09:00' -Instructor 'Jones, Bob' -Id 'P2' -Student 'Able, Amy'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                # P1 sees P2 as a reservation (no mission yet); P2 sees P1's completed FT5001
                                                (Get-TRow $r 'P1').Notes | Should -Be 'OVERRIDE: Student is on hold; Turn time 2.0 hr < 3.25 hr minimum (vs 0900)'
                                                (Get-TRow $r 'P2').Notes | Should -Be 'OVERRIDE: Student is on hold; Turn time 2.0 hr < 3.25 hr minimum (vs 0700 FT5001)'
                                                Get-NameText (Get-TRow $r 'P1').Mission | Should -Be 'FT5001'
                                                Get-NameText (Get-TRow $r 'P2').Mission | Should -Be 'FT5102'

                                                @($r.Overrides) | Should -HaveCount 2
                                                @($r.Overrides | ForEach-Object { $_.LineId }) | Should -Be @('P1', 'P2')
                                                @($r.Overrides[1].Violations) | Should -Be @('Student is on hold', 'Turn time 2.0 hr < 3.25 hr minimum (vs 0700 FT5001)')
                                                $r.Stats.Overrides | Should -Be 2
                                }

                                It 'flags an unknown pinned student, keeps the line, and logs it' {
                                                $students = @(New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00'))
                                                $rows = @(New-TestRow -ETD '2026-10-20 12:00' -Instructor 'Jones, Bob' -Id 'P3' -Student 'Nobody, Ned' -Mission 'FT5001')
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                $p3 = Get-TRow $r 'P3'
                                                $p3.Student | Should -Be 'Nobody, Ned'
                                                $p3.FilledBy | Should -Be 'Pin'
                                                $p3.Notes | Should -Be "OVERRIDE: student 'Nobody, Ned' not found in MasterData"
                                                (Get-TEntry $students[0] 'FT5001').IsComplete | Should -BeFalse

                                                @($r.Overrides) | Should -HaveCount 1
                                                $r.Overrides[0].Student | Should -Be 'Nobody, Ned'
                                                $r.Overrides[0].LineId | Should -Be 'P3'
                                                @($r.Overrides[0].Violations) | Should -Be @("student 'Nobody, Ned' not found in MasterData")
                                                @($r.Log) | Should -Contain "Pinned student 'Nobody, Ned' on 2026-10-20 12:00:00 not found in MasterData"
                                }

                                It 'flags a pinned mission that is not in the syllabus and keeps the typed mission text' {
                                                $students = @(New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00'))
                                                $rows = @(New-TestRow -ETD '2026-10-20 07:00' -Instructor 'Lee, Ann' -Id 'P1' -Student 'Able, Amy' -Mission 'FX9999')
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                $p1 = Get-TRow $r 'P1'
                                                $p1.Mission | Should -Be 'FX9999'
                                                $p1.Notes | Should -BeLike "OVERRIDE: FX9999 isn't in Able, Amy's syllabus*"
                                                @($r.Overrides[0].Violations) | Should -Contain "FX9999 isn't in Able, Amy's syllabus"
                                                $students[0].SimDebt | Should -Be 0
                                }

                                It 'flags the same mission pinned on two lines and applies its SimDebt once' {
                                                $students = @(New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00'))
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 07:00' -Instructor 'Lee, Ann' -Id 'P1' -Student 'Able, Amy' -Mission 'FT5001'
                                                                New-TestRow -ETD '2026-10-20 11:00' -Instructor 'Jones, Bob' -Id 'P2' -Student 'Able, Amy' -Mission 'FT5001'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                (Get-TRow $r 'P1').Notes | Should -Be "OVERRIDE: FT5001 isn't eligible here (already pinned on another line)"
                                                (Get-TRow $r 'P2').Notes | Should -Be "OVERRIDE: FT5001 isn't eligible here (already complete)"
                                                $students[0].SimDebt | Should -Be 1
                                }
                }

                Context 'engine picks and reservations' {

                                It 'marks engine picks complete so no student gets a turn under 3.25 hr (old $SelectedMission bug)' {
                                                $students = @(
                                                                New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00')
                                                                New-TestStudent -Name 'Baker, Bo' -Done @('FI3101|2026-10-14 08:00', 'FT5001|2026-10-15 08:00|Jones, Bob')
                                                )
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 07:30' -Instructor 'Jones, Bob' -Id 'L1'
                                                                New-TestRow -ETD '2026-10-20 08:30' -Instructor 'Lee, Ann' -Id 'L2'
                                                                New-TestRow -ETD '2026-10-20 09:30' -Instructor 'Ortiz, Sam' -Id 'L3'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                (Get-TRow $r 'L1').FilledBy | Should -Be 'Engine'
                                                (Get-TRow $r 'L2').FilledBy | Should -Be 'Engine'
                                                (Get-TRow $r 'L3').Student | Should -Be 'For Sale'
                                                @(Get-TDayEvents $students[0] '2026-10-20') | Should -HaveCount 1
                                                @(Get-TDayEvents $students[1] '2026-10-20') | Should -HaveCount 1

                                                Get-NameText (Get-TRow $r 'L1').Student | Should -Be 'Able, Amy'
                                                $e = Get-TEntry $students[0] 'FT5001'
                                                $e.IsComplete | Should -BeTrue
                                                $e.Instructor | Should -Be 'Jones, Bob'
                                                $e.ETD | Should -Be '2026-10-20 07:30:00'

                                                # The completed pick is what the hard rules now see
                                                $probe = New-TestRow -ETD '2026-10-20 09:30' -Instructor 'Ortiz, Sam' -Id 'PX'
                                                $hr = Test-HardRules -Student $students[0] -Sortie $probe -Instructors (New-TestInstructors) -Explain
                                                @($hr.Violations | ForEach-Object { $_.Message }) | Should -Contain 'Turn time 2.0 hr < 3.25 hr minimum (vs 0730 FT5001)'
                                }

                                It 'does not triple-turn a student the engine already scheduled twice' {
                                                $students = @(New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00'))
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 07:00' -Instructor 'Jones, Bob' -Id 'L1'
                                                                New-TestRow -ETD '2026-10-20 10:30' -SortieType 'OFT' -Id 'L2'
                                                                New-TestRow -ETD '2026-10-20 14:00' -Instructor 'Lee, Ann' -Id 'L3'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                Get-NameText (Get-TRow $r 'L1').Mission | Should -Be 'FT5001'
                                                Get-NameText (Get-TRow $r 'L2').Mission | Should -Be 'FE3102'
                                                (Get-TRow $r 'L3').Student | Should -Be 'For Sale'
                                                @(Get-TDayEvents $students[0] '2026-10-20') | Should -Be @('0700 FT5001', '1030 FE3102')

                                                $probe = New-TestRow -ETD '2026-10-20 14:00' -Instructor 'Lee, Ann' -Id 'PX'
                                                $hr = Test-HardRules -Student $students[0] -Sortie $probe -Instructors (New-TestInstructors) -Explain
                                                @($hr.Violations | ForEach-Object { $_.Message }) | Should -Contain 'Already has 2 events today (0700 FT5001, 1030 FE3102) - triple turn'
                                }

                                It 'lets a LATER pinned line block an earlier fill that would break turn time' {
                                                # Control: without the pin the student gets the 0800 line
                                                $ctlStudents = @(New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00'))
                                                $ctl = Invoke-TFill -Schedule @(New-TestRow -ETD '2026-10-20 08:00' -Instructor 'Lee, Ann' -Id 'L1') -Students $ctlStudents
                                                Get-NameText (Get-TRow $ctl 'L1').Student | Should -Be 'Able, Amy'

                                                $students = @(New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00'))
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 08:00' -Instructor 'Lee, Ann' -Id 'L1'
                                                                New-TestRow -ETD '2026-10-20 10:00' -SortieType 'OFT' -Id 'P1' -Student 'Able, Amy' -Mission 'FE3102'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                $l1 = Get-TRow $r 'L1'
                                                $l1.Student | Should -Be 'For Sale'
                                                $l1.FilledBy | Should -Be 'ForSale'
                                                (Get-TEntry $students[0] 'FT5001').IsComplete | Should -BeFalse
                                                (Get-TRow $r 'P1').Notes | Should -BeNullOrEmpty
                                                @(Get-TDayEvents $students[0] '2026-10-20') | Should -Be @('1000 FE3102')
                                }

                                It 'does not let a later pin block an earlier fill with a legal turn' {
                                                $students = @(New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00'))
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 08:00' -Instructor 'Lee, Ann' -Id 'L1'
                                                                New-TestRow -ETD '2026-10-20 11:30' -SortieType 'OFT' -Id 'P1' -Student 'Able, Amy' -Mission 'FE3102'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                Get-NameText (Get-TRow $r 'L1').Student | Should -Be 'Able, Amy'
                                                Get-NameText (Get-TRow $r 'L1').Mission | Should -Be 'FT5001'
                                                (Get-TRow $r 'P1').Notes | Should -BeNullOrEmpty
                                                @($r.Overrides) | Should -HaveCount 0
                                }

                                It "holds the time of a pinned O&B's return leg too" {
                                                $students = @(New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00'))
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 06:00' -Instructor 'Lee, Ann' -Id 'L1'
                                                                New-TestRow -ETD '2026-10-20 10:00' -SortieType 'O&B' -Instructor 'Jones, Bob' -Id 'OB1' -LinkedLineId 'RT1' -Student 'Able, Amy'
                                                                New-TestRow -ETD '2026-10-20 13:00' -SortieType 'O&B Return' -Instructor 'Jones, Bob' -Id 'RT1'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                (Get-TRow $r 'L1').Student | Should -Be 'For Sale'
                                                @(Get-TDayEvents $students[0] '2026-10-20') | Should -Be @('1000 FT5001', '1300 FT5102')
                                                (Get-TRow $r 'OB1').Notes | Should -BeNullOrEmpty
                                }

                                It 'never schedules a pinned mission on another line' {
                                                $students = @(New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00'))
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 07:00' -Instructor 'Lee, Ann' -Id 'L1'
                                                                New-TestRow -ETD '2026-10-20 14:00' -Instructor 'Jones, Bob' -Id 'P1' -Student 'Able, Amy' -Mission 'FT5001'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                (Get-TRow $r 'L1').Student | Should -Be 'For Sale'
                                                $e = Get-TEntry $students[0] 'FT5001'
                                                $e.ETD | Should -Be '2026-10-20 14:00:00'
                                                $e.Instructor | Should -Be 'Jones, Bob'
                                                (Get-TRow $r 'P1').Notes | Should -BeNullOrEmpty
                                                $students[0].SimDebt | Should -Be 1
                                }

                                It "keeps prerequisite order: an earlier engine pick can satisfy a later pin's prereq" {
                                                # Control: the FT5102 pin alone is missing its prereq
                                                $ctlStudents = @(New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00'))
                                                $ctl = Invoke-TFill -Schedule @(New-TestRow -ETD '2026-10-20 14:00' -Instructor 'Jones, Bob' -Id 'P1' -Student 'Able, Amy' -Mission 'FT5102') -Students $ctlStudents
                                                (Get-TRow $ctl 'P1').Notes | Should -Be "OVERRIDE: FT5102 isn't eligible here (prereq FT5001 not complete)"

                                                $students = @(New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00'))
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 08:00' -Instructor 'Lee, Ann' -Id 'L1'
                                                                New-TestRow -ETD '2026-10-20 14:00' -Instructor 'Jones, Bob' -Id 'P1' -Student 'Able, Amy' -Mission 'FT5102'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                Get-NameText (Get-TRow $r 'L1').Mission | Should -Be 'FT5001'
                                                Get-NameText (Get-TRow $r 'P1').Mission | Should -Be 'FT5102'
                                                (Get-TRow $r 'P1').Notes | Should -Not -BeLike '*OVERRIDE*'
                                                @($r.Overrides) | Should -HaveCount 0
                                }

                                It 'breaks score ties by name, whatever order the students come in' {
                                                $a = @(New-TestStudent -Name 'Zulu, Zed' -Done @('FI3101|2026-10-15 08:00'); New-TestStudent -Name 'Alpha, Al' -Done @('FI3101|2026-10-15 08:00'))
                                                $b = @(New-TestStudent -Name 'Alpha, Al' -Done @('FI3101|2026-10-15 08:00'); New-TestStudent -Name 'Zulu, Zed' -Done @('FI3101|2026-10-15 08:00'))
                                                $ra = Invoke-TFill -Schedule @(New-TestRow -ETD '2026-10-20 07:00' -Instructor 'Lee, Ann' -Id 'L1') -Students $a
                                                $rb = Invoke-TFill -Schedule @(New-TestRow -ETD '2026-10-20 07:00' -Instructor 'Lee, Ann' -Id 'L1') -Students $b

                                                Get-NameText (Get-TRow $ra 'L1').Student | Should -Be 'Alpha, Al'
                                                Get-NameText (Get-TRow $rb 'L1').Student | Should -Be 'Alpha, Al'
                                }

                                It 'refills lines that still hold an engine placeholder' {
                                                $students = @(
                                                                New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00')
                                                                New-TestStudent -Name 'Baker, Bo' -Done @('FI3101|2026-10-15 08:00')
                                                )
                                                $l1 = New-TestRow -ETD '2026-10-20 07:00' -Instructor 'Lee, Ann' -Id 'L1' -Student 'For Sale'
                                                $l1.FilledBy = 'ForSale'
                                                $l2 = New-TestRow -ETD '2026-10-20 07:10' -Instructor 'Jones, Bob' -Id 'L2' -Student 'Need Form SP'
                                                $l2.FilledBy = 'DS'
                                                $r = Invoke-TFill -Schedule @($l1, $l2) -Students $students

                                                (Get-TRow $r 'L1').FilledBy | Should -Be 'Engine'
                                                (Get-TRow $r 'L2').FilledBy | Should -Be 'Engine'
                                                @((Get-TRow $r 'L1'), (Get-TRow $r 'L2') | ForEach-Object { Get-NameText $_.Student } | Sort-Object) | Should -Be @('Able, Amy', 'Baker, Bo')
                                                $r.Stats.Pinned | Should -Be 0
                                }
                }

                Context 'For Sale lines' {

                                It "marks a Day Local nobody can fly For Sale with 'No Available SP | Potentially fill if we have another line'" {
                                                $students = @(New-TestStudent -Name 'Hold, Hal' -Done @('FI3101|2026-10-15 08:00') -OnHold)
                                                $r = Invoke-TFill -Schedule @(New-TestRow -ETD '2026-10-20 08:00' -Instructor 'Jones, Bob' -Id 'L1') -Students $students

                                                $l1 = Get-TRow $r 'L1'
                                                $l1.Student | Should -BeExactly 'For Sale'
                                                $l1.Mission | Should -BeNullOrEmpty
                                                $l1.FilledBy | Should -Be 'ForSale'
                                                $l1.Notes | Should -BeExactly 'No Available SP | Potentially fill if we have another line'
                                                $r.Stats.ForSale | Should -Be 1
                                }

                                It "gives sim lines (OFT/IFT) only 'No Available SP'" {
                                                $students = @(New-TestStudent -Name 'Hold, Hal' -Done @('FI3101|2026-10-15 08:00') -OnHold)
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 08:00' -SortieType 'OFT' -Id 'S1'
                                                                New-TestRow -ETD '2026-10-20 09:00' -SortieType 'IFT' -Id 'S2'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                (Get-TRow $r 'S1').Notes | Should -BeExactly 'No Available SP'
                                                (Get-TRow $r 'S2').Notes | Should -BeExactly 'No Available SP'
                                                (Get-TRow $r 'S1').Student | Should -Be 'For Sale'
                                                (Get-TRow $r 'S2').Student | Should -Be 'For Sale'
                                }

                                It 'appends the For Sale note to notes the line already had' {
                                                $students = @(New-TestStudent -Name 'Hold, Hal' -Done @('FI3101|2026-10-15 08:00') -OnHold)
                                                $r = Invoke-TFill -Schedule @(New-TestRow -ETD '2026-10-20 13:00' -SortieType 'OFT' -Id 'S1' -Notes 'Bring charts') -Students $students

                                                (Get-TRow $r 'S1').Notes | Should -BeExactly 'Bring charts | No Available SP'
                                }

                                It "marks an open O&B Return with no O&B For Sale with 'O&B Leg 1 Not Filled'" {
                                                $students = @(New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00'))
                                                $r = Invoke-TFill -Schedule @(New-TestRow -ETD '2026-10-20 11:00' -SortieType 'O&B Return' -Instructor 'Jones, Bob' -Id 'RT1') -Students $students

                                                $rt = Get-TRow $r 'RT1'
                                                $rt.Student | Should -BeExactly 'For Sale'
                                                $rt.Mission | Should -BeNullOrEmpty
                                                $rt.FilledBy | Should -Be 'ForSale'
                                                $rt.Notes | Should -BeExactly 'O&B Leg 1 Not Filled'
                                                (Get-TEntry $students[0] 'FT5001').IsComplete | Should -BeFalse
                                }

                                It 'puts the return up for sale too when nobody can fly the O&B' {
                                                $students = @(New-TestStudent -Name 'Hold, Hal' -Done @('FI3101|2026-10-15 08:00') -OnHold)
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 08:00' -SortieType 'O&B' -Instructor 'Jones, Bob' -Id 'OB1' -LinkedLineId 'RT1'
                                                                New-TestRow -ETD '2026-10-20 11:00' -SortieType 'O&B Return' -Instructor 'Jones, Bob' -Id 'RT1'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                (Get-TRow $r 'OB1').Notes | Should -BeExactly 'No Available SP | Potentially fill if we have another line'
                                                (Get-TRow $r 'RT1').Notes | Should -BeExactly 'O&B Leg 1 Not Filled'
                                                (Get-TRow $r 'RT1').Student | Should -Be 'For Sale'
                                                $r.Stats.ForSale | Should -Be 2
                                }
                }

                Context 'O&B return legs' {

                                It 'fills the linked return (LinkedLineId) with the next event and the same IP when the engine picks an O&B' {
                                                $students = @(New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00'))
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 08:00' -SortieType 'O&B' -Instructor 'Jones, Bob' -Id 'OB1' -LinkedLineId 'RT1'
                                                                New-TestRow -ETD '2026-10-20 11:00' -SortieType 'O&B Return' -Instructor 'Lee, Ann' -Id 'RT1'
                                                                # Unlinked decoy that the legacy same-IP lookup would match
                                                                New-TestRow -ETD '2026-10-20 10:30' -SortieType 'O&B Return' -Instructor 'Jones, Bob' -Id 'RT2'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                $ob = Get-TRow $r 'OB1'
                                                $ob.FilledBy | Should -Be 'Engine'
                                                Get-NameText $ob.Mission | Should -Be 'FT5001'

                                                $rt = Get-TRow $r 'RT1'
                                                [object]::ReferenceEquals($rt.Student, $students[0]) | Should -BeTrue
                                                Get-NameText $rt.Mission | Should -Be 'FT5102'
                                                Get-NameText $rt.Instructor | Should -Be 'Jones, Bob'
                                                $rt.Callsign | Should -Be $ob.Callsign
                                                $rt.Num | Should -Be $ob.Num
                                                $rt.FilledBy | Should -Be 'Return'

                                                $e = Get-TEntry $students[0] 'FT5102'
                                                $e.IsComplete | Should -BeTrue
                                                $e.Instructor | Should -BeOfType [string]
                                                $e.Instructor | Should -BeExactly 'Jones, Bob'
                                                $e.ETD | Should -Be '2026-10-20 11:00:00'

                                                (Get-TRow $r 'RT2').Notes | Should -BeExactly 'O&B Leg 1 Not Filled'
                                                $r.Stats.Return | Should -Be 1
                                                $students[0].SimDebt | Should -Be 2
                                }

                                It 'fills the linked return of a pinned O&B' {
                                                $students = @(New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00') -SimDebt 2)
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 08:00' -SortieType 'O&B' -Instructor 'Jones, Bob' -Id 'OB1' -LinkedLineId 'RT1' -Student 'Able, Amy'
                                                                New-TestRow -ETD '2026-10-20 11:00' -SortieType 'O&B Return' -Instructor 'Jones, Bob' -Id 'RT1'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                (Get-TRow $r 'OB1').FilledBy | Should -Be 'Pin'
                                                Get-NameText (Get-TRow $r 'OB1').Mission | Should -Be 'FT5001'
                                                (Get-TRow $r 'RT1').FilledBy | Should -Be 'Return'
                                                Get-NameText (Get-TRow $r 'RT1').Mission | Should -Be 'FT5102'
                                                @($r.Overrides) | Should -HaveCount 0
                                                $students[0].SimDebt | Should -Be 4
                                                $students[0].Reservations | Should -BeNullOrEmpty
                                }

                                It 'finds the return by the legacy same-day same-IP lookup for IngestData rows with no board ids' {
                                                $students = @(New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00'))
                                                $ob = [PSCustomObject]@{ ETD = '2026-10-20 08:00:00'; Instructor = 'Jones, Bob'; SortieType = 'O&B'; Student = $null; Mission = $null; Callsign = $null; Num = $null; Notes = $null }
                                                $other = [PSCustomObject]@{ ETD = '2026-10-20 11:00:00'; Instructor = 'Lee, Ann'; SortieType = 'O&B Return'; Student = $null; Mission = $null; Callsign = $null; Num = $null; Notes = $null }
                                                $mine = [PSCustomObject]@{ ETD = '2026-10-20 11:00:00'; Instructor = 'Jones, Bob'; SortieType = 'O&B Return'; Student = $null; Mission = $null; Callsign = $null; Num = $null; Notes = $null }
                                                $r = Invoke-TFill -Schedule @($ob, $other, $mine) -Students $students

                                                @($r.Schedule | ForEach-Object { $_.BoardLineId }) | Should -Be @('row0', 'row1', 'row2')
                                                $mine.FilledBy | Should -Be 'Return'
                                                Get-NameText $mine.Mission | Should -Be 'FT5102'
                                                $other.Student | Should -Be 'For Sale'
                                                $other.Notes | Should -BeExactly 'O&B Leg 1 Not Filled'
                                }
                }

                Context 'formations' {

                                It 'gives an engine formation lead a student wingman with FormCallsigns and notes' {
                                                $students = @(
                                                                New-TestStudent -Name 'Adams, Al' -Done $script:FormDone
                                                                New-TestStudent -Name 'Brown, Bea' -Done $script:FormDone
                                                )
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 08:00' -Instructor 'Jones, Bob' -Id 'F1'
                                                                New-TestRow -ETD '2026-10-20 08:10' -Instructor 'Lee, Ann' -Id 'F2'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                $lead = Get-TRow $r 'F1'
                                                Get-NameText $lead.Student | Should -Be 'Adams, Al'
                                                Get-NameText $lead.Mission | Should -Be 'FF5201'
                                                $lead.FilledBy | Should -Be 'Engine'
                                                $lead.Callsign | Should -Be 'HAZE'
                                                $lead.Num | Should -Be 11
                                                $lead.Notes | Should -BeExactly 'Wingman: Brown, Bea'

                                                $wing = Get-TRow $r 'F2'
                                                Get-NameText $wing.Student | Should -Be 'Brown, Bea'
                                                Get-NameText $wing.Mission | Should -Be 'FF5201'
                                                $wing.FilledBy | Should -Be 'Wing'
                                                $wing.Callsign | Should -Be 'HAZE'
                                                $wing.Num | Should -Be 12
                                                $wing.Notes | Should -BeExactly 'Wingman to Adams, Al'

                                                $e = Get-TEntry $students[1] 'FF5201'
                                                $e.IsComplete | Should -BeTrue
                                                $e.Instructor | Should -Be 'Lee, Ann'
                                                $e.ETD | Should -Be '2026-10-20 08:10:00'
                                                $students[1].SimDebt | Should -Be 1
                                                $r.Stats.Filled | Should -Be 1
                                                $r.Stats.Wing | Should -Be 1
                                }

                                It 'gives a pinned formation lead a student wingman' {
                                                $students = @(
                                                                New-TestStudent -Name 'Adams, Al' -Done $script:FormDone
                                                                New-TestStudent -Name 'Brown, Bea' -Done $script:FormDone
                                                )
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 08:00' -Instructor 'Jones, Bob' -Id 'F1' -Student 'Adams, Al' -Mission 'FF5201'
                                                                New-TestRow -ETD '2026-10-20 08:10' -Instructor 'Lee, Ann' -Id 'F2'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                (Get-TRow $r 'F1').FilledBy | Should -Be 'Pin'
                                                (Get-TRow $r 'F1').Notes | Should -BeExactly 'Wingman: Brown, Bea'
                                                (Get-TRow $r 'F1').Num | Should -Be 11
                                                (Get-TRow $r 'F2').FilledBy | Should -Be 'Wing'
                                                (Get-TRow $r 'F2').Num | Should -Be 12
                                                (Get-TEntry $students[1] 'FF5201').IsComplete | Should -BeTrue
                                                @($r.Overrides) | Should -HaveCount 0
                                }

                                It "falls back to Direct Support ('Need Form SP') when no wingman is available, skipping sim lines" {
                                                $students = @(New-TestStudent -Name 'Adams, Al' -Done $script:FormDone)
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 08:00' -Instructor 'Jones, Bob' -Id 'F1'
                                                                New-TestRow -ETD '2026-10-20 08:05' -SortieType 'IFT' -Id 'S1'
                                                                New-TestRow -ETD '2026-10-20 08:07' -SortieType 'OFT' -Id 'S2'
                                                                New-TestRow -ETD '2026-10-20 08:15' -Instructor 'Lee, Ann' -Id 'F2'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                $lead = Get-TRow $r 'F1'
                                                Get-NameText $lead.Mission | Should -Be 'FF5201'
                                                $lead.Notes | Should -BeExactly 'Direct Support Required | Wingman: DS'
                                                $lead.Num | Should -Be 11

                                                $ds = Get-TRow $r 'F2'
                                                $ds.Student | Should -BeExactly 'Need Form SP'
                                                $ds.Mission | Should -BeNullOrEmpty
                                                $ds.FilledBy | Should -Be 'DS'
                                                $ds.Notes | Should -BeExactly 'Direct Support for Adams, Al'
                                                $ds.Callsign | Should -Be 'HAZE'
                                                $ds.Num | Should -Be 12

                                                (Get-TRow $r 'S1').FilledBy | Should -Be 'ForSale'
                                                (Get-TRow $r 'S2').FilledBy | Should -Be 'ForSale'
                                                $r.Stats.DS | Should -Be 1
                                }

                                It 'pairs two pinned lines through WingOfLineId without taking another open line' {
                                                $students = @(
                                                                New-TestStudent -Name 'Adams, Al' -Done $script:FormDone
                                                                New-TestStudent -Name 'Brown, Bea' -Done $script:FormDone
                                                )
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 08:00' -Instructor 'Jones, Bob' -Id 'F1' -Student 'Adams, Al' -Mission 'FF5201'
                                                                New-TestRow -ETD '2026-10-20 08:10' -Instructor 'Lee, Ann' -Id 'F2' -Student 'Brown, Bea' -Mission 'FF5201' -WingOfLineId 'F1'
                                                                New-TestRow -ETD '2026-10-20 08:05' -Instructor 'Ortiz, Sam' -Id 'X'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                $lead = Get-TRow $r 'F1'
                                                $lead.FilledBy | Should -Be 'Pin'
                                                $lead.Callsign | Should -Be 'HAZE'
                                                $lead.Num | Should -Be 11
                                                @(Get-TNotes $lead) | Should -Contain 'Wingman: Brown, Bea'

                                                $wing = Get-TRow $r 'F2'
                                                $wing.FilledBy | Should -Be 'Pin'
                                                $wing.Callsign | Should -Be 'HAZE'
                                                $wing.Num | Should -Be 12
                                                $wing.Notes | Should -BeExactly 'Wingman to Adams, Al'

                                                $x = Get-TRow $r 'X'
                                                $x.Student | Should -Be 'For Sale'
                                                $x.Notes | Should -Not -BeLike '*Direct Support for*'

                                                $e = Get-TEntry $students[1] 'FF5201'
                                                $e.Instructor | Should -Be 'Lee, Ann'
                                                $e.ETD | Should -Be '2026-10-20 08:10:00'
                                                $students[1].SimDebt | Should -Be 1
                                                $r.Stats.Wing | Should -Be 0
                                }

                                It "gives an '(auto)' pin on a wing line a formation event even when a single-ship event scores higher" {
                                                $students = @(
                                                                New-TestStudent -Name 'Adams, Al' -Done $script:FormDone
                                                                New-TestStudent -Name 'Chk, Carl' -Done $script:CheckDone
                                                )
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 08:00' -Instructor 'Lee, Ann' -Id 'F1' -Student 'Adams, Al' -Mission 'FF5201'
                                                                New-TestRow -ETD '2026-10-20 08:10' -Instructor 'Jones, Bob' -Id 'F2' -Student 'Chk, Carl' -WingOfLineId 'F1'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                # On Jones' line FT5190 would win on score (+50 Checkride Priority)
                                                Get-NameText (Get-TRow $r 'F2').Mission | Should -Be 'FF5201'
                                                (Get-TRow $r 'F2').Num | Should -Be 12
                                                (Get-TEntry $students[1] 'FT5190').IsComplete | Should -BeFalse
                                }

                                It "does not pick a wingman whose later pin would break turn time" {
                                                $students = @(
                                                                New-TestStudent -Name 'Adams, Al' -Done $script:FormDone
                                                                New-TestStudent -Name 'Brown, Bea' -Done $script:FormDone
                                                )
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 08:00' -Instructor 'Jones, Bob' -Id 'F1' -Student 'Adams, Al' -Mission 'FF5201'
                                                                New-TestRow -ETD '2026-10-20 08:10' -Instructor 'Lee, Ann' -Id 'F2'
                                                                New-TestRow -ETD '2026-10-20 09:30' -SortieType 'OFT' -Id 'P2' -Student 'Brown, Bea'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                (Get-TRow $r 'F2').FilledBy | Should -Be 'DS'
                                                (Get-TRow $r 'F2').Student | Should -Be 'Need Form SP'
                                                (Get-TRow $r 'F1').Notes | Should -BeExactly 'Wingman: DS'
                                                (Get-TEntry $students[1] 'FF5201').IsComplete | Should -BeFalse
                                }

                                It 'does not flag an OVERRIDE on a lead whose wing is pinned through WingOfLineId' {
                                                # Regression: the lead's open-wing-line check must count the pinned wing.
                                                $students = @(
                                                                New-TestStudent -Name 'Adams, Al' -Done $script:FormDone
                                                                New-TestStudent -Name 'Brown, Bea' -Done $script:FormDone
                                                )
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 08:00' -Instructor 'Jones, Bob' -Id 'F1' -Student 'Adams, Al' -Mission 'FF5201'
                                                                New-TestRow -ETD '2026-10-20 08:10' -Instructor 'Lee, Ann' -Id 'F2' -Student 'Brown, Bea' -Mission 'FF5201' -WingOfLineId 'F1'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                (Get-TRow $r 'F1').Notes | Should -BeExactly 'Wingman: Brown, Bea'
                                                @($r.Overrides) | Should -HaveCount 0
                                }

                                It "does not note 'Direct Support Required' on a lead whose wing is pinned through WingOfLineId" {
                                                # Regression: the wing's reserved FF5201 must not make the lead ask for Direct Support.
                                                $students = @(
                                                                New-TestStudent -Name 'Adams, Al' -Done $script:FormDone
                                                                New-TestStudent -Name 'Brown, Bea' -Done $script:FormDone
                                                )
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 08:00' -Instructor 'Jones, Bob' -Id 'F1' -Student 'Adams, Al' -Mission 'FF5201'
                                                                New-TestRow -ETD '2026-10-20 08:10' -Instructor 'Lee, Ann' -Id 'F2' -Student 'Brown, Bea' -Mission 'FF5201' -WingOfLineId 'F1'
                                                                New-TestRow -ETD '2026-10-20 08:05' -Instructor 'Ortiz, Sam' -Id 'X'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                @(Get-TNotes (Get-TRow $r 'F1')) | Should -Not -Contain 'Direct Support Required'
                                }

                                It 'never uses an IFT / OFT line as a wingman, even when the board paired it' {
                                                $students = @(
                                                                New-TestStudent -Name 'Adams, Al' -Done $script:FormDone
                                                                New-TestStudent -Name 'Brown, Bea' -Done $script:FormDone
                                                )
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 08:00' -Instructor 'Jones, Bob' -Id 'F1' -Student 'Adams, Al' -Mission 'FF5201'
                                                                New-TestRow -ETD '2026-10-20 08:05' -Instructor 'Ortiz, Sam' -Id 'X'
                                                                New-TestRow -ETD '2026-10-20 08:10' -SortieType 'IFT' -Id 'S1' -Student 'Brown, Bea' -WingOfLineId 'F1'
                                                                New-TestRow -ETD '2026-10-20 08:15' -SortieType 'OFT' -Id 'S2'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                (Get-TRow $r 'S1').Notes | Should -Not -BeLike '*Wingman to*'
                                                (Get-TRow $r 'S1').Callsign | Should -Not -Be 'HAZE'
                                                (Get-TRow $r 'S2').FilledBy | Should -Not -BeIn @('Wing', 'DS')
                                                (Get-TRow $r 'F1').Notes | Should -BeLike '*Wingman: DS*'
                                                (Get-TRow $r 'X').FilledBy | Should -Be 'DS'
                                }

                                It 'advances FormIndex within a day' {
                                                $r = Invoke-TFill -Schedule (New-TMixedRows) -Students (New-TMixedStudents)

                                                (Get-TRow $r 'A1').Num | Should -Be 11
                                                (Get-TRow $r 'A2').Num | Should -Be 12
                                                (Get-TRow $r 'C1').Callsign | Should -Be 'HAZE'
                                                (Get-TRow $r 'C1').Num | Should -Be 21
                                                (Get-TRow $r 'C2').Num | Should -Be 22
                                                (Get-TRow $r 'C1').FilledBy | Should -Be 'Engine'
                                                (Get-TRow $r 'C2').FilledBy | Should -Be 'Wing'
                                }

                                It 'resets FormIndex each day' {
                                                $students = @(
                                                                New-TestStudent -Name 'Adams, Al' -Done $script:FormDone
                                                                New-TestStudent -Name 'Brown, Bea' -Done $script:FormDone
                                                )
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 08:00' -Instructor 'Jones, Bob' -Id 'A1'
                                                                New-TestRow -ETD '2026-10-20 08:10' -Instructor 'Lee, Ann' -Id 'A2'
                                                                New-TestRow -ETD '2026-10-21 08:00' -Instructor 'Jones, Bob' -Id 'B1'
                                                                New-TestRow -ETD '2026-10-21 08:10' -Instructor 'Lee, Ann' -Id 'B2'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                Get-NameText (Get-TRow $r 'B1').Mission | Should -Be 'FF5202'
                                                (Get-TRow $r 'B2').FilledBy | Should -Be 'Wing'
                                                (Get-TRow $r 'A1').Num | Should -Be 11
                                                (Get-TRow $r 'A2').Num | Should -Be 12
                                                (Get-TRow $r 'B1').Num | Should -Be 11
                                                (Get-TRow $r 'B2').Num | Should -Be 12
                                }
                }

                Context 'multi-day crew rest' {

                                It 'blocks an early day-2 line after an engine pick late on day 1 (< 15 hr)' {
                                                $students = @(New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00'))
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-21 07:00' -Instructor 'Lee, Ann' -Id 'D2'
                                                                New-TestRow -ETD '2026-10-20 18:00' -Instructor 'Jones, Bob' -Id 'D1'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                @($r.Schedule | ForEach-Object { $_.BoardLineId }) | Should -Be @('D1', 'D2')
                                                Get-NameText (Get-TRow $r 'D1').Mission | Should -Be 'FT5001'
                                                (Get-TRow $r 'D2').Student | Should -Be 'For Sale'
                                                (Get-TEntry $students[0] 'FT5102').IsComplete | Should -BeFalse
                                }

                                It 'lets a pinned early day-2 line block a late day-1 fill (crew rest before the next day)' {
                                                $students = @(New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00'))
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 18:00' -Instructor 'Jones, Bob' -Id 'D1'
                                                                New-TestRow -ETD '2026-10-21 07:00' -SortieType 'OFT' -Id 'D2' -Student 'Able, Amy' -Mission 'FE3102'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                (Get-TRow $r 'D1').Student | Should -Be 'For Sale'
                                                (Get-TRow $r 'D2').Notes | Should -BeNullOrEmpty
                                                (Get-TEntry $students[0] 'FT5001').IsComplete | Should -BeFalse
                                }

                                It 'keeps pins on both days and flags crew rest in both directions' {
                                                $students = @(New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00'))
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 18:00' -Instructor 'Jones, Bob' -Id 'D1' -Student 'Able, Amy' -Mission 'FT5001'
                                                                New-TestRow -ETD '2026-10-21 07:00' -SortieType 'OFT' -Id 'D2' -Student 'Able, Amy' -Mission 'FE3102'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                $next = ([datetime]'2026-10-21 07:00').ToString('ddd HHmm')
                                                $prev = ([datetime]'2026-10-20 18:00').ToString('ddd HHmm')
                                                (Get-TRow $r 'D1').Notes | Should -BeExactly "OVERRIDE: Crew rest 13.0 hr < 15 hr before $next FE3102"
                                                (Get-TRow $r 'D2').Notes | Should -BeExactly "OVERRIDE: Crew rest 13.0 hr < 15 hr since $prev FT5001"
                                                @($r.Overrides) | Should -HaveCount 2
                                                (Get-TEntry $students[0] 'FT5001').IsComplete | Should -BeTrue
                                                (Get-TEntry $students[0] 'FE3102').IsComplete | Should -BeTrue
                                }

                                It 'allows day 2 when crew rest is exactly 15 hr' {
                                                $students = @(New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00'))
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 16:00' -Instructor 'Jones, Bob' -Id 'D1'
                                                                New-TestRow -ETD '2026-10-21 07:00' -Instructor 'Lee, Ann' -Id 'D2'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                Get-NameText (Get-TRow $r 'D1').Mission | Should -Be 'FT5001'
                                                Get-NameText (Get-TRow $r 'D2').Student | Should -Be 'Able, Amy'
                                                Get-NameText (Get-TRow $r 'D2').Mission | Should -Be 'FT5102'
                                }
                }

                Context 'SimDebt' {

                                It 'applies SimDebt for pins, engine picks, wingmen and return legs' {
                                                $students = @(
                                                                New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00') -SimDebt 1
                                                                New-TestStudent -Name 'Dunn, Di' -Done @('FI3101|2026-10-14 08:00')
                                                                New-TestStudent -Name 'Adams, Al' -Done $script:FormDone -SimDebt 0.5
                                                                New-TestStudent -Name 'Brown, Bea' -Done $script:FormDone -SimDebt 2
                                                )
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 06:00' -SortieType 'OFT' -Id 'P1' -Student 'Able, Amy' -Mission 'FE3102'
                                                                New-TestRow -ETD '2026-10-20 05:00' -SortieType 'O&B' -Instructor 'Ortiz, Sam' -Id 'OB1' -LinkedLineId 'RT1'
                                                                New-TestRow -ETD '2026-10-20 08:00' -SortieType 'O&B Return' -Instructor 'Ortiz, Sam' -Id 'RT1'
                                                                New-TestRow -ETD '2026-10-20 08:00' -Instructor 'Jones, Bob' -Id 'F1'
                                                                New-TestRow -ETD '2026-10-20 08:10' -Instructor 'Lee, Ann' -Id 'F2'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                (Get-TRow $r 'P1').FilledBy | Should -Be 'Pin'
                                                Get-NameText (Get-TRow $r 'OB1').Student | Should -Be 'Dunn, Di'
                                                (Get-TRow $r 'RT1').FilledBy | Should -Be 'Return'
                                                Get-NameText (Get-TRow $r 'F1').Student | Should -Be 'Adams, Al'
                                                Get-NameText (Get-TRow $r 'F2').Student | Should -Be 'Brown, Bea'

                                                (Get-TStudent $r 'Able, Amy').SimDebt | Should -Be 0.5     # pinned transition sim: -0.5
                                                (Get-TStudent $r 'Dunn, Di').SimDebt | Should -Be 2        # engine O&B + return: +1 +1
                                                (Get-TStudent $r 'Adams, Al').SimDebt | Should -Be 1.5     # engine lead: +1
                                                (Get-TStudent $r 'Brown, Bea').SimDebt | Should -Be 3      # wing: +1
                                }
                }

                Context 'unsat clean-up' {

                                It 'an OFT re-sim earlier in the day lets an UnsatSim student fly a later line' {
                                                $students = @(New-TestStudent -Name 'Able, Amy' -UnsatSim -UnsatSimType 'OFT' -Done @('FI3101|2026-10-15 08:00'))
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 07:30' -SortieType 'OFT' -Id 'S1'
                                                                New-TestRow -ETD '2026-10-20 12:00' -Instructor 'Jones, Bob' -Id 'F1'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                Get-NameText (Get-TRow $r 'S1').Mission | Should -Be 'FE3102'
                                                Get-TNotes (Get-TRow $r 'S1') | Should -Contain 'Clears Unsat Sim (OFT)'
                                                Get-NameText (Get-TRow $r 'F1').Student | Should -Be 'Able, Amy'
                                                (Get-TStudent $r 'Able, Amy').UnsatSim | Should -BeFalse
                                }

                                It 'an IFT re-sim cleans up an IFT hook but not an OFT hook' -TestCases @(
                                                @{ Kind = 'IFT'; Flies = 'Able, Amy'; Unsat = $false }
                                                @{ Kind = 'OFT'; Flies = 'For Sale'; Unsat = $true }
                                ) {
                                                $done = @('FI3101|2026-10-10 08:00', 'FE3102|2026-10-11 08:00', 'FT5001|2026-10-12 08:00|Ortiz, Sam', 'FT5102|2026-10-13 08:00|Ortiz, Sam')
                                                $students = @(New-TestStudent -Name 'Able, Amy' -UnsatSim -UnsatSimType $Kind -Done $done)
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 07:30' -SortieType 'IFT' -Id 'S1'
                                                                New-TestRow -ETD '2026-10-20 12:00' -Instructor 'Jones, Bob' -Id 'F1'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                Get-NameText (Get-TRow $r 'S1').Student | Should -Be 'Able, Amy'
                                                Get-NameText (Get-TRow $r 'F1').Student | Should -Be $Flies
                                                (Get-TStudent $r 'Able, Amy').UnsatSim | Should -Be $Unsat
                                }

                                It 'a flight cleans up an unsat flight so a later sim can be filled' {
                                                $students = @(New-TestStudent -Name 'Able, Amy' -UnsatFlight -Done @('FI3101|2026-10-15 08:00'))
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 07:30' -Instructor 'Jones, Bob' -Id 'F1'
                                                                New-TestRow -ETD '2026-10-20 12:00' -SortieType 'OFT' -Id 'S1'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                Get-NameText (Get-TRow $r 'F1').Mission | Should -Be 'FT5001'
                                                Get-TNotes (Get-TRow $r 'F1') | Should -Contain 'Clears Unsat Flight'
                                                Get-NameText (Get-TRow $r 'S1').Student | Should -Be 'Able, Amy'
                                                (Get-TStudent $r 'Able, Amy').UnsatFlight | Should -BeFalse
                                }

                                It 'a re-sim later in the day does not clear an earlier hardcoded flight' {
                                                $students = @(New-TestStudent -Name 'Able, Amy' -UnsatSim -UnsatSimType 'OFT' -Done @('FI3101|2026-10-15 08:00'))
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 07:30' -Instructor 'Jones, Bob' -Id 'P1' -Student 'Able, Amy' -Mission 'FT5001'
                                                                New-TestRow -ETD '2026-10-20 12:00' -SortieType 'OFT' -Id 'P2' -Student 'Able, Amy' -Mission 'FE3102'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                @(Get-TNotes (Get-TRow $r 'P1') | Where-Object { $_ -like 'OVERRIDE:*Unsat sim*' }).Count | Should -Be 1
                                                Get-TNotes (Get-TRow $r 'P2') | Should -Contain 'Clears Unsat Sim (OFT)'
                                                (Get-TStudent $r 'Able, Amy').UnsatSim | Should -BeFalse
                                }

                                It 'never clears the flag on the students passed in when they are copies' {
                                                $base = New-TestStudent -Name 'Able, Amy' -UnsatSim -UnsatSimType 'OFT' -Done @('FI3101|2026-10-15 08:00')
                                                $rows = @(New-TestRow -ETD '2026-10-20 07:30' -SortieType 'OFT' -Id 'S1')
                                                $r = Invoke-TFill -Schedule $rows -Students @(Copy-StudentList @($base))
                                                (Get-TStudent $r 'Able, Amy').UnsatSim | Should -BeFalse
                                                $base.UnsatSim | Should -BeTrue
                                }
                }

                Context 'determinism, stats and results' {

                                It 'gives identical results for two runs on identical student copies' {
                                                $base = New-TMixedStudents
                                                $r1 = Invoke-TFill -Schedule (New-TMixedRows) -Students (Copy-StudentList $base)
                                                $r2 = Invoke-TFill -Schedule (New-TMixedRows) -Students (Copy-StudentList $base)

                                                $s1 = Get-TSnapshot $r1
                                                $s1 | Should -Not -BeNullOrEmpty
                                                $s1 | Should -BeLike '*|Engine|*'
                                                Get-TSnapshot $r2 | Should -BeExactly $s1
                                }

                                It 'gives the same schedule whatever order the students come in' {
                                                $base = New-TMixedStudents
                                                $rev = @($base[3], $base[2], $base[1], $base[0])
                                                $r1 = Invoke-TFill -Schedule (New-TMixedRows) -Students (Copy-StudentList $base)
                                                $r2 = Invoke-TFill -Schedule (New-TMixedRows) -Students (Copy-StudentList $rev)

                                                Get-TSnapshot $r2 | Should -BeExactly (Get-TSnapshot $r1)
                                }

                                It 'leaves the original students untouched when filling copies' {
                                                $base = New-TMixedStudents
                                                $before = $base | ConvertTo-Json -Depth 6
                                                $r = Invoke-TFill -Schedule (New-TMixedRows) -Students (Copy-StudentList $base)

                                                $r.Stats.Filled | Should -BeGreaterThan 0
                                                ($base | ConvertTo-Json -Depth 6) | Should -BeExactly $before
                                }

                                It 'counts each kind of line in Stats' {
                                                $r = Invoke-TFill -Schedule (New-TMixedRows) -Students (New-TMixedStudents)

                                                (Get-TRow $r 'P1').FilledBy | Should -Be 'Pin'
                                                (Get-TRow $r 'OB1').FilledBy | Should -Be 'Engine'
                                                Get-NameText (Get-TRow $r 'OB1').Student | Should -Be 'Cole, Cy'
                                                (Get-TRow $r 'RT1').FilledBy | Should -Be 'Return'
                                                (Get-TRow $r 'A1').FilledBy | Should -Be 'Engine'
                                                (Get-TRow $r 'A2').FilledBy | Should -Be 'Wing'
                                                (Get-TRow $r 'S1').FilledBy | Should -Be 'ForSale'
                                                (Get-TRow $r 'U1').FilledBy | Should -Be 'Pin'

                                                $r.Stats.Lines | Should -Be 9
                                                $r.Stats.Pinned | Should -Be 2
                                                $r.Stats.Filled | Should -Be 3
                                                $r.Stats.Wing | Should -Be 2
                                                $r.Stats.Return | Should -Be 1
                                                $r.Stats.DS | Should -Be 0
                                                $r.Stats.ForSale | Should -Be 1
                                                $r.Stats.Overrides | Should -Be 1
                                                @($r.Overrides) | Should -HaveCount 1
                                }

                                It 'returns rows in ETD order and gives rows without a board id one' {
                                                $students = @(New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00'))
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 08:00' -SortieType 'OFT'
                                                                New-TestRow -ETD '2026-10-20 07:00' -Instructor 'Lee, Ann'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students

                                                @($r.Schedule | ForEach-Object { $_.BoardLineId }) | Should -Be @('row1', 'row0')
                                                @($r.Schedule | ForEach-Object { $_.ETD }) | Should -Be @('2026-10-20 07:00:00', '2026-10-20 08:00:00')
                                                [object]::ReferenceEquals($r.Students[0], $students[0]) | Should -BeTrue
                                }

                                It 'reports progress for each line it processes and survives a failing callback' {
                                                $calls = New-Object System.Collections.Generic.List[string]
                                                $students = @(New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00'))
                                                $rows = @(
                                                                New-TestRow -ETD '2026-10-20 08:00' -SortieType 'OFT'
                                                                New-TestRow -ETD '2026-10-20 07:00' -Instructor 'Lee, Ann'
                                                )
                                                $r = Invoke-TFill -Schedule $rows -Students $students -OnProgress { param($i, $n, $label) $calls.Add("$i/$n $label") }
                                                $calls.ToArray() | Should -Be @('1/2 0700 Day Local', '2/2 0800 OFT')

                                                $students2 = @(New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00'))
                                                $r2 = Invoke-TFill -Schedule @(New-TestRow -ETD '2026-10-20 07:00' -Instructor 'Lee, Ann' -Id 'L1') -Students $students2 -OnProgress { throw 'boom' }
                                                (Get-TRow $r2 'L1').FilledBy | Should -Be 'Engine'
                                }
                }
}

Describe 'Process-HardcodedSortie' {

                BeforeAll {
                                $script:Ips = New-TestInstructors
                                $script:W = New-TestSoftWeights
                                $script:F = New-TestFormations
                }

                It 'records the pin, completes the mission with SimDebt and keeps a preset callsign' {
                                $s = New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00') -SimDebt 1
                                $row = New-TestRow -ETD '2026-10-20 07:30' -Instructor 'Jones, Bob' -Id 'P1'
                                $row.Callsign = 'VIPER'
                                $row.Num = 7
                                $m = Get-TEntry $s 'FT5001'
                                $out = @(Process-HardcodedSortie -Sortie $row -Student $s -Mission $m -Schedule @($row) -Students @($s) -Instructors $script:Ips -SoftWeights $script:W -Formations $script:F -FormIndex 2 6>$null)

                                $out | Should -Be @(2)
                                $row.Instructor.Name | Should -Be 'Jones, Bob'
                                [object]::ReferenceEquals($row.Student, $s) | Should -BeTrue
                                [object]::ReferenceEquals($row.Mission, $m) | Should -BeTrue
                                $row.FilledBy | Should -Be 'Pin'
                                $row.Callsign | Should -Be 'VIPER'
                                $row.Num | Should -Be 7
                                $m.IsComplete | Should -BeTrue
                                $m.Instructor | Should -BeExactly 'Jones, Bob'
                                $m.ETD | Should -Be '2026-10-20 07:30:00'
                                $s.SimDebt | Should -Be 2
                }

                It 'takes the callsign from the IP when the line has none' {
                                $s = New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00')
                                $row = New-TestRow -ETD '2026-10-20 07:30' -Instructor 'Ortiz, Sam' -Id 'P1'
                                [void](Process-HardcodedSortie -Sortie $row -Student $s -Mission (Get-TEntry $s 'FT5001') -Schedule @($row) -Students @($s) -Instructors $script:Ips -SoftWeights $script:W 6>$null)

                                $row.Callsign | Should -Be 'TALON'
                                $row.Num | Should -Be 34
                }

                It 'pairs a formation lead, skipping sim lines, and returns the advanced FormIndex' {
                                $a = New-TestStudent -Name 'Adams, Al' -Done $script:FormDone
                                $b = New-TestStudent -Name 'Brown, Bea' -Done $script:FormDone
                                $lead = New-TestRow -ETD '2026-10-20 08:00' -Instructor 'Jones, Bob' -Id 'F1'
                                $ift = New-TestRow -ETD '2026-10-20 08:05' -SortieType 'IFT' -Id 'S1'
                                $oft = New-TestRow -ETD '2026-10-20 08:07' -SortieType 'OFT' -Id 'S2'
                                $wing = New-TestRow -ETD '2026-10-20 08:22' -Instructor 'Lee, Ann' -Id 'F2'
                                $out = @(Process-HardcodedSortie -Sortie $lead -Student $a -Mission (Get-TEntry $a 'FF5201') -Schedule @($lead, $ift, $oft, $wing) -Students @($a, $b) `
                                                -Instructors $script:Ips -SoftWeights $script:W -Formations $script:F -FormIndex 2 6>$null)

                                $out | Should -Be @(0)
                                $lead.Num | Should -Be 21
                                $lead.Notes | Should -BeExactly 'Wingman: Brown, Bea'
                                $wing.Num | Should -Be 22
                                $wing.FilledBy | Should -Be 'Wing'
                                $ift.Student | Should -BeNullOrEmpty
                                $oft.Student | Should -BeNullOrEmpty
                                $ift.FilledBy | Should -BeNullOrEmpty
                }

                It 'does not pair a line that is already a wingman' {
                                $a = New-TestStudent -Name 'Adams, Al' -Done $script:FormDone
                                $b = New-TestStudent -Name 'Brown, Bea' -Done $script:FormDone
                                $row = New-TestRow -ETD '2026-10-20 08:00' -Instructor 'Jones, Bob' -Id 'F1' -Notes 'Wingman to Carter, Cy'
                                $open = New-TestRow -ETD '2026-10-20 08:10' -Instructor 'Lee, Ann' -Id 'F2'
                                $out = @(Process-HardcodedSortie -Sortie $row -Student $a -Mission (Get-TEntry $a 'FF5201') -Schedule @($row, $open) -Students @($a, $b) `
                                                -Instructors $script:Ips -SoftWeights $script:W -Formations $script:F -FormIndex 0 6>$null)

                                $out | Should -Be @(0)
                                $row.Notes | Should -BeExactly 'Wingman to Carter, Cy'
                                $open.Student | Should -BeNullOrEmpty
                                (Get-TEntry $b 'FF5201').IsComplete | Should -BeFalse
                }

                It 'records a pin with no mission without completing anything' {
                                $s = New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00')
                                $row = New-TestRow -ETD '2026-10-20 07:30' -Instructor 'Jones, Bob' -Id 'P1'
                                $out = @(Process-HardcodedSortie -Sortie $row -Student $s -Mission $null -Schedule @($row) -Students @($s) -Instructors $script:Ips -SoftWeights $script:W -FormIndex 3 3>$null 6>$null)

                                $out | Should -Be @(3)
                                $row.FilledBy | Should -Be 'Pin'
                                $row.Mission | Should -BeNullOrEmpty
                                @($s.Syllabus | Where-Object { $_.IsComplete }) | Should -HaveCount 1
                                $s.SimDebt | Should -Be 0
                }

                It 'fills the O&B return leg' {
                                $s = New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00') -SimDebt 1
                                $ob = New-TestRow -ETD '2026-10-20 08:00' -SortieType 'O&B' -Instructor 'Jones, Bob' -Id 'OB1' -LinkedLineId 'RT1'
                                $rt = New-TestRow -ETD '2026-10-20 11:00' -SortieType 'O&B Return' -Instructor 'Lee, Ann' -Id 'RT1'
                                $out = @(Process-HardcodedSortie -Sortie $ob -Student $s -Mission (Get-TEntry $s 'FT5001') -Schedule @($ob, $rt) -Students @($s) -Instructors $script:Ips -SoftWeights $script:W -Formations $script:F -FormIndex 1 6>$null)

                                $out | Should -Be @(1)
                                $rt.FilledBy | Should -Be 'Return'
                                [object]::ReferenceEquals($rt.Instructor, $ob.Instructor) | Should -BeTrue
                                Get-NameText $rt.Mission | Should -Be 'FT5102'
                                $s.SimDebt | Should -Be 3
                }
}

Describe 'Pair-FormationWingman' {

                BeforeAll {
                                $script:Ips = New-TestInstructors
                                $script:W = New-TestSoftWeights
                                $script:F = New-TestFormations

                                function New-TLead {
                                                param ($Student, [string]$Notes = $null)
                                                $row = New-TestRow -ETD '2026-10-20 08:00' -Instructor 'Jones, Bob' -Id 'F1' -Notes $Notes
                                                $row.Instructor = Resolve-SortieInstructor -Instructors $script:Ips -Instructor 'Jones, Bob'
                                                $row.Student = $Student
                                                $row.Mission = Get-TEntry $Student 'FF5201'
                                                $row.Callsign = 'TALON'
                                                $row.Num = 31
                                                return $row
                                }
                }

                It 'picks a wingman inside the look-ahead window, skipping sim lines and lines before the lead' {
                                $a = New-TestStudent -Name 'Adams, Al' -Done $script:FormDone
                                $b = New-TestStudent -Name 'Brown, Bea' -Done $script:FormDone
                                $lead = New-TLead $a
                                $before = New-TestRow -ETD '2026-10-20 07:55' -Instructor 'Lee, Ann' -Id 'F0'
                                $ift = New-TestRow -ETD '2026-10-20 08:05' -SortieType 'IFT' -Id 'S1'
                                $oft = New-TestRow -ETD '2026-10-20 08:07' -SortieType 'OFT' -Id 'S2'
                                $wing = New-TestRow -ETD '2026-10-20 08:15' -Instructor 'Lee, Ann' -Id 'F2'
                                $out = @(Pair-FormationWingman -PrimarySortie $lead -PrimaryStudent $a -PrimaryMission $lead.Mission -Schedule @($before, $lead, $ift, $oft, $wing) `
                                                -Students @($a, $b) -Instructors $script:Ips -SoftWeights $script:W -FormIndex 0 -Formations $script:F)

                                $out | Should -Be @(2)
                                $lead.Callsign | Should -Be 'HAZE'
                                $lead.Num | Should -Be 11
                                $lead.Notes | Should -BeExactly 'Wingman: Brown, Bea'
                                [object]::ReferenceEquals($wing.Student, $b) | Should -BeTrue
                                Get-NameText $wing.Mission | Should -Be 'FF5201'
                                $wing.FilledBy | Should -Be 'Wing'
                                $wing.Num | Should -Be 12
                                $wing.Notes | Should -BeExactly 'Wingman to Adams, Al'
                                foreach ($x in @($before, $ift, $oft)) {
                                                $x.Student | Should -BeNullOrEmpty
                                                $x.FilledBy | Should -BeNullOrEmpty
                                }
                                $e = Get-TEntry $b 'FF5201'
                                $e.IsComplete | Should -BeTrue
                                $e.Instructor | Should -BeExactly 'Lee, Ann'
                                $e.ETD | Should -Be '2026-10-20 08:15:00'
                                $b.SimDebt | Should -Be 1
                }

                It 'includes a line exactly at the look-ahead limit and wraps FormIndex' {
                                $a = New-TestStudent -Name 'Adams, Al' -Done $script:FormDone
                                $b = New-TestStudent -Name 'Brown, Bea' -Done $script:FormDone
                                $lead = New-TLead $a
                                $wing = New-TestRow -ETD '2026-10-20 08:22' -Instructor 'Lee, Ann' -Id 'F2'
                                $out = @(Pair-FormationWingman -PrimarySortie $lead -PrimaryStudent $a -PrimaryMission $lead.Mission -Schedule @($lead, $wing) `
                                                -Students @($a, $b) -Instructors $script:Ips -SoftWeights $script:W -FormIndex 2 -Formations $script:F)

                                $out | Should -Be @(0)
                                $lead.Num | Should -Be 21
                                $wing.Num | Should -Be 22
                                $wing.FilledBy | Should -Be 'Wing'
                }

                It 'resets an out-of-range FormIndex' {
                                $a = New-TestStudent -Name 'Adams, Al' -Done $script:FormDone
                                $b = New-TestStudent -Name 'Brown, Bea' -Done $script:FormDone
                                $lead = New-TLead $a
                                $wing = New-TestRow -ETD '2026-10-20 08:10' -Instructor 'Lee, Ann' -Id 'F2'
                                $out = @(Pair-FormationWingman -PrimarySortie $lead -PrimaryStudent $a -PrimaryMission $lead.Mission -Schedule @($lead, $wing) `
                                                -Students @($a, $b) -Instructors $script:Ips -SoftWeights $script:W -FormIndex 9 -Formations $script:F)

                                $out | Should -Be @(2)
                                $lead.Num | Should -Be 11
                                $wing.Num | Should -Be 12
                }

                It "appends notes, keeps callsigns without FormCallsigns, and notes 'Wingman: none found' outside the window" {
                                $a = New-TestStudent -Name 'Adams, Al' -Done $script:FormDone
                                $b = New-TestStudent -Name 'Brown, Bea' -Done $script:FormDone
                                $lead = New-TLead $a 'Brief 0630'
                                $late = New-TestRow -ETD '2026-10-20 08:23' -Instructor 'Lee, Ann' -Id 'F2'
                                $out = @(Pair-FormationWingman -PrimarySortie $lead -PrimaryStudent $a -PrimaryMission $lead.Mission -Schedule @($lead, $late) `
                                                -Students @($a, $b) -Instructors $script:Ips -SoftWeights $script:W -FormIndex 0 -Formations @())

                                $out | Should -Be @(0)
                                $lead.Notes | Should -BeExactly 'Brief 0630 | Wingman: none found'
                                $lead.Callsign | Should -Be 'TALON'
                                $lead.Num | Should -Be 31
                                $late.Student | Should -BeNullOrEmpty
                }

                It 'with an existing pinned wing only sets its callsign and notes' {
                                $a = New-TestStudent -Name 'Adams, Al' -Done $script:FormDone
                                $b = New-TestStudent -Name 'Brown, Bea' -Done $script:FormDone
                                $lead = New-TLead $a
                                $pinned = New-TestRow -ETD '2026-10-20 08:10' -Instructor 'Lee, Ann' -Id 'F2' -WingOfLineId 'F1' -Notes 'Brief together'
                                $pinned.Student = $b
                                $open = New-TestRow -ETD '2026-10-20 08:05' -Instructor 'Ortiz, Sam' -Id 'X'
                                $out = @(Pair-FormationWingman -PrimarySortie $lead -PrimaryStudent $a -PrimaryMission $lead.Mission -Schedule @($lead, $open, $pinned) `
                                                -Students @($a, $b) -Instructors $script:Ips -SoftWeights $script:W -FormIndex 0 -Formations $script:F -ExistingWingSortie $pinned)

                                $out | Should -Be @(2)
                                $pinned.Callsign | Should -Be 'HAZE'
                                $pinned.Num | Should -Be 12
                                $pinned.Notes | Should -BeExactly 'Brief together | Wingman to Adams, Al'
                                $lead.Notes | Should -BeExactly 'Wingman: Brown, Bea'
                                $open.Student | Should -BeNullOrEmpty
                                (Get-TEntry $b 'FF5201').IsComplete | Should -BeFalse
                }

                It 'skips a candidate who fails the hard rules and falls back to Direct Support' {
                                $a = New-TestStudent -Name 'Adams, Al' -Done $script:FormDone
                                $b = New-TestStudent -Name 'Brown, Bea' -Done $script:FormDone -OnHold
                                $lead = New-TLead $a
                                $ds = New-TestRow -ETD '2026-10-20 08:10' -Instructor 'Lee, Ann' -Id 'F2'
                                $out = @(Pair-FormationWingman -PrimarySortie $lead -PrimaryStudent $a -PrimaryMission $lead.Mission -Schedule @($lead, $ds) `
                                                -Students @($a, $b) -Instructors $script:Ips -SoftWeights $script:W -FormIndex 0 -Formations $script:F)

                                $out | Should -Be @(2)
                                $ds.Student | Should -BeExactly 'Need Form SP'
                                $ds.FilledBy | Should -Be 'DS'
                                $ds.Notes | Should -BeExactly 'Direct Support for Adams, Al'
                                $ds.Num | Should -Be 12
                                $lead.Notes | Should -BeExactly 'Wingman: DS'
                                (Get-TEntry $b 'FF5201').IsComplete | Should -BeFalse
                                $b.SimDebt | Should -Be 0
                }

                It 'chooses between equal wingmen by name' {
                                $a = New-TestStudent -Name 'Adams, Al' -Done $script:FormDone
                                $c = New-TestStudent -Name 'Carter, Cy' -Done $script:FormDone
                                $b = New-TestStudent -Name 'Brown, Bea' -Done $script:FormDone
                                $lead = New-TLead $a
                                $wing = New-TestRow -ETD '2026-10-20 08:10' -Instructor 'Lee, Ann' -Id 'F2'
                                [void](Pair-FormationWingman -PrimarySortie $lead -PrimaryStudent $a -PrimaryMission $lead.Mission -Schedule @($lead, $wing) `
                                                -Students @($a, $c, $b) -Instructors $script:Ips -SoftWeights $script:W -FormIndex 0 -Formations $script:F)

                                Get-NameText $wing.Student | Should -Be 'Brown, Bea'
                                (Get-TEntry $c 'FF5201').IsComplete | Should -BeFalse
                }
}

Describe 'Schedule-ReturnLeg' {

                BeforeAll {
                                $script:Ips = New-TestInstructors
                                $script:W = New-TestSoftWeights

                                # Student whose O&B leg (FT5001 at 0800 with Jones) is already recorded
                                function New-TObStudent {
                                                return (New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00', 'FT5001|2026-10-20 08:00|Jones, Bob') -SimDebt 1)
                                }
                                function New-TObRow {
                                                param ([string]$LinkedLineId = '')
                                                $row = New-TestRow -ETD '2026-10-20 08:00' -SortieType 'O&B' -Instructor 'Jones, Bob' -Id 'OB1' -LinkedLineId $LinkedLineId
                                                $row.Instructor = Resolve-SortieInstructor -Instructors $script:Ips -Instructor 'Jones, Bob'
                                                $row.Callsign = 'TALON'
                                                $row.Num = 31
                                                $row.FilledBy = 'Engine'
                                                return $row
                                }
                }

                It 'fills the linked return before a row the legacy lookup would match' {
                                $s = New-TObStudent
                                $ob = New-TObRow -LinkedLineId 'RT1'
                                $ob.Student = $s
                                $decoy = New-TestRow -ETD '2026-10-20 10:30' -SortieType 'O&B Return' -Instructor 'Jones, Bob' -Id 'RT2'
                                $rt = New-TestRow -ETD '2026-10-20 11:00' -SortieType 'O&B Return' -Instructor 'Lee, Ann' -Id 'RT1'
                                Schedule-ReturnLeg -Sortie $ob -Student $s -Schedule @($ob, $decoy, $rt) -Instructors $script:Ips -SoftWeights $script:W

                                [object]::ReferenceEquals($rt.Student, $s) | Should -BeTrue
                                $rt.FilledBy | Should -Be 'Return'
                                $decoy.Student | Should -BeNullOrEmpty
                                $decoy.FilledBy | Should -BeNullOrEmpty
                }

                It "stores the O&B IP's name (not the object) in the syllabus and copies the callsign" {
                                $s = New-TObStudent
                                $ob = New-TObRow -LinkedLineId 'RT1'
                                $rt = New-TestRow -ETD '2026-10-20 11:00' -SortieType 'O&B Return' -Instructor 'Lee, Ann' -Id 'RT1'
                                Schedule-ReturnLeg -Sortie $ob -Student $s -Schedule @($ob, $rt) -Instructors $script:Ips -SoftWeights $script:W

                                [object]::ReferenceEquals($rt.Instructor, $ob.Instructor) | Should -BeTrue
                                $rt.Callsign | Should -Be 'TALON'
                                $rt.Num | Should -Be 31
                                Get-NameText $rt.Mission | Should -Be 'FT5102'
                                $e = Get-TEntry $s 'FT5102'
                                $e.IsComplete | Should -BeTrue
                                $e.Instructor | Should -BeOfType [string]
                                $e.Instructor | Should -BeExactly 'Jones, Bob'
                                $e.ETD | Should -Be '2026-10-20 11:00:00'
                                $s.SimDebt | Should -Be 2
                }

                It "releases the student's reservation on the return line" {
                                $s = New-TObStudent
                                Add-StudentReservation -Student $s -LineId 'RT1' -ETD '2026-10-20 11:00' -SortieType 'O&B Return'
                                Add-StudentReservation -Student $s -LineId 'Z9' -ETD '2026-10-22 08:00' -SortieType 'Day Local' -Mission 'FT5103'
                                $ob = New-TObRow -LinkedLineId 'RT1'
                                $rt = New-TestRow -ETD '2026-10-20 11:00' -SortieType 'O&B Return' -Instructor 'Jones, Bob' -Id 'RT1'
                                Schedule-ReturnLeg -Sortie $ob -Student $s -Schedule @($ob, $rt) -Instructors $script:Ips -SoftWeights $script:W

                                @($s.Reservations | ForEach-Object { $_.LineId }) | Should -Be @('Z9')
                }

                It 'falls back to the same-day same-IP open return after the O&B when there is no link' {
                                $s = New-TObStudent
                                $ob = New-TObRow
                                $otherIp = New-TestRow -ETD '2026-10-20 11:00' -SortieType 'O&B Return' -Instructor 'Lee, Ann' -Id 'RTa'
                                $otherDay = New-TestRow -ETD '2026-10-21 11:00' -SortieType 'O&B Return' -Instructor 'Jones, Bob' -Id 'RTb'
                                $earlier = New-TestRow -ETD '2026-10-20 07:00' -SortieType 'O&B Return' -Instructor 'Jones, Bob' -Id 'RTd'
                                $match = New-TestRow -ETD '2026-10-20 11:00' -SortieType 'O&B Return' -Instructor 'Jones, Bob' -Id 'RTc'
                                Schedule-ReturnLeg -Sortie $ob -Student $s -Schedule @($ob, $otherIp, $otherDay, $earlier, $match) -Instructors $script:Ips -SoftWeights $script:W

                                $match.FilledBy | Should -Be 'Return'
                                foreach ($x in @($otherIp, $otherDay, $earlier)) { $x.FilledBy | Should -BeNullOrEmpty }
                }

                It 'notes the O&B and warns when there is no return line' {
                                $s = New-TObStudent
                                $ob = New-TObRow
                                $out = @(Schedule-ReturnLeg -Sortie $ob -Student $s -Schedule @($ob) -Instructors $script:Ips -SoftWeights $script:W 3>&1)

                                $ob.Notes | Should -BeExactly 'No O&B Return line found'
                                @($out | ForEach-Object { [string]$_.Message }) | Should -Contain 'Could not find an empty O&B Return row on 2026-10-20 for Jones, Bob.'
                                $s.SimDebt | Should -Be 1
                }

                It 'notes the return when the student has no flyable event left' {
                                $s = New-TestStudent -Name 'Able, Amy' -Done @('FI3101|2026-10-15 08:00', 'FT5001|2026-10-16 08:00|Jones, Bob', 'FT5102|2026-10-17 08:00|Jones, Bob', 'FT5103|2026-10-18 08:00|Jones, Bob', 'FT5104|2026-10-19 07:00|Lee, Ann', 'FT5190|2026-10-20 08:00|Jones, Bob')
                                $ob = New-TObRow -LinkedLineId 'RT1'
                                $rt = New-TestRow -ETD '2026-10-20 11:00' -SortieType 'O&B Return' -Instructor 'Jones, Bob' -Id 'RT1'
                                Schedule-ReturnLeg -Sortie $ob -Student $s -Schedule @($ob, $rt) -Instructors $script:Ips -SoftWeights $script:W

                                $rt.FilledBy | Should -Be 'Return'
                                $rt.Mission | Should -BeNullOrEmpty
                                $rt.Notes | Should -BeExactly 'No valid mission for return leg'
                                [object]::ReferenceEquals($rt.Student, $s) | Should -BeTrue
                }
}
