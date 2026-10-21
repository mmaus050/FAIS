# Pester 5 tests for Test-HardRules, SoftRules and the Syllabus helpers they use.
# Dates used throughout (2026): Fri 20 Feb .. Fri 27 Feb history, Mon 02 Mar,
# Tue 03 Mar = the day being scheduled, Wed 04 Mar = the next day.

BeforeAll {
                . "$PSScriptRoot/Fixtures/New-FaisTestData.ps1"
                Import-FaisModules

                $script:IPs = New-TestInstructors
                $script:W = New-TestSoftWeights

                # Everything except the formation events: Day Local opts only FF5201
                $script:FormDone = @(
                                'FI3101|2026-02-20 08:00',
                                'FE3102|2026-02-20 13:00',
                                'FT5001|2026-02-23 08:00|Lee, Ann',
                                'FT5102|2026-02-24 08:00|Lee, Ann',
                                'FT5103|2026-02-25 08:00|Lee, Ann',
                                'FN3201|2026-02-25 13:00',
                                'FT5104|2026-02-26 08:00|Lee, Ann',
                                'FT5190|2026-02-27 08:00|Jones, Bob'
                )
                # Everything except the checkride (FT5104 flown by Jones)
                $script:AllButCheck = @(
                                'FI3101|2026-02-20 08:00',
                                'FE3102|2026-02-20 13:00',
                                'FT5001|2026-02-23 08:00|Lee, Ann',
                                'FT5102|2026-02-24 08:00|Lee, Ann',
                                'FT5103|2026-02-25 08:00|Ortiz, Sam',
                                'FN3201|2026-02-25 13:00',
                                'FF5201|2026-02-26 08:00|Lee, Ann',
                                'FF5202|2026-02-26 13:00|Lee, Ann',
                                'FT5104|2026-02-27 08:00|Jones, Bob'
                )

                function Invoke-HR {
                                param (
                                                $Student,
                                                $Sortie,
                                                [array]$Schedule = @(),
                                                [array]$Students = @(),
                                                $Mission = $null,
                                                [double]$FormLookAhead = 22,
                                                [hashtable]$LineContext = $null,
                                                [switch]$Explain,
                                                [switch]$IsWingCandidate
                                )
                                $splat = @{
                                                Student         = $Student
                                                Sortie          = $Sortie
                                                Schedule        = $Schedule
                                                Students        = $Students
                                                Instructors     = $script:IPs
                                                FormLookAhead   = $FormLookAhead
                                                Explain         = $Explain
                                                IsWingCandidate = $IsWingCandidate
                                }
                                if ($null -ne $Mission) { $splat['Mission'] = $Mission }
                                if ($null -ne $LineContext) { $splat['LineContext'] = $LineContext }
                                return (Test-HardRules @splat)
                }

                function Get-RuleCodes {
                                param ($Result)
                                return @(@($Result.Violations) | ForEach-Object { $_.Rule })
                }

                function Get-Viol {
                                param ($Result, [string]$Rule)
                                return (@($Result.Violations) | Where-Object { $_.Rule -eq $Rule } | Select-Object -First 1)
                }

                function Get-EvtResult {
                                param ($Result, [string]$Mission)
                                return (@($Result.EventResults) | Where-Object { $_.Mission -eq $Mission } | Select-Object -First 1)
                }

                function Get-EvtViol {
                                param ($Result, [string]$Mission, [string]$Rule)
                                $er = Get-EvtResult $Result $Mission
                                if ($null -eq $er) { return $null }
                                return (@($er.Violations) | Where-Object { $_.Rule -eq $Rule } | Select-Object -First 1)
                }

                function New-Absence {
                                param ([string]$Start, [string]$End, [string]$Reason = '')
                                return [PSCustomObject]@{ Student = 'x'; StartTime = [datetime]$Start; EndTime = [datetime]$End; Reason = $Reason }
                }

                function Get-MissionNames {
                                param ($Events)
                                return @(@($Events) | Where-Object { $null -ne $_ } | ForEach-Object { [string]$_.Mission })
                }

                function Get-Soft {
                                param (
                                                $Student,
                                                [string]$ETD = '2026-03-03 10:00',
                                                [string]$SortieType = 'Day Local',
                                                [string]$Instructor = '',
                                                [string]$Mission,
                                                $Weights = $script:W,
                                                [array]$Instructors = $script:IPs
                                )
                                $row = New-TestRow -ETD $ETD -SortieType $SortieType -Instructor $Instructor
                                $row.Mission = Find-SyllabusMission -Student $Student -Mission $Mission
                                if ($null -eq $row.Mission) { throw "test setup: $Mission not in syllabus" }
                                return (SoftRules -Student $Student -Config $Weights -Sortie $row -Instructors $Instructors -Detailed)
                }

                function New-MiniSyllabus {
                                param ([array]$Rows)
                                return @(foreach ($r in $Rows) {
                                                [PSCustomObject]@{
                                                                SyllabusID = $r[0]; Mission = $r[1]; Prereqs = $r[2]; Type = $r[3]
                                                                IsCheck = $false; IsComplete = $r[4]; Instructor = ''; ETD = ''; Level = 1
                                                }
                                })
                }
}


# =================================================================================
#  SYLLABUS HELPERS
# =================================================================================
Describe 'Get-OptedEvents' {
                It 'opts only the first event on a fresh syllabus' {
                                Get-MissionNames (Get-OptedEvents -SyllabusData (New-TestSyllabus)) | Should -Be @('FI3101')
                }

                It 'opts every incomplete event whose prereqs are all complete' {
                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00'
                                Get-MissionNames (Get-OptedEvents -SyllabusData $s.Syllabus) | Should -Be @('FE3102', 'FT5001')
                }

                It 'returns an empty result for an empty syllabus' {
                                @(Get-OptedEvents -SyllabusData @()).Count | Should -Be 0
                }

                It 'treats text booleans correctly (FALSE text is not complete, TRUE text is)' {
                                $syl = New-TestSyllabus
                                $syl[0].IsComplete = 'FALSE'
                                Get-MissionNames (Get-OptedEvents -SyllabusData $syl) | Should -Be @('FI3101')
                                $syl[0].IsComplete = 'TRUE'
                                Get-MissionNames (Get-OptedEvents -SyllabusData $syl) | Should -Be @('FE3102', 'FT5001')
                }

                It 'handles comma-separated, case-insensitive prereqs and blank prereqs' {
                                $syl = New-MiniSyllabus @(
                                                , @(1, 'AA1001', '0', 'Single', $true)
                                                , @(2, 'AA1002', 'aa1001', 'Single', $false)
                                                , @(3, 'AA1003', 'AA1001, AA1002', 'Single', $false)
                                                , @(4, 'AA1004', '', 'Single', $false)
                                )
                                Get-MissionNames (Get-OptedEvents -SyllabusData $syl) | Should -Be @('AA1002', 'AA1004')
                                $syl[1].IsComplete = $true
                                Get-MissionNames (Get-OptedEvents -SyllabusData $syl) | Should -Be @('AA1003', 'AA1004')
                }

                It '-AssumeComplete treats extra missions as complete without touching the syllabus' {
                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00'
                                $before = $s.Syllabus | ConvertTo-Json -Depth 5
                                Get-MissionNames (Get-OptedEvents -SyllabusData $s.Syllabus -AssumeComplete @('ft5001')) | Should -Be @('FE3102', 'FT5102')
                                ($s.Syllabus | ConvertTo-Json -Depth 5) | Should -Be $before
                }

                It '-ExcludeMissions drops reserved missions (and does not complete them)' {
                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00'
                                Get-MissionNames (Get-OptedEvents -SyllabusData $s.Syllabus -ExcludeMissions @('FE3102')) | Should -Be @('FT5001')
                                Get-MissionNames (Get-OptedEvents -SyllabusData $s.Syllabus -ExcludeMissions @('FT5001')) | Should -Be @('FE3102')
                }
}

Describe 'Get-ValidOptedEvents' {
                It 'filters by sortie type: <SortieType>' -TestCases @(
                                @{ SortieType = 'Day Local'; Expected = @('FT5103') }
                                @{ SortieType = 'OFT'; Expected = @('FE3102', 'FN3201') }
                                @{ SortieType = 'IFT'; Expected = @('FN3201') }
                ) {
                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00', 'FT5001|2026-02-23 08:00|Jones, Bob', 'FT5102|2026-02-24 08:00|Jones, Bob'
                                $row = New-TestRow -ETD '2026-03-03 08:00' -SortieType $SortieType
                                Get-MissionNames (Get-ValidOptedEvents -Student $s -Sortie $row) | Should -Be $Expected
                }

                It 'returns nothing for an unknown sortie type' {
                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00'
                                @(Get-ValidOptedEvents -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -SortieType 'Ground')).Count | Should -Be 0
                }

                It 'sorts by SyllabusID numerically even when IDs are text' {
                                $s = New-TestStudent -Name 'A'
                                $s.Syllabus = New-MiniSyllabus @(
                                                , @('10', 'AA1010', '0', 'Single', $false)
                                                , @('9', 'AA1009', '0', 'Single', $false)
                                )
                                Get-MissionNames (Get-ValidOptedEvents -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00')) | Should -Be @('AA1009', 'AA1010')
                }

                It 'never offers solo events on an O&B' {
                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00', 'FT5001|2026-02-23 08:00|Jones, Bob', 'FT5102|2026-02-24 08:00|Jones, Bob'
                                @(Get-ValidOptedEvents -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -SortieType 'O&B')).Count | Should -Be 0
                }

                It 'O&B look-ahead: keeps the event when a flyable second leg follows' {
                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00'
                                Get-MissionNames (Get-ValidOptedEvents -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -SortieType 'O&B')) | Should -Be @('FT5001')
                }

                It 'O&B look-ahead: a solo cannot be the second leg' {
                                # FT5102 would be leg 1, but after it only FT5103 (SingleSolo) and sims open up
                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00', 'FT5001|2026-02-23 08:00|Jones, Bob'
                                @(Get-ValidOptedEvents -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -SortieType 'O&B')).Count | Should -Be 0
                                Get-MissionNames (Get-ValidOptedEvents -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -SortieType 'Day Local')) | Should -Be @('FT5102')
                }

                It 'O&B look-ahead: nothing when the last event has no second leg; -NoLookahead skips the check' {
                                $s = New-TestStudent -Name 'A' -Done $script:AllButCheck
                                $ob = New-TestRow -ETD '2026-03-03 08:00' -SortieType 'O&B'
                                @(Get-ValidOptedEvents -Student $s -Sortie $ob).Count | Should -Be 0
                                Get-MissionNames (Get-ValidOptedEvents -Student $s -Sortie $ob -NoLookahead) | Should -Be @('FT5190')
                }

                It 'O&B look-ahead never mutates the syllabus' {
                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00'
                                $before = $s | ConvertTo-Json -Depth 6
                                $null = Get-ValidOptedEvents -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -SortieType 'O&B')
                                ($s | ConvertTo-Json -Depth 6) | Should -Be $before
                }

                It '-AssumeComplete simulates completions' {
                                $s = New-TestStudent -Name 'A'
                                Get-MissionNames (Get-ValidOptedEvents -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00') -AssumeComplete @('FI3101')) | Should -Be @('FT5001')
                                $s.Syllabus[0].IsComplete | Should -BeFalse
                }

                It 'excludes missions reserved on other lines but not on the sortie''s own line' {
                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00'
                                Add-StudentReservation -Student $s -LineId 'P1' -ETD '2026-03-05 10:00' -SortieType 'Day Local' -Mission 'FT5001' -Instructor 'Jones, Bob'
                                @(Get-ValidOptedEvents -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -Id 'L1')).Count | Should -Be 0
                                Get-MissionNames (Get-ValidOptedEvents -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -Id 'P1')) | Should -Be @('FT5001')
                }
}

Describe 'Sim debt and Complete-StudentMission' {
                It 'Get-SimDebtDelta: flights add, transition and formation sims pay down' {
                                Get-SimDebtDelta -MissionName 'FT5001' -SimStopMult 1 -TransSimRatio 0.5 -FormSimRatio 0.25 | Should -Be 1
                                Get-SimDebtDelta -MissionName 'FF5201' -SimStopMult 1 -TransSimRatio 0.5 -FormSimRatio 0.25 | Should -Be 1
                                Get-SimDebtDelta -MissionName 'FE3102' -SimStopMult 1 -TransSimRatio 0.5 -FormSimRatio 0.25 | Should -Be -0.5
                                Get-SimDebtDelta -MissionName 'FI3101' -SimStopMult 1 -TransSimRatio 0.5 -FormSimRatio 0.25 | Should -Be -0.5
                                Get-SimDebtDelta -MissionName 'FN3201' -SimStopMult 1 -TransSimRatio 0.5 -FormSimRatio 0.25 | Should -Be -0.25
                                Get-SimDebtDelta -MissionName 'GR1001' -SimStopMult 1 -TransSimRatio 0.5 -FormSimRatio 0.25 | Should -Be 0
                }

                It 'marks the mission complete with IP and ETD and applies flight sim debt once' {
                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00' -SimDebt 2
                                $entry = Complete-StudentMission -Student $s -Mission 'FT5001' -InstructorName 'Jones, Bob' -ETD '2026-03-03 07:30' -SoftWeights $script:W
                                $entry.Mission | Should -Be 'FT5001'
                                $m = Find-SyllabusMission -Student $s -Mission 'FT5001'
                                $m.IsComplete | Should -BeTrue
                                $m.Instructor | Should -Be 'Jones, Bob'
                                $m.ETD | Should -Be '2026-03-03 07:30:00'
                                $s.SimDebt | Should -Be 3
                                $s.SyllabusRevision | Should -Be 1

                                # Completing it again does not double the debt
                                $null = Complete-StudentMission -Student $s -Mission 'FT5001' -InstructorName 'Ortiz, Sam' -ETD '2026-03-03 09:00' -SoftWeights $script:W
                                $s.SimDebt | Should -Be 3
                                $m.Instructor | Should -Be 'Ortiz, Sam'
                                $s.SyllabusRevision | Should -Be 2
                }

                It 'sims pay sim debt down (<Mission>)' -TestCases @(
                                @{ Mission = 'FE3102'; Done = @('FI3101|2026-02-20 08:00'); Expected = 1.5 }
                                @{ Mission = 'FN3201'; Done = @('FI3101|2026-02-20 08:00', 'FT5001|2026-02-23 08:00|Jones, Bob', 'FT5102|2026-02-24 08:00|Jones, Bob'); Expected = 1.5 }
                ) {
                                $s = New-TestStudent -Name 'A' -Done $Done -SimDebt 2
                                $evt = Find-SyllabusMission -Student $s -Mission $Mission
                                $null = Complete-StudentMission -Student $s -Mission $evt -InstructorName 'SIM IP' -ETD '2026-03-03 07:30' -SoftWeights $script:W
                                $s.SimDebt | Should -Be $Expected
                }

                It 'leaves sim debt alone with -NoSimDebt or without SoftWeights' {
                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00' -SimDebt 2
                                $null = Complete-StudentMission -Student $s -Mission 'FT5001' -InstructorName 'Jones, Bob' -ETD '2026-03-03 07:30' -SoftWeights $script:W -NoSimDebt
                                $s.SimDebt | Should -Be 2
                                $null = Complete-StudentMission -Student $s -Mission 'FE3102' -InstructorName 'SIM IP' -ETD '2026-03-03 11:00'
                                $s.SimDebt | Should -Be 2
                }

                It 'returns $null and changes nothing for a mission not in the syllabus' {
                                $s = New-TestStudent -Name 'A' -SimDebt 2
                                $before = $s | ConvertTo-Json -Depth 6
                                Complete-StudentMission -Student $s -Mission 'ZZ9999' -InstructorName 'Jones, Bob' -ETD '2026-03-03 07:30' -SoftWeights $script:W | Should -BeNullOrEmpty
                                ($s | ConvertTo-Json -Depth 6) | Should -Be $before
                }

                It 'Update-StudentSimDebtFromSyllabus recomputes from every completed event' {
                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00', 'FT5001|2026-02-23 08:00|Jones, Bob', 'FT5102|2026-02-24 08:00|Jones, Bob' -SimDebt 99
                                Update-StudentSimDebtFromSyllabus -Student $s -SoftWeights $script:W | Should -Be 1.5
                                $s.SimDebt | Should -Be 1.5
                }
}

Describe 'Undo-StudentMission and Reset-PlannedCompletions' {
                It 'Undo-StudentMission backs out only that event''s sim debt (a hand-set SimDebt is kept)' {
                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00', 'FT5001|2026-02-23 08:00|Jones, Bob' -SimDebt 7
                                Undo-StudentMission -Student $s -Mission 'FT5001' -SoftWeights $script:W | Should -BeTrue
                                $m = Find-SyllabusMission -Student $s -Mission 'FT5001'
                                $m.IsComplete | Should -BeFalse
                                $m.Instructor | Should -Be ''
                                $m.ETD | Should -Be ''
                                $s.SimDebt | Should -Be 6
                                $s.SyllabusRevision | Should -Be 1

                                # A sim pays debt down, so undoing one puts it back
                                Undo-StudentMission -Student $s -Mission 'FI3101' -SoftWeights $script:W | Should -BeTrue
                                $s.SimDebt | Should -Be 6.5
                }

                It 'Undo-StudentMission changes nothing for an event that is not complete' {
                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00' -SimDebt 7
                                $before = $s | ConvertTo-Json -Depth 6
                                Undo-StudentMission -Student $s -Mission 'FT5001' -SoftWeights $script:W | Should -BeFalse
                                Undo-StudentMission -Student $s -Mission 'ZZ9999' -SoftWeights $script:W | Should -BeFalse
                                ($s | ConvertTo-Json -Depth 6) | Should -Be $before
                }

                It 'Undo-StudentMission leaves SimDebt alone without SoftWeights' {
                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00', 'FT5001|2026-02-23 08:00|Jones, Bob' -SimDebt 7
                                Undo-StudentMission -Student $s -Mission 'FT5001' | Should -BeTrue
                                $s.SimDebt | Should -Be 7
                }

                It 'Reset-PlannedCompletions clears events on/after the day and backs out their debt only' {
                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00', 'FT5001|2026-03-03 07:30|Jones, Bob', 'FT5102|2026-03-04 07:30|Jones, Bob' -SimDebt 10
                                $other = New-TestStudent -Name 'B' -Done 'FI3101|2026-02-20 08:00' -SimDebt 3
                                $changed = @(Reset-PlannedCompletions -Students @($s, $other) -FromDate ([datetime]'2026-03-03') -SoftWeights $script:W)
                                @($changed | ForEach-Object { $_.Name }) | Should -Be @('A')
                                (Find-SyllabusMission -Student $s -Mission 'FI3101').IsComplete | Should -BeTrue
                                (Find-SyllabusMission -Student $s -Mission 'FT5001').IsComplete | Should -BeFalse
                                (Find-SyllabusMission -Student $s -Mission 'FT5102').IsComplete | Should -BeFalse
                                $s.SimDebt | Should -Be 8
                                $other.SimDebt | Should -Be 3
                }

                It 'Reset-PlannedCompletions puts back an unsat that a re-planned line cleaned up' {
                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00'
                                $f = New-TestStudent -Name 'B' -Done 'FI3101|2026-02-20 08:00'
                                $old = New-TestStudent -Name 'C' -Done 'FI3101|2026-02-20 08:00'
                                $prev = @(
                                                (New-TestRow -ETD '2026-03-03 07:30' -SortieType 'IFT' -Student 'A' -Mission 'FE3102' -Notes 'Clears Unsat Sim (IFT)')
                                                (New-TestRow -ETD '2026-03-03 07:30' -Instructor 'Jones, Bob' -Student $f -Mission 'FT5001' -Notes 'Solo Flight | Clears Unsat Flight')
                                                (New-TestRow -ETD '2026-03-02 07:30' -SortieType 'OFT' -Student 'C' -Mission 'FE3102' -Notes 'Clears Unsat Sim (OFT)')
                                )
                                $changed = @(Reset-PlannedCompletions -Students @($s, $f, $old) -FromDate ([datetime]'2026-03-03') -SoftWeights $script:W -PrevSchedule $prev)
                                @($changed | ForEach-Object { $_.Name } | Sort-Object) | Should -Be @('A', 'B')
                                $s.UnsatSim | Should -BeTrue
                                $s.UnsatSimType | Should -Be 'IFT'
                                $f.UnsatFlight | Should -BeTrue
                                $f.UnsatSim | Should -BeFalse
                                # Cleaned up on a day that is not being re-planned: stays cleaned up
                                $old.UnsatSim | Should -BeFalse
                }
}

Describe 'Unsat clean-up' {
                It 'an OFT hook is cleaned up by an OFT but not an IFT' {
                                $s = New-TestStudent -Name 'A' -UnsatSim -UnsatSimType 'OFT'
                                Clear-CleanedUnsat -Student $s -SortieType 'IFT' | Should -BeNullOrEmpty
                                $s.UnsatSim | Should -BeTrue
                                Clear-CleanedUnsat -Student $s -SortieType 'Day Local' | Should -BeNullOrEmpty
                                $s.UnsatSim | Should -BeTrue
                                Clear-CleanedUnsat -Student $s -SortieType 'OFT' | Should -Be 'Clears Unsat Sim (OFT)'
                                $s.UnsatSim | Should -BeFalse
                }

                It 'an IFT hook is cleaned up by an <SortieType>' -TestCases @(
                                @{ SortieType = 'IFT' }
                                @{ SortieType = 'OFT' }
                ) {
                                $s = New-TestStudent -Name 'A' -UnsatSim -UnsatSimType 'IFT'
                                Clear-CleanedUnsat -Student $s -SortieType $SortieType | Should -Be 'Clears Unsat Sim (IFT)'
                                $s.UnsatSim | Should -BeFalse
                }

                It 'a hook of unknown type needs an OFT' {
                                $s = New-TestStudent -Name 'A' -UnsatSim
                                Get-UnsatSimKind $s | Should -Be 'OFT'
                                Clear-CleanedUnsat -Student $s -SortieType 'IFT' | Should -BeNullOrEmpty
                                Clear-CleanedUnsat -Student $s -SortieType 'OFT' | Should -Be 'Clears Unsat Sim (OFT)'
                }

                It 'an unsat flight is cleaned up by a <SortieType> but not a sim' -TestCases @(
                                @{ SortieType = 'Day Local' }
                                @{ SortieType = 'O&B' }
                                @{ SortieType = 'O&B Return' }
                ) {
                                $s = New-TestStudent -Name 'A' -UnsatFlight
                                Clear-CleanedUnsat -Student $s -SortieType 'OFT' | Should -BeNullOrEmpty
                                $s.UnsatFlight | Should -BeTrue
                                Clear-CleanedUnsat -Student $s -SortieType $SortieType | Should -Be 'Clears Unsat Flight'
                                $s.UnsatFlight | Should -BeFalse
                }

                It 'does nothing for a student with no unsat' {
                                $s = New-TestStudent -Name 'A'
                                Clear-CleanedUnsat -Student $s -SortieType 'OFT' | Should -BeNullOrEmpty
                                Clear-CleanedUnsat -Student $s -SortieType 'Day Local' | Should -BeNullOrEmpty
                                $s.UnsatSim | Should -BeFalse
                                $s.UnsatFlight | Should -BeFalse
                }

                It 'Complete-StudentMission clears the unsat the line cleans up and notes the line' {
                                $s = New-TestStudent -Name 'A' -UnsatSim -UnsatSimType 'OFT' -Done 'FI3101|2026-02-20 08:00'
                                $row = New-TestRow -ETD '2026-03-03 07:30' -SortieType 'OFT'
                                $null = Complete-StudentMission -Student $s -Mission 'FE3102' -InstructorName 'SIM IP' -ETD $row.ETD -SoftWeights $script:W -Sortie $row
                                $s.UnsatSim | Should -BeFalse
                                $row.Notes | Should -Be 'Clears Unsat Sim (OFT)'
                }

                It 'Complete-StudentMission without a sortie type leaves the unsat alone' {
                                $s = New-TestStudent -Name 'A' -UnsatSim -Done 'FI3101|2026-02-20 08:00'
                                $null = Complete-StudentMission -Student $s -Mission 'FE3102' -InstructorName 'SIM IP' -ETD '2026-03-03 07:30' -SoftWeights $script:W
                                $s.UnsatSim | Should -BeTrue
                }

                It 'Get-LastSimType is the latest sim flown up to the cutoff' {
                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00', 'FE3102|2026-02-21 08:00', 'FT5001|2026-02-23 08:00|Jones, Bob', 'FT5102|2026-02-24 08:00|Jones, Bob', 'FN3201|2026-03-05 08:00'
                                Get-LastSimType -Student $s -Before ([datetime]'2026-03-01') | Should -Be 'OFT'
                                Get-LastSimType -Student $s -Before ([datetime]'2026-03-06') | Should -Be 'IFT'
                                Get-LastSimType -Student (New-TestStudent -Name 'B') -Before ([datetime]'2026-03-06') | Should -Be ''
                }
}

Describe 'Student reservations' {
                It 'adds a reservation with a parsed start and IsSim from the sortie type' {
                                $s = New-TestStudent -Name 'A'
                                Add-StudentReservation -Student $s -LineId 'P1' -ETD '2026-03-03 13:00:00' -SortieType 'OFT' -Mission 'FE3102' -Instructor 'SIM IP'
                                Add-StudentReservation -Student $s -LineId 'P2' -ETD '2026-03-04 08:00:00' -SortieType 'Day Local' -Mission 'FT5001' -Instructor 'Jones, Bob'
                                $res = @($s.Reservations)
                                $res.Count | Should -Be 2
                                $res[0].LineId | Should -Be 'P1'
                                $res[0].Start | Should -BeOfType [datetime]
                                $res[0].Start | Should -Be ([datetime]'2026-03-03 13:00')
                                $res[0].IsSim | Should -BeTrue
                                $res[1].IsSim | Should -BeFalse
                                $res[1].Instructor | Should -Be 'Jones, Bob'
                }

                It 'replaces a reservation for the same line and ignores a bad ETD' {
                                $s = New-TestStudent -Name 'A'
                                Add-StudentReservation -Student $s -LineId 'P1' -ETD '2026-03-03 13:00' -SortieType 'Day Local' -Mission 'FT5001'
                                Add-StudentReservation -Student $s -LineId 'P1' -ETD '2026-03-03 15:00' -SortieType 'Day Local' -Mission 'FT5102'
                                Add-StudentReservation -Student $s -LineId 'P3' -ETD 'not a time' -SortieType 'Day Local'
                                $res = @($s.Reservations)
                                $res.Count | Should -Be 1
                                $res[0].Mission | Should -Be 'FT5102'
                                $res[0].Start | Should -Be ([datetime]'2026-03-03 15:00')
                }

                It 'adds the Reservations property when the student has none' {
                                $s = [PSCustomObject]@{ Name = 'Bare'; Syllabus = @() }
                                Add-StudentReservation -Student $s -LineId 'P1' -ETD '2026-03-03 13:00' -SortieType 'Day Local'
                                @($s.Reservations).Count | Should -Be 1
                }

                It 'removes reservations by line id' {
                                $s = New-TestStudent -Name 'A'
                                foreach ($id in 'P1', 'P2', 'P3') { Add-StudentReservation -Student $s -LineId $id -ETD '2026-03-03 13:00' -SortieType 'Day Local' -Mission "M$id" }
                                Remove-StudentReservation -Student $s -LineId @('P1', 'P3')
                                @($s.Reservations | ForEach-Object { $_.LineId }) | Should -Be @('P2')
                }

                It 'Get-ReservedMissionNames skips blank missions and the excepted line' {
                                $s = New-TestStudent -Name 'A'
                                Add-StudentReservation -Student $s -LineId 'P1' -ETD '2026-03-03 10:00' -SortieType 'O&B' -Mission 'FT5001'
                                Add-StudentReservation -Student $s -LineId 'R1' -ETD '2026-03-03 13:00' -SortieType 'O&B Return'
                                Add-StudentReservation -Student $s -LineId 'P2' -ETD '2026-03-03 16:00' -SortieType 'OFT' -Mission 'FE3102'
                                @(Get-ReservedMissionNames -Student $s) | Should -Be @('FT5001', 'FE3102')
                                @(Get-ReservedMissionNames -Student $s -ExceptLineId 'P2') | Should -Be @('FT5001')
                }
}

Describe 'Get-StudentTimeline' {
                It 'merges completed events and reservations oldest first, sim/flight by mission Type' {
                                # FE3102 flown with a named IP is still a sim (Type OFT)
                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00', 'FE3102|2026-03-03 06:00|Jones, Bob', 'FT5001|2026-02-23 08:00|Lee, Ann'
                                Add-StudentReservation -Student $s -LineId 'P1' -ETD '2026-03-03 13:00' -SortieType 'Day Local' -Mission 'FT5102' -Instructor 'Ortiz, Sam'
                                Add-StudentReservation -Student $s -LineId 'P2' -ETD '2026-02-21 13:00' -SortieType 'IFT' -Instructor 'SIM IP'
                                $tl = @(Get-StudentTimeline -Student $s)
                                @($tl | ForEach-Object { $_.Start.ToString('MM-dd HHmm') }) | Should -Be @('02-20 0800', '02-21 1300', '02-23 0800', '03-03 0600', '03-03 1300')
                                @($tl | ForEach-Object { $_.Mission }) | Should -Be @('FI3101', '', 'FT5001', 'FE3102', 'FT5102')
                                @($tl | ForEach-Object { $_.IsSim }) | Should -Be @($true, $true, $false, $true, $false)
                                @($tl | ForEach-Object { $_.IsReservation }) | Should -Be @($false, $true, $false, $false, $true)
                                $tl[3].Instructor | Should -Be 'Jones, Bob'
                                $tl[4].LineId | Should -Be 'P1'
                }

                It '-ExceptLineId leaves out that line and -NoReservations leaves out all of them' {
                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00'
                                Add-StudentReservation -Student $s -LineId 'P1' -ETD '2026-03-03 13:00' -SortieType 'Day Local' -Mission 'FT5001'
                                Add-StudentReservation -Student $s -LineId 'P2' -ETD '2026-03-03 17:00' -SortieType 'OFT' -Mission 'FE3102'
                                @(Get-StudentTimeline -Student $s -ExceptLineId 'P1' | ForEach-Object { $_.Mission }) | Should -Be @('FI3101', 'FE3102')
                                @(Get-StudentTimeline -Student $s -NoReservations | ForEach-Object { $_.Mission }) | Should -Be @('FI3101')
                }

                It 'skips completed events with no ETD and returns nothing for a null student' {
                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00'
                                $s.Syllabus[2].IsComplete = $true
                                $s.Syllabus[2].ETD = ''
                                @(Get-StudentTimeline -Student $s | ForEach-Object { $_.Mission }) | Should -Be @('FI3101')
                                @(Get-StudentTimeline -Student $null).Count | Should -Be 0
                }
}


# =================================================================================
#  TEST-HARDRULES
# =================================================================================
Describe 'Test-HardRules' {

                Context 'result shape and line-level rules' {
                                It 'returns Passed, Violations, Notes, PassedEvents and EventResults' {
                                                $s = New-TestStudent -Name 'Clean, Carl' -Done 'FI3101|2026-02-20 08:00'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -Instructor 'Jones, Bob' -Id 'L1')
                                                $r.Passed | Should -BeTrue
                                                @($r.Violations).Count | Should -Be 0
                                                @($r.Notes).Count | Should -Be 0
                                                Get-MissionNames $r.PassedEvents | Should -Be @('FT5001')
                                                $er = @($r.EventResults)
                                                $er.Count | Should -Be 1
                                                $er[0].Mission | Should -Be 'FT5001'
                                                $er[0].Passed | Should -BeTrue
                                                $er[0].Event.SyllabusID | Should -Be 3
                                                @($er[0].Violations).Count | Should -Be 0
                                }

                                It 'violations carry Rule, Severity, Message, Mission and Blocks' {
                                                $s = New-TestStudent -Name 'Held, Hank' -OnHold -Done 'FI3101|2026-02-20 08:00'
                                                $v = Get-Viol (Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -Instructor 'Jones, Bob')) 'ON_HOLD'
                                                $v.Severity | Should -Be 'Hard'
                                                $v.Message | Should -Be 'Student is on hold'
                                                $v.Mission | Should -Be ''
                                                $v.Blocks | Should -BeTrue
                                }

                                It 'BAD_ETD: a line with no valid ETD' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00'
                                                $row = New-TestRow -ETD '2026-03-03 08:00' -Instructor 'Jones, Bob'
                                                $row.ETD = 'sometime'
                                                $r = Invoke-HR -Student $s -Sortie $row -Explain
                                                Get-RuleCodes $r | Should -Be @('BAD_ETD')
                                                (Get-Viol $r 'BAD_ETD').Message | Should -Be "Line has no valid ETD ('sometime')"
                                                $r.Passed | Should -BeFalse
                                }

                                It 'RETURN_LINE: return legs are never filled on their own' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 13:00' -SortieType 'O&B Return' -Instructor 'Jones, Bob')
                                                Get-RuleCodes $r | Should -Be @('RETURN_LINE')
                                                (Get-Viol $r 'RETURN_LINE').Message | Should -Be 'O&B Return lines are filled from their O&B'
                                                $r.Passed | Should -BeFalse
                                }
                }

                Context 'student-level rules' {
                                It 'ON_HOLD' {
                                                $s = New-TestStudent -Name 'A' -OnHold -Done 'FI3101|2026-02-20 08:00'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -Instructor 'Jones, Bob')
                                                Get-RuleCodes $r | Should -Be @('ON_HOLD')
                                                (Get-Viol $r 'ON_HOLD').Message | Should -Be 'Student is on hold'
                                                $r.Passed | Should -BeFalse
                                }

                                It 'ABSENCE: <Name>' -TestCases @(
                                                @{ Name = 'overlap after the ETD (2.5 hr buffer)'; Start = '2026-03-03 12:00'; End = '2026-03-03 13:00'; Reason = 'Dentist'; Expected = 'Unavailable 1200-1300 (Dentist); sortie window 0830-1230' }
                                                @{ Name = 'overlap before the ETD (1.5 hr buffer), no reason'; Start = '2026-03-03 07:00'; End = '2026-03-03 08:31'; Reason = ''; Expected = 'Unavailable 0700-0831; sortie window 0830-1230' }
                                                @{ Name = 'multi-day absence shows dates'; Start = '2026-03-02 08:00'; End = '2026-03-04 17:00'; Reason = 'Leave'; Expected = 'Unavailable 02 Mar 0800-04 Mar 1700 (Leave); sortie window 0830-1230' }
                                                @{ Name = 'ends exactly at window start'; Start = '2026-03-03 07:00'; End = '2026-03-03 08:30'; Reason = 'Appt'; Expected = $null }
                                                @{ Name = 'starts exactly at window end'; Start = '2026-03-03 12:30'; End = '2026-03-03 14:00'; Reason = 'Appt'; Expected = $null }
                                                @{ Name = 'different day'; Start = '2026-03-04 09:00'; End = '2026-03-04 11:00'; Reason = 'Appt'; Expected = $null }
                                ) {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00' -Absences @(New-Absence $Start $End $Reason)
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 10:00' -Instructor 'Jones, Bob')
                                                if ($null -eq $Expected) {
                                                                Get-RuleCodes $r | Should -Not -Contain 'ABSENCE'
                                                                $r.Passed | Should -BeTrue
                                                }
                                                else {
                                                                Get-RuleCodes $r | Should -Be @('ABSENCE')
                                                                (Get-Viol $r 'ABSENCE').Message | Should -Be $Expected
                                                                $r.Passed | Should -BeFalse
                                                }
                                }

                                It 'ABSENCE reports only the first overlapping absence' {
                                                $abs = @((New-Absence '2026-03-03 09:00' '2026-03-03 09:30' 'One'), (New-Absence '2026-03-03 11:00' '2026-03-03 11:30' 'Two'))
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00' -Absences $abs
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 10:00' -Instructor 'Jones, Bob') -Explain
                                                @($r.Violations | Where-Object { $_.Rule -eq 'ABSENCE' }).Count | Should -Be 1
                                                (Get-Viol $r 'ABSENCE').Message | Should -Be 'Unavailable 0900-0930 (One); sortie window 0830-1230'
                                }

                                It 'CAP_OB' {
                                                $s = New-TestStudent -Name 'A' -IsCap -Done 'FI3101|2026-02-20 08:00'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -SortieType 'O&B' -Instructor 'Jones, Bob')
                                                Get-RuleCodes $r | Should -Be @('CAP_OB')
                                                (Get-Viol $r 'CAP_OB').Message | Should -Be "CAP students can't fly O&Bs"
                                }

                                It 'CAP_SINGLE_TURN: a CAP student who already flew today' {
                                                $s = New-TestStudent -Name 'A' -IsCap -Done 'FI3101|2026-02-20 08:00', 'FT5001|2026-03-03 07:00|Jones, Bob'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 13:00' -Instructor 'Jones, Bob')
                                                Get-RuleCodes $r | Should -Be @('CAP_SINGLE_TURN')
                                                (Get-Viol $r 'CAP_SINGLE_TURN').Message | Should -Be 'CAP student already flies today (0700 FT5001)'
                                }

                                It 'CAP_SINGLE_TURN counts a later pinned flight (reservation)' {
                                                $s = New-TestStudent -Name 'A' -IsCap -Done 'FI3101|2026-02-20 08:00'
                                                Add-StudentReservation -Student $s -LineId 'P1' -ETD '2026-03-03 15:00' -SortieType 'Day Local' -Instructor 'Jones, Bob'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -Instructor 'Jones, Bob' -Id 'L1')
                                                (Get-Viol $r 'CAP_SINGLE_TURN').Message | Should -Be 'CAP student already flies today (1500)'
                                }

                                It 'CAP_SINGLE_TURN ignores sims today and does not apply on a sim line' {
                                                # A sim flown with a named IP is still a sim (decided by mission Type)
                                                $s = New-TestStudent -Name 'A' -IsCap -Done 'FI3101|2026-02-20 08:00', 'FE3102|2026-03-03 07:00|Lee, Ann'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 13:00' -Instructor 'Jones, Bob')
                                                Get-RuleCodes $r | Should -Not -Contain 'CAP_SINGLE_TURN'
                                                $r.Passed | Should -BeTrue

                                                $s2 = New-TestStudent -Name 'B' -IsCap -Done 'FI3101|2026-02-20 08:00', 'FT5001|2026-03-03 07:00|Jones, Bob'
                                                $r2 = Invoke-HR -Student $s2 -Sortie (New-TestRow -ETD '2026-03-03 13:00' -SortieType 'OFT')
                                                Get-RuleCodes $r2 | Should -Not -Contain 'CAP_SINGLE_TURN'
                                                $r2.Passed | Should -BeTrue
                                }

                                It 'TRIPLE_TURN: two events already today' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-03-03 06:00', 'FT5001|2026-03-03 09:30|Jones, Bob'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 13:00' -Instructor 'Jones, Bob')
                                                Get-RuleCodes $r | Should -Be @('TRIPLE_TURN')
                                                (Get-Viol $r 'TRIPLE_TURN').Message | Should -Be 'Already has 2 events today (0600 FI3101, 0930 FT5001) - triple turn'
                                }

                                It 'TRIPLE_TURN counts a completed event plus a later reservation' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-03-03 06:00'
                                                Add-StudentReservation -Student $s -LineId 'P1' -ETD '2026-03-03 16:30' -SortieType 'OFT' -Mission 'FE3102' -Instructor 'SIM IP'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 11:00' -Instructor 'Jones, Bob' -Id 'L1')
                                                (Get-Viol $r 'TRIPLE_TURN').Message | Should -Be 'Already has 2 events today (0600 FI3101, 1630 FE3102) - triple turn'
                                }

                                It 'OB_NOT_ONLY_EVENT: an O&B must be the only event of the day' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-03-03 06:00'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 10:00' -SortieType 'O&B' -Instructor 'Jones, Bob')
                                                Get-RuleCodes $r | Should -Be @('OB_NOT_ONLY_EVENT')
                                                (Get-Viol $r 'OB_NOT_ONLY_EVENT').Message | Should -Be "O&B must be the day's only event (has 0600 FI3101)"
                                }

                                It 'OB_NOT_ONLY_EVENT does not apply to a Day Local with one event today' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-03-03 06:00'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 10:00' -Instructor 'Jones, Bob') -Explain
                                                @($r.Violations).Count | Should -Be 0
                                                $r.Passed | Should -BeTrue
                                }
                }

                Context 'turn time and duty day' {
                                BeforeAll {
                                                $script:TT = @('FI3101|2026-02-27 08:00', 'FT5001|2026-03-03 07:00|Jones, Bob')
                                }

                                It 'TURN_TIME: under 3.25 hr (plan example message)' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-23 08:00', 'FT5001|2026-02-24 08:00|Jones, Bob', 'FT5102|2026-02-25 08:00|Jones, Bob', 'FT5103|2026-03-03 07:30|Jones, Bob'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 10:15' -Instructor 'Jones, Bob')
                                                Get-RuleCodes $r | Should -Be @('TURN_TIME')
                                                (Get-Viol $r 'TURN_TIME').Message | Should -Be 'Turn time 2.8 hr < 3.25 hr minimum (vs 0730 FT5103)'
                                                $r.Passed | Should -BeFalse
                                }

                                It 'turn time <Etd>: rule <Rule>, note <Note>' -TestCases @(
                                                @{ Etd = '2026-03-03 09:30'; Rule = 'TURN_TIME'; Message = 'Turn time 2.5 hr < 3.25 hr minimum (vs 0700 FT5001)'; Note = $null }
                                                @{ Etd = '2026-03-03 10:15'; Rule = $null; Message = $null; Note = 'Less than 3.5 hr TT, Flt/CC Approval Required' }
                                                @{ Etd = '2026-03-03 10:29'; Rule = $null; Message = $null; Note = 'Less than 3.5 hr TT, Flt/CC Approval Required' }
                                                @{ Etd = '2026-03-03 10:30'; Rule = $null; Message = $null; Note = $null }
                                                @{ Etd = '2026-03-03 12:59'; Rule = $null; Message = $null; Note = $null }
                                                @{ Etd = '2026-03-03 13:00'; Rule = $null; Message = $null; Note = 'If this is a Night Flight SP will not have Duty Day' }
                                                @{ Etd = '2026-03-03 16:00'; Rule = $null; Message = $null; Note = 'If this is a Night Flight SP will not have Duty Day' }
                                                @{ Etd = '2026-03-03 16:30'; Rule = 'DUTY_DAY'; Message = 'Duty day: 9.5 hr between 0700 and 1630 events (> 9 hr max)'; Note = $null }
                                ) {
                                                $s = New-TestStudent -Name 'A' -Done $script:TT
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD $Etd -Instructor 'Jones, Bob') -Explain
                                                if ($null -eq $Rule) {
                                                                @($r.Violations).Count | Should -Be 0
                                                                $r.Passed | Should -BeTrue
                                                }
                                                else {
                                                                Get-RuleCodes $r | Should -Be @($Rule)
                                                                (Get-Viol $r $Rule).Message | Should -Be $Message
                                                                $r.Passed | Should -BeFalse
                                                }
                                                if ($null -eq $Note) { @($r.Notes).Count | Should -Be 0 }
                                                else { @($r.Notes) | Should -Be @($Note) }
                                }

                                It 'DUTY_DAY orders the times when the other event is later in the day' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-27 08:00'
                                                Add-StudentReservation -Student $s -LineId 'P1' -ETD '2026-03-03 16:30' -SortieType 'OFT' -Mission 'FE3102' -Instructor 'SIM IP'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 07:00' -Instructor 'Jones, Bob' -Id 'L1')
                                                (Get-Viol $r 'DUTY_DAY').Message | Should -Be 'Duty day: 9.5 hr between 0700 and 1630 events (> 9 hr max)'
                                }

                                It 'reservations count for turn time in both directions' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-27 08:00'
                                                Add-StudentReservation -Student $s -LineId 'P1' -ETD '2026-03-03 07:00' -SortieType 'OFT' -Mission 'FE3102' -Instructor 'SIM IP'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 09:00' -Instructor 'Jones, Bob' -Id 'L1')
                                                (Get-Viol $r 'TURN_TIME').Message | Should -Be 'Turn time 2.0 hr < 3.25 hr minimum (vs 0700 FE3102)'

                                                $s2 = New-TestStudent -Name 'B' -Done 'FI3101|2026-02-27 08:00'
                                                Add-StudentReservation -Student $s2 -LineId 'P2' -ETD '2026-03-03 11:00' -SortieType 'OFT' -Mission 'FE3102' -Instructor 'SIM IP'
                                                $r2 = Invoke-HR -Student $s2 -Sortie (New-TestRow -ETD '2026-03-03 09:00' -Instructor 'Jones, Bob' -Id 'L1')
                                                (Get-Viol $r2 'TURN_TIME').Message | Should -Be 'Turn time 2.0 hr < 3.25 hr minimum (vs 1100 FE3102)'
                                }

                                It 'a reservation with no mission is labelled by time only' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-27 08:00'
                                                Add-StudentReservation -Student $s -LineId 'R9' -ETD '2026-03-03 11:00' -SortieType 'O&B Return' -Instructor 'Ortiz, Sam'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 09:00' -Instructor 'Jones, Bob' -Id 'L1')
                                                (Get-Viol $r 'TURN_TIME').Message | Should -Be 'Turn time 2.0 hr < 3.25 hr minimum (vs 1100)'
                                }

                                It 'the reservation for the line being evaluated is ignored' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-27 08:00'
                                                Add-StudentReservation -Student $s -LineId 'L1' -ETD '2026-03-03 09:00' -SortieType 'Day Local' -Instructor 'Jones, Bob'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 09:00' -Instructor 'Jones, Bob' -Id 'L1') -Explain
                                                @($r.Violations).Count | Should -Be 0
                                                $r.Passed | Should -BeTrue
                                }

                                It 'notes survive a later violation on the fast path' {
                                                $s = New-TestStudent -Name 'A' -UnsatSim -Done $script:TT
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 10:15' -Instructor 'Jones, Bob' -Id 'L1')
                                                Get-RuleCodes $r | Should -Be @('UNSAT_SIM')
                                                @($r.Notes) | Should -Be @('Less than 3.5 hr TT, Flt/CC Approval Required')
                                }

                                It 'the same note is listed once even when two events trigger it' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-27 08:00', 'FT5001|2026-03-03 07:00|Jones, Bob'
                                                Add-StudentReservation -Student $s -LineId 'P1' -ETD '2026-03-03 13:30' -SortieType 'OFT' -Mission 'FE3102' -Instructor 'SIM IP'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 10:15' -Instructor 'Jones, Bob' -Id 'L1') -Explain
                                                Get-RuleCodes $r | Should -Be @('TRIPLE_TURN')
                                                @($r.Notes) | Should -Be @('Less than 3.5 hr TT, Flt/CC Approval Required')
                                }
                }

                Context 'O&B return leg' {
                                It 'RETURN_TURN_TIME: the linked return leg is too close to another event today' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-27 08:00'
                                                Add-StudentReservation -Student $s -LineId 'S1' -ETD '2026-03-03 16:00' -SortieType 'OFT' -Mission 'FE3102' -Instructor 'SIM IP'
                                                $ob = New-TestRow -ETD '2026-03-03 09:00' -SortieType 'O&B' -Instructor 'Jones, Bob' -Id 'OB1' -LinkedLineId 'R1'
                                                $rt = New-TestRow -ETD '2026-03-03 14:00' -SortieType 'O&B Return' -Instructor 'Jones, Bob' -Id 'R1'
                                                $r = Invoke-HR -Student $s -Sortie $ob -Schedule @($ob, $rt) -Explain
                                                Get-RuleCodes $r | Should -Be @('OB_NOT_ONLY_EVENT', 'RETURN_TURN_TIME')
                                                (Get-Viol $r 'OB_NOT_ONLY_EVENT').Message | Should -Be "O&B must be the day's only event (has 1600 FE3102)"
                                                (Get-Viol $r 'RETURN_TURN_TIME').Message | Should -Be 'Return leg 1400: turn time 2.0 hr < 3.25 hr (vs 1600 FE3102)'
                                                @($r.Notes) | Should -Be @('If this is a Night Flight SP will not have Duty Day')
                                }

                                It 'return leg notes and RETURN_DUTY_DAY: return at <ReturnEtd>' -TestCases @(
                                                @{ ReturnEtd = '2026-03-03 09:15'; Rule = $null; Message = $null; Note = 'Return Leg: Less than 3.5 hr TT, Flt/CC Approval Required' }
                                                @{ ReturnEtd = '2026-03-03 13:00'; Rule = $null; Message = $null; Note = 'Return Leg: Night Flight Warning' }
                                                @{ ReturnEtd = '2026-03-03 15:30'; Rule = 'RETURN_DUTY_DAY'; Message = 'Return leg 1530: duty day 9.5 hr (> 9 hr max, vs 0600 FI3101)'; Note = $null }
                                ) {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-03-03 06:00'
                                                $ob = New-TestRow -ETD '2026-03-03 10:00' -SortieType 'O&B' -Instructor 'Jones, Bob' -Id 'OB1' -LinkedLineId 'R1'
                                                $rt = New-TestRow -ETD $ReturnEtd -SortieType 'O&B Return' -Instructor 'Jones, Bob' -Id 'R1'
                                                $r = Invoke-HR -Student $s -Sortie $ob -Schedule @($ob, $rt) -Explain
                                                $expectedRules = @('OB_NOT_ONLY_EVENT')
                                                if ($null -ne $Rule) { $expectedRules += $Rule }
                                                Get-RuleCodes $r | Should -Be $expectedRules
                                                if ($null -ne $Rule) { (Get-Viol $r $Rule).Message | Should -Be $Message }
                                                if ($null -eq $Note) { @($r.Notes).Count | Should -Be 0 }
                                                else { @($r.Notes) | Should -Be @($Note) }
                                }

                                It 'finds an unlinked return line by the legacy lookup (same day, same IP, open)' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-27 08:00'
                                                Add-StudentReservation -Student $s -LineId 'S1' -ETD '2026-03-03 16:00' -SortieType 'OFT' -Mission 'FE3102' -Instructor 'SIM IP'
                                                $ob = New-TestRow -ETD '2026-03-03 09:00' -SortieType 'O&B' -Instructor 'Jones, Bob' -Id 'OB1'
                                                $other = New-TestRow -ETD '2026-03-03 12:00' -SortieType 'O&B Return' -Instructor 'Ortiz, Sam' -Id 'R0'
                                                $rt = New-TestRow -ETD '2026-03-03 14:00' -SortieType 'O&B Return' -Instructor 'Jones, Bob' -Id 'R1'
                                                $r = Invoke-HR -Student $s -Sortie $ob -Schedule @($ob, $other, $rt) -Explain
                                                (Get-Viol $r 'RETURN_TURN_TIME').Message | Should -Be 'Return leg 1400: turn time 2.0 hr < 3.25 hr (vs 1600 FE3102)'
                                }
                }

                Context 'crew rest' {
                                It 'CREW_REST: under 15 hr since the previous day''s last event' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-27 08:00', 'FT5001|2026-03-02 07:00|Jones, Bob', 'FT5102|2026-03-02 18:30|Jones, Bob'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 07:30' -Instructor 'Jones, Bob')
                                                Get-RuleCodes $r | Should -Be @('CREW_REST')
                                                (Get-Viol $r 'CREW_REST').Message | Should -Be 'Crew rest 13.0 hr < 15 hr since Mon 1830 FT5102'
                                                $r.Passed | Should -BeFalse
                                }

                                It 'CREW_REST: exactly 15 hr is legal' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-27 08:00', 'FT5001|2026-03-02 07:00|Jones, Bob', 'FT5102|2026-03-02 18:30|Jones, Bob'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 09:30' -Instructor 'Jones, Bob') -Explain
                                                @($r.Violations).Count | Should -Be 0
                                                $r.Passed | Should -BeTrue
                                }

                                It 'CREW_REST only looks at the previous calendar day' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-27 08:00', 'FT5001|2026-03-01 23:00|Jones, Bob'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 06:00' -Instructor 'Jones, Bob') -Explain
                                                Get-RuleCodes $r | Should -Not -Contain 'CREW_REST'
                                }

                                It 'CREW_REST counts a previous-day reservation' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-27 08:00'
                                                Add-StudentReservation -Student $s -LineId 'P0' -ETD '2026-03-02 20:00' -SortieType 'OFT' -Mission 'FE3102' -Instructor 'SIM IP'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 07:30' -Instructor 'Jones, Bob' -Id 'L1')
                                                (Get-Viol $r 'CREW_REST').Message | Should -Be 'Crew rest 11.5 hr < 15 hr since Mon 2000 FE3102'
                                }

                                It 'CREW_REST_NEXT: under 15 hr before the next day''s first event (reservation)' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-27 08:00', 'FT5001|2026-02-27 13:00|Jones, Bob'
                                                Add-StudentReservation -Student $s -LineId 'W1' -ETD '2026-03-04 06:00' -SortieType 'OFT' -Mission 'FE3102' -Instructor 'SIM IP'
                                                Add-StudentReservation -Student $s -LineId 'W2' -ETD '2026-03-04 12:00' -SortieType 'Day Local' -Instructor 'Lee, Ann'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 18:00' -Instructor 'Jones, Bob' -Id 'L1')
                                                Get-RuleCodes $r | Should -Be @('CREW_REST_NEXT')
                                                (Get-Viol $r 'CREW_REST_NEXT').Message | Should -Be 'Crew rest 12.0 hr < 15 hr before Wed 0600 FE3102'
                                }

                                It 'CREW_REST_NEXT: a completed event on the next day counts too' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-27 08:00', 'FT5001|2026-03-04 06:00|Jones, Bob'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 18:00' -Instructor 'Jones, Bob')
                                                (Get-Viol $r 'CREW_REST_NEXT').Message | Should -Be 'Crew rest 12.0 hr < 15 hr before Wed 0600 FT5001'
                                }

                                It 'CREW_REST_NEXT measures from the O&B return leg when there is one' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-27 08:00'
                                                Add-StudentReservation -Student $s -LineId 'W1' -ETD '2026-03-04 06:00' -SortieType 'OFT' -Mission 'FE3102' -Instructor 'SIM IP'
                                                $ob = New-TestRow -ETD '2026-03-03 13:00' -SortieType 'O&B' -Instructor 'Jones, Bob' -Id 'OB1' -LinkedLineId 'R1'
                                                $rt = New-TestRow -ETD '2026-03-03 16:00' -SortieType 'O&B Return' -Instructor 'Jones, Bob' -Id 'R1'
                                                $r = Invoke-HR -Student $s -Sortie $ob -Schedule @($ob, $rt) -Explain
                                                Get-RuleCodes $r | Should -Be @('CREW_REST_NEXT')
                                                (Get-Viol $r 'CREW_REST_NEXT').Message | Should -Be 'Crew rest 14.0 hr < 15 hr before Wed 0600 FE3102'

                                                # Without the return leg the O&B itself is 17 hr out
                                                $r2 = Invoke-HR -Student $s -Sortie $ob -Schedule @($ob) -Explain
                                                Get-RuleCodes $r2 | Should -Not -Contain 'CREW_REST_NEXT'
                                }

                                It 'CREW_REST_NEXT ignores events two days out' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-27 08:00'
                                                Add-StudentReservation -Student $s -LineId 'T1' -ETD '2026-03-05 05:00' -SortieType 'OFT' -Mission 'FE3102' -Instructor 'SIM IP'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 18:00' -Instructor 'Jones, Bob' -Id 'L1') -Explain
                                                @($r.Violations).Count | Should -Be 0
                                }
                }

                Context 'experienced IPs and unsat events' {
                                It 'EXP_IP: <Who> student with a non-experienced IP' -TestCases @(
                                                @{ Who = 'CAP'; Cap = $true; Intl = $false; Ip = 'Lee, Ann'; Expected = "CAP student needs an experienced IP (Lee, Ann isn't)" }
                                                @{ Who = 'International'; Cap = $false; Intl = $true; Ip = 'Lee, Ann'; Expected = "International student needs an experienced IP (Lee, Ann isn't)" }
                                                @{ Who = 'CAP+International'; Cap = $true; Intl = $true; Ip = 'Lee, Ann'; Expected = "CAP student needs an experienced IP (Lee, Ann isn't)" }
                                                @{ Who = 'CAP (no IP)'; Cap = $true; Intl = $false; Ip = ''; Expected = "CAP student needs an experienced IP (NEED IP isn't)" }
                                ) {
                                                $s = New-TestStudent -Name 'A' -IsCap:$Cap -IsInternational:$Intl -Done 'FI3101|2026-02-20 08:00'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -Instructor $Ip)
                                                Get-RuleCodes $r | Should -Be @('EXP_IP')
                                                (Get-Viol $r 'EXP_IP').Message | Should -Be $Expected
                                }

                                It 'EXP_IP also applies on an O&B' {
                                                $s = New-TestStudent -Name 'A' -IsInternational -Done 'FI3101|2026-02-20 08:00'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -SortieType 'O&B' -Instructor 'Lee, Ann')
                                                (Get-Viol $r 'EXP_IP').Message | Should -Be "International student needs an experienced IP (Lee, Ann isn't)"
                                }

                                It 'EXP_IP passes with an experienced IP and does not apply on sims' {
                                                $s = New-TestStudent -Name 'A' -IsCap -IsInternational -Done 'FI3101|2026-02-20 08:00'
                                                (Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -Instructor 'Jones, Bob')).Passed | Should -BeTrue
                                                (Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -SortieType 'OFT')).Passed | Should -BeTrue
                                }

                                It 'UNSAT_SIM blocks <SortieType> but not sims' -TestCases @(
                                                @{ SortieType = 'Day Local' }
                                                @{ SortieType = 'O&B' }
                                ) {
                                                $s = New-TestStudent -Name 'A' -UnsatSim -Done 'FI3101|2026-02-20 08:00'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -SortieType $SortieType -Instructor 'Jones, Bob')
                                                Get-RuleCodes $r | Should -Be @('UNSAT_SIM')
                                                (Get-Viol $r 'UNSAT_SIM').Message | Should -Be 'Unsat sim on record - no Day Local/O&B until a re-sim in an OFT'
                                                (Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -SortieType 'OFT')).Passed | Should -BeTrue
                                }

                                It 'UNSAT_SIM names the sim that cleans it up (<Kind>)' -TestCases @(
                                                @{ Kind = 'OFT'; Expected = 'Unsat sim on record - no Day Local/O&B until a re-sim in an OFT' }
                                                @{ Kind = 'IFT'; Expected = 'Unsat sim on record - no Day Local/O&B until a re-sim in an OFT or IFT' }
                                                @{ Kind = ''; Expected = 'Unsat sim on record - no Day Local/O&B until a re-sim in an OFT' }
                                ) {
                                                $s = New-TestStudent -Name 'A' -UnsatSim -UnsatSimType $Kind -Done 'FI3101|2026-02-20 08:00'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -Instructor 'Jones, Bob')
                                                (Get-Viol $r 'UNSAT_SIM').Message | Should -Be $Expected
                                }

                                It 'UNSAT_FLIGHT blocks <SortieType> but not flights' -TestCases @(
                                                @{ SortieType = 'OFT' }
                                                @{ SortieType = 'IFT' }
                                ) {
                                                $s = New-TestStudent -Name 'A' -UnsatFlight -Done 'FI3101|2026-02-20 08:00', 'FT5001|2026-02-23 08:00|Jones, Bob', 'FT5102|2026-02-24 08:00|Jones, Bob'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -SortieType $SortieType)
                                                Get-RuleCodes $r | Should -Be @('UNSAT_FLIGHT')
                                                (Get-Viol $r 'UNSAT_FLIGHT').Message | Should -Be 'Unsat flight on record - no OFT/IFT until re-fly'
                                                (Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -Instructor 'Jones, Bob')).Passed | Should -BeTrue
                                }
                }

                Context 'eligible events and explicit missions' {
                                It 'NO_ELIGIBLE_EVENT: next events do not fit the line' {
                                                $s = New-TestStudent -Name 'A'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -Instructor 'Jones, Bob')
                                                Get-RuleCodes $r | Should -Be @('NO_ELIGIBLE_EVENT')
                                                (Get-Viol $r 'NO_ELIGIBLE_EVENT').Message | Should -Be 'No eligible Day Local events (next: FI3101 IFT)'
                                                @($r.EventResults).Count | Should -Be 0
                                                $r.Passed | Should -BeFalse
                                }

                                It 'NO_ELIGIBLE_EVENT lists up to three next events' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00', 'FT5001|2026-02-23 08:00|Jones, Bob', 'FT5102|2026-02-24 08:00|Jones, Bob'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -SortieType 'O&B' -Instructor 'Jones, Bob')
                                                (Get-Viol $r 'NO_ELIGIBLE_EVENT').Message | Should -Be 'No eligible O&B events (next: FE3102 OFT, FT5103 SingleSolo, FN3201 IFT)'
                                }

                                It 'NO_ELIGIBLE_EVENT: O&B with no second leg' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00', 'FT5001|2026-02-23 08:00|Jones, Bob'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -SortieType 'O&B' -Instructor 'Jones, Bob')
                                                (Get-Viol $r 'NO_ELIGIBLE_EVENT').Message | Should -Be 'No eligible O&B events (no second O&B leg available after FT5102)'
                                }

                                It 'NO_ELIGIBLE_EVENT: syllabus complete or waiting on reserved events' {
                                                $done = @(New-TestSyllabus | ForEach-Object { "$($_.Mission)|2026-02-2$($_.SyllabusID % 7) 08:00|Lee, Ann" })
                                                $s = New-TestStudent -Name 'A' -Done $done
                                                (Get-Viol (Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -Instructor 'Jones, Bob')) 'NO_ELIGIBLE_EVENT').Message |
                                                                Should -Be 'No eligible Day Local events (syllabus complete or waiting on reserved events)'

                                                $s2 = New-TestStudent -Name 'B' -Done 'FI3101|2026-02-20 08:00'
                                                Add-StudentReservation -Student $s2 -LineId 'P1' -ETD '2026-03-05 08:00' -SortieType 'Day Local' -Mission 'FT5001' -Instructor 'Jones, Bob'
                                                Add-StudentReservation -Student $s2 -LineId 'P2' -ETD '2026-03-05 13:00' -SortieType 'OFT' -Mission 'FE3102' -Instructor 'SIM IP'
                                                (Get-Viol (Invoke-HR -Student $s2 -Sortie (New-TestRow -ETD '2026-03-03 08:00' -Instructor 'Jones, Bob' -Id 'L1')) 'NO_ELIGIBLE_EVENT').Message |
                                                                Should -Be 'No eligible Day Local events (syllabus complete or waiting on reserved events)'
                                }

                                It 'MISSION_NOT_FOUND' {
                                                $s = New-TestStudent -Name 'Smith, John' -Done 'FI3101|2026-02-20 08:00'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -Instructor 'Jones, Bob') -Mission 'ZZ9999'
                                                Get-RuleCodes $r | Should -Be @('MISSION_NOT_FOUND')
                                                $v = Get-Viol $r 'MISSION_NOT_FOUND'
                                                $v.Message | Should -Be "ZZ9999 isn't in Smith, John's syllabus"
                                                $v.Mission | Should -Be 'ZZ9999'
                                                @($r.EventResults).Count | Should -Be 0
                                                $r.Passed | Should -BeFalse
                                }

                                It 'NOT_OPTED (<Why>)' -TestCases @(
                                                @{ Why = 'already complete'; Done = @('FI3101|2026-02-20 08:00'); Type = 'IFT'; Ip = ''; Mission = 'FI3101'; Expected = 'Not opted for FI3101: already complete' }
                                                @{ Why = 'prereq'; Done = @(); Type = 'Day Local'; Ip = 'Jones, Bob'; Mission = 'FT5102'; Expected = "Not opted for FT5102: FT5001 isn't complete yet" }
                                                @{ Why = 'sim before its prereq flight'; Done = @('FI3101|2026-02-20 08:00', 'FT5001|2026-02-23 08:00|Jones, Bob'); Type = 'OFT'; Ip = ''; Mission = 'FN3201'; Expected = "Not opted for FN3201: FT5102 isn't complete yet" }
                                ) {
                                                $s = New-TestStudent -Name 'A' -Done $Done
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -SortieType $Type -Instructor $Ip) -Mission $Mission -Explain
                                                @($r.EventResults).Count | Should -Be 1
                                                $v = Get-EvtViol $r $Mission 'NOT_OPTED'
                                                $v.Message | Should -Be $Expected
                                                $v.Mission | Should -Be $Mission
                                                $v.Severity | Should -Be 'Hard'
                                                $v.Blocks | Should -BeTrue
                                                @((Get-EvtResult $r $Mission).Violations | Where-Object { $_.Rule -eq 'MISSION_NOT_ELIGIBLE' }).Count | Should -Be 0
                                                (Get-EvtResult $r $Mission).Passed | Should -BeFalse
                                                $r.Passed | Should -BeFalse
                                }

                                It 'NOT_OPTED: a sim scheduled before the dollar ride (the sim needs FT5001)' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00'
                                                ($s.Syllabus | Where-Object { $_.Mission -eq 'FE3102' }).Prereqs = 'FI3101, FT5001'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -SortieType 'OFT') -Mission 'FE3102' -Explain
                                                (Get-EvtViol $r 'FE3102' 'NOT_OPTED').Message | Should -Be "Not opted for FE3102: FT5001 isn't complete yet"
                                                $r.Passed | Should -BeFalse
                                }

                                It 'NOT_OPTED lists every missing prereq' {
                                                $s = New-TestStudent -Name 'A'
                                                ($s.Syllabus | Where-Object { $_.Mission -eq 'FT5190' }).Prereqs = 'FT5104, FN3201'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -Instructor 'Jones, Bob') -Mission 'FT5190' -Explain
                                                (Get-EvtViol $r 'FT5190' 'NOT_OPTED').Message | Should -Be "Not opted for FT5190: FT5104, FN3201 aren't complete yet"
                                }

                                It 'a pinned line with a mission the student is not opted for is kept and flagged OVERRIDE' {
                                                $s = @(New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00')
                                                $row = New-TestRow -ETD '2026-03-03 08:00' -Instructor 'Jones, Bob' -Id 'P1' -Student 'A' -Mission 'FT5102'
                                                $fill = Invoke-FillSchedule -Schedule @($row) -Students $s -Instructors (New-TestInstructors) -SoftWeights (New-TestSoftWeights) 6>$null 3>$null
                                                $out = $fill.Schedule[0]
                                                (Get-NameText $out.Mission) | Should -Be 'FT5102'
                                                $out.Notes | Should -BeLike "*OVERRIDE: Not opted for FT5102: FT5001 isn't complete yet*"
                                                @($fill.Overrides).Count | Should -Be 1
                                }

                                It 'a typed mission that is not in the syllabus is kept as typed and flagged once' {
                                                $s = @(New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00')
                                                $row = New-TestRow -ETD '2026-03-03 08:00' -Instructor 'Jones, Bob' -Id 'P1' -Student 'A' -Mission 'XX1234'
                                                $fill = Invoke-FillSchedule -Schedule @($row) -Students $s -Instructors (New-TestInstructors) -SoftWeights (New-TestSoftWeights) 6>$null 3>$null
                                                $out = $fill.Schedule[0]
                                                (Get-NameText $out.Mission) | Should -Be 'XX1234'
                                                $out.Notes | Should -BeLike "*OVERRIDE: XX1234 isn't in A's syllabus*"
                                                $out.Notes | Should -Not -BeLike '*mission left blank*'
                                }

                                It 'MISSION_NOT_ELIGIBLE (<Why>)' -TestCases @(
                                                @{ Why = 'type'; Done = @('FI3101|2026-02-20 08:00'); Type = 'Day Local'; Ip = 'Jones, Bob'; Mission = 'FE3102'; Expected = "FE3102 isn't eligible here (Type OFT not allowed on Day Local)" }
                                                @{ Why = 'no second O&B leg'; Done = @('FI3101|2026-02-20 08:00', 'FT5001|2026-02-23 08:00|Jones, Bob'); Type = 'O&B'; Ip = 'Jones, Bob'; Mission = 'FT5102'; Expected = "FT5102 isn't eligible here (no second O&B leg available)" }
                                ) {
                                                $s = New-TestStudent -Name 'A' -Done $Done
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -SortieType $Type -Instructor $Ip) -Mission $Mission -Explain
                                                @($r.EventResults).Count | Should -Be 1
                                                $v = Get-EvtViol $r $Mission 'MISSION_NOT_ELIGIBLE'
                                                $v.Message | Should -Be $Expected
                                                $v.Mission | Should -Be $Mission
                                                $v.Severity | Should -Be 'Hard'
                                                (Get-EvtResult $r $Mission).Passed | Should -BeFalse
                                                $r.Passed | Should -BeFalse
                                }

                                It 'MISSION_NOT_ELIGIBLE (reserved on another line), but fine on its own line' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00'
                                                Add-StudentReservation -Student $s -LineId 'P9' -ETD '2026-03-05 10:00' -SortieType 'Day Local' -Mission 'FT5001' -Instructor 'Jones, Bob'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -Instructor 'Jones, Bob' -Id 'L1') -Mission 'FT5001'
                                                (Get-EvtViol $r 'FT5001' 'MISSION_NOT_ELIGIBLE').Message | Should -Be "FT5001 isn't eligible here (already pinned on another line)"
                                                $r.Passed | Should -BeFalse

                                                $own = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -Instructor 'Jones, Bob' -Id 'P9') -Mission 'FT5001'
                                                $own.Passed | Should -BeTrue
                                                Get-MissionNames $own.PassedEvents | Should -Be @('FT5001')
                                }

                                It 'an eligible explicit mission is the only event tested' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00', 'FT5001|2026-02-23 08:00|Jones, Bob', 'FT5102|2026-02-24 08:00|Jones, Bob'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -SortieType 'OFT') -Mission 'FN3201'
                                                @($r.EventResults | ForEach-Object { $_.Mission }) | Should -Be @('FN3201')
                                                $r.Passed | Should -BeTrue
                                }
                }

                Context 'O&B and checkride event rules' {
                                It 'FORM_ON_OB: formation events cannot go on an O&B' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00', 'FT5001|2026-02-23 08:00|Jones, Bob', 'FT5102|2026-02-24 08:00|Jones, Bob', 'FN3201|2026-02-24 13:00'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -SortieType 'O&B' -Instructor 'Jones, Bob' -Id 'OB1')
                                                @($r.Violations).Count | Should -Be 0
                                                $v = Get-EvtViol $r 'FF5201' 'FORM_ON_OB'
                                                $v.Message | Should -Be "Formation events can't be flown on an O&B"
                                                $v.Mission | Should -Be 'FF5201'
                                                @((Get-EvtResult $r 'FF5201').Violations | ForEach-Object { $_.Rule }) | Should -Be @('FORM_ON_OB')
                                                $r.Passed | Should -BeFalse
                                }

                                It 'CHECK_ON_OB and OB_NO_SECOND_LEG for a checkride pinned on an O&B' {
                                                $s = New-TestStudent -Name 'A' -Done $script:AllButCheck
                                                # Ortiz flew FT5103 with this student, so he is also not independent
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -SortieType 'O&B' -Instructor 'Ortiz, Sam') -Mission 'FT5190' -Explain
                                                $er = Get-EvtResult $r 'FT5190'
                                                @($er.Violations | ForEach-Object { $_.Rule }) | Should -Be @('MISSION_NOT_ELIGIBLE', 'CHECK_ON_OB', 'OB_NO_SECOND_LEG', 'CHECK_SAME_IP', 'CHECK_NEEDS_CHECK_IP')
                                                (Get-EvtViol $r 'FT5190' 'MISSION_NOT_ELIGIBLE').Message | Should -Be "FT5190 isn't eligible here (no second O&B leg available)"
                                                (Get-EvtViol $r 'FT5190' 'CHECK_ON_OB').Message | Should -Be "Checkrides can't be flown on an O&B"
                                                (Get-EvtViol $r 'FT5190' 'OB_NO_SECOND_LEG').Message | Should -Be 'No flyable return-leg event after FT5190'
                                                (Get-EvtViol $r 'FT5190' 'CHECK_SAME_IP').Message | Should -Be 'Evaluator Ortiz, Sam already flew with this student (FT5103 25 Feb)'
                                                (Get-EvtViol $r 'FT5190' 'CHECK_NEEDS_CHECK_IP').Message | Should -Be "Checkride needs a Check IP (Ortiz, Sam isn't)"
                                                @($r.Violations).Count | Should -Be 0
                                                $r.Passed | Should -BeFalse
                                }

                                It 'OB_NO_SECOND_LEG when the only follow-on is a solo' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00', 'FT5001|2026-02-23 08:00|Jones, Bob'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -SortieType 'O&B' -Instructor 'Jones, Bob') -Mission 'FT5102' -Explain
                                                (Get-EvtViol $r 'FT5102' 'OB_NO_SECOND_LEG').Message | Should -Be 'No flyable return-leg event after FT5102'
                                }

                                It 'OB_RETURN_CHECK: the return leg would be a checkride' {
                                                $done = @($script:AllButCheck | Where-Object { $_ -notlike 'FT5104*' })
                                                $s = New-TestStudent -Name 'A' -Done $done
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -SortieType 'O&B' -Instructor 'Jones, Bob')
                                                @($r.Violations).Count | Should -Be 0
                                                $v = Get-EvtViol $r 'FT5104' 'OB_RETURN_CHECK'
                                                $v.Message | Should -Be 'Return leg would be checkride FT5190'
                                                $v.Mission | Should -Be 'FT5104'
                                                $r.Passed | Should -BeFalse
                                }

                                It 'CHECK_SAME_IP: the evaluator already flew with the student' {
                                                $s = New-TestStudent -Name 'A' -Done $script:AllButCheck
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -Instructor 'Jones, Bob')
                                                @((Get-EvtResult $r 'FT5190').Violations | ForEach-Object { $_.Rule }) | Should -Be @('CHECK_SAME_IP')
                                                (Get-EvtViol $r 'FT5190' 'CHECK_SAME_IP').Message | Should -Be 'Evaluator Jones, Bob already flew with this student (FT5104 27 Feb)'
                                                $r.Passed | Should -BeFalse
                                }

                                It 'CHECK_SAME_IP ignores sims flown with that IP' {
                                                $done = @($script:AllButCheck | ForEach-Object { if ($_ -like 'FT5104*') { 'FT5104|2026-02-27 08:00|Lee, Ann' } elseif ($_ -like 'FE3102*') { 'FE3102|2026-02-20 13:00|Jones, Bob' } else { $_ } })
                                                $s = New-TestStudent -Name 'A' -Done $done
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -Instructor 'Jones, Bob') -Explain
                                                $r.Passed | Should -BeTrue
                                                @((Get-EvtResult $r 'FT5190').Violations).Count | Should -Be 0
                                }

                                It 'CHECK_NEEDS_CHECK_IP' {
                                                $s = New-TestStudent -Name 'A' -Done $script:AllButCheck
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -Instructor 'Leech, Tom')
                                                @((Get-EvtResult $r 'FT5190').Violations | ForEach-Object { $_.Rule }) | Should -Be @('CHECK_NEEDS_CHECK_IP')
                                                (Get-EvtViol $r 'FT5190' 'CHECK_NEEDS_CHECK_IP').Message | Should -Be "Checkride needs a Check IP (Leech, Tom isn't)"
                                }
                }

                Context 'formation lead block' {
                                BeforeAll {
                                                $script:LeadRow = New-TestRow -ETD '2026-03-03 08:00' -Instructor 'Jones, Bob' -Id 'L1'
                                                $script:OpenRow = New-TestRow -ETD '2026-03-03 08:10' -Instructor 'Ortiz, Sam' -Id 'L2'
                                }

                                It 'FORM_NO_WING_LINE when no open live-fly line is inside the look-ahead' {
                                                $lead = New-TestStudent -Name 'Lead' -Done $script:FormDone
                                                $late = New-TestRow -ETD '2026-03-03 08:30' -Instructor 'Ortiz, Sam' -Id 'L3'
                                                $r = Invoke-HR -Student $lead -Sortie $script:LeadRow -Schedule @($script:LeadRow, $late) -Students @($lead)
                                                @($r.Violations).Count | Should -Be 0
                                                $v = Get-EvtViol $r 'FF5201' 'FORM_NO_WING_LINE'
                                                $v.Message | Should -Be 'No open Day Local line within 22 min for a wingman/DS'
                                                $v.Severity | Should -Be 'Hard'
                                                $r.Passed | Should -BeFalse

                                                # A longer look-ahead reaches the 0830 line (message reports the window)
                                                $r30 = Invoke-HR -Student $lead -Sortie $script:LeadRow -Schedule @($script:LeadRow, $late) -Students @($lead) -FormLookAhead 30
                                                $r30.Passed | Should -BeTrue
                                                $r15 = Invoke-HR -Student $lead -Sortie $script:LeadRow -Schedule @($script:LeadRow) -Students @($lead) -FormLookAhead 15
                                                (Get-EvtViol $r15 'FF5201' 'FORM_NO_WING_LINE').Message | Should -Be 'No open Day Local line within 15 min for a wingman/DS'
                                }

                                It 'lines that cannot take a wingman do not count: <Why>' -TestCases @(
                                                @{ Why = 'O&B'; SortieType = 'O&B'; Student = $null; FilledBy = $null }
                                                @{ Why = 'O&B Return'; SortieType = 'O&B Return'; Student = $null; FilledBy = $null }
                                                @{ Why = 'IFT'; SortieType = 'IFT'; Student = $null; FilledBy = $null }
                                                @{ Why = 'OFT'; SortieType = 'OFT'; Student = $null; FilledBy = $null }
                                                @{ Why = 'has a student'; SortieType = 'Day Local'; Student = 'Someone'; FilledBy = $null }
                                                @{ Why = 'already filled'; SortieType = 'Day Local'; Student = $null; FilledBy = 'DS' }
                                ) {
                                                $lead = New-TestStudent -Name 'Lead' -Done $script:FormDone
                                                $row = New-TestRow -ETD '2026-03-03 08:10' -SortieType $SortieType -Instructor 'Ortiz, Sam' -Id 'L2' -Student $Student
                                                $row.FilledBy = $FilledBy
                                                $r = Invoke-HR -Student $lead -Sortie $script:LeadRow -Schedule @($script:LeadRow, $row) -Students @($lead)
                                                (Get-EvtViol $r 'FF5201' 'FORM_NO_WING_LINE') | Should -Not -BeNullOrEmpty
                                }

                                It 'FORM_PREFER_SINGLE (Advisory, still blocks) when only single-ship students are available' {
                                                $lead = New-TestStudent -Name 'Lead' -Done $script:FormDone
                                                $single = New-TestStudent -Name 'Single' -Done 'FI3101|2026-02-20 08:00'
                                                $r = Invoke-HR -Student $lead -Sortie $script:LeadRow -Schedule @($script:LeadRow, $script:OpenRow) -Students @($lead, $single)
                                                $v = Get-EvtViol $r 'FF5201' 'FORM_PREFER_SINGLE'
                                                $v.Message | Should -Be 'No formation wingman available; another student can fly single-ship here'
                                                $v.Severity | Should -Be 'Advisory'
                                                $v.Blocks | Should -BeTrue
                                                $v.Mission | Should -Be 'FF5201'
                                                (Get-EvtResult $r 'FF5201').Passed | Should -BeFalse
                                                $r.Passed | Should -BeFalse
                                }

                                It 'Direct Support Required is an event note (not a Sortie note) when nobody else can fly' {
                                                $lead = New-TestStudent -Name 'Lead' -Done $script:FormDone
                                                $newbie = New-TestStudent -Name 'Newbie'
                                                $r = Invoke-HR -Student $lead -Sortie $script:LeadRow -Schedule @($script:LeadRow, $script:OpenRow) -Students @($lead, $newbie)
                                                $r.Passed | Should -BeTrue
                                                $er = Get-EvtResult $r 'FF5201'
                                                @($er.Violations).Count | Should -Be 0
                                                @($er.Notes) | Should -Be @('Direct Support Required')
                                                @($r.Notes).Count | Should -Be 0
                                                $script:LeadRow.Notes | Should -BeNullOrEmpty
                                }

                                It 'passes cleanly when another student can fly Form' {
                                                $lead = New-TestStudent -Name 'Lead' -Done $script:FormDone
                                                $wing = New-TestStudent -Name 'Wing' -Done $script:FormDone
                                                $single = New-TestStudent -Name 'Single' -Done 'FI3101|2026-02-20 08:00'
                                                $r = Invoke-HR -Student $lead -Sortie $script:LeadRow -Schedule @($script:LeadRow, $script:OpenRow) -Students @($lead, $wing, $single)
                                                $r.Passed | Should -BeTrue
                                                $er = Get-EvtResult $r 'FF5201'
                                                @($er.Violations).Count | Should -Be 0
                                                @($er.Notes).Count | Should -Be 0
                                }

                                It '-IsWingCandidate skips the formation lead block' {
                                                $lead = New-TestStudent -Name 'Lead' -Done $script:FormDone
                                                $r = Invoke-HR -Student $lead -Sortie $script:LeadRow -Schedule @($script:LeadRow) -Students @($lead) -IsWingCandidate -Explain
                                                $r.Passed | Should -BeTrue
                                                @((Get-EvtResult $r 'FF5201').Violations).Count | Should -Be 0
                                }

                                It 'caches the formation context in -LineContext and reuses it' {
                                                $lead = New-TestStudent -Name 'Lead' -Done $script:FormDone
                                                $wing = New-TestStudent -Name 'Wing' -Done $script:FormDone
                                                $ctx = @{}
                                                $r = Invoke-HR -Student $lead -Sortie $script:LeadRow -Schedule @($script:LeadRow, $script:OpenRow) -Students @($lead, $wing) -LineContext $ctx
                                                $r.Passed | Should -BeTrue
                                                $ctx.ContainsKey('Formation') | Should -BeTrue
                                                @($ctx['Formation'].OpenWingLines | ForEach-Object { $_.BoardLineId }) | Should -Be @('L2')
                                                $ctx['Formation'].FormCapable.Contains('Wing') | Should -BeTrue

                                                # A pre-filled context wins over the live schedule
                                                $fake = @{ Formation = [PSCustomObject]@{ OpenWingLines = @(); FormCapable = @(); SingleCapable = @() } }
                                                $r2 = Invoke-HR -Student $lead -Sortie $script:LeadRow -Schedule @($script:LeadRow, $script:OpenRow) -Students @($lead, $wing) -LineContext $fake
                                                (Get-EvtViol $r2 'FF5201' 'FORM_NO_WING_LINE') | Should -Not -BeNullOrEmpty
                                }
                }

                Context 'solo flight notes' {
                                BeforeAll {
                                                $script:SoloDone = @('FI3101|2026-02-23 08:00', 'FT5001|2026-02-24 08:00|Jones, Bob', 'FT5102|2026-02-25 08:00|Jones, Bob')
                                }

                                It 'IP not in the last five flights: <Ip>' -TestCases @(
                                                @{ Ip = 'Lee, Ann'; Expected = @('Solo Flight', 'Need IP in Continuity to Brief') }
                                                @{ Ip = 'Jones, Bob'; Expected = @('Solo Flight') }
                                                @{ Ip = ''; Expected = @('Solo Flight') }
                                ) {
                                                $s = New-TestStudent -Name 'A' -Done $script:SoloDone
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -Instructor $Ip)
                                                $r.Passed | Should -BeTrue
                                                $er = Get-EvtResult $r 'FT5103'
                                                $er.Passed | Should -BeTrue
                                                @($er.Notes) | Should -Be $Expected
                                                @($r.Notes).Count | Should -Be 0
                                }

                                It 'only counts flights before this sortie for continuity' {
                                                # Lee flies the student later (completed on the board) - not a brief-continuity flight
                                                $s = New-TestStudent -Name 'A' -Done ($script:SoloDone + 'FT5104|2026-03-05 08:00|Lee, Ann')
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -Instructor 'Lee, Ann')
                                                @((Get-EvtResult $r 'FT5103').Notes) | Should -Be @('Solo Flight', 'Need IP in Continuity to Brief')
                                }
                }

                Context 'fast path vs -Explain' {
                                It 'the fast path stops at the first violation; -Explain returns them all' {
                                                $abs = @(New-Absence '2026-03-03 09:00' '2026-03-03 09:30' 'Appt')
                                                $s = New-TestStudent -Name 'A' -OnHold -UnsatSim -IsCap -Absences $abs -Done 'FI3101|2026-02-20 08:00'
                                                $row = New-TestRow -ETD '2026-03-03 10:00' -Instructor 'Lee, Ann'

                                                $fast = Invoke-HR -Student $s -Sortie $row
                                                Get-RuleCodes $fast | Should -Be @('ON_HOLD')
                                                @($fast.EventResults).Count | Should -Be 0
                                                $fast.Passed | Should -BeFalse

                                                $full = Invoke-HR -Student $s -Sortie $row -Explain
                                                Get-RuleCodes $full | Should -Be @('ON_HOLD', 'ABSENCE', 'EXP_IP', 'UNSAT_SIM')
                                                (Get-Viol $full 'EXP_IP').Message | Should -Be "CAP student needs an experienced IP (Lee, Ann isn't)"
                                                Get-MissionNames @($full.EventResults | ForEach-Object { $_.Event }) | Should -Be @('FT5001')
                                                $full.Passed | Should -BeFalse
                                }

                                It '-Explain keeps evaluating time rules past a triple turn' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-03-03 06:00', 'FT5001|2026-03-03 09:30|Jones, Bob'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 11:00' -Instructor 'Jones, Bob') -Explain
                                                Get-RuleCodes $r | Should -Be @('TRIPLE_TURN', 'TURN_TIME')
                                                (Get-Viol $r 'TURN_TIME').Message | Should -Be 'Turn time 1.5 hr < 3.25 hr minimum (vs 0930 FT5001)'
                                }

                                It 'the fast path skips the formation block once an event is already blocked; -Explain adds it' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00', 'FT5001|2026-02-23 08:00|Jones, Bob', 'FT5102|2026-02-24 08:00|Jones, Bob', 'FN3201|2026-02-24 13:00'
                                                $ob = New-TestRow -ETD '2026-03-03 08:00' -SortieType 'O&B' -Instructor 'Jones, Bob' -Id 'OB1'
                                                $fast = Invoke-HR -Student $s -Sortie $ob -Schedule @($ob)
                                                @((Get-EvtResult $fast 'FF5201').Violations | ForEach-Object { $_.Rule }) | Should -Be @('FORM_ON_OB')
                                                $full = Invoke-HR -Student $s -Sortie $ob -Schedule @($ob) -Explain
                                                @((Get-EvtResult $full 'FF5201').Violations | ForEach-Object { $_.Rule }) | Should -Be @('FORM_ON_OB', 'FORM_NO_WING_LINE')
                                }
                }

                Context 'per-event evaluation' {
                                It 'one blocked event does not reject the student when another passes' {
                                                $done = @($script:FormDone | Where-Object { $_ -notlike 'FT5190*' })
                                                $s = New-TestStudent -Name 'Lead' -Done $done
                                                $wing = New-TestStudent -Name 'Wing' -Done $done
                                                $lead = New-TestRow -ETD '2026-03-03 08:00' -Instructor 'Ortiz, Sam' -Id 'L1'
                                                $open = New-TestRow -ETD '2026-03-03 08:10' -Instructor 'Lee, Ann' -Id 'L2'
                                                foreach ($explain in @($false, $true)) {
                                                                $r = Invoke-HR -Student $s -Sortie $lead -Schedule @($lead, $open) -Students @($s, $wing) -Explain:$explain
                                                                @($r.EventResults | ForEach-Object { $_.Mission }) | Should -Be @('FF5201', 'FT5190')
                                                                (Get-EvtResult $r 'FF5201').Passed | Should -BeTrue
                                                                (Get-EvtResult $r 'FT5190').Passed | Should -BeFalse
                                                                (Get-EvtViol $r 'FT5190' 'CHECK_NEEDS_CHECK_IP').Message | Should -Be "Checkride needs a Check IP (Ortiz, Sam isn't)"
                                                                Get-MissionNames $r.PassedEvents | Should -Be @('FF5201')
                                                                $r.Passed | Should -BeTrue
                                                }
                                }

                                It 'Passed is false when no event passes, even with no student-level violations' {
                                                $s = New-TestStudent -Name 'A' -Done $script:AllButCheck
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 08:00' -Instructor 'Jones, Bob')
                                                @($r.Violations).Count | Should -Be 0
                                                @($r.PassedEvents).Count | Should -Be 0
                                                $r.Passed | Should -BeFalse
                                }
                }

                Context 'Excel-shaped values' {
                                It 'reads text flags, text IsComplete, OADate ETDs and [EXP] name tags' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00', 'FT5001|2026-02-23 08:00|Jones, Bob [EXP]'
                                                # "FALSE" text must not put the student on hold (legacy truthiness bug)
                                                $s.OnHold = 'FALSE'
                                                $s.UnsatSim = 'no'
                                                $s.IsCap = 'TRUE'
                                                $s.Syllabus[0].IsComplete = 'TRUE'
                                                $s.Syllabus[2].IsComplete = 'True'
                                                $s.Syllabus[2].ETD = 46084.3125
                                                $s.Syllabus[3].IsComplete = 'FALSE'
                                                $r = Invoke-HR -Student $s -Sortie (New-TestRow -ETD '2026-03-03 09:00' -Instructor 'Jones, Bob') -Explain
                                                Get-RuleCodes $r | Should -Be @('CAP_SINGLE_TURN', 'TURN_TIME')
                                                (Get-Viol $r 'CAP_SINGLE_TURN').Message | Should -Be 'CAP student already flies today (0730 FT5001)'
                                                (Get-Viol $r 'TURN_TIME').Message | Should -Be 'Turn time 1.5 hr < 3.25 hr minimum (vs 0730 FT5001)'
                                                Get-MissionNames $r.PassedEvents | Should -Be @('FT5102')

                                                $soft = Get-Soft -Student $s -ETD '2026-03-04 10:00' -Instructor 'Jones, Bob' -Mission 'FT5102'
                                                $soft.Breakdown['Continue String'] | Should -Be 10
                                                $soft.Breakdown['CAP Bonus (string)'] | Should -Be 8
                                                $soft.Breakdown['Counter Continuity'] | Should -Be 5
                                }

                                It 'Get-MissionKind decides by Type, falling back to the IP name only when Type is blank' {
                                                Get-MissionKind -Mission ([PSCustomObject]@{ Type = 'OFT' }) -InstructorName 'Jones, Bob' | Should -Be 'Sim'
                                                Get-MissionKind -Mission ([PSCustomObject]@{ Type = ' ift ' }) -InstructorName '' | Should -Be 'Sim'
                                                Get-MissionKind -Mission ([PSCustomObject]@{ Type = 'Single' }) -InstructorName 'SIM IP' | Should -Be 'Flight'
                                                Get-MissionKind -Mission ([PSCustomObject]@{ Type = '' }) -InstructorName 'SIM IP' | Should -Be 'Sim'
                                                Get-MissionKind -Mission ([PSCustomObject]@{ Type = '' }) -InstructorName 'Lee, Ann' | Should -Be 'Flight'
                                }
                }

                Context 'no mutation' {
                                It 'leaves the student, syllabus, sortie and schedule untouched (<Name>)' -TestCases @(
                                                @{ Name = 'O&B checkride with look-ahead'; Kind = 'ob' }
                                                @{ Name = 'formation lead needing DS'; Kind = 'form' }
                                                @{ Name = 'O&B formation with return leg'; Kind = 'obform' }
                                ) {
                                                $others = @()
                                                if ($Kind -eq 'ob') {
                                                                $s = New-TestStudent -Name 'A' -Done $script:AllButCheck -SimDebt 2
                                                                $row = New-TestRow -ETD '2026-03-03 08:00' -SortieType 'O&B' -Instructor 'Jones, Bob' -Id 'OB1'
                                                                $sched = @($row)
                                                                $mission = 'FT5190'
                                                }
                                                elseif ($Kind -eq 'form') {
                                                                $s = New-TestStudent -Name 'Lead' -Done $script:FormDone -SimDebt 2
                                                                $row = New-TestRow -ETD '2026-03-03 08:00' -Instructor 'Jones, Bob' -Id 'L1'
                                                                $sched = @($row, (New-TestRow -ETD '2026-03-03 08:10' -Instructor 'Ortiz, Sam' -Id 'L2'))
                                                                $others = @((New-TestStudent -Name 'Newbie'), (New-TestStudent -Name 'Single' -Done 'FI3101|2026-02-20 08:00'))
                                                                $mission = $null
                                                }
                                                else {
                                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00', 'FT5001|2026-02-23 08:00|Jones, Bob', 'FT5102|2026-02-24 08:00|Jones, Bob', 'FN3201|2026-02-24 13:00'
                                                                Add-StudentReservation -Student $s -LineId 'W1' -ETD '2026-03-04 06:00' -SortieType 'OFT' -Mission 'FE3102' -Instructor 'SIM IP'
                                                                $row = New-TestRow -ETD '2026-03-03 13:00' -SortieType 'O&B' -Instructor 'Jones, Bob' -Id 'OB1' -LinkedLineId 'R1'
                                                                $sched = @($row, (New-TestRow -ETD '2026-03-03 16:00' -SortieType 'O&B Return' -Instructor 'Jones, Bob' -Id 'R1'))
                                                                $mission = $null
                                                }
                                                $all = @($s) + $others
                                                $snapStudents = $all | ConvertTo-Json -Depth 10
                                                $snapSchedule = $sched | ConvertTo-Json -Depth 10
                                                $props = @($s.PSObject.Properties | ForEach-Object { $_.Name })

                                                foreach ($explain in @($false, $true)) {
                                                                $null = Invoke-HR -Student $s -Sortie $row -Schedule $sched -Students $all -Mission $mission -Explain:$explain
                                                }

                                                ($all | ConvertTo-Json -Depth 10) | Should -Be $snapStudents
                                                ($sched | ConvertTo-Json -Depth 10) | Should -Be $snapSchedule
                                                @($s.PSObject.Properties | ForEach-Object { $_.Name }) | Should -Be $props
                                                $row.Instructor | Should -BeOfType [string]
                                }
                }
}


# =================================================================================
#  SOFTRULES
# =================================================================================
Describe 'SoftRules -Detailed' {
                It 'returns Score, Breakdown, Notes, Student, Mission and ETD; Score is the Breakdown total' {
                                $s = New-TestStudent -Name 'Smith, John' -Done 'FI3101|2026-03-02 08:00'
                                $r = Get-Soft -Student $s -Instructor 'Jones, Bob' -Mission 'FT5001'
                                $r.Student | Should -Be 'Smith, John'
                                $r.Mission | Should -Be 'FT5001'
                                $r.ETD | Should -Be '2026-03-03 10:00:00'
                                @($r.Breakdown.Keys) | Should -Be @('Dollar Ride Bonus', 'Sim Debt (flight)', 'Exhaustion', 'Student Priority')
                                $r.Breakdown['Dollar Ride Bonus'] | Should -Be 20
                                $r.Breakdown['Sim Debt (flight)'] | Should -Be 0
                                $r.Breakdown['Exhaustion'] | Should -Be -1
                                $r.Score | Should -Be 19
                                $total = 0.0
                                foreach ($k in $r.Breakdown.Keys) { $total += $r.Breakdown[$k] }
                                $r.Score | Should -Be $total
                                @($r.Notes).Count | Should -Be 0
                }

                It 'without -Detailed returns the same score as a number' {
                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-03-02 08:00' -Priority 3
                                $row = New-TestRow -ETD '2026-03-03 10:00' -Instructor 'Jones, Bob'
                                $row.Mission = Find-SyllabusMission -Student $s -Mission 'FT5001'
                                $plain = SoftRules -Student $s -Config $script:W -Sortie $row -Instructors $script:IPs
                                $plain | Should -BeOfType [double]
                                $plain | Should -Be 22
                }

                It 'Sim Debt: flights pay -SimDebt, sims earn +SimDebt (<Mission> on <SortieType>)' -TestCases @(
                                @{ Mission = 'FT5102'; SortieType = 'Day Local'; Label = 'Sim Debt (flight)'; Points = -3 }
                                @{ Mission = 'FE3102'; SortieType = 'OFT'; Label = 'Sim Debt (sim)'; Points = 3 }
                                @{ Mission = 'FN3201'; SortieType = 'IFT'; Label = 'Sim Debt (sim)'; Points = 3 }
                ) {
                                $s = New-TestStudent -Name 'A' -SimDebt 3
                                $r = Get-Soft -Student $s -SortieType $SortieType -Mission $Mission
                                $r.Breakdown[$Label] | Should -Be $Points
                                $other = if ($Label -eq 'Sim Debt (flight)') { 'Sim Debt (sim)' } else { 'Sim Debt (flight)' }
                                $r.Breakdown.Contains($other) | Should -BeFalse
                }

                It 'CAP Bonus' {
                                $s = New-TestStudent -Name 'A' -IsCap
                                $r = Get-Soft -Student $s -SortieType 'IFT' -Mission 'FI3101'
                                $r.Breakdown['CAP Bonus'] | Should -Be 8
                                $r.Breakdown.Contains('CAP Bonus (string)') | Should -BeFalse
                }

                Context 'strings and counters' {
                                It 'Continue String (+ Counter Continuity) when the last IP flies again' {
                                                $s = New-TestStudent -Name 'A' -Done 'FT5001|2026-02-24 08:00|Lee, Ann', 'FT5102|2026-02-26 08:00|Jones, Bob'
                                                $r = Get-Soft -Student $s -Instructor 'Jones, Bob' -Mission 'FT5103'
                                                $r.Breakdown['Continue String'] | Should -Be 10
                                                $r.Breakdown['Counter Continuity'] | Should -Be 5
                                                foreach ($k in 'String 4', 'String 5', 'Zipper', 'Break String', 'Non-Counter', 'CAP Bonus (string)') { $r.Breakdown.Contains($k) | Should -BeFalse }
                                }

                                It 'CAP Bonus (string) reads the student''s IsCap' {
                                                $s = New-TestStudent -Name 'A' -IsCap -Done 'FT5001|2026-02-24 08:00|Lee, Ann', 'FT5102|2026-02-26 08:00|Jones, Bob'
                                                $r = Get-Soft -Student $s -Instructor 'Jones, Bob' -Mission 'FT5103'
                                                $r.Breakdown['CAP Bonus'] | Should -Be 8
                                                $r.Breakdown['Continue String'] | Should -Be 10
                                                $r.Breakdown['CAP Bonus (string)'] | Should -Be 8
                                }

                                It 'String 4 after three in a row with the same IP' {
                                                $s = New-TestStudent -Name 'A' -Done 'FT5001|2026-02-24 08:00|Jones, Bob', 'FT5102|2026-02-25 08:00|Jones, Bob', 'FT5103|2026-02-26 08:00|Jones, Bob'
                                                $r = Get-Soft -Student $s -Instructor 'Jones, Bob' -Mission 'FT5104'
                                                $r.Breakdown['String 4'] | Should -Be -10
                                                $r.Breakdown.Contains('Continue String') | Should -BeFalse
                                                $r.Breakdown.Contains('String 5') | Should -BeFalse
                                }

                                It 'String 5 after four in a row with the same IP' {
                                                $s = New-TestStudent -Name 'A' -Done 'FT5001|2026-02-23 08:00|Jones, Bob', 'FT5102|2026-02-24 08:00|Jones, Bob', 'FT5103|2026-02-25 08:00|Jones, Bob', 'FT5104|2026-02-26 08:00|Jones, Bob'
                                                $r = Get-Soft -Student $s -Instructor 'Jones, Bob' -Mission 'FF5201'
                                                $r.Breakdown['String 5'] | Should -Be -100
                                                $r.Breakdown.Contains('String 4') | Should -BeFalse
                                }

                                It 'Zipper when the IP flew the ride before last' {
                                                $s = New-TestStudent -Name 'A' -Done 'FT5001|2026-02-24 08:00|Jones, Bob', 'FT5102|2026-02-26 08:00|Lee, Ann'
                                                $r = Get-Soft -Student $s -Instructor 'Jones, Bob' -Mission 'FT5103'
                                                $r.Breakdown['Zipper'] | Should -Be 5
                                                $r.Breakdown['Counter Continuity'] | Should -Be 5
                                                $r.Breakdown.Contains('Break String') | Should -BeFalse
                                }

                                It 'Break String for a new IP (no Counter Continuity)' {
                                                $s = New-TestStudent -Name 'A' -Done 'FT5001|2026-02-24 08:00|Lee, Ann', 'FT5102|2026-02-26 08:00|Ortiz, Sam'
                                                $r = Get-Soft -Student $s -Instructor 'Jones, Bob' -Mission 'FT5103'
                                                $r.Breakdown['Break String'] | Should -Be -5
                                                $r.Breakdown.Contains('Counter Continuity') | Should -BeFalse
                                }

                                It 'Counter Continuity for an IP who flew the student before (older than the zipper)' {
                                                $s = New-TestStudent -Name 'A' -Done 'FT5001|2026-02-24 08:00|Jones, Bob', 'FT5102|2026-02-25 08:00|Lee, Ann', 'FT5103|2026-02-26 08:00|Ortiz, Sam'
                                                $r = Get-Soft -Student $s -Instructor 'Jones, Bob' -Mission 'FT5104'
                                                $r.Breakdown['Break String'] | Should -Be -5
                                                $r.Breakdown['Counter Continuity'] | Should -Be 5
                                }

                                It 'a sim flown with a named IP is not a flight for strings (decided by Type)' {
                                                $s = New-TestStudent -Name 'A' -Done 'FT5001|2026-02-24 08:00|Jones, Bob', 'FE3102|2026-02-26 08:00|Lee, Ann'
                                                $r = Get-Soft -Student $s -Instructor 'Jones, Bob' -Mission 'FT5102'
                                                $r.Breakdown['Continue String'] | Should -Be 10
                                                $r.Breakdown.Contains('Break String') | Should -BeFalse
                                }

                                It 'Non-Counter for Leech (non-counter) when he did not fly the last two rides' {
                                                $s = New-TestStudent -Name 'A' -Done 'FT5001|2026-02-24 08:00|Jones, Bob', 'FT5102|2026-02-26 08:00|Ortiz, Sam'
                                                $r = Get-Soft -Student $s -Instructor 'Leech, Tom' -Mission 'FT5103'
                                                $r.Breakdown['Non-Counter'] | Should -Be -15
                                                foreach ($k in 'Continue String', 'Zipper', 'Break String', 'Counter Continuity') { $r.Breakdown.Contains($k) | Should -BeFalse }
                                }

                                It 'Leech becomes a counter if he flew IP1 (Continue String + Counter Continuity)' {
                                                $s = New-TestStudent -Name 'A' -Done 'FT5001|2026-02-24 08:00|Jones, Bob', 'FT5102|2026-02-26 08:00|Leech, Tom'
                                                $r = Get-Soft -Student $s -Instructor 'Leech, Tom' -Mission 'FT5103'
                                                $r.Breakdown.Contains('Non-Counter') | Should -BeFalse
                                                $r.Breakdown['Continue String'] | Should -Be 10
                                                $r.Breakdown['Counter Continuity'] | Should -Be 5
                                }

                                It 'Leech becomes a counter if he flew IP2 (Zipper)' {
                                                $s = New-TestStudent -Name 'A' -Done 'FT5001|2026-02-24 08:00|Leech, Tom', 'FT5102|2026-02-26 08:00|Jones, Bob'
                                                $r = Get-Soft -Student $s -Instructor 'Leech, Tom' -Mission 'FT5103'
                                                $r.Breakdown.Contains('Non-Counter') | Should -BeFalse
                                                $r.Breakdown['Zipper'] | Should -Be 5
                                }

                                It 'Leech stays a non-counter if he only flew IP3' {
                                                $s = New-TestStudent -Name 'A' -Done 'FT5001|2026-02-24 08:00|Leech, Tom', 'FT5102|2026-02-25 08:00|Jones, Bob', 'FT5103|2026-02-26 08:00|Ortiz, Sam'
                                                $r = Get-Soft -Student $s -Instructor 'Leech, Tom' -Mission 'FT5104'
                                                $r.Breakdown['Non-Counter'] | Should -Be -15
                                                $r.Breakdown.Contains('Counter Continuity') | Should -BeFalse
                                }

                                It 'IsCounter cell: blank counts as a counter, FALSE text does not' {
                                                $ips = @(
                                                                [PSCustomObject]@{ Name = 'Blank, Bo'; IsCounter = ''; IsExperienced = $true; IsCheck = $false; Callsign = 'X'; Num = 1 }
                                                                [PSCustomObject]@{ Name = 'Text, Ty'; IsCounter = 'FALSE'; IsExperienced = $true; IsCheck = $false; Callsign = 'X'; Num = 2 }
                                                )
                                                $s = New-TestStudent -Name 'A' -Done 'FT5001|2026-02-24 08:00|Jones, Bob', 'FT5102|2026-02-26 08:00|Ortiz, Sam'
                                                $blank = Get-Soft -Student $s -Instructor 'Blank, Bo' -Mission 'FT5103' -Instructors $ips
                                                $blank.Breakdown.Contains('Non-Counter') | Should -BeFalse
                                                $blank.Breakdown['Break String'] | Should -Be -5
                                                $text = Get-Soft -Student $s -Instructor 'Text, Ty' -Mission 'FT5103' -Instructors $ips
                                                $text.Breakdown['Non-Counter'] | Should -Be -15
                                }

                                It 'no string or counter scoring with SIM IP or no IP' {
                                                $s = New-TestStudent -Name 'A' -Done 'FT5001|2026-02-24 08:00|Jones, Bob', 'FT5102|2026-02-26 08:00|Jones, Bob'
                                                foreach ($r in @((Get-Soft -Student $s -SortieType 'OFT' -Mission 'FE3102'), (Get-Soft -Student $s -Instructor '' -Mission 'FT5103'))) {
                                                                foreach ($k in 'Non-Counter', 'Continue String', 'String 4', 'String 5', 'Zipper', 'Break String', 'Counter Continuity') { $r.Breakdown.Contains($k) | Should -BeFalse }
                                                }
                                }
                }

                Context 'sim matching, exhaustion, BIT, checkrides and priority' {
                                It 'OFT/IFT Mismatch uses SoftWeights.OFTIFTPen and adds a note' {
                                                $s = New-TestStudent -Name 'A'
                                                $r = Get-Soft -Student $s -SortieType 'OFT' -Mission 'FI3101'
                                                $r.Breakdown['OFT/IFT Mismatch'] | Should -Be -7
                                                @($r.Notes) | Should -Be @('IFT in an OFT')

                                                $w = New-TestSoftWeights
                                                $w.OFTIFTPen = -3
                                                (Get-Soft -Student $s -SortieType 'OFT' -Mission 'FI3101' -Weights $w).Breakdown['OFT/IFT Mismatch'] | Should -Be -3
                                }

                                It 'no OFT/IFT Mismatch for <Mission> on <SortieType>' -TestCases @(
                                                @{ Mission = 'FI3101'; SortieType = 'IFT' }
                                                @{ Mission = 'FE3102'; SortieType = 'OFT' }
                                ) {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-20 08:00'
                                                $r = Get-Soft -Student $s -SortieType $SortieType -Mission $Mission
                                                $r.Breakdown.Contains('OFT/IFT Mismatch') | Should -BeFalse
                                }

                                It 'Exhaustion counts flights x2 and sims x1 in the last ExhaustionDays, before this sortie' {
                                                # FI3101 (6 days back) is out; FE3102 (5 days) in; FT5001 yesterday and FT5102 earlier today in; FN3201 later today out
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-25 08:00', 'FE3102|2026-02-26 08:00', 'FT5001|2026-03-02 08:00|Jones, Bob', 'FT5102|2026-03-03 06:00|Jones, Bob', 'FN3201|2026-03-03 15:00'
                                                $r = Get-Soft -Student $s -Instructor 'Jones, Bob' -Mission 'FT5103'
                                                $r.Breakdown['Exhaustion'] | Should -Be -5
                                }

                                It 'Exhaustion is always reported (0 when rested)' {
                                                $s = New-TestStudent -Name 'A'
                                                (Get-Soft -Student $s -SortieType 'IFT' -Mission 'FI3101').Breakdown['Exhaustion'] | Should -Be 0
                                }

                                It 'BIT Tracker: last flight <LastFlight> -> <Label> <Points>' -TestCases @(
                                                @{ LastFlight = '2026-02-28 08:00'; Cap = $false; Label = 'BIT Tracker'; Points = 3 }
                                                @{ LastFlight = '2026-02-28 08:00'; Cap = $true; Label = 'BIT Tracker'; Points = 6 }
                                                @{ LastFlight = '2026-02-24 08:00'; Cap = $false; Label = 'BIT Tracker'; Points = 9 }
                                                @{ LastFlight = '2026-02-20 08:00'; Cap = $false; Label = 'BIT Tracker (capped)'; Points = 4.5 }
                                ) {
                                                $s = New-TestStudent -Name 'A' -IsCap:$Cap -Done "FT5001|$LastFlight|Lee, Ann"
                                                $r = Get-Soft -Student $s -Instructor 'Ortiz, Sam' -Mission 'FT5102'
                                                $r.Breakdown[$Label] | Should -Be $Points
                                                if ($Label -eq 'BIT Tracker (capped)') {
                                                                @($r.Notes) | Should -Be @('BIT! Consider an 86')
                                                                $r.Breakdown.Contains('BIT Tracker') | Should -BeFalse
                                                }
                                                else {
                                                                @($r.Notes).Count | Should -Be 0
                                                                $r.Breakdown.Contains('BIT Tracker (capped)') | Should -BeFalse
                                                }
                                }

                                It 'no BIT Tracker after yesterday''s flight or on a sim line' {
                                                $s = New-TestStudent -Name 'A' -Done 'FT5001|2026-03-02 08:00|Lee, Ann'
                                                (Get-Soft -Student $s -Instructor 'Ortiz, Sam' -Mission 'FT5102').Breakdown.Contains('BIT Tracker') | Should -BeFalse
                                                $s2 = New-TestStudent -Name 'B' -Done 'FT5001|2026-02-20 08:00|Lee, Ann'
                                                $sim = Get-Soft -Student $s2 -SortieType 'OFT' -Mission 'FE3102'
                                                $sim.Breakdown.Contains('BIT Tracker') | Should -BeFalse
                                                $sim.Breakdown.Contains('BIT Tracker (capped)') | Should -BeFalse
                                }

                                It 'Checkride Priority with a Check IP' {
                                                $s = New-TestStudent -Name 'A' -Done 'FT5104|2026-03-02 08:00|Lee, Ann'
                                                $r = Get-Soft -Student $s -Instructor 'Jones, Bob' -Mission 'FT5190'
                                                $r.Breakdown['Checkride Priority'] | Should -Be 50
                                                @($r.Notes) | Should -Contain 'Checkride! Requires Formal'
                                                (Get-Soft -Student $s -Instructor 'Ortiz, Sam' -Mission 'FT5190').Breakdown.Contains('Checkride Priority') | Should -BeFalse
                                                (Get-Soft -Student $s -Instructor 'Jones, Bob' -Mission 'FT5102').Breakdown.Contains('Checkride Priority') | Should -BeFalse
                                }

                                It 'Student Priority' {
                                                $s = New-TestStudent -Name 'A' -Priority 7
                                                $r = Get-Soft -Student $s -SortieType 'IFT' -Mission 'FI3101'
                                                $r.Breakdown['Student Priority'] | Should -Be 7
                                                $r.Score | Should -Be 7
                                }

                                It 'adds up a full breakdown with the fixture weights' {
                                                # CAP, sim debt 2, Jones flew last ride 3 days ago, plus an older sim
                                                $s = New-TestStudent -Name 'A' -IsCap -SimDebt 2 -Priority 1 -Done 'FI3101|2026-02-26 08:00', 'FT5001|2026-02-28 08:00|Jones, Bob'
                                                $r = Get-Soft -Student $s -Instructor 'Jones, Bob' -Mission 'FT5102'
                                                @($r.Breakdown.Keys) | Should -Be @('Sim Debt (flight)', 'CAP Bonus', 'Continue String', 'CAP Bonus (string)', 'Exhaustion', 'BIT Tracker', 'Counter Continuity', 'Student Priority')
                                                @($r.Breakdown.Values) | Should -Be @(-2, 8, 10, 8, -3, 6, 5, 1)
                                                $r.Score | Should -Be 33
                                }

                                It 'Get-EventScore scores a copy of the line and leaves the real line alone' {
                                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-03-02 08:00'
                                                $row = New-TestRow -ETD '2026-03-03 10:00' -Instructor 'Jones, Bob'
                                                $evt = Find-SyllabusMission -Student $s -Mission 'FT5001'
                                                $r = Get-EventScore -Student $s -Sortie $row -Evt $evt -SoftWeights $script:W -Instructors $script:IPs
                                                $r.Mission | Should -Be 'FT5001'
                                                $r.Score | Should -Be 19
                                                $row.Mission | Should -BeNullOrEmpty
                                }

                                It 'does not write PendingNote or change the student or sortie' {
                                                $s = New-TestStudent -Name 'A' -Done 'FT5001|2026-02-20 08:00|Lee, Ann'
                                                $row = New-TestRow -ETD '2026-03-03 10:00' -Instructor 'Jones, Bob'
                                                $row.Mission = Find-SyllabusMission -Student $s -Mission 'FT5102'
                                                $snap = $s | ConvertTo-Json -Depth 6
                                                $rowSnap = $row | ConvertTo-Json -Depth 6
                                                $r = SoftRules -Student $s -Config $script:W -Sortie $row -Instructors $script:IPs -Detailed
                                                @($r.Notes) | Should -Be @('BIT! Consider an 86')
                                                $s.PSObject.Properties['PendingNote'] | Should -BeNullOrEmpty
                                                ($s | ConvertTo-Json -Depth 6) | Should -Be $snap
                                                ($row | ConvertTo-Json -Depth 6) | Should -Be $rowSnap
                                }
                }
}


# =================================================================================
#  LEGACY WRAPPER (plan section 2: "HardRules wrapper")
# =================================================================================
Describe 'HardRules legacy wrapper' {
                It 'exists and keeps the legacy bool result and side effects (OptedEvents, Sortie.Mission, PendingNote)' {
                                Get-Command -Name 'HardRules' -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty -Because 'the plan keeps a HardRules wrapper with the legacy bool result'
                                $s = New-TestStudent -Name 'A' -Done 'FI3101|2026-02-23 08:00', 'FT5001|2026-02-24 08:00|Jones, Bob', 'FT5102|2026-02-25 08:00|Jones, Bob'
                                $row = New-TestRow -ETD '2026-03-03 08:00' -Instructor 'Lee, Ann'
                                $row.Instructor = Resolve-SortieInstructor -Instructors $script:IPs -Instructor $row.Instructor
                                $ok = HardRules -Student $s -Sortie $row -Schedule @($row) -Students @($s) -FormLookAhead 22
                                $ok | Should -BeTrue
                                Get-MissionNames $s.OptedEvents | Should -Be @('FT5103')
                                (Get-NameText $row.Mission) | Should -Be 'FT5103'
                                $s.PendingNote | Should -Be 'Solo Flight | Need IP in Continuity to Brief'
                }
}
