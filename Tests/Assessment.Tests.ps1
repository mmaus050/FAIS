# Pester 5 tests for Modules/Assessment.psm1: Get-PlacementAssessment (what if this
# student goes on this board line?), Get-AssessmentLines, Get-StudentPuckInfo and
# Get-InstructorPuckInfo. Boards are built with BoardModel.
# Dates (2026): Fri 20 Feb .. Fri 27 Feb history, Tue 03 Mar = the day being
# scheduled, Wed 04 Mar = the next day.

BeforeAll {
                . "$PSScriptRoot/Fixtures/New-FaisTestData.ps1"
                Import-FaisModules

                $script:IPs = New-TestInstructors
                $script:W = New-TestSoftWeights
                $script:F = New-TestFormations

                # Day Local opts FT5001 (dollar ride); nothing fits an OFT or IFT line
                $script:AbleDone = @(
                                'FI3101|2026-02-20 08:00',
                                'FE3102|2026-02-20 13:00'
                )
                # Day Local opts FT5103 (SingleSolo); IFT opts FN3201
                $script:SoloDone = @(
                                'FI3101|2026-02-20 08:00',
                                'FE3102|2026-02-20 13:00',
                                'FT5001|2026-02-23 08:00|Lee, Ann',
                                'FT5102|2026-02-24 08:00|Lee, Ann'
                )
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
                # Day Local opts FF5201 (ID 7) and the checkride FT5190 (ID 10); every flight with Lee
                $script:ChkDone = @(
                                'FI3101|2026-02-20 08:00',
                                'FE3102|2026-02-20 13:00',
                                'FT5001|2026-02-23 08:00|Lee, Ann',
                                'FT5102|2026-02-24 08:00|Lee, Ann',
                                'FT5103|2026-02-25 08:00|Lee, Ann',
                                'FN3201|2026-02-25 13:00',
                                'FT5104|2026-02-26 08:00|Lee, Ann'
                )

                # Same as ChkDone, but FT5104 was flown with Jones (the only Check IP)
                $script:JonesDone = @(
                                'FI3101|2026-02-20 08:00',
                                'FE3102|2026-02-20 13:00',
                                'FT5001|2026-02-23 08:00|Lee, Ann',
                                'FT5102|2026-02-24 08:00|Lee, Ann',
                                'FT5103|2026-02-25 08:00|Lee, Ann',
                                'FN3201|2026-02-25 13:00',
                                'FT5104|2026-02-26 08:00|Jones, Bob'
                )

                $script:FooterText = 'Based on pinned lines only; auto-filled lines can change turns.'

                # The result cache is keyed by board identity and Version only (Data is fixed while
                # the board is open), so every Describe clears it before each test

                function New-AsData {
                                param ([array]$Students)
                                return @{ Students = $Students; Instructors = $script:IPs; SoftWeights = $script:W; Formations = $script:F }
                }

                # Specs: @{ Id; ETD; Type; IP; Student; Mission; WingOf }
                function New-AsBoard {
                                param ([array]$Lines = @(), [string[]]$Days = @('2026-03-03'))
                                $b = New-Board -Days $Days
                                $made = New-Object System.Collections.Generic.List[object]
                                foreach ($s in $Lines) {
                                                $p = @{ ETD = [datetime]$s.ETD; SortieType = [string]$s.Type }
                                                if ($s.Id) { $p['Id'] = [string]$s.Id }
                                                if ($s.IP) { $p['Instructor'] = [string]$s.IP }
                                                if ($s.Student) { $p['Student'] = [string]$s.Student }
                                                if ($s.Mission) { $p['Mission'] = [string]$s.Mission }
                                                if ($s.WingOf) { $p['WingOfLineId'] = [string]$s.WingOf }
                                                $made.Add((New-BoardLine @p))
                                }
                                if ($made.Count -gt 0) { Add-BoardLine -Board $b -Line $made.ToArray() }
                                return $b
                }

                # Adds quick-entry lines (O&B + linked return etc.) to a board
                function Add-AsQuickEntry {
                                param ($Board, [string]$Text, [string]$Day = '2026-03-03', [array]$Students = @())
                                $qe = ConvertFrom-QuickEntry -Text $Text -Day $Day -Instructors $script:IPs -Students $Students -Board $Board
                                if (@($qe.Errors).Count -gt 0) { throw "test setup: $($qe.Errors -join '; ')" }
                                Add-BoardLine -Board $Board -Line $qe.Lines
                                return $qe.Lines
                }

                function Get-Assess {
                                param (
                                                $Board,
                                                [string]$LineId,
                                                [string]$Name,
                                                $Data,
                                                [string]$Mission = '',
                                                [string[]]$Exclude = @(),
                                                [switch]$Quick
                                )
                                return (Get-PlacementAssessment -Board $Board -LineId $LineId -StudentName $Name -MissionName $Mission `
                                                -ExcludeLineIds $Exclude -Data $Data -Quick:$Quick)
                }

                function Get-RuleList {
                                param ($Assessment)
                                return ((@($Assessment.Violations) | Where-Object { $null -ne $_ } | ForEach-Object { $_.Rule }) -join ',')
                }

                function Get-Viol {
                                param ($Assessment, [string]$Rule)
                                return (@($Assessment.Violations) | Where-Object { $null -ne $_ -and $_.Rule -eq $Rule } | Select-Object -First 1)
                }

                function Get-Detail {
                                param ($Assessment, [string]$Mission)
                                return (@($Assessment.EventDetails) | Where-Object { $_.Mission -eq $Mission } | Select-Object -First 1)
                }

                function Get-Points {
                                param ($Detail, [string]$Label)
                                $b = @($Detail.Breakdown) | Where-Object { $_.Label -eq $Label } | Select-Object -First 1
                                if ($null -eq $b) { return $null }
                                return $b.Points
                }

                function Get-LineText {
                                param ($Lines)
                                return @(@($Lines) | ForEach-Object { "$($_.Kind)|$($_.Text)" })
                }

                function New-Absence {
                                param ([string]$Start, [string]$End, [string]$Reason = '')
                                return [PSCustomObject]@{ Student = 'x'; StartTime = [datetime]$Start; EndTime = [datetime]$End; Reason = $Reason }
                }
}



# =================================================================================
#  STATUS, VIOLATIONS, NOTES, EVENT DETAILS
# =================================================================================
Describe 'Get-PlacementAssessment: status and result shape' {
                BeforeEach {
                                Clear-AssessmentCache
                                $able = New-TestStudent -Name 'Able, Amy' -Done $script:AbleDone
                                $solo = New-TestStudent -Name 'Solo, Sue' -Done $script:SoloDone
                                $data = New-AsData @($able, $solo)
                                $board = New-AsBoard @(
                                                @{ Id = 'L1'; ETD = '2026-03-03 07:30'; Type = 'Day Local'; IP = 'Lee, Ann' },
                                                @{ Id = 'L2'; ETD = '2026-03-03 09:00'; Type = 'Day Local'; IP = 'Ortiz, Sam' },
                                                @{ Id = 'L3'; ETD = '2026-03-03 13:00'; Type = 'OFT' }
                                )
                }

                It 'is Clean for a legal placement, with the auto mission and its score' {
                                $a = Get-Assess $board 'L1' 'Able, Amy' $data
                                $a.Status | Should -Be 'Clean'
                                @($a.Violations).Count | Should -Be 0
                                @($a.Notes).Count | Should -Be 0
                                $a.AutoMission | Should -Be 'FT5001'
                                $a.Mission | Should -Be 'FT5001'
                                $a.Score | Should -Be 20
                                $a.BestScore | Should -Be 20
                                (@($a.EligibleEvents) -join ',') | Should -Be 'FT5001'
                                $a.LineId | Should -Be 'L1'
                                $a.StudentName | Should -Be 'Able, Amy'
                                $a.BoardVersion | Should -Be $board.Version
                                $a.Quick | Should -BeFalse
                }

                It 'builds the summary from name, ETD, sortie type, mission and signed score' {
                                $a = Get-Assess $board 'L1' 'Able, Amy' $data
                                $a.Summary | Should -Be 'Able, Amy -> 0730 Day Local: FT5001 +20.0'
                }

                It 'gives EventDetails {Mission, Type, Passed, Score, Breakdown, Violations, Notes}' {
                                $a = Get-Assess $board 'L1' 'Able, Amy' $data
                                @($a.EventDetails).Count | Should -Be 1
                                $d = $a.EventDetails[0]
                                foreach ($p in @('Mission', 'Type', 'Passed', 'Score', 'Breakdown', 'Violations', 'Notes')) {
                                                $d.PSObject.Properties[$p] | Should -Not -BeNullOrEmpty -Because "EventDetails need $p"
                                }
                                $d.Mission | Should -Be 'FT5001'
                                $d.Type | Should -Be 'Single'
                                $d.Passed | Should -BeTrue
                                $d.Score | Should -Be 20
                }

                It 'gives the breakdown as {Label, Points} items that add up to the score' {
                                $d = (Get-Assess $board 'L1' 'Able, Amy' $data).EventDetails[0]
                                foreach ($b in @($d.Breakdown)) {
                                                $b.PSObject.Properties['Label'] | Should -Not -BeNullOrEmpty
                                                $b.PSObject.Properties['Points'] | Should -Not -BeNullOrEmpty
                                }
                                (@($d.Breakdown) | ForEach-Object { $_.Label }) -join ',' | Should -Be 'Dollar Ride Bonus,Sim Debt (flight),Exhaustion,Student Priority'
                                Get-Points $d 'Dollar Ride Bonus' | Should -Be 20
                                $sum = 0.0
                                foreach ($b in @($d.Breakdown)) { $sum += $b.Points }
                                $sum | Should -Be $d.Score
                }

                It 'is Warn when the chosen event only carries notes (solo flight)' {
                                $a = Get-Assess $board 'L2' 'Solo, Sue' $data
                                $a.Status | Should -Be 'Warn'
                                @($a.Violations).Count | Should -Be 0
                                $a.Mission | Should -Be 'FT5103'
                                ($a.Notes -join '|') | Should -Be 'Solo Flight|Need IP in Continuity to Brief'
                                (Get-Detail $a 'FT5103').Notes -join '|' | Should -Be 'Solo Flight|Need IP in Continuity to Brief'
                }

                It 'is Warn for a student-level note alone (3.25 hr turn after an earlier pin)' {
                                $b2 = New-AsBoard @(
                                                @{ Id = 'A1'; ETD = '2026-03-03 07:30'; Type = 'Day Local'; IP = 'Lee, Ann'; Student = 'Able, Amy' },
                                                @{ Id = 'A2'; ETD = '2026-03-03 10:45'; Type = 'Day Local'; IP = 'Lee, Ann' }
                                )
                                $a = Get-Assess $b2 'A2' 'Able, Amy' $data
                                $a.Status | Should -Be 'Warn'
                                @($a.Violations).Count | Should -Be 0
                                ($a.Notes -join '|') | Should -Be 'Less than 3.5 hr TT, Flt/CC Approval Required'
                }

                It 'is Hard for a student-level violation and counts it in the summary' {
                                $hold = New-TestStudent -Name 'Hold, Hal' -OnHold -Done $script:AbleDone
                                $a = Get-Assess $board 'L1' 'Hold, Hal' (New-AsData @($hold))
                                $a.Status | Should -Be 'Hard'
                                Get-RuleList $a | Should -Be 'ON_HOLD'
                                $v = Get-Viol $a 'ON_HOLD'
                                $v.Severity | Should -Be 'Hard'
                                $v.Message | Should -Be 'Student is on hold'
                                $a.Summary | Should -Be 'Hold, Hal -> 0730 Day Local: FT5001 +20.0, 1 rule flag'
                }

                It 'lists every violation (explain mode) and pluralizes the summary' {
                                $hold = New-TestStudent -Name 'Hold, Hal' -OnHold -Done $script:AbleDone -Absences @(New-Absence '2026-03-03 07:00' '2026-03-03 09:00' 'Dentist')
                                $a = Get-Assess $board 'L1' 'Hold, Hal' (New-AsData @($hold))
                                $a.Status | Should -Be 'Hard'
                                Get-RuleList $a | Should -Be 'ON_HOLD,ABSENCE'
                                (Get-Viol $a 'ABSENCE').Message | Should -Be 'Unavailable 0700-0900 (Dentist); sortie window 0600-1000'
                                $a.Summary | Should -Be 'Hold, Hal -> 0730 Day Local: FT5001 +20.0, 2 rule flags'
                }

                It 'is Hard with no event when nothing in the syllabus fits the line' {
                                $a = Get-Assess $board 'L3' 'Able, Amy' $data
                                $a.Status | Should -Be 'Hard'
                                Get-RuleList $a | Should -Be 'NO_ELIGIBLE_EVENT'
                                (Get-Viol $a 'NO_ELIGIBLE_EVENT').Message | Should -Be 'No eligible OFT events (next: FT5001 Single)'
                                $a.AutoMission | Should -Be ''
                                $a.Mission | Should -Be ''
                                $a.Score | Should -BeNullOrEmpty
                                @($a.EventDetails).Count | Should -Be 0
                                @($a.EligibleEvents).Count | Should -Be 0
                                $a.Summary | Should -Be 'Able, Amy -> 1300 OFT: no event, 1 rule flag'
                }

                It 'flags an O&B Return line as Hard RETURN_LINE' {
                                $b2 = New-AsBoard @()
                                [void](Add-AsQuickEntry -Board $b2 -Text '1015 O&B Ortiz')
                                $ret = @($b2.Lines | Where-Object { $_.SortieType -eq 'O&B Return' })[0]
                                foreach ($q in @($false, $true)) {
                                                Clear-AssessmentCache
                                                $a = Get-Assess $b2 $ret.Id 'Able, Amy' $data -Quick:$q
                                                $a.Status | Should -Be 'Hard'
                                                Get-RuleList $a | Should -Be 'RETURN_LINE'
                                                $a.Violations[0].Severity | Should -Be 'Hard'
                                                $a.Violations[0].Message | Should -Be 'O&B Return lines are filled from their O&B'
                                                $a.Summary | Should -Be 'Return lines follow their O&B'
                                                $a.AutoMission | Should -Be ''
                                                @($a.EventDetails).Count | Should -Be 0
                                }
                }

                It 'flags a student who is not in MasterData as Hard UNKNOWN_STUDENT' {
                                $a = Get-Assess $board 'L1' 'Nobody, Ned' $data
                                $a.Status | Should -Be 'Hard'
                                Get-RuleList $a | Should -Be 'UNKNOWN_STUDENT'
                                $a.Violations[0].Message | Should -Be "'Nobody, Ned' is not in MasterData"
                                $a.Summary | Should -Be "'Nobody, Ned' is not in MasterData"
                                $a.StudentName | Should -Be 'Nobody, Ned'
                }

                It 'resolves a unique last name to the student' {
                                $a = Get-Assess $board 'L1' 'Able' $data
                                $a.Status | Should -Be 'Clean'
                                $a.Summary | Should -Be 'Able, Amy -> 0730 Day Local: FT5001 +20.0'
                }

                It 'returns Error / Line not found for an unknown line id' {
                                $a = Get-Assess $board 'NOPE' 'Able, Amy' $data
                                $a.Status | Should -Be 'Error'
                                $a.Summary | Should -Be 'Line not found'
                }
}


# =================================================================================
#  WHICH EVENT: AUTO MISSION, EXPLICIT MISSION
# =================================================================================
Describe 'Get-PlacementAssessment: auto mission and explicit -MissionName' {
                BeforeEach {
                                Clear-AssessmentCache
                                $able = New-TestStudent -Name 'Able, Amy' -Done $script:AbleDone
                                $chk = New-TestStudent -Name 'Chk, Cal' -Done $script:ChkDone
                                $data = New-AsData @($able, $chk)
                                # Chk alone: with a single-ship student around, FF5201 would get FORM_PREFER_SINGLE
                                $chkData = New-AsData @($chk)
                                $board = New-AsBoard @(
                                                @{ Id = 'C1'; ETD = '2026-03-03 07:30'; Type = 'Day Local'; IP = 'Jones, Bob' },
                                                @{ Id = 'C2'; ETD = '2026-03-03 07:45'; Type = 'Day Local'; IP = 'Lee, Ann' },
                                                @{ Id = 'C3'; ETD = '2026-03-03 13:00'; Type = 'OFT' }
                                )
                }

                It 'AutoMission is the best-scoring passing event, not the lowest syllabus ID' {
                                # FF5201 (ID 7) passes with a DS note; the checkride FT5190 (ID 10) scores +50 with a Check IP
                                $a = Get-Assess $board 'C1' 'Chk, Cal' $chkData
                                (Get-Detail $a 'FF5201').Passed | Should -BeTrue
                                (Get-Detail $a 'FT5190').Passed | Should -BeTrue
                                (Get-Detail $a 'FF5201').Score | Should -Be -1
                                (Get-Detail $a 'FT5190').Score | Should -Be 49
                                Get-Points (Get-Detail $a 'FT5190') 'Checkride Priority' | Should -Be 50
                                $a.AutoMission | Should -Be 'FT5190'
                                $a.Mission | Should -Be 'FT5190'
                                $a.BestScore | Should -Be 49
                                $a.Score | Should -Be 49
                                (@($a.EligibleEvents) -join ',') | Should -Be 'FF5201,FT5190'
                }

                It 'keeps every candidate in EventDetails with its own notes' {
                                $a = Get-Assess $board 'C1' 'Chk, Cal' $chkData
                                (@($a.EventDetails) | ForEach-Object { $_.Mission }) -join ',' | Should -Be 'FF5201,FT5190'
                                (Get-Detail $a 'FF5201').Notes -join '|' | Should -Be 'Direct Support Required'
                                (Get-Detail $a 'FT5190').Notes -join '|' | Should -Be 'Checkride! Requires Formal'
                                # Assessment notes come from the chosen event only
                                ($a.Notes -join '|') | Should -Be 'Checkride! Requires Formal'
                                $a.Status | Should -Be 'Warn'
                }

                It 'AutoMission skips a higher-scoring event that fails its event rules' {
                                # Jones flew FT5104 with this student, so his checkride fails CHECK_SAME_IP
                                $jay = New-TestStudent -Name 'Jay, Cee' -Done $script:JonesDone
                                $a = Get-Assess $board 'C1' 'Jay, Cee' (New-AsData @($jay))
                                $ft = Get-Detail $a 'FT5190'
                                $ff = Get-Detail $a 'FF5201'
                                $ft.Passed | Should -BeFalse
                                (@($ft.Violations) | ForEach-Object { $_.Rule }) -join ',' | Should -Be 'CHECK_SAME_IP'
                                @($ft.Violations)[0].Message | Should -Be 'Evaluator Jones, Bob already flew with this student (FT5104 26 Feb)'
                                $ft.Score | Should -BeGreaterThan $ff.Score
                                $ff.Passed | Should -BeTrue
                                $a.AutoMission | Should -Be 'FF5201'
                                $a.BestScore | Should -Be $ff.Score
                                $a.Status | Should -Be 'Warn'
                }

                It 'lists both checkride rules for an IP who flew with the student and is not a Check IP' {
                                $a = Get-Assess $board 'C2' 'Chk, Cal' $chkData
                                $ft = Get-Detail $a 'FT5190'
                                (@($ft.Violations) | ForEach-Object { $_.Rule }) -join ',' | Should -Be 'CHECK_SAME_IP,CHECK_NEEDS_CHECK_IP'
                                (@($ft.Violations) | ForEach-Object { $_.Message }) -join '|' | Should -Be "Evaluator Lee, Ann already flew with this student (FT5001 23 Feb)|Checkride needs a Check IP (Lee, Ann isn't)"
                                # nothing passes: no auto mission, the first event is shown, and it is Hard
                                $a.AutoMission | Should -Be ''
                                $a.Mission | Should -Be 'FF5201'
                                $a.Status | Should -Be 'Hard'
                }

                It 'evaluates an explicit -MissionName even when it is not eligible (MISSION_NOT_ELIGIBLE)' {
                                $a = Get-Assess $board 'C1' 'Able, Amy' $data -Mission 'FT5102'
                                $a.Status | Should -Be 'Hard'
                                $a.Mission | Should -Be 'FT5102'
                                Get-RuleList $a | Should -Be 'MISSION_NOT_ELIGIBLE'
                                $v = Get-Viol $a 'MISSION_NOT_ELIGIBLE'
                                $v.Severity | Should -Be 'Hard'
                                $v.Mission | Should -Be 'FT5102'
                                $v.Message | Should -Be "FT5102 isn't eligible here (prereq FT5001 not complete)"
                                $d = Get-Detail $a 'FT5102'
                                $d | Should -Not -BeNullOrEmpty
                                $d.Passed | Should -BeFalse
                                $d.Score | Should -Not -BeNullOrEmpty
                                @($d.Breakdown).Count | Should -BeGreaterThan 0
                                $a.Summary | Should -Be 'Able, Amy -> 0730 Day Local: FT5102 0, 1 rule flag'
                }

                It 'still lists the valid events in EligibleEvents when the explicit mission is ineligible' {
                                $a = Get-Assess $board 'C1' 'Able, Amy' $data -Mission 'FT5102'
                                (@($a.EligibleEvents) -join ',') | Should -Be 'FT5001'
                }

                It 'explains why an explicit mission is ineligible: <Mission> on <LineId>' -TestCases @(
                                @{ LineId = 'C1'; Mission = 'FE3102'; Expected = "FE3102 isn't eligible here (already complete)" }
                                @{ LineId = 'C3'; Mission = 'FT5001'; Expected = "FT5001 isn't eligible here (Type Single not allowed on OFT)" }
                ) {
                                param ($LineId, $Mission, $Expected)
                                $a = Get-Assess $board $LineId 'Able, Amy' $data -Mission $Mission
                                $a.Status | Should -Be 'Hard'
                                (Get-Viol $a 'MISSION_NOT_ELIGIBLE').Message | Should -Be $Expected
                                $a.Mission | Should -Be $Mission
                                $a.RequestedMission | Should -Be $Mission
                }

                It 'scores an ineligible explicit mission with its real breakdown (FT5001 on an OFT line)' {
                                $a = Get-Assess $board 'C3' 'Able, Amy' $data -Mission 'FT5001'
                                $a.Score | Should -Be 20
                                Get-Points (Get-Detail $a 'FT5001') 'Dollar Ride Bonus' | Should -Be 20
                                @($a.EligibleEvents).Count | Should -Be 0
                }

                It 'an eligible explicit mission is Clean and RequestedMission echoes it' {
                                $a = Get-Assess $board 'C2' 'Able, Amy' $data -Mission 'FT5001'
                                $a.Status | Should -Be 'Clean'
                                $a.Mission | Should -Be 'FT5001'
                                $a.RequestedMission | Should -Be 'FT5001'
                }

                It 'matches the explicit mission case-insensitively and echoes it as given' {
                                $a = Get-Assess $board 'C2' 'Able, Amy' $data -Mission 'ft5001'
                                $a.Status | Should -Be 'Clean'
                                $a.Mission | Should -Be 'FT5001'
                                $a.RequestedMission | Should -Be 'ft5001'
                }

                It 'RequestedMission is blank without -MissionName' {
                                $a = Get-Assess $board 'C2' 'Able, Amy' $data
                                $a.RequestedMission | Should -Be ''
                }

                It 'flags an explicit mission that is not in the syllabus as MISSION_NOT_FOUND' {
                                $a = Get-Assess $board 'C2' 'Able, Amy' $data -Mission 'ZZ9999'
                                $a.Status | Should -Be 'Hard'
                                Get-RuleList $a | Should -Be 'MISSION_NOT_FOUND'
                                (Get-Viol $a 'MISSION_NOT_FOUND').Message | Should -Be "ZZ9999 isn't in Able, Amy's syllabus"
                                $a.Mission | Should -Be ''
                                $a.RequestedMission | Should -Be 'ZZ9999'
                                (@($a.EligibleEvents) -join ',') | Should -Be 'FT5001'
                }
}


# =================================================================================
#  THE STUDENT'S OTHER PINS
# =================================================================================
Describe "Get-PlacementAssessment: the student's other pins" {
                BeforeEach {
                                Clear-AssessmentCache
                                $able = New-TestStudent -Name 'Able, Amy' -Done $script:AbleDone
                                $baker = New-TestStudent -Name 'Baker, Ben' -Done $script:AbleDone
                                $data = New-AsData @($able, $baker)
                }

                It 'applies an earlier pin as completed (next event moves on, turn time vs that pin)' {
                                $b = New-AsBoard @(
                                                @{ Id = 'E1'; ETD = '2026-03-03 07:30'; Type = 'Day Local'; IP = 'Lee, Ann'; Student = 'Able, Amy' },
                                                @{ Id = 'E2'; ETD = '2026-03-03 09:00'; Type = 'Day Local'; IP = 'Ortiz, Sam' }
                                )
                                $a = Get-Assess $b 'E2' 'Able, Amy' $data
                                $a.Status | Should -Be 'Hard'
                                Get-RuleList $a | Should -Be 'TURN_TIME'
                                (Get-Viol $a 'TURN_TIME').Message | Should -Be 'Turn time 1.5 hr < 3.25 hr minimum (vs 0730 FT5001)'
                                (@($a.EligibleEvents) -join ',') | Should -Be 'FT5102'
                                $a.AutoMission | Should -Be 'FT5102'
                                # the earlier FT5001 added sim debt and a Lee string
                                $d = Get-Detail $a 'FT5102'
                                Get-Points $d 'Sim Debt (flight)' | Should -Be -1
                                Get-Points $d 'Break String' | Should -Be -5
                }

                It 'uses the explicit mission of an earlier pin' {
                                # FT5102 made flyable right after FI3101, so the 0730 pin has two choices
                                $syl = $able.Syllabus | Where-Object { $_.Mission -eq 'FT5102' }
                                $syl.Prereqs = 'FI3101'
                                $b = New-AsBoard @(
                                                @{ Id = 'E1'; ETD = '2026-03-03 07:30'; Type = 'Day Local'; IP = 'Lee, Ann'; Student = 'Able, Amy'; Mission = 'FT5102' },
                                                @{ Id = 'E2'; ETD = '2026-03-03 11:00'; Type = 'Day Local'; IP = 'Ortiz, Sam' }
                                )
                                $a = Get-Assess $b 'E2' 'Able, Amy' $data
                                # FT5102 is flown at 0730, so FT5001 and FT5103 are open at 1100
                                (@($a.EligibleEvents) -join ',') | Should -Be 'FT5001,FT5103'
                                $a.AutoMission | Should -Be 'FT5001'
                                $a.Status | Should -Be 'Clean'

                                Set-BoardMission -Board $b -Id 'E1' -Mission 'FT5001'
                                $a2 = Get-Assess $b 'E2' 'Able, Amy' $data
                                (@($a2.EligibleEvents) -join ',') | Should -Be 'FT5102'
                }

                It 'applies an earlier (auto) pin as the event its own assessment picks' {
                                # The 0730 pin's own assessment (and the fill) fly the checkride FT5190 there:
                                # FF5201 has no wing line. The 1200 line must then see FT5190 as flown,
                                # leaving FF5201. Instead the context marks the lowest-ID valid event
                                # (FF5201) complete and offers FF5202 and FT5190 again.
                                $chk = New-TestStudent -Name 'Chk, Cal' -Done $script:ChkDone
                                $cd = New-AsData @($chk)
                                $b = New-AsBoard @(
                                                @{ Id = 'E1'; ETD = '2026-03-03 07:30'; Type = 'Day Local'; IP = 'Jones, Bob'; Student = 'Chk, Cal' },
                                                @{ Id = 'E2'; ETD = '2026-03-03 12:00'; Type = 'Day Local'; IP = 'Lee, Ann' }
                                )
                                (Get-Assess $b 'E1' 'Chk, Cal' $cd).AutoMission | Should -Be 'FT5190'
                                $a = Get-Assess $b 'E2' 'Chk, Cal' $cd
                                (@($a.EligibleEvents) -join ',') | Should -Be 'FF5201'
                }

                It 'applies a pin on the previous day as completed (CREW_REST)' {
                                $b = New-AsBoard -Days @('2026-03-03', '2026-03-04') -Lines @(
                                                @{ Id = 'D1'; ETD = '2026-03-03 18:00'; Type = 'Day Local'; IP = 'Lee, Ann'; Student = 'Able, Amy' },
                                                @{ Id = 'D2'; ETD = '2026-03-04 07:00'; Type = 'Day Local'; IP = 'Ortiz, Sam' }
                                )
                                $a = Get-Assess $b 'D2' 'Able, Amy' $data
                                $a.Status | Should -Be 'Hard'
                                Get-RuleList $a | Should -Be 'CREW_REST'
                                (Get-Viol $a 'CREW_REST').Message | Should -Be 'Crew rest 13.0 hr < 15 hr since Tue 1800 FT5001'
                                $a.Mission | Should -Be 'FT5102'
                }

                It 'applies an earlier O&B pin and its return leg as two completed events (triple turn)' {
                                $b = New-AsBoard @(@{ Id = 'T1'; ETD = '2026-03-03 13:30'; Type = 'Day Local'; IP = 'Lee, Ann' })
                                [void](Add-AsQuickEntry -Board $b -Text '0700 O&B Ortiz sp:Able' -Students @($able, $baker))
                                $a = Get-Assess $b 'T1' 'Able, Amy' $data
                                $a.Status | Should -Be 'Hard'
                                Get-RuleList $a | Should -Be 'TRIPLE_TURN'
                                (Get-Viol $a 'TRIPLE_TURN').Message | Should -Be 'Already has 2 events today (0700 FT5001, 1000 FT5102) - triple turn'
                                $a.Mission | Should -Be 'FT5103'
                }

                It 'turns a later pin into a reservation (turn time against it)' {
                                $b = New-AsBoard @(
                                                @{ Id = 'R1'; ETD = '2026-03-03 07:30'; Type = 'Day Local'; IP = 'Lee, Ann' },
                                                @{ Id = 'R2'; ETD = '2026-03-03 10:00'; Type = 'Day Local'; IP = 'Ortiz, Sam'; Student = 'Able, Amy' }
                                )
                                $a = Get-Assess $b 'R1' 'Able, Amy' $data
                                $a.Status | Should -Be 'Hard'
                                Get-RuleList $a | Should -Be 'TURN_TIME'
                                (Get-Viol $a 'TURN_TIME').Message | Should -Be 'Turn time 2.5 hr < 3.25 hr minimum (vs 1000)'
                                # nothing is marked complete by a later pin
                                $a.AutoMission | Should -Be 'FT5001'
                }

                It "reserves a later pin's explicit mission so it can't be used earlier" {
                                $b = New-AsBoard @(
                                                @{ Id = 'R1'; ETD = '2026-03-03 07:30'; Type = 'Day Local'; IP = 'Lee, Ann' },
                                                @{ Id = 'R2'; ETD = '2026-03-03 10:00'; Type = 'Day Local'; IP = 'Ortiz, Sam'; Student = 'Able, Amy'; Mission = 'FT5001' }
                                )
                                $a = Get-Assess $b 'R1' 'Able, Amy' $data
                                Get-RuleList $a | Should -Be 'TURN_TIME,NO_ELIGIBLE_EVENT'
                                (Get-Viol $a 'TURN_TIME').Message | Should -Be 'Turn time 2.5 hr < 3.25 hr minimum (vs 1000 FT5001)'
                                (Get-Viol $a 'NO_ELIGIBLE_EVENT').Message | Should -Be 'No eligible Day Local events (syllabus complete or waiting on reserved events)'
                                @($a.EligibleEvents).Count | Should -Be 0
                                $a.Summary | Should -Be 'Able, Amy -> 0730 Day Local: no event, 2 rule flags'
                }

                It 'turns a later pin on the next day into a reservation (CREW_REST_NEXT)' {
                                $b = New-AsBoard -Days @('2026-03-03', '2026-03-04') -Lines @(
                                                @{ Id = 'D1'; ETD = '2026-03-03 18:00'; Type = 'Day Local'; IP = 'Lee, Ann' },
                                                @{ Id = 'D2'; ETD = '2026-03-04 07:00'; Type = 'Day Local'; IP = 'Ortiz, Sam'; Student = 'Able, Amy' }
                                )
                                $a = Get-Assess $b 'D1' 'Able, Amy' $data
                                $a.Status | Should -Be 'Hard'
                                Get-RuleList $a | Should -Be 'CREW_REST_NEXT'
                                (Get-Viol $a 'CREW_REST_NEXT').Message | Should -Be 'Crew rest 13.0 hr < 15 hr before Wed 0700'
                }

                It "checks an O&B target's return leg against a later pin" {
                                $b = New-AsBoard @(@{ Id = 'P1'; ETD = '2026-03-03 12:00'; Type = 'Day Local'; IP = 'Lee, Ann'; Student = 'Able, Amy' })
                                $ob = @(Add-AsQuickEntry -Board $b -Text '0700 O&B Ortiz') | Where-Object { $_.SortieType -eq 'O&B' }
                                $a = Get-Assess $b $ob.Id 'Able, Amy' $data
                                $a.Status | Should -Be 'Hard'
                                Get-RuleList $a | Should -Be 'OB_NOT_ONLY_EVENT,RETURN_TURN_TIME'
                                (Get-Viol $a 'OB_NOT_ONLY_EVENT').Message | Should -Be "O&B must be the day's only event (has 1200)"
                                (Get-Viol $a 'RETURN_TURN_TIME').Message | Should -Be 'Return leg 1000: turn time 2.0 hr < 3.25 hr (vs 1200)'
                }

                It 'flags a pin at the same time as a 0.0 hr turn' {
                                $b = New-AsBoard @(
                                                @{ Id = 'S1'; ETD = '2026-03-03 07:30'; Type = 'Day Local'; IP = 'Ortiz, Sam'; Student = 'Able, Amy' },
                                                @{ Id = 'S2'; ETD = '2026-03-03 07:30'; Type = 'Day Local'; IP = 'Lee, Ann' }
                                )
                                $a = Get-Assess $b 'S2' 'Able, Amy' $data
                                (Get-Viol $a 'TURN_TIME').Message | Should -Be 'Turn time 0.0 hr < 3.25 hr minimum (vs 0730)'
                }

                It "ignores the student's own pin on the target line" {
                                $b = New-AsBoard @(@{ Id = 'O1'; ETD = '2026-03-03 07:30'; Type = 'Day Local'; IP = 'Lee, Ann'; Student = 'Able, Amy' })
                                $a = Get-Assess $b 'O1' 'Able, Amy' $data
                                $a.Status | Should -Be 'Clean'
                                $a.Mission | Should -Be 'FT5001'
                }

                It "does not count other students' pins" {
                                $b = New-AsBoard @(
                                                @{ Id = 'E1'; ETD = '2026-03-03 07:30'; Type = 'Day Local'; IP = 'Lee, Ann'; Student = 'Baker, Ben' },
                                                @{ Id = 'E2'; ETD = '2026-03-03 09:00'; Type = 'Day Local'; IP = 'Ortiz, Sam' }
                                )
                                $a = Get-Assess $b 'E2' 'Able, Amy' $data
                                $a.Status | Should -Be 'Clean'
                                $a.Mission | Should -Be 'FT5001'
                }

                It '-ExcludeLineIds drops the earlier line the student is dragged from' {
                                $b = New-AsBoard @(
                                                @{ Id = 'E1'; ETD = '2026-03-03 07:30'; Type = 'Day Local'; IP = 'Lee, Ann'; Student = 'Able, Amy' },
                                                @{ Id = 'E2'; ETD = '2026-03-03 09:00'; Type = 'Day Local'; IP = 'Ortiz, Sam' }
                                )
                                $with = Get-Assess $b 'E2' 'Able, Amy' $data
                                $moved = Get-Assess $b 'E2' 'Able, Amy' $data -Exclude @('E1')
                                $with.Status | Should -Be 'Hard'
                                $moved.Status | Should -Be 'Clean'
                                @($moved.Violations).Count | Should -Be 0
                                $moved.AutoMission | Should -Be 'FT5001'
                                (@($moved.EligibleEvents) -join ',') | Should -Be 'FT5001'
                }

                It '-ExcludeLineIds drops a later line the student is dragged from' {
                                $b = New-AsBoard @(
                                                @{ Id = 'R1'; ETD = '2026-03-03 07:30'; Type = 'Day Local'; IP = 'Lee, Ann' },
                                                @{ Id = 'R2'; ETD = '2026-03-03 10:00'; Type = 'Day Local'; IP = 'Ortiz, Sam'; Student = 'Able, Amy'; Mission = 'FT5001' }
                                )
                                $a = Get-Assess $b 'R1' 'Able, Amy' $data -Exclude @('R2')
                                $a.Status | Should -Be 'Clean'
                                $a.Mission | Should -Be 'FT5001'
                }

                It '-ExcludeLineIds on an O&B also drops its return leg' {
                                $b = New-AsBoard @(@{ Id = 'T1'; ETD = '2026-03-03 13:30'; Type = 'Day Local'; IP = 'Lee, Ann' })
                                $ob = @(Add-AsQuickEntry -Board $b -Text '0700 O&B Ortiz sp:Able' -Students @($able)) | Where-Object { $_.SortieType -eq 'O&B' }
                                $a = Get-Assess $b 'T1' 'Able, Amy' $data -Exclude @($ob.Id)
                                $a.Status | Should -Be 'Clean'
                                $a.Mission | Should -Be 'FT5001'
                }

                It 'holds the time of a later pin that has no flyable event' {
                                # IFT pin: Able has no IFT event left, but she is still on that line at 0730
                                $b = New-AsBoard @(
                                                @{ Id = 'S0'; ETD = '2026-03-03 06:00'; Type = 'Day Local'; IP = 'Lee, Ann' },
                                                @{ Id = 'S1'; ETD = '2026-03-03 07:30'; Type = 'IFT'; Student = 'Able, Amy' }
                                )
                                $a = Get-Assess $b 'S0' 'Able, Amy' $data
                                (Get-Viol $a 'TURN_TIME').Message | Should -Be 'Turn time 1.5 hr < 3.25 hr minimum (vs 0730)'
                }

                It 'holds the time of an earlier pin that has no flyable event' {
                                # Same IFT pin, target after it: the pin is neither completed (no event) nor
                                # reserved, so it drops out of the timeline and the 1.5 hr turn is not flagged
                                $b = New-AsBoard @(
                                                @{ Id = 'S1'; ETD = '2026-03-03 07:30'; Type = 'IFT'; Student = 'Able, Amy' },
                                                @{ Id = 'S2'; ETD = '2026-03-03 09:00'; Type = 'Day Local'; IP = 'Lee, Ann' }
                                )
                                $a = Get-Assess $b 'S2' 'Able, Amy' $data
                                Get-RuleList $a | Should -Match 'TURN_TIME'
                                (Get-Viol $a 'TURN_TIME').Message | Should -BeLike 'Turn time 1.5 hr < 3.25 hr minimum (vs 0730*'
                                $a.Status | Should -Be 'Hard'
                }
}


# =================================================================================
#  NO MUTATION
# =================================================================================
Describe 'Get-PlacementAssessment: inputs are never changed' {
                BeforeEach { Clear-AssessmentCache }

                It 'leaves the original student objects untouched (JSON snapshot)' {
                                $able = New-TestStudent -Name 'Able, Amy' -SimDebt 1.5 -Done $script:AbleDone -Absences @(New-Absence '2026-03-03 15:00' '2026-03-03 16:00' 'Appt')
                                $lead = New-TestStudent -Name 'Lead, Lou' -Done $script:FormDone
                                $data = New-AsData @($able, $lead)
                                $b = New-AsBoard -Days @('2026-03-03', '2026-03-04') -Lines @(
                                                @{ Id = 'M1'; ETD = '2026-03-03 07:30'; Type = 'Day Local'; IP = 'Lee, Ann'; Student = 'Able, Amy' },
                                                @{ Id = 'M2'; ETD = '2026-03-03 09:00'; Type = 'Day Local'; IP = 'Ortiz, Sam' },
                                                @{ Id = 'M3'; ETD = '2026-03-03 13:00'; Type = 'Day Local'; IP = 'Jones, Bob'; Student = 'Able, Amy'; Mission = 'FT5103' },
                                                @{ Id = 'M4'; ETD = '2026-03-04 07:00'; Type = 'Day Local'; IP = 'Lee, Ann'; Student = 'Able, Amy' },
                                                @{ Id = 'M5'; ETD = '2026-03-03 09:10'; Type = 'Day Local'; IP = 'Lee, Ann'; Student = 'Lead, Lou' }
                                )
                                [void](Add-AsQuickEntry -Board $b -Text '0800 O&B Jones sp:Able' -Day '2026-03-04' -Students @($able, $lead))
                                $before = ConvertTo-Json -InputObject @($data.Students) -Depth 10 -Compress

                                [void](Get-Assess $b 'M2' 'Able, Amy' $data)
                                [void](Get-Assess $b 'M2' 'Able, Amy' $data -Mission 'FT5102')
                                [void](Get-Assess $b 'M1' 'Able, Amy' $data)
                                [void](Get-Assess $b 'M2' 'Able, Amy' $data -Exclude @('M1') -Quick)
                                [void](Get-Assess $b 'M4' 'Able, Amy' $data)
                                [void](Get-Assess $b 'M2' 'Lead, Lou' $data)
                                [void](Get-Assess $b 'M5' 'Lead, Lou' $data)

                                $after = ConvertTo-Json -InputObject @($data.Students) -Depth 10 -Compress
                                $after | Should -Be $before
                                @($able.Reservations).Count | Should -Be 0
                                $able.SimDebt | Should -Be 1.5
                                $able.PSObject.Properties['SyllabusRevision'] | Should -BeNullOrEmpty
                                @($able.Syllabus | Where-Object { $_.IsComplete }).Count | Should -Be 2
                }

                It 'leaves the board lines and Version untouched' {
                                $able = New-TestStudent -Name 'Able, Amy' -Done $script:AbleDone
                                $data = New-AsData @($able)
                                $b = New-AsBoard @(
                                                @{ Id = 'M1'; ETD = '2026-03-03 07:30'; Type = 'Day Local'; IP = 'Lee, Ann'; Student = 'Able, Amy' },
                                                @{ Id = 'M2'; ETD = '2026-03-03 09:00'; Type = 'Day Local'; IP = 'Ortiz, Sam' }
                                )
                                $before = ConvertTo-Json -InputObject @($b.Lines.ToArray()) -Depth 5 -Compress
                                $v = $b.Version
                                $undo = $b.Undo.Count
                                [void](Get-Assess $b 'M2' 'Able, Amy' $data)
                                [void](Get-Assess $b 'M1' 'Able, Amy' $data -Quick)
                                (ConvertTo-Json -InputObject @($b.Lines.ToArray()) -Depth 5 -Compress) | Should -Be $before
                                $b.Version | Should -Be $v
                                $b.Undo.Count | Should -Be $undo
                }
}


# =================================================================================
#  CACHE
# =================================================================================
Describe 'Get-PlacementAssessment: cache by Board.Version' {
                BeforeEach {
                                Clear-AssessmentCache
                                $able = New-TestStudent -Name 'Able, Amy' -Done $script:AbleDone
                                $data = New-AsData @($able)
                                $board = New-AsBoard @(
                                                @{ Id = 'K1'; ETD = '2026-03-03 07:30'; Type = 'Day Local'; IP = 'Lee, Ann' },
                                                @{ Id = 'K2'; ETD = '2026-03-03 09:00'; Type = 'Day Local'; IP = 'Ortiz, Sam' }
                                )
                }

                It 'returns the same object for the same arguments and Version' {
                                $a1 = Get-Assess $board 'K1' 'Able, Amy' $data
                                $a2 = Get-Assess $board 'K1' 'Able, Amy' $data
                                [object]::ReferenceEquals($a1, $a2) | Should -BeTrue
                }

                It 'recomputes after an edit bumps Version (new pin causes a turn-time flag)' {
                                $a1 = Get-Assess $board 'K1' 'Able, Amy' $data
                                $v1 = $board.Version
                                [void](Set-BoardPin -Board $board -Id 'K2' -Student 'Able, Amy')
                                $board.Version | Should -Be ($v1 + 1)
                                $a2 = Get-Assess $board 'K1' 'Able, Amy' $data
                                [object]::ReferenceEquals($a1, $a2) | Should -BeFalse
                                $a2.BoardVersion | Should -Be ($v1 + 1)
                                $a1.Status | Should -Be 'Clean'
                                $a2.Status | Should -Be 'Hard'
                                (Get-Viol $a2 'TURN_TIME').Message | Should -Be 'Turn time 1.5 hr < 3.25 hr minimum (vs 0900)'
                }

                It 'recomputes after an IP change (non-counter penalty appears)' {
                                $a1 = Get-Assess $board 'K1' 'Able, Amy' $data
                                Set-BoardInstructor -Board $board -Id 'K1' -Name 'Leech, Tom'
                                $a2 = Get-Assess $board 'K1' 'Able, Amy' $data
                                $a1.Score | Should -Be 20
                                $a2.Score | Should -Be 5
                                Get-Points $a2.EventDetails[0] 'Non-Counter' | Should -Be -15
                }

                It 'recomputes after Undo-Board' {
                                [void](Set-BoardPin -Board $board -Id 'K2' -Student 'Able, Amy')
                                $a1 = Get-Assess $board 'K1' 'Able, Amy' $data
                                (Undo-Board $board) | Should -BeTrue
                                $a2 = Get-Assess $board 'K1' 'Able, Amy' $data
                                [object]::ReferenceEquals($a1, $a2) | Should -BeFalse
                                $a1.Status | Should -Be 'Hard'
                                $a2.Status | Should -Be 'Clean'
                }

                It 'does not reuse results between two boards at the same Version' {
                                $other = New-AsBoard @(@{ Id = 'K1'; ETD = '2026-03-03 07:30'; Type = 'Day Local'; IP = 'Leech, Tom' })
                                $other.Version | Should -Be $board.Version
                                $a1 = Get-Assess $board 'K1' 'Able, Amy' $data
                                $a2 = Get-Assess $other 'K1' 'Able, Amy' $data
                                $a1.Score | Should -Be 20
                                $a2.Score | Should -Be 5
                }

                It 'keeps separate entries per mission, exclude list and -Quick' {
                                $base = Get-Assess $board 'K1' 'Able, Amy' $data
                                $m = Get-Assess $board 'K1' 'Able, Amy' $data -Mission 'FT5001'
                                $x = Get-Assess $board 'K1' 'Able, Amy' $data -Exclude @('K2')
                                $q = Get-Assess $board 'K1' 'Able, Amy' $data -Quick
                                [object]::ReferenceEquals($base, $m) | Should -BeFalse
                                [object]::ReferenceEquals($base, $x) | Should -BeFalse
                                [object]::ReferenceEquals($base, $q) | Should -BeFalse
                                $m.RequestedMission | Should -Be 'FT5001'
                                $q.Quick | Should -BeTrue
                                $base.Quick | Should -BeFalse
                                [object]::ReferenceEquals($q, (Get-Assess $board 'K1' 'Able, Amy' $data -Quick)) | Should -BeTrue
                }

                It 'Clear-AssessmentCache forces a recompute' {
                                $a1 = Get-Assess $board 'K1' 'Able, Amy' $data
                                Clear-AssessmentCache
                                $a2 = Get-Assess $board 'K1' 'Able, Amy' $data
                                [object]::ReferenceEquals($a1, $a2) | Should -BeFalse
                                $a2.Summary | Should -Be $a1.Summary
                }
}


# =================================================================================
#  QUICK (DRAG HEAT)
# =================================================================================
Describe 'Get-PlacementAssessment: -Quick' {
                BeforeEach {
                                Clear-AssessmentCache
                                $able = New-TestStudent -Name 'Able, Amy' -Done $script:AbleDone
                                $solo = New-TestStudent -Name 'Solo, Sue' -Done $script:SoloDone
                                $hold = New-TestStudent -Name 'Hold, Hal' -OnHold -Done $script:AbleDone -Absences @(New-Absence '2026-03-03 07:00' '2026-03-03 09:00')
                                $data = New-AsData @($able, $solo, $hold)
                                $board = New-AsBoard @(
                                                @{ Id = 'Q1'; ETD = '2026-03-03 07:30'; Type = 'Day Local'; IP = 'Lee, Ann' },
                                                @{ Id = 'Q2'; ETD = '2026-03-03 09:00'; Type = 'Day Local'; IP = 'Ortiz, Sam'; Student = 'Able, Amy' }
                                )
                }

                It 'returns a status without scores, breakdown or eligible-event list' {
                                $q = Get-Assess $board 'Q1' 'Solo, Sue' $data -Quick
                                $q.Quick | Should -BeTrue
                                $q.Status | Should -Be 'Warn'
                                $q.Mission | Should -Be 'FT5103'
                                $q.Score | Should -BeNullOrEmpty
                                $q.BestScore | Should -BeNullOrEmpty
                                @($q.EligibleEvents).Count | Should -Be 0
                                @($q.EventDetails).Count | Should -Be 1
                                @($q.EventDetails[0].Breakdown).Count | Should -Be 0
                                $q.EventDetails[0].Score | Should -BeNullOrEmpty
                                $q.Summary | Should -Be 'Solo, Sue -> 0730 Day Local: FT5103'
                }

                It 'gives the same status as the full assessment: <Name> on <LineId>' -TestCases @(
                                @{ Name = 'Able, Amy'; LineId = 'Q1'; Expected = 'Hard' }
                                @{ Name = 'Solo, Sue'; LineId = 'Q1'; Expected = 'Warn' }
                                @{ Name = 'Solo, Sue'; LineId = 'Q2'; Expected = 'Warn' }
                                @{ Name = 'Hold, Hal'; LineId = 'Q1'; Expected = 'Hard' }
                                @{ Name = 'Baker, Nobody'; LineId = 'Q1'; Expected = 'Hard' }
                ) {
                                param ($Name, $LineId, $Expected)
                                $full = Get-Assess $board $LineId $Name $data
                                $quick = Get-Assess $board $LineId $Name $data -Quick
                                $full.Status | Should -Be $Expected
                                $quick.Status | Should -Be $Expected
                }

                It 'is Clean for a clean placement' {
                                $b2 = New-AsBoard @(@{ Id = 'Q9'; ETD = '2026-03-03 07:30'; Type = 'Day Local'; IP = 'Lee, Ann' })
                                (Get-Assess $b2 'Q9' 'Able, Amy' $data -Quick).Status | Should -Be 'Clean'
                }

                It 'stops at the first student-level violation (fast path)' {
                                $full = Get-Assess $board 'Q1' 'Hold, Hal' $data
                                $quick = Get-Assess $board 'Q1' 'Hold, Hal' $data -Quick
                                Get-RuleList $full | Should -Be 'ON_HOLD,ABSENCE'
                                Get-RuleList $quick | Should -Be 'ON_HOLD'
                                $quick.Summary | Should -Be 'Hold, Hal -> 0730 Day Local: no event, 1 rule flag'
                }
}


# =================================================================================
#  FORMATIONS (LEAD / WING PAIRED ON THE BOARD)
# =================================================================================
Describe 'Get-PlacementAssessment: formation leads and wings' {
                BeforeEach {
                                Clear-AssessmentCache
                                $lead = New-TestStudent -Name 'Lead, Lou' -Done $script:FormDone
                                $chk = New-TestStudent -Name 'Chk, Cal' -Done $script:ChkDone
                                $ift = New-TestStudent -Name 'Ift, Ivy' -Done $script:SoloDone
                                $able = New-TestStudent -Name 'Able, Amy' -Done $script:AbleDone
                                $data = New-AsData @($lead, $chk, $ift, $able)
                                # F2 is pinned, so there is no open Day Local line within 22 min of F1
                                $board = New-AsBoard @(
                                                @{ Id = 'F1'; ETD = '2026-03-03 07:30'; Type = 'Day Local'; IP = 'Ortiz, Sam'; Student = 'Lead, Lou' },
                                                @{ Id = 'F2'; ETD = '2026-03-03 07:40'; Type = 'Day Local'; IP = 'Jones, Bob'; Student = 'Chk, Cal' },
                                                @{ Id = 'F3'; ETD = '2026-03-03 07:45'; Type = 'IFT'; Student = 'Ift, Ivy' },
                                                @{ Id = 'F4'; ETD = '2026-03-03 07:50'; Type = 'OFT' }
                                )
                }

                It 'flags an unpaired lead with no open line as FORM_NO_WING_LINE' {
                                $a = Get-Assess $board 'F1' 'Lead, Lou' $data
                                $a.Status | Should -Be 'Hard'
                                Get-RuleList $a | Should -Be 'FORM_NO_WING_LINE'
                                (Get-Viol $a 'FORM_NO_WING_LINE').Message | Should -Be 'No open Day Local line within 22 min for a wingman/DS'
                                (Get-Detail $a 'FF5201').Passed | Should -BeFalse
                                $a.AutoMission | Should -Be ''
                                $a.Mission | Should -Be 'FF5201'
                }

                It 'does not flag a lead whose wing is paired on the board (Set-FormationPair)' {
                                (Set-FormationPair -Board $board -WingId 'F2' -LeadId 'F1') | Should -BeTrue
                                $a = Get-Assess $board 'F1' 'Lead, Lou' $data
                                Get-RuleList $a | Should -Be ''
                                (Get-Detail $a 'FF5201').Passed | Should -BeTrue
                                $a.Notes -contains 'Direct Support Required' | Should -BeFalse
                                $a.AutoMission | Should -Be 'FF5201'
                                $a.Status | Should -Be 'Clean'
                }

                It 'flags the lead again once the pairing is removed' {
                                [void](Set-FormationPair -Board $board -WingId 'F2' -LeadId 'F1')
                                $paired = Get-Assess $board 'F1' 'Lead, Lou' $data
                                (Set-FormationPair -Board $board -WingId 'F2' -LeadId '') | Should -BeTrue
                                $a = Get-Assess $board 'F1' 'Lead, Lou' $data
                                $paired.Status | Should -Be 'Clean'
                                Get-RuleList $a | Should -Be 'FORM_NO_WING_LINE'
                                $a.Status | Should -Be 'Hard'
                }

                It 'finds an open wing line once the paired wing pin is cleared' {
                                [void](Set-FormationPair -Board $board -WingId 'F2' -LeadId 'F1')
                                Clear-BoardPin -Board $board -Id 'F2'
                                (Get-BoardLine $board 'F2').WingOfLineId | Should -Be ''
                                $a = Get-Assess $board 'F1' 'Lead, Lou' $data
                                # F2 is open now, and Chk can still fly Form there: no flag, no DS note
                                Get-RuleList $a | Should -Be ''
                                @($a.Notes).Count | Should -Be 0
                                $a.AutoMission | Should -Be 'FF5201'
                                $a.Status | Should -Be 'Clean'
                }

                It 'skips the lead check on a wing-paired Day Local line and auto-picks a Form event' {
                                $unpaired = Get-Assess $board 'F2' 'Chk, Cal' $data
                                (@((Get-Detail $unpaired 'FF5201').Violations) | ForEach-Object { $_.Rule }) -join ',' | Should -Be 'FORM_NO_WING_LINE'
                                $unpaired.AutoMission | Should -Be 'FT5190'

                                [void](Set-FormationPair -Board $board -WingId 'F2' -LeadId 'F1')
                                $a = Get-Assess $board 'F2' 'Chk, Cal' $data
                                $ff = Get-Detail $a 'FF5201'
                                $ff.Passed | Should -BeTrue
                                @($ff.Violations).Count | Should -Be 0
                                # the checkride still scores higher, but a wing only flies a Form event
                                (Get-Detail $a 'FT5190').Score | Should -BeGreaterThan $ff.Score
                                $a.AutoMission | Should -Be 'FF5201'
                                $a.Mission | Should -Be 'FF5201'
                                $a.Status | Should -Be 'Clean'
                                (@($a.EligibleEvents) -join ',') | Should -Be 'FF5201,FT5190'
                }

                It 'refuses to pair an IFT or OFT line as a wing' {
                                (Set-FormationPair -Board $board -WingId 'F3' -LeadId 'F1') | Should -BeFalse
                                (Set-FormationPair -Board $board -WingId 'F4' -LeadId 'F1') | Should -BeFalse
                                (Get-BoardLine $board 'F3').WingOfLineId | Should -Be ''
                }

                It 'never lets a forced WingOfLineId on an IFT line satisfy the lead' {
                                # e.g. an old saved board: the field is set without Set-FormationPair
                                (Get-BoardLine $board 'F3').WingOfLineId = 'F1'
                                $a = Get-Assess $board 'F1' 'Lead, Lou' $data
                                Get-RuleList $a | Should -Be 'FORM_NO_WING_LINE'
                                $a.Status | Should -Be 'Hard'
                }

                It 'assesses an IFT line with a forced WingOfLineId as a normal line (IFT event, not a wing)' {
                                (Get-BoardLine $board 'F3').WingOfLineId = 'F1'
                                $a = Get-Assess $board 'F3' 'Ift, Ivy' $data
                                $a.AutoMission | Should -Be 'FN3201'
                                $a.Status | Should -Be 'Clean'
                }

                It 'never lets a forced WingOfLineId on an O&B line satisfy the lead; the O&B is assessed normally' {
                                $ob = @(Add-AsQuickEntry -Board $board -Text '0745 O&B Lee sp:Able' -Students @($able)) | Where-Object { $_.SortieType -eq 'O&B' }
                                (Set-FormationPair -Board $board -WingId $ob.Id -LeadId 'F1') | Should -BeFalse
                                $ob.WingOfLineId = 'F1'
                                $a = Get-Assess $board 'F1' 'Lead, Lou' $data
                                Get-RuleList $a | Should -Be 'FORM_NO_WING_LINE'
                                $w = Get-Assess $board $ob.Id 'Able, Amy' $data
                                $w.AutoMission | Should -Be 'FT5001'
                }

                It 'does not treat a Day Local line paired to an OFT line as a wing' {
                                (Get-BoardLine $board 'F2').WingOfLineId = 'F4'
                                $a = Get-Assess $board 'F2' 'Chk, Cal' $data
                                (@((Get-Detail $a 'FF5201').Violations) | ForEach-Object { $_.Rule }) -join ',' | Should -Be 'FORM_NO_WING_LINE'
                                $a.AutoMission | Should -Be 'FT5190'
                }

                It 'notes Direct Support Required when an open line exists but no other student can fly Form or single' {
                                $b2 = New-AsBoard @(
                                                @{ Id = 'P1'; ETD = '2026-03-03 07:30'; Type = 'Day Local'; IP = 'Ortiz, Sam' },
                                                @{ Id = 'P2'; ETD = '2026-03-03 07:45'; Type = 'Day Local'; IP = 'Lee, Ann' }
                                )
                                $a = Get-Assess $b2 'P1' 'Lead, Lou' (New-AsData @($lead))
                                $a.Status | Should -Be 'Warn'
                                ($a.Notes -join '|') | Should -Be 'Direct Support Required'
                                $a.AutoMission | Should -Be 'FF5201'
                }

                It 'is Warn with an advisory FORM_PREFER_SINGLE when only a single-ship student is available' {
                                $b2 = New-AsBoard @(
                                                @{ Id = 'P1'; ETD = '2026-03-03 07:30'; Type = 'Day Local'; IP = 'Ortiz, Sam' },
                                                @{ Id = 'P2'; ETD = '2026-03-03 07:45'; Type = 'Day Local'; IP = 'Lee, Ann' }
                                )
                                $a = Get-Assess $b2 'P1' 'Lead, Lou' (New-AsData @($lead, $able))
                                $a.Status | Should -Be 'Warn'
                                Get-RuleList $a | Should -Be 'FORM_PREFER_SINGLE'
                                $v = Get-Viol $a 'FORM_PREFER_SINGLE'
                                $v.Severity | Should -Be 'Advisory'
                                $v.Message | Should -Be 'No formation wingman available; another student can fly single-ship here'
                                $a.Mission | Should -Be 'FF5201'
                }
}


# =================================================================================
#  DISPLAY LINES
# =================================================================================
Describe 'Get-AssessmentLines' {
                BeforeEach { Clear-AssessmentCache }

                It 'maps every part of an assessment to its Kind, omitting zero-point items' {
                                $fake = [PSCustomObject]@{
                                                Summary      = 'Title text'
                                                Quick        = $false
                                                Mission      = 'M1'
                                                Violations   = @(
                                                                [PSCustomObject]@{ Rule = 'H'; Severity = 'Hard'; Message = 'hard one' }
                                                                [PSCustomObject]@{ Rule = 'A'; Severity = 'Advisory'; Message = 'advisory one' }
                                                )
                                                Notes        = @('note one')
                                                EventDetails = @(
                                                                [PSCustomObject]@{ Mission = 'M1'; Passed = $true; Score = 10.5; Violations = @(); Notes = @(); Breakdown = @(
                                                                                [PSCustomObject]@{ Label = 'Small'; Points = 3 }
                                                                                [PSCustomObject]@{ Label = 'Zero'; Points = 0 }
                                                                                [PSCustomObject]@{ Label = 'Neg'; Points = -2.5 }
                                                                                [PSCustomObject]@{ Label = 'Big'; Points = 10 }
                                                                ) }
                                                                [PSCustomObject]@{ Mission = 'M2'; Passed = $true; Score = 4; Violations = @(); Notes = @(); Breakdown = @() }
                                                                [PSCustomObject]@{ Mission = 'M3'; Passed = $true; Score = $null; Violations = @(); Notes = @(); Breakdown = @() }
                                                                [PSCustomObject]@{ Mission = 'M4'; Passed = $false; Score = 1; Notes = @(); Breakdown = @(); Violations = @(
                                                                                [PSCustomObject]@{ Rule = 'X'; Severity = 'Hard'; Message = 'why it fails' }
                                                                                [PSCustomObject]@{ Rule = 'Y'; Severity = 'Hard'; Message = 'second reason' }
                                                                ) }
                                                )
                                }
                                $lines = @(Get-AssessmentLines $fake)
                                (Get-LineText $lines) -join "`n" | Should -Be (@(
                                                'Title|Title text',
                                                'Hard|hard one',
                                                'Warn|advisory one',
                                                'Note|note one',
                                                'Plus|Big +10',
                                                'Plus|Small +3',
                                                'Minus|Neg -2.5',
                                                'Info|Other events: M2 4.0; M3 ok; M4 blocked (why it fails)',
                                                "Info|$script:FooterText"
                                ) -join "`n")
                }

                It 'leaves out the footer for a -Quick assessment' {
                                $fake = [PSCustomObject]@{ Summary = 'T'; Quick = $true; Mission = ''; Violations = @(); Notes = @(); EventDetails = @() }
                                (Get-LineText @(Get-AssessmentLines $fake)) -join "`n" | Should -Be 'Title|T'
                }

                It 'returns nothing for $null' {
                                @(Get-AssessmentLines $null).Count | Should -Be 0
                }

                It 'renders a real assessment: plus and minus items, zero items dropped' {
                                Clear-AssessmentCache
                                $s = New-TestStudent -Name 'Debt, Dee' -SimDebt 2 -Done $script:AbleDone
                                $b = New-AsBoard @(@{ Id = 'G1'; ETD = '2026-03-03 07:30'; Type = 'Day Local'; IP = 'Lee, Ann' })
                                $a = Get-Assess $b 'G1' 'Debt, Dee' (New-AsData @($s))
                                # the breakdown holds zero items (Exhaustion, Student Priority) that must not show
                                @($a.EventDetails[0].Breakdown | Where-Object { $_.Points -eq 0 }).Count | Should -BeGreaterThan 0
                                (Get-LineText @(Get-AssessmentLines $a)) -join "`n" | Should -Be (@(
                                                'Title|Debt, Dee -> 0730 Day Local: FT5001 +18.0',
                                                'Plus|Dollar Ride Bonus +20',
                                                'Minus|Sim Debt (flight) -2',
                                                "Info|$script:FooterText"
                                ) -join "`n")
                }

                It 'renders the chosen event first and the other events in the Info line' {
                                Clear-AssessmentCache
                                $chk = New-TestStudent -Name 'Chk, Cal' -Done $script:ChkDone
                                $b = New-AsBoard @(
                                                @{ Id = 'G1'; ETD = '2026-03-03 07:30'; Type = 'Day Local'; IP = 'Jones, Bob' },
                                                @{ Id = 'G2'; ETD = '2026-03-03 07:45'; Type = 'Day Local'; IP = 'Lee, Ann' }
                                )
                                $a = Get-Assess $b 'G1' 'Chk, Cal' (New-AsData @($chk))
                                (Get-LineText @(Get-AssessmentLines $a)) -join "`n" | Should -Be (@(
                                                'Title|Chk, Cal -> 0730 Day Local: FT5190 +49.0',
                                                'Note|Checkride! Requires Formal',
                                                'Plus|Checkride Priority +50',
                                                'Plus|BIT Tracker +6',
                                                'Minus|Break String -5',
                                                'Minus|Exhaustion -2',
                                                'Info|Other events: FF5201 -1.0',
                                                "Info|$script:FooterText"
                                ) -join "`n")
                }

                It 'shows a blocked other event with its first reason' {
                                Clear-AssessmentCache
                                $chk = New-TestStudent -Name 'Chk, Cal' -Done $script:ChkDone
                                $b = New-AsBoard @(@{ Id = 'G1'; ETD = '2026-03-03 07:30'; Type = 'Day Local'; IP = 'Jones, Bob' })
                                $lines = @(Get-AssessmentLines (Get-Assess $b 'G1' 'Chk, Cal' (New-AsData @($chk))))
                                (@($lines | Where-Object { $_.Kind -eq 'Info' }) | ForEach-Object { $_.Text }) -join "`n" | Should -Be (@(
                                                'Other events: FF5201 blocked (No open Day Local line within 22 min for a wingman/DS)',
                                                $script:FooterText
                                ) -join "`n")
                }

                It 'renders hard and advisory violations from a real assessment' {
                                Clear-AssessmentCache
                                $lead = New-TestStudent -Name 'Lead, Lou' -Done $script:FormDone
                                $hold = New-TestStudent -Name 'Hold, Hal' -OnHold -Done $script:AbleDone
                                $b = New-AsBoard @(
                                                @{ Id = 'G1'; ETD = '2026-03-03 07:30'; Type = 'Day Local'; IP = 'Ortiz, Sam' },
                                                @{ Id = 'G2'; ETD = '2026-03-03 07:45'; Type = 'Day Local'; IP = 'Lee, Ann' }
                                )
                                $d = New-AsData @($lead, $hold, (New-TestStudent -Name 'Able, Amy' -Done $script:AbleDone))
                                $warn = @(Get-AssessmentLines (Get-Assess $b 'G1' 'Lead, Lou' $d))
                                $hard = @(Get-AssessmentLines (Get-Assess $b 'G1' 'Hold, Hal' $d))
                                (Get-LineText $warn) -contains 'Warn|No formation wingman available; another student can fly single-ship here' | Should -BeTrue
                                (Get-LineText $hard) -contains 'Hard|Student is on hold' | Should -BeTrue
                                $warn[0].Kind | Should -Be 'Title'
                }
}


# =================================================================================
#  PUCKS
# =================================================================================
Describe 'Get-StudentPuckInfo' {
                BeforeEach {
                                Clear-AssessmentCache
                                $able = New-TestStudent -Name 'Able, Amy' -Done $script:AbleDone -SimDebt 1.5
                                $baker = New-TestStudent -Name 'Baker, Ben' -Done $script:AbleDone
                                $board = New-AsBoard -Days @('2026-03-03', '2026-03-04') -Lines @(
                                                @{ Id = 'U1'; ETD = '2026-03-03 14:00'; Type = 'Day Local'; IP = 'Jones, Bob'; Student = 'Able, Amy' },
                                                @{ Id = 'U2'; ETD = '2026-03-03 15:00'; Type = 'IFT' },
                                                @{ Id = 'U3'; ETD = '2026-03-03 16:00'; Type = 'Day Local'; IP = 'Lee, Ann'; Student = 'Baker, Ben' },
                                                @{ Id = 'U4'; ETD = '2026-03-04 08:00'; Type = 'Day Local'; IP = 'Jones, Bob'; Student = 'Able, Amy' }
                                )
                                [void](Add-AsQuickEntry -Board $board -Text '0700 O&B Jones sp:Able' -Students @($able, $baker))
                }

                It 'shows every flag badge in order' {
                                $s = New-TestStudent -Name 'All, Al' -IsCap -IsInternational -OnHold -UnsatSim -UnsatFlight
                                $p = Get-StudentPuckInfo -Student $s -Board $board -Day '2026-03-03'
                                ($p.Badges -join ',') | Should -Be 'CAP,INTL,HOLD,U-SIM,U-FLT'
                                $p.Blocked | Should -BeTrue
                                $p.Name | Should -Be 'All, Al'
                }

                It 'reads text flags through ConvertTo-FaisBool' {
                                $s = New-TestStudent -Name 'Text, Tia'
                                $s.IsCap = 'TRUE'
                                $s.IsInternational = 'FALSE'
                                $s.OnHold = ''
                                $p = Get-StudentPuckInfo -Student $s -Board $board -Day '2026-03-03'
                                ($p.Badges -join ',') | Should -Be 'CAP'
                                $p.Blocked | Should -BeFalse
                }

                It 'has no badges for a plain student' {
                                $p = Get-StudentPuckInfo -Student $baker -Board $board -Day '2026-03-03'
                                @($p.Badges).Count | Should -Be 0
                                @($p.Absences).Count | Should -Be 0
                                $p.Blocked | Should -BeFalse
                }

                It 'shows an ABS hhmm-hhmm badge for an absence on the day, with the reason in Absences' {
                                $s = New-TestStudent -Name 'Abs, Ann' -Absences @(New-Absence '2026-03-03 09:00' '2026-03-03 11:00' 'Dentist')
                                $p = Get-StudentPuckInfo -Student $s -Board $board -Day '2026-03-03'
                                ($p.Badges -join ',') | Should -Be 'ABS 0900-1100'
                                ($p.Absences -join ',') | Should -Be '0900-1100 Dentist'
                }

                It 'clips a multi-day absence to the day (<Day> -> <Expected>)' -TestCases @(
                                @{ Day = '2026-03-02'; Expected = 'ABS 1500-2359' }
                                @{ Day = '2026-03-03'; Expected = 'ABS 0000-2359' }
                                @{ Day = '2026-03-04'; Expected = 'ABS 0000-0800' }
                                @{ Day = '2026-03-05'; Expected = '' }
                ) {
                                param ($Day, $Expected)
                                $s = New-TestStudent -Name 'Abs, Ann' -Absences @(New-Absence '2026-03-02 15:00' '2026-03-04 08:00')
                                $p = Get-StudentPuckInfo -Student $s -Board $board -Day $Day
                                ($p.Badges -join ',') | Should -Be $Expected
                }

                It 'trims the absence text when there is no reason' {
                                $s = New-TestStudent -Name 'Abs, Ann' -Absences @(New-Absence '2026-03-03 09:00' '2026-03-03 11:00')
                                $p = Get-StudentPuckInfo -Student $s -Board $board -Day '2026-03-03'
                                ($p.Absences -join ',') | Should -Be '0900-1100'
                }

                It 'puts flag badges before absence badges' {
                                $s = New-TestStudent -Name 'All, Al' -IsCap -Absences @(New-Absence '2026-03-03 09:00' '2026-03-03 11:00')
                                $p = Get-StudentPuckInfo -Student $s -Board $board -Day '2026-03-03'
                                ($p.Badges -join ',') | Should -Be 'CAP,ABS 0900-1100'
                }

                It 'counts turns on the day, an O&B as 2' {
                                (Get-StudentPuckInfo -Student $able -Board $board -Day '2026-03-03').Turns | Should -Be 3
                                (Get-StudentPuckInfo -Student $able -Board $board -Day '2026-03-04').Turns | Should -Be 1
                                (Get-StudentPuckInfo -Student $baker -Board $board -Day '2026-03-03').Turns | Should -Be 1
                                (Get-StudentPuckInfo -Student $baker -Board $board -Day '2026-03-04').Turns | Should -Be 0
                }

                It 'counts 0 turns without a board' {
                                (Get-StudentPuckInfo -Student $able -Board $null -Day '2026-03-03').Turns | Should -Be 0
                }

                It 'shows the next one or two opted events, or complete' {
                                (Get-StudentPuckInfo -Student (New-TestStudent -Name 'New, Ned') -Board $board -Day '2026-03-03').Next | Should -Be 'FI3101'
                                $s = New-TestStudent -Name 'Two, Tom' -Done 'FI3101|2026-02-20 08:00'
                                (Get-StudentPuckInfo -Student $s -Board $board -Day '2026-03-03').Next | Should -Be 'FE3102 / FT5001'
                                $all = @('FI3101', 'FE3102', 'FT5001', 'FT5102', 'FT5103', 'FN3201', 'FF5201', 'FF5202', 'FT5104', 'FT5190') | ForEach-Object { "$_|2026-02-20 08:00" }
                                $done = New-TestStudent -Name 'Done, Dan' -Done $all
                                (Get-StudentPuckInfo -Student $done -Board $board -Day '2026-03-03').Next | Should -Be 'complete'
                }

                It 'reports SimDebt as a number' {
                                (Get-StudentPuckInfo -Student $able -Board $board -Day '2026-03-03').SimDebt | Should -Be 1.5
                                $s = New-TestStudent -Name 'Blank, Bea'
                                $s.SimDebt = ''
                                (Get-StudentPuckInfo -Student $s -Board $board -Day '2026-03-03').SimDebt | Should -Be 0
                }
}

Describe 'Get-InstructorPuckInfo' {
                BeforeEach {
                                Clear-AssessmentCache
                                $board = New-AsBoard -Days @('2026-03-03', '2026-03-04') -Lines @(
                                                @{ Id = 'I1'; ETD = '2026-03-03 14:00'; Type = 'Day Local'; IP = 'Jones, Bob' },
                                                @{ Id = 'I2'; ETD = '2026-03-03 16:00'; Type = 'Day Local'; IP = 'Lee, Ann' },
                                                @{ Id = 'I3'; ETD = '2026-03-04 08:00'; Type = 'Day Local'; IP = 'Jones, Bob' }
                                )
                                # O&B at 0700 with its linked return at 1000, both Jones
                                [void](Add-AsQuickEntry -Board $board -Text '0700 O&B Jones')
                }

                It "shows badges '<Badges>' for <Name>" -TestCases @(
                                @{ Name = 'Jones, Bob'; Badges = 'EXP,CHK' }
                                @{ Name = 'Lee, Ann'; Badges = '' }
                                @{ Name = 'Leech, Tom'; Badges = 'EXP,NC' }
                                @{ Name = 'Ortiz, Sam'; Badges = 'EXP' }
                ) {
                                param ($Name, $Badges)
                                $ip = $script:IPs | Where-Object { $_.Name -eq $Name }
                                $p = Get-InstructorPuckInfo -Instructor $ip -Board $board -Day '2026-03-03'
                                ($p.Badges -join ',') | Should -Be $Badges
                                $p.Name | Should -Be $Name
                }

                It 'reads text flags and treats a blank IsCounter as a counter' {
                                $blank = [PSCustomObject]@{ Name = 'Blank, Bo'; IsCounter = ''; IsExperienced = 'TRUE'; IsCheck = 'no'; Callsign = ''; Num = '' }
                                $text = [PSCustomObject]@{ Name = 'Text, Ty'; IsCounter = 'FALSE'; IsExperienced = 'FALSE'; IsCheck = 'TRUE'; Callsign = 'VIPER'; Num = '' }
                                ((Get-InstructorPuckInfo -Instructor $blank -Board $null -Day '2026-03-03').Badges -join ',') | Should -Be 'EXP'
                                ((Get-InstructorPuckInfo -Instructor $text -Board $null -Day '2026-03-03').Badges -join ',') | Should -Be 'CHK,NC'
                }

                It 'joins callsign and number, trimming a blank part' {
                                (Get-InstructorPuckInfo -Instructor $script:IPs[0] -Board $board -Day '2026-03-03').Callsign | Should -Be 'TALON 31'
                                $text = [PSCustomObject]@{ Name = 'Text, Ty'; IsCounter = $true; IsExperienced = $false; IsCheck = $false; Callsign = 'VIPER'; Num = '' }
                                (Get-InstructorPuckInfo -Instructor $text -Board $null -Day '2026-03-03').Callsign | Should -Be 'VIPER'
                }

                It 'counts the day''s lines in ETD order without the O&B return line' {
                                $p = Get-InstructorPuckInfo -Instructor $script:IPs[0] -Board $board -Day '2026-03-03'
                                @($board.Lines | Where-Object { $_.Day -eq '2026-03-03' -and $_.Instructor -eq 'Jones, Bob' }).Count | Should -Be 3
                                $p.Lines | Should -Be 2
                                ($p.Times -join ',') | Should -Be '0700,1400'
                }

                It 'counts only the requested day' {
                                $p = Get-InstructorPuckInfo -Instructor $script:IPs[0] -Board $board -Day '2026-03-04'
                                $p.Lines | Should -Be 1
                                ($p.Times -join ',') | Should -Be '0800'
                                (Get-InstructorPuckInfo -Instructor $script:IPs[1] -Board $board -Day '2026-03-03').Lines | Should -Be 1
                                (Get-InstructorPuckInfo -Instructor $script:IPs[2] -Board $board -Day '2026-03-03').Lines | Should -Be 0
                }

                It 'counts 0 lines without a board' {
                                $p = Get-InstructorPuckInfo -Instructor $script:IPs[0] -Board $null -Day '2026-03-03'
                                $p.Lines | Should -Be 0
                                @($p.Times).Count | Should -Be 0
                }
}
