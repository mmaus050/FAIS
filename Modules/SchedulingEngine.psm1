# SCHEDULINGENGINE.PSM1 DIRECTORY:
#             Contains all functions that run the logic for the scheduling
# Test-HardRules          (every hard rule, with a reason for each failure)
# SoftRules               (score for one student on one event on one line)
# Get-FormationLineContext / Get-LinkedReturnRow
# Process-HardcodedSortie (pinned / hardcoded lines)
# Pair-FormationWingman   (wingman or direct support for a formation lead)
# Schedule-ReturnLeg      (O&B second leg)
# Invoke-FillSchedule     (fills every line that is not hardcoded)
# Trace-SchedulingDecision


# =================================================================================
#  SMALL HELPERS
# =================================================================================

# Reads a SoftWeights value as a number, with the legacy default when blank
function Get-Weight {
                param (
                                $Config,
                                [string]$Name,
                                [double]$Default = 0
                )

                $v = Get-OptionalProperty $Config $Name $null
                if ($null -eq $v -or [string]$v -eq '') { return $Default }
                $num = 0.0
                if ([double]::TryParse(([string]$v).Trim(), [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$num)) { return $num }
                return $Default
}

function New-RuleViolation {
                param (
                                [string]$Rule,
                                [string]$Message,
                                [ValidateSet('Hard', 'Advisory')] [string]$Severity = 'Hard',
                                [string]$Mission = '',
                                [bool]$Blocks = $true
                )

                return [PSCustomObject]@{
                                Rule     = $Rule
                                Severity = $Severity
                                Message  = $Message
                                Mission  = $Mission
                                Blocks   = $Blocks
                }
}

function Test-AnyBlocking {
                param ($Violations)

                if ($null -eq $Violations) { return $false }
                foreach ($v in $Violations) { if ($null -ne $v -and $v.Blocks) { return $true } }
                return $false
}

function Format-Hours {
                param ([double]$Hours)
                return $Hours.ToString('0.0', [System.Globalization.CultureInfo]::InvariantCulture)
}

# 2.5, 3, 0.25 (no trailing zeros)
function Format-Number {
                param ([double]$Value)
                return $Value.ToString('0.##', [System.Globalization.CultureInfo]::InvariantCulture)
}

# "1 Flight", "2 Flights"
function Format-Count {
                param ([int]$Count, [string]$Word)
                if ($Count -eq 1) { return "1 $Word" }
                return "$Count $($Word)s"
}

# IP last name for score reasons: 'Jones, Bob [EXP]' -> 'Jones'
function Get-ShortIPName {
                param ([string]$Name)

                $base = Get-BaseName $Name
                if ($base -match '^([^,]+),') { return $Matches[1].Trim() }
                return $base
}

# How many of the newest flights in a row were with this IP ($Flights newest first)
function Get-IPRunLength {
                param ([array]$Flights, [string]$IPBase)

                $n = 0
                foreach ($f in $Flights) {
                                if ((Get-BaseName (Get-NameText $f.Instructor)) -ne $IPBase) { break }
                                $n++
                }
                return $n
}


# =================================================================================
#  CONTEXT LOOKUPS
# =================================================================================

# Finds the O&B Return line that belongs to an O&B: by link first, otherwise the
# legacy lookup (same day, same IP, still open).
function Get-LinkedReturnRow {
                param (
                                $Sortie,
                                [array]$Schedule,
                                [string]$StudentName = ''
                )

                $linked = [string](Get-OptionalProperty $Sortie 'LinkedLineId' '')
                if ($linked -ne '') {
                                foreach ($r in $Schedule) {
                                                if ([string](Get-OptionalProperty $r 'BoardLineId' '') -ne $linked) { continue }
                                                # A line retyped to Day Local / sim, or pinned to someone else, is no longer the return leg
                                                $who = Get-NameText $r.Student
                                                if ((Test-IsReturnSortieType ([string]$r.SortieType)) -and ($who -eq '' -or $who -ieq $StudentName -or (Test-IsSentinelStudent $who))) { return $r }
                                }
                }

                $start = Parse-ETD $Sortie.ETD
                if ($null -eq $start) { return $null }
                $ipBase = Get-BaseName (Get-NameText $Sortie.Instructor)

                $candidates = @(
                                foreach ($r in $Schedule) {
                                                if (-not (Test-IsReturnSortieType ([string]$r.SortieType))) { continue }
                                                if ($null -ne (Get-OptionalProperty $r 'FilledBy' $null) -and [string]$r.FilledBy -ne '') { continue }
                                                $rs = Parse-ETD $r.ETD
                                                if ($null -eq $rs -or $rs.Date -ne $start.Date) { continue }
                                                $who = Get-NameText $r.Student
                                                if ($who -ne '' -and $who -ine $StudentName) { continue }
                                                if ((Get-BaseName (Get-NameText $r.Instructor)) -ine $ipBase) { continue }
                                                [PSCustomObject]@{ Row = $r; Start = $rs; After = ($rs -ge $start) }
                                }
                )
                if ($candidates.Count -eq 0) { return $null }
                $pick = $candidates | Sort-Object -Property @{ Expression = { -not $_.After } }, Start | Select-Object -First 1
                return $pick.Row
}

# Open live-fly lines a formation lead could take a wingman / DS from, and which
# other students could fly Form or single-ship events on this line.
function Get-FormationLineContext {
                param (
                                $Sortie,
                                [array]$Schedule,
                                [array]$Students,
                                [double]$FormLookAhead = 22
                )

                $start = Parse-ETD $Sortie.ETD
                $open = New-Object System.Collections.Generic.List[object]
                if ($null -ne $start) {
                                $end = $start.AddMinutes($FormLookAhead)
                                foreach ($r in $Schedule) {
                                                if (Test-SameSortie $r $Sortie) { continue }
                                                if (-not (Test-IsOpenWingLine $r)) { continue }
                                                $rs = Parse-ETD $r.ETD
                                                if ($null -ne $rs -and $rs -ge $start -and $rs -le $end) { $open.Add($r) }
                                }
                }

                $formCapable = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
                $singleCapable = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
                foreach ($s in @($Students)) {
                                if ([string]::IsNullOrWhiteSpace([string]$s.Name)) { continue }
                                foreach ($evt in @(Get-ValidOptedEvents -Student $s -Sortie $Sortie)) {
                                                $t = [string]$evt.Type
                                                if ($t -like '*Form*') { [void]$formCapable.Add([string]$s.Name) }
                                                elseif ($t -eq 'Single' -or $t -eq 'SingleSolo') { [void]$singleCapable.Add([string]$s.Name) }
                                }
                }

                return [PSCustomObject]@{
                                OpenWingLines = $open.ToArray()
                                FormCapable   = $formCapable
                                SingleCapable = $singleCapable
                }
}

# A line a wingman or DS can be placed on: open, live-fly Day Local (never OFT/IFT, never an O&B leg)
function Test-IsOpenWingLine {
                param ($Row)

                if ((Get-NameText $Row.Student) -ne '') { return $false }
                if ([string](Get-OptionalProperty $Row 'FilledBy' '') -ne '') { return $false }
                return (Test-IsWingCapableType ([string]$Row.SortieType))
}

# The line the board paired as wingman to this lead (WingOfLineId), or $null.
# Only counts when both lines can fly formation (OFT / IFT / O&B lines never can).
function Get-PinnedWingRow {
                param ($Sortie, [array]$Schedule)

                $myId = [string](Get-OptionalProperty $Sortie 'BoardLineId' '')
                if ($myId -eq '' -or -not (Test-IsWingCapableType ([string]$Sortie.SortieType))) { return $null }
                foreach ($r in $Schedule) {
                                if ([string](Get-OptionalProperty $r 'WingOfLineId' '') -eq $myId -and (Test-IsWingCapableType ([string]$r.SortieType))) { return $r }
                }
                return $null
}

# True when this line is pinned as a wingman to another line the board paired it with
function Test-IsPairedWing {
                param ($Sortie, [array]$Schedule)

                $leadId = [string](Get-OptionalProperty $Sortie 'WingOfLineId' '')
                if ($leadId -eq '' -or -not (Test-IsWingCapableType ([string]$Sortie.SortieType))) { return $false }
                foreach ($r in $Schedule) {
                                if ([string](Get-OptionalProperty $r 'BoardLineId' '') -eq $leadId) { return (Test-IsWingCapableType ([string]$r.SortieType)) }
                }
                return $false
}

function Test-SetHasOther {
                param ($Set, [string]$Self)

                foreach ($n in $Set) { if ($n -ine $Self) { return $true } }
                return $false
}


# =================================================================================
#  _   _               _   ____        _
# | | | | __ _ _ __ __| | |  _ \ _   _| | ___  ___
# | |_| |/ _` | '__/ _` | | |_) | | | | |/ _ \/ __|
# |  _  | (_| | | | (_| | |  _ <| |_| | |  __/\__ \
# |_| |_|\__,_|_|  \__,_| |_| \_\\__,_|_|\___||___/
#
# Test-HardRules never changes the student, the sortie or the syllabus.
# Without -Explain it stops at the first student-level failure (fast path for the
# fill and for drag heat). With -Explain it evaluates every rule so the board can
# list all of them.
# =================================================================================
function Test-HardRules {
                param (
                                [Parameter(Mandatory)] $Student,
                                [Parameter(Mandatory)] $Sortie,
                                [array]$Schedule = @(),
                                [array]$Students = @(),
                                [array]$Instructors = @(),
                                [double]$FormLookAhead = 22,
                                $Mission,
                                [switch]$IsWingCandidate,
                                [switch]$Explain,
                                [hashtable]$LineContext
                )

                $violations = New-Object System.Collections.Generic.List[object]
                $notes = New-Object System.Collections.Generic.List[string]
                $eventResults = New-Object System.Collections.Generic.List[object]

                $start = Parse-ETD $Sortie.ETD
                $sortieType = [string]$Sortie.SortieType
                $isSimLine = Test-IsSimSortieType $sortieType
                $isOB = Test-IsOBSortieType $sortieType
                $isDayLocal = $sortieType -like '*Day Local*'
                $ip = Resolve-SortieInstructor -Instructors $Instructors -Instructor $Sortie.Instructor
                $ipName = Get-NameText $ip
                $ipBase = Get-BaseName $ipName
                $ipIsSim = $ipName -like '*SIM IP*'
                $lineId = [string](Get-OptionalProperty $Sortie 'BoardLineId' '')
                $studentName = [string]$Student.Name

                if ($null -eq $start) {
                                $violations.Add((New-RuleViolation 'BAD_ETD' "Line has no valid ETD ('$($Sortie.ETD)')"))
                                return (New-HardRuleResult $violations $notes $eventResults)
                }

                $timeline = @(Get-StudentTimeline -Student $Student -ExceptLineId $lineId)
                $today = @($timeline | Where-Object { $_.Start.Date -eq $start.Date })
                $stop = { (-not $Explain) -and (Test-AnyBlocking $violations) }

                # ---------------- O&B RETURN LINES ----------------
                # Return legs are filled from their O&B, never on their own
                if (Test-IsReturnSortieType $sortieType) {
                                $violations.Add((New-RuleViolation 'RETURN_LINE' 'O&B Return lines are filled from their O&B'))
                                if (& $stop) { return (New-HardRuleResult $violations $notes $eventResults) }
                }

                # ---------------- HOLD STATUS ----------------
                if (ConvertTo-FaisBool $Student.OnHold) {
                                $violations.Add((New-RuleViolation 'ON_HOLD' 'Student is on hold'))
                                if (& $stop) { return (New-HardRuleResult $violations $notes $eventResults) }
                }

                # ---------------- STUDENT UNAVAILABILITY / ABSENCES ----------------
                # The sortie "really begins" 1.5 hr before the ETD and "ends" 2.5 hr after it.
                # Example: a 1000 sortie occupies ~0830-1230 for the absence check.
                $windowStart = $start.AddHours(-1.5)
                $windowEnd = $start.AddHours(2.5)
                foreach ($absence in @($Student.Absences)) {
                                if ($null -eq $absence -or $null -eq $absence.StartTime -or $null -eq $absence.EndTime) { continue }
                                if ($windowStart -lt $absence.EndTime -and $windowEnd -gt $absence.StartTime) {
                                                $sameDay = ($absence.StartTime.Date -eq $start.Date -and $absence.EndTime.Date -eq $start.Date)
                                                $fmt = if ($sameDay) { 'HHmm' } else { 'dd MMM HHmm' }
                                                $reason = if ([string]::IsNullOrWhiteSpace([string]$absence.Reason)) { '' } else { " ($($absence.Reason))" }
                                                $violations.Add((New-RuleViolation 'ABSENCE' ("Unavailable {0}-{1}{2}; sortie window {3}-{4}" -f $absence.StartTime.ToString($fmt), $absence.EndTime.ToString($fmt), $reason, $windowStart.ToString('HHmm'), $windowEnd.ToString('HHmm'))))
                                                break
                                }
                }
                if (& $stop) { return (New-HardRuleResult $violations $notes $eventResults) }

                # ---------------- CAPSTERS ----------------
                $isCap = ConvertTo-FaisBool $Student.IsCap
                $isIntl = ConvertTo-FaisBool $Student.IsInternational
                if ($isCap -and $sortieType -like '*O&B*') {
                                $violations.Add((New-RuleViolation 'CAP_OB' "CAP students can't fly O&Bs"))
                                if (& $stop) { return (New-HardRuleResult $violations $notes $eventResults) }
                }
                # CAPsters single turn ONLY (flights)
                if ($isCap -and -not $isSimLine) {
                                $flights = @($today | Where-Object { -not $_.IsSim })
                                if ($flights.Count -ge 1) {
                                                $violations.Add((New-RuleViolation 'CAP_SINGLE_TURN' ("CAP student already flies today ({0})" -f ((@($flights | ForEach-Object { Format-EventLabel $_.Start $_.Mission })) -join ', '))))
                                                if (& $stop) { return (New-HardRuleResult $violations $notes $eventResults) }
                                }
                }

                # ---------------- TRIPLE TURN PREVENTION ----------------
                $todayLabels = (@($today | ForEach-Object { Format-EventLabel $_.Start $_.Mission })) -join ', '
                if ($today.Count -ge 2) {
                                $violations.Add((New-RuleViolation 'TRIPLE_TURN' "Already has $($today.Count) events today ($todayLabels) - triple turn"))
                                if (& $stop) { return (New-HardRuleResult $violations $notes $eventResults) }
                }
                elseif ($isOB -and $today.Count -ge 1) {
                                $violations.Add((New-RuleViolation 'OB_NOT_ONLY_EVENT' "O&B must be the day's only event (has $todayLabels)"))
                                if (& $stop) { return (New-HardRuleResult $violations $notes $eventResults) }
                }

                # ---------------- TURN TIME, DUTY DAY, CREW REST ----------------
                $returnStart = $null
                if ($isOB) {
                                $retRow = Get-LinkedReturnRow -Sortie $Sortie -Schedule $Schedule -StudentName $studentName
                                if ($null -ne $retRow) { $returnStart = Parse-ETD $retRow.ETD }
                }

                foreach ($evt in $today) {
                                $label = Format-EventLabel $evt.Start $evt.Mission
                                $tt = [Math]::Abs(($start - $evt.Start).TotalHours)
                                if ($tt -lt 3.25) {
                                                $violations.Add((New-RuleViolation 'TURN_TIME' ("Turn time {0} hr < 3.25 hr minimum (vs {1})" -f (Format-Hours $tt), $label)))
                                }
                                elseif ($tt -lt 3.5) {
                                                [void]$notes.Add('Less than 3.5 hr TT, Flt/CC Approval Required')
                                }
                                elseif ($tt -ge 6 -and $tt -le 9) {
                                                [void]$notes.Add('If this is a Night Flight SP will not have Duty Day')
                                }
                                elseif ($tt -gt 9) {
                                                $early = if ($evt.Start -lt $start) { $evt.Start } else { $start }
                                                $late = if ($evt.Start -lt $start) { $start } else { $evt.Start }
                                                $violations.Add((New-RuleViolation 'DUTY_DAY' ("Duty day: {0} hr between {1} and {2} events (> 9 hr max)" -f (Format-Hours $tt), $early.ToString('HHmm'), $late.ToString('HHmm'))))
                                }

                                if ($null -ne $returnStart) {
                                                $rtt = [Math]::Abs(($returnStart - $evt.Start).TotalHours)
                                                if ($rtt -lt 3.25) {
                                                                $violations.Add((New-RuleViolation 'RETURN_TURN_TIME' ("Return leg {0}: turn time {1} hr < 3.25 hr (vs {2})" -f $returnStart.ToString('HHmm'), (Format-Hours $rtt), $label)))
                                                }
                                                elseif ($rtt -lt 3.5) {
                                                                [void]$notes.Add('Return Leg: Less than 3.5 hr TT, Flt/CC Approval Required')
                                                }
                                                elseif ($rtt -ge 6 -and $rtt -le 9) {
                                                                [void]$notes.Add('Return Leg: Night Flight Warning')
                                                }
                                                elseif ($rtt -gt 9) {
                                                                $violations.Add((New-RuleViolation 'RETURN_DUTY_DAY' ("Return leg {0}: duty day {1} hr (> 9 hr max, vs {2})" -f $returnStart.ToString('HHmm'), (Format-Hours $rtt), $label)))
                                                }
                                }
                }
                if (& $stop) { return (New-HardRuleResult $violations $notes $eventResults) }

                # Crew rest from the previous day's last event
                $before = @($timeline | Where-Object { $_.Start.Date -lt $start.Date })
                if ($before.Count -gt 0) {
                                $last = $before[$before.Count - 1]
                                if ($last.Start.Date -eq $start.Date.AddDays(-1)) {
                                                $rest = ($start - $last.Start).TotalHours
                                                if ($rest -lt 15) {
                                                                $violations.Add((New-RuleViolation 'CREW_REST' ("Crew rest {0} hr < 15 hr since {1}" -f (Format-Hours $rest), (Format-EventLabel $last.Start $last.Mission -WithDay))))
                                                }
                                }
                }
                # Crew rest before the next day's first event (multi-day boards)
                $after = @($timeline | Where-Object { $_.Start.Date -gt $start.Date })
                if ($after.Count -gt 0) {
                                $next = $after[0]
                                $dayEnd = if ($null -ne $returnStart -and $returnStart -gt $start) { $returnStart } else { $start }
                                # A later event the same day is checked on its own line, so only the day's last event answers for tomorrow
                                $laterToday = @($today | Where-Object { $_.Start -gt $dayEnd })
                                if ($laterToday.Count -eq 0 -and $next.Start.Date -eq $start.Date.AddDays(1)) {
                                                $rest = ($next.Start - $dayEnd).TotalHours
                                                if ($rest -lt 15) {
                                                                $violations.Add((New-RuleViolation 'CREW_REST_NEXT' ("Crew rest {0} hr < 15 hr before {1}" -f (Format-Hours $rest), (Format-EventLabel $next.Start $next.Mission -WithDay))))
                                                }
                                }
                }
                if (& $stop) { return (New-HardRuleResult $violations $notes $eventResults) }

                # ---------------- CAP / INTERNATIONAL NEED EXPERIENCED IPs ----------------
                if (($isCap -or $isIntl) -and ($isDayLocal -or $sortieType -like '*O&B*')) {
                                if (-not (ConvertTo-FaisBool $ip.IsExperienced)) {
                                                $who = if ($isCap) { 'CAP' } else { 'International' }
                                                $violations.Add((New-RuleViolation 'EXP_IP' "$who student needs an experienced IP ($ipBase isn't)"))
                                                if (& $stop) { return (New-HardRuleResult $violations $notes $eventResults) }
                                }
                }

                # ---------------- TRAINING FOLLOWING UNSAT EVENT ----------------
                # Cleared on the student once a line that cleans it up is filled (Clear-CleanedUnsat)
                if ((ConvertTo-FaisBool $Student.UnsatSim) -and ($isDayLocal -or $sortieType -like '*O&B*')) {
                                $need = if ((Get-UnsatSimKind $Student) -eq 'IFT') { 'an OFT or IFT' } else { 'an OFT' }
                                $violations.Add((New-RuleViolation 'UNSAT_SIM' "Unsat sim on record - no Day Local/O&B until a re-sim in $need"))
                }
                if ((ConvertTo-FaisBool $Student.UnsatFlight) -and $isSimLine) {
                                $violations.Add((New-RuleViolation 'UNSAT_FLIGHT' 'Unsat flight on record - no OFT/IFT until re-fly'))
                }
                if (& $stop) { return (New-HardRuleResult $violations $notes $eventResults) }

                # ---------------- COLLECT EVENTS TO TEST ----------------
                $valid = @(Get-ValidOptedEvents -Student $Student -Sortie $Sortie)
                $eventsToTest = New-Object System.Collections.Generic.List[object]
                $explicitName = Get-NameText $Mission
                $explicitIneligible = $null

                if ($explicitName -ne '') {
                                $explicit = Find-SyllabusMission -Student $Student -Mission $explicitName
                                if ($null -eq $explicit) {
                                                $violations.Add((New-RuleViolation 'MISSION_NOT_FOUND' "$explicitName isn't in $studentName's syllabus" -Mission $explicitName))
                                }
                                else {
                                                $eventsToTest.Add($explicit)
                                                $isValid = $false
                                                foreach ($v in $valid) { if ([string]$v.Mission -ieq [string]$explicit.Mission) { $isValid = $true; break } }
                                                if (-not $isValid) {
                                                                # Not opted (already flown, or a prereq still to fly) is its own flag; anything
                                                                # else is about this line (wrong type, pinned elsewhere, no O&B second leg)
                                                                $notOpted = Get-NotOptedReason -Student $Student -Mission $explicit
                                                                if ($notOpted -ne '') {
                                                                                $explicitIneligible = New-RuleViolation 'NOT_OPTED' ("Not opted for {0}: {1}" -f $explicit.Mission, $notOpted) -Mission ([string]$explicit.Mission)
                                                                }
                                                                else {
                                                                                $explicitIneligible = New-RuleViolation 'MISSION_NOT_ELIGIBLE' ("{0} isn't eligible here ({1})" -f $explicit.Mission, (Get-IneligibleReason -Student $Student -Mission $explicit -Sortie $Sortie)) -Mission ([string]$explicit.Mission)
                                                                }
                                                }
                                }
                }
                else {
                                if ($valid.Count -eq 0) {
                                                $violations.Add((New-RuleViolation 'NO_ELIGIBLE_EVENT' (Get-NoEligibleMessage -Student $Student -Sortie $Sortie)))
                                }
                                foreach ($v in $valid) { $eventsToTest.Add($v) }
                }

                $completedReal = @($timeline | Where-Object { -not $_.IsReservation })
                # A wingman the board already paired to this line satisfies the formation-lead checks
                $pinnedWing = if ($IsWingCandidate) { $null } else { Get-PinnedWingRow -Sortie $Sortie -Schedule $Schedule }

                foreach ($evt in $eventsToTest) {
                                $ev = New-Object System.Collections.Generic.List[object]
                                $en = New-Object System.Collections.Generic.List[string]
                                $mName = ([string]$evt.Mission).Trim()
                                $mType = ([string]$evt.Type).Trim()
                                $mIsCheck = ConvertTo-FaisBool $evt.IsCheck

                                if ($null -ne $explicitIneligible) { $ev.Add($explicitIneligible) }

                                # Formations are blocked from O&Bs (temporary rule)
                                if ($mType -like '*Form*' -and $sortieType -like '*O&B*') {
                                                $ev.Add((New-RuleViolation 'FORM_ON_OB' "Formation events can't be flown on an O&B" -Mission $mName))
                                }
                                # Checkrides are blocked from O&Bs entirely
                                if ($sortieType -like '*O&B*' -and $mIsCheck) {
                                                $ev.Add((New-RuleViolation 'CHECK_ON_OB' "Checkrides can't be flown on an O&B" -Mission $mName))
                                }
                                # O&B second leg has to exist and can't be a checkride
                                if ($isOB) {
                                                $reserved = @(Get-ReservedMissionNames -Student $Student -ExceptLineId $lineId)
                                                $nextEvents = @(Get-OptedEvents -SyllabusData $Student.Syllabus -AssumeComplete @($mName) -ExcludeMissions $reserved |
                                                                Where-Object { ([string]$_.Type).Trim() -in @('Single', 'Formation', 'Both') } |
                                                                Sort-Object -Property { Get-SyllabusSortKey $_ })
                                                if ($nextEvents.Count -eq 0) {
                                                                $ev.Add((New-RuleViolation 'OB_NO_SECOND_LEG' "No flyable return-leg event after $mName" -Mission $mName))
                                                }
                                                elseif (ConvertTo-FaisBool $nextEvents[0].IsCheck) {
                                                                $ev.Add((New-RuleViolation 'OB_RETURN_CHECK' "Return leg would be checkride $($nextEvents[0].Mission)" -Mission $mName))
                                                }
                                }

                                # Checkride evaluator independence, and only Check IPs give checkrides
                                if ($mIsCheck) {
                                                foreach ($past in $completedReal) {
                                                                if ($past.IsSim -or [string]::IsNullOrWhiteSpace($past.Instructor)) { continue }
                                                                if ((Get-BaseName $past.Instructor) -ieq $ipBase -and $ipBase -ne '') {
                                                                                $ev.Add((New-RuleViolation 'CHECK_SAME_IP' ("Evaluator {0} already flew with this student ({1} {2})" -f $ipBase, $past.Mission, $past.Start.ToString('dd MMM')) -Mission $mName))
                                                                                break
                                                                }
                                                }
                                                if (-not (ConvertTo-FaisBool $ip.IsCheck)) {
                                                                $ev.Add((New-RuleViolation 'CHECK_NEEDS_CHECK_IP' "Checkride needs a Check IP ($ipBase isn't)" -Mission $mName))
                                                }
                                }

                                # Formation lead: needs an open line for a wingman / DS. If no other
                                # student can fly Form here but one can fly single-ship, prefer single-ship.
                                if ($mType -like '*Form*' -and -not $IsWingCandidate -and $null -eq $pinnedWing -and ($Explain -or -not (Test-AnyBlocking $ev))) {
                                                $ctx = $null
                                                if ($null -ne $LineContext -and $LineContext.ContainsKey('Formation')) { $ctx = $LineContext['Formation'] }
                                                else {
                                                                $ctx = Get-FormationLineContext -Sortie $Sortie -Schedule $Schedule -Students $Students -FormLookAhead $FormLookAhead
                                                                if ($null -ne $LineContext) { $LineContext['Formation'] = $ctx }
                                                }
                                                if (@($ctx.OpenWingLines).Count -lt 1) {
                                                                $ev.Add((New-RuleViolation 'FORM_NO_WING_LINE' ("No open Day Local line within {0} min for a wingman/DS" -f [int]$FormLookAhead) -Mission $mName))
                                                }
                                                elseif (-not (Test-SetHasOther $ctx.FormCapable $studentName)) {
                                                                if (Test-SetHasOther $ctx.SingleCapable $studentName) {
                                                                                $ev.Add((New-RuleViolation 'FORM_PREFER_SINGLE' 'No formation wingman available; another student can fly single-ship here' -Severity Advisory -Mission $mName))
                                                                }
                                                                else {
                                                                                [void]$en.Add('Direct Support Required')
                                                                }
                                                }
                                }

                                # Solo flight continuity (notes only)
                                if (-not $isSimLine -and $mType -like '*Solo*') {
                                                [void]$en.Add('Solo Flight')
                                                if (-not $ipIsSim -and $ipBase -ne '' -and $ipBase -ne 'NEED IP') {
                                                                $last5 = @($completedReal | Where-Object { -not $_.IsSim -and $_.Start -lt $start -and -not [string]::IsNullOrWhiteSpace($_.Instructor) } |
                                                                                Sort-Object -Property Start -Descending | Select-Object -First 5)
                                                                $found = $false
                                                                foreach ($f in $last5) { if ((Get-BaseName $f.Instructor) -ieq $ipBase) { $found = $true; break } }
                                                                if (-not $found) { [void]$en.Add('Need IP in Continuity to Brief') }
                                                }
                                }

                                $eventResults.Add([PSCustomObject]@{
                                                Event      = $evt
                                                Mission    = $mName
                                                Passed     = (-not (Test-AnyBlocking $ev))
                                                Violations = $ev.ToArray()
                                                Notes      = $en.ToArray()
                                })
                }

                return (New-HardRuleResult $violations $notes $eventResults)
}

# Legacy entry point: $true / $false like the old HardRules, with its side effects
# (Student.OptedEvents, Sortie.Mission = first passing event by Syllabus ID, Student.PendingNote)
function HardRules {
                param (
                                $Student,
                                $Sortie,
                                $Schedule,
                                $FormLookAhead = 22,
                                $Students,
                                [switch]$IsWingCandidate
                )

                $preset = $null
                if ($null -ne $Sortie.Mission -and $Sortie.Mission -isnot [string] -and (Get-NameText $Sortie.Mission) -ne '') { $preset = $Sortie.Mission }
                $r = Test-HardRules -Student $Student -Sortie $Sortie -Schedule @($Schedule) -Students @($Students) -FormLookAhead $FormLookAhead `
                                -Mission $preset -IsWingCandidate:$IsWingCandidate
                if (-not $r.Passed) {
                                $Student | Add-Member -MemberType NoteProperty -Name 'PendingNote' -Value $null -Force
                                return $false
                }

                $passed = @($r.PassedEvents | Sort-Object -Property { Get-SyllabusSortKey $_ })
                $Student | Add-Member -MemberType NoteProperty -Name 'OptedEvents' -Value $passed -Force
                $Sortie.Mission = $passed[0]
                $notes = @(Get-ChosenEventNotes -Hard $r -Evt $passed[0] -Soft $null)
                $joined = if ($notes.Count -gt 0) { $notes -join ' | ' } else { $null }
                $Student | Add-Member -MemberType NoteProperty -Name 'PendingNote' -Value $joined -Force
                return $true
}

function New-HardRuleResult {
                param ($Violations, $Notes, $EventResults)

                $events = @($EventResults.ToArray())
                $passedEvents = @($events | Where-Object { $_.Passed } | ForEach-Object { $_.Event })
                $studentBlocked = Test-AnyBlocking $Violations.ToArray()
                $uniqueNotes = @($Notes.ToArray() | Select-Object -Unique)
                return [PSCustomObject]@{
                                Passed       = ((-not $studentBlocked) -and $passedEvents.Count -gt 0)
                                Violations   = $Violations.ToArray()
                                Notes        = $uniqueNotes
                                PassedEvents = $passedEvents
                                EventResults = $events
                }
}

# Why the student has nothing to fly on this sortie type
function Get-NoEligibleMessage {
                param ($Student, $Sortie)

                $sortieType = [string]$Sortie.SortieType
                $label = if ([string]::IsNullOrWhiteSpace($sortieType)) { 'this line' } else { $sortieType }
                $lineId = [string](Get-OptionalProperty $Sortie 'BoardLineId' '')
                $reserved = @(Get-ReservedMissionNames -Student $Student -ExceptLineId $lineId)
                $opted = @(Get-OptedEvents -SyllabusData $Student.Syllabus -ExcludeMissions $reserved | Sort-Object -Property { Get-SyllabusSortKey $_ })
                if ($opted.Count -eq 0) { return "No eligible $label events (syllabus complete or waiting on reserved events)" }

                $types = Get-AllowedEventTypes $sortieType
                $fits = @($opted | Where-Object { $null -ne $types -and ([string]$_.Type).Trim() -in $types })
                if ($fits.Count -gt 0 -and (Test-IsOBSortieType $sortieType)) {
                                return "No eligible $label events (no second O&B leg available after $($fits[0].Mission))"
                }
                $nextList = (@($opted | Select-Object -First 3 | ForEach-Object { "$($_.Mission) $($_.Type)" })) -join ', '
                return "No eligible $label events (next: $nextList)"
}

# Mission name (upper case) -> complete, for a student's whole syllabus
function Get-MissionDoneMap {
                param ($Student)

                $done = @{}
                foreach ($m in @($Student.Syllabus)) { if ($null -ne $m) { $done[([string]$m.Mission).Trim().ToUpperInvariant()] = (Test-MissionComplete $m) } }
                return $done
}

# Why a student isn't opted for an event ('' when they are): it is already complete, or a
# prereq isn't complete yet (e.g. a sim whose prereq is the dollar ride).
# -DoneMap: a Get-MissionDoneMap result, when checking many events of one student.
function Get-NotOptedReason {
                param ($Student, $Mission, [hashtable]$DoneMap)

                if (Test-MissionComplete $Mission) { return 'already complete' }
                $prereqText = [string]$Mission.Prereqs
                if ($prereqText.Trim() -eq '0' -or [string]::IsNullOrWhiteSpace($prereqText)) { return '' }
                $done = if ($null -ne $DoneMap) { $DoneMap } else { Get-MissionDoneMap $Student }
                $missing = @($prereqText -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' -and -not $done[$_.ToUpperInvariant()] })
                if ($missing.Count -eq 0) { return '' }
                $verb = if ($missing.Count -eq 1) { "isn't" } else { "aren't" }
                return "$($missing -join ', ') $verb complete yet"
}

# Why one specific event can't go on this line
function Get-IneligibleReason {
                param ($Student, $Mission, $Sortie)

                if (Test-MissionComplete $Mission) { return 'already complete' }

                $done = @{}
                foreach ($m in @($Student.Syllabus)) { $done[([string]$m.Mission).Trim().ToUpperInvariant()] = (Test-MissionComplete $m) }
                $prereqText = [string]$Mission.Prereqs
                if ($prereqText.Trim() -ne '0' -and -not [string]::IsNullOrWhiteSpace($prereqText)) {
                                $missing = @($prereqText -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' -and -not $done[$_.ToUpperInvariant()] })
                                if ($missing.Count -gt 0) { return "prereq $($missing -join ', ') not complete" }
                }

                $lineId = [string](Get-OptionalProperty $Sortie 'BoardLineId' '')
                $reserved = @(Get-ReservedMissionNames -Student $Student -ExceptLineId $lineId)
                if ($reserved -contains [string]$Mission.Mission) { return 'already pinned on another line' }

                $types = Get-AllowedEventTypes ([string]$Sortie.SortieType)
                if ($null -eq $types -or ([string]$Mission.Type).Trim() -notin $types) {
                                return "Type $($Mission.Type) not allowed on $($Sortie.SortieType)"
                }
                if (Test-IsOBSortieType ([string]$Sortie.SortieType)) { return 'no second O&B leg available' }
                return 'not available'
}


# =================================================================================
#  ____         __ _     ____        _
# / ___|  ___  / _| |_  |  _ \ _   _| | ___  ___
# \___ \ / _ \| |_| __| | |_) | | | | |/ _ \/ __|
#  ___) | (_) |  _| |_  |  _ <| |_| | |  __/\__ \
# |____/ \___/|_|  \__| |_| \_\\__,_|_|\___||___/
#
# Score for one student flying $Sortie.Mission on $Sortie. Higher is better.
# -Detailed returns {Score, Breakdown (label -> points), Reasons (label -> text for
# the board), Notes, Student, Mission, ETD}.
# =================================================================================
function SoftRules {
                param (
                                $Student,
                                $Config,
                                $Sortie,
                                [array]$Instructors = @(),
                                [switch]$Detailed
                )

                $Score = 0.0
                $Breakdown = [ordered]@{}
                $Reasons = [ordered]@{}
                $NewNotes = New-Object System.Collections.Generic.List[string]
                $SortieDate = Parse-ETD $Sortie.ETD
                $missionName = Get-NameText $Sortie.Mission
                $missionType = [string](Get-OptionalProperty $Sortie.Mission 'Type' '')
                $missionIsCheck = ConvertTo-FaisBool (Get-OptionalProperty $Sortie.Mission 'IsCheck' $false)
                $sortieType = [string]$Sortie.SortieType
                $ip = Resolve-SortieInstructor -Instructors $Instructors -Instructor $Sortie.Instructor
                $CurrentIP = Get-NameText $ip
                $CurrentIP_Base = ''
                $CurrentIP_Short = ''

                ####################################
                # RULE : DOLLAR RIDE BONUS
                # Gets dollar rides out the door
                ####################################
                if ($missionName -like '*FT5001*') {
                                $pts = Get-Weight $Config 'DollarRideBonus'
                                $Score += $pts
                                $Breakdown['Dollar Ride Bonus'] = $pts
                                $Reasons['Dollar Ride Bonus'] = "Dollar Ride $missionName"
                }

                ###############################################################
                # RULE : Sim Stop Prevention
                # Every flight that gets scheduled creates a 'debt' that is paid off by completing sims
                ###############################################################
                $simDebt = 0.0
                if ($null -ne $Student.SimDebt -and [string]$Student.SimDebt -ne '') { $simDebt = [double]$Student.SimDebt }
                # A negative SimDebt means the student is ahead on sims. One decimal, like the points
                # the board shows next to it.
                $debtShown = [math]::Round([math]::Abs($simDebt), 1)
                $debtText = if ($simDebt -lt 0) { "Sim credit $(Format-Number $debtShown)" } else { "Sim Debt $(Format-Number $debtShown)" }
                if ($missionName -like '*FT5*' -or $missionName -like '*FF5*' -or $missionName -like '*AA5*') {
                                $pts = -$simDebt
                                $Score += $pts
                                $Breakdown['Sim Debt (flight)'] = $pts
                                $Reasons['Sim Debt (flight)'] = if ($simDebt -lt 0) { "$debtText, flight uses it" } else { "$debtText, flight adds to it" }
                }
                elseif ($missionName -like '*FI3*' -or $missionName -like '*FE3*' -or $missionName -like '*FI4*' -or $missionName -like '*FE4*' -or
                                $missionName -like '*FN3*' -or $missionName -like '*FN4*' -or $missionName -like '*AA4*') {
                                $pts = $simDebt
                                $Score += $pts
                                $Breakdown['Sim Debt (sim)'] = $pts
                                $Reasons['Sim Debt (sim)'] = if ($simDebt -lt 0) { "$debtText, sim adds to it" } else { "$debtText, sim pays it down" }
                }

                ####################################
                # RULE: CAP BONUS
                # CAPsters don't usually double turn, so give them a bonus to get on the schedule
                ####################################
                $isCap = ConvertTo-FaisBool $Student.IsCap
                if ($isCap) {
                                $pts = Get-Weight $Config 'CAPBonus'
                                $Score += $pts
                                $Breakdown['CAP Bonus'] = $pts
                                $Reasons['CAP Bonus'] = 'CAP student'
                }

                #####################################
                # RULE: STRINGS AND COUNTERS
                # Rewards three strings, penalizes everything else to varying degrees
                # Flights only: a sim line never scores here, and past sims never count
                #####################################
                $PastFlights = @(
                                $Student.Syllabus | Where-Object {
                                                (Test-MissionComplete $_) -and
                                                $null -ne (Parse-ETD $_.ETD) -and
                                                (Get-MissionKind -Mission $_ -InstructorName (Get-NameText $_.Instructor)) -eq 'Flight' -and
                                                -not [string]::IsNullOrWhiteSpace((Get-NameText $_.Instructor)) -and
                                                (Parse-ETD $_.ETD) -lt $SortieDate
                                } | Sort-Object -Property { Parse-ETD $_.ETD } -Descending
                )
                $IP1_Base = if ($PastFlights.Count -ge 1) { Get-BaseName (Get-NameText $PastFlights[0].Instructor) } else { '' }
                $IP2_Base = if ($PastFlights.Count -ge 2) { Get-BaseName (Get-NameText $PastFlights[1].Instructor) } else { '' }
                $IP3_Base = if ($PastFlights.Count -ge 3) { Get-BaseName (Get-NameText $PastFlights[2].Instructor) } else { '' }
                $IP4_Base = if ($PastFlights.Count -ge 4) { Get-BaseName (Get-NameText $PastFlights[3].Instructor) } else { '' }

                $IsEffectiveCounter = ConvertTo-FaisBool $ip.IsCounter -Default $true
                $hasRealIP = ($CurrentIP -ne '' -and $CurrentIP -notmatch 'SIM IP' -and $CurrentIP -ne 'NEED IP')
                $isFlightSortie = -not (Test-IsSimSortieType $sortieType) -and $missionType -notmatch '(?i)^\s*(IFT|OFT)\s*$'

                if ($hasRealIP -and $isFlightSortie) {
                                $CurrentIP_Base = Get-BaseName $CurrentIP
                                $CurrentIP_Short = Get-ShortIPName $CurrentIP
                                $nonCounterTag = ''

                                # A non-counter who flew the student's last ride (or the one before) counts as a counter
                                if (-not $IsEffectiveCounter -and ($CurrentIP_Base -eq $IP1_Base -or ($CurrentIP_Base -eq $IP2_Base -and $CurrentIP_Base -ne ''))) {
                                                $IsEffectiveCounter = $true
                                                $nonCounterTag = ' (non-counter)'
                                }

                                if (-not $IsEffectiveCounter) {
                                                $pts = Get-Weight $Config 'NonCounter'
                                                $Score += $pts
                                                $Breakdown['Non-Counter'] = $pts
                                                $Reasons['Non-Counter'] = "$CurrentIP_Short is a non-counter"
                                }
                                else {
                                                if ($IP1_Base -eq $CurrentIP_Base) {
                                                                # This ride plus the rides in a row this IP already flew
                                                                $StringText = "$(1 + (Get-IPRunLength -Flights $PastFlights -IPBase $CurrentIP_Base)) String$nonCounterTag"
                                                                if ($IP1_Base -eq $IP2_Base -and $IP2_Base -eq $IP3_Base -and $IP3_Base -eq $IP4_Base) {
                                                                                $pts = Get-Weight $Config 'String5' -100
                                                                                $Score += $pts
                                                                                $Breakdown['String 5'] = $pts
                                                                                $Reasons['String 5'] = $StringText
                                                                }
                                                                elseif ($IP1_Base -eq $IP2_Base -and $IP2_Base -eq $IP3_Base) {
                                                                                $pts = Get-Weight $Config 'String4' -10
                                                                                $Score += $pts
                                                                                $Breakdown['String 4'] = $pts
                                                                                $Reasons['String 4'] = $StringText
                                                                }
                                                                else {
                                                                                $pts = Get-Weight $Config 'ContString'
                                                                                $Score += $pts
                                                                                $Breakdown['Continue String'] = $pts
                                                                                $Reasons['Continue String'] = $StringText
                                                                                if ($isCap) {
                                                                                                $pts2 = Get-Weight $Config 'CAPBonus'
                                                                                                $Score += $pts2
                                                                                                $Breakdown['CAP Bonus (string)'] = $pts2
                                                                                                $Reasons['CAP Bonus (string)'] = "CAP student, $StringText"
                                                                                }
                                                                }
                                                }
                                                elseif ($IP2_Base -eq $CurrentIP_Base -and $IP1_Base -ne $CurrentIP_Base -and $IP2_Base -ne '') {
                                                                $pts = Get-Weight $Config 'Zipper'
                                                                $Score += $pts
                                                                $Breakdown['Zipper'] = $pts
                                                                $arrow = " $([char]0x2192) "
                                                                $Reasons['Zipper'] = "Zipper: $CurrentIP_Short$arrow$(Get-ShortIPName (Get-NameText $PastFlights[0].Instructor))$arrow$CurrentIP_Short$nonCounterTag"
                                                }
                                                elseif ($IP1_Base -ne '' -and $IP1_Base -ne $CurrentIP_Base) {
                                                                $pts = Get-Weight $Config 'BreakString'
                                                                $Score += $pts
                                                                $Breakdown['Break String'] = $pts
                                                                $brokenRun = Get-IPRunLength -Flights $PastFlights -IPBase $IP1_Base
                                                                $lastIP = Get-ShortIPName (Get-NameText $PastFlights[0].Instructor)
                                                                $Reasons['Break String'] = if ($brokenRun -ge 2) { "Breaks $brokenRun String w/ $lastIP" } else { "New IP after 1 ride w/ $lastIP" }
                                                }
                                }
                }

                ##################################################
                # RULE : SIMULATOR TYPE MATCHING
                # Prioritizes putting OFT events in OFT slots and IFT events in IFT slots.
                ##################################################
                if ($sortieType -match 'OFT' -and $missionType -match 'IFT') {
                                $pts = Get-Weight $Config 'OFTIFTPen'
                                $Score += $pts
                                $Breakdown['OFT/IFT Mismatch'] = $pts
                                $Reasons['OFT/IFT Mismatch'] = "IFT event $missionName in an OFT"
                                [void]$NewNotes.Add("$missionName (IFT) in an OFT")
                }

                ############################################
                # RULE: EXHAUSTION POINTS
                # Starts subtracting points the busier and busier the SP is
                ############################################
                $ExhaustionDays = Get-Weight $Config 'ExhaustionDays' 5
                $ExhaustionMult = Get-Weight $Config 'ExhaustionMult' 1
                $RecentFlights = 0
                $RecentSims = 0
                foreach ($m in @($Student.Syllabus)) {
                                if (-not (Test-MissionComplete $m)) { continue }
                                $EventDate = Parse-ETD $m.ETD
                                if ($null -eq $EventDate -or $null -eq $SortieDate) { continue }
                                $DaysDiff = ($SortieDate.Date - $EventDate.Date).TotalDays
                                if ($DaysDiff -ge 0 -and $DaysDiff -le $ExhaustionDays -and $EventDate -lt $SortieDate) {
                                                if ((Get-MissionKind -Mission $m -InstructorName (Get-NameText $m.Instructor)) -eq 'Sim') { $RecentSims++ } else { $RecentFlights++ }
                                }
                }
                $pts = (($RecentFlights * 2) + $RecentSims) * $ExhaustionMult
                $Score += $pts
                $Breakdown['Exhaustion'] = $pts
                $recentParts = @(
                                if ($RecentFlights -gt 0) { Format-Count $RecentFlights 'Flight' }
                                if ($RecentSims -gt 0) { Format-Count $RecentSims 'Sim' }
                )
                $recentText = if ($recentParts.Count -gt 0) { $recentParts -join ' and ' } else { 'No Flights or Sims' }
                $dayWord = if ($ExhaustionDays -eq 1) { 'Day' } else { 'Days' }
                $Reasons['Exhaustion'] = "$recentText in last $(Format-Number $ExhaustionDays) $dayWord"

                ##########################################################
                # RULE: BIT TRACKER
                # Only run the BIT Tracker if the event being scheduled is a flight (Not IFT/OFT)
                ##########################################################
                if (-not (Test-IsSimSortieType $sortieType) -and $PastFlights.Count -gt 0 -and $null -ne $SortieDate) {
                                $BIT_X = Get-Weight $Config 'BIT_X' 1.5
                                $BIT_Y = Get-Weight $Config 'BIT_Y' 6
                                $LastFlightDate = Parse-ETD $PastFlights[0].ETD
                                $EffectiveBreakDays = ($SortieDate.Date - $LastFlightDate.Date).TotalDays - 1
                                $capMult = if ($isCap) { 2 } else { 1 }
                                $daysOff = "$(Format-Count ([int]$EffectiveBreakDays) 'day') without a flight"

                                if ($EffectiveBreakDays -gt 0) {
                                                if ($EffectiveBreakDays -le $BIT_Y) {
                                                                $pts = $EffectiveBreakDays * $BIT_X * $capMult
                                                                $Score += $pts
                                                                $Breakdown['BIT Tracker'] = $pts
                                                                $lastFlew = $LastFlightDate.ToString('ddd M/d', [System.Globalization.CultureInfo]::InvariantCulture)
                                                                $Reasons['BIT Tracker'] = "$daysOff (last flew $lastFlew)$(if ($isCap) { ', CAP x2' })"
                                                }
                                                else {
                                                                $pts = $BIT_X * ($BIT_Y / 2.0)
                                                                $Score += $pts
                                                                $Breakdown['BIT Tracker (capped)'] = $pts
                                                                $Reasons['BIT Tracker (capped)'] = "$daysOff, past the $(Format-Number $BIT_Y)-day BIT limit"
                                                                [void]$NewNotes.Add("BIT! $daysOff, consider an 86")
                                                }
                                }
                }

                ##################################################
                # RULE : Counter Limit
                # Tries to keep the total number of counters low
                ##################################################
                if ($IsEffectiveCounter -and $CurrentIP_Base -ne '') {
                                $AllPastIPs = @($PastFlights | ForEach-Object { Get-BaseName (Get-NameText $_.Instructor) } | Select-Object -Unique)
                                if ($AllPastIPs -contains $CurrentIP_Base) {
                                                $pts = Get-Weight $Config 'CounterContinuity' 5
                                                $Score += $pts
                                                $Breakdown['Counter Continuity'] = $pts
                                                $ridesWithIP = @($PastFlights | Where-Object { (Get-BaseName (Get-NameText $_.Instructor)) -eq $CurrentIP_Base }).Count
                                                $Reasons['Counter Continuity'] = "$CurrentIP_Short already flew $ridesWithIP of $(Format-Count $PastFlights.Count 'ride')"
                                }
                }

                ################################################
                # RULE : PRIORITIZE CHECKRIDES
                # If we have a check IP, we should use them for checkrides
                ################################################
                if ((ConvertTo-FaisBool $ip.IsCheck) -and $missionIsCheck) {
                                $pts = 50
                                $Score += $pts
                                $Breakdown['Checkride Priority'] = $pts
                                $Reasons['Checkride Priority'] = "Checkride $missionName with check IP $(Get-ShortIPName $CurrentIP)"
                                [void]$NewNotes.Add('Checkride! Requires Formal')
                }

                ##########################################################
                # RULE: STUDENT PRIORITY
                # This just spots them whatever their priority is
                ##########################################################
                $pts = 0.0
                if ($null -ne $Student.Priority -and [string]$Student.Priority -ne '') { [void][double]::TryParse([string]$Student.Priority, [ref]$pts) }
                $Score += $pts
                $Breakdown['Student Priority'] = $pts
                $Reasons['Student Priority'] = 'Student Priority'

                if ($Detailed) {
                                return [PSCustomObject]@{
                                                Score     = $Score
                                                Breakdown = $Breakdown
                                                Reasons   = $Reasons
                                                Notes     = $NewNotes.ToArray()
                                                Student   = [string]$Student.Name
                                                Mission   = $missionName
                                                ETD       = $Sortie.ETD
                                }
                }
                return $Score
}

# Runs SoftRules on a copy of the line with the event set, so the real line is never
# left half-assigned
function Get-EventScore {
                param ($Student, $Sortie, $Evt, $SoftWeights, [array]$Instructors)

                $probe = $Sortie.PSObject.Copy()
                $probe.Mission = $Evt
                return (SoftRules -Student $Student -Config $SoftWeights -Sortie $probe -Instructors $Instructors -Detailed)
}

# Higher score wins; ties go to higher Priority, then name (so runs are repeatable)
function Test-CandidateBetter {
                param ($A, $B)

                if ($null -eq $B) { return $true }
                if ($A.Score -ne $B.Score) { return ($A.Score -gt $B.Score) }
                $pa = 0.0; $pb = 0.0
                [void][double]::TryParse([string]$A.Student.Priority, [ref]$pa)
                [void][double]::TryParse([string]$B.Student.Priority, [ref]$pb)
                if ($pa -ne $pb) { return ($pa -gt $pb) }
                $c = [string]::Compare([string]$A.Student.Name, [string]$B.Student.Name, $true)
                if ($c -ne 0) { return ($c -lt 0) }
                return ((Get-SyllabusSortKey $A.Event) -lt (Get-SyllabusSortKey $B.Event))
}


# =================================================================================
#  ____                 _   _               _               _        ____             _   _
# |  _ \ __ _ ___ ___  | | | | __ _ _ __ __| | ___ ___   __| | ___  / ___|  ___  _ __| |_(_) ___
# | |_) / _` / __/ __| | |_| |/ _` | '__/ _` |/ __/ _ \ / _` |/ _ \ \___ \ / _ \| '__| __| |/ _ \
# |  __/ (_| \__ \__ \ |  _  | (_| | | | (_| | (_| (_) | (_| |  __/  ___) | (_) | |  | |_| |  __/
# |_|   \__,_|___/___/ |_| |_|\__,_|_|  \__,_|\___\___/ \__,_|\___| |____/ \___/|_|   \__|_|\___|
#
# Writes a pinned student onto their line. Hard rules are NOT enforced here (the
# caller flags them); this only records the event, callsigns, formation and O&B leg.
# Returns the updated FormIndex.
# =================================================================================
function Process-HardcodedSortie {
                param (
                                [Parameter(Mandatory)] $Sortie,
                                [Parameter(Mandatory)] $Student,
                                $Mission,
                                [array]$Schedule = @(),
                                [array]$Students = @(),
                                [array]$Instructors = @(),
                                $SoftWeights,
                                [array]$Formations = @(),
                                [int]$FormIndex = 0
                )

                $ip = Resolve-SortieInstructor -Instructors $Instructors -Instructor $Sortie.Instructor
                $Sortie.Instructor = $ip
                $Sortie.Student = $Student
                $Sortie.Mission = $Mission
                [void](Initialize-SortieRow $Sortie)
                $Sortie.FilledBy = 'Pin'

                if ($null -ne $Mission) {
                                [void](Complete-StudentMission -Student $Student -Mission $Mission -InstructorName (Get-NameText $ip) -ETD $Sortie.ETD -SoftWeights $SoftWeights -Sortie $Sortie)
                                Write-Host " [PINNED] '$($Student.Name)' on '$($Mission.Mission)' (ETD: $($Sortie.ETD))" -ForegroundColor Cyan
                }
                else {
                                Write-Warning " [PINNED] '$($Student.Name)' has no mission on ETD $($Sortie.ETD)."
                }

                if ([string]::IsNullOrWhiteSpace([string]$Sortie.Callsign)) {
                                $Sortie.Callsign = $ip.Callsign
                                $Sortie.Num = $ip.Num
                }

                # Formation pairing only for true leads (not a line pinned as someone's wing)
                $isWing = (Test-IsPairedWing -Sortie $Sortie -Schedule $Schedule) -or ([string]$Sortie.Notes -like '*Wingman to*')
                if ($null -ne $Mission -and ([string]$Mission.Type) -like '*Form*' -and -not $isWing) {
                                $existingWing = Get-PinnedWingRow -Sortie $Sortie -Schedule $Schedule
                                $FormIndex = Pair-FormationWingman -PrimarySortie $Sortie -PrimaryStudent $Student -PrimaryMission $Mission `
                                                -Schedule $Schedule -Students $Students -Instructors $Instructors -SoftWeights $SoftWeights `
                                                -FormIndex $FormIndex -Formations $Formations -ExistingWingSortie $existingWing
                }

                # O&B second leg
                if ($null -ne $Mission -and (Test-IsOBSortieType ([string]$Sortie.SortieType))) {
                                Schedule-ReturnLeg -Sortie $Sortie -Student $Student -Schedule $Schedule -Instructors $Instructors -SoftWeights $SoftWeights
                }

                return $FormIndex
}


# =================================================================================
#  FORMATION PAIRING
#  First tries to find another student to fly as wingman. If none is available, the
#  next open live-fly line inside the look-ahead window becomes Direct Support.
#  FormCallsigns is consecutive Lead/Wing pairs (e.g. HAZE 11, HAZE 12); each pair
#  attempt consumes two entries and advances FormIndex by 2.
# =================================================================================
function Pair-FormationWingman {
                param (
                                $PrimarySortie,
                                $PrimaryStudent,
                                $PrimaryMission,
                                [array]$Schedule,
                                [array]$Students,
                                [array]$Instructors,
                                $SoftWeights,
                                [int]$FormIndex = 0,
                                [array]$Formations = @(),
                                $ExistingWingSortie
                )

                try {
                                $LeadEntry = $null
                                $WingEntry = $null
                                $haveCallsigns = ($null -ne $Formations -and $Formations.Count -gt 0)
                                if ($haveCallsigns) {
                                                if ($FormIndex -ge $Formations.Count) { $FormIndex = 0 }
                                                $LeadEntry = $Formations[$FormIndex]
                                                $WingIndex = $FormIndex + 1
                                                if ($WingIndex -ge $Formations.Count) { $WingIndex = 0 }
                                                $WingEntry = $Formations[$WingIndex]
                                                $PrimarySortie.Callsign = $LeadEntry.Callsign
                                                $PrimarySortie.Num = $LeadEntry.Num
                                }
                                $leadName = Get-NameText $PrimaryStudent

                                # The board already paired a pinned wingman to this lead
                                if ($null -ne $ExistingWingSortie) {
                                                if ($haveCallsigns) {
                                                                $ExistingWingSortie.Callsign = $WingEntry.Callsign
                                                                $ExistingWingSortie.Num = $WingEntry.Num
                                                }
                                                Add-SortieNote -Sortie $ExistingWingSortie -Note "Wingman to $leadName"
                                                Add-SortieNote -Sortie $PrimarySortie -Note "Wingman: $(Get-NameText $ExistingWingSortie.Student)"
                                }
                                else {
                                                $FormLookAhead = Get-Weight $SoftWeights 'FormLookAhead' 22
                                                $LevelDifference = Get-Weight $SoftWeights 'LevelDifference' 0
                                                $CurrentETD = Parse-ETD $PrimarySortie.ETD
                                                $TimeWindowEnd = $CurrentETD.AddMinutes($FormLookAhead)
                                                $primaryLvl = 0
                                                [void][int]::TryParse([string]$PrimaryMission.Level, [ref]$primaryLvl)

                                                $openLines = @(foreach ($r in $Schedule) {
                                                                if (Test-SameSortie $r $PrimarySortie) { continue }
                                                                if (-not (Test-IsOpenWingLine $r)) { continue }
                                                                $rs = Parse-ETD $r.ETD
                                                                if ($null -ne $rs -and $rs -ge $CurrentETD -and $rs -le $TimeWindowEnd) { $r }
                                                })

                                                $best = $null
                                                foreach ($WingSortie in $openLines) {
                                                                $WingSortie.Instructor = Resolve-SortieInstructor -Instructors $Instructors -Instructor $WingSortie.Instructor
                                                                foreach ($FormStudent in $Students) {
                                                                                if ([string]::IsNullOrWhiteSpace([string]$FormStudent.Name) -or [string]$FormStudent.Name -eq $leadName) { continue }

                                                                                $TentativeWingMission = $null
                                                                                foreach ($evt in @(Get-ValidOptedEvents -Student $FormStudent -Sortie $WingSortie)) {
                                                                                                if ([string]$evt.Type -like '*Form*') { $TentativeWingMission = $evt; break }
                                                                                }
                                                                                if ($null -eq $TentativeWingMission) { continue }

                                                                                $hr = Test-HardRules -Student $FormStudent -Sortie $WingSortie -Schedule $Schedule -Students $Students -Instructors $Instructors `
                                                                                                -FormLookAhead $FormLookAhead -Mission $TentativeWingMission -IsWingCandidate
                                                                                if (-not $hr.Passed) { continue }

                                                                                $soft = Get-EventScore -Student $FormStudent -Sortie $WingSortie -Evt $TentativeWingMission -SoftWeights $SoftWeights -Instructors $Instructors
                                                                                $wingLvl = 0
                                                                                [void][int]::TryParse([string]$TentativeWingMission.Level, [ref]$wingLvl)
                                                                                $cand = [PSCustomObject]@{
                                                                                                Student = $FormStudent
                                                                                                Sortie  = $WingSortie
                                                                                                Event   = $TentativeWingMission
                                                                                                Score   = $soft.Score + (($wingLvl - $primaryLvl) * $LevelDifference)
                                                                                                Hard    = $hr
                                                                                                Soft    = $soft
                                                                                }
                                                                                if (Test-CandidateBetter $cand $best) { $best = $cand }
                                                                }
                                                }

                                                if ($null -ne $best) {
                                                                $w = $best.Sortie
                                                                $w.Student = $best.Student
                                                                $w.Mission = $best.Event
                                                                [void](Initialize-SortieRow $w)
                                                                $w.FilledBy = 'Wing'
                                                                if ($haveCallsigns) {
                                                                                $w.Callsign = $WingEntry.Callsign
                                                                                $w.Num = $WingEntry.Num
                                                                }
                                                                [void](Complete-StudentMission -Student $best.Student -Mission $best.Event -InstructorName (Get-NameText $w.Instructor) -ETD $w.ETD -SoftWeights $SoftWeights -Sortie $w)
                                                                Add-SortieNote -Sortie $w -Note "Wingman to $leadName"
                                                                foreach ($n in (Get-ChosenEventNotes -Hard $best.Hard -Evt $best.Event -Soft $best.Soft)) { Add-SortieNote -Sortie $w -Note $n }
                                                                Add-SortieNote -Sortie $PrimarySortie -Note "Wingman: $($best.Student.Name)"
                                                }
                                                else {
                                                                # No student wingman: claim the soonest open live-fly line as Direct Support
                                                                $ds = $openLines | Sort-Object -Property { Parse-ETD $_.ETD } | Select-Object -First 1
                                                                if ($null -ne $ds) {
                                                                                $ds.Student = 'Need Form SP'
                                                                                [void](Initialize-SortieRow $ds)
                                                                                $ds.FilledBy = 'DS'
                                                                                if ($haveCallsigns) {
                                                                                                $ds.Callsign = $WingEntry.Callsign
                                                                                                $ds.Num = $WingEntry.Num
                                                                                }
                                                                                Add-SortieNote -Sortie $ds -Note "Direct Support for $leadName"
                                                                                Add-SortieNote -Sortie $PrimarySortie -Note 'Wingman: DS'
                                                                }
                                                                else {
                                                                                Add-SortieNote -Sortie $PrimarySortie -Note 'Wingman: none found'
                                                                }
                                                }
                                }

                                # Always consume the full Lead/Wing pair
                                if ($haveCallsigns) {
                                                $FormIndex += 2
                                                if ($FormIndex -ge $Formations.Count) { $FormIndex = 0 }
                                }
                                return $FormIndex
                } catch {
                                Write-Warning "Error resolving formation pairing: $($_.Exception.Message)"
                                return $FormIndex
                }
}

# Hard-rule notes for the chosen event, then its soft-rule notes
function Get-ChosenEventNotes {
                param ($Hard, $Evt, $Soft)

                $out = New-Object System.Collections.Generic.List[string]
                if ($null -ne $Hard) {
                                foreach ($n in @($Hard.Notes)) { if ($n) { $out.Add($n) } }
                                $er = Find-EventResult -Hard $Hard -Evt $Evt
                                if ($null -ne $er) { foreach ($n in @($er.Notes)) { if ($n) { $out.Add($n) } } }
                }
                if ($null -ne $Soft) { foreach ($n in @($Soft.Notes)) { if ($n) { $out.Add($n) } } }
                return @($out | Select-Object -Unique)
}

function Find-EventResult {
                param ($Hard, $Evt)

                $name = Get-NameText $Evt
                foreach ($er in @($Hard.EventResults)) { if ($er.Mission -ieq $name) { return $er } }
                return $null
}


# =================================================================================
#  O&B RETURN LEG
#  Called after the O&B event has been marked complete on the student.
# =================================================================================
function Schedule-ReturnLeg {
                param (
                                $Sortie,
                                $Student,
                                [array]$Schedule,
                                [array]$Instructors = @(),
                                $SoftWeights
                )

                $OBReturn = Get-LinkedReturnRow -Sortie $Sortie -Schedule $Schedule -StudentName (Get-NameText $Student)
                if ($null -eq $OBReturn) {
                                $day = (Parse-ETD $Sortie.ETD)
                                Write-Warning "Could not find an empty O&B Return row on $(if ($day) { $day.ToString('yyyy-MM-dd') }) for $(Get-NameText $Sortie.Instructor)."
                                Add-SortieNote -Sortie $Sortie -Note 'No O&B Return line found'
                                return
                }

                [void](Initialize-SortieRow $OBReturn)
                $OBReturn.Student = $Student
                $OBReturn.Instructor = $Sortie.Instructor
                $OBReturn.Callsign = $Sortie.Callsign
                $OBReturn.Num = $Sortie.Num
                $OBReturn.FilledBy = 'Return'
                Remove-StudentReservation -Student $Student -LineId @([string]$OBReturn.BoardLineId)

                # Select the mission the SP flies home: next valid event by Syllabus ID
                $ValidEvents = @(Get-ValidOptedEvents -Student $Student -Sortie $OBReturn -NoLookahead)
                if ($ValidEvents.Count -gt 0) {
                                $OBReturn.Mission = $ValidEvents[0]
                                [void](Complete-StudentMission -Student $Student -Mission $ValidEvents[0] -InstructorName (Get-NameText $OBReturn.Instructor) -ETD $OBReturn.ETD -SoftWeights $SoftWeights -Sortie $OBReturn)
                }
                else {
                                $OBReturn.Mission = $null
                                Add-SortieNote -Sortie $OBReturn -Note 'No valid mission for return leg'
                }
}


# =================================================================================
#  ____   ____ _   _ _____ ____  _   _ _     ___ _   _  ____
# / ___| / ___| | | | ____|  _ \| | | | |   |_ _| \ | |/ ___|
# \___ \| |   | |_| |  _| | | | | | | | |    | ||  \| | |  _
#  ___) | |___|  _  | |___| |_| | |_| | |___ | || |\  | |_| |
# |____/ \____|_| |_|_____|____/ \___/|_____|___|_| \_|\____|
#
# Fills every line that is not hardcoded. Pinned lines are kept (rule violations
# are flagged with OVERRIDE notes). Lines are processed in ETD order; pinned lines
# that come later in the day hold the student's time as reservations, so earlier
# fills respect them.
# Returns {Schedule, Students, Overrides, Stats, Log}.
# =================================================================================
function Invoke-FillSchedule {
                param (
                                [Parameter(Mandatory)] [AllowEmptyCollection()] [array]$Schedule,
                                [Parameter(Mandatory)] [array]$Students,
                                [array]$Instructors = @(),
                                $SoftWeights,
                                [array]$Formations = @(),
                                [scriptblock]$OnProgress,
                                [switch]$Trace
                )

                $log = New-Object System.Collections.Generic.List[string]
                $overrides = New-Object System.Collections.Generic.List[object]
                $FormLookAhead = Get-Weight $SoftWeights 'FormLookAhead' 22

                # ---------------- PREP ROWS ----------------
                $prepared = New-Object System.Collections.Generic.List[object]
                $idx = 0
                foreach ($row in $Schedule) {
                                [void](Initialize-SortieRow $row)
                                if ([string]::IsNullOrWhiteSpace([string]$row.BoardLineId)) { $row.BoardLineId = "row$idx" }
                                $row.FilledBy = $null
                                $ip = Resolve-SortieInstructor -Instructors $Instructors -Instructor $row.Instructor
                                $row.Instructor = $ip
                                if ([string]::IsNullOrWhiteSpace([string]$row.Callsign)) {
                                                $row.Callsign = $ip.Callsign
                                                $row.Num = $ip.Num
                                }
                                if (Test-IsSentinelStudent $row.Student) { $row.Student = $null }
                                if ((Get-NameText $row.Mission) -eq '') { $row.Mission = $null }
                                $key = Parse-ETD $row.ETD
                                if ($null -eq $key) { $key = [datetime]::MaxValue }
                                $prepared.Add([PSCustomObject]@{ Row = $row; Key = $key; Index = $idx })
                                $idx++
                }
                $rows = @($prepared | Sort-Object -Property Key, Index | ForEach-Object { $_.Row })

                # ---------------- PINNED LINES BECOME RESERVATIONS ----------------
                $pinStudent = @{}
                foreach ($row in $rows) {
                                $name = Get-NameText $row.Student
                                if ($name -eq '') { continue }
                                $s = Resolve-StudentByName -Students $Students -Name $name
                                if ($null -eq $s) {
                                                $log.Add("Pinned student '$name' on $($row.ETD) not found in MasterData")
                                                continue
                                }
                                $pinStudent[[string]$row.BoardLineId] = $s
                                Add-StudentReservation -Student $s -LineId ([string]$row.BoardLineId) -ETD $row.ETD -SortieType ([string]$row.SortieType) -Mission (Get-NameText $row.Mission) -Instructor (Get-NameText $row.Instructor)
                                if (Test-IsOBSortieType ([string]$row.SortieType)) {
                                                $ret = Get-LinkedReturnRow -Sortie $row -Schedule $rows -StudentName $name
                                                if ($null -ne $ret) {
                                                                Add-StudentReservation -Student $s -LineId ([string]$ret.BoardLineId) -ETD $ret.ETD -SortieType ([string]$ret.SortieType) -Instructor (Get-NameText $row.Instructor)
                                                }
                                }
                }

                # ---------------- MAIN LOOP ----------------
                $FormIndex = 0
                $CurrentDay = $null
                $n = $rows.Count
                $i = 0
                foreach ($Sortie in $rows) {
                                $i++
                                $day = Parse-ETD $Sortie.ETD
                                if ($null -ne $day -and $CurrentDay -ne $day.Date) {
                                                $CurrentDay = $day.Date
                                                $FormIndex = 0
                                }
                                if ([string]$Sortie.FilledBy -ne '') { continue }

                                if ($null -ne $OnProgress) {
                                                try { & $OnProgress $i $n ("{0} {1}" -f (Format-EtdShort $Sortie.ETD), $Sortie.SortieType) } catch { }
                                }
                                Write-Host "Scheduling Sortie $i of $n ($(Format-EtdShort $Sortie.ETD) $($Sortie.SortieType))"

                                # ---------------- HARDCODED (PINNED) ----------------
                                if ((Get-NameText $Sortie.Student) -ne '') {
                                                $FormIndex = Invoke-PinnedLine -Sortie $Sortie -Student $pinStudent[[string]$Sortie.BoardLineId] -Rows $rows -Students $Students `
                                                                -Instructors $Instructors -SoftWeights $SoftWeights -Formations $Formations -FormIndex $FormIndex -FormLookAhead $FormLookAhead -Overrides $overrides
                                                continue
                                }

                                # ---------------- RETURN LEG WITH NO O&B ----------------
                                if (Test-IsReturnSortieType ([string]$Sortie.SortieType)) {
                                                $Sortie.Student = 'For Sale'
                                                $Sortie.Mission = $null
                                                $Sortie.FilledBy = 'ForSale'
                                                Add-SortieNote -Sortie $Sortie -Note 'O&B Leg 1 Not Filled'
                                                continue
                                }

                                # ---------------- OPEN LINE ----------------
                                $FormIndex = Invoke-OpenLine -Sortie $Sortie -Rows $rows -Students $Students -Instructors $Instructors -SoftWeights $SoftWeights `
                                                -Formations $Formations -FormIndex $FormIndex -FormLookAhead $FormLookAhead -Trace:$Trace
                }

                $stats = [ordered]@{ Lines = $n; Pinned = 0; Filled = 0; Wing = 0; Return = 0; DS = 0; ForSale = 0; Overrides = $overrides.Count }
                foreach ($r in $rows) {
                                switch ([string]$r.FilledBy) {
                                                'Pin' { $stats.Pinned++ }
                                                'Engine' { $stats.Filled++ }
                                                'Wing' { $stats.Wing++ }
                                                'Return' { $stats.Return++ }
                                                'DS' { $stats.DS++ }
                                                'ForSale' { $stats.ForSale++ }
                                }
                }

                return [PSCustomObject]@{
                                Schedule  = $rows
                                Students  = $Students
                                Overrides = $overrides.ToArray()
                                Stats     = [PSCustomObject]$stats
                                Log       = $log.ToArray()
                }
}

# One pinned line: keep the student, pick the event, record it, flag rule breaks
# Which event a pinned student flies on a line: the explicit mission, otherwise the best-scoring
# event that passes the rules (Form events only for a wing, and preferred for a lead with a paired wing).
# Returns {Mission, Soft, Hard}. Mission is $null when nothing can be flown.
function Select-PinnedMission {
                param ($Student, $Sortie, [array]$Rows, [array]$Students, [array]$Instructors, $SoftWeights, [double]$FormLookAhead = 22)

                $isWing = Test-IsPairedWing -Sortie $Sortie -Schedule $Rows
                $hasWing = $null -ne (Get-PinnedWingRow -Sortie $Sortie -Schedule $Rows)
                $explicitName = Get-NameText $Sortie.Mission
                $missionArg = if ($explicitName -ne '') { $explicitName } else { $null }
                $hr = Test-HardRules -Student $Student -Sortie $Sortie -Schedule $Rows -Students $Students -Instructors $Instructors `
                                -FormLookAhead $FormLookAhead -Mission $missionArg -IsWingCandidate:$isWing -Explain

                $mission = $null
                $soft = $null
                if ($explicitName -ne '') {
                                $mission = Find-SyllabusMission -Student $Student -Mission $explicitName
                }
                else {
                                # (auto) = the best-scoring event that passes the event-level rules
                                $passed = @($hr.EventResults | Where-Object { $_.Passed })
                                $formPassed = @($passed | Where-Object { [string]$_.Event.Type -like '*Form*' })
                                $pool = if (($isWing -or $hasWing) -and $formPassed.Count -gt 0) { $formPassed } elseif ($isWing) { @() } else { $passed }
                                $best = $null
                                foreach ($er in $pool) {
                                                $sr = Get-EventScore -Student $Student -Sortie $Sortie -Evt $er.Event -SoftWeights $SoftWeights -Instructors $Instructors
                                                $cand = [PSCustomObject]@{ Student = $Student; Event = $er.Event; Score = $sr.Score; Soft = $sr }
                                                if (Test-CandidateBetter $cand $best) { $best = $cand }
                                }
                                if ($null -ne $best) {
                                                $mission = $best.Event
                                                $soft = $best.Soft
                                }
                                else {
                                                $fallback = @($hr.EventResults | ForEach-Object { $_.Event })
                                                if ($isWing) { $formOnly = @($fallback | Where-Object { [string]$_.Type -like '*Form*' }); if ($formOnly.Count -gt 0) { $fallback = $formOnly } }
                                                if ($fallback.Count -gt 0) { $mission = $fallback[0] }
                                }
                }
                if ($null -ne $mission -and $null -eq $soft) {
                                $soft = Get-EventScore -Student $Student -Sortie $Sortie -Evt $mission -SoftWeights $SoftWeights -Instructors $Instructors
                }
                return [PSCustomObject]@{ Mission = $mission; Soft = $soft; Hard = $hr }
}

function Invoke-PinnedLine {
                param ($Sortie, $Student, [array]$Rows, [array]$Students, [array]$Instructors, $SoftWeights, [array]$Formations, [int]$FormIndex, [double]$FormLookAhead, $Overrides)

                $name = Get-NameText $Sortie.Student
                if ($null -eq $Student) {
                                $Sortie.FilledBy = 'Pin'
                                $msg = "student '$name' not found in MasterData"
                                Add-SortieNote -Sortie $Sortie -Note "OVERRIDE: $msg"
                                $Overrides.Add([PSCustomObject]@{ LineId = $Sortie.BoardLineId; ETD = $Sortie.ETD; Student = $name; Violations = @($msg) })
                                return $FormIndex
                }

                $release = @([string]$Sortie.BoardLineId)
                if (Test-IsOBSortieType ([string]$Sortie.SortieType)) {
                                $ret = Get-LinkedReturnRow -Sortie $Sortie -Schedule $Rows -StudentName $Student.Name
                                if ($null -ne $ret) { $release += [string]$ret.BoardLineId }
                }
                Remove-StudentReservation -Student $Student -LineId $release

                $explicitName = Get-NameText $Sortie.Mission
                $pick = Select-PinnedMission -Student $Student -Sortie $Sortie -Rows $Rows -Students $Students -Instructors $Instructors -SoftWeights $SoftWeights -FormLookAhead $FormLookAhead
                $hr = $pick.Hard
                $mission = $pick.Mission
                $soft = $pick.Soft

                $FormIndex = Process-HardcodedSortie -Sortie $Sortie -Student $Student -Mission $mission -Schedule $Rows -Students $Students `
                                -Instructors $Instructors -SoftWeights $SoftWeights -Formations $Formations -FormIndex $FormIndex
                if ($null -eq $mission -and $explicitName -ne '') {
                                # A typed mission that isn't in the syllabus: kept as typed. The student still flies an O&B's
                                # return leg (the next event they can fly home, or a note that there is none).
                                $Sortie.Mission = $explicitName
                                if (Test-IsOBSortieType ([string]$Sortie.SortieType)) {
                                                Schedule-ReturnLeg -Sortie $Sortie -Student $Student -Schedule $Rows -Instructors $Instructors -SoftWeights $SoftWeights
                                                $release = @([string]$Sortie.BoardLineId)
                                }
                }
                if ($null -eq $mission) {
                                # Nothing was marked complete, so keep holding the student's time on this line (and its return)
                                # or later lines could double- or triple-book them
                                foreach ($rid in $release) {
                                                $held = $Rows | Where-Object { [string]$_.BoardLineId -eq $rid } | Select-Object -First 1
                                                if ($null -ne $held) {
                                                                Add-StudentReservation -Student $Student -LineId $rid -ETD $held.ETD -SortieType ([string]$held.SortieType) -Mission $(if ($rid -eq [string]$Sortie.BoardLineId) { $explicitName } else { '' }) -Instructor (Get-NameText $Sortie.Instructor)
                                                }
                                }
                }

                foreach ($note in (Get-ChosenEventNotes -Hard $hr -Evt $mission -Soft $soft)) { Add-SortieNote -Sortie $Sortie -Note $note }

                $hard = New-Object System.Collections.Generic.List[string]
                foreach ($v in @($hr.Violations)) { if ($v.Severity -eq 'Hard') { $hard.Add($v.Message) } }
                if ($null -ne $mission) {
                                $er = Find-EventResult -Hard $hr -Evt $mission
                                if ($null -ne $er) { foreach ($v in @($er.Violations)) { if ($v.Severity -eq 'Hard') { $hard.Add($v.Message) } } }
                }
                elseif ($explicitName -eq '') {
                                # (A typed mission that isn't in the syllabus is kept as typed; MISSION_NOT_FOUND says why)
                                $hard.Add('no eligible event - mission left blank')
                }
                $hardList = @($hard | Select-Object -Unique)
                if ($hardList.Count -gt 0) {
                                Add-SortieNote -Sortie $Sortie -Note ("OVERRIDE: " + ($hardList -join '; '))
                                $Overrides.Add([PSCustomObject]@{ LineId = $Sortie.BoardLineId; ETD = $Sortie.ETD; Student = $Student.Name; Violations = $hardList })
                }
                return $FormIndex
}

# One open line: run every student through the hard rules, score the survivors,
# put the best one on the line
function Invoke-OpenLine {
                param ($Sortie, [array]$Rows, [array]$Students, [array]$Instructors, $SoftWeights, [array]$Formations, [int]$FormIndex, [double]$FormLookAhead, [switch]$Trace)

                $lineContext = @{}
                $best = $null
                $traceList = New-Object System.Collections.Generic.List[object]

                foreach ($Student in $Students) {
                                if ([string]::IsNullOrWhiteSpace([string]$Student.Name)) { continue }
                                $hr = Test-HardRules -Student $Student -Sortie $Sortie -Schedule $Rows -Students $Students -Instructors $Instructors `
                                                -FormLookAhead $FormLookAhead -LineContext $lineContext
                                if (-not $hr.Passed) { continue }

                                foreach ($er in @($hr.EventResults | Where-Object { $_.Passed })) {
                                                $sr = Get-EventScore -Student $Student -Sortie $Sortie -Evt $er.Event -SoftWeights $SoftWeights -Instructors $Instructors
                                                $cand = [PSCustomObject]@{ Student = $Student; Event = $er.Event; Score = $sr.Score; Soft = $sr; Hard = $hr }
                                                if ($Trace) { $traceList.Add($cand) }
                                                if (Test-CandidateBetter $cand $best) { $best = $cand }
                                }
                }

                if ($Trace) { Trace-SchedulingDecision -Sortie $Sortie -Candidates $traceList.ToArray() -Best $best }

                # ---------------- UNABLE TO FILL ----------------
                # Puts the line up for sale if we don't have any valid students to put in the line
                if ($null -eq $best) {
                                Write-Host "No available Student for $($Sortie.ETD) : $($Sortie.SortieType)"
                                $Sortie.Student = 'For Sale'
                                $Sortie.Mission = $null
                                $Sortie.FilledBy = 'ForSale'
                                Add-SortieNote -Sortie $Sortie -Note 'No Available SP'
                                if ([string]$Sortie.SortieType -notlike '*FT*') {
                                                Add-SortieNote -Sortie $Sortie -Note 'Potentially fill if we have another line'
                                }
                                return $FormIndex
                }

                Write-Host "The Selected Student's Score is $($best.Score)"
                $Sortie.Student = $best.Student
                $Sortie.Mission = $best.Event
                $Sortie.FilledBy = 'Engine'
                [void](Complete-StudentMission -Student $best.Student -Mission $best.Event -InstructorName (Get-NameText $Sortie.Instructor) -ETD $Sortie.ETD -SoftWeights $SoftWeights -Sortie $Sortie)
                foreach ($note in (Get-ChosenEventNotes -Hard $best.Hard -Evt $best.Event -Soft $best.Soft)) { Add-SortieNote -Sortie $Sortie -Note $note }

                if ([string]$best.Event.Type -like '*Form*') {
                                $FormIndex = Pair-FormationWingman -PrimarySortie $Sortie -PrimaryStudent $best.Student -PrimaryMission $best.Event `
                                                -Schedule $Rows -Students $Students -Instructors $Instructors -SoftWeights $SoftWeights -FormIndex $FormIndex -Formations $Formations
                }
                if (Test-IsOBSortieType ([string]$Sortie.SortieType)) {
                                Schedule-ReturnLeg -Sortie $Sortie -Student $best.Student -Schedule $Rows -Instructors $Instructors -SoftWeights $SoftWeights
                }
                return $FormIndex
}

function Trace-SchedulingDecision {
                param (
                                $Sortie,
                                [array]$Candidates,
                                $Best
                )

                Write-Host "`n=== DIAGNOSTIC: Scheduling Decision for ETD $($Sortie.ETD) ===" -ForegroundColor Yellow
                Write-Host "Eligible candidates: $(@($Candidates).Count)"
                foreach ($c in ($Candidates | Sort-Object -Property Score -Descending)) {
                                Write-Host ("  - {0} | {1} | Score: {2}" -f $c.Student.Name, (Get-NameText $c.Event), $c.Score)
                }
                if ($null -ne $Best) {
                                Write-Host "--> WINNER: $($Best.Student.Name) on $(Get-NameText $Best.Event) with Score $($Best.Score)" -ForegroundColor Green
                }
                Write-Host "===========================================================`n" -ForegroundColor Yellow
}
