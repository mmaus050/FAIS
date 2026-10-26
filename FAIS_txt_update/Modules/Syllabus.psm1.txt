# SYLLABUS.PSM1 DIRECTORY:
#             Everything about a student's syllabus: which events they can fly next,
#             marking events complete, sim debt, and the timeline of what they fly
# Test-MissionComplete / Get-SyllabusSortKey / Find-SyllabusMission / Get-MissionKind
# Get-OptedEvents / Get-AllowedEventTypes / Get-ValidOptedEvents
# Get-SimDebtDelta / Get-MissionSimDebtDelta / Update-StudentSimDebtFromSyllabus
# Complete-StudentMission / Undo-StudentMission / Reset-PlannedCompletions
# Get-UnsatSimKind / Get-LastSimType / Clear-CleanedUnsat / Restore-PlannedUnsat
# Add-StudentReservation / Remove-StudentReservation / Get-ReservedMissionNames
# Get-StudentTimeline
# Measure-String / Get-StringLength / Get-StringLengths / Get-StringCounts (String Tracker)
# Get-MissionCategory / Get-MissionPhase / Get-StringFlights / Get-StudentNonCounters
# Get-StringOutlook / Get-StringPlacement (string and counter rules for the scheduler)


# =================================================================================
#  SYLLABUS LOOKUPS
# =================================================================================

function Test-MissionComplete {
                param ($Mission)

                if ($null -eq $Mission) { return $false }
                return (ConvertTo-FaisBool $Mission.IsComplete)
}

# SyllabusID can come in as a number or text; sort numerically when possible
function Get-SyllabusSortKey {
                param ($Mission)

                $raw = $Mission.SyllabusID
                $num = 0.0
                if ($null -ne $raw -and [double]::TryParse(([string]$raw).Trim(), [ref]$num)) { return $num }
                return [double]::MaxValue
}

function Find-SyllabusMission {
                param (
                                $Student,
                                $Mission
                )

                $name = Get-NameText $Mission
                if ($null -eq $Student -or [string]::IsNullOrWhiteSpace($name)) { return $null }
                foreach ($m in @($Student.Syllabus)) {
                                if ($null -ne $m -and ([string]$m.Mission).Trim() -ieq $name) { return $m }
                }
                return $null
}

# 'Sim' for OFT/IFT syllabus events, 'Flight' for everything else
function Get-MissionKind {
                param (
                                $Mission,
                                [string]$InstructorName
                )

                $type = [string](Get-OptionalProperty $Mission 'Type' '')
                if ($type -match '(?i)^\s*(IFT|OFT)\s*$') { return 'Sim' }
                if (-not [string]::IsNullOrWhiteSpace($type)) { return 'Flight' }
                if ($InstructorName -like '*SIM IP*') { return 'Sim' }
                return 'Flight'
}


# =================================================================================
#  OPTED EVENTS
# =================================================================================

# Finds the events a student is "opted" for: incomplete, with every prereq complete.
# -AssumeComplete treats extra missions as complete (used for O&B look-ahead).
# -ExcludeMissions drops missions already reserved on another line.
function Get-OptedEvents {
                param (
                                [AllowEmptyCollection()]
                                [array]$SyllabusData,
                                [string[]]$AssumeComplete = @(),
                                [string[]]$ExcludeMissions = @()
                )

                $completionStatus = @{}
                foreach ($row in @($SyllabusData)) {
                                if ($null -eq $row) { continue }
                                $completionStatus[([string]$row.Mission).Trim().ToUpperInvariant()] = (Test-MissionComplete $row)
                }
                foreach ($name in $AssumeComplete) {
                                if (-not [string]::IsNullOrWhiteSpace($name)) { $completionStatus[$name.Trim().ToUpperInvariant()] = $true }
                }
                $excluded = @{}
                foreach ($name in $ExcludeMissions) {
                                if (-not [string]::IsNullOrWhiteSpace($name)) { $excluded[$name.Trim().ToUpperInvariant()] = $true }
                }

                $optedEvents = New-Object System.Collections.Generic.List[object]
                foreach ($row in @($SyllabusData)) {
                                if ($null -eq $row) { continue }
                                $key = ([string]$row.Mission).Trim().ToUpperInvariant()
                                if ($key -eq '' -or $completionStatus[$key] -or $excluded.ContainsKey($key)) { continue }

                                $prereqsMet = $true
                                $prereqText = [string]$row.Prereqs
                                if ($prereqText.Trim() -ne '0' -and -not [string]::IsNullOrWhiteSpace($prereqText)) {
                                                foreach ($prereq in ($prereqText -split ',')) {
                                                                $p = $prereq.Trim().ToUpperInvariant()
                                                                if ($p -eq '') { continue }
                                                                if (-not $completionStatus[$p]) { $prereqsMet = $false; break }
                                                }
                                }

                                if ($prereqsMet) { $optedEvents.Add($row) }
                }

                return $optedEvents.ToArray()
}

# Which syllabus event Types may be flown on a sortie type ($null = none)
function Get-AllowedEventTypes {
                param ([string]$SortieType)

                if ($SortieType -match '(?i)O\s*&\s*B') { return @('Single', 'Formation', 'Both') }
                if ($SortieType -like '*Day Local*') { return @('Single', 'SingleSolo', 'Formation', 'FormSolo', 'Both') }
                if ($SortieType -like '*IFT*') { return @('IFT') }
                if ($SortieType -like '*OFT*') { return @('IFT', 'OFT') }
                return $null
}

# Opted events that fit the sortie type, minus missions reserved on other lines.
# For an O&B, also requires a flyable second leg after the first event.
# Never mutates the syllabus.
function Get-ValidOptedEvents {
                param (
                                $Student,
                                $Sortie,
                                [string[]]$AssumeComplete = @(),
                                [switch]$NoLookahead
                )

                if ($null -eq $Student) { return }
                $types = Get-AllowedEventTypes ([string]$Sortie.SortieType)
                if ($null -eq $types) { return }

                $lineId = [string](Get-OptionalProperty $Sortie 'BoardLineId' '')
                $reserved = @(Get-ReservedMissionNames -Student $Student -ExceptLineId $lineId)

                $all = Get-OptedEvents -SyllabusData $Student.Syllabus -AssumeComplete $AssumeComplete -ExcludeMissions $reserved
                $OptedEvents = @($all | Where-Object { ([string]$_.Type).Trim() -in $types } | Sort-Object -Property { Get-SyllabusSortKey $_ })
                if ($OptedEvents.Count -eq 0) { return }

                # O&B DOUBLE RIDE PREREQ (LOOK-AHEAD): the return leg must have an event to fly
                if (-not $NoLookahead -and (Test-IsOBSortieType ([string]$Sortie.SortieType))) {
                                $first = [string]$OptedEvents[0].Mission
                                $next = Get-OptedEvents -SyllabusData $Student.Syllabus -AssumeComplete (@($AssumeComplete) + $first) -ExcludeMissions $reserved
                                $second = @($next | Where-Object { ([string]$_.Type).Trim() -in @('Single', 'Formation', 'Both') })
                                if ($second.Count -eq 0) { return }
                }

                return $OptedEvents
}


# =================================================================================
#  SIM DEBT
#  Every flight creates "debt" that sims pay down.
# =================================================================================

function Get-SimDebtDelta {
                param (
                                [string]$MissionName,
                                [double]$SimStopMult,
                                [double]$TransSimRatio,
                                [double]$FormSimRatio
                )

                # Flights increase debt
                if ($MissionName -like '*FT5*' -or $MissionName -like '*FF5*' -or $MissionName -like '*AA5*') { return $SimStopMult }
                # Transition sims decrease debt
                if ($MissionName -like '*FI3*' -or $MissionName -like '*FE3*' -or $MissionName -like '*FI4*' -or $MissionName -like '*FE4*') { return -$TransSimRatio }
                # Formation sims decrease debt
                if ($MissionName -like '*FN3*' -or $MissionName -like '*FN4*' -or $MissionName -like '*AA4*') { return -$FormSimRatio }
                return 0.0
}

function Get-MissionSimDebtDelta {
                param (
                                [string]$MissionName,
                                $SoftWeights
                )

                if ($null -eq $SoftWeights) { return 0.0 }
                return (Get-SimDebtDelta -MissionName $MissionName `
                                -SimStopMult ([double](Get-OptionalProperty $SoftWeights 'SimStopMult' 0)) `
                                -TransSimRatio ([double](Get-OptionalProperty $SoftWeights 'TransSimRatio' 0)) `
                                -FormSimRatio ([double](Get-OptionalProperty $SoftWeights 'FormSimRatio' 0)))
}

# Recomputes SimDebt from every completed syllabus event
function Update-StudentSimDebtFromSyllabus {
                param(
                                [Parameter(Mandatory)]
                                $Student,
                                [Parameter(Mandatory)]
                                $SoftWeights
                )

                if ($null -eq $Student -or $null -eq $Student.Syllabus) {
                                Write-Warning "Cannot recalculate SimDebt: student or syllabus is missing."
                                return $null
                }

                foreach ($propertyName in @('SimStopMult', 'TransSimRatio', 'FormSimRatio')) {
                                if ($null -eq $SoftWeights.$propertyName -or $SoftWeights.$propertyName -eq '') {
                                                Write-Warning "Cannot recalculate SimDebt for '$($Student.Name)': SoftWeights.$propertyName is missing."
                                                return $null
                                }
                }

                $debt = 0.0
                foreach ($mission in $Student.Syllabus) {
                                if (-not (Test-MissionComplete $mission)) { continue }
                                $missionName = [string]$mission.Mission
                                if ([string]::IsNullOrWhiteSpace($missionName)) { continue }
                                $debt += Get-MissionSimDebtDelta -MissionName $missionName -SoftWeights $SoftWeights
                }

                $Student | Add-Member -MemberType NoteProperty -Name 'SimDebt' -Value ([math]::Round($debt, 2)) -Force
                return $Student.SimDebt
}


# =================================================================================
#  COMPLETING EVENTS
# =================================================================================

# The one place a scheduled event is written onto a student's syllabus.
# Marks it complete with the IP name and ETD, and applies its sim debt once.
# -SortieType (or -Sortie) is the line it is flown on: an unsat that line cleans up is
# cleared, and -Sortie gets a note saying so.
function Complete-StudentMission {
                param (
                                [Parameter(Mandatory)] $Student,
                                [Parameter(Mandatory)] $Mission,
                                [string]$InstructorName,
                                $ETD,
                                $SoftWeights,
                                [string]$SortieType = '',
                                $Sortie,
                                [switch]$NoSimDebt
                )

                $entry = Find-SyllabusMission -Student $Student -Mission $Mission
                if ($null -eq $entry) { return $null }

                $wasComplete = Test-MissionComplete $entry
                $entry.IsComplete = $true
                $entry.Instructor = $InstructorName
                $dt = Parse-ETD $ETD
                $entry.ETD = if ($null -ne $dt) { $dt.ToString('yyyy-MM-dd HH:mm:ss') } else { [string]$ETD }

                if (-not $wasComplete -and -not $NoSimDebt -and $null -ne $SoftWeights) {
                                $current = 0.0
                                if ($null -ne $Student.SimDebt -and [string]$Student.SimDebt -ne '') { $current = [double]$Student.SimDebt }
                                $Student.SimDebt = [math]::Round($current + (Get-MissionSimDebtDelta -MissionName ([string]$entry.Mission) -SoftWeights $SoftWeights), 2)
                }

                if ($SortieType -eq '' -and $null -ne $Sortie) { $SortieType = [string]$Sortie.SortieType }
                if ($SortieType -ne '') {
                                $cleared = Clear-CleanedUnsat -Student $Student -SortieType $SortieType
                                if ($null -ne $cleared -and $null -ne $Sortie) { Add-SortieNote -Sortie $Sortie -Note $cleared }
                }

                $rev = [int](Get-OptionalProperty $Student 'SyllabusRevision' 0)
                $Student | Add-Member -MemberType NoteProperty -Name 'SyllabusRevision' -Value ($rev + 1) -Force
                return $entry
}

# Takes a completed event back off a student's syllabus (logged incomplete, or a planned
# event being re-planned) and backs out only that event's sim debt, so the SimDebt read
# from MasterData keeps any hand adjustment. Returns $true when the event was complete.
function Undo-StudentMission {
                param (
                                [Parameter(Mandatory)] $Student,
                                [Parameter(Mandatory)] $Mission,
                                $SoftWeights
                )

                # The exact syllabus entry when one is passed, else the first with that name
                $entry = $null
                foreach ($m in @($Student.Syllabus)) { if ([object]::ReferenceEquals($m, $Mission)) { $entry = $m; break } }
                if ($null -eq $entry) { $entry = Find-SyllabusMission -Student $Student -Mission $Mission }
                if ($null -eq $entry -or -not (Test-MissionComplete $entry)) { return $false }

                $entry.IsComplete = $false
                $entry.Instructor = ''
                $entry.ETD = ''
                if ($null -ne $SoftWeights) {
                                $current = 0.0
                                if ($null -ne $Student.SimDebt -and [string]$Student.SimDebt -ne '') { $current = [double]$Student.SimDebt }
                                $Student.SimDebt = [math]::Round($current - (Get-MissionSimDebtDelta -MissionName ([string]$entry.Mission) -SoftWeights $SoftWeights), 2)
                }

                $rev = [int](Get-OptionalProperty $Student 'SyllabusRevision' 0)
                $Student | Add-Member -MemberType NoteProperty -Name 'SyllabusRevision' -Value ($rev + 1) -Force
                return $true
}

# Clears completions dated on/after a day. A previous run that scheduled these days
# wrote them as complete; they are plans, not flown events, when the days are re-built.
# Their sim debt is backed out, and an unsat that one of those planned lines cleaned up
# (its "Clears Unsat" note in -PrevSchedule) goes back on the student.
# Returns the students that changed.
function Reset-PlannedCompletions {
                param (
                                [array]$Students,
                                [datetime]$FromDate,
                                $SoftWeights,
                                [array]$PrevSchedule = @()
                )

                $changed = New-Object System.Collections.Generic.List[object]
                foreach ($s in @($Students)) {
                                $hit = $false
                                foreach ($m in @($s.Syllabus)) {
                                                if (-not (Test-MissionComplete $m)) { continue }
                                                $dt = Parse-ETD $m.ETD
                                                if ($null -ne $dt -and $dt.Date -ge $FromDate.Date) {
                                                                if (Undo-StudentMission -Student $s -Mission $m -SoftWeights $SoftWeights) { $hit = $true }
                                                }
                                }
                                if ($hit) { $changed.Add($s) }
                }

                foreach ($row in @($PrevSchedule)) {
                                if ($null -eq $row) { continue }
                                $dt = Parse-ETD $row.ETD
                                if ($null -eq $dt -or $dt.Date -lt $FromDate.Date) { continue }
                                $name = Get-NameText $row.Student
                                if (Test-IsSentinelStudent $name) { continue }
                                $s = Resolve-StudentByName -Students $Students -Name $name
                                if ($null -ne $s -and (Restore-PlannedUnsat -Student $s -Notes ([string]$row.Notes)) -and -not $changed.Contains($s)) { $changed.Add($s) }
                }
                return $changed.ToArray()
}


# =================================================================================
#  UNSAT EVENTS
#  Unsat sim: no flying until a re-sim. An OFT hook needs an OFT; an IFT hook can be
#  cleaned up in an OFT or an IFT. Unsat flight: no sims until a flight.
#  UnsatSimType ('OFT' / 'IFT') says which sim was hooked; unknown is treated as OFT.
# =================================================================================

function Get-UnsatSimKind {
                param ($Student)

                $type = [string](Get-OptionalProperty $Student 'UnsatSimType' '')
                if ($type -match '(?i)^\s*IFT\s*$') { return 'IFT' }
                return 'OFT'
}

# 'OFT' / 'IFT' for the student's latest completed sim event up to -Before ('' if none).
# MasterData only holds UnsatSim TRUE/FALSE, so this stands in for the hooked sim.
function Get-LastSimType {
                param (
                                $Student,
                                [datetime]$Before = (Get-Date)
                )

                $last = $null
                $lastStart = $null
                foreach ($m in @($Student.Syllabus)) {
                                if (-not (Test-MissionComplete $m)) { continue }
                                $start = Parse-ETD $m.ETD
                                if ($null -eq $start -or $start -gt $Before) { continue }
                                if ((Get-MissionKind -Mission $m -InstructorName (Get-NameText $m.Instructor)) -ne 'Sim') { continue }
                                if ($null -eq $lastStart -or $start -gt $lastStart) { $last = $m; $lastStart = $start }
                }
                if ($null -eq $last) { return '' }
                $type = ([string]$last.Type).Trim().ToUpperInvariant()
                if ($type -in @('OFT', 'IFT')) { return $type }
                return ''
}

# Clears the unsat a line of -SortieType cleans up. Returns the note for that line, or $null.
function Clear-CleanedUnsat {
                param (
                                [Parameter(Mandatory)] $Student,
                                [string]$SortieType
                )

                if (Test-IsSimSortieType $SortieType) {
                                if (-not (ConvertTo-FaisBool $Student.UnsatSim)) { return $null }
                                $kind = Get-UnsatSimKind $Student
                                if ($kind -eq 'OFT' -and $SortieType -notlike '*OFT*') { return $null }
                                $Student.UnsatSim = $false
                                return "Clears Unsat Sim ($kind)"
                }
                if (-not (ConvertTo-FaisBool $Student.UnsatFlight)) { return $null }
                $Student.UnsatFlight = $false
                return 'Clears Unsat Flight'
}

# Puts an unsat back on the student when the planned line that cleared it is re-planned.
# Returns $true when a flag was set.
function Restore-PlannedUnsat {
                param (
                                [Parameter(Mandatory)] $Student,
                                [string]$Notes
                )

                $hit = $false
                if ($Notes -match 'Clears Unsat Sim \((OFT|IFT)\)' -and -not (ConvertTo-FaisBool $Student.UnsatSim)) {
                                $Student | Add-Member -NotePropertyName 'UnsatSim' -NotePropertyValue $true -Force
                                $Student | Add-Member -NotePropertyName 'UnsatSimType' -NotePropertyValue $Matches[1].ToUpperInvariant() -Force
                                $hit = $true
                }
                if ($Notes -match 'Clears Unsat Flight' -and -not (ConvertTo-FaisBool $Student.UnsatFlight)) {
                                $Student | Add-Member -NotePropertyName 'UnsatFlight' -NotePropertyValue $true -Force
                                $hit = $true
                }
                return $hit
}


# =================================================================================
#  RESERVATIONS
#  A hardcoded (pinned) line that has not been processed yet holds the student's
#  time so earlier lines respect it (turn time, triple turn, crew rest).
# =================================================================================

function Add-StudentReservation {
                param (
                                [Parameter(Mandatory)] $Student,
                                [Parameter(Mandatory)] [string]$LineId,
                                $ETD,
                                [string]$SortieType,
                                [string]$Mission,
                                [string]$Instructor
                )

                $start = Parse-ETD $ETD
                if ($null -eq $start) { return }
                if ($null -eq $Student.PSObject.Properties['Reservations'] -or $null -eq $Student.Reservations) {
                                $Student | Add-Member -MemberType NoteProperty -Name 'Reservations' -Value @() -Force
                }

                $kept = @($Student.Reservations | Where-Object { $_.LineId -ne $LineId })
                $kept += [PSCustomObject]@{
                                LineId     = $LineId
                                Start      = $start
                                SortieType = $SortieType
                                Mission    = $Mission
                                Instructor = $Instructor
                                IsSim      = (Test-IsSimSortieType $SortieType)
                }
                $Student.Reservations = $kept
}

function Remove-StudentReservation {
                param (
                                [Parameter(Mandatory)] $Student,
                                [string[]]$LineId
                )

                if ($null -eq $Student.PSObject.Properties['Reservations'] -or $null -eq $Student.Reservations) { return }
                $Student.Reservations = @($Student.Reservations | Where-Object { $LineId -notcontains $_.LineId })
}

function Get-ReservedMissionNames {
                param (
                                $Student,
                                [string]$ExceptLineId = ''
                )

                if ($null -eq $Student -or $null -eq $Student.PSObject.Properties['Reservations']) { return @() }
                return @($Student.Reservations |
                                Where-Object { $_.LineId -ne $ExceptLineId -and -not [string]::IsNullOrWhiteSpace($_.Mission) } |
                                ForEach-Object { $_.Mission })
}


# =================================================================================
#  TIMELINE
# =================================================================================

# Every dated event a student has: completed syllabus events plus reservations,
# oldest first. -ExceptLineId leaves out the line being evaluated.
function Get-StudentTimeline {
                param (
                                $Student,
                                [string]$ExceptLineId = '',
                                [switch]$NoReservations
                )

                $items = New-Object System.Collections.Generic.List[object]
                if ($null -eq $Student) { return }

                foreach ($m in @($Student.Syllabus)) {
                                if (-not (Test-MissionComplete $m)) { continue }
                                $start = Parse-ETD $m.ETD
                                if ($null -eq $start) { continue }
                                $ipName = Get-NameText $m.Instructor
                                $items.Add([PSCustomObject]@{
                                                Start         = $start
                                                Mission       = ([string]$m.Mission).Trim()
                                                Instructor    = $ipName
                                                IsSim         = ((Get-MissionKind -Mission $m -InstructorName $ipName) -eq 'Sim')
                                                IsReservation = $false
                                                LineId        = ''
                                                SortieType    = ''
                                })
                }

                if (-not $NoReservations -and $null -ne $Student.PSObject.Properties['Reservations']) {
                                foreach ($r in @($Student.Reservations)) {
                                                if ($null -eq $r -or $r.LineId -eq $ExceptLineId) { continue }
                                                $items.Add([PSCustomObject]@{
                                                                Start         = $r.Start
                                                                Mission       = [string]$r.Mission
                                                                Instructor    = [string]$r.Instructor
                                                                IsSim         = [bool]$r.IsSim
                                                                IsReservation = $true
                                                                LineId        = $r.LineId
                                                                SortieType    = $r.SortieType
                                                })
                                }
                }

                return @($items | Sort-Object -Property Start)
}


# =================================================================================
#  STRINGS AND COUNTERS (String Tracker, and the scheduler's string rules)
#    AETCI instructor continuity: an IP who flies a sortie with a student is a counter and
#    that sortie is part of a string. Three sorties in a row with one counter is the goal;
#    a 2 or 4-sortie string is acceptable (to prevent BITs).
#    - Each category (FT, FF, AA) is its own sequence, in mission order: FT5001, FT5002, ...
#      even when rides from another category were flown in between
#    - Non-counters (supervisors) are skipped, any number in a row, unless they flew the student
#      two rides in a row (strictly back to back): then they are a counter for that student
#    - One different counter between two of an IP's rides does not break that IP's string
#      (A,A,B,A is a 3-string with A); two different counters in a row end it
#    - Sims, solos and check rides (IsCheck) are left out entirely
#    - Counter limits per phase: 6 in Fundamental (FT, FF), 3 in Airmanship (AA)
# =================================================================================

$script:CounterLimits = @{ Fundamental = 6; Airmanship = 3 }

# Walks back one ride at a time; each step is a recursive call.
# $History is newest first: $History[0] is IP1 (current ride), $History[1] is IP2, ...
# $SkippedOne is True once a different counter has been passed over since IP1's last ride.
function Measure-String {
                param ([string[]]$History, [string[]]$NonCounters, [int]$n, [bool]$SkippedOne)

                $ip = $History[$n]  # $null once we run out of rides, which answers False to everything

                # IP1 = IPn?  True: one more ride in the string, and the skip resets
                if ($ip -eq $History[0]) { return 1 + (Measure-String $History $NonCounters ($n + 1) $false) }

                # Is IPn a non-counter?  True: skip it
                if ($ip -in $NonCounters) { return Measure-String $History $NonCounters ($n + 1) $SkippedOne }

                # First different counter: look past it
                if ($null -ne $ip -and -not $SkippedOne) { return Measure-String $History $NonCounters ($n + 1) $true }

                # Second different counter in a row, or no more rides: the string ends
                return 0
}

# Length of the last ride's string. $Sorties is IP names oldest to newest: A,A,B,A,A,A -> 5
function Get-StringLength {
                param ([string[]]$Sorties, [string[]]$NonCounters = @())

                if ($Sorties.Count -eq 0) { return 0 }
                [string[]]$History = $Sorties[($Sorties.Count - 1)..0]  # typed, so one ride stays an array
                if ($History[0] -in $NonCounters) { return 0 }  # a non-counter's ride is not part of a string
                return 1 + (Measure-String $History $NonCounters 1 $false)
}

# Length of every string in one category's rides (IP names in mission order). A,A,B,A,A,A -> 5, 1
function Get-StringLengths {
                param ([string[]]$Sorties, [string[]]$NonCounters = @())

                # String length at each ride, counted back from that ride
                $Length = @(for ($i = 0; $i -lt $Sorties.Count; $i++) { Get-StringLength $Sorties[0..$i] $NonCounters })

                # A string is finished at a ride when that IP's next ride starts over at 1, or the IP never flies the category again
                for ($i = 0; $i -lt $Sorties.Count; $i++) {
                                if ($Length[$i] -eq 0) { continue }  # non-counter
                                $next = $null
                                for ($j = $i + 1; $j -lt $Sorties.Count; $j++) {
                                                if ($Sorties[$j] -eq $Sorties[$i]) { $next = $j; break }
                                }
                                if ($null -eq $next -or $Length[$next] -eq 1) { $Length[$i] }
                }
}

# 'FT5103' -> Category FT, Number 5103 (no number sorts last)
function Get-MissionCategory {
                param ([string]$Mission)

                $null = $Mission.Trim() -match '^(\D*)(\d*)'
                return [PSCustomObject]@{
                                Category = $Matches[1].ToUpper()
                                Number   = if ($Matches[2] -ne '') { [double]$Matches[2] } else { [double]::MaxValue }
                }
}

# The counter-limit phase of a category: Airmanship for AA, Fundamental for everything else (FT, FF)
function Get-MissionPhase {
                param ([string]$Category)

                if ($Category -ieq 'AA') { return 'Airmanship' }
                return 'Fundamental'
}

# A ride that can be part of a string: a flight with an IP in the cockpit, not a sim, solo or check ride
function Test-StringRide {
                param ($Mission, [string]$Instructor)

                if ($null -eq $Mission) { return $false }
                $ipName = Get-NameText $Instructor
                return ((-not (ConvertTo-FaisBool $Mission.IsCheck)) -and
                                [string]$Mission.Type -notlike '*Solo*' -and
                                (Get-MissionKind -Mission $Mission -InstructorName $ipName) -eq 'Flight' -and
                                -not [string]::IsNullOrWhiteSpace($ipName))
}

function New-StringFlight {
                param ($Mission, [string]$Instructor)

                $cat = Get-MissionCategory ([string]$Mission.Mission)
                return [PSCustomObject]@{
                                Mission  = ([string]$Mission.Mission).Trim()
                                Category = $cat.Category
                                Number   = $cat.Number
                                ID       = Get-SyllabusSortKey $Mission
                                Phase    = Get-MissionPhase $cat.Category
                                IP       = Get-BaseName (Get-NameText $Instructor)
                }
}

# A student's completed string rides (Test-StringRide), by category and in mission order inside each.
# -Before keeps only rides with an ETD before then.
function Get-StringFlights {
                param ([array]$Syllabus, $Before = $null)

                # A loop, not a pipeline: the rules call this for every student on every line
                $rides = New-Object System.Collections.Generic.List[object]
                foreach ($m in $Syllabus) {
                                if (-not (Test-MissionComplete $m)) { continue }
                                $ipName = Get-NameText $m.Instructor
                                if (-not (Test-StringRide $m $ipName)) { continue }
                                if ($null -ne $Before) {
                                                $etd = Parse-ETD $m.ETD
                                                if ($null -eq $etd -or $etd -ge $Before) { continue }
                                }
                                $rides.Add((New-StringFlight $m $ipName))
                }
                return @($rides | Sort-Object Category, Number, ID)
}

# A student's non-counters: the Instructors sheet's non-counters (a blank IsCounter counts as a
# counter) and -Extra, except any who flew the student two rides in a row in a category.
# $Flights: Get-StringFlights (in order)
function Get-StudentNonCounters {
                param ([array]$Flights = @(), [array]$Instructors = @(), [string[]]$Extra = @())

                # IPs who flew the student two rides in a row in a category (strictly back to back, not zippered)
                $BackToBack = @(
                                for ($i = 1; $i -lt $Flights.Count; $i++) {
                                                if ($Flights[$i].Category -eq $Flights[$i - 1].Category -and $Flights[$i].IP -eq $Flights[$i - 1].IP) { $Flights[$i].IP }
                                }
                )

                $listed = @($Instructors | Where-Object { -not (ConvertTo-FaisBool $_.IsCounter -Default $true) } | ForEach-Object { Get-BaseName (Get-NameText $_.Name) })
                return @(@($listed) + @($Extra) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and $_ -notin $BackToBack } | Sort-Object -Unique)
}

# A student's strings and counters from their syllabus and the Instructors list.
# Returns @(5-strings, 4-strings, 3-strings, 2-strings, 1-strings, counters). Strings longer than 5 count as 5-strings.
# Counters are the unique IPs who are not non-counters.
function Get-StringCounts {
                param ([array]$Syllabus, [array]$Instructors = @())

                $Flights = @(Get-StringFlights -Syllabus $Syllabus)
                $NonCounters = @(Get-StudentNonCounters -Flights $Flights -Instructors $Instructors)

                # Strings, counted inside each category
                $Strings = @{ 5 = 0; 4 = 0; 3 = 0; 2 = 0; 1 = 0 }
                foreach ($Group in ($Flights | Group-Object Category)) {
                                foreach ($Len in (Get-StringLengths ([string[]]@($Group.Group | ForEach-Object { $_.IP })) $NonCounters)) {
                                                $Strings[[Math]::Min($Len, 5)]++
                                }
                }

                $Counters = @($Flights | ForEach-Object { $_.IP } | Where-Object { $_ -notin $NonCounters } | Sort-Object -Unique).Count

                return @($Strings[5], $Strings[4], $Strings[3], $Strings[2], $Strings[1], $Counters)
}

# What one more ride with counter $IP does to a category's strings. $Sorties is the category's IP
# names oldest to newest, before this ride.
#   Length       the IP's string with this ride (1: a new string)
#   Zipper       the string goes on over one different counter (A,A,B + A)
#   LastIP       the counter who flew the ride before (the one zippered over)
#   Alternating  that counter's own ride was a zipper too (A,B,A + B): the pattern the AETCI rules out
#   Broken       strings this ride ends (two different counters in a row): { IP, Length }
#   Bridged      strings this ride sits in the middle of, still alive: { IP, Length }
function Get-StringOutlook {
                param ([string[]]$Sorties = @(), [string]$IP, [string[]]$NonCounters = @())

                [string[]]$after = @(@($Sorties) + $IP)
                $length = Get-StringLength $after $NonCounters

                # Positions of the counters' rides (non-counters are skipped, as they never break a string)
                $counterAt = @(for ($i = 0; $i -lt $Sorties.Count; $i++) { if ($Sorties[$i] -notin $NonCounters) { $i } })
                $lastIP = if ($counterAt.Count -gt 0) { $Sorties[$counterAt[$counterAt.Count - 1]] } else { '' }

                $zipper = ($length -ge 2 -and $lastIP -ne '' -and $lastIP -ne $IP)
                $alternating = $false
                if ($zipper -and $counterAt.Count -ge 2) {
                                $j = $counterAt[$counterAt.Count - 1]
                                $alternating = ($Sorties[$counterAt[$counterAt.Count - 2]] -ne $lastIP) -and ((Get-StringLength $Sorties[0..$j] $NonCounters) -ge 2)
                }

                $broken = New-Object System.Collections.Generic.List[object]
                $bridged = New-Object System.Collections.Generic.List[object]
                foreach ($other in @($counterAt | ForEach-Object { $Sorties[$_] } | Sort-Object -Unique)) {
                                if ($other -eq $IP) { continue }
                                # Still alive before this ride: one more ride with them would add to their string
                                if ((Get-StringLength ([string[]]@(@($Sorties) + $other)) $NonCounters) -lt 2) { continue }
                                $lastAt = @($counterAt | Where-Object { $Sorties[$_] -eq $other })[-1]
                                $entry = [PSCustomObject]@{ IP = $other; Length = (Get-StringLength $Sorties[0..$lastAt] $NonCounters) }
                                if ((Get-StringLength ([string[]]@(@($after) + $other)) $NonCounters) -lt 2) { $broken.Add($entry) } else { $bridged.Add($entry) }
                }

                return [PSCustomObject]@{
                                Length      = $length
                                Zipper      = $zipper
                                LastIP      = $lastIP
                                Alternating = $alternating
                                Broken      = $broken.ToArray()
                                Bridged     = $bridged.ToArray()
                }
}

# What a student flying $Mission with $Instructor does to their strings and counters (scheduler
# rules). $IsCounter is the IP's IsCounter on the Instructors sheet. Only rides before -Before count.
# -Flights: Get-StringFlights for the student, when the caller already has them.
# $null when the ride is never part of a string (a sim, solo or check ride, or no IP). Otherwise:
#   Category, Phase, IP     FF5003 with 'Jones, Bob [Capt]' -> FF, Fundamental, Jones, Bob
#   IsCounter               False for a non-counter who did not fly the category's ride before:
#                           they break no strings and are not a counter
#   TwoInRow                a non-counter who flew the ride before too, so now counts
#   Outlook                 Get-StringOutlook for this ride (counters only)
#   NewCounter              the IP is not yet one of the phase's counters
#   CounterCount, Limit     the phase's counters with this ride, and the phase's limit
function Get-StringPlacement {
                param (
                                $Student,
                                $Mission,
                                [string]$Instructor,
                                [bool]$IsCounter = $true,
                                $Before = $null,
                                [array]$Instructors = @(),
                                [array]$Flights
                )

                if (-not (Test-StringRide $Mission $Instructor)) { return $null }
                $ride = New-StringFlight $Mission $Instructor
                if ($null -eq $Flights) { $Flights = @(Get-StringFlights -Syllabus @($Student.Syllabus) -Before $Before) }
                $past = @($Flights | Where-Object { $_.Mission -ine $ride.Mission })

                # Non-counters with this ride in place, so a non-counter flying the ride before counts now
                $extra = if ($IsCounter) { @() } else { @($ride.IP) }
                $listed = @(Get-StudentNonCounters -Instructors $Instructors -Extra $extra)
                $nonCounters = [string[]]@(Get-StudentNonCounters -Flights @(@($past) + $ride | Sort-Object Category, Number, ID) -Instructors $Instructors -Extra $extra)
                $counts = $ride.IP -notin $nonCounters

                $counters = @($past | Where-Object { $_.Phase -eq $ride.Phase -and $_.IP -notin $nonCounters } | ForEach-Object { $_.IP } | Sort-Object -Unique)
                $newCounter = $counts -and $ride.IP -notin $counters
                $outlook = $null
                if ($counts) {
                                $sorties = [string[]]@($past | Where-Object { $_.Category -eq $ride.Category } | ForEach-Object { $_.IP })
                                $outlook = Get-StringOutlook -Sorties $sorties -IP $ride.IP -NonCounters $nonCounters
                }

                return [PSCustomObject]@{
                                Category     = $ride.Category
                                Phase        = $ride.Phase
                                IP           = $ride.IP
                                IsCounter    = $counts
                                TwoInRow     = ($counts -and $ride.IP -in $listed)
                                Outlook      = $outlook
                                NewCounter   = $newCounter
                                CounterCount = ($counters.Count + [int]$newCounter)
                                Limit        = $script:CounterLimits[$ride.Phase]
                }
}
