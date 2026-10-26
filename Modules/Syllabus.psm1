# SYLLABUS.PSM1 DIRECTORY:
#             Everything about a student's syllabus: which events they can fly next,
#             marking events complete, sim debt, and the timeline of what they fly
# Test-MissionComplete / Get-SyllabusSortKey / Find-SyllabusMission / Get-MissionKind
# Get-OptedEvents / Get-AllowedEventTypes / Get-ValidOptedEvents
# Get-SimDebtDelta / Get-MissionSimDebtDelta / Update-StudentSimDebtFromSyllabus
# Complete-StudentMission / Undo-StudentMission / Reset-PlannedCompletions / Test-ChangesStrings
# Get-UnsatSimKind / Get-LastSimType / Clear-CleanedUnsat / Restore-PlannedUnsat
# Add-StudentReservation / Remove-StudentReservation / Get-ReservedMissionNames
# Get-StudentTimeline
# Measure-String / Get-StringLength / Get-StringLengths / Test-StringRide / Get-StringCounts (String Tracker)
# Get-DiscontinuityWeight / Get-DiscontinuityScore / Add-StudentDiscontinuity / Initialize-StudentDiscontinuity


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
# A ride that counts toward strings moves the discontinuity score when it is flown or its IP
# changes (-Instructors says who the non-counters are; -NoDiscontinuity leaves it alone).
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
                                [array]$Instructors = @(),
                                [switch]$NoSimDebt,
                                [switch]$NoDiscontinuity
                )

                $entry = Find-SyllabusMission -Student $Student -Mission $Mission
                if ($null -eq $entry) { return $null }

                $wasComplete = Test-MissionComplete $entry
                $discBefore = $null
                if (-not $NoDiscontinuity -and $null -ne $SoftWeights -and (Test-ChangesStrings $entry $InstructorName)) {
                                $discBefore = Get-DiscontinuityScore -Syllabus $Student.Syllabus -Instructors $Instructors -SoftWeights $SoftWeights
                }
                $entry.IsComplete = $true
                $entry.Instructor = $InstructorName
                $dt = Parse-ETD $ETD
                $entry.ETD = if ($null -ne $dt) { $dt.ToString('yyyy-MM-dd HH:mm:ss') } else { [string]$ETD }

                if (-not $wasComplete -and -not $NoSimDebt -and $null -ne $SoftWeights) {
                                $current = 0.0
                                if ($null -ne $Student.SimDebt -and [string]$Student.SimDebt -ne '') { $current = [double]$Student.SimDebt }
                                $Student.SimDebt = [math]::Round($current + (Get-MissionSimDebtDelta -MissionName ([string]$entry.Mission) -SoftWeights $SoftWeights), 2)
                }
                if ($null -ne $discBefore) {
                                Add-StudentDiscontinuity -Student $Student -Delta ((Get-DiscontinuityScore -Syllabus $Student.Syllabus -Instructors $Instructors -SoftWeights $SoftWeights) - $discBefore)
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
# event being re-planned) and backs out only that event's sim debt and discontinuity, so the
# SimDebt and DiscontinuityScore read from MasterData keep any hand adjustment
# (-NoDiscontinuity leaves the discontinuity alone). Returns $true when the event was complete.
function Undo-StudentMission {
                param (
                                [Parameter(Mandatory)] $Student,
                                [Parameter(Mandatory)] $Mission,
                                $SoftWeights,
                                [array]$Instructors = @(),
                                [switch]$NoDiscontinuity
                )

                # The exact syllabus entry when one is passed, else the first with that name
                $entry = $null
                foreach ($m in @($Student.Syllabus)) { if ([object]::ReferenceEquals($m, $Mission)) { $entry = $m; break } }
                if ($null -eq $entry) { $entry = Find-SyllabusMission -Student $Student -Mission $Mission }
                if ($null -eq $entry -or -not (Test-MissionComplete $entry)) { return $false }

                $discBefore = $null
                if (-not $NoDiscontinuity -and $null -ne $SoftWeights -and (Test-ChangesStrings $entry -Undo)) {
                                $discBefore = Get-DiscontinuityScore -Syllabus $Student.Syllabus -Instructors $Instructors -SoftWeights $SoftWeights
                }
                $entry.IsComplete = $false
                $entry.Instructor = ''
                $entry.ETD = ''
                if ($null -ne $SoftWeights) {
                                $current = 0.0
                                if ($null -ne $Student.SimDebt -and [string]$Student.SimDebt -ne '') { $current = [double]$Student.SimDebt }
                                $Student.SimDebt = [math]::Round($current - (Get-MissionSimDebtDelta -MissionName ([string]$entry.Mission) -SoftWeights $SoftWeights), 2)
                }
                if ($null -ne $discBefore) {
                                Add-StudentDiscontinuity -Student $Student -Delta ((Get-DiscontinuityScore -Syllabus $Student.Syllabus -Instructors $Instructors -SoftWeights $SoftWeights) - $discBefore)
                }

                $rev = [int](Get-OptionalProperty $Student 'SyllabusRevision' 0)
                $Student | Add-Member -MemberType NoteProperty -Name 'SyllabusRevision' -Value ($rev + 1) -Force
                return $true
}

# Clears completions dated on/after a day. A previous run that scheduled these days
# wrote them as complete; they are plans, not flown events, when the days are re-built.
# Their sim debt and discontinuity are backed out (each student's discontinuity once, for all
# of their events), and an unsat that one of those planned lines cleaned up (its
# "Clears Unsat" note in -PrevSchedule) goes back on the student.
# Returns the students that changed.
function Reset-PlannedCompletions {
                param (
                                [array]$Students,
                                [datetime]$FromDate,
                                $SoftWeights,
                                [array]$PrevSchedule = @(),
                                [array]$Instructors = @()
                )

                $changed = New-Object System.Collections.Generic.List[object]
                foreach ($s in @($Students)) {
                                $hit = $false
                                $discBefore = $null
                                foreach ($m in @($s.Syllabus)) {
                                                if (-not (Test-MissionComplete $m)) { continue }
                                                $dt = Parse-ETD $m.ETD
                                                if ($null -ne $dt -and $dt.Date -ge $FromDate.Date) {
                                                                if ($null -eq $discBefore -and $null -ne $SoftWeights -and (Test-ChangesStrings $m -Undo)) {
                                                                                $discBefore = Get-DiscontinuityScore -Syllabus $s.Syllabus -Instructors $Instructors -SoftWeights $SoftWeights
                                                                }
                                                                if (Undo-StudentMission -Student $s -Mission $m -SoftWeights $SoftWeights -NoDiscontinuity) { $hit = $true }
                                                }
                                }
                                if ($null -ne $discBefore) {
                                                Add-StudentDiscontinuity -Student $s -Delta ((Get-DiscontinuityScore -Syllabus $s.Syllabus -Instructors $Instructors -SoftWeights $SoftWeights) - $discBefore)
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

# True when marking the event complete with -InstructorName (or, with -Undo, taking it back)
# can change the student's strings and counters, and so their discontinuity score
function Test-ChangesStrings {
                param ($Mission, [string]$InstructorName, [switch]$Undo)

                $wasComplete = Test-MissionComplete $Mission
                $oldIP = Get-NameText $Mission.Instructor
                $wasRide = $wasComplete -and (Test-StringRide $Mission $oldIP)
                if ($Undo) { return $wasRide }
                if ($wasComplete -and $oldIP -eq $InstructorName) { return $false }
                return ($wasRide -or (Test-StringRide $Mission $InstructorName))
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
#  STRINGS AND COUNTERS (String Tracker)
#    - Each category (FT, FF, AA) is its own sequence, in mission order: FT5001, FT5002, ...
#      even when rides from another category were flown in between
#    - Non-counters are skipped, any number in a row, unless they flew the student two
#      rides in a row (strictly back to back): then they are a counter for that student
#    - One different counter between two of an IP's rides does not break that IP's string;
#      two different counters in a row end it
#    - Sims, solos and check rides (IsCheck) are left out entirely
# =================================================================================

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
                $History = $Sorties[($Sorties.Count - 1)..0]
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

# True when the event, flown with -InstructorName, is a ride that counts toward strings and
# counters: a flight with an IP in the cockpit, not a solo or a check ride
function Test-StringRide {
                param ($Mission, [string]$InstructorName)

                return (-not (ConvertTo-FaisBool $Mission.IsCheck) -and
                                [string]$Mission.Type -notlike '*Solo*' -and
                                (Get-MissionKind -Mission $Mission -InstructorName $InstructorName) -eq 'Flight' -and
                                -not [string]::IsNullOrWhiteSpace($InstructorName))
}

# A student's strings and counters from their syllabus and the Instructors list.
# Returns @(5-strings, 4-strings, 3-strings, 2-strings, 1-strings, counters). Strings longer than 5 count as 5-strings.
# Counters are the unique IPs who are not non-counters.
function Get-StringCounts {
                param ([array]$Syllabus, [array]$Instructors = @())

                # Completed rides with an IP in the cockpit, by category (FT5001 -> FT, 5001) and in mission order inside each
                $Flights = @(
                                $Syllabus | Where-Object {
                                                (Test-MissionComplete $_) -and (Test-StringRide $_ (Get-NameText $_.Instructor))
                                } | ForEach-Object {
                                                $null = ([string]$_.Mission).Trim() -match '^(\D*)(\d*)'
                                                [PSCustomObject]@{
                                                                Category = $Matches[1].ToUpper()
                                                                Number   = if ($Matches[2] -ne '') { [double]$Matches[2] } else { [double]::MaxValue }
                                                                ID       = Get-SyllabusSortKey $_
                                                                IP       = Get-BaseName (Get-NameText $_.Instructor)
                                                }
                                } | Sort-Object Category, Number, ID
                )

                # IPs who flew the student two rides in a row in a category (strictly back to back, not zippered)
                $BackToBack = @(
                                for ($i = 1; $i -lt $Flights.Count; $i++) {
                                                if ($Flights[$i].Category -eq $Flights[$i - 1].Category -and $Flights[$i].IP -eq $Flights[$i - 1].IP) { $Flights[$i].IP }
                                }
                )

                # Non-counters from the Instructors sheet (a blank IsCounter counts as a counter), except back-to-back ones
                $NonCounters = @(
                                $Instructors | Where-Object { -not (ConvertTo-FaisBool $_.IsCounter -Default $true) } |
                                                ForEach-Object { Get-BaseName (Get-NameText $_.Name) } |
                                                Where-Object { $_ -notin $BackToBack }
                )

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


# =================================================================================
#  DISCONTINUITY SCORE
#    How broken up a student's rides have been, from the String Tracker counts:
#    Disc1String per 1-string + DiscCounter per counter + Disc3PlusString per 3+ string
#    (SoftWeights; +1, +1 and -1 when blank). The Continue String rule adds it, so a
#    student with a lot of 1-strings and counters is the one kept with an IP.
#    Kept like SimDebt: read from MasterData's Students sheet (counted from the syllabus
#    when the cell is blank), moved by each ride that is flown or taken back, written back.
# =================================================================================

# One discontinuity weight from SoftWeights, or its default when the cell is blank
function Get-DiscontinuityWeight {
                param ($SoftWeights, [string]$Name, [double]$Default)

                $v = ([string](Get-OptionalProperty $SoftWeights $Name '')).Trim()
                $num = 0.0
                if ($v -ne '' -and [double]::TryParse($v, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$num)) { return $num }
                return $Default
}

# A student's discontinuity score, counted from their syllabus (Get-StringCounts)
function Get-DiscontinuityScore {
                param ([array]$Syllabus, [array]$Instructors = @(), $SoftWeights)

                # @(5+, 4, 3, 2, 1-strings, counters)
                $c = @(Get-StringCounts -Syllabus $Syllabus -Instructors $Instructors)
                $score = (Get-DiscontinuityWeight $SoftWeights 'Disc1String' 1) * $c[4] +
                                (Get-DiscontinuityWeight $SoftWeights 'DiscCounter' 1) * $c[5] +
                                (Get-DiscontinuityWeight $SoftWeights 'Disc3PlusString' -1) * ($c[0] + $c[1] + $c[2])
                return [math]::Round($score, 2)
}

# Moves the student's DiscontinuityScore by -Delta (a blank score counts as 0)
function Add-StudentDiscontinuity {
                param ($Student, [double]$Delta)

                $current = 0.0
                $v = Get-OptionalProperty $Student 'DiscontinuityScore' $null
                if ($null -ne $v -and [string]$v -ne '') { $current = [double]$v }
                $Student | Add-Member -MemberType NoteProperty -Name 'DiscontinuityScore' -Value ([math]::Round($current + $Delta, 2)) -Force
}

# Counts the score from the syllabus for each student whose DiscontinuityScore is blank (no
# DiscontinuityScore column in MasterData yet, or a cell cleared so FAIS counts it again).
# Returns the names of the students it counted.
function Initialize-StudentDiscontinuity {
                param ([array]$Students, [array]$Instructors = @(), $SoftWeights)

                $counted = New-Object System.Collections.Generic.List[string]
                foreach ($s in @($Students)) {
                                if ($null -eq $s) { continue }
                                $v = Get-OptionalProperty $s 'DiscontinuityScore' $null
                                if ($null -ne $v -and [string]$v -ne '') { continue }
                                $score = Get-DiscontinuityScore -Syllabus @($s.Syllabus) -Instructors $Instructors -SoftWeights $SoftWeights
                                $s | Add-Member -MemberType NoteProperty -Name 'DiscontinuityScore' -Value $score -Force
                                $counted.Add([string]$s.Name)
                }
                return $counted.ToArray()
}
