# SYLLABUS.PSM1 DIRECTORY:
#             Everything about a student's syllabus: which events they can fly next,
#             marking events complete, sim debt, and the timeline of what they fly
# Test-MissionComplete / Get-SyllabusSortKey / Find-SyllabusMission / Get-MissionKind
# Get-OptedEvents / Get-AllowedEventTypes / Get-ValidOptedEvents
# Get-SimDebtDelta / Get-MissionSimDebtDelta / Update-StudentSimDebtFromSyllabus
# Complete-StudentMission / Reset-PlannedCompletions
# Add-StudentReservation / Remove-StudentReservation / Get-ReservedMissionNames
# Get-StudentTimeline


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
function Complete-StudentMission {
                param (
                                [Parameter(Mandatory)] $Student,
                                [Parameter(Mandatory)] $Mission,
                                [string]$InstructorName,
                                $ETD,
                                $SoftWeights,
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

                $rev = [int](Get-OptionalProperty $Student 'SyllabusRevision' 0)
                $Student | Add-Member -MemberType NoteProperty -Name 'SyllabusRevision' -Value ($rev + 1) -Force
                return $entry
}

# Clears completions dated on/after a day. A previous run that scheduled these days
# wrote them as complete; they are plans, not flown events, when the days are re-built.
# Returns the students that changed (SimDebt recomputed for them).
function Reset-PlannedCompletions {
                param (
                                [array]$Students,
                                [datetime]$FromDate,
                                $SoftWeights
                )

                $changed = New-Object System.Collections.Generic.List[object]
                foreach ($s in @($Students)) {
                                $hit = $false
                                foreach ($m in @($s.Syllabus)) {
                                                if (-not (Test-MissionComplete $m)) { continue }
                                                $dt = Parse-ETD $m.ETD
                                                if ($null -ne $dt -and $dt.Date -ge $FromDate.Date) {
                                                                $m.IsComplete = $false
                                                                $m.Instructor = ''
                                                                $m.ETD = ''
                                                                $hit = $true
                                                }
                                }
                                if ($hit) {
                                                if ($null -ne $SoftWeights) { [void](Update-StudentSimDebtFromSyllabus -Student $s -SoftWeights $SoftWeights) }
                                                $rev = [int](Get-OptionalProperty $s 'SyllabusRevision' 0)
                                                $s | Add-Member -MemberType NoteProperty -Name 'SyllabusRevision' -Value ($rev + 1) -Force
                                                $changed.Add($s)
                                }
                }
                return $changed.ToArray()
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
