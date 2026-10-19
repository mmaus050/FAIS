# SCHEDULINGENGINE.PSM1 DIRECTORY:

#             Contains all functions that run the logic for the scheduling

# HardRules

# SoftRules

# PairWingman

#

#

#

#

#

#

#

#

#

#

#

#

#

#

 

 

 

# =================================================================================

#  ____                 _   _               _               _        ____             _   _     

# |  _ \ __ _ ___ ___  | | | | __ _ _ __ __| | ___ ___   __| | ___  / ___|  ___  _ __| |_(_) ___

# | |_) / _` / __/ __| | |_| |/ _` | '__/ _` |/ __/ _ \ / _` |/ _ \ \___ \ / _ \| '__| __| |/ _ \

# |  __/ (_| \__ \__ \ |  _  | (_| | | | (_| | (_| (_) | (_| |  __/  ___) | (_) | |  | |_| |  __/

# |_|   \__,_|___/___/ |_| |_|\__,_|_|  \__,_|\___\___/ \__,_|\___| |____/ \___/|_|   \__|_|\___|

# =================================================================================

function PassHardcodeSortie {

                param (

                                [Parameter(Mandatory=$true)] $Sortie,

                                [Parameter(Mandatory=$true)] $SortieInstructor,

                                [Parameter(Mandatory=$true)] [array]$Students,

                                [Parameter(Mandatory=$true)] [array]$Schedule,

                                [Parameter(Mandatory=$true)] [array]$Instructors,

                                $SoftWeights,

                                $Formations,

                                [int]$FormIndex = 0

                )

                # 1. Resolve Student Name and locate their object in $Students

                $rawStudentName = if ($Sortie.Student -is [PSCustomObject]) { $Sortie.Student.Name } else { [string]$Sortie.Student }

                $normalizedRawName = Get-NormalizedLastName -Name $rawStudentName

                $matchedStudent = $Students | Where-Object {

                                (Get-NormalizedLastName -Name $_.Name).ToLower() -eq $normalizedRawName.ToLower()

                } | Select-Object -First 1

                if ($null -ne $matchedStudent) {

                                # Link the full student object

                                $Sortie.Student = $matchedStudent

                                # 2. Resolve Mission (from row or auto-find top valid opted event if blank)

                                $rawMissionName = if ($Sortie.Mission -is [PSCustomObject]) { $Sortie.Mission.Mission } else { [string]$Sortie.Mission }

                                $matchedMission = $null

                                if (-not [string]::IsNullOrWhiteSpace($rawMissionName)) {

                                                $matchedMission = $matchedStudent.Syllabus | Where-Object { $_.Mission.Trim() -eq $rawMissionName.Trim() } | Select-Object -First 1

                                } else {

                                                $validEvents = Get-ValidOptedEvents -Student $matchedStudent -Sortie $Sortie

                                                if ($null -ne $validEvents) {

                                                                $matchedMission = $validEvents | Sort-Object -Property SyllabusID | Select-Object -First 1

                                                }

                                }

                                # 3. Update the student's syllabus event in memory

                                if ($null -ne $matchedMission) {

                                                $Sortie.Mission = $matchedMission

                                                foreach ($m in $matchedStudent.Syllabus) {

                                                                if ($m.Mission -eq $matchedMission.Mission) {

                                                                                $m.IsComplete = $true

                                                                                $m.Instructor = if ($SortieInstructor -is [PSCustomObject]) { $SortieInstructor.Name } else { $SortieInstructor }

                                                                                $m.ETD = $Sortie.ETD

                                                                                break

                                                                }

                                                }

                                                Write-Host " [BYPASS] Scheduled '$($matchedStudent.Name)' on '$($matchedMission.Mission)' (ETD: $($Sortie.ETD)) - Syllabus Updated." -ForegroundColor Cyan

                                } else {

                                                Write-Warning " [BYPASS] Student '$($matchedStudent.Name)' has no valid mission matching '$rawMissionName' on ETD $($Sortie.ETD)."

                                }

                                # 4. Ensure Instructor and Callsigns are populated

                                if (-not ($Sortie.Instructor -is [PSCustomObject])) {

                                                $Sortie.Instructor = $SortieInstructor

                                }

                                if ([string]::IsNullOrWhiteSpace($Sortie.Callsign)) {

                                                $Sortie.Callsign = $SortieInstructor.Callsign

                                                $Sortie.Num = $SortieInstructor.Num

                                }

 

 

                                # 5. Handle Formation Pairing only for true leads.

                                #    Skip if this line is already a wingman (Notes already contain "Wingman to").

                                $alreadyWingman = $false

                                if (-not [string]::IsNullOrWhiteSpace($Sortie.Notes) -and $Sortie.Notes -like "*Wingman to*") {

                                                $alreadyWingman = $true

                                }

 

                                if ($null -ne $matchedMission -and $matchedMission.Type -like "*Form*" -and -not $alreadyWingman) {

                                                $FormIndex = Pair-FormationWingman -PrimarySortie $Sortie `

                                                                                   -PrimaryStudent $matchedStudent `

                                                                                   -PrimaryMission $matchedMission `

                                                                                   -Schedule $Schedule `

                                                                                   -Students $Students `

                                                                                   -Instructors $Instructors `

                                                                                   -SoftWeights $SoftWeights `

                                                                                   -FormLookAhead $SoftWeights.FormLookAhead `

                                                                                   -LevelDifference $SoftWeights.LevelDifference `

                                                                                   -FormIndex $FormIndex `

                                                                                   -Formations $Formations

                                }

 

 

                                # 6. Handle O&B Return Leg if applicable

                                if ($Sortie.SortieType -eq "O&B" -and $null -ne $matchedMission) {

                                                Schedule-ReturnLeg -Sortie $Sortie -BestMatch $matchedStudent -Schedule $Schedule -Students $Students

                                }

                } else {

                                Write-Warning " [BYPASS] Student '$rawStudentName' on ETD $($Sortie.ETD) was not found in MasterData."

                }

                # Return updated FormIndex so formation callsigns remain in sequence

                return $FormIndex

}

 

 

 

 

 

 

 

 

# =============================================================================================
#  _   _               _   ____        _          
# | | | | __ _ _ __ __| | |  _ \ _   _| | ___  ___
# | |_| |/ _` | '__/ _` | | |_) | | | | |/ _ \/ __|
# |  _  | (_| | | | (_| | |  _ <| |_| | |  __/\__ \
# |_| |_|\__,_|_|  \__,_| |_| \_\\__,_|_|\___||___/
# =============================================================================================

function HardRules {

                param (

                                $Student,

                                $Sortie,

                                $Schedule,

                                $FormLookAhead = 22,

                                $Students,

                                [switch]$IsWingCandidate

                )

 

# Set Up: This calculates

# This gets all events that the SP has completed for crew rest, solo continuity etc

                $NextStart = Parse-ETD $Sortie.ETD

                $completedEvents = @(

                                $Student.Syllabus | Where-Object {

                                                ($_.IsComplete -eq $true -or $_.IsComplete -eq 'True') -and

                                                -not [string]::IsNullOrWhiteSpace($_.ETD)

                                } | Sort-Object -Property { Parse-ETD $_.ETD } -Descending

                )

                $todaysEvents = @(

                                $completedEvents | Where-Object {

                                                $evtDate = Parse-ETD $_.ETD

                                                $null -ne $evtDate -and $evtDate.Date -eq $NextStart.Date

                                }

                )

 

 

############################################################

# STUDENT UNAVAILABILITY / ABSENCES

# This blocks a student from being scheduled if the proposed sortie

# overlaps with any time they are marked unavailable in the StudentAbsences sheet.

# NEW BUFFER LOGIC:

#   - The sortie is considered to "really begin" 1.5 hours BEFORE the ETD

#   - The sortie is considered to "end"       2.5 hours AFTER  the ETD

#

# Example: A 1000 sortie will now be treated as occupying ~0830 – 1230

#          when checking against student absences.

############################################################

                if ($null -ne $Student.Absences -and $Student.Absences.Count -gt 0) {

                                $BufferBeforeHours = -1.5

                                $BufferAfterHours  =  2.5

                                $SortieEffectiveStart = $NextStart.AddHours($BufferBeforeHours)

                                $SortieEffectiveEnd   = $NextStart.AddHours($BufferAfterHours)

 

                                foreach ($absence in $Student.Absences) {

                                                if ($SortieEffectiveStart -lt $absence.EndTime -and $SortieEffectiveEnd -gt $absence.StartTime) {

                                                                $NewNotes += "Student unavailable ($($absence.Reason))"

                                                                return $false

                                                }

                                }

                }

 

 

#################################################

# Prevent O&B Return Overwrite

# - TESTED

#################################################

                if ($Sortie.SortieType -like "*Return*") {

                                return $false

                }

 

 

###############################################

# BANS CAPsters FROM O&Bs

# - TESTED

###############################################

                if ($Student.IsCap -and $Sortie.SortieType -like "*O&B*") {

                                return $false

                }

 

 

##############################################

# CAPsters single turn ONLY (Flights)

# - TESTED

##############################################

                if ($Student.IsCAP) {

                                if ($Sortie.SortieType -notlike "*FT*") {

                                                $flightsToday = 0

                                                foreach ($event in $todaysEvents) {

                                                                if ($event.Instructor -notlike "*SIM IP*") {

                                                                                $flightsToday += 1

                                                                }

                                                }

                                                if ($flightsToday -ge 1) { return $false }

                                }

                }

 

 

##############################################

# HOLD Status

# - TESTED

##############################################

                if ($Student.OnHold) { return $false }

 

 

##############################################

# TRIPLE TURN PREVENTION

# - TESTED

##############################################

                if ($todaysEvents.Count -ge 2) {

                                return $false

                }

                elseif ($Sortie.SortieType -eq "O&B" -and $todaysEvents.Count -ge 1) {

                                return $false

                }

 

 

##############################################

# O&B RETURN LEG HARD RULES LOOKAHEAD

#

##############################################

                $ReturnStart = $null

                $ReturnMission = $null

                if ($Sortie.SortieType -eq "O&B") {

                                $TargetName = if ($Sortie.Instructor -is [PSCustomObject]) { $Sortie.Instructor.Name } else { $Sortie.Instructor }

                                $ReturnSortie = $Schedule | Where-Object {

                                                $_.SortieType -like "*Return*" -and

                                                [string]::IsNullOrWhiteSpace($_.Student) -and

                                                ([datetime]$_.ETD).Date -eq $NextStart.Date

                                } | Where-Object {

                                                $rowInst = if ($_.Instructor -is [PSCustomObject]) { $_.Instructor.Name } else { $_.Instructor }

                                                [string]$rowInst.Trim() -eq [string]$TargetName.Trim()

                                } | Select-Object -First 1

 

                                if ($null -ne $ReturnSortie) {

                                                $ReturnStart = Parse-ETD $ReturnSortie.ETD

                                }

 

                                $SimulatedMission = $Sortie.Mission

                                $SimulatedMission.IsComplete = $true

                                $NextEvents = Get-OptedEvents -SyllabusData $Student.Syllabus

                                if ($null -ne $NextEvents) {

                                                foreach ($e in $NextEvents) {

                                                                if ($e.Type -eq "Single" -or $e.Type -eq "Formation" -or $e.Type -eq "Both") {

                                                                                $ReturnMission = $e

                                                                                break

                                                                }

                                                }

                                }

                                $SimulatedMission.IsComplete = $false

 

                                # Ensure return leg isn't a checkride

                                if ($null -ne $ReturnMission -and $ReturnMission.IsCheck) {

                                                return $false

                                }

                }

 

 

######################################################

# TURN TIME, DUTY DAY, CREW REST

# Pretty Thoroughly Tested, Complex Rule so lots of room for error

######################################################

                $CrewRestCheck = $true

                $TurnTimeCheck = $true

                $DutyDayCheck  = $true

 

                if ($Sortie.SortieType -notlike "*Return*") {

                                foreach ($evt in $todaysEvents) {

                                                $evtDate = Parse-ETD $evt.ETD

                                                if ($null -ne $evtDate) {

                                                                $TurnTime = [Math]::Abs(($NextStart - $evtDate).TotalHours)

                                                               

                                                                if ($TurnTime -lt 3.25) {

                                                                                $TurnTimeCheck = $false

                                                                                break

                                                                }

                                                                elseif ($TurnTime -ge 3.25 -and $TurnTime -lt 3.5) {

                                                                                $TurnTimeCheck = $true

                                                                                $NewNotes += "Less than 3.5 hr TT, Flt/CC Approval Required"

                                                                }

                                                                elseif ($TurnTime -ge 6 -and $TurnTime -le 9) {

                                                                                $TurnTimeCheck = $true

                                                                                $NewNotes += "If this is a Night Flight SP will not have Duty Day"

                                                                }

                                                                elseif ($TurnTime -gt 9) {

                                                                                $DutyDayCheck = $false

                                                                                break

                                                                }

 

                                                                if ($null -ne $ReturnStart) {

                                                                                $ReturnTurnTime = [Math]::Abs(($ReturnStart - $evtDate).TotalHours)

                                                                                if ($ReturnTurnTime -lt 3.25) {

                                                                                                $TurnTimeCheck = $false

                                                                                                break

                                                                                }

                                                                                elseif ($ReturnTurnTime -ge 3.25 -and $ReturnTurnTime -lt 3.5) {

                                                                                                $TurnTimeCheck = $true

                                                                                                $NewNotes += "Return Leg: Less than 3.5 hr TT, Flt/CC Approval Required"

                                                                                }

                                                                                elseif ($ReturnTurnTime -ge 6 -and $ReturnTurnTime -le 9) {

                                                                                                $TurnTimeCheck = $true

                                                                                                $NewNotes += "Return Leg: Night Flight Warning"

                                                                                }

                                                                                elseif ($ReturnTurnTime -gt 9) {

                                                                                                $DutyDayCheck = $false

                                                                                                break

                                                                                }

                                                                }

                                                }

                                }

                }

 

                $PastEvents = @($completedEvents | Where-Object { (Parse-ETD $_.ETD).Date -lt $NextStart.Date })

                if ($PastEvents.Count -gt 0) {

                                $LastStart = Parse-ETD $PastEvents[0].ETD

                                if ($null -ne $LastStart -and $LastStart.Date -eq $NextStart.Date.AddDays(-1)) {

                                                $CrewRest = ($NextStart - $LastStart).TotalHours

                                                if ($CrewRest -lt 15) {

                                                                $CrewRestCheck = $false

                                                }

                                }

                }

 

                if (-not $CrewRestCheck) { return $false }

                if (-not $TurnTimeCheck) { return $false }

                if (-not $DutyDayCheck) { return $false }

 

 

###########################################################

# CAP/INTERNATIONAL AND EXP IPs

# - TESTED

###########################################################

                if ($Student.IsCAP -or $Student.IsInternational) { 

                                if ($Sortie.SortieType -like "*Day Local*" -or $Sortie.SortieType -like "*O&B*") {

                                                if (-not $Sortie.Instructor.IsExperienced) { return $false }

                                }

                }

 

 

##########################################################

# TRAINING FOLLOWING UNSAT EVENT

#

##########################################################

                if ($Student.UnsatSim -and ($Sortie.SortieType -like "*Day Local*" -or $Sortie.SortieType -like "*O&B*")) { return $false }

                if ($Student.UnsatFlight -and ($Sortie.SortieType -like "*IFT*" -or $Sortie.SortieType -like "*OFT*")) { return $false }

 

 

##########################################################

##########################################################

                # =========================================================

                # COLLECT EVENTS TO TEST

                # =========================================================

                $eventsToTest = @()

 

                if ($null -eq $Sortie.Mission) {

                                $ValidEvents = Get-ValidOptedEvents -Student $Student -Sortie $Sortie

                                if ($null -eq $ValidEvents) { return $false }

                                $eventsToTest = @($ValidEvents)

                }

                else {

                                # Caller already supplied a specific mission (e.g. wing candidate path)

                                $eventsToTest = @($Sortie.Mission)

                }

 

                $PassedEvents = @()

 

##########################################################

##########################################################

 

foreach ($Event in $eventsToTest) {

                                $Sortie.Mission = $Event

                                $NewNotes = @()          # reset notes for this particular event

 

 

#################################################

# TEMPORARY Block Formations from O&B

# - TESTED

#################################################

                                if ($Sortie.Mission.Type -like "*Form*" -and $Sortie.SortieType -like "*O&B*") {

                                                return $false

                                }

 

 

##################################################

# Block checkrides on O&Bs entirely

# - TESTED

##################################################

                                if ($Sortie.SortieType -like "*O&B*") {

                                                if ($Sortie.Mission.IsCheck) {

                                                                return $false

                                                }

                                }

 

 

###############################################

# CHECKRIDE EVALUATOR INDEPENDENCE

#

###############################################

                if ($Sortie.Mission.IsCheck) {

                                $CurrentIP_Base = $Sortie.Instructor.Name.Split("[")[0].Trim()

                                $PastIPs = @(

                                                $completedEvents | Where-Object {

                                                                -not [string]::IsNullOrWhiteSpace($_.Instructor) -and $_.Instructor -notlike "*SIM IP*"

                                                } | ForEach-Object {

                                                                $ipName = if ($_.Instructor -is [PSCustomObject]) { $_.Instructor.Name } else { [string]$_.Instructor }

                                                                $ipName.Split("[")[0].Trim()

                                                }

                                )

                                if ($PastIPs -contains $CurrentIP_Base) {

                                                return $false

                                }

                               

                                # Checkrides with only Check IPs

                                if (-not $Sortie.Instructor.IsCheck) {

                                                return $false

                                }

                }

 

 

################################################

                # FORMATION BLOCK

                # Only applied when the student is being considered as a lead.

                # A wing candidate already has a lead that claimed this line, so the

                # open-line count is skipped for wing evaluations.

################################################

                # This shouldn't matter because the sortie should be hardcoded so this gets skipped

 

 

###########################################################

# SOLO FLIGHT CONTINUITY

# - TESTED

###########################################################

                if ($Sortie.SortieType -notlike "*OFT*" -and $Sortie.SortieType -notlike "*IFT*" -and -not [string]::IsNullOrWhiteSpace($Sortie.Instructor)) {

                               

                                $completionStatus = @{}

                                foreach ($r in $Student.Syllabus) { $completionStatus[$r.Mission] = ([bool]$r.IsComplete -or $r.IsComplete -match 'True') }

                               

                                $IsNextFlightSolo = $false

                                foreach ($r in ($Student.Syllabus | Sort-Object SyllabusID)) {

                                                if (-not $completionStatus[$r.Mission]) {

                                                                $prereqsMet = $true

                                                                if ($r.Prereqs -ne '0' -and -not [string]::IsNullOrWhiteSpace($r.Prereqs)) {

                                                                                foreach ($p in ($r.Prereqs -split ',')) {

                                                                                                if (-not $completionStatus[$p.Trim()]) { $prereqsMet = $false; break }

                                                                                }

                                                                }

                                                               

                                                                if ($prereqsMet -and ($r.Type -eq "Single" -or $r.Type -eq "SingleSolo" -or $r.Type -eq "Formation" -or $r.Type -eq "FormSolo" -or $r.Type -eq "Both")) {

                                                                                if ($r.Type -like "*Solo*") { $IsNextFlightSolo = $true }

                                                                                break

                                                                }

                                                }

                                }

 

                                if ($IsNextFlightSolo) {

                                                $NewNotes += "Solo Flight"

                                               

                                                $CurrentIP_Base = $Sortie.Instructor.Name.Split("[")[0].Trim()

                                               

                                                if ($CurrentIP_Base -notlike "*SIM IP*") {

                                                                $Last5Flights = @(

                                                                                $completedEvents | Where-Object {

                                                                                                # Safely get the instructor name whether it's an object or a string

                                                                                                $instName = if ($_.Instructor -is [PSCustomObject]) { $_.Instructor.Name } else { [string]$_.Instructor }

                                                                                                $instName -notlike "*SIM IP*" -and -not [string]::IsNullOrWhiteSpace($instName)

                                                                                } | Select-Object -First 5

                                                                )

                                                               

                                                                $IPFound = $false

                                                                foreach ($flight in $Last5Flights) {

                                                                                # Safely extract the string before attempting to Split

                                                                                $flightInstName = if ($flight.Instructor -is [PSCustomObject]) { $flight.Instructor.Name } else { [string]$flight.Instructor }

                                                                                $flightIP_Base = $flightInstName.Split("[")[0].Trim()

                                                                               

                                                                                if ($flightIP_Base -eq $CurrentIP_Base) {

                                                                                                $IPFound = $true

                                                                                                break

                                                                                }

                                                                }

                                                               

                                                                if (-not $IPFound) {

                                                                                $NewNotes += "Need IP in Continuity to Brief"

                                                                }

                                                }

                                }

                }

 

 

 

 

 

                if ($NewNotes.Count -gt 0) {

                                $joinedNotes = $NewNotes -join " | "

                                $Student | Add-Member -MemberType NoteProperty -Name "PendingNote" -Value $joinedNotes -Force

                } else {

                                $Student | Add-Member -MemberType NoteProperty -Name "PendingNote" -Value $null -Force

                }

               

                return $true

}

 

 

$PassedEvents += $Event

}

 

 

                # =========================================================

                # FINAL DECISION

                # =========================================================

                if ($PassedEvents.Count -eq 0) { return $false }

 

                # Only the events that passed hard rules go forward to SoftRules

                $Student | Add-Member -MemberType NoteProperty -Name "OptedEvents" -Value $PassedEvents -Force

 

                # Give the rest of the scheduler a default mission (highest priority survivor)

                $Sortie.Mission = $PassedEvents | Sort-Object -Property SyllabusID | Select-Object -First 1

 

                # Optional: apply notes from the highest-priority surviving event

                if ($NewNotes.Count -gt 0) {

                                $joinedNotes = $NewNotes -join " | "

                                $Student | Add-Member -MemberType NoteProperty -Name "PendingNote" -Value $joinedNotes -Force

                }

                else {

                                $Student | Add-Member -MemberType NoteProperty -Name "PendingNote" -Value $null -Force

                }
