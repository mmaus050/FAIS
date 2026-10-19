# DATAMANIPULATION.PSM1 DIRECTORY:

#             Contains all functions that change, move, format data during the program run

# MatchInstructors

# ParseETD

# Convert-SortieObjectsToText

# SyllabusSimDebtUpdate

#

#

#

#

#

#

#

 

 

 

# =================================================================================

#  __  __       _       _       ___           _                   _                

# |  \/  | __ _| |_ ___| |__   |_ _|_ __  ___| |_ _ __ _   _  ___| |_ ___  _ __ ___

# | |\/| |/ _` | __/ __| '_ \   | || '_ \/ __| __| '__| | | |/ __| __/ _ \| '__/ __|

# | |  | | (_| | || (__| | | |  | || | | \__ \ |_| |  | |_| | (__| || (_) | |  \__ \

# |_|  |_|\__,_|\__\___|_| |_| |___|_| |_|___/\__|_|   \__,_|\___|\__\___/|_|  |___/

# ==================================================================================

function MatchInstructors {

                param (

                                [array]$Instructors,

                                [string]$TargetName

                )

 

                if ([string]::IsNullOrWhiteSpace($TargetName)) {

                                return $null

                }

 

                try {

                                $BaseTarget = $TargetName -replace '\s*\[.*?\]', ''

                                $BaseTarget = $BaseTarget.Trim()

 

                                $Match = $Instructors | Where-Object {

                                                $BaseSource = $_.Name -replace '\s*\[.*?\]', ''

                                                $BaseSource = $BaseSource.Trim()

                                                $BaseSource -match [regex]::Escape($BaseTarget)

                                } | Select-Object -First 1

 

                                return $Match

                } catch {

                                Write-Warning "Error matching instructor: $($_.Exception.Message)"

                                return $null

                }

}

 

 

 

 

 

function SyllabusSimDebtUpdate {

    param (

        $Sortie,

        $Students,

        $SimStopMult,

        $TransSimRatio,

        $FormSimRatio

    )

 

 

    $Sortie.Student.SimDebt = SimDebtUpdate -Sortie $Sortie -SimStopMult $SimStopMult -TransSimRatio $TransSimRatio -FormSimRatio $FormSimRatio

 

 

 

 

    foreach ($Student in $Students){

        if ($Student.Name -eq $Sortie.Student.Name){

            foreach ($m in $Student.Syllabus) {

                                                                if ($m.Mission -eq $Sortie.Mission) {

                                                                                $m.IsComplete = $true

                                                                                $m.Instructor = $Sortie.Instructor.Name

                                                                                $m.ETD = $Sortie.ETD

                                                                                break

                                                                }

                                                }

                                                break

 

        }

    }

 

}

 

function SyllabusSimDebtUpdate {

                param (

                                $Sortie,

                                $SimStopMult,

                                $TransSimRatio,

                                $FormSimRatio

                )

 

                # Update SimDebt on the student attached to this sortie.

                # That student object is the same one sitting in the main $Students array.

                $Sortie.Student.SimDebt = SimDebtUpdate -Sortie $Sortie -SimStopMult $SimStopMult -TransSimRatio $TransSimRatio -FormSimRatio $FormSimRatio

 

                # Mark this sortie s mission complete on that student s syllabus

                foreach ($m in $Sortie.Student.Syllabus) {

                                if ($m.Mission -eq $Sortie.Mission.Mission) {

                                                $m.IsComplete = $true

                                                $m.Instructor = $Sortie.Instructor.Name

                                                $m.ETD = $Sortie.ETD

                                                break

                                }

                }

}

 

 

 

 

 

# ===============================================================================================

#  ____                                ____       _              _       _          _   ____             _   _     

# |  _ \ _ __ ___   ___ ___  ___ ___  / ___|  ___| |__   ___  __| |_   _| | ___  __| | / ___|  ___  _ __| |_(_) ___

# | |_) | '__/ _ \ / __/ _ \/ __/ __| \___ \ / __| '_ \ / _ \/ _` | | | | |/ _ \/ _` | \___ \ / _ \| '__| __| |/ _ \

# |  __/| | | (_) | (_|  __/\__ \__ \  ___) | (__| | | |  __/ (_| | |_| | |  __/ (_| |  ___) | (_) | |  | |_| |  __/

# |_|   |_|  \___/ \___\___||___/___/ |____/ \___|_| |_|\___|\__,_|\__,_|_|\___|\__,_| |____/ \___/|_|   \__|_|\___|

# ================================================================================================

 

function HardcodedSortie {

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