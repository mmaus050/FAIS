#FAIS
# cd C:\Users\1595161645A\Desktop\FAIS7.1


# =================================================================================
#  ____ ___  _   _ _____ ___ ____ _   _ ____      _  _____ ___ ___  _   _
# / ___/ _ \| \ | |  ___|_ _/ ___| | | |  _ \    / \|_   _|_ _/ _ \| \ | |
#| |  | | | |  \| | |_   | | |  _| | | | |_) |  / _ \ | |  | | | | |  \| |
#| |__| |_| | |\  |  _|  | | |_| | |_| |  _ <  / ___ \| |  | | |_| | |\  |
# \____\___/|_| \_|_|   |___\____|\___/|_| \_\/_/   \_\_| |___\___/|_| \_|
# =================================================================================


# ==========================================
#       CONFIGURATION
# ==========================================
$excelInputPath = Join-Path (Split-Path $PSScriptRoot -Parent) "MasterData.xlsx"
$desktopPath = [Environment]::GetFolderPath('Desktop')
$excelOutputPath = Join-Path -Path $desktopPath -ChildPath "OutputSchedule.xlsx"


# ==========================================
#      IMPORT MODULES AND FUNCTIONS
# ==========================================
# Makes sure you have excel
try {
    Import-Module ImportExcel -ErrorAction Stop
}
catch {
    Write-Host "Please install or copy the ImportExcel module, then try again." -ForegroundColor Yellow
    return
}

# Imports all the functions you need
Import-Module (Join-Path $PSScriptRoot "Modules\ConnectsSyllabusToStudent.psm1") -Force -DisableNameChecking
Import-Module (Join-Path $PSScriptRoot "Modules\GetOptedEvents1x2.psm1") -Force -DisableNameChecking
Import-Module (Join-Path $PSScriptRoot "Modules\ScheduleHandling1x3.psm1") -Force -DisableNameChecking
Import-Module (Join-Path $PSScriptRoot "Modules\HardAndSoftRules1x4.psm1") -Force -DisableNameChecking
Import-Module (Join-Path $PSScriptRoot "Modules\Diagnostics.psm1") -Force -DisableNameChecking
Import-Module (Join-Path $PSScriptRoot "Modules\FormationPairing1x8.psm1") -Force -DisableNameChecking
Import-Module (Join-Path $PSScriptRoot "Modules\PreRunMenu.psm1") -Force -DisableNameChecking
Import-Module (Join-Path $PSScriptRoot "Modules\OutputToTeams.psm1") -Force -DisableNameChecking

Import-Module (Join-Path $PSScriptRoot "Modules\DataDownload.psm1") -Force -DisableNameChecking
Import-Module (Join-Path $PSScriptRoot "Modules\DataManipulation.psm1") -Force -DisableNameChecking
Import-Module (Join-Path $PSScriptRoot "Modules\SchedulingEngine.psm1") -Force -DisableNameChecking

# Gets Program ready to run
# Close-AllOpenExcelWorkbooks


# ==========================================
#       INGEST DATA
# ==========================================
$Ingested = IngestData -excelInputPath $excelInputPath -desktopPath $desktopPath -excelOutputPath $excelOutputPath
$Schedule = $Ingested.Schedule
$Instructors = $Ingested.Instructors
$Students = $Ingested.Students
$Formations = $Ingested.Formations
$SoftWeights = $Ingested.SoftWeights
$StudentAbsences = $Ingested.StudentAbsences


<#
# ==========================================
#      SCHEDULING ENGINE
# ==========================================
$numForm = 0
$num = 0
$CurrentDay = $null


# ==========================================
#      PRE RUN MENU
# ==========================================
# This brings the previous day's schedule if you want to MND lines or Unsat them
# Brings in the Previous Schedule with all the complex objects attached
$PrevSchedule = DownloadPrevSchedule -Students $Students -Instructors $Instructors -desktopPath $desktopPath -excelOutputPath $excelOutputPath

# This needs to be reworked at least somewhat

#$MasterDataPath = Join-Path $PSScriptRoot 'MasterData.xlsx'
#Invoke-PreRunMenu -Schedule $PrevSchedule -Students $Students -SoftWeights $SoftWeights -MasterDataPath $excelPath


# ==================================================================================================
#  ____   ____ _   _ _____ ____  _   _ _     ___ _   _  ____   _     ___   ___  ____
# / ___| / ___| | | | ____|  _ \| | | | |   |_ _| \ | |/ ___| | |   / _ \ / _ \|  _ \
# \___ \| |   | |_| |  _| | | | | | | | |    | ||  \| | |  _  | |  | | | | | | | |_) |
#  ___) | |___|  _  | |___| |_| | |_| | |___ | || |\  | |_| | | |__| |_| | |_| |  __/
# |____/ \____|_| |_|_____|____/ \___/|_____|___|_| \_|\____| |_____\___/ \___/|_|
# ==================================================================================================
foreach ($Sortie in $Schedule){


#=======================================================
# SCHEDULE SET UP:
# Setup for the run, clearing arrays, etc
#
#=======================================================
                # Purely for Testing UI
                $num++
                Write-Host "Scheduling Sortie $num"

                # Empties the Candidates for each new sortie
                $SPCandidates = @()
                $NewNotes = @()


#=======================================================
# CONNECTS INSTRUCTOR DATA:
# Finds the Instructor based off the Instructor Data in Master Data
# $Sortie.Instructor is a string; $SortieInstructor is the object
#=======================================================
                # This converts the Instructor String into the complete Object with all their data
                if ($Sortie.Instructor -is [PSCustomObject]) {
                                $SortieInstructor = $Sortie.Instructor
                } else {
                                $SortieInstructor = MatchInstructors -Instructors $Instructors -TargetName $Sortie.Instructor
                                if ($null -eq $SortieInstructor) {
                                                $fallbackName = if ([string]::IsNullOrWhiteSpace($Sortie.Instructor)) { "NEED IP" } else { $Sortie.Instructor }
                                                $SortieInstructor = [PSCustomObject]@{ Name = $fallbackName; IsCounter = $true; IsExperienced = $false; IsCheck = $false; Callsign = ""; Num = "" }
                                }
                }


#=======================================================
# HARD CODED SORTIES:
# If there is input, track that, input into Output, Update Student Schedule, etc
#
#=======================================================
                if (-not [string]::IsNullOrWhiteSpace($Sortie.Student)) {
                                $numForm = Process-HardcodedSortie -Sortie $Sortie -SortieInstructor $SortieInstructor -Students                                         $Students -Schedule $Schedule -Instructors $Instructors -SoftWeights                                                                                                 $SoftWeights -Formations $Formations -FormIndex $numForm
                                continue
                #If there is no input, input the IP's Callsign and add in the Instructor Object
                } else {
                                $Sortie.Callsign = $SortieInstructor.Callsign
                                $Sortie.Num = $SortieInstructor.Num
                                $Sortie.Instructor = $SortieInstructor
                }


#=======================================================
# HARD RULES:
# For each student, runs the hard rules, doesn't run if the line is filled
#
#=======================================================
                foreach ($Student in $Students){
                                if ([string]::IsNullOrWhiteSpace($Student.Name)) { continue }
                                $Sortie.Mission = $null

                                # HardRules now obtains valid opted events itself after the early independent checks
                                if (HardRules -Student $Student -Sortie $Sortie -Schedule $Schedule -FormLookAhead $SoftWeights.FormLookAhead -Students $Students) {
                                                $SPCandidates += $Student
                                }
                }


#=======================================================
# SOFT RULES:
# Runs the soft rules, tries every event that the SP is opted for
# Maybe put own restrictions in about when flights can be flown
#=======================================================
                foreach ($Candidate in $SPCandidates) {
                                $BestScoreForStudent = $null
                                $BestEventForStudent = $null

                                foreach ($Event in $Candidate.OptedEvents) {
                                                $Sortie.Mission = $Event
                                                $Score = SoftRules -Student $Candidate -Config $SoftWeights -Sortie $Sortie

                                                if ($null -eq $BestScoreForStudent -or $Score -gt $BestScoreForStudent) {
                                                                $BestScoreForStudent = $Score
                                                                $BestEventForStudent = $Event
                                                }
                                }

                                $Candidate | Add-Member -MemberType NoteProperty -Name "Score" -Value $BestScoreForStudent -Force
                                $Candidate | Add-Member -MemberType NoteProperty -Name "BestEvent" -Value $BestEventForStudent -Force
                }

                $BestMatch = $SPCandidates | Sort-Object -Property Score -Descending | Select-Object -First 1
                $Sortie.Student = $BestMatch
                $Sortie.Mission = $BestMatch.BestEvent
                Write-Host "The Selected Student's Score is $($BestMatch.Score)"


#=======================================================
# UNABLE TO FILL:
# Puts the line up for sale if we don't have any valid
# students to put in the line
#=======================================================
                if ($SPCandidates.Count -eq 0) {
                                Write-Host "No available Student for $($Sortie.ETD) : $($Sortie.Type)"
                                $Sortie.Notes += "No Available SP"
                                if ($Sortie.SortieType -notlike "*FT*") {
                                                $Sortie.Notes += " | Potentially fill if we have another line"
                                }
                                $Sortie.Mission = " "
                                $Sortie.Student = "For Sale"
                                continue
                }


#=======================================================
# SIM DEBT/SYLLABUS/NOTES:
# Updates the Sim Debt and Syllabus for whatever event thats scheduled
# Also carrys the Notes over to be outputted
#
#
# This needs to change, apply the Sim Debt to the Student Object and then
# We are going to output that at the very end
#=======================================================
                # Mission is already set to the highest-scoring event that survived hard rules
                $Sortie.Student.SimDebt = SimDebtUpdate -Sortie $Sortie `
                                                      -SimStopMult $SoftWeights.SimStopMult `
                                                      -TransSimRatio $SoftWeights.TransSimRatio `
                                                      -FormSimRatio $SoftWeights.FormSimRatio
                if ($null -ne $BestMatch.PendingNote) {
                                $Sortie.Notes += " | $($BestMatch.PendingNote)"
                }


# =====================================================
# STUDENT SYLLABUS UPDATE:
# This updates the student's syllabus locally, just keeps it in the local memory
#
# =====================================================
                foreach ($Student in $Students) {
                                if ($Student.Name -eq $BestMatch.Name) {
                                                $Student.SimDebt = $Sortie.Student.SimDebt
                                                foreach ($m in $Student.Syllabus) {
                                                                if ($m.Mission -eq $SelectedMission.Mission) {
                                                                                $m.IsComplete = $true
                                                                                $m.Instructor = $SortieInstructor.Name
                                                                                $m.ETD = $Sortie.ETD
                                                                                break
                                                                }
                                                }
                                                break
                                }
                }


# =====================================================================================
#   ____ _     _____    _    _   _      _    _   _ ____     ___  _   _ _____ ____  _   _ _____
#  / ___| |   | ____|  / \  | \ | |    / \  | \ | |  _ \   / _ \| | | |_   _|  _ \| | | |_   _|
# | |   | |   |  _|   / _ \ |  \| |   / _ \ |  \| | | | | | | | | | | | | | | |_) | | | | | |
# | |___| |___| |___ / ___ \| |\  |  / ___ \| |\  | |_| | | |_| | |_| | | | |  __/| |_| | | |
#  \____|_____|_____/_/   \_\_| \_| /_/   \_\_| \_|____/   \___/ \___/  |_| |_|    \___/  |_|
# ====================================================================================

#>
