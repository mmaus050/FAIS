# PRERUNMENU.PSM1 DIRECTORY:
#             The pop-ups that run before the puck board: log hooked (unsat) and
#             incomplete events from the previous schedule
# Initialize-FaisWinForms
# Show-PreRunMenu
# Show-EventSelectionMenu
# Process-HookedEvents
# Process-IncompleteEvents
# Invoke-PreRunMenu


# Loads WinForms only when a menu is shown, so the module also imports where
# WinForms is unavailable (tests)
function Initialize-FaisWinForms {
                Add-Type -AssemblyName System.Windows.Forms
                Add-Type -AssemblyName System.Drawing
                [System.Windows.Forms.Application]::EnableVisualStyles()
}


# ========================================================================================
#  ____                 ____                __  __
# |  _ \ _ __ ___      |  _ \ _   _ _ __   |  \/  | ___ _ __  _   _
# | |_) | '__/ _ \_____| |_) | | | | '_ \  | |\/| |/ _ \ '_ \| | | |
# |  __/| | |  __/_____|  _ <| |_| | | | | | |  | |  __/ | | | |_| |
# |_|   |_|  \___|     |_| \_\\__,_|_| |_| |_|  |_|\___|_| |_|\__,_|
# ========================================================================================
# Returns "Hook", "Incomplete", "Run" or "Exit"
function Show-PreRunMenu {
                Initialize-FaisWinForms

                $Form = New-Object System.Windows.Forms.Form
                $Form.Text = "Pre-Run Operations Menu"
                $Form.Size = New-Object System.Drawing.Size(400, 300)
                $Form.StartPosition = "CenterScreen"
                $Form.Font = New-Object System.Drawing.Font("Segoe UI", 11)

                $Label = New-Object System.Windows.Forms.Label
                $Label.Text = "Select an action before running the scheduler:"
                $Label.Location = New-Object System.Drawing.Point(20, 20)
                $Label.AutoSize = $true
                $Form.Controls.Add($Label)

                $btnHook = New-Object System.Windows.Forms.Button
                $btnHook.Text = "Log Hooked Event (Unsat)"
                $btnHook.Location = New-Object System.Drawing.Point(20, 60)
                $btnHook.Size = New-Object System.Drawing.Size(340, 40)
                $btnHook.Add_Click({ $Form.Tag = "Hook"; $Form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $Form.Close() })
                $Form.Controls.Add($btnHook)

                $btnIncomplete = New-Object System.Windows.Forms.Button
                $btnIncomplete.Text = "Log Incomplete Event"
                $btnIncomplete.Location = New-Object System.Drawing.Point(20, 110)
                $btnIncomplete.Size = New-Object System.Drawing.Size(340, 40)
                $btnIncomplete.Add_Click({ $Form.Tag = "Incomplete"; $Form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $Form.Close() })
                $Form.Controls.Add($btnIncomplete)

                $btnRun = New-Object System.Windows.Forms.Button
                $btnRun.Text = "Open Puck Board ->"
                $btnRun.Location = New-Object System.Drawing.Point(20, 180)
                $btnRun.Size = New-Object System.Drawing.Size(340, 40)
                $btnRun.BackColor = [System.Drawing.Color]::LightGreen
                $btnRun.Add_Click({ $Form.Tag = "Run"; $Form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $Form.Close() })
                $Form.Controls.Add($btnRun)

                $result = $Form.ShowDialog()
                if ($result -eq [System.Windows.Forms.DialogResult]::OK) { return $Form.Tag } else { return "Exit" }
}


# ========================================================================================
#  _____                 _     ____       _           _   _
# | ____|_   _____ _ __ | |_  / ___|  ___| | ___  ___| |_(_) ___  _ __
# |  _| \ \ / / _ \ '_ \| __| \___ \ / _ \ |/ _ \/ __| __| |/ _ \| '_ \
# | |___ \ V /  __/ | | | |_   ___) |  __/ |  __/ (__| |_| | (_) | | | |
# |_____| \_/ \___|_| |_|\__| |____/ \___|_|\___|\___|\__|_|\___/|_| |_|
# ========================================================================================
# Displays the list of scheduled events to select from. Returns the checked items
# ({Display; Sortie}) or $null on cancel.
function Show-EventSelectionMenu {
                param (
                                $Schedule,
                                $Title
                )

                Initialize-FaisWinForms

                $Form = New-Object System.Windows.Forms.Form
                $Form.Text = $Title
                $Form.Size = New-Object System.Drawing.Size(850, 500)
                $Form.StartPosition = "CenterScreen"
                $Form.Font = New-Object System.Drawing.Font("Segoe UI", 11)

                $HeaderLabel = New-Object System.Windows.Forms.Label
                $HeaderLabel.Text = "Select the scheduled flights/sims:"
                $HeaderLabel.Location = New-Object System.Drawing.Point(20, 20)
                $HeaderLabel.AutoSize = $true
                $Form.Controls.Add($HeaderLabel)

                $CheckedListBox = New-Object System.Windows.Forms.CheckedListBox
                $CheckedListBox.Location = New-Object System.Drawing.Point(20, 60)
                $CheckedListBox.Size = New-Object System.Drawing.Size(790, 320)
                $CheckedListBox.CheckOnClick = $true
                $CheckedListBox.Font = New-Object System.Drawing.Font("Consolas", 10)

                # Only lines that had a real student on them
                $ScheduledSorties = @(
                                $Schedule |
                                                Where-Object { -not (Test-IsSentinelStudent $_.Student) } |
                                                Sort-Object -Property { $d = Parse-ETD $_.ETD; if ($null -eq $d) { [datetime]::MinValue } else { $d } } -Descending
                )

                foreach ($Sortie in $ScheduledSorties) {
                                $studentName = Get-NameText $Sortie.Student
                                $missionName = Get-NameText $Sortie.Mission
                                if ($missionName -eq '') { $missionName = 'NONE' }
                                $ipName = Get-NameText $Sortie.Instructor
                                if ($ipName -eq '') { $ipName = 'NO IP' }

                                $displayString = "{0,-20} | {1,-30} | SP: {2,-15} | Mis: {3,-10} | {4}" -f `
                                                $Sortie.ETD, $ipName, $studentName, $missionName, $Sortie.SortieType

                                [void]$CheckedListBox.Items.Add([PSCustomObject]@{
                                                Display = $displayString
                                                Sortie  = $Sortie
                                })
                }

                $CheckedListBox.DisplayMember = "Display"
                $Form.Controls.Add($CheckedListBox)

                $SubmitBtn = New-Object System.Windows.Forms.Button
                $SubmitBtn.Location = New-Object System.Drawing.Point(20, 400)
                $SubmitBtn.Size = New-Object System.Drawing.Size(200, 40)
                $SubmitBtn.Text = "Confirm Selection"
                $SubmitBtn.BackColor = [System.Drawing.Color]::LightCoral
                $SubmitBtn.Add_Click({
                                if ($CheckedListBox.CheckedItems.Count -gt 0) {
                                                $Form.Tag = @($CheckedListBox.CheckedItems)
                                                $Form.DialogResult = [System.Windows.Forms.DialogResult]::OK
                                                $Form.Close()
                                }
                                else {
                                                [void][System.Windows.Forms.MessageBox]::Show(
                                                                "Please select at least one event.",
                                                                "None Selected",
                                                                [System.Windows.Forms.MessageBoxButtons]::OK,
                                                                [System.Windows.Forms.MessageBoxIcon]::Warning
                                                )
                                }
                })
                $Form.Controls.Add($SubmitBtn)

                $ExitBtn = New-Object System.Windows.Forms.Button
                $ExitBtn.Location = New-Object System.Drawing.Point(610, 400)
                $ExitBtn.Size = New-Object System.Drawing.Size(200, 40)
                $ExitBtn.Text = "Cancel"
                $ExitBtn.BackColor = [System.Drawing.Color]::LightGray
                $ExitBtn.Add_Click({
                                $Form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                                $Form.Close()
                })
                $Form.Controls.Add($ExitBtn)

                $result = $Form.ShowDialog()
                if ($result -eq [System.Windows.Forms.DialogResult]::OK) { return $Form.Tag }
                return $null
}


# ========================================================================================
# HOOKED (UNSAT): sim unsat blocks later flights; flight unsat blocks later sims
# ========================================================================================
function Process-HookedEvents {
                param (
                                [Parameter(Mandatory)] $SelectedItems,
                                [Parameter(Mandatory)] $Students
                )

                $updatedCount = 0
                foreach ($Item in $SelectedItems) {
                                $Sortie = $Item.Sortie
                                if ($null -eq $Sortie -or (Test-IsSentinelStudent $Sortie.Student)) {
                                                Write-Host "Skipped a hooked event because it has no assigned student." -ForegroundColor Yellow
                                                continue
                                }

                                $studentName = Get-NameText $Sortie.Student
                                $matchedStudent = $Students | Where-Object { $_.Name -eq $studentName } | Select-Object -First 1
                                if ($null -eq $matchedStudent) {
                                                Write-Host "Could not find student '$studentName' in the student list. Unsat flag was not set." -ForegroundColor Yellow
                                                continue
                                }

                                if (Test-IsSimSortieType ([string]$Sortie.SortieType)) {
                                                $matchedStudent.UnsatSim = $true
                                                Write-Host "Logged UnsatSim for $($matchedStudent.Name) (sortie type: $($Sortie.SortieType))." -ForegroundColor Cyan
                                }
                                else {
                                                $matchedStudent.UnsatFlight = $true
                                                Write-Host "Logged UnsatFlight for $($matchedStudent.Name) (sortie type: $($Sortie.SortieType))." -ForegroundColor Cyan
                                }

                                # Annotate the schedule row so the unsat is visible on output
                                Add-SortieNote -Sortie $Sortie -Note 'Event Hooked (Unsat)'
                                $updatedCount++
                }
                return $updatedCount
}


# ========================================================================================
# INCOMPLETE: the event goes back to not-complete and SimDebt is recalculated
# Returns the students whose SimDebt changed
# ========================================================================================
function Process-IncompleteEvents {
                param(
                                [Parameter(Mandatory)] $SelectedItems,
                                [Parameter(Mandatory)] $Students,
                                $SoftWeights
                )

                if ($null -eq $SoftWeights) {
                                Write-Warning "SoftWeights was not provided. Incomplete events can be processed, but SimDebt will not be recalculated or saved."
                }

                $updatedStudents = New-Object System.Collections.Generic.List[object]
                foreach ($Item in $SelectedItems) {
                                $Sortie = $Item.Sortie
                                if ($null -eq $Sortie -or (Test-IsSentinelStudent $Sortie.Student)) {
                                                Write-Warning "Skipped an incomplete event because it has no assigned student."
                                                continue
                                }

                                $studentName = Get-NameText $Sortie.Student
                                $missionName = Get-NameText $Sortie.Mission
                                $matchedStudent = $Students | Where-Object { $_.Name -eq $studentName } | Select-Object -First 1

                                if ($null -eq $matchedStudent) {
                                                Write-Warning "Could not find student '$studentName' in the student list."
                                }
                                elseif ([string]::IsNullOrWhiteSpace($missionName)) {
                                                Write-Warning "Could not identify a mission for student '$studentName'."
                                }
                                else {
                                                $syllabusMission = Find-SyllabusMission -Student $matchedStudent -Mission $missionName
                                                if ($null -ne $syllabusMission) {
                                                                $syllabusMission.IsComplete = $false
                                                                $syllabusMission.Instructor = ''
                                                                $syllabusMission.ETD = ''
                                                                if ($null -ne $SoftWeights) {
                                                                                $newDebt = Update-StudentSimDebtFromSyllabus -Student $matchedStudent -SoftWeights $SoftWeights
                                                                                if ($null -ne $newDebt) {
                                                                                                Write-Host "Updated SimDebt for $($matchedStudent.Name): $newDebt" -ForegroundColor Cyan
                                                                                                if (-not $updatedStudents.Contains($matchedStudent)) { $updatedStudents.Add($matchedStudent) }
                                                                                }
                                                                }
                                                }
                                                else {
                                                                Write-Warning "Mission '$missionName' was not found in $studentName's syllabus. SimDebt was not changed."
                                                }
                                }

                                # Preserve the existing schedule update behavior
                                $Sortie.Student = $null
                                $Sortie.Mission = $null
                                Add-SortieNote -Sortie $Sortie -Note 'Incomplete'
                }
                return $updatedStudents.ToArray()
}


# ========================================================================================
# The pre-run loop. Returns $true to continue to the puck board, $false to stop.
# ========================================================================================
function Invoke-PreRunMenu {
                param (
                                $Schedule,
                                $Students,
                                $SoftWeights,
                                [string]$MasterDataPath
                )

                while ($true) {
                                $UserChoice = Show-PreRunMenu

                                if ($UserChoice -eq "Hook") {
                                                $SelectedItems = Show-EventSelectionMenu -Schedule $Schedule -Title "Select Hooked Events"
                                                if ($null -ne $SelectedItems) {
                                                                $hookedCount = Process-HookedEvents -SelectedItems $SelectedItems -Students $Students
                                                                if ($hookedCount -gt 0) {
                                                                                [void][System.Windows.Forms.MessageBox]::Show("Hooked events recorded successfully ($hookedCount).", "Success",
                                                                                                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                                                                }
                                                                else {
                                                                                [void][System.Windows.Forms.MessageBox]::Show("No hooked events were applied. Check the console for details.", "Nothing Updated",
                                                                                                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                                                                }
                                                }
                                }
                                elseif ($UserChoice -eq "Incomplete") {
                                                $SelectedItems = Show-EventSelectionMenu -Schedule $Schedule -Title "Select Incomplete Events"
                                                if ($null -ne $SelectedItems) {
                                                                $updatedStudents = @(Process-IncompleteEvents -SelectedItems $SelectedItems -Students $Students -SoftWeights $SoftWeights)

                                                                # Persist SimDebt only for students changed during this selection
                                                                $simDebtByName = @{}
                                                                foreach ($student in $updatedStudents) {
                                                                                if (-not [string]::IsNullOrWhiteSpace([string]$student.Name) -and $null -ne $student.SimDebt) {
                                                                                                $simDebtByName[$student.Name] = [double]$student.SimDebt
                                                                                }
                                                                }
                                                                if ($simDebtByName.Count -gt 0) {
                                                                                if (-not [string]::IsNullOrWhiteSpace($MasterDataPath) -and (Test-Path -LiteralPath $MasterDataPath)) {
                                                                                                try {
                                                                                                                Update-MasterDataSimDebtColumnP -SimDebtByName $simDebtByName -MasterDataPath $MasterDataPath
                                                                                                }
                                                                                                catch {
                                                                                                                Write-Warning "SimDebt changed in memory but could not be saved to MasterData.xlsx: $($_.Exception.Message)"
                                                                                                }
                                                                                }
                                                                                else {
                                                                                                Write-Warning "SimDebt changed in memory but was not saved because MasterDataPath is unavailable."
                                                                                }
                                                                }

                                                                [void][System.Windows.Forms.MessageBox]::Show("Incomplete events processed and SimDebt recalculated.", "Success",
                                                                                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                                                }
                                }
                                elseif ($UserChoice -eq "Run") {
                                                return $true
                                }
                                else {
                                                Write-Host "Exiting Program..." -ForegroundColor Yellow
                                                return $false
                                }
                }
}
