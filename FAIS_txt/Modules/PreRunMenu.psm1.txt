# PRERUNMENU.PSM1 DIRECTORY:
#             The pop-ups that run before the puck board: log hooked (unsat) and
#             incomplete events from the previous schedule, and change students' HOLD status
# Initialize-FaisWinForms
# Show-PreRunMenu
# Show-EventSelectionMenu
# Show-HoldStatusMenu
# Process-HookedEvents
# Process-IncompleteEvents
# Update-StudentHoldStatus
# Save-PreRunChanges
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
# Returns "Hook", "Incomplete", "Hold", "Run" or "Exit"
function Show-PreRunMenu {
                Initialize-FaisWinForms

                $Form = New-Object System.Windows.Forms.Form
                $Form.Text = "Pre-Run Operations Menu"
                $Form.Size = New-Object System.Drawing.Size(400, 350)
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

                $btnHold = New-Object System.Windows.Forms.Button
                $btnHold.Text = "Change Student HOLD status"
                $btnHold.Location = New-Object System.Drawing.Point(20, 160)
                $btnHold.Size = New-Object System.Drawing.Size(340, 40)
                $btnHold.Add_Click({ $Form.Tag = "Hold"; $Form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $Form.Close() })
                $Form.Controls.Add($btnHold)

                $btnRun = New-Object System.Windows.Forms.Button
                $btnRun.Text = "Open Puck Board ->"
                $btnRun.Location = New-Object System.Drawing.Point(20, 230)
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
# HOLD STATUS: every student with a check box, checked = on hold. Returns name -> $true/$false
# for every student listed, or $null on cancel.
# ========================================================================================
function Show-HoldStatusMenu {
                param ($Students)

                Initialize-FaisWinForms

                $Form = New-Object System.Windows.Forms.Form
                $Form.Text = "Change Student HOLD status"
                $Form.Size = New-Object System.Drawing.Size(450, 560)
                $Form.StartPosition = "CenterScreen"
                $Form.Font = New-Object System.Drawing.Font("Segoe UI", 11)

                $HeaderLabel = New-Object System.Windows.Forms.Label
                $HeaderLabel.Text = "Checked = on hold. Click a student to change it:"
                $HeaderLabel.Location = New-Object System.Drawing.Point(20, 20)
                $HeaderLabel.AutoSize = $true
                $Form.Controls.Add($HeaderLabel)

                $CheckedListBox = New-Object System.Windows.Forms.CheckedListBox
                $CheckedListBox.Location = New-Object System.Drawing.Point(20, 60)
                $CheckedListBox.Size = New-Object System.Drawing.Size(390, 380)
                $CheckedListBox.CheckOnClick = $true
                foreach ($Student in $Students) {
                                $name = ([string]$Student.Name).Trim()
                                if ($name -eq '') { continue }
                                $index = $CheckedListBox.Items.Add($name)
                                $CheckedListBox.SetItemChecked($index, (ConvertTo-FaisBool $Student.OnHold))
                }
                $Form.Controls.Add($CheckedListBox)

                $SaveBtn = New-Object System.Windows.Forms.Button
                $SaveBtn.Location = New-Object System.Drawing.Point(20, 460)
                $SaveBtn.Size = New-Object System.Drawing.Size(180, 40)
                $SaveBtn.Text = "Save HOLD Status"
                $SaveBtn.BackColor = [System.Drawing.Color]::LightGreen
                $SaveBtn.Add_Click({
                                $holdByName = @{}
                                for ($i = 0; $i -lt $CheckedListBox.Items.Count; $i++) {
                                                $holdByName[[string]$CheckedListBox.Items[$i]] = $CheckedListBox.GetItemChecked($i)
                                }
                                $Form.Tag = $holdByName
                                $Form.DialogResult = [System.Windows.Forms.DialogResult]::OK
                                $Form.Close()
                })
                $Form.Controls.Add($SaveBtn)

                $ExitBtn = New-Object System.Windows.Forms.Button
                $ExitBtn.Location = New-Object System.Drawing.Point(230, 460)
                $ExitBtn.Size = New-Object System.Drawing.Size(180, 40)
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
# HOOKED (UNSAT): sim unsat blocks later flights; flight unsat blocks later sims.
# A hooked OFT needs an OFT to clean it up; a hooked IFT an OFT or IFT.
# Returns the students whose unsat flag was set
# ========================================================================================
function Process-HookedEvents {
                param (
                                [Parameter(Mandatory)] $SelectedItems,
                                [Parameter(Mandatory)] $Students
                )

                $hooked = New-Object System.Collections.Generic.List[object]
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
                                                $kind = if ([string]$Sortie.SortieType -like '*OFT*') { 'OFT' } else { 'IFT' }
                                                $matchedStudent.UnsatSim = $true
                                                $matchedStudent | Add-Member -NotePropertyName 'UnsatSimType' -NotePropertyValue $kind -Force
                                                Write-Host "Logged UnsatSim for $($matchedStudent.Name) (sortie type: $($Sortie.SortieType))." -ForegroundColor Cyan
                                }
                                else {
                                                $matchedStudent.UnsatFlight = $true
                                                Write-Host "Logged UnsatFlight for $($matchedStudent.Name) (sortie type: $($Sortie.SortieType))." -ForegroundColor Cyan
                                }

                                # Annotate the schedule row so the unsat is visible on output
                                Add-SortieNote -Sortie $Sortie -Note 'Event Hooked (Unsat)'
                                if (-not $hooked.Contains($matchedStudent)) { $hooked.Add($matchedStudent) }
                }
                return $hooked.ToArray()
}


# ========================================================================================
# INCOMPLETE: the event goes back to not-complete and only its sim debt is backed out
# Returns the students that changed
# ========================================================================================
function Process-IncompleteEvents {
                param(
                                [Parameter(Mandatory)] $SelectedItems,
                                [Parameter(Mandatory)] $Students,
                                $SoftWeights
                )

                if ($null -eq $SoftWeights) {
                                Write-Warning "SoftWeights was not provided. Incomplete events can be processed, but SimDebt will not change."
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
                                                if ($null -eq $syllabusMission) {
                                                                Write-Warning "Mission '$missionName' was not found in $studentName's syllabus. SimDebt was not changed."
                                                }
                                                elseif (Undo-StudentMission -Student $matchedStudent -Mission $syllabusMission -SoftWeights $SoftWeights) {
                                                                Write-Host "$missionName is incomplete for $($matchedStudent.Name); SimDebt is now $($matchedStudent.SimDebt)." -ForegroundColor Cyan
                                                                if (-not $updatedStudents.Contains($matchedStudent)) { $updatedStudents.Add($matchedStudent) }
                                                }
                                                else {
                                                                Write-Host "$missionName was already incomplete for $($matchedStudent.Name); nothing changed." -ForegroundColor Yellow
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
# HOLD STATUS: sets OnHold on the students whose box changed. Returns those changes in the
# Update-MasterDataStudents -FlagsByName shape (name -> @{ OnHold = $true/$false })
# ========================================================================================
function Update-StudentHoldStatus {
                param (
                                [Parameter(Mandatory)] [hashtable]$HoldByName,
                                [Parameter(Mandatory)] $Students
                )

                $changes = @{}
                foreach ($Student in $Students) {
                                $name = ([string]$Student.Name).Trim()
                                if ($name -eq '' -or -not $HoldByName.ContainsKey($name)) { continue }
                                $hold = [bool]$HoldByName[$name]
                                if ($hold -eq (ConvertTo-FaisBool $Student.OnHold)) { continue }
                                $Student.OnHold = $hold
                                $changes[$name] = @{ OnHold = $hold }
                                Write-Host "$name is $(if ($hold) { 'now on hold' } else { 'off hold' })." -ForegroundColor Cyan
                }
                return $changes
}


# ========================================================================================
# Saves a pre-run change right away, both files or neither: SimDebt / unsat flags to
# MasterData and, for incomplete events, those events' cells on the students' tabs in
# OutputSchedule (so the two files agree even if the board is closed without exporting).
# A file open in Excel is saved and closed for the write, then opened again.
# A change that can't be saved stays in memory and goes out with the Export.
# Returns $true when saved.
# ========================================================================================
function Save-PreRunChanges {
                param (
                                [string]$MasterDataPath,
                                [string]$OutputPath,
                                [hashtable]$SimDebtByName = @{},
                                [hashtable]$FlagsByName = @{},
                                [array]$TabStudents = @(),
                                [scriptblock]$LockedPrompt
                )

                if ([string]::IsNullOrWhiteSpace($MasterDataPath) -or -not (Test-Path -LiteralPath $MasterDataPath)) {
                                Write-Warning "MasterData.xlsx was not found, so this change is only kept for this run (Export writes it)."
                                return $false
                }
                $writeTabs = ($TabStudents.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($OutputPath) -and (Test-Path -LiteralPath $OutputPath))
                $toSave = @($MasterDataPath)
                if ($writeTabs) { $toSave += $OutputPath }
                $closed = New-Object System.Collections.Generic.List[string]
                try {
                                foreach ($file in $toSave) {
                                                if (-not (Request-FaisFileWritable -Path $file -Prompt $LockedPrompt -Closed $closed)) {
                                                                Write-Warning "$file is still open, so this change is only kept for this run (Export writes it)."
                                                                return $false
                                                }
                                }

                                try {
                                                Invoke-WithFileBackup -Files $toSave -Action {
                                                                Update-MasterDataStudents -MasterDataPath $MasterDataPath -SimDebtByName $SimDebtByName -FlagsByName $FlagsByName
                                                                if ($writeTabs) { Export-StudentSyllabusTabs -Students $TabStudents -FilePath $OutputPath }
                                                }
                                                return $true
                                }
                                catch {
                                                Write-Warning "This change could not be saved, so it is only kept for this run (Export writes it): $($_.Exception.Message)"
                                                return $false
                                }
                }
                finally {
                                if ($closed.Count -gt 0) { Open-FaisWorkbooks -Path $closed.ToArray() }
                }
}


# ========================================================================================
# The pre-run loop. Returns $true to continue to the puck board, $false to stop.
# ========================================================================================
function Invoke-PreRunMenu {
                param (
                                $Schedule,
                                $Students,
                                $SoftWeights,
                                [string]$MasterDataPath,
                                [string]$OutputPath,
                                [scriptblock]$LockedPrompt
                )

                if ($null -eq $LockedPrompt) {
                                $LockedPrompt = {
                                                param($path)
                                                $answer = [System.Windows.Forms.MessageBox]::Show("$path is still open and FAIS could not close it (a cell may be in edit mode, or someone else has it open).`n`nClose it, then press OK to try again.", 'File is open',
                                                                [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Warning)
                                                return ($answer -eq [System.Windows.Forms.DialogResult]::OK)
                                }
                }

                while ($true) {
                                $UserChoice = Show-PreRunMenu

                                if ($UserChoice -eq "Hook") {
                                                $SelectedItems = Show-EventSelectionMenu -Schedule $Schedule -Title "Select Hooked Events"
                                                if ($null -ne $SelectedItems) {
                                                                $hooked = @(Process-HookedEvents -SelectedItems $SelectedItems -Students $Students)
                                                                if ($hooked.Count -gt 0) {
                                                                                # UnsatSim / UnsatFlight = TRUE in MasterData now
                                                                                $flags = Get-UnsatFlagUpdates -Students $hooked
                                                                                $saved = ($flags.Count -eq 0) -or (Save-PreRunChanges -MasterDataPath $MasterDataPath -OutputPath $OutputPath -FlagsByName $flags -LockedPrompt $LockedPrompt)
                                                                                if ($saved) {
                                                                                                foreach ($s in $hooked) {
                                                                                                                $s | Add-Member -NotePropertyName 'UnsatSimInMasterData' -NotePropertyValue (ConvertTo-FaisBool $s.UnsatSim) -Force
                                                                                                                $s | Add-Member -NotePropertyName 'UnsatFlightInMasterData' -NotePropertyValue (ConvertTo-FaisBool $s.UnsatFlight) -Force
                                                                                                }
                                                                                }
                                                                                $where = if ($saved) { 'and saved to MasterData' } else { 'for this run only (MasterData could not be saved; Export will write it)' }
                                                                                [void][System.Windows.Forms.MessageBox]::Show("Hooked events recorded for $($hooked.Count) student(s) $where.", "Success",
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

                                                                # SimDebt and those events' tab cells, for the students changed by this selection
                                                                $simDebtByName = @{}
                                                                foreach ($student in $updatedStudents) {
                                                                                if (-not [string]::IsNullOrWhiteSpace([string]$student.Name) -and $null -ne $student.SimDebt) {
                                                                                                $simDebtByName[([string]$student.Name).Trim()] = [double]$student.SimDebt
                                                                                }
                                                                }
                                                                $text = "No events changed (they were already incomplete). Check the console for details."
                                                                if ($updatedStudents.Count -gt 0) {
                                                                                $saved = Save-PreRunChanges -MasterDataPath $MasterDataPath -OutputPath $OutputPath -SimDebtByName $simDebtByName -TabStudents $updatedStudents -LockedPrompt $LockedPrompt
                                                                                $text = "Incomplete events processed. SimDebt and the students' tabs are saved."
                                                                                if (-not $saved) { $text = "Incomplete events processed for this run only (the files could not be saved; Export will write them)." }
                                                                }
                                                                [void][System.Windows.Forms.MessageBox]::Show($text, "Incomplete Events",
                                                                                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                                                }
                                }
                                elseif ($UserChoice -eq "Hold") {
                                                $holdByName = Show-HoldStatusMenu -Students $Students
                                                if ($null -ne $holdByName) {
                                                                $changes = Update-StudentHoldStatus -HoldByName $holdByName -Students $Students
                                                                $text = "No HOLD status changed."
                                                                if ($changes.Count -gt 0) {
                                                                                # OnHold = TRUE / FALSE in MasterData now
                                                                                $saved = Save-PreRunChanges -MasterDataPath $MasterDataPath -OutputPath $OutputPath -FlagsByName $changes -LockedPrompt $LockedPrompt
                                                                                if ($saved) {
                                                                                                foreach ($s in $Students) {
                                                                                                                if ($changes.ContainsKey(([string]$s.Name).Trim())) {
                                                                                                                                $s | Add-Member -NotePropertyName 'OnHoldInMasterData' -NotePropertyValue (ConvertTo-FaisBool $s.OnHold) -Force
                                                                                                                }
                                                                                                }
                                                                                }
                                                                                $where = if ($saved) { 'and saved to MasterData' } else { 'for this run only (MasterData could not be saved; Export will write it)' }
                                                                                $text = "HOLD status changed for $($changes.Count) student(s) $where."
                                                                }
                                                                [void][System.Windows.Forms.MessageBox]::Show($text, "HOLD Status",
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
