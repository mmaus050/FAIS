# UI.PSM1 DIRECTORY:
#             Contains all functions that make the program easier to understand
# InstructionManualPopUp
# SortFormations
# PreRunMenu
# EventSelectionMenu
# SortieNotes
# ApplyBorders
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
#
#
# ============================================================================================================


# ========================================================================================
#  ____                 ____                __  __
# |  _ \ _ __ ___      |  _ \ _   _ _ __   |  \/  | ___ _ __  _   _
# | |_) | '__/ _ \_____| |_) | | | | '_ \  | |\/| |/ _ \ '_ \| | | |
# |  __/| | |  __/_____|  _ <| |_| | | | | | |  | |  __/ | | | |_| |
# |_|   |_|  \___|     |_| \_\\__,_|_| |_| |_|  |_|\___|_| |_|\__,_|
# ========================================================================================
function PreRunMenu {
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
                $btnHook.Text = "Log Hooked Event (Unsat) TESTING"
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
                $btnRun.Text = "Ready to Run Program ->"
                $btnRun.Location = New-Object System.Drawing.Point(20, 180)
                $btnRun.Size = New-Object System.Drawing.Size(340, 40)
                $btnRun.BackColor = [System.Drawing.Color]::LightGreen
                $btnRun.Add_Click({ $Form.Tag = "Run"; $Form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $Form.Close() })
                $Form.Controls.Add($btnRun)

                $result = $Form.ShowDialog()
                if ($result -eq [System.Windows.Forms.DialogResult]::OK) { return $Form.Tag } else { return "Exit" }
}


# ========================================================================================
#  _____                 _     ____       _           _   _               __  __
# | ____|_   _____ _ __ | |_  / ___|  ___| | ___  ___| |_(_) ___  _ __   |  \/  | ___ _ __  _   _
# |  _| \ \ / / _ \ '_ \| __| \___ \ / _ \ |/ _ \/ __| __| |/ _ \| '_ \  | |\/| |/ _ \ '_ \| | | |
# | |___ \ V /  __/ | | | |_   ___) |  __/ |  __/ (__| |_| | (_) | | | | | |  | |  __/ | | | |_| |
# |_____| \_/ \___|_| |_|\__| |____/ \___|_|\___|\___|\__|_|\___/|_| |_| |_|  |_|\___|_| |_|\__,_|
# ========================================================================================
# Displays the list of events to select from
function Show-EventSelectionMenu {
    param (
        $Schedule,
        $Title
    )

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

    # Excel imports Student values as strings. Do not require a PSCustomObject.
    $ScheduledSorties = @(
        $Schedule |
            Where-Object {
                $null -ne $_.Student -and
                -not [string]::IsNullOrWhiteSpace([string]$_.Student)
            } |
            Sort-Object -Property {
                try {
                    [datetime]$_.ETD
                }
                catch {
                    [datetime]::MinValue
                }
            } -Descending
    )

    foreach ($Sortie in $ScheduledSorties) {

        # Supports Student as either an object with .Name or an Excel text value.
        $studentName = if (
            $Sortie.Student -isnot [string] -and
            $null -ne $Sortie.Student.PSObject.Properties['Name']
        ) {
            [string]$Sortie.Student.Name
        }
        else {
            [string]$Sortie.Student
        }

        # Supports Mission as either an object with .Mission or an Excel text value.
        $missionName = if (
            $null -ne $Sortie.Mission -and
            $Sortie.Mission -isnot [string] -and
            $null -ne $Sortie.Mission.PSObject.Properties['Mission']
        ) {
            [string]$Sortie.Mission.Mission
        }
        elseif ($null -ne $Sortie.Mission) {
            [string]$Sortie.Mission
        }
        else {
            'NONE'
        }

        # Supports Instructor as either an object with .Name or an Excel text value.
        $ipName = if (
            $null -ne $Sortie.Instructor -and
            $Sortie.Instructor -isnot [string] -and
            $null -ne $Sortie.Instructor.PSObject.Properties['Name']
        ) {
            [string]$Sortie.Instructor.Name
        }
        elseif (-not [string]::IsNullOrWhiteSpace([string]$Sortie.Instructor)) {
            [string]$Sortie.Instructor
        }
        else {
            'NO IP'
        }

        $displayString = "{0,-20} | {1,-30} | SP: {2,-15} | Mis: {3,-10} | {4}" -f `
            $Sortie.ETD,
            $ipName,
            $studentName,
            $missionName,
            $Sortie.SortieType

        $listItem = [PSCustomObject]@{
            Display = $displayString
            Sortie  = $Sortie
        }

        [void]$CheckedListBox.Items.Add($listItem)
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
            $Form.Tag = $CheckedListBox.CheckedItems
            $Form.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $Form.Close()
        }
        else {
            [System.Windows.Forms.MessageBox]::Show(
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

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $Form.Tag
    }

    return $null
}
