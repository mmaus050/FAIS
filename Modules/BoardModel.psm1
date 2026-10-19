# BOARDMODEL.PSM1 DIRECTORY:
#             The puck board's data: days, lines, pins. No WPF in here, so all of it can
#             be tested without a window.
# Get-DefaultBoardSettings / New-Board / New-BoardLine / Get-BoardLine / Get-BoardDayLines
# Undo: Save-BoardUndoPoint / Undo-Board
# Edits: Add-BoardDay / Remove-BoardDay / Add-BoardLine / Remove-BoardLine / Set-BoardLineEtd
#        Set-BoardLineType / Set-BoardInstructor / Set-BoardPin / Clear-BoardPin / Move-BoardPin
#        Set-BoardMission / Set-BoardLineField / Set-FormationPair / Unlink-ReturnLine
# Typing: ConvertTo-BoardTime / Format-BoardTime / ConvertTo-BoardSortieType / ConvertFrom-QuickEntry
# Engine: ConvertTo-ScheduleRows / ConvertFrom-ScheduleRows / Get-FillOverlay / Test-BoardReadyToFill
# Files:  Save-BoardState / Open-BoardState / Get-SavedBoardDays
#         Save-BoardTemplate / Get-BoardTemplateList / Import-BoardTemplate / Copy-BoardPreviousDay
# Get-NextFlyingDay


# =================================================================================
#  BOARD AND LINES
# =================================================================================

function Get-DefaultBoardSettings {
                return @{
                                OBReturnOffsetMinutes = 180
                                ReturnSortieType      = 'O&B Return'
                                SimInstructor         = 'SIM IP'
                                UndoDepth             = 50
                }
}

function New-Board {
                param (
                                [string[]]$Days = @(),
                                [hashtable]$Settings
                )

                $s = Get-DefaultBoardSettings
                if ($null -ne $Settings) { foreach ($k in $Settings.Keys) { $s[$k] = $Settings[$k] } }

                $dayList = New-Object System.Collections.Generic.List[string]
                foreach ($d in $Days) { $n = ConvertTo-BoardDay $d; if ($n -and -not $dayList.Contains($n)) { $dayList.Add($n) } }
                $dayList.Sort()

                return [PSCustomObject]@{
                                SchemaVersion = 1
                                Days          = $dayList
                                Lines         = (New-Object System.Collections.Generic.List[object])
                                Version       = 0
                                Undo          = (New-Object System.Collections.Generic.List[string])
                                Settings      = $s
                }
}

# 'yyyy-MM-dd' from a date or text
function ConvertTo-BoardDay {
                param ($Day)

                if ($null -eq $Day) { return $null }
                if ($Day -is [datetime]) { return $Day.ToString('yyyy-MM-dd') }
                $dt = Parse-ETD $Day
                if ($null -eq $dt) { return $null }
                return $dt.ToString('yyyy-MM-dd')
}

function New-BoardId {
                return [guid]::NewGuid().ToString('N').Substring(0, 8)
}

function New-BoardLine {
                param (
                                [Parameter(Mandatory)] [datetime]$ETD,
                                [Parameter(Mandatory)] [string]$SortieType,
                                [string]$Instructor = '',
                                [string]$Student = '',
                                [string]$Mission = '',
                                [string]$Callsign = '',
                                [string]$Num = '',
                                [string]$Notes = '',
                                [string]$Source = 'Manual',
                                [string]$Id = '',
                                [string]$LinkedLineId = '',
                                [string]$WingOfLineId = ''
                )

                if ($Id -eq '') { $Id = New-BoardId }
                return [PSCustomObject]@{
                                Id           = $Id
                                Day          = $ETD.ToString('yyyy-MM-dd')
                                ETD          = $ETD
                                SortieType   = $SortieType
                                Instructor   = $Instructor
                                Student      = $Student
                                Mission      = $Mission
                                Callsign     = $Callsign
                                Num          = $Num
                                Notes        = $Notes
                                LinkedLineId = $LinkedLineId
                                WingOfLineId = $WingOfLineId
                                Source       = $Source
                }
}

function Get-BoardLine {
                param ($Board, [string]$Id)

                if ([string]::IsNullOrWhiteSpace($Id)) { return $null }
                foreach ($l in $Board.Lines) { if ($l.Id -eq $Id) { return $l } }
                return $null
}

# A day's lines in ETD order (ties keep the order they were added)
function Get-BoardDayLines {
                param ($Board, [string]$Day)

                $i = 0
                $tagged = foreach ($l in $Board.Lines) {
                                if ([string]::IsNullOrEmpty($Day) -or $l.Day -eq $Day) { [PSCustomObject]@{ L = $l; I = $i } }
                                $i++
                }
                return @($tagged | Sort-Object -Property @{ Expression = { $_.L.ETD } }, I | ForEach-Object { $_.L })
}

function Test-BoardLineIsReturn {
                param ($Line)
                return (Test-IsReturnSortieType ([string]$Line.SortieType))
}

function Test-BoardLineIsOB {
                param ($Line)
                return (Test-IsOBSortieType ([string]$Line.SortieType))
}


# =================================================================================
#  UNDO
#  Each edit saves a JSON snapshot first. Version always goes up (also on undo) so
#  cached assessments are never reused after a change.
# =================================================================================

function ConvertTo-BoardSnapshot {
                param ($Board)

                $dto = [PSCustomObject]@{
                                Days  = $Board.Days.ToArray()
                                Lines = @(foreach ($l in $Board.Lines) { ConvertTo-LineDto $l })
                }
                return (ConvertTo-Json -InputObject $dto -Depth 6 -Compress)
}

function Restore-BoardSnapshot {
                param ($Board, [string]$Json)

                $dto = ConvertFrom-Json $Json
                $Board.Days.Clear()
                foreach ($d in @($dto.Days)) { if ($d) { $Board.Days.Add([string]$d) } }
                $Board.Lines.Clear()
                foreach ($l in @($dto.Lines)) { if ($null -ne $l) { $Board.Lines.Add((ConvertFrom-LineDto $l)) } }
}

function Save-BoardUndoPoint {
                param ($Board)

                $Board.Undo.Add((ConvertTo-BoardSnapshot $Board))
                $depth = [int]$Board.Settings.UndoDepth
                if ($depth -lt 1) { $depth = 50 }
                while ($Board.Undo.Count -gt $depth) { $Board.Undo.RemoveAt(0) }
}

function Step-BoardVersion {
                param ($Board)
                $Board.Version = [int]$Board.Version + 1
}

function Undo-Board {
                param ($Board)

                if ($Board.Undo.Count -eq 0) { return $false }
                $json = $Board.Undo[$Board.Undo.Count - 1]
                $Board.Undo.RemoveAt($Board.Undo.Count - 1)
                Restore-BoardSnapshot -Board $Board -Json $json
                Step-BoardVersion $Board
                return $true
}


# =================================================================================
#  EDITS
# =================================================================================

function Add-BoardDay {
                param ($Board, $Day)

                $d = ConvertTo-BoardDay $Day
                if ($null -eq $d -or $Board.Days.Contains($d)) { return $d }
                Save-BoardUndoPoint $Board
                $Board.Days.Add($d)
                $Board.Days.Sort()
                Step-BoardVersion $Board
                return $d
}

function Remove-BoardDay {
                param ($Board, [string]$Day)

                if (-not $Board.Days.Contains($Day)) { return }
                Save-BoardUndoPoint $Board
                [void]$Board.Days.Remove($Day)
                $keep = @($Board.Lines | Where-Object { $_.Day -ne $Day })
                $Board.Lines.Clear()
                foreach ($l in $keep) { $Board.Lines.Add($l) }
                Step-BoardVersion $Board
}

# Adds lines (one undo step). Days are added for lines on new dates.
function Add-BoardLine {
                param (
                                $Board,
                                [array]$Line,
                                [switch]$NoUndo
                )

                if (-not $NoUndo) { Save-BoardUndoPoint $Board }
                foreach ($l in $Line) {
                                if ($null -eq $l) { continue }
                                if (-not $Board.Days.Contains($l.Day)) { $Board.Days.Add($l.Day); $Board.Days.Sort() }
                                $Board.Lines.Add($l)
                }
                Step-BoardVersion $Board
}

# Removes a line. An O&B takes its linked return with it.
function Remove-BoardLine {
                param ($Board, [string]$Id, [switch]$NoUndo)

                $line = Get-BoardLine $Board $Id
                if ($null -eq $line) { return }
                if (-not $NoUndo) { Save-BoardUndoPoint $Board }

                $gone = @($Id)
                if ((Test-BoardLineIsOB $line) -and $line.LinkedLineId) { $gone += $line.LinkedLineId }
                foreach ($l in $Board.Lines.ToArray()) {
                                if ($gone -contains $l.Id) { [void]$Board.Lines.Remove($l); continue }
                                if ($gone -contains $l.LinkedLineId) { $l.LinkedLineId = '' }
                                if ($gone -contains $l.WingOfLineId) { $l.WingOfLineId = '' }
                }
                Step-BoardVersion $Board
}

# New ETD for a line. An O&B drags its return leg by the same amount.
function Set-BoardLineEtd {
                param (
                                $Board,
                                [string]$Id,
                                [datetime]$ETD
                )

                $line = Get-BoardLine $Board $Id
                if ($null -eq $line) { return }
                Save-BoardUndoPoint $Board
                $delta = $ETD - $line.ETD
                $line.ETD = $ETD
                $line.Day = $ETD.ToString('yyyy-MM-dd')
                if (-not $Board.Days.Contains($line.Day)) { $Board.Days.Add($line.Day); $Board.Days.Sort() }

                if ((Test-BoardLineIsOB $line) -and $line.LinkedLineId) {
                                $ret = Get-BoardLine $Board $line.LinkedLineId
                                if ($null -ne $ret) {
                                                $ret.ETD = $ret.ETD + $delta
                                                $ret.Day = $ret.ETD.ToString('yyyy-MM-dd')
                                }
                }
                Step-BoardVersion $Board
}

# Changes a line's type. Becoming an O&B adds a linked return; leaving O&B removes it.
function Set-BoardLineType {
                param ($Board, [string]$Id, [string]$SortieType)

                $line = Get-BoardLine $Board $Id
                if ($null -eq $line -or $line.SortieType -eq $SortieType) { return }
                Save-BoardUndoPoint $Board

                $wasOB = Test-BoardLineIsOB $line
                $line.SortieType = $SortieType
                $isOB = Test-BoardLineIsOB $line

                if ($wasOB -and -not $isOB -and $line.LinkedLineId) {
                                $retId = $line.LinkedLineId
                                $line.LinkedLineId = ''
                                foreach ($l in $Board.Lines.ToArray()) { if ($l.Id -eq $retId) { [void]$Board.Lines.Remove($l) } }
                }
                if ($isOB -and -not $wasOB) {
                                $ret = New-ReturnLineFor -Board $Board -OBLine $line
                                $Board.Lines.Add($ret)
                }
                if (Test-BoardLineIsReturn $line) { $line.Student = ''; $line.Mission = '' }
                Step-BoardVersion $Board
}

function New-ReturnLineFor {
                param ($Board, $OBLine, [Nullable[datetime]]$ReturnETD)

                $offset = [int]$Board.Settings.OBReturnOffsetMinutes
                $etd = if ($null -ne $ReturnETD) { [datetime]$ReturnETD } else { $OBLine.ETD.AddMinutes($offset) }
                $ret = New-BoardLine -ETD $etd -SortieType ([string]$Board.Settings.ReturnSortieType) -Instructor $OBLine.Instructor -Source $OBLine.Source
                $OBLine.LinkedLineId = $ret.Id
                $ret.LinkedLineId = $OBLine.Id
                return $ret
}

# Puts an IP on a line. An O&B's IP flies the return too.
function Set-BoardInstructor {
                param ($Board, [string]$Id, [string]$Name)

                $line = Get-BoardLine $Board $Id
                if ($null -eq $line) { return }
                Save-BoardUndoPoint $Board
                $line.Instructor = $Name
                if ((Test-BoardLineIsOB $line) -and $line.LinkedLineId) {
                                $ret = Get-BoardLine $Board $line.LinkedLineId
                                if ($null -ne $ret) { $ret.Instructor = $Name }
                }
                Step-BoardVersion $Board
}

# Hardcodes a student on a line. Return lines can't be pinned (they follow their O&B).
function Set-BoardPin {
                param ($Board, [string]$Id, [string]$Student, [string]$Mission = '')

                $line = Get-BoardLine $Board $Id
                if ($null -eq $line -or (Test-BoardLineIsReturn $line)) { return $false }
                Save-BoardUndoPoint $Board
                if ($line.Student -ne $Student) { $line.WingOfLineId = '' }
                $line.Student = $Student
                $line.Mission = $Mission
                Step-BoardVersion $Board
                return $true
}

function Clear-BoardPin {
                param ($Board, [string]$Id)

                $line = Get-BoardLine $Board $Id
                if ($null -eq $line) { return }
                Save-BoardUndoPoint $Board
                $line.Student = ''
                $line.Mission = ''
                $line.WingOfLineId = ''
                foreach ($l in $Board.Lines) { if ($l.WingOfLineId -eq $Id) { $l.WingOfLineId = '' } }
                Step-BoardVersion $Board
}

# Moves a pinned student to another line (swaps if that line is pinned). -Copy keeps the original.
function Move-BoardPin {
                param ($Board, [string]$FromId, [string]$ToId, [switch]$Copy)

                $from = Get-BoardLine $Board $FromId
                $to = Get-BoardLine $Board $ToId
                if ($null -eq $from -or $null -eq $to -or $FromId -eq $ToId -or (Test-BoardLineIsReturn $to)) { return $false }
                Save-BoardUndoPoint $Board

                $s = $from.Student; $m = $from.Mission
                if ($Copy) {
                                $to.Student = $s
                                $to.Mission = ''
                }
                else {
                                $from.Student = $to.Student
                                $from.Mission = $to.Mission
                                $to.Student = $s
                                $to.Mission = $m
                }
                foreach ($l in @($from, $to)) { $l.WingOfLineId = '' }
                foreach ($l in $Board.Lines) { if ($l.WingOfLineId -eq $FromId -or $l.WingOfLineId -eq $ToId) { $l.WingOfLineId = '' } }
                Step-BoardVersion $Board
                return $true
}

function Set-BoardMission {
                param ($Board, [string]$Id, [string]$Mission)

                $line = Get-BoardLine $Board $Id
                if ($null -eq $line -or $line.Mission -eq $Mission) { return }
                Save-BoardUndoPoint $Board
                $line.Mission = $Mission
                Step-BoardVersion $Board
}

function Set-BoardLineField {
                param (
                                $Board,
                                [string]$Id,
                                [ValidateSet('Callsign', 'Num', 'Notes')] [string]$Field,
                                [string]$Value
                )

                $line = Get-BoardLine $Board $Id
                if ($null -eq $line -or [string]$line.$Field -eq $Value) { return }
                Save-BoardUndoPoint $Board
                $line.$Field = $Value
                Step-BoardVersion $Board
}

# Marks a pinned line as flying wing to a pinned lead line ('' clears)
function Set-FormationPair {
                param ($Board, [string]$WingId, [string]$LeadId = '')

                $wing = Get-BoardLine $Board $WingId
                if ($null -eq $wing) { return $false }
                if ($LeadId -ne '') {
                                $lead = Get-BoardLine $Board $LeadId
                                if ($null -eq $lead -or $LeadId -eq $WingId) { return $false }
                }
                Save-BoardUndoPoint $Board
                foreach ($l in $Board.Lines) { if ($l.WingOfLineId -eq $LeadId -and $LeadId -ne '') { $l.WingOfLineId = '' } }
                $wing.WingOfLineId = $LeadId
                Step-BoardVersion $Board
                return $true
}

function Unlink-ReturnLine {
                param ($Board, [string]$Id)

                $line = Get-BoardLine $Board $Id
                if ($null -eq $line -or -not $line.LinkedLineId) { return }
                Save-BoardUndoPoint $Board
                $other = Get-BoardLine $Board $line.LinkedLineId
                $line.LinkedLineId = ''
                if ($null -ne $other) { $other.LinkedLineId = '' }
                Step-BoardVersion $Board
}


# =================================================================================
#  TYPING TIMES AND QUICK ENTRY
# =================================================================================

# 0730 / 730 / 7:30 / 07:30 / 07.30 / 0730L / 7 -> [timespan]; $null when invalid
function ConvertTo-BoardTime {
                param ([string]$Text)

                if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
                $t = $Text.Trim()
                $h = -1; $m = -1
                if ($t -match '^(\d{1,2})[:.h](\d{2})\s*[lLzZ]?$') { $h = [int]$Matches[1]; $m = [int]$Matches[2] }
                elseif ($t -match '^(\d{3,4})\s*[lLzZ]?$') { $v = $Matches[1].PadLeft(4, '0'); $h = [int]$v.Substring(0, 2); $m = [int]$v.Substring(2, 2) }
                elseif ($t -match '^(\d{1,2})\s*[lLzZ]?$') { $h = [int]$Matches[1]; $m = 0 }
                else { return $null }
                if ($h -lt 0 -or $h -gt 23 -or $m -lt 0 -or $m -gt 59) { return $null }
                return New-TimeSpan -Hours $h -Minutes $m
}

function Format-BoardTime {
                param ($Value)

                if ($null -eq $Value) { return '' }
                if ($Value -is [timespan]) { return ('{0:00}{1:00}' -f $Value.Hours, $Value.Minutes) }
                $dt = Parse-ETD $Value
                if ($null -eq $dt) { return '' }
                return $dt.ToString('HHmm')
}

# Short names the user can type for each sortie type ($null = not a type, 'AMBIGUOUS' = SIM)
function ConvertTo-BoardSortieType {
                param ([string]$Token, $Settings)

                $returnType = 'O&B Return'
                if ($null -ne $Settings -and $Settings.ReturnSortieType) { $returnType = [string]$Settings.ReturnSortieType }
                $k = ($Token -replace '[\s\-_]', '').ToUpperInvariant()
                switch ($k) {
                                { $_ -in @('DL', 'LOCAL', 'DAYLOCAL', 'D', 'L') } { return 'Day Local' }
                                { $_ -in @('OB', 'O&B', 'ONB', 'OANDB', 'O/B') } { return 'O&B' }
                                { $_ -in @('RTN', 'RET', 'RETURN', 'R', 'O&BRETURN', 'OBRETURN', 'OBRTN', 'O&BRTN') } { return $returnType }
                                'OFT' { return 'OFT' }
                                'IFT' { return 'IFT' }
                                { $_ -in @('SIM', 'S') } { return 'AMBIGUOUS' }
                }
                return $null
}

# Parses what the user types into the quick-entry box. Several entries can be
# separated by "," or ";". Returns {Lines; Errors; Warnings}.
#   0730 DL                 Day Local at 0730
#   1015 O&B / 1015 OB 1330 O&B plus its linked return (default +3 hr, or at 1330)
#   1330 RTN                return line linked to the latest unlinked O&B that day
#   1300 OFT x2             two OFT lines
#   0800 DL Jones sp:Smith  IP Jones, student Smith hardcoded (m:FT5103 sets the mission)
function ConvertFrom-QuickEntry {
                param (
                                [string]$Text,
                                [Parameter(Mandatory)] $Day,
                                [array]$Instructors = @(),
                                [array]$Students = @(),
                                $Board,
                                $Settings
                )

                if ($null -eq $Settings) {
                                $Settings = if ($null -ne $Board) { $Board.Settings } else { Get-DefaultBoardSettings }
                }
                $dayText = ConvertTo-BoardDay $Day
                $dayDate = [datetime]::ParseExact($dayText, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
                $lines = New-Object System.Collections.Generic.List[object]
                $errors = New-Object System.Collections.Generic.List[string]
                $warnings = New-Object System.Collections.Generic.List[string]
                $offset = [int]$Settings.OBReturnOffsetMinutes
                if ($offset -le 0) { $offset = 180 }

                # Multi-word type names become one token first
                $norm = $Text -replace '(?i)\bday\s+local\b', 'DL' -replace '(?i)\bo\s*&\s*b\s+(return|rtn)\b', 'RTN' -replace '(?i)\bo\s*&\s*b\b', 'O&B'

                foreach ($entry in ($norm -split '[,;\r\n]+')) {
                                $e = $entry.Trim()
                                if ($e -eq '') { continue }
                                $tokens = @($e -split '\s+' | Where-Object { $_ -ne '' })

                                $time = ConvertTo-BoardTime $tokens[0]
                                if ($null -eq $time) { $errors.Add("Invalid time '$($tokens[0])' in '$e' (use 0730, 730 or 7:30)"); continue }

                                $type = $null
                                $returnTime = $null
                                $count = 1
                                $ipText = New-Object System.Collections.Generic.List[string]
                                $studentText = ''
                                $missionText = ''
                                $bad = $false

                                for ($i = 1; $i -lt $tokens.Count; $i++) {
                                                $tok = $tokens[$i]
                                                if ($tok -match '^(?i)(sp|stu|student)[:=](.+)$') { $studentText = $Matches[2]; continue }
                                                if ($tok -match '^(?i)(m|msn|mission)[:=](.+)$') { $missionText = $Matches[2]; continue }
                                                if ($tok -match '^(?i)ip[:=](.+)$') { $ipText.Add($Matches[1]); continue }
                                                if ($tok -match '^(?i)x(\d)$') { $count = [int]$Matches[1]; continue }
                                                if ($null -eq $type) {
                                                                $tt = ConvertTo-BoardSortieType $tok $Settings
                                                                if ($tt -eq 'AMBIGUOUS') { $errors.Add("'$tok' is ambiguous - use OFT or IFT ('$e')"); $bad = $true; break }
                                                                if ($null -ne $tt) { $type = $tt; continue }
                                                }
                                                elseif ($type -eq 'O&B' -and $null -eq $returnTime -and $ipText.Count -eq 0) {
                                                                $rt = ConvertTo-BoardTime $tok
                                                                if ($null -ne $rt) { $returnTime = $rt; continue }
                                                }
                                                $ipText.Add($tok)
                                }
                                if ($bad) { continue }
                                if ($null -eq $type) { $errors.Add("Missing sortie type in '$e' (DL, O&B, RTN, OFT, IFT)"); continue }
                                if ($count -lt 1) { $count = 1 }

                                # Instructor: last name, full name or callsign
                                $ipName = ''
                                if ($ipText.Count -gt 0) {
                                                $raw = ($ipText -join ' ').Trim()
                                                $ip = Find-InstructorForEntry -Instructors $Instructors -Text $raw
                                                if ($null -eq $ip) { $errors.Add("IP '$raw' not found ('$e')"); continue }
                                                $ipName = [string]$ip.Name
                                }

                                $studentName = ''
                                if ($studentText -ne '') {
                                                $s = Resolve-StudentByName -Students $Students -Name ($studentText -replace '_', ' ')
                                                if ($null -eq $s) { $errors.Add("Student '$studentText' not found ('$e')"); continue }
                                                if (Test-IsReturnSortieType $type) { $warnings.Add("Return lines follow their O&B; student ignored on '$e'") }
                                                else { $studentName = [string]$s.Name }
                                }

                                $etd = $dayDate.Add($time)
                                for ($c = 0; $c -lt $count; $c++) {
                                                $line = New-BoardLine -ETD $etd -SortieType $type -Instructor $ipName -Student $(if ($c -eq 0) { $studentName } else { '' }) -Mission $(if ($c -eq 0) { $missionText } else { '' })
                                                $lines.Add($line)
                                                if ($type -eq 'O&B') {
                                                                $retEtd = if ($null -ne $returnTime) { $dayDate.Add($returnTime) } else { $etd.AddMinutes($offset) }
                                                                if ($retEtd -le $etd) { $warnings.Add("Return time $(Format-BoardTime $retEtd) is not after the O&B in '$e'") }
                                                                $lines.Add((New-ReturnLineFor -Board ([PSCustomObject]@{ Settings = $Settings }) -OBLine $line -ReturnETD $retEtd))
                                                }
                                                elseif (Test-IsReturnSortieType $type) {
                                                                $ob = Find-UnlinkedOBForReturn -Board $Board -Pending $lines -ReturnLine $line
                                                                if ($null -ne $ob) {
                                                                                $ob.LinkedLineId = $line.Id
                                                                                $line.LinkedLineId = $ob.Id
                                                                                if ($line.Instructor -eq '') { $line.Instructor = $ob.Instructor }
                                                                }
                                                                else { $warnings.Add("No O&B without a return before $(Format-BoardTime $etd); return line added unlinked") }
                                                }
                                }
                }

                return [PSCustomObject]@{
                                Lines    = $lines.ToArray()
                                Errors   = $errors.ToArray()
                                Warnings = $warnings.ToArray()
                }
}

function Find-InstructorForEntry {
                param ([array]$Instructors, [string]$Text)

                $t = $Text.Trim().TrimEnd(',')
                if ($t -eq '') { return $null }
                $byLast = @($Instructors | Where-Object { (Get-NormalizedLastName (Get-BaseName ([string]$_.Name))) -ieq $t })
                if ($byLast.Count -eq 1) { return $byLast[0] }
                return (Get-MatchedInstructor -Instructors $Instructors -TargetName $t)
}

function Find-UnlinkedOBForReturn {
                param ($Board, $Pending, $ReturnLine)

                $pool = New-Object System.Collections.Generic.List[object]
                if ($null -ne $Board) { foreach ($l in $Board.Lines) { $pool.Add($l) } }
                foreach ($l in $Pending) { if ($l.Id -ne $ReturnLine.Id) { $pool.Add($l) } }
                $cands = @($pool | Where-Object {
                                (Test-BoardLineIsOB $_) -and -not $_.LinkedLineId -and $_.Day -eq $ReturnLine.Day -and $_.ETD -lt $ReturnLine.ETD
                } | Sort-Object -Property ETD -Descending)
                if ($cands.Count -gt 0) { return $cands[0] }
                return $null
}


# =================================================================================
#  ENGINE HAND-OFF
# =================================================================================

# The board's lines as schedule rows, in exactly the shape IngestData gave the
# engine from the Schedule sheet (plus the board ids)
function ConvertTo-ScheduleRows {
                param (
                                $Board,
                                [string[]]$Days
                )

                $simIP = 'SIM IP'
                if ($Board.Settings.SimInstructor) { $simIP = [string]$Board.Settings.SimInstructor }
                $useDays = if ($null -ne $Days -and $Days.Count -gt 0) { $Days } else { @($Board.Days) }

                $out = New-Object System.Collections.Generic.List[object]
                foreach ($day in ($useDays | Sort-Object)) {
                                foreach ($l in (Get-BoardDayLines -Board $Board -Day $day)) {
                                                $ip = ([string]$l.Instructor).Trim()
                                                if ($ip -eq '' -and (Test-IsSimSortieType ([string]$l.SortieType))) { $ip = $simIP }
                                                $isReturn = Test-BoardLineIsReturn $l
                                                $out.Add([PSCustomObject]@{
                                                                ETD          = $l.ETD.ToString('yyyy-MM-dd HH:mm:ss')
                                                                Instructor   = $(if ($ip -eq '') { $null } else { $ip })
                                                                SortieType   = [string]$l.SortieType
                                                                Student      = $(if ($isReturn -or [string]::IsNullOrWhiteSpace($l.Student)) { $null } else { ([string]$l.Student).Trim() })
                                                                Mission      = $(if ($isReturn -or [string]::IsNullOrWhiteSpace($l.Mission)) { $null } else { ([string]$l.Mission).Trim() })
                                                                Callsign     = $(if ([string]::IsNullOrWhiteSpace($l.Callsign)) { $null } else { [string]$l.Callsign })
                                                                Num          = $(if ([string]::IsNullOrWhiteSpace($l.Num)) { $null } else { [string]$l.Num })
                                                                Notes        = $(if ([string]::IsNullOrWhiteSpace($l.Notes)) { $null } else { [string]$l.Notes })
                                                                BoardLineId  = $l.Id
                                                                LinkedLineId = $(if ($l.LinkedLineId) { $l.LinkedLineId } else { $null })
                                                                WingOfLineId = $(if ($l.WingOfLineId) { $l.WingOfLineId } else { $null })
                                                                FilledBy     = $null
                                                })
                                }
                }
                return $out.ToArray()
}

# Schedule rows (Schedule sheet or a previous MasterSchedule) as board lines.
# Engine placeholders (For Sale, Need Form SP) are dropped; O&Bs are linked to their returns.
function ConvertFrom-ScheduleRows {
                param (
                                [array]$Rows,
                                [string]$Source = 'Imported',
                                [switch]$KeepStudents,
                                [switch]$KeepNotes,
                                [switch]$NoInstructors,
                                [string]$TargetDay = ''
                )

                $lines = New-Object System.Collections.Generic.List[object]
                foreach ($r in $Rows) {
                                $etd = Parse-ETD $r.ETD
                                if ($null -eq $etd) { continue }
                                if ($TargetDay -ne '') {
                                                $etd = ([datetime]::ParseExact($TargetDay, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)).Add($etd.TimeOfDay)
                                }
                                $type = [string]$r.SortieType
                                if ([string]::IsNullOrWhiteSpace($type)) { continue }
                                $ip = Get-NameText $r.Instructor
                                if ($NoInstructors -or $ip -like '*SIM IP*' -or $ip -eq 'NEED IP') { $ip = '' }
                                $student = ''
                                $mission = ''
                                if ($KeepStudents -and -not (Test-IsSentinelStudent $r.Student) -and -not (Test-IsReturnSortieType $type)) {
                                                $student = Get-NameText $r.Student
                                                $mission = Get-NameText $r.Mission
                                }
                                $notes = if ($KeepNotes) { [string]$r.Notes } else { '' }
                                $callsign = if ($NoInstructors) { '' } else { [string](Get-OptionalProperty $r 'Callsign' '') }
                                $num = if ($NoInstructors) { '' } else { [string](Get-OptionalProperty $r 'Num' '') }
                                if (-not $KeepStudents) { $callsign = ''; $num = '' }
                                $lines.Add((New-BoardLine -ETD $etd -SortieType $type -Instructor $ip -Student $student -Mission $mission -Callsign $callsign -Num $num -Notes $notes -Source $Source))
                }

                # Link each O&B to the next unlinked return that day with the same IP (any IP if blank)
                $sorted = @($lines | Sort-Object -Property ETD)
                foreach ($ob in $sorted) {
                                if (-not (Test-BoardLineIsOB $ob)) { continue }
                                $ret = $sorted | Where-Object {
                                                (Test-BoardLineIsReturn $_) -and -not $_.LinkedLineId -and $_.Day -eq $ob.Day -and $_.ETD -gt $ob.ETD -and
                                                ($ob.Instructor -eq '' -or $_.Instructor -eq '' -or $_.Instructor -ieq $ob.Instructor)
                                } | Select-Object -First 1
                                if ($null -ne $ret) {
                                                $ob.LinkedLineId = $ret.Id
                                                $ret.LinkedLineId = $ob.Id
                                }
                }
                return $lines.ToArray()
}

# What the fill put on each board line: lineId -> {Student, Mission, FilledBy, Notes, Callsign, Num, Instructor}
function Get-FillOverlay {
                param ($FillResult)

                $map = @{}
                foreach ($r in @($FillResult.Schedule)) {
                                $id = [string](Get-OptionalProperty $r 'BoardLineId' '')
                                if ($id -eq '') { continue }
                                $map[$id] = [PSCustomObject]@{
                                                Student    = Get-NameText $r.Student
                                                Mission    = Get-NameText $r.Mission
                                                FilledBy   = [string]$r.FilledBy
                                                Notes      = [string]$r.Notes
                                                Callsign   = [string]$r.Callsign
                                                Num        = [string]$r.Num
                                                Instructor = Get-NameText $r.Instructor
                                                Override   = ([string]$r.Notes -like '*OVERRIDE:*')
                                }
                }
                return $map
}

# Things worth a second look before filling. Returns a list of strings.
function Test-BoardReadyToFill {
                param (
                                $Board,
                                [array]$Students = @(),
                                [array]$Instructors = @()
                )

                $w = New-Object System.Collections.Generic.List[string]
                if ($Board.Lines.Count -eq 0) { $w.Add('There are no lines on the board.'); return $w.ToArray() }

                $seen = @{}
                foreach ($l in (Get-BoardDayLines -Board $Board -Day '')) {
                                $label = "$($l.ETD.ToString('ddd HHmm')) $($l.SortieType)"
                                $isSim = Test-IsSimSortieType ([string]$l.SortieType)
                                if (-not $isSim -and [string]::IsNullOrWhiteSpace($l.Instructor)) { $w.Add("$label has no IP (it will be NEED IP)") }
                                elseif (-not $isSim -and $null -eq (Get-MatchedInstructor -Instructors $Instructors -TargetName $l.Instructor)) { $w.Add("$label IP '$($l.Instructor)' is not in MasterData") }
                                if ((Test-BoardLineIsOB $l) -and -not $l.LinkedLineId) { $w.Add("$label has no linked O&B Return line") }
                                if (-not [string]::IsNullOrWhiteSpace($l.Student)) {
                                                if ($null -eq (Resolve-StudentByName -Students $Students -Name $l.Student)) { $w.Add("$label student '$($l.Student)' is not in MasterData") }
                                                $key = "$($l.Student)|$($l.ETD.ToString('s'))"
                                                if ($seen.ContainsKey($key)) { $w.Add("$($l.Student) is pinned twice at $($l.ETD.ToString('ddd HHmm'))") }
                                                $seen[$key] = $true
                                }
                }
                foreach ($d in $Board.Days) {
                                if (@($Board.Lines | Where-Object { $_.Day -eq $d }).Count -eq 0) { $w.Add("$d has no lines") }
                }
                return $w.ToArray()
}


# =================================================================================
#  FILES: AUTOSAVE AND TEMPLATES
#  Dates are stored as text so Windows PowerShell 5.1 round-trips them exactly.
# =================================================================================

function ConvertTo-LineDto {
                param ($Line)

                return [ordered]@{
                                Id           = $Line.Id
                                ETD          = $Line.ETD.ToString('yyyy-MM-ddTHH:mm')
                                SortieType   = [string]$Line.SortieType
                                Instructor   = [string]$Line.Instructor
                                Student      = [string]$Line.Student
                                Mission      = [string]$Line.Mission
                                Callsign     = [string]$Line.Callsign
                                Num          = [string]$Line.Num
                                Notes        = [string]$Line.Notes
                                LinkedLineId = [string]$Line.LinkedLineId
                                WingOfLineId = [string]$Line.WingOfLineId
                                Source       = [string]$Line.Source
                }
}

function ConvertFrom-LineDto {
                param ($Dto)

                $etd = [datetime]::ParseExact([string]$Dto.ETD, 'yyyy-MM-ddTHH:mm', [System.Globalization.CultureInfo]::InvariantCulture)
                return (New-BoardLine -ETD $etd -SortieType ([string]$Dto.SortieType) -Instructor ([string]$Dto.Instructor) -Student ([string]$Dto.Student) `
                                -Mission ([string]$Dto.Mission) -Callsign ([string]$Dto.Callsign) -Num ([string]$Dto.Num) -Notes ([string]$Dto.Notes) `
                                -Source ([string]$Dto.Source) -Id ([string]$Dto.Id) -LinkedLineId ([string]$Dto.LinkedLineId) -WingOfLineId ([string]$Dto.WingOfLineId))
}

# Write-then-rename so a crash never leaves a half-written file
function Write-TextFileAtomic {
                param ([string]$Path, [string]$Text)

                $dir = Split-Path -Parent $Path
                if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
                $tmp = "$Path.tmp"
                [System.IO.File]::WriteAllText($tmp, $Text, (New-Object System.Text.UTF8Encoding($false)))
                Move-Item -LiteralPath $tmp -Destination $Path -Force
}

# One file per day: <Dir>\yyyy-MM-dd.json
function Save-BoardState {
                param ($Board, [Parameter(Mandatory)] [string]$Dir)

                foreach ($day in $Board.Days) {
                                $dto = [ordered]@{
                                                SchemaVersion = 1
                                                Day           = $day
                                                SavedAt       = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
                                                Lines         = @(foreach ($l in (Get-BoardDayLines -Board $Board -Day $day)) { ConvertTo-LineDto $l })
                                }
                                Write-TextFileAtomic -Path (Join-Path $Dir "$day.json") -Text (ConvertTo-Json -InputObject $dto -Depth 6)
                }
}

function Read-BoardDayFile {
                param ([string]$Path)

                if (-not (Test-Path -LiteralPath $Path)) { return }
                try {
                                $dto = ConvertFrom-Json ([System.IO.File]::ReadAllText($Path))
                                return @(foreach ($l in @($dto.Lines)) { if ($null -ne $l) { ConvertFrom-LineDto $l } })
                } catch {
                                Write-Warning "Could not read saved board $Path : $($_.Exception.Message)"
                }
}

# Days on/after -From that have a saved board file
function Get-SavedBoardDays {
                param ([string]$Dir, [datetime]$From = [datetime]::MinValue)

                if (-not (Test-Path -LiteralPath $Dir)) { return }
                $out = foreach ($f in (Get-ChildItem -LiteralPath $Dir -Filter '*.json' -File)) {
                                $d = $null
                                if ([datetime]::TryParseExact($f.BaseName, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture, 'None', [ref]$d) -and $d.Date -ge $From.Date) { $f.BaseName }
                }
                return @($out | Sort-Object)
}

# Board for -Days, filled from saved files where they exist
function Open-BoardState {
                param ([string]$Dir, [string[]]$Days, [hashtable]$Settings)

                $board = New-Board -Days $Days -Settings $Settings
                foreach ($day in $board.Days) {
                                foreach ($l in @(Read-BoardDayFile (Join-Path $Dir "$day.json"))) { $board.Lines.Add($l) }
                }
                return $board
}

# Saves a day's lines as a named template (times and types; IPs optional).
# Return lines are stored inside their O&B as ReturnTime.
function Save-BoardTemplate {
                param (
                                $Board,
                                [string]$Day,
                                [Parameter(Mandatory)] [string]$Name,
                                [Parameter(Mandatory)] [string]$Dir,
                                [switch]$IncludeInstructors
                )

                $safe = ($Name -replace '[^\w\- ]', '').Trim()
                if ($safe -eq '') { throw "Template name '$Name' has no usable characters" }
                $items = foreach ($l in (Get-BoardDayLines -Board $Board -Day $Day)) {
                                if ((Test-BoardLineIsReturn $l) -and $l.LinkedLineId) { continue }
                                $rt = ''
                                if ((Test-BoardLineIsOB $l) -and $l.LinkedLineId) {
                                                $ret = Get-BoardLine $Board $l.LinkedLineId
                                                if ($null -ne $ret) { $rt = $ret.ETD.ToString('HH:mm') }
                                }
                                [ordered]@{
                                                Time       = $l.ETD.ToString('HH:mm')
                                                SortieType = [string]$l.SortieType
                                                Instructor = $(if ($IncludeInstructors) { [string]$l.Instructor } else { '' })
                                                ReturnTime = $rt
                                }
                }
                $dto = [ordered]@{ SchemaVersion = 1; Name = $Name; SavedAt = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss'); Lines = @($items) }
                $path = Join-Path $Dir "$safe.json"
                Write-TextFileAtomic -Path $path -Text (ConvertTo-Json -InputObject $dto -Depth 5)
                return $path
}

function Get-BoardTemplateList {
                param ([string]$Dir)

                if (-not (Test-Path -LiteralPath $Dir)) { return }
                $out = foreach ($f in (Get-ChildItem -LiteralPath $Dir -Filter '*.json' -File)) {
                                try {
                                                $dto = ConvertFrom-Json ([System.IO.File]::ReadAllText($f.FullName))
                                                [PSCustomObject]@{ Name = [string]$dto.Name; Path = $f.FullName; Count = @($dto.Lines).Count }
                                } catch { }
                }
                return @($out | Sort-Object -Property Name)
}

# Adds a template's lines to a day (one undo step). -Replace clears the day first.
function Import-BoardTemplate {
                param (
                                $Board,
                                [string]$Day,
                                [Parameter(Mandatory)] [string]$Path,
                                [switch]$Replace
                )

                $dto = ConvertFrom-Json ([System.IO.File]::ReadAllText($Path))
                $dayDate = [datetime]::ParseExact($Day, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
                Save-BoardUndoPoint $Board
                if ($Replace) {
                                foreach ($l in @($Board.Lines | Where-Object { $_.Day -eq $Day })) { [void]$Board.Lines.Remove($l) }
                }
                $added = 0
                foreach ($t in @($dto.Lines)) {
                                $time = ConvertTo-BoardTime ([string]$t.Time)
                                if ($null -eq $time) { continue }
                                $line = New-BoardLine -ETD $dayDate.Add($time) -SortieType ([string]$t.SortieType) -Instructor ([string]$t.Instructor) -Source 'Template'
                                $Board.Lines.Add($line); $added++
                                if (Test-BoardLineIsOB $line) {
                                                $rt = ConvertTo-BoardTime ([string]$t.ReturnTime)
                                                $retEtd = if ($null -ne $rt) { $dayDate.Add($rt) } else { $null }
                                                $Board.Lines.Add((New-ReturnLineFor -Board $Board -OBLine $line -ReturnETD $retEtd)); $added++
                                }
                }
                if (-not $Board.Days.Contains($Day)) { $Board.Days.Add($Day); $Board.Days.Sort() }
                Step-BoardVersion $Board
                return $added
}

# Copies the most recent earlier day's lines (times, types, IPs) onto -TargetDay.
# Looks at saved boards first (up to 7 days back, so Monday finds Friday), then the
# previous MasterSchedule. Students are not copied. Returns a short description.
function Copy-BoardPreviousDay {
                param (
                                $Board,
                                [Parameter(Mandatory)] [string]$TargetDay,
                                [string]$Dir,
                                [array]$PrevSchedule = @(),
                                [switch]$IncludeInstructors
                )

                $target = [datetime]::ParseExact($TargetDay, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
                $sourceRows = $null
                $sourceName = ''

                # 1. Lines already on the board for an earlier day
                for ($back = 1; $back -le 7 -and $null -eq $sourceRows; $back++) {
                                $d = $target.AddDays(-$back).ToString('yyyy-MM-dd')
                                $onBoard = @($Board.Lines | Where-Object { $_.Day -eq $d })
                                if ($onBoard.Count -gt 0) {
                                                $sourceRows = @(ConvertTo-ScheduleRows -Board $Board -Days @($d))
                                                $sourceName = "board $d"
                                                break
                                }
                                if ($Dir) {
                                                $saved = @(Read-BoardDayFile (Join-Path $Dir "$d.json"))
                                                if ($saved.Count -gt 0) {
                                                                $tmp = New-Board -Days @($d)
                                                                foreach ($l in $saved) { $tmp.Lines.Add($l) }
                                                                $sourceRows = @(ConvertTo-ScheduleRows -Board $tmp -Days @($d))
                                                                $sourceName = "saved board $d"
                                                }
                                }
                }

                # 2. The previous MasterSchedule's latest day before the target
                if ($null -eq $sourceRows -and $PrevSchedule.Count -gt 0) {
                                $days = @($PrevSchedule | ForEach-Object { $x = Parse-ETD $_.ETD; if ($null -ne $x -and $x.Date -lt $target.Date) { $x.Date } } | Sort-Object -Unique)
                                if ($days.Count -gt 0) {
                                                $last = $days[$days.Count - 1]
                                                $sourceRows = @($PrevSchedule | Where-Object { $x = Parse-ETD $_.ETD; $null -ne $x -and $x.Date -eq $last })
                                                $sourceName = "MasterSchedule $($last.ToString('yyyy-MM-dd'))"
                                }
                }

                if ($null -eq $sourceRows -or $sourceRows.Count -eq 0) { return $null }

                $lines = @(ConvertFrom-ScheduleRows -Rows $sourceRows -Source 'Copied' -NoInstructors:(-not $IncludeInstructors) -TargetDay $TargetDay)
                Add-BoardLine -Board $Board -Line $lines
                return [PSCustomObject]@{ Source = $sourceName; Count = $lines.Count }
}

# Next Monday-Friday after a date
function Get-NextFlyingDay {
                param ([datetime]$From = (Get-Date))

                $d = $From.Date.AddDays(1)
                while ($d.DayOfWeek -eq 'Saturday' -or $d.DayOfWeek -eq 'Sunday') { $d = $d.AddDays(1) }
                return $d
}
