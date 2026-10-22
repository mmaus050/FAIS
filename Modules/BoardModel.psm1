# BOARDMODEL.PSM1 DIRECTORY:
#             The puck board's data: days, lines, pins. No WPF in here, so all of it can
#             be tested without a window.
# Get-DefaultBoardSettings / New-Board / New-BoardLine / Get-BoardLine / Get-BoardDayLines
# Undo: Save-BoardUndoPoint / Undo-Board
# Edits: Add-BoardDay / Remove-BoardDay / Add-BoardLine / Remove-BoardLine / Set-BoardLineEtd
#        Set-BoardLineType / Set-BoardInstructor / Set-BoardPin / Clear-BoardPin / Move-BoardPin
#        Set-BoardMission / Set-BoardLineField / Set-FormationPair / Unlink-ReturnLine
# Typing: ConvertTo-BoardTime / Format-BoardTime / ConvertTo-BoardSortieType / ConvertFrom-QuickEntry
#         New-BulkBoardLines (Add Bulk Lines)
# Engine: ConvertTo-ScheduleRows / ConvertFrom-ScheduleRows / Get-FillOverlay / Test-BoardReadyToFill
# One day at a time: Resolve-SheetRowDays / ConvertFrom-SheetRows / Import-BoardDayLines / Get-BoardSheetState / Get-SheetBannerInfo / Get-SheetDaysToAdd
#         Set-BoardSheetSync / ConvertFrom-PublishedRows / Get-PublishedContextRows / Split-FillResult
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
                                # Day -> what MasterData's Schedule sheet held for that day when it was last loaded
                                # or ignored: {Fingerprint; Action 'Loaded'|'Ignored'; At}
                                SheetSync     = @{}
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
                                Sync  = @(foreach ($d in @(Get-BoardSheetSyncDays $Board)) { [ordered]@{ Day = $d; Entry = $Board.SheetSync[$d] } })
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
                if ($null -ne $Board.PSObject.Properties['SheetSync']) {
                                $Board.SheetSync = @{}
                                foreach ($e in @($dto.Sync)) { if ($null -ne $e -and $e.Day) { $Board.SheetSync[[string]$e.Day] = ConvertTo-SheetSyncEntry $e.Entry } }
                }
}

# Days with a sheet sync record, in order (empty for boards made before sync existed)
function Get-BoardSheetSyncDays {
                param ($Board)

                if ($null -eq $Board.PSObject.Properties['SheetSync'] -or $null -eq $Board.SheetSync) { return @() }
                return @($Board.SheetSync.Keys | Sort-Object)
}

function ConvertTo-SheetSyncEntry {
                param ($Entry)

                if ($null -eq $Entry) { return $null }
                return [PSCustomObject]@{
                                Fingerprint      = [string]$Entry.Fingerprint
                                Action           = [string]$Entry.Action
                                At               = [string]$Entry.At
                                BoardFingerprint = [string](Get-OptionalProperty $Entry 'BoardFingerprint' '')
                }
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
                if (@(Get-BoardSheetSyncDays $Board) -contains $Day) { $Board.SheetSync.Remove($Day) }
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
                                # O&B -> return links a quick-entry parse found for lines already on the board.
                                # Applied here, after the undo point, so a failed entry changes nothing and undo reverses them.
                                [array]$Link = @(),
                                [switch]$NoUndo
                )

                if (-not $NoUndo) { Save-BoardUndoPoint $Board }
                foreach ($l in $Line) {
                                if ($null -eq $l) { continue }
                                if (-not $Board.Days.Contains($l.Day)) { $Board.Days.Add($l.Day); $Board.Days.Sort() }
                                $Board.Lines.Add($l)
                }
                foreach ($k in $Link) {
                                if ($null -eq $k) { continue }
                                $ob = Get-BoardLine $Board ([string]$k.ObId)
                                if ($null -ne $ob -and -not $ob.LinkedLineId) { $ob.LinkedLineId = [string]$k.ReturnId }
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
                                                if (-not $Board.Days.Contains($ret.Day)) { $Board.Days.Add($ret.Day); $Board.Days.Sort() }
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
                $wasReturn = Test-BoardLineIsReturn $line
                $line.SortieType = $SortieType
                $isOB = Test-BoardLineIsOB $line

                # A return line retyped to something else is no longer anyone's return leg
                if ($wasReturn -and -not (Test-BoardLineIsReturn $line) -and $line.LinkedLineId) {
                                $ob = Get-BoardLine $Board $line.LinkedLineId
                                if ($null -ne $ob) { $ob.LinkedLineId = '' }
                                $line.LinkedLineId = ''
                }
                # OFT / IFT / O&B / Return lines can never fly formation: drop their pairings both ways
                if (-not (Test-IsWingCapableType ([string]$line.SortieType))) {
                                $line.WingOfLineId = ''
                                foreach ($l in $Board.Lines.ToArray()) { if ($l.WingOfLineId -eq $line.Id) { $l.WingOfLineId = '' } }
                }

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

# Moves an IP from one line to another (swaps if the target has one). -Copy keeps the original.
function Move-BoardInstructor {
                param ($Board, [string]$FromId, [string]$ToId, [switch]$Copy)

                $from = Get-BoardLine $Board $FromId
                $to = Get-BoardLine $Board $ToId
                if ($null -eq $from -or $null -eq $to -or $FromId -eq $ToId) { return $false }
                Save-BoardUndoPoint $Board

                $moving = $from.Instructor
                $displaced = $to.Instructor
                $to.Instructor = $moving
                if (-not $Copy) { $from.Instructor = $displaced }
                foreach ($l in @($from, $to)) {
                                if ((Test-BoardLineIsOB $l) -and $l.LinkedLineId) {
                                                $ret = Get-BoardLine $Board $l.LinkedLineId
                                                if ($null -ne $ret) { $ret.Instructor = $l.Instructor }
                                }
                }
                Step-BoardVersion $Board
                return $true
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
                if ($null -eq $line -or [string]$line.Mission -ceq $Mission) { return }
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
                                # OFT / IFT / O&B lines can never fly formation, as lead or wing
                                $lead = Get-BoardLine $Board $LeadId
                                if ($null -eq $lead -or $LeadId -eq $WingId) { return $false }
                                if (-not (Test-IsWingCapableType ([string]$wing.SortieType)) -or -not (Test-IsWingCapableType ([string]$lead.SortieType))) { return $false }
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
                $links = New-Object System.Collections.Generic.List[object]
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
                                                                $claimed = @($links | ForEach-Object { $_.ObId })
                                                                $ob = Find-UnlinkedOBForReturn -Board $Board -Pending $lines -ReturnLine $line -ClaimedIds $claimed
                                                                if ($null -ne $ob) {
                                                                                # Lines from this same entry can be linked now; board lines are linked by Add-BoardLine
                                                                                if ($lines.Contains($ob)) { $ob.LinkedLineId = $line.Id }
                                                                                $links.Add(@{ ObId = $ob.Id; ReturnId = $line.Id })
                                                                                $line.LinkedLineId = $ob.Id
                                                                                if ($line.Instructor -eq '') { $line.Instructor = $ob.Instructor }
                                                                }
                                                                else { $warnings.Add("No O&B without a return before $(Format-BoardTime $etd); return line added unlinked") }
                                                }
                                }
                }

                return [PSCustomObject]@{
                                Lines    = $lines.ToArray()
                                Links    = $links.ToArray()
                                Errors   = $errors.ToArray()
                                Warnings = $warnings.ToArray()
                }
}

# Add Bulk Lines: -Count lines of -SortieType on -Day, the first at -StartTime and each one
# -Every minutes after the one before (3 by default). Values come straight from the dialog's
# boxes, so they are checked here. An O&B gets its linked return leg, like a typed one.
# Returns {Lines; Times ('HHmm' of each line); Errors}. Lines is empty when there are errors.
function New-BulkBoardLines {
                param (
                                [Parameter(Mandatory)] $Day,
                                [string]$StartTime,
                                [string]$Count,
                                [string]$Every = '3',
                                [string]$SortieType = 'Day Local',
                                [string]$Instructor = '',
                                $Settings
                )

                $defaults = Get-DefaultBoardSettings
                if ($null -eq $Settings) { $Settings = $defaults }
                $errors = New-Object System.Collections.Generic.List[string]
                $lines = New-Object System.Collections.Generic.List[object]
                $times = New-Object System.Collections.Generic.List[string]
                $maxLines = 60

                $start = ConvertTo-BoardTime $StartTime
                if ($null -eq $start) { $errors.Add("Start time '$StartTime' isn't a time (type 0730, 730 or 7:30).") }
                $n = 0
                if (-not [int]::TryParse(([string]$Count).Trim(), [ref]$n) -or $n -lt 1 -or $n -gt $maxLines) { $errors.Add("Number of lines must be a whole number from 1 to $maxLines.") }
                $gap = 0
                if (-not [int]::TryParse(([string]$Every).Trim(), [ref]$gap) -or $gap -lt 0 -or $gap -gt 240) { $errors.Add('Minutes between lines must be a whole number from 0 to 240.') }
                $type = ([string]$SortieType).Trim()
                if ($type -eq '') { $errors.Add('Pick a sortie type.') }
                $dayText = ConvertTo-BoardDay $Day
                if ($null -eq $dayText) { $errors.Add("'$Day' isn't a day.") }

                if ($errors.Count -eq 0) {
                                $dayDate = [datetime]::ParseExact($dayText, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
                                $first = $dayDate.Add($start)
                                $last = $first.AddMinutes($gap * ($n - 1))
                                $offset = [int]$Settings.OBReturnOffsetMinutes
                                if ($offset -le 0) { $offset = [int]$defaults.OBReturnOffsetMinutes }
                                $returnType = [string]$Settings.ReturnSortieType
                                if ($returnType.Trim() -eq '') { $returnType = [string]$defaults.ReturnSortieType }
                                $returnBoard = [PSCustomObject]@{ Settings = @{ OBReturnOffsetMinutes = $offset; ReturnSortieType = $returnType } }
                                $isOB = Test-IsOBSortieType $type
                                if ($last.Date -ne $dayDate) {
                                                $errors.Add(("The last line would be at {0} the next day. Start earlier, add fewer lines or put them closer together." -f $last.ToString('HHmm')))
                                }
                                elseif ($isOB -and $last.AddMinutes($offset).Date -ne $dayDate) {
                                                $errors.Add(("The last O&B's return leg would be at {0} the next day. Start earlier or add fewer lines." -f $last.AddMinutes($offset).ToString('HHmm')))
                                }
                                else {
                                                for ($i = 0; $i -lt $n; $i++) {
                                                                $etd = $first.AddMinutes($gap * $i)
                                                                $line = New-BoardLine -ETD $etd -SortieType $type -Instructor $Instructor -Source 'Bulk'
                                                                $lines.Add($line)
                                                                $times.Add($etd.ToString('HHmm'))
                                                                if ($isOB) { $lines.Add((New-ReturnLineFor -Board $returnBoard -OBLine $line -ReturnETD $etd.AddMinutes($offset))) }
                                                }
                                }
                }
                return [PSCustomObject]@{
                                Lines  = $(if ($errors.Count -eq 0) { $lines.ToArray() } else { @() })
                                Times  = $(if ($errors.Count -eq 0) { $times.ToArray() } else { @() })
                                Errors = $errors.ToArray()
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
                param ($Board, $Pending, $ReturnLine, [array]$ClaimedIds = @())

                $pool = New-Object System.Collections.Generic.List[object]
                if ($null -ne $Board) { foreach ($l in $Board.Lines) { $pool.Add($l) } }
                foreach ($l in $Pending) { if ($l.Id -ne $ReturnLine.Id) { $pool.Add($l) } }
                $cands = @($pool | Where-Object {
                                (Test-BoardLineIsOB $_) -and -not $_.LinkedLineId -and $ClaimedIds -notcontains $_.Id -and $_.Day -eq $ReturnLine.Day -and $_.ETD -lt $ReturnLine.ETD
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
                $newIdByOldId = @{}
                $oldLinkByNewId = @{}
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
                                # "Clears Unsat ..." is written by the fill, so it never comes in from a sheet
                                $notes = if ($KeepNotes) { (@(([string]$r.Notes) -split '\s*\|\s*' | Where-Object { $_ -ne '' -and $_ -notlike 'Clears Unsat*' })) -join ' | ' } else { '' }
                                $callsign = if ($NoInstructors) { '' } else { [string](Get-OptionalProperty $r 'Callsign' '') }
                                $num = if ($NoInstructors) { '' } else { [string](Get-OptionalProperty $r 'Num' '') }
                                if (-not $KeepStudents) { $callsign = ''; $num = '' }
                                $newLine = New-BoardLine -ETD $etd -SortieType $type -Instructor $ip -Student $student -Mission $mission -Callsign $callsign -Num $num -Notes $notes -Source $Source
                                $lines.Add($newLine)
                                $oldId = [string](Get-OptionalProperty $r 'BoardLineId' '')
                                $oldLink = [string](Get-OptionalProperty $r 'LinkedLineId' '')
                                if ($oldId -ne '') { $newIdByOldId[$oldId] = $newLine }
                                if ($oldLink -ne '') { $oldLinkByNewId[$newLine.Id] = $oldLink }
                }

                # Rows that carry their own link ids (from the board) keep their exact O&B / return pairs
                foreach ($l in $lines) {
                                if (-not $oldLinkByNewId.ContainsKey($l.Id) -or $l.LinkedLineId) { continue }
                                $other = $newIdByOldId[$oldLinkByNewId[$l.Id]]
                                if ($null -ne $other -and -not $other.LinkedLineId -and $other.Id -ne $l.Id -and
                                                (((Test-BoardLineIsOB $l) -and (Test-BoardLineIsReturn $other)) -or ((Test-BoardLineIsReturn $l) -and (Test-BoardLineIsOB $other)))) {
                                                $l.LinkedLineId = $other.Id
                                                $other.LinkedLineId = $l.Id
                                }
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
# -Days limits the checks to the days being filled (default: every board day).
function Test-BoardReadyToFill {
                param (
                                $Board,
                                [array]$Students = @(),
                                [array]$Instructors = @(),
                                [string[]]$Days = @()
                )

                $w = New-Object System.Collections.Generic.List[string]
                if ($Board.Lines.Count -eq 0) { $w.Add('There are no lines on the board.'); return $w.ToArray() }
                $checkDays = if ($Days.Count -gt 0) { $Days } else { @($Board.Days) }

                $seen = @{}
                foreach ($l in (Get-BoardDayLines -Board $Board -Day '')) {
                                if ($checkDays -notcontains $l.Day) { continue }
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
                foreach ($d in $checkDays) {
                                if (@($Board.Lines | Where-Object { $_.Day -eq $d }).Count -eq 0) { $w.Add("$d has no lines") }
                }
                return $w.ToArray()
}


# =================================================================================
#  ONE DAY AT A TIME
#  A day can be loaded from MasterData's Schedule sheet or from the last exported
#  schedule (OutputSchedule's MasterSchedule), filled on its own and exported on its own.
#  Exported days the fill does not touch go in as fixed context so nobody is double-booked.
# =================================================================================

function Format-BoardDayLabel {
                param ([string]$Day)

                if ([string]::IsNullOrWhiteSpace($Day)) { return 'no date' }
                return ([datetime]::ParseExact($Day, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)).ToString('ddd dd MMM')
}

# Sheet rows that have a time but no date belong to the day being built (-DefaultDay, the
# board's first day). Returns copies of those rows dated that day; dated rows are unchanged.
function Resolve-SheetRowDays {
                param ([array]$Rows = @(), [string]$DefaultDay = '')

                return @(foreach ($r in $Rows) {
                                if ($null -eq $r) { continue }
                                if ([string]$r.Day -ne '' -or $DefaultDay -eq '') { $r; continue }
                                $c = $r.PSObject.Copy()
                                $c.Day = $DefaultDay
                                $c.ETD = "$DefaultDay $($r.Time):00"
                                $c
                })
}

# The Schedule sheet's lines for one day as board lines (Read-ScheduleSheet rows).
# -FromDay takes another day's rows ('' = the rows with no date) and puts them on -Day.
# -IncludeTimeOnly also adds the rows that have no date. -KeepStudents hardcodes the
# students (and missions) the sheet names.
function ConvertFrom-SheetRows {
                param (
                                [array]$Rows = @(),
                                [Parameter(Mandatory)] [string]$Day,
                                [AllowEmptyString()] [string]$FromDay,
                                [switch]$IncludeTimeOnly,
                                [switch]$KeepStudents,
                                [switch]$KeepNotes
                )

                if (-not $PSBoundParameters.ContainsKey('FromDay')) { $FromDay = $Day }
                $pick = @(foreach ($r in $Rows) {
                                $match = if ($FromDay -eq '') { [bool]$r.TimeOnly -and [string]$r.Day -eq '' } else { [string]$r.Day -eq $FromDay }
                                if (-not $match -and -not ($IncludeTimeOnly -and [string]$r.Day -eq '')) { continue }
                                $c = $r.PSObject.Copy()
                                $c.ETD = "$Day $($r.Time):00"
                                $c
                })
                return @(ConvertFrom-ScheduleRows -Rows $pick -Source 'Sheet' -KeepStudents:$KeepStudents -KeepNotes:$KeepNotes)
}

function New-SheetSyncEntry {
                param ([string]$Fingerprint, [ValidateSet('Loaded', 'Ignored')] [string]$Action = 'Loaded', [string]$BoardFingerprint = '')

                return [PSCustomObject]@{ Fingerprint = $Fingerprint; Action = $Action; At = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss'); BoardFingerprint = $BoardFingerprint }
}

# A board day's lines as a short text that changes with any edit to them (times, types, IPs,
# students, missions, callsigns, notes, O&B and wing links). Line ids do not count.
function Get-BoardDayFingerprint {
                param ($Board, [string]$Day)

                $parts = foreach ($l in (Get-BoardDayLines -Board $Board -Day $Day)) {
                                $link = ''; $wing = ''
                                if ($l.LinkedLineId) { $o = Get-BoardLine $Board $l.LinkedLineId; $link = if ($null -ne $o) { $o.ETD.ToString('HHmm') } else { '?' } }
                                if ($l.WingOfLineId) { $o = Get-BoardLine $Board $l.WingOfLineId; $wing = if ($null -ne $o) { $o.ETD.ToString('HHmm') + [string]$o.Student } else { '?' } }
                                (@($l.ETD.ToString('HHmm'), $l.SortieType, $l.Instructor, $l.Student, $l.Mission, $l.Callsign, $l.Num, $l.Notes, $link, $wing) | ForEach-Object { ([string]$_).Trim() }) -join '|'
                }
                return (Get-FaisFingerprint -Parts @($parts))
}

# Records what the sheet held for a day when the user loaded it or chose to keep the board's lines
function Set-BoardSheetSync {
                param ($Board, [Parameter(Mandatory)] [string]$Day, $Entry)

                if ($null -eq $Board.PSObject.Properties['SheetSync']) { $Board | Add-Member -MemberType NoteProperty -Name 'SheetSync' -Value @{} }
                if ($null -eq $Entry) { $Board.SheetSync.Remove($Day) } else { $Board.SheetSync[$Day] = ConvertTo-SheetSyncEntry $Entry }
}

# Puts one day's lines on the board as one undo step. -Replace clears the day first.
# -Sync records what the Schedule sheet held for that day. Returns {Added; Removed}.
function Import-BoardDayLines {
                param (
                                $Board,
                                [Parameter(Mandatory)] [string]$Day,
                                [array]$Lines = @(),
                                [switch]$Replace,
                                $Sync
                )

                Save-BoardUndoPoint $Board
                $removed = 0
                if ($Replace) {
                                $gone = @($Board.Lines | Where-Object { $_.Day -eq $Day } | ForEach-Object { $_.Id })
                                foreach ($l in $Board.Lines.ToArray()) {
                                                if ($gone -contains $l.Id) { [void]$Board.Lines.Remove($l); $removed++; continue }
                                                if ($gone -contains $l.LinkedLineId) { $l.LinkedLineId = '' }
                                                if ($gone -contains $l.WingOfLineId) { $l.WingOfLineId = '' }
                                }
                }
                if (-not $Board.Days.Contains($Day)) { $Board.Days.Add($Day); $Board.Days.Sort() }
                $added = 0
                foreach ($l in $Lines) {
                                if ($null -eq $l) { continue }
                                if (-not $Board.Days.Contains($l.Day)) { $Board.Days.Add($l.Day); $Board.Days.Sort() }
                                $Board.Lines.Add($l)
                                $added++
                }
                if ($null -ne $Sync) {
                                # Remember how the day looked right after loading, to tell later whether it was edited
                                $entry = ConvertTo-SheetSyncEntry $Sync
                                $entry.BoardFingerprint = Get-BoardDayFingerprint -Board $Board -Day $Day
                                Set-BoardSheetSync -Board $Board -Day $Day -Entry $entry
                }
                Step-BoardVersion $Board
                return [PSCustomObject]@{ Added = $added; Removed = $removed }
}

# How a board day stands against MasterData's Schedule sheet (Read-ScheduleSheet rows):
#   NoSheetLines  the sheet has no dated lines for the day
#   InSync        the board last loaded (or chose to keep its own lines over) exactly what the sheet has now
#   EmptyDay      the board has no lines that day and the sheet does
#   Outdated      the sheet changed since the day was loaded and nobody has touched the day on
#                 the board since, so it can be updated without losing anything
#   Changed       the sheet changed since the day was loaded or kept, and the board's day was edited
#   New           the sheet has lines for a day the board built some other way
# Returns {Day; State; SheetCount; BoardCount; Fingerprint}
function Get-BoardSheetState {
                param ($Board, [array]$SheetRows = @(), [Parameter(Mandatory)] [string]$Day)

                $count = @($SheetRows | Where-Object { [string]$_.Day -eq $Day }).Count
                $fp = Get-ScheduleDayFingerprint -Rows $SheetRows -Day $Day
                $boardCount = @($Board.Lines | Where-Object { $_.Day -eq $Day }).Count
                $entry = $null
                if (@(Get-BoardSheetSyncDays $Board) -contains $Day) { $entry = $Board.SheetSync[$Day] }

                $state = 'New'
                if ($count -eq 0) { $state = 'NoSheetLines' }
                elseif ($null -ne $entry -and [string]$entry.Fingerprint -eq $fp) { $state = 'InSync' }
                elseif ($boardCount -eq 0) { $state = 'EmptyDay' }
                elseif ($null -ne $entry -and $entry.Action -eq 'Loaded' -and $entry.BoardFingerprint -ne '' -and $entry.BoardFingerprint -eq (Get-BoardDayFingerprint -Board $Board -Day $Day)) { $state = 'Outdated' }
                elseif ($null -ne $entry) { $state = 'Changed' }
                return [PSCustomObject]@{ Day = $Day; State = $state; SheetCount = $count; BoardCount = $boardCount; Fingerprint = $fp }
}

# The one thing about the Schedule sheet worth telling the user, or $null. First a board day
# (the active one first) whose sheet lines changed or were never loaded, else a later day the
# sheet has lines for that is not on the board yet, else (PastDay) why an empty board got
# nothing: the sheet only has days before the board's first day. -Ignore holds
# 'day|fingerprint' keys of days the user already waved off.
# Returns {Day; TargetDay (where loading puts it); State; Count; Fingerprint; Text; ActionText}
function Get-SheetBannerInfo {
                param ($Board, [array]$SheetRows = @(), [string]$ActiveDay = '', [string[]]$Ignore = @())

                if (@($SheetRows).Count -eq 0) { return $null }
                $days = @($Board.Days)
                if ($ActiveDay -and $days -contains $ActiveDay) { $days = @($ActiveDay) + @($days | Where-Object { $_ -ne $ActiveDay }) }
                foreach ($d in $days) {
                                $st = Get-BoardSheetState -Board $Board -SheetRows $SheetRows -Day $d
                                $label = Format-BoardDayLabel $d
                                $text = $null; $action = $null
                                switch ($st.State) {
                                                'Outdated' { $text = "MasterData's Schedule sheet changed for $label ($($st.SheetCount) line(s) there now)."; $action = "Update $label from the sheet" }
                                                'Changed' { $text = "MasterData's Schedule sheet changed for $label since it was loaded, and the board's $label has been edited since ($($st.SheetCount) line(s) on the sheet now)."; $action = "Replace $label with the sheet" }
                                                'New' { $text = "MasterData's Schedule sheet has $($st.SheetCount) line(s) for $label that are not on the board."; $action = "Load them (replaces $label)" }
                                                'EmptyDay' { $text = "MasterData's Schedule sheet has $($st.SheetCount) line(s) for $label."; $action = "Load $label" }
                                }
                                if ($null -ne $text) {
                                                return [PSCustomObject]@{ Day = $d; TargetDay = $d; State = $st.State; Count = $st.SheetCount; Fingerprint = $st.Fingerprint; Text = $text; ActionText = $action }
                                }
                }
                $later = @(Get-SheetDaysToAdd -Board $Board -SheetRows $SheetRows -Ignore $Ignore)
                foreach ($g in $later) {
                                $label = Format-BoardDayLabel $g.Day
                                return [PSCustomObject]@{
                                                Day         = $g.Day
                                                TargetDay   = $g.Day
                                                State       = 'OtherDay'
                                                Count       = $g.Count
                                                Fingerprint = $g.Fingerprint
                                                Text        = "MasterData's Schedule sheet also has $($g.Count) line(s) for $label."
                                                ActionText  = "Add $label to the board"
                                }
                }
                # A sheet with lines only for days before the board's first day loads nothing by itself,
                # so while that first day is empty say why, and offer the latest of those days for it
                $first = if ($Board.Days.Count -gt 0) { [string]$Board.Days[0] } else { '' }
                if ($first -eq '' -or @(Get-BoardDayLines -Board $Board -Day $first).Count -gt 0) { return $null }
                $sheetDays = @(Get-ScheduleSheetDays -Rows $SheetRows | Where-Object { $_.Day -ne '' })
                if (@($sheetDays | Where-Object { $_.Day -ge $first }).Count -gt 0) { return $null }
                $past = @($sheetDays | Where-Object { $_.Day -lt $first })
                if ($past.Count -eq 0) { return $null }
                $g = $past[-1]
                $fp = Get-ScheduleDayFingerprint -Rows $SheetRows -Day $g.Day
                if ($Ignore -contains "$($g.Day)|$fp") { return $null }
                $label = Format-BoardDayLabel $g.Day
                $firstLabel = Format-BoardDayLabel $first
                return [PSCustomObject]@{
                                Day         = $g.Day
                                TargetDay   = $first
                                State       = 'PastDay'
                                Count       = $g.Count
                                Fingerprint = $fp
                                Text        = "MasterData's Schedule sheet only has lines for days before $firstLabel (the board's first day): $($g.Count) line(s) for $label. They were not loaded."
                                ActionText  = "Put them on $firstLabel"
                }
}

# The days MasterData's Schedule sheet has lines for that are not on the board yet: from the
# board's first day on (earlier days are not offered), and not waved off. -Ignore holds
# 'day|fingerprint' keys. In date order, each {Day; Count; Fingerprint}.
function Get-SheetDaysToAdd {
                param ($Board, [array]$SheetRows = @(), [string[]]$Ignore = @())

                $first = if ($Board.Days.Count -gt 0) { [string]$Board.Days[0] } else { '' }
                return @(foreach ($g in @(Get-ScheduleSheetDays -Rows $SheetRows)) {
                                if ($g.Day -eq '' -or $Board.Days.Contains($g.Day) -or ($first -ne '' -and $g.Day -lt $first)) { continue }
                                $fp = Get-ScheduleDayFingerprint -Rows $SheetRows -Day $g.Day
                                if ($Ignore -contains "$($g.Day)|$fp") { continue }
                                [PSCustomObject]@{ Day = $g.Day; Count = $g.Count; Fingerprint = $fp }
                })
}

# Days the last exported MasterSchedule has lines for, in order
function Get-PublishedDays {
                param ([array]$Rows = @())

                $days = foreach ($r in $Rows) { $d = Parse-ETD $r.ETD; if ($null -ne $d) { $d.ToString('yyyy-MM-dd') } }
                return @($days | Sort-Object -Unique)
}

# Pairs each "Wingman to <lead>" line with its lead (same day, the lead's student) through
# WingOfLineId. Works on board lines (-IdProperty Id) and engine rows (-IdProperty BoardLineId).
function Set-WingLinksFromNotes {
                param ([array]$Items = @(), [string]$IdProperty = 'Id')

                foreach ($w in $Items) {
                                if ([string]$w.Notes -notmatch '(?i)Wingman to (?<lead>[^|]+)') { continue }
                                $leadName = $Matches['lead'].Trim()
                                if (-not (Test-IsWingCapableType ([string]$w.SortieType))) { continue }
                                $wd = Parse-ETD $w.ETD
                                if ($null -eq $wd) { continue }
                                foreach ($c in $Items) {
                                                if ([string]$c.$IdProperty -eq [string]$w.$IdProperty -or (Get-NameText $c.Student) -ine $leadName) { continue }
                                                $cd = Parse-ETD $c.ETD
                                                if ($null -eq $cd -or $cd.Date -ne $wd.Date -or -not (Test-IsWingCapableType ([string]$c.SortieType))) { continue }
                                                $w.WingOfLineId = [string]$c.$IdProperty
                                                break
                                }
                }
}

# Text copies of exported rows for one day (they may carry student / IP / mission objects)
function Get-PublishedDayRows {
                param ([array]$Rows = @(), [string]$Day)

                return @(foreach ($r in $Rows) {
                                $d = Parse-ETD $r.ETD
                                if ($null -eq $d -or ($Day -and $d.ToString('yyyy-MM-dd') -ne $Day)) { continue }
                                [PSCustomObject]@{
                                                ETD        = $d.ToString('yyyy-MM-dd HH:mm:ss')
                                                Instructor = Get-NameText $r.Instructor
                                                SortieType = [string]$r.SortieType
                                                Student    = Get-NameText $r.Student
                                                Mission    = Get-NameText $r.Mission
                                                Callsign   = $r.Callsign
                                                Num        = $r.Num
                                                Notes      = [string]$r.Notes
                                }
                })
}

# One day of the last exported schedule as board lines with everyone kept where they were:
# students and missions hardcoded, wingmen paired to their leads. For Sale and direct support
# lines come back open. Notes and callsigns are left off; the next fill writes them again.
# -TargetDay puts that day's lines on another day.
function ConvertFrom-PublishedRows {
                param ([array]$Rows = @(), [Parameter(Mandatory)] [string]$Day, [string]$TargetDay = '')

                $lines = @(ConvertFrom-ScheduleRows -Rows (Get-PublishedDayRows -Rows $Rows -Day $Day) -Source 'Exported' -KeepStudents -KeepNotes -TargetDay $TargetDay)
                Set-WingLinksFromNotes -Items $lines -IdProperty 'Id'
                foreach ($l in $lines) { $l.Notes = ''; $l.Callsign = ''; $l.Num = '' }
                return $lines
}

# The last exported schedule's lines for the days a fill does not touch (on/after -FromDay and
# not in -ExcludeDays) as hardcoded engine rows. Filled together with the board's days they
# keep every student's exported events where they are, so filling one day never double-books
# a student or flies their events out of order. For Sale / direct support lines are left out;
# return legs follow their O&B as usual.
function Get-PublishedContextRows {
                param ([array]$Rows = @(), [string]$FromDay = '', [string[]]$ExcludeDays = @())

                $out = New-Object System.Collections.Generic.List[object]
                $i = 0
                foreach ($r in (Get-PublishedDayRows -Rows $Rows)) {
                                $day = ([datetime]$r.ETD).ToString('yyyy-MM-dd')
                                if (($FromDay -ne '' -and $day -lt $FromDay) -or $ExcludeDays -contains $day) { continue }
                                $type = [string]$r.SortieType
                                if ([string]::IsNullOrWhiteSpace($type)) { continue }
                                $isReturn = Test-IsReturnSortieType $type
                                if (-not $isReturn -and (Test-IsSentinelStudent $r.Student)) { continue }
                                $i++
                                $wingNote = $null
                                if ($r.Notes -match '(?i)Wingman to [^|]+') { $wingNote = $Matches[0].Trim() }
                                $out.Add([PSCustomObject]@{
                                                ETD          = $r.ETD
                                                Instructor   = $(if ($r.Instructor -eq '') { $null } else { $r.Instructor })
                                                SortieType   = $type
                                                Student      = $(if ($isReturn) { $null } else { $r.Student })
                                                Mission      = $(if ($isReturn -or $r.Mission -eq '') { $null } else { $r.Mission })
                                                Callsign     = $r.Callsign
                                                Num          = $r.Num
                                                Notes        = $wingNote
                                                BoardLineId  = "pub$i"
                                                LinkedLineId = $null
                                                WingOfLineId = $null
                                                FilledBy     = $null
                                })
                }
                $context = $out.ToArray()
                Set-WingLinksFromNotes -Items $context -IdProperty 'BoardLineId'
                return $context
}

# A fill's lines split into the days that were filled and the exported days that only rode along
# as context. Stats and overrides count the filled days only; overrides on context lines (an edit
# that now breaks a day already exported) are kept as ContextOverrides.
# Returns {Schedule; Context; Students; Overrides; ContextOverrides; Stats; Log; Days}
function Split-FillResult {
                param ($FillResult, [string[]]$Days = @())

                $mine = New-Object System.Collections.Generic.List[object]
                $ctx = New-Object System.Collections.Generic.List[object]
                foreach ($r in @($FillResult.Schedule)) {
                                $d = Parse-ETD $r.ETD
                                $day = if ($null -ne $d) { $d.ToString('yyyy-MM-dd') } else { '' }
                                if ($Days -contains $day) { $mine.Add($r) } else { $ctx.Add($r) }
                }
                $ids = @{}
                foreach ($r in $mine) { $ids[[string]$r.BoardLineId] = $true }
                $ov = @($FillResult.Overrides | Where-Object { $null -ne $_ -and $ids.ContainsKey([string]$_.LineId) })
                $ctxOv = @($FillResult.Overrides | Where-Object { $null -ne $_ -and -not $ids.ContainsKey([string]$_.LineId) })
                $stats = [ordered]@{ Lines = $mine.Count; Pinned = 0; Filled = 0; Wing = 0; Return = 0; DS = 0; ForSale = 0; Overrides = $ov.Count }
                foreach ($r in $mine) {
                                switch ([string]$r.FilledBy) {
                                                'Pin' { $stats.Pinned++ }
                                                'Engine' { $stats.Filled++ }
                                                'Wing' { $stats.Wing++ }
                                                'Return' { $stats.Return++ }
                                                'DS' { $stats.DS++ }
                                                'ForSale' { $stats.ForSale++ }
                                }
                }
                return [PSCustomObject]@{
                                Schedule         = $mine.ToArray()
                                Context          = $ctx.ToArray()
                                Students         = $FillResult.Students
                                Overrides        = $ov
                                ContextOverrides = $ctxOv
                                Stats            = [PSCustomObject]$stats
                                Log              = $FillResult.Log
                                Days             = @($Days | Sort-Object)
                }
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
                                if (@(Get-BoardSheetSyncDays $Board) -contains $day -and $null -ne $Board.SheetSync[$day]) { $dto.Sheet = $Board.SheetSync[$day] }
                                Write-TextFileAtomic -Path (Join-Path $Dir "$day.json") -Text (ConvertTo-Json -InputObject $dto -Depth 6)
                }
}

function Read-BoardDayFile {
                param ([string]$Path)

                $dto = Read-BoardDayDto $Path
                if ($null -eq $dto) { return }
                return @(foreach ($l in @($dto.Lines)) { if ($null -ne $l) { ConvertFrom-LineDto $l } })
}

# A saved day file as read from disk ($null when missing or damaged)
function Read-BoardDayDto {
                param ([string]$Path)

                if (-not (Test-Path -LiteralPath $Path)) { return $null }
                try {
                                $dto = ConvertFrom-Json ([System.IO.File]::ReadAllText($Path))
                                # Check the lines read before handing the file back
                                foreach ($l in @($dto.Lines)) { if ($null -ne $l) { [void](ConvertFrom-LineDto $l) } }
                                return $dto
                } catch {
                                Write-Warning "Could not read saved board $Path : $($_.Exception.Message)"
                                return $null
                }
}

# Days on/after -From that have a saved board file
function Get-SavedBoardDays {
                param ([string]$Dir, [datetime]$From = [datetime]::MinValue)

                if (-not (Test-Path -LiteralPath $Dir)) { return }
                $out = foreach ($f in (Get-ChildItem -LiteralPath $Dir -Filter '*.json' -File)) {
                                $d = [datetime]::MinValue
                                if ([datetime]::TryParseExact($f.BaseName, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$d) -and $d.Date -ge $From.Date) { $f.BaseName }
                }
                return @($out | Sort-Object)
}

# Board for -Days, filled from saved files where they exist
function Open-BoardState {
                param ([string]$Dir, [string[]]$Days, [hashtable]$Settings)

                $board = New-Board -Days $Days -Settings $Settings
                foreach ($day in $board.Days) {
                                $dto = Read-BoardDayDto (Join-Path $Dir "$day.json")
                                if ($null -eq $dto) { continue }
                                foreach ($l in @($dto.Lines)) { if ($null -ne $l) { $board.Lines.Add((ConvertFrom-LineDto $l)) } }
                                if ($null -ne $dto.PSObject.Properties['Sheet'] -and $null -ne $dto.Sheet) { $board.SheetSync[$day] = ConvertTo-SheetSyncEntry $dto.Sheet }
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
