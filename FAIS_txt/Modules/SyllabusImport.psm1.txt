# SYLLABUSIMPORT.PSM1 DIRECTORY:
#             "Update Syllabi": reads a training file (the day's flying schedule or the Big
#             Board), works out which one it is, and brings each FAIS student's syllabus in
#             line with it: on the board, on their OutputSchedule tab and in their SimDebt
# Read-TrainingFile / Get-SheetBlock / ConvertTo-TrainingEtd
# Find-DailyScheduleHeadings / Read-DailyScheduleSheet / Find-BigBoardHeadings / Read-BigBoardSheet / Get-BigBoardInstructor
# ConvertTo-PersonName / Get-PersonNameMatch / Resolve-TrainingStudents / Resolve-TrainingInstructor
# Get-SyllabusUpdatePlan / Get-SyllabusTargetResult / New-SyllabusChange
# Update-SyllabusTabTargets / Invoke-SyllabusUpdate / Save-SyllabusUpdate / Format-SyllabusUpdateSummary
#
# The rules:
#  - Flying schedule (its Flying and Sims tabs): a student's line with a Line Status of CM or
#    EXEC marks its mission complete with the line's IP and ETD; INC or CX makes sure it is not
#    complete. Any other status (or none) changes nothing.
#  - Big Board: anything in an event's columns marks that event complete with the entry's IP and
#    date; nothing there makes sure it is not complete. An event FAIS has planned for later than
#    now is left alone (the Big Board can't show it yet).
#  - Only events in the student's syllabus change; every other event or column is ignored.


# What a line's status on the flying schedule means for its event
$script:CompleteStatuses = @('CM', 'EXEC')
$script:NotCompleteStatuses = @('INC', 'CX')

# Flying schedule headings (letters and digits only, lower case), best first
$script:DailyScheduleHeadings = [ordered]@{
                Student    = @('student', 'studentname')
                Mission    = @('mission', 'eventname', 'event')
                Instructor = @('instructor', 'simulatorinstructor', 'simip', 'ipname')
                ETD        = @('etd', 'starttime', 'startdatetime', 'start')
                Status     = @('linestatus', 'scheduledlinestatus')
}

# Big Board columns that describe the student rather than an event
$script:BigBoardInfoHeadings = @('track', 'org', 'organization', 'class', 'flight')


# =================================================================================
#  READING THE FILE
# =================================================================================

# Reads a training file and works out what it is. Returns
#   {Path; Kind ('DailySchedule' / 'BigBoard' / ''); Label; Day; Sheets; Events; Facts; People; Error}
# Each fact: {Student; Mission; Complete; Explicit; Status; Instructor; ETD ([datetime] or $null)}
#   Explicit is $false only for the Big Board's empty columns (they never take back an event
#   FAIS has planned for later). People: everyone the file lists, matched to students later.
function Read-TrainingFile {
                param ([string]$Path)

                $result = [PSCustomObject]@{ Path = $Path; Kind = ''; Label = ''; Day = $null; Sheets = @(); Events = @(); Facts = @(); People = @(); Error = '' }
                $fileName = [System.IO.Path]::GetFileName([string]$Path)
                if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
                                $result.Error = "$Path was not found."
                                return $result
                }

                $book = Open-FaisWorkbookForRead -Path $Path
                try {
                                if ($null -eq $book.Package) {
                                                $result.Error = "Could not open $fileName ($($book.Error)). Pick the .xlsx file."
                                                return $result
                                }
                                $sheets = @($book.Package.Workbook.Worksheets | Where-Object { $null -ne $_.Dimension })

                                # The flying schedule: every tab with Student, Mission and Line Status columns (Flying, Sims)
                                $tabs = @(foreach ($ws in $sheets) {
                                                $headings = Find-DailyScheduleHeadings $ws
                                                if ($null -ne $headings) { Read-DailyScheduleSheet -Worksheet $ws -Headings $headings }
                                })
                                if ($tabs.Count -gt 0) {
                                                $result.Kind = 'DailySchedule'
                                                $result.Sheets = @($tabs | ForEach-Object { $_.Name })
                                                $result.Facts = @($tabs | ForEach-Object { $_.Facts })
                                                $result.People = @($tabs | ForEach-Object { $_.People } | Sort-Object -Unique)
                                                $days = @($tabs | Where-Object { $null -ne $_.Day } | ForEach-Object { $_.Day })
                                                if ($days.Count -gt 0) { $result.Day = $days[0] }
                                                $result.Label = if ($null -ne $result.Day) { "Flying schedule for $($result.Day.ToString('ddd dd MMM yyyy'))" } else { 'Flying schedule' }
                                                return $result
                                }

                                foreach ($ws in $sheets) {
                                                $headings = Find-BigBoardHeadings $ws
                                                if ($null -eq $headings) { continue }
                                                $board = Read-BigBoardSheet -Worksheet $ws -Headings $headings
                                                $result.Kind = 'BigBoard'
                                                $result.Sheets = @($ws.Name)
                                                $result.Events = @($headings.Events | ForEach-Object { $_.Name })
                                                $result.Facts = $board.Facts
                                                $result.People = $board.People
                                                $result.Label = if ($board.Title) { "Big Board ($($board.Title))" } else { 'Big Board' }
                                                return $result
                                }
                                $result.Error = "$fileName is not a flying schedule (Flying / Sims tabs with Student, Mission and Line Status columns) or a Big Board (a Student Name column followed by the events), so nothing was changed."
                } catch {
                                $result.Error = "Could not read $fileName : $($_.Exception.Message)"
                } finally {
                                Close-FaisWorkbookForRead $book
                }
                return $result
}

# Cell values from row 1 / column 1 to -ToRow / -ToCol as a 2-D array, [row, column] from 0
function Get-SheetBlock {
                param ($Worksheet, [int]$ToRow, [int]$ToCol)

                $v = $Worksheet.Cells[1, 1, $ToRow, $ToCol].Value
                if ($v -is [System.Array] -and $v.Rank -eq 2) { return , $v }
                $one = New-Object 'object[,]' 1, 1
                $one[0, 0] = $v
                return , $one
}

# A training file's date / time cell as [datetime], or $null. Reads real dates, Excel serials,
# the Big Board's "07-24-2026-1500", and times with no date (put on -Day; $null without one).
function ConvertTo-TrainingEtd {
                param ($Value, $Day)

                $text = ConvertTo-ScheduleCellText $Value
                if ($text -match '^\d{1,2}-\d{1,2}-\d{4}-\d{4}$') {
                                $dt = [datetime]::MinValue
                                if ([datetime]::TryParseExact($text, 'M-d-yyyy-HHmm', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) { return $dt }
                                return $null
                }
                $tv = ConvertTo-FaisTimeValue $Value
                if ($null -eq $tv) { return $null }
                if ($tv.HasDate) { return $tv.Value }
                if ($null -ne $Day) { return ([datetime]$Day).Date.Add($tv.Value.TimeOfDay) }
                return $null
}


# =================================================================================
#  THE FLYING SCHEDULE
#  One tab per kind of line (Flying, Sims, ...), headings on the first row. The Flying tab
#  has Student / Mission / Instructor / ETD / Line Status; the Sims tab has Student /
#  Event Name / Simulator Instructor / Start Time / Scheduled Line Status.
# =================================================================================

# A flying schedule tab's heading row: {Row; Columns (field -> column)}, or $null when none of
# its first 10 rows has Student, Mission (or Event Name) and Line Status headings
function Find-DailyScheduleHeadings {
                param ($Worksheet)

                $lastRow = [math]::Min(10, $Worksheet.Dimension.End.Row)
                $lastCol = $Worksheet.Dimension.End.Column
                $values = Get-SheetBlock -Worksheet $Worksheet -ToRow $lastRow -ToCol $lastCol
                for ($r = 0; $r -lt $lastRow; $r++) {
                                # heading -> its first column
                                $at = @{}
                                for ($c = $lastCol - 1; $c -ge 0; $c--) {
                                                $v = $values[$r, $c]
                                                if ($null -eq $v) { continue }
                                                $key = ((ConvertTo-ScheduleCellText $v) -replace '[^a-zA-Z0-9]', '').ToLowerInvariant()
                                                if ($key -ne '') { $at[$key] = $c + 1 }
                                }
                                $map = @{}
                                foreach ($field in $script:DailyScheduleHeadings.Keys) {
                                                foreach ($heading in $script:DailyScheduleHeadings[$field]) {
                                                                if ($at.ContainsKey($heading)) { $map[$field] = $at[$heading]; break }
                                                }
                                }
                                if ($map.ContainsKey('Student') -and $map.ContainsKey('Mission') -and $map.ContainsKey('Status')) {
                                                return [PSCustomObject]@{ Row = $r + 1; Columns = $map }
                                }
                }
                return $null
}

# One flying schedule tab: {Name; Day; Facts; People}. A line is a fact when it has a student,
# a mission and a Line Status of CM / EXEC (complete) or INC / CX (not complete).
# An ETD with no date takes the tab's date ("Flying (21-Sep-2026)").
function Read-DailyScheduleSheet {
                param ($Worksheet, $Headings)

                $day = $null
                if ($Worksheet.Name -match '\((\d{1,2}-[A-Za-z]{3}-\d{4})\)') {
                                $parsed = [datetime]::MinValue
                                if ([datetime]::TryParseExact($Matches[1], 'd-MMM-yyyy', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) { $day = $parsed }
                }
                $lastRow = $Worksheet.Dimension.End.Row
                $values = Get-SheetBlock -Worksheet $Worksheet -ToRow $lastRow -ToCol $Worksheet.Dimension.End.Column
                $col = $Headings.Columns

                $facts = New-Object System.Collections.Generic.List[object]
                $people = New-Object System.Collections.Generic.List[string]
                # Row index r (from 0) is sheet row r + 1, so the lines start at index Row
                for ($r = $Headings.Row; $r -lt $lastRow; $r++) {
                                $student = ConvertTo-ScheduleCellText $values[$r, ($col.Student - 1)]
                                if ($student -eq '' -or (Test-IsSentinelStudent $student)) { continue }
                                $people.Add($student)
                                $status = (ConvertTo-ScheduleCellText $values[$r, ($col.Status - 1)]).ToUpperInvariant()
                                $complete = ($script:CompleteStatuses -contains $status)
                                if (-not $complete -and $script:NotCompleteStatuses -notcontains $status) { continue }
                                $mission = ConvertTo-ScheduleCellText $values[$r, ($col.Mission - 1)]
                                if ($mission -eq '') { continue }
                                $ip = if ($col.ContainsKey('Instructor')) { ConvertTo-ScheduleCellText $values[$r, ($col.Instructor - 1)] } else { '' }
                                $etd = if ($col.ContainsKey('ETD')) { ConvertTo-TrainingEtd -Value $values[$r, ($col.ETD - 1)] -Day $day } else { $null }
                                if ($null -eq $day -and $null -ne $etd) { $day = $etd.Date }
                                $facts.Add([PSCustomObject]@{ Student = $student; Mission = $mission; Complete = $complete; Explicit = $true; Status = $status; Instructor = $ip; ETD = $etd })
                }
                return [PSCustomObject]@{ Name = $Worksheet.Name; Day = $day; Facts = $facts.ToArray(); People = $people.ToArray() }
}


# =================================================================================
#  THE BIG BOARD
#  One row per student under a heading row: Student Name, Track, Org, Class, then one
#  heading per event (BA101, FT5001, ...) over a block of columns. A flown event fills
#  its block: date "MM-DD-YYYY-HHMM", status (Complete), grade, IP.
# =================================================================================

# The Big Board's heading row: {Row; StudentCol; Events (each {Name; First; Last} column)}, or
# $null. It is the first row (of the first 20) with a Student Name heading and at least three
# event codes after it; each heading to its right starts a block that runs to the next heading.
function Find-BigBoardHeadings {
                param ($Worksheet)

                $lastRow = [math]::Min(20, $Worksheet.Dimension.End.Row)
                $lastCol = $Worksheet.Dimension.End.Column
                $values = Get-SheetBlock -Worksheet $Worksheet -ToRow $lastRow -ToCol $lastCol
                for ($r = 0; $r -lt $lastRow; $r++) {
                                $studentCol = 0
                                for ($c = 0; $c -lt $lastCol; $c++) {
                                                $v = $values[$r, $c]
                                                if ($null -eq $v) { continue }
                                                $key = ((ConvertTo-ScheduleCellText $v) -replace '[^a-zA-Z0-9]', '').ToLowerInvariant()
                                                if ($key -eq 'studentname' -or $key -eq 'student') { $studentCol = $c + 1; break }
                                }
                                if ($studentCol -eq 0) { continue }

                                $starts = @(for ($c = $studentCol; $c -lt $lastCol; $c++) { $v = $values[$r, $c]; if ($null -ne $v -and (ConvertTo-ScheduleCellText $v) -ne '') { $c + 1 } })
                                $events = New-Object System.Collections.Generic.List[object]
                                for ($i = 0; $i -lt $starts.Count; $i++) {
                                                $name = ConvertTo-ScheduleCellText $values[$r, ($starts[$i] - 1)]
                                                if ($script:BigBoardInfoHeadings -contains (($name -replace '[^a-zA-Z0-9]', '').ToLowerInvariant())) { continue }
                                                $last = if ($i + 1 -lt $starts.Count) { $starts[$i + 1] - 1 } else { $lastCol }
                                                $events.Add([PSCustomObject]@{ Name = $name; First = $starts[$i]; Last = $last })
                                }
                                $codes = @($events.ToArray() | Where-Object { $_.Name -match '^[A-Za-z]{1,4}\d{2,5}' })
                                if ($codes.Count -lt 3) { continue }
                                return [PSCustomObject]@{ Row = $r + 1; StudentCol = $studentCol; Events = $events.ToArray() }
                }
                return $null
}

# The Big Board's students: {Title; Facts; People}. Each student gets one fact per event:
# complete when anything is in the event's columns (its first cell is the date), with that
# date as the ETD and the entry's IP as the instructor; not complete when they are empty.
function Read-BigBoardSheet {
                param ($Worksheet, $Headings)

                $lastRow = $Worksheet.Dimension.End.Row
                $lastCol = $Worksheet.Dimension.End.Column
                $values = Get-SheetBlock -Worksheet $Worksheet -ToRow $lastRow -ToCol $lastCol

                # The title above the headings ("Scheduling Aid for Organization ...")
                $title = ''
                for ($r = 0; $r -lt $Headings.Row - 1 -and $title -eq ''; $r++) {
                                for ($c = 0; $c -lt $lastCol; $c++) {
                                                if ($null -eq $values[$r, $c]) { continue }
                                                $title = ConvertTo-ScheduleCellText $values[$r, $c]
                                                if ($title -ne '') { break }
                                }
                }

                $facts = New-Object System.Collections.Generic.List[object]
                $people = New-Object System.Collections.Generic.List[string]
                $texts = New-Object System.Collections.Generic.List[string]
                for ($r = $Headings.Row; $r -lt $lastRow; $r++) {
                                $student = ConvertTo-ScheduleCellText $values[$r, ($Headings.StudentCol - 1)]
                                if ($student -eq '' -or (Test-IsSentinelStudent $student)) { continue }
                                $people.Add($student)
                                foreach ($ev in $Headings.Events) {
                                                $etd = $null
                                                $texts.Clear()
                                                for ($c = $ev.First - 1; $c -lt $ev.Last; $c++) {
                                                                # Most cells are empty: test them here rather than in a function call
                                                                $v = $values[$r, $c]
                                                                if ($null -eq $v) { continue }
                                                                $text = if ($v -is [string]) { $v.Trim() } else { ConvertTo-ScheduleCellText $v }
                                                                if ($text -eq '') { continue }
                                                                if ($texts.Count -eq 0 -and $null -eq $etd) {
                                                                                $etd = ConvertTo-TrainingEtd -Value $v
                                                                                if ($null -ne $etd) { continue }
                                                                }
                                                                $texts.Add($text)
                                                }
                                                $has = ($null -ne $etd -or $texts.Count -gt 0)
                                                $ip = if ($has) { Get-BigBoardInstructor -Texts $texts.ToArray() } else { '' }
                                                $facts.Add([PSCustomObject]@{ Student = $student; Mission = $ev.Name; Complete = $has; Explicit = $has; Status = ''; Instructor = $ip; ETD = $etd })
                                }
                }
                return [PSCustomObject]@{ Title = $title; Facts = $facts.ToArray(); People = $people.ToArray() }
}

# The IP in a Big Board entry: its last text that is not a status (Complete), a grade
# (G, E, U) or a number
function Get-BigBoardInstructor {
                param ([string[]]$Texts = @())

                for ($i = $Texts.Count - 1; $i -ge 0; $i--) {
                                $t = $Texts[$i]
                                if ($t -match '^[A-Za-z]{1,2}$' -or $t -match '^(?i)((in)?completed?|inc|exec|cnx|sched(uled)?)$') { continue }
                                if ($t -notmatch '[A-Za-z]{2}') { continue }
                                return $t
                }
                return ''
}


# =================================================================================
#  WHO IS WHO
#  The flying schedule lists the whole wing ("Quill, Casey B. [2d Lt]"), so a FAIS student
#  is matched on more than a last name whenever FAIS has more than a last name for them.
# =================================================================================

# A name split for matching: "Quill, Casey B. [2d Lt]" -> Last 'quill', Given @('casey', 'b').
# Tags in [ ], periods, apostrophes and case don't count; a hyphen counts as a space.
# A name with no comma is all last name ("Quill").
function ConvertTo-PersonName {
                param ([string]$Name)

                $text = Get-BaseName $Name
                $last = $text
                $given = ''
                $comma = $text.IndexOf(',')
                if ($comma -ge 0) { $last = $text.Substring(0, $comma); $given = $text.Substring($comma + 1) }
                $clean = {
                                param ([string]$Part)
                                $s = ($Part.ToLowerInvariant() -replace "['.]", '') -replace '[^\p{L}\p{Nd}]', ' '
                                return ($s -replace '\s+', ' ').Trim()
                }
                $l = & $clean $last
                $g = & $clean $given
                # Assigned straight from @() so a single first name stays a one-item array
                $tokens = @()
                if ($g -ne '') { $tokens = @($g -split ' ') }
                return [PSCustomObject]@{ Last = $l; Given = $tokens; Key = ($l + '|' + ($tokens -join ' ')) }
}

# How a FAIS name (-Known) matches a name from a file (-Other), both from ConvertTo-PersonName:
#   'Exact'    same name
#   'Given'    same last name, and the first names (and middle initials, as far as both go) agree;
#              an initial agrees with any name it starts ("c" = "casey")
#   'LastOnly' same last name and one of them has no first name
#   ''         not the same person
function Get-PersonNameMatch {
                param ($Known, $Other)

                if ($Known.Last -eq '' -or $Known.Last -ne $Other.Last) { return '' }
                if ($Known.Key -eq $Other.Key) { return 'Exact' }
                if ($Known.Given.Count -eq 0 -or $Other.Given.Count -eq 0) { return 'LastOnly' }
                $n = [math]::Min($Known.Given.Count, $Other.Given.Count)
                for ($i = 0; $i -lt $n; $i++) {
                                $a = [string]$Known.Given[$i]
                                $b = [string]$Other.Given[$i]
                                if ($a -eq $b) { continue }
                                if (($a.Length -eq 1 -and $b.StartsWith($a)) -or ($b.Length -eq 1 -and $a.StartsWith($b))) { continue }
                                return ''
                }
                return 'Given'
}

# Matches the people a training file names to FAIS students. Returns {Map; Unknown; Ambiguous}:
#   Map: file name -> student, by the closest match (Exact, then Given, then LastOnly). A student
#   FAIS knows by last name only matches when just one person in the file has that last name.
#   Unknown: names that are no FAIS student. Ambiguous: names left alone because they could be
#   more than one student, or a student who could be more than one person in the file.
function Resolve-TrainingStudents {
                param ([string[]]$Names = @(), [array]$Students = @())

                $known = @(foreach ($s in $Students) {
                                if ($null -ne $s -and -not [string]::IsNullOrWhiteSpace([string]$s.Name)) { [PSCustomObject]@{ Student = $s; Person = ConvertTo-PersonName ([string]$s.Name) } }
                })
                $people = @(foreach ($n in @($Names | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)) {
                                [PSCustomObject]@{ Raw = $n; Person = ConvertTo-PersonName $n }
                })
                # Last name -> how many different people in the file have it
                $keysByLast = @{}
                foreach ($p in $people) {
                                if (-not $keysByLast.ContainsKey($p.Person.Last)) { $keysByLast[$p.Person.Last] = New-Object System.Collections.Generic.HashSet[string] }
                                [void]$keysByLast[$p.Person.Last].Add($p.Person.Key)
                }

                $map = @{}
                $unknown = New-Object System.Collections.Generic.List[string]
                $ambiguous = New-Object System.Collections.Generic.List[string]
                foreach ($p in $people) {
                                $found = @{ Exact = @(); Given = @(); LastOnly = @() }
                                foreach ($k in $known) {
                                                $how = Get-PersonNameMatch -Known $k.Person -Other $p.Person
                                                if ($how) { $found[$how] += $k.Student }
                                }
                                $matched = $false
                                foreach ($how in @('Exact', 'Given', 'LastOnly')) {
                                                $hits = @($found[$how])
                                                if ($hits.Count -eq 0) { continue }
                                                $matched = $true
                                                if ($hits.Count -gt 1) { $ambiguous.Add("$($p.Raw) (could be $(@($hits | ForEach-Object { $_.Name }) -join ' or '))") }
                                                elseif ($how -eq 'LastOnly' -and $keysByLast[$p.Person.Last].Count -gt 1) {
                                                                $line = "$($hits[0].Name) (more than one person on the file is called $(Get-NormalizedLastName (Get-BaseName $p.Raw)))"
                                                                if (-not $ambiguous.Contains($line)) { $ambiguous.Add($line) }
                                                }
                                                else { $map[$p.Raw] = $hits[0] }
                                                break
                                }
                                if (-not $matched) { $unknown.Add($p.Raw) }
                }

                # A student matched by two names that aren't the same person is left alone
                $rawsOf = @{}
                foreach ($raw in @($map.Keys)) {
                                $name = [string]$map[$raw].Name
                                if (-not $rawsOf.ContainsKey($name)) { $rawsOf[$name] = New-Object System.Collections.Generic.List[string] }
                                $rawsOf[$name].Add($raw)
                }
                foreach ($name in @($rawsOf.Keys | Sort-Object)) {
                                $raws = @($rawsOf[$name].ToArray() | Sort-Object)
                                $same = $true
                                for ($i = 0; $i -lt $raws.Count -and $same; $i++) {
                                                for ($j = $i + 1; $j -lt $raws.Count; $j++) {
                                                                if (-not (Get-PersonNameMatch -Known (ConvertTo-PersonName $raws[$i]) -Other (ConvertTo-PersonName $raws[$j]))) { $same = $false; break }
                                                }
                                }
                                if (-not $same) {
                                                foreach ($raw in $raws) { $map.Remove($raw) }
                                                $ambiguous.Add("$name (the file has $($raws -join ' and '))")
                                }
                }
                return [PSCustomObject]@{ Map = $map; Unknown = $unknown.ToArray(); Ambiguous = $ambiguous.ToArray() }
}

# The IP name to put on a syllabus event: the FAIS instructor it matches (as students are
# matched, but no other file names are checked), else the name as the file has it, without
# its rank tag. '' for no name.
function Resolve-TrainingInstructor {
                param ([string]$Name, [array]$Instructors = @())

                $plain = ((Get-BaseName $Name) -replace '\s+', ' ').Trim()
                if ($plain -eq '') { return '' }
                $person = ConvertTo-PersonName $Name
                $found = @{ Exact = @(); Given = @(); LastOnly = @() }
                foreach ($ip in @($Instructors)) {
                                if ($null -eq $ip -or [string]::IsNullOrWhiteSpace([string]$ip.Name)) { continue }
                                $how = Get-PersonNameMatch -Known (ConvertTo-PersonName ([string]$ip.Name)) -Other $person
                                if ($how) { $found[$how] += $ip }
                }
                foreach ($how in @('Exact', 'Given', 'LastOnly')) {
                                $hits = @($found[$how])
                                if ($hits.Count -eq 1) { return ([string]$hits[0].Name).Trim() }
                                if ($hits.Count -gt 1) { break }
                }
                return $plain
}


# =================================================================================
#  WHAT CHANGES
# =================================================================================

# Works out what a training file changes. -File is a Read-TrainingFile result. Returns
#   {Kind; Targets; Changes; Planned; Matched; Unknown; Ambiguous; NotInSyllabus; NotOnFile}
#   Targets: one per student and syllabus event the file speaks for:
#            {Student; Mission; Complete; Explicit; Status; Instructor; ETD}. A complete line
#            wins over a not-complete one for the same event (the latest, when there are several).
#   Changes: the targets that change something, against each student's tab in -OutputPath (what
#            is saved), else their syllabus in memory (see New-SyllabusChange).
#   Planned: events the Big Board has no entry for that FAIS has planned after -Now (left alone).
#   NotInSyllabus: "student mission" for flying schedule lines whose mission is not in the
#            student's syllabus (the Big Board's other events are simply ignored).
#   NotOnFile: FAIS students the Big Board does not list.
function Get-SyllabusUpdatePlan {
                param (
                                [Parameter(Mandatory)] $File,
                                [array]$Students = @(),
                                [array]$Instructors = @(),
                                [string]$OutputPath = '',
                                $SoftWeights,
                                [datetime]$Now = (Get-Date)
                )

                $isBoard = ($File.Kind -eq 'BigBoard')
                $match = Resolve-TrainingStudents -Names @($File.People) -Students $Students

                # Facts grouped by student and syllabus event; each student's events looked up by name once
                $eventsOf = @{}
                $groups = [ordered]@{}
                $notInSyllabus = New-Object System.Collections.Generic.List[string]
                foreach ($fact in @($File.Facts)) {
                                $student = $match.Map[[string]$fact.Student]
                                if ($null -eq $student) { continue }
                                $name = [string]$student.Name
                                if (-not $eventsOf.ContainsKey($name)) {
                                                $byName = @{}
                                                foreach ($m in @($student.Syllabus)) {
                                                                if ($null -eq $m) { continue }
                                                                $k = ([string]$m.Mission).Trim().ToUpperInvariant()
                                                                if ($k -ne '' -and -not $byName.ContainsKey($k)) { $byName[$k] = $m }
                                                }
                                                $eventsOf[$name] = $byName
                                }
                                $key = ([string]$fact.Mission).Trim().ToUpperInvariant()
                                $entry = $eventsOf[$name][$key]
                                if ($null -eq $entry) {
                                                if (-not $isBoard) { $notInSyllabus.Add("$name $(([string]$fact.Mission).Trim())") }
                                                continue
                                }
                                $groupKey = "$name|$key"
                                if (-not $groups.Contains($groupKey)) { $groups[$groupKey] = New-Object System.Collections.Generic.List[object] }
                                $groups[$groupKey].Add([PSCustomObject]@{ Fact = $fact; Student = $student; Entry = $entry })
                }

                $ipOf = @{}
                $targets = New-Object System.Collections.Generic.List[object]
                foreach ($groupKey in $groups.Keys) {
                                $list = $groups[$groupKey]
                                $done = @($list | Where-Object { $_.Fact.Complete } | Sort-Object -Property { if ($null -ne $_.Fact.ETD) { $_.Fact.ETD } else { [datetime]::MinValue } })
                                $pick = if ($done.Count -gt 0) { $done[$done.Count - 1] } else { $list[0] }
                                $fact = $pick.Fact
                                $rawIp = [string]$fact.Instructor
                                if (-not $ipOf.ContainsKey($rawIp)) { $ipOf[$rawIp] = Resolve-TrainingInstructor -Name $rawIp -Instructors $Instructors }
                                $targets.Add([PSCustomObject]@{
                                                Student    = $pick.Student
                                                Mission    = ([string]$pick.Entry.Mission).Trim()
                                                Complete   = [bool]$fact.Complete
                                                Explicit   = [bool]$fact.Explicit
                                                Status     = [string]$fact.Status
                                                Instructor = $ipOf[$rawIp]
                                                ETD        = $fact.ETD
                                })
                }

                # What changes: against the student's tab where they have one (that is what is saved)
                $byStudent = [ordered]@{}
                foreach ($t in $targets) {
                                $name = [string]$t.Student.Name
                                if (-not $byStudent.Contains($name)) { $byStudent[$name] = New-Object System.Collections.Generic.List[object] }
                                $byStudent[$name].Add($t)
                }
                $changes = New-Object System.Collections.Generic.List[object]
                $book = Open-FaisWorkbookForRead -Path $OutputPath
                try {
                                foreach ($name in $byStudent.Keys) {
                                                $ws = $null
                                                if ($null -ne $book.Package) { $ws = $book.Package.Workbook.Worksheets[$name] }
                                                $mine = $byStudent[$name].ToArray()
                                                $found = if ($null -ne $ws -and $null -ne $ws.Dimension) {
                                                                Update-SyllabusTabTargets -Worksheet $ws -Targets $mine -SoftWeights $SoftWeights -Now $Now -Preview
                                                }
                                                else {
                                                                Invoke-SyllabusUpdate -Targets $mine -SoftWeights $SoftWeights -Now $Now -Preview
                                                }
                                                foreach ($c in @($found)) { if ($null -ne $c) { $changes.Add($c) } }
                                }
                } finally {
                                Close-FaisWorkbookForRead $book
                }

                $matched = @($match.Map.Values | ForEach-Object { [string]$_.Name } | Sort-Object -Unique)
                $notOnFile = @()
                if ($isBoard) { $notOnFile = @($Students | Where-Object { $null -ne $_ -and $matched -notcontains [string]$_.Name } | ForEach-Object { [string]$_.Name }) }
                return [PSCustomObject]@{
                                Kind          = $File.Kind
                                Targets       = $targets.ToArray()
                                Changes       = @($changes.ToArray() | Where-Object { $_.Kind -ne 'Planned' })
                                Planned       = @($changes.ToArray() | Where-Object { $_.Kind -eq 'Planned' })
                                Matched       = $matched
                                Unknown       = @($match.Unknown)
                                Ambiguous     = @($match.Ambiguous)
                                NotInSyllabus = @($notInSyllabus.ToArray() | Sort-Object -Unique)
                                NotOnFile     = $notOnFile
                }
}

# What one event becomes under a target. -Current is the event as it is now ({IsComplete;
# Instructor; ETD}: a syllabus entry or a tab row). Returns $null when it already matches,
# Kind 'Planned' for a Big Board "not complete" on an event FAIS has planned after -Now (left
# as it is), else {Kind (Complete / Updated / NotComplete); IsComplete; Instructor; ETD}.
# Updated = it was complete already and its IP or ETD changes. A target with no IP or ETD
# keeps the event's own when it was already complete.
function Get-SyllabusTargetResult {
                param ($Current, $Target, [datetime]$Now = (Get-Date))

                $wasComplete = ConvertTo-FaisBool $Current.IsComplete
                $etdNow = Format-TabEtd $Current.ETD
                $ipNow = Get-NameText $Current.Instructor
                if ($Target.Complete) {
                                $etd = if ($null -ne $Target.ETD) { ([datetime]$Target.ETD).ToString('yyyy-MM-dd HH:mm:ss') } elseif ($wasComplete) { $etdNow } else { '' }
                                $ip = if ([string]$Target.Instructor -ne '') { [string]$Target.Instructor } elseif ($wasComplete) { $ipNow } else { '' }
                                if ($wasComplete -and $etd -eq $etdNow -and $ip -ceq $ipNow) { return $null }
                                $kind = if ($wasComplete) { 'Updated' } else { 'Complete' }
                                return [PSCustomObject]@{ Kind = $kind; IsComplete = $true; Instructor = $ip; ETD = $etd }
                }
                if (-not $wasComplete) { return $null }
                if (-not $Target.Explicit) {
                                $planned = Parse-ETD $Current.ETD
                                if ($null -ne $planned -and $planned -gt $Now) { return [PSCustomObject]@{ Kind = 'Planned'; IsComplete = $true; Instructor = $ipNow; ETD = $etdNow } }
                }
                return [PSCustomObject]@{ Kind = 'NotComplete'; IsComplete = $false; Instructor = ''; ETD = '' }
}

# One change, for the summary and the save: {Student; Mission; Kind; Instructor; ETD; Status;
# SimDebt (what it does to the student's sim debt)}
function New-SyllabusChange {
                param ($Target, $Result, $SoftWeights)

                $debt = 0.0
                if ($Result.Kind -eq 'Complete') { $debt = Get-MissionSimDebtDelta -MissionName ([string]$Target.Mission) -SoftWeights $SoftWeights }
                elseif ($Result.Kind -eq 'NotComplete') { $debt = - (Get-MissionSimDebtDelta -MissionName ([string]$Target.Mission) -SoftWeights $SoftWeights) }
                return [PSCustomObject]@{
                                Student    = [string]$Target.Student.Name
                                Mission    = [string]$Target.Mission
                                Kind       = [string]$Result.Kind
                                Instructor = [string]$Result.Instructor
                                ETD        = [string]$Result.ETD
                                Status     = [string]$Target.Status
                                SimDebt    = [double]$debt
                }
}


# =================================================================================
#  MAKING THE CHANGES
# =================================================================================

# Brings a student's tab in OutputSchedule in line with -Targets (theirs only): the Instructor,
# ETD and IsComplete cells of each target's event (its first row on the tab) change when they
# differ; nothing else on the tab is touched. -Preview only works the changes out.
# Returns the changes (New-SyllabusChange), events left as planned included (Kind 'Planned').
function Update-SyllabusTabTargets {
                param (
                                [Parameter(Mandatory)] $Worksheet,
                                [array]$Targets = @(),
                                $SoftWeights,
                                [datetime]$Now = (Get-Date),
                                [switch]$Preview
                )

                $lastRow = $Worksheet.Dimension.End.Row
                $lastCol = $Worksheet.Dimension.End.Column
                $col = @{}
                for ($c = 1; $c -le $lastCol; $c++) {
                                $heading = ([string]$Worksheet.Cells[1, $c].Value).Trim()
                                if ($heading -ne '' -and -not $col.ContainsKey($heading)) { $col[$heading] = $c }
                }
                if (-not $col.ContainsKey('Mission')) {
                                throw "The '$($Worksheet.Name)' tab has no Mission heading in row 1, so its events can't be matched."
                }
                if (-not $Preview) {
                                foreach ($heading in @('Instructor', 'ETD', 'IsComplete')) {
                                                if (-not $col.ContainsKey($heading)) {
                                                                $lastCol++
                                                                $Worksheet.Cells[1, $lastCol].Value = $heading
                                                                $col[$heading] = $lastCol
                                                }
                                }
                }

                # Mission -> its first row on the tab
                $rowOf = @{}
                for ($r = 2; $r -le $lastRow; $r++) {
                                $key = ([string]$Worksheet.Cells[$r, $col['Mission']].Value).Trim().ToUpperInvariant()
                                if ($key -ne '' -and -not $rowOf.ContainsKey($key)) { $rowOf[$key] = $r }
                }

                $out = New-Object System.Collections.Generic.List[object]
                foreach ($t in $Targets) {
                                $key = ([string]$t.Mission).Trim().ToUpperInvariant()
                                if (-not $rowOf.ContainsKey($key)) { continue }
                                $r = $rowOf[$key]
                                $current = @{}
                                foreach ($heading in @('IsComplete', 'Instructor', 'ETD')) {
                                                $current[$heading] = if ($col.ContainsKey($heading)) { $Worksheet.Cells[$r, $col[$heading]].Value } else { $null }
                                }
                                $res = Get-SyllabusTargetResult -Current ([PSCustomObject]$current) -Target $t -Now $Now
                                if ($null -eq $res) { continue }
                                if (-not $Preview -and $res.Kind -ne 'Planned') {
                                                $Worksheet.Cells[$r, $col['IsComplete']].Value = [bool]$res.IsComplete
                                                $Worksheet.Cells[$r, $col['ETD']].Value = $(if ($res.ETD -eq '') { $null } else { $res.ETD })
                                                $Worksheet.Cells[$r, $col['Instructor']].Value = $(if ($res.Instructor -eq '') { $null } else { $res.Instructor })
                                }
                                $out.Add((New-SyllabusChange -Target $t -Result $res -SoftWeights $SoftWeights))
                }
                return $out.ToArray()
}

# Applies -Targets to the students' syllabi in memory, through Complete-StudentMission and
# Undo-StudentMission (so SimDebt moves with each event, and unsat flags are left alone).
# -Preview only works the changes out. Returns the changes, as Update-SyllabusTabTargets does.
function Invoke-SyllabusUpdate {
                param ([array]$Targets = @(), $SoftWeights, [datetime]$Now = (Get-Date), [switch]$Preview)

                $out = New-Object System.Collections.Generic.List[object]
                foreach ($t in $Targets) {
                                $entry = Find-SyllabusMission -Student $t.Student -Mission $t.Mission
                                if ($null -eq $entry) { continue }
                                $res = Get-SyllabusTargetResult -Current $entry -Target $t -Now $Now
                                if ($null -eq $res) { continue }
                                if (-not $Preview) {
                                                if ($res.Kind -eq 'NotComplete') { [void](Undo-StudentMission -Student $t.Student -Mission $entry -SoftWeights $SoftWeights) }
                                                elseif ($res.Kind -ne 'Planned') { [void](Complete-StudentMission -Student $t.Student -Mission $entry -InstructorName $res.Instructor -ETD $res.ETD -SoftWeights $SoftWeights) }
                                }
                                $out.Add((New-SyllabusChange -Target $t -Result $res -SoftWeights $SoftWeights))
                }
                return $out.ToArray()
}

# Saves an update right away, both files or neither: each student's tab in OutputSchedule (only
# the cells of the events that change; a student with no tab gets a whole one, from their
# syllabus in memory, so run Invoke-SyllabusUpdate first) and their SimDebt in MasterData,
# moved by what the tab changes did (-MemoryChanges: what Invoke-SyllabusUpdate returned, used
# for the new tabs). A file open in Excel is saved and closed for the write, then opened again.
# Returns {Saved; Changes; Message}.
function Save-SyllabusUpdate {
                param (
                                [array]$Targets = @(),
                                [array]$MemoryChanges = @(),
                                [string]$OutputPath,
                                [string]$MasterDataPath,
                                $SoftWeights,
                                [scriptblock]$LockedPrompt,
                                [datetime]$Now = (Get-Date)
                )

                $result = [PSCustomObject]@{ Saved = $false; Changes = @(); Message = '' }
                if ([string]::IsNullOrWhiteSpace($OutputPath) -or [string]::IsNullOrWhiteSpace($MasterDataPath) -or -not (Test-Path -LiteralPath $MasterDataPath)) {
                                $result.Message = 'MasterData.xlsx or OutputSchedule.xlsx was not found'
                                return $result
                }

                $byStudent = [ordered]@{}
                foreach ($t in $Targets) {
                                $name = [string]$t.Student.Name
                                if (-not $byStudent.Contains($name)) { $byStudent[$name] = New-Object System.Collections.Generic.List[object] }
                                $byStudent[$name].Add($t)
                }
                $files = @($OutputPath, $MasterDataPath)
                $closed = New-Object System.Collections.Generic.List[string]
                try {
                                foreach ($f in $files) {
                                                if (-not (Request-FaisFileWritable -Path $f -Prompt $LockedPrompt -Closed $closed)) {
                                                                $result.Message = "$([System.IO.Path]::GetFileName($f)) is still open"
                                                                return $result
                                                }
                                }

                                $written = New-Object System.Collections.Generic.List[object]
                                $newTabs = New-Object System.Collections.Generic.List[object]
                                $debt = @{}
                                try {
                                                Invoke-WithFileBackup -Files $files -Action {
                                                                $package = Open-ExcelPackage -Path $OutputPath -Create
                                                                try {
                                                                                foreach ($name in $byStudent.Keys) {
                                                                                                $ws = $package.Workbook.Worksheets[$name]
                                                                                                if ($null -eq $ws -or $null -eq $ws.Dimension) { $newTabs.Add($byStudent[$name][0].Student); continue }
                                                                                                foreach ($c in @(Update-SyllabusTabTargets -Worksheet $ws -Targets $byStudent[$name].ToArray() -SoftWeights $SoftWeights -Now $Now)) { $written.Add($c) }
                                                                                }
                                                                                $toSave = @($written | Where-Object { $_.Kind -ne 'Planned' }).Count -gt 0
                                                                                Close-ExcelPackage -ExcelPackage $package -NoSave:(-not $toSave)
                                                                                $package = $null
                                                                }
                                                                catch {
                                                                                if ($null -ne $package) { Close-ExcelPackage -ExcelPackage $package -NoSave }
                                                                                throw
                                                                }
                                                                if ($newTabs.Count -gt 0) {
                                                                                Export-StudentSyllabusTabs -Students $newTabs.ToArray() -FilePath $OutputPath
                                                                                $newNames = @($newTabs | ForEach-Object { [string]$_.Name })
                                                                                foreach ($c in $MemoryChanges) { if ($newNames -contains $c.Student) { $written.Add($c) } }
                                                                }
                                                                foreach ($c in $written) { if ($c.SimDebt -ne 0) { $debt[$c.Student] = [double]$debt[$c.Student] + $c.SimDebt } }
                                                                foreach ($name in @($debt.Keys)) { if ([math]::Abs($debt[$name]) -lt 0.000001) { $debt.Remove($name) } }
                                                                if ($debt.Count -gt 0) { Update-MasterDataStudents -MasterDataPath $MasterDataPath -SimDebtDeltaByName $debt }
                                                }
                                }
                                catch {
                                                $result.Message = $_.Exception.Message
                                                return $result
                                }
                                $result.Saved = $true
                                $result.Changes = $written.ToArray()
                                return $result
                }
                finally {
                                if ($closed.Count -gt 0) { Open-FaisWorkbooks -Path $closed.ToArray() }
                }
}


# =================================================================================
#  TELLING THE USER
# =================================================================================

# The text Update Syllabi shows before it changes anything: what the file is, what changes for
# whom, and who or what is left alone. -MaxLines caps the students listed and -MaxMissions the
# events listed for each (a message box can't scroll; the console gets them all).
function Format-SyllabusUpdateSummary {
                param ($File, $Plan, [int]$MaxLines = 15, [int]$MaxMissions = 8)

                $capped = {
                                param ([string[]]$Items, [int]$Max = 12)
                                $list = @($Items | Where-Object { $_ })
                                if ($list.Count -le $Max) { return ($list -join '; ') }
                                return ((@($list | Select-Object -First $Max) -join '; ') + "; ...and $($list.Count - $Max) more")
                }

                $events = {
                                param ([string[]]$Items)
                                if ($Items.Count -le $MaxMissions) { return ($Items -join ' ') }
                                return ((@($Items | Select-Object -First $MaxMissions) -join ' ') + " +$($Items.Count - $MaxMissions) more")
                }

                $lines = New-Object System.Collections.Generic.List[string]
                $what = if ($File.Kind -eq 'BigBoard') { "$(@($File.People).Count) students, $(@($File.Events).Count) events" } else { "tabs: $(@($File.Sheets) -join ', ')" }
                $lines.Add("$($File.Label)  ($what)")
                $lines.Add("FAIS students on it: $(@($Plan.Matched).Count)")
                $lines.Add('')

                $changes = @($Plan.Changes)
                if ($changes.Count -eq 0) {
                                $lines.Add('Every syllabus already matches it: nothing to change.')
                }
                else {
                                $byStudent = [ordered]@{}
                                foreach ($c in @($changes | Sort-Object -Property Student, Mission)) {
                                                if (-not $byStudent.Contains($c.Student)) { $byStudent[$c.Student] = New-Object System.Collections.Generic.List[object] }
                                                $byStudent[$c.Student].Add($c)
                                }
                                $count = { param($kind) @($changes | Where-Object { $_.Kind -eq $kind }).Count }
                                $lines.Add(("Changes for {0} student(s): {1} event(s) marked complete, {2} marked NOT complete, {3} IP / ETD updated." -f $byStudent.Count, (& $count 'Complete'), (& $count 'NotComplete'), (& $count 'Updated')))
                                $shown = 0
                                foreach ($name in $byStudent.Keys) {
                                                if ($shown -ge $MaxLines) {
                                                                $lines.Add("  ...and $($byStudent.Count - $shown) more student(s) (the console lists them all)")
                                                                break
                                                }
                                                $mine = $byStudent[$name].ToArray()
                                                $parts = @()
                                                $done = @($mine | Where-Object { $_.Kind -eq 'Complete' } | ForEach-Object { $_.Mission })
                                                if ($done.Count -gt 0) { $parts += "complete $(& $events $done)" }
                                                $undone = @($mine | Where-Object { $_.Kind -eq 'NotComplete' } | ForEach-Object { if ($_.Status) { "$($_.Mission) ($($_.Status))" } else { $_.Mission } })
                                                if ($undone.Count -gt 0) { $parts += "NOT complete $(& $events $undone)" }
                                                $updated = @($mine | Where-Object { $_.Kind -eq 'Updated' } | ForEach-Object { $_.Mission })
                                                if ($updated.Count -gt 0) { $parts += "IP / ETD updated $(& $events $updated)" }
                                                $debt = 0.0
                                                foreach ($c in $mine) { $debt += [double]$c.SimDebt }
                                                if ([math]::Abs($debt) -ge 0.005) { $parts += ("SimDebt {0:+0.##;-0.##}" -f $debt) }
                                                $lines.Add("  ${name}: $($parts -join '; ')")
                                                $shown++
                                }
                }

                $alone = New-Object System.Collections.Generic.List[string]
                if (@($Plan.Planned).Count -gt 0) { $alone.Add("Left as planned: $(@($Plan.Planned).Count) event(s) FAIS has scheduled after now, which the Big Board can't show yet.") }
                if (@($Plan.NotInSyllabus).Count -gt 0) { $alone.Add("Not in the student's syllabus: $(& $capped @($Plan.NotInSyllabus))") }
                if (@($Plan.Ambiguous).Count -gt 0) { $alone.Add("Could be more than one person: $(& $capped @($Plan.Ambiguous))") }
                if (@($Plan.Unknown).Count -gt 0) {
                                if ($File.Kind -eq 'BigBoard') { $alone.Add("Not FAIS students: $(& $capped @($Plan.Unknown))") }
                                else { $alone.Add("$(@($Plan.Unknown).Count) other people on the schedule are not FAIS students.") }
                }
                if (@($Plan.NotOnFile).Count -gt 0) { $alone.Add("FAIS students not on the Big Board: $(& $capped @($Plan.NotOnFile))") }
                if ($alone.Count -gt 0) {
                                $lines.Add('')
                                $lines.Add('Left alone:')
                                foreach ($a in $alone) { $lines.Add("  $a") }
                }
                return ($lines.ToArray() -join "`n")
}
