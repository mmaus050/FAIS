# ASSESSMENT.PSM1 DIRECTORY:
#             "What if this student goes on this line?" for the puck board.
#             Runs the real hard rules and soft rules against the board's other pins.
# Clear-AssessmentCache
# Get-BoardEvalRows / Get-BoardStudentContext
# Get-PlacementAssessment
# Get-AssessmentLines (text for tooltips / popups)
# Get-MissionChoiceTexts / ConvertFrom-MissionChoiceText / Get-MissionAdvice (the mission dropdown)
# Get-StudentPuckInfo / Get-InstructorPuckInfo


$script:AssessCache = @{
                Version  = -1
                BoardKey = $null
                Rows     = $null
                Lines    = @{}
                Results  = @{}
}

function Clear-AssessmentCache {
                $script:AssessCache.Version = -1
                $script:AssessCache.BoardKey = $null
                $script:AssessCache.Rows = $null
                $script:AssessCache.Lines = @{}
                $script:AssessCache.Results = @{}
}

function Sync-AssessmentCache {
                param ($Board)

                $key = [System.Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($Board)
                if ($script:AssessCache.Version -ne $Board.Version -or $script:AssessCache.BoardKey -ne $key) {
                                Clear-AssessmentCache
                                $script:AssessCache.Version = $Board.Version
                                $script:AssessCache.BoardKey = $key
                }
}

# Board lines as engine rows with instructor objects resolved (cached per board version)
function Get-BoardEvalRows {
                param ($Board, $Data)

                Sync-AssessmentCache $Board
                if ($null -eq $script:AssessCache.Rows) {
                                $rows = @(ConvertTo-ScheduleRows -Board $Board)
                                foreach ($r in $rows) { $r.Instructor = Resolve-SortieInstructor -Instructors $Data.Instructors -Instructor $r.Instructor }
                                $script:AssessCache.Rows = $rows
                }
                return $script:AssessCache.Rows
}

# A copy of the student as they would be on -TargetRow's day: their other pins earlier
# than the target are flown (marked complete), later ones hold the time as reservations.
function Get-BoardStudentContext {
                param (
                                [array]$Rows,
                                $Student,
                                $TargetRow,
                                [string[]]$ExcludeLineIds = @(),
                                $Data
                )

                $clone = Copy-StudentForEval $Student
                $targetStart = Parse-ETD $TargetRow.ETD
                $pins = @($Rows | Where-Object {
                                (Get-NameText $_.Student) -ieq [string]$Student.Name -and
                                $_.BoardLineId -ne $TargetRow.BoardLineId -and
                                $ExcludeLineIds -notcontains $_.BoardLineId
                } | Sort-Object -Property { Parse-ETD $_.ETD })

                foreach ($pin in $pins) {
                                $start = Parse-ETD $pin.ETD
                                $ipName = Get-NameText $pin.Instructor
                                $isOB = Test-IsOBSortieType ([string]$pin.SortieType)
                                $ret = if ($isOB) { Get-LinkedReturnRow -Sortie $pin -Schedule $Rows -StudentName $Student.Name } else { $null }

                                if ($null -ne $start -and $null -ne $targetStart -and $start -lt $targetStart) {
                                                # The same pick the fill makes: explicit mission, else the best-scoring event that passes
                                                $pick = Select-PinnedMission -Student $clone -Sortie $pin -Rows $Rows -Students $Data.Students -Instructors $Data.Instructors -SoftWeights $Data.SoftWeights `
                                                                -FormLookAhead (Get-Weight $Data.SoftWeights 'FormLookAhead' 22)
                                                $m = $pick.Mission
                                                if ($null -ne $m -and -not (Test-MissionComplete $m)) {
                                                                [void](Complete-StudentMission -Student $clone -Mission $m -InstructorName $ipName -ETD $pin.ETD -SoftWeights $Data.SoftWeights -SortieType ([string]$pin.SortieType))
                                                }
                                                else {
                                                                # No event could be flown: the student still occupies that time
                                                                Add-StudentReservation -Student $clone -LineId ([string]$pin.BoardLineId) -ETD $pin.ETD -SortieType ([string]$pin.SortieType) -Mission (Get-NameText $pin.Mission) -Instructor $ipName
                                                }
                                                if ($null -ne $ret) {
                                                                $rm = @(Get-ValidOptedEvents -Student $clone -Sortie $ret -NoLookahead) | Select-Object -First 1
                                                                if ($null -ne $m -and $null -ne $rm) { [void](Complete-StudentMission -Student $clone -Mission $rm -InstructorName $ipName -ETD $ret.ETD -SoftWeights $Data.SoftWeights -SortieType ([string]$ret.SortieType)) }
                                                                else { Add-StudentReservation -Student $clone -LineId ([string]$ret.BoardLineId) -ETD $ret.ETD -SortieType ([string]$ret.SortieType) -Instructor $ipName }
                                                }
                                }
                                else {
                                                Add-StudentReservation -Student $clone -LineId ([string]$pin.BoardLineId) -ETD $pin.ETD -SortieType ([string]$pin.SortieType) -Mission (Get-NameText $pin.Mission) -Instructor $ipName
                                                if ($null -ne $ret) {
                                                                Add-StudentReservation -Student $clone -LineId ([string]$ret.BoardLineId) -ETD $ret.ETD -SortieType ([string]$ret.SortieType) -Instructor $ipName
                                                }
                                }
                }
                return $clone
}

function Get-AssessmentStatus {
                param ($StudentViolations, $Chosen, $Notes)

                foreach ($v in @($StudentViolations)) { if ($v.Severity -eq 'Hard') { return 'Hard' } }
                if ($null -eq $Chosen) { return 'Hard' }
                foreach ($v in @($Chosen.Violations)) { if ($v.Severity -eq 'Hard') { return 'Hard' } }
                if (@($StudentViolations).Count -gt 0 -or @($Chosen.Violations).Count -gt 0 -or @($Notes).Count -gt 0) { return 'Warn' }
                return 'Clean'
}

# Evaluates one student on one board line.
# Returns {Status 'Clean'|'Warn'|'Hard'|'Error', Violations, Notes, AutoMission, Mission, Score,
#          BestScore, EligibleEvents, EventDetails, Summary, LineId, StudentName, BoardVersion}
function Get-PlacementAssessment {
                param (
                                [Parameter(Mandatory)] $Board,
                                [Parameter(Mandatory)] [string]$LineId,
                                [Parameter(Mandatory)] [string]$StudentName,
                                [string]$MissionName = '',
                                [string[]]$ExcludeLineIds = @(),
                                [Parameter(Mandatory)] $Data,
                                [switch]$Quick
                )

                Sync-AssessmentCache $Board
                $cacheKey = "$LineId|$StudentName|$MissionName|$([bool]$Quick)|$($ExcludeLineIds -join ',')"
                if ($script:AssessCache.Results.ContainsKey($cacheKey)) { return $script:AssessCache.Results[$cacheKey] }

                $result = [PSCustomObject]@{
                                LineId         = $LineId
                                StudentName    = $StudentName
                                BoardVersion   = $Board.Version
                                Status         = 'Error'
                                Violations     = @()
                                Notes          = @()
                                AutoMission    = ''
                                Mission        = ''
                                Score          = $null
                                BestScore      = $null
                                EligibleEvents = @()
                                EventDetails   = @()
                                Summary        = ''
                                Quick          = [bool]$Quick
                                RequestedMission = $MissionName
                }

                $line = Get-BoardLine $Board $LineId
                if ($null -eq $line) { $result.Summary = 'Line not found'; return $result }
                if (Test-IsReturnSortieType ([string]$line.SortieType)) {
                                $result.Status = 'Hard'
                                $result.Violations = @(New-RuleViolation 'RETURN_LINE' 'O&B Return lines are filled from their O&B')
                                $result.Summary = 'Return lines follow their O&B'
                                return $result
                }
                $student = Resolve-StudentByName -Students $Data.Students -Name $StudentName
                if ($null -eq $student) {
                                $result.Status = 'Hard'
                                $result.Violations = @(New-RuleViolation 'UNKNOWN_STUDENT' "'$StudentName' is not in MasterData")
                                $result.Summary = $result.Violations[0].Message
                                return $result
                }

                $rows = Get-BoardEvalRows -Board $Board -Data $Data
                $target = $null
                foreach ($r in $rows) { if ($r.BoardLineId -eq $LineId) { $target = $r; break } }
                $probe = $target.PSObject.Copy()
                $probe.Student = $null
                $probe.Mission = $null

                $clone = Get-BoardStudentContext -Rows $rows -Student $student -TargetRow $probe -ExcludeLineIds $ExcludeLineIds -Data $Data
                $formLookAhead = Get-Weight $Data.SoftWeights 'FormLookAhead' 22
                if (-not $script:AssessCache.Lines.ContainsKey($LineId)) { $script:AssessCache.Lines[$LineId] = @{} }
                $lineCtx = $script:AssessCache.Lines[$LineId]
                $isWing = Test-IsPairedWing -Sortie $probe -Schedule $rows
                $missionArg = if ($MissionName -ne '') { $MissionName } else { $null }

                $hr = Test-HardRules -Student $clone -Sortie $probe -Schedule $rows -Students $Data.Students -Instructors $Data.Instructors `
                                -FormLookAhead $formLookAhead -Mission $missionArg -IsWingCandidate:$isWing -Explain:(-not $Quick) -LineContext $lineCtx

                $details = New-Object System.Collections.Generic.List[object]
                $best = $null
                foreach ($er in @($hr.EventResults)) {
                                $d = [PSCustomObject]@{
                                                Mission    = $er.Mission
                                                Type       = [string]$er.Event.Type
                                                Passed     = $er.Passed
                                                Score      = $null
                                                Breakdown  = @()
                                                Violations = @($er.Violations)
                                                Notes      = @($er.Notes)
                                }
                                if (-not $Quick) {
                                                $sr = Get-EventScore -Student $clone -Sortie $probe -Evt $er.Event -SoftWeights $Data.SoftWeights -Instructors $Data.Instructors
                                                $d.Score = [math]::Round([double]$sr.Score, 1)
                                                $d.Breakdown = @(foreach ($k in $sr.Breakdown.Keys) { [PSCustomObject]@{ Label = $k; Points = [math]::Round([double]$sr.Breakdown[$k], 1); Text = [string]$sr.Reasons[$k] } })
                                                $d.Notes = @(@($er.Notes) + @($sr.Notes) | Select-Object -Unique)
                                                if ($er.Passed -and (-not $isWing -or $d.Type -like '*Form*')) {
                                                                $cand = [PSCustomObject]@{ Student = $student; Event = $er.Event; Score = $sr.Score; Detail = $d }
                                                                if (Test-CandidateBetter $cand $best) { $best = $cand }
                                                }
                                }
                                elseif ($er.Passed -and $null -eq $best) {
                                                $best = [PSCustomObject]@{ Detail = $d }
                                }
                                $details.Add($d)
                }

                $chosen = $null
                if ($MissionName -ne '') { $chosen = $details | Where-Object { $_.Mission -ieq $MissionName } | Select-Object -First 1 }
                elseif ($null -ne $best) { $chosen = $best.Detail }
                elseif ($details.Count -gt 0) { $chosen = $details[0] }

                $notes = @(@($hr.Notes) + @(if ($null -ne $chosen) { $chosen.Notes }) | Where-Object { $_ } | Select-Object -Unique)
                $violations = @(@($hr.Violations) + @(if ($null -ne $chosen) { $chosen.Violations }) | Where-Object { $null -ne $_ })

                $result.Status = Get-AssessmentStatus -StudentViolations $hr.Violations -Chosen $chosen -Notes $notes
                $result.Violations = $violations
                $result.Notes = $notes
                $result.AutoMission = if ($null -ne $best) { [string]$best.Detail.Mission } else { '' }
                $result.Mission = if ($null -ne $chosen) { [string]$chosen.Mission } else { '' }
                $result.Score = if ($null -ne $chosen) { $chosen.Score } else { $null }
                $result.BestScore = if ($null -ne $best) { $best.Detail.Score } else { $null }
                $result.EventDetails = $details.ToArray()
                if (-not $Quick) {
                                $result.EligibleEvents = @(Get-ValidOptedEvents -Student $clone -Sortie $probe | ForEach-Object { [string]$_.Mission })
                }

                $flagCount = @($violations | Where-Object { $_.Severity -eq 'Hard' }).Count
                $scoreText = if ($null -ne $result.Score) { ' {0:+0.0;-0.0;0}' -f $result.Score } else { '' }
                $flagText = if ($flagCount -eq 1) { ', 1 rule flag' } elseif ($flagCount -gt 1) { ", $flagCount rule flags" } else { '' }
                $result.Summary = "{0} -> {1} {2}: {3}{4}{5}" -f $student.Name, (Format-BoardTime $line.ETD), $line.SortieType, $(if ($result.Mission) { $result.Mission } else { 'no event' }), $scoreText, $flagText

                $script:AssessCache.Results[$cacheKey] = $result
                return $result
}

# Assessment as display lines: {Kind 'Title'|'Hard'|'Warn'|'Note'|'Plus'|'Minus'|'Info'; Text}
function Get-AssessmentLines {
                param ($Assessment)

                $out = New-Object System.Collections.Generic.List[object]
                if ($null -eq $Assessment) { return }
                $out.Add([PSCustomObject]@{ Kind = 'Title'; Text = $Assessment.Summary })

                foreach ($v in @($Assessment.Violations)) {
                                $kind = if ($v.Severity -eq 'Hard') { 'Hard' } else { 'Warn' }
                                $out.Add([PSCustomObject]@{ Kind = $kind; Text = $v.Message })
                }
                foreach ($n in @($Assessment.Notes)) { $out.Add([PSCustomObject]@{ Kind = 'Note'; Text = $n }) }

                $chosen = $Assessment.EventDetails | Where-Object { $_.Mission -eq $Assessment.Mission } | Select-Object -First 1
                if ($null -ne $chosen) {
                                foreach ($b in @($chosen.Breakdown | Sort-Object -Property { [math]::Abs($_.Points) } -Descending)) {
                                                # The descriptive reason when the soft rules gave one, else the rule's label
                                                $label = [string](Get-OptionalProperty $b 'Text' '')
                                                if ($label -eq '') { $label = $b.Label }
                                                if ($b.Points -gt 0) { $out.Add([PSCustomObject]@{ Kind = 'Plus'; Text = ('{0} +{1:0.#}' -f $label, $b.Points) }) }
                                                elseif ($b.Points -lt 0) { $out.Add([PSCustomObject]@{ Kind = 'Minus'; Text = ('{0} {1:0.#}' -f $label, $b.Points) }) }
                                }
                }

                $others = @($Assessment.EventDetails | Where-Object { $_.Mission -ne $Assessment.Mission })
                if ($others.Count -gt 0) {
                                $parts = foreach ($o in $others) {
                                                if ($o.Passed -and $null -ne $o.Score) { '{0} {1:0.0}' -f $o.Mission, $o.Score }
                                                elseif ($o.Passed) { "$($o.Mission) ok" }
                                                else {
                                                                $why = @($o.Violations | Select-Object -First 1 | ForEach-Object { $_.Message })
                                                                "$($o.Mission) blocked ($($why -join ''))"
                                                }
                                }
                                $out.Add([PSCustomObject]@{ Kind = 'Info'; Text = 'Other events: ' + ($parts -join '; ') })
                }
                if (-not $Assessment.Quick) {
                                $out.Add([PSCustomObject]@{ Kind = 'Info'; Text = 'Based on pinned lines only; auto-filled lines can change turns.' })
                }
                return $out.ToArray()
}

# =================================================================================
#  CHOOSING A MISSION BY HAND
#  The mission dropdown offers the events the student can fly on the line, then the
#  next ones they aren't opted for yet, then "Type a mission..." for anything else.
#  A choice that breaks a rule is checked with Get-MissionAdvice before it goes on.
# =================================================================================

$script:MissionChoice = @{
                Auto      = '(auto)'
                Type      = 'Type a mission...'
                NotOpted  = ' (not opted)'
}

# The dropdown's entries (text) for a hardcoded student on a line of -SortieType.
# -Eligible: events they can fly there; -Current: the mission on the line now.
function Get-MissionChoiceTexts {
                param (
                                $Student,
                                [string]$SortieType = '',
                                [string[]]$Eligible = @(),
                                [string]$Current = '',
                                [int]$MaxNotOpted = 8
                )

                $out = New-Object System.Collections.Generic.List[string]
                $have = @{}
                $out.Add($script:MissionChoice.Auto)
                foreach ($e in $Eligible) {
                                $name = ([string]$e).Trim()
                                if ($name -eq '' -or $have.ContainsKey($name.ToUpperInvariant())) { continue }
                                $out.Add($name)
                                $have[$name.ToUpperInvariant()] = $true
                }
                $cur = ([string]$Current).Trim()
                if ($cur -ne '' -and -not $have.ContainsKey($cur.ToUpperInvariant())) {
                                $out.Add($cur)
                                $have[$cur.ToUpperInvariant()] = $true
                }

                # The next events of the line's type the student isn't opted for (prereqs still to fly)
                if ($null -ne $Student -and $MaxNotOpted -gt 0) {
                                $types = Get-AllowedEventTypes $SortieType
                                $added = 0
                                foreach ($m in @($Student.Syllabus | Where-Object { $null -ne $_ } | Sort-Object -Property { Get-SyllabusSortKey $_ })) {
                                                if ($added -ge $MaxNotOpted) { break }
                                                $name = ([string]$m.Mission).Trim()
                                                if ($name -eq '' -or $have.ContainsKey($name.ToUpperInvariant()) -or (Test-MissionComplete $m)) { continue }
                                                if ($null -ne $types -and ([string]$m.Type).Trim() -notin $types) { continue }
                                                if ((Get-NotOptedReason -Student $Student -Mission $m) -eq '') { continue }
                                                $out.Add($name + $script:MissionChoice.NotOpted)
                                                $have[$name.ToUpperInvariant()] = $true
                                                $added++
                                }
                }
                $out.Add($script:MissionChoice.Type)
                return $out.ToArray()
}

# What a dropdown entry means: {Kind 'Auto'|'Type'|'Mission'; Mission}
function ConvertFrom-MissionChoiceText {
                param ([string]$Text)

                $t = ([string]$Text).Trim()
                if ($t -eq '' -or $t -eq $script:MissionChoice.Auto) { return [PSCustomObject]@{ Kind = 'Auto'; Mission = '' } }
                if ($t -eq $script:MissionChoice.Type) { return [PSCustomObject]@{ Kind = 'Type'; Mission = '' } }
                if ($t.EndsWith($script:MissionChoice.NotOpted)) { $t = $t.Substring(0, $t.Length - $script:MissionChoice.NotOpted.Length).Trim() }
                return [PSCustomObject]@{ Kind = 'Mission'; Mission = $t }
}

# Whether a mission picked by hand needs a second look before it goes on, and what to fly instead.
# -Requested: the assessment with that mission; -Auto: the assessment with (auto).
# Problems are the rule flags that come from the mission itself (not opted, not in the syllabus,
# wrong line, a checkride rule...), not ones any mission would have (turn time, absence...).
# Returns {NeedsConfirm; Problems; Suggested; SuggestedScore; Others; Mission; Title; Text; Buttons}
function Get-MissionAdvice {
                param ($Requested, $Auto, [Parameter(Mandatory)] [string]$Mission)

                $missionRules = @('NOT_OPTED', 'MISSION_NOT_ELIGIBLE', 'MISSION_NOT_FOUND')
                $hard = @($Requested.Violations | Where-Object { $null -ne $_ -and $_.Severity -eq 'Hard' -and ($missionRules -contains $_.Rule -or ([string]$_.Mission) -ne '') })
                $problems = @($hard | ForEach-Object { [string]$_.Message } | Select-Object -Unique)
                $name = if ($null -ne $Requested -and $Requested.StudentName) { [string]$Requested.StudentName } else { 'This student' }
                $first = (($name -split ',')[0]).Trim()

                $suggested = ''
                $score = $null
                if ($null -ne $Auto -and -not [string]::IsNullOrWhiteSpace([string]$Auto.AutoMission) -and [string]$Auto.AutoMission -ine $Mission) {
                                $suggested = [string]$Auto.AutoMission
                                $score = $Auto.BestScore
                }
                $others = @()
                if ($null -ne $Auto) { $others = @(@($Auto.EligibleEvents) | Where-Object { $_ -and $_ -ine $suggested -and $_ -ine $Mission } | Select-Object -First 5) }

                $headline = "$Mission breaks a rule for $name on this line."
                if ($problems.Count -eq 0) { $headline = "$Mission is fine for $name on this line." }
                elseif (@($hard | Where-Object { $_.Rule -eq 'NOT_OPTED' }).Count -gt 0) { $headline = "$name isn't opted for $Mission." }
                elseif (@($hard | Where-Object { $_.Rule -eq 'MISSION_NOT_FOUND' }).Count -gt 0) { $headline = "$Mission isn't in $name's syllabus." }
                elseif (@($hard | Where-Object { $_.Rule -eq 'MISSION_NOT_ELIGIBLE' }).Count -gt 0) { $headline = "$name can't fly $Mission on this line." }

                $lines = New-Object System.Collections.Generic.List[string]
                $lines.Add($headline)
                $lines.Add('')
                foreach ($p in $problems) { $lines.Add("  X  $p") }
                $lines.Add('')
                if ($suggested -ne '') {
                                $scoreText = if ($null -ne $score) { ' (score {0:+0.0;-0.0;0})' -f [double]$score } else { '' }
                                $lines.Add("Recommended: $suggested, the mission the scheduler would put here$scoreText.")
                }
                else {
                                $lines.Add("There is no mission $first can fly on this line right now.")
                }
                if ($others.Count -gt 0) { $lines.Add("Other missions $first can fly here: $($others -join ', ').") }
                $lines.Add('')
                if ($suggested -ne '') {
                                $lines.Add("Yes = put $suggested on the line (recommended)")
                                $lines.Add("No = keep $Mission anyway (it is flagged OVERRIDE on the output)")
                }
                else {
                                $lines.Add("OK = keep $Mission anyway (it is flagged OVERRIDE on the output)")
                }
                $lines.Add('Cancel = leave the mission as it was')

                return [PSCustomObject]@{
                                NeedsConfirm   = ($problems.Count -gt 0)
                                Problems       = $problems
                                Suggested      = $suggested
                                SuggestedScore = $score
                                Others         = $others
                                Mission        = $Mission
                                Title          = "Check the mission for $name"
                                Text           = ($lines -join "`n")
                                Buttons        = $(if ($suggested -ne '') { 'YesNoCancel' } else { 'OKCancel' })
                }
}

# Badges and counts for a student puck on one day
function Get-StudentPuckInfo {
                param ($Student, $Board, [string]$Day)

                $badges = New-Object System.Collections.Generic.List[string]
                if (ConvertTo-FaisBool $Student.IsCap) { $badges.Add('CAP') }
                if (ConvertTo-FaisBool $Student.IsInternational) { $badges.Add('INTL') }
                if (ConvertTo-FaisBool $Student.OnHold) { $badges.Add('HOLD') }
                if (ConvertTo-FaisBool $Student.UnsatSim) { $badges.Add('U-SIM') }
                if (ConvertTo-FaisBool $Student.UnsatFlight) { $badges.Add('U-FLT') }

                $absText = New-Object System.Collections.Generic.List[string]
                $dayDate = $null
                if ($Day) { $dayDate = [datetime]::ParseExact($Day, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture) }
                foreach ($a in @($Student.Absences)) {
                                if ($null -eq $a -or $null -eq $dayDate) { continue }
                                if ($a.StartTime.Date -le $dayDate -and $a.EndTime.Date -ge $dayDate) {
                                                $s = if ($a.StartTime.Date -lt $dayDate) { '0000' } else { $a.StartTime.ToString('HHmm') }
                                                $e = if ($a.EndTime.Date -gt $dayDate) { '2359' } else { $a.EndTime.ToString('HHmm') }
                                                $badges.Add("ABS $s-$e")
                                                $absText.Add("$s-$e $($a.Reason)".Trim())
                                }
                }

                $turns = 0
                if ($null -ne $Board) {
                                foreach ($l in $Board.Lines) {
                                                if ($l.Day -ne $Day -or [string]::IsNullOrWhiteSpace($l.Student)) { continue }
                                                if ($l.Student -ine [string]$Student.Name) { continue }
                                                $turns += $(if (Test-IsOBSortieType ([string]$l.SortieType)) { 2 } else { 1 })
                                }
                }

                $opted = @(Get-OptedEvents -SyllabusData $Student.Syllabus | Sort-Object -Property { Get-SyllabusSortKey $_ })
                $next = if ($opted.Count -gt 0) { ((@($opted | Select-Object -First 2 | ForEach-Object { $_.Mission })) -join ' / ') } else { 'complete' }
                $debt = 0.0
                if ($null -ne $Student.SimDebt -and [string]$Student.SimDebt -ne '') { [void][double]::TryParse([string]$Student.SimDebt, [ref]$debt) }

                return [PSCustomObject]@{
                                Name     = [string]$Student.Name
                                Badges   = $badges.ToArray()
                                Absences = $absText.ToArray()
                                Turns    = $turns
                                Next     = $next
                                SimDebt  = $debt
                                Blocked  = ((ConvertTo-FaisBool $Student.OnHold))
                }
}

# Badges and line count for an IP puck on one day
function Get-InstructorPuckInfo {
                param ($Instructor, $Board, [string]$Day)

                $badges = New-Object System.Collections.Generic.List[string]
                if (ConvertTo-FaisBool $Instructor.IsExperienced) { $badges.Add('EXP') }
                if (ConvertTo-FaisBool $Instructor.IsCheck) { $badges.Add('CHK') }
                if (-not (ConvertTo-FaisBool $Instructor.IsCounter -Default $true)) { $badges.Add('NC') }

                $count = 0
                $times = New-Object System.Collections.Generic.List[string]
                if ($null -ne $Board) {
                                foreach ($l in (Get-BoardDayLines -Board $Board -Day $Day)) {
                                                if ([string]$l.Instructor -ieq [string]$Instructor.Name -and -not (Test-IsReturnSortieType ([string]$l.SortieType))) {
                                                                $count++
                                                                $times.Add($l.ETD.ToString('HHmm'))
                                                }
                                }
                }
                return [PSCustomObject]@{
                                Name     = [string]$Instructor.Name
                                Callsign = ("{0} {1}" -f $Instructor.Callsign, $Instructor.Num).Trim()
                                Badges   = $badges.ToArray()
                                Lines    = $count
                                Times    = $times.ToArray()
                }
}
