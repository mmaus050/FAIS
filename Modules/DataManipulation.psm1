# DATAMANIPULATION.PSM1 DIRECTORY:
#             Shared helpers that change, move, match and format data during the program run
# ConvertTo-FaisBool / Get-NameText / Get-BaseName
# Parse-ETD / Format-EtdShort / Format-EventLabel
# New-FallbackInstructor / Get-MatchedInstructor / Resolve-SortieInstructor
# Get-NormalizedLastName / Resolve-StudentByName / Test-IsSentinelStudent
# Test-IsSimSortieType / Test-IsReturnSortieType / Test-IsOBSortieType / Get-SortieCategory
# Add-SortieNote / Convert-SortieObjectsToText / Test-SameSortie
# Copy-StudentForEval / Copy-StudentList


# =================================================================================
#  VALUE HELPERS
#  Excel hands us strings, numbers, booleans or blanks for the same column.
#  These keep every rule reading values the same way.
# =================================================================================

# Converts an Excel cell value to a real boolean. "FALSE" as text is truthy in
# PowerShell, so every flag column goes through here.
function ConvertTo-FaisBool {
                param (
                                $Value,
                                [bool]$Default = $false
                )

                if ($null -eq $Value) { return $Default }
                if ($Value -is [bool]) { return $Value }
                if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) { return ($Value -ne 0) }

                $text = ([string]$Value).Trim().ToLowerInvariant()
                if ($text -eq '') { return $Default }
                if ($text -in @('true', 'yes', 'y', '1', 'x', '-1')) { return $true }
                if ($text -in @('false', 'no', 'n', '0')) { return $false }
                return $Default
}

# Returns the display name of a Student / Instructor / Mission field whether it
# holds an object, a string or nothing.
function Get-NameText {
                param ($Value)

                if ($null -eq $Value) { return '' }
                if ($Value -is [string]) { return $Value.Trim() }
                $props = $Value.PSObject.Properties
                if ($null -ne $props['Name']) { return ([string]$Value.Name).Trim() }
                if ($null -ne $props['Mission']) { return ([string]$Value.Mission).Trim() }
                return ([string]$Value).Trim()
}

# Strips "[...]" tags from an instructor name ("Smith [EXP]" -> "Smith")
function Get-BaseName {
                param ([string]$Name)

                if ([string]::IsNullOrWhiteSpace($Name)) { return '' }
                return ($Name -replace '\s*\[.*?\]', '').Trim()
}


# =================================================================================
#  TIME HELPERS
# =================================================================================

# Safely converts DateTimes, Excel serial dates and normal date strings
function Parse-ETD {
                param ($rawDate)

                if ($null -eq $rawDate) { return $null }
                if ($rawDate -is [datetime]) { return $rawDate }
                if ([string]::IsNullOrWhiteSpace([string]$rawDate)) { return $null }

                $text = ([string]$rawDate).Trim()
                if ($text -match '^\d+(\.\d+)?$') {
                                return [datetime]::FromOADate([double]$text)
                }

                $parsed = $text -as [datetime]
                return $parsed
}

# 0730 style time for messages
function Format-EtdShort {
                param ($Date)

                $dt = Parse-ETD $Date
                if ($null -eq $dt) { return '----' }
                return $dt.ToString('HHmm')
}

# "0730 FT5103" or "Mon 0730 FT5103" when the event is on a different day
function Format-EventLabel {
                param (
                                $Start,
                                [string]$Mission,
                                [switch]$WithDay
                )

                $dt = Parse-ETD $Start
                $time = if ($null -eq $dt) { '----' } elseif ($WithDay) { $dt.ToString('ddd HHmm') } else { $dt.ToString('HHmm') }
                if ([string]::IsNullOrWhiteSpace($Mission)) { return $time }
                return "$time $Mission"
}


# =================================================================================
#  INSTRUCTORS
# =================================================================================

# The stand-in object used when a line's IP is blank or not in MasterData
function New-FallbackInstructor {
                param ([string]$Name)

                if ([string]::IsNullOrWhiteSpace($Name)) { $Name = 'NEED IP' }
                return [PSCustomObject]@{
                                Name          = $Name
                                IsCounter     = $true
                                IsExperienced = $false
                                IsCheck       = $false
                                Callsign      = ''
                                Num           = ''
                                IsFallback    = $true
                }
}

# Finds an instructor by name. Exact (tag-stripped, case-insensitive) match first,
# then callsign, then the legacy partial match.
function Get-MatchedInstructor {
                param (
                                [array]$Instructors,
                                [string]$TargetName
                )

                if ([string]::IsNullOrWhiteSpace($TargetName) -or $null -eq $Instructors) {
                                return $null
                }

                try {
                                $BaseTarget = Get-BaseName $TargetName

                                foreach ($ip in $Instructors) {
                                                if ((Get-BaseName ([string]$ip.Name)) -ieq $BaseTarget) { return $ip }
                                }
                                foreach ($ip in $Instructors) {
                                                if (-not [string]::IsNullOrWhiteSpace([string]$ip.Callsign) -and ([string]$ip.Callsign).Trim() -ieq $BaseTarget) { return $ip }
                                }
                                foreach ($ip in $Instructors) {
                                                $BaseSource = Get-BaseName ([string]$ip.Name)
                                                if ($BaseSource -match [regex]::Escape($BaseTarget)) { return $ip }
                                }
                                return $null
                } catch {
                                Write-Warning "Error matching instructor: $($_.Exception.Message)"
                                return $null
                }
}

# Turns a line's Instructor field (object, name or blank) into an instructor object.
# Unknown or blank names get the fallback object so rules never read $null.
function Resolve-SortieInstructor {
                param (
                                [array]$Instructors,
                                $Instructor
                )

                if ($null -ne $Instructor -and $Instructor -isnot [string] -and $null -ne $Instructor.PSObject.Properties['Name']) {
                                return $Instructor
                }

                $name = Get-NameText $Instructor
                if ([string]::IsNullOrWhiteSpace($name)) { return New-FallbackInstructor 'NEED IP' }
                if ($name -like '*SIM IP*') { return New-FallbackInstructor $name }

                $match = Get-MatchedInstructor -Instructors $Instructors -TargetName $name
                if ($null -ne $match) { return $match }
                return New-FallbackInstructor $name
}


# =================================================================================
#  STUDENTS
# =================================================================================

# Returns only the last-name portion of a name string (everything before the first comma)
function Get-NormalizedLastName {
                param ([string]$Name)

                if ([string]::IsNullOrWhiteSpace($Name)) { return '' }
                $normalized = $Name.Trim()
                $commaIndex = $normalized.IndexOf(',')
                if ($commaIndex -ge 0) {
                                $normalized = $normalized.Substring(0, $commaIndex).Trim()
                }
                return $normalized
}

# Finds a student by exact name first, then by a last name that only one student has
function Resolve-StudentByName {
                param (
                                [array]$Students,
                                $Name
                )

                $text = Get-NameText $Name
                if ([string]::IsNullOrWhiteSpace($text) -or $null -eq $Students) { return $null }

                foreach ($s in $Students) {
                                if (([string]$s.Name).Trim() -ieq $text) { return $s }
                }

                $last = Get-NormalizedLastName $text
                $found = @($Students | Where-Object { (Get-NormalizedLastName ([string]$_.Name)) -ieq $last })
                if ($found.Count -eq 1) { return $found[0] }
                if ($found.Count -gt 1) {
                                Write-Warning "Student name '$text' matches more than one student ($(@($found | ForEach-Object { $_.Name }) -join '; ')). Use the full name."
                }
                return $null
}

# Placeholder text the engine writes into the Student column
function Test-IsSentinelStudent {
                param ($Value)

                $text = Get-NameText $Value
                if ([string]::IsNullOrWhiteSpace($text)) { return $true }
                return ($text -in @('For Sale', 'Need Form SP', 'NEED IP', 'SIM IP'))
}

# Deep-enough copy of a student so rules can be run "what if" without touching the original.
# Syllabus entries and reservations are copied; absences are shared (never mutated).
function Copy-StudentForEval {
                param ($Student)

                if ($null -eq $Student) { return $null }

                $copy = $Student.PSObject.Copy()
                $copy.Syllabus = @(foreach ($m in @($Student.Syllabus)) { if ($null -ne $m) { $m.PSObject.Copy() } })

                $res = @()
                if ($null -ne $Student.PSObject.Properties['Reservations']) {
                                $res = @(foreach ($r in @($Student.Reservations)) { if ($null -ne $r) { $r.PSObject.Copy() } })
                }
                $copy | Add-Member -MemberType NoteProperty -Name 'Reservations' -Value $res -Force
                $rev = if ($null -ne $Student.PSObject.Properties['SyllabusRevision']) { [int]$Student.SyllabusRevision } else { 0 }
                $copy | Add-Member -MemberType NoteProperty -Name 'SyllabusRevision' -Value $rev -Force
                return $copy
}

function Copy-StudentList {
                param ([array]$Students)

                return @(foreach ($s in @($Students)) { if ($null -ne $s) { Copy-StudentForEval $s } })
}


# =================================================================================
#  SORTIES
# =================================================================================

function Test-IsSimSortieType {
                param ([string]$SortieType)
                return ($SortieType -like '*OFT*' -or $SortieType -like '*IFT*')
}

function Test-IsReturnSortieType {
                param ([string]$SortieType)
                return ($SortieType -like '*Return*')
}

# True for the outbound O&B leg only (not the return line)
function Test-IsOBSortieType {
                param ([string]$SortieType)
                return ($SortieType -match '(?i)O\s*&\s*B' -and -not (Test-IsReturnSortieType $SortieType))
}

function Get-SortieCategory {
                [CmdletBinding()]
                param(
                                [Parameter(ValueFromPipeline = $true)]
                                [string]$SortieType
                )
                process {
                                if ([string]::IsNullOrWhiteSpace($SortieType)) { return 'Unknown' }
                                if (Test-IsSimSortieType $SortieType) { return 'Sim' }
                                return 'Flight'
                }
}

# Safely appends notes to a Sortie without overwriting or leaving ugly formatting
function Add-SortieNote {
                param (
                                $Sortie,
                                [string]$Note
                )

                if ($null -eq $Sortie -or [string]::IsNullOrWhiteSpace($Note)) { return }
                $existing = [string]$Sortie.Notes
                if ([string]::IsNullOrWhiteSpace($existing)) {
                                $Sortie.Notes = $Note.Trim()
                } elseif (($existing -split '\s*\|\s*') -notcontains $Note.Trim()) {
                                $Sortie.Notes = "$($existing.Trim()) | $($Note.Trim())"
                }
}

# Converts Instructor, Mission and Student objects to their display strings.
# Mutates the rows in place, so only call it on copies meant for export.
function Convert-SortieObjectsToText {
                param(
                                [AllowEmptyCollection()]
                                [array]$Schedule
                )

                foreach ($Sortie in $Schedule) {
                                if ($null -ne $Sortie.Instructor -and $Sortie.Instructor -isnot [string]) { $Sortie.Instructor = Get-NameText $Sortie.Instructor }
                                if ($null -ne $Sortie.Mission -and $Sortie.Mission -isnot [string]) { $Sortie.Mission = Get-NameText $Sortie.Mission }
                                if ($null -ne $Sortie.Student -and $Sortie.Student -isnot [string]) { $Sortie.Student = Get-NameText $Sortie.Student }
                }
}

# Board rows are copied a lot, so compare by line id when both rows have one
function Test-SameSortie {
                param ($A, $B)

                if ($null -eq $A -or $null -eq $B) { return $false }
                $ida = if ($null -ne $A.PSObject.Properties['BoardLineId']) { [string]$A.BoardLineId } else { '' }
                $idb = if ($null -ne $B.PSObject.Properties['BoardLineId']) { [string]$B.BoardLineId } else { '' }
                if ($ida -ne '' -and $idb -ne '') { return ($ida -eq $idb) }
                return [object]::ReferenceEquals($A, $B)
}

# Reads an optional property without tripping over rows that do not have it
function Get-OptionalProperty {
                param (
                                $Object,
                                [string]$Name,
                                $Default = $null
                )

                if ($null -eq $Object) { return $Default }
                $p = $Object.PSObject.Properties[$Name]
                if ($null -eq $p) { return $Default }
                return $p.Value
}

# Adds the board bookkeeping properties to a schedule row if it does not have them
function Initialize-SortieRow {
                param ($Sortie)

                foreach ($name in @('BoardLineId', 'LinkedLineId', 'WingOfLineId', 'FilledBy')) {
                                if ($null -eq $Sortie.PSObject.Properties[$name]) {
                                                $Sortie | Add-Member -MemberType NoteProperty -Name $name -Value $null
                                }
                }
                return $Sortie
}
