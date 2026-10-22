# Synthetic MasterData-shaped objects for tests. Dot-source this file.
# Shapes match what IngestData returns.

$script:FaisRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent

function Import-FaisModules {
                param ([switch]$IncludeUI)

                $names = @('DataManipulation', 'ExcelFiles', 'Syllabus', 'DataDownload', 'SchedulingEngine', 'DataUpload', 'SyllabusImport', 'PreRunMenu', 'BoardModel', 'Assessment')
                if ($IncludeUI) { $names += 'PuckBoard' }
                foreach ($n in $names) {
                                $path = Join-Path $script:FaisRoot "Modules\$n.psm1"
                                $path = $path -replace '\\', [IO.Path]::DirectorySeparatorChar
                                if (Test-Path $path) { Import-Module $path -Force -DisableNameChecking -Global -ErrorAction Stop }
                }
}

# Ten-event syllabus with a prereq chain covering every event Type
function New-TestSyllabus {
                $rows = @(
                                @{ ID = 1;  Mission = 'FI3101'; Prereqs = '0';      Type = 'IFT';        IsCheck = $false; Level = 1 }
                                @{ ID = 2;  Mission = 'FE3102'; Prereqs = 'FI3101'; Type = 'OFT';        IsCheck = $false; Level = 1 }
                                @{ ID = 3;  Mission = 'FT5001'; Prereqs = 'FI3101'; Type = 'Single';     IsCheck = $false; Level = 1 }
                                @{ ID = 4;  Mission = 'FT5102'; Prereqs = 'FT5001'; Type = 'Single';     IsCheck = $false; Level = 2 }
                                @{ ID = 5;  Mission = 'FT5103'; Prereqs = 'FT5102'; Type = 'SingleSolo'; IsCheck = $false; Level = 2 }
                                @{ ID = 6;  Mission = 'FN3201'; Prereqs = 'FT5102'; Type = 'IFT';        IsCheck = $false; Level = 3 }
                                @{ ID = 7;  Mission = 'FF5201'; Prereqs = 'FN3201'; Type = 'Formation';  IsCheck = $false; Level = 3 }
                                @{ ID = 8;  Mission = 'FF5202'; Prereqs = 'FF5201'; Type = 'Formation';  IsCheck = $false; Level = 4 }
                                @{ ID = 9;  Mission = 'FT5104'; Prereqs = 'FT5103'; Type = 'Single';     IsCheck = $false; Level = 3 }
                                @{ ID = 10; Mission = 'FT5190'; Prereqs = 'FT5104'; Type = 'Single';     IsCheck = $true;  Level = 4 }
                )
                return @(foreach ($r in $rows) {
                                [PSCustomObject]@{
                                                SyllabusID = $r.ID
                                                Mission    = $r.Mission
                                                Prereqs    = $r.Prereqs
                                                Instructor = ''
                                                ETD        = ''
                                                Type       = $r.Type
                                                IsCheck    = $r.IsCheck
                                                IsComplete = $false
                                                NumOpting  = 0
                                                Level      = $r.Level
                                }
                })
}

# -Done: array of 'Mission|yyyy-MM-dd HH:mm|Instructor' strings (Instructor defaults to SIM IP)
function New-TestStudent {
                param (
                                [string]$Name,
                                [string[]]$Done = @(),
                                [switch]$IsCap,
                                [switch]$IsInternational,
                                [switch]$OnHold,
                                [switch]$UnsatSim,
                                [string]$UnsatSimType = '',
                                [switch]$UnsatFlight,
                                [double]$Priority = 0,
                                [double]$SimDebt = 0,
                                [array]$Absences = @()
                )

                $syl = New-TestSyllabus
                foreach ($spec in $Done) {
                                $d = @($spec -split '\|')
                                $m = $syl | Where-Object { $_.Mission -eq $d[0] } | Select-Object -First 1
                                if ($null -eq $m) { throw "Unknown test mission $($d[0])" }
                                $m.IsComplete = $true
                                $m.ETD = ([datetime]$d[1]).ToString('yyyy-MM-dd HH:mm:ss')
                                $m.Instructor = if ($d.Count -ge 3) { $d[2] } else { 'SIM IP' }
                }

                return [PSCustomObject]@{
                                ID              = $Name
                                Name            = $Name
                                IsCap           = [bool]$IsCap
                                IsInternational = [bool]$IsInternational
                                Priority        = $Priority
                                UnsatSim        = [bool]$UnsatSim
                                UnsatSimType    = $UnsatSimType
                                UnsatFlight     = [bool]$UnsatFlight
                                OnHold          = [bool]$OnHold
                                SimDebt         = $SimDebt
                                Syllabus        = $syl
                                Absences        = $Absences
                                Reservations    = @()
                }
}

function New-TestInstructors {
                return @(
                                [PSCustomObject]@{ Name = 'Jones, Bob';  IsCounter = $true;  IsExperienced = $true;  IsCheck = $true;  Callsign = 'TALON'; Num = 31; OSC = ''; OSNum = '' }
                                [PSCustomObject]@{ Name = 'Lee, Ann';    IsCounter = $true;  IsExperienced = $false; IsCheck = $false; Callsign = 'TALON'; Num = 32; OSC = ''; OSNum = '' }
                                [PSCustomObject]@{ Name = 'Leech, Tom';  IsCounter = $false; IsExperienced = $true;  IsCheck = $false; Callsign = 'TALON'; Num = 33; OSC = ''; OSNum = '' }
                                [PSCustomObject]@{ Name = 'Ortiz, Sam';  IsCounter = $true;  IsExperienced = $true;  IsCheck = $false; Callsign = 'TALON'; Num = 34; OSC = ''; OSNum = '' }
                )
}

function New-TestFormations {
                return @(
                                [PSCustomObject]@{ NumFormation = 1; Callsign = 'HAZE'; Num = 11 }
                                [PSCustomObject]@{ NumFormation = 2; Callsign = 'HAZE'; Num = 12 }
                                [PSCustomObject]@{ NumFormation = 3; Callsign = 'HAZE'; Num = 21 }
                                [PSCustomObject]@{ NumFormation = 4; Callsign = 'HAZE'; Num = 22 }
                )
}

function New-TestSoftWeights {
                return [PSCustomObject]@{
                                ContString        = 10
                                String4           = -10
                                String5           = -100
                                Zipper            = 5
                                NonCounter        = -15
                                BreakString       = -5
                                ExhaustionMult    = -1
                                ExhaustionDays    = 5
                                BIT_X             = 1.5
                                BIT_Y             = 6
                                CAPBonus          = 8
                                FormLookAhead     = 22
                                LevelDifference   = 2
                                CounterContinuity = 5
                                OFTIFTPen         = -7
                                DollarRideBonus   = 20
                                TransSimRatio     = 0.5
                                FormSimRatio      = 0.5
                                SimStopMult       = 1
                }
}

# Schedule row in the IngestData shape (plus optional board ids)
function New-TestRow {
                param (
                                [string]$ETD,
                                [string]$SortieType = 'Day Local',
                                [string]$Instructor = '',
                                $Student = $null,
                                $Mission = $null,
                                [string]$Id = '',
                                [string]$LinkedLineId = '',
                                [string]$WingOfLineId = '',
                                [string]$Notes = $null
                )

                if ([string]::IsNullOrWhiteSpace($Instructor) -and ($SortieType -like '*OFT*' -or $SortieType -like '*IFT*')) { $Instructor = 'SIM IP' }
                return [PSCustomObject]@{
                                ETD          = ([datetime]$ETD).ToString('yyyy-MM-dd HH:mm:ss')
                                Instructor   = $Instructor
                                SortieType   = $SortieType
                                Student      = $Student
                                Mission      = $Mission
                                Callsign     = $null
                                Num          = $null
                                Notes        = $Notes
                                BoardLineId  = $(if ($Id) { $Id } else { $null })
                                LinkedLineId = $(if ($LinkedLineId) { $LinkedLineId } else { $null })
                                WingOfLineId = $(if ($WingOfLineId) { $WingOfLineId } else { $null })
                                FilledBy     = $null
                }
}
