# EXCELFILES.PSM1 DIRECTORY:
#             Keeps Excel out of FAIS's way. A workbook FAIS reads is saved first, so edits
#             not saved yet are read too. A workbook FAIS writes is saved and closed first,
#             so the write never fails because Excel holds the file. Works through the
#             running Excel (COM) on Windows and quietly does nothing anywhere else.
# Test-FileWritable / Wait-FileWritable
# Get-FaisExcelApp / Invoke-FaisExcelCom / Test-FaisExcelBusyError
# Find-FaisOpenWorkbook / Close-FaisWorkbookObject / Save-FaisWorkbook
# Close-FaisExcelFiles / Request-FaisFileWritable / Open-FaisWorkbooks


# =================================================================================
#  IS THE FILE FREE?
# =================================================================================

# True when the file does not exist yet or can be opened for writing (not open in Excel)
function Test-FileWritable {
                param ([string]$Path)

                if (-not (Test-Path -LiteralPath $Path)) { return $true }
                try {
                                $fs = [System.IO.File]::Open($Path, 'Open', 'ReadWrite', 'None')
                                $fs.Close()
                                return $true
                } catch {
                                return $false
                }
}

# Waits for a file to be closed in Excel. -Prompt is a scriptblock that gets the path
# and returns $true to retry. Returns $false if the user gave up.
function Wait-FileWritable {
                param (
                                [string]$Path,
                                [scriptblock]$Prompt
                )

                while (-not (Test-FileWritable $Path)) {
                                $retry = $false
                                if ($null -ne $Prompt) { $retry = [bool](& $Prompt $Path) }
                                else {
                                                $answer = Read-Host "$Path is open (probably in Excel). Close it and press Enter to retry, or type N to skip"
                                                $retry = ($answer -notmatch '^(?i)n')
                                }
                                if (-not $retry) { return $false }
                }
                return $true
}

# Excel can take a moment to let go of a file it just closed
function Wait-FaisFileReleased {
                param ([string]$Path, [int]$TimeoutMs = 3000)

                $sw = [System.Diagnostics.Stopwatch]::StartNew()
                while (-not (Test-FileWritable $Path) -and $sw.ElapsedMilliseconds -lt $TimeoutMs) { Start-Sleep -Milliseconds 100 }
                return (Test-FileWritable $Path)
}


# =================================================================================
#  TALKING TO EXCEL
# =================================================================================

# The Excel the user has open, or $null (not Windows, Excel not running, or no COM)
function Get-FaisExcelApp {
                if ([System.Environment]::OSVersion.Platform -ne [System.PlatformID]::Win32NT) { return $null }
                try { return [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { return $null }
}

# Excel refuses calls while a cell is being edited or a dialog is open
function Test-FaisExcelBusyError {
                param ($Exception)

                # RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER, VBA_E_IGNORE
                $busy = @(-2147418111, -2147417846, -2146777998)
                $e = $Exception
                while ($null -ne $e) {
                                if ($busy -contains [int]$e.HResult) { return $true }
                                $e = $e.InnerException
                }
                return $false
}

# Runs one COM call, retrying for a few seconds while Excel is busy
function Invoke-FaisExcelCom {
                param ([scriptblock]$ComCall, [int]$ComTries = 10, [int]$ComDelayMs = 300)

                for ($try = 1; $try -le $ComTries; $try++) {
                                try { return (& $ComCall) }
                                catch {
                                                if ($try -ge $ComTries -or -not (Test-FaisExcelBusyError $_.Exception)) { throw }
                                                Start-Sleep -Milliseconds $ComDelayMs
                                }
                }
}

function Remove-FaisComObject {
                param ($Object)

                if ($null -eq $Object) { return }
                try { if ([System.Runtime.InteropServices.Marshal]::IsComObject($Object)) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object) } } catch { }
}

# The workbook Excel has open for this file, or $null. OneDrive / SharePoint workbooks report a
# web address as FullName, so those match on the file name. A file name match from another
# folder only counts when this file is locked (a mapped drive and its UNC path look different).
function Find-FaisOpenWorkbook {
                param ($ExcelApp, [string]$Path)

                if ($null -eq $ExcelApp -or [string]::IsNullOrWhiteSpace($Path)) { return $null }
                $full = [System.IO.Path]::GetFullPath($Path)
                $leaf = [System.IO.Path]::GetFileName($full)
                $byName = $null
                foreach ($wb in @(Invoke-FaisExcelCom { @($ExcelApp.Workbooks) })) {
                                $wbFull = [string]$wb.FullName
                                if ($wbFull -ieq $full) { return $wb }
                                if ([string]$wb.Name -ieq $leaf) {
                                                if ($wbFull -match '^(?i)(https?:)?//') { return $wb }
                                                if ($null -eq $byName) { $byName = $wb }
                                }
                }
                if ($null -ne $byName -and -not (Test-FileWritable $full)) { return $byName }
                return $null
}

# Saves one open workbook when it has changes, then closes it with -Close.
# A read-only workbook with changes, or one never saved anywhere, is left open (closing it would lose work).
# Returns {Path; Name; WasOpen; Saved; Closed; Message}. Never throws.
function Close-FaisWorkbookObject {
                param ($Workbook, [string]$Path = '', [switch]$Close)

                $r = [PSCustomObject]@{ Path = $Path; Name = ''; WasOpen = $true; Saved = $false; Closed = $false; Message = '' }
                try {
                                $r.Name = [string](Invoke-FaisExcelCom { $Workbook.Name })
                                if ($r.Path -eq '') { $r.Path = [string](Invoke-FaisExcelCom { $Workbook.FullName }) }
                                $dirty = -not [bool](Invoke-FaisExcelCom { $Workbook.Saved })
                                $readOnly = [bool](Invoke-FaisExcelCom { $Workbook.ReadOnly })
                                $neverSaved = [string]::IsNullOrWhiteSpace([string](Invoke-FaisExcelCom { $Workbook.Path }))
                                if ($dirty -and ($readOnly -or $neverSaved)) {
                                                $why = if ($readOnly) { 'is open read-only' } else { 'has never been saved' }
                                                $r.Message = "$($r.Name) $why in Excel and has unsaved changes, so FAIS left it open."
                                                return $r
                                }
                                if ($dirty) {
                                                [void](Invoke-FaisExcelCom { $Workbook.Save() })
                                                $r.Saved = $true
                                }
                                if ($Close) {
                                                [void](Invoke-FaisExcelCom { $Workbook.Close($false) })
                                                $r.Closed = $true
                                }
                                $did = @()
                                if ($r.Saved) { $did += 'saved' }
                                if ($r.Closed) { $did += 'closed' }
                                if ($did.Count -gt 0) { $r.Message = "$($r.Name) $($did -join ' and ') in Excel." }
                } catch {
                                $why = if (Test-FaisExcelBusyError $_.Exception) { 'Excel is busy (finish editing the cell or close the dialog in Excel)' } else { $_.Exception.Message }
                                $r.Message = "Could not save $($r.Name) in Excel: $why"
                }
                return $r
}

# Saves this file in Excel if Excel has it open (and closes it with -Close).
# Returns {Path; Name; WasOpen; Saved; Closed; Message}. Never throws.
function Save-FaisWorkbook {
                param ([string]$Path, [switch]$Close, $ExcelApp)

                $none = [PSCustomObject]@{ Path = $Path; Name = [System.IO.Path]::GetFileName($Path); WasOpen = $false; Saved = $false; Closed = $false; Message = '' }
                if ([string]::IsNullOrWhiteSpace($Path)) { return $none }
                $app = if ($null -ne $ExcelApp) { $ExcelApp } else { Get-FaisExcelApp }
                if ($null -eq $app) { return $none }
                $wb = $null
                try {
                                $wb = Find-FaisOpenWorkbook -ExcelApp $app -Path $Path
                                if ($null -eq $wb) { return $none }
                                $r = Close-FaisWorkbookObject -Workbook $wb -Path $Path -Close:$Close
                                if ($r.Closed) { [void](Wait-FaisFileReleased -Path $Path) }
                                return $r
                } catch {
                                $none.Message = "Could not check Excel for $($none.Name): $($_.Exception.Message)"
                                return $none
                } finally {
                                Remove-FaisComObject $wb
                                if ($null -eq $ExcelApp) { Remove-FaisComObject $app; [System.GC]::Collect() }
                }
}

# Saves the workbooks Excel has open for these files, and closes them unless -SaveOnly.
# -All also saves and closes every other workbook that has been saved somewhere before.
# Returns a result for each workbook that was open, and prints one line for each.
function Close-FaisExcelFiles {
                param (
                                [string[]]$Path = @(),
                                [switch]$SaveOnly,
                                [switch]$All,
                                $ExcelApp,
                                [switch]$Quiet
                )

                $results = New-Object System.Collections.Generic.List[object]
                $app = if ($null -ne $ExcelApp) { $ExcelApp } else { Get-FaisExcelApp }
                if ($null -eq $app) { return $results.ToArray() }
                try {
                                foreach ($p in $Path) {
                                                if ([string]::IsNullOrWhiteSpace($p)) { continue }
                                                $r = Save-FaisWorkbook -Path $p -Close:(-not $SaveOnly) -ExcelApp $app
                                                if ($r.WasOpen) { $results.Add($r) }
                                }
                                if ($All) {
                                                foreach ($wb in @(Invoke-FaisExcelCom { @($app.Workbooks) })) {
                                                                # Excel's hidden personal macro workbook stays open
                                                                if ([string]$wb.Name -match '^(?i)PERSONAL\.') { Remove-FaisComObject $wb; continue }
                                                                $r = Close-FaisWorkbookObject -Workbook $wb -Close:(-not $SaveOnly)
                                                                $results.Add($r)
                                                                Remove-FaisComObject $wb
                                                }
                                }
                } catch {
                                if (-not $Quiet) { Write-Host "Could not reach Excel to save the open workbooks: $($_.Exception.Message)" -ForegroundColor Yellow }
                } finally {
                                if ($null -eq $ExcelApp) { Remove-FaisComObject $app; [System.GC]::Collect() }
                }
                if (-not $Quiet) {
                                foreach ($r in $results) {
                                                if ($r.Message) {
                                                                $color = if ($r.Message -like 'Could not*' -or $r.Message -like '*left it open*') { 'Yellow' } else { 'DarkCyan' }
                                                                Write-Host "  $($r.Message)" -ForegroundColor $color
                                                }
                                }
                }
                return $results.ToArray()
}

# Called right before FAIS writes a file: saves and closes it in Excel, then makes sure the
# file is free. -Prompt (gets the path, returns $true to retry) is shown if it is still locked.
# Paths FAIS closed are added to -Closed (a List[string]) so they can be reopened afterwards.
# Returns $true when the file can be written.
function Request-FaisFileWritable {
                param (
                                [string]$Path,
                                [scriptblock]$Prompt,
                                $Closed,
                                $ExcelApp,
                                [switch]$NoExcel
                )

                if (Test-FileWritable $Path) { return $true }
                if (-not $NoExcel) {
                                $r = Save-FaisWorkbook -Path $Path -Close -ExcelApp $ExcelApp
                                if ($r.Message) {
                                                $color = if ($r.Closed) { 'DarkCyan' } else { 'Yellow' }
                                                Write-Host "  $($r.Message)" -ForegroundColor $color
                                }
                                if ($r.Closed -and $null -ne $Closed -and -not $Closed.Contains($Path)) { $Closed.Add($Path) }
                }
                return (Wait-FileWritable -Path $Path -Prompt $Prompt)
}

# Opens files again in Excel (the ones FAIS closed to write them)
function Open-FaisWorkbooks {
                param ([string[]]$Path = @())

                foreach ($p in $Path) {
                                if ([string]::IsNullOrWhiteSpace($p) -or -not (Test-Path -LiteralPath $p)) { continue }
                                try { Start-Process -FilePath $p -ErrorAction Stop }
                                catch { Write-Host "  Could not reopen $p : $($_.Exception.Message)" -ForegroundColor Yellow }
                }
}
