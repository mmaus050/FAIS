# PUCKBOARD.PSM1 DIRECTORY:
#             The WPF puck board window. Layout lives in PuckBoard.xaml; the data and
#             rules live in BoardModel / Assessment / SchedulingEngine.
# Show-PuckBoard            (opens the window, returns @{Action='Export'; Board; FillResult} or $null)
# Initialize-PuckBoardAssemblies
#
# How it hangs together:
#  - Every row, puck and slot is a DataTemplate bound to a plain PSCustomObject "view model".
#    A row is refreshed by replacing its view model in the day's ObservableCollection.
#  - Mouse / drag / key / focus events are handled once at window level and find the
#    element they belong to by its Tag (IPSlot, StudentSlot, EtdBox, ...).
#  - Anything that changes the board goes through Request-PBCommand, which runs it after
#    the current WPF event finishes, so rows are never rebuilt inside their own events.
#  - Pinned students are re-checked in the background after every change (DispatcherTimer),
#    which keeps the corner markers current without freezing the window.

$script:PB = $null
$script:PBKnownTags = @('IPPuck', 'StudentPuck', 'IPSlot', 'StudentSlot', 'IPPool', 'StudentPool', 'EtdBox', 'TypeBox',
                'MissionBox', 'CallsignBox', 'NumBox', 'NotesBox', 'DeleteLine', 'Row')
$script:PBCellTags = @('EtdBox', 'CallsignBox', 'NumBox', 'NotesBox')
$script:PBAutoText = '(auto)'
$script:PBBrushCache = @{}

$script:PBColors = @{
                StripeDayLocal = '#3B82F6'
                StripeOB       = '#8B5CF6'
                StripeReturn   = '#C4B5FD'
                StripeSim      = '#F59E0B'
                StripeOther    = '#94A3B8'
                MarkerNone     = 'Transparent'
                MarkerPending  = '#9CA3AF'
                MarkerClean    = '#16A34A'
                MarkerWarn     = '#F59E0B'
                MarkerHard     = '#DC2626'
                HeatClean      = '#BBF7D0'
                HeatWarn       = '#FDE68A'
                HeatHard       = '#FECACA'
}

# Runs queued board commands after the WPF event that asked for them has finished
$script:PBDrainAction = { Invoke-PBSafe { Invoke-PBDrain } }


# =================================================================================
#  ENTRY POINT
# =================================================================================

function Initialize-PuckBoardAssemblies {
                Add-Type -AssemblyName PresentationFramework
                Add-Type -AssemblyName PresentationCore
                Add-Type -AssemblyName WindowsBase
                Add-Type -AssemblyName System.Xaml
                Add-Type -AssemblyName Microsoft.VisualBasic
}

# -Data:     @{ Students; Instructors; SoftWeights; Formations; PrevSchedule; ImportSchedule }
# -Days:     'yyyy-MM-dd' days to open (saved future boards are added automatically)
# -Settings: BoardDir, TemplateDir, OBReturnOffsetMinutes, ReturnSortieType, SimInstructor, Trace
# Returns @{ Action = 'Export'; Board; FillResult } or $null when closed without exporting.
function Show-PuckBoard {
                param (
                                [Parameter(Mandatory)] $Data,
                                [string[]]$Days = @(),
                                [hashtable]$Settings = @{}
                )

                $window = Initialize-PuckBoardWindow -Data $Data -Days $Days -Settings $Settings
                [void]$window.ShowDialog()
                return $script:PB.Result
}

# Builds the window and all board state without showing it
function Initialize-PuckBoardWindow {
                param (
                                [Parameter(Mandatory)] $Data,
                                [string[]]$Days = @(),
                                [hashtable]$Settings = @{}
                )

                if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
                                throw "The puck board needs PowerShell running in STA mode. Start it with: powershell.exe -STA -File FAIS.ps1"
                }
                Initialize-PuckBoardAssemblies

                $set = Get-DefaultBoardSettings
                foreach ($k in $Settings.Keys) { $set[$k] = $Settings[$k] }
                $set.BoardDir = Resolve-PBWritableDir -Preferred $set.BoardDir -Leaf 'Boards'
                $set.TemplateDir = Resolve-PBWritableDir -Preferred $set.TemplateDir -Leaf 'Templates'

                if ($Days.Count -eq 0) { $Days = @((Get-NextFlyingDay).ToString('yyyy-MM-dd')) }
                $firstDay = @($Days | Sort-Object)[0]
                $firstDate = [datetime]::ParseExact($firstDay, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
                $savedDays = @(Get-SavedBoardDays -Dir $set.BoardDir -From $firstDate)
                $allDays = @(@($Days) + $savedDays | Sort-Object -Unique)
                $board = Open-BoardState -Dir $set.BoardDir -Days $allDays -Settings $set

                # Events a previous run planned on these days are plans, not flown events
                $reset = @(Reset-PlannedCompletions -Students $Data.Students -FromDate $firstDate -SoftWeights $Data.SoftWeights)
                Clear-AssessmentCache

                $xamlPath = Join-Path $PSScriptRoot 'PuckBoard.xaml'
                [xml]$xaml = Get-Content -LiteralPath $xamlPath -Raw -Encoding UTF8
                $reader = New-Object System.Xml.XmlNodeReader $xaml
                $window = [System.Windows.Markup.XamlReader]::Load($reader)

                $controls = @{}
                foreach ($name in @('UndoButton', 'HelpButton', 'QuickEntryBox', 'AddLineButton', 'QuickEntryFeedback', 'TemplateCombo',
                                'LoadTemplateButton', 'SaveTemplateButton', 'CopyPrevDayButton', 'ImportSheetButton', 'AddDayPicker', 'AddDayButton',
                                'RemoveDayButton', 'CancelButton', 'BackToEditButton', 'RefillButton', 'ExportButton', 'FillButton', 'StatusText',
                                'CountsText', 'IPSearchBox', 'IPPool', 'StudentSearchBox', 'StudentPool', 'ModeBanner', 'ModeBannerText', 'DayTabs',
                                'BusyOverlay', 'BusyText', 'BusyProgress', 'DropPopup', 'DropPopupPanel')) {
                                $c = $window.FindName($name)
                                if ($null -eq $c) { throw "PuckBoard.xaml is missing the element '$name'" }
                                $controls[$name] = $c
                }

                $script:PB = @{
                                Window      = $window
                                C           = $controls
                                Board       = $board
                                Data        = $Data
                                Settings    = $set
                                DayColls    = @{}
                                Assess      = @{}
                                RowAssess   = @{}
                                Queue       = (New-Object System.Collections.Generic.List[string])
                                Commands    = (New-Object System.Collections.Generic.List[object])
                                Drag        = $null
                                PendingDrop = $null
                                Heat        = $null
                                Mode        = 'Edit'
                                Overlay     = $null
                                FillResult  = $null
                                Result      = $null
                                Busy        = $false
                                Draining    = $false
                                TypeOptions = @('Day Local', 'O&B', [string]$set.ReturnSortieType, 'OFT', 'IFT')
                                LastFocusDay = $firstDay
                }

                Initialize-PBTimers
                Register-PBHandlers
                $controls.ImportSheetButton.Visibility = if (@($Data.ImportSchedule).Count -gt 0) { 'Visible' } else { 'Collapsed' }
                $controls.AddDayPicker.SelectedDate = (Get-NextFlyingDay -From ([datetime]::ParseExact(@($allDays)[-1], 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)))
                Update-PBTemplateList
                Set-PBMode 'Edit'
                Update-PBAll
                Start-PBAssessQueue

                $msg = "Board opened for $(@($board.Days) -join ', ')."
                if ($reset.Count -gt 0) { $msg += " Cleared planned events on/after $firstDay for $($reset.Count) student(s) so they can be rescheduled." }
                Set-PBStatus $msg
                return $window
}

# Boards / Templates folder next to FAIS.ps1, else %LOCALAPPDATA%\FAIS
function Resolve-PBWritableDir {
                param ([string]$Preferred, [string]$Leaf)

                $candidates = @()
                if ($Preferred) { $candidates += $Preferred }
                $candidates += (Join-Path (Split-Path $PSScriptRoot -Parent) $Leaf)
                if ($env:LOCALAPPDATA) { $candidates += (Join-Path (Join-Path $env:LOCALAPPDATA 'FAIS') $Leaf) }
                foreach ($dir in $candidates) {
                                try {
                                                if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null }
                                                $probe = Join-Path $dir '.write-test'
                                                [System.IO.File]::WriteAllText($probe, 'ok')
                                                Remove-Item -LiteralPath $probe -Force
                                                return $dir
                                } catch { }
                }
                throw "No writable folder for $Leaf"
}


# =================================================================================
#  SMALL HELPERS
# =================================================================================

function Write-PBLog {
                param ($Message)

                try {
                                $dir = if ($env:TEMP) { $env:TEMP } else { [System.IO.Path]::GetTempPath() }
                                Add-Content -LiteralPath (Join-Path $dir 'FAIS_board.log') -Value ("{0} {1}" -f (Get-Date -Format 's'), $Message)
                } catch { }
}

# Every event handler runs through here so one bad click can't close the window
function Invoke-PBSafe {
                param ([scriptblock]$Action)

                try { & $Action }
                catch {
                                $msg = $_.Exception.Message
                                Write-PBLog ("{0}`n{1}" -f $msg, $_.ScriptStackTrace)
                                try { Set-PBStatus "Something went wrong: $msg (details in %TEMP%\FAIS_board.log)" 'Hard' } catch { }
                }
}

function Get-PBBrush {
                param ([string]$Color)

                if (-not $script:PBBrushCache.ContainsKey($Color)) {
                                $b = (New-Object System.Windows.Media.BrushConverter).ConvertFromString($Color)
                                $b.Freeze()
                                $script:PBBrushCache[$Color] = $b
                }
                return $script:PBBrushCache[$Color]
}

function Set-PBStatus {
                param ([string]$Text, [string]$Kind = 'Info')

                $tb = $script:PB.C.StatusText
                $tb.Text = $Text
                $color = switch ($Kind) { 'Hard' { '#B91C1C' } 'Warn' { '#B45309' } 'Good' { '#15803D' } default { '#0F172A' } }
                $tb.Foreground = Get-PBBrush $color
}

# Lets the window repaint during long work
function Invoke-PBDoEvents {
                $script:PB.Window.Dispatcher.Invoke([System.Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
}

function Get-PBActiveDay {
                $tab = $script:PB.C.DayTabs.SelectedItem
                if ($null -ne $tab -and $tab.Tag) { return [string]$tab.Tag }
                if ($script:PB.Board.Days.Count -gt 0) { return $script:PB.Board.Days[0] }
                return $null
}

function Get-PBDayLabel {
                param ([string]$Day)

                $d = [datetime]::ParseExact($Day, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
                $n = @($script:PB.Board.Lines | Where-Object { $_.Day -eq $Day }).Count
                return ("{0}  ({1})" -f $d.ToString('ddd dd MMM'), $n)
}

# Walks up from a clicked element to the nearest element with one of our Tags
function Find-PBTagged {
                param ($Element, [string[]]$Tags = $script:PBKnownTags)

                $el = $Element
                while ($null -ne $el) {
                                if ($el -is [System.Windows.FrameworkElement] -and $el.Tag -is [string] -and $Tags -contains $el.Tag) { return $el }
                                $el = Get-PBParent $el
                }
                return $null
}

function Get-PBParent {
                param ($Element)

                if ($Element -is [System.Windows.Media.Visual] -or $Element -is [System.Windows.Media.Media3D.Visual3D]) {
                                $p = [System.Windows.Media.VisualTreeHelper]::GetParent($Element)
                                if ($null -ne $p) { return $p }
                }
                if ($Element -is [System.Windows.DependencyObject]) {
                                return [System.Windows.LogicalTreeHelper]::GetParent($Element)
                }
                return $null
}

# True when the click landed in something that needs the mouse itself (text box, combo, button, scroll bar)
function Test-PBInsideInput {
                param ($Element)

                $el = $Element
                while ($null -ne $el) {
                                if ($el -is [System.Windows.Controls.Primitives.TextBoxBase] -or $el -is [System.Windows.Controls.ComboBox] -or
                                                $el -is [System.Windows.Controls.Primitives.ButtonBase] -or $el -is [System.Windows.Controls.Primitives.ScrollBar] -or
                                                $el -is [System.Windows.Controls.DatePicker]) { return $true }
                                if ($el -is [System.Windows.FrameworkElement] -and $el.Tag -is [string] -and @('IPPuck', 'StudentPuck', 'IPSlot', 'StudentSlot') -contains $el.Tag) { return $false }
                                $el = Get-PBParent $el
                }
                return $false
}

# All rendered elements with a Tag inside an element (used to find a line's slot)
function Get-PBTaggedDescendants {
                param ($Root, [string]$Tag)

                $found = New-Object System.Collections.Generic.List[object]
                if ($null -eq $Root) { return }
                $stack = New-Object System.Collections.Generic.Stack[object]
                $stack.Push($Root)
                while ($stack.Count -gt 0) {
                                $node = $stack.Pop()
                                if ($node -is [System.Windows.FrameworkElement] -and [string]$node.Tag -eq $Tag) { $found.Add($node) }
                                if ($node -is [System.Windows.Media.Visual]) {
                                                $n = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($node)
                                                for ($i = 0; $i -lt $n; $i++) { $stack.Push([System.Windows.Media.VisualTreeHelper]::GetChild($node, $i)) }
                                }
                }
                return $found.ToArray()
}

# The rendered slot (StudentSlot / IPSlot) for a line on the active tab
function Get-PBSlotElement {
                param ([string]$LineId, [string]$Tag = 'StudentSlot')

                $tab = $script:PB.C.DayTabs.SelectedItem
                if ($null -eq $tab) { return $null }
                foreach ($el in @(Get-PBTaggedDescendants -Root $tab.Content -Tag $Tag)) {
                                if ($null -ne $el.DataContext -and [string]$el.DataContext.LineId -eq $LineId) { return $el }
                }
                return $null
}


# =================================================================================
#  TIMERS AND EVENTS
# =================================================================================

function Initialize-PBTimers {
                $w = $script:PB.Window

                $assess = New-Object System.Windows.Threading.DispatcherTimer
                $assess.Interval = [TimeSpan]::FromMilliseconds(40)
                $assess.Add_Tick({ Invoke-PBSafe { Invoke-PBAssessTick } })
                $script:PB.AssessTimer = $assess

                $save = New-Object System.Windows.Threading.DispatcherTimer
                $save.Interval = [TimeSpan]::FromSeconds(2)
                $save.Add_Tick({ Invoke-PBSafe { $script:PB.SaveTimer.Stop(); Save-PBBoard } })
                $script:PB.SaveTimer = $save

                $popup = New-Object System.Windows.Threading.DispatcherTimer
                $popup.Interval = [TimeSpan]::FromSeconds(7)
                $popup.Add_Tick({ Invoke-PBSafe { $script:PB.PopupTimer.Stop(); $script:PB.C.DropPopup.IsOpen = $false } })
                $script:PB.PopupTimer = $popup

                $heat = New-Object System.Windows.Threading.DispatcherTimer
                $heat.Interval = [TimeSpan]::FromMilliseconds(30)
                $heat.Add_Tick({ Invoke-PBSafe { Invoke-PBHeatTick } })
                $script:PB.HeatTimer = $heat

                $w.Add_Closing({
                                param($s, $e)
                                Invoke-PBSafe {
                                                foreach ($t in @($script:PB.AssessTimer, $script:PB.SaveTimer, $script:PB.PopupTimer, $script:PB.HeatTimer)) { $t.Stop() }
                                                Save-PBBoard
                                }
                })
}

function Register-PBHandlers {
                $w = $script:PB.Window
                $C = $script:PB.C

                $w.AddHandler([System.Windows.UIElement]::PreviewMouseLeftButtonDownEvent,
                                [System.Windows.Input.MouseButtonEventHandler]{ param($s, $e) Invoke-PBSafe { Invoke-PBMouseDown $e } }, $true)
                $w.AddHandler([System.Windows.UIElement]::PreviewMouseMoveEvent,
                                [System.Windows.Input.MouseEventHandler]{ param($s, $e) Invoke-PBSafe { Invoke-PBMouseMove $e } }, $true)
                $w.AddHandler([System.Windows.UIElement]::DragEnterEvent,
                                [System.Windows.DragEventHandler]{ param($s, $e) Invoke-PBSafe { Invoke-PBDragOver $e -Enter } }, $true)
                $w.AddHandler([System.Windows.UIElement]::DragOverEvent,
                                [System.Windows.DragEventHandler]{ param($s, $e) Invoke-PBSafe { Invoke-PBDragOver $e } }, $true)
                $w.AddHandler([System.Windows.UIElement]::DropEvent,
                                [System.Windows.DragEventHandler]{ param($s, $e) Invoke-PBSafe { Invoke-PBDrop $e } }, $true)
                $w.AddHandler([System.Windows.UIElement]::PreviewKeyDownEvent,
                                [System.Windows.Input.KeyEventHandler]{ param($s, $e) Invoke-PBSafe { Invoke-PBKeyDown $e } }, $true)
                $w.AddHandler([System.Windows.UIElement]::LostKeyboardFocusEvent,
                                [System.Windows.Input.KeyboardFocusChangedEventHandler]{ param($s, $e) Invoke-PBSafe { Invoke-PBLostFocus $e } }, $true)
                $w.AddHandler([System.Windows.Controls.Primitives.Selector]::SelectionChangedEvent,
                                [System.Windows.Controls.SelectionChangedEventHandler]{ param($s, $e) Invoke-PBSafe { Invoke-PBSelectionChanged $e } }, $true)
                $w.AddHandler([System.Windows.FrameworkElement]::ToolTipOpeningEvent,
                                [System.Windows.Controls.ToolTipEventHandler]{ param($s, $e) Invoke-PBSafe { Invoke-PBToolTipOpening $e } }, $true)
                $w.AddHandler([System.Windows.FrameworkElement]::ContextMenuOpeningEvent,
                                [System.Windows.Controls.ContextMenuEventHandler]{ param($s, $e) Invoke-PBSafe { Invoke-PBContextMenuOpening $e } }, $true)
                $w.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent,
                                [System.Windows.RoutedEventHandler]{ param($s, $e) Invoke-PBSafe { Invoke-PBRowClick $e } }, $true)

                $C.AddLineButton.Add_Click({ Invoke-PBSafe { Invoke-PBQuickEntry } })
                $C.UndoButton.Add_Click({ Invoke-PBSafe { Request-PBCommand @{ Cmd = 'Undo' } } })
                $C.HelpButton.Add_Click({ Invoke-PBSafe { Show-PBHelp } })
                $C.LoadTemplateButton.Add_Click({ Invoke-PBSafe { Invoke-PBLoadTemplate } })
                $C.SaveTemplateButton.Add_Click({ Invoke-PBSafe { Invoke-PBSaveTemplate } })
                $C.CopyPrevDayButton.Add_Click({ Invoke-PBSafe { Request-PBCommand @{ Cmd = 'CopyPrevDay' } } })
                $C.ImportSheetButton.Add_Click({ Invoke-PBSafe { Invoke-PBImportSheet } })
                $C.AddDayButton.Add_Click({ Invoke-PBSafe { Invoke-PBAddDay } })
                $C.RemoveDayButton.Add_Click({ Invoke-PBSafe { Invoke-PBRemoveDay } })
                $C.FillButton.Add_Click({ Invoke-PBSafe { Request-PBCommand @{ Cmd = 'Fill' } } })
                $C.RefillButton.Add_Click({ Invoke-PBSafe { Request-PBCommand @{ Cmd = 'Fill' } } })
                $C.BackToEditButton.Add_Click({ Invoke-PBSafe { Exit-PBReview; Update-PBAll; Set-PBStatus 'Back to editing. The fill was discarded.' } })
                $C.ExportButton.Add_Click({ Invoke-PBSafe { Invoke-PBExport } })
                $C.CancelButton.Add_Click({ Invoke-PBSafe { Invoke-PBCancel } })
                $C.IPSearchBox.Add_TextChanged({ Invoke-PBSafe { Update-PBPools } })
                $C.StudentSearchBox.Add_TextChanged({ Invoke-PBSafe { Update-PBPools } })
}

# Board changes asked for from inside an element's event run right after that event
function Request-PBCommand {
                param ([hashtable]$Command)

                $script:PB.Commands.Add($Command)
                [void]$script:PB.Window.Dispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Background, [System.Action]$script:PBDrainAction)
}

function Invoke-PBDrain {
                if ($script:PB.Draining) { return }
                $script:PB.Draining = $true
                try {
                                while ($script:PB.Commands.Count -gt 0) {
                                                $cmd = $script:PB.Commands[0]
                                                $script:PB.Commands.RemoveAt(0)
                                                Invoke-PBSafe { Invoke-PBCommand $cmd }
                                }
                } finally { $script:PB.Draining = $false }
}


# =================================================================================
#  REFRESHING THE VIEW
# =================================================================================

# Runs a board edit then refreshes everything that depends on it
function Invoke-PBChange {
                param ([scriptblock]$Action, [string]$Status, [switch]$KeepReview)

                if ($script:PB.Mode -eq 'Review' -and -not $KeepReview) { Exit-PBReview }
                $script:PB.Busy = $true
                try { & $Action } finally { $script:PB.Busy = $false }
                Update-PBAll
                Request-PBAutosave
                Start-PBAssessQueue
                if ($Status) { Set-PBStatus $Status }
}

function Update-PBAll {
                $script:PB.Busy = $true
                try {
                                Update-PBTabs
                                foreach ($day in $script:PB.Board.Days) { Update-PBDay $day }
                                Update-PBPools
                                Update-PBCounts
                } finally { $script:PB.Busy = $false }
}

# One tab per board day; rebuilt only when the days change
function Update-PBTabs {
                $tabs = $script:PB.C.DayTabs
                $want = @($script:PB.Board.Days)
                $have = @(foreach ($t in $tabs.Items) { [string]$t.Tag })
                if (($want -join '|') -ne ($have -join '|')) {
                                $selected = Get-PBActiveDay
                                if ($null -eq $selected -or $want -notcontains $selected) { $selected = $script:PB.LastFocusDay }
                                $tabs.Items.Clear()
                                $script:PB.DayColls = @{}
                                foreach ($day in $want) {
                                                $coll = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
                                                $script:PB.DayColls[$day] = $coll

                                                $ic = New-Object System.Windows.Controls.ItemsControl
                                                $ic.ItemTemplate = $script:PB.Window.FindResource('LineRowTemplate')
                                                $ic.ItemsSource = $coll
                                                $sv = New-Object System.Windows.Controls.ScrollViewer
                                                $sv.VerticalScrollBarVisibility = [System.Windows.Controls.ScrollBarVisibility]::Auto
                                                $sv.Content = $ic
                                                $hdr = New-Object System.Windows.Controls.ContentControl
                                                $hdr.ContentTemplate = $script:PB.Window.FindResource('LineHeaderTemplate')
                                                $hdr.Content = ' '
                                                $dp = New-Object System.Windows.Controls.DockPanel
                                                [System.Windows.Controls.DockPanel]::SetDock($hdr, [System.Windows.Controls.Dock]::Top)
                                                [void]$dp.Children.Add($hdr)
                                                [void]$dp.Children.Add($sv)

                                                $ti = New-Object System.Windows.Controls.TabItem
                                                $ti.Tag = $day
                                                $ti.Content = $dp
                                                [void]$tabs.Items.Add($ti)
                                }
                                foreach ($t in $tabs.Items) { if ([string]$t.Tag -eq $selected) { $tabs.SelectedItem = $t } }
                                if ($null -eq $tabs.SelectedItem -and $tabs.Items.Count -gt 0) { $tabs.SelectedIndex = 0 }
                }
                foreach ($t in $tabs.Items) { $t.Header = Get-PBDayLabel ([string]$t.Tag) }
}

function Update-PBDay {
                param ([string]$Day)

                $coll = $script:PB.DayColls[$Day]
                if ($null -eq $coll) { return }
                $coll.Clear()
                foreach ($line in (Get-BoardDayLines -Board $script:PB.Board -Day $Day)) { $coll.Add((New-PBLineVM $line)) }
}

# Replaces one row's view model (skipped while the user is typing in that row)
function Update-PBLineRow {
                param ([string]$LineId)

                $line = Get-BoardLine $script:PB.Board $LineId
                if ($null -eq $line) { return $true }
                $coll = $script:PB.DayColls[$line.Day]
                if ($null -eq $coll) { return $true }
                $focused = [System.Windows.Input.Keyboard]::FocusedElement
                if ($focused -is [System.Windows.Controls.Primitives.TextBoxBase] -or $focused -is [System.Windows.Controls.ComboBox]) {
                                if ($null -ne $focused.DataContext -and $null -ne $focused.DataContext.PSObject.Properties['LineId'] -and $focused.DataContext.LineId -eq $LineId) { return $false }
                }
                for ($i = 0; $i -lt $coll.Count; $i++) {
                                if ($coll[$i].LineId -eq $LineId) { $coll[$i] = New-PBLineVM $line; break }
                }
                return $true
}

function Get-PBCurrentAssessment {
                param ($Line)

                $a = $script:PB.Assess[$Line.Id]
                if ($null -eq $a) { return $null }
                if ($a.StudentName -ne $Line.Student) { return $null }
                return $a
}

function Test-PBAssessmentFresh {
                param ($Line)

                $a = Get-PBCurrentAssessment $Line
                return ($null -ne $a -and $a.BoardVersion -eq $script:PB.Board.Version -and [string]$a.RequestedMission -ieq [string]$Line.Mission)
}

# Missions offered in a pinned line's dropdown
function Get-PBMissionChoices {
                param ($Line)

                $choices = New-Object System.Collections.Generic.List[string]
                $choices.Add($script:PBAutoText)
                $a = Get-PBCurrentAssessment $Line
                $names = @()
                if ($null -ne $a -and @($a.EligibleEvents).Count -gt 0) { $names = @($a.EligibleEvents) }
                else {
                                $student = Resolve-StudentByName -Students $script:PB.Data.Students -Name $Line.Student
                                $types = Get-AllowedEventTypes ([string]$Line.SortieType)
                                if ($null -ne $student -and $null -ne $types) {
                                                $names = @(Get-OptedEvents -SyllabusData $student.Syllabus | Where-Object { ([string]$_.Type).Trim() -in $types } |
                                                                Sort-Object -Property { Get-SyllabusSortKey $_ } | ForEach-Object { [string]$_.Mission })
                                }
                }
                foreach ($n in $names) { if (-not $choices.Contains($n)) { $choices.Add($n) } }
                if ($Line.Mission -and -not $choices.Contains([string]$Line.Mission)) { $choices.Add([string]$Line.Mission) }
                return $choices.ToArray()
}

# The view model a row's DataTemplate binds to
function New-PBLineVM {
                param ($Line)

                $K = $script:PBColors
                $review = ($script:PB.Mode -eq 'Review')
                $ov = $null
                if ($review -and $null -ne $script:PB.Overlay) { $ov = $script:PB.Overlay[$Line.Id] }
                $isReturn = Test-IsReturnSortieType ([string]$Line.SortieType)
                $isSim = Test-IsSimSortieType ([string]$Line.SortieType)
                $isOB = Test-IsOBSortieType ([string]$Line.SortieType)
                $pinned = -not [string]::IsNullOrWhiteSpace($Line.Student)

                $stripe = if ($isReturn) { $K.StripeReturn } elseif ($isOB) { $K.StripeOB } elseif ($isSim) { $K.StripeSim } elseif ($Line.SortieType -like '*Day Local*') { $K.StripeDayLocal } else { $K.StripeOther }

                # IP slot
                $ipText = [string]$Line.Instructor
                $ipFg = '#1E3A8A'; $ipBg = '#DBEAFE'; $ipBorder = '#93C5FD'; $ipStyle = 'Normal'
                $ipTip = 'Drag an IP puck here. Drag it to another line to move it (Ctrl = copy).'
                if ([string]::IsNullOrWhiteSpace($ipText)) {
                                $ipFg = '#94A3B8'; $ipBg = '#F8FAFC'; $ipBorder = '#CBD5E1'; $ipStyle = 'Italic'
                                $ipText = if ($isSim) { 'SIM IP' } else { 'drop IP' }
                }
                if ($isReturn -and $Line.LinkedLineId) { $ipTip = 'Follows its O&B (change the IP on the O&B line).' }

                # Student slot
                $studentText = ''; $sub = ''; $fg = '#14532D'; $weight = 'SemiBold'; $style = 'Normal'
                $bg = '#F8FAFC'; $border = '#CBD5E1'; $thick = '1'; $marker = $K.MarkerNone; $allowDrop = $true
                if ($pinned) {
                                $studentText = [string]$Line.Student
                                $bg = '#DCFCE7'; $border = '#22C55E'; $thick = '2'
                                $a = Get-PBCurrentAssessment $Line
                                $marker = $K.MarkerPending
                                if ($null -ne $a) {
                                                $marker = switch ($a.Status) { 'Clean' { $K.MarkerClean } 'Warn' { $K.MarkerWarn } 'Hard' { $K.MarkerHard } default { $K.MarkerPending } }
                                                if (-not $Line.Mission -and $a.AutoMission) { $sub = "-> $($a.AutoMission)" }
                                }
                                if ($Line.WingOfLineId) {
                                                $lead = Get-BoardLine $script:PB.Board $Line.WingOfLineId
                                                if ($null -ne $lead) { $sub = ("wing to {0} {1}" -f $lead.ETD.ToString('HHmm'), $lead.Student).Trim() }
                                }
                                if ($review -and $null -ne $ov -and $ov.Override) { $marker = $K.MarkerHard }
                }
                elseif ($isReturn) {
                                $allowDrop = $false
                                $style = 'Italic'; $weight = 'Normal'; $fg = '#7C3AED'; $bg = '#F5F3FF'; $border = '#DDD6FE'
                                $ob = if ($Line.LinkedLineId) { Get-BoardLine $script:PB.Board $Line.LinkedLineId } else { $null }
                                if ($review -and $null -ne $ov -and $ov.Student) {
                                                $studentText = $ov.Student
                                                $sub = $ov.Mission
                                                if ($ov.FilledBy -eq 'ForSale') { $fg = '#B91C1C'; $bg = '#FEE2E2'; $border = '#FCA5A5' }
                                }
                                elseif ($null -ne $ob -and $ob.Student) { $studentText = "<- $($ob.Student)" ; $sub = 'with O&B' }
                                elseif ($null -ne $ob) { $studentText = '<- follows O&B' }
                                else { $studentText = 'not linked to an O&B'; $fg = '#B45309' }
                }
                elseif ($review -and $null -ne $ov -and $ov.Student) {
                                $studentText = $ov.Student
                                $sub = $ov.Mission
                                $style = 'Italic'
                                switch ($ov.FilledBy) {
                                                'Engine' { $fg = '#1E3A8A'; $bg = '#EFF6FF'; $border = '#60A5FA'; $sub = "$($ov.Mission) AUTO".Trim() }
                                                'Wing' { $fg = '#1E3A8A'; $bg = '#EFF6FF'; $border = '#60A5FA'; $sub = "$($ov.Mission) WING".Trim() }
                                                'DS' { $fg = '#92400E'; $bg = '#FEF3C7'; $border = '#F59E0B' }
                                                'ForSale' { $fg = '#B91C1C'; $bg = '#FEE2E2'; $border = '#FCA5A5'; $sub = '' }
                                                default { $fg = '#1E3A8A' }
                                }
                }
                else {
                                $studentText = 'drop student'
                                $fg = '#94A3B8'; $weight = 'Normal'; $style = 'Italic'
                }

                # Mission dropdown
                $missionOptions = @('')
                $missionSelected = ''
                $missionEnabled = $false
                if ($pinned) {
                                $missionOptions = Get-PBMissionChoices $Line
                                $missionSelected = if ($Line.Mission) { [string]$Line.Mission } else { $script:PBAutoText }
                                $missionEnabled = $true
                }
                elseif ($review -and $null -ne $ov -and $ov.Mission) {
                                $missionOptions = @([string]$ov.Mission)
                                $missionSelected = [string]$ov.Mission
                }

                $typeOptions = $script:PB.TypeOptions
                if ($typeOptions -notcontains [string]$Line.SortieType) { $typeOptions = @($typeOptions) + [string]$Line.SortieType }
                $script:PB.RowAssess[$Line.Id] = $script:PB.Assess[$Line.Id]

                # Callsign / notes (review shows what the fill wrote)
                $callsign = [string]$Line.Callsign
                $num = [string]$Line.Num
                $notes = [string]$Line.Notes
                $readOnly = $false
                $notesFg = '#0F172A'
                $rowBg = '#FFFFFF'
                $csTip = 'Blank = the IP''s callsign (formations get FormCallsigns)'
                if ($review -and $null -ne $ov) {
                                $callsign = $ov.Callsign; $num = $ov.Num; $notes = $ov.Notes; $readOnly = $true
                                if ($ov.Override) { $notesFg = '#B91C1C'; $rowBg = '#FEF2F2' }
                }

                return [PSCustomObject]@{
                                LineId                 = $Line.Id
                                EtdText                = $Line.ETD.ToString('HHmm')
                                SortieType             = [string]$Line.SortieType
                                TypeOptions            = $typeOptions
                                StripeBrush            = $stripe
                                InstructorRaw          = [string]$Line.Instructor
                                InstructorText         = $ipText
                                IPForeground           = $ipFg
                                IPBackground           = $ipBg
                                IPBorder               = $ipBorder
                                IPFontStyle            = $ipStyle
                                IPToolTip              = $ipTip
                                StudentRaw             = [string]$Line.Student
                                StudentText            = $studentText
                                StudentSubText         = $sub
                                StudentForeground      = $fg
                                StudentWeight          = $weight
                                StudentFontStyle       = $style
                                StudentBackground      = $bg
                                StudentBorder          = $border
                                StudentBorderThickness = $thick
                                AllowStudentDrop       = $allowDrop
                                MarkerBrush            = $marker
                                MissionRaw             = [string]$Line.Mission
                                MissionOptions         = $missionOptions
                                MissionSelected        = $missionSelected
                                MissionEnabled         = $missionEnabled
                                CallsignText           = $callsign
                                NumText                = $num
                                NotesText              = $notes
                                NotesForeground        = $notesFg
                                NotesToolTip           = $(if ([string]::IsNullOrWhiteSpace($notes)) { $null } else { $notes })
                                ReadOnlyFields         = $readOnly
                                CallsignToolTip        = $csTip
                                RowBackground          = $rowBg
                                IsReturn               = $isReturn
                }
}

function Update-PBPools {
                $day = Get-PBActiveDay
                $board = $script:PB.Board
                $C = $script:PB.C

                $ipFilter = ([string]$C.IPSearchBox.Text).Trim()
                $ipVMs = foreach ($ip in @($script:PB.Data.Instructors | Sort-Object -Property Name)) {
                                if ($ipFilter -and ([string]$ip.Name -notlike "*$ipFilter*") -and ([string]$ip.Callsign -notlike "*$ipFilter*")) { continue }
                                $info = Get-InstructorPuckInfo -Instructor $ip -Board $board -Day $day
                                [PSCustomObject]@{
                                                Name            = $info.Name
                                                SubText         = ("{0}  {1}" -f $info.Callsign, ($info.Badges -join ' ')).Trim()
                                                CountText       = "x$($info.Lines)"
                                                CountVisibility = $(if ($info.Lines -gt 0) { 'Visible' } else { 'Collapsed' })
                                                ToolTipText     = $(if ($info.Lines -gt 0) { "On $($info.Lines) line(s) this day: $($info.Times -join ', ')" } else { 'Not on any line this day' })
                                }
                }
                $C.IPPool.ItemsSource = @($ipVMs)

                $spFilter = ([string]$C.StudentSearchBox.Text).Trim()
                $spVMs = foreach ($s in @($script:PB.Data.Students | Sort-Object -Property @{ Expression = { [double]$_.Priority }; Descending = $true }, Name)) {
                                if ([string]::IsNullOrWhiteSpace([string]$s.Name)) { continue }
                                if ($spFilter -and ([string]$s.Name -notlike "*$spFilter*")) { continue }
                                $info = Get-StudentPuckInfo -Student $s -Board $board -Day $day
                                $bg = '#DCFCE7'; $bd = '#86EFAC'
                                if ($info.Blocked) { $bg = '#E5E7EB'; $bd = '#9CA3AF' }
                                elseif ($info.Badges | Where-Object { $_ -like 'U-*' -or $_ -like 'ABS*' }) { $bg = '#FEF3C7'; $bd = '#FCD34D' }
                                $turnsBrush = if ($info.Turns -ge 3) { '#DC2626' } elseif ($info.Turns -eq 2) { '#D97706' } else { '#16A34A' }
                                $tip = @("Next: $($info.Next)", ("Sim debt: {0:0.##}" -f $info.SimDebt), "Priority: $($s.Priority)")
                                if ($info.Badges.Count -gt 0) { $tip += "Flags: $($info.Badges -join ', ')" }
                                foreach ($a in $info.Absences) { $tip += "Absent $a" }
                                [PSCustomObject]@{
                                                Name            = $info.Name
                                                SubText         = ("next {0}   debt {1:0.#}" -f $info.Next, $info.SimDebt)
                                                BadgeText       = ($info.Badges -join '  ')
                                                BadgeVisibility = $(if ($info.Badges.Count -gt 0) { 'Visible' } else { 'Collapsed' })
                                                TurnsText       = $(if ($info.Turns -eq 1) { '1 turn' } else { "$($info.Turns) turns" })
                                                TurnsVisibility = $(if ($info.Turns -gt 0) { 'Visible' } else { 'Collapsed' })
                                                TurnsBrush      = $turnsBrush
                                                Background      = $bg
                                                BorderBrush     = $bd
                                                ToolTipText     = ($tip -join "`n")
                                }
                }
                $C.StudentPool.ItemsSource = @($spVMs)
}

function Update-PBCounts {
                $board = $script:PB.Board
                $day = Get-PBActiveDay
                $lines = @($board.Lines | Where-Object { $_.Day -eq $day })
                $pins = @($lines | Where-Object { $_.Student })
                $flags = 0
                foreach ($p in $pins) { $a = Get-PBCurrentAssessment $p; if ($null -ne $a -and $a.Status -eq 'Hard') { $flags++ } }
                $text = "{0} lines | {1} hardcoded | {2} flagged" -f $lines.Count, $pins.Count, $flags
                if ($script:PB.Queue.Count -gt 0) { $text += " | checking $($script:PB.Queue.Count)..." }
                if ($script:PB.Mode -eq 'Review' -and $null -ne $script:PB.FillResult) {
                                $st = $script:PB.FillResult.Stats
                                $text += "   ||   Fill (all days): {0} auto, {1} wing, {2} return, {3} DS, {4} for sale, {5} override(s)" -f $st.Filled, $st.Wing, $st.Return, $st.DS, $st.ForSale, $st.Overrides
                }
                $script:PB.C.CountsText.Text = $text
}

function Set-PBMode {
                param ([ValidateSet('Edit', 'Review')] [string]$Mode)

                $C = $script:PB.C
                $script:PB.Mode = $Mode
                $review = ($Mode -eq 'Review')
                $C.FillButton.Visibility = if ($review) { 'Collapsed' } else { 'Visible' }
                foreach ($b in @($C.BackToEditButton, $C.RefillButton, $C.ExportButton)) { $b.Visibility = if ($review) { 'Visible' } else { 'Collapsed' } }
                $C.ModeBanner.Visibility = if ($review) { 'Visible' } else { 'Collapsed' }
                if ($review) {
                                $C.ModeBannerText.Text = 'REVIEW: italic blue = filled by the scheduler, green = hardcoded, red notes = OVERRIDE (hardcoded against a rule), red = For Sale. Export writes OutputSchedule.xlsx. Any change goes back to editing.'
                }
}

function Exit-PBReview {
                if ($script:PB.Mode -ne 'Review') { return }
                $script:PB.Overlay = $null
                $script:PB.FillResult = $null
                Set-PBMode 'Edit'
}

function Request-PBAutosave {
                $script:PB.SaveTimer.Stop()
                $script:PB.SaveTimer.Start()
}

function Save-PBBoard {
                try {
                                Save-BoardState -Board $script:PB.Board -Dir $script:PB.Settings.BoardDir
                } catch {
                                Write-PBLog "Autosave failed: $($_.Exception.Message)"
                }
}


# =================================================================================
#  BACKGROUND CHECKS (corner markers)
# =================================================================================

# Queues every hardcoded line, active day first
function Start-PBAssessQueue {
                $q = $script:PB.Queue
                $q.Clear()
                $active = Get-PBActiveDay
                $pins = @($script:PB.Board.Lines | Where-Object { $_.Student -and -not (Test-IsReturnSortieType ([string]$_.SortieType)) })
                foreach ($l in @($pins | Where-Object { $_.Day -eq $active } | Sort-Object -Property ETD)) { $q.Add($l.Id) }
                foreach ($l in @($pins | Where-Object { $_.Day -ne $active } | Sort-Object -Property ETD)) { $q.Add($l.Id) }
                if ($q.Count -gt 0) { $script:PB.AssessTimer.Start() }
                Update-PBCounts
}

function Invoke-PBAssessTick {
                if ($script:PB.Busy -or $script:PB.Draining) { return }
                $q = $script:PB.Queue
                $sw = [System.Diagnostics.Stopwatch]::StartNew()
                $deferred = New-Object System.Collections.Generic.List[string]
                while ($q.Count -gt 0 -and $sw.ElapsedMilliseconds -lt 35) {
                                $id = $q[0]
                                $q.RemoveAt(0)
                                $line = Get-BoardLine $script:PB.Board $id
                                if ($null -eq $line -or -not $line.Student) { continue }
                                if (-not (Test-PBAssessmentFresh $line)) {
                                                $script:PB.Assess[$id] = Get-PlacementAssessment -Board $script:PB.Board -LineId $id -StudentName $line.Student `
                                                                -MissionName ([string]$line.Mission) -Data $script:PB.Data
                                }
                                elseif ([object]::ReferenceEquals($script:PB.RowAssess[$id], $script:PB.Assess[$id])) { continue }
                                if (-not (Update-PBLineRow $id)) { $deferred.Add($id) }
                }
                foreach ($d in $deferred) { $q.Add($d) }
                if ($q.Count -eq 0 -or $q.Count -eq $deferred.Count) {
                                if ($q.Count -eq 0) { $script:PB.AssessTimer.Stop() }
                }
                Update-PBCounts
}


# =================================================================================
#  DRAG AND DROP
#  Payload is the text "Kind|Name|FromLineId" (Kind = Student or IP; FromLineId blank
#  when the puck came from a pool).
# =================================================================================

function Invoke-PBMouseDown {
                param ($e)

                $script:PB.Drag = $null
                $src = $e.OriginalSource
                if (Test-PBInsideInput $src) { return }
                $el = Find-PBTagged $src
                if ($null -eq $el) { return }
                $vm = $el.DataContext
                $kind = $null; $name = $null; $from = ''
                switch ([string]$el.Tag) {
                                'IPPuck' { $kind = 'IP'; $name = $vm.Name }
                                'StudentPuck' { $kind = 'Student'; $name = $vm.Name }
                                'IPSlot' {
                                                [void]$el.Focus()
                                                if ($vm.InstructorRaw -and -not ($vm.IsReturn)) { $kind = 'IP'; $name = $vm.InstructorRaw; $from = $vm.LineId }
                                }
                                'StudentSlot' {
                                                [void]$el.Focus()
                                                if ($vm.StudentRaw) { $kind = 'Student'; $name = $vm.StudentRaw; $from = $vm.LineId }
                                }
                }
                if ($null -eq $kind) { return }
                $script:PB.Drag = @{ Start = $e.GetPosition($script:PB.Window); Source = $el; Kind = $kind; Name = [string]$name; From = $from; Active = $false }
}

function Invoke-PBMouseMove {
                param ($e)

                $d = $script:PB.Drag
                if ($null -eq $d -or $d.Active) { return }
                if ($e.LeftButton -ne [System.Windows.Input.MouseButtonState]::Pressed) { $script:PB.Drag = $null; return }
                $p = $e.GetPosition($script:PB.Window)
                if ([math]::Abs($p.X - $d.Start.X) -lt [System.Windows.SystemParameters]::MinimumHorizontalDragDistance -and
                                [math]::Abs($p.Y - $d.Start.Y) -lt [System.Windows.SystemParameters]::MinimumVerticalDragDistance) { return }

                $d.Active = $true
                $script:PB.PendingDrop = $null
                $payload = "{0}|{1}|{2}" -f $d.Kind, $d.Name, $d.From
                $data = New-Object System.Windows.DataObject('FAIS.Puck', $payload)
                if ($d.Kind -eq 'Student') { Start-PBHeat $d }
                try {
                                [void][System.Windows.DragDrop]::DoDragDrop($d.Source, $data, ([System.Windows.DragDropEffects]::Copy -bor [System.Windows.DragDropEffects]::Move))
                }
                finally {
                                Stop-PBHeat
                                $script:PB.Drag = $null
                }
                if ($null -ne $script:PB.PendingDrop) {
                                $drop = $script:PB.PendingDrop
                                $script:PB.PendingDrop = $null
                                Request-PBCommand $drop
                }
}

# What would happen if the payload were dropped here: a command, or $null
function Get-PBDropCommand {
                param ($e)

                if (-not $e.Data.GetDataPresent('FAIS.Puck')) { return $null }
                $parts = ([string]$e.Data.GetData('FAIS.Puck')) -split '\|', 3
                if ($parts.Count -lt 3) { return $null }
                $kind = $parts[0]; $name = $parts[1]; $from = $parts[2]
                $target = Find-PBTagged $e.OriginalSource @('IPSlot', 'StudentSlot', 'IPPool', 'StudentPool')
                if ($null -eq $target) { return $null }
                $ctrl = ($e.KeyStates -band [System.Windows.DragDropKeyStates]::ControlKey) -eq [System.Windows.DragDropKeyStates]::ControlKey
                $tag = [string]$target.Tag
                $toId = if ($null -ne $target.DataContext -and $null -ne $target.DataContext.PSObject.Properties['LineId']) { [string]$target.DataContext.LineId } else { '' }

                if ($kind -eq 'Student' -and $tag -eq 'StudentSlot') {
                                if (-not $target.DataContext.AllowStudentDrop -or $toId -eq $from) { return $null }
                                if ($from) { return @{ Cmd = 'MoveStudent'; From = $from; LineId = $toId; Name = $name; Copy = $ctrl; Effect = $(if ($ctrl) { 'Copy' } else { 'Move' }); Target = $target } }
                                return @{ Cmd = 'PinStudent'; LineId = $toId; Name = $name; Effect = 'Copy'; Target = $target }
                }
                if ($kind -eq 'Student' -and $tag -eq 'StudentPool' -and $from) {
                                return @{ Cmd = 'RemoveStudent'; LineId = $from; Effect = 'Move'; Target = $target }
                }
                if ($kind -eq 'IP' -and $tag -eq 'IPSlot') {
                                if ($target.DataContext.IsReturn -or $toId -eq $from) { return $null }
                                if ($from) { return @{ Cmd = 'MoveIP'; From = $from; LineId = $toId; Copy = $ctrl; Effect = $(if ($ctrl) { 'Copy' } else { 'Move' }); Target = $target } }
                                return @{ Cmd = 'SetIP'; LineId = $toId; Name = $name; Effect = 'Copy'; Target = $target }
                }
                if ($kind -eq 'IP' -and $tag -eq 'IPPool' -and $from) {
                                return @{ Cmd = 'RemoveIP'; LineId = $from; Effect = 'Move'; Target = $target }
                }
                return $null
}

function Invoke-PBDragOver {
                param ($e, [switch]$Enter)

                $cmd = Get-PBDropCommand $e
                if ($null -eq $cmd) { $e.Effects = [System.Windows.DragDropEffects]::None }
                else {
                                $e.Effects = [System.Windows.DragDropEffects]$cmd.Effect
                                if ($Enter -and ($cmd.Cmd -eq 'PinStudent' -or $cmd.Cmd -eq 'MoveStudent')) {
                                                $name = if ($cmd.Cmd -eq 'PinStudent') { $cmd.Name } else { $cmd.Name }
                                                Set-PBHeatFor -Element $cmd.Target -LineId $cmd.LineId
                                }
                }
                $e.Handled = $true
}

function Invoke-PBDrop {
                param ($e)

                $cmd = Get-PBDropCommand $e
                if ($null -eq $cmd) { $e.Effects = [System.Windows.DragDropEffects]::None; $e.Handled = $true; return }
                $cmd.Remove('Target')
                $e.Effects = [System.Windows.DragDropEffects]$cmd.Effect
                $e.Handled = $true
                # Applied after DoDragDrop returns (Invoke-PBMouseMove), never inside the drag loop
                $script:PB.PendingDrop = $cmd
}

# Heat: while a student is dragged, slots turn green / amber / red for that student
function Start-PBHeat {
                param ($Drag)

                $map = @{}
                $tab = $script:PB.C.DayTabs.SelectedItem
                if ($null -ne $tab) {
                                foreach ($el in @(Get-PBTaggedDescendants -Root $tab.Content -Tag 'StudentSlot')) {
                                                if ($null -ne $el.DataContext -and $el.DataContext.AllowStudentDrop) { $map[[string]$el.DataContext.LineId] = $el }
                                }
                }
                $queue = New-Object System.Collections.Generic.List[string]
                foreach ($k in $map.Keys) { if ($k -ne $Drag.From) { $queue.Add($k) } }
                $script:PB.Heat = @{ Map = $map; Queue = $queue; Name = $Drag.Name; From = $Drag.From; Touched = (New-Object System.Collections.Generic.List[object]) }
                $script:PB.HeatTimer.Start()
}

function Invoke-PBHeatTick {
                $h = $script:PB.Heat
                if ($null -eq $h) { $script:PB.HeatTimer.Stop(); return }
                $sw = [System.Diagnostics.Stopwatch]::StartNew()
                while ($h.Queue.Count -gt 0 -and $sw.ElapsedMilliseconds -lt 25) {
                                $id = $h.Queue[0]
                                $h.Queue.RemoveAt(0)
                                Set-PBHeatFor -Element $h.Map[$id] -LineId $id
                }
                if ($h.Queue.Count -eq 0) { $script:PB.HeatTimer.Stop() }
}

function Set-PBHeatFor {
                param ($Element, [string]$LineId)

                $h = $script:PB.Heat
                if ($null -eq $h -or $null -eq $Element -or $LineId -eq $h.From) { return }
                $exclude = if ($h.From) { @($h.From) } else { @() }
                $a = Get-PlacementAssessment -Board $script:PB.Board -LineId $LineId -StudentName $h.Name -ExcludeLineIds $exclude -Data $script:PB.Data -Quick
                $color = switch ($a.Status) { 'Clean' { $script:PBColors.HeatClean } 'Warn' { $script:PBColors.HeatWarn } default { $script:PBColors.HeatHard } }
                $Element.SetCurrentValue([System.Windows.Controls.Border]::BackgroundProperty, (Get-PBBrush $color))
                if (-not $h.Touched.Contains($Element)) { $h.Touched.Add($Element) }
}

function Stop-PBHeat {
                $script:PB.HeatTimer.Stop()
                $h = $script:PB.Heat
                $script:PB.Heat = $null
                if ($null -eq $h) { return }
                foreach ($el in $h.Touched) {
                                $be = $el.GetBindingExpression([System.Windows.Controls.Border]::BackgroundProperty)
                                if ($null -ne $be) { $be.UpdateTarget() }
                }
}


# =================================================================================
#  KEYBOARD, TYPING AND COMBOS
# =================================================================================

function Invoke-PBKeyDown {
                param ($e)

                $C = $script:PB.C
                $src = $e.OriginalSource
                $key = $e.Key
                $mods = [System.Windows.Input.Keyboard]::Modifiers
                $ctrl = ($mods -band [System.Windows.Input.ModifierKeys]::Control) -eq [System.Windows.Input.ModifierKeys]::Control
                $shift = ($mods -band [System.Windows.Input.ModifierKeys]::Shift) -eq [System.Windows.Input.ModifierKeys]::Shift
                $KeyT = [System.Windows.Input.Key]

                if ([object]::ReferenceEquals($src, $C.QuickEntryBox) -and $key -eq $KeyT::Enter) {
                                Invoke-PBQuickEntry
                                $e.Handled = $true
                                return
                }

                $tb = $src -as [System.Windows.Controls.TextBox]
                if ($null -ne $tb -and $script:PBCellTags -contains [string]$tb.Tag) {
                                if ([string]$tb.Tag -eq 'EtdBox' -and ($key -eq $KeyT::Up -or $key -eq $KeyT::Down)) {
                                                $t = ConvertTo-BoardTime $tb.Text
                                                if ($null -eq $t) { $t = ConvertTo-BoardTime $tb.DataContext.EtdText }
                                                $step = if ($shift) { 30 } else { 5 }
                                                if ($key -eq $KeyT::Down) { $step = -$step }
                                                $mins = [int](($t.TotalMinutes + $step) % 1440)
                                                if ($mins -lt 0) { $mins += 1440 }
                                                $tb.Text = Format-BoardTime ([TimeSpan]::FromMinutes($mins))
                                                $tb.SelectAll()
                                                $e.Handled = $true
                                                return
                                }
                                if ($key -eq $KeyT::Enter) { Submit-PBCell $tb; $e.Handled = $true; return }
                                if ($key -eq $KeyT::Escape) { Reset-PBCell $tb; $e.Handled = $true; return }
                                return
                }
                if ($src -is [System.Windows.Controls.Primitives.TextBoxBase]) { return }

                if ($ctrl -and $key -eq $KeyT::Z) { Request-PBCommand @{ Cmd = 'Undo' }; $e.Handled = $true; return }
                if ($ctrl -and $key -eq $KeyT::L) { [void]$C.QuickEntryBox.Focus(); $C.QuickEntryBox.SelectAll(); $e.Handled = $true; return }
                if ($ctrl -and $key -eq $KeyT::Enter) { Request-PBCommand @{ Cmd = 'Fill' }; $e.Handled = $true; return }
                if ($key -eq $KeyT::F5) {
                                $script:PB.Assess = @{}
                                Clear-AssessmentCache
                                Start-PBAssessQueue
                                Set-PBStatus 'Re-checking every hardcoded line...'
                                $e.Handled = $true
                                return
                }
                if ($key -eq $KeyT::Delete) {
                                $slot = Find-PBTagged ([System.Windows.Input.Keyboard]::FocusedElement) @('IPSlot', 'StudentSlot')
                                if ($null -ne $slot) {
                                                if ([string]$slot.Tag -eq 'StudentSlot' -and $slot.DataContext.StudentRaw) { Request-PBCommand @{ Cmd = 'RemoveStudent'; LineId = $slot.DataContext.LineId } }
                                                elseif ([string]$slot.Tag -eq 'IPSlot' -and $slot.DataContext.InstructorRaw) { Request-PBCommand @{ Cmd = 'RemoveIP'; LineId = $slot.DataContext.LineId } }
                                                $e.Handled = $true
                                }
                }
}

function Invoke-PBLostFocus {
                param ($e)

                if ($script:PB.Busy) { return }
                $tb = $e.OriginalSource -as [System.Windows.Controls.TextBox]
                if ($null -ne $tb -and $script:PBCellTags -contains [string]$tb.Tag) { Submit-PBCell $tb }
}

# Commits a typed ETD / callsign / num / notes (does nothing if unchanged)
function Submit-PBCell {
                param ($TextBox)

                $vm = $TextBox.DataContext
                if ($null -eq $vm -or $null -eq $vm.PSObject.Properties['LineId'] -or $vm.ReadOnlyFields) { return }
                $line = Get-BoardLine $script:PB.Board $vm.LineId
                if ($null -eq $line) { return }
                $text = ([string]$TextBox.Text).Trim()

                switch ([string]$TextBox.Tag) {
                                'EtdBox' {
                                                $t = ConvertTo-BoardTime $text
                                                if ($null -eq $t) {
                                                                $TextBox.BorderBrush = Get-PBBrush '#DC2626'
                                                                Set-PBStatus "'$text' is not a time. Type 0730, 730 or 7:30." 'Warn'
                                                                return
                                                }
                                                $newEtd = $line.ETD.Date.Add($t)
                                                if ($newEtd -ne $line.ETD) { Request-PBCommand @{ Cmd = 'SetEtd'; LineId = $line.Id; Value = $newEtd } }
                                                elseif ($text -ne $vm.EtdText) { $TextBox.Text = $vm.EtdText }
                                }
                                'CallsignBox' { if ($text -ne [string]$line.Callsign) { Request-PBCommand @{ Cmd = 'SetField'; LineId = $line.Id; Field = 'Callsign'; Value = $text } } }
                                'NumBox' { if ($text -ne [string]$line.Num) { Request-PBCommand @{ Cmd = 'SetField'; LineId = $line.Id; Field = 'Num'; Value = $text } } }
                                'NotesBox' { if ($text -ne ([string]$line.Notes).Trim()) { Request-PBCommand @{ Cmd = 'SetField'; LineId = $line.Id; Field = 'Notes'; Value = $text } } }
                }
}

function Reset-PBCell {
                param ($TextBox)

                $vm = $TextBox.DataContext
                if ($null -eq $vm) { return }
                switch ([string]$TextBox.Tag) {
                                'EtdBox' { $TextBox.Text = $vm.EtdText }
                                'CallsignBox' { $TextBox.Text = $vm.CallsignText }
                                'NumBox' { $TextBox.Text = $vm.NumText }
                                'NotesBox' { $TextBox.Text = $vm.NotesText }
                }
                $TextBox.ClearValue([System.Windows.Controls.Control]::BorderBrushProperty)
}

function Invoke-PBSelectionChanged {
                param ($e)

                if ($script:PB.Busy) { return }
                $src = $e.OriginalSource
                if ([object]::ReferenceEquals($src, $script:PB.C.DayTabs)) {
                                $script:PB.LastFocusDay = Get-PBActiveDay
                                Update-PBPools
                                Update-PBCounts
                                return
                }
                $cb = $src -as [System.Windows.Controls.ComboBox]
                if ($null -eq $cb -or $e.AddedItems.Count -eq 0) { return }
                $vm = $cb.DataContext
                if ($null -eq $vm -or $null -eq $vm.PSObject.Properties['LineId']) { return }
                $value = [string]$e.AddedItems[0]
                switch ([string]$cb.Tag) {
                                'TypeBox' {
                                                if ($value -and $value -ne $vm.SortieType) { Request-PBCommand @{ Cmd = 'SetType'; LineId = $vm.LineId; Value = $value } }
                                }
                                'MissionBox' {
                                                $mission = if ($value -eq $script:PBAutoText) { '' } else { $value }
                                                if ($vm.MissionEnabled -and $mission -ne $vm.MissionRaw) { Request-PBCommand @{ Cmd = 'SetMission'; LineId = $vm.LineId; Value = $mission } }
                                }
                }
}

function Invoke-PBRowClick {
                param ($e)

                $btn = $e.OriginalSource -as [System.Windows.Controls.Button]
                if ($null -eq $btn -or [string]$btn.Tag -ne 'DeleteLine') { return }
                $vm = $btn.DataContext
                if ($null -ne $vm -and $null -ne $vm.PSObject.Properties['LineId']) {
                                Request-PBCommand @{ Cmd = 'DeleteLine'; LineId = $vm.LineId }
                                $e.Handled = $true
                }
}


# =================================================================================
#  TOOLTIPS, POP-UPS AND MENUS
# =================================================================================

# Coloured lines for an assessment: X hard rule, ! warning, + and - score reasons
function New-PBAssessmentPanel {
                param ($Assessment, [string]$Footer)

                $panel = New-Object System.Windows.Controls.StackPanel
                $panel.MaxWidth = 460
                foreach ($ln in @(Get-AssessmentLines $Assessment)) {
                                $tb = New-Object System.Windows.Controls.TextBlock
                                $tb.TextWrapping = [System.Windows.TextWrapping]::Wrap
                                $tb.Margin = New-Object System.Windows.Thickness(0, 1, 0, 1)
                                switch ($ln.Kind) {
                                                'Title' { $tb.Text = $ln.Text; $tb.FontWeight = [System.Windows.FontWeights]::Bold; $tb.Margin = New-Object System.Windows.Thickness(0, 0, 0, 4) }
                                                'Hard' { $tb.Text = "$([char]0x2716)  $($ln.Text)"; $tb.Foreground = Get-PBBrush '#B91C1C'; $tb.FontWeight = [System.Windows.FontWeights]::SemiBold }
                                                'Warn' { $tb.Text = "$([char]0x26A0)  $($ln.Text)"; $tb.Foreground = Get-PBBrush '#B45309' }
                                                'Note' { $tb.Text = "$([char]0x2022)  $($ln.Text)"; $tb.Foreground = Get-PBBrush '#92400E' }
                                                'Plus' { $tb.Text = "+  $($ln.Text)"; $tb.Foreground = Get-PBBrush '#15803D' }
                                                'Minus' { $tb.Text = "$([char]0x2212)  $($ln.Text)"; $tb.Foreground = Get-PBBrush '#B91C1C' }
                                                default { $tb.Text = $ln.Text; $tb.Foreground = Get-PBBrush '#64748B'; $tb.FontStyle = [System.Windows.FontStyles]::Italic; $tb.FontSize = 11 }
                                }
                                [void]$panel.Children.Add($tb)
                }
                if ($Assessment.Status -eq 'Hard') {
                                $tb = New-Object System.Windows.Controls.TextBlock
                                $tb.Text = 'Allowed: it stays hardcoded and the line is flagged OVERRIDE on the output.'
                                $tb.TextWrapping = [System.Windows.TextWrapping]::Wrap
                                $tb.Foreground = Get-PBBrush '#64748B'
                                $tb.FontSize = 11
                                $tb.Margin = New-Object System.Windows.Thickness(0, 4, 0, 0)
                                [void]$panel.Children.Add($tb)
                }
                if ($Footer) {
                                $tb = New-Object System.Windows.Controls.TextBlock
                                $tb.Text = $Footer
                                $tb.Foreground = Get-PBBrush '#64748B'
                                $tb.FontSize = 11
                                [void]$panel.Children.Add($tb)
                }
                return $panel
}

function New-PBTextPanel {
                param ([string[]]$Lines, [string]$Title)

                $panel = New-Object System.Windows.Controls.StackPanel
                $panel.MaxWidth = 460
                if ($Title) {
                                $tb = New-Object System.Windows.Controls.TextBlock
                                $tb.Text = $Title
                                $tb.FontWeight = [System.Windows.FontWeights]::Bold
                                [void]$panel.Children.Add($tb)
                }
                foreach ($l in $Lines) {
                                if (-not $l) { continue }
                                $tb = New-Object System.Windows.Controls.TextBlock
                                $tb.Text = $l
                                $tb.TextWrapping = [System.Windows.TextWrapping]::Wrap
                                if ($l -like '*OVERRIDE*') { $tb.Foreground = Get-PBBrush '#B91C1C' }
                                [void]$panel.Children.Add($tb)
                }
                return $panel
}

# Full assessment for a pinned line (cached per board version)
function Get-PBLineAssessment {
                param ($Line)

                if (-not (Test-PBAssessmentFresh $Line)) {
                                $script:PB.Assess[$Line.Id] = Get-PlacementAssessment -Board $script:PB.Board -LineId $Line.Id -StudentName $Line.Student `
                                                -MissionName ([string]$Line.Mission) -Data $script:PB.Data
                }
                return $script:PB.Assess[$Line.Id]
}

function Invoke-PBToolTipOpening {
                param ($e)

                $el = Find-PBTagged $e.OriginalSource @('StudentSlot')
                if ($null -eq $el -or -not [object]::ReferenceEquals($el, $e.OriginalSource)) { return }
                $vm = $el.DataContext
                $line = if ($null -ne $vm) { Get-BoardLine $script:PB.Board $vm.LineId } else { $null }
                if ($null -eq $line) { $e.Handled = $true; return }

                $content = $null
                if ($line.Student) {
                                $content = New-PBAssessmentPanel (Get-PBLineAssessment $line)
                }
                elseif ($script:PB.Mode -eq 'Review' -and $null -ne $script:PB.Overlay -and $null -ne $script:PB.Overlay[$line.Id]) {
                                $ov = $script:PB.Overlay[$line.Id]
                                $by = switch ($ov.FilledBy) { 'Engine' { 'Picked by the scheduler' } 'Wing' { 'Wingman picked by the scheduler' } 'Return' { 'Return leg of the O&B' } 'DS' { 'Direct support for a formation' } 'ForSale' { 'No student could fly this line' } default { $ov.FilledBy } }
                                $content = New-PBTextPanel -Title ("{0} {1}" -f $ov.Student, $ov.Mission).Trim() -Lines (@($by) + @(([string]$ov.Notes) -split '\s*\|\s*'))
                }
                elseif (Test-IsReturnSortieType ([string]$line.SortieType)) {
                                $content = New-PBTextPanel -Lines @('O&B Return lines are filled with the O&B''s student when the schedule is filled.')
                }
                if ($null -eq $content) { $e.Handled = $true; return }
                $el.ToolTip = $content
}

function Show-PBPopupFor {
                param ([string]$LineId, $Content)

                $C = $script:PB.C
                $target = Get-PBSlotElement -LineId $LineId -Tag 'StudentSlot'
                if ($null -eq $target) { return }
                $C.DropPopupPanel.Children.Clear()
                [void]$C.DropPopupPanel.Children.Add($Content)
                $C.DropPopup.PlacementTarget = $target
                $C.DropPopup.IsOpen = $false
                $C.DropPopup.IsOpen = $true
                $script:PB.PopupTimer.Stop()
                $script:PB.PopupTimer.Start()
}

# Pop-up with the + / - reasons and rule flags right after a student lands on a line
function Show-PBPlacementPopup {
                param ([string]$LineId)

                $line = Get-BoardLine $script:PB.Board $LineId
                if ($null -eq $line -or -not $line.Student) { return }
                $script:PB.Window.UpdateLayout()
                $a = Get-PBLineAssessment $line
                [void](Update-PBLineRow $LineId)
                Update-PBCounts
                $script:PB.Window.UpdateLayout()
                Show-PBPopupFor -LineId $LineId -Content (New-PBAssessmentPanel $a -Footer 'Hover the corner marker any time to see this again.')
                $kind = switch ($a.Status) { 'Hard' { 'Hard' } 'Warn' { 'Warn' } default { 'Good' } }
                Set-PBStatus $a.Summary $kind
}

function New-PBMenuItem {
                param ([string]$Header, [hashtable]$Command)

                $mi = New-Object System.Windows.Controls.MenuItem
                $mi.Header = $Header
                if ($null -ne $Command) {
                                $mi.Tag = $Command
                                $mi.Add_Click({ param($s, $a) Invoke-PBSafe { Request-PBCommand $s.Tag } })
                }
                return $mi
}

function Invoke-PBContextMenuOpening {
                param ($e)

                $el = Find-PBTagged $e.OriginalSource
                if ($null -eq $el) { return }
                $tag = [string]$el.Tag
                if (@('StudentSlot', 'IPSlot', 'Row') -notcontains $tag) { return }
                $menu = $el.ContextMenu
                if ($null -eq $menu) { return }
                $vm = $el.DataContext
                $line = if ($null -ne $vm) { Get-BoardLine $script:PB.Board $vm.LineId } else { $null }
                if ($null -eq $line) { $e.Handled = $true; return }
                Set-PBMenuItems -Menu $menu -Tag $tag -Line $line
}

# Fills a slot's / row's right-click menu
function Set-PBMenuItems {
                param ($Menu, [string]$Tag, $Line)

                $Menu.Items.Clear()
                $id = $Line.Id
                $review = ($script:PB.Mode -eq 'Review')
                $canFormate = Test-IsWingCapableType ([string]$Line.SortieType)

                if ($Tag -eq 'StudentSlot') {
                                if ($Line.Student) {
                                                [void]$Menu.Items.Add((New-PBMenuItem "Explain $($Line.Student) on this line" @{ Cmd = 'Explain'; LineId = $id }))
                                                [void]$Menu.Items.Add((New-PBMenuItem 'Take student off this line' @{ Cmd = 'RemoveStudent'; LineId = $id }))
                                                # Formation pairing is only for Day Local lines: OFT / IFT can never fly wing
                                                if ($canFormate) {
                                                                $wingMenu = New-PBMenuItem 'Fly wing to...' $null
                                                                $leads = @(Get-BoardDayLines -Board $script:PB.Board -Day $Line.Day | Where-Object {
                                                                                $_.Id -ne $id -and $_.Student -and (Test-IsWingCapableType ([string]$_.SortieType))
                                                                })
                                                                foreach ($l in $leads) {
                                                                                [void]$wingMenu.Items.Add((New-PBMenuItem ("{0} {1} ({2})" -f $l.ETD.ToString('HHmm'), $l.Student, $l.Instructor) @{ Cmd = 'WingTo'; LineId = $id; Arg = $l.Id }))
                                                                }
                                                                $wingMenu.IsEnabled = ($leads.Count -gt 0)
                                                                [void]$Menu.Items.Add($wingMenu)
                                                }
                                                if ($Line.WingOfLineId) { [void]$Menu.Items.Add((New-PBMenuItem 'Clear wing pairing' @{ Cmd = 'ClearWing'; LineId = $id })) }
                                }
                                elseif ($review -and $null -ne $script:PB.Overlay -and $null -ne $script:PB.Overlay[$id] -and $script:PB.Overlay[$id].FilledBy -in @('Engine', 'Wing')) {
                                                $ov = $script:PB.Overlay[$id]
                                                [void]$Menu.Items.Add((New-PBMenuItem "Hardcode $($ov.Student) on $($ov.Mission) here" @{ Cmd = 'PinStudent'; LineId = $id; Name = $ov.Student; Mission = $ov.Mission }))
                                }
                }
                if ($Tag -eq 'IPSlot' -and $Line.Instructor) {
                                [void]$Menu.Items.Add((New-PBMenuItem "Take $($Line.Instructor) off this line" @{ Cmd = 'RemoveIP'; LineId = $id }))
                }
                if ($Menu.Items.Count -gt 0) { [void]$Menu.Items.Add((New-Object System.Windows.Controls.Separator)) }
                [void]$Menu.Items.Add((New-PBMenuItem 'Duplicate line' @{ Cmd = 'DuplicateLine'; LineId = $id }))
                if ($Line.LinkedLineId) { [void]$Menu.Items.Add((New-PBMenuItem 'Unlink O&B / return' @{ Cmd = 'Unlink'; LineId = $id })) }
                [void]$Menu.Items.Add((New-PBMenuItem 'Delete line' @{ Cmd = 'DeleteLine'; LineId = $id }))
}


# =================================================================================
#  COMMANDS (every board edit lands here)
# =================================================================================

function Invoke-PBCommand {
                param ([hashtable]$Cmd)

                $board = $script:PB.Board
                $id = [string]$Cmd.LineId
                switch ($Cmd.Cmd) {
                                'PinStudent' {
                                                $mission = if ($Cmd.ContainsKey('Mission')) { [string]$Cmd.Mission } else { '' }
                                                Invoke-PBChange { [void](Set-BoardPin -Board $board -Id $id -Student $Cmd.Name -Mission $mission) }
                                                Show-PBPlacementPopup $id
                                }
                                'MoveStudent' {
                                                Invoke-PBChange { [void](Move-BoardPin -Board $board -FromId $Cmd.From -ToId $id -Copy:([bool]$Cmd.Copy)) }
                                                Show-PBPlacementPopup $id
                                }
                                'RemoveStudent' { Invoke-PBChange { Clear-BoardPin -Board $board -Id $id } -Status 'Student taken off the line.' }
                                'SetIP' { Invoke-PBChange { Set-BoardInstructor -Board $board -Id $id -Name $Cmd.Name } -Status "$($Cmd.Name) put on the line." }
                                'MoveIP' { Invoke-PBChange { [void](Move-BoardInstructor -Board $board -FromId $Cmd.From -ToId $id -Copy:([bool]$Cmd.Copy)) } -Status 'IP moved.' }
                                'RemoveIP' { Invoke-PBChange { Set-BoardInstructor -Board $board -Id $id -Name '' } -Status 'IP taken off the line.' }
                                'SetEtd' { Invoke-PBChange { Set-BoardLineEtd -Board $board -Id $id -ETD $Cmd.Value } -Status ("ETD set to {0}." -f ([datetime]$Cmd.Value).ToString('HHmm')) }
                                'SetType' { Invoke-PBChange { Set-BoardLineType -Board $board -Id $id -SortieType $Cmd.Value } -Status "Line changed to $($Cmd.Value)." }
                                'SetMission' {
                                                Invoke-PBChange { Set-BoardMission -Board $board -Id $id -Mission $Cmd.Value }
                                                Show-PBPlacementPopup $id
                                }
                                'SetField' {
                                                # Notes / callsign don't change any rule, so only this row is redrawn
                                                Set-BoardLineField -Board $board -Id $id -Field $Cmd.Field -Value $Cmd.Value
                                                [void](Update-PBLineRow $id)
                                                Request-PBAutosave
                                }
                                'DeleteLine' { Invoke-PBChange { Remove-BoardLine -Board $board -Id $id } -Status 'Line deleted (Ctrl+Z to undo).' }
                                'DuplicateLine' {
                                                $l = Get-BoardLine $board $id
                                                if ($null -ne $l) {
                                                                Invoke-PBChange {
                                                                                $copy = New-BoardLine -ETD $l.ETD -SortieType $l.SortieType -Instructor $l.Instructor
                                                                                $new = @($copy)
                                                                                if (Test-IsOBSortieType ([string]$l.SortieType)) {
                                                                                                $ret = if ($l.LinkedLineId) { Get-BoardLine $board $l.LinkedLineId } else { $null }
                                                                                                $retEtd = if ($null -ne $ret) { $ret.ETD } else { $null }
                                                                                                $new += New-ReturnLineFor -Board $board -OBLine $copy -ReturnETD $retEtd
                                                                                }
                                                                                Add-BoardLine -Board $board -Line $new
                                                                } -Status 'Line duplicated.'
                                                }
                                }
                                'Unlink' { Invoke-PBChange { Unlink-ReturnLine -Board $board -Id $id } -Status 'O&B and return unlinked.' }
                                'WingTo' { Invoke-PBChange { [void](Set-FormationPair -Board $board -WingId $id -LeadId $Cmd.Arg) } -Status 'Wing pairing set.' }
                                'ClearWing' { Invoke-PBChange { [void](Set-FormationPair -Board $board -WingId $id -LeadId '') } -Status 'Wing pairing cleared.' }
                                'Explain' { Show-PBPlacementPopup $id }
                                'AddLines' {
                                                $lines = @($Cmd.Lines)
                                                Invoke-PBChange { Add-BoardLine -Board $board -Line $lines } -Status $Cmd.Status
                                }
                                'Undo' {
                                                if ($script:PB.Mode -eq 'Review') { Exit-PBReview }
                                                if (Undo-Board $board) {
                                                                Update-PBAll; Request-PBAutosave; Start-PBAssessQueue
                                                                Set-PBStatus 'Undone.'
                                                }
                                                else { Set-PBStatus 'Nothing to undo.' 'Warn' }
                                }
                                'CopyPrevDay' { Invoke-PBCopyPrevDay }
                                'Fill' { Invoke-PBFill }
                }
}


# =================================================================================
#  TOOLBAR ACTIONS
# =================================================================================

function Invoke-PBQuickEntry {
                $C = $script:PB.C
                $text = [string]$C.QuickEntryBox.Text
                if ([string]::IsNullOrWhiteSpace($text)) { return }
                $day = Get-PBActiveDay
                if ($null -eq $day) { Set-PBStatus 'Add a day first.' 'Warn'; return }

                $r = ConvertFrom-QuickEntry -Text $text -Day $day -Instructors $script:PB.Data.Instructors -Students $script:PB.Data.Students -Board $script:PB.Board
                if ($r.Errors.Count -gt 0) {
                                $C.QuickEntryFeedback.Foreground = Get-PBBrush '#FCA5A5'
                                $C.QuickEntryFeedback.Text = ($r.Errors -join '  |  ')
                                return
                }
                if ($r.Lines.Count -eq 0) { return }
                $C.QuickEntryBox.Text = ''
                $msg = "Added $($r.Lines.Count) line(s)."
                if ($r.Warnings.Count -gt 0) { $msg += '  ' + ($r.Warnings -join '  |  ') }
                $C.QuickEntryFeedback.Foreground = Get-PBBrush $(if ($r.Warnings.Count -gt 0) { '#FDE68A' } else { '#86EFAC' })
                $C.QuickEntryFeedback.Text = $msg
                Request-PBCommand @{ Cmd = 'AddLines'; Lines = $r.Lines; Status = $msg }
}

function Update-PBTemplateList {
                $script:PB.C.TemplateCombo.ItemsSource = @(Get-BoardTemplateList -Dir $script:PB.Settings.TemplateDir)
}

function Invoke-PBLoadTemplate {
                $t = $script:PB.C.TemplateCombo.SelectedItem
                if ($null -eq $t) { Set-PBStatus 'Pick a template first.' 'Warn'; return }
                $day = Get-PBActiveDay
                $hasLines = @($script:PB.Board.Lines | Where-Object { $_.Day -eq $day }).Count -gt 0
                $replace = $false
                if ($hasLines) {
                                $ans = [System.Windows.MessageBox]::Show("Replace this day's lines with '$($t.Name)'?`n`nYes = replace the day`nNo = add the template's lines to the day",
                                                'Load template', [System.Windows.MessageBoxButton]::YesNoCancel, [System.Windows.MessageBoxImage]::Question)
                                if ($ans -eq [System.Windows.MessageBoxResult]::Cancel) { return }
                                $replace = ($ans -eq [System.Windows.MessageBoxResult]::Yes)
                }
                $path = $t.Path
                Invoke-PBChange { [void](Import-BoardTemplate -Board $script:PB.Board -Day $day -Path $path -Replace:$replace) } -Status "Template '$($t.Name)' loaded."
}

function Invoke-PBSaveTemplate {
                $day = Get-PBActiveDay
                if (@($script:PB.Board.Lines | Where-Object { $_.Day -eq $day }).Count -eq 0) { Set-PBStatus 'This day has no lines to save.' 'Warn'; return }
                $name = [Microsoft.VisualBasic.Interaction]::InputBox('Template name (e.g. Mon 3-turn):', 'Save day as template', '')
                if ([string]::IsNullOrWhiteSpace($name)) { return }
                $ans = [System.Windows.MessageBox]::Show('Save the IP on each line too?', 'Save template', [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
                $path = Save-BoardTemplate -Board $script:PB.Board -Day $day -Name $name -Dir $script:PB.Settings.TemplateDir -IncludeInstructors:($ans -eq [System.Windows.MessageBoxResult]::Yes)
                Update-PBTemplateList
                Set-PBStatus "Template saved: $path" 'Good'
}

function Invoke-PBCopyPrevDay {
                $day = Get-PBActiveDay
                if ($null -eq $day) { return }
                $holder = @{ Res = $null }
                Invoke-PBChange {
                                $holder.Res = Copy-BoardPreviousDay -Board $script:PB.Board -TargetDay $day -Dir $script:PB.Settings.BoardDir -PrevSchedule @($script:PB.Data.PrevSchedule) -IncludeInstructors
                }
                $res = $holder.Res
                if ($null -eq $res) { Set-PBStatus 'No earlier day found to copy (looked back 7 days and at the last MasterSchedule).' 'Warn' }
                else { Set-PBStatus "Copied $($res.Count) line(s) from $($res.Source). Students were not copied." 'Good' }
}

function Invoke-PBImportSheet {
                $rows = @($script:PB.Data.ImportSchedule)
                if ($rows.Count -eq 0) { return }
                $ans = [System.Windows.MessageBox]::Show("Add the $($rows.Count) line(s) from the MasterData Schedule sheet to the board (students on them stay hardcoded)?",
                                'Import Schedule sheet', [System.Windows.MessageBoxButton]::OKCancel, [System.Windows.MessageBoxImage]::Question)
                if ($ans -ne [System.Windows.MessageBoxResult]::OK) { return }
                $lines = @(ConvertFrom-ScheduleRows -Rows $rows -Source 'Imported' -KeepStudents -KeepNotes)
                Invoke-PBChange { Add-BoardLine -Board $script:PB.Board -Line $lines } -Status "Imported $($lines.Count) line(s) from the Schedule sheet."
                Reset-PBPlannedFromBoard
}

function Invoke-PBAddDay {
                $date = $script:PB.C.AddDayPicker.SelectedDate
                if ($null -eq $date) { Set-PBStatus 'Pick a date first.' 'Warn'; return }
                $day = ([datetime]$date).ToString('yyyy-MM-dd')
                Invoke-PBChange { [void](Add-BoardDay -Board $script:PB.Board -Day $day) }
                Reset-PBPlannedFromBoard
                foreach ($t in $script:PB.C.DayTabs.Items) { if ([string]$t.Tag -eq $day) { $script:PB.C.DayTabs.SelectedItem = $t } }
                $script:PB.C.AddDayPicker.SelectedDate = (Get-NextFlyingDay -From ([datetime]$date))
                Set-PBStatus "Added $day." 'Good'
}

# A newly added earlier day: events planned on/after it are plans, not flown events
function Reset-PBPlannedFromBoard {
                if ($script:PB.Board.Days.Count -eq 0) { return }
                $first = [datetime]::ParseExact($script:PB.Board.Days[0], 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
                $changed = @(Reset-PlannedCompletions -Students $script:PB.Data.Students -FromDate $first -SoftWeights $script:PB.Data.SoftWeights)
                if ($changed.Count -gt 0) {
                                Clear-AssessmentCache
                                $script:PB.Assess = @{}
                                Update-PBAll
                                Start-PBAssessQueue
                }
}

function Invoke-PBRemoveDay {
                $day = Get-PBActiveDay
                if ($null -eq $day) { return }
                $n = @($script:PB.Board.Lines | Where-Object { $_.Day -eq $day }).Count
                $ans = [System.Windows.MessageBox]::Show("Remove $day and its $n line(s) from the board?", 'Remove day', [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
                if ($ans -ne [System.Windows.MessageBoxResult]::Yes) { return }
                Invoke-PBChange { Remove-BoardDay -Board $script:PB.Board -Day $day } -Status "$day removed (Ctrl+Z to undo)."
                $file = Join-Path $script:PB.Settings.BoardDir "$day.json"
                if (Test-Path -LiteralPath $file) { Remove-Item -LiteralPath $file -Force -ErrorAction SilentlyContinue }
}

function Show-BusyOverlay {
                param ([string]$Text)

                $C = $script:PB.C
                $C.BusyText.Text = $Text
                $C.BusyProgress.Value = 0
                $C.BusyOverlay.Visibility = 'Visible'
                Invoke-PBDoEvents
}

function Hide-BusyOverlay {
                $script:PB.C.BusyOverlay.Visibility = 'Collapsed'
}

# "Fill the rest of the lines": runs the scheduler on a copy of the students and shows
# the result on the board for review. Nothing is written until Export.
function Invoke-PBFill {
                $board = $script:PB.Board
                $warnings = @(Test-BoardReadyToFill -Board $board -Students $script:PB.Data.Students -Instructors $script:PB.Data.Instructors)
                if ($board.Lines.Count -eq 0) { Set-PBStatus 'There are no lines to fill. Type some in the Add lines box.' 'Warn'; return }
                if ($warnings.Count -gt 0) {
                                $shown = @($warnings | Select-Object -First 15)
                                $more = if ($warnings.Count -gt 15) { "`n...and $($warnings.Count - 15) more" } else { '' }
                                $ans = [System.Windows.MessageBox]::Show(("Before filling:`n`n- " + ($shown -join "`n- ") + $more + "`n`nFill anyway?"),
                                                'Fill the rest of the lines', [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
                                if ($ans -ne [System.Windows.MessageBoxResult]::Yes) { return }
                }

                Exit-PBReview
                Show-BusyOverlay 'Filling the schedule...'
                try {
                                $rows = @(ConvertTo-ScheduleRows -Board $board)
                                $students = @(Copy-StudentList $script:PB.Data.Students)
                                $progress = {
                                                param($i, $n, $label)
                                                $script:PB.C.BusyProgress.Value = [int](100 * $i / [math]::Max(1, $n))
                                                $script:PB.C.BusyText.Text = "Filling line $i of $n  ($label)"
                                                Invoke-PBDoEvents
                                }
                                $result = Invoke-FillSchedule -Schedule $rows -Students $students -Instructors $script:PB.Data.Instructors `
                                                -SoftWeights $script:PB.Data.SoftWeights -Formations @($script:PB.Data.Formations) -OnProgress $progress -Trace:([bool]$script:PB.Settings.Trace)
                }
                finally { Hide-BusyOverlay }

                $script:PB.FillResult = $result
                $script:PB.Overlay = Get-FillOverlay -FillResult $result
                Set-PBMode 'Review'
                Update-PBAll
                $st = $result.Stats
                $kind = if ($st.Overrides -gt 0 -or $st.ForSale -gt 0) { 'Warn' } else { 'Good' }
                Set-PBStatus ("Filled: {0} by the scheduler, {1} hardcoded, {2} wing, {3} return legs, {4} for sale, {5} override flag(s). Check the board, then Export." -f $st.Filled, $st.Pinned, $st.Wing, $st.Return, $st.ForSale, $st.Overrides) $kind
}

function Invoke-PBExport {
                if ($script:PB.Mode -ne 'Review' -or $null -eq $script:PB.FillResult) { Set-PBStatus 'Fill the lines first.' 'Warn'; return }
                Save-PBBoard
                $script:PB.Result = [PSCustomObject]@{ Action = 'Export'; Board = $script:PB.Board; FillResult = $script:PB.FillResult }
                try { $script:PB.Window.DialogResult = $true } catch { $script:PB.Window.Close() }
}

function Invoke-PBCancel {
                if ($script:PB.Mode -eq 'Review') {
                                $ans = [System.Windows.MessageBox]::Show("Close without exporting? The board is saved; the fill is not.", 'Close', [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
                                if ($ans -ne [System.Windows.MessageBoxResult]::Yes) { return }
                }
                $script:PB.Result = $null
                $script:PB.Window.Close()
}

function Show-PBHelp {
                $text = @"
ADDING LINES (box at the top, press Enter)
  0730 DL                 Day Local at 0730   (730, 7:30 also work)
  1015 O&B                O&B plus its return leg 3 hr later
  1015 OB 1330            O&B with the return at 1330
  1330 RTN                return line, linked to the latest O&B before it
  1300 OFT x2             two OFT lines (IFT works the same)
  0800 DL Jones sp:Smith  IP Jones, student Smith hardcoded (m:FT5103 sets the mission)
  Several at once:        0730 DL, 0745 DL, 0800 DL
Templates save a day's times; "Copy previous day's lines" copies the last flying day.

PUCKS
  Drag an IP puck onto a line's IP slot, a student puck onto a line's student slot.
  Drag between lines to move (hold Ctrl to copy). Drop back on a pool to take it off.
  While dragging a student, slots turn green (ok), amber (needs a look), red (breaks a rule).
  Rule breaks are allowed - the line is flagged OVERRIDE on the output.

MARKERS
  The corner triangle on a hardcoded student: green clean, amber warnings, red rule broken,
  grey still checking. Hover it for the + / - reasons. Right-click for more.

TIMES
  Click an ETD, type a new time, Enter. Up/Down = 5 min, Shift = 30 min, Esc cancels.

KEYS
  Ctrl+L quick entry   Ctrl+Z undo   Delete clears the selected slot
  F5 re-check everything   Ctrl+Enter fill

FILL
  "Fill the rest of the lines" schedules every open line like before. Review the result,
  then Export (writes OutputSchedule.xlsx), Re-fill, or go Back to edit.
"@
                [void][System.Windows.MessageBox]::Show($text, 'FAIS Puck Board help', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
}


# =================================================================================
#  WINDOWS SELF-TEST
#  Builds the real window off-screen, drives it through the same commands the mouse
#  and keyboard use, fills, and reports problems plus any WPF binding errors.
#  Used by Tests\Invoke-SelfTest.ps1. Returns a list of problems (empty = all good).
# =================================================================================
function Test-PuckBoardWindow {
                param (
                                [Parameter(Mandatory)] $Data,
                                [Parameter(Mandatory)] [string]$Day,
                                [Parameter(Mandatory)] [string]$WorkDir
                )

                $problems = New-Object System.Collections.Generic.List[string]
                Initialize-PuckBoardAssemblies
                $writer = New-Object System.IO.StringWriter
                $listener = New-Object System.Diagnostics.TextWriterTraceListener($writer)
                [System.Diagnostics.PresentationTraceSources]::Refresh()
                $source = [System.Diagnostics.PresentationTraceSources]::DataBindingSource
                [void]$source.Listeners.Add($listener)
                $source.Switch.Level = [System.Diagnostics.SourceLevels]::Warning

                $step = 'start'
                $check = {
                                param([bool]$Ok, [string]$What)
                                if (-not $Ok) { $problems.Add("[$step] $What") }
                                $status = [string]$script:PB.C.StatusText.Text
                                if ($status -like 'Something went wrong*') { $problems.Add("[$step] $status"); $script:PB.C.StatusText.Text = '' }
                }
                try {
                                $step = 'open window'
                                $settings = @{ BoardDir = (Join-Path $WorkDir 'Boards'); TemplateDir = (Join-Path $WorkDir 'Templates') }
                                $w = Initialize-PuckBoardWindow -Data $Data -Days @($Day) -Settings $settings
                                $w.ShowActivated = $false
                                $w.WindowStartupLocation = [System.Windows.WindowStartupLocation]::Manual
                                $w.Left = -20000
                                $w.Top = -20000
                                $w.Show()
                                Invoke-PBDoEvents
                                & $check ($script:PB.C.DayTabs.Items.Count -eq 1) 'one day tab'

                                $step = 'quick entry'
                                $ip1 = Get-NormalizedLastName ([string]$Data.Instructors[0].Name)
                                $ip2 = Get-NormalizedLastName ([string]$Data.Instructors[1].Name)
                                $q = ConvertFrom-QuickEntry -Text "0730 DL $ip1, 0745 DL $ip2, 1015 OB 1315 $ip1, 1300 OFT" -Day $Day -Instructors $Data.Instructors -Students $Data.Students -Board $script:PB.Board
                                & $check ($q.Errors.Count -eq 0) ("quick entry errors: " + ($q.Errors -join '; '))
                                Invoke-PBCommand @{ Cmd = 'AddLines'; Lines = $q.Lines; Status = 'test' }
                                Invoke-PBDoEvents
                                $lines = @(Get-BoardDayLines -Board $script:PB.Board -Day $Day)
                                & $check ($lines.Count -eq 5) "expected 5 lines, got $($lines.Count)"
                                & $check ($script:PB.DayColls[$Day].Count -eq 5) 'rows on screen match lines'
                                & $check ($null -ne (Get-PBSlotElement -LineId $lines[0].Id)) 'student slot rendered'
                                & $check ($null -ne (Get-PBSlotElement -LineId $lines[0].Id -Tag 'IPSlot')) 'IP slot rendered'

                                $step = 'pin student'
                                $sp = [string]$Data.Students[0].Name
                                Invoke-PBCommand @{ Cmd = 'PinStudent'; LineId = $lines[0].Id; Name = $sp }
                                Invoke-PBDoEvents
                                & $check ((Get-BoardLine $script:PB.Board $lines[0].Id).Student -eq $sp) 'student pinned'
                                & $check ([bool]$script:PB.C.DropPopup.IsOpen) 'drop pop-up opened'
                                & $check ($script:PB.C.DropPopupPanel.Children.Count -gt 0) 'drop pop-up has content'

                                $step = 'move, IP, ETD, type, mission, undo'
                                Invoke-PBCommand @{ Cmd = 'MoveStudent'; From = $lines[0].Id; LineId = $lines[1].Id; Name = $sp; Copy = $false }
                                & $check ((Get-BoardLine $script:PB.Board $lines[1].Id).Student -eq $sp) 'student moved'
                                Invoke-PBCommand @{ Cmd = 'Undo' }
                                & $check ((Get-BoardLine $script:PB.Board $lines[0].Id).Student -eq $sp) 'undo put the student back'
                                Invoke-PBCommand @{ Cmd = 'MoveIP'; From = $lines[0].Id; LineId = $lines[1].Id; Copy = $false }
                                Invoke-PBCommand @{ Cmd = 'Undo' }
                                Invoke-PBCommand @{ Cmd = 'SetEtd'; LineId = $lines[1].Id; Value = $lines[1].ETD.AddMinutes(5) }
                                & $check ((Get-BoardLine $script:PB.Board $lines[1].Id).ETD -eq $lines[1].ETD.AddMinutes(5)) 'ETD changed'
                                Invoke-PBCommand @{ Cmd = 'SetType'; LineId = $lines[1].Id; Value = 'O&B' }
                                & $check (@($script:PB.Board.Lines).Count -eq 6) 'O&B added its return line'
                                Invoke-PBCommand @{ Cmd = 'Undo' }
                                & $check (@($script:PB.Board.Lines).Count -eq 5) 'undo removed the return line'
                                $choices = @(Get-PBMissionChoices (Get-BoardLine $script:PB.Board $lines[0].Id))
                                & $check ($choices.Count -ge 2) 'mission dropdown has events'
                                if ($choices.Count -ge 2) { Invoke-PBCommand @{ Cmd = 'SetMission'; LineId = $lines[0].Id; Value = $choices[1] } }
                                Invoke-PBCommand @{ Cmd = 'SetField'; LineId = $lines[0].Id; Field = 'Notes'; Value = 'self-test note' }
                                Invoke-PBDoEvents

                                $step = 'background checks'
                                $n = 0
                                while ($script:PB.Queue.Count -gt 0 -and $n -lt 400) { Invoke-PBAssessTick; $n++ }
                                & $check ($script:PB.Queue.Count -eq 0) 'background check queue drained'
                                $pinned = Get-BoardLine $script:PB.Board $lines[0].Id
                                $panel = New-PBAssessmentPanel (Get-PBLineAssessment $pinned)
                                & $check ($panel.Children.Count -gt 1) 'tooltip panel has lines'
                                $menu = New-Object System.Windows.Controls.ContextMenu
                                Set-PBMenuItems -Menu $menu -Tag 'StudentSlot' -Line $pinned
                                & $check ($menu.Items.Count -ge 3) 'right-click menu built'
                                & $check (@($script:PB.C.StudentPool.ItemsSource).Count -eq @($Data.Students).Count) 'student pool shows every student'
                                & $check (@($script:PB.C.IPPool.ItemsSource).Count -eq @($Data.Instructors).Count) 'IP pool shows every IP'

                                $step = 'heat'
                                Start-PBHeat @{ Kind = 'Student'; Name = [string]$Data.Students[1].Name; From = '' }
                                $n = 0
                                while ($null -ne $script:PB.Heat -and $script:PB.Heat.Queue.Count -gt 0 -and $n -lt 200) { Invoke-PBHeatTick; $n++ }
                                Stop-PBHeat

                                $step = 'fill and review'
                                Invoke-PBFill
                                Invoke-PBDoEvents
                                & $check ($script:PB.Mode -eq 'Review') 'board in review mode after fill'
                                & $check ($null -ne $script:PB.FillResult) 'fill result kept'
                                Exit-PBReview
                                Update-PBAll

                                $step = 'template and autosave'
                                $tplPath = Save-BoardTemplate -Board $script:PB.Board -Day $Day -Name 'SelfTest' -Dir $script:PB.Settings.TemplateDir
                                & $check (Test-Path -LiteralPath $tplPath) 'template saved'
                                Update-PBTemplateList
                                Save-PBBoard
                                & $check (Test-Path -LiteralPath (Join-Path $script:PB.Settings.BoardDir "$Day.json")) 'board autosaved'

                                $step = 'export'
                                Invoke-PBFill
                                Invoke-PBExport
                                & $check ($null -ne $script:PB.Result -and $script:PB.Result.Action -eq 'Export') 'export returned the fill'
                }
                catch {
                                $problems.Add("[$step] Exception: $($_.Exception.Message) $($_.ScriptStackTrace)")
                }
                finally {
                                try { if ($null -ne $script:PB -and $script:PB.Window.IsVisible) { $script:PB.Window.Close() } } catch { }
                                $listener.Flush()
                                $source.Listeners.Remove($listener)
                }

                foreach ($ln in ($writer.ToString() -split "`r?`n")) {
                                if ($ln -match 'BindingExpression|System\.Windows\.Data Error') { $problems.Add("Binding: $ln") }
                }
                return $problems.ToArray()
}
