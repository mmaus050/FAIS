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
#  - MasterData's Schedule sheet is re-read whenever MasterData is saved (another timer watches
#    the file). Board days nobody has edited since they were loaded update by themselves; for
#    the rest a banner offers the change. "Load a day..." loads one day from the sheet or from
#    the last exported schedule.
#  - Pucks go on a line by drag and drop, or by click and place: a click picks a puck up
#    (PB.Held), the next slot clicked gets it, and a double-click puts it on the selected line.
#  - Clicking a line selects it (PB.SelectedLineId). The Line details panel under the days shows
#    that line's student against every rule and ranks every student for the line (DetailsTimer).
#  - The menu bar has every action; the Line menu acts on the selected line.

$script:PB = $null
$script:PBDlg = $null
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

# -Data:     @{ Students; Instructors; SoftWeights; Formations; PrevSchedule; ScheduleSheet; MasterDataPath; OutputPath }
#            ScheduleSheet is a Read-ScheduleSheet result; PrevSchedule the last exported MasterSchedule.
# -Days:     'yyyy-MM-dd' days to open (saved future boards are added automatically)
# -Settings: BoardDir, TemplateDir, OBReturnOffsetMinutes, ReturnSortieType, SimInstructor, Trace
# Returns @{ Action = 'Export'; Board; FillResult; Days } or $null when closed without exporting.
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
                $reset = @(Reset-PlannedCompletions -Students $Data.Students -FromDate $firstDate -SoftWeights $Data.SoftWeights -PrevSchedule @($Data.PrevSchedule))
                Clear-AssessmentCache

                $xamlPath = Join-Path $PSScriptRoot 'PuckBoard.xaml'
                [xml]$xaml = Get-Content -LiteralPath $xamlPath -Raw -Encoding UTF8
                $reader = New-Object System.Xml.XmlNodeReader $xaml
                $window = [System.Windows.Markup.XamlReader]::Load($reader)

                $controls = @{}
                foreach ($name in @('UndoButton', 'HelpButton', 'QuickEntryBox', 'AddLineButton', 'QuickEntryFeedback', 'TemplateCombo',
                                'LoadTemplateButton', 'SaveTemplateButton', 'CopyPrevDayButton', 'LoadDayButton', 'AddDayPicker', 'AddDayButton',
                                'RemoveDayButton', 'BulkLinesButton', 'UpdateSyllabiButton', 'CancelButton', 'BackToEditButton', 'RefillButton', 'ExportButton',
                                'FillDayButton', 'FillAllButton', 'ColorKeyButton', 'StatusText',
                                'CountsText', 'IPSearchBox', 'IPPool', 'StudentSearchBox', 'StudentPool', 'ModeBanner', 'ModeBannerText', 'DayTabs',
                                'SheetBanner', 'SheetBannerText', 'SheetBannerLoadButton', 'SheetBannerIgnoreButton', 'LogoImage',
                                'HoldBanner', 'HoldBannerText', 'HoldLetGoButton', 'CenterGrid', 'DetailsSplitter', 'DetailsPanel', 'DetailsHideButton',
                                'DetailsTitle', 'DetailsChecks', 'DetailsPlaceButton', 'DetailsHideHardCheck', 'DetailsListTitle', 'DetailsListInfo', 'DetailsList',
                                'MainMenu', 'MenuLoadDay', 'MenuApplyTemplate', 'MenuSaveTemplate', 'MenuCopyPrevDay', 'MenuUpdateSyllabi', 'MenuClose',
                                'MenuUndo', 'MenuTypeLines', 'MenuBulkLines', 'MenuRecheck', 'MenuLetGo', 'MenuLine', 'MenuPrevDay', 'MenuNextDay',
                                'MenuAddNextDay', 'MenuAddPickedDay', 'MenuRemoveDay', 'MenuFillDay', 'MenuFillAll', 'MenuRefill', 'MenuBackToEdit',
                                'MenuExport', 'MenuShowDetails', 'MenuColorKey', 'MenuHelp', 'MenuHelpColorKey',
                                'BusyOverlay', 'BusyText', 'BusyProgress', 'DropPopup', 'DropPopupPanel')) {
                                $c = $window.FindName($name)
                                if ($null -eq $c) { throw "PuckBoard.xaml is missing the element '$name'" }
                                $controls[$name] = $c
                }
                Set-PBLogo -Image $controls.LogoImage -Path (Join-Path $PSScriptRoot 'BoardLogo.png')

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
                                Filling     = $false
                                HoverKey    = $null
                                HoverElement = $null
                                TypeOptions = @('Day Local', 'O&B', [string]$set.ReturnSortieType, 'OFT', 'IFT')
                                LastFocusDay = $firstDay
                                FillDays    = @()
                                SheetIgnore = (New-Object System.Collections.Generic.List[string])
                                SheetBannerInfo = $null
                                SheetStamp  = ''
                                SheetPending = ''
                                SheetAutoPending = $false
                                ImportDir   = ''
                                SelectedLineId = ''
                                Held        = $null
                                Press       = $null
                                Details     = $null
                                DetailsVisible = $true
                                DetailsHeight = 240
                                ColorKeyWindow = $null
                }

                Initialize-PBTimers
                Register-PBHandlers
                Register-PBMenuHandlers
                Initialize-PBDetails
                $controls.AddDayPicker.SelectedDate = (Get-NextFlyingDay -From ([datetime]::ParseExact($board.Days[$board.Days.Count - 1], 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)))
                Update-PBTemplateList
                Set-PBMode 'Edit'
                Update-PBAll
                # A Schedule sheet problem never stops the board from opening
                $sheetMsg = ''
                try { $sheetMsg = Invoke-PBSheetStartup }
                catch {
                                Write-PBLog ("Schedule sheet at start: {0}`n{1}" -f $_.Exception.Message, $_.ScriptStackTrace)
                                $sheetMsg = "Could not load MasterData's Schedule sheet ($($_.Exception.Message)). Use Load a day... to try again."
                }
                Start-PBAssessQueue
                if (-not [string]::IsNullOrWhiteSpace([string]$Data.MasterDataPath)) { $script:PB.SheetTimer.Start() }

                $msg = "Board opened for $(@($board.Days | ForEach-Object { Format-BoardDayLabel $_ }) -join ', ')."
                if ($reset.Count -gt 0) { $msg += " Cleared planned events on/after $firstDay for $($reset.Count) student(s) so they can be rescheduled." }
                if ($sheetMsg) { $msg += " $sheetMsg" }
                Set-PBStatus $msg
                return $window
}

# Boards / Templates folder next to FAIS.ps1, else %LOCALAPPDATA%\FAIS
# Shows the logo file in the title bar. A missing or unreadable file just hides the logo.
# Read through a stream (not a file URI) so a path with '#' or '%' in it still works and the file isn't locked.
function Set-PBLogo {
                param ($Image, [string]$Path)

                $Image.Visibility = 'Collapsed'
                if (-not (Test-Path -LiteralPath $Path)) { return }
                try {
                                $stream = [System.IO.File]::OpenRead($Path)
                                try {
                                                $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
                                                $bitmap.BeginInit()
                                                $bitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
                                                $bitmap.StreamSource = $stream
                                                $bitmap.EndInit()
                                                $bitmap.Freeze()
                                } finally { $stream.Dispose() }
                                $Image.Source = $bitmap
                                $Image.Visibility = 'Visible'
                }
                catch { }
}

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
                                # A menu item counts too: a click on a right-click menu must not pick up or place the slot under it
                                if ($el -is [System.Windows.Controls.Primitives.TextBoxBase] -or $el -is [System.Windows.Controls.ComboBox] -or
                                                $el -is [System.Windows.Controls.Primitives.ButtonBase] -or $el -is [System.Windows.Controls.Primitives.ScrollBar] -or
                                                $el -is [System.Windows.Controls.DatePicker] -or $el -is [System.Windows.Controls.MenuItem]) { return $true }
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
                $popup.Add_Tick({ Invoke-PBSafe { Close-PBPopup } })
                $script:PB.PopupTimer = $popup

                $heat = New-Object System.Windows.Threading.DispatcherTimer
                $heat.Interval = [TimeSpan]::FromMilliseconds(30)
                $heat.Add_Tick({ Invoke-PBSafe { Invoke-PBHeatTick } })
                $script:PB.HeatTimer = $heat

                # Watches MasterData so a save in Excel shows up on the board
                $sheet = New-Object System.Windows.Threading.DispatcherTimer
                $sheet.Interval = [TimeSpan]::FromSeconds(2)
                $sheet.Add_Tick({ Invoke-PBSafe { Invoke-PBSheetTick } })
                $script:PB.SheetTimer = $sheet

                # Ranks every student for the selected line, a few at a time
                $details = New-Object System.Windows.Threading.DispatcherTimer
                $details.Interval = [TimeSpan]::FromMilliseconds(30)
                $details.Add_Tick({ Invoke-PBSafe { Invoke-PBDetailsTick } })
                $script:PB.DetailsTimer = $details

                $w.Add_Closing({
                                param($s, $e)
                                Invoke-PBSafe {
                                                foreach ($t in @($script:PB.AssessTimer, $script:PB.SaveTimer, $script:PB.PopupTimer, $script:PB.HeatTimer, $script:PB.SheetTimer, $script:PB.DetailsTimer)) { $t.Stop() }
                                                Save-PBBoard
                                }
                })
}

function Register-PBHandlers {
                $w = $script:PB.Window
                $C = $script:PB.C

                $w.AddHandler([System.Windows.UIElement]::PreviewMouseLeftButtonDownEvent,
                                [System.Windows.Input.MouseButtonEventHandler]{ param($s, $e) Invoke-PBSafe { Invoke-PBMouseDown $e } }, $true)
                $w.AddHandler([System.Windows.UIElement]::PreviewMouseLeftButtonUpEvent,
                                [System.Windows.Input.MouseButtonEventHandler]{ param($s, $e) Invoke-PBSafe { Invoke-PBMouseUp $e } }, $true)
                # A right-click selects the line too, so the Line menu and its keys act on the line clicked
                $w.AddHandler([System.Windows.UIElement]::PreviewMouseRightButtonDownEvent,
                                [System.Windows.Input.MouseButtonEventHandler]{ param($s, $e) Invoke-PBSafe { Select-PBLineAt $e.OriginalSource } }, $true)
                $w.AddHandler([System.Windows.Input.Keyboard]::GotKeyboardFocusEvent,
                                [System.Windows.Input.KeyboardFocusChangedEventHandler]{ param($s, $e) Invoke-PBSafe { if (-not $script:PB.Busy) { Select-PBLineAt $e.NewFocus } } }, $true)
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
                $w.Add_Deactivated({ Invoke-PBSafe { Close-PBPopup } })
                $w.AddHandler([System.Windows.FrameworkElement]::ContextMenuOpeningEvent,
                                [System.Windows.Controls.ContextMenuEventHandler]{ param($s, $e) Invoke-PBSafe { Invoke-PBContextMenuOpening $e } }, $true)
                $w.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent,
                                [System.Windows.RoutedEventHandler]{ param($s, $e) Invoke-PBSafe { Invoke-PBRowClick $e } }, $true)

                $C.AddLineButton.Add_Click({ Invoke-PBSafe { Invoke-PBQuickEntry } })
                $C.BulkLinesButton.Add_Click({ Invoke-PBSafe { Invoke-PBBulkLines } })
                $C.UpdateSyllabiButton.Add_Click({ Invoke-PBSafe { Invoke-PBUpdateSyllabi } })
                $C.UndoButton.Add_Click({ Invoke-PBSafe { Request-PBCommand @{ Cmd = 'Undo' } } })
                $C.HelpButton.Add_Click({ Invoke-PBSafe { Show-PBHelp } })
                $C.ColorKeyButton.Add_Click({ Invoke-PBSafe { Show-PBColorKey } })
                $C.HoldLetGoButton.Add_Click({ Invoke-PBSafe { Request-PBCommand @{ Cmd = 'LetGo' } } })
                $C.LoadTemplateButton.Add_Click({ Invoke-PBSafe { Invoke-PBLoadTemplate } })
                $C.SaveTemplateButton.Add_Click({ Invoke-PBSafe { Invoke-PBSaveTemplate } })
                $C.CopyPrevDayButton.Add_Click({ Invoke-PBSafe { Request-PBCommand @{ Cmd = 'CopyPrevDay' } } })
                $C.LoadDayButton.Add_Click({ Invoke-PBSafe { Invoke-PBLoadDayDialog } })
                $C.SheetBannerLoadButton.Add_Click({ Invoke-PBSafe { Invoke-PBSheetBannerLoad } })
                $C.SheetBannerIgnoreButton.Add_Click({ Invoke-PBSafe { Invoke-PBSheetBannerIgnore } })
                $C.AddDayButton.Add_Click({ Invoke-PBSafe { Invoke-PBAddDay } })
                $C.RemoveDayButton.Add_Click({ Invoke-PBSafe { Invoke-PBRemoveDay } })
                $C.FillDayButton.Add_Click({ Invoke-PBSafe { Request-PBCommand @{ Cmd = 'Fill'; Scope = 'Day' } } })
                $C.FillAllButton.Add_Click({ Invoke-PBSafe { Request-PBCommand @{ Cmd = 'Fill'; Scope = 'All' } } })
                $C.RefillButton.Add_Click({ Invoke-PBSafe { Request-PBCommand @{ Cmd = 'Fill'; Refill = $true } } })
                $C.BackToEditButton.Add_Click({ Invoke-PBSafe { Invoke-PBBackToEdit } })
                $C.ExportButton.Add_Click({ Invoke-PBSafe { Invoke-PBExport } })
                $C.CancelButton.Add_Click({ Invoke-PBSafe { Invoke-PBCancel } })
                $C.IPSearchBox.Add_TextChanged({ Invoke-PBSafe { Update-PBPools } })
                $C.StudentSearchBox.Add_TextChanged({ Invoke-PBSafe { Update-PBPools } })

                # Line details panel
                $C.DetailsHideButton.Add_Click({ Invoke-PBSafe { Show-PBDetailsPanel -Visible $false } })
                $C.DetailsPlaceButton.Add_Click({ Invoke-PBSafe { Invoke-PBDetailsPlace } })
                $C.DetailsHideHardCheck.Add_Checked({ Invoke-PBSafe { Update-PBDetailsFilter } })
                $C.DetailsHideHardCheck.Add_Unchecked({ Invoke-PBSafe { Update-PBDetailsFilter } })
                $C.DetailsList.Add_SelectionChanged({ Invoke-PBSafe { Invoke-PBDetailsPick } })
                $C.DetailsList.Add_MouseDoubleClick({
                                param($s, $e)
                                Invoke-PBSafe { if ($null -ne (Find-PBTagged $e.OriginalSource @('Candidate'))) { Invoke-PBDetailsPlace } }
                })
}

# The menu bar. Each item does what its button, key or right-click item does.
function Register-PBMenuHandlers {
                $C = $script:PB.C

                $C.MenuLoadDay.Add_Click({ Invoke-PBSafe { Invoke-PBLoadDayDialog } })
                $C.MenuSaveTemplate.Add_Click({ Invoke-PBSafe { Invoke-PBSaveTemplate } })
                $C.MenuCopyPrevDay.Add_Click({ Invoke-PBSafe { Request-PBCommand @{ Cmd = 'CopyPrevDay' } } })
                $C.MenuUpdateSyllabi.Add_Click({ Invoke-PBSafe { Invoke-PBUpdateSyllabi } })
                $C.MenuClose.Add_Click({ Invoke-PBSafe { Invoke-PBCancel } })
                $C.MenuUndo.Add_Click({ Invoke-PBSafe { Request-PBCommand @{ Cmd = 'Undo' } } })
                $C.MenuTypeLines.Add_Click({ Invoke-PBSafe { Request-PBCommand @{ Cmd = 'FocusQuickEntry' } } })
                $C.MenuBulkLines.Add_Click({ Invoke-PBSafe { Invoke-PBBulkLines } })
                $C.MenuRecheck.Add_Click({ Invoke-PBSafe { Invoke-PBRecheck } })
                $C.MenuLetGo.Add_Click({ Invoke-PBSafe { Request-PBCommand @{ Cmd = 'LetGo' } } })
                $C.MenuPrevDay.Add_Click({ Invoke-PBSafe { Select-PBDayOffset -1 } })
                $C.MenuNextDay.Add_Click({ Invoke-PBSafe { Select-PBDayOffset 1 } })
                $C.MenuAddNextDay.Add_Click({ Invoke-PBSafe { Invoke-PBAddDay -AddDate (Get-PBNextBoardDay) } })
                $C.MenuAddPickedDay.Add_Click({ Invoke-PBSafe { Invoke-PBAddDay } })
                $C.MenuRemoveDay.Add_Click({ Invoke-PBSafe { Invoke-PBRemoveDay } })
                $C.MenuFillDay.Add_Click({ Invoke-PBSafe { Request-PBCommand @{ Cmd = 'Fill'; Scope = 'Day' } } })
                $C.MenuFillAll.Add_Click({ Invoke-PBSafe { Request-PBCommand @{ Cmd = 'Fill'; Scope = 'All' } } })
                $C.MenuRefill.Add_Click({ Invoke-PBSafe { Request-PBCommand @{ Cmd = 'Fill'; Refill = $true } } })
                $C.MenuBackToEdit.Add_Click({ Invoke-PBSafe { Invoke-PBBackToEdit } })
                $C.MenuExport.Add_Click({ Invoke-PBSafe { Invoke-PBExport } })
                $C.MenuShowDetails.Add_Checked({ Invoke-PBSafe { Show-PBDetailsPanel -Visible $true } })
                $C.MenuShowDetails.Add_Unchecked({ Invoke-PBSafe { Show-PBDetailsPanel -Visible $false } })
                $C.MenuColorKey.Add_Click({ Invoke-PBSafe { Show-PBColorKey } })
                $C.MenuHelpColorKey.Add_Click({ Invoke-PBSafe { Show-PBColorKey } })
                $C.MenuHelp.Add_Click({ Invoke-PBSafe { Show-PBHelp } })

                # SubmenuOpened bubbles up from nested submenus too ("Fly wing to..."): only rebuild
                # a menu when it is the one opening, never while one of its own submenus opens
                $C.MenuLine.Add_SubmenuOpened({
                                param($s, $e)
                                Invoke-PBSafe { if ([object]::ReferenceEquals($e.OriginalSource, $script:PB.C.MenuLine)) { Update-PBLineMenu } }
                })
                $C.MenuApplyTemplate.Add_SubmenuOpened({
                                param($s, $e)
                                Invoke-PBSafe { if ([object]::ReferenceEquals($e.OriginalSource, $script:PB.C.MenuApplyTemplate)) { Update-PBTemplateMenu } }
                })
                $C.MainMenu.AddHandler([System.Windows.Controls.MenuItem]::SubmenuOpenedEvent, [System.Windows.RoutedEventHandler]{
                                param($s, $e)
                                Invoke-PBSafe { if ($script:PB.C.MainMenu.Items.Contains($e.OriginalSource)) { Update-PBMenuState } }
                })
}

# Board changes asked for from inside an element's event run right after that event
function Request-PBCommand {
                param ([hashtable]$Command)

                if ($script:PB.Filling) { return }

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
                $focus = Get-PBFocusState
                $script:PB.Busy = $true
                try { & $Action } finally { $script:PB.Busy = $false }
                Update-PBAll
                Restore-PBFocus $focus
                Request-PBAutosave
                Start-PBAssessQueue
                if ($Status) { Set-PBStatus $Status }
}

# Which row cell has the keyboard, so it can be given back after the rows are rebuilt
function Get-PBFocusState {
                $fe = [System.Windows.Input.Keyboard]::FocusedElement -as [System.Windows.FrameworkElement]
                if ($null -eq $fe -or $fe.Tag -isnot [string]) { return $null }
                if (@('EtdBox', 'CallsignBox', 'NumBox', 'NotesBox', 'TypeBox', 'MissionBox') -notcontains $fe.Tag) { return $null }
                $vm = $fe.DataContext
                if ($null -eq $vm -or $null -eq $vm.PSObject.Properties['LineId']) { return $null }
                $state = @{ LineId = [string]$vm.LineId; Tag = [string]$fe.Tag; Start = 0; Length = 0 }
                if ($fe -is [System.Windows.Controls.TextBox]) { $state.Start = $fe.SelectionStart; $state.Length = $fe.SelectionLength }
                return $state
}

function Restore-PBFocus {
                param ($State)

                if ($null -eq $State) { return }
                $tab = $script:PB.C.DayTabs.SelectedItem
                if ($null -eq $tab) { return }
                $script:PB.Window.UpdateLayout()
                foreach ($el in @(Get-PBTaggedDescendants -Root $tab.Content -Tag $State.Tag)) {
                                if ($null -ne $el.DataContext -and [string]$el.DataContext.LineId -eq $State.LineId) {
                                                [void]$el.Focus()
                                                if ($el -is [System.Windows.Controls.TextBox]) {
                                                                $len = ([string]$el.Text).Length
                                                                $el.Select([math]::Min($State.Start, $len), [math]::Min($State.Length, [math]::Max(0, $len - $State.Start)))
                                                }
                                                return
                                }
                }
}

function Update-PBAll {
                $script:PB.Busy = $true
                try {
                                Update-PBTabs
                                Sync-PBSelection -NoRedraw
                                foreach ($day in $script:PB.Board.Days) { Update-PBDay $day }
                                Update-PBPools
                                Update-PBCounts
                                Update-PBFillScope
                                Update-PBSheetBanner
                                Update-PBDetails
                                Update-PBHeldHeat
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
                # An open drop-down list has focus on one of its items; the owner is the combo box
                if ($focused -is [System.Windows.Controls.ComboBoxItem]) { $focused = [System.Windows.Controls.ItemsControl]::ItemsControlFromItemContainer($focused) }
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

# Missions offered in a pinned line's dropdown: (auto), the events the student can fly here,
# the next ones they aren't opted for yet (tagged), and "Type a mission...". Both lists come
# from the line's assessment, which sees the student's other lines on the board, so the tags
# agree with the flags. Until it is in: the events they are opted for in MasterData, no tags.
function Get-PBMissionChoices {
                param ($Line)

                $a = Get-PBCurrentAssessment $Line
                $names = @()
                $notOpted = @()
                if ($null -ne $a -and -not $a.Quick) {
                                $names = @($a.EligibleEvents)
                                $notOpted = @($a.NotOptedEvents)
                }
                else {
                                $student = Resolve-StudentByName -Students $script:PB.Data.Students -Name $Line.Student
                                $types = Get-AllowedEventTypes ([string]$Line.SortieType)
                                if ($null -ne $student -and $null -ne $types) {
                                                $names = @(Get-OptedEvents -SyllabusData $student.Syllabus | Where-Object { ([string]$_.Type).Trim() -in $types } |
                                                                Sort-Object -Property { Get-SyllabusSortKey $_ } | ForEach-Object { [string]$_.Mission })
                                }
                }
                return @(Get-MissionChoiceTexts -Eligible $names -Current ([string]$Line.Mission) -NotOpted $notOpted)
}

# What picking -Value in a row's mission dropdown asks for: a ChooseMission command, or $null
# when nothing changes ((auto) on an (auto) line, the line's own mission again, a disabled row)
function Get-PBMissionCommand {
                param ($VM, [string]$Value)

                if ($null -eq $VM -or -not $VM.MissionEnabled -or @($VM.MissionOptions) -notcontains $Value) { return $null }
                $choice = ConvertFrom-MissionChoiceText $Value
                if ($choice.Kind -ne 'Type' -and $choice.Mission -ceq [string]$VM.MissionRaw) { return $null }
                return @{ Cmd = 'ChooseMission'; LineId = [string]$VM.LineId; Value = $Value }
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
                                $missionOptions = @(Get-PBMissionChoices $Line)
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
                $csTip = 'Blank = the IP''s callsign (formations get FormCallsigns)'
                if ($review -and $null -ne $ov) {
                                $callsign = $ov.Callsign; $num = $ov.Num; $notes = $ov.Notes; $readOnly = $true
                                if ($ov.Override) { $notesFg = '#B91C1C' }
                }
                $rowColors = Get-PBRowColors -Line $Line -Selected ([string]$Line.Id -eq [string]$script:PB.SelectedLineId)

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
                                RowBackground          = $rowColors.Background
                                RowBorder              = $rowColors.Border
                                IsReturn               = $isReturn
                }
}

# A row's fill and outline: pink for an OVERRIDE in review, a blue outline (and tint) when selected
function Get-PBRowColors {
                param ($Line, [bool]$Selected)

                $bg = '#FFFFFF'
                $border = '#E2E8F0'
                if ($script:PB.Mode -eq 'Review' -and $null -ne $script:PB.Overlay) {
                                $ov = $script:PB.Overlay[$Line.Id]
                                if ($null -ne $ov -and $ov.Override) { $bg = '#FEF2F2' }
                }
                if ($Selected) {
                                $border = '#2563EB'
                                if ($bg -eq '#FFFFFF') { $bg = '#EFF6FF' }
                }
                return @{ Background = $bg; Border = $border }
}

function Update-PBPools {
                $day = Get-PBActiveDay
                $board = $script:PB.Board
                $C = $script:PB.C

                # A puck picked up with a click gets a thick dark outline
                $held = $script:PB.Held
                $clickTip = "`nClick to pick up, then click a slot. Double-click puts it on the selected line."
                $ipFilter = ([string]$C.IPSearchBox.Text).Trim()
                $ipVMs = foreach ($ip in @($script:PB.Data.Instructors | Sort-Object -Property Name)) {
                                if ($ipFilter -and ([string]$ip.Name -notlike "*$ipFilter*") -and ([string]$ip.Callsign -notlike "*$ipFilter*")) { continue }
                                $info = Get-InstructorPuckInfo -Instructor $ip -Board $board -Day $day
                                $isHeld = ($null -ne $held -and $held.Kind -eq 'IP' -and [string]$held.Name -eq [string]$info.Name)
                                [PSCustomObject]@{
                                                Name            = $info.Name
                                                SubText         = ("{0}  {1}" -f $info.Callsign, ($info.Badges -join ' ')).Trim()
                                                CountText       = "x$($info.Lines)"
                                                CountVisibility = $(if ($info.Lines -gt 0) { 'Visible' } else { 'Collapsed' })
                                                ToolTipText     = $(if ($info.Lines -gt 0) { "On $($info.Lines) line(s) this day: $($info.Times -join ', ')" } else { 'Not on any line this day' }) + $clickTip
                                                BorderBrush     = $(if ($isHeld) { '#1E3A8A' } else { '#93C5FD' })
                                                PuckBorderThickness = $(if ($isHeld) { '3' } else { '1' })
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
                                $tip += $clickTip.Trim()
                                $isHeld = ($null -ne $held -and $held.Kind -eq 'Student' -and [string]$held.Name -eq [string]$info.Name)
                                if ($isHeld) { $bd = '#1D4ED8' }
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
                                                PuckBorderThickness = $(if ($isHeld) { '3' } else { '1' })
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
                                $text += "   ||   Fill ({0}): {1} auto, {2} wing, {3} return, {4} DS, {5} for sale, {6} override(s)" -f (Format-PBDays @($script:PB.FillDays)), $st.Filled, $st.Wing, $st.Return, $st.DS, $st.ForSale, $st.Overrides
                }
                $script:PB.C.CountsText.Text = $text
}

function Set-PBMode {
                param ([ValidateSet('Edit', 'Review')] [string]$Mode)

                $C = $script:PB.C
                $script:PB.Mode = $Mode
                $review = ($Mode -eq 'Review')
                Update-PBFillScope
                foreach ($b in @($C.BackToEditButton, $C.RefillButton, $C.ExportButton)) { $b.Visibility = if ($review) { 'Visible' } else { 'Collapsed' } }
                $C.ModeBanner.Visibility = if ($review) { 'Visible' } else { 'Collapsed' }
                if ($review) {
                                $scope = Format-PBDays @($script:PB.FillDays)
                                $C.ModeBannerText.Text = "REVIEW ($scope): italic blue = filled by the scheduler, green = hardcoded, red notes = OVERRIDE (hardcoded against a rule), red = For Sale. Export writes $scope to OutputSchedule.xlsx; days already exported stay as they are. Any change goes back to editing."
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
                if ($script:PB.Busy -or $script:PB.Draining -or $script:PB.Filling) { return }
                if ($null -ne $script:PB.Heat -or ($null -ne $script:PB.Drag -and $script:PB.Drag.Active)) { return }
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
                $script:PB.Press = $null
                if ($script:PB.Filling) { $e.Handled = $true; return }
                # The pop-up never blocks the click that dismisses it
                if ($script:PB.C.DropPopup.IsOpen) { Close-PBPopup }
                $src = $e.OriginalSource
                # A click anywhere on a line (its boxes included) selects it for the Line details panel
                Select-PBLineAt $src
                if (Test-PBInsideInput $src) { return }
                $el = Find-PBTagged $src
                if ($null -eq $el) { return }
                # A press and release on the same puck or slot without dragging is a click (Invoke-PBMouseUp)
                if (@('IPPuck', 'StudentPuck', 'IPSlot', 'StudentSlot') -contains [string]$el.Tag) {
                                $script:PB.Press = @{ Element = $el; Count = $e.ClickCount }
                }
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
                if ($null -eq $d) { Update-PBHoverToolTip $e; return }
                if ($d.Active) { return }
                if ($e.LeftButton -ne [System.Windows.Input.MouseButtonState]::Pressed) { $script:PB.Drag = $null; return }
                $p = $e.GetPosition($script:PB.Window)
                if ([math]::Abs($p.X - $d.Start.X) -lt [System.Windows.SystemParameters]::MinimumHorizontalDragDistance -and
                                [math]::Abs($p.Y - $d.Start.Y) -lt [System.Windows.SystemParameters]::MinimumVerticalDragDistance) { return }

                $d.Active = $true
                $script:PB.Press = $null
                $script:PB.PendingDrop = $null
                # Dragging lets go of a puck picked up with a click (after the drag, so nothing moves under it)
                $letGo = ($null -ne $script:PB.Held)
                if ($letGo) { Stop-PBHeat }
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
                if ($letGo) { Request-PBCommand @{ Cmd = 'LetGo'; Quiet = $true } }
                if ($null -ne $script:PB.PendingDrop) {
                                $drop = $script:PB.PendingDrop
                                $script:PB.PendingDrop = $null
                                Request-PBCommand $drop
                }
}

# A click (press and release on the same puck or slot, no drag). On a puck: pick it up, or put
# it down again. Double-click on a puck: put it on the selected line. On a slot while holding a
# puck: put the puck there (Ctrl keeps holding it, for the next slot).
function Invoke-PBMouseUp {
                param ($e)

                $press = $script:PB.Press
                $script:PB.Press = $null
                if ($null -eq $press -or $script:PB.Filling) { return }
                $el = Find-PBTagged $e.OriginalSource @('IPPuck', 'StudentPuck', 'IPSlot', 'StudentSlot')
                if ($null -eq $el -or -not [object]::ReferenceEquals($el, $press.Element)) { return }
                $vm = $el.DataContext
                if ($null -eq $vm) { return }
                $tag = [string]$el.Tag
                if ($tag -eq 'IPPuck' -or $tag -eq 'StudentPuck') {
                                $kind = if ($tag -eq 'IPPuck') { 'IP' } else { 'Student' }
                                if ($press.Count -ge 2) { Request-PBCommand @{ Cmd = 'PlaceOnSelected'; Kind = $kind; Name = [string]$vm.Name } }
                                else { Request-PBCommand @{ Cmd = 'Hold'; Kind = $kind; Name = [string]$vm.Name } }
                                return
                }
                $keep = ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Control) -eq [System.Windows.Input.ModifierKeys]::Control
                [void](Request-PBPlaceOnSlot -Slot $el -Keep:$keep)
}

# Puts the puck being held on a slot (a click, or Enter on a focused slot). Says so when the
# slot is the wrong kind for the puck.
function Request-PBPlaceOnSlot {
                param ($Slot, [switch]$Keep)

                $held = $script:PB.Held
                if ($null -eq $held -or $null -eq $Slot -or $null -eq $Slot.DataContext) { return $false }
                $want = if ([string]$Slot.Tag -eq 'IPSlot') { 'IP' } else { 'Student' }
                if ($held.Kind -ne $want) {
                                $hint = if ($held.Kind -eq 'IP') { "$($held.Name) is an IP: click a line's IP slot (left of the student)." } else { "$($held.Name) is a student: click a line's student slot." }
                                Set-PBStatus $hint 'Warn'
                                return $false
                }
                Request-PBCommand @{ Cmd = 'PlaceHeld'; LineId = [string]$Slot.DataContext.LineId; Keep = [bool]$Keep }
                return $true
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
                if ($script:PB.Filling) { $e.Handled = $true; return }
                if ($key -eq [System.Windows.Input.Key]::Escape -and $C.DropPopup.IsOpen) { Close-PBPopup }
                $mods = [System.Windows.Input.Keyboard]::Modifiers
                $ctrl = ($mods -band [System.Windows.Input.ModifierKeys]::Control) -eq [System.Windows.Input.ModifierKeys]::Control
                $shift = ($mods -band [System.Windows.Input.ModifierKeys]::Shift) -eq [System.Windows.Input.ModifierKeys]::Shift
                $KeyT = [System.Windows.Input.Key]
                # An open menu handles its own keys
                if ($src -is [System.Windows.Controls.MenuItem]) { return }

                # Ctrl+PgUp / Ctrl+PgDn change day from anywhere, a box being typed in included
                if ($ctrl -and ($key -eq $KeyT::PageUp -or $key -eq $KeyT::PageDown)) {
                                Select-PBDayOffset $(if ($key -eq $KeyT::PageUp) { -1 } else { 1 })
                                $e.Handled = $true
                                return
                }

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
                if ($key -eq $KeyT::Escape -and $null -ne $script:PB.Held) { Request-PBCommand @{ Cmd = 'LetGo' }; $e.Handled = $true; return }
                if ($src -is [System.Windows.Controls.Primitives.TextBoxBase]) { return }

                # Arrow keys on a closed mission dropdown open it rather than step through it: moving in the
                # open list picks nothing until Enter or a click (a pick can ask about a rule break)
                $combo = $src -as [System.Windows.Controls.ComboBox]
                if ($null -ne $combo -and [string]$combo.Tag -eq 'MissionBox' -and -not $combo.IsDropDownOpen -and
                                @($KeyT::Up, $KeyT::Down, $KeyT::Left, $KeyT::Right, $KeyT::Home, $KeyT::End, $KeyT::PageUp, $KeyT::PageDown) -contains $key) {
                                $combo.IsDropDownOpen = $true
                                $e.Handled = $true
                                return
                }

                $selected = [string]$script:PB.SelectedLineId
                if ($ctrl -and $key -eq $KeyT::Z) { Request-PBCommand @{ Cmd = 'Undo' }; $e.Handled = $true; return }
                if ($ctrl -and $key -eq $KeyT::L) { [void]$C.QuickEntryBox.Focus(); $C.QuickEntryBox.SelectAll(); $e.Handled = $true; return }
                # Ctrl+Enter fills every day, Ctrl+Shift+Enter only the day on screen
                if ($ctrl -and $key -eq $KeyT::Enter) { Request-PBCommand @{ Cmd = 'Fill'; Scope = $(if ($shift) { 'Day' } else { 'All' }) }; $e.Handled = $true; return }
                if ($ctrl -and $key -eq $KeyT::D -and $selected) { Request-PBCommand @{ Cmd = 'DuplicateLine'; LineId = $selected }; $e.Handled = $true; return }
                if ($ctrl -and $key -eq $KeyT::Delete -and $selected) { Request-PBCommand @{ Cmd = 'DeleteLine'; LineId = $selected }; $e.Handled = $true; return }
                if ($key -eq $KeyT::F5) { Invoke-PBRecheck; $e.Handled = $true; return }
                if ($key -eq $KeyT::Enter -and $null -ne $script:PB.Held) {
                                $slot = Find-PBTagged ([System.Windows.Input.Keyboard]::FocusedElement) @('IPSlot', 'StudentSlot')
                                if ($null -ne $slot) { [void](Request-PBPlaceOnSlot -Slot $slot -Keep:$ctrl); $e.Handled = $true; return }
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
                if ($null -eq $vm -or $null -eq $vm.PSObject.Properties['LineId']) { return }
                # In Review mode callsign / num / notes show the fill's result and are read-only; an ETD edit
                # is allowed and sends the board back to editing
                if ($vm.ReadOnlyFields -and [string]$TextBox.Tag -ne 'EtdBox') { return }
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
                                Close-PBPopup
                                $script:PB.LastFocusDay = Get-PBActiveDay
                                Sync-PBSelection
                                Update-PBPools
                                Update-PBCounts
                                Update-PBFillScope
                                Update-PBSheetBanner
                                Update-PBDetails
                                Update-PBHeldHeat
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
                                                # Checked (and typed, for "Type a mission...") after this event, not inside it
                                                $cmd = Get-PBMissionCommand -VM $vm -Value $value
                                                if ($null -ne $cmd) { Request-PBCommand $cmd }
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

# ToolTipOpening is a Direct routed event (it never bubbles to the window), so the tooltip is
# built while the pointer moves over a slot, before WPF decides to show it
function Update-PBHoverToolTip {
                param ($e)

                $el = Find-PBTagged $e.OriginalSource @('StudentSlot')
                if ($null -eq $el) { $script:PB.HoverKey = $null; return }
                $vm = $el.DataContext
                if ($null -eq $vm -or $null -eq $vm.PSObject.Properties['LineId']) { return }
                $key = "{0}|{1}|{2}" -f $vm.LineId, $script:PB.Board.Version, $script:PB.Mode
                if ($script:PB.HoverKey -eq $key -and [object]::ReferenceEquals($script:PB.HoverElement, $el)) { return }
                $script:PB.HoverKey = $key
                $script:PB.HoverElement = $el
                Set-PBSlotToolTip $el
}

# Sets (or clears) a student slot's tooltip from the line it shows
function Set-PBSlotToolTip {
                param ($Element)

                $vm = $Element.DataContext
                $line = if ($null -ne $vm) { Get-BoardLine $script:PB.Board $vm.LineId } else { $null }
                if ($null -eq $line) { $Element.ToolTip = $null; return }

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
                $Element.ToolTip = $content
}

function Close-PBPopup {
                $script:PB.PopupTimer.Stop()
                $script:PB.C.DropPopup.IsOpen = $false
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
                param ([string]$Header, [hashtable]$Command, [string]$Gesture = '')

                $mi = New-Object System.Windows.Controls.MenuItem
                $mi.Header = $Header
                if ($Gesture) { $mi.InputGestureText = $Gesture }
                if ($null -ne $Command) {
                                $mi.Tag = $Command
                                $mi.Add_Click({ param($s, $a) Invoke-PBSafe { Request-PBCommand $s.Tag } })
                }
                return $mi
}

function Invoke-PBContextMenuOpening {
                param ($e)

                $el = Find-PBTagged $e.OriginalSource @('StudentSlot', 'IPSlot', 'Row')
                if ($null -eq $el) { return }
                $tag = [string]$el.Tag
                $menu = $el.ContextMenu
                if ($null -eq $menu) { return }
                $vm = $el.DataContext
                $line = if ($null -ne $vm) { Get-BoardLine $script:PB.Board $vm.LineId } else { $null }
                if ($null -eq $line) { $e.Handled = $true; return }
                Set-PBMenuItems -Menu $menu -Tag $tag -Line $line
}

# Fills a slot's / row's right-click menu
# -Tag Line is the menu bar's Line menu: everything for the student and the IP, with the keys
# that act on the selected line
function Set-PBMenuItems {
                param ($Menu, [string]$Tag, $Line)

                $Menu.Items.Clear()
                $id = $Line.Id
                $review = ($script:PB.Mode -eq 'Review')
                $canFormate = Test-IsWingCapableType ([string]$Line.SortieType)
                $isReturn = Test-IsReturnSortieType ([string]$Line.SortieType)
                $forStudent = ($Tag -eq 'StudentSlot' -or $Tag -eq 'Line')
                $forIP = ($Tag -eq 'IPSlot' -or $Tag -eq 'Line')

                [void]$Menu.Items.Add((New-PBMenuItem 'Show in Line details' @{ Cmd = 'ShowDetails'; LineId = $id }))
                # The puck picked up with a click can go here too
                $held = $script:PB.Held
                if ($null -ne $held -and -not $isReturn -and (($held.Kind -eq 'Student' -and ($forStudent -or $Tag -eq 'Row')) -or ($held.Kind -eq 'IP' -and ($forIP -or $Tag -eq 'Row')))) {
                                [void]$Menu.Items.Add((New-PBMenuItem "Put $($held.Name) on this line" @{ Cmd = 'PlaceHeld'; LineId = $id }))
                }
                [void]$Menu.Items.Add((New-Object System.Windows.Controls.Separator))

                if ($forStudent) {
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
                if ($forIP -and $Line.Instructor) {
                                [void]$Menu.Items.Add((New-PBMenuItem "Take $($Line.Instructor) off this line" @{ Cmd = 'RemoveIP'; LineId = $id }))
                }
                if ($Menu.Items[$Menu.Items.Count - 1] -isnot [System.Windows.Controls.Separator]) { [void]$Menu.Items.Add((New-Object System.Windows.Controls.Separator)) }
                $keys = ($Tag -eq 'Line')
                [void]$Menu.Items.Add((New-PBMenuItem 'Duplicate line' @{ Cmd = 'DuplicateLine'; LineId = $id } -Gesture $(if ($keys) { 'Ctrl+D' } else { '' })))
                if ($Line.LinkedLineId) { [void]$Menu.Items.Add((New-PBMenuItem 'Unlink O&B / return' @{ Cmd = 'Unlink'; LineId = $id })) }
                [void]$Menu.Items.Add((New-PBMenuItem 'Delete line' @{ Cmd = 'DeleteLine'; LineId = $id } -Gesture $(if ($keys) { 'Ctrl+Delete' } else { '' })))
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
                                'ChooseMission' { Invoke-PBChooseMission -LineId $id -Choice ([string]$Cmd.Value) }
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
                                                $links = if ($Cmd.ContainsKey('Links')) { @($Cmd.Links) } else { @() }
                                                Invoke-PBChange { Add-BoardLine -Board $board -Line $lines -Link $links } -Status $Cmd.Status
                                                # A student typed with a mission (sp: m:): show the rule check, like a pick in the dropdown
                                                $typed = @($lines | Where-Object { $null -ne $_ -and $_.Student -and $_.Mission } | Select-Object -First 1)
                                                if ($typed.Count -gt 0) { Show-PBPlacementPopup ([string]$typed[0].Id) }
                                }
                                'Undo' {
                                                $wasReview = ($script:PB.Mode -eq 'Review')
                                                if ($wasReview) { Exit-PBReview }
                                                if (Undo-Board $board) {
                                                                Update-PBAll; Request-PBAutosave; Start-PBAssessQueue
                                                                Set-PBStatus 'Undone.'
                                                }
                                                else {
                                                                if ($wasReview) { Update-PBAll }
                                                                Set-PBStatus 'Nothing to undo.' 'Warn'
                                                }
                                }
                                'CopyPrevDay' { Invoke-PBCopyPrevDay }
                                'Fill' { Invoke-PBFill -Refill:([bool]$Cmd.Refill) -Scope $(if ($Cmd.Scope -eq 'Day') { 'Day' } else { 'All' }) }
                                # Click and place
                                'Hold' { Set-PBHeld -Kind $Cmd.Kind -Name $Cmd.Name }
                                'LetGo' { Clear-PBHeld -Quiet:([bool]$Cmd.Quiet) }
                                'PlaceHeld' {
                                                $held = $script:PB.Held
                                                if ($null -ne $held) { Invoke-PBPlace -Kind $held.Kind -Name $held.Name -LineId $id -KeepHolding:([bool]$Cmd.Keep) }
                                }
                                'PlaceOnSelected' {
                                                $sel = [string]$script:PB.SelectedLineId
                                                if (-not $sel -or $null -eq (Get-BoardLine $board $sel)) {
                                                                Set-PBStatus "Click a line first, then double-click $($Cmd.Name) to put them on it (or click a slot while holding $($Cmd.Name))." 'Warn'
                                                }
                                                else { Invoke-PBPlace -Kind $Cmd.Kind -Name $Cmd.Name -LineId $sel }
                                }
                                'PlaceOnLine' { Invoke-PBPlace -Kind $Cmd.Kind -Name $Cmd.Name -LineId $id }
                                'ShowDetails' {
                                                Select-PBLine $id
                                                Show-PBDetailsPanel -Visible $true
                                }
                                'FocusQuickEntry' { [void]$script:PB.C.QuickEntryBox.Focus(); $script:PB.C.QuickEntryBox.SelectAll() }
                }
}


# =================================================================================
#  CLICK AND PLACE
#  A click on a puck picks it up (PB.Held): its outline goes thick, the banner says so and, for
#  a student, the slots turn green / amber / red like a drag. The next slot clicked gets it.
#  A double-click on a puck puts it straight on the selected line. Esc lets go.
# =================================================================================

# Picks a puck up, or lets go of it when it is the one already held
function Set-PBHeld {
                param ([string]$Kind, [string]$Name)

                $held = $script:PB.Held
                if ($null -ne $held -and $held.Kind -eq $Kind -and $held.Name -eq $Name) { Clear-PBHeld; return }
                Stop-PBHeat
                $script:PB.Held = @{ Kind = $Kind; Name = $Name }
                Update-PBPools
                Update-PBHoldBanner
                Update-PBHeldHeat
                $where = if ($Kind -eq 'Student') { 'a student slot' } else { "a line's IP slot" }
                Set-PBStatus "Holding $Name. Click $where to put them there, or press Esc to let go."
}

function Clear-PBHeld {
                param ([switch]$Quiet)

                $held = $script:PB.Held
                if ($null -eq $held) { return }
                $script:PB.Held = $null
                Stop-PBHeat
                Update-PBPools
                Update-PBHoldBanner
                if (-not $Quiet) { Set-PBStatus "Let go of $($held.Name)." }
}

function Update-PBHoldBanner {
                $C = $script:PB.C
                $held = $script:PB.Held
                if ($null -eq $held) { $C.HoldBanner.Visibility = 'Collapsed'; return }
                $C.HoldBannerText.Text = if ($held.Kind -eq 'Student') {
                                "Holding $($held.Name): click a student slot to put them there (green fits, amber needs a look, red breaks a rule). Hold Ctrl as you click to keep holding them for another line."
                }
                else {
                                "Holding $($held.Name): click a line's IP slot to put them there. Hold Ctrl as you click to keep holding them for another line."
                }
                $C.HoldBanner.Visibility = 'Visible'
}

# Colors the slots for a student being held (again after the rows are redrawn)
function Update-PBHeldHeat {
                $held = $script:PB.Held
                if ($null -eq $held -or $held.Kind -ne 'Student') { return }
                if ($null -ne $script:PB.Drag -and $script:PB.Drag.Active) { return }
                Stop-PBHeat
                $script:PB.Window.UpdateLayout()
                Start-PBHeat @{ Name = [string]$held.Name; From = '' }
}

# Puts an IP or a student on a line: click and place, a double-click, or the Line details list.
# The puck held is let go of unless -KeepHolding.
function Invoke-PBPlace {
                param ([string]$Kind, [string]$Name, [string]$LineId, [switch]$KeepHolding)

                $line = Get-BoardLine $script:PB.Board $LineId
                if ($null -eq $line -or [string]::IsNullOrWhiteSpace($Name)) { return }
                if (Test-IsReturnSortieType ([string]$line.SortieType)) {
                                Set-PBStatus "$(Format-BoardTime $line.ETD) $($line.SortieType) follows its O&B: put $Name on the O&B line instead." 'Warn'
                                return
                }
                if (-not $KeepHolding -and $null -ne $script:PB.Held) {
                                # Let go first, so the redraw that follows the change shows the puck put down
                                $script:PB.Held = $null
                                Stop-PBHeat
                                Update-PBHoldBanner
                }
                if ($Kind -eq 'IP') { Invoke-PBCommand @{ Cmd = 'SetIP'; LineId = $LineId; Name = $Name } }
                else {
                                if ($null -ne $script:PB.Details) { $script:PB.Details.Focus = '' }
                                Invoke-PBCommand @{ Cmd = 'PinStudent'; LineId = $LineId; Name = $Name }
                }
}


# =================================================================================
#  THE SELECTED LINE AND THE LINE DETAILS PANEL
#  Clicking a line (or right-clicking it, or tabbing into it) selects it: a blue outline, and
#  the panel under the days shows
#    - left: the line's student (or the one picked in the list) against the hard rules, the
#      warnings and the soft-rule score: the same check as the pop-up after a drop;
#    - right: every student ranked for the line, best first, checked a few at a time
#      (DetailsTimer). Double-click one, or Put on this line, to hardcode them there.
# =================================================================================

# Selects the line an element belongs to
function Select-PBLineAt {
                param ($Element)

                if ($null -eq $Element) { return }
                $row = Find-PBTagged $Element @('Row')
                if ($null -eq $row -or $null -eq $row.DataContext -or $null -eq $row.DataContext.PSObject.Properties['LineId']) { return }
                Select-PBLine ([string]$row.DataContext.LineId)
}

function Select-PBLine {
                param ([string]$LineId)

                $old = [string]$script:PB.SelectedLineId
                if ($old -eq $LineId) { return }
                $script:PB.SelectedLineId = $LineId
                foreach ($id in @($old, $LineId)) { if ($id) { Set-PBRowHighlight -LineId $id } }
                Update-PBDetails
}

# Recolors one row for the selection without rebuilding it (a rebuild inside the click that
# selected it would swallow the click)
function Set-PBRowHighlight {
                param ([string]$LineId)

                $line = Get-BoardLine $script:PB.Board $LineId
                if ($null -eq $line) { return }
                $el = Get-PBSlotElement -LineId $LineId -Tag 'Row'
                if ($null -eq $el) { return }
                $colors = Get-PBRowColors -Line $line -Selected ($LineId -eq [string]$script:PB.SelectedLineId)
                $el.SetCurrentValue([System.Windows.Controls.Border]::BackgroundProperty, (Get-PBBrush $colors.Background))
                $el.SetCurrentValue([System.Windows.Controls.Border]::BorderBrushProperty, (Get-PBBrush $colors.Border))
}

# Drops the selection when its line is gone or is not on the day on screen.
# -NoRedraw: the caller redraws every row anyway.
function Sync-PBSelection {
                param ([switch]$NoRedraw)

                $id = [string]$script:PB.SelectedLineId
                if (-not $id) { return }
                $line = Get-BoardLine $script:PB.Board $id
                if ($null -ne $line -and $line.Day -eq (Get-PBActiveDay)) { return }
                $script:PB.SelectedLineId = ''
                if ($null -ne $line -and -not $NoRedraw) {
                                $script:PB.Busy = $true
                                try { Update-PBDay $line.Day } finally { $script:PB.Busy = $false }
                }
}

# "0730 Day Local  |  IP Jones, Bob  |  Adams, Al on FT5001"
function Get-PBLineCaption {
                param ($Line)

                $isSim = Test-IsSimSortieType ([string]$Line.SortieType)
                $ip = if ($Line.Instructor) { "IP $($Line.Instructor)" } elseif ($isSim) { 'SIM IP' } else { 'no IP yet' }
                $who = 'no student yet'
                if ($Line.Student) {
                                $who = [string]$Line.Student
                                if ($Line.Mission) { $who += " on $($Line.Mission)" }
                }
                elseif (Test-IsReturnSortieType ([string]$Line.SortieType)) { $who = 'flies with its O&B' }
                elseif ($script:PB.Mode -eq 'Review' -and $null -ne $script:PB.Overlay -and $null -ne $script:PB.Overlay[$Line.Id] -and $script:PB.Overlay[$Line.Id].Student) {
                                $ov = $script:PB.Overlay[$Line.Id]
                                $who = ("{0} {1}" -f $ov.Student, $ov.Mission).Trim() + ' (filled by the scheduler)'
                }
                return ("{0} {1}   |   {2}   |   {3}" -f (Format-BoardTime $Line.ETD), $Line.SortieType, $ip, $who)
}

function Initialize-PBDetails {
                $script:PB.Details = @{
                                Key           = ''
                                LineId        = ''
                                Focus         = ''
                                Queue         = (New-Object System.Collections.Generic.List[string])
                                All           = (New-Object System.Collections.Generic.List[object])
                                Items         = (New-Object 'System.Collections.ObjectModel.ObservableCollection[object]')
                                Total         = 0
                                ChecksPending = $false
                }
                $script:PB.C.DetailsList.ItemsSource = $script:PB.Details.Items
}

# Shows or hides the panel (the Hide button, View > Line details panel). Its height is kept.
function Show-PBDetailsPanel {
                param ([bool]$Visible)

                $C = $script:PB.C
                if ([bool]$script:PB.DetailsVisible -eq $Visible) { return }
                $script:PB.DetailsVisible = $Visible
                $row = $C.CenterGrid.RowDefinitions[3]
                if ($Visible) {
                                $row.Height = New-Object System.Windows.GridLength([double]$script:PB.DetailsHeight)
                                $C.DetailsPanel.Visibility = 'Visible'
                                $C.DetailsSplitter.Visibility = 'Visible'
                }
                else {
                                if ($row.ActualHeight -gt 60) { $script:PB.DetailsHeight = [double]$row.ActualHeight }
                                $row.Height = New-Object System.Windows.GridLength(0)
                                $C.DetailsPanel.Visibility = 'Collapsed'
                                $C.DetailsSplitter.Visibility = 'Collapsed'
                }
                if ([bool]$C.MenuShowDetails.IsChecked -ne $Visible) { $C.MenuShowDetails.IsChecked = $Visible }
                Update-PBDetails -Force
}

# Starts the panel over for the selected line, unless it already shows this line as the board
# is now. -Force starts over anyway.
function Update-PBDetails {
                param ([switch]$Force)

                $d = $script:PB.Details
                if ($null -eq $d) { return }
                $C = $script:PB.C
                $id = [string]$script:PB.SelectedLineId
                $line = if ($id) { Get-BoardLine $script:PB.Board $id } else { $null }
                $key = if ($null -ne $line -and $script:PB.DetailsVisible) { '{0}|{1}|{2}' -f $id, $script:PB.Board.Version, $script:PB.Mode } else { '' }
                if (-not $Force -and $key -eq $d.Key) { return }
                # The student picked in the list stays picked while the same line is shown
                if ($id -ne $d.LineId) { $d.Focus = '' }
                $d.Key = $key
                $d.LineId = $id
                $d.Queue.Clear()
                $d.All.Clear()
                $d.Items.Clear()
                $d.Total = 0
                $d.ChecksPending = $false
                $script:PB.DetailsTimer.Stop()
                $C.DetailsPlaceButton.IsEnabled = $false
                if (-not $script:PB.DetailsVisible) { return }

                if ($null -eq $line) {
                                $C.DetailsTitle.Text = 'Click a line to see who fits it and why.'
                                Set-PBDetailsChecks @(New-PBDetailsText 'Click anywhere on a line to see its student checked against every hard rule, warning and soft rule, and every student ranked for that line.' -Kind 'Info')
                                $C.DetailsListInfo.Text = ''
                                return
                }
                $C.DetailsTitle.Text = Get-PBLineCaption $line
                if (Test-IsReturnSortieType ([string]$line.SortieType)) {
                                Update-PBDetailsChecks
                                $C.DetailsListInfo.Text = "Return lines fly with their O&B's student, so nobody is put on them directly."
                                return
                }
                foreach ($s in @($script:PB.Data.Students | Sort-Object -Property Name)) {
                                if (-not [string]::IsNullOrWhiteSpace([string]$s.Name)) { $d.Queue.Add([string]$s.Name) }
                }
                $d.Total = $d.Queue.Count
                $d.ChecksPending = $true
                Set-PBDetailsChecks @(New-PBDetailsText 'Checking...' -Kind 'Info')
                Update-PBDetailsListInfo
                $script:PB.DetailsTimer.Start()
}

# Every 30 ms: the left side first, then as many students as fit in 30 ms
function Invoke-PBDetailsTick {
                if ($script:PB.Busy -or $script:PB.Draining -or $script:PB.Filling) { return }
                if ($null -ne $script:PB.Drag -and $script:PB.Drag.Active) { return }
                $d = $script:PB.Details
                $line = if ($null -ne $d -and $d.LineId) { Get-BoardLine $script:PB.Board $d.LineId } else { $null }
                if ($null -eq $line) { $script:PB.DetailsTimer.Stop(); return }
                # A change the panel was not told about (a note, a callsign): start over
                if ($d.Key -ne ('{0}|{1}|{2}' -f $d.LineId, $script:PB.Board.Version, $script:PB.Mode)) { Update-PBDetails; return }
                if ($d.ChecksPending) {
                                $d.ChecksPending = $false
                                Update-PBDetailsChecks
                                return
                }
                $sw = [System.Diagnostics.Stopwatch]::StartNew()
                while ($d.Queue.Count -gt 0 -and $sw.ElapsedMilliseconds -lt 30) {
                                $name = $d.Queue[0]
                                $d.Queue.RemoveAt(0)
                                Add-PBCandidate (New-PBCandidateVM -Line $line -Name $name -Assessment (Get-PBCandidateAssessment -Line $line -Name $name))
                }
                Update-PBDetailsListInfo
                if ($d.Queue.Count -eq 0) { $script:PB.DetailsTimer.Stop() }
}

# One student on the line: the line's own student with its mission, anyone else on (auto)
function Get-PBCandidateAssessment {
                param ($Line, [string]$Name, [string]$Mission = '')

                if (-not $Mission -and $Name -ieq [string]$Line.Student) { $Mission = [string]$Line.Mission }
                return Get-PlacementAssessment -Board $script:PB.Board -LineId $Line.Id -StudentName $Name -MissionName $Mission -Data $script:PB.Data
}

function Get-PBCandidateRank {
                param ([string]$Status)

                switch ($Status) {
                                'Clean' { return 0 }
                                'Warn' { return 1 }
                                'Hard' { return 2 }
                }
                return 3
}

# The biggest + / - score reasons, e.g. "Continue String +10, Exhaustion -6"
function Get-PBTopScoreReason {
                param ($Assessment)

                $chosen = $Assessment.EventDetails | Where-Object { $_.Mission -eq $Assessment.Mission } | Select-Object -First 1
                if ($null -eq $chosen) { return 'Fits' }
                $top = @($chosen.Breakdown | Where-Object { $_.Points -ne 0 } | Sort-Object -Property { [math]::Abs($_.Points) } -Descending | Select-Object -First 2)
                if ($top.Count -eq 0) { return 'Fits' }
                return (@($top | ForEach-Object { '{0} {1:+0.#;-0.#}' -f $_.Label, $_.Points }) -join ', ')
}

# A row of the "Every student on this line" list
function New-PBCandidateVM {
                param ($Line, [string]$Name, $Assessment)

                $a = $Assessment
                $hard = @($a.Violations | Where-Object { $_.Severity -eq 'Hard' })
                $soft = @($a.Violations | Where-Object { $_.Severity -ne 'Hard' })
                $reasonFg = '#475569'
                if ($hard.Count -gt 0) {
                                $reason = [string]$hard[0].Message
                                if ($hard.Count -gt 1) { $reason += " (and $($hard.Count - 1) more)" }
                                $reasonFg = '#B91C1C'
                }
                elseif ($a.Status -eq 'Hard') { $reason = 'No event they can fly fits this line'; $reasonFg = '#B91C1C' }
                elseif ($soft.Count -gt 0) { $reason = [string]$soft[0].Message; $reasonFg = '#B45309' }
                elseif (@($a.Notes).Count -gt 0) { $reason = [string]@($a.Notes)[0]; $reasonFg = '#B45309' }
                else { $reason = Get-PBTopScoreReason $a; $reasonFg = '#15803D' }

                $where = ''
                $ov = $null
                if ($script:PB.Mode -eq 'Review' -and $null -ne $script:PB.Overlay) { $ov = $script:PB.Overlay[$Line.Id] }
                if ($Name -ieq [string]$Line.Student) { $where = '  (on this line)' }
                elseif ($null -ne $ov -and [string]$ov.Student -ieq $Name) { $where = '  (scheduler''s pick)' }
                $icon = '!'
                $iconFg = '#D97706'
                if ($a.Status -eq 'Clean') { $icon = [string][char]0x2714; $iconFg = '#16A34A' }
                elseif ($a.Status -ne 'Warn') { $icon = [string][char]0x2716; $iconFg = '#DC2626' }
                $tip = @(Get-AssessmentLines $a | Where-Object { $_.Kind -ne 'Info' } | ForEach-Object { $_.Text }) -join "`n"

                return [PSCustomObject]@{
                                StudentName = $Name
                                NameText    = "$Name$where"
                                Icon        = $icon
                                IconBrush   = $iconFg
                                MissionText = $(if ($a.Mission) { [string]$a.Mission } else { '-' })
                                ScoreText   = $(if ($null -ne $a.Score) { '{0:+0.0;-0.0;0.0}' -f [double]$a.Score } else { '' })
                                ReasonText  = $reason
                                ReasonBrush = $reasonFg
                                ToolTipText = "$tip`nDouble-click to hardcode $Name on this line."
                                Rank        = (Get-PBCandidateRank ([string]$a.Status))
                                SortScore   = $(if ($null -ne $a.Score) { [double]$a.Score } else { [double]::MinValue })
                                Status      = [string]$a.Status
                }
}

# True when -A goes above -B: fits, then warnings, then rule breaks; best score first; then by name
function Test-PBCandidateBefore {
                param ($A, $B)

                if ($A.Rank -ne $B.Rank) { return ($A.Rank -lt $B.Rank) }
                if ($A.SortScore -ne $B.SortScore) { return ($A.SortScore -gt $B.SortScore) }
                return ([string]::Compare([string]$A.StudentName, [string]$B.StudentName, $true) -lt 0)
}

# "Hide rule breaks" leaves out the students who break a hard rule
function Test-PBCandidateShown {
                param ($Candidate)

                return -not ([bool]$script:PB.C.DetailsHideHardCheck.IsChecked -and $Candidate.Rank -ge 2)
}

# Puts a student into the list in order, keeping the student picked in it picked
function Add-PBCandidate {
                param ($Candidate)

                $d = $script:PB.Details
                $at = 0
                while ($at -lt $d.All.Count -and -not (Test-PBCandidateBefore $Candidate $d.All[$at])) { $at++ }
                $d.All.Insert($at, $Candidate)
                if (-not (Test-PBCandidateShown $Candidate)) { return }
                $pos = 0
                for ($i = 0; $i -lt $at; $i++) { if (Test-PBCandidateShown $d.All[$i]) { $pos++ } }
                $d.Items.Insert($pos, $Candidate)
                if ($d.Focus -and $Candidate.StudentName -eq $d.Focus -and $null -eq $script:PB.C.DetailsList.SelectedItem) { $script:PB.C.DetailsList.SelectedItem = $Candidate }
}

function Update-PBDetailsFilter {
                $d = $script:PB.Details
                if ($null -eq $d) { return }
                $d.Items.Clear()
                foreach ($cand in $d.All) {
                                if (-not (Test-PBCandidateShown $cand)) { continue }
                                $d.Items.Add($cand)
                                if ($d.Focus -and $cand.StudentName -eq $d.Focus) { $script:PB.C.DetailsList.SelectedItem = $cand }
                }
                Update-PBDetailsListInfo
}

# "Checking 30 of 98 students... 4 fit, 6 with warnings, 20 break a hard rule"
function Update-PBDetailsListInfo {
                $d = $script:PB.Details
                $fit = 0; $warn = 0; $hard = 0
                foreach ($cand in $d.All) {
                                if ($cand.Rank -eq 0) { $fit++ } elseif ($cand.Rank -eq 1) { $warn++ } else { $hard++ }
                }
                $text = '{0} fit, {1} with warnings, {2} break a hard rule' -f $fit, $warn, $hard
                if ($d.Queue.Count -gt 0) { $text = "Checking $($d.All.Count) of $($d.Total) students... $text" }
                else { $text = "$($d.Total) students: $text." }
                $hidden = $d.All.Count - $d.Items.Count
                if ($hidden -gt 0) { $text += " ($hidden hidden)" }
                if ($d.Queue.Count -eq 0) { $text += ' Double-click a student, or pick one and press Put on this line, to hardcode them here.' }
                $script:PB.C.DetailsListInfo.Text = $text
}

# A student picked in the list: their check goes on the left
function Invoke-PBDetailsPick {
                $d = $script:PB.Details
                if ($null -eq $d) { return }
                $item = $script:PB.C.DetailsList.SelectedItem
                $line = if ($d.LineId) { Get-BoardLine $script:PB.Board $d.LineId } else { $null }
                $script:PB.C.DetailsPlaceButton.IsEnabled = ($null -ne $item -and $null -ne $line -and -not (Test-IsReturnSortieType ([string]$line.SortieType)) -and
                                [string]$item.StudentName -ine [string]$line.Student)
                # The list emptying while it is rebuilt keeps the pick
                if ($null -eq $item -or [string]$item.StudentName -eq $d.Focus) { return }
                $d.Focus = [string]$item.StudentName
                Update-PBDetailsChecks
}

# Put on this line, or a double-click in the list
function Invoke-PBDetailsPlace {
                $item = $script:PB.C.DetailsList.SelectedItem
                $id = if ($null -ne $script:PB.Details) { [string]$script:PB.Details.LineId } else { '' }
                if ($null -eq $item -or -not $id) { return }
                Request-PBCommand @{ Cmd = 'PlaceOnLine'; Kind = 'Student'; Name = [string]$item.StudentName; LineId = $id }
}

# The left side: who is on the line (or picked in the list) and how they do against every rule
function Update-PBDetailsChecks {
                $d = $script:PB.Details
                $line = if ($d.LineId) { Get-BoardLine $script:PB.Board $d.LineId } else { $null }
                if ($null -eq $line) { return }
                $parts = New-Object System.Collections.Generic.List[object]

                if (Test-IsReturnSortieType ([string]$line.SortieType)) {
                                $ob = if ($line.LinkedLineId) { Get-BoardLine $script:PB.Board $line.LinkedLineId } else { $null }
                                $parts.Add((New-PBDetailsText 'O&B Return line' -Kind 'Head'))
                                if ($null -eq $ob) { $parts.Add((New-PBDetailsText 'Not linked to an O&B, so nobody will fly it. Delete it, or type the return again with RTN after its O&B.' -Kind 'Warn')) }
                                else {
                                                $who = if ($ob.Student) { "$($ob.Student) is hardcoded on it, so they fly this return too." } else { 'Whoever is put on the O&B flies this return too.' }
                                                $parts.Add((New-PBDetailsText ("Flies with the {0} O&B. {1}" -f (Format-BoardTime $ob.ETD), $who)))
                                }
                                Set-PBDetailsChecks $parts.ToArray()
                                return
                }

                $pinned = [string]$line.Student
                $ov = $null
                if ($script:PB.Mode -eq 'Review' -and $null -ne $script:PB.Overlay) { $ov = $script:PB.Overlay[$line.Id] }
                $fromFill = ($null -ne $ov -and -not $pinned -and $ov.Student -and @('Engine', 'Wing', 'DS') -contains [string]$ov.FilledBy)
                $name = [string]$d.Focus
                if (-not $name) { $name = $pinned }
                if (-not $name -and $fromFill) { $name = [string]$ov.Student }
                if (-not $name) {
                                if ($null -ne $ov -and [string]$ov.FilledBy -eq 'ForSale') {
                                                $parts.Add((New-PBDetailsText 'For Sale' -Kind 'Head'))
                                                $parts.Add((New-PBDetailsText 'No student could fly this line when it was filled. The list on the right shows why each one can''t.'))
                                }
                                else {
                                                $parts.Add((New-PBDetailsText 'Nobody is on this line yet' -Kind 'Head'))
                                                $parts.Add((New-PBDetailsText 'The scheduler fills it when you press Fill. To hardcode someone, pick a student in the list to see how they would do here, then double-click them (or press Put on this line).'))
                                }
                                Set-PBDetailsChecks $parts.ToArray()
                                return
                }

                $mission = ''
                $how = 'if they flew this line'
                if ($name -ieq $pinned) { $how = 'hardcoded on this line' }
                elseif ($fromFill -and [string]$ov.Student -ieq $name) { $how = 'picked by the scheduler'; $mission = [string]$ov.Mission }
                $a = Get-PBCandidateAssessment -Line $line -Name $name -Mission $mission
                $parts.Add((New-PBDetailsText "$name, $how" -Kind 'Head'))
                foreach ($p in @(New-PBDetailsSections $a)) { $parts.Add($p) }
                if ($pinned -and $name -ine $pinned) { $parts.Add((New-PBDetailsText "$pinned is on this line now: pick them in the list to see theirs." -Kind 'Info')) }
                Set-PBDetailsChecks $parts.ToArray()
}

# An assessment as headed sections: the verdict, HARD RULES, WARNINGS, SOFT RULES (the score)
function New-PBDetailsSections {
                param ($Assessment)

                $a = $Assessment
                $out = New-Object System.Collections.Generic.List[object]
                $lines = @(Get-AssessmentLines $a)
                $hardLines = @($lines | Where-Object { $_.Kind -eq 'Hard' })
                $warnLines = @($lines | Where-Object { $_.Kind -eq 'Warn' -or $_.Kind -eq 'Note' })
                $softLines = @($lines | Where-Object { $_.Kind -eq 'Plus' -or $_.Kind -eq 'Minus' })
                $infoLines = @($lines | Where-Object { $_.Kind -eq 'Info' })
                $tick = [string][char]0x2714
                $cross = [string][char]0x2716

                if ($a.Status -eq 'Clean') { $out.Add((New-PBDetailsText "$tick  Passes every hard rule." -Kind 'Good')) }
                elseif ($a.Status -eq 'Warn') { $out.Add((New-PBDetailsText "$tick  Passes the hard rules, with warnings worth a look." -Kind 'Warn')) }
                else { $out.Add((New-PBDetailsText "$cross  Breaks a hard rule. Allowed if you hardcode it, but the line is flagged OVERRIDE on the output." -Kind 'Hard')) }
                $scoreText = if ($null -ne $a.Score) { '{0:+0.0;-0.0;0.0}' -f [double]$a.Score } else { '' }
                $mission = if ($a.Mission) { [string]$a.Mission } else { 'no event' }
                $how = if ($a.RequestedMission) { 'hardcoded' } else { 'the best-scoring event that passes, as (auto) picks' }
                $out.Add((New-PBDetailsText ("Mission {0} ({1}){2}" -f $mission, $how, $(if ($scoreText) { ", score $scoreText" } else { '' }))))

                $out.Add((New-PBDetailsText 'HARD RULES' -Kind 'Section'))
                if ($hardLines.Count -eq 0 -and $a.Status -eq 'Hard') { $out.Add((New-PBDetailsText "$cross  No event they can fly fits this line" -Kind 'Hard')) }
                elseif ($hardLines.Count -eq 0) { $out.Add((New-PBDetailsText "$tick  None broken" -Kind 'Good')) }
                foreach ($ln in $hardLines) { $out.Add((New-PBDetailsText "$cross  $($ln.Text)" -Kind 'Hard')) }

                $out.Add((New-PBDetailsText 'WARNINGS' -Kind 'Section'))
                if ($warnLines.Count -eq 0) { $out.Add((New-PBDetailsText 'None' -Kind 'Muted')) }
                foreach ($ln in $warnLines) {
                                $mark = if ($ln.Kind -eq 'Warn') { [string][char]0x26A0 } else { [string][char]0x2022 }
                                $out.Add((New-PBDetailsText "$mark  $($ln.Text)" -Kind 'Warn'))
                }

                $out.Add((New-PBDetailsText ('SOFT RULES' + $(if ($scoreText) { " (SCORE $scoreText)" } else { '' })) -Kind 'Section'))
                if ($softLines.Count -eq 0) { $out.Add((New-PBDetailsText 'No score reasons' -Kind 'Muted')) }
                foreach ($ln in $softLines) {
                                if ($ln.Kind -eq 'Plus') { $out.Add((New-PBDetailsText "+  $($ln.Text)" -Kind 'Good')) }
                                else { $out.Add((New-PBDetailsText "$([char]0x2212)  $($ln.Text)" -Kind 'Hard')) }
                }
                foreach ($ln in $infoLines) { $out.Add((New-PBDetailsText $ln.Text -Kind 'Info')) }
                return $out.ToArray()
}

# One line of text in the panel. -Kind: Head, Section, Text, Good, Warn, Hard, Muted, Info
function New-PBDetailsText {
                param ([string]$Text, [string]$Kind = 'Text')

                $tb = New-Object System.Windows.Controls.TextBlock
                $tb.Text = $Text
                $tb.TextWrapping = [System.Windows.TextWrapping]::Wrap
                $tb.Margin = New-Object System.Windows.Thickness(0, 1, 0, 1)
                switch ($Kind) {
                                'Head' { $tb.FontWeight = [System.Windows.FontWeights]::Bold; $tb.FontSize = 13; $tb.Margin = New-Object System.Windows.Thickness(0, 0, 0, 3) }
                                'Section' { $tb.FontWeight = [System.Windows.FontWeights]::SemiBold; $tb.FontSize = 10; $tb.Foreground = Get-PBBrush '#64748B'; $tb.Margin = New-Object System.Windows.Thickness(0, 7, 0, 1) }
                                'Good' { $tb.Foreground = Get-PBBrush '#15803D' }
                                'Warn' { $tb.Foreground = Get-PBBrush '#B45309' }
                                'Hard' { $tb.Foreground = Get-PBBrush '#B91C1C'; $tb.FontWeight = [System.Windows.FontWeights]::SemiBold }
                                'Muted' { $tb.Foreground = Get-PBBrush '#94A3B8'; $tb.FontStyle = [System.Windows.FontStyles]::Italic }
                                'Info' { $tb.Foreground = Get-PBBrush '#64748B'; $tb.FontStyle = [System.Windows.FontStyles]::Italic; $tb.FontSize = 11 }
                }
                return $tb
}

function Set-PBDetailsChecks {
                param ([array]$Parts)

                $panel = $script:PB.C.DetailsChecks
                $panel.Children.Clear()
                foreach ($p in $Parts) { [void]$panel.Children.Add($p) }
}


# =================================================================================
#  MENU BAR
#  The items are wired in Register-PBMenuHandlers. A top menu brings its items up to date as
#  it opens: what can't be done now is greyed out, and the Day and Fill items name their day.
# =================================================================================

function Update-PBMenuState {
                $C = $script:PB.C
                $review = ($script:PB.Mode -eq 'Review')
                $active = Get-PBActiveDay
                $tabs = $C.DayTabs
                $C.MenuUndo.IsEnabled = ($script:PB.Board.Undo.Count -gt 0)
                $C.MenuLetGo.IsEnabled = ($null -ne $script:PB.Held)
                $C.MenuPrevDay.IsEnabled = ($tabs.SelectedIndex -gt 0)
                $C.MenuNextDay.IsEnabled = ($tabs.SelectedIndex -ge 0 -and $tabs.SelectedIndex -lt $tabs.Items.Count - 1)
                $C.MenuAddNextDay.Header = "Add the next flying day ($(Format-BoardDayLabel ((Get-PBNextBoardDay).ToString('yyyy-MM-dd'))))"
                $picked = $C.AddDayPicker.SelectedDate
                $C.MenuAddPickedDay.IsEnabled = ($null -ne $picked)
                $C.MenuAddPickedDay.Header = if ($null -ne $picked) { "Add the day in the date box ($(Format-BoardDayLabel (([datetime]$picked).ToString('yyyy-MM-dd'))))" } else { 'Add the day in the date box' }
                $C.MenuRemoveDay.IsEnabled = ($null -ne $active)
                $C.MenuApplyTemplate.IsEnabled = ($null -ne $active)
                $label = if ($active) { Format-BoardDayLabel $active } else { 'this day' }
                $C.MenuFillDay.Header = "Fill _this day ($label)"
                $C.MenuFillAll.Header = "Fill _all days ($($script:PB.Board.Days.Count))"
                $C.MenuFillDay.IsEnabled = (-not $review -and $null -ne $active)
                $C.MenuFillAll.IsEnabled = (-not $review)
                foreach ($mi in @($C.MenuRefill, $C.MenuBackToEdit, $C.MenuExport)) { $mi.IsEnabled = $review }
                $C.MenuShowDetails.IsChecked = [bool]$script:PB.DetailsVisible
}

# The Line menu: the selected line's right-click menu, with the keys that act on it
function Update-PBLineMenu {
                $menu = $script:PB.C.MenuLine
                $id = [string]$script:PB.SelectedLineId
                $line = if ($id) { Get-BoardLine $script:PB.Board $id } else { $null }
                if ($null -eq $line) {
                                $menu.Items.Clear()
                                $none = New-PBMenuItem 'Click a line on the board first' $null
                                $none.IsEnabled = $false
                                [void]$menu.Items.Add($none)
                                return
                }
                Set-PBMenuItems -Menu $menu -Tag 'Line' -Line $line
                $head = New-PBMenuItem (Get-PBLineCaption $line) $null
                $head.IsEnabled = $false
                $menu.Items.Insert(0, $head)
}

# File > Apply a template: one item per saved template
function Update-PBTemplateMenu {
                $menu = $script:PB.C.MenuApplyTemplate
                $menu.Items.Clear()
                $list = @(Get-BoardTemplateList -Dir $script:PB.Settings.TemplateDir)
                if ($list.Count -eq 0) {
                                $none = New-PBMenuItem '(no templates saved yet: File > Save this day as a template)' $null
                                $none.IsEnabled = $false
                                [void]$menu.Items.Add($none)
                                return
                }
                foreach ($t in $list) {
                                # A single _ in a menu header would be read as an access key
                                $mi = New-PBMenuItem (([string]$t.Name) -replace '_', '__') $null
                                $mi.Tag = $t
                                $mi.Add_Click({ param($s, $a) Invoke-PBSafe { Invoke-PBLoadTemplate -Template $s.Tag } })
                                [void]$menu.Items.Add($mi)
                }
}

# The next flying day after the board's last day
function Get-PBNextBoardDay {
                $days = $script:PB.Board.Days
                if ($days.Count -eq 0) { return (Get-NextFlyingDay) }
                return (Get-NextFlyingDay -From (ConvertTo-PBDayDate ([string]$days[$days.Count - 1])))
}

# The day tab -Offset tabs along (Ctrl+PgUp / Ctrl+PgDn, Day menu)
function Select-PBDayOffset {
                param ([int]$Offset)

                $tabs = $script:PB.C.DayTabs
                if ($tabs.Items.Count -eq 0) { return }
                $i = [math]::Max(0, [math]::Min($tabs.Items.Count - 1, $tabs.SelectedIndex + $Offset))
                if ($i -ne $tabs.SelectedIndex) { $tabs.SelectedIndex = $i }
}

# F5 / Edit > Re-check every hardcoded line
function Invoke-PBRecheck {
                $script:PB.Assess = @{}
                Clear-AssessmentCache
                Start-PBAssessQueue
                Update-PBDetails -Force
                Set-PBStatus 'Re-checking every hardcoded line...'
}

function Invoke-PBBackToEdit {
                Exit-PBReview
                Update-PBAll
                Set-PBStatus 'Back to editing. The fill was discarded.'
}


# =================================================================================
#  COLOR KEY (the Colors button, View or Help menu)
#  A small window that stays open beside the board.
# =================================================================================

# Every color on the board: sections of { Title; Items = @{ Swatch; Fill; Border; Thick; Label; Text } }.
# Swatch: Stripe (a line's left edge), Marker (a hardcoded slot's corner), Pill (a slot or puck),
# Badge (a count on a puck), Row (a whole line), None (text only).
function Get-PBColorKey {
                $K = $script:PBColors
                $item = {
                                param ([string]$Swatch, [string]$Fill, [string]$Text, [string]$Border = '', [int]$Thick = 1, [string]$Label = '')
                                @{ Swatch = $Swatch; Fill = $Fill; Text = $Text; Border = $Border; Thick = $Thick; Label = $Label }
                }
                return @(
                                @{ Title = 'Line type (the stripe at the left of a line)'; Items = @(
                                                (& $item 'Stripe' $K.StripeDayLocal 'Day Local')
                                                (& $item 'Stripe' $K.StripeOB 'O&B')
                                                (& $item 'Stripe' $K.StripeReturn 'O&B Return (flies with its O&B''s student)')
                                                (& $item 'Stripe' $K.StripeSim 'OFT / IFT')
                                                (& $item 'Stripe' $K.StripeOther 'Any other type')) }
                                @{ Title = 'Corner marker on a hardcoded student (hover it for the reasons)'; Items = @(
                                                (& $item 'Marker' $K.MarkerClean 'Passes every hard rule')
                                                (& $item 'Marker' $K.MarkerWarn 'Passes, with warnings worth a look')
                                                (& $item 'Marker' $K.MarkerHard 'Breaks a hard rule: allowed, but flagged OVERRIDE on the output')
                                                (& $item 'Marker' $K.MarkerPending 'Still being checked')) }
                                @{ Title = 'Student slots while you drag or hold a student'; Items = @(
                                                (& $item 'Pill' $K.HeatClean 'Fits' '#CBD5E1')
                                                (& $item 'Pill' $K.HeatWarn 'Fits, with warnings' '#CBD5E1')
                                                (& $item 'Pill' $K.HeatHard 'Breaks a hard rule' '#CBD5E1')) }
                                @{ Title = 'Student slot on a line'; Items = @(
                                                (& $item 'Pill' '#DCFCE7' 'Hardcoded student' '#22C55E' 2)
                                                (& $item 'Pill' '#F8FAFC' 'Open: the scheduler fills it' '#CBD5E1')
                                                (& $item 'Pill' '#F5F3FF' 'Return leg: follows its O&B' '#DDD6FE')) }
                                @{ Title = 'After a fill, while you review it'; Items = @(
                                                (& $item 'Pill' '#EFF6FF' 'Picked by the scheduler (AUTO), or its wingman (WING), in italics' '#60A5FA')
                                                (& $item 'Pill' '#FEF3C7' 'Direct support for a formation' '#F59E0B')
                                                (& $item 'Pill' '#FEE2E2' 'For Sale: no student could fly the line' '#FCA5A5')
                                                (& $item 'Row' '#FEF2F2' 'Pink line, red notes: OVERRIDE (hardcoded against a rule)' '#E2E8F0')) }
                                @{ Title = 'Student pucks (right)'; Items = @(
                                                (& $item 'Pill' '#DCFCE7' 'Available' '#86EFAC')
                                                (& $item 'Pill' '#FEF3C7' 'Has an unsat (U-SIM, U-FLT) or is absent part of the day (ABS)' '#FCD34D')
                                                (& $item 'Pill' '#E5E7EB' 'On hold (HOLD)' '#9CA3AF')
                                                (& $item 'Badge' '#16A34A' '1 turn that day' '' 1 '1 turn')
                                                (& $item 'Badge' '#D97706' '2 turns that day' '' 1 '2 turns')
                                                (& $item 'Badge' '#DC2626' '3 or more turns that day' '' 1 '3 turns')
                                                (& $item 'None' '' 'CAP: one flight a day, no O&Bs, needs an experienced IP. INTL: international, needs an experienced IP. U-SIM / U-FLT: unsat sim / flight. ABS: absent at those times.')) }
                                @{ Title = 'IP pucks (left)'; Items = @(
                                                (& $item 'Pill' '#DBEAFE' 'IP' '#93C5FD')
                                                (& $item 'Badge' '#1D4ED8' 'Lines that IP has that day' '' 1 'x2')
                                                (& $item 'None' '' 'EXP: experienced IP. CHK: check IP. NC: non-counter.')) }
                                @{ Title = 'Picking and holding'; Items = @(
                                                (& $item 'Row' '#EFF6FF' 'The selected line (Line details shows it)' '#2563EB')
                                                (& $item 'Pill' '#DCFCE7' 'A puck you are holding: click a slot to put it there' '#1D4ED8' 3)) }
                )
}

function Show-PBColorKey {
                $win = $script:PB.ColorKeyWindow
                if ($null -ne $win -and $win.IsVisible) { [void]$win.Activate(); return }
                $win = New-PBColorKeyWindow
                $script:PB.ColorKeyWindow = $win
                try { if ($script:PB.Window.IsVisible) { $win.Owner = $script:PB.Window } } catch { }
                $win.Show()
}

function New-PBColorKeyWindow {
                $win = New-Object System.Windows.Window
                $win.Title = 'FAIS color key'
                $win.Width = 470
                $win.Height = 640
                $win.ShowInTaskbar = $false
                $win.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterOwner
                $win.Background = Get-PBBrush '#F8FAFC'
                $win.FontFamily = New-Object System.Windows.Media.FontFamily('Segoe UI')
                $win.FontSize = 12

                $panel = New-Object System.Windows.Controls.StackPanel
                $panel.Margin = New-Object System.Windows.Thickness(14)
                $title = New-Object System.Windows.Controls.TextBlock
                $title.Text = 'What the colors mean'
                $title.FontSize = 15
                $title.FontWeight = [System.Windows.FontWeights]::Bold
                [void]$panel.Children.Add($title)
                foreach ($section in @(Get-PBColorKey)) {
                                $head = New-Object System.Windows.Controls.TextBlock
                                $head.Text = $section.Title
                                $head.FontWeight = [System.Windows.FontWeights]::SemiBold
                                $head.Foreground = Get-PBBrush '#334155'
                                $head.Margin = New-Object System.Windows.Thickness(0, 10, 0, 3)
                                [void]$panel.Children.Add($head)
                                foreach ($it in $section.Items) {
                                                $rowPanel = New-Object System.Windows.Controls.DockPanel
                                                $rowPanel.Margin = New-Object System.Windows.Thickness(0, 1, 0, 1)
                                                if ($it.Swatch -ne 'None') {
                                                                $swatch = New-PBColorSwatch $it
                                                                [System.Windows.Controls.DockPanel]::SetDock($swatch, [System.Windows.Controls.Dock]::Left)
                                                                [void]$rowPanel.Children.Add($swatch)
                                                }
                                                $text = New-Object System.Windows.Controls.TextBlock
                                                $text.Text = $it.Text
                                                $text.TextWrapping = [System.Windows.TextWrapping]::Wrap
                                                $text.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
                                                if ($it.Swatch -eq 'None') { $text.Foreground = Get-PBBrush '#475569'; $text.FontSize = 11 }
                                                [void]$rowPanel.Children.Add($text)
                                                [void]$panel.Children.Add($rowPanel)
                                }
                }
                $scroll = New-Object System.Windows.Controls.ScrollViewer
                $scroll.VerticalScrollBarVisibility = [System.Windows.Controls.ScrollBarVisibility]::Auto
                $scroll.Content = $panel
                $win.Content = $scroll
                return $win
}

# A small sample of one color, drawn the way the board draws it
function New-PBColorSwatch {
                param ($Item)

                $box = New-Object System.Windows.Controls.Grid
                $box.Width = 52
                $box.Height = 20
                $box.Margin = New-Object System.Windows.Thickness(0, 1, 10, 1)
                $shape = New-Object System.Windows.Controls.Border
                $shape.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Left
                $shape.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
                switch ($Item.Swatch) {
                                'Stripe' {
                                                $shape.Width = 8; $shape.Height = 18
                                                $shape.CornerRadius = New-Object System.Windows.CornerRadius(2)
                                                $shape.Background = Get-PBBrush $Item.Fill
                                }
                                'Marker' {
                                                # A hardcoded slot with its corner triangle
                                                $shape.Width = 44; $shape.Height = 18
                                                $shape.CornerRadius = New-Object System.Windows.CornerRadius(9)
                                                $shape.Background = Get-PBBrush '#DCFCE7'
                                                $shape.BorderBrush = Get-PBBrush '#22C55E'
                                                $shape.BorderThickness = New-Object System.Windows.Thickness(2)
                                                $tri = New-Object System.Windows.Shapes.Polygon
                                                $tri.Points = [System.Windows.Media.PointCollection]::Parse('0,0 10,0 10,10')
                                                $tri.Fill = Get-PBBrush $Item.Fill
                                                $tri.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Right
                                                $tri.VerticalAlignment = [System.Windows.VerticalAlignment]::Top
                                                $tri.Margin = New-Object System.Windows.Thickness(0, 0, 4, 0)
                                                $shape.Child = $tri
                                }
                                'Badge' {
                                                $shape.CornerRadius = New-Object System.Windows.CornerRadius(8)
                                                $shape.Padding = New-Object System.Windows.Thickness(6, 0, 6, 0)
                                                $shape.Background = Get-PBBrush $Item.Fill
                                                $label = New-Object System.Windows.Controls.TextBlock
                                                $label.Text = $Item.Label
                                                $label.Foreground = Get-PBBrush '#FFFFFF'
                                                $label.FontSize = 10
                                                $label.FontWeight = [System.Windows.FontWeights]::SemiBold
                                                $shape.Child = $label
                                }
                                default {
                                                # Pill (a slot or puck) or Row (a whole line)
                                                $shape.Width = 44; $shape.Height = 18
                                                $shape.CornerRadius = New-Object System.Windows.CornerRadius($(if ($Item.Swatch -eq 'Row') { 3 } else { 9 }))
                                                $shape.Background = Get-PBBrush $Item.Fill
                                                if ($Item.Border) { $shape.BorderBrush = Get-PBBrush $Item.Border }
                                                $shape.BorderThickness = New-Object System.Windows.Thickness([double]$Item.Thick)
                                }
                }
                [void]$box.Children.Add($shape)
                return $box
}


# =================================================================================
#  MISSIONS PICKED BY HAND
#  Any mission can go on a pinned line. One the scheduler wouldn't put there (not opted,
#  not in the syllabus, wrong line...) is explained first, with the right mission offered;
#  if the user keeps it anyway it goes on and is flagged OVERRIDE on the output.
# =================================================================================

# A choice from a pinned line's mission dropdown: (auto), a mission, or "Type a mission..."
function Invoke-PBChooseMission {
                param ([string]$LineId, [string]$Choice)

                $board = $script:PB.Board
                $line = Get-BoardLine $board $LineId
                if ($null -eq $line -or [string]::IsNullOrWhiteSpace($line.Student)) { return }
                $picked = ConvertFrom-MissionChoiceText $Choice
                $mission = [string]$picked.Mission
                if ($picked.Kind -eq 'Type') {
                                $mission = (Read-PBText -Prompt "Mission for $($line.Student) (e.g. FT5102):" -Title 'Type a mission' -Default ([string]$line.Mission)).Trim()
                                if ($mission -eq '') { Reset-PBLineRow $LineId; return }
                                $mission = Resolve-PBMissionName -Line $line -Mission $mission
                }

                $override = $false
                if ($mission -ne '') {
                                $advice = Get-PBMissionAdvice -Line $line -Mission $mission
                                if ($advice.NeedsConfirm) {
                                                $answer = Show-PBMessage -Text $advice.Text -Title $advice.Title -Buttons $advice.Buttons -Icon 'Warning'
                                                if ($answer -eq 'Yes') { $mission = $advice.Suggested }
                                                elseif ($answer -eq 'No' -or $answer -eq 'OK') { $override = $true }
                                                else { Reset-PBLineRow $LineId; Set-PBStatus 'Mission left as it was.'; return }
                                }
                }
                if ($mission -ceq [string]$line.Mission) { Reset-PBLineRow $LineId; return }

                $status = if ($mission -eq '') { 'Mission set to (auto).' } elseif ($override) { "$mission kept - it is flagged OVERRIDE on the output." } else { "Mission set to $mission." }
                Invoke-PBChange { Set-BoardMission -Board $board -Id $LineId -Mission $mission } -Status $status
                Show-PBPlacementPopup $LineId
                # The popup puts the line's summary on the status bar; a kept rule break says so instead
                if ($override) { Set-PBStatus $status 'Warn' }
}

# The rule check for a mission picked by hand (Get-MissionAdvice with both assessments)
function Get-PBMissionAdvice {
                param ($Line, [string]$Mission)

                $board = $script:PB.Board
                $requested = Get-PlacementAssessment -Board $board -LineId $Line.Id -StudentName $Line.Student -MissionName $Mission -Data $script:PB.Data
                $auto = Get-PlacementAssessment -Board $board -LineId $Line.Id -StudentName $Line.Student -MissionName '' -Data $script:PB.Data
                return Get-MissionAdvice -Requested $requested -Auto $auto -Mission $Mission
}

# A typed mission spelled the way the student's syllabus has it ("ft5102" -> "FT5102")
function Resolve-PBMissionName {
                param ($Line, [string]$Mission)

                $student = Resolve-StudentByName -Students $script:PB.Data.Students -Name $Line.Student
                if ($null -ne $student) {
                                foreach ($m in @($student.Syllabus)) {
                                                if ($null -ne $m -and ([string]$m.Mission).Trim() -ieq $Mission) { return ([string]$m.Mission).Trim() }
                                }
                }
                return $Mission
}

# Redraws a line's row the way the board has it (after a cancelled choice in its dropdown)
function Reset-PBLineRow {
                param ([string]$LineId)

                $line = Get-BoardLine $script:PB.Board $LineId
                if ($null -eq $line) { return }
                $focus = Get-PBFocusState
                $script:PB.Busy = $true
                try { Update-PBDay $line.Day } finally { $script:PB.Busy = $false }
                Restore-PBFocus $focus
}

# Message box and text prompt (wrapped so the tests can answer them).
# Show-PBMessage returns 'OK', 'Yes', 'No' or 'Cancel'; Read-PBText returns '' when cancelled.
function Show-PBMessage {
                param ([string]$Text, [string]$Title, [string]$Buttons = 'OK', [string]$Icon = 'Information')

                $answer = [System.Windows.MessageBox]::Show($Text, $Title, [System.Windows.MessageBoxButton]$Buttons, [System.Windows.MessageBoxImage]$Icon)
                return [string]$answer
}

function Read-PBText {
                param ([string]$Prompt, [string]$Title, [string]$Default = '')

                return [string][Microsoft.VisualBasic.Interaction]::InputBox($Prompt, $Title, $Default)
}


# =================================================================================
#  ADD BULK LINES
# =================================================================================

# The "Add Bulk Lines" button: asks for the start time, how many lines and the minutes
# between them, then puts them on the day being looked at (one undo step)
function Invoke-PBBulkLines {
                $day = Get-PBActiveDay
                if ($null -eq $day) { Set-PBStatus 'Add a day first.' 'Warn'; return }
                $request = Show-PBBulkLinesDialog -Day $day
                if ($null -eq $request) { return }
                [void](Add-PBBulkLines -Day $day -StartTime $request.StartTime -Count $request.Count -Every $request.Every -SortieType $request.SortieType)
}

# Returns $true when the lines went on
function Add-PBBulkLines {
                param ([string]$Day, [string]$StartTime, [string]$Count, [string]$Every = '3', [string]$SortieType = 'Day Local')

                $made = New-BulkBoardLines -Day $Day -StartTime $StartTime -Count $Count -Every $Every -SortieType $SortieType -Settings $script:PB.Board.Settings
                if (@($made.Errors).Count -gt 0) { Set-PBStatus (@($made.Errors) -join ' ') 'Warn'; return $false }
                $newLines = @($made.Lines)
                $summary = 'Added ' + (Format-PBBulkTimes -Times $made.Times -Every $Every -SortieType $SortieType -DayText (ConvertTo-PBDayDate $Day).ToString('ddd dd MMM')) + '.'
                Invoke-PBChange { Add-BoardLine -Board $script:PB.Board -Line $newLines } -Status $summary
                return $true
}

# "10 Day Local lines: 0730 to 0757, every 3 min" (-DayText adds " on Tue 20 Oct" before the colon)
function Format-PBBulkTimes {
                param ([string[]]$Times, [string]$Every, [string]$SortieType, [string]$DayText = '')

                $n = @($Times).Count
                if ($n -eq 0) { return '' }
                $noun = if ($n -eq 1) { 'line' } else { 'lines' }
                $on = if ($DayText) { " on $DayText" } else { '' }
                $text = "$n $SortieType $noun$on at $($Times[0])"
                if ($n -gt 1) { $text = "$n $SortieType $noun$on`: $($Times[0]) to $($Times[$n - 1]), every $($Every.Trim()) min" }
                if (Test-IsOBSortieType $SortieType) { $text += ' (each with its return leg)' }
                return $text
}

function ConvertTo-PBDayDate {
                param ([string]$Day)

                return [datetime]::ParseExact($Day, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
}

# The boxes' starting values: the next start after the day's last line (else 0730), 3 minutes apart
function Get-PBBulkDefaults {
                param ([string]$Day)

                $every = 3
                $start = '0730'
                $last = @(Get-BoardDayLines -Board $script:PB.Board -Day $Day | Where-Object { -not (Test-IsReturnSortieType ([string]$_.SortieType)) } | Sort-Object -Property ETD | Select-Object -Last 1)
                if ($last.Count -gt 0) {
                                $next = $last[0].ETD.AddMinutes($every)
                                if ($next.Date -eq $last[0].ETD.Date) { $start = $next.ToString('HHmm') }
                }
                return @{ StartTime = $start; Count = ''; Every = [string]$every; SortieType = 'Day Local' }
}

# The Add bulk lines window. Returns @{ StartTime; Count; Every; SortieType } or $null when cancelled.
function Show-PBBulkLinesDialog {
                param ([string]$Day)

                $win = Initialize-PBBulkLinesDialog -Day $Day
                try { [void]$win.ShowDialog() } catch { Write-PBLog "Add bulk lines dialog: $($_.Exception.Message)" }
                $result = $script:PBBulk.Result
                $script:PBBulk = $null
                return $result
}

# Builds the Add bulk lines window without showing it. Its state lives in $script:PBBulk.
function Initialize-PBBulkLinesDialog {
                param ([string]$Day)

                [xml]$xaml = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'BulkLinesDialog.xaml') -Raw -Encoding UTF8
                $win = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
                $bulk = @{ Window = $win; Day = $Day; Result = $null }
                foreach ($field in @('DayText', 'StartTimeBox', 'CountBox', 'IntervalBox', 'TypeCombo', 'PreviewText', 'AddButton', 'CancelButton')) {
                                $c = $win.FindName($field)
                                if ($null -eq $c) { throw "BulkLinesDialog.xaml is missing the element '$field'" }
                                $bulk[$field] = $c
                }
                $script:PBBulk = $bulk
                try { if ($script:PB.Window.IsVisible) { $win.Owner = $script:PB.Window } } catch { }

                $start = Get-PBBulkDefaults -Day $Day
                $bulk.DayText.Text = "Lines go on " + (ConvertTo-PBDayDate $Day).ToString('dddd dd MMM yyyy') + '.'
                $bulk.StartTimeBox.Text = $start.StartTime
                $bulk.CountBox.Text = $start.Count
                $bulk.IntervalBox.Text = $start.Every
                $returnType = [string]$script:PB.Board.Settings.ReturnSortieType
                $bulk.TypeCombo.ItemsSource = @($script:PB.TypeOptions | Where-Object { $_ -and $_ -ne $returnType })
                $bulk.TypeCombo.SelectedItem = $start.SortieType
                if ($null -eq $bulk.TypeCombo.SelectedItem) { $bulk.TypeCombo.SelectedIndex = 0 }

                $bulk.StartTimeBox.Add_TextChanged({ Invoke-PBSafe { Update-PBBulkPreview } })
                $bulk.CountBox.Add_TextChanged({ Invoke-PBSafe { Update-PBBulkPreview } })
                $bulk.IntervalBox.Add_TextChanged({ Invoke-PBSafe { Update-PBBulkPreview } })
                $bulk.TypeCombo.Add_SelectionChanged({ Invoke-PBSafe { Update-PBBulkPreview } })
                $bulk.AddButton.Add_Click({ Invoke-PBSafe { Submit-PBBulkLinesDialog } })
                $win.Add_Loaded({ Invoke-PBSafe { [void]$script:PBBulk.CountBox.Focus() } })
                Update-PBBulkPreview
                return $win
}

# What the boxes ask for right now
function Get-PBBulkRequest {
                $bulk = $script:PBBulk
                return @{
                                StartTime  = [string]$bulk.StartTimeBox.Text
                                Count      = [string]$bulk.CountBox.Text
                                Every      = [string]$bulk.IntervalBox.Text
                                SortieType = [string]$bulk.TypeCombo.SelectedItem
                }
}

# Shows the times the lines would get, or what is wrong with the boxes. Returns $true when they are good.
function Update-PBBulkPreview {
                $bulk = $script:PBBulk
                if ($null -eq $bulk) { return $false }
                $request = Get-PBBulkRequest
                if ([string]::IsNullOrWhiteSpace($request.Count)) {
                                $bulk.PreviewText.Text = 'Type how many lines to add.'
                                $bulk.PreviewText.Foreground = Get-PBBrush '#64748B'
                                $bulk.AddButton.IsEnabled = $false
                                return $false
                }
                $made = New-BulkBoardLines -Day $bulk.Day -StartTime $request.StartTime -Count $request.Count -Every $request.Every -SortieType $request.SortieType -Settings $script:PB.Board.Settings
                if (@($made.Errors).Count -gt 0) {
                                $bulk.PreviewText.Text = (@($made.Errors) -join "`n")
                                $bulk.PreviewText.Foreground = Get-PBBrush '#B91C1C'
                                $bulk.AddButton.IsEnabled = $false
                                return $false
                }
                $times = @($made.Times)
                $shown = if ($times.Count -le 12) { $times -join ' ' } else { (@($times[0..7]) -join ' ') + ' ... ' + (@($times[($times.Count - 2)..($times.Count - 1)]) -join ' ') }
                $bulk.PreviewText.Text = (Format-PBBulkTimes -Times $times -Every $request.Every -SortieType $request.SortieType) + "`n" + $shown
                $bulk.PreviewText.Foreground = Get-PBBrush '#0F172A'
                $bulk.AddButton.IsEnabled = $true
                return $true
}

function Submit-PBBulkLinesDialog {
                $bulk = $script:PBBulk
                if (-not (Update-PBBulkPreview)) { return }
                $bulk.Result = Get-PBBulkRequest
                try { $bulk.Window.DialogResult = $true } catch { $bulk.Window.Close() }
}


# =================================================================================
#  UPDATE SYLLABI
#  Reads a training file (the day's flying schedule or the Big Board), works out which one
#  it is and brings the students' syllabi in line with it (SyllabusImport). What will change
#  is shown first; on Yes the board's syllabi change and are saved right away.
# =================================================================================

# The "Update Syllabi" button
function Invoke-PBUpdateSyllabi {
                $path = Show-PBOpenFileDialog -Title 'Update Syllabi: pick the flying schedule or the Big Board'
                if ([string]::IsNullOrWhiteSpace($path)) { return }
                $fileName = [System.IO.Path]::GetFileName($path)
                $data = $script:PB.Data
                $now = Get-Date

                Show-BusyOverlay "Reading $fileName..."
                try {
                                $read = Read-TrainingFile -Path $path
                                $plan = $null
                                if (-not $read.Error) {
                                                $plan = Get-SyllabusUpdatePlan -File $read -Students @($data.Students) -Instructors @($data.Instructors) -OutputPath ([string]$data.OutputPath) -SoftWeights $data.SoftWeights -Now $now
                                }
                }
                finally { Hide-BusyOverlay }
                if ($read.Error) {
                                [void](Show-PBMessage -Text $read.Error -Title 'Update Syllabi' -Buttons 'OK' -Icon 'Warning')
                                Set-PBStatus "Update Syllabi: $fileName was not used - nothing was changed." 'Warn'
                                return
                }

                $summary = Format-SyllabusUpdateSummary -File $read -Plan $plan
                Write-Host "`nUpdate Syllabi from $path" -ForegroundColor Cyan
                Write-Host (Format-SyllabusUpdateSummary -File $read -Plan $plan -MaxLines 100000 -MaxMissions 100000)
                if (@($plan.Changes).Count -eq 0) {
                                [void](Show-PBMessage -Text $summary -Title 'Update Syllabi' -Buttons 'OK' -Icon 'Information')
                                Set-PBStatus "Update Syllabi: every syllabus already matches $fileName - nothing was changed." 'Good'
                                return
                }
                $question = "$summary`n`nUpdate these syllabi? Each student's tab in OutputSchedule and their SimDebt in MasterData are saved right away."
                if ($script:PB.Mode -eq 'Review') { $question += "`n`nThe fill you are reviewing is thrown away (it was made from the old syllabi). Fill again afterwards." }
                if ((Show-PBMessage -Text $question -Title 'Update Syllabi' -Buttons 'YesNo' -Icon 'Question') -ne 'Yes') {
                                Set-PBStatus 'Update Syllabi cancelled - nothing was changed.'
                                return
                }

                # A fill made from the old syllabi would put them back on Export
                Exit-PBReview
                $memory = @(Invoke-SyllabusUpdate -Targets $plan.Targets -SoftWeights $data.SoftWeights -Now $now)
                $prompt = {
                                param($file)
                                $text = "$file is still open and FAIS could not close it (a cell may be in edit mode, or someone else has it open).`n`nClose it, then press OK to try again."
                                return ((Show-PBMessage -Text $text -Title 'File is open' -Buttons 'OKCancel' -Icon 'Warning') -eq 'OK')
                }
                Show-BusyOverlay 'Saving the syllabi...'
                try {
                                $saved = Save-SyllabusUpdate -Targets $plan.Targets -MemoryChanges $memory -OutputPath ([string]$data.OutputPath) -MasterDataPath ([string]$data.MasterDataPath) `
                                                -SoftWeights $data.SoftWeights -LockedPrompt $prompt -Now $now
                }
                finally { Hide-BusyOverlay }

                # Every hardcoded student is checked again against the new syllabi
                Clear-AssessmentCache
                $script:PB.Assess = @{}
                Update-PBAll
                Start-PBAssessQueue

                $students = @($plan.Changes | ForEach-Object { $_.Student } | Sort-Object -Unique).Count
                $done = "Updated $(@($plan.Changes).Count) event(s) for $students student(s) from $fileName"
                if ($saved.Saved) {
                                Set-PBStatus "$done and saved them to OutputSchedule and MasterData." 'Good'
                }
                else {
                                $text = "$done for this run only: they could not be saved ($($saved.Message)). Export writes them with the schedule."
                                [void](Show-PBMessage -Text $text -Title 'Update Syllabi' -Buttons 'OK' -Icon 'Warning')
                                Set-PBStatus $text 'Warn'
                }
}

# The Excel file picker for Update Syllabi (wrapped so the tests can answer it). Opens where
# the last file came from, else Downloads. Returns the file picked, or '' when cancelled.
function Show-PBOpenFileDialog {
                param ([string]$Title)

                $dlg = New-Object Microsoft.Win32.OpenFileDialog
                $dlg.Title = $Title
                $dlg.Filter = 'Excel workbooks (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All files (*.*)|*.*'
                $dlg.CheckFileExists = $true
                $dir = [string]$script:PB.ImportDir
                if ($dir -eq '' -or -not (Test-Path -LiteralPath $dir)) { $dir = Join-Path ([Environment]::GetFolderPath('UserProfile')) 'Downloads' }
                if (Test-Path -LiteralPath $dir) { $dlg.InitialDirectory = $dir }
                $owner = $script:PB.Window
                $picked = if ($null -ne $owner -and $owner.IsVisible) { $dlg.ShowDialog($owner) } else { $dlg.ShowDialog() }
                if ($picked -ne $true) { return '' }
                $script:PB.ImportDir = [System.IO.Path]::GetDirectoryName($dlg.FileName)
                return [string]$dlg.FileName
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
                Request-PBCommand @{ Cmd = 'AddLines'; Lines = $r.Lines; Links = $r.Links; Status = $msg }
}

function Update-PBTemplateList {
                $script:PB.C.TemplateCombo.ItemsSource = @(Get-BoardTemplateList -Dir $script:PB.Settings.TemplateDir)
}

# -Template: one from Get-BoardTemplateList (the File menu); else the one picked in the toolbar
function Invoke-PBLoadTemplate {
                param ($Template)

                $t = if ($null -ne $Template) { $Template } else { $script:PB.C.TemplateCombo.SelectedItem }
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

# -AddDate: the day to add (the Day menu); else the one in the toolbar's date box
function Invoke-PBAddDay {
                param ($AddDate)

                $date = if ($null -ne $AddDate) { $AddDate } else { $script:PB.C.AddDayPicker.SelectedDate }
                if ($null -eq $date) { Set-PBStatus 'Pick a date first.' 'Warn'; return }
                $day = ([datetime]$date).ToString('yyyy-MM-dd')
                Invoke-PBChange { [void](Add-BoardDay -Board $script:PB.Board -Day $day) }
                Reset-PBPlannedFromBoard
                Select-PBDay $day
                $script:PB.C.AddDayPicker.SelectedDate = (Get-NextFlyingDay -From ([datetime]$date))
                # A new day the Schedule sheet has lines for comes with them
                $msg = "Added $(Format-BoardDayLabel $day)."
                $st = Get-BoardSheetState -Board $script:PB.Board -SheetRows @(Get-PBSheetRows) -Day $day
                if ($st.State -eq 'EmptyDay') {
                                $res = Invoke-PBLoadDay -Source 'Sheet' -SourceDay $day -TargetDay $day -Replace -KeepStudents -Quiet
                                if ($null -ne $res) { $msg = "Added $(Format-BoardDayLabel $day) with its $($res.Added) line(s) from MasterData's Schedule sheet (Ctrl+Z takes them off)." }
                }
                Set-PBStatus $msg 'Good'
}

# A newly added earlier day: events planned on/after it are plans, not flown events
function Reset-PBPlannedFromBoard {
                if ($script:PB.Board.Days.Count -eq 0) { return }
                $first = [datetime]::ParseExact($script:PB.Board.Days[0], 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
                $changed = @(Reset-PlannedCompletions -Students $script:PB.Data.Students -FromDate $first -SoftWeights $script:PB.Data.SoftWeights -PrevSchedule @($script:PB.Data.PrevSchedule))
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

# "Fill this day" / "Fill all days": runs the scheduler on a copy of the students and shows
# the result on the board for review. Nothing is written until Export.
# -Scope Day fills only the day on screen (and later exports only it). Days already
# exported that are not being filled go in as fixed context, so nobody is
# double-booked and their events stay in order. -Refill repeats the last fill's days.
function Invoke-PBFill {
                param ([switch]$Refill, [ValidateSet('All', 'Day')] [string]$Scope = 'All')

                $board = $script:PB.Board
                if ($board.Lines.Count -eq 0) { Set-PBStatus 'There are no lines to fill. Type some in the Add lines box, or use Load a day...' 'Warn'; return }
                $days = @(Get-PBFillScope -Scope $Scope)
                if ($Refill -and @($script:PB.FillDays).Count -gt 0) {
                                $again = @($script:PB.FillDays | Where-Object { $board.Days.Contains([string]$_) })
                                if ($again.Count -gt 0) { $days = $again }
                }
                if (@($board.Lines | Where-Object { $days -contains $_.Day }).Count -eq 0) { Set-PBStatus "$(Format-PBDays $days) has no lines to fill." 'Warn'; return }
                $warnings = @(Test-BoardReadyToFill -Board $board -Students $script:PB.Data.Students -Instructors $script:PB.Data.Instructors -Days $days)
                if ($warnings.Count -gt 0) {
                                $shown = @($warnings | Select-Object -First 15)
                                $more = if ($warnings.Count -gt 15) { "`n...and $($warnings.Count - 15) more" } else { '' }
                                $ans = [System.Windows.MessageBox]::Show(("Before filling:`n`n- " + ($shown -join "`n- ") + $more + "`n`nFill anyway?"),
                                                "Fill $(Format-PBDays $days)", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
                                if ($ans -ne [System.Windows.MessageBoxResult]::Yes) { return }
                }

                Exit-PBReview
                Show-BusyOverlay 'Filling the schedule...'
                $script:PB.Filling = $true
                try {
                                $rows = @(ConvertTo-ScheduleRows -Board $board -Days $days)
                                $context = @(Get-PublishedContextRows -Rows @($script:PB.Data.PrevSchedule) -FromDay ([string]$board.Days[0]) -ExcludeDays $days)
                                $rows = @($rows) + @($context)
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
                finally { $script:PB.Filling = $false; Hide-BusyOverlay }

                $split = Split-FillResult -FillResult $result -Days $days
                $script:PB.FillDays = @($split.Days)
                $script:PB.FillResult = $split
                $script:PB.Overlay = Get-FillOverlay -FillResult $split
                Set-PBMode 'Review'
                Update-PBAll
                $st = $split.Stats
                $kind = if ($st.Overrides -gt 0 -or $st.ForSale -gt 0) { 'Warn' } else { 'Good' }
                $msg = "Filled {0}: {1} by the scheduler, {2} hardcoded, {3} wing, {4} return legs, {5} for sale, {6} override flag(s). Check the board, then Export." -f (Format-PBDays $days), $st.Filled, $st.Pinned, $st.Wing, $st.Return, $st.ForSale, $st.Overrides
                $ctxOv = @($split.ContextOverrides)
                if ($ctxOv.Count -gt 0) {
                                $kind = 'Warn'
                                $what = @($ctxOv | Select-Object -First 2 | ForEach-Object { "{0} {1}: {2}" -f (Format-EventLabel -Start $_.ETD -WithDay), $_.Student, (@($_.Violations) -join '; ') })
                                $msg += " HEADS UP: $($ctxOv.Count) line(s) already exported on another day now break a rule ($($what -join ' | ')). Load and fill that day too."
                }
                Set-PBStatus $msg $kind
}

function Invoke-PBExport {
                if ($script:PB.Mode -ne 'Review' -or $null -eq $script:PB.FillResult) { Set-PBStatus 'Fill the lines first.' 'Warn'; return }
                Save-PBBoard
                $script:PB.Result = [PSCustomObject]@{ Action = 'Export'; Board = $script:PB.Board; FillResult = $script:PB.FillResult; Days = @($script:PB.FillDays) }
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
  Add Bulk Lines          asks for a start time, how many lines and the minutes between them
                          (3 by default), e.g. 0730, 10 lines -> 0730 0733 ... 0757
Templates save a day's times; "Copy previous day's lines" copies the last flying day.

ONE DAY AT A TIME
  Load a day...           Loads one day from MasterData's Schedule sheet, or from the schedule
                          you last exported (to change a day already sent out: everyone stays
                          where they were, so a fill only changes the lines you open up).
  MasterData is watched:  save the Schedule sheet in Excel and the board follows: a day you have
                          not edited since it was loaded (or that is still empty) updates by
                          itself; for a day you have edited, the blue banner offers the change.
  Fill this day           Fills and exports just the day on screen (Ctrl+Shift+Enter). Days
                          already exported stay exactly as they were.

PUCKS (drag and drop, or click and place)
  Drag an IP puck onto a line's IP slot, a student puck onto a line's student slot.
  Or click a puck to pick it up (thick outline, dark banner), then click the slot to put it
  there (hold Ctrl as you click to keep holding it for another line). Esc lets go.
  Double-click a puck to put it straight on the selected line.
  Drag between lines to move (hold Ctrl to copy). Drop back on a pool to take it off.
  While dragging or holding a student, slots turn green (ok), amber (needs a look), red
  (breaks a rule). Rule breaks are allowed - the line is flagged OVERRIDE on the output.

LINE DETAILS (the panel under the days)
  Click anywhere on a line to select it (blue outline). The panel shows its student against
  every HARD RULE, WARNING and SOFT RULE (the + / - score), and ranks every student for the
  line: fits first, then warnings, then rule breaks, best score first. Click a student in the
  list to see their check; double-click (or Put on this line) to hardcode them there.
  Drag the bar above the panel to resize it; Hide (or View menu) to close it.

MISSIONS
  A hardcoded student's mission dropdown lists (auto), the events they can fly on the line,
  the next events they are NOT opted for yet (marked), and "Type a mission..." for anything.
  A mission that breaks a rule (not opted - e.g. a sim before the dollar ride - not in the
  syllabus, or the wrong line type) is explained first, with the scheduler's pick offered:
  Yes = use the recommended mission, No = keep yours (flagged OVERRIDE), Cancel = no change.
  A line typed with sp: and m: goes on as typed and its rule check pops up.

MARKERS
  The corner triangle on a hardcoded student: green clean, amber warnings, red rule broken,
  grey still checking. Hover it for the + / - reasons. Right-click for more.

TIMES
  Click an ETD, type a new time, Enter. Up/Down = 5 min, Shift = 30 min, Esc cancels.

MENUS
  Every action is in the menu bar; the Line menu acts on the selected line (the same as its
  right-click menu). Help > What the colors mean (or the Colors button) explains every color.

KEYS
  Ctrl+L quick entry   Ctrl+Z undo   Delete clears the selected slot   Esc lets go of a puck
  Ctrl+D duplicate the selected line   Ctrl+Delete delete it   Ctrl+PgUp / PgDn change day
  F5 re-check everything   Ctrl+Enter fill all days   Ctrl+Shift+Enter fill this day

FILL
  "Fill <day>" schedules the open lines on the day on screen; "Fill all N days" (shown when
  the board has more than one day) schedules every day. Review the result, then Export
  (writes OutputSchedule.xlsx), Re-fill, or go Back to edit. A window then says what was
  written, or why nothing was, with buttons to open the files.
  Excel files FAIS writes are saved and closed for you first, and opened again afterwards.

UPDATE SYLLABI
  Pick a training file; FAIS works out which one it is:
  - the day's flying schedule (Flying and Sims tabs): a student's line marked CM or EXEC
    marks its mission complete with the line's IP and ETD; INC or CX makes sure it is not.
  - the Big Board: an entry in an event's columns marks it complete with its IP and date;
    no entry makes sure it is not complete (an event FAIS has planned for later stays).
  Only events in a student's syllabus change. You see what will change first; Yes saves
  each student's tab in OutputSchedule and their SimDebt in MasterData right away.
"@
                [void][System.Windows.MessageBox]::Show($text, 'FAIS Puck Board help', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
}


# =================================================================================
#  ONE DAY AT A TIME
#  "Load a day..." puts one day on the board from MasterData's Schedule sheet or from the
#  last exported schedule. The sheet is re-read whenever MasterData is saved, and the
#  banner offers any day the board is behind on. "Fill this day" fills and exports one day.
# =================================================================================

function Select-PBDay {
                param ([string]$Day)

                foreach ($t in $script:PB.C.DayTabs.Items) { if ([string]$t.Tag -eq $Day) { $script:PB.C.DayTabs.SelectedItem = $t } }
}

# "Tue 20 Oct" for one day, "Tue 20 Oct - Thu 22 Oct (3 days)" for several
function Format-PBDays {
                param ([string[]]$Days = @())

                $list = @($Days | Where-Object { $_ } | Sort-Object -Unique)
                if ($list.Count -eq 0) { return 'no days' }
                if ($list.Count -eq 1) { return (Format-BoardDayLabel $list[0]) }
                return ("{0} - {1} ({2} days)" -f (Format-BoardDayLabel $list[0]), (Format-BoardDayLabel $list[-1]), $list.Count)
}

# The days a fill covers: -Scope Day is the day on screen ("Fill this day"), All every board day
function Get-PBFillScope {
                param ([ValidateSet('All', 'Day')] [string]$Scope = 'All')

                $active = Get-PBActiveDay
                if ($Scope -eq 'Day' -and $active) { return @([string]$active) }
                return @($script:PB.Board.Days)
}

# The two fill buttons: "Fill Tue 20 Oct" (the day on screen) and "Fill all 3 days" (only when
# the board has more than one day). Both are hidden while a fill is being reviewed.
function Update-PBFillScope {
                $C = $script:PB.C
                $active = Get-PBActiveDay
                $n = $script:PB.Board.Days.Count
                $review = ($script:PB.Mode -eq 'Review')
                $C.FillDayButton.Content = if ($active) { "Fill $(Format-BoardDayLabel $active)" } else { 'Fill this day' }
                $C.FillAllButton.Content = "Fill all $n days"
                $C.FillDayButton.Visibility = if ($review) { 'Collapsed' } else { 'Visible' }
                $C.FillAllButton.Visibility = if ($review -or $n -lt 2) { 'Collapsed' } else { 'Visible' }
}

function Get-PBFileStamp {
                param ([string]$Path)

                try {
                                $f = Get-Item -LiteralPath $Path -ErrorAction Stop
                                return ('{0}|{1}' -f $f.LastWriteTimeUtc.Ticks, $f.Length)
                } catch { return '' }
}

# MasterData's Schedule sheet rows as last read, with the rows that have a time but no
# date put on the day being built (the board's first day)
function Get-PBSheetRows {
                $sheet = $script:PB.Data.ScheduleSheet
                if ($null -eq $sheet) { return @() }
                $first = if ($script:PB.Board.Days.Count -gt 0) { [string]$script:PB.Board.Days[0] } else { '' }
                return @(Resolve-SheetRowDays -Rows @($sheet.Rows) -DefaultDay $first)
}

# Reads MasterData's Schedule sheet again. -SaveFirst saves MasterData in Excel first, so
# edits not saved yet are read too. Returns the read; a good one replaces Data.ScheduleSheet.
function Update-PBSheetRead {
                param ([switch]$SaveFirst)

                $path = [string]$script:PB.Data.MasterDataPath
                if ([string]::IsNullOrWhiteSpace($path)) { return $script:PB.Data.ScheduleSheet }
                if ($SaveFirst) {
                                $saved = Save-FaisWorkbook -Path $path
                                if ($saved.Message) { Write-PBLog $saved.Message }
                }
                $read = Read-ScheduleSheet -Path $path -ReturnType ([string]$script:PB.Settings.ReturnSortieType) -SimInstructor ([string]$script:PB.Settings.SimInstructor)
                if (-not $read.Error) { $script:PB.Data.ScheduleSheet = $read }
                $script:PB.SheetStamp = Get-PBFileStamp $path
                return $read
}

# Reads the last exported MasterSchedule again (saving OutputSchedule in Excel first)
function Update-PBPublishedRead {
                $path = [string]$script:PB.Data.OutputPath
                if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -LiteralPath $path)) { return }
                $saved = Save-FaisWorkbook -Path $path
                if ($saved.Message) { Write-PBLog $saved.Message }
                $script:PB.Data.PrevSchedule = @(DownloadPrevSchedule -Students $script:PB.Data.Students -Instructors $script:PB.Data.Instructors -excelOutputPath $path)
}

# Loads the sheet onto the board days that can take it without losing anything: days still
# empty, and days nobody has edited since they were loaded. Returns 'Tue 20 Oct (14 line(s))'
# for each day loaded.
function Invoke-PBSheetAutoUpdate {
                $done = New-Object System.Collections.Generic.List[string]
                $rows = @(Get-PBSheetRows)
                if ($rows.Count -eq 0) { return @() }
                foreach ($day in @($script:PB.Board.Days)) {
                                $st = Get-BoardSheetState -Board $script:PB.Board -SheetRows $rows -Day $day
                                if ($st.State -ne 'EmptyDay' -and $st.State -ne 'Outdated') { continue }
                                $res = Invoke-PBLoadDay -Source 'Sheet' -SourceDay $day -TargetDay $day -Replace -KeepStudents -Quiet -StayOnDay
                                if ($null -ne $res) { $done.Add("$(Format-BoardDayLabel $day) ($($res.Added) line(s))") }
                }
                return $done.ToArray()
}

# When the board opens: re-read the sheet if MasterData was saved after it was read, bring the
# days the board can safely take up to date, and let the banner show anything else.
# Returns a short message for the status line ('' when nothing was loaded).
function Invoke-PBSheetStartup {
                $path = [string]$script:PB.Data.MasterDataPath
                if ($path -and (Test-Path -LiteralPath $path)) {
                                $sheet = $script:PB.Data.ScheduleSheet
                                $readAt = if ($null -ne $sheet -and $null -ne $sheet.ReadAt) { [datetime]$sheet.ReadAt } else { [datetime]::MinValue }
                                if ((Get-Item -LiteralPath $path).LastWriteTime -gt $readAt) { [void](Update-PBSheetRead) }
                                else { $script:PB.SheetStamp = Get-PBFileStamp $path }
                }
                $loaded = @(Invoke-PBSheetAutoUpdate)
                Update-PBSheetBanner
                if ($loaded.Count -eq 0) { return '' }
                return "Loaded $($loaded -join ', ') from MasterData's Schedule sheet (Ctrl+Z takes them off)."
}

# Every couple of seconds: once MasterData has been saved since it was last read (and the file
# has stopped changing), read the Schedule sheet again, update the days that can safely take
# it (in edit mode, and not while a row is being typed in), and refresh the banner
function Invoke-PBSheetTick {
                if ($script:PB.Busy -or $script:PB.Draining -or $script:PB.Filling) { return }
                if ($null -ne $script:PB.Drag -and $script:PB.Drag.Active) { return }
                $path = [string]$script:PB.Data.MasterDataPath
                if ([string]::IsNullOrWhiteSpace($path)) { $script:PB.SheetTimer.Stop(); return }
                $stamp = Get-PBFileStamp $path
                if ($stamp -eq '' -or $stamp -eq $script:PB.SheetStamp) { $script:PB.SheetPending = '' }
                elseif ($stamp -ne $script:PB.SheetPending) { $script:PB.SheetPending = $stamp; return }
                else {
                                $script:PB.SheetPending = ''
                                $read = Update-PBSheetRead
                                if ($read.Error) { Write-PBLog "Schedule sheet re-read failed: $($read.Error)"; return }
                                $script:PB.SheetAutoPending = $true
                                Update-PBSheetBanner
                }
                if (-not $script:PB.SheetAutoPending -or $script:PB.Mode -ne 'Edit' -or $null -ne (Get-PBFocusState)) { return }
                $script:PB.SheetAutoPending = $false
                $updated = @(Invoke-PBSheetAutoUpdate)
                if ($updated.Count -gt 0) { Set-PBStatus "MasterData was saved: updated $($updated -join ', ') from the Schedule sheet. Ctrl+Z undoes." 'Good' }
                elseif ($null -ne $script:PB.SheetBannerInfo) { Set-PBStatus "MasterData was saved. $($script:PB.SheetBannerInfo.Text)" 'Warn' }
}

# Shows or hides the banner about MasterData's Schedule sheet
function Update-PBSheetBanner {
                $C = $script:PB.C
                $info = $null
                try {
                                $rows = @(Get-PBSheetRows)
                                if ($rows.Count -gt 0) {
                                                $info = Get-SheetBannerInfo -Board $script:PB.Board -SheetRows $rows -ActiveDay ([string](Get-PBActiveDay)) -Ignore $script:PB.SheetIgnore.ToArray()
                                }
                } catch { Write-PBLog "Sheet banner: $($_.Exception.Message)" }
                $script:PB.SheetBannerInfo = $info
                if ($null -eq $info) { $C.SheetBanner.Visibility = 'Collapsed'; return }
                $C.SheetBannerText.Text = $info.Text
                $C.SheetBannerLoadButton.Content = $info.ActionText
                $C.SheetBannerIgnoreButton.Content = if ($info.State -eq 'OtherDay') { 'Not now' } else { "Keep the board's lines" }
                $C.SheetBanner.Visibility = 'Visible'
}

function Invoke-PBSheetBannerLoad {
                $info = $script:PB.SheetBannerInfo
                if ($null -eq $info) { return }
                [void](Invoke-PBLoadDay -Source 'Sheet' -SourceDay $info.Day -TargetDay $info.Day -Replace -KeepStudents)
}

function Invoke-PBSheetBannerIgnore {
                $info = $script:PB.SheetBannerInfo
                if ($null -eq $info) { return }
                if ($info.State -eq 'OtherDay') { $script:PB.SheetIgnore.Add("$($info.Day)|$($info.Fingerprint)") }
                else {
                                Set-BoardSheetSync -Board $script:PB.Board -Day $info.Day -Entry (New-SheetSyncEntry -Fingerprint $info.Fingerprint -Action 'Ignored')
                                Request-PBAutosave
                }
                Update-PBSheetBanner
                Set-PBStatus "Kept the board's lines for $(Format-BoardDayLabel $info.Day). You will be asked again if the Schedule sheet changes."
}

# Puts one day's lines on the board (one undo step) from MasterData's Schedule sheet
# (-Source Sheet) or from the last exported schedule (-Source Exported).
# -SourceDay is the day taken ('' = sheet rows with no date), -TargetDay where they go.
# Returns {Added; Removed; Students; Day}, or $null when there was nothing to load.
function Invoke-PBLoadDay {
                param (
                                [ValidateSet('Sheet', 'Exported')] [string]$Source = 'Sheet',
                                [AllowEmptyString()] [string]$SourceDay = '',
                                [Parameter(Mandatory)] [string]$TargetDay,
                                [switch]$Replace,
                                [switch]$KeepStudents,
                                [switch]$Quiet,
                                [switch]$StayOnDay
                )

                $board = $script:PB.Board
                $sync = $null
                if ($Source -eq 'Sheet') {
                                $rows = @(Get-PBSheetRows)
                                $lines = @(ConvertFrom-SheetRows -Rows $rows -Day $TargetDay -FromDay $SourceDay -KeepStudents:$KeepStudents -KeepNotes)
                                if ($SourceDay -eq $TargetDay) { $sync = New-SheetSyncEntry -Fingerprint (Get-ScheduleDayFingerprint -Rows $rows -Day $TargetDay) }
                                $from = "MasterData's Schedule sheet"
                }
                else {
                                $moveTo = if ($TargetDay -ne $SourceDay) { $TargetDay } else { '' }
                                $lines = @(ConvertFrom-PublishedRows -Rows @($script:PB.Data.PrevSchedule) -Day $SourceDay -TargetDay $moveTo)
                                if (-not $KeepStudents) { foreach ($l in $lines) { $l.Student = ''; $l.Mission = ''; $l.WingOfLineId = '' } }
                                $from = 'the last exported schedule'
                }
                if ($lines.Count -eq 0) {
                                if (-not $Quiet) { Set-PBStatus "$from has no lines for $(Format-BoardDayLabel $SourceDay)." 'Warn' }
                                return $null
                }

                $newDay = -not $board.Days.Contains($TargetDay)
                $holder = @{ Res = $null }
                Invoke-PBChange { $holder.Res = Import-BoardDayLines -Board $board -Day $TargetDay -Lines $lines -Replace:$Replace -Sync $sync }
                if ($newDay) { Reset-PBPlannedFromBoard }
                if (-not $StayOnDay) { Select-PBDay $TargetDay }
                $res = $holder.Res
                $pinned = @($lines | Where-Object { $_.Student }).Count
                if (-not $Quiet) {
                                $msg = "Loaded $($res.Added) line(s) onto $(Format-BoardDayLabel $TargetDay) from $from"
                                if ($pinned -gt 0) { $msg += ", $pinned with students hardcoded" }
                                if ($res.Removed -gt 0) { $msg += ", replacing $($res.Removed) line(s)" }
                                $msg += '. Ctrl+Z undoes.'
                                $kind = 'Good'
                                if ($Source -eq 'Sheet' -and $null -ne $script:PB.Data.ScheduleSheet -and @($script:PB.Data.ScheduleSheet.Skipped).Count -gt 0) {
                                                $msg += " $(@($script:PB.Data.ScheduleSheet.Skipped).Count) sheet row(s) could not be read (Load a day... lists them)."
                                                $kind = 'Warn'
                                }
                                Set-PBStatus $msg $kind
                }
                return [PSCustomObject]@{ Added = $res.Added; Removed = $res.Removed; Students = $pinned; Day = $TargetDay }
}

# "Load a day...": pick where from and which day, then load it
function Invoke-PBLoadDayDialog {
                $choice = Show-PBLoadDayDialog
                if ($null -eq $choice) { return }
                [void](Invoke-PBLoadDay -Source $choice.Source -SourceDay $choice.SourceDay -TargetDay $choice.TargetDay -Replace:$choice.Replace -KeepStudents:$choice.KeepStudents)
}

function Show-PBLoadDayDialog {
                $win = Initialize-PBLoadDayDialog
                try { [void]$win.ShowDialog() } catch { Write-PBLog "Load a day dialog: $($_.Exception.Message)" }
                $result = $script:PBDlg.Result
                $script:PBDlg = $null
                return $result
}

# Builds the Load a day window without showing it. Its state lives in $script:PBDlg.
function Initialize-PBLoadDayDialog {
                [xml]$xaml = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'LoadDayDialog.xaml') -Raw -Encoding UTF8
                $win = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
                $dlg = @{ Window = $win; Result = $null }
                foreach ($n in @('SourceSheetRadio', 'SourceExportRadio', 'SourceInfoText', 'DayList', 'TargetDayPicker', 'TargetInfoText',
                                'ReplaceRadio', 'AddRadio', 'KeepStudentsCheck', 'SkippedText', 'RefreshButton', 'LoadButton', 'CancelButton')) {
                                $c = $win.FindName($n)
                                if ($null -eq $c) { throw "LoadDayDialog.xaml is missing the element '$n'" }
                                $dlg[$n] = $c
                }
                $script:PBDlg = $dlg
                try { if ($script:PB.Window.IsVisible) { $win.Owner = $script:PB.Window } } catch { }
                $dlg.SourceSheetRadio.Add_Checked({ Invoke-PBSafe { Update-PBDialogList } })
                $dlg.SourceExportRadio.Add_Checked({ Invoke-PBSafe { Update-PBDialogList } })
                $dlg.DayList.Add_SelectionChanged({ Invoke-PBSafe { Update-PBDialogTarget } })
                $dlg.TargetDayPicker.Add_SelectedDateChanged({ Invoke-PBSafe { Update-PBDialogTargetInfo } })
                $dlg.RefreshButton.Add_Click({ Invoke-PBSafe { Update-PBDialogList -Reread } })
                $dlg.LoadButton.Add_Click({ Invoke-PBSafe { Submit-PBDialog } })
                # The file may have changed since the board opened: read it fresh (saving it in Excel first)
                Update-PBDialogList -Reread
                return $win
}

function New-PBDialogItem {
                param ([string]$Day, [int]$Count, [int]$Students)

                $label = if ($Day -eq '') { 'No date' } else { ([datetime]::ParseExact($Day, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)).ToString('ddd dd MMM yyyy') }
                $n = if ($Day -ne '') { @($script:PB.Board.Lines | Where-Object { $_.Day -eq $Day }).Count } else { 0 }
                $where = if ($Day -ne '' -and $script:PB.Board.Days.Contains($Day)) { "on the board now: $n" } else { 'not on the board' }
                return [PSCustomObject]@{ Day = $Day; Text = ('{0,-16} {1,3} line(s), {2,3} with students   ({3})' -f $label, $Count, $Students, $where) }
}

function Get-PBSheetInfoText {
                param ($Sheet)

                $path = [string]$script:PB.Data.MasterDataPath
                if ($null -eq $Sheet) { return "MasterData's Schedule sheet has not been read ($path)." }
                if ($Sheet.Error) { return [string]$Sheet.Error }
                if (-not $Sheet.Found) { return "MasterData has no 'Schedule' sheet ($path). Add one with the columns ETD, Instructor, SortieType, Student, Mission, Callsign, Num, Notes (headings in any order)." }
                $all = @($Sheet.Rows)
                if ($all.Count -eq 0) { return "MasterData's Schedule sheet has no lines ($path)." }
                $text = "Read at $($Sheet.ReadAt.ToString('HHmm')) from $path ($($Sheet.Layout))."
                $noDate = @($all | Where-Object { $_.TimeOnly }).Count
                if ($noDate -gt 0 -and $script:PB.Board.Days.Count -gt 0) { $text += " $noDate line(s) have a time but no date: they count as $(Format-BoardDayLabel $script:PB.Board.Days[0]) (the first day on the board)." }
                return $text
}

# Fills the day list for the chosen source. -Reread reads the file again first.
function Update-PBDialogList {
                param ([switch]$Reread)

                $dlg = $script:PBDlg
                if ($null -eq $dlg) { return }
                $items = New-Object System.Collections.Generic.List[object]
                if ([bool]$dlg.SourceExportRadio.IsChecked) {
                                if ($Reread) { Update-PBPublishedRead }
                                $pub = @($script:PB.Data.PrevSchedule)
                                foreach ($day in @(Get-PublishedDays -Rows $pub)) {
                                                $dayRows = @(Get-PublishedDayRows -Rows $pub -Day $day)
                                                $withSp = @($dayRows | Where-Object { -not (Test-IsSentinelStudent $_.Student) -and -not (Test-IsReturnSortieType ([string]$_.SortieType)) }).Count
                                                $items.Add((New-PBDialogItem -Day $day -Count $dayRows.Count -Students $withSp))
                                }
                                $path = [string]$script:PB.Data.OutputPath
                                $dlg.SourceInfoText.Text = if ($items.Count -gt 0) { "From $path. Use this to change a day already sent out: everyone stays where they were (hardcoded), so a fill only changes the lines you open up." } else { "No exported schedule was found ($path)." }
                                $dlg.KeepStudentsCheck.Content = 'Keep the students and missions as exported (recommended)'
                                $dlg.SkippedText.Text = ''
                }
                else {
                                if ($Reread) { [void](Update-PBSheetRead -SaveFirst) }
                                $sheet = $script:PB.Data.ScheduleSheet
                                foreach ($g in @(Get-ScheduleSheetDays -Rows @(Get-PBSheetRows))) { $items.Add((New-PBDialogItem -Day $g.Day -Count $g.Count -Students $g.Students)) }
                                $dlg.SourceInfoText.Text = Get-PBSheetInfoText $sheet
                                $dlg.KeepStudentsCheck.Content = 'Keep the students and missions on the sheet (they go on as hardcoded)'
                                $skipped = @()
                                if ($null -ne $sheet) { $skipped = @($sheet.Skipped) }
                                $dlg.SkippedText.Text = ''
                                if ($skipped.Count -gt 0) {
                                                $more = if ($skipped.Count -gt 6) { "; and $($skipped.Count - 6) more" } else { '' }
                                                $dlg.SkippedText.Text = "Rows that could not be read: " + (@($skipped | Select-Object -First 6) -join '; ') + $more
                                }
                }
                $dlg.DayList.ItemsSource = $items.ToArray()
                # Start on the day on screen when the list has it, else the first day listed
                $active = [string](Get-PBActiveDay)
                $pick = $null
                foreach ($it in $items) { if ($it.Day -eq $active) { $pick = $it } }
                if ($null -eq $pick -and $items.Count -gt 0) { $pick = $items[0] }
                $dlg.DayList.SelectedItem = $pick
                Update-PBDialogTarget
}

function Update-PBDialogTarget {
                $dlg = $script:PBDlg
                if ($null -eq $dlg) { return }
                $item = $dlg.DayList.SelectedItem
                $day = if ($null -ne $item -and $item.Day) { [string]$item.Day } else { [string](Get-PBActiveDay) }
                if ($day) { $dlg.TargetDayPicker.SelectedDate = [datetime]::ParseExact($day, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture) }
                Update-PBDialogTargetInfo
}

function Update-PBDialogTargetInfo {
                $dlg = $script:PBDlg
                if ($null -eq $dlg) { return }
                $date = $dlg.TargetDayPicker.SelectedDate
                if ($null -eq $date) { $dlg.TargetInfoText.Text = ''; return }
                $day = ([datetime]$date).ToString('yyyy-MM-dd')
                $n = @($script:PB.Board.Lines | Where-Object { $_.Day -eq $day }).Count
                $dlg.TargetInfoText.Text = if ($n -gt 0) { "has $n line(s) on the board now" } elseif ($script:PB.Board.Days.Contains($day)) { 'on the board, no lines yet' } else { 'adds this day to the board' }
}

function Submit-PBDialog {
                $dlg = $script:PBDlg
                if ($null -eq $dlg) { return }
                $item = $dlg.DayList.SelectedItem
                if ($null -eq $item) {
                                [void][System.Windows.MessageBox]::Show('Pick a day in the list first.', 'Load a day', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                                return
                }
                $date = $dlg.TargetDayPicker.SelectedDate
                if ($null -eq $date) {
                                [void][System.Windows.MessageBox]::Show('Pick the day to put the lines on.', 'Load a day', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                                return
                }
                $dlg.Result = [PSCustomObject]@{
                                Source       = $(if ([bool]$dlg.SourceExportRadio.IsChecked) { 'Exported' } else { 'Sheet' })
                                SourceDay    = [string]$item.Day
                                TargetDay    = ([datetime]$date).ToString('yyyy-MM-dd')
                                Replace      = [bool]$dlg.ReplaceRadio.IsChecked
                                KeepStudents = [bool]$dlg.KeepStudentsCheck.IsChecked
                }
                try { $dlg.Window.DialogResult = $true } catch { $dlg.Window.Close() }
}


# =================================================================================
#  EXPORT RESULT
#  After Export (FAIS.ps1) a window says what was written, or why nothing was, with
#  buttons that open the files. The same text still goes to the console.
# =================================================================================

# What the export result window says: {Ok; Title; Subtitle; Lines = @({Kind; Text})}.
# Kind: Head, Text, Path, Good, Warn, Hard.
#   -UnsatChanges: Get-UnsatFlagUpdates from before the export (name -> @{ UnsatSim; UnsatFlight })
#   -Warnings: the export's warning messages (why it failed, or names MasterData didn't have)
function Get-FaisExportSummary {
                param (
                                [bool]$Ok,
                                $FillResult,
                                [string[]]$Days = @(),
                                [string]$OutputPath = '',
                                [string]$MasterDataPath = '',
                                [string]$AtisPath = '',
                                [hashtable]$UnsatChanges = @{},
                                [string[]]$Warnings = @()
                )

                $lines = New-Object System.Collections.Generic.List[object]
                $add = { param([string]$Kind, [string]$Text) $lines.Add([PSCustomObject]@{ Kind = $Kind; Text = $Text }) }
                $dayText = Format-PBDays $Days
                $warnText = @($Warnings | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)

                if (-not $Ok) {
                                & $add 'Head' 'What went wrong'
                                if ($warnText.Count -eq 0) { & $add 'Hard' 'The files could not be written. The PowerShell window has the details.' }
                                foreach ($w in $warnText) { & $add 'Hard' $w }
                                & $add 'Head' 'What to do'
                                & $add 'Text' 'Close OutputSchedule, MasterData and ATIS if they are open somewhere FAIS cannot reach (another computer, or a cell still being edited), then run FAIS again and Export. Your board is saved, so the lines are still there.'
                                return [PSCustomObject]@{
                                                Ok       = $false
                                                Title    = 'Export failed: nothing was written'
                                                Subtitle = "OutputSchedule, MasterData and ATIS are exactly as they were ($dayText was not exported)."
                                                Lines    = $lines.ToArray()
                                }
                }

                $st = $FillResult.Stats
                & $add 'Head' 'What was scheduled'
                & $add 'Text' ("{0} lines: {1} hardcoded, {2} filled by the scheduler, {3} wingmen, {4} return legs, {5} direct support, {6} for sale" -f `
                                $st.Lines, $st.Pinned, $st.Filled, $st.Wing, $st.Return, $st.DS, $st.ForSale)
                if ([int]$st.ForSale -gt 0) { & $add 'Warn' "$($st.ForSale) line(s) are For Sale: no student could fly them." }

                & $add 'Head' 'Files written'
                & $add 'Text' "OutputSchedule: the MasterSchedule tab for $dayText (other days already exported are kept), and each student's syllabus tab"
                & $add 'Path' $OutputPath
                $unsatCount = @($UnsatChanges.Keys).Count
                & $add 'Text' $(if ($unsatCount -gt 0) { "MasterData: SimDebt, and the unsat flags of $unsatCount student(s)" } else { 'MasterData: SimDebt' })
                & $add 'Path' $MasterDataPath
                if ($AtisPath) {
                                & $add 'Text' "ATIS: one tab per day for $dayText, flights then sims"
                                & $add 'Path' $AtisPath
                }

                if ($unsatCount -gt 0) {
                                & $add 'Head' 'Unsats'
                                foreach ($name in @($UnsatChanges.Keys | Sort-Object)) {
                                                $flags = $UnsatChanges[$name]
                                                foreach ($key in @($flags.Keys | Sort-Object)) {
                                                                if ($flags[$key]) { & $add 'Warn' "$key set for $name" }
                                                                else { & $add 'Good' "$key cleared for $name (cleaned up on this schedule)" }
                                                }
                                }
                }

                $overrides = @($FillResult.Overrides)
                if ($overrides.Count -gt 0) {
                                & $add 'Head' 'Hardcoded against a rule (marked OVERRIDE on the output)'
                                foreach ($o in $overrides) { & $add 'Hard' ("{0} {1}: {2}" -f (Format-EtdShort $o.ETD), $o.Student, (@($o.Violations) -join '; ')) }
                }
                if ($warnText.Count -gt 0) {
                                & $add 'Head' 'Heads up'
                                foreach ($w in $warnText) { & $add 'Warn' $w }
                }
                return [PSCustomObject]@{
                                Ok       = $true
                                Title    = "Schedule exported: $dayText"
                                Subtitle = 'Everything below was written. Run FAIS again to change a day or build the next one.'
                                Lines    = $lines.ToArray()
                }
}

# Shows the export result (a Get-FaisExportSummary) until it is closed
function Show-FaisExportResult {
                param ([Parameter(Mandatory)] $Summary, [string]$OutputPath = '', [string]$AtisPath = '')

                $win = Initialize-FaisExportResultWindow -Summary $Summary -OutputPath $OutputPath -AtisPath $AtisPath
                [void]$win.ShowDialog()
}

# Builds the export result window without showing it
function Initialize-FaisExportResultWindow {
                param ([Parameter(Mandatory)] $Summary, [string]$OutputPath = '', [string]$AtisPath = '')

                Initialize-PuckBoardAssemblies
                [xml]$xaml = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'ExportResultDialog.xaml') -Raw -Encoding UTF8
                $win = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
                $ui = @{}
                foreach ($part in @('ResultBadge', 'ResultIcon', 'ResultTitle', 'ResultSubtitle', 'ResultLines', 'OpenOutputButton', 'OpenAtisButton', 'CloseResultButton')) {
                                $found = $win.FindName($part)
                                if ($null -eq $found) { throw "ExportResultDialog.xaml is missing the element '$part'" }
                                $ui[$part] = $found
                }

                $ok = [bool]$Summary.Ok
                $win.Title = if ($ok) { 'FAIS: schedule exported' } else { 'FAIS: export failed' }
                $ui.ResultBadge.Background = Get-PBBrush $(if ($ok) { '#16A34A' } else { '#DC2626' })
                $ui.ResultIcon.Text = if ($ok) { [string][char]0x2714 } else { [string][char]0x2716 }
                $ui.ResultTitle.Text = [string]$Summary.Title
                $ui.ResultTitle.Foreground = Get-PBBrush $(if ($ok) { '#14532D' } else { '#991B1B' })
                $ui.ResultSubtitle.Text = [string]$Summary.Subtitle
                $first = $true
                foreach ($ln in @($Summary.Lines)) {
                                $tb = New-Object System.Windows.Controls.TextBlock
                                $tb.Text = [string]$ln.Text
                                $tb.TextWrapping = [System.Windows.TextWrapping]::Wrap
                                $tb.Margin = New-Object System.Windows.Thickness(0, 1, 0, 1)
                                switch ([string]$ln.Kind) {
                                                'Head' {
                                                                $tb.FontWeight = [System.Windows.FontWeights]::SemiBold
                                                                $tb.Foreground = Get-PBBrush '#334155'
                                                                $tb.Margin = New-Object System.Windows.Thickness(0, $(if ($first) { 0 } else { 10 }), 0, 2)
                                                }
                                                'Path' { $tb.Foreground = Get-PBBrush '#64748B'; $tb.FontSize = 11; $tb.Margin = New-Object System.Windows.Thickness(12, 0, 0, 3) }
                                                'Good' { $tb.Foreground = Get-PBBrush '#15803D' }
                                                'Warn' { $tb.Foreground = Get-PBBrush '#B45309' }
                                                'Hard' { $tb.Foreground = Get-PBBrush '#B91C1C' }
                                }
                                $first = $false
                                [void]$ui.ResultLines.Children.Add($tb)
                }

                # The buttons carry their file, so their handlers need nothing from this function
                $ui.OpenOutputButton.Tag = $OutputPath
                $ui.OpenAtisButton.Tag = $AtisPath
                $ui.OpenOutputButton.Visibility = if ($ok -and $OutputPath -and (Test-Path -LiteralPath $OutputPath)) { 'Visible' } else { 'Collapsed' }
                $ui.OpenAtisButton.Visibility = if ($ok -and $AtisPath -and (Test-Path -LiteralPath $AtisPath)) { 'Visible' } else { 'Collapsed' }
                foreach ($b in @($ui.OpenOutputButton, $ui.OpenAtisButton)) {
                                $b.Add_Click({ param($s, $a) try { Open-FaisWorkbooks -Path @([string]$s.Tag) } catch { Write-Host "Could not open $($s.Tag): $($_.Exception.Message)" -ForegroundColor Yellow } })
                }
                # The board has just closed: come to the front once, without staying on top of Excel
                $win.Topmost = $true
                $win.Add_ContentRendered({ param($s, $a) try { $s.Topmost = $false; [void]$s.Activate() } catch { } })
                return $win
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
                                & $check ($null -ne $script:PB.C.LogoImage.Source -and $script:PB.C.LogoImage.Visibility -eq 'Visible') 'logo shown in the title bar'

                                $step = 'quick entry'
                                $ip1 = Get-NormalizedLastName ([string]$Data.Instructors[0].Name)
                                $ip2 = Get-NormalizedLastName ([string]$Data.Instructors[1].Name)
                                $q = ConvertFrom-QuickEntry -Text "0730 DL $ip1, 0745 DL $ip2, 1015 OB 1315 $ip1, 1300 OFT" -Day $Day -Instructors $Data.Instructors -Students $Data.Students -Board $script:PB.Board
                                & $check ($q.Errors.Count -eq 0) ("quick entry errors: " + ($q.Errors -join '; '))
                                Invoke-PBCommand @{ Cmd = 'AddLines'; Lines = $q.Lines; Links = $q.Links; Status = 'test' }
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
                                & $check ($script:PB.Board.Lines.Count -eq 6) 'O&B added its return line'
                                Invoke-PBCommand @{ Cmd = 'Undo' }
                                & $check ($script:PB.Board.Lines.Count -eq 5) 'undo removed the return line'
                                $choices = @(Get-PBMissionChoices (Get-BoardLine $script:PB.Board $lines[0].Id))
                                $events = @($choices | Where-Object { (ConvertFrom-MissionChoiceText $_).Kind -eq 'Mission' -and $_ -notlike '* (not opted)' })
                                & $check ($events.Count -ge 1) 'mission dropdown has events'
                                if ($events.Count -ge 1) {
                                                # The real pick path, unless this event would ask about a rule break (a message box would stop the test)
                                                $advice = Get-PBMissionAdvice -Line (Get-BoardLine $script:PB.Board $lines[0].Id) -Mission $events[0]
                                                $pickCmd = if ($advice.NeedsConfirm) { 'SetMission' } else { 'ChooseMission' }
                                                Invoke-PBCommand @{ Cmd = $pickCmd; LineId = $lines[0].Id; Value = $events[0] }
                                                & $check ((Get-BoardLine $script:PB.Board $lines[0].Id).Mission -eq $events[0]) 'mission picked from the dropdown'
                                }
                                Invoke-PBCommand @{ Cmd = 'SetField'; LineId = $lines[0].Id; Field = 'Notes'; Value = 'self-test note' }
                                Invoke-PBDoEvents

                                $step = 'tooltip and mission list'
                                $slot = Get-PBSlotElement -LineId $lines[0].Id
                                & $check ($null -ne $slot) 'student slot found for the tooltip check'
                                if ($null -ne $slot) {
                                                Set-PBSlotToolTip $slot
                                                & $check ($slot.ToolTip -is [System.Windows.Controls.StackPanel]) 'student slot tooltip has the +/- reasons'
                                }
                                $vm0 = New-PBLineVM (Get-BoardLine $script:PB.Board $lines[0].Id)
                                & $check ($vm0.MissionOptions -is [System.Array] -and $vm0.MissionOptions.Count -ge 1 -and $vm0.MissionOptions[0] -eq '(auto)') 'mission list is a real list starting with (auto)'

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

                                $step = 'load a day from the Schedule sheet'
                                $inv = [System.Globalization.CultureInfo]::InvariantCulture
                                $day2 = (Get-NextFlyingDay -From ([datetime]::ParseExact($Day, 'yyyy-MM-dd', $inv))).ToString('yyyy-MM-dd')
                                $ipName = [string]$Data.Instructors[0].Name
                                $newSheetRow = {
                                                param([string]$At, [string]$Type, [string]$IP, [string]$Student, [int]$Row)
                                                [PSCustomObject]@{ ETD = "$day2 $($At):00"; Day = $day2; Time = $At; TimeOnly = $false; Instructor = $IP; SortieType = $Type
                                                                Student = $(if ($Student) { $Student } else { $null }); Mission = $null; Callsign = $null; Num = $null; Notes = $null; SheetRow = $Row }
                                }
                                $sheetRows = @((& $newSheetRow '07:30' 'Day Local' $ipName ([string]$Data.Students[0].Name) 2), (& $newSheetRow '13:00' 'OFT' 'SIM IP' '' 3))
                                $script:PB.Data.ScheduleSheet = [PSCustomObject]@{ Path = ''; Worksheet = 'Schedule'; Found = $true; Error = ''; Layout = 'test'; HeaderRow = 1; Rows = $sheetRows; Skipped = @(); ReadAt = (Get-Date) }
                                Update-PBSheetBanner
                                & $check ($script:PB.C.SheetBanner.Visibility -eq 'Visible' -and $script:PB.SheetBannerInfo.State -eq 'OtherDay') 'banner offers the sheet''s other day'
                                Invoke-PBSheetBannerLoad
                                Invoke-PBDoEvents
                                & $check (@(Get-BoardDayLines -Board $script:PB.Board -Day $day2).Count -eq 2) 'sheet day loaded onto the board'
                                & $check ((Get-PBActiveDay) -eq $day2) 'the loaded day is on screen'
                                & $check ($script:PB.DayColls[$day2].Count -eq 2) 'rows on screen match the loaded lines'
                                & $check ($script:PB.C.SheetBanner.Visibility -ne 'Visible') 'banner gone once the day is loaded'
                                $edited = @($sheetRows | ForEach-Object { $_.PSObject.Copy() })
                                $edited[1].Time = '13:30'; $edited[1].ETD = "$day2 13:30:00"
                                $script:PB.Data.ScheduleSheet.Rows = $edited
                                Update-PBSheetBanner
                                & $check ($script:PB.SheetBannerInfo.State -eq 'Outdated') 'a sheet change on an untouched day is safe to apply'
                                $updated = @(Invoke-PBSheetAutoUpdate)
                                & $check ($updated.Count -eq 1 -and @(Get-BoardDayLines -Board $script:PB.Board -Day $day2 | Where-Object { $_.ETD.ToString('HHmm') -eq '1330' }).Count -eq 1) 'untouched day updated from the sheet by itself'
                                $first2 = @(Get-BoardDayLines -Board $script:PB.Board -Day $day2)[0]
                                Invoke-PBCommand @{ Cmd = 'SetField'; LineId = $first2.Id; Field = 'Notes'; Value = 'edited on the board' }
                                $edited[0].Instructor = [string]$Data.Instructors[1].Name
                                Update-PBSheetBanner
                                & $check ($script:PB.SheetBannerInfo.State -eq 'Changed') 'banner asks before replacing an edited day'
                                & $check (@(Invoke-PBSheetAutoUpdate).Count -eq 0) 'an edited day is never replaced by itself'
                                Invoke-PBSheetBannerIgnore
                                & $check ($script:PB.C.SheetBanner.Visibility -ne 'Visible') 'keeping the board''s lines hides the banner'

                                $step = 'load a day dialog'
                                [void](Initialize-PBLoadDayDialog)
                                & $check (@($script:PBDlg.DayList.ItemsSource).Count -ge 1) 'dialog lists the sheet''s days'
                                $script:PBDlg.SourceExportRadio.IsChecked = $true
                                & $check ([string]$script:PBDlg.SourceInfoText.Text -ne '') 'dialog explains the exported source'
                                $script:PBDlg.SourceSheetRadio.IsChecked = $true
                                foreach ($it in @($script:PBDlg.DayList.ItemsSource)) { if ($it.Day -eq $day2) { $script:PBDlg.DayList.SelectedItem = $it } }
                                Submit-PBDialog
                                & $check ($null -ne $script:PBDlg.Result -and $script:PBDlg.Result.SourceDay -eq $day2 -and $script:PBDlg.Result.TargetDay -eq $day2) 'dialog returns the picked day'
                                $script:PBDlg = $null

                                $step = 'mission dropdown'
                                $pinned = @(Get-BoardDayLines -Board $script:PB.Board -Day $day2 | Where-Object { $_.Student })
                                & $check ($pinned.Count -ge 1) 'a student is pinned on the loaded day'
                                if ($pinned.Count -ge 1) {
                                                $opts = @(Get-PBMissionChoices $pinned[0])
                                                & $check ($opts[0] -eq '(auto)' -and $opts[-1] -eq 'Type a mission...') 'mission dropdown offers (auto) and typing a mission'
                                }

                                $step = 'add bulk lines dialog'
                                [void](Initialize-PBBulkLinesDialog -Day $day2)
                                & $check (-not $script:PBBulk.AddButton.IsEnabled) 'Add lines waits for the number of lines'
                                & $check ([string]$script:PBBulk.IntervalBox.Text -eq '3') 'minutes between lines starts at 3'
                                $script:PBBulk.StartTimeBox.Text = '1400'
                                $script:PBBulk.CountBox.Text = '400'
                                & $check (-not $script:PBBulk.AddButton.IsEnabled) 'too many lines is refused'
                                $script:PBBulk.CountBox.Text = '4'
                                & $check ($script:PBBulk.AddButton.IsEnabled -and [string]$script:PBBulk.PreviewText.Text -like '*1400 1403 1406 1409*') 'preview shows the times'
                                Submit-PBBulkLinesDialog
                                $bulkAsk = $script:PBBulk.Result
                                $script:PBBulk = $null
                                & $check ($null -ne $bulkAsk -and $bulkAsk.Count -eq '4' -and $bulkAsk.StartTime -eq '1400' -and $bulkAsk.SortieType -eq 'Day Local') 'dialog returns what was typed'
                                $before = @(Get-BoardDayLines -Board $script:PB.Board -Day $day2).Count
                                if ($null -ne $bulkAsk) {
                                                & $check (Add-PBBulkLines -Day $day2 -StartTime $bulkAsk.StartTime -Count $bulkAsk.Count -Every $bulkAsk.Every -SortieType $bulkAsk.SortieType) 'bulk lines added'
                                                & $check (@(Get-BoardDayLines -Board $script:PB.Board -Day $day2).Count -eq $before + 4 -and $script:PB.DayColls[$day2].Count -eq $before + 4) 'four new rows on the day'
                                                Invoke-PBCommand @{ Cmd = 'Undo' }
                                                & $check (@(Get-BoardDayLines -Board $script:PB.Board -Day $day2).Count -eq $before) 'one undo takes them all off'
                                }

                                $step = 'menu bar'
                                Select-PBDay $day2
                                Invoke-PBDoEvents
                                Update-PBMenuState
                                & $check ($script:PB.C.MainMenu.Items.Count -ge 7) 'menu bar has its menus'
                                & $check ([string]$script:PB.C.MenuFillDay.Header -like "*$(Format-BoardDayLabel $day2)*") 'Fill menu names the day on screen'
                                Update-PBTemplateMenu
                                & $check ($script:PB.C.MenuApplyTemplate.Items.Count -ge 1 -and $script:PB.C.MenuApplyTemplate.Items[0].IsEnabled) 'Apply a template lists the saved template'

                                $step = 'select a line and its details'
                                $open = @(Get-BoardDayLines -Board $script:PB.Board -Day $day2 | Where-Object { -not $_.Student -and -not (Test-IsReturnSortieType ([string]$_.SortieType)) })[0]
                                Select-PBLine $open.Id
                                & $check ($script:PB.SelectedLineId -eq $open.Id) 'line selected'
                                $n = 0
                                while (($script:PB.Details.ChecksPending -or $script:PB.Details.Queue.Count -gt 0) -and $n -lt 1000) { Invoke-PBDetailsTick; $n++ }
                                $named = @($Data.Students | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.Name) }).Count
                                & $check ($script:PB.Details.Items.Count -eq $named) 'details list ranks every student'
                                & $check ($script:PB.C.DetailsChecks.Children.Count -ge 1) 'details left side filled'
                                Update-PBLineMenu
                                & $check ($script:PB.C.MenuLine.Items.Count -ge 4) 'Line menu lists the selected line''s actions'
                                $top = $script:PB.Details.Items[0]
                                $script:PB.C.DetailsList.SelectedItem = $top
                                & $check ([string]$script:PB.Details.Focus -eq [string]$top.StudentName) 'picking a student in the list shows their check'
                                & $check ($script:PB.C.DetailsChecks.Children.Count -ge 6) 'their check has the rule sections'
                                & $check ([bool]$script:PB.C.DetailsPlaceButton.IsEnabled) 'Put on this line is ready'

                                $step = 'click and place'
                                $holdName = [string]$Data.Students[1].Name
                                Invoke-PBCommand @{ Cmd = 'Hold'; Kind = 'Student'; Name = $holdName }
                                Invoke-PBDoEvents
                                & $check ($null -ne $script:PB.Held -and $script:PB.C.HoldBanner.Visibility -eq 'Visible') 'clicking a puck picks it up'
                                Invoke-PBCommand @{ Cmd = 'PlaceHeld'; LineId = $open.Id }
                                Invoke-PBDoEvents
                                & $check ((Get-BoardLine $script:PB.Board $open.Id).Student -eq $holdName -and $null -eq $script:PB.Held) 'clicking a slot puts the held student there'
                                & $check ($script:PB.C.HoldBanner.Visibility -ne 'Visible') 'the banner goes once the student is placed'
                                Invoke-PBCommand @{ Cmd = 'Undo' }
                                Invoke-PBCommand @{ Cmd = 'PlaceOnSelected'; Kind = 'IP'; Name = $ipName }
                                & $check ((Get-BoardLine $script:PB.Board $open.Id).Instructor -eq $ipName) 'double-clicking an IP puts them on the selected line'
                                Invoke-PBCommand @{ Cmd = 'Undo' }
                                Close-PBPopup

                                $step = 'color key and export result windows'
                                $key = New-PBColorKeyWindow
                                & $check ($null -ne $key.Content) 'color key window builds'
                                $stats = [PSCustomObject]@{ Lines = 2; Pinned = 1; Filled = 1; Wing = 0; Return = 0; DS = 0; ForSale = 0 }
                                $summary = Get-FaisExportSummary -Ok $true -FillResult ([PSCustomObject]@{ Stats = $stats; Overrides = @() }) -Days @($day2) -OutputPath 'x.xlsx' -MasterDataPath 'y.xlsx'
                                $resultWin = Initialize-FaisExportResultWindow -Summary $summary
                                & $check ($resultWin.FindName('ResultLines').Children.Count -ge 4) 'export result window lists what was written'

                                $step = 'fill only this day'
                                Select-PBDay $day2
                                Invoke-PBDoEvents
                                & $check ([string]$script:PB.C.FillDayButton.Content -eq "Fill $(Format-BoardDayLabel $day2)") 'fill button names the day on screen'
                                & $check ($script:PB.C.FillAllButton.Visibility -eq 'Visible') 'Fill all days shows with two days on the board'
                                Invoke-PBFill -Scope Day
                                Invoke-PBDoEvents
                                & $check ($script:PB.Mode -eq 'Review' -and (@($script:PB.FillDays) -join ',') -eq $day2) 'fill covered only the day on screen'
                                & $check (@($script:PB.FillResult.Schedule | Where-Object { (Parse-ETD $_.ETD).ToString('yyyy-MM-dd') -ne $day2 }).Count -eq 0) 'fill result holds only that day'
                                & $check ($script:PB.C.FillDayButton.Visibility -ne 'Visible') 'fill buttons hide while reviewing'
                                Exit-PBReview
                                Update-PBAll

                                $step = 'export'
                                Invoke-PBFill
                                Invoke-PBExport
                                & $check ($null -ne $script:PB.Result -and $script:PB.Result.Action -eq 'Export') 'export returned the fill'
                                & $check ((@($script:PB.Result.Days) -join ',') -eq (@($script:PB.Board.Days) -join ',')) 'export covers every day after a full fill'
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
