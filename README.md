# FAIS

Flight scheduler for student pilots. You build the day on a **puck board**: drag IPs and students onto lines, type ETDs, see why a student fits a line or breaks a rule, then let the scheduler **fill the rest of the lines**.

## Running it

```
powershell.exe -STA -File .\FAIS.ps1
```
Add `-Trace` to print every scheduling decision to the console. FAIS restarts itself in STA mode if needed.

**Needs:** Windows PowerShell 5.1 and the ImportExcel module. `MasterData.xlsx` lives in the FAIS folder, next to `FAIS.ps1`. `OutputSchedule.xlsx` and `ATIS.xlsx` are written to the Desktop.

**You can leave Excel open.** If MasterData or OutputSchedule is open in Excel, FAIS saves it before reading it (so edits you have not saved yet are read too), and saves and closes it before writing it (so there is no "file is open" error), then opens it again afterwards. You only see a "file is open" box if Excel can't be reached, for example while a cell is still being edited.

**What happens:**
1. **Pre-run pop-ups.** Log any hooked (unsat) or incomplete events from the last schedule, then click **Open Puck Board**. These are saved right away (see *Data in and out*).
2. **Puck board.** Each day on MasterData's **Schedule** sheet opens as its own tab with its lines (or you type them), and you hardcode whoever you want. *(details below)*
3. **Fill.** The scheduler fills every open line: formation wingmen, direct support, O&B return legs and For Sale lines included. **Fill <day>** fills just the day on screen; **Fill all N days** (shown when the board has more than one day) fills them all. You review the result on the board.
4. **Export.** All or nothing: if any part fails, none of the files change. It writes the days you filled:
   - OutputSchedule: the **MasterSchedule** sheet (first tab). Only the filled days are replaced; other days already exported (today onward) are kept, so you can update **one day at a time**,
   - OutputSchedule: **each student's syllabus tab** (tomorrow's run starts from it),
   - MasterData: **SimDebt**, and **UnsatSim / UnsatFlight** reset to FALSE for students whose schedule cleans up the unsat.

   It also writes **ATIS.xlsx**: one tab per day (named like `Wednesday. October 21, 2026`). Each tab lists that day's flights in ETD order, then a thick border, then that day's sims (OFT / IFT) in ETD order, in the same columns as MasterSchedule. Each export rebuilds the tabs for the days it covers and keeps every other day's tab, so ATIS builds up a day-by-day history. The file opens on the first day you just exported. Set `WriteAtis = $false` in `$OutputSettings` to turn it off.

   Afterwards a window says what was written (days, line counts, files, unsats cleared, OVERRIDE lines), or, if the export failed, that nothing was written and why. **Open OutputSchedule** and **Open ATIS** open the files.

Closing the board without exporting writes nothing (only the pre-run pop-ups have already saved). Your lines are autosaved either way.

## Data in and out

| When | Reads | Writes |
|---|---|---|
| Start | MasterData: Schedule (see *MasterData's Schedule sheet*), Instructors, Students (first 100), SoftWeights, FormCallsigns, StudentAbsences. OutputSchedule: each student's tab (a student with no tab starts from the MasterData Syllabus sheet). Workbooks open in Excel are saved first. | Nothing |
| **Log Hooked Event** | | MasterData: UnsatSim or UnsatFlight = TRUE |
| **Log Incomplete Event** | | MasterData: SimDebt (only that event's debt comes off). OutputSchedule: that event's cells on the student's tab |
| Puck board and fill | MasterData's Schedule sheet again whenever MasterData is saved; **Load a day...** reads it (or the last exported MasterSchedule) fresh | Nothing (the board's own autosave in `Boards\` aside) |
| **Update Syllabi** | The training file you pick (flying schedule or Big Board) | OutputSchedule: the changed events' cells on the students' tabs (a student with no tab gets one). MasterData: SimDebt |
| **Export** | | OutputSchedule: MasterSchedule (the filled days), student tabs. MasterData: SimDebt, unsat resets. ATIS: the exported days' tabs |

- **MasterSchedule is the only sheet of OutputSchedule that is rewritten**, and only the filled days in it change (ATIS rebuilds the tabs of the days it exports). A student's tab is updated in place: only the Instructor, ETD and IsComplete cells of events that changed are written, so your own columns, notes and formatting stay. A student with no tab gets a new one.
- **MasterData Students sheet:** students are matched by the column headed SPName and SimDebt goes to the column headed SimDebt (columns B and P when those headings are missing). UnsatSim and UnsatFlight are written to the columns with those headings.
- **Absences** are attached only when the name matches a student's name exactly (capital letters aside). A name that matches no one is reported in the console.
- **Unsat clean-up:** after a hooked OFT, the student can't fly until an OFT. After a hooked IFT, an OFT or IFT will do. After a hooked flight, the student can't sim until they fly. The scheduler and the board treat the unsat as cleaned up for every line after that sim or flight, and the line's Notes say `Clears Unsat Sim (OFT)` (or IFT, or `Clears Unsat Flight`). MasterData only says TRUE or FALSE, so for an unsat typed into MasterData the hooked sim is taken to be the student's last sim flown (an OFT if unknown). If you re-plan a day whose schedule cleaned up an unsat, the unsat comes back until the new schedule cleans it up again.

## MasterData's Schedule sheet

The board reads the **Schedule** sheet and keeps watching it:

- **When the board opens,** every day the sheet has lines for gets its own day tab with those lines, from the board's first day (the next flying day) on. A sheet with lines for Fri, Mon and Tue opens three tabs. A day on the board that you built yourself (typed lines, a template) is not replaced: the blue banner asks first.
- **Lines dated before the board's first day are not loaded.** That includes today, because the board starts on the next flying day. If that is all the sheet has, the blue banner says so and offers **Put them on <first day>**. **Load a day...** can put any sheet day on any board day.
- **Whenever you save MasterData in Excel,** the board reads the sheet again (within a few seconds). A new day on the sheet gets its own tab. A day you have not changed on the board since it was loaded updates by itself. If you have changed that day on the board, a blue banner offers **Replace ... with the sheet** or **Keep the board's lines**. A day you take off with **Remove day** stays off until its lines on the sheet change.
- **Load a day...** (toolbar) lists the days the sheet has, with line counts. Pick one, pick the day to put it on, choose **Replace** or **Add**, and whether to keep the students and missions it names (they go on as hardcoded). The sheet is saved in Excel and read fresh when the dialog opens. Rows that could not be read are listed with the reason.

**The sheet's layout:**
- **Columns are found by their headings**, in any order: `ETD` (or `Time`), `Date`, `Instructor` (or `IP`), `SortieType` (or `Type`), `Student` (or `SP`), `Mission` (or `Event`), `Callsign`, `Num` (or `#`), `Notes` (or `Remarks`). With no heading row, the old fixed order is used: ETD, Instructor, SortieType, Student, Mission, Callsign, Num, Notes.
- **ETD** can be a date and time, or just a time (`0730`, `730`, `7:30`, `7:30 AM`, a time-formatted cell). A time-only row takes its date from the `Date` column, or from a row above it that holds only a date (for example `Tuesday, October 21, 2026`). With neither, it counts as the first day on the board.
- **SortieType** can be written short: `DL`, `OB`, `RTN`, `OFT`, `IFT`.
- **Blank rows** and repeated heading rows are skipped. (They used to stop the read, so days after a blank row never loaded.)

## Update Syllabi

**Update Syllabi** (toolbar) brings each student's syllabus in line with a training file. Pick the file and FAIS works out which one it is from its headings:

- **The day's flying schedule** (a tab with Student, Mission and Line Status columns). Its **Flying** and **Sims** tabs are read; the other tabs are ignored. For each FAIS student on it:
  - Line Status **CM** or **EXEC**: the line's mission is marked complete, with the line's Instructor and ETD (Start Time on the Sims tab).
  - Line Status **INC** or **CX**: the mission is made not complete (its IP and ETD are cleared).
  - Any other status, or none: nothing changes.
  - If the same student flies the same mission more than once that day, a complete line wins over an INC / CX one.
- **The Big Board** (a Student Name column, then a block of columns for each event: date, status, grade, IP). For each FAIS student on it and each event in their syllabus:
  - Anything in the event's columns: the event is marked complete, with the IP and the date (`MM-DD-YYYY-HHMM`) as the ETD.
  - Nothing there: the event is made not complete, except an event FAIS has planned for later than now (the Big Board can't show it yet), which is left as planned.

**What it leaves alone:**
- Events that aren't in the student's syllabus. Most Big Board columns are academics, and they are simply ignored.
- People who aren't FAIS students. The flying schedule covers the whole wing.
- Unsat flags. Log a hook with the pre-run pop-up as before.

**Names:**
- Rank tags (`[2d Lt]`), case, periods and extra spaces don't matter.
- A student matches on last name **and** first name (and middle initial when both have one). An initial matches the name it starts: `Quill, C.` matches `Quill, Casey B.`.
- A student MasterData knows by last name only matches when exactly one person on the file has that last name.
- A name that could be more than one student is left alone and listed.
- The IP is written the way MasterData's Instructors sheet spells them when they are on it, otherwise as the file has them, without the rank.

**Before anything is saved** you see what will change, student by student: events marked complete, events marked NOT complete (with the status, e.g. `(CX)`), IP / ETD updates, and each student's SimDebt change. Then come the people, events and planned events left alone. The console lists everything. **Yes** saves right away:
- the changed events' Instructor, ETD and IsComplete cells on each student's tab in OutputSchedule (nothing else on the tab is touched; a student with no tab gets one);
- SimDebt in MasterData, moved by each changed event, the same way a scheduled event or a logged incomplete moves it.

Both files are saved or neither. If they can't be saved (still open), the change is kept for this run and Export writes it. If you are reviewing a fill, it is thrown away (it was made from the old syllabi), so fill again.

## Changing a day you already exported

1. Open FAIS, click **Load a day...**, choose **Last exported schedule (OutputSchedule)** and the day (it is added to the board if it isn't there). Everyone goes back on the board where they were exported, hardcoded.
2. Change what changed (take a student off, swap an IP, add a line).
3. Press **Fill <day>** (the day on screen), then **Export**. Only the open lines are filled, and only that day is replaced in OutputSchedule. Other exported days are filled in as fixed context, so no student is double-booked or flown out of order; if your change breaks a line on another exported day, the status line says which.

## The puck board

| Area | What it does |
|---|---|
| **Add lines** box (top) | Type lines and press Enter. See *Typing lines* below. |
| **Add Bulk Lines** (top) | Asks for a start time, how many lines and the minutes between them (3 by default), then adds them to the day on screen: `0730`, `10` lines, `3` min gives 0730, 0733 ... 0757. Pick the line type too (an O&B gets its return leg). One **Ctrl+Z** takes the whole run off. |
| **Template / Copy previous day's lines** | Reuse a saved wave, or copy the last flying day's times, types and IPs. |
| **Load a day...** | Load one day from MasterData's Schedule sheet or from the last exported schedule. |
| **Update Syllabi** | Pick the day's flying schedule or the Big Board; FAIS works out which it is and marks each student's events complete or not complete from it. See *Update Syllabi*. |
| **Blue banner** | MasterData's Schedule sheet has lines the board doesn't: load them or keep the board's lines. |
| **Menu bar** (very top) | Every action in one place, with its key: File (load a day, templates, Update Syllabi), Edit (undo, type lines, bulk lines, re-check), **Line** (what can be done to the selected line), Day, Fill, View, Help. |
| **Fill <day> / Fill all N days** (bottom) | Fill and export just the day on screen, or every day on the board. |
| **Line details** (under the days) | The selected line's student against every rule, and every student ranked for that line. See *Line details* below. |
| **Colors** (top right) | A small window that says what every color on the board means (also Help > What the colors mean). |
| **Day tabs** | One tab per day. **Add day** builds several days at once; crew rest is checked across them. |
| **IP pucks** (left) | Drag onto a line's IP slot. The badge shows how many lines that IP has that day. |
| **Student pucks** (right) | Drag onto a line's student slot to hardcode them. Badges: CAP, INTL, HOLD, U-SIM, U-FLT, ABS times, turns that day. |

- **Two ways to put a puck on a line:** drag it onto the slot, or **click it to pick it up** (it gets a thick outline and a dark banner says what you are holding) and then **click the slot**. Hold **Ctrl** as you click the slot to keep holding the puck for another line; **Esc** or **Let go** puts it down. **Double-click** a puck to put it straight on the selected line.
- **Moving pucks:** drag between lines to move a puck (hold **Ctrl** to copy). Drop it back on its pool to take it off the line.
- **While you drag or hold a student,** slots turn **green** (fine), **amber** (needs a look) or **red** (breaks a hard rule).
- **When you drop a student,** a pop-up lists:
  - hard-rule flags (red ✖), for example `Turn time 2.8 hr < 3.25 hr minimum (vs 0730 FT5103)` or `Not opted for FE3102: FT5001 isn't complete yet` (a sim whose prerequisite is the dollar ride);
  - warnings, for example `Less than 3.5 hr TT, Flt/CC Approval Required`;
  - the **+ / −** score reasons, spelled out, for example `3 String +10`, `Breaks 4 String w/ Lee −5` or `2 Flights and 2 Sims in last 5 Days −6`.
- **Rule breaks are allowed.** The line is hardcoded anyway and marked `OVERRIDE: ...` on the output.
- **Corner marker** on every hardcoded student: green, amber, red, or grey while checking. Hover for the reasons; right-click for more (explain, take off, *Fly wing to...*, duplicate, delete line).
- **Mission dropdown:** `(auto)` picks the best-scoring event that passes the rules. Choose an event to hardcode it. After the events the student can fly on the line come the next ones they are **not opted for** yet (marked `(not opted)`), then **Type a mission...** for anything else.
- **A mission picked or typed in the dropdown that breaks a rule** (not opted, already flown, not in the syllabus, wrong line type) is explained before it goes on, with the mission the scheduler would put there: **Yes** puts the recommended mission on, **No** keeps yours (the line is flagged `OVERRIDE: Not opted for ...` on the output), **Cancel** changes nothing. You always have the final say. A line typed with `sp:` and `m:` in the Add lines box goes on as typed and opens the rule-check pop-up, so its flags show straight away.
- **Keyboard:** arrow keys on a mission dropdown open it; move with the arrows and press **Enter** to pick.
- **IFT / OFT lines** are never wingmen or direct support.

### Line details
Click anywhere on a line (or right-click it, or tab into it) to select it: it gets a blue outline, and the **Line details** panel under the days shows
- **left:** the line's student against the **hard rules** (✖ broken, or none broken), the **warnings**, and the **soft rules** (the + / − score reasons and the total), the same check as the pop-up after a drop. For an open line it explains that the scheduler fills it.
- **right:** **every student ranked for that line**: ✔ fits, ! fits with warnings, ✖ breaks a hard rule (with the first rule it breaks), each with the mission (auto) would pick and its score. Fits come first, best score first. The list fills in a few students at a time. **Hide rule breaks** leaves out the ✖ students.

Click a student in the list to see their full check on the left. **Double-click** them, or press **Put on this line**, to hardcode them there. Drag the bar above the panel to resize it; **Hide** (or View > Line details panel) closes it.

### Typing lines
| Type this | Get |
|---|---|
| `0730 DL` (or `730 dl`, `7:30 local`) | Day Local at 0730 |
| `1015 O&B` | O&B, plus its **O&B Return** 3 hr later (same IP, linked) |
| `1015 OB 1330` | O&B with the return at 1330 |
| `1330 RTN` | Return line linked to the latest O&B before it |
| `1300 OFT x2` | Two OFT lines (IFT works the same) |
| `0800 DL Jones sp:Smith m:FT5103` | IP Jones, student Smith hardcoded on FT5103 |
| `0730 DL, 0745 DL, 0800 DL` | Several lines at once |

**ETD box:** type a new time and press Enter. **Up/Down** moves it 5 min; **Shift+Up/Down** moves it 30 min. **Esc** cancels.

**Keys:**
- **Ctrl+L**: jump to the Add lines box
- **Ctrl+Z**: undo
- **Delete**: clear the selected slot
- **Ctrl+D** / **Ctrl+Delete**: duplicate / delete the selected line
- **Ctrl+PgUp** / **Ctrl+PgDn**: previous / next day
- **Esc**: let go of a puck you picked up
- **F5**: re-check everything
- **Ctrl+Enter**: fill all days; **Ctrl+Shift+Enter**: fill the day on screen

## Files

| Path | What |
|---|---|
| `FAIS.ps1` | Entry point and settings (`$BoardSettings`, `$OutputSettings`; the Teams copy is off by default) |
| `Modules\` | Code: engine (`SchedulingEngine`, `Syllabus`), board (`PuckBoard.xaml/.psm1`, `LoadDayDialog.xaml`, `BulkLinesDialog.xaml`, `ExportResultDialog.xaml`, `BoardLogo.png` (the title-bar logo), `BoardModel`, `Assessment`), Excel in/out (`DataDownload`, `DataUpload`, `ExcelFiles`), Update Syllabi (`SyllabusImport`), pre-run pop-ups (`PreRunMenu`), helpers (`DataManipulation`) |
| `Boards\` | Autosaved board, one file per day (created on first run) |
| `Templates\` | Saved wave templates |
| `OldFunctions.psm1` | Legacy code, kept for reference only; not loaded |
| `Tests\` | Self-test and Pester suites |

The MasterData **Schedule** sheet is optional: without it, type the lines or use templates.

**Output settings** (`$OutputSettings` in FAIS.ps1):

| Setting | Default | What |
|---|---|---|
| `WriteAtis` | `$true` | Write ATIS.xlsx on the Desktop |
| `SaveAndCloseExcel` | `$true` | Save and close a workbook in Excel before FAIS writes it |
| `ReopenExcel` | `$true` | Open those workbooks again afterwards |
| `CloseAllExcelWorkbooks` | `$false` | Save and close *every* open workbook before writing, not just FAIS's (never-saved workbooks are left open) |
| `KeepPastDays` | `0` | Exported days MasterSchedule keeps before today |

## Checking an install

```
powershell.exe -STA -ExecutionPolicy Bypass -File .\Tests\Invoke-SelfTest.ps1
```
This checks the engine, an Excel round trip (temp files only) and the puck board window (it may flash briefly). It prints PASS/FAIL and exits 0 when everything passes. The developer suites (`Tests\*.Tests.ps1`) need Pester 5. `DataIO.Tests.ps1` also needs ImportExcel and is skipped without it.

## Scoring changes from the old version
These are deliberate fixes, so expect some different picks:
- **Non-Counter now applies:** an IP's `IsCounter` is read. A blank cell counts as a counter.
- **OFT/IFT mismatch penalty now applies:** `OFTIFTPen` is read from SoftWeights.
- **Picks count toward turns:** students the scheduler picks are marked complete, so turn time and triple turns see them.
- **Hardcoded lines later in the day hold the student's time,** so earlier fills respect them.
- **SimDebt moves for every scheduled event:** hardcoded lines, wingmen and return legs included.
- **Taking an event back only takes back its own SimDebt:** for an event logged incomplete, or a planned event on a day being re-planned. A SimDebt you set by hand in MasterData is no longer recalculated from scratch.
- **Unsats clear themselves** once the schedule has the right re-sim or flight (see *Data in and out*).
- **Formation callsigns** restart at the first FormCallsigns pair each day.
- **Teams day sheets** are replaced when a day is exported again (they used to keep the old lines too).
- **SimDebt** is written to the MasterData column headed `SimDebt` (column P only when there is no such heading).
