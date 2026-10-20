# FAIS

Flight scheduler for student pilots. You build the day on a **puck board**: drag IPs and students onto lines, type ETDs, see why a student fits a line or breaks a rule, then let the scheduler **fill the rest of the lines**.

## Running it

```
powershell.exe -STA -File .\FAIS.ps1
```
Add `-Trace` to print every scheduling decision to the console. FAIS restarts itself in STA mode if needed.

**Needs:** Windows PowerShell 5.1 and the ImportExcel module. `MasterData.xlsx` lives in the folder *above* the FAIS folder. `OutputSchedule.xlsx` and `ATIS.xlsx` are written to the Desktop.

**You can leave Excel open.** If MasterData or OutputSchedule is open in Excel, FAIS saves it before reading it (so edits you have not saved yet are read too), and saves and closes it before writing it (so there is no "file is open" error), then opens it again afterwards. You only see a "file is open" box if Excel can't be reached, for example while a cell is still being edited.

**What happens:**
1. **Pre-run pop-ups.** Log any hooked (unsat) or incomplete events from the last schedule, then click **Open Puck Board**. These are saved right away (see *Data in and out*).
2. **Puck board.** The lines for each day come from MasterData's **Schedule** sheet (or you type them), and you hardcode whoever you want. *(details below)*
3. **Fill.** The scheduler fills every open line: formation wingmen, direct support, O&B return legs and For Sale lines included. Tick **Only this day** to fill just the day on screen. You review the result on the board.
4. **Export.** All or nothing: if any part fails, none of the files change. It writes the days you filled:
   - OutputSchedule: the **MasterSchedule** sheet (first tab). Only the filled days are replaced; other days already exported (today onward) are kept, so you can update **one day at a time**,
   - OutputSchedule: **each student's syllabus tab** (tomorrow's run starts from it),
   - MasterData: **SimDebt**, and **UnsatSim / UnsatFlight** reset to FALSE for students whose schedule cleans up the unsat.

   It also writes **ATIS.xlsx**: one tab per day (named like `Wednesday. October 21, 2026`). Each tab lists that day's flights in ETD order, then a thick border, then that day's sims (OFT / IFT) in ETD order, in the same columns as MasterSchedule. Each export rebuilds the tabs for the days it covers and keeps every other day's tab, so ATIS builds up a day-by-day history. The file opens on the first day you just exported. Set `WriteAtis = $false` in `$OutputSettings` to turn it off.

Closing the board without exporting writes nothing (only the pre-run pop-ups have already saved). Your lines are autosaved either way.

## Data in and out

| When | Reads | Writes |
|---|---|---|
| Start | MasterData: Schedule (see *MasterData's Schedule sheet*), Instructors, Students (first 100), SoftWeights, FormCallsigns, StudentAbsences. OutputSchedule: each student's tab (a student with no tab starts from the MasterData Syllabus sheet). Workbooks open in Excel are saved first. | Nothing |
| **Log Hooked Event** | | MasterData: UnsatSim or UnsatFlight = TRUE |
| **Log Incomplete Event** | | MasterData: SimDebt (only that event's debt comes off). OutputSchedule: that event's cells on the student's tab |
| Puck board and fill | MasterData's Schedule sheet again whenever MasterData is saved; **Load a day...** reads it (or the last exported MasterSchedule) fresh | Nothing (the board's own autosave in `Boards\` aside) |
| **Export** | | OutputSchedule: MasterSchedule (the filled days), student tabs. MasterData: SimDebt, unsat resets. ATIS: the exported days' tabs |

- **MasterSchedule is the only sheet of OutputSchedule that is rewritten**, and only the filled days in it change (ATIS rebuilds the tabs of the days it exports). A student's tab is updated in place: only the Instructor, ETD and IsComplete cells of events that changed are written, so your own columns, notes and formatting stay. A student with no tab gets a new one.
- **MasterData Students sheet:** students are matched by the column headed SPName and SimDebt goes to the column headed SimDebt (columns B and P when those headings are missing). UnsatSim and UnsatFlight are written to the columns with those headings.
- **Absences** are attached only when the name matches a student's name exactly (capital letters aside). A name that matches no one is reported in the console.
- **Unsat clean-up:** after a hooked OFT, the student can't fly until an OFT. After a hooked IFT, an OFT or IFT will do. After a hooked flight, the student can't sim until they fly. The scheduler and the board treat the unsat as cleaned up for every line after that sim or flight, and the line's Notes say `Clears Unsat Sim (OFT)` (or IFT, or `Clears Unsat Flight`). MasterData only says TRUE or FALSE, so for an unsat typed into MasterData the hooked sim is taken to be the student's last sim flown (an OFT if unknown). If you re-plan a day whose schedule cleaned up an unsat, the unsat comes back until the new schedule cleans it up again.

## MasterData's Schedule sheet

The board reads the **Schedule** sheet one day at a time and keeps watching it:

- **When the board opens,** each day that is still empty gets its lines from the sheet.
- **Whenever you save MasterData in Excel,** the board reads the sheet again (within a few seconds). A day you have not changed on the board since it was loaded updates by itself. If you have changed that day on the board, a blue banner offers **Replace ... with the sheet** or **Keep the board's lines**. The banner also offers later days the sheet has that are not on the board yet.
- **Load a day...** (toolbar) lists the days the sheet has, with line counts. Pick one, pick the day to put it on, choose **Replace** or **Add**, and whether to keep the students and missions it names (they go on as hardcoded). The sheet is saved in Excel and read fresh when the dialog opens. Rows that could not be read are listed with the reason.

**The sheet's layout:**
- **Columns are found by their headings**, in any order: `ETD` (or `Time`), `Date`, `Instructor` (or `IP`), `SortieType` (or `Type`), `Student` (or `SP`), `Mission` (or `Event`), `Callsign`, `Num` (or `#`), `Notes` (or `Remarks`). With no heading row, the old fixed order is used: ETD, Instructor, SortieType, Student, Mission, Callsign, Num, Notes.
- **ETD** can be a date and time, or just a time (`0730`, `730`, `7:30`, `7:30 AM`, a time-formatted cell). A time-only row takes its date from the `Date` column, or from a row above it that holds only a date (for example `Tuesday, October 21, 2026`). With neither, it counts as the first day on the board.
- **SortieType** can be written short: `DL`, `OB`, `RTN`, `OFT`, `IFT`.
- **Blank rows** and repeated heading rows are skipped. (They used to stop the read, so days after a blank row never loaded.)

## Changing a day you already exported

1. Open FAIS, click **Load a day...**, choose **Last exported schedule (OutputSchedule)** and the day (it is added to the board if it isn't there). Everyone goes back on the board where they were exported, hardcoded.
2. Change what changed (take a student off, swap an IP, add a line).
3. Tick **Only this day**, **Fill this day**, then **Export**. Only the open lines are filled, and only that day is replaced in OutputSchedule. Other exported days are filled in as fixed context, so no student is double-booked or flown out of order; if your change breaks a line on another exported day, the status line says which.

## The puck board

| Area | What it does |
|---|---|
| **Add lines** box (top) | Type lines and press Enter. See *Typing lines* below. |
| **Template / Copy previous day's lines** | Reuse a saved wave, or copy the last flying day's times, types and IPs. |
| **Load a day...** | Load one day from MasterData's Schedule sheet or from the last exported schedule. |
| **Blue banner** | MasterData's Schedule sheet has lines the board doesn't: load them or keep the board's lines. |
| **Only this day** (bottom) | Fill and export just the day on screen. |
| **Day tabs** | One tab per day. **Add day** builds several days at once; crew rest is checked across them. |
| **IP pucks** (left) | Drag onto a line's IP slot. The badge shows how many lines that IP has that day. |
| **Student pucks** (right) | Drag onto a line's student slot to hardcode them. Badges: CAP, INTL, HOLD, U-SIM, U-FLT, ABS times, turns that day. |

- **Moving pucks:** drag between lines to move a puck (hold **Ctrl** to copy). Drop it back on its pool to take it off the line.
- **While you drag a student,** slots turn **green** (fine), **amber** (needs a look) or **red** (breaks a hard rule).
- **When you drop a student,** a pop-up lists:
  - hard-rule flags (red ✖), for example `Turn time 2.8 hr < 3.25 hr minimum (vs 0730 FT5103)`;
  - warnings, for example `Less than 3.5 hr TT, Flt/CC Approval Required`;
  - the **+ / −** score reasons, for example `Continue String +10` or `Exhaustion −6`.
- **Rule breaks are allowed.** The line is hardcoded anyway and marked `OVERRIDE: ...` on the output.
- **Corner marker** on every hardcoded student: green, amber, red, or grey while checking. Hover for the reasons; right-click for more (explain, take off, *Fly wing to...*, duplicate, delete line).
- **Mission dropdown:** `(auto)` picks the best-scoring event that passes the rules. Choose an event to hardcode it.
- **IFT / OFT lines** are never wingmen or direct support.

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
- **F5**: re-check everything
- **Ctrl+Enter**: fill (just this day when **Only this day** is ticked)

## Files

| Path | What |
|---|---|
| `FAIS.ps1` | Entry point and settings (`$BoardSettings`, `$OutputSettings`; the Teams copy is off by default) |
| `Modules\` | Code: engine (`SchedulingEngine`, `Syllabus`), board (`PuckBoard.xaml/.psm1`, `LoadDayDialog.xaml`, `BoardModel`, `Assessment`), Excel in/out (`DataDownload`, `DataUpload`, `ExcelFiles`), pre-run pop-ups (`PreRunMenu`), helpers (`DataManipulation`) |
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
