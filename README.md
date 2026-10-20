# FAIS

Flight scheduler for student pilots. You build the day on a **puck board**: drag IPs and students onto lines, type ETDs, see why a student fits a line or breaks a rule, then let the scheduler **fill the rest of the lines**.

## Running it

```
powershell.exe -STA -File .\FAIS.ps1
```
Add `-Trace` to print every scheduling decision to the console. FAIS restarts itself in STA mode if needed.

**Needs:** Windows PowerShell 5.1 and the ImportExcel module. `MasterData.xlsx` lives in the folder *above* the FAIS folder. `OutputSchedule.xlsx` is written to the Desktop.

**What happens:**
1. **Pre-run pop-ups.** Log any hooked (unsat) or incomplete events from the last schedule, then click **Open Puck Board**.
2. **Puck board.** Build the day's lines and hardcode whoever you want. *(details below)*
3. **Fill the rest of the lines.** The scheduler fills every open line: formation wingmen, direct support, O&B return legs and For Sale lines included. You review the result on the board.
4. **Export.** This writes OutputSchedule.xlsx:
   - **MasterSchedule** sheet (first tab),
   - **each student's syllabus tab** (tomorrow's run starts from it),
   - **SimDebt** to MasterData.

Closing the board without exporting writes nothing. Your lines are autosaved either way.

## The puck board

| Area | What it does |
|---|---|
| **Add lines** box (top) | Type lines and press Enter. See *Typing lines* below. |
| **Template / Copy previous day's lines** | Reuse a saved wave, or copy the last flying day's times, types and IPs. |
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
- **Ctrl+Enter**: fill

## Files

| Path | What |
|---|---|
| `FAIS.ps1` | Entry point and settings (`$BoardSettings`, `$OutputSettings`; the Teams copy is off by default) |
| `Modules\` | Code: engine (`SchedulingEngine`, `Syllabus`), board (`PuckBoard.xaml/.psm1`, `BoardModel`, `Assessment`), Excel in/out (`DataDownload`, `DataUpload`), pre-run pop-ups (`PreRunMenu`), helpers (`DataManipulation`) |
| `Boards\` | Autosaved board, one file per day (created on first run) |
| `Templates\` | Saved wave templates |
| `OldFunctions.psm1` | Legacy code, kept for reference only; not loaded |
| `Tests\` | Self-test and Pester suites |

The MasterData **Schedule** sheet is no longer needed. If it has rows, the board offers **Import Schedule sheet**.

## Checking an install

```
powershell.exe -STA -ExecutionPolicy Bypass -File .\Tests\Invoke-SelfTest.ps1
```
This checks the engine, an Excel round trip (temp files only) and the puck board window (it may flash briefly). It prints PASS/FAIL and exits 0 when everything passes. The developer suites (`Tests\*.Tests.ps1`) need Pester 5.

## Scoring changes from the old version
These are deliberate fixes, so expect some different picks:
- **Non-Counter now applies:** an IP's `IsCounter` is read. A blank cell counts as a counter.
- **OFT/IFT mismatch penalty now applies:** `OFTIFTPen` is read from SoftWeights.
- **Picks count toward turns:** students the scheduler picks are marked complete, so turn time and triple turns see them.
- **Hardcoded lines later in the day hold the student's time,** so earlier fills respect them.
- **SimDebt moves for every scheduled event:** hardcoded lines, wingmen and return legs included.
- **Formation callsigns** restart at the first FormCallsigns pair each day.
