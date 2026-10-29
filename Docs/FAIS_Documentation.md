# FAIS Documentation and How-To Guide

This guide covers FAIS from several directions on purpose. The same ideas (lines, pins, hard rules, scores, strings, export) come up in the user guide, the logic sections, the code sections and the worked examples, each time from a different angle. If one explanation doesn't click, the next one usually will. Cross-references like **(see §9)** point to the other places a topic is covered.

> **The README is the short version.** `README.md` in the FAIS folder is the quick reference. This document is the long version: it explains *why* as well as *what*, and walks through the code.

---

## Contents

- **Part A: Getting oriented**
  - [§0 How to open this document](#0-how-to-open-this-document)
  - [§1 FAIS in one page](#1-fais-in-one-page)
  - [§2 Glossary](#2-glossary)
  - [§3 Installing and first run](#3-installing-and-first-run)
- **Part B: How to use FAIS**
  - [§4 The daily workflow, step by step](#4-the-daily-workflow-step-by-step)
  - [§5 The pre-run pop-ups](#5-the-pre-run-pop-ups)
  - [§6 The puck board, area by area](#6-the-puck-board-area-by-area)
  - [§7 Typing lines, keys and shortcuts](#7-typing-lines-keys-and-shortcuts)
  - [§8 Filling, reviewing and exporting](#8-filling-reviewing-and-exporting)
  - [§9 Changing a day you already exported](#9-changing-a-day-you-already-exported)
  - [§10 Update Syllabi](#10-update-syllabi)
  - [§11 Add Student Absence](#11-add-student-absence)
- **Part C: The logic (how FAIS decides)**
  - [§12 How one line gets filled: the selection process](#12-how-one-line-gets-filled-the-selection-process)
  - [§13 Hard rules (pass / fail)](#13-hard-rules-pass--fail)
  - [§14 Soft rules (the score)](#14-soft-rules-the-score)
  - [§15 Strings and counters](#15-strings-and-counters)
  - [§16 Sim debt, discontinuity, exhaustion and BIT](#16-sim-debt-discontinuity-exhaustion-and-bit)
  - [§17 Formations, wingmen and direct support](#17-formations-wingmen-and-direct-support)
  - [§18 O&Bs and return legs](#18-obs-and-return-legs)
  - [§19 Unsats, holds and absences](#19-unsats-holds-and-absences)
  - [§20 Tie-breaks and the order lines are filled in](#20-tie-breaks-and-the-order-lines-are-filled-in)
- **Part D: How the code works**
  - [§21 Architecture at a glance](#21-architecture-at-a-glance)
  - [§22 Module by module](#22-module-by-module)
  - [§23 The data objects](#23-the-data-objects)
  - [§24 Life of a run, through the code](#24-life-of-a-run-through-the-code)
  - [§25 The puck board's internals](#25-the-puck-boards-internals)
  - [§26 Excel in and out](#26-excel-in-and-out)
- **Part E: Data reference**
  - [§27 MasterData.xlsx, sheet by sheet](#27-masterdataxlsx-sheet-by-sheet)
  - [§28 OutputSchedule.xlsx and ATIS.xlsx](#28-outputschedulexlsx-and-atisxlsx)
  - [§29 What is read and written, and when](#29-what-is-read-and-written-and-when)
  - [§30 Settings](#30-settings)
- **Part F: The same thing, explained again**
  - [§31 Life of a student](#31-life-of-a-student)
  - [§32 Life of a line](#32-life-of-a-line)
  - [§33 Worked scoring example](#33-worked-scoring-example)
  - [§34 Scenarios and recipes](#34-scenarios-and-recipes)
- **Part G: Help**
  - [§35 Troubleshooting](#35-troubleshooting)
  - [§36 FAQ](#36-faq)
  - [§37 Testing](#37-testing)
  - [§38 Developer guide: changing FAIS safely](#38-developer-guide-changing-fais-safely)
  - [§39 File map](#39-file-map)
  - [§40 Quick index](#40-quick-index)

---

# Part A: Getting oriented

## §0 How to open this document

This guide comes in two copies in the `Docs` folder. They have the same content:

| File | Best way to open it |
|---|---|
| `Docs\FAIS_Documentation.html` | **Double-click it.** It opens in your web browser (Edge, Chrome, Firefox) with a clickable table of contents. No internet needed. |
| `Docs\FAIS_Documentation.md` | Read it on **GitHub** (it renders there), or open it in VS Code and press **Ctrl+Shift+V** for the preview. In Notepad it shows as plain text with `#` and `**` marks. |

**Getting it onto your computer:**

1. **With Git:** in the FAIS folder run `git pull`, or check out the branch the docs are on. The files appear in `Docs\`.
2. **Without Git:** on GitHub, open the repository, pick the branch, click **Code > Download ZIP**, unzip it, and open `Docs\FAIS_Documentation.html`.
3. **Just one file:** on GitHub, open `Docs/FAIS_Documentation.html`, click **Download raw file** (the download icon), save it, then double-click it.

> If Windows shows the HTML file as blocked (it came from the internet), right-click it, choose **Properties**, tick **Unblock**, and click OK. The page has no scripts, so it's safe to open.

## §1 FAIS in one page

**What it is.** FAIS builds the daily flying schedule for student pilots (SPs). It is a Windows PowerShell program with a WPF window called the **puck board**. It reads its data from Excel (`MasterData.xlsx`, `OutputSchedule.xlsx`) and writes the finished schedule back to Excel (`OutputSchedule.xlsx`, `ATIS.xlsx`).

**The idea.** A schedule is a list of **lines** (ETD, line type, IP). Every line needs a student and a mission. You put on the students you already know about (**hardcoding**, or **pinning**). The **scheduler** fills the rest:

1. It throws out every student who **can't** fly the line (**hard rules**: crew rest, turn time, prereqs and so on).
2. It **scores** everyone left (**soft rules**: keep students with the same IP, pay down sim debt, prevent long breaks).
3. The highest score gets the line.

**The run, in five steps:**

```
 Start FAIS ─► Pre-run pop-ups ─► Puck board ─► Fill ─► Export
               (unsats, holds,     (build lines,  (scheduler   (OutputSchedule,
                incompletes,        pin people)    fills open   MasterData,
                absences)                          lines)       ATIS)
```

**You always have the final say.** Anything the scheduler picks can be moved. Anything that breaks a rule can still be placed: it goes out flagged `OVERRIDE`.

**Nothing is written until Export.** The exceptions are the pre-run pop-ups, Update Syllabi and Add Student Absence, which save right away. The board also autosaves its own lines to `Boards\` so you never lose work.

## §2 Glossary

| Term | Meaning |
|---|---|
| **SP** | Student pilot. |
| **IP** | Instructor pilot. Listed on MasterData's **Instructors** sheet. |
| **Line** / **sortie** | One row of the schedule: an ETD, a line type, an IP, and (once filled) a student and a mission. The code calls them *sorties* or *rows*; the board calls them *lines*. |
| **ETD** | Estimated time of departure. The line's start time. |
| **Line type** / **SortieType** | `Day Local` (DL), `O&B`, `O&B Return` (RTN), `OFT`, `IFT`. Decides which events can be flown (§13). |
| **Event** / **mission** | One item of the syllabus, for example `FT5001`. |
| **Syllabus** | The ordered list of events every student goes through, with prereqs. Each student has their own copy (their tab in OutputSchedule) that records what they have completed. |
| **Opted** | A student is *opted* for an event when it isn't complete and all its prereqs are complete. Only opted events can be scheduled automatically. |
| **Hardcode** / **pin** | Putting a specific student (and maybe a mission) on a line yourself. The scheduler never moves a pinned student. |
| **Fill** | Running the scheduler over the open lines. |
| **Hard rule** | A pass/fail rule (crew rest, turn time...). A student who breaks one is never picked by the scheduler. |
| **Soft rule** | A scoring rule. Points are added or taken off; highest total wins. |
| **SoftWeights** | The MasterData sheet that holds the soft-rule point values. |
| **OVERRIDE** | Note written on a line whose pinned student breaks a hard rule. Allowed, but visible. |
| **Counter** | An IP who flies a sortie with a student. AETCI continuity rules limit how many a student should have. |
| **Non-counter** | An IP with `IsCounter` = FALSE (supervisors). Doesn't count, unless they fly the student two rides in a row. |
| **String** | Rides in a row with the same counter, in one category (FT, FF, AA). 3 is the goal. |
| **Zipper** | One different counter in the middle of a string (`Lee, Lee, Ortiz, Lee`). The string survives. |
| **Sim debt** | A running number: flights add to it, sims pay it down. Keeps students from running out of sims (a *sim stop*). |
| **Discontinuity score** | How broken up a student's strings have been. High = lots of 1-strings and counters. Gives those students priority to stay with an IP. |
| **BIT** | Break in training: too many days without a flight. |
| **Unsat** / **hook** | A failed event. An unsat sim blocks flying until a re-sim; an unsat flight blocks sims until a re-fly. |
| **CAP** | Students flagged `IsCAP`. One flight a day, no O&Bs, need an experienced IP. |
| **INTL** | International students (`IsInternational`). Need an experienced IP. |
| **Wingman** | A second student flying formation with a formation lead. |
| **DS (Direct support)** | When no student can fly wing, an open line is claimed as `Need Form SP` to support the formation lead. |
| **For Sale** / **NEED SP** | A line with no student. The scheduler writes For Sale when no one can fly it; Export writes NEED SP for any open line. |
| **Reservation** | A pinned line later in the day "holding" the student's time so earlier lines respect it. |
| **Context rows** | Days already exported that aren't being filled, loaded as fixed lines so the fill doesn't double-book anyone. |

## §3 Installing and first run

**Needs:**
- Windows with **Windows PowerShell 5.1** (the blue PowerShell that comes with Windows).
- The **ImportExcel** PowerShell module. If it's missing, FAIS says *"Please install or copy the ImportExcel module"*. Install it with `Install-Module ImportExcel -Scope CurrentUser`, or copy the module folder into `Documents\WindowsPowerShell\Modules`.
- `MasterData.xlsx` in the FAIS folder, next to `FAIS.ps1`.
- `OutputSchedule.xlsx` in the same folder (created on the first export if it isn't there).

**Run it:**

```
cd <your FAIS folder>
powershell.exe -STA -ExecutionPolicy Bypass -File .\FAIS.ps1
```

- `-STA` is needed by the WPF window. If you forget it, FAIS restarts itself in STA mode.
- `-ExecutionPolicy Bypass` avoids the "running scripts is disabled" error on locked-down machines.
- Add `-Trace` to print every scheduling decision (every candidate and score) to the console.

**Check an install** without touching your real files:

```
powershell.exe -STA -ExecutionPolicy Bypass -File .\Tests\Invoke-SelfTest.ps1
```

It prints PASS/FAIL for each check and exits 0 when all pass (see §37).

**Text-only copies.** If your machine blocks `.ps1` downloads, the repository also has `FAIS_txt.zip`, `FAIS_catalog.txt` and `FAIS_txt_update\` (only the changed files as `.txt`). `FAIS_txt_update\Restore.txt` is a short script that renames `X.psm1.txt` back to `X.psm1` and unblocks the files.

---

# Part B: How to use FAIS

## §4 The daily workflow, step by step

This is the normal day. Each step is covered in more detail later.

1. **Close nothing.** You can leave `MasterData.xlsx` and `OutputSchedule.xlsx` open in Excel. FAIS saves them before reading (so unsaved edits are read too) and saves/closes/reopens them around writes.
2. **Start FAIS** (§3). The console shows what was loaded: *"Loaded 14 students, 103 IPs, 72 formation callsigns..."* and the days found on the Schedule sheet.
3. **Pre-run pop-ups** (§5). Log anything that went wrong on the last schedule:
   - **Log Hooked Event** for unsats,
   - **Log Incomplete Event** for events that didn't happen,
   - **Change Student HOLD status**,
   - **Add Student Absence**.

   Then click **Open Puck Board**.
4. **Build the day on the puck board** (§6, §7). The lines usually come from MasterData's **Schedule** sheet automatically, one tab per day. Otherwise type them (`0730 DL`), use **Add Bulk Lines**, a template, or **Copy previous day's lines**.
5. **Put IPs on lines** (drag IP pucks onto the IP slots) if the sheet didn't already.
6. **Hardcode the students you already know about** (drag student pucks onto student slots). Watch the corner marker: green = fine, amber = warning, red = breaks a hard rule.
7. **Fill.** Press **Fill <day>** (Ctrl+Shift+Enter) or **Fill all N days** (Ctrl+Enter). The scheduler's picks go on the board in blue italics, marked `AUTO`.
8. **Review.** Click any line to see **Line details**: why that student fits, their + / − score, and **See options** to rank every other student for that line. Move anyone you want. **Clear fill** / **Re-fill** / **Ctrl+Z** are there if you want to start over.
9. **Export** (bottom bar). Writes the board exactly as it stands. A window tells you what was written; **Open OutputSchedule** and **Open ATIS** open the files.

Closing the board without exporting writes nothing to OutputSchedule or MasterData (only the pop-ups have saved). The board's lines are kept in `Boards\` for next time.

## §5 The pre-run pop-ups

Before the board opens, FAIS reads the **last exported schedule** (the MasterSchedule sheet of OutputSchedule) and shows a small menu. Use it to record what happened since then. Each choice saves straight away.

| Button | What you do | What it changes |
|---|---|---|
| **Log Hooked Event** | Tick the lines from the last schedule that were unsat (hooked). | `UnsatSim` (for an OFT / IFT line) or `UnsatFlight` (for a flight) = TRUE on MasterData's Students sheet. The student's next lines are limited (§19). |
| **Log Incomplete Event** | Tick the lines that didn't happen (weather, maintenance, CX). | The event is made **not complete** on the student's tab (IP and ETD cleared), and only that event's part of **SimDebt** and **DiscontinuityScore** is taken back off. |
| **Change Student HOLD status** | Tick students to put on hold (checked = on hold). | `OnHold` = TRUE / FALSE. A student on hold is never scheduled. |
| **Add Student Absence** | Block students off for a time (§11). | A line on the **StudentAbsences** sheet. |
| **Open Puck Board** | | Goes on to the board. |
| **Exit** | | Stops FAIS. |

If a file can't be saved (for example a cell is still being edited in Excel), the change still applies for this run and **Export** writes it later. The message box tells you which happened.

> **Why do this before the board?** The scheduler works from each student's syllabus. If FT5003 is still marked complete when it was really cancelled, the scheduler thinks the student is ready for FT5004. Logging the incomplete puts the syllabus right first.

## §6 The puck board, area by area

```
┌──────────────────────────────────────────────────────────────────────────────┐
│ File  Edit  Line  Day  Fill  View  Help                       (menu bar)     │
│ [Add lines: 0730 DL ...] [Add Bulk Lines] [Template] [Load a day...]         │
│ [Update Syllabi] [Add Student Absence]                        [Colors]       │
│ ── blue banner: "MasterData's Schedule sheet has lines for ..." ──           │
├────────────┬──────────────────────────────────────────────┬──────────────────┤
│ IP pucks   │  Day tabs:  Wed 21 Oct | Thu 22 Oct | +       │ Student pucks    │
│ (search)   │  ┌─ETD──Type──────IP slot──────Student slot──┐│ (search)         │
│ Albertson  │  │0730  Day Local [Lee      ] [Smith  FT5004]││ Smith  CAP  1t   │
│ Avara  NC  │  │0745  O&B       [Ortiz    ] [  open       ]││ Jones  U-SIM     │
│ ...        │  │1045  O&B Return[Ortiz    ] [ follows O&B ]││ Diaz   HOLD      │
│            │  └───────────────────────────────────────────┘│ ...              │
│            │ ═══ Line details ═══════════════════════════  │                  │
│            │ left: hard rules / warnings / + − score       │                  │
│            │ right: [See options] ranked list of students  │                  │
├────────────┴──────────────────────────────────────────────┴──────────────────┤
│ status line                [Fill Wed 21 Oct] [Fill all 2 days] [Export]       │
└──────────────────────────────────────────────────────────────────────────────┘
```

| Area | What it does |
|---|---|
| **Add lines** box | Type lines and press Enter (§7). |
| **Add Bulk Lines** | Start time, number of lines, minutes between them (3 by default) and the line type. `0730`, `10`, `3` gives 0730, 0733 ... 0757. An O&B gets its return leg. One Ctrl+Z takes the whole run off. A return type on its own is refused. |
| **Template / Copy previous day's lines** | Reuse a saved wave (`Templates\`), or copy the last flying day's times, types and IPs. |
| **Load a day...** | Load one day from MasterData's Schedule sheet, or from the last exported schedule (§9). Choose **Replace** or **Add**, and whether to keep the students and missions it names. |
| **Update Syllabi** | Bring syllabi in line with the day's flying schedule or the Big Board (§10). |
| **Add Student Absence** | Block students off (§11). |
| **Blue banner** | MasterData's Schedule sheet changed and has lines the board doesn't: load them or keep the board's lines. |
| **Menu bar** | Every action in one place: File, Edit, **Line** (acts on the selected line), Day, Fill, View, Help. |
| **Day tabs** | One tab per day. **Add day** builds several days at once; crew rest is checked across them. **Remove day** takes a day off. |
| **IP pucks** (left) | Drag onto a line's IP slot. The badge shows how many lines that IP has that day. Non-counter IPs are **purple** (NC); counters are blue. Badges: EXP (experienced), CHK (check IP), NC. |
| **Student pucks** (right) | Drag onto a student slot to hardcode. Badges: CAP, INTL, HOLD, U-SIM, U-FLT, ABS times, turns that day (green 1, amber 2, red 3+). Hover for next events, sim debt, discontinuity score and priority. |
| **Line details** (under the days) | The selected line's student against every rule, and **See options** (§8). |
| **Colors** (top right) | A key to every color on the board (also Help > What the colors mean). |
| **Fill / Clear fill / Re-fill / Export** (bottom) | §8. |

**Putting pucks on lines.** Two ways:
- **Drag and drop.** Drag a puck onto a slot. Drag between lines to move (hold **Ctrl** to copy). Drop back on its pool to take it off.
- **Click and place.** Click a puck to pick it up (thick outline; a dark banner says what you're holding), then click a slot. Hold **Ctrl** while clicking the slot to keep holding it for another line. **Esc** or **Let go** puts it down. **Double-click** a puck to put it straight on the selected line.

Nothing is checked while you drag (that used to be slow). When you drop a student, their rules are checked once and the line is selected so **Line details** shows the result.

**Colors, in short:**

| Where | Color | Meaning |
|---|---|---|
| Stripe at the left of a line | (by type) | Day Local, O&B, O&B Return, OFT / IFT, other. |
| Corner marker on a hardcoded student | green / amber / red / grey | Passes / passes with warnings / breaks a hard rule (OVERRIDE) / still checking. |
| Student slot | green | Hardcoded. |
| Student slot | white-grey | Open: the scheduler fills it. |
| Student slot | lilac | Return leg: follows its O&B. |
| Student slot | blue italics, `AUTO` | Put on by the scheduler (or its wingman). |
| Student slot | amber | Direct support for a formation. |
| Student slot | red, `NEED SP` | No student could fly the line. |
| Student puck | green / amber / grey | Available / has an unsat or absence / on hold. |

**Right-click** a hardcoded student for more: explain, take off, *Fly wing to...*, duplicate, delete line.

**The mission dropdown** on a hardcoded student:
- `(auto)` picks the best-scoring event that passes the rules (§12).
- Then the events the student can fly on this line.
- Then the next events they are **not opted for** yet, marked `(not opted)`.
- Then **Type a mission...** for anything else.

If you pick a mission that breaks a rule (not opted, already flown, not in the syllabus, wrong line type), FAIS explains why and offers the mission it would choose: **Yes** = use the recommended one, **No** = keep yours (flagged `OVERRIDE: Not opted for ...`), **Cancel** = no change.

## §7 Typing lines, keys and shortcuts

**The Add lines box:**

| Type this | You get |
|---|---|
| `0730 DL` (or `730 dl`, `7:30 local`) | Day Local at 0730 |
| `1015 O&B` | O&B, plus its **O&B Return** 3 hr later (same IP, linked) |
| `1015 OB 1330` | O&B with the return at 1330 |
| `1330 RTN` | Return line linked to the latest O&B before it that has no return |
| `1300 OFT x2` | Two OFT lines (IFT works the same) |
| `0800 DL Jones sp:Smith m:FT5103` | IP Jones, student Smith hardcoded on FT5103 |
| `0730 DL, 0745 DL, 0800 DL` | Several lines at once (`,` or `;`) |

Line type short names accepted: `DL`, `LOCAL`, `DAYLOCAL` → Day Local; `OB`, `O&B`, `ONB`, `O/B` → O&B; `RTN`, `RET`, `RETURN`, `R` → O&B Return; `OFT`; `IFT`. `SIM` is ambiguous and asks you to pick OFT or IFT.

Times accepted: `0730`, `730`, `7:30`, `07:30`, `07.30`, `0730L`, `7`.

**ETD box:** click an ETD, type a new time, press Enter. **Up/Down** = 5 min, **Shift+Up/Down** = 30 min, **Esc** cancels. Moving an O&B moves its return leg by the same amount.

**Keys:**

| Key | Does |
|---|---|
| Ctrl+L | Jump to the Add lines box |
| Ctrl+Z | Undo (up to 50 steps; a whole fill is one step) |
| Delete | Clear the selected slot |
| Ctrl+D / Ctrl+Delete | Duplicate / delete the selected line |
| Ctrl+PgUp / Ctrl+PgDn | Previous / next day |
| Esc | Let go of a puck you picked up |
| F5 | Re-check everything |
| Ctrl+Enter | Fill all days |
| Ctrl+Shift+Enter | Fill the day on screen |
| Arrow keys on a mission dropdown | Open it; arrows to move, Enter to pick |

## §8 Filling, reviewing and exporting

### Fill

- **Fill <day>** fills only the day on screen. **Fill all N days** (shown when there's more than one day) fills every day on the board.
- Before filling, FAIS checks the day for problems and asks *"Fill anyway?"* if it finds any: a line with no IP (it will be `NEED IP`), an IP or student not on MasterData, an O&B with no return line, a student pinned twice at the same time, a day with no lines.
- Only **open** lines are filled. Hardcoded students, and the picks of an earlier fill, stay where they are. So you can fill, change things, and fill again.
- The picks go **on the board** as ordinary lines (blue italics, `AUTO`). You can drag them, swap them, change missions.
- A line nobody can fly shows **NEED SP** in red.
- The status line sums it up: *"Filled Wed 21 Oct: the scheduler put 18 student(s) on (2 wingmen), 3 return legs, 1 line(s) with no student (NEED SP), 0 override flag(s)."*
- If filling this day makes a line on **another day you already exported** break a rule, the status line says *HEADS UP* and names it (§9).

### Clear fill, Re-fill, Undo

- **Clear fill** takes the scheduler's picks off the day on screen. Students you hardcoded, or picks you moved, stay.
- **Re-fill** = Clear fill, then Fill.
- **Ctrl+Z** takes the whole fill back off in one step.

### Review with Line details

Click anywhere on a line. The **Line details** panel shows:

- **Left:** the line's student against the **hard rules** (✖ broken), the **warnings**, and the **soft rules** (every + / − reason and the total). For an open line it says the scheduler will fill it (or that Export writes NEED SP).
- **Right:** **See options** ranks **every student** for that line: ✔ fits, ! fits with warnings, ✖ breaks a hard rule (with the first rule broken). Each shows the mission `(auto)` would pick and its score. Fits come first, best score first. **Hide rule breaks** hides the ✖ students. Click a student to see their full check on the left; double-click, or **Put on this line**, to hardcode them there.

See options is only run when you press it (it takes a while), and any change to the board clears the list.

### Export

**Export** writes the board **as it stands**. Nobody new is picked; every student stays where they are. Rule breaks are flagged `OVERRIDE`. A line with no student goes out as **NEED SP**.

It writes the day on screen **and** every day you filled since the board opened (the button names them). **All or nothing:** every file is backed up first and put back if any part fails.

| File | What changes |
|---|---|
| OutputSchedule: **MasterSchedule** (first tab) | The exported days are replaced. Other days already exported (today onward) are kept, so you can update one day at a time. |
| OutputSchedule: **each student's tab** | Only the Instructor, ETD and IsComplete cells of events that changed. Your own columns, notes and formatting stay. A student with no tab gets one. |
| MasterData: **Students** | SimDebt, DiscontinuityScore, and UnsatSim / UnsatFlight reset to FALSE where the schedule cleans them up. |
| MasterData: **String Tracker** | One row per student: 5+, 4, 3, 2 and 1-strings and unique counters (§15). |
| MasterData: **StudentAbsences** | Any absence Add Student Absence couldn't save earlier. |
| **ATIS.xlsx** (Desktop) | One tab per exported day (`Wednesday. October 21, 2026`): flights in ETD order, a thick border, then sims. Other days' tabs are kept. |

Afterwards a window says what was written (days, line counts, files, unsats cleared, OVERRIDE lines) or, if it failed, that **nothing** was written and why.

## §9 Changing a day you already exported

1. Start FAIS, click **Load a day...**, choose **Last exported schedule (OutputSchedule)** and the day. Everyone goes back on the board where they were exported, hardcoded.
2. Change what changed: take a student off, swap an IP, add a line.
3. Press **Fill <day>**, adjust, then **Export**.

Only the open lines are filled, and only that day is replaced in MasterSchedule. **Other exported days ride along as fixed context** (§12, §24): their students are treated as already booked, so nobody is double-booked or flown out of order. If your change breaks a line on another exported day (say you moved a student's FT5004 to today, and tomorrow's exported line has them on FT5004 too), the status line says which. Load and fill that day too.

## §10 Update Syllabi

**Update Syllabi** (toolbar) brings each student's syllabus in line with an official training file. Pick the file; FAIS works out which kind it is from its headings.

**The day's flying schedule** (tabs with Student, Mission and Line Status columns). Its **Flying** and **Sims** tabs are read:
- Line Status **CM** or **EXEC** → the mission is marked complete, with that line's Instructor and ETD (Start Time on Sims).
- Line Status **INC** or **CX** → the mission is made not complete.
- Anything else → nothing changes.
- The same student and mission more than once that day → a complete line wins.

**The Big Board** (a Student Name column, then a block of columns per event: date, status, grade, IP):
- Anything in the event's columns → complete, with the IP and the date as the ETD.
- Nothing there → not complete, **except** an event FAIS has planned for later than now (the Big Board can't show it yet).

**Left alone:** events not in the student's syllabus (most Big Board columns are academics), people who aren't FAIS students, and unsat flags.

**Name matching.** `Quill, Casey B. [2d Lt]`, `Quill, Casey`, `Quill` and `Casey B. Quill` are all the same student. Ranks, suffixes, case, periods and hyphens don't matter. Last name **and** first name must agree (an initial matches the name it starts). A last-name-only MasterData name matches only when exactly one person in the file has that last name. A name that could be more than one student is left alone and listed.

**Before saving** you see every change, student by student, including each SimDebt and Discontinuity change. **Yes** saves the changed tab cells in OutputSchedule and SimDebt / DiscontinuityScore in MasterData, both or neither. Students already on lines are re-checked against their new syllabi.

## §11 Add Student Absence

Blocks students off for a stretch of time (appointment, leave, a duty). The same as typing a line on MasterData's **StudentAbsences** sheet.

1. **Put students in the box:** drag from the list, click then click the box, double-click, or paste names (**Ctrl+V** or **Paste names**, one per line). Pasted names are matched like Update Syllabi.
2. **Pick the day** on the calendar.
3. **Type From and To**: `0730`, `1530`, `0330 PM`, `330pm`, `3:30 p.m.`. **on** is the end day, for an absence past midnight or over several days.
4. Optional **Reason** (shows on the puck and in the rule flag).
5. **Add absence.**

The absence applies at once: pucks show `ABS 0900-1130`, the scheduler keeps the student off any line it overlaps, and a line they're already on turns red. It is also saved to StudentAbsences (IsActive TRUE). If MasterData can't be saved, the absence holds for this run and Export writes it. An absence with the same start and end isn't added twice.

**How overlap is judged** (§13, §19): a sortie occupies **1.5 hr before** to **2.5 hr after** its ETD. A 1000 line conflicts with any absence touching 0830–1230.

---

# Part C: The logic (how FAIS decides)

## §12 How one line gets filled: the selection process

This is the heart of FAIS. Here it is in plain English first; §24 walks the same path through the code, and §32 tells it as the story of one line.

### The short version

For each open line, in ETD order:

```
for every student:
    run the HARD RULES for this student on this line
    if any rule fails → skip this student
    for every event the student could fly here:
        compute the SOFT RULE score
        keep the best (student, event) pair so far
put the best student on the line with that event
mark the event complete on that student (so later lines see it)
if it's a formation event → find a wingman (or direct support)
if it's an O&B → fill the return leg with the same student
if nobody passed → the line is "For Sale" (NEED SP on the board)
```

### The longer version

1. **Prepare.** Every line gets an id, its IP is looked up on the Instructors sheet (unknown → a placeholder; blank → `NEED IP`), and the callsign defaults to the IP's callsign. Placeholders like `For Sale` or `NEED SP` are cleared.
2. **Sort** the lines by ETD (lines without a valid time go last).
3. **Reserve pinned time.** Every hardcoded student gets a **reservation** for their line (and its O&B return). Reservations sit in the student's timeline like completed events, so an *earlier* open line won't give the student a turn that clashes with a *later* hardcoded one.
4. **Walk the lines in order.** For each line:
   - **Hardcoded line:** keep the student. Pick the mission (the one you chose, or the best-scoring one that passes). Mark it complete. Pair a wingman if it's a formation lead. Fill the return leg if it's an O&B. Any hard rule broken is written as `OVERRIDE: ...`.
   - **Return leg with no O&B student:** `For Sale` (*O&B Leg 1 Not Filled*). Return legs are only ever filled from their O&B.
   - **Open line:** test every student (step 5).
5. **Test every student** on the open line:
   - **Hard rules** (§13). Fail → out.
   - For each event that passed, compute the **score** (§14).
   - Keep the best by score, then priority, then name, then syllabus order (§20).
6. **Place the winner.** Write the student and mission on the line, mark the event complete on their syllabus (moving SimDebt and DiscontinuityScore, and clearing any unsat this line cleans up), add the notes. Then wingman / DS (§17) and return leg (§18).
7. **No winner** → `For Sale`, note *No Available SP*.

Because each pick is marked complete straight away, **later lines see it**: turn time, triple turns and sim debt all count it. That's why the order matters (§20).

### Where the event comes from

The scheduler never invents events. For a student and a line it takes:

1. the student's **opted** events (not complete, every prereq complete),
2. minus missions already **reserved** on another line,
3. that **fit the line type**:

| Line type | Event types allowed |
|---|---|
| Day Local | Single, SingleSolo, Formation, FormSolo, Both |
| O&B | Single, Formation, Both (and there must be a flyable event left for the return leg) |
| OFT | IFT, OFT |
| IFT | IFT |
| O&B Return | (filled only from its O&B) |

Then the hard rules check each of those events, and the soft rules score each one that passes. So one student can be scored on several events for the same line; the best (student, event) pair wins.

### When you hardcode

The same machinery runs on your pins, but **nothing is refused**:
- `(auto)` mission = the best-scoring event that passes. A wingman line only considers formation events; a lead with a paired wing prefers formation events.
- A typed mission is used as typed, even if it's not in the syllabus.
- Every hard rule broken goes in the note as `OVERRIDE: ...` and in the export summary.

## §13 Hard rules (pass / fail)

Hard rules decide **who is allowed** on a line. The scheduler never picks a student who breaks one. The board shows each broken rule as a red ✖ with a reason, and still lets you place the student (flagged `OVERRIDE`).

They're checked in this order. When filling, FAIS stops at the first failure (fast); on the board it checks them all (so it can list them).

### Student-level rules (about the student and the time)

| # | Rule (code) | Fails when | Example message |
|---|---|---|---|
| 1 | No ETD (`BAD_ETD`) | The line has no valid time. | `Line has no valid ETD ('')` |
| 2 | Return line (`RETURN_LINE`) | The line is an O&B Return. They're filled from their O&B, never on their own. | `O&B Return lines are filled from their O&B` |
| 3 | On hold (`ON_HOLD`) | `OnHold` is TRUE. | `Student is on hold` |
| 4 | Absence (`ABSENCE`) | An active absence overlaps **ETD − 1.5 hr to ETD + 2.5 hr**. | `Unavailable 0900-1130 (Dentist); sortie window 0830-1230` |
| 5 | CAP on O&B (`CAP_OB`) | CAP student on an O&B. | `CAP students can't fly O&Bs` |
| 6 | CAP single turn (`CAP_SINGLE_TURN`) | CAP student already flies a flight that day (sims don't count). | `CAP student already flies today (0730 FT5003)` |
| 7 | Triple turn (`TRIPLE_TURN`) | Already 2 events that day. | `Already has 2 events today (...) - triple turn` |
| 8 | O&B not only event (`OB_NOT_ONLY_EVENT`) | An O&B when the student already has an event that day. | `O&B must be the day's only event (has 0730 FT5003)` |
| 9 | Turn time (`TURN_TIME`) | Less than **3.25 hr** from another event that day. | `Turn time 2.8 hr < 3.25 hr minimum (vs 0730 FT5103)` |
| 10 | Duty day (`DUTY_DAY`) | More than **9 hr** between the day's events. | `Duty day: 9.5 hr between 0700 and 1630 events (> 9 hr max)` |
| 11 | Return-leg turn / duty (`RETURN_TURN_TIME`, `RETURN_DUTY_DAY`) | The same two checks for an O&B's return leg. | `Return leg 1315: turn time 3.0 hr < 3.25 hr (vs ...)` |
| 12 | Crew rest (`CREW_REST`) | Less than **15 hr** since yesterday's last event (ETD to ETD). | `Crew rest 13.5 hr < 15 hr since Tue 1730 FT5004` |
| 13 | Crew rest next day (`CREW_REST_NEXT`) | Less than **15 hr** before tomorrow's first event (multi-day boards). Only the day's last event is checked. | `Crew rest 14.0 hr < 15 hr before Thu 0700 ...` |
| 14 | Experienced IP (`EXP_IP`) | CAP or international student on a Day Local / O&B with an IP whose `IsExperienced` is FALSE. | `CAP student needs an experienced IP (Lee isn't)` |
| 15 | Unsat sim (`UNSAT_SIM`) | `UnsatSim` TRUE and the line is a Day Local / O&B. | `Unsat sim on record - no Day Local/O&B until a re-sim in an OFT` |
| 16 | Unsat flight (`UNSAT_FLIGHT`) | `UnsatFlight` TRUE and the line is an OFT / IFT. | `Unsat flight on record - no OFT/IFT until re-fly` |
| 17 | No eligible event (`NO_ELIGIBLE_EVENT`) | Nothing in the syllabus is opted and fits this line type. | `No eligible Day Local events (next: FE4002 OFT, ...)` |

**Warnings** (shown amber, never block):
- Turn time 3.25–3.5 hr: `Less than 3.5 hr TT, Flt/CC Approval Required`.
- 6–9 hr between events: `If this is a Night Flight SP will not have Duty Day`.
- The same two for a return leg.

### Event-level rules (about one event on this line)

Each candidate event is checked too. A student passes only if at least one event passes.

| Rule (code) | Fails when |
|---|---|
| `NOT_OPTED` | You chose an event the student isn't opted for (already complete, or a prereq isn't). *e.g.* `Not opted for FE3102: FT5001 isn't complete yet`. |
| `MISSION_NOT_ELIGIBLE` | You chose an event that can't go on this line (wrong type, pinned elsewhere, no O&B second leg). |
| `MISSION_NOT_FOUND` | The mission you typed isn't in the student's syllabus. |
| `FORM_ON_OB` | A formation event on an O&B. |
| `CHECK_ON_OB` | A check ride on an O&B. |
| `OB_NO_SECOND_LEG` | An O&B event with no flyable event left for the return. |
| `OB_RETURN_CHECK` | The return leg would be a check ride. |
| `CHECK_SAME_IP` | Check ride with an IP who already flew with the student (evaluator independence). |
| `CHECK_NEEDS_CHECK_IP` | Check ride with an IP whose `IsCheck` is FALSE. |
| `FORM_NO_WING_LINE` | A formation lead with no open Day Local line within **FormLookAhead** minutes (22 by default) for a wingman or DS. |
| `FORM_PREFER_SINGLE` | *(Advisory, doesn't block)* No other student can fly formation here, but another can fly single-ship. |

**Notes** added to lines (no effect on the pick): `Direct Support Required`, `Solo Flight`, `Need IP in Continuity to Brief` (solo whose IP hasn't flown the student in their last 5 flights).

> **Another way to see it.** Hard rules answer *"can this student legally fly this line?"*. Soft rules (§14) answer *"of everyone who can, who should?"*.

## §14 Soft rules (the score)

Every student who passes the hard rules gets a **score** for each event they could fly. Higher wins. Each rule adds or takes away points; the values come from MasterData's **SoftWeights** sheet (first data row). The board lists every rule that scored as a + / − line, for example `3 String (zipper over Lee) +25`.

| Rule (board label) | SoftWeights column | Points | When |
|---|---|---|---|
| **Dollar Ride Bonus** | `DollarRideBonus` | +weight | Mission is FT5001. Gets new students flying. |
| **Sim Debt (flight)** | (student's SimDebt) | **− SimDebt** | Mission is a flight (FT5, FF5, AA5). A student deep in sim debt is pushed toward sims instead. A negative debt (sim credit) *adds* points. |
| **Sim Debt (sim)** | (student's SimDebt) | **+ SimDebt** | Mission is a sim (FI3, FE3, FI4, FE4, FN3, FN4, AA4). Students in debt are pulled into sims. |
| **CAP Bonus** | `CAPBonus` | +weight | CAP student (they can only fly once a day, so they get a boost). |
| **Continue String** | `ContString` | +weight | The ride makes a 2- or 3-string with this IP (a zipper return included) (§15). |
| **Discontinuity (string)** | (student's DiscontinuityScore) | + score | Added on top of Continue String only (§16). |
| **CAP Bonus (string)** | `CAPBonus` | +weight | A CAP student continuing a string gets the CAP bonus a second time. |
| **String 4** | `String4` | weight (default −10) | The ride makes a 4-string. |
| **String 5** | `String5` | weight (default −100) | The ride makes a 5+ string. |
| **String Swap** | — | 0 | Ends a 2-string to make a new 2-string. Even. |
| **Break String** | `BreakString` | weight × strings broken | Really ends a 2-string (two different counters in a row). |
| **Zipper** | `Zipper` | weight | The zipper ride itself, or alternating counters (§15). |
| **Non-Counter** | — | 0 | The IP is a non-counter (and didn't fly the ride before). |
| **Counter Continuity** | `CounterContinuity` | weight (default 5) | The IP already counts for this student in this phase. |
| **New Counter** | `NewCounter` (optional) | weight (default 0) | A new counter. Shown only when it scores or reaches the phase limit. |
| **OFT/IFT Mismatch** | `OFTIFTPen` | weight | An IFT event in an OFT slot. |
| **Exhaustion** | `ExhaustMult`, `ExhaustDays` | (2 × flights + sims) × mult | Events in the last `ExhaustDays` days (default 5, mult 1). Use a negative mult to penalise busy students. |
| **BIT Tracker** | `BITPerDay`, `BITDays` | days off × BITPerDay (× 2 for CAP) | Flight lines only. Days without a flight (not counting the day before), up to `BITDays`. |
| **BIT Tracker (capped)** | same | BITPerDay × BITDays ÷ 2 | Past the BIT limit; the line gets a note `BIT! n days without a flight, consider an 86`. |
| **Checkride Priority** | — | +50 | A check ride with a check IP. Note `Checkride! Requires Formal`. |
| **Student Priority** | (student's `Priority`) | + priority | Always. Your manual nudge. |

**The values on the current MasterData** (`SoftWeights` row 2) are, for reference:

```
ContString 25   String4 -20   String5 -60   Zipper -3   NonCounter 5 (no longer used)
BreakString -10 ExhaustMult -0.25  ExhaustDays 5   BITPerDay 6   BITDays 5
CAPBonus 5      FormLookAhead 22   LevelDifference -10   CounterContinuity 5
OFTIFTPen -10   DollarRideBonus 10 TransSimRatio 2   FormSimRatio 3.3333   SimStopMult 3
```

Optional columns you can add: `NewCounter`, `Disc1String`, `DiscCounter`, `Disc3PlusString`. A blank or missing cell uses the default in the table above (0 if none is listed).

> **The same rules, grouped by purpose:**
> - **Keep training moving:** Dollar Ride Bonus, BIT Tracker, Checkride Priority, CAP Bonus.
> - **Instructor continuity (AETCI):** Continue String, String 4/5, Break String, Zipper, String Swap, Counter Continuity, New Counter, Discontinuity.
> - **Balance sims and flights:** Sim Debt (flight / sim), OFT/IFT Mismatch.
> - **Don't overwork anyone:** Exhaustion.
> - **Your override:** Student Priority.

§33 works a full example by hand.

## §15 Strings and counters

FAIS follows the AETCI instructor continuity rules. This section explains them three ways: plain rules, examples, and how the code counts.

### The rules in plain English

- **An IP who flies a sortie with a student is a counter**, and that sortie is part of a **string**.
- **Three sorties in a row with one counter is the goal.** 2 or 4 is acceptable (to prevent BITs); 5 or more is not.
- **Each category (FT, FF, AA) is its own string**, counted in mission order (FT5001, FT5002, ...), even if rides from another category were flown in between.
- **One different counter in between doesn't break a string.** `Lee, Lee, Ortiz, Lee` is a 3-string with Lee (a **zipper** over Ortiz). Two different counters in a row end it.
- **Non-counters** (`IsCounter` FALSE, usually supervisors) score 0, break no strings and aren't counters, **unless** they fly the student two rides in a row: then they count for that student like anyone else.
- **Sims, solos and check rides are left out** entirely.
- **Counter limit:** 6 counters in Fundamental (FT, FF) and 3 in Airmanship (AA). FAIS shows where a student stands (`Diaz is a new counter (4 of 3 Airmanship counters)`) but doesn't enforce it.

### The rules as examples

History is the student's past rides in one category, oldest first. "+ X" is the ride being scored.

| History | Next ride | What it is | Score |
|---|---|---|---|
| Lee | + Lee | 2-string | Continue String (+ Discontinuity) |
| Lee, Lee | + Lee | 3-string | Continue String (+ Discontinuity) |
| Lee, Lee, Lee | + Lee | 4-string | String 4 |
| Lee ×4 | + Lee | 5-string | String 5 |
| Lee, Lee | + Ortiz | Zipper ride: Lee's 2-string waits on the next ride | Zipper only (not Break String) |
| Lee, Lee, Ortiz | + Lee | Zipper return: 3-string with Lee | Continue String |
| Lee, Lee, Ortiz | + Smith | Two different counters in a row: Lee's 2-string ends | Break String |
| Lee, Lee, Ortiz | + Ortiz | Ends Lee's 2-string to make Ortiz's 2-string | String Swap (0) |
| Lee, Ortiz, Lee | + Ortiz | Alternating counters: not a string | Zipper only (no Continue String) |
| Lee, Lee, Lee | + Ortiz | Lee's string is already 3: nothing to lose | (no Break String) |
| Lee | + Avara (non-counter) | Non-counter | Non-Counter (0) |
| Lee, Avara | + Avara | Non-counter two in a row: now counts | Scored like a counter's 2-string |

On top of those, **Counter Continuity** (+5) applies whenever the IP already counts in this phase, and **New Counter** whenever they don't.

### How the String Tracker counts

At Export, FAIS writes one row per student to MasterData's **String Tracker**: number of 5+, 4, 3, 2 and 1-strings, and unique counters. The counting:

1. Take the student's completed rides that count (`Test-StringRide`: a flight, not a sim, solo or check ride, with an IP).
2. Group them by category (FT, FF, AA) and sort by mission number.
3. Work out who the student's non-counters are: the Instructors sheet's non-counters, minus any who flew the student two rides back to back.
4. For each category, measure every string (below), and count a string as *finished* at a ride when that IP's next ride starts over at 1, or the IP never flies the category again.
5. Counters = the unique IPs who aren't non-counters.

### How the code measures one string

`Measure-String` (in `Syllabus.psm1`) walks back from the newest ride one step at a time:

```
IP1 = the newest ride's IP
step back one ride (IPn):
    IPn is IP1?             → +1, and forget any skip
    IPn is a non-counter?   → skip it (doesn't break anything)
    IPn is a different counter, and we haven't skipped one yet?
                            → skip it (the zipper), remember we skipped
    otherwise (second different counter, or no more rides)
                            → stop
```

So `A, A, B, A, A, A` measured from the end is 5: A A A, skip B, A A.

`Get-StringOutlook` uses this to answer "what does one more ride with IP X do?": the new length, whether it's a zipper, whether the pattern is alternating, which other IPs' strings it **breaks** (they could no longer reach 2) and which it **bridges** (still alive, waiting on the next ride). `Get-StringPlacement` adds the counter information (non-counter? new counter? count vs limit). `SoftRules` turns that into points (§14).

## §16 Sim debt, discontinuity, exhaustion and BIT

These four numbers are what make FAIS "remember" a student's recent history.

### Sim debt

**Why:** students need sims at the right pace. If they fly too far ahead, they hit a *sim stop* and can't fly until they catch up.

**How it moves:** every event marked complete changes the student's SimDebt by:

| Event | Change |
|---|---|
| Flight (FT5, FF5, AA5) | **+ SimStopMult** (3 on the current sheet) |
| Transition sim (FI3, FE3, FI4, FE4) | **− TransSimRatio × SimStopMult** (2 × 3 = 6) |
| Formation sim (FN3, FN4, AA4) | **− FormSimRatio × SimStopMult** (3.3333 × 3 = 10) |

**How it scores:** a flight scores −SimDebt, a sim scores +SimDebt. A student 12 in debt loses 12 points on every flight and gains 12 on every sim, so sims win.

**Where it's kept:** the `SimDebt` column of MasterData's Students sheet. Read at start, moved by every scheduled event (hardcoded lines, picks, wingmen, return legs), by logged incompletes, by re-planning, and by Update Syllabi; written back on Export, Log Incomplete Event and Update Syllabi. Taking an event back only takes back *its own* change, so a number you typed in by hand is kept and moved from there.

### Discontinuity score

**Why:** two students could both stay with an IP. The one whose training has been more broken up (lots of different IPs) should get to.

**How it's counted** (from the same strings and counters as the String Tracker):

| Each | Points (SoftWeights column, default) |
|---|---|
| 1-string | `Disc1String`, +1 |
| unique counter | `DiscCounter`, +1 |
| 3+ string (3, 4, 5+) | `Disc3PlusString`, −1 |

**How it scores:** only on a **Continue String** ride, as its own line: `Discontinuity 4, 2 String +4`. Negative scores take points off. String 4/5, String Swap, Zipper, Break String and Non-Counter lines don't get it.

**Where it's kept:** like SimDebt, in the `DiscontinuityScore` column (FAIS adds the heading the first time). A blank cell is counted from the syllabus at start and written on Export; clear a cell to have FAIS recount it. Every ride that changes the student's strings moves it by exactly what that ride changed.

### Exhaustion

Counts the student's completed events in the last `ExhaustDays` days (5) before this line: **(2 × flights + sims) × ExhaustMult**. With the current −0.25, two flights and two sims cost (4 + 2) × −0.25 = −1.5. The board shows `2 Flights and 2 Sims in last 5 Days`.

### BIT tracker

On flight lines only. **Days off** = days between the last flight and this line, minus one (flying yesterday = 0 days off).
- 1 to `BITDays` days: **days × BITPerDay** (doubled for CAP). With the current sheet, 3 days off = 3 × 6 = +18.
- More than `BITDays`: capped at **BITPerDay × BITDays ÷ 2**, and the line gets `BIT! n days without a flight, consider an 86`.

## §17 Formations, wingmen and direct support

When a student is placed on a **formation** event (type contains "Form") as a lead:

1. **Callsigns.** The lead takes the next pair on the **FormCallsigns** sheet (`HAZE 11` lead, `HAZE 12` wing). Each formation uses two entries; the list restarts at the first pair **each day**.
2. **Look for open lines** within **FormLookAhead** minutes (22) of the lead's ETD that a wingman can fly: open local flying lines (Day Local). Never an OFT / IFT, an O&B or a return leg.
3. **Look for a student wingman** on those lines: anyone (not the lead) with an opted formation event who passes the hard rules there. Score each with the soft rules, plus **(wing level − lead level) × LevelDifference** (from the syllabus `Level` column). With LevelDifference −10, a wingman far ahead of the lead loses points. The best becomes the wingman (`Wing`), noted `Wingman to <lead>`.
4. **No student wingman →** the soonest open line becomes **Direct Support**: student `Need Form SP`, noted `Direct Support for <lead>`.
5. **No open line at all →** the lead gets `Wingman: none found`.

On the board you can also pair a wing yourself: right-click a hardcoded student and choose **Fly wing to...**. A paired wing satisfies the lead's formation check.

The hard rule `FORM_NO_WING_LINE` (§13) stops the scheduler choosing a formation event when there's no line for a wing at all. `FORM_PREFER_SINGLE` warns when nobody else can fly formation but someone could fly single-ship.

## §18 O&Bs and return legs

An **O&B** (out-and-back) is two lines: the O&B and its **O&B Return** (by default 3 hr later, same IP). The student flies both.

- **On the board:** typing `1015 O&B` makes both lines, linked. Moving the O&B moves the return; deleting it deletes the return; changing a line to O&B adds a return. The return's student slot just says it follows its O&B.
- **Hard rules for the O&B:** it must be the day's only event, no CAP students, no formation or check events, and there must be a flyable event left for the return leg that isn't a check ride. Turn time and duty day are checked for the return leg too.
- **When the O&B is filled** (by you or the scheduler), `Schedule-ReturnLeg` puts the same student and IP on the return, and gives it the student's next valid event in syllabus order.
- **A return whose O&B is open** goes out `For Sale` (*O&B Leg 1 Not Filled*).
- **Which return belongs to which O&B:** by the link first; for lines from old sheets with no link, the return on the same day with the same IP that's still open (preferring one after the O&B).

## §19 Unsats, holds and absences

**Unsats** (logged with *Log Hooked Event*, or TRUE in MasterData):

| Hooked | Blocked | Cleaned up by |
|---|---|---|
| OFT | Day Local and O&B | an **OFT** |
| IFT | Day Local and O&B | an **OFT or IFT** |
| Flight | OFT and IFT | any **flight** |

MasterData only stores TRUE / FALSE, so for an unsat typed there FAIS takes the hooked sim to be the student's last sim flown (an OFT if unknown).

As soon as a line that cleans the unsat is scheduled, the unsat is cleared for every line after it, and the line's Notes say `Clears Unsat Sim (OFT)` (or IFT, or `Clears Unsat Flight`). Export writes FALSE to MasterData. If you re-plan a day whose schedule cleaned up an unsat, the unsat comes back until the new schedule cleans it up again.

**Hold:** a student on hold breaks the `ON_HOLD` rule on every line.

**Absences:** come from MasterData's StudentAbsences sheet (active rows only, IsActive TRUE) or from Add Student Absence. A name must match the student's name **exactly** (case aside); a name that matches no one is reported in the console. A line clashes when the absence overlaps ETD − 1.5 hr to ETD + 2.5 hr.

## §20 Tie-breaks and the order lines are filled in

**Order.** Lines are filled in **ETD order**, across every day being filled. Each pick is marked complete straight away, so it counts toward later lines' turn time, triple turns, crew rest and sim debt.

**Why reservations exist.** If a student is hardcoded at 1400, filling the 1100 line first would otherwise not know about it. The reservation puts 1400 in the student's timeline before anything is filled, so the 1100 line won't give them a 2.5 hr turn.

**Tie-breaks**, when two (student, event) pairs score the same:
1. higher **Priority**,
2. then **name** alphabetically,
3. then the event earlier in the syllabus.

This makes runs repeatable: the same board and the same data always give the same fill.

**Formation callsign counter** resets at the start of each day.

---

# Part D: How the code works

## §21 Architecture at a glance

FAIS is plain PowerShell modules plus WPF XAML files. No compiled code, no database.

```
             ┌──────────────────────────────────────────────────────────┐
             │ FAIS.ps1 (entry point, settings, run order)              │
             └──────────────────────────────────────────────────────────┘
                 │ imports, in order
                 ▼
 DataManipulation ─► ExcelFiles ─► Syllabus ─► DataDownload ─► SchedulingEngine
   (helpers)          (Excel COM)   (syllabus,   (read Excel)    (hard rules,
                                     strings,                      soft rules,
                                     sim debt)                     fill)
                 ─► DataUpload ─► SyllabusImport ─► PreRunMenu ─► BoardModel
                    (write Excel)  (Update Syllabi) (WinForms     (board data,
                                                     pop-ups)      no UI)
                 ─► Assessment ─► PuckBoard (+ .xaml files)
                    ("what if"     (the WPF window)
                     for the board)
```

**Layers:**

| Layer | Modules | Knows about |
|---|---|---|
| Helpers | DataManipulation, ExcelFiles | Values, names, dates, Excel files |
| Domain | Syllabus, SchedulingEngine | Students, events, rules, scores. No UI, no files. |
| Data in / out | DataDownload, DataUpload, SyllabusImport | Excel sheets ↔ objects |
| Board model | BoardModel, Assessment | Days, lines, pins, undo, autosave. No WPF, so it's testable. |
| UI | PreRunMenu (WinForms), PuckBoard (WPF) | Windows, buttons, drag and drop |

The engine never touches Excel and the board never makes scheduling decisions itself: the board turns its lines into **schedule rows**, hands them to the engine, and turns the result back into lines.

## §22 Module by module

### `FAIS.ps1` (entry point)
1. Checks FullLanguage mode and STA (restarts itself with `-STA` if needed).
2. Settings: paths, `$BoardSettings`, `$OutputSettings` (§30).
3. Imports ImportExcel, then the modules **in order** (later ones call earlier ones).
4. Saves open workbooks in Excel, then `IngestData`. Stops if OutputSchedule can't be opened (carrying on would start every student from a blank syllabus) or SoftWeights can't be read.
5. `DownloadPrevSchedule` and `Invoke-PreRunMenu`.
6. `Show-PuckBoard`. If the board returns anything but Export, stops ("nothing was written").
7. `Export-FaisOutput`, then `Show-FaisExportResult`.

### `DataManipulation.psm1` (shared helpers)
- `ConvertTo-FaisBool`: Excel gives `TRUE`, `"FALSE"`, 1, blanks... and `"FALSE"` as text is *true* in PowerShell. Every flag goes through here.
- `Get-NameText`, `Get-BaseName`: a name whether the field holds an object, a string or nothing; the base name without `[rank]`.
- `Parse-ETD`, `ConvertTo-FaisTimeValue`: dates, times, Excel serial numbers, `0730`, `7:30 PM`...
- `Resolve-SortieInstructor`, `Get-MatchedInstructor`: a line's IP as an instructor object (placeholder `NEED IP` when blank).
- `Resolve-StudentByName`: exact name, then a last name only one student has.
- `Test-IsSimSortieType`, `Test-IsReturnSortieType`, `Test-IsOBSortieType`, `Test-IsWingCapableType`, `ConvertTo-CanonicalSortieType`: what a line type is.
- `Test-IsSentinelStudent`: `For Sale`, `NEED SP`, `Need Form SP`, `NEED IP`, `SIM IP` aren't real students.
- `Copy-StudentForEval`, `Copy-StudentList`: deep-enough copies so "what if" never changes the real students.

### `ExcelFiles.psm1` (keeping Excel out of the way)
Talks to a running Excel through COM (Windows only; does nothing elsewhere). `Close-FaisExcelFiles -SaveOnly` saves a workbook before reading; `Request-FaisFileWritable` saves and closes before writing (and shows the "file is open" box only if Excel can't be reached); `Open-FaisWorkbooks` reopens afterwards.

### `Syllabus.psm1` (the student's syllabus)
- **Lookups:** `Test-MissionComplete`, `Find-SyllabusMission`, `Get-SyllabusSortKey`, `Get-MissionKind` (Sim for OFT/IFT types).
- **Eligibility:** `Get-OptedEvents` (incomplete with prereqs met), `Get-AllowedEventTypes` (by line type), `Get-ValidOptedEvents` (both, minus reserved, plus the O&B look-ahead).
- **Completing:** `Complete-StudentMission` is *the one place* an event is marked complete: sets IsComplete, IP, ETD; moves SimDebt once; moves DiscontinuityScore by what the ride changed; clears an unsat the line cleans up. `Undo-StudentMission` reverses exactly that. `Reset-PlannedCompletions` undoes events dated from the board's first day on (they were plans from an earlier run, not flown events).
- **Reservations:** `Add-StudentReservation`, `Remove-StudentReservation`, `Get-ReservedMissionNames`.
- **Timeline:** `Get-StudentTimeline`: completed events plus reservations, oldest first. The time-based hard rules read this.
- **Strings:** `Measure-String`, `Get-StringLength(s)`, `Get-StringCounts`, `Get-StringFlights`, `Get-StudentNonCounters`, `Get-StringOutlook`, `Get-StringPlacement` (§15).
- **Discontinuity:** `Get-DiscontinuityScore`, `Add-StudentDiscontinuity`, `Initialize-StudentDiscontinuity` (§16).
- **Unsats:** `Clear-CleanedUnsat`, `Restore-PlannedUnsat`, `Get-LastSimType`.

### `DataDownload.psm1` (reading Excel)
- `IngestData`: Instructors, Students (first 100) with each student's syllabus from their OutputSchedule tab (or MasterData's Syllabus sheet if they have none), the Schedule sheet, FormCallsigns, SoftWeights, StudentAbsences. Returns one object.
- `Open-FaisWorkbookForRead`: reads a workbook even while Excel has it open.
- `Read-ScheduleSheet`, `Get-ScheduleSheetDays`, `Get-ScheduleDayFingerprint`: the Schedule sheet, headings found by name, blank rows skipped, a fingerprint per day so the board can tell when the sheet changed.
- `DownloadPrevSchedule`: the last MasterSchedule with students, missions and IPs linked back to objects (for the pre-run pop-ups and for context rows).

### `SchedulingEngine.psm1` (the decisions)
- `Test-HardRules` (§13): returns `{Passed, Violations, Notes, PassedEvents, EventResults}`. Never changes anything. `-Explain` checks every rule (for the board); without it, stops at the first failure (for the fill).
- `SoftRules` (§14): one student, one event, one line → a score. `-Detailed` adds the breakdown and the reason text the board shows.
- `Get-EventScore`: scores on a *copy* of the line so the real line is never half-assigned.
- `Test-CandidateBetter`: the tie-break (§20).
- `Invoke-FillSchedule`: the main loop (§12, §24).
- `Invoke-OpenLine`, `Invoke-PinnedLine`, `Select-PinnedMission`, `Process-HardcodedSortie`: one line each way.
- `Pair-FormationWingman` (§17), `Schedule-ReturnLeg` (§18).
- `Trace-SchedulingDecision`: the `-Trace` output.

### `DataUpload.psm1` (writing Excel)
- `ConvertTo-ExportRows`: rows as plain text.
- `Read-MasterScheduleRows`, `Merge-MasterScheduleRows`: one day at a time; other days kept.
- `Export-FormattedMasterSchedule`, `Apply-GroupBorders`, `Set-MasterScheduleFirst`.
- `Export-StudentSyllabusTabs`, `Update-SyllabusTabCells`: in-place cell updates.
- `Update-MasterDataStudents`, `Get-UnsatFlagUpdates`, `Update-StringTracker`.
- `Add-MasterDataAbsences`, `Get-PendingStudentAbsences`, `Save-StudentAbsences`.
- `Export-AtisWorkbook`, `Export-DailyScheduleSheets` (optional Teams copy).
- `Invoke-WithFileBackup`: copies every target file to Temp, runs the writes, and copies them all back if anything throws. This is what makes Export **all or nothing**.
- `Export-FaisOutput`: ties it together (§26).

### `SyllabusImport.psm1` (Update Syllabi)
Reads the training file, detects its kind from headings (`Find-DailyScheduleHeadings`, `Find-BigBoardHeadings`), matches names (`ConvertTo-PersonName`, `Get-PersonNameMatch`, `New-PersonIndex`, `Find-PersonMatch`), builds a plan (`Get-SyllabusUpdatePlan`), shows it (`Format-SyllabusUpdateSummary`) and saves it (`Save-SyllabusUpdate`) (§10).

### `PreRunMenu.psm1` (pop-ups)
WinForms menus: `Show-PreRunMenu`, `Show-EventSelectionMenu`, `Show-HoldStatusMenu`, then `Process-HookedEvents`, `Process-IncompleteEvents`, `Update-StudentHoldStatus`, `Save-PreRunChanges`. `Invoke-PreRunMenu` loops until Open Puck Board or Exit (§5).

### `BoardModel.psm1` (the board's data, no UI)
- **Board and lines:** `New-Board`, `New-BoardLine`, `Get-BoardDayLines`. A board is `{Days, Lines, Version, Undo, Settings, SheetSync}`.
- **Undo:** each edit saves a JSON snapshot first (`Save-BoardUndoPoint`, up to 50). `Version` always goes up, so cached checks are never reused after a change.
- **Edits:** add/remove days and lines, ETD, type, IP, pin, move, mission, formation pair. O&B rules (return follows) live here.
- **Typing:** `ConvertTo-BoardTime`, `ConvertTo-BoardSortieType`, `ConvertFrom-QuickEntry`, `New-BulkBoardLines`.
- **Absences:** `ConvertTo-ClockTime`, `Resolve-PastedStudents`, `New-StudentAbsences`, `Add-StudentAbsences`.
- **Engine hand-off:** `ConvertTo-ScheduleRows` (lines → engine rows), `ConvertFrom-ScheduleRows` (rows → lines), `Set-BoardFill` / `Clear-BoardFill`, `Test-BoardReadyToFill`.
- **One day at a time:** reading sheet days, banner logic, `Get-PublishedContextRows` (exported days as fixed context), `Split-FillResult` (separate the filled days from the context).
- **Files:** `Save-BoardState` / `Open-BoardState` (one JSON per day in `Boards\`, written to a temp file then renamed so a crash never leaves half a file), templates, `Copy-BoardPreviousDay`, `Get-NextFlyingDay` (next Monday–Friday).

### `Assessment.psm1` ("what if" for the board)
`Get-PlacementAssessment` runs the **real** hard and soft rules for one student on one board line, against the board's other pins, on copies of the students. Status is `Clean`, `Warn`, `Hard` or `Error` (that's the corner marker color). Results are cached per board version. Also the mission dropdown texts (`Get-MissionChoiceTexts`, `Get-MissionAdvice`) and puck tooltips (`Get-StudentPuckInfo`, `Get-InstructorPuckInfo`).

### `PuckBoard.psm1` + XAML files (the window)
§25.

### `OldFunctions.psm1`
Legacy code kept for reference. **Not loaded.**

## §23 The data objects

These are the shapes passed between modules. Knowing them makes the code much easier to read.

**Student** (built by `IngestData`):

```
Name, ID, IsCap, IsInternational, Priority
OnHold, UnsatSim, UnsatSimType ('OFT'/'IFT'), UnsatFlight
  (+ ...InMasterData copies: what MasterData holds, so Export only writes changes)
SimDebt, DiscontinuityScore
Syllabus      = [ syllabus entries ]
Absences      = [ {Student, StartTime, EndTime, Reason} ]
Reservations  = [ {LineId, Start, SortieType, Mission, Instructor, IsSim} ]
SyllabusRevision (goes up on every change, for caching)
```

**Syllabus entry:**

```
SyllabusID, Mission ('FT5003'), Prereqs ('FT5002' or 'FE4001, FT5005' or 0),
Type ('Single','SingleSolo','Formation','FormSolo','Both','OFT','IFT'),
IsCheck, IsComplete, Instructor, ETD ('yyyy-MM-dd HH:mm:ss'), NumOpting, Level
```

**Instructor:**

```
Name ('Lee, Pat A. [Capt]'), IsCounter (blank = TRUE), IsExperienced, IsCheck,
Callsign, Num, OSC, OSNum
```

**Schedule row** (what the engine works on; `ConvertTo-ScheduleRows` makes these from board lines):

```
ETD, Instructor (name, then an instructor object), SortieType,
Student (name, then a student object), Mission (name, then a syllabus entry),
Callsign, Num, Notes,
BoardLineId, LinkedLineId (O&B → its return), WingOfLineId (wing → its lead),
FilledBy: Pin | Engine | Wing | Return | DS | ForSale | NeedSP
```

**Board line** (what the board stores and autosaves):

```
Id, Day ('yyyy-MM-dd'), ETD, SortieType, Instructor, Student, Mission,
Callsign, Num, Notes, LinkedLineId, WingOfLineId,
Source: Manual | Sheet | Exported | Copied | Template ...
Filled: '' (typed/dropped) | Engine | Wing | ForSale | DS   (what the last fill did)
```

**Hard rule result** (`Test-HardRules`):

```
Passed       (no student-level block AND at least one event passed)
Violations   [ {Rule, Severity 'Hard'|'Advisory', Message, Mission, Blocks} ]
Notes        [ warnings ]
PassedEvents [ syllabus entries ]
EventResults [ {Event, Mission, Passed, Violations, Notes} ]
```

**Soft rule result** (`SoftRules -Detailed`):

```
Score, Breakdown (label → points), Reasons (label → text), Notes, Student, Mission, ETD
```

**Fill result** (`Invoke-FillSchedule`, then `Split-FillResult`):

```
Schedule (the filled days' rows), Context (exported days that rode along),
Students (the copies, with picks marked complete), Overrides, ContextOverrides,
Stats {Lines, Pinned, Filled, Wing, Return, DS, ForSale, NeedSP, Overrides}, Log, Days
```

## §24 Life of a run, through the code

The same run as §4, but following the function calls.

```
FAIS.ps1
 ├─ Close-FaisExcelFiles -SaveOnly            save open workbooks so edits are read
 ├─ IngestData                                → Instructors, Students (+ syllabi), ScheduleSheet,
 │    └─ Initialize-StudentDiscontinuity        Formations, SoftWeights, absences attached
 ├─ DownloadPrevSchedule                      → the last MasterSchedule, linked to objects
 ├─ Invoke-PreRunMenu                         loop: Hook / Incomplete / Hold / Absence
 │    ├─ Process-HookedEvents → Save-PreRunChanges        (UnsatSim / UnsatFlight)
 │    ├─ Process-IncompleteEvents → Undo-StudentMission → Save-PreRunChanges
 │    └─ Update-StudentHoldStatus → Save-PreRunChanges
 ├─ Show-PuckBoard
 │    ├─ Open-BoardState (Boards\*.json)       your autosaved lines
 │    ├─ Reset-PlannedCompletions              events planned from the first day on → not complete
 │    ├─ load Schedule sheet days, start timers (assess, autosave, sheet watch)
 │    ├─ ... you edit: every change → Request-PBCommand → BoardModel edit → Save-BoardUndoPoint
 │    ├─ Fill → Invoke-PBFill
 │    │    ├─ Test-BoardReadyToFill            "Fill anyway?"
 │    │    ├─ Invoke-PBScheduler
 │    │    │    ├─ ConvertTo-ScheduleRows       board lines → engine rows
 │    │    │    ├─ Get-PublishedContextRows     other exported days as fixed rows
 │    │    │    ├─ Copy-StudentList             work on copies
 │    │    │    ├─ Invoke-FillSchedule          ← §12
 │    │    │    └─ Split-FillResult
 │    │    └─ Set-BoardFill                     picks → board lines (Filled = Engine/Wing)
 │    └─ Export → Invoke-PBExport
 │         └─ Invoke-PBScheduler -NoAutoFill    same pass, but open lines → NEED SP
 ├─ Get-UnsatFlagUpdates, Get-PendingStudentAbsences
 ├─ Export-FaisOutput                          ← §26
 └─ Show-FaisExportResult
```

**Inside `Invoke-FillSchedule`:**

```
prep every row (ids, IP objects, callsigns, clear placeholders)
sort by ETD
for every pinned row: Add-StudentReservation (and for its O&B return)
for every row in order:
    new day → FormIndex = 0
    already filled (e.g. a return leg or wingman done earlier) → skip
    pinned   → Invoke-PinnedLine
                 Remove-StudentReservation (this line's own hold)
                 Select-PinnedMission → Test-HardRules -Explain, best-scoring passing event
                 Process-HardcodedSortie → Complete-StudentMission,
                                           Pair-FormationWingman (lead), Schedule-ReturnLeg (O&B)
                 hard violations → "OVERRIDE: ..." note + Overrides list
    -NoAutoFill → NEED SP
    return leg  → For Sale (its O&B wasn't filled)
    open     → Invoke-OpenLine
                 for each student: Test-HardRules; for each passed event: Get-EventScore
                 best by Test-CandidateBetter
                 → Complete-StudentMission, notes, Pair-FormationWingman, Schedule-ReturnLeg
                 nobody → For Sale
count FilledBy into Stats
```

**Why Export runs the engine again with `-NoAutoFill`:** the board stores names. Export needs the full picture: missions resolved, events marked complete on the students' syllabi (so tabs, SimDebt, DiscontinuityScore, unsat clean-ups and the String Tracker are right), wingman notes, return legs, OVERRIDE flags. Running the same engine with "don't pick anyone new" produces exactly that from the board as it stands.

## §25 The puck board's internals

`PuckBoard.psm1` (the biggest file) builds the window from `PuckBoard.xaml`. The design notes at the top of the file, expanded:

- **View models.** Every row, puck and slot is a WPF `DataTemplate` bound to a plain PSCustomObject. A row is refreshed by replacing its view model in the day's `ObservableCollection`.
- **One handler per event type.** Mouse, drag, key and focus events are handled once at window level; the element is found by its `Tag` (`IPSlot`, `StudentSlot`, `EtdBox`, ...).
- **Commands run after the event.** Anything that changes the board goes through `Request-PBCommand`, which runs it after the current WPF event finishes, so rows are never rebuilt inside their own events.
- **Background checking.** Pinned students are re-checked after every change by a `DispatcherTimer` (every 40 ms, a few at a time), keeping the corner markers current without freezing the window.
- **Autosave** every 2 s when something changed (`Boards\yyyy-MM-dd.json`).
- **Sheet watch.** Another timer (2 s) watches MasterData's modified time. When it changes, the Schedule sheet is re-read; each day gets a **fingerprint**. A new day gets a tab; a day you haven't edited since it was loaded updates by itself; a day you've edited gets the blue banner. A day you removed stays off until its fingerprint changes.
- **Click and place.** `PB.Held` is the puck you're holding; the next slot clicked gets it.
- **Line details.** `PB.SelectedLineId` is the selected line; `DetailsTimer` (30 ms) fills the See options list a few students at a time.
- **Fill / Export** call `Invoke-PBScheduler` (§24) behind a busy overlay with a progress bar.
- **Dialogs:** `LoadDayDialog.xaml`, `BulkLinesDialog.xaml`, `AbsenceDialog.xaml`, `ExportResultDialog.xaml`.

## §26 Excel in and out

**Reading:** ImportExcel (EPPlus underneath) reads sheets as rows of objects keyed by the heading row. `Open-FaisWorkbookForRead` opens the file with *read, allow others to write* sharing, so a workbook open in Excel can still be read (its last saved copy, which is why FAIS saves it first).

**Writing (`Export-FaisOutput`):**

1. Work out the rows for the days being exported, and the SimDebt, DiscontinuityScore, unsat flags and pending absences to write.
2. (Optional) save and close every open workbook (`CloseAllExcelWorkbooks`).
3. For each target file, `Request-FaisFileWritable`: save and close it in Excel; if that's impossible, show the "file is open" box (OK = retry, Cancel = give up and write nothing).
4. Read the existing MasterSchedule and merge: drop past days (before today minus `KeepPastDays`), replace the exported days, keep the rest.
5. Inside `Invoke-WithFileBackup` (all or nothing):
   - MasterSchedule (formatted, group borders, first tab),
   - each student's tab (changed cells only),
   - MasterData Students (SimDebt, DiscontinuityScore, flags),
   - String Tracker,
   - pending absences,
   - ATIS.
6. Optional Teams copy (`PublishToTeams`).
7. Reopen the workbooks FAIS closed (`ReopenExcel`).

---

# Part E: Data reference

## §27 MasterData.xlsx, sheet by sheet

**Don't rename columns.** FAIS finds most columns by their heading.

### Students
| Column | Used for |
|---|---|
| `SPKey` | Id |
| `SPName` | The student's name. Must match their tab name in OutputSchedule and their name on StudentAbsences. (Only the first 100 named rows are used.) |
| `IsCAP` | CAP rules (§13) and bonus (§14) |
| `IsInternational` | Needs an experienced IP |
| `Priority` | Added to every score; first tie-break |
| `UnsatSim`, `UnsatFlight` | §19 |
| `OnHold` | Never scheduled |
| `SimDebt` | §16 (column P if no heading) |
| `DiscontinuityScore` | §16 (added after the last heading the first time) |
| `CAPCounter1..7` | Not read by the current code |

### Instructors
`IPName`, `IsCounter` (blank = counter), `IsExperienced`, `IsCheck`, `Callsign`, `Num`, `OSC`, `OSNum`.

### Syllabus
The master syllabus a student starts from when they have no tab in OutputSchedule: `SyllabusID` (order), `Mission`, `Prereqs` (comma list, or 0), `OptingEvents`, `Type`, `IsCheck`, `Level` (formation level, used by LevelDifference), `Instructor`, `ETD`, `IsComplete`.

Current flow: FT5001 (dollar ride) opens the FE/FI sims and FT5002; FT5001–FT5008, solo FT5201, FT5101–FT5105, check ride **FT5390**; formation sims FN3001/FN3002/FN4001; formation FF5001–FF5008, solo FF5201, FF5101–FF5106, check ride **FF5390**; then AA5001–AA5007 with AA4001–AA4003 sims, check ride **AA5190**.

### Schedule (optional)
The day's lines. Headings found by name, in any order: `ETD` (or `Time`), `Date`, `Instructor` (or `IP`), `SortieType` (or `Type`), `Student` (or `SP`), `Mission` (or `Event`), `Callsign`, `Num` (or `#`), `Notes` (or `Remarks`). With no heading row, the old fixed order: ETD, Instructor, SortieType, Student, Mission, Callsign, Num, Notes.
- **ETD** can be a date and time or just a time (`0730`, `7:30 AM`...). A time-only row takes its date from the `Date` column or a row above that holds only a date.
- **SortieType** can be short: `DL`, `OB`, `RTN`, `OFT`, `IFT`.
- Blank rows and repeated heading rows are skipped.

### SoftWeights
One data row of weights (§14). Header names: `ContString`, `String4`, `String5`, `Zipper`, `NonCounter`, `BreakString`, `ExhaustMult`, `ExhaustDays`, `BITPerDay`, `BITDays`, `CAPBonus`, `FormLookAhead`, `LevelDifference`, `CounterContinuity`, `OFTIFTPen`, `DollarRideBonus`, `TransSimRatio`, `FormSimRatio`, `SimStopMult`, optional `NewCounter`, `Disc1String`, `DiscCounter`, `Disc3PlusString`.

### FormCallsigns
`Callsign`, `Num`, in lead/wing pairs (`HAZE 11`, `HAZE 12`, ...).

### StudentAbsences
`Student`, `StartTime`, `EndTime`, `Reason`, `IsActive`. Only active rows with both times are used.

### String Tracker (written by FAIS)
`StudentName`, `5+ Strings`, `4-Strings`, `3-Strings`, `2-Strings`, `1-Strings`, `TotalUniqueCounters`.

### Other sheets
`Instructions`, `Formation Sorties`, `All Students`, `Full Syllabus`, `Schedule People to Checkrides` are kept for people; the code doesn't read them.

## §28 OutputSchedule.xlsx and ATIS.xlsx

**OutputSchedule.xlsx** (FAIS folder):
- **MasterSchedule** (first tab): every exported day from today on. Columns ETD, Instructor, SortieType, Student, Mission, Callsign, Num, Notes. Only this sheet is rewritten, and only the exported days in it change.
- **One tab per student** (named exactly as `SPName`): their syllabus with Instructor, ETD and IsComplete for what they've done (and what's been scheduled). This is where each run **starts** from. FAIS writes only the changed cells, so your own columns and formatting are safe.

**ATIS.xlsx** (Desktop): one tab per exported day, flights in ETD order, a thick border, then sims. Rebuilt for the exported days; other tabs kept.

## §29 What is read and written, and when

| When | Reads | Writes |
|---|---|---|
| Start | MasterData: Schedule, Instructors, Students, SoftWeights, FormCallsigns, StudentAbsences, (Syllabus). OutputSchedule: each student's tab, MasterSchedule. | Nothing |
| Log Hooked Event | | MasterData: UnsatSim / UnsatFlight = TRUE |
| Log Incomplete Event | | MasterData: SimDebt, DiscontinuityScore. OutputSchedule: that event's cells |
| Change HOLD status | | MasterData: OnHold |
| Puck board | Schedule sheet again whenever MasterData is saved | `Boards\*.json` (autosave), `Templates\` when you save one |
| Update Syllabi | The training file | OutputSchedule: changed cells. MasterData: SimDebt, DiscontinuityScore |
| Add Student Absence | | MasterData: StudentAbsences |
| Export | | OutputSchedule: MasterSchedule, student tabs. MasterData: SimDebt, DiscontinuityScore, unsat resets, String Tracker, unsaved absences. ATIS |

## §30 Settings

In `FAIS.ps1`:

**`$BoardSettings`**

| Setting | Default | What |
|---|---|---|
| `OBReturnOffsetMinutes` | 180 | Typing `1015 O&B` adds the return this many minutes later |
| `ReturnSortieType` | `O&B Return` | Line type written on return lines |
| `SimInstructor` | `SIM IP` | IP shown on OFT / IFT lines with no IP |
| `BoardDir` | `Boards` | Autosave folder |
| `TemplateDir` | `Templates` | Templates folder |
| `Trace` | `-Trace` switch | Print every decision |

**`$OutputSettings`**

| Setting | Default | What |
|---|---|---|
| `WriteAtis` | `$true` | Write ATIS.xlsx |
| `KeepPastDays` | 0 | Exported days before today MasterSchedule keeps |
| `SaveAndCloseExcel` | `$true` | Save and close a workbook before writing it |
| `ReopenExcel` | `$true` | Reopen those workbooks afterwards |
| `CloseAllExcelWorkbooks` | `$false` | Save and close *every* open workbook first |
| `PublishToTeams` | `$false` | Also write daily sheets to a Teams/OneDrive workbook |
| `TeamsFolder`, `TeamsFileName` | | Where |

**Rules that are fixed in code** (change them in `SchedulingEngine.psm1`, see §38): turn time 3.25 / 3.5 hr, duty day 9 hr, crew rest 15 hr, absence window −1.5 / +2.5 hr, Checkride Priority +50, counter limits 6 / 3.

---

# Part F: The same thing, explained again

## §31 Life of a student

Follow one student, **Smith**, through a run.

1. **Start.** FAIS reads Smith's row on Students (`SimDebt 6`, `Priority 1`, `DiscontinuityScore 3`) and Smith's tab in OutputSchedule: FT5001 complete with Diaz, FT5002 and FT5003 with Lee.
2. **Reset.** Anything on Smith's tab dated from the board's first day on was a *plan* from an earlier run. It's made not complete again (and its sim debt and discontinuity backed out) because this run will plan those days again.
3. **Pre-run.** Nothing logged for Smith.
4. **Opted events.** FT5004 (prereq FT5003 done) and the sims whose prereqs are met.
5. **Board.** Smith's puck shows no badges. Hover: next events, sim debt 6, discontinuity 3, priority 1.
6. **Fill, 0800 Day Local with Lee.** Hard rules: not on hold, no absence, no other events, crew rest fine, FT5004 fits a Day Local. Pass. Score: Continue String (3-string with Lee) +25, Discontinuity +3, Counter Continuity +5, Sim Debt (flight) −6, Exhaustion −0.75, Priority +1 = **27.25**. Highest on the line → Smith gets it.
7. **Marked complete.** FT5004 is now complete on Smith's (copied) syllabus with Lee at 0800. SimDebt 6 → 9. DiscontinuityScore moves by whatever the new 3-string changed (one more 3+ string: −1 → 2).
8. **Later lines.** At 1000 Smith fails turn time (2 hr after 0800). At 1200 Smith could pass, but FT5004 is now done, so FT5005 is Smith's next flight; with SimDebt 9 a flight costs Smith 9 points and a sim on a sim line earns 9, so Smith is pulled toward the OFT / IFT lines.
9. **Export.** Smith's tab gets FT5004's Instructor, ETD and IsComplete. MasterData gets SimDebt 9 and DiscontinuityScore 2. The String Tracker row for Smith is updated.
10. **Tomorrow.** The next run starts from that tab.

## §32 Life of a line

Follow one line, **0800 Day Local, IP Lee**.

1. **Created** by the Schedule sheet (or typed `0800 DL Lee`). Board line: `{ETD 0800, SortieType Day Local, Instructor 'Lee, Pat [Capt]', Student '', Filled ''}`.
2. **On the board** it shows a Day Local stripe, Lee in the IP slot, an open student slot.
3. **Fill pressed.** `ConvertTo-ScheduleRows` turns it into an engine row with `BoardLineId`. Lee is looked up: counter, experienced, not a check IP, callsign `TWEET 04`.
4. **Its turn** in ETD order. Every student is tested. Jones has an unsat sim → `UNSAT_SIM`, out. Diaz is on hold → out. Smith and Brown pass.
5. **Scoring.** Smith 27.25 (FT5004), Brown 12.5 (FT5004). Smith wins (§33 shows the arithmetic).
6. **Filled.** `Student Smith, Mission FT5004, FilledBy Engine`, note added if any. Not a formation, not an O&B, so nothing more.
7. **Back on the board** (`Set-BoardFill`): blue italics, `AUTO`, `Filled = Engine`.
8. **You review** it: click it, Line details shows Smith's + / − list. See options shows Brown second.
9. **You might move Brown here.** Then the line is a plain pin (Filled cleared); the marker shows Brown's check.
10. **Export.** The engine runs once more with `-NoAutoFill`. The line is pinned, so `Invoke-PinnedLine` re-checks and records it. It's written to MasterSchedule and ATIS: `0800 | Lee, Pat [Capt] | Day Local | Smith | FT5004 | TWEET | 04 | ...`.

## §33 Worked scoring example

Two students can both fly **FT5004** on **0800 Day Local with Lee** (a counter). Weights from the current SoftWeights row.

**Smith:** FT rides so far: FT5001 Diaz, FT5002 Lee, FT5003 Lee. SimDebt 6. DiscontinuityScore 3 (one 1-string with Diaz, two counters). Last flew yesterday (FT5003); 1 flight and 1 sim in the last 5 days. Priority 1.

**Brown:** FT rides: FT5001 Diaz, FT5002 Ortiz, FT5003 Ortiz. SimDebt −3 (ahead on sims). Last flew 3 days ago (FT5003), and that is the only event in the last 5 days. Priority 1. Lee has never flown Brown.

| Rule | Smith | Brown |
|---|---|---|
| String | Diaz, Lee, Lee + **Lee** = 3-string with Lee → **Continue String +25** | Diaz, Ortiz, Ortiz + **Lee** = zipper ride (Ortiz's 2-string waits on the next ride) → **Zipper −3** |
| Discontinuity (string) | **+3** (only with Continue String) | — |
| Counter | Lee already counts → **Counter Continuity +5** | Lee is new → New Counter 0 (not shown) |
| Sim Debt (flight) | −SimDebt = **−6** | −(−3) = **+3** |
| Exhaustion | (2×1 + 1) × −0.25 = **−0.75** | (2×1) × −0.25 = **−0.5** |
| BIT Tracker | flew yesterday → 0 days off → none | 3 days − 1 = 2 days × 6 = **+12** |
| Student Priority | **+1** | **+1** |
| **Total** | **27.25** | **12.5** |

Smith gets the line: keeping Smith's string with Lee is worth more than Brown's days off. If Brown had been off 5 days, the BIT points (4 × 6 = 24) would have pushed Brown to 24.5 and made it close. If Lee had flown Brown once before, Brown would also have had Counter Continuity +5.

After Smith flies it, Smith's DiscontinuityScore moves from 3 to 2: the Lee rides are now a 3-string (−1), while the Diaz 1-string (+1) and the two counters (+2) stay.

On the board, Smith's Line details would read something like:

```
✔ No hard rules broken
+ 3 String                                             +25
+ Discontinuity 3, 3 String                             +3
+ Lee already counts (2 of 6 Fundamental counters)      +5
− Sim Debt 6, flight adds to it                         −6
− 1 Flight and 1 Sim in last 5 Days                  −0.75
+ Student Priority                                      +1
  Total                                              27.25
```

## §34 Scenarios and recipes

**"A student must fly a specific event at a specific time."**
Drag them onto the line and pick the mission in the dropdown (or type `0800 DL Lee sp:Smith m:FT5103`). If it breaks a rule, choose **No** to keep it; it goes out `OVERRIDE`.

**"Keep a student off the schedule today."**
Put them on hold (pre-run), or add an absence covering the whole day.

**"The scheduler keeps picking the same student."**
Click the line, press **See options**, and compare the + / − lists. Usually it's a big Continue String or BIT. Lower that student's `Priority`, or raise another's.

**"A line comes out NEED SP."**
See options with **Hide rule breaks** off: the ✖ next to each student says the first rule they break. Common causes: everyone has 2 events already, crew rest, no opted event for that line type (e.g. everyone needs sims and it's a Day Local).

**"An event got cancelled after I exported."**
Next run: **Log Incomplete Event** in the pre-run. Or, same run: **Load a day... > Last exported schedule**, take the student off, Fill, Export.

**"I updated the Schedule sheet in Excel while the board was open."**
Save it. The board follows within a few seconds: a day you haven't edited updates itself; otherwise the blue banner asks.

**"Formation lead but no wingman."**
Add a Day Local line within 22 minutes of the lead, or pair one yourself with right-click > *Fly wing to...*.

**"I want to tune the scoring."**
Change values on the SoftWeights sheet, save, and restart FAIS. Use `-Trace` and See options to see the effect.

**"Count a student's discontinuity again from scratch."**
Clear their `DiscontinuityScore` cell in MasterData. The next start recounts it and Export writes it.

**"Plan two days at once."**
**Add day**, build both, **Fill all 2 days**. Crew rest is checked across them.

---

# Part G: Help

## §35 Troubleshooting

| Symptom | Cause / fix |
|---|---|
| *"Please install or copy the ImportExcel module"* | Install it (§3). |
| *"FAIS needs PowerShell in FullLanguage mode"* | The machine runs PowerShell in Constrained Language mode (policy). WPF can't run; ask your admin. |
| *"running scripts is disabled on this system"* | Start with `-ExecutionPolicy Bypass`. If files came from the internet, `Get-ChildItem -Recurse | Unblock-File`. |
| *"OutputSchedule.xlsx could not be opened"* and FAIS stops | It's corrupt or locked by another user. Fix or close it. FAIS stops on purpose: carrying on would start everyone from a blank syllabus. |
| *"The SoftWeights sheet could not be read"* | Check the SoftWeights sheet exists with its heading row. |
| *"File is open"* box at Export | Excel couldn't be reached (a cell is being edited, a dialog is open, or someone else has it). Finish the edit, click OK. Cancel writes nothing. |
| Export failed, *nothing was written* | Read the reason in the result window. The board is saved; fix and Export again. |
| A student's absence doesn't apply | The name on StudentAbsences doesn't exactly match `SPName`. The console lists it at start. |
| *"No tab for X in OutputSchedule yet"* | Normal for a new student; they start from the master Syllabus and get a tab at Export. |
| *"student 'X' not found in MasterData"* | A name on a line doesn't match a student. Check spelling, or use the full name if two students share a last name. |
| Lines from the Schedule sheet don't appear | They're dated before the board's first day (today or earlier). The banner offers **Put them on <first day>**, or use Load a day... |
| A day you removed came back | Its lines on the Schedule sheet changed. |
| Board is slow | Avoid pressing See options on many lines; it ranks every student. |
| Picks look odd | Run with `-Trace`, or use See options, to see every candidate's score. |

## §36 FAQ

**Does the scheduler ever move a student I put on?** No. Pinned students stay, even if they break a rule (they're flagged OVERRIDE).

**Why did it put no one on a line?** Nobody passed the hard rules. See options shows why each student failed.

**Does Export fill open lines?** No. Export writes the board as it stands; open lines go out NEED SP. Press Fill first if you want them filled.

**Can I fill, change things, and fill again?** Yes. Only open lines are filled; your changes stay.

**What happens if I close the board without exporting?** Nothing is written to OutputSchedule or MasterData. Your lines are autosaved and come back next time.

**Why does FAIS make planned events "not complete" when it starts?** A previous run marked them complete when it scheduled them. They're plans, not flown events, so they're reset from the board's first day on, and re-planned. Flown events before that day stay complete.

**Is the counter limit enforced?** No. It's shown in the score reasons only.

**Why is a non-counter purple?** Non-counters (IsCounter FALSE) don't count toward strings unless they fly a student twice in a row.

**Where's the old code?** `OldFunctions.psm1`, kept for reference, not loaded.

**Can I run it on a Mac or Linux?** The engine and tests can run under PowerShell 7, but the board needs WPF and Excel COM, so day-to-day use is Windows only.

## §37 Testing

- **Self-test** (no Pester needed, no real files touched):
  `powershell.exe -STA -ExecutionPolicy Bypass -File .\Tests\Invoke-SelfTest.ps1`
  Checks the engine, board model and placement checks with synthetic students; an Excel round trip on temp files (if ImportExcel is installed); and the board window driven off-screen (it may flash). `-SkipWindow` skips the window. Exit code 0 = all passed.
- **Pester 5 suites** in `Tests\*.Tests.ps1`, one per area: `HardRules`, `FillSchedule`, `BoardModel`, `Assessment`, `PuckBoardLogic`, `DataIO` (needs ImportExcel), `DataUpload`, `ExcelFiles`, `Export`, `ScheduleSheet`, `SyllabusImport`, `OneDay`, `Static`. Test data comes from `Tests\Fixtures\New-FaisTestData.ps1`.
  Run with `Invoke-Pester .\Tests`.

## §38 Developer guide: changing FAIS safely

**Golden rules:**
1. **The engine decides, the board displays.** Put rule logic in `SchedulingEngine.psm1` / `Syllabus.psm1`, never in `PuckBoard.psm1`. The board's checks (`Assessment.psm1`) call the same functions, so they stay in step automatically.
2. **Never change the real students in a "what if".** Use `Copy-StudentForEval` / `Copy-StudentList`.
3. **Mark events complete only through `Complete-StudentMission`** (and take them back through `Undo-StudentMission`), so SimDebt, DiscontinuityScore and unsats move exactly once.
4. **Read flags through `ConvertTo-FaisBool`.**
5. **Keep module order** in `FAIS.ps1` (a module can only call ones loaded before it).
6. **Write tests** next to the matching suite and run the self-test.

**Add a hard rule:** in `Test-HardRules`, add a block in the right section:
```powershell
if (<condition>) {
    $violations.Add((New-RuleViolation 'MY_RULE' "Plain reason with the numbers"))
    if (& $stop) { return (New-HardRuleResult $violations $notes $eventResults) }
}
```
Use `$ev.Add(...)` instead inside the per-event loop for an event-level rule. Use `-Severity Advisory` for a warning that shouldn't block. Add a case to `Tests\HardRules.Tests.ps1`.

**Add a soft rule with a new weight:**
1. Add the column to MasterData's SoftWeights sheet.
2. Read it in `IngestData` (the `$SoftWeights` object in `DataDownload.psm1`).
3. In `SoftRules`, add points with `Get-Weight $Config 'MyWeight' <default>`, and set both `$Breakdown['My Rule']` and `$Reasons['My Rule']` (the text the board shows).
4. Add a test in `Tests\FillSchedule.Tests.ps1` or `Assessment.Tests.ps1`.

**Change a fixed number** (turn time, crew rest...): search `Test-HardRules` for the number (e.g. `3.25`, `15`), and update the README and this guide.

**Add a line type:** update `Get-AllowedEventTypes` (Syllabus), the `Test-Is...SortieType` helpers and `ConvertTo-CanonicalSortieType` (DataManipulation), and `ConvertTo-BoardSortieType` (BoardModel).

**Keep the text-only copies in step** if you use them: refresh `FAIS_txt_update\` with the changed files as `.txt`.

## §39 File map

| Path | What |
|---|---|
| `FAIS.ps1` | Entry point, settings, run order |
| `Modules\DataManipulation.psm1` | Shared helpers |
| `Modules\ExcelFiles.psm1` | Save / close / reopen workbooks in Excel |
| `Modules\Syllabus.psm1` | Syllabus, opted events, sim debt, strings, discontinuity |
| `Modules\DataDownload.psm1` | Reading MasterData / OutputSchedule |
| `Modules\SchedulingEngine.psm1` | Hard rules, soft rules, the fill |
| `Modules\DataUpload.psm1` | Export |
| `Modules\SyllabusImport.psm1` | Update Syllabi |
| `Modules\PreRunMenu.psm1` | Pre-run pop-ups |
| `Modules\BoardModel.psm1` | Board data, undo, typing, autosave |
| `Modules\Assessment.psm1` | Board checks and tooltips |
| `Modules\PuckBoard.psm1` + `PuckBoard.xaml` | The board window |
| `Modules\LoadDayDialog.xaml`, `BulkLinesDialog.xaml`, `AbsenceDialog.xaml`, `ExportResultDialog.xaml` | Dialogs |
| `Modules\BoardLogo.png` | Title-bar logo |
| `MasterData.xlsx` | Input data (§27) |
| `OutputSchedule.xlsx` | Output and each student's progress (§28) |
| `Big Board.xlsx`, `21-Sep, 14 OG Phase 2 + sub.xlsx` | Sample training files for Update Syllabi |
| `Boards\` | Autosaved board, one JSON per day |
| `Templates\` | Saved wave templates |
| `Tests\` | Self-test and Pester suites |
| `OldFunctions.psm1` | Legacy, not loaded |
| `FAIS_txt.zip`, `FAIS_catalog.txt`, `FAIS_txt_update\` | Text-only copies for machines that block scripts |
| `Docs\` | This guide |

## §40 Quick index

| Topic | Where |
|---|---|
| Absences | §11, §13 (#4), §19, §27 |
| BIT | §14, §16 |
| CAP students | §13 (#5, #6, #14), §14 |
| Check rides | §13 (event rules), §14 (Checkride Priority) |
| Colors | §6 |
| Context rows (other exported days) | §9, §24, §2 |
| Crew rest, turn time, duty day | §13 (#9–#13) |
| Direct support | §17 |
| Discontinuity score | §16, §33 |
| Export | §8, §26, §29 |
| Formations / wingmen | §17 |
| Hard rules | §13 |
| Line details / See options | §8 |
| O&B / return legs | §18 |
| OVERRIDE | §12, §13, §36 |
| Pre-run pop-ups | §5 |
| Scoring / soft rules | §14, §33 |
| Selection process | §12, §24, §32 |
| Sim debt | §16 |
| Strings / counters / zipper | §15 |
| Tie-breaks | §20 |
| Unsats | §19 |
| Update Syllabi | §10 |
