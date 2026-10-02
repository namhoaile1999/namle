# Implementation Spec - CreoRegister v1.0  (Spec v0.2, 2026-10-02, APPROVED + BUILT, NOT YET TESTED)

Status: approved by Nam 2026-10-02 and built as `CreoRegister.bas` v1.0. Not yet compiled or run.
Test with `docs/CreoRegister_TEST_PROTOCOL.md`.
Symbol legend: [?] = open question / default you may veto.

## What it does
A small Excel tool workbook, `CreoRegister.xlsm`, sits next to Creo. It shows
the next free part number on each register sheet. When you make a new part or
assembly in Creo, type the name (for example `M281975`), fill DESCRIPTION in
the New File Options dialog and click OK. Then click one button in the tool:
**Register active Creo part**. The tool reads the part's name and DESCRIPTION
from Creo, opens the shared Part Number Register, finds the matching row,
writes the Project ID and Description, saves and closes it. It also logs what
it did.

## Who it is for / NOT for
- For: Nam (and later colleagues) creating new M / A / B parts in Creo who
  need them recorded in `Part_number_register.xlsx` on the shared drive.
- NOT for: editing or renaming rows that already exist, Elec parts (E...),
  Fasteners (F...), Documents (D...), and bulk registration of old parts.

## Core problem
Today the same information is typed twice, once in Creo and once in the
register. The register is protected, shared and 4000 rows deep, so finding
the free row is slow. Typing it twice also risks the two records disagreeing
(wrong row, wrong project, a typo in the description).

## Facts found in the real register (checked 2026-10-02 on the uploaded copy)
| Sheet | Part No rule | Key column | Columns the tool writes | Sheet protected? |
|---|---|---|---|---|
| Manufactured parts | formula `="M" & D & E` | E = Number sequence (4 digits, text) | C Description, D Project ID | Yes (C, D unlocked) |
| Assemblies | formula `="A" & D & E` | E = Number sequence (4 digits, text) | C Description, D Project ID | Yes (C, D unlocked; format cells NOT allowed) |
| Bought parts | typed value, e.g. `B002678` | B = Part No | C Description | No |
| Elec | `"E" & proj & ElecID digit & 3-digit seq` | - | (out of scope v1) | Yes |

- **The number sequence is shared by ALL projects on one sheet.** M...1975 is
  one row. Project 28 or 21 only changes column D. So `M281975` and
  `M211975` cannot both exist. The tool enforces this.
- **Trap: Project ID cells are formatted "General".** Writing `03` the normal
  way turns into the number 3, and the Part No becomes `M30053` (wrong). The
  tool writes Project ID as text (with an apostrophe prefix). It cannot change
  cell formats, because Assemblies forbids that under its protection.
- Rows with a Project ID but no Description exist (26 on M, 11 on A, for
  example M030053). The tool treats them as **used**.
- Fully empty holes exist in the middle (M 0240, M 1619, A 0235, A 0236).
  They are free if you type them on purpose. "Next free" ignores them and
  shows the number after the last used row.
- Next free on the uploaded copy: **M 1975, A 0561, B B002678.**
- Bought part numbers are not all standard: 27 have suffixes (`B001050A`,
  `B001100-2`). B lookup is an exact match on column B.

## Success criteria (measurable)
1. One click after Creo's OK writes the correct row. Afterwards the register's
   Part No cell (column B) shows exactly the Creo name. The tool reads it back
   to prove this.
2. The tool never writes into a used row. A used row gives a popup that names
   the row's current description, and nothing changes.
3. The tool writes only these cells: C and D on M/A, C on B. It never deletes
   rows and never changes formats, formulas, locked cells or other sheets.
4. If a colleague has the register open, the tool refuses with a plain-language
   message, and nothing is written or saved.
5. Every attempt (OK or FAIL plus reason) is logged on the tool's Log sheet
   with a timestamp.
6. Non-coder use: set the register path once (Pick button), then use one button
   per part. No VBA editor after setup.

## Out of scope (v1)
- Elec (E...), Fasteners, Documents, Blacktrace PN sheets.
- Updating a row when DESCRIPTION changes later in Creo.
- Checking that an A-number is really an assembly in Creo (see open questions).
- Creating parts in Creo from Excel.
- A Creo toolbar button or mapkey trigger (possible v2).
- Changing anything in the register's design, protection or formats.

## Environment / prerequisites
Same as your verified CreoExport setup (lessons doc section 1): Creo VB API
registered, "Creo VB API Type Library" reference ticked (exactly one),
`PRO_COMM_MSG_EXE` set, one Creo session running. Read and write access to the
register's shared folder.

## Workflow (what you do)
1. Open `CreoRegister.xlsm`. Click **Refresh next free** to see e.g.
   `M: 1975 | A: 0561 | B: B002678`.
2. In Creo: File > New > Part > name `M281975` > fill DESCRIPTION > OK.
3. In Excel: click **Register active Creo part** (or press Ctrl+Shift+R [?]).
4. Popup: `OK - M281975 written to Manufactured parts row 2042: Project 28,
   "Bracket for X"`. The register is already saved and closed.

## Build steps + decisions
| # | Step | Decision | Default chosen [?] = veto welcome | Why |
|---|---|---|---|---|
| 1 | Tool sheet + SETUP macro | Layout | One sheet "Register": B1 = register path, Pick button, Refresh next free, Register button, next-free panel. A "Log" sheet. | Same pattern as your other tools |
| 2 | Trigger | Excel button vs Creo mapkey vs Excel-first | **Excel button** (you did not pick one, so I took the recommended option) [?] | Fewest moving parts; only verified API calls |
| 3 | Read Creo | Which model | `session.CurrentModel` = the model in Creo's active window | Model you just created is the active one |
| 4 | Parse name | Accepted formats | `M`+6 digits, `A`+6 digits (2 project + 4 sequence); `B`+anything (exact match). Anything else is refused. Suffixes like `M281975_V2` are refused [?] | Strict = no wrong rows |
| 5 | Project check | Validate project ID | Must exist in register sheet **Index** column A (01-28 today) | Catches typos like M821975 |
| 6 | Description | Bad values | Refuse if parameter missing, blank, or equal to `START_PART` | Template default must never reach the register |
| 7 | Open register | Shared-drive lock | If already open in this Excel, use it (refuse if it has unsaved edits). Otherwise open it without the "file in use" prompt. If it opened read-only (someone else has it), close it and refuse | No writes to a copy that cannot be saved |
| 8 | Free-row check | What counts as used | M/A: C or D not empty. B: C not empty = used; D-F filled but C empty = Yes/No warning | Matches the 26/11 "project but no description" rows |
| 9 | Same part twice | Idempotency | Row already has the same project + description: "already registered", no write | Safe to click twice |
| 10 | Write + prove | Read-back | Write D (text) and C, then read column B. If it is not the Creo name, clear the 2 cells the tool just wrote and FAIL | Catches the "03 becomes 3" trap and any wrong row |
| 11 | Save | Backup copy? | Save in place, **no backup file** on the shared drive; the Log sheet records every write so it can be undone by hand [?] | Avoids clutter in the company folder |
| 12 | Log | Where | Tool workbook "Log" sheet: time, Creo name, sheet, row, project, description, result | Audit trail on a shared file |

## Known weak spots you accept
- **Race with colleagues:** "Next free" can be stale. Registration re-checks
  the row at write time, so the register itself never gets a double entry.
  But if a colleague took 1975 a minute ago, your Creo part is already named
  M281975 and must be renamed in Creo. The tool tells you; it cannot rename.
- **Register busy:** while anyone has the register open, registering is blocked
  until they close it. This is by design: it is the only safe option on a
  shared drive.
- **Active window only:** if another model's window is active in Creo when you
  click, that model is registered instead. The popup shows the name, so check it.
- **API verification gap:** `session.CurrentModel` and
  `IpfcModel.InstanceName` were confirmed via PTC docs page titles and community
  posts in search results. The PTC docs site is blocked from my environment, so
  I could not read the pages directly. Test step 1 proves them on your machine
  first. Every other Creo call is from your verified lessons doc (connect,
  parameter read).
- Not tested against Creo with Windchill. Your setup appears to be local files.

## Data-safety statement (skill section 9)
- Disk writes: exactly one `Save` of the register workbook, after a proven
  read-back. No Kill, no MkDir, no rename, no SaveAs, no backup files.
- Register cell writes: at most 2 cells (C, D) in one row. The only "undo" path
  clears exactly the cells the tool wrote in the same click.
- Creo: read-only. No Save, no parameter writes, no window changes. The tool
  disconnects on every path.
- Source = destination hard block: not applicable (no export folder).

## Verification plan (you run it; numbered protocol comes with the build)
Run everything against a **COPY** of the register in a local test folder first.
1. Compile + Refresh next free: expect M 1975, A 0561, B B002678 on the copy.
2. Happy path M: new part `M281975`, DESCRIPTION "TEST PART" > Register. In the
   copy: row with sequence 1975 has D=28 (text), C=TEST PART, B shows M281975.
3. Happy path A (`A280561`) and B (`B002678`).
4. Click Register again on the same part: "already registered", no change.
5. Broken cases (each must refuse with a clear reason and change nothing):
   `M280053` (row used by project 03) | DESCRIPTION left `START_PART` |
   `M991976` (unknown project 99) | `E213167` (Elec not in v1) | `TEST1` (bad
   format) | copy set to Read-only in Windows file Properties (simulates a
   colleague holding it).
6. Quality: open the copy by hand and check the Part No cells, the protection
   still on, and nothing else changed.
7. Only then point B1 at the real register.

## Builder handoff
Built by the senior session directly (single small module, about 400 lines).
Integrity gate (skill section 10) and manifest apply: ASCII-only .bas, a
version header, procedure list and tail check on the saved file.

## Decisions (Nam, 2026-10-02)
- Q1 Trigger: Excel button (default kept; no explicit pick).
- Q2 Suffixed names: REFUSE.
- Q3 Part/assembly type check: not in v1 (default kept).
- Q4 Shortcut: Ctrl+Shift+R.
- Q5 Users: colleagues too, so a user manual was added (`docs/CreoRegister_MANUAL.md`).

## Build notes (deviations from v0.1, all on the safe side)
1. A **confirm popup** (name / description / sheet / project, Yes/No) appears
   BEFORE the register is opened. It covers the "wrong Creo window active"
   weak spot, and it does not hold the shared-file lock while you read it.
2. Descriptions that Excel would convert (leading = + - @, numbers, dates such
   as "1/2") are written as text with an apostrophe prefix.
3. Read-back gives a separate, clear message when the row has no Part No
   formula. Assemblies has numbered rows 0737+ (sheet rows 755+) without the
   formula.
4. "Prepared rows left" counts only rows that have a Part No.
5. The tool workbook saves itself after each attempt so the log survives,
   but only if it has already been saved as .xlsm.
6. Senior built directly, with no builder or critic subagents (the user did
   not ask for agents). A self-critic pass caught 6 issues before delivery:
   close-after-save reported as FAIL; date/number auto-conversion of
   descriptions; Save As popup on an unsaved tool workbook; the tool path
   picked as the register; overstated rows-left on Assemblies; a parameter
   named `caption` re-casing `.Caption`.

## Open after build
- [?] The register file contains Sheet Views + threaded comments, which are
  SharePoint/OneDrive co-authoring features. If the register is co-authored
  in SharePoint/Teams instead of sitting on a plain network drive, a colleague
  having it open does NOT make it read-only. The "register busy" guard then
  does not trigger, and protection rests on the at-write re-check plus
  co-authoring merge. Ask Nam where the file lives (test protocol, last
  section).
- [?] `session.CurrentModel` / `IpfcModel.InstanceName` are proved only by
  test step 2.
