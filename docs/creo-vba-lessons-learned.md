# Creo Parametric + Excel VBA — Lessons Learned (WORKING COPY, compacted 2026-07-10)

Project: batch export STEP (.prt/.asm) + PDF/DXF (.drw) from Creo 10/11 via
Excel VBA. Status: **CreoExport.bas v1.25 + CreoBackup.bas v1.1 — BOTH
USER-TESTED, VERIFIED BASELINES 2026-07-12** (version history in section 8;
v1.25 = filter dropdown, CreoBackup v1.1 = batch Save-a-Backup, both passed
their full test protocols). This doc = verified facts + every error hit
+ fix, so the next task starts from here, not from zero. Append new lessons
here. Full pre-compaction narrative (665 lines):
`archive/creo-vba-lessons-learned-FULL-2026-07-10.md` — read it only if a
detail below is insufficient.

User: mechanical engineer, NOT a coder. Complete .bas files only, numbered
steps, never hand-edit lines. Full behavior rules: skill sections 0-1, 8b.

## 1. Architecture + environment (verified)

- Creo has NO internal VBA. Automation via **Creo VB API** (pfcls, async COM
  only; free optional installer component). Excel connects to a RUNNING Creo
  session; both must be 64-bit.
- Zero-code alternative: **Creo Distributed Batch** does STEP+PDF batch —
  separately licensed; THIS user does NOT have it. Check first on new sites.
- Setup checklist (in order):
  1. `<Creo>\Common Files\<build>\vbapi\` exists (else rerun installer).
  2. Run `vbapi\bin\vb_api_register.bat` **as admin** (re-run after any Creo
     minor-version update — API breaks otherwise).
  3. Excel Tools > References > "Creo VB API Type Library" — tick exactly
     ONE Creo library (two ticked = type-mismatch bugs).
  4. Env var **`PRO_COMM_MSG_EXE`** = full path to
     `...\x86e_win64\obj\pro_comm_msg.exe` ← THE actual "cannot connect"
     fix on this machine.
  5. Restart Excel AND Creo (env vars read at startup). Still failing → run
     both as admin.
- TWO Creo sessions on one machine → ALL connects fail with the generic
  "Cannot connect" (PTC ambiguity design, manual 3.5). NOT a bug; Nam
  mistook it for one — expect users to. Backlog: mention it in the MsgBox.

## 2. Errors solved (never re-debug these)

1. Compile "Syntax error" on `(New Class).Method(...)` — VBA has no inline
   New (VB.NET forum examples mislead). Two lines: Dim + Set = New.
2. Compile "Variable not defined" on an enum — a GUESSED name. Never guess
   pfcls names; verify in PTC docs/working code. Every guessed name was
   wrong, every verified name compiled first try. Qualify enums fully.
3. Same compile error after a "fix" — user imported the OLD file (browser
   saved `CreoExport (1).bas`). Version in header; confirm after every swap.
4. Runtime "Cannot connect" though Creo open — ladder: `Connect("", "", ".",
   20)` (never Nothing×4) → run as admin → set PRO_COMM_MSG_EXE (actual fix)
   → check duplicate references → check for a second Creo session.
5. Compile "Method or data member not found" on an inherited member — VBA
   early binding can't reach parent-interface members through a child
   variable. Hold the SAME object in two variables, one per interface:
   ```vba
   Dim ft As pfcls.IpfcFeature:       Set ft = feats.Item(i) ' -> ft.Status
   Dim cf As pfcls.IpfcComponentFeat: Set cf = feats.Item(i) ' -> cf.ModelDescr
   ```
   Applies to IpfcFeature/IpfcComponentFeat, IpfcSheetOwner,
   IpfcParameterOwner.
6. "Sub or Function not defined" on MakeDir (twice!) = tail-TRUNCATED .bas
   file, not a logic bug. See integrity gate (section 6). A repeat compile
   error matching a documented truncation → suspect the file first.
7. `XToolkitNotDisplayed` on Export — model loaded but not displayed. Fix:
   CreateModelWindow + Display BEFORE Export (param READS need no display).
8. `XToolkitObsoleteFunc` reading `param.Value.StringValue` directly (Creo
   9+) — read via `GetScaledValue` (section 3 param block).
9. `XToolkitNotFound` on RetrieveModel = model renamed in Windows Explorer
   (12/367 in production). Data problem, not code. The "(can't open)"
   scan marker now catches these before export.
10. STEP colors lost = invalid config-option name silently swallowed by
    On Error wrapper for weeks. Real option: `step_export_format` =
    `ap214_is`. Error-silencing hides wrong names — use sparingly, comment
    why.

## 3. Verified API names (safe to reuse verbatim)

Connection / session:
```vba
Dim connMaker As pfcls.CCpfcAsyncConnection      ' Set = New, then:
Set conn = connMaker.Connect("", "", ".", 20)    ' IpfcAsyncConnection
Set session = conn.Session                        ' IpfcBaseSession
session.ChangeDirectory folder
session.SetConfigOption "name", "value"           ' wrap in On Error Resume Next
session.EraseUndisplayedModels
conn.Disconnect 2
```

Open + display model:
```vba
CCpfcModelDescriptor.CreateFromFileName("name.prt")  ' latest version auto
session.RetrieveModel(descr)                          ' IpfcModel
session.CreateModelWindow(mdl) -> mdl.Display         ' only for Export
win.Close                                             ' win.Activate NOT needed
' (Activate removal verified safe v1.16: exports + Close work without it)
```

STEP export:
```vba
CCpfcGeometryFlags.Create() ; flags.AsSolids = True
CCpfcSTEP3DExportInstructions.Create( _
    EpfcAssemblyConfiguration.EpfcEXPORT_ASM_SINGLE_FILE, flags)
mdl.Export outPath, inst
' config: "step_export_format" = "ap214_is" (AP203 default drops colors)
```

PDF export (drawings):
```vba
CCpfcPDFExportInstructions.Create()
' options collection: New pfcls.CpfcPDFOptions   (Cpfc, not CCpfc!)
' option:  CCpfcPDFOption.Create() ; .OptionType ; .OptionValue
' arg values: New pfcls.CMpfcArgument -> .CreateBoolArgValue / .CreateIntArgValue
' verified enums: EpfcPDFOptionType.EpfcPDFOPT_PENTABLE (bool),
'   EpfcPDFOPT_RASTER_DPI (int 100-600), EpfcPDFOPT_SHEETS (default ALL),
'   EpfcPDFOPT_LAUNCH_VIEWER (bool False = no popup, v1.16 D1)
inst.Options = opts ; mdl.Export outPath, inst
```

DXF export (drawings, LAST SHEET, v1.20 VERIFIED pattern):
```vba
Dim so As pfcls.IpfcSheetOwner: Set so = mdl     ' two-handle rule (err 5)
oldSheet = so.CurrentSheetNumber
so.CurrentSheetNumber = so.NumberOfSheets        ' set-and-confirm (sync)
mdl.Display                                       ' MANDATORY - see section 4
' CCpfcDXFExportInstructions.Create() -> mdl.Export
' then WAIT until file exists + size stable (2x100ms polls) BEFORE restoring
' oldSheet (restore-while-writing captured the wrong sheet, v1.15 bug);
' timeout 30 s = loud FAIL. Restore on failure path too (contamination).
' kill *.log* after.
```

Assembly tree walk:
```vba
solid = asmMdl                                     ' IpfcSolid handle
feats = solid.ListFeaturesByType(False, EpfcFeatureType.EpfcFEATTYPE_COMPONENT)
ft = feats.Item(i)   ' IpfcFeature  -> ft.Status (0 = active, skip non-zero)
cf = feats.Item(i)   ' IpfcComponentFeat -> cf.ModelDescr
cd.InstanceName ; cd.Type = EpfcModelType.EpfcMDL_ASSEMBLY / EpfcMDL_PART
' recurse: session.GetModel(nm, EpfcMDL_ASSEMBLY), fallback RetrieveModel
' NOTE: the boolean = visible-vs-internal, NOT active-vs-suppressed.
' Filter suppressed via Feature.Status yourself.
```

Parameter read (verified from PTC Community thread 127448):
```vba
Dim paramOwner As pfcls.IpfcParameterOwner
Set paramOwner = mdl                          ' same object, new handle
Dim p As pfcls.IpfcParameter
Set p = paramOwner.GetParam("DESCRIPTION")    ' Nothing if param absent
Dim bp As pfcls.IpfcBaseParameter
Set bp = p                                    ' .Value.discr lives HERE
Select Case bp.Value.discr                    ' EpfcParamValueType.*
    Case EpfcParamValueType.EpfcPARAM_STRING:  v = p.GetScaledValue.StringValue
    Case EpfcParamValueType.EpfcPARAM_INTEGER: v = CStr(p.GetScaledValue.IntValue)
    Case EpfcParamValueType.EpfcPARAM_BOOLEAN: v = CStr(p.GetScaledValue.BoolValue)
    Case EpfcParamValueType.EpfcPARAM_DOUBLE:  v = CStr(p.GetScaledValue.DoubleValue)
End Select
' NEVER p.Value.StringValue directly -> XToolkitObsoleteFunc on Creo 9+.
```

Class-name pattern: `Ipfc...` interface (variables) | `CCpfc...` creator
(New + .Create) | `Cpfc...` collection (New directly) | `CMpfcArgument` arg
factory | `Epfc...` enums, fully qualified | inherited members need a
variable OF the declaring interface.

## 4. DEAD ENDS — verified, do NOT revisit

- **`IpfcDXFExportInstructions.OptionValue` and everything on it**
  (CCpfcExport2DOption / IpfcExport2DOption, ExportSheetOption,
  ModelSpaceSheet, EpfcExport2DSheetOption enums, the Sheets Iintseq/
  Cintseq sequence): compiles, but throws `pfcExceptions::XUnimplemented
  ... "supported in Object TOOLKIT only"` at mdl.Export, killing EVERY DXF
  (v1.19 user test, 2026-07-10). PTC's doc caveat is enforced at runtime.
  ExportPDF's `IpfcExportInstructions.OptionValue` (pentable/DPI/viewer) is
  a DIFFERENT, working mechanism — don't confuse them.
- **COM lazy validation**: pfcls may accept a property SET client-side and
  reject it only at the CONSUMING call. Error-wrapping the SET with a
  "fallback" proves nothing — catch at the consuming call or don't attach.
- **No config.pro sheet-selection option** for DXF export exists (PTC r12
  export docs, checked 2026-07-10) — dialog-only choice.
- **Timing theories for the DXF wrong-sheet bug**: sheet-set race (v1.17)
  and restore race (v1.18) both FALSIFIED — retries = 0, file-stable wait
  at 100 ms floor, and the SAME 3 files failed deterministically. A race
  cannot pick the same files every run.
- **"Smart" param-refresh skip** (by disk date): rejected deliberately —
  disk dates can't see unsaved in-session edits, which are exactly what a
  refresh is for. Full re-read; memory bounded by the 25-file erase sweep.

## 5. The DXF wrong-sheet saga (v1.15→v1.20) — root cause + prevention

SYMPTOM: multi-sheet drawing, DXF contains the wrong sheet, Status = OK
(silent bad output). Same 3 of 7 files every run.

ROOT CAUSE (proven by v1.20 verification, 2026-07-10): setting
`CurrentSheetNumber` changes the DATA-side sheet but not the DISPLAYED one,
and Creo's DXF exporter follows the DISPLAYED sheet. Drawings saved/opening
on sheet 1 therefore always exported sheet 1. THE FIX: `mdl.Display` after
the confirmed sheet switch, before Export. Verified: 7 drawings (incl. the
3 bad ones, 2 sheets each, opening on sheet 1) × 2 runs — all last sheet.

PREVENTION RULES:
- Any API-side sheet/view change an exporter must honor → `mdl.Display`
  BEFORE Export.
- Restore the user's sheet AFTER the file is proven complete on disk
  (exists + size stable on two 100 ms polls; restore-while-writing was the
  v1.15 bug) and on the failure path too.
- "Status OK" never guaranteed output content: verify output QUALITY once
  per feature. Old verifications are sample-limited, not absolute (v1.14.1
  "DXF verified" rested on 4 files that lacked the bad property).
- Evidence beats theory: a diagnostic counter that can falsify your
  hypothesis (retry counter, max-wait metric) is worth shipping even when
  the fix built on it is wrong.
- RESIDUAL (open): DXF delivered by v1.16 or earlier may contain silent
  wrong sheets — spot-check; re-export suspects with v1.20.

## 6. Process + multi-agent lessons (hard-won)

- **Agent-reported results are CLAIMS.** A builder reported balanced
  procedure counts for a file it delivered TRUNCATED (v1.14.1); the user's
  compile caught it. Senior re-runs all structural checks on the saved file.
- **File-integrity gate** (skill section 10): every saved .bas (delivery OR
  backup) — re-read from disk, check tail, count procedures, DIFF THE
  PROCEDURE LIST vs previous (count balance can't catch a wholly deleted
  proc), write a manifest. The v1.15 BACKUP was found truncated at line
  1240 a day later — backups are deliverables.
- **NEVER move deliverables through the sandbox shell** — the mount lags
  host writes (saw 1762 of 1780 lines once; `sync` + waiting did NOT help;
  reconfirmed 2026-07-10 on .md files, and AGAIN in the v1.22 build: bash
  tail/wc saw a stale 2010-line file while host Read had 2106). Promote
  host-side (Read → Write), verify host-side. Sandbox = computation only.
- **Cross-validator earns its cost**: in three audits it corrected upstream
  claims (overstated CRITICAL severity, wrong counts, wrong comment). Add
  it between auditors and senior for audit tasks.
- **Critic-layer track record** (why the independent Sonnet critic stays ON
  for logic changes): it caught a BLOCKING defect the builder's self-check
  missed on three consecutive builds — v1.21 (Intersect->Nothing crash on
  the selection Subs, section 10), v1.24 (unclosed BOM workbook on the
  fallback-InputBox path, section 12), CreoBackup v1.0 (RunCrash: unclosed
  Creo connection, section 13). Senior cross-validation then caught what the
  critic itself missed (a second v1.21 empty-list ReDim crash; a v1.22
  critic misattribution). Every layer catches something the previous one
  passed — skill section 8 rule 4.
- **Docs agents invent verification status** ("has been tested") — audit
  every such claim. A plausible-sounding wrong comment is the hardest
  defect class (three "0 = clean run" comments shipped past builder AND
  auditor).
- **Error markers separate data vs tool problems**: "(missing)" = part
  lacks the param (user fixes data); "(can't open)" = drawing broken;
  "(error)" = the READ failed (user reports the tool). One marker for both
  would make a systematic tool failure look like mass data errors.
- **Test hygiene**: observations taken in an invalid config (second Creo
  session open) are not defect evidence — re-run clean first. Forcing one
  output to fail: block the path with a FOLDER named like the target file
  (file can't overwrite folder → XToolkitCantWrite FAIL for that half,
  batch continues) — clean, reversible, verified technique.
- **A verified version is an asset**: any change voids test results.
  Bundle refactors with the next feature change; never ship them alone.
  Deferred refactors: COL_* constants, dedup of filter-parse/E1-validation
  blocks, one ERASE_SWEEP_EVERY const.
- **Requirements**: ask what the naming SCHEME means (suffix IS part of the
  part number → exact match); "sort" meant filter, "include suppressed"
  meant exclude — restate + confirm before coding.
- **A "new" symptom after removing a louder one** may be an old condition
  becoming visible (model-window parade "appeared" when the PDF popup was
  removed; code was byte-identical) — confirm the code actually changed
  before hunting a regression.

## 7. Design facts (v1.20 baseline; UI layout superseded by the v1.22+ band, see section 11)

- Excel sheet as UI; SETUP macro builds all; user never opens the VBA
  editor. Buttons stacked in column I (SUPERSEDED v1.22+ by the pixel-offset
  band; see section 11).
- Per-file error handling; FAIL text in status cell; batch never stops;
  OK/FAIL counts at end. Resume skips OK rows; Reset status re-enables.
- Memory: erase per file + EraseUndisplayedModels every 25 files.
- Names: strip trailing `.N` version, dedupe; filter = starts-with keywords
  ("M A B"); family tables: generic only (instances have no file).
- Exports: STEP ap214 colors; PDF pentable + 600 dpi, no viewer popup, all
  sheets; DXF last sheet only (v1.20 pattern, section 3); logs killed.
- Export-only invariant PROVEN line-by-line (v1.15 audit): no statement can
  touch .prt/.asm/.drw; Kill limited to .stp/.pdf/.dxf/.log under
  <outRoot>\STEP|PDF|DXF\; zero Save/Backup/Rename on models. Sheet
  mutation restored after export AND on failure (manual-Save contamination
  guard). B2 = source folder → Yes/No warning.
- In-Creo model-window parade during runs = REQUIRED export precondition
  (D8: the API needs a displayed window). Removing it = Stage 2 spec.
  Accepted: Creo drops trail/logs in its working dir; SETUP re-run wipes
  the Files sheet.
- Param columns: C Description / D Material (module constants PARAM_DESC =
  "DESCRIPTION", PARAM_MAT = "MATERIALID"); silent-connect fill at scan
  ("?" if Creo closed), loud
  Refresh button re-reads ALL rows (session values for open models, disk
  for closed); >50 rows = convenience prompt, not a cap.

## 8. Version history (all VERIFIED unless noted)

v1.0-v1.7: connect, STEP/PDF/DXF export, filter, name-list exact match.
v1.8-v1.13: ASM tree scan (active-only), per-drawing PDF/DXF/BOTH (col E→G),
Clear session, E3 asm cell. 367-file run: 355 OK, 12 FAIL (renamed files).
v1.14/.1: Description/Material columns + "(missing)/(can't open)/(error)"
markers; production 172 files 7:38. Truncated-delivery incident (section 6).
v1.15: security audit fixes (sheet restore S3-A — later found buggy; B2
warning); backup-truncation incident. v1.16: Stage 1 — no PDF popups, no
focus steal (win.Activate removed), col H timing + TimeLog sheet, resume +
Reset status; production 7.2 min vs 7:38 baseline. v1.17: sheet-set
confirm-retry (falsified theory, kept as diagnostic). v1.18: post-export
file-stability wait before sheet restore (falsified theory, kept as
correctness guard). v1.19: Export2DOption attach — DEAD END (section 4),
never production. v1.20 (VERIFIED 2026-07-10): attach removed,
mdl.Display kept = THE DXF fix (section 5). TimeLog rows: 6 retries,
7 max wait, per-type table row 9. v1.21 (VERIFIED 2026-07-11,
full 8-step protocol on 390-part real list): connect diagnostics
(PRO_COMM_MSG_EXE check + two-session hint), pre-export audit popup,
StatusBar progress, live B3 highlight, Toggle-tick/Cycle-Fmt selection
buttons, regrouped buttons (green EXPORT zone / amber danger zone),
retry-once rider + mParamOpenRetries counter (section 9).
v1.22-v1.23 (VERIFIED 2026-07-12): freeze panes + Ctrl+Shift shortcut keys
+ a scroll-independent button band (closes the v1.21 stationary-button gap;
v1.22 point-vs-pixel overlap bug fixed in v1.23 — section 11) + the
same-source-folder HARD BLOCK. v1.24 (VERIFIED 2026-07-12): CSV/BOM import
(611/809 exact on the real list) + button-polish rider; Stage 2 B0 speed
baseline captured (510 files / 590.5 s — section 12). v1.25 (CURRENT,
VERIFIED 2026-07-12): B3 filter dropdown + tooltip legend + typed combos
(section 12). Separate module CreoBackup.bas v1.1 (CURRENT, VERIFIED
2026-07-12, full 10-step protocol): batch IpfcModel.Backup of an assembly's
active components + matching drawings, names unchanged, in its own workbook
sheet (section 13).

Still open: spot-check old delivered DXF (v1.16 or earlier, silent
wrong-sheet risk — zero-code re-export with v1.20+); Stage 2 speed
experiments (spec written 2026-07-12, S2-A lock-screen run already PASS);
REOPENED intermittent PRT "(can't open)" at scan (WATCH AND WAIT, section
9). DONE since: obsolete v1.15/v1.18 backups deleted (2026-07-11);
connect-MsgBox two-session hint shipped in v1.21.

## 9. REOPENED: intermittent PRT "(can't open)" params (2026-07-10)

SYMPTOM (Nam, recurring since ~v1.18): at scan/refresh some PRT rows show
Description "(can't open)" + BLANK Material although the parts have both
params; the affected COUNT varies between scans of the same folder.

- LESSON (bug-closure discipline): this was closed once ("item 7") after a
  SINGLE clean isolation run. An INTERMITTENT symptom can never be closed
  by one passing run - that is the sample-limited-verification lesson
  applied to bug closure. Closing it was a mistake; reopened.
- CODE FACTS (v1.20 read 2026-07-10): PRT-row "(can't open)" + blank D has
  exactly one producer pair - ReadTwoParams' OpenFail handler (fires when
  session.RetrieveModel throws for the part) or FillParams' per-row wrap
  (fires when ReadTwoParams itself raises). Param-READ failures show
  "(error)", absent params "(missing)" - so these rows failed at OPEN.
  Varying counts => TRANSIENT retrieve failures, not per-file data.
  E1=DXF has no code path into param reads (re-verified) - the plausible
  real correlate is session state: scanning right after an export batch,
  many models in session, Creo busy; possibly interplay with the
  every-25-rows EraseUndisplayedModels sweep inside FillParams.
- MARKER-RULE VIOLATION noted: for PRT rows, a transient session problem
  currently wears the same "(can't open)" marker as a permanently broken
  file - a tool problem displayed as a data problem. Any eventual fix must
  restore the distinction (e.g. retry once, then a transient-specific
  marker), but ONLY after the root cause is proven.
- PROTOCOL RESULT (Nam, 2026-07-11, after closing/reopening everything):
  single xtop.exe confirmed; fresh Creo scan x2 = 0 failures; scan right
  after a small export batch = 0 failures. NOT REPRODUCIBLE in a clean
  state. Proven safe: fresh session, single Creo, small post-export scan.
  NOT tested (never reproduced): Refresh-heals, same-rows-or-random.
  Suspected failing condition: LONG-RUNNING session (hours/days, several
  big batches) or machine load at that moment - unproven.
- DECISION (Nam, 2026-07-11): WATCH AND WAIT. No code change; v1.20 stays
  a verified asset. When it recurs in real use, Nam records (1) how long
  Creo has been up + what ran before, (2) whether "Refresh params" heals
  the marked rows. Durability fix (retry-once on PRT open + distinct
  transient marker + retry counter for field data) is SPECCED IN CONCEPT
  and gets bundled with the NEXT code change, whatever it is.
- RIDER DELIVERED in v1.21 (same day): ReadTwoParams retries the initial
  RetrieveModel once after YieldMs 300; successful retry increments
  mParamOpenRetries (shown in Refresh params MsgBox + pre-export audit).
  A nonzero count = the transient failure happened and self-healed.

## 10. v1.21 build lessons (2026-07-11)

- CRITIC CAUGHT A BLOCKING DEFECT the builder's self-check missed (running
  tally of critic/senior catches lives in section 6): both new selection
  Subs called .Cells on Intersect(...) that returns Nothing whenever a
  selected area misses column A (the realistic case: users select cells
  in the E or G columns the buttons act on) -> unhandled runtime 91.
  Senior review then found a SECOND crash the critic missed: ReDim
  rows(1 To 0) on an empty file list.
- FIX PATTERN (reusable): for "act on selected rows" VBA, intersect
  Selection area's .EntireRow with the data range's key column and guard
  Nothing; guard the empty-list ReDim with an early "No files listed"
  exit. Selecting cells in ANY column then selects those rows.
- VBA facts verified this build: form-control Buttons cannot be colored
  (zone CELLS behind them instead); Optional params on a Private Sub are
  fine for existing callers; FormatConditions.Delete before .Add keeps
  SETUP idempotent; a variable named rows() shadows nothing harmful
  inside its own procedure (but avoid the name in new code anyway).
- Gemini-review triage lesson: an external review's PREMISES must be
  checked against the actual code before acting - 4 of 12 items rested
  on false claims (blank-B2 behavior, "no progress indicator", "config
  needs VBA editor", modal-removal-as-speedup), and 2 more would have
  reintroduced deliberately-designed-out bugs (stale-OK re-scan, split
  BOTH resume). Reviews suggest; the lessons doc + code decide.

## 11. v1.22/v1.23 build lessons (2026-07-11 / 2026-07-12)

- Verified API (safe to reuse): `Application.MacroOptions Macro:="Name",
  HasShortcutKey:=True, ShortcutKey:="T"` - UPPERCASE letter =
  Ctrl+Shift+letter, persists in the workbook (unlike OnKey = session
  only), target must be a Public parameterless Sub. MS Learn checked.
  FreezePanes pattern: ws.Activate -> FreezePanes = False -> select the
  split cell -> FreezePanes = True (Activate first, else runtime 1004;
  False-then-True keeps SETUP idempotent).
- Self-check DESIGN error (senior's, caught by builder): "grep the new
  parameter name at call sites" fails when arguments are passed
  POSITIONALLY - VBA calls carry the value, not the name. Write
  grep-based checks against strings that actually appear in code.
- Critic misattribution (caught by senior cross-validation; tally in
  section 6): critic quoted the SPEC's wording as if it were a CODE
  comment and flagged the comment as overstating. The code was accurate;
  the spec got a one-line build note instead of touching verified code.
  Check WHERE a quoted string lives before editing anything.
- Layout physics for Excel button bands: a vertical stack cannot be
  frozen (freeze is row-based); a naive one-button-per-column band
  right of the data overflows 1920 px. Fix: pixel offsets from ONE
  anchor cell (AddBtn dxPx) so the band starts inside wide unused
  column F. Zone colors = fixed cell-range estimates, cosmetic only.
- USER TEST RESULT (2026-07-12, 7/8 PASS): the band WORKED (buttons
  fire, shortcuts, freeze, idempotent SETUP, regression export OK) but
  overlapped rows 3-4, hiding the header row + I4 reminder; zones
  landed misplaced; 100% zoom needed horizontal scroll. ROOT CAUSE OF
  ALL THREE: **Excel Buttons.Add Left/Top/Width/Height and RowHeight/
  Column .Left/.Width are in POINTS, not pixels.** The v1.22 geometry
  was designed in pixels: a 26pt button anchored at row 3 (top=30pt)
  reaches 56pt while row 4 starts at 45pt -> 11pt overlap on EVERY
  machine (deterministic, not DPI). The "118 px" pitch was really
  118pt = ~157px, blowing the band past 1920px at 100% zoom.
  PREVENTION: design ALL Excel shape/row/column geometry in points
  (row default = 15pt); verify stacked elements as sum-of-points
  BEFORE delivery; never trust px estimates. Fixes (v1.23): explicit
  RowHeight for band rows, tighter point pitch, and a dynamic zone
  painter that finds columns by comparing Columns(c).Left/.Width
  (points) against button x-ranges instead of hardcoding cell letters.
- v1.23 USER-TESTED 2026-07-12: all 8 protocol steps FUNCTIONAL PASS
  (incl. hard-block broken case: popup fired, zero disk writes) ->
  verified baseline. Residual cosmetics (accepted for now, polish only
  as a bundled v1.24): zone cell-fills snap to COLUMN boundaries while
  buttons sit at arbitrary point offsets -> edge overlap is inherent;
  perfect alignment would need the band on dedicated narrow columns
  right of the data, which pushes the band off-screen (~1750pt) — dead
  end, don't revisit. Nam's screen still needs ~80% zoom for the full
  band: point->pixel size depends on Windows display SCALE (125% =
  1.25x pixels per point), so a band that fits 1920px at 100% scale
  overflows at 125% — ask the user's display scale BEFORE sizing
  full-width Excel UI. Danger pair (Reset/Clear) stacked at one dx =
  misclick risk; cheap fix = diagonal stagger (different dx per row).

## 12. v1.24/v1.25 build lessons (2026-07-12)

- CRITIC CAUGHT A BLOCKING DEFECT the builder's self-check passed (tally in
  section 6): in ImportBomList the section AFTER Workbooks.Open had no error
  handler — a fat-fingered column number in the fallback InputBox (realistic
  for a non-coder) would raise a raw VBA error and leave the user's BOM file
  open/locked. RULE (reusable — this is the canonical statement; the
  CreoBackup build in section 13 generalized it to a held Creo connection):
  any Sub that OPENS/acquires a resource (file, workbook, or connection)
  must guarantee the release on EVERY path — arm "On Error GoTo <CloseLabel>"
  immediately after the successful Open/Connect, re-arm it (never GoTo 0)
  after any tight Resume-Next block, and give the handler its own Resume
  Next around the Close/Disconnect itself.
- BaseCreoName is FILENAME-ONLY: returns "" when the input has no
  .prt/.asm/.drw extension — NEVER feed it bare part numbers (BOM PNs).
  CleanBomName is the pass-through-safe variant. Verify a helper's
  contract in code before reusing it; the name suggested more than it does.
- Real-file facts (A210265.csv, verified 2026-07-12): Creo BOM exports
  here are comma CSVs; header row 1; the PN header is "PN " WITH a
  trailing space (Trim headers before matching!); duplicates are normal
  (fasteners x11); PNs never contain spaces/dots; _N suffixes are part
  of the name. 809 data rows -> 611 unique.
- InputBox bounds discipline: numeric-looking input still needs a range
  check against the file's real column count, and REJECTIONS must say
  why (silent cancel = confused user, wasted round trip).
- v1.24 USER-TESTED 2026-07-12: all protocol steps PASS (611/809 exact
  on the real BOM; production regression clean). STAGE 2 B0 BASELINE
  captured: 510 files = 210 PDF + 36 DXF + 264 STEP in 590.5 s
  (~1.16 s/file average) on the production laptop, v1.24, quality
  spot-checked (pentable, DXF last sheet). Compare every Stage 2
  experiment against these numbers. Untested residual: the no-PN-header
  column prompt path (critic-verified in code only).
- v1.25 USER-TESTED 2026-07-12: 4/4 PASS (dropdown, tooltip legend,
  typed combos, regression with manual quality checks). Validation
  facts verified in the field: xlValidateList + ShowError=False gives
  a suggest-only dropdown (typed non-list values accepted silently);
  InputMessage tooltip = free-real-estate legend when the sheet has no
  visibly free cell; Validation coexists with FormatConditions on the
  same cell.
- **S2-A LOCK-SCREEN: PASS (2026-07-12, production evidence).** The
  510-file B0 batch ran with the desktop LOCKED and completed; outputs
  = the quality-checked set. FACT: Creo window creation/display and
  all three exporters work while the desktop is locked (Windows does
  not suspend processes on lock; GDI concern did not materialize).
  Locked unattended runs approved. Prereqs: Sleep = Never, lid-close =
  Do nothing. Hibernate/sleep still stops Excel+Creo+COM dead.

## 13. CreoBackup v1.0/v1.1 build lessons (2026-07-12)

- **UTF-8 vs ANSI .bas TRAP (v1.0 field test, step 5): any non-ASCII
  character in a delivered .bas becomes MOJIBAKE after import.** The
  Write tool saves UTF-8; the VBA editor imports .bas as ANSI, so an
  em dash "—" renders as three garbage characters in cells, MsgBoxes
  and StatusBar (logic unaffected — comparisons used "YES" only).
  PREVENTION RULE: .bas string literals and comments are ASCII-ONLY
  (hyphens, not em dashes; no curly quotes, no ellipsis). Add a
  non-ASCII grep (0 required) to every .bas integrity gate + builder
  self-check. Fixed in v1.1; v1.1 light retest PASS 2026-07-12 (compile
  + scan + small run) -> v1.1 = VERIFIED BASELINE. Rule promoted to
  skill master section 10 same day.
- Excel auto-converts date-looking STRINGS written to cells into real
  dates -> narrow column shows ##########. Force text with a leading
  apostrophe when a timestamp must display in a narrow cell (v1.1).
- API **FIELD-VERIFIED 2026-07-12 (Nam, full 10-step protocol PASS)**:
  `IpfcModel.Backup(WhereTo As IpfcModelDescriptor)` = File > Save a
  Backup. Pattern: fresh descriptor via CreateFromFileName, set
  `.Path` = destination folder, `mdl.Backup descr`. CONFIRMED IN THE
  FIELD: (a) Backup DOES honor descriptor.Path as the target folder;
  (b) Backup needs NO window/Display (save-type op, not an exporter);
  (c) drawing backup carries its models; asm backup carries all
  components; batch per-item FAIL isolation works (renamed-.drw
  broken case: one FAIL row, batch continued). Creo behavior (PTC docs +
  community, matches Nam's field report): backing up an asm/drawing/
  mfg saves ALL dependents — a drawing backup carries its models; an
  asm backup carries all components INCLUDING suppressed ones (cannot
  be filtered).
- `rename_drawings_with_object` (free config.pro) auto-copies drawings
  ONLY for Save As / Save a Copy / Rename — all force a NEW name. It
  does NOT apply to Backup. Checked 2026-07-12; don't re-research.
- CRITIC CAUGHT A BLOCKING DEFECT the builder's self-check passed (tally in
  section 6): BackupRun's loop had no outer error trap — an unexpected
  non-Creo error (locked sheet, COM hiccup) would strand the Creo connection
  and leave the StatusBar stuck. Fix = RunCrash net armed after connect,
  re-armed (never GoTo 0) after each Resume-Next block — the same
  guarantee-release-on-EVERY-path rule stated in section 12, applied here to
  the held Creo connection.
- Two-module workbook rule: a second .bas in the same workbook
  compiles into the SAME global namespace — duplicate PUBLIC names =
  "Ambiguous name" compile error. Convention used: exactly 6 Public
  entry points with a Backup* prefix, everything else Private with a
  Bk prefix. Check the other module's Public list BEFORE naming.
- Cowork subagent gotcha: a fresh builder agent answered the MCP
  auth notices instead of the task prompt (lost the whole prompt).
  Fix: prefix subagent prompts with "TASK: ... ignore MCP notices"
  and re-spawn once if the reply contains no work product.

## 14. THE SPLIT — CreoTree extraction + CreoBOM v2.0 (2026-07-16, BUILT, awaiting test)

- Design decision (Fable-5 senior arch view, arch-note Revision 2):
  the four tools link by WORKFLOW, not by shared mutable state. The
  quantity rollup that two pipelines need lives in ONE stateless helper
  `CreoTree.bas` (pure functions: no sheet, no shortcut, no Creo). This
  is the SAFE kind of coupling - categorically unlike the shared-sheet
  coupling that silently broke Ctrl+Shift+T. Consumers: CreoBOM v2.0
  (bought parts) now; CreoPackage v2.0 (machined parts) in Spec 2.
- STANDING RULE (write into any future CreoTree change): a change to
  CreoTree voids BOTH consumers' verification -> retest BOTH. Design the
  seam (parts=nested dict keyed UCase PN with name/qty/desc/mat/kind;
  problems=dict of Array(pn,reason,detail)) for both callers up front.
- Extraction technique that worked: builder COPIES the verified regions
  byte-for-byte (do not re-derive), adds only the classify step, returns
  via dictionaries instead of writing sheet rows. Independent critic
  (Fable-5) then DIFFS the moved regions against the original - "identical"
  is the pass. This caught nothing major here precisely because the rule
  was "copy, don't rewrite"; the one MINOR was in the NEW glue (blank
  Problem sheet when TreeRollup returns False on an all-rows-bad csv -
  the caller must flush `problems` even on the False path).
- Mount-lag RECURRED and was caught (section 10 rule 1 held): right after
  a senior Edit, `mcp__workspace__bash` reported CreoBOM.bas truncated
  mid-line at ~839 lines with an unbalanced proc count. Host-side Read +
  Grep confirmed 855 lines, 31/31 balanced, clean tail. LESSON REINFORCED:
  never trust the sandbox shell to verify a just-written file; host tools
  are authoritative. Structural gates run on host Read/Grep, not bash.
- Refactor cost honored: extracting from verified CreoBOM v1.3 voids its
  baseline. Success is defined as EQUIVALENCE (bought output identical to
  v1.3 on With_assembly.csv), not "no error" - the known numbers
  (42 rows / 122 pieces / the B+FC counts) are the objective check.

## 15. CreoPackage v2.0 build - 2nd CreoTree consumer (2026-07-17, BUILT, awaiting test)

- CreoPackage became the MANUFACTURING pipeline: two-step (BUILD LIST via
  CreoTree kind=MFG, then BUILD PACKAGE fills the template + matches drawings).
  It is now CreoTree's SECOND consumer, so the standing rule bites: a CreoTree
  change voids CreoBOM AND CreoPackage - retest both.
- NEW LESSON (Fable-5 critic MAJOR-1): a helper that opens a workbook and runs
  inside a caller's `On Error GoTo Fail` MUST close that workbook on its own
  error before re-raising. Otherwise the open handle LOCKS the package folder
  the caller's error message tells the user to delete - a self-inflicted
  "cannot delete" trap. Pattern: local `On Error GoTo Fail` in the helper ->
  `If Not wb Is Nothing Then wb.Close SaveChanges:=False` -> `Err.Raise` to
  bubble back. Same discipline the txt-handle already had (close on error so it
  doesn't lock the folder). Applies to ANY Open/Save inside a build that a Fail
  handler tells the user to clean up.
- Read-only cross-tool link done safely: Package auto-links its export root by
  READING CreoExport Files!B2 (else B1\EXPORT). Rules that keep it out of the
  14b coupling class: read one cell, never write CreoExport's sheet; manual
  value wins; graceful "" fallback if the Files sheet is missing; and persist
  the auto-resolved value into Package's OWN cell only AFTER it validates
  (critic MINOR-2 - a stale path pinned into the cell would suppress the
  auto-link forever, since "manual wins").
- Validate-before-persist and refuse-don't-coerce recurred as minors: a picker
  that reports success must first confirm the folder exists + has PDF\/STEP\
  (MINOR-3); a Type cell must be REFUSED if invalid at package time, not
  silently coerced to NORMAL and written to the shop xlsx (MINOR-4, "loud miss"
  doctrine). Guard `Selection` with `TypeName(Selection)<>"Range"` before
  `.EntireRow` in any button that reads the selection (MINOR-5, same class as
  the v1.21 CreoExport catch).
- xlsm->xlsx copy trap (MINOR-6): FileCopy of a .xlsm template to a hard-coded
  .xlsx name yields a content/extension mismatch Excel may refuse. Fix: restrict
  the template picker to *.xlsx (or SaveAs FileFormat:=51). Was latent in
  CreoBOM v1.3 too.
- V4 BUNDLE LESSONS (2026-07-17, critic catches on NEW code only - the
  no-logic-edit rule kept all verified bodies clean):
  (a) FROZEN-PANE BUTTON TRAP: a button column that fits the BAND width
  can still cross the FREEZE LINE and scroll away with the data (Export
  freeze at A5 = ~72pt tall; Backup rows 1-5 x 18pt = exactly 90pt).
  Check button bottoms against the freeze row's cumulative HEIGHT, not
  just against other buttons. 2x2 grids / shorter buttons fix it.
  (b) MsgBox SILENTLY TRUNCATES at ~1024 chars. Any report that
  concatenates many full paths (worse: "was X, now Y" = two paths per
  line) must be chunked into multiple boxes past ~900 chars - the
  truncated tail is invisible, and it eats exactly the warnings the
  report exists to deliver.
  (c) Alt+F8 hiding via `Optional ByVal dummyHide As Variant` shipped on
  ~40 Publics; shortcut targets (ToggleTick/ToggleTickSelected/
  CycleFmtSelected) deliberately stay parameterless per section 11's
  verified MacroOptions form - UNVERIFIED whether a dummy-arg sub still
  fires via shortcut, so we did not gamble three working shortcuts on it.
  >>> FIELD-REFUTED 2026-07-17 (v4b): the Optional-dummyHide trick did
  NOT hide macros in Nam's Excel - every dummyHide sub still listed in
  Alt+F8 (screenshot evidence). The claim came from the architecture
  discussion and shipped unverified - exactly the never-guess violation
  the rules exist to stop. REPLACEMENT (v4b rider): `Option Private
  Module` in all five tool/master modules (documented: removes ALL the
  module's macros from the Macro dialog) + a new normal module CreoMenu
  as the only visible surface (SETUP_1_MASTER..SETUP_5_BACKUP wrappers +
  KeyB/KeyT/KeyC shortcut wrappers; MacroOptions re-registered onto the
  wrappers - parameterless Publics in a normal module = the verified
  form). REMAINING UNVERIFIED FACT (flagged to Nam, tested FIRST in the
  protocol): sheet-button OnAction firing into an Option Private Module
  sub. If it fails, every button errors loudly, no data risk; rollback =
  delete the five one-line Option Private Module statements. dummyHide
  signatures were left in place (harmless, and buttons already point at
  them).
  (e) Naming rider (Nam): BackupSetup was invisible among Backup* names -
  the numbered SETUP_1..SETUP_5 wrapper convention fixes discoverability
  for good.
  (d) A pushed default path the tool then REFUSES (nonexistent
  <export root>\PACKAGE) = advertised happy path that fails on first
  use. When a config layer invents a default, the consuming tool must
  either create it (one safe level) or the spec must say who does.
- FIELD ERROR + WRONG FIRST DIAGNOSIS (2026-07-17, compile): `Dim eNum As
  Long` = bare "Syntax error". ROOT CAUSE: VBA identifiers are
  CASE-INSENSITIVE, so `eNum` collides with the reserved keyword `Enum`.
  The senior's first diagnosis ("Dim not allowed after a line label") was
  WRONG - a Dim after a label is legal VBA - and cost one round trip: the
  moved line failed identically, which is what exposed the real cause
  (same statement failing in two positions = the statement, not the
  position). PREVENTION: (a) never name a variable a case-variant of any
  VBA keyword (eNum/Enum, tYpe/Type, nExt/Next...); (b) when a compile
  error survives a fix, re-suspect the DIAGNOSIS, not just the file
  (extends section 10 rule 4); (c) fix applied: errNum/errDesc.

## 16. CreoRegister spec facts (2026-10-02, spec stage, nothing built yet)

- Part_number_register.xlsx (shared drive, protected): M/A Part No is a
  formula ="M"&D&E; Number sequence (col E) is GLOBAL per sheet, not per
  project. Project ID col D is General-formatted -> writing "03" as a value
  becomes 3 -> Part No "M30053". Write text with an apostrophe prefix;
  Assemblies protection forbids format changes, so NumberFormat="@" is not
  an option there.
- Rows with a Project ID but no Description = used (reserved/abandoned).
- API names for this tool: session.CurrentModel, IpfcModel.InstanceName -
  confirmed via search-result snippets only (support.ptc.com and
  community.ptc.com are blocked from the cloud sandbox). Prove them on the
  user's machine in test step 1 before trusting them.
