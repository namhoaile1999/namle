Attribute VB_Name = "CreoRegister"
' =============================================================================
' CreoRegister v1.0   (2026-10-02)
'
' Registers the ACTIVE Creo part/assembly in the shared Part Number Register.
'   M + 2-digit project + 4-digit number  -> sheet "Manufactured parts"
'   A + 2-digit project + 4-digit number  -> sheet "Assemblies"
'   B...  (exact Part No match)           -> sheet "Bought parts"
' It writes ONLY: Description (col C) and, on M/A, Project ID (col D) as text.
' It refuses when the row is used, the register is busy, or anything is unsure.
'
' Setup once: Tools > References > tick "Creo VB API Type Library" (exactly
' one), run RegSetup, click "Pick register file".
' Daily use:  create the part in Creo (name + DESCRIPTION), then click
'             "REGISTER ACTIVE CREO PART" or press Ctrl+Shift+R.
'
' Public macros (everything else is Private):
'   RegSetup, RegPickRegister, RegRefreshNextFree, RegRegisterActivePart
' Spec: docs/CreoRegister_SPEC.md
' =============================================================================
Option Explicit

Private Const TOOL_VERSION As String = "CreoRegister v1.0"

' ---- tool workbook layout ----
Private Const UI_SHEET As String = "Register"
Private Const LOG_SHEET As String = "Log"
Private Const PATH_CELL As String = "B3"
Private Const PANEL_FIRST_ROW As Long = 7      ' rows 7, 8, 9 = M, A, B
Private Const REFRESHED_CELL As String = "B11"
Private Const RESULT_CELL As String = "B13"
Private Const BUTTON_ROW As Long = 15

' ---- Creo ----
Private Const PARAM_DESC As String = "DESCRIPTION"
Private Const TEMPLATE_DESC As String = "START_PART"   ' start_part default value

' ---- register layout (checked against the real file 2026-10-02) ----
Private Const SH_MFG As String = "Manufactured parts"
Private Const SH_ASM As String = "Assemblies"
Private Const SH_BUY As String = "Bought parts"
Private Const SH_INDEX As String = "Index"
Private Const HEADER_ROW As Long = 2
Private Const FIRST_ROW As Long = 3
Private Const COL_PN As Long = 2          ' B  Part No (formula on M/A)
Private Const COL_DESC As Long = 3        ' C  Description
Private Const COL_PROJ As Long = 4        ' D  Project ID (M/A)
Private Const COL_SEQ As Long = 5         ' E  Number sequence (M/A)
Private Const COL_BUY_EXTRA_LAST As Long = 6   ' Bought: D..F = MPN, supplier, note
Private Const BLOCK_COLS As Long = 6      ' columns A..F are read in one block

' =============================================================================
' PUBLIC MACROS
' =============================================================================

' Builds (or rebuilds) the Register sheet and buttons. Keeps the register path.
' Never clears the Log sheet (it is the audit trail).
Public Sub RegSetup()
    Dim ws As Worksheet, lg As Worksheet
    Dim keepPath As String, shortcutNote As String

    Set ws = rgEnsureSheet(UI_SHEET)
    keepPath = rgText(ws.Range(PATH_CELL).Value)

    On Error Resume Next   ' Buttons.Delete may raise when there are no buttons yet
    ws.Buttons.Delete
    On Error GoTo 0
    ws.Cells.Clear

    ws.Range("A1").Value = TOOL_VERSION & " - register the active Creo part in the Part Number Register"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 13
    ws.Range("A3").Value = "Register file:"
    ws.Range("A3").Font.Bold = True
    If Len(keepPath) > 0 Then ws.Range(PATH_CELL).Value = "'" & keepPath

    ws.Range("A5").Value = "Next free numbers (click Refresh; re-checked at every Register click)"
    ws.Range("A5").Font.Bold = True
    ws.Range("A6").Value = "Sheet"
    ws.Range("B6").Value = "Next free number"
    ws.Range("C6").Value = "Type in Creo as"
    ws.Range("D6").Value = "Register row"
    ws.Range("E6").Value = "Prepared rows left"
    ws.Range("F6").Value = "Last used description"
    ws.Range("A6:F6").Font.Bold = True
    ws.Range("A6:F6").Interior.Color = RGB(221, 235, 247)
    ws.Cells(PANEL_FIRST_ROW, 1).Value = SH_MFG & " (M)"
    ws.Cells(PANEL_FIRST_ROW + 1, 1).Value = SH_ASM & " (A)"
    ws.Cells(PANEL_FIRST_ROW + 2, 1).Value = SH_BUY & " (B)"

    ws.Range("A11").Value = "Refreshed at:"
    ws.Range("A13").Value = "Last result:"
    ws.Range("A11,A13").Font.Bold = True
    ws.Range("A17").Value = "## in 'Type in Creo as' = your 2-digit project ID (see the register's Index sheet)."
    ws.Range("A18").Value = "Workflow: Creo File > New > type the name > fill DESCRIPTION > OK, then click REGISTER (Ctrl+Shift+R)."
    ws.Range("A17:A18").Font.Italic = True

    ws.Columns("A").ColumnWidth = 26
    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C").ColumnWidth = 22
    ws.Columns("D").ColumnWidth = 13
    ws.Columns("E").ColumnWidth = 18
    ws.Columns("F").ColumnWidth = 50

    ' Buttons: geometry in POINTS (lessons doc section 11). Row made tall
    ' enough that 28pt buttons never overlap neighbouring rows.
    ws.Rows(BUTTON_ROW).RowHeight = 34
    rgAddButton ws, ws.Cells(BUTTON_ROW, 1), 0, 135, "Pick register file", "RegPickRegister"
    rgAddButton ws, ws.Cells(BUTTON_ROW, 1), 145, 135, "Refresh next free", "RegRefreshNextFree"
    rgAddButton ws, ws.Cells(BUTTON_ROW, 1), 290, 270, "REGISTER ACTIVE CREO PART (Ctrl+Shift+R)", "RegRegisterActivePart"

    Set lg = rgEnsureSheet(LOG_SHEET)
    If Len(rgText(lg.Range("A1").Value)) = 0 Then
        lg.Range("A1:I1").Value = Array("Time", "Result", "Creo name", "Sheet", "Row", _
                                        "Project", "Description", "Detail", "User")
        lg.Range("A1:I1").Font.Bold = True
        lg.Columns("A").ColumnWidth = 19
        lg.Columns("B").ColumnWidth = 11
        lg.Columns("C").ColumnWidth = 14
        lg.Columns("D").ColumnWidth = 18
        lg.Columns("E:F").ColumnWidth = 8
        lg.Columns("G").ColumnWidth = 40
        lg.Columns("H").ColumnWidth = 70
        lg.Columns("I").ColumnWidth = 16
    End If

    ' Ctrl+Shift+R. Verified form (lessons section 11): uppercase letter,
    ' Public parameterless Sub in a normal module.
    On Error Resume Next   ' a failure here only loses the shortcut; reported below
    Application.MacroOptions Macro:="RegRegisterActivePart", HasShortcutKey:=True, ShortcutKey:="R"
    If Err.Number <> 0 Then shortcutNote = vbLf & "WARNING: shortcut Ctrl+Shift+R could not be set (" & Err.Description & "). Use the button."
    On Error GoTo 0

    ws.Activate
    ws.Range(PATH_CELL).Select
    MsgBox TOOL_VERSION & " setup done." & vbLf & vbLf & _
           "Next: click 'Pick register file' and choose the Part Number Register." & shortcutNote, _
           vbInformation, TOOL_VERSION
End Sub

' Lets the user choose the register file, stores the path, refreshes the panel.
Public Sub RegPickRegister()
    Dim f As Variant
    f = Application.GetOpenFilename("Excel workbooks (*.xlsx;*.xlsm),*.xlsx;*.xlsm", , _
                                    "Pick the Part Number Register")
    If VarType(f) = vbBoolean Then Exit Sub       ' user pressed Cancel
    If StrComp(CStr(f), ThisWorkbook.FullName, vbTextCompare) = 0 Then
        MsgBox "That is this tool workbook, not the Part Number Register. Pick the register file.", _
               vbExclamation, TOOL_VERSION
        Exit Sub
    End If
    ThisWorkbook.Worksheets(UI_SHEET).Range(PATH_CELL).Value = "'" & CStr(f)
    RegRefreshNextFree
End Sub

' Opens the register READ-ONLY (never blocks colleagues), fills the
' "next free" panel, closes it again.
Public Sub RegRefreshNextFree()
    Dim wb As Workbook, openedByUs As Boolean, why As String

    On Error GoTo Crash
    Application.ScreenUpdating = False
    Application.StatusBar = "CreoRegister: reading the register..."
    If Not rgOpenRegister(rgRegisterPath(), True, wb, openedByUs, why) Then GoTo ShowProblem
    why = rgUpdatePanel(wb)
    ThisWorkbook.Worksheets(UI_SHEET).Range(REFRESHED_CELL).Value = "'" & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    GoTo CloseIt

Crash:
    why = "Unexpected error while reading the register (tool error, please report): " & _
          Err.Number & " " & Err.Description
    Resume CloseIt

CloseIt:
    On Error Resume Next   ' closing a read-only copy cannot lose data; keep going
    If openedByUs And Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0

ShowProblem:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    If Len(why) > 0 Then MsgBox why, vbExclamation, TOOL_VERSION
End Sub

' MAIN BUTTON. Reads the active Creo model, writes its row, saves the register.
Public Sub RegRegisterActivePart()
    Dim regPath As String, nm As String, desc As String, why As String
    Dim kind As String, proj As String, seq As String, sheetName As String
    Dim wb As Workbook, ws As Worksheet, openedByUs As Boolean
    Dim r As Long, curDesc As String, curProj As String, extra As String
    Dim descToWrite As String, pnNow As String, stepName As String
    Dim wroteProj As Boolean, wroteDesc As Boolean, savedOk As Boolean
    Dim outcome As String, answer As VbMsgBoxResult

    outcome = "FAIL"
    On Error GoTo Crash

    stepName = "read Creo"
    Application.StatusBar = "CreoRegister: reading the active model from Creo..."
    If Not rgReadCreo(nm, desc, why) Then GoTo FailClose

    stepName = "check name"
    If Not rgParseName(nm, kind, proj, seq, why) Then GoTo FailClose
    If Len(desc) = 0 Then
        why = "DESCRIPTION of " & nm & " is blank (missing)." & vbLf & _
              "Fill it in Creo (File > Prepare > Model Properties > Parameters), then try again."
        GoTo FailClose
    End If
    If StrComp(desc, TEMPLATE_DESC, vbTextCompare) = 0 Then
        why = "DESCRIPTION of " & nm & " is still the template default '" & TEMPLATE_DESC & "'." & vbLf & _
              "Change it in Creo, then try again."
        GoTo FailClose
    End If
    sheetName = rgSheetFor(kind)

    ' Guards against the wrong Creo window being active.
    answer = MsgBox("Register this Creo model?" & vbLf & vbLf & _
                    "Name:  " & nm & vbLf & _
                    "Description:  " & desc & vbLf & _
                    "Sheet:  " & sheetName & _
                    IIf(kind = "B", "", vbLf & "Project:  " & proj), _
                    vbYesNo + vbQuestion, TOOL_VERSION)
    If answer <> vbYes Then
        outcome = "CANCELLED"
        why = "Cancelled by user before anything was written."
        GoTo FailClose
    End If

    stepName = "open register"
    Application.StatusBar = "CreoRegister: opening the register..."
    Application.ScreenUpdating = False
    regPath = rgRegisterPath()
    If Not rgOpenRegister(regPath, False, wb, openedByUs, why) Then GoTo FailClose

    stepName = "check register layout"
    Set ws = rgGetSheet(wb, sheetName)
    If ws Is Nothing Then
        why = "The register has no sheet named '" & sheetName & "'."
        GoTo FailClose
    End If
    why = rgHeaderProblem(ws, kind)
    If Len(why) > 0 Then
        why = "The register layout changed on sheet '" & sheetName & "': " & why & vbLf & _
              "This tool version does not know the new layout - nothing written."
        GoTo FailClose
    End If

    stepName = "find row"
    If kind = "B" Then
        r = rgFindRow(ws, COL_PN, nm, 0, why)
    Else
        If Not rgProjectExists(wb, proj, why) Then GoTo FailClose
        r = rgFindRow(ws, COL_SEQ, seq, 4, why)
    End If
    If r = 0 Then GoTo FailClose

    stepName = "check row is free"
    curDesc = rgText(ws.Cells(r, COL_DESC).Value)
    If kind = "B" Then
        If Len(curDesc) > 0 Then
            If curDesc = desc Then GoTo AlreadyDone
            why = nm & " is already used in '" & sheetName & "' row " & r & ":" & vbLf & _
                  "  Description: " & curDesc & vbLf & _
                  "Use the next free B number (click Refresh next free)."
            GoTo FailClose
        End If
        extra = rgBoughtExtras(ws, r)
        If Len(extra) > 0 Then
            Application.ScreenUpdating = True
            answer = MsgBox(nm & " (row " & r & ") has no description, but other cells are filled:" & vbLf & _
                            extra & vbLf & vbLf & "Someone may have reserved it. Write the description anyway?", _
                            vbYesNo + vbExclamation + vbDefaultButton2, TOOL_VERSION)
            Application.ScreenUpdating = False
            If answer <> vbYes Then
                outcome = "CANCELLED"
                why = "Cancelled: row " & r & " already has other data (" & extra & ")."
                GoTo FailClose
            End If
        End If
    Else
        curProj = rgPad(ws.Cells(r, COL_PROJ).Value, 2)
        If Len(curDesc) > 0 Or Len(curProj) > 0 Then
            If curProj = proj And curDesc = desc Then GoTo AlreadyDone
            why = nm & " cannot be registered: number " & seq & " (row " & r & ") is already used." & vbLf & _
                  "  Project ID: " & IIf(Len(curProj) = 0, "(blank)", curProj) & vbLf & _
                  "  Description: " & IIf(Len(curDesc) = 0, "(blank)", curDesc) & vbLf & _
                  "Rename the part in Creo to a free number (click Refresh next free)."
            GoTo FailClose
        End If
    End If

    If rgCellBlocked(ws, r, COL_DESC) Or (kind <> "B" And rgCellBlocked(ws, r, COL_PROJ)) Then
        why = "Row " & r & " on '" & sheetName & "' is locked by the register's protection." & vbLf & _
              "Ask the register owner to unlock it."
        GoTo FailClose
    End If

    stepName = "write"
    ' Force TEXT when Excel would otherwise convert it: a leading = + - @
    ' (formula), a number ("1e5") or a date ("1/2"). Normal text is written as is.
    descToWrite = desc
    If InStr(1, "=+-@", Left$(desc, 1)) > 0 Or IsNumeric(desc) Or IsDate(desc) Then descToWrite = "'" & desc
    If kind <> "B" Then
        ' Apostrophe = store as TEXT. Column D is "General": a plain 03 would
        ' become the number 3 and the Part No formula would give M3xxxx.
        ws.Cells(r, COL_PROJ).Value = "'" & proj
        wroteProj = True
    End If
    ws.Cells(r, COL_DESC).Value = descToWrite
    wroteDesc = True

    stepName = "read-back check"
    On Error Resume Next   ' if Calculate is refused, the read-back below fails loudly instead
    ws.Cells(r, COL_PN).Calculate
    On Error GoTo Crash
    pnNow = UCase$(rgText(ws.Cells(r, COL_PN).Value))
    If Len(pnNow) = 0 Then
        why = "Register row " & r & " has no Part No formula in column B (the prepared table ends earlier)." & vbLf & _
              "The cells written were cleared again; nothing was saved." & vbLf & _
              "Ask the register owner to extend the table, or use a lower free number."
        GoTo FailClose
    End If
    If pnNow <> nm Or rgText(ws.Cells(r, COL_DESC).Value) <> desc Then
        why = "Read-back check failed: register row " & r & " shows Part No '" & pnNow & _
              "' and Description '" & rgText(ws.Cells(r, COL_DESC).Value) & "'" & vbLf & _
              "instead of '" & nm & "' / '" & desc & "'." & vbLf & _
              "The cells written were cleared again; nothing was saved. Please report this (tool error)."
        GoTo FailClose
    End If

    stepName = "save register"
    Application.StatusBar = "CreoRegister: saving the register..."
    wb.Save
    savedOk = True

    On Error Resume Next   ' panel refresh is cosmetic; the register is already saved
    rgUpdatePanel wb
    ThisWorkbook.Worksheets(UI_SHEET).Range(REFRESHED_CELL).Value = "'" & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    On Error GoTo Crash

    stepName = "close register"
    If openedByUs Then wb.Close SaveChanges:=False
    Set wb = Nothing
    outcome = "OK"
    why = "Register saved and closed."
    GoTo Finish

AlreadyDone:
    outcome = "ALREADY"
    why = "Row " & r & " already holds exactly this part - nothing written."
    GoTo FailClose

Crash:
    If savedOk Then
        ' The register IS saved; only the tidy-up after it failed. Not a FAIL.
        outcome = "OK"
        why = "Written and saved, but closing the register failed (" & Err.Description & _
              "). Close it by hand if it is still open."
    Else
        why = "Unexpected error at step '" & stepName & "' (tool error, please report): " & _
              Err.Number & " " & Err.Description
    End If
    Resume FailClose

FailClose:
    ' Undo + release on EVERY non-OK path (lessons section 12 rule).
    On Error Resume Next   ' cleanup must run to the end even if one step fails
    If Not savedOk And Not ws Is Nothing Then
        If wroteDesc Then ws.Cells(r, COL_DESC).ClearContents
        If wroteProj Then ws.Cells(r, COL_PROJ).ClearContents
        ' Workbook was clean before we touched it (rgOpenRegister refuses dirty
        ' ones), so after the undo it is clean again.
        If (wroteDesc Or wroteProj) And Not openedByUs Then wb.Saved = True
    End If
    If openedByUs And Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set wb = Nothing
    On Error GoTo 0

Finish:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    rgReport outcome, nm, sheetName, r, proj, desc, why
End Sub

' =============================================================================
' CREO
' =============================================================================

' Reads the active model's name and DESCRIPTION. Always disconnects.
Private Function rgReadCreo(ByRef nm As String, ByRef desc As String, ByRef why As String) As Boolean
    Dim connMaker As pfcls.CCpfcAsyncConnection
    Dim conn As pfcls.IpfcAsyncConnection
    Dim session As pfcls.IpfcBaseSession
    Dim mdl As pfcls.IpfcModel
    Dim paramOwner As pfcls.IpfcParameterOwner
    Dim p As pfcls.IpfcParameter
    Dim bp As pfcls.IpfcBaseParameter
    Dim stage As String

    nm = ""
    desc = ""
    On Error GoTo CreoFail

    stage = "connect"
    Set connMaker = New pfcls.CCpfcAsyncConnection
    Set conn = connMaker.Connect("", "", ".", 20)
    Set session = conn.Session

    stage = "get active model"
    Set mdl = session.CurrentModel
    If mdl Is Nothing Then
        why = "Creo has no active model. Click into the new part's window in Creo, then try again."
        GoTo CleanUp
    End If
    nm = UCase$(Trim$(mdl.InstanceName))

    stage = "read DESCRIPTION"
    Set paramOwner = mdl                        ' same object, parameter-owner handle
    Set p = paramOwner.GetParam(PARAM_DESC)     ' Nothing if the parameter is absent
    If p Is Nothing Then
        why = "Model " & nm & " has no DESCRIPTION parameter (missing)." & vbLf & _
              "Add it in Creo (File > Prepare > Model Properties > Parameters), then try again."
        GoTo CleanUp
    End If
    Set bp = p                                  ' .Value.discr lives on the base interface
    Select Case bp.Value.discr
        Case EpfcParamValueType.EpfcPARAM_STRING:  desc = p.GetScaledValue.StringValue
        Case EpfcParamValueType.EpfcPARAM_INTEGER: desc = CStr(p.GetScaledValue.IntValue)
        Case EpfcParamValueType.EpfcPARAM_DOUBLE:  desc = CStr(p.GetScaledValue.DoubleValue)
        Case EpfcParamValueType.EpfcPARAM_BOOLEAN: desc = CStr(p.GetScaledValue.BoolValue)
    End Select
    desc = Trim$(desc)
    rgReadCreo = True

CleanUp:
    On Error Resume Next   ' a failed disconnect is harmless: we are done with Creo
    If Not conn Is Nothing Then conn.Disconnect 2
    Exit Function

CreoFail:
    If stage = "connect" Then
        why = "Cannot connect to Creo (error): " & Err.Description & vbLf & vbLf & _
              "Check: Creo is running; only ONE Creo session is open; the PRO_COMM_MSG_EXE " & _
              "variable is set; Excel and Creo were restarted after setting it."
    ElseIf stage = "get active model" Then
        why = "Could not get the active model from Creo (error): " & Err.Description & vbLf & _
              "Is a model window open and active in Creo?"
    Else
        why = "Reading from Creo failed at step '" & stage & "' (error): " & Err.Description
    End If
    rgReadCreo = False
    Resume CleanUp
End Function

' =============================================================================
' NAME RULES
' =============================================================================

' M/A: letter + 2-digit project + 4-digit number, nothing else (suffixes refused).
' B:   any name starting with B; matched exactly against column B later.
Private Function rgParseName(ByVal nm As String, ByRef kind As String, ByRef proj As String, _
                             ByRef seq As String, ByRef why As String) As Boolean
    kind = ""
    proj = ""
    seq = ""
    Select Case Left$(nm, 1)
        Case "M", "A"
            If Len(nm) = 7 And rgAllDigits(Mid$(nm, 2)) Then
                kind = Left$(nm, 1)
                proj = Mid$(nm, 2, 2)
                seq = Mid$(nm, 4, 4)
                rgParseName = True
            Else
                why = "Name '" & nm & "' must be exactly " & Left$(nm, 1) & _
                      " + 2-digit project + 4-digit number, e.g. " & Left$(nm, 1) & "281975." & vbLf & _
                      "Suffixes such as _V2 or -01 are not accepted."
            End If
        Case "B"
            If Len(nm) >= 2 Then
                kind = "B"
                rgParseName = True
            Else
                why = "Name '" & nm & "' is not a Bought part number."
            End If
        Case "E"
            why = "Elec parts (E...) are not handled by " & TOOL_VERSION & ". Fill the Elec sheet by hand."
        Case Else
            why = "Name '" & nm & "' does not start with M, A or B." & vbLf & _
                  TOOL_VERSION & " handles Manufactured (M), Assemblies (A) and Bought (B) parts only."
    End Select
End Function

Private Function rgSheetFor(ByVal kind As String) As String
    Select Case kind
        Case "M": rgSheetFor = SH_MFG
        Case "A": rgSheetFor = SH_ASM
        Case "B": rgSheetFor = SH_BUY
    End Select
End Function

' =============================================================================
' REGISTER ACCESS
' =============================================================================

' forReading = True : read-only open (never blocks colleagues)
' forReading = False: write open; refuses if it comes up read-only (in use)
' Re-uses the register if it is already open in this Excel.
Private Function rgOpenRegister(ByVal regPath As String, ByVal forReading As Boolean, _
                                ByRef wb As Workbook, ByRef openedByUs As Boolean, _
                                ByRef why As String) As Boolean
    Dim w As Workbook, fileName As String

    Set wb = Nothing
    openedByUs = False
    If Len(regPath) = 0 Then
        why = "No register file is set. Click 'Pick register file' first."
        Exit Function
    End If
    If StrComp(regPath, ThisWorkbook.FullName, vbTextCompare) = 0 Then
        why = "The register path points at this tool workbook. Click 'Pick register file' and choose the register."
        Exit Function
    End If
    If Not rgFileExists(regPath) Then
        why = "Register file not found (is the shared drive connected?):" & vbLf & regPath
        Exit Function
    End If
    fileName = Mid$(regPath, InStrRev(regPath, "\") + 1)

    For Each w In Application.Workbooks
        If StrComp(w.FullName, regPath, vbTextCompare) = 0 Then
            Set wb = w
            Exit For
        ElseIf StrComp(w.Name, fileName, vbTextCompare) = 0 Then
            why = "A file named '" & fileName & "' is already open in Excel from another path:" & vbLf & _
                  w.FullName & vbLf & "Close it first; the tool must open the register itself."
            Exit Function
        End If
    Next w

    If Not wb Is Nothing Then
        If Not forReading Then
            If wb.ReadOnly Then
                why = "The register is open READ-ONLY in your Excel (someone else probably has it). " & _
                      "Close it, then try again."
                Set wb = Nothing
                Exit Function
            End If
            If Not wb.Saved Then
                why = "The register is open in your Excel with UNSAVED changes." & vbLf & _
                      "Save or close it first, so the tool never saves your own edits by accident."
                Set wb = Nothing
                Exit Function
            End If
        End If
        rgOpenRegister = True
        Exit Function
    End If

    On Error GoTo OpenFail
    ' Notify:=False -> if the file is in use it opens read-only WITHOUT a prompt.
    Set wb = Application.Workbooks.Open(Filename:=regPath, UpdateLinks:=0, ReadOnly:=forReading, _
                                        IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
    On Error GoTo 0
    openedByUs = True

    If Not forReading And wb.ReadOnly Then
        wb.Close SaveChanges:=False
        Set wb = Nothing
        openedByUs = False
        why = "The register is in use by someone else (or you have no write permission), " & _
              "so it opened read-only. Nothing was written." & vbLf & _
              "Try again when they have closed it."
        Exit Function
    End If
    rgOpenRegister = True
    Exit Function

OpenFail:
    why = "Could not open the register (error): " & Err.Description & vbLf & regPath
    Set wb = Nothing
    openedByUs = False
End Function

Private Function rgFileExists(ByVal p As String) As Boolean
    Dim fso As Object
    On Error Resume Next   ' a malformed or unreachable path simply counts as "not found"
    Set fso = CreateObject("Scripting.FileSystemObject")
    rgFileExists = fso.FileExists(p)
End Function

' Sheet lookup ignoring case and stray spaces in tab names.
Private Function rgGetSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim s As Worksheet
    For Each s In wb.Worksheets
        If StrComp(Trim$(s.Name), Trim$(sheetName), vbTextCompare) = 0 Then
            Set rgGetSheet = s
            Exit Function
        End If
    Next s
End Function

' Returns "" when row-2 headers match what this version expects.
Private Function rgHeaderProblem(ByVal ws As Worksheet, ByVal kind As String) As String
    If Not rgHeaderStarts(ws, COL_PN, "Part No") Then rgHeaderProblem = "cell B2 should start with 'Part No'": Exit Function
    If Not rgHeaderStarts(ws, COL_DESC, "Description") Then rgHeaderProblem = "cell C2 should start with 'Description'": Exit Function
    If kind = "B" Then Exit Function
    If Not rgHeaderStarts(ws, COL_PROJ, "Project ID") Then rgHeaderProblem = "cell D2 should start with 'Project ID'": Exit Function
    If Not rgHeaderStarts(ws, COL_SEQ, "Number sequence") Then rgHeaderProblem = "cell E2 should start with 'Number sequence'": Exit Function
End Function

Private Function rgHeaderStarts(ByVal ws As Worksheet, ByVal col As Long, ByVal expected As String) As Boolean
    Dim s As String
    s = rgText(ws.Cells(HEADER_ROW, col).Value)
    rgHeaderStarts = (StrComp(Left$(s, Len(expected)), expected, vbTextCompare) = 0)
End Function

' Project ID must be listed in column A of the Index sheet (row 1 = header).
Private Function rgProjectExists(ByVal wb As Workbook, ByVal proj As String, ByRef why As String) As Boolean
    Dim ix As Worksheet, lastR As Long, i As Long
    Set ix = rgGetSheet(wb, SH_INDEX)
    If ix Is Nothing Then
        why = "The register has no '" & SH_INDEX & "' sheet to check project IDs against."
        Exit Function
    End If
    lastR = ix.Cells(ix.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastR
        If rgPad(ix.Cells(i, 1).Value, 2) = proj Then
            rgProjectExists = True
            Exit Function
        End If
    Next i
    why = "Project ID '" & proj & "' is not listed on the register's Index sheet." & vbLf & _
          "Check the part name, or ask the register owner to add the project."
End Function

' Returns the sheet row whose key column equals key (exactly one match), else 0.
' padLen > 0: key is a zero-padded number (sequence); 0: plain text (Part No).
Private Function rgFindRow(ByVal ws As Worksheet, ByVal keyCol As Long, ByVal key As String, _
                           ByVal padLen As Long, ByRef why As String) As Long
    Dim block As Variant, lastR As Long, i As Long, hits As Long, found As Long
    Dim s As String, lastKey As String

    lastR = ws.Cells(ws.Rows.Count, keyCol).End(xlUp).Row
    If lastR < FIRST_ROW Then
        why = "Sheet '" & ws.Name & "' has no prepared rows."
        Exit Function
    End If
    block = ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(lastR, BLOCK_COLS)).Value   ' always 2-D

    For i = 1 To UBound(block, 1)
        If padLen > 0 Then
            s = rgPad(block(i, keyCol), padLen)
        Else
            s = UCase$(rgText(block(i, keyCol)))
        End If
        If s = key Then
            hits = hits + 1
            If hits = 1 Then found = FIRST_ROW + i - 1
        End If
    Next i

    If hits = 1 Then
        rgFindRow = found
    ElseIf hits = 0 Then
        lastKey = rgText(block(UBound(block, 1), keyCol))
        why = "'" & key & "' was not found on sheet '" & ws.Name & "'." & vbLf & _
              "Prepared rows end at '" & lastKey & "'. Check the part name, or ask the register " & _
              "owner to add more rows."
    Else
        why = "'" & key & "' appears " & hits & " times on sheet '" & ws.Name & _
              "' (first at row " & found & "). Register data problem - fix it by hand first."
    End If
End Function

' Bought rows: lists D..F cells someone already filled ("header: value; ...").
Private Function rgBoughtExtras(ByVal ws As Worksheet, ByVal r As Long) As String
    Dim c As Long, s As String, v As String
    For c = COL_DESC + 1 To COL_BUY_EXTRA_LAST
        v = rgText(ws.Cells(r, c).Value)
        If Len(v) > 0 Then s = s & rgText(ws.Cells(HEADER_ROW, c).Value) & ": " & v & "; "
    Next c
    rgBoughtExtras = s
End Function

Private Function rgCellBlocked(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As Boolean
    rgCellBlocked = ws.ProtectContents And ws.Cells(r, c).Locked
End Function

' =============================================================================
' NEXT-FREE PANEL
' =============================================================================

' Fills panel rows 7..9 from an open register. Returns "" or problem text.
Private Function rgUpdatePanel(ByVal wb As Workbook) As String
    Dim ui As Worksheet, kinds As Variant, k As Long, kind As String
    Dim ws As Worksheet, problems As String, rowNo As Long
    Dim nextNum As String, nextRow As Long, rowsLeft As Long, lastDesc As String, typeAs As String

    Set ui = ThisWorkbook.Worksheets(UI_SHEET)
    kinds = Array("M", "A", "B")
    For k = 0 To 2
        kind = kinds(k)
        rowNo = PANEL_FIRST_ROW + k
        ui.Range(ui.Cells(rowNo, 2), ui.Cells(rowNo, 6)).ClearContents
        Set ws = rgGetSheet(wb, rgSheetFor(kind))
        If ws Is Nothing Then
            ui.Cells(rowNo, 2).Value = "(sheet missing)"
            problems = problems & "Sheet '" & rgSheetFor(kind) & "' not found in the register." & vbLf
        ElseIf Len(rgHeaderProblem(ws, kind)) > 0 Then
            ui.Cells(rowNo, 2).Value = "(layout changed)"
            problems = problems & "Sheet '" & ws.Name & "': " & rgHeaderProblem(ws, kind) & vbLf
        ElseIf rgNextFree(ws, kind, nextNum, nextRow, rowsLeft, lastDesc) Then
            If kind = "B" Then typeAs = nextNum Else typeAs = kind & "##" & nextNum
            If nextNum = "FULL" Then typeAs = "(no prepared rows left)"
            ui.Cells(rowNo, 2).Value = "'" & nextNum
            ui.Cells(rowNo, 3).Value = "'" & typeAs
            If nextRow > 0 Then ui.Cells(rowNo, 4).Value = nextRow
            ui.Cells(rowNo, 5).Value = rowsLeft
            ui.Cells(rowNo, 6).Value = "'" & lastDesc
        Else
            ui.Cells(rowNo, 2).Value = "(no rows)"
        End If
    Next k
    rgUpdatePanel = problems
End Function

' Next free = the row right after the LAST used row (holes in the middle are
' not offered). Used: M/A = C or D filled; B = any of C..F filled.
Private Function rgNextFree(ByVal ws As Worksheet, ByVal kind As String, ByRef nextNum As String, _
                            ByRef nextRow As Long, ByRef rowsLeft As Long, ByRef lastDesc As String) As Boolean
    Dim block As Variant, keyCol As Long, lastR As Long, i As Long, lastUsed As Long

    nextNum = ""
    nextRow = 0
    rowsLeft = 0
    lastDesc = ""
    If kind = "B" Then keyCol = COL_PN Else keyCol = COL_SEQ
    lastR = ws.Cells(ws.Rows.Count, keyCol).End(xlUp).Row
    If lastR < FIRST_ROW Then Exit Function
    block = ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(lastR, BLOCK_COLS)).Value   ' always 2-D

    For i = 1 To UBound(block, 1)
        If rgRowUsed(block, i, kind) Then lastUsed = i
    Next i
    If lastUsed > 0 Then lastDesc = rgText(block(lastUsed, COL_DESC))

    If lastUsed = UBound(block, 1) Then
        nextNum = "FULL"
    Else
        nextRow = FIRST_ROW + lastUsed            ' block index lastUsed+1
        ' Count only rows that really have a Part No (Assemblies has numbered
        ' rows without the Part No formula at the bottom).
        For i = lastUsed + 1 To UBound(block, 1)
            If Len(rgText(block(i, COL_PN))) > 0 Then rowsLeft = rowsLeft + 1
        Next i
        If kind = "B" Then
            nextNum = UCase$(rgText(block(lastUsed + 1, COL_PN)))
        Else
            nextNum = rgPad(block(lastUsed + 1, COL_SEQ), 4)
        End If
    End If
    rgNextFree = True
End Function

Private Function rgRowUsed(ByRef block As Variant, ByVal i As Long, ByVal kind As String) As Boolean
    Dim c As Long
    If kind = "B" Then
        For c = COL_DESC To COL_BUY_EXTRA_LAST
            If Len(rgText(block(i, c))) > 0 Then rgRowUsed = True: Exit Function
        Next c
    Else
        rgRowUsed = (Len(rgText(block(i, COL_DESC))) > 0 Or Len(rgText(block(i, COL_PROJ))) > 0)
    End If
End Function

' =============================================================================
' SMALL HELPERS
' =============================================================================

' Cell value as trimmed text. Empty -> "". Error cells count as filled, so a
' broken row is never treated as free.
Private Function rgText(ByVal v As Variant) As String
    If IsError(v) Then
        rgText = "(cell error)"
    ElseIf IsEmpty(v) Or IsNull(v) Then
        rgText = ""
    Else
        rgText = Trim$(CStr(v))
    End If
End Function

' Digits-only text padded with leading zeros (3 -> "03", 53 -> "0053").
' Anything else is returned unchanged (trimmed).
Private Function rgPad(ByVal v As Variant, ByVal padLen As Long) As String
    Dim s As String
    s = rgText(v)
    If Len(s) > 0 And Len(s) < padLen And rgAllDigits(s) Then s = String$(padLen - Len(s), "0") & s
    rgPad = s
End Function

Private Function rgAllDigits(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    rgAllDigits = True
End Function

Private Function rgRegisterPath() As String
    On Error Resume Next   ' missing Register sheet -> "" -> caller says "run setup / pick file"
    rgRegisterPath = rgText(ThisWorkbook.Worksheets(UI_SHEET).Range(PATH_CELL).Value)
End Function

Private Function rgEnsureSheet(ByVal sheetName As String) As Worksheet
    Dim s As Worksheet
    Set s = rgGetSheet(ThisWorkbook, sheetName)
    If s Is Nothing Then
        Set s = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        s.Name = sheetName
    End If
    Set rgEnsureSheet = s
End Function

Private Sub rgAddButton(ByVal ws As Worksheet, ByVal anchor As Range, ByVal dx As Double, _
                        ByVal btnWidth As Double, ByVal btnCaption As String, ByVal macroName As String)
    Dim b As Object   ' form-control Button
    Set b = ws.Buttons.Add(anchor.Left + dx, anchor.Top + 3, btnWidth, anchor.Height - 6)
    b.OnAction = macroName
    b.Caption = btnCaption
End Sub

' Writes one Log row, the "Last result" cell, saves the tool workbook (so the
' log survives) and shows the popup. Logging problems never hide the popup.
Private Sub rgReport(ByVal outcome As String, ByVal nm As String, ByVal sheetName As String, _
                     ByVal r As Long, ByVal proj As String, ByVal desc As String, ByVal detail As String)
    Dim lg As Worksheet, nr As Long, icon As VbMsgBoxStyle, title As String

    On Error Resume Next   ' a logging failure must not turn a finished write into an error
    Set lg = rgGetSheet(ThisWorkbook, LOG_SHEET)
    If Not lg Is Nothing Then
        nr = lg.Cells(lg.Rows.Count, 1).End(xlUp).Row + 1
        lg.Cells(nr, 1).Value = Now
        lg.Cells(nr, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
        lg.Cells(nr, 2).Value = outcome
        lg.Cells(nr, 3).Value = "'" & nm
        lg.Cells(nr, 4).Value = sheetName
        If r > 0 Then lg.Cells(nr, 5).Value = r
        lg.Cells(nr, 6).Value = "'" & proj
        lg.Cells(nr, 7).Value = "'" & desc
        lg.Cells(nr, 8).Value = "'" & Replace(detail, vbLf, " | ")
        lg.Cells(nr, 9).Value = "'" & Application.UserName
    End If
    ThisWorkbook.Worksheets(UI_SHEET).Range(RESULT_CELL).Value = "'" & outcome & " - " & nm & " - " & Replace(detail, vbLf, " ")
    ' Only a tool workbook already saved to disk (as .xlsm) is saved; a brand-new
    ' unsaved workbook would otherwise pop a Save As dialog.
    If Not ThisWorkbook.ReadOnly And Len(ThisWorkbook.Path) > 0 Then ThisWorkbook.Save
    On Error GoTo 0

    If outcome = "CANCELLED" Then Exit Sub    ' user already knows
    If outcome = "OK" Then
        MsgBox "OK - registered" & vbLf & vbLf & _
               nm & "  ->  '" & sheetName & "' row " & r & vbLf & _
               IIf(Len(proj) > 0, "Project: " & proj & vbLf, "") & _
               "Description: " & desc & vbLf & vbLf & detail, _
               vbInformation, TOOL_VERSION
        Exit Sub
    End If
    If outcome = "ALREADY" Then
        icon = vbInformation
        title = "Already registered"
    Else
        icon = vbExclamation
        title = "NOT registered"
    End If
    MsgBox title & IIf(Len(nm) > 0, " - " & nm, "") & vbLf & vbLf & detail, icon, TOOL_VERSION
End Sub
