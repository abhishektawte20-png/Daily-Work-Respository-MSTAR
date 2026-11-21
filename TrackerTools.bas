Attribute VB_Name = "TrackerTools"
Option Explicit

' ============================================================================
'  Daily Update Tracker - button macros
'  Every macro asks simple questions (InputBox) and writes to Roster / Teams /
'  Weightages / Base_Targets / Daily_Log. Nothing is deleted.
' ============================================================================

Private Const ROST_FIRST As Long = 7
Private Const ROST_LAST As Long = 406
Private Const LOG_FIRST As Long = 7
Private Const LOG_LASTCOL As Long = 45

Private Function Sh(ByVal n As String) As Worksheet
    Set Sh = ThisWorkbook.Worksheets(n)
End Function

Private Function FirstEmptyRow(ByVal w As Worksheet, ByVal col As Long, ByVal r1 As Long, ByVal r2 As Long) As Long
    Dim r As Long
    For r = r1 To r2
        If Len(Trim$(CStr(w.Cells(r, col).Value))) = 0 Then
            FirstEmptyRow = r
            Exit Function
        End If
    Next r
    FirstEmptyRow = 0
End Function

Private Function ChooseFrom(ByVal ttl As String, ByVal rng As Range) As String
    Dim c As Range, msg As String, items() As String, n As Long, ans As String, i As Long
    For Each c In rng.Cells
        If Len(Trim$(CStr(c.Value))) > 0 Then
            n = n + 1
            ReDim Preserve items(1 To n)
            items(n) = Trim$(CStr(c.Value))
            msg = msg & n & ") " & items(n) & vbLf
        End If
    Next c
    If n = 0 Then
        MsgBox "The list is empty - add entries first.", vbExclamation
        Exit Function
    End If
    ans = Trim$(InputBox(msg & vbLf & "Type the number (or the exact name):", ttl))
    If ans = "" Then Exit Function
    If IsNumeric(ans) Then
        i = CLng(ans)
        If i >= 1 And i <= n Then
            ChooseFrom = items(i)
            Exit Function
        End If
    End If
    For i = 1 To n
        If StrComp(items(i), ans, vbTextCompare) = 0 Then
            ChooseFrom = items(i)
            Exit Function
        End If
    Next i
    MsgBox "'" & ans & "' is not in the list.", vbExclamation
End Function

Private Function AskDate(ByVal prompt As String, ByVal ttl As String, ByVal dflt As String, ByRef ok As Boolean) As Date
    Dim s As String
    ok = False
    s = Trim$(InputBox(prompt, ttl, dflt))
    If s = "" Then Exit Function
    If Not IsDate(s) Then
        MsgBox "'" & s & "' is not a valid date. Please use a format like 15-Mar-2027.", vbExclamation
        Exit Function
    End If
    AskDate = CDate(s)
    ok = True
End Function

Private Function PickEmployeeRow(ByVal ttl As String) As Long
    Dim s As String, r As Long, w As Worksheet, hits(1 To 20) As Long, n As Long
    Dim msg As String, ans As String, i As Long
    Set w = Sh("Roster")
    s = Trim$(InputBox("Type any part of the employee's name:", ttl))
    If s = "" Then Exit Function
    For r = ROST_FIRST To ROST_LAST
        If Len(CStr(w.Cells(r, 2).Value)) > 0 And Len(CStr(w.Cells(r, 6).Value)) = 0 Then
            If InStr(1, CStr(w.Cells(r, 2).Value), s, vbTextCompare) > 0 Then
                If n < 20 Then
                    n = n + 1
                    hits(n) = r
                End If
            End If
        End If
    Next r
    If n = 0 Then
        MsgBox "No current employee matches '" & s & "'.", vbExclamation
        Exit Function
    End If
    If n = 1 Then
        PickEmployeeRow = hits(1)
        Exit Function
    End If
    For i = 1 To n
        msg = msg & i & ") " & w.Cells(hits(i), 2).Value & "   (" & w.Cells(hits(i), 3).Value & ")" & vbLf
    Next i
    ans = Trim$(InputBox(msg & vbLf & "Type the number:", ttl))
    If IsNumeric(ans) Then
        i = CLng(ans)
        If i >= 1 And i <= n Then PickEmployeeRow = hits(i)
    End If
End Function

Private Function AppendLogRows(ByVal nm As String, ByVal d1 As Date, ByVal d2 As Date) As Long
    Dim wl As Worksheet, n As Long, firstR As Long, lastF As Long, arr() As Variant
    Dim i As Long, c As Long, oldCalc As Long
    Set wl = Sh("Daily_Log")
    n = CLng(d2 - d1) + 1
    If n < 1 Then Exit Function
    firstR = wl.Cells(wl.Rows.Count, 2).End(xlUp).Row + 1
    If firstR < LOG_FIRST Then firstR = LOG_FIRST
    ReDim arr(1 To n, 1 To 2)
    For i = 1 To n
        arr(i, 1) = d1 + (i - 1)
        arr(i, 2) = nm
    Next i
    Application.ScreenUpdating = False
    oldCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    lastF = wl.Cells(wl.Rows.Count, 4).End(xlUp).Row
    If firstR + n - 1 > lastF Then
        For c = 3 To LOG_LASTCOL
            If wl.Cells(lastF, c).HasFormula Then
                wl.Range(wl.Cells(lastF, c), wl.Cells(firstR + n - 1, c)).FillDown
            End If
        Next c
    End If
    wl.Range(wl.Cells(firstR, 1), wl.Cells(firstR + n - 1, 2)).Value = arr
    wl.Range(wl.Cells(firstR, 1), wl.Cells(firstR + n - 1, 1)).NumberFormat = "dd-mmm-yy"
    Application.Calculation = oldCalc
    Application.ScreenUpdating = True
    AppendLogRows = n
End Function

' ---------------------------------------------------------------- JOINER
Private Sub AppendRampRow(ByVal nm As String, ByVal d As Date, ByVal stg As String)
    Dim w As Worksheet, r As Long
    Set w = Sh("Ramping")
    r = FirstEmptyRow(w, 8, 7, 206)
    If r = 0 Then
        MsgBox "The ramp-progress table is full (200 rows).", vbExclamation
        Exit Sub
    End If
    w.Cells(r, 8).Value = nm
    w.Cells(r, 9).Value = d
    w.Cells(r, 9).NumberFormat = "dd-mmm-yyyy"
    w.Cells(r, 10).Value = stg
End Sub

Private Function FirstStage() As String
    Dim c As Range
    For Each c In Sh("Ramping").Range("F7:F16").Cells
        If Len(Trim$(CStr(c.Value))) > 0 Then
            FirstStage = Trim$(CStr(c.Value))
            Exit Function
        End If
    Next c
End Function

' returns number of daily rows created, or -1 if the person could not be added
Private Function CreateJoiner(ByVal nm As String, ByVal id As String, ByVal lead As String, ByVal grp As String, _
                              ByVal startD As Date, ByVal ramp As Boolean, ByVal batch As String) As Long
    Dim wr As Worksheet, r As Long, yr As Long
    Set wr = Sh("Roster")
    CreateJoiner = -1
    If Application.WorksheetFunction.CountIf(wr.Range("B" & ROST_FIRST & ":B" & ROST_LAST), nm) > 0 Then Exit Function
    yr = CLng(Sh("Settings").Range("B3").Value)
    If Year(startD) > yr Then Exit Function
    r = FirstEmptyRow(wr, 2, ROST_FIRST, ROST_LAST)
    If r = 0 Then Exit Function
    wr.Cells(r, 1).Value = id
    wr.Cells(r, 2).Value = nm
    wr.Cells(r, 3).Value = lead
    wr.Cells(r, 4).Value = batch
    wr.Cells(r, 5).Value = startD
    wr.Cells(r, 8).Value = grp
    wr.Cells(r, 10).Value = startD
    If ramp Then AppendRampRow nm, startD, FirstStage()
    CreateJoiner = AppendLogRows(nm, startD, DateSerial(yr, 12, 31))
End Function

Public Sub AddJoiner()
    Dim nm As String, id As String, lead As String, grp As String, ok As Boolean
    Dim startD As Date, ramp As Boolean, n As Long, ans As String
    nm = Trim$(InputBox("Full name of the new joiner:", "Add joiner (1/5)"))
    If nm = "" Then Exit Sub
    If Application.WorksheetFunction.CountIf(Sh("Roster").Range("B" & ROST_FIRST & ":B" & ROST_LAST), nm) > 0 Then
        MsgBox "'" & nm & "' is already on the Roster. To change their team use 'Move Team'.", vbExclamation
        Exit Sub
    End If
    id = Trim$(InputBox("Employee ID (can be left blank):", "Add joiner (2/5)"))
    lead = ChooseFrom("Add joiner (3/5) - team lead", Sh("Teams").Range("A7:A26"))
    If lead = "" Then Exit Sub
    grp = ChooseFrom("Add joiner (4/5) - target group (used after ramp-up)", Sh("Base_Targets").Range("A7:A16"))
    If grp = "" Then Exit Sub
    startD = AskDate("First working day (dd-mmm-yyyy):", "Add joiner (5/5)", Format(Date, "dd-mmm-yyyy"), ok)
    If Not ok Then Exit Sub
    ans = MsgBox("Is " & nm & " going through the RAMPING programme (week-by-week targets)?", vbYesNo + vbQuestion, "Ramping?")
    ramp = (ans = vbYes)
    n = CreateJoiner(nm, id, lead, grp, startD, ramp, "")
    If n < 0 Then
        MsgBox "Could not add " & nm & " (start date after the tracker year, or Roster full).", vbExclamation
        Exit Sub
    End If
    MsgBox nm & " added to " & lead & "." & vbLf & n & " daily rows created in Daily_Log (weekends included, shaded)." & _
           IIf(ramp, vbLf & "Started in " & FirstStage() & " on the Ramping tab.", ""), vbInformation, "Done"
End Sub

Public Sub AddRampBatch()
    Dim names As String, parts() As String, i As Long, nm As String, lead As String, grp As String
    Dim ok As Boolean, startD As Date, batch As String, n As Long, added As Long, skipped As String
    names = Trim$(InputBox("Type ALL the names in the new batch, separated by a semicolon ( ; )" & vbLf & _
            "Example: Asha Rao; Ravi Patil; Meena Shah", "Add ramp batch (1/4)"))
    If names = "" Then Exit Sub
    lead = ChooseFrom("Add ramp batch (2/4) - team lead", Sh("Teams").Range("A7:A26"))
    If lead = "" Then Exit Sub
    grp = ChooseFrom("Add ramp batch (3/4) - target group after ramp-up", Sh("Base_Targets").Range("A7:A16"))
    If grp = "" Then Exit Sub
    startD = AskDate("Batch start date (dd-mmm-yyyy):", "Add ramp batch (4/4)", Format(Date, "dd-mmm-yyyy"), ok)
    If Not ok Then Exit Sub
    batch = Format(startD, "mmm-yy") & " batch"
    parts = Split(names, ";")
    For i = LBound(parts) To UBound(parts)
        nm = Trim$(parts(i))
        If Len(nm) > 0 Then
            n = CreateJoiner(nm, "", lead, grp, startD, True, batch)
            If n >= 0 Then
                added = added + 1
            Else
                skipped = skipped & nm & "; "
            End If
        End If
    Next i
    MsgBox added & " researcher(s) added to " & lead & " and started in " & FirstStage() & "." & _
           IIf(Len(skipped) > 0, vbLf & "NOT added (already on Roster / invalid): " & skipped, ""), vbInformation, "Done"
End Sub

' ---------------------------------------------------------------- RAMP STAGE (manual decision)
Public Sub AdvanceRampStage()
    Dim r As Long, w As Worksheet, cur As String, nxt As String, c As Range, found As Boolean
    Dim ok As Boolean, d As Date, wp As Worksheet
    r = PickEmployeeRow("Advance ramp stage")
    If r = 0 Then Exit Sub
    Set w = Sh("Roster")
    Set wp = Sh("Ramping")
    cur = Trim$(CStr(w.Cells(r, 7).Value))
    If cur = "" Or StrComp(cur, "Graduated", vbTextCompare) = 0 Then
        MsgBox w.Cells(r, 2).Value & " is not currently in the ramping programme.", vbExclamation
        Exit Sub
    End If
    For Each c In wp.Range("F7:F16").Cells
        If found And Len(Trim$(CStr(c.Value))) > 0 Then
            nxt = Trim$(CStr(c.Value))
            Exit For
        End If
        If StrComp(Trim$(CStr(c.Value)), cur, vbTextCompare) = 0 Then found = True
    Next c
    If nxt = "" Then nxt = "Graduated"
    d = AskDate(w.Cells(r, 2).Value & ": move from " & cur & " to " & nxt & " starting (dd-mmm-yyyy):", "Advance ramp stage", Format(Date, "dd-mmm-yyyy"), ok)
    If Not ok Then Exit Sub
    AppendRampRow CStr(w.Cells(r, 2).Value), d, nxt
    MsgBox w.Cells(r, 2).Value & " is now " & nxt & " from " & Format(d, "dd-mmm-yyyy") & ".", vbInformation
End Sub

Public Sub GraduateRamp()
    Dim r As Long, w As Worksheet, ok As Boolean, d As Date
    r = PickEmployeeRow("Graduate from ramping")
    If r = 0 Then Exit Sub
    Set w = Sh("Roster")
    d = AskDate("Graduation date for " & w.Cells(r, 2).Value & " (dd-mmm-yyyy):", "Graduate", Format(Date, "dd-mmm-yyyy"), ok)
    If Not ok Then Exit Sub
    AppendRampRow CStr(w.Cells(r, 2).Value), d, "Graduated"
    MsgBox w.Cells(r, 2).Value & " graduates on " & Format(d, "dd-mmm-yyyy") & " and gets the normal target from then on.", vbInformation
End Sub

Public Sub NewRampTarget()
    Dim w As Worksheet, stg As String, d As Date, ok As Boolean, v As String, r As Long
    Set w = Sh("Ramping")
    stg = ChooseFrom("Ramp target - which stage?", w.Range("F7:F16"))
    If stg = "" Then Exit Sub
    d = AskDate("New target for '" & stg & "' applies FROM which date (dd-mmm-yyyy)?", "Ramp target", Format(Date, "dd-mmm-yyyy"), ok)
    If Not ok Then Exit Sub
    v = Trim$(InputBox("Target PPD for '" & stg & "':", "Ramp target"))
    If Not IsNumeric(v) Or v = "" Then
        MsgBox "Please enter a number.", vbExclamation
        Exit Sub
    End If
    If CDbl(v) <= 0 Then
        MsgBox "The target must be greater than 0.", vbExclamation
        Exit Sub
    End If
    r = FirstEmptyRow(w, 1, 7, 46)
    If r = 0 Then
        MsgBox "The stage-target table is full.", vbExclamation
        Exit Sub
    End If
    w.Cells(r, 1).Value = d
    w.Cells(r, 1).NumberFormat = "dd-mmm-yyyy"
    w.Cells(r, 2).Value = stg
    w.Cells(r, 3).Value = CDbl(v)
    w.Cells(r, 4).Value = "Added via button on " & Format(Date, "dd-mmm-yyyy")
    MsgBox "'" & stg & "' target becomes " & v & " from " & Format(d, "dd-mmm-yyyy") & ".", vbInformation
End Sub

' ---------------------------------------------------------------- LEAVER
Public Sub EndEmployee()
    Dim r As Long, ok As Boolean, d As Date, w As Worksheet
    r = PickEmployeeRow("Remove leaver")
    If r = 0 Then Exit Sub
    Set w = Sh("Roster")
    d = AskDate("Last working day for " & w.Cells(r, 2).Value & " (dd-mmm-yyyy):", "Remove leaver", Format(Date, "dd-mmm-yyyy"), ok)
    If Not ok Then Exit Sub
    If Len(CStr(w.Cells(r, 5).Value)) > 0 Then
        If d < CDate(w.Cells(r, 5).Value) Then
            MsgBox "That is before their start date.", vbExclamation
            Exit Sub
        End If
    End If
    If MsgBox("Set the last day of " & w.Cells(r, 2).Value & " to " & Format(d, "dd-mmm-yyyy") & "?" & vbLf & _
              "Rows after this date drop out of all totals.", vbYesNo + vbQuestion, "Confirm") <> vbYes Then Exit Sub
    w.Cells(r, 6).Value = d
    w.Cells(r, 6).NumberFormat = "dd-mmm-yyyy"
    MsgBox "Done. " & w.Cells(r, 2).Value & " is excluded after " & Format(d, "dd-mmm-yyyy") & ".", vbInformation
End Sub

' ---------------------------------------------------------------- TEAM MOVE
Public Sub MoveTeam()
    Dim r As Long, r2 As Long, w As Worksheet, newLead As String, ok As Boolean, eff As Date, c As Long
    r = PickEmployeeRow("Move to another team")
    If r = 0 Then Exit Sub
    Set w = Sh("Roster")
    newLead = ChooseFrom("Move team - NEW team lead for " & w.Cells(r, 2).Value, Sh("Teams").Range("A7:A26"))
    If newLead = "" Then Exit Sub
    If StrComp(newLead, CStr(w.Cells(r, 3).Value), vbTextCompare) = 0 Then
        MsgBox "They are already in that team.", vbExclamation
        Exit Sub
    End If
    eff = AskDate("First day in the new team (dd-mmm-yyyy):", "Move team", Format(Date, "dd-mmm-yyyy"), ok)
    If Not ok Then Exit Sub
    If Len(CStr(w.Cells(r, 5).Value)) > 0 Then
        If eff <= CDate(w.Cells(r, 5).Value) Then
            MsgBox "The move date must be after their current start date.", vbExclamation
            Exit Sub
        End If
    End If
    r2 = FirstEmptyRow(w, 2, ROST_FIRST, ROST_LAST)
    If r2 = 0 Then
        MsgBox "The Roster is full (400 rows).", vbExclamation
        Exit Sub
    End If
    For c = 1 To 11
        If c <> 7 Then w.Cells(r2, c).Value = w.Cells(r, c).Value
    Next c
    w.Cells(r2, 3).Value = newLead
    w.Cells(r2, 5).Value = eff
    w.Cells(r2, 6).ClearContents
    w.Cells(r, 6).Value = eff - 1
    w.Cells(r, 6).NumberFormat = "dd-mmm-yyyy"
    MsgBox w.Cells(r, 2).Value & " moves to " & newLead & " from " & Format(eff, "dd-mmm-yyyy") & "." & vbLf & _
           "History before that date stays with the old team.", vbInformation
End Sub

' ---------------------------------------------------------------- TEAM LEAD
Public Sub AddTeamLead()
    Dim nm As String, w As Worksheet, r As Long
    Set w = Sh("Teams")
    nm = Trim$(InputBox("Name of the new team lead:", "Add team lead"))
    If nm = "" Then Exit Sub
    If Application.WorksheetFunction.CountIf(w.Range("A7:A26"), nm) > 0 Then
        MsgBox "'" & nm & "' is already a team lead.", vbExclamation
        Exit Sub
    End If
    r = FirstEmptyRow(w, 1, 7, 26)
    If r = 0 Then
        MsgBox "The Teams list is full (20).", vbExclamation
        Exit Sub
    End If
    w.Cells(r, 1).Value = nm
    w.Cells(r, 2).Value = "All Rounders"
    w.Cells(r, 3).Value = "Y"
    MsgBox nm & " added. They now appear in every dropdown, Team_View and Dashboard." & vbLf & _
           "Next: use 'Add Joiner' or 'Move Team' to put people in their team.", vbInformation
End Sub

' ---------------------------------------------------------------- WEIGHTAGES
Public Sub NewWeightVersion()
    Dim w As Worksheet, lastR As Long, r As Long, d As Date, ok As Boolean, c As Long
    Set w = Sh("Weightages")
    lastR = w.Cells(w.Rows.Count, 1).End(xlUp).Row
    If lastR < 6 Then lastR = 6
    If lastR >= 21 Then
        MsgBox "The weightage table is full (15 versions). Clear an old row first.", vbExclamation
        Exit Sub
    End If
    If lastR = 6 Then
        MsgBox "Please type the first weightage row by hand.", vbExclamation
        Exit Sub
    End If
    d = AskDate("New weightages apply FROM which date (dd-mmm-yyyy)?" & vbLf & "(The latest row is dated " & Format(w.Cells(lastR, 1).Value, "dd-mmm-yyyy") & ")", "New weightage version", Format(Date, "dd-mmm-yyyy"), ok)
    If Not ok Then Exit Sub
    If d <= CDate(w.Cells(lastR, 1).Value) Then
        MsgBox "The new date must be later than the latest row.", vbExclamation
        Exit Sub
    End If
    r = lastR + 1
    w.Cells(r, 1).Value = d
    For c = 2 To 17
        w.Cells(r, c).Value = w.Cells(lastR, c).Value
    Next c
    w.Cells(r, 18).Value = "New version - edit the changed weights"
    w.Activate
    w.Cells(r, 2).Select
    MsgBox "A new row dated " & Format(d, "dd-mmm-yyyy") & " was added (copy of the previous weights)." & vbLf & _
           "Now type over the weights that changed. Earlier dates keep their old weights.", vbInformation
End Sub

' ---------------------------------------------------------------- BASE TARGETS
Public Sub NewTargetVersion()
    Dim w As Worksheet, grp As String, d As Date, ok As Boolean, v As String, r As Long, i As Long
    Set w = Sh("Base_Targets")
    grp = ChooseFrom("Base target change - which group?", w.Range("A7:A16"))
    If grp = "" Then Exit Sub
    d = AskDate("New base target applies FROM which date (dd-mmm-yyyy)?", "Base target change", Format(Date, "dd-mmm-yyyy"), ok)
    If Not ok Then Exit Sub
    v = Trim$(InputBox("New base target (PPD) for '" & grp & "':", "Base target change"))
    If Not IsNumeric(v) Or v = "" Then
        MsgBox "Please enter a number.", vbExclamation
        Exit Sub
    End If
    If CDbl(v) <= 0 Then
        MsgBox "The target must be greater than 0.", vbExclamation
        Exit Sub
    End If
    For i = 7 To 46
        If StrComp(CStr(w.Cells(i, 5).Value), grp, vbTextCompare) = 0 And Len(CStr(w.Cells(i, 4).Value)) > 0 Then
            If CDate(w.Cells(i, 4).Value) = d Then
                MsgBox "A target for '" & grp & "' on that exact date already exists (row " & i & "). Edit that row instead.", vbExclamation
                Exit Sub
            End If
        End If
    Next i
    r = FirstEmptyRow(w, 4, 7, 46)
    If r = 0 Then
        MsgBox "The target table is full.", vbExclamation
        Exit Sub
    End If
    w.Cells(r, 4).Value = d
    w.Cells(r, 4).NumberFormat = "dd-mmm-yyyy"
    w.Cells(r, 5).Value = grp
    w.Cells(r, 6).Value = CDbl(v)
    w.Cells(r, 7).Value = "Added via button on " & Format(Date, "dd-mmm-yyyy")
    MsgBox "'" & grp & "' base target becomes " & v & " from " & Format(d, "dd-mmm-yyyy") & ". Earlier dates are unchanged.", vbInformation
End Sub

' ---------------------------------------------------------------- LOG HELPERS
Public Sub ToggleConfCL()
    Dim w As Worksheet
    Set w = Sh("Daily_Log")
    w.Columns("R:Z").Hidden = Not w.Columns("R:Z").Hidden
End Sub

Public Sub SortLogByDate()
    Dim wl As Worksheet, lastR As Long
    Set wl = Sh("Daily_Log")
    lastR = wl.Cells(wl.Rows.Count, 2).End(xlUp).Row
    If lastR <= LOG_FIRST Then Exit Sub
    If MsgBox("Sort Daily_Log by Date, then Employee? (takes a few seconds)", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    Application.ScreenUpdating = False
    wl.Range(wl.Cells(LOG_FIRST, 1), wl.Cells(lastR, LOG_LASTCOL)).Sort _
        Key1:=wl.Cells(LOG_FIRST, 1), Order1:=xlAscending, _
        Key2:=wl.Cells(LOG_FIRST, 2), Order2:=xlAscending, Header:=xlNo
    Application.ScreenUpdating = True
    MsgBox "Sorted.", vbInformation
End Sub
