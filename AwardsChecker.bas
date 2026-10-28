Option Explicit

'==============================================================================
' AWARDS CHECKER
' Run the macro "RunAwardsCheck" while the Awards file is the active workbook.
'
' Step 1  Cleanse: numbers sitting in column T (PAY_ADDITIONS) where column R
'         (PAY_EMPLOYEERATE) is blank are moved into column R.
' Step 2  Builds a unique list of Job Levels (column D & column I) and runs the
'         pay item validations on every row.
'
' Output sheets (re-created each run):
'   Cleanse Log        - every value moved from T to R
'   Job Levels         - unique job level list with error counts
'   Validation Errors  - one line per failed rule
' Failed rate cells on the Awards sheet are shaded red; moved cells yellow.
'==============================================================================

' Column positions on the Awards sheet (A = 1)
Private Const COL_AWARD As Long = 2      ' B  AWC_DESCRIPTION
Private Const COL_JOB As Long = 4        ' D  AJB_DESCRIPTION
Private Const COL_LEVEL As Long = 9      ' I  ALV_DESCRIPTION
Private Const COL_PAYCODE As Long = 11   ' K  ADT_PAYITEMCODE
Private Const COL_EMP As Long = 18       ' R  PAY_EMPLOYEERATE
Private Const COL_CLIENT As Long = 19    ' S  PAY_CLIENTRATE
Private Const COL_ADD As Long = 20       ' T  PAY_ADDITIONS
Private Const LAST_COL As Long = 21      ' U  PAY_DEDUCTIONS

' Two rates within half a cent are treated as equal (allows rounding to the cent,
' e.g. 2.5 x 16.91 = 42.275 stored as 42.28)
Private Const TOL As Double = 0.0051

Private Const CLR_ERR As Long = 13551615     ' light red
Private Const CLR_MOVED As Long = 10284031   ' light yellow
Private Const CLR_HEAD As Long = 7949855     ' dark blue
Private Const CLR_OK As Long = 13561798      ' light green

Private Const ERR_COLS As Long = 11

Private mWs As Worksheet
Private mData As Variant
Private mErrs() As Variant
Private mErrCount As Long

Public Sub RunAwardsCheck()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim calcMode As Long

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "Open the Awards file first, then run this macro.", vbExclamation
        Exit Sub
    End If

    Set ws = FindAwardsSheet(wb)
    If ws Is Nothing Then
        MsgBox "Could not find a worksheet named 'Awards' in " & wb.Name & ".", vbExclamation
        Exit Sub
    End If

    If UCase$(SafeStr(ws.Cells(1, COL_EMP).Value)) <> "PAY_EMPLOYEERATE" _
       Or UCase$(SafeStr(ws.Cells(1, COL_CLIENT).Value)) <> "PAY_CLIENTRATE" _
       Or UCase$(SafeStr(ws.Cells(1, COL_ADD).Value)) <> "PAY_ADDITIONS" Then
        If MsgBox("Columns R/S/T on '" & ws.Name & "' don't have the expected headers " & _
                  "(PAY_EMPLOYEERATE, PAY_CLIENTRATE, PAY_ADDITIONS)." & vbCrLf & _
                  "Continue anyway?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    End If

    Dim lastRow As Long
    With ws.UsedRange
        lastRow = .Row + .Rows.Count - 1
    End With
    If lastRow < 2 Then
        MsgBox "No data rows found on '" & ws.Name & "'.", vbExclamation
        Exit Sub
    End If

    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo CleanFail

    Set mWs = ws
    mData = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, LAST_COL)).Value

    ' Clear shading from a previous run
    ws.Range(ws.Cells(2, COL_EMP), ws.Cells(lastRow, COL_CLIENT)).Interior.ColorIndex = xlNone

    Dim r As Long

    '--------------------------------------------------------------------------
    ' STEP 1: cleanse - move numbers from T to R
    '--------------------------------------------------------------------------
    Dim logArr() As Variant, logCount As Long, movedCount As Long
    ReDim logArr(1 To lastRow, 1 To 6)

    For r = 2 To lastRow
        If IsNum(mData(r, COL_ADD)) Then
            logCount = logCount + 1
            logArr(logCount, 1) = r
            logArr(logCount, 2) = JobKey(r)
            logArr(logCount, 3) = PayCode(r)
            logArr(logCount, 4) = ToNum(mData(r, COL_ADD))
            If IsBlank(mData(r, COL_EMP)) Then
                mData(r, COL_EMP) = ToNum(mData(r, COL_ADD))
                mData(r, COL_ADD) = Empty
                ws.Cells(r, COL_EMP).Value = mData(r, COL_EMP)
                ws.Cells(r, COL_ADD).ClearContents
                ws.Cells(r, COL_EMP).Interior.Color = CLR_MOVED
                logArr(logCount, 5) = mData(r, COL_EMP)
                logArr(logCount, 6) = "Moved from T to R"
                movedCount = movedCount + 1
            Else
                logArr(logCount, 5) = ToNum(mData(r, COL_EMP))
                logArr(logCount, 6) = "NOT moved - column R already has a value, please review"
            End If
        End If
    Next r

    '--------------------------------------------------------------------------
    ' STEP 2a: unique job levels + NT employee rate per job level / award
    '--------------------------------------------------------------------------
    Dim dLevel As Object, dNT As Object, dNTSeen As Object
    Set dLevel = CreateObject("Scripting.Dictionary")
    Set dNT = CreateObject("Scripting.Dictionary")
    Set dNTSeen = CreateObject("Scripting.Dictionary")

    ' Job Levels columns: 1 Job Level, 2 AJB_DESCRIPTION, 3 ALV_DESCRIPTION, 4 Award(s),
    ' 5 No. of Awards, 6 Pay Items, 7 NT Employee Rate, 8 Errors, 9 Status
    Dim lv() As Variant, nLv As Long
    ReDim lv(1 To lastRow, 1 To 9)

    Dim key As String, award As String, code As String, ntKey As String
    Dim idx As Long

    For r = 2 To lastRow
        If RowHasData(r) Then
            key = JobKey(r)
            award = SafeStr(mData(r, COL_AWARD))
            If Not dLevel.Exists(key) Then
                nLv = nLv + 1
                dLevel.Add key, nLv
                lv(nLv, 1) = key
                lv(nLv, 2) = SafeStr(mData(r, COL_JOB))
                lv(nLv, 3) = SafeStr(mData(r, COL_LEVEL))
                lv(nLv, 4) = award
                lv(nLv, 5) = 1
                lv(nLv, 6) = 0
                lv(nLv, 7) = "(no NT row)"
                lv(nLv, 8) = 0
            Else
                idx = dLevel(key)
                If InStr(1, "; " & lv(idx, 4) & ";", "; " & award & ";", vbBinaryCompare) = 0 Then
                    lv(idx, 4) = lv(idx, 4) & "; " & award
                    lv(idx, 5) = lv(idx, 5) + 1
                End If
            End If
            idx = dLevel(key)
            lv(idx, 6) = lv(idx, 6) + 1

            If PayCode(r) = "NT" Then
                ntKey = key & "|" & award
                If Not dNT.Exists(ntKey) Then
                    dNT.Add ntKey, ToNum(mData(r, COL_EMP))
                    If VarType(lv(idx, 7)) = vbString Then lv(idx, 7) = ToNum(mData(r, COL_EMP))
                End If
            End If
        End If
    Next r

    '--------------------------------------------------------------------------
    ' STEP 2b: validations
    '--------------------------------------------------------------------------
    ReDim mErrs(1 To lastRow * 4, 1 To ERR_COLS)
    mErrCount = 0

    Dim emp As Double, client As Double, nt As Double, expected As Double
    Dim hasNT As Boolean, needZeroClient As Boolean
    Dim empRule As Long, errsBefore As Long
    Dim ruleText As String

    For r = 2 To lastRow
        If RowHasData(r) Then
            key = JobKey(r)
            award = SafeStr(mData(r, COL_AWARD))
            code = PayCode(r)
            emp = ToNum(mData(r, COL_EMP))
            client = ToNum(mData(r, COL_CLIENT))
            ntKey = key & "|" & award
            hasNT = dNT.Exists(ntKey)
            If hasNT Then nt = dNT(ntKey) Else nt = 0
            errsBefore = mErrCount

            If code = "NT" Then
                If dNTSeen.Exists(ntKey) Then
                    AddError r, "More than one NT row for this job level", nt, "", COL_EMP
                Else
                    dNTSeen.Add ntKey, True
                End If
            End If

            ' empRule: 0 = none, 1 = equals NT, 2 = negative NT, 3 = 2.5 x NT
            needZeroClient = False
            empRule = 0
            Select Case code
                Case "AL", "ALCASH", "AL5", "PL"
                    needZeroClient = True: empRule = 1
                Case "WELLB", "COMLEA"
                    empRule = 1
                Case "PHNW"
                    needZeroClient = True: empRule = 3
                Case Else
                    If MatchesNumPlusNum(code, "RDOACC") Then
                        needZeroClient = True: empRule = 2
                    ElseIf MatchesNumPlusNum(code, "RDO") Then
                        needZeroClient = True: empRule = 1
                    End If
            End Select

            If needZeroClient And client <> 0 Then
                AddError r, code & ": client rate should be 0", IIf(hasNT, nt, ""), 0, COL_CLIENT
            End If

            If empRule <> 0 Then
                If Not hasNT Then
                    AddError r, code & ": no NT row found for this job level, cannot check employee rate", "", "", COL_EMP
                Else
                    Select Case empRule
                        Case 1
                            expected = nt
                            ruleText = code & ": employee rate should equal NT employee rate"
                        Case 2
                            expected = -nt
                            ruleText = code & ": employee rate should be negative NT employee rate"
                        Case 3
                            expected = nt * 2.5
                            ruleText = code & ": employee rate should be 2.5 x NT employee rate"
                    End Select
                    If Abs(emp - expected) > TOL Then
                        AddError r, ruleText, nt, expected, COL_EMP
                    End If
                End If
            End If

            If client <> 0 And Not (client > emp) Then
                AddError r, code & ": client rate should be greater than employee rate", IIf(hasNT, nt, ""), "> " & Format$(emp, "0.0000"), COL_CLIENT
            End If

            idx = dLevel(key)
            lv(idx, 8) = lv(idx, 8) + (mErrCount - errsBefore)
        End If
    Next r

    For idx = 1 To nLv
        If lv(idx, 8) = 0 Then lv(idx, 9) = "OK" Else lv(idx, 9) = "Errors"
    Next idx

    '--------------------------------------------------------------------------
    ' OUTPUT
    '--------------------------------------------------------------------------
    Dim shLog As Worksheet, shLv As Worksheet, shErr As Worksheet

    Set shLog = PrepSheet(wb, "Cleanse Log", ws)
    WriteHeader shLog, Array("Awards Row", "Job Level", "Pay Code", "Value in T", "Employee Rate (R) now", "Action")
    If logCount > 0 Then
        shLog.Range("A2").Resize(logCount, 6).Value = logArr
        shLog.Range("D2").Resize(logCount, 2).NumberFormat = "0.0000"
    Else
        shLog.Range("A2").Value = "No numbers found in column T."
    End If
    FinishSheet shLog, 6

    Set shLv = PrepSheet(wb, "Job Levels", ws)
    WriteHeader shLv, Array("Job Level (D & I)", "AJB_DESCRIPTION", "ALV_DESCRIPTION", "Award(s)", _
                            "No. of Awards", "Pay Items", "NT Employee Rate", "Errors", "Status")
    If nLv > 0 Then
        shLv.Range("A2").Resize(nLv, 9).Value = lv
        shLv.Range("G2").Resize(nLv, 1).NumberFormat = "0.0000"
        shLv.Range("A1").Resize(nLv + 1, 9).Sort Key1:=shLv.Range("A2"), Order1:=xlAscending, Header:=xlYes
        For idx = 2 To nLv + 1
            If shLv.Cells(idx, 9).Value = "OK" Then
                shLv.Cells(idx, 9).Interior.Color = CLR_OK
            Else
                shLv.Cells(idx, 9).Interior.Color = CLR_ERR
            End If
        Next idx
    End If
    FinishSheet shLv, 9

    Set shErr = PrepSheet(wb, "Validation Errors", ws)
    WriteHeader shErr, Array("Awards Row", "Job Level (D & I)", "Award", "AJB_DESCRIPTION", "ALV_DESCRIPTION", _
                             "Pay Code", "Issue", "Employee Rate (R)", "Client Rate (S)", "NT Employee Rate", "Expected")
    If mErrCount > 0 Then
        shErr.Range("A2").Resize(mErrCount, ERR_COLS).Value = mErrs
        shErr.Range("H2").Resize(mErrCount, 4).NumberFormat = "0.0000"
    Else
        shErr.Range("A2").Value = "No validation errors found."
    End If
    FinishSheet shErr, ERR_COLS

    Application.Calculation = calcMode
    Application.ScreenUpdating = True

    shErr.Activate
    ActiveWindow.FreezePanes = False
    shErr.Range("A2").Select
    ActiveWindow.FreezePanes = True

    MsgBox "Awards check complete." & vbCrLf & vbCrLf & _
           "Data rows checked:  " & (lastRow - 1) & vbCrLf & _
           "Values moved T -> R:  " & movedCount & vbCrLf & _
           "Unique job levels:  " & nLv & vbCrLf & _
           "Validation errors:  " & mErrCount, vbInformation, "Awards Checker"

    Erase mErrs
    mData = Empty
    Set mWs = Nothing
    Exit Sub

CleanFail:
    Application.Calculation = calcMode
    Application.ScreenUpdating = True
    MsgBox "The check stopped with an error:" & vbCrLf & Err.Description, vbCritical, "Awards Checker"
End Sub

'------------------------------------------------------------------------------
' Helpers
'------------------------------------------------------------------------------

Private Sub AddError(ByVal r As Long, ByVal issue As String, ByVal ntRate As Variant, _
                     ByVal expected As Variant, ByVal highlightCol As Long)
    mErrCount = mErrCount + 1
    mErrs(mErrCount, 1) = r
    mErrs(mErrCount, 2) = JobKey(r)
    mErrs(mErrCount, 3) = SafeStr(mData(r, COL_AWARD))
    mErrs(mErrCount, 4) = SafeStr(mData(r, COL_JOB))
    mErrs(mErrCount, 5) = SafeStr(mData(r, COL_LEVEL))
    mErrs(mErrCount, 6) = PayCode(r)
    mErrs(mErrCount, 7) = issue
    If IsBlank(mData(r, COL_EMP)) Then mErrs(mErrCount, 8) = "" Else mErrs(mErrCount, 8) = ToNum(mData(r, COL_EMP))
    If IsBlank(mData(r, COL_CLIENT)) Then mErrs(mErrCount, 9) = "" Else mErrs(mErrCount, 9) = ToNum(mData(r, COL_CLIENT))
    mErrs(mErrCount, 10) = ntRate
    mErrs(mErrCount, 11) = expected
    mWs.Cells(r, highlightCol).Interior.Color = CLR_ERR
End Sub

Private Function FindAwardsSheet(ByVal wb As Workbook) As Worksheet
    Dim sh As Worksheet
    For Each sh In wb.Worksheets
        If LCase$(sh.Name) = "awards" Then Set FindAwardsSheet = sh: Exit Function
    Next sh
    ' A CSV opened in Excel names the sheet after the file, e.g. "Awards (2)"
    For Each sh In wb.Worksheets
        If LCase$(sh.Name) Like "awards*" Then Set FindAwardsSheet = sh: Exit Function
    Next sh
End Function

Private Function PrepSheet(ByVal wb As Workbook, ByVal sheetName As String, ByVal dataSheet As Worksheet) As Worksheet
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = wb.Worksheets(sheetName)
    On Error GoTo 0
    If sh Is Nothing Then
        Set sh = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        sh.Name = sheetName
    Else
        If sh.AutoFilterMode Then sh.AutoFilterMode = False
        sh.Cells.Clear
    End If
    Set PrepSheet = sh
End Function

Private Sub WriteHeader(ByVal sh As Worksheet, ByVal headers As Variant)
    Dim n As Long
    n = UBound(headers) - LBound(headers) + 1
    With sh.Range("A1").Resize(1, n)
        .Value = headers
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = CLR_HEAD
    End With
End Sub

Private Sub FinishSheet(ByVal sh As Worksheet, ByVal nCols As Long)
    sh.Cells.Font.Name = "Arial"
    sh.Cells.Font.Size = 10
    sh.Range("A1").Resize(1, nCols).AutoFilter
    sh.Columns(1).Resize(, nCols).AutoFit
End Sub

Private Function JobKey(ByVal r As Long) As String
    JobKey = SafeStr(mData(r, COL_JOB)) & SafeStr(mData(r, COL_LEVEL))
End Function

Private Function PayCode(ByVal r As Long) As String
    PayCode = UCase$(Trim$(SafeStr(mData(r, COL_PAYCODE))))
End Function

Private Function RowHasData(ByVal r As Long) As Boolean
    RowHasData = Len(SafeStr(mData(r, COL_JOB))) + Len(SafeStr(mData(r, COL_LEVEL))) _
                 + Len(SafeStr(mData(r, COL_PAYCODE))) > 0
End Function

' True when code = prefix & <number> & "+" & <number>, e.g. RDO38+2, RDOACC36+4
Private Function MatchesNumPlusNum(ByVal code As String, ByVal prefix As String) As Boolean
    Dim rest As String, parts() As String
    If Len(code) <= Len(prefix) Then Exit Function
    If Left$(code, Len(prefix)) <> prefix Then Exit Function
    rest = Mid$(code, Len(prefix) + 1)
    parts = Split(rest, "+")
    If UBound(parts) <> 1 Then Exit Function
    MatchesNumPlusNum = IsPlainNumber(parts(0)) And IsPlainNumber(parts(1))
End Function

' Digits with at most one decimal point, e.g. 38, 2, 2.5
Private Function IsPlainNumber(ByVal s As String) As Boolean
    Dim i As Long, ch As String, dots As Long, digits As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = "." Then
            dots = dots + 1
            If dots > 1 Then Exit Function
        ElseIf ch >= "0" And ch <= "9" Then
            digits = digits + 1
        Else
            Exit Function
        End If
    Next i
    IsPlainNumber = digits > 0
End Function

Private Function SafeStr(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    SafeStr = CStr(v)
End Function

Private Function IsBlank(ByVal v As Variant) As Boolean
    IsBlank = (Len(Trim$(SafeStr(v))) = 0)
End Function

Private Function IsNum(ByVal v As Variant) As Boolean
    If IsBlank(v) Or IsError(v) Then Exit Function
    IsNum = IsNumeric(v)
End Function

Private Function ToNum(ByVal v As Variant) As Double
    If IsBlank(v) Or IsError(v) Then Exit Function
    If IsNumeric(v) Then ToNum = CDbl(v)
End Function
