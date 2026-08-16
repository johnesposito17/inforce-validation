Attribute VB_Name = "Module2"
' In-Force Data Validation Utility
' Author: John Esposito
' Purpose: Applies six data-quality checks to a life in-force extract
'          and produces a severity-ranked exception report.
' Valuation date: 2024-12-31
' Last updated: 2026-08-16

Option Explicit

'==================================================================
'  IN-FORCE DATA VALIDATION
'
'  Scans the in-force extract, applies six data-quality checks,
'  and produces a severity-ranked exception report plus a summary.
'==================================================================

'--- Configuration: change these, not the logic below -------------
Private Const DATA_SHEET     As String = "InForce"
Private Const REPORT_SHEET   As String = "Exceptions"
Private Const SUMMARY_SHEET  As String = "Summary"
Private Const VALID_PRODUCTS As String = "TRM10,TRM20,TRM30,WL,UL"
Private Const MIN_AGE        As Long = 18
Private Const MAX_AGE        As Long = 85
'------------------------------------------------------------------

Private exRows As Collection    ' holds every exception found


Sub RunDataChecks()

    Dim wsData As Worksheet
    Dim data As Variant
    Dim lastRow As Long, r As Long, srcRow As Long
    Dim valDate As Date
    Dim seen As Collection
    Dim firstRow As Long
    Dim polNum As String, prod As String, sexCode As String
    Dim ageVal As Variant, faceVal As Variant, issDate As Variant

    Application.ScreenUpdating = False

    Set wsData = ThisWorkbook.Worksheets(DATA_SHEET)
    valDate = DateSerial(2024, 12, 31)          ' the valuation date

    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No data found on " & DATA_SHEET, vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Pull the whole block into memory in one go. Reading cell by cell
    ' from the sheet would be roughly a thousand times slower.
    data = wsData.Range("A2:G" & lastRow).Value

    Set exRows = New Collection
    Set seen = New Collection      ' remembers policy numbers already seen

    For r = 1 To UBound(data, 1)

        srcRow = r + 1                          ' array row 1 = sheet row 2

        polNum = Trim(CStr(data(r, 1)))
        prod = Trim(CStr(data(r, 2)))
        sexCode = Trim(CStr(data(r, 3)))
        ageVal = data(r, 4)
        issDate = data(r, 5)
        faceVal = data(r, 6)

        '--- CHECK 1: policy number present and unique --------------
        If Len(polNum) = 0 Then
            AddException "High", "(blank)", srcRow, "Missing policy number", _
                "PolicyNumber", "", _
                "Policy number is blank, so this record cannot be identified " & _
                "or matched back to the prior valuation."
        Else
            ' Look the policy number up in the "seen" collection. Asking
            ' for a key that isn't there raises an error, so we swallow
            ' it and leave firstRow at 0, meaning "not seen before".
            firstRow = 0
            On Error Resume Next
            firstRow = seen(polNum)
            On Error GoTo 0

            If firstRow > 0 Then
                AddException "High", polNum, srcRow, "Duplicate policy number", _
                    "PolicyNumber", polNum, _
                    "Policy number " & polNum & " already appears on row " & _
                    firstRow & ". Duplicates double-count reserves and " & _
                    "overstate the company's liability."
            Else
                ' Collection.Add takes the ITEM first, then the KEY --
                ' the opposite order to most languages. Easy to get wrong.
                seen.Add srcRow, polNum
            End If
        End If

        '--- CHECK 2: issue age within plausible bounds -------------
        If Trim(CStr(ageVal)) = "" Or Not IsNumeric(ageVal) Then
            AddException "Medium", polNum, srcRow, "Invalid issue age", _
                "IssueAge", CStr(ageVal), _
                "Issue age is missing or non-numeric, so no mortality rate " & _
                "can be looked up for this life."

        ElseIf CDbl(ageVal) < MIN_AGE Or CDbl(ageVal) > MAX_AGE Then
            AddException "Medium", polNum, srcRow, "Issue age out of range", _
                "IssueAge", CStr(ageVal), _
                "Issue age of " & ageVal & " falls outside the plausible " & _
                "issue range of " & MIN_AGE & " to " & MAX_AGE & ". Likely a " & _
                "keying error or a date-of-birth conversion fault."
        End If

        '--- CHECK 3: face amount is positive -----------------------
        If Trim(CStr(faceVal)) = "" Or Not IsNumeric(faceVal) Then
            AddException "High", polNum, srcRow, "Invalid face amount", _
                "FaceAmount", CStr(faceVal), _
                "Face amount is missing or non-numeric, so no death benefit " & _
                "or reserve can be calculated for this policy."

        ElseIf CDbl(faceVal) <= 0 Then
            AddException "High", polNum, srcRow, "Non-positive face amount", _
                "FaceAmount", CStr(faceVal), _
                "Face amount of " & Format(CDbl(faceVal), "#,##0") & " is zero " & _
                "or negative. An in-force policy must carry a positive death " & _
                "benefit; a negative value would produce a negative reserve."
        End If

        '--- CHECK 4: issue date not after the valuation date -------
        If Trim(CStr(issDate)) = "" Or Not IsDate(issDate) Then
            AddException "High", polNum, srcRow, "Invalid issue date", _
                "IssueDate", CStr(issDate), _
                "Issue date is missing or unreadable, so policy duration and " & _
                "the applicable reserve factors cannot be determined."

        ElseIf CDate(issDate) > valDate Then
            AddException "High", polNum, srcRow, "Issue date after valuation date", _
                "IssueDate", Format(CDate(issDate), "yyyy-mm-dd"), _
                "Policy is dated " & Format(CDate(issDate), "yyyy-mm-dd") & _
                ", which is after the valuation date of " & _
                Format(valDate, "yyyy-mm-dd") & ". It should not appear in " & _
                "this extract at all."
        End If

        '--- CHECK 5: sex code present and recognized ---------------
        If sexCode = "" Then
            AddException "Medium", polNum, srcRow, "Missing sex code", _
                "Sex", "", _
                "Sex code is blank, so sex-distinct mortality rates cannot be " & _
                "applied to this life."

        ElseIf UCase(sexCode) <> "M" And UCase(sexCode) <> "F" Then
            AddException "Medium", polNum, srcRow, "Unrecognized sex code", _
                "Sex", sexCode, _
                "Sex code '" & sexCode & "' is not one of the expected values " & _
                "M or F."
        End If

        '--- CHECK 6: product code on the approved list -------------
        If prod = "" Then
            AddException "High", polNum, srcRow, "Missing product code", _
                "ProductCode", "", _
                "Product code is blank, so the policy cannot be assigned to a " & _
                "valuation basis or reported in the correct product line."

        ElseIf InStr(1, "," & VALID_PRODUCTS & ",", "," & UCase(prod) & ",", _
                     vbTextCompare) = 0 Then
            AddException "High", polNum, srcRow, "Unknown product code", _
                "ProductCode", prod, _
                "Product code '" & prod & "' is not on the approved list (" & _
                VALID_PRODUCTS & "). Either the extract is corrupted or a new " & _
                "product has not been set up in the valuation system."
        End If

    Next r

    WriteReport
    WriteSummary

    Application.ScreenUpdating = True

    MsgBox "Validation complete." & vbCrLf & vbCrLf & _
           Format(lastRow - 1, "#,##0") & " policies scanned." & vbCrLf & _
           exRows.Count & " exceptions raised.", vbInformation, "Data Validation"

End Sub


'--- records one exception, tagging it with a sort rank -----------
Private Sub AddException(sev As String, polNum As String, srcRow As Long, _
                         checkName As String, fieldName As String, _
                         valueFound As String, descr As String)
    Dim rank As Long
    Select Case sev
        Case "High":   rank = 1
        Case "Medium": rank = 2
        Case Else:     rank = 3
    End Select
    exRows.Add Array(rank, sev, polNum, srcRow, checkName, fieldName, _
                     valueFound, descr)
End Sub


'--- writes the ranked exception report ---------------------------
Private Sub WriteReport()

    Dim wsRep As Worksheet
    Dim out() As Variant
    Dim item As Variant
    Dim i As Long, j As Long, n As Long

    Set wsRep = GetOrCreateSheet2(REPORT_SHEET)
    wsRep.Cells.Clear

    wsRep.Range("A1:H1").Value = Array("Rank", "Severity", "Policy Number", _
        "Source Row", "Check", "Field", "Value Found", "Description")

    n = exRows.Count
    If n = 0 Then
        wsRep.Range("A2").Value = "No exceptions found."
        Exit Sub
    End If

    ReDim out(1 To n, 1 To 8)
    For i = 1 To n
        item = exRows(i)
        For j = 0 To 7
            out(i, j + 1) = item(j)
        Next j
    Next i

    wsRep.Range("A2").Resize(n, 8).Value = out

    ' sort by severity rank, then by source row
    With wsRep.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsRep.Range("A2:A" & n + 1), Order:=xlAscending
        .SortFields.Add Key:=wsRep.Range("D2:D" & n + 1), Order:=xlAscending
        .SetRange wsRep.Range("A1:H" & n + 1)
        .Header = xlYes
        .Apply
    End With

    wsRep.Columns("A").Delete          ' drop the helper rank column

    ' formatting
    With wsRep.Range("A1:G1")
        .Font.Bold = True
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
    End With

    For i = 2 To n + 1
        Select Case wsRep.Cells(i, 1).Value
            Case "High":   wsRep.Cells(i, 1).Interior.Color = RGB(255, 199, 206)
            Case "Medium": wsRep.Cells(i, 1).Interior.Color = RGB(255, 235, 156)
            Case "Low":    wsRep.Cells(i, 1).Interior.Color = RGB(198, 239, 206)
        End Select
    Next i

    wsRep.Columns("A:F").AutoFit
    wsRep.Columns("G").ColumnWidth = 90
    wsRep.Range("G:G").WrapText = True
    wsRep.Rows(1).AutoFilter
    wsRep.Activate
    wsRep.Range("A2").Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True

End Sub


'--- writes the management summary --------------------------------
Private Sub WriteSummary()

    Dim wsSum As Worksheet
    Dim item As Variant, checkNames As Variant
    Dim i As Long, r As Long, c As Long, tally As Long
    Dim nHigh As Long, nMed As Long, nLow As Long

    ' Every check name the macro can raise. Listing them all means the
    ' summary shows a zero for checks that found nothing, which proves
    ' the check ran rather than leaving the reader to wonder.
    checkNames = Array( _
        "Missing policy number", "Duplicate policy number", _
        "Invalid issue age", "Issue age out of range", _
        "Invalid face amount", "Non-positive face amount", _
        "Invalid issue date", "Issue date after valuation date", _
        "Missing sex code", "Unrecognized sex code", _
        "Missing product code", "Unknown product code")

    For i = 1 To exRows.Count
        item = exRows(i)
        Select Case item(1)                    ' item(1) = severity
            Case "High":   nHigh = nHigh + 1
            Case "Medium": nMed = nMed + 1
            Case Else:     nLow = nLow + 1
        End Select
    Next i

    Set wsSum = GetOrCreateSheet2(SUMMARY_SHEET)
    wsSum.Cells.Clear

    wsSum.Range("A1").Value = "In-Force Data Validation - Summary"
    wsSum.Range("A1").Font.Bold = True
    wsSum.Range("A1").Font.Size = 14

    wsSum.Range("A2").Value = "Run at:"
    wsSum.Range("B2").Value = Now
    wsSum.Range("B2").NumberFormat = "yyyy-mm-dd hh:mm"
    wsSum.Range("A3").Value = "Total exceptions:"
    wsSum.Range("B3").Value = exRows.Count

    r = 5
    wsSum.Cells(r, 1).Value = "By severity"
    wsSum.Cells(r, 1).Font.Bold = True
    wsSum.Cells(r + 1, 1).Value = "High"
    wsSum.Cells(r + 1, 2).Value = nHigh
    wsSum.Cells(r + 2, 1).Value = "Medium"
    wsSum.Cells(r + 2, 2).Value = nMed
    wsSum.Cells(r + 3, 1).Value = "Low"
    wsSum.Cells(r + 3, 2).Value = nLow

    r = r + 5
    wsSum.Cells(r, 1).Value = "By check"
    wsSum.Cells(r, 1).Font.Bold = True
    r = r + 1

    For c = LBound(checkNames) To UBound(checkNames)
        tally = 0
        For i = 1 To exRows.Count
            item = exRows(i)
            If item(4) = checkNames(c) Then tally = tally + 1   ' item(4) = check name
        Next i
        wsSum.Cells(r, 1).Value = checkNames(c)
        wsSum.Cells(r, 2).Value = tally
        r = r + 1
    Next c

    wsSum.Columns("A:B").AutoFit

End Sub


'--- same helper as in Module1, kept local so the modules are independent
Private Function GetOrCreateSheet2(nm As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
                 After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nm
    End If
    Set GetOrCreateSheet2 = ws
End Function
