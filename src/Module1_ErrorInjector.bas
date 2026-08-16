Attribute VB_Name = "Module1"
Option Explicit

'==================================================================
'  ERROR INJECTOR
'  Deliberately corrupts ~3% of rows and logs what it did to an
'  "AnswerKey" sheet, so the validation macro can be tested
'  against a known truth.
'==================================================================
Sub InjectErrors()

    Dim ws As Worksheet, wsKey As Worksheet
    Dim lastRow As Long, nBreak As Long
    Dim i As Long, r As Long, srcR As Long, errType As Long
    Dim keyRow As Long

    Set ws = ThisWorkbook.Worksheets("InForce")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "No data on the InForce sheet.", vbExclamation
        Exit Sub
    End If

    ' 3% of the data rows, minimum 1
    nBreak = Application.WorksheetFunction.Max(1, Int((lastRow - 1) * 0.03))

    Set wsKey = GetOrCreateSheet("AnswerKey")
    wsKey.Cells.Clear
    wsKey.Range("A1:C1").Value = Array("Row", "Field broken", "What was done")
    wsKey.Range("A1:C1").Font.Bold = True
    keyRow = 2

    Randomize

    For i = 1 To nBreak

        r = Int(Rnd() * (lastRow - 1)) + 2      ' a random data row
        errType = Int(Rnd() * 6) + 1            ' one of six error types

        Select Case errType

            Case 1  ' duplicate policy number
                srcR = Int(Rnd() * (lastRow - 1)) + 2
                If srcR = r Then srcR = IIf(r = 2, 3, r - 1)
                ws.Cells(r, 1).Value = ws.Cells(srcR, 1).Value
                LogKey wsKey, keyRow, r, "PolicyNumber", _
                       "Overwrote with the policy number from row " & srcR

            Case 2  ' implausible issue age
                If Rnd() < 0.5 Then
                    ws.Cells(r, 4).Value = 0
                    LogKey wsKey, keyRow, r, "IssueAge", "Set issue age to 0"
                Else
                    ws.Cells(r, 4).Value = 150
                    LogKey wsKey, keyRow, r, "IssueAge", "Set issue age to 150"
                End If

            Case 3  ' non-positive face amount
                If Rnd() < 0.5 Then
                    ws.Cells(r, 6).Value = 0
                    LogKey wsKey, keyRow, r, "FaceAmount", "Set face amount to 0"
                Else
                    ws.Cells(r, 6).Value = -Abs(ws.Cells(r, 6).Value)
                    LogKey wsKey, keyRow, r, "FaceAmount", "Made face amount negative"
                End If

            Case 4  ' issue date after the valuation date
                ws.Cells(r, 5).Value = DateSerial(2025, 6, 15)
                LogKey wsKey, keyRow, r, "IssueDate", _
                       "Set issue date to 2025-06-15, after the valuation date"

            Case 5  ' missing required field
                If Rnd() < 0.5 Then
                    ws.Cells(r, 3).ClearContents
                    LogKey wsKey, keyRow, r, "Sex", "Blanked the sex code"
                Else
                    ws.Cells(r, 2).ClearContents
                    LogKey wsKey, keyRow, r, "ProductCode", "Blanked the product code"
                End If

            Case 6  ' product code not on the approved list
                ws.Cells(r, 2).Value = "XX99"
                LogKey wsKey, keyRow, r, "ProductCode", _
                       "Replaced product code with unapproved value XX99"

        End Select

    Next i

    wsKey.Columns("A:C").AutoFit

    MsgBox "Injected " & nBreak & " errors. See the AnswerKey sheet.", vbInformation

End Sub


'--- writes one line to the answer key and advances the row counter
Private Sub LogKey(wsKey As Worksheet, ByRef keyRow As Long, _
                   r As Long, fieldName As String, note As String)
    wsKey.Cells(keyRow, 1).Value = r
    wsKey.Cells(keyRow, 2).Value = fieldName
    wsKey.Cells(keyRow, 3).Value = note
    keyRow = keyRow + 1
End Sub


'--- returns a sheet by name, creating it if it doesn't exist yet
Private Function GetOrCreateSheet(nm As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
                 After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nm
    End If
    Set GetOrCreateSheet = ws
End Function

