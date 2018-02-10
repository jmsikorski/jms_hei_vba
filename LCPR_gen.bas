Attribute VB_Name = "LCPR"
Private Function open_file() As String
    On Error GoTo 10
    Dim strFileToOpen As Variant
    Dim f As Variant
    Dim wb As Workbook
1:
    strFileToOpen = Application.GetOpenFilename _
    (Title:="Please select file to import", _
    FileFilter:="Excel Files *.xls* (*.xls*),", _
    MultiSelect:=False)
    If Not IsArray(strFileToOpen) Then
        If strFileToOpen = False Then
            MsgBox "No file selected.", vbExclamation, "Sorry!"
            GoTo 1
        Else
            Set wb = Workbooks.Open(Filename:=strFileToOpen)
            open_file = wb.Name
            Exit Function
        End If
    Else
        For Each f In strFileToOpen
            Dim tmp() As String
            Set wb = Workbooks.Open(Filename:=strFileToOpen)
            open_file = wb.Name
            Exit Function
        Next f
    End If
    open_file = "ERROR"
    Exit Function
10:
    MsgBox "ERROR: Import file can not be open!", vbCritical
    ThisWorkbook.Close False
    open_file = "ERROR"
    On Error GoTo 0
    Exit Function
test:
    On Error Resume Next
    Workbooks.Open "C:\Users\jsikorski\Desktop\461705 LCPR.xls"
    open_file = "461705 LCPR.xls"
End Function

Public Sub new_report()
Attribute new_report.VB_ProcData.VB_Invoke_Func = "N\n14"
    On Error GoTo e_msg
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim ogwb As Workbook
    Dim wb As Workbook
    Dim rng As Range
    Dim ws As Worksheet
    Dim rep_wb As Workbook
    Dim j_name As String
    Dim j_num As String
    Set ogwb = ThisWorkbook
    Set rep_wb = Workbooks.Add
    ogwb.Worksheets("REPORT").Copy before:=rep_wb.Worksheets(1)
add_job:
    Set wb = Workbooks(open_file)
    Set ws = wb.Worksheets(1)
    Set rng = ws.Range("A1")
    j_num = ws.Range("B1")
    j_name = ws.Range("C1")
    On Error Resume Next
    rep_wb.Worksheets("Sheet1").Delete
    On Error GoTo e_msg
    rep_wb.Worksheets(1).Name = "REPORT"
    Do While rng <> "Description"
        Set rng = rng.Offset(0, 1)
    Loop
    Set rng = ws.Range(ws.Range("A1"), rng.End(xlDown))
    rng.EntireColumn.Delete
    Set rng = ws.Range("U1")
    Do While Left(rng, 5) <> "Total"
        Set rng = rng.Offset(1, 0)
    Loop
    Set rng = ws.Range(rng, rng.End(xlToRight))
    rng.Copy
    ws.Range("A1").End(xlDown).Offset(2, 0).PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    ws.Range(rng, rng.End(xlToRight)).EntireColumn.Delete
    Dim r_rng As Range
    Set r_rng = rep_wb.Worksheets("REPORT").Range("rep_start")
    Set rng = ws.Range("A1")
    Do While rng.Cells(1, 1) <> vbNullString
        r_rng = j_num
        r_rng.Offset(0, 1) = j_name
        For i = 0 To 13
            r_rng.Offset(0, i + 2) = rng.Offset(0, i)
        Next i
        Set r_rng = r_rng.Offset(1, 0)
        r_rng.EntireRow.Copy
        r_rng.EntireRow.Insert
        Set r_rng = r_rng.Offset(-1, 0)
        r_rng.Offset(1, 0).PasteSpecial xlPasteFormats
        Set rng = rng.Offset(1, 0)
        
    Loop
    Application.CutCopyMode = False
    With rep_wb.Worksheets("REPORT")
        Set r_rng = .Range(.Range("rep_start").End(xlDown).Offset(1, 0), .Range("rep_start").End(xlDown).Offset(2, 0))
        r_rng.EntireRow.Delete
        Set r_rng = .Range("rep_start").Offset(0, 17)
        For i = 0 To 5
            r_rng.Offset(0, i * 2).Copy
            .Range(r_rng.Offset(0, i * 2), r_rng.Offset(0, i * 2).End(xlDown)).PasteSpecial xlPasteFormulas
        Next i
    End With
    ws.Parent.Close False
    rep_wb.SaveAs "LCPR SPREADSHEET_CMiC_" & Format(Now(), "MM.DD.YY") & ".xlsx"
    Dim ans As Integer
    ans = MsgBox("Would you like to import another job?", vbYesNo, "Add Job")
    If ans = vbYes Then
        ogwb.Worksheets("REPORT").Range("new_job").Copy
        Set rng = rep_wb.Worksheets(1).Range("rep_start").End(xlDown).Offset(7, 0)
        rng.PasteSpecial xlPasteAll
        rep_wb.Names.Item("rep_start").RefersTo = "=" & rng.Address
        GoTo add_job
    Else
        rep_wb.Worksheets("REPORT").Range("A1").Activate
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        ogwb.Close True
    End If
e_msg:
    Dim e_ans As Integer
    e_ans = MsgBox("Error importing file", vbAbortRetryIgnore + vbCritical, "ERROR")
    If e_ans = vbAbort Then
        ThisWorkbook.Close False
    ElseIf e_ans = vbRetry Then
        GoTo add_job
    ElseIf e_ans = vbIgnore Then
        Exit Sub
    Else
        GoTo e_msg
    End If
End Sub

Public Sub res_update()
    Application.ScreenUpdating = True
End Sub
