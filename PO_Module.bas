Attribute VB_Name = "PO_Module"
Const shts = 3

Public Function get_req(Optional num As Integer) As Worksheet
Attribute get_req.VB_ProcData.VB_Invoke_Func = "r\n14"
    Dim ws As Worksheet
    Dim file As String
    Dim tmp() As String
    Dim new_file As String
    Dim wb As Workbook
    Dim xFile As String
    Dim xFodler As String
    Dim job() As String
    job = Split(get_job, " ")
    If num = 0 Then
        xFile = select_req
    Else
        xFile = job(0) & "-" & get_req_number(num) & ".xls*"
    End If
    xFolder = ThisWorkbook.Path & "\FIELD PO REQ\"
    If file_exists(xFolder & xFile) Then
        Set wb = Workbooks.Open(FileName:=xFolder & xFile)
    Else
        Set wb = Workbooks.Open(FileName:=select_req)
    End If
    Set get_req = wb.Worksheets(1)
'    If open_workbook_FileDialog <> 1 Then
'        MsgBox "ERROR GETTING SHEETS", vbCritical, "ERROR"
'    End If
End Function

'Public Sub test()
'    Dim ws As Worksheet
'    Dim po_ws As Worksheet
'    Dim po_wb As Workbook
'    Dim wb As Workbook
'    Set wb = ThisWorkbook
'    Set ws = gen_req
'    Set po_ws = get_req
'    Set po_wb = po_ws.Parent
'    Stop
'    wb.Activate
'    import_req ws, po_ws
'    po_wb.Close savechanges:=False
'End Sub

Private Function select_req() As String
    On Error GoTo 10
    Dim strFileToOpen As Variant
    Dim f As Variant
    Dim wb As Workbook
1:
    
    strFileToOpen = Application.GetOpenFilename _
    (Title:="Please select PO REQ file", _
    FileFilter:="Excel Files *.xls* (*.xls*),", _
    MultiSelect:=False)
    If Not IsArray(strFileToOpen) Then
        If strFileToOpen = False Then
            MsgBox "No file selected.", vbExclamation, "Sorry!"
            GoTo 1
        Else
            Set wb = Workbooks.Open(FileName:=strFileToOpen)
            select_req = wb.Name
            Exit Function
        End If
    Else
        For Each f In strFileToOpen
            Dim tmp() As String
            Set wb = Workbooks.Open(FileName:=strFileToOpen)
            select_req = wb.Name
            Exit Function
'            tmp = Split(f, "\")
'            Set wb = Workbooks(tmp(6))
'            Set ws = wb.Worksheets(1)
'            ws.Copy after:=ThisWorkbook.Worksheets(ThisWorkbook.Sheets.Count)
'            new_file = wb.Path & "\ NEW PO'S\"
'            new_file = new_file & tmp(3) & "_PO# " & ThisWorkbook.Sheets.Count - shts & "_" & Format(ws.Range("H9").Value, "mm.dd.yy") & ".xlsx"
'            wb.SaveAs new_file, xlOpenXMLWorkbook
'            wb.Close
'            ThisWorkbook.Worksheets(ThisWorkbook.Sheets.Count).Name = "REQ " & ThisWorkbook.Sheets.Count - shts
        Next f
    End If
    select_req = "ERROR"
    Exit Function
10:
    select_req = "ERROR"
    On Error GoTo 0
End Function

Private Function get_job() As String
    Dim job() As String
    job = Split(ThisWorkbook.Path, "\")
    For i = 1 To UBound(job)
        If Left(job(i), 2) = "46" Then
            get_job = job(i)
            Exit Function
        End If
    Next i
    get_job = "46XXXX TEST JOB"
End Function


Private Function file_exists(xFile As String) As Boolean
     file_exists = (Dir(xFile) > "")
End Function

'Private Function open_workbook_FileDialog() As Integer
'    On Error GoTo 10
'    Dim strFileToOpen As Variant
'    Dim f As Variant
'1:
'    strFileToOpen = Application.GetOpenFilename _
'    (Title:="Please choose a file to open", _
'    FileFilter:="Excel Files *.xls* (*.xls*),", _
'    MultiSelect:=True)
'    If Not IsArray(strFileToOpen) Then
'        If strFileToOpen = False Then
'            MsgBox "No file selected.", vbExclamation, "Sorry!"
'            GoTo 1
'        Else
'        Workbooks.Open Filename:=strFileToOpen
'        End If
'    Else
'        For Each f In strFileToOpen
'            Dim tmp() As String
'            Workbooks.Open Filename:=f
'            tmp = Split(f, "\")
'            Set wb = Workbooks(tmp(6))
'            Set ws = wb.Worksheets(1)
'            ws.Copy after:=ThisWorkbook.Worksheets(ThisWorkbook.Sheets.Count)
'            new_file = wb.Path & "\ NEW PO'S\"
'            new_file = new_file & tmp(3) & "_PO# " & ThisWorkbook.Sheets.Count - shts & "_" & Format(ws.Range("H9").Value, "mm.dd.yy") & ".xlsx"
'            wb.SaveAs new_file, xlOpenXMLWorkbook
'            wb.Close
'            ThisWorkbook.Worksheets(ThisWorkbook.Sheets.Count).Name = "REQ " & ThisWorkbook.Sheets.Count - shts
'        Next f
'    End If
'    open_workbook_FileDialog = 1
'    Exit Function
'10:
'    open_workbook_FileDialog = -1
'    On Error GoTo 0
'End Function

Private Function get_req_number(num As Integer) As String
    get_req_number = Right("0000" & num, 4)
End Function

Private Function gen_req() As Worksheet
    Dim ws As Worksheet
    Dim rng As Range
    Dim ans As Integer
    
    ThisWorkbook.Worksheets("MASTER").Copy after:=ThisWorkbook.Worksheets(Sheets.Count)
    Set rng = ThisWorkbook.Worksheets("PO LOG").Range("Req_start").Offset(Sheets.Count - shts, 0)
    Set ws = ThisWorkbook.Worksheets(Sheets.Count)
    ws.Unprotect
    If rng.Offset(-1, 0).Value = "REQ #" Then
        ws.Name = "REQ " & get_req_number(Range("Req_No"))
    Else
        ws.Name = "REQ " & get_req_number(rng.Offset(-1, 0).Value + 1)
    End If
    ws.Range("date").Value = Format(Now, "mm/dd/yyyy")
    
    ws.Range("start").Activate
    Set gen_req = ws
End Function

Public Sub new_req()
Attribute new_req.VB_ProcData.VB_Invoke_Func = "n\n14"
    Dim ws As Worksheet
    Dim rng As Range
    Dim ans As Integer
    Dim po_ws As Worksheet
    Dim po_wb As Workbook
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Set rng = ThisWorkbook.Worksheets("PO LOG").Range("Req_start").Offset(ThisWorkbook.Sheets.Count - shts, 0)
         
    ans = MsgBox("Import Field PO Req?", vbYesNoCancel + vbQuestion, "Import?")
    If ans = vbYes Then
        Set ws = gen_req
        If rng = vbNullString Then
            Set po_ws = get_req(1)
        Else
            Set po_ws = get_req(rng.Offset(1, 0).Value)
        End If
        Set po_wb = po_ws.Parent
        wb.Activate
        import_req ws, po_ws
    ElseIf ans = vbNo Then
        Set ws = gen_req
        ws.Range("start").Select
    ElseIf ans = vbCancel Then
        Exit Sub
    End If
End Sub

Private Function import_req(new_ws As Worksheet, import_ws As Worksheet) As Integer
    Application.ScreenUpdating = False
    Dim rng As Range
    Dim used_rng As Range
    Dim i As Integer
    Dim max As Integer
    Dim main_wb As Workbook
    Dim i_wb As Workbook
    Set i_wb = import_ws.Parent
    Dim xFile As String
    xFile = i_wb.Path & "\" & i_wb.Name
    Set main_wb = new_ws.Parent
    i = 0
    max = 23
    Set used_rng = import_ws.Range("po_list")
    With new_ws.Range("start")
    For Each rng In used_rng
        If rng <> vbNullString Then
            .Offset(i, 0) = rng
            .Offset(i, 1) = "16000"
            .Offset(i, 2) = import_ws.Range("po_qty").Offset(i + 1, 0)
            i = i + 1
        Else
            Exit For
        End If
    Next
    End With
    Set rng = main_wb.Worksheets("PO LOG").Range("Req_start").Offset(main_wb.Sheets.Count - shts, 3)
    main_wb.Worksheets("PO LOG").Unprotect
    main_wb.Worksheets("PO LOG").Hyperlinks.Add rng, xFile, , , "FIELD XLS"
    main_wb.Worksheets("PO LOG").Protect
    import_req = 1
    i_wb.Close savechanges:=False
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Function
10:
    MsgBox "Unable to import PO REQ", vbCritical
    import_req = -1
    i_wb.Close False
    new_ws.Delete
    On Error GoTo 0
    Application.ScreenUpdating = True
End Function

'Private Function import_req(new_ws As Worksheet, import_ws As Worksheet) As Integer
'    Dim rng As Range
'    Dim used_rng As Range
'    Stop
'    Set used_rng = import_ws.Range("po_list")
'    For Each rng In used_rng
'        With new_ws.Range("start")
'        If rng <> vbNullString Then
'
'    Next
'End Function

Private Sub save_req()
Attribute save_req.VB_ProcData.VB_Invoke_Func = "S\n14"
    Dim xSht As Worksheet
    Dim xFileDlg As FileDialog
    Dim xFolder As String
    Dim xYesorNo As Integer
    Dim xOutlookObj As Object
    Dim xEmailObj As Object
    Dim xUsedRng As Range
    Dim job() As String
    Dim rng As Range
    Set xSht = ActiveSheet
    Set rng = ThisWorkbook.Worksheets("PO LOG").Range("link_start")
    job = Split(ThisWorkbook.Path, "\")
            
    Set rng = ThisWorkbook.Worksheets("PO LOG").Range("link_start").Offset(xSht.Index - 3, 0)
    
    On Error Resume Next
    xFolder = ThisWorkbook.Path & "\" & ThisWorkbook.Worksheets("PO LOG").Range("pdf") & "\"
    MkDir xFolder
    On Error GoTo 0
    xFolder = xFolder & job(3) & "_PO " & get_req_number(xSht.Index - (shts + 1) + ThisWorkbook.Worksheets("PO LOG").Range("Req_No")) _
    & "_" & Format(Now, "mm.dd.yy") & ".pdf"
    'Check if file already exist
    If Len(Dir(xFolder)) > 0 Then
        xYesorNo = MsgBox(xFolder & " already exists." & vbCrLf & vbCrLf & "Do you want to overwrite it?", _
                          vbYesNo + vbQuestion, "File Exists")
        On Error Resume Next
        If xYesorNo = vbYes Then
            Kill xFolder
        Else
            MsgBox "if you don't overwrite the existing PDF, I can't continue." _
                        & vbCrLf & vbCrLf & "Press OK to exit this macro.", vbCritical, "Exiting Macro"
            Exit Sub
        End If
        If Err.Number <> 0 Then
            MsgBox "Unable to delete existing file.  Please make sure the file is not open or write protected." _
                        & vbCrLf & vbCrLf & "Press OK to exit this macro.", vbCritical, "Unable to Delete File"
            Exit Sub
        End If
    End If
     
    Set xUsedRng = xSht.UsedRange
    If Application.WorksheetFunction.CountA(xUsedRng.Cells) <> 0 Then
        'Save as PDF file
        xSht.ExportAsFixedFormat Type:=xlTypePDF, FileName:=xFolder, Quality:=xlQualityStandard
        ThisWorkbook.Worksheets("PO LOG").Unprotect
        ThisWorkbook.Worksheets("PO LOG").Hyperlinks.Add rng, xFolder, , , "PDF FILE"
        ThisWorkbook.Worksheets("PO LOG").Protect
    Else
      MsgBox "The active worksheet cannot be blank"
      Exit Sub
    End If
End Sub

Private Sub send_req()
    Dim xSht As Worksheet
    Dim xOutlookObj As Object
    Dim xEmailObj As Outlook.MailItem
    Dim job() As String
    Dim send_to As String
    Dim xYesorNo As Integer
    Set xSht = ActiveSheet
'GET DEFAULT EMAIL SIGNATURE
    Dim signature As String
    signature = Environ("appdata") & "\Microsoft\Signatures\"
    If Dir(signature, vbDirectory) <> vbNullString Then
        signature = signature & Dir$(signature & "*.txt")
    Else:
        signature = ""
    End If
    signature = CreateObject("Scripting.FileSystemObject").GetFile(signature).OpenAsTextStream(1, -2).ReadAll
    
    job = Split(ThisWorkbook.Path, "\")
    
    xFolder = ThisWorkbook.Path & "\" & ThisWorkbook.Worksheets("PO LOG").Range("pdf") & "\"
    xFolder = xFolder & job(3) & "_PO " & get_req_number(xSht.Index - (shts + 1) + ThisWorkbook.Worksheets("PO LOG").Range("Req_No")) _
    & "_" & Format(Now, "mm.dd.yy") & ".pdf"
    
    Set xOutlookObj = CreateObject("Outlook.Application")
    Set xEmailObj = xOutlookObj.CreateItem(olMailItem)
    With xEmailObj
        .To = ThisWorkbook.Worksheets("INSTRUCTIONS").Range("email_to")
        .CC = ThisWorkbook.Worksheets("INSTRUCTIONS").Range("email_cc")
        .Subject = job(3) & " PO " & xSht.Name
        
        xYesorNo = MsgBox("Preview E-mail?", vbYesNoCancel + vbQuestion, "Send?")
        If xYesorNo = vbYes Then
            .Display
            .Attachments.Add xFolder
        ElseIf xYesorNo = vbCancel Then
            MsgBox "PDF created, but not submitted", vbInformation
        Else
            .Body = "Hello," & vbNewLine & vbNewLine & "Attached is a PO Request for " & job(3) & vbNewLine & vbNewLine & signature
            .Attachments.Add xFolder
            .Send
        End If
    End With
End Sub

Public Sub save_send_req()
Attribute save_send_req.VB_ProcData.VB_Invoke_Func = "S\n14"
    If ActiveSheet.Name = "PO LOG" Or ActiveSheet.Name = "MASTER" Then
        MsgBox ("ERROR: Can not send " & ActiveSheet.Name)
        Exit Sub
    End If
    save_req
    Dim xYesorNo As Integer
    xYesorNo = MsgBox("Email Request?", vbYesNoCancel + vbQuestion, "Send?")
    If xYesorNo = vbYes Then
        send_req
    ElseIf xYesorNo = vbCancel Then
        MsgBox "PDF created, but not submitted", vbInformation
    Else
        ActiveSheet.PrintOut preview:=True
    End If
End Sub

