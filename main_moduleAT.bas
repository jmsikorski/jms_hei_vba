Attribute VB_Name = "main_module"
Public Sub main()
    Dim d As Date
    Dim rng As Range
    Dim ws As Worksheet
    Dim tRange As Range
    Set ws = ThisWorkbook.Worksheets("ROSTER")
    ws.Unprotect "hei3078"
    d = calcWeek(Date)
    Dim x As ListObject
    Dim found As Boolean
    Dim i As Integer, cnt As Integer
    found = False
    Set rng = ws.UsedRange
    rng.EntireColumn.Hidden = False
    Set rng = ws.ListObjects(1).Range(1, 1)
    For i = 1 To ws.ListObjects(1).HeaderRowRange.Columns.Count
        If rng.offset(0, i).Value = Format(d, "m/d/yy") Then
            found = True
            Set tRange = ws.Range(Cells(1, 8), Cells(1, rng.offset(0, i).Column - 1))
        End If
        If found Then
            For cnt = 1 To 20
                If rng.offset(0, i + cnt).Value <> Format(d + cnt, "m/d/yy") Then
                    With ws.ListObjects(1)
                        .ListColumns.Add
                        .ListColumns(ws.ListObjects(1).ListColumns.Count).name = Format(d + cnt, "m/d/yy")
                        rng.AutoFilter Field:=.ListColumns(ws.ListObjects(1).ListColumns.Count).Range.Column, visibledropdown:=False
                    End With
                End If
            Next
            Set tRange = ws.Range(Cells(1, rng.offset(1, i + cnt).Column), Cells(1, ws.ListObjects(1).HeaderRowRange.Columns.Count))
            Exit For
        End If
    Next
    ws.Protect "hei3078"
    Set ws = Nothing
    Set rng = Nothing
End Sub

Public Sub pdf_report()
    Dim d As Date
    Dim rng As Range
    Dim ws As Worksheet
    Dim tRange As Range
    Dim pdfName As String
    Set ws = ThisWorkbook.Worksheets("ROSTER")
    ws.Unprotect "hei3078"
    d = calcWeek(Date)
    Debug.Print d - 6
    Dim x As ListObject
    Dim found As Boolean
    Dim i As Integer, cnt As Integer
    found = False
    Set rng = ws.UsedRange
    rng.EntireColumn.Hidden = False
    Set rng = ws.ListObjects(1).Range(1, 1)
    For i = 1 To ws.ListObjects(1).HeaderRowRange.Columns.Count
        If rng.offset(0, i).Value = Format(d, "m/d/yy") Then
            found = True
            Set tRange = ws.Range(Cells(1, 8), Cells(1, rng.offset(0, i).Column - 1))
            tRange.EntireColumn.Hidden = True
        End If
        If found Then
            Set tRange = ws.Range(Cells(1, rng.offset(1, i + 20).Column + 1), Cells(1, ws.ListObjects(1).HeaderRowRange.Columns.Count))
            tRange.EntireColumn.Hidden = True
            Exit For
        End If
    Next
    pdfName = Application.GetSaveAsFilename(InitialFileName:="Attendance_Report_" & Format(d, "m.d.yy") & ".pdf", FileFilter:="PDF Files (*.pdf), *.pdf")
    Debug.Print pdfName
    Set rng = ws.UsedRange
    rng.EntireColumn.Hidden = False
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfName, Quality:=xlQualityStandard
    ws.Protect "hei3078"
    Set ws = Nothing
    Set rng = Nothing
End Sub

Public Sub print_report()
    Dim d As Date
    Dim rng As Range
    Dim ws As Worksheet
    Dim tRange As Range
    Dim pdfName As String
    Set ws = ThisWorkbook.Worksheets("ROSTER")
    ws.Unprotect "hei3078"
    d = calcWeek(Date) - 6
    Debug.Print d
    Dim x As ListObject
    Dim found As Boolean
    Dim i As Integer, cnt As Integer
    found = False
    Set rng = ws.UsedRange
    rng.EntireColumn.Hidden = False
    Set rng = ws.ListObjects(1).Range(1, 1)
    For i = 1 To ws.ListObjects(1).HeaderRowRange.Columns.Count
        If rng.offset(0, i).Value = Format(d, "m/d/yy") Then
            found = True
            Set tRange = ws.Range(Cells(1, 8), Cells(1, rng.offset(0, i).Column - 1))
            tRange.EntireColumn.Hidden = True
        End If
        If found Then
            Set tRange = ws.Range(Cells(1, rng.offset(1, i + 20).Column + 1), Cells(1, ws.ListObjects(1).HeaderRowRange.Columns.Count))
            tRange.EntireColumn.Hidden = True
            Exit For
        End If
    Next
    ws.PrintOut preview:=True
    Set rng = ws.UsedRange
    rng.EntireColumn.Hidden = False
    ws.Protect "hei3078"
    Set ws = Nothing
    Set rng = Nothing
End Sub

