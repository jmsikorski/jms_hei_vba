VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Dim rng As Range
    Dim used_rng As Range
    Dim xFile As String
    Dim job As String
    job = PO_Module.get_job
    Me.Worksheets("PO LOG").Unprotect
    Me.Worksheets("PO LOG").Range("path") = Me.Path
    Me.Worksheets("PO LOG").Protect
    xFile = "R:\Data\Jobfiles\Purchasing Department\" & Dir("R:\Data\Jobfiles\Purchasing Department\" & Left(job, 6) & "*", vbDirectory)
    Set rng = Me.Worksheets("PO LOG").Range("link_start").Offset(1, 0)
    If rng <> vnullstring Then
        Set used_rng = Me.Worksheets("PO LOG").Range(Worksheets("PO LOG").Range("link_start").Offset(1, 0), Worksheets("PO LOG").Range("link_start").End(xlDown))
        For Each rng In used_rng
            If rng.Offset(0, 1) = vbNullString Then
                Dim xLink As String
                xLink = Dir(xFile & "\" & Left(job, 6) & "-" & rng.Offset(0, -1) & "*.pdf") 'the first one found
                If xLink <> "" Then
                    ThisWorkbook.Worksheets("PO LOG").Unprotect
                    xLink = xFile & "\" & xLink
                    ThisWorkbook.Worksheets("PO LOG").Hyperlinks.Add rng.Offset(0, 1), xLink, , , "PDF FILE"
                    ThisWorkbook.Worksheets("PO LOG").Protect
                End If
            End If
        Next
    End If
    Me.Worksheets("INSTRUCTIONS").Activate
End Sub
