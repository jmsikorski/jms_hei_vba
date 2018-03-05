VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Log_Attendance 
   Caption         =   "Attendance Tracking"
   ClientHeight    =   9300.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8280.001
   OleObjectBlob   =   "Log_Attendance.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Log_Attendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ty() As String
Const Holiday = 15773696
Const PAF_Vacation = 5296274
Const PAF_DOWP = 65535
Const PAF_Unpaid = 49407
Const Sick_Other = 255

Private Sub Attendance_Enter_Click()
    Dim tBox As Object
    Dim day As Date
    Dim eNumber As Integer
    Dim clr(4) As Double
    Dim rng As Range
    Dim r As Integer, c As Integer
    Dim note As String
    If Me.ComboBox1.Value = vbNullString Then
        MsgBox "Invalid Date", vbCritical + vbOK, "DATE ERROR"
        Exit Sub
    End If
    If Me.ComboBox2.Value = vbNullString Then
        MsgBox "Invalid Employee Number", vbCritical + vbOKCancel, "EMPLOYEE NUMBER ERROR"
        Exit Sub
    End If
    For Each rng In ThisWorkbook.Worksheets("ROSTER").ListObjects(1).HeaderRowRange
        If rng.Value = Format(Me.ComboBox1.Value, "m/d/yy") Then
            c = rng.Column
        End If
    Next
    For Each rng In ThisWorkbook.Worksheets("ROSTER").ListObjects(1).ListColumns("EMP #").DataBodyRange
        Dim tmp As String
        tmp = rng.Value
        If tmp = Me.ComboBox2.Value Then
            r = rng.Row
        End If
    Next
    note = TextBox1.Value
    Set rng = ThisWorkbook.Worksheets("ROSTER").Cells(r, c)
    clr(0) = Holiday
    clr(1) = PAF_Vacation
    clr(2) = PAF_DOWP
    clr(3) = PAF_Unpaid
    clr(4) = Sick_Other
    day = Me.ComboBox1.Value
    eNumber = Me.ComboBox2
    Set tBox = Me.Controls.Item("tBoxList")
    For i = 0 To tBox.ListCount - 1
        If tBox.Selected(i) Then
            If tBox.List(i) <> "Clear" Then
                rng.Value = note
                rng.Interior.color = clr(i)
                tBox.Selected(i) = False
                Exit For
            Else
                rng.Interior.color = ThisWorkbook.Worksheets("ROSTER").ListObjects(1).Range(rng.Row, 1).Interior.ThemeColor
                rng.Value = vbNullString
            End If
        End If
    Next
    Me.ComboBox1.ListIndex = Me.ComboBox1.ListIndex + 1
    'Me.ComboBox2.Value = vbNullString
    Set rng = Nothing
End Sub

Private Sub clr_Click()
    Me.ComboBox2.Value = vbNullString
    Me.ComboBox3.Value = vbNullString
    For i = 0 To Me.Controls.Item("tBoxList").ListCount - 1
        Me.Controls.Item("tBoxList").Selected(i) = False
    Next
End Sub

Private Sub ComboBox1_Change()
    Dim ans As Integer
    If Me.ComboBox1.Value = vbNullString Then
        ans = MsgBox("Invalid Date", vbExclamation + vbOKCancel, "DATE ERROR")
        If ans = vbOK Then
            Me.ComboBox1.Value = Date
        End If
    End If
End Sub

Private Sub ComboBox2_Change()
    On Error Resume Next
    Dim r As Integer
    Dim name As String
    Dim rng As Range
    Dim x As ListObject
    Me.ComboBox3.ListIndex = Me.ComboBox2.ListIndex
'    r = Me.ComboBox2.ListIndex
'    With ThisWorkbook.Worksheets("ROSTER").ListObjects("emp_roster")
'        Set rng = .ListRows(r + 1).Range
'        name = rng.Cells(1, .ListColumns("FIRST").Range.Column).Value & " " & rng.Cells(1, .ListColumns("LAST NAME").Range.Column)
'        TextBox1.Value = name
'    End With
    On Error GoTo 0
End Sub

Private Sub ComboBox3_Change()
    Me.ComboBox2.ListIndex = Me.ComboBox3.ListIndex
End Sub

Private Sub done_Click()
    Dim ans As Integer
    ans = vbYes
    Set tBox = Me.Controls.Item("tBoxList")
    For i = 0 To Me.Controls.Item("tBoxList").ListCount - 1
        If Me.Controls.Item("tBoxList").Selected(i) = True And Me.ComboBox2.Value <> vbNullString Then
            ans = MsgBox("Are you sure you're done?", vbQuestion + vbYesNo, "Edit " & Me.ComboBox3.Value & "?")
        End If
    Next
    If ans = vbNo Or ans = vbCancel Then
        Exit Sub
    Else
        Unload Me
    End If
End Sub

Private Sub Frame1_Click()

End Sub

'Private Sub TextBox1_Change()
'    On Error Resume Next
'    Dim name As String
'    Dim rng As Range
'    For Each rng In ThisWorkbook.Worksheets("ROSTER").ListObjects("emp_roster").ListColumns("FIRST").DataBodyRange
'        name = rng & " " & rng.Offset(0, -1)
'        If name = Me.TextBox1.Value Then
'            Me.ComboBox2.ListIndex = rng.Row - 2
'            Exit Sub
'        End If
'    Next
'    Me.ComboBox2.Value = vbNullString
'End Sub

Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
    Dim i As Integer
    Dim day As Date
    Dim t As Range
    day = calcWeek(Date)
    day = day - 7
    With Me.ComboBox1
        For i = 0 To 20
            .AddItem day + i
        Next i
    End With
    Me.ComboBox1.Value = Date
    Me.ComboBox1.Value = Cells(1, ActiveCell.Column)
    With Me.ComboBox2
        For Each t In ThisWorkbook.Worksheets("ROSTER").ListObjects("emp_roster").ListColumns("EMP #").DataBodyRange
            .AddItem t.Value
        Next
    End With
    Me.ComboBox2.Value = Cells(ActiveCell.Row, 5)
    With Me.ComboBox3
        For Each t In ThisWorkbook.Worksheets("ROSTER").ListObjects("emp_roster").ListColumns("FIRST").DataBodyRange
            .AddItem t.Value & " " & t.offset(0, -1).Value
        Next
    End With
    Me.ComboBox3.ListIndex = Me.ComboBox2.ListIndex
    With Frame1
        .Height = 147
        .Width = 210
        .Left = 102
        .Top = 243
        .Caption = "Select Attendance Event"
    End With
    Dim tBox As Control
    ty = Split("Holiday/PAF Vacation/PAF DOWP/PAF Unpaid/SICK, LATE, NO CALL/Clear", "/")
    Set tBox = Frame1.Controls.Add("Forms.ListBox.1", "tBoxList", True)
    With tBox
        .Height = Frame1.Height - 18
        .Width = Frame1.Width - 12
        .Left = 6
        .Top = 6
        .SpecialEffect = 0
        .ListStyle = 1
        .MultiSelect = 0
    End With
    For i = 0 To UBound(ty)
        tBox.AddItem ty(i)
    Next
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
       done_Click
    End If
End Sub

