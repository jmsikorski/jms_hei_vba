VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Edit_Roster 
   Caption         =   "Edit Roster"
   ClientHeight    =   7185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8280.001
   OleObjectBlob   =   "Edit_Roster.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Edit_Roster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim editRange As Range

Private Sub ComboBox2_Change()
    Dim ans As Integer
    If checkYes(Mid(Me.ComboBox2.Text, 1, 1)) Then
        Me.ComboBox2.ListIndex = 0
        Exit Sub
    ElseIf checkNo(Mid(Me.ComboBox2.Text, 1, 1)) Then
        Me.ComboBox2.ListIndex = 1
        Exit Sub
    ElseIf Me.ComboBox2.Value = vbNullString Then
        Exit Sub
    Else
        ans = MsgBox("Invalid Entry", vbExclamation + vbOKCancel, "YES OR NO")
        If ans = vbOK Then
            ComboBox2.Text = vbNullString
        Else
            Edit_Roster_Cancel_Click
        End If
    End If
End Sub

Private Sub ComboBox3_Change()
    Dim ans As Integer
    If checkYes(Mid(Me.ComboBox3.Text, 1, 1)) Then
        Me.ComboBox3.ListIndex = 0
        Exit Sub
    ElseIf checkNo(Mid(Me.ComboBox3.Text, 1, 1)) Then
        Me.ComboBox3.ListIndex = 1
        Exit Sub
    ElseIf Me.ComboBox3.Text = vbNullString Then
        Exit Sub
    Else
        ans = MsgBox("Invalid Entry", vbExclamation + vbOKCancel, "YES OR NO")
        If ans = vbOK Then
            ComboBox3.Text = vbNullString
        Else
            Edit_Roster_Cancel_Click
        End If
    End If
End Sub

Private Sub Edit_Roster_Cancel_Click()
    Unload Me
End Sub

Private Sub Edit_Roster_Enter_Click()
    Dim rng As Range
    Dim ans As Integer
    Dim padding As Integer
    Dim class As String
    Dim lName As String
    Dim fName As String
    Dim emNum As Double
    Dim perDiem As String
    Dim active As String
    padding = 20
    If Me.ComboBox1.Value = vbNullString Then
        MsgBox "Invalid Class", vbCritical + vbOK, "CLASS ERROR"
        Me.ComboBox1.Value = editRange.Cells(1, 1)
        Exit Sub
    End If
    If Me.ComboBox2.Value = vbNullString Then
        MsgBox "Invalid Per Diem Entry", vbCritical + vbOK, "PER DIEM ERROR"
        Me.ComboBox1.Value = editRange.Cells(1, 5)
        Exit Sub
    End If
    If Me.ComboBox3.Value = vbNullString Then
        MsgBox "Invalid Active Entry", vbCritical + vbOK, "ACTIVE ERROR"
        Me.ComboBox1.Value = editRange.Cells(1, 6)
        Exit Sub
    End If
    For Each rng In ThisWorkbook.Worksheets("ROSTER").ListObjects(1).ListColumns("EMP #").DataBodyRange
        If rng.Value = CDbl(Me.TextBox3.Value) Then
            If Application.Intersect(rng, editRange) Is Nothing Then
                MsgBox "That Employee Number already exsits", vbCritical + vbOK, "EMPLOYEE NUMBER ERROR"
                Me.TextBox3.Value = editRange.Cells(1, 4)
                Exit Sub
            End If
        End If
    Next
    class = Me.ComboBox1.Value
    lName = Me.TextBox1.Value
    fName = Me.TextBox2.Value
    emNum = CDbl(Me.TextBox3.Value)
    If Me.ComboBox2.Value = "YES" Then
        perDiem = True
    Else
        perDiem = False
    End If
    If Me.ComboBox3.Value = "YES" Then
        active = True
    Else
        active = False
    End If
    Dim confirm As String
        
    If editRange.Cells(1, 1) <> class Then
        confirm = "Change Class to " & class & "?"
        On Error GoTo clean_up
        If checkConfirm(confirm) Then
            editRange.Cells(1, 1) = class
        Else
            Me.ComboBox1.Value = editRange.Cells(1, 1)
            GoTo clean_up
        End If
    End If
    If editRange.Cells(1, 2) <> lName Then
        confirm = "Change Last Name to " & lName & "?"
        On Error GoTo clean_up
        If checkConfirm(confirm) Then
            editRange.Cells(1, 2) = lName
        Else
            Me.TextBox1.Value = editRange.Cells(1, 2)
            GoTo clean_up
        End If
    End If
    If editRange.Cells(1, 3) <> fName Then
        confirm = "Change First Name to " & fName & "?"
        On Error GoTo clean_up
        If checkConfirm(confirm) Then
            editRange.Cells(1, 3) = fName
        Else
            Me.TextBox2.Value = editRange.Cells(1, 3)
            GoTo clean_up
        End If
    End If
    If editRange.Cells(1, 4) <> emNum Then
        confirm = "Change Employee Number to " & emNum & "?"
        On Error GoTo clean_up
        If checkConfirm(confirm) Then
            editRange.Cells(1, 4) = emNum
        Else
            Me.TextBox3.Value = editRange.Cells(1, 4)
            GoTo clean_up
        End If
    End If
    If perDiem Then
        If editRange.Cells(1, 5) <> "YES" Then
            confirm = "Change PerDiem to YES?"
            On Error GoTo clean_up
            If checkConfirm(confirm) Then
                editRange.Cells(1, 5) = "YES"
            Else
                Me.ComboBox2.Value = editRange.Cells(1, 5)
                GoTo clean_up
            End If
        End If
    Else
        If editRange.Cells(1, 5) <> "NO" Then
            confirm = "Change PerDiem to NO?"
            On Error GoTo clean_up
            If checkConfirm(confirm) Then
                editRange.Cells(1, 5) = "NO"
            Else
                Me.ComboBox2.Value = editRange.Cells(1, 5)
                GoTo clean_up
            End If
        End If
    End If
    If active Then
        If editRange.Cells(1, 6) <> "YES" Then
            confirm = "Change Active to YES?"
            On Error GoTo clean_up
            If checkConfirm(confirm) Then
                editRange.Cells(1, 6) = "YES"
            Else
                Me.ComboBox2.Value = editRange.Cells(1, 6)
                GoTo clean_up
            End If
        End If
    Else
        If editRange.Cells(1, 6) <> "NO" Then
            confirm = "Change Active to NO?"
            On Error GoTo clean_up
            If checkConfirm(confirm) Then
                editRange.Cells(1, 6) = "NO"
            Else
                Me.ComboBox2.Value = editRange.Cells(1, 6)
                GoTo clean_up
            End If
        End If
    End If
clean_up:
    Set rng = Nothing
    Edit_Roster_Cancel_Click
End Sub

Private Function checkConfirm(t As String) As Integer
    ans = MsgBox(t, vbYesNoCancel, "CONFIRM")
    If ans = vbNo Then
        checkConfirm = 0
    ElseIf ans = vbCancel Then
        checkConfirm = -1
    ElseIf ans = vbYes Then
        checkConfirm = 1
    End If
End Function

Private Sub UserForm_Initialize()
    Dim buffer As Integer
    Dim vOffset As Integer
    Dim cBox() As Integer
    Dim tBox() As Integer
    Dim rng As Range
    ReDim cBox(2)
    ReDim tBox(2)
    cBox(0) = 1
    cBox(1) = 5
    cBox(2) = 6
    tBox(0) = 2
    tBox(1) = 3
    tBox(2) = 4
    vOffset = -3
    buffer = 12
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
    Dim t As Range
    Dim ctrl As Control
    Dim ctrlCnt As Integer
    Dim ctrlName As String
    Dim i As Integer
    Dim a As Integer
    Dim cls() As String
    Dim clsCnt As Integer
    clsCnt = 0
    ctrlCnt = 1
    Set editRange = ThisWorkbook.Worksheets("ROSTER").Range(Cells(ActiveCell.Row, 2), Cells(ActiveCell.Row, 7))
    For Each ctrl In Me.Controls
        If ctrl.name = "Label" & ctrlCnt Then
            ctrlCnt = ctrlCnt + 1
        End If
    Next
    For i = 1 To ctrlCnt - 1
        ctrlName = vbNullString
        For a = 0 To UBound(tBox)
            If tBox(a) = i Then
                ctrlName = "TextBox" & a + 1
                Exit For
            End If
        Next
        If ctrlName = vbNullString Then
            For a = 0 To UBound(cBox)
                If cBox(a) = i Then
                    ctrlName = "ComboBox" & a + 1
                    Exit For
                End If
            Next
        End If
        With Me.Controls("Label" & i)
            Debug.Print ctrlName
            Me.Controls(ctrlName).Left = .Left + .Width + buffer
            Me.Controls(ctrlName).Top = .Top + vbuffer
        End With
        Me.Controls(ctrlName).TabIndex = i
    Next
    Dim found As Boolean
    found = False
    ReDim cls(0)
    With Me.ComboBox1
        For Each rng In ActiveSheet.ListObjects(1).ListColumns(2).DataBodyRange
            For i = 0 To UBound(cls)
                If cls(i) = rng.Value Then
                    found = True
                    Exit For
                End If
                found = False
            Next
            If Not found Then
                ReDim Preserve cls(clsCnt)
                cls(clsCnt) = rng.Value
                clsCnt = clsCnt + 1
            End If
        Next
        For i = 0 To UBound(cls)
            .AddItem cls(i)
        Next i
        .Value = Cells(ActiveCell.Row, 2)
    End With
    With Me.ComboBox2
        .AddItem "YES"
        .AddItem "NO"
        .Value = Cells(ActiveCell.Row, 6)
    End With
    With Me.ComboBox3
        .AddItem "YES"
        .AddItem "NO"
        .Value = Cells(ActiveCell.Row, 7)
    End With
    Me.TextBox1.Value = Cells(ActiveCell.Row, 3)
    Me.TextBox2.Value = Cells(ActiveCell.Row, 4)
    Me.TextBox3.Value = Cells(ActiveCell.Row, 5)
        
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
       Edit_Roster_Cancel_Click
    End If
End Sub

Public Function checkYes(t As String) As Boolean
    If t = "Y" Or t = "y" Then
        checkYes = True
    ElseIf Len(t) > 1 Then
        checkYes = False
    End If
End Function

Public Function checkNo(t As String) As Boolean
    If t = "N" Or t = "n" Then
        checkNo = True
    ElseIf Len(t) > 1 Then
        checkNo = False
    End If
    
End Function

