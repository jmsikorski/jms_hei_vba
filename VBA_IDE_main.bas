Attribute VB_Name = "VBA_IDE_main"
Option Explicit

Public Sub testGitCommit()
    gitCommit "Commit Passed Through VBA", "C:\Users\jsikorski\Documents\VBAProjectFiles\ALL VBA CODE\jms_hei_vba\"
End Sub

Public Sub gitCommit(tString As String, gitRepPath As String)
'Written by: Jason Sikorski: sikorsk4@gmail.com
'Argument 1: tString - String value for commit message
'Argument 2: gitRepPath - Full path of git repositroy
'Requirements: VBS_Concole.vbs in gitRepPath
    Dim gitVBS As String
    Dim cmdString As String
    Dim arg1 As String
    gitVBS = "VBS_Console.vbs"
    arg1 = wrapString(tString)
    gitRepPath = wrapString(gitRepPath & " " & gitVBS)
    cmdString = "cscript " & gitRepPath & " " & arg1
    Debug.Print cmdString
'    Shell cmdString, vbNormalFocus
End Sub

Private Function wrapString(txt As String) As String
    wrapString = """" & txt & """"
    Debug.Print wrapString
End Function
