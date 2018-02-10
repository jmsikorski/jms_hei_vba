
Dim Arg, var1
Set Arg = WScript.Arguments

'Parameter1, begin with index0
var1 = Arg(0)

msgbox "Commit message is " & var1

'Clear the objects at the end of your script.

stdout = Console(var1)
MsgBox stdout
set Arg = Nothing

Function Console(strMessage)
'@description: Run command prompt command and get its output.
'@author: Jeremy England ( SimplyCoded )
  Dim Wss, Cmd, Return, Output
  Set Wss = CreateObject("WScript.Shell")
  Set Cmd = Wss.Exec("cmd.exe")
  Cmd.StdIn.WriteLine "cd C:\ProgramData\Git" & " 2>&1"
  Cmd.StdIn.WriteLine "git-cmd.exe" & " 2>&1"
  Cmd.StdIn.WriteLine "cd C:\Users\jsikorski\Documents\VBAProjectFiles\ALL VBA CODE\jms_hei_vba" & " 2>&1"
  Cmd.StdIn.WriteLine "git add *" & " 2>&1"
  Cmd.StdIn.WriteLine "git commit -m """ & strMessage  & """ 2>&1"
  Cmd.StdIn.Close
  While InStr(Cmd.StdOut.ReadLine, ">" & strCmd) = 0 : Wend
  Do : Output = Cmd.StdOut.ReadLine
    If Cmd.StdOut.AtEndOfStream Then Exit Do _
    Else Return = Return & Output & vbLf
  Loop
  Console = Return
End Function
