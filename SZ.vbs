' Sluzhebnye zapiski: zapusk bez okna konsoli.
' Rabotaet i iz setevoy papki (\\server\...), v otlichie ot SZ.bat:
' cmd.exe ne umeet delat' takuyu papku tekushchey.
Option Explicit
Dim fso, shell, folder, pythonw, script
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
folder = fso.GetParentFolderName(WScript.ScriptFullName)
pythonw = folder & "\environment\python\pythonw.exe"
script = folder & "\main.pyw"
If Not fso.FileExists(pythonw) Then
  MsgBox "Ne naiden " & pythonw & vbCrLf & "Skopiruite papku programmy celikom, vmeste s environment.", vbCritical, "SZ"
  WScript.Quit 1
End If
shell.Run """" & pythonw & """ """ & script & """", 1, False
