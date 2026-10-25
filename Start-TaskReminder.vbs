' TaskReminder - silent launcher
' Called by wscript.exe so no console window flashes.
Option Explicit

Dim fso, shell, scriptDir, ps1, cmd, extra, i
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
ps1 = scriptDir & "\TaskReminder.ps1"

If Not fso.FileExists(ps1) Then
    MsgBox "TaskReminder.ps1 not found in:" & vbCrLf & scriptDir, 16, "TaskReminder"
    WScript.Quit 1
End If

extra = ""
For i = 0 To WScript.Arguments.Count - 1
    extra = extra & " " & WScript.Arguments(i)
Next

cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & ps1 & """" & extra
' 0 = hidden, False = do not wait
shell.Run cmd, 0, False
