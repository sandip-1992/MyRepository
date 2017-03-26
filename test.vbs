Dim objShell
Set objShell = CreateObject ("WScript.Shell")
filepath="D:\New folder\program.exe"
objShell.Run (chr(34) & filepath & chr(34))