Option Explicit
Dim s, f, dir
Set s = CreateObject("WScript.Shell")
Set f = CreateObject("Scripting.FileSystemObject")
dir = f.GetParentFolderName(WScript.ScriptFullName)
s.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & dir & "\agent_bridge.ps1""", 0, False
