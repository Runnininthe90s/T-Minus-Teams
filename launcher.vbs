Dim objShell, fso, strPath, strCommand
Set objShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Get the exact folder this VBScript is running from
strPath = fso.GetParentFolderName(WScript.ScriptFullName)

' Build the PowerShell launch command
strCommand = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & strPath & "\T-Minus-Teams.ps1"""

' The '0' argument tells WScript to run this completely invisibly
objShell.Run strCommand, 0, False
