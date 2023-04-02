Option Explicit

Dim oShell, oShortCut, sDeskTopPath

Set oShell = CreateObject("Wscript.Shell")

sDeskTopPath = oShell.SpecialFolders("Desktop")

Set oShortCut = oShell.CreateShortcut(sDeskTopPath & "\Печать маркировки.lnk")

oShortCut.IconLocation = "C:\Windows\System32\wscript.exe , 0"

oShortCut.TargetPath = "C:\Windows\System32\wscript.exe"

oShortCut.Arguments = "DataMatrixPRN.vbs"

oShortCut.WorkingDirectory = oShell.CurrentDirectory

oShortCut.Save