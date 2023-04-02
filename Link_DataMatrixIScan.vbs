Option Explicit

Dim oShell, oShortCut, sDeskTopPath

Set oShell = CreateObject("Wscript.Shell")

sDeskTopPath = oShell.SpecialFolders("Desktop")

Set oShortCut = oShell.CreateShortcut(sDeskTopPath & "\Сканер учета.lnk")

oShortCut.IconLocation = "C:\Windows\System32\wscript.exe, 0"

oShortCut.TargetPath = "C:\Windows\System32\wscript.exe"

oShortCut.Arguments = "DataMatrixIScan.vbs"

oShortCut.WorkingDirectory = oShell.CurrentDirectory

oShortCut.Save