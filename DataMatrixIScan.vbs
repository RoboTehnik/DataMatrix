Set fso = CreateObject("Scripting.FileSystemObject")
MyBaseFileName = fso.GetBaseName(WScript.ScriptName)

Set WshShell = CreateObject("WScript.Shell")
strComputerName = UCase(WshShell.ExpandEnvironmentStrings("%computername%"))

DT = Now
DateVal = Year(DT) & _
          Right("0" & Month(DT), 2) & _
          Right("0" & Day(DT), 2) & _
          Right("0" & Hour(DT), 2) & _
          Right("0" & Minute(DT), 2) & _
          Right("0" & Second(DT), 2)

Select Case strComputerName
  Case "HOST","ACER"
    Set OutputFile = fso.OpenTextFile("C:\WORK\DataMatrixInvent\" & MyBaseFileName & "_" & strComputerName & "_" & DateVal & ".txt", 2, True)
  Case Else
    Set OutputFile = fso.OpenTextFile("\\aptsrv\iadistrib\soft\" & MyBaseFileName & "_" & strComputerName & "_" & DateVal & ".txt", 2, True)
End Select

i = 0
Do
  RAW = ""
  RAW = InputBox("��� ������ ����������: " & strComputerName & vbCr & "������������ ��������� ��� ����������" & vbCr & "������������� �����: " & i, "������ ����� � vmsirenko@gmail.com")
  i = i + 1
  OutputFile.WriteLine(RuEn(RAW))
Loop Until RAW = ""
OutputFile.Close
WScript.Echo "������������ ���������!"
WScript.Quit

Function RuEn(InputText)
  RuEn = InputText
  RuEn = Replace(RuEn,"�","Q")
  RuEn = Replace(RuEn,"�","W")
  RuEn = Replace(RuEn,"�","E")
  RuEn = Replace(RuEn,"�","R")
  RuEn = Replace(RuEn,"�","T")
  RuEn = Replace(RuEn,"�","Y")
  RuEn = Replace(RuEn,"�","U")
  RuEn = Replace(RuEn,"�","I")
  RuEn = Replace(RuEn,"�","O")
  RuEn = Replace(RuEn,"�","P")
  RuEn = Replace(RuEn,"�","A")
  RuEn = Replace(RuEn,"�","S")
  RuEn = Replace(RuEn,"�","D")
  RuEn = Replace(RuEn,"�","F")
  RuEn = Replace(RuEn,"�","G")
  RuEn = Replace(RuEn,"�","H")
  RuEn = Replace(RuEn,"�","J")
  RuEn = Replace(RuEn,"�","K")
  RuEn = Replace(RuEn,"�","L")
  RuEn = Replace(RuEn,"�","Z")
  RuEn = Replace(RuEn,"�","X")
  RuEn = Replace(RuEn,"�","C")
  RuEn = Replace(RuEn,"�","V")
  RuEn = Replace(RuEn,"�","B")
  RuEn = Replace(RuEn,"�","N")
  RuEn = Replace(RuEn,"�","M")

  RuEn = Replace(RuEn,"�","q")
  RuEn = Replace(RuEn,"�","w")
  RuEn = Replace(RuEn,"�","e")
  RuEn = Replace(RuEn,"�","r")
  RuEn = Replace(RuEn,"�","t")
  RuEn = Replace(RuEn,"�","y")
  RuEn = Replace(RuEn,"�","u")
  RuEn = Replace(RuEn,"�","i")
  RuEn = Replace(RuEn,"�","o")
  RuEn = Replace(RuEn,"�","p")
  RuEn = Replace(RuEn,"�","a")
  RuEn = Replace(RuEn,"�","s")
  RuEn = Replace(RuEn,"�","d")
  RuEn = Replace(RuEn,"�","f")
  RuEn = Replace(RuEn,"�","g")
  RuEn = Replace(RuEn,"�","h")
  RuEn = Replace(RuEn,"�","j")
  RuEn = Replace(RuEn,"�","k")
  RuEn = Replace(RuEn,"�","l")
  RuEn = Replace(RuEn,"�","z")
  RuEn = Replace(RuEn,"�","x")
  RuEn = Replace(RuEn,"�","c")
  RuEn = Replace(RuEn,"�","v")
  RuEn = Replace(RuEn,"�","b")
  RuEn = Replace(RuEn,"�","n")
  RuEn = Replace(RuEn,"�","m")

End Function