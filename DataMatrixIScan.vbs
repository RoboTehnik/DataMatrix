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
  RAW = InputBox("Имя вашего компьютера: " & strComputerName & vbCr & "Отсканируйте следующий код маркировки" & vbCr & "Отсканировано кодов: " & i, "Сканер учета © vmsirenko@gmail.com")
  i = i + 1
  OutputFile.WriteLine(RuEn(RAW))
Loop Until RAW = ""
OutputFile.Close
WScript.Echo "Сканирование закончено!"
WScript.Quit

Function RuEn(InputText)
  RuEn = InputText
  RuEn = Replace(RuEn,"Й","Q")
  RuEn = Replace(RuEn,"Ц","W")
  RuEn = Replace(RuEn,"У","E")
  RuEn = Replace(RuEn,"К","R")
  RuEn = Replace(RuEn,"Е","T")
  RuEn = Replace(RuEn,"Н","Y")
  RuEn = Replace(RuEn,"Г","U")
  RuEn = Replace(RuEn,"Ш","I")
  RuEn = Replace(RuEn,"Щ","O")
  RuEn = Replace(RuEn,"З","P")
  RuEn = Replace(RuEn,"Ф","A")
  RuEn = Replace(RuEn,"Ы","S")
  RuEn = Replace(RuEn,"В","D")
  RuEn = Replace(RuEn,"А","F")
  RuEn = Replace(RuEn,"П","G")
  RuEn = Replace(RuEn,"Р","H")
  RuEn = Replace(RuEn,"О","J")
  RuEn = Replace(RuEn,"Л","K")
  RuEn = Replace(RuEn,"Д","L")
  RuEn = Replace(RuEn,"Я","Z")
  RuEn = Replace(RuEn,"Ч","X")
  RuEn = Replace(RuEn,"С","C")
  RuEn = Replace(RuEn,"М","V")
  RuEn = Replace(RuEn,"И","B")
  RuEn = Replace(RuEn,"Т","N")
  RuEn = Replace(RuEn,"Ь","M")

  RuEn = Replace(RuEn,"й","q")
  RuEn = Replace(RuEn,"ц","w")
  RuEn = Replace(RuEn,"у","e")
  RuEn = Replace(RuEn,"к","r")
  RuEn = Replace(RuEn,"е","t")
  RuEn = Replace(RuEn,"н","y")
  RuEn = Replace(RuEn,"г","u")
  RuEn = Replace(RuEn,"ш","i")
  RuEn = Replace(RuEn,"щ","o")
  RuEn = Replace(RuEn,"з","p")
  RuEn = Replace(RuEn,"ф","a")
  RuEn = Replace(RuEn,"ы","s")
  RuEn = Replace(RuEn,"в","d")
  RuEn = Replace(RuEn,"а","f")
  RuEn = Replace(RuEn,"п","g")
  RuEn = Replace(RuEn,"р","h")
  RuEn = Replace(RuEn,"о","j")
  RuEn = Replace(RuEn,"л","k")
  RuEn = Replace(RuEn,"д","l")
  RuEn = Replace(RuEn,"я","z")
  RuEn = Replace(RuEn,"ч","x")
  RuEn = Replace(RuEn,"с","c")
  RuEn = Replace(RuEn,"м","v")
  RuEn = Replace(RuEn,"и","b")
  RuEn = Replace(RuEn,"т","n")
  RuEn = Replace(RuEn,"ь","m")

End Function