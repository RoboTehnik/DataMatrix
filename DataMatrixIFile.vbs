Set fso = CreateObject("Scripting.FileSystemObject")
MyBaseFileName = fso.GetBaseName(WScript.ScriptName)

Set xmlDoc = CreateObject("Microsoft.XMLDOM")
xmlDoc.Async = "False"
If fso.FileExists(MyBaseFileName & ".xml") = False Then Call MakeXML()
xmlDoc.Load(MyBaseFileName & ".xml")
Set xmlINVENT = xmlDoc.documentElement

Set InputFile = fso.OpenTextFile(MyBaseFileName & ".txt", 1, False)
Set OutputFile = fso.OpenTextFile(MyBaseFileName & ".csv", 2, True)
Do Until InputFile.AtEndOfStream
  RAW = InputFile.ReadLine
  If Left(RAW,2) = "01" Then
      GTIN = GetGTIN(RAW)
      SN = GetSN(RAW)
      SGTIN = GTIN & SN
      IID = GetIID(SGTIN)
      If Not IsEmpty(IID) Then 
        If xmlDoc.selectNodes("//INVENT/SGTIN[text()='" & SGTIN & "']").length = 0 Then
          OutputFile.WriteLine(IID & ",1")
          Set xmlSGTINnew = xmlINVENT.appendChild(xmlDoc.createElement("SGTIN"))
          xmlSGTINnew.Text = SGTIN
          xmlDoc.Save(MyBaseFileName & ".xml")
          WScript.Echo IID, SGTIN
        End If
      End If
  End If
Loop
InputFile.Close
OutputFile.Close
WScript.Echo "All Done!"
WScript.Quit

Function GetIID(SGTIN)
  Set ODBCconnect = CreateObject("ADODB.Connection")
  'ODBCconnect.Open "Driver=Firebird/InterBase(r) driver;CHARSET=WIN1251;UID=SYSDBA;PWD=masterkey;DbName=aptsrv/3052:C:\IADB\IAPTEKA.FDB"
  ODBCconnect.Open "Driver=Firebird/InterBase(r) driver;CHARSET=WIN1251;UID=SYSDBA;PWD=masterkey;DbName=coserv/3052:C:\IADB\IAPTEKA.FDB"
  Set FBrecordset = ODBCconnect.Execute("select dis.iid from docitem_sgtin dis where dis.sgtin='" & SGTIN & "'")
    While Not FBrecordset.EOF
      GetIID = FBrecordset.Fields(0).Value
      FBrecordset.MoveNext
    Wend
  ODBCconnect.Close
End Function

Function GetGTIN(RAW)
  GetGTIN = Mid(RAW,3,14)
End Function

Function GetSN(RAW)
  GetSN = RuEn(Mid(RAW,19,13))
End Function

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


Sub MakeXML()
  xmlDoc.appendChild(xmlDoc.createProcessingInstruction("xml", "version='1.0' encoding='windows-1251'"))
  Set rootNode = xmlDoc.appendChild(xmlDoc.createElement("INVENT"))
  xmlDoc.Save(MyBaseFileName & ".xml")
End Sub