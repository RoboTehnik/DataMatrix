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
  RuEn = Replace(RuEn,"É","Q")
  RuEn = Replace(RuEn,"Ö","W")
  RuEn = Replace(RuEn,"Ó","E")
  RuEn = Replace(RuEn,"Ê","R")
  RuEn = Replace(RuEn,"Å","T")
  RuEn = Replace(RuEn,"Í","Y")
  RuEn = Replace(RuEn,"Ã","U")
  RuEn = Replace(RuEn,"Ø","I")
  RuEn = Replace(RuEn,"Ù","O")
  RuEn = Replace(RuEn,"Ç","P")
  RuEn = Replace(RuEn,"Ô","A")
  RuEn = Replace(RuEn,"Û","S")
  RuEn = Replace(RuEn,"Â","D")
  RuEn = Replace(RuEn,"À","F")
  RuEn = Replace(RuEn,"Ï","G")
  RuEn = Replace(RuEn,"Ð","H")
  RuEn = Replace(RuEn,"Î","J")
  RuEn = Replace(RuEn,"Ë","K")
  RuEn = Replace(RuEn,"Ä","L")
  RuEn = Replace(RuEn,"ß","Z")
  RuEn = Replace(RuEn,"×","X")
  RuEn = Replace(RuEn,"Ñ","C")
  RuEn = Replace(RuEn,"Ì","V")
  RuEn = Replace(RuEn,"È","B")
  RuEn = Replace(RuEn,"Ò","N")
  RuEn = Replace(RuEn,"Ü","M")

  RuEn = Replace(RuEn,"é","q")
  RuEn = Replace(RuEn,"ö","w")
  RuEn = Replace(RuEn,"ó","e")
  RuEn = Replace(RuEn,"ê","r")
  RuEn = Replace(RuEn,"å","t")
  RuEn = Replace(RuEn,"í","y")
  RuEn = Replace(RuEn,"ã","u")
  RuEn = Replace(RuEn,"ø","i")
  RuEn = Replace(RuEn,"ù","o")
  RuEn = Replace(RuEn,"ç","p")
  RuEn = Replace(RuEn,"ô","a")
  RuEn = Replace(RuEn,"û","s")
  RuEn = Replace(RuEn,"â","d")
  RuEn = Replace(RuEn,"à","f")
  RuEn = Replace(RuEn,"ï","g")
  RuEn = Replace(RuEn,"ð","h")
  RuEn = Replace(RuEn,"î","j")
  RuEn = Replace(RuEn,"ë","k")
  RuEn = Replace(RuEn,"ä","l")
  RuEn = Replace(RuEn,"ÿ","z")
  RuEn = Replace(RuEn,"÷","x")
  RuEn = Replace(RuEn,"ñ","c")
  RuEn = Replace(RuEn,"ì","v")
  RuEn = Replace(RuEn,"è","b")
  RuEn = Replace(RuEn,"ò","n")
  RuEn = Replace(RuEn,"ü","m")
End Function


Sub MakeXML()
  xmlDoc.appendChild(xmlDoc.createProcessingInstruction("xml", "version='1.0' encoding='windows-1251'"))
  Set rootNode = xmlDoc.appendChild(xmlDoc.createElement("INVENT"))
  xmlDoc.Save(MyBaseFileName & ".xml")
End Sub