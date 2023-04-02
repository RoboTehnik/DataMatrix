Set fso = CreateObject("Scripting.FileSystemObject")
Set xmlDoc = CreateObject("Microsoft.XMLDOM")
MyBaseFileName = fso.GetBaseName(WScript.ScriptName)
If fso.FileExists(MyBaseFileName & ".xml") = False Then Call MakeXML()
xmlDoc.Async = "False"
xmlDoc.Load(MyBaseFileName & ".xml")
barcode_port = xmlDoc.selectNodes("//BARCODE/PORT")(0).Text
barcode_prn = MyBaseFileName & ".prn"
If DB_PATH_Exist() = "" Then Call DB_PATH_UPD("APTSRV/3052:C:\IADB\IAPTEKA.FDB")
DB_PATH = xmlDoc.selectNodes("//BARCODE/DB_PATH")(0).Text

Do
  SGTIN = InputBox("Введите SGTIN для печати кода маркировки", "Печать кода маркировки © vmsirenko@gmail.com")
  If SGTIN = "" Then Exit Do
  If Len(SGTIN) = 83 Then
    SGTIN = RuEn(SGTIN)
    MARK_CODE = Ink_GS(SGTIN)
    SN = Mid(SGTIN,19,13)
    Call PrintLabel(barcode_port,MARK_CODE,SN)
  End If
  If Len(SGTIN) = 27 Then
    EAN13 = Left(SGTIN,14)
    SN = Right(SGTIN,13)
    NODE_SOURCE = GetNodeSource(SGTIN)
    If Not NODE_SOURCE = "" Then
      WScript.Echo "Товар поступил из аптеки №" & NODE_SOURCE & vbCr & "Сейчас будет выполнена попытка подключения к другой аптеке"
      DB_PATH_HOST = DB_PATH
      DB_PATH = "APTSRV/40" & Left("0" & NODE_SOURCE,2) & ":C:\IADB\IAPTEKA.FDB"
    End if
    MARK_CODE = GetMarkCode(SGTIN)
    MARK_CODE = Ink_GS(MARK_CODE)
    Call PrintLabel(barcode_port,MARK_CODE,SN)
    DB_PATH = DB_PATH_HOST
  End If
  If Left(SGTIN,2) = "* " Then
    Set fput = fso.OpenTextFile(barcode_prn, 2, True)
    Set colNodes=xmlDoc.selectNodes("//BARCODE/LABEL/INIT/COMMAND")
    For Each objNode in colNodes
      fput.WriteLine(objNode.Text)
    Next
        fput.WriteLine "A55,25,0,4,1,1,N,""" & win1251_cp866(Mid(SGTIN,3)) & """"
        fput.WriteLine("P10")
        fput.Close
        fso.CopyFile barcode_prn, barcode_port
        fso.DeleteFile barcode_prn, True
  End If
  If SGTIN = "***" Then
    Set fput = fso.OpenTextFile(barcode_prn, 2, True)
    Set colNodes=xmlDoc.selectNodes("//BARCODE/LABEL/INIT/COMMAND")
    For Each objNode in colNodes
      fput.WriteLine(objNode.Text)
    Next
        fput.WriteLine "A55,25,0,4,1,1,N,""ЃЂ’Ђђ…џ"""
        fput.WriteLine "A30,55,0,4,1,1,N,""“‘’ЂЌЋ‚‹…ЌЂ"""
        fput.WriteLine "A25,100,0,4,1,1,N,""" & Date() & """"
        fput.WriteLine("P1")
        fput.Close
        fso.CopyFile barcode_prn, barcode_port
        fso.DeleteFile barcode_prn, True
  End If
Loop
WScript.Quit

Sub PrintLabel(barcode_port,MARK_CODE,SN)
  Select Case LCase(barcode_port)
    Case "gif", "jpg", "png", "svg"
      dim xHttp: Set xHttp = createobject("Microsoft.XMLHTTP")
      dim bStrm: Set bStrm = createobject("Adodb.Stream")
      'https://barcode.tec-it.com/barcode.ashx?data=010460221000338121100000148030191EE0692UQ4EFdSUB%2ByGfv4ZUneDIy%2FJxRvajviZB4jV3By0YZI%3D&code=GS1DataMatrix&translate-esc=true&dmsize=Default
      xHttp.Open "GET", "https://barcode.tec-it.com/barcode.ashx?data=" & MARK_CODE & "&code=DataMatrix&multiplebarcodes=false&translate-esc=true&unit=Fit&dpi=96&imagetype=" & barcode_port & "&rotation=0&color=%23000000&bgcolor=%23ffffff&codepage=Default&qunit=Mm&quiet=0&dmsize=Default", "False"
      xHttp.Send
      with bStrm
        .type = 1 '//binary
        .open
        .write xHttp.responseBody
        .savetofile MyBaseFileName & "." & barcode_port, 2 '//overwrite
      end with
    Case Else
      Set fput = fso.OpenTextFile(barcode_prn, 2, True)
      Set colNodes=xmlDoc.selectNodes("//BARCODE/LABEL/INIT/COMMAND")
      For Each objNode in colNodes
        fput.WriteLine(objNode.Text)
      Next
      fput.WriteLine(xmlDoc.selectNodes("//BARCODE/LABEL/MARK_CODE")(0).Text & ",""" & MARK_CODE & """")
      fput.WriteLine(xmlDoc.selectNodes("//BARCODE/LABEL/SGTIN")(0).Text & ",""" & SN & """")
      fput.WriteLine("P1")
      fput.Close
      fso.CopyFile barcode_prn, barcode_port
      fso.DeleteFile barcode_prn, True
  End Select
End Sub

Function Ink_GS(Mark_Code)
  SegmentA = Left(Mark_Code,31)
  SegmentB = Mid(Mark_Code,32,6)
  SegmentC = Mid(Mark_Code,38)
  Ink_GS = SegmentA & chr(29) & SegmentB & chr(29) & SegmentC
End Function

Sub DB_PATH_UPD(DB_PATH)
  Set xmlDocRoot = xmlDoc.selectSingleNode("//BARCODE")
  Set xmlDocDB_PATH = xmlDocRoot.insertBefore(xmlDoc.createElement("DB_PATH"),xmlDoc.selectSingleNode("//BARCODE/PORT"))
  xmlDocDB_PATH.Text = DB_PATH
  xmlDoc.Save(MyBaseFileName & ".xml")
  WScript.Echo "Путь к базе данных установлен по умолчанию " & vbCr & DB_PATH
End Sub

Function DB_PATH_Exist()
  DB_PATH_Exist = ""
  Set xmlDocDB_PATH = xmlDoc.selectNodes("//BARCODE/*")
  For Each n In xmlDocDB_PATH
    If n.NodeName = "DB_PATH" Then DB_PATH_Exist = n.NodeName
  Next
End Function

Function GetMarkCode(SGTIN)
  Set ODBCconnect = CreateObject("ADODB.Connection")
  ODBCconnect.Open "Driver=Firebird/InterBase(r) driver;CHARSET=WIN1251;UID=SYSDBA;PWD=masterkey;DbName=" & DB_PATH
  Set FBrecordset = ODBCconnect.Execute("select first 1 dis.mark_code from docitem_sgtin dis where dis.sgtin='" & SGTIN & "'")
    While Not FBrecordset.EOF
      MARK_CODE = FBrecordset.Fields(0).Value
      FBrecordset.MoveNext
    Wend
  ODBCconnect.Close
  GetMarkCode = MARK_CODE
End Function

Function GetNodeSource(SGTIN)
  Set ODBCconnect = CreateObject("ADODB.Connection")
  ODBCconnect.Open "Driver=Firebird/InterBase(r) driver;CHARSET=WIN1251;UID=SYSDBA;PWD=masterkey;DbName=" & DB_PATH
  Set FBrecordset = ODBCconnect.Execute("select first 1 p.node_id from docitem_sgtin d inner join items i on d.iid=i.iid inner join parties p on i.part_id=p.part_id left join organization o on p.node_id=o.node_id where d.sgtin='" & SGTIN & "' and o.pred_id is null")
    While Not FBrecordset.EOF
      NODE_SOURCE = FBrecordset.Fields(0).Value
      FBrecordset.MoveNext
    Wend
  ODBCconnect.Close
  GetNodeSource = NODE_SOURCE
End Function

Sub MakeXML()
  Set xmlDocRoot = xmlDoc.appendChild(xmlDoc.createProcessingInstruction("xml", "version='1.0' encoding='windows-1251'"))
  Set xmlDocRoot = xmlDoc.appendChild(xmlDoc.createElement("BARCODE"))
  Set xmlDocPort = xmlDocRoot.appendChild(xmlDoc.createElement("PORT"))
  xmlDocPort.Text = "\\zav\zebra"
  Set xmlDocPort = xmlDocRoot.appendChild(xmlDoc.createElement("PORT"))
  xmlDocPort.Text = "gif"
  Set xmlDocPort = xmlDocRoot.appendChild(xmlDoc.createElement("PORT"))
  xmlDocPort.Text = "png"
  Set xmlDocPort = xmlDocRoot.appendChild(xmlDoc.createElement("PORT"))
  xmlDocPort.Text = "jpg"
  Set xmlDocPort = xmlDocRoot.appendChild(xmlDoc.createElement("PORT"))
  xmlDocPort.Text = "svg"
  Set xmlDocLabel = xmlDocRoot.appendChild(xmlDoc.createElement("LABEL"))
  Set xmlDocInit = xmlDocLabel.appendChild(xmlDoc.createElement("INIT"))
  Set xmlDocCommand = xmlDocInit.appendChild(xmlDoc.createElement("COMMAND"))
  xmlDocCommand.Text = "N"
  Set xmlDocCommand = xmlDocInit.appendChild(xmlDoc.createElement("COMMAND"))
  xmlDocCommand.Text = "S4"
  Set xmlDocCommand = xmlDocInit.appendChild(xmlDoc.createElement("COMMAND"))
  xmlDocCommand.Text = "D7"
  Set xmlDocCommand = xmlDocInit.appendChild(xmlDoc.createElement("COMMAND"))
  xmlDocCommand.Text = "I8,10,001"
  Set xmlDocCommand = xmlDocInit.appendChild(xmlDoc.createElement("COMMAND"))
  xmlDocCommand.Text = "N"
  Set xmlDocCommand = xmlDocInit.appendChild(xmlDoc.createElement("COMMAND"))
  xmlDocCommand.Text = "R100,0"
  Set xmlDocMark_Code = xmlDocLabel.appendChild(xmlDoc.createElement("MARK_CODE"))
  xmlDocMark_Code.Text = "b80,0,D,h3"
  Set xmlDocSGTIN = xmlDocLabel.appendChild(xmlDoc.createElement("SGTIN"))
  xmlDocSGTIN.Text = "A70,125,0,1,1,1,N"
  xmlDoc.Save(MyBaseFileName & ".xml")
  WScript.Echo "При первом запуске скрипта был создан файл со стандартными настройками: " & MyBaseFileName & ".xml " & "Запустите скрипт еще раз."
  WScript.Quit
End Sub

Function utf8_win1251(strinput)
  adReadAll = -1
  adTypeText = 2
  Set objStream = CreateObject("ADODB.Stream")    
  objStream.Open()
  objStream.Type = adTypeText
  objStream.Charset = "windows-1251"
  objStream.WriteText(strinput)
  objStream.Flush()
  objStream.Position = 0
  objStream.Charset = "utf-8"
  utf8_win1251 = objStream.ReadText(adReadAll)
  objStream.Close()
End Function

Function win1251_utf8(strinput)
  adReadAll = -1
  adTypeText = 2
  Set objStream = CreateObject("ADODB.Stream")    
  objStream.Open()
  objStream.Type = adTypeText
  objStream.Charset = "utf-8"
  objStream.WriteText(strinput)
  objStream.Flush()
  objStream.Position = 0
  objStream.Charset = "windows-1251"
  win1251_utf8 = objStream.ReadText(adReadAll)
  objStream.Close()
End Function

Function RuEn(InputText)
  RuEn = InputText
  RuEn = Replace(RuEn,"E","Q")
  RuEn = Replace(RuEn,"O","W")
  RuEn = Replace(RuEn,"O","E")
  RuEn = Replace(RuEn,"E","R")
  RuEn = Replace(RuEn,"A","T")
  RuEn = Replace(RuEn,"I","Y")
  RuEn = Replace(RuEn,"A","U")
  RuEn = Replace(RuEn,"O","I")
  RuEn = Replace(RuEn,"U","O")
  RuEn = Replace(RuEn,"C","P")
  RuEn = Replace(RuEn,"O","A")
  RuEn = Replace(RuEn,"U","S")
  RuEn = Replace(RuEn,"A","D")
  RuEn = Replace(RuEn,"A","F")
  RuEn = Replace(RuEn,"I","G")
  RuEn = Replace(RuEn,"?","H")
  RuEn = Replace(RuEn,"I","J")
  RuEn = Replace(RuEn,"E","K")
  RuEn = Replace(RuEn,"A","L")
  RuEn = Replace(RuEn,"?","Z")
  RuEn = Replace(RuEn,"?","X")
  RuEn = Replace(RuEn,"N","C")
  RuEn = Replace(RuEn,"I","V")
  RuEn = Replace(RuEn,"E","B")
  RuEn = Replace(RuEn,"O","N")
  RuEn = Replace(RuEn,"U","M")

  RuEn = Replace(RuEn,"e","q")
  RuEn = Replace(RuEn,"o","w")
  RuEn = Replace(RuEn,"o","e")
  RuEn = Replace(RuEn,"e","r")
  RuEn = Replace(RuEn,"a","t")
  RuEn = Replace(RuEn,"i","y")
  RuEn = Replace(RuEn,"a","u")
  RuEn = Replace(RuEn,"o","i")
  RuEn = Replace(RuEn,"u","o")
  RuEn = Replace(RuEn,"c","p")
  RuEn = Replace(RuEn,"o","a")
  RuEn = Replace(RuEn,"u","s")
  RuEn = Replace(RuEn,"a","d")
  RuEn = Replace(RuEn,"a","f")
  RuEn = Replace(RuEn,"i","g")
  RuEn = Replace(RuEn,"?","h")
  RuEn = Replace(RuEn,"i","j")
  RuEn = Replace(RuEn,"e","k")
  RuEn = Replace(RuEn,"a","l")
  RuEn = Replace(RuEn,"y","z")
  RuEn = Replace(RuEn,"?","x")
  RuEn = Replace(RuEn,"n","c")
  RuEn = Replace(RuEn,"i","v")
  RuEn = Replace(RuEn,"e","b")
  RuEn = Replace(RuEn,"o","n")
  RuEn = Replace(RuEn,"u","m")

End Function

Function win1251_cp866(strinput)
  adReadAll = -1
  adTypeText = 2
  Set objStream = CreateObject("ADODB.Stream")    
  objStream.Open()
  objStream.Type = adTypeText
  objStream.Charset = "cp866"
  objStream.WriteText(strinput)
  objStream.Flush()
  objStream.Position = 0
  objStream.Charset = "windows-1251"
  win1251_cp866 = objStream.ReadText(adReadAll)
  objStream.Close()
End Function
