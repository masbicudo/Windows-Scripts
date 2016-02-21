Const adTypeBinary = 1
Const adTypeText = 2

allCharSets = Array( _
    "us-ascii"    ,_
    "utf-7"       ,_
    "utf-8"       ,_
    "iso-8859-1"  ,_
    "iso-8859-2"  ,_
    "iso-8859-3"  ,_
    "iso-8859-4"  ,_
    "iso-8859-5"  ,_
    "iso-8859-6"  ,_
    "iso-8859-7"  ,_
    "iso-8859-8"  ,_
    "iso-8859-9"  ,_
    "koi8-r"      ,_
    "shift-jis"   ,_
    "big5"        ,_
    "euc-jp"      ,_
    "euc-kr"      ,_
    "gb2312"      ,_
    "iso-2022-jp" ,_
    "iso-2022-kr" )

Private Function BuildByteTable()
    Redim result(255)
    Set BinaryStream = CreateObject("ADODB.Stream")
    BinaryStream.Open
    For l = 0 To UBound(allCharSets)
        charset = allCharSets(l)
        For x = 0 To 255
            BinaryStream.Position = 0
            BinaryStream.SetEOS
            BinaryStream.Type = adTypeText
            BinaryStream.CharSet = charset
            BinaryStream.WriteText Chr(x)
            BinaryStream.Position = 0
            BinaryStream.Type = 1
            byteArray = BinaryStream.Read
            For y = 0 To UBound(byteArray)
                valb = CLng(AscB(MidB(byteArray, y+1, 1)))
                If VarType(result(valb)) = 0 Then
                    BinaryStream.Position = y
                    singleByteArray = BinaryStream.Read(1)
                    result(valb) = singleByteArray
                End If
            Next
        Next
    Next
    BuildByteTable = result
    BinaryStream.Close()
End Function

Private Sub SaveTable()
    table = BuildByteTable()
    Set BinaryStream = CreateObject("ADODB.Stream")
    BinaryStream.Open
    BinaryStream.Type = 1
    For x = 0 To 255
        BinaryStream.Write table(x)
    Next
    BinaryStream.SaveToFile "ByteTable"
    BinaryStream.Close()
End Sub

SaveTable()