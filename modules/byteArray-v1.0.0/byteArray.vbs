' Version: 1.0.0

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

tableFile = VBImport_LibInfo.LibDirectory & "ByteTable"

Class ByteArrayClass

    Private Function BuildByteTable()
        Redim result(255)
        Set stream = CreateObject("ADODB.Stream")
        stream.Open
        For l = 0 To UBound(allCharSets)
            charset = allCharSets(l)
            For x = 0 To 255
                stream.Position = 0
                stream.SetEOS
                stream.Type = adTypeText
                stream.CharSet = charset
                stream.WriteText Chr(x)
                stream.Position = 0
                stream.Type = 1
                byteArray = stream.Read
                For y = 0 To UBound(byteArray)
                    valb = CLng(AscB(MidB(byteArray, y+1, 1)))
                    If VarType(result(valb)) = 0 Then
                        stream.Position = y
                        singleByteArray = stream.Read(1)
                        result(valb) = singleByteArray
                    End If
                Next
            Next
        Next
        BuildByteTable = result
        stream.Close()
    End Function

    Function GetByteTable
        On Error Resume Next
        GetByteTable = OpenTable()
        hasError = Err.Number <> 0
        On Error GoTo 0
        If hasError Then SaveTable(): GetByteTable = OpenTable()
    End Function

    Function ToByteArray(variantArray)
        table = GetByteTable()
        Set stream = CreateObject("ADODB.Stream")
        stream.Open
        stream.Type = 1
        For x = 0 To UBound(variantArray)
            stream.Write table(variantArray(x))
        Next
        stream.Position = 0
        ToByteArray = stream.Read
        stream.Close()
    End Function

    Function CreateEmptyByteArray(Size)
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = adTypeText
        stream.Open
        stream.CharSet = "ASCII"
        stream.WriteText String(Size, 0)
        stream.Position = 0
        stream.Type = adTypeBinary
        CreateEmptyByteArray = stream.Read
        stream.Close()
    End Function

    Function CopyBytes(destBuffer, destPos, srcBuffer, srcPos, size)
        If size > 0 Then
            Set stream = CreateObject("ADODB.Stream")
            stream.Type = adTypeBinary
            stream.Open
            stream.Write srcBuffer
            stream.Position = srcPos
            srcCrop = stream.Read(size)

            stream.Position = 0
            stream.Write destBuffer
            stream.Position = destPos
            stream.Write srcCrop
            stream.Position = 0
            CopyBytes = stream.Read(UBound(destBuffer)+1)
            stream.Close()
        Else
            CopyBytes = destBuffer
        End If
    End Function

    Private Sub SaveTable()
        table = BuildByteTable()
        Set stream = CreateObject("ADODB.Stream")
        stream.Open
        stream.Type = 1
        For x = 0 To 255
            stream.Write table(x)
        Next
        stream.SaveToFile tableFile
        stream.Close()
    End Sub

    Private Function OpenTable()
        Redim result(255)
        Set stream = CreateObject("ADODB.Stream")
        stream.Open
        stream.Type = 1
        stream.LoadFromFile tableFile
        For x = 0 To 255
            singleByteArray = stream.Read(1)
            result(x) = singleByteArray
        Next
        OpenTable = result
        stream.Close()
    End Function

End Class

Set Export = New ByteArrayClass
src = Export.ToByteArray(Array(0,1,2,3,4,5,6,7,8,9))
