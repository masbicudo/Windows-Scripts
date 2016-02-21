' Using ActiveX to load and save text and binary files.
Const adTypeBinary = 1
Const adTypeText = 2
Const adSaveCreateOverWrite = 2

Class FileUtils

    Sub SaveFile(fname, ResponseBody)
        On Error Resume Next
        Set objADOStream = CreateObject("Adodb.Stream")
        If objADOStream Is Nothing Then
            Set objFile = objFSO.OpenTextFile(fname, ForWriting, True)
            For i = 1 To LenB(ResponseBody)
                If Err.Number = 0 Then
                    objFile.Write Chr(AscB(MidB(ResponseBody, i, 1)))
                End If
            Next
            objFile.Close
        Else
            with objADOStream
                .type = adTypeBinary '//binary
                .open
                .write ResponseBody
                .savetofile fname, adSaveCreateOverWrite '//overwrite
            end with
        End If
        If Err.Number <> 0 Then
            Err.Clear
            WScript.Echo "Could not write to file"
        End If
        On Error GoTo 0
    End Sub

    Function RemoveInvalidChars(str)
        str = Replace(str, "\", "")
        str = Replace(str, "/", "")
        str = Replace(str, ":", "")
        str = Replace(str, "*", "")
        str = Replace(str, "?", "")
        str = Replace(str, """", "")
        str = Replace(str, "<", "")
        str = Replace(str, ">", "")
        str = Replace(str, "|", "")
        RemoveInvalidChars = str
    End Function

    Function SaveBinaryData(fileName, byteArray)
      Set stream = CreateObject("ADODB.Stream")
      stream.Type = adTypeBinary
      stream.Open
      stream.Write byteArray
      stream.SaveToFile fileName, adSaveCreateOverWrite
      stream.Close
    End Function

    Function SaveTextData(fileName, text, charSet)
      Set stream = CreateObject("ADODB.Stream")
      stream.Type = adTypeText
      If Len(charSet) > 0 Then stream.CharSet = charSet
      stream.Open
      stream.WriteText text
      stream.SaveToFile fileName, adSaveCreateOverWrite
      stream.Close
    End Function

    Function ReadBinaryFile(fileName)
      Set stream = CreateObject("ADODB.Stream")
      stream.Type = adTypeBinary
      stream.Open
      stream.LoadFromFile fileName
      ReadBinaryFile = stream.Read
      stream.Close
    End Function

    Function ReadTextFile(fileName, charSet)
      Set stream = CreateObject("ADODB.Stream")
      stream.Type = adTypeText
      If Len(charSet) > 0 Then stream.CharSet = charSet
      stream.Open
      stream.LoadFromFile fileName
      ReadTextFile = stream.ReadText
      stream.Close
    End Function

End Class

Set Export = New FileUtils