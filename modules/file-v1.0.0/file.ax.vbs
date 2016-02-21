' Using ActiveX to load and save text and binary files.

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
                .type = 1 '//binary
                .open
                .write ResponseBody
                .savetofile fname, 2 '//overwrite
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

End Class

Set Export = New FileUtils