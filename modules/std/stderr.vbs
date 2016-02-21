Class StdErr

    Private s
    Private hasObj
    Private iswscript

    Private Sub Class_Initialize()
        iswscript = InStr(UCase(WScript.FullName), "WSCRIPT") <> 0
        On Error Resume Next
        Set s = WScript.StdErr
        s.Write ""
        hasObj = (Err.Number = 0)
        On Error GoTo 0
    End Sub

    Sub WriteLine(str)
        If Not hasObj Then Exit Sub
        On Error Resume Next
        s.Write str
        s.WriteBlankLines 1
        hasObj = (Err.Number = 0)
        On Error GoTo 0
    End Sub

    Sub NextLine()
        If Not hasObj Then Exit Sub
        On Error Resume Next
        s.WriteBlankLines 1
        hasObj = (Err.Number = 0)
        On Error GoTo 0
    End Sub

    Sub Write(str)
        If Not hasObj Then Exit Sub
        On Error Resume Next
        s.Write str
        hasObj = (Err.Number = 0)
        On Error GoTo 0
    End Sub
    
    Function CanWrite()
        if Not hasObj Then
            CanWrite = False
        Else
            CanWrite = Not iswscript
        End if
    End Function

End Class

Set Export = New StdErr