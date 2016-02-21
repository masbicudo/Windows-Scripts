Dim VBImport_Items
Function VBImport(F)
    Set W=WScript
    T="Scripting."
    Set M=W.CreateObject(T&"FileSystemObject")
    If Not IsObject(VBImport_Items)Then
        Set I=CreateObject(T&"Dictionary")
        Set VBImport_Items=I
    End If
    Set I=VBImport_Items
    A=Array("","modules\")
    G=M.BuildPath(CreateObject("WScript.Shell").ExpandEnvironmentStrings("%WSRTGLB%"),".")&"\"
    J=M.BuildPath(M.GetParentFolderName(W.ScriptFullName),".")&"\"
    P=""
    For Z=0 To 99
        V=0:If Z=0 And InStr(G,":\")>0 Then V=1
        For Y=0 To V
            Q=J&P:If Y=1 Then Q=G
            For X=0 To 1
                U=X=0 And Y=0 And Z=0
                N=Q&A(X)&F
                B=M.GetAbsolutePathName(N)
                If U Then C=B
                If I.Exists(B)Then
                    'W.Echo "found entry", B
                    Set R=I.Item(B)
                    If Not U Then I.Add C,R
                    Set VBImport=R
                    Exit Function
                ElseIf M.FileExists(B)Then
                    'W.Echo "found file", B
                    Set S=M.OpenTextFile(B,1)
                    Set info = New VBLibraryInfoClass
                    info.LibPath = B
                    Set R=Rt_ImpVb(S.ReadAll(), info)
                    Set info = Nothing
                    S.Close
                    I.Add B,R
                    If Not U Then I.Add C,R
                    Set VBImport=R
                    Exit Function
                End If
            Next
        Next
        D=M.GetAbsolutePathName(J&P)
        If Right(D,2)=":\"Then Err.Raise -10937, "VBImport", "Cannot find "&F
        P="..\"&P
    Next
End Function
Function Rt_ImpVb(D, VBImport_LibInfo)
    Set Export=Nothing
    Execute D
    Set Rt_ImpVb=Export
End Function
Class VBLibraryInfoClass
    Public LibPath
    Public Property Get LibDirectory
        pos = InStrRev(LibPath, "\")
        LibDirectory = Mid(LibPath, 1, pos)
    End Property
End Class