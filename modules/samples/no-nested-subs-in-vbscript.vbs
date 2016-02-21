Function XRef()
    Function X()
        X = 10
    End Function
    Set XRef = GetRef("X")
End Function
Set Ref = XRef
WScript.Echo Ref()