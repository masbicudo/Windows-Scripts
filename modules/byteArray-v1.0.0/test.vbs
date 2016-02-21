VBSLoadRunTime
Set ba = VBImport("byteArray.vbs")

Redim testArray(255)
For x = 0 To 255
    testArray(x) = x
Next
test = ba.ToByteArray(testArray)
WScript.Echo "VarType(test)", VarType(test)
For x = 0 To 255
    valb = CLng(AscB(MidB(test, x+1, 1)))
    If valb <> x Then WScript.Echo "Falhou!", valb, x Else WScript.Echo "Ok!", valb, x
Next


test = ba.CreateEmptyByteArray(10)
For x = 0 To 9
    valb = CLng(AscB(MidB(test, x+1, 1)))
    If valb <> 0 Then WScript.Echo "Falhou!", valb, x Else WScript.Echo "Ok!", valb, x
Next


src = ba.ToByteArray(Array(0,1,2,3,4,5,6,7,8,9))
dst = ba.ToByteArray(Array(1,5,0,5,0,5,0,6))
result = Array(1,5,3,4,5,6,0,6)
test = ba.CopyBytes(dst, 2, src, 3, 4)
For x = 0 To UBound(test)
    valb = CLng(AscB(MidB(test, x+1, 1)))
    If valb <> result(x) Then WScript.Echo "Falhou!", valb, x Else WScript.Echo "Ok!", valb, x
Next


''' IMPORT RUNTIME.VBS - MINIFIED VERSION - v1.0.0 '''
' Author: Miguel Angelo Santos Bicudo
Sub VBSLoadRunTime():Set W=WScript:F="RunTime.vbs":Set M=W.CreateObject("Scripting.FileSystemObject")
A=Array("","modules\"):G=M.BuildPath(CreateObject("WScript.Shell").ExpandEnvironmentStrings("%WSRTGLB%"),".")&"\"
J=M.BuildPath(M.GetParentFolderName(W.ScriptFullName),".")&"\":P=""
For Z=0 To 99:V=0:If Z=0 And InStr(G,":\")>0 Then V=1
For Y=0 To V:Q=J&P:If Y=1 Then Q=G
For X=0 To 1:N=Q&A(X)&F:If M.FileExists(N) Then:Set S=M.OpenTextFile(N,1):ExecuteGlobal S.ReadAll():S.Close:Exit Sub
Next:Next:B=M.GetAbsolutePathName(J&P&F):P="..\"&P
If InStr(B,"\")=InstrRev(B,"\")Then Exit For
Next:Err.Raise -10937, "VBSLoadRunTime", "Cannot find "&F:End Sub
''' END IMPORT RUNTIME.VBS '''