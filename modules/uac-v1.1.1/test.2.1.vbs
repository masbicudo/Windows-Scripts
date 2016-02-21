VBSLoadRunTime
Set test = VBImport("tester.vbs")

test.DoTest 2, 1

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