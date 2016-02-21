VBSLoadRunTime
Set mime = VBImport("mime.vbs")

WScript.Echo "Testing relations between mimes and extensions"

    Test_IsExtensionOfMimeType "exe", "application/octet-stream"
    Test_IsExtensionOfMimeType "bin", "application/octet-stream"
    Test_IsExtensionOfMimeType "jpg", "application/octet-stream"
    Test_IsExtensionOfMimeType "jpg", "image/jpeg"

WScript.Echo ""
WScript.Echo "Testing relations between mimes"

    Test_GetMimesInGraphOfMime "image/jpeg"
    Test_GetMimesInGraphOfMime "text/plain"
    Test_GetMimesInGraphOfMime "application/octet-stream"

    allMimes = mime.GetAllMimeTypes()
    For it = 0 To UBound(allMimes)
        eachMime = allMimes(it)
        relatedMimes = mime.GetMimesInGraphOfMime(eachMime)
        If UBound(relatedMimes) > 0 Then
            Test_GetMimesInGraphOfMime eachMime
        End If
    Next

WScript.Echo ""
WScript.Echo "Testing relations between extensions"

    Test_GetExtsInGraphOfMime "image/jpeg"
    Test_GetExtsInGraphOfMime "text/plain"
    Test_GetExtsInGraphOfMime "application/octet-stream"



Sub Test_IsExtensionOfMimeType(a, b)
    result = "F"
    If mime.IsExtensionOfMimeType(a, b) Then result = "T"
    WScript.Echo "  " & a & " is " & b & ": " & result
End Sub
Sub Test_GetMimesInGraphOfMime(a)
    mimes = mime.GetMimesInGraphOfMime(a)
    WScript.Echo "  " & a & " => " & Join(mimes, " ")
End Sub
Sub Test_GetExtsInGraphOfMime(a)
    exts = mime.GetExtsInGraphOfMime(a)
    WScript.Echo "  " & a & " => " & Join(exts, " ")
End Sub

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