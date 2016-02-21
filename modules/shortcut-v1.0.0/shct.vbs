' Version: 1.0.0
VBSLoadRunTime
Set dbg = VBImport("..\debug-v1.2.0\debug.vbs")
Set su = VBImport("..\scripting-v1.0.1\force.cscript.vbs")
Set uac = VBImport("..\uac-v1.2.0\uac.vbs")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set oWS = WScript.CreateObject("WScript.Shell")

'WScript.Echo objFSO.GetAbsolutePathName("C:")
'WScript.Echo objFSO.GetAbsolutePathName("C:\")
'WScript.Echo objFSO.GetAbsolutePathName("C:\abc")
'WScript.Echo objFSO.GetAbsolutePathName("C:\abc\")
'WScript.Echo objFSO.GetFileName("C:")
'WScript.Echo objFSO.GetFileName("C:\")
'WScript.Echo objFSO.GetFileName("C:\abc")
'WScript.Echo objFSO.GetFileName("C:\abc\")
'WScript.Echo "GetDriveName", objFSO.GetDriveName("C:")
'WScript.Echo "GetDriveName", objFSO.GetDriveName("C:\")
'WScript.Echo "GetDriveName", objFSO.GetDriveName("C:\abc")
'WScript.Echo "GetDriveName", objFSO.GetDriveName("C:\abc\")
'WScript.Echo "GetDriveName", objFSO.FolderExists("C:")
'WScript.Echo "GetDriveName", objFSO.FolderExists("C:\")
'WScript.Echo "GetDriveName", objFSO.FolderExists("C:\Windows")
'WScript.Echo "GetDriveName", objFSO.FolderExists("C:\Windows\")

dbg.EnterBlock "shct.vbs"

silent = True
Set NamedArgs = WScript.Arguments.Named

'FORCE CSCRIPT
If Not silent Then su.ForceCScriptExecution "close-cmd"

ItDl=0
prefix="Creating shortcut"
If NamedArgs.Exists("target") And NamedArgs.Exists("link") Then
    target = NamedArgs.Item("target")
    link = NamedArgs.Item("link")
    CreateShortcut target, link
'ElseIf WScript.Arguments.Count = 2 Then
'    arg0 = WScript.Arguments(0)
'    arg1 = WScript.Arguments(1)
'    isFile0 = objFSO.FileExists(arg0)
'    isFile1 = objFSO.FileExists(arg1)
'    isDir0 = objFSO.FolderExists(arg0)
'    isDir1 = objFSO.FolderExists(arg1)
'    ext0 = LCase(objFSO.GetExtensionName(arg0))
'    ext1 = LCase(objFSO.GetExtensionName(arg1))
'    lnk0 = (ext0 = "lnk")
'    lnk1 = (ext1 = "lnk")
'    If isFile0 Then link = arg1: target = arg0
'    If isFile1 Then link = arg0: target = arg1
'    If ext0 = "lnk" Then link = arg0: target = arg1
'    If ext1 = "lnk" Then link = arg1: target = arg0
'    Download Url, File
Else
    WScript.Echo "Usage:" & vbCrLf & "shct /target:target-object [/link:(lnk-file|destination-folder)]"
End If
If Not silent Then WScript.Echo "Finished!"

Sub CreateShortcut(target, link)
WScript.Echo target, link
    link = Trim(link)
    extLink = LCase(objFSO.GetExtensionName(link))
    target = objFSO.GetAbsolutePathName(target)
    If extLink <> "lnk" Then
        If objFSO.FolderExists(link) Or Right(link, 1) = "\" Then
            If Right(link, 1) <> "\" Then link = link & "\"
            proposedName = objFSO.GetFileName(target)
            If proposedName = "" Then proposedName = objFSO.GetDriveName(target)
            If Right(proposedName, 1) = ":" Then proposedName = Mid(proposedName, 1, Len(proposedName) - 1)
            link = link & proposedName & ".lnk"
        Else
            link = Mid(link, 1, Len(link) - Len(extLink))
            If Right(link, 1) <> "." Then link = link & "."
            link = link & "lnk"
        End If
    End If
    link = objFSO.GetAbsolutePathName(link)

    If Not silent Then WScript.Echo Replace(prefix,"#","#"&ItDl)&": " & Url
    WScript.Echo Replace(prefix,"#","#"&ItDl)&": " & link & " => " & target
    Set oLink = oWS.CreateShortcut(link)
    oLink.TargetPath = target
    oLink.Save
End Sub

dbg.ExitBlock "shct.vbs"

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