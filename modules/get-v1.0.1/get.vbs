' Version: 1.0.1

'Set uac = VBImport("uac.vbs")
'Set su = VBImport("force.cscript.vbs")
Set http = VBImport("..\http-v1.0.0\http.vbs")

'uac.Elevate
'su.ForceCScriptExecution 0

Set NamedArgs = WScript.Arguments.Named
If NamedArgs.Exists("list") Then

    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 2
    objStream.CharSet = "utf-8"
    objStream.Open
    objStream.LoadFromFile NamedArgs.Item("list")

    Dim File, Url, DlYet, Ignore, objFSO
    Url = ""
    File = ""
    Set Ignore = CreateObject("Scripting.Dictionary")
    DlYet = True
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set wshShell = CreateObject( "WScript.Shell" )
    Do Until objStream.EOS
        Line = Trim(objStream.ReadText(-2)) ' read line
        If Line = "" And Not DlYet And Url <> "" And File <> "" Then
            DlYet = True
            On Error Resume Next
            objStdOut.WriteLine "Downloading: " & Url
            On Error Goto 0
            result = http.Download(Url, File)
            'WScript.Echo result, Left(result, 6), objFSO.GetExtensionName(result)
            If Left(result, 6) <> "ERROR:" And Ignore.Exists(objFSO.GetExtensionName(result)) Then
                objFSO.DeleteFile result
            End If
        ElseIf InStr(Line, ":") > 0 Then
            DlYet = False
            LineSplit = Split(Line, ":", 2)
            Label = UCase(LineSplit(0))
            If Label = "URL" Then Url = LineSplit(1)
            If Label = "FILE" Then File = LineSplit(1)
            If Label = "FILE-EXPAND" Then File = wshShell.ExpandEnvironmentStrings( LineSplit(1) )
            If Label = "IGNORE" And Not Ignore.Exists(LineSplit(1)) Then Ignore.Add LineSplit(1), True
        End If
    Loop
    If Not DlYet Then
        On Error Resume Next
        objStdOut.WriteLine "Downloading: " & Url
        On Error Goto 0
        result = http.Download(Url, File)
        If Left(result, 6) <> "ERROR:" And Ignore.Exists(objFSO.GetExtensionName(result)) Then
            objFSO.DeleteFile result
        End If
    End If

    objStream.Close

ElseIf NamedArgs.Exists("url") And NamedArgs.Exists("file") Then
    http.Download NamedArgs.Item("url"), NamedArgs.Item("file")
ElseIf WScript.Arguments.Count = 2 Then
    http.Download WScript.Arguments(1), WScript.Arguments(0)
End If

''' IMPORT.VBS - MINIFIED VERSION - v2.0.0 '''
' Author: Miguel Angelo Santos Bicudo
Dim VBImport_Items:Function VBImport(F):A="Scripting."
If Not IsObject(VBImport_Items)Then Set I=CreateObject(A&"Dictionary"):Set VBImport_Items=I
Set W=WScript:Set M=W.CreateObject(A&"FileSystemObject"):Set I=VBImport_Items
If InStr(F,":\")>0Then N=F Else N=M.GetAbsolutePathName(M.BuildPath(M.GetParentFolderName(W.ScriptFullName),F))
If I.Exists(N)Then Set R=I.Item(N)Else If M.FileExists(N)Then Set S=M.OpenTextFile(N,1):Set R=E_ImpVb(S.ReadAll()):_
  S.Close()Else If InStr(N,"\")<>InstrRev(N,"\")Then Set R=VBImport(M.BuildPath("..",F)) Else Set R=Nothing
If Not I.Exists(N)Then I.Add N,R
Set VBImport=R:End Function:Function E_ImpVb(D):Set Export=Nothing:Execute D:Set E_ImpVb=Export:End Function
''' END IMPORT.VBS '''
