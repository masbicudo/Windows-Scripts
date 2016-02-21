' Version: 1.0.3

hasWork = WScript.Arguments.Length <> 0
Set NamedArgs = WScript.Arguments.Named
silent = NamedArgs.Exists("silent") Or Not hasWork

If Not silent Then _
    Set su = VBImport("..\scripting-v1.0.1\force.cscript.vbs")

If hasWork Then Set uac = VBImport("..\uac-v1.0.0\uac.vbs")
Set http = VBImport("..\http-v1.0.0\http.vbs")

If hasWork Then uac.Elevate
If Not silent Then su.ForceCScriptExecution "close-cmd"

Set objFSO = CreateObject("Scripting.FileSystemObject")

listFile = ""
If NamedArgs.Exists("list") Then
    listFile = NamedArgs.Item("list")
ElseIf WScript.Arguments.Length = 1 Then
    arg0 = Trim(WScript.Arguments(0))
    If Right(Left(arg0, 3), 2) = ":\" Or objFSO.FileExists(arg0) Then
        listFile = arg0
    End If
End If

ItDl=0

If listFile <> "" Then
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 2
    objStream.CharSet = "utf-8"
    objStream.Open
    objStream.LoadFromFile listFile

    Dim File, Url, DlYet, Ignore, objFSO
    Url = ""
    File = ""
    Set Ignore = CreateObject("Scripting.Dictionary")
    DlYet = True
    Set wshShell = CreateObject( "WScript.Shell" )
    Do Until objStream.EOS
        Line = Trim(objStream.ReadText(-2)) ' read line
        If Line = "" And Not DlYet And Url <> "" And File <> "" Then
            DlYet = True
            Download Url, File
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
    If Not DlYet Then Download Url, File
    
    objStream.Close
ElseIf NamedArgs.Exists("url") And NamedArgs.Exists("file") Then
    File = NamedArgs.Item("file")
    Url = NamedArgs.Item("url")
    Download Url, File
ElseIf WScript.Arguments.Count = 2 Then
    File = WScript.Arguments(0)
    Url = WScript.Arguments(1)
    Download Url, File
Else
    WScript.Echo "An URI must be indicated, or a list file with multiple items to download."
End If
If Not silent Then WScript.Echo "Finished!"

Sub Download(Url, File)
    If Not silent Then WScript.Echo "Downloading #"&ItDl&": " & Url
    result = http.Download(Url, File)
    ItDl=ItDl+1
    If Left(result, 6) <> "ERROR:" And Ignore.Exists(objFSO.GetExtensionName(result)) Then
        objFSO.DeleteFile result
    End If
End Sub

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
