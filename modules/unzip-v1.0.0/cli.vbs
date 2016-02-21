' Version: 1.0.0

Set unzip = VBImport("unzip.vbs")

src = ""
dest = ""

Set NamedArgs = WScript.Arguments.Named
If NamedArgs.Exists("src") Then
    src = NamedArgs.Item("src")
ElseIf WScript.Arguments.Length > 0 Then
    src = WScript.Arguments(0)
End If

If NamedArgs.Exists("dest") Then
    dest = NamedArgs.Item("dest")
ElseIf WScript.Arguments.Length > 1 Then
    dest = WScript.Arguments(1)
End If

If src <> "" Then
    unzip.Unzip src, dest
Else
    WScript.Echo "You must indicate a source Zip file."
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

