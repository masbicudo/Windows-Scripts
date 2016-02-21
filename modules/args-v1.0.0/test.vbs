Set args = VBImport("args.vbs")

WScript.Echo args.CommandLineArgs()

WScript.Echo "All args"
it = 0
For Each Arg In WScript.Arguments
    WScript.Echo "  #"&it&": "&args.ToCommandLine(Arg)
    it=it+1
Next

WScript.Echo "All names"
it = 0
For Each Arg In WScript.Arguments
    WScript.Echo "  #"&it&": "&args.GetArgName(Arg)
    it=it+1
Next

WScript.Echo "All values"
it = 0
For Each Arg In WScript.Arguments
    WScript.Echo "  #"&it&": "&args.GetArgValue(Arg)
    it=it+1
Next

WScript.Echo "IsNamed"
it = 0
For Each Arg In WScript.Arguments
    WScript.Echo "  #"&it&": "&-CInt(args.IsNamed(Arg))
    it=it+1
Next

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
