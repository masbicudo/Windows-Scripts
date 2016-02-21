Dim VBImport_Items
Function VBImport(From)
    If Not IsObject(VBImport_Items) Then
        Set VBImport_Items = CreateObject("Scripting.Dictionary")
    End If
    If VBImport_Items.Exists(From) Then
        Set VBImport = VBImport_Items.Item(From)
    Else
        Set FSO = WScript.CreateObject("Scripting.FileSystemObject")
        FileName = FSO.BuildPath(FSO.GetFile(WScript.ScriptFullName).ParentFolder,F)
        Set FileStream = FSO.OpenTextFile(FileName, 1) ' ForReading
        Set VBImport = EvalImport(FileStream.ReadAll())
        VBImport_Items.Add From, VBImport
        FileStream.Close()
    End If
End Function
Function EvalImport(Code)
    Set Export = Nothing
    Execute Code
    Set EvalImport = Export
End Function

'''' IMPORT.VBS - MINIFIED VERSION '''
'Dim VBImport_Items:Function VBImport(F)
'If Not IsObject(VBImport_Items)Then Set VBImport_Items=CreateObject("Scripting.Dictionary")
'If VBImport_Items.Exists(F)Then Set VBImport=VBImport_Items.Item(F)Else Set FSO=WScript.CreateObject(_
'"Scripting.FileSystemObject"):Set FS=FSO.OpenTextFile(FSO.BuildPath(FSO.GetFile(WScript.ScriptFullName _
').ParentFolder,F),1):Set VBImport=E_ImpVb(FS.ReadAll()):VBImport_Items.Add F,VBImport:FS.Close()
'End Function:Function E_ImpVb(S):Set Export=Nothing:Execute S:Set E_ImpVb=Export:End Function
'''' END IMPORT.VBS '''
