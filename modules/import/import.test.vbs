VBImport "import.test.res0.vbs.dat"

''' IMPORT.VBS - MINIFIED VERSION '''
Dim VBImport_Items:Function VBImport(F)
If Not IsObject(VBImport_Items)Then Set VBImport_Items=CreateObject("Scripting.Dictionary")
If VBImport_Items.Exists(F)Then Set VBImport=VBImport_Items.Item(F)Else Set FSO=WScript.CreateObject(_
"Scripting.FileSystemObject"):Set FS=FSO.OpenTextFile(FSO.BuildPath(FSO.GetFile(WScript.ScriptFullName _
).ParentFolder,F),1):Set VBImport=E_ImpVb(FS.ReadAll()):VBImport_Items.Add F,VBImport:FS.Close()
End Function:Function E_ImpVb(S):Set Export=Nothing:Execute S:Set E_ImpVb=Export:End Function
''' END IMPORT.VBS '''
