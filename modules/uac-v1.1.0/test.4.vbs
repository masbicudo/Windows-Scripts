Set uac = VBImport("uac.vbs")
Set dbg = VBImport("..\debug-v1.0.0\debug.vbs")
dbg.Enable
dbg.SleepUnit = 0
dbg.Echo 1, 0, "Elevated: " & uac.IsElevated

uac.FixCurrentDirectory
uac.Elevate

dbg.Echo 0, 0, "Elevated code running. Waiting 5 seconds."
WScript.Sleep 5000

''' IMPORT.VBS - MINIFIED VERSION '''
Dim VBImport_Items:Function VBImport(F)
If Not IsObject(VBImport_Items)Then Set VBImport_Items=CreateObject("Scripting.Dictionary")
If VBImport_Items.Exists(F)Then Set VBImport=VBImport_Items.Item(F)Else Set FSO=WScript.CreateObject(_
"Scripting.FileSystemObject"):Set FS=FSO.OpenTextFile(FSO.BuildPath(FSO.GetFile(WScript.ScriptFullName _
).ParentFolder,F),1):Set VBImport=E_ImpVb(FS.ReadAll()):VBImport_Items.Add F,VBImport:FS.Close()
End Function:Function E_ImpVb(S):Set Export=Nothing:Execute S:Set E_ImpVb=Export:End Function
''' END IMPORT.VBS '''