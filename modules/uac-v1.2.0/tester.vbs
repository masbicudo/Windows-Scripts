Class TestParams
    Sub DoTest(a,b)
        Set uac = VBImport("uac.vbs")
        Set dbg = VBImport("..\debug-v1.1.0\debug.vbs")
        Set FSO = CreateObject("Scripting.FileSystemObject")
        dbg.Enable
        dbg.SleepUnit = 0
        fileName = FSO.GetFileName(WScript.ScriptFullName)
        dbg.EnterBlock fileName
        dbg.Echo 1, 0, "Elevated: " & uac.IsElevated

        If b = 1 Then NoOp
        If b = 2 Then uac.FixCurrentDirectory
        If b = 3 Then uac.ResetCurrentDirectory

        If a = 1 Then uac.Elevate
        If a = 2 Then uac.ElevateWait
        If a = 3 Then uac.TryElevate

        timeToWait = 0
        dbg.Echo 0, 0, "Elevated code running. Waiting "&timeToWait&"s."
        WScript.Sleep timeToWait*1000
        dbg.ExitBlock fileName
    End Sub
    Sub NoOp
    End Sub
End Class

Set Export = New TestParams