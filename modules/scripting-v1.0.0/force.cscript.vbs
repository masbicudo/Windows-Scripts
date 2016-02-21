Class ForceCScript

    Sub ForceCScriptExecution(useCmd)
        Dim Arg, Str
        If Not IsCScript Then
            For Each Arg In WScript.Arguments
                If InStr(Arg, " ") Then Arg = """" & Arg & """"
                Str = Str & " " & Arg
            Next
            Set WSHShell = CreateObject("WScript.Shell")
            Command = WScript.Path &_
                "\cscript.exe //nologo """ &_
                WScript.ScriptFullName &_
                """ " + Str
            If useCmd Then
                ExitCode = WSHShell.Run("cmd.exe /k " + Command, 1, True)
                WScript.Quit ExitCode
            Else
                Set Dbg = WSHShell.Exec(command)
                While Dbg.Status = 0
                    WScript.Sleep 1
                Wend
                WScript.Quit Dbg.ExitCode
            End If
        End If
    End Sub

    Function IsCScript
        IsCScript = LCase(Right(WScript.FullName, 12)) = "\cscript.exe"
    End Function
    
End Class

Set Export = New ForceCScript