Class ForceCScript

    Private vrb
    Private Sub Class_Initialize()
        vrb = False
    End Sub

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
                """" + Str
            If useCmd = "close-cmd" Then
                If vrb Then WScript.Echo "Command:", "cmd.exe /c " + Command
                ExitCode = WSHShell.Run("cmd.exe /c " + Command, 1, True)
                'WScript.Echo "Saiu do force.vbs"
                WScript.Quit ExitCode
            ElseIf useCmd = "keep-cmd" Then
                If vrb Then WScript.Echo "Command:", "cmd.exe /k " + Command
                ExitCode = WSHShell.Run("cmd.exe /k " + Command, 1, True)
                'WScript.Echo "Saiu do force.vbs"
                WScript.Quit ExitCode
            ElseIf useCmd = "no-cmd" Then
                If vrb Then WScript.Echo "Command:", Command
                Set Dbg = WSHShell.Exec(command)
                While Dbg.Status = 0
                    WScript.Sleep 1
                Wend
                'WScript.Echo "Saiu do force.vbs"
                WScript.Quit Dbg.ExitCode
            End If
        End If
    End Sub

    Function IsCScript
        IsCScript = LCase(Right(WScript.FullName, 12)) = "\cscript.exe"
    End Function
    
    Sub Verbose()
        vrb = True
    End Sub
    
    Sub Silent()
        vrb = False
    End Sub
    
End Class

Set Export = New ForceCScript