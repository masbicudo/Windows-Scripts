Class UAC

    Sub Elevate()
        Set OSList = GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem")
        For Each OS In OSList
            If InStr(1, OS.Caption, "XP") = 0 And InStr(1, OS.Caption, "Server 2003") = 0 Then
                Set WSHShell = CreateObject("WScript.Shell")
                Set FSO = CreateObject("Scripting.FileSystemObject")
                If IsElevated Then
                    WSHShell.CurrentDirectory = FSO.GetFile(WScript.ScriptFullName).ParentFolder
                Else
                    Dim AllArgs
                    For Each Arg In WScript.Arguments
                        If InStr( Arg, " " ) Then Arg = """" & Arg & """"
                        AllArgs = AllArgs & " " & Arg
                    Next
                    Command = """" & WScript.ScriptFullName & """" & AllArgs
                    CreateObject("Shell.Application").ShellExecute _
                        FSO.GetFileName(WScript.FullName),_
                        " //nologo " & Command,_
                        FSO.GetFile(WScript.ScriptFullName).ParentFolder,_
                        "runas",_
                        1
                    WScript.Quit
                End If
            End If
        Next
    End Sub
    
    Function IsElevated
        IsElevated = CreateObject("WScript.Shell").Run("cmd.exe /c ""whoami /groups|findstr S-1-16-12288""", 0, true) = 0
    End function

End Class

Set Export = New UAC