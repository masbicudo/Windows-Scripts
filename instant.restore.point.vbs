''''Script to create manual system restore point without user intervention''''

Set objShell = CreateObject("Shell.Application")

If WScript.Arguments.Count = 0 Then
    objShell.ShellExecute "wscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " Run" & Chr(34) & "ManualRestorePoint" & Chr(34), , "runas", 1

ElseIf WScript.Arguments.Count = 1 Then
    objShell.ShellExecute "wscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " Run " & Chr(34) & WScript.Arguments(0) & Chr(34), , "runas", 1

Else
    WScript.Echo WScript.Arguments(1)
    Set objSysRestore = GetObject("winmgmts:\\.\root\default:Systemrestore")
    If (objSysRestore.CreateRestorePoint(WScript.Arguments(1), 0, 100)) = 0 Then
        'objSysRestore.CreateRestorePoint WScript.Arguments(1), 0, 101
        WScript.Echo "Success"
    Else
        WScript.Echo "Failed"
    End If

End If