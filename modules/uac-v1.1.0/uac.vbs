Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

Class UAC

    Private signalFileStream

    Private m_ElevateCalled
    Sub Elevate()
        DebugEnter "Elevate"
        If Not m_ElevateCalled Then
            Set OSList = GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem")
            For Each OS In OSList
                If InStr(1, OS.Caption, "XP") = 0 And InStr(1, OS.Caption, "Server 2003") = 0 Then
                    ElevateCore(False)
                End If
            Next
        End If
        DebugExit "Elevate"
    End Sub

    Sub ElevateWait()
        DebugEnter "ElevateWait"
        If Not m_ElevateCalled Then
            Set OSList = GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem")
            For Each OS In OSList
                If InStr(1, OS.Caption, "XP") = 0 And InStr(1, OS.Caption, "Server 2003") = 0 Then
                    ElevateCore(True)
                End If
            Next
        End If
        DebugExit "ElevateWait"
    End Sub

    Private Sub ElevateCore(wait)
        DebugEnterArgs "ElevateCore", "wait = "&wait
        m_ElevateCalled = True
        Set WSHShell = CreateObject("WScript.Shell")
        Set FSO = CreateObject("Scripting.FileSystemObject")

        'Signaling the execution of the elevated script
        SignalExecution()

        mustRestart = FixCurrentDirectory()
        DebugEcho 1, 0, "mustRestart" &" "& mustRestart
        Dim AllArgs
        If mustRestart Then
            ItA = 0
            For Each Arg In WScript.Arguments
                DebugEcho 0, 1, "Arg #" & ItA &" "& Arg
                If Not IsArgName(Arg, "UAC-Fix-CurrentDirectory") And Not IsArgName(Arg, "UAC-Signal-File-GUID") Then
                WScript.Echo Arg
                    If InStr(Arg, " ") Then Arg = """" & Arg & """"
                    AllArgs = AllArgs & " " & Arg
                End If
                ItA = ItA+1
            Next
            DebugEcho 1, 1, AllArgs
            Command = WScript.FullName & " """ & WScript.ScriptFullName & """" & AllArgs
            ExitCode = WSHShell.Run("cmd.exe /c " + Command, 1, True)
            DebugEcho 1, 0, "Quit: uac.vbs"
            WScript.Quit ExitCode
        ElseIf Not IsElevated Then
            For Each Arg In WScript.Arguments
                If InStr( Arg, " " ) Then Arg = """" & Arg & """"
                AllArgs = AllArgs & " " & Arg
            Next
            FixArg = "/UAC-Fix-CurrentDirectory:"""&WSHShell.CurrentDirectory&""""
            Command = """" & WScript.ScriptFullName & """" & AllArgs & " " & FixArg

            'Creating a temp file that will be used for start and exit signals
            Dim guid
            If wait Then
                guid = CreateSignal()
                If guid = "" Then DebugEcho 1, 0, "Could not create file for signals": Err.Raise -10000
                Command = Command & " /UAC-Signal-File-GUID:" & guid
            End If

            'Recall the script as administrator
            DebugEcho 1, 0, "Recalling the script with 'runas' verb"
            CreateObject("Shell.Application").ShellExecute _
                FSO.GetFileName(WScript.FullName),_
                " //nologo " & Command,_
                FSO.GetFile(WScript.ScriptFullName).ParentFolder,_
                "runas",_
                1

            'Waiting for execution to finish
            If wait Then If Not WaitForExit(guid) Then DebugEcho 1, 0, "Elevation canceled by user."

            DebugEcho 1, 0, "Quit: uac.vbs"
            WScript.Quit
        End If
        DebugExit "ElevateCore"
    End Sub

    Function IsElevated
        IsElevated = CreateObject("WScript.Shell").Run("cmd.exe /c ""whoami /groups|findstr S-1-16-12288""", 0, true) = 0
    End function

    Private Function IsArgName(Arg, name)
        leftArg = "/"&name&":"
        IsArgName = (Left(Arg, Len(leftArg)) = leftArg)
    End Function

    Private Sub SignalExecution()
        DebugEnter "SignalExecution"
        Set WSHShell = CreateObject("WScript.Shell")
        Set FSO = CreateObject("Scripting.FileSystemObject")
        If WScript.Arguments.Named.Exists("UAC-Signal-File-GUID") Then
            guid = WScript.Arguments.Named.Item("UAC-Signal-File-GUID")
            DebugEcho 1,0,"guid = "&guid
            signalFile = WSHShell.ExpandEnvironmentStrings("%TEMP%\"&guid)
            DebugEcho 1,0,"signalFile = "&signalFile
            Set signalFileStream = FSO.OpenTextFile(signalFile, ForWriting, False)
            signalFileStream.Write "STARTED"
            'MUST NOT CLOSE THE STREAM
        End If
        DebugExit "SignalExecution"
    End Sub

    Private Function CreateSignal()
        DebugEnter "CreateSignal"
        Set WSHShell = CreateObject("WScript.Shell")
        Set FSO = CreateObject("Scripting.FileSystemObject")
        errNum = False
        signalFile = ""
        guid = ""
        For rep = 0 to 9
            guid = CreateGUID
            DebugEcho 1, 0, "guid = CreateGUID ' "&guid
            signalFile = WSHShell.ExpandEnvironmentStrings("%TEMP%\"&guid)
            DebugEcho 1, 0, "signalFile = "&signalFile
            On Error Resume Next
            FSO.CreateTextFile(signalFile).Close()
            errNum = Err.Number
            errMessage = Err.Description
            On Error GoTo 0
            If errNum = 0 Then Exit For
            DebugEcho 1, 0, "errNum" &" = "&errNum &" ' "&errMessage
            WScript.Sleep 500
        Next
        If FSO.FileExists(signalFile) Then CreateSignal = guid Else CreateSignal = ""
        DebugExit "CreateSignal"
    End Function

    Private Function WaitForExit(guid)
        Set WSHShell = CreateObject("WScript.Shell")
        Set FSO = CreateObject("Scripting.FileSystemObject")
        DebugEnterArgs "WaitForExit", "guid = "&guid
        signalFile = WSHShell.ExpandEnvironmentStrings("%TEMP%\"&guid)

        'Waiting for start signal
        DebugEnter "Read Start Msg"
        text = ""
        For X = 1 To 8
            On Error Resume Next
            Set stream = CreateObject("Scripting.FileSystemObject")_
                .OpenTextFile(signalFile, ForReading)
            text = stream.ReadAll()
            stream.Close()
            On Error GoTo 0
            'DebugEcho 1,0,"Try #"&X
            If text = "STARTED" Then Exit For
            WScript.Sleep X*50
        Next
        DebugExit "Read Start Msg"

        If text = "" Then
            WaitForExit = False
        Else
            DebugEnter "Read Exit Msg"
            'Waiting for exit signal
            X = 0
            Do While FSO.FileExists(signalFile)
                On Error Resume Next
                FSO.DeleteFile(signalFile)
                On Error GoTo 0
                'DebugEcho 1,0,"Try #"&X
                WScript.Sleep 50
                X=X+1
            Loop
            WaitForExit = True
            DebugExit "Read Exit Msg"
        End If
        DebugExitResult "WaitForExit", WaitForExit
    End Function

    Private m_FixedCurDir
    Function FixCurrentDirectory()
        DebugEnter "FixCurrentDirectory"
        If Not m_FixedCurDir Then
            m_FixedCurDir = True
            Set FSO = CreateObject("Scripting.FileSystemObject")
            Set WSHShell = CreateObject("WScript.Shell")
            Set Nargs = WScript.Arguments.Named
            FixCurrentDirectory = False
            DebugEcho 0, 0, "Nargs.Exists('UAC-Fix-CurrentDirectory')" &" "& Nargs.Exists("UAC-Fix-CurrentDirectory") &" "& Nargs.Item("UAC-Fix-CurrentDirectory")
            DebugEcho 1, 0, "FSO.FileExists('cscript.exe')" &" "& FSO.FileExists("cscript.exe")
            If Nargs.Exists("UAC-Fix-CurrentDirectory") Then
                WSHShell.CurrentDirectory = Nargs.Item("UAC-Fix-CurrentDirectory")
                FixCurrentDirectory = True
            ElseIf FSO.FileExists("cscript.exe") Then
                WSHShell.CurrentDirectory = FSO.GetFile(WScript.ScriptFullName).ParentFolder
            End If
        End If
        DebugExitResult "FixCurrentDirectory", FixCurrentDirectory
    End Function

    Private Function IsCScript
        IsCScript = LCase(Right(WScript.FullName, 12)) = "\cscript.exe"
    End Function

    Private Function CreateGUID
      Dim TypeLib
      Set TypeLib = CreateObject("Scriptlet.TypeLib")
      CreateGUID = Replace(Mid(TypeLib.Guid, 2, 36), "-", "")
    End Function

    Private dbg
    Private Sub Class_Initialize()
        On Error Resume Next
        Set dbg = VBImport("..\debug-v1.0.0\debug.vbs")
        On Error GoTo 0
        m_FixedCurDir = False
        m_ElevateCalled = False
    End Sub

    Private Function DebugTime(t)
        DebugTime = t
        If VarType(dbg) <> 0 Then DebugTime = t*dbg.SleepUnit + dbg.SleepBase
    End Function
    Private Sub DebugEcho(t,i,m)
        If VarType(dbg) <> 0 Then dbg.Echo t,i,m
    End Sub
    Private Sub DebugEnter(n)
        If VarType(dbg) <> 0 Then dbg.EnterBlock n
    End Sub
    Private Sub DebugExit(n)
        If VarType(dbg) <> 0 Then dbg.ExitBlock n
    End Sub
    Private Sub DebugEnterArgs(n,m)
        If VarType(dbg) <> 0 Then dbg.EnterBlockArgs n,m
    End Sub
    Private Sub DebugExitResult(n,m)
        If VarType(dbg) <> 0 Then dbg.ExitBlockResult n,m
    End Sub

End Class

Set Export = New UAC