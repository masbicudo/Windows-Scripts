' Version: 1.2.0
On Error Resume Next
Set dbg = VBImport("..\debug-v1.2.0\debug.vbs")
On Error GoTo 0

Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
selfTest = False
Class UAC

    Private signalFileStream

    Private m_ElevateCalled
    Function Elevate()
        DebugEnter "UAC:Elevate"
        Elevate = False
        If Not m_ElevateCalled Then
            Set OSList = GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem")
            For Each OS In OSList
                If InStr(1, OS.Caption, "XP") = 0 And InStr(1, OS.Caption, "Server 2003") = 0 Then
                    ElevateCore False, False, False
                    Elevate = True
                    Exit For
                End If
            Next
        End If
        DebugExit "UAC:Elevate"
    End Function

    Function ElevateWait()
        DebugEnter "UAC:ElevateWait"
        ElevateWait = False
        If Not m_ElevateCalled Then
            Set OSList = GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem")
            For Each OS In OSList
                If InStr(1, OS.Caption, "XP") = 0 And InStr(1, OS.Caption, "Server 2003") = 0 Then
                    ElevateCore True, False, False
                    ElevateWait = True
                    Exit For
                End If
            Next
        End If
        DebugExit "UAC:ElevateWait"
    End Function

    Function ElevateWithOutput()
        DebugEnter "UAC:ElevateWithOutput"
        ElevateWithOutput = False
        If Not m_ElevateCalled Then
            Set OSList = GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem")
            For Each OS In OSList
                If InStr(1, OS.Caption, "XP") = 0 And InStr(1, OS.Caption, "Server 2003") = 0 Then
                    ElevateCore True, False, True
                    ElevateWithOutput = True
                    Exit For
                End If
            Next
        End If
        DebugExit "UAC:ElevateWithOutput"
    End Function

    Function TryElevate()
        DebugEnter "UAC:TryElevate"
        TryElevate = False
        If Not m_ElevateCalled Then
            Set OSList = GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem")
            For Each OS In OSList
                If InStr(1, OS.Caption, "XP") = 0 And InStr(1, OS.Caption, "Server 2003") = 0 Then
                    ElevateCore True, True, False
                    TryElevate = True
                    Exit For
                End If
            Next
        End If
        DebugExit "UAC:TryElevate"
    End Function

    Function TryElevateWithOutput()
        DebugEnter "UAC:TryElevateWithOutput"
        TryElevateWithOutput = False
        If Not m_ElevateCalled Then
            Set OSList = GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem")
            For Each OS In OSList
                If InStr(1, OS.Caption, "XP") = 0 And InStr(1, OS.Caption, "Server 2003") = 0 Then
                    ElevateCore True, True, True
                    TryElevateWithOutput = True
                    Exit For
                End If
            Next
        End If
        DebugExit "UAC:TryElevateWithOutput"
    End Function

    Private Sub ElevateCore(wait, allowNonElevated, redirectOutput)
        If Not wait And redirectOutput Then Err.Raise -10264, "uac.vbs", "Cannot redirect output without waiting."

        DebugEnterArgs "ElevateCore", "wait = "&wait&", allowNonElevated = "&allowNonElevated&", redirectOutput = "&redirectOutput
        m_ElevateCalled = True
        Set WSHShell = CreateObject("WScript.Shell")
        Set FSO = CreateObject("Scripting.FileSystemObject")

        'Signaling the execution of the elevated script
        SignalExecution()

        'Fix the current directory if not fixed yet
        fixed = FixCurrentDirectory()

        Set namedArgs = WScript.Arguments.Named
        hasUacArgs = namedArgs.Exists("UAC-Fix-CurrentDirectory") _
            Or namedArgs.Exists("UAC-Signal-File-GUID")
        mustRestartWithoutUacArgs = fixed And hasUacArgs
        DebugEcho 1, 0, "mustRestartWithoutUacArgs" &" "& mustRestartWithoutUacArgs

        'Output file
        DebugEcho 0, 0, "redirectOutput = " & redirectOutput
        OutFile = ""
        If redirectOutput And namedArgs.Exists("UAC-Signal-File-GUID") Then
            OutFile = namedArgs.Item("UAC-Signal-File-GUID") & ".txt"
            OutFile = WSHShell.ExpandEnvironmentStrings("%TEMP%\" & OutFile)
            DebugEcho 0, 0, "OutFile = " & OutFile
        End If
        mustRestartWithOutputToFile = (OutFile <> "")

        'Restarting the script without the special UAC arguments
        Dim AllArgs
        If mustRestartWithoutUacArgs Or mustRestartWithOutputToFile Then
            ItA = 0
            For Each Arg In WScript.Arguments
                isUacGuidArg = IsArgName(Arg, "UAC-Signal-File-GUID")
                isUacFixArg = IsArgName(Arg, "UAC-Fix-CurrentDirectory")
                willUseArg = (Not mustRestartWithoutUacArgs Or Not isUacFixArg) And Not isUacGuidArg
                If willUseArg Then DebugEcho 0, 1, "Arg #" & ItA &" "& Arg _
                    Else DebugEcho 0, 1, "Arg (Discarded) #" & ItA &" "& Arg
                If willUseArg Then
                    If InStr(Arg, " ") Then Arg = """" & Arg & """"
                    AllArgs = AllArgs & " " & Arg
                End If
                ItA = ItA+1
            Next
            If mustRestartWithOutputToFile Then Command = WScript.Path & "\cscript.exe" _
                Else Command = WScript.FullName
            Command = Command & " """ & WScript.ScriptFullName & """" & AllArgs
            Command = "cmd.exe /c " & Command
            If mustRestartWithOutputToFile Then Command = Command & " > " & OutFile
            DebugEcho 1, 1, "Command = "&Command
            ExitCode = WSHShell.Run(Command, 1, True)
            DebugEcho 1, 0, "Quit: uac.vbs"
            WScript.Quit ExitCode
        End If

        'Restarting the script with elevated privileges
        If Not IsElevated Then
            For Each Arg In WScript.Arguments
                If InStr( Arg, " " ) Then Arg = """" & Arg & """"
                AllArgs = AllArgs & " " & Arg
            Next

            If Not m_ResetCurDir Then _
                AllArgs = AllArgs & " /UAC-Fix-CurrentDirectory:"""&WSHShell.CurrentDirectory&""""

            'Creating a temp file that will be used for start and exit signals
            Dim guid
            If wait Then
                guid = CreateSignal()
                If guid = "" Then DebugEcho 1, 0, "Could not create file for signals": Err.Raise -10000
                AllArgs = AllArgs & " /UAC-Signal-File-GUID:" & guid
            End If

            'Recall the script as administrator
            CommandArgs = """" & WScript.ScriptFullName & """" & AllArgs & " " & FixArg
            CommandArgs = " //nologo " & CommandArgs
            DebugEcho 1, 0, "Recalling the script with 'runas' verb"
            DebugEcho 0, 1, "WScript.FullName = "&WScript.FullName
            DebugEcho 1, 1, "CommandArgs = "&CommandArgs
            If Not selfTest Then CreateObject("Shell.Application").ShellExecute _
                WScript.FullName,_
                CommandArgs,_
                FSO.GetFile(WScript.ScriptFullName).ParentFolder,_
                "runas",_
                1

            'Waiting for execution to finish
            If wait Then
                If WaitForExitSignal(guid) Then
                    OutFile = ""
                    If redirectOutput Then
                        'Outputing the elevated execution output - NO ERRORS here
                        On Error Resume Next
                        OutFile = guid & ".txt"
                        OutFile = WSHShell.ExpandEnvironmentStrings("%TEMP%\" & OutFile)
                        If FSO.FileExists(OutFile) Then
                            ForceEcho 0, 0, "ElevatedOutput" & "(" & OutFile & ")"
                            Set MyFile = fso.OpenTextFile(OutFile, ForReading)
                            Do While MyFile.AtEndOfStream <> True
                                TextLine = MyFile.ReadLine()
                                ForceEcho 0, 1, TextLine
                            Loop
                            MyFile.Close()
                            FSO.DeleteFile(OutFile)
                            ForceEcho 0, 0, "ElevatedOutput"
                        End If
                        On Error GoTo 0
                    End If
                Else
                    DebugEcho 1, 0, "Elevation canceled by user."
                    noQuit = allowNonElevated
                End If
            End If

            If Not noQuit Then
                DebugEcho 1, 0, "Quit: uac.vbs"
                WScript.Quit
            End if
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
        Set namedArgs = WScript.Arguments.Named
        If namedArgs.Exists("UAC-Signal-File-GUID") Then
            guid = namedArgs.Item("UAC-Signal-File-GUID")
            DebugEcho 1,0,"guid = "&guid
            signalFile = WSHShell.ExpandEnvironmentStrings("%TEMP%\"&guid)
            DebugEcho 1,0,"signalFile = "&signalFile
            On Error Resume Next
            Set signalFileStream = FSO.OpenTextFile(signalFile, ForWriting, False)
            If Err.Number <> 0 Then ForceEcho 1,0,"Error("&Err.Number&"): "&Err.Description&" - "&Err.Source
            On Error GoTo 0
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

    Private Function WaitForExitSignal(guid)
        Set WSHShell = CreateObject("WScript.Shell")
        Set FSO = CreateObject("Scripting.FileSystemObject")
        DebugEnterArgs "WaitForExitSignal", "guid = "&guid
        signalFile = WSHShell.ExpandEnvironmentStrings("%TEMP%\"&guid)

        'Waiting for start signal
        DebugEnter "Read Start Msg"
        text = ""
        For X = 1 To 10
            On Error Resume Next
            Set stream = CreateObject("Scripting.FileSystemObject")_
                .OpenTextFile(signalFile, ForReading)
            text = stream.ReadAll()
            stream.Close()
            On Error GoTo 0
            DebugEcho 1,0,"Try #"&X
            If text = "STARTED" Then Exit For
            WScript.Sleep X*50
        Next
        DebugExit "Read Start Msg"

        If text = "" Then
            WaitForExitSignal = False
        Else
            DebugEnter "Read Exit Msg"
            'Waiting for exit signal
            X = 0
            Do While FSO.FileExists(signalFile)
                On Error Resume Next
                FSO.DeleteFile(signalFile)
                On Error GoTo 0
                DebugEcho 1,0,"Try #"&X
                WScript.Sleep 50
                X=X+1
            Loop
            WaitForExitSignal = True
            DebugExit "Read Exit Msg"
        End If
        DebugExitResult "WaitForExitSignal", WaitForExitSignal
    End Function

    Private m_ResetCurDir
    Function ResetCurrentDirectory()
        DebugEnter "UAC:ResetCurrentDirectory"
        ResetCurrentDirectory = False
        If Not m_ResetCurDir And Not m_FixedCurDir Then
            m_ResetCurDir = True
            m_FixedCurDir = True
            Set FSO = CreateObject("Scripting.FileSystemObject")
            Set WSHShell = CreateObject("WScript.Shell")
            DebugEcho 1, 0, "FSO.FileExists('cscript.exe')" &" "& FSO.FileExists("cscript.exe")
            If FSO.FileExists("cscript.exe") Then
                WSHShell.CurrentDirectory = FSO.GetFile(WScript.ScriptFullName).ParentFolder
                ResetCurrentDirectory = True
            End If
        End If
        DebugExitResult "UAC:ResetCurrentDirectory", ResetCurrentDirectory
    End Function

    Private m_FixedCurDir
    Function FixCurrentDirectory()
        DebugEnter "UAC:FixCurrentDirectory"
        FixCurrentDirectory = False
        If Not m_FixedCurDir Then
            m_FixedCurDir = True
            Set FSO = CreateObject("Scripting.FileSystemObject")
            Set WSHShell = CreateObject("WScript.Shell")
            Set Nargs = WScript.Arguments.Named
            DebugEcho 0, 0, "Nargs.Exists('UAC-Fix-CurrentDirectory')" &" "& Nargs.Exists("UAC-Fix-CurrentDirectory") &" "& Nargs.Item("UAC-Fix-CurrentDirectory")
            DebugEcho 1, 0, "FSO.FileExists('cscript.exe')" &" "& FSO.FileExists("cscript.exe")
            If Nargs.Exists("UAC-Fix-CurrentDirectory") Then
                WSHShell.CurrentDirectory = Nargs.Item("UAC-Fix-CurrentDirectory")
                FixCurrentDirectory = True
            ElseIf FSO.FileExists("cscript.exe") Then
                WSHShell.CurrentDirectory = FSO.GetFile(WScript.ScriptFullName).ParentFolder
                FixCurrentDirectory = True
            End If
        End If
        DebugExitResult "UAC:FixCurrentDirectory", FixCurrentDirectory
    End Function

    Private Function IsCScript
        IsCScript = LCase(Right(WScript.FullName, 12)) = "\cscript.exe"
    End Function

    Private Function CreateGUID
      Dim TypeLib
      Set TypeLib = CreateObject("Scriptlet.TypeLib")
      CreateGUID = Replace(Mid(TypeLib.Guid, 2, 36), "-", "")
    End Function

    Private Sub Class_Initialize()
        m_FixedCurDir = False
        m_ElevateCalled = False
        lvl = 0
    End Sub

    Private lvl
    Private Function DebugTime(t)
        DebugTime = t
        If VarType(dbg) <> 0 Then DebugTime = t*dbg.SleepUnit + dbg.SleepBase
    End Function
    Private Sub DebugEcho(t,i,m)
        If VarType(dbg) <> 0 Then dbg.Echo t,i,m Else WScript.Echo String((lvl+i)*4," ")&m:_
            If IsCScript Then WScript.Sleep t*1000
    End Sub
    Private Sub DebugEnter(n)
        lvl=lvl+1
        If VarType(dbg) <> 0 Then dbg.EnterBlock n Else WScript.Echo String((lvl+i)*4," ")&n
    End Sub
    Private Sub DebugExit(n)
        lvl=lvl-1:If lvl<0 Then lvl=0
        If VarType(dbg) <> 0 Then dbg.ExitBlock n Else WScript.Echo String((lvl+i)*4," ")&n
    End Sub
    Private Sub DebugEnterArgs(n,m)
        lvl=lvl+1
        If VarType(dbg) <> 0 Then dbg.EnterBlockArgs n,m Else WScript.Echo String((lvl+i)*4," ")&n&"("&m&")"
    End Sub
    Private Sub DebugExitResult(n,m)
        lvl=lvl-1:If lvl<0 Then lvl=0
        If VarType(dbg) <> 0 Then dbg.ExitBlockResult n,m Else WScript.Echo String((lvl+i)*4," ")&n&" = "&m
    End Sub

    Private Sub ForceEcho(t,i,m)
        'On Error Resume Next
        'WScript.StdErr.WriteLine String((lvl+i)*4," ")&m
        'hasErr = Err.Number <> 0
        'On Error GoTo 0
        'If Not hasErr Then WScript.Sleep t*1000
        'If hasErr Then
            If VarType(dbg) <> 0 Then
                If dbg.IsEnabled Then
                    dbg.Echo t,i,m
                Else
                    WScript.Echo String(i*4," ")&m
                End If
            Else
                DebugEcho t,i,m
            End If
        'End If
    End Sub

    Function PatternMatch(nonElev, elevUacParams, elevFinal)
        If uac.IsElevated Then
            Set namedArgs = WScript.Arguments.Named
            hasUacArgs = namedArgs.Exists("UAC-Fix-CurrentDirectory") _
                Or namedArgs.Exists("UAC-Signal-File-GUID")
            If hasUacArgs Then PatternMatch = elevUacParams Else PatternMatch = elevFinal
        Else
            PatternMatch = nonElev
        End If
    End Function

End Class

Set Export = New UAC