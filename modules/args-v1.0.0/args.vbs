Class Args

    Function CommandLineArgs()
        Dim joinAll(): it = 0
        Redim joinAll(WScript.Arguments.Length)
        For Each Arg In WScript.Arguments
            name = GetArgName(Arg)
            If VarType(name) = vbString Then
                Arg = WScript.Arguments.Named(name)
                If InStr(name, " ") Then name = """"&name&""""
                If VarType(Arg) = vbEmpty Then
                    joinAll(it) = "/" & name
                Else
                    If InStr(Arg, " ") > 0 Then joinAll(it) = "/"&name&":"&""""&Replace(Arg,"""","\""")&"""" _
                        Else joinAll(it) = "/"&name&":"&Arg
                End If
            Else
                If InStr(Arg, " ") > 0 Then joinAll(it) = """"&Replace(Arg,"""","\""")&"""" _
                    Else joinAll(it) = Arg
            End If
            it = it + 1
        Next
        CommandLineArgs = Join(joinAll, " ")
    End Function

    Function ToCommandLine(Arg)
        name = GetArgName(Arg)
        If VarType(name) = vbString Then
            Val = GetArgValue(Arg)
            If InStr(name, " ") Then name = """"&name&""""
            If VarType(Val) = vbEmpty Then
                ToCommandLine = "/" & name
            Else
                If InStr(Val, " ") > 0 Then ToCommandLine = "/"&name&":"&""""&Replace(Val,"""","\""")&"""" _
                    Else ToCommandLine = "/"&name&":"&Val
            End If
        Else
            If InStr(Arg, " ") > 0 Then ToCommandLine = """"&Replace(Arg,"""","\""")&"""" _
                Else ToCommandLine = Arg
        End If
    End Function

    Function ArgNames()
        Dim names(): it = 0
        Redim names(WScript.Arguments.Length)
        For Each Arg In WScript.Arguments
            name = GetArgName(Arg)
        Next
        Redim Preserve names(it)
        ArgNames = names
    End Function

    Function IsNamed(arg)
        If Left(Arg, 1) = "/" And Len(Arg) > 1 Then
            IsNamed = True
        Else
            IsNamed = False
        End If
    End Function

    Function GetArgName(Arg)
        If Left(Arg, 1) = "/" And Len(Arg) > 1 Then
            pos = InStr(2, Arg, ":")
            If pos > 1 Then GetArgName = Mid(Arg, 2, pos-2) Else GetArgName = Mid(Arg, 2)
        End If
    End Function

    Function GetArgValue(Arg)
        If Left(Arg, 1) = "/" And Len(Arg) > 1 Then
            pos = InStr(2, Arg, ":")
            If pos > 1 Then GetArgValue = Mid(Arg, pos+1)
        Else
            GetArgValue = Arg
        End If
    End Function

End Class

'Class CmdArgs
'
'End Class
'
'Class CmdArg
'    Sub Init(Arg)
'        Name = GetArgName(Arg)
'        Value = GetArgValue(Arg)
'    End Sub
'End Class

Set Export = New Args