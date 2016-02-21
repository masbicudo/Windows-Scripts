' cscript rem-arg.vbs /f: /x:"m a" /"n a":"k l" 10 /n --name

'' Joined line of arguments
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
            If InStr(Arg, " ") > 0 Then joinAll(it) = "/"&name&":"&""""&Replace(Arg,"""","""""")&"""" _
                Else joinAll(it) = "/"&name&":"&Arg
        End If
    Else
        If InStr(Arg, " ") > 0 Then joinAll(it) = """"&Replace(Arg,"""","""""")&"""" _
            Else joinAll(it) = Arg
    End If
    it = it + 1
Next
WScript.Echo "Joined: "+Join(joinAll, " ")

'' Each argument alone
WScript.Echo ""
WScript.Echo "Args: "&WScript.Arguments.Length
Dim names(): it = 0
Redim names(WScript.Arguments.Length)
For Each Arg In WScript.Arguments
    name = GetArgName(Arg)
    If VarType(name) = vbString _
        Then strNameType = "(Named)": names(it) = name: it = it + 1 _
        Else strNameType = "(Unnamed)"
    WScript.Echo "  Arg #"&it&" "&strNameType&": """&Arg&""""
Next
Redim Preserve names(it)

'' Each named argument and it's value
WScript.Echo ""
WScript.Echo "Named: "&WScript.Arguments.Named.Length
For it = 0 To WScript.Arguments.Named.Length-1
    argVal = WScript.Arguments.Named(names(it))
    If VarType(argVal) = vbEmpty _
        Then argVal = "[none]" _
        Else argVal = """"+argVal+""""
    WScript.Echo "  Named #"&it&" ("""&names(it)&"""): "&argVal
Next

'' Each unnamed argument value
WScript.Echo ""
WScript.Echo "Unnamed: "&WScript.Arguments.Unnamed.Length
it = 0
For Each Arg In WScript.Arguments.Unnamed
    WScript.Echo "  Unnamed #"&it&": """&Arg&""""
    it = it + 1
Next

Function GetArgName(Arg)
    If Left(Arg, 1) = "/" Then
        pos = InStr(2, Arg, ":")
        If pos > 1 Then GetArgName = Mid(Arg, 2, pos-2) Else GetArgName = Mid(Arg, 2)
    End If
End Function

'Set objSWbemServices = GetObject("WinMgmts:Root\Cimv2") 
'Set colProcess = objSWbemServices.ExecQuery("Select * From Win32_Process") 
'For Each objProcess In colProcess 
'    If InStr (objProcess.CommandLine, WScript.ScriptName) <> 0 Then 
'        strLine = Mid(objProcess.CommandLine, InStr(objProcess.CommandLine , WScript.ScriptName) + Len(WScript.ScriptName) + 1)
'    End If 
'Next
'WScript.Echo(strLine)
