Class Dbg

    Private isdebug
    Private level
    Private blockNames()

    Private Sub Class_Initialize()
        IndentString = "  "
        SleepBase = 0
        SleepUnit = 1000
        SleepEnabled = True
    End Sub

    Sub Enable()
        isdebug = True
    End Sub

    Sub Echo(time, indent, msg)
        If VarType(isdebug) = vbBoolean Then
            If isdebug Then
                ind = ""
                For X = 1 To indent+level: ind = ind&IndentString: Next
                WScript.Echo ind & msg
                If SleepEnabled Then WScript.Sleep time*SleepUnit+SleepBase
            End If
        End If
    End Sub

    Sub EnterBlock(name)
        Redim Preserve blockNames(level+1)
        blockNames(level) = name
        Echo 0, 0, "Enter " & name
        level = level + 1
    End Sub

    Sub ExitBlock(name)
        level2 = level
        Do
            level2 = level2 - 1
            If level2 < 0 Then
                Echo 0, 0, "ERROR: Tried to exit " & name & " but scope was not found"
                Exit Do
            ElseIf blockNames(level2) = name Then
                Redim Preserve blockNames(level2)
                level = level2
                Echo 0, 0, "Exit " & name
                Exit Do
            Else
                Echo 0, 0, "ERROR: Tried to exit " & name & " but scope is " & blockNames(level2)
            End If
        Loop
    End Sub

    Sub EnterBlockArgs(name, strArgs)
        msg = strArgs
        msg = "(" & msg & ")"
        Redim Preserve blockNames(level+1)
        blockNames(level) = name
        Echo 0, 0, "Enter " & name & msg
        level = level + 1
    End Sub

    Sub ExitBlockResult(name, strResult)
        msg = strResult
        If msg <> "" Then msg = " = " & msg
        level2 = level
        Do
            level2 = level2 - 1
            If level2 < 0 Then
                Echo 0, 0, "ERROR: Tried to exit " & name & " but it was not found"
                Exit Do
            ElseIf blockNames(level2) = name Then
                Redim Preserve blockNames(level2)
                level = level2
                Echo 0, 0, "Exit " & name & msg
                Exit Do
            Else
                Echo 0, 0, "ERROR: Tried to exit " & name & " but scope is " & blockNames(level2)
            End If
        Loop
    End Sub

    Private m_IndentString
    Public Property Get IndentString
        IndentString = m_IndentString
    End Property
    Public Property Let IndentString(value)
        m_IndentString = value
    End Property

    Private m_SleepUnit
    Public Property Get SleepUnit
        SleepUnit = m_SleepUnit
    End Property
    Public Property Let SleepUnit(value)
        m_SleepUnit = value
    End Property

    Private m_SleepBase
    Public Property Get SleepBase
        SleepBase = m_SleepBase
    End Property
    Public Property Let SleepBase(value)
        m_SleepBase = value
    End Property

    Private m_SleepEnabled
    Public Property Get SleepEnabled
        SleepEnabled = m_SleepEnabled
    End Property
    Public Property Let SleepEnabled(value)
        m_SleepEnabled = value
    End Property

End Class

Set Export = New Dbg