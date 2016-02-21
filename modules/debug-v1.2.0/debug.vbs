' Version: 1.1.0
Class Dbg

    Private isdebug
    Private level
    Private blockNames
    Private blockIgnores
    Private ignoredBlocks

    Private Sub Class_Initialize()
        IndentString = "  "
        SleepBase = 0
        SleepUnit = 1000
        SleepEnabled = IsCScript
        SleepOnBlockChange = 1
        Redim blockNames(127) '128 element
        Redim blockIgnores(127) '128 element
        Redim ignoredBlocks(-1) '0 element
    End Sub

    Sub RedimArrays(size)
        theSize = UBound(blockNames)+1
        If size > theSize Then
            While size > theSize: theSize = theSize * 2: Wend
            Redim Preserve blockNames(theSize-1)
            Redim Preserve blockIgnores(theSize-1)
        End If
    End Sub

    Sub Enable()
        isdebug = True
    End Sub

    Sub IgnoreBlock(name)
        lenBlk = UBound(ignoredBlocks)
        For X = 0 To lenBlk
            If ignoredBlocks(X) = name Then Exit Sub
        Next
        Redim Preserve ignoredBlocks(lenBlk+1)
        ignoredBlocks(lenBlk+1) = name
    End Sub

    Private Function IsBlockIgnored(name)
        IsBlockIgnored = False
        pos = 0
        Do
            pos = InStr(pos+1, name, ":")
            If pos = 0 Then subname = name Else subname = Mid(name, 1, pos-1)

            lenBlk = UBound(ignoredBlocks)
            For X = 0 To lenBlk
                If ignoredBlocks(X) = subname Then
                    IsBlockIgnored = True
                    Exit Function
                End If
            Next

            If pos = 0 Then Exit Do
        Loop
    End Function

    Sub Echo(time, indent, msg)
        If level > 0 Then If blockIgnores(level-1) Then Exit Sub
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
        nextLevel = level+1
        RedimArrays(nextLevel)
        blockNames(level) = name
        If level = 0 Then
            blockIgnores(level) = IsBlockIgnored(name)
        ElseIf blockIgnores(level-1) Then
            blockIgnores(level) = True
        Else
            blockIgnores(level) = IsBlockIgnored(name)
        End If
        If Not blockIgnores(level) Then _
            Echo SleepOnBlockChange, 0, "Enter " & name
        level = nextLevel
    End Sub

    Sub ExitBlock(name)
        level2 = level
        Do
            level2 = level2 - 1
            If level2 < 0 Then
                Echo SleepOnBlockChange, 0, "ERROR: Tried to exit " & name & " but scope was not found"
                Exit Do
            ElseIf blockNames(level2) = name Then
                level = level2
                If Not blockIgnores(level) Then _
                    Echo SleepOnBlockChange, 0, "Exit " & name
                Exit Do
            Else
                Echo SleepOnBlockChange, 0, "ERROR: Tried to exit " & name & " but scope is " & blockNames(level2)
            End If
        Loop
    End Sub

    Sub EnterBlockArgs(name, strArgs)
        msg = strArgs
        msg = "(" & msg & ")"
        nextLevel = level+1
        RedimArrays(nextLevel)
        blockNames(level) = name
        If level = 0 Then
            blockIgnores(level) = IsBlockIgnored(name)
        ElseIf blockIgnores(level-1) Then
            blockIgnores(level) = True
        Else
            blockIgnores(level) = IsBlockIgnored(name)
        End If
        If Not blockIgnores(level) Then _
            Echo SleepOnBlockChange, 0, "Enter " & name & msg
        level = nextLevel
    End Sub

    Sub ExitBlockResult(name, strResult)
        msg = strResult
        If msg <> "" Then msg = " = " & msg
        level2 = level
        Do
            level2 = level2 - 1
            If level2 < 0 Then
                Echo SleepOnBlockChange, 0, "ERROR: Tried to exit " & name & " but it was not found"
                Exit Do
            ElseIf blockNames(level2) = name Then
                level = level2
                If Not blockIgnores(level) Then _
                    Echo SleepOnBlockChange, 0, "Exit " & name & msg
                Exit Do
            Else
                Echo SleepOnBlockChange, 0, "ERROR: Tried to exit " & name & " but scope is " & blockNames(level2)
            End If
        Loop
    End Sub

    Public Property Get IsEnabled
        IsEnabled = isdebug
    End Property

    Private m_IndentString
    Public Property Get IndentString
        IndentString = m_IndentString
    End Property
    Public Property Let IndentString(value)
        m_IndentString = value
    End Property

    Private m_SleepOnBlockChange
    Public Property Get SleepOnBlockChange
        SleepOnBlockChange = m_SleepOnBlockChange
    End Property
    Public Property Let SleepOnBlockChange(value)
        m_SleepOnBlockChange = value
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

    Private Function IsCScript
        IsCScript = LCase(Right(WScript.FullName, 12)) = "\cscript.exe"
    End Function

End Class

Set Export = New Dbg