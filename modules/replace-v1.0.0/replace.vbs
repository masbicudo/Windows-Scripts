Set objFSO = CreateObject("Scripting.FileSystemObject")
Const ForReading = 1    
Const ForWriting = 2

'WScript.Echo CreateObject("WScript.Shell").CurrentDirectory
'WScript.Echo objFSO.GetAbsolutePathName(".")

Set unnamed = Wscript.Arguments.Unnamed
Set named = Wscript.Arguments.Named

useRegex = named.Exists("R") _
        Or named.Exists("Regex")

ignoreCase = named.Exists("CI") _
        Or named.Exists("CaseInsensitive")

useStd = named.Exists("S") _
        Or named.Exists("StandardStreams")

'If unnamed.Length > 2 Then _
'        WScript.Echo "0: "&unnamed(0)&vbCrLf&_
'             "1: "&unnamed(1)&vbCrLf&_
'             "2: "&unnamed(2)
'If unnamed.Length = 2 Then _
'        WScript.Echo "0: "&unnamed(0)&vbCrLf&_
'             "1: "&unnamed(1)

If (useStd And unnamed.Length <> 2) Or (Not useStd And unnamed.Length <> 3) Then
    WScript.Echo "Usage:" & vbCrLf &_
        "    replace [Options] fileName searchText replaceText" & vbCrLf &_
        "" & vbCrLf &_
        "Options:" & vbCrLf &_
        "    /CI - case insensitive (short form of /CaseInsensitive)" & vbCrLf &_
        "    /R - regular expression (short form of /Regex)" & vbCrLf &_
        "    /S - read and write standard streams" & vbCrLf &_
        "                   (don't indicate a fileName in this case," & vbCrLf &_
        "                   short form of /StandardStreams)" & vbCrLf &_
        "" & vbCrLf &_
        "Arguments:" & vbCrLf &_
        "    fileName - any text file name (absent when /In flag is provided)" & vbCrLf &_
        "    searchText - text to search and then replace" & vbCrLf &_
        "    replaceText - replace with this text"
    WScript.Quit
End If


'Reading the source stream
If useStd Then
    strOldText = unnamed(0)
    strNewText = unnamed(1)
    If Not WScript.StdIn.AtEndOfStream Then
        strText = WScript.StdIn.ReadAll
    End If
Else
    strFileName = unnamed(0)
    strOldText = unnamed(1)
    strNewText = unnamed(2)
    Set objFile = objFSO.OpenTextFile(strFileName, ForReading)
    strText = objFile.ReadAll
    objFile.Close
End If

'Replacing the text
If useRegex Then
    Set regexp = New RegExp
    regexp.IgnoreCase = ignoreCase
    regexp.Pattern = strOldText
    regexp.Global = True
    strNewText = Replace(strNewText, "\r", vbCr)
    strNewText = Replace(strNewText, "\n", vbLf)
    strNewText = Replace(strNewText, "\t", vbTab)
    strNewText = Replace(strNewText, "\\", "\")
    strNewText = regexp.Replace(strText, strNewText)
Else
    If ignoreCase Then strNewText = Replace(strText, strOldText, strNewText, 1, -1, 1)_
        Else strNewText = Replace(strText, strOldText, strNewText)
End If

'Writing the destination stream
If useStd Then
    If strNewText <> "" Then
        WScript.StdOut.Write strNewText
    End If
Else
    Set objFile = objFSO.OpenTextFile(strFileName, ForWriting)
    objFile.Write strNewText
    objFile.Close
End If
