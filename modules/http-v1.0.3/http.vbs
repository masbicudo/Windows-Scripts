' Version: 1.0.3
Class HTTP

    Function Download(myURL, myPath)
        If InStr(myPath, ":") > 2 Then Err.Raise -10937, "get.vbs", "File name cannot contain ':' in the path."

        Set fileax = VBImport("..\file-v1.0.0\file.ax.vbs")
        Set mime = VBImport("..\mime-v1.0.0\mime.vbs")

        ' This Sub downloads the FILE specified in myURL to the path specified in myPath.
        '
        ' Improved by Miguel Angelo Santos Bicudo
        ' http://www.masbicudo.com
        '
        ' Based on script written by Rob van der Woude
        ' http://www.robvanderwoude.com
        '
        ' Based on a script found on the Thai Visa forum
        ' http://www.thaivisa.com/forum/index.php?showtopic=21832
        
        ' Constants
        Const ForReading = 1, ForWriting = 2, ForAppending = 8

        ' Create a File System Object
        Set objFSO = CreateObject("Scripting.FileSystemObject")

        ' Check if the specified target file or folder exists,
        ' and build the fully qualified path of the target file
        Set oShell = CreateObject("WScript.Shell")
        If myPath = "" Then
            myPath = oShell.CurrentDirectory
        End If

        'ScriptDir = objFSO.GetParentFolderName(WScript.ScriptFullName)
        'OldWorkDir = oShell.CurrentDirectory
        'oShell.CurrentDirectory = ScriptDir
        myPath = Trim(myPath)
        IsDir = Right(myPath, 1) = "\" Or objFSO.FolderExists(myPath)
        myPath = objFSO.GetAbsolutePathName(myPath)
        'oShell.CurrentDirectory = OldWorkDir
        strFile = myPath
        If IsDir Then
            strFile = objFSO.BuildPath(myPath, "::RCFN::")
        End If
        DownloadDir = objFSO.GetParentFolderName(strFile)
        If Not objFSO.FolderExists(DownloadDir) Then
            objFSO.CreateFolder DownloadDir
        End If
        If Not objFSO.FolderExists(DownloadDir) Then
            Download = "ERROR: Target folder not found. Could not create it."
            Exit Function
        End If

        ' Download the specified URL
        ' and get the file name from server
        On Error Resume Next
        Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
        If objHTTP Is Nothing Then
            Set objHTTP = CreateObject("MSXML2.XmlHttp.6.0")
        End If
        objHTTP.Open "GET", myURL, False
        objHTTP.Send
        If Err.Number <> 0 Then
            Err.Clear
            Download = "ERROR: Could not request URL"
            Exit Function
        End If
        ResponseBody = objHTTP.ResponseBody

        contentTypeHeader = objHTTP.GetResponseHeader("Content-Type")
        WScript.Echo contentDispositionHeader
        
        contentDispositionHeader = objHTTP.GetResponseHeader("Content-Disposition")
        WScript.Echo contentDispositionHeader
        SrvFN = ServerFileName(contentDispositionHeader)
        isUrlFileName = False
        If SrvFN = "" Then
            SrvFN = Mid(myURL, InStrRev(myURL, "/") + 1)
            isUrlFileName = True
        End If
        SrvFN = fileax.RemoveInvalidChars(SrvFN)
        objHTTP.Close
        strFile = Replace(strFile, "::RCFN::", SrvFN)
        On Error GoTo 0
        CurrExt = objFSO.GetExtensionName(strFile)
        If IsDir And isUrlFileName Then
            If Not mime.IsExtensionOfMimeType(CurrExt, contentTypeHeader) Then
                mimeType = LCase(Split(contentTypeHeader, ";")(0))
                If mimeType <> "application/octet-stream" Then
                    Ext = mime.GetExtensionFromMimeType()
                    strFile = strFile & "." & Ext
                End If
            End If
        End If

        fileax.SaveFile strFile, ResponseBody
        Download = strFile
    End Function

    Function ServerFileName(Header)
        ServerFileName = ""
        If VarType(Header) = vbString Then
            If Header = "" Then Exit Function
        ElseIf VarType(Header) = vbObject Then
            If Header Is Nothing Then Exit Function
        Else
            Exit Function
        End If
        FileNamePos = 0
        Do
            FileNamePos = Instr(FileNamePos+1, Header,"filename")
            If FileNamePos > 0 Then
                EqSignPos = InStr(FileNamePos+1, Header, "=")
                strBetween = Trim(Mid(Header, FileNamePos+8, EqSignPos-FileNamePos-8))
                If strBetween = "" Or strBetween = "*" Then
                    If EqSignPos > 0 Then
                        LngSignPos = InStr(EqSignPos+1, Header, "'")
                        If LngSignPos > 0 Then
                            LngSignPos = InStr(LngSignPos+1, Header, "'")
                            If LngSignPos > 0 Then
                                ServerFileName = UrlDecode(Right(Header, Len(Header) - LngSignPos))
                            End If
                        Else
                            ServerFileName = UrlDecode(Right(Header, Len(Header) - EqSignPos))
                        End If
                    End If
                End If
            End If
        Loop While FileNamePos > 0
    End Function

    Function UrlDecode(ByVal str)
        'http://dwarf1711.blogspot.com.br/2007/10/vbscript-urldecode-function.html
        Dim intI, strChar, strRes
        str = Replace(str, "+", " ")
        For intI = 1 To Len(str)
            strChar = Mid(str, intI, 1)
            If strChar = "%" Then
                If intI + 2 < Len(str) Then
                    strRes = strRes & Chr(CLng("&H" & Mid(str, intI+1, 2)))
                    intI = intI + 2
                End If
            Else
                strRes = strRes & strChar
            End If
        Next
        UrlDecode = strRes
    End Function

End Class

Set Export = New HTTP