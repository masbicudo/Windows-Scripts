Class Unzip

    Function Unzip(sourceFile, destFolder)
        Set stdout = VBImport("..\std\stdout.vbs")

        ' This Sub extracts FILE specified in sourceFile to the path specified in destFolder.
        '
        ' Improved by Miguel Angelo Santos Bicudo
        ' http://www.masbicudo.com
        '
        ' Based on script written By Justin Godden 2010
        ' http://stackoverflow.com/a/7947219/195417
        
        ' Constants
        Const ForReading = 1, ForWriting = 2, ForAppending = 8

        ' Create a File System Object
        Set objFSO = CreateObject("Scripting.FileSystemObject")

        ' Get the absolute path of the source file
        sourceFile = objFSO.GetAbsolutePathName(sourceFile)
        
        ' Check if the specified target file or folder exists,
        ' and build the fully qualified path of the target file
        Set oShell = CreateObject("WScript.Shell")
        If destFolder = "" Then
            destFolder = objFSO.GetParentFolderName(sourceFile)
        End If
        
        ' Get the absolute path of the output directory
        destFolder = objFSO.GetAbsolutePathName(destFolder)
        If Right(destFolder, 1) <> "\" Then destFolder = destFolder & "\"

        stdout.WriteLine "From """ & sourceFile & """ To """ & destFolder & """"
        stdout.WriteLine "Extracting file"

        If Not objFSO.FileExists(sourceFile) Then
            Err.Raise -10837, "unzip.vbs", "File to unzip not found."
        Else
            Set shellApp = CreateObject("Shell.Application")
            Set objSource = shellApp.NameSpace(sourceFile).Items()
            Set objTarget = shellApp.NameSpace(destFolder)
            intOptions = 256
            objTarget.CopyHere objSource, intOptions

            stdout.WriteLine "File extracted"
        End If

    End Function

End Class

Set Export = New Unzip