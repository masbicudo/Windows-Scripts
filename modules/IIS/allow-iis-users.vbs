Set oShell = WScript.CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")
directory = oFSO.GetParentFolderName(WScript.ScriptFullName)
SetPerms directory, "IIS_IUSRS"
SetPerms directory, "IUSR"

Sub SetPerms(strFile, usr)
    setPermsStr = "%COMSPEC% /c echo Y| C:\windows\system32\icacls.exe " _
        & Chr(34) _
        & strFile _
        & Chr(34) _
        & " /grant " _
        & usr _
        & ":(OI)(CI)F"

    WScript.Echo setPermsStr
    oShell.Run setPermsStr
End Sub