var WshShell = CreateObject("WScript.Shell")
var objEnv = WshShell.Enviroment("Process")
Text = objEnv("PATH")
