var su = importjs("string.utils.js");
_export = {
    forceCScript: ForceCScriptExecution
}
function ForceCScriptExecution(useCmd) {
    if ((su.right(WScript.FullName, 12)).toLowerCase() != "\\cscript.exe") {
        var args = [];
        for(var i = 0; i < WScript.Arguments.length; i++) {
            var a = WScript.Arguments(i);
            args.push(a.indexOf(" ") >= 0 ? '"'+a+'"' : a);
        }
        var shell = new ActiveXObject("WScript.Shell");
        var command = WScript.Path +
            "\\cscript.exe //nologo \"" +
            WScript.ScriptFullName +
            "\" " + args.join(" ");
        if (useCmd){
            // Run doc: http://ss64.com/vb/run.html
            var exitCode = shell.Run("cmd.exe /k " + command, 1, true);
            WScript.Quit(exitCode);
        }
        else {
            // Run doc: http://ss64.com/vb/exec.html
            var dbg = shell.Exec(command);
            while(dbg.Status == 0)
                WScript.Sleep(1);
            WScript.Quit(dbg.ExitCode);
        }
    }
}
