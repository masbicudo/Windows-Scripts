// cscript rem-arg.js /f: /x:"m a" /"n a":"k l" 10 /n --name

// Joined line of arguments
var joinAll = [];
for(var i=0;i<WScript.Arguments.Length;i++){
    var arg = WScript.Arguments(i);
    var name = getArgName(arg);
    if (name!==undefined) {
        arg = WScript.Arguments.Named(name);
        name = name.indexOf(" ") > -1 ? '"'+name+'"' : name;
        if (arg===undefined)
            joinAll.push("/"+name);
        else
            joinAll.push("/"+name+":"+
                    (arg.indexOf(" ") > -1 ? '"'+arg.replace(/\"/g,'""')+'"' : arg));
    } else {
        joinAll.push(arg.indexOf(" ") > -1 ? '"'+arg.replace(/\"/g,'""')+'"' : arg);
    }
}
WScript.Echo("Joined: "+joinAll.join(" "));

// Each argument alone
WScript.Echo("");
WScript.Echo("Args: "+WScript.Arguments.Length);
var names = [];
for(var i=0;i<WScript.Arguments.Length;i++){
    var arg = WScript.Arguments(i);
    var name = getArgName(arg);
    if (name!==undefined) names.push(name);
    WScript.Echo("  Arg #"+i+" "+(name!==undefined?"(Named)":"(Unnamed)")+": \""+WScript.Arguments(i)+"\"");
}

// Each named argument and it's value
WScript.Echo("");
WScript.Echo("Named: "+WScript.Arguments.Named.Length);
for(var i=0;i<WScript.Arguments.Named.Length;i++){
    var argVal = WScript.Arguments.Named(names[i]);
    argVal = argVal===undefined ? "[none]" : "\""+argVal+"\"";
    WScript.Echo("  Named #"+i+" (\""+names[i]+"\"): "+argVal);
}

// Each unnamed argument value
WScript.Echo("");
WScript.Echo("Unnamed: "+WScript.Arguments.Unnamed.Length);
for(var i=0;i<WScript.Arguments.Unnamed.Length;i++){
    WScript.Echo("  Unnamed #"+i+": \""+WScript.Arguments.Unnamed(i)+"\"");
}

function getArgName(arg){
    var match = /\/([^\:]+)/g.exec(arg);
    if (match) return match[1];
}

//var objSWbemServices = new ActiveXObject("WinMgmts:Root\Cimv2");
//var colProcess = objSWbemServices.ExecQuery("Select * From Win32_Process");
//for(var objProcess in colProcess){
//    var pos = objProcess.CommandLine.indexOf(WScript.ScriptName);
//    if (pos > -1)
//        strLine = objProcess.CommandLine.substr(pos + WScript.ScriptName.length);
//}
//WScript.Echo(strLine)
