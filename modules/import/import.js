function jsimport(from){
    var jsimportMap = jsimport.items = jsimport.items || {};
    if (jsimportMap[from]) return jsimportMap[from];
    var fso = WScript.CreateObject("Scripting.FileSystemObject");
    var fileName = fso.BuildPath(fso.GetFile(WScript.ScriptFullName).ParentFolder, from);
    var fileStream = fso.OpenTextFile(fileName, 1 /*ForReading*/);
    try {
        var _export = evalImport(fileStream.ReadAll());
        jsimportMap[from] = _export;
        return _export;
    }
    finally{ fileStream.close(); }
}
function evalImport(s){
    var _export = {};
    eval(s);
    return _export;
}

///** IMPORT.JS - MINIFIED VERSION **/
//function jsimport(f){var r=jsimport.items=jsimport.items||{};if(r[f])return r[f];
//var fso=WScript.CreateObject("Scripting.FileSystemObject"),e=fso.OpenTextFile(
//fso.BuildPath(fso.GetFile(WScript.ScriptFullName).ParentFolder,f),1);
//try{var i=e_impjs(e.ReadAll());return r[f]=i,i}finally{e.close()}}
//function e_impjs(s){var _export={};return eval(s),_export}
///** END IMPORT.JS **/
