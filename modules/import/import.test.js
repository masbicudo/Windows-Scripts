jsimport("import.test.res0.js.dat");

function jsimport(f){var r=jsimport.items=jsimport.items||{};if(r[f])return r[f];
var fso=WScript.CreateObject("Scripting.FileSystemObject"),e=fso.OpenTextFile(
fso.BuildPath(fso.GetFile(WScript.ScriptFullName).ParentFolder,f),1);
try{var i=e_impjs(e.ReadAll());return r[f]=i,i}finally{e.close()}}
function e_impjs(s){var _export={};return eval(s),_export}
