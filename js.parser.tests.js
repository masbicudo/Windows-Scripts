var parser = importjs("js.parser.js");
var cscript = importjs("force.cscript.js");

cscript.forceCScript(1);

var m = parser.matchIdentifier("_", 0)||{};
WScript.Echo('"'+m.value+'"'+'"'+m.length+'"');

var m = parser.matchFunctionHead("function xpto /**/ (  ) { }", 0)||{};
WScript.Echo('"'+m.name+'"'+'"'+m.args+'"'+'"'+m.length+'"');

var m = parser.matchStringOrRegexLiteral("\'daf\\\'gfas\'", 0)||{};
WScript.Echo('"'+[m.type,m.value,m.start,m.length].join('"\n"')+'"');

var m = parser.matchStringOrRegexLiteral("\"daf\\\"gfas\"", 0)||{};
WScript.Echo('"'+[m.type,m.value,m.start,m.length].join('"\n"')+'"');

var m = parser.matchStringOrRegexLiteral("\/daf\\\/gfas\/", 0)||{};
WScript.Echo('"'+[m.type,m.value,m.start,m.length].join('"\n"')+'"');

/** IMPORT.JS - MINIFIED VERSION **/
function importjs(t){var r=importjs.list=importjs.list||{};if(r[t])return r[t];
var e=WScript.CreateObject("Scripting.FileSystemObject").OpenTextFile(t,1);
try{var i=e_impjs(e.ReadAll());return r[t]=i,i}finally{e.close()}}
function e_impjs(s){var _export={};return eval(s),_export}
/** END IMPORT.JS **/