var decorator = importjs("js.function.decorator.js");
var cscript = importjs("force.cscript.js");

cscript.forceCScript(1);

WScript.Echo(decorate(function () {
    return {a:function(){}};
}));
WScript.Echo(decorate(function () {
}));
WScript.Echo(decorate(function F() {
}));
WScript.Echo(decorate(function F1( a ) {
}));
WScript.Echo(decorate(function F2( a, b ) {
}));
WScript.Echo(decorate(function/**/()/**/
{//
}));
WScript.Echo(decorate(function/**/Fc() {//
}));
WScript.Echo(decorate(function/**/F1c( a )//
{//
}));
WScript.Echo(decorate(function/**/F2c( a, b ) {//
}));

function decorate(fn){
    var tks = decorator.decorate(fn.toString(), 0, function(t,n){
        return t=="start"?"callStack.push("+n+");try{":"}finally{callStack.pop(/*"+n+"*/);}";
    });
    return tks.join("");
}

/** IMPORT.JS - MINIFIED VERSION **/
function importjs(t){var r=importjs.list=importjs.list||{};if(r[t])return r[t];
var e=WScript.CreateObject("Scripting.FileSystemObject").OpenTextFile(t,1);
try{var i=e_impjs(e.ReadAll());return r[t]=i,i}finally{e.close()}}
function e_impjs(s){var _export={};return eval(s),_export}