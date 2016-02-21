var rgxFnHeader = new RegExp("function\\b name(\\( args\\)) \\{"
    .replace(/name/g, "(?:(id) )?")
    .replace(/args/g, "(?:(?:id , )*id )?")
    .replace(/id/g, "(?:[_\\w][_\\w\\d]*)")
    .replace(/ /g, "(?:\\s+|\\/\\/[^\\n]*|\\/\\*(?:(?!=\\*\\/).)*\\*\\/)*"));

var f = function () {
}

function F() {
}

function F1( a ) {
}

function F2( a, b ) {
}

var fc = function/**/()/**/
{//
}

function/**/Fc() {//
}

function/**/F1c( a )//
{//
}

function/**/F2c( a, b ) {//
}

WScript.Echo(r.exec(f));
WScript.Echo(r.exec(F));
WScript.Echo(r.exec(F1));
WScript.Echo(r.exec(F2));

WScript.Echo(r.exec(fc));
WScript.Echo(r.exec(Fc));
WScript.Echo(r.exec(F1c));
WScript.Echo(r.exec(F2c));

var r = new RegExp("/(?:[^/\\\\]|\\\\.)+/[mgiuy]*|'(?:[^'\\\\]|\\\\.)*'|\"(?:[^\"\\\\]|\\\\.)*\"");

WScript.Echo(r.exec("'sdg\\'fg'"));
WScript.Echo(r.exec('"fdf\\"dsg"'));
WScript.Echo(r.exec("/ghd\\/sj/gim"));
