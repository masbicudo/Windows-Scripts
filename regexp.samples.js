// RegExp test file

// Matching at a specific position and beyond
var rgx = /fn/gm;
WScript.Echo(rgx.lastIndex);
rgx.lastIndex = 3;
WScript.Echo(rgx.exec("as fn fadf fn fs"));
WScript.Echo(rgx.lastIndex);

// Match at a specific position only
var rgx = /^.{3}(fn)/g;
WScript.Echo(rgx.lastIndex);
rgx.lastIndex = 0;
WScript.Echo(rgx.exec("012fn fadf fn fs"));
WScript.Echo(rgx.lastIndex);

