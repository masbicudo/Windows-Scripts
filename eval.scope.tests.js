function test(ok, msg){
    if (WScript)
        WScript.Echo((ok ? "GOOD: " : "FAIL: ") + msg);
}

var global = this;
var x = 'outer';
(function() {
  var x = 'inner';
  eval('test(x == "inner", "Inner")'); 
  (1,eval)('test(x == "outer", "Outer")');
})();

var x = 'outer';
(function() {
  var x = 'inner';
  eval('test(x == "inner", "Inner")'); 
  (eval)('test(x == "outer", "Outer")');
})();

var x = 'outer';
(function() {
  var x = 'inner';
  eval('test(x == "inner", "Inner")');
  var i = eval;
  i('test(x == "outer", "Outer")');
})();

var x = 'outer';
var k = eval;
(function() {
  var x = 'inner';
  eval('test(x == "inner", "Inner")');
  k('test(x == "outer", "Outer")');
})();
