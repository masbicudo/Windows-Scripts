// Parsing JavaScript to decorate functions.
var parser = importjs("js.parser.js");

_export = {
    decorate: decorate
};

function decorate(code, pos, fn) {
    var tks = [];
    var num = 0;
    var bkt = [];
    var prev = 0;
    for(; pos < code.length; ) {
        // reading functions
        var m = parser.matchFunctionHead(code, pos);
        if (m) {
            tks.push(code.substr(prev, m.start-prev));
            prev = m.start + m.length;
            pos += m.length;
            num++;
            bkt.push(num);
            tks.push(code.substr(m.start, m.length));
            tks.push(fn("start", num));
            continue;
        }

        // reading string literals
        var m = parser.matchStringOrRegexLiteral(code, pos);
        if (m) {
            tks.push(code.substr(prev, m.start-prev));
            prev = m.start + m.length;
            pos += m.length;
            tks.push(m.value);
            continue;
        }
        
        // reading comments
        var cnt = parser.matchCodeSpaces(code, pos);
        if (cnt) {
            tks.push(code.substr(prev, pos-prev));
            tks.push(code.substr(pos, cnt));
            pos += cnt;
            prev = pos;
            continue;
        }
        
        // reading open/close block
        if (code.charCodeAt(pos) == 123)//{
            bkt.push(0);
        else if (code.charCodeAt(pos) == 125)//}
        {
            var outNum = bkt.pop();
            if (outNum) {
                tks.push(code.substr(prev, pos-prev));
                prev=pos;
                tks.push(fn("end", outNum));
            }
        }
        
        pos++;
    }
    tks.push(code.substr(prev, pos-prev));
    return tks;
}
