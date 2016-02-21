// JavaScript parser.
var ch = importjs("char.js");

// RegExp's that based the parser.
//
// var rgxFnHeader = new RegExp("function\\b name(\\( args\\)) \\{"
//     .replace(/name/g, "(?:(id) )?")
//     .replace(/args/g, "(?:(?:id , )*id )?")
//     .replace(/id/g, "(?:[_\\w][_\\w\\d]*)")
//     .replace(/ /g, "(?:\\s+|\\/\\/[^\\n]*|\\/\\*(?:(?!=\\*\\/).)*\\*\\/)*"));

_export = {
    matchStringOrRegexLiteral: matchStringOrRegexLiteral,
    matchDelimLiteral: matchDelimLiteral,
    matchFunctionHead: matchFunctionHead,
    matchIdentifierAndSpaces: matchIdentifierAndSpaces,
    matchCodeSpaces: matchCodeSpaces,
    matchIdentifier: matchIdentifier,
    isCharCodeAt: isCharCodeAt,
    isBoundary: isBoundary,
    isLetter: isLetter,
    isDigit: isDigit,
    matchString: matchString
};

function decorate(code, pos, fn) {
    var tks = [];
    var num = 0;
    var bkt = [];
    var prev = 0;
    for(; pos < code.length; ) {
        // reading functions
        var m = matchFunctionHead(code, pos);
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
        var m = matchStringOrRegexLiteral(code, pos);
        if (m) {
            tks.push(code.substr(prev, m.start-prev));
            prev = m.start + m.length;
            pos += m.length;
            tks.push(m.value);
            continue;
        }
        
        // reading comments
        var cnt = matchCodeSpaces(code, pos);
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

function matchStringOrRegexLiteral(code, pos){
    var m = matchDelimLiteral(code, pos, 34);
    if (m) {
        m.type = "literal-string-double-quoted";
        return m;
    }
    var m = matchDelimLiteral(code, pos, 39);
    if (m) {
        m.type = "literal-string-single-quoted";
        return m;
    }
    var m = matchDelimLiteral(code, pos, 47);
    if (m) {
        m.type = "literal-regex";
        return m;
    }
    return null;
}

function matchDelimLiteral(code, pos, delim) {
    if (!isCharCodeAt(code, pos, delim))
        return null;
    
    for(var len = 1; pos+len<code.length; len++){
        var c = code.charCodeAt(pos+len);
        if (c == delim)
            return {
                value: code.substr(pos, len+1),
                start: pos,
                length: len+1
            };
        if (c == 92)
            len++;
    }
    return null;
}

function matchFunctionHead(code, pos) {
    var len = 0;
    if (!isBoundary(code, pos+len)) return null;
    var ln = matchString(code, pos, "function");
    if (!ln) return null;
    len+=ln;
    if (!isBoundary(code, pos+len)) return null;
    var ln = matchCodeSpaces(code, pos+len);
    len+=ln;
    // match function name
    var m = matchIdentifierAndSpaces(code, pos+len);
    var name = "";
    if (m) {
        name = m.id;
        len+=m.length;
    }
    // match arguments
    if (pos+len>=code.length) return null;
    if (code.charCodeAt(pos+len++) != 40)
        return null;
    var ln = matchCodeSpaces(code, pos+len);
    len+=ln;
    var args = [];
    
    for(;pos+len<code.length;){
        var m = matchIdentifierAndSpaces(code, pos+len);
        if (!m)
            break;
        args.push(m.id);
        len+=m.length;
        if (pos+len>=code.length) return null;
        if (code.charCodeAt(pos+len) != 44)
            break;
        len++;
        var ln = matchCodeSpaces(code, pos+len);
        len+=ln;
    }

    if (pos+len>=code.length) return null;
    if (code.charCodeAt(pos+len++) != 41)
        return null;
    
    var ln = matchCodeSpaces(code, pos+len);
    len+=ln;
    
    if (pos+len>=code.length) return null;
    if (code.charCodeAt(pos+len++) != 123)
        return null;

    return {
        type: "function-head",
        name: name,
        args: args,
        start: pos,
        length: len
    };
}

function matchIdentifierAndSpaces(code, pos){
    var len = 0;
    var m = matchIdentifier(code, pos);
    if (!m) return null;
    var id = m.value;
    len+=m.length;
    var ln = matchCodeSpaces(code, pos+len);
    len+=ln;
    return {
        type: "function-name",
        id: id,
        start: pos,
        length: len
    };
}

function matchCodeSpaces(code, pos) {
    var len = 0;
    var cont = true;
    var start = pos;
    while(cont){
        for(;pos+len<code.length;len++){
            var c = code.charCodeAt(pos+len);
            if (c != 32 && c != 13 && c != 10 && c != 9)
                break;
        }
        cont = !!len;
        pos+=len;
        len = 0;
        if (pos+len+2<code.length) {
            if (code.charCodeAt(pos+len) == 47 && code.charCodeAt(pos+len+1) == 47) {
                len+=2;
                for(;pos+len<code.length;len++){
                    var c = code.charCodeAt(pos);
                    if (c == 13)
                        break;
                }
                cont = cont || !!len;
                pos+=len;
                len = 0;
            }
            if (code.charCodeAt(pos+len) == 47 && code.charCodeAt(pos+len+1) == 42) {
                len+=2;
                for(;pos+len<code.length;len++){
                    if (pos+len+2<code.length)
                        if (code.charCodeAt(pos+len) == 42 && code.charCodeAt(pos+len+1) == 47) {
                            len+=2;
                            break;
                        }
                }
                cont = cont || !!len;
                pos+=len;
                len = 0;
            }
        }
    }
    return pos - start;
}

function matchIdentifier(code, pos) {
    var isU = isCharCodeAt(code, pos, 95);
    var isB = isBoundary(code, pos);
    var isL = isLetter(code, pos);
    if (!isU && !isB || !isU && !isL)
        return null;
    for(var len=1; ; len++){
        if (!isLetter(code, pos+len) && !isDigit(code, pos+len) && !isCharCodeAt(code, pos+len, 95))
            return {
                type: "identifier",
                value: code.substr(pos, len),
                start: pos,
                length: len
            };
    }
}

function isCharCodeAt(code, pos, chr) {
    var c = code.charCodeAt(pos);
    return pos>=0 && pos<code.length && c == chr;
}

function isBoundary(code, pos){
    var current = isLetter(code, pos) || isDigit(code, pos);
    var previous = isLetter(code, pos-1) || isDigit(code, pos-1);
    return current != previous;
}

function isLetter(code, pos){
    var c = code.charCodeAt(pos);
    return pos>=0 && pos<code.length && ch.isLetter(c);
}

function isDigit(code, pos){
    var c = code.charCodeAt(pos);
    return pos>=0 && pos<code.length && ch.isDigit(c);
}

function matchString(code, pos, str) {
    if (pos+str.length>code.length)
        return 0;
    for(var i=0;i<str.length;i++){
        if (code.charCodeAt(pos+i)!=str.charCodeAt(i))
            return 0;
    }
    return str.length;
}
