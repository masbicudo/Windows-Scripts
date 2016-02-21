// Using ActiveX to load and save text and binary files.

_export = {
    saveFile: saveFile,
    saveTextFile: saveTextFile,
    binaryToText: binaryToText
};

function saveFile(fname, data) {
    var objADOStream = new ActiveXObject("ADODB.Stream");
    try {
        objADOStream.open();
        objADOStream.type = 1; // Binary
        objADOStream.write(data);
        objADOStream.position = 0;
        objADOStream.saveToFile(fname.replace(/[\\\/\:\*\?\"\<\>\|]/, ""), 2);
    }
    finally {
        objADOStream.close();
    }
}

function saveTextFile(fname, text, format) {
    var objADOStream = new ActiveXObject("ADODB.Stream");
    try {
        objADOStream.open();
        objADOStream.type = 2; // Binary
        objADOStream.CharSet = format || "utf-8";
        objADOStream.writeText(text);
        objADOStream.position = 0;
        objADOStream.saveToFile(fname.replace(/[\\\/\:\*\?\"\<\>\|]/, ""), 2);
    }
    finally {
        objADOStream.close();
    }
}

function binaryToText(data, format) {
    var objADOStream = new ActiveXObject("ADODB.Stream");
    try {
        objADOStream.open();
        objADOStream.type = 1; // Binary
        objADOStream.write(data);
        objADOStream.position = 0;
        objADOStream.Type = 2;//adTypeText
        objADOStream.CharSet = format;
        return objADOStream.ReadText();
    }
    finally {
        objADOStream.close();
    }
}

