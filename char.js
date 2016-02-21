_export = {
    isLetter: isLetter,
    isDigit: isDigit
};

function isLetter(c){
    return c >= 65 && c <= 90 || c >= 97 && c <= 122;
}

function isDigit(c){
    return c >= 48 && c <= 57;
}
