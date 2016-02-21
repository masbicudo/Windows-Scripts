var callStack = [];

// creating variables needed by the scripts
var window = this;

var document = WScript.CreateObject("HTMLFile"); // "MSXML2.DOMDocument"
document.open();
document.close();
document.write('\
<!DOCTYPE html>\n\
<html xmlns="http://www.w3.org/1999/xhtml">\n\
  <head>\n\
  </head>\n\
  <body style="font-family: Arial;border: 0 none;">\n\
    <div id="branding"  style="float: left;"></div><br />\n\
    <div id="content">Loading...</div>\n\
  </body>\n\
</html>\
');

navigator = {
   appCodeName: "Mozilla",
   appName: "Netscape",
   appVersion: "5.0 (Windows NT 6.1; WOW64; Trident/7.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; .NET4.0C; .NET4.0E; rv:11.0) like Gecko",
   cookieEnabled: true,
   language: "en-US",
   maxTouchPoints: 0,
   mimeTypes: { },
   onLine: true,
   platform: "Win32",
   plugins: { },
   product: "Gecko",
   userAgent: "Mozilla/5.0 (Windows NT 6.1; WOW64; Trident/7.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; .NET4.0C; .NET4.0E; rv:11.0) like Gecko",
   vendor: ""
};

var Global = this;

var xhr = new ActiveXObject('MSXML2.XMLHTTP');
//xhr.onreadystatechange = XhrEventHandler;
xhr.Open('GET', "https://www.google.com/jsapi", false);
xhr.Send();

XhrEventHandler();

function XhrEventHandler()
{
    if (xhr.readyState == 4)
    {
        var text = BinaryToText(xhr.ResponseBody, "utf-8");
        SaveTextFile("jsapi", text);




        var tks = decorate(text, 0, function(t,n){
            return t=="start"?"callStack.push("+n+");try{":"}finally{callStack.pop(/*"+n+"*/);}";
        });
        var newCode = tks.join("");
        //SaveTextFile("new.js", newCode);






        (1, eval).call(Global, newCode);
        google.load('search', '1');
        //SaveTextFile("doc.html", document.body.innerHTML);
        var scripts = document.getElementsByTagName("script");
        while (scripts && scripts.length) {
            for(var i = 0; i<scripts.length; i++){
                var script = scripts[i]; 
                xhr.Open('GET', script.src, false);
                xhr.Send();
                if (xhr.readyState == 4){
                    script.parentNode.removeChild(script);
                    var text = BinaryToText(xhr.ResponseBody, "utf-8");
                    //WScript.Echo(script.src.substr(script.src.lastIndexOf("/")+1));
                    SaveTextFile(script.src.substr(script.src.lastIndexOf("/")+1), text);
                    //WScript.Echo(script.type, script.src, script.text);
                    (1, eval).call(Global, text);
                    SaveTextFile("doc2.html", document.body.innerHTML);
                }
            }
        }

        // Create an Image Search instance.
        var imageSearch = new google.search.ImageSearch();

        // Set searchComplete as the callback function when a search is 
        // complete.  The imageSearch object will have results in it.
        imageSearch.setSearchCompleteCallback(this, searchComplete, null);

        // Find me a beautiful car.
        imageSearch.execute("Subaru STI");
        WScript.Echo('imageSearch.execute("Subaru STI");');
        
        // Include the required Google branding
        google.search.Search.getBranding('branding');
        WScript.Echo("google.search.Search.getBranding('branding');");
        
        function searchComplete() {
            WScript.Echo("searchComplete");
            // Check that we got results
            if (imageSearch.results && imageSearch.results.length > 0) {
                var results = imageSearch.results;
                var items = [];
                for (var i = 0; i < results.length; i++) {
                    var result = results[i];
                    var subitems = [];
                    for(var k in result) {
                        subitems.push(k+":\""+result[k]+"\"");
                    }
                    items.push("{"+subitems.join(",")+"}");
                }
                var str = "["+items.join(",")+"]";
                SaveTextFile("data.txt", str);
            }
        }
    }
}



