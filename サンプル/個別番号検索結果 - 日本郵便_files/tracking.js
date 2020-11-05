//import ua_analyze.js
document.write("<script src=\""+ala_protocol+"//ala.durasite.net/ala_ua_analyze.js?ord="+ala_noCacheParam+"\" type=\"text/javascript\"></script>");

//ala_addEvent(window,"unload",ala_unloadEvent);
ala_addEvent(window,"load",ala_loaded);
var ala_UA = navigator.userAgent;
if(ala_UA.indexOf('Firefox/2') < 0 ){
    ala_addEvent(window,"click",ala_clickEvent);
}
alaThisPage = "unknow"
alaStartTime = ""+new Date().getTime()
alaNextLinkURL = null
alaJsVersion = 4;

var protocol="http";
if (document.location.toString().indexOf("https://", 0) == 0){
    protocol = "https";
    alaJsVersion=3;
}else {
    protocol = "http";
    loadHeatMapJs();
}

function loadHeatMapJs(){
    // LOAD Heat Map JS
    var hmElement = document.createElement("script");
    hmElement.src="http://hm.durasite.net/js/hm_tracking2.js?cid=72";
    var hmobjBody = document.getElementsByTagName("body").item(0);
    if (hmobjBody == null){
        hmobjBody = document.body;
    }
    hmobjBody.appendChild(hmElement);
    // END OF HEATMAP LOADER
}

//document.write("startTime:"+alaStartTime);
//document.write("<br>")
//document.write("ref:"+document.referrer);
ala_setCookie("ALACookieChecker","true");
var ala_cookieEnable = ala_getCookie("ALACookieChecker");
var ala_uid = ala_getCookie("ala_uid");
var ala_sid= ala_getCookie("ala_sid");
var ala_la = ala_getCookie("ala_la");
var alaThisPage=""+document.location

var params = (function(){
    var el = (function(el){
        if(el.nodeName.toLowerCase() == "script"){
            return el;
        }

        return arguments.callee(el.lastChild);
    })(document);

    var src = el ? el.src ? el.src : "" : "";
    var tokens = src.match(/([a-z]+)=([^&]+)/g);
    if (tokens == null){
        return "1";
    }
    var result = {};

    for(var i = 0; i < tokens.length; i++){
        var token = tokens[i];
        var idx = token.indexOf("=");
        result[token.substring(0, idx)] = token.substring(idx + 1);
    }

    return result;
})();

var ala_cid = params.cid;
if (ala_cid == undefined || ala_cid == null){
    ala_cid = "72";
}

function ala_loaded(){
    ala_loadImageForCountupAtLoad();
}

function ala_addEvent(elm,listener,fn){
    try{
        elm.addEventListener(listener,fn,false);
    }catch(e){
        try{
            elm.attachEvent("on"+listener,fn);
        }catch(ex){
        }
    }
}

function ala_encode(data){
    data = data.replace(/\&/g,"|@");
    data = data.replace(/#/g,"%23");
    return encodeURI(data);
}

function ala_manualCountup(url, prevUrl, title){
    var element = document.createElement("script");
    var title = document.title;
    if (title == undefined){
        title ="";
    }
    element.src=protocol+"://ala.durasite.net/A-LogAnalyzer/AccessStart?cid="+ala_cid+"&uid="+ala_uid+"&sid="+ala_sid+"&la="+ala_la+"&ref="+ala_encode(prevUrl)+"&st="+alaStartTime+"&cookie="+ala_cookieEnable+"&t="+ala_encode(title)+"&c="+document.charset+"&url="+ala_encode(url)
    var objBody = document.getElementsByTagName("body").item(0);
    if (objBody == null){
        objBody = document.body;
    }
    objBody.appendChild(element);
}

function ala_loadImageForCountupAtLoad(){
    if(ala_denyUrl(alaThisPage) == true){
      return;
    }

    var element = document.createElement("script");
    //var title = document.title;
    var title = DURASIET_ANALYZER_BROWSER_DATA+"#"+DURASIET_ANALYZER_OS_DATA+"#"+DURASIET_ANALYZER_DEVICE_DATA;
    if (title == undefined){
        title ="";
    }
    element.src=protocol+"://ala.durasite.net/A-LogAnalyzer/AccessStart?cid="+ala_cid+"&uid="+ala_uid+"&sid="+ala_sid+"&la="+ala_la+"&ref="+ala_encode(document.referrer)+"&st="+alaStartTime+"&cookie="+ala_cookieEnable+"&t="+ala_encode(title)+"&c="+document.charset+"&url="+ala_encode(""+document.location)
    var objBody = document.getElementsByTagName("body").item(0);
    if (objBody == null){
        objBody = document.body;
    }
    objBody.appendChild(element);
}

function ala_loadImageForCountupAtUnload(){
    var element = document.createElement("script");
    if (alaNextLinkURL != null){
        alaNextLinkURL = ala_encode(alaNextLinkURL)

        element.src=protocol+"://ala.durasite.net/A-LogAnalyzer/AccessEnd?st="+alaStartTime+"&outUrl="+ala_encode(alaNextLinkURL)+"&url="+ala_encode(alaThisPage)+"&ref="+ala_encode(document.referrer)+"&uid="+ala_uid+"&sid="+ala_sid+"&la="+ala_la;
    }else {
        element.src=protocol+"://ala.durasite.net/A-LogAnalyzer/AccessEnd?st="+alaStartTime+"&url="+ala_encode(alaThisPage)+"&ref="+ala_encode(document.referrer)+"&uid="+ala_uid+"&sid="+ala_sid+"&la="+ala_la;
    }

    var objBody = document.getElementsByTagName("body").item(0);
    objBody.appendChild(element);

}

function ala_unloadEvent(){
    ala_loadImageForCountupAtUnload()
}

function ala_clickEvent(){
    if (document.activeElement.href != undefined){
        alaNextLinkURL = ""+document.activeElement.href
    }
}

function ala_setCookie(key,value,hour){
    var exp = new Date();
    exp.setTime(exp.getTime()+(hour*60*60*1000));
    var itemData =  key + "=" + escape(value) + ";";
    var expires = "expires="+exp.toGMTString()+";";
    document.cookie =  itemData + expires+" path=/;";
}


function ala_getCookie(key){
    var tmp=document.cookie+";";
    var pos=tmp.indexOf(key+"=",0);
    var pos2 = tmp.indexOf("; "+key+"=",0);

    if(pos>0 && pos2 >=0){
        pos = pos2+2;
    }
    if (pos >0 && pos2 <0){
        return "";
    }
    if (pos >= 0){
        tmp=tmp.substring(pos,tmp.length);
        var start=tmp.indexOf("=",0) + 1;
        var end = tmp.indexOf(";",start);
        return(unescape(tmp.substring(start,end)));
    }

    return("");
}

var ala_denyArray = [];
function ala_denyUrl(url){
    for (var i=0; i<ala_denyArray.length; i++) {
        if (url.indexOf(ala_denyArray[i],0) != -1) {
            return true;
        }
    }
}

