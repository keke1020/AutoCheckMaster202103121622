var DURASIET_ANALYZER_BROWSER_DATA = ALA_getBrowser();
var DURASIET_ANALYZER_OS_DATA = ALA_getOs();
var DURASIET_ANALYZER_DEVICE_DATA = ALA_getDevice();

function ALA_getBrowser(){
	var ua = navigator.userAgent;
	//alert(ua);	
	
	if (ua.indexOf('Opera') != -1) {
		return 'Opera';
	
	} else if (ua.indexOf('MSIE') != -1) {
  		return 'IE' + ALA_getEI_Version(ua);

	} else if (ua.indexOf('Firefox') != -1) {
		return 'FireFox' + ALA_getFF_Version(ua);
	
	} else if (ua.indexOf('Safari') != -1) {
		return ALA_getSafariUA(ua);
	
	} else {
		return "Other";
	}
}

function ALA_getOs(){
	var os,ua = navigator.userAgent;
	if (ua.match(/Win(dows )?NT 6\.2/)) {
		os = "Windows_8";
	}else if (ua.match(/Win(dows )?NT 6\.1/)) {
		os = "Windows_7";
	}else if (ua.match(/Win(dows )?NT 6\.0/)) {
		os = "Windows_Vista";
	}else if (ua.match(/Win(dows )?NT 5\.2/)) {
		os = "Windows_Server_2003";
	}else if (ua.match(/Win(dows )?(NT 5\.1|XP)/)) {
		os = "Windows_XP";
	}else if (ua.match(/Win(dows)? (9x 4\.90|ME)/)) {
		os = "Windows_ME";
	}else if (ua.match(/Win(dows )?(NT 5\.0|2000)/)) {
		os = "Windows_2000";
	}else if (ua.match(/Win(dows )?98/)) {
		os = "Windows_98";
	}else if (ua.match(/Win(dows )?NT( 4\.0)?/)) {
		os = "Windows_NT";
	}else if (ua.match(/Win(dows )?95/)) {
		os = "Windows_95";
	}else if (ua.match(/Mac|PPC/)) {
		
		if(ua.indexOf('Macintosh')!=-1){
			if(ua.indexOf('Mac OS X')!=-1){
				os='Mac_OS_X';
			}else if(ua.indexOf('PPC')!=-1){
				os='Mac_Powor_PC';
			}else{
				os='Mac_OS';
			}
		}else if(ua.indexOf('like')){	
			//iOSの判定を追加する
			os='iOS';
		}
	}else if (ua.match(/Linux/)) {
		if(ua.match(/Android/)){
			//Android系の判定
			os = "Android";
		}else{
			// Linux の処理
			os = "Linux";
		}
	}else if (ua.match(/^.*\s([A-Za-z]+BSD)/)) {
		os = RegExp.$1;					// BSD 系の処理
	}else if (ua.match(/SunOS/)) {
		os = "Solaris";					// Solaris の処理
	}else {
		os = "Other";					// 上記以外 OS の処理
	}
	return os;
}

function ALA_getDevice(){
	var device,ua = navigator.userAgent;
	if(ua.match(/iPad/)){
		device = "iPad";
	}else if(ua.match(/iPod/)){
		device = "iPod";
	}else if(ua.match(/iPhone/)){
		device = "iPhone";
	}else if(ua.match(/Android/)){
		device = "Android";
	}else{
		device = "Other";
	}
	
	return device;
}

function ALA_getEI_Version(UA){
	var START,END;  //切り取り開始点と終了点
	var VERSION;    //ブラウザバージョンを格納する変数

	//最初にInternet Explorerか調べます
	START = UA.indexOf("MSIE");
	//バージョン番号の後に「;」があるのでそれを検索
	END = UA.indexOf(";",START);

	//「MSIE」＋半角スペースの5文字を加えた位置から切り取る
	VERSION = UA.substring(START+5,END);
	VERSION = VERSION.split(".");

	return VERSION[0];
}

function ALA_getFF_Version(UA){
	var START,END;  //切り取り開始点と終了点
	var VERSION;    //ブラウザバージョンを格納する変数
	
	START = UA.indexOf("Firefox");
	//FireFoxの場合、バージョンは最後。つまり終了点は文字列全体の長さ。
	END = UA.length;
	//「Firefox」＋「/」の8文字を加えた位置から切り取ります
	VERSION = UA.substring(START+8,END);
	
	//頭に「FireFox」を付けてバージョンを書き出す
	VERSION = VERSION.split(".");
	return VERSION[0];
}

function ALA_getSafariUA(UA){
	//「Chrome」が含まれるか調べる
	if ( UA.indexOf("Chrome") != -1 ){
		START = UA.indexOf("Chrome");
		END   = UA.indexOf(" ",START);
		VERSION =UA.substring(START+7,END);
    		VERSION = VERSION.split(".");
		
		return "GoogleChrome"+VERSION[0];

	}else{
    		START = UA.indexOf("Version");
    		END   = UA.indexOf(" ",START);    
    		VERSION = UA.substring(START+8,END);
		VERSION = VERSION.split(".");
		//VERSION = VERSION.replace("/","");
	    	var ver = VERSION[0].replace("/","");
		return "Safari"+ver;

	}
}
