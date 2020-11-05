
//*****************************************************************
//	Cookie生成処理
//	　ユーザID、パスワードによりCookie情報の生成
//*****************************************************************
function createCookie( contextPath ) {
	var userid  = burauza('user').value;
	var passwd  = burauza('pass').value;
	var keepVal = burauza('idkeep').value;

	/* Cookieを保存 */
	if (keepVal == '01') {
		if (userid.length > 0 && passwd.length > 0) {
			document.cookie = "userid=" + escape(userid) 
							+ "; domain=" + document.domain
							+ "; path=" + contextPath
							+ "; expires=Tue, 31-Dec-2030 23:59:59 GMT; ";
		}
	}
	/* Cookieを消去 */
	if (keepVal == '02') {
		document.cookie = "userid=xx"
						+ "; domain=" + document.domain
						+ "; path=" + contextPath
						+ "; expires=Wed, 1-Jan-1997 00:00:00 GMT;";
	}
}

//*****************************************************************
//	Cookie情報設定処理
//	　CookieからユーザIDを取得して、画面に設定する
//*****************************************************************
function SetFromCookie() {
	var tmp1, tmp2, xx1, xx2, xx3;
	var userid  = burauza('user');
	var passwd  = burauza('pass');
	
	if (userid == void(0) || passwd == void(0)) return;
	if (userid.value == null || passwd.value == null) return;
	tmp1 = " " + document.cookie + ";";
	xx1 = xx2 = 0;
	len = tmp1.length;
	while (xx1 < len) {
		xx2 = tmp1.indexOf(";", xx1);
		tmp2 = tmp1.substring(xx1 + 1, xx2);
		xx3 = tmp2.indexOf("=");
		if (tmp2.substring(0, xx3) == "userid") {
			userid.value = (unescape(tmp2.substring(xx3 + 1, xx2 - xx1 - 1)));
			passwd.value = "";
			return;
		}
		xx1 = xx2 + 1;
	}
	userid.value="";
}
