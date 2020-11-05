/* 別ＷｉｎｄｏｗＯｐｅｎ */
/* openwinもsubwinに直す */
function subwin(urlname, tate, yoko, wndname){
	var tmp;
	var ftemp;

	tmp = "height=" + (tate + "") + ",width=" + (yoko + "") + ",resizable=Yes, scrollbars=yes";
	ftemp = window.open(urlname, wndname, tmp);
	ftemp.focus();
}
