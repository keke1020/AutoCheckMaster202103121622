/* PRE-PLUGINS SECTION */
var do_PrePlugins = function() {
// your customized code is here
// Channel Manager Parameter config

	var _pagename = "";
	var sDomainInfo = "";

	// pageName
	if(!s.pageType&&!s.pageName){

		// Get Page Name
		_pagename=s.getPageName()
		if(_pagename)_pagename=_pagename.replace(/\.[a-z]+$/,"").replace(":index","")
		if(!_pagename)_pagename="top"

		// Get Domain Information from Hostname
		var sHostName = location.hostname;
		var _i=0;
		var _bFoundRms = false;

		var aData = sHostName.split(".");

		for(_i=0;_i<aData.length;_i++)
		{
			if(aData[_i] == "rms")
			{
				_bFoundRms = true;
				break;
			}

			sDomainInfo = aData[_i];
		}

		// Initialize sDomainInfo if host name does not include "rms"
		if(_bFoundRms==false)
		{
			sDomainInfo = "";
			s.pageName = _pagename;
		}
		else
		{
			s.pageName = sDomainInfo + ":" + _pagename;
			
		}
		
		//remove:?
		s.pageName = s.pageName.replace(/\:\?/,"");
		

	}

	//channel
	if(!s.channel){
		//s.channel=s.pageName.split(":")[0]

		if(sDomainInfo == "")
		{
			s.channel = _pagename.split(":")[0];
		}
		else
		{
			s.channel = sDomainInfo + ":" + _pagename.split(":")[0];
		}
		
		s.channel = s.channel.replace(/\:\?/,"");
	}

	//RakutenPageType
	if(s.RakutenPageType && (typeof s.eo == "undefined")){
		s.pageName+="["+s.RakutenPageType+"]"
		s.channel+="["+s.RakutenPageType+"]"
	}

	//deeper analysis of channel
	//s.prop6
	if(!s.prop6){
		if(_pagename.split(":")[1])
		{
			s.prop6 = s.channel + ":" + _pagename.split(":")[1];
			s.prop6 = s.prop6.replace(/\:\?/,"");
		}
		else
		{
			s.prop6 = s.channel;
		}
	}
	
	//deeper analysis of channel
	//s.prop7
	
	if(!s.prop7){
		if(_pagename.split(":")[2])
		{
			s.prop7 = s.prop6 + ":" + _pagename.split(":")[2];
			s.prop7 = s.prop7.replace(/\:\?/,"");
		}else{
			s.prop7 = s.prop6;
		}
	}
	
	//get Query
	s.prop10 = location.search.replace(/\?/,"")
	
	//pageType for c40
	
	
	if(!s.prop40){
		if(location.host.match(/mainmenu\.rms/) && location.search.match(/param/)){
			if(!s.getQueryParam('params')){
				s.prop40 = "manual: top"
			}else{
				s.prop40 = s.getQueryParam('params').replace(/\.[a-z]+$/,"")
			}
		}else if(location.host.match(/help\.rms\.rakuten\.co\.jp/) && location.pathname.match(/^\/mw\//)){
			s.prop40 = s.pageName + ":" + "hid:" + s.getQueryParam('hid')
		}
		
		else{
			s.prop40 = s.pageName;
		}
	}

	
}

/* POST-PLUGINS SECTION */
var do_PostPlugins = function() {
// your customized code is here

	// for MST global tracking
	if(s.events&&s.events.match(/purchase/)){
		//s.events=s.apl(s.events,"event71")
		//s.eVar49=s.prop50+":"+"purchase"
	}
	if(!s.eo&&!s.lnk&&!s.pageType&&!s.un.match(/dev/)&&!s.un.match(/rakutenglobal/)){
		if(s.campaign.match(/_gmx/)||s.campaign.match(/_upc/)||s.eVar49){s.un=s.apl(s.un,"rakutenglobalprod")}
	}
}


/* CUSTOM-PLUGIN SECTION */


/* CODE SECTION - DON'T TOUCH BELOW */
if(s.usePrePlugins)s.doPrePlugins = do_PrePlugins;
if(s.usePostPlugins)s.doPostPlugins = do_PostPlugins;

/************* To Stop Google Preview From Being Counted *************/
if(navigator.userAgent.match(/Google Web Preview/i)){
	s.t=new Function();
	s.tl=new Function();
}
