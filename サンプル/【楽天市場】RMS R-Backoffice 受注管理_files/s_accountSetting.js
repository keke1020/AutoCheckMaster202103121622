var accountSetting = {};

// set RSID for your environment
accountSetting.useLog = false;
accountSetting.listingParamName = "sclid";
accountSetting.campaignParamName = "scid,sclid";
accountSetting.defaultRSID = "rakutenjprmsdev";
//change bellow to false for DEV/STG environment
accountSetting.dynamicAccountSelection=true
accountSetting.dynamicAccountList="rakutenjprmsprod=rms.rakuten.co.jp";
accountSetting.serviceName = "jprms";
accountSetting.cookieDomainPeriods="3"
accountSetting.currencyCode = "JPY";
accountSetting.trackDownloadLinks = false;
accountSetting.trackExternalLinks = false;
accountSetting.usePrePlugins = true;
accountSetting.usePostPlugins = true;
accountSetting._internalSite = new Array();
accountSetting._internalSite[0] = "javascript:";
accountSetting._internalSite[1] = "rms.rakuten.co.jp";


/*** DON'T TOUCH ***/
accountSetting.linkInternalFilters = accountSetting._internalSite.join(",");