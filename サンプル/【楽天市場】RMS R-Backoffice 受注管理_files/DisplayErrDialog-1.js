/*
 DisplayErrDialog-1.0.0.min.js
 Copyright (c) 2014 Rakuten.Inc
 Date : 2014/9/18 18:00:00
*/
function DisplayErrDialog(labelDialog){this.labelDialog=labelDialog}
DisplayErrDialog.prototype={dispSysErr:function(){this.emptyDialog();var data=this.getSysMsgs();this.setMsgData(data);this.set$messages();this.printDialog()},setMsgData:function(data){var nowMsgData=this.labelDialog.getMsgData();var msgData=$.extend(true,{},nowMsgData,data);this.labelDialog.setMsgData(msgData)},getSysMsgs:function(){var errMsgs={};return errMsgs},display:function(){this.emptyDialog();this.set$messages();this.printDialog()},set$messages:function(){var $messageErr={};var data=this.labelDialog.getMsgData();
var tmplOptions=this.labelDialog.getTmplOptions();try{$messageErr=$.tmpl("messageErr",data,tmplOptions)}catch(e){}this.labelDialog.set$messageErr($messageErr)},printDialog:function(){var $Dialog=this.labelDialog.get$Dialog();var $messageErr=this.labelDialog.get$messageErr();$Dialog.append($messageErr)},emptyDialog:function(){var $Dialog=this.labelDialog.get$Dialog();$Dialog.empty()}};
