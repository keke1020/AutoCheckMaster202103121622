/*
 MakeTemplates-1.0.0.min.js
 Copyright (c) 2014 Rakuten.Inc
 Date : 2014/9/18 18:00:00
*/
function MakeTemplates(labelDialog){this.labelDialog=labelDialog}
MakeTemplates.prototype={make:function(){this.set$templates();this.makeTemplates()},set$templates:function(){var $templates=$("div.LabelDialogTemplates");this.labelDialog.set$templates($templates)},makeTemplates:function(){var $templates=this.labelDialog.get$templates();this.makeMessageList($templates);this.makeMessageDetail($templates);this.makeMessageEdit($templates);this.makeMessageArea($templates);this.mekeMessageErr($templates)},makeMessageList:function($templates){var markup=$templates.children("div#messageList").get(0).outerHTML;
$.template("messageList",markup)},makeMessageDetail:function($templates){var markup=$templates.children("div#messageDetail").get(0).outerHTML;$.template("messageDetail",markup)},makeMessageEdit:function($templates){var markup=$templates.children("div#messageEdit").get(0).outerHTML;$.template("messageEdit",markup)},makeMessageArea:function($templates){var markup=$templates.children("div#messageArea").get(0).outerHTML;$.template("messageArea",markup)},mekeMessageErr:function($templates){var markup=
$templates.children("div#messageErr").get(0).outerHTML;$.template("messageErr",markup)}};
