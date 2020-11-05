/*
 MakeMedicineLabel-1.0.0.min.js
 Copyright (c) 2014 Rakuten.Inc
 Date : 2014/9/18 18:00:00
*/
function MakeMedicineLabel(stickLabel){this.stickLabel=stickLabel}
MakeMedicineLabel.prototype={make:function(){if(!this.isMedicine())return;this.overwrite();this.addClass()},overwrite:function(){if(this.isUnsent()==false)return;var $AdditionalLabelChild=this.stickLabel.get$AdditionalLabelChild();var labelData={};var url="drug_msg_unsend.gif";labelData.labelPath=url;$AdditionalLabelChild.data(labelData)},addClass:function(){var $AdditionalLabelChild=this.stickLabel.get$AdditionalLabelChild();if(this.isLabelTypeNull())$AdditionalLabelChild.addClass("PostLabelCreate");else $AdditionalLabelChild.addClass("GetMessageList")},
isUnsent:function(){if(this.isLabelTypeNull())return true;if(this.isLabelTypeZero())return true;return false},isIconDrugTrue:function(){var $AdditionalLabel=this.stickLabel.get$AdditionalLabel();if($AdditionalLabel.children("span.IconDrug").text()=="true")return true;return false},isLabelTypeZero:function(){var $AdditionalLabelChild=this.stickLabel.get$AdditionalLabelChild();var labelType=$AdditionalLabelChild.data("labelType");if(labelType==0)return true;return false},isLabelTypeNull:function(){var $AdditionalLabelChild=
this.stickLabel.get$AdditionalLabelChild();var labelType=$AdditionalLabelChild.data("labelType");if(labelType==null)return true;return false},isMedicine:function(){var $AdditionalLabelChild=this.stickLabel.get$AdditionalLabelChild();var labelType=$AdditionalLabelChild.data("labelType");if(this.isIconDrugTrue())return true;else if(labelType==0)return true;else if(labelType==1)return true;else if(labelType==2)return true;else if(labelType==3)return true;return false}};
