/*
 StickLabelController-1.0.0.min.js
 Copyright (c) 2014 Rakuten.Inc
 Date : 2014/9/18 18:00:00
*/
function StickLabelController(stickLabel){this.stickLabel=stickLabel}
StickLabelController.prototype={update:function(message){if(message=="InitializeStickLabel#initialize")this.initializeStickLabel_initialize();else if(message=="GetLabelList#get")this.getLabelList_get();else if(message=="PrintLabel#print")this.printLabel_print();else if(message=="MakeMedicineLabel#make")this.makeMedicineLabel_make();else if(message=="GetLabelListAll#get")this.getLabelListAll_get();else if(message=="DisplayLabel#display")this.displayLabel_display();else if(message=="InitializeStickLabel#initBefore")this.initializeStickLabel_initBefore()},
initializeStickLabel_initialize:function(){var initializeStickLabel=this.stickLabel.getInitializeStickLabel();initializeStickLabel.initialize()},getLabelList_get:function(){if(this.isLimited())return;var getLabelList=this.stickLabel.getGetLabelList();getLabelList.get()},printLabel_print:function(){var printLabel=this.stickLabel.getPrintLabel();printLabel.print()},makeMedicineLabel_make:function(){var makeMedicineLabel=this.stickLabel.getMakeMedicineLabel();makeMedicineLabel.make()},getLabelListAll_get:function(){if(this.isLimited())return;
var getLabelListAll=this.stickLabel.getGetLabelListAll();getLabelListAll.get()},displayLabel_display:function(){var displayLabel=this.stickLabel.getDisplayLabel();displayLabel.display()},initializeStickLabel_initBefore:function(){var initializeStickLabel=this.stickLabel.getInitializeStickLabel();initializeStickLabel.initBefore()},isLimited:function(){var limited=this.stickLabel.getLimited();if(limited=="E16-002")return true;return false}};
