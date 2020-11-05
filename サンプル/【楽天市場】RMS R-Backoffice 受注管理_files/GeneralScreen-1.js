/*
 GeneralScreen-1.0.0.min.js
 Copyright (c) 2014 Rakuten.Inc
 Date : 2014/9/18 18:00:00
*/
function GeneralScreen(){this.stickLabel=new StickLabel(this);this.labelDialog=new LabelDialog(this)}
GeneralScreen.prototype={observers:[],seens:[],seen:-1,registerObserver:function(observer){this.observers.push(observer)},removeObserver:function(){throw"removeObserverは実装されていません。";},notifyObservers:function(message){var i=0,len=this.observers.length,self=this;(function loop(){if(i>len-1)return;try{self.observers[i].update(message)}catch(e){}i++;loop()})()},broadcast:function(message){this.notifyObservers(message)},next:function(){this.seen=this.seen+1;this.seens[this.seen]()},clearSeens:function(){this.seens=
[];this.seen=-1},initLabelRelation:function(){this.clearSeens();var self=this;this.seens.push(function(){var message="InitializeLabelDialog#initialize";self.broadcast(message)});this.seens.push(function(){var message="InitializeStickLabel#initBefore";self.broadcast(message)});this.seens.push(function(){var message="InitializeStickLabel#initialize";self.broadcast(message)})}};var generalScreen;$(document).ready(function(){generalScreen=new GeneralScreen;generalScreen.initLabelRelation();generalScreen.next()});
