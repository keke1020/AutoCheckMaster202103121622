/**
 * common.js
 *
 *  version --- 2.0.1
 *  updated --- 2014/11/25
 */


/* ! viewportの上書き --------------------------------------------------- */
;(function () {
	var w = $(window).width();
	var original = $('meta[name="viewport"]').attr('content');
	var pc = 'width=1024, user-scalable=yes';
	if( w > 736 ) {
		$('meta[name="viewport"]').attr('content', pc);
	} else {
		$('meta[name="viewport"]').attr('content', original);
	}
}());

// modal
$(function(){
	$('.js_modalOpen').magnificPopup({
		type: 'iframe',
		mainClass: 'mfp-fade',
		removalDelay: 200,
		preloader: false,
		fixedContentPos: false
	});
});

// linkify images
$(function(){
	$(document).on('click', 'img.linkify', function ( e ) {
		if( $(window).width() > 736 ) {
			e.preventDefault();
		} else {
			window.open(this.src, '_blank');
		}
	});
});


/* !stack ------------------------------------------------------------------- */
jQuery(document).ready(function($) {
	pageScroll();
	rollover();
	localNav();
	localNav02();
	addCss();
	scrollTop();
	tileHeight();
	picColumnWidth();
	picCaption01();
	serviceCatIconSets();
});

/* !共通部分 ------------------------------------------------------------------- */
function cmnInclude (name, is_www2) { 'use strict';
	var BASE_URL = 'http://www.sagawa-exp.co.jp';
	var INC_PATH = '../common/pc/inc/';

	$.ajax({
		url: INC_PATH + name + '.html',
		cache: false,
		async: false,
		success: function(html){
			if (is_www2) {
				html = html.replace(/ href="\//gi, ' href="' + BASE_URL + '/');
			}
			document.write(html);
		}
	});
}

// 別ドメイン（システム周り）からの呼び出しの場合は、
// 第2引数に`true`を指定することで`BASE_URL`の追加された絶対パスに書き換わります。
// ただし、書き換わるのはルートパスのみです。

function cmnHeader(is_www2) { cmnInclude("header", is_www2); } // ヘッダー
function cmnFooter(is_www2) { cmnInclude("footer", is_www2); } // フッター

/* ローカルナビ */
function cmnSub_dummy(is_www2) { cmnInclude("sub_dummy", is_www2); } // ダミー
function cmnSub_styleguide(is_www2) { cmnInclude("sub_styleguide", is_www2); } // スタイルガイド
function cmnSub_goal(is_www2) { cmnInclude("sub_goal", is_www2); } // ゴール
function cmnSub_service01(is_www2) { cmnInclude("sub_service01", is_www2); } // サービス一覧 Standard
function cmnSub_service02(is_www2) { cmnInclude("sub_service02", is_www2); } // サービス一覧 Support
function cmnSub_service03(is_www2) { cmnInclude("sub_service03", is_www2); } // サービス一覧 Solutions
function cmnSub_service04(is_www2) { cmnInclude("sub_service04", is_www2); } // サービス一覧 その他
function cmnSub_delivery(is_www2) { cmnInclude("sub_delivery", is_www2); } // 送る・受け取る
function cmnSub_company(is_www2) { cmnInclude("sub_company", is_www2); } // 会社案内
function cmnSub_csr(is_www2) { cmnInclude("sub_csr", is_www2); } // CSR
function cmnSub_contact(is_www2) { cmnInclude("sub_contact", is_www2); } // お問い合わせ

/* 採用 */
function cmnHeader_recruit(is_www2) { cmnInclude("header_recruit", is_www2); } // ヘッダー
function cmnFooter_recruit(is_www2) { cmnInclude("footer_recruit", is_www2); } // フッター
function cmnSub_recruit(is_www2) { cmnInclude("sub_recruit", is_www2); } // 新卒

/* 多言語 */
/* English */
function cmnHeader_lang_en(is_www2) { cmnInclude("header_lang_en", is_www2); } // Englishヘッダー
function cmnFooter_lang_en(is_www2) { cmnInclude("footer_lang_en", is_www2); } // Englishフッター
function cmnSub_company_en(is_www2) { cmnInclude("sub_company_en", is_www2); } // company
function cmnSub_price_en(is_www2) { cmnInclude("sub_price_en", is_www2); } // price
function cmnSub_service01_en(is_www2) { cmnInclude("sub_service01_en", is_www2); } // Service Standard
function cmnSub_service02_en(is_www2) { cmnInclude("sub_service02_en", is_www2); } // Service Support
function cmnSub_service03_en(is_www2) { cmnInclude("sub_service03_en", is_www2); } // Service Solutions

/* China */
function cmnHeader_lang_cn(is_www2) { cmnInclude("header_lang_cn", is_www2); } // Chineseヘッダー
function cmnFooter_lang_cn(is_www2) { cmnInclude("footer_lang_cn", is_www2); } // Chineseフッター
function cmnSub_company_cn(is_www2) { cmnInclude("sub_company_cn", is_www2); } // company
function cmnSub_price_cn(is_www2) { cmnInclude("sub_price_cn", is_www2); } // price
function cmnSub_service01_cn(is_www2) { cmnInclude("sub_service01_cn", is_www2); } // Service Standard
function cmnSub_service02_cn(is_www2) { cmnInclude("sub_service02_cn", is_www2); } // Service Support
function cmnSub_service03_cn(is_www2) { cmnInclude("sub_service03_cn", is_www2); } // Service Solutions

/* ナビ無し */
function cmnHeader_compact(is_www2) { cmnInclude("header_compact", is_www2); } // ヘッダー
function cmnFooter_compact(is_www2) { cmnInclude("footer_compact", is_www2); } // フッター

/* SGHヘッダー */
function cmnHeader_sgh(is_www2) { cmnInclude("header_sgh", is_www2); } // ヘッダー
function cmnFooter_sgh(is_www2) { cmnInclude("footer_sgh", is_www2); } // フッター

/* GOAL */
function cmnGoalNavi(is_www2) { cmnInclude("goal_navi", is_www2); }
function cmnGoalContact(is_www2) { cmnInclude("goal_contact", is_www2); }


/* !isUA -------------------------------------------------------------------- */
var isUA = (function(){
	var ua = navigator.userAgent.toLowerCase();
	indexOfKey = function(key){ return (ua.indexOf(key) != -1)? true: false;}
	var o = {};
	o.ie      = function(){ return indexOfKey("msie"); }
	o.fx      = function(){ return indexOfKey("firefox"); }
	o.chrome  = function(){ return indexOfKey("chrome"); }
	o.opera   = function(){ return indexOfKey("opera"); }
	o.android = function(){ return indexOfKey("android"); }
	o.ipad    = function(){ return indexOfKey("ipad"); }
	o.ipod    = function(){ return indexOfKey("ipod"); }
	o.iphone  = function(){ return indexOfKey("iphone"); }
	return o;
})();

/* !rollover ---------------------------------------------------------------- */
var rollover = function(){
	var suffix = { normal : '_no.', over   : '_on.'}
	$('a.over, img.over, input.over').each(function(){
		var a = null;
		var img = null;

		var elem = $(this).get(0);
		if( elem.nodeName.toLowerCase() == 'a' ){
			a = $(this);
			img = $('img',this);
		}else if( elem.nodeName.toLowerCase() == 'img' || elem.nodeName.toLowerCase() == 'input' ){
			img = $(this);
		}

		var src_no = img.attr('src');
		var src_on = src_no.replace(suffix.normal, suffix.over);

		if( elem.nodeName.toLowerCase() == 'a' ){
			a.bind("mouseover focus",function(){ img.attr('src',src_on); })
			 .bind("mouseout blur",  function(){ img.attr('src',src_no); });
		}else if( elem.nodeName.toLowerCase() == 'img' ){
			img.bind("mouseover",function(){ img.attr('src',src_on); })
			   .bind("mouseout", function(){ img.attr('src',src_no); });
		}else if( elem.nodeName.toLowerCase() == 'input' ){
			img.bind("mouseover focus",function(){ img.attr('src',src_on); })
			   .bind("mouseout blur",  function(){ img.attr('src',src_no); });
		}

		var cacheimg = document.createElement('img');
		cacheimg.src = src_on;
	});
};
/* !pageScroll -------------------------------------------------------------- */
var pageScroll = function(){
	jQuery.easing.easeInOutCubic = function (x, t, b, c, d) {
		if ((t/=d/2) < 1) return c/2*t*t*t + b;
		return c/2*((t-=2)*t*t + 2) + b;
	};
	$('a.scroll, .scroll a, .pageTop a').each(function(){
		var target = this.hash;
		if ( target && $(target).length ) {
			$(this).on("click keypress", function(e){
				e.preventDefault();
				var targetY = $(target).offset().top;
				var parent  = ( isUA.opera() )? (document.compatMode == 'BackCompat') ? 'body': 'html' : 'html,body';
				$(parent).animate(
					{scrollTop: targetY },
					400,
					'easeInOutCubic'
				);
				return false;
			});
		}
	});
}
/* !localNav ---------------------------------------------------------------- */
var localNav = function(){
	var navClass = document.body.className.toLowerCase(),
		parent = $("#lNavi01"),
		prefix = 'lNav',
		current = 'current',
		regex = {
			a  : /l/,
			dp : [
				/l[\d]+_[\d]+_[\d]+_[\d]+/,
				/l[\d]+_[\d]+_[\d]+/,
				/l[\d]+_[\d]+/,
				/l[\d]+/
			]
		},
		route = [],
		i,
		l,
		temp,
		node;

	$("ul ul", parent).hide();

	if( navClass.indexOf("ldef") >= -1 ){
		for(i = 0, l = regex.dp.length; i < l; i++){
			temp = regex.dp[i].exec( navClass );
			if( temp ){
				route[i] = temp[0].replace(regex.a, prefix);
			}
		}
		///console.log(route);
		if( route[0] ){
			// depth 4
			node = $("a."+route[0], parent);
			node.addClass(current);
			//node.next().show();
			node.parent().parent().show()
				.parent().parent().show()
				.parent().parent().show();
			node.parent().parent().prev().addClass('parent');
			node.parent().parent()
				.parent().parent().prev().addClass('parent');
			node.parent().parent()
				.parent().parent()
				.parent().parent().prev().addClass('parent');

		}else if( route[1] ){
			// depth 3
			node = $("a."+route[1], parent);
			node.addClass(current);
			node.next().show();
			node.parent().parent().show()
				.parent().parent().show();
			node.parent().parent().prev().addClass('parent');
			node.parent().parent()
				.parent().parent().prev().addClass('parent');


		}else if( route[2] ){
			// depth 2
			node = $("a."+route[2], parent);
			node.addClass(current);
			node.next().show();
			node.parent().parent().show();
			node.parent().parent().prev().addClass('parent');

		}else if( route[3] ){
			// depth 1
			node = $("a."+route[3], parent);
			node.addClass(current);
			node.next().show();

		}else{
		}
	}
}

/* ! localNav(service) --------------------------------------------------- */
var localNav02 = function(){
	var navClass = document.body.className.toLowerCase(),
		parent = $(".lNavi_service"),
		prefix = 'lNav',
		current = 'current',
		regex = {
			a  : /l/,
			dp : [
				/l[\d]+_[\d]+_[\d]+_[\d]+/,
				/l[\d]+_[\d]+_[\d]+/,
				/l[\d]+_[\d]+/,
				/l[\d]+/
			]
		},
		route = [],
		i,
		l,
		temp,
		node;

	$("ul ul", parent).hide();

	if( navClass.indexOf("ldef") >= -1 ){
		for(i = 0, l = regex.dp.length; i < l; i++){
			temp = regex.dp[i].exec( navClass );
			if( temp ){
				route[i] = temp[0].replace(regex.a, prefix);
			}
		}
		///console.log(route);
		if( route[0] ){
			// depth 4
			node = $("a."+route[0], parent);
			node.addClass(current);
			node.next().show();
			node.parent().parent().show()
				.parent().parent().show()
				.parent().parent().show();
			node.parent().parent().prev().addClass('parent');
			node.parent().parent()
				.parent().parent().prev().addClass('parent');
			node.parent().parent()
				.parent().parent()
				.parent().parent().prev().addClass('parent');

		}else if( route[1] ){
			// depth 3
			node = $("a."+route[1], parent);
			node.addClass(current);
			node.next().show();
			node.parent().parent().show()
				.parent().parent().show();
			node.parent().parent().prev().addClass('parent');
			node.parent().parent()
				.parent().parent().prev().addClass('parent');


		}else if( route[2] ){
			// depth 2
			node = $("a."+route[2], parent);
			node.addClass(current);
			node.next().show();
			node.parent().parent().show();
			node.parent().parent().prev().addClass('parent');

		}else if( route[3] ){
			// depth 1
			node = $("a."+route[3], parent);
			node.addClass(current);
			node.next().show();

		}else{
		}
	}
}

$(function(){

	$(".lNavi_service > li > a").on('click',function(){
		if($(this).is(".open")){
			$(this).next().stop(true,false).slideToggle(200);
			$(this).removeClass("open");
		}else{
			$(this).next().stop(true,false).slideToggle(200);
			$(this).addClass("open");
		}

		return false;
	}).next().hide();

	//service ローカルナビ bodyのクラス名の頭文字で開くナビを判断
	var isBodyClass = document.body.className.toLowerCase();

	if(isBodyClass.indexOf("l01_") !=-1) {
		$(".lNavi_service > li a").removeClass('open');
		$(".lNavi_service > li a.lNav01").addClass('open');
	}
	else if(isBodyClass.indexOf("l02_") !=-1) {
		$(".lNavi_service > li a").removeClass('open');
		$(".lNavi_service > li a.lNav02").addClass('open');
	}
	else if(isBodyClass.indexOf("l03_") !=-1) {
		$(".lNavi_service > li a").removeClass('open');
		$(".lNavi_service > li a.lNav03").addClass('open');
	}
	else if(isBodyClass.indexOf("l04_") !=-1) {
		$(".lNavi_service > li a").removeClass('open');
		$(".lNavi_service > li a.lNav04").addClass('open');
	}


	$('.lNavi_service > li a.open').next().show();
});

/*	!scrollTop-----------------------------------------------------------------------------------------------------------------*/
var scrollTop = (function(){
	var $pageTop = $('.pageTop');
	$pageTop.hide();
	$(window).scroll(function(){
		if ($(this).scrollTop() > 10) {
		$pageTop.fadeIn(); // スクロール量100以上のとき、フェードイン
		} else {
			$pageTop.fadeOut();
		}
	});

	// 位置揃え
	var $fib = $(".footerInnerBottom");
	fibHeight = $fib.outerHeight() - parseInt($fib.css('border-top-width'), 10);
	$pageTop.css({ bottom: fibHeight });
});

/* !Add Class --------------------------------------------------- */
var addCss = (function(){
	//fareTable
	$('.fareTable tbody tr:nth-child(even)').addClass('even');
});

/* ! 高さ揃え --------------------------------------------------- */
var tileHeight = (function(){
	$(window).on('load resize orientationchange', function(){
		var w = $(window).width();
		// PCサイズ
		if( w > 736 ) {
			// 扉パーツ
			if ( !$('body#g03').length ) {
				$('.indexBox_module').each(function(){
					var $box = $(this).find('.box_level01:visible a');
					if ( $(this).is('.wide_module') || $(this).is('.quarter_module') ) {
						$box.tile(4);
					} else if ( $(this).is('.onethird_module') ) {
						$box.tile(3);
					} else {
						$box.tile(2);
					}
				});
			}
			$('.list_innerLink01 > li').tile(2); // ページ内リンク
			$('.infoList.full li:not(.important) > *').tile(2); // お知らせ
			$('.module_col201:has(".indexBox_module") .leadBox').tile(2); // 採用
			$('#subNaviSp > li a').removeAttr('style'); // SPナビ

			$('.topIndexBox .topIndexBoxInner > .box').tile(4); // グローバル版(English,China)扉パーツ

			$('.column_module .module_col201.column_basic').tile(2);
			$('.column_module .module_col201 .column_module').tile(2);

		// SPサイズ
		} else {
			// 扉パーツ
			if ( !$('body#g03').length ) {
				$('.indexBox_module').each(function(){
					var $box = $(this).find('.box_level01:visible a');
					if ( $(this).is('.quarter_module') ) {
						$box.tile(2);
					} else {
						$box.removeAttr('style');
					}
				});
			}
			$('.list_innerLink01 > li').removeAttr('style'); // ページ内リンク
			$('.infoList.full li:not(.important) > *').removeAttr('style'); // お知らせ
			$('.module_col201:has(".indexBox_module") .leadBox').removeAttr('style'); // 採用
			$('#subNaviSp > li a').tile(2); // SPナビ

			$('.topIndexBox .topIndexBoxInner > .box').removeAttr('style'); // グローバル版(English,China)扉パーツ
		}
	});
});

/* ! 画像2カラム（可変幅） --------------------------------------------------- */
var picColumnWidth = (function(){
$(window).on('load',function() {
	//画面幅取得
	var w = $(window).width();
	//PCサイズ
	if(w > 736){
		var adImgW = $('.picture_adjust .module_col202.adjust img').attr('width');
		$('.picture_adjust .module_col202.adjust').width(adImgW);

		var adLastW = 680 - adImgW;
		$('.picture_adjust .module_col202:last-child').width(adLastW);
	}
});
});

/* ! 画像1カラム キャプション位置揃え(画像のクラスに.wAutoがついているもの限定) --------------------------------------------------- */
var picCaption01 = (function(){
	$(window).on('load resize',function() {
		var w = $(window).width();
		if(w > 736){ /* pcサイズ */

			$('.picture_module > .module_col101 .pic.wAuto').each(function(){

				var imgW = $(this).find('img').attr('width');
				$(this).parent().width(imgW); //.module_col101に値をセット
			});

			$('.picture_module > .module_col201 .pic.wAuto').each(function(){

				var img201W = $(this).find('img').attr('width');
				$(this).next('.caption').width(img201W); //.module_col201に値をセット
			});
		}

		else {
			$('.picture_module > .module_col101').removeAttr('style');
			$('.picture_module > .module_col201').removeAttr('style');
		}

	});
});


/* ! タブ切替え --------------------------------------------------- */
$(function() {
	$('.list_tabBtn01 > li').on('click',function(){
		var index = $('.list_tabBtn01 > li').index(this);

		$('.expressArea01 .expressAreaList01').hide();
		$('.expressArea01 .expressAreaList01').eq(index).fadeIn(700);

		$('.list_tabBtn01 > li').removeClass('current');

		$(this).addClass('current');

		return false;
	});

});

/* ! アコーディオン --------------------------------------------------- */
$(function(){
	$(".expressAccordionBtn01").on('click',function(){
		if($(this).is(".open")){
			$(this).next().stop(true,false).slideToggle("slow");
			$(this).removeClass("open");
			tileHeight();
		}else{
			$(this).next().stop(true,false).slideToggle("slow");
			$(this).addClass("open");
			tileHeight();
		}
	}).next('.expressAreaList01').hide();

	$('.expressAccordionBtn01:first-child').addClass("open");
	$('.expressAccordionBtn01:first-child').next('.expressAreaList01').show();
});


$(function(){
$(window).on('load resize',function(){
	var modalW = $('.modalContents').width();
	if(modalW > 300) {
		$('.modalContents').addClass('pcModal');
	}
	else {
		$('.modalContents').removeClass('pcModal');
	}
});
});

// ヘッダー
$(function () {
	$('.headerUtility').each(function () {
		var $this = $(this);
		$this.find('.sp-gNavi').append($('#gNavi').eq(0).find('li').clone(true));
		$this.find('.sp-sitemap').append($this.find('.sitemap li').clone(true));

		$this.find('>dl>dt:visible')
			.each(function () {
				$(this).data('orig-text', $(this).text());
			})
			.on('click',function () {
				var $target = $(this).next('dd');
				$this.find('>dl>dt:visible').each(function () {
					var $dt = $(this);
					var $dd = $dt.next('dd');
					if ( $target.is($dd) && $dd.is(':hidden') ) {
						$dd.stop(true).slideDown('fast');
						$dt.text($dt.data('close-text'));
						$dt.addClass('is-open');
					} else {
						$dd.stop(true).slideUp('fast');
						$dt.text($dt.data('orig-text'));
						$dt.removeClass('is-open');
					}
				});
			});
	});
});



/* ! IEのdisabled属性 --------------------------------------------------- */
$(function(){
	if ( isUA.ie() ) {
		$('input:disabled').each(function() {
			$(this).removeAttr('disabled');
			$(this).addClass('disabled');
			$(this).attr('onclick', 'javascript:return false;');
		});
	}
});

/* ! サービスのカテゴリアイコン設定 --------------------------------------------------- */
/*
2015.02.11 追加
HTML上で削除した場合、将来的にそのカテゴリのサービスが有効になった場合、HTMLの書き直しが必要になるため、JSで非表示操作をする */

var serviceCatIconSets = (function(){
	$('.pageTtl > .catIcon').each(function(){

	if( $('.cat01').hasClass('no') && $('.cat02').hasClass('no') && $('.cat03').hasClass('no') && $('.cat04').hasClass('no')){
		$(this).hide(); // すべてのカテゴリのリストにnoというclassがついている場合非表示にする
	}
	});
});


/* ! 見た目調整 --------------------------------------------------- */
$(function(){
	//.table_module01の中に.fareTableがある場合、クラスを付ける
	if($('.table_module01').find('.fareTable')){
		$('.table_module01').addClass('fareTable_module');
	}
	//.table_moduleの兄弟要素が.box_listing01だった場合
	if($('.table_module01').next('.box_listing01')){
		$('.table_module01').addClass('mbN');
	}
	//サービス　マテリアル個別ページ
	if($('.ftBox_material').find('.captionImage04')){
		$('.ftBox_material > .captionImage04').each(function(){
			$(this)
		});
	}

	$(window).on('load resize',function() {
		var w = $(window).width();
		if(w <= 736){ /* pcサイズ */

			$('.ftBox_material > .captionImage04').each(function(){

				var capImg04H = $(this).height();
				$(this).next('.detail01').css("padding-bottom", capImg04H);
			});
		}
		else {
			$('.ftBox_material > .detail01').css("padding-bottom", "0");
		}

	});
});
