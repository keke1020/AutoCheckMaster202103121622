/*global YAHOO:true, document:true, window:true, navigator:true, $:true, jQuery:true, clearInterval:true, setInterval:true */

// Set Global YAHOO
if (typeof YAHOO === "undefined") {
    YAHOO = {};
}
if (typeof YAHOO.JP === "undefined") {
    YAHOO.JP = {};
}
if (typeof YAHOO.JP.commerce === "undefined") {
    YAHOO.JP.commerce = {};
}
if (typeof YAHOO.JP.commerce.sellerfrontB === "undefined") {
    YAHOO.JP.commerce.sellerfrontB = {};
}

(function () {
    var Y = YAHOO.JP.commerce.sellerfrontB;

    /**
     * 入力フォームの追加・削除   2011.11.22
     *
     * @namespace
     * @author v 1.0.0  Masatoshi Wakizaki
     *
     * @param {Object} setting セッティング情報オブジェクト
     * @param {String} setting.setName="addInputSet" セット名（ID）
     * @param {String} [setting.max="10"] 最大追加可能数
     * @param {String} [setting.htmlArea="addInputHtml"] 項目１セットを囲ったclass
     * @param {String} [setting.btnAreaOn="addInputAreaOn"] 追加ボタン（有効）を囲ったclass
     * @param {String} [setting.btnAreaOff="addInputAreaOff"] 追加ボタン（無効）を囲ったclass
     * @param {String} [setting.btnClass="addInputBtn"] 追加ボタン（有効）のクリック部分に設定するclass
     * @param {String} [setting.delClass="addInputDelBtn"] 削除ボタンのクリック部分に設定するclass
     * @param {String} setting.repeatHtml 追加されるHTML
     *
     * @example
     *  数件追加済みでページをロードしたい時は、HTML部分にclass「addInputHtml」を含むHTMLを必要個数出力しておいてください。
     *  ■HTML
     *  <div class="tdBody" id="addInputSet">
     *  <div class="ycMdTextBtn addInputHtml"><ul class="exCfx"><li class="text"><span><input type="text" class="w300" name="test"></span></li></ul></div>
     *  <!--<div class="ycMdTextBtn"><ul class="exCfx"><li class="text"><span><input name="" type="text" class="w300"></span></li><li class="text ml8"><span><input type="checkbox" name="" id=""><label for="">この設定を削除</label></span></li></ul></div>-->
     *  <span class="blk mb5">（全角10文字）</span>
     *  <div class="ycMdBtn exCfx addInputAreaOn"><div class="btnGrS"><a class="addInputBtn" href=""><span>追加</span></a></div></div>
     *  <div class="exCfx addInputAreaOff" style="display: none;"><div class="btnDisableS"><span class="btnR"><span class="btnL">追加</span></span></div></div>
     *  </div>
     *
     *  ■JS
     *  YAHOO.JP.commerce.sellerfrontB.addInput.set({
     *      setName    : 'addInputSet',     //セット名（ID）
     *      max        : '10',              //最大追加可能数
     *      htmlArea   : 'addInputHtml',    //項目１セットを囲ったclass
     *      btnAreaOn  : 'addInputAreaOn',  //追加ボタン（有効）を囲ったclass
     *      btnAreaOff : 'addInputAreaOff', //追加ボタン（無効）を囲ったclass
     *      btnClass   : 'addInputBtn',     //追加ボタン（有効）のクリック部分に設定するclass
     *      delClass   : 'addInputDelBtn',  //削除ボタンのクリック部分に設定するclass
     *      nameClass  : 'addInputName',    //出力タグのnameを設定するclass
     *      repeatHtml : '<div class="ycMdTextBtn addInputHtml"><ul class="exCfx"><li class="text"><span class="unitTxt"><input name="" type="text" class="w60 txtRt"></span><span class="unitTxt">回払い</span></li><li class="btn ml8"><div class="ycMdBtn"><div class="btnGrS"><a href="" class="addInputDelBtn"><span>削除</span></a></div></div></li></ul></div>'          //追加されるHTML
     *  });
     */
    Y.addInput = {
        set : function (setting) {
            var mySetId = (setting && setting.setName)    ? setting.setName    : "addInputSet",
            repeatMax   = (setting && setting.max)        ? setting.max        : "10",
            htmlArea    = (setting && setting.htmlArea)   ? setting.htmlArea   : "addInputHtml",
            btnAreaOn   = (setting && setting.btnAreaOn)  ? setting.btnAreaOn  : "addInputAreaOn",
            btnAreaOff  = (setting && setting.btnAreaOff) ? setting.btnAreaOff : "addInputAreaOff",
            btnClass    = (setting && setting.btnClass)   ? setting.btnClass   : "addInputBtn",
            delClass    = (setting && setting.delClass)   ? setting.delClass   : "addInputDelBtn",
            repeatHtml  = (setting && setting.repeatHtml) ? setting.repeatHtml : "";

            if(repeatHtml === ""){
                return false;
            }

            //最大表示数判定
            var checkMax = function(){
                if($("#"+ mySetId).find("."+htmlArea).size() < repeatMax-0){
                    //追加ＯＫ
                    $("#"+ mySetId).find("."+btnAreaOn).show();
                    $("#"+ mySetId).find("."+btnAreaOff).hide();
                }else{
                    //追加ＮＧ
                    $("#"+ mySetId).find("."+btnAreaOn).hide();
                    $("#"+ mySetId).find("."+btnAreaOff).show();
                }
            };

            //一回実行
            checkMax();

            //追加ボタンイベント設定
            $("#"+ mySetId).find("."+btnClass).click(function(e){
                if (e) {
                    e.preventDefault();
                }
                // name属性を追加する為の準備
                var size       = $("#"+ mySetId).find("." + htmlArea).size() + 1;
                // name属性の連番を変更
                var repeatwork = repeatHtml.replace(/%_NO_%/g, size);

                $("#"+ mySetId).find("."+htmlArea+":last").after(repeatwork);
                checkMax();
            });
        }
    };//addInput

    /**
     * 入力フォームの内容をコピー   2011.12.02
     *
     * @namespace
     * @author v 1.0.0  Masatoshi Wakizaki
     *
     * @param {Object} setting セッティング情報オブジェクト
     * @param {String} [setting.setName="inputCopySet"] セットのclass名
     * @param {String} [setting.btnClass="inputCopyBtn"] 反映ボタンのclass名
     * @param {String} [setting.MasterClass="inputCopyMaster"] コピー元のclass名
     * @param {String} [setting.ChildClass="inputCopyChild"] コピー先のclass名
     *
     * @example
     *  全体をclass「inputCopySet」で囲い、その中に以下classを設定
     *  1.反映ボタンに「inputCopyBtn」
     *  2.コピー元のinputに「inputCopyMaster」（一つ）
     *  3.コピー先のinputに「inputCopyChild」（複数ＯＫ）
     *  このセットを同一ページ内に複数設置可能。
     *
     *  <form>
     *  <div class="inputCopySet">
     *  <table>
     *  <tr>
     *  <td>テスト</td>
     *  <td><input type="text" name="textfield" class="inputCopyChild"></td>
     *  </tr>
     *  <tr>
     *  <td>テスト</td>
     *  <td><input type="text" name="textfield2" class="inputCopyChild"></td>
     *  </tr>
     *  <tr>
     *  <td>テスト</td>
     *  <td><input type="text" name="textfield3" class="inputCopyChild"></td>
     *  </tr>
     *  </table>
     *  <input type="text" name="textfield3" class="inputCopyMaster">
     *  <input type="button" name="button" value="コピー" class="inputCopyBtn">
     *  </div>
     *  </form>
     *
     *  ■JS
     *  <script type="text/javascript">
     *  <!--
     *  (function () {
     *  	$(document).ready(function(){
     *  		YAHOO.JP.commerce.sellerfrontB.copyInput.set();
     *  	});
     *  }());
     *  -->
     *  </script>
     */
    Y.copyInput = {
        set : function (setting) {
            var mySetClass = (setting && setting.setName) ? setting.setName : "inputCopySet",
            btnClass = (setting && setting.btnClass) ? setting.btnClass : "inputCopyBtn",
            masterClass = (setting && setting.MasterClass) ? setting.MasterClass : "inputCopyMaster",
            ChildClass = (setting && setting.ChildClass) ? setting.ChildClass : "inputCopyChild";
            $("." + mySetClass).each(function(){
                var _that = $(this);
                _that.find("." + btnClass).on('click', function(e){
                    e.preventDefault();
                    var myInput = _that.find("." + masterClass).val() || "";
                    _that.find("." + ChildClass).val(myInput);
                });
            });
        }
    };//copyInput


            /**
     * モーダルボックス    2011.08.03
     *
     * @namespace
     * @version v 2.0.0  Masatoshi Wakizaki CoRe用に修正
     * @author v 1.0.0  Natsume Suzuoki
     *
     * @param {Object} setting セッティング情報オブジェクト
     * @param {String} setting.MODALBOX モーダル表示するボックスのID名（兼 識別名）
     * @param {String} [setting.OPEN=""] 起動ボタンのID
     * @param {String} [setting.CLOSE=""] 閉じる機能をもつ要素のIDまたはクラス名
     * @param {Number} [setting.SPEED=200] フェードインの速度(ms)
     * @param {String} [setting.MASK="#000"] 背景マスクの色
     * @param {Number} [setting.MASK_OP=0.5] 背景マスクの透明度
     *
     * @example
     * <!-- モーダル -->
     * <div id="PostageSet" class="ModalBox">
     * <div class="wrppr">
     * <div class="hd">
     * <h4>表示状況</h4>
     * <a href="Javascript:void(0);" class="closeBtn close">Close</a>
     * </div><!-- /.hd -->
     * <div class="bd ClearFix">
     * モーダルの中身
     * </div><!-- /.bd -->
     * </div><!-- /.wrppr -->
     * </div><!-- /#ImageViewer -->
     * <!-- モーダル -->
     * <p><a href="Javascript:void(0);" onClick="YAHOO.JP.commerce.sellerfrontB.ModalBox.show('PostageSet', this)">直接起動コードで開く</a></p>
     * <p><a href="Javascript:void(0);" id="PostageSetBtn">イベントリスナーから追加して開く</a></p>
     * <script type="text/javascript">
     * <!--
     * (function () {
     * // -----------------------------------------------------------------------------
     *     $(document).ready(function(){
     *         // モーダルボックス
     *         YAHOO.JP.commerce.sellerfrontB.ModalBox.set({
     *             MODALBOX    : 'PostageSet',    //モーダル表示するボックスのID名（兼 識別名）
     *             OPEN        : '#PostageSetBtn',    //起動ボタンのID[省略可]
     *             CLOSE        : '.close'        //閉じる機能をもつ要素のIDまたはクラス名
     *         });
     *     });
     * // -----------------------------------------------------------------------------
     * }());
     * -->
     * </script>
     */
    // モーダル表示フラグを初期化
    var modalOpenFlg = false;

    Y.ModalBox = {
        _CUSTOM : {},
        set : function (setting) {
            var _this = this,
            SETTING = setting || {},
            SetName = setting.MODALBOX;     //識別名
            //呼び出し毎に設定値を保存
            if (!this._CUSTOM[SetName]) {
                 this._CUSTOM[SetName] = {};
            }
            //データ不備があれば終了
            if (!SETTING.MODALBOX || !SETTING.CLOSE) {
                 return false;
            }
            //変数に格納
            this._CUSTOM[SetName].MODALBOX = "#" + SETTING.MODALBOX || "";
            this._CUSTOM[SetName].OPEN = SETTING.OPEN || "";
            this._CUSTOM[SetName].CLOSE = SETTING.CLOSE || "";
            this._CUSTOM[SetName].MASK_LAYER = "";
            this._CUSTOM[SetName].SPEED = SETTING.SPEED - 0 || 200;
            this._CUSTOM[SetName].MASK = SETTING.MASK || "#000";
            this._CUSTOM[SetName].MASK_OP = SETTING.MASK_OP || 0.3;

            this._CUSTOM[SetName].MASK_CSS = {
                position                : 'absolute',
                left                    : 0,
                top                    : 0,
                zIndex                : 9000,
                backgroundColor    : this._CUSTOM[SetName].MASK,
                display                : 'none'
            };
            //create mask
            if (!this._CUSTOM[SetName].MASK_LAYER) {
                this._CUSTOM[SetName].MASK_LAYER = $('<div></div>');
                this._CUSTOM[SetName].MASK_LAYER.css(this._CUSTOM[SetName].MASK_CSS);
                this._CUSTOM[SetName].MASK_LAYER.addClass("ModalMask");
                $("body").append(this._CUSTOM[SetName].MASK_LAYER);
                //mask click
                this._CUSTOM[SetName].MASK_LAYER.click(function () {
                    //_this.hide(SetName);
                });
                //IE6は背景マスクに中身が無いとクリック出来ない
                if (!jQuery.support.style && typeof document.documentElement.style.maxHeight === "undefined") {
                    //ie6
                    //this._CUSTOM[SetName].MASK_LAYER.append('<span style="display:block;height:100%">&nbsp;</span>');
                }
            }

            // キー押下時イベント
            KeyDownAction = function(event) {
                var modalId = "#" + setting.MODALBOX;
                var parentsId = $(':focus').parents(modalId);

                // 画面外からの操作、且つ、モーダル表示の場合
                if (parentsId.length == 0 && modalOpenFlg == true) {
                    doFocus(event);
                }
            };

            window.document.onkeydown = KeyDownAction;

            // フォーカス設定
            doFocus = function(event) {
               var keyEvent = event || window.event;
               var keyCode  = keyEvent.keyCode;

                // キーコードがタブ：9の場合
                if (keyCode == 9) {
                    $('.close').focus();
                }
            };

            //モーダル中身を最後に移動
            $(this._CUSTOM[SetName].MODALBOX).appendTo($("body"));

            //open
            if (this._CUSTOM[SetName].OPEN !== "") {
                $(this._CUSTOM[SetName].OPEN).on("click", function (e) {
                    e.preventDefault();
                    //イベント発生元の子要素なら親を選択(IDとclass両方に対応)
                    var t;
                    if ($(e.target).is(_this._CUSTOM[SetName].OPEN)) {
                        t = e.target;
                    } else {
                        t = $(e.target).closest(_this._CUSTOM[SetName].OPEN);
                    }
                    _this.show(SetName, t);
                });
            }
            //close
            $(this._CUSTOM[SetName].MODALBOX + " " + this._CUSTOM[SetName].CLOSE).on("click", function (e) {
                e.preventDefault();
                _this.hide(SetName);
            });
        }, /* set */

        show : function (SetName, _that) {
            modalOpenFlg = true; // モーダル表示フラグ=true
            var docW = $(document).width(),
            docH = $(document).height(),
            modalObj = $(this._CUSTOM[SetName].MODALBOX),
            maskObj = this._CUSTOM[SetName].MASK_LAYER,
            getXY = function (_Obj) {
                var myXY = {},
                scrll = $(document).scrollTop(),
                boxW = _Obj.outerWidth(),
                boxH = _Obj.outerHeight(),
                posX = $(window).width(),
                posY = $(window).height(),
                scrollLeft = document.body.scrollLeft || document.documentElement.scrollLeft;
                myXY.top = (function () {
                    var newX = (posY / 2 - boxH / 2) + scrll;
                    if (newX > scrll) {
                        return newX;
                    } else {
                        return scrll;
                    }
                })();
                myXY.left = (function () {
                    var newX = (posX / 2 - boxW / 2) + scrollLeft;
                    if (newX > 0) {
                        return newX;
                    } else {
                        return 0;
                    }
                })();
                return myXY;
            },
            myXY = getXY(modalObj);
            modalObj.css('top', myXY.top);
            modalObj.css('left', myXY.left);

            //背景マスクの位置IE6はfixedが使えないので計算
            if (!jQuery.support.style && typeof document.documentElement.style.maxHeight === "undefined") {
                //ie6
                maskObj.css({'width' : docW, 'height' : docH, 'top' : 0, 'left' : 0});
            } else {
                maskObj.css({'width' : '100%', 'height' : '100%', 'top' : '0', 'left' : '0', 'position' : 'fixed'});
            }
            maskObj.fadeTo(this._CUSTOM[SetName].SPEED, this._CUSTOM[SetName].MASK_OP);
            modalObj.fadeIn(this._CUSTOM[SetName].SPEED);
            //IE6対策 背景iframe
            //$(maskObj).bgiframe();
            //背景リサイズ
            $(window).bind("resize.ModalBox", function () {
                var myXY = getXY(modalObj);
                modalObj.css('left', myXY.left);
                Y.ModalBox._CUSTOM[SetName].MASK_LAYER.css({'width' : $(document).width(), 'height' : $(document).height()});
            });
        },/* show */

        hide : function (SetName) {
            modalOpenFlg = false; // モーダル表示フラグ=false
            this._CUSTOM[SetName].MASK_LAYER.hide();
            $(this._CUSTOM[SetName].MODALBOX).hide();
            $(window).unbind("resize.ModalBox");
        }/* hide */

    };//ModalBox


        /**
    * 送料計算フォーム（str_cartset_C-2-1.html）のtoggle  2012.02.24
    *
    * @namespace
    * @version v 1.0.0  Takashi yamakawa
    * @author v 1.0.0  Takashi yamakawa
    *
    * id名「#disableToggle」で囲う。
    * その中にtoggle用radioを配置。class名「chkDisBtn」を付与。
    * 有効化するradioに、class名「chkDisActive」を付与。
    * 銀ボタンはオン版は「.ycMdBtn」、オフ版は「.addInputAreaOff」。
    */
    Y.cartBasisToggle = {
        set: function(setting) {
            if (!($("div.ycMdCartBasis").length > 0)) { return false; }

            var _that = $("div.ycMdCartBasis"),
                        changeFnc = function() {
                                if (_that.find("input.chkDisActive").attr("checked")) {
                                        //radio：設定するの時
                                        _that.find("div.addInputAreaOff").hide();
                                        _that.find("div.ycMdBtn").show();
                                        _that.find("input:text").prop("disabled", false);

                                } else {
                                        //radio：設定しないの時
                                        _that.find("div.addInputAreaOff").show();
                                        _that.find("div.ycMdBtn").hide();
                                        _that.find("input:text").prop("disabled", "disabled");
                                }
                        };
            //radioボタン
            _that.find("input.chkDisBtn").change(function() {
                changeFnc();
            });
            //最初に一回実行
            changeFnc();

        }
    }; //cartBasisToggle

}());