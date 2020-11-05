
/*  */
$(function() {
    /******************
     * カレンダー表示 *
     ******************/
    // Datepicker の初期化
    // $(".datepicker").datepicker();
    $(".datepicker").datepicker({
        showAnim: '',
        showOn: 'button',
        buttonText: '',
        buttonImageOnly: true,
        buttonImage: JS_GLOBAL_IMAGE_URL + '/common/ic_calendar.png'
    });

    $(".ui-datepicker-trigger").mouseover(function() {
        $(this).css('cursor', 'pointer');
    });

    /************************************
     * 注文詳細ページ(別ウインドウ)表示 *
     ************************************/
    openOrderDetailWindow = function(url, order_id) {
        var target = 'order_detail_view_' + order_id;
        while (target.indexOf('-',0) != -1) {
            target = target.replace('-', '_');
        }

        var orderDetailView;
        orderDetailView = window.open(url, target, 'width=1010, scrollbars=yes');
        orderDetailView.focus();
    };

    /**********************************************
     * カスタムメールプレビュー(別ウインドウ)表示 *
     **********************************************/
    openPreviewCustomMail = function(f) {
        var windowOpen;
        f.target = "hoge";
        windowOpen = window.open("about:blank", f.target, 'width=990, scrollbars=yes');
        windowOpen.focus();
        f.submit();
    };

    /*************************************
     * 検索ボタン押下(page番号を1に戻す) *
     *************************************/
    searchStart = function (url) {
        $('#page').val(1);
        document.f.action = url;
        doSubmit(document.f);
    };

    /********************************
     * 検索パッドの表示時の状態判断 *
     ********************************/
    defaultSearchPad = function(value)
    {
        if (value === 0) {
            $('#divContent').show();
            $('#searchPadShow').hide();
            $('#searchPadHide').show();
        } else {
            $('#divContent').hide();
            $('#searchPadHide').hide();
            $('#searchPadShow').show();
        }
    };

    /**************************************
     * 検索時一覧をファーストビューにする *
     **************************************/
    moveSearchResultView = function()
    {
        window.scroll(0,$('#searchResultArea').offset().top);
    }

    /***************************************
     * 注文検索 メールテンプレート設定別窓 *
     ***************************************/
    openMailSetting = function(url)
    {
        var target = 'mailSetting';
        var mailSettingView;
        mailSettingView = window.open(url, target, 'width=1010, scrollbars=yes');
        mailSettingView.focus();
    }
});
jQuery(function() {
    // 一括処理パッド(上)開閉
    // パッド Open
    $('#batchProcessPadShow').click(function() {
        $('#batchProcessContent').show();
        $('#batchProcessPadShow').hide();
        $('#batchProcessPadHide').show();
        document.getElementById("process_pad_head").value     = 1;
        document.getElementById("process_pad_headPage").value = 1;
        // aタグ無効
        return false;
    });
    // パッド Close
    $('#batchProcessPadHide').click(function() {
        $('#batchProcessContent').hide();
        $('#batchProcessPadHide').hide();
        $('#batchProcessPadShow').show();
        document.getElementById("process_pad_head").value     = 0;
        document.getElementById("process_pad_headPage").value = 0;
        // aタグ無効
        return false;
    });

    // 一括処理パッド(下)開閉
    // パッド Open
    $('#batchProcessPad2Show').click(function() {
        $('#batchProcessContent2').show();
        $('#batchProcessPad2Show').hide();
        $('#batchProcessPad2Hide').show();
        document.getElementById("process_pad_foot").value     = 1;
        document.getElementById("process_pad_footPage").value = 1;
        // aタグ無効
        return false;
    });
    // パッド Close
    $('#batchProcessPad2Hide').click(function() {
        $('#batchProcessContent2').hide();
        $('#batchProcessPad2Hide').hide();
        $('#batchProcessPad2Show').show();
        document.getElementById("process_pad_foot").value     = 0;
        document.getElementById("process_pad_footPage").value = 0;
        // aタグ無効
        return false;
    });

    // パッド Open
    $('#searchPadShow').click(function() {
        $('#divContent').show();
        $('#searchPadShow').hide();
        $('#searchPadHide').show();
        document.getElementById("SearchPadStatus").value     = 0;
        document.getElementById("SearchPadStatusPage").value = 0;
        return false;
    });
    // パッド Close
    $('#searchPadHide').click(function() {
        $('#divContent').hide();
        $('#searchPadHide').hide();
        $('#searchPadShow').show();
        document.getElementById("SearchPadStatus").value     = 1;
        document.getElementById("SearchPadStatusPage").value = 1;
        return false;
    });

    // 未選択でのエラー表示用
    $('.fakeBtn').click(function() {
        $('#notCheckId').addClass('ycMdErrMsg');
        document.getElementById("notCheckId").innerHTML = "<p><strong>処理が行えません。注文か処理ステータスが未選択です。<\/strong><\/p>";
    });

    // 未選択でのエラー表示用 注文
    $('.fakeBtnOrder').click(function() {
        $('#notCheckId').addClass('ycMdErrMsg');
        document.getElementById("notCheckId").innerHTML = "<p><strong>処理が行えません。注文が未選択です。<\/strong><\/p>";
    });

    // 未選択でのエラー表示用 注文 + 更新内容
    $('.fakeBtnOrderOrUpdate').click(function() {
        $('#notCheckId').addClass('ycMdErrMsg');
        document.getElementById("notCheckId").innerHTML = "<p><strong>処理が行えません。注文か更新内容が未選択です。<\/strong><\/p>";
    });

    // 未選択でのエラー表示用 注文 + メールテンプレート
    $('.fakeBtnOrderOrMail').click(function() {
        $('#notCheckId').addClass('ycMdErrMsg');
        document.getElementById("notCheckId").innerHTML = "<p><strong>処理が行えません。注文かメールテンプレートが未選択です。<\/strong><\/p>";
    });

    // 未選択でのエラー表示用 メールテンプレート
    $('.fakeBtnMail').click(function() {
        $('#notCheckId').addClass('ycMdErrMsg');
        document.getElementById("notCheckId").innerHTML = "<p><strong>処理が行えません。メールテンプレートが未選択です。<\/strong><\/p>";
    });


    // エラー表示削除
    $('.updateConfirmBtn').click(function() {
        $('#notCheckId').removeClass('ycMdErrMsg');
        document.getElementById("notCheckId").innerHTML = "";
    });
});