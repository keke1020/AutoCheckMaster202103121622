var sending = false;

$(window).bind("unload",function(){}); // FireFoxのスクリプトキャッシュ対策

$(window).load(function () {
    sending = false;
});

$(document).ready(function() {
    $('input').keypress(function(ev) {
        if ((ev.which && ev.which === 13) || (ev.keyCode && ev.keyCode === 13)) {
            return false;
        } else {
            return true;
        }
    });
});

function doSubmit(form) {
    if (!sending) {
        makeHiddenParam(form);
        form.submit();
        sending = true;
    }
}

function doWindowLocation(url) {
    if (!sending) {
        window.location.href = url;
        sending = true;
    }
}

function makeHiddenParam(form) {
    var fid = form.id;
    if (fid == '') {
        fid = "d_form";
        form.id = fid;
    }
    // textボックス
    $('#' + fid + ' input[type=text]').each(function() {
        name  = $(this).attr('name');
        if ($('[name=' + name + ']').is(':disabled') == true) {
            $('<input />').attr('type', 'hidden').attr('name' , name).attr('value', $('[name=' + name + ']').val()).appendTo('#' + fid);
        }
    });

    // radioボタン
    $('#' + fid + ' input[type=radio]').each(function() {
        name  = $(this).attr('name');
        if ($(this).is(':disabled') == true && $(this).is(':checked') == true) {
            $('<input />').attr('type', 'hidden').attr('name' , name).attr('value', $('input:radio[name=' + name + ']:checked').val()).appendTo('#' + fid);
        }
    });

    // checkbox
    $('#' + fid + ' input[type=checkbox]').each(function() {
        name  = $(this).attr('name');
        if ($('[name=' + name + ']').is(':disabled') == true) {
            if (typeof $('input:checkbox[name=' + name + ']:checked').val() != "undefined") {
                $('<input />').attr('type', 'hidden').attr('name' , name).attr('value', $('input:checkbox[name=' + name + ']:checked').val()).appendTo('#' + fid);
            }
        }
    });

    // selectボックス
    $('#' + fid + ' select').each(function() {
        name  = $(this).attr('name');
        if ($('[name=' + name + ']').is(':disabled') == true) {
            $('<input />').attr('type', 'hidden').attr('name' , name).attr('value', $('select[name=' + name + ']').val()).appendTo('#' + fid);
        }
    });

    // textarea
    $('#' + fid + ' textarea').each(function() {
        name  = $(this).attr('name');
        if ($('[name=' + name + ']').is(':disabled') == true) {
            $('<input />').attr('type', 'hidden').attr('name' , name).attr('value', $('[name=' + name + ']').val()).appendTo('#' + fid);
        }
    });
}

function numberFormat(num) {
    num = zeroSuppress(num);
    return num.toString().replace( /([0-9]+?)(?=(?:[0-9]{3})+$)/g , '$1,');
}

function zeroSuppress(num) {
    return num.toString().replace( /^0+([0-9]+.*)/, '$1');
}

function doSearchByPressEnter(formName, searchButtonId) {
    $('form[name=' + formName + '] input[type=text]').keypress(function (e) {
        if ((e.which && e.which === 13) || (e.keyCode && e.keyCode === 13)) {
            $('#' + searchButtonId).trigger('click');
            return true;
        }
    });
}

(function () {
    String.prototype.escapeHtml = function(){
        var obj = document.createElement('pre');
        if (typeof obj.textContent != 'undefined') {
            obj.textContent = this;
        } else {
            obj.innerText = this;
        }
        return obj.innerHTML;
    }
}());