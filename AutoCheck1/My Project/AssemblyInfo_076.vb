<?php

//----------------------------------------------------------------------------//
try {
    $ms_dbname = $db_name;
    $dbstr = "mysql:host=".$ms_server.";dbname=".$ms_dbname.";charset=utf8";
    $pdo = new PDO($dbstr,$ms_usrid,$ms_pass,array(PDO::ATTR_EMULATE_PREPARES => false));
} catch (PDOException $e) {
    exit('データベース接続失敗。'.$e->getMessage());
}
//$sql = "SELECT * FROM `list1` WHERE (`ID` = 1)";

//----------------------------------------------------------------------------//
//変更の際の添付ファイルリスト
if ($_REQUEST['mode'] === "changeWrite") {
    $mfiles = $_REQUEST['mfiles'];
    $mfileArray = explode(",",$mfiles);
    
    //チェックがあればファイル削除
    for ($i = 0; $i <= 5; $i++) {
        $k = $i + 1;
        $de = "delfile".$k;
        if ($_REQUEST[$de] == "yes") {
            $filename = explode("|",$mfileArray[$i]);
            $fn = "./file/".$filename[1];
            unlink($fn);
            $mfileArray[$i] = "";
        }
    }
}

//----------------------------------------------------------------------------//
//ファイル登録
if (($_REQUEST['mode'] === "newWright") or ($_REQUEST['mode'] === "changeWrite")) {
    //時間で採番する//
    $mid = time();

    $filelist = "";
    for ($i = 1; $i <= 6; $i++) {
        $formName = 'upfile'.$i;
        $tempfile = $_FILES[$for