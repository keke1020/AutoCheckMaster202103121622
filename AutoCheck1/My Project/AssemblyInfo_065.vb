﻿<?xml version="1.0"?>
<doc>
<assembly>
<name>
AutoCheck1
</name>
</assembly>
<members>
<member name="P:AutoCheck1.My.Resources.Resources.ResourceManager">
	<summary>
  このクラスで使用されているキャッシュされた ResourceManager インスタンスを返します。
</summary>
</member><member name="P:AutoCheck1.My.Resources.Resources.Culture">
	<summary>
  厳密に型指定されたこのリソース クラスを使用して、すべての検索リソースに対し、
  現在のスレッドの CurrentUICulture プロパティをオーバーライドします。
</summary>
</member><member name="T:AutoCheck1.My.Resources.Resources">
	<summary>
  ローカライズされた文字列などを検索するための、厳密に型指定されたリソース クラスです。
</summary>
</member>
</members>
</doc>                                                                                                                                                                  End Using

        RadioButton3.Checked = True
        ComboBox6.SelectedIndex = 0

        'tenpoフォルダチェック
        If Not Directory.Exists(Form1.appPath & "\tenpo") Then
            Dim di2 As System.IO.DirectoryInfo = System.IO.Directory.CreateDirectory(Path.GetDirectoryName(Form1.appPath) & "\tenpo")
        End If

        DataGridView8.Rows.Add("マスタ")
   