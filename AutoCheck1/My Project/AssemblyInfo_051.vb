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
</doc>                                                                                                                                                             

    Private Sub ToolStripComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        If ToolStripComboBox1.SelectedItem <> "選択してください" Then
            Dim folder As String = Path.GetDirectoryName(Form1.appPath) & "\formchange"
            Dim p As String = folder & "\" & ToolStripComboBox1.SelectedIt