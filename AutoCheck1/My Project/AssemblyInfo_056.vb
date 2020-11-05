Imports System.Net
Imports System.IO
Imports EncodingOperation
Imports System.Text.RegularExpressions

Public Class HTMLdialog
    Dim secValue As String = ""
    Dim tenpo As String() = New String() {"FKstyle", "Lucky9", "あかねYahoo", "暁", "アリス", "あかね楽天", "KuraNavi", "問よか", "雑貨倉庫"}


    Private Sub HTMLdialog_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        secValue = "html_dialog"   'iniファイルセクション

        Dim di As New System.IO.DirectoryInfo(Path.GetDirectoryName(Form1.appPath) & "\htmltemplate")
        Dim files As System.IO.FileInfo() = di.GetFiles("*.txt", System.IO.SearchOption.AllDirectories)
        ComboBox1.Items.Clear()
        For Each f As System.IO.FileInfo In files
            Dim fname As String = System.IO.Path.GetFileNameWithoutExtension(f.Name)
            ComboBox1.Items.Add(fname)
        Next

        serverListComb()

        ComboBox1.SelectedIndex = 0

        Try '動作見直し
            If Form1.inif.ReadINI(secValue, "tenpoCheck1", "") <> "" Then
                RadioButton1.Checked = Form1.inif.ReadINI(secValue, "tenpoCheck1", "")
                RadioButton2.Checked = Form1.inif.ReadINI(secValue, "tenpoCheck2", "")
                RadioButton3.Checked = Form1.inif.ReadINI(secValue, "tenpoCheck3", "")
                RadioButton4.Checked = Form1.in