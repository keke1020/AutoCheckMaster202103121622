Imports System.IO
Imports EncodingOperation

Public Class Form1_F_kanri
    Dim listAll As String() = New String() {"アップデートファイル,\config\folder.txt",
                                            "ブックマーク,\bookmark.txt",
                                            "チェックリスト,\CheckList.txt",
                                            "お届け日数,\config\お届け日数.txt",
                                            "ログイン情報,\login.txt",
                                            "送料,\config\送料.txt",
                                            "追跡HP設定,\config\追跡HP設定.txt",
                                            "追跡番号,\config\追跡番号.txt",
                                            "会社情報,\会社情報.txt",
                                            "伝票用住所,\伝票用住所.txt",
                                            "伝票設定用,\config\denpyoConfig.txt",
                                            "コードチェック,\config\CodeCheck.txt",
                                            "強調リスト,\強調リスト.txt",
                                            "祝日リスト,\Holiday.dat"}

    Private Sub Opt_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        For i As Integer = 0 To listAll.Length - 1
            Dim str As String() = Split(listAll(i), ",")
            ListBox1.Items.Add(str(0))
        Next

    End Sub

    Private Sub ListBox1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListBox1.MouseDoubleClick
        Dim fName As String() = Split(listAll(ListBox1.SelectedIndex), ",")
        Dim fPath As String = Path.GetDirectoryName(Form1.appPath) & fName(1)
        Using sr As New System.IO.StreamReader(fPath, System.Text.Encoding.Default)
            AzukiControl1.Text = sr.ReadToEnd
        End Using
        ToolStripStatusLabel1.Text = ListBox1.SelectedIndex
    End Sub

    Private Sub 保存ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 保存ToolStripMenuItem.Click
        Dim fName As String() = Split(listAll(ToolStripStatusLabel1.Text), ",")
        Dim fPath As String = Path.GetDirectoryName(Form1.appPath) & fName(1)
        SaveTxt(fPath)
        MsgBox("「" & fName(1) & "」を保存しました", MsgBoxStyle.SystemModal)
    End Sub

    Public Sub SaveTxt(ByVal fp As String)
        ' CSVファイルオープン
        Dim sw As IO.StreamWriter = New IO.StreamWriter(fp, False, System.Text.Encoding.GetEncoding("SHIFT-JIS"))
        sw.Write(AzukiControl1.Text)
        sw.Close()
    End Sub

End Class