Imports System.IO
Imports System.Text.RegularExpressions

Public Class csv_denpyo

    Dim secValue As String = ""
    Dim dgv As DataGridView = DataGridView1
    Dim sakiPath As String = ""
    Dim motoPath As String = ""

    Private Shared DenpyoKind As String = ""


    Private Sub csv_denpyo_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'sakiPath = "Z:\中村作成データ\AutoCheck\update"
        sakiPath = "\\SERVER-PC\Users\Public\program\autocheck\update"
        secValue = "csv_denpyo"   'iniファイルセクション

        dgv = DataGridView1

        templateFolderRead()
        ComboBox1.SelectedIndex = 0
        Panel4.Size = New Size(Panel4.Width, 22)
        SplitContainer1.SplitterDistance = 323

        '配送マスタ読み込み
        DataGridView6.Rows.Clear()
        Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\配送マスタ.csv"
        Dim csvRecords As New ArrayList()
        csvRecords = CSV_READ(fName)

        'ヘッダー行
        Dim header As String() = New String() {"商品コード", "商品名", "商品分類タグ", "代表商品コード", "ship-weight", "ロケーション"}
        For c As Integer = 0 To header.Length - 1
            DataGridView6.Columns.Add(c, c)
            DataGridView6.Columns(c).SortMode = DataGridViewColumnSortMode.NotSortable
      