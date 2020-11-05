Imports System.Net
Imports System.IO
Imports System.Text.RegularExpressions

Public Class Dialog
    Private _TextBox As New List(Of TextBox)
    Dim tempoCodeArray As ArrayList = New ArrayList
    Dim tenpo As String() = New String() {"FKstyle", "Lucky9", "あかねYahoo", "暁", "アリス", "あかね楽天", "KuraNavi", "問よか", "通販の雑貨倉庫"}

    Dim csvForm As Csv

    Private Sub Dialog_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        _TextBox.Add(TextBox3)
        _TextBox.Add(TextBox5)
        _TextBox.Add(TextBox6)
        _TextBox.Add(TextBox4)
        _TextBox.Add(TextBox8)
        _TextBox.Add(TextBox7)

        'ファイル読み取り
        Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\config\CodeCheck.txt"
        Using sr As New System.IO.StreamReader(fName, System.Text.Encoding.Default)
            Do While Not sr.EndOfStream
                Dim s As String = sr.ReadLine
                tempoCodeArray.Add(s)
            Loop
        End Using

        RadioButton3.Checked = True
        ComboBox6.SelectedIndex = 0
        ComboBox8.SelectedIndex = 0
        ComboBox9.SelectedIndex = 0

        'tenpoフォルダチェック
        If Not Directory.Exists(Form1.appPath & "\tenpo") Then
            Dim di2 As System.IO.DirectoryInfo = System.IO.Directory.CreateDirectory(Path.GetDirectoryName(Form1.appPath) & "\tenpo")
        End If

        DataGridView8.Rows.Add("マスタ")
        For i As Integer = 0 To tenpo.Length - 1
            DataGridView8.Rows.Add(tenpo(i))
            'DataGridView10.Columns.Add("Column" & i, tenpo(i))
        Next

        For i As Integer = 0 To FileArray.Length - 1
            ListBox1.Items.Add(FileArray(i))
        Next

        FileGetLastTime()   'ファイルの更新時間を調べる
    End Sub

    Private Sub FileGetLastTime()
        DataGridView9.Rows.Clear()

        Dim FileArray As String() = New String() {"マスタ", "FKstyle", "Lucky9", "あかねYahoo", "通販の暁", "雑貨の国のアリス", "あかね楽天", "KuraNavi", "問屋よかろうもん", "通販の雑貨倉庫"}
        Dim folderPath As String = Path.GetDirectoryName(Form1.appPath) & "\tenpo\"

        DataGridView9.Rows.Add(FileArray(0))
        If File.Exists(folderPath & FileArray(0) & "_master.csv") = True Then
            DataGridView9.Item(1, 0).Value = File.GetLastWriteTime(folderPath & FileArray(0) & "_master.csv")
        End If

        For r As Integer = 1 To FileArray.Length - 1
            DataGridView9.Rows.Add(FileArray(r))
            If File.Exists(folderPath & FileArray(r) & "_item.csv") = True Then
                DataGridView9.Item(1, r).Value = File.GetLastWriteTime(folderPath & FileArray(r) & "_item.csv")
            End If
            If File.Exists(folderPath & FileArray(r) & "_select.csv") = True Then
                DataGridView9.Item(2, r).Value = File.GetLastWriteTime(folderPath & FileArray(r) & "_select.csv")
            End If
        Next
    End Sub

    '行番号を表示する
    Private Sub DataGridView_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles _
        DataGridView1.RowPostPaint, DataGridView2.RowPostPaint, DataGridView3.RowPostPaint, DataGridView6.RowPostPaint, DataGridView7.RowPostPaint, DataGridView10.RowPostPaint
        Dim rect As New Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, sender.RowHeadersWidth - 4, sender.Rows(e.RowIndex).Height)
        TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), sender.RowHeadersDefaultCellStyle.Font, rect, sender.RowHeadersDefaultCellStyle.ForeColor, TextFormatFlags.VerticalCenter Or TextFormatFlags.Right)
    End Sub

    '処理
    Dim nowR As Integer = 0
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        wMode = 2
        For r As Integer = 0 To DataGridView2.RowCount - 2
            DataGridView2.FirstDisplayedScrollingRowIndex = r
            If RadioButton1.Checked = True Then
                For c As Integer = 1 To DataGridView2.ColumnCount - 1
                    Dim flag As Boolean = False
                    Select Case c
                        Case 1
                            If CheckBox4.Checked = True Then
                                flag = True
                            End If
                        Case 2, 9
                            If CheckBox1.Checked = True Then
                                flag = True
                            End If
                        Case 3, 6
                            If CheckBox2.Checked = True Then
                                flag = True
                            End If
                        Case 4
                            If CheckBox3.Checked = True Then
                                flag = True
                            End If
                            'Case 7 'オークション
                            '    If CheckBox5.Checked = True Then
                            '        flag = True
                            '    End If
                        Case 7
                            If CheckBox6.Checked = True Then
                                flag = True
                            End If
                        Case 8
                            If CheckBox7.Checked = True Then
                                flag = True
                            End If
                        Case Else
                            flag = False
                    End Select
                    If flag = True Then
                        ComboBox1.SelectedIndex = c - 1
                        WebRun(r)
                        flag = False
                    End If
                Next
            Else
                WebRun(r)
            End If

            Application.DoEvents()
            'System.Threading.Thread.Sleep(500 * 1)
        Next

        MsgBox("終了しました", MsgBoxStyle.SystemModal)
    End Sub

    Private Sub WebRun(ByVal r As Integer)
        tempo = ComboBox1.SelectedIndex
        WebBrowser1.Document.GetElementsByTagName("select")("tenpo").SetAttribute("Selectedindex", tempo)
        code = DataGridView2.Item(0, r).Value
        WebBrowser1.Document.GetElementsByTagName("input")("syohin_code").SetAttribute("value", code)
        WebBrowser1.Document.Forms(0).InvokeMember("submit")

        webReadNow = True
        Do
            Application.DoEvents()
        Loop Until WebBrowser1.ReadyState = WebBrowserReadyState.Complete And WebBrowser1.IsBusy = False

        Dim html As String = WebBrowser1.Document.Body.InnerText
        If InStr(html, "商品が正しく設定されており") > 0 Then
            DataGridView2.Item(tempo + 1, r).Value = "●"
        ElseIf InStr(html, "紐づいて正しく設定されており") > 0 Then
            DataGridView2.Item(tempo + 1, r).Value = "◎"
        Else
            DataGridView2.Item(tempo + 1, r).Value = "×"
        End If

        '●＝商品コード[zk176-bk]の商品が正しく設定されており、在庫連携済みです。
        '×＝商品コード[ad009-bl]の商品がモール商品CSVアップロードで登録されていません。
        'MX＝商品コード[ad009-mi]の在庫が一度もシステムに登録されていません。
        '△＝商品コード[e007a]は、[e007-a]と紐づいて正しく設定されており、在庫連携待ちです。在庫連携が開始するまでお待ち下さい。
        '▲＝商品コード[m605-lbk]は、[m605-bk-l]と紐づいて正しく設定されており、在庫連携済みです。
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        DataGridView2.Rows.Clear()
    End Sub


    Private Sub HTMLGet(ByVal doc As HtmlAgilityPack.HtmlDocument)
        Dim res As String = ""
        Dim rArray As String() = Nothing

        Dim selecter As String = ""
        selecter = "//div[@class='alert alert-info']"
        Dim nodes As HtmlAgilityPack.HtmlNodeCollection = doc.DocumentNode.SelectNodes(selecter)
        If Not nodes Is Nothing Then
            DataGridView2.Item(2, nowR).Value = nodes(0).InnerText
        Else
            DataGridView2.Item(2, nowR).Value = "False"
        End If
    End Sub

    Private Sub コピーToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles コピーToolStripMenuItem.Click
        Dim dgv As DataGridView = DataGridView2
        Dim selCell = dgv.SelectedCells
        COPYS(dgv)
    End Sub

    Public Shared DGV1tb As TextBox
    Private Sub COPYS(ByVal dgv As DataGridView)
        If dgv.GetCellCount(DataGridViewElementStates.Selected) > 0 Then
            If dgv.CurrentCell.IsInEditMode = True Then
                DGV1tb.Copy()
            Else
                Try
                    'Clipboard.SetText(dgv.GetClipboardContent.GetText)
                    Dim returnValue As DataObject
                    returnValue = dgv.GetClipboardContent()
                    Dim strTab = returnValue.GetText()
                    Clipboard.SetDataObject(strTab, True)
                Catch ex As System.Runtime.InteropServices.ExternalException
                    MessageBox.Show("コピーできませんでした。")
                End Try
            End If
        End If
    End Sub

    Private Sub 貼り付けToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 貼り付けToolStripMenuItem.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim dgv As DataGridView = DataGridView2
        Dim selCell = dgv.SelectedCells
        PASTES(dgv, selCell)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Public ColumnChars() As String
    Private _editingColumn As Integer   '編集中のカラム番号
    Private _editingCtrl As DataGridViewTextBoxEditingControl
    Private Sub PASTES(ByVal dgv As DataGridView, ByVal selCell As Object)
        Dim x As Integer = dgv.CurrentCellAddress.X
        Dim y As Integer = dgv.CurrentCellAddress.Y
        ' クリップボードの内容を取得
        Dim clipText As String = Clipboard.GetText()
        ' 改行を変換
        clipText = clipText.Replace(vbCrLf, vbLf)
        clipText = clipText.Replace(vbCr, vbLf)
        ' 改行で分割
        Dim lines() As String = clipText.Split(vbLf)

        Dim oneFlag As Boolean = False
        If lines.Length = 1 Then
            Dim vals() As String = lines(0).Split(vbTab)
            If vals.Length = 1 Then
                oneFlag = True
            End If
        End If

        If oneFlag = True Then
            If dgv.IsCurrentCellInEditMode = False Then
                dgv(selCell(0).ColumnIndex, selCell(0).RowIndex).Value = lines(0)
            Else
                'datagridviewのセルをテキストボックスとして扱い
                '選択文字列を置き換えればpasteになる
                Dim TextEditCtrl As DataGridViewTextBoxEditingControl = DirectCast(dgv.EditingControl, DataGridViewTextBoxEditingControl)
                TextEditCtrl.SelectedText = lines(0)
                'dgv(selCell(0).ColumnIndex, selCell(0).RowIndex).Value = Replace(dgv(selCell(0).ColumnIndex, selCell(0).RowIndex).Value, TextEditCtrl.SelectedText, lines(0))
                'dgv.EndEdit()
            End If
        Else
            Dim r As Integer
            Dim nflag As Boolean = True
            For r = 0 To lines.GetLength(0) - 1
                If dgv Is DataGridView3 Then    'コードチェック欄では空欄は追加しない
                    If lines(r) = "" Then
                        Continue For
                    End If
                End If

                ' 最後のNULL行をコピーするかどうか
                If r >= lines.GetLength(0) - 1 And
                   "".Equals(lines(r)) And nflag = False Then Exit For
                If "".Equals(lines(r)) = False Then nflag = False

                ' タブで分割
                Dim vals() As String = lines(r).Split(vbTab)

                ' 各セルの値を設定
                Dim c As Integer = 0
                Dim c2 As Integer = 0
                For c = 0 To vals.GetLength(0) - 1
                    ' セルが存在しなければ貼り付けない
                    If Not (x + c2 >= 0 And x + c2 < dgv.ColumnCount And
                            y + r >= 0 And y + r < dgv.RowCount) Then
                        Continue For
                    End If
                    ' 非表示セルには貼り付けない
                    If dgv(x + c2, y + r).Visible = False Then
                        c = c - 1
                        Continue For
                    End If
                    '' 貼り付け処理(入力可能文字チェック無しの時)------------
                    '' 行追加モード&(最終行の時は行追加)
                    'If y + r = dgv.RowCount - 1 And _
                    '   dgv.AllowUserToAddRows = True Then
                    '    dgv.RowCount = dgv.RowCount + 1
                    'End If
                    '' 貼り付け
                    'dgv(x + c2, y + r).Value = vals(c)
                    ' ------------------------------------------------------
                    ' 貼り付け処理(入力可能文字チェック有りの時)------------
                    Dim pststr As String = ""
                    For i As Long = 0 To vals(c).Length - 1
                        _editingColumn = x + c2
                        Dim tmpe As KeyPressEventArgs =
                            New KeyPressEventArgs(vals(c).Substring(i, 1)) With {
                            .Handled = False
                            }
                        DataGridView_CellKeyPress(dgv, tmpe)
                        If tmpe.Handled = False Then
                            pststr = pststr & vals(c).Substring(i, 1)
                        End If
                    Next
                    ' 行追加モード＆最終行の時は行追加
                    If y + r = dgv.RowCount - 1 And
                       dgv.AllowUserToAddRows = True Then
                        dgv.RowCount = dgv.RowCount + 1
                    End If
                    ' 貼り付け
                    dgv(x + c2, y + r).Value = pststr
                    ' ------------------------------------------------------
                    ' 次のセルへ
                    c2 = c2 + 1
                Next
            Next
        End If
    End Sub

    Private Sub DataGridView_CellKeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs)
        ' カラムへの入力可能文字を指定するための配列が指定されているかチェック
        If IsArray(ColumnChars) Then
            ' カラムへの入力可能文字を指定するための配列数チェック
            If ColumnChars.GetLength(0) - 1 >= _editingColumn Then
                ' カラムへの入力可能文字が指定されているかチェック
                If ColumnChars(_editingColumn) <> "" Then
                    ' カラムへの入力可能文字かチェック
                    If InStr(ColumnChars(_editingColumn), e.KeyChar) <= 0 AndAlso
                       e.KeyChar <> Chr(Keys.Back) Then
                        ' カラムへの入力可能文字では無いので無効
                        e.Handled = True
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub DataGridView_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DataGridView2.KeyUp
        Dim selCell = DataGridView2.SelectedCells

        If e.Modifiers = Keys.Control And e.KeyCode = Keys.Enter Then '複数セル一括入力
            '-----------------------------------------------------------------------
            PreDIV()
            '-----------------------------------------------------------------------
            Dim str As String = selCell(0).Value
            For i As Integer = 0 To selCell.Count - 1
                selCell(i).Value = str
            Next
            '-----------------------------------------------------------------------
            RetDIV()
            '-----------------------------------------------------------------------
        End If
    End Sub

    Dim cellChange As Boolean = True
    Dim sc As Integer = 0
    Dim sr As Integer = 0
    Public Sub PreDIV()
        '-----------------------------------------------------------------------
        sc = DataGridView2.HorizontalScrollingOffset
        sr = DataGridView2.FirstDisplayedScrollingRowIndex
        cellChange = False
        DataGridView2.ScrollBars = ScrollBars.None
        DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        DataGridView2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
        DataGridView2.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        DataGridView2.ScrollBars = ScrollBars.Both
        '-----------------------------------------------------------------------
    End Sub

    Public Sub RetDIV()
        '-----------------------------------------------------------------------
        DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        cellChange = True
        '***************************************************
        DataGridView2.HorizontalScrollingOffset = sc
        '***************************************************
        If sr > DataGridView2.RowCount - 2 Then
            sr = DataGridView2.RowCount - 2
        End If
        If sr < 0 Then
            sr = 0
        End If
        DataGridView2.FirstDisplayedScrollingRowIndex = sr
        '-----------------------------------------------------------------------
    End Sub

    Dim wMode As Integer = 0
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        wMode = 1
        WebBrowser1.Navigate("https://ne01.next-engine.com/usertools/checkzaikosetting")
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        WebBrowser1.Navigate("https://base.next-engine.org/apps/launch/?id=65459")
    End Sub

    Dim webReadNow As Boolean = False
    Dim tempo As Integer = 0
    Dim code As String = ""
    Private Sub WebBrowser1_DocumentCompleted(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted
        Try
            Select Case wMode
                Case 1
                    WebBrowser1.Document.GetElementsByTagName("input")("user_login_code").SetAttribute("value", "ping@yongshun006.com")
                    WebBrowser1.Document.GetElementsByTagName("input")("user_password").SetAttribute("value", "ph_123456")
                Case 2

            End Select
        Catch ex As Exception

        End Try

        webReadNow = False
    End Sub


    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        Dim dgv As DataGridView = DataGridView3
        Dim selCell = dgv.SelectedCells
        COPYS(dgv)
    End Sub

    Private Sub ToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem2.Click
        Dim dgv As DataGridView = DataGridView3
        Dim selCell = dgv.SelectedCells
        PASTES(dgv, selCell)
    End Sub

    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        RadioButton_Change("Yahoo")
        Label2.Text = "1、data"
        Label3.Text = "2、なし"
        TextBox2.Enabled = False
        Button5.Enabled = False
    End Sub

    Private Sub RadioButton4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton4.CheckedChanged
        RadioButton_Change("楽天")
        Label2.Text = "1、item"
        Label3.Text = "2、select"
        TextBox2.Enabled = True
        Button5.Enabled = True
    End Sub

    Private Sub RadioButton5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton5.CheckedChanged
        RadioButton_Change("Wowma")
        Label2.Text = "1、data"
        Label3.Text = "2、なし"
        TextBox2.Enabled = False
        Button5.Enabled = False
    End Sub

    Private Sub RadioButton6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton6.CheckedChanged
        RadioButton_Change("Amazon")
        Label2.Text = "1、出品"
        Label3.Text = "2、なし"
        TextBox2.Enabled = False
        Button5.Enabled = False
    End Sub

    Private Sub RadioButton7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton7.CheckedChanged
        RadioButton_Change("Qoo10")
        Label2.Text = "1、item"
        Label3.Text = "2、なし"
        TextBox2.Enabled = False
        Button5.Enabled = False
    End Sub

    Private Sub RadioButton8_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton8.CheckedChanged
        RadioButton_Change("ヤフオク")
        Label2.Text = "1、item"
        Label3.Text = "2、select"
        TextBox2.Enabled = True
        Button5.Enabled = True
    End Sub

    Private Sub RadioButton_Change(ByVal tempoName As String)
        For i As Integer = 0 To tempoCodeArray.Count - 1
            Dim tArray As String() = Split(tempoCodeArray(i), "=")
            If tArray(0) = tempoName Then
                Dim sArray As String() = Split(tArray(1), ",")
                For t As Integer = 0 To 5
                    If sArray(t) <> "" Then
                        _TextBox(t).Enabled = True
                    Else
                        _TextBox(t).Enabled = False
                    End If
                    _TextBox(t).Text = sArray(t)
                Next

                Exit For
            End If
        Next
    End Sub

    Private Sub DataGridView3_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles DataGridView3.DragDrop
        Dim dataPathArray As ArrayList = New ArrayList
        For Each filename As String In e.Data.GetData(DataFormats.FileDrop)
            dataPathArray.Add(filename)
        Next
        For Each filename As String In dataPathArray
            Dim dgv As DataGridView = DataGridView3
            Dim csvRecords As New ArrayList()
            csvRecords = CSV_READ(filename)
            For r As Integer = 0 To csvRecords.Count - 1
                Dim sArray As String() = Split(csvRecords(r), "|=|")
                If dgv.ColumnCount <= sArray.Length - 1 Then
                    For c As Integer = dgv.ColumnCount To sArray.Length - 1
                        dgv.Columns.Add(c, c)
                        dgv.Columns(c).SortMode = DataGridViewColumnSortMode.NotSortable
                    Next
                End If
                dgv.Rows.Add(sArray)
            Next

            Exit For
        Next
    End Sub

    Private Sub DataGridView3_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles DataGridView3.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Dim cArray1 As ArrayList = New ArrayList    'ハイフン入りコード
    Dim cArray2 As ArrayList = New ArrayList    'ハイフン無しコード（横縦順）
    Dim cArray3 As ArrayList = New ArrayList    'ハイフン無しコード（縦横順）
    Dim cArrayAri As ArrayList = New ArrayList  '入力済リスト
    Dim cArrayZan As ArrayList = New ArrayList  '残り抽出用
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        FileRead(TextBox1, DataGridView4)

        Dim numName As ArrayList = New ArrayList From {
            TextBox3.Text,
            TextBox5.Text,
            TextBox6.Text
        }
        CArrayCreate(DataGridView4, numName)
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        FileRead(TextBox2, DataGridView5)

        Dim numName As ArrayList = New ArrayList From {
            TextBox4.Text,
            TextBox8.Text,
            TextBox7.Text
        }
        CArrayCreate(DataGridView5, numName)
    End Sub

    Private Sub CArrayCreate(ByVal dgv As DataGridView, ByVal numName As ArrayList)
        Dim mode As Integer = 0
        If RadioButton3.Checked Then    'Yahoo
            mode = 1
        ElseIf RadioButton4.Checked Then    '楽天
            mode = 0
        ElseIf RadioButton5.Checked Then    'Wowma
            mode = 2
        ElseIf RadioButton6.Checked Then    'Amazon
            mode = 3
        ElseIf RadioButton7.Checked Then    'Qoo10
            mode = 4
        End If

        Dim colNum1 As Integer = -1
        Dim colNum2 As Integer = -1
        Dim colNum3 As Integer = -1
        For c As Integer = 0 To dgv.ColumnCount - 1
            If numName(1) <> "" Then
                If dgv.Item(c, 0).Value = numName(0) Then
                    colNum1 = c
                ElseIf dgv.Item(c, 0).Value = numName(1) Then
                    colNum2 = c
                ElseIf dgv.Item(c, 0).Value = numName(2) Then
                    colNum3 = c
                End If
            Else
                If DataGridView4.Item(c, 0).Value = numName(0) Then
                    colNum1 = c
                    Exit For
                End If
            End If
        Next
        If colNum1 = -1 Then
            MsgBox("「" & numName(0) & "」が見つかりません", MsgBoxStyle.SystemModal)
            Exit Sub
        Else
            If numName(1) <> "" And numName(2) <> "" Then   'メインと横軸・縦軸の時
                Select Case mode
                    Case 0  '楽天専用
                        For r As Integer = 1 To dgv.RowCount - 1
                            If dgv.Item(colNum2, r).Value <> "" Then
                                Dim str As String = dgv.Item(colNum1, r).Value & dgv.Item(colNum2, r).Value & dgv.Item(colNum3, r).Value
                                If cArray1.Contains(str) = False Then   '重複登録しない
                                    cArray1.Add(str)
                                    str = Replace(str, "-", "")
                                    str = str.ToLower()
                                    cArray2.Add(str)
                                    str = dgv.Item(colNum1, r).Value & dgv.Item(colNum3, r).Value & dgv.Item(colNum2, r).Value
                                    str = Replace(str, "-", "")
                                    cArray3.Add(str)
                                End If
                            End If
                        Next
                        '-------------------------------------------------------------------
                    Case 2  'Wowma専用
                        For r As Integer = 1 To dgv.RowCount - 1
                            Dim oyaCode As String = dgv.Item(colNum1, r).Value
                            If dgv.Item(colNum2, r).Value = "" Then     '枝番号が無い時
                                If cArray1.Contains(oyaCode) = False Then   '重複登録しない
                                    Dim str As String = oyaCode
                                    cArray1.Add(str)
                                    str = Replace(str, "-", "")
                                    str = str.ToLower()
                                    cArray2.Add(str)
                                    str = Replace(str, "-", "")
                                    cArray3.Add(str)
                                End If
                            Else    '枝番号がある時 SkuOptionTypeCol、SkuOptionTypeRowに「:」区切りで枝番号が入っている
                                Dim aArray1 As String() = Split(dgv.Item(colNum2, r).Value, ":")
                                Dim aArray2 As String() = Split(dgv.Item(colNum3, r).Value, ":")
                                For i As Integer = 0 To aArray1.Length - 1
                                    For k As Integer = 0 To aArray2.Length - 1
                                        Dim edaCode1 As String = oyaCode & aArray1(i) & aArray2(k)
                                        If cArray1.Contains(edaCode1) = False Then   '重複登録しない
                                            cArray1.Add(edaCode1)
                                            edaCode1 = Replace(edaCode1, "-", "")
                                            edaCode1 = edaCode1.ToLower()
                                            cArray2.Add(edaCode1)
                                            Dim edaCode2 As String = oyaCode & aArray2(k) & aArray1(i)
                                            edaCode2 = Replace(edaCode2, "-", "")
                                            edaCode2 = edaCode2.ToLower()
                                            cArray3.Add(edaCode2)
                                        End If
                                    Next
                                Next
                            End If
                        Next
                        '-------------------------------------------------------------------
                End Select
            ElseIf numName(1) <> "" And numName(2) = "" Then    'メインとサブだけの時
                Select Case mode
                    Case 1  'Yahoo専用
                        For r As Integer = 1 To dgv.RowCount - 1
                            If dgv.Item(colNum2, r).Value = "" Then     '枝番号が無い時
                                Dim str As String = dgv.Item(colNum1, r).Value
                                If cArray1.Contains(str) = False Then   '重複登録しない
                                    cArray1.Add(str)
                                    str = Replace(str, "-", "")
                                    str = str.ToLower()
                                    cArray2.Add(str)
                                    str = Replace(str, "-", "")
                                    cArray3.Add(str)
                                End If
                            Else    '枝番号がある時 sub-codeに「選択肢=コード&選択肢=コード」
                                Dim aArray1 As String() = Split(dgv.Item(colNum2, r).Value, "&")
                                For i As Integer = 0 To aArray1.Length - 1
                                    Dim aArray2 As String() = Split(aArray1(i), "=")
                                    Dim str As String = aArray2(1)
                                    If cArray1.Contains(str) = False Then   '重複登録しない
                                        cArray1.Add(str)
                                        str = Replace(str, "-", "")
                                        str = str.ToLower()
                                        cArray2.Add(str)
                                        str = Replace(str, "-", "")
                                        cArray3.Add(str)
                                    End If
                                Next
                            End If
                        Next
                        '-------------------------------------------------------------------
                    Case 4  'Qoo10専用
                        For r As Integer = 1 To dgv.RowCount - 1
                            If dgv.Item(colNum2, r).Value = "" Then     '枝番号が無い時
                                Dim str As String = dgv.Item(colNum1, r).Value
                                If cArray1.Contains(str) = False Then   '重複登録しない
                                    cArray1.Add(str)
                                    str = Replace(str, "-", "")
                                    str = str.ToLower()
                                    cArray2.Add(str)
                                    str = Replace(str, "-", "")
                                    cArray3.Add(str)
                                End If
                            Else    '枝番号がある時 Inventory Infoに「横軸||*縦軸||*0.00||*在庫||*コード$$」
                                Dim aArray1 As String() = Split(dgv.Item(colNum2, r).Value, "$$")
                                For i As Integer = 0 To aArray1.Length - 1
                                    Dim aArray2 As String() = Split(aArray1(i), "||*")
                                    Dim str As String = aArray2(aArray2.Length - 1)
                                    If str = "" Then
                                        str = dgv.Item(colNum1, r).Value    '枝番号が登録されていない場合がある
                                    End If
                                    If cArray1.Contains(str) = False Then   '重複登録しない
                                        cArray1.Add(str)
                                        str = Replace(str, "-", "")
                                        str = str.ToLower()
                                        cArray2.Add(str)
                                        str = Replace(str, "-", "")
                                        cArray3.Add(str)
                                    End If
                                Next
                            End If
                        Next
                        '-------------------------------------------------------------------
                End Select
            Else                                            'メインだけの時
                cArray1.Clear()
                cArray2.Clear()
                cArray3.Clear()

                For r As Integer = 1 To dgv.RowCount - 1
                    Dim str As String = dgv.Item(colNum1, r).Value
                    If cArray1.Contains(str) = False Then   '重複登録しない
                        cArray1.Add(str)
                        str = Replace(str, "-", "")
                        str = str.ToLower()
                        cArray2.Add(str)
                        str = Replace(str, "-", "")
                        cArray3.Add(str)
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        cArrayZan = cArray1.Clone   'シャローコピー（ディープコピーすると連動してしまう）

        Dim searchW As String = ""
        Dim deleteW As String = ""
        For r As Integer = 0 To DataGridView3.RowCount - 2
            If DataGridView3.Item(0, r).Value <> "" Then
                Dim flag As Boolean = False
                searchW = DataGridView3.Item(0, r).Value
                If cArray1.Contains(searchW) Then
                    DataGridView3.Item(1, r).Value = cArray1(cArray1.IndexOf(searchW))
                    cArrayAri.Add(cArray1(cArray1.IndexOf(searchW)))
                    deleteW = cArray1(cArray1.IndexOf(searchW))
                    flag = True
                Else
                    searchW = Replace(DataGridView3.Item(0, r).Value, "-", "")
                    searchW = searchW.ToLower()
                    If cArray2.Contains(searchW) = True Then
                        DataGridView3.Item(2, r).Value = cArray1(cArray2.IndexOf(searchW))
                        cArrayAri.Add(cArray1(cArray2.IndexOf(searchW)))
                        deleteW = cArray1(cArray2.IndexOf(searchW))
                        flag = True
                    Else
                        If cArray3.Contains(searchW) = True Then
                            DataGridView3.Item(2, r).Value = cArray1(cArray3.IndexOf(searchW))
                            cArrayAri.Add(cArray1(cArray3.IndexOf(searchW)))
                            deleteW = cArray1(cArray3.IndexOf(searchW))
                            flag = True
                        End If
                    End If
                End If

                If flag = True Then '重複があるとnumがマイナスになる
                    Dim num As Integer = cArrayZan.IndexOf(deleteW)
                    If num >= 0 Then
                        cArrayZan.RemoveAt(num)
                    Else
                        'MsgBox(deleteW & vbCrLf & num, MsgBoxStyle.SystemModal)
                    End If
                End If
            End If
        Next

        '残リストから代表コード等が重複しているものを削除する
        Dim cArrayZan2 As ArrayList = New ArrayList
        Dim flag2 As Boolean = False
        For i As Integer = 0 To cArrayZan.Count - 1
            For k As Integer = 0 To cArrayAri.Count - 1
                If InStr(cArrayAri(k), cArrayZan(i)) > 0 Then
                    flag2 = True
                    Exit For
                End If
            Next
            If flag2 = False Then
                cArrayZan2.Add(cArrayZan(i))
            Else
                flag2 = False
            End If
        Next

        cArrayZan.Clear()
        cArrayZan = cArrayZan2.Clone

        TextBox9.Text = cArrayZan.Count
        MsgBox("終了しました", MsgBoxStyle.SystemModal)

        'Yahoo
        '　codeに代表コード、sub-codeに「選択肢=コード&選択肢=コード」で入っている
        '楽天
        '　dl-itemファイルの、商品番号に代表コード
        '　dl-selectファイルの、商品管理番号（商品URL）に代表コード、項目選択肢別在庫用横軸選択肢子番号と項目選択肢別在庫用縦軸選択肢子番号に子番号が入っている
    End Sub

    Private Sub FileRead(ByVal tbox As TextBox, ByVal dgv As DataGridView)
        Dim ofd As New OpenFileDialog With {
                    .AutoUpgradeEnabled = False,
            .RestoreDirectory = True
        }
        If ofd.ShowDialog() = DialogResult.OK Then
            dgv.Rows.Clear()

            tbox.Text = ofd.FileName
            Dim csvRecords As New ArrayList()
            csvRecords = CSV_READ(ofd.FileName)
            For r As Integer = 0 To csvRecords.Count - 1
                Dim sArray As String() = Split(csvRecords(r), "|=|")
                If dgv.ColumnCount <= sArray.Length - 1 Then
                    For c As Integer = dgv.ColumnCount To sArray.Length - 1
                        dgv.Columns.Add(c, c)
                        dgv.Columns(c).SortMode = DataGridViewColumnSortMode.NotSortable
                    Next
                End If
                dgv.Rows.Add(sArray)
            Next
        End If
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim str As String = ""
        For i As Integer = 0 To cArrayZan.Count - 1
            str &= cArrayZan(i) & vbCrLf
        Next
        Dim fp As String = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) & "\残コード一覧.txt"
        Dim sw As IO.StreamWriter = New IO.StreamWriter(fp, False, System.Text.Encoding.GetEncoding("SHIFT-JIS"))
        sw.Write(str)
        sw.Close()
    End Sub

    Private Function CSV_READ(ByRef path As String)
        Dim csvRecords As New ArrayList()

        'CSVファイル名
        Dim csvFileName As String = path

        'Shift JISで読み込む
        Dim tfp As New FileIO.TextFieldParser(csvFileName, System.Text.Encoding.GetEncoding(932)) With {
            .TextFieldType = FileIO.FieldType.Delimited,
            .Delimiters = New String() {","},
            .HasFieldsEnclosedInQuotes = True,
            .TrimWhiteSpace = True
        }

        While Not tfp.EndOfData
            'フィールドを読み込む
            Dim fields As String() = tfp.ReadFields()
            Dim str As String = ""
            For i As Integer = 0 To fields.Length - 1
                If i = 0 Then
                    str = fields(i)
                Else
                    str &= "|=|" & fields(i)
                End If
            Next
            '保存
            csvRecords.Add(str)
        End While

        '後始末
        tfp.Close()
        Return csvRecords
    End Function

    Private Sub フォームのリセットToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles フォームのリセットToolStripMenuItem.Click
        Dim frm As Form = New Dialog
        Me.Dispose()
        frm.Show()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        DataGridView3.Rows.Clear()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        For r As Integer = 0 To DataGridView3.RowCount - 2
            DataGridView3.Item(1, r).Value = ""
            DataGridView3.Item(2, r).Value = ""
        Next
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Dim img As String = "http(?<text>.*?).*>"
        Dim DLcount As Integer = 0
        Dim fol As String = ""

        DataGridView6.Rows.Clear()
        NumericUpDown1.Value = 0

        Dim htmlStr As String = csvForm.DataGridView1.SelectedCells(0).Value
        Dim htmlLines As String() = Split(htmlStr, " ")
        For k As Integer = 0 To htmlLines.Length - 1
            Select Case True
                Case InStr(htmlLines(k), ".jpg"), InStr(htmlLines(k), ".gif")
                    DataGridView6.Rows.Add(htmlLines(k))
                    NumericUpDown1.Value += 1
            End Select
        Next
        NumericUpDown1.Value -= 1

        Dim DGVheaderCheck As ArrayList = New ArrayList
        For c As Integer = 0 To csvForm.DataGridView1.Columns.Count - 1
            DGVheaderCheck.Add(csvForm.DataGridView1.Item(c, 0).Value)
        Next c
        Dim r1 As Integer = csvForm.DataGridView1.SelectedCells(0).RowIndex
        TextBox13.Text = csvForm.DataGridView1.Item(DGVheaderCheck.IndexOf(ComboBox4.SelectedItem), r1).Value
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Dim DGVheaderCheck As ArrayList = New ArrayList
        For c As Integer = 0 To csvForm.DataGridView1.Columns.Count - 1
            DGVheaderCheck.Add(csvForm.DataGridView1.Item(c, 0).Value)
        Next c

        Dim res1 As String = ""
        Dim res2 As String = ""
        For i As Integer = 0 To NumericUpDown1.Value
            Dim ext As String = ".jpg"
            Try
                If InStr(DataGridView6.Item(0, i).Value, ".gif") Then
                    ext = ".gif"
                ElseIf InStr(DataGridView6.Item(0, i).Value, ".png") Then
                    ext = ".png"
                End If
            Catch ex As Exception
                ext = ".jpg"
            End Try
            If i = 0 Then
                res1 = TextBox11.Text & TextBox12.Text & "/" & TextBox13.Text & ext
                res2 = TextBox13.Text
            Else
                res1 &= " " & TextBox11.Text & TextBox12.Text & "/" & TextBox13.Text & "-" & i & ext
                res2 &= " " & TextBox13.Text
            End If
        Next

        '追加分
        If CheckBox9.Checked = True Then
            res1 &= " " & TextBox11.Text & TextBox10.Text
            res2 &= " " & TextBox13.Text
        End If
        If CheckBox10.Checked = True Then
            res1 &= " " & TextBox11.Text & TextBox14.Text
            res2 &= " " & TextBox13.Text
        End If
        If CheckBox11.Checked = True Then
            res1 &= " " & TextBox11.Text & TextBox18.Text
            res2 &= " " & TextBox13.Text
        End If
        If CheckBox13.Checked = True Then
            res1 &= " " & TextBox11.Text & TextBox20.Text
            res2 &= " " & TextBox13.Text
        End If
        If CheckBox12.Checked = True Then
            res1 &= " " & TextBox11.Text & TextBox19.Text
            res2 &= " " & TextBox13.Text
        End If

        Dim r1 As Integer = csvForm.DataGridView1.SelectedCells(0).RowIndex
        Dim head1 As String = ComboBox2.SelectedItem
        csvForm.DataGridView1.Item(DGVheaderCheck.IndexOf(head1), r1).Value = res1
        Dim head2 As String = ComboBox3.SelectedItem
        csvForm.DataGridView1.Item(DGVheaderCheck.IndexOf(head2), r1).Value = res2

        CheckBox9.Checked = False
        CheckBox10.Checked = False
        CheckBox11.Checked = False
        CheckBox13.Checked = False
        CheckBox12.Checked = False
    End Sub

    Private Sub DataGridView6_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView6.SelectionChanged
        If DataGridView6.SelectedCells.Count > 0 Then
            Dim url As String = DataGridView6.SelectedCells(0).Value
            TextBox17.Text = Path.GetFileNameWithoutExtension(url)
        End If
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        Dim fbd As New FolderBrowserDialog With {
            .Description = "フォルダを指定してください。",
            .RootFolder = Environment.SpecialFolder.Desktop,
            .SelectedPath = "",
            .ShowNewFolderButton = True
        }

        If fbd.ShowDialog(Me) = DialogResult.OK Then
            TextBox15.Text = fbd.SelectedPath
        End If
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        Dim fbd As New FolderBrowserDialog With {
            .Description = "フォルダを指定してください。",
            .RootFolder = Environment.SpecialFolder.Desktop,
            .SelectedPath = "",
            .ShowNewFolderButton = True
        }

        If fbd.ShowDialog(Me) = DialogResult.OK Then
            TextBox16.Text = fbd.SelectedPath
        End If
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        For r As Integer = 0 To DataGridView6.RowCount - 1
            Try
                Dim fileName1 As String = Path.GetFileName(DataGridView6.Item(0, r).Value)
                Dim fileName2 As String = Path.GetFileName(DataGridView6.Item(0, r).Value)
                Dim ext As String = ".jpg"
                If InStr(fileName2, ".gif") Then
                    ext = ".gif"
                ElseIf InStr(fileName2, ".png") Then
                    ext = ".png"
                End If
                If CheckBox8.Checked = True Then
                    If r = 0 Then
                        fileName2 = TextBox13.Text & ext
                    Else
                        fileName2 = TextBox13.Text & "-" & r & ext
                    End If
                End If
                Dim stFilePathes As String() = GetFilesMostDeep(TextBox15.Text, fileName1)
                For Each stFilePath As String In stFilePathes
                    File.Copy(stFilePath, TextBox16.Text & "\" & fileName2)
                Next stFilePath
            Catch ex As Exception

            End Try
        Next
    End Sub

    Public Shared Function GetFilesMostDeep(ByVal stRootPath As String, ByVal stPattern As String) As String()
        Dim hStringCollection As New System.Collections.Specialized.StringCollection()

        ' このディレクトリ内のすべてのファイルを検索する
        For Each stFilePath As String In System.IO.Directory.GetFiles(stRootPath, stPattern)
            hStringCollection.Add(stFilePath)
        Next stFilePath

        ' このディレクトリ内のすべてのサブディレクトリを検索する (再帰)
        For Each stDirPath As String In System.IO.Directory.GetDirectories(stRootPath)
            Dim stFilePathes As String() = GetFilesMostDeep(stDirPath, stPattern)

            ' 条件に合致したファイルがあった場合は、ArrayList に加える
            If Not stFilePathes Is Nothing Then
                hStringCollection.AddRange(stFilePathes)
            End If
        Next stDirPath

        ' StringCollection を 1 次元の String 配列にして返す
        Dim stReturns As String() = New String(hStringCollection.Count - 1) {}
        hStringCollection.CopyTo(stReturns, 0)

        Return stReturns
    End Function

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        Dim DGVheaderCheck As ArrayList = New ArrayList
        For c As Integer = 0 To csvForm.DataGridView1.Columns.Count - 1
            DGVheaderCheck.Add(csvForm.DataGridView1.Item(c, 0).Value)
        Next c

        Dim sakiC As Integer = DGVheaderCheck.IndexOf(ComboBox5.SelectedItem)
        Dim motoC As Integer = DGVheaderCheck.IndexOf(ComboBox2.SelectedItem)
        For r As Integer = 0 To csvForm.DataGridView1.RowCount - 1
            Dim res As String = ""
            If CheckBox14.Checked = True Then
                res &= "<p>" & vbLf
            End If
            Dim rArray As String() = Split(csvForm.DataGridView1.Item(motoC, r).Value, " ")
            For i As Integer = 0 To rArray.Length - 1
                If CheckBox15.Checked = True Then
                    If InStr(rArray(i), "other") = 0 Then
                        Dim url As String = Replace(TextBox28.Text, "[画像URL]", rArray(i))
                        'res &= "<img src='" & rArray(i) & "' width=100%>" & vbLf
                        res &= url & vbLf
                    End If
                Else
                    Dim url As String = Replace(TextBox28.Text, "[画像URL]", rArray(i))
                    'res &= "<img src='" & rArray(i) & "' width=100%>" & vbLf
                    res &= url & vbLf
                End If
            Next
            If CheckBox14.Checked = True Then
                res &= "</p>" & vbLf & "<br>" & vbLf
            End If
            csvForm.DataGridView1.Item(sakiC, r).Value = Replace(csvForm.DataGridView1.Item(sakiC, r).Value, TextBox21.Text, res)
        Next
        MsgBox("終了しました", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
    End Sub

    '--------------------------------------------------------------------------------
    Private Sub DataGridView7_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles DataGridView7.DragDrop
        Dim fCount As Integer = 0
        Dim dataPathArray As ArrayList = New ArrayList
        For Each filename As String In e.Data.GetData(DataFormats.FileDrop)
            dataPathArray.Add(filename)
            fCount += 1
        Next

        Dim newR As Integer = 0
        If RadioButton11.Checked = True Then
            DataGridView7.Rows.Clear()
        Else
            newR = DataGridView7.RowCount
        End If
        For i As Integer = 0 To dataPathArray.Count - 1
            If i = 0 Then
                TextBox22.Text = Path.GetFileNameWithoutExtension(dataPathArray(i))
            End If
            DataGridView7.Rows.Add()
            DataGridView7.Item(0, i + newR).Value = Path.GetFileNameWithoutExtension(dataPathArray(i))
            DataGridView7.Item(2, i + newR).Value = Path.GetExtension(dataPathArray(i))
            DataGridView7.Item(3, i + newR).Value = Path.GetDirectoryName(dataPathArray(i))
        Next

        FileNameChange()
    End Sub

    Private Sub DataGridView7_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles DataGridView7.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub RadioButton9_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton9.CheckedChanged
        FileNameChange()
    End Sub

    Private Sub TextBox22_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox22.TextChanged
        FileNameChange()
    End Sub

    Private Sub TextBox25_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox25.TextChanged
        FileNameChange()
    End Sub

    Private Sub NumericUpDown4_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown4.ValueChanged
        FileNameChange()
    End Sub

    Private Sub NumericUpDown2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown2.ValueChanged
        FileNameChange()
    End Sub

    Private Sub NumericUpDown3_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown3.ValueChanged
        FileNameChange()
    End Sub

    Private Sub TextBox23_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox23.TextChanged
        FileNameChange()
    End Sub

    Private Sub TextBox24_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox24.TextChanged
        FileNameChange()
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox6.SelectedIndexChanged
        FileNameChange()
    End Sub

    Sub FileNameChange()
        If DataGridView7.RowCount > 0 Then
            For r As Integer = 0 To DataGridView7.RowCount - 1
                If RadioButton9.Checked = True Then
                    Dim num As String = NumericUpDown2.Value
                    If r < NumericUpDown4.Value - 1 Then
                        DataGridView7.Item(1, r).Value = TextBox22.Text
                    Else
                        num = r + num - (NumericUpDown4.Value - 1)
                        num = num.PadLeft(NumericUpDown3.Value, "0"c)
                        DataGridView7.Item(1, r).Value = TextBox22.Text & TextBox25.Text & num
                    End If
                ElseIf RadioButton10.Checked = True Then
                    Dim fName As String = DataGridView7.Item(0, r).Value
                    Dim nName As String = ""
                    If ComboBox6.SelectedItem = "全て置換" Then
                        fName = Replace(fName, TextBox23.Text, TextBox24.Text)
                    ElseIf ComboBox6.SelectedItem = "前方1つ" Then
                        fName = Replace(fName, TextBox23.Text, TextBox24.Text, 1, 1)
                    ElseIf ComboBox6.SelectedItem = "後方1つ" Then
                        Dim intFindPos As Integer = fName.LastIndexOf(TextBox23.Text) + 1
                        If intFindPos > 0 Then
                            nName = fName.Substring(0, intFindPos - 1)
                            fName = nName & Replace(fName, TextBox23.Text, TextBox24.Text, intFindPos, 1)
                        End If
                    End If
                    DataGridView7.Item(1, r).Value = fName
                End If

                '文字列生成
                Dim mName As String = DataGridView7.Item(1, r).Value
                Dim mNameWithExt As String = mName & DataGridView7.Item(2, r).Value
                If CheckBox17.Checked Then
                    DataGridView7.Item(4, r).Value = Replace(TextBox27.Text, TextBox30.Text, mNameWithExt)
                Else
                    DataGridView7.Item(4, r).Value = Replace(TextBox27.Text, TextBox30.Text, mName)
                End If
                If CheckBox18.Checked Then
                    DataGridView7.Item(5, r).Value = Replace(TextBox29.Text, TextBox30.Text, mNameWithExt)
                Else
                    DataGridView7.Item(5, r).Value = Replace(TextBox29.Text, TextBox30.Text, mName)
                End If
            Next
        End If
    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        TextBox23.Text = DataGridView7.Item(0, DataGridView7.SelectedCells(0).RowIndex).Value
    End Sub

    'ファイル名置換
    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        If DataGridView7.RowCount > 0 Then
            For r As Integer = 0 To DataGridView7.RowCount - 1
                Dim source As String = DataGridView7.Item(3, r).Value & "\" & DataGridView7.Item(0, r).Value & DataGridView7.Item(2, r).Value
                Dim dest As String = DataGridView7.Item(3, r).Value & "\" & DataGridView7.Item(1, r).Value & DataGridView7.Item(2, r).Value
                If File.Exists(source) Then
                    File.Move(source, dest)
                Else
                    MsgBox("元ファイルが見つかりません", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
                    Exit Sub
                End If
            Next
            MsgBox("変更しました", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
        End If
    End Sub

    'csv取り込み
    Private Sub Button28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button28.Click
        Dim ofd As New OpenFileDialog With {
            .FileName = "default.html",
            .InitialDirectory = Environment.SpecialFolder.DesktopDirectory,
            .Filter = "CSVファイル(*.csv)|*.csv|すべてのファイル(*.*)|*.*",
                     .AutoUpgradeEnabled = False,
          .FilterIndex = 2,
            .Title = "開くファイルを選択してください",
            .RestoreDirectory = True,
            .CheckFileExists = True,
            .CheckPathExists = True
        }

        'ダイアログを表示する
        If ofd.ShowDialog() = DialogResult.OK Then
            Dim fName As String = ofd.FileName
            CsvGet(fName)
        End If
    End Sub

    Private Sub CsvGet(ByVal fName As String)
        Dim itemList_web As ArrayList = CSV_READ(fName)
        Select Case True
            Case Regex.IsMatch(itemList_web(0), ".*商品分類タグ.*")
                ComboBox7.SelectedItem = "マスタ"
            Case Regex.IsMatch(itemList_web(1), ".*lucky9.*")
                ComboBox7.SelectedItem = "Lucky9"
            Case Regex.IsMatch(itemList_web(1), ".*akaneashop.*")
                ComboBox7.SelectedItem = "あかねYahoo"
            Case Regex.IsMatch(itemList_web(1), ".*fkstyle.*")
                ComboBox7.SelectedItem = "FKstyle"
            Case Regex.IsMatch(itemList_web(1), ".*patri.*")
                ComboBox7.SelectedItem = "通販の暁"
            Case Regex.IsMatch(itemList_web(1), ".*alice-zk.*")
                ComboBox7.SelectedItem = "雑貨の国のアリス"
            Case Regex.IsMatch(itemList_web(1), ".*ashop.*")
                ComboBox7.SelectedItem = "あかね楽天"
            Case Regex.IsMatch(itemList_web(0), ".*SeqExhibitId.*")
                ComboBox7.SelectedItem = "KuraNavi"
            Case Regex.IsMatch(itemList_web(0), ".*商品特定コード指定.*")
                ComboBox7.SelectedItem = "問屋よかろうもん"
            Case Regex.IsMatch(itemList_web(0), ".*Seller Code.*")
                ComboBox7.SelectedItem = "通販の雑貨倉庫"
        End Select
        Application.DoEvents()

        Dim pathEnd As String = ""
        If InStr(itemList_web(0), "原価|=|商品区分") > 0 Then
            pathEnd = Path.GetDirectoryName(Form1.appPath) & "\tenpo\" & ComboBox7.SelectedItem & "_master.csv"
        ElseIf InStr(itemList_web(0), "path|=|name") > 0 Then
            pathEnd = Path.GetDirectoryName(Form1.appPath) & "\tenpo\" & ComboBox7.SelectedItem & "_item.csv"
        ElseIf InStr(itemList_web(0), "option-name-1|=|option-value-1") > 0 Then
            pathEnd = Path.GetDirectoryName(Form1.appPath) & "\tenpo\" & ComboBox7.SelectedItem & "_select.csv"
        ElseIf InStr(itemList_web(0), "項目選択肢用コントロールカラム|=|商品管理番号（商品URL）") > 0 Then
            pathEnd = Path.GetDirectoryName(Form1.appPath) & "\tenpo\" & ComboBox7.SelectedItem & "_select.csv"
        ElseIf InStr(itemList_web(0), "コントロールカラム|=|商品管理番号（商品URL）") > 0 Then
            pathEnd = Path.GetDirectoryName(Form1.appPath) & "\tenpo\" & ComboBox7.SelectedItem & "_item.csv"
        ElseIf InStr(itemList_web(0), "SeqExhibitId|=|Code") > 0 Then
            pathEnd = Path.GetDirectoryName(Form1.appPath) & "\tenpo\" & ComboBox7.SelectedItem & "_item.csv"
        ElseIf InStr(itemList_web(0), "Code|=|SkuOptionTypeCol") > 0 Then
            pathEnd = Path.GetDirectoryName(Form1.appPath) & "\tenpo\" & ComboBox7.SelectedItem & "_select.csv"
        ElseIf InStr(itemList_web(0), "商品特定コード指定|=|更新時間フラグ") > 0 Then
            pathEnd = Path.GetDirectoryName(Form1.appPath) & "\tenpo\" & ComboBox7.SelectedItem & "_item.csv"
        ElseIf InStr(itemList_web(0), "商品特定コード指定|=|オプション特定コード指定") > 0 Then
            pathEnd = Path.GetDirectoryName(Form1.appPath) & "\tenpo\" & ComboBox7.SelectedItem & "_select.csv"
        ElseIf InStr(itemList_web(0), "Item Code|=|Seller Code") > 0 Then
            pathEnd = Path.GetDirectoryName(Form1.appPath) & "\tenpo\" & ComboBox7.SelectedItem & "_item.csv"
        End If

        Dim DR As DialogResult = MsgBox(pathEnd & "でよろしいですか？", MsgBoxStyle.OkCancel Or MsgBoxStyle.SystemModal)
        If DR = Windows.Forms.DialogResult.Cancel Then
            Exit Sub
        Else
            File.Copy(fName, pathEnd, True)
            MsgBox("データを登録しました", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
        End If

        FileGetLastTime()
    End Sub


    Private Sub Button29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button29.Click
        Dim FileArray As String() = New String() {"FKstyle", "Lucky9", "あかねYahoo", "通販の暁", "雑貨の国のアリス", "あかね楽天", "KuraNavi", "問屋よかろうもん", "通販の雑貨倉庫"}
        Dim folderPath As String = Path.GetDirectoryName(Form1.appPath) & "\tenpo\"

        For r As Integer = 0 To DataGridView8.RowCount - 1
            For c As Integer = 1 To DataGridView8.ColumnCount - 1
                DataGridView8.Item(c, r).Value = ""
            Next
        Next
        Application.DoEvents()

        ToolStripProgressBar2.Value = 0
        ToolStripProgressBar2.Maximum = FileArray.Length

        For r As Integer = 0 To FileArray.Length - 1
            If File.Exists(folderPath & FileArray(r) & "_item.csv") = True Then
                Dim itemList_web As ArrayList = CSV_READ(folderPath & FileArray(r) & "_item.csv")

                'ヘッダーテキストをコレクションに取り込む
                Dim headerArray As String() = Split(itemList_web(0), "|=|")
                Dim DGVheaderCheck As ArrayList = New ArrayList
                For c As Integer = 0 To headerArray.Length - 1
                    DGVheaderCheck.Add(headerArray(c))
                Next c

                Dim codeName As String = ""
                Dim priceName As String = ""
                Dim saleName As String = ""
                Dim haisouName As String = ""
                Dim KsouryoName As String = ""
                Dim soukoName As String = ""
                Select Case FileArray(r)
                    Case "FKstyle", "Lucky9", "あかねYahoo"
                        codeName = "code"
                        priceName = "price"
                        saleName = "sale-price"
                        haisouName = "ship-weight"
                        KsouryoName = ""
                        soukoName = "display"
                    Case "通販の暁", "雑貨の国のアリス", "あかね楽天"
                        codeName = "商品管理番号（商品URL）"
                        priceName = "販売価格"
                        saleName = "表示価格"
                        haisouName = "配送方法セット管理番号"
                        KsouryoName = "個別送料"
                        soukoName = "倉庫指定"
                    Case "KuraNavi"
                        'codeName = "Code"
                        codeName = "itemCode"
                        'priceName = "Price"
                        priceName = "itemPrice"
                        saleName = ""
                        'haisouName = "DeliveryType"
                        haisouName = "deliveryMethodName1"
                        'KsouryoName = "ShipFee"
                        KsouryoName = ""
                        soukoName = ""
                    Case "問屋よかろうもん"
                        codeName = "独自商品コード"
                        priceName = "販売価格"
                        saleName = ""
                        haisouName = "送料個別設定"
                        KsouryoName = ""
                        soukoName = "商品表示可否"
                    Case "通販の雑貨倉庫"
                        codeName = "Seller Code"
                        priceName = "Sell Price"
                        saleName = ""
                        haisouName = "Shipping Group No"
                        KsouryoName = ""
                        soukoName = ""
                    Case Else
                        MsgBox("チェック失敗")
                        Exit Sub
                End Select

                For i As Integer = 0 To itemList_web.Count - 1
                    Dim res As String() = Split(itemList_web(i), "|=|")
                    If res(DGVheaderCheck.IndexOf(codeName)) = TextBox26.Text Then
                        ToolStripStatusLabel8.Text = i
                        DataGridView8.Item(1, r + 1).Value = res(DGVheaderCheck.IndexOf(priceName))

                        '送料
                        Dim souryou As Integer = 0
                        Select Case FileArray(r)
                            Case "FKstyle"
                                Dim s As Integer = CInt(res(DGVheaderCheck.IndexOf(haisouName)))
                                If 2 <= s And s <= 50 Then
                                    souryou = 189
                                ElseIf 51 <= s And s <= 100 Then
                                    souryou = 378
                                ElseIf 101 <= s And s <= 150 Then
                                    souryou = 567
                                ElseIf s <= 1 Or (150 <= s And s <= 9999) Then
                                    souryou = 750
                                Else
                                    souryou = 972
                                End If
                            Case "通販の暁"
                                If res(DGVheaderCheck.IndexOf(haisouName)) <> "" Then
                                    Dim s As Integer = CInt(res(DGVheaderCheck.IndexOf(haisouName)))
                                    If s = 1 Or s = 2 Then
                                        souryou = 760
                                    ElseIf s = 3 Then
                                        souryou = 1260
                                    Else
                                        souryou = 240
                                    End If
                                Else    '空欄の場合（宅配便）
                                    souryou = 760
                                End If
                            Case "KuraNavi"
                                Dim s As String = res(DGVheaderCheck.IndexOf(haisouName))
                                If s = "M" Then
                                    souryou = 240
                                Else
                                    'If res(DGVheaderCheck.IndexOf("ShipFeeType")) = "E" Then
                                    If res(DGVheaderCheck.IndexOf("deliveryMethodId1")) = "E" Then
                                        souryou = 0
                                    Else
                                        souryou = 756
                                    End If
                                End If
                            Case "問屋よかろうもん"
                                Dim s As String = res(DGVheaderCheck.IndexOf(haisouName))
                                If s = "20170407121422" Then
                                    souryou = 240
                                ElseIf s = "20170407121342" Then
                                    souryou = 750
                                Else
                                    souryou = 1250
                                End If
                            Case "通販の雑貨倉庫"
                                Dim s As String = res(DGVheaderCheck.IndexOf(haisouName))
                                Select Case s
                                    Case "395750", "380887", "379225", "379224", "359409"
                                        souryou = 756
                                    Case "395082", "394665", "394313", "394255", "373781"
                                        souryou = 240
                                    Case "393930"
                                        souryou = 340
                                    Case "359413"
                                        souryou = 972
                                    Case Else
                                        souryou = 0
                                End Select
                            Case Else
                                souryou = 0
                        End Select
                        If KsouryoName <> "" Then
                            If res(DGVheaderCheck.IndexOf(KsouryoName)) <> "" Then
                                souryou = res(DGVheaderCheck.IndexOf(KsouryoName))
                            End If
                        End If
                        DataGridView8.Item(3, r + 1).Value = souryou

                        If saleName <> "" Then
                            If res(DGVheaderCheck.IndexOf(saleName)) <> "" Then
                                Select Case FileArray(r)
                                    Case "通販の暁", "雑貨の国のアリス", "あかね楽天"
                                        DataGridView8.Item(2, r + 1).Value = DataGridView8.Item(1, r + 1).Value
                                        DataGridView8.Item(1, r + 1).Value = res(DGVheaderCheck.IndexOf(saleName))
                                    Case Else
                                        DataGridView8.Item(2, r + 1).Value = res(DGVheaderCheck.IndexOf(saleName))
                                End Select
                                DataGridView8.Item(4, r + 1).Value = CInt(DataGridView8.Item(2, r + 1).Value) + CInt(DataGridView8.Item(3, r + 1).Value)
                            Else
                                DataGridView8.Item(2, r + 1).Value = 0
                                DataGridView8.Item(4, r + 1).Value = CInt(DataGridView8.Item(1, r + 1).Value) + CInt(DataGridView8.Item(3, r + 1).Value)
                            End If
                        Else
                            DataGridView8.Item(2, r + 1).Value = 0
                            DataGridView8.Item(4, r + 1).Value = CInt(DataGridView8.Item(1, r + 1).Value) + CInt(DataGridView8.Item(3, r + 1).Value)
                        End If

                        '倉庫
                        Dim souko As String = ""
                        If soukoName = "" Then
                            souko = "-"
                        Else
                            souko = res(DGVheaderCheck.IndexOf(soukoName))
                        End If
                        Select Case FileArray(r)
                            Case "FKstyle", "Lucky9", "あかねYahoo"
                                If souko = "1" Then
                                    souko = "表示"
                                Else
                                    souko = "✕"
                                End If
                            Case "通販の暁", "雑貨の国のアリス", "あかね楽天"
                                If souko = "0" Then
                                    souko = "表示"
                                Else
                                    souko = "✕"
                                End If
                            Case "問屋よかろうもん"
                                If souko = "Y" Then
                                    souko = "表示"
                                Else
                                    souko = "✕"
                                End If
                            Case Else
                                souko = "-"
                        End Select
                        DataGridView8.Item(5, r + 1).Value = souko

                        Exit For
                    End If
                Next

                ToolStripProgressBar2.Value += 1
            End If
        Next

        ToolStripProgressBar2.Value = 0
    End Sub

    Dim FileArray As String() = New String() {"FKstyle", "Lucky9", "あかねYahoo", "通販の暁", "雑貨の国のアリス", "あかね楽天", "KuraNavi", "問屋よかろうもん", "通販の雑貨倉庫"}
    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        If ListBox1.SelectedItems.Count = 0 Then
            Exit Sub
        End If

        DataGridView10.Rows.Clear()
        If DataGridView10.ColumnCount > 3 Then
            For c As Integer = DataGridView10.ColumnCount - 1 To 3
                DataGridView10.Columns.RemoveAt(DataGridView10.ColumnCount - 1)
            Next
        End If

        ToolStripStatusLabel8.Text = "検索開始"

        Dim folderPath As String = Path.GetDirectoryName(Form1.appPath) & "\tenpo\"

        If File.Exists(folderPath & "マスタ_master.csv") = True Then
            Dim itemList_web As ArrayList = CSV_READ(folderPath & "マスタ_master.csv")

            'ヘッダーテキストをコレクションに取り込む
            Dim headerArray As String() = Split(itemList_web(0), "|=|")
            Dim DGVheaderCheck As ArrayList = New ArrayList
            For c As Integer = 0 To headerArray.Length - 1
                DGVheaderCheck.Add(headerArray(c))
            Next c

            DataGridView10.Rows.Add()
            Dim normalArray As ArrayList = New ArrayList     '通常の商品検索のため、通常リスト作成
            Dim yoyakuArrayD As ArrayList = New ArrayList      '予約リスト（代表コード）
            For i As Integer = 0 To itemList_web.Count - 1
                Dim res As String() = Split(itemList_web(i), "|=|")
                Dim Scode As String = res(DGVheaderCheck.IndexOf("商品コード"))
                Dim Dcode As String = res(DGVheaderCheck.IndexOf("代表商品コード"))
                If res(DGVheaderCheck.IndexOf("商品区分")) = "予約" Then
                    Dim sArray As String() = New String() {Scode, Dcode, res(DGVheaderCheck.IndexOf("商品区分"))}
                    DataGridView10.Rows.Add(sArray)

                    '--------------------------
                    If Dcode = "" Then
                        yoyakuArrayD.Add(Scode)
                    Else
                        yoyakuArrayD.Add(Dcode)
                    End If
                    '--------------------------
                Else
                    '通常商品はサブコードで他に予約になっていればリストに入れない
                    Dim Xcode As String = Dcode
                    If Dcode = "" Then
                        Xcode = Scode
                    End If
                    If Not yoyakuArrayD.Contains(Xcode) Then
                        normalArray.Add(Scode & "," & Dcode & "," & res(DGVheaderCheck.IndexOf("商品区分")))
                    End If
                End If
            Next

            Application.DoEvents()

            '各店データを調べる
            Dim colNum As Integer = 0
            Dim fileName As String = ListBox1.SelectedItem
            If File.Exists(folderPath & fileName & "_item.csv") Then
                Dim itemList1 As ArrayList = CSV_READ(folderPath & fileName & "_item.csv")
                Dim itemList2 As ArrayList = New ArrayList
                If fileName <> "通販の雑貨倉庫" Then
                    itemList2 = CSV_READ(folderPath & fileName & "_select.csv")
                End If

                'ヘッダーテキストをコレクションに取り込む
                headerArray = Split(itemList1(0), "|=|")
                Dim DGVheaderCheck1 As ArrayList = New ArrayList
                For c As Integer = 0 To headerArray.Length - 1
                    DGVheaderCheck1.Add(headerArray(c))
                Next c
                Dim DGVheaderCheck2 As ArrayList = New ArrayList
                If fileName <> "通販の雑貨倉庫" Then
                    headerArray = Split(itemList2(0), "|=|")
                    For c As Integer = 0 To headerArray.Length - 1
                        DGVheaderCheck2.Add(headerArray(c))
                    Next c
                End If

                DataGridView10.Columns.Add("Column" & colNum + 3, fileName)
                DataGridView10.Columns.Add("Column" & colNum + 4, "option")
                DataGridView10.Item(colNum + 3, 0).Value = fileName

                '*************************************************************************************
                ' 予約状態調査
                '*************************************************************************************
                Select Case fileName
                    Case "FKstyle", "Lucky9", "あかねYahoo"

                        For r As Integer = 1 To DataGridView10.RowCount - 1
                            Dim Scode As String = DataGridView10.Item(0, r).Value
                            Dim Dcode As String = DataGridView10.Item(1, r).Value   '代表コード
                            Scode = Scode.ToLower
                            Dcode = Dcode.ToLower
                            'Yahooは大文字小文字を区別しないので.ToLowerで処理している

                            '倉庫に入っているか調べる
                            Dim souko As Boolean = False
                            If Dcode <> "" Then
                                For k As Integer = 0 To itemList1.Count - 1
                                    Dim res As String() = Split(itemList1(k), "|=|")
                                    If Dcode = res(DGVheaderCheck1.IndexOf("code")).ToLower Then
                                        If res(DGVheaderCheck1.IndexOf("display")) = "0" Then
                                            DataGridView10.Item(colNum + 3, r).Value = "倉庫"
                                            souko = True
                                            Application.DoEvents()
                                            Exit For
                                        End If
                                    End If
                                Next
                            Else
                                For k As Integer = 0 To itemList1.Count - 1
                                    Dim res As String() = Split(itemList1(k), "|=|")
                                    If Scode = res(DGVheaderCheck1.IndexOf("code")).ToLower Then
                                        If res(DGVheaderCheck1.IndexOf("display")) = "0" Then
                                            DataGridView10.Item(colNum + 3, r).Value = "倉庫"
                                            souko = True
                                            Application.DoEvents()
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If

                            If souko = False Then
                                If Dcode <> "" Then
                                    For c As Integer = 0 To itemList2.Count - 1
                                        Dim res As String() = Split(itemList2(c), "|=|")
                                        If Scode = res(DGVheaderCheck2.IndexOf("sub-code")).ToLower Then
                                            DataGridView10.Item(colNum + 3, r).Value = res(DGVheaderCheck2.IndexOf("lead-time-instock"))
                                            ToolStripStatusLabel8.Text = c
                                            Application.DoEvents()
                                            Exit For
                                        End If
                                    Next
                                    If DataGridView10.Item(colNum + 3, r).Value = "" Then
                                        For c As Integer = 0 To itemList1.Count - 1
                                            Dim res As String() = Split(itemList1(c), "|=|")
                                            If Dcode = res(DGVheaderCheck1.IndexOf("code")).ToLower Then
                                                DataGridView10.Item(colNum + 3, r).Value = res(DGVheaderCheck1.IndexOf("lead-time-instock"))
                                                ToolStripStatusLabel8.Text = c
                                                Application.DoEvents()
                                                Exit For
                                            End If
                                        Next
                                    End If
                                Else
                                    For c As Integer = 0 To itemList1.Count - 1
                                        Dim res As String() = Split(itemList1(c), "|=|")
                                        If Scode = res(DGVheaderCheck1.IndexOf("code")).ToLower Then
                                            DataGridView10.Item(colNum + 3, r).Value = res(DGVheaderCheck1.IndexOf("lead-time-instock"))
                                            ToolStripStatusLabel8.Text = c
                                            Application.DoEvents()
                                            Exit For
                                        End If
                                    Next
                                End If

                                Select Case DataGridView10.Item(colNum + 3, r).Value
                                    Case "1000", "2"
                                        DataGridView10.Item(colNum + 3, r).Style.BackColor = Color.Yellow
                                End Select
                            End If
                        Next

                        If CheckBox16.Checked Then
                            For r As Integer = 0 To normalArray.Count - 1
                                Dim codeArray As String() = Split(normalArray(r), ",")
                                Dim Scode As String = codeArray(0)
                                Dim Dcode As String = codeArray(1)
                                Scode = Scode.ToLower
                                Dcode = Dcode.ToLower

                                Dim yFlag As Boolean = False
                                Dim yStatus As String = ""
                                If Dcode <> "" Then
                                    For c As Integer = 0 To itemList2.Count - 1
                                        Dim res As String() = Split(itemList2(c), "|=|")
                                        If Scode = res(DGVheaderCheck2.IndexOf("sub-code")).ToLower Then
                                            If Regex.IsMatch(res(DGVheaderCheck2.IndexOf("lead-time-instock")), "3000|4000|5000") Then
                                                yFlag = True
                                                yStatus = res(DGVheaderCheck2.IndexOf("lead-time-instock"))
                                                ToolStripStatusLabel8.Text = c
                                                Exit For
                                            End If
                                        End If
                                    Next
                                    If yFlag = False Then
                                        For c As Integer = 0 To itemList1.Count - 1
                                            Dim res As String() = Split(itemList1(c), "|=|")
                                            If Dcode = res(DGVheaderCheck1.IndexOf("code")).ToLower Then
                                                If Regex.IsMatch(res(DGVheaderCheck1.IndexOf("lead-time-instock")), "3000|4000|5000") Then
                                                    yFlag = True
                                                    yStatus = res(DGVheaderCheck1.IndexOf("lead-time-instock"))
                                                    ToolStripStatusLabel8.Text = c
                                                    Exit For
                                                End If
                                            End If
                                        Next
                                    End If
                                Else
                                    For c As Integer = 0 To itemList1.Count - 1
                                        Dim res As String() = Split(itemList1(c), "|=|")
                                        If Scode = res(DGVheaderCheck1.IndexOf("code")).ToLower Then
                                            If Regex.IsMatch(res(DGVheaderCheck1.IndexOf("lead-time-instock")), "3000|4000|5000") Then
                                                yFlag = True
                                                yStatus = res(DGVheaderCheck1.IndexOf("lead-time-instock"))
                                                ToolStripStatusLabel8.Text = c
                                                Exit For
                                            End If
                                        End If
                                    Next
                                End If

                                If yFlag = True Then
                                    Dim sArray As String() = New String() {Scode, Dcode, codeArray(2), yStatus}
                                    DataGridView10.Rows.Add(sArray)
                                    DataGridView10.Rows(DataGridView10.Rows.Count - 1).DefaultCellStyle.BackColor = Color.Yellow
                                    Application.DoEvents()
                                End If

                                If r Mod 100 = 0 Then
                                    ToolStripStatusLabel8.Text = "検索中：" & normalArray.Count - r
                                    Application.DoEvents()
                                End If
                            Next

                            colNum += 1
                        End If
                    Case "通販の暁", "雑貨の国のアリス", "あかね楽天"
                        For r As Integer = 1 To DataGridView10.RowCount - 1
                            Dim Scode As String = DataGridView10.Item(0, r).Value
                            Dim Dcode As String = DataGridView10.Item(1, r).Value
                            '--------------------------
                            Dim Xcode As String = Dcode     'サブコード無しのコード計算
                            If Dcode = "" Then
                                Xcode = Scode
                            End If
                            '--------------------------

                            '倉庫に入っているか調べる
                            Dim souko As Boolean = False
                            If Dcode <> "" Then
                                For k As Integer = 0 To itemList1.Count - 1
                                    Dim res As String() = Split(itemList1(k), "|=|")
                                    If Dcode = res(DGVheaderCheck1.IndexOf("商品管理番号（商品URL）")) Then
                                        If res(DGVheaderCheck1.IndexOf("倉庫指定")) = "1" Then
                                            DataGridView10.Item(colNum + 3, r).Value = "倉庫"
                                            souko = True
                                            Exit For
                                        End If
                                    End If
                                Next
                            Else
                                For k As Integer = 0 To itemList1.Count - 1
                                    Dim res As String() = Split(itemList1(k), "|=|")
                                    If Scode = res(DGVheaderCheck1.IndexOf("商品管理番号（商品URL）")) Then
                                        If res(DGVheaderCheck1.IndexOf("倉庫指定")) = "1" Then
                                            DataGridView10.Item(colNum + 3, r).Value = "倉庫"
                                            souko = True
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If

                            If souko = False Then
                                If Dcode <> "" Then
                                    For c As Integer = 0 To itemList2.Count - 1
                                        Dim res As String() = Split(itemList2(c), "|=|")
                                        If Dcode = res(DGVheaderCheck2.IndexOf("商品管理番号（商品URL）")) Then
                                            Dim CoBango As String = Replace(Scode, Dcode, "")
                                            Dim Ccode As String = res(DGVheaderCheck2.IndexOf("項目選択肢別在庫用横軸選択肢子番号")) & res(DGVheaderCheck2.IndexOf("項目選択肢別在庫用縦軸選択肢子番号"))
                                            If CoBango = Ccode Then
                                                DataGridView10.Item(colNum + 3, r).Value = res(DGVheaderCheck2.IndexOf("在庫あり時納期管理番号"))
                                                ToolStripStatusLabel8.Text = c
                                                Exit For
                                            End If
                                        End If
                                    Next
                                    If DataGridView10.Item(colNum + 3, r).Value = "" Then
                                        For c As Integer = 0 To itemList1.Count - 1
                                            Dim res As String() = Split(itemList1(c), "|=|")
                                            If Dcode = res(DGVheaderCheck1.IndexOf("商品管理番号（商品URL）")) Then
                                                DataGridView10.Item(colNum + 3, r).Value = res(DGVheaderCheck1.IndexOf("在庫あり時納期管理番号"))
                                                ToolStripStatusLabel8.Text = c
                                                Exit For
                                            End If
                                        Next
                                    End If
                                Else
                                    For c As Integer = 0 To itemList1.Count - 1
                                        Dim res As String() = Split(itemList1(c), "|=|")
                                        If Scode = res(DGVheaderCheck1.IndexOf("商品管理番号（商品URL）")) Then
                                            DataGridView10.Item(colNum + 3, r).Value = res(DGVheaderCheck1.IndexOf("在庫あり時納期管理番号"))
                                            ToolStripStatusLabel8.Text = c
                                            Exit For
                                        End If
                                    Next
                                End If
                                Application.DoEvents()

                                'オプション検索
                                For c As Integer = 0 To itemList2.Count - 1
                                    Dim res As String() = Split(itemList2(c), "|=|")
                                    If Xcode = res(DGVheaderCheck2.IndexOf("商品管理番号（商品URL）")) Then
                                        If InStr(res(DGVheaderCheck2.IndexOf("Select/Checkbox用項目名")), "予約") > 0 Then
                                            DataGridView10.Item(colNum + 4, r).Value &= res(DGVheaderCheck2.IndexOf("Select/Checkbox用項目名"))
                                        End If
                                    End If
                                Next

                                Select Case DataGridView10.Item(colNum + 3, r).Value
                                    Case "1"
                                        DataGridView10.Item(colNum + 3, r).Style.BackColor = Color.Yellow
                                        Application.DoEvents()
                                End Select
                            End If
                        Next

                        If CheckBox16.Checked Then
                            For r As Integer = 0 To normalArray.Count - 1
                                Dim codeArray As String() = Split(normalArray(r), ",")
                                Dim Scode As String = codeArray(0)
                                Dim Dcode As String = codeArray(1)
                                '--------------------------
                                Dim Xcode As String = Dcode     'サブコード無しのコード計算
                                If Dcode = "" Then
                                    Xcode = Scode
                                End If
                                '--------------------------

                                Dim yFlag As Boolean = False
                                Dim yStatus As String = ""
                                If Dcode <> "" Then
                                    For c As Integer = 0 To itemList2.Count - 1
                                        Dim res As String() = Split(itemList2(c), "|=|")
                                        If Dcode = res(DGVheaderCheck2.IndexOf("商品管理番号（商品URL）")) Then
                                            Dim CoBango As String = Replace(Scode, Dcode, "")
                                            Dim Ccode As String = res(DGVheaderCheck2.IndexOf("項目選択肢別在庫用横軸選択肢子番号")) & res(DGVheaderCheck2.IndexOf("項目選択肢別在庫用縦軸選択肢子番号"))
                                            If CoBango = Ccode Then
                                                If Regex.IsMatch(res(DGVheaderCheck2.IndexOf("在庫あり時納期管理番号")), "[2-4]") Then
                                                    yFlag = True
                                                    yStatus = res(DGVheaderCheck2.IndexOf("在庫あり時納期管理番号"))
                                                    ToolStripStatusLabel8.Text = c
                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    Next
                                    If yFlag = False Then
                                        For c As Integer = 0 To itemList1.Count - 1
                                            Dim res As String() = Split(itemList1(c), "|=|")
                                            If Dcode = res(DGVheaderCheck1.IndexOf("商品管理番号（商品URL）")) Then
                                                If Regex.IsMatch(res(DGVheaderCheck1.IndexOf("在庫あり時納期管理番号")), "[2-4]") Then
                                                    yFlag = True
                                                    yStatus = res(DGVheaderCheck1.IndexOf("在庫あり時納期管理番号"))
                                                    ToolStripStatusLabel8.Text = c
                                                    Exit For
                                                End If
                                            End If
                                        Next
                                    End If
                                Else
                                    For c As Integer = 0 To itemList1.Count - 1
                                        Dim res As String() = Split(itemList1(c), "|=|")
                                        If Scode = res(DGVheaderCheck1.IndexOf("商品管理番号（商品URL）")) Then
                                            If Regex.IsMatch(res(DGVheaderCheck1.IndexOf("在庫あり時納期管理番号")), "[2-4]") Then
                                                yFlag = True
                                                yStatus = res(DGVheaderCheck1.IndexOf("在庫あり時納期管理番号"))
                                                ToolStripStatusLabel8.Text = c
                                                Exit For
                                            End If
                                        End If
                                    Next
                                End If

                                If yFlag = True Then
                                    Dim sArray As String() = New String() {Scode, Dcode, codeArray(2), yStatus}
                                    DataGridView10.Rows.Add(sArray)
                                    DataGridView10.Rows(DataGridView10.Rows.Count - 1).DefaultCellStyle.BackColor = Color.Yellow
                                    Application.DoEvents()
                                End If

                                'オプション検索
                                For c As Integer = 0 To itemList2.Count - 1
                                    Dim res As String() = Split(itemList2(c), "|=|")
                                    If Xcode = res(DGVheaderCheck2.IndexOf("商品管理番号（商品URL）")) Then
                                        If InStr(res(DGVheaderCheck2.IndexOf("Select/Checkbox用項目名")), "予約") > 0 Then
                                            Dim sArray As String() = New String() {Scode, Dcode, codeArray(2), yStatus}
                                            DataGridView10.Rows.Add(sArray)
                                            DataGridView10.Item(colNum + 4, DataGridView10.Rows.Count - 1).Value &= res(DGVheaderCheck2.IndexOf("Select/Checkbox用項目名"))
                                            DataGridView10.Item(colNum + 4, DataGridView10.Rows.Count - 1).Style.BackColor = Color.Yellow
                                        End If
                                    End If
                                Next

                                If r Mod 100 = 0 Then
                                    ToolStripStatusLabel8.Text = "検索中：" & normalArray.Count - r
                                    Application.DoEvents()
                                End If
                            Next

                            colNum += 1
                        End If
                    Case "KuraNavi"
                        For r As Integer = 1 To DataGridView10.RowCount - 1
                            Dim Scode As String = DataGridView10.Item(0, r).Value
                            Dim Dcode As String = DataGridView10.Item(1, r).Value
                            If Dcode <> "" Then
                                For c As Integer = 0 To itemList1.Count - 1
                                    Dim res As String() = Split(itemList1(c), "|=|")
                                    If Dcode = res(DGVheaderCheck1.IndexOf("Code")) Then
                                        DataGridView10.Item(colNum + 3, r).Value = res(DGVheaderCheck1.IndexOf("Caption"))
                                        ToolStripStatusLabel8.Text = c
                                        Application.DoEvents()
                                        Exit For
                                    End If
                                Next
                            Else
                                For c As Integer = 0 To itemList1.Count - 1
                                    Dim res As String() = Split(itemList1(c), "|=|")
                                    If Scode = res(DGVheaderCheck1.IndexOf("Code")) Then
                                        DataGridView10.Item(colNum + 3, r).Value = res(DGVheaderCheck1.IndexOf("Caption"))
                                        ToolStripStatusLabel8.Text = c
                                        Application.DoEvents()
                                        Exit For
                                    End If
                                Next
                            End If

                            If DataGridView10.Item(colNum + 3, r).Value <> "" Then
                                If InStr(DataGridView10.Item(colNum + 3, r).Value, "予約") = 0 Then
                                    DataGridView10.Item(colNum + 3, r).Style.BackColor = Color.Yellow
                                End If
                            End If
                        Next

                        colNum += 1
                    Case "問屋よかろうもん"
                        For r As Integer = 1 To DataGridView10.RowCount - 1
                            Dim Scode As String = DataGridView10.Item(0, r).Value
                            Dim Dcode As String = DataGridView10.Item(1, r).Value

                            '倉庫に入っているか調べる
                            Dim souko As Boolean = False
                            If Dcode <> "" Then
                                For k As Integer = 0 To itemList1.Count - 1
                                    Dim res As String() = Split(itemList1(k), "|=|")
                                    If Dcode = res(DGVheaderCheck1.IndexOf("独自商品コード")) Then
                                        If res(DGVheaderCheck1.IndexOf("商品表示可否")) = "N" Then
                                            DataGridView10.Item(colNum + 3, r).Value = "倉庫"
                                            souko = True
                                            Application.DoEvents()
                                            Exit For
                                        End If
                                    End If
                                Next
                            Else
                                For k As Integer = 0 To itemList1.Count - 1
                                    Dim res As String() = Split(itemList1(k), "|=|")
                                    If Scode = res(DGVheaderCheck1.IndexOf("独自商品コード")) Then
                                        If res(DGVheaderCheck1.IndexOf("商品表示可否")) = "N" Then
                                            DataGridView10.Item(colNum + 3, r).Value = "倉庫"
                                            souko = True
                                            Application.DoEvents()
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If

                            If souko = False Then
                                If Dcode <> "" Then
                                    For c As Integer = 0 To itemList1.Count - 1
                                        Dim res As String() = Split(itemList1(c), "|=|")
                                        If Dcode = res(DGVheaderCheck2.IndexOf("独自商品コード")) Then
                                            DataGridView10.Item(colNum + 3, r).Value = res(DGVheaderCheck1.IndexOf("製造元")) & "/" & res(DGVheaderCheck1.IndexOf("商品別特殊表示"))
                                            ToolStripStatusLabel8.Text = c
                                            Application.DoEvents()
                                            Exit For
                                        End If
                                    Next
                                Else
                                    For c As Integer = 0 To itemList1.Count - 1
                                        Dim res As String() = Split(itemList1(c), "|=|")
                                        If Scode = res(DGVheaderCheck1.IndexOf("独自商品コード")) Then
                                            DataGridView10.Item(colNum + 3, r).Value = res(DGVheaderCheck1.IndexOf("製造元")) & "/" & res(DGVheaderCheck1.IndexOf("商品別特殊表示"))
                                            ToolStripStatusLabel8.Text = c
                                            Application.DoEvents()
                                            Exit For
                                        End If
                                    Next
                                End If
                            End If

                            If DataGridView10.Item(colNum + 3, r).Value <> "" Then
                                If InStr(DataGridView10.Item(colNum + 3, r).Value, "予約/") = 0 Then
                                    DataGridView10.Item(colNum + 3, r).Style.BackColor = Color.Orange
                                Else
                                    If InStr(DataGridView10.Item(colNum + 3, r).Value, "予約") = 0 Then
                                        DataGridView10.Item(colNum + 3, r).Style.BackColor = Color.Yellow
                                    End If
                                End If
                            End If
                        Next

                        colNum += 1
                    Case "通販の雑貨倉庫"
                        For r As Integer = 1 To DataGridView10.RowCount - 1
                            Dim Scode As String = DataGridView10.Item(0, r).Value
                            Dim Dcode As String = DataGridView10.Item(1, r).Value
                            Scode = Scode.ToLower
                            Dcode = Dcode.ToLower

                            If Dcode <> "" Then
                                For c As Integer = 0 To itemList1.Count - 1
                                    Dim res As String() = Split(itemList1(c), "|=|")
                                    If Dcode = res(DGVheaderCheck1.IndexOf("Seller Code")).ToLower Then
                                        DataGridView10.Item(colNum + 3, r).Value = res(DGVheaderCheck1.IndexOf("Available Date")) & "/" & res(DGVheaderCheck1.IndexOf("Option Info"))
                                        ToolStripStatusLabel8.Text = c
                                        Application.DoEvents()
                                        Exit For
                                    End If
                                Next
                            Else
                                For c As Integer = 0 To itemList1.Count - 1
                                    Dim res As String() = Split(itemList1(c), "|=|")
                                    If Scode = res(DGVheaderCheck1.IndexOf("Seller Code")).ToLower Then
                                        DataGridView10.Item(colNum + 3, r).Value = res(DGVheaderCheck1.IndexOf("Available Date")) & "/" & res(DGVheaderCheck1.IndexOf("Option Info"))
                                        ToolStripStatusLabel8.Text = c
                                        Application.DoEvents()
                                        Exit For
                                    End If
                                Next
                            End If

                            If DataGridView10.Item(colNum + 3, r).Value <> "" Then
                                If InStr(DataGridView10.Item(colNum + 3, r).Value, "予約") = 0 Then
                                    DataGridView10.Item(colNum + 3, r).Style.BackColor = Color.Yellow
                                End If
                            End If
                        Next

                        colNum += 1
                End Select

            End If

            ToolStripStatusLabel8.Text = ""
            MsgBox("チェック終了しました")
        End If
    End Sub

    Private Sub DataGridView10_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView10.SelectionChanged
        Try
            If DataGridView10.SelectedCells(0).Value <> Nothing Or DataGridView10.SelectedCells(0).Value <> "" Then
                AzukiControl1.Document.Text = DataGridView10.SelectedCells(0).Value
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub DataGridView9_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles DataGridView9.DragDrop
        For Each filename As String In e.Data.GetData(DataFormats.FileDrop)
            CsvGet(filename)
        Next

    End Sub

    Private Sub DataGridView9_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles DataGridView9.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        If DataGridView10.RowCount = 0 Then
            Exit Sub
        End If

        Dim sfd As New SaveFileDialog With {
            .InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
            .Filter = "CSVファイル(*.csv)|*.csv|すべてのファイル(*.*)|*.*",
            .AutoUpgradeEnabled = False,
            .FilterIndex = 0,
            .Title = "保存先のファイルを選択してください",
            .RestoreDirectory = True,
            .OverwritePrompt = True,
            .CheckPathExists = True
        }
        If InStr(Me.Text, "\") > 0 Then
            Dim sPath As String = IO.Path.GetFileNameWithoutExtension(Me.Text)
            sfd.FileName = sPath & ".csv"
        Else
            sfd.FileName = "新しいファイル.csv"
        End If

        'ダイアログを表示する
        If sfd.ShowDialog(Me) = DialogResult.OK Then
            SaveCsv(sfd.FileName)
            MsgBox(sfd.FileName & vbCrLf & "保存しました", MsgBoxStyle.SystemModal)
        End If
    End Sub


    Public Sub SaveCsv(ByVal fp As String)
        ' DataGridViewをCSV出力する
        DataGridView10.EndEdit()

        ' CSVファイルオープン
        Dim str As String = ""
        Dim sw As IO.StreamWriter = New IO.StreamWriter(fp, False, System.Text.Encoding.GetEncoding("SHIFT-JIS"))
        For r As Integer = 0 To DataGridView10.Rows.Count - 2
            If DataGridView10.Rows(r).Visible = True Then
                For c As Integer = 0 To DataGridView10.Columns.Count - 1
                    ' DataGridViewのセルのデータ取得
                    Dim dt As String = ""
                    If DataGridView10.Rows(r).Cells(c).Value Is Nothing = False Then
                        dt = DataGridView10.Rows(r).Cells(c).Value.ToString()
                        dt = Replace(dt, vbCrLf, vbLf)
                        'dt = Replace(dt, vbLf, "")
                        If Not dt Is Nothing Then
                            dt = Replace(dt, """", """""")
                            Select Case True
                                Case InStr(dt, ","), InStr(dt, vbLf), InStr(dt, vbCr), InStr(dt, """") 'Regex.IsMatch(dt, ".*,.*|.*" & vbLf & ".*|.*" & vbCr & ".*|.*"".*")
                                    dt = """" & dt & """"
                            End Select
                        End If
                    End If
                    If c < DataGridView10.Columns.Count - 1 Then
                        dt = dt & ","
                    End If
                    ' CSVファイル書込
                    sw.Write(dt)
                Next
                sw.Write(vbLf)
            End If
        Next

        ' CSVファイルクローズ
        sw.Close()
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        Dim str As String = ""
        For r As Integer = 0 To DataGridView7.Rows.Count - 1
            str &= DataGridView7.Item(1, r).Value & DataGridView7.Item(2, r).Value
            If ComboBox8.SelectedItem = "横" Then
                str &= vbTab
                If r <> DataGridView7.Rows.Count - 1 Then
                    If ComboBox9.SelectedItem = "連続" Then
                        str &= ""
                    ElseIf ComboBox9.SelectedItem = "1つ空き" Then
                        str &= vbTab
                    Else
                        str &= vbTab & vbTab
                    End If
                End If
            Else
                str &= vbCrLf
                If r <> DataGridView7.Rows.Count - 1 Then
                    If ComboBox9.SelectedItem = "連続" Then
                        str &= ""
                    ElseIf ComboBox9.SelectedItem = "1つ空き" Then
                        str &= vbCrLf
                    Else
                        str &= vbCrLf & vbCrLf
                    End If
                End If
            End If
        Next
        Clipboard.SetText(str)
    End Sub

    '列データ移行
    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        Dim DGV1 As DataGridView = csvForm.DataGridView1
        Dim dH1 As ArrayList = TM_HEADER_1ROW_GET(csvForm.DataGridView1)

        Dim selArray As New ArrayList
        For k As Integer = 0 To DGV1.SelectedCells.Count - 1
            If Not selArray.Contains(DGV1.SelectedCells(k).RowIndex) Then
                selArray.Add(DGV1.SelectedCells(k).RowIndex)
            End If
        Next

        csvForm.selChangeFlag = False
        For k As Integer = 0 To selArray.Count - 1
            Dim selRow As Integer = selArray(k)
            Dim motoStr As String = DGV1.Item(dH1.IndexOf(TextBox32.Text), selRow).Value
            Dim sakiStr As String = ""
            If motoStr <> "" Then
                Dim m As MatchCollection = Regex.Matches(motoStr, TextBox31.Text)
                For i As Integer = 0 To m.Count - 1
                    If TextBox34.Text <> "" Then
                        If Not Regex.IsMatch(m(i).Value, TextBox34.Text) Then
                            sakiStr &= m(i).Value
                            If RadioButton14.Checked Then
                                sakiStr &= vbLf
                            ElseIf RadioButton15.Checked Then
                                sakiStr &= vbCr
                            Else
                                sakiStr &= vbCrLf
                            End If
                        End If
                    Else
                        sakiStr &= m(i).Value
                        If RadioButton14.Checked Then
                            sakiStr &= vbLf
                        ElseIf RadioButton15.Checked Then
                            sakiStr &= vbCr
                        Else
                            sakiStr &= vbCrLf
                        End If
                    End If
                Next
                DGV1.Item(dH1.IndexOf(TextBox33.Text), selRow).Value = sakiStr
                DGV1.Item(dH1.IndexOf(TextBox33.Text), selRow).Style.BackColor = Color.Yellow
                DGV1.CurrentCell = DGV1(dH1.IndexOf(TextBox33.Text), selRow)
            End If
        Next
        csvForm.selChangeFlag = True

        MsgBox("終了しました")
    End Sub

    Private Sub TabPage4_GotFocus(sender As Object, e As EventArgs) Handles TabPage4.GotFocus
        If Form1.CSV_FORMS(1).ToolStripStatusLabel5.Text = "2" Then
            csvForm = Form1.CSV_FORMS(1)
        Else
            csvForm = csvForm
        End If
    End Sub
End Class