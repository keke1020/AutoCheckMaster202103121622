Imports EncodingOperation
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic.FileIO
Imports System.IO
Imports ZetaHtmlTidy
Imports Hnx8
Imports System.Net
Imports NPOI.HSSF.UserModel
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel

Public Class Csv
    Dim secValue As String = ""

    Dim undoStack As New Stack(Of Object()())
    Dim redoStack As New Stack(Of Object()())
    Dim ignore As Boolean = False
    Dim cellChange As Boolean = True
    Dim cellChangeArray As ArrayList = New ArrayList

    Dim sc As Integer = 0
    Dim sr As Integer = 0

    Dim splitter As String = ","    '区切り文字


    Private Sub Form5_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        AddHandler My.Settings.SettingChanging, AddressOf Settings_SettingChanging
        secValue = "csv"   'iniファイルセクション

        '******************************************
        'dgv画面チラつき防止
        '******************************************
        'DataGirdViewのTypeを取得
        Dim dgvtype As Type = GetType(DataGridView)
        'プロパティ設定の取得
        Dim dgvPropertyInfo As Reflection.PropertyInfo = dgvtype.GetProperty("DoubleBuffered", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)

        '対象のDataGridViewにtrueをセットする
        Dim dgvArray As DataGridView() = {DataGridView1, DataGridView2}
        For i As Integer = 0 To dgvArray.Length - 1
            dgvPropertyInfo.SetValue(dgvArray(i), True, Nothing)
        Next

        If Form1.inif.ReadINI(secValue, "csvSortMode", True) Then
            注文順自動並べ替えToolStripMenuItem.Checked = True
        Else
            注文順自動並べ替えToolStripMenuItem.Checked = False
        End If

        If Form1.inif.ReadINI(secValue, "csvSortAfterSave", True) Then
            並べ替え後自動上書き保存ToolStripMenuItem.Checked = True
        Else
            並べ替え後自動上書き保存ToolStripMenuItem.Checked = False
        End If

        Me.TopMost = True
        Me.TopMost = False
    End Sub

    Private Sub Form5_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If e.CloseReason = CloseReason.FormOwnerClosing Or e.CloseReason = CloseReason.UserClosing Then
            Me.Visible = False
            e.Cancel = True
        End If
    End Sub

    'フォームのサイズ最小化バグ修正
    Private Sub Settings_SettingChanging(ByVal sender As Object, ByVal e As System.Configuration.SettingChangingEventArgs)
        If Not Me.WindowState = FormWindowState.Normal Then
            If e.SettingName = "Form5Size" Or e.SettingName = "Form5Location" Then
                e.Cancel = True
            End If
        End If
    End Sub

    '*****************************************************************************************
    'undo-redo(extensions.vb)
    Private Sub DataGridView1_CellValidated(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellValidated
        If ignore Then Return
        If undoStack.NeedsItem(DataGridView1) Then
            undoStack.Push(DataGridView1.Rows.Cast(Of DataGridViewRow).Where(Function(r) Not r.IsNewRow).Select(Function(r) r.Cells.Cast(Of DataGridViewCell).Select(Function(c) c.Value).ToArray).ToArray)
        End If
        元に戻すToolStripMenuItem.Enabled = undoStack.Count > 1
    End Sub

    Private Sub 元に戻すToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 元に戻すToolStripMenuItem.Click
        Dim dgBack As ArrayList = New ArrayList
        For r As Integer = 0 To DataGridView1.RowCount - 1
            For c As Integer = 0 To DataGridView1.ColumnCount - 1
                Select Case DataGridView1.Item(c, r).Style.BackColor.Name
                    Case 0, Color.White.Name

                    Case Else
                        dgBack.Add(c & splitter & r & splitter & DataGridView1.Item(c, r).Style.BackColor.Name)
                End Select
            Next
        Next
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        If redoStack.Count = 0 OrElse redoStack.NeedsItem(DataGridView1) Then
            redoStack.Push(DataGridView1.Rows.Cast(Of DataGridViewRow).Where(Function(r) Not r.IsNewRow).Select(Function(r) r.Cells.Cast(Of DataGridViewCell).Select(Function(c) c.Value).ToArray).ToArray)
        End If
        Dim rows()() As Object = undoStack.Pop
        While rows.ItemEquals(DataGridView1.Rows.Cast(Of DataGridViewRow).Where(Function(r) Not r.IsNewRow).ToArray)
            rows = undoStack.Pop
        End While
        ignore = True
        DataGridView1.Rows.Clear()
        For x As Integer = 0 To rows.GetUpperBound(0)
            DataGridView1.Rows.Add(rows(x))
        Next
        ignore = False
        元に戻すToolStripMenuItem.Enabled = undoStack.Count > 0
        やり直しToolStripMenuItem.Enabled = redoStack.Count > 0
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
        For i As Integer = 0 To dgBack.Count - 1
            Dim cr As String() = Split(dgBack(i), splitter)
            Dim colorName = New ColorConverter().ConvertFromString(cr(2))   'カラーコンバーター
            DataGridView1.Item(CInt(cr(0)), CInt(cr(1))).Style.BackColor = colorName
        Next
    End Sub

    Private Sub やり直しToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles やり直しToolStripMenuItem.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        If undoStack.Count = 0 OrElse undoStack.NeedsItem(DataGridView1) Then
            undoStack.Push(DataGridView1.Rows.Cast(Of DataGridViewRow).Where(Function(r) Not r.IsNewRow).Select(Function(r) r.Cells.Cast(Of DataGridViewCell).Select(Function(c) c.Value).ToArray).ToArray)
        End If
        Dim rows()() As Object = redoStack.Pop
        While rows.ItemEquals(DataGridView1.Rows.Cast(Of DataGridViewRow).Where(Function(r) Not r.IsNewRow).ToArray)
            rows = redoStack.Pop
        End While
        ignore = True
        DataGridView1.Rows.Clear()
        For x As Integer = 0 To rows.GetUpperBound(0)
            DataGridView1.Rows.Add(rows(x))
        Next
        ignore = False
        やり直しToolStripMenuItem.Enabled = redoStack.Count > 0
        元に戻すToolStripMenuItem.Enabled = undoStack.Count > 0
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub
    '*****************************************************************************************


    Private Sub 検索と置換ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 検索と置換ToolStripMenuItem.Click
        Search.Size = New Size(412, 454)
        'Search.TabControl1.SelectedIndex = 0
        Form1.SetTabVisible(Search.TabControl1, "検索と置換")
        Search.Show()
    End Sub

    Public dataPathArray As New ArrayList
    Public headerAdd As Boolean = True
    Public addFlag As Boolean = True
    Private Sub DataGridView1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles DataGridView1.DragDrop
        dataPathArray.Clear()

        Dim fCount As Integer = 0
        For Each filename As String In e.Data.GetData(DataFormats.FileDrop)
            dataPathArray.Add(filename)
            fCount += 1
        Next

        If dataPathArray.Count > 1 Then
            CSV_F_files.DGV1.Rows.Clear()
            For i As Integer = 0 To dataPathArray.Count - 1
                CSV_F_files.DGV1.Rows.Add(True, Path.GetFileName(dataPathArray(i)))
            Next
            CSV_F_files.dataPathArray = dataPathArray
            Dim DR As DialogResult = CSV_F_files.ShowDialog()
            If DR = DialogResult.OK Then
                DropRun(dataPathArray, fCount, headerAdd, addFlag)
            Else
                Exit Sub
            End If
        Else
            DropRun(dataPathArray, fCount, True, True)
        End If
    End Sub

    Private Shared Rows_Filename As New ArrayList
    Public Sub DropRun(dataPathArray As ArrayList, fCount As Integer, Optional headerAdd As Boolean = False, Optional addFlag As Boolean = True)
        If DataGridView1.RowCount > 1 Then
            If ToolStripStatusLabel4.Text = "Qoo10追跡番号リストモード" Then
                If fCount <> 1 Then
                    Dim DR2 As DialogResult = MsgBox("1ファイル以外は同時に解析できません。", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
                    Exit Sub
                End If
                For Each filename As String In dataPathArray
                    Qoo10Tuiseki(filename)
                    Exit Sub
                Next
            ElseIf ToolStripStatusLabel4.Text = "Qoo10発送予定リストモード" Then
                If fCount <> 1 Then
                    Dim DR2 As DialogResult = MsgBox("1ファイル以外は同時に解析できません。", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
                    Exit Sub
                End If
                For Each filename As String In dataPathArray
                    Qoo10Hassou(filename)
                    Exit Sub
                Next
            ElseIf ToolStripStatusLabel4.Text = "ヤフオクモード" Then
                If fCount <> 1 Then
                    Dim DR2 As DialogResult = MsgBox("1ファイル以外は同時に解析できません。", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
                    Exit Sub
                End If
                For Each filename As String In dataPathArray
                    YahooAucNyukin(filename)
                    Exit Sub
                Next
            Else    '通常モード
                If addFlag Then
                    If dataPathArray.Count = 1 Then
                        Dim DR As DialogResult = MsgBox("追加しますか？", MsgBoxStyle.YesNoCancel Or MsgBoxStyle.SystemModal)
                        If DR = DialogResult.No Then
                            DataGridView1.Columns.Clear()
                            checkA = 0
                            checkB = 0
                            checkC = 0
                        ElseIf DR = DialogResult.Cancel Then
                            Exit Sub
                        End If
                    End If
                Else
                    DataGridView1.Columns.Clear()
                    checkA = 0
                    checkB = 0
                    checkC = 0
                End If
            End If
        Else
            If ToolStripStatusLabel4.Text = "在庫チェックモード" Then
                If fCount <> 2 Then
                    Dim DR2 As DialogResult = MsgBox("2ファイル以外は同時に解析できません。", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
                    Exit Sub
                End If
                Dim DR As DialogResult = MsgBox("在庫変動のチェックをします", MsgBoxStyle.OkCancel Or MsgBoxStyle.SystemModal)
                If DR = Windows.Forms.DialogResult.Cancel Then
                    Exit Sub
                End If
            End If
        End If

        DataGridView2.Visible = False
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------

        Rows_Filename.Clear()   '行毎のファイル名を保存する
        Dim num As Integer = 0
        Dim colorArray As Color() = {Color.LightYellow, Color.LightCyan}
        For Each filename As String In dataPathArray
            Dim csvRecords As New ArrayList()
            csvRecords = CSV_READ(filename)

            If csvRecords Is Nothing Then
                Exit For
            End If

            Dim startRowNum As Integer = 0
            If num = 0 And DataGridView1.RowCount > 0 Then
                Dim DR As DialogResult = MsgBox("追加ファイルのヘッダーを削除して良いですか？", MsgBoxStyle.YesNo Or MsgBoxStyle.SystemModal)
                If DR = DialogResult.Yes Then
                    startRowNum = 1
                Else
                    startRowNum = 0
                End If
            ElseIf num > 0 And Not headerAdd Then
                startRowNum = 1
            End If

            For r As Integer = startRowNum To csvRecords.Count - 1
                Dim sArray As String() = Split(csvRecords(r), "|=|")
                If DataGridView1.ColumnCount < sArray.Length - 1 Then
                    For c As Integer = DataGridView1.ColumnCount To sArray.Length - 1
                        DataGridView1.Columns.Add(c, c)
                        DataGridView1.Columns(c).SortMode = DataGridViewColumnSortMode.NotSortable
                    Next
                End If
                DataGridView1.Rows.Add(sArray)
                If num > 0 Then
                    Dim cNum As Integer = num Mod 2
                    DataGridView1.Item(0, DataGridView1.RowCount - 2).Style.BackColor = colorArray(cNum) '追加した時だけ、一時的に色を付ける
                End If

                Rows_Filename.Add(filename)
            Next

            If num = 0 Then
                Me.Text = filename
            End If
            num += 1
        Next

        DGVoneHeadCheck()

        If 注文順自動並べ替えToolStripMenuItem.Checked = True Then
            ChumonSort(0)
        End If

        Dim checkStr As String = ""
        For c As Integer = 0 To DataGridView1.ColumnCount - 1
            checkStr &= DataGridView1.Item(c, 0).Value
        Next
        If InStr(checkStr, "pathnamecode") > 0 Then
            YahooCheck()
        End If
        If InStr(checkStr, "コントロールカラム商品管理番号") > 0 Then
            RakutenCheck()
        End If
        If InStr(checkStr, "商品特定コード指定更新時間フラグ") > 0 Then
            MakeShopCheck()
        End If
        If InStr(checkStr, "Item CodeSeller Code") > 0 Then
            Qoo10Check()
        End If

        '-----------------------------------------------------------------------
        If DataGridView1.Rows.Count > 0 Then
            RetDIV()
        End If

        '-----------------------------------------------------------------------

        Me.TopMost = True
        Me.TopMost = False
    End Sub

    Private Sub DataGridView1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles DataGridView1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub YahooCheck()
        For c As Integer = 0 To DataGridView1.ColumnCount - 1
            Dim res As String = DataGridView1.Item(c, 0).Value
            If res <> "" Then
                Select Case True
                    'Yahoo用
                    Case Regex.IsMatch(res, "^path|name|^code|^price|sale-price")
                        DataGridView1.Item(c, 0).Style.BackColor = Color.LightYellow
                    Case Regex.IsMatch(res, "sub-code|options|product-category|spec1$|spec2|spec3|spec4|spec5")
                        DataGridView1.Item(c, 0).Style.BackColor = Color.LightBlue
                    Case Regex.IsMatch(res, "sale-period-start|sale-period-end")
                        DataGridView1.Item(c, 0).Style.BackColor = Color.GreenYellow
                    Case Regex.IsMatch(res, "pr-rate")
                        DataGridView1.Item(c, 0).Style.BackColor = Color.Red
                End Select
            End If
        Next
    End Sub

    Private Sub RakutenCheck()
        '↓楽天用
        For c As Integer = 0 To DataGridView1.ColumnCount - 1
            Dim res As String = DataGridView1.Item(c, 0).Value
            If res <> "" Then
                Select Case True
                    Case Regex.IsMatch(res, "コントロールカラム|商品管理番号（商品URL）|商品番号")
                        DataGridView1.Item(c, 0).Style.BackColor = Color.LightYellow
                    Case Regex.IsMatch(res, "項目選択肢用コントロールカラム|商品管理番号（商品URL）|選択肢タイプ|項目選択肢別在庫用横軸選択肢|項目選択肢別在庫用横軸選択肢子番号|項目選択肢別在庫用縦軸選択肢|項目選択肢別在庫用縦軸選択肢子番号")
                        DataGridView1.Item(c, 0).Style.BackColor = Color.LightYellow
                End Select
            End If
        Next
    End Sub

    Private Sub MakeShopCheck()
        '↓MakeShop用
        For c As Integer = 0 To DataGridView1.ColumnCount - 1
            Dim res As String = DataGridView1.Item(c, 0).Value
            If res <> "" Then
                Select Case True
                    Case Regex.IsMatch(res, "商品特定コード指定")
                        DataGridView1.Item(c, 0).Style.BackColor = Color.Orange
                    Case Regex.IsMatch(res, "独自商品コード|^商品名$|販売価格")
                        DataGridView1.Item(c, 0).Style.BackColor = Color.LightBlue
                    Case Regex.IsMatch(res, "更新時間フラグ|システム商品コード|カテゴリ名|サブカテゴリ名|重量|定価|^ポイント$|仕入価格|製造元|原産地|原産地表示フラグ|数量|数量表示フラグ|最小注文限度数|最大注文限度数")
                        DataGridView1.Item(c, 0).Style.BackColor = Color.LightYellow
                    Case Regex.IsMatch(res, "陳列位置|送料個別設定|割引使用フラグ|割引率|割引期間|^商品グループ$|商品検索語|商品別特殊表示|オプション１名称|オプション２名称|オプショングループ")
                        DataGridView1.Item(c, 0).Style.BackColor = Color.LightYellow
                    Case Regex.IsMatch(res, "拡大画像名|普通画像名|縮小画像名|モバイル画像名|モバイル商品説明|^商品説明$|JANコード|商品表示可否")
                        DataGridView1.Item(c, 0).Style.BackColor = Color.LightYellow
                End Select
            End If
        Next
    End Sub

    Private Sub Qoo10Check()
        '↓Qoo10用
        For c As Integer = 0 To DataGridView1.ColumnCount - 1
            Dim res As String = DataGridView1.Item(c, 0).Value
            If res <> "" Then
                Select Case True
                    Case Regex.IsMatch(res, "Item Code|Seller Code|Status|2nd Cat Code|Item Name|Origin Type")
                        DataGridView1.Item(c, 0).Style.BackColor = Color.LightYellow
                    Case Regex.IsMatch(res, "Sell Price|Sell Qty")
                        DataGridView1.Item(c, 0).Style.BackColor = Color.LightBlue
                End Select
            End If
        Next
    End Sub

    Private Sub DGVoneHeadCheck()
        Dim str As String = ""
        For c As Integer = 0 To DataGridView1.ColumnCount - 1
            str &= DataGridView1.Item(c, 0).Value
        Next

        Select Case True
            Case Regex.IsMatch(str, ".*商品名.*|.*注文日.*|.*受注番号.*|.*syohin_code.*|.*price.*|.*path.*|.*商品管理番号.*|.*品名.*|.*OrderId.*")
                CheckBox1.Checked = True
            Case Else
                CheckBox1.Checked = False
        End Select
    End Sub

    Private Function CSV_READ2(ByRef path As String)    'これは使用しない
        Dim csvRecords As New ArrayList()

        'CSVファイル名
        Dim csvFileName As String = path

        'TextFieldParserは空行を読み飛ばすので、対策。
        'Dim csvStr As String = ""
        'Dim fName As String = csvFileName
        'ToolStripStatusLabel8.Text = 0
        'Using sr As New System.IO.StreamReader(fName, System.Text.Encoding.Default)
        '    Do While Not sr.EndOfStream
        '        Dim s As String = sr.ReadLine
        '        If s = "" Then
        '            csvStr &= "{LF}" & vbLf
        '        Else
        '            csvStr &= s & vbLf
        '        End If
        '        ToolStripStatusLabel8.Text += 1
        '        Application.DoEvents()
        '    Loop
        'End Using

        'Dim csvStr As String = System.IO.File.ReadAllText(csvFileName, System.Text.Encoding.Default)
        'csvStr = Replace(csvStr, vbLf, "{LF}")



        'Shift JISで読み込む
        Dim e1 As EncodingJP
        Dim e2 As Boolean
        TextFileEncoding.LoadFromFile(csvFileName, e1, True, e2)    'エンコードを調べる
        ToolStripDropDownButton10.Text = DecodingName(e1)
        Dim tfp As New TextFieldParser(csvFileName, EncodingName(e1)) With {
            .TextFieldType = FileIO.FieldType.Delimited,
            .Delimiters = New String() {splitter},
            .HasFieldsEnclosedInQuotes = True,
            .TrimWhiteSpace = False
        }
        'Dim rs As New System.IO.StringReader(csvStr)
        'Dim tfp As New FileIO.TextFieldParser(rs)

        While Not tfp.EndOfData
            'フィールドを読み込む
            Dim fields As String() = tfp.ReadFields()
            Dim str As String = ""
            Dim doubleStr As Double = 0
            For i As Integer = 0 To fields.Length - 1
                'fields(i) = Replace(fields(i), vbLf, vbCrLf)
                If i = 0 Then
                    str = fields(i)
                Else
                    str &= "|=|" & fields(i)
                End If
            Next
            '保存
            csvRecords.Add(str)
        End While

        'For i As Integer = 0 To csvRecords.Count - 1
        '    csvRecords(i) = Replace(csvRecords(i), "{LF}", "")
        'Next

        '後始末
        tfp.Close()
        Return csvRecords
    End Function

    Public Function CSV_READ(ByRef path As String)
        Dim csvText As String = ""
        Dim file As FileInfo = New FileInfo(path)
        Dim fileEnc As String = ""
        Using reader As ReadJEnc.FileReader = New ReadJEnc.FileReader(file)
            Dim c As ReadJEnc.CharCode = reader.Read(file)
            'fileEnc = c.Name
            csvText = reader.Text
        End Using
        'If fileEnc = "ShiftJIS" Then
        '    fileEnc = "Shift-JIS"
        'End If
        'ToolStripDropDownButton10.Text = fileEnc
        'ToolStripDropDownButton10.Text = "Shift-JIS"

        'Dim csvText As String = file.ReadAllText(path, System.Text.Encoding.Default)

        If csvText Is Nothing Then
            MsgBox(path + "を読み込めないです", MsgBoxStyle.SystemModal)
            Exit Function
        End If
        'カンマ、タブ等の区切り文字認識
        Dim header As String() = Regex.Split(csvText, vbCrLf & "|" & vbCr & "|" & vbLf)
        If InStr(header(0), vbTab) Then
            splitter = vbTab
            ToolStripStatusLabel9.Text = "tab"
        Else
            splitter = ","
            ToolStripStatusLabel9.Text = "comma"
        End If

        '前後の改行を削除しておく
        csvText = csvText.Trim(
            New Char() {ControlChars.Cr, ControlChars.Lf})

        Dim csvRecords As New System.Collections.ArrayList
        Dim csvFields As New System.Collections.ArrayList

        Dim csvTextLength As Integer = csvText.Length
        Dim startPos As Integer = 0
        Dim endPos As Integer = 0
        Dim field As String = ""

        While True
            '空白を飛ばす
            'While startPos < csvTextLength _
            '    AndAlso (csvText.Chars(startPos) = " "c OrElse csvText.Chars(startPos) = ControlChars.Tab)
            '    startPos += 1
            'End While

            'データの最後の位置を取得
            If startPos < csvTextLength _
                AndAlso csvText.Chars(startPos) = ControlChars.Quote Then
                '"で囲まれているとき
                '最後の"を探す
                endPos = startPos
                While True
                    endPos = csvText.IndexOf(ControlChars.Quote, endPos + 1)
                    If endPos < 0 Then
                        Throw New ApplicationException("""が不正")
                    End If
                    '"が2つ続かない時は終了
                    If endPos + 1 = csvTextLength OrElse csvText.Chars((endPos + 1)) <> ControlChars.Quote Then
                        Exit While
                    End If
                    '"が2つ続く
                    endPos += 1
                End While

                '一つのフィールドを取り出す
                field = csvText.Substring(startPos, endPos - startPos + 1)
                '""を"にする
                field = field.Substring(1, field.Length - 2).Replace("""""", """")

                endPos += 1
                '空白を飛ばす
                While endPos < csvTextLength AndAlso
                    csvText.Chars(endPos) <> splitter AndAlso
                    csvText.Chars(endPos) <> ControlChars.Lf
                    endPos += 1
                End While
            Else
                '"で囲まれていない
                'カンマか改行の位置
                endPos = startPos
                While endPos < csvTextLength AndAlso
                    csvText.Chars(endPos) <> splitter AndAlso
                    csvText.Chars(endPos) <> ControlChars.Lf
                    endPos += 1
                End While

                '一つのフィールドを取り出す
                field = csvText.Substring(startPos, endPos - startPos)
                '後の空白を削除
                field = field.TrimEnd()
            End If

            'フィールドの追加
            csvFields.Add(field)

            '行の終了か調べる
            If endPos >= csvTextLength OrElse
                csvText.Chars(endPos) = ControlChars.Lf Then
                '行の終了
                'レコードの追加
                csvFields.TrimToSize()
                Dim str As String = ""
                For i As Integer = 0 To csvFields.Count - 1
                    If i = 0 Then
                        str = csvFields(i)
                    Else
                        str &= "|=|" & csvFields(i)
                    End If
                Next
                'csvRecords.Add(csvFields)
                csvRecords.Add(str)
                csvFields = New System.Collections.ArrayList(csvFields.Count)

                If endPos >= csvTextLength Then
                    '終了
                    Exit While
                End If
            End If

            '次のデータの開始位置
            startPos = endPos + 1
        End While

        csvRecords.TrimToSize()

        For r As Integer = 0 To csvRecords.Count - 1
            If InStr(csvRecords(r), "E+") Then
                MsgBox("エクセルで保存し直した形跡があります。データを作成し直してください", MsgBoxStyle.SystemModal)
                Exit For
            End If
        Next

        'If Format(System.IO.File.GetCreationTime(path), "yyyy/MM/dd H:mm:ss") <> Format(System.IO.File.GetLastWriteTime(path), "yyyy/MM/dd H:mm:ss") Then
        '    MsgBox("エクセルで保存し直した形跡があります。データを作成し直してください", MsgBoxStyle.SystemModal)
        'End If

        Return csvRecords
    End Function

    Private Function EncodingName(ByVal e As EncodingJP) As System.Text.Encoding
        Select Case e
            Case EncodingJP.ASCII
                Return System.Text.Encoding.ASCII
            Case EncodingJP.UTF8
                Return System.Text.Encoding.UTF8
            Case EncodingJP.UTF7
                Return System.Text.Encoding.UTF7
            Case EncodingJP.SJIS
                Return System.Text.Encoding.GetEncoding("SHIFT-JIS")
            Case Else
                Return System.Text.Encoding.Default
        End Select
    End Function

    Private Function EncodingName2(ByVal e As String) As System.Text.Encoding
        Select Case e
            Case "欧文"
                Return System.Text.Encoding.ASCII
            Case "UTF-8"
                Return System.Text.Encoding.UTF8
            Case "UTF-7"
                Return System.Text.Encoding.UTF7
            Case "SJIS"
                Return System.Text.Encoding.GetEncoding("SHIFT-JIS")
            Case Else
                Return System.Text.Encoding.Default
        End Select
    End Function

    Private Function DecodingName(ByVal e As String) As String
        Select Case e
            Case EncodingJP.ASCII
                Return "欧文"
            Case EncodingJP.SJIS
                Return "SJIS"
            Case EncodingJP.EUC
                Return "EUC"
            Case EncodingJP.JIS
                Return "JIS"
            Case EncodingJP.UTF16_LE
                Return "Unicode"
            Case EncodingJP.UTF16_BE
                Return "Unicode(Big Endian)"
            Case EncodingJP.UTF8
                Return "UTF-8"
            Case EncodingJP.UTF8_BOM
                Return "UTF-8(BOM)"
            Case EncodingJP.UTF7
                Return "UTF-7"
            Case Else
                Return ""
        End Select
    End Function

    'スクロール
    Private Sub DataGridView1_Scroll(ByVal sender As Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles DataGridView1.Scroll
        Try
            If DataGridView2.Visible = True Then
                DataGridView2.FirstDisplayedScrollingRowIndex = DataGridView1.FirstDisplayedScrollingRowIndex
                'DataGridView2.HorizontalScrollingOffset = DataGridView1.HorizontalScrollingOffset
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub DataGridView2_Scroll(ByVal sender As Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles DataGridView2.Scroll
        Try
            DataGridView1.FirstDisplayedScrollingRowIndex = DataGridView2.FirstDisplayedScrollingRowIndex
            'DataGridView1.HorizontalScrollingOffset = DataGridView2.HorizontalScrollingOffset
        Catch ex As Exception

        End Try
    End Sub

    Private Sub 新規作成ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 新規作成ToolStripMenuItem.Click
        If DataGridView1.RowCount = 0 Then
            DataGridView1.Rows.Clear()
            DataGridView1.Columns.Clear()
        End If

        For c As Integer = 0 To 10
            Dim textColumn As New DataGridViewTextBoxColumn With {
                .Name = "Column" & c,
                .HeaderText = c
            }
            DataGridView1.Columns.Add(textColumn)
        Next
    End Sub

    Private Sub 上書き保存ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 上書き保存ToolStripMenuItem.Click
        If Me.Text <> "..." Then
            Dim FileName As String = Me.Text
            SaveCsv(FileName, 0)
            MsgBox("上書き保存しました", MsgBoxStyle.SystemModal)
        Else
            名前を付けて保存ToolStripMenuItem_Click(sender, e)
        End If
    End Sub

    Private Sub 名前を付けて保存全てで囲むToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 名前を付けて保存全てで囲むToolStripMenuItem.Click
        NameChangeSave(1)
    End Sub

    Private Sub 名前を付けて保存ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 名前を付けて保存ToolStripMenuItem.Click
        NameChangeSave(0)
    End Sub

    Private Sub 名前を付けて保存タブ区切りToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 名前を付けて保存タブ区切りToolStripMenuItem.Click
        NameChangeSave(2)
    End Sub

    Private Sub NameChangeSave(ByVal mode As Integer)
        If DataGridView1.RowCount = 0 Then
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
            SaveCsv(sfd.FileName, mode)
            MsgBox(sfd.FileName & vbCrLf & "保存しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    Private Sub 名前を付けて保存エクセルxlsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 名前を付けて保存エクセルxlsToolStripMenuItem.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If

        Dim sfd As New SaveFileDialog()
        If InStr(Me.Text, "\") > 0 Then
            Dim sPath As String = IO.Path.GetFileNameWithoutExtension(Me.Text)
            sfd.FileName = sPath & ".xlsx"
        Else
            sfd.FileName = "新しいファイル.xlsx"
        End If
        sfd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        sfd.Filter = "XLSXファイル(*.xlsx)|*.xlsx|すべてのファイル(*.*)|*.*"
        sfd.AutoUpgradeEnabled = False
        sfd.FilterIndex = 0
        sfd.Title = "保存先のファイルを選択してください"
        sfd.RestoreDirectory = True
        sfd.OverwritePrompt = True
        sfd.CheckPathExists = True

        DataGridView1.EndEdit()

        'ダイアログを表示する
        If sfd.ShowDialog(Me) = DialogResult.OK Then
            Dim visibleFlag As Boolean = False
            If マイナス在庫ToolStripMenuItem.Checked = True Or 変更有りのみToolStripMenuItem.Checked = True Then
                visibleFlag = True
            End If

            Dim xWorkbook As IWorkbook = New XSSFWorkbook()
            Dim sheet1 As ISheet = xWorkbook.CreateSheet("Sheet1")

            Dim sheet As Integer = 0

            Dim inRow As Integer = 0
            For r As Integer = 0 To DataGridView1.RowCount - 2
                If visibleFlag = False Or (visibleFlag = True And DataGridView1.Rows(r).Visible = True) Then
                    sheet1.CreateRow(r)
                    For c As Integer = 0 To DataGridView1.ColumnCount - 1
                        If Not DataGridView1.Item(c, r).Value Is DBNull.Value Then
                            Dim v As String = DataGridView1.Item(c, r).Value
                            If IsNumeric(v) Then
                                sheet1.GetRow(r).CreateCell(c).SetCellValue(CDbl(v))
                            Else
                                sheet1.GetRow(r).CreateCell(c).SetCellValue(v)
                            End If
                        End If
                    Next
                    inRow += 1
                End If
            Next

            'Dim desk As String = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
            Using wfs = File.Create(sfd.FileName)
                xWorkbook.Write(wfs)
            End Using
            sheet1 = Nothing
            xWorkbook = Nothing

            MsgBox(sfd.FileName & vbCrLf & "保存しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    '*******************************************************************************
    'ToolStripDropDownButton2
    '*******************************************************************************
    Private Sub 切り取りToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 切り取りToolStripMenuItem.Click
        Dim dgv As DataGridView = DataGridView1
        If AzukiControl1.BackColor = Color.White Then
            AzukiControl1.Cut()
        ElseIf TextBox2.BackColor = Color.White Then
            TextBox2.Cut()
        Else
            '-----------------------------------------------------------------------
            PreDIV()
            '-----------------------------------------------------------------------
            Dim selCell = dgv.SelectedCells
            CUTS(dgv)
            '-----------------------------------------------------------------------
            RetDIV()
            '-----------------------------------------------------------------------
        End If
    End Sub

    Private Sub コピーToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles コピーToolStripMenuItem.Click
        Dim dgv As DataGridView = DataGridView1
        If AzukiControl1.BackColor = Color.White Then
            AzukiControl1.Copy()
        ElseIf TextBox2.BackColor = Color.White Then
            TextBox2.Copy()
        Else
            Dim selCell = dgv.SelectedCells
            COPYS(dgv, 1)
        End If
    End Sub

    Private Sub コピー囲み無しToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles コピー囲み無しToolStripMenuItem.Click
        Dim dgv As DataGridView = DataGridView1
        If AzukiControl1.BackColor = Color.White Then
            AzukiControl1.Copy()
        ElseIf TextBox2.BackColor = Color.White Then
            TextBox2.Copy()
        Else
            Dim selCell = dgv.SelectedCells
            COPYS(dgv, 0)
        End If
    End Sub

    Private Sub 貼り付けToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 貼り付けToolStripMenuItem.Click
        Dim dgv As DataGridView = DataGridView1
        If AzukiControl1.BackColor = Color.White Then
            AzukiControl1.Paste()
        ElseIf TextBox2.BackColor = Color.White Then
            TextBox2.Paste()
        Else
            '-----------------------------------------------------------------------
            PreDIV()
            '-----------------------------------------------------------------------
            Dim selCell = dgv.SelectedCells
            PASTES(dgv, selCell)
            '-----------------------------------------------------------------------
            RetDIV()
            '-----------------------------------------------------------------------
        End If
    End Sub

    Private Sub 削除ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 削除ToolStripMenuItem.Click
        Dim dgv As DataGridView = DataGridView1
        If AzukiControl1.BackColor = Color.White Then
            AzukiControl1.Cut()
        ElseIf TextBox2.BackColor = Color.White Then
            TextBox2.Cut()
        Else
            If Not DataGridView1.IsCurrentCellInEditMode Then
                '-----------------------------------------------------------------------
                PreDIV()
                '-----------------------------------------------------------------------
                Dim selCell = dgv.SelectedCells
                DELS(dgv, selCell)
                '-----------------------------------------------------------------------
                RetDIV()
                '-----------------------------------------------------------------------
            End If
        End If
    End Sub

    Private Sub 上に挿入ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 上に挿入ToolStripMenuItem1.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim dgv As DataGridView = DataGridView1
        Dim selCell = dgv.SelectedCells
        ROWSADD(dgv, selCell, 1)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 下に挿入ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 下に挿入ToolStripMenuItem1.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim dgv As DataGridView = DataGridView1
        Dim selCell = dgv.SelectedCells
        ROWSADD(dgv, selCell, 0)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 右に挿入ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 右に挿入ToolStripMenuItem1.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim dgv As DataGridView = DataGridView1
        Dim selCell = dgv.SelectedCells
        COLSADD(dgv, selCell, 0)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 左に挿入ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 左に挿入ToolStripMenuItem1.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim dgv As DataGridView = DataGridView1
        Dim selCell = dgv.SelectedCells
        COLSADD(dgv, selCell, 1)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 右に挿入ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 右に挿入ToolStripMenuItem.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim dgv As DataGridView = DataGridView1
        Dim selCell = dgv.SelectedCells
        COLSADD(dgv, selCell, 0)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 左に挿入ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 左に挿入ToolStripMenuItem.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim dgv As DataGridView = DataGridView1
        Dim selCell = dgv.SelectedCells
        COLSADD(dgv, selCell, 1)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 行を削除ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 行を削除ToolStripMenuItem.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim dgv As DataGridView = DataGridView1
        Dim selCell = dgv.SelectedCells
        ROWSCUT(dgv, selCell)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
        'selCell(selCell.Count - 1).Selected = True
        'DataGridView1.CurrentCell = selCell(selCell.Count - 1)
    End Sub

    Private Sub 非表示行を削除ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 非表示行を削除ToolStripMenuItem.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim dgv As DataGridView = DataGridView1
        For r As Integer = dgv.RowCount - 1 To 0 Step -1
            If dgv.Rows(r).IsNewRow = False And dgv.Rows(r).Visible = False Then
                dgv.Rows.RemoveAt(r)
            End If
        Next
        dgv.Refresh()
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
        MsgBox("削除処理しました", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
    End Sub

    Private Sub 列を削除ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 列を削除ToolStripMenuItem.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim dgv As DataGridView = DataGridView1
        Dim selCell = dgv.SelectedCells
        COLSCUT(dgv, selCell)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 列選択ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 列選択ToolStripMenuItem.Click
        ColSelect(0)
    End Sub

    Private Sub 列選択ヘッダー除外ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 列選択ヘッダー除外ToolStripMenuItem1.Click
        ColSelect(1)
    End Sub

    Private Sub ColSelect(ByVal mode As Integer)
        Panel2.Visible = True
        Panel2.Location = New Point(SplitContainer1.Width / 2 - 40, SplitContainer1.Height / 2 - 30)
        Application.DoEvents()

        selChangeFlag = False

        Dim dgv As DataGridView = DataGridView1
        Dim selCell = dgv.SelectedCells

        Dim selCol As ArrayList = New ArrayList
        For i As Integer = 0 To selCell.Count - 1
            If Not selCol.Contains(selCell(i).ColumnIndex) Then
                selCol.Add(selCell(i).ColumnIndex)
            End If
        Next

        Dim start As Integer = 0
        If mode <> 0 Then
            start = 1
        End If
        For i As Integer = 0 To selCol.Count - 1
            For r As Integer = start To DataGridView1.RowCount - 1
                If DataGridView1.Rows(r).IsNewRow Then
                    Continue For
                End If
                If DataGridView1.Rows(r).Visible = True Then
                    DataGridView1.Item(selCol(i), r).selected = True
                End If
            Next
        Next

        selChangeFlag = True
        Panel2.Visible = False
    End Sub

    Private Sub 列選択モードToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 列選択モードToolStripMenuItem.Click
        If 列選択モードToolStripMenuItem.Text = "列選択モード" Then
            列選択モードToolStripMenuItem.Text = "列選択モード解除"
            DataGridView1.SelectionMode = DataGridViewSelectionMode.FullColumnSelect
        Else
            列選択モードToolStripMenuItem.Text = "列選択モード"
            DataGridView1.SelectionMode = DataGridViewSelectionMode.CellSelect
        End If
    End Sub

    Private Sub 列選択ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 列選択ToolStripMenuItem1.Click
        列選択ToolStripMenuItem.PerformClick()
    End Sub

    Private Sub 列選択ヘッダー除外ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 列選択ヘッダー除外ToolStripMenuItem.Click
        列選択ヘッダー除外ToolStripMenuItem1.PerformClick()
    End Sub

    Private Sub 全て選択ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 全て選択ToolStripMenuItem.Click
        Dim dgv As DataGridView = DataGridView1
        If AzukiControl1.BackColor = Color.White Then
            AzukiControl1.SelectAll()
        ElseIf TextBox2.BackColor = Color.White Then
            TextBox2.SelectAll()
        Else
            dgv.SelectAll()
        End If
    End Sub

    Public Sub PreDIV()
        '-----------------------------------------------------------------------
        'sc = DataGridView1.FirstDisplayedScrollingColumnIndex
        sc = DataGridView1.HorizontalScrollingOffset
        sr = DataGridView1.FirstDisplayedScrollingRowIndex
        'For r As Integer = 0 To DataGridView1.RowCount - 1
        '    If DataGridView1.Rows(r).Visible And Not DataGridView1.Rows(r).IsNewRow Then
        '        sr = DataGridView1.FirstDisplayedScrollingRowIndex
        '    End If
        'Next
        cellChange = False
        'DataGridView1.ScrollBars = ScrollBars.None
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
        DataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        DataGridView1.ScrollBars = ScrollBars.Both
        '-----------------------------------------------------------------------
    End Sub

    Public Sub RetDIV()
        '-----------------------------------------------------------------------
        If セルの自動幅調整ToolStripMenuItem.Checked = True Then
            DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        End If
        cellChange = True
        '***************************************************
        'If sc > DataGridView1.ColumnCount - 2 Then
        '    sc = DataGridView1.ColumnCount - 2
        'End If
        'If sc < 0 Then
        '    sc = 0
        'End If
        'DataGridView1.FirstDisplayedScrollingColumnIndex = sc
        DataGridView1.HorizontalScrollingOffset = sc
        '***************************************************
        'If sr < 0 Then
        '    sr = 0
        'End If
        'For r As Integer = sr To DataGridView1.RowCount - 1
        '    If DataGridView1.Rows(r).Visible And Not DataGridView1.Rows(r).IsNewRow Then
        '        'sr = DataGridView1.FirstDisplayedScrollingRowIndex
        '        sr = r
        '    End If
        'Next

        If sr > DataGridView1.RowCount - 2 Then
            sr = DataGridView1.RowCount - 2
        End If
        If sr < 0 Then
            sr = 0
        End If
        YahooCheck()    'Yahooチェック
        Try
            DataGridView1.FirstDisplayedScrollingRowIndex = sr
        Catch ex As Exception

        End Try

        DataGridView1.Rows(0).Frozen = True
        DataGridView1.Rows(0).DividerHeight = 3
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 選択セルの左で固定ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 選択セルの左で固定ToolStripMenuItem.Click
        Dim dgv As DataGridView = DataGridView1
        Dim selCell = dgv.SelectedCells

        For c As Integer = 0 To dgv.ColumnCount - 1
            dgv.Columns(c).Frozen = False
            dgv.Columns(c).DividerWidth = 0
        Next

        Dim koteiR As New ArrayList
        For i As Integer = 0 To selCell.Count - 1
            If Not koteiR.Contains(selCell(i).ColumnIndex) Then
                koteiR.Add(selCell(i).ColumnIndex)
            End If
        Next
        koteiR.Sort()
        Dim fdsc As Integer = dgv.FirstDisplayedScrollingColumnIndex
        If koteiR(0) > 0 Then
            For i As Integer = 0 To dgv.Columns.Count - 1
                If i < fdsc Then
                    dgv.Columns(i).Visible = False
                Else
                    If koteiR(0) >= fdsc Then
                        dgv.Columns(koteiR(0) - 1).Frozen = True
                        dgv.Columns(koteiR(0) - 1).DividerWidth = 3
                        Exit For
                    End If
                End If
            Next
        End If
    End Sub

    Private Sub 固定解除ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 固定解除ToolStripMenuItem.Click
        Dim dgv As DataGridView = DataGridView1
        For c As Integer = 0 To dgv.Columns.Count - 1
            dgv.Columns(c).Visible = True
        Next
        For c As Integer = 0 To dgv.ColumnCount - 1
            If dgv.Columns(c).Frozen = True Then
                dgv.Columns(c).Frozen = False
            Else
                dgv.Columns(c).DividerWidth = 0
                'Exit Sub
            End If
        Next
    End Sub

    'Private Sub 選択セルの左で固定ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 選択セルの左で固定ToolStripMenuItem.Click
    '    Dim dgv As DataGridView = DataGridView1
    '    Dim selCell = dgv.SelectedCells

    '    For c As Integer = 0 To dgv.ColumnCount - 1
    '        dgv.Columns(c).Frozen = False
    '        dgv.Columns(c).DividerWidth = 0
    '    Next

    '    Dim koteiR As New ArrayList
    '    For i As Integer = 0 To selCell.Count - 1
    '        If Not koteiR.Contains(selCell(i).ColumnIndex) Then
    '            koteiR.Add(selCell(i).ColumnIndex)
    '        End If
    '    Next
    '    koteiR.Sort()
    '    If koteiR(0) > 0 Then
    '        dgv.Columns(koteiR(0) - 1).Frozen = True
    '        dgv.Columns(koteiR(0) - 1).DividerWidth = 3
    '    End If
    'End Sub

    'Private Sub 固定解除ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 固定解除ToolStripMenuItem.Click
    '    Dim dgv As DataGridView = DataGridView1
    '    For c As Integer = 0 To dgv.ColumnCount - 1
    '        If dgv.Columns(c).Frozen = True Then
    '            dgv.Columns(c).Frozen = False
    '        Else
    '            dgv.Columns(c).DividerWidth = 0
    '            'Exit Sub
    '        End If
    '    Next
    'End Sub

    Private Sub セルの自動幅調整ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles セルの自動幅調整ToolStripMenuItem.Click
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        セルの自動幅調整ToolStripMenuItem.Checked = True
        セルの自動幅調整切ToolStripMenuItem.Checked = False
    End Sub

    Private Sub セルの自動幅調整切ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles セルの自動幅調整切ToolStripMenuItem.Click
        DataGridView1.ScrollBars = ScrollBars.None
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        DataGridView1.ScrollBars = ScrollBars.Both
        セルの自動幅調整ToolStripMenuItem.Checked = False
        セルの自動幅調整切ToolStripMenuItem.Checked = True
    End Sub

    Private Sub 上へToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 上へToolStripMenuItem.Click
        If DataGridView1.SelectedCells.Count > 0 Then
            Dim dgRow As ArrayList = New ArrayList
            For i As Integer = 0 To DataGridView1.SelectedCells.Count - 1
                If dgRow.Contains(DataGridView1.SelectedCells(i).RowIndex) = False Then
                    dgRow.Add(DataGridView1.SelectedCells(i).RowIndex)
                End If
            Next
            dgRow.Sort()

            If dgRow(0) > 0 Then
                DataGridView1.ClearSelection()
                Dim ueData As ArrayList = New ArrayList
                For c As Integer = 0 To DataGridView1.ColumnCount - 1
                    ueData.Add(DataGridView1.Item(c, dgRow(0) - 1).Value)
                Next

                For r As Integer = 0 To dgRow.Count - 1
                    For c As Integer = 0 To DataGridView1.ColumnCount - 1
                        DataGridView1.Item(c, dgRow(r) - 1).Value = DataGridView1.Item(c, dgRow(r)).Value
                        DataGridView1.Item(c, dgRow(r) - 1).Selected = True
                    Next
                Next

                For c As Integer = 0 To DataGridView1.ColumnCount - 1
                    DataGridView1.Item(c, dgRow(dgRow.Count - 1)).Value = ueData(c)
                Next
            End If
        End If
    End Sub

    Private Sub 下へToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 下へToolStripMenuItem.Click
        If DataGridView1.SelectedCells.Count > 0 Then
            Dim dgRow As ArrayList = New ArrayList
            For i As Integer = 0 To DataGridView1.SelectedCells.Count - 1
                If dgRow.Contains(DataGridView1.SelectedCells(i).RowIndex) = False Then
                    dgRow.Add(DataGridView1.SelectedCells(i).RowIndex)
                End If
            Next
            dgRow.Sort()
            dgRow.Reverse()

            If dgRow(0) < DataGridView1.RowCount - 2 Then
                DataGridView1.ClearSelection()
                Dim sitaData As ArrayList = New ArrayList
                For c As Integer = 0 To DataGridView1.ColumnCount - 1
                    sitaData.Add(DataGridView1.Item(c, dgRow(0) + 1).Value)
                Next

                For r As Integer = 0 To dgRow.Count - 1
                    For c As Integer = 0 To DataGridView1.ColumnCount - 1
                        DataGridView1.Item(c, dgRow(r) + 1).Value = DataGridView1.Item(c, dgRow(r)).Value
                        DataGridView1.Item(c, dgRow(r) + 1).Selected = True
                    Next
                Next

                For c As Integer = 0 To DataGridView1.ColumnCount - 1
                    DataGridView1.Item(c, dgRow(dgRow.Count - 1)).Value = sitaData(c)
                Next
            End If
        End If
    End Sub

    '*******************************************************************************
    'ToolStripDropDownButton2ここまで
    '*******************************************************************************

    '*******************************************************************************
    'ContextMenuStrip
    '*******************************************************************************
    Private Sub 切り取りToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 切り取りToolStripMenuItem1.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim dgv As DataGridView = DataGridView1
        Dim selCell = dgv.SelectedCells
        CUTS(dgv)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub コピーToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles コピーToolStripMenuItem1.Click
        Dim dgv As DataGridView = DataGridView1
        Dim selCell = dgv.SelectedCells
        COPYS(dgv, 1)
    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        Dim dgv As DataGridView = DataGridView1
        Dim selCell = dgv.SelectedCells
        COPYS(dgv, 0)
    End Sub

    Private Sub 貼り付けToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 貼り付けToolStripMenuItem1.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim dgv As DataGridView = DataGridView1
        Dim selCell = dgv.SelectedCells
        PASTES(dgv, selCell)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 下に挿入ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 下に挿入ToolStripMenuItem.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim dgv As DataGridView = DataGridView1
        Dim selCell = dgv.SelectedCells
        ROWSADD(dgv, selCell, 0)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 上に挿入ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 上に挿入ToolStripMenuItem.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim dgv As DataGridView = DataGridView1
        Dim selCell = dgv.SelectedCells
        ROWSADD(dgv, selCell, 1)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 行を選択直下に複製ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 行を選択直下に複製ToolStripMenuItem.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim dgv As DataGridView = DataGridView1

        Dim dgRow As ArrayList = New ArrayList
        For i As Integer = 0 To dgv.SelectedCells.Count - 1
            If dgRow.Contains(dgv.SelectedCells(i).RowIndex) = False Then
                dgRow.Add(dgv.SelectedCells(i).RowIndex)
            End If
        Next
        dgRow.Sort()
        dgRow.Reverse()

        For r As Integer = 0 To dgRow.Count - 1
            dgv.Rows.InsertCopy(dgRow(r), dgRow(r) + 1)
            For c As Integer = 0 To dgv.ColumnCount - 1
                If dgv.Item(c, dgRow(r)).Value <> "" Then
                    dgv.Item(c, dgRow(r) + 1).Value = dgv.Item(c, dgRow(r)).Value
                End If
            Next
        Next
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 行を上へ移動ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 行を上へ移動ToolStripMenuItem.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim dgv As DataGridView = DataGridView1

        Dim dgRow As ArrayList = New ArrayList
        For i As Integer = 0 To dgv.SelectedCells.Count - 1
            If dgRow.Contains(dgv.SelectedCells(i).RowIndex) = False Then
                dgRow.Add(dgv.SelectedCells(i).RowIndex)
            End If
        Next
        dgRow.Sort()
        dgRow.Reverse()

        For r As Integer = 0 To dgRow.Count - 1
            If dgRow(r) = 0 Or dgRow(r) = 1 Then
                Continue For
            End If
            dgv.Rows.InsertCopy(dgRow(r), dgRow(r) + 1)
            For c As Integer = 0 To dgv.ColumnCount - 1
                If dgv.Item(c, dgRow(r)).Value <> "" Then
                    dgv.Item(c, dgRow(r) + 1).Value = dgv.Item(c, dgRow(r) - 1).Value
                End If
            Next
            dgv.Rows.RemoveAt(dgRow(r) - 1)
        Next
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 行を下へ移動ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 行を下へ移動ToolStripMenuItem.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim dgv As DataGridView = DataGridView1

        Dim dgRow As ArrayList = New ArrayList
        For i As Integer = 0 To dgv.SelectedCells.Count - 1
            If dgRow.Contains(dgv.SelectedCells(i).RowIndex) = False Then
                dgRow.Add(dgv.SelectedCells(i).RowIndex)
            End If
        Next
        dgRow.Sort()
        dgRow.Reverse()

        For r As Integer = 0 To dgRow.Count - 1
            If dgRow(r) = 0 Or dgRow(r) + 2 > dgv.RowCount - 1 Then
                Continue For
            End If
            dgv.Rows.InsertCopy(dgRow(r), dgRow(r) + 2)
            For c As Integer = 0 To dgv.ColumnCount - 1
                If dgv.Item(c, dgRow(r)).Value <> "" Then
                    dgv.Item(c, dgRow(r) + 2).Value = dgv.Item(c, dgRow(r)).Value
                End If
            Next
            dgv.Rows.RemoveAt(dgRow(r))
            dgv.CurrentCell = dgv(dgv.SelectedCells(0).ColumnIndex, dgRow(r) + 1)
        Next
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 削除ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 削除ToolStripMenuItem1.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim dgv As DataGridView = DataGridView1
        Dim selCell = dgv.SelectedCells
        DELS(dgv, selCell)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 行を削除ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 行を削除ToolStripMenuItem1.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim dgv As DataGridView = DataGridView1
        Dim selCell = dgv.SelectedCells
        selChangeFlag = False
        ROWSCUT(dgv, selCell)
        selChangeFlag = True
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 空行一括削除ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 空行一括削除ToolStripMenuItem.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim dgv As DataGridView = DataGridView1
        For r As Integer = dgv.RowCount - 1 To 0 Step -1
            If dgv.Rows(r).IsNewRow Then
                Continue For
            End If
            Dim delFlag As Boolean = True
            For c As Integer = 0 To dgv.ColumnCount - 1
                If CStr(dgv.Item(c, r).Value) <> "" Then
                    delFlag = False
                    Continue For
                End If
            Next
            If delFlag Then
                dgv.Rows.RemoveAt(r)
            End If
        Next
        selChangeFlag = True
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 列を削除ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 列を削除ToolStripMenuItem1.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim dgv As DataGridView = DataGridView1
        Dim selCell = dgv.SelectedCells
        COLSCUT(dgv, selCell)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 指定配列から空要素を削除ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles 指定配列から空要素を削除ToolStripMenuItem2.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim dgv As DataGridView = DataGridView1
        Dim selCell = dgv.SelectedCells

        Dim delC As New ArrayList
        For i As Integer = 0 To selCell.Count - 1
            If Not delC.Contains(selCell(i).ColumnIndex) Then
                delC.Add(selCell(i).ColumnIndex)
            End If
        Next

        For r As Integer = dgv.RowCount - 1 To 0 Step -1
            If dgv.Rows(r).IsNewRow Then
                Continue For
            End If
            Dim delFlag As Boolean = True
            For c As Integer = 0 To delC.Count - 1
                If CStr(dgv.Item(delC(c), r).Value) <> "" Then
                    delFlag = False
                    Continue For
                Else
                    Exit For
                End If
            Next
            If delFlag Then
                dgv.Rows.RemoveAt(r)
                Continue For
            End If
        Next
    End Sub

    '↑↑=======================================================================↑↑
    'ContextMenuStripここまで
    '===============================================================================

    ' ----------------------------------------------------------------------------------
    ' DataGridViewでCtrl-Vキー押下時にクリップボードから貼り付けるための実装↓↓↓↓↓↓
    ' DataGridViewでDelやBackspaceキー押下時にセルの内容を消去するための実装↓↓↓↓↓↓
    ' ----------------------------------------------------------------------------------
    Public Shared ColumnChars() As String
    Private _editingColumn As Integer   '編集中のカラム番号
    Private _editingCtrl As DataGridViewTextBoxEditingControl

    Private Sub DataGridView_EditingControlShowing(
        ByVal sender As Object, ByVal e As DataGridViewEditingControlShowingEventArgs) _
        Handles DataGridView1.EditingControlShowing
        ' 編集中のカラム番号を保存
        _editingColumn = CType(sender, DataGridView).CurrentCellAddress.X
        Try
            ' 編集中のTextBoxEditingControlにKeyPressイベント設定
            _editingCtrl = CType(e.Control, DataGridViewTextBoxEditingControl)
            AddHandler _editingCtrl.KeyPress, AddressOf DataGridView_CellKeyPress
        Catch
        End Try
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

    Private Sub DataGridView_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DataGridView1.KeyUp
        Dim dgv As DataGridView = CType(sender, DataGridView)
        Dim selCell = dgv.SelectedCells

        If e.KeyCode = Keys.Back Or e.KeyCode = Keys.Delete Then    ' セルの内容を消去
            If Not DataGridView1.IsCurrentCellInEditMode Then
                DELS(dgv, selCell)
            End If
        ElseIf e.Modifiers = Keys.Control And e.KeyCode = Keys.Enter Then '複数セル一括入力
            '-----------------------------------------------------------------------
            PreDIV()
            '-----------------------------------------------------------------------
            Dim str As String = dgv.CurrentCell.Value
            For i As Integer = 0 To selCell.Count - 1
                If dgv.Rows(selCell(i).RowIndex).Visible = True Then    '表示している行のみ
                    If selCell(i).Value <> str Then
                        selCell(i).Value = str
                        selCell(i).Style.BackColor = Color.Yellow
                    End If
                End If
            Next
            '-----------------------------------------------------------------------
            RetDIV()
            '-----------------------------------------------------------------------
        End If
    End Sub
    ' ----------------------------------------------------------------------------------
    ' DataGridViewでCtrl-Vキー押下時にクリップボードから貼り付けるための実装↑↑↑↑↑↑
    ' DataGridViewでDelやBackspaceキー押下時にセルの内容を消去するための実装↑↑↑↑↑↑
    ' ----------------------------------------------------------------------------------


    '*******************************************************************************
    '切り取り、コピー、貼り付けなど
    '*******************************************************************************
    Private Sub DELS(ByVal dgv As DataGridView, ByVal selCell As Object)
        'Dim dgv As DataGridView = DataGridView1
        'Dim selCell = dgv.SelectedCells
        For i As Integer = 0 To selCell.Count - 1
            If dgv.Rows(selCell(i).RowIndex).Visible Then
                dgv(selCell(i).ColumnIndex, selCell(i).RowIndex).Value = ""
            End If
        Next
    End Sub

    Public Shared DGV1tb As TextBox
    Private Sub DataGridView1_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        If TypeOf e.Control Is DataGridViewTextBoxEditingControl Then
            Dim dgv As DataGridView = CType(sender, DataGridView)

            '編集のために表示されているコントロールを取得
            DGV1tb = CType(e.Control, TextBox)
        End If
    End Sub

    Private Sub COPYS(ByVal dgv As DataGridView, ByVal mode As Integer)
        If dgv.GetCellCount(DataGridViewElementStates.Selected) > 0 Then
            If dgv.CurrentCell.IsInEditMode = True Then
                DGV1tb.Copy()
            Else
                Try
                    If dgv.SelectedCells.Count > 0 Then
                        Dim lstStrBuf As New List(Of String)
                        For Each cell As DataGridViewCell In dgv.SelectedCells
                            If IsDBNull(cell.Value) Then
                                lstStrBuf.Add(Nothing)
                            Else
                                lstStrBuf.Add(cell.Value)
                                'cell.Value = (cell.Value.ToString()).Replace(vbNewLine, vbLf)
                                cell.Value = Replace(cell.Value, vbCrLf, vbLf)
                                If mode = 1 Then    '通常コピーはダブルコーテーションで囲む
                                    cell.Value = Replace(cell.Value, """", """""")
                                    cell.Value = """" & cell.Value & """"
                                End If
                            End If
                        Next

                        'Clipboard.SetText(dgv.GetClipboardContent.GetText)
                        Dim returnValue As DataObject
                        returnValue = dgv.GetClipboardContent()
                        Dim strTab = returnValue.GetText()
                        Clipboard.SetDataObject(strTab, True)

                        Dim i As Integer = 0
                        For Each cell As DataGridViewCell In dgv.SelectedCells
                            cell.Value = lstStrBuf(i)
                            i += 1
                        Next
                    End If
                Catch ex As System.Runtime.InteropServices.ExternalException
                    MessageBox.Show("コピーできませんでした。")
                End Try
            End If
        End If
    End Sub

    Private Sub CUTS(ByVal dgv As DataGridView)
        If dgv.GetCellCount(DataGridViewElementStates.Selected) > 0 Then
            If dgv.CurrentCell.IsInEditMode = True Then
                DGV1tb.Cut()
            Else
                Try
                    Clipboard.SetText(dgv.GetClipboardContent.GetText)

                    '切り取りなので、選択セルの内容をクリアする
                    For i As Integer = 0 To dgv.SelectedCells.Count - 1
                        Dim c As Integer = dgv.SelectedCells.Item(i).ColumnIndex
                        Dim r As Integer = dgv.SelectedCells.Item(i).RowIndex
                        dgv.Item(c, r).Value = ""
                    Next i
                Catch ex As System.Runtime.InteropServices.ExternalException
                    MessageBox.Show("コピーできませんでした。")
                End Try
            End If
        End If
    End Sub

    Private Sub PASTES(ByVal dgv As DataGridView, ByVal selCell As Object)
        Dim x As Integer = dgv.CurrentCellAddress.X
        Dim y As Integer = dgv.CurrentCellAddress.Y
        ' クリップボードの内容を取得
        Dim clipText As String = Clipboard.GetText()

        ' 改行を変換
        'clipText = clipText.Replace(vbCrLf, vbLf)
        'clipText = clipText.Replace(vbCr, vbLf)
        ' 改行で分割
        'Dim lines() As String = clipText.Split(vbNewLine)

        If AzukiControl1.BackColor = Color.White Then
            AzukiControl1.Paste()
        ElseIf TextBox2.BackColor = Color.White Then
            TextBox2.Paste()
        Else
            '================================================================
            '改行を含むテキストを取得するためTextFieldParserで解析
            Dim csvRecords As New ArrayList()
            'Dim clipText2 As String = Replace(clipText, vbLf, "|=|")
            'Dim clipCSV As String() = Split(clipText2, vbCrLf)

            'For i As Integer = 0 To clipCSV.Length - 1
            '    Dim rs As New StringReader(clipText)
            '    Dim tfp As New TextFieldParser(rs) With {
            '        .TextFieldType = FieldType.Delimited,
            '        .Delimiters = New String() {vbTab, splitter},
            '        .HasFieldsEnclosedInQuotes = True,
            '        .TrimWhiteSpace = False
            '    }

            '    While Not tfp.EndOfData
            '        Dim str As String = ""
            '        Dim fields As String() = tfp.ReadFields()   'フィールドを読み込む

            '        For k As Integer = 0 To fields.Length - 1
            '            If k = 0 Then
            '                str = fields(k)
            '            Else
            '                str &= vbTab & fields(k)
            '            End If
            '        Next
            '        csvRecords.Add(str)
            '    End While

            '    tfp.Close()
            '    rs.Close()
            'Next

            clipText = Replace(clipText, vbCrLf, "<<CRLF>>")
            clipText = Replace(clipText, vbCr, "<<CR>>")
            clipText = Replace(clipText, vbLf, "<<LF>>")
            Dim clipTextArray As String() = Split(clipText, "<<CRLF>>")
            'Dim clipTextArray As String() = Split(clipText, vbCrLf)

            For i As Integer = 0 To clipTextArray.Length - 1
                Dim flag As Boolean = False
                Dim rs As New StringReader(clipTextArray(i))
                Using tfp As New TextFieldParser(rs)
                    tfp.TextFieldType = FieldType.Delimited
                    tfp.Delimiters = New String() {vbTab} ', splitter}
                    tfp.HasFieldsEnclosedInQuotes = True
                    tfp.TrimWhiteSpace = False

                    While Not tfp.EndOfData
                        Dim str As String = ""
                        Dim fields As String() = tfp.ReadFields()   'フィールドを読み込む

                        For k As Integer = 0 To fields.Length - 1
                            If k = 0 Then
                                str = fields(k)
                            Else
                                str &= vbTab & fields(k)
                            End If
                        Next
                        str = Replace(str, "<<CR>>", vbCr)
                        str = Replace(str, "<<LF>>", vbLf)
                        csvRecords.Add("" & str)
                        flag = True
                    End While
                End Using
                rs.Close()

                'データが無い行をチェックし追加する
                If flag = False Then
                    csvRecords.Add("")
                End If
            Next

            '================================================================

            Dim lines() As String = Nothing
            lines = DirectCast(csvRecords.ToArray(GetType(String)), String())

            Dim oneFlag As Boolean = False
            If lines.Length = 1 Then
                Dim vals() As String = lines(0).Split(vbTab)
                If vals.Length = 1 Then
                    oneFlag = True
                End If
            End If

            If oneFlag = True Then
                If dgv.IsCurrentCellInEditMode = False Then
                    For i As Integer = 0 To selCell.count - 1
                        If dgv.Item(selCell(i).ColumnIndex, selCell(i).RowIndex).visible Then
                            dgv(selCell(i).ColumnIndex, selCell(i).RowIndex).Value = lines(0)
                            dgv(selCell(i).ColumnIndex, selCell(i).RowIndex).style.backcolor = Color.Yellow
                        End If
                    Next
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
                        dgv(x + c2, y + r).Style.BackColor = Color.Yellow
                        ' ------------------------------------------------------
                        ' 次のセルへ
                        c2 = c2 + 1
                    Next
                Next
            End If
        End If

    End Sub

    Private Sub ROWSCUT(ByVal dgv As DataGridView, ByVal selCell As Object)
        Dim delR As New ArrayList
        For i As Integer = 0 To selCell.Count - 1
            If Not delR.Contains(selCell(i).RowIndex) Then
                delR.Add(selCell(i).RowIndex)
            End If
        Next
        delR.Sort()

        'delR.Reverse()
        If delR.Count > 100 Then
            Panel2.Visible = True
            dgv.Visible = False
            Application.DoEvents()
        End If
        For r As Integer = delR.Count - 1 To 0 Step -1
            If delR(r) < dgv.RowCount - 1 Then
                dgv.Rows.RemoveAt(delR(r))
            End If
        Next
        If delR.Count > 100 Then
            dgv.Visible = True
            Panel2.Visible = False
        End If
    End Sub

    Private Sub ROWSADD(ByVal dgv As DataGridView, ByVal selCell As Object, ByVal mode As Integer)
        Dim delR As New ArrayList
        For i As Integer = 0 To selCell.Count - 1
            If Not delR.Contains(selCell(i).RowIndex) Then
                delR.Add(selCell(i).RowIndex)
            End If
        Next
        delR.Sort()
        If mode = 1 Then
            dgv.Rows.Insert(delR(0), delR.Count)
        Else
            delR.Reverse()
            If delR(0) + 1 >= dgv.RowCount - 1 Then
                dgv.Rows.Add(delR.Count)
            Else
                delR(0) += 1
                dgv.Rows.Insert(delR(0), delR.Count)
            End If
        End If
    End Sub

    Private Sub COLSADD(ByVal dgv As DataGridView, ByVal selCell As Object, ByVal mode As Integer)
        Dim delR As New ArrayList
        For i As Integer = 0 To selCell.Count - 1
            If Not delR.Contains(selCell(i).ColumnIndex) Then
                delR.Add(selCell(i).ColumnIndex)
            End If
        Next
        delR.Sort()
        Dim num As String = dgv.ColumnCount
        Dim textColumn As New DataGridViewTextBoxColumn With {
            .Name = "Column" & num,
            .HeaderText = num
        }
        If mode = 1 Then
            dgv.Columns.Insert(delR(0), textColumn)
        Else
            delR.Reverse()
            delR(0) += 1
            If delR(0) + 1 <= dgv.ColumnCount - 1 Then
                dgv.Columns.Insert(delR(0), textColumn)
            Else
                dgv.Columns.Add(textColumn)
            End If
        End If
    End Sub

    Private Sub COLSCUT(ByVal dgv As DataGridView, ByVal selCell As Object)
        Dim delC As New ArrayList
        For i As Integer = 0 To selCell.Count - 1
            If Not delC.Contains(selCell(i).ColumnIndex) Then
                delC.Add(selCell(i).ColumnIndex)
            End If
        Next
        delC.Sort()
        delC.Reverse()
        For c As Integer = 0 To delC.Count - 1
            If delC(c) < dgv.ColumnCount Then
                dgv.Columns.RemoveAt(delC(c))
            End If
        Next
    End Sub

    '↑↑=======================================================================↑↑
    'ContextMenuStripここまで
    '===============================================================================

    ' DataGridViewをCSV出力する
    Public Sub SaveCsv(ByVal fp As String, ByVal mode As Integer)
        Dim visibleFlag As Boolean = False
        If マイナス在庫ToolStripMenuItem.Checked = True Or 変更有りのみToolStripMenuItem.Checked = True Then
            visibleFlag = True
        Else
            For r As Integer = 0 To DataGridView1.Rows.Count - 2
                If DataGridView1.Rows(r).Visible = False Then
                    Dim DR As DialogResult = MsgBox("表示行のみ保存しますか？", MsgBoxStyle.YesNo Or MsgBoxStyle.SystemModal)
                    If DR = Windows.Forms.DialogResult.Yes Then
                        visibleFlag = True
                        Exit For
                    Else
                        Exit For
                    End If
                End If
            Next
        End If

        DataGridView1.EndEdit()

        Dim CsvDelimiter As String = splitter
        If mode = 2 Then
            CsvDelimiter = vbTab
        End If

        ToolStripProgressBar1.Maximum = DataGridView1.Rows.Count
        ToolStripProgressBar1.Value = 0
        ' CSVファイルオープン
        Dim str As String = ""
        'Dim sw As StreamWriter = New StreamWriter(fp, False, EncodingName2(ToolStripDropDownButton10.Text))
        Dim sw As StreamWriter = New StreamWriter(fp, False, System.Text.Encoding.GetEncoding(ToolStripDropDownButton10.Text))
        For r As Integer = 0 To DataGridView1.Rows.Count - 2
            If visibleFlag = False Or (visibleFlag = True And DataGridView1.Rows(r).Visible = True) Then
                For c As Integer = 0 To DataGridView1.Columns.Count - 1
                    ' DataGridViewのセルのデータ取得
                    Dim dt As String = ""
                    If DataGridView1.Rows(r).Cells(c).Value Is Nothing = False Then
                        dt = DataGridView1.Rows(r).Cells(c).Value.ToString()
                        dt = Replace(dt, vbCrLf, vbLf)
                        'dt = Replace(dt, vbLf, "")
                        If mode = 1 Then
                            dt = """" & dt & """"
                        Else
                            If Not dt Is Nothing Then
                                dt = Replace(dt, """", """""")
                                Select Case True
                                    Case InStr(dt, splitter), InStr(dt, vbLf), InStr(dt, vbCr), InStr(dt, """") 'Regex.IsMatch(dt, ".*,.*|.*" & vbLf & ".*|.*" & vbCr & ".*|.*"".*")
                                        dt = """" & dt & """"
                                End Select
                            End If
                        End If
                    End If
                    If c < DataGridView1.Columns.Count - 1 Then
                        dt = dt & CsvDelimiter
                    End If
                    ' CSVファイル書込
                    sw.Write(dt)
                Next
                sw.Write(vbLf)
            End If
            ToolStripProgressBar1.Value += 1
            If ToolStripProgressBar1.Value Mod 500 = 0 Then
                Application.DoEvents()
            End If
        Next

        ' CSVファイルクローズ
        sw.Close()

        ToolStripProgressBar1.Value = 0

        Me.Text = fp
    End Sub


    '行番号を表示する
    Private Sub DataGridView1_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView1.RowPostPaint
        ' 行ヘッダのセル領域を、行番号を描画する長方形とする（ただし右端に4ドットのすき間を空ける）
        Dim rect As New Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, DataGridView1.RowHeadersWidth - 4, DataGridView1.Rows(e.RowIndex).Height)

        ' 上記の長方形内に行番号を縦方向中央＆右詰で描画する　フォントや色は行ヘッダの既定値を使用する
        TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), DataGridView1.RowHeadersDefaultCellStyle.Font, rect, DataGridView1.RowHeadersDefaultCellStyle.ForeColor, TextFormatFlags.VerticalCenter Or TextFormatFlags.Right)
    End Sub

    '===============================================================================
    'チェックここから
    '↓↓=======================================================================↓↓
    Dim checkA As Integer = 0
    Dim checkB As Integer = 0
    Dim checkC As Integer = 0

    Private Sub DefaltCell()
        If DataGridView1.RowCount > 0 Then
            DataGridView2.Visible = True
            DataGridView2.Rows.Clear()
            For i As Integer = 0 To DataGridView2.SelectedCells.Count - 1
                DataGridView2.SelectedCells(i).Selected = False
            Next i

            For r As Integer = 0 To DataGridView1.RowCount - 2
                DataGridView1.Rows(r).DefaultCellStyle.BackColor = Color.White
                DataGridView2.Rows.Add("")
            Next

            DataGridView2.Rows(0).Frozen = True
            DataGridView2.Rows(0).DividerHeight = 3
        End If
    End Sub

    Private Sub 重複ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 重複ToolStripMenuItem.Click
        DefaltCell()
        Juuhuku(0)
        DataGridView1.CurrentCell = DataGridView1(checkA, 0)
    End Sub

    Private Sub 重複チェック行全て対象ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 重複チェック行全て対象ToolStripMenuItem.Click
        DefaltCell()
        Juuhuku(1)
        DataGridView1.CurrentCell = DataGridView1(checkA, 0)
    End Sub

    Private Sub Juuhuku(mode As Integer)
        Dim checkArray As New ArrayList
        For r As Integer = 1 To DataGridView1.RowCount - 1
            If DataGridView1.Rows(r).IsNewRow Then
                Continue For
            End If

            Dim nowStr As String = ""
            For c As Integer = 0 To DataGridView1.ColumnCount - 1
                If mode = 0 Then
                    Select Case DataGridView1.Item(c, 0).Value
                        Case "伝票番号", "受注日", "状態", "受注番号", "店舗", "購入者名", "送り先名", "送り先〒", "送り先住所", "発送方法", "明細行", "商品ｺｰﾄﾞ", "jan", "伝票no", "受注no", "受注名", "お届け先名", "お届け先郵便番号", "お届け先住所", "お届け先電話", "発送先名", "発送先郵便番号", "発送先住所", "発送先電話番号"
                            nowStr &= DataGridView1.Item(c, r).Value
                    End Select
                Else
                    nowStr &= DataGridView1.Item(c, r).Value
                End If
            Next

            If checkArray.Contains(nowStr) Then
                DataGridView1.Rows(r).DefaultCellStyle.BackColor = Color.Khaki
                DataGridView2.Item(0, r).Value &= "重複:" & checkArray.IndexOf(nowStr) + 1
            Else
                checkArray.Add(nowStr)
            End If
        Next
    End Sub

    Private Sub DataGridView2_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellDoubleClick
        Dim jStr As String = DataGridView2.CurrentCell.Value
        If InStr(jStr, "重複:") > 0 Then
            Dim gyou As String() = Split(jStr, ":")
            CSV_F_juhuku.Show()
            CSV_F_juhuku.DataGridView1.Rows.Clear()
            CSV_F_juhuku.DataGridView1.Columns.Clear()

            CSV_F_juhuku.DataGridView1.Columns.Add(0, "行番号")
            CSV_F_juhuku.DataGridView1.Columns.Add(0, "ファイル名")
            For c As Integer = 0 To DataGridView1.ColumnCount - 1
                CSV_F_juhuku.DataGridView1.Columns.Add(c + 1, DataGridView1.Item(c, 0).Value)
            Next

            CSV_F_juhuku.DataGridView1.Rows.Add()
            Dim nRow As Integer = CSV_F_juhuku.DataGridView1.RowCount - 1
            CSV_F_juhuku.DataGridView1.Item(0, nRow).Value = gyou(1)
            CSV_F_juhuku.DataGridView1.Item(1, nRow).Value = Path.GetFileNameWithoutExtension(Rows_Filename(CInt(gyou(1))))
            For c As Integer = 0 To DataGridView1.ColumnCount - 1
                CSV_F_juhuku.DataGridView1.Item(c + 2, nRow).Value = DataGridView1.Item(c, CInt(gyou(1))).Value
            Next

            For rr As Integer = 0 To DataGridView2.RowCount - 1
                If DataGridView2.Item(0, rr).Value = jStr Then
                    CSV_F_juhuku.DataGridView1.Rows.Add()
                    nRow = CSV_F_juhuku.DataGridView1.RowCount - 1
                    CSV_F_juhuku.DataGridView1.Item(0, nRow).Value = rr
                    CSV_F_juhuku.DataGridView1.Item(1, nRow).Value = Path.GetFileNameWithoutExtension(Rows_Filename(rr))
                    For c As Integer = 0 To DataGridView1.ColumnCount - 1
                        CSV_F_juhuku.DataGridView1.Item(c + 2, nRow).Value = DataGridView1.Item(c, rr).Value
                    Next
                End If
            Next
        End If
    End Sub

    Private Sub チェック欄非表示ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles チェック欄非表示ToolStripMenuItem.Click
        DataGridView2.Visible = False
    End Sub
    '↑↑=======================================================================↑↑
    'ContextMenuStripここまで
    '===============================================================================


    Private Sub フォームのリセットToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles フォームのリセットToolStripMenuItem.Click
        'Me.Dispose()
        'Dim frm As Form = New csv
        'frm.Show()
        DataGridView1.Columns.Clear()
        AzukiControl1.Text = ""
        TextBox2.Text = ""
        DataGridView2.Rows.Clear()
    End Sub

    Private Sub 通常モードToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 通常モードToolStripMenuItem.Click
        ToolStripStatusLabel4.Text = "通常モード"
        通常モードToolStripMenuItem.Checked = True
        在庫チェックモードToolStripMenuItem.Checked = False
        Qoo10追跡番号リストモードToolStripMenuItem.Checked = False
        Qoo10発送予定リストモードToolStripMenuItem.Checked = False
        ヤフオク入金リストモードToolStripMenuItem.Checked = False
    End Sub

    Private Sub 在庫チェックモードToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 在庫チェックモードToolStripMenuItem.Click
        ToolStripStatusLabel4.Text = "在庫チェックモード"
        通常モードToolStripMenuItem.Checked = False
        在庫チェックモードToolStripMenuItem.Checked = True
        Qoo10追跡番号リストモードToolStripMenuItem.Checked = False
        Qoo10発送予定リストモードToolStripMenuItem.Checked = False
        ヤフオク入金リストモードToolStripMenuItem.Checked = False
    End Sub

    Private Sub Qoo10追跡番号リストモードToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Qoo10追跡番号リストモードToolStripMenuItem.Click
        ToolStripStatusLabel4.Text = "Qoo10追跡番号リストモード"
        通常モードToolStripMenuItem.Checked = False
        在庫チェックモードToolStripMenuItem.Checked = False
        Qoo10追跡番号リストモードToolStripMenuItem.Checked = True
        Qoo10発送予定リストモードToolStripMenuItem.Checked = False
        ヤフオク入金リストモードToolStripMenuItem.Checked = False
        Qoo10CartNo()
    End Sub

    Private Sub ヤフオク入金リストモードToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ヤフオク入金リストモードToolStripMenuItem.Click
        ToolStripStatusLabel4.Text = "ヤフオクモード"
        通常モードToolStripMenuItem.Checked = False
        在庫チェックモードToolStripMenuItem.Checked = False
        Qoo10追跡番号リストモードToolStripMenuItem.Checked = False
        Qoo10発送予定リストモードToolStripMenuItem.Checked = False
        ヤフオク入金リストモードToolStripMenuItem.Checked = True
    End Sub

    Private Sub Qoo10CartNo()
        'Qoo10からのDeliveryデータが正しいかチェック（要約・詳細どちらでも可）
        If DataGridView1.RowCount > 0 Then
            Dim hCheckCount As Integer = 0
            For c As Integer = 0 To DataGridView1.ColumnCount - 1
                Select Case DataGridView1.Item(c, 0).Value
                    Case "受取人名", "カート番号", "注文番号", "配送会社", "送り状番号", "発送日"
                        hCheckCount += 1
                    Case Else

                End Select
            Next
            If hCheckCount < 6 Then
                MsgBox("Qoo10csvが正常なデータで無い可能性が高いです。正常なデータを読み込んでください。", MsgBoxStyle.SystemModal)
                Exit Sub
            End If
        End If

        'datagridviewのヘッダーテキストをコレクションに取り込む
        Dim DGVheaderCheck As ArrayList = New ArrayList
        For c As Integer = 0 To DataGridView1.Columns.Count - 1
            DGVheaderCheck.Add(DataGridView1.Item(c, 0).Value)
        Next c

        Dim cartNoArray As ArrayList = New ArrayList
        Dim ShipPersonArray As ArrayList = New ArrayList
        For r As Integer = 0 To DataGridView1.RowCount - 1
            Dim searchA As String = Replace(DataGridView1.Item(DGVheaderCheck.IndexOf("受取人名"), r).Value, "　", "")
            searchA = Replace(searchA, " ", "")
            searchA = StrConv(searchA, VbStrConv.Narrow)
            Dim searchB As String = DataGridView1.Item(DGVheaderCheck.IndexOf("カート番号"), r).Value

            For r2 As Integer = r + 1 To DataGridView1.RowCount - 1
                Dim searchAA As String = Replace(DataGridView1.Item(DGVheaderCheck.IndexOf("受取人名"), r2).Value, "　", "")
                searchAA = Replace(searchAA, " ", "")
                searchAA = StrConv(searchAA, VbStrConv.Narrow)
                If searchA = searchAA Then
                    If searchB <> DataGridView1.Item(DGVheaderCheck.IndexOf("カート番号"), r2).Value Then
                        For c As Integer = 0 To DataGridView1.ColumnCount - 1
                            DataGridView1.Item(c, r).Style.BackColor = Color.Lavender
                        Next
                        For c As Integer = 0 To DataGridView1.ColumnCount - 1
                            DataGridView1.Item(c, r2).Style.BackColor = Color.Lavender
                        Next
                    End If
                End If
            Next
        Next
    End Sub

    Private Sub Qoo10発送予定リストモードToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Qoo10発送予定リストモードToolStripMenuItem.Click
        ToolStripStatusLabel4.Text = "Qoo10発送予定リストモード"
        通常モードToolStripMenuItem.Checked = False
        在庫チェックモードToolStripMenuItem.Checked = False
        Qoo10追跡番号リストモードToolStripMenuItem.Checked = False
        Qoo10発送予定リストモードToolStripMenuItem.Checked = True
        ヤフオク入金リストモードToolStripMenuItem.Checked = False
    End Sub

    '*********************************************************************************************************
    'Qoo10追跡番号リストモード
    '*********************************************************************************************************
    Private Sub Qoo10Tuiseki(ByVal filename As String)
        Dim csvRecords As New ArrayList()
        csvRecords = CSV_READ(filename)

        '読み込んだファイルを調べる
        If csvRecords.Count > 0 Then
            Dim sArray0 As String() = Split(csvRecords(0), "|=|")
            Dim sArray1 As String() = Split(csvRecords(1), "|=|")
            For i As Integer = 0 To sArray0.Count - 1
                Select Case sArray0(i)
                    Case "お問い合わせ番号"
                        OtherDialog.qoo10mode = 1
                        Exit For
                    Case "伝票番号"
                        Dim joutai As Integer = 0   '「状態」列の番号を調べる
                        For k As Integer = 0 To sArray0.Count - 1
                            If sArray0(k) = "状態" Then
                                joutai = k
                                Exit For
                            End If
                        Next
                        If sArray1(joutai) = "キャンセル" Then
                            OtherDialog.qoo10mode = 5
                            Exit For
                        End If
                    Case "受注分類タグ"
                        Dim joutai As Integer = 0   '「受注分類タグ」列の番号を調べる
                        For k As Integer = 0 To sArray0.Count - 1
                            If sArray0(k) = "受注分類タグ" Then
                                joutai = k
                                Exit For
                            End If
                        Next
                        If InStr(sArray1(joutai), "予約") > 0 Then
                            OtherDialog.qoo10mode = 6
                            Exit For
                        Else
                            OtherDialog.qoo10mode = 3
                            Exit For
                        End If
                End Select
            Next
        Else
            MsgBox("読み込みファイルのデータに不備があります", MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        '選択画面を表示
        OtherDialog.Size = New Size(230, 240)
        OtherDialog.TabControl1.SelectedIndex = 0
        If OtherDialog.qoo10mode = 1 Then
            OtherDialog.RadioButton3.Checked = True
        ElseIf OtherDialog.qoo10mode = 5 Then
            OtherDialog.RadioButton4.Checked = True
        ElseIf OtherDialog.qoo10mode = 3 Then
            OtherDialog.RadioButton2.Checked = True
        ElseIf OtherDialog.qoo10mode = 6 Then
            OtherDialog.RadioButton6.Checked = True
        End If

        Dim DR As DialogResult = OtherDialog.ShowDialog()
        If DR = Windows.Forms.DialogResult.Cancel Then
            Exit Sub
        End If

        'datagridviewのヘッダーテキストをコレクションに取り込む
        Dim DGVheaderCheck As ArrayList = New ArrayList
        For c As Integer = 0 To DataGridView1.Columns.Count - 1
            'DGVheaderCheck.Add(DataGridView1.Columns(c).HeaderText)
            DGVheaderCheck.Add(DataGridView1.Item(c, 0).Value)
        Next c

        Dim titleNum As Integer() = New Integer(4) {}
        Dim delNum As Integer = 9999
        Dim titleR As String() = Split(csvRecords(0), "|=|")
        If OtherDialog.qoo10mode = 0 Then   '佐川ビズロジ
            titleNum(0) = 49
            titleNum(1) = 18
            titleNum(2) = 1
            titleNum(3) = 99
            titleNum(4) = 99
        ElseIf OtherDialog.qoo10mode = 1 Then   'ゆうパケット
            titleNum(0) = 4
            titleNum(1) = 99
            titleNum(2) = 0
            titleNum(3) = 2
            titleNum(4) = 3
        Else
            titleNum(3) = 99
            titleNum(4) = 99
            For c As Integer = 0 To titleR.Length - 1
                Select Case titleR(c)
                    Case "品名２", "記事", "お客様側管理番号", "受注番号"
                        If titleNum(0) = 0 Then '佐川とヤマトの重複防止（品名２）
                            titleNum(0) = c
                        End If
                    Case "お届け先名称１", "お届け先名", "お届け先　名称", "お届け先　名称２", "送り先名"
                        titleNum(1) = c
                    Case "お問合せ送り状№", "伝票番号", "お問い合わせ番号", "発送伝票番号"
                        titleNum(2) = c
                    Case "発送方法"
                        titleNum(3) = c
                    Case "状態"
                        titleNum(4) = c
                    Case "削除区分"
                        delNum = c
                    Case Else

                End Select
            Next
        End If

        'デフォルトの配送方法
        Dim haisou As String = ""
        If OtherDialog.qoo10mode = 0 Then
            haisou = "佐川急便"
        ElseIf OtherDialog.qoo10mode = 4 Then
            haisou = "ヤマト運輸"
        ElseIf OtherDialog.qoo10mode = 1 Or OtherDialog.qoo10mode = 3 Or OtherDialog.qoo10mode = 5 Then
            haisou = "ゆうパケット"
        Else
            haisou = "ゆうパック"
        End If

        For r1 As Integer = 0 To DataGridView1.RowCount - 2
            Dim start As Integer = 0
            If InStr(csvRecords(0), "お届け先") > 0 Then
                start = 1
            End If

            For r2 As Integer = start To csvRecords.Count - 1
                Dim sArray As String() = Split(csvRecords(r2), "|=|")

                '追跡番号が2つ以上入っている時は最初の1つにしておく
                If InStr(sArray(titleNum(2)), splitter) > 0 Then
                    Dim res As String() = Split(sArray(titleNum(2)), splitter)
                    sArray(titleNum(2)) = res(0)
                End If

                Dim flag As Boolean = False
                If delNum = 9999 Then
                    flag = True
                ElseIf sArray(delNum) <> 1 Then
                    flag = True
                End If

                '0=ビズロジ 1=ゆうパケット 2=ゆうパック 3=e飛伝2 4=B2
                If flag = True Then
                    If OtherDialog.qoo10mode = 4 Then
                        Dim searchA As String = Replace(DataGridView1.Item(DGVheaderCheck.IndexOf("受取人名"), r1).Value, "　", "")
                        searchA = Replace(searchA, " ", "")
                        searchA = StrConv(searchA, VbStrConv.Narrow)
                        Dim searchB As String = Replace(sArray(titleNum(1)), "　", "")
                        searchB = Replace(searchB, " ", "")
                        searchB = StrConv(searchB, VbStrConv.Narrow)

                        '注文番号が入っている時は優先する
                        If sArray(titleNum(0)) <> "" Then
                            If DataGridView1.Item(DGVheaderCheck.IndexOf("注文番号"), r1).Value = sArray(titleNum(0)) Then
                                DataGridView1.Item(DGVheaderCheck.IndexOf("送り状番号"), r1).Value = sArray(titleNum(2))
                                DataGridView1.Item(DGVheaderCheck.IndexOf("配送会社"), r1).Value = haisou
                                Exit For
                            End If
                        Else
                            If searchA = searchB Then
                                DataGridView1.Item(DGVheaderCheck.IndexOf("送り状番号"), r1).Value = sArray(titleNum(2))
                                DataGridView1.Item(DGVheaderCheck.IndexOf("配送会社"), r1).Value = haisou
                                Exit For
                            End If
                        End If
                    ElseIf OtherDialog.qoo10mode = 1 Or OtherDialog.qoo10mode = 2 Then
                        If DataGridView1.Item(DGVheaderCheck.IndexOf("注文番号"), r1).Value = sArray(titleNum(0)) Then
                            DataGridView1.Item(DGVheaderCheck.IndexOf("送り状番号"), r1).Value = sArray(titleNum(2))
                            If sArray(titleNum(3)) = 0 Then
                                haisou = "ゆうパック"
                            Else
                                haisou = "ゆうパケット"
                            End If
                            DataGridView1.Item(DGVheaderCheck.IndexOf("配送会社"), r1).Value = haisou
                            Exit For
                        End If
                    ElseIf OtherDialog.qoo10mode = 3 Then   'ネクストエンジン（追跡番号）
                        If DataGridView1.Item(DGVheaderCheck.IndexOf("注文番号"), r1).Value = sArray(titleNum(0)) Then
                            DataGridView1.Item(DGVheaderCheck.IndexOf("送り状番号"), r1).Value = sArray(titleNum(2))
                            Dim tn3 As String = sArray(titleNum(3))
                            tn3 = Regex.Replace(tn3, "\(.*\)", "")
                            If DataGridView1.Item(DGVheaderCheck.IndexOf("配送会社"), r1).Value <> tn3 Then
                                DataGridView1.Item(DGVheaderCheck.IndexOf("配送会社"), r1).Value = tn3
                            End If
                        End If
                    ElseIf OtherDialog.qoo10mode = 5 Then   'ネクストエンジン（同梱キャンセル）
                        If DataGridView1.Item(DGVheaderCheck.IndexOf("注文番号"), r1).Value = sArray(titleNum(0)) Then
                            '既にデータが入っている時は目印を入れない
                            If DataGridView1.Item(DGVheaderCheck.IndexOf("送り状番号"), r1).Value = "" Then
                                DataGridView1.Item(DGVheaderCheck.IndexOf("送り状番号"), r1).Value = DataGridView1.Item(DGVheaderCheck.IndexOf("送り状番号"), r1).Value & "同梱C"
                                DataGridView1.Item(DGVheaderCheck.IndexOf("送り状番号"), r1).Style.BackColor = Color.LightBlue
                                Dim tn3 As String = sArray(titleNum(3))
                                tn3 = Regex.Replace(tn3, "\(.*\)", "")    'ゆうパック(C)←(C)を消す
                                haisou = tn3
                            Else
                                If Not IsNumeric(DataGridView1.Item(DGVheaderCheck.IndexOf("送り状番号"), r1).Value) Then
                                    DataGridView1.Item(DGVheaderCheck.IndexOf("送り状番号"), r1).Value = DataGridView1.Item(DGVheaderCheck.IndexOf("送り状番号"), r1).Value & "/同梱C"
                                End If
                                DataGridView1.Item(DGVheaderCheck.IndexOf("送り状番号"), r1).Style.BackColor = Color.LightBlue
                            End If

                            Dim cartNo As String = DataGridView1.Item(DGVheaderCheck.IndexOf("カート番号"), r1).Value
                            For r3 As Integer = 0 To DataGridView1.RowCount - 1
                                If DataGridView1.Rows(r3).IsNewRow = False Then
                                    If cartNo = DataGridView1.Item(DGVheaderCheck.IndexOf("カート番号"), r3).Value Then
                                        DataGridView1.Item(DGVheaderCheck.IndexOf("カート番号"), r3).Style.BackColor = Color.LightBlue
                                    End If
                                End If
                            Next
                        End If
                    ElseIf OtherDialog.qoo10mode = 6 Then   'ネクストエンジン（予約）
                        If DataGridView1.Item(DGVheaderCheck.IndexOf("注文番号"), r1).Value = sArray(titleNum(0)) Then
                            '既にデータが入っている時は目印を入れない
                            If DataGridView1.Item(DGVheaderCheck.IndexOf("送り状番号"), r1).Value = "" Then
                                DataGridView1.Item(DGVheaderCheck.IndexOf("送り状番号"), r1).Value = DataGridView1.Item(DGVheaderCheck.IndexOf("送り状番号"), r1).Value & "予約"
                                DataGridView1.Item(DGVheaderCheck.IndexOf("送り状番号"), r1).Style.BackColor = Color.Orange
                                Dim tn3 As String = sArray(titleNum(3))
                                tn3 = Regex.Replace(tn3, "\(.*\)", "")
                                haisou = tn3
                            Else
                                If Not IsNumeric(DataGridView1.Item(DGVheaderCheck.IndexOf("送り状番号"), r1).Value) Then
                                    DataGridView1.Item(DGVheaderCheck.IndexOf("送り状番号"), r1).Value = DataGridView1.Item(DGVheaderCheck.IndexOf("送り状番号"), r1).Value & "/予約"
                                End If
                                DataGridView1.Item(DGVheaderCheck.IndexOf("送り状番号"), r1).Style.BackColor = Color.Orange
                            End If

                            Dim cartNo As String = DataGridView1.Item(DGVheaderCheck.IndexOf("カート番号"), r1).Value
                            For r3 As Integer = 0 To DataGridView1.RowCount - 1
                                If DataGridView1.Rows(r3).IsNewRow = False Then
                                    If cartNo = DataGridView1.Item(DGVheaderCheck.IndexOf("カート番号"), r3).Value Then
                                        DataGridView1.Item(DGVheaderCheck.IndexOf("発送日"), r3).Style.BackColor = Color.Orange
                                    End If
                                End If
                            Next
                        End If
                    ElseIf OtherDialog.qoo10mode = 7 Then
                        If DataGridView1.Item(DGVheaderCheck.IndexOf("注文番号"), r1).Value = sArray(titleNum(0)) Then
                            DataGridView1.Item(DGVheaderCheck.IndexOf("送り状番号"), r1).Value = sArray(titleNum(4))
                            DataGridView1.Item(DGVheaderCheck.IndexOf("送り状番号"), r1).Style.BackColor = Color.Yellow
                        End If
                    Else
                        If DataGridView1.Item(DGVheaderCheck.IndexOf("注文番号"), r1).Value = sArray(titleNum(0)) Then
                            DataGridView1.Item(DGVheaderCheck.IndexOf("送り状番号"), r1).Value = sArray(titleNum(2))
                            DataGridView1.Item(DGVheaderCheck.IndexOf("配送会社"), r1).Value = haisou
                            Exit For
                        End If
                    End If
                End If
            Next
        Next

        Dim maeName As ArrayList = New ArrayList
        Dim maeNo As ArrayList = New ArrayList
        Dim maeHaiso As ArrayList = New ArrayList
        Dim maeCartNo As ArrayList = New ArrayList

        '一度データをリストに入れる
        For r1 As Integer = 0 To DataGridView1.RowCount - 2
            If DataGridView1.Item(DGVheaderCheck.IndexOf("送り状番号"), r1).Value <> "" Then
                If maeNo.Contains(DataGridView1.Item(DGVheaderCheck.IndexOf("送り状番号"), r1).Value) = False Then
                    Dim searchA As String = Replace(DataGridView1.Item(DGVheaderCheck.IndexOf("受取人名"), r1).Value, "　", "")
                    searchA = Replace(searchA, " ", "")
                    searchA = StrConv(searchA, VbStrConv.Narrow)
                    maeName.Add(searchA)
                    maeNo.Add(DataGridView1.Item(DGVheaderCheck.IndexOf("送り状番号"), r1).Value)
                    maeHaiso.Add(DataGridView1.Item(DGVheaderCheck.IndexOf("配送会社"), r1).Value)
                    maeCartNo.Add(DataGridView1.Item(DGVheaderCheck.IndexOf("カート番号"), r1).Value)
                End If
            End If
        Next

        For r1 As Integer = 0 To DataGridView1.RowCount - 2
            If DataGridView1.Item(DGVheaderCheck.IndexOf("送り状番号"), r1).Value = "" Then
                Dim searchA As String = DataGridView1.Item(DGVheaderCheck.IndexOf("カート番号"), r1).Value
                If maeCartNo.Contains(searchA) = True Then
                    Dim no As Integer = maeCartNo.IndexOf(searchA)
                    DataGridView1.Item(DGVheaderCheck.IndexOf("送り状番号"), r1).Value = maeNo(no)
                    DataGridView1.Item(DGVheaderCheck.IndexOf("配送会社"), r1).Value = maeHaiso(no)
                End If
            End If
        Next

        MsgBox("リスト合成終了しました", MsgBoxStyle.SystemModal)
    End Sub

    Private Sub Qoo10Hassou(ByVal filename As String)
        'datagridviewのヘッダーテキストをコレクションに取り込む
        Dim DGVheaderCheck As ArrayList = New ArrayList
        For c As Integer = 0 To DataGridView1.Columns.Count - 1
            'DGVheaderCheck.Add(DataGridView1.Columns(c).HeaderText)
            DGVheaderCheck.Add(DataGridView1.Item(c, 0).Value)
        Next c

        Dim csvRecords As New ArrayList()
        csvRecords = CSV_READ(filename)

        '読み込みcsvの列番号を取得する
        Dim titleR As String() = Split(csvRecords(0), "|=|")
        Dim orderNum2 As Integer = 0
        Dim orderNum3 As Integer = 0
        Dim hassouD As Integer = 0
        Dim joutaiNum As Integer = 0    '同梱キャンセルデータか調べる
        Dim tagNum As Integer = 0       '予約データか調べる
        For c As Integer = 0 To titleR.Length - 1
            Select Case titleR(c)
                Case "受注番号"
                    orderNum2 = c
                Case "発送方法"
                    orderNum3 = c
                Case "状態"
                    joutaiNum = c
                Case "受注分類タグ"
                    tagNum = c
                Case Else

            End Select
        Next

        '「印刷待ち・済み」データか、同梱キャンセルデータか調べる
        Dim dataKind As Integer = 0 '0=通常、1=同梱キャンセル、2=予約
        Dim titleR2 As String() = Split(csvRecords(1), "|=|")
        If InStr(titleR2(joutaiNum), "キャンセル") > 0 Then
            dataKind = 1
        ElseIf InStr(titleR2(tagNum), "予約") > 0 Then
            dataKind = 2
        End If

        Dim cartArray As ArrayList = New ArrayList
        Dim hassouArray As ArrayList = New ArrayList

        For r1 As Integer = 0 To DataGridView1.RowCount - 2
            If InStr(csvRecords(0), "販売者商品コード") = 0 Then
                If cartArray.Contains(DataGridView1.Item(DGVheaderCheck.IndexOf("カート番号"), r1).Value) Then
                    DataGridView1.Item(DGVheaderCheck.IndexOf("発送予定日"), r1).Value = Format(DateAdd(DateInterval.Day, 1, Now()), "yyyy/MM/dd")
                    Dim haisou As String = hassouArray(cartArray.IndexOf(DataGridView1.Item(DGVheaderCheck.IndexOf("カート番号"), r1).Value))
                    haisou = Regex.Replace(haisou, "\(.*\)", "")
                    DataGridView1.Item(DGVheaderCheck.IndexOf("配送会社"), r1).Value = haisou
                    DataGridView1.Item(DGVheaderCheck.IndexOf("発送予定日"), r1).Style.BackColor = Color.Wheat
                Else
                    For r2 As Integer = 0 To csvRecords.Count - 1
                        Dim sArray As String() = Split(csvRecords(r2), "|=|")
                        If DataGridView1.Item(DGVheaderCheck.IndexOf("注文番号"), r1).Value = sArray(orderNum2) Then
                            If DataGridView1.Item(DGVheaderCheck.IndexOf("発送予定日"), r1).Value = "" Then
                                DataGridView1.Item(DGVheaderCheck.IndexOf("発送予定日"), r1).Value = Format(DateAdd(DateInterval.Day, 1, Now()), "yyyy/MM/dd")
                                cartArray.Add(DataGridView1.Item(DGVheaderCheck.IndexOf("カート番号"), r1).Value)
                                Dim haisou As String = Regex.Replace(sArray(orderNum3), "\(.*\)", "")
                                DataGridView1.Item(DGVheaderCheck.IndexOf("配送会社"), r1).Value = haisou
                                hassouArray.Add(haisou)
                            End If

                            If dataKind = 1 Then
                                DataGridView1.Item(DGVheaderCheck.IndexOf("発送予定日"), r1).Style.BackColor = Color.LightBlue
                            ElseIf dataKind = 2 Then
                                DataGridView1.Item(DGVheaderCheck.IndexOf("発送予定日"), r1).Value = "予約"
                                DataGridView1.Item(DGVheaderCheck.IndexOf("発送予定日"), r1).Style.BackColor = Color.Orange
                            End If
                        End If
                    Next
                End If
            End If
        Next

        MsgBox("リスト検索終了しました", MsgBoxStyle.SystemModal)
    End Sub

    ''*********************************************************************************************************
    ''在庫チェック（EC店長）
    ''*********************************************************************************************************
    'Private Sub ZaikoCheck(ByVal filename As String)
    '    DataGridView1.Enabled = False
    '    '-----------------------------------------------------------------------
    '    PreDIV()
    '    '-----------------------------------------------------------------------

    '    Dim textColumn As New DataGridViewTextBoxColumn()
    '    DataGridView1.Columns.Add(textColumn)
    '    Dim lastColumnNum As Integer = DataGridView1.ColumnCount - 1
    '    DataGridView1.Columns(lastColumnNum).HeaderText = lastColumnNum
    '    Dim textColumn2 As New DataGridViewTextBoxColumn()
    '    DataGridView1.Columns.Add(textColumn2)
    '    DataGridView1.Columns(lastColumnNum + 1).HeaderText = lastColumnNum + 1

    '    'datagridviewのヘッダーテキストをコレクションに取り込む
    '    Dim DGVheaderCheck As ArrayList = New ArrayList
    '    For c As Integer = 0 To DataGridView1.Columns.Count - 1
    '        'DGVheaderCheck.Add(DataGridView1.Columns(c).HeaderText)
    '        DGVheaderCheck.Add(DataGridView1.Item(c, 0).Value)
    '    Next c

    '    Dim csvRecords As New ArrayList()
    '    csvRecords = CSV_READ(filename)

    '    'For r As Integer = 0 To DataGridView1.RowCount - 2
    '    '    Dim a As Integer = DataGridView1.RowCount - 2 - r
    '    '    If DataGridView1.Item(0, a).Value Is "" Then
    '    '        DataGridView1.Rows.RemoveAt(a)
    '    '    End If
    '    'Next

    '    ToolStripProgressBar1.Value = 0
    '    ToolStripProgressBar1.Maximum = csvRecords.Count

    '    Dim sNum As Integer() = Nothing
    '    sNum = New Integer() {
    '        DGVheaderCheck.IndexOf("EC店長商品コード"),
    '        DGVheaderCheck.IndexOf("顧客社商品コード"),
    '        DGVheaderCheck.IndexOf("オプションコード"),
    '        DGVheaderCheck.IndexOf("商品名"),
    '        DGVheaderCheck.IndexOf("オプション"),
    '        DGVheaderCheck.IndexOf("在庫数")}

    '    Dim rr As Integer = 0
    '    For z As Integer = 0 To csvRecords.Count - 1
    '        If InStr(csvRecords(z), "|=|") > 0 Then
    '            Dim sArray As String() = Split(csvRecords(z), "|=|")
    '            Dim flag As Boolean = False
    '            For r As Integer = 0 To DataGridView1.RowCount - 2
    '                If DataGridView1.Item(lastColumnNum, r).Value = "" Then
    '                    If sArray(sNum(0)) = DataGridView1.Item(0, r).Value And sArray(sNum(1)) = DataGridView1.Item(1, r).Value Then
    '                        If DataGridView1.Item(2, r).Value <> "" Then
    '                            If sArray(sNum(2)) = DataGridView1.Item(2, r).Value Then
    '                                DataGridView1.Item(lastColumnNum, r).Value = sArray(sNum(5))
    '                                rr = r
    '                                flag = True
    '                                Exit For
    '                            Else

    '                            End If
    '                        Else
    '                            DataGridView1.Item(lastColumnNum, r).Value = sArray(sNum(5))
    '                            rr = r
    '                            flag = True
    '                            Exit For
    '                        End If
    '                    Else

    '                    End If
    '                End If
    '            Next

    '            If flag = False Then    '最初のCSVに無い物はアクア
    '                DataGridView1.Rows.Add(sArray(sNum(0)), sArray(sNum(1)), sArray(sNum(2)), sArray(sNum(3)), sArray(sNum(4)))
    '                DataGridView1.Item(lastColumnNum, DataGridView1.RowCount - 2).Value = sArray(sNum(5))
    '                DataGridView1.Item(lastColumnNum, DataGridView1.RowCount - 2).Style.BackColor = Color.Aqua
    '            Else
    '                If DataGridView1.Item(lastColumnNum, rr).Value <> DataGridView1.Item(lastColumnNum - 1, rr).Value Then
    '                    If DataGridView1.Item(lastColumnNum, rr).Value = 0 And Not DataGridView1.Item(lastColumnNum, rr).Value Is Nothing Then
    '                        DataGridView1.Item(lastColumnNum, rr).Style.BackColor = Color.Red
    '                    Else
    '                        DataGridView1.Item(lastColumnNum, rr).Style.BackColor = Color.Yellow
    '                    End If
    '                    DataGridView1.Item(lastColumnNum + 1, rr).Value = DataGridView1.Item(lastColumnNum, rr).Value - DataGridView1.Item(lastColumnNum - 1, rr).Value
    '                End If
    '            End If

    '            ToolStripProgressBar1.Value += 1
    '            If z Mod 100 = 0 Then
    '                Application.DoEvents()
    '            End If
    '        End If
    '    Next

    '    '-----------------------------------------------------------------------
    '    RetDIV()
    '    '-----------------------------------------------------------------------
    '    Me.Text = filename
    '    ToolStripProgressBar1.Value = 0
    '    DataGridView1.Enabled = True
    'End Sub

    'Private Sub ZaikoCheck2(ByVal filename As String)
    '    DataGridView1.Enabled = False
    '    '-----------------------------------------------------------------------
    '    PreDIV()
    '    '-----------------------------------------------------------------------

    '    'datagridviewのヘッダーテキストをコレクションに取り込む
    '    Dim DGVheaderCheck As ArrayList = New ArrayList
    '    For c As Integer = 0 To DataGridView1.Columns.Count - 1
    '        'DGVheaderCheck.Add(DataGridView1.Columns(c).HeaderText)
    '        DGVheaderCheck.Add(DataGridView1.Item(c, 0).Value)
    '    Next c

    '    Dim sNum As Integer() = Nothing
    '    sNum = New Integer() {
    '        DGVheaderCheck.IndexOf("商品コード"),
    '        DGVheaderCheck.IndexOf("在庫数")}

    '    '列削除する
    '    For c As Integer = 3 To DataGridView1.ColumnCount - 1
    '        DataGridView1.Columns.RemoveAt(DataGridView1.ColumnCount - 1)
    '    Next

    '    Dim textColumn As New DataGridViewTextBoxColumn()
    '    DataGridView1.Columns.Add(textColumn)
    '    Dim lastColumnNum As Integer = DataGridView1.ColumnCount - 1
    '    DataGridView1.Columns(lastColumnNum).HeaderText = lastColumnNum
    '    Dim textColumn2 As New DataGridViewTextBoxColumn()
    '    DataGridView1.Columns.Add(textColumn2)
    '    DataGridView1.Columns(lastColumnNum + 1).HeaderText = lastColumnNum + 1

    '    '1行目を保存
    '    topArray = New String(DataGridView1.ColumnCount) {}
    '    For c As Integer = 0 To DataGridView1.ColumnCount - 1
    '        topArray(c) = DataGridView1.Item(c, 0).Value
    '    Next
    '    DataGridView1.Rows.RemoveAt(0)

    '    'ソートする
    '    direction = System.ComponentModel.ListSortDirection.Ascending
    '    DataGridView1.Sort(DataGridView1.Columns(0), direction)

    '    Dim csvRecords As New ArrayList()
    '    csvRecords = CSV_READ(filename)

    '    ToolStripProgressBar1.Value = 0
    '    ToolStripProgressBar1.Maximum = csvRecords.Count

    '    Dim rr As Integer = 0
    '    For z As Integer = 0 To csvRecords.Count - 1
    '        If InStr(csvRecords(z), "|=|") > 0 Then
    '            Dim sArray As String() = Split(csvRecords(z), "|=|")
    '            Dim flag As Boolean = False
    '            For r As Integer = 0 To DataGridView1.RowCount - 2
    '                If DataGridView1.Item(lastColumnNum, r).Value = "" Then
    '                    If sArray(sNum(0)) = DataGridView1.Item(0, r).Value Then
    '                        DataGridView1.Item(lastColumnNum, r).Value = sArray(sNum(1))
    '                        rr = r
    '                        flag = True
    '                        Exit For
    '                    Else

    '                    End If
    '                End If
    '            Next

    '            If flag = False Then    '最初のCSVに無い物はアクア
    '                DataGridView1.Rows.Add(sArray(sNum(0)), sArray(sNum(1)))
    '                DataGridView1.Item(lastColumnNum, DataGridView1.RowCount - 2).Value = sArray(sNum(1))
    '                DataGridView1.Item(lastColumnNum, DataGridView1.RowCount - 2).Style.BackColor = Color.Aqua
    '            Else
    '                If DataGridView1.Item(lastColumnNum, rr).Value <> DataGridView1.Item(lastColumnNum - 1, rr).Value Then
    '                    If DataGridView1.Item(lastColumnNum, rr).Value = 0 And Not DataGridView1.Item(lastColumnNum, rr).Value Is Nothing Then
    '                        DataGridView1.Item(lastColumnNum, rr).Style.BackColor = Color.Red
    '                    Else
    '                        DataGridView1.Item(lastColumnNum, rr).Style.BackColor = Color.Yellow
    '                    End If
    '                    DataGridView1.Item(lastColumnNum + 1, rr).Value = DataGridView1.Item(lastColumnNum, rr).Value - DataGridView1.Item(lastColumnNum - 1, rr).Value
    '                End If
    '            End If

    '            ToolStripProgressBar1.Value += 1
    '            If z Mod 100 = 0 Then
    '                Application.DoEvents()
    '            End If
    '        End If
    '    Next

    '    '1行目を戻す
    '    DataGridView1.Rows.Insert(0)
    '    For c As Integer = 0 To DataGridView1.ColumnCount - 1
    '        DataGridView1.Item(c, 0).Value = topArray(c)
    '    Next

    '    '-----------------------------------------------------------------------
    '    RetDIV()
    '    '-----------------------------------------------------------------------
    '    Me.Text = filename
    '    ToolStripProgressBar1.Value = 0
    '    DataGridView1.Enabled = True
    'End Sub

    'ヤフオク入金リストモード
    Private Sub YahooAucNyukin(filename As String)
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------

        Dim csvRecords As New ArrayList()
        csvRecords = CSV_READ(filename)
        Dim cH1 As String() = Split(csvRecords(0), "|=|")
        Dim dH1 As ArrayList = TM_HEADER_1ROW_GET(DataGridView1)

        Dim nH1 As String() = New String() {"受注番号", "金額", "承認番号"}
        For c As Integer = 0 To nH1.Length - 1
            Dim cc As Integer = DataGridView1.ColumnCount
            DataGridView1.Columns.Add(cc, cc)
            DataGridView1.Item(cc, 0).Value = nH1(c)
        Next
        dH1 = TM_HEADER_1ROW_GET(DataGridView1)

        For r As Integer = 1 To DataGridView1.RowCount - 1
            If DataGridView1.Rows(r).IsNewRow Then
                Exit For
            End If

            Dim kID As String = DataGridView1.Item(dH1.IndexOf("決済ID"), r).Value
            If kID <> "" Then
                For i As Integer = 0 To csvRecords.Count - 1
                    Dim line As String() = Split(csvRecords(i), "|=|")
                    If kID = line(Array.IndexOf(cH1, "決済ID")) Then
                        DataGridView1.Item(dH1.IndexOf("受注番号"), r).Value = DataGridView1.Item(dH1.IndexOf("商品ID"), r).Value & "-" & line(Array.IndexOf(cH1, "取引番号"))
                        DataGridView1.Item(dH1.IndexOf("金額"), r).Value = DataGridView1.Item(dH1.IndexOf("支払金額"), r).Value
                        DataGridView1.Item(dH1.IndexOf("承認番号"), r).Value = line(Array.IndexOf(cH1, "決済ID"))
                        'Dim nyukinbi As String = Format(CDate(DataGridView1.Item(dH1.IndexOf("取扱日"), r).Value), "yyyy/MM/dd HH:mm:ss")
                        'DataGridView1.Item(dH1.IndexOf("入金日"), r).Value = nyukinbi
                        DataGridView1.Item(dH1.IndexOf("受注番号"), r).Style.BackColor = Color.Yellow
                        Exit For
                    End If
                Next
            End If
        Next
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
        Me.Text = filename
        MsgBox("処理が終了しました")
    End Sub

    Private Sub CSV整形処理ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CSV整形処理ToolStripMenuItem.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------

        Dim nH1 As String() = New String() {"受注番号", "金額", "承認番号"}
        Dim dH1 As ArrayList = TM_HEADER_1ROW_GET(DataGridView1)
        For c As Integer = DataGridView1.ColumnCount - 1 To 0 Step -1
            If Not nH1.Contains(DataGridView1.Item(c, 0).Value) Then
                DataGridView1.Columns.RemoveAt(c)
            End If
        Next

        For r As Integer = DataGridView1.RowCount - 1 To 1 Step -1
            If DataGridView1.Rows(r).IsNewRow Then
                Continue For
            End If

            If CStr(DataGridView1.Item(0, r).Value) = "" Then
                DataGridView1.Rows.RemoveAt(r)
            End If
        Next

        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
        MsgBox("処理が終了しました")
    End Sub


    Private Sub 表示リセットToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 表示リセットToolStripMenuItem.Click
        DataGridView1.Enabled = False
        '-----------------------------------------------------------------------
        DataGridView1.ScrollBars = ScrollBars.None
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        DataGridView1.ScrollBars = ScrollBars.Both
        '-----------------------------------------------------------------------
        ToolStripProgressBar1.Value = 0
        ToolStripProgressBar1.Maximum = DataGridView1.RowCount
        For r As Integer = 0 To DataGridView1.RowCount - 2
            DataGridView1.Rows(r).Visible = True
            ToolStripProgressBar1.Value += 1
            If r Mod 500 = 0 Then
                Application.DoEvents()
            End If
        Next
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        ToolStripProgressBar1.Value = 0
        DataGridView1.Enabled = True

        マイナス在庫ToolStripMenuItem.Checked = False
        変更有りのみToolStripMenuItem.Checked = False
    End Sub

    Private Sub 変更有りのみToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 変更有りのみToolStripMenuItem.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        If 変更有りのみToolStripMenuItem.Checked = True Then
            Exit Sub
        End If

        DataGridView1.Enabled = False
        '-----------------------------------------------------------------------
        DataGridView1.ScrollBars = ScrollBars.None
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        DataGridView1.ScrollBars = ScrollBars.Both
        '-----------------------------------------------------------------------
        ToolStripProgressBar1.Value = 0
        ToolStripProgressBar1.Maximum = DataGridView1.RowCount
        Dim lastColumnNum As Integer = DataGridView1.ColumnCount - 2
        For r As Integer = 1 To DataGridView1.RowCount - 2
            If IsNumeric(DataGridView1.Item(lastColumnNum, r).Value) Then
                If IsNumeric(DataGridView1.Item(lastColumnNum - 1, r).Value) Then
                    If DataGridView1.Item(lastColumnNum, r).Value = DataGridView1.Item(lastColumnNum - 1, r).Value Then
                        DataGridView1.Rows(r).Visible = False
                    Else
                        DataGridView1.Rows(r).Visible = True
                    End If
                Else
                    DataGridView1.Rows(r).Visible = True
                End If
            Else
                DataGridView1.Rows(r).Visible = False
            End If
            ToolStripProgressBar1.Value += 1
            If r Mod 500 = 0 Then
                Application.DoEvents()
            End If
        Next
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        ToolStripProgressBar1.Value = 0
        DataGridView1.Enabled = True

        変更有りのみToolStripMenuItem.Checked = True
        マイナス在庫ToolStripMenuItem.Checked = False
    End Sub

    Private Sub マイナス在庫ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles マイナス在庫ToolStripMenuItem.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If
        If マイナス在庫ToolStripMenuItem.Checked = True Then
            Exit Sub
        End If

        DataGridView1.Enabled = False
        '-----------------------------------------------------------------------
        DataGridView1.ScrollBars = ScrollBars.None
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        DataGridView1.ScrollBars = ScrollBars.Both
        '-----------------------------------------------------------------------
        ToolStripProgressBar1.Value = 0
        ToolStripProgressBar1.Maximum = DataGridView1.RowCount
        Dim lastColumnNum As Integer = DataGridView1.ColumnCount - 2
        For r As Integer = 1 To DataGridView1.RowCount - 2
            If IsNumeric(DataGridView1.Item(lastColumnNum, r).Value) Then
                If DataGridView1.Item(lastColumnNum, r).Value >= 0 Then
                    DataGridView1.Rows(r).Visible = False
                End If
            Else
                DataGridView1.Rows(r).Visible = False
            End If
            ToolStripProgressBar1.Value += 1
            If r Mod 500 = 0 Then
                Application.DoEvents()
            End If
        Next
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        ToolStripProgressBar1.Value = 0
        DataGridView1.Enabled = True

        マイナス在庫ToolStripMenuItem.Checked = True
        変更有りのみToolStripMenuItem.Checked = False
    End Sub

    Private Sub 佐川伝票データ自動削除ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 佐川伝票データ自動削除ToolStripMenuItem.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If

        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        'datagridviewのヘッダーテキストをコレクションに取り込む
        Dim DGVheaderCheck As ArrayList = New ArrayList
        For c As Integer = 0 To DataGridView1.Columns.Count - 1
            DGVheaderCheck.Add(DataGridView1.Item(c, 0).Value)
        Next c

        Dim delNum As Integer = 0
        For r As Integer = DataGridView1.Rows.Count - 1 To 1 Step -1
            If DataGridView1.Item(DGVheaderCheck.IndexOf("削除区分"), r).Value = "1" Then
                DataGridView1.Rows.RemoveAt(r)
                delNum += 1
            End If
        Next

        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------

        MsgBox(delNum & " 行削除し、" & DataGridView1.Rows.Count - 1 & "行残りました。" & vbCrLf & "このまま保存処理を行ないます。", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
        名前を付けて保存ToolStripMenuItem.PerformClick()
    End Sub

    '*********************************************************************************************************
    '伝票出力
    '*********************************************************************************************************
    Private Sub 佐川伝票変換出力ToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 佐川伝票変換出力ToolStripMenuItem.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If

        Dim headText As String() = New String() {"住所録コード", "お届け先電話番号", "お届け先郵便番号", "お届け先住所１", "お届け先住所２", "お届け先住所３", "お届け先名称１", "お届け先名称２", "お客様管理ナンバー", "お客様コード", "部署・担当者", "荷送人電話番号", "ご依頼主電話番号", "ご依頼主郵便番号", "ご依頼主住所１", "ご依頼主住所２", "ご依頼主名称１", "ご依頼主名称２", "荷姿コード", "品名１", "品名２", "品名３", "品名４", "品名５", "出荷個数", "便種（スピードを選択）", "便種（クール便指定）", "配達日", "配達指定時間帯", "配達指定時間（時分）", "代引金額", "消費税", "保険金額", "指定シール１", "指定シール２", "指定シール３", "営業所止めフラグ", "営業店コード", "元着区分"}
        Dim motoHeadText As ArrayList = HeadCollection()
        Dim selRow As ArrayList = SelCollection()

        Dim str As String = ""
        For i As Integer = 0 To headText.Length - 1
            If i = 0 Then
                str = headText(i)
            Else
                str &= splitter & headText(i)
            End If
        Next
        str &= vbCrLf

        For r As Integer = 1 To DataGridView1.RowCount - 2
            If selRow.Contains(r) = True Then
                Dim yubin As String = ""
                Dim ken As String = ""

                For i As Integer = 0 To headText.Length - 1
                    Select Case i
                        Case 0
                            str &= ""
                        Case 1
                            str &= splitter & DataGridView1.Item(motoHeadText.IndexOf("送付先電話番号"), r).Value
                        Case 2
                            If InStr(DataGridView1.Item(motoHeadText.IndexOf("送付先住所"), r).Value, " ") > 0 Then
                                Dim sArray As String() = Split(DataGridView1.Item(motoHeadText.IndexOf("送付先住所"), r).Value, " ")
                                str &= splitter & sArray(0)
                                yubin = sArray(0) & " "
                            Else
                                str &= splitter
                            End If
                        Case 3
                            If InStr(DataGridView1.Item(motoHeadText.IndexOf("送付先住所"), r).Value, " ") > 0 Then
                                Dim sArray As String() = Split(DataGridView1.Item(motoHeadText.IndexOf("送付先住所"), r).Value, " ")
                                str &= splitter & sArray(1)
                                ken = sArray(1) & " "
                            Else
                                str &= splitter
                            End If
                        Case 4
                            Dim res As String = Replace(DataGridView1.Item(motoHeadText.IndexOf("送付先住所"), r).Value, yubin, "")
                            res = Replace(res, ken, "")
                            str &= splitter & res
                        Case 6
                            str &= splitter & DataGridView1.Item(motoHeadText.IndexOf("送付先氏名"), r).Value
                        Case 7
                            str &= splitter & StrConv(DataGridView1.Item(motoHeadText.IndexOf("送付先氏名カナ"), r).Value, VbStrConv.Narrow)
                        Case 8
                            str &= splitter & "137345180009"
                        Case 10
                            str &= splitter & "KuraNavi(AU)"
                        Case 11
                            str &= splitter & "0925866859"
                        Case 18
                            str &= splitter & "バッグ類"
                        Case 19
                            str &= splitter & "雑貨"
                        Case 20
                            str &= splitter & DataGridView1.Item(motoHeadText.IndexOf("取引ナンバー"), r).Value
                        Case 24
                            str &= splitter & "1"
                        Case 28
                            Dim res As String = DataGridView1.Item(motoHeadText.IndexOf("取引オプション"), r).Value
                            If res <> "" Then
                                Select Case res
                                    Case "配達時間帯=午前中"
                                        str &= splitter & "01"
                                    Case "配達時間帯=12:00～14:00"
                                        str &= splitter & "12"
                                    Case "配達時間帯=14:00～16:00"
                                        str &= splitter & "14"
                                    Case "配達時間帯=16:00～18:00"
                                        str &= splitter & "16"
                                    Case "配達時間帯=18:00～21:00"
                                        str &= splitter & "04"
                                    Case Else
                                        str &= splitter
                                End Select
                            Else
                                str &= splitter
                            End If
                        Case 30
                            If DataGridView1.Item(motoHeadText.IndexOf("購入者希望取引方法"), r).Value = "代金引換便" Then
                                str &= splitter & DataGridView1.Item(motoHeadText.IndexOf("請求金額"), r).Value
                            Else
                                str &= splitter
                            End If
                        Case 31
                            If DataGridView1.Item(motoHeadText.IndexOf("購入者希望取引方法"), r).Value = "代金引換便" Then
                                str &= splitter & "0"
                            Else
                                str &= splitter
                            End If
                        Case Else
                            str &= splitter
                    End Select
                Next
                str &= vbCrLf
            End If
        Next

        SaveDenpyo("佐川", str)
    End Sub

    Private Sub ヤマト伝票変換出力ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ヤマト伝票変換出力ToolStripMenuItem.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If

        Dim headText As String() = New String() {"お客様管理番号", "送り状種別", "温度区分", "伝票番号", "出荷予定日", "配達指定日", "配達時間帯区分", "届け先コード", "届け先電話番号", "届け先電話番号(枝番)", "届け先郵便番号", "お届け先住所", "お届け先住所（アパートマンション名）", "会社・部門名１", "会社・部門名２", "届け先名(漢字)", "届け先名(カナ)", "敬称", "依頼主コード", "依頼主電話番号", "依頼主電話番号(枝番)", "依頼主郵便番号", "ご依頼主住所", "ご依頼主住所（アパートマンション名）", "依頼主名（漢字）", "依頼主名カナ", "品名コード１", "品名１", "品名コード２", "品名２", "荷扱い１", "荷扱い２", "記事", "コレクト代金引換額(税込)", "内消費税額等", "営業所止め置き", "止め置き営業所コード", "発行枚数", "個数口枠の印字", "請求先顧客コード", "請求先分類コード", "運賃管理番号"}
        Dim motoHeadText As ArrayList = HeadCollection()
        Dim selRow As ArrayList = SelCollection()

        Dim str As String = ""
        For i As Integer = 0 To headText.Length - 1
            If i = 0 Then
                str = headText(i)
            Else
                str &= splitter & headText(i)
            End If
        Next
        str &= vbCrLf

        For r As Integer = 1 To DataGridView1.RowCount - 2
            If selRow.Contains(r) = True Then
                Dim yubin As String = ""
                Dim ken As String = ""

                For i As Integer = 0 To headText.Length - 1
                    Select Case i
                        Case 0
                            str &= ""
                        Case 1
                            str &= splitter & "0"
                        Case 4
                            str &= splitter & Format(Now, "yyyy/MM/dd")
                        Case 6
                            Dim res As String = DataGridView1.Item(motoHeadText.IndexOf("取引オプション"), r).Value
                            If res <> "" Then
                                Select Case res
                                    Case "配達時間帯=午前中"
                                        str &= splitter & "0812"
                                    Case "配達時間帯=12:00～14:00"
                                        str &= splitter & "1214"
                                    Case "配達時間帯=14:00～16:00"
                                        str &= splitter & "1416"
                                    Case "配達時間帯=16:00～18:00"
                                        str &= splitter & "1618"
                                    Case "配達時間帯=18:00～21:00"
                                        str &= splitter & "1821"
                                    Case Else
                                        str &= splitter
                                End Select
                            Else
                                str &= splitter
                            End If
                        Case 8
                            str &= splitter & DataGridView1.Item(motoHeadText.IndexOf("送付先電話番号"), r).Value
                        Case 10
                            If InStr(DataGridView1.Item(motoHeadText.IndexOf("送付先住所"), r).Value, " ") > 0 Then
                                Dim sArray As String() = Split(DataGridView1.Item(motoHeadText.IndexOf("送付先住所"), r).Value, " ")
                                str &= splitter & sArray(0)
                                yubin = sArray(0) & " "
                            Else
                                str &= splitter
                            End If
                        Case 11
                            If InStr(DataGridView1.Item(motoHeadText.IndexOf("送付先住所"), r).Value, " ") > 0 Then
                                Dim sArray As String() = Split(DataGridView1.Item(motoHeadText.IndexOf("送付先住所"), r).Value, " ")
                                str &= splitter & sArray(1)
                                ken = sArray(1) & " "
                            Else
                                str &= splitter
                            End If
                        Case 12
                            Dim res As String = Replace(DataGridView1.Item(motoHeadText.IndexOf("送付先住所"), r).Value, yubin, "")
                            res = Replace(res, ken, "")
                            str &= splitter & res
                        Case 15
                            str &= splitter & DataGridView1.Item(motoHeadText.IndexOf("送付先氏名"), r).Value
                        Case 16
                            str &= splitter & StrConv(DataGridView1.Item(motoHeadText.IndexOf("送付先氏名カナ"), r).Value, VbStrConv.Narrow)
                        Case 19
                            str &= splitter & "092-586-6859"
                        Case 21
                            str &= splitter & "816-0921"
                        Case 22
                            str &= splitter & "福岡県大野城市"
                        Case 23
                            str &= splitter & "仲畑1-37-17"
                        Case 24
                            str &= splitter & "KuraNavi(AU)"
                        Case 27
                            str &= splitter & "キャンプ・ボディ用品"
                        Case 32
                            str &= splitter & DataGridView1.Item(motoHeadText.IndexOf("取引ナンバー"), r).Value
                        Case 33
                            If DataGridView1.Item(motoHeadText.IndexOf("購入者希望取引方法"), r).Value = "代金引換便" Then
                                str &= splitter & DataGridView1.Item(motoHeadText.IndexOf("請求金額"), r).Value
                            Else
                                str &= splitter
                            End If
                        Case 34
                            If DataGridView1.Item(motoHeadText.IndexOf("購入者希望取引方法"), r).Value = "代金引換便" Then
                                str &= splitter & "0"
                            Else
                                str &= splitter
                            End If
                        Case 38
                            str &= splitter & "1"
                        Case 39
                            str &= splitter & "929801214"
                        Case 41
                            str &= splitter & "1"
                        Case Else
                            str &= splitter
                    End Select
                Next
                str &= vbCrLf
            End If
        Next

        SaveDenpyo("ヤマト", str)
    End Sub

    Private Sub ゆうパケット伝票変換出力ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ゆうパケット伝票変換出力ToolStripMenuItem.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If

        Dim headText As String() = New String() {"お届け先　郵便番号", "お届け先　住所1", "お届け先　名称", "お届け先　名称２", "お届け先　電話番号", "ご依頼主　郵便番号", "ご依頼主　住所１", "ご依頼主　名称１", "ご依頼主　電話番号", "商品サイズ/厚さ区分", "お届け先　メールアドレス１"}
        Dim motoHeadText As ArrayList = HeadCollection()
        Dim selRow As ArrayList = SelCollection()

        Dim str As String = ""
        For i As Integer = 0 To headText.Length - 1
            If i = 0 Then
                str = headText(i)
            Else
                str &= splitter & headText(i)
            End If
        Next
        str &= vbCrLf

        For r As Integer = 1 To DataGridView1.RowCount - 2
            If selRow.Contains(r) = True Then
                Dim yubin As String = ""
                Dim ken As String = ""

                For i As Integer = 0 To headText.Length - 1
                    Select Case i
                        Case 0
                            If InStr(DataGridView1.Item(motoHeadText.IndexOf("送付先住所"), r).Value, " ") > 0 Then
                                Dim sArray As String() = Split(DataGridView1.Item(motoHeadText.IndexOf("送付先住所"), r).Value, " ")
                                str &= sArray(0)
                                yubin = sArray(0) & " "
                            Else
                                str &= splitter
                            End If
                        Case 1
                            If InStr(DataGridView1.Item(motoHeadText.IndexOf("送付先住所"), r).Value, " ") > 0 Then
                                Dim sArray As String() = Split(DataGridView1.Item(motoHeadText.IndexOf("送付先住所"), r).Value, " ")
                                ken = sArray(1)
                                Dim res As String = Replace(DataGridView1.Item(motoHeadText.IndexOf("送付先住所"), r).Value, yubin, "")
                                res = Replace(res, " ", "")
                                str &= splitter & res
                            Else
                                str &= splitter
                            End If
                        Case 2
                            str &= splitter & DataGridView1.Item(motoHeadText.IndexOf("送付先氏名"), r).Value
                        Case 4
                            str &= splitter & DataGridView1.Item(motoHeadText.IndexOf("送付先電話番号"), r).Value
                        Case 5
                            str &= splitter & "816-0921"
                        Case 6
                            str &= splitter & "福岡県大野城市仲畑1-37-17"
                        Case 7
                            str &= splitter & "KuraNavi(AU)"
                        Case 8
                            str &= splitter & "092-586-6859"
                        Case 10
                            str &= splitter & DataGridView1.Item(motoHeadText.IndexOf("購入者メールアドレス"), r).Value
                        Case Else
                            str &= splitter
                    End Select
                Next
                str &= vbCrLf
            End If
        Next

        SaveDenpyo("ゆうパケット", str)
    End Sub

    Private Sub SaveDenpyo(ByVal fn As String, ByVal str As String)
        Dim sfd As New SaveFileDialog With {
            .FileName = fn & "伝票用出力データ.csv",
            .InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
            .Filter = "CSVファイル(*.csv)|*.csv|すべてのファイル(*.*)|*.*",
                    .AutoUpgradeEnabled = False,
            .FilterIndex = 0,
            .Title = "保存先のファイルを選択してください",
            .RestoreDirectory = True,
            .OverwritePrompt = True,
            .CheckPathExists = True
        }

        'ダイアログを表示する
        If sfd.ShowDialog(Me) = DialogResult.OK Then
            Dim sw As IO.StreamWriter = New IO.StreamWriter(sfd.FileName, False, System.Text.Encoding.GetEncoding("SHIFT-JIS"))
            sw.Write(str)
            sw.Close()
            MsgBox(sfd.FileName & vbCrLf & "保存しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    Private Function HeadCollection()
        Dim motoHeadText As ArrayList = New ArrayList
        For c As Integer = 0 To DataGridView1.ColumnCount - 1
            motoHeadText.Add(DataGridView1.Item(c, 0).Value)
        Next
        Return motoHeadText
    End Function

    Private Function SelCollection()
        Dim selRow As ArrayList = New ArrayList
        For i As Integer = 0 To DataGridView1.SelectedCells.Count - 1
            If selRow.Contains(DataGridView1.SelectedCells(i).RowIndex) = 0 Then
                selRow.Add(DataGridView1.SelectedCells(i).RowIndex)
            End If
        Next
        selRow.Sort()
        Return selRow
    End Function

    Public selChangeFlag As Boolean = True
    Dim BeforeColor
    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        If selChangeFlag = False Then
            Exit Sub
        End If

        If DataGridView1.SelectedCells.Count > 50 Then
            Exit Sub
        End If

        selChangeFlag = False
        '選択行が表示されていない時は選択を解除
        Dim selCell As New ArrayList
        For i As Integer = 0 To DataGridView1.SelectedCells.Count - 1
            If DataGridView1.Rows(DataGridView1.SelectedCells(i).RowIndex).Visible = False Then
                selCell.Add(DataGridView1.SelectedCells(i).ColumnIndex & "," & DataGridView1.SelectedCells(i).RowIndex)
            End If
        Next
        For i As Integer = 0 To selCell.Count - 1
            Dim index As String() = Split(selCell(i), ",")
            DataGridView1.Item(CInt(index(0)), CInt(index(1))).Selected = False
        Next
        selChangeFlag = True

        If DataGridView1.SelectedCells.Count = 1 And Panel2.Visible = False Then
            AzukiControl1.Text = DataGridView1.SelectedCells(0).Value
            If CheckBox2.Checked = True Then
                Dim GyouBegin As Integer = AzukiControl1.GetLineHeadIndex(AzukiControl1.LineCount - 1)
                Dim GyouEnd As Integer = AzukiControl1.GetLineLength(AzukiControl1.LineCount - 1)
                AzukiControl1.Document.SetSelection(GyouBegin, GyouBegin)
                AzukiControl1.ScrollToCaret()
            End If
            DataGridView1.Focus()
            If DataGridView1.SelectedCells(0).Value <> "" Then
                ToolStripStatusLabel7.Text = "[" & DataGridView1.SelectedCells(0).Value.length & "]"    '文字数カウント
            End If
        End If

        '選択行の背景色を変更する
        Dim dgv As DataGridView = DataGridView1
        If dgv.SelectedCells.Count > 50 Then
            Exit Sub
        End If

        For r As Integer = 0 To dgv.RowCount - 1
            If DataGridView1.Rows(r).DefaultCellStyle.BackColor <> BeforeColor Then
                DataGridView1.Rows(r).DefaultCellStyle.BackColor = BeforeColor
            End If
        Next

        Dim selRow As ArrayList = New ArrayList
        For i As Integer = 0 To dgv.SelectedCells.Count - 1
            If Not selRow.Contains(dgv.SelectedCells(i).RowIndex) Then
                selRow.Add(dgv.SelectedCells(i).RowIndex)
            End If
        Next

        For i As Integer = 0 To selRow.Count - 1
            BeforeColor = DataGridView1.Rows(selRow(i)).DefaultCellStyle.BackColor
            DataGridView1.Rows(selRow(i)).DefaultCellStyle.BackColor = Color.LightCyan
        Next
    End Sub

    Private Sub AzukiControl1_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AzukiControl1.SizeChanged
        AzukiControl1.ViewWidth = AzukiControl1.Width - 30
    End Sub

    Private Sub AzukiControl1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AzukiControl1.TextChanged
        If DataGridView1.GridColor = Color.LightGray Then
            If DataGridView1.SelectedCells.Count > 0 Then
                DataGridView1.SelectedCells(0).Value = AzukiControl1.Text
            End If
        End If
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If ToolStripStatusLabel2.Text <> DataGridView1.RowCount - 2 Then
            ToolStripStatusLabel2.Text = DataGridView1.RowCount - 2
        End If
        If ToolStripStatusLabel1.Text <> DataGridView1.SelectedCells.Count Then
            ToolStripStatusLabel1.Text = DataGridView1.SelectedCells.Count
        End If
    End Sub

    'Private Sub 列のソート不可ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 列のソート不可ToolStripMenuItem.Click
    '    If 列のソート不可ToolStripMenuItem.Checked = False Then
    '        For Each c As DataGridViewColumn In DataGridView1.Columns
    '            c.SortMode = DataGridViewColumnSortMode.NotSortable
    '        Next c
    '        列のソート不可ToolStripMenuItem.Checked = True
    '    Else
    '        For Each c As DataGridViewColumn In DataGridView1.Columns
    '            c.SortMode = DataGridViewColumnSortMode.Programmatic
    '        Next c
    '        列のソート不可ToolStripMenuItem.Checked = False
    '    End If
    'End Sub

    'datagridviewのソート
    Dim topArray As String() = Nothing
    Dim direction As System.ComponentModel.ListSortDirection = System.ComponentModel.ListSortDirection.Descending
    Private Sub DataGridView1_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.ColumnHeaderMouseClick
        Dim DR As DialogResult = MsgBox(e.ColumnIndex & "列目でソートします。" & vbCrLf & "1行目除外する。", MsgBoxStyle.YesNoCancel Or MsgBoxStyle.SystemModal)
        If DR = Windows.Forms.DialogResult.Yes Then
            topArray = New String(DataGridView1.ColumnCount) {}
            For c As Integer = 0 To DataGridView1.ColumnCount - 1
                topArray(c) = DataGridView1.Item(c, 0).Value
            Next
            DataGridView1.Rows.RemoveAt(0)
        ElseIf DR = Windows.Forms.DialogResult.Cancel Then
            Exit Sub
        End If

        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        If direction = System.ComponentModel.ListSortDirection.Descending Then
            direction = System.ComponentModel.ListSortDirection.Ascending
            DataGridView1.Sort(DataGridView1.Columns(e.ColumnIndex), direction)
            ToolStripStatusLabel6.Text = "[sort/" & e.ColumnIndex & ":昇順]"
        Else
            direction = System.ComponentModel.ListSortDirection.Descending
            DataGridView1.Sort(DataGridView1.Columns(e.ColumnIndex), direction)
            ToolStripStatusLabel6.Text = "[sort/" & e.ColumnIndex & ":降順]"
        End If

        If DR = Windows.Forms.DialogResult.Yes Then
            DataGridView1.Rows.Insert(0)
            For c As Integer = 0 To DataGridView1.ColumnCount - 1
                DataGridView1.Item(c, 0).Value = topArray(c)
            Next
        End If
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 古い順通常ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 古い順通常ToolStripMenuItem.Click
        ChumonSort(0)
    End Sub

    Private Sub 新しい順ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 新しい順ToolStripMenuItem.Click
        ChumonSort(1)
    End Sub

    Private Sub ChumonSort(ByVal mode As Integer)
        Dim sortColumnName As String() = {"フリー項目２", "品名１", "記事欄06"}
        Dim dH1 As ArrayList = TM_HEADER_1ROW_GET(DataGridView1, 0)

        Dim flag1 As Boolean = False
        Dim sortHeaderStr As String = ""
        For i As Integer = 0 To sortColumnName.Length - 1
            If dH1.Contains(sortColumnName(i)) Then
                flag1 = True
                sortHeaderStr = sortColumnName(i)
                Exit For
            End If
        Next

        If Not flag1 Then
            Exit Sub
        End If

        Dim DR As DialogResult = MsgBox("伝票データを認識しました。" & vbCrLf & "ロケーション順ソートする場合は「OK」" & vbCrLf & "ソートしない場合は「キャンセル」", MsgBoxStyle.OkCancel Or MsgBoxStyle.SystemModal)
        If DR = DialogResult.OK Then
            topArray = New String(DataGridView1.ColumnCount) {}
            For c As Integer = 0 To DataGridView1.ColumnCount - 1
                topArray(c) = DataGridView1.Item(c, 0).Value
            Next
            DataGridView1.Rows.RemoveAt(0)
        ElseIf DR = Windows.Forms.DialogResult.Cancel Then
            Exit Sub
        End If

        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        direction = System.ComponentModel.ListSortDirection.Ascending
        DataGridView1.Sort(DataGridView1.Columns(dH1.IndexOf(sortHeaderStr)), direction)
        ToolStripStatusLabel6.Text = "[sort/" & dH1.IndexOf(sortHeaderStr) & ":昇順]"

        direction = System.ComponentModel.ListSortDirection.Descending
        DataGridView1.Sort(DataGridView1.Columns(dH1.IndexOf(sortHeaderStr)), direction)
        ToolStripStatusLabel6.Text = "[sort/" & dH1.IndexOf(sortHeaderStr) & ":降順]"

        If mode = 0 Then
            direction = System.ComponentModel.ListSortDirection.Ascending
            DataGridView1.Sort(DataGridView1.Columns(dH1.IndexOf(sortHeaderStr)), direction)
            ToolStripStatusLabel6.Text = "[sort/" & dH1.IndexOf(sortHeaderStr) & ":昇順]"
        End If

        If DR = Windows.Forms.DialogResult.OK Then
            DataGridView1.Rows.Insert(0)
            For c As Integer = 0 To DataGridView1.ColumnCount - 1
                DataGridView1.Item(c, 0).Value = topArray(c)
            Next
        End If
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------

        If 並べ替え後自動上書き保存ToolStripMenuItem.Checked = True Then
            ChumonSave()
        End If
    End Sub

    Private Sub ChumonSave()
        Dim DR As DialogResult = MsgBox("引き続き名前を付けて保存しますか？", MsgBoxStyle.YesNo Or MsgBoxStyle.SystemModal)
        If DR = DialogResult.Yes Then
            名前を付けて保存ToolStripMenuItem.PerformClick()
        End If
    End Sub

    Private Sub 表示フォントを変更ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 表示フォントを変更ToolStripMenuItem.Click
        Dim fd As New FontDialog()
        If fd.ShowDialog() <> DialogResult.Cancel Then
            DataGridView1.DefaultCellStyle.Font = fd.Font
        End If
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim num As Integer = 0
        For i As Integer = 0 To DataGridView1.SelectedCells.Count - 1
            If (CheckBox1.Checked = True And DataGridView1.SelectedCells(i).RowIndex <> 0) Or (CheckBox1.Checked = False) Then
                If DataGridView1.Rows(DataGridView1.SelectedCells(i).RowIndex).Visible = True Then
                    DataGridView1.SelectedCells(i).Value = TextBox2.Text & DataGridView1.SelectedCells(i).Value
                    num += 1
                End If
            End If
        Next
        MsgBox(num & "件追加しました", MsgBoxStyle.SystemModal)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim num As Integer = 0
        For i As Integer = 0 To DataGridView1.SelectedCells.Count - 1
            If (CheckBox1.Checked = True And DataGridView1.SelectedCells(i).RowIndex <> 0) Or (CheckBox1.Checked = False) Then
                If DataGridView1.Rows(DataGridView1.SelectedCells(i).RowIndex).Visible = True Then
                    DataGridView1.SelectedCells(i).Value = DataGridView1.SelectedCells(i).Value & TextBox2.Text
                    num += 1
                End If
            End If
        Next
        MsgBox(num & "件追加しました", MsgBoxStyle.SystemModal)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim num As Integer = 0
        For i As Integer = 0 To DataGridView1.SelectedCells.Count - 1
            If (CheckBox1.Checked = True And DataGridView1.SelectedCells(i).RowIndex <> 0) Or (CheckBox1.Checked = False) Then
                If DataGridView1.Rows(DataGridView1.SelectedCells(i).RowIndex).Visible = True Then
                    If InStr(DataGridView1.SelectedCells(i).Value, TextBox2.Text) > 0 Then
                        DataGridView1.SelectedCells(i).Value = Replace(DataGridView1.SelectedCells(i).Value, TextBox2.Text, "")
                        num += 1
                    End If
                End If
            End If
        Next
        MsgBox(num & "件削除しました", MsgBoxStyle.SystemModal)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub DataGridView1_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        If cellChange = True Then
            If e.RowIndex <> DataGridView1.RowCount - 1 Then
                DataGridView1.Item(e.ColumnIndex, e.RowIndex).Style.BackColor = Color.Yellow
            End If
        End If
    End Sub

    Private Sub 行目を残すToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 行目を残すToolStripMenuItem.Click
        ChangeRowSel(1)
    End Sub

    Private Sub 行目を残さないToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 行目を残さないToolStripMenuItem.Click
        ChangeRowSel(0)
    End Sub

    Private Sub ChangeRowSel(ByVal mode As Integer)
        If DataGridView1.RowCount <= 1 Then
            Exit Sub
        End If

        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------

        Dim roC As Integer = DataGridView1.RowCount - 2
        Dim rowEnd As Integer = roC - mode
        For r As Integer = 0 To rowEnd
            Dim flag As Boolean = True
            Dim backR As Integer = roC - r
            For c As Integer = 0 To DataGridView1.ColumnCount - 1
                If DataGridView1.Item(c, backR).Style.BackColor = Color.Yellow Then
                    flag = False
                End If
            Next
            If flag = True Then
                DataGridView1.Rows.RemoveAt(backR)
            End If
        Next

        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 抽出ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 抽出ToolStripMenuItem.Click
        If DataGridView1.RowCount > 1 Then
            HTMLdialog.Size = New Size(308, 451)
            HTMLdialog.TabControl2.SelectTab("TabPage8")
            HTMLdialog.ComboBox10.SelectedItem = "1:CSV1"

            HTMLdialog.ComboBox6.Items.Clear()
            HTMLdialog.ComboBox7.Items.Clear()
            HTMLdialog.ComboBox8.Items.Clear()
            For c As Integer = 0 To DataGridView1.ColumnCount - 1
                HTMLdialog.ComboBox6.Items.Add(DataGridView1.Columns(c).HeaderText)
                HTMLdialog.ComboBox7.Items.Add(DataGridView1.Columns(c).HeaderText)
                HTMLdialog.ComboBox8.Items.Add(DataGridView1.Columns(c).HeaderText)
            Next
            HTMLdialog.ComboBox6.SelectedIndex = 0
            HTMLdialog.ComboBox7.SelectedIndex = 0
            HTMLdialog.ComboBox8.SelectedIndex = 0

            HTMLdialog.Show()
        End If
    End Sub

    Private Sub 抽出リストリセットToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 抽出リストリセットToolStripMenuItem.Click
        If DataGridView1.RowCount > 1 Then
            For r As Integer = 0 To DataGridView1.RowCount - 2
                DataGridView1.Rows(r).Visible = True
            Next
        End If
    End Sub

    Private Sub DataGridView1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.GotFocus
        DataGridView1.GridColor = Color.DarkCyan
        AzukiControl1.BackColor = Color.LightCyan
        TextBox2.BackColor = Color.LightCyan
    End Sub

    Private Sub AzukiControl1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles AzukiControl1.GotFocus
        DataGridView1.GridColor = Color.LightGray
        AzukiControl1.BackColor = Color.White
        AzukiControl1.Refresh()
        TextBox2.BackColor = Color.LightCyan
    End Sub

    Private Sub TextBox2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.GotFocus
        DataGridView1.GridColor = Color.LightGray
        AzukiControl1.BackColor = Color.LightCyan
        TextBox2.BackColor = Color.White
    End Sub

    Private Sub 英数全角半角変換ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 英数全角半角変換ToolStripMenuItem.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        ToolStripProgressBar1.Maximum = DataGridView1.SelectedCells.Count
        ToolStripProgressBar1.Value = 0
        For i As Integer = 0 To DataGridView1.SelectedCells.Count - 1
            DataGridView1.SelectedCells(i).Value = Form1.StrConvNumeric(DataGridView1.SelectedCells(i).Value)
            DataGridView1.SelectedCells(i).Value = Form1.StrConvEnglish(DataGridView1.SelectedCells(i).Value)
            ToolStripProgressBar1.Value += 1
            Application.DoEvents()
        Next
        ToolStripProgressBar1.Value = 0
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
        MsgBox("変換終了", MsgBoxStyle.SystemModal)
    End Sub

    Private Sub スペース半角変換ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles スペース半角変換ToolStripMenuItem.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        ToolStripProgressBar1.Maximum = DataGridView1.SelectedCells.Count
        ToolStripProgressBar1.Value = 0
        For i As Integer = 0 To DataGridView1.SelectedCells.Count - 1
            DataGridView1.SelectedCells(i).Value = Form1.StrConvZSpaceToH(DataGridView1.SelectedCells(i).Value)
            ToolStripProgressBar1.Value += 1
            Application.DoEvents()
        Next
        ToolStripProgressBar1.Value = 0
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
        MsgBox("変換終了", MsgBoxStyle.SystemModal)
    End Sub

    Private Sub 大文字小文字変換選択セルのみToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 大文字小文字変換選択セルのみToolStripMenuItem.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        ToolStripProgressBar1.Maximum = DataGridView1.SelectedCells.Count
        ToolStripProgressBar1.Value = 0
        For i As Integer = 0 To DataGridView1.SelectedCells.Count - 1
            DataGridView1.SelectedCells(i).Value = Form1.StrConvToNarrow(DataGridView1.SelectedCells(i).Value)
        Next
        ToolStripProgressBar1.Value = 0
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
        MsgBox("変換終了", MsgBoxStyle.SystemModal)
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        If DataGridView1.RowCount < 1 Then
            Exit Sub
        ElseIf DataGridView1.Rows(0).IsNewRow Then
            Exit Sub
        End If

        If Csv_change.Visible = False Then
            Csv_change.Show()
        End If

        If Csv_change.DataGridView1.RowCount > 0 Then
            Csv_change.DataGridView1.Rows.Clear()
            Csv_change.DataGridView1.Columns.Clear()
        End If

        '-----------------------------------------------------------------------
        Csv_change.PreDIV()
        '-----------------------------------------------------------------------
        For c As Integer = 0 To DataGridView1.ColumnCount - 1
            Csv_change.DataGridView1.Columns.Add("column" & c + 1, c)
        Next

        Dim newR As Integer = 0
        For r As Integer = 0 To DataGridView1.RowCount - 1
            If DataGridView1.Rows(r).IsNewRow Then
                Continue For
            End If
            If DataGridView1.Rows(r).Visible = True Then
                Csv_change.DataGridView1.Rows.Add(1)
                For c As Integer = 0 To DataGridView1.ColumnCount - 1
                    Csv_change.DataGridView1.Item(c, newR).Value = DataGridView1.Item(c, r).Value
                Next
                newR += 1
            End If
        Next
        '-----------------------------------------------------------------------
        Csv_change.RetDIV()
        '-----------------------------------------------------------------------

        MsgBox("移行完了", MsgBoxStyle.SystemModal)
    End Sub

    Private Sub 区切り処理ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 区切り処理ToolStripMenuItem.Click
        Search.Size = New Size(187, 215)
        'Search.TabControl1.SelectedIndex = 1
        Form1.SetTabVisible(Search.TabControl1, "区切り処理")
        Search.Show()
    End Sub

    Private Sub 佐川急便ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 佐川急便ToolStripMenuItem.Click
        OtherDialog.Size = New Size(199, 199)
        OtherDialog.TabControl1.SelectedIndex = 1
        OtherDialog.Show()
    End Sub

    Private Sub 注文順自動並べ替えToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 注文順自動並べ替えToolStripMenuItem.Click
        If 注文順自動並べ替えToolStripMenuItem.Checked = True Then
            注文順自動並べ替えToolStripMenuItem.Checked = False
            '*********************************************
            Form1.inif.WriteINI(secValue, "csvSortMode", False)
            '*********************************************
            'My.Settings.csvSortMode = False
        Else
            注文順自動並べ替えToolStripMenuItem.Checked = True
            '*********************************************
            Form1.inif.WriteINI(secValue, "csvSortMode", True)
            '*********************************************
            'My.Settings.csvSortMode = True
        End If
    End Sub

    Private Sub 並べ替え後自動上書き保存ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 並べ替え後自動上書き保存ToolStripMenuItem.Click
        If 並べ替え後自動上書き保存ToolStripMenuItem.Checked = True Then
            並べ替え後自動上書き保存ToolStripMenuItem.Checked = False
            '*********************************************
            Form1.inif.WriteINI(secValue, "csvSortAfterSave", False)
            '*********************************************
        Else
            並べ替え後自動上書き保存ToolStripMenuItem.Checked = True
            '*********************************************
            Form1.inif.WriteINI(secValue, "csvSortAfterSave", True)
            '*********************************************
        End If
    End Sub

    Private Sub ファイル一覧取得ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ファイル一覧取得ToolStripMenuItem.Click
        Dim fbd As New FolderBrowserDialog With {
            .Description = "フォルダを指定してください。",
            .RootFolder = Environment.SpecialFolder.Desktop,
            .SelectedPath = "",
            .ShowNewFolderButton = True
        }

        If fbd.ShowDialog(Me) = DialogResult.OK Then
            If DataGridView1.RowCount > 0 Then
                DataGridView1.Rows.Clear()
                DataGridView1.Columns.Clear()
            End If
            DataGridView1.Columns.Add(0, 0)

            Dim di As New System.IO.DirectoryInfo(fbd.SelectedPath)
            Dim files As System.IO.FileInfo() = di.GetFiles("*.*", System.IO.SearchOption.AllDirectories)
            For Each f As System.IO.FileInfo In files
                Dim fname As String = System.IO.Path.GetFileName(f.Name)
                DataGridView1.Rows.Add(fname)
            Next
        End If
    End Sub

    Private Sub 修正用一覧取得ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 修正用一覧取得ToolStripMenuItem.Click
        Dim fbd As New FolderBrowserDialog With {
            .Description = "フォルダを指定してください。",
            .RootFolder = Environment.SpecialFolder.Desktop,
            .SelectedPath = "",
            .ShowNewFolderButton = True
        }

        If fbd.ShowDialog(Me) = DialogResult.OK Then
            If DataGridView1.RowCount > 0 Then
                DataGridView1.Rows.Clear()
                DataGridView1.Columns.Clear()
            End If
            DataGridView1.Columns.Add(0, 0)
            DataGridView1.Columns.Add(1, 1)

            Dim di As New System.IO.DirectoryInfo(fbd.SelectedPath)
            Dim files As System.IO.FileInfo() = di.GetFiles("*.*", System.IO.SearchOption.AllDirectories)
            For Each f As System.IO.FileInfo In files
                Dim fname As String = System.IO.Path.GetFileName(f.Name)
                DataGridView1.Rows.Add(fname)
            Next

            Me.Text = fbd.SelectedPath
        End If
    End Sub

    Private Sub 修正実行ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 修正実行ToolStripMenuItem.Click
        If DataGridView1.RowCount > 0 Then
            ToolStripProgressBar1.Value = 0
            ToolStripProgressBar1.Maximum = DataGridView1.RowCount
            For r As Integer = 0 To DataGridView1.RowCount - 1
                If DataGridView1.Item(0, r).Value <> "" And DataGridView1.Item(1, r).Value <> "" Then
                    Dim motoPath As String = Me.Text & "\" & DataGridView1.Item(0, r).Value
                    Dim sakiPath As String = Me.Text & "\" & DataGridView1.Item(1, r).Value
                    If File.Exists(motoPath) Then
                        File.Move(motoPath, sakiPath)
                    End If
                End If
                ToolStripProgressBar1.Value += 1
            Next
            ToolStripProgressBar1.Value = 0
            MsgBox("修正しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    Private Sub URLでダウンロードToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles URLでダウンロードToolStripMenuItem.Click
        If URLでダウンロードToolStripMenuItem.Checked = True Then
            URLでダウンロードToolStripMenuItem.Checked = False
            ファイル名でフォルダ検索ToolStripMenuItem.Checked = True
        Else
            URLでダウンロードToolStripMenuItem.Checked = True
            ファイル名でフォルダ検索ToolStripMenuItem.Checked = False
        End If
    End Sub

    Private Sub ファイル名でフォルダ検索ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ファイル名でフォルダ検索ToolStripMenuItem.Click
        If ファイル名でフォルダ検索ToolStripMenuItem.Checked = False Then
            URLでダウンロードToolStripMenuItem.Checked = False
            ファイル名でフォルダ検索ToolStripMenuItem.Checked = True
        Else
            URLでダウンロードToolStripMenuItem.Checked = True
            ファイル名でフォルダ検索ToolStripMenuItem.Checked = False
        End If
    End Sub

    Private Sub ImgSrctextToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImgSrctextToolStripMenuItem.Click
        FileDL1(1, "<img src=""(?<text>.*?)"".*>|<img src='(?<text>.*?)'.*>")
    End Sub

    Private Sub 行ずつToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 行ずつToolStripMenuItem.Click
        FileDL1(2, "")
    End Sub

    Private Sub 行複数アドレスToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 行複数アドレスToolStripMenuItem.Click
        FileDL1(3, "")
    End Sub

    Private Sub FileDL1(ByRef mode As Integer, ByRef img As String)
        ServicePointManager.SecurityProtocol = DirectCast(3072, SecurityProtocolType)
        Dim DLcount As Integer = 0
        Dim fol1 As String = ""
        Dim fol2 As String = ""
        Dim urlMode As Integer = 0

        If URLでダウンロードToolStripMenuItem.Checked = True Then
            urlMode = 0
        Else
            urlMode = 1

            Dim fbd1 As New FolderBrowserDialog With {
                .Description = "検索するフォルダを指定してください。",
                .RootFolder = Environment.SpecialFolder.Desktop,
                .ShowNewFolderButton = True
            }
            If fbd1.ShowDialog(Me) = DialogResult.OK Then
                fol1 = fbd1.SelectedPath
            End If
        End If

        Dim fbd As New FolderBrowserDialog With {
            .Description = "保存するフォルダを指定してください。",
            .RootFolder = Environment.SpecialFolder.Desktop,
            .ShowNewFolderButton = True
        }
        If fbd.ShowDialog(Me) = DialogResult.OK Then
            fol2 = fbd.SelectedPath
        End If

        ToolStripProgressBar1.Maximum = DataGridView1.SelectedCells.Count
        ToolStripProgressBar1.Value = 0

        For i As Integer = 0 To DataGridView1.SelectedCells.Count - 1
            Dim htmlStr As String = DataGridView1.SelectedCells(i).Value
            htmlStr = Replace(htmlStr, vbCrLf, vbLf)
            htmlStr = Replace(htmlStr, vbCr, vbLf)
            Dim htmlLines As String() = Split(htmlStr, vbLf)
            For k As Integer = 0 To htmlLines.Length - 1
                Select Case True
                    Case InStr(htmlLines(k), ".jpg"), InStr(htmlLines(k), ".gif")
                        If mode = 1 Then
                            Dim re = New Regex(img, RegexOptions.IgnoreCase Or RegexOptions.Singleline)
                            Dim m As Match = re.Match(htmlLines(k))
                            While m.Success
                                Dim text As String = m.Groups("text").Value
                                If urlMode = 0 Then
                                    'MsgBox(htmlLines(k) & "/" & text, MsgBoxStyle.SystemModal)
                                    If ImgDL(fol2, text) = True Then
                                        DLcount += 1
                                    End If
                                    m = m.NextMatch()
                                Else
                                    Dim fname As String = Path.GetFileName(text)
                                    Dim stFilePathes As String() = Dialog.GetFilesMostDeep(fol1, fname)
                                    For Each stFilePath As String In stFilePathes
                                        Try
                                            File.Copy(stFilePath, fol2 & "\" & fname)
                                            DLcount += 1
                                        Catch ex As Exception

                                        End Try
                                    Next stFilePath
                                End If
                            End While
                        ElseIf mode = 2 Then
                            If urlMode = 0 Then
                                If ImgDL(fol2, htmlLines(k)) = True Then
                                    DLcount += 1
                                End If
                            Else
                                Dim fname As String = Path.GetFileName(htmlLines(k))
                                Dim stFilePathes As String() = Dialog.GetFilesMostDeep(fol1, fname)
                                For Each stFilePath As String In stFilePathes
                                    Try
                                        File.Copy(stFilePath, fol2 & "\" & fname)
                                        DLcount += 1
                                    Catch ex As Exception

                                    End Try
                                Next stFilePath
                            End If
                        ElseIf mode = 3 Then
                            If urlMode = 0 Then
                                Dim fArray As String() = Regex.Split(htmlLines(k), ",| ")
                                For Each fName As String In fArray
                                    If ImgDL(fol2, fName) = True Then
                                        DLcount += 1
                                    End If
                                Next
                            Else
                                Dim fArray As String() = Regex.Split(htmlLines(k), ",| ")
                                For Each f As String In fArray
                                    Dim fname As String = Path.GetFileName(f)
                                    Dim stFilePathes As String() = Dialog.GetFilesMostDeep(fol1, fname)
                                    For Each stFilePath As String In stFilePathes
                                        Try
                                            File.Copy(stFilePath, fol2 & "\" & fname)
                                            DLcount += 1
                                        Catch ex As Exception

                                        End Try
                                    Next stFilePath
                                Next
                            End If
                        End If
                End Select
            Next

            ToolStripProgressBar1.Value += 1
        Next

        ToolStripProgressBar1.Value = 0
        MsgBox(DLcount & "ファイルダウンロードしました", MsgBoxStyle.SystemModal)
    End Sub

    Private Function ImgDL(ByVal fol As String, ByVal url As String)
        If url <> "" Then
            Dim flag As Boolean = True
            Dim fName As String = Path.GetFileName(url)

            Try
                Dim wc As New System.Net.WebClient()
                wc.DownloadFile(url, fol & "\" & fName)
                wc.Dispose()
            Catch ex As Exception
                flag = False
            End Try
        End If

        Return True
    End Function

    'コードフォーマット
    Private Function Array_Create(ByVal s As String) As ArrayList
        s = Regex.Replace(s, ">", ">" & vbCrLf)
        Dim array As String() = Split(s, vbCrLf)

        For i As Integer = 0 To array.Length - 1
            array(i) = array(i).TrimStart()
        Next

        Dim str As New ArrayList(array)
        Return str
    End Function

    Private Function BR_ADD(ByVal htmlRecords As ArrayList) As ArrayList
        Dim str As ArrayList = New ArrayList
        Dim tagBR As String() = New String() {"<center", "</center", "<table", "</table>", "<td ", "</td>", "</tr>", "<img ", "<a ", "<b>", "<font", "br>", "<div", "-->"}

        For i As Integer = 0 To htmlRecords.Count - 1
            Dim flag As Boolean = False
            If InStr(htmlRecords(i), "<table") > 0 Then
                htmlRecords(i) = Replace(htmlRecords(i), "<table", vbCrLf & "<table")
            End If
            For c As Integer = 0 To tagBR.Length - 1
                If InStr(htmlRecords(i), tagBR(c)) > 0 Then
                    str.Add(htmlRecords(i) & vbCrLf)
                    flag = True
                    Exit For
                End If
            Next
            If flag = False Then
                str.Add(htmlRecords(i))
            End If
        Next

        Return str
    End Function

    Private Sub HTMLTIDYANDFORMATToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HTMLTIDYANDFORMATToolStripMenuItem.Click
        Dim htmlRecords As New ArrayList

        ToolStripProgressBar1.Value = 0
        ToolStripProgressBar1.Maximum = DataGridView1.SelectedCells.Count
        For i As Integer = 0 To DataGridView1.SelectedCells.Count - 1
            'ZetaHtmlTidy
            Dim html As String = DataGridView1.SelectedCells(i).Value
            Dim s As String = ""
            Using tidy As HtmlTidy = New HtmlTidy()
                s = tidy.CleanHtml(html, HtmlTidyOptions.ConvertToXhtml)
            End Using

            s = Replace(s, vbCrLf, "")
            s = Replace(s, vbLf, "")
            s = Replace(s, vbCr, "")
            s = Regex.Replace(s, "<!DOCTYPE h.+HTML Tidy.+head><body>", "")
            's = Regex.Replace(s, "<!DOCTYPE h.+HTML Tidy.+/>", "")
            's = Regex.Replace(s, "<meta.+HTML Tidy.+head><body>", "")
            s = Replace(s, "</body></html>", "")
            s = Replace(s, " />", ">")
            s = LTrim(s)
            s = Replace(s, "  <", "<")
            s = Replace(s, " <", "<")
            s = Replace(s, vbTab, "")

            htmlRecords = Array_Create(s)
            htmlRecords = BR_ADD(htmlRecords)

            Dim str As String = ""
            For k As Integer = 0 To htmlRecords.Count - 1
                str &= htmlRecords(k)
            Next

            DataGridView1.SelectedCells(i).Value = str
            ToolStripProgressBar1.Value += 1
        Next

        ToolStripProgressBar1.Value = 0
        MsgBox("終了しました", MsgBoxStyle.SystemModal)
    End Sub

    Private Sub 英数全角チェックToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 英数全角チェックToolStripMenuItem.Click
        Dim c As Integer = 0
        For i As Integer = 0 To DataGridView1.SelectedCells.Count - 1
            If DataGridView1.SelectedCells(i).Value <> "" Then
                Dim checkStr As String = DataGridView1.SelectedCells(i).Value
                checkStr = Regex.Replace(checkStr, "[^0-9A-Za-z０-９Ａ-Ｚａ-ｚ]", "")
                If Not (IsOneByteChar(checkStr)) Then
                    DataGridView1.SelectedCells(i).Style.BackColor = Color.Red
                    c += 1
                End If
            End If
        Next
        MsgBox(c & "件ありました。", MsgBoxStyle.SystemModal)
    End Sub

    '1バイト文字で構成された文字列であるか判定
    Private Function IsOneByteChar(ByVal str As String) As Boolean
        Dim byte_data As Byte() = System.Text.Encoding.GetEncoding(932).GetBytes(str)
        If (byte_data.Length = str.Length) Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub 行目ヘッダー固定ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 行目ヘッダー固定ToolStripMenuItem.Click
        If 行目ヘッダー固定ToolStripMenuItem.Checked = True Then
            DataGridView1.Rows(0).Frozen = False
            DataGridView1.Rows(0).DividerHeight = 0
            行目ヘッダー固定ToolStripMenuItem.Checked = False
        Else
            DataGridView1.Rows(0).Frozen = True
            DataGridView1.Rows(0).DividerHeight = 3
            行目ヘッダー固定ToolStripMenuItem.Checked = True
        End If
    End Sub

    Private Sub イメージリスト操作楽天ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles イメージリスト操作楽天ToolStripMenuItem.Click
        If DataGridView1.RowCount > 0 Then
            Dialog.ComboBox2.Items.Clear()

            Dim DGVheaderCheck As ArrayList = New ArrayList
            For c As Integer = 0 To DataGridView1.Columns.Count - 1
                DGVheaderCheck.Add(DataGridView1.Item(c, 0).Value)
                Dialog.ComboBox2.Items.Add(DataGridView1.Item(c, 0).Value)
                Dialog.ComboBox3.Items.Add(DataGridView1.Item(c, 0).Value)
                Dialog.ComboBox4.Items.Add(DataGridView1.Item(c, 0).Value)
                Dialog.ComboBox5.Items.Add(DataGridView1.Item(c, 0).Value)
            Next c

            Dialog.ComboBox2.SelectedItem = "商品画像URL"
            Dialog.ComboBox3.SelectedItem = "商品画像名（ALT）"
            Dialog.ComboBox4.SelectedItem = "商品番号"
            Dialog.ComboBox5.SelectedItem = "PC用販売説明文"

            Dialog.TabControl1.SelectedIndex = 3
            Form1.SetTabVisible(Dialog.TabControl1, "イメージリスト")
            Dialog.Show()
        End If
    End Sub

    Private Sub ゆうプリ出力データより出荷分解析ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ゆうプリ出力データより出荷分解析ToolStripMenuItem.Click
        If DataGridView1.RowCount > 0 Then
            Dim DGVheaderCheck As ArrayList = New ArrayList
            For c As Integer = 0 To DataGridView1.Columns.Count - 1
                DGVheaderCheck.Add(DataGridView1.Item(c, 0).Value)
            Next c

            Dim StArray As ArrayList = New ArrayList
            Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\config\ゆうプリ配達ステータス.txt"
            Using sr As New System.IO.StreamReader(fName, System.Text.Encoding.Default)
                Do While Not sr.EndOfStream
                    Dim s As String = sr.ReadLine
                    StArray.Add(s)
                Loop
            End Using

            DataGridView1.Columns.Add("99", "配達ステータス")

            For r As Integer = 0 To DataGridView1.RowCount - 1
                If DataGridView1.Rows(r).IsNewRow = False Then
                    Dim sCode As String = DataGridView1.Item(DGVheaderCheck.IndexOf("配達ステータス集約コード"), r).Value
                    Dim mCode As String = DataGridView1.Item(DGVheaderCheck.IndexOf("配達ステータス明細コード"), r).Value
                    If sCode = "" And mCode = "" Then
                        'ステータスが入っていない行
                        DataGridView1.Rows(r).Visible = False
                    Else
                        For i As Integer = 0 To StArray.Count - 1
                            If InStr(StArray(i), sCode & splitter & mCode) > 0 Then
                                Dim st As String() = Split(StArray(i), splitter)
                                DataGridView1.Item(DataGridView1.ColumnCount - 1, r).Value = st(2)
                                If st(3) = "0" Then
                                    DataGridView1.Rows(r).Visible = False
                                Else
                                    DataGridView1.Rows(r).Visible = True
                                End If
                                Exit For
                            End If
                        Next
                    End If
                End If
            Next

            MsgBox("分析終了しました", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
        End If
    End Sub

    Private Sub Csv_Activated(sender As Form, e As EventArgs) Handles Me.Activated
        If sender Is Form1.CSV_FORMS(1) Then
            If Form1.CSV_FORMS(0).ToolStripStatusLabel5.Text <> "2" Then
                Form1.CSV_FORMS(0).ToolStripStatusLabel5.Text = "2"
                Form1.CSV_FORMS(1).ToolStripStatusLabel5.Text = "2"
                HTMLcreate.Label1.Text = "2"
            End If
        Else
            If Form1.CSV_FORMS(0).ToolStripStatusLabel5.Text <> "1" Then
                Form1.CSV_FORMS(0).ToolStripStatusLabel5.Text = "1"
                Form1.CSV_FORMS(1).ToolStripStatusLabel5.Text = "1"
                HTMLcreate.Label1.Text = "1"
            End If
        End If
    End Sub

    '=================================================================================================================================
    '文字コード、改行コード変更
    Private Sub ToolStripDropDownButton10_Click(sender As ToolStripMenuItem, e As EventArgs) Handles _
            ShiftJISToolStripMenuItem.Click, UTF8ToolStripMenuItem.Click, EUCJPToolStripMenuItem.Click, utf16ToolStripMenuItem.Click,
            EUCCNToolStripMenuItem.Click
        ToolStripDropDownButton10.Text = sender.Text
    End Sub

    Private Sub ToolStripDropDownButton9_Click(sender As ToolStripMenuItem, e As EventArgs) Handles _
            CRLFToolStripMenuItem.Click, LFToolStripMenuItem.Click, CRToolStripMenuItem.Click
        ToolStripDropDownButton9.Text = sender.Text
    End Sub



    '=================================================================================================================================


End Class