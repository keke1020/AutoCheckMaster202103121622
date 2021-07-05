Imports System.IO
Imports System.Text.RegularExpressions
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel

Public Class Csv_change
    Dim secValue As String = ""
    Dim dgv As DataGridView = DataGridView1

    Dim splitter As String = splitter    '区切り文字


    '*****************************************************************
    '一括処理用
    '*****************************************************************
    Dim AutoMode As Boolean = False

    Private Sub Csv_change_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        secValue = "csv_change"   'iniファイルセクション

        '******************************************
        'dgv画面チラつき防止
        '******************************************
        'DataGirdViewのTypeを取得
        Dim dgvtype As Type = GetType(DataGridView)
        'プロパティ設定の取得
        Dim dgvPropertyInfo As Reflection.PropertyInfo = dgvtype.GetProperty("DoubleBuffered", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)

        '対象のDataGridViewにtrueをセットする
        Dim dgvArray As DataGridView() = {DataGridView1, DataGridView2, DataGridView3, DataGridView4, DataGridView5, DataGridView6, DataGridView7, DataGridView8}
        For i As Integer = 0 To dgvArray.Length - 1
            dgvPropertyInfo.SetValue(dgvArray(i), True, Nothing)
        Next

        dgv = DataGridView1
        TemplateFolderRead()
        TemplateFolderRead2()
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        ComboBox8.SelectedIndex = 0
        ComboBox9.SelectedIndex = 0

        SplitContainer2.SplitterDistance = SplitContainer2.Height
        Button11.Visible = False
        Button12.Visible = False
        Button13.Visible = False
        Button14.Visible = False
        Button15.Visible = False

        If Form1.inif.ReadINI(secValue, "yupackCheck", "") <> "" Then
            CheckBox2.Checked = Form1.inif.ReadINI(secValue, "yupackCheck", "")
        End If

    End Sub

    Dim combobox1Array As ArrayList = New ArrayList
    Public Sub TemplateFolderRead()
        Dim di As New DirectoryInfo(Path.GetDirectoryName(Form1.appPath) & "\template")
        Dim files As FileInfo() = di.GetFiles("*.dat", SearchOption.AllDirectories)

        Dim fArray As ArrayList = New ArrayList
        For Each f As FileInfo In files
            Dim fname As String() = Split(Path.GetFileNameWithoutExtension(f.Name), "、")
            If Regex.IsMatch(fname(0), "★|出荷|発送完了|Qoo10\(select\)変換|Qoo10直（通販堂）配（日本郵便）変換|伝票起票変換|ヤフオク") Then
                fArray.Add(fname(0))
            ElseIf InStr(fname(0), "合併") > 0 Then
                ComboBox2.Items.Add(fname(0))
            End If
        Next

        'ソートしてComboBox1に結果を表示する
        fArray.Sort()
        combobox1Array.Clear()
        combobox1Array.Add("選択してください")
        For i As Integer = 0 To fArray.Count - 1
            ComboBox1.Items.Add(fArray(i))
            combobox1Array.Add(fArray(i))
        Next
    End Sub

    Private Sub TemplateFolderRead2()
        Dim di As New DirectoryInfo(Path.GetDirectoryName(Form1.appPath) & "\template")
        Dim files As FileInfo() = di.GetFiles("*.dat", SearchOption.AllDirectories)

        Dim fArray As ArrayList = New ArrayList
        For Each f As FileInfo In files
            Dim fname As String = Path.GetFileNameWithoutExtension(f.Name)
            If InStr(fname, "Program") > 0 Then
                ComboBox8.Items.Add(fname)
            End If
        Next
    End Sub

    Private Sub DataGridView1_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As DataGridViewCellMouseEventArgs) Handles DataGridView1.ColumnHeaderMouseClick
        ColumnHeaderMouseClick(sender, e)
    End Sub

    Private Sub DataGridView5_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As DataGridViewCellMouseEventArgs) Handles DataGridView5.ColumnHeaderMouseClick
        ColumnHeaderMouseClick(sender, e)
    End Sub

    Private Sub ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As DataGridViewCellMouseEventArgs)
        dgv.ClearSelection()
        For r As Integer = 0 To dgv.RowCount - 1
            dgv.Item(e.ColumnIndex, r).Selected = True
        Next
    End Sub

    Private Sub DataGridView1_DragDrop(ByVal sender As Object, ByVal e As DragEventArgs) Handles DataGridView1.DragDrop
        DgvDragDrop(sender, e)
    End Sub

    Private Sub DataGridView5_DragDrop(ByVal sender As Object, ByVal e As DragEventArgs) Handles DataGridView5.DragDrop
        DgvDragDrop(sender, e)
    End Sub

    Private Sub DgvDragDrop(ByVal sender As Object, ByVal e As DragEventArgs)
        BT6_push = False    '処理したかチェックフラグ

        Dim fCount As Integer = 0
        Dim dataPathArray As ArrayList = New ArrayList
        For Each filename As String In e.Data.GetData(DataFormats.FileDrop)
            dataPathArray.Add(filename)
            fCount += 1
        Next

        If dgv.RowCount > 1 Then
            Dim DR As DialogResult = MsgBox("処理しますか", MsgBoxStyle.OkCancel Or MsgBoxStyle.SystemModal)
            If DR = Windows.Forms.DialogResult.Cancel Then
                Exit Sub
            End If
        End If

        ToolStripProgressBar1.Value = 0
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------

        If dgv.RowCount > 1 Then
            For Each filename As String In dataPathArray
                Dim csvRecords As New ArrayList()
                csvRecords = CSV_READ(filename)
                ToolStripProgressBar1.Maximum = csvRecords.Count

                If TabControl1.SelectedIndex = 1 Then
                    CSV_GAPPEI(csvRecords)
                End If

                Exit For
            Next
        Else
            For Each filename As String In dataPathArray
                If dgv Is DataGridView1 Then
                    ToolStripLabel2.Text = filename
                Else
                    ToolStripStatusLabel1.Text = filename
                End If
                Dim csvRecords As New ArrayList()
                csvRecords = CSV_READ(filename)
                ToolStripProgressBar1.Maximum = csvRecords.Count

                For r As Integer = 0 To csvRecords.Count - 1
                    Dim sArray As String() = Split(csvRecords(r), "|=|")
                    If dgv.ColumnCount <= sArray.Length - 1 Then
                        For c As Integer = dgv.ColumnCount To sArray.Length - 1
                            dgv.Columns.Add(c, c)
                            dgv.Columns(c).SortMode = DataGridViewColumnSortMode.NotSortable
                        Next
                    End If
                    dgv.Rows.Add(sArray)

                    ToolStripProgressBar1.Value += 1
                Next

                Exit For
            Next
        End If

        ToolStripProgressBar1.Value = 0
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------

        CsvCheck()
    End Sub

    Private Sub CsvCheck()
        If TabControl1.SelectedTab.Text = "データ変換" Then
            Dim header As ArrayList = New ArrayList
            For c As Integer = 0 To DataGridView1.ColumnCount - 1
                header.Add(DataGridView1.Item(c, 0).Value)
            Next

            TextBox10.BackColor = Color.White
            TextBox10.Text = "正常"
            For i As Integer = 0 To DataGridView2.RowCount - 1
                Dim res As String = DataGridView2.Item(1, i).Value
                If res <> "" Then
                    Dim outputRes As String = Regex.Replace(res, "\[.*\]", "", RegexOptions.Singleline)
                    If outputRes <> "" Then
                        If header.Contains(outputRes) = False Then
                            TextBox10.Text = outputRes
                            TextBox10.BackColor = Color.Yellow
                            Exit For
                        End If
                    End If
                End If
            Next
        End If
    End Sub

    Private Sub DataGridView1_DragEnter(ByVal sender As Object, ByVal e As DragEventArgs) Handles DataGridView1.DragEnter
        DgvDragEnter(sender, e)
    End Sub

    Private Sub DataGridView5_DragEnter(ByVal sender As Object, ByVal e As DragEventArgs) Handles DataGridView5.DragEnter
        DgvDragEnter(sender, e)
    End Sub

    Private Sub DgvDragEnter(ByVal sender As Object, ByVal e As DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Function CSV_READ2(ByRef path As String)
        Dim csvRecords As New ArrayList()

        'CSVファイル名
        Dim csvFileName As String = path

        'Shift JISで読み込む
        Dim tfp As New FileIO.TextFieldParser(csvFileName, System.Text.Encoding.GetEncoding(932)) With {
            .TextFieldType = FileIO.FieldType.Delimited,
            .Delimiters = New String() {splitter},
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

    Private Function CSV_READ(ByRef path As String)
        Dim csvText As String = File.ReadAllText(path, System.Text.Encoding.Default)

        'カンマ、タブ等の区切り文字認識
        Dim header As String() = Regex.Split(csvText, vbCrLf & "|" & vbCr & "|" & vbLf)
        'Dim header As String() = Split(csvText, vbCrLf)
        If InStr(header(0), vbTab) Then
            splitter = vbTab
            ToolStripStatusLabel2.Text = "tab"
        Else
            splitter = ","
            ToolStripStatusLabel2.Text = "comma"
        End If

        '前後の改行を削除しておく
        csvText = csvText.Trim(
            New Char() {ControlChars.Cr, ControlChars.Lf})

        Dim csvRecords As New ArrayList
        Dim csvFields As New ArrayList

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
                csvFields = New ArrayList(csvFields.Count)

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

        Return csvRecords
    End Function

    Dim sc As Integer = 0
    Dim sr As Integer = 0
    Public Sub PreDIV()
        '-----------------------------------------------------------------------
        sc = dgv.HorizontalScrollingOffset
        sr = dgv.FirstDisplayedScrollingRowIndex
        dgv.ScrollBars = ScrollBars.None
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
        dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        dgv.ScrollBars = ScrollBars.Both
        '-----------------------------------------------------------------------
    End Sub

    Public Sub RetDIV()
        '-----------------------------------------------------------------------
        dgv.HorizontalScrollingOffset = sc
        '***************************************************
        If sr > dgv.RowCount - 2 Then
            sr = dgv.RowCount - 2
        End If
        If sr < 0 Then
            sr = 0
        End If
        dgv.FirstDisplayedScrollingRowIndex = sr
        'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        '-----------------------------------------------------------------------
    End Sub

    Private Sub CSV_GAPPEI(ByVal csvRecords As ArrayList, Optional mode As Integer = 0)
        Dim DB1 As DataGridView = dgv
        Dim DB2 As DataGridView = DataGridView3
        Dim kensaku1 As ArrayList = New ArrayList
        Dim kensaku2 As ArrayList = New ArrayList
        Dim kensakuNum1 As ArrayList = New ArrayList
        Dim kensakuNum2 As ArrayList = New ArrayList
        Dim tuika As ArrayList = New ArrayList
        Dim tuikaNum As ArrayList = New ArrayList
        Dim keisan1 As ArrayList = New ArrayList
        Dim keisan2 As ArrayList = New ArrayList
        Dim keisanNum1 As ArrayList = New ArrayList
        Dim keisanNum2 As ArrayList = New ArrayList
        Dim okikae1 As ArrayList = New ArrayList
        Dim okikae2 As ArrayList = New ArrayList
        Dim okikaeNum1 As ArrayList = New ArrayList
        Dim okikaeNum2 As ArrayList = New ArrayList

        For r As Integer = 0 To DB2.RowCount - 2
            If DB2.Item(0, r).Value = "検索" Then
                kensaku1.Add(DB2.Item(1, r).Value)
                kensaku2.Add(DB2.Item(2, r).Value)
            ElseIf DB2.Item(0, r).Value = "追加" Then
                tuika.Add(DB2.Item(2, r).Value)
            ElseIf DB2.Item(0, r).Value = "置換" Then   '置換未実装
                okikae1.Add(DB2.Item(1, r).Value)
                okikae2.Add(DB2.Item(1, r).Value)
            Else
                keisan1.Add(DB2.Item(1, r).Value)
                keisan2.Add(DB2.Item(2, r).Value)
            End If
        Next

        '*****************************************************
        ' 各コラム番号を取得
        '*****************************************************

        '元データ（コラム番号を取得）
        If RadioButton1.Checked = True Then
            For i As Integer = 0 To kensaku1.Count - 1
                For c As Integer = 0 To DB1.ColumnCount - 1
                    If kensaku1(i) = DB1.Item(c, 0).Value Then
                        Dim a = DB1.Item(c, 0).Value
                        kensakuNum1.Add(c)
                        Exit For
                    End If
                Next
            Next
        Else
            For i As Integer = 0 To kensaku1.Count - 1
                kensakuNum1.Add(kensaku1(i))
            Next
        End If

        '読込データ（コラム番号を取得）
        If RadioButton1.Checked = True Then
            For i As Integer = 0 To kensaku2.Count - 1
                Dim CR As String() = Split(csvRecords(0), "|=|")
                For c As Integer = 0 To CR.Length - 1
                    If kensaku2(i) = CR(c) Then
                        kensakuNum2.Add(c)
                        Exit For
                    End If
                Next
            Next
        Else
            For i As Integer = 0 To kensaku2.Count - 1
                kensakuNum2.Add(kensaku2(i))
            Next
        End If

        '検索データ整合
        If kensakuNum1.Count <> kensakuNum2.Count Then
            MsgBox("2ファイルの検索データ（検索）が合わないため合併できません", MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        '追加データ（コラム番号を取得）
        If RadioButton1.Checked = True Then
            For i As Integer = 0 To tuika.Count - 1
                Dim CR As String() = Split(csvRecords(0), "|=|")
                For c As Integer = 0 To CR.Length - 1
                    If tuika(i) = CR(c) Then
                        tuikaNum.Add(c)
                    End If
                Next
            Next
        Else
            For i As Integer = 0 To tuika.Count - 1
                tuikaNum.Add(tuika(i))
            Next
        End If

        '追加データ整合
        If tuika.Count <> tuikaNum.Count Then
            MsgBox("読込ファイルの追加データ（追加）が合わないため合併できません", MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        '置換データ（コラム番号を取得）
        If RadioButton1.Checked = True Then
            For i As Integer = 0 To okikae1.Count - 1
                For c As Integer = 0 To DB1.ColumnCount - 1
                    If okikae1(i) = DB1.Item(c, 0).Value Then
                        Dim a = DB1.Item(c, 0).Value
                        okikaeNum1.Add(c)
                        Exit For
                    End If
                Next
            Next
        Else
            For i As Integer = 0 To okikae1.Count - 1
                okikaeNum1.Add(okikae1(i))
            Next
        End If

        If RadioButton1.Checked = True Then
            For i As Integer = 0 To okikae2.Count - 1
                Dim CR As String() = Split(csvRecords(0), "|=|")
                For c As Integer = 0 To CR.Length - 1
                    If okikae2(i) = CR(c) Then
                        okikaeNum2.Add(c)
                    End If
                Next
            Next
        Else
            For i As Integer = 0 To okikae2.Count - 1
                okikaeNum2.Add(tuika(i))
            Next
        End If

        '置換データ整合
        If okikae2.Count <> okikaeNum2.Count Then
            MsgBox("読込ファイルの置換データ（置換）が合わないため合併できません", MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        'ヘッダー追加
        For i As Integer = 0 To tuikaNum.Count - 1
            DB1.Columns.Add("column" & DB1.ColumnCount, DB1.ColumnCount)
            DB1.Columns(DB1.ColumnCount - 1).SortMode = DataGridViewColumnSortMode.NotSortable
            DB1.Item(DB1.ColumnCount - 1, 0).Value = tuika(i)
        Next

        ToolStripProgressBar1.Value = 0
        ToolStripProgressBar1.Maximum = DB1.RowCount

        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        If mode = 0 Then
            'データ検索
            For r As Integer = 0 To DB1.RowCount - 1
                Dim ken1 As String = ""
                For c As Integer = 0 To kensakuNum1.Count - 1
                    ken1 &= DB1.Item(kensakuNum1(c), r).value
                Next
                For rr As Integer = 0 To csvRecords.Count - 1
                    'If InStr(csvRecords(rr), "|=|") > 0 Then
                    Dim CR As String() = Split(csvRecords(rr), "|=|")
                    Dim ken2 As String = ""
                    For c As Integer = 0 To kensakuNum2.Count - 1
                        ken2 &= CR(kensakuNum2(c))
                    Next
                    If CheckBox3.Checked = True Then
                        ken1 = StrConv(ken1, VbStrConv.Narrow)
                        ken2 = StrConv(ken2, VbStrConv.Narrow)
                        ken1 = StrConv(ken1, VbStrConv.Lowercase)
                        ken2 = StrConv(ken2, VbStrConv.Lowercase)
                    End If
                    If CheckBox8.Checked = True Then    '文字列検索（通常は部分一致）
                        Dim pattern As String = ""
                        Select Case ComboBox9.SelectedItem
                            Case "部分一致"
                                pattern = ".*" & ken2 & ".*"
                            Case "完全一致"
                                pattern = "^" & ken2 & "$"
                            Case "前方一致"
                                pattern = "^" & ken2 & ".*"
                            Case "後方一致"
                                pattern = ".*" & ken2 & "$"
                        End Select
                        If Regex.IsMatch(ken1, pattern) Then
                            If tuikaNum.Count > 0 Then
                                Dim cnum As Integer = DB1.ColumnCount - 1
                                DB1.Item(cnum, r).Value = CR(tuikaNum(0))
                            End If
                            If okikaeNum1.Count > 0 Then
                                DB1.Item(okikaeNum1(0), r).Value = CR(okikaeNum2(0))
                                DB1.Item(okikaeNum1(0), r).Style.BackColor = Color.LightYellow
                            End If
                            Exit For
                        End If
                    Else
                        If ken1 = ken2 Then
                            If tuikaNum.Count > 0 Then
                                Dim cnum As Integer = DB1.ColumnCount - 1
                                DB1.Item(cnum, r).Value = CR(tuikaNum(0))
                            End If
                            If okikaeNum1.Count > 0 Then
                                DB1.Item(okikaeNum1(0), r).Value = CR(okikaeNum2(0))
                                DB1.Item(okikaeNum1(0), r).Style.BackColor = Color.LightYellow
                            End If
                            Exit For
                        End If
                    End If
                    'End If
                Next

                ToolStripProgressBar1.Value += 1
                If r Mod 100 = 0 Then
                    Application.DoEvents()
                End If
            Next
        Else
            Dim dgv1 As DataGridView = DataGridView1
            Dim dgv5 As DataGridView = DataGridView5
            Dim dgv5Array As New ArrayList
            For r5 As Integer = 0 To dgv5.RowCount - 1
                Dim ken2 As String = ""
                For c As Integer = 0 To kensakuNum2.Count - 1
                    ken2 &= dgv5.Item(kensakuNum2(c), r5).value
                Next
                dgv5Array.Add(ken2)
            Next
            For r As Integer = 0 To dgv1.RowCount - 1
                Dim ken1 As String = ""
                For c As Integer = 0 To kensakuNum1.Count - 1
                    ken1 &= dgv1.Item(kensakuNum1(c), r).value
                Next
                For r5 As Integer = 0 To dgv5.RowCount - 1
                    If CheckBox8.Checked = True Then    '文字列検索（通常は部分一致）
                        Dim ken2 As String = ""
                        For c As Integer = 0 To kensakuNum2.Count - 1
                            ken2 &= dgv5.Item(kensakuNum2(c), r5).value
                        Next
                        If CheckBox3.Checked = True Then
                            ken1 = StrConv(ken1, VbStrConv.Narrow)
                            ken2 = StrConv(ken2, VbStrConv.Narrow)
                            ken1 = StrConv(ken1, VbStrConv.Lowercase)
                            ken2 = StrConv(ken2, VbStrConv.Lowercase)
                        End If
                        Dim pattern As String = ""
                        Select Case ComboBox9.SelectedItem
                            Case "部分一致"
                                pattern = ".*" & ken2 & ".*"
                            Case "完全一致"
                                pattern = "^" & ken2 & "$"
                            Case "前方一致"
                                pattern = "^" & ken2 & ".*"
                            Case "後方一致"
                                pattern = ".*" & ken2 & "$"
                        End Select
                        If Regex.IsMatch(ken1, pattern) Then
                            If tuikaNum.Count > 0 Then
                                Dim cnum As Integer = dgv1.ColumnCount - 1
                                dgv1.Item(cnum, r).Value = dgv5.Item(tuikaNum(0), r5).value
                            End If
                            If okikaeNum1.Count > 0 Then
                                dgv1.Item(okikaeNum1(0), r).Value = dgv5.Item(okikaeNum2(0), r5).value
                                dgv1.Item(okikaeNum1(0), r).Style.BackColor = Color.LightYellow
                            End If
                            Exit For
                        End If
                    Else
                        If dgv5Array.Contains(ken1) Then
                            Dim r5row As Integer = dgv5Array.IndexOf(ken1)
                            If tuikaNum.Count > 0 Then
                                Dim cnum As Integer = DB1.ColumnCount - 1
                                dgv1.Item(cnum, r).Value = dgv5.Item(tuikaNum(0), r5row).value
                            End If
                            If okikaeNum1.Count > 0 Then
                                dgv1.Item(okikaeNum1(0), r).Value = dgv5.Item(okikaeNum2(0), r5row).value
                                dgv1.Item(okikaeNum1(0), r).Style.BackColor = Color.LightYellow
                            End If
                            Exit For
                        End If
                    End If
                Next

                ToolStripProgressBar1.Value += 1
                If r Mod 100 = 0 Then
                    Application.DoEvents()
                End If
            Next
        End If
        ToolStripProgressBar1.Value = 0

        '-----------------------------------------------------------------------
        '計算一括処理
        If keisan1.Count > 0 Then
            '元）計算データ（コラム番号を取得）
            If RadioButton1.Checked = True Then
                For i As Integer = 0 To keisan1.Count - 1
                    For c As Integer = 0 To DB1.ColumnCount - 1
                        If keisan1(i) = DB1.Item(c, 0).Value Then
                            Dim a = DB1.Item(c, 0).Value
                            keisanNum1.Add(c)
                            Exit For
                        End If
                    Next
                Next
            Else
                For i As Integer = 0 To keisan1.Count - 1
                    keisanNum1.Add(keisan1(i))
                Next
            End If

            '読込）計算データ（コラム番号を取得）
            If RadioButton1.Checked = True Then
                For i As Integer = 0 To keisan2.Count - 1
                    For c As Integer = 0 To DB1.ColumnCount - 1
                        If keisan2(i) = DB1.Item(c, 0).Value Then
                            Dim a = DB1.Item(c, 0).Value
                            keisanNum2.Add(c)
                            Exit For
                        End If
                    Next
                Next
            Else
                For i As Integer = 0 To keisan2.Count - 1
                    keisanNum2.Add(keisan2(i))
                Next
            End If

            '計算データ整合
            If keisanNum1.Count <> keisanNum2.Count Then
                MsgBox("読込ファイルの追加データ（計算）が合わないため合併できません", MsgBoxStyle.SystemModal)
                Exit Sub
            End If

            'ヘッダー追加（2列以上の計算にこれでは対応していない）
            For i As Integer = 0 To tuikaNum.Count - 1
                DB1.Columns.Add("column" & DB1.ColumnCount, DB1.ColumnCount)
                DB1.Columns(DB1.ColumnCount - 1).SortMode = DataGridViewColumnSortMode.NotSortable
                DB1.Item(DB1.ColumnCount - 1, 0).Value = "計算"
            Next

            For r As Integer = 0 To DB1.RowCount - 1
                If IsNumeric(DB1.Item(keisanNum1(0), r).value) = True And IsNumeric(DB1.Item(keisanNum2(0), r).value) = True Then
                    Dim cnum As Integer = DB1.ColumnCount - 1
                    DB1.Item(cnum, r).Value = DB1.Item(keisanNum2(0), r).value - DB1.Item(keisanNum1(0), r).value
                End If
                ToolStripProgressBar1.Value += 1
            Next
        End If
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------

        ToolStripProgressBar1.Value = 0
    End Sub

    'データ合併 保存
    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button10.Click
        Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\template\" & ComboBox2.Text & ".dat"
        If File.Exists(fName) = True Then
            MsgBox("同じ名前のファイルがあります", MsgBoxStyle.SystemModal)
            Exit Sub
        ElseIf InStr(fName, "合併") = 0 Then
            MsgBox("ファイル名に「合併」を入れてください", MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        Dim str As String = ""
        For r As Integer = 0 To DataGridView3.RowCount - 2
            str &= DataGridView3.Item(0, r).Value & splitter
            str &= DataGridView3.Item(1, r).Value & splitter
            str &= DataGridView3.Item(2, r).Value & vbCrLf
        Next

        Dim sw As StreamWriter = New StreamWriter(fName, False, System.Text.Encoding.GetEncoding("SHIFT-JIS"))
        sw.Write(str)
        sw.Close()
        ReloadCombobox()

        If AutoMode = False Then
            MsgBox("保存しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button9.Click
        If ComboBox2.SelectedIndex = 0 Then
            Exit Sub
        End If

        Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\template\" & ComboBox2.SelectedItem & ".dat"
        File.Delete(fName)
        ReloadCombobox()
    End Sub

    'データ変換 保存
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As EventArgs)
        Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\template\" & ComboBox1.Text & ".dat"
        If InStr(fName, "変換") = 0 Then
            MsgBox("ファイル名に「変換」を入れてください", MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        If File.Exists(fName) = True Then
            Dim DR As DialogResult = MsgBox("同名のテンプレがあります。上書きしますか？", MsgBoxStyle.SystemModal)
            If DR = Windows.Forms.DialogResult.Cancel Then
                Exit Sub
            End If
        End If

        Dim str As String = ""
        For r As Integer = 0 To DataGridView2.RowCount - 2
            str &= DataGridView2.Item(0, r).Value & splitter
            str &= DataGridView2.Item(1, r).Value & splitter
            str &= DataGridView2.Item(2, r).Value & vbCrLf
        Next

        SaveCsv(fName, DataGridView2, False, 1)
        'Dim sw As IO.StreamWriter = New IO.StreamWriter(fName, False, System.Text.Encoding.GetEncoding("SHIFT-JIS"))
        'sw.Write(str)
        'sw.Close()
        ReloadCombobox()
        MsgBox("保存しました", MsgBoxStyle.SystemModal)
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As EventArgs)
        If ComboBox1.SelectedIndex = 0 Then
            Exit Sub
        End If

        Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\template\" & ComboBox1.SelectedItem & ".dat"
        File.Delete(fName)
        ReloadCombobox()
    End Sub

    Private Sub ReloadCombobox()
        ComboBox1.Items.Clear()
        ComboBox1.Items.Add("選択してください")
        ComboBox2.Items.Clear()
        ComboBox2.Items.Add("選択してください")
        TemplateFolderRead()
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
    End Sub

    '合併コンボボックス
    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.SelectedIndex = 0 Then
            Exit Sub
        End If

        DataGridView3.Rows.Clear()

        'ファイル読み取り
        Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\template\" & ComboBox2.SelectedItem & ".dat"
        Using sr As New StreamReader(fName, System.Text.Encoding.Default)
            Do While Not sr.EndOfStream
                Dim s As String = sr.ReadLine
                Dim sArray As String() = Split(s, splitter)
                DataGridView3.Rows.Add(sArray)
            Loop
        End Using
    End Sub

    Private Sub ComboBox1_DropDown(ByVal sender As Object, ByVal e As EventArgs) Handles ComboBox1.DropDown
        ComboBox1.Items.Clear()

        Dim nArray As ArrayList = New ArrayList
        '*********************************************
        Dim iniStr As String = Form1.inif.ReadINI(secValue, "csvCList", "")
        Dim sArray As String() = Split(iniStr, splitter)
        '*********************************************
        'Dim sArray As String() = Split(My.Settings.csvCList, splitter)
        For i As Integer = 0 To combobox1Array.Count - 1
            Dim flag As Boolean = False
            If InStr(combobox1Array(i), "配") > 0 Then
                For k As Integer = 0 To sArray.Count - 1
                    Dim sa As String = RightStr(sArray(k), 2)
                    '特殊な文字だけ変換
                    If sArray(k) = "通販の暁" Then
                        sa = "あかつき"
                    ElseIf sArray(k) = "ヤフオク よかろうもん" Then
                        sa = "ヤフオク"
                    ElseIf sArray(k) = "よかろうもん" Then
                        sa = "Yよか"
                    ElseIf sArray(k) = "通販の雑貨倉庫" Then
                        sa = "雑貨"
                    End If
                    If InStr(combobox1Array(i), sa) > 0 Then
                        flag = True
                        Exit For
                    End If
                Next
            Else
                flag = True
            End If
            If flag = True Then
                nArray.Add(combobox1Array(i))
            End If
        Next

        For i As Integer = 0 To nArray.Count - 1
            ComboBox1.Items.Add(nArray(i))
        Next
    End Sub

    Public Shared Function RightStr(ByVal stTarget As String, ByVal iLength As Integer) As String
        If iLength <= stTarget.Length Then
            Return stTarget.Substring(stTarget.Length - iLength)
        End If

        Return stTarget
    End Function

    '変換コンボボックス
    Dim hKaisya As String = ""
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedIndex = 0 Then
            Exit Sub
        End If

        DataGridView2.Rows.Clear()
        BT6_push = False    '処理したかチェックフラグ

        'ファイル読み取り
        Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\template\" & ComboBox1.SelectedItem & ".dat"
        Dim csvRecords As New ArrayList()
        csvRecords = CSV_READ(fName)
        For r As Integer = 0 To csvRecords.Count - 1
            Dim sArray As String() = Split(csvRecords(r), "|=|")
            If Regex.IsMatch(sArray(0), "^//") Then
                Continue For
            ElseIf Regex.IsMatch(sArray(0), "^//END") Then   '//ENDで読込終了
                Exit For
            End If
            DataGridView2.Rows.Add(sArray)
        Next

        Select Case True
            Case Regex.IsMatch(ComboBox1.SelectedItem, "日本郵便")
                hKaisya = "日本郵便"
            Case Regex.IsMatch(ComboBox1.SelectedItem, "e飛伝2")
                hKaisya = "e飛伝2"
            Case Regex.IsMatch(ComboBox1.SelectedItem, "佐川BIZ")
                hKaisya = "佐川BIZ"
        End Select

        CsvCheck()
    End Sub

    '変換処理
    Dim BT6_push As Boolean = False
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button6.Click
        If DataGridView2.Rows.Count <= 1 Then
            MsgBox("テンプレートが読み込まれていません", MsgBoxStyle.Exclamation Or MsgBoxStyle.SystemModal)
            Exit Sub
        ElseIf DataGridView1.Rows.Count = 0 Then
            MsgBox("処理をするファイルが読み込まれていません", MsgBoxStyle.Exclamation Or MsgBoxStyle.SystemModal)
            Exit Sub
        ElseIf BT6_push = True Then
            MsgBox("既に処理が行われています。", MsgBoxStyle.Exclamation Or MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        BT6_push = True

        Dim DGV1 As DataGridView = DataGridView1    'メイン
        Dim DGV4 As DataGridView = DataGridView4    '一時退避
        Dim DGV2 As DataGridView = DataGridView2    'テンプレート
        Dim henkanIraiTel As Integer = 0
        Dim hassouSakiMei As Integer = 0

        Dim ProgramList As New ArrayList    'プログラム行は別途処理するため保存

        'コラム番号に変換（変換先）
        For r As Integer = 0 To DGV2.RowCount - 1
            If Not DGV2.Rows(r).IsNewRow Then
                If DGV2.Item(0, r).Value = "プログラム" Then
                    ProgramList.Add(DGV2.Item(1, r).Value & splitter & DGV2.Item(2, r).Value)
                End If
            End If
        Next

        DGV4.Rows.Clear()
        DGV4.Columns.Clear()
        DGV4.Visible = True
        For c As Integer = 0 To DGV1.ColumnCount - 1
            DGV4.Columns.Add(c, c)
        Next
        For r As Integer = 0 To DGV1.RowCount - 1
            DGV4.Rows.Add(1)
            For c As Integer = 0 To DGV1.ColumnCount - 1
                DGV4.Item(c, r).Value = DGV1.Item(c, r).Value
            Next
        Next
        DGV1.Rows.Clear()
        DGV1.Columns.Clear()

        ToolStripProgressBar1.Value = 0
        ToolStripProgressBar1.Maximum = DGV1.RowCount
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------

        'ヘッダー追加
        For c As Integer = 0 To DGV2.Rows.Count - 1
            If Not DGV2.Rows(c).IsNewRow And DGV2.Item(0, c).Value <> "プログラム" Then
                DGV1.Columns.Add("dgv1" & c, c)
                If c = 0 Then
                    DGV1.Rows.Add(1)
                End If
                DGV1.Item(c, 0).Value = DGV2.Item(0, c).Value
            End If
        Next

        Dim dH1 As ArrayList = TM_HEADER_1ROW_GET(DGV1, 0)
        Dim dH4 As ArrayList = TM_HEADER_1ROW_GET(DGV4, 0)

        Dim henkanO As ArrayList = New ArrayList  '置換
        For r As Integer = 0 To DGV2.RowCount - 1
            If Not DGV2.Rows(r).IsNewRow Then
                henkanO.Add(DGV2.Item(3, r).Value)
            End If
        Next

        '伝票設定用ファイル読み取り
        Dim regStr As String = ""
        Dim regIrai As String = ""
        Dim dFlag As Integer = 0
        Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\config\denpyoConfig.txt"
        Using sr As New StreamReader(fName, System.Text.Encoding.Default)
            Do While Not sr.EndOfStream
                Dim s As String = sr.ReadLine
                If s = "[伝票依頼主用]" Then
                    dFlag = 1
                ElseIf s = "[依頼主変換]" Then
                    dFlag = 2
                ElseIf dFlag = 1 Then
                    regStr = s
                ElseIf dFlag = 2 Then
                    regIrai = s
                    Exit Do
                End If
            Loop
        End Using

        'コラム番号に変換（変換元）9999=空欄、10000=固定
        'ヘッダー+"-"+ヘッダー,空欄（"は'でもOK）
        '[固定],文字列
        '[規定],今日、明日、日付=20170220,フォーマット（省略時 yyyyMMdd）
        '[変換]ヘッダー,search=word|search=word
        '[検索],ヘッダー=search/ヘッダー
        '[検索],ヘッダー=search/[固定]=TrueWord|FalseWord
        '　　　↑search 文字、>0
        '[分割]ヘッダー,splitword=number,
        '　　　↑number=last最後
        '[合併]ヘッダー1,ヘッダー2|ヘッダー3（1つでもOK）
        '[選択],ヘッダー1|ヘッダー2|ヘッダー3,notWord
        '[計算],ヘッダー*0.08,floor or ceiling
        '[特殊]ヘッダー,[Qoo10日]
        '[特殊]ヘッダー,[Qoo10時]
        '[特殊]ヘッダー,[ヤフオク日時]
        '[特殊]ヘッダー,[住所]1or2
        '[特殊]ヘッダー,[発送備考欄個口]
        '[特殊]企業名,[依頼主]
        '[特殊]企業名,[電話]000-000-0000

        Dim koguchi As New ArrayList    'メール便用個口リスト
        For r As Integer = 1 To DGV4.RowCount - 1
            DGV1.Rows.Add(1)
            For c As Integer = 0 To DGV2.Rows.Count - 1
                If DGV2.Rows(c).IsNewRow Then
                    Continue For
                ElseIf DGV2.Item(0, c).Value = "プログラム" Then
                    Continue For
                End If

                Select Case True
                    '==================================================
                    Case Regex.IsMatch(DGV2.Item(1, c).Value, "\[規定\]")
                        Dim formatStr As String = DGV2.Item(3, c).Value
                        If formatStr = "" Then
                            formatStr = "yyyyMMdd"
                        End If
                        Select Case True
                            Case Regex.IsMatch(DGV2.Item(2, c).Value, "今日")
                                DGV1.Item(c, r).Value = Format(Today, formatStr)
                            Case Regex.IsMatch(DGV2.Item(2, c).Value, "明日")
                                DGV1.Item(c, r).Value = Format(DateAdd(DateInterval.Day, 1, Today), formatStr)
                            Case Regex.IsMatch(DGV2.Item(2, c).Value, "日付")
                                Dim dStr As String = Replace(DGV2.Item(2, c).Value, "日付=", "")
                                DGV1.Item(c, r).Value = Format(dStr, formatStr)
                        End Select
                    '==================================================
                    Case Regex.IsMatch(DGV2.Item(1, c).Value, "\[固定\]")
                        DGV1.Item(c, r).Value = DGV2.Item(2, c).Value
                    '==================================================
                    Case Regex.IsMatch(DGV2.Item(1, c).Value, "\[変換\]")
                        Dim henkanC As String = Replace(DGV2.Item(1, c).Value, "[変換]", "")
                        Dim sArray1 As String() = Split(DGV2.Item(2, c).Value, "|")
                        For i As Integer = 0 To sArray1.Length - 1
                            Dim sArray2 As String() = Split(sArray1(i), "=")
                            If sArray2(0) = "" Then
                                If DGV4.Item(TM_ArIndexof(dH4, henkanC), r).Value = "" Then
                                    DGV1.Item(c, r).Value = sArray2(1)
                                    If DGV1.Item(c, r).Value = "ヤマト運輸" Then
                                        DGV1.Item(c - 1, r).Value = "YAMATO"
                                    ElseIf DGV1.Item(c, r).Value = "佐川急便" Then
                                        DGV1.Item(c - 1, r).Value = "SAGAWA"
                                    End If
                                    Exit For
                                End If
                            Else
                                If InStr(DGV4.Item(TM_ArIndexof(dH4, henkanC), r).Value, sArray2(0)) > 0 Then
                                    DGV1.Item(c, r).Value = sArray2(1)

                                    If DGV1.Item(c, r).Value = "ヤマト運輸" Then
                                        DGV1.Item(c - 1, r).Value = "YAMATO"
                                    ElseIf DGV1.Item(c, r).Value = "佐川急便" Then
                                        DGV1.Item(c - 1, r).Value = "SAGAWA"
                                    End If
                                    Exit For
                                End If
                            End If




                        Next
                    '==================================================
                    Case Regex.IsMatch(DGV2.Item(1, c).Value, "\[分割\]")
                        Dim henkanC As String = Replace(DGV2.Item(1, c).Value, "[分割]", "")
                        Dim splitword As String = DGV2.Item(2, c).Value
                        splitword = Replace(splitword, "\s", " ")
                        Dim sArray1 As String() = Split(splitword, "=")
                        If DGV4.Item(TM_ArIndexof(dH4, henkanC), r).Value <> "" Then
                            Dim sArray2 As String() = Split(DGV4.Item(TM_ArIndexof(dH4, henkanC), r).Value, sArray1(0))
                            If sArray1(1) = "last" Then '最後のワードを取得
                                DGV1.Item(c, r).Value = sArray2(sArray2.Length - 1)
                            ElseIf InStr(sArray1(1), "*") > 0 Then
                                Dim str As String = ""
                                sArray1(1) = Replace(sArray1(1), "*", "")
                                For k As Integer = 0 To sArray2.Length - 1
                                    If k >= sArray1(1) - 1 Then
                                        If str = "" Then
                                            str = sArray2(k)
                                        Else
                                            str &= " " & sArray2(k)
                                        End If
                                    End If
                                Next
                                DGV1.Item(c, r).Value = str
                            ElseIf sArray2.Length >= sArray1(1) Then
                                DGV1.Item(c, r).Value = sArray2(sArray1(1) - 1)
                            Else
                                DGV1.Item(c, r).Value = ""
                            End If
                        Else
                            DGV1.Item(c, r).Value = ""
                        End If
                     '==================================================
                    Case Regex.IsMatch(DGV2.Item(1, c).Value, "\[検索\]")
                        Dim henkanC As String = DGV2.Item(2, c).Value
                        Dim sArray1 As String() = Split(henkanC, "/")       '合計金額=>0/008|   '商品番号=ad/True|False     '商品番号=ad/送料
                        Dim sArray2 As String() = Split(sArray1(0), "=")    '合計金額=>0        '商品番号=ad
                        If InStr(sArray2(1), ">") > 0 Then
                            Dim num As String = Replace(InStr(sArray2(1), ">"), ">", "")
                            If DGV4.Item(TM_ArIndexof(dH4, sArray2(0)), r).Value > num Then
                                Dim sArray4 As String() = Split(sArray1(1), "|")
                                DGV1.Item(c, r).Value = sArray4(0)
                            Else
                                Dim sArray4 As String() = Split(sArray1(1), "|")
                                DGV1.Item(c, r).Value = sArray4(1)
                            End If
                        ElseIf InStr(DGV4.Item(TM_ArIndexof(dH4, sArray2(0)), r).Value, sArray2(1)) > 0 Then   'TrueWord
                            If InStr(sArray1(1), "[固定]") <> 0 Then
                                Dim henkanD As String = Replace(DGV2.Item(2, c).Value, "[固定]=", "")
                                Dim sArray4 As String() = Split(henkanD, "|")
                                DGV1.Item(c, r).Value = sArray4(0)
                            Else
                                Dim sArray4 As String() = Split(sArray1(1), "|")
                                If InStr(sArray4(0), "+") > 0 Then                      '「+」で連結
                                    Dim s As String() = Split(sArray4(0), "+")
                                    Dim res As String = ""
                                    For i As Integer = 0 To s.Length - 1
                                        res &= DGV4.Item(TM_ArIndexof(dH4, s(i)), r).Value
                                    Next
                                    DGV1.Item(c, r).Value = res
                                Else
                                    DGV1.Item(c, r).Value = DGV4.Item(TM_ArIndexof(dH4, sArray4(0)), r).Value
                                End If
                            End If
                        Else                                                                 'FalseWord
                            If InStr(sArray1(1), "[固定]") <> 0 Then
                                Dim henkanD As String = Replace(DGV2.Item(2, c).Value, "[固定]=", "")
                                Dim sArray4 As String() = Split(henkanD, "|")
                                DGV1.Item(c, r).Value = sArray4(1)
                            Else
                                Dim sArray4 As String() = Split(sArray1(1), "|")
                                If InStr(sArray4(1), "+") > 0 Then                      '「+」で連結
                                    Dim s As String() = Split(sArray4(1), "+")
                                    Dim res As String = ""
                                    For i As Integer = 0 To s.Length - 1
                                        res &= DGV4.Item(TM_ArIndexof(dH4, s(i)), r).Value
                                    Next
                                    DGV1.Item(c, r).Value = res
                                Else
                                    DGV1.Item(c, r).Value = DGV4.Item(TM_ArIndexof(dH4, sArray4(1)), r).Value
                                End If
                            End If
                        End If
                     '==================================================
                    Case Regex.IsMatch(DGV2.Item(1, c).Value, "\[選択\]")
                        Dim sentakuC As String = DGV2.Item(2, c).Value
                        If InStr(sentakuC, "|") > 0 Then
                            Dim arrayA As String() = Split(sentakuC, "|")
                            Dim res As String = ""
                            For i As Integer = 0 To arrayA.Length - 1
                                If DGV4.Item(TM_ArIndexof(dH4, arrayA(i)), r).Value <> DGV2.Item(3, c).Value Then
                                    res = DGV4.Item(TM_ArIndexof(dH4, arrayA(i)), r).Value
                                    Exit For
                                ElseIf DGV4.Item(TM_ArIndexof(dH4, arrayA(i)), r).Value <> "" Then
                                    res = DGV4.Item(TM_ArIndexof(dH4, arrayA(i)), r).Value
                                    Exit For
                                End If
                            Next
                            DGV1.Item(c, r).Value = res
                        Else
                            If DGV4.Item(TM_ArIndexof(dH4, sentakuC), r).Value <> DGV2.Item(3, c).Value Then
                                DGV1.Item(c, r).Value = DGV4.Item(TM_ArIndexof(dH4, sentakuC), r).Value
                            End If
                        End If
                     '==================================================
                    Case Regex.IsMatch(DGV2.Item(1, c).Value, "\[計算\]")
                        Dim keisan As String = DGV2.Item(2, c).Value
                        Dim parts As String() = Regex.Split(keisan, "(?<1>\*|\/|\-|\+)")
                        Dim exp As String = ""
                        For i As Integer = 0 To parts.Length - 1
                            If Not IsNumeric(parts(i)) And Not Regex.IsMatch(parts(i), "[\*|\/|\-|\+]{1,1}") Then
                                exp &= DGV4.Item(TM_ArIndexof(dH4, parts(i)), r).Value
                            Else
                                exp &= parts(i)
                            End If
                        Next
                        '文字列を計算する
                        Dim dt As New DataTable()
                        Dim result As Double = CDbl(dt.Compute(exp, ""))
                        If DGV2.Item(3, c).Value = "floor" Then
                            result = Math.Floor(result)
                        ElseIf DGV2.Item(3, c).Value = "ceiling" Then
                            result = Math.Ceiling(result)
                        End If
                        DGV1.Item(c, r).Value = result
                    '==================================================
                    Case Regex.IsMatch(DGV2.Item(1, c).Value, "\[特殊\]")
                        If InStr(DGV2.Item(2, c).Value, "Qoo10") > 0 Then
                            Dim tokuC As String = Replace(DGV2.Item(1, c).Value, "[特殊]", "")
                            Dim colNum As Integer = TM_ArIndexof(dH4, tokuC)
                            If InStr(ComboBox1.SelectedItem, "佐川") > 0 Then
                                If Regex.IsMatch(DGV4.Item(colNum, r).Value, "-") And Regex.IsMatch(DGV4.Item(colNum, r).Value, ":") Then
                                    Dim arrayA As String() = Split(DGV4.Item(colNum, r).Value, " ")
                                    If DGV2.Item(2, c).Value = "[Qoo10日]" Then
                                        Dim arrayB As String() = Split(arrayA(0), "-")
                                        DGV1.Item(c, r).Value = arrayB(1) & "月" & arrayB(2) & "日"
                                    ElseIf DGV2.Item(2, c).Value = "[Qoo10時]" Then
                                        DGV1.Item(c, r).Value = Qoo10time(arrayA(1))
                                    End If
                                ElseIf Regex.IsMatch(DGV4.Item(colNum, r).Value, "-") And Not Regex.IsMatch(DGV4.Item(colNum, r).Value, ":") Then
                                    If DGV2.Item(2, c).Value = "[Qoo10日]" Then
                                        Dim arrayB As String() = Split(DGV4.Item(colNum, r).Value, "-")
                                        DGV1.Item(c, r).Value = arrayB(1) & "月" & arrayB(2) & "日"
                                    End If
                                ElseIf Not Regex.IsMatch(DGV4.Item(colNum, r).Value, "-") And Regex.IsMatch(DGV4.Item(colNum, r).Value, ":") Then
                                    If DGV2.Item(2, c).Value = "[Qoo10時]" Then
                                        DGV1.Item(c, r).Value = Qoo10time(DGV4.Item(colNum, r).Value)
                                    End If
                                End If
                            Else
                                If Regex.IsMatch(DGV4.Item(colNum, r).Value, "-") And Regex.IsMatch(DGV4.Item(colNum, r).Value, ":") Then
                                    Dim arrayA As String() = Split(DGV4.Item(colNum, r).Value, " ")
                                    If DGV2.Item(2, c).Value = "[Qoo10日]" Then
                                        DGV1.Item(c, r).Value = Replace(arrayA(0), "-", "")
                                    ElseIf DGV2.Item(2, c).Value = "[Qoo10時]" Then
                                        DGV1.Item(c, r).Value = Qoo10time(arrayA(1))
                                    End If
                                ElseIf Regex.IsMatch(DGV4.Item(colNum, r).Value, "-") And Not Regex.IsMatch(DGV4.Item(colNum, r).Value, ":") Then
                                    If DGV2.Item(2, c).Value = "[Qoo10日]" Then
                                        DGV1.Item(c, r).Value = Replace(DGV4.Item(colNum, r).Value, "-", "")
                                    End If
                                ElseIf Not Regex.IsMatch(DGV4.Item(colNum, r).Value, "-") And Regex.IsMatch(DGV4.Item(colNum, r).Value, ":") Then
                                    If DGV2.Item(2, c).Value = "[Qoo10時]" Then
                                        DGV1.Item(c, r).Value = Qoo10time(DGV4.Item(colNum, r).Value)
                                    End If
                                End If
                            End If
                        ElseIf Regex.IsMatch(DGV2.Item(2, c).Value, "\[住所\]") Then
                            Dim tokuC As String = Replace(DGV2.Item(1, c).Value, "[特殊]", "")
                            Dim mode As String = Replace(DGV2.Item(2, c).Value, "[住所]", "")
                            Dim address As String = DGV4.Item(TM_ArIndexof(dH4, tokuC), r).Value

                            Dim addressCheck As Boolean = False
                            Dim work As String
                            Dim i As Integer
                            For i = 0 To address.Length - 1 '住所の文字数分繰り返す
                                work = address.Substring(i, 1)
                                If IsNumeric(work) Then
                                    addressCheck = True
                                    Exit For
                                End If
                            Next i
                            If mode = 1 Then
                                DGV1.Item(c, r).Value = address.Substring(0, i)

                                If addressCheck = False Then    '番地未入力チェック
                                    DGV1.Item(c, r).Style.BackColor = Color.Yellow
                                    'For cc As Integer = 0 To dgv1.ColumnCount - 1
                                    '    dgv1.Item(cc, r).Style.BackColor = Color.Yellow
                                    'Next
                                End If
                            ElseIf mode = 2 Then
                                DGV1.Item(c, r).Value = address.Substring(i)
                            End If
                        ElseIf Regex.IsMatch(DGV2.Item(2, c).Value, "\[個口\]") Then '受注者電話番号の後ろに「,2」とし、複数個口を自動処理する
                            Dim tokuC As String = Replace(DGV2.Item(1, c).Value, "[特殊]", "")
                            Dim denwa As String = DGV4.Item(TM_ArIndexof(dH4, tokuC), r).Value

                            If InStr(denwa, splitter) > 0 Then
                                Dim kosuu As String() = Split(denwa, splitter)
                                If IsNumeric(kosuu(1)) Then '個口が数字で、1個以上の場合のみ
                                    If kosuu(1) > 1 Then
                                        DGV1.Item(c, r).Value = kosuu(1)
                                        koguchi.Add(r & splitter & kosuu(1))
                                    End If
                                End If
                            Else
                                DGV1.Item(c, r).Value = 1
                                koguchi.Add(r & splitter & 1)
                            End If
                        ElseIf Regex.IsMatch(DGV2.Item(2, c).Value, "\[依頼主\]") Then
                            Dim tokuC As String = Replace(DGV2.Item(1, c).Value, "[特殊]", "")
                            Dim kigyoumei As String = Replace(DGV4.Item(TM_ArIndexof(dH4, tokuC), r).Value, " ", "")
                            kigyoumei = Replace(kigyoumei, "　", "")

                            Select Case True
                                Case Regex.IsMatch(kigyoumei, regStr)   '伝票設定用ファイルから設定読み取り
                                    DGV1.Item(c, r).Value = ""
                                Case Else
                                    Dim todokeSaki As String = Replace(DGV4.Item(hassouSakiMei, r).Value, " ", "")
                                    todokeSaki = Replace(todokeSaki, "　", "")
                                    If todokeSaki = kigyoumei Then
                                        DGV1.Item(c, r).Value = ""  '企業名と届け先が同じなら変更しない
                                    Else
                                        '依頼主変換のチェック
                                        Dim iraiArray As String() = Split(regIrai, "|")
                                        Dim iraiFlag As Boolean = False
                                        Dim iraiHenkan As String() = Nothing
                                        For i As Integer = 0 To iraiArray.Length - 1
                                            Dim namaeArray As String() = Split(iraiArray(i), splitter)
                                            If kigyoumei = namaeArray(0) Then
                                                iraiHenkan = namaeArray
                                                iraiFlag = True
                                                Exit For
                                            End If
                                        Next

                                        Select Case hKaisya
                                            Case "日本郵便"
                                                If iraiFlag = False Then
                                                    DGV1.Item(c, r).Value = "(依頼主)[" & DGV4.Item(henkanIraiTel, r).Value & "]" & kigyoumei & "様"
                                                Else
                                                    DGV1.Item(c, r).Value = "(依頼主)[" & iraiHenkan(2) & "]" & iraiHenkan(1) & "様"
                                                End If
                                            Case Else
                                                If iraiFlag = False Then
                                                    DGV1.Item(c, r).Value = "(依頼)" & kigyoumei & "様"    ' & dgv4.Item(henkanIraiTel, r).Value
                                                Else
                                                    DGV1.Item(c, r).Value = "(依頼)" & iraiHenkan(1) & "様" ' & iraiHenkan(2)
                                                End If
                                        End Select
                                    End If
                            End Select
                        ElseIf Regex.IsMatch(DGV2.Item(2, c).Value, "\[電話\]") Then
                            Dim tokuC As String = Replace(DGV2.Item(1, c).Value, "[特殊]", "")
                            Dim kigyoumei As String = Replace(DGV4.Item(TM_ArIndexof(dH4, tokuC), r).Value, " ", "")
                            kigyoumei = Replace(kigyoumei, "　", "")

                            Select Case True
                                Case Regex.IsMatch(kigyoumei, regStr)   '伝票設定用ファイルから設定読み取り
                                    Dim res As String = Replace(DGV2.Item(2, c).Value, "[電話]", "")
                                    DGV1.Item(c, r).Value = res
                                Case Else
                                    Dim todokeSaki As String = Replace(DGV4.Item(hassouSakiMei, r).Value, " ", "")
                                    todokeSaki = Replace(todokeSaki, "　", "")
                                    If todokeSaki = kigyoumei Then
                                        Dim res As String = Replace(DGV2.Item(2, c).Value, "[電話]", "")
                                        DGV1.Item(c, r).Value = res
                                    Else
                                        DGV1.Item(c, r).Value = ""
                                    End If
                            End Select
                        ElseIf Regex.IsMatch(DGV2.Item(2, c).Value, "\[ヤフオク日時\]") Then
                            Dim tokuC As String = Replace(DGV2.Item(1, c).Value, "[特殊]", "")
                            tokuC = DGV4.Item(TM_ArIndexof(dH4, tokuC), r).Value
                            DGV1.Item(c, r).Value = Format(CDate(tokuC), "yyyy/MM/dd HH:mm:ss")
                        End If
                    '==================================================
                    Case Regex.IsMatch(DGV2.Item(1, c).Value, "\[テンプレ\]")
                        Dim files As String = Path.GetDirectoryName(Form1.appPath) & "\template\" & DGV2.Item(2, c).Value & ".html"
                        If File.Exists(files) Then
                            Dim lines As String = File.ReadAllText(files)
                            DGV1.Item(c, r).Value = lines
                        Else
                            DGV1.Item(c, r).Value = "Template Error"
                        End If
                        '==================================================
                    Case Else
                        If InStr(DGV2.Item(1, c).Value, "+") > 0 Then   '+で連結
                            Dim arrayA As String() = Split(DGV2.Item(1, c).Value, "+")
                            Dim res As String = ""
                            For i As Integer = 0 To arrayA.Length - 1
                                If dH4.Contains(arrayA(i)) Then
                                    Dim resA As String = DGV4.Item(TM_ArIndexof(dH4, arrayA(i)), r).Value
                                    resA = Regex.Replace(resA, Chr(34) & "|" & Chr(39), "")
                                    res &= resA
                                Else
                                    res &= arrayA(i)
                                End If
                            Next
                            DGV1.Item(c, r).Value = res
                        Else
                            Dim header As String = DGV2.Item(1, c).Value
                            If header <> "" Then
                                DGV1.Item(c, r).Value = DGV4.Item(TM_ArIndexof(dH4, header), r).Value
                            End If
                        End If
                End Select

                '文字列の置換
                If henkanO(c) <> "" Then
                    Dim rp1 As String() = Split(henkanO(c), "|")
                    For i As Integer = 0 To rp1.Length - 1
                        If InStr(rp1(i), "=") > 0 Then
                            Dim rp2 As String() = Split(rp1(i), "=")
                            DGV1.Item(c, r).Value = Replace(DGV1.Item(c, r).Value, rp2(0), rp2(1))
                        End If
                    Next
                End If
            Next
        Next
        DGV4.Visible = False

        '最後にプログラム処理を行なう
        'DGVheaderCheck1.Clear()
        'For c As Integer = 0 To DGV1.Columns.Count - 1
        '    DGVheaderCheck1.Add(DGV1.Item(c, 0).Value)
        'Next c
        dH1 = TM_HEADER_1ROW_GET(DGV1, 0)
        For i As Integer = 0 To ProgramList.Count - 1
            If Regex.IsMatch(ProgramList(i), "^伝票チェック") Then
                Dim PL As String() = Split(ProgramList(i), splitter)
                Dim Bangou1 As String = ""
                Dim Bangou2 As String = ""
                Dim Bangou3 As String = ""
                If PL(1) = "楽天" Then
                    Bangou1 = "受注番号"
                    Bangou2 = "お荷物伝票番号"
                ElseIf PL(1) = "楽天ペイ" Then
                    Bangou1 = "注文番号"
                    Bangou2 = "お荷物伝票番号"
                ElseIf PL(1) = "Yahoo" Then
                    Bangou1 = "OrderId"
                    Bangou2 = "ShipInvoiceNumber1"
                    'Bangou3 = "ShipInvoiceNumber2"
                ElseIf PL(1) = "MakeShop" Then
                    Bangou1 = "注文番号"
                    Bangou2 = "伝票番号"
                ElseIf PL(1) = "Amazon" Then
                    Bangou1 = "order-id"
                    Bangou2 = "tracking-number"
                ElseIf PL(1) = "Wowma" Then
                    Bangou1 = "orderId"
                    Bangou2 = "shippingNumber"
                End If

                Dim allArray As ArrayList = New ArrayList       '伝票番号
                Dim poolArray As ArrayList = New ArrayList      '伝票番号
                Dim removeArray As ArrayList = New ArrayList    '行
                For r As Integer = DGV1.RowCount - 1 To 1 Step -1
                    Dim DNum As String = DGV1.Item(dH1.IndexOf(Bangou1), r).Value

                    If InStr(DNum, "-fk-") > 0 Then '複写伝票は削除
                        Dim DNum1 As String() = Regex.Split(DNum, "-fk-")
                        allArray.Add(DNum1(0))
                        removeArray.Add(r)
                    ElseIf InStr(DNum, "-bk-") > 0 Then
                        '分割フラグ（-bk）が付いている受注番号が他にあれば削除対象にする
                        Dim DNum1 As String() = Regex.Split(DNum, "-bk-|-h")
                        allArray.Add(DNum1(0))

                        If Not poolArray.Contains(DNum1(0)) Then
                            DGV1.Item(dH1.IndexOf(Bangou1), r).Value = DNum1(0) '分割番号を正しい番号に直す
                            poolArray.Add(DNum1(0))
                        Else
                            Dim motoGyou As Integer = allArray.IndexOf(DNum1(0)) + 1
                            Select Case PL(1)
                                Case "楽天"
                                    DGV1.Item(dH1.IndexOf(Bangou2), motoGyou).Value &= splitter & DGV1.Item(dH1.IndexOf(Bangou2), r).Value
                                Case "楽天ペイ"
                                    DGV1.Rows.Add()
                                    For c As Integer = 0 To DGV1.ColumnCount - 1
                                        DGV1.Item(dH1.IndexOf(Bangou2), motoGyou + 1).Value = DGV1.Item(dH1.IndexOf(Bangou2), motoGyou).Value
                                    Next
                                Case "Yahoo"
                                    If DGV1.Item(dH1.IndexOf(Bangou2), motoGyou).Value = "" Then   'ShipInvoiceNumber1が空いている時
                                        DGV1.Item(dH1.IndexOf(Bangou2), motoGyou).Value = DGV1.Item(dH1.IndexOf(Bangou2), r).Value
                                        'ElseIf DGV1.Item(DGVheaderCheck1.IndexOf(Bangou3), motoGyou).Value = "" Then   'ShipInvoiceNumber2が空いている時
                                        '    DGV1.Item(DGVheaderCheck1.IndexOf(Bangou3), motoGyou).Value = DGV1.Item(DGVheaderCheck1.IndexOf(Bangou2), r).Value
                                    Else
                                        Dim a = 0
                                    End If
                                Case "MakeShop", "Amazon"
                                    '分割伝票の分は削除する
                            End Select
                            removeArray.Add(r)
                        End If
                    Else
                        '各モールの規定文字列ではない受注番号の場合削除対象にする
                        Dim kiteiFalseFlag As Boolean = True
                        Select Case PL(1)
                            Case "楽天", "楽天ペイ"
                                If Not Regex.IsMatch(DNum, "[0-9]{6}-[0-9]{8}-[0-9]{8}") Then
                                    kiteiFalseFlag = False
                                End If
                            Case "Yahoo"
                                If Not Regex.IsMatch(DNum, "[0-9]{8}") Then
                                    kiteiFalseFlag = False
                                End If
                            Case "MakeShop"
                                If Not Regex.IsMatch(DNum, "P[0-9]{18}") Then
                                    kiteiFalseFlag = False
                                End If
                            Case "Amazon"
                                If Not Regex.IsMatch(DNum, "[0-9]{3}-[0-9]{7}-[0-9]{7}") Then
                                    kiteiFalseFlag = False
                                End If
                            Case "Wowma"
                                If Not Regex.IsMatch(DNum, "[0-9]{6,9}") Then
                                    kiteiFalseFlag = False
                                End If
                        End Select

                        If kiteiFalseFlag = True Then
                            If Not poolArray.Contains(DNum) Then
                                allArray.Add(DNum)
                                poolArray.Add(DNum)
                            Else
                                Dim motoGyou As Integer = allArray.IndexOf(DNum) + 1
                                Select Case PL(1)
                                    Case "楽天", "楽天ペイ"
                                        DGV1.Item(dH1.IndexOf(Bangou2), motoGyou).Value &= splitter & DGV1.Item(dH1.IndexOf(Bangou2), r).Value
                                    Case "Yahoo"
                                        If DGV1.Item(dH1.IndexOf(Bangou2), motoGyou).Value = "" Then   'ShipInvoiceNumber1が空いている時
                                            DGV1.Item(dH1.IndexOf(Bangou2), motoGyou).Value = DGV1.Item(dH1.IndexOf(Bangou2), r).Value
                                            'ElseIf DGV1.Item(DGVheaderCheck1.IndexOf(Bangou3), motoGyou).Value = "" Then   'ShipInvoiceNumber2が空いている時
                                            '    DGV1.Item(DGVheaderCheck1.IndexOf(Bangou3), motoGyou).Value = DGV1.Item(DGVheaderCheck1.IndexOf(Bangou2), r).Value
                                        Else
                                            Dim a = 0
                                        End If
                                    Case "MakeShop", "Amazon", "Wowma"
                                        '分割伝票の分は削除する
                                End Select
                                removeArray.Add(r)
                            End If
                        Else
                            removeArray.Add(r)
                        End If

                    End If

                    Select Case PL(1)
                        Case "楽天", "楽天ペイ"
                            If DGV1.Item(dH1.IndexOf("配送会社"), r).Value = "1003" And DGV1.Item(dH1.IndexOf(Bangou2), r).Value = "" Then
                                DGV1.Item(dH1.IndexOf(Bangou2), r).Value = "定形外郵便発送(追跡番号無し)"
                            End If
                            If PL(1) = "楽天ペイ" Then
                                If InStr(DGV1.Item(dH1.IndexOf("お荷物伝票番号"), r).Value, ",") > 0 Then
                                    Dim dNo As String() = Split(DGV1.Item(dH1.IndexOf("お荷物伝票番号"), r).Value, ",")
                                    For d As Integer = 0 To dNo.Length - 1
                                        If d <> dNo.Length - 1 Then
                                            DGV1.Rows.Insert(r, 1)
                                        End If
                                        For c As Integer = 0 To DGV1.ColumnCount - 1
                                            If c = dH1.IndexOf("お荷物伝票番号") Then
                                                DGV1.Item(c, r + d).Value = dNo(d)
                                            Else
                                                DGV1.Item(c, r + d).Value = DGV1.Item(c, r + 1).Value
                                            End If
                                        Next
                                    Next
                                End If
                            End If
                        Case "Amazon"
                            If DGV1.Item(dH1.IndexOf("carrier-name"), r).Value = "日本郵便" And DGV1.Item(dH1.IndexOf(Bangou2), r).Value = "" Then
                                DGV1.Item(dH1.IndexOf(Bangou2), r).Value = Format(Today(), "yyyyMMdd")
                            End If
                    End Select
                Next
                removeArray.Sort()
                For k As Integer = removeArray.Count - 1 To 0 Step -1
                    DGV1.Rows.RemoveAt(removeArray(k))
                Next
            ElseIf Regex.IsMatch(ProgramList(i), "^ヤフオク画像") Then
                Dim imgFolder As String = Path.GetDirectoryName(ToolStripLabel2.Text)
                Dim di As New DirectoryInfo(imgFolder)
                Dim files As FileInfo() = di.GetFiles("*", SearchOption.AllDirectories)
                'ファイル名をArrayListに変換しソートする
                Dim filesArray As New ArrayList
                For Each f As FileInfo In files
                    filesArray.Add(f.FullName)
                Next
                filesArray.Sort()
                '画像1～10に管理番号と合致するファイルを探して書き込む
                Dim usedFileArray As New ArrayList
                Dim headerList As String() = New String() {"画像1", "画像2", "画像3", "画像4", "画像5", "画像6", "画像7", "画像8", "画像9", "画像10"}
                Dim num As Integer = 0
                For r As Integer = 1 To DGV1.RowCount - 1
                    num = 0
                    Dim code As String = DGV1.Item(TM_ArIndexof(dH1, "管理番号"), r).Value
                    For Each f As String In filesArray
                        If Regex.IsMatch(Path.GetFileNameWithoutExtension(f), "^" & code) Then
                            If Not usedFileArray.Contains(Path.GetFileNameWithoutExtension(f)) Then
                                Dim hNum As Integer = TM_ArIndexof(dH1, headerList(num))
                                DGV1.Item(hNum, r).Value = Path.GetFileName(f)
                                usedFileArray.Add(Path.GetFileNameWithoutExtension(f))
                                num += 1
                            End If
                        End If
                    Next
                Next
                '説明文の画像を変更する
                For r As Integer = 1 To DGV1.RowCount - 1
                    For k As Integer = 0 To headerList.Length - 1
                        Dim imgStr As String = "<img src=https://shopping.c.yimg.jp/lib/lucky9/_AA_ width=100%><br>" & vbCrLf
                        Dim koumoku As String = DGV1.Item(TM_ArIndexof(dH1, headerList(k)), r).Value
                        If koumoku <> "" Then
                            imgStr = Replace(imgStr, "_AA_", koumoku)
                            DGV1.Item(TM_ArIndexof(dH1, "説明"), r).Value = Replace(DGV1.Item(TM_ArIndexof(dH1, "説明"), r).Value, "<!--IMG-->", imgStr & "<!--IMG-->")
                        End If
                        If k = 6 Then   '7枚まで
                            Exit For
                        End If
                    Next
                    DGV1.Item(TM_ArIndexof(dH1, "説明"), r).Value = Replace(DGV1.Item(TM_ArIndexof(dH1, "説明"), r).Value, "<!--IMG-->", "")
                Next
                '説明を移行
                For r As Integer = 1 To DGV1.RowCount - 1
                    Dim str As String = DGV4.Item(TM_ArIndexof(dH4, "explanation"), r).Value
                    str = Regex.Replace(str, "\r|\n|\r\n", "<br>" & vbCrLf)
                    DGV1.Item(TM_ArIndexof(dH1, "説明"), r).Value = Replace(DGV1.Item(TM_ArIndexof(dH1, "説明"), r).Value, "<!--CON-->", str & vbCrLf)
                Next
            End If
        Next

        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------

        ToolStripProgressBar1.Value = 0

    End Sub

    '変換処理
    'Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
    '    Dim DV1 As DataGridView = dgv
    '    Dim DV2 As DataGridView = DataGridView2
    '    Dim DV3 As DataGridView = DataGridView4
    '    Dim henkanSHeader As ArrayList = New ArrayList
    '    Dim henkanSNo As ArrayList = New ArrayList  '変換先
    '    Dim henkanMNo As ArrayList = New ArrayList  '変換元
    '    Dim henkanIraiTel As Integer = 0
    '    Dim hassouSakiMei As Integer = 0

    '    Dim ProgramList As ArrayList = New ArrayList    'プログラム行は別途処理するため保存

    '    'コラム番号に変換（変換先）
    '    If CheckBox1.Checked = True Then
    '        For r As Integer = 0 To DV2.RowCount - 2
    '            If DV2.Item(0, r).Value = "プログラム" Then
    '                Continue For
    '            End If
    '            henkanSHeader.Add(DV2.Item(0, r).Value)
    '            henkanSNo.Add(r)
    '        Next
    '    Else
    '        For r As Integer = 0 To DV2.RowCount - 2
    '            If DV2.Item(0, r).Value = "プログラム" Then
    '                Continue For
    '            End If
    '            henkanSHeader.Add(r)
    '            henkanSNo.Add(r)
    '        Next
    '    End If

    '    Dim henkanO As ArrayList = New ArrayList  '置換
    '    For r As Integer = 0 To DV2.RowCount - 2
    '        If DV2.Item(0, r).Value = "プログラム" Then
    '            Continue For
    '        End If
    '        henkanO.Add(DV2.Item(3, r).Value)
    '    Next

    '    If CheckBox2.Checked = True Then
    '        If InStr(ComboBox1.SelectedItem, "ゆパック元") > 0 Or InStr(ComboBox1.SelectedItem, "ゆパック代引") > 0 Then
    '            If InStr(ComboBox1.SelectedItem, "ゆパック元") > 0 Then
    '                Dim DR As DialogResult = MsgBox("ゆうパック元払いではない「代引き」の伝票は削除して処理をします", MsgBoxStyle.OkCancel Or MsgBoxStyle.SystemModal)
    '                If DR = Windows.Forms.DialogResult.Cancel Then
    '                    Exit Sub
    '                End If
    '            ElseIf InStr(ComboBox1.SelectedItem, "ゆパック代引") > 0 Then
    '                Dim DR As DialogResult = MsgBox("ゆうパック代引きではない「元払い」の伝票は削除して処理をします", MsgBoxStyle.OkCancel Or MsgBoxStyle.SystemModal)
    '                If DR = Windows.Forms.DialogResult.Cancel Then
    '                    Exit Sub
    '                End If
    '            End If
    '            Dim goukeiC As Integer = -1
    '            For c As Integer = 0 To DV1.ColumnCount - 1
    '                If DV1.Item(c, 0).Value = "合計金額" Then
    '                    goukeiC = c
    '                End If
    '            Next
    '            If goukeiC = -1 Then
    '                MsgBox("合計金額の列が見つかりません。", MsgBoxStyle.SystemModal)
    '                Exit Sub
    '            End If
    '            Dim maxRow As Integer = DV1.RowCount - 1
    '            For r As Integer = 0 To maxRow
    '                Dim delRow As Integer = maxRow - r
    '                If delRow <= 0 Then
    '                    Exit For
    '                End If
    '                If DV1.Item(goukeiC, delRow).Value <> "0" And InStr(ComboBox1.SelectedItem, "ゆパック元") > 0 Then
    '                    DV1.Rows.RemoveAt(delRow)
    '                ElseIf DV1.Item(goukeiC, delRow).Value = "0" And InStr(ComboBox1.SelectedItem, "ゆパック代引") > 0 Then
    '                    DV1.Rows.RemoveAt(delRow)
    '                End If
    '            Next
    '        End If
    '    End If

    '    'Try
    '    'コラム番号に変換（変換元）9999=空欄、10000=固定
    '    'ヘッダー+"-"+ヘッダー,空欄（"は'でもOK）
    '    '[固定],文字列
    '    '[規定],今日、明日、日付=20170220,フォーマット（省略時 yyyyMMdd）
    '    '[変換]ヘッダー,search=word|search=word
    '    '[検索],ヘッダー=search/ヘッダー
    '    '[検索],ヘッダー=search/[固定]=TrueWord|FalseWord
    '    '　　　↑search 文字、>0
    '    '[分割]ヘッダー,splitword=number,
    '    '　　　↑number=last最後
    '    '[合併]ヘッダー1,ヘッダー2|ヘッダー3（1つでもOK）
    '    '[選択],ヘッダー1|ヘッダー2|ヘッダー3,notWord
    '    '[特殊]ヘッダー,[Qoo10日]
    '    '[特殊]ヘッダー,[Qoo10時]
    '    '[特殊]ヘッダー,[住所]1or2
    '    '[特殊]ヘッダー,[発送備考欄個口]
    '    '[特殊]企業名,[依頼主]
    '    '[特殊]企業名,[電話]000-000-0000
    '    For r As Integer = 0 To DV2.RowCount - 2
    '        If DV2.Item(0, r).Value = "プログラム" Then
    '            ProgramList.Add(DV2.Item(1, r).Value & splitter & DV2.Item(2, r).Value)
    '            Continue For
    '        End If

    '        Dim flag As Boolean = False
    '        Dim nCol As String() = Nothing
    '        If DV2.Item(1, r).Value = "[固定]" Then
    '            henkanMNo.Add(10000)
    '        ElseIf DV2.Item(1, r).Value = "[規定]" Then
    '            henkanMNo.Add(10002)
    '        ElseIf InStr(DV2.Item(1, r).Value, "[特殊]") > 0 Then
    '            If RadioButton16.Checked = True Then
    '                Dim hStr As String = Replace(DV2.Item(1, r).Value, "[特殊]", "")
    '                For c As Integer = 0 To DV1.ColumnCount - 1
    '                    If DV1.Item(c, 0).Value = hStr Then
    '                        henkanMNo.Add("[特殊]" & c)
    '                        flag = True
    '                        Exit For
    '                    End If
    '                Next
    '                If flag = False Then
    '                    henkanMNo.Add(9999)
    '                End If
    '                '依頼主の電話番号列を取得する
    '                For c As Integer = 0 To DV1.ColumnCount - 1
    '                    Select Case DV1.Item(c, 0).Value
    '                        Case "企業電話番号", "代行ご依頼主電話"     'ネクストエンジンのテンプレ変えたらここも
    '                            henkanIraiTel = c
    '                        Case "発送先名", "お届け先名"   '確認用
    '                            hassouSakiMei = c
    '                    End Select
    '                Next
    '            Else
    '                henkanMNo.Add(DV2.Item(1, r).Value)
    '            End If
    '        ElseIf InStr(DV2.Item(1, r).Value, "[変換]") > 0 Then
    '            If RadioButton16.Checked = True Then
    '                Dim hStr As String = Replace(DV2.Item(1, r).Value, "[変換]", "")
    '                For c As Integer = 0 To DV1.ColumnCount - 1
    '                    If DV1.Item(c, 0).Value = hStr Then
    '                        henkanMNo.Add("[変換]" & c)
    '                        flag = True
    '                        Exit For
    '                    End If
    '                Next
    '                If flag = False Then
    '                    henkanMNo.Add(9999)
    '                End If
    '            Else
    '                henkanMNo.Add(DV2.Item(1, r).Value)
    '            End If
    '        ElseIf InStr(DV2.Item(1, r).Value, "[分割]") > 0 Then
    '            If RadioButton16.Checked = True Then
    '                Dim hStr As String = Replace(DV2.Item(1, r).Value, "[分割]", "")
    '                For c As Integer = 0 To DV1.ColumnCount - 1
    '                    If DV1.Item(c, 0).Value = hStr Then
    '                        henkanMNo.Add("[分割]" & c)
    '                        flag = True
    '                        Exit For
    '                    End If
    '                Next
    '                If flag = False Then
    '                    henkanMNo.Add(10003)
    '                End If
    '            Else
    '                henkanMNo.Add(DV2.Item(1, r).Value)
    '            End If
    '        ElseIf InStr(DV2.Item(1, r).Value, "[合併]") > 0 Then
    '            If RadioButton16.Checked = True Then
    '                Dim res As String = ""
    '                Dim hStr As String = Replace(DV2.Item(1, r).Value, "[合併]", "")
    '                For c As Integer = 0 To DV1.ColumnCount - 1
    '                    If DV1.Item(c, 0).Value = hStr Then
    '                        res = c
    '                        flag = True
    '                        Exit For
    '                    End If
    '                Next
    '                Dim gStr As String() = Split(DV2.Item(2, r).Value, "|")
    '                For i As Integer = 0 To gStr.Length - 1
    '                    For c As Integer = 0 To DV1.ColumnCount - 1
    '                        If DV1.Item(c, 0).Value = gStr(i) Then
    '                            res &= "|" & c
    '                            Exit For
    '                        End If
    '                    Next
    '                Next
    '                If flag = False Then
    '                    henkanMNo.Add(9999)
    '                Else
    '                    henkanMNo.Add("[合併]" & res)
    '                End If
    '            Else
    '                henkanMNo.Add("[合併]" & DV2.Item(1, r).Value & "|" & DV2.Item(2, r).Value)
    '            End If
    '        ElseIf InStr(DV2.Item(1, r).Value, "[検索]") > 0 Then
    '            If RadioButton16.Checked = True Then
    '                Dim str As String = "[検索]"
    '                Dim sArray1 As String() = Split(DV2.Item(2, r).Value, "/")
    '                Dim sArray2 As String() = Split(sArray1(0), "=")
    '                For c As Integer = 0 To DV1.ColumnCount - 1
    '                    If InStr(DV2.Item(c, 0).Value, ">0") > 0 Then
    '                        str &= c & "=" & sArray2(1)
    '                        Exit For
    '                    Else
    '                        If DV1.Item(c, 0).Value = sArray2(0) Then
    '                            str &= c & "=" & sArray2(1)
    '                            Exit For
    '                        End If
    '                    End If
    '                Next
    '                If InStr(sArray1(1), "[固定]") = 0 Then
    '                    Dim hed2 As String = sArray1(1)
    '                    For c As Integer = 0 To DV1.ColumnCount - 1
    '                        If DV1.Item(c, 0).Value = hed2 Then
    '                            str &= "/" & c
    '                            Exit For
    '                        End If
    '                    Next
    '                Else
    '                    str &= "/" & sArray1(1)
    '                End If
    '                henkanMNo.Add(str)
    '            Else
    '                henkanMNo.Add(DV2.Item(1, r).Value)
    '            End If
    '        ElseIf InStr(DV2.Item(1, r).Value, "[選択]") > 0 Then
    '            If RadioButton16.Checked = True Then
    '                Dim str As String = Replace(DV2.Item(1, r).Value, "[選択]", "")
    '                nCol = Split(str, "|")
    '                Dim res As String = ""
    '                For i As Integer = 0 To nCol.Length - 1
    '                    For c As Integer = 0 To DV1.ColumnCount - 1
    '                        If nCol(i) = DV1.Item(c, 0).Value Then
    '                            If res = "" Then
    '                                res = c
    '                                flag = True
    '                            Else
    '                                res &= "|" & c
    '                                flag = True
    '                            End If
    '                        End If
    '                    Next
    '                Next
    '                If flag = True Then
    '                    henkanMNo.Add("[選択]" & res)
    '                Else
    '                    henkanMNo.Add(9999)
    '                End If
    '            Else
    '                henkanMNo.Add(DV2.Item(1, r).Value)
    '            End If
    '        ElseIf InStr(DV2.Item(1, r).Value, "+") > 0 Then    '複数コラムを元にする時
    '            If RadioButton16.Checked = True Then
    '                nCol = Split(DV2.Item(1, r).Value, "+")
    '                Dim res As String = ""
    '                For i As Integer = 0 To nCol.Length - 1
    '                    If InStr(nCol(i), Chr(34)) = 0 And InStr(nCol(i), Chr(39)) = 0 Then 'Chr(34)=ダブルコーテーション、Chr(39)=シングルコーテーション
    '                        For c As Integer = 0 To DV1.ColumnCount - 1
    '                            If nCol(i) = DV1.Item(c, 0).Value Then
    '                                If res = "" Then
    '                                    res = c
    '                                    flag = True
    '                                    Exit For
    '                                Else
    '                                    res &= "+" & c
    '                                    flag = True
    '                                    Exit For
    '                                End If
    '                            End If
    '                        Next
    '                    Else
    '                        If res = "" Then
    '                            res = nCol(i)
    '                        Else
    '                            res &= "+" & nCol(i)
    '                        End If
    '                    End If
    '                Next
    '                If flag = True Then
    '                    henkanMNo.Add(res)
    '                Else
    '                    henkanMNo.Add(9999)
    '                End If
    '            Else
    '                henkanMNo.Add(DV2.Item(1, r).Value)
    '            End If
    '        Else
    '            If RadioButton16.Checked = True Then
    '                For c As Integer = 0 To DV1.ColumnCount - 1
    '                    If DV2.Item(1, r).Value = DV1.Item(c, 0).Value Then
    '                        henkanMNo.Add(c)
    '                        flag = True
    '                        Exit For
    '                    End If
    '                Next
    '                If flag = False Then
    '                    henkanMNo.Add(9999)
    '                End If
    '            Else
    '                If DV2.Item(1, r).Value = "" Then
    '                    henkanMNo.Add("9999")
    '                Else
    '                    henkanMNo.Add(DV2.Item(1, r).Value)
    '                End If
    '            End If
    '        End If
    '    Next

    '    'データ一時退避
    '    DV3.Rows.Clear()
    '    DV3.Columns.Clear()
    '    DV3.Visible = True
    '    For c As Integer = 0 To DV1.ColumnCount - 1
    '        DV3.Columns.Add(c, c)
    '    Next
    '    For r As Integer = 0 To DV1.RowCount - 1
    '        DV3.Rows.Add(1)
    '        For c As Integer = 0 To DV1.ColumnCount - 1
    '            DV3.Item(c, r).Value = DV1.Item(c, r).Value
    '        Next
    '    Next

    '    ToolStripProgressBar1.Value = 0
    '    ToolStripProgressBar1.Maximum = DV1.RowCount

    '    '-----------------------------------------------------------------------
    '    PreDIV()
    '    '-----------------------------------------------------------------------
    '    'メインを消す
    '    DV1.Rows.Clear()
    '    DV1.Columns.Clear()

    '    For c As Integer = 0 To henkanSNo.Count - 1
    '        DV1.Columns.Add(c, c)
    '    Next

    '    If CheckBox1.Checked = True Then
    '        DV1.Rows.Add(1)
    '        For c As Integer = 0 To henkanSNo.Count - 1
    '            DV1.Item(c, 0).Value = henkanSHeader(c)
    '        Next
    '    End If

    '    Dim start As Integer = 0    'コラム番号を使用する際は0行から
    '    Dim addRow As Integer = 0
    '    If RadioButton16.Checked = True Then
    '        start = 1
    '        addRow = 0
    '    Else
    '        start = 0
    '        addRow = 1
    '    End If

    '    '伝票設定用ファイル読み取り
    '    Dim regStr As String = ""
    '    Dim regIrai As String = ""
    '    Dim dFlag As Integer = 0
    '    Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\config\denpyoConfig.txt"
    '    Using sr As New System.IO.StreamReader(fName, System.Text.Encoding.Default)
    '        Do While Not sr.EndOfStream
    '            Dim s As String = sr.ReadLine
    '            If s = "[伝票依頼主用]" Then
    '                dFlag = 1
    '            ElseIf s = "[依頼主変換]" Then
    '                dFlag = 2
    '            ElseIf dFlag = 1 Then
    '                regStr = s
    '            ElseIf dFlag = 2 Then
    '                regIrai = s
    '                Exit Do
    '            End If
    '        Loop
    '    End Using

    '    'datagridviewのヘッダーテキストをコレクションに取り込む
    '    Dim DGVheaderCheck3 As ArrayList = New ArrayList
    '    For c As Integer = 0 To DV3.Columns.Count - 1
    '        DGVheaderCheck3.Add(DV3.Item(c, 0).Value)
    '    Next c

    '    Dim koguchi As ArrayList = New ArrayList    'メール便用個口リスト
    '    For r As Integer = start To DV3.RowCount - 1
    '        DV1.Rows.Add(1)
    '        For c As Integer = 0 To henkanSNo.Count - 1
    '            Select Case True
    '                Case Regex.IsMatch(henkanMNo(c), "9999")
    '                    DV1.Item(c, r + addRow).Value = ""
    '                Case Regex.IsMatch(henkanMNo(c), "10002")
    '                    Dim formatStr As String = DV2.Item(3, c).Value
    '                    If formatStr = "" Then
    '                        formatStr = "yyyyMMdd"
    '                    End If
    '                    Select Case True
    '                        Case InStr(DV2.Item(2, c).Value, "今日")
    '                            DV1.Item(c, r + addRow).Value = Format(Today, formatStr)
    '                        Case InStr(DV2.Item(2, c).Value, "明日")
    '                            DV1.Item(c, r + addRow).Value = Format(DateAdd(DateInterval.Day, 1, Today), formatStr)
    '                        Case InStr(DV2.Item(2, c).Value, "日付")
    '                            Dim dStr As String = Replace(DV2.Item(2, c).Value, "日付=", "")
    '                            If IsNumeric(dStr) Then
    '                                DV1.Item(c, r + addRow).Value = Format(dStr, formatStr)
    '                            Else
    '                                dStr = DV3.Item(DGVheaderCheck3.IndexOf(dStr), r).Value
    '                                DV1.Item(c, r + addRow).Value = Format(CDate(dStr), formatStr)
    '                            End If
    '                    End Select
    '                Case Regex.IsMatch(henkanMNo(c), "10000")
    '                    DV1.Item(c, r + addRow).Value = DV2.Item(2, c).Value
    '                Case Regex.IsMatch(henkanMNo(c), "[変換].*")
    '                    Dim henkanC As String = Replace(henkanMNo(c), "[変換]", "")
    '                    Dim sArray1 As String() = Split(DV2.Item(2, c).Value, "|")
    '                    For i As Integer = 0 To sArray1.Length - 1
    '                        Dim sArray2 As String() = Split(sArray1(i), "=")
    '                        If sArray2(0) = "" Then
    '                            If DV3.Item(henkanC, r).Value = "" Then
    '                                DV1.Item(c, r + addRow).Value = sArray2(1)
    '                                Exit For
    '                            End If
    '                        Else
    '                            If InStr(DV3.Item(henkanC, r).Value, sArray2(0)) > 0 Then
    '                                DV1.Item(c, r + addRow).Value = sArray2(1)
    '                                Exit For
    '                            End If
    '                        End If
    '                    Next
    '                Case Regex.IsMatch(henkanMNo(c), "[分割].*")
    '                    Dim henkanC As String = Replace(henkanMNo(c), "[分割]", "")
    '                    Dim splitword As String = DV2.Item(2, c).Value
    '                    splitword = Replace(splitword, "\s", " ")
    '                    Dim sArray1 As String() = Split(splitword, "=")
    '                    If DV3.Item(henkanC, r).Value <> "" Then
    '                        Dim sArray2 As String() = Split(DV3.Item(henkanC, r).Value, sArray1(0))
    '                        If sArray1(1) = "last" Then '最後のワードを取得
    '                            DV1.Item(c, r + addRow).Value = sArray2(sArray2.Length - 1)
    '                        ElseIf sArray2.Length >= sArray1(1) Then
    '                            DV1.Item(c, r + addRow).Value = sArray2(sArray1(1) - 1)
    '                        Else
    '                            DV1.Item(c, r + addRow).Value = ""
    '                        End If
    '                    Else
    '                        DV1.Item(c, r + addRow).Value = ""
    '                    End If
    '                Case Regex.IsMatch(henkanMNo(c), "[合併].*")
    '                    Dim henkanC As String() = Split(Replace(henkanMNo(c), "[合併]", ""), "|")
    '                    Dim res As String = ""
    '                    For i As Integer = 0 To henkanC.Length - 1
    '                        If DV3.Item(henkanC(i), r).Value <> "" Then
    '                            res &= DV3.Item(henkanC(i), r).Value
    '                        End If
    '                    Next
    '                    DV1.Item(c, r + addRow).Value = res
    '                Case Regex.IsMatch(henkanMNo(c), "[検索].*")
    '                    Dim henkanC As String = Replace(henkanMNo(c), "[検索]", "")
    '                    Dim sArray1 As String() = Split(henkanC, "/")
    '                    Dim sArray2 As String() = Split(sArray1(0), "=")
    '                    If InStr(sArray2(1), ">") > 0 Then
    '                        Dim num As String = Replace(InStr(sArray2(1), ">"), ">", "")
    '                        If DV3.Item(sArray2(0), r).Value > num Then
    '                            Dim sArray3 As String() = Split(DV2.Item(2, c).Value, "/")
    '                            Dim sArray4 As String() = Split(sArray3(1), "|")
    '                            DV1.Item(c, r + addRow).Value = sArray4(0)
    '                        Else
    '                            Dim sArray3 As String() = Split(DV2.Item(2, c).Value, "/")
    '                            Dim sArray4 As String() = Split(sArray3(1), "|")
    '                            DV1.Item(c, r + addRow).Value = sArray4(1)
    '                        End If
    '                    ElseIf InStr(DV3.Item(sArray2(0), r).Value, sArray2(1)) > 0 Then
    '                        If InStr(sArray1(1), "[固定]") = 0 Then
    '                            DV1.Item(c, r + addRow).Value = DV3.Item(sArray1(1), r).Value
    '                        Else
    '                            Dim sArray3 As String() = Split(DV2.Item(2, c).Value, "/")
    '                            Dim sArray4 As String() = Split(sArray3(1), "|")
    '                            DV1.Item(c, r + addRow).Value = sArray4(0)
    '                        End If
    '                    Else
    '                        If InStr(sArray1(1), "[固定]") = 0 Then
    '                            DV1.Item(c, r + addRow).Value = ""
    '                        Else
    '                            Dim sArray3 As String() = Split(DV2.Item(2, c).Value, "/")
    '                            Dim sArray4 As String() = Split(sArray3(1), "|")
    '                            DV1.Item(c, r + addRow).Value = sArray4(1)
    '                        End If
    '                    End If
    '                Case Regex.IsMatch(henkanMNo(c), "[選択].*")
    '                    Dim sentakuC As String = Replace(henkanMNo(c), "[選択]", "")
    '                    If InStr(sentakuC, "|") > 0 Then
    '                        Dim arrayA As String() = Split(sentakuC, "|")
    '                        Dim res As String = ""
    '                        For i As Integer = 0 To arrayA.Length - 1
    '                            Dim a = DV2.Item(2, c).Value
    '                            If DV3.Item(arrayA(i), r).Value <> DV2.Item(2, c).Value And DV3.Item(arrayA(i), r).Value <> "" Then
    '                                res = DV3.Item(arrayA(i), r).Value
    '                                Exit For
    '                            End If
    '                        Next
    '                        DV1.Item(c, r + addRow).Value = res
    '                    Else
    '                        DV1.Item(c, r + addRow).Value = DV3.Item(henkanMNo(c), r).Value
    '                    End If
    '                Case Regex.IsMatch(henkanMNo(c), "[特殊].*")
    '                    If InStr(DV2.Item(2, c).Value, "Qoo10") > 0 Then
    '                        If InStr(ComboBox1.SelectedItem, "佐川") > 0 Then
    '                            Dim tokuC As String = Replace(henkanMNo(c), "[特殊]", "")
    '                            If InStr(DV3.Item(tokuC, r).Value, "-") > 0 And InStr(DV3.Item(tokuC, r).Value, ":") > 0 Then
    '                                Dim arrayA As String() = Split(DV3.Item(tokuC, r).Value, " ")
    '                                If DV2.Item(2, c).Value = "[Qoo10日]" Then
    '                                    Dim arrayB As String() = Split(arrayA(0), "-")
    '                                    DV1.Item(c, r + addRow).Value = arrayB(1) & "月" & arrayB(2) & "日"
    '                                ElseIf DV2.Item(2, c).Value = "[Qoo10時]" Then
    '                                    DV1.Item(c, r + addRow).Value = Qoo10time(arrayA(1))
    '                                End If
    '                            ElseIf InStr(DV3.Item(tokuC, r).Value, "-") > 0 And InStr(DV3.Item(tokuC, r).Value, ":") = 0 Then
    '                                If DV2.Item(2, c).Value = "[Qoo10日]" Then
    '                                    Dim arrayB As String() = Split(DV3.Item(tokuC, r).Value, "-")
    '                                    DV1.Item(c, r + addRow).Value = arrayB(1) & "月" & arrayB(2) & "日"
    '                                End If
    '                            ElseIf InStr(DV3.Item(tokuC, r).Value, "-") = 0 And InStr(DV3.Item(tokuC, r).Value, ":") > 0 Then
    '                                If DV2.Item(2, c).Value = "[Qoo10時]" Then
    '                                    DV1.Item(c, r + addRow).Value = Qoo10time(DV3.Item(tokuC, r).Value)
    '                                End If
    '                            End If
    '                        Else
    '                            Dim tokuC As String = Replace(henkanMNo(c), "[特殊]", "")
    '                            If InStr(DV3.Item(tokuC, r).Value, "-") > 0 And InStr(DV3.Item(tokuC, r).Value, ":") > 0 Then
    '                                Dim arrayA As String() = Split(DV3.Item(tokuC, r).Value, " ")
    '                                If DV2.Item(2, c).Value = "[Qoo10日]" Then
    '                                    DV1.Item(c, r + addRow).Value = Replace(arrayA(0), "-", "")
    '                                ElseIf DV2.Item(2, c).Value = "[Qoo10時]" Then
    '                                    DV1.Item(c, r + addRow).Value = Qoo10time(arrayA(1))
    '                                End If
    '                            ElseIf InStr(DV3.Item(tokuC, r).Value, "-") > 0 And InStr(DV3.Item(tokuC, r).Value, ":") = 0 Then
    '                                If DV2.Item(2, c).Value = "[Qoo10日]" Then
    '                                    DV1.Item(c, r + addRow).Value = Replace(DV3.Item(tokuC, r).Value, "-", "")
    '                                End If
    '                            ElseIf InStr(DV3.Item(tokuC, r).Value, "-") = 0 And InStr(DV3.Item(tokuC, r).Value, ":") > 0 Then
    '                                If DV2.Item(2, c).Value = "[Qoo10時]" Then
    '                                    DV1.Item(c, r + addRow).Value = Qoo10time(DV3.Item(tokuC, r).Value)
    '                                End If
    '                            End If
    '                        End If
    '                    ElseIf InStr(DV2.Item(2, c).Value, "住所") > 0 Then
    '                        Dim tokuC As String = Replace(henkanMNo(c), "[特殊]", "")
    '                        Dim mode As String = Replace(DV2.Item(2, c).Value, "[住所]", "")
    '                        Dim address As String = DV3.Item(tokuC, r).Value

    '                        Dim addressCheck As Boolean = False
    '                        Dim work As String
    '                        Dim i As Integer
    '                        For i = 0 To address.Length - 1 '住所の文字数分繰り返す
    '                            work = address.Substring(i, 1)
    '                            If IsNumeric(work) Then
    '                                addressCheck = True
    '                                Exit For
    '                            End If
    '                        Next i
    '                        If mode = 1 Then
    '                            DV1.Item(c, r + addRow).Value = address.Substring(0, i)

    '                            If addressCheck = False Then    '番地未入力チェック
    '                                For cc As Integer = 0 To DV1.ColumnCount - 1
    '                                    DV1.Item(cc, r).Style.BackColor = Color.Yellow
    '                                Next
    '                            End If
    '                        ElseIf mode = 2 Then
    '                            DV1.Item(c, r + addRow).Value = address.Substring(i)
    '                        End If
    '                    ElseIf InStr(DV2.Item(2, c).Value, "個口") > 0 Then '受注者電話番号の後ろに「,2」とし、複数個口を自動処理する
    '                        Dim tokuC As String = Replace(henkanMNo(c), "[特殊]", "")
    '                        Dim denwa As String = DV3.Item(tokuC, r).Value

    '                        If InStr(denwa, splitter) > 0 Then
    '                            Dim kosuu As String() = Split(denwa, splitter)
    '                            If IsNumeric(kosuu(1)) Then '個口が数字で、1個以上の場合のみ
    '                                If kosuu(1) > 1 Then
    '                                    DV1.Item(c, r + addRow).Value = kosuu(1)
    '                                    koguchi.Add(r + addRow & splitter & kosuu(1))
    '                                End If
    '                            End If
    '                        End If
    '                    ElseIf InStr(DV2.Item(2, c).Value, "依頼主") > 0 Then
    '                        Dim tokuC As String = Replace(henkanMNo(c), "[特殊]", "")
    '                        Dim kigyoumei As String = Replace(DV3.Item(tokuC, r).Value, " ", "")
    '                        kigyoumei = Replace(kigyoumei, "　", "")

    '                        Select Case True
    '                            Case Regex.IsMatch(kigyoumei, regStr)   '伝票設定用ファイルから設定読み取り
    '                                DV1.Item(c, r + addRow).Value = ""
    '                            Case Else
    '                                Dim todokeSaki As String = Replace(DV3.Item(hassouSakiMei, r).Value, " ", "")
    '                                todokeSaki = Replace(todokeSaki, "　", "")
    '                                If todokeSaki = kigyoumei Then
    '                                    DV1.Item(c, r + addRow).Value = ""  '企業名と届け先が同じなら変更しない
    '                                Else
    '                                    '依頼主変換のチェック
    '                                    Dim iraiArray As String() = Split(regIrai, "|")
    '                                    Dim iraiFlag As Boolean = False
    '                                    Dim iraiHenkan As String() = Nothing
    '                                    For i As Integer = 0 To iraiArray.Length - 1
    '                                        Dim namaeArray As String() = Split(iraiArray(i), splitter)
    '                                        If kigyoumei = namaeArray(0) Then
    '                                            iraiHenkan = namaeArray
    '                                            iraiFlag = True
    '                                            Exit For
    '                                        End If
    '                                    Next
    '                                    If iraiFlag = False Then
    '                                        DV1.Item(c, r + addRow).Value = "(依頼主)[" & DV3.Item(henkanIraiTel, r).Value & "]" & kigyoumei & "様"
    '                                    Else
    '                                        DV1.Item(c, r + addRow).Value = "(依頼主)[" & iraiHenkan(2) & "]" & iraiHenkan(1) & "様"
    '                                    End If
    '                                End If
    '                        End Select
    '                    ElseIf InStr(DV2.Item(2, c).Value, "電話") > 0 Then
    '                        Dim tokuC As String = Replace(henkanMNo(c), "[特殊]", "")
    '                        Dim kigyoumei As String = Replace(DV3.Item(tokuC, r).Value, " ", "")
    '                        kigyoumei = Replace(kigyoumei, "　", "")

    '                        Select Case True
    '                            Case Regex.IsMatch(kigyoumei, regStr)   '伝票設定用ファイルから設定読み取り
    '                                Dim res As String = Replace(DV2.Item(2, c).Value, "[電話]", "")
    '                                DV1.Item(c, r + addRow).Value = res
    '                            Case Else
    '                                Dim todokeSaki As String = Replace(DV3.Item(hassouSakiMei, r).Value, " ", "")
    '                                todokeSaki = Replace(todokeSaki, "　", "")
    '                                If todokeSaki = kigyoumei Then
    '                                    Dim res As String = Replace(DV2.Item(2, c).Value, "[電話]", "")
    '                                    DV1.Item(c, r + addRow).Value = res
    '                                Else
    '                                    DV1.Item(c, r + addRow).Value = ""
    '                                End If
    '                        End Select
    '                    End If
    '                Case Else
    '                    If InStr(henkanMNo(c), "+") > 0 Then
    '                        Dim arrayA As String() = Split(henkanMNo(c), "+")
    '                        Dim res As String = ""
    '                        For i As Integer = 0 To arrayA.Length - 1
    '                            If InStr(arrayA(i), Chr(34)) = 0 And InStr(arrayA(i), Chr(39)) = 0 Then
    '                                res &= DV3.Item(arrayA(i), r).Value
    '                            Else
    '                                Dim resA As String = Replace(arrayA(i), Chr(34), "")
    '                                resA = Replace(arrayA(i), Chr(39), "")
    '                                res &= resA
    '                            End If
    '                        Next
    '                        DV1.Item(c, r + addRow).Value = res
    '                    Else
    '                        DV1.Item(c, r + addRow).Value = DV3.Item(henkanMNo(c), r).Value
    '                    End If
    '            End Select

    '            '文字列の置換
    '            If henkanO(c) <> "" Then
    '                Dim rp1 As String() = Split(henkanO(c), "|")
    '                For i As Integer = 0 To rp1.Length - 1
    '                    If InStr(rp1(i), "=") > 0 Then
    '                        Dim rp2 As String() = Split(rp1(i), "=")
    '                        DV1.Item(c, r + addRow).Value = Replace(DV1.Item(c, r + addRow).Value, rp2(0), rp2(1))
    '                    End If
    '                Next
    '            End If
    '        Next
    '    Next
    '    DV3.Visible = False

    '    'メール便のみ個口増加処理（複数個口で動作するので却下）
    '    'If InStr(ComboBox1.SelectedItem, "ゆパケ") > 0 Then
    '    '    For i As Integer = koguchi.Count - 1 To 0 Step -1
    '    '        Dim kogStr As String() = Split(koguchi(i), splitter) '対象行番号,個口数
    '    '        For k As Integer = 0 To kogStr(1) - 2           '2個以上
    '    '            Dim selR As Integer = kogStr(0) + k
    '    '            dgv.Rows.InsertCopy(selR, selR + 1)
    '    '            For c As Integer = 0 To dgv.ColumnCount - 1
    '    '                If dgv.Item(c, selR).Value <> "" Then
    '    '                    dgv.Item(c, selR + 1).Value = dgv.Item(c, selR).Value
    '    '                End If
    '    '            Next
    '    '        Next
    '    '    Next
    '    'End If

    '    '最後にプログラム処理を行なう
    '    For i As Integer = 0 To ProgramList.Count - 1
    '        If Regex.IsMatch(ProgramList(i), "^伝票チェック") Then
    '            Dim DGVheaderCheck1 As ArrayList = New ArrayList
    '            For c As Integer = 0 To DV1.Columns.Count - 1
    '                DGVheaderCheck1.Add(DV1.Item(c, 0).Value)
    '            Next c

    '            Dim PL As String() = Split(ProgramList(i), splitter)
    '            Dim Bangou1 As String = ""
    '            Dim Bangou2 As String = ""
    '            Dim Bangou3 As String = ""
    '            If PL(1) = "楽天" Then
    '                Bangou1 = "受注番号"
    '                Bangou2 = "お荷物伝票番号"
    '            ElseIf PL(1) = "Yahoo" Then
    '                Bangou1 = "OrderId"
    '                Bangou2 = "ShipInvoiceNumber1"
    '                Bangou3 = "ShipInvoiceNumber2"
    '            ElseIf PL(1) = "MakeShop" Then
    '                Bangou1 = "注文番号"
    '                Bangou2 = "伝票番号"
    '            ElseIf PL(1) = "Amazon" Then
    '                Bangou1 = "order-id"
    '                Bangou2 = "tracking-number"
    '            End If

    '            Dim allArray As ArrayList = New ArrayList       '伝票番号
    '            Dim poolArray As ArrayList = New ArrayList      '伝票番号
    '            Dim removeArray As ArrayList = New ArrayList    '行
    '            For r As Integer = 1 To DV1.RowCount - 1
    '                Dim DNum As String = DataGridView1.Item(DGVheaderCheck1.IndexOf(Bangou1), r).Value

    '                If InStr(DNum, "-fk-") > 0 Then '複写伝票は削除
    '                    Dim DNum1 As String() = Regex.Split(DNum, "-fk-")
    '                    allArray.Add(DNum1(0))
    '                    removeArray.Add(r)
    '                ElseIf InStr(DNum, "-bk-") > 0 Then
    '                    '分割フラグ（-bk）が付いている受注番号が他にあれば削除対象にする
    '                    Dim DNum1 As String() = Regex.Split(DNum, "-bk-|-h")
    '                    allArray.Add(DNum1(0))

    '                    If Not poolArray.Contains(DNum1(0)) Then
    '                        DataGridView1.Item(DGVheaderCheck1.IndexOf(Bangou1), r).Value = DNum1(0) '分割番号を正しい番号に直す
    '                        poolArray.Add(DNum1(0))
    '                    Else
    '                        Dim motoGyou As Integer = allArray.IndexOf(DNum1(0)) + 1
    '                        Select Case PL(1)
    '                            Case "楽天"
    '                                DataGridView1.Item(DGVheaderCheck1.IndexOf(Bangou2), motoGyou).Value &= splitter & DataGridView1.Item(DGVheaderCheck1.IndexOf(Bangou2), r).Value
    '                            Case "Yahoo"
    '                                If DataGridView1.Item(DGVheaderCheck1.IndexOf(Bangou2), motoGyou).Value = "" Then   'ShipInvoiceNumber1が空いている時
    '                                    DataGridView1.Item(DGVheaderCheck1.IndexOf(Bangou2), motoGyou).Value = DataGridView1.Item(DGVheaderCheck1.IndexOf(Bangou2), r).Value
    '                                ElseIf DataGridView1.Item(DGVheaderCheck1.IndexOf(Bangou3), motoGyou).Value = "" Then   'ShipInvoiceNumber2が空いている時
    '                                    DataGridView1.Item(DGVheaderCheck1.IndexOf(Bangou3), motoGyou).Value = DataGridView1.Item(DGVheaderCheck1.IndexOf(Bangou2), r).Value
    '                                Else
    '                                    Dim a = 0
    '                                End If
    '                            Case "MakeShop", "Amazon"
    '                                '分割伝票の分は削除する
    '                        End Select
    '                        removeArray.Add(r)
    '                    End If
    '                Else
    '                    '各モールの規定文字列ではない受注番号の場合削除対象にする
    '                    Dim kiteiFalseFlag As Boolean = True
    '                    Select Case PL(1)
    '                        Case "楽天"
    '                            If Not Regex.IsMatch(DNum, "[0-9]{6}-[0-9]{8}-[0-9]{8}") Then
    '                                kiteiFalseFlag = False
    '                            End If
    '                        Case "Yahoo"
    '                            If Not Regex.IsMatch(DNum, "[0-9]{8}") Then
    '                                kiteiFalseFlag = False
    '                            End If
    '                        Case "MakeShop"
    '                            If Not Regex.IsMatch(DNum, "P[0-9]{18}") Then
    '                                kiteiFalseFlag = False
    '                            End If
    '                        Case "Amazon"
    '                            If Not Regex.IsMatch(DNum, "[0-9]{3}-[0-9]{7}-[0-9]{7}") Then
    '                                kiteiFalseFlag = False
    '                            End If
    '                    End Select

    '                    If kiteiFalseFlag = True Then
    '                        If Not poolArray.Contains(DNum) Then
    '                            allArray.Add(DNum)
    '                            poolArray.Add(DNum)
    '                        Else
    '                            Dim motoGyou As Integer = allArray.IndexOf(DNum) + 1
    '                            Select Case PL(1)
    '                                Case "楽天"   'エラー消えてしまう（未完成）
    '                                    DataGridView1.Item(DGVheaderCheck1.IndexOf(Bangou2), motoGyou).Value &= splitter & DataGridView1.Item(DGVheaderCheck1.IndexOf(Bangou2), r).Value
    '                                Case "Yahoo"
    '                                    If DataGridView1.Item(DGVheaderCheck1.IndexOf(Bangou2), motoGyou).Value = "" Then   'ShipInvoiceNumber1が空いている時
    '                                        DataGridView1.Item(DGVheaderCheck1.IndexOf(Bangou2), motoGyou).Value = DataGridView1.Item(DGVheaderCheck1.IndexOf(Bangou2), r).Value
    '                                    ElseIf DataGridView1.Item(DGVheaderCheck1.IndexOf(Bangou3), motoGyou).Value = "" Then   'ShipInvoiceNumber2が空いている時
    '                                        DataGridView1.Item(DGVheaderCheck1.IndexOf(Bangou3), motoGyou).Value = DataGridView1.Item(DGVheaderCheck1.IndexOf(Bangou2), r).Value
    '                                    Else
    '                                        Dim a = 0
    '                                    End If
    '                                Case "MakeShop", "Amazon"
    '                                    '分割伝票の分は削除する
    '                            End Select
    '                            removeArray.Add(r)
    '                        End If
    '                    Else
    '                        removeArray.Add(r)
    '                    End If

    '                End If
    '            Next
    '            For k As Integer = removeArray.Count - 1 To 0 Step -1
    '                DataGridView1.Rows.RemoveAt(removeArray(k))
    '            Next
    '        End If
    '    Next

    '    '-----------------------------------------------------------------------
    '    RetDIV()
    '    '-----------------------------------------------------------------------

    '    ToolStripProgressBar1.Value = 0
    '    'Catch ex As Exception
    '    '    DataGridView4.Visible = False
    '    '    MsgBox("読込ファイルの形式が異なるため処理できません" & vbCrLf & ex.Message, MsgBoxStyle.SystemModal)
    '    'End Try
    'End Sub

    Private Function Qoo10time(ByVal str As String)
        If InStr(ComboBox1.SelectedItem, "佐川") > 0 Then
            Select Case str
                Case "09:00~12:00"
                    str = "01"
                Case "12:00~14:00"
                    str = "12"
                Case "14:00~16:00"
                    str = "14"
                Case "16:00~18:00"
                    str = "16"
                Case "18:00~20:00"
                    str = "18"
                Case "19:00~21:00"
                    str = "19"
                Case "18:00~21:00"
                    str = "04"
            End Select
        Else    'ゆうパック
            Select Case str
                Case "09:00~12:00"
                    str = "午前中"
                Case "12:00~14:00"
                    str = "12時-14時"
                Case "14:00~16:00"
                    str = "14時-16時"
                Case "16:00~18:00"
                    str = "16時-18時"
                Case "18:00~20:00"
                    str = "18時-20時"
                Case "19:00~21:00"
                    str = "19時-21時"
                Case "18:00~21:00"
                    str = "18時-21時"
                Case "20:00~21:00"
                    str = "19時-21時"
            End Select
        End If
        Return str
    End Function

    'フォームのリセット
    Private Sub フォームのリセット再描画ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles フォームのリセット再描画ToolStripMenuItem.Click
        Dim frm As Form = New Csv_change
        Me.Dispose()
        frm.Show()
    End Sub

    Private Sub フォームのリセットToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles フォームのリセットToolStripMenuItem.Click
        InitControl(Me)
    End Sub

    Private Sub InitControl(ByVal Ctrl As Control)
        For Each CtrlItem As Control In Ctrl.Controls
            If TypeOf CtrlItem Is TextBox Then
                CtrlItem.Text = ""
            ElseIf TypeOf CtrlItem Is DataGridView Then
                Select Case CtrlItem.Name
                    Case "DataGridView1", "DataGridView5", "DataGridView4"
                        DirectCast(CtrlItem, DataGridView).Columns.Clear()
                    Case "DataGridView2"
                        'リセットしない
                    Case Else
                        DirectCast(CtrlItem, DataGridView).Rows.Clear()
                End Select
                If CtrlItem.Name = "DataGridView4" Then
                    DataGridView4.Visible = False
                End If
            End If

            If CtrlItem.Controls.Count > 0 Then
                InitControl(CtrlItem)
            End If
        Next
    End Sub

    Private Sub 右上グリッドのリセットToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles 右上グリッドのリセットToolStripMenuItem.Click
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
    End Sub

    Private Sub 右下グリッドのリセットToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles 右下グリッドのリセットToolStripMenuItem.Click
        DataGridView5.Rows.Clear()
        DataGridView5.Columns.Clear()
    End Sub

    Private Sub DataGridView1_GotFocus(ByVal sender As Object, ByVal e As EventArgs) Handles DataGridView1.GotFocus
        Panel4.BackColor = Color.Yellow
        Panel5.BackColor = Color.Gainsboro
        dgv = DataGridView1
    End Sub

    Private Sub DataGridView5_GotFocus(ByVal sender As Object, ByVal e As EventArgs) Handles DataGridView5.GotFocus
        Panel4.BackColor = Color.Gainsboro
        Panel5.BackColor = Color.Yellow
        dgv = DataGridView5
    End Sub

    '行番号を表示する
    Private Sub DataGridView1_RowPostPaint(ByVal sender As DataGridView, ByVal e As DataGridViewRowPostPaintEventArgs) Handles _
            DataGridView1.RowPostPaint, DataGridView2.RowPostPaint, DataGridView3.RowPostPaint, DataGridView5.RowPostPaint, DataGridView6.RowPostPaint
        ' 行ヘッダのセル領域を、行番号を描画する長方形とする（ただし右端に4ドットのすき間を空ける）
        Dim rect As New Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, sender.RowHeadersWidth - 4, sender.Rows(e.RowIndex).Height)
        ' 上記の長方形内に行番号を縦方向中央＆右詰で描画する　フォントや色は行ヘッダの既定値を使用する
        TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), sender.RowHeadersDefaultCellStyle.Font, rect, sender.RowHeadersDefaultCellStyle.ForeColor, TextFormatFlags.VerticalCenter Or TextFormatFlags.Right)
    End Sub

    Private Sub 上書き保存ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles 上書き保存ToolStripMenuItem.Click
        If dgv.RowCount = 0 Then
            Exit Sub
        End If

        If dgv Is DataGridView1 And ToolStripLabel2.Text = "..." Then
            MsgBox("保存できません", MsgBoxStyle.SystemModal)
            Exit Sub
        ElseIf dgv Is DataGridView5 And ToolStripStatusLabel1.Text = "..." Then
            MsgBox("保存できません", MsgBoxStyle.SystemModal)
            Exit Sub
        Else
            If dgv Is DataGridView1 Then
                SaveCsv(ToolStripLabel2.Text, dgv, False, 0)
                MsgBox(ToolStripLabel2.Text & vbCrLf & "上書き保存しました", MsgBoxStyle.SystemModal)
            Else
                SaveCsv(ToolStripStatusLabel1.Text, dgv, False, 0)
                MsgBox(ToolStripStatusLabel1.Text & vbCrLf & "上書き保存しました", MsgBoxStyle.SystemModal)
            End If
        End If
    End Sub

    Private Sub 保存ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles 保存ToolStripMenuItem.Click
        If dgv.RowCount = 0 Then
            Exit Sub
        End If

        Dim sfd As New SaveFileDialog With {
            .FileName = "新しいファイル.csv",
            .InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
            .Filter = "CSVファイル(*.csv)|*.csv|すべてのファイル(*.*)|*.*",
                    .AutoUpgradeEnabled = False,
            .FilterIndex = 0,
            .Title = "保存先のファイルを選択してください",
            .RestoreDirectory = True,
            .OverwritePrompt = True,
            .CheckPathExists = True
        }

        'ダイアログを表示する
        If sfd.ShowDialog() = DialogResult.OK Then
            SaveCsv(sfd.FileName, dgv, False, 0)
            MsgBox(sfd.FileName & vbCrLf & "保存しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    Private Sub 出荷通知テキストタブ区切りToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 出荷通知テキストタブ区切りToolStripMenuItem.Click
        AmazonTabTxt(System.Text.Encoding.GetEncoding("SHIFT-JIS"), "SHIFT-JIS")
    End Sub

    Private Sub 中国語OS用ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 中国語OS用ToolStripMenuItem.Click
        AmazonTabTxt(System.Text.Encoding.GetEncoding("GBK"), "GBK")
    End Sub

    Private Sub 中国語OS用utf16ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 中国語OS用utf16ToolStripMenuItem.Click
        AmazonTabTxt(System.Text.Encoding.GetEncoding("utf-16"), "utf-16")
    End Sub

    Private Sub AmazonTabTxt(aEnc As Text.Encoding, bEnc As String)
        If dgv.RowCount = 0 Then
            Exit Sub
        End If

        Dim sfd As New SaveFileDialog With {
            .FileName = "Amazon出荷通知_" & Format(Now(), "yyMMddHHmmss") & ".txt",
            .InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
            .Filter = "TXTファイル(*.txt)|*.txt|すべてのファイル(*.*)|*.*",
                    .AutoUpgradeEnabled = False,
            .FilterIndex = 0,
            .Title = "保存先のファイルを選択してください",
            .RestoreDirectory = True,
            .OverwritePrompt = True,
            .CheckPathExists = True
        }

        'Amazon用テンプレート読み出し
        Dim loadPath As String = Path.GetDirectoryName(Form1.appPath) & "\config\save_template\Amazon_Shipment.xls"
        Dim loadArray As ArrayList = TM_XLS_READ(loadPath, "出荷通知テンプレート")
        Dim sheet As Integer = 1
        Dim maxCol As Integer = 9   '1～
        Dim maxRow As Integer = 2

        Dim headerStr As String = ""
        For r As Integer = 0 To maxRow - 1
            Dim line As String() = Split(loadArray(r), "|=|")
            For c As Integer = 0 To maxCol - 1
                If c < line.Length Then
                    headerStr &= line(c) & vbTab
                End If
            Next
            headerStr = headerStr.TrimEnd()
            headerStr &= vbLf
        Next

        'ダイアログを表示する
        If sfd.ShowDialog(Me) = DialogResult.OK Then
            File.WriteAllText(sfd.FileName, headerStr, aEnc)  'ヘッダー書き込み
            SaveCsv(sfd.FileName, dgv, True, 2, bEnc)
            MsgBox(sfd.FileName & vbCrLf & "保存しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    Private Sub Xlsxで保存ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Xlsxで保存ToolStripMenuItem.Click
        If DataGridView1.RowCount = 0 Then
            Exit Sub
        End If

        Dim sfd As New SaveFileDialog With {
            .InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
            .Filter = "XLSXファイル(*.xlsx)|*.xlsx|すべてのファイル(*.*)|*.*",
                    .AutoUpgradeEnabled = False,
            .FilterIndex = 0,
            .Title = "保存先のファイルを選択してください",
            .RestoreDirectory = True,
            .OverwritePrompt = True,
            .CheckPathExists = True
        }
        If InStr(Me.Text, "\") > 0 Then
            Dim sPath As String = Path.GetFileNameWithoutExtension(ToolStripLabel2.Text)
            sfd.FileName = sPath & ".xlsx"
        Else
            sfd.FileName = "新しいファイル.xlsx"
        End If

        DataGridView1.EndEdit()

        'ダイアログを表示する
        If sfd.ShowDialog(Me) = DialogResult.OK Then
            Dim visibleFlag As Boolean = False

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

            'Dim desk As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
            Using wfs = File.Create(sfd.FileName)
                xWorkbook.Write(wfs)
            End Using
            sheet1 = Nothing
            xWorkbook = Nothing

            MsgBox(sfd.FileName & vbCrLf & "保存しました", MsgBoxStyle.SystemModal)
        End If
    End Sub


    ' DataGridViewをCSV出力する
    Public Sub SaveCsv(ByVal fp As String, ByVal dg As DataGridView, ByVal addFlag As Boolean, ByVal mode As Integer, Optional aEnc As String = "SHIFT-JIS")
        ' CSVファイルオープン（addFlag：true=追加、false=上書き）
        Dim sw As StreamWriter = New StreamWriter(fp, addFlag, System.Text.Encoding.GetEncoding(aEnc))
        Dim minusRow As Integer = 1
        If dg.AllowUserToAddRows = True Then
            minusRow = 2
        End If

        Dim sagawaFlag As Boolean = False
        If InStr(ComboBox1.SelectedItem, "佐川") > 0 And InStr(ComboBox1.SelectedItem, "変換") > 0 And mode = 0 Then
            sagawaFlag = True
        End If

        Dim CsvDelimiter As String = splitter
        If mode = 2 Then    'タブ区切り
            CsvDelimiter = vbTab
        End If

        DataGridView1.EndEdit()

        Dim writeStr As String = ""
        Dim printRow As Integer = 0
        For r As Integer = 0 To dg.Rows.Count - minusRow
            If CheckBox6.Checked = True And sagawaFlag = True Then    '佐川出力逆順対応
                printRow = dg.Rows.Count - minusRow - r
            Else
                printRow = r
            End If

            Dim headerStr As String = ""    'ヘッダーを調べるため文字列を作る
            For c As Integer = 0 To dg.Columns.Count - 1
                headerStr &= dg.Item(c, printRow).Value
            Next

            Dim dl As String = ""
            For c As Integer = 0 To dg.Columns.Count - 1
                ' DataGridViewのセルのデータ取得
                Dim dt As String = ""
                If dg.Rows(printRow).Cells(c).Value Is Nothing = False Then
                    dt = dg.Rows(printRow).Cells(c).Value.ToString()
                    dt = Replace(dt, vbCrLf, vbLf)
                    dt = Replace(dt, vbLf, "")
                    If Not dt Is Nothing Then
                        dt = Replace(dt, """", """""")
                        Select Case True
                            Case Regex.IsMatch(dt, ".*,.*|.*" & vbLf & ".*|.*" & vbCr & ".*|.*"".*")
                                dt = """" & dt & """"
                        End Select
                    End If
                End If
                If c < dg.Columns.Count - 1 Then
                    dt = dt & CsvDelimiter
                End If
                dl &= dt
            Next

            If InStr(headerStr, "住所録コードお届け先電話番号") > 0 Then
                writeStr = dl & vbLf & writeStr
            Else
                writeStr &= dl & vbLf
            End If
        Next
        sw.Write(writeStr)
        ' CSVファイルクローズ
        sw.Close()
    End Sub

    '名前を付けて保存（件数分割）
    Private Sub 件ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles 件ToolStripMenuItem.Click
        SaveA(200)
    End Sub

    Private Sub 件ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles 件ToolStripMenuItem1.Click
        SaveA(300)
    End Sub

    Private Sub 件ToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles 件ToolStripMenuItem2.Click
        SaveA(400)
    End Sub

    Private Sub SaveA(ByVal kensu As Integer)
        If dgv.RowCount = 0 Then
            Exit Sub
        End If

        Dim sfd As New SaveFileDialog With {
            .FileName = "新しいファイル.csv",
            .InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
            .Filter = "CSVファイル(*.csv)|*.csv|すべてのファイル(*.*)|*.*",
                    .AutoUpgradeEnabled = False,
            .FilterIndex = 0,
            .Title = "保存先のファイルを選択してください",
            .RestoreDirectory = True,
            .OverwritePrompt = True,
            .CheckPathExists = True
        }

        DataGridView1.EndEdit()

        'ダイアログを表示する
        If sfd.ShowDialog(Me) = DialogResult.OK Then
            SaveCsv(sfd.FileName, dgv, False, 0)

            If kensu > 0 Then
                'ファイル読み取り
                Dim fPath As String = sfd.FileName
                Dim motoDir As String = Path.GetDirectoryName(fPath)
                Dim motoName As String = Path.GetFileNameWithoutExtension(fPath)
                Dim header As String = ""
                Dim FileNum As Integer = 1
                'Dim lineCount As Integer = cntLineFromFile(fPath)
                'Dim maxNum As Integer = lineCount / kensu

                Dim str As String = ""
                Dim r As Integer = 0
                Using sr As New StreamReader(fPath, System.Text.Encoding.Default)
                    Do While Not sr.EndOfStream
                        Dim s As String = sr.ReadLine
                        If r = 0 Then
                            header = s & vbCrLf
                        Else
                            If r Mod kensu = 0 Then
                                str = header & str & s
                                Dim saveName As String = motoDir & "\" & motoName & "_" & FileNum & ".csv"
                                FileNum += 1
                                My.Computer.FileSystem.WriteAllText(saveName, str, False, Form1.enc)
                                str = ""
                            Else
                                str &= s & vbCrLf
                            End If
                        End If
                        r += 1
                    Loop
                End Using
                If str <> "" Then
                    str = header & str
                    Dim saveName As String = motoDir & "\" & motoName & "_" & FileNum & ".csv"
                    My.Computer.FileSystem.WriteAllText(saveName, str, False, Form1.enc)
                End If
                My.Computer.FileSystem.DeleteFile(fPath)
            End If

            MsgBox(sfd.FileName & vbCrLf & "保存しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    'ファイル行数取得
    Public Function CntLineFromFile(ByVal filePath As String) As Long
        Dim StReader As New StreamReader(filePath)
        Dim cnt As Long
        While (StReader.Peek() >= 0)
            StReader.ReadLine()
            cnt += 1
        End While
        Return cnt
    End Function


    '空行削除
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button1.Click
        If dgv.RowCount = 0 Then
            Exit Sub
        End If

        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim nowR As Integer = 0
        Dim maxR As Integer = dgv.RowCount - 1
        Dim delR As Integer = 0
        For r As Integer = 0 To maxR
            nowR = maxR - r
            Dim sFlag As Boolean = True
            For c As Integer = 0 To dgv.ColumnCount - 1
                If dgv.Item(c, nowR).Value <> "" Then
                    sFlag = False
                    Exit For
                End If
            Next
            If sFlag = True Then
                dgv.Rows.RemoveAt(nowR)
                delR += 1
            End If
        Next
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------

        If AutoMode = False Then
            MsgBox(delR & "行削除しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    Private Sub Csv1ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Csv1ToolStripMenuItem.Click
        CsvIkou(Form1.CSV_FORMS(0), Form1.CSV_FORMS(0).DataGridView1)
    End Sub

    Private Sub Csv2ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Csv2ToolStripMenuItem.Click
        CsvIkou(Form1.CSV_FORMS(1), Form1.CSV_FORMS(1).DataGridView1)
    End Sub

    Private Sub CsvIkou(ByVal frm As Object, ByVal csvDg As DataGridView)
        If frm.Visible = False Then
            frm.Show()
        End If

        If csvDg.RowCount > 0 Then
            csvDg.Rows.Clear()
            csvDg.Columns.Clear()
        End If

        '-----------------------------------------------------------------------
        frm.preDIV()
        '-----------------------------------------------------------------------
        For c As Integer = 0 To dgv.ColumnCount - 1
            csvDg.Columns.Add("column" & c + 1, c)
        Next

        Dim flag As Boolean = True
        For r As Integer = 0 To dgv.RowCount - 1
            If r = dgv.RowCount - 1 Then    '最後の行にデータが入っているか調べる
                flag = False
                For c As Integer = 0 To dgv.ColumnCount - 1
                    If dgv.Item(c, r).Value <> "" Then
                        flag = True
                        Exit For
                    End If
                Next
            End If

            If flag = True Then
                csvDg.Rows.Add(1)
                For c As Integer = 0 To dgv.ColumnCount - 1
                    csvDg.Item(c, r).Value = dgv.Item(c, r).Value
                Next
            End If
        Next
        '-----------------------------------------------------------------------
        frm.retDIV()
        '-----------------------------------------------------------------------

        If AutoMode = False Then
            MsgBox("移行完了", MsgBoxStyle.SystemModal)
        End If
    End Sub

    '選択行のみ
    Private Sub Csv1ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles Csv1ToolStripMenuItem1.Click
        SelRowIkou(1)
    End Sub

    Private Sub Csv2ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles Csv2ToolStripMenuItem1.Click
        SelRowIkou(2)
    End Sub

    Private Sub SelRowIkou(mode As Integer)
        Dim csvForm As Csv = Form1.CSV_FORMS(0)
        If mode = 2 Then
            csvForm = Form1.CSV_FORMS(1)
        End If

        If csvForm.Visible = False Then
            csvForm.Show()
        End If

        If csvForm.DataGridView1.RowCount = 0 Then
            For c As Integer = 0 To dgv.ColumnCount - 1
                csvForm.DataGridView1.Columns.Add("column" & c + 1, c)
            Next
        End If

        If dgv.ColumnCount <> csvForm.DataGridView1.ColumnCount Then
            MsgBox("移行先の列数が異なります", MsgBoxStyle.SystemModal)
            Exit Sub
        Else
            '-----------------------------------------------------------------------
            csvForm.PreDIV()
            '-----------------------------------------------------------------------
            Dim selRow As ArrayList = New ArrayList
            For i As Integer = 0 To dgv.SelectedCells.Count - 1
                If selRow.Contains(dgv.SelectedCells(i).RowIndex) = False Then
                    selRow.Add(dgv.SelectedCells(i).RowIndex)
                End If
            Next
            selRow.Sort()
            For r As Integer = 0 To selRow.Count - 1
                csvForm.DataGridView1.Rows.Add(1)
                For c As Integer = 0 To dgv.ColumnCount - 1
                    csvForm.DataGridView1.Item(c, csvForm.DataGridView1.RowCount - 2).Value = dgv.Item(c, selRow(r)).Value
                Next
            Next
            '-----------------------------------------------------------------------
            csvForm.RetDIV()
            '-----------------------------------------------------------------------
        End If
    End Sub

    '列取得
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button3.Click
        ColGet()
    End Sub

    Private Sub Button29_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button29.Click
        ColGet()
    End Sub

    Private Sub Button30_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button30.Click
        ColGet()
    End Sub

    Private Sub ColGet()
        If dgv.RowCount = 0 Then
            Exit Sub
        End If

        ComboBox3.Items.Clear()
        ComboBox4.Items.Clear()
        ComboBox5.Items.Clear()
        ComboBox6.Items.Clear()
        ComboBox7.Items.Clear()
        For c As Integer = 0 To dgv.ColumnCount - 1
            ComboBox3.Items.Add(dgv.Columns(c).HeaderText)
            ComboBox4.Items.Add(dgv.Columns(c).HeaderText)
            ComboBox5.Items.Add(dgv.Columns(c).HeaderText)
            ComboBox6.Items.Add(dgv.Columns(c).HeaderText)
            ComboBox7.Items.Add(dgv.Columns(c).HeaderText)
        Next

        ComboBox3.SelectedIndex = 0
        ComboBox4.SelectedIndex = 0
        ComboBox5.SelectedIndex = 0
        ComboBox6.SelectedIndex = 0
        ComboBox7.SelectedIndex = 0
    End Sub

    '一括置換
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button2.Click
        If ComboBox3.Items.Count <= 1 Then
            MsgBox("先に列取得をしてください", MsgBoxStyle.SystemModal)
            Exit Sub
        End If
        If dgv.SelectedCells.Count = 0 Then
            Exit Sub
        End If

        ToolStripProgressBar1.Value = 0
        ToolStripProgressBar1.Maximum = dgv.SelectedCells.Count

        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim countR As Integer = 0
        Dim col As Integer = ComboBox3.SelectedIndex
        For i As Integer = 0 To dgv.SelectedCells.Count - 1
            If dgv.SelectedCells(i).Value = TextBox1.Text Then
                Dim row As Integer = dgv.SelectedCells(i).RowIndex
                dgv.SelectedCells(i).Value = dgv.Item(col, row).Value
                countR += 1
            End If
            ToolStripProgressBar1.Value += 1
            Application.DoEvents()
        Next
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------

        ToolStripProgressBar1.Value = 0
        If AutoMode = False Then
            MsgBox(countR & "件を一括置換終了しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    '一括行削除
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button4.Click
        If ComboBox3.Items.Count <= 1 Then
            MsgBox("先に列取得をしてください", MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        ToolStripProgressBar1.Value = 0
        ToolStripProgressBar1.Maximum = dgv.RowCount

        Dim mode As Integer = 0
        Dim mStr As String = ""
        Dim mInt As Double = 0
        If RadioButton5.Checked = True Then         '一致
            mode = 0
        ElseIf RadioButton6.Checked = True Then     '=<
            mode = 1
        ElseIf RadioButton7.Checked = True Then     '>=
            mode = 2
        ElseIf RadioButton8.Checked = True Then     '<
            mode = 3
        ElseIf RadioButton9.Checked = True Then     '>
            mode = 4
        ElseIf RadioButton10.Checked = True Then    'like
            mode = 5
        ElseIf RadioButton3.Checked = True Then     '不一致
            mode = 6
        ElseIf RadioButton17.Checked = True Then     '不一致
            mode = 7
        End If
        mStr = Replace(TextBox2.Text, " ", "")
        If mode >= 1 And mode <= 4 Then
            If IsNumeric(mStr) = False Then
                MsgBox("数字以外を指定できません", MsgBoxStyle.SystemModal)
                Exit Sub
            Else
                mInt = CInt(mStr)
            End If
        End If

        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim countR As Integer = 0
        Dim col As Integer = ComboBox4.SelectedIndex
        Dim nowR As Integer = 0
        Dim maxR As Integer = dgv.RowCount - 1
        If dgv.AllowUserToAddRows Then
            maxR = dgv.RowCount - 2
        End If

        Dim startR As Integer = 0
        If CheckBox7.Checked = True Then
            startR = 1
        End If

        For r As Integer = maxR To startR Step -1
            Dim delFlag As Boolean = True
            'nowR = maxR - r
            Select Case mode
                Case 1
                    If IsNumeric(dgv.Item(col, r).Value) = True Then
                        If dgv.Item(col, r).Value <= mInt Then
                            delFlag = False
                        End If
                    End If
                Case 2
                    If IsNumeric(dgv.Item(col, r).Value) = True Then
                        If dgv.Item(col, r).Value >= mInt Then
                            delFlag = False
                        End If
                    End If
                Case 3
                    If IsNumeric(dgv.Item(col, r).Value) = True Then
                        If dgv.Item(col, r).Value < mInt Then
                            delFlag = False
                        End If
                    End If
                Case 4
                    If IsNumeric(dgv.Item(col, r).Value) = True Then
                        If dgv.Item(col, r).Value > mInt Then
                            delFlag = False
                        End If
                    End If
                Case 5
                    If InStr(dgv.Item(col, r).Value, mStr) > 0 Then
                        delFlag = False
                    End If
                Case 6
                    If dgv.Item(col, r).Value <> mStr Then
                        delFlag = False
                    End If
                Case 7
                    If InStr(dgv.Item(col, r).Value, mStr) > 0 Then
                        delFlag = True
                    Else
                        delFlag = False
                    End If
                Case Else
                    If dgv.Item(col, r).Value = mStr Then
                        delFlag = False
                    End If
            End Select

            If delFlag = False Then
                dgv.Rows.RemoveAt(r)
                countR += 1
            End If
            ToolStripProgressBar1.Value += 1
            Application.DoEvents()
        Next
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------

        ToolStripProgressBar1.Value = 0
        If AutoMode = False Then
            MsgBox(countR & "件を一括行削除終了しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    '一括処理
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button5.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        If RadioButton11.Checked = True Then
            For r As Integer = 0 To dgv.SelectedCells.Count - 1
                dgv.SelectedCells(r).Value = TextBox3.Text & dgv.SelectedCells(r).Value
            Next
        ElseIf RadioButton12.Checked = True Then
            For r As Integer = 0 To dgv.SelectedCells.Count - 1
                dgv.SelectedCells(r).Value = dgv.SelectedCells(r).Value & TextBox3.Text
            Next
        Else
            For r As Integer = 0 To dgv.SelectedCells.Count - 1
                dgv.SelectedCells(r).Value = Replace(dgv.SelectedCells(r).Value, TextBox3.Text, "")
            Next
        End If
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    '英数全角半角変換
    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button20.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        For i As Integer = 0 To dgv.SelectedCells.Count - 1
            If dgv.SelectedCells(i).Value <> "" Then
                dgv.SelectedCells(i).Value = Form1.StrConvNumeric(dgv.SelectedCells(i).Value)
                dgv.SelectedCells(i).Value = Form1.StrConvEnglish(dgv.SelectedCells(i).Value)
            End If
        Next
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
        MsgBox("変換終了", MsgBoxStyle.SystemModal)
    End Sub

    '置換
    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button21.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim num As Integer = 0
        For i As Integer = 0 To dgv.SelectedCells.Count - 1
            If dgv.SelectedCells(i).Value <> "" Then
                If InStr(dgv.SelectedCells(i).Value, TextBox5.Text) > 0 Then
                    num += 1
                End If
                dgv.SelectedCells(i).Value = Replace(dgv.SelectedCells(i).Value, TextBox5.Text, TextBox6.Text)
            End If
        Next
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
        MsgBox(num & "件、変換終了", MsgBoxStyle.SystemModal)
    End Sub

    '一括処理
    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button18.Click
        If ComboBox5.Items.Count <= 1 Then
            MsgBox("先に列取得をしてください", MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        Dim cA As Integer = ComboBox5.SelectedItem
        Dim cB As Integer = ComboBox6.SelectedItem
        Dim cC As Integer = ComboBox7.SelectedItem

        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        For r As Integer = 0 To dgv.RowCount - 1
            If RadioButton4.Checked = True Then
                dgv.Item(cC, r).Value = dgv.Item(cA, r).Value & TextBox4.Text & dgv.Item(cB, r).Value
            Else
                Dim sArray As String() = Split(dgv.Item(cA, r).Value, TextBox4.Text)
                For i As Integer = 0 To sArray.Length - 1
                    If i = 0 Then
                        dgv.Item(cB, r).Value = sArray(i)
                    ElseIf i = 1 Then
                        dgv.Item(cC, r).Value = sArray(i)
                    Else
                        dgv.Item(cC, r).Value &= sArray(i)
                    End If
                Next
            End If
        Next
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub RadioButton4_CheckedChanged(ByVal sender As System.Object, ByVal e As EventArgs) Handles RadioButton4.CheckedChanged
        Label10.Text = "列1"
        Label11.Text = "列2"
        Label12.Text = "結果"
    End Sub

    Private Sub RadioButton14_CheckedChanged(ByVal sender As System.Object, ByVal e As EventArgs) Handles RadioButton14.CheckedChanged
        Label10.Text = "列1"
        Label11.Text = "結果1"
        Label12.Text = "結果2"
    End Sub

    '縦横変換
    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button16.Click
        If dgv.RowCount > 500 Then
            MsgBox("500行以上のデータを縦横変換できません", MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim rStr As ArrayList = New ArrayList
        For r As Integer = 0 To dgv.RowCount - 1
            Dim aStr As String = ""
            For c As Integer = 0 To dgv.ColumnCount - 1
                If aStr = "" Then
                    aStr = dgv.Item(c, r).Value
                Else
                    aStr &= "|=|" & dgv.Item(c, r).Value
                End If
            Next
            rStr.Add(aStr)
        Next
        Dim gyou As Integer = dgv.ColumnCount

        dgv.Rows.Clear()
        dgv.Columns.Clear()

        For c As Integer = 0 To rStr.Count - 1
            dgv.Columns.Add(c, c)
            dgv.Columns(c).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        dgv.Rows.Add(gyou)

        For c As Integer = 0 To rStr.Count - 1
            Dim tStr As String() = Split(rStr(c), "|=|")
            For r As Integer = 0 To tStr.Length - 1
                dgv.Item(c, r).Value = tStr(r)
            Next
        Next
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    '移動↓
    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button13.Click
        Ido(DataGridView1, DataGridView5)
    End Sub

    '移動↑
    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button14.Click
        Ido(DataGridView5, DataGridView1)
    End Sub

    Private Sub Ido(ByVal dgv1 As DataGridView, ByVal dgv2 As DataGridView)
        If dgv2.RowCount > 0 Then
            Dim DR As DialogResult = MsgBox("移動しますか？", MsgBoxStyle.OkCancel Or MsgBoxStyle.SystemModal)
            If DR = Windows.Forms.DialogResult.Cancel Then
                Exit Sub
            End If
        End If

        '-----------------------------------------------------------------------
        dgv2.ScrollBars = ScrollBars.None
        dgv2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgv2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
        dgv2.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        dgv2.ScrollBars = ScrollBars.Both
        '-----------------------------------------------------------------------
        dgv2.Rows.Clear()
        dgv2.Columns.Clear()

        For c As Integer = 0 To dgv1.ColumnCount - 1
            dgv2.Columns.Add(c, c)
            dgv2.Columns(c).SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        For r As Integer = 0 To dgv1.RowCount - 1
            dgv2.Rows.Add(1)
            For c As Integer = 0 To dgv1.ColumnCount - 1
                If dgv1.Item(c, r).Value <> "" Then
                    dgv2.Item(c, r).Value = dgv1.Item(c, r).Value
                End If
            Next
        Next
        '-----------------------------------------------------------------------
        dgv2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        '-----------------------------------------------------------------------
        If AutoMode = False Then
            MsgBox("移動しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    '↓交換↑
    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button15.Click
        Dim dgvArray1 As ArrayList = New ArrayList
        Dim dgvArray2 As ArrayList = New ArrayList
        Dim d1col As Integer = DataGridView1.ColumnCount
        Dim d2col As Integer = DataGridView5.ColumnCount

        For i As Integer = 0 To 1
            Dim dgvA As DataGridView
            If i = 0 Then
                dgvA = DataGridView1
            Else
                dgvA = DataGridView5
            End If
            '-----------------------------------------------------------------------
            dgvA.ScrollBars = ScrollBars.None
            dgvA.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgvA.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            dgvA.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            dgvA.ScrollBars = ScrollBars.Both
            '-----------------------------------------------------------------------
            For r As Integer = 0 To dgvA.RowCount - 1
                Dim str As String = ""
                For c As Integer = 0 To dgvA.ColumnCount - 1
                    If str = "" Then
                        str = dgvA.Item(c, r).Value
                    Else
                        str &= "|=|" & dgvA.Item(c, r).Value
                    End If
                Next
                If i = 0 Then
                    dgvArray1.Add(str)
                Else
                    dgvArray2.Add(str)
                End If
            Next
        Next

        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        DataGridView5.Rows.Clear()
        DataGridView5.Columns.Clear()

        For c As Integer = 0 To d1col - 1
            DataGridView5.Columns.Add(c, c)
            DataGridView5.Columns(c).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        For c As Integer = 0 To d2col - 1
            DataGridView1.Columns.Add(c, c)
            DataGridView1.Columns(c).SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        ToolStripProgressBar1.Value = 0
        ToolStripProgressBar1.Maximum = dgvArray1.Count
        For r As Integer = 0 To dgvArray1.Count - 1
            DataGridView5.Rows.Add(1)
            Dim res As String() = Split(dgvArray1(r), "|=|")
            For c As Integer = 0 To res.Length - 1
                If res(c) <> "" Then
                    DataGridView5.Item(c, r).Value = res(c)
                End If
            Next
            ToolStripProgressBar1.Value += 1
        Next
        ToolStripProgressBar1.Value = 0
        ToolStripProgressBar1.Maximum = dgvArray2.Count
        For r As Integer = 0 To dgvArray2.Count - 1
            DataGridView1.Rows.Add(1)
            Dim res As String() = Split(dgvArray2(r), "|=|")
            For c As Integer = 0 To res.Length - 1
                If res(c) <> "" Then
                    DataGridView1.Item(c, r).Value = res(c)
                End If
            Next
            ToolStripProgressBar1.Value += 1
        Next
        ToolStripProgressBar1.Value = 0
        '-----------------------------------------------------------------------
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        DataGridView5.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        '-----------------------------------------------------------------------
        If AutoMode = False Then
            MsgBox("交換しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    '↓処理
    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button11.Click
        If TabControl1.SelectedIndex = 1 Then
            Panel4.BackColor = Color.Yellow
            Panel5.BackColor = Color.Gainsboro
            dgv = DataGridView1
            Dim csvRecords As New ArrayList()
            For r As Integer = 0 To DataGridView1.RowCount - 1
                Dim str As String = ""
                For c As Integer = 0 To DataGridView1.ColumnCount - 1
                    If c = 0 Then
                        str = DataGridView1.Item(c, r).Value
                    Else
                        str &= "|=|" & DataGridView1.Item(c, r).Value
                    End If
                Next
                csvRecords.Add(str)
            Next
            Panel5.BackColor = Color.Yellow
            Panel4.BackColor = Color.Gainsboro
            dgv = DataGridView5
            CSV_GAPPEI(csvRecords)
        Else
            MsgBox("処理ボタンで処理してください", MsgBoxStyle.SystemModal)
            Exit Sub
        End If
    End Sub

    '処理↑
    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button12.Click
        If TabControl1.SelectedIndex = 1 Then
            Panel5.BackColor = Color.Yellow
            Panel4.BackColor = Color.Gainsboro
            dgv = DataGridView5
            Dim csvRecords As New ArrayList()
            For r As Integer = 0 To DataGridView5.RowCount - 1
                Dim str As String = ""
                For c As Integer = 0 To DataGridView5.ColumnCount - 1
                    If c = 0 Then
                        str = DataGridView5.Item(c, r).Value
                    Else
                        str &= "|=|" & DataGridView5.Item(c, r).Value
                    End If
                Next
                csvRecords.Add(str)
            Next
            Panel4.BackColor = Color.Yellow
            Panel5.BackColor = Color.Gainsboro
            dgv = DataGridView1
            CSV_GAPPEI(csvRecords)
        Else
            MsgBox("処理ボタンで処理してください", MsgBoxStyle.SystemModal)
            Exit Sub
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        If TabControl1.SelectedIndex = 1 Then
            Panel5.BackColor = Color.Yellow
            Panel4.BackColor = Color.Gainsboro
            dgv = DataGridView5
            Dim csvRecords As New ArrayList()
            For r As Integer = 0 To DataGridView5.RowCount - 1
                Dim str As String = ""
                For c As Integer = 0 To DataGridView5.ColumnCount - 1
                    If c = 0 Then
                        str = DataGridView5.Item(c, r).Value
                    Else
                        str &= "|=|" & DataGridView5.Item(c, r).Value
                    End If
                Next
                csvRecords.Add(str)
            Next
            Panel4.BackColor = Color.Yellow
            Panel5.BackColor = Color.Gainsboro
            dgv = DataGridView1
            CSV_GAPPEI(csvRecords, 1)
        Else
            MsgBox("処理ボタンで処理してください", MsgBoxStyle.SystemModal)
            Exit Sub
        End If
    End Sub

    Private Sub 列を削除ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles 列を削除ToolStripMenuItem.Click
        Panel4.BackColor = Color.Yellow
        Panel5.BackColor = Color.Gainsboro
        dgv = DataGridView1
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim selCell = dgv.SelectedCells
        COLSCUT(dgv, selCell)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles ToolStripMenuItem1.Click
        Panel5.BackColor = Color.Yellow
        Panel4.BackColor = Color.Gainsboro
        dgv = DataGridView5
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim selCell = dgv.SelectedCells
        COLSCUT(dgv, selCell)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
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

    Private Sub 行を削除ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles 行を削除ToolStripMenuItem.Click
        Panel5.BackColor = Color.Yellow
        Panel4.BackColor = Color.Gainsboro
        dgv = DataGridView1
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim selCell = dgv.SelectedCells
        ROWSCUT(dgv, selCell)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 行を削除ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles 行を削除ToolStripMenuItem1.Click
        Panel5.BackColor = Color.Yellow
        Panel4.BackColor = Color.Gainsboro
        dgv = DataGridView5
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim selCell = dgv.SelectedCells
        ROWSCUT(dgv, selCell)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub ROWSCUT(ByVal dgv As DataGridView, ByVal selCell As Object)
        Dim delR As New ArrayList
        For i As Integer = 0 To selCell.Count - 1
            If Not delR.Contains(selCell(i).RowIndex) Then
                delR.Add(selCell(i).RowIndex)
            End If
        Next
        delR.Sort()
        delR.Reverse()
        selChangeFlag = False
        For r As Integer = 0 To delR.Count - 1
            If delR(r) <= dgv.RowCount - 1 Then
                dgv.Rows.RemoveAt(delR(r))
            End If
        Next
        selChangeFlag = True
    End Sub

    Private Sub 左に追加ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles 左に追加ToolStripMenuItem1.Click
        Panel4.BackColor = Color.Yellow
        Panel5.BackColor = Color.Gainsboro
        dgv = DataGridView1
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim selCell = dgv.SelectedCells
        COLSADD(dgv, selCell, 0)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 右に追加ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles 右に追加ToolStripMenuItem1.Click
        Panel4.BackColor = Color.Yellow
        Panel5.BackColor = Color.Gainsboro
        dgv = DataGridView1
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim selCell = dgv.SelectedCells
        COLSADD(dgv, selCell, 1)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 左に追加ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles 左に追加ToolStripMenuItem.Click
        Panel5.BackColor = Color.Yellow
        Panel4.BackColor = Color.Gainsboro
        dgv = DataGridView5
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim selCell = dgv.SelectedCells
        COLSADD(dgv, selCell, 0)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 右に追加ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles 右に追加ToolStripMenuItem.Click
        Panel5.BackColor = Color.Yellow
        Panel4.BackColor = Color.Gainsboro
        dgv = DataGridView5
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim selCell = dgv.SelectedCells
        COLSADD(dgv, selCell, 1)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub COLSADD(ByVal dgv As DataGridView, ByVal selCell As Object, ByVal num As Integer)
        Dim addC As New ArrayList
        For i As Integer = 0 To selCell.Count - 1
            If Not addC.Contains(selCell(i).ColumnIndex) Then
                addC.Add(selCell(i).ColumnIndex)
            End If
        Next
        For c As Integer = 0 To addC.Count - 1
            If num = 0 Then
                Dim TextBoxColumn As New DataGridViewTextBoxColumn With {
                    .Name = dgv.ColumnCount,
                    .HeaderText = dgv.ColumnCount
                }
                Dim colNum As Integer = addC(c)
                dgv.Columns.Insert(colNum, TextBoxColumn)
                dgv.Columns(colNum).SortMode = DataGridViewColumnSortMode.NotSortable
            Else
                If addC(c) <> dgv.ColumnCount - 1 Then
                    Dim TextBoxColumn As New DataGridViewTextBoxColumn With {
                        .Name = dgv.ColumnCount,
                        .HeaderText = dgv.ColumnCount
                    }
                    Dim colNum As Integer = addC(c) + 1
                    dgv.Columns.Insert(colNum, TextBoxColumn)
                    dgv.Columns(colNum).SortMode = DataGridViewColumnSortMode.NotSortable
                Else
                    Dim colNum As Integer = dgv.ColumnCount
                    dgv.Columns.Add(colNum, colNum)
                    dgv.Columns(colNum).SortMode = DataGridViewColumnSortMode.NotSortable
                End If
            End If
        Next
    End Sub

    Private Sub 上に挿入ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles 上に挿入ToolStripMenuItem.Click
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

    Private Sub 下に挿入ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles 下に挿入ToolStripMenuItem.Click
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

    Private Sub 値を削除ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles 値を削除ToolStripMenuItem.Click
        Panel4.BackColor = Color.Yellow
        Panel5.BackColor = Color.Gainsboro
        dgv = DataGridView1
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim selCell = dgv.SelectedCells
        DELS(dgv, selCell)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 行を選択直下に複製ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles 行を選択直下に複製ToolStripMenuItem.Click
        Panel4.BackColor = Color.Yellow
        Panel5.BackColor = Color.Gainsboro
        dgv = DataGridView1
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim selR As Integer = dgv.SelectedCells(0).RowIndex
        dgv.Rows.InsertCopy(selR, selR + 1)
        For c As Integer = 0 To dgv.ColumnCount - 1
            If dgv.Item(c, selR).Value <> "" Then
                dgv.Item(c, selR + 1).Value = dgv.Item(c, selR).Value
            End If
        Next
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub ToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles ToolStripMenuItem2.Click
        Panel5.BackColor = Color.Yellow
        Panel4.BackColor = Color.Gainsboro
        dgv = DataGridView5
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim selCell = dgv.SelectedCells
        DELS(dgv, selCell)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub DELS(ByVal dgv As DataGridView, ByVal selCell As Object)
        For i As Integer = 0 To selCell.Count - 1
            dgv(selCell(i).ColumnIndex, selCell(i).RowIndex).Value = ""
        Next
    End Sub

    '重複削除
    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button19.Click
        JuuhukuDel()
    End Sub

    Private Sub JuuhukuDel()
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim zanArray As ArrayList = New ArrayList
        Dim delArray As ArrayList = New ArrayList
        For i As Integer = 0 To dgv.SelectedCells.Count - 1
            If zanArray.Contains(dgv.SelectedCells(i).Value) = True Then
                Dim dRow As Integer = dgv.SelectedCells(i).RowIndex
                If delArray.Contains(dRow) = False Then
                    delArray.Add(dRow)
                End If
            End If
            zanArray.Add(dgv.SelectedCells(i).Value)
        Next
        delArray.Sort()
        delArray.Reverse()
        For i As Integer = 0 To delArray.Count - 1
            dgv.Rows.RemoveAt(delArray(i))
        Next
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            SplitContainer2.SplitterDistance = SplitContainer2.Height * 0.65
            Button11.Visible = True
            Button12.Visible = True
            Button13.Visible = True
            Button14.Visible = True
            Button15.Visible = True
        Else
            SplitContainer2.SplitterDistance = SplitContainer2.Height
            Button11.Visible = False
            Button12.Visible = False
            Button13.Visible = False
            Button14.Visible = False
            Button15.Visible = False
        End If
    End Sub

    Private Sub 佐川急便ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles 佐川急便ToolStripMenuItem.Click
        OtherDialog.Size = New Size(199, 199)
        OtherDialog.TabControl1.SelectedIndex = 1
        OtherDialog.Show()
    End Sub

    Private Sub 伝票表示ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles 伝票表示ToolStripMenuItem.Click
        OtherDialog.Size = New Size(745, 442)
        OtherDialog.TabControl1.SelectedIndex = 2
        OtherDialog.Show()
        OtherDialog.dgvWrite = dgv
        OtherDialog.CsvToDenpyo(dgv)
    End Sub

    'Private Sub CheckBox5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox5.CheckedChanged
    '    If CheckBox5.Checked = True Then
    '        Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\会社情報.txt"
    '        Dim flag As Boolean = False
    '        CheckedListBox1.Items.Clear()
    '        Using sr As New System.IO.StreamReader(fName, System.Text.Encoding.Default)
    '            Do While Not sr.EndOfStream
    '                Dim s As String = sr.ReadLine
    '                If s = "//店舗リスト" Then
    '                    flag = True
    '                ElseIf s = "" Then

    '                ElseIf flag = True Then
    '                    Dim tArray As String() = Split(s, "=")
    '                    CheckedListBox1.Items.Add(tArray(0))
    '                End If
    '            Loop
    '        End Using
    '        Panel6.Location = New Point(160, 40)
    '        Panel6.Visible = True

    '        '*********************************************
    '        Dim iniStr As String = Form1.inif.ReadINI(secValue, "csvCList", "")
    '        Dim sArray As String() = Split(iniStr, splitter)
    '        '*********************************************
    '        'Dim sArray As String() = Split(My.Settings.csvCList, splitter)
    '        For i As Integer = 0 To CheckedListBox1.Items.Count - 1
    '            Dim res As String = CheckedListBox1.Items(i)
    '            If sArray.Contains(res) = True Then
    '                CheckedListBox1.SetItemChecked(i, True)
    '            End If
    '        Next
    '    Else
    '        Panel6.Visible = False
    '    End If
    'End Sub

    'Private Sub CheckedListBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckedListBox1.SelectedIndexChanged
    '    Dim str As String = ""
    '    For Each clb As String In CheckedListBox1.CheckedItems
    '        If str = "" Then
    '            str = clb
    '        Else
    '            str &= splitter & clb
    '        End If
    '    Next
    '    '*********************************************
    '    Form1.inif.WriteINI(secValue, "csvCList", str)
    '    '*********************************************
    '    'My.Settings.csvCList = str
    'End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As EventArgs) Handles CheckBox2.CheckedChanged
        '*********************************************
        Form1.inif.WriteINI(secValue, "yupackCheck", CheckBox2.Checked)
        '*********************************************
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles TabControl1.SelectedIndexChanged
        If dgv.RowCount > 0 Then
            ComboBox10.Items.Clear()
            For c As Integer = 0 To dgv.ColumnCount - 1
                ComboBox10.Items.Add(dgv.Columns(c).HeaderText)
            Next
            ComboBox10.SelectedIndex = 0
        End If
    End Sub

    '抽出実行
    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button23.Click
        If CheckBox9.Checked = True Then    '英数半角変換
            TextBox16.Text = StrConv(TextBox16.Text, VbStrConv.Lowercase)
        End If
        If CheckBox11.Checked = True Then    '全角半角変換
            TextBox16.Text = Form1.StrConvNumeric(TextBox16.Text)
            TextBox16.Text = Form1.StrConvEnglish(TextBox16.Text)
        End If

        Dim colA As Integer = ComboBox10.SelectedIndex
        Dim searchA As String() = Split(TextBox16.Text, vbCrLf)

        Dim header As Integer = False
        If CheckBox10.Checked = True Then
            If dgv.RowCount > 0 Then
                dgv.Rows(0).Visible = True
                header = True
            End If
        End If

        For r As Integer = 0 To dgv.RowCount - 1
            If header = True And r <> 0 Then
                Dim vFlag As Boolean = False
                For i As Integer = 0 To searchA.Length - 1
                    If searchA(i) <> "" Then
                        If CB3_1.Checked = True Then
                            If InStr(dgv.Item(colA, r).Value, searchA(i)) > 0 Then
                                dgv.Rows(r).Visible = True
                                vFlag = True
                                Exit For
                            End If
                        Else
                            If dgv.Item(colA, r).Value = searchA(i) Then
                                dgv.Rows(r).Visible = True
                                vFlag = True
                                Exit For
                            End If
                        End If
                    End If
                Next
                If vFlag = False Then
                    dgv.Rows(r).Visible = False
                End If
            End If
        Next

        selChangeFlag = False
        Dim maxRow As Integer = dgv.RowCount - 1
        Dim nowRow As Integer = 0
        For r As Integer = 0 To maxRow
            nowRow = maxRow - r
            If dgv.Rows(nowRow).Visible = False Then
                dgv.Rows.RemoveAt(nowRow)
            End If
        Next
        selChangeFlag = True
    End Sub

    '削除実行
    Private Sub Button31_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button31.Click
        If CheckBox9.Checked = True Then    '英数半角変換
            TextBox16.Text = StrConv(TextBox16.Text, VbStrConv.Lowercase)
        End If
        If CheckBox11.Checked = True Then    '全角半角変換
            TextBox16.Text = Form1.StrConvNumeric(TextBox16.Text)
            TextBox16.Text = Form1.StrConvEnglish(TextBox16.Text)
        End If

        Dim colA As Integer = ComboBox10.SelectedIndex
        Dim searchA As String() = Split(TextBox16.Text, vbCrLf)

        Dim header As Integer = False
        If CheckBox10.Checked = True Then
            If dgv.RowCount > 0 Then
                dgv.Rows(0).Visible = True
                header = True
            End If
        End If

        For r As Integer = 0 To dgv.RowCount - 1
            If header = True And r <> 0 Then
                Dim vFlag As Boolean = True

                If CB3_1.Checked = True Then
                    For i As Integer = 0 To searchA.Length - 1
                        If searchA(i) <> "" Then
                            If CB3_1.Checked = True Then
                                If InStr(dgv.Item(colA, r).Value, searchA(i)) > 0 Then
                                    dgv.Rows(r).Visible = False
                                    vFlag = False
                                    Exit For
                                End If
                            Else
                                If dgv.Item(colA, r).Value = searchA(i) Then
                                    dgv.Rows(r).Visible = False
                                    vFlag = False
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                Else
                    Dim a = dgv.Item(colA, r).Value
                    If searchA.Contains(dgv.Item(colA, r).Value) Then
                        vFlag = False
                    End If
                End If

                If vFlag = True Then
                    dgv.Rows(r).Visible = True
                Else
                    dgv.Rows(r).Visible = False
                End If
            End If
        Next

        Dim maxRow As Integer = dgv.RowCount - 1
        Dim nowRow As Integer = 0
        For r As Integer = 0 To maxRow
            nowRow = maxRow - r
            If dgv.Rows(nowRow).Visible = False Then
                dgv.Rows.RemoveAt(nowRow)
            End If
        Next
    End Sub

    Private Sub TextBox16_KeyUp(ByVal sender As Object, ByVal e As KeyEventArgs) Handles TextBox16.KeyUp

    End Sub

    '=======================================================
    'プログラム処理
    '=======================================================
    Private Sub ComboBox8_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As EventArgs) Handles ComboBox8.SelectedIndexChanged
        If ComboBox8.SelectedIndex = 0 Then
            Exit Sub
        End If

        DataGridView6.Rows.Clear()

        'ファイル読み取り
        Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\template\" & ComboBox8.SelectedItem & ".dat"
        Using sr As New StreamReader(fName, System.Text.Encoding.Default)
            Do While Not sr.EndOfStream
                Dim s As String = sr.ReadLine
                If InStr(s, "//") > 0 Then
                    DataGridView6.Rows.Add(1)
                    DataGridView6.Item(0, DataGridView6.RowCount - 2).Value = "コメント"
                    DataGridView6.Item(1, DataGridView6.RowCount - 2).Value = s
                Else
                    Dim sArray As String() = Split(s, splitter)
                    DataGridView6.Rows.Add(sArray)
                End If
            Loop
        End Using
    End Sub

    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button25.Click
        Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\template\" & ComboBox8.Text & ".dat"
        If File.Exists(fName) = True Then
            MsgBox("同じ名前のファイルがあります", MsgBoxStyle.SystemModal)
            Exit Sub
        ElseIf InStr(fName, "Program") = 0 Then
            MsgBox("ファイル名に「Program」を入れてください", MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        Dim str As String = ""
        For r As Integer = 0 To DataGridView6.RowCount - 2
            str &= DataGridView6.Item(0, r).Value & splitter
            str &= DataGridView6.Item(1, r).Value & splitter
            str &= DataGridView6.Item(2, r).Value & splitter
            str &= DataGridView6.Item(3, r).Value & vbCrLf
        Next

        Dim sw As StreamWriter = New StreamWriter(fName, False, System.Text.Encoding.GetEncoding("SHIFT-JIS"))
        sw.Write(str)
        sw.Close()
        ReloadCombobox()

        MsgBox("保存しました", MsgBoxStyle.SystemModal)
    End Sub

    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button24.Click
        If ComboBox8.SelectedIndex = 0 Then
            Exit Sub
        End If

        Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\template\" & ComboBox8.SelectedItem & ".dat"
        File.Delete(fName)
        ReloadCombobox()
    End Sub

    Private Sub Button28_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button28.Click
        TextBox7.Text = 0
        Panel8.Visible = False
        MsgBox("リセットしました", MsgBoxStyle.SystemModal)
    End Sub

    Private Sub Button26_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button26.Click
        RunProgram()
    End Sub

    Private Sub Button27_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button27.Click
        Panel8.Visible = False
        TextBox8.Text = ""
        TextBox8.BackColor = Color.WhiteSmoke
    End Sub

    '処理
    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button22.Click
        RunProgram()
    End Sub

    Private Sub RunProgram()
        Dim runRow As Integer = 0
        If TextBox7.Text = "" Then
            TextBox7.Text = 0
        End If
        If TextBox7.Text <= DataGridView6.RowCount - 2 Then
            runRow = TextBox7.Text
            Panel8.Visible = False
            TextBox8.Text = "実行中"
            TextBox8.BackColor = Color.Aqua
            Application.DoEvents()
        Else
            MsgBox("これ以上プログラムが実行できません", MsgBoxStyle.SystemModal)
            Panel8.Visible = False
            TextBox8.Text = ""
            TextBox8.BackColor = Color.WhiteSmoke
            Exit Sub
        End If

        Dim runFlag As Boolean = True
        For r As Integer = runRow To DataGridView6.RowCount - 2
            If InStr(DataGridView6.Item(0, r).Value, "コメント") = 0 Then
                Select Case DataGridView6.Item(0, r).Value
                    Case "メッセージ"
                        MsgBox(DataGridView6.Item(2, r).Value, MsgBoxStyle.SystemModal)
                    Case "表示"
                        PText(r)
                    Case "ダイアログ"
                        PDialog1()
                        runRow += 1
                        runFlag = False
                    Case "タブ"
                        PTabSelect(r)
                    Case "コントロール"
                        PControl(r)
                    Case "グリッド"
                        PGrid(r)
                    Case "行追加"
                        PRowAdd(r)
                    Case "終了"
                        PEnd()
                        Exit For
                End Select

                runRow += 1
                TextBox7.Text = runRow
                Application.DoEvents()
                If runFlag = False Then
                    Exit For
                End If
            End If
        Next
    End Sub

    Private Sub PText(ByVal r As Integer)
        TextBox9.Text = DataGridView6.Item(2, r).Value
    End Sub

    Private Sub PDialog1()
        Panel8.Visible = True
        TextBox8.Text = "待機中"
        TextBox8.BackColor = Color.Yellow
    End Sub

    Private Sub PTabSelect(ByVal r As Integer)
        For i As Integer = 0 To TabControl1.TabCount - 1
            If TabControl1.TabPages(i).Text = DataGridView6.Item(1, r).Value Then
                TabControl1.SelectedIndex = i
                Exit For
            End If
        Next
    End Sub

    Private Sub PControl(ByVal r As Integer)
        Select Case DataGridView6.Item(1, r).Value
            Case "TB1_1"    '特殊処理1
                Button1.PerformClick()
            Case "TB1_2"
                Button16.PerformClick()
            Case "TB1_3"
                Button3.PerformClick()
            Case "TB1_4"
                If DataGridView6.Item(2, r).Value = "clear" Then
                    TextBox1.Clear()
                Else
                    TextBox1.Text = Replace(DataGridView6.Item(2, r).Value, "|", vbCrLf)
                End If
            Case "TB1_5"
                ComboBox3.SelectedIndex = DataGridView6.Item(2, r).Value
            Case "TB1_6"
                Button2.PerformClick()



            Case "TB3_1"    '特殊処理3
                CB3_1.Checked = DataGridView6.Item(2, r).Value
            Case "TB3_2"
                CheckBox10.Checked = DataGridView6.Item(2, r).Value
            Case "TB3_3"
                ComboBox10.SelectedIndex = DataGridView6.Item(2, r).Value
            Case "TB3_4"
                Button23.PerformClick()
                Application.DoEvents()
            Case "TB3_5"
                If DataGridView6.Item(2, r).Value = "clear" Then
                    TextBox16.Clear()
                Else
                    TextBox16.Text = Replace(DataGridView6.Item(2, r).Value, "|", vbCrLf)
                End If
        End Select
    End Sub

    Private Sub PGrid(ByVal r As Integer)
        Select Case DataGridView6.Item(1, r).Value
            Case "列抽出"
                Dim gArray As String() = Split(DataGridView6.Item(2, r).Value, "|")
                For c As Integer = dgv.ColumnCount - 1 To 0 Step -1
                    If gArray.Contains(dgv.Item(c, 0).Value) = False Then
                        dgv.Columns.RemoveAt(c)
                    End If
                Next
            Case "検索"
                Dim gArray1 As String() = Split(DataGridView6.Item(2, r).Value, "=")
                Dim gArray2 As String() = Split(DataGridView6.Item(3, r).Value, "|")

                Dim dHeader As ArrayList = New ArrayList
                For c As Integer = 0 To dgv.ColumnCount - 1
                    dHeader.Add(dgv.Item(c, 0).Value)
                Next

                Dim header1 As Integer = dHeader.IndexOf(gArray1(0))
                Dim header2 As Integer = -1
                Dim value2 As String = ""
                Dim header3 As Integer = -1
                Dim value3 As String = ""
                If InStr(gArray2(0), "=") > 0 Then
                    Dim array As String() = Split(gArray2(0), "=")
                    header2 = dHeader.IndexOf(array(0))
                    value2 = array(1)
                End If
                If InStr(gArray2(1), "=") > 0 Then
                    Dim array As String() = Split(gArray2(1), "=")
                    header3 = dHeader.IndexOf(array(0))
                    value3 = array(1)
                End If

                For row As Integer = 1 To dgv.RowCount - 1
                    If InStr(dgv.Item(header1, row).Value, gArray1(1)) Then
                        If header2 <> -1 Then
                            dgv.Item(header2, row).Value = value2
                        End If
                    Else
                        If header3 <> -1 Then
                            dgv.Item(header3, row).Value = value3
                        End If
                    End If
                Next
            Case "追加"
                Dim gArray1 As String() = Split(DataGridView6.Item(2, r).Value, "=")
                Dim tStr As String = gArray1(1)
                tStr = Replace(tStr, "\n", vbLf)
                Dim gStr As String = DataGridView6.Item(3, r).Value

                Dim dHeader As ArrayList = New ArrayList
                For c As Integer = 0 To dgv.ColumnCount - 1
                    dHeader.Add(dgv.Item(c, 0).Value)
                Next

                Dim header1 As Integer = dHeader.IndexOf(gArray1(0))
                Dim value1 As String = ""

                For row As Integer = 1 To dgv.RowCount - 1
                    If gStr = "前" Then
                        dgv.Item(header1, row).Value = tStr & dgv.Item(header1, row).Value
                    Else
                        dgv.Item(header1, row).Value = dgv.Item(header1, row).Value & tStr
                    End If
                Next
            Case "削除" '使用できないかも（不完全）
                Dim gArray1 As String() = Split(DataGridView6.Item(2, r).Value, "=")
                Dim gStr As String = DataGridView6.Item(3, r).Value

                Dim dHeader As ArrayList = New ArrayList
                For c As Integer = 0 To dgv.ColumnCount - 1
                    dHeader.Add(dgv.Item(c, 0).Value)
                Next

                Dim header1 As Integer = dHeader.IndexOf(gArray1(0))
                Dim value1 As String = ""

                For row As Integer = 1 To dgv.RowCount - 1
                    If gStr = "行" Then
                        Dim str As String = dgv.Item(header1, row).Value
                        Dim str2 As String = ""
                        str = Replace(str, vbCr, vbLf)
                        str = Replace(str, vbCrLf, vbLf)
                        Dim lines As String() = Split(str, vbLf)
                        For i As Integer = 0 To lines.Length - 1
                            If InStr(lines(i), gArray1(1)) = 0 Then
                                If str2 = "" Then
                                    str2 = lines(i)
                                Else
                                    str2 = vbLf & lines(i)
                                End If
                            End If
                        Next
                    Else
                        dgv.Item(header1, row).Value = Regex.Replace(dgv.Item(header1, row).Value, gArray1(1), gArray1(2))
                    End If
                Next
            Case "行複製"
                Dim gArray1 As String() = Split(DataGridView6.Item(2, r).Value, "=")
                Dim gArray2 As String() = Split(DataGridView6.Item(3, r).Value, "|")

                Dim dHeader As ArrayList = New ArrayList
                For c As Integer = 0 To dgv.ColumnCount - 1
                    dHeader.Add(dgv.Item(c, 0).Value)
                Next

                Dim header1 As Integer = dHeader.IndexOf(gArray1(0))

                For row As Integer = dgv.RowCount - 1 To 1 Step -1
                    If InStr(dgv.Item(header1, row).Value, gArray1(1)) Then
                        dgv.Rows.InsertCopy(row, row + 1)   '複製でコピーできていないので下記でコピー
                        For c As Integer = 0 To dgv.ColumnCount - 1
                            If dgv.Item(c, row).Value <> "" Then
                                dgv.Item(c, row + 1).Value = dgv.Item(c, row).Value
                            End If
                        Next
                        '複製した行のカラムを書き換える
                        For i As Integer = 0 To gArray2.Length - 1
                            Dim cArray As String() = Split(gArray2(i), "=")
                            Dim header2 As Integer = dHeader.IndexOf(cArray(0))
                            dgv.Item(header2, row + 1).Value = cArray(1)
                        Next
                    End If
                Next
        End Select
    End Sub

    Private Sub PRowAdd(ByVal r As Integer)

    End Sub

    Private Sub PEnd()
        For i As Integer = 0 To TabControl1.TabCount - 1
            If TabControl1.TabPages(i).Text = "プログラム処理" Then
                TabControl1.SelectedIndex = i
                Exit For
            End If
        Next
        TextBox7.Text = 0
        Panel8.Visible = False
        TextBox8.Text = ""
        TextBox8.BackColor = Color.WhiteSmoke
        MsgBox("終了しました", MsgBoxStyle.SystemModal)
    End Sub

    Dim selChangeFlag As Boolean = True
    Dim BeforeColor
    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs) Handles DataGridView1.SelectionChanged
        If selChangeFlag = False Then
            Exit Sub
        End If

        If DataGridView1.SelectedCells.Count > 50 Then
            Exit Sub
        End If

        If DataGridView1.SelectedCells.Count > 0 Then
            TextBox17.Text = DataGridView1.SelectedCells(0).Value
        End If

        '-------------------------------------------------------------------------------------
        '選択行の色付
        Dim dgv As DataGridView = DataGridView1
        If dgv.SelectedCells.Count > 50 Then
            Exit Sub
        End If

        For r As Integer = 0 To dgv.RowCount - 1
            DataGridView1.Rows(r).DefaultCellStyle.BackColor = BeforeColor
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
        '-------------------------------------------------------------------------------------

        If CheckBox12.Checked = False Or TabControl1.SelectedTab.Text <> "伝票合成" Then
            Exit Sub
        ElseIf DataGridView1.RowCount = 0 Or DataGridView5.RowCount = 0 Then
            Exit Sub
        ElseIf DataGridView1.SelectedCells.Count = 0 Then
            Exit Sub
        End If

        'データのチェック

        'datagridviewのヘッダーテキストをコレクションに取り込む
        Dim DGVheaderCheck1 As ArrayList = New ArrayList
        For c As Integer = 0 To DataGridView1.Columns.Count - 1
            DGVheaderCheck1.Add(DataGridView1.Item(c, 0).Value)
        Next c
        Dim DGVheaderCheck2 As ArrayList = New ArrayList
        For c As Integer = 0 To DataGridView5.Columns.Count - 1
            DGVheaderCheck2.Add(DataGridView5.Item(c, 0).Value)
        Next c

        Dim selR As Integer = DataGridView1.SelectedCells(0).RowIndex
        Dim denpyoNum As String = ""
        If DGVheaderCheck1.Contains("伝票no") Then
            denpyoNum = DataGridView1.Item(DGVheaderCheck1.IndexOf("伝票no"), selR).Value
        ElseIf DGVheaderCheck1.Contains("伝票番号") Then
            denpyoNum = DataGridView1.Item(DGVheaderCheck1.IndexOf("伝票番号"), selR).Value
        ElseIf DGVheaderCheck1.Contains("注文日") Then
            denpyoNum = DataGridView1.Item(DGVheaderCheck1.IndexOf("注文日"), selR).Value
        End If

        DataGridView8.Rows.Clear()
        For r As Integer = 0 To DataGridView5.RowCount - 1
            Dim sDenpyoNum As String = DataGridView5.Item(DGVheaderCheck2.IndexOf("伝票番号"), r).Value
            If denpyoNum = sDenpyoNum Then
                Dim sCode As String = DataGridView5.Item(DGVheaderCheck2.IndexOf("商品ｺｰﾄﾞ"), r).Value
                Dim sCount As String = DataGridView5.Item(DGVheaderCheck2.IndexOf("受注数"), r).Value
                DataGridView8.Rows.Add(sCode, sCount)
            End If
        Next
    End Sub



    Private Sub Button32_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button32.Click
        If TabControl1.SelectedTab.Text <> "伝票合成" Then
            Exit Sub
        ElseIf DataGridView1.RowCount = 0 Or DataGridView5.RowCount = 0 Then
            Exit Sub
        ElseIf DataGridView1.SelectedCells.Count = 0 Then
            Exit Sub
        End If

        'データのチェック

        'datagridviewのヘッダーテキストをコレクションに取り込む
        Dim DGVheaderCheck1 As ArrayList = New ArrayList
        For c As Integer = 0 To DataGridView1.Columns.Count - 1
            DGVheaderCheck1.Add(DataGridView1.Item(c, 0).Value)
        Next c
        Dim DGVheaderCheck2 As ArrayList = New ArrayList
        For c As Integer = 0 To DataGridView5.Columns.Count - 1
            DGVheaderCheck2.Add(DataGridView5.Item(c, 0).Value)
        Next c

        DataGridView1.Columns.Add("column" & DataGridView1.ColumnCount, DataGridView1.ColumnCount)
        Dim addColumn As Integer = DataGridView1.ColumnCount - 1
        DataGridView1.Item(addColumn, 0).Value = "商品群"

        For r1 As Integer = 1 To DataGridView1.RowCount - 1
            Dim denpyoNum As String = ""
            If DGVheaderCheck1.Contains("伝票no") Then
                denpyoNum = DataGridView1.Item(DGVheaderCheck1.IndexOf("伝票no"), r1).Value
            ElseIf DGVheaderCheck1.Contains("伝票番号") Then
                denpyoNum = DataGridView1.Item(DGVheaderCheck1.IndexOf("伝票番号"), r1).Value
            ElseIf DGVheaderCheck1.Contains("注文日") Then
                denpyoNum = DataGridView1.Item(DGVheaderCheck1.IndexOf("注文日"), r1).Value
            End If

            For r2 As Integer = 0 To DataGridView5.RowCount - 1
                Dim sDenpyoNum As String = DataGridView5.Item(DGVheaderCheck2.IndexOf("伝票番号"), r2).Value
                If denpyoNum = sDenpyoNum Then
                    Dim sCode As String = DataGridView5.Item(DGVheaderCheck2.IndexOf("商品ｺｰﾄﾞ"), r2).Value
                    Dim sCount As String = DataGridView5.Item(DGVheaderCheck2.IndexOf("受注数"), r2).Value
                    If DataGridView1.Item(addColumn, r1).Value <> "" Then
                        DataGridView1.Item(addColumn, r1).Value &= "、"
                    End If
                    DataGridView1.Item(addColumn, r1).Value &= sCode & "*" & sCount
                End If
            Next
        Next
    End Sub

    Private Sub CheckBox8_CheckedChanged(ByVal sender As System.Object, ByVal e As EventArgs) Handles CheckBox8.CheckedChanged
        If CheckBox8.Checked = True Then
            ComboBox9.Enabled = True
        Else
            ComboBox9.Enabled = False
        End If
    End Sub

    '抽出用コード作成
    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button17.Click
        If TextBox11.Text <> "" Then
            Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\location.csv"
            Dim csvRecords As New ArrayList()
            csvRecords = CSV_READ(fName)

            TextBox11.Text = Form1.StrConvToNarrow(TextBox11.Text)

            Dim answerArray As ArrayList = New ArrayList
            Dim lines As String() = Split(TextBox11.Text, vbCrLf)
            For i As Integer = 0 To lines.Length - 1
                For r As Integer = 0 To csvRecords.Count - 1
                    Dim sArray As String() = Split(csvRecords(r), "|=|")
                    If Regex.IsMatch(sArray(0), "^" & lines(i) & "$") Then
                        If sArray(3) = "" Then
                            answerArray.Add(sArray(0))
                        Else
                            answerArray.Add(sArray(3))
                        End If
                        Exit For
                    End If
                Next
            Next

            DataGridView1.Rows.Clear()
            DataGridView1.Columns.Clear()

            DataGridView1.Columns.Add("1", "コード")
            For i As Integer = 0 To answerArray.Count - 1
                DataGridView1.Rows.Add(answerArray(i))
            Next

            '重複削除する
            If CheckBox13.Checked Then
                For r As Integer = 0 To DataGridView1.RowCount - 1
                    DataGridView1.Item(0, r).Selected = True
                Next
                JuuhukuDel()
            End If

            'ソート
            If CheckBox14.Checked Then
                DataGridView1.Sort(DataGridView1.Columns(0), ComponentModel.ListSortDirection.Ascending)
            End If

            DataGridView1.ClearSelection()
        End If
    End Sub

    Private Sub TextBox11_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles TextBox11.KeyDown
        If e.Modifiers = Keys.Control And e.KeyCode = Keys.A Then
            TextBox11.SelectAll()
        End If
    End Sub

    '楽天予約文追加
    Private Sub Button33_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button33.Click
        Dim DGVheaderCheck1 As ArrayList = New ArrayList
        For c As Integer = 0 To DataGridView1.Columns.Count - 1
            DGVheaderCheck1.Add(DataGridView1.Item(c, 0).Value)
        Next c

        Dim cArray As String() = Split(TextBox12.Text, vbCrLf)
        For i As Integer = 0 To cArray.Length - 1
            If cArray(i) <> "" Then
                DataGridView1.Rows.Add()
                Dim r As Integer = DataGridView1.RowCount - 1
                DataGridView1.Item(DGVheaderCheck1.IndexOf("項目選択肢用コントロールカラム"), r).Value = "n"
                DataGridView1.Item(DGVheaderCheck1.IndexOf("商品管理番号（商品URL）"), r).Value = cArray(i)
                DataGridView1.Item(DGVheaderCheck1.IndexOf("選択肢タイプ"), r).Value = "s"
                DataGridView1.Item(DGVheaderCheck1.IndexOf("Select/Checkbox用項目名"), r).Value = TextBox13.Text & TextBox14.Text & TextBox15.Text
                DataGridView1.Item(DGVheaderCheck1.IndexOf("Select/Checkbox用選択肢"), r).Value = "了承済み！"
            End If
        Next
    End Sub

    Private Sub Button34_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button34.Click
        Try
            Dim DGVheaderCheck1 As ArrayList = New ArrayList
            For c As Integer = 0 To DataGridView1.Columns.Count - 1
                DGVheaderCheck1.Add(DataGridView1.Item(c, 0).Value)
            Next c

            Dim str As String = ""
            For r As Integer = 1 To DataGridView1.RowCount - 1
                str &= DataGridView1.Item(DGVheaderCheck1.IndexOf("商品番号"), r).Value & vbCrLf
            Next
            TextBox12.Text = str
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button35_Click(sender As Object, e As EventArgs) Handles Button35.Click
        Dim copySrc As String = Form1.ロケーションToolStripMenuItem.Text & "\location.csv"
        Dim copyDst As String = Path.GetDirectoryName(Form1.appPath) & "\location.csv"
        If File.Exists(copySrc) Then
            My.Computer.FileSystem.CopyFile(copySrc, copyDst, True)
            MsgBox("コピーしました。ファイルの更新日は" & vbCrLf & File.GetLastWriteTime(copyDst))
        Else
            MsgBox("サーバーのファイルが見つかりません")
        End If
    End Sub

    Private Sub 件ToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles 件ToolStripMenuItem3.Click
        SaveA(100)
    End Sub

    Private Sub 件ToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles 件ToolStripMenuItem4.Click
        SaveA(150)
    End Sub

    Private Sub 件ToolStripMenuItem5_Click(sender As Object, e As EventArgs) Handles 件ToolStripMenuItem5.Click
        SaveA(250)
    End Sub

    Private Sub 件ToolStripMenuItem6_Click(sender As Object, e As EventArgs) Handles 件ToolStripMenuItem6.Click
        SaveA(500)
    End Sub

    Private Sub ToolStripTextBox1_KeyUp(sender As Object, e As KeyEventArgs) Handles ToolStripTextBox1.KeyUp
        If e.KeyCode = Keys.Enter Then
            ToolStripDropDownButton2.DropDown.Close()
            Application.DoEvents()
            If IsNumeric(ToolStripTextBox1.Text) Then
                If ToolStripTextBox1.Text > 0 Then
                    SaveA(ToolStripTextBox1.Text)
                End If
            End If
        End If
    End Sub
End Class