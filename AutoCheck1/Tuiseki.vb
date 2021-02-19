Imports EncodingOperation
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic.FileIO
Imports System.IO
Imports System.Net
Imports System.ComponentModel
Imports System.Text
Imports System.Threading

Public Class Tuiseki
    Dim enc As Encoding = Encoding.GetEncoding("utf-8")
    Dim encUTF8 As Encoding = Encoding.GetEncoding("utf-8")

    Dim splitter As String = ","    '区切り文字

    Dim sagawa As String = "https://k2k.sagawa-exp.co.jp/p/web/okurijosearch.do?okurijoNo="
    Dim sagawa2 As String = "http://k2k.sagawa-exp.co.jp/cgi-bin/mall.mmcgi?"
    'oku01=358611435781&oku02=358611435663&oku03=358611441473
    Dim yamato As String = "http://jizen.kuronekoyamato.co.jp/jizen/servlet/crjz.b.NQ0010?id="
    Dim yubin1 As String = "https://trackings.post.japanpost.jp/services/srv/search/?requestNo1="
    Dim yubin2 As String = "&search=追跡スタート"

    Dim haitatu As Integer = 0
    Dim compCount As Integer = 0
    Dim kekka As String = ""
    Dim webReadComplete As Boolean = False
    Dim numberArray As New ArrayList
    Dim kensakuTitle1 As String = ""
    Dim kensakuTitle2 As String = ""

    Dim delMoji As String = vbTab & "|" & vbLf & "|" & vbCr & "|\?"
    Dim CMS1_kihon As New ArrayList

    Dim dH1 As New ArrayList
    Dim dH3 As New ArrayList
    Dim dH_hibetu As New ArrayList

    Dim serverPath As String = "\\SERVER2-PC\files\発送状況調査結果"

    Dim TB_Counter As TextBox() = Nothing

    Private Sub Tuiseki_Load(sender As Object, e As EventArgs) Handles Me.Load
        'SplitContainer2.Panel2Collapsed = True

        DGV1.Columns(3).Frozen = True
        DGV1.Columns(3).DividerWidth = 3
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0

        For i As Integer = 0 To ContextMenuStrip1.Items.Count - 1
            CMS1_kihon.Add(ContextMenuStrip1.Items(i))
        Next
        Dim now_str As String = Format(Now, "yyyy/MM/dd")
        DateTimePicker1.Value = CDate(now_str & " 00:00:00")
        DateTimePicker2.Value = CDate(now_str & " 23:59:59")

        TB_Counter = {TextBox1, TextBox2, TextBox6, TextBox14, TextBox4, TextBox5, TextBox10, TextBox15, TextBox13, TextBox16, TextBox17}
        For Each cK As String In CounterKeys
            Counter.Add(cK, 0)
        Next

        Dim dgv As DataGridView() = {DGV1, DGV2, DGV3, DGV_hibetu}
        For i As Integer = 0 To dgv.Length - 1
            For Each c As DataGridViewColumn In dgv(i).Columns
                c.SortMode = DataGridViewColumnSortMode.NotSortable
            Next c
        Next
        dgv = {DGV2, DGV_hibetu}
        For i As Integer = 0 To dgv.Length - 1
            dgv(i).BringToFront()
            dgv(i).Dock = DockStyle.Fill
        Next

        dH1 = TM_HEADER_GET(DGV1)
        dH3 = TM_HEADER_GET(DGV3)
        dH_hibetu = TM_HEADER_GET(DGV_hibetu)

        Dim haisou As String() = New String() {"佐川", "ゆパケ", "定形外", "ヤマト"}
        For r As Integer = 0 To 3
            DGV3.Rows.Add(haisou(r))
            For c As Integer = 1 To DGV3.ColumnCount - 1
                DGV3.Item(c, DGV3.RowCount - 1).Value = 0
                'DGV3.Item(c, DGV3.RowCount - 1).Style.Alignment = DataGridViewContentAlignment.MiddleRight
            Next
        Next

        Me.Invoke(New CallDelegate(AddressOf CB3get))
        'CB3get()
    End Sub

    Dim cb3_mae As Integer = 0
    Private dPath As String = ""
    Private Sub CB3get()
        'Me.Invoke(New CallDelegate(AddressOf CB3_selectmae))
        If ComboBox3.SelectedIndex >= 0 Then
            cb3_mae = ComboBox3.SelectedIndex
        End If

        'Try
        'Me.Invoke(New CallDelegate(AddressOf CB3_clear))
        ComboBox3.Items.Clear()

        If InStr(Environment.MachineName, "TAK") > 0 Then
            dPath = "N:\●追跡"
            CheckBox7.Enabled = True
            'ElseIf InStr(Environment.MachineName, "NAKA") > 0 And InStr(appPath, "Debug") > 0 Then
            '    dPath = "I:\●追跡"
            '    If Not Directory.Exists(dPath) Then
            '        dPath = serverPath
            '    End If
            '    CheckBox7.Enabled = True
        Else
            dPath = serverPath
        End If

        DGV_hibetu.Rows.Clear()

        Dim kubetu As String() = New String() {"未", "該当無", "対応確認", "返送", "C", "不明"}
        Dim files As String() = Directory.GetFiles(dPath, "*.csv", IO.SearchOption.TopDirectoryOnly)
        '1 次元 Array 内または Array の一部内の要素の順序を反転させます
        Array.Reverse(files)
        For i As Integer = 0 To files.Length - 1
            Dim fName As String = Path.GetFileNameWithoutExtension(files(i))
            Dim hiduke1 As String = Split(fName, "(")(0)
            Dim hiduke2 As String = hiduke1.Substring(0, 2) & "/" & hiduke1.Substring(2, 2) & "/" & hiduke1.Substring(4, 2)
            '指定された 2 つの日時の差分を指定の単位で返す関数です
            'If DateDiff(DateInterval.Day, CDate(hiduke2), Now) > 60 Then
            '    'ループを終了し次のループに移る
            '    Continue For
            'End If

            ComboBox3.Items.Add(fName)

            DGV_hibetu.Rows.Add()
            Dim r As Integer = DGV_hibetu.RowCount - 1
            DGV_hibetu.Item(dH_hibetu.IndexOf("日付"), r).Value = Split(fName, "(全")(0)
            Dim all As String = Split(Split(fName, "(全")(1), ")")(0)
            DGV_hibetu.Item(dH_hibetu.IndexOf("すべて"), r).Value = all
            Dim fN As String() = Split(fName, " ")
            If fN.Length > 1 Then
                For Each fNstr As String In fN
                    For k As Integer = 0 To kubetu.Length - 1
                        If InStr(fNstr, kubetu(k).Substring(0, 1)) > 0 Then
                            fNstr = Replace(fNstr, kubetu(k).Substring(0, 1), "")
                            DGV_hibetu.Item(dH_hibetu.IndexOf(kubetu(k)), r).Value = fNstr
                            DGV_hibetu.Item(dH_hibetu.IndexOf(kubetu(k)), r).Style.BackColor = Color.LemonChiffon
                            Exit For
                        End If
                    Next
                Next
            End If

            If i > 60 Then
                Exit For
            End If
        Next

        DGV_hibetu.Rows.Add("合計")
        Dim newRow As Integer = DGV_hibetu.RowCount - 1
        For c As Integer = 1 To DGV_hibetu.ColumnCount - 1
            Dim rowSum As Integer = 0
            For r As Integer = 0 To DGV_hibetu.RowCount - 2
                If IsNumeric(DGV_hibetu.Item(c, r).Value) Then
                    rowSum += CInt(DGV_hibetu.Item(c, r).Value)
                End If
            Next
            DGV_hibetu.Item(c, newRow).Value = rowSum
        Next

        'Me.Invoke(New CallDelegate(AddressOf CB3_selectindex))
        ComboBox3.SelectedIndex = cb3_mae
        'Catch ex As Exception

        'End Try
    End Sub
    'Private Sub CB3_clear()
    '    ComboBox3.Items.Clear()
    'End Sub
    'Private Sub CB3_selectmae()
    '    If ComboBox3.SelectedIndex >= 0 Then
    '        cb3_mae = ComboBox3.SelectedIndex
    '    End If
    'End Sub
    'Private Sub CB3_selectindex()
    '    ComboBox3.SelectedIndex = cb3_mae
    'End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        Dim cb3 As String = Split(ComboBox3.Text, "(全")(0)
        Dim cb3_date As String = "20" & cb3.Substring(0, 2) & "/" & cb3.Substring(2, 2) & "/" & cb3.Substring(4, 2)
        DateTimePicker1.Value = CDate(cb3_date & " 00:00:00")
        DateTimePicker2.Value = CDate(cb3_date & " 23:59:59")
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.Invoke(New CallDelegate(AddressOf TBText_Counter_Reset))
        DGV1.Rows.Clear()
        Me.Invoke(New CallDelegate(AddressOf CB3get))
        'CB3get()

        Dim files As String() = Directory.GetFiles(dPath, "*.csv", System.IO.SearchOption.AllDirectories)
        Array.Reverse(files)

        Dim cb3_flag As Boolean = True
        Dim startD As Date = DateTimePicker1.Value
        While cb3_flag
            Dim cName As String = Format(startD, "yyMMdd")

            For i As Integer = 0 To files.Length - 1
                Dim fName As String = Path.GetFileNameWithoutExtension(files(i))
                'Dim cb As String = Split(ComboBox3.Text, "(全")(0)
                If Regex.IsMatch(fName, "^" & cName) Then
                    Me.Text = files(i)
                    CSVtoDGV(files(i))

                    ToolStripProgressBar1.Value = 0
                    ToolStripProgressBar1.Maximum = DGV1.RowCount

                    ComboBox1.Items.Clear()
                    ComboBox1.Items.Add("店舗")
                    For r As Integer = 0 To DGV1.RowCount - 2
                        Dim tenpo As String = DGV1.Item(dH1.IndexOf("店舗"), r).Value
                        If tenpo Is Nothing Then
                            Continue For
                        End If

                        If Not ComboBox1.Items.Contains(tenpo) Then
                            ComboBox1.Items.Add(tenpo)
                        End If

                        ToolStripProgressBar1.Value += 1
                    Next
                    ComboBox1.SelectedIndex = 0

                    ToolStripProgressBar1.Value = 0
                    ToolStripProgressBar1.Maximum = 100

                    CountDisplay()
                    CountAll()

                    Exit For
                End If

                Application.DoEvents()
            Next

            Dim endD As Date = DateAdd(DateInterval.Day, 1, startD)
            If endD > DateTimePicker2.Value Then
                cb3_flag = False
            Else
                startD = endD
            End If
        End While

        If Format(DateTimePicker1.Value, "yyMMdd") <> Format(DateTimePicker2.Value, "yyMMdd") Then
            Me.Text = Format(DateTimePicker1.Value, "yyMMdd") & "～" & Format(DateTimePicker2.Value, "yyMMdd")
        End If

        If SplitContainer2.Panel2Collapsed = False Then
            NissuIchiran()
        End If
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        If ComboBox3.SelectedIndex < ComboBox3.Items.Count - 1 Then
            ComboBox3.SelectedIndex += 1
        End If
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        If ComboBox3.SelectedIndex > 0 Then
            ComboBox3.SelectedIndex -= 1
        End If
    End Sub

    '行番号を表示する
    Private Sub DataGridView1_RowPostPaint(ByVal sender As DataGridView, ByVal e As DataGridViewRowPostPaintEventArgs) Handles DGV1.RowPostPaint, DGV2.RowPostPaint, DGV_hibetu.RowPostPaint
        ' 行ヘッダのセル領域を、行番号を描画する長方形とする（ただし右端に4ドットのすき間を空ける）
        Dim rect As New Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, sender.RowHeadersWidth - 4, sender.Rows(e.RowIndex).Height)

        ' 上記の長方形内に行番号を縦方向中央＆右詰で描画する　フォントや色は行ヘッダの既定値を使用する
        TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), sender.RowHeadersDefaultCellStyle.Font, rect, sender.RowHeadersDefaultCellStyle.ForeColor, TextFormatFlags.VerticalCenter Or TextFormatFlags.Right)
    End Sub

    'コンボボックスカラー
    Private Sub ComboBox3_DrawItem(sender As Object, e As DrawItemEventArgs) Handles ComboBox3.DrawItem
        If e.Index = -1 Then Exit Sub
        Dim cb As ComboBox = sender
        Dim str As String = cb.Items(e.Index).ToString
        Dim fnt As Font = New Font(e.Font, FontStyle.Regular)
        Dim brs As Brush = Brushes.Black
        If InStr(str, "完") > 0 Then
            If InStr(str, "該") > 0 Then
                brs = Brushes.DarkGray
                fnt = New Font(e.Font, FontStyle.Bold)
            Else
                brs = Brushes.Gray
            End If
        ElseIf InStr(str, "該") > 0 Then
            brs = Brushes.Red
            fnt = New Font(e.Font, FontStyle.Bold)
        End If
        e.DrawBackground()
        e.Graphics.DrawString(str, fnt, brs, e.Bounds.X, e.Bounds.Y)
        e.DrawFocusRectangle()
    End Sub

    Private Sub DGV1_DragDrop(ByVal sender As Object, ByVal e As Windows.Forms.DragEventArgs) Handles DGV1.DragDrop
        If DGV1.RowCount > 1 Then
            Dim DR As DialogResult = MsgBox("追加しますか", MsgBoxStyle.YesNo Or MsgBoxStyle.SystemModal)
            If DR = Windows.Forms.DialogResult.No Then
                DGV1.Rows.Clear()
            End If
        End If

        Me.TopMost = True
        Me.TopMost = False

        If DGV1.RowCount > 1 Then
            Dim DR As DialogResult = MsgBox("該当なしの情報の追加ですか？", MsgBoxStyle.YesNo Or MsgBoxStyle.SystemModal)
            If DR = Windows.Forms.DialogResult.Yes Then
                For Each filename As String In e.Data.GetData(DataFormats.FileDrop)
                    DropCsvCheck(filename)
                Next

                MsgBox("終了しました")
                CountDisplay()
                CountAll()
                Exit Sub
            End If
        End If

        For Each filename As String In e.Data.GetData(DataFormats.FileDrop)
            Me.Text = filename
            CSVtoDGV(filename)
        Next

        CB1get()

        ToolStripProgressBar1.Value = 0
        ToolStripProgressBar1.Maximum = 100

        CountDisplay()
        CountAll()

        Me.TopMost = True
        Me.TopMost = False

        MsgBox("終了しました")
    End Sub

    Private Sub DropCsvCheck(filename As String)
        Dim csvRecords As New ArrayList()
        csvRecords = TM_CSV_READ(filename)(0)

        If Regex.IsMatch(csvRecords(0), "伝票番号.*ﾋﾟｯｷﾝｸﾞ指示内容.*発送伝票番号") Then
            '伝票検索画面csv
            Dim addHeader As String() = New String() {"配送会社|発送方法", "追跡番号|発送伝票番号", "ﾋﾟｯｷﾝｸﾞ指示内容|ﾋﾟｯｷﾝｸﾞ指示内容", "状態|状態"}
            Dim csvH As String() = Split(csvRecords(0), "|=|")
            For i As Integer = 1 To csvRecords.Count - 1
                Dim sArray As String() = Split(csvRecords(i), "|=|")
                Dim denpyoNo1 As String = sArray(Array.IndexOf(csvH, "伝票番号"))
                For r As Integer = 0 To DGV1.RowCount - 1
                    If DGV1.Rows(r).IsNewRow Then
                        Continue For
                    End If
                    If DGV1.Item(dH1.IndexOf("伝票番号"), r).Value = denpyoNo1 Then
                        For Each aH As String In addHeader
                            Dim aHArray As String() = Split(aH, "|")
                            If aHArray(0) = "追跡番号" Then
                                Dim csvTuiseki As String = sArray(Array.IndexOf(csvH, aHArray(1)))
                                If InStr(csvTuiseki, DGV1.Item(dH1.IndexOf(aHArray(0)), r).Value) = 0 Then
                                    If InStr(csvTuiseki, ",") > 0 Then
                                        DGV1.Item(dH1.IndexOf(aHArray(0)), r).Value = sArray(Array.IndexOf(csvH, aHArray(1)))
                                        DGV1.Item(dH1.IndexOf(aHArray(0)), r).Style.BackColor = Color.Red
                                    Else
                                        DGV1.Item(dH1.IndexOf(aHArray(0)), r).Value = sArray(Array.IndexOf(csvH, aHArray(1)))
                                        DGV1.Item(dH1.IndexOf(aHArray(0)), r).Style.BackColor = Color.Orange
                                    End If
                                End If
                            Else
                                If DGV1.Item(dH1.IndexOf(aHArray(0)), r).Value <> sArray(Array.IndexOf(csvH, aHArray(1))) Then
                                    DGV1.Item(dH1.IndexOf(aHArray(0)), r).Value = sArray(Array.IndexOf(csvH, aHArray(1)))
                                    If aHArray(0) = "状態" Then
                                        If InStr(DGV1.Item(dH1.IndexOf(aHArray(0)), r).Value, "キャンセル") > 0 Then
                                            DGV1.Item(dH1.IndexOf(aHArray(0)), r).Style.BackColor = Color.Red
                                        Else
                                            DGV1.Item(dH1.IndexOf(aHArray(0)), r).Style.BackColor = Color.Orange
                                        End If
                                    Else
                                        DGV1.Item(dH1.IndexOf(aHArray(0)), r).Style.BackColor = Color.Orange
                                    End If
                                End If
                            End If
                        Next
                    End If
                Next
            Next
        ElseIf Regex.IsMatch(csvRecords(0), "受注日.*伝票番号.*商品ｺｰﾄﾞ") Then
            '受注明細csv
            Dim addHeader As String() = New String() {"商品|商品ｺｰﾄﾞ"}
            Dim csvH As String() = Split(csvRecords(0), "|=|")
            Dim dNoArray As New ArrayList   '複数商品を入力するため、伝票番号を記録し、今回の入力分だけ記入するようにする
            For i As Integer = 1 To csvRecords.Count - 1
                Dim sArray As String() = Split(csvRecords(i), "|=|")
                Dim denpyoNo1 As String = sArray(Array.IndexOf(csvH, "伝票番号"))
                For r As Integer = 0 To DGV1.RowCount - 1
                    If DGV1.Rows(r).IsNewRow Then
                        Continue For
                    End If
                    If DGV1.Item(dH1.IndexOf("伝票番号"), r).Value = denpyoNo1 Then
                        For Each aH As String In addHeader
                            Dim aHArray As String() = Split(aH, "|")
                            If aHArray(0) = "商品" Then
                                If Not dNoArray.Contains(denpyoNo1) Then    '今回初めてならまず削除
                                    DGV1.Item(dH1.IndexOf(aHArray(0)), r).Value = ""
                                    Dim sCode As String = sArray(Array.IndexOf(csvH, aHArray(1))) & "*" & sArray(Array.IndexOf(csvH, "受注数"))
                                    DGV1.Item(dH1.IndexOf(aHArray(0)), r).Value = sCode
                                Else
                                    Dim sCode As String = sArray(Array.IndexOf(csvH, aHArray(1))) & "*" & sArray(Array.IndexOf(csvH, "受注数"))
                                    DGV1.Item(dH1.IndexOf(aHArray(0)), r).Value &= "、" & sCode
                                End If
                                DGV1.Item(dH1.IndexOf(aHArray(0)), r).Style.BackColor = Color.Orange
                                dNoArray.Add(denpyoNo1)
                            End If
                        Next
                    End If
                Next
            Next
        Else
            '不明
        End If
    End Sub

    Private Sub CB1get()
        Me.Invoke(New CallDelegate(AddressOf CB1_clear))

        callCB1 = "店舗"
        Me.Invoke(New CallDelegate(AddressOf CB1_add))
        For r As Integer = 0 To DGV1.RowCount - 2
            Dim tenpo As String = DGV1.Item(dH1.IndexOf("店舗"), r).Value
            If Not ComboBox1.Items.Contains(tenpo) Then
                callCB1 = tenpo
                Me.Invoke(New CallDelegate(AddressOf CB1_add))
            End If
        Next
        Me.Invoke(New CallDelegate(AddressOf CB1_select))
    End Sub

    Private Sub CB3save()
        Dim flag As Boolean = False
        If InStr(Environment.MachineName, "TAK") > 0 Or InStr(Environment.MachineName, "PING2") Or InStr(Environment.MachineName, "PING") Or InStr(Environment.MachineName, "MAO") Or InStr(Environment.MachineName, "SERVER5") Then
            flag = True
        ElseIf InStr(Environment.MachineName, "NAKA") > 0 And InStr(appPath, "Debug") > 0 Then
            If Directory.Exists(dPath) Then
                flag = True
            End If
        End If

        If flag Then
            Dim str As String = ""
            For Each cbStr As String In ComboBox3.Items
                str &= cbStr & vbCrLf
            Next

            Dim saveName As String = dPath & "\log\" & Format(Now, "yyyyMMdd_HHmmss") & ".txt"
            If Not Directory.Exists(dPath & "\log") Then
                Directory.CreateDirectory(dPath & "\log")
            End If
            File.WriteAllText(saveName, str)
        End If
    End Sub

    Dim pattern1 As String() = New String() {"配送会社|発送方法", "追跡番号|発送伝票番号", "伝票番号", "注文番号|受注番号", "届け先", "結果", "済", "受注日", "納品書発行日", "出荷確定日", "出荷日", "配達完了日", "店舗", "送り先名", "送り先住所", "受注分類タグ", "ﾋﾟｯｷﾝｸﾞ指示内容", "状態", "調査日", "個数", "詳細", "初回反応", "所要日数", "商品"}
    Dim pattern2 As String() = New String() {"配送会社", "追跡番号", "伝票番号", "注文番号", "届け先", "結果", "済", "受注日", "納品書発行日", "出荷確定日", "出荷日", "配達完了日", "店舗", "送り先名", "送り先住所", "受注分類タグ", "ﾋﾟｯｷﾝｸﾞ指示内容", "状態", "調査日", "個数", "詳細", "初回反応", "所要日数", "商品"}
    Private Sub CSVtoDGV(filename As String)
        Dim csvRecords As New ArrayList()
        csvRecords = TM_CSV_READ(filename)(0)

        Dim csvH As String() = Split(csvRecords(0), "|=|")

        ToolStripProgressBar1.Value = 0
        ToolStripProgressBar1.Maximum = csvRecords.Count
        For i As Integer = 1 To csvRecords.Count - 1
            Dim sArray As String() = Split(csvRecords(i), "|=|")
            DGV1.Rows.Add()
            Dim r As Integer = DGV1.RowCount - 2

            For k As Integer = 0 To pattern1.Length - 1
                Dim m As Match = Regex.Match(csvRecords(0), pattern1(k))
                If csvH.Contains(m.Value) Then
                    Dim mStr As String = sArray(Array.IndexOf(csvH, m.Value))
                    If mStr <> "" Then
                        mStr = Regex.Replace(mStr, delMoji, "")
                    End If
                    '0000-0000-0000 => 000000000000
                    If pattern2(k) = "追跡番号" Then
                        If mStr <> "" Then
                            mStr = Regex.Replace(mStr, "\-", "")
                        End If
                    End If
                    DGV1.Item(dH1.IndexOf(pattern2(k)), r).Value = mStr

                    If pattern2(k) = "済" Then
                        If DGV1.Item(dH1.IndexOf(pattern2(k)), r).Value = "済" Then
                            DGV1.Item(dH1.IndexOf(pattern2(k)), r).Style.BackColor = Color.Yellow
                        ElseIf DGV1.Item(dH1.IndexOf(pattern2(k)), r).Value = "返送" Then
                            DGV1.Item(dH1.IndexOf(pattern2(k)), r).Style.BackColor = Color.Orange
                        ElseIf DGV1.Item(dH1.IndexOf(pattern2(k)), r).Value = "対応確認" Then
                            DGV1.Item(dH1.IndexOf(pattern2(k)), r).Style.BackColor = Color.LightBlue
                        ElseIf DGV1.Item(dH1.IndexOf(pattern2(k)), r).Value = "不明" Then
                            DGV1.Item(dH1.IndexOf(pattern2(k)), r).Style.BackColor = Color.Orange
                        End If
                    End If

                    If pattern2(k) = "状態" Then
                        If InStr(DGV1.Item(dH1.IndexOf(pattern2(k)), r).Value, "キャンセル") > 0 Then
                            DGV1.Item(dH1.IndexOf(pattern2(k)), r).Style.BackColor = Color.Red
                        End If
                    End If

                    If pattern2(k) = "商品" Then
                        If DGV1.Item(dH1.IndexOf(pattern2(k)), r).Value <> "" Then
                            DGV1.Item(dH1.IndexOf(pattern2(k)), r).Style.BackColor = Color.PaleGreen
                        End If
                    End If
                End If

                If pattern2(k) = "所要日数" Then
                    'If DGV1.Item(dH1.IndexOf("所要日数"), r).Value = "" Then
                    If DGV1.Item(dH1.IndexOf("出荷日"), r).Value <> "" And DGV1.Item(dH1.IndexOf("配達完了日"), r).Value <> "" Then
                        If IsDate(DGV1.Item(dH1.IndexOf("出荷日"), r).Value) And IsDate(DGV1.Item(dH1.IndexOf("配達完了日"), r).Value) Then
                            Dim syukkabi As Date = Format(CDate(DGV1.Item(dH1.IndexOf("出荷日"), r).Value), "yyyy/MM/dd 12:00:00")
                            Dim kanryobi As Date = Format(CDate(DGV1.Item(dH1.IndexOf("配達完了日"), r).Value), "yyyy/MM/dd 12:00:00")
                            Dim sNissu As Integer = DateDiff(DateInterval.Day, syukkabi, kanryobi)
                            If sNissu <> 0 Then
                                DGV1.Item(dH1.IndexOf(pattern2(k)), r).Value = sNissu
                            End If
                        End If
                    End If
                    'End If
                End If
            Next

            '追跡番号が複数ある時は行複製する
            If InStr(DGV1.Item(dH1.IndexOf("追跡番号"), r).Value, ",") > 0 Then
                Dim tuisekiArray As String() = Split(DGV1.Item(dH1.IndexOf("追跡番号"), r).Value, ",")
                For t As Integer = 0 To tuisekiArray.Length - 1
                    If t = 0 Then
                        DGV1.Item(dH1.IndexOf("追跡番号"), r).Value = tuisekiArray(0)
                        DGV1.Item(dH1.IndexOf("伝票番号"), r).Style.BackColor = Color.LightBlue
                        Application.DoEvents()
                    Else
                        Dim oldR As Integer = r
                        DGV1.Rows.Add()
                        Dim newR As Integer = DGV1.RowCount - 2
                        For c As Integer = 0 To DGV1.ColumnCount - 1
                            Application.DoEvents()
                            If c = dH1.IndexOf("追跡番号") Then
                                DGV1.Item(dH1.IndexOf("追跡番号"), newR).Value = tuisekiArray(t)
                            Else
                                If DGV1.Item(c, oldR).Value <> "" Then
                                    DGV1.Item(c, newR).Value = DGV1.Item(c, oldR).Value
                                End If
                            End If
                        Next
                        DGV1.Item(dH1.IndexOf("伝票番号"), newR).Style.BackColor = Color.LightBlue
                    End If
                Next
            End If

            ToolStripProgressBar1.Value += 1
        Next
    End Sub

    Private Sub DGV1_DragEnter(ByVal sender As Object, ByVal e As Windows.Forms.DragEventArgs) Handles DGV1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        NameChangeSave(1)
    End Sub

    Private Sub NameChangeSave(mode As Integer, Optional dialogON As Boolean = True)
        If DGV1.RowCount <= 1 Then
            Exit Sub
        End If

        Dim savePath As String = ""
        If dialogON Then
            '.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
            '.RestoreDirectory = True,
            Dim sfd As New SaveFileDialog With {
                .InitialDirectory = dPath,
                .Filter = "CSVファイル(*.csv)|*.csv|すべてのファイル(*.*)|*.*",
                .AutoUpgradeEnabled = False,
                .FilterIndex = 0,
                .Title = "保存先のファイルを選択してください",
                .OverwritePrompt = True,
                .CheckPathExists = True
            }

            'If InStr(Me.Text, "(全") = 0 Then
            '    Dim dH0 As ArrayList = TM_HEADER_GET(DGV1)
            '    Dim filenameA As String = DGV1.Item(dH0.IndexOf("出荷確定日"), 0).Value
            '    Dim filenameB As String = Format(CDate(filenameA), "yyMMdd")
            '    Dim week As String() = {"日", "月", "火", "水", "木", "金", "土"}
            '    Dim filenameC As String = week(CDate(filenameA).DayOfWeek)
            '    sfd.FileName = filenameB & "(" & filenameC & ")(全" & ".csv"
            'ElseIf InStr(Me.Text, "\") > 0 Then
            '    Dim sPath As String = Path.GetFileNameWithoutExtension(Me.Text)
            '    sfd.FileName = sPath & ".csv"
            'Else
            '    sfd.FileName = "新しいファイル.csv"
            'End If

            Dim dH0 As ArrayList = TM_HEADER_GET(DGV1)
            Dim filenameA As String = DGV1.Item(dH0.IndexOf("出荷確定日"), 0).Value
            Dim filenameB As String = Format(CDate(filenameA), "yyMMdd")
            Dim week As String() = {"日", "月", "火", "水", "木", "金", "土"}
            Dim filenameC As String = week(CDate(filenameA).DayOfWeek)
            sfd.FileName = filenameB & "(" & filenameC & ")(全" & ".csv"

            'ダイアログを表示する
            If sfd.ShowDialog(Me) = DialogResult.OK Then
                savePath = sfd.FileName
            Else
                Exit Sub
            End If
        Else
            savePath = Me.Text
        End If

        SaveCsv(DGV1, savePath, mode)

        Dim fName As String = Path.GetFileNameWithoutExtension(Me.Text)
        Dim sName As String = ""
        Dim newName As String = ""
        Dim naiyou As String = ""
        naiyou = "(全" & TextBox1.Text & ")"
        If TextBox16.Text > 0 Then
            naiyou &= " 未" & TextBox16.Text
        End If
        If TextBox5.Text > 0 Then
            naiyou &= " 該" & TextBox5.Text
        End If
        If TextBox14.Text > 0 Then
            naiyou &= " 対" & TextBox14.Text
        End If
        If TextBox13.Text > 0 Then
            naiyou &= " 返" & TextBox13.Text
        End If
        If TextBox17.Text > 0 Then
            naiyou &= " 不" & TextBox17.Text
        End If
        If TextBox15.Text > 0 Then
            naiyou &= " C" & TextBox15.Text
        End If
        If InStr(fName, "(全") > 0 Then
            Dim fn As String() = Split(fName, "(全")
            newName = fn(0) & naiyou
            sName = fn(0)
        Else
            newName = fName & naiyou
            sName = fName
        End If
        Dim newPath As String = Replace(savePath, fName, newName)
        File.Move(savePath, newPath)

        'サーバーに保存
        Dim serverNewPath As String = ""
        If Regex.IsMatch(Environment.MachineName, "TAK|NAKA|PING|PING2|MAO|SERVER5") Then
            Try
                Dim saveFlag As Boolean = False
                Dim files As String() = Directory.GetFiles(serverPath, "*.csv", System.IO.SearchOption.AllDirectories)
                Array.Reverse(files)
                For i As Integer = 0 To files.Length - 1
                    fName = Path.GetFileNameWithoutExtension(files(i))
                    If InStr(fName, sName) > 0 Then
                        If InStr(fName, "(全") > 0 Then
                            Dim fn As String() = Split(fName, "(全")
                            newName = fn(0) & naiyou
                        Else
                            newName = fName & naiyou
                        End If
                        serverNewPath = Replace(files(i), fName, newName)
                        File.Move(files(i), serverNewPath)
                        SaveCsv(DGV1, serverNewPath, mode)
                        saveFlag = True
                        Exit For
                    End If
                    If i > 100 Then
                        Exit For
                    End If
                Next
                If Not saveFlag Then
                    serverNewPath = serverPath & "\" & Path.GetFileName(serverNewPath)
                    SaveCsv(DGV1, serverNewPath, mode)
                End If
            Catch ex As Exception

            End Try
        End If

        If dialogON Then
            Me.Text = newPath
            Me.Invoke(New CallDelegate(AddressOf CB3get))
            'CB3get()
            MsgBox(newPath & vbCrLf & "保存しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    Public Sub SaveCsv(dgv As DataGridView, fp As String, mode As Integer)
        dgv.EndEdit()

        Dim CsvDelimiter As String = ","
        If mode = 2 Then
            CsvDelimiter = vbTab
        End If

        callTSP1_maximum = dgv.Rows.Count
        Me.Invoke(New CallDelegate(AddressOf TSProgressBarMaximum))
        callTSP1 = 0
        Me.Invoke(New CallDelegate(AddressOf TSProgressBar1))

        ' CSVファイルオープン
        Dim str As String = ""
        'Dim sw As StreamWriter = New StreamWriter(fp, False, EncodingName2(ToolStripDropDownButton10.Text))
        Dim sw As StreamWriter = New StreamWriter(fp, False, System.Text.Encoding.GetEncoding("SHIFT-JIS"))

        Dim dc As String = ""
        For c As Integer = 0 To dgv.ColumnCount - 1
            If c <> 0 Then
                dc &= ","
            End If
            dc &= """" & dgv.Columns(c).HeaderText & """"
        Next
        dc &= vbLf
        sw.Write(dc)

        For r As Integer = 0 To dgv.Rows.Count - 1
            If dgv.Rows(r).IsNewRow Then
                Continue For
            End If

            For c As Integer = 0 To dgv.Columns.Count - 1
                ' DataGridViewのセルのデータ取得
                Dim dt As String = ""
                If dgv.Rows(r).Cells(c).Value Is Nothing = False Then
                    dt = dgv.Rows(r).Cells(c).Value.ToString()
                    dt = Replace(dt, vbCrLf, vbLf)
                    'dt = Replace(dt, vbLf, "")
                    If mode = 1 Then
                        dt = """" & dt & """"
                    Else
                        If Not dt Is Nothing Then
                            dt = Replace(dt, """", """""")
                            Select Case True
                                Case InStr(dt, ","), InStr(dt, vbLf), InStr(dt, vbCr), InStr(dt, """") 'Regex.IsMatch(dt, ".*,.*|.*" & vbLf & ".*|.*" & vbCr & ".*|.*"".*")
                                    dt = """" & dt & """"
                            End Select
                        End If
                    End If
                End If
                If c < dgv.Columns.Count - 1 Then
                    dt = dt & CsvDelimiter
                End If
                ' CSVファイル書込
                sw.Write(dt)
            Next
            sw.Write(vbLf)

            Me.Invoke(New CallDelegate(AddressOf TSProgressBar1Plus))
            If ToolStripProgressBar1.Value Mod 500 = 0 Then
                Application.DoEvents()
            End If
        Next

        ' CSVファイルクローズ
        sw.Close()

        callTSP1 = 0
        Me.Invoke(New CallDelegate(AddressOf TSProgressBar1))
        callMeText = fp
        Me.Invoke(New CallDelegate(AddressOf MeText))
    End Sub

    Private TM1 As Boolean = True
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TM1 = False
        BackgroundWorker1.CancelAsync()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        selOn = False
        Button4.Enabled = True

        If BackgroundWorker1.IsBusy Then
            '保留中のバックグラウンド操作のキャンセルを要求します
            BackgroundWorker1.CancelAsync()
        End If

        bg_message = ""
        'バックグラウンド操作の実行を開始します
        BackgroundWorker1.RunWorkerAsync()

        'ToolStripProgressBar1.Value = 0
        'CountAll()
        'Button4.Enabled = False
        'TM1 = True
        'MsgBox("終了しました", MsgBoxStyle.SystemModal)
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        If Not BackgroundWorker1.IsBusy Then
            Exit Sub
        ElseIf UpdatePanel.Visible Then
            Exit Sub
        End If

        numberArray.Clear()
        dH1 = TM_HEADER_GET(DGV1)

        Dim tuiseki_url As String = ""
        'Dim kensakuTitle1 As String = ""
        'Dim kensakuTitle2 As String = ""
        Dim haitatuKaisya As String = ""
        Dim doc As HtmlAgilityPack.HtmlDocument = Nothing
        Dim kekka As String = ""

        Dim selRowArray As New ArrayList
        If CheckBox5.Checked Then
            For i As Integer = 0 To DGV1.SelectedCells.Count - 1
                Dim selRow As Integer = DGV1.SelectedCells(i).RowIndex
                If Not selRowArray.Contains(selRow) Then
                    selRowArray.Add(selRow)
                End If
            Next
        End If

        If CheckBox7.Checked Then
            Dim countNow As Integer = 0
            Dim kArraySagawa As New ArrayList
            Dim kArrayYubin As New ArrayList
            Dim kArrayYamato As New ArrayList
            Dim kArrayNasi As New ArrayList

            For r As Integer = 0 To DGV1.RowCount - 1
                '強制終了
                If TM1 = False Then
                    BackgroundWorker1.CancelAsync()
                    Exit For
                End If

                '除外
                If DGV1.Rows(r).IsNewRow Then
                    Continue For
                End If

                Dim lastFlag As Boolean = False
                If r = DGV1.RowCount - 2 Then
                    If kArraySagawa.Count > 0 Or kArrayYubin.Count > 0 Or kArrayYamato.Count > 0 Then
                        lastFlag = True
                    End If
                End If

                '除外（最後の行のみ配列にデータがあったら除外しない）
                Dim visibleFlag As Boolean = True
                If DGV1.Rows(r).Visible = False Then
                    visibleFlag = False
                ElseIf CheckBox1.Checked = True Then
                    If DGV1.Item(dH1.IndexOf("済"), r).Value <> "" Then
                        If Regex.IsMatch(DGV1.Item(dH1.IndexOf("済"), r).Value, "済|返送|対応確認|不明") Then
                            visibleFlag = False
                        End If
                    End If
                End If
                If CheckBox5.Checked = True Then
                    If Not selRowArray.Contains(r) Then
                        visibleFlag = False
                    End If
                End If

                If Not CheckBox2.Checked And InStr(DGV1.Item(dH1.IndexOf("配送会社"), r).Value, "佐川") > 0 Then
                    visibleFlag = False
                ElseIf Not CheckBox3.Checked And InStr(DGV1.Item(dH1.IndexOf("配送会社"), r).Value, "ゆうパケット") > 0 Then
                    visibleFlag = False
                ElseIf Not CheckBox4.Checked And InStr(DGV1.Item(dH1.IndexOf("配送会社"), r).Value, "ヤマト") > 0 Then
                    visibleFlag = False
                End If

                If Not lastFlag And Not visibleFlag Then
                    Continue For
                End If

                If CheckBox12.Checked Then
                    Dim checkTime As String = DGV1.Item(dH1.IndexOf("調査日"), r).Value
                    If checkTime <> "" Then
                        Dim a = DateDiff(DateInterval.Minute, CDate(checkTime), Now)
                        If DateDiff(DateInterval.Minute, CDate(checkTime), Now) < 10 Then
                            visibleFlag = False
                        End If
                    End If
                End If

                CountDisplay()

                If visibleFlag Then
                    Dim number As String = DGV1.Item(dH1.IndexOf("追跡番号"), r).Value
                    If DGV1.Item(dH1.IndexOf("配送会社"), r).Value = "" Then
                        haitatuKaisya = ""
                    ElseIf number = "" And Not Regex.IsMatch(DGV1.Item(dH1.IndexOf("配送会社"), r).Value, "定形外") Then
                        DGV1.Item(dH1.IndexOf("結果"), r).Value = "番号無し"
                        DGV1.Item(dH1.IndexOf("済"), r).Value = "済"
                        DGV1.Item(dH1.IndexOf("済"), r).Style.BackColor = Color.Yellow
                    ElseIf Regex.IsMatch(DGV1.Item(dH1.IndexOf("配送会社"), r).Value, "佐川") Then
                        haitatuKaisya = "佐川"
                        kArraySagawa.Add(r & "=" & number)
                    ElseIf Regex.IsMatch(DGV1.Item(dH1.IndexOf("配送会社"), r).Value, "ヤマト") Then
                        haitatuKaisya = "ヤマト"
                        kArrayYamato.Add(r & "=" & number)
                    ElseIf Regex.IsMatch(DGV1.Item(dH1.IndexOf("配送会社"), r).Value, "ゆうパケット") Then
                        haitatuKaisya = "郵便"
                        kArrayYubin.Add(r & "=" & number)
                    ElseIf Regex.IsMatch(DGV1.Item(dH1.IndexOf("配送会社"), r).Value, "定形外") Then
                        DGV1.Item(dH1.IndexOf("結果"), r).Value = "不可"
                        DGV1.Item(dH1.IndexOf("済"), r).Value = "済"
                        DGV1.Item(dH1.IndexOf("済"), r).Style.BackColor = Color.Yellow
                    Else
                        haitatuKaisya = ""
                    End If
                End If

                '佐川
                If kArraySagawa.Count > 9 Or (lastFlag And kArraySagawa.Count > 0) Then
                    Dim templeSagawa As String = Path.GetDirectoryName(Form1.appPath) & "\tuiseki_sagawa.html"
                    Dim templeSagawa_temp As String = Path.GetDirectoryName(Form1.appPath) & "\tuiseki_sagawa_temp.html"
                    Dim templeStr As String = File.ReadAllText(templeSagawa, Encoding.GetEncoding("utf-8"))
                    For i As Integer = 0 To 9
                        If i > kArraySagawa.Count - 1 Then
                            templeStr = Replace(templeStr, "_$" & i & "_", "")
                        Else
                            Dim kArray As String() = Split(kArraySagawa(i), "=")
                            templeStr = Replace(templeStr, "_$" & i & "_", kArray(1))
                        End If
                    Next
                    File.WriteAllText(templeSagawa_temp, templeStr)

                    '****************
                    webReadComplete = False
                    kensakuTitle2 = "【お荷物問い合わせサービス】"
                    WebBrowser1.Navigate(templeSagawa_temp)
                    'Threading.Thread.Sleep(1000 * 1)
                    Application.DoEvents()
                    Dim html As String = ""
                    Dim doFlag As Boolean = True
                    Do While doFlag = True
                        If InStr(wb_txt, Split(kArraySagawa(0), "=")(1)) > 0 And InStr(wb_title, kensakuTitle2) > 0 Then
                            html = wb_txt
                            html = WebUtility.HtmlDecode(html)
                            doFlag = False
                        Else
                            Threading.Thread.Sleep(1000 * 1)
                            Application.DoEvents()
                        End If
                    Loop
                    '****************

                    doc = New HtmlAgilityPack.HtmlDocument
                    doc.LoadHtml(html)

                    'デバッグ
                    Me.Invoke(New CallDelegate(AddressOf WB_TXT_DEBUG))
                    Dim wb_txt2s As String = WebUtility.HtmlDecode(wb_txt2)
                    doc.LoadHtml(wb_txt2s)

                    Dim domArray As ArrayList = HTMLgetArray(doc, New String() {"td", "th", "dd"})
                    For i As Integer = 0 To kArraySagawa.Count - 1
                        Dim kArray As String() = Split(kArraySagawa(i), "=")
                        Dim nowRow As Integer = CInt(kArray(0))
                        Dim numCount As Integer = 0
                        For k As Integer = 0 To domArray.Count - 1
                            If domArray(k) = kArray(1) Then
                                callR = nowRow
                                Me.Invoke(New CallDelegate(AddressOf DGV1_CurrentCell))
                                Application.DoEvents()

                                If numCount = 0 Then
                                    DGV1.Item(dH1.IndexOf("結果"), nowRow).Value = domArray(k + 1)
                                    If Regex.IsMatch(DGV1.Item(dH1.IndexOf("結果"), nowRow).Value, "配達完了") Then
                                        DGV1.Item(dH1.IndexOf("済"), nowRow).Value = "済"
                                        DGV1.Item(dH1.IndexOf("済"), nowRow).Style.BackColor = Color.Yellow
                                    ElseIf Regex.IsMatch(DGV1.Item(dH1.IndexOf("結果"), nowRow).Value, "返送") Then
                                        DGV1.Item(dH1.IndexOf("済"), nowRow).Value = "返送"
                                        DGV1.Item(dH1.IndexOf("済"), nowRow).Style.BackColor = Color.Orange
                                    End If
                                    Dim kanryo As String = domArray(k + 3)
                                    If kanryo.Length < 15 Then
                                        DGV1.Item(dH1.IndexOf("配達完了日"), nowRow).Value = kanryo
                                    End If
                                    If DGV1.Item(dH1.IndexOf("初回反応"), nowRow).Value Is Nothing Then
                                        If Not Regex.IsMatch(DGV1.Item(dH1.IndexOf("結果"), nowRow).Value, "該当なし") Then
                                            DGV1.Item(dH1.IndexOf("初回反応"), nowRow).Value = Format(Now, "yyyy/MM/dd HH:mm:ss")
                                        End If
                                    End If
                                ElseIf numCount = 1 Then
                                    DGV1.Item(dH1.IndexOf("出荷日"), nowRow).Value = domArray(k + 2)
                                    DGV1.Item(dH1.IndexOf("届け先"), nowRow).Value = domArray(k + 6)
                                    DGV1.Item(dH1.IndexOf("個数"), nowRow).Value = domArray(k + 8)
                                    DGV1.Item(dH1.IndexOf("調査日"), nowRow).Value = Format(Now, "yyyy/MM/dd HH:mm:ss")
                                    If InStr(DGV1.Item(dH1.IndexOf("結果"), nowRow).Value, "登録されておりません") = 0 Then
                                        Dim syousai As String = ""
                                        For p As Integer = 0 To 100
                                            Dim pNum As Integer = (k + 11) + (p * 3)
                                            If pNum + 2 > domArray.Count - 1 Then
                                                Exit For
                                            End If
                                            If Not IsNumeric(domArray(pNum)) And (InStr(domArray(pNum), "登録されておりません") = 0 And InStr(domArray(pNum), "お問い合わせ") = 0) Then
                                                syousai &= domArray(pNum) & " " & domArray(pNum + 1) & " " & domArray(pNum + 2) & "||"
                                            Else
                                                Exit For
                                            End If
                                        Next
                                        DGV1.Item(dH1.IndexOf("詳細"), nowRow).Value = syousai
                                    End If
                                Else
                                    Exit For
                                End If
                                numCount += 1
                            End If
                        Next
                    Next

                    kArraySagawa.Clear()
                    Threading.Thread.Sleep(500 * 1)
                End If

                '郵便
                If kArrayYubin.Count > 9 Or (lastFlag And kArrayYubin.Count > 0) Then
                    Dim templeYubin As String = Path.GetDirectoryName(Form1.appPath) & "\tuiseki_yubin.html"
                    Dim templeYubin_temp As String = Path.GetDirectoryName(Form1.appPath) & "\tuiseki_yubin_temp.html"
                    Dim templeStr As String = File.ReadAllText(templeYubin, Encoding.GetEncoding("utf-8"))
                    '追跡番号1つだと詳細検索に飛ぶのでNo2にダミーを入れて検索
                    If kArrayYubin.Count = 1 Then
                        Dim kArray As String() = Split(kArrayYubin(0), "=")
                        templeStr = Replace(templeStr, "_$0_", kArray(1))
                        templeStr = Replace(templeStr, "_$1_", "0")
                    Else
                        For i As Integer = 0 To 9
                            If i > kArrayYubin.Count - 1 Then
                                templeStr = Replace(templeStr, "_$" & i & "_", "")
                            Else
                                Dim kArray As String() = Split(kArrayYubin(i), "=")
                                templeStr = Replace(templeStr, "_$" & i & "_", kArray(1))
                            End If
                        Next
                    End If
                    File.WriteAllText(templeYubin_temp, templeStr)

                    '****************
                    webReadComplete = False
                    kensakuTitle2 = "個別番号検索結果"
                    WebBrowser1.Navigate(templeYubin_temp)
                    'Threading.Thread.Sleep(1000 * 1)
                    Application.DoEvents()
                    Dim html As String = ""
                    Dim doFlag As Boolean = True
                    Do While doFlag = True
                        Dim tNo1 As String = Split(kArrayYubin(0), "=")(1)
                        Dim tNo2 As String = tNo1.Substring(0, 4) & "-" & tNo1.Substring(4, 4) & "-" & tNo1.Substring(8, 4)
                        If (InStr(wb_txt, tNo2) > 0 Or InStr(wb_txt, tNo1) > 0) And InStr(wb_title, kensakuTitle2) > 0 Then
                            html = wb_txt
                            html = WebUtility.HtmlDecode(html)
                            doFlag = False
                        Else
                            Threading.Thread.Sleep(500 * 1)
                            Application.DoEvents()
                        End If
                    Loop
                    '****************

                    doc = New HtmlAgilityPack.HtmlDocument
                    doc.LoadHtml(html)

                    'デバッグ
                    Me.Invoke(New CallDelegate(AddressOf WB_TXT_DEBUG))
                    Dim wb_txt2s As String = WebUtility.HtmlDecode(wb_txt2)
                    doc.LoadHtml(wb_txt2s)

                    Dim domArray As ArrayList = HTMLgetArray(doc, New String() {"td"})
                    For i As Integer = 0 To kArrayYubin.Count - 1
                        Dim kArray As String() = Split(kArrayYubin(i), "=")
                        Dim nowRow As Integer = CInt(kArray(0))
                        For k As Integer = 0 To domArray.Count - 1
                            If domArray(k) = kArray(1) Then
                                callR = nowRow
                                Me.Invoke(New CallDelegate(AddressOf DGV1_CurrentCell))
                                Application.DoEvents()

                                If InStr(domArray(k + 1), "見つかりません") > 0 Then
                                    DGV1.Item(dH1.IndexOf("結果"), nowRow).Value = "該当なし"
                                Else
                                    DGV1.Item(dH1.IndexOf("結果"), nowRow).Value = domArray(k + 3)
                                    If Regex.IsMatch(DGV1.Item(dH1.IndexOf("結果"), nowRow).Value, "お届け済|窓口でお渡し") Then
                                        DGV1.Item(dH1.IndexOf("済"), nowRow).Value = "済"
                                        DGV1.Item(dH1.IndexOf("済"), nowRow).Style.BackColor = Color.Yellow
                                    ElseIf Regex.IsMatch(DGV1.Item(dH1.IndexOf("結果"), nowRow).Value, "返送") Then
                                        DGV1.Item(dH1.IndexOf("済"), nowRow).Value = "返送"
                                        DGV1.Item(dH1.IndexOf("済"), nowRow).Style.BackColor = Color.Orange
                                    End If
                                    If DGV1.Item(dH1.IndexOf("初回反応"), nowRow).Value Is Nothing Then
                                        If Not Regex.IsMatch(DGV1.Item(dH1.IndexOf("結果"), nowRow).Value, "該当なし") Then
                                            DGV1.Item(dH1.IndexOf("初回反応"), nowRow).Value = Format(Now, "yyyy/MM/dd HH:mm:ss")
                                        End If
                                    End If
                                    Dim hiduke As String = domArray(k + 2)
                                    If hiduke.Length > 10 Then
                                        hiduke = hiduke.Substring(0, 10) & " " & hiduke.Substring(10, 5)
                                    End If
                                    If Regex.IsMatch(DGV1.Item(dH1.IndexOf("結果"), nowRow).Value, "お届け済|窓口でお渡し") Then
                                        DGV1.Item(dH1.IndexOf("配達完了日"), nowRow).Value = hiduke
                                    End If
                                    If DGV1.Item(dH1.IndexOf("出荷日"), nowRow).Value = "" And Regex.IsMatch(DGV1.Item(dH1.IndexOf("結果"), nowRow).Value, "引受") Then
                                        DGV1.Item(dH1.IndexOf("出荷日"), nowRow).Value = hiduke
                                    ElseIf DGV1.Item(dH1.IndexOf("出荷日"), nowRow).Value = "" Then
                                        DGV1.Item(dH1.IndexOf("出荷日"), nowRow).Value = DGV1.Item(dH1.IndexOf("出荷確定日"), nowRow).Value
                                    End If
                                    DGV1.Item(dH1.IndexOf("届け先"), nowRow).Value = domArray(k + 4) & "(" & domArray(k + 5) & ")"
                                    DGV1.Item(dH1.IndexOf("個数"), nowRow).Value = ""
                                End If
                                DGV1.Item(dH1.IndexOf("調査日"), nowRow).Value = Format(Now, "yyyy/MM/dd HH:mm:ss")
                            End If
                        Next
                    Next

                    kArrayYubin.Clear()
                    Threading.Thread.Sleep(500 * 1)
                End If

                'ヤマト
                If kArrayYamato.Count > 9 Or (lastFlag And kArrayYamato.Count > 0) Then
                    Dim templeYamato As String = Path.GetDirectoryName(Form1.appPath) & "\tuiseki_yamato.html"
                    Dim templeYamato_temp As String = Path.GetDirectoryName(Form1.appPath) & "\tuiseki_yamato_temp.html"
                    Dim templeStr As String = File.ReadAllText(templeYamato, Encoding.GetEncoding("shift-jis"))
                    For i As Integer = 0 To 9
                        If i > kArrayYamato.Count - 1 Then
                            templeStr = Replace(templeStr, "_$" & i & "_", "")
                        Else
                            Dim kArray As String() = Split(kArrayYamato(i), "=")
                            templeStr = Replace(templeStr, "_$" & i & "_", kArray(1))
                        End If
                    Next
                    File.WriteAllText(templeYamato_temp, templeStr)

                    '****************
                    webReadComplete = False
                    kensakuTitle2 = "お問い合わせシステム"
                    WebBrowser1.Navigate(templeYamato_temp)
                    'Threading.Thread.Sleep(1000 * 1)
                    Application.DoEvents()
                    Dim html As String = ""
                    Dim doFlag As Boolean = True
                    Do While doFlag = True
                        If InStr(wb_txt, Split(kArrayYamato(0), "=")(1)) > 0 And InStr(wb_title, kensakuTitle2) > 0 Then
                            html = wb_txt
                            html = WebUtility.HtmlDecode(html)
                            doFlag = False
                        Else
                            Threading.Thread.Sleep(1000 * 1)
                            Application.DoEvents()
                        End If
                    Loop
                    '****************

                    doc = New HtmlAgilityPack.HtmlDocument
                    doc.LoadHtml(html)

                    'デバッグ
                    Me.Invoke(New CallDelegate(AddressOf WB_TXT_DEBUG))
                    Dim wb_txt2s As String = WebUtility.HtmlDecode(wb_txt2)
                    doc.LoadHtml(wb_txt2s)

                    Dim domArray As ArrayList = HTMLgetArray(doc, New String() {"td"})
                    For i As Integer = 0 To kArrayYamato.Count - 1
                        Dim kArray As String() = Split(kArrayYamato(i), "=")
                        Dim yamatoNo As String = kArray(1).ToString.Substring(0, 4) & "-" & kArray(1).ToString.Substring(4, 4) & "-" & kArray(1).ToString.Substring(8, 4)
                        Dim nowRow As Integer = CInt(kArray(0))
                        Dim numCount As Integer = 0
                        For k As Integer = 0 To domArray.Count - 1
                            If domArray(k) = kArray(1) Or domArray(k) = "伝票番号" & kArray(1) Then
                                callR = nowRow
                                Me.Invoke(New CallDelegate(AddressOf DGV1_CurrentCell))
                                Application.DoEvents()

                                If numCount = 0 Then
                                    DGV1.Item(dH1.IndexOf("結果"), nowRow).Value = domArray(k + 2)
                                    If Regex.IsMatch(DGV1.Item(dH1.IndexOf("結果"), nowRow).Value, "完了") Then
                                        DGV1.Item(dH1.IndexOf("済"), nowRow).Value = "済"
                                        DGV1.Item(dH1.IndexOf("済"), nowRow).Style.BackColor = Color.Yellow
                                    ElseIf Regex.IsMatch(DGV1.Item(dH1.IndexOf("結果"), nowRow).Value, "返送") Then
                                        DGV1.Item(dH1.IndexOf("済"), nowRow).Value = "返送"
                                        DGV1.Item(dH1.IndexOf("済"), nowRow).Style.BackColor = Color.Orange
                                    End If
                                    DGV1.Item(dH1.IndexOf("配達完了日"), nowRow).Value = domArray(k + 1)
                                    If DGV1.Item(dH1.IndexOf("初回反応"), nowRow).Value Is Nothing Then
                                        If Not Regex.IsMatch(DGV1.Item(dH1.IndexOf("結果"), nowRow).Value, "該当なし") Then
                                            DGV1.Item(dH1.IndexOf("初回反応"), nowRow).Value = Format(Now, "yyyy/MM/dd HH:mm:ss")
                                        End If
                                    End If
                                ElseIf numCount = 1 Then
                                    DGV1.Item(dH1.IndexOf("出荷日"), nowRow).Value = domArray(k + 8)
                                    DGV1.Item(dH1.IndexOf("個数"), nowRow).Value = ""
                                    DGV1.Item(dH1.IndexOf("調査日"), nowRow).Value = Format(Now, "yyyy/MM/dd HH:mm:ss")
                                    If InStr(DGV1.Item(dH1.IndexOf("結果"), nowRow).Value, "誤り") = 0 Then
                                        Dim syousai As String = domArray(k + 5) & "||"  'クロネコDM便
                                        Dim basyo As String = ""
                                        For p As Integer = 0 To 100
                                            Dim pNum As Integer = (k + 7) + (p * 7)
                                            If pNum > domArray.Count - 1 Then
                                                Exit For
                                            End If
                                            If Not IsNumeric(domArray(pNum)) And InStr(domArray(pNum), "誤り") = 0 And domArray(pNum) <> "" And InStr(domArray(pNum), "<") = 0 Then
                                                syousai &= domArray(pNum) & " " & domArray(pNum + 1) & " " & domArray(pNum + 2) & " " & domArray(pNum + 3) & " " & domArray(pNum + 4) & "||"
                                                basyo = domArray(pNum + 3)
                                            Else
                                                Exit For
                                            End If
                                        Next
                                        DGV1.Item(dH1.IndexOf("詳細"), nowRow).Value = syousai
                                        DGV1.Item(dH1.IndexOf("届け先"), nowRow).Value = basyo
                                    End If
                                Else
                                    Exit For
                                End If
                                numCount += 1
                            End If
                        Next
                    Next

                    kArraySagawa.Clear()
                    Threading.Thread.Sleep(500 * 1)
                End If

                countNow += 1
                If countNow Mod 50 = 0 Then
                    CountAll()
                End If
                Application.DoEvents()
            Next
            '***************************************************************************************************
            '***************************************************************************************************
            '***************************************************************************************************
        Else    '1行ずつ検索
            Dim countNow As Integer = 0
            For r As Integer = 0 To DGV1.RowCount - 1
                '強制終了
                If TM1 = False Then
                    BackgroundWorker1.CancelAsync()
                    Exit For
                End If

                '除外
                If DGV1.Rows(r).IsNewRow Then
                    Continue For
                ElseIf DGV1.Rows(r).Visible = False Then
                    Continue For
                ElseIf CheckBox1.Checked = True Then
                    If DGV1.Item(dH1.IndexOf("済"), r).Value <> "" Then
                        If Regex.IsMatch(DGV1.Item(dH1.IndexOf("済"), r).Value, "済|返送|対応確認|不明") Then
                            Continue For
                        End If
                    End If
                ElseIf CheckBox5.Checked = True Then
                    If Not selRowArray.Contains(r) Then
                        Continue For
                    End If
                End If

                If Not CheckBox2.Checked And InStr(DGV1.Item(dH1.IndexOf("配送会社"), r).Value, "佐川") > 0 Then
                    Continue For
                ElseIf Not CheckBox3.Checked And InStr(DGV1.Item(dH1.IndexOf("配送会社"), r).Value, "ゆうパケット") > 0 Then
                    Continue For
                ElseIf Not CheckBox4.Checked And InStr(DGV1.Item(dH1.IndexOf("配送会社"), r).Value, "ヤマト") > 0 Then
                    Continue For
                End If

                compCount = 0
                webReadComplete = False
                callR = r
                Me.Invoke(New CallDelegate(AddressOf DGV1_CurrentCell))
                'DGV1.CurrentCell = DGV1(dH1.IndexOf("結果"), r)
                CountDisplay()

                Dim number As String = DGV1.Item(dH1.IndexOf("追跡番号"), r).Value
                If Regex.IsMatch(DGV1.Item(dH1.IndexOf("配送会社"), r).Value, "佐川") Then
                    tuiseki_url = sagawa & number
                    'kensakuTitle1 = "佐川急便"
                    'kensakuTitle2 = "【お荷物問い合わせサービス】"
                    haitatuKaisya = "佐川"
                    'enc = System.Text.Encoding.GetEncoding("csWindows31J")
                    enc = System.Text.Encoding.GetEncoding("utf-8")
                ElseIf Regex.IsMatch(DGV1.Item(dH1.IndexOf("配送会社"), r).Value, "ヤマト") Then
                    tuiseki_url = yamato & number
                    'kensakuTitle1 = "クロネコヤマト"
                    'kensakuTitle2 = "荷物お問い合わせシステム"
                    haitatuKaisya = "ヤマト"
                    enc = System.Text.Encoding.GetEncoding("csWindows31J")
                ElseIf Regex.IsMatch(DGV1.Item(dH1.IndexOf("配送会社"), r).Value, "ゆうパケット") Then
                    tuiseki_url = yubin1 & number & yubin2
                    'kensakuTitle1 = "検索結果"
                    'kensakuTitle2 = "日本郵便"
                    haitatuKaisya = "郵便"
                    enc = System.Text.Encoding.GetEncoding("utf-8")
                Else
                    haitatuKaisya = ""
                End If

                ToolStripStatusLabel1.Text = number
                callTSP1 = 10
                Me.Invoke(New CallDelegate(AddressOf TSProgressBar1))
                'ToolStripProgressBar1.Value = 10

                If Regex.IsMatch(haitatuKaisya, "佐川|郵便") Then
                    Dim html As String = ""
                    ServicePointManager.SecurityProtocol = DirectCast(3072, SecurityProtocolType)
                    'Dim req As WebRequest = WebRequest.Create(tuiseki_url)
                    'req.UseDefaultCredentials = False

                    Dim req As HttpWebRequest = CType(WebRequest.Create(tuiseki_url), HttpWebRequest)
                    req.Timeout = 10000
                    req.KeepAlive = True
                    req.AllowAutoRedirect = True
                    req.PreAuthenticate = True
                    req.UserAgent = "Mozilla/4.0 (compatible; MSIE 6.0; Windows XP)"
                    Dim cookieContainer As CookieContainer = New CookieContainer()
                    req.CookieContainer = cookieContainer
                    Dim resW As WebResponse = req.GetResponse()
                    Dim st As Stream = resW.GetResponseStream()
                    Dim sr As StreamReader = New StreamReader(st, enc)

                    html = sr.ReadToEnd()
                    html = WebUtility.HtmlDecode(html)
                    sr.Close()
                    st.Close()
                    doc = New HtmlAgilityPack.HtmlDocument
                    doc.LoadHtml(html)
                    callStr3 = html
                    Me.Invoke(New CallDelegate(AddressOf TBText3))
                    'TextBox3.Text = html
                    Application.DoEvents()
                ElseIf Regex.IsMatch(haitatuKaisya, "ヤマト") Then
                    WebBrowser1.Navigate(tuiseki_url)

                    Dim doCount As Integer = 0
                    Do While webReadComplete = False
                        If doCount > 100 Then
                            If ToolStripProgressBar1.Value < 90 Then
                                ToolStripProgressBar1.Value += 1
                            Else
                                ToolStripProgressBar1.Value = 0
                            End If
                            doCount = 0
                        Else
                            doCount += 1
                        End If
                        Application.DoEvents()
                        'System.Threading.Thread.Sleep(1000)
                    Loop
                Else
                    kekka = "不可"
                    DGV1.Item(dH1.IndexOf("結果"), r).Value = kekka
                    DGV1.Item(dH1.IndexOf("済"), r).Value = "済"
                    DGV1.Item(dH1.IndexOf("済"), r).Style.BackColor = Color.Yellow
                End If

                Dim res = ""
                If haitatuKaisya = "佐川" Then
                    res = HTMLgetTxt(doc, "th", 5, 0, "")(0)
                    If InStr(res, "該当なし") > 0 Or res = "error" Then
                        kekka = res
                        DGV1.Item(dH1.IndexOf("結果"), r).Value = kekka
                    Else
                        kekka = res
                        kekka = Regex.Replace(kekka, " |　", "")
                        DGV1.Item(dH1.IndexOf("結果"), r).Value = kekka
                        If Regex.IsMatch(kekka, "配達完了") Then
                            DGV1.Item(dH1.IndexOf("済"), r).Value = "済"
                            DGV1.Item(dH1.IndexOf("済"), r).Style.BackColor = Color.Yellow
                        ElseIf Regex.IsMatch(kekka, "返送") Then
                            DGV1.Item(dH1.IndexOf("済"), r).Value = "返送"
                            DGV1.Item(dH1.IndexOf("済"), r).Style.BackColor = Color.Orange
                        End If
                        Dim todokesaki As String = HTMLgetTxt(doc, "td", 4, 0, "")(0)
                        todokesaki = Regex.Replace(todokesaki, "[A-Za-z0-9\:\- ]", "")
                        DGV1.Item(dH1.IndexOf("届け先"), r).Value = todokesaki
                        DGV1.Item(dH1.IndexOf("出荷日"), r).Value = HTMLgetTxt(doc, "td", 2, 0, "")(0)
                        Dim kanryou As String = HTMLgetTxt(doc, "dd", 0, 0, "")(0)
                        If kanryou.Length < 15 Then
                            kanryou = Regex.Replace(kanryou, "　| ", "")
                            DGV1.Item(dH1.IndexOf("配達完了日"), r).Value = kanryou
                            Dim kosuu As String = HTMLgetTxt(doc, "td", 5, 0, "")(0)
                            DGV1.Item(dH1.IndexOf("個数"), r).Value = kosuu
                        End If
                        DGV1.Item(dH1.IndexOf("調査日"), r).Value = Format(Now, "yyyy/MM/dd HH:mm:ss")
                        If DGV1.Item(dH1.IndexOf("初回反応"), r).Value Is Nothing Then
                            DGV1.Item(dH1.IndexOf("初回反応"), r).Value = Format(Now, "yyyy/MM/dd HH:mm:ss")
                        End If
                    End If
                ElseIf haitatuKaisya = "郵便" Then
                    Dim resArray As String() = HTMLgetTxt(doc, "td", 3, 6, "郵便局")
                    If InStr(resArray(0), "見つかりません") > 0 Or resArray(0) = "error" Then
                        kekka = "該当なし"
                        DGV1.Item(dH1.IndexOf("結果"), r).Value = kekka
                    Else
                        kekka = resArray(0)
                        kekka = Regex.Replace(kekka, " |　", "")
                        DGV1.Item(dH1.IndexOf("結果"), r).Value = kekka
                        If Regex.IsMatch(kekka, "お届け済|窓口でお渡し") Then
                            DGV1.Item(dH1.IndexOf("済"), r).Value = "済"
                            DGV1.Item(dH1.IndexOf("済"), r).Style.BackColor = Color.Yellow
                        ElseIf Regex.IsMatch(kekka, "返送") Then
                            DGV1.Item(dH1.IndexOf("済"), r).Value = "返送"
                            DGV1.Item(dH1.IndexOf("済"), r).Style.BackColor = Color.Orange
                        End If
                        DGV1.Item(dH1.IndexOf("出荷日"), r).Value = HTMLgetTxt(doc, "td", 2, 0, "")(0)
                        'DGV1.Item(dH1.IndexOf("届け先"), r).Value = HTMLgetTxt(doc, "td", 5, 6, "")
                        'DGV1.Item(dH1.IndexOf("配達完了日"), r).Value = HTMLgetTxt(doc, "td", 2, 6, kekka)
                        DGV1.Item(dH1.IndexOf("届け先"), r).Value = HTMLgetTxt(doc, "td", resArray(1) + 2, 0, "")(0) & "(" & HTMLgetTxt(doc, "td", resArray(1) + 3, 0, "")(0) & ")"
                        If Regex.IsMatch(kekka, "お届け済|窓口でお渡し") Then
                            DGV1.Item(dH1.IndexOf("配達完了日"), r).Value = HTMLgetTxt(doc, "td", resArray(1) - 1, 0, kekka)(0)
                        End If
                        DGV1.Item(dH1.IndexOf("調査日"), r).Value = Format(Now, "yyyy/MM/dd HH:mm:ss")
                        If DGV1.Item(dH1.IndexOf("初回反応"), r).Value Is Nothing Then
                            DGV1.Item(dH1.IndexOf("初回反応"), r).Value = Format(Now, "yyyy/MM/dd HH:mm:ss")
                        End If
                    End If
                ElseIf haitatuKaisya = "ヤマト" Then
                    res = HTMLgetTxt2("td", 15, 0, "")
                    If InStr(res, "未登録") > 0 Or res = "error" Then
                        kekka = res
                    Else
                        kekka = res
                    End If
                    DGV1.Item(dH1.IndexOf("結果"), r).Value = kekka
                End If

                numberArray.Add(number)
                Application.DoEvents()

                callTSP1 = 100
                Me.Invoke(New CallDelegate(AddressOf TSProgressBar1))
                'ToolStripProgressBar1.Value = 100
                Threading.Thread.Sleep(500)

                countNow += 1
                If countNow Mod 50 = 0 Then
                    CountAll()
                End If
            Next
        End If
    End Sub

    Dim callR As Integer = 0
    Dim callStr7 As String = ""
    Dim callTSP1 As Integer = 0
    Dim callTSP1_maximum As Integer = 100
    Dim callStr3 As String = ""
    Dim CounterD As Integer = 0
    Dim callCB1 As String = ""
    Private Delegate Sub CallDelegate()
    Private Sub DGV1_CurrentCell()
        DGV1.CurrentCell = DGV1(dH1.IndexOf("結果"), callR)
    End Sub
    Private Sub TBText7()
        TextBox7.Text = callStr7
    End Sub
    Private Sub TBText8()
        TextBox8.Text = CounterD
    End Sub
    Private Sub TSProgressBar1()
        ToolStripProgressBar1.Value = callTSP1
    End Sub
    Private Sub TSProgressBarMaximum()
        ToolStripProgressBar1.Maximum = callTSP1_maximum
    End Sub
    Private Sub TSProgressBar1Plus()
        ToolStripProgressBar1.Value += 1
    End Sub
    Private Sub TBText3()
        TextBox3.Text = callStr3
    End Sub

    'Dim CounterKeys As String() = New String() {"all", "sumi", "hannouAri", "taiou", "notSumi", "gaitouNasi", "kekkaNasi", "cansel", "hensou", "mi"}
    'Dim TB_Counter As TextBox() = {TextBox1, TextBox2, TextBox6, TextBox14, TextBox4, TextBox5, TextBox10, TextBox15, TextBox13, TextBox16}
    Private Sub TBText_Counter()
        For i As Integer = 0 To TB_Counter.Length - 1
            TB_Counter(i).Text = Counter(CounterKeys(i))
        Next

        'TextBox1.Text = Counter("all")
        'TextBox2.Text = Counter("sumi")
        'TextBox6.Text = Counter("hannouAri")
        'TextBox14.Text = Counter("taiou")

        'TextBox4.Text = Counter("notSumi")
        'TextBox5.Text = Counter("gaitouNasi")
        'TextBox10.Text = Counter("kekkaNasi")
        'TextBox15.Text = Counter("cansel")

        'TextBox13.Text = Counter("hensou")
        'TextBox16.Text = Counter("mi")
    End Sub
    Private Sub TBText_Counter_Reset()
        For i As Integer = 0 To TB_Counter.Length - 1
            TB_Counter(i).Text = 0
        Next
    End Sub
    Private Sub CB1_clear()
        ComboBox1.Items.Clear()
    End Sub
    Private Sub CB1_add()
        ComboBox1.Items.Add(callCB1)
    End Sub
    Private Sub CB1_select()
        ComboBox1.SelectedIndex = 0
    End Sub
    Dim wb_txt2 As String = ""
    Private Sub WB_TXT_DEBUG()
        wb_txt2 = WebBrowser1.Document.Body.Parent.OuterHtml
    End Sub

    Private bg_message As String = ""
    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        If bg_message = "" Then
            ToolStripProgressBar1.Value = 0
            CountAll()
            Button4.Enabled = False
            TM1 = True
            selOn = True

            If SplitContainer2.Panel2Collapsed = False Then
                NissuIchiran()
            End If

            MsgBox("終了しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    Private Function HTMLgetTxt(ByVal doc As HtmlAgilityPack.HtmlDocument, ByVal elName As String, ByVal elNum As Integer, ByVal elList As Integer, ByVal notList As String) As String()
        Dim res As String = ""
        Dim resB As String = ""
        Dim resC As Integer = 0
        Dim resX As String() = {"", ""}

        Dim selecter As String = ""
        selecter = "//" & elName
        Dim nodes As HtmlAgilityPack.HtmlNodeCollection = doc.DocumentNode.SelectNodes(selecter)

        If elList = 0 Then
            Try
                res = nodes(elNum).InnerText
                resC = elNum
            Catch ex As Exception
                res = "error"
            End Try
        Else
            For i As Integer = 0 To 100
                Try
                    'Dim resA As String = WebBrowser1.Document.GetElementsByTagName(elName)(elNum + (elList * i)).InnerText
                    Dim resA As String = nodes(elNum + (elList * i)).InnerText
                    If notList <> "" Then
                        If InStr(resA, notList) = 0 Then
                            res = resA
                            resC = elNum + (elList * i)
                        End If
                    Else
                        res = resA
                        resC = elNum + (elList * i)
                    End If
                Catch ex As Exception
                    resB = "error"
                End Try
                If i <> 0 And resB = "error" Then
                    Exit For
                ElseIf i = 0 And resB = "error" Then
                    res = "error"
                End If
            Next
        End If

        If res <> "" Then
            res = Regex.Replace(res, delMoji, "")
        End If

        resX = {res, resC}
        Return resX
    End Function

    Private Function HTMLgetTxt2(elName As String, elNum As Integer, elList As Integer, notList As String) As String
        Dim res As String = ""
        Dim resB As String = ""
        If elList = 0 Then
            Try
                res = WebBrowser1.Document.GetElementsByTagName(elName)(elNum).InnerText
            Catch ex As Exception
                res = "error"
            End Try
        Else    'ゆうメはリストなので最新を取得
            For i As Integer = 0 To 100
                Try
                    Dim resA As String = WebBrowser1.Document.GetElementsByTagName(elName)(elNum + (elList * i)).InnerText
                    If InStr(resA, notList) = 0 Then
                        res = resA
                    End If
                Catch ex As Exception
                    resB = "error"
                End Try
                If i <> 0 And resB = "error" Then
                    Exit For
                ElseIf i = 0 And resB = "error" Then
                    res = "error"
                End If
            Next
        End If
        Return res
    End Function

    Private Function HTMLgetArray(doc As HtmlAgilityPack.HtmlDocument, elName As String())
        Dim selecter As String = ""
        For k As Integer = 0 To elName.Length - 1
            If k <> 0 Then
                selecter &= " | "
            End If
            selecter &= "//" & elName(k)
        Next
        Dim nodes As HtmlAgilityPack.HtmlNodeCollection = doc.DocumentNode.SelectNodes(selecter)
        If nodes Is Nothing Then
            Return Nothing
            Exit Function
        End If

        Dim resArray As New ArrayList
        For i As Integer = 0 To nodes.Count - 1
            Try
                Dim res As String = nodes(i).InnerText
                If res <> "" Then
                    res = Regex.Replace(res, vbCr & "|" & vbLf & "|" & vbTab & "| |　|\-", "")
                End If
                resArray.Add(res)
            Catch ex As Exception
                resArray.Add("error")
            End Try
        Next

        Return resArray
    End Function

    Private wb_title As String = ""
    Private wb_txt As String = ""
    Private Sub WebBrowser1_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted
        'compCount += 1

        'ToolStripTextBox1.Text = WebBrowser1.Url.ToString

        'If compCount = 1 Then
        '    MsgBox(kensakuTitle1 & " " & kensakuTitle2 & vbCrLf & WebBrowser1.DocumentTitle, MsgBoxStyle.SystemModal)
        'End If
        'If haitatu = 3 Then
        '    If InStr(WebBrowser1.DocumentTitle, kensakuTitle1) > 0 And InStr(WebBrowser1.DocumentTitle, kensakuTitle2) > 0 And numberArray.Contains(ToolStripStatusLabel1.Text) = False Then
        '        webReadComplete = True
        '        System.Threading.Thread.Sleep(1000)
        '    End If
        'Else
        '    If compCount = 1 And numberArray.Contains(ToolStripStatusLabel1.Text) = False Then
        '        webReadComplete = True
        '        System.Threading.Thread.Sleep(1000)
        '    End If
        'End If

        'wb_title = WebBrowser1.DocumentTitle
        'If WebBrowser1.DocumentTitle <> "" Then
        '    If Regex.IsMatch(WebBrowser1.DocumentTitle, kensakuTitle2) Then
        '        Dim dammy1 As String = WebBrowser1.DocumentTitle
        '        Do While InStr(WebBrowser1.DocumentText.ToLower, "</html>") = 0
        '            Application.DoEvents()
        '            Dim dammy2 As String = WebBrowser1.DocumentText
        '            '読み込み待ち
        '        Loop
        '        'wb_txt = WebBrowser1.DocumentText
        '        wb_txt = WebBrowser1.Document.Body.OuterHtml
        '        wb_txt = WebUtility.HtmlDecode(wb_txt)
        '        webReadComplete = True
        '    End If
        'End If
    End Sub

    Private Sub WebBrowser1_DocumentTitleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles WebBrowser1.DocumentTitleChanged
        'compCount += 1

        wb_title = WebBrowser1.DocumentTitle
        If WebBrowser1.DocumentTitle <> "" Then
            If Regex.IsMatch(WebBrowser1.DocumentTitle, kensakuTitle2) Then
                Dim dammy1 As String = WebBrowser1.DocumentTitle
                Do While InStr(WebBrowser1.DocumentText.ToLower, "</html>") = 0
                    Application.DoEvents()
                    Dim dammy2 As String = WebBrowser1.DocumentText
                    '読み込み待ち
                Loop
                'wb_txt = WebBrowser1.DocumentText
                If WebBrowser1.Document.Body Is Nothing Then
                    wb_txt = ""
                Else
                    wb_txt = WebBrowser1.Document.Body.Parent.OuterHtml
                    wb_txt = WebUtility.HtmlDecode(wb_txt)
                End If
                webReadComplete = True
            End If
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        selOn = False
        Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)

        ToolStripProgressBar1.Value = 0
        ToolStripProgressBar1.Maximum = DGV1.RowCount

        For r As Integer = 0 To DGV1.RowCount - 1
            If DGV1.Rows(r).IsNewRow Then
                Continue For
            End If

            If CheckBox6.Checked And InStr(DGV1.Item(dH1.IndexOf("配送会社"), r).Value, "定形外") > 0 Then
                DGV1.Rows(r).Visible = False
            ElseIf ComboBox1.SelectedIndex = 0 Or DGV1.Item(dH1.IndexOf("店舗"), r).Value = ComboBox1.SelectedItem Then
                If ComboBox2.SelectedIndex = 0 Or InStr(DGV1.Item(dH1.IndexOf("配送会社"), r).Value, ComboBox2.SelectedItem) > 0 Then
                    If RadioButton2.Checked Then
                        If DGV1.Item(dH1.IndexOf("済"), r).Value = "済" Then
                            DGV1.Rows(r).Visible = True
                        Else
                            DGV1.Rows(r).Visible = False
                        End If
                    ElseIf RadioButton3.Checked Then
                        If DGV1.Item(dH1.IndexOf("済"), r).Value = "済" Then
                            DGV1.Rows(r).Visible = False
                        Else
                            DGV1.Rows(r).Visible = True
                        End If
                    ElseIf RadioButton4.Checked Then
                        If DGV1.Item(dH1.IndexOf("結果"), r).Value = "" Then
                            DGV1.Rows(r).Visible = False
                        ElseIf Regex.IsMatch(DGV1.Item(dH1.IndexOf("結果"), r).Value, "該当なし|error") Then
                            DGV1.Rows(r).Visible = True
                        Else
                            DGV1.Rows(r).Visible = False
                        End If
                    ElseIf RadioButton5.Checked Then
                        If DGV1.Item(dH1.IndexOf("初回反応"), r).Value = "" Then
                            DGV1.Rows(r).Visible = False
                        Else
                            DGV1.Rows(r).Visible = True
                        End If
                    ElseIf RadioButton6.Checked Then
                        If DGV1.Item(dH1.IndexOf("結果"), r).Value = "" Then
                            DGV1.Rows(r).Visible = True
                        Else
                            DGV1.Rows(r).Visible = False
                        End If
                    ElseIf RadioButton7.Checked Then
                        If DGV1.Item(dH1.IndexOf("済"), r).Value = "返送" Then
                            DGV1.Rows(r).Visible = True
                        Else
                            DGV1.Rows(r).Visible = False
                        End If
                    ElseIf RadioButton8.Checked Then
                        If DGV1.Item(dH1.IndexOf("済"), r).Value = "不明" Then
                            DGV1.Rows(r).Visible = True
                        Else
                            DGV1.Rows(r).Visible = False
                        End If
                    ElseIf RadioButton9.Checked Then
                        If DGV1.Item(dH1.IndexOf("済"), r).Value = "対応確認" Then
                            DGV1.Rows(r).Visible = True
                        Else
                            DGV1.Rows(r).Visible = False
                        End If
                    ElseIf RadioButton10.Checked Then
                        If InStr(DGV1.Item(dH1.IndexOf("状態"), r).Value, "キャンセル") > 0 Then
                            DGV1.Rows(r).Visible = True
                        Else
                            DGV1.Rows(r).Visible = False
                        End If
                    ElseIf RadioButton11.Checked Then
                        If DGV1.Item(dH1.IndexOf("済"), r).Value = "" Then
                            DGV1.Rows(r).Visible = True
                        Else
                            DGV1.Rows(r).Visible = False
                        End If
                    Else
                        DGV1.Rows(r).Visible = True
                    End If
                Else
                    DGV1.Rows(r).Visible = False
                End If
            Else
                DGV1.Rows(r).Visible = False
            End If

            If r Mod 100 = 0 Then
                ToolStripProgressBar1.Value = r
            End If
        Next

        ToolStripProgressBar1.Value = 0

        For r As Integer = 0 To DGV1.RowCount - 1
            If DGV1.Rows(r).IsNewRow Then
                Continue For
            End If

            If DGV1.Rows(r).Visible Then
                Dim CBArray As CheckBox() = {CheckBox8, CheckBox9, CheckBox10, CheckBox11, CheckBox13}
                Dim ColorArray As Color() = {Color.Yellow, Color.Orange, Color.Red, Color.LightBlue, Color.PaleGreen}
                Dim flagA As Boolean = True
                For i As Integer = 0 To CBArray.Length - 1
                    If CBArray(i).Checked And flagA Then
                        Dim flagB As Boolean = False
                        For c As Integer = 0 To DGV1.ColumnCount - 1
                            If DGV1.Item(c, r).Style.BackColor = ColorArray(i) Then
                                flagB = True
                                Exit For
                            End If
                        Next
                        If Not flagB Then
                            flagA = False
                        End If
                    End If
                Next

                If Not flagA Then
                    DGV1.Rows(r).Visible = False
                End If
            End If
            ToolStripProgressBar1.Value = r
        Next

        ToolStripProgressBar1.Value = 0
        ToolStripProgressBar1.Maximum = 100

        CountDisplay()
        selOn = True
    End Sub

    Private Sub DGV3_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV3.CellDoubleClick
        If e.ColumnIndex >= 1 Then
            Select Case e.ColumnIndex
                Case 1
                    RadioButton1.Checked = True
                Case 2
                    RadioButton2.Checked = True
                Case 3
                    RadioButton11.Checked = True
                Case 4
                    RadioButton4.Checked = True
                Case 5
                    RadioButton9.Checked = True
                Case 6
                    RadioButton7.Checked = True
                Case 7
                    RadioButton10.Checked = True
            End Select
            Select Case e.RowIndex
                Case 0
                    ComboBox2.SelectedIndex = 1
                    CheckBox6.Checked = True
                Case 1
                    ComboBox2.SelectedIndex = 2
                    CheckBox6.Checked = True
                Case 2
                    ComboBox2.SelectedIndex = 3
                    CheckBox6.Checked = False
                Case 3
                    ComboBox2.SelectedIndex = 4
                    CheckBox6.Checked = True
            End Select
            Button3.PerformClick()
        End If
    End Sub

    Private Sub DGV3_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DGV3.ColumnHeaderMouseClick
        If e.ColumnIndex >= 1 Then
            Select Case e.ColumnIndex
                Case 1
                    RadioButton1.Checked = True
                Case 2
                    RadioButton2.Checked = True
                Case 3
                    RadioButton11.Checked = True
                Case 4
                    RadioButton4.Checked = True
                Case 5
                    RadioButton9.Checked = True
                Case 6
                    RadioButton7.Checked = True
                Case 7
                    RadioButton10.Checked = True
            End Select
            ComboBox2.SelectedIndex = 0
            CheckBox6.Checked = True
            Button3.PerformClick()
        Else
            Exit Sub
        End If
    End Sub

    Private Sub CountDisplay()
        CounterD = 0
        For r As Integer = 0 To DGV1.RowCount - 2
            If DGV1.Rows(r).Visible Then
                CounterD += 1
                If r = DGV1.SelectedCells(0).RowIndex Then
                    callStr7 = CounterD
                    Me.Invoke(New CallDelegate(AddressOf TBText7))
                    'TextBox7.Text = Counter
                End If
            End If
        Next
        'TextBox8.Text = CounterD
        Me.Invoke(New CallDelegate(AddressOf TBText8))
    End Sub

    Dim Counter As New Hashtable
    Dim CounterKeys As String() = New String() {"all", "sumi", "hannouAri", "taiou", "notSumi", "gaitouNasi", "kekkaNasi", "cansel", "hensou", "mi", "humei"}
    Dim haisouRow As String() = New String() {"佐川", "ゆうパケット", "定形外", "ヤマト"}
    Private Sub CountAll()
        Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
        For Each cK As String In CounterKeys
            Counter(cK) = 0
        Next
        For r As Integer = 0 To DGV3.RowCount - 1
            For c As Integer = 1 To DGV3.ColumnCount - 1
                DGV3.Item(c, r).Value = 0
            Next
        Next

        For r As Integer = 0 To DGV1.RowCount - 1
            If DGV1.Rows(r).IsNewRow Then
                Continue For
            End If

            '配送会社
            Dim hR As Integer = 0
            For i As Integer = 0 To haisouRow.Length - 1
                If InStr(DGV1.Item(dH1.IndexOf("配送会社"), r).Value, haisouRow(i)) > 0 Then
                    hR = i
                    Exit For
                End If
            Next

            Counter("all") += 1
            DGV3.Item(dH3.IndexOf("全"), hR).Value += 1

            If DGV1.Item(dH1.IndexOf("済"), r).Value = "済" Then
                Counter("sumi") += 1
                DGV3.Item(dH3.IndexOf("済"), hR).Value += 1
            ElseIf DGV1.Item(dH1.IndexOf("済"), r).Value = "返送" Then
                Counter("hensou") += 1
                DGV3.Item(dH3.IndexOf("返"), hR).Value += 1
            ElseIf DGV1.Item(dH1.IndexOf("済"), r).Value = "対応確認" Then
                Counter("taiou") += 1
                DGV3.Item(dH3.IndexOf("対"), hR).Value += 1
            ElseIf DGV1.Item(dH1.IndexOf("済"), r).Value = "不明" Then
                Counter("humei") += 1
                DGV3.Item(dH3.IndexOf("不"), hR).Value += 1
            End If
            If DGV1.Item(dH1.IndexOf("済"), r).Value <> "済" Then
                Counter("notSumi") += 1
            End If
            If DGV1.Item(dH1.IndexOf("済"), r).Value = "" Then
                Counter("mi") += 1
                DGV3.Item(dH3.IndexOf("未"), hR).Value += 1
            End If
            If DGV1.Item(dH1.IndexOf("結果"), r).Value <> "" Then
                If Regex.IsMatch(DGV1.Item(dH1.IndexOf("結果"), r).Value, "該当なし|error") Then
                    If DGV1.Item(dH1.IndexOf("済"), r).Value = "" Then
                        Counter("gaitouNasi") += 1
                        DGV3.Item(dH3.IndexOf("該"), hR).Value += 1
                    ElseIf Not Regex.IsMatch(DGV1.Item(dH1.IndexOf("済"), r).Value, "対応確認|不明") Then
                        Counter("gaitouNasi") += 1
                        DGV3.Item(dH3.IndexOf("該"), hR).Value += 1
                    End If
                End If
            Else
                Counter("kekkaNasi") += 1
            End If
            If DGV1.Item(dH1.IndexOf("初回反応"), r).Value <> "" Then
                Counter("hannouAri") += 1
            End If
            If InStr(DGV1.Item(dH1.IndexOf("状態"), r).Value, "キャンセル") > 0 Then
                Counter("cansel") += 1
                DGV3.Item(dH3.IndexOf("C"), hR).Value += 1
            End If
        Next

        For r As Integer = 0 To DGV3.RowCount - 1
            For c As Integer = 3 To DGV3.ColumnCount - 1
                If DGV3.Item(c, r).Value > 0 Then
                    DGV3.Item(c, r).Style.BackColor = Color.LightSkyBlue
                Else
                    DGV3.Item(c, r).Style.BackColor = Color.Empty
                End If
            Next
        Next

        Me.Invoke(New CallDelegate(AddressOf TBText_Counter))
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim searchCount As Integer = 0
        For r As Integer = 0 To DGV1.RowCount - 2
            Dim vFlag As Boolean = False
            For c As Integer = 0 To DGV1.ColumnCount - 1
                Dim v As String = DGV1.Item(c, r).Value
                If v <> "" Then
                    If Regex.IsMatch(v, TextBox12.Text) Then
                        DGV1.Rows(r).Visible = True
                        searchCount += 1
                        vFlag = True
                        Exit For
                    End If
                End If
            Next
            If Not vFlag Then
                DGV1.Rows(r).Visible = False
            End If
        Next

        If searchCount = 0 Then
            MsgBox("検索結果はありませんでした")
        End If
    End Sub

    Private selOn As Boolean = True
    Private Sub DGV1_SelectionChanged(sender As Object, e As EventArgs) Handles DGV1.SelectionChanged
        If selOn And DGV1.SelectedCells.Count > 0 Then
            TextBox11.Text = DGV1.SelectedCells(0).Value
        End If
    End Sub

    '#########################################################
    '特殊検索
    '#########################################################
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Matome_kensaku(0)
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Matome_kensaku(1)
    End Sub

    Private Sub Matome_kensaku(mode As Integer)
        DGV1.Rows.Clear()
        Application.DoEvents()

        Dim nowCSV As String = ""
        Dim files As String() = Directory.GetFiles(dPath, "*.csv", System.IO.SearchOption.AllDirectories)
        Array.Reverse(files)
        For k As Integer = 0 To ComboBox3.Items.Count - 1
            For i As Integer = 0 To files.Length - 1
                Dim fName As String = Path.GetFileNameWithoutExtension(files(i))
                Dim cb As String() = Split(ComboBox3.Items(k), "(全")
                If InStr(fName, cb(0)) > 0 Then
                    If Regex.IsMatch(cb(1), "該|対|返|C") Then
                        Me.Text = files(i)

                        Dim csvRecords As New ArrayList()
                        csvRecords = TM_CSV_READ(files(i))(0)

                        Dim csvH As String() = Split(csvRecords(0), "|=|")

                        ToolStripProgressBar1.Value = 0
                        ToolStripProgressBar1.Maximum = csvRecords.Count
                        For p As Integer = 1 To csvRecords.Count - 1
                            Dim sArray As String() = Split(csvRecords(p), "|=|")
                            Dim flag As Boolean = False
                            If mode = 0 And InStr(sArray(Array.IndexOf(csvH, "結果")), "該当なし") > 0 Then
                                If Regex.IsMatch(sArray(Array.IndexOf(csvH, "済")), "対応|返送|不明|確認") Then
                                    flag = False
                                Else
                                    flag = True
                                End If
                            ElseIf mode = 1 Then
                                If InStr(sArray(Array.IndexOf(csvH, "結果")), "該当なし") > 0 Then
                                    flag = True
                                End If
                                If Not flag And sArray(Array.IndexOf(csvH, "済")) <> "" Then
                                    If Regex.IsMatch(sArray(Array.IndexOf(csvH, "済")), "対応|返送") Then
                                        flag = True
                                    End If
                                End If
                                If Not flag And InStr(sArray(Array.IndexOf(csvH, "状態")), "キャンセル") > 0 Then
                                    flag = True
                                End If
                            End If

                            If flag Then
                                DGV1.Rows.Add()
                                Dim newR As Integer = DGV1.RowCount - 2
                                For c As Integer = 0 To pattern2.Length - 1
                                    DGV1.Item(dH1.IndexOf(pattern2(c)), newR).Value = sArray(Array.IndexOf(csvH, pattern2(c)))
                                Next
                            End If

                            ToolStripProgressBar1.Value += 1
                        Next

                        ToolStripProgressBar1.Value = 0
                        ToolStripProgressBar1.Maximum = 100

                        Exit For
                    End If
                End If
            Next
        Next

        Me.Text = ""

        CB1get()
        CountDisplay()
        CountAll()
    End Sub


    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If BackgroundWorker2.IsBusy Then
            BackgroundWorker2.CancelAsync()
        End If
        BackgroundWorker2.RunWorkerAsync()
    End Sub

    Private Sub BackgroundWorker2_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        selOn = False

        Dim nowCSV As String = ""
        Dim files As String() = Directory.GetFiles(dPath, "*.csv", System.IO.SearchOption.AllDirectories)
        Array.Reverse(files)
        For k As Integer = ComboBox3.Items.Count - 1 To 0 Step -1
            For i As Integer = 0 To files.Length - 1
                Dim fName As String = Path.GetFileNameWithoutExtension(files(i))
                Dim cb As String() = Split(ComboBox3.Items(k), "(全")
                If InStr(fName, cb(0)) > 0 Then
                    If InStr(cb(1), "該0") = 0 Or InStr(cb(1), "未0") = 0 Then
                        If InStr(cb(0), "完") = 0 And RadioButton13.Checked Then
                            Continue For
                        ElseIf InStr(cb(0), "完") > 0 And RadioButton12.Checked Then
                            Continue For
                        End If

                        Me.Invoke(New CallDelegate(AddressOf TBText_Counter_Reset))
                        Me.Invoke(New CallDelegate(AddressOf DGV1_RowsClear))
                        callMeText = files(i)
                        Me.Invoke(New CallDelegate(AddressOf MeText))

                        Dim csvRecords As New ArrayList()
                        csvRecords = TM_CSV_READ(files(i))(0)

                        Dim csvH As String() = Split(csvRecords(0), "|=|")

                        callTSP1 = 0
                        Me.Invoke(New CallDelegate(AddressOf TSProgressBar1))
                        callTSP1_maximum = csvRecords.Count
                        Me.Invoke(New CallDelegate(AddressOf TSProgressBarMaximum))
                        For p As Integer = 1 To csvRecords.Count - 1
                            Dim sArray As String() = Split(csvRecords(p), "|=|")
                            Me.Invoke(New CallDelegate(AddressOf DGV1_RowsAdd))
                            Dim newR As Integer = DGV1.RowCount - 2
                            For c As Integer = 0 To pattern2.Length - 1
                                If csvH.Contains(pattern2(c)) Then
                                    DGV1.Item(dH1.IndexOf(pattern2(c)), newR).Value = sArray(Array.IndexOf(csvH, pattern2(c)))
                                End If
                            Next
                            Me.Invoke(New CallDelegate(AddressOf TSProgressBar1Plus))
                        Next

                        CountAll()
                        Me.Invoke(New CallDelegate(AddressOf RB3))
                        Me.Invoke(New CallDelegate(AddressOf Button3_PerformClick))

                        callTSP1 = 0
                        Me.Invoke(New CallDelegate(AddressOf TSProgressBar1))
                        ToolStripProgressBar1.Maximum = 100

                        Me.Invoke(New CallDelegate(AddressOf Button4_Enabled))
                        bg_message = "jidou"
                        If BackgroundWorker1.IsBusy Then
                            BackgroundWorker1.CancelAsync()
                        End If
                        BackgroundWorker1.RunWorkerAsync()

                        Do While BackgroundWorker1.IsBusy
                            '完了待ち
                        Loop
                        Threading.Thread.Sleep(1000 * 2)

                        If TM1 = False Then
                            BackgroundWorker2.CancelAsync()
                            Exit Sub
                        End If

                        Exit For
                    End If
                End If
            Next

            CountAll()
            Application.DoEvents()
            NameChangeSave(1, False)
        Next

        CB1get()
        'CountDisplay()
        CountAll()
        Me.Invoke(New CallDelegate(AddressOf CB3get))
        Me.Invoke(New CallDelegate(AddressOf CB3save))

        selOn = True
        MsgBox("終了しました")
        BackgroundWorker2.CancelAsync()
    End Sub

    Dim callMeText As String = ""
    Private Sub MeText()
        Me.Text = callMeText
    End Sub
    Private Sub DGV1_RowsClear()
        DGV1.Rows.Clear()
    End Sub
    Private Sub DGV1_RowsAdd()
        DGV1.Rows.Add()
    End Sub
    Private Sub RB3()
        RadioButton3.Checked = True
    End Sub
    Private Sub Button3_PerformClick()
        Button3.PerformClick()
    End Sub
    Private Sub Button4_Enabled()
        Button4.Enabled = True
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If BackgroundWorker1.IsBusy Then
            ProgressBar1.Value += 1
            If ProgressBar1.Value = ProgressBar1.Maximum Then
                ProgressBar1.Value = 0
            End If
        End If
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        NeListCheck(0)
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        NeListCheck(1)
    End Sub

    Private Sub NeListCheck(mode As Integer)
        '.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
        '.RestoreDirectory = True,
        Dim ofd As New OpenFileDialog With {
            .InitialDirectory = dPath,
            .Filter = "csvファイル(*.csv)|*.csv|すべてのファイル(*.*)|*.*",
            .FilterIndex = 1,
            .Title = "チェックするファイルを選択してください",
            .CheckFileExists = True,
            .CheckPathExists = True
        }

        If ofd.ShowDialog() = DialogResult.OK Then
            selOn = False
            Dim cancelFile As String = ofd.FileName

            Dim files As String() = Directory.GetFiles(dPath, "*.csv", System.IO.SearchOption.AllDirectories)
            Array.Reverse(files)
            Dim cancelCount As Integer = 0
            For k As Integer = ComboBox3.Items.Count - 1 To 0 Step -1
                For i As Integer = 0 To files.Length - 1
                    Dim fName As String = Path.GetFileNameWithoutExtension(files(i))
                    Dim cb As String() = Split(ComboBox3.Items(k), "(全")
                    If InStr(fName, cb(0)) > 0 Then
                        Me.Invoke(New CallDelegate(AddressOf TBText_Counter_Reset))
                        Me.Invoke(New CallDelegate(AddressOf DGV1_RowsClear))
                        callMeText = files(i)
                        Me.Invoke(New CallDelegate(AddressOf MeText))

                        Dim csvRecords As New ArrayList()
                        csvRecords = TM_CSV_READ(files(i))(0)

                        Dim csvH As String() = Split(csvRecords(0), "|=|")

                        callTSP1 = 0
                        Me.Invoke(New CallDelegate(AddressOf TSProgressBar1))
                        callTSP1_maximum = csvRecords.Count
                        Me.Invoke(New CallDelegate(AddressOf TSProgressBarMaximum))
                        For p As Integer = 1 To csvRecords.Count - 1
                            Dim sArray As String() = Split(csvRecords(p), "|=|")
                            Me.Invoke(New CallDelegate(AddressOf DGV1_RowsAdd))
                            Dim newR As Integer = DGV1.RowCount - 2
                            For c As Integer = 0 To pattern2.Length - 1
                                If csvH.Contains(pattern2(c)) Then
                                    DGV1.Item(dH1.IndexOf(pattern2(c)), newR).Value = sArray(Array.IndexOf(csvH, pattern2(c)))
                                End If
                            Next
                            Me.Invoke(New CallDelegate(AddressOf TSProgressBar1Plus))
                        Next

                        CountAll()
                        Application.DoEvents()
                        Dim cancelCountA As String = TextBox15.Text

                        DropCsvCheck(cancelFile)
                        CountAll()
                        Application.DoEvents()
                        Dim cancelCountB As String = TextBox15.Text

                        If mode = 0 Then
                            NameChangeSave(1, False)
                        ElseIf mode = 1 And cancelCountA <> cancelCountB Then
                            cancelCount += CInt(cancelCountB) - CInt(cancelCountA)
                            NameChangeSave(1, False)
                        End If
                    End If
                Next
            Next

            Me.Invoke(New CallDelegate(AddressOf CB3get))
            Me.Invoke(New CallDelegate(AddressOf CB3save))

            selOn = True
            ToolStripProgressBar1.Value = 0
            If mode = 0 Then
                MsgBox(cancelCount & "終了しました。")
            Else
                MsgBox(cancelCount & "件修正しました。")
            End If
        End If
    End Sub

    '##################################################################################

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        TabControl1.SelectedTab = TabPage2
        Application.DoEvents()

        Try
            If BackgroundWorker3.IsBusy Then
                BackgroundWorker3.CancelAsync()
            End If
            BackgroundWorker3.RunWorkerAsync()
        Catch ex As Exception
            MsgBox("くまを再起動してください")
        End Try

        TabControl1.SelectedTab = TabPage3
    End Sub

    Private Sub BackgroundWorker3_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        selOn = False
        Call_LB2visible = True
        Me.Invoke(New CallDelegate(AddressOf LB2_visible))

        'ログイン確認
        Dim defaultURL As String = "https://base.next-engine.org/apps/launch/?id=65459"
        kensakuTitle2 = "TOP|ネクストエンジン"
        WebBrowser1.Navigate(defaultURL)
        Me.Invoke(New CallDelegate(AddressOf WB_getDocTxt))
        Do While InStr(WB_doc, "ネクストエンジン") = 0
            '読み込み待ち
            Threading.Thread.Sleep(1000 * 1)
            Application.DoEvents()
            Me.Invoke(New CallDelegate(AddressOf WB_getDocTxt))
        Loop
        Dim html As String = WebUtility.HtmlDecode(WB_doc)
        If InStr(html, "売れ筋ランキング") > 0 Then
            'ログインOK
        Else
            webReadComplete = False
            Me.Invoke(New CallDelegate(AddressOf WB_login))
            Do While Not Regex.IsMatch(WB_doc, "ログインしました|新着のお知らせ一覧")
                Me.Invoke(New CallDelegate(AddressOf WB_getDocTxt))
                Application.DoEvents()
                '読み込み待ち
                Threading.Thread.Sleep(1000 * 1)
            Loop
        End If

        'Dim doc As HtmlAgilityPack.HtmlDocument = Nothing
        For r As Integer = 0 To DGV1.RowCount - 1
            If Not DGV1.Rows(r).Visible Or DGV1.Rows(r).IsNewRow Then
                Continue For
            End If

            If Label2.BackColor = Color.Khaki Then
                Call_LB2color = Color.Yellow
                Me.Invoke(New CallDelegate(AddressOf LB2_color))
            Else
                Call_LB2color = Color.Khaki
                Me.Invoke(New CallDelegate(AddressOf LB2_color))
            End If

            Dim selRow As Integer = r
            Dim dNo As String = DGV1.Item(dH1.IndexOf("伝票番号"), selRow).Value
            Dim url As String = "https://ne81.next-engine.com/Userjyuchu/jyuchuInp?kensaku_denpyo_no="
            url &= dNo
            url &= "&jyuchu_meisai_order=jyuchu_meisai_gyo"

            callR = selRow
            Me.Invoke(New CallDelegate(AddressOf DGV1_CurrentCell))

            '****************
            webReadComplete = False
            kensakuTitle2 = "伝票番号"
            WebBrowser1.Navigate(url)
            Application.DoEvents()
            Dim doFlag As Boolean = True
            Do While doFlag = True
                Me.Invoke(New CallDelegate(AddressOf WB_getDocTxt))
                If InStr(WB_doc, dNo) > 0 And InStr(WB_doc, "</html>") > 0 Then
                    html = WB_doc
                    html = WebUtility.HtmlDecode(html)
                    doFlag = False
                Else
                    Threading.Thread.Sleep(1000 * 1)
                    Application.DoEvents()
                End If
            Loop
            '****************

            'doc = New HtmlAgilityPack.HtmlDocument
            'doc.LoadHtml(html)

            Call_WBdata1_row = r
            Me.Invoke(New CallDelegate(AddressOf WB_data1))
        Next

        Call_LB2visible = False
        Me.Invoke(New CallDelegate(AddressOf LB2_visible))
        CountAll()
        selOn = True
    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        MsgBox("終了しました", MsgBoxStyle.SystemModal)
    End Sub

    Private Sub WB_login()
        kensakuTitle2 = "ネクストエンジン"
        WebBrowser1.Document.GetElementById("user_login_code").SetAttribute("value", "jidou2@yongshun006.com")
        WebBrowser1.Document.GetElementById("user_password").SetAttribute("value", "aaa111777-")
        WebBrowser1.Document.Forms(0).InvokeMember("submit")
    End Sub
    Private WB_doc As String = ""
    Private Sub WB_getDocTxt()
        WB_doc = WebBrowser1.DocumentText
    End Sub
    Private Call_WBdata1_row As Integer = 0
    Private Sub WB_data1()
        Dim hdNo As String = WebBrowser1.Document.GetElementById("hasou_denpyo_no").GetAttribute("value")
        Dim picS As String = WebBrowser1.Document.GetElementById("pic_siji_naiyou").InnerText
        Dim joutai As String = WebBrowser1.Document.GetElementById("jyuchu_jyotai_kbn").GetAttribute("value")
        Dim joutai_cansel As String = WebBrowser1.Document.GetElementById("jyuchu_cancel_kbn").GetAttribute("value")
        If hdNo <> "" Then
            hdNo = Regex.Replace(hdNo, "-", "")
        End If
        If InStr(hdNo, DGV1.Item(dH1.IndexOf("追跡番号"), Call_WBdata1_row).Value) = 0 Then
            DGV1.Item(dH1.IndexOf("追跡番号"), Call_WBdata1_row).Value = hdNo
            DGV1.Item(dH1.IndexOf("追跡番号"), Call_WBdata1_row).Style.BackColor = Color.Orange
        End If
        If picS <> "" Then
            picS = Regex.Replace(picS, vbCrLf & "|" & vbCr & "|" & vbLf, "")
        End If
        If DGV1.Item(dH1.IndexOf("ﾋﾟｯｷﾝｸﾞ指示内容"), Call_WBdata1_row).Value <> picS Then
            DGV1.Item(dH1.IndexOf("ﾋﾟｯｷﾝｸﾞ指示内容"), Call_WBdata1_row).Value = picS
            DGV1.Item(dH1.IndexOf("ﾋﾟｯｷﾝｸﾞ指示内容"), Call_WBdata1_row).Style.BackColor = Color.Orange
        End If
        If InStr(DGV1.Item(dH1.IndexOf("状態"), Call_WBdata1_row).Value, "キャンセル") = 0 And joutai_cansel <> 0 Then
            DGV1.Item(dH1.IndexOf("状態"), Call_WBdata1_row).Value = "キャンセル"
            DGV1.Item(dH1.IndexOf("状態"), Call_WBdata1_row).Style.BackColor = Color.Red
        ElseIf InStr(DGV1.Item(dH1.IndexOf("状態"), Call_WBdata1_row).Value, "キャンセル") > 0 And joutai_cansel = 0 Then
            If joutai = 50 Then
                DGV1.Item(dH1.IndexOf("状態"), Call_WBdata1_row).Value = "出荷確定済（完了）"
                DGV1.Item(dH1.IndexOf("状態"), Call_WBdata1_row).Style.BackColor = Color.Orange
            Else
                For i As Integer = 0 To joutaiArray.Length - 1
                    Dim jA As String() = Split(joutaiArray(i), ":")
                    If joutai = jA(0) Then
                        DGV1.Item(dH1.IndexOf("状態"), Call_WBdata1_row).Value = jA(1)
                        DGV1.Item(dH1.IndexOf("状態"), Call_WBdata1_row).Style.BackColor = Color.Orange
                        Exit For
                    End If
                Next
            End If
        End If

        'MsgBox(hdNo & vbCrLf & picS)
    End Sub
    Private Call_LB2visible As Boolean = True
    Private Call_LB2color As Color = Color.Yellow
    Private Sub LB2_visible()
        Label2.Visible = Call_LB2visible
    End Sub
    Private Sub LB2_color()
        Label2.BackColor = Call_LB2color
    End Sub

    Private joutaiArray As String() = New String() {"0:取込情報不足", "1:受注メール取込済", "2:起票済(CSV/手入力)", "20:納品書印刷待ち", "30:納品書印刷中", "40:納品書印刷済", "50:出荷確定済（完了）"}

    '##################################################################################

    Private Sub 詳細をToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 詳細をToolStripMenuItem.Click
        Dim selRow As Integer = DGV1.SelectedCells(0).RowIndex
        Dim url As String = ""
        Select Case True
            Case InStr(DGV1.Item(dH1.IndexOf("配送会社"), selRow).Value, "佐川") > 0
                url = sagawa & DGV1.Item(dH1.IndexOf("追跡番号"), selRow).Value
            Case InStr(DGV1.Item(dH1.IndexOf("配送会社"), selRow).Value, "ゆうパケット") > 0
                url = yubin1 & DGV1.Item(dH1.IndexOf("追跡番号"), selRow).Value & yubin2
            Case InStr(DGV1.Item(dH1.IndexOf("配送会社"), selRow).Value, "ヤマト") > 0
                url = yamato & DGV1.Item(dH1.IndexOf("追跡番号"), selRow).Value
            Case Else
                Exit Sub
        End Select
        Process.Start(url)
    End Sub

    Private Sub ネクストエンジン情報をブラウザで開くToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ネクストエンジン情報をブラウザで開くToolStripMenuItem.Click
        Dim selRow As Integer = DGV1.SelectedCells(0).RowIndex
        Dim url As String = "https://ne81.next-engine.com/Userjyuchu/jyuchuInp?kensaku_denpyo_no="
        url &= DGV1.Item(dH1.IndexOf("伝票番号"), selRow).Value
        url &= "&jyuchu_meisai_order=jyuchu_meisai_gyo"
        Process.Start(url)
    End Sub


    Private Sub 切り取りToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 切り取りToolStripMenuItem1.Click
        CUTS(DGV1)
    End Sub

    Private Sub コピーToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles コピーToolStripMenuItem1.Click
        COPYS(DGV1, 1)
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        COPYS(DGV1, 0)
    End Sub

    Private Sub 貼り付けToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 貼り付けToolStripMenuItem1.Click
        Dim selCell = DGV1.SelectedCells
        PASTES(DGV1, selCell)
    End Sub

    Private Sub 列選択ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 列選択ToolStripMenuItem.Click
        ColSelect(0)
    End Sub

    ' ----------------------------------------------------------------------------------
    ' DataGridViewでCtrl-Vキー押下時にクリップボードから貼り付けるための実装↓↓↓↓↓↓
    ' DataGridViewでDelやBackspaceキー押下時にセルの内容を消去するための実装↓↓↓↓↓↓
    ' ----------------------------------------------------------------------------------
    Public Shared ColumnChars() As String
    Private _editingColumn As Integer   '編集中のカラム番号
    Private _editingCtrl As DataGridViewTextBoxEditingControl

    Private Sub DataGridView_EditingControlShowing(
        ByVal sender As Object, ByVal e As DataGridViewEditingControlShowingEventArgs) _
        Handles DGV1.EditingControlShowing
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

    Private Sub DataGridView_KeyUp(ByVal sender As Object, ByVal e As Windows.Forms.KeyEventArgs) Handles DGV1.KeyUp
        Dim dgv As DataGridView = CType(sender, DataGridView)
        Dim selCell = dgv.SelectedCells

        If e.KeyCode = Keys.Back Or e.KeyCode = Keys.Delete Then    ' セルの内容を消去
            If Not DGV1.IsCurrentCellInEditMode Then
                DELS(dgv, selCell)
            End If
        ElseIf e.Modifiers = Keys.Control And e.KeyCode = Keys.Enter Then '複数セル一括入力
            Dim str As String = dgv.CurrentCell.Value
            For i As Integer = 0 To selCell.Count - 1
                If dgv.Rows(selCell(i).RowIndex).Visible = True Then    '表示している行のみ
                    If selCell(i).Value <> str Then
                        selCell(i).Value = str
                        selCell(i).Style.BackColor = Color.Yellow
                    End If
                End If
            Next
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
    Private Sub DataGridView1_EditingControlShowing(ByVal sender As Object, ByVal e As Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles DGV1.EditingControlShowing
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

    Dim CMS1Array As String() = New String() {
        "佐川急便(宅配便)|ゆうパケット|ヤマト(DM便)|定形外郵便",
        "済|返送|対応確認|不明"
        }

    Private Sub ContextMenuStrip1_Opening(sender As Object, e As CancelEventArgs) Handles ContextMenuStrip1.Opening
        For i As Integer = ContextMenuStrip1.Items.Count - 1 To 0 Step -1
            If Not CMS1_kihon.Contains(ContextMenuStrip1.Items(i)) Then
                ContextMenuStrip1.Items.RemoveAt(i)
            End If
        Next

        Dim selCol As Integer = DGV1.SelectedCells(0).ColumnIndex

        Select Case selCol
            Case dH1.IndexOf("配送会社")
                Dim addTxt As String() = Split(CMS1Array(0), "|")
                For Each aTxt As String In addTxt
                    ContextMenuStrip1.Items.Add(aTxt)
                Next
            Case dH1.IndexOf("済")
                Dim addTxt As String() = Split(CMS1Array(1), "|")
                For Each aTxt As String In addTxt
                    ContextMenuStrip1.Items.Add(aTxt)
                Next
        End Select
    End Sub

    Private Sub ContextMenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles ContextMenuStrip1.ItemClicked
        If Not CMS1_kihon.Contains(e.ClickedItem) Then
            For i As Integer = 0 To DGV1.SelectedCells.Count - 1
                DGV1.SelectedCells(i).Value = e.ClickedItem.ToString
            Next
        End If
    End Sub

    Private Sub ColSelect(ByVal mode As Integer)
        selOn = False

        Dim dgv As DataGridView = DGV1
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
            For r As Integer = start To dgv.RowCount - 1
                If dgv.Rows(r).IsNewRow Then
                    Continue For
                End If
                If dgv.Rows(r).Visible = True Then
                    dgv.Item(selCol(i), r).selected = True
                End If
            Next
        Next

        selOn = True
    End Sub

    '↑↑=======================================================================↑↑
    'ContextMenuStripここまで
    '===============================================================================

    Private Sub Tuiseki_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If e.CloseReason = CloseReason.FormOwnerClosing Or e.CloseReason = CloseReason.UserClosing Then
            Me.Visible = False
            e.Cancel = True
        End If
    End Sub


    Private Sub リセットToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles リセットToolStripMenuItem.Click
        Dim frm As Form = New Tuiseki
        Me.Dispose()
        frm.Show()
    End Sub

    Private prefecture As String() = New String() {"北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県", "茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川", "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県", "岐阜県", "静岡県", "愛知県", "三重県", "滋賀県", "京都府", "大阪府", "兵庫県", "奈良県", "和歌山", "鳥取県", "島根県", "岡山県", "広島県", "山口県", "徳島県", "香川県", "愛媛県", "高知県", "福岡県", "佐賀県", "長崎県", "熊本県", "大分県", "宮崎県", "鹿児島", "沖縄県"}
    Private Sub 所要日数一覧ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 所要日数一覧ToolStripMenuItem.Click
        SplitContainer2.Panel2Collapsed = False
        NissuIchiran()
    End Sub

    Private Sub NissuIchiran()
        DGV2.Rows.Clear()
        For i As Integer = 0 To prefecture.Length - 1
            DGV2.Rows.Add(prefecture(i))
        Next

        Dim dH2 As ArrayList = TM_HEADER_GET(DGV2)
        For r As Integer = 0 To DGV1.RowCount - 1
            If IsNumeric(DGV1.Item(dH1.IndexOf("所要日数"), r).Value) Then
                Dim nissu As Integer = DGV1.Item(dH1.IndexOf("所要日数"), r).Value
                Dim haisou As String = DGV1.Item(dH1.IndexOf("配送会社"), r).Value

                Dim kenmei As String = DGV1.Item(dH1.IndexOf("送り先住所"), r).Value.ToString.Substring(0, 3)
                If Not prefecture.Contains(kenmei) Then
                    Continue For
                End If
                Dim nowRow As Integer = Array.IndexOf(prefecture, kenmei)

                Dim minCount As Integer = 0
                Dim maxCount As Integer = 0
                If haisou <> "" Then
                    Dim headerMin As String = ""
                    Dim headerMax As String = ""
                    Dim headerKosu As String = ""
                    Select Case True
                        Case Regex.IsMatch(haisou, "佐川")
                            headerMin = "佐川MIN"
                            headerMax = "佐川MAX"
                            headerKosu = "佐川個数"
                        Case Regex.IsMatch(haisou, "ゆうパケット")
                            headerMin = "ゆパケMIN"
                            headerMax = "ゆパケMAX"
                            headerKosu = "ゆパケ個数"
                        Case Else
                            Continue For
                    End Select

                    If DGV2.Item(dH2.IndexOf(headerMin), nowRow).Value Is Nothing Then
                        DGV2.Item(dH2.IndexOf(headerMin), nowRow).Value = nissu
                    Else
                        Dim a = DGV2.Item(dH2.IndexOf(headerMin), nowRow).Value
                        If CInt(DGV2.Item(dH2.IndexOf(headerMin), nowRow).Value) > nissu Then
                            DGV2.Item(dH2.IndexOf(headerMin), nowRow).Value = nissu
                        End If
                    End If
                    If DGV2.Item(dH2.IndexOf(headerMax), nowRow).Value Is Nothing Then
                        DGV2.Item(dH2.IndexOf(headerMax), nowRow).Value = nissu
                    Else
                        Dim a = DGV2.Item(dH2.IndexOf(headerMax), nowRow).Value
                        If CInt(DGV2.Item(dH2.IndexOf(headerMax), nowRow).Value) < nissu Then
                            DGV2.Item(dH2.IndexOf(headerMax), nowRow).Value = nissu
                        End If
                    End If
                    If DGV2.Item(dH2.IndexOf(headerKosu), nowRow).Value Is Nothing Then
                        DGV2.Item(dH2.IndexOf(headerKosu), nowRow).Value = "1=" & nissu
                    Else
                        Dim c As String() = Split(DGV2.Item(dH2.IndexOf(headerKosu), nowRow).Value, "=")
                        Dim c1 As Integer = CInt(c(0)) + 1
                        Dim c2 As String = ""
                        If c.Length > 1 Then
                            c2 = c(1) & "," & nissu
                        Else
                            c2 = nissu
                        End If
                        DGV2.Item(dH2.IndexOf(headerKosu), nowRow).Value = c1 & "=" & c2
                    End If
                End If
            End If
        Next

        Dim nameArray As String() = New String() {"佐川個数|佐川平均", "ゆパケ個数|ゆパケ平均"}
        For Each n As String In nameArray
            Dim nA As String() = Split(n, "|")

            For r As Integer = 0 To DGV2.RowCount - 1
                If Not DGV2.Item(dH2.IndexOf(nA(0)), r).Value Is Nothing Then
                    Dim c As String() = Split(DGV2.Item(dH2.IndexOf(nA(0)), r).Value, "=")
                    If c.Length > 1 Then
                        Dim c1 As String() = Split(c(1), ",")
                        Dim c2 As Integer() = Array.ConvertAll(Of String, Integer)(c1, AddressOf Integer.Parse)
                        DGV2.Item(dH2.IndexOf(nA(1)), r).Value = Math.Round(c2.Average, 1)
                    End If
                End If
            Next
        Next
    End Sub

    Private Sub 閉じるToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 閉じるToolStripMenuItem.Click
        SplitContainer2.Panel2Collapsed = True
    End Sub

    Private Sub 保存ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 保存ToolStripMenuItem.Click
        If DGV2.RowCount = 0 Then
            Exit Sub
        End If

        Dim savePath As String = ""
        '.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
        '.RestoreDirectory = True,
        Dim sfd As New SaveFileDialog With {
            .InitialDirectory = dPath,
            .Filter = "CSVファイル(*.csv)|*.csv|すべてのファイル(*.*)|*.*",
            .AutoUpgradeEnabled = False,
            .FilterIndex = 0,
            .Title = "保存先のファイルを選択してください",
            .OverwritePrompt = True,
            .CheckPathExists = True
        }

        sfd.FileName = "所要日数.csv"

        'ダイアログを表示する
        If sfd.ShowDialog(Me) = DialogResult.OK Then
            savePath = sfd.FileName
        Else
            Exit Sub
        End If

        SaveCsv(DGV2, savePath, 1)
    End Sub


    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        DGV_hibetu.BringToFront()
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        DGV2.BringToFront()
    End Sub

    Private Sub DGV_hibetu_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV_hibetu.CellClick
        Dim selDate As String = DGV_hibetu.Item(dH_hibetu.IndexOf("日付"), e.RowIndex).Value
        For i As Integer = 0 To ComboBox3.Items.Count - 1
            If InStr(ComboBox3.Items(i), selDate) > 0 Then
                ComboBox3.SelectedIndex = i
                Exit For
            End If
        Next
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        CB3save()
        MsgBox("保存しました")
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Dim cbSelStr As String = ComboBox3.SelectedItem
        If InStr(cbSelStr, " 完 ") > 0 Then
            MsgBox("完 追加済み")
        Else
            Dim cbNewStr = Replace(cbSelStr, "(全", " 完 (全")
            File.Move(dPath & "\" & cbSelStr & ".csv", dPath & "\" & cbNewStr & ".csv")
            MsgBox("「" & cbNewStr & "」修正しました")
        End If
        CB3get()
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        Dim cbSelStr As String = ComboBox3.SelectedItem
        If InStr(cbSelStr, " 完 ") > 0 Then
            Dim cbNewStr = Replace(cbSelStr, " 完 ", "")
            File.Move(dPath & "\" & cbSelStr & ".csv", dPath & "\" & cbNewStr & ".csv")
            MsgBox("「" & cbNewStr & "」修正しました")
        Else
            MsgBox("完 付いていません")
        End If
        CB3get()
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        System.Diagnostics.Process.Start(dPath)
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        System.Diagnostics.Process.Start(dPath)
    End Sub

    'エクスポート
    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        If DGV1.Rows.Count = 0 Then
            MsgBox("データがないです", MsgBoxStyle.SystemModal)
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
            sfd.FileName = "荷物追跡結果.csv"
        End If

        'ダイアログを表示する
        If sfd.ShowDialog(Me) = DialogResult.OK Then
            SaveCsv(sfd.FileName, 0, DGV1, True)
            MsgBox(sfd.FileName & vbCrLf & "保存しました", MsgBoxStyle.SystemModal)
        End If

    End Sub

    ' DataGridViewをCSV出力する
    Public Sub SaveCsv(ByVal fp As String, ByVal mode As Integer, dgv As DataGridView, header_flag As Boolean)

        dgv.EndEdit()


        Dim CsvDelimiter As String = Splitter
        If mode = 2 Then
            CsvDelimiter = vbTab
        End If

        Dim dt_header As String = ""
        If header_flag Then
            Dim dH As ArrayList = TM_HEADER_GET(dgv)
            For i As Integer = 0 To dH.Count - 1
                dt_header = dt_header & dH(i) & CsvDelimiter
            Next
        End If

        ' CSVファイルオープン
        Dim str As String = ""
        'Dim sw As StreamWriter = New StreamWriter(fp, False, EncodingName2(ToolStripDropDownButton10.Text))
        Dim sw As StreamWriter = New StreamWriter(fp, False, System.Text.Encoding.GetEncoding("Shift-JIS"))
        If header_flag Then
            sw.Write(dt_header)
            sw.Write(vbLf)
        End If
        For r As Integer = 0 To dgv.Rows.Count - 1
            For c As Integer = 0 To dgv.Columns.Count - 1
                ' DataGridViewのセルのデータ取得
                Dim dt As String = ""
                If dgv.Rows(r).Cells(c).Value Is Nothing = False Then
                    dt = dgv.Rows(r).Cells(c).Value.ToString()
                    dt = Replace(dt, vbCrLf, vbLf)
                    'dt = Replace(dt, vbLf, "")
                    If mode = 1 Then
                        dt = """" & dt & """"
                    Else
                        If Not dt Is Nothing Then
                            dt = Replace(dt, """", """""")
                            Select Case True
                                Case InStr(dt, Splitter), InStr(dt, vbLf), InStr(dt, vbCr), InStr(dt, """") 'Regex.IsMatch(dt, ".*,.*|.*" & vbLf & ".*|.*" & vbCr & ".*|.*"".*")
                                    dt = """" & dt & """"
                            End Select
                        End If
                    End If
                End If
                If c < dgv.Columns.Count - 1 Then
                    dt = dt & CsvDelimiter
                End If
                ' CSVファイル書込
                sw.Write(dt)
            Next
            sw.Write(vbLf)
            Application.DoEvents()
        Next

        ' CSVファイルクローズ
        sw.Close()
        Me.Text = fp
    End Sub
End Class