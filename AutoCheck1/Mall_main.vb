Imports System.Data.OleDb
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic.FileIO
Imports Sgry.Azuki.WinForms
Imports WinSCP
Imports ZetaHtmlTidy
Imports System.Linq
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel

Public Class Mall_main
    Public Shared CnAccdb As String = ""
    Public Shared CnAccdb_genre As String = ""
    Public Shared CnAccdb_amazon As String = ""

    Private Cn0 As New OleDbConnection
    Private SQLCm0 As OleDbCommand
    Private Adapter0 As New OleDbDataAdapter
    Private Table0 As New DataTable
    Private ds0 As New DataSet

    Private TableName As String = ""
    Private setField As String = ""
    Private setStr As String = ""
    Private where As String = ""

    Private dgv5table As New DataTable
    Private dgv5table_p As New DataTable

    Private dgv22table As New DataTable

    Private dbimg_yahooa As String = Path.GetDirectoryName(Form1.appPath) & "\dbimg\yahooa"

    Private TenpoTable As New DataTable '店舗データを確保する


    Private Sub Mall_main_Load(sender As Object, e As EventArgs) Handles Me.Load
        If Regex.IsMatch(My.Computer.Name, "ABCD|takashi|NAKA|ping|LIU", RegexOptions.IgnoreCase) Or Regex.IsMatch(appPath, "debug", RegexOptions.IgnoreCase) Then
            CheckBox17.Enabled = True
            CheckBox17.Checked = False
            CheckBox18.Enabled = True
            CheckBox18.Checked = True
            CheckBox19.Enabled = True
            CheckBox19.Checked = True
        Else
            Button1.Enabled = False
            Button2.Enabled = False
            DGV1.ReadOnly = True
            SplitContainer8.Enabled = False
            CheckBox17.Enabled = False
            CheckBox17.Checked = True
            CheckBox18.Enabled = False
            CheckBox18.Checked = False
            CheckBox19.Enabled = False
            CheckBox19.Checked = True
        End If

        CnAccdb = Path.GetDirectoryName(Form1.appPath) & "\db\miseDB.accdb"
        CnAccdb_genre = Path.GetDirectoryName(Form1.appPath) & "\db\category.accdb"
        CnAccdb_amazon = Path.GetDirectoryName(Form1.appPath) & "\db\AmazonNode.accdb"

        'フォルダ作成
        If Not Directory.Exists(dbimg_yahooa) Then
            Directory.CreateDirectory(dbimg_yahooa)
        End If
        TextBox32.Text = dbimg_yahooa

        '初期選択状態変更
        Dim cb As ComboBox() = New ComboBox() {ComboBox1, ComboBox3, ComboBox5, ComboBox6, ComboBox8, ComboBox9, ComboBox15, ComboBox16}
        For i As Integer = 0 To cb.Length - 1
            If cb(i).Items.Count > 0 Then
                cb(i).SelectedIndex = 0
            End If
        Next
        cb = New ComboBox() {ComboBox23, ComboBox24, ComboBox25}
        For i As Integer = 0 To cb.Length - 1
            If cb(i).Items.Count > 0 Then
                cb(i).SelectedIndex = 1
            End If
        Next
        Dim tscb As ToolStripComboBox() = New ToolStripComboBox() {ToolStripComboBox1, ToolStripComboBox3, ToolStripComboBox4}
        For i As Integer = 0 To tscb.Length - 1
            If tscb(i).Items.Count > 0 Then
                tscb(i).SelectedIndex = 0
            End If
        Next

        Dim cb1_num As Integer = 0

        TableName = "T_基本情報"
        '=================================================
        Dim Table As DataTable = MALLDB_CONNECT_SELECT(CnAccdb, TableName, "表示順")
        '=================================================
        For i As Integer = 0 To Table.Rows.Count - 1
            ComboBox1.Items.Add(Table.Rows(i)("店名").ToString)
            ComboBox7.Items.Add(Table.Rows(i)("店名").ToString)
            ToolStripComboBox2.Items.Add(Table.Rows(i)("店名").ToString)
            DGV9.Rows.Add(Table.Rows(i)("店名"), "", "", "", Table.Rows(i)("モール"), Table.Rows(i)("店ID"))
            If Table.Rows(i)("初期選択").ToString = "yes" Then
                cb1_num = i
            End If
        Next
        TenpoTable = Table.Copy
        '----------
        Table.Clear()
        '----------

        ComboBox1_items1()                      'ComboBox1（店舗名）
        ComboBox2_items1()                      'ComboBox2（csv店舗名）
        ComboBox1.SelectedIndex = cb1_num
        ComboBox7.SelectedIndex = 0
        ToolStripComboBox2.SelectedIndex = 0
        TableGetLastTime(0)
        ComboBox18.SelectedIndex = 0

        If Form1.AdminFlag Then
            Button59.Visible = True
        End If

        TableName = "T_設定_Amazon"
        '=================================================
        Table0 = MALLDB_CONNECT_SELECT(CnAccdb, TableName, "名称='テンプレート名'")
        '=================================================
        Dim hs1 As String() = Split(Table0.Rows(0)("設定1").ToString, "|=|")
        Dim hs2 As String() = Split(Table0.Rows(0)("設定2").ToString, "|=|")
        ComboBox13.Items.Clear()
        ComboBox14.Items.Clear()
        DGV21.Rows.Clear()
        For i As Integer = 0 To hs1.Length - 1
            ComboBox13.Items.Add(hs1(i) & "," & hs2(i))
            ComboBox14.Items.Add(hs1(i) & "," & hs2(i))
            DGV21.Rows.Add(hs1(i), 0)
        Next

        '******************************************
        'dgv画面チラつき防止
        '******************************************
        'DataGirdViewのTypeを取得
        Dim dgvtype As Type = GetType(DataGridView)
        'プロパティ設定の取得
        Dim dgvPropertyInfo As Reflection.PropertyInfo = dgvtype.GetProperty("DoubleBuffered", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)

        '対象のDataGridViewにtrueをセットする
        Dim dgvArray As DataGridView() = {DGV1, DGV2, DGV3, DGV4, DGV5, DGV6, DGV7, DGV8, DGV9, DGV10, DGV11, DGV12, DGV13, DGV14, DGV15, DGV16, DGV17, DGV18, DGV19, DGV20, DGV21}
        For i As Integer = 0 To dgvArray.Length - 1
            dgvPropertyInfo.SetValue(dgvArray(i), True, Nothing)
        Next

        FocusDGV = DGV2
    End Sub

    Private Sub Form1_Shown(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Shown
        Label19.Visible = True
        Application.DoEvents()
        MasterDB()
        HelpDB()
        DGV20_Header()
        DGV22DB()

        Label19.Visible = False
    End Sub

    '行番号を表示する
    Private Sub DataGridView_RowPostPaint(ByVal sender As Object, ByVal e As Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles _
            DGV1.RowPostPaint, DGV2.RowPostPaint, DGV3.RowPostPaint, DGV4.RowPostPaint, DGV5.RowPostPaint,
            DGV6.RowPostPaint, DGV7.RowPostPaint, DGV8.RowPostPaint, DGV9.RowPostPaint, DGV10.RowPostPaint,
            DGV11.RowPostPaint, DGV12.RowPostPaint, DGV13.RowPostPaint, DGV14.RowPostPaint, DGV15.RowPostPaint,
            DGV16.RowPostPaint, DGV17.RowPostPaint, DGV19.RowPostPaint, DGV20.RowPostPaint, DGV20.RowPostPaint
        Dim rect As New Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, sender.RowHeadersWidth - 4, sender.Rows(e.RowIndex).Height)
        TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), sender.RowHeadersDefaultCellStyle.Font, rect, sender.RowHeadersDefaultCellStyle.ForeColor, TextFormatFlags.VerticalCenter Or TextFormatFlags.Right)
    End Sub

    Private Sub DataGridView16_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DGV16.DataError
        If Not (e.Exception Is Nothing) Then
            ListBox9.Items.Add(e.ColumnIndex & ":" & e.RowIndex + 1 & "=>「" & DGV16.Item(e.ColumnIndex, e.RowIndex).Value & "」は入力不可")
            DGV16.Item(e.ColumnIndex, e.RowIndex).Value = ""
            DGV16.Item(e.ColumnIndex, e.RowIndex).Style.BackColor = Color.Orange

            'MessageBox.Show(Me,
            'String.Format("({0}, {1}) のセルでエラーが発生しました。" +
            '    vbCrLf + vbCrLf + "説明: {2}" & vbCrLf & "「" & dgv16.Item(e.ColumnIndex, e.RowIndex).Value & "」は入力できません。",
            '    e.ColumnIndex, e.RowIndex, e.Exception.Message),
            '"エラーが発生しました",
            'MessageBoxButtons.OK,
            'MessageBoxIcon.Error)

            'If InStr(RichTextBox3.Text, "ERROR/") = 0 Then
            '    RichTextBox3.Text = ""
            'End If
            'RichTextBox3.Text &= "ERROR/" & e.ColumnIndex & ":" & e.RowIndex & "「" & DGV16.Item(e.ColumnIndex, e.RowIndex).Value & "」は入力不可" & vbCrLf
        End If
    End Sub

    Private Sub ListBox9_DoubleClick(sender As Object, e As EventArgs) Handles ListBox9.DoubleClick
        Dim selStr As String = Split(ListBox9.SelectedItem, "=>")(0)
        Dim selNum As String() = Split(selStr, ":")
        Try
            DGV16.CurrentCell = DGV16(CInt(selNum(0)), CInt(selNum(1)) - 1)
        Catch ex As Exception

        End Try
    End Sub

    Public FocusDGV As DataGridView = DGV2
    Private Sub DataGridView_GotFocus(sender As DataGridView, e As EventArgs) Handles _
            DGV1.GotFocus, DGV2.GotFocus, DGV3.GotFocus, DGV4.GotFocus, DGV5.GotFocus,
            DGV6.GotFocus, DGV7.GotFocus, DGV8.GotFocus, DGV9.GotFocus, DGV10.GotFocus,
            DGV11.GotFocus, DGV12.GotFocus, DGV13.GotFocus, DGV14.GotFocus, DGV15.GotFocus,
            DGV16.GotFocus, DGV17.GotFocus, DGV18.GotFocus, DGV19.GotFocus, DGV20.GotFocus
        If sender Is Nothing Then
            FocusDGV = DGV2
        Else
            FocusDGV = sender
        End If
    End Sub

    Private Sub DataGridView_DragDrop(ByVal sender As Object, ByVal e As Windows.Forms.DragEventArgs) Handles DGV9.DragDrop
        Me.Activate()
        For Each filename As String In e.Data.GetData(DataFormats.FileDrop)
            LIST4VIEW("ドラッグドロップ認識", "s")
            CsvGet(filename)
        Next
    End Sub

    Private Sub DataGridView12_DragDrop(ByVal sender As Object, ByVal e As Windows.Forms.DragEventArgs) Handles _
            DGV12.DragDrop, DGV13.DragDrop, DGV17.DragDrop
        Me.Activate()
        For Each filename As String In e.Data.GetData(DataFormats.FileDrop)
            LIST4VIEW("ドラッグドロップ認識", "s")
            AmazonTemplateXlsGet(filename)
        Next
    End Sub

    Private Sub DataGridView14_DragDrop(ByVal sender As Object, ByVal e As Windows.Forms.DragEventArgs) Handles DGV14.DragDrop
        Me.Activate()
        For Each filename As String In e.Data.GetData(DataFormats.FileDrop)
            LIST4VIEW("ドラッグドロップ認識", "s")
            AmazonTemplateXlsGet(filename)
        Next
    End Sub

    Private Sub DataGridView16_DragDrop(ByVal sender As Object, ByVal e As Windows.Forms.DragEventArgs) Handles _
            DGV15.DragDrop, DGV16.DragDrop
        Me.Activate()
        For Each filename As String In e.Data.GetData(DataFormats.FileDrop)
            LIST4VIEW("ドラッグドロップ認識", "s")
            AmazonItemXlsGet(filename)
        Next
    End Sub

    Private Sub DataGridView20_DragDrop(ByVal sender As Object, ByVal e As Windows.Forms.DragEventArgs) Handles _
            DGV20.DragDrop
        Me.Activate()
        For Each filename As String In e.Data.GetData(DataFormats.FileDrop)
            LIST4VIEW("ドラッグドロップ認識", "s")
            AmazonCsvGet(filename)
        Next
    End Sub

    Private Sub DataGridView11_DragDrop(ByVal sender As Object, ByVal e As Windows.Forms.DragEventArgs) Handles _
            DGV11.DragDrop
        Me.Activate()
        For Each filename As String In e.Data.GetData(DataFormats.FileDrop)
            LIST4VIEW("ドラッグドロップ認識", "s")
            CategoryCSVGet(filename)
        Next
    End Sub

    Private Sub DataGridView_DragEnter(ByVal sender As Object, ByVal e As Windows.Forms.DragEventArgs) Handles _
            DGV9.DragEnter, DGV12.DragEnter, DGV13.DragEnter, DGV17.DragEnter, DGV14.DragEnter,
            DGV15.DragEnter, DGV16.DragEnter, DGV20.DragEnter, DGV11.DragEnter, DGV22.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub Mall_main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If e.CloseReason = CloseReason.FormOwnerClosing Or e.CloseReason = CloseReason.UserClosing Then
            Me.Visible = False
            e.Cancel = True
        End If
    End Sub

    Private Sub LIST4VIEW(ByRef mes As String, ByRef se As String)
        If ListBox4.Items.Count > 100 Then
            ListBox4.Items.Clear()
        End If

        Dim seStr As String = ""
        Select Case se
            Case "start", "s"
                seStr = "START"
                ListBox4.Items.Add(mes & ">>" & seStr)
            Case "end", "e"
                seStr = "END"
                ListBox4.Items.Add("..." & seStr & ":" & mes)
            Case "error", "r"
                seStr = "ERROR"
                ListBox4.Items.Add("[" & seStr & "]:" & mes)
            Case Else
                seStr = se
                ListBox4.Items.Add("  =>" & mes & "/" & seStr)
        End Select

        ListBox4.SelectedIndex = ListBox4.Items.Count - 1
        ListBox4.SelectedItems.Clear()
        Application.DoEvents()
    End Sub

    Private Sub TextBox_KeyDown(sender As TextBox, e As KeyEventArgs) Handles _
        TextBox1.KeyDown, TextBox2.KeyDown, TextBox3.KeyDown, TextBox9.KeyDown, TextBox15.KeyDown,
        TextBox16.KeyDown
        If e.Modifiers = Keys.Control And e.KeyCode = Keys.A Then
            sender.SelectAll()
        End If
    End Sub

    Private Sub AzukiControl_SizeChanged(sender As AzukiControl, e As EventArgs) Handles _
            AzukiControl1.SizeChanged, AzukiControl2.SizeChanged, AzukiControl4.SizeChanged
        sender.ViewWidth = sender.Width - 30
    End Sub

    Private Sub フォームのリセットToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles フォームのリセットToolStripMenuItem.Click
        Me.Dispose()
        Dim frm As Form = New Mall_main
        frm.Show()
    End Sub

    Private Sub ToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem3.Click
        If ToolStripMenuItem3.Text = ">>>" Then
            ToolStripMenuItem3.Text = "<<<"
            SplitContainer1.Panel1Collapsed = False
        Else
            ToolStripMenuItem3.Text = ">>>"
            SplitContainer1.Panel1Collapsed = True
        End If
    End Sub



    '====================================================================================================================
    '====================================================================================================================
    '====================================================================================================================


    'ComboBox1_SelectedIndexChanged
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedIndex = 0 Then
            ComboBox2.Enabled = True
            ComboBox2_items1()
        Else
            ComboBox2.Enabled = False

            '先に削除してから、初期選択状態を保存
            LIST4VIEW("初期選択状態読込", "s")
            setStr = "[初期選択] = ''"
            where = "[初期選択] = 'yes'"
            TableName = "T_基本情報"
            '=================================================
            MALLDB_CONNECT_UPDATE(CnAccdb, TableName, setStr, where)
            '=================================================
            setStr = "[初期選択] = 'yes'"
            where = "[店名] = '" & ComboBox1.SelectedItem & "'"
            TableName = "T_基本情報"
            '=================================================
            MALLDB_CONNECT_UPDATE(CnAccdb, TableName, setStr, where)
            '=================================================

            '設定を読み込み
            LIST4VIEW("店舗設定読込", "s")
            where = " [店名] = '" & ComboBox1.SelectedItem & "'"
            TableName = "T_基本情報"
            '=================================================
            Dim Table1 As DataTable = MALLDB_CONNECT_SELECT(CnAccdb, TableName, where, "表示順")
            '=================================================
            where = " [モール] = '" & Table1.Rows(0)("モール").ToString & "'"
            TableName = "T_設定_基本情報"
            '=================================================
            Dim Table2 As DataTable = MALLDB_CONNECT_SELECT(CnAccdb, TableName, where)
            '=================================================
            DGV1.Rows.Clear()
            If Table2.Rows.Count > 0 Then
                For c As Integer = 0 To Table2.Columns.Count - 1
                    Dim cName As String = Table2.Columns(c).ColumnName
                    If c = 0 Then
                        DGV1.Rows.Add("モール", Table1.Rows(0)("モール").ToString)
                        DGV1.Item(1, DGV1.Rows.Count - 1).ReadOnly = True
                        DGV1.Item(1, DGV1.Rows.Count - 1).Style.BackColor = Color.LightCyan
                    ElseIf Table2.Rows(0)(cName).ToString = "yes" Then
                        DGV1.Rows.Add(cName, Table1.Rows(0)(cName).ToString)
                        If Regex.IsMatch(cName, "店名|モール") Then
                            DGV1.Item(1, DGV1.Rows.Count - 1).ReadOnly = True
                            DGV1.Item(1, DGV1.Rows.Count - 1).Style.BackColor = Color.LightCyan
                        ElseIf Regex.IsMatch(cName, "ログイン2|パス2|FTPPASS") Then
                            DGV1.Item(1, DGV1.Rows.Count - 1).ReadOnly = False
                            'DataGridView1.Item(1, DataGridView1.Rows.Count - 1).Value = "********"
                        End If
                    End If
                    DGV1.Item(0, DGV1.Rows.Count - 1).ReadOnly = True
                    DGV1.Item(0, DGV1.Rows.Count - 1).Style.BackColor = Color.LightCyan
                Next
            End If
            '----------
            Table2.Clear()
            Table1.Clear()
            '----------
        End If
    End Sub

    '店舗一覧読み込み
    Private Sub ComboBox1_items1()
        LIST4VIEW("店舗一覧読込", "s")
        ComboBox1.Items.Clear()
        ComboBox1.Items.Add("新規店舗登録")
        ComboBox7.Items.Clear()
        ToolStripComboBox2.Items.Clear()
        DGV9.Rows.Clear()

        Dim cb1_num As Integer = 0
        TableName = "T_基本情報"
        '=================================================
        Dim Table As DataTable = MALLDB_CONNECT_SELECT(CnAccdb, TableName, "表示順")
        '=================================================
        For i As Integer = 0 To Table.Rows.Count - 1
            If Table.Rows(i)("店名") <> "マスタ" Then
                ComboBox1.Items.Add(Table.Rows(i)("店名").ToString)
            End If
            ComboBox7.Items.Add(Table.Rows(i)("店名").ToString)
            ToolStripComboBox2.Items.Add(Table.Rows(i)("店名").ToString)
            DGV9.Rows.Add(Table.Rows(i)("店名"), "", "", "", Table.Rows(i)("モール"), Table.Rows(i)("店ID"))
            If Table.Rows(i)("初期選択").ToString = "yes" Then
                cb1_num = i + 1
            End If
        Next
        '----------
        Table.Clear()
        '----------

        TableGetLastTime(0)
    End Sub

    'テンプレート選択用モール一覧読み込み
    Private Sub ComboBox2_items1()
        LIST4VIEW("テンプレート選択用一覧読込", "s")
        TableName = "T_設定_基本情報"
        '=================================================
        Dim Table As DataTable = MALLDB_CONNECT_SELECT(CnAccdb, TableName)
        '=================================================
        ComboBox2.Items.Clear()
        DGV1.Rows.Clear()
        If Table.Rows.Count > 0 Then
            For i As Integer = 0 To Table.Rows.Count - 1
                ComboBox2.Items.Add(Table.Rows(i)("モール").ToString)
            Next
        End If
        ComboBox2.SelectedIndex = 0
        '----------
        Table.Clear()
        '----------
    End Sub

    'モール基本情報テンプレート読み込み
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        LIST4VIEW("モール基本情報テンプレート読込", "s")
        where = "[モール] = '" & ComboBox2.SelectedItem & "'"
        TableName = "T_設定_基本情報"
        '=================================================
        Dim Table As DataTable = MALLDB_CONNECT_SELECT(CnAccdb, TableName, where)
        '=================================================
        DGV1.Rows.Clear()
        If Table.Rows.Count > 0 Then
            For c As Integer = 0 To Table.Columns.Count - 1
                If Table.Rows(0)(c).ToString = "yes" Then
                    DGV1.Rows.Add(Table.Columns(c).ColumnName, "")
                End If
            Next
        End If
        '----------
        Table.Clear()
        '----------
    End Sub

    '店舗の登録・編集
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        LIST4VIEW("店舗の登録・編集", "s")
        Dim checkName As String() = New String() {"店名", "TB名", "ログイン1", "パス1", "店ID"}
        For r As Integer = 0 To DGV1.Rows.Count - 1
            If checkName.Contains(DGV1.Item(0, r).Value) Then
                If DGV1.Item(1, r).Value = "" Then
                    DGV1.Item(1, r).Style.BackColor = Color.Yellow
                    MsgBox("入力必須項目が未入力です", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
                    Exit Sub
                Else
                    DGV1.Item(1, r).Style.BackColor = Color.LightCyan
                End If
            End If
        Next

        Application.DoEvents()

        where = "[店名] = '"
        For r As Integer = 0 To DGV1.Rows.Count - 1
            If DGV1.Item(0, r).Value = "店名" Then
                where &= DGV1.Item(1, r).Value
                Exit For
            End If
        Next
        where &= "'"
        TableName = "T_基本情報"
        '=================================================
        Dim Table As DataTable = MALLDB_CONNECT_SELECT(CnAccdb, TableName, where, "表示順")
        '=================================================
        LIST4VIEW("入力情報チェック", "e")

        If Table.Rows.Count = 0 Then
            setField = ""
            setStr = ""
            For r As Integer = 0 To DGV1.Rows.Count - 1
                If r = 0 And DGV1.Item(0, r).Value <> "モール" Then
                    setField &= "`モール`,"
                    setStr &= "'" & ComboBox2.SelectedItem & "',"
                End If

                setField &= "`" & DGV1.Item(0, r).Value & "`,"
                setStr &= "'" & DGV1.Item(1, r).Value & "',"
            Next
            setField = setField.TrimEnd(",")
            setStr = setStr.TrimEnd(",")
            TableName = "T_基本情報"
            '=================================================
            MALLDB_CONNECT_INSERT_ONE(CnAccdb, TableName, setField, setStr)
            '=================================================

            MsgBox("登録しました", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
        Else
            setStr = "[初期選択] = 'yes',"
            For r As Integer = 0 To DGV1.Rows.Count - 1
                setStr &= "[" & DGV1.Item(0, r).Value & "] = '" & DGV1.Item(1, r).Value & "',"
            Next
            setStr = setStr.TrimEnd(",")
            where = "[店名] = '" & ComboBox1.SelectedItem & "'"
            TableName = "T_基本情報"
            '=================================================
            MALLDB_CONNECT_UPDATE(CnAccdb, TableName, setStr, where)
            '=================================================

            MsgBox("更新しました", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
        End If

        '----------
        Table.Clear()
        '----------

        ComboBox1_items1()
        ComboBox1.SelectedIndex = ComboBox1.Items.Count - 1
    End Sub

    '店舗の削除
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If MsgBox(ComboBox1.SelectedItem & "を削除しますか？", MsgBoxStyle.OkCancel Or MsgBoxStyle.SystemModal) = MsgBoxResult.Cancel Then
            Exit Sub
        End If

        LIST4VIEW("店舗データ削除", "s")
        where = "[店名] = '" & ComboBox1.SelectedItem & "'"
        TableName = "T_基本情報"
        '=================================================
        MALLDB_CONNECT_DELETE(CnAccdb, TableName, where)
        '=================================================

        ComboBox1_items1()
        ComboBox1.SelectedIndex = 0
        MsgBox(ComboBox1.SelectedItem & "を削除しました", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
    End Sub

    '更新情報アップデート
    Private Sub LinkLabel1_LinkClicked_1(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        TableGetLastTime(0)
    End Sub

    Dim tt As Integer = 0
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If tt = 60 * 5 Then
            TableGetLastTime(0)
        End If
        tt += 1
    End Sub

    'サーバーデータ同期
    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        Dim serverPath As String = Form1.サーバーToolStripMenuItem.Text
        If Regex.IsMatch(appPath, "takashi", RegexOptions.IgnoreCase) And Regex.IsMatch(appPath, "debug", RegexOptions.IgnoreCase) Then
            serverPath = SpecialDirectories.MyDocuments & "\test"
        End If

        If CheckBox17.Checked Then
            Dim fPath2 As String = serverPath & "\tenpo\settei_lasttime.dat"
            Dim strArray1 As String = File.ReadAllText(fPath2)
            Dim strArray As String() = Split(strArray1, vbCrLf)
            MALLDB_CONNECT(CnAccdb)
            Dim tableName As String = ""
            Dim num As Integer = 0
            Dim FieldArray As New ArrayList

            ProgressBar1.Value = 0
            ProgressBar1.Maximum = strArray.Length

            For i As Integer = 0 To strArray.Length - 1
                If strArray(i) = "" Then
                    Continue For
                End If
                Select Case True
                    Case Regex.IsMatch(strArray(i), "^#####")
                        tableName = Replace(strArray(i), "#####", "")
                        FieldArray.Clear()
                        num = 0
                    Case Else
                        Dim sArray As String() = Split(strArray(i), "|$$|")
                        If num = 0 Then
                            For k As Integer = 1 To sArray.Length - 1
                                FieldArray.Add(sArray(k))
                            Next
                        Else
                            Dim SetArray As New ArrayList
                            For k As Integer = 1 To sArray.Length - 1
                                SetArray.Add(sArray(k))
                            Next
                            MALLDB_UPSERT(tableName, "[ID]=" & sArray(0) & "", FieldArray, SetArray)
                        End If
                        num += 1
                End Select

                ProgressBar1.Value += 1
            Next

            ProgressBar1.Value = 0
            MALLDB_DISCONNECT()
        End If

        Dim savePath As String = ""
        For r As Integer = 0 To DGV9.Rows.Count - 1
            For c As Integer = 1 To 3
                If DGV9.Item(c, r).Value.ToString = "" Then
                    Continue For
                ElseIf DGV9.Item(c, r).Style.BackColor = Color.Yellow Then
                    Dim fName As String = DGV9.Item(0, r).Value & "_" & DGV9.Columns(c).HeaderText & ".csv"
                    If r = 0 And c = 1 Then
                        fName = DGV9.Item(0, r).Value & "_" & "master" & ".csv"
                    ElseIf r = 0 And c = 2 Then
                        fName = DGV9.Item(0, r).Value & "_" & "page" & ".csv"
                    End If
                    savePath = serverPath & "\tenpo\" & fName
                    LIST4VIEW(DGV9.Item(0, r).Value & " " & Path.GetFileNameWithoutExtension(fName) & " csv取得", "s")
                    'MsgBox(dgv9.Item(0, r).Value & "_" & dgv9.Columns(c).HeaderText & vbCrLf & fName & vbCrLf & c & vbCrLf & r)
                    CsvGet(savePath, "1", DGV9.Item(0, r).Value)
                End If
            Next
        Next

        '管理用設定アップロード
        If CheckBox18.Checked Then
            Settei_write()
        End If

        'Amazonデータコピー
        If CheckBox19.Checked Then
            AmazonDBcopy()
        End If

        MsgBox("サーバー同期終了しました。フォームをリセットします。", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
        フォームのリセットToolStripMenuItem.PerformClick()
    End Sub

    Private Sub AmazonDBcopy()
        Dim serverPath As String = Form1.サーバーToolStripMenuItem.Text
        Dim CnAccdb_server As String = serverPath & "\update\db\AmazonNode.accdb"
        Dim CnAccdb_amazon_temp As String = Path.GetDirectoryName(Form1.appPath) & "\db\AmazonNodeTemp.accdb"
        File.Copy(CnAccdb_server, CnAccdb_amazon_temp, True)

        '更新情報を開いて新しいテーブルを調べる
        Dim tableS As DataTable = TM_DB_CONNECT_SELECT(CnAccdb_amazon_temp, "更新情報")
        Dim tableT As DataTable = TM_DB_CONNECT_SELECT(CnAccdb_amazon, "更新情報")

        ProgressBar1.Value = 0
        ProgressBar1.Maximum = tableS.Rows.Count

        Dim tableName As String = ""
        For i As Integer = 0 To tableS.Rows.Count - 1
            If tableS.Rows(i)("更新時間") < tableT.Rows(i)("更新時間") Then
                'ローカルからサーバー
                'MsgBox(tableS.Rows(i)("テーブル名") & "<" & tableT.Rows(i)("テーブル名"))
                tableName = tableS.Rows(i)("テーブル名")
                TM_DB_CONNECT_TABLE_COPY(CnAccdb_amazon, tableName, CnAccdb_amazon_temp, tableName)
            Else
                'サーバーからローカル
                'MsgBox(tableS.Rows(i)("テーブル名") & ">" & tableT.Rows(i)("テーブル名"))
                tableName = tableS.Rows(i)("テーブル名")
                TM_DB_CONNECT_TABLE_COPY(CnAccdb_amazon_temp, tableName, CnAccdb_amazon, tableName)
            End If
            ProgressBar1.Value += 1
        Next

        ProgressBar1.Value = 0
    End Sub

    'csvファイルの分別と保存
    Private Sub CsvGet(ByVal fName As String, Optional mode As String = "", Optional dgvTenpo As String = "")
        LIST4VIEW("店舗認識", "s")
        Dim HeaderMark As ArrayList = New ArrayList From {
            "商品分類タグ,マスタ"
        }
        TableName = "T_基本情報"
        '=================================================
        Table0 = MALLDB_CONNECT_SELECT(CnAccdb, TableName, "表示順")
        '=================================================
        For i As Integer = Table0.Rows.Count - 1 To 0 Step -1
            HeaderMark.Add(Table0.Rows(i)("店ID").ToString & "," & Table0.Rows(i)("店名").ToString)
        Next

        Dim itemList_web As ArrayList = TM_CSV_READ(fName)(0)
        If dgvTenpo = "" Then   'csvから店舗を特定する
            If Regex.IsMatch(itemList_web(0), "商品分類タグ|daihyo_syohin_code") Then
                ComboBox7.SelectedItem = "マスタ"
                ToolStripComboBox2.SelectedItem = "マスタ"
            ElseIf InStr(itemList_web(0), "管理番号|=|カテゴリ|=|タイトル") > 0 Then
                ComboBox7.SelectedItem = "ヤフオク雑貨ショップKT"
                ToolStripComboBox2.SelectedItem = "ヤフオク雑貨ショップKT"
            Else
                For i As Integer = Table0.Rows.Count - 1 To 0 Step -1   '1～2行目から店舗を特定
                    If InStr(itemList_web(1), "www.qoo10.jp") > 0 Then
                        ComboBox7.SelectedItem = "Qoo10雑貨倉庫"
                        ToolStripComboBox2.SelectedItem = "Qoo10雑貨倉庫"
                        Exit For
                    Else
                        Dim hm As String() = Split(HeaderMark(i), ",")
                        'MsgBox(itemList_web(1) & vbCrLf & Table0.Rows(i)("店ID").ToString)
                        If Regex.IsMatch(itemList_web(1), Table0.Rows(i)("店ID").ToString & "|" & Table0.Rows(i)("店名").ToString) Then
                            ComboBox7.SelectedItem = Table0.Rows(i)("店名").ToString
                            ToolStripComboBox2.SelectedItem = Table0.Rows(i)("店名").ToString
                            Exit For
                        End If
                    End If
                Next
            End If
        Else
            ComboBox7.SelectedItem = dgvTenpo
            ToolStripComboBox2.SelectedItem = dgvTenpo
        End If
        Application.DoEvents()

        '-----------------------------------------------------------------------------------------
        LIST4VIEW("テンプレート種類認識", "s")
        Dim KindMark As ArrayList = New ArrayList From {
            "原価\|=\|商品区分,master"
        }
        TableName = "T_設定_テンプレート"
        '=================================================
        Table0 = MALLDB_CONNECT_SELECT(CnAccdb, TableName)
        '=================================================
        For i As Integer = Table0.Rows.Count - 1 To 0 Step -1
            KindMark.Add(Table0.Rows(i)("検索").ToString & "," & Table0.Rows(i)("種類").ToString)
        Next

        Dim kmName As String = ""
        Dim pathEnd As String = ""
        For i As Integer = 0 To KindMark.Count - 1
            Dim km As String() = Split(KindMark(i), ",")
            If Regex.IsMatch(itemList_web(0), km(0)) Then
                pathEnd = Path.GetDirectoryName(Form1.appPath) & "\tenpo\" & ComboBox7.SelectedItem & "_" & km(1) & ".csv"
                kmName = km(1)
                Exit For
            End If
        Next

        '----------
        Table0.Clear()
        '----------

        If Path.GetFileNameWithoutExtension(pathEnd) = "" Then
            MsgBox("ファイルが認識できません", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        '-----------------------------------------------------------------------------------------
        Dim DR As DialogResult = Nothing
        If mode = "" Then
            DR = MsgBox(Path.GetFileNameWithoutExtension(pathEnd) & "でよろしいですか？", MsgBoxStyle.OkCancel Or MsgBoxStyle.SystemModal)
        End If

        If DR = Windows.Forms.DialogResult.Cancel Then
            Exit Sub
        Else
            'csvファイルを確認する
            Dim tobasuFlag As Boolean = False
            Dim res As String = CSVcheck(fName, ComboBox7.SelectedItem, kmName)
            If Regex.IsMatch(res, "^error:") Then
                Dim DRR As DialogResult = MsgBox("CSVデータが正しくありません。データ同期を飛ばしますか？" & vbCrLf & res, MsgBoxStyle.YesNo Or MsgBoxStyle.SystemModal)
                If DRR = DialogResult.Yes Then
                    tobasuFlag = True
                End If
            End If

            'データが正しく無ければ同期をしない選択肢。飛ばさなければエラーが出るのでエラー箇所を特定できる。
            If tobasuFlag = False Then
                'フォルダにファイルを保存
                File.Copy(fName, pathEnd, True)
                'データベースにファイルを登録
                CSV_TO_DB(fName, kmName, itemList_web)
                'ファイル更新時間をデータベースに保存
                FileGetLastTime(pathEnd)

                If ComboBox7.SelectedItem = "マスタ" Then
                    MasterDB()
                End If

                If mode = "" Then
                    MsgBox("データを登録しました", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
                End If
            End If

        End If

        TableGetLastTime(0)
    End Sub

    Private Sub DGV22_DragDrop(sender As Object, e As DragEventArgs) Handles DGV22.DragDrop
        Dim serverPath As String = Form1.サーバーToolStripMenuItem.Text
        Dim serverPath2 As String = serverPath & "\tenpo\ヤフオク雑貨ショップKT_auctions.csv"
        Dim pathEnd = Path.GetDirectoryName(Form1.appPath) & "\tenpo\ヤフオク雑貨ショップKT_auctions.csv"
        Dim TableDst As String = "T_KTYA_auctions"
        Dim fPath As String = serverPath & "\tenpo\tenpo_lasttime2.txt"
        Dim action = False

        For Each filename As String In e.Data.GetData(DataFormats.FileDrop)
            Dim itemList_web As ArrayList = TM_CSV_READ(filename)(0)
            If InStr(itemList_web(0), "管理番号|=|カテゴリ|=|タイトル|=|説明|=|ストア内商品検索用キーワード|=|開始価格|=|即決価格|=|値下げ交渉") Then

                File.Copy(filename, pathEnd, True)
                File.Copy(filename, serverPath2, True)

                MALLDB_CONNECT(CnAccdb)
                '=================================================
                MALLDB_CONNECT_DELETE(CnAccdb, TableDst, "")
                Dim csvlines As String() = Split(itemList_web(0), "|=|")
                setField = ""
                For i As Integer = 0 To csvlines.Length - 1
                    If csvlines(i) <> "" Then
                        If InStr(csvlines(i), "[") > 0 Then
                            setField &= csvlines(i) & ","
                        Else
                            setField &= "[" & csvlines(i) & "],"
                        End If
                    End If
                Next
                setField = setField.TrimEnd(",")

                For r As Integer = 1 To itemList_web.Count - 1
                    setStr = ""
                    itemList_web(r) = AccessStrConvert(itemList_web(r))     'Access登録用文字変換
                    Dim lines As String() = Split(itemList_web(r), "|=|")
                    For i As Integer = 0 To lines.Count - 1
                        setStr &= "'" & lines(i) & "',"
                    Next
                    setStr = setStr.TrimEnd(",")
                    MALLDB_SET_STR(TableDst, setField, setStr)
                    Application.DoEvents()
                Next

                Dim da As Date = File.GetLastWriteTime(filename)
                Dim where As String = "[ファイル名]='T_KTYA_auctions'"
                Dim fieldArray As New ArrayList From {
                        "ファイル名", "更新"
                    }
                Dim setArray As New ArrayList From {
                        "T_KTYA_auctions", da
                    }
                MALLDB_CONNECT_UPSERT(CnAccdb, "T_ファイル情報", where, fieldArray, setArray)

                'TextBox4.Text = da

                Dim file_lasttime As New System.IO.StreamWriter(fPath)
                file_lasttime.WriteLine(da)
                file_lasttime.Close()

                action = True
            Else
                MessageBox.Show("ファイルを確認してください")
            End If
        Next

        MALLDB_DISCONNECT()

        If action = True Then
            DGV22DB()
            MessageBox.Show("データを登録しました")
        End If
    End Sub


    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim serverPath As String = Form1.サーバーToolStripMenuItem.Text
        Dim fPath As String = serverPath & "\tenpo\tenpo_lasttime2.txt"
        Dim file As New System.IO.StreamReader(fPath)
        Dim serverfiletime As Date = file.ReadToEnd()
        Dim fName As String = "\\SERVER2-PC\Users\Public\program\autocheck\tenpo\ヤフオク雑貨ショップKT_auctions.csv"


        MALLDB_CONNECT(CnAccdb)
        Dim tableName As String = "T_KTYA_auctions"
        Dim itemList_web As ArrayList = TM_CSV_READ(fName)(0)
        Dim csvlines As String() = Split(itemList_web(0), "|=|")
        MALLDB_CONNECT_DELETE(CnAccdb, tableName, "")
        setField = ""
        For i As Integer = 0 To csvlines.Length - 1
            If csvlines(i) <> "" Then
                If InStr(csvlines(i), "[") > 0 Then
                    setField &= csvlines(i) & ","
                Else
                    setField &= "[" & csvlines(i) & "],"
                End If
            End If
        Next
        setField = setField.TrimEnd(",")
        For r As Integer = 1 To itemList_web.Count - 1
            setStr = ""
            itemList_web(r) = AccessStrConvert(itemList_web(r))     'Access登録用文字変換
            Dim lines As String() = Split(itemList_web(r), "|=|")
            For i As Integer = 0 To lines.Count - 1
                setStr &= "'" & lines(i) & "',"
            Next
            setStr = setStr.TrimEnd(",")
            MALLDB_SET_STR(tableName, setField, setStr)
            Application.DoEvents()
        Next

        Dim where As String = "[ファイル名]='T_KTYA_auctions'"
        Dim fieldArray As New ArrayList From {
                "ファイル名", "更新"
            }
        Dim setArray As New ArrayList From {
                "T_KTYA_auctions", serverfiletime
            }
        MALLDB_CONNECT_UPSERT(CnAccdb, "T_ファイル情報", where, fieldArray, setArray)
        MALLDB_DISCONNECT()

        DGV22DB()
    End Sub

    'csvファイルの中身を確認する
    Private Function CSVcheck(fName As String, tName As String, kindMark As String) As String
        Dim csvArray As ArrayList = TM_CSV_READ(fName)(0)
        Dim csvHeaderArray As String() = Split(csvArray(0).ToString, "|=|")

        'モールを特定
        Dim Table1 As DataTable = MALLDB_CONNECT_SELECT(CnAccdb, "T_基本情報", "[店名]='" & tName & "'", "表示順")
        Dim mall As String = Table1.Rows(0)("モール")
        Table1.Clear()

        'テンプレート読出
        Table1 = MALLDB_CONNECT_SELECT(CnAccdb, "T_設定_テンプレート", "[モール]='" & mall & "' AND [種類]='" & kindMark & "'")
        If Table1.Rows.Count > 0 Then
            Dim res As String = Replace(Table1.Rows(0)("ヘッダー"), " MEMO", "")

            '変換必要かチェック
            Dim comArray As String() = Split(res, "|#|")
            If comArray(0) = "変換" Then
                res = comArray(2)
                Dim headerArray As String() = Split(res, ",")
                '変換の場合は、テンプレートのヘッダーがcsvにあるか確認する
                For i As Integer = 0 To headerArray.Length - 1
                    If headerArray(i) = "" Then
                        Continue For
                    End If
                    If Not csvHeaderArray.Contains(headerArray(i)) Then
                        'MsgBox(csvHeaderArray(i) & "が適合しません", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
                        Return "error:" & headerArray(i)
                        Exit Function
                    End If
                Next
                Return "True"
            Else
                res = Regex.Replace(res, "\[|\]", "")
                Dim dgvHeaderArray As String() = Split(res, ",")
                'csvのヘッダーがテンプレートにあるか確認する
                For i As Integer = 0 To csvHeaderArray.Length - 1
                    If Not dgvHeaderArray.Contains(csvHeaderArray(i)) Then
                        'MsgBox(csvHeaderArray(i) & "が適合しません", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
                        Return "error:" & csvHeaderArray(i)
                        Exit Function
                    End If
                Next
                Return "True"
            End If
        Else
            'MsgBox("csvの種類が特定できません", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
            Return "error:csvの種類が特定できません"
            Exit Function
        End If

        Return "True"
    End Function

    'ファイルの更新時間を表示する
    Private Sub FileGetLastTime(fName As String)
        LIST4VIEW("ファイル更新時間チェック", "s")
        Dim da As Date = File.GetLastWriteTime(fName)
        Dim where As String = "[ファイル名]='" & Path.GetFileNameWithoutExtension(fName) & "'"
        Dim fieldArray As New ArrayList From {
                "ファイル名", "更新"
            }
        Dim setArray As New ArrayList From {
                Path.GetFileNameWithoutExtension(fName), da
            }
        MALLDB_CONNECT_UPSERT(CnAccdb, "T_ファイル情報", where, fieldArray, setArray)

        'サーバーのファイルに更新時間を保存
        TableGetLastTime(1)

        LIST4VIEW("ファイル更新時間チェック", "e")
    End Sub

    'サーバーにファイルの更新時間を保存
    'サーバーのファイルより新しい時は、ファイルを保存
    'Private Sub LastTimeSave(fName As String)
    '    TableName = "T_ファイル情報"
    '    '=================================================
    '    Table0 = MALLDB_CONNECT_SELECT(CnAccdb, TableName)
    '    '=================================================

    '    Dim str As String = ""
    '    For r As Integer = 0 To dgv9.RowCount - 1
    '        str &= Table0.Rows(r)("ファイル名") & "," & Table0.Rows(r)("更新") & vbCrLf
    '    Next

    '    Dim serverPath As String = Form1.サーバーToolStripMenuItem.Text
    '    Dim fPath As String = serverPath & "\tenpo\tenpo_lasttime.dat"
    '    Dim savePath As String = serverPath & "\tenpo\" & Path.GetFileName(fName)

    '    Try
    '        'サーバーのファイルから更新時間を読み込み比較
    '        TableGetLastTime(1)


    '        My.Computer.FileSystem.WriteAllText(fPath, str, True)
    '        My.Computer.FileSystem.CopyFile(fName, savePath, True)
    '    Catch ex As Exception
    '        'デバッグ環境の場合エラー出るので対応
    '    End Try
    'End Sub

    ''' <summary>
    ''' サーバーのデータと更新時間を比較
    ''' </summary>
    ''' <param name="mode">0=データ表示用、1=サーバー保存用</param>
    Private Sub TableGetLastTime(mode As Integer)
        LIST4VIEW("更新時間を読出し", "s")

        Dim serverPath As String = Form1.サーバーToolStripMenuItem.Text
        If InStr(Form1.appPath, "takashi") > 0 And InStr(Form1.appPath, "debug") > 0 Then
            serverPath = SpecialDirectories.MyDocuments & "\test"
        End If
        Dim localPath As String = Path.GetDirectoryName(Form1.appPath) & "\tenpo\"
        Dim fPath As String = serverPath & "\tenpo\tenpo_lasttime.dat"
        Dim lines As String() = Split(My.Computer.FileSystem.ReadAllText(fPath), vbCrLf)
        Dim sNameArray As New ArrayList
        Dim timeArray As New ArrayList
        For i As Integer = 0 To lines.Length - 1
            Dim t As String() = Split(lines(i), ",")
            If t(0) <> "" Then
                sNameArray.Add(t(0))
                timeArray.Add(t(1))
            End If
        Next

        TableName = "T_ファイル情報"
        '=================================================
        Table0 = MALLDB_CONNECT_SELECT(CnAccdb, TableName)
        '=================================================

        Dim str As String = ""
        DGV9 = DGV9
        For r As Integer = 0 To DGV9.RowCount - 1
            Dim rowName1 As String = DGV9.Item(0, r).Value
            Dim rowName2 As String = DGV9.Item(5, r).Value
            For c As Integer = 1 To DGV9.ColumnCount - 1 - 2
                Dim colName As String = DGV9.Columns(c).HeaderText
                If rowName1 = "マスタ" And colName = "item" Then colName = "master"
                If rowName1 = "マスタ" And colName = "select" Then colName = "page"

                Dim flag As Boolean = True
                For i As Integer = 0 To Table0.Rows.Count - 1
                    If Table0.Rows(i)("ファイル名") = rowName1 & "_" & colName Or Table0.Rows(i)("ファイル名") = rowName2 & "_" & colName Then
                        DGV9.Item(c, r).Value = Table0.Rows(i)("更新")
                        Dim da As Date = Table0.Rows(i)("更新")
                        If mode = 0 Then
                            If DateDiff(DateInterval.Day, Now, da) < 0 Then
                                DGV9.Item(c, r).Style.BackColor = Color.LightGray
                            Else
                                DGV9.Item(c, r).Style.BackColor = Color.White
                            End If
                        End If

                        'サーバーのファイル更新日をチェック
                        If sNameArray.Contains(Table0.Rows(i)("ファイル名")) Then
                            Dim sa As Date = timeArray(sNameArray.IndexOf(Table0.Rows(i)("ファイル名")))
                            If DateDiff(DateInterval.Second, sa, da) < 0 Then   '更新時間比較
                                If mode = 0 Then
                                    DGV9.Item(c, r).Style.BackColor = Color.Yellow
                                ElseIf mode = 1 Then
                                    str &= Table0.Rows(i)("ファイル名") & "," & sa & vbCrLf
                                    flag = False
                                End If
                            ElseIf DateDiff(DateInterval.Second, sa, da) = 0 Then

                            Else
                                If mode = 1 Then    'これだと全部飛んでしまう
                                    Try
                                        'ファイル保存
                                        Dim sPath As String = localPath & Table0.Rows(i)("ファイル名") & ".csv"
                                        Dim dPath As String = serverPath & "\tenpo\" & Table0.Rows(i)("ファイル名") & ".csv"
                                        File.Copy(sPath, dPath, True)
                                    Catch ex As Exception
                                        'デバッグ環境の場合エラー出るので対応
                                    End Try
                                End If
                            End If
                        End If

                        Exit For
                    End If
                Next

                If flag = True And mode = 1 Then
                    If DGV9.Item(c, r).Value.ToString <> "" Then
                        str &= rowName1 & "_" & colName & "," & DGV9.Item(c, r).Value.ToString & vbCrLf
                    End If
                End If
            Next
        Next

        If mode = 1 Then
            Try
                My.Computer.FileSystem.WriteAllText(fPath, str, False)
            Catch ex As Exception

            End Try
        End If

        '----------
        Table0.Clear()
        '----------
        LIST4VIEW("更新時間を読出し", "e")
    End Sub

    'Private Sub FileGetLastTime()
    '    LIST4VIEW("ファイル更新時間チェック", "s")
    '    For r As Integer = 0 To DataGridView9.Rows.Count - 1
    '        For c As Integer = 1 To DataGridView9.Columns.Count - 1
    '            DataGridView9.Item(c, r).Value = ""
    '        Next
    '    Next

    '    Dim folderPath As String = Path.GetDirectoryName(Form1.appPath) & "\tenpo\"
    '    If File.Exists(folderPath & DataGridView9.Item(0, 0).Value & "_master.csv") = True Then
    '        Dim da As Date = File.GetLastWriteTime(folderPath & DataGridView9.Item(0, 0).Value & "_master.csv")
    '        DataGridView9.Item(1, 0).Value = da
    '        If DateDiff(DateInterval.Day, Now, da) < 0 Then
    '            DataGridView9.Item(1, 0).Style.BackColor = Color.LightGray
    '        Else
    '            DataGridView9.Item(1, 0).Style.BackColor = Color.White
    '        End If
    '    End If

    '    Dim fName As String() = New String() {"_item", "_select", "_item-cat"}
    '    For r As Integer = 1 To DataGridView9.Rows.Count - 1
    '        For i As Integer = 0 To 2
    '            If File.Exists(folderPath & DataGridView9.Item(0, r).Value & fName(i) & ".csv") = True Then
    '                Dim da As Date = File.GetLastWriteTime(folderPath & DataGridView9.Item(0, r).Value & fName(i) & ".csv")
    '                DataGridView9.Item(i + 1, r).Value = da
    '                Dim a = DateDiff(DateInterval.Day, Now, da)
    '                If DateDiff(DateInterval.Day, Now, da) < 0 Then
    '                    DataGridView9.Item(i + 1, r).Style.BackColor = Color.LightGray
    '                Else
    '                    DataGridView9.Item(i + 1, r).Style.BackColor = Color.White
    '                End If
    '            End If
    '        Next
    '    Next
    'End Sub

    Private Sub CSV_TO_DB(fName As String, kmName As String, Optional linesArray As ArrayList = Nothing)
        MALLDB_CONNECT(CnAccdb)

        'テーブル名を決める
        TableName = "T_基本情報"
        where = "店名='" & ComboBox7.SelectedItem & "'"
        Dim Table1 As DataTable = MALLDB_GET_STR(TableName, "モール,TB名", where)
        Dim TableDst As String = "T_" & Table1(0)("TB名") & "_" & kmName

        'テーブルが存在しているか確認する
        'New Object() {TABLE_CATALOG, TABLE_SCHEMA, TABLE_NAME, TABLE_TYPE}
        'Dim schemaTable As DataTable = Cn0.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, TableDst, "TABLE"})
        Dim schemaTable As DataTable = MALLDB_TABLE_EXISTS(Nothing, Nothing, TableDst, "TABLE")

        'If schemaTable.Rows.Count = 0 Then
        '    Exit Sub
        'End If

        'ヘッダーを読み出す
        TableName = "T_設定_テンプレート"
        'If Table1(0)("モール") = "マスタ" Then
        '    kmName = "page"
        'End If
        where = "モール='" & Table1(0)("モール") & "' AND 種類='" & kmName & "'"
        Dim Table2 As DataTable = MALLDB_GET_STR(TableName, "コード,ヘッダー", where)
        Dim createHeaderList As String = Table2.Rows(0)("ヘッダー")
        Dim codeHeader As String = Table2.Rows(0)("コード")

        '--------------------------------------------------------------------
        '変換かチェック
        Dim henkanFlag As Boolean = False
        Dim yahuokuFlag As Boolean = False
        Dim motoHeaderList As String() = Nothing
        Dim comArray As String() = Split(createHeaderList, "|#|")
        If comArray(0) = "変換" Then
            createHeaderList = comArray(1)
            motoHeaderList = Split(comArray(2), ",")    '変換元ヘッダー
            henkanFlag = True
            yahuokuFlag = False
        ElseIf Regex.IsMatch(comArray(0), "^管理番号") Then
            henkanFlag = True
            yahuokuFlag = True
        End If
        '--------------------------------------------------------------------

        If Not henkanFlag Then
            '存在するテーブルを削除する
            If schemaTable.Rows.Count > 0 Then
                LIST4VIEW("テーブル削除", "s")
                Dim SQL1 As String = "DROP TABLE [" & TableDst & "] "
                SQLCm0.CommandText = SQL1
                SQLCm0.ExecuteNonQuery()
            End If

            'テーブルを作成する
            LIST4VIEW("テーブル作成", "s")
            Dim SQL As String = "CREATE TABLE [" & TableDst & "] "
            SQL &= "(" & createHeaderList & ")"
            SQLCm0.CommandText = SQL
            SQLCm0.ExecuteNonQuery()
        End If

        '----------
        Table2.Clear()
        '----------

        'ヘッダー生成
        Dim csvlines As String() = Split(linesArray(0), "|=|")                          '読込csvのヘッダー
        Dim chlist As String() = Split(Replace(createHeaderList, " MEMO", ""), ",")     'テーブル作成のヘッダー
        setField = ""
        If Not henkanFlag Then
            For i As Integer = 0 To csvlines.Length - 1
                If csvlines(i) <> "" Then
                    If InStr(csvlines(i), "[") > 0 Then
                        setField &= csvlines(i) & ","
                    Else
                        setField &= "[" & csvlines(i) & "],"
                    End If
                End If
            Next
            setField = setField.TrimEnd(",")
        Else
            For i As Integer = 0 To chlist.Length - 1
                If chlist(i) <> "" Then
                    If InStr(chlist(i), "[") > 0 Then
                        setField &= chlist(i) & ","
                    Else
                        setField &= "[" & chlist(i) & "],"
                    End If
                End If
            Next
            setField = setField.TrimEnd(",")
        End If

        'データ書き込み
        LIST4VIEW("テーブル書き込み", "s")
        ToolStripProgressBar1.Maximum = linesArray.Count
        ToolStripProgressBar1.Value = 0
        For r As Integer = 1 To linesArray.Count - 1
            setStr = ""
            linesArray(r) = AccessStrConvert(linesArray(r))     'Access登録用文字変換
            Dim lines As String() = Split(linesArray(r), "|=|")
            If Not henkanFlag And Not yahuokuFlag Then
                For i As Integer = 0 To lines.Count - 1
                    setStr &= "'" & lines(i) & "',"
                Next
                setStr = setStr.TrimEnd(",")
                '=================================================
                Dim flag As String = MALLDB_SET_STR(TableDst, setField, setStr)
                If flag <> "True" Then
                    MsgBox(flag)
                    Exit For
                End If
                '=================================================
            ElseIf henkanFlag And yahuokuFlag Then      'ヤフオクデータは追加
                Dim setArray As New ArrayList
                For i As Integer = 0 To chlist.Count - 1
                    If csvlines.Contains(chlist(i)) Then
                        Dim n As Integer = Array.IndexOf(csvlines, chlist(i))
                        setArray.Add(lines(n))
                    Else
                        setArray.Add("")
                    End If
                Next
                '=================================================
                Dim where As String = ""
                If csvlines.Contains(codeHeader) Then
                    where &= "[" & codeHeader & "]='" & lines(Array.IndexOf(csvlines, codeHeader)) & "'"
                End If

                Dim FieldArray As New ArrayList(chlist)
                MALLDB_UPSERT(TableDst, where, FieldArray, setArray)
                '=================================================
            Else    '変換の時
                Dim setArray As New ArrayList
                For i As Integer = 0 To chlist.Count - 1
                    If csvlines.Contains(motoHeaderList(i)) Then
                        Dim n As Integer = Array.IndexOf(csvlines, motoHeaderList(i))
                        setArray.Add(lines(n))
                    Else
                        setArray.Add("")
                    End If
                Next
                '=================================================
                Dim where As String = ""
                Dim cH As String() = Split(codeHeader, "#")
                If csvlines.Contains(cH(1)) Then
                    where &= "[" & cH(0) & "]='" & lines(Array.IndexOf(csvlines, cH(1))) & "'"
                End If

                Dim FieldArray As New ArrayList(chlist)
                MALLDB_UPSERT(TableDst, where, FieldArray, setArray)
                '=================================================
            End If

            ToolStripProgressBar1.Value += 1
            Application.DoEvents()
        Next
        ToolStripProgressBar1.Value = 0

        MALLDB_DISCONNECT()
        '----------
        Table1.Clear()
        '----------

        LIST4VIEW("テーブル生成", "e")
    End Sub

    'マスターデータ読込
    Private Sub MasterDB()
        MALLDB_CONNECT(CnAccdb)
        'Dim schemaTable As DataTable = MALLDB_TABLE_EXISTS(Nothing, Nothing, TableDst, "TABLE")
        'If schemaTable.Rows.Count = 0 Then
        '    MALLDB_DISCONNECT()
        '    Exit Sub
        'End If
        Dim TableDst1 As String = "T_MASTER_master"
        dgv5table = MALLDB_GET_SELECT(TableDst1, "", "商品コード")
        Table_to_DGV(DGV5, dgv5table)
        MALLDB_DISCONNECT()
    End Sub

    'ヤフオク欠品
    Private Sub DGV22DB()
        MALLDB_CONNECT(CnAccdb)
        Dim TableDst1 As String = "T_KTYA_auctions"
        Dim serverPath As String = Form1.サーバーToolStripMenuItem.Text
        Dim fPath As String = serverPath & "\tenpo\tenpo_lasttime2.txt"

        Dim schemaTable As DataTable = MALLDB_TABLE_EXISTS(Nothing, Nothing, TableDst1, "TABLE")
        If schemaTable.Rows.Count = 0 Then
            Dim createSql As String = "Create TABLE [T_KTYA_auctions](
                                    [管理番号] longtext, [カテゴリ] longtext, 
                                    [タイトル] longtext, [説明] longtext, 
                                    [ストア内商品検索用キーワード] longtext, 
                                    [開始価格] longtext, [即決価格] longtext, 
                                    [値下げ交渉] longtext, [個数] longtext, 
                                    [入札個数制限] longtext, [期間] longtext, 
                                    [終了時間] longtext, 
                                    [商品発送元の都道府県] longtext, 
                                    [商品発送元の市区町村] longtext, 
                                    [送料負担] longtext, 
                                    [代金先払い、後払い] longtext, 
                                    [落札ナビ決済方法設定] longtext, 
                                    [商品の状態] longtext, 
                                    [商品の状態備考] longtext, 
                                    [返品の可否] longtext, 
                                    [返品の可否備考] longtext, 
                                    [画像1] longtext, 
                                    [画像1コメント] longtext, 
                                    [画像2] longtext, 
                                    [画像2コメント] longtext, 
                                    [画像3] longtext, 
                                    [画像3コメント] longtext, 
                                    [画像4] longtext, 
                                    [画像4コメント] longtext, 
                                    [画像5] longtext, 
                                    [画像5コメント] longtext, 
                                    [画像6] longtext, 
                                    [画像6コメント] longtext, 
                                    [画像7] longtext, 
                                    [画像7コメント] longtext, 
                                    [画像8] longtext, 
                                    [画像8コメント] longtext, 
                                    [画像9] longtext, 
                                    [画像9コメント] longtext, 
                                    [画像10] longtext, 
                                    [画像10コメント] longtext, 
                                    [最低評価] longtext, 
                                    [悪評割合制限] longtext, 
                                    [入札者認証制限] longtext, 
                                    [自動延長] longtext, 
                                    [商品の自動再出品] longtext, 
                                    [自動値下げ] longtext, 
                                    [最低落札価格] longtext, 
                                    [チャリティー] longtext, 
                                    [注目のオークション] longtext, 
                                    [太字テキスト] longtext, 
                                    [背景色] longtext, 
                                    [ストアホットオークション] longtext, 
                                    [目立ちアイコン] longtext, 
                                    [贈答品アイコン] longtext, 
                                    [Tポイントオプション] longtext, 
                                    [アフィリエイトオプション] longtext, 
                                    [荷物の大きさ] longtext, 
                                    [荷物の重量] longtext, 
                                    [その他配送方法1] longtext, 
                                    [その他配送方法1料金表ページリンク] longtext, 
                                    [その他配送方法1全国一律価格] longtext, 
                                    [その他配送方法2] longtext, 
                                    [その他配送方法2料金表ページリンク] longtext, 
                                    [その他配送方法2全国一律価格] longtext, 
                                    [その他配送方法3] longtext, 
                                    [その他配送方法3料金表ページリンク] longtext, 
                                    [その他配送方法3全国一律価格] longtext, 
                                    [その他配送方法4] longtext, 
                                    [その他配送方法4料金表ページリンク] longtext, 
                                    [その他配送方法4全国一律価格] longtext, 
                                    [その他配送方法5] longtext, 
                                    [その他配送方法5料金表ページリンク] longtext, 
                                    [その他配送方法5全国一律価格] longtext, 
                                    [その他配送方法6] longtext, 
                                    [その他配送方法6料金表ページリンク] longtext, 
                                    [その他配送方法6全国一律価格] longtext, 
                                    [その他配送方法7] longtext, 
                                    [その他配送方法7料金表ページリンク] longtext, 
                                    [その他配送方法7全国一律価格] longtext, 
                                    [その他配送方法8] longtext, 
                                    [その他配送方法8料金表ページリンク] longtext, 
                                    [その他配送方法8全国一律価格] longtext, 
                                    [その他配送方法9] longtext, 
                                    [その他配送方法9料金表ページリンク] longtext, 
                                    [その他配送方法9全国一律価格] longtext, 
                                    [その他配送方法10] longtext, 
                                    [その他配送方法10料金表ページリンク] longtext, 
                                    [その他配送方法10全国一律価格] longtext, 
                                    [海外発送] longtext, 
                                    [配送方法・送料設定] longtext, 
                                    [代引手数料設定] longtext, 
                                    [消費税設定] longtext, 
                                    [JANコード・ISBNコード] longtext, 
                                    [メンバーズオークション] longtext, 
                                    [ブランドID] longtext, 
                                    [商品スペックサイズ種別] longtext, 
                                    [商品スペックサイズID] longtext, 
                                    [商品分類ID] longtext)"
            Dim Cn1 As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CnAccdb)
            Dim SQLCm1 = Cn1.CreateCommand
            Cn1.Open()
            SQLCm1.CommandText = createSql
            SQLCm1.ExecuteNonQuery()
            Cn1.Close()
            SQLCm1.Dispose()
            Cn1.Dispose()
        End If

        dgv22table = MALLDB_GET_SELECT(TableDst1, "", "")
        Table_to_DGV(DGV22, dgv22table)

        Dim file As New System.IO.StreamReader(fPath)
        Dim servertime As Date = file.ReadToEnd()

        Dim Table1 As DataTable = MALLDB_CONNECT_SELECT(CnAccdb, "T_ファイル情報", "[ファイル名]='T_KTYA_auctions'", "")
        If Table1.Rows.Count = 0 Then
            Dim createSql As String = "insert into T_ファイル情報(ファイル名,更新) values ('T_KTYA_auctions', #2020/01/01 00:00:00#)"
            Dim Cn1 As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CnAccdb)
            Dim SQLCm1 = Cn1.CreateCommand
            Cn1.Open()
            SQLCm1.CommandText = createSql
            SQLCm1.ExecuteNonQuery()
            Cn1.Close()
            SQLCm1.Dispose()
            Cn1.Dispose()

            Button5.Enabled = True
            Button5.BackColor = Color.LightYellow
        Else
            For i As Integer = 0 To Table1.Rows.Count - 1
                Dim da As Date = Table1.Rows(i)("更新")
                TextBox4.Text = da
                If DateDiff(DateInterval.Second, servertime, da) < 0 Then
                    Button5.Enabled = True
                    Button5.BackColor = Color.LightYellow
                Else
                    Button5.Enabled = False
                    Button5.BackColor = Color.Gray
                End If
            Next
        End If

        MALLDB_DISCONNECT()
    End Sub

    'ヘルプデータ読込
    Dim helpTable As DataTable = Nothing
    Dim helpArray As New ArrayList
    Private Sub HelpDB()
        MALLDB_CONNECT(CnAccdb)
        Dim TableDst As String = "T_ヘルプ"
        Dim schemaTable As DataTable = MALLDB_TABLE_EXISTS(Nothing, Nothing, TableDst, "TABLE")
        If schemaTable.Rows.Count = 0 Then
            MALLDB_DISCONNECT()
            Exit Sub
        End If

        helpTable = MALLDB_GET_SELECT(TableDst, "")
        For i As Integer = 0 To helpTable.Rows.Count - 1
            ListBox3.Items.Add(helpTable(i)("タイトル"))
            helpArray.Add(helpTable(i)("タイトル"))
        Next
        MALLDB_DISCONNECT()
    End Sub

    Private Sub ListBox3_DoubleClick(sender As Object, e As EventArgs) Handles ListBox3.DoubleClick
        TextBox14.Text = Regex.Replace(helpTable(helpArray.IndexOf(ListBox3.SelectedItem))("説明").ToString, "\|=\|", vbCrLf)
    End Sub

    ''' <summary>
    ''' Access登録用文字変換
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    Private Function AccessStrConvert(str As String) As String
        str = Replace(str, vbCrLf, "\n\r")
        str = Replace(str, vbCr, "\r")
        str = Replace(str, vbLf, "\n")
        str = Replace(str, "'", "Chr(39)")
        'str = Replace(str, ",", "Chr(44)")
        'str = Replace(str, "|", "Chr(124)")
        Return str
    End Function

    ''' <summary>
    ''' Access読出用文字変換
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    Private Function StrAccessConvert(str As String) As String
        str = Replace(str, "\n\r", vbCrLf)
        str = Replace(str, "\r", vbCr)
        str = Replace(str, "\n", vbLf)
        str = Replace(str, "Chr(39)", "'")
        str = Replace(str, "Chr(44)", ",")
        str = Replace(str, "Chr(124)", "|")
        Return str
    End Function

    ''' <summary>
    ''' Access登録用文字変換(ArrayList)
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    Private Function AccessArrayConvert(str As ArrayList) As ArrayList
        For i As Integer = 0 To str.Count - 1
            str(i) = Replace(str(i), vbCrLf, "\n\r")
            str(i) = Replace(str(i), vbCr, "\r")
            str(i) = Replace(str(i), vbLf, "\n")
            str(i) = Replace(str(i), "'", "Chr(39)")
        Next
        Return str
    End Function


    ''' <summary>
    ''' Access読出用文字変換(ArrayList)
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    Private Function ArrayAccessConvert(str As ArrayList) As ArrayList
        For i As Integer = 0 To str.Count - 1
            str(i) = Replace(str(i), "\n\r", vbCrLf)
            str(i) = Replace(str(i), "\r", vbCr)
            str(i) = Replace(str(i), "\n", vbLf)
            str(i) = Replace(str(i), "Chr(39)", "'")
            str(i) = Replace(str(i), "Chr(44)", ",")
            str(i) = Replace(str(i), "Chr(124)", "|")
        Next
        Return str
    End Function

    Private Sub ToolStripComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ToolStripComboBox2.SelectedIndexChanged
        Select Case True
            Case Regex.IsMatch(ToolStripComboBox2.SelectedItem, "マスタ")
                ToolStripComboBox1.SelectedItem = "master"
                ToolStripComboBox3.SelectedItem = "無し"
            Case Regex.IsMatch(ToolStripComboBox2.SelectedItem, "問屋|Sarada|トココ|Hewflit|ヤフオク|Qoo10")
                ToolStripComboBox1.SelectedItem = "item"
                ToolStripComboBox3.SelectedItem = "無し"
            Case Else
                ToolStripComboBox1.SelectedItem = "item"
                ToolStripComboBox3.SelectedItem = "select"
        End Select

        'Amazon画面の店舗表示
        Dim tenpoName As String = ""
        Select Case ToolStripComboBox2.SelectedItem
            Case "Sarada"
                tenpoName = "サラダ"
            Case "トココ"
                tenpoName = "トココ"
            Case "Hewflit"
                tenpoName = "フリット"
            Case Else
                tenpoName = "..."
        End Select
        ComboBox17.Text = tenpoName
    End Sub

    Private Sub ToolStripComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        Select Case ToolStripComboBox1.SelectedItem
            Case "select", "page", "item-cat"
                ToolStripComboBox3.SelectedItem = "無し"
        End Select
    End Sub

    '読込
    Dim CsvRead_TableName(1) As String
    Dim DgvTable(1) As DataTable
    Dim d As New DataView
    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        selChangeFlag = False

        LIST4VIEW("登録データ読込", "s")
        TableName = "T_基本情報"
        where = "店名='" & ToolStripComboBox2.SelectedItem & "'"
        '=================================================
        MALLDB_CONNECT(CnAccdb)
        Dim Table1 As DataTable = MALLDB_GET_SELECT(TableName, where).Copy
        '=================================================

        Label19.Visible = True
        Application.DoEvents()

        ToolStripStatusLabel4.Text = Table1(0)("モール")
        Dim order As New ArrayList
        Select Case ToolStripStatusLabel4.Text
            Case "楽天"
                order.Add("商品管理番号（商品URL）")
            Case "Yahoo"
                order.Add("code")
                order.Add("sub-code")
            Case "ヤフオク"
                order.Add("管理番号")
            Case "MakeShop"
                order.Add("独自商品コード")
            Case "Amazon"
                order.Add("出品者SKU")
            Case "Wowma"
                order.Add("itemCode")
            Case "Qoo10"
                order.Add("Seller Code")
            Case "マスタ"
                order.Add("商品コード")
                order.Add("代表商品コード")
        End Select

        Dim tscb As ToolStripComboBox() = New ToolStripComboBox() {ToolStripComboBox1, ToolStripComboBox3}
        Dim dgv As DataGridView() = New DataGridView() {DGV2, DGV4}
        For i As Integer = 0 To 1
            dgv(i).Rows.Clear()
            dgv(i).Columns.Clear()

            Dim TableDst As String = "T_" & Table1(0)("TB名") & "_" & tscb(i).SelectedItem
            CsvRead_TableName(i) = TableDst
            If order(0) = "出品者SKU" And tscb(i).SelectedItem = "select" Then
                order(0) = "seller-sku"
            End If
            Dim schemaTable As DataTable = MALLDB_TABLE_EXISTS(Nothing, Nothing, TableDst, "TABLE")
            If i = 0 Then
                If schemaTable.Rows.Count <> 0 Then
                    DgvTable(i) = MALLDB_GET_SELECT(TableDst, "", order(0))
                    Table_to_DGV(DGV2, DgvTable(i))
                End If
            Else
                If schemaTable.Rows.Count = 0 Then
                    SplitContainer7.Panel2Collapsed = True
                Else
                    SplitContainer7.Panel2Collapsed = False
                    DgvTable(i) = MALLDB_GET_SELECT(TableDst, "", order(0))
                    Table_to_DGV(DGV4, DgvTable(i))
                End If
            End If

            Dim dH As ArrayList = TM_HEADER_GET(dgv(i))

            '必須項目確認
            LIST4VIEW("必須項目確認読込", "s")
            TableName = "T_設定_テンプレート"
            If i = 0 Then
                where = "モール='" & ToolStripStatusLabel4.Text & "' AND 種類='" & ToolStripComboBox1.SelectedItem & "'"
            Else
                where = "モール='" & ToolStripStatusLabel4.Text & "' AND 種類='" & ToolStripComboBox3.SelectedItem & "'"
            End If
            Dim Table2 As DataTable = MALLDB_GET_SELECT(TableName, where)
            If Table2.Rows.Count > 0 Then
                'デフォルト保存ファイル名読込
                If i = 0 Then
                    TextBox29.Text = Table2.Rows(0)("ファイル名").ToString()
                    TextBox30.Text = ""     '前に読み込んだ時の片付け
                Else
                    TextBox30.Text = Table2.Rows(0)("ファイル名").ToString()
                End If

                '必須項目読込
                For k As Integer = 0 To 2
                    Dim headerName As String = ""
                    Dim backColor As Color = Color.LightYellow
                    If k = 0 Then
                        headerName = Table2.Rows(0)("必須項目").ToString
                        backColor = Color.LightYellow
                    ElseIf k = 1 Then
                        headerName = Table2.Rows(0)("必須項目2").ToString
                        backColor = Color.LightBlue
                    ElseIf k = 2 Then
                        headerName = Table2.Rows(0)("必須項目3").ToString
                        backColor = Color.GreenYellow
                    End If
                    If headerName <> "" Then
                        Dim reqHeaderArray As String() = Split(headerName, "|=|")
                        For p As Integer = 0 To reqHeaderArray.Length - 1
                            dgv(i).Columns(dH.IndexOf(reqHeaderArray(p))).HeaderCell.Style.BackColor = backColor
                        Next
                    End If
                Next
            End If
            Table2.Clear()
        Next
        Table1.Clear()

        Label19.Visible = False

        '=================================================
        MALLDB_DISCONNECT()
        '=================================================
        LIST4VIEW("登録データ読込", "e")

        ListBox1.Items.Clear()
        ComboBox22.Items.Clear()
        For c As Integer = 0 To DGV2.ColumnCount - 1
            ListBox1.Items.Add(DGV2.Columns(c).HeaderText)
            ComboBox22.Items.Add(DGV2.Columns(c).HeaderText)
        Next
        For i As Integer = order.Count - 1 To 0 Step -1
            If ListBox1.Items.Contains(order(i)) Then
                ListBox1.SelectedItem = order(i)
            End If
        Next
        If ComboBox22.Items.Count > 0 Then
            ComboBox22.SelectedIndex = 0
        End If

        Program_List_CB()

        'ヤフオク画像存在チェック
        If CheckBox31.Checked And order(0) = "管理番号" Then
            Dim dH2 As ArrayList = TM_HEADER_GET(DGV2)
            Dim headerList1 As String() = New String() {"画像1", "画像2", "画像3", "画像4", "画像5", "画像6", "画像7", "画像8", "画像9", "画像10"}
            For r As Integer = 0 To DGV2.RowCount - 1
                For c As Integer = 0 To headerList1.Length - 1
                    Dim fName As String = DGV2.Item(dH2.IndexOf(headerList1(c)), r).Value
                    If fName = "" Then
                        Continue For
                    End If
                    If Not File.Exists(TextBox32.Text & "\" & fName) Then
                        DGV2.Item(dH2.IndexOf(headerList1(c)), r).Style.BackColor = Color.Orange
                    End If
                Next
            Next
        End If

        ComboBox10.SelectedItem = ToolStripStatusLabel4.Text
        selChangeFlag = True
    End Sub

    Private Sub Program_List_CB()
        If ToolStripComboBox2.SelectedIndex >= 0 Then
            where = ""
            Dim orderby As String = "ID"
            TableName = "T_プログラム"
            '=================================================
            Dim Table As DataTable = MALLDB_CONNECT_SELECT(CnAccdb, TableName, where, orderby)
            '=================================================

            ComboBox5.Items.Clear()
            ComboBox6.Items.Clear()
            ComboBox11.Items.Clear()
            ComboBox5.ResetText()
            ComboBox6.ResetText()
            ComboBox11.ResetText()
            For i As Integer = 0 To Table.Rows.Count - 1
                If Table.Rows(i)("モール") = ToolStripStatusLabel4.Text Then
                    If Table.Rows(i)("名称") = "csv列表示変更" Then
                        ComboBox5.Items.Add(Table.Rows(i)("名称2"))
                    ElseIf Table.Rows(i)("名称") = "csv修正プログラム" Then
                        'If InStr(Table.Rows(i)("名称2"), "予約に変更") = False Then
                        ComboBox6.Items.Add(Table.Rows(i)("名称2"))
                        'End If
                    ElseIf Table.Rows(i)("名称") = "csv変換" Then
                        ComboBox11.Items.Add(Table.Rows(i)("名称2"))
                    End If
                End If
            Next

            If ComboBox5.Items.Count > 0 Then
                ComboBox5.SelectedIndex = 0
            End If
            If ComboBox6.Items.Count > 0 Then
                ComboBox6.SelectedIndex = 0
            End If
            If ComboBox11.Items.Count > 0 Then
                ComboBox11.SelectedIndex = 0
            End If

            '=================================================
            Table.Clear()
            '=================================================
        End If
    End Sub

    '保存
    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        Dim DR As DialogResult = MsgBox("データベースに保存しても良いですか？", MsgBoxStyle.YesNo)
        If DR = DialogResult.No Then
            Exit Sub
        End If

        LIST4VIEW("登録データ読込", "s")
        TableName = "T_基本情報"
        where = "店名='" & ToolStripComboBox2.SelectedItem & "'"
        '=================================================
        MALLDB_CONNECT(CnAccdb)
        Dim Table1 As DataTable = MALLDB_GET_SELECT(TableName, where).Copy
        '=================================================

        Label19.Visible = True
        Application.DoEvents()

        Dim order As String = ""
        Select Case ToolStripStatusLabel4.Text
            Case "楽天"
                order = "商品管理番号（商品URL）"
            Case "Yahoo"
                order = "code"
            Case "ヤフオク"
                order = "管理番号"
            Case "MakeShop"
                order = "独自商品コード"
            Case "Amazon"
                order = "出品者SKU"
            Case "マスタ"
                order = "商品コード"
        End Select

        Dim tscb As ToolStripComboBox() = New ToolStripComboBox() {ToolStripComboBox1, ToolStripComboBox3}
        Dim dgv As DataGridView() = New DataGridView() {DGV2, DGV4}
        For i As Integer = 0 To 1
            If dgv(i).Rows.Count = 0 Then
                Continue For
            End If

            Dim TableDst As String = "T_" & Table1(0)("TB名") & "_" & tscb(i).SelectedItem
            Dim schemaTable As DataTable = MALLDB_TABLE_EXISTS(Nothing, Nothing, TableDst, "TABLE")
            If schemaTable.Rows.Count = 0 Then
                Continue For
            End If

            Dim dH As ArrayList = TM_HEADER_GET(dgv(i))

            'データが正しいかチェックする
            Dim whereC As String = "[モール]='" & ToolStripStatusLabel4.Text & "' AND [種類]='" & tscb(i).SelectedItem & "'"
            Dim tableC As DataTable = MALLDB_GET_STR("T_設定_テンプレート", "保存必須", whereC)
            If Not dH.Contains(tableC(0)("保存必須")) Then
                MsgBox("保存データが正しくありません")
                Exit Sub
            End If

            If order = "出品者SKU" And tscb(i).SelectedItem = "select" Then
                order = "seller-sku"
            End If

            MALLDB_CONNECT(CnAccdb)
            '=================================================
            Dim FieldArray As New ArrayList
            For c As Integer = 0 To dgv(i).ColumnCount - 1
                FieldArray.Add(dgv(i).Columns(c).HeaderText)
            Next

            For r As Integer = 0 To dgv(i).RowCount - 1
                Dim where As String = "[" & order & "]='" & dgv(i).Item(dH.IndexOf(order), r).Value & "'"
                Dim setArray As New ArrayList
                For c As Integer = 0 To dgv(i).ColumnCount - 1
                    setArray.Add(dgv(i).Item(c, r).Value)
                Next
                MALLDB_UPSERT(TableDst, where, FieldArray, setArray)
            Next
            '=================================================

            'フォルダにファイルを保存
            Dim table0 As DataTable = MALLDB_GET_SQL(TableDst, "")
            Dim pathEnd As String = Path.GetDirectoryName(Form1.appPath) & "\tenpo\" & ToolStripComboBox2.SelectedItem & "_" & tscb(i).SelectedItem & ".csv"
            ConvertDataTableToCsv(table0, pathEnd, True)

            MALLDB_DISCONNECT()

            'ファイル更新時間をデータベースに保存
            FileGetLastTime(pathEnd)
        Next

        Label19.Visible = False
        LIST4VIEW("登録データ保存終了", "e")
        MsgBox("保存しました")
    End Sub

    '削除
    Private Sub ToolStripButton12_Click(sender As Object, e As EventArgs) Handles ToolStripButton12.Click
        Dim DR As DialogResult = MsgBox("表示行を元にデータベースから削除しても良いですか？", MsgBoxStyle.YesNo)
        If DR = DialogResult.No Then
            Exit Sub
        End If

        LIST4VIEW("登録データ読込", "s")
        TableName = "T_基本情報"
        where = "店名='" & ToolStripComboBox2.SelectedItem & "'"
        '=================================================
        MALLDB_CONNECT(CnAccdb)
        Dim Table1 As DataTable = MALLDB_GET_SELECT(TableName, where).Copy
        '=================================================

        Label19.Visible = True
        Application.DoEvents()

        Dim order As String = ""
        Select Case ToolStripStatusLabel4.Text
            Case "楽天"
                order = "商品管理番号（商品URL）"
            Case "Yahoo"
                order = "code"
            Case "ヤフオク"
                order = "管理番号"
            Case "MakeShop"
                order = "独自商品コード"
            Case "Amazon"
                order = "出品者SKU"
            Case "マスタ"
                order = "商品コード"
        End Select

        Dim tscb As ToolStripComboBox() = New ToolStripComboBox() {ToolStripComboBox1, ToolStripComboBox3}
        Dim dgv As DataGridView() = New DataGridView() {DGV2, DGV4}
        For i As Integer = 0 To 1
            If dgv(i).Rows.Count = 0 Then
                Continue For
            End If

            Dim TableDst As String = "T_" & Table1(0)("TB名") & "_" & tscb(i).SelectedItem
            Dim schemaTable As DataTable = MALLDB_TABLE_EXISTS(Nothing, Nothing, TableDst, "TABLE")
            If schemaTable.Rows.Count = 0 Then
                Continue For
            End If

            Dim dH As ArrayList = TM_HEADER_GET(dgv(i))

            'データが正しいかチェックする
            Dim whereC As String = "[モール]='" & ToolStripStatusLabel4.Text & "' AND [種類]='" & tscb(i).SelectedItem & "'"
            Dim tableC As DataTable = MALLDB_GET_STR("T_設定_テンプレート", "保存必須", whereC)
            If Not dH.Contains(tableC(0)("保存必須")) Then
                MsgBox("保存データが正しくありません")
                Exit Sub
            End If

            If order = "出品者SKU" And tscb(i).SelectedItem = "select" Then
                order = "seller-sku"
            End If

            MALLDB_CONNECT(CnAccdb)
            '=================================================
            For r As Integer = 0 To dgv(i).RowCount - 1
                Dim where As String = "[" & order & "]='" & dgv(i).Item(dH.IndexOf(order), r).Value & "'"
                MALLDB_DELETE(TableDst, where)
            Next
            '=================================================

            'フォルダにファイルを保存
            Dim table0 As DataTable = MALLDB_GET_SQL(TableDst, "")
            Dim pathEnd As String = Path.GetDirectoryName(Form1.appPath) & "\tenpo\" & ToolStripComboBox2.SelectedItem & "_" & tscb(i).SelectedItem & ".csv"
            ConvertDataTableToCsv(table0, pathEnd, True)

            MALLDB_DISCONNECT()

            'ファイル更新時間をデータベースに保存
            FileGetLastTime(pathEnd)
        Next

        ToolStripButton1.PerformClick()
        MsgBox("削除しました")
    End Sub



    '連結（メモ：必要なくなったら消してOK）
    'Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles ToolStripButton5.Click
    '    DataGridView2.Rows.Clear()
    '    DataGridView2.Columns.Clear()

    '    LIST4VIEW("登録データ読込", "s")
    '    TableName = "T_基本情報"
    '    where = "店名='" & ToolStripComboBox2.SelectedItem & "'"
    '    '=================================================
    '    MALLDB_CONNECT(CnAccdb)
    '    Dim Table1 As DataTable = MALLDB_GET_SELECT(TableName, where)
    '    '=================================================
    '    Dim TableDst As String = "T_" & Table1(0)("TB名") & "_" & ToolStripComboBox1.SelectedItem
    '    Dim TableDst2 As String = "T_" & Table1(0)("TB名") & "_select"

    '    Dim schemaTable As DataTable = MALLDB_TABLE_EXISTS(Nothing, Nothing, TableDst, "TABLE")
    '    If schemaTable.Rows.Count = 0 Then
    '        Exit Sub
    '    End If

    '    where = "INNER JOIN " & TableDst2 & " ON " & TableDst & ".商品管理番号（商品URL）=" & TableDst2 & ".商品管理番号（商品URL）"
    '    DataGridView2table = MALLDB_GET_SQL(TableDst, where)
    '    For r As Integer = 0 To DataGridView2table.Rows.Count - 1
    '        If r = 0 Then   'ヘッダー
    '            For c As Integer = 0 To DataGridView2table.Columns.Count - 1
    '                DataGridView2.Columns.Add("c" & c, DataGridView2table.Columns(c).ColumnName)
    '            Next
    '        End If

    '        DataGridView2.Rows.Add()
    '        For c As Integer = 0 To DataGridView2table.Columns.Count - 1
    '            DataGridView2.Item(c, r).Value = DataGridView2table.Rows(r)(c)
    '        Next
    '    Next

    '    '=================================================
    '    MALLDB_DISCONNECT()
    '    '=================================================

    '    ListBox1.Items.Clear()
    '    For c As Integer = 0 To DataGridView2.ColumnCount - 1
    '        ListBox1.Items.Add(DataGridView2.Columns(c).HeaderText)
    '    Next
    'End Sub

    Private Sub TabControl7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl7.SelectedIndexChanged
        'LIST4VIEW("タブ選択認識", "")
        Select Case TabControl3.SelectedTab.Text
            Case "データベース"
                上に表示されているデータを保存ToolStripMenuItem.Enabled = True
                下に表示されているデータを保存ToolStripMenuItem.Enabled = True
            Case Else
                上に表示されているデータを保存ToolStripMenuItem.Enabled = False
                下に表示されているデータを保存ToolStripMenuItem.Enabled = False
        End Select
    End Sub

    Private Sub TabControl3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl3.SelectedIndexChanged
        'LIST4VIEW("タブ選択認識", "")
        Select Case TabControl3.SelectedTab.Text
            Case "CSV"
                TabControl1.Enabled = True
                SetTabVisible(TabControl11, "csv")
            Case "Amazon"
                TabControl1.Enabled = False
                SetTabVisible(TabControl11, "Amazon")
            Case "ヤフオク"
                TabControl1.Enabled = False
                LISTVIEW_UPDATE2()
            Case Else
                TabControl1.Enabled = False
        End Select
    End Sub

    'TabColtrol表示切替
    Public Sub SetTabVisible(oTabControl As TabControl, tText As String)
        Dim nameArray As ArrayList = New ArrayList
        'For i As Integer = 0 To oTabControl.TabPages.Count - 1
        '    nameArray.Add(oTabControl.TabPages(i).Text)
        'Next
        nameArray.Add("csv")
        nameArray.Add("Amazon")

        For i As Integer = 0 To nameArray.Count - 1
            If Regex.IsMatch(nameArray(i), tText) Then
                SetTabVisible2(oTabControl, i, True)
            Else
                SetTabVisible2(oTabControl, i, False)
            End If
        Next
    End Sub

    Public Sub SetTabVisible2(oTabControl As TabControl, nIndex As Integer, bVisible As Boolean)
        Dim oAllTabPages As List(Of KeyValuePair(Of TabPage, Boolean))
        If oTabControl.Tag Is Nothing Then
            ' 全タブと、その表示状態を保持
            oAllTabPages = New List(Of KeyValuePair(Of TabPage, Boolean))
            For Each oTabPage As TabPage In oTabControl.TabPages
                Dim oTabAndVisible As New KeyValuePair(Of TabPage, Boolean)(oTabPage, True)
                oAllTabPages.Add(oTabAndVisible)
            Next
            oTabControl.Tag = oAllTabPages
        Else
            ' 全タブと、その表示状態を取得
            oAllTabPages = CType(oTabControl.Tag, List(Of KeyValuePair(Of TabPage, Boolean)))
        End If

        ' タブの表示状態を設定
        oAllTabPages(nIndex) = New KeyValuePair(Of TabPage, Boolean)(oAllTabPages(nIndex).Key, bVisible)

        oTabControl.TabPages.Clear()
        For Each oTabAndVisible As KeyValuePair(Of TabPage, Boolean) In oAllTabPages
            If oTabAndVisible.Value = True Then
                oTabControl.TabPages.Add(oTabAndVisible.Key)
            End If
        Next
    End Sub

    Private Function REGEX_REPLACE(pattern As String) As String
        If pattern <> "" Then
            pattern = Replace(pattern, "-", "\-")
            pattern = pattern.TrimEnd("|")
        End If

        Return pattern
    End Function

    '抽出実行
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If ListBox1.SelectedIndices.Count = 0 Or (TextBox1.Text = "" And TextBox3.Text = "") Then
            LIST4VIEW("リスト・テキスト未入力", "r")
            Exit Sub
        End If

        LIST4VIEW("抽出処理", "s")
        dH2 = TM_HEADER_GET(DGV2)

        TextBox9.Text = ""
        Dim pattern As String = ""
        If CheckBox5.Checked And TextBox3.Text <> "" Then
            TextBox3.Text = Regex.Replace(TextBox3.Text, "　| ", "")
            Dim listArray As String() = Split(TextBox3.Text, vbCrLf)
            Array.Sort(listArray)
            Dim str As String = ""
            For i As Integer = 0 To listArray.Count - 1
                If listArray(i) <> "" Then
                    str &= listArray(i) & vbCrLf
                End If
            Next
            TextBox3.Text = str

            If CheckBox34.Checked Then
                Dim lines As String = CODE_EXTRACTION(TextBox3.Text)
                TextBox9.Text = Replace(lines, "|", vbCrLf)
                pattern = Regex.Replace(TextBox9.Text, "\r\n|\r|\n", "|")
            Else
                pattern = Regex.Replace(TextBox3.Text, "\r\n|\r|\n", "|")
            End If
        ElseIf TextBox1.Text <> "" Then
            If CheckBox3.Checked Then
                pattern = Regex.Replace(TextBox1.Text, "\s|　", "|")
            Else
                pattern = TextBox1.Text
            End If
        Else
            Exit Sub
        End If

        If CheckBox3.Checked Then
            pattern = REGEX_REPLACE(pattern)

            Select Case ComboBox3.SelectedItem
                Case "完全一致"
                    pattern = Replace(pattern, "|", "$|^")
                    pattern = "^" & pattern & "$"
                Case "前方一致"
                    pattern = Replace(pattern, "|", "|^")
                    pattern = "^" & pattern
                Case "後方一致"
                    pattern = Replace(pattern, "|", "$|")
                    pattern = pattern & "$"
                Case Else

            End Select
        End If

        Dim dgvArray As DataGridView() = New DataGridView() {DGV2, DGV4}
        For d As Integer = 0 To dgvArray.Length - 1
            Dim dHX As ArrayList = TM_HEADER_GET(dgvArray(d))
            For r As Integer = dgvArray(d).RowCount - 1 To 0 Step -1
                Dim delFlag As Boolean = True
                For i As Integer = 0 To ListBox1.SelectedIndices.Count - 1
                    Dim h As Integer = dHX.IndexOf(ListBox1.SelectedItems(i))   'ヘッダー検索する
                    Dim str As String = dgvArray(d).Item(h, r).Value

                    If str <> "" Then
                        If CheckBox3.Checked Then   '正規表現
                            If CheckBox4.Checked Then
                                If Not Regex.IsMatch(str, pattern) Then
                                    delFlag = False
                                End If
                            Else
                                If Regex.IsMatch(str, pattern) Then
                                    delFlag = False
                                End If
                            End If
                        Else
                            If CheckBox4.Checked Then
                                If InStr(str, TextBox1.Text) = 0 Then
                                    delFlag = False
                                End If
                            Else
                                If InStr(str, TextBox1.Text) > 0 Then
                                    delFlag = False
                                End If
                            End If
                        End If
                    End If

                    If i = ListBox1.SelectedIndices.Count - 1 And delFlag = True Then
                        REMOVE_ROW_6(CheckBox6.Checked, dgvArray(d), r)
                    End If
                Next
            Next

            For r As Integer = dgvArray(d).RowCount - 1 To 0 Step -1
                If CheckBox32.Checked Then
                    Dim h As Integer = dHX.IndexOf(ComboBox22.SelectedItem)
                    Dim dgvStr As String = dgvArray(d).Item(h, r).Value
                    If IsNumeric(dgvStr) Then
                        If RadioButton13.Checked Then
                            If CInt(dgvStr) <= CInt(TextBox33.Text) Then
                                dgvArray(d).Item(h, r).Style.BackColor = Color.Yellow
                            End If
                        Else
                            If CInt(dgvStr) >= CInt(TextBox33.Text) Then
                                dgvArray(d).Item(h, r).Style.BackColor = Color.Yellow
                            End If
                        End If
                    End If
                End If
            Next


            '楽天「i」「s」のみ
            For r As Integer = dgvArray(d).RowCount - 1 To 0 Step -1
                If dgvArray(d).Rows(r).Visible = False Then
                    Continue For
                End If
                If RadioButton7.Checked Or RadioButton8.Checked Then
                    If dHX.Contains("選択肢タイプ") Then
                        If RadioButton7.Checked And dgvArray(d).Item(dHX.IndexOf("選択肢タイプ"), r).Value <> "i" Then
                            REMOVE_ROW_6(CheckBox6.Checked, dgvArray(d), r)
                        ElseIf RadioButton8.Checked And dgvArray(d).Item(dHX.IndexOf("選択肢タイプ"), r).Value <> "s" Then
                            REMOVE_ROW_6(CheckBox6.Checked, dgvArray(d), r)
                        End If
                    End If
                End If
            Next

            If d = 1 And CheckBox5.Checked Then
                For r As Integer = dgvArray(d).RowCount - 1 To 0 Step -1
                    CodeHojo(ToolStripStatusLabel4.Text, dgvArray(d), r)
                Next

                '枝番検索して削除
                If CheckBox37.Checked Then
                    For r As Integer = dgvArray(d).RowCount - 1 To 0 Step -1
                        CodeHojoDel(ToolStripStatusLabel4.Text, dgvArray(d), r)
                    Next
                End If
            End If
        Next
        LIST4VIEW("抽出処理", "e")
    End Sub

    Private Sub REMOVE_ROW_6(cb6 As Boolean, dgv As DataGridView, r As Integer)
        If cb6 Then
            dgv.Rows.RemoveAt(r)
        Else
            dgv.Rows(r).Visible = False
        End If
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked Then
            TextBox1.Enabled = False
            TextBox3.Enabled = True
            TextBox9.Enabled = True
        Else
            TextBox1.Enabled = True
            TextBox3.Enabled = False
            TextBox9.Enabled = False
        End If
    End Sub

    '商品に枝番号ある時、サブコードをチェックし、そのコードにマーク（背景）を付けておく
    Private Sub CodeHojo(mall As String, dgv As DataGridView, r As Integer)
        Dim dH As ArrayList = TM_HEADER_GET(dgv)
        Dim listArray As String() = Split(TextBox3.Text, vbCrLf)
        Select Case mall
            Case "楽天"
                If dgv.Item(dH.IndexOf("選択肢タイプ"), r).Value = "i" Then
                    Dim itemCode As String = dgv.Item(dH.IndexOf("商品管理番号（商品URL）"), r).Value
                    Dim hCode As String = Replace(dgv.Item(dH.IndexOf("項目選択肢別在庫用横軸選択肢子番号"), r).Value, "=", "")
                    Dim vCode As String = Replace(dgv.Item(dH.IndexOf("項目選択肢別在庫用縦軸選択肢子番号"), r).Value, "=", "")
                    Dim codeArray As New ArrayList From {
                        itemCode & hCode & vCode,
                        itemCode & Replace(hCode, "-", "") & vCode,
                        itemCode & hCode & Replace(vCode, "-", ""),
                        itemCode & vCode & hCode,
                        itemCode & Replace(vCode, "-", "") & hCode,
                        itemCode & vCode & Replace(hCode, "-", "")
                    }

                    For i As Integer = 0 To codeArray.Count - 1
                        If listArray.Contains(codeArray(i)) Then
                            dgv.Item(dH.IndexOf("商品管理番号（商品URL）"), r).Style.BackColor = Color.LightSkyBlue
                            dgv.Item(dH.IndexOf("項目選択肢別在庫用横軸選択肢子番号"), r).Style.BackColor = Color.LightSkyBlue
                            dgv.Item(dH.IndexOf("項目選択肢別在庫用縦軸選択肢子番号"), r).Style.BackColor = Color.LightSkyBlue
                            Exit Sub
                        End If
                    Next
                End If
            Case "Yahoo"
                If dgv.Item(dH.IndexOf("sub-code"), r).Value <> "" Then
                    If listArray.Contains(dgv.Item(dH.IndexOf("sub-code"), r).Value) Then
                        dgv.Item(dH.IndexOf("sub-code"), r).Style.BackColor = Color.LightSkyBlue
                    End If
                End If
            Case "Wowma"
                If dgv.Item(dH.IndexOf("stockSegment"), r).Value = "2" Then
                    Dim itemCode As String = dgv.Item(dH.IndexOf("itemCode"), r).Value
                    Dim hCode As String = Replace(dgv.Item(dH.IndexOf("choicesStockHorizontalCode"), r).Value, "=", "")
                    Dim vCode As String = Replace(dgv.Item(dH.IndexOf("choicesStockVerticalCode"), r).Value, "=", "")
                    Dim codeArray As New ArrayList From {
                        itemCode & hCode & vCode,
                        itemCode & Replace(hCode, "-", "") & vCode,
                        itemCode & hCode & Replace(vCode, "-", ""),
                        itemCode & vCode & hCode,
                        itemCode & Replace(vCode, "-", "") & hCode,
                        itemCode & vCode & Replace(hCode, "-", "")
                    }

                    For i As Integer = 0 To codeArray.Count - 1
                        If listArray.Contains(codeArray(i)) Then
                            dgv.Item(dH.IndexOf("itemCode"), r).Style.BackColor = Color.LightSkyBlue
                            dgv.Item(dH.IndexOf("choicesStockHorizontalCode"), r).Style.BackColor = Color.LightSkyBlue
                            dgv.Item(dH.IndexOf("choicesStockVerticalCode"), r).Style.BackColor = Color.LightSkyBlue
                            Exit Sub
                        End If
                    Next
                End If
        End Select
    End Sub

    '対象でないコードを見つけて削除する
    Private Sub CodeHojoDel(mall As String, dgv As DataGridView, r As Integer)
        Dim dH As ArrayList = TM_HEADER_GET(dgv)
        Dim listArray As String() = Split(TextBox3.Text, vbCrLf)
        Select Case mall
            Case "楽天"
                If dgv.Item(dH.IndexOf("選択肢タイプ"), r).Value = "i" Then
                    '枝番検索して削除
                    Dim eFlag As Boolean = False
                    For c As Integer = 0 To dgv.ColumnCount - 1
                        If dgv.Item(c, r).Style.BackColor = Color.LightSkyBlue Then
                            eFlag = True
                            Exit For
                        End If
                    Next
                    If eFlag = False Then
                        REMOVE_ROW_6(CheckBox6.Checked, dgv, r)
                    End If
                End If
            Case "Yahoo"
                If dgv.Item(dH.IndexOf("lead-time-instock"), r).Value <> "" Then
                    Dim eFlag As Boolean = False
                    For c As Integer = 0 To dgv.ColumnCount - 1
                        If dgv.Item(c, r).Style.BackColor = Color.LightSkyBlue Then
                            eFlag = True
                            Exit For
                        End If
                    Next
                    If eFlag = False Then
                        dgv.Item(dH.IndexOf("lead-time-instock"), r).Style.BackColor = Color.Pink
                    End If
                End If
            Case "Wowma"
                Dim eFlag As Boolean = False
                If dgv.Item(dH.IndexOf("stockSegment"), r).Value = "2" Then
                    '枝番検索して削除
                    For c As Integer = 0 To dgv.ColumnCount - 1
                        If dgv.Item(c, r).Style.BackColor = Color.LightSkyBlue Then
                            eFlag = True
                            Exit For
                        End If
                    Next
                ElseIf dgv.Item(dH.IndexOf("stockSegment"), r).Value = "1" Then
                    eFlag = False
                End If
                If eFlag = False Then
                    REMOVE_ROW_6(CheckBox6.Checked, dgv, r)
                End If
        End Select
    End Sub

    '''<summary>抽出コード作成 ＝ 返り値：正規表現用文字列（重複削除・ソート済）</summary>
    Private Function CODE_EXTRACTION(tStr As String) As String
        Dim res As String = ""
        tStr = Form1.StrConvToNarrow(tStr)

        Dim dH5 As ArrayList = TM_HEADER_GET(DGV5)
        Dim code As New ArrayList
        Dim dcode As New ArrayList
        For r As Integer = 0 To DGV5.Rows.Count - 1
            code.Add(DGV5.Item(dH5.IndexOf("商品コード"), r).Value)
            dcode.Add(DGV5.Item(dH5.IndexOf("代表商品コード"), r).Value)
        Next

        Dim answerArray As New ArrayList
        Dim errorArray As New ArrayList
        TextBox18.Text = ""
        Dim lines As String() = Split(tStr, vbCrLf)
        For i As Integer = 0 To lines.Length - 1
            If lines(i) = "" Then
                Continue For
            End If
            If code.Contains(lines(i)) Then
                Dim num As Integer = code.IndexOf(lines(i))
                If dcode(num) = "" Then
                    If Not answerArray.Contains(lines(i)) Then
                        answerArray.Add(lines(i))
                    End If
                Else
                    If Not answerArray.Contains(dcode(num)) Then
                        answerArray.Add(dcode(num))
                    End If
                End If
            ElseIf dcode.Contains(lines(i)) Then
                Dim num As Integer = dcode.IndexOf(lines(i))
                If Not answerArray.Contains(dcode(num)) Then
                    answerArray.Add(dcode(num))
                End If
            Else
                errorArray.Add(lines(i))
            End If
        Next

        If errorArray.Count > 0 Then
            For i As Integer = 0 To errorArray.Count - 1
                TextBox18.Text = errorArray(i) & vbCrLf
            Next
            TabControl5.SelectedTab = TabControl5.TabPages("TabPage10")
        End If

        answerArray.Sort()

        For i As Integer = 0 To answerArray.Count - 1
            res &= answerArray(i) & "|"
        Next
        res = res.TrimEnd("|")
        Return res
    End Function

    'Private Function CODE_EXTRACTION(tStr As String) As String
    '    Dim res As String = ""
    '    Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\location.csv"
    '    Dim csvRecords As New ArrayList()
    '    csvRecords =  TM_CSV_READ(fName)

    '    tStr = Form1.StrConvToNarrow(tStr)

    '    Dim answerArray As ArrayList = New ArrayList
    '    Dim lines As String() = Split(tStr, vbCrLf)
    '    For i As Integer = 0 To lines.Length - 1
    '        If lines(i) = "" Then
    '            Continue For
    '        End If
    '        For r As Integer = 0 To csvRecords.Count - 1
    '            Dim sArray As String() = Split(csvRecords(r), "|=|")
    '            If Regex.IsMatch(sArray(0), "^" & lines(i)) Then
    '                If sArray(3) = "" Then
    '                    If Not answerArray.Contains(sArray(0)) Then
    '                        answerArray.Add(sArray(0))
    '                    End If
    '                Else
    '                    If Not answerArray.Contains(sArray(3)) Then
    '                        answerArray.Add(sArray(3))
    '                    End If
    '                End If
    '                Exit For
    '            End If
    '        Next
    '    Next

    '    answerArray.Sort()

    '    For i As Integer = 0 To answerArray.Count - 1
    '        res &= answerArray(i) & "|"
    '    Next
    '    res = res.TrimEnd("|")
    '    Return res
    'End Function

    ''' <summary>
    ''' データを戻すボタン
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        UndoDataDGV()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        UndoDataDGV()
    End Sub

    Private Sub UndoDataDGV()
        Dim dgv As DataGridView() = New DataGridView() {DGV2, DGV4}
        For i As Integer = 0 To dgv.Length - 1
            If i = 1 And SplitContainer7.Panel2Collapsed Then
                Exit For
            End If
            Table_to_DGV(dgv(i), DgvTable(i))
        Next
    End Sub

    ''' <summary>
    ''' データグリッドビュー選択時処理
    ''' </summary>
    Dim selChangeFlag As Boolean = True
    Dim BeforeColor
    Private Sub DGV_SelectionChanged(sender As Object, e As EventArgs) Handles _
        DGV2.SelectionChanged, DGV4.SelectionChanged
        Dim dgv As DataGridView = sender
        Dim dH As ArrayList = TM_HEADER_GET(dgv)

        If selChangeFlag = False Or dgv.Rows.Count = 0 Then
            Exit Sub
        End If

        If dgv.SelectedCells.Count > 50 Then
            Exit Sub
        End If

        Dim code As String = ""
        If dgv.SelectedCells.Count > 0 Then
            If Not dgv.SelectedCells(0).Value Is Nothing Then
                AzukiControl1.Text = dgv.SelectedCells(0).Value.ToString
                If dgv Is DGV2 Then
                    ToolStripButton5.BackColor = Color.Yellow
                    ToolStripButton6.BackColor = SystemColors.GradientInactiveCaption
                Else
                    ToolStripButton5.BackColor = SystemColors.GradientInactiveCaption
                    ToolStripButton6.BackColor = Color.Yellow
                End If
            End If
            If ToolStripComboBox4.SelectedItem = "最終行" Then    'テキスト最終行表示
                Dim GyouBegin As Integer = AzukiControl1.GetLineHeadIndex(AzukiControl1.LineCount - 1)
                Dim GyouEnd As Integer = AzukiControl1.GetLineLength(AzukiControl1.LineCount - 1)
                AzukiControl1.Document.SetSelection(GyouBegin, GyouBegin)
                AzukiControl1.ScrollToCaret()
            End If
            dgv.Focus()
            If Not dgv.SelectedCells(0).Value Is Nothing Then
                ToolStripStatusLabel1.Text = "[文字数：" & dgv.SelectedCells(0).Value.ToString.Length & "]"    '文字数カウント
            End If

            'item、selectの連携表示
            If SplitContainer7.Panel2Collapsed = False And dgv Is DGV2 Then
                selChangeFlag = False
                Dim dgv4 As DataGridView = Me.DGV4
                If dgv4.RowCount > 0 Then
                    Try
                        For c As Integer = 0 To DGV2.ColumnCount - 1
                            If Regex.IsMatch(DGV2.Columns(c).HeaderText, "商品管理番号（商品URL）|code|itemCode") Then
                                code = DGV2.Item(c, DGV2.SelectedCells(0).RowIndex).Value
                                Exit For
                            End If
                        Next
                        Dim col As Integer = 0
                        For c As Integer = 0 To dgv4.ColumnCount - 1
                            If Regex.IsMatch(dgv4.Columns(c).HeaderText, "商品管理番号（商品URL）|code|itemCode") Then
                                col = c
                                Exit For
                            End If
                        Next
                        For r As Integer = 0 To dgv4.RowCount - 1
                            If code = dgv4.Item(col, r).Value Then
                                dgv4.Item(col, r).Selected = True
                                dgv4.CurrentCell = dgv4(col, r)
                                Exit For
                            End If
                        Next
                    Catch ex As Exception

                    End Try
                End If
                selChangeFlag = True
            End If
        End If

        '選択行が表示されていない時は選択を解除
        Dim selCell As ArrayList = New ArrayList
        For i As Integer = 0 To dgv.SelectedCells.Count - 1
            If dgv.Rows(dgv.SelectedCells(i).RowIndex).Visible = False Then
                selCell.Add(dgv.SelectedCells(i).ColumnIndex & "," & dgv.SelectedCells(i).RowIndex)
            End If
        Next
        For i As Integer = 0 To selCell.Count - 1
            Dim index As String() = Split(selCell(i), ",")
            dgv.Item(CInt(index(0)), CInt(index(1))).Selected = False
        Next

        '選択行の背景色を変更する
        If dgv.SelectedCells.Count > 50 Then
            Exit Sub
        End If

        For r As Integer = 0 To dgv.RowCount - 1
            dgv.Rows(r).DefaultCellStyle.BackColor = BeforeColor
        Next

        Dim selRow As ArrayList = New ArrayList
        For i As Integer = 0 To dgv.SelectedCells.Count - 1
            If Not selRow.Contains(dgv.SelectedCells(i).RowIndex) Then
                selRow.Add(dgv.SelectedCells(i).RowIndex)
            End If
        Next

        For i As Integer = 0 To selRow.Count - 1
            BeforeColor = dgv.Rows(selRow(i)).DefaultCellStyle.BackColor
            dgv.Rows(selRow(i)).DefaultCellStyle.BackColor = Color.LightCyan
        Next

        '選択状態の情報
        Label34.Text = dgv.Rows.GetRowCount(DataGridViewElementStates.Selected) & "/" & dgv.RowCount _
            & ":" & dgv.Columns.GetColumnCount(DataGridViewElementStates.Selected) & "/" & dgv.ColumnCount _
            & "-" & dgv.GetCellCount(DataGridViewElementStates.Selected)

        If Mall_subwindow.Visible Then
            If dgv.SelectedCells.Count > 0 Then
                Dim htmlStr As String = dgv.SelectedCells(0).Value
                'htmlStr = "<div width=750px>" & htmlStr & "</div>"
                If InStr(htmlStr, "<br>") = 0 Then
                    htmlStr = "<p>" & htmlStr & "</p>"
                End If
                Mall_subwindow.WebBrowser1.DocumentText = htmlStr
            End If
        End If

    End Sub

    Private Sub DGV_SelectionChanged2(sender As Object, e As EventArgs) Handles _
        DGV22.SelectionChanged
        Dim dgv As DataGridView = sender
        Dim dH As ArrayList = TM_HEADER_GET(dgv)

        If dgv.Rows.Count = 0 Then
            Exit Sub
        End If

        Dim code As String = ""
        If dgv.SelectedCells.Count > 0 Then
            If Not dgv.SelectedCells(0).Value Is Nothing Then
                AzukiControl3.Text = dgv.SelectedCells(0).Value.ToString
            End If
            If ToolStripComboBox4.SelectedItem = "最終行" Then    'テキスト最終行表示
                Dim GyouBegin As Integer = AzukiControl3.GetLineHeadIndex(AzukiControl3.LineCount - 1)
                Dim GyouEnd As Integer = AzukiControl3.GetLineLength(AzukiControl3.LineCount - 1)
                AzukiControl3.Document.SetSelection(GyouBegin, GyouBegin)
                AzukiControl3.ScrollToCaret()
            End If
            dgv.Focus()
        End If

        '選択行が表示されていない時は選択を解除
        Dim selCell As ArrayList = New ArrayList
        For i As Integer = 0 To dgv.SelectedCells.Count - 1
            If dgv.Rows(dgv.SelectedCells(i).RowIndex).Visible = False Then
                selCell.Add(dgv.SelectedCells(i).ColumnIndex & "," & dgv.SelectedCells(i).RowIndex)
            End If
        Next
        For i As Integer = 0 To selCell.Count - 1
            Dim index As String() = Split(selCell(i), ",")
            dgv.Item(CInt(index(0)), CInt(index(1))).Selected = False
        Next

        For r As Integer = 0 To dgv.RowCount - 1
            dgv.Rows(r).DefaultCellStyle.BackColor = BeforeColor
        Next

        Dim selRow As ArrayList = New ArrayList
        For i As Integer = 0 To dgv.SelectedCells.Count - 1
            If Not selRow.Contains(dgv.SelectedCells(i).RowIndex) Then
                selRow.Add(dgv.SelectedCells(i).RowIndex)
            End If
        Next

        For i As Integer = 0 To selRow.Count - 1
            BeforeColor = dgv.Rows(selRow(i)).DefaultCellStyle.BackColor
            dgv.Rows(selRow(i)).DefaultCellStyle.BackColor = Color.LightCyan
        Next

    End Sub

    Private Sub DGV_CellClick(sender As DataGridView, e As DataGridViewCellEventArgs) Handles DGV2.CellClick, DGV4.CellClick
        Dim code As String = ""
        For c As Integer = 0 To sender.ColumnCount - 1
            If Regex.IsMatch(sender.Columns(c).HeaderText, "商品管理番号（商品URL）|code|itemCode|独自商品コード|Seller Code") Then
                code = sender.Item(c, sender.SelectedCells(0).RowIndex).Value
                Exit For
            End If
        Next

        If code = "" Then
            RichTextBox1.Text = ""
            Exit Sub
        End If

        Dim dCode As String() = Split(code, "-")
        If dCode.Length > 0 Then
            'マスタ表示
            Dim dH5 As ArrayList = TM_HEADER_GET(DGV5)
            RichTextBox1.Text = ""
            For i As Integer = 0 To DGV5.RowCount - 1
                If dCode(0) = DGV5.Item(dH5.IndexOf("商品コード"), i).Value Or dCode(0) = DGV5.Item(dH5.IndexOf("代表商品コード"), i).Value Then
                    If RichTextBox1.Text = "" Then
                        RichTextBox1.Text &= dCode(0) & " " & DGV5.Item(dH5.IndexOf("商品名"), i).Value & vbCrLf
                    End If
                    RichTextBox1.Text &= "<" & DGV5.Item(dH5.IndexOf("商品区分"), i).Value & ">"
                    RichTextBox1.Text &= DGV5.Item(dH5.IndexOf("商品コード"), i).Value & " "
                    RichTextBox1.Text &= DGV5.Item(dH5.IndexOf("商品分類タグ"), i).Value & vbCrLf
                    RichTextBox1.Visible = True
                    'Exit For
                End If
            Next
        Else
            RichTextBox1.Text = ""
        End If
    End Sub

    ''' <summary>
    ''' テキストボックス入力をDataGridViewに反映
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub AzukiControl1_Validated(sender As Object, e As EventArgs) Handles AzukiControl1.Validated
        If AzukiControl1.Text = "" Then
            Exit Sub
        End If

        If ToolStripButton5.BackColor = Color.Yellow Then
            DGV2.CurrentCell.Value = AzukiControl1.Text
        Else
            DGV4.CurrentCell.Value = AzukiControl1.Text
        End If
    End Sub

    Private Sub AzukiControl1_PreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs) Handles AzukiControl1.PreviewKeyDown
        Dim dgv As DataGridView = DGV2
        If ToolStripButton5.BackColor = Color.Yellow Then
            dgv = DGV2
        Else
            dgv = DGV4
        End If

        '未完成
        'If e.Modifiers = Keys.Control And e.KeyCode = Keys.Enter Then
        '    AzukiControl1.Document.Text.Insert(AzukiControl1.Document.CaretIndex, vbLf)
        'ElseIf e.KeyCode = Keys.Enter Then
        '    dgv.EndEdit()
        'End If
    End Sub

    ''' <summary>
    ''' 名前を付けて保存
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub 名前を付けて保存ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 名前を付けて保存ToolStripMenuItem.Click
        Select Case TabControl7.SelectedTab.Text
            Case "データベース"
                If TabControl3.SelectedTab.Text = "Amazon" Then
                    If DGV16.RowCount = 0 Then
                        Exit Sub
                    End If
                    AmazonSave()
                Else
                    If DGV2.RowCount = 0 Then
                        Exit Sub
                    End If
                    NameChangeSave(0, "", DGV2)
                End If
            Case "各店比較"
                If DGV7.RowCount = 0 Then
                    Exit Sub
                End If
                NameChangeSave(0, "", DGV7, "9")
        End Select
    End Sub

    Private Sub 上に表示されているデータを保存ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 上に表示されているデータを保存ToolStripMenuItem.Click
        If DGV2.RowCount = 0 Then
            Exit Sub
        End If

        NameChangeSave(0, "", DGV2, "1")
    End Sub

    Private Sub 下に表示されているデータを保存ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 下に表示されているデータを保存ToolStripMenuItem.Click
        If DGV4.RowCount = 0 Then
            Exit Sub
        End If

        NameChangeSave(0, "", DGV4, "2")
    End Sub

    ''' <summary>
    ''' 名前を付けて保存
    ''' </summary>
    ''' <param name="mode">0=csv、1=ダブルコーテーション付、2=区切りタブ</param>
    ''' <param name="defaultPath"></param>
    ''' <param name="dgv"></param>
    ''' <param name="sel">0=上下、1=上、2=下、9=各店比較</param>
    Public Sub NameChangeSave(mode As Integer, defaultPath As String, dgv As DataGridView, Optional sel As String = "0")
        Dim sfd As New SaveFileDialog With {
            .Filter = "CSVファイル(*.csv)|*.csv|すべてのファイル(*.*)|*.*",
                    .AutoUpgradeEnabled = False,
            .FilterIndex = 0,
            .Title = "保存先のファイルを選択してください",
            .RestoreDirectory = True,
            .OverwritePrompt = True,
            .CheckPathExists = True
        }
        '.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),

        If InStr(defaultPath, "\") > 0 Then
            Dim sPath As String = Path.GetFileNameWithoutExtension(defaultPath)
            sfd.FileName = sPath & ".csv"
        Else
            sfd.FileName = "ファイル名自動.csv"
        End If

        'ダイアログを表示する
        If sfd.ShowDialog(Me) = DialogResult.OK Then
            Dim pathA As String = ""
            Dim pathB As String = ""
            Dim pathC As String = ""
            If Regex.IsMatch(sel, "[0|1]") Then
                If TextBox29.Text <> "" Then
                    pathA = Replace(sfd.FileName, "ファイル名自動.csv", TextBox29.Text)
                End If
                SaveCsv(pathA, mode, DGV2)
            End If
            If SplitContainer7.Panel2Collapsed = False And Regex.IsMatch(sel, "[0|2]") Then
                If TextBox30.Text <> "" Then
                    pathB = Replace(sfd.FileName, "ファイル名自動.csv", TextBox30.Text)
                End If
                SaveCsv(pathB, mode, DGV4)
            End If
            If sel = "9" Then
                pathC = Replace(sfd.FileName, "ファイル名自動.csv", "各店比較.csv")
                SaveCsv(sfd.FileName, mode, dgv)
            End If
            MsgBox(pathA & vbCrLf & pathB & vbCrLf & pathC & vbCrLf & "保存しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    Public Sub SaveCsv(fp As String, mode As Integer, dgv As DataGridView, Optional ENC As String = "SHIFT-JIS")
        Dim visibleFlag As Boolean = False
        For r As Integer = 0 To dgv.Rows.Count - 1
            If dgv.Rows(r).Visible = False Then
                Dim DR As DialogResult = MsgBox("表示行のみ保存しますか？", MsgBoxStyle.YesNo Or MsgBoxStyle.SystemModal)
                If DR = Windows.Forms.DialogResult.Yes Then
                    visibleFlag = True
                    Exit For
                Else
                    Exit For
                End If
            End If
        Next

        ' CSVファイルオープン
        Dim str As String = ""
        Dim sw As StreamWriter = Nothing
        Try
            sw = New StreamWriter(fp, False, Encoding.GetEncoding(ENC))
        Catch ex As Exception
            MsgBox("ファイルがエクセル等で開かれていませんか？閉じてから再度試してください。")
            sw.Dispose()
            Exit Sub
        End Try

        '区切り
        Dim delimiter As String = ","
        If mode = 2 Then
            delimiter = vbTab
        End If

        'ヘッダー行
        Dim dt As String = ""
        For c As Integer = 0 To dgv.Columns.Count - 1
            dt &= dgv.Columns(c).HeaderText & delimiter
        Next
        dt = dt.TrimEnd(delimiter) & vbLf
        sw.Write(dt)

        For r As Integer = 0 To dgv.Rows.Count - 1
            If visibleFlag = False Or (visibleFlag = True And dgv.Rows(r).Visible = True) Then
                For c As Integer = 0 To dgv.Columns.Count - 1
                    ' DataGridViewのセルのデータ取得
                    dt = ""
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
                        dt = dt & delimiter
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

    Public Sub AmazonSave()
        Dim sfd As New SaveFileDialog With {
            .Filter = "TSVファイル(*.tsv)|*.tsv|すべてのファイル(*.*)|*.*",
                    .AutoUpgradeEnabled = False,
            .FilterIndex = 0,
            .Title = "保存先のファイルを選択してください",
            .RestoreDirectory = True,
            .OverwritePrompt = True,
            .CheckPathExists = True,
            .FileName = "amazon.tsv"
        }
        '.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),

        'ダイアログを表示する
        If sfd.ShowDialog(Me) = DialogResult.OK Then
            Dim str As String = ""
            Dim dgv As DataGridView() = {DGV15, DGV16}
            For i As Integer = 0 To dgv.Length - 1
                For r As Integer = 0 To dgv(i).RowCount - 1
                    If Not dgv(i).Rows(r).IsNewRow Then
                        For c As Integer = 0 To dgv(i).ColumnCount - 1
                            str &= """" & dgv(i).Item(c, r).Value & """" & vbTab
                        Next
                        str = str.TrimEnd(vbTab)
                        str &= vbCrLf
                    End If
                Next
            Next
            File.WriteAllText(sfd.FileName, str, Encoding.Default)
            MsgBox(sfd.FileName & vbCrLf & "保存しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    Private Sub 名前を付けて保存xlsxToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 名前を付けて保存xlsxToolStripMenuItem.Click
        CSV_SAVE_XLSX()
    End Sub

    Private Sub CSV_SAVE_XLSX()
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

        Try
            FocusDGV.EndEdit()
        Catch ex As Exception

        End Try
        If FocusDGV Is Nothing Then
            FocusDGV = DGV2
        End If

        'ダイアログを表示する
        If sfd.ShowDialog(Me) = DialogResult.OK Then
            Dim sheet As Integer = 0
            Dim xWorkbook As IWorkbook = New XSSFWorkbook()
            Dim sheet1 As ISheet = xWorkbook.CreateSheet(sheet)
            Dim inRow As Integer = 0
            sheet1.CreateRow(0)
            For c As Integer = 0 To FocusDGV.ColumnCount - 1
                sheet1.GetRow(inRow).CreateCell(c).SetCellValue(FocusDGV.Columns(c).HeaderText)
            Next
            inRow += 1
            For r As Integer = 0 To FocusDGV.RowCount - 1
                If FocusDGV.Rows(r).IsNewRow Then
                    Continue For
                End If
                If FocusDGV.Rows(r).Visible = True Then
                    sheet1.CreateRow(inRow)
                    For c As Integer = 0 To FocusDGV.ColumnCount - 1
                        If Not FocusDGV.Item(c, r).Value Is DBNull.Value Then
                            Dim v As String = FocusDGV.Item(c, r).Value
                            If IsNumeric(v) Then
                                sheet1.GetRow(inRow).CreateCell(c).SetCellValue(CDbl(v))
                            Else
                                sheet1.GetRow(inRow).CreateCell(c).SetCellValue(v)
                            End If
                        End If

                        If FocusDGV.Item(c, r).Style.BackColor <> Color.Empty Then
                            Dim bk As Color = FocusDGV.Item(c, r).Style.BackColor
                            sheet1.GetRow(inRow).GetCell(c).CellStyle.FillBackgroundColor = IndexedColors.ValueOf(bk.Name).Index
                        End If
                    Next
                    inRow += 1
                End If
            Next

            Dim desk As String = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
            Using wfs = File.Create(sfd.FileName)
                xWorkbook.Write(wfs)
            End Using
            sheet1 = Nothing
            xWorkbook = Nothing
            MsgBox(sfd.FileName & vbCrLf & "保存しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    Private Sub ヤフオク専用保存csv画像ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ヤフオク専用保存csv画像ToolStripMenuItem.Click
        Dim saveFolder As String = ""
        Dim dlg = VistaFolderBrowserDialog1
        dlg.RootFolder = Environment.SpecialFolder.Desktop
        dlg.Description = "データを保存するフォルダを選択してください"
        If dlg.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
            saveFolder = dlg.SelectedPath
        End If
        If saveFolder = "" Then
            Exit Sub
        End If

        Dim csvPath As String = saveFolder & "\data_" & Format(Now, "yyyyMMddHHmmss") & ".csv"
        YahuokuCsv(csvPath, 0, DGV2)

        Dim fileDir As String = TextBox32.Text & "\"
        Dim dH2 As ArrayList = TM_HEADER_GET(DGV2)
        Dim errorStr As String = ""
        For r As Integer = 0 To DGV2.RowCount - 1
            For c As Integer = 1 To 10
                Dim fName As String = DGV2.Item(dH2.IndexOf("画像" & c), r).Value
                If fName <> "" Then
                    If File.Exists(fileDir & fName) Then
                        File.Copy(fileDir & fName, saveFolder & "\" & fName, True)
                    Else
                        errorStr &= fName & vbCrLf
                    End If
                End If
            Next
        Next

        Dim str As String = csvPath & vbCrLf & "保存しました"
        If errorStr <> "" Then
            str &= vbCrLf & vbCrLf & "下記ファイルエラー" & errorStr
        End If
        MsgBox(str, MsgBoxStyle.SystemModal)
    End Sub

    Public Sub YahuokuCsv(fp As String, mode As Integer, dgv As DataGridView, Optional ENC As String = "SHIFT-JIS")
        Dim visibleFlag As Boolean = False
        For r As Integer = 0 To dgv.Rows.Count - 1
            If dgv.Rows(r).Visible = False Then
                Dim DR As DialogResult = MsgBox("表示行のみ保存しますか？", MsgBoxStyle.YesNo Or MsgBoxStyle.SystemModal)
                If DR = Windows.Forms.DialogResult.Yes Then
                    visibleFlag = True
                    Exit For
                Else
                    Exit For
                End If
            End If
        Next

        ' CSVファイルオープン
        Dim str As String = ""
        Dim sw As StreamWriter = New StreamWriter(fp, False, Encoding.GetEncoding(ENC))

        '区切り
        Dim delimiter As String = ","
        If mode = 2 Then
            delimiter = vbTab
        End If

        'ヘッダー行
        Dim dt As String = ""
        For c As Integer = 0 To dgv.Columns.Count - 1
            dt &= dgv.Columns(c).HeaderText & delimiter
        Next
        dt = dt.TrimEnd(delimiter) & vbLf
        sw.Write(dt)

        Dim dH2 As ArrayList = TM_HEADER_GET(DGV2)

        'カテゴリを登録しないと書き出しできないようにする
        'カテゴリ部分は「0000|1111」で2つ以上登録できる。2つ以上登録すると、行数が2行になる
        For r As Integer = 0 To dgv.Rows.Count - 1
            If visibleFlag = False Or (visibleFlag = True And dgv.Rows(r).Visible = True) Then
                Dim cateCount As String() = Split(dgv.Item(dH2.IndexOf("カテゴリ"), r).Value, "|")
                For cc As Integer = 0 To cateCount.Length - 1

                    For c As Integer = 0 To dgv.Columns.Count - 1
                        ' DataGridViewのセルのデータ取得
                        dt = ""
                        If c = dH2.IndexOf("カテゴリ") Then
                            dt = cateCount(cc)
                        ElseIf dgv.Rows(r).Cells(c).Value Is Nothing = False Then
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
                            dt = dt & delimiter
                        End If
                        ' CSVファイル書込
                        sw.Write(dt)
                    Next
                    sw.Write(vbLf)

                Next
            End If
        Next

        ' CSVファイルクローズ
        sw.Close()
    End Sub

    '修正履歴取得
    Private CellErrFlg As Boolean = False
    Private CellValue
    Private Sub DataGridView_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles _
        DGV2.DataError, DGV4.DataError
        CellErrFlg = True
    End Sub

    Private Sub DataGridView_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles _
            DGV1.CellBeginEdit, DGV2.CellBeginEdit, DGV3.CellBeginEdit, DGV4.CellBeginEdit, DGV5.CellBeginEdit,
            DGV6.CellBeginEdit, DGV7.CellBeginEdit, DGV8.CellBeginEdit, DGV9.CellBeginEdit, DGV10.CellBeginEdit,
            DGV11.CellBeginEdit, DGV16.CellBeginEdit
        Dim dgv As DataGridView = sender
        CellErrFlg = False
        CellValue = dgv.CurrentCell.Value   '該当セルの値を格納しておく
    End Sub

    Private Sub DataGridView_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles _
            DGV1.CellEndEdit, DGV2.CellEndEdit, DGV3.CellEndEdit, DGV4.CellEndEdit, DGV5.CellEndEdit,
            DGV6.CellEndEdit, DGV7.CellEndEdit, DGV8.CellEndEdit, DGV9.CellEndEdit, DGV10.CellEndEdit,
            DGV11.CellEndEdit, DGV16.CellEndEdit
        Dim dgv As DataGridView = sender
        Dim dgv3 As DataGridView = Me.DGV3
        Dim CellChgFlg As Boolean = False
        If CellErrFlg = False Then
            For r As Integer = 0 To dgv3.Rows.Count - 1
                If dgv3.Item(0, r).Value = dgv.CurrentRow.Index Then
                    If dgv3.Item(1, r).Value = dgv.CurrentCell.ColumnIndex Then
                        dgv3.Item(2, r).Value = CellValue
                        dgv3.Item(3, r).Value = dgv.CurrentCell.Value
                        CellChgFlg = True
                        Exit For
                    End If
                End If
            Next
            If CellChgFlg = False Then
                Dim str As String() = New String() {dgv.CurrentRow.Index, dgv.CurrentCell.ColumnIndex, CellValue, dgv.CurrentCell.Value}
                dgv3.Rows.Add(str)
            End If
        End If

        '変更したセルに色をつける
        dgv.CurrentCell.Style.BackColor = Color.Yellow

        If sender Is DGV16 Then
            CheckDGV16("one", dgv.CurrentCell.ColumnIndex, dgv.CurrentCell.RowIndex)
        End If
    End Sub


    '************************************************************************************************************
    Dim masterCodeArray As New ArrayList
    Dim masterDaihyoArray As New ArrayList
    Dim dH2 As New ArrayList
    Dim dH4 As New ArrayList
    Dim dH5 As New ArrayList
    Dim errorColor As Color = Color.Yellow

    Private Structure HeaderName
        Public mall As String
        Public codeName As String
        Public subcode1 As String
        Public subcode2 As String
    End Structure

    'チェック実行
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        If CheckBox10.Checked And DGV2.Rows.Count = 0 Then
            Exit Sub
        ElseIf CheckBox11.Checked And DGV4.Rows.Count = 0 Then
            Select Case ToolStripStatusLabel4.Text
                Case "楽天", "Yahoo"
                    MsgBox("下のリストは読み込まれていません", MsgBoxStyle.OkOnly And MsgBoxStyle.SystemModal)
                    Exit Sub
            End Select
        End If

        '初期設定
        ListBox5.Items.Clear()
        GroupBox2.Text = "一括データチェック(" & ListBox5.Items.Count & ")"
        dH2 = TM_HEADER_GET(DGV2)
        dH4 = TM_HEADER_GET(DGV4)
        dH5 = TM_HEADER_GET(DGV5)

        LIST4VIEW("店舗データ読出し", "s")
        Dim hnNo As Integer = 0
        Dim hnArray As New List(Of HeaderName)
        Dim hn As New HeaderName
        For i As Integer = 0 To TenpoTable.Rows.Count - 1
            If InStr(TenpoTable.Rows(i)("店名"), ToolStripComboBox2.Text) > 0 Then
                LIST4VIEW(TenpoTable.Rows(i)("モール") & "仕様チェック", "s")
                Select Case TenpoTable.Rows(i)("モール")
                    Case "楽天"
                        hn.mall = "楽天"
                        hn.codeName = "商品管理番号（商品URL）"
                        hn.subcode1 = "項目選択肢別在庫用横軸選択肢子番号"
                        hn.subcode2 = "項目選択肢別在庫用縦軸選択肢子番号"
                        hnArray.Add(hn)
                    Case "Yahoo"
                        hn.mall = "Yahoo"
                        hn.codeName = "code"
                        hn.subcode1 = "sub-code"
                        hn.subcode2 = ""
                        hnArray.Add(hn)
                    Case "MakeShop"
                        hn.mall = "MakeShop"
                        hn.codeName = "独自商品コード"
                        hn.subcode1 = ""
                        hn.subcode2 = ""
                        hnArray.Add(hn)
                    Case "Qoo10"
                        hn.mall = "Qoo10"
                        hn.codeName = "Seller Code"
                        hn.subcode1 = ""
                        hn.subcode2 = ""
                        hnArray.Add(hn)
                End Select
                Exit For
            End If
        Next

        'プログラムリスト取得
        Dim clbArray As New ArrayList
        For i As Integer = 0 To CheckedListBox1.Items.Count - 1
            clbArray.Add(CheckedListBox1.Items(i))
        Next

        'マスタコード
        masterCodeArray.Clear()
        masterDaihyoArray.Clear()
        For r As Integer = 0 To DGV5.RowCount - 1
            If Not masterCodeArray.Contains(DGV5.Item(dH5.IndexOf("商品コード"), r).Value.ToString.ToLower) Then
                masterCodeArray.Add(DGV5.Item(dH5.IndexOf("商品コード"), r).Value.ToString.ToLower)
            End If
            Dim dCode As String = DGV5.Item(dH5.IndexOf("代表商品コード"), r).Value
            If dCode <> "" Then
                If Not masterDaihyoArray.Contains(dCode.ToString.ToLower) Then
                    masterDaihyoArray.Add(dCode.ToString.ToLower)
                End If
            End If
        Next

        '--------------------------------------------
        'マスターコード存在チェック
        If CheckedListBox1.GetItemChecked(clbArray.IndexOf("マスターコード存在確認")) Then
            Master_Check1(hnArray, hnNo)
        End If
        '予約状態チェック
        If CheckedListBox1.GetItemChecked(clbArray.IndexOf("予約状態")) Then
            Yoyaku_Check(hnArray, hnNo)
        End If
        '予約文言記載チェック
        If CheckedListBox1.GetItemChecked(clbArray.IndexOf("予約文言記載")) Then
            Yoyaku_Mongon_Check(hnArray, hnNo)
        End If
        '販売期間・納品時期チェック
        If CheckedListBox1.GetItemChecked(clbArray.IndexOf("販売期間・納品時期")) Then
            SellDate_Check1(hnArray, hnNo)
        End If
        '配送方法チェック
        If CheckedListBox1.GetItemChecked(clbArray.IndexOf("配送方法確認")) Then
            Haisou_Check1(hnArray, hnNo)
        End If
        '--------------------------------------------

        MsgBox("終了しました", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
    End Sub

    'マスターコード存在チェック
    Private Sub Master_Check1(hnArray As List(Of HeaderName), hnNo As Integer)
        If CheckBox10.Checked Then
            dH2 = TM_HEADER_GET(DGV2)
            For r As Integer = 0 To DGV2.RowCount - 1
                Select Case hnArray(hnNo).mall
                    Case "楽天"
                        If DGV2.Item(dH2.IndexOf("在庫タイプ"), r).Value = "1" Then
                            Dim code As String = DGV2.Item(dH2.IndexOf(hnArray(hnNo).codeName), r).Value
                            code = code.ToString.ToLower
                            Master_Check2("i", hnArray(hnNo).codeName, code, DGV2, dH2, r)
                        End If
                    Case "Yahoo"
                        If DGV2.Item(dH2.IndexOf("sub-code"), r).Value = "" Then
                            Dim code As String = DGV2.Item(dH2.IndexOf(hnArray(hnNo).codeName), r).Value
                            code = code.ToString.ToLower
                            Master_Check2("i", hnArray(hnNo).codeName, code, DGV2, dH2, r)
                        End If
                End Select
            Next
        End If
        If CheckBox11.Checked Then
            Dim okCodeArray As New ArrayList
            Dim noMasterArray As New ArrayList
            dH4 = TM_HEADER_GET(DGV4)
            For r As Integer = 0 To DGV4.RowCount - 1
                Select Case hnArray(hnNo).mall
                    Case "楽天"
                        If DGV4.Item(dH4.IndexOf("選択肢タイプ"), r).Value = "i" Then
                            Dim code As String = DGV4.Item(dH4.IndexOf(hnArray(hnNo).codeName), r).Value & DGV4.Item(dH4.IndexOf(hnArray(hnNo).subcode1), r).Value & DGV4.Item(dH4.IndexOf(hnArray(hnNo).subcode2), r).Value
                            code = code.ToString.ToLower
                            Dim dCode As String = DGV4.Item(dH4.IndexOf(hnArray(hnNo).codeName), r).Value
                            If Master_Check2("s", hnArray(hnNo).codeName, code, DGV4, dH4, r) Then
                                If Not okCodeArray.Contains(dCode) Then
                                    okCodeArray.Add(dCode)   '代表コードを入れる
                                End If
                            Else
                                If Not okCodeArray.Contains(dCode) Then
                                    noMasterArray.Add(dCode & ",s" & "/" & dH4.IndexOf(hnArray(hnNo).codeName) & ":" & r & "/M無し" & "/" & code)
                                End If
                            End If
                        End If
                    Case "Yahoo"
                        If DGV4.Item(dH4.IndexOf("sub-code"), r).Value <> "" Then
                            Dim code As String = DGV4.Item(dH4.IndexOf(hnArray(hnNo).subcode1), r).Value
                            code = code.ToString.ToLower
                            Dim dCode As String = DGV4.Item(dH4.IndexOf(hnArray(hnNo).codeName), r).Value
                            If Master_Check2("s", hnArray(hnNo).codeName, code, DGV4, dH4, r) Then
                                If Not okCodeArray.Contains(dCode) Then
                                    okCodeArray.Add(dCode)   '代表コードを入れる
                                End If
                            Else
                                If Not okCodeArray.Contains(dCode) Then
                                    noMasterArray.Add(dCode & ",s" & "/" & dH4.IndexOf(hnArray(hnNo).codeName) & ":" & r & "/M無し" & "/" & code)
                                End If
                            End If
                        End If
                End Select
            Next

            '紐付けデータを除外する（未作成）

            For i As Integer = 0 To noMasterArray.Count - 1
                Dim str As String() = Split(noMasterArray(i), ",")
                If Not okCodeArray.Contains(str(0)) Then
                    If Not ListBox5.Items.Contains(str(1)) Then
                        ListBox5.Items.Add(str(1))
                    End If
                    GroupBox2.Text = "一括データチェック(" & ListBox5.Items.Count & ")"
                End If
            Next
        End If
    End Sub

    'マスターにコードがあるかチェック
    Private Function Master_Check2(f As String, codeName As String, code As String, dgv As DataGridView, dH As ArrayList, r As Integer) As Boolean
        If InStr(code, "sp") <> 0 Then
            Return True
        End If

        code = Master_Code_Convert(code)
        Mall_subwindow.DataGridView2.Rows.Add(code)
        If masterCodeArray.Contains(code) Then
            Return True
        Else
            dgv.Item(dH.IndexOf(codeName), r).Style.BackColor = errorColor
            'selectは他のサブコードを調べてから表示する
            If f = "i" Then
                Dim str As String = f & "/" & dH.IndexOf(codeName) & ":" & r & "/M無し" & "/" & code
                If Not ListBox5.Items.Contains(str) Then
                    ListBox5.Items.Add(str)
                    GroupBox2.Text = "一括データチェック(" & ListBox5.Items.Count & ")"
                End If
            End If
            Return False
        End If
    End Function

    ''' <summary>
    ''' マスターコードに変換する
    ''' </summary>
    ''' <param name="code"></param>
    ''' <param name="sub1"></param>
    ''' <param name="sub2"></param>
    ''' <returns></returns>
    Private Function Master_Code_Convert(code As String, Optional sub1 As String = "", Optional sub2 As String = "") As String
        Dim res As String = ""
        Dim codeA As String() = Nothing
        If sub1 <> "" And sub2 <> "" Then
            codeA = New String() {code.ToString.ToLower, sub1.ToString.ToLower, sub2.ToString.ToLower}
        ElseIf sub1 <> "" And sub2 = "" Then
            codeA = New String() {code.ToString.ToLower, sub1.ToString.ToLower}
        Else
            codeA = Split(code.ToString.ToLower, "-")
        End If

        If InStr(code, "ap022") > 0 Then
            Dim test As String = "0"
        End If

        Dim codeArray As New ArrayList
        If codeA.Length = 4 Then
            codeArray.Add(codeA(0) & "-" & codeA(1) & "-" & codeA(2) & "-" & codeA(3))
            codeArray.Add(codeA(0) & "-" & codeA(1) & "-" & codeA(3) & "-" & codeA(2))
            codeArray.Add(codeA(0) & "-" & codeA(2) & "-" & codeA(1) & "-" & codeA(3))
            codeArray.Add(codeA(0) & "-" & codeA(2) & "-" & codeA(3) & "-" & codeA(1))
            codeArray.Add(codeA(0) & "-" & codeA(3) & "-" & codeA(1) & "-" & codeA(2))
            codeArray.Add(codeA(0) & "-" & codeA(3) & "-" & codeA(2) & "-" & codeA(1))
        ElseIf codeA.Length = 3 Then
            Dim re = New Regex("-")
            codeA(1) = re.Replace(codeA(1), "", 1)   '「-」が2個入っている場合に対応（1回だけ置換）
            codeA(2) = Replace(codeA(2), "-", "")
            codeArray.Add(codeA(0) & codeA(1) & codeA(2))
            codeArray.Add(codeA(0) & "-" & codeA(1) & codeA(2))
            codeArray.Add(codeA(0) & codeA(1) & "-" & codeA(2))
            codeArray.Add(codeA(0) & "-" & codeA(1) & "-" & codeA(2))
            codeArray.Add(codeA(0) & codeA(2) & codeA(1))
            codeArray.Add(codeA(0) & "-" & codeA(2) & codeA(1))
            codeArray.Add(codeA(0) & codeA(2) & "-" & codeA(1))
            codeArray.Add(codeA(0) & "-" & codeA(2) & "-" & codeA(1))
        ElseIf codeA.Length = 2 Then
            codeArray.Add(codeA(0) & "-" & codeA(1))
            codeArray.Add(codeA(0) & Replace(codeA(1), "-", ""))
        Else
            codeArray.Add(code.ToString.ToLower & sub1.ToString.ToLower & sub2.ToString.ToLower)
        End If

        For i As Integer = 0 To codeArray.Count - 1
            If masterCodeArray.Contains(codeArray(i)) Then
                res = codeArray(i)
                Exit For
            End If
        Next

        Return res
    End Function

    '予約状態のチェック
    Private Sub Yoyaku_Check(hnArray As List(Of HeaderName), hnNo As Integer)
        If CheckBox10.Checked Then
            LIST4VIEW("予約リストチェック", "s")

            Dim soukoArray As New ArrayList
            Dim dgv As DataGridView() = New DataGridView() {DGV2, DGV4}
            For i As Integer = 0 To 1
                If dgv(i).Rows.Count = 0 Then
                    Continue For
                End If
                Dim dH As ArrayList = TM_HEADER_GET(dgv(i))

                ToolStripProgressBar1.Value = 0
                ToolStripProgressBar1.Maximum = dgv(i).RowCount

                'https://navi-manual.faq.rakuten.net/item/000009624
                '倉庫指定 0： 販売中 倉庫に入れる 空欄の場合は0（販売中）
                '在庫タイプ 1： 通常在庫設定 項目選択肢別在庫設定
                '選択肢タイプ:　s：セレクトボックス　c：チェックボックス　f：フリーテキスト i：項目選択肢別在庫
                For r As Integer = 0 To dgv(i).RowCount - 1
                    Dim code As String = dgv(i).Item(dH.IndexOf(hnArray(hnNo).codeName), r).Value
                    code = code.ToString.ToLower
                    Select Case hnArray(hnNo).mall
                        Case "楽天"
                            If i = 0 Then
                                If Not soukoArray.Contains(code) Then
                                    If dgv(i).Item(dH.IndexOf("倉庫指定"), r).Value = "0" Then
                                        If dgv(i).Item(dH.IndexOf("在庫タイプ"), r).Value = "1" Then  'サブコード無しだけチェック
                                            Yoyaku_Check1(hnArray(hnNo).mall, "i", code, dgv(i), dH, r)
                                        End If
                                    Else
                                        soukoArray.Add(code)
                                    End If
                                End If
                            Else
                                If Not soukoArray.Contains(code) Then
                                    If dgv(i).Item(dH.IndexOf("選択肢タイプ"), r).Value = "i" Then  'サブコード有りだけチェック
                                        code &= dgv(i).Item(dH.IndexOf(hnArray(hnNo).subcode1), r).Value
                                        code &= dgv(i).Item(dH.IndexOf(hnArray(hnNo).subcode2), r).Value
                                        code = code.ToLower
                                        Yoyaku_Check1(hnArray(hnNo).mall, "s", code, dgv(i), dH, r)
                                    End If
                                End If
                            End If
                        Case "Yahoo"
                            If i = 0 Then
                                If Not soukoArray.Contains(code) Then
                                    If dgv(i).Item(dH.IndexOf("display"), r).Value = "0" Then
                                        If dgv(i).Item(dH.IndexOf("sub-code"), r).Value = "" Then
                                            Yoyaku_Check1(hnArray(hnNo).mall, "i", code, dgv(i), dH, r)
                                        End If
                                    Else
                                        soukoArray.Add(code)
                                    End If
                                End If
                            Else
                                If Not soukoArray.Contains(code) Then
                                    If dgv(i).Item(dH.IndexOf("sub-code"), r).Value <> "" Then
                                        Yoyaku_Check1(hnArray(hnNo).mall, "i", code, dgv(i), dH, r)
                                    End If
                                End If
                            End If
                        Case "MakeShop"
                            If i = 0 Then
                                If Not soukoArray.Contains(code) Then
                                    If dgv(i).Item(dH.IndexOf("商品表示可否"), r).Value = "Y" Then
                                        Yoyaku_Check1(hnArray(hnNo).mall, "i", code, dgv(i), dH, r)
                                    Else
                                        soukoArray.Add(code)
                                    End If
                                End If
                            Else

                            End If
                        Case "Qoo10"
                            If i = 0 Then
                                Yoyaku_Check1(hnArray(hnNo).mall, "i", code, dgv(i), dH, r)
                            Else

                            End If
                    End Select

                    ToolStripProgressBar1.Value += 1
                    Application.DoEvents()
                Next
            Next

            ToolStripProgressBar1.Value = 0
            LIST4VIEW("予約リストチェック", "e")
        End If
    End Sub

    '予約管理番号チェック
    Private Function Yoyaku_Check1(mall As String, f As String, code As String, dgv As DataGridView, dH As ArrayList, r As Integer)
        If Not masterCodeArray.Contains(code.ToLower) Then
            Return True
        End If

        Dim kubun As String = DGV5.Item(dH5.IndexOf("商品区分"), masterCodeArray.IndexOf(code.ToLower)).Value
        Dim leadTimeName As String = ""
        Dim No As String = ""
        If mall = "楽天" Then
            leadTimeName = "在庫あり時納期管理番号"
            'No = "1"
        ElseIf mall = "Yahoo" Then
            leadTimeName = "lead-time-instock"
            No = "2"
        ElseIf mall = "MakeShop" Then
            leadTimeName = "製造元"
            No = ""     '即納時文字
        ElseIf mall = "Qoo10" Then
            leadTimeName = "Available Date"
            No = ""     '即納時文字
        End If
        If code = "zk285-a" Then
            Dim a = ""
        End If
        Dim res As Boolean = True
        If mall = "楽天" Then
            Select Case kubun
                Case "通常"
                    If dgv.Item(dH.IndexOf(leadTimeName), r).Value = "1" Or dgv.Item(dH.IndexOf(leadTimeName), r).Value = "4" Then
                    Else
                        If ToolStripComboBox2.Text = "FKstyle" Then
                            If dgv.Item(dH.IndexOf(leadTimeName), r).Value = "1000" Then
                                Return True
                                Exit Function
                            End If
                        End If

                        dgv.Item(dH.IndexOf(leadTimeName), r).Style.BackColor = errorColor
                        Dim str As String = f & "/" & dH.IndexOf(leadTimeName) & ":" & r & " / 即納になってない"
                        If Not ListBox5.Items.Contains(str) Then
                            ListBox5.Items.Add(str)
                            GroupBox2.Text = "一括データチェック(" & ListBox5.Items.Count & ")"
                        End If
                        res = False
                    End If
                Case "予約"
                    If dgv.Item(dH.IndexOf(leadTimeName), r).Value = "1" Or dgv.Item(dH.IndexOf(leadTimeName), r).Value = "3" Or dgv.Item(dH.IndexOf(leadTimeName), r).Value = "4" Then
                        dgv.Item(dH.IndexOf(leadTimeName), r).Style.BackColor = errorColor
                        Dim str As String = f & "/" & dH.IndexOf(leadTimeName) & ":" & r & " / 予約になってない"
                        If Not ListBox5.Items.Contains(str) Then
                            ListBox5.Items.Add(str)
                            GroupBox2.Text = "一括データチェック(" & ListBox5.Items.Count & ")"
                        End If
                        res = False
                    End If
            End Select
        Else
            Select Case kubun
                Case "通常"
                    If dgv.Item(dH.IndexOf(leadTimeName), r).Value <> No Then
                        If ToolStripComboBox2.Text = "FKstyle" Then
                            If dgv.Item(dH.IndexOf(leadTimeName), r).Value = "1000" Then
                                Return True
                                Exit Function
                            End If
                        End If

                        dgv.Item(dH.IndexOf(leadTimeName), r).Style.BackColor = errorColor
                        Dim str As String = f & "/" & dH.IndexOf(leadTimeName) & ":" & r & " / 即納になってない"
                        If Not ListBox5.Items.Contains(str) Then
                            ListBox5.Items.Add(str)
                            GroupBox2.Text = "一括データチェック(" & ListBox5.Items.Count & ")"
                        End If
                        res = False
                    End If
                Case "予約"
                    If dgv.Item(dH.IndexOf(leadTimeName), r).Value = No Then
                        dgv.Item(dH.IndexOf(leadTimeName), r).Style.BackColor = errorColor
                        Dim str As String = f & "/" & dH.IndexOf(leadTimeName) & ":" & r & " / 予約になってない"
                        If Not ListBox5.Items.Contains(str) Then
                            ListBox5.Items.Add(str)
                            GroupBox2.Text = "一括データチェック(" & ListBox5.Items.Count & ")"
                        End If
                        res = False
                    End If
            End Select
        End If
        Return res
    End Function

    '予約文言チェック
    Private Sub Yoyaku_Mongon_Check(hnArray As List(Of HeaderName), hnNo As Integer)
        If CheckBox10.Checked Then
            LIST4VIEW("予約文言リストチェック", "s")

            Dim dgv As DataGridView() = New DataGridView() {DGV2, DGV4}
            For i As Integer = 0 To 1
                If dgv(i).Rows.Count = 0 Then
                    Continue For
                End If
                Dim dH As ArrayList = TM_HEADER_GET(dgv(i))
                Dim checkArray As New ArrayList

                ToolStripProgressBar1.Value = 0
                ToolStripProgressBar1.Maximum = dgv(i).RowCount

                For r As Integer = 0 To dgv(i).RowCount - 1
                    Dim code As String = dgv(i).Item(dH.IndexOf(hnArray(hnNo).codeName), r).Value
                    code = code.ToString.ToLower
                    Select Case hnArray(hnNo).mall
                        Case "楽天"
                            If i = 0 Then
                                '楽天itemは検索該当無し
                            Else
                                If dgv(i).Item(dH.IndexOf("選択肢タイプ"), r).Value = "s" Then  'サブコード有りだけチェック
                                    Dim dCode As String = dgv(i).Item(dH.IndexOf("商品管理番号（商品URL）"), r).Value
                                    If Not checkArray.Contains(dCode) Then
                                        Dim TenpoRow As DataRow() = DgvTable(i).Select("商品管理番号（商品URL）='" & dCode & "' AND 選択肢タイプ='s'")
                                        Yoyaku_Check2("s", hnArray(hnNo).mall, TenpoRow, dCode, dgv(i), dH, r)
                                        checkArray.Add(dCode)
                                    End If
                                Else
                                    Dim dCode As String = dgv(i).Item(dH.IndexOf("商品管理番号（商品URL）"), r).Value.ToString.ToLower
                                    Dim sCode As String = dgv(i).Item(dH.IndexOf("商品管理番号（商品URL）"), r).Value.ToString.ToLower
                                    If CStr(dgv(i).Item(dH.IndexOf("項目選択肢別在庫用横軸選択肢子番号"), r).Value) <> "" Then
                                        sCode &= dgv(i).Item(dH.IndexOf("項目選択肢別在庫用横軸選択肢子番号"), r).Value.ToString.ToLower
                                    End If
                                    If CStr(dgv(i).Item(dH.IndexOf("項目選択肢別在庫用縦軸選択肢子番号"), r).Value) <> "" Then
                                        sCode &= dgv(i).Item(dH.IndexOf("項目選択肢別在庫用縦軸選択肢子番号"), r).Value.ToString.ToLower
                                    End If
                                    If Not checkArray.Contains(sCode) Then
                                        'Dim TenpoRow As DataRow() = DgvTable(i).Select("商品管理番号（商品URL）='" & dCode & "' AND 選択肢タイプ='i'")
                                        Yoyaku_Check1(hnArray(hnNo).mall, "s", sCode, dgv(i), dH, r)
                                        checkArray.Add(sCode)
                                    End If
                                End If
                            End If
                        Case "Yahoo"
                            If i = 0 Then
                                If dgv(i).Item(dH.IndexOf("sub-code"), r).Value = "" Then
                                    Dim dCode As String = dgv(i).Item(dH.IndexOf("code"), r).Value
                                    Dim TenpoRow As DataRow() = DgvTable(i).Select("[code]='" & dCode & "'") ' AND [sub-code]=''")
                                    Yoyaku_Check2("i", hnArray(hnNo).mall, TenpoRow, dCode, dgv(i), dH, r)
                                End If
                            Else
                                If dgv(i).Item(dH.IndexOf("sub-code"), r).Value <> "" Then
                                    Dim sCode As String = dgv(i).Item(dH.IndexOf("sub-code"), r).Value
                                    If Trim(sCode) <> "" Then
                                        Dim TenpoRow As DataRow() = DgvTable(i).Select("[sub-code]='" & sCode & "'")
                                        Yoyaku_Check2("s", hnArray(hnNo).mall, TenpoRow, sCode, dgv(i), dH, r)
                                    End If
                                End If
                            End If
                        Case "MakeShop"
                            If i = 0 Then
                                Dim dCode As String = dgv(i).Item(dH.IndexOf("独自商品コード"), r).Value
                                Dim TenpoRow As DataRow() = DgvTable(i).Select("[独自商品コード]='" & dCode & "'") ' AND [sub-code]=''")
                                Yoyaku_Check2("i", hnArray(hnNo).mall, TenpoRow, dCode, dgv(i), dH, r)
                            Else

                            End If
                        Case "Qoo10"
                            If i = 0 Then
                                Dim dCode As String = dgv(i).Item(dH.IndexOf("Seller Code"), r).Value
                                Dim TenpoRow As DataRow() = DgvTable(i).Select("[Seller Code]='" & dCode & "'") ' AND [sub-code]=''")
                                Yoyaku_Check2("i", hnArray(hnNo).mall, TenpoRow, dCode, dgv(i), dH, r)
                            Else

                            End If
                    End Select

                    ToolStripProgressBar1.Value += 1
                Next
            Next

            ToolStripProgressBar1.Value = 0
            LIST4VIEW("予約文言リストチェック", "e")
        End If
    End Sub

    'selectの予約文章チェック
    Private Sub Yoyaku_Check2(f As String, mall As String, TableRow As DataRow(), code As String, dgv As DataGridView, dH As ArrayList, r As Integer)
        Dim kubun As String = "通常"
        Dim MasterRow As DataRow() = dgv5table.Select("商品コード='" & code & "' or 代表商品コード='" & code & "'")
        For i As Integer = 0 To MasterRow.Length - 1
            If MasterRow(i)("商品区分").ToString = "予約" Then
                kubun = "予約"
                Exit For
            End If
        Next

        Dim koumokuName As String = ""
        If mall = "楽天" Then
            'koumokuName = "Select/Checkbox用項目名"
            koumokuName = "項目選択肢項目名"
        ElseIf mall = "Yahoo" Then
            If f = "i" Then
                koumokuName = "options"
            Else
                koumokuName = "option-name-1"
            End If
        ElseIf mall = "MakeShop" Then
            koumokuName = "商品別特殊表示"
        ElseIf mall = "Qoo10" Then
            koumokuName = "Option Info"
        End If

        Dim res As Boolean = True
        Select Case kubun
            Case "通常"
                For i As Integer = 0 To TableRow.Count - 1
                    If InStr(TableRow(i)(koumokuName).ToString, "予約") > 0 Then
                        Dim str As String = f & "/" & dH.IndexOf(koumokuName) & ":" & r + i & " / 予約文言削除漏れ"
                        If Not ListBox5.Items.Contains(str) Then
                            ListBox5.Items.Add(str)
                            GroupBox2.Text = "一括データチェック(" & ListBox5.Items.Count & ")"
                        End If
                        Syori_Hyouji(r) '処理行を表示
                        Exit For
                    End If
                Next
            Case "予約"
                Dim yFlag As Boolean = False
                For i As Integer = 0 To TableRow.Count - 1
                    If InStr(TableRow(i)(koumokuName).ToString, "予約") > 0 Then
                        yFlag = True
                    End If
                Next
                If yFlag = False Then
                    Dim str As String = f & "/" & dH.IndexOf(koumokuName) & ":" & r & " / 予約文言ない"
                    If Not ListBox5.Items.Contains(str) Then
                        ListBox5.Items.Add(str)
                        GroupBox2.Text = "一括データチェック(" & ListBox5.Items.Count & ")"
                    End If
                    Syori_Hyouji(r) '処理行を表示
                End If
        End Select

        'For i As Integer = 0 To TableRow.Count - 1
        '    Select Case kubun
        '        Case "通常"
        '            If InStr(TableRow(i)(koumokuName).ToString, "予約") > 0 Then
        '                If mall = "楽天" Then
        '                    ListBox5.Items.Add("s/" & dH.IndexOf(koumokuName) & ":" & r & " / 予約文言削除漏れ")
        '                End If
        '                res = False
        '            End If
        '        Case "予約"
        '            If InStr(TableRow(i)(koumokuName).ToString, "予約") > 0 Then
        '                res = True
        '            End If
        '    End Select
        'Next
        'If kubun = "予約" Then
        '    If mall = "楽天" Then
        '        ListBox5.Items.Add("s/" & dH.IndexOf(koumokuName) & ":" & r & " / 予約文言ない")
        '    End If
        '    res = False
        'End If
    End Sub

    '販売期間チェック
    Private Sub SellDate_Check1(hnArray As List(Of HeaderName), hnNo As Integer)
        If CheckBox10.Checked Then
            dH2 = TM_HEADER_GET(DGV2)
            ToolStripProgressBar1.Value = 0
            ToolStripProgressBar1.Maximum = DGV2.RowCount

            For r As Integer = 0 To DGV2.RowCount - 1
                Select Case hnArray(hnNo).mall
                    Case "楽天"
                        If DGV2.Item(dH2.IndexOf("倉庫指定"), r).Value = "0" Then
                            SellDate_Check2("楽天", DGV2, dH2, r)
                        End If
                    Case "Yahoo"
                        If DGV2.Item(dH2.IndexOf("display"), r).Value = "1" Then
                            SellDate_Check2("Yahoo", DGV2, dH2, r)
                        End If
                    Case "MakeShop"
                        If DGV2.Item(dH2.IndexOf("商品表示可否"), r).Value = "Y" Then
                            SellDate_Check2("MakeShop", DGV2, dH2, r)
                        End If
                End Select
                ToolStripProgressBar1.Value += 1
            Next

            ToolStripProgressBar1.Value = 0
        End If
    End Sub

    Private Sub SellDate_Check2(mall As String, dgv As DataGridView, dH As ArrayList, r As Integer)
        Dim koumokuNameArray As String() = Nothing
        If mall = "楽天" Then       '2018/09/04 20:00 2018/09/11 01:59
            koumokuNameArray = {"販売期間指定"}
            If dgv.Item(dH.IndexOf(koumokuNameArray(0)), r).Value <> "" Then
                Dim str As String() = Split(dgv.Item(dH.IndexOf(koumokuNameArray(0)), r).Value, " ")
                Dim d As Date = CDate(str(2) & " " & str(3) & ":00")
                If d < Now Then
                    Dim str2 As String = "i" & "/" & dH.IndexOf(koumokuNameArray(0)) & ":" & r & " / 期間超過"
                    If Not ListBox5.Items.Contains(str2) Then
                        ListBox5.Items.Add(str2)
                        GroupBox2.Text = "一括データチェック(" & ListBox5.Items.Count & ")"
                    End If
                    Syori_Hyouji(r) '処理行を表示
                End If
            End If
        ElseIf mall = "Yahoo" Then  '2018081700
            koumokuNameArray = {"sale-period-start", "sale-period-end"}
            If dgv.Item(dH.IndexOf(koumokuNameArray(1)), r).Value <> "" Then
                Dim str As String = dgv.Item(dH.IndexOf(koumokuNameArray(1)), r).Value
                Dim d As Date = CDate(str.Substring(0, 4) & "/" & str.Substring(4, 2) & "/" & str.Substring(6, 2) & " " & str.Substring(8, 2) & ":59:59")
                If d < Now Then
                    Dim str2 As String = "i" & "/" & dH.IndexOf(koumokuNameArray(1)) & ":" & r & " / 期間超過"
                    If Not ListBox5.Items.Contains(str2) Then
                        ListBox5.Items.Add(str2)
                        GroupBox2.Text = "一括データチェック(" & ListBox5.Items.Count & ")"
                    End If
                    Syori_Hyouji(r) '処理行を表示
                End If
            End If
        ElseIf mall = "MakeShop" Then   '20180825-20180825
            koumokuNameArray = {"割引期間"}
            If dgv.Item(dH.IndexOf(koumokuNameArray(0)), r).Value <> "" Then
                Dim str As String() = Split(dgv.Item(dH.IndexOf(koumokuNameArray(0)), r).Value, "-")
                Dim d As Date = CDate(str(1).Substring(0, 4) & "/" & str(1).Substring(4, 2) & "/" & str(1).Substring(6, 2) & "23:59:59")
                If d < Now Then
                    Dim str2 As String = "i" & "/" & dH.IndexOf(koumokuNameArray(1)) & ":" & r & " / 期間超過"
                    If Not ListBox5.Items.Contains(str2) Then
                        ListBox5.Items.Add(str)
                        GroupBox2.Text = "一括データチェック(" & ListBox5.Items.Count & ")"
                    End If
                    Syori_Hyouji(r) '処理行を表示
                End If
            End If
        End If
    End Sub

    '配送方法チェック
    Private Sub Haisou_Check1(hnArray As List(Of HeaderName), hnNo As Integer)
        If CheckBox10.Checked Then
            dH2 = TM_HEADER_GET(DGV2)
            ToolStripProgressBar1.Value = 0
            ToolStripProgressBar1.Maximum = DGV2.RowCount

            For r As Integer = 0 To DGV2.RowCount - 1
                Select Case hnArray(hnNo).mall
                    Case "楽天"
                        If DGV2.Item(dH2.IndexOf("倉庫指定"), r).Value = "0" Then
                            Dim dCode As String = DGV2.Item(dH2.IndexOf("商品管理番号（商品URL）"), r).Value
                            Haisou_Check2("楽天", DGV2, dH2, r, dCode)
                        End If
                    Case "MakeShop"
                        If DGV2.Item(dH2.IndexOf("商品表示可否"), r).Value = "Y" Then
                            Dim dCode As String = DGV2.Item(dH2.IndexOf("独自商品コード"), r).Value
                            Haisou_Check2("MakeShop", DGV2, dH2, r, dCode)
                        End If
                End Select
                ToolStripProgressBar1.Value += 1
            Next

            ToolStripProgressBar1.Value = 0
        End If
    End Sub

    Private Sub Haisou_Check2(mall As String, dgv As DataGridView, dH As ArrayList, r As Integer, code As String)
        Dim kubun As String() = Nothing
        Dim MasterRow As DataRow() = dgv5table.Select("商品コード='" & code & "' or 代表商品コード='" & code & "'")
        If MasterRow.Length = 0 Then
            Exit Sub
        End If

        For i As Integer = 0 To MasterRow.Length - 1
            If InStr(MasterRow(i)("商品分類タグ").ToString, "宅配便") Then
                If InStr(ToolStripComboBox2.SelectedItem, "暁") > 0 Then
                    kubun = {"6", "宅配便"}
                Else
                    kubun = {"2|3", "宅配便"}
                End If
                Exit For
            ElseIf InStr(MasterRow(i)("商品分類タグ").ToString, "メール便") Then
                If InStr(ToolStripComboBox2.SelectedItem, "暁") > 0 Then
                    kubun = {"7", "メール便"}
                Else
                    kubun = {"1", "メール便"}
                End If
                Exit For
            ElseIf InStr(MasterRow(i)("商品分類タグ").ToString, "定形外") Then
                If InStr(ToolStripComboBox2.SelectedItem, "暁") > 0 Then
                    kubun = {"8", "定形外"}
                Else
                    kubun = {"4", "定形外"}
                End If
                Exit For
            Else
                Exit Sub
            End If
        Next

        Dim koumokuName As String = ""
        If mall = "楽天" Then
            koumokuName = "配送方法セット管理番号"
            If dgv.Item(dH.IndexOf(koumokuName), r).Value <> "" Then
                Dim str As String = dgv.Item(dH.IndexOf(koumokuName), r).Value
                If Not Regex.IsMatch(str, kubun(0)) Then
                    Dim str2 As String = "i" & "/" & dH.IndexOf(koumokuName) & ":" & r & " / 配送相違(" & kubun(1) & ")に"
                    If Not ListBox5.Items.Contains(str2) Then
                        ListBox5.Items.Add(str2)
                        GroupBox2.Text = "一括データチェック(" & ListBox5.Items.Count & ")"
                    End If
                    Syori_Hyouji(r) '処理行を表示
                End If
            Else
                Dim str2 As String = "i" & "/" & dH.IndexOf(koumokuName) & ":" & r & " / 配送未設定"
                If Not ListBox5.Items.Contains(str2) Then
                    ListBox5.Items.Add(str2)
                    GroupBox2.Text = "一括データチェック(" & ListBox5.Items.Count & ")"
                End If
                Syori_Hyouji(r) '処理行を表示
            End If
        ElseIf mall = "MakeShop" Then
            koumokuName = "送料個別設定"
            If dgv.Item(dH.IndexOf(koumokuName), r).Value <> "" Then
                Dim str As String = dgv.Item(dH.IndexOf(koumokuName), r).Value
                Select Case str
                    Case "20170407121342", "20170407121633"
                        str = "宅配便"
                    Case "20170407121422"
                        str = "メール便"
                    Case "20180402101403"
                        str = "定形外"
                End Select
                If str <> kubun(1) Then
                    Dim str2 As String = "i" & "/" & dH.IndexOf(koumokuName) & ":" & r & " / 配送相違(" & kubun(1) & ")に"
                    If Not ListBox5.Items.Contains(str2) Then
                        ListBox5.Items.Add(str2)
                        GroupBox2.Text = "一括データチェック(" & ListBox5.Items.Count & ")"
                    End If
                    Syori_Hyouji(r) '処理行を表示
                End If
            End If
        End If
    End Sub

    Private Sub Syori_Hyouji(r As Integer)
        ToolStripStatusLabel2.Text = "[処理：" & String.Format("{0:D4}", r) & "]"
        Application.DoEvents()
    End Sub
    '************************************************************************************************************




    Private Sub Data_Check(mall As String)
        ListBox5.Items.Clear()
        GroupBox2.Text = "一括データチェック(" & ListBox5.Items.Count & ")"
        dH2 = TM_HEADER_GET(DGV2)
        dH4 = TM_HEADER_GET(DGV4)
        dH5 = TM_HEADER_GET(DGV5)

        Dim codeName As String = ""
        Select Case mall
            Case "楽天"
                codeName = "商品管理番号（商品URL）"
            Case "Yahoo"
                codeName = "code"
        End Select

        'マスタコード
        For r As Integer = 0 To DGV5.RowCount - 1
            masterCodeArray.Add(DGV5.Item(dH5.IndexOf("商品コード"), r).Value.ToString.ToLower)
            If DGV5.Item(dH5.IndexOf("代表商品コード"), r).Value <> "" Then
                'dCodeArray.Add(dgv5.Item(dH5.IndexOf("代表商品コード"), r).Value.ToString.ToLower)
            End If
        Next

        ToolStripProgressBar1.Value = 0
        ToolStripProgressBar1.Maximum = DGV2.RowCount

        'チェック
        Dim code As String = ""
        For r As Integer = 0 To DGV2.RowCount - 1
            'マスタチェック
            Select Case mall
                Case "楽天"
                    code = DGV2.Item(dH2.IndexOf(codeName), r).Value
                    If InStr(code, "sp") = 0 Then   'spはセット商品
                        If Not Master_Check2(mall, "i", code.ToLower, DGV2, dH2, r) Then
                            Continue For
                        End If
                        '予約状態
                        If DGV2.Item(dH2.IndexOf("在庫タイプ"), r).Value = "1" Then  'サブコード無しだけチェック
                            Yoyaku_Check1(mall, "i", code.ToLower, DGV2, dH2, r)
                        End If
                    End If
                Case "Yahoo"
                    '予約状態（サブコードは調べない）
                    If DGV2.Item(dH2.IndexOf("sub-code"), r).Value = "" Then
                        code = DGV2.Item(dH2.IndexOf(codeName), r).Value
                        Yoyaku_Check1(mall, "i", code, DGV2, dH2, r)
                    End If
            End Select

            ToolStripProgressBar1.Value += 1
        Next

        ToolStripProgressBar1.Value = 0
        ToolStripProgressBar1.Maximum = DGV4.RowCount

        Select Case mall
            Case "楽天"
                For r As Integer = 0 To DGV4.RowCount - 1
                    If DGV4.Item(dH4.IndexOf("選択肢タイプ"), r).Value = "i" Then
                        code = DGV4.Item(dH4.IndexOf("商品管理番号（商品URL）"), r).Value &
                            DGV4.Item(dH4.IndexOf("項目選択肢別在庫用横軸選択肢子番号"), r).Value &
                            DGV4.Item(dH4.IndexOf("項目選択肢別在庫用縦軸選択肢子番号"), r).Value

                        'マスタチェック
                        If Not Master_Check2(mall, "s", code.ToLower, DGV4, dH4, r) Then
                            Continue For
                        End If
                    Else
                        '予約状態
                        Dim dCode As String = DGV4.Item(dH4.IndexOf("商品管理番号（商品URL）"), r).Value
                        Dim TenpoRow As DataRow() = DgvTable(1).Select("商品管理番号（商品URL）='" & dCode & "' AND 選択肢タイプ='s'")
                        Dim codeArray As New ArrayList From {
                            DGV4.Item(dH4.IndexOf("商品管理番号（商品URL）"), r).Value & DGV4.Item(dH4.IndexOf("項目選択肢別在庫用横軸選択肢子番号"), r).Value & DGV4.Item(dH4.IndexOf("項目選択肢別在庫用縦軸選択肢子番号"), r).Value,
                            DGV4.Item(dH4.IndexOf("商品管理番号（商品URL）"), r).Value & Replace(DGV4.Item(dH4.IndexOf("項目選択肢別在庫用横軸選択肢子番号"), r).Value, "-", "") & DGV4.Item(dH4.IndexOf("項目選択肢別在庫用縦軸選択肢子番号"), r).Value,
                            DGV4.Item(dH4.IndexOf("商品管理番号（商品URL）"), r).Value & DGV4.Item(dH4.IndexOf("項目選択肢別在庫用横軸選択肢子番号"), r).Value & Replace(DGV4.Item(dH4.IndexOf("項目選択肢別在庫用縦軸選択肢子番号"), r).Value, "-", ""),
                            DGV4.Item(dH4.IndexOf("商品管理番号（商品URL）"), r).Value & DGV4.Item(dH4.IndexOf("項目選択肢別在庫用縦軸選択肢子番号"), r).Value & DGV4.Item(dH4.IndexOf("項目選択肢別在庫用横軸選択肢子番号"), r).Value,
                            DGV4.Item(dH4.IndexOf("商品管理番号（商品URL）"), r).Value & Replace(DGV4.Item(dH4.IndexOf("項目選択肢別在庫用縦軸選択肢子番号"), r).Value, "-", "") & DGV4.Item(dH4.IndexOf("項目選択肢別在庫用横軸選択肢子番号"), r).Value,
                            DGV4.Item(dH4.IndexOf("商品管理番号（商品URL）"), r).Value & DGV4.Item(dH4.IndexOf("項目選択肢別在庫用縦軸選択肢子番号"), r).Value & Replace(DGV4.Item(dH4.IndexOf("項目選択肢別在庫用横軸選択肢子番号"), r).Value, "-", "")
                        }
                        Dim xCode As String = ""
                        For i As Integer = 0 To codeArray.Count - 1
                            If masterCodeArray.Contains(codeArray(i).ToString.ToLower) Then
                                xCode = codeArray(i).ToString.ToLower
                                Exit For
                            End If
                        Next
                        If xCode <> "" Then
                            Dim kubun As String = DGV5.Item(dH5.IndexOf("商品区分"), masterCodeArray.IndexOf(xCode)).Value
                            'Yoyaku_Check2("楽天", TenpoRow, xCode, dgv4, dH4, r)
                        End If
                    End If

                    ToolStripProgressBar1.Value += 1
                Next
            Case "Yahoo"
                Dim xCode As String = ""
                Dim yoyakuArray As New ArrayList
                For r As Integer = 0 To DGV4.RowCount - 1
                    Dim dCode As String = DGV4.Item(dH4.IndexOf("code"), r).Value
                    Dim TenpoRow As DataRow() = DgvTable(1).Select("code='" & dCode & "' AND sub-code=''")

                    If DGV2.Item(dH2.IndexOf("sub-code"), r).Value = "" Then
                        'Yoyaku_Check2("楽天", TenpoRow, xCode, dgv4, dH4, r)
                    Else
                        'Yoyaku_Check2("楽天", TenpoRow, xCode, dgv4, dH4, r)
                    End If
                Next
        End Select

        ToolStripProgressBar1.Value = 0
        ToolStripProgressBar1.Maximum = DGV4.RowCount

        Select Case mall
            Case "楽天"

            Case "Yahoo"

        End Select

        ToolStripProgressBar1.Value = 0
    End Sub

    Private Sub ListBox5_DoubleClick(sender As Object, e As EventArgs) Handles ListBox5.DoubleClick
        If ListBox5.SelectedIndices.Count > 0 Then
            Dim selStr As String() = Split(ListBox5.SelectedItem, "/")
            Dim cr As String() = Split(selStr(1), ":")
            If selStr(0) = "i" Then
                DGV2.Item(CInt(cr(0)), CInt(cr(1))).Selected = True
                DGV2.CurrentCell = DGV2(CInt(cr(0)), CInt(cr(1)))
            Else
                DGV4.Item(CInt(cr(0)), CInt(cr(1))).Selected = True
                DGV4.CurrentCell = DGV4(CInt(cr(0)), CInt(cr(1)))
            End If
        End If
    End Sub

    'グリッド表示変更
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        If ComboBox5.SelectedItem = "" Then
            Exit Sub
        End If

        If ComboBox5.SelectedItem = "全て" Then
            UndoDataDGV()
        Else
            LIST4VIEW("設定読出し", "s")
            where = "[名称] = 'csv列表示変更'"
            where &= " AND [名称2] = '" & ComboBox5.SelectedItem & "'"
            TableName = "T_プログラム"
            '=================================================
            Dim Table As DataTable = MALLDB_CONNECT_SELECT(CnAccdb, TableName, where)
            '=================================================

            Dim pStrArray1 As String() = Split(Table.Rows(0)("プログラム").ToString, "|=|")
            Dim dgv As DataGridView = DGV2
            Dim dH As ArrayList = TM_HEADER_GET(dgv)
            Dim headerFlag As Boolean = False
            For i As Integer = 0 To pStrArray1.Length - 1
                Select Case pStrArray1(i)
                    Case "1"
                        dgv = DGV2
                    Case "2"
                        If SplitContainer7.Panel2Collapsed Then
                            Exit For
                        Else
                            dgv = DGV4
                        End If
                    Case "header"
                        headerFlag = True
                    Case Else
                        If headerFlag Then
                            Dim columnList As String() = Split(pStrArray1(i), "|")
                            For c As Integer = 0 To columnList.Length - 1
                                If InStr(columnList(c), "#") > 0 Then
                                    Dim cl As String() = Split(columnList(c), "#")
                                    If dH.Contains(cl(0)) Then  '既に存在する列はヘッダーを書き換え
                                        dgv.Columns(c).HeaderText = cl(1)
                                    Else    '存在しない列は追加
                                        dgv.Columns.Add(columnList(c), cl(1))
                                    End If
                                Else
                                    dgv.Columns(c).HeaderText = columnList(c)
                                End If
                            Next
                        Else
                            Dim columnList As String() = Split(pStrArray1(i), "|")
                            Select Case columnList(0)
                                Case "＜非表示＞"       '非表示
                                    For c As Integer = dgv.ColumnCount - 1 To 0 Step -1
                                        If Not columnList.Contains(dgv.Columns(c).HeaderText) Then
                                            dgv.Columns(c).Visible = False
                                        Else
                                            dgv.Columns(c).Visible = True
                                        End If
                                    Next
                                Case Else               '削除
                                    For c As Integer = dgv.ColumnCount - 1 To 0 Step -1
                                        If Not columnList.Contains(dgv.Columns(c).HeaderText) Then
                                            dgv.Columns.RemoveAt(c)
                                        End If
                                    Next
                            End Select
                        End If
                End Select
            Next

            '----------
            Table.Clear()
            '----------
            LIST4VIEW("グリッド表示変更", "e")
        End If
    End Sub

    'グリッドプログラム処理
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        If ComboBox6.SelectedItem = "" Then
            Exit Sub
        End If

        LIST4VIEW("設定読出し", "s")
        where = "[名称] = 'csv修正プログラム'"
        where &= " AND [名称2] = '" & ComboBox6.SelectedItem & "'"
        TableName = "T_プログラム"
        '=================================================
        Dim Table As DataTable = MALLDB_CONNECT_SELECT(CnAccdb, TableName, where)
        '=================================================

        dH2 = TM_HEADER_GET(DGV2)
        dH4 = TM_HEADER_GET(DGV4)

        '代表コード変換
        TextBox40.Text = ""
        If CheckBox35.Checked And TextBox39.Text <> "" Then
            Dim tbArray As TextBox() = {TextBox39, TextBox10}
            For i As Integer = 0 To tbArray.Length - 1
                tbArray(i).Text = Regex.Replace(tbArray(i).Text, " |　", "")
                Dim listArray As String() = Split(tbArray(i).Text, vbCrLf)
                Array.Sort(listArray)
                Dim str As String = ""
                For k As Integer = 0 To listArray.Count - 1
                    If listArray(k) <> "" Then
                        str &= listArray(k) & vbCrLf
                    End If
                Next
                tbArray(i).Text = str
            Next

            If CheckBox34.Checked Then
                Dim lines As String = CODE_EXTRACTION(TextBox39.Text)
                TextBox40.Text = Replace(lines, "|", vbCrLf)
                Dim lines2 As String = CODE_EXTRACTION(TextBox10.Text)
                TextBox42.Text = Replace(lines2, "|", vbCrLf)
            End If
        End If

        Dim pStrArray1 As String() = Split(Table.Rows(0)("プログラム").ToString, "|=|")
        Dim dgv As New DataGridView
        Dim dH As New ArrayList
        Dim dH5 As ArrayList = TM_HEADER_GET(DGV5)
        Dim codeArray As New ArrayList  'item、selectで変更する代表コードを通しで保存する
        For i As Integer = 0 To pStrArray1.Length - 1
            Select Case pStrArray1(i)
                Case "1"
                    dgv = DGV2
                    dH = dH2
                Case "2"
                    If SplitContainer7.Panel2Collapsed Then
                        Exit For
                    Else
                        dgv = DGV4
                        dH = dH4
                    End If
                Case Else
                    Select Case True
                        Case Regex.IsMatch(pStrArray1(i), "^検索|^REG検索|^行削除|^削除予定|^REG削除予定|^背景追加")
                            'Dim pStrArray2 As String() = Split(pStrArray1(i), ",")
                            Dim pStrArray2 As String() = TM_CSV_SPLIT(pStrArray1(i))    'プログラムにダブルコーテーション囲み対応
                            Dim pList1 As String() = Split(pStrArray2(1), "#")
                            For r As Integer = dgv.Rows.Count - 1 To 0 Step -1

                                Dim pFlag As Boolean = True
                                If Regex.IsMatch(pStrArray1(i), "^REG検索|^REG削除予定") Then
                                    pFlag = False
                                    Dim hantenFlag As Boolean = False
                                    For p As Integer = 0 To pList1.Length - 1
                                        If InStr(pList1(p), "=") > 0 Then
                                            Dim pList2 As String() = Split(pList1(p), "=")
                                            If Regex.IsMatch(pList2(1), "^!") Then
                                                hantenFlag = True
                                            End If
                                            Dim pattern As String = Regex.Replace(pList2(1), "^!", "")
                                            For c As Integer = 0 To dgv.ColumnCount - 1
                                                If Regex.IsMatch(dgv.Columns(c).HeaderText, pList2(0)) Then
                                                    If dgv.Item(c, r).Value <> "" Then
                                                        If Regex.IsMatch(dgv.Item(c, r).Value, pattern) Then
                                                            pFlag = True
                                                            Exit For
                                                        End If
                                                    End If
                                                End If
                                            Next
                                            If pFlag = True Then
                                                Exit For
                                            End If
                                        End If
                                    Next
                                    If hantenFlag Then  '!予約の時に反転
                                        If pFlag Then
                                            pFlag = False
                                        Else
                                            pFlag = True
                                        End If
                                    End If
                                Else
                                    Dim inFlag As Boolean = True
                                    For p As Integer = 0 To pList1.Length - 1
                                        If InStr(pList1(p), "=") > 0 Then
                                            Dim pList2 As String() = Split(pList1(p), "=")
                                            If Regex.IsMatch(pList2(1), "^color\.") Then
                                                pList2(1) = Regex.Replace(pList2(1), "^color\.", "")
                                                If Regex.IsMatch(pList2(1), "^!") Then
                                                    pList2(1) = Regex.Replace(pList2(1), "^!", "")
                                                    If Not dgv.Item(dH.IndexOf(pList2(0)), r).Style.BackColor.Name.ToString = pList2(1) Then
                                                        inFlag = False
                                                    Else
                                                        inFlag = True
                                                    End If
                                                Else
                                                    If dgv.Item(dH.IndexOf(pList2(0)), r).Style.BackColor.Name.ToString = pList2(1) Then
                                                        inFlag = True
                                                    Else
                                                        inFlag = False
                                                    End If
                                                End If
                                            Else
                                                Dim pattern As String = Regex.Replace(pList2(1), "^!", "")
                                                If pattern = "" Then
                                                    If dgv.Item(dH.IndexOf(pList2(0)), r).Value = "" Then
                                                        inFlag = True
                                                    Else
                                                        inFlag = False
                                                    End If
                                                ElseIf dgv.Item(dH.IndexOf(pList2(0)), r).Value = "" Then
                                                    If dgv.Item(dH.IndexOf(pList2(0)), r).Value = pattern Then
                                                        inFlag = True
                                                    ElseIf InStr(pattern, "[none]") > 0 Then
                                                        inFlag = True
                                                    Else
                                                        inFlag = False
                                                    End If
                                                Else
                                                    pattern = Replace(pattern, "[none]", "")
                                                    If Regex.IsMatch(dgv.Item(dH.IndexOf(pList2(0)), r).Value, pattern) Then
                                                        inFlag = True
                                                    Else
                                                        inFlag = False
                                                    End If
                                                End If
                                                '!を付けると状態を反転する
                                                If Regex.IsMatch(pList2(1), "^!") Then
                                                    Dim a = dgv.Item(dH.IndexOf(pList2(0)), r).Value
                                                    If inFlag = True Then
                                                        inFlag = False
                                                    Else
                                                        inFlag = True
                                                    End If
                                                End If
                                            End If

                                            If pFlag = True Then
                                                If inFlag = True Then
                                                    pFlag = True
                                                Else
                                                    pFlag = False
                                                End If
                                            Else
                                                pFlag = False
                                            End If
                                        End If
                                    Next
                                End If

                                If pFlag Then
                                    Select Case True
                                        Case Regex.IsMatch(pStrArray1(i), "^検索|^REG検索")
                                            Dim pList3 As String() = Split(pStrArray2(2), "#")
                                            For p As Integer = 0 To pList3.Length - 1
                                                If InStr(pList3(p), "=") > 0 Then
                                                    Dim pList4 As String() = Split(pList3(p), "=")
                                                    If Regex.IsMatch(CStr(pList4(1)), "%[0-9]%") Then
                                                        Dim tb As String = TextBox39.Text
                                                        Dim tb2 As String = TextBox10.Text
                                                        If CheckBox35.Checked Then
                                                            tb = TextBox40.Text
                                                            tb2 = TextBox10.Text
                                                        End If
                                                        If tb <> "" Or tb2 <> "" Then
                                                            '2つめの予約期間適用
                                                            Dim sArrayStr As String = Regex.Replace(tb, " |　", "")
                                                            Dim sArray As String() = Split(sArrayStr, vbCrLf)
                                                            Dim sArrayStr2 As String = Regex.Replace(tb2, " |　", "")
                                                            Dim sArray2 As String() = Split(sArrayStr2, vbCrLf)

                                                            Dim dgvStr As String = ""
                                                            For c As Integer = 0 To dgv.ColumnCount - 1
                                                                dgvStr &= dgv.Item(c, r).Value
                                                            Next
                                                            Dim sFlag As Boolean = False
                                                            For k As Integer = 0 To sArray.Length - 1
                                                                If InStr(dgvStr, sArray(k)) > 0 Then
                                                                    If sArray(k) = "" Then
                                                                        Continue For
                                                                    End If
                                                                    sFlag = True
                                                                    Exit For
                                                                End If
                                                            Next
                                                            If sFlag Then
                                                                pList4(1) = Replace(pList4(1), "%1%", TextBox38.Text)     'ボックス置換
                                                                pList4(1) = Replace(pList4(1), "%2%", TextBox37.Text)
                                                                pList4(1) = Replace(pList4(1), "%3%", ComboBox24.SelectedItem)
                                                            Else
                                                                For k As Integer = 0 To sArray2.Length - 1
                                                                    If sArray2(k) = "" Then
                                                                        Continue For
                                                                    End If
                                                                    If InStr(dgvStr, sArray2(k)) > 0 Then
                                                                        sFlag = True
                                                                        Exit For
                                                                    End If
                                                                Next
                                                                If sFlag Then
                                                                    pList4(1) = Replace(pList4(1), "%1%", TextBox43.Text)     'ボックス置換
                                                                    pList4(1) = Replace(pList4(1), "%2%", TextBox44.Text)
                                                                    pList4(1) = Replace(pList4(1), "%3%", ComboBox25.SelectedItem)
                                                                Else
                                                                    pList4(1) = Replace(pList4(1), "%1%", TextBox12.Text)     'ボックス置換
                                                                    pList4(1) = Replace(pList4(1), "%2%", TextBox13.Text)
                                                                    pList4(1) = Replace(pList4(1), "%3%", ComboBox23.SelectedItem)
                                                                End If
                                                            End If
                                                        Else
                                                            pList4(1) = Replace(pList4(1), "%1%", TextBox12.Text)     'ボックス置換
                                                            pList4(1) = Replace(pList4(1), "%2%", TextBox13.Text)
                                                            pList4(1) = Replace(pList4(1), "%3%", ComboBox23.SelectedItem)
                                                        End If
                                                    End If

                                                    If Not Regex.IsMatch(pList4(1), "^\[end\]") Then
                                                        'ピンクは変更しない
                                                        If dgv.Item(dH.IndexOf(pList4(0)), r).Style.BackColor <> Color.Pink Then
                                                            If InStr(pList4(1), "$1") > 0 Then
                                                                dgv.Item(dH.IndexOf(pList4(0)), r).Value = Replace(pList4(1), "$1", dgv.Item(dH.IndexOf(pList4(0)), r).Value)
                                                            Else
                                                                If Regex.IsMatch(pList4(0), "在庫あり時納期管理番号") And Regex.IsMatch(ComboBox6.SelectedItem, "予約解除") Then
                                                                    If dgv.Item(dH.IndexOf("商品管理番号（商品URL）"), r).Value <> "" Then
                                                                        dgv.Item(dH.IndexOf(pList4(0)), r).Value = dgv.Item(dH.IndexOf("商品管理番号（商品URL）"), r).Value
                                                                        Dim findByd5 As Boolean = False
                                                                        For d5i As Integer = 0 To DGV5.RowCount - 1
                                                                            If dgv.Item(dH.IndexOf("商品管理番号（商品URL）"), r).Value = DGV5.Item(dH5.IndexOf("商品コード"), d5i).Value Or dgv.Item(dH.IndexOf("商品管理番号（商品URL）"), r).Value = DGV5.Item(dH5.IndexOf("代表商品コード"), d5i).Value Then
                                                                                If DGV5.Item(dH5.IndexOf("商品分類タグ"), d5i).Value <> "" Then
                                                                                    If Regex.IsMatch(DGV5.Item(dH5.IndexOf("商品分類タグ"), d5i).Value, "宅配便") Then
                                                                                        dgv.Item(dH.IndexOf(pList4(0)), r).Value = "1"
                                                                                        findByd5 = True
                                                                                    ElseIf Regex.IsMatch(DGV5.Item(dH5.IndexOf("商品分類タグ"), d5i).Value, "メール便") Or Regex.IsMatch(DGV5.Item(dH5.IndexOf("商品分類タグ"), d5i).Value, "定形外") Then
                                                                                        dgv.Item(dH.IndexOf(pList4(0)), r).Value = "4"
                                                                                        findByd5 = True
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                            If findByd5 = False Then
                                                                                dgv.Item(dH.IndexOf(pList4(0)), r).Value = pList4(1)
                                                                            End If
                                                                        Next
                                                                    Else
                                                                        dgv.Item(dH.IndexOf(pList4(0)), r).Value = pList4(1)
                                                                    End If
                                                                Else
                                                                    dgv.Item(dH.IndexOf(pList4(0)), r).Value = pList4(1)
                                                                End If
                                                            End If
                                                            dgv.Item(dH.IndexOf(pList4(0)), r).Style.BackColor = Color.Yellow
                                                        End If
                                                    Else
                                                        For c As Integer = 0 To dgv.ColumnCount - 1
                                                            If Regex.IsMatch(dgv.Columns(c).HeaderText, pList4(0)) Then
                                                                If dgv.Item(c, r).Value = "" Then
                                                                    Dim wStr As String = Replace(pList4(1), "[end]", "")
                                                                    dgv.Item(c, r).Value = wStr
                                                                    dgv.Item(c, r).Style.BackColor = Color.Yellow
                                                                    Exit For
                                                                End If
                                                            End If
                                                        Next
                                                    End If
                                                End If
                                            Next
                                        Case Regex.IsMatch(pStrArray1(i), "^削除予定|^REG削除予定|^背景追加")
                                            Dim pList2 As String() = Split(pStrArray2(1), "#")
                                            For p As Integer = 0 To pList2.Length - 1
                                                If InStr(pList2(p), "=") > 0 Then
                                                    Dim pList3 As String() = Split(pList2(p), "=")
                                                    Select Case True
                                                        Case Regex.IsMatch(pStrArray1(i), "^削除予定")
                                                            dgv.Item(dH.IndexOf(pList3(0)), r).Style.BackColor = Color.Orange
                                                        Case Regex.IsMatch(pStrArray1(i), "^REG削除予定")
                                                            For c As Integer = 0 To dgv.ColumnCount - 1
                                                                If Regex.IsMatch(dgv.Columns(c).HeaderText, pList3(0)) Then
                                                                    If dgv.Item(c, r).Value <> "" Then
                                                                        If Regex.IsMatch(dgv.Item(c, r).Value, pList3(1)) Then
                                                                            dgv.Item(c, r).Style.BackColor = Color.Orange
                                                                        End If
                                                                    End If
                                                                End If
                                                            Next
                                                        Case Regex.IsMatch(pStrArray1(i), "^背景追加")
                                                            dgv.Item(dH.IndexOf(pStrArray2(2)), r).Style.BackColor = Color.Yellow
                                                    End Select
                                                End If
                                            Next
                                        Case Regex.IsMatch(pStrArray1(i), "^行削除,")
                                            dgv.Rows.RemoveAt(r)
                                    End Select
                                End If

                            Next
                        Case Regex.IsMatch(pStrArray1(i), "^行追加,|^行追加1,|^行追加2,")
                            '行追加＝処理DGVに追加、行追加2＝DGV2とDGV4を指定
                            Dim dgvX As DataGridView = dgv
                            If Regex.IsMatch(pStrArray1(i), "^行追加1,") Then
                                dgvX = DGV2
                                dH = dH2
                            ElseIf Regex.IsMatch(pStrArray1(i), "^行追加2,") Then
                                dgvX = DGV4
                                dH = dH4
                            End If

                            Dim pStrArray2 As String() = Split(pStrArray1(i), ",")
                            Dim changeList As String = pStrArray2(1)
                            For r As Integer = 0 To dgvX.Rows.Count - 1
                                Dim code As String = dgvX.Item(dH.IndexOf(changeList), r).Value
                                If Not codeArray.Contains(code) Then
                                    Dim list As String = pStrArray2(2)
                                    list = Replace(list, "_CODE_", code)

                                    If Regex.IsMatch(CStr(list), "%1%|%2%") Then
                                        Dim tb As String = TextBox39.Text
                                        Dim tb2 As String = TextBox10.Text
                                        If CheckBox35.Checked Then
                                            tb = TextBox40.Text
                                            tb2 = TextBox10.Text
                                        End If
                                        If tb <> "" Or tb2 <> "" Then
                                            '2つめの予約期間適用
                                            Dim sArrayStr As String = Regex.Replace(tb, " |　", "")
                                            Dim sArray As String() = Split(sArrayStr, vbCrLf)
                                            Dim sArrayStr2 As String = Regex.Replace(tb2, " |　", "")
                                            Dim sArray2 As String() = Split(sArrayStr2, vbCrLf)

                                            Dim dgvStr As String = ""
                                            For c As Integer = 0 To dgvX.ColumnCount - 1
                                                dgvStr &= dgvX.Item(c, r).Value
                                            Next
                                            Dim sFlag As Boolean = False
                                            For k As Integer = 0 To sArray.Length - 1
                                                If sArray(k) = "" Then
                                                    Continue For
                                                ElseIf InStr(dgvStr, sArray(k)) > 0 Then
                                                    sFlag = True
                                                    Exit For
                                                End If
                                            Next
                                            If sFlag Then
                                                list = Replace(list, "%1%", TextBox38.Text)     'ボックス置換
                                                list = Replace(list, "%2%", TextBox37.Text)
                                                list = Replace(list, "%3%", ComboBox24.SelectedItem)
                                            Else
                                                For k As Integer = 0 To sArray2.Length - 1
                                                    If sArray2(k) = "" Then
                                                        Continue For
                                                    ElseIf InStr(dgvStr, sArray2(k)) > 0 Then
                                                        sFlag = True
                                                        Exit For
                                                    End If
                                                Next
                                                If sFlag Then
                                                    list = Replace(list, "%1%", TextBox43.Text)     'ボックス置換
                                                    list = Replace(list, "%2%", TextBox44.Text)
                                                    list = Replace(list, "%3%", ComboBox25.SelectedItem)
                                                Else
                                                    list = Replace(list, "%1%", TextBox12.Text)     'ボックス置換
                                                    list = Replace(list, "%2%", TextBox13.Text)
                                                    list = Replace(list, "%3%", ComboBox23.SelectedItem)
                                                End If
                                            End If
                                        Else
                                            list = Replace(list, "%1%", TextBox12.Text)     'ボックス置換
                                            list = Replace(list, "%2%", TextBox13.Text)
                                            list = Replace(list, "%3%", ComboBox23.SelectedItem)
                                        End If
                                    End If

                                    'list = Replace(list, "%1%", TextBox12.Text)     'ボックス置換
                                    'list = Replace(list, "%2%", TextBox13.Text)

                                    '選択項目肢名を追加
                                    Dim listRes As String = Yoyaku_Items(dgvX, dH, list, code)
                                    If listRes <> "" Then
                                        list = listRes
                                    End If

                                    Dim lists As String() = Split(list, "#")
                                    dgvX.Rows.Add(lists)
                                    dgvX.Item(dH.IndexOf(changeList), dgvX.Rows.Count - 1).Style.BackColor = Color.Yellow
                                    codeArray.Add(code)
                                End If
                            Next
                        Case Regex.IsMatch(pStrArray1(i), "^並べ替え,")
                            Dim dgvX As DataGridView = dgv
                            Dim pStrArray2 As String() = Split(pStrArray1(i), ",")
                            Dim pStrArray3 As String() = Split(pStrArray2(1), "#")
                            Dim sortStr As String = ""
                            For k As Integer = 0 To pStrArray3.Length - 1
                                sortStr &= pStrArray3(k) & ","
                            Next
                            sortStr = sortStr.TrimEnd(",")
                            Dim dataT As DataTable = TM_DGV_TO_DATATABLE(dgvX)
                            '------------
                            '背景カラー保存
                            For r As Integer = 0 To dataT.Rows.Count - 1
                                If r = 0 Then
                                    dataT.Columns.Add("背景カラー")
                                End If
                                Dim colorStr As String = ""
                                For c As Integer = 0 To dgvX.ColumnCount - 1
                                    Dim cc As String = "empty"
                                    If dgvX.Item(c, r).Style.BackColor = Color.Empty Then
                                        colorStr &= "empty" & ","
                                    Else
                                        colorStr &= New ColorConverter().ConvertToString(dgvX.Item(c, r).Style.BackColor) & ","
                                    End If
                                Next
                                colorStr = colorStr.TrimEnd(",")
                                dataT.Rows(r)("背景カラー") = colorStr
                            Next
                            '------------
                            Dim view As New DataView(dataT) With {
                                .Sort = sortStr
                                }
                            Dim td As New DataTable
                            td = view.ToTable
                            '------------
                            'ソートした後に背景カラー取得
                            Dim colorArray As New ArrayList
                            For r As Integer = 0 To td.Rows.Count - 1
                                colorArray.Add(td.Rows(r)("背景カラー"))
                            Next
                            td.Columns.Remove("背景カラー")
                            '------------
                            TM_DATATABLE_TO_DGV(dgvX, td)
                            dataT.Dispose()
                            td.Dispose()
                            '------------
                            '背景カラーを戻す
                            For r As Integer = 0 To dgvX.RowCount - 1
                                Dim ca As String() = Split(colorArray(r), ",")
                                For c As Integer = 0 To dgvX.ColumnCount - 1
                                    If ca(c) = "" Or ca(c) = "empty" Then
                                        dgvX.Item(c, r).Style.BackColor = Color.Empty
                                    Else
                                        dgvX.Item(c, r).Style.BackColor = New ColorConverter().ConvertFromString(ca(c))
                                    End If
                                Next
                            Next
                            '------------


                            'Dim dgvX As DataGridView = dgv
                            'Dim pStrArray2 As String() = Split(pStrArray1(i), ",")
                            'Dim sortColumn As DataGridViewColumn = dgvX.Columns(dH.IndexOf(pStrArray2(1)))  'DataGridView並べ替え
                            'Dim view As DataView = CType(dgvX.DataSource, DataTable).DefaultView
                            'view.Sort = "支店 ASC,部署名 ASC,達成率 DESC"
                            'dgvX.Sort(sortColumn, System.ComponentModel.ListSortDirection.Ascending)
                        Case Regex.IsMatch(pStrArray1(i), "^整形,")
                            Dim dgvX As DataGridView = dgv
                            Dim pStrArray2 As String() = Split(pStrArray1(i), ",")
                            Select Case pStrArray2(2)
                                Case "Yahoo改行"
                                    For r As Integer = 0 To dgvX.Rows.Count - 1
                                        Dim str As String = dgvX.Item(dH.IndexOf(pStrArray2(1)), r).Value
                                        str = Regex.Replace(str, "\n\r|\n|\r", vbCrLf & vbCrLf)
                                        str = Replace(str, vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf)
                                        dgvX.Item(dH.IndexOf(pStrArray2(1)), r).Value = str
                                    Next
                            End Select
                        Case Regex.IsMatch(pStrArray1(i), "^固定,")
                            Dim dgvX As DataGridView = dgv
                            Dim pStrArray2 As String() = Split(pStrArray1(i), ",")
                            For r As Integer = 0 To dgvX.Rows.Count - 1
                                dgvX.Item(dH.IndexOf(pStrArray2(1)), r).Value = pStrArray2(2)
                            Next
                        Case Else
                            '空行などは飛ばす
                    End Select

                    'ddd
            End Select
        Next

        '----------
        Table.Clear()
        '----------
        LIST4VIEW("プログラム処理", "e")
    End Sub

    Private Sub Table_to_DGV(dgv As DataGridView, table As DataTable)
        dgv.Rows.Clear()
        dgv.Columns.Clear()

        For r As Integer = 0 To table.Rows.Count - 1
            If r = 0 Then   'ヘッダー
                For c As Integer = 0 To table.Columns.Count - 1
                    dgv.Columns.Add("c" & c, table.Columns(c).ColumnName)
                Next
            End If

            Dim strArray As New ArrayList
            For c As Integer = 0 To table.Columns.Count - 1
                strArray.Add(StrAccessConvert(table.Rows(r)(c).ToString))
            Next
            Dim strs As String() = DirectCast(strArray.ToArray(GetType(String)), String())
            dgv.Rows.Add(strs)
        Next
    End Sub

    '選択項目肢名を追加
    Private Function Yoyaku_Items(dgvX As DataGridView, dH As ArrayList, list As String, code As String) As String
        Select Case ToolStripStatusLabel4.Text
            Case "楽天"
                If dgvX.Name = DGV4.Name Then
                    Dim kStr As String = ""
                    For q As Integer = 0 To dgvX.Rows.Count - 1
                        If dgvX.Item(dH.IndexOf("商品管理番号（商品URL）"), q).Value = code Then
                            If dgvX.Item(dH.IndexOf("商品管理番号（商品URL）"), q).Style.BackColor = Color.LightSkyBlue Then
                                kStr &= Replace(dgvX.Item(dH.IndexOf("項目選択肢別在庫用横軸選択肢"), q).Value, "/", "")
                                kStr &= Replace(dgvX.Item(dH.IndexOf("項目選択肢別在庫用縦軸選択肢"), q).Value, "/", "")
                                kStr = Replace(kStr, "種類", "")  '横軸に「種類」の時があるので削除
                                kStr &= "/"
                            End If
                        End If
                    Next
                    If kStr <> "" Then
                        kStr = kStr.TrimEnd("/")
                        kStr = "（" & kStr & "）"
                        list = Replace(list, "予約商品", "予約商品" & kStr)
                        list = Replace(list, "予約につき", "予約につき" & kStr)
                    End If
                    Return list
                End If
            Case "Yahoo"
                If dgvX.Name = DGV4.Name Then
                    Dim kStr As String = ""
                    For q As Integer = 0 To dgvX.Rows.Count - 1
                        If dgvX.Item(dH.IndexOf("code"), q).Value = code Then
                            If dgvX.Item(dH.IndexOf("sub-code"), q).Style.BackColor = Color.LightSkyBlue Then
                                kStr &= Replace(dgvX.Item(dH.IndexOf("option-value-1"), q).Value, "/", "")
                                kStr &= Replace(dgvX.Item(dH.IndexOf("option-value-2"), q).Value, "/", "")
                                kStr &= "/"
                            End If
                        End If
                    Next
                    If kStr <> "" Then
                        kStr = kStr.TrimEnd("/")
                        kStr = "（" & kStr & "）"
                        list = Replace(list, "予約商品", "予約商品" & kStr)
                        list = Replace(list, "予約につき", "予約につき" & kStr)
                    End If
                    Return list
                End If
        End Select

        Return ""
    End Function

    ''' <summary>
    ''' 一括適用
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Button10.PerformClick()
        Button13.PerformClick()
    End Sub

    ''' <summary>
    ''' 設定読込
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        '=================================================
        Table0 = MALLDB_CONNECT_SELECT(CnAccdb, "T_" & ComboBox8.SelectedItem)
        '=================================================
        Table_to_DGV(DGV6, Table0)
        '----------
        Table0.Clear()
        '----------
    End Sub

    ''' <summary>
    ''' 設定保存
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        MALLDB_CONNECT_UPSERT_DGV(CnAccdb, "T_" & ComboBox8.SelectedItem, DGV6, 0)
        'Settei_write()
        MsgBox("保存しました", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
    End Sub

    Private Sub Settei_write()
        Dim serverPath As String = Form1.サーバーToolStripMenuItem.Text
        If InStr(Form1.appPath, "takashi") > 0 And InStr(Form1.appPath, "debug") > 0 Then  '自宅テスト用
            serverPath = SpecialDirectories.MyDocuments & "\test"
        End If

        Dim localPath As String = Path.GetDirectoryName(Form1.appPath) & "\tenpo\"
        Dim fPath As String = serverPath & "\tenpo\settei_lasttime.dat"
        File.WriteAllText(fPath, Now)

        Dim str As String = ""
        Dim tableName As String() = New String() {"基本情報", "設定_基本情報", "設定_テンプレート", "プログラム", "その他設定", "ヘルプ", "設定_Amazon"}
        MALLDB_CONNECT(CnAccdb)

        ProgressBar1.Value = 0
        ProgressBar1.Maximum = tableName.Length

        For i As Integer = 0 To tableName.Length - 1
            Dim Table As DataTable = MALLDB_GET_SQL("T_" & tableName(i), "")
            str &= "#####" & "T_" & tableName(i) & vbNewLine
            For c As Integer = 0 To Table.Columns.Count - 1
                Dim res As String = Replace(Table.Columns(c).ColumnName.ToString, vbCrLf, vbLf)
                str &= res & "|$$|"
            Next
            str &= vbCrLf
            For r As Integer = 0 To Table.Rows.Count - 1
                For c As Integer = 0 To Table.Columns.Count - 1
                    Dim res As String = Replace(Table.Rows(r)(c).ToString, vbCrLf, vbLf)
                    str &= res & "|$$|"
                Next
                str &= vbCrLf
            Next
            Table.Clear()
            ProgressBar1.Value += 1
            Application.DoEvents()
        Next
        Dim fPath2 As String = serverPath & "\tenpo\settei_lasttime.dat"
        File.WriteAllText(fPath2, str)

        ProgressBar1.Value = 0
        MALLDB_DISCONNECT()
    End Sub

    ''' <summary>
    ''' 設定画面SelectionChanged
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DataGridView6_SelectionChanged(sender As Object, e As EventArgs) Handles DGV6.SelectionChanged
        Dim dgv As DataGridView = sender

        If selChangeFlag = False Or dgv.Rows.Count = 0 Then
            Exit Sub
        End If

        If dgv.SelectedCells.Count > 0 Then
            If Not dgv.SelectedCells(0).Value Is Nothing Then
                Dim str As String = dgv.SelectedCells(0).Value.ToString
                If ComboBox9.SelectedIndex <> 0 Then
                    Dim pattern As String() = Split(ComboBox9.SelectedItem, "/")
                    str = Regex.Replace(str, pattern(0), vbCrLf)
                End If
                AzukiControl2.Text = str
            End If
            dgv.Focus()
        End If

        '選択行が表示されていない時は選択を解除
        Dim selCell As ArrayList = New ArrayList
        For i As Integer = 0 To dgv.SelectedCells.Count - 1
            If dgv.Rows(dgv.SelectedCells(i).RowIndex).Visible = False Then
                selCell.Add(dgv.SelectedCells(i).ColumnIndex & "," & dgv.SelectedCells(i).RowIndex)
            End If
        Next
        For i As Integer = 0 To selCell.Count - 1
            Dim index As String() = Split(selCell(i), ",")
            dgv.Item(CInt(index(0)), CInt(index(1))).Selected = False
        Next

        '選択行の背景色を変更する
        If dgv.SelectedCells.Count > 50 Then
            Exit Sub
        End If

        For r As Integer = 0 To dgv.RowCount - 1
            dgv.Rows(r).DefaultCellStyle.BackColor = BeforeColor
        Next

        Dim selRow As ArrayList = New ArrayList
        For i As Integer = 0 To dgv.SelectedCells.Count - 1
            If Not selRow.Contains(dgv.SelectedCells(i).RowIndex) Then
                selRow.Add(dgv.SelectedCells(i).RowIndex)
            End If
        Next

        For i As Integer = 0 To selRow.Count - 1
            BeforeColor = dgv.Rows(selRow(i)).DefaultCellStyle.BackColor
            dgv.Rows(selRow(i)).DefaultCellStyle.BackColor = Color.LightCyan
        Next
    End Sub

    Private Sub AzukiControl2_Validated(sender As Object, e As EventArgs) Handles AzukiControl2.Validated
        Dim str As String = AzukiControl2.Text
        If ComboBox9.SelectedIndex <> 0 Then
            Dim pattern As String() = Split(ComboBox9.SelectedItem, "/")
            str = Regex.Replace(str, vbCrLf, pattern(1))
        End If
        DGV6.CurrentCell.Value = str
    End Sub

    '************************************************************************************************************
    'マスター検索
    Private Sub ToolStripButton9_Click(sender As Object, e As EventArgs) Handles ToolStripButton9.Click
        Dim search As String = Regex.Replace(ToolStripTextBox2.Text, "\s|　", "|")

        Dim num As Integer = DGV5.CurrentCell.RowIndex * (DGV5.ColumnCount - 1)
        num += DGV5.CurrentCell.ColumnIndex
        For r As Integer = 0 To DGV5.Rows.Count - 1
            For c As Integer = 0 To DGV5.Columns.Count - 1
                Dim nows As Integer = (r * (DGV5.ColumnCount - 1)) + c
                If num < nows Then
                    If CStr(DGV5.Item(c, r).Value) <> "" Then
                        If Regex.IsMatch(DGV5.Item(c, r).Value, search) Then
                            DGV5.CurrentCell = DGV5(c, r)
                            Exit Sub
                        End If
                    End If
                End If
            Next
        Next

        MsgBox("検索結果がありませんでした", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
    End Sub

    Private Sub ToolStripTextBox2_KeyUp(sender As Object, e As KeyEventArgs) Handles ToolStripTextBox2.KeyUp
        If e.KeyCode = Keys.Enter Then
            ToolStripButton9.PerformClick()
        End If
    End Sub

    '抽出
    Private Sub ToolStripButton7_Click(sender As Object, e As EventArgs) Handles ToolStripButton7.Click
        Table_to_DGV(DGV5, dgv5table)
        Dim search As String = Regex.Replace(ToolStripTextBox2.Text, "\s|　", "|")

        For r As Integer = DGV5.Rows.Count - 1 To 0 Step -1
            Dim flag As Boolean = False
            For c As Integer = 0 To DGV5.Columns.Count - 1
                If DGV5.Item(c, r).Value <> "" Then
                    If Regex.IsMatch(DGV5.Item(c, r).Value, search) Then
                        flag = True
                        Continue For
                    End If
                End If
            Next
            If flag = False Then
                DGV5.Rows.RemoveAt(r)
            End If
        Next
    End Sub

    'マスターの×ボタン
    Private Sub ToolStripButton8_Click(sender As Object, e As EventArgs) Handles ToolStripButton8.Click
        Table_to_DGV(DGV5, dgv5table)
    End Sub

    Private Sub MasterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MasterToolStripMenuItem.Click
        If ToolStripDropDownButton1.Text = "master" Then
            Exit Sub
        End If
        ToolStripDropDownButton1.Text = "master"
        Table_to_DGV(DGV5, dgv5table)
    End Sub

    Private Sub PageToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PageToolStripMenuItem.Click
        If ToolStripDropDownButton1.Text = "page" Then
            Exit Sub
        End If
        Label19.Visible = True
        Application.DoEvents()
        ToolStripDropDownButton1.Text = "page"
        If dgv5table_p.Rows.Count = 0 Then
            Dim TableDst2 As String = "T_MASTER_page"
            dgv5table_p = MALLDB_CONNECT_SELECT(CnAccdb, TableDst2, "", "商品コード")
        End If
        Table_to_DGV(DGV5, dgv5table_p)
        Label19.Visible = False
    End Sub
    '************************************************************************************************************

    '************************************************************************************************************
    Private Sub 変更行のみ残すToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 変更行のみ残すToolStripMenuItem.Click
        Dim dgv As DataGridView() = New DataGridView() {DGV2, DGV4}
        For i As Integer = 0 To 1
            If dgv(i).Rows.Count = 0 Then
                Continue For
            End If
            For r As Integer = dgv(i).Rows.Count - 1 To 0 Step -1
                If Not ChangeRowSearch(dgv(i), r) Then  '黄色が無ければ削除する
                    dgv(i).Rows.RemoveAt(r)
                End If
            Next
        Next
    End Sub

    Private Sub 削除予定オレンジ行を削除ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 削除予定オレンジ行を削除ToolStripMenuItem.Click
        Dim dgv As DataGridView() = New DataGridView() {DGV2, DGV4}
        For i As Integer = 0 To 1
            If dgv(i).Rows.Count = 0 Then
                Continue For
            End If
            For r As Integer = dgv(i).Rows.Count - 1 To 0 Step -1
                If DeleteRowSearch(dgv(i), r) Then  'オレンジがあれば削除する
                    dgv(i).Rows.RemoveAt(r)
                End If
            Next
        Next
    End Sub

    Private Sub 変更のある代表コードを残すToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 変更のある代表コードを残すToolStripMenuItem.Click
        'モールを調べる
        Dim codeHeader As String = ""
        For i As Integer = 0 To TenpoTable.Rows.Count - 1
            If InStr(TenpoTable.Rows(i)("店名"), ToolStripComboBox2.Text) > 0 Then
                Select Case TenpoTable.Rows(i)("モール")
                    Case "楽天"
                        codeHeader = "商品管理番号（商品URL）"
                    Case "Yahoo"
                        codeHeader = "code"
                End Select
                Exit For
            End If
        Next

        dH2 = TM_HEADER_GET(DGV2)
        dH4 = TM_HEADER_GET(DGV4)

        Dim dgv As DataGridView() = New DataGridView() {DGV2, DGV4}
        For i As Integer = 0 To 1
            If dgv(i).Rows.Count = 0 Then
                Continue For
            End If
            Dim codeArray As New ArrayList  '変更のあるコードを保持する
            Dim codeNo As Integer = 0
            If i = 0 Then
                codeNo = dH2.IndexOf(codeHeader)
            Else
                codeNo = dH4.IndexOf(codeHeader)
            End If
            For r As Integer = dgv(i).Rows.Count - 1 To 0 Step -1
                If ChangeRowSearch(dgv(i), r) Then
                    If Not codeArray.Contains(dgv(i).Item(codeNo, r).Value) Then
                        codeArray.Add(dgv(i).Item(codeNo, r).Value)
                    End If
                End If
            Next
            '保持したコードと合致する行以外を削除する
            For r As Integer = dgv(i).Rows.Count - 1 To 0 Step -1
                If Not codeArray.Contains(dgv(i).Item(codeNo, r).Value) Then
                    dgv(i).Rows.RemoveAt(r)
                End If
            Next
        Next
    End Sub

    Private Function ChangeRowSearch(dgv As DataGridView, r As Integer) As Boolean
        Dim flag As Boolean = False
        For c As Integer = 0 To dgv.ColumnCount - 1
            If dgv.Item(c, r).Style.BackColor = Color.Yellow Then
                Return True
                Exit For
            End If
        Next
        If flag = False Then
            Return False
        End If
    End Function

    Private Function DeleteRowSearch(dgv As DataGridView, r As Integer) As Boolean
        Dim flag As Boolean = False
        For c As Integer = 0 To dgv.ColumnCount - 1
            If dgv.Item(c, r).Style.BackColor = Color.Orange Then
                Return True
                Exit For
            End If
        Next
        If flag = False Then
            Return False
        End If
    End Function

    Private Sub 選択セルの左で固定ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 選択セルの左で固定ToolStripMenuItem.Click
        Dim dgv As DataGridView = FocusDGV
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
        Dim dgv As DataGridView = FocusDGV
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

    Private Sub HTMLTIDYANDFORMATToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HTMLTIDYANDFORMATToolStripMenuItem.Click
        Dim htmlRecords As New ArrayList

        ToolStripProgressBar1.Value = 0
        ToolStripProgressBar1.Maximum = FocusDGV.SelectedCells.Count
        For i As Integer = 0 To FocusDGV.SelectedCells.Count - 1
            'ZetaHtmlTidy
            Dim html As String = FocusDGV.SelectedCells(i).Value
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

            FocusDGV.SelectedCells(i).Value = str
            ToolStripProgressBar1.Value += 1
        Next

        ToolStripProgressBar1.Value = 0
        MsgBox("終了しました", MsgBoxStyle.SystemModal)
    End Sub

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

    Private Sub 英数全角半角変換ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 英数全角半角変換ToolStripMenuItem.Click
        ToolStripProgressBar1.Maximum = FocusDGV.SelectedCells.Count
        ToolStripProgressBar1.Value = 0
        For i As Integer = 0 To FocusDGV.SelectedCells.Count - 1
            FocusDGV.SelectedCells(i).Value = Form1.StrConvNumeric(FocusDGV.SelectedCells(i).Value)
            FocusDGV.SelectedCells(i).Value = Form1.StrConvEnglish(FocusDGV.SelectedCells(i).Value)
            ToolStripProgressBar1.Value += 1
            Application.DoEvents()
        Next
        ToolStripProgressBar1.Value = 0
        MsgBox("変換終了", MsgBoxStyle.SystemModal)
    End Sub

    Private Sub スペース半角変換ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles スペース半角変換ToolStripMenuItem.Click
        ToolStripProgressBar1.Maximum = FocusDGV.SelectedCells.Count
        ToolStripProgressBar1.Value = 0
        For i As Integer = 0 To FocusDGV.SelectedCells.Count - 1
            FocusDGV.SelectedCells(i).Value = Form1.StrConvZSpaceToH(FocusDGV.SelectedCells(i).Value)
            ToolStripProgressBar1.Value += 1
            Application.DoEvents()
        Next
        ToolStripProgressBar1.Value = 0
        MsgBox("変換終了", MsgBoxStyle.SystemModal)
    End Sub

    Private Sub 大文字小文字変換選択セルのみToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 大文字小文字変換選択セルのみToolStripMenuItem.Click
        ToolStripProgressBar1.Maximum = FocusDGV.SelectedCells.Count
        ToolStripProgressBar1.Value = 0
        For i As Integer = 0 To FocusDGV.SelectedCells.Count - 1
            FocusDGV.SelectedCells(i).Value = Form1.StrConvToNarrow(FocusDGV.SelectedCells(i).Value)
        Next
        ToolStripProgressBar1.Value = 0
        MsgBox("変換終了", MsgBoxStyle.SystemModal)
    End Sub

    Private Sub 片仮名半角全角ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 片仮名半角全角ToolStripMenuItem.Click
        ToolStripProgressBar1.Maximum = FocusDGV.SelectedCells.Count
        ToolStripProgressBar1.Value = 0
        For i As Integer = 0 To FocusDGV.SelectedCells.Count - 1
            FocusDGV.SelectedCells(i).Value = FKanaHan2Zen(FocusDGV.SelectedCells(i).Value)
        Next
        ToolStripProgressBar1.Value = 0
        MsgBox("変換終了", MsgBoxStyle.SystemModal)
    End Sub

    Private Function FKanaHan2Zen(ByVal myString As String) As String
        Dim i As Integer, chrKana As String, strKana As String = ""
        Dim strTemp As String = ""
        If myString Like "*[｡-ﾟ]*" Then
            For i = 0 To myString.Length - 1
                chrKana = myString.Chars(i)
                Select Case Asc(chrKana)
                    Case 161 To 223
                        strKana &= chrKana  '半角が続いたら文字をつなぐ
                    Case Else  '全角文字になったら半角の文字を全角に変換
                        If strKana.Length > 0 Then
                            strTemp &= StrConv(strKana, VbStrConv.Wide)
                            strKana = ""
                        End If
                        strTemp &= chrKana
                End Select
            Next i
            If strKana.Length > 0 Then  '最後の文字が半角の場合の処理
                strTemp &= StrConv(strKana, VbStrConv.Wide)
            End If
        Else
            Return myString
        End If
        Return strTemp
    End Function

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

        Dim dgX As DataGridView = DGV2
        If ToolStripButton6.BackColor = Color.Yellow Then
            dgX = DGV4
        End If

        ToolStripProgressBar1.Maximum = dgX.SelectedCells.Count
        ToolStripProgressBar1.Value = 0

        For i As Integer = 0 To dgX.SelectedCells.Count - 1
            Dim htmlStr As String = dgX.SelectedCells(i).Value
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

    '画像ダウンロード新ロジック（18/10/17）
    Private Sub 画像URL全てToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 画像URL全てToolStripMenuItem.Click
        Dim saveFolder As String = ""
        Dim dlg = VistaFolderBrowserDialog1
        dlg.RootFolder = Environment.SpecialFolder.Desktop
        dlg.Description = "画像を保存するフォルダを選択してください"
        If dlg.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
            saveFolder = dlg.SelectedPath
        End If
        If saveFolder = "" Then
            Exit Sub
        End If

        Dim dgX As DataGridView = DGV2
        If ToolStripButton6.BackColor = Color.Yellow Then
            dgX = DGV4
        End If

        ToolStripProgressBar1.Maximum = dgX.SelectedCells.Count
        ToolStripProgressBar1.Value = 0

        Dim pattern As String = "(?<text>http.*?(\.jpg|\.jpeg|\.gif|\.png|\.tif|\.bmp|""|'|\s))"
        For p As Integer = 0 To dgX.SelectedCells.Count - 1
            Dim htmlStr As String = dgX.SelectedCells(p).Value
            htmlStr = Replace(htmlStr, vbCrLf, vbLf)
            htmlStr = Replace(htmlStr, vbCr, vbLf)
            Dim htmlLines As String() = Split(htmlStr, vbLf)
            For k As Integer = 0 To htmlLines.Length - 1
                Dim m As MatchCollection = Regex.Matches(htmlLines(k), pattern, RegexOptions.IgnoreCase Or RegexOptions.Singleline)
                For i As Integer = 0 To m.Count - 1
                    Dim res As String = m(i).Groups("text").Value
                    'res = Replace(res, """", "")
                    'res = Replace(res, "'", "")
                    TM_IMGDL(saveFolder, res)
                Next
            Next

            ToolStripProgressBar1.Value += 1
        Next

        ToolStripProgressBar1.Value = 0
        MsgBox("ファイルダウンロードしました", MsgBoxStyle.SystemModal)
    End Sub

    Private Sub 整形のみToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 整形のみToolStripMenuItem.Click
        Seikei()
    End Sub

    Private Sub 行頭SPTAB除去ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 行頭SPTAB除去ToolStripMenuItem.Click
        Sptabdel()
        Seikei()
    End Sub

    Private Sub Sptabdel()
        ToolStripProgressBar1.Maximum = FocusDGV.SelectedCells.Count
        ToolStripProgressBar1.Value = 0
        For i As Integer = 0 To FocusDGV.SelectedCells.Count - 1
            Dim str As String = FocusDGV.SelectedCells(i).Value
            If str = "" Then
                Continue For
            End If
            str = LTrim(str)
            str = Replace(str, "  <", "<")
            str = Replace(str, " <", "<")
            str = Replace(str, vbTab, "")
            FocusDGV.SelectedCells(i).Value = str
        Next
        ToolStripProgressBar1.Value = 0
        MsgBox("変換終了", MsgBoxStyle.SystemModal)
    End Sub

    Private Sub Seikei()
        ToolStripProgressBar1.Maximum = FocusDGV.SelectedCells.Count
        ToolStripProgressBar1.Value = 0
        For i As Integer = 0 To FocusDGV.SelectedCells.Count - 1
            Dim str As String = FocusDGV.SelectedCells(i).Value
            If str = "" Then
                Continue For
            End If
            SeikeiHTML(str)
            FocusDGV.SelectedCells(i).Value = str
        Next
        ToolStripProgressBar1.Value = 0
        MsgBox("変換終了", MsgBoxStyle.SystemModal)
    End Sub

    Public Function SeikeiHTML(ByVal html As String) As String
        Dim htmlRecords As ArrayList = New ArrayList

        'ZetaHtmlTidy
        Dim s As String = ""
        Using tidy As HtmlTidy = New HtmlTidy()
            s = tidy.CleanHtml(html, HtmlTidyOptions.ConvertToXhtml)
        End Using

        s = Replace(s, vbCrLf, "")
        s = Replace(s, vbLf, "")
        s = Replace(s, vbCr, "")
        s = Regex.Replace(s, "<!DOCTYPE h.+HTML Tidy.+head><body>", "")
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

        Return str
    End Function

    Private Sub Br追加ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Br追加ToolStripMenuItem.Click
        ToolStripProgressBar1.Maximum = FocusDGV.SelectedCells.Count
        ToolStripProgressBar1.Value = 0
        For i As Integer = 0 To FocusDGV.SelectedCells.Count - 1
            Dim str As String = FocusDGV.SelectedCells(i).Value
            If str = "" Then
                Continue For
            End If
            str = Regex.Replace(str, vbCrLf & "|" & vbCr, vbLf)
            str = Replace(str, vbLf, "<br>" & vbLf)
            FocusDGV.SelectedCells(i).Value = str
        Next
        ToolStripProgressBar1.Value = 0
        MsgBox("変換終了", MsgBoxStyle.SystemModal)
    End Sub

    Private Sub Br削除ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Br削除ToolStripMenuItem.Click
        ToolStripProgressBar1.Maximum = FocusDGV.SelectedCells.Count
        ToolStripProgressBar1.Value = 0
        For i As Integer = 0 To FocusDGV.SelectedCells.Count - 1
            Dim str As String = FocusDGV.SelectedCells(i).Value
            If str = "" Then
                Continue For
            End If
            str = Replace(str, "<br>", "")
            FocusDGV.SelectedCells(i).Value = str
        Next
        ToolStripProgressBar1.Value = 0
        MsgBox("変換終了", MsgBoxStyle.SystemModal)
    End Sub

    '************************************************************************************************************

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        'Dim imgStr As String = "<img src=""(?<t>.*?)"".*>|<img src='(?<t>.*?)'.*>"
        Dim htmlLines As String = FocusDGV.SelectedCells(0).Value
        Dim tStr As String = ""
        'Select Case True
        '    Case Regex.IsMatch(htmlLines, ".jpg|.gif|.png")
        '        Dim results As MatchCollection = Regex.Matches(htmlLines, imgStr, RegexOptions.IgnoreCase Or RegexOptions.Singleline)
        '        For Each m As Match In results
        '            tStr &= m.Groups("t").Value & vbNewLine
        '        Next
        'End Select
        Dim strRegex As String = "http.*\.(jpg|jpeg|gif|png)"
        Dim myRegex As New Regex(strRegex, RegexOptions.None)
        For Each myMatch As Match In myRegex.Matches(htmlLines)
            If myMatch.Success Then
                tStr &= myMatch.Value & vbNewLine
            End If
        Next

        If tStr = "" Then
            MsgBox("対象データがありません", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
        Else
            Clipboard.SetText(tStr)
            MsgBox("クリップボードに保存しました", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
        End If
    End Sub

    '************************************************************************************************************

    Private Sub サブウィンドウToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles サブウィンドウToolStripMenuItem.Click
        Mall_subwindow.Show()
        Mall_subwindow.BringToFront()
        Mall_subwindow.Location = New Point(Me.Location.X + Me.Size.Width, Me.Location.Y)
        'Mall_subwindow.TabControl1.SelectedTab = Newitem.TabControl1.TabPages("TabPage1")
    End Sub

    Private Sub Mall_main_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        If Mall_subwindow.Visible Then
            Mall_subwindow.BringToFront()
            Mall_subwindow.Location = New Point(Me.Location.X + Me.Size.Width, Me.Location.Y)
        End If
    End Sub

    Private Sub Mall_main_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged
        If Mall_subwindow.Visible Then
            Mall_subwindow.BringToFront()
            Mall_subwindow.Location = New Point(Me.Location.X + Me.Size.Width, Me.Location.Y)
        End If
    End Sub

    '************************************************************************************************************

    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        If TabControl1.SelectedTab.Text = "P" Then
            RichTextBox1.Visible = True
        ElseIf TabControl1.SelectedTab.Text = "計算" Then
            ComboBox19.Items.Clear()
            ComboBox20.Items.Clear()
            ComboBox21.Items.Clear()
            If DGV2.Rows.Count > 0 Or DGV4.Rows.Count > 0 Then
                ComboBox19.Items.Add("選択してください")
                ComboBox20.Items.Add("選択してください")
                ComboBox21.Items.Add("選択してください")
                If DGV2.Rows.Count > 0 Then
                    For c As Integer = 0 To DGV2.ColumnCount - 1
                        ComboBox19.Items.Add("上：" & DGV2.Columns(c).HeaderText)
                        ComboBox20.Items.Add("上：" & DGV2.Columns(c).HeaderText)
                        ComboBox21.Items.Add("上：" & DGV2.Columns(c).HeaderText)
                    Next
                End If
                If DGV4.Rows.Count > 0 Then
                    For c As Integer = 0 To DGV4.ColumnCount - 1
                        ComboBox19.Items.Add("下：" & DGV4.Columns(c).HeaderText)
                        ComboBox20.Items.Add("下：" & DGV4.Columns(c).HeaderText)
                        ComboBox21.Items.Add("下：" & DGV4.Columns(c).HeaderText)
                    Next
                End If
                ComboBox19.SelectedIndex = 0
                ComboBox20.SelectedIndex = 0
                ComboBox21.SelectedIndex = 0
            End If
        Else
            RichTextBox1.Visible = False
        End If
    End Sub

    '(日1-10)=下#(日11-20)=上#(日21-31)=中
    '(日1-10)=(月+1)月中下旬#(日11-20)=(月+1)月下旬～(月+2)月上旬#(日21-31)=(月+2)月上中旬
    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        where = "[名称2] = '" & ComboBox6.SelectedItem & "'"
        TableName = "T_プログラム"
        '=================================================
        Dim Table As DataTable = MALLDB_CONNECT_SELECT(CnAccdb, TableName, where)
        '=================================================

        If TabControl1.SelectedTab.Text = "P" Then
            RichTextBox1.Visible = True
        Else
            RichTextBox1.Visible = False
        End If
        If Table.Rows.Count > 0 Then
            RichTextBox1.Text = Table.Rows(0)("説明")
        Else
            RichTextBox1.Text = ""
        End If

        Dim header As String() = New String() {"replace1", "replace2"}
        Dim tb As TextBox() = New TextBox() {TextBox12, TextBox13}
        For i As Integer = 0 To 1
            If Table.Rows(0)(header(i)).ToString <> "" Then
                Dim str As String = Table.Rows(0)(header(i)).ToString
                str = Regex.Replace(str, "\(月.*?\)", New MatchEvaluator(AddressOf REGEX_TUKI))

                If Regex.IsMatch(str, "\(日.*?\)") Then
                    str = REGEX_JIKI(str)
                End If

                tb(i).Text = str
            Else
                tb(i).Text = ""
            End If
        Next
        TextBox38.Text = TextBox12.Text
        TextBox37.Text = TextBox13.Text
    End Sub

    '(月+2)→現在月+2を返す
    Private Shared Function REGEX_TUKI(ByVal m As System.Text.RegularExpressions.Match) As String
        Dim str As String = m.Value
        str = Regex.Replace(str, "\(月|\)|\+", "")
        Dim res As String = DateAdd(DateInterval.Month, Double.Parse(str), Now).Month
        Return res
    End Function

    Private Shared Function REGEX_JIKI(m As String) As String
        Dim sArray1 As String() = Split(m, "#")
        For i As Integer = 0 To sArray1.Length - 1
            Dim str As String = ""
            Dim sArray2 As String() = Split(sArray1(i), "=")
            str = Regex.Replace(sArray2(0), "\(日|\)", "")
            If Regex.IsMatch(Now.Date.ToString, "[" & str & "]") Then
                Return sArray2(1)
            End If
        Next
        Return ""
    End Function

    '************************************************************************************************************
    '検索
    Private Sub TextBox15_Click(sender As Object, e As EventArgs) Handles TextBox15.Click
        'TB15_change()
    End Sub

    Private Sub TextBox15_Validated(sender As Object, e As EventArgs) Handles TextBox15.Validated
        'TB15_change()
    End Sub

    Private Sub TB15_change()
        'If TextBox15.ForeColor = Color.DarkGray Then
        '    TextBox15.Text = ""
        '    TextBox15.ForeColor = Color.Black
        'Else
        '    If TextBox15.Text = "" Then
        '        TextBox15.ForeColor = Color.DarkGray
        '        TextBox15.Text = "検索文字"
        '    End If
        'End If
    End Sub

    Private Sub TextBox16_Click(sender As Object, e As EventArgs) Handles TextBox16.Click
        TB16_change()
    End Sub

    Private Sub TextBox16_Validated(sender As Object, e As EventArgs) Handles TextBox16.Validated
        TB16_change()
    End Sub

    Private Sub TB16_change()
        'If TextBox16.ForeColor = Color.DarkGray Then
        '    TextBox16.Text = ""
        '    TextBox16.ForeColor = Color.Black
        'Else
        '    If TextBox16.Text = "" Then
        '        TextBox16.ForeColor = Color.DarkGray
        '        TextBox16.Text = "置換文字"
        '    End If
        'End If
    End Sub

    '検索
    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        If TextBox15.Text = "" Then
            MsgBox("検索する文字が入力されていません", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
            Exit Sub
        ElseIf FocusDGV.Rows.Count = 0 Then
            MsgBox("データが読み込まれていません", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        Dim search As String = Regex.Replace(TextBox15.Text, "\s|　", "|")

        Dim num As Integer = FocusDGV.CurrentCell.RowIndex * (FocusDGV.ColumnCount - 1)
        num += FocusDGV.CurrentCell.ColumnIndex
        For r As Integer = 0 To FocusDGV.Rows.Count - 1
            For c As Integer = 0 To FocusDGV.Columns.Count - 1
                Dim nows As Integer = (r * (FocusDGV.ColumnCount - 1)) + c
                If num < nows Then
                    If FocusDGV.Item(c, r).Value <> "" Then
                        If CheckBox16.Checked Then
                            If Regex.IsMatch(FocusDGV.Item(c, r).Value, search) Then
                                FocusDGV.CurrentCell = FocusDGV(c, r)
                                FocusDGV.CurrentCell.Style.BackColor = Color.Yellow
                                Exit Sub
                            End If
                        Else
                            If InStr(FocusDGV.Item(c, r).Value, search) > 0 Then
                                FocusDGV.CurrentCell = FocusDGV(c, r)
                                FocusDGV.CurrentCell.Style.BackColor = Color.Yellow
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            Next
        Next

        MsgBox("検索結果がありませんでした", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
    End Sub

    '置換
    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        If TextBox15.Text = "" Then
            MsgBox("検索する文字が入力されていません", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
            Exit Sub
        ElseIf FocusDGV.Rows.Count = 0 Then
            MsgBox("データが読み込まれていません", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        Dim search As String = TextBox15.Text
        If CheckBox30.Checked Then
            search = Regex.Replace(search, "\s|　", "|")
        End If

        Dim op As RegexOptions
        Select Case ComboBox18.SelectedItem
            Case "複数行"
                op = RegexOptions.Multiline
            Case "単一行"
                op = RegexOptions.Singleline
            Case "小文字"
                op = RegexOptions.IgnoreCase
            Case Else
                op = RegexOptions.None
        End Select

        Dim num As Integer = 0
        ToolStripProgressBar1.Value = 0
        ToolStripProgressBar1.Maximum = FocusDGV.Rows.Count
        For r As Integer = 0 To FocusDGV.Rows.Count - 1
            For c As Integer = 0 To FocusDGV.Columns.Count - 1
                If FocusDGV.Item(c, r).Selected Then
                    If FocusDGV.Item(c, r).Value <> "" Then
                        If CheckBox16.Checked Then
                            If Regex.IsMatch(FocusDGV.Item(c, r).Value, search, op) Then
                                FocusDGV.Item(c, r).Value = Regex.Replace(FocusDGV.Item(c, r).Value, search, TextBox16.Text, op)
                                FocusDGV.Item(c, r).Style.BackColor = Color.Yellow
                                num += 1
                            End If
                        Else
                            If InStr(FocusDGV.Item(c, r).Value, search) > 0 Then
                                FocusDGV.Item(c, r).Value = Replace(FocusDGV.Item(c, r).Value, search, TextBox16.Text)
                                FocusDGV.Item(c, r).Style.BackColor = Color.Yellow
                                num += 1
                            End If
                        End If
                    End If
                End If
            Next
            ToolStripProgressBar1.Value += 1
        Next

        ToolStripProgressBar1.Value = 0
        MsgBox(num & "件 置換しました。", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
    End Sub

    '削除
    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        If TextBox15.Text = "" Then
            MsgBox("検索する文字が入力されていません", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
            Exit Sub
        ElseIf FocusDGV.Rows.Count = 0 Then
            MsgBox("データが読み込まれていません", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        Dim search As String = Regex.Replace(TextBox15.Text, "\s|　", "|")

        Dim num As Integer = 0
        For r As Integer = 0 To FocusDGV.Rows.Count - 1
            For c As Integer = 0 To FocusDGV.Columns.Count - 1
                If FocusDGV.Item(c, r).Selected Then
                    If FocusDGV.Item(c, r).Value <> "" Then
                        If CheckBox16.Checked Then
                            If Regex.IsMatch(FocusDGV.Item(c, r).Value, search) Then
                                FocusDGV.Item(c, r).Value = Regex.Replace(FocusDGV.Item(c, r).Value, search, "")
                                FocusDGV.Item(c, r).Style.BackColor = Color.Yellow
                                num += 1
                            End If
                        Else
                            If InStr(FocusDGV.Item(c, r).Value, search) > 0 Then
                                FocusDGV.Item(c, r).Value = Replace(FocusDGV.Item(c, r).Value, search, "")
                                FocusDGV.Item(c, r).Style.BackColor = Color.Yellow
                                num += 1
                            End If
                        End If
                    End If
                End If
            Next
        Next

        MsgBox(num & "件 削除しました。", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
    End Sub

    '前に一括追加
    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        Dim dgvX As DataGridView = DGV2
        If ToolStripButton6.BackColor = Color.Yellow Then
            dgvX = DGV4
        End If

        Dim num As Integer = 0
        For i As Integer = 0 To dgvX.SelectedCells.Count - 1
            If dgvX.Rows(dgvX.SelectedCells(i).RowIndex).Visible = True Then
                dgvX.SelectedCells(i).Value = TextBox2.Text & dgvX.SelectedCells(i).Value
                num += 1
            End If
        Next
        MsgBox(num & "件 追加しました", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
    End Sub

    '後に一括追加
    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        Dim dgvX As DataGridView = DGV2
        If ToolStripButton6.BackColor = Color.Yellow Then
            dgvX = DGV4
        End If

        Dim num As Integer = 0
        For i As Integer = 0 To dgvX.SelectedCells.Count - 1
            If dgvX.Rows(dgvX.SelectedCells(i).RowIndex).Visible = True Then
                dgvX.SelectedCells(i).Value = dgvX.SelectedCells(i).Value & TextBox2.Text
                num += 1
            End If
        Next
        MsgBox(num & "件 追加しました", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        Dim dgvX As DataGridView = DGV2
        If ToolStripButton6.BackColor = Color.Yellow Then
            dgvX = DGV4
        End If

        Dim num As Integer = 0
        For i As Integer = 0 To dgvX.SelectedCells.Count - 1
            If dgvX.Rows(dgvX.SelectedCells(i).RowIndex).Visible = True Then
                dgvX.SelectedCells(i).Value = Replace(dgvX.SelectedCells(i).Value, TextBox2.Text, "")
                num += 1
            End If
        Next
        MsgBox(num & "件 削除しました", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
    End Sub

    '************************************************************************************************************
    '計算
    '************************************************************************************************************
    Private Sub Button53_Click(sender As Object, e As EventArgs) Handles Button53.Click
        Dim dgvX As DataGridView = DGV2
        If ToolStripButton6.BackColor = Color.Yellow Then
            dgvX = DGV4
        End If
        Dim dHX As ArrayList = TM_HEADER_GET(dgvX)
        For k As Integer = 0 To dgvX.SelectedCells.Count - 1
            Dim selRow As Integer = dgvX.SelectedCells(k).RowIndex
            Dim keisan As String = TextBox34.Text
            Dim h As String = ""
            If ComboBox19.Text <> "選択してください" Then
                h = Regex.Replace(ComboBox19.Text, "上：|下：", "")
                If CStr(dgvX.Item(dHX.IndexOf(h), selRow).Value) = "" Then
                    Continue For
                Else
                    keisan = Replace(keisan, "_A_", dgvX.Item(dHX.IndexOf(h), selRow).Value)
                End If
            End If
            If ComboBox20.Text <> "選択してください" Then
                h = Regex.Replace(ComboBox20.Text, "上：|下：", "")
                If CStr(dgvX.Item(dHX.IndexOf(h), selRow).Value) = "" Then
                    Continue For
                Else
                    keisan = Replace(keisan, "_B_", dgvX.Item(dHX.IndexOf(h), selRow).Value)
                End If
            End If
            If ComboBox21.Text <> "選択してください" Then
                h = Regex.Replace(ComboBox21.Text, "上：|下：", "")
                If CStr(dgvX.Item(dHX.IndexOf(h), selRow).Value) = "" Then
                    Continue For
                Else
                    keisan = Replace(keisan, "_C_", dgvX.Item(dHX.IndexOf(h), selRow).Value)
                End If
            End If

            '文字列を計算する
            Dim dt As New System.Data.DataTable()
            Dim result As Double = CDbl(dt.Compute(keisan, ""))
            If RadioButton12.Checked Then
                result = Math.Floor(result)
            ElseIf RadioButton11.Checked Then
                result = Math.Ceiling(result)
            ElseIf RadioButton10.Checked Then
                result = Math.Round(result)
            End If
            dgvX.SelectedCells(k).Value = result
        Next
    End Sub

    Private Sub LinkLabel6_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel6.LinkClicked
        Dim selectPos As Integer = TextBox34.SelectionStart
        TextBox34.Text = TextBox34.Text.Insert(selectPos, "_A_")
        TextBox34.SelectionStart = selectPos
    End Sub

    Private Sub LinkLabel7_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel7.LinkClicked
        Dim selectPos As Integer = TextBox34.SelectionStart
        TextBox34.Text = TextBox34.Text.Insert(selectPos, "_B_")
        TextBox34.SelectionStart = selectPos
    End Sub

    Private Sub LinkLabel8_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel8.LinkClicked
        Dim selectPos As Integer = TextBox34.SelectionStart
        TextBox34.Text = TextBox34.Text.Insert(selectPos, "_C_")
        TextBox34.SelectionStart = selectPos
    End Sub

    '************************************************************************************************************
    'ランダム文字列置換
    '************************************************************************************************************
    Private Sub Button57_Click(sender As Object, e As EventArgs) Handles Button57.Click
        Dim dgvX As DataGridView = DGV2
        If ToolStripButton6.BackColor = Color.Yellow Then
            dgvX = DGV4
        End If
        Dim dHX As ArrayList = TM_HEADER_GET(dgvX)
        For i As Integer = 0 To dgvX.SelectedCells.Count - 1
            Dim str As String = dgvX.SelectedCells(i).Value
            str = Replace(str, "　", " ")
            str = Regex.Replace(str, "\s{1,5}", " ")
            Dim mae As New ArrayList
            Dim mc As MatchCollection = Regex.Matches(str, "【.*】\s{0,2}")
            For Each m As Match In mc
                mae.Add(m.Value)
            Next
            Dim ato As New ArrayList
            mc = Regex.Matches(str, "\s{0,2}[a-zA-z]{1,2}[0-9]{3}$")
            For Each m As Match In mc
                ato.Add(m.Value)
            Next
            For k As Integer = 0 To mae.Count - 1
                str = Replace(str, mae(k), "")
            Next
            For k As Integer = 0 To ato.Count - 1
                str = Replace(str, ato(k), "")
            Next

            'ランダム並び替え
            Dim strArray As String() = Split(str, " ")
            Dim strArray2 As String() = strArray.OrderBy(Function(s) Guid.NewGuid()).ToArray()
            For k As Integer = 0 To strArray2.Length - 1
                If k = 0 Then
                    str = strArray2(k)
                Else
                    str &= " " & strArray2(k)
                End If
            Next

            For k As Integer = mae.Count - 1 To 0 Step -1
                str = mae(k) & str
            Next
            For k As Integer = ato.Count - 1 To 0 Step -1
                str = str & ato(k)
            Next
            dgvX.SelectedCells(i).Value = str
        Next
    End Sub

    '************************************************************************************************************
    'プログラムcsv変換
    '************************************************************************************************************
    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        'DataGridViewコピー
        DGV8.Rows.Clear()
        DGV8.Columns.Clear()
        For r As Integer = 0 To DGV2.RowCount - 1
            For c As Integer = 0 To DGV2.ColumnCount - 1
                If r = 0 Then
                    DGV8.Columns.Add(c, DGV2.Columns(c).HeaderText)
                End If
                If c = 0 Then
                    DGV8.Rows.Add()
                End If
                DGV8.Item(c, r).Value = DGV2.Item(c, r).Value
            Next
        Next

        where = "[名称2] = '" & ComboBox11.SelectedItem & "'"
        TableName = "T_プログラム"
        '=================================================
        Dim Table As DataTable = MALLDB_CONNECT_SELECT(CnAccdb, TableName, where)
        '=================================================
        Dim pList As String() = Split(Table(0)("プログラム"), "|=|")
        '=================================================
        Table.Clear()
        '=================================================

        Dim p2List As New ArrayList
        Dim p3List As New ArrayList
        Dim p4List As New ArrayList
        DGV2.Rows.Clear()
        DGV2.Columns.Clear()
        Dim pFlag As Integer = 0
        For i As Integer = 0 To pList.Length - 1
            If pList(i) = "" Or Regex.IsMatch(pList(i), "^//|^プログラム") Then
                Continue For
            End If

            Dim h As String() = Split(pList(i), ",")
            DGV2.Columns.Add(i, h(0))
            p2List.Add(h(1))
            p3List.Add(h(2))
            If h.Length = 4 Then
                p4List.Add(h(3))
            End If
        Next

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

        'datagridviewのヘッダーテキストをコレクションに取り込む
        Dim dH2 As ArrayList = TM_HEADER_GET(DGV2)
        Dim dH8 As ArrayList = TM_HEADER_GET(DGV8)

        Dim ProgramList As New ArrayList    'プログラム行は別途処理するため保存
        Dim tokuHozonArray1 As New ArrayList    '特殊用保存リスト
        Dim tokuHozonArray2 As New ArrayList
        For r As Integer = 0 To DGV8.RowCount - 1
            ProgramList.Clear()
            DGV2.Rows.Add(1)
            For c As Integer = 0 To pList.Count - 1
                If pList(c) = "" Then
                    Continue For
                ElseIf Regex.IsMatch(pList(c), "^プログラム") Then
                    Dim pArray As String() = Split(pList(c), ",")
                    ProgramList.Add(Replace(pList(c), "プログラム,", ""))
                    Continue For
                End If

                Select Case True
                    '==================================================
                    Case Regex.IsMatch(p2List(c), "\[規定\]")
                        Dim formatStr As String = p4List(c)
                        If formatStr = "" Then
                            formatStr = "yyyyMMdd"
                        End If
                        Select Case True
                            Case Regex.IsMatch(p3List(c), "今日")
                                DGV2.Item(c, r).Value = Format(Today, formatStr)
                            Case Regex.IsMatch(p3List(c), "明日")
                                DGV2.Item(c, r).Value = Format(DateAdd(DateInterval.Day, 1, Today), formatStr)
                            Case Regex.IsMatch(p3List(c), "日付")
                                Dim dStr As String = Replace(p3List(c), "日付=", "")
                                DGV2.Item(c, r).Value = Format(dStr, formatStr)
                        End Select
                    '==================================================
                    Case Regex.IsMatch(p2List(c), "\[固定\]")
                        DGV2.Item(c, r).Value = p3List(c)
                    '==================================================
                    Case Regex.IsMatch(p2List(c), "\[変換\]")
                        Dim henkanC As String = Replace(p2List(c), "[変換]", "")
                        Dim sArray1 As String() = Split(p3List(c), "|")
                        For i As Integer = 0 To sArray1.Length - 1
                            Dim sArray2 As String() = Split(sArray1(i), "=")
                            If sArray2(0) = "" Then
                                If DGV8.Item(TM_ArIndexof(dH8, henkanC), r).Value = "" Then
                                    DGV2.Item(c, r).Value = sArray2(1)
                                    Exit For
                                End If
                            Else
                                If InStr(DGV8.Item(TM_ArIndexof(dH8, henkanC), r).Value, sArray2(0)) > 0 Then
                                    DGV2.Item(c, r).Value = sArray2(1)
                                    Exit For
                                End If
                            End If
                        Next
                    '==================================================
                    Case Regex.IsMatch(p2List(c), "\[分割\]")
                        Dim henkanC As String = Replace(p2List(c), "[分割]", "")
                        Dim splitword As String = p3List(c)
                        splitword = Replace(splitword, "\s", " ")
                        Dim sArray1 As String() = Split(splitword, "=")
                        If DGV8.Item(TM_ArIndexof(dH8, henkanC), r).Value <> "" Then
                            Dim sArray2 As String() = Split(DGV8.Item(TM_ArIndexof(dH8, henkanC), r).Value, sArray1(0))
                            If sArray1(1) = "last" Then '最後のワードを取得
                                DGV2.Item(c, r).Value = sArray2(sArray2.Length - 1)
                            ElseIf sArray2.Length >= sArray1(1) Then
                                DGV2.Item(c, r).Value = sArray2(sArray1(1) - 1)
                            Else
                                DGV2.Item(c, r).Value = ""
                            End If
                        Else
                            DGV2.Item(c, r).Value = ""
                        End If
                     '==================================================
                    Case Regex.IsMatch(p2List(c), "\[検索\]")
                        Dim henkanC As String = p3List(c)
                        Dim sArray1 As String() = Split(henkanC, "/")       '合計金額=>0/008|   '商品番号=ad/True|False     '商品番号=ad/送料
                        Dim sArray2 As String() = Split(sArray1(0), "=")    '合計金額=>0        '商品番号=ad
                        If InStr(sArray2(1), ">") > 0 Then
                            Dim num As String = Replace(InStr(sArray2(1), ">"), ">", "")
                            If DGV8.Item(TM_ArIndexof(dH8, sArray2(0)), r).Value > num Then
                                Dim sArray4 As String() = Split(sArray1(1), "|")
                                DGV2.Item(c, r).Value = sArray4(0)
                            Else
                                Dim sArray4 As String() = Split(sArray1(1), "|")
                                DGV2.Item(c, r).Value = sArray4(1)
                            End If
                        ElseIf InStr(DGV8.Item(TM_ArIndexof(dH8, sArray2(0)), r).Value, sArray2(1)) > 0 Then   'TrueWord
                            If InStr(sArray1(1), "[固定]") <> 0 Then
                                Dim henkanD As String = Replace(sArray1(1), "[固定]=", "")
                                Dim sArray4 As String() = Split(henkanD, "|")
                                DGV2.Item(c, r).Value = sArray4(0)
                            Else
                                Dim sArray4 As String() = Split(sArray1(1), "|")
                                If InStr(sArray4(0), "+") > 0 Then                      '「+」で連結
                                    Dim s As String() = Split(sArray4(0), "+")
                                    Dim res As String = ""
                                    For i As Integer = 0 To s.Length - 1
                                        res &= DGV8.Item(TM_ArIndexof(dH8, s(i)), r).Value
                                    Next
                                    DGV2.Item(c, r).Value = res
                                Else
                                    DGV2.Item(c, r).Value = DGV8.Item(TM_ArIndexof(dH8, sArray4(0)), r).Value
                                End If
                            End If
                        Else                                                                 'FalseWord
                            If InStr(sArray1(1), "[固定]") <> 0 Then
                                Dim henkanD As String = Replace(sArray1(1), "[固定]=", "")
                                Dim sArray4 As String() = Split(henkanD, "|")
                                DGV2.Item(c, r).Value = sArray4(1)
                            Else
                                Dim sArray4 As String() = Split(sArray1(1), "|")
                                If InStr(sArray4(1), "+") > 0 Then                      '「+」で連結
                                    Dim s As String() = Split(sArray4(1), "+")
                                    Dim res As String = ""
                                    For i As Integer = 0 To s.Length - 1
                                        res &= DGV8.Item(TM_ArIndexof(dH8, s(i)), r).Value
                                    Next
                                    DGV2.Item(c, r).Value = res
                                Else
                                    DGV2.Item(c, r).Value = DGV8.Item(TM_ArIndexof(dH8, sArray4(1)), r).Value
                                End If
                            End If
                        End If
                     '==================================================
                    Case Regex.IsMatch(p2List(c), "\[選択\]")
                        Dim sentakuC As String = p3List(c)
                        If InStr(sentakuC, "|") > 0 Then
                            Dim arrayA As String() = Split(sentakuC, "|")
                            Dim res As String = ""
                            For i As Integer = 0 To arrayA.Length - 1
                                If DGV8.Item(TM_ArIndexof(dH8, arrayA(i)), r).Value <> p4List(c) Then
                                    res = DGV8.Item(TM_ArIndexof(dH8, arrayA(i)), r).Value
                                    Exit For
                                ElseIf DGV8.Item(TM_ArIndexof(dH8, arrayA(i)), r).Value <> "" Then
                                    res = DGV8.Item(TM_ArIndexof(dH8, arrayA(i)), r).Value
                                    Exit For
                                End If
                            Next
                            DGV2.Item(c, r).Value = res
                        Else
                            If DGV8.Item(TM_ArIndexof(dH8, sentakuC), r).Value <> p4List(c) Then
                                DGV2.Item(c, r).Value = DGV8.Item(TM_ArIndexof(dH8, sentakuC), r).Value
                            End If
                        End If
                    '==================================================
                    Case Regex.IsMatch(p2List(c), "\[計算\]")
                        Dim keisan As String = p3List(c)
                        Dim parts As String() = Regex.Split(keisan, "(?<1>\*|\/|\-|\+)")
                        Dim exp As String = ""
                        For i As Integer = 0 To parts.Length - 1
                            If Not IsNumeric(parts(i)) And Not Regex.IsMatch(parts(i), "[\*|\/|\-|\+]{1,1}") Then
                                exp &= DGV8.Item(TM_ArIndexof(dH8, parts(i)), r).Value
                            Else
                                exp &= parts(i)
                            End If
                        Next
                        '文字列を計算する
                        Dim dt As New System.Data.DataTable()
                        Dim result As Double = CDbl(dt.Compute(exp, ""))
                        DGV2.Item(c, r).Value = result
                    '==================================================
                    Case Regex.IsMatch(p2List(c), "\[特殊\]")
                        If Regex.IsMatch(p3List(c), "\[ヤフオク日時\]") Then
                            Dim tokuC As String = Replace(p3List(c), "[特殊]", "")
                            tokuC = DGV8.Item(TM_ArIndexof(dH8, tokuC), r).Value
                            DGV2.Item(c, r).Value = Format(CDate(tokuC), "yyyy/MM/dd HH:mm:ss")
                            'ElseIf Regex.IsMatch(p3List(c), "^Wowma順") Then    'プログラム行で処理した方が良いので没
                            '    Dim choicesStockCode As String() = Split(p4List(c), "|")
                            '    Dim data1 As String = dgv2.Item(dH2.IndexOf(choicesStockCode(0)), r).Value
                            '    Dim data2 As String = dgv2.Item(dH2.IndexOf(choicesStockCode(1)), r).Value
                            '    If Regex.IsMatch(p3List(c), "^Wowma順H") Then
                            '        Dim seq As Integer = 0
                            '        If tokuHozonArray1.Contains(data1) Then
                            '            seq = tokuHozonArray1.IndexOf(data1)
                            '        Else
                            '            seq = tokuHozonArray1.Count
                            '            tokuHozonArray1.Add(data1)
                            '        End If
                            '        If data1 <> "" And data2 = "" Then
                            '            dgv2.Item(c, r).Value = seq + 1
                            '        ElseIf data1 <> "" And data2 <> "" Then
                            '            dgv2.Item(c, r).Value = seq + 1
                            '        End If
                            '    ElseIf Regex.IsMatch(p3List(c), "^Wowma順V") Then
                            '        Dim seq As Integer = 0
                            '        If tokuHozonArray2.Contains(data2) Then
                            '            seq = tokuHozonArray2.IndexOf(data2)
                            '        Else
                            '            seq = tokuHozonArray2.Count
                            '            tokuHozonArray2.Add(data2)
                            '        End If
                            '        If data1 = "" And data2 <> "" Then
                            '            dgv2.Item(c, r).Value = seq + 1
                            '        ElseIf data1 <> "" And data2 <> "" Then
                            '            dgv2.Item(c, r).Value = seq + 1
                            '        End If
                            '    End If
                        End If
                    '==================================================
                    Case Regex.IsMatch(p2List(c), "\[テンプレ\]")
                        Dim files As String = Path.GetDirectoryName(Form1.appPath) & "\template\" & p3List(c) & ".html"
                        If File.Exists(files) Then
                            Dim lines As String = File.ReadAllText(files)
                            DGV2.Item(c, r).Value = lines
                        Else
                            DGV2.Item(c, r).Value = "Template Error"
                        End If
                        '==================================================
                    Case Else
                        If InStr(p2List(c), "+") > 0 Then   '+で連結
                            Dim arrayA As String() = Split(p2List(c), "+")
                            Dim res As String = ""
                            For i As Integer = 0 To arrayA.Length - 1
                                Dim aStr As String = ""
                                Dim aByt As String = ""
                                If InStr(arrayA(i), "[") > 0 Then
                                    aStr = Regex.Replace(arrayA(i), "\[.*\]", "")
                                    aByt = Regex.Match(arrayA(i), "(?<byt>\[.*\])").Groups("byt").Value
                                Else
                                    aStr = arrayA(i)
                                End If
                                If dH8.Contains(aStr) Then
                                    Dim resA As String = DGV8.Item(TM_ArIndexof(dH8, aStr), r).Value
                                    If resA <> "" Then
                                        resA = Regex.Replace(resA, Chr(34) & "|" & Chr(39), "")
                                        '[50]でバイト取得
                                        If aByt <> "" Then
                                            aByt = Regex.Replace(aByt, "\[|\]", "")
                                            resA = SubstringByte(resA, 0, CInt(aByt))
                                        End If
                                    End If
                                    res &= resA
                                Else
                                    res &= aStr
                                End If
                            Next
                            DGV2.Item(c, r).Value = res
                        Else
                            If InStr(p2List(c), "[") > 0 Then
                                Dim aStr As String = Regex.Replace(p2List(c), "\[.*\]", "")
                                Dim aByt As String = Regex.Match(p2List(c), "(?<byt>\[.*\])").Groups("byt").Value
                                Dim res As String = ""
                                If dH8.Contains(aStr) Then
                                    Dim resA As String = DGV8.Item(TM_ArIndexof(dH8, aStr), r).Value
                                    If resA <> "" Then
                                        resA = Regex.Replace(resA, Chr(34) & "|" & Chr(39), "")
                                        '[50]でバイト取得
                                        If aByt <> "" Then
                                            aByt = Regex.Replace(aByt, "\[|\]", "")
                                            resA = SubstringByte(resA, 0, CInt(aByt))
                                        End If
                                    End If
                                    res &= resA
                                Else
                                    res &= aStr
                                End If
                                DGV2.Item(c, r).Value = res
                            Else
                                Dim header As String = p2List(c)
                                If header <> "" Then
                                    DGV2.Item(c, r).Value = DGV8.Item(TM_ArIndexof(dH8, header), r).Value
                                End If
                            End If
                        End If
                End Select
            Next

            'プログラム処理
            For i As Integer = 0 To ProgramList.Count - 1
                If Regex.IsMatch(ProgramList(i), "^画像") Then
                    'モール毎設定
                    Dim headerList1 As String() = Nothing
                    Dim headerList2 As String() = Nothing
                    Dim codeName As String = ""
                    Dim plusUrlMae As String = ""
                    Dim pl As String() = Split(ProgramList(i), "#")
                    Select Case pl(1)
                        Case "ヤフオク,,"
                            headerList1 = New String() {"画像1", "画像2", "画像3", "画像4", "画像5", "画像6", "画像7", "画像8", "画像9", "画像10"}
                            headerList2 = Nothing
                            codeName = "管理番号"
                            plusUrlMae = ""
                        Case "Wowma,,"
                            headerList1 = New String() {"imageUrl1", "imageUrl2", "imageUrl3", "imageUrl4", "imageUrl5", "imageUrl6", "imageUrl7", "imageUrl8", "imageUrl9", "imageUrl10"}
                            headerList2 = New String() {"imageName1", "imageName2", "imageName3", "imageName4", "imageName5", "imageName6", "imageName7", "imageName8", "imageName9", "imageName10"}
                            codeName = "itemCode"
                            plusUrlMae = "https://image.wowma.jp/38238702/img/"
                    End Select
                    Dim code As String = DGV2.Item(TM_ArIndexof(dH2, codeName), r).Value

                    Dim savePath As String = ""
                    Dim sfd As New OpenFileDialog With {
                        .Filter = "すべてのファイル(*.*)|*.*",
                        .AutoUpgradeEnabled = False,
                        .FilterIndex = 0,
                        .Title = "「" & code & "」画像保存先のファイルを選択してください",
                        .RestoreDirectory = True,
                        .CheckPathExists = True
                    }
                    If sfd.ShowDialog() = DialogResult.OK Then
                        savePath = sfd.FileName
                    Else
                        Exit For
                    End If

                    Dim imgFolder As String = Path.GetDirectoryName(savePath)
                    Dim di As New DirectoryInfo(imgFolder)
                    Dim files As FileInfo() = di.GetFiles("*", IO.SearchOption.AllDirectories)
                    'ファイル名をArrayListに変換しソートする
                    Dim filesArray As New ArrayList
                    For Each f As FileInfo In files
                        filesArray.Add(f.FullName)
                    Next
                    'ファイル名のソート修正処理
                    filesArray = FILENAME_SORT(filesArray)

                    '画像1～10に管理番号と合致するファイルを探して書き込む
                    Dim usedFileArray As New ArrayList
                    For k As Integer = 0 To headerList1.Length - 1
                        For Each f As String In filesArray
                            If Regex.IsMatch(Path.GetFileNameWithoutExtension(f), code) Then
                                If Not usedFileArray.Contains(Path.GetFileNameWithoutExtension(f)) Then
                                    Dim hNum As Integer = TM_ArIndexof(dH2, headerList1(k))
                                    DGV2.Item(hNum, r).Value = plusUrlMae & Path.GetFileName(f)
                                    If headerList2 IsNot Nothing Then
                                        Dim hNum2 As Integer = TM_ArIndexof(dH2, headerList2(k))
                                        DGV2.Item(hNum2, r).Value = code
                                    End If
                                    usedFileArray.Add(Path.GetFileNameWithoutExtension(f))
                                    Exit For
                                End If
                            End If
                        Next
                    Next
                ElseIf Regex.IsMatch(ProgramList(i), "^ヤフオク説明文画像") Then
                    '説明文の画像を変更する
                    Dim headerList As String() = New String() {"画像1", "画像2", "画像3", "画像4", "画像5", "画像6", "画像7", "画像8", "画像9", "画像10"}
                    For k As Integer = 0 To headerList.Length - 1
                        Dim imgStr As String = "<img src=https://item-shopping.c.yimg.jp/i/f/_AA_ width=100%><br>" & vbCrLf
                        Dim koumoku As String = DGV2.Item(TM_ArIndexof(dH2, headerList(k)), r).Value
                        If koumoku <> "" Then
                            koumoku = Path.GetFileNameWithoutExtension(koumoku)
                            imgStr = Replace(imgStr, "_AA_", koumoku)
                            DGV2.Item(TM_ArIndexof(dH2, "説明"), r).Value = Replace(DGV2.Item(TM_ArIndexof(dH2, "説明"), r).Value, "<!--IMG-->", imgStr & "<!--IMG-->")
                        End If
                        If k = 6 Then   '7枚まで
                            Exit For
                        End If
                    Next
                    DGV2.Item(TM_ArIndexof(dH2, "説明"), r).Value = Replace(DGV2.Item(TM_ArIndexof(dH2, "説明"), r).Value, "<!--IMG-->", "")
                    '説明を移行
                    Dim str As String = DGV8.Item(TM_ArIndexof(dH8, "explanation"), r).Value
                    str = Regex.Replace(str, "\r|\n|\r\n", "<br>" & vbCrLf)
                    DGV2.Item(TM_ArIndexof(dH2, "説明"), r).Value = Replace(DGV2.Item(TM_ArIndexof(dH2, "説明"), r).Value, "<!--CON-->", str & vbCrLf)
                ElseIf Regex.IsMatch(ProgramList(i), "^\[byt\]") Then
                    'バイトチェックする
                    Dim pListB As String() = Split(ProgramList(i), ",")
                    Dim str As String = DGV2.Item(dH2.IndexOf(pListB(1)), r).Value
                    If str <> "" Then
                        If Encoding.GetEncoding("Shift_JIS").GetByteCount(str) > pListB(2) Then
                            DGV2.Item(dH2.IndexOf(pListB(1)), r).Style.BackColor = Color.Yellow
                        End If
                    End If
                ElseIf Regex.IsMatch(ProgramList(i), "^WowmaSeq") Then
                    'choicesStockHorizontalSeq、choicesStockVerticalSeqの番号入力
                    Dim codeA As String = ""
                    Dim checkArray1 As New ArrayList
                    Dim checkArray2 As New ArrayList
                    For r2 As Integer = 0 To DGV2.RowCount - 1
                        Dim code As String = DGV2.Item(TM_ArIndexof(dH2, "itemCode"), r2).Value
                        If codeA = code Then
                            If CStr(DGV2.Item(TM_ArIndexof(dH2, "choicesStockHorizontalCode"), r2).Value) <> "" Then
                                If Not checkArray1.Contains(DGV2.Item(TM_ArIndexof(dH2, "choicesStockHorizontalCode"), r2).Value) Then
                                    checkArray1.Add(DGV2.Item(TM_ArIndexof(dH2, "choicesStockHorizontalCode"), r2).Value)
                                    DGV2.Item(TM_ArIndexof(dH2, "choicesStockHorizontalSeq"), r2).Value = checkArray1.Count
                                Else
                                    DGV2.Item(TM_ArIndexof(dH2, "choicesStockHorizontalSeq"), r2).Value = checkArray1.IndexOf(DGV2.Item(TM_ArIndexof(dH2, "choicesStockHorizontalCode"), r2).Value) + 1
                                End If
                            End If
                            If CStr(DGV2.Item(TM_ArIndexof(dH2, "choicesStockVerticalCode"), r2).Value) <> "" Then
                                If Not checkArray2.Contains(DGV2.Item(TM_ArIndexof(dH2, "choicesStockVerticalCode"), r2).Value) Then
                                    checkArray2.Add(DGV2.Item(TM_ArIndexof(dH2, "choicesStockVerticalCode"), r2).Value)
                                    DGV2.Item(TM_ArIndexof(dH2, "choicesStockVerticalSeq"), r2).Value = checkArray2.Count
                                Else
                                    DGV2.Item(TM_ArIndexof(dH2, "choicesStockVerticalSeq"), r2).Value = checkArray2.IndexOf(DGV2.Item(TM_ArIndexof(dH2, "choicesStockVerticalCode"), r2).Value) + 1
                                End If
                            End If
                        Else
                            codeA = code
                            checkArray1.Clear()
                            checkArray2.Clear()
                            For r3 As Integer = 0 To r2
                                If code = DGV2.Item(TM_ArIndexof(dH2, "itemCode"), r3).Value Then
                                    If Not checkArray1.Contains(DGV2.Item(TM_ArIndexof(dH2, "choicesStockHorizontalCode"), r3).Value) Then
                                        checkArray1.Add(DGV2.Item(TM_ArIndexof(dH2, "choicesStockHorizontalCode"), r3).Value)
                                    End If
                                    If Not checkArray2.Contains(DGV2.Item(TM_ArIndexof(dH2, "choicesStockVerticalCode"), r3).Value) Then
                                        checkArray2.Add(DGV2.Item(TM_ArIndexof(dH2, "choicesStockVerticalCode"), r3).Value)
                                    End If
                                End If
                            Next
                            If CStr(DGV2.Item(TM_ArIndexof(dH2, "choicesStockHorizontalCode"), r2).Value) <> "" Then
                                DGV2.Item(TM_ArIndexof(dH2, "choicesStockHorizontalSeq"), r2).Value = checkArray1.Count
                            End If
                            If CStr(DGV2.Item(TM_ArIndexof(dH2, "choicesStockVerticalCode"), r2).Value) <> "" Then
                                DGV2.Item(TM_ArIndexof(dH2, "choicesStockVerticalSeq"), r2).Value = checkArray2.Count
                            End If
                        End If
                    Next
                End If
            Next

        Next    '*****************************************************************************

        MsgBox("プログラム処理終了")
    End Sub

    'ファイル名ソート（_1.jpg、_10.jpg、_2.jpgとなるのを、_1.jpg、_2.jpg、_10.jpgに修正）
    '「_」はヤフオク用、「-」はWowma用
    Private Function FILENAME_SORT(filesArray As ArrayList) As ArrayList
        For k As Integer = 1 To 9
            For i As Integer = 0 To filesArray.Count - 1
                If Regex.IsMatch(filesArray(i), "_" & k & "\.") Then
                    Dim check As String = Regex.Replace(filesArray(i), "_" & k & "\.", "_0" & k & ".")
                    If Not filesArray.Contains(check) Then
                        filesArray(i) = check
                    End If
                ElseIf Regex.IsMatch(filesArray(i), "-" & k & "\.") Then
                    Dim check As String = Regex.Replace(filesArray(i), "-" & k & "\.", "-0" & k & ".")
                    If Not filesArray.Contains(check) Then
                        filesArray(i) = check
                    End If
                End If
            Next
        Next
        filesArray.Sort()
        For k As Integer = 1 To 9
            For i As Integer = 0 To filesArray.Count - 1
                If Regex.IsMatch(filesArray(i), "_0" & k & "\.") Then
                    Dim check As String = Regex.Replace(filesArray(i), "_0" & k & "\.", "_" & k & ".")
                    If Not filesArray.Contains(check) Then
                        filesArray(i) = check
                    End If
                ElseIf Regex.IsMatch(filesArray(i), "-0" & k & "\.") Then
                    Dim check As String = Regex.Replace(filesArray(i), "-0" & k & "\.", "-" & k & ".")
                    If Not filesArray.Contains(check) Then
                        filesArray(i) = check
                    End If
                End If
            Next
        Next
        Return filesArray
    End Function

    '画像URL追加
    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        dH2 = TM_HEADER_GET(DGV2)

        Dim sArray As New ArrayList
        For s As Integer = 0 To DGV2.SelectedCells.Count - 1
            If Not sArray.Contains(s) Then
                sArray.Add(DGV2.SelectedCells(s).RowIndex)
            End If
        Next
        For r As Integer = 0 To sArray.Count - 1
            For i As Integer = 0 To 19
                If i = 0 Then
                    DGV2.Item(dH2.IndexOf("imageName1"), sArray(r)).Value = DGV2.Item(dH2.IndexOf("itemCode"), sArray(r)).Value
                    DGV2.Item(dH2.IndexOf("imageUrl1"), sArray(r)).Value = "https://image.wowma.jp/38238702/img/" & DGV2.Item(dH2.IndexOf("itemCode"), sArray(r)).Value & ".jpg"
                    DGV2.Item(dH2.IndexOf("imageName1"), sArray(r)).Style.BackColor = Color.Yellow
                    DGV2.Item(dH2.IndexOf("imageUrl1"), sArray(r)).Style.BackColor = Color.Yellow
                Else
                    Select Case ComboBox10.SelectedItem
                        Case "Wowma"
                            If i <= NumericUpDown2.Value Then
                                DGV2.Item(dH2.IndexOf("imageName" & i + 1), sArray(r)).Value = DGV2.Item(dH2.IndexOf("itemCode"), sArray(r)).Value
                                DGV2.Item(dH2.IndexOf("imageUrl" & i + 1), sArray(r)).Value = "https://image.wowma.jp/38238702/img/" & DGV2.Item(dH2.IndexOf("itemCode"), sArray(r)).Value & "-" & i & ".jpg"
                                DGV2.Item(dH2.IndexOf("imageName" & i + 1), sArray(r)).Style.BackColor = Color.Yellow
                                DGV2.Item(dH2.IndexOf("imageUrl" & i + 1), sArray(r)).Style.BackColor = Color.Yellow
                            Else
                                DGV2.Item(dH2.IndexOf("imageName" & i + 1), sArray(r)).Value = ""
                                DGV2.Item(dH2.IndexOf("imageUrl" & i + 1), sArray(r)).Value = ""
                                DGV2.Item(dH2.IndexOf("imageName" & i + 1), sArray(r)).Style.BackColor = DefaultBackColor
                                DGV2.Item(dH2.IndexOf("imageUrl" & i + 1), sArray(r)).Style.BackColor = DefaultBackColor
                            End If
                            If i >= 9 Then  'Wowmaは10枚まで
                                Exit For
                            End If
                    End Select
                End If
            Next
        Next
    End Sub





















































    '''' <summary>ヘッダー取得</summary>
    '''' <param name="dgv">データグリッドビュー</param>
    'Private Function TM_HEADER_GET(dgv As DataGridView) As ArrayList
    '    Dim DGVheaderCheck As ArrayList = New ArrayList
    '    For c As Integer = 0 To dgv.Columns.Count - 1
    '        DGVheaderCheck.Add(dgv.Columns(c).HeaderText)
    '    Next c
    '    Return DGVheaderCheck
    'End Function

    ''' <summary>テーブルが存在しているか確認する</summary>
    ''' <param name="str1">テーブル名、型、所有者、スキーマ</param>
    ''' <param name="str2">テーブル名、型、所有者、スキーマ</param>
    ''' <param name="str3">テーブル名、型、所有者、スキーマ</param>
    ''' <param name="str4">テーブル名、型、所有者、スキーマ</param>
    Private Function MALLDB_TABLE_EXISTS(str1 As String, str2 As String, str3 As String, str4 As String) As DataTable
        Dim schemaTable As DataTable = Cn0.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {str1, str2, str3, str4})
        Return schemaTable
    End Function

    ''' <summary>接続してTableを取得し閉じる</summary>
    ''' <param name="CnAccdb">データベース名</param>
    ''' <param name="sql">SQL文</param>
    Private Function MALLDB_CONNECT_SELECT_SQL(CnAccdb As String, sql As String) As DataTable
        '--------------------------------------------------------------
        Dim Cn1 As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CnAccdb)
        Dim SQLCm1 = Cn1.CreateCommand
        Dim Adapter1 As New OleDbDataAdapter(SQLCm1)
        Dim Table1 As New DataTable
        Cn1.Open()
        ToolStripStatusLabel3.Text = "接続○"
        ToolStripStatusLabel3.BackColor = Color.Yellow
        '--------------------------------------------------------------
        SQLCm1.CommandText = sql
        Adapter1.Fill(Table1)
        '--------------------------------------------------------------
        ToolStripStatusLabel3.Text = "接続×"
        ToolStripStatusLabel3.BackColor = Color.Lavender
        Cn1.Close()
        Adapter1.Dispose()
        SQLCm1.Dispose()
        Cn1.Dispose()
        '--------------------------------------------------------------
        Return Table1
    End Function

    ''' <summary>接続してTableを取得し閉じる</summary>
    ''' <param name="CnAccdb">データベース名</param>
    ''' <param name="TableName">テーブル名</param>
    ''' <param name="where">where文：WHERE (---)</param>
    ''' <param name="orderby">orderby文：ORDER BY ---</param>
    ''' <param name="limit">limit文：LIMIT ---</param>
    Private Function MALLDB_CONNECT_SELECT(CnAccdb As String, TableName As String, Optional where As String = "", Optional orderby As String = "", Optional limit As String = "") As DataTable
        Try
            '--------------------------------------------------------------
            Dim Cn1 As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CnAccdb)
            Dim SQLCm1 = Cn1.CreateCommand
            Dim Adapter1 As New OleDbDataAdapter(SQLCm1)
            Dim Table1 As New DataTable
            Cn1.Open()
            ToolStripStatusLabel3.Text = "接続○"
            ToolStripStatusLabel3.BackColor = Color.Yellow
            '--------------------------------------------------------------
            If where <> "" Then
                where = " WHERE (" & where & ")"
            End If
            If orderby <> "" Then
                orderby = " ORDER BY " & orderby & ""
            End If
            If limit <> "" Then
                limit = " LIMIT " & limit & ""
            End If
            SQLCm1.CommandText = "SELECT * FROM [" & TableName & "]" & where & orderby & limit
            Adapter1.Fill(Table1)
            '--------------------------------------------------------------
            ToolStripStatusLabel3.Text = "接続×"
            ToolStripStatusLabel3.BackColor = Color.Lavender
            Cn1.Close()
            Adapter1.Dispose()
            SQLCm1.Dispose()
            Cn1.Dispose()
            '--------------------------------------------------------------
            Return Table1
        Catch ex As Exception
            Dim Table1 As New DataTable
            MALLDB_DISCONNECT()
            MsgBox("データベースエラーです。", MsgBoxStyle.Exclamation)
            Return Table1
        End Try
    End Function

    ''' <summary>INSERT_ONE</summary>
    ''' <param name="CnAccdb">データベース名</param>
    ''' <param name="TableName">テーブル名</param>
    ''' <param name="setField">フィールド名（複数可）</param>
    ''' <param name="setStr">セットする文字（複数可）</param>
    Private Function MALLDB_CONNECT_INSERT_ONE(CnAccdb As String, TableName As String, setField As String, setStr As String) As Boolean
        '--------------------------------------------------------------
        Dim Cn1 As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CnAccdb)
        Dim SQLCm1 = Cn1.CreateCommand
        Cn1.Open()
        ToolStripStatusLabel3.Text = "接続○"
        ToolStripStatusLabel3.BackColor = Color.Yellow
        '--------------------------------------------------------------
        SQLCm1.CommandText = "INSERT INTO [" & TableName & "] (" & setField & ") VALUES (" & setStr & ")"
        SQLCm1.ExecuteNonQuery()
        '--------------------------------------------------------------
        ToolStripStatusLabel3.Text = "接続×"
        ToolStripStatusLabel3.BackColor = Color.Lavender
        Cn1.Close()
        SQLCm1.Dispose()
        Cn1.Dispose()
        '--------------------------------------------------------------
        Return True
    End Function

    ''' <summary>UPDATE</summary>
    ''' <param name="CnAccdb">データベース名</param>
    ''' <param name="TableName">テーブル名</param>
    ''' <param name="setStr">セットする文字</param>
    ''' <param name="where">where文：WHERE (---)</param>
    Private Function MALLDB_CONNECT_UPDATE(CnAccdb As String, TableName As String, setStr As String, where As String) As Boolean
        '--------------------------------------------------------------
        Dim Cn1 As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CnAccdb)
        Dim SQLCm1 = Cn1.CreateCommand
        Cn1.Open()
        ToolStripStatusLabel3.Text = "接続○"
        ToolStripStatusLabel3.BackColor = Color.Yellow
        '--------------------------------------------------------------
        If where <> "" Then
            where = " WHERE (" & where & ")"
        End If
        SQLCm1.CommandText = "UPDATE [" & TableName & "] SET " & setStr & where
        SQLCm1.ExecuteNonQuery()
        '--------------------------------------------------------------
        ToolStripStatusLabel3.Text = "接続×"
        ToolStripStatusLabel3.BackColor = Color.Lavender
        Cn1.Close()
        SQLCm1.Dispose()
        Cn1.Dispose()
        '--------------------------------------------------------------
        Return True
    End Function

    Private Sub MALLDB_CONNECT_UPSERT(CnAccdb As String, TableName As String, where As String, FieldArray As ArrayList, SetArray As ArrayList)
        '--------------------------------------------------------------
        Dim Cn1 As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CnAccdb)
        Dim SQLCm1 = Cn1.CreateCommand
        Dim Adapter1 As New OleDbDataAdapter(SQLCm1)
        Dim Table1 As New DataTable
        Cn1.Open()
        ToolStripStatusLabel3.Text = "接続○"
        ToolStripStatusLabel3.BackColor = Color.Yellow
        '--------------------------------------------------------------

        SQLCm1.CommandText = "SELECT * FROM [" & TableName & "] WHERE " & where
        Adapter1.Fill(Table1)

        If Table1.Rows.Count > 0 Then
            setStr = ""
            For i As Integer = 0 To FieldArray.Count - 1
                setStr &= "[" & FieldArray(i) & "] = '" & SetArray(i) & "',"
            Next
            setStr = setStr.TrimEnd(",")
            SQLCm1.CommandText = "UPDATE [" & TableName & "] SET " & setStr & " WHERE " & where
            SQLCm1.ExecuteNonQuery()
        Else
            setField = ""
            setStr = ""
            For i As Integer = 0 To FieldArray.Count - 1
                setField &= "[" & FieldArray(i) & "],"
                setStr &= "'" & SetArray(i) & "',"
            Next
            setField = setField.TrimEnd(",")
            setStr = setStr.TrimEnd(",")
            SQLCm1.CommandText = "INSERT INTO [" & TableName & "] (" & setField & ") VALUES (" & setStr & ")"
            SQLCm1.ExecuteNonQuery()
        End If

        '--------------------------------------------------------------
        ToolStripStatusLabel3.Text = "接続×"
        ToolStripStatusLabel3.BackColor = Color.Lavender
        Cn1.Close()
        Table1.Dispose()
        Adapter1.Dispose()
        SQLCm1.Dispose()
        Cn1.Dispose()
        '--------------------------------------------------------------
    End Sub

    Private Function MALLDB_CONNECT_UPSERT_DGV(CnAccdb As String, TableName As String, dgv As DataGridView, KeyCol As Integer, Optional rows As Integer() = Nothing) As Boolean
        '--------------------------------------------------------------
        Dim Cn1 As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CnAccdb)
        Dim SQLCm1 = Cn1.CreateCommand
        Dim Adapter1 As New OleDbDataAdapter(SQLCm1)
        Dim Table1 As New DataTable
        Cn1.Open()
        ToolStripStatusLabel3.Text = "接続○"
        ToolStripStatusLabel3.BackColor = Color.Yellow
        '--------------------------------------------------------------
        Dim query As String = ""
        For r As Integer = 0 To dgv.Rows.Count - 1
            If Not dgv.Rows(r).IsNewRow Then
                'データを読み出し
                Dim keyStr As String = ""
                If dgv.Columns(0).HeaderText = "ID" Then
                    keyStr &= dgv.Columns(KeyCol).HeaderText & "=" & dgv.Item(KeyCol, r).Value & ""
                Else
                    keyStr &= dgv.Columns(KeyCol).HeaderText & "='" & dgv.Item(KeyCol, r).Value & "'"
                End If
                Table1.Clear()
                If dgv.Item(KeyCol, r).Value <> "" Then  '新規
                    query = "SELECT ID FROM [" & TableName & "] WHERE " & keyStr
                    SQLCm1.CommandText = query
                    Adapter1.Fill(Table1)
                End If

                'データの有無で別処理
                If Table1.Rows.Count = 1 Then
                    setStr = ""
                    For c As Integer = 1 To dgv.ColumnCount - 1
                        Dim res As String = Replace(dgv.Item(c, r).Value, vbCrLf, vbLf)
                        setStr &= "[" & dgv.Columns(c).HeaderText & "] = '" & dgv.Item(c, r).Value & "',"
                    Next
                    setStr = setStr.TrimEnd(",")
                    where = keyStr
                    SQLCm1.CommandText = "UPDATE [" & TableName & "] SET " & setStr & " WHERE " & where
                    SQLCm1.ExecuteNonQuery()
                Else
                    setField = ""
                    setStr = ""
                    For c As Integer = 1 To dgv.ColumnCount - 1
                        Dim res As String = Replace(dgv.Item(c, r).Value, vbCrLf, vbLf)
                        setField &= "[" & dgv.Columns(c).HeaderText & "],"
                        setStr &= "'" & res & "',"
                    Next
                    setField = setField.TrimEnd(",")
                    setStr = setStr.TrimEnd(",")
                    SQLCm1.CommandText = "INSERT INTO [" & TableName & "] (" & setField & ") VALUES (" & setStr & ")"
                    SQLCm1.ExecuteNonQuery()
                End If
            End If
        Next
        '--------------------------------------------------------------
        ToolStripStatusLabel3.Text = "接続×"
        ToolStripStatusLabel3.BackColor = Color.Lavender
        Cn1.Close()
        Table1.Dispose()
        Adapter1.Dispose()
        SQLCm1.Dispose()
        Cn1.Dispose()
        '--------------------------------------------------------------
        Return True
    End Function

    ''' <summary>DELETE</summary>
    ''' <param name="CnAccdb">データベース名</param>
    ''' <param name="TableName">テーブル名</param>
    ''' <param name="where">where文：WHERE (---)</param>
    Private Function MALLDB_CONNECT_DELETE(CnAccdb As String, TableName As String, where As String) As Boolean
        '--------------------------------------------------------------
        Dim Cn1 As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CnAccdb)
        Dim SQLCm1 = Cn1.CreateCommand
        Cn1.Open()
        ToolStripStatusLabel3.Text = "接続○"
        ToolStripStatusLabel3.BackColor = Color.Yellow
        '--------------------------------------------------------------
        If where <> "" Then
            where = " WHERE (" & where & ")"
        End If
        SQLCm1.CommandText = "DELETE FROM [" & TableName & "] " & where
        SQLCm1.ExecuteNonQuery()
        '--------------------------------------------------------------
        ToolStripStatusLabel3.Text = "接続×"
        ToolStripStatusLabel3.BackColor = Color.Lavender
        Cn1.Close()
        SQLCm1.Dispose()
        Cn1.Dispose()
        '--------------------------------------------------------------
        Return True
    End Function

    ''' <summary>接続維持</summary>
    ''' <param name="CnAccdb">データベース名</param>
    Private Sub MALLDB_CONNECT(CnAccdb As String)
        '--------------------------------------------------------------
        Cn0 = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CnAccdb)
        SQLCm0 = Cn0.CreateCommand
        Adapter0 = New OleDbDataAdapter(SQLCm0)
        Cn0.Open()
        ToolStripStatusLabel3.Text = "接続○"
        ToolStripStatusLabel3.BackColor = Color.Yellow
        '--------------------------------------------------------------
    End Sub

    ''' <summary>接続解除</summary>
    Private Sub MALLDB_DISCONNECT()
        '--------------------------------------------------------------
        ToolStripStatusLabel3.Text = "接続×"
        ToolStripStatusLabel3.BackColor = Color.Lavender
        Cn0.Close()
        Adapter0.Dispose()
        SQLCm0.Dispose()
        Cn0.Dispose()
        '--------------------------------------------------------------
    End Sub

    ''' <summary>接続状態でデータ取得1</summary>
    ''' <param name="TableName">テーブル名</param>
    ''' <param name="where">where文：WHERE (---)</param>
    ''' <param name="orderby">orderby文：ORDER BY ---</param>
    ''' <param name="limit">limit文：LIMIT ---</param>
    Private Function MALLDB_GET_SELECT(TableName As String, where As String, Optional orderby As String = "", Optional limit As String = "") As DataTable
        If where <> "" Then
            where = " WHERE (" & where & ")"
        End If
        If orderby <> "" Then
            orderby = " ORDER BY [" & orderby & "]"
        End If
        If limit <> "" Then
            limit = " LIMIT " & limit & ""
        End If
        Dim Table As New DataTable
        SQLCm0.CommandText = "SELECT * FROM [" & TableName & "]" & where & orderby & limit
        Adapter0.Fill(Table)

        Return Table
    End Function

    ''' <summary>接続状態でデータ取得2</summary>
    ''' <param name="TableName">テーブル名</param>
    ''' <param name="getField">フィールド名</param>
    ''' <param name="where">where文：WHERE (---)</param>
    ''' <param name="orderby">orderby文：ORDER BY ---</param>
    Private Function MALLDB_GET_STR(TableName As String, getField As String, where As String, Optional orderby As String = "") As DataTable
        If orderby <> "" Then
            orderby = " ORDER BY " & orderby & ""
        End If
        Dim Table As New DataTable
        SQLCm0.CommandText = "SELECT " & getField & " FROM [" & TableName & "] WHERE (" & where & ")" & orderby
        Adapter0.Fill(Table)

        Return Table
    End Function

    ''' <summary>接続状態でデータ取得3</summary>
    ''' <param name="TableName">テーブル名</param>
    ''' <param name="sql">sql文</param>
    Private Function MALLDB_GET_SQL(TableName As String, sql As String) As DataTable
        Dim Table As New DataTable
        If sql <> "" Then
            SQLCm0.CommandText = "SELECT * FROM [" & TableName & "] " & sql
        Else
            SQLCm0.CommandText = "SELECT * FROM [" & TableName & "]"
        End If
        Adapter0.Fill(Table)

        Return Table
    End Function

    ''' <summary>接続状態でデータ上書き</summary>
    ''' <param name="TableName">テーブル名</param>
    ''' <param name="setField">フィールド名（複数可）</param>
    ''' <param name="setStr">セットする文字（複数可）</param>
    Private Function MALLDB_SET_STR(TableName As String, setField As String, setStr As String)
        Try
            SQLCm0.CommandText = "INSERT INTO [" & TableName & "] (" & setField & ") VALUES (" & setStr & ")"
            SQLCm0.ExecuteNonQuery()
            Return "True"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Private Function MALLDB_UPSERT(TableName As String, where As String, FieldArray As ArrayList, SetArray As ArrayList) As String
        Table0.Rows.Clear()
        Table0.Columns.Clear()
        SQLCm0.CommandText = "SELECT * FROM [" & TableName & "] WHERE " & where
        Adapter0.Fill(Table0)

        If Table0.Rows.Count > 0 Then
            setStr = ""
            For i As Integer = 0 To FieldArray.Count - 1
                If FieldArray(i) <> "" Then
                    If SetArray(i) <> "" Then
                        If Regex.IsMatch(SetArray(i), "^NULL$|^null$") Then
                            setStr &= "[" & FieldArray(i) & "] = '',"
                        Else
                            setStr &= "[" & FieldArray(i) & "] = '" & SetArray(i) & "',"
                        End If
                    Else
                        If CStr(Table0.Rows(0)(FieldArray(i)).ToString) <> "" Then     'フィールド255まで（元128+更新128でもダメなので対策）
                            setStr &= "[" & FieldArray(i) & "] = '',"   '元データに文字が入っている時のみクリア
                        End If
                    End If
                End If
            Next
            setStr = setStr.TrimEnd(",")
            SQLCm0.CommandText = "UPDATE [" & TableName & "] SET " & setStr & " WHERE " & where
            SQLCm0.ExecuteNonQuery()

            Return "update"
        Else
            setField = ""
            setStr = ""
            For i As Integer = 0 To FieldArray.Count - 1
                If FieldArray(i) <> "" Then
                    setField &= "[" & FieldArray(i) & "],"
                    setStr &= "'" & SetArray(i) & "',"
                End If
            Next
            setField = setField.TrimEnd(",")
            setStr = setStr.TrimEnd(",")
            SQLCm0.CommandText = "INSERT INTO [" & TableName & "] (" & setField & ") VALUES (" & setStr & ")"
            SQLCm0.ExecuteNonQuery()

            Return "insert"
        End If
    End Function

    ''' <summary>
    ''' 接続状態でデータ削除
    ''' </summary>
    ''' <param name="TableName">テーブル名</param>
    ''' <param name="where">where文：WHERE (---)</param>
    ''' <returns></returns>
    Private Function MALLDB_DELETE(TableName As String, where As String) As Boolean
        If where <> "" Then
            where = " WHERE (" & where & ")"
        End If
        SQLCm0.CommandText = "DELETE FROM [" & TableName & "] " & where
        SQLCm0.ExecuteNonQuery()

        Return True
    End Function





    ''' <summary>ブラウザフォームに記入</summary>
    ''' <param name="elName">フォーム名</param>
    ''' <param name="elNum">番号</param>
    ''' <param name="setStr">セットする文字</param>
    Private Sub HTMLsetTxt(ByVal elName As String, ByVal elNum As Integer, ByVal setStr As String)
        Try
            Form1.TabBrowser1.SelectedTab.WebBrowser.Document.All.GetElementsByName(elName)(elNum).InnerText = setStr
        Catch ex As Exception
            MsgBox("Error:HTMLsetTxt", MsgBoxStyle.SystemModal)
        End Try
    End Sub

    '****************************************************************************************************
    ' DataGridView選択
    '****************************************************************************************************
    Private Sub 切り取りToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles _
        切り取りToolStripMenuItem1.Click, 切り取りToolStripMenuItem.Click
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        CUTS(FocusDGV)
    End Sub

    Private Sub コピーToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles _
        コピーToolStripMenuItem1.Click, コピーToolStripMenuItem.Click
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        COPYS(FocusDGV, 0)
    End Sub

    '「"」で囲んでコピー
    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles _
        ToolStripMenuItem1.Click, コピーToolStripMenuItem2.Click
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        COPYS(FocusDGV, 1)
    End Sub

    Private Sub 貼り付けToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles _
        貼り付けToolStripMenuItem1.Click, 貼り付けToolStripMenuItem.Click
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        PASTES(FocusDGV, selCell)
    End Sub

    '子項目をクリックした時、SourceControlがnullになるので先に取得
    Private contextMenuSourceControl As Control = Nothing
    Private Sub ContextMenuStrip1_Opened(sender As Object, e As EventArgs) Handles _
            ContextMenuStrip1.Opened
        Dim menu As ContextMenuStrip = DirectCast(sender, ContextMenuStrip)
        'SourceControlプロパティの内容を覚えておく
        contextMenuSourceControl = menu.SourceControl
    End Sub

    Private Sub 上に挿入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles _
        上に挿入ToolStripMenuItem.Click, 上に挿入ToolStripMenuItem1.Click
        If Not contextMenuSourceControl Is Nothing Then
            FocusDGV = contextMenuSourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        ROWSADD(FocusDGV, selCell, 1)
    End Sub

    Private Sub 下に挿入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles _
        下に挿入ToolStripMenuItem.Click, 下に挿入ToolStripMenuItem1.Click
        If Not contextMenuSourceControl Is Nothing Then
            FocusDGV = contextMenuSourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        ROWSADD(FocusDGV, selCell, 0)
    End Sub

    Private Sub 行を選択直下に複製ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles _
        行を選択直下に複製ToolStripMenuItem.Click, 行を選択直下に複製ToolStripMenuItem1.Click
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim dgRow As ArrayList = New ArrayList
        For i As Integer = 0 To FocusDGV.SelectedCells.Count - 1
            If dgRow.Contains(FocusDGV.SelectedCells(i).RowIndex) = False Then
                dgRow.Add(FocusDGV.SelectedCells(i).RowIndex)
            End If
        Next
        dgRow.Sort()
        dgRow.Reverse()

        For r As Integer = 0 To dgRow.Count - 1
            Dim addRow As Integer = dgRow(0) + r + 1
            FocusDGV.Rows.InsertCopy(dgRow(r), addRow)
            For c As Integer = 0 To FocusDGV.ColumnCount - 1
                If FocusDGV.Item(c, dgRow(r)).Value <> "" Then
                    FocusDGV.Item(c, addRow).Value = FocusDGV.Item(c, dgRow(r)).Value
                End If
            Next
        Next
    End Sub

    Private Sub 右に挿入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles _
        右に挿入ToolStripMenuItem.Click, 右に挿入ToolStripMenuItem.Click
        If Not contextMenuSourceControl Is Nothing Then
            FocusDGV = contextMenuSourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        COLSADD(FocusDGV, selCell, 0)
    End Sub

    Private Sub 左に挿入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles _
        左に挿入ToolStripMenuItem.Click, 左に挿入ToolStripMenuItem.Click
        If Not contextMenuSourceControl Is Nothing Then
            FocusDGV = contextMenuSourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        COLSADD(FocusDGV, selCell, 1)
    End Sub

    Private Sub 削除ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles _
        削除ToolStripMenuItem1.Click, 削除ToolStripMenuItem1.Click
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        DELS(FocusDGV, selCell)
    End Sub

    Private Sub 背景色を消すToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles _
        背景色を消すToolStripMenuItem.Click, 背景色を消すToolStripMenuItem1.Click
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        For i As Integer = 0 To FocusDGV.SelectedCells.Count - 1
            FocusDGV.SelectedCells(i).Style.BackColor = DefaultBackColor
        Next
    End Sub

    Private Sub 列選択ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles _
        列選択ToolStripMenuItem1.Click, 列選択ToolStripMenuItem1.Click
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        selChangeFlag = False
        ColSelect(FocusDGV, 0)
        selChangeFlag = True
    End Sub

    Private Sub 行を削除ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles _
        行を削除ToolStripMenuItem1.Click, 行を削除ToolStripMenuItem.Click
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        ROWSCUT(FocusDGV, selCell)
    End Sub

    Private Sub 列を削除ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles _
        列を削除ToolStripMenuItem1.Click, 列を削除ToolStripMenuItem.Click
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        COLSCUT(FocusDGV, selCell)

        If FocusDGV Is DGV16 Then
            COLSCUT(DGV15, selCell)
        End If
    End Sub

    ' ----------------------------------------------------------------------------------
    ' DataGridViewでCtrl-Vキー押下時にクリップボードから貼り付けるための実装↓↓↓↓↓↓
    ' DataGridViewでDelやBackspaceキー押下時にセルの内容を消去するための実装↓↓↓↓↓↓
    ' ----------------------------------------------------------------------------------
    Public Shared ColumnChars() As String
    Private _editingColumn As Integer   '編集中のカラム番号
    Private _editingCtrl As DataGridViewTextBoxEditingControl
    Public Shared innerTextBox As TextBox

    Private Sub DataGridView_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles _
        DGV1.EditingControlShowing, DGV2.EditingControlShowing, DGV3.EditingControlShowing, DGV4.EditingControlShowing, DGV5.EditingControlShowing,
        DGV6.EditingControlShowing, DGV7.EditingControlShowing, DGV8.EditingControlShowing, DGV9.EditingControlShowing, DGV10.EditingControlShowing,
        DGV11.EditingControlShowing, DGV16.EditingControlShowing
        If TypeOf e.Control Is DataGridViewTextBoxEditingControl Then
            Dim dgv As DataGridView = CType(sender, DataGridView)
            innerTextBox = CType(e.Control, TextBox)    '編集のために表示されているコントロールを取得
        End If
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

    Private Sub DataGridView_KeyUp(ByVal sender As Object, ByVal e As Windows.Forms.KeyEventArgs) Handles _
        DGV1.KeyUp, DGV2.KeyUp, DGV3.KeyUp, DGV4.KeyUp, DGV5.KeyUp,
        DGV6.KeyUp, DGV7.KeyUp, DGV8.KeyUp, DGV9.KeyUp, DGV10.KeyUp,
        DGV11.KeyUp, DGV16.KeyUp
        Dim dgv As DataGridView = CType(sender, DataGridView)
        Dim selCell = dgv.SelectedCells

        If e.KeyCode = Keys.Back Then    ' セルの内容を消去
            If DGV1.IsCurrentCellInEditMode Then

            Else
                DELS(dgv, selCell)
            End If
        ElseIf e.Modifiers = Keys.Control And e.KeyCode = Keys.Enter Then '複数セル一括入力
            Dim str As String = dgv.CurrentCell.Value
            For i As Integer = 0 To selCell.Count - 1
                If selCell(i).Value <> str Then
                    selCell(i).Value = str
                    selCell(i).Style.BackColor = Color.Yellow
                End If
            Next
        End If
    End Sub
    ' ----------------------------------------------------------------------------------
    ' DataGridViewでCtrl-Vキー押下時にクリップボードから貼り付けるための実装↑↑↑↑↑↑
    ' DataGridViewでDelやBackspaceキー押下時にセルの内容を消去するための実装↑↑↑↑↑↑
    ' ----------------------------------------------------------------------------------

    Private Sub DELS(ByVal dgv As DataGridView, ByVal selCell As Object)
        For i As Integer = 0 To selCell.Count - 1
            dgv(selCell(i).ColumnIndex, selCell(i).RowIndex).Value = ""
        Next
    End Sub

    ''' <summary>
    ''' 切り取り（汎用版）
    ''' </summary>
    ''' <param name="dgv"></param>
    Private Sub CUTS(ByVal dgv As DataGridView)
        If dgv.GetCellCount(DataGridViewElementStates.Selected) > 0 Then
            If dgv.CurrentCell.IsInEditMode = True Then
                If innerTextBox IsNot Nothing Then
                    innerTextBox.Cut()
                End If
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
                    MessageBox.Show("切り取りできませんでした。")
                End Try
            End If
        End If
    End Sub

    ''' <summary>
    ''' コピー（汎用版）
    ''' </summary>
    ''' <param name="dgv"></param>
    ''' <param name="mode"></param>
    Private Sub COPYS(ByVal dgv As DataGridView, ByVal mode As Integer)
        If dgv.GetCellCount(DataGridViewElementStates.Selected) > 0 Then
            If dgv.CurrentCell.IsInEditMode = True Then
                If innerTextBox IsNot Nothing Then
                    innerTextBox.Copy()
                End If
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

    ''' <summary>
    ''' 貼り付け（汎用版）
    ''' </summary>
    ''' <param name="dgv"></param>
    ''' <param name="selCell"></param>
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
        '    Dim rs As New System.IO.StringReader(clipText)
        '    Dim tfp As New TextFieldParser(rs) With {
        '            .TextFieldType = FileIO.FieldType.Delimited,
        '            .Delimiters = New String() {vbTab, ","},
        '            .HasFieldsEnclosedInQuotes = True
        '        }

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
                tfp.Delimiters = New String() {vbTab} ', ","}
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
            Dim gAddFlag As Boolean = False     '行追加モードの場合は背景黄色にしない
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
                    If y + r = dgv.RowCount - 1 And dgv.AllowUserToAddRows = True Then
                        dgv.RowCount = dgv.RowCount + 1
                        gAddFlag = True
                    End If
                    ' 貼り付け
                    dgv(x + c2, y + r).Value = pststr
                    If Not gAddFlag Then
                        dgv(x + c2, y + r).Style.BackColor = Color.Yellow
                    End If
                    ' ------------------------------------------------------
                    ' 次のセルへ
                    c2 = c2 + 1
                Next
            Next
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

    ''' <summary>
    ''' 列選択（汎用版）
    ''' </summary>
    ''' <param name="dgv"></param>
    ''' <param name="mode">mode=0 1行目を選択する、mode=1 1行目を選択除外</param>
    Private Sub ColSelect(dgv As DataGridView, mode As Integer)
        Dim selCell = dgv.SelectedCells
        Dim selCol As ArrayList = New ArrayList
        For i As Integer = 0 To selCell.Count - 1
            If Not selCol.Contains(selCell(i).ColumnIndex) Then
                selCol.Add(selCell(i).ColumnIndex)
            End If
        Next

        If mode = 0 Then
            For i As Integer = 0 To selCol.Count - 1
                For r As Integer = 0 To dgv.RowCount - 1
                    If Not dgv.Rows(r).IsNewRow Then
                        dgv.Item(selCol(i), r).selected = True
                    End If
                Next
            Next
        Else
            For i As Integer = 0 To selCol.Count - 1
                For r As Integer = 1 To dgv.RowCount - 1
                    If Not dgv.Rows(r).IsNewRow Then
                        dgv.Item(selCol(i), r).selected = True
                    End If
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

        'データ処理を軽くするため画面を一時的に消す
        If delR.Count > 100 Then
            dgv.Visible = False
            Application.DoEvents()
        End If

        For r As Integer = delR.Count - 1 To 0 Step -1
            If Not dgv.Rows(delR(r)).IsNewRow Then
                dgv.Rows.RemoveAt(delR(r))
            End If
        Next
        If delR.Count > 100 Then
            dgv.Visible = True
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


    '--------------------------------------------------------------------------------------------------
    '##################################################################################################
    '
    ' 各店比較
    '
    '##################################################################################################
    '--------------------------------------------------------------------------------------------------

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        KakutenKeisan(True)
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        KakutenKeisan(False)
    End Sub

    Private Sub KakutenKeisan(souryo As Boolean)
        Dim dgv7 As DataGridView = Me.DGV7
        dgv7.Rows.Clear()
        dgv7.Columns.Clear()
        FocusDGV = dgv7

        MALLDB_CONNECT(CnAccdb)

        Dim checkColArray As New ArrayList
        Dim masterArray As New ArrayList
        Dim Table1 As DataTable = MALLDB_GET_SELECT("T_基本情報", "", "表示順")
        For i As Integer = 0 To Table1.Rows.Count - 1
            If Table1.Rows(i)("モール") = "マスタ" Then
                dgv7.Columns.Add("c" & i, "code")
                dgv7.Columns.Add("c" & i, "name")
                Table0 = MALLDB_GET_SELECT("T_" & Table1.Rows(i)("TB名") & "_master", "", "商品コード")
                For r As Integer = 0 To Table0.Rows.Count - 1
                    Dim code As String = Table0.Rows(r)("代表商品コード").ToString.ToLower
                    If Table0.Rows(r)("代表商品コード") = "" Then
                        code = Table0.Rows(r)("商品コード").ToString.ToLower
                    End If
                    Dim shohinmei As String = Table0.Rows(r)("商品名")
                    If Not masterArray.Contains(code) Then
                        dgv7.Rows.Add(code, shohinmei)
                        masterArray.Add(code)
                    End If
                Next
            Else
                '表示切替
                If Not CheckBox1.Checked And Table1.Rows(i)("モール") = "楽天" Then
                    Continue For
                End If
                If Not CheckBox2.Checked And Table1.Rows(i)("モール") = "Yahoo" Then
                    Continue For
                End If
                If Not CheckBox27.Checked And Table1.Rows(i)("モール") = "Amazon" Then
                    Continue For
                End If
                If Not CheckBox12.Checked And Not Regex.IsMatch(Table1.Rows(i)("モール"), "楽天|Yahoo|Amazon") Then
                    Continue For
                End If

                Dim souryoArray As String() = Nothing
                If souryo Then
                    Dim Table2 As DataTable = MALLDB_GET_SELECT("T_その他設定", "[op1]='送料' AND [op2]='" & Table1.Rows(i)("店名") & "'")
                    souryoArray = New String() {Table2.Rows(0)("op3"), Table2.Rows(0)("op4"), Table2.Rows(0)("op5"), Table2.Rows(0)("op6")}
                End If

                Dim settei As String() = Nothing
                Dim haiso As String() = Nothing
                Select Case Table1.Rows(i)("モール")
                    Case "楽天"
                        settei = New String() {"商品管理番号（商品URL）", "販売価格", "表示価格", "倉庫指定", "0"}
                        haiso = New String() {"配送方法セット管理番号", "個別送料"}
                    Case "Yahoo"
                        settei = New String() {"code", "price", "sale-price", "display", "1"}
                        haiso = New String() {"ship-weight", ""}
                    Case "MakeShop"
                        settei = New String() {"独自商品コード", "定価", "販売価格", "商品表示可否", "Y"}
                        haiso = New String() {"送料個別設定", ""}
                    Case "Wowma"
                        settei = New String() {"itemCode", "itemPrice", "itemPrice", "saleStatus", "1"}
                        haiso = New String() {"deliveryMethodName1", ""}
                    Case "Qoo10"
                        settei = New String() {"Seller Code", "Sell Price", "Sell Price", "", ""}
                        haiso = New String() {"Shipping Group No", ""}
                    Case "Amazon"
                        settei = New String() {"出品者SKU", "価格", "価格", "", ""}
                        haiso = New String() {"", ""}
                    Case Else
                        Continue For
                End Select

                'チェック切替
                Dim mode As Integer = 0
                If CheckBox15.Checked And Table1.Rows(i)("モール") = "楽天" Then
                    checkColArray.Add(dgv7.ColumnCount)
                    mode = 0
                End If
                If CheckBox14.Checked And Table1.Rows(i)("モール") = "Yahoo" Then
                    checkColArray.Add(dgv7.ColumnCount)
                    mode = 1
                End If
                If CheckBox26.Checked And Table1.Rows(i)("モール") = "Amazon" Then
                    checkColArray.Add(dgv7.ColumnCount)
                    mode = 1
                End If
                If CheckBox13.Checked And Not Regex.IsMatch(Table1.Rows(i)("モール"), "楽天|Yahoo|Amazon") Then
                    checkColArray.Add(dgv7.ColumnCount)
                    mode = 1
                End If

                dgv7.Columns.Add("c" & i, Table1.Rows(i)("店名"))
                Dim colNum As Integer = dgv7.ColumnCount - 1
                Table0 = MALLDB_GET_SELECT("T_" & Table1.Rows(i)("TB名") & "_item", "")
                For r As Integer = 0 To Table0.Rows.Count - 1
                    Dim code As String = Table0.Rows(r)(settei(0)).ToString.ToLower
                    If masterArray.Contains(code) Then
                        Dim price1 As String = Table0.Rows(r)(settei(1))
                        Dim price2 As String = Table0.Rows(r)(settei(2))

                        If settei(3) = "" Then
                            If IsNumeric(price2) Then
                                If mode = 1 Then
                                    dgv7.Item(colNum, masterArray.IndexOf(code)).Value = price1 & "/" & price2
                                Else
                                    dgv7.Item(colNum, masterArray.IndexOf(code)).Value = price2 & "/" & price1
                                End If
                            Else
                                dgv7.Item(colNum, masterArray.IndexOf(code)).Value = price1
                            End If
                        Else
                            If Table0.Rows(r)(settei(3)) = settei(4) Then   '倉庫確認
                                Dim addSouryo As Integer = 0
                                If souryo Then
                                    If haiso(1) <> "" Then
                                        If IsNumeric(Table0.Rows(r)(haiso(1))) Then
                                            addSouryo = Table0.Rows(r)(haiso(1))
                                            Dim a = 0
                                        Else
                                            For k As Integer = 0 To souryoArray.Length - 1
                                                If souryoArray(k) = "" Then
                                                    Continue For
                                                End If
                                                Dim sa As String() = Split(souryoArray(k), "=")
                                                If Regex.IsMatch(Table0.Rows(r)(haiso(0)), sa(0)) Then
                                                    addSouryo = sa(1)
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                    Else
                                        For k As Integer = 0 To souryoArray.Length - 1
                                            If souryoArray(k) = "" Then
                                                Continue For
                                            End If
                                            Dim sa As String() = Split(souryoArray(k), "=")
                                            If Regex.IsMatch(Table0.Rows(r)(haiso(0)), sa(0)) Then
                                                addSouryo = sa(1)
                                                Exit For
                                            End If
                                        Next
                                    End If
                                End If

                                Dim sStr As String = ""
                                If souryo Then
                                    sStr = "(" & addSouryo & ")"
                                End If
                                If IsNumeric(price2) Then
                                    If mode = 1 Then
                                        dgv7.Item(colNum, masterArray.IndexOf(code)).Value = (CInt(price2) + addSouryo) & sStr & "/" & (CInt(price1) + addSouryo)
                                    Else
                                        dgv7.Item(colNum, masterArray.IndexOf(code)).Value = (CInt(price1) + addSouryo) & sStr & "/" & (CInt(price2) + addSouryo)
                                    End If
                                Else
                                    dgv7.Item(colNum, masterArray.IndexOf(code)).Value = (CInt(price1) + addSouryo) & sStr
                                End If
                            Else
                                dgv7.Item(colNum, masterArray.IndexOf(code)).Value = "倉庫"
                            End If
                        End If
                    End If
                Next
            End If

            Application.DoEvents()
        Next

        MALLDB_DISCONNECT()

        '表を整形
        For r As Integer = 0 To dgv7.Rows.Count - 1
            Dim minK As Integer = -1
            Dim maxK As Integer = -1
            For i As Integer = 0 To checkColArray.Count - 1
                Dim c As Integer = checkColArray(i)
                Dim k As String = Split(dgv7.Item(c, r).Value, "/")(0)
                If k <> "" Then
                    k = Regex.Replace(k, "\(.*\)", "")
                End If
                If IsNumeric(k) Then
                    If minK = -1 Or minK > k Then
                        minK = k
                    End If
                    If maxK = -1 Or maxK < k Then
                        maxK = k
                    End If
                End If
            Next
            For i As Integer = 0 To checkColArray.Count - 1
                Dim c As Integer = checkColArray(i)
                Dim k As String = Split(dgv7.Item(c, r).Value, "/")(0)
                If k <> "" Then
                    k = Regex.Replace(k, "\(.*\)", "")
                End If
                If IsNumeric(k) Then
                    If minK = k Then
                        dgv7.Item(c, r).Style.BackColor = Color.Aqua
                    End If
                    If maxK = k Then
                        dgv7.Item(c, r).Style.BackColor = Color.Orange
                    End If
                End If
            Next
        Next

        MsgBox("検索終了しました", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
    End Sub

    Private Sub DataGridView7_SelectionChanged(sender As Object, e As EventArgs) Handles DGV7.SelectionChanged
        If DGV7.SelectedCells.Count > 0 Then
            TextBox35.Text = DGV7.SelectedCells(0).Value
        End If
    End Sub

    '--------------------------------------------------------------------------------------------------
    '##################################################################################################
    '
    ' カテゴリ検索
    '
    '##################################################################################################
    '--------------------------------------------------------------------------------------------------

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        Dim header As String() = Nothing
        If RadioButton4.Checked Then
            Dim arrayStr As New ArrayList
            arrayStr.Add("全商品ディレクトリID")
            arrayStr.Add("第一階層ディレクトリ")
            arrayStr.Add("第一階層以下ディレクトリ")
            'For i As Integer = 1 To 69
            '    arrayStr.Add("選択可能なタグの分類" & i)
            'Next
            header = DirectCast(arrayStr.ToArray(GetType(String)), String())
            TableName = "T_楽天_ジャンル"
        ElseIf RadioButton5.Checked Then
            header = New String() {"name", "path_name", "relation"}
            TableName = "T_Yahoo_カテゴリー"
        ElseIf RadioButton9.Checked Then
            header = New String() {"カテゴリ"}
            TableName = "T_ヤフオク_カテゴリ"
        End If

        For i As Integer = 0 To header.Length - 1
            If i = 0 Then
                where = "[" & header(i) & "] LIKE '%" & TextBox17.Text & "%'"
            Else
                where &= " OR " & "[" & header(i) & "] LIKE '%" & TextBox17.Text & "%'"
            End If
        Next

        '=================================================
        Dim Table As DataTable = MALLDB_CONNECT_SELECT(CnAccdb_genre, TableName, where)
        '=================================================

        If Table.Rows.Count > 0 Then
            Table_to_DGV(DGV11, Table)
        Else
            DGV11.Rows.Clear()
            DGV11.Columns.Clear()
            DGV11.Columns.Add("1", "1")
            DGV11.Rows.Add("検索結果がありません")
        End If

        '=================================================
        Table.Dispose()
        '=================================================
    End Sub

    Private Sub CategoryCSVGet(filename)
        Dim csvArray As ArrayList = TM_CSV_READ(filename)(0)
        Dim csvHeaderArray As String() = Split(csvArray(0).ToString, "|=|")

        If DGV11.RowCount > 0 Then DGV11.Rows.Clear()
        If DGV11.ColumnCount > 0 Then DGV11.Columns.Clear()
        For c As Integer = 0 To csvHeaderArray.Length - 1
            DGV11.Columns.Add(c, csvHeaderArray(c))
        Next
        For r As Integer = 1 To csvArray.Count - 1
            DGV11.Rows.Add()
            Dim cArray As String() = Split(csvArray(r).ToString, "|=|")
            For c As Integer = 0 To cArray.Length - 1
                DGV11.Item(c, r - 1).Value = cArray(c)
            Next
        Next
    End Sub

    'カテゴリマスタ保存
    Private Sub Button59_Click(sender As Object, e As EventArgs) Handles Button59.Click
        If DGV11.RowCount < 1000 Then
            MsgBox("正しいカテゴリマスタを読み込んでいますか？")
            Exit Sub
        End If

        MALLDB_CONNECT(CnAccdb_genre)

        'テーブル名を決める
        If RadioButton4.Checked Then
            TableName = "T_楽天_ジャンル"
        ElseIf RadioButton5.Checked Then
            TableName = "T_Yahoo_カテゴリー"
        ElseIf RadioButton9.Checked Then
            TableName = "T_ヤフオク_カテゴリ"
        Else
            Exit Sub
        End If

        'テーブルが存在しているか確認する
        'New Object() {TABLE_CATALOG, TABLE_SCHEMA, TABLE_NAME, TABLE_TYPE}
        'Dim schemaTable As DataTable = Cn0.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, TableDst, "TABLE"})
        Dim schemaTable As DataTable = MALLDB_TABLE_EXISTS(Nothing, Nothing, TableName, "TABLE")

        If schemaTable.Rows.Count <> 0 Then
            '存在するテーブルを削除する
            If schemaTable.Rows.Count > 0 Then
                LIST4VIEW("テーブル削除", "s")
                Dim SQL1 As String = "DROP TABLE [" & TableName & "] "
                SQLCm0.CommandText = SQL1
                SQLCm0.ExecuteNonQuery()
            End If
        End If
        '----------
        schemaTable.Clear()
        '----------

        Dim createHeaderList As String = ""
        For c As Integer = 0 To DGV11.ColumnCount - 1
            If RadioButton9.Checked Then
                createHeaderList &= "" & DGV11.Columns(c).HeaderText & " MEMO,"
            Else
                createHeaderList &= "[" & DGV11.Columns(c).HeaderText & "] MEMO,"
            End If
            setField &= "" & DGV11.Columns(c).HeaderText & ","
        Next
        createHeaderList = createHeaderList.TrimEnd(",")
        setField = setField.TrimEnd(",")

        'テーブルを作成する
        Dim SQL As String = "CREATE TABLE [" & TableName & "] "
        SQL &= "(" & createHeaderList & ")"
        SQLCm0.CommandText = SQL
        SQLCm0.ExecuteNonQuery()

        'データ書き込み
        ToolStripProgressBar1.Maximum = DGV11.RowCount
        ToolStripProgressBar1.Value = 0
        For r As Integer = 0 To DGV11.RowCount - 1
            setStr = ""
            For c As Integer = 0 To DGV11.ColumnCount - 1
                setStr &= "'" & DGV11.Item(c, r).Value & "',"
            Next
            setStr = setStr.TrimEnd(",")
            MALLDB_SET_STR(TableName, setField, setStr)

            ToolStripProgressBar1.Value += 1
            Application.DoEvents()
        Next
        ToolStripProgressBar1.Value = 0

        MALLDB_DISCONNECT()

        Dim fileSrc As String = Path.GetDirectoryName(Form1.appPath) & "\db\category.accdb"
        Dim fileDst As String = Form1.サーバーToolStripMenuItem.Text & "\update\db\category.accdb"
        File.Copy(fileSrc, fileDst, True)

        MsgBox("終了しました")
    End Sub

    Private Sub Button58_Click(sender As Object, e As EventArgs) Handles Button58.Click
        Dim fileSrc As String = Path.GetDirectoryName(Form1.appPath) & "\db\category.accdb"
        Dim fileDst As String = Form1.サーバーToolStripMenuItem.Text & "\update\db\category.accdb"
        File.Copy(fileDst, fileSrc, True)
        MsgBox("更新しました")
    End Sub

    Private Sub TextBox17_KeyUp(sender As Object, e As KeyEventArgs) Handles TextBox17.KeyUp
        If e.KeyCode = Keys.Enter Then
            Button26.PerformClick()
        End If
    End Sub






    '=====================================
    '-------------------------------------
    ' Amazon
    '-------------------------------------
    '=====================================
#Region "Amazon"

    Private Sub DGV20_Header()
        'Amazon入力画面
        Dim aPath As String = Path.GetDirectoryName(Form1.appPath) & "\config\AmazonDefault.txt"
        Dim dgv20headerArray As String() = File.ReadAllLines(aPath, Encoding.GetEncoding("SHIFT-JIS"))
        DGV20.Rows.Clear()
        DGV20.Columns.Clear()
        For c As Integer = 0 To dgv20headerArray.Length - 1
            DGV20.Columns.Add(c, dgv20headerArray(c))
        Next
    End Sub

    Private Sub AmazonCsvGet(filename)
        Dim csvArray As ArrayList = TM_CSV_READ(filename)(0)
        Dim csvHeaderArray As String() = Split(csvArray(0).ToString, "|=|")

        DGV20.Rows.Clear()
        DGV20.Columns.Clear()
        For c As Integer = 0 To csvHeaderArray.Length - 1
            DGV20.Columns.Add(c, csvHeaderArray(c))
        Next
        For r As Integer = 1 To csvArray.Count - 1
            DGV20.Rows.Add()
            Dim cArray As String() = Split(csvArray(r).ToString, "|=|")
            For c As Integer = 0 To cArray.Length - 1
                DGV20.Item(c, r).Value = cArray(c)
            Next
        Next
    End Sub

    Private Sub DataGridView20_SelectionChanged(sender As Object, e As EventArgs) Handles DGV20.SelectionChanged
        If DGV20.SelectedCells.Count > 0 Then
            TextBox41.Text = DGV20.SelectedCells(0).Value
        End If
    End Sub

    Private Sub TextBox41_Validated(sender As Object, e As EventArgs) Handles TextBox41.Validated
        If Not DGV20.Rows(0).IsNewRow Then
            DGV20.SelectedCells(0).Value = TextBox41.Text
        End If
    End Sub

    '分析
    Private Sub Button54_Click(sender As Object, e As EventArgs) Handles Button54.Click
        Dim aPath As String = Path.GetDirectoryName(Form1.appPath) & "\config\Amazon入力カテゴリ対応表.txt"
        Dim cArray As String() = File.ReadAllLines(aPath, Encoding.GetEncoding("SHIFT-JIS"))

        For r As Integer = 0 To DGV21.RowCount - 1
            DGV21.Rows(r).Visible = True
            DGV21.Item(1, r).Value = 0
            DGV21.Item(1, r).Style.BackColor = Color.Empty
        Next

        Dim dH20 As ArrayList = TM_HEADER_GET(DGV20)
        If Not dH20.Contains("分析") Then
            DGV20.Columns.Add("bunseki", "分析")
            dH20.Add("分析")
        End If

        For r As Integer = 0 To DGV20.RowCount - 1
            Dim countFlag As Boolean = False
            Dim node As String = DGV20.Item(dH20.IndexOf("ノードカテゴリ"), r).Value
            If node <> "" Then
                node = Regex.Replace(node, "【|】", "")
                For i As Integer = 0 To cArray.Length - 1
                    Dim cA As String() = Split(cArray(i), ",")
                    For k As Integer = 0 To cA.Length - 1
                        If cA(k) <> "" Then
                            If Regex.IsMatch(node, "^" & cA(k)) Then
                                DGV21.Item(1, i).Value += 1
                                DGV21.Item(1, i).Style.BackColor = Color.Yellow
                                DGV20.Item(dH20.IndexOf("分析"), r).Value = DGV21.Item(0, i).Value
                                countFlag = True
                                Exit For
                            End If
                        End If
                    Next

                    If countFlag Then
                        Exit For
                    End If
                Next
            End If
        Next

        For r As Integer = 0 To DGV21.RowCount - 1
            If DGV21.Item(1, r).Value = 0 Then
                DGV21.Rows(r).Visible = False
            End If
        Next
    End Sub

    Private Sub AmazonTemplateXlsGet(fName As String, Optional mode As String = "")
        Dim fn As String = Path.GetFileNameWithoutExtension(fName)
        For i As Integer = 0 To ComboBox13.Items.Count - 1
            If InStr(ComboBox13.Items(i), fn) > 0 Then
                ComboBox13.SelectedIndex = i
                Exit For
            End If
        Next

        Dim itemList_web As ArrayList = TM_XLS_READ(fName, "テンプレート")
        DGV12.Rows.Clear()
        DGV12.Columns.Clear()

        Dim kajougakiFlag As Boolean = True
        Dim colPlus As Integer = 0
        'If itemList_web(2).contains("bullet_point") Then
        '    colPlus += 5
        '    For i As Integer = 1 To 10      'テンプレートが箇条書き11～15に間違っている場合があるので補正
        '        itemList_web(1) = Replace(itemList_web(1), "商品説明の箇条書き" & i + 10, "商品説明の箇条書き" & i)
        '    Next
        '    kajougakiFlag = True
        'Else
        '    colPlus += 10
        '    kajougakiFlag = False
        'End If
        For i As Integer = 0 To itemList_web.Count - 1
            Dim colArray As String() = Split(itemList_web(i), "|=|")
            If DGV12.ColumnCount < colArray.Length Then
                Dim startNum As Integer = DGV12.ColumnCount
                Dim endNum As Integer = colArray.Length - 1
                For c As Integer = startNum To endNum + colPlus
                    DGV12.Columns.Add(c, c)
                Next
            End If
            DGV12.Rows.Add(colArray)
            Dim counter As Integer = 0
            For c As Integer = 0 To colArray.Length - 1
                If counter <= DGV12.ColumnCount - 1 Then
                    DGV12.Item(counter, DGV12.RowCount - 1).Value = colArray(c)
                    counter += 1
                End If
            Next
        Next


        'For i As Integer = 0 To itemList_web.Count - 1
        '    Dim colArray As String() = Split(itemList_web(i), "|=|")
        '    If dgv12.ColumnCount < colArray.Length Then
        '        Dim startNum As Integer = dgv12.ColumnCount
        '        Dim endNum As Integer = colArray.Length - 1
        '        For c As Integer = startNum To endNum + colPlus
        '            dgv12.Columns.Add(c, c)
        '        Next
        '    End If
        '    dgv12.Rows.Add(colArray)
        '    Dim counter As Integer = 0
        '    For c As Integer = 0 To colArray.Length - 1
        '        If colArray(c) = "商品説明の箇条書き5" Then
        '            dgv12.Item(counter, dgv12.RowCount - 1).Value = colArray(c)
        '            counter += 1
        '            For k As Integer = 6 To 10
        '                dgv12.Item(counter, dgv12.RowCount - 1).Value = "商品説明の箇条書き" & k
        '                counter += 1
        '            Next
        '        ElseIf colArray(c) = "商品の仕様5" Then
        '            dgv12.Item(counter, dgv12.RowCount - 1).Value = colArray(c)
        '            counter += 1
        '            For k As Integer = 6 To 10
        '                dgv12.Item(counter, dgv12.RowCount - 1).Value = "商品の仕様" & k
        '                counter += 1
        '            Next
        '        ElseIf colArray(c) = "bullet_point5" Then
        '            dgv12.Item(counter, dgv12.RowCount - 1).Value = colArray(c)
        '            counter += 1
        '            For k As Integer = 6 To 10
        '                dgv12.Item(counter, dgv12.RowCount - 1).Value = "bullet_point" & k
        '                counter += 1
        '            Next
        '        ElseIf kajougakiFlag = False And colArray(c) = "検索キーワード" Then
        '            For k As Integer = 1 To 10
        '                dgv12.Item(counter, dgv12.RowCount - 1).Value = "商品説明の箇条書き" & k
        '                counter += 1
        '            Next
        '            dgv12.Item(counter, dgv12.RowCount - 1).Value = colArray(c)
        '            counter += 1
        '        ElseIf kajougakiFlag = False And colArray(c) = "generic_keywords" Then
        '            For k As Integer = 1 To 10
        '                dgv12.Item(counter, dgv12.RowCount - 1).Value = "bullet_point" & k
        '                counter += 1
        '            Next
        '            dgv12.Item(counter, dgv12.RowCount - 1).Value = colArray(c)
        '            counter += 1
        '        ElseIf InStr(colArray(c), "商品検索情報") > 0 Then
        '            dgv12.Item(counter, dgv12.RowCount - 1).Value = colArray(c)
        '            counter += 1
        '            If kajougakiFlag Then
        '                '箇条書き6～10を増やす
        '                For k As Integer = 6 To 10
        '                    dgv12.Item(counter, dgv12.RowCount - 1).Value = ""
        '                    counter += 1
        '                Next
        '            Else
        '                '箇条書きが無いテンプレートは1～10を増やす
        '                For k As Integer = 1 To 10
        '                    dgv12.Item(counter, dgv12.RowCount - 1).Value = ""
        '                    counter += 1
        '                Next
        '            End If
        '        Else
        '            If counter <= dgv12.ColumnCount - 1 Then
        '                dgv12.Item(counter, dgv12.RowCount - 1).Value = colArray(c)
        '                counter += 1
        '            End If
        '        End If
        '    Next
        'Next

        itemList_web = TM_XLS_READ(fName, "推奨値")
        DGV13.Rows.Clear()
        DGV13.Columns.Clear()
        DGV13.Columns.Add(0, 0)
        For c As Integer = 0 To itemList_web.Count - 1
            Dim rowArray As String() = Split(itemList_web(c), "|=|")
            For r As Integer = 0 To rowArray.Length - 1
                If c = 0 Then
                    DGV13.Rows.Add()
                End If
                If rowArray(r) <> "" Then
                    If c = 0 Then
                        DGV13.Item(0, r).Value = rowArray(r)
                    Else
                        DGV13.Item(0, r).Value &= "|" & rowArray(r)
                    End If
                End If
            Next
        Next
        '推奨値のカラム数が多すぎ
        'For i As Integer = 0 To itemList_web.Count - 1
        '    Dim colArray As String() = Split(itemList_web(i), "|=|")
        '    If dgv13.ColumnCount < colArray.Length Then
        '        Dim startNum As Integer = dgv13.ColumnCount
        '        Dim endNum As Integer = colArray.Length - 1
        '        For c As Integer = startNum To endNum
        '            dgv13.Columns.Add(c, c)
        '        Next
        '    End If
        '    dgv13.Rows.Add(colArray)
        'Next

        itemList_web = TM_XLS_READ(fName, "データ定義")
        DGV17.Rows.Clear()
        Dim num As Integer = 0
        For i As Integer = 0 To itemList_web.Count - 1
            Dim colArray As String() = Split(itemList_web(i), "|=|")
            If colArray(0) = "" Then
                DGV17.Rows.Add()
                DGV17.Item(0, num).Value = colArray(2)
                DGV17.Item(1, num).Value = colArray(1)
                Dim str As String = "(例)" & colArray(5) & vbLf & colArray(3) & vbLf & colArray(4)
                DGV17.Item(2, num).Value = Replace(str, vbLf, "[\n]")
                DGV17.Item(3, num).Value = colArray(6)
                DGV17.Item(4, num).Value = colArray(5)
                num += 1
            End If
        Next
    End Sub

    Private Sub AmazonItemXlsGet(fName As String)
        If ComboBox14.SelectedIndex <= 0 Then
            MsgBox("テンプレートを先に読み込んでください")
            Exit Sub
        End If

        Dim dH15 As ArrayList = TM_HEADER_1ROW_GET(DGV15, 2)
        Dim itemList_web As ArrayList = TM_XLS_READ(fName, "テンプレート")
        DGV16.Rows.Clear()

        Dim itemListHeader As String() = Nothing
        For i As Integer = 0 To itemList_web.Count - 1
            If i = 0 Or i = 1 Then
                Continue For
            ElseIf i = 2 Then   '読み込んだエクセルのヘッダー
                itemListHeader = Split(itemList_web(i), "|=|")
            Else
                If itemList_web(i) <> "" Then
                    DGV16.Rows.Add()
                    Dim colArray As String() = Split(itemList_web(i), "|=|")
                    For k As Integer = 0 To colArray.Length - 1
                        If dH15.Contains(itemListHeader(k)) Then
                            DGV16.Item(dH15.IndexOf(itemListHeader(k)), DGV16.RowCount - 2).Value = colArray(k)
                        End If
                    Next
                End If
            End If
        Next

        MsgBox("読み込み完了しました")
    End Sub

    Private Sub DataGridView12_SelectionChanged(sender As Object, e As EventArgs) Handles DGV12.SelectionChanged
        If DGV12.RowCount > 0 And DGV12.SelectedCells.Count > 0 Then
            TextBox21.Text = CStr(DGV12.SelectedCells(0).Value)
        End If
    End Sub

    '保存
    Private Sub Button29_Click(sender As Object, e As EventArgs) Handles Button29.Click
        Dim ueArray As New ArrayList
        For r As Integer = 0 To 2
            Dim str As String = ""
            For c As Integer = 0 To DGV12.ColumnCount - 1
                If c = 0 Then
                    str = DGV12.Item(c, r).Value
                Else
                    str &= "|" & DGV12.Item(c, r).Value
                End If
            Next
            ueArray.Add(str)
        Next

        Dim sitaStr As String = ""
        For r As Integer = 0 To DGV13.RowCount - 1
            Dim str As String = ""
            For c As Integer = 0 To DGV13.ColumnCount - 1
                If c = 0 Then
                    str = DGV13.Item(c, r).Value
                Else
                    str &= "|" & DGV13.Item(c, r).Value
                End If
            Next
            sitaStr &= str & vbCrLf
        Next
        sitaStr = Replace(sitaStr, "'", "[c39]")
        sitaStr = Replace(sitaStr, """", "[c34]")

        Dim setumeiStr As String = ""
        Dim hissuStr As String = ""
        Dim formatStr As String = ""
        For r As Integer = 0 To DGV17.RowCount - 1
            Dim str As String = ""
            For c As Integer = 0 To 2
                If c = 0 Then
                    str = DGV17.Item(c, r).Value
                Else
                    str &= "|-|" & DGV17.Item(c, r).Value
                End If
            Next
            setumeiStr &= str & "|=|"
            hissuStr &= DGV17.Item(3, r).Value & "[\n]"
            formatStr &= DGV17.Item(4, r).Value & "[\n]"
        Next
        setumeiStr = Replace(setumeiStr, "'", "[c39]")
        setumeiStr = Replace(setumeiStr, """", "[c34]")

        Dim setStr As String = ""
        Dim FieldArray As String() = {"設定1", "設定2", "設定3"}
        For i As Integer = 0 To FieldArray.Length - 1
            setStr &= "[" & FieldArray(i) & "] = '" & ueArray(i) & "',"
        Next
        setStr &= "[設定4]='" & sitaStr & "',"
        setStr &= "[説明]='" & setumeiStr & "',"
        setStr &= "[必須]='" & hissuStr & "',"
        setStr &= "[フォーマット]='" & formatStr & "'"

        If ComboBox13.SelectedIndex >= 0 Then
            Dim fName As String = Split(ComboBox13.SelectedItem, ",")(1)
            fName = Split(fName, "|")(0)
            Dim where As String = "[名称]='" & fName & "'"
            TM_DB_CONNECT_UPDATE(CnAccdb, "T_設定_Amazon", setStr, where)
            MsgBox("保存しました")
        Else
            MsgBox("ファイルが認識できていません。テンプレートを選択してください。")
        End If
    End Sub

    '読込
    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click
        If ComboBox13.SelectedIndex >= 0 Then
            Dim tempName As String() = Split(ComboBox13.SelectedItem, ",")
            tempName(1) = Split(tempName(1), "|")(0)
            Dim data As DataTable = TM_DB_CONNECT_SELECT(CnAccdb, "T_設定_Amazon", "[名称]='" & tempName(1) & "'")

            DGV12.Rows.Clear()
            DGV12.Columns.Clear()

            Dim strArray As String() = Nothing
            strArray = Split(data.Rows(0)("設定1"), "|")
            For i As Integer = 0 To strArray.Length - 1
                DGV12.Columns.Add(i, i)
            Next
            DGV12.Rows.Add(strArray)

            strArray = Split(data.Rows(0)("設定2"), "|")
            DGV12.Rows.Add(strArray)

            strArray = Split(data.Rows(0)("設定3"), "|")
            DGV12.Rows.Add(strArray)

            '----------------------------

            DGV13.Rows.Clear()
            DGV13.Columns.Clear()

            strArray = Split(data.Rows(0)("設定4"), vbLf)
            DGV13.Columns.Add(0, 0)
            For i As Integer = 0 To strArray.Length - 1
                If strArray(i) <> "" Then
                    strArray(i) = Replace(strArray(i), "[c39]", "'")
                    strArray(i) = Replace(strArray(i), "[c34]", """")
                    DGV13.Rows.Add(strArray(i))
                End If
            Next

            '----------------------------
            DGV17.Rows.Clear()

            Dim strArray2 As String() = Nothing
            strArray = Split(data.Rows(0)("説明"), "|=|")
            Dim strArray3 As String() = Split(data.Rows(0)("必須"), "[\n]")
            Dim strArray4 As String() = Split(data.Rows(0)("フォーマット"), "[\n]")
            For i As Integer = 0 To strArray.Length - 1
                If strArray(i) <> "" Then
                    strArray(i) = Replace(strArray(i), "[c39]", "'")
                    strArray(i) = Replace(strArray(i), "[c34]", """")
                    strArray2 = Split(strArray(i), "|-|")
                    DGV17.Rows.Add(strArray2)
                    DGV17.Item(3, i).Value = strArray3(i)
                    DGV17.Item(4, i).Value = strArray4(i)
                End If
            Next

        End If
    End Sub

    'ブラウズノードリスト
    Private Sub NodeXlsGet(fName As String)
        TextBox23.Text = "読み込み中..."
        Application.DoEvents()
        Dim itemList_web As ArrayList = TM_XLS_READ(fName, "分類")
        TextBox23.Text = "データを展開しています..."
        Application.DoEvents()
        DGV14.Rows.Clear()
        For i As Integer = 2 To itemList_web.Count - 1
            Dim colArray As String() = Split(itemList_web(i), "|=|")
            DGV14.Rows.Add()
            DGV14.Item(0, i - 2).Value = colArray(1)
            DGV14.Item(1, i - 2).Value = colArray(2)
        Next
        Button30.Enabled = True
        TextBox23.Text = ""
    End Sub

    Private Sub DataGridView14_SelectionChanged(sender As Object, e As EventArgs) Handles DGV14.SelectionChanged
        If DGV14.RowCount > 0 And DGV14.SelectedCells.Count > 0 Then
            TextBox23.Text = CStr(DGV14.SelectedCells(0).Value)
        End If
    End Sub

    '検索
    Private Sub Button32_Click(sender As Object, e As EventArgs) Handles Button32.Click
        TableName = "T_Amazonノード"
        where = "[ノードID] LIKE '%" & TextBox22.Text & "%'"
        where &= " OR " & "[分類] LIKE '%" & TextBox22.Text & "%'"

        '=================================================
        Dim Table As DataTable = MALLDB_CONNECT_SELECT(CnAccdb_amazon, TableName, where)
        '=================================================

        If Table.Rows.Count > 0 Then
            Table_to_DGV(DGV14, Table)
        Else
            DGV14.Rows.Clear()
            DGV14.Columns.Clear()
            DGV14.Columns.Add("1", "1")
            DGV14.Rows.Add("検索結果がありません")
        End If

        '=================================================
        Table.Dispose()
        '=================================================
    End Sub

    Private Sub TextBox22_KeyUp(sender As Object, e As KeyEventArgs) Handles TextBox22.KeyUp
        If e.KeyCode = Keys.Enter Then
            Button32.PerformClick()
        End If
    End Sub

    'データベース更新
    Private Sub Button30_Click(sender As Object, e As EventArgs) Handles Button30.Click
        MALLDB_CONNECT(CnAccdb_amazon)

        'テーブルを削除する
        Dim SQL1 As String = "DROP TABLE [T_Amazonノード]"
        SQLCm0.CommandText = SQL1
        SQLCm0.ExecuteNonQuery()

        'テーブルを作成する
        Dim SQL As String = "CREATE TABLE [T_Amazonノード] "
        SQL &= "(ノードID MEMO, 分類 MEMO)"
        SQLCm0.CommandText = SQL
        SQLCm0.ExecuteNonQuery()

        ToolStripProgressBar1.Maximum = DGV14.RowCount
        ToolStripProgressBar1.Value = 0

        Dim setField As String = "[ノードID],[分類]"
        Dim setStr As String = ""
        For r As Integer = 0 To DGV14.RowCount - 1
            setStr = "'" & DGV14.Item(0, r).Value & "','" & DGV14.Item(1, r).Value & "'"
            MALLDB_SET_STR("T_Amazonノード", setField, setStr)
            ToolStripProgressBar1.Value += 1
        Next

        ToolStripProgressBar1.Value = 0

        MALLDB_DISCONNECT()
        MsgBox("保存しました")
    End Sub

    'JANコード読込
    Private Sub Button33_Click(sender As Object, e As EventArgs) Handles Button33.Click
        TableName = "T_" & ComboBox15.SelectedItem & "JANコード"
        where = ""
        '=================================================
        Dim Table As DataTable = MALLDB_CONNECT_SELECT(CnAccdb_amazon, TableName, where)
        '=================================================

        If Table.Rows.Count > 0 Then
            Table_to_DGV(DGV14, Table)
            For r As Integer = 0 To DGV14.RowCount - 1
                DGV14.Item(2, r).Value = Calc_Check_Digit(DGV14.Item(0, r).Value & DGV14.Item(1, r).Value)
                DGV14.Item(3, r).Value = DGV14.Item(0, r).Value & DGV14.Item(1, r).Value & DGV14.Item(2, r).Value
            Next
        Else
            DGV14.Rows.Clear()
            DGV14.Columns.Clear()
            DGV14.Columns.Add("1", "1")
            DGV14.Rows.Add("検索結果がありません")
        End If

        '=================================================
        Table.Dispose()
        '=================================================
    End Sub

    'JANコード保存
    Private Sub Button34_Click(sender As Object, e As EventArgs) Handles Button34.Click
        MALLDB_CONNECT(CnAccdb_amazon)

        'テーブルを削除する
        TableName = "T_" & ComboBox15.SelectedItem & "JANコード"
        Dim SQL1 As String = "DROP TABLE [" & TableName & "]"
        SQLCm0.CommandText = SQL1
        SQLCm0.ExecuteNonQuery()

        'テーブルを作成する
        Dim SQL As String = "CREATE TABLE [" & TableName & "] "
        SQL &= "(企業番号 MEMO, 商品番号 MEMO, cデジット MEMO, JANコード MEMO, 商品コード MEMO, メモ MEMO)"
        SQLCm0.CommandText = SQL
        SQLCm0.ExecuteNonQuery()

        ToolStripProgressBar1.Maximum = DGV14.RowCount
        ToolStripProgressBar1.Value = 0

        Dim setField As String = "[企業番号],[商品番号],[cデジット],[JANコード],[商品コード],[メモ]"
        For r As Integer = 0 To DGV14.RowCount - 1
            Dim setStr As String = ""
            For c As Integer = 0 To DGV14.ColumnCount - 1
                setStr &= "'" & DGV14.Item(c, r).Value & "',"
            Next
            setStr = setStr.TrimEnd(",")
            MALLDB_SET_STR(TableName, setField, setStr)
            ToolStripProgressBar1.Value += 1
        Next

        ToolStripProgressBar1.Value = 0

        MALLDB_DISCONNECT()
        MsgBox("保存しました")
    End Sub

    'JAN増加実行
    Private Sub Button35_Click(sender As Object, e As EventArgs) Handles Button35.Click
        Dim dH14 As ArrayList = TM_HEADER_GET(DGV14)
        Dim dEnd As Integer = 0
        If ComboBox16.SelectedItem = "1増やす" Then
            dEnd = 1
        ElseIf ComboBox16.SelectedItem = "10増やす" Then
            dEnd = 10
        ElseIf ComboBox16.SelectedItem = "100増やす" Then
            dEnd = 100
        ElseIf ComboBox16.SelectedItem = "1000増やす" Then
            dEnd = 1000
        End If

        For i As Integer = 0 To dEnd - 1
            DGV14.Rows.Add()
            Dim num As Integer = DGV14.RowCount - 1
            DGV14.Item(dH14.IndexOf("企業番号"), num).Value = DGV14.Item(dH14.IndexOf("企業番号"), num - 1).Value
            Dim sNo As String = DGV14.Item(dH14.IndexOf("商品番号"), num - 1).Value
            sNo = CInt(sNo) + 1
            If sNo > 999 Then
                sNo = 0
            End If
            sNo = sNo.PadLeft(3, "0"c)
            DGV14.Item(dH14.IndexOf("商品番号"), num).Value = sNo
            DGV14.Item(dH14.IndexOf("cデジット"), num).Value = Calc_Check_Digit(DGV14.Item(dH14.IndexOf("企業番号"), num).Value & DGV14.Item(dH14.IndexOf("商品番号"), num).Value)
            DGV14.Item(dH14.IndexOf("JANコード"), num).Value = DGV14.Item(dH14.IndexOf("企業番号"), num).Value & DGV14.Item(dH14.IndexOf("商品番号"), num).Value & DGV14.Item(dH14.IndexOf("cデジット"), num).Value
        Next
    End Sub

    Dim BeforeColor16 As Color = Color.Empty
    Private Sub DataGridView16_SelectionChanged(sender As DataGridView, e As EventArgs) Handles DGV15.SelectionChanged, DGV16.SelectionChanged
        If sender.RowCount > 0 And sender.SelectedCells.Count > 0 Then
            TextBox24.Text = CStr(sender.SelectedCells(0).Value)

            If sender Is DGV16 Then
                '選択行の背景色を変更する
                If sender.SelectedCells.Count > 50 Then
                    Exit Sub
                End If
                For r As Integer = 0 To sender.RowCount - 1
                    sender.Rows(r).DefaultCellStyle.BackColor = BeforeColor16
                Next
                Dim selRow As ArrayList = New ArrayList
                For i As Integer = 0 To sender.SelectedCells.Count - 1
                    If Not selRow.Contains(sender.SelectedCells(i).RowIndex) Then
                        selRow.Add(sender.SelectedCells(i).RowIndex)
                    End If
                Next
                For i As Integer = 0 To selRow.Count - 1
                    BeforeColor16 = sender.Rows(selRow(i)).DefaultCellStyle.BackColor
                    sender.Rows(selRow(i)).DefaultCellStyle.BackColor = Color.LightCyan
                Next

                If Not DGV16.SelectedCells(0).Value Is Nothing Then
                    ToolStripStatusLabel1.Text = "[文字数："
                    ToolStripStatusLabel1.Text &= DGV16.SelectedCells(0).Value.ToString.Length    '文字数カウント
                    ToolStripStatusLabel1.Text &= "(" & SJIS.GetByteCount(DGV16.SelectedCells(0).Value.ToString) & "byt)"
                    ToolStripStatusLabel1.Text &= "/" & CInt(SJIS.GetByteCount(DGV16.SelectedCells(0).Value.ToString)) + CInt(DGV16.SelectedCells(0).Value.ToString.Length) & "byt)"
                    ToolStripStatusLabel1.Text &= "]"
                End If
            End If

            'データ定義呼び出し
            If DGV15.RowCount >= 2 Then
                Dim selCol As Integer = sender.SelectedCells(0).ColumnIndex
                Dim selHeader As String = DGV15.Item(selCol, 2).Value
                selHeader = Regex.Replace(selHeader, "[0-9]{1,3}$", "")
                If dataTeigiDB IsNot Nothing Then
                    For i As Integer = 0 To dataTeigiDB.Length - 1
                        If InStr(dataTeigiDB(i), "|-|") = 0 Then
                            Continue For
                        End If
                        Dim dtArray As String = Split(dataTeigiDB(i), "|-|")(1)
                        If Regex.IsMatch(dtArray, "^" & selHeader) Then
                            'RichTextBox2.Text = dataTeigiDB(i)
                            Dim setumei As String = Split(dataTeigiDB(i), "|-|")(2)
                            setumei = Replace(setumei, "[\n]", vbLf)
                            setumei = Replace(setumei, "[c39]", "'")
                            setumei = Replace(setumei, "[c34]", """")
                            RichTextBox2.Text = setumei
                            Exit For
                        End If
                        If i = dataTeigiDB.Length - 1 Then
                            RichTextBox2.Text = ""
                        End If
                    Next
                End If
            End If
        End If
    End Sub

    Private Sub DataGridView19_SelectionChanged(sender As DataGridView, e As EventArgs) Handles DGV19.SelectionChanged
        If sender.RowCount > 0 And sender.SelectedCells.Count > 0 Then
            TextBox24.Text = CStr(sender.SelectedCells(0).Value)
        End If
    End Sub

    Private Sub TextBox24_Validated(sender As Object, e As EventArgs) Handles TextBox24.Validated
        For i As Integer = 0 To DGV16.SelectedCells.Count - 1
            DGV16.SelectedCells(i).Value = TextBox24.Text
            CheckDGV16("one", DGV16.SelectedCells(i).ColumnIndex, DGV16.SelectedCells(i).RowIndex)
            If DGV16.SelectedCells(i).Style.BackColor = Color.Empty Then
                DGV16.SelectedCells(i).Style.BackColor = Color.Yellow
            End If
        Next
    End Sub

    'カテゴリーテンプレート自動選択
    Private Sub DGV21_SelectionChanged(sender As DataGridView, e As EventArgs) Handles DGV21.SelectionChanged
        If DGV21.SelectedCells(0).ColumnIndex <> 1 Then
            Exit Sub
        End If

        If TabControl8.SelectedTab.Text <> "データ生成" Then
            TabControl8.SelectedTab = TabControl8.TabPages("TabPage27")
        End If

        If TabControl8.SelectedTab.Text = "データ生成" And CheckBox38.Checked Then
            Dim selStr As String = DGV21.Item(0, sender.SelectedCells(0).RowIndex).Value
            For i As Integer = 0 To ComboBox14.Items.Count - 1
                If InStr(ComboBox14.Items(i), selStr) > 0 Then
                    ComboBox14.SelectedIndex = i
                    Exit For
                End If
            Next

            If ComboBox14.SelectedIndex >= 0 Then
                Button31.PerformClick()
                Application.DoEvents()
                Threading.Thread.Sleep(1)
                Button55.PerformClick()
            End If
        End If
    End Sub

    '読込
    Dim dataTeigiDB As String() = Nothing
    Dim dataHenkanDB As String() = Nothing
    Dim tempAmazonHeader As New ArrayList
    Dim tempAmazonDB As New ArrayList
    Private Sub Button31_Click(sender As Object, e As EventArgs) Handles Button31.Click
        If ComboBox14.SelectedIndex >= 0 Then
            'データ一時保存
            If CheckBox22.Checked Then
                If DGV16.RowCount > 0 Then
                    tempAmazonHeader = TM_HEADER_1ROW_GET(DGV15, 2)
                    tempAmazonDB.Clear()
                    For r As Integer = 0 To DGV16.RowCount - 1
                        If DGV16.Rows(r).IsNewRow Then
                            Continue For
                        End If

                        Dim line As String = ""
                        For c As Integer = 0 To DGV16.ColumnCount - 1
                            If c = 0 Then
                                line = DGV16.Item(c, r).Value
                            Else
                                line &= "|=|" & DGV16.Item(c, r).Value
                            End If
                        Next
                        tempAmazonDB.Add(line)
                    Next
                End If
            End If

            'テンプレートの読込（テンプレート複数ファイル名ある際の処理未完成）
            Dim tempName As String() = Split(ComboBox14.SelectedItem, ",")
            Dim tName As String() = Split(tempName(1), "|")
            Dim data As DataTable = TM_DB_CONNECT_SELECT(CnAccdb, "T_設定_Amazon", "[名称]='" & tName(0) & "'")

            CheckBox25.Checked = False
            DGV15.Rows.Clear()
            DGV15.Columns.Clear()
            DGV16.Rows.Clear()
            DGV16.Columns.Clear()
            DGV19.Rows.Clear()
            DGV19.Columns.Clear()

            Dim strArray As String() = Nothing
            strArray = Split(data.Rows(0)("設定1"), "|")
            For i As Integer = 0 To strArray.Length - 1
                DGV15.Columns.Add(i, i)
                DGV15.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                DGV16.Columns.Add(i, i)
                DGV16.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                DGV19.Columns.Add(i, i)
                DGV19.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
            Next
            DGV15.Rows.Add(strArray)

            strArray = Split(data.Rows(0)("設定2"), "|")
            DGV15.Rows.Add(strArray)
            ListBox6.Items.AddRange(strArray)

            strArray = Split(data.Rows(0)("設定3"), "|")
            DGV15.Rows.Add(strArray)

            Dim dH15 As ArrayList = TM_HEADER_1ROW_GET(DGV15, 2)

            '----------------------------
            '推奨値の読み込み
            Dim str As String = data.Rows(0)("設定4")
            suiArray = Split(data.Rows(0)("設定4"), vbLf)
            str = Replace(str, "[c39]", "'")
            str = Replace(str, "[c34]", """")
            Dim setteiArray As String() = Split(str, vbLf)
            For c As Integer = 0 To dH15.Count - 1
                Dim suiName As String = Regex.Replace(dH15(c), "[0-9]{1,3}$", "")
                For i As Integer = 0 To setteiArray.Length - 1
                    If setteiArray(i) = "" Then
                        Continue For
                    End If
                    'Amazonの設定がおかしいので修正
                    Select Case suiName
                        Case "brand_name"
                            Continue For
                    End Select

                    Dim settei As String() = Split(setteiArray(i), "|")
                    If Regex.IsMatch(settei(1), "[0-9]{1,3}$") Then 'エクセルの設定でlifestyle1、lifestyle2となっている時、検索を戻す
                        suiName = dH15(c)
                    End If
                    If Regex.IsMatch(suiName, "^" & settei(1)) Then
                        If settei.Length >= 2 Then
                            Dim itemArray As New ArrayList
                            For k As Integer = 2 To settei.Length - 1
                                itemArray.Add(settei(k))
                            Next
                            DGV_COLUMN_COMBOBOX(DGV16, c, itemArray)
                        End If
                        Exit For
                    End If
                Next
            Next

            'データ定義の読み込み
            Dim hissuDB As New ArrayList
            Dim formatDB As New ArrayList
            If data.Rows(0)("説明").ToString <> "" Then
                dataTeigiDB = Split(data.Rows(0)("説明"), "|=|")     '表示用
                Dim hArray As String() = Split(data.Rows(0)("必須"), "[\n]")
                Dim fArray As String() = Split(data.Rows(0)("フォーマット"), "[\n]")
                For i As Integer = 0 To dataTeigiDB.Length - 1
                    If dataTeigiDB(i) <> "" Then
                        Dim dd As String = Split(dataTeigiDB(i), "|-|")(1)
                        hissuDB.Add(dd & "," & hArray(i))
                        formatDB.Add(dd & "," & fArray(i))
                    End If
                Next
            End If

            'Amazon入力データ変換表の読込
            Dim aPath As String = Path.GetDirectoryName(Form1.appPath) & "\config\Amazon入力データ変換表.txt"
            dataHenkanDB = File.ReadAllLines(aPath, Encoding.GetEncoding("SHIFT-JIS"))

            '必須の読み込み
            If data.Rows(0)("必須").ToString <> "" Then
                For i As Integer = 0 To hissuDB.Count - 1
                    Dim hh As String() = Split(hissuDB(i), ",")
                    If hh.Length > 1 Then
                        If dH15.Contains(hh(0)) Then
                            If hh(1) = "必須" Then
                                DGV15.Item(dH15.IndexOf(hh(0)), 2).Style.BackColor = Color.Yellow
                            ElseIf hh(1) = "推奨" Then
                                DGV15.Item(dH15.IndexOf(hh(0)), 2).Style.BackColor = Color.LightPink
                            End If
                        End If
                    End If
                Next
            End If

            'サンプルの読み込み
            If data.Rows(0)("フォーマット").ToString <> "" Then
                DGV19.Rows.Add()
                For i As Integer = 0 To formatDB.Count - 1
                    Dim ff As String() = Split(formatDB(i), ",")
                    If ff.Length > 1 Then
                        If dH15.Contains(ff(0)) Then
                            DGV19.Item(dH15.IndexOf(ff(0)), 0).Value = ff(1)
                        End If
                    End If
                Next
            End If
            '----------------------------

            ListBox9.Items.Clear()

            If CheckBox22.Checked Then
                TEMPLATE_CHANGE()
            End If
        End If
    End Sub

    '列に推奨値を入力する
    Private Sub DGV_COLUMN_COMBOBOX(dgv As DataGridView, cols As Integer, strArray As ArrayList)
        Dim column As New DataGridViewComboBoxColumn With {
                .DisplayStyleForCurrentCellOnly = True,
                .DataPropertyName = dgv.Columns(cols).DataPropertyName,
                .HeaderText = cols
            }
        For i As Integer = 0 To strArray.Count - 1
            column.Items.Add(strArray(i))
        Next
        dgv.Columns.Insert(cols, column)
        dgv.Columns.RemoveAt(cols + 1)
    End Sub

    'feed_product_typeにより推奨値が変わるのに対応
    Dim suiArray As String() = Nothing
    Private Sub DataGridView16_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DGV16.CellValueChanged
        If DGV15.RowCount > 0 And DGV16.RowCount > 1 Then
            Dim dH15 As ArrayList = TM_HEADER_1ROW_GET(DGV15, 2)
            If dH15(e.ColumnIndex) = "feed_product_type" Then
                Dim fpt = DGV16.Item(dH15.IndexOf("feed_product_type"), e.RowIndex).Value
                For i As Integer = 0 To suiArray.Length - 1
                    Dim sArray As String() = Split(suiArray(i), "|")
                    If InStr(sArray(0), "[ " & fpt & " ]") > 0 Or InStr(sArray(0), "[" & fpt & "]") > 0 Then
                        If dH15.Contains(sArray(1)) Then
                            Dim cell As DataGridViewComboBoxCell = DGV16.Item(dH15.IndexOf(sArray(1)), e.RowIndex)
                            For k As Integer = cell.Items.Count - 1 To 0 Step -1
                                cell.Items.RemoveAt(k)
                            Next
                            For k As Integer = 2 To sArray.Length - 1
                                cell.Items.Add(sArray(k))
                            Next
                        End If
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub TEMPLATE_CHANGE()
        'テーブル名
        Dim dH15 As ArrayList = TM_HEADER_1ROW_GET(DGV15, 2)

        '展開
        Dim dArray As String() = Split(tempAmazonDB(0), ",")
        For i As Integer = 0 To tempAmazonDB.Count - 1
            DGV16.Rows.Add()
            Dim rowNum As Integer = DGV16.RowCount - 2
            Dim line As String() = Split(tempAmazonDB(i), "|=|")
            For k As Integer = 0 To line.Length - 1
                If dH15.Contains(tempAmazonHeader(k)) Then
                    DGV16.Item(dH15.IndexOf(tempAmazonHeader(k)), rowNum).Value = line(k)
                End If
            Next
        Next
    End Sub

    'スクロール追随
    Private Sub DataGridView16_Scroll(sender As Object, e As ScrollEventArgs) Handles DGV16.Scroll
        Try
            DGV15.HorizontalScrollingOffset = DGV16.HorizontalScrollingOffset
            DGV19.HorizontalScrollingOffset = DGV16.HorizontalScrollingOffset
        Catch ex As Exception

        End Try
    End Sub

    Private Sub DataGridView19_Scroll(sender As Object, e As ScrollEventArgs) Handles DGV19.Scroll
        Try
            DGV15.HorizontalScrollingOffset = DGV19.HorizontalScrollingOffset
            DGV16.HorizontalScrollingOffset = DGV19.HorizontalScrollingOffset
        Catch ex As Exception

        End Try
    End Sub

    'カラム横幅同期
    Private Sub DataGridView15_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles DGV15.ColumnWidthChanged
        If DGV16.ColumnCount > e.Column.Index Then
            DGV16.Columns(e.Column.Index).Width = e.Column.Width
        End If
        If DGV19.ColumnCount > e.Column.Index Then
            DGV19.Columns(e.Column.Index).Width = e.Column.Width
        End If
    End Sub

    Private Sub DataGridView16_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles DGV16.ColumnWidthChanged
        If DGV15.ColumnCount > e.Column.Index Then
            DGV15.Columns(e.Column.Index).Width = e.Column.Width
        End If
        If DGV19.ColumnCount > e.Column.Index Then
            DGV19.Columns(e.Column.Index).Width = e.Column.Width
        End If
    End Sub

    '下削除
    Private Sub Button45_Click(sender As Object, e As EventArgs) Handles Button45.Click
        DGV16.Rows.Clear()
    End Sub

    'データ入力のデータで生成
    Private Sub Button55_Click(sender As Object, e As EventArgs) Handles Button55.Click
        If DGV20.RowCount = 0 Then
            Exit Sub
        End If
        Dim dh20 As ArrayList = TM_HEADER_GET(DGV20)
        If Not dh20.Contains("分析") Then
            MsgBox("先に分析を行なってください。")
            Exit Sub
        End If

        Dim fieldArray As New ArrayList
        Dim motoArray As New ArrayList
        Dim aPath As String = Path.GetDirectoryName(Form1.appPath) & "\config\Amazon入力データ変換表.txt"
        Dim cArray As String() = File.ReadAllLines(aPath, Encoding.GetEncoding("SHIFT-JIS"))
        For i As Integer = 0 To cArray.Length - 1
            If dh20.Contains(Split(cArray(i), ",")(1)) Or InStr(Split(cArray(i), ",")(1), "[固定]") > 0 Or InStr(Split(cArray(i), ",")(1), "[キーワード]") > 0 Then
                fieldArray.Add(Split(cArray(i), ",")(0))
                motoArray.Add(Split(cArray(i), ",")(1))
            End If
        Next

        Dim dH16 As ArrayList = TM_HEADER_1ROW_GET(DGV15, 2)
        Dim category As String = Split(ComboBox14.SelectedItem, ",")(0)

        For i As Integer = 0 To fieldArray.Count - 1
            If fieldArray(i) <> "" Or fieldArray(i) <> "//" Then
                'プログラム行にデータが無ければ処理しない、//はコメントアウト
            ElseIf Regex.IsMatch(fieldArray(i), "^プログラム") Then  '（amazon未完成）

            Else

            End If

            Dim start As Integer = -1
            For r2 As Integer = 0 To DGV20.RowCount - 1
                If DGV20.Rows(r2).IsNewRow Then
                    Continue For
                ElseIf DGV20.Item(dh20.IndexOf("分析"), r2).Value <> category Then
                    Continue For
                End If

                start += 1
                If i = 0 Then
                    DGV16.Rows.Add()
                End If
                If dH16.Contains(fieldArray(i)) Or Regex.IsMatch(fieldArray(i), "generic_keywords") Then
                    If Regex.IsMatch(motoArray(i), "^\[固定\]") Then
                        Dim str As String = Replace(motoArray(i), "[固定]", "")
                        DGV16.Item(dH16.IndexOf(fieldArray(i)), start).Value = str
                    ElseIf Regex.IsMatch(motoArray(i), "^\[キーワード\]") Then
                        Dim header As String = Replace(motoArray(i), "[キーワード]", "")
                        Dim mStr As String = ""
                        If dh20.Contains(header) Then
                            If InStr(motoArray(i), "+") > 0 Then
                                mStr &= DGV20.Item(dh20.IndexOf(header), r2).Value
                            Else
                                If CStr(DGV20.Item(dh20.IndexOf(header), r2).Value) <> "" Then
                                    mStr = DGV20.Item(dh20.IndexOf(header), r2).Value
                                End If
                            End If
                        End If
                        mStr = Replace(mStr, "　", " ")
                        mStr = Replace(mStr, "　", " ")
                        If dH16.Contains("検索キーワード1") Then
                            DGV16.Item(dH16.IndexOf(header), start).Value = mStr
                        Else
                            DGV16.Item(dH16.IndexOf("generic_keywords"), start).Value &= mStr
                        End If
                    Else
                        Dim mA As String() = Regex.Split(motoArray(i), "\||\+")
                        Dim mStr As String = ""
                        For k As Integer = 0 To mA.Length - 1
                            If dh20.Contains(mA(k)) Then
                                If InStr(motoArray(i), "+") > 0 Then
                                    mStr &= DGV20.Item(dh20.IndexOf(mA(k)), r2).Value
                                Else
                                    'MsgBox(mA(k) & "/" & dgv2.Item(dh2.IndexOf(mA(k)), r2).Value)
                                    If CStr(DGV20.Item(dh20.IndexOf(mA(k)), r2).Value) <> "" Then
                                        mStr = DGV20.Item(dh20.IndexOf(mA(k)), r2).Value
                                        Exit For
                                    End If
                                End If
                            End If
                        Next
                        DGV16.Item(dH16.IndexOf(fieldArray(i)), start).Value = mStr
                    End If
                End If
            Next

        Next

        BrandInsert()
        CheckDGV16("all")

        'DataErrorを出すために、スクロールして確認する
        Label38.Show()
        Application.DoEvents()
        DGV16.ResumeLayout()
        RichTextBox2.Visible = False
        ListBox9.Visible = False
        For r As Integer = 0 To DGV16.RowCount - 1
            For c As Integer = 0 To DGV16.ColumnCount - 1
                DGV16.CurrentCell = DGV16(c, r)
            Next
        Next
        DGV16.CurrentCell = DGV16(0, 0)
        DGV16.SuspendLayout()
        RichTextBox2.Visible = True
        ListBox9.Visible = True
        Label38.Hide()
        Application.DoEvents()

        'リストの並べ替え
        Dim listArray As New ArrayList
        For Each lA As String In ListBox9.Items
            listArray.Add(lA)
        Next
        listArray.Sort()
        ListBox9.Items.Clear()
        For Each lA As String In listArray
            ListBox9.Items.Add(lA)
        Next
    End Sub

    '文字数のチェック
    Dim SJIS As Encoding = Encoding.GetEncoding("Shift_JIS")
    Private Sub CheckDGV16(mode As String, Optional c As Integer = 0, Optional r As Integer = 0)
        Dim dH15 As ArrayList = TM_HEADER_1ROW_GET(DGV15, 2)
        If mode = "all" Then
            For r2 As Integer = 0 To DGV16.RowCount - 1
                If DGV16.Rows(r2).IsNewRow Then
                    Continue For
                End If
                For c2 As Integer = 0 To DGV16.ColumnCount - 1
                    CheckDGV16_2(dH15, c2, r2)
                Next
                CheckDGV16_3(dH15, r2)
            Next
        Else
            CheckDGV16_2(dH15, c, r)
        End If
    End Sub

    Private Sub CheckDGV16_2(dH15 As ArrayList, c As Integer, r As Integer)
        Dim dvc As DataGridViewCell = DGV16.Item(c, r)
        If dvc.Value = "" Then
            Exit Sub
        End If

        Dim headerName As String = dH15(c)
        For i As Integer = 0 To dataHenkanDB.Length - 1
            Dim setumeiArray As String() = Split(dataHenkanDB(i), ",")
            Dim dt As String = Regex.Replace(setumeiArray(0), "[0-9]{1,3}$", "")
            If Regex.IsMatch(headerName, "^" & dt & "[0-9]{0,3}$") Then
                Dim bytC As Integer = SJIS.GetByteCount(dvc.Value)
                Dim charC As Integer = dvc.Value.ToString.Length
                If setumeiArray(3) = "" Then
                    '規制無し
                ElseIf InStr(setumeiArray(3), "C") > 0 Then
                    If charC > Replace(setumeiArray(3), "C", "") Then
                        dvc.Style.BackColor = Color.Orange
                    End If
                ElseIf setumeiArray(3) >= 40 Then   '40バイト以上はバイト数＋文字数（Amazonの仕様）
                    If (bytC + charC) > setumeiArray(3) Then
                        dvc.Style.BackColor = Color.Orange
                    Else
                        If dvc.Style.BackColor <> Color.Empty Then
                            dvc.Style.BackColor = Color.Empty
                        End If
                    End If
                Else
                    If bytC > setumeiArray(3) Then
                        dvc.Style.BackColor = Color.Orange
                    Else
                        If dvc.Style.BackColor <> Color.Empty Then
                            dvc.Style.BackColor = Color.Empty
                        End If
                    End If
                End If

                Exit For
            End If
        Next
    End Sub

    Private Sub CheckDGV16_3(dH15 As ArrayList, r2 As Integer)
        Dim checkArray As String() = {"color_name|color_map", "color_map|color_name", "size_name|size_map", "size_map|size_name"}

        'nameとmapで同じのがあれば自動入力
        For i As Integer = 0 To checkArray.Length - 1
            Dim checkStr As String() = Split(checkArray(i), "|")
            If dH15.Contains(checkStr(0)) And dH15.Contains(checkStr(1)) Then
                If DGV16.Item(dH15.IndexOf(checkStr(0)), r2).Value <> "" Then
                    If DGV16.Item(dH15.IndexOf(checkStr(1)), r2).Value = "" Then
                        Try
                            DGV16.Item(dH15.IndexOf(checkStr(1)), r2).Value = DGV16.Item(dH15.IndexOf(checkStr(0)), r2).Value
                            DGV16.Item(dH15.IndexOf(checkStr(1)), r2).Style.BackColor = Color.SkyBlue
                        Catch ex As Exception
                            DGV16.Item(dH15.IndexOf(checkStr(1)), r2).Style.BackColor = Color.Orange
                            ListBox9.Items.Add(dH15.IndexOf(checkStr(1)) & ":" & r2 + 1 & "=>" & checkStr(0) & "○ " & checkStr(1) & "×")
                        End Try
                    Else
                        DGV16.Item(dH15.IndexOf(checkStr(1)), r2).Style.BackColor = Color.Empty
                    End If
                End If
            End If
        Next

        '価格のフォーマット確認
        For c As Integer = 0 To DGV16.ColumnCount - 1
            If Regex.IsMatch(dH15(c), "price", RegexOptions.IgnoreCase) Then
                Dim price As String = DGV16.Item(c, r2).Value
                If price <> "" Then
                    If Not IsNumeric(price) Then
                        DGV16.Item(c, r2).Style.BackColor = Color.Orange
                        ListBox9.Items.Add(c & ":" & r2 + 1 & "=>" & "価格が数字でない")
                    Else
                        If Regex.IsMatch(price, "\\|-|,") Then
                            price = Replace(price, "\", "")
                            price = Replace(price, "-", "")
                            DGV16.Item(c, r2).Value = Replace(price, ",", "")
                            DGV16.Item(c, r2).Style.BackColor = Color.SkyBlue
                        End If
                    End If
                End If
            End If
        Next
    End Sub

    Private Sub LinkLabel9_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel9.LinkClicked
        BrandInsert()
    End Sub

    Private Sub BrandInsert()
        If ComboBox17.SelectedItem = "" Or ComboBox17.SelectedItem = "..." Then
            Exit Sub
        End If

        Dim brand As String = ""
        Dim maker As String = ""
        Select Case ComboBox17.SelectedItem
            Case "サラダ"
                brand = "Sarada"
                maker = "Sarada"
            Case "フリット"
                brand = "Hewflit"
                maker = "Hewflit"
            Case "トココ"
                brand = "通販のトココ"
                maker = "サンパーシー"
            Case "アリス"
                brand = "雑貨の国のアリス"
                maker = "クロマチックフーガ"
        End Select

        Dim dH16 As ArrayList = TM_HEADER_1ROW_GET(DGV15, 2)
        For r As Integer = 0 To DGV16.RowCount - 1
            If Not DGV16.Rows(r).IsNewRow Then
                If dH16.Contains("item_name") Then
                    '商品の名称に入っている店名を消して新しい店名を入れる
                    Dim regStr As String = "Sarada\s{0,1}|Hewflit\s{0,1}|通販のトココ\s{0,1}|雑貨の国のアリス\s{0,1}"
                    Dim itemName As String = Regex.Replace(DGV16.Item(dH16.IndexOf("item_name"), r).Value, regStr, "")
                    DGV16.Item(dH16.IndexOf("item_name"), r).Value = brand & " " & itemName
                End If
                If dH16.Contains("brand_name") Then
                    DGV16.Item(dH16.IndexOf("brand_name"), r).Value = brand
                End If
                If dH16.Contains("manufacturer") Then
                    DGV16.Item(dH16.IndexOf("manufacturer"), r).Value = maker
                End If
            End If
        Next

        CheckDGV16("all")
    End Sub

    'csvタブのデータで生成
    Private Sub Button36_Click(sender As Object, e As EventArgs) Handles Button36.Click
        'Dim fieldArray As String() = {"item_sku", "external_product_id", "item_name", "product_description", "standard_price", "external_product_id_type"}
        'Dim motoArray As String() = {"出品者SKU", "商品ID", "商品名", "商品の説明", "価格", "[固定]EAN"}

        If DGV15.RowCount = 0 Then
            Exit Sub
        End If

        Dim fieldArray As New ArrayList
        Dim motoArray As New ArrayList
        Dim data As DataTable = TM_DB_CONNECT_SELECT(CnAccdb, "T_設定_Amazon", "[名称]='出品データ移行'")
        Dim setteiArray As String() = Split(data.Rows(0)("設定1"), "|=|")
        For i As Integer = 0 To setteiArray.Length - 1
            fieldArray.Add(Split(setteiArray(i), ",")(0))
            motoArray.Add(Split(setteiArray(i), ",")(1))
        Next

        Dim dh2 As ArrayList = TM_HEADER_GET(DGV2)
        Dim dH16 As ArrayList = TM_HEADER_1ROW_GET(DGV15, 2)

        Dim start As Integer = DGV16.RowCount - 1
        For i As Integer = 0 To fieldArray.Count - 1
            If fieldArray(i) <> "" Or fieldArray(i) <> "//" Then
                'プログラム行にデータが無ければ処理しない、//はコメントアウト
            ElseIf Regex.IsMatch(fieldArray(i), "^プログラム") Then  '（amazon未完成）

            Else
            End If

            For r2 As Integer = 0 To DGV2.RowCount - 1
                If i = 0 Then
                    DGV16.Rows.Add()
                End If
                If dH16.Contains(fieldArray(i)) Then
                    If Regex.IsMatch(motoArray(i), "^\[固定\]") Then
                        Dim str As String = Replace(motoArray(i), "[固定]", "")
                        DGV16.Item(dH16.IndexOf(fieldArray(i)), start + r2).Value = str
                    Else
                        Dim mA As String() = Regex.Split(motoArray(i), "\||\+")
                        Dim mStr As String = ""
                        For k As Integer = 0 To mA.Length - 1
                            If dh2.Contains(mA(k)) Then
                                If InStr(motoArray(i), "+") > 0 Then
                                    mStr &= DGV2.Item(dh2.IndexOf(mA(k)), r2).Value
                                Else
                                    'MsgBox(mA(k) & "/" & dgv2.Item(dh2.IndexOf(mA(k)), r2).Value)
                                    If CStr(DGV2.Item(dh2.IndexOf(mA(k)), r2).Value) <> "" Then
                                        mStr = DGV2.Item(dh2.IndexOf(mA(k)), r2).Value
                                        Exit For
                                    End If
                                End If
                            End If
                        Next
                        DGV16.Item(dH16.IndexOf(fieldArray(i)), start + r2).Value = mStr
                    End If
                End If
            Next

        Next
    End Sub

    'ノード別保存
    Private Sub Button38_Click(sender As Object, e As EventArgs) Handles Button38.Click
        If DGV16.RowCount <= 1 Then
            Exit Sub
        ElseIf ComboBox17.SelectedIndex < 0 Then
            MsgBox("店舗を選択してください")
            Exit Sub
        End If

        'テーブル名
        Dim tableName As String = "T_" & ComboBox17.Text & "データ"

        Dim dH15 As ArrayList = TM_HEADER_1ROW_GET(DGV15, 1)
        Dim dH15_2 As ArrayList = TM_HEADER_1ROW_GET(DGV15, 2)
        ToolStripProgressBar1.Maximum = DGV16.RowCount
        ToolStripProgressBar1.Value = 0

        TM_DB_CONNECT(CnAccdb_amazon)

        For r As Integer = 0 To DGV16.RowCount - 1
            If DGV16.Rows(r).IsNewRow Then
                Exit For
            End If
            If dH15_2.Contains("parent_child") Then
                If DGV16.Item(dH15_2.IndexOf("parent_child"), r).Value = "Parent" Then
                    Continue For
                End If
            End If

            Dim FieldArray As New ArrayList
            Dim SetArray As New ArrayList
            Dim sku As String = ""
            Dim data As String = ""

            FieldArray.Add("出品者SKU")
            sku = DGV16.Item(dH15.IndexOf("出品者SKU"), r).Value
            SetArray.Add(sku)
            FieldArray.Add("カテゴリ")
            SetArray.Add(Split(ComboBox14.SelectedItem, ",")(1))
            If dH15.Contains("商品タイプ") Then
                FieldArray.Add("商品タイプ")
                SetArray.Add(DGV16.Item(dH15.IndexOf("商品タイプ"), r).Value)
            End If
            If dH15.Contains("推奨ブラウズノード") Then
                FieldArray.Add("推奨ブラウズノード")
                SetArray.Add(DGV16.Item(dH15.IndexOf("推奨ブラウズノード"), r).Value)
            End If
            For c As Integer = 0 To DGV16.ColumnCount - 1
                data &= DGV15.Item(c, 2).Value & "="
                data &= DGV16.Item(c, r).Value & "|"
            Next
            data = data.TrimEnd("|")
            FieldArray.Add("データ")
            SetArray.Add(data)

            SetArray = AccessArrayConvert(SetArray)
            TM_DB_UPSERT(tableName, "[出品者SKU]='" & sku & "'", FieldArray, SetArray)
            ToolStripProgressBar1.Value += 1
        Next

        TM_DB_DISCONNECT()

        ToolStripProgressBar1.Value = 0
        MsgBox("保存しました")
    End Sub

    '読み込み（カテゴリー）
    Private Sub Button37_Click(sender As Object, e As EventArgs) Handles Button37.Click
        Dim comboSel As String = Split(ComboBox14.SelectedItem, ",")(1)
        AmazonDB_GET("[カテゴリ]='" & comboSel & "'")

        If CheckBox36.Checked Then
            DGV19.Rows.Clear()
            DGV19.Columns.Clear()
            For c As Integer = 0 To DGV16.ColumnCount - 1
                DGV19.Columns.Add(c, c)
            Next
            For r As Integer = 0 To DGV16.RowCount - 1
                DGV19.Rows.Add()
                For c As Integer = 0 To DGV16.ColumnCount - 1
                    If DGV16.Item(c, r).Value <> "" Then
                        DGV19.Item(c, r).Value = DGV16.Item(c, r).Value
                    End If
                Next
            Next
        End If
    End Sub

    Private Sub CheckBox36_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox36.CheckedChanged
        If CheckBox36.Checked Then
            DGV19.Visible = True
        Else
            DGV19.Visible = False
        End If
    End Sub

    '全て
    Private Sub Button41_Click(sender As Object, e As EventArgs) Handles Button41.Click
        AmazonDB_GET("")
    End Sub

    'コード
    Private Sub Button40_Click(sender As Object, e As EventArgs) Handles Button40.Click
        If CheckBox23.Checked = False And ComboBox14.SelectedItem Is Nothing Then
            Exit Sub
        End If

        Dim strArray As String() = Split(TextBox25.Text, vbCrLf)
        Dim where As String = ""
        For i As Integer = 0 To strArray.Length - 1
            If strArray(i) <> "" Then
                If where = "" Then
                    where &= "[出品者SKU] LIKE '%" & strArray(i) & "%'"
                    where &= " OR " & "[データ] LIKE '%" & strArray(i) & "%'"
                Else
                    where &= " OR " & "[出品者SKU] LIKE '%" & strArray(i) & "%'"
                    where &= " OR " & "[データ] LIKE '%" & strArray(i) & "%'"
                End If
            End If
        Next

        AmazonDB_GET(where, CheckBox23.Checked)
    End Sub

    Private Sub AmazonDB_GET(Optional where As String = "", Optional NodeReadFlag As Boolean = False)
        If NodeReadFlag = False And DGV15.RowCount = 0 Then
            MsgBox("読込するノードを選んでください")
            Exit Sub
        End If

        If ComboBox17.Text = "..." Then
            MsgBox("店舗が選択されていません")
            Exit Sub
        End If

        'テーブル名
        Dim tableName As String = "T_" & ComboBox17.Text & "データ"
        Dim table As New DataTable
        If where <> "" Then
            table = TM_DB_CONNECT_SELECT(CnAccdb_amazon, tableName, where)
        Else
            table = TM_DB_CONNECT_SELECT(CnAccdb_amazon, tableName)
        End If
        If table.Rows.Count = 0 Then
            MsgBox("データが見つかりません")
            Exit Sub
        End If
        Dim view As New DataView(table) With {
            .Sort = "出品者SKU ASC"
        }

        If NodeReadFlag Then
            '1行目のノードでコンボボックスを読み込む
            For n As Integer = 0 To view.Count - 1
                Dim node As String = view(n)("カテゴリ").ToString
                If node = "" Then
                    Continue For
                End If
                For i As Integer = 0 To ComboBox14.Items.Count - 1
                    If InStr(ComboBox14.Items(i), node) > 0 Then
                        ComboBox14.SelectedIndex = i
                        Exit For
                    End If
                Next
                Exit For
            Next
            Button31.PerformClick()
        End If

        '展開
        Dim dH15 As ArrayList = TM_HEADER_1ROW_GET(DGV15, 2)
        For Each row As DataRowView In view
            DGV16.Rows.Add()
            Dim rowNum As Integer = DGV16.RowCount - 2
            Dim res As String = StrAccessConvert(row("データ"))    'Access用文字を戻す
            Dim dArray As String() = Split(res, "|")
            For k As Integer = 0 To dArray.Length - 1
                Dim lists As String() = Split(dArray(k), "=")
                If dH15.Contains(lists(0)) Then
                    DGV16.Item(dH15.IndexOf(lists(0)), rowNum).Value = lists(1)
                End If
            Next
        Next
    End Sub

    '抽出
    Private Sub Button44_Click(sender As Object, e As EventArgs) Handles Button44.Click
        If ListBox6.SelectedItems.Count > 0 Then
            Dim dH15 As ArrayList = TM_HEADER_1ROW_GET(DGV15, 1)
            Dim LBArray As New ArrayList
            For i As Integer = 0 To ListBox6.SelectedItems.Count - 1
                If dH15.Contains(ListBox6.SelectedItems(i)) Then
                    LBArray.Add(dH15.IndexOf(ListBox6.SelectedItems(i)))
                End If
            Next
            For c As Integer = DGV15.ColumnCount - 1 To 0 Step -1
                If Not LBArray.Contains(c) Then
                    If CheckBox24.Checked Then
                        DGV15.Columns.RemoveAt(c)
                        DGV16.Columns.RemoveAt(c)
                        LinkLabel2.Enabled = False
                    Else
                        DGV15.Columns(c).Visible = False
                        DGV16.Columns(c).Visible = False
                        LinkLabel2.Enabled = True
                    End If
                Else
                    If CheckBox24.Checked = False Then
                        DGV15.Columns(c).Visible = True
                        DGV16.Columns(c).Visible = True
                    End If
                End If
            Next
        End If
    End Sub

    '表示戻す
    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        For c As Integer = 0 To DGV15.ColumnCount - 1
            DGV15.Columns(c).Visible = True
            DGV16.Columns(c).Visible = True
        Next
    End Sub

    '必須列のみ表示
    Private Sub LinkLabel4_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel4.LinkClicked
        For c As Integer = 0 To DGV15.ColumnCount - 1
            If DGV15.Item(c, 2).Style.BackColor = Color.Yellow Then
                DGV15.Columns(c).Visible = True
                DGV16.Columns(c).Visible = True
            Else
                DGV15.Columns(c).Visible = False
                DGV16.Columns(c).Visible = False
            End If
        Next
        LinkLabel2.Enabled = True
    End Sub

    '必須・推奨列のみ表示
    Private Sub LinkLabel5_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel5.LinkClicked
        For c As Integer = 0 To DGV15.ColumnCount - 1
            If DGV15.Item(c, 2).Style.BackColor = Color.Yellow Or DGV15.Item(c, 2).Style.BackColor = Color.LightPink Then
                DGV15.Columns(c).Visible = True
                DGV16.Columns(c).Visible = True
            Else
                DGV15.Columns(c).Visible = False
                DGV16.Columns(c).Visible = False
            End If
        Next
        LinkLabel2.Enabled = True
    End Sub

    '検索
    Private Sub Button43_Click(sender As Object, e As EventArgs) Handles Button43.Click
        If TextBox26.Text = "" Then
            Exit Sub
        End If

        Dim dgvS As DataGridView = DGV16
        Dim search As String = Regex.Replace(TextBox26.Text, "\s|　", "|")

        Dim num As Integer = dgvS.CurrentCell.RowIndex * (dgvS.ColumnCount - 1)
        If dgvS.CurrentCell Is Nothing Then
            num = 0
        End If
        num += dgvS.CurrentCell.ColumnIndex
        For r As Integer = 0 To dgvS.Rows.Count - 1
            For c As Integer = 0 To dgvS.Columns.Count - 1
                Dim nows As Integer = (r * (dgvS.ColumnCount - 1)) + c
                If num < nows Then
                    If CStr(dgvS.Item(c, r).Value) <> "" Then
                        If Regex.IsMatch(dgvS.Item(c, r).Value, search) Then
                            dgvS.CurrentCell = dgvS(c, r)
                            Exit Sub
                        End If
                    End If
                End If
            Next
        Next

        MsgBox("検索結果がありませんでした", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
    End Sub

    '置換
    Private Sub Button42_Click(sender As Object, e As EventArgs) Handles Button42.Click
        If TextBox26.Text = "" Then
            Exit Sub
        End If

        Dim dgvS As DataGridView = DGV16
        Dim search As String = ""
        If CheckBox21.Checked Then
            search = Regex.Replace(TextBox26.Text, "\s|　", "|")
        Else
            search = TextBox26.Text
        End If

        Dim num As Integer = 0
        For i As Integer = 0 To dgvS.SelectedCells.Count - 1
            If CStr(dgvS.SelectedCells(i).Value) <> "" Then
                If CheckBox21.Checked Then
                    If Regex.IsMatch(dgvS.SelectedCells(i).Value, search) Then
                        num += 1
                        dgvS.SelectedCells(i).Value = Regex.Replace(dgvS.SelectedCells(i).Value, search, TextBox27.Text)
                    End If
                Else
                    If InStr(dgvS.SelectedCells(i).Value, search) > 0 Then
                        num += 1
                        dgvS.SelectedCells(i).Value = Replace(dgvS.SelectedCells(i).Value, search, TextBox27.Text)
                    End If
                End If
            End If
        Next

        If num = 0 Then
            MsgBox("置換結果がありませんでした", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
        Else
            MsgBox(num & "件、置換しました。", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
        End If
    End Sub

    'ヘッダー検索
    Private Sub LinkLabel3_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        If TextBox28.Text = "" Then
            Exit Sub
        End If

        Dim dgvS As DataGridView = DGV15
        Dim search As String = Regex.Replace(TextBox28.Text, "\s|　", "|")

        Dim num As Integer = dgvS.CurrentCell.RowIndex * (dgvS.ColumnCount - 1)
        If dgvS.CurrentCell Is Nothing Then
            num = 0
        End If
        num += dgvS.CurrentCell.ColumnIndex
        For r As Integer = 0 To dgvS.Rows.Count - 1
            For c As Integer = 0 To dgvS.Columns.Count - 1
                Dim nows As Integer = (r * (dgvS.ColumnCount - 1)) + c
                If num < nows Then
                    If CStr(dgvS.Item(c, r).Value) <> "" Then
                        If Regex.IsMatch(dgvS.Item(c, r).Value, search) Then
                            dgvS.CurrentCell = dgvS(c, r)
                            Exit Sub
                        End If
                    End If
                End If
            Next
        Next

        MsgBox("検索結果がありませんでした", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
    End Sub

    '選択セルの文字列をランダム置換
    Private Sub Button56_Click(sender As Object, e As EventArgs) Handles Button56.Click
        For i As Integer = 0 To DGV16.SelectedCells.Count - 1
            Dim str As String = DGV16.SelectedCells(i).Value
            str = Replace(str, "　", " ")
            str = Regex.Replace(str, "\s{1,5}", " ")

            '店名
            Dim tenmei As String = ""
            If InStr(str, "Sarada") > 0 Then
                str = Regex.Replace(str, "Sarada\s{0,3}", "")
                tenmei = "Sarada "
            ElseIf InStr(str, "Hewflit") > 0 Then
                str = Regex.Replace(str, "Hewflit\s{0,3}", "")
                tenmei = "Hewflit "
            ElseIf InStr(str, "通販のトココ") > 0 Then
                str = Regex.Replace(str, "通販のトココ\s{0,3}", "")
                tenmei = "通販のトココ "
            ElseIf InStr(str, "雑貨の国のアリス") > 0 Then
                str = Regex.Replace(str, "雑貨の国のアリス\s{0,3}", "")
                tenmei = "雑貨の国のアリス "
            End If

            'ランダム並び替え
            Dim strArray As String() = Split(str, " ")
            Dim strArray2 As String() = strArray.OrderBy(Function(s) Guid.NewGuid()).ToArray()
            For k As Integer = 0 To strArray2.Length - 1
                If k = 0 Then
                    str = tenmei & strArray2(k)
                Else
                    str &= " " & strArray2(k)
                End If
            Next

            DGV16.SelectedCells(i).Value = str
        Next
    End Sub


    '選択行をDBから削除
    Private Sub Button39_Click(sender As Object, e As EventArgs) Handles Button39.Click
        Dim DR As DialogResult = MsgBox("選択行を削除してもよろしいですか？", MsgBoxStyle.OkCancel Or MsgBoxStyle.SystemModal)
        If DR = DialogResult.Cancel Then
            Exit Sub
        End If

        Dim delArray As New ArrayList
        For i As Integer = 0 To DGV16.SelectedCells.Count - 1
            If Not delArray.Contains(DGV16.SelectedCells(i).RowIndex) Then
                delArray.Add(DGV16.SelectedCells(i).RowIndex)
            End If
        Next
        delArray.Sort()

        Dim tableName As String = "T_" & ComboBox17.Text & "データ"
        Dim dH15 As ArrayList = TM_HEADER_1ROW_GET(DGV15, 1)
        Dim num As Integer = 0
        For i As Integer = delArray.Count - 1 To 0 Step -1
            Dim code As String = DGV16.Item(dH15.IndexOf("出品者SKU"), delArray(i)).Value
            If TM_DB_CONNECT_DELETE(CnAccdb_amazon, tableName, "[出品者SKU]='" & code & "'") Then
                DGV16.Rows.RemoveAt(delArray(i))
                num += 1
            End If
        Next

        MsgBox(num & "件削除しました")
    End Sub

    '親ページ生成
    Private Sub CheckBox25_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox25.CheckedChanged
        If DGV16 Is Nothing Then
            Exit Sub
        ElseIf DGV16.RowCount <= 1 Then
            Exit Sub
        End If

        Dim dH15 As ArrayList = TM_HEADER_1ROW_GET(DGV15, 2)

        If CheckBox25.Checked Then
            Dim dgvStyle As New DataGridViewCellStyle With {
                .Alignment = DataGridViewContentAlignment.MiddleLeft,
                .ForeColor = Color.Navy,
                .BackColor = Color.LightGreen
            }
            Dim pData As DataTable = TM_DB_CONNECT_SELECT(CnAccdb, "T_設定_Amazon", "[名称]='親ページ作成'")
            Dim pD As String() = Split(pData(0)("設定1"), "|=|")
            Dim parentCode As New ArrayList
            For r As Integer = 0 To (DGV16.RowCount - 1) * 2
                If DGV16.Rows(r).IsNewRow Then
                    Continue For
                End If
                Dim check1 As String = DGV16.Item(dH15.IndexOf("parent_child"), r).Value
                Dim check2 As String = DGV16.Item(dH15.IndexOf("parent_sku"), r).Value
                If check1 <> "" Then
                    If Regex.IsMatch(check1, "Child", RegexOptions.IgnoreCase) Then
                        If Not parentCode.Contains(check2) Then
                            DGV16.Rows.Insert(r)

                            For i As Integer = 0 To pD.Length - 1
                                Dim pArray As String() = Split(pD(i), ",")
                                If dH15.Contains(pArray(0)) Then
                                    DGV16.Item(dH15.IndexOf(pArray(0)), r).Value = DGV16.Item(dH15.IndexOf(pArray(1)), r + 1).Value
                                End If
                            Next
                            DGV16.Item(dH15.IndexOf("parent_child"), r).Value = "Parent"

                            For c As Integer = 0 To DGV16.ColumnCount - 1
                                DGV16.Item(c, r).Style = dgvStyle
                            Next
                            parentCode.Add(check2)
                        End If
                    End If
                Else '枝番号なし
                    For c As Integer = 0 To DGV16.ColumnCount - 1
                        DGV16.Item(c, r).Style.BackColor = Color.Aqua
                    Next
                End If
                If r = DGV16.RowCount - 2 Then
                    Exit For
                End If
            Next
        Else
            For r As Integer = DGV16.RowCount - 1 To 0 Step -1
                Dim check1 As String = DGV16.Item(dH15.IndexOf("parent_child"), r).Value
                If check1 <> "" Then
                    If Regex.IsMatch(check1, "Parent", RegexOptions.IgnoreCase) Then
                        DGV16.Rows.RemoveAt(r)
                    End If
                Else
                    For c As Integer = 0 To DGV16.ColumnCount - 1
                        DGV16.Item(c, r).Style.BackColor = Color.Empty
                    Next
                End If
            Next
        End If
    End Sub


    Private Sub LISTVIEW_UPDATE2()
        ListView2.Items.Clear()
        ListView2.Visible = False
        ProgressBar2.Visible = True

        Dim dic As DirectoryInfo = New DirectoryInfo(dbimg_yahooa)
        Dim fileInfo As FileInfo
        ProgressBar2.Maximum = dic.GetFiles.Count
        ProgressBar2.Value = 0

        For Each fileInfo In dic.GetFiles()
            Dim item() As String = {fileInfo.Name, fileInfo.LastWriteTime.ToString(), CStr(Math.Floor(fileInfo.Length / 1000)) & " KB"}
            ListView2.Items.Add(New ListViewItem(item))
            ProgressBar2.Value += 1
        Next

        ProgressBar2.Value = 0
        ProgressBar2.Visible = False
        ListView2.Visible = True
    End Sub

    Private Sub 再表示ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 再表示ToolStripMenuItem.Click
        LISTVIEW_UPDATE2()
    End Sub

    Private Sub 詳細表示ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 詳細表示ToolStripMenuItem.Click
        ListView2.View = View.Details
    End Sub

    Private Sub リストビューToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles リストビューToolStripMenuItem.Click
        ListView2.View = View.List
    End Sub

    Private Sub 小アイコンToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 小アイコンToolStripMenuItem.Click
        ListView2.View = View.SmallIcon
    End Sub

    Private Sub 大アイコンToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 大アイコンToolStripMenuItem.Click
        ListView2.View = View.LargeIcon
    End Sub

    Private Sub タイルToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles タイルToolStripMenuItem.Click
        ListView2.View = View.Tile
    End Sub

    Private Sub ListView2_ItemSelectionChanged(sender As Object, e As ListViewItemSelectionChangedEventArgs) Handles ListView2.ItemSelectionChanged
        Dim imgPath As String = TextBox32.Text & "\" & ListView2.SelectedItems(0).Text
        If File.Exists(imgPath) Then
            PictureBox1.ImageLocation = imgPath
        Else
            PictureBox1.Image = Nothing
        End If
    End Sub

    Private Sub ListView2_DragDrop(sender As Object, e As DragEventArgs) Handles ListView2.DragDrop
        Me.Activate()
        ListBox10.Items.Clear()
        For Each filename As String In e.Data.GetData(DataFormats.FileDrop)
            ListBox10.Items.Add(Path.GetFileName(filename))
        Next
        Application.DoEvents()

        For Each filename As String In e.Data.GetData(DataFormats.FileDrop)
            LIST4VIEW("ドラッグドロップ認識", "s")
            Dim fileDst As String = TextBox32.Text & "\" & Path.GetFileName(filename)
            File.Copy(filename, fileDst)

            ListBox10.Items.RemoveAt(0)
            Application.DoEvents()
        Next

        LISTVIEW_UPDATE2()
        MsgBox("追加しました")
    End Sub

    Private Sub ListView2_DragEnter(sender As Object, e As DragEventArgs) Handles ListView2.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    '追加
    Private Sub ToolStripButton10_Click(sender As Object, e As EventArgs) Handles ToolStripButton10.Click

    End Sub

    '削除
    Private Sub ToolStripButton11_Click(sender As Object, e As EventArgs) Handles ToolStripButton11.Click
        Dim DR As DialogResult = MsgBox("選択ファイルを削除してもよろしいですか？", MsgBoxStyle.OkCancel)
        If DR = DialogResult.Cancel Then
            Exit Sub
        End If

        Dim serverDir As String = Form1.サーバーToolStripMenuItem.Text & "\dbimg\"
        Dim fileDir As String = TextBox32.Text & "\"
        For i As Integer = 0 To ListView2.SelectedItems.Count - 1
            File.Delete(fileDir & ListView2.SelectedItems(i).Text)
            If File.Exists(serverDir & ListView2.SelectedItems(i).Text) Then
                File.Delete(serverDir & ListView2.SelectedItems(i).Text)
            End If
        Next

        LISTVIEW_UPDATE2()
        MsgBox("削除しました")
    End Sub

    Private Sub Button52_Click(sender As Object, e As EventArgs) Handles Button52.Click
        Process.Start(TextBox32.Text)
    End Sub

    '画像サーバー同期
    Private Sub Button46_Click(sender As Object, e As EventArgs) Handles Button46.Click
        Dim serverDir As String = Form1.サーバーToolStripMenuItem.Text & "\dbimg\"
        Dim fileDir As String = TextBox32.Text & "\"

        Dim dicServer As DirectoryInfo = New DirectoryInfo(serverDir)
        Dim dicLocal As DirectoryInfo = New DirectoryInfo(fileDir)
        Dim fileInfo As FileInfo
        For Each fileInfo In dicServer.GetFiles()
            If Not File.Exists(fileDir & fileInfo.Name) Then
                If CheckBox29.Checked Then
                    File.Delete(serverDir & fileInfo.Name)
                    LIST4VIEW(fileInfo.Name & " S DEL", "s")
                Else
                    File.Copy(serverDir & fileInfo.Name, fileDir & fileInfo.Name, True)
                    LIST4VIEW(fileInfo.Name & " S->L", "s")
                End If
            Else
                If File.GetLastWriteTime(fileDir & fileInfo.Name) < File.GetLastWriteTime(serverDir & fileInfo.Name) Then
                    File.Copy(serverDir & fileInfo.Name, fileDir & fileInfo.Name, True)
                    LIST4VIEW(fileInfo.Name & " S->L UP", "s")
                End If
            End If
        Next

        For Each fileInfo In dicLocal.GetFiles()
            If Not File.Exists(serverDir & fileInfo.Name) Then
                File.Copy(fileDir & fileInfo.Name, serverDir & fileInfo.Name)
                LIST4VIEW(fileInfo.Name & " L->S", "s")
            Else
                If File.GetLastWriteTime(fileDir & fileInfo.Name) > File.GetLastWriteTime(serverDir & fileInfo.Name) Then
                    File.Copy(fileDir & fileInfo.Name, serverDir & fileInfo.Name, True)
                    LIST4VIEW(fileInfo.Name & " L->S UP", "s")
                End If
            End If
        Next

        MsgBox("同期終了しました")
    End Sub

    '検索
    Private Sub Button47_Click(sender As Object, e As EventArgs) Handles Button47.Click
        For i As Integer = 0 To ListView2.Items.Count - 1
            If InStr(ListView2.Items(i).Text, TextBox31.Text) > 0 Then
                ListView2.Items(i).BackColor = Color.Yellow
            Else
                ListView2.Items(i).BackColor = Color.Empty
            End If
        Next
    End Sub

    Private Sub ToolStripTextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles ToolStripTextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim SR As Sgry.Azuki.SearchResult = AzukiControl1.Document.FindNext(ToolStripTextBox3.Text, AzukiControl1.Document.CaretIndex)
            If Not SR Is Nothing Then
                AzukiControl1.Document.SetSelection(SR.Begin, SR.End)
                AzukiControl1.ScrollToCaret()
            Else
                MsgBox("指定の文字は見つかりませんでした", MsgBoxStyle.SystemModal)
            End If
        End If
    End Sub

    '**************************************
    'ヤフオク在庫チェック
    '**************************************
    Private Sub ListBox7_DragDrop(ByVal sender As Object, ByVal e As Windows.Forms.DragEventArgs) Handles _
            ListBox7.DragDrop, ListBox8.DragDrop
        Me.Activate()
        ListBox7.Items.Clear()
        Dim codeArray1 As New ArrayList
        Dim codeArray2 As New ArrayList
        For Each filename As String In e.Data.GetData(DataFormats.FileDrop)
            LIST4VIEW("ドラッグドロップ認識", "s")
            Dim readLines As String() = File.ReadAllLines(filename, Encoding.GetEncoding("SHIFT-JIS"))
            Dim header As String() = Split(readLines(0), ",")
            Dim header_kanri As String = ""
            If header.Contains("管理番号") Then
                header_kanri = "管理番号"
            Else
                MsgBox("ヤフオクの出品中一覧からcsvをダウンロードしてください")
                Exit Sub
            End If

            '訳あり商品              aa000-w
            '訳ありで新宮発送        aa000-ws
            '予約商品　              aa000-y
            '井相田にある新品        aa000-h
            '上記当てはまらない特殊　aa000-x
            For i As Integer = 1 To readLines.Length - 1
                If readLines(i) <> "" Then
                    Dim code As String = Split(readLines(i), ",")(Array.IndexOf(header, header_kanri))
                    code = Replace(code, """", "")
                    If Not Regex.IsMatch(code, "\-w|\-ws|\-y|\-h|\-x") Then '訳ありを除く
                        If Not codeArray1.Contains(code) Then
                            codeArray1.Add(code)
                        End If
                    Else
                        Dim code2 As String = Regex.Replace(code, "\-w|\-ws|\-y|\-h|\-x", "")
                        If Not codeArray1.Contains(code2) Then
                            codeArray1.Add(code2)
                        End If
                        If Not codeArray2.Contains(code) Then
                            codeArray2.Add(code)
                        End If
                    End If
                End If
            Next
        Next

        codeArray1.Sort()
        For Each c As String In codeArray1
            ListBox7.Items.Add(c)
        Next
        codeArray2.Sort()
        For Each c As String In codeArray2
            ListBox8.Items.Add(c)
        Next
    End Sub

    Private Sub ListBox7_DragEnter(ByVal sender As Object, ByVal e As Windows.Forms.DragEventArgs) Handles _
            ListBox7.DragEnter, ListBox8.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub CheckBox33_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox33.CheckedChanged
        TextBox36.Visible = True
    End Sub

    'マスター呼出検索
    Private Sub Button51_Click(sender As Object, e As EventArgs) Handles Button51.Click
        If CheckBox33.Checked Then
            Dim readLines As String() = Split(TextBox36.Text, vbCrLf)
            Dim codeArray1 As New ArrayList
            Dim codeArray2 As New ArrayList
            For i As Integer = 0 To readLines.Length - 1
                If readLines(i) <> "" Then
                    readLines(i) = Replace(readLines(i), """", "")
                    If Not Regex.IsMatch(readLines(i), "\-w|\-ws|\-y|\-h|\-x") Then '訳ありを除く
                        If Not codeArray1.Contains(readLines(i)) Then
                            codeArray1.Add(readLines(i))
                        End If
                    Else
                        Dim code2 As String = Regex.Replace(readLines(i), "\-w|\-ws|\-y|\-h|\-x", "")
                        If Not codeArray1.Contains(code2) Then
                            codeArray1.Add(code2)
                        End If
                        If Not codeArray2.Contains(readLines(i)) Then
                            codeArray2.Add(readLines(i))
                        End If
                    End If
                End If
            Next
            codeArray1.Sort()
            For Each c As String In codeArray1
                ListBox7.Items.Add(c)
            Next
            codeArray2.Sort()
            For Each c As String In codeArray2
                ListBox8.Items.Add(c)
            Next
        End If

        If ListBox7.Items.Count = 0 Then
            Exit Sub
        End If

        Dim codeArray As New ArrayList
        For i As Integer = 0 To ListBox7.Items.Count - 1
            codeArray.Add(ListBox7.Items(i).ToString)
        Next

        Dim dH5 As ArrayList = TM_HEADER_GET(DGV5)
        Dim dH18 As ArrayList = TM_HEADER_GET(DGV18)
        For r1 As Integer = 0 To DGV5.RowCount - 1
            Dim sCode As String = DGV5.Item(dH5.IndexOf("商品コード"), r1).Value
            Dim dCode As String = DGV5.Item(dH5.IndexOf("代表商品コード"), r1).Value

            If codeArray.Contains(sCode) Or codeArray.Contains(dCode) Then
                DGV18.Rows.Add()
                Dim newRow As Integer = DGV18.RowCount - 1

                DGV18.Item(dH18.IndexOf("商品コード"), newRow).Value = DGV5.Item(dH5.IndexOf("商品コード"), r1).Value
                DGV18.Item(dH18.IndexOf("代表商品コード"), newRow).Value = DGV5.Item(dH5.IndexOf("代表商品コード"), r1).Value
                DGV18.Item(dH18.IndexOf("商品区分"), newRow).Value = DGV5.Item(dH5.IndexOf("商品区分"), r1).Value
                If DGV18.Item(dH18.IndexOf("商品区分"), newRow).Value = "予約" Then
                    DGV18.Item(dH18.IndexOf("商品区分"), newRow).Style.BackColor = Color.LightPink
                End If
                DGV18.Item(dH18.IndexOf("在庫数"), newRow).Value = DGV5.Item(dH5.IndexOf("在庫数"), r1).Value
                DGV18.Item(dH18.IndexOf("引当数"), newRow).Value = DGV5.Item(dH5.IndexOf("引当数"), r1).Value
                DGV18.Item(dH18.IndexOf("フリー在庫"), newRow).Value = DGV5.Item(dH5.IndexOf("フリー在庫"), r1).Value
                DGV18.Item(dH18.IndexOf("予約在庫"), newRow).Value = DGV5.Item(dH5.IndexOf("予約在庫"), r1).Value
                DGV18.Item(dH18.IndexOf("予約引当"), newRow).Value = DGV5.Item(dH5.IndexOf("予約引当"), r1).Value
                DGV18.Item(dH18.IndexOf("予約フリー"), newRow).Value = DGV5.Item(dH5.IndexOf("予約フリー"), r1).Value

                If codeArray.Contains(sCode) And Not codeArray.Contains(dCode) Then
                    DGV18.Item(dH18.IndexOf("商品コード"), newRow).Style.BackColor = Color.LightBlue
                Else
                    DGV18.Item(dH18.IndexOf("代表商品コード"), newRow).Style.BackColor = Color.LightBlue
                End If
            End If
        Next

        Dim ue_dCode As String = ""
        For r1 As Integer = 0 To DGV18.RowCount - 1
            Dim dCode As String = DGV18.Item(dH18.IndexOf("代表商品コード"), r1).Value
            Dim dfZaiko As Integer = 0
            If dCode <> "" Then
                For r2 As Integer = 0 To DGV18.RowCount - 1
                    If dCode = DGV18.Item(dH18.IndexOf("代表商品コード"), r2).Value Then
                        dfZaiko += DGV18.Item(dH18.IndexOf("フリー在庫"), r2).Value
                        dfZaiko += DGV18.Item(dH18.IndexOf("予約フリー"), r2).Value
                    End If
                Next
            Else
                dfZaiko += DGV18.Item(dH18.IndexOf("フリー在庫"), r1).Value
                dfZaiko += DGV18.Item(dH18.IndexOf("予約フリー"), r1).Value
            End If
            DGV18.Item(dH18.IndexOf("DF在庫"), r1).Value = dfZaiko

            If DGV18.Item(dH18.IndexOf("商品コード"), r1).Style.BackColor = Color.LightBlue Then
                If DGV18.Item(dH18.IndexOf("商品区分"), r1).Value = "予約" Then
                    If dfZaiko < 5 Then
                        DGV18.Item(dH18.IndexOf("予約フリー"), r1).Style.BackColor = Color.Yellow
                    ElseIf dfZaiko < 10 Then
                        DGV18.Item(dH18.IndexOf("予約フリー"), r1).Style.BackColor = Color.LightPink
                    Else
                        DGV18.Item(dH18.IndexOf("予約フリー"), r1).Style.BackColor = Color.WhiteSmoke
                    End If
                Else
                    If dfZaiko < 5 Then
                        DGV18.Item(dH18.IndexOf("フリー在庫"), r1).Style.BackColor = Color.Yellow
                    ElseIf dfZaiko < 10 Then
                        DGV18.Item(dH18.IndexOf("フリー在庫"), r1).Style.BackColor = Color.LightPink
                    Else
                        DGV18.Item(dH18.IndexOf("フリー在庫"), r1).Style.BackColor = Color.WhiteSmoke
                    End If
                End If
            Else
                If dfZaiko < 5 Then
                    DGV18.Item(dH18.IndexOf("DF在庫"), r1).Style.BackColor = Color.Yellow
                ElseIf dfZaiko < 10 Then
                    DGV18.Item(dH18.IndexOf("DF在庫"), r1).Style.BackColor = Color.LightPink
                Else
                    DGV18.Item(dH18.IndexOf("DF在庫"), r1).Style.BackColor = Color.WhiteSmoke
                End If
            End If

            ue_dCode = dCode
        Next

        For i As Integer = 0 To ListBox8.Items.Count - 1
            Dim code As String = Regex.Replace(ListBox8.Items(i), "\-w|\-ws|\-y|\-h|\-x", "")
            For r2 As Integer = 0 To DGV18.RowCount - 1
                If code = DGV18.Item(dH18.IndexOf("商品コード"), r2).Value Or code = DGV18.Item(dH18.IndexOf("代表商品コード"), r2).Value Then
                    Select Case True
                        Case Regex.IsMatch(ListBox8.Items(i), "\-w")
                            DGV18.Item(dH18.IndexOf("メモ"), r2).Value = "訳"
                        Case Regex.IsMatch(ListBox8.Items(i), "\-ws")
                            DGV18.Item(dH18.IndexOf("メモ"), r2).Value = "訳新宮"
                        Case Regex.IsMatch(ListBox8.Items(i), "\-y")
                            DGV18.Item(dH18.IndexOf("メモ"), r2).Value = "予約"
                            If DGV18.Item(dH18.IndexOf("商品区分"), r2).Value <> "予約" Then
                                DGV18.Item(dH18.IndexOf("商品区分"), r2).Style.BackColor = Color.Yellow
                            End If
                        Case Regex.IsMatch(ListBox8.Items(i), "\-h")
                            DGV18.Item(dH18.IndexOf("メモ"), r2).Value = "井相田"
                        Case Regex.IsMatch(ListBox8.Items(i), "\-x")
                            DGV18.Item(dH18.IndexOf("メモ"), r2).Value = "特殊"
                    End Select
                End If
            Next
        Next
    End Sub

    Private Sub ListBox7_DoubleClick(sender As Object, e As EventArgs) Handles ListBox7.DoubleClick
        Dim code As String = Split(ListBox7.SelectedItem, "-")(0)
        Dim dH18 As ArrayList = TM_HEADER_GET(DGV18)
        For r As Integer = 0 To DGV18.RowCount - 1
            If Regex.IsMatch(DGV18.Item(dH18.IndexOf("商品コード"), r).Value, "^" & code) Then
                DGV18.ClearSelection()
                DGV18.Item(dH18.IndexOf("商品コード"), r).Selected = True
                DGV18.CurrentCell = DGV18(dH18.IndexOf("商品コード"), r)
                Exit Sub
            End If
        Next
    End Sub

    Private Sub ListBox8_DoubleClick(sender As Object, e As EventArgs) Handles ListBox8.DoubleClick
        Dim code As String = Split(ListBox8.SelectedItem, "-")(0)
        Dim dH18 As ArrayList = TM_HEADER_GET(DGV18)
        For r As Integer = 0 To DGV18.RowCount - 1
            If Regex.IsMatch(DGV18.Item(dH18.IndexOf("商品コード"), r).Value, "^" & code) Then
                DGV18.ClearSelection()
                DGV18.Item(dH18.IndexOf("商品コード"), r).Selected = True
                DGV18.CurrentCell = DGV18(dH18.IndexOf("商品コード"), r)
                Exit Sub
            End If
        Next
    End Sub


    '**************************************
    'マスターの処理
    '**************************************
    'マスターpageデータ整合処理
    Private Sub Button49_Click(sender As Object, e As EventArgs) Handles Button49.Click
        MsgBox("マスターpageデータ整合処理を開始します")

        TM_DB_CONNECT(CnAccdb)
        Dim tableMaster As DataTable = TM_DB_GET_STR("T_MASTER_master", "商品コード", "")
        Dim masterArray As New ArrayList
        For i As Integer = 0 To tableMaster.Rows.Count - 1
            masterArray.Add(tableMaster.Rows(i)("商品コード"))
        Next
        Dim tablePage As DataTable = TM_DB_GET_STR("T_Master_page", "商品コード", "")

        Dim delStr As String = ""
        For i As Integer = 0 To tablePage.Rows.Count - 1
            Dim code As String = tablePage.Rows(i)("商品コード")
            If Not masterArray.Contains(code) Then
                TM_DB_DELETE("T_Master_page", "[商品コード]='" & code & "'")
                delStr &= code & vbCrLf
            End If
        Next

        If delStr <> "" Then
            delStr &= vbCrLf & "上記を削除しました" & vbCrLf
        End If
        MsgBox(delStr & "マスターpageデータ整合処理を完了しました")

        '引き続き保存処理
        Dim sfd As New SaveFileDialog With {
            .Filter = "CSVファイル(*.csv)|*.csv|すべてのファイル(*.*)|*.*",
                    .AutoUpgradeEnabled = False,
            .FilterIndex = 0,
            .Title = "保存先のファイルを選択してください",
            .RestoreDirectory = True,
            .OverwritePrompt = True,
            .CheckPathExists = True,
            .FileName = "page.csv"
        }
        '.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),

        TM_DB_DISCONNECT()

        ToolStripComboBox2.SelectedItem = "マスタ"
        ToolStripComboBox1.SelectedItem = "page"
        ToolStripComboBox3.SelectedItem = "無し"
        ToolStripButton1.PerformClick()
        Application.DoEvents()
    End Sub

    Private Sub Button50_Click(sender As Object, e As EventArgs) Handles Button50.Click
        If MsgBox("データを登録してよろしいですか？", MsgBoxStyle.OkCancel Or MsgBoxStyle.SystemModal) = MsgBoxResult.Cancel Then
            Exit Sub
        End If

        MALLDB_CONNECT(CnAccdb)
        Dim dH2 As ArrayList = TM_HEADER_GET(DGV2)

        'テーブルを削除する
        TableName = "T_Master_page"
        Dim SQL1 As String = "DROP TABLE [" & TableName & "]"
        SQLCm0.CommandText = SQL1
        SQLCm0.ExecuteNonQuery()

        'テーブルを作成する
        Dim headerStr As String = ""
        For c As Integer = 0 To dH2.Count - 1
            headerStr &= "[" & dH2(c) & "] MEMO,"
        Next
        headerStr = headerStr.Trim(",")
        Dim SQL As String = "CREATE TABLE [" & TableName & "] "
        SQL &= "(" & headerStr & ")"
        SQLCm0.CommandText = SQL
        SQLCm0.ExecuteNonQuery()

        Dim setField As String = ""
        For c As Integer = 0 To dH2.Count - 1
            setField &= "[" & dH2(c) & "],"
        Next
        setField = setField.Trim(",")

        ToolStripProgressBar1.Maximum = DGV2.RowCount
        ToolStripProgressBar1.Value = 0
        For r As Integer = 0 To DGV2.RowCount - 1
            Dim setStr As String = ""
            For c As Integer = 0 To DGV2.ColumnCount - 1
                setStr &= "'" & DGV2.Item(c, r).Value & "',"
            Next
            setStr = setStr.TrimEnd(",")
            MALLDB_SET_STR(TableName, setField, setStr)
            ToolStripProgressBar1.Value += 1
        Next

        'フォルダにファイルを保存
        Dim table0 As DataTable = MALLDB_GET_SQL(TableName, "")
        Dim pathEnd As String = Path.GetDirectoryName(Form1.appPath) & "\tenpo\" & ToolStripComboBox2.SelectedItem & "_" & ToolStripComboBox1.SelectedItem & ".csv"
        ConvertDataTableToCsv(table0, pathEnd, True)

        MALLDB_DISCONNECT()

        'ファイル更新時間をデータベースに保存
        FileGetLastTime(pathEnd)

        MsgBox("データを登録しました", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)

        TableGetLastTime(0)
    End Sub

    Private Sub AzukiControl_SizeChanged(sender As Object, e As EventArgs) Handles AzukiControl4.SizeChanged, AzukiControl2.SizeChanged, AzukiControl1.SizeChanged

    End Sub

    Private Sub ToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem4.Click
        Shiji.Show()
    End Sub





#End Region

End Class