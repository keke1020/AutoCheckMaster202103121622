Imports System.Data.OleDb
Imports System.IO
Imports System.Text
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel


Public Class SagawaDenpyo
    Dim cnbackupdb_path As String = "\\SERVER2-PC\backup\BZ伝票\BZ.accdb"
    Private Cn0 As New OleDbConnection
    Private SQLCm0 As OleDbCommand
    Private Adapter0 As New OleDbDataAdapter

    Private setField As String = ""
    Private setStr As String = ""
    Private where As String = ""

    Dim checkbz_header1 As String = "お届け先コード,お届け先郵便番号,お届け先住所１,お届け先住所２,お届け先住所３,お届け先名１,お届け先名２,お届け先電話,お届け先メールアドレス,代行ご依頼主郵便番号,代行ご依頼主住所１,代行ご依頼主住所２,代行ご依頼主名１,代行ご依頼主名２,代行ご依頼主電話,代行ご依頼主メールアドレス,佐川急便顧客コード,問い合せ№,発送日,個数,顧客管理番号,便種コード,配達指定日,時間帯コード,代引金額,代引消費税,保険金額,元着区分（未使用）,営止め区分,営止精算店コード,営止精算店ローカルコード,記事欄01,記事欄02,記事欄03,記事欄04,記事欄05,記事欄06,記事欄07,記事欄08,記事欄09,記事欄10,記事欄11,記事欄12,マスタコード,マスタ配送,マスタ個口,マスタ備考,マーク,処理用,処理用2"

    Private Sub SagawaDenpyo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
        TextBox2.Text = ""
        datagridview1.Rows.Clear()
        DataGridView3.Rows.Clear()
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        ComboBox3.SelectedIndex = -1
        loadDgv2()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        NameChangeSave(1)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        NameChangeSave(2)
    End Sub

    Private Sub NameChangeSave(ByVal name As Integer)
        Dim ofd As New OpenFileDialog With {
            .InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
            .Filter = "CSVファイル(*.csv)|*.csv",
                    .AutoUpgradeEnabled = False,
            .FilterIndex = 0,
            .Title = "開くファイルを選択してください",
            .RestoreDirectory = True,
            .CheckFileExists = True,
            .CheckPathExists = True
        }

        'ダイアログを表示する
        If ofd.ShowDialog(Me) = DialogResult.OK Then
            If name = "1" Then
                TextBox1.Text = ofd.FileName
            ElseIf name = "2" Then
                TextBox2.Text = ofd.FileName
            End If
        End If
    End Sub

    Private Sub KryptonButton3_Click(sender As Object, e As EventArgs) Handles KryptonButton3.Click
        upload(1)
        loadDgv2()
    End Sub

    Private Sub KryptonButton1_Click(sender As Object, e As EventArgs) Handles KryptonButton1.Click
        upload(2)
        loadDgv2()
    End Sub

    Private Sub upload(name As Integer)
        Dim uploadFile As String = ""
        Dim backupFolder As String = ""
        Dim BZ As String = ""
        Dim BZFolder As String = ""

        If name = 1 Then
            uploadFile = TextBox1.Text
            BZFolder = ComboBox1.SelectedItem
            backupFolder = TextBox5.Text
        ElseIf name = 2 Then
            uploadFile = TextBox2.Text
            BZFolder = ComboBox2.SelectedItem
            backupFolder = TextBox6.Text
        End If

        If uploadFile = "" Then
            MsgBox("開くファイルを選択してください。")
            Exit Sub
        End If

        'check フォルダー
        If System.IO.Directory.Exists(BZFolder) = False Then
            MessageBox.Show("BIZSERVER-PCのフォルダにアクセスできない、ネットワークの問題かもしれない、管理員に連絡してください。")
            Exit Sub
        End If

        If System.IO.Directory.Exists(backupFolder) = False Then
            MessageBox.Show("バックアップフォルダにアクセスできない、ネットワークの問題かもしれない、管理員に連絡してください。")
            Exit Sub
        End If

        If System.IO.File.Exists(uploadFile) = False Then
            MessageBox.Show("アップロードファイルが存在しない。")
            Exit Sub
        End If

        If checkFileHeader(uploadFile) Then
            Dim time_now As String = Format(Now, "yyyy/MM/dd HH:mm:ss")
            Dim time_csv As String = Format(Now, "yyyyMMddHHmmss") & ".csv"

            BZFolder &= "\" & time_csv
            backupFolder &= "\" & time_csv

            '成功的话倒数3秒 防止重发
            If name = 1 Then
                KryptonButton3.Enabled = False
                File.Copy(uploadFile, BZFolder, True)
                File.Copy(uploadFile, backupFolder, True)

                BZ = "1号"
                saveDB(BZ, uploadFile, backupFolder, time_now)

                System.Threading.Thread.Sleep(3000 * 1)
                KryptonButton3.Enabled = True
            ElseIf name = 2 Then
                KryptonButton1.Enabled = False
                File.Copy(uploadFile, BZFolder, True)
                File.Copy(uploadFile, backupFolder, True)

                BZ = "2号"
                saveDB(BZ, uploadFile, backupFolder, time_now)


                System.Threading.Thread.Sleep(3000 * 1)
                KryptonButton1.Enabled = True
            End If
        End If

        TextBox1.Text = ""
        TextBox2.Text = ""
        KryptonButton3.Enabled = True
        KryptonButton1.Enabled = True
    End Sub

    Private Function checkFileHeader(uploadFile As String)
        Dim csvRecords As New ArrayList()
        Dim flag As Boolean = True
        csvRecords = TM_CSV_READ(uploadFile)(0)
        Dim cHeader As String = Replace(csvRecords(0), "|=|", ",")

        If checkbz_header1 = cHeader Then
        Else
            MsgBox(uploadFile & vbCrLf & "ファイルの認識ができません。")
            flag = False
        End If
        Return flag
    End Function




    Private Sub saveDB(BZ As String, uploadFile As String, backupFolder As String, time As String)
        Dim TableName As String = "リスト"
        Dim Table As DataTable = MALLDB_CONNECT_SELECT(cnbackupdb_path, TableName)

        setField = ""
        setStr = ""

        setField = "`ビズロジ`,`元ファイル`,`bkファイル`,`時間`,`結果`,`担当者`,`PC`"
        setStr = "'" & BZ & "'," & "'" & uploadFile & "'," & "'" & backupFolder & "'," & "'" & time & "'," & "'UK'," & "'" & Environment.UserName & "','" & Environment.MachineName & "'"

        '=================================================
        MALLDB_CONNECT_INSERT_ONE(cnbackupdb_path, TableName, setField, setStr)
        '=================================================
        Dim dr As DataGridViewRow = New DataGridViewRow
        dr.CreateCells(datagridview1)
        dr.Cells(0).Value = time
        dr.Cells(1).Value = BZ
        dr.Cells(2).Value = uploadFile
        dr.Cells(3).Value = backupFolder
        dr.Cells(4).Value = "UK"
        dr.Cells(5).Value = Environment.UserName
        dr.Cells(6).Value = Environment.MachineName

        datagridview1.Rows.Insert(0, dr)

    End Sub

    Private Sub フォームのリセットToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles フォームのリセットToolStripMenuItem.Click
        datagridview1.Rows.Clear()
        TextBox1.Text = ""
        TextBox2.Text = ""
        AzukiControl1.Text = ""
        AzukiControl2.Text = ""
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        ComboBox3.SelectedIndex = -1
        KryptonButton3.Enabled = True
        KryptonButton1.Enabled = True
        loadDgv2()
    End Sub

    'Public Shared Function FileToStream(ByVal fileName As String) As Stream
    '    '打开文件
    '    Dim fileStream As FileStream
    '    fileStream = New FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read)
    '    '读取文件的 byte()
    '    Dim bytes() As Byte = New Byte(fileStream.Length) {}
    '    fileStream.Read(bytes, 0, bytes.Length)
    '    fileStream.Close()
    '    '把 byte()转换成 Stream
    '    Dim stream As Stream = New MemoryStream(bytes)
    '    Return stream
    'End Function

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
            SQLCm1.CommandText = "Select * FROM [" & TableName & "]" & where & orderby & limit
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
        SQLCm1.CommandText = "UPDATE [" & TableName & "] Set " & setStr & where
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

        SQLCm1.CommandText = "Select * FROM [" & TableName & "] WHERE " & where
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

    Private Sub Datagridview1_RowPostPaint(sender As Object, e As DataGridViewRowPostPaintEventArgs) Handles datagridview1.RowPostPaint, datagridview2.RowPostPaint, DataGridView3.RowPostPaint
        ' 行ヘッダのセル領域を、行番号を描画する長方形とする（ただし右端に4ドットのすき間を空ける）
        Dim rect As New Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, sender.RowHeadersWidth - 4, sender.Rows(e.RowIndex).Height)
        ' 上記の長方形内に行番号を縦方向中央＆右詰で描画する　フォントや色は行ヘッダの既定値を使用する
        TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), sender.RowHeadersDefaultCellStyle.Font, rect, sender.RowHeadersDefaultCellStyle.ForeColor, TextFormatFlags.VerticalCenter Or TextFormatFlags.Right)
    End Sub

    Dim BeforeColor As Color() = New Color() {Color.Empty, Color.Empty, Color.Empty, Color.Empty, Color.Empty}
    Private Sub Datagridview1_SelectionChanged(sender As Object, e As EventArgs) Handles datagridview1.SelectionChanged
        If sender.Rows.Count = 0 Then
            Exit Sub
        End If

        Dim dHSender As ArrayList = TM_HEADER_GET(sender)
        Dim dgv As DataGridView

        Dim num As Integer = 0
        Select Case True
            Case sender Is datagridview1
                'num = 0
                dgv = datagridview1
                If dgv.SelectedCells.Count <> 0 Then
                    AzukiControl1.Text = dgv.SelectedCells(0).Value
                End If
        End Select

        '選択行に色を付ける
        For r As Integer = 0 To sender.RowCount - 1
            If sender.Rows(r).DefaultCellStyle.BackColor <> BeforeColor(num) Then
                sender.Rows(r).DefaultCellStyle.BackColor = BeforeColor(num)
            End If
        Next
        Dim selRow As ArrayList = New ArrayList
        For i As Integer = 0 To sender.SelectedCells.Count - 1
            If Not selRow.Contains(sender.SelectedCells(i).RowIndex) Then
                selRow.Add(sender.SelectedCells(i).RowIndex)
            End If
        Next
        For i As Integer = 0 To selRow.Count - 1
            BeforeColor(num) = sender.Rows(selRow(i)).DefaultCellStyle.BackColor
            sender.Rows(selRow(i)).DefaultCellStyle.BackColor = Color.LightCyan
        Next
    End Sub

    Private Sub Datagridview2_SelectionChanged(sender As Object, e As EventArgs) Handles datagridview2.SelectionChanged
        If sender.Rows.Count = 0 Or sender.Rows.Count = 1 Then
            Exit Sub
        End If

        Dim dHSender As ArrayList = TM_HEADER_GET(sender)
        Dim dgv As DataGridView

        Dim num As Integer = 0
        Select Case True
            Case sender Is datagridview2
                'num = 0
                dgv = datagridview2
                If dgv.SelectedCells.Count <> 0 Then
                    AzukiControl2.Text = dgv.SelectedCells(0).Value
                End If
        End Select

        '選択行に色を付ける
        For r As Integer = 0 To sender.RowCount - 1
            If sender.Rows(r).DefaultCellStyle.BackColor <> BeforeColor(num) Then
                sender.Rows(r).DefaultCellStyle.BackColor = BeforeColor(num)
            End If
        Next
        Dim selRow As ArrayList = New ArrayList
        For i As Integer = 0 To sender.SelectedCells.Count - 1
            If Not selRow.Contains(sender.SelectedCells(i).RowIndex) Then
                selRow.Add(sender.SelectedCells(i).RowIndex)
            End If
        Next
        For i As Integer = 0 To selRow.Count - 1
            BeforeColor(num) = sender.Rows(selRow(i)).DefaultCellStyle.BackColor
            sender.Rows(selRow(i)).DefaultCellStyle.BackColor = Color.LightCyan
        Next
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        loadDgv2()
    End Sub

    Private Sub loadDgv2()
        datagridview2.Rows.Clear()
        datagridview2.Columns.Clear()

        Dim TableName As String = "リスト"
        Dim Table As DataTable = MALLDB_CONNECT_SELECT(cnbackupdb_path, TableName)

        If Table.Rows.Count > 0 Then
            Table_to_DGV(datagridview2, Table)
        Else
            datagridview2.Rows.Clear()
            datagridview2.Columns.Clear()
        End If

        If datagridview2.Rows.Count > 0 Then
            'datagridview2.Columns(5).HeaderStyle.Width = Unit.Pixel(120)
            For r As Integer = 0 To datagridview2.Rows.Count - 1
                If datagridview2.Item(5, r).Value = "OK" Then
                    datagridview2.Item(5, r).Style.BackColor = Color.Honeydew
                ElseIf datagridview2.Item(5, r).Value = "NG" Then
                    datagridview2.Item(5, r).Style.BackColor = Color.LightPink
                End If
            Next
        End If
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

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If datagridview2.SelectedCells.Count <> 0 Then
            Dim TableName As String = "リスト"
            For c As Integer = 0 To datagridview2.SelectedCells.Count - 1
                Dim id As String = datagridview2.Item(0, datagridview2.SelectedCells(c).RowIndex).Value
                setStr = "[結果] = 'OK'"
                where = "[ID] = " & id
                MALLDB_CONNECT_UPDATE(cnbackupdb_path, TableName, setStr, where)
            Next
        Else
            Exit Sub
        End If
        loadDgv2()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If datagridview2.SelectedCells.Count <> 0 Then
            Dim TableName As String = "リスト"
            For c As Integer = 0 To datagridview2.SelectedCells.Count - 1
                Dim id As String = datagridview2.Item(0, datagridview2.SelectedCells(c).RowIndex).Value
                setStr = "[結果] = 'NG'"
                where = "[ID] = " & id
                MALLDB_CONNECT_UPDATE(cnbackupdb_path, TableName, setStr, where)
            Next
        Else
            Exit Sub
        End If
        loadDgv2()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        If datagridview2.SelectedCells.Count <> 0 Then
            Dim TableName As String = "リスト"
            For c As Integer = 0 To datagridview2.SelectedCells.Count - 1
                Dim id As String = datagridview2.Item(0, datagridview2.SelectedCells(c).RowIndex).Value
                setStr = "[結果] = 'UK'"
                where = "[ID] = " & id
                MALLDB_CONNECT_UPDATE(cnbackupdb_path, TableName, setStr, where)
            Next
        Else
            Exit Sub
        End If
        loadDgv2()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        openBackUpFolder(1)
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        openBackUpFolder(2)
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        openBackUpFolder(3)
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        openBackUpFolder(4)
    End Sub

    Private Sub openBackUpFolder(index As Integer)
        Select Case index
            Case 1
                System.Diagnostics.Process.Start(ComboBox1.SelectedItem)
            Case 2
                System.Diagnostics.Process.Start(ComboBox2.SelectedItem)
            Case 3
                System.Diagnostics.Process.Start(TextBox5.Text)
            Case 4
                System.Diagnostics.Process.Start(TextBox6.Text)
        End Select

    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        If datagridview2.SelectedCells.Count <> 0 Then
            If MessageBox.Show("確認しましたか？", "キャンセル", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                KryptonButton3.Enabled = False
                KryptonButton1.Enabled = False

                Dim TableName As String = "リスト"
                For c As Integer = 0 To datagridview2.SelectedCells.Count - 1
                    Dim BZ As String = datagridview2.Item(1, datagridview2.SelectedCells(c).RowIndex).Value
                    Dim uploadFile As String = ""
                    uploadFile = datagridview2.Item(3, datagridview2.SelectedCells(c).RowIndex).Value
                    If uploadFile <> "" And checkFileHeader(uploadFile) Then
                        Dim saveFolder As String = ""
                        Dim time_now As String = Format(Now, "yyyy/MM/dd HH:mm:ss")
                        Dim time_csv As String = Format(Now, "yyyyMMddHHmmss") & ".csv"
                        Dim backupFolder As String = ""

                        If BZ = "1号" Then
                            saveFolder = ComboBox1.SelectedItem
                            backupFolder = TextBox5.Text
                        ElseIf BZ = "2号" Then
                            saveFolder = ComboBox2.SelectedItem
                            backupFolder = TextBox6.Text
                        End If

                        'check フォルダー
                        If System.IO.Directory.Exists(saveFolder) = False Then
                            MessageBox.Show("BIZSERVER-PCのフォルダにアクセスできない、ネットワークの問題かもしれない、管理員に連絡してください。")
                            Exit Sub
                        End If

                        If System.IO.Directory.Exists(backupFolder) = False Then
                            MessageBox.Show("バックアップフォルダにアクセスできない、ネットワークの問題かもしれない、管理員に連絡してください。")
                            Exit Sub
                        End If

                        If System.IO.File.Exists(uploadFile) = False Then
                            MessageBox.Show("アップロードファイルが存在しない。")
                            Exit Sub
                        End If

                        '成功的话倒数3秒 防止重发
                        If BZ = "1号" Then
                            saveFolder &= "\" & time_csv
                            backupFolder = TextBox5.Text & "\" & time_csv
                            File.Copy(uploadFile, saveFolder, True)
                            File.Copy(uploadFile, backupFolder, True)

                            System.Threading.Thread.Sleep(3500 * 1)
                        ElseIf BZ = "2号" Then
                            saveFolder &= "\" & time_csv
                            backupFolder = TextBox6.Text & "\" & time_csv
                            File.Copy(uploadFile, saveFolder, True)
                            File.Copy(uploadFile, backupFolder, True)

                            System.Threading.Thread.Sleep(3500 * 1)
                        End If

                        TextBox1.Text = ""
                        TextBox2.Text = ""

                        setField = ""
                        setStr = ""

                        setField = "`ビズロジ`,`元ファイル`,`bkファイル`,`時間`,`結果`,`担当者`,`PC`"
                        setStr = "'" & BZ & "'," & "'" & uploadFile & "'," & "'" & backupFolder & "'," & "'" & time_now & "'," & "'UK'," & "'" & Environment.UserName & "','" & Environment.MachineName & "'"

                        '=================================================
                        MALLDB_CONNECT_INSERT_ONE(cnbackupdb_path, TableName, setField, setStr)
                        '=================================================
                    End If
                Next
            Else
                Exit Sub
            End If
        Else
            Exit Sub
        End If
        loadDgv2()
        KryptonButton3.Enabled = True
        KryptonButton1.Enabled = True
    End Sub

    Private Sub DataGridView3_DragEnter(sender As Object, e As DragEventArgs) Handles DataGridView3.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub DataGridView3_DragDrop(sender As Object, e As DragEventArgs) Handles DataGridView3.DragDrop
        DataGridView3.Rows.Clear()
        Dim dataPathArray As New ArrayList
        For Each filename As String In e.Data.GetData(DataFormats.FileDrop)
            dataPathArray.Add(filename)
        Next

        Dim addflag As Boolean = True

        If dataPathArray.Count > 0 Then
            For i As Integer = 0 To dataPathArray.Count - 1
                If System.IO.Path.GetExtension(dataPathArray(i)) <> ".csv" Then
                    addflag = False
                Else
                    If dataPathArray(i) <> "" And checkFileHeader(dataPathArray(i)) Then
                    Else
                        addflag = False
                    End If
                End If
                DataGridView3.Rows.Add(dataPathArray(i))
            Next
            If addflag = False Then
                MessageBox.Show("ファイルの拡張子あるいは中身を確認してください。")
                DataGridView3.Rows.Clear()
                Exit Sub
            End If
        End If
    End Sub

    Private Sub DataGridView3_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView3.SelectionChanged
        If sender.Rows.Count = 0 Then
            Exit Sub
        End If

        Dim dHSender As ArrayList = TM_HEADER_GET(sender)
        Dim dgv As DataGridView

        Dim num As Integer = 0
        Select Case True
            Case sender Is DataGridView3
                'num = 0
                dgv = DataGridView3
        End Select

        '選択行に色を付ける
        For r As Integer = 0 To sender.RowCount - 1
            If sender.Rows(r).DefaultCellStyle.BackColor <> BeforeColor(num) Then
                sender.Rows(r).DefaultCellStyle.BackColor = BeforeColor(num)
            End If
        Next
        Dim selRow As ArrayList = New ArrayList
        For i As Integer = 0 To sender.SelectedCells.Count - 1
            If Not selRow.Contains(sender.SelectedCells(i).RowIndex) Then
                selRow.Add(sender.SelectedCells(i).RowIndex)
            End If
        Next
        For i As Integer = 0 To selRow.Count - 1
            BeforeColor(num) = sender.Rows(selRow(i)).DefaultCellStyle.BackColor
            sender.Rows(selRow(i)).DefaultCellStyle.BackColor = Color.LightCyan
        Next
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        If ComboBox3.SelectedIndex = -1 Then
            MessageBox.Show("1号あるいは２号を選択してください。")
            Exit Sub
        Else
            If DataGridView3.Rows.Count = 0 Then
                MessageBox.Show("ファイルを入れてください。")
                Exit Sub
            Else
                If MessageBox.Show(ComboBox3.SelectedItem & "に発送しますか？", "キャンセル", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                    KryptonButton3.Enabled = False
                    KryptonButton1.Enabled = False

                    Dim TableName As String = "リスト"
                    For c As Integer = 0 To DataGridView3.SelectedCells.Count - 1

                        Dim uploadFile As String = DataGridView3.Item(0, DataGridView3.SelectedCells(c).RowIndex).Value

                        If uploadFile <> "" And checkFileHeader(uploadFile) Then
                            Dim BZ As String = ""
                            Dim saveFolder As String = ""
                            Dim time_now As String = Format(Now, "yyyy/MM/dd HH:mm:ss")
                            Dim time_csv As String = Format(Now, "yyyyMMddHHmmss") & ".csv"
                            Dim backupFolder As String = ""

                            If ComboBox3.SelectedItem = "1号" Then
                                BZ = "1号"
                                saveFolder = ComboBox1.SelectedItem
                                backupFolder = TextBox5.Text
                            ElseIf ComboBox3.SelectedItem = "2号" Then
                                BZ = "2号"
                                saveFolder = ComboBox2.SelectedItem
                                backupFolder = TextBox6.Text
                            End If

                            'check フォルダー
                            If System.IO.Directory.Exists(saveFolder) = False Then
                                MessageBox.Show("BIZSERVER-PCのフォルダにアクセスできない、ネットワークの問題かもしれない、管理員に連絡してください。")
                                Exit Sub
                            End If

                            If System.IO.Directory.Exists(backupFolder) = False Then
                                MessageBox.Show("バックアップフォルダにアクセスできない、ネットワークの問題かもしれない、管理員に連絡してください。")
                                Exit Sub
                            End If

                            If System.IO.File.Exists(uploadFile) = False Then
                                MessageBox.Show("アップロードファイルが存在しない。")
                                Exit Sub
                            End If

                            '成功的话倒数3秒 防止重发
                            If BZ = "1号" Then
                                saveFolder &= "\" & time_csv
                                backupFolder = TextBox5.Text & "\" & time_csv
                                File.Copy(uploadFile, saveFolder, True)
                                File.Copy(uploadFile, backupFolder, True)

                                System.Threading.Thread.Sleep(3500 * 1)
                            ElseIf BZ = "2号" Then
                                saveFolder &= "\" & time_csv
                                backupFolder = TextBox6.Text & "\" & time_csv
                                File.Copy(uploadFile, saveFolder, True)
                                File.Copy(uploadFile, backupFolder, True)

                                System.Threading.Thread.Sleep(3500 * 1)
                            End If

                            TextBox1.Text = ""
                            TextBox2.Text = ""

                            setField = ""
                            setStr = ""

                            setField = "`ビズロジ`,`元ファイル`,`bkファイル`,`時間`,`結果`,`担当者`,`PC`"
                            setStr = "'" & BZ & "'," & "'" & uploadFile & "'," & "'" & backupFolder & "'," & "'" & time_now & "'," & "'UK'," & "'" & Environment.UserName & "','" & Environment.MachineName & "'"

                            '=================================================
                            MALLDB_CONNECT_INSERT_ONE(cnbackupdb_path, TableName, setField, setStr)
                            '=================================================

                            Dim dr As DataGridViewRow = New DataGridViewRow
                            dr.CreateCells(datagridview1)
                            dr.Cells(0).Value = time_now
                            dr.Cells(1).Value = BZ
                            dr.Cells(2).Value = uploadFile
                            dr.Cells(3).Value = backupFolder
                            dr.Cells(4).Value = "UK"
                            dr.Cells(5).Value = Environment.UserName
                            dr.Cells(6).Value = Environment.MachineName

                            datagridview1.Rows.Insert(0, dr)
                        End If
                    Next
                Else
                    Exit Sub
                End If
            End If
        End If

        loadDgv2()
        KryptonButton3.Enabled = True
        KryptonButton1.Enabled = True
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        DataGridView3.Rows.Clear()
    End Sub
End Class