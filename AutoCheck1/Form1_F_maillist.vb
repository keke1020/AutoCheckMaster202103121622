Imports System.Data.OleDb
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

Public Class Form1_F_maillist
    Dim ENC_UTF8 As Encoding = Encoding.UTF8
    Public Shared CnAccdb As String = ""

    Private Cn0 As New OleDbConnection
    Private SQLCm0 As OleDbCommand
    Private Adapter0 As New OleDbDataAdapter
    Private Table0 As New DataTable
    Private ds0 As New DataSet

    Private TableName As String = ""
    Private setField As String = ""
    Private setStr As String = ""
    Private where As String = ""

    Private Sub Form1_F_maillist_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For c As Integer = 0 To DGV1.ColumnCount - 1
            DGV1.Columns(c).SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        Dim serverPath As String = Form1.サーバーToolStripMenuItem.Text
        CnAccdb = serverPath & "\update\db\mail.accdb"

        If System.IO.File.Exists(CnAccdb) Then
            Dim Table As DataTable = MALLDB_CONNECT_SELECT(CnAccdb, "メール設定", "")

            For i As Integer = 0 To Table.Rows.Count - 1
                DGV1.Rows.Add(Table.Rows(i)("NAME"), Table.Rows(i)("ADDRESS"), Table.Rows(i)("IMAP"), Table.Rows(i)("PORT"), Table.Rows(i)("PASS"))
            Next
        Else
            MsgBox("メール情報が存在しないです。", MsgBoxStyle.Exclamation)
        End If
    End Sub

    Private Sub フォームのリセットToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles フォームのリセットToolStripMenuItem.Click
        Me.Dispose()
        Dim frm As Form = New Form1_F_maillist
        frm.Show()
    End Sub

    Private Sub 閉じるToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 閉じるToolStripMenuItem1.Click
        Me.Dispose()
    End Sub

    Private Sub 保存ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 保存ToolStripMenuItem.Click
        If DGV1.RowCount = 0 Then
            Exit Sub
        End If
        MALLDB_CONNECT(CnAccdb)
        MALLDB_DELETE("メール設定", "")
        setField = "`NAME`,`ADDRESS`,`IMAP`,`PORT`,`PASS`"
        For c As Integer = 0 To DGV1.RowCount - 1
            setStr = "'" & DGV1.Item(0, c).Value & "'," & "'" & DGV1.Item(1, c).Value & "'," & "'" & DGV1.Item(2, c).Value & "'," & "'" & DGV1.Item(3, c).Value & "'," & "'" & DGV1.Item(4, c).Value & "'"
            MALLDB_CONNECT_INSERT_ONE(CnAccdb, "メール設定", setField, setStr)
        Next
        MALLDB_DISCONNECT()
        MsgBox("保存しました。", MsgBoxStyle.Exclamation)
        ToolStripDropDownButton6.PerformClick()
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

    '行番号を表示する
    Private Sub DataGridView_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles _
            DGV1.RowPostPaint
        Dim rect As New Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, sender.RowHeadersWidth - 4, sender.Rows(e.RowIndex).Height)
        TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), sender.RowHeadersDefaultCellStyle.Font, rect, sender.RowHeadersDefaultCellStyle.ForeColor, TextFormatFlags.VerticalCenter Or TextFormatFlags.Right)
    End Sub
End Class