Imports System.Data.OleDb
Imports System.IO
Imports Microsoft.VisualBasic.FileIO

'================================================
'------------------------------------------------
'
'汎用プログラムモジュール
'
'------------------------------------------------
'================================================

Module TakModule2
    Public CnAccdb As String = ""

    Dim Cn0 As New OleDbConnection
    Dim SQLCm0 As OleDbCommand
    Dim Adapter0 As New OleDbDataAdapter
    Dim Table0 As New DataTable

    Dim setField As String = ""
    Dim setStr As String = ""
    Dim where As String = ""



    '********************************************
    '--------------------------------------------
    ' データベース関連
    '--------------------------------------------
    '********************************************

    ''' <summary>テーブルが存在しているか確認する</summary>
    ''' <param name="str1">テーブル名、型、所有者、スキーマ</param>
    ''' <param name="str2">テーブル名、型、所有者、スキーマ</param>
    ''' <param name="str3">テーブル名、型、所有者、スキーマ</param>
    ''' <param name="str4">テーブル名、型、所有者、スキーマ</param>
    Public Function TM_DB_TABLE_EXISTS(str1 As String, str2 As String, str3 As String, str4 As String) As DataTable
        Dim schemaTable As DataTable = Cn0.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {str1, str2, str3, str4})
        Return schemaTable
    End Function

    ''' <summary>接続してTableを取得し閉じる</summary>
    ''' <param name="CnAccdb">データベース名</param>
    ''' <param name="sql">SQL文</param>
    Public Function TM_DB_CONNECT_SELECT_SQL(CnAccdb As String, sql As String) As DataTable
        '--------------------------------------------------------------
        Dim Cn1 As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CnAccdb)
        Dim SQLCm1 = Cn1.CreateCommand
        Dim Adapter1 As New OleDbDataAdapter(SQLCm1)
        Dim Table1 As New DataTable
        Cn1.Open()
        '--------------------------------------------------------------
        SQLCm1.CommandText = sql
        Adapter1.Fill(Table1)
        '--------------------------------------------------------------
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
    Public Function TM_DB_CONNECT_SELECT(CnAccdb As String, TableName As String, Optional where As String = "", Optional orderby As String = "", Optional limit As String = "") As DataTable
        '--------------------------------------------------------------
        Dim Cn1 As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CnAccdb)
        Dim SQLCm1 = Cn1.CreateCommand
        Dim Adapter1 As New OleDbDataAdapter(SQLCm1)
        Dim Table1 As New DataTable
        Cn1.Open()
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
        Cn1.Close()
        Adapter1.Dispose()
        SQLCm1.Dispose()
        Cn1.Dispose()
        '--------------------------------------------------------------
        Return Table1
    End Function

    ''' <summary>接続してTableを取得し閉じる（カラムヘッダー変わらない場合）</summary>
    ''' <param name="CnAccdb1">コピー元データベース</param>
    ''' <param name="TableName1">コピー元テーブル名</param>
    ''' <param name="CnAccdb2">コピー先データベース</param>
    ''' <param name="TableName2">コピー先テーブル名</param>
    Public Sub TM_DB_CONNECT_TABLE_COPY(CnAccdb1 As String, TableName1 As String, CnAccdb2 As String, TableName2 As String)
        '--------------------------------------------------------------
        Dim Cn1 As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CnAccdb2)
        Dim SQLCm1 = Cn1.CreateCommand
        Cn1.Open()
        '--------------------------------------------------------------
        SQLCm1.CommandText = "DELETE [" & CnAccdb2 & "].[" & TableName2 & "].* FROM " & "[" & CnAccdb2 & "].[" & TableName2 & "]"
        SQLCm1.ExecuteNonQuery()

        SQLCm1.CommandText = "INSERT INTO [" & CnAccdb2 & "].[" & TableName2 & "] Select * FROM [" & CnAccdb1 & "].[" & TableName1 & "]"
        SQLCm1.ExecuteNonQuery()
        '--------------------------------------------------------------
        Cn1.Close()
        SQLCm1.Dispose()
        Cn1.Dispose()
        '--------------------------------------------------------------
    End Sub

    ''' <summary>INSERT_ONE</summary>
    ''' <param name="CnAccdb">データベース名</param>
    ''' <param name="TableName">テーブル名</param>
    ''' <param name="setField">フィールド名（複数可）</param>
    ''' <param name="setStr">セットする文字（複数可）</param>
    Public Function TM_DB_CONNECT_INSERT_ONE(CnAccdb As String, TableName As String, setField As String, setStr As String) As Boolean
        '--------------------------------------------------------------
        Dim Cn1 As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CnAccdb)
        Dim SQLCm1 = Cn1.CreateCommand
        Cn1.Open()
        '--------------------------------------------------------------
        SQLCm1.CommandText = "INSERT INTO [" & TableName & "] (" & setField & ") VALUES (" & setStr & ")"
        SQLCm1.ExecuteNonQuery()
        '--------------------------------------------------------------
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
    Public Function TM_DB_CONNECT_UPDATE(CnAccdb As String, TableName As String, setStr As String, where As String) As Boolean
        '--------------------------------------------------------------
        Dim Cn1 As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CnAccdb)
        Dim SQLCm1 = Cn1.CreateCommand
        Cn1.Open()
        '--------------------------------------------------------------
        If where <> "" Then
            where = " WHERE (" & where & ")"
        End If
        SQLCm1.CommandText = "UPDATE [" & TableName & "] Set " & setStr & where
        SQLCm1.ExecuteNonQuery()
        '--------------------------------------------------------------
        Cn1.Close()
        SQLCm1.Dispose()
        Cn1.Dispose()
        '--------------------------------------------------------------
        Return True
    End Function

    Public Sub TM_DB_CONNECT_UPSERT(CnAccdb As String, TableName As String, where As String, FieldArray As ArrayList, SetArray As ArrayList)
        '--------------------------------------------------------------
        Dim Cn1 As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CnAccdb)
        Dim SQLCm1 = Cn1.CreateCommand
        Dim Adapter1 As New OleDbDataAdapter(SQLCm1)
        Dim Table1 As New DataTable
        Cn1.Open()
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

        Table1.Clear()

        '--------------------------------------------------------------
        Cn1.Close()
        Table1.Dispose()
        Adapter1.Dispose()
        SQLCm1.Dispose()
        Cn1.Dispose()
        '--------------------------------------------------------------
    End Sub

    Public Function TM_DB_CONNECT_UPSERT_DGV(CnAccdb As String, TableName As String, dgv As DataGridView, KeyCol As Integer, Optional rows As Integer() = Nothing) As Boolean
        '--------------------------------------------------------------
        Dim Cn1 As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CnAccdb)
        Dim SQLCm1 = Cn1.CreateCommand
        Dim Adapter1 As New OleDbDataAdapter(SQLCm1)
        Dim Table1 As New DataTable
        Cn1.Open()
        '--------------------------------------------------------------
        Dim query As String = ""
        For r As Integer = 0 To dgv.Rows.Count - 1
            If Not dgv.Rows(r).IsNewRow Then
                Dim keyStr As String = ""
                If dgv.Columns(0).HeaderText = "ID" Then
                    keyStr &= dgv.Columns(KeyCol).HeaderText & "=" & dgv.Item(KeyCol, r).Value & ""
                Else
                    keyStr &= dgv.Columns(KeyCol).HeaderText & "='" & dgv.Item(KeyCol, r).Value & "'"
                End If
                If dgv.Item(KeyCol, r).Value = "" Then  '新規
                    Table1.Clear()
                Else
                    query = "SELECT ID FROM [" & TableName & "] WHERE " & keyStr
                    SQLCm1.CommandText = query
                    Adapter1.Fill(Table1)
                End If

                If Table1.Rows.Count > 0 Then
                    setStr = ""
                    For c As Integer = 1 To dgv.ColumnCount - 1
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
                        setField &= "[" & dgv.Columns(c).HeaderText & "],"
                        setStr &= "'" & dgv.Item(c, r).Value & "',"
                    Next
                    setField = setField.TrimEnd(",")
                    setStr = setStr.TrimEnd(",")
                    SQLCm1.CommandText = "INSERT INTO [" & TableName & "] (" & setField & ") VALUES (" & setStr & ")"
                    SQLCm1.ExecuteNonQuery()
                End If
            End If
        Next
        '--------------------------------------------------------------
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
    Public Function TM_DB_CONNECT_DELETE(CnAccdb As String, TableName As String, where As String) As Boolean
        '--------------------------------------------------------------
        Dim Cn1 As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CnAccdb)
        Dim SQLCm1 = Cn1.CreateCommand
        Cn1.Open()
        '--------------------------------------------------------------
        If where <> "" Then
            where = " WHERE (" & where & ")"
        End If
        SQLCm1.CommandText = "DELETE FROM [" & TableName & "] " & where
        SQLCm1.ExecuteNonQuery()
        '--------------------------------------------------------------
        Cn1.Close()
        SQLCm1.Dispose()
        Cn1.Dispose()
        '--------------------------------------------------------------
        Return True
    End Function

    ''' <summary>接続維持</summary>
    ''' <param name="CnAccdb">データベース名</param>
    Public Sub TM_DB_CONNECT(CnAccdb As String)
        '--------------------------------------------------------------
        Cn0 = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CnAccdb)
        SQLCm0 = Cn0.CreateCommand
        Adapter0 = New OleDbDataAdapter(SQLCm0)
        Cn0.Open()
        '--------------------------------------------------------------
    End Sub

    ''' <summary>接続解除</summary>
    Public Sub TM_DB_DISCONNECT()
        '--------------------------------------------------------------
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
    Public Function TM_DB_GET_SELECT(TableName As String, where As String, Optional orderby As String = "", Optional limit As String = "") As DataTable
        If where <> "" Then
            where = " WHERE (" & where & ")"
        End If
        If orderby <> "" Then
            orderby = " ORDER BY " & orderby & ""
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
    Public Function TM_DB_GET_STR(TableName As String, getField As String, where As String, Optional orderby As String = "") As DataTable
        If orderby <> "" Then
            orderby = " ORDER BY " & orderby & ""
        End If
        Dim Table As New DataTable
        If where = "" Then
            SQLCm0.CommandText = "SELECT " & getField & " FROM [" & TableName & "]" & orderby
        Else
            SQLCm0.CommandText = "SELECT " & getField & " FROM [" & TableName & "] WHERE (" & where & ")" & orderby
        End If
        Adapter0.Fill(Table)

        Return Table
    End Function

    ''' <summary>接続状態でデータ取得3</summary>
    ''' <param name="TableName">テーブル名</param>
    ''' <param name="sql">sql文</param>
    Public Function TM_DB_GET_SQL(TableName As String, sql As String) As DataTable
        Dim Table As New DataTable
        SQLCm0.CommandText = "SELECT * FROM [" & TableName & "] " & sql
        Adapter0.Fill(Table)

        Return Table
    End Function

    ''' <summary>接続状態でデータ上書き</summary>
    ''' <param name="TableName">テーブル名</param>
    ''' <param name="setField">フィールド名（複数可）</param>
    ''' <param name="setStr">セットする文字（複数可）</param>
    Public Function TM_DB_SET_STR(TableName As String, setField As String, setStr As String)
        Try
            SQLCm0.CommandText = "INSERT INTO [" & TableName & "] (" & setField & ") VALUES (" & setStr & ")"
            SQLCm0.ExecuteNonQuery()
            Return "True"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function TM_DB_UPSERT(TableName As String, where As String, FieldArray As ArrayList, SetArray As ArrayList) As String
        Table0.Rows.Clear()
        Table0.Columns.Clear()
        SQLCm0.CommandText = "SELECT * FROM [" & TableName & "] WHERE " & where
        Adapter0.Fill(Table0)

        If Table0.Rows.Count > 0 Then
            setStr = ""
            For i As Integer = 0 To FieldArray.Count - 1
                setStr &= "[" & FieldArray(i) & "] = '" & SetArray(i) & "',"
            Next
            setStr = setStr.TrimEnd(",")
            SQLCm0.CommandText = "UPDATE [" & TableName & "] SET " & setStr & " WHERE " & where
            SQLCm0.ExecuteNonQuery()

            Return "update"
        Else
            setField = ""
            setStr = ""
            For i As Integer = 0 To FieldArray.Count - 1
                setField &= "[" & FieldArray(i) & "],"
                setStr &= "'" & SetArray(i) & "',"
            Next
            setField = setField.TrimEnd(",")
            setStr = setStr.TrimEnd(",")
            SQLCm0.CommandText = "INSERT INTO [" & TableName & "] (" & setField & ") VALUES (" & setStr & ")"
            SQLCm0.ExecuteNonQuery()

            Return "insert"
        End If
    End Function

    ''' <summary>DELETE</summary>
    ''' <param name="TableName">テーブル名</param>
    ''' <param name="where">where文：WHERE (---)</param>
    Public Function TM_DB_DELETE(TableName As String, where As String) As Boolean
        If where <> "" Then
            where = " WHERE (" & where & ")"
        Else
            Return False
            Exit Function
        End If
        SQLCm0.CommandText = "DELETE FROM [" & TableName & "] " & where
        SQLCm0.ExecuteNonQuery()

        Return True
    End Function



    '********************************************
    '--------------------------------------------
    ' DataGridView関連
    '--------------------------------------------
    '********************************************
    ''子項目をクリックした時、SourceControlがnullになるので先に取得
    'Private contextMenuSourceControl As Control = Nothing
    'Private Sub ContextMenuStrip1_Opened(sender As Object, e As EventArgs) Handles _
    '        ContextMenuStrip1.Opened
    '    Dim menu As ContextMenuStrip = DirectCast(sender, ContextMenuStrip)
    '    'SourceControlプロパティの内容を覚えておく
    '    contextMenuSourceControl = menu.SourceControl
    'End Sub

    Public FocusDGV As DataGridView

    Public Sub DGV_CUT(sender As Object, e As EventArgs)
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        CUTS(FocusDGV)
    End Sub

    Public Sub DGV_COPY(sender As Object, e As EventArgs)
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        COPYS(FocusDGV, 0)
    End Sub

    '「"」で囲んでコピー
    Public Sub DGV_COPY_DUBLE_QUOTATION(sender As Object, e As EventArgs)
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        COPYS(FocusDGV, 1)
    End Sub

    Public Sub DGV_PASTE(sender As Object, e As EventArgs)
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        PASTES(FocusDGV, selCell)
    End Sub

    Public Sub DGV_UPROW_INSERT(sender As Object, e As EventArgs)
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        ROWSADD(FocusDGV, selCell, 0)
    End Sub

    Public Sub DGV_DOWNROW_INSERT(sender As Object, e As EventArgs)
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        ROWSADD(FocusDGV, selCell, 1)
    End Sub

    Public Sub DGV_DOWNROW_CLONE(sender As Object, e As EventArgs)
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

    Public Sub DGV_RIGHTCOL_INSERT(sender As Object, e As EventArgs)
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        COLSADD(FocusDGV, selCell, 0)
    End Sub

    Public Sub DGV_LEFTCOL_INSERT(sender As Object, e As EventArgs)
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        COLSADD(FocusDGV, selCell, 1)
    End Sub

    Public Sub DGV_DELS(sender As Object, e As EventArgs)
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        DELS(FocusDGV, selCell)
    End Sub

    Public Sub DGV_BACKCOLOR_DEL(sender As Object, e As EventArgs)
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        For i As Integer = 0 To FocusDGV.SelectedCells.Count - 1
            FocusDGV.SelectedCells(i).Style.BackColor = Control.DefaultBackColor
        Next
    End Sub

    Dim selChangeFlag As Boolean = True
    Public Sub DGV_SELECT_COL(sender As Object, e As EventArgs)
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        selChangeFlag = False
        ColSelect(FocusDGV, 0)
        selChangeFlag = True
    End Sub

    Public Sub DGV_ROWS_DEL(sender As Object, e As EventArgs)
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        ROWSCUT(FocusDGV, selCell)
    End Sub

    Public Sub DGV_COLS_DEL(sender As Object, e As EventArgs)
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        COLSCUT(FocusDGV, selCell)
    End Sub

    ' ----------------------------------------------------------------------------------
    ' DataGridViewでCtrl-Vキー押下時にクリップボードから貼り付けるための実装↓↓↓↓↓↓
    ' DataGridViewでDelやBackspaceキー押下時にセルの内容を消去するための実装↓↓↓↓↓↓
    ' ----------------------------------------------------------------------------------
    Public ColumnChars() As String
    Public _editingColumn As Integer   '編集中のカラム番号
    Public _editingCtrl As DataGridViewTextBoxEditingControl
    Public innerTextBox As TextBox

    'Private Sub DataGridView_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles _
    '    DataGridView1.EditingControlShowing, DataGridView2.EditingControlShowing, DataGridView3.EditingControlShowing, DataGridView4.EditingControlShowing, DataGridView5.EditingControlShowing,
    '    DataGridView6.EditingControlShowing, DataGridView7.EditingControlShowing, DataGridView8.EditingControlShowing, DataGridView9.EditingControlShowing, DataGridView10.EditingControlShowing,
    '    DataGridView11.EditingControlShowing
    '    If TypeOf e.Control Is DataGridViewTextBoxEditingControl Then
    '        Dim dgv As DataGridView = CType(sender, DataGridView)
    '        innerTextBox = CType(e.Control, TextBox)    '編集のために表示されているコントロールを取得
    '    End If
    '    ' 編集中のカラム番号を保存
    '    _editingColumn = CType(sender, DataGridView).CurrentCellAddress.X
    '    Try
    '        ' 編集中のTextBoxEditingControlにKeyPressイベント設定
    '        _editingCtrl = CType(e.Control, DataGridViewTextBoxEditingControl)
    '        AddHandler _editingCtrl.KeyPress, AddressOf DataGridView_CellKeyPress
    '    Catch
    '    End Try
    'End Sub

    Public Sub DataGridView_CellKeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs)
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

    'Private Sub DataGridView_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles _
    '    DataGridView1.KeyUp, DataGridView2.KeyUp, DataGridView3.KeyUp, DataGridView4.KeyUp, DataGridView5.KeyUp,
    '    DataGridView6.KeyUp, DataGridView7.KeyUp, DataGridView8.KeyUp, DataGridView9.KeyUp, DataGridView10.KeyUp,
    '    DataGridView11.KeyUp
    '    Dim dgv As DataGridView = CType(sender, DataGridView)
    '    Dim selCell = dgv.SelectedCells

    '    If e.KeyCode = Keys.Back Then    ' セルの内容を消去
    '        If DataGridView1.IsCurrentCellInEditMode Then

    '        Else
    '            DELS(dgv, selCell)
    '        End If
    '    ElseIf e.Modifiers = Keys.Control And e.KeyCode = Keys.Enter Then '複数セル一括入力
    '        Dim str As String = dgv.CurrentCell.Value
    '        For i As Integer = 0 To selCell.Count - 1
    '            selCell(i).Value = str
    '        Next
    '    End If
    'End Sub
    ' ----------------------------------------------------------------------------------
    ' DataGridViewでCtrl-Vキー押下時にクリップボードから貼り付けるための実装↑↑↑↑↑↑
    ' DataGridViewでDelやBackspaceキー押下時にセルの内容を消去するための実装↑↑↑↑↑↑
    ' ----------------------------------------------------------------------------------

    Public Sub DELS(ByVal dgv As DataGridView, ByVal selCell As Object)
        For i As Integer = 0 To selCell.Count - 1
            dgv(selCell(i).ColumnIndex, selCell(i).RowIndex).Value = ""
        Next
    End Sub

    ''' <summary>
    ''' 切り取り（汎用版）
    ''' </summary>
    ''' <param name="dgv"></param>
    Public Sub CUTS(ByVal dgv As DataGridView)
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
    Public Sub COPYS(ByVal dgv As DataGridView, ByVal mode As Integer)
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
    Public Sub PASTES(ByVal dgv As DataGridView, ByVal selCell As Object)
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

    Public Sub ROWSADD(ByVal dgv As DataGridView, ByVal selCell As Object, ByVal mode As Integer)
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

    Public Sub COLSADD(ByVal dgv As DataGridView, ByVal selCell As Object, ByVal mode As Integer)
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
    Public Sub ColSelect(dgv As DataGridView, mode As Integer)
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

    Public Sub ROWSCUT(ByVal dgv As DataGridView, ByVal selCell As Object)
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

    Public Sub COLSCUT(ByVal dgv As DataGridView, ByVal selCell As Object)
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

    ''' <summary>
    ''' DataTableをDataGridViewに表示
    ''' </summary>
    ''' <param name="dgv">DataGridView</param>
    ''' <param name="table">DataTable</param>
    Public Sub TM_DATATABLE_TO_DGV(dgv As DataGridView, table As DataTable)
        dgv.Rows.Clear()
        dgv.Columns.Clear()

        For r As Integer = 0 To table.Rows.Count - 1
            If r = 0 Then   'ヘッダー
                For c As Integer = 0 To table.Columns.Count - 1
                    dgv.Columns.Add("c" & c, table.Columns(c).ColumnName)
                Next
            End If

            dgv.Rows.Add()
            For c As Integer = 0 To table.Columns.Count - 1
                dgv.Item(c, r).Value = table.Rows(r)(c).ToString
            Next
        Next
    End Sub

    ''' <summary>
    ''' DataGridViewをDataTableに変換
    ''' </summary>
    ''' <param name="dgv">DataGridView</param>
    ''' <returns>DataTable</returns>
    Public Function TM_DGV_TO_DATATABLE(dgv As DataGridView)
        Dim resTable As New DataTable
        For Each column As DataGridViewColumn In dgv.Columns
            resTable.Columns.Add(CStr(column.HeaderCell.Value))
        Next
        For rowIndex As Integer = 0 To dgv.Rows.Count - 1
            Dim values(dgv.Columns.Count - 1)
            For columnIndex As Integer = 0 To dgv.Columns.Count - 1
                values(columnIndex) = dgv.Rows(rowIndex).Cells(columnIndex).Value
            Next
            resTable.Rows.Add(values)
        Next
        Return resTable
        resTable.Dispose()
    End Function

    ''' <summary>
    ''' ArrayListをStrに変換
    ''' </summary>
    ''' <param name="al">ArrayList</param>
    ''' <param name="br">行毎に改行など追加の有無（default=""）</param>
    ''' <returns>String</returns>
    Public Function TM_ARRAYLIST_TO_STR(al As ArrayList, Optional br As String = "")
        Dim str As String = ""
        For Each s As String In al
            str &= s & br
        Next
        str = str.TrimEnd(br)
        Return str
    End Function

    ''' <summary>
    ''' 文字列のcsv型分割（ダブルコーテーション囲み対応）
    ''' </summary>
    ''' <param name="str">文字列</param>
    ''' <returns>配列</returns>
    Public Function TM_CSV_SPLIT(str As String)
        Dim row As String()
        Dim rs As New StringReader(str)
        Using parser As New TextFieldParser(rs)
            parser.TextFieldType = FieldType.Delimited
            parser.SetDelimiters(",")
            parser.HasFieldsEnclosedInQuotes = True
            parser.TrimWhiteSpace = False
            row = parser.ReadFields()
        End Using
        Return row
    End Function


End Module
