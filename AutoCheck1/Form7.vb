Imports System.IO
Imports System.Data.OleDb
Imports WiseupSoft.Win.Control

Public Class Form7
    Private Sub DataGridView1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles DataGridView1.DragDrop
        For Each filename As String In e.Data.GetData(DataFormats.FileDrop)
            If DataGridView1.RowCount > 1 Then
                Dim DR As DialogResult = MsgBox("追加しますか", MsgBoxStyle.YesNo Or MsgBoxStyle.SystemModal)
                If DR = Windows.Forms.DialogResult.No Then
                    DataGridView1.Rows.Clear()
                End If
            End If

            Dim csvRecords As New ArrayList()
            csvRecords = CSV_READ(filename)

            Dim cFlag As Integer = 0
            If InStr(csvRecords(0), "受注番号,受注ステータス") > 0 Then '楽天
                cFlag = 1
            ElseIf InStr(csvRecords(0), "OrderId,LineId") > 0 Then 'YahooS
                cFlag = 2
            Else
                cFlag = 0
            End If

            For i As Integer = 0 To csvRecords.Count - 1
                Dim sArray As String() = Split(csvRecords(i), ",")
                If cFlag = 0 Then
                    If IsNumeric(sArray(0)) Then
                        DataGridView1.Rows.Add(sArray)
                    End If
                ElseIf cFlag = 1 Then
                    If IsNumeric(sArray(104)) Then
                        Dim hDate As String = ""
                        If sArray(4) <> "" Then
                            hDate = Format(CDate(sArray(4)), "MM/dd")
                        End If
                        Dim sKei As Integer = sArray(104) * sArray(105)
                        Dim code As String() = Split(sArray(102), "-")
                        Dim code1 As String = ""
                        Dim code2 As String = ""
                        If code.Length = 1 Then
                            code1 = code(0)
                            code2 = "-"
                        Else
                            code1 = code(0)
                            code2 = sArray(102)
                        End If
                        DataGridView1.Rows.Add(hDate, sArray(0), code1, code2, sArray(104), sArray(105), sKei, "")
                    End If
                ElseIf cFlag = 2 Then
                    If IsNumeric(sArray(12)) Then
                        Dim sKei As Integer = sArray(12) * sArray(3)
                        If sArray(5) = "" Then
                            sArray(5) = "-"
                        End If
                        DataGridView1.Rows.Add("", sArray(2), sArray(4), sArray(5), sArray(12), sArray(3), sKei, "")
                    End If
                End If
            Next

            Exit For
        Next

        Kei(DataGridView1)
    End Sub

    Private Sub DataGridView1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles DataGridView1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Function CSV_READ(ByRef path As String)
        Dim csvRecords As New ArrayList()

        'CSVファイル名
        Dim csvFileName As String = path

        'Shift JISで読み込む
        'フィールドが文字で区切られているとする
        'デフォルトでDelimitedなので、必要なし
        '区切り文字を,とする
        'フィールドを"で囲み、改行文字、区切り文字を含めることができるか
        'デフォルトでtrueなので、必要なし
        'フィールドの前後からスペースを削除する
        'デフォルトでtrueなので、必要なし
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
                If str = "" Then
                    str = fields(i)
                Else
                    str &= "," & fields(i)
                End If
            Next
            '保存
            csvRecords.Add(str)
        End While

        '後始末
        tfp.Close()
        Return csvRecords
    End Function

    '保存
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        ToolStripProgressBar1.Maximum = DataGridView1.RowCount
        ToolStripProgressBar1.Value = 0
        countNum.Clear()
        countNum.Add(0)
        countNum.Add(0)

        For r As Integer = 0 To DataGridView1.RowCount - 2
            Dim str1 As String = ""
            Dim str2 As String = ""
            Dim inup As String = ""

            If DataGridView1.Item(8, r).Value = "" Then 'INSERT
                inup = 99999999
            Else 'UPDATE
                inup = DataGridView1.Item(8, r).Value
            End If

            'INSERT用
            If DataGridView1.Item(0, r).Value = "" Then
                str1 &= "#" & Now & "#"
            Else
                str1 &= "#" & DataGridView1.Item(0, r).Value & "#"
            End If
            For i As Integer = 1 To 7
                If DataGridView1.Item(i, r).Value.ToString = "" Then
                    str1 &= ",null"
                Else
                    str1 &= ",'" & DataGridView1.Item(i, r).Value & "'"
                End If
            Next

            'UPDATE用
            If DataGridView1.Item(0, r).Value <> "" Then
                str2 = "[発送日]=#" & DataGridView1.Item(0, r).Value & "#"
            Else
                str2 = "[発送日]=#" & Now & "#"
            End If
            str2 &= ",[受注番号]='" & DataGridView1.Item(1, r).Value & "'"
            str2 &= ",[コード]='" & DataGridView1.Item(2, r).Value & "'"
            str2 &= ",[サブコード]='" & DataGridView1.Item(3, r).Value & "'"
            str2 &= ",[単価]='" & DataGridView1.Item(4, r).Value & "'"
            str2 &= ",[個数]='" & DataGridView1.Item(5, r).Value & "'"
            str2 &= ",[小計]='" & DataGridView1.Item(6, r).Value & "'"
            str2 &= ",[その他]='" & DataGridView1.Item(7, r).Value & "'"

            DB_save(str1, str2, inup, r)
            ToolStripProgressBar1.Value += 1
            Application.DoEvents()
        Next

        MsgBox("上書き " & countNum(0) & "件、新規 " & countNum(1) & "件 保存しました", MsgBoxStyle.SystemModal)
        ToolStripProgressBar1.Value = 0
        ToolStripProgressBar1.Maximum = 100
    End Sub

    '***************************************************************************************************
    'データベース関連
    '***************************************************************************************************
    Public Shared CnAccdb As String = ""
    Dim countNum As New ArrayList

    Private Sub DB_save(ByVal str_line1 As String, ByVal str_line2 As String, ByVal inup As Integer, ByVal r As Integer)
        'データベース登録作業===================================
        CnAccdb = Path.GetDirectoryName(Form1.appPath) & "\db\Database2.accdb"
        '--------------------------------------------------------------
        Dim Cn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CnAccdb)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable
        Cn.Open()
        '--------------------------------------------------------------
        Dim List_Start As Date = CDate(Format(DateAdd(DateInterval.Day, -30, Now), "yyyy/MM/dd 00:00:00"))
        Dim List_End As Date = CDate(Format(DateAdd(DateInterval.Day, 1, Now), "yyyy/MM/dd 00:00:00"))
        Dim DateStr As String = " (発送日 BETWEEN #" & CDate(List_Start) & "# AND #" & CDate(List_End) & "#)"
        Dim bango As String = DataGridView1.Item(1, r).Value
        Dim id1 As String = DataGridView1.Item(2, r).Value
        Dim id2 As String = DataGridView1.Item(3, r).Value
        Dim whereCheck As String = " AND ([受注番号] = '" & bango & "') AND ([コード] = '" & id1 & "') AND ([サブコード] = '" & id2 & "')"
        SQLCm.CommandText = "SELECT * FROM 売行動向 WHERE" & DateStr & whereCheck
        Adapter.Fill(Table)

        If Table.Rows.Count > 0 Then
            inup = Table.Rows(0)("ID")
            SQLCm.CommandText = "UPDATE 売行動向 SET " & str_line2 & " WHERE ID = " & inup
            countNum(0) += 1
        Else
            If inup = 99999999 Then
                SQLCm.CommandText = "INSERT INTO 売行動向 (発送日, 受注番号, コード, サブコード, 単価, 個数, 小計, その他) VALUES (" & str_line1 & ")"
                countNum(1) += 1
            Else
                SQLCm.CommandText = "UPDATE 売行動向 SET " & str_line2 & " WHERE ID = " & inup
                countNum(0) += 1
            End If
        End If
        SQLCm.ExecuteNonQuery()
        '--------------------------------------------------------------
        Table.Dispose()
        Adapter.Dispose()
        SQLCm.Dispose()
        Cn.Dispose()
        '--------------------------------------------------------------
        '=======================================================

        Kei(DataGridView1)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        DataGridView1.Rows.Clear()

        Dim List_Start As Date = CDate(Format(DateTimePicker1.Value, "yyyy/MM/dd 00:00:00"))
        Dim List_End As Date = CDate(Format(DateTimePicker2.Value, "yyyy/MM/dd 23:59:59"))
        Dim DateStr As String = " (発送日 BETWEEN #" & CDate(List_Start) & "# AND #" & CDate(List_End) & "#)"

        '=======================================================
        CnAccdb = Path.GetDirectoryName(Form1.appPath) & "\db\Database2.accdb"
        '--------------------------------------------------------------
        Dim Cn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CnAccdb)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable
        Dim ds As New DataSet
        Cn.Open()
        '--------------------------------------------------------------
        Dim whereCheck As String = ""

        SQLCm.CommandText = "SELECT * FROM 売行動向 WHERE" & DateStr & whereCheck
        Adapter.Fill(Table)

        Dim rowNum As Integer = 0
        For r As Integer = 0 To Table.Rows.Count - 1
            Dim searchFlag As Boolean = True
            If TextBox1.Text <> "" Then
                Dim textSearch As String = ""
                For i As Integer = 0 To Table.Columns.Count - 1
                    textSearch &= Table.Rows(r)(i)
                Next
                If InStr(textSearch, TextBox1.Text) = 0 Then
                    searchFlag = False
                End If
            End If

            If searchFlag = True Then
                DataGridView1.Rows.Add()
                DataGridView1.Item(0, rowNum).Value = Table.Rows(r)("発送日").ToString
                DataGridView1.Item(1, rowNum).Value = Table.Rows(r)("受注番号").ToString
                DataGridView1.Item(2, rowNum).Value = Table.Rows(r)("コード").ToString
                DataGridView1.Item(3, rowNum).Value = Table.Rows(r)("サブコード").ToString
                DataGridView1.Item(4, rowNum).Value = Table.Rows(r)("単価").ToString
                DataGridView1.Item(5, rowNum).Value = Table.Rows(r)("個数").ToString
                DataGridView1.Item(6, rowNum).Value = Table.Rows(r)("小計").ToString
                DataGridView1.Item(7, rowNum).Value = Table.Rows(r)("その他").ToString
                DataGridView1.Item(8, rowNum).Value = Table.Rows(r)("ID").ToString
                rowNum += 1
            End If
        Next
        '--------------------------------------------------------------
        Table.Dispose()
        Adapter.Dispose()
        SQLCm.Dispose()
        Cn.Dispose()
        '--------------------------------------------------------------
        '=======================================================

        Kei(DataGridView1)
    End Sub

    '一括書込
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        For r As Integer = 0 To DataGridView1.RowCount - 2
            DataGridView1.Item(0, r).Value = Format(DateTimePicker3.Value, "MM/dd")
        Next
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        DataGridView2.Rows.Clear()
        TabControl1.SelectedIndex = 1

        Dim List_Start As Date = CDate(Format(DateTimePicker1.Value, "yyyy/MM/dd 00:00:00"))
        Dim List_End As Date = CDate(Format(DateTimePicker2.Value, "yyyy/MM/dd 23:59:59"))
        Dim DateStr As String = " (発送日 BETWEEN #" & CDate(List_Start) & "# AND #" & CDate(List_End) & "#)"

        '=======================================================
        CnAccdb = Path.GetDirectoryName(Form1.appPath) & "\db\Database2.accdb"
        '--------------------------------------------------------------
        Dim Cn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CnAccdb)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable
        Dim ds As New DataSet
        Cn.Open()
        '--------------------------------------------------------------

        Dim whereCheck As String = " GROUP BY [コード]"

        SQLCm.CommandText = "SELECT コード, AVG(単価), SUM(個数), SUM(小計) FROM 売行動向 " & " WHERE" & DateStr & whereCheck
        Adapter.Fill(Table)

        For r As Integer = 0 To Table.Rows.Count - 1
            DataGridView2.Rows.Add()
            DataGridView2.Item(0, r).Value = Table.Rows(r)(0).ToString
            DataGridView2.Item(1, r).Value = Math.Round(Table.Rows(r)(1))
            DataGridView2.Item(2, r).Value = Table.Rows(r)(2).ToString
            DataGridView2.Item(3, r).Value = Table.Rows(r)(3).ToString
            'DataGridView2.Item(1, r).Value = Table.Rows(r)("サブコード").ToString
            'DataGridView2.Item(2, r).Value = Table.Rows(r)("単価").ToString
            'DataGridView2.Item(3, r).Value = Table.Rows(r)("個数").ToString
            'DataGridView2.Item(4, r).Value = Table.Rows(r)("小計").ToString
        Next

        '--------------------------------------------------------------
        Table.Dispose()
        Adapter.Dispose()
        SQLCm.Dispose()
        Cn.Dispose()
        '--------------------------------------------------------------
        '=======================================================

        Kei(DataGridView2)
    End Sub

    Private Sub Kei(ByRef dg As DataGridView)
        TextBox2.Text = 0
        TextBox3.Text = 0

        For r As Integer = 0 To dg.RowCount - 2
            For c As Integer = 0 To dg.ColumnCount - 1
                If dg.Columns(c).HeaderText = "小計" Then
                    TextBox3.Text = CInt(TextBox3.Text) + CInt(dg.Item(c, r).Value)
                ElseIf dg.Columns(c).HeaderText = "個数" Then
                    TextBox2.Text = CInt(TextBox2.Text) + CInt(dg.Item(c, r).Value)
                End If
            Next
        Next
    End Sub

    Private Sub DataGridView2_SortCompare(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewSortCompareEventArgs) Handles DataGridView2.SortCompare
        e.SortResult = CInt(e.CellValue1) - CInt(e.CellValue2)
        e.Handled = True
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If DataGridView2.RowCount > 0 Then
            Dim ur As String = "http://shopping.c.yimg.jp/lib/fkstyle/" & DataGridView2.Item(0, DataGridView2.SelectedCells(0).RowIndex).Value & ".jpg"
            WebBrowser1.Navigate(ur)
        End If
    End Sub

    Private Sub WebBrowser1_DocumentCompleted(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted
        Try
            WebBrowser1.Document.Images(0).Style = "zoom:35%;"
        Catch ex As Exception

        End Try
    End Sub




    Dim nowMonth As Integer = 0

    Private Sub MonthCalendarM1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MonthCalendarM1.Load
        'nowMonth = Now.Month
        'Dim start As Date = CDate(Now.Year & "/" & Now.Month & "/1")
        'DBdate(start)
    End Sub

    Private Sub MonthCalendarM1_SelectedDateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MonthCalendarM1.SelectedDateChanged
        If MonthCalendarM1.SelectedDate.Month <> nowMonth Then
            nowMonth = MonthCalendarM1.SelectedDate.Month
            Dim start As Date = CDate(MonthCalendarM1.SelectedDate.Year & "/" & MonthCalendarM1.SelectedDate.Month & "/1")
            DBdate(start)
        End If
    End Sub

    Private Sub DBdate(ByVal start As Date)
        Dim List_Start As Date = CDate(Format(start, "yyyy/MM/dd 00:00:00")).AddDays(-10)
        Dim endD As Date = DateSerial(start.Year, start.Month + 1, 1).AddDays(-1)
        Dim List_End As Date = CDate(Format(endD, "yyyy/MM/dd 23:59:59")).AddDays(+10)
        Dim DateStr As String = " (発送日 BETWEEN #" & CDate(List_Start) & "# AND #" & CDate(List_End) & "#)"

        '=======================================================
        CnAccdb = Path.GetDirectoryName(Form1.appPath) & "\db\Database2.accdb"
        '--------------------------------------------------------------
        Dim Cn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CnAccdb)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable
        Dim ds As New DataSet
        Cn.Open()
        '--------------------------------------------------------------

        Dim whereCheck As String = " GROUP BY [発送日]"

        SQLCm.CommandText = "SELECT 発送日 FROM 売行動向 " & " WHERE" & DateStr & whereCheck
        'SQLCm.CommandText = "SELECT * FROM 売行動向 WHERE" & DateStr
        Adapter.Fill(Table)
        'MsgBox(Table.Rows.Count, MsgBoxStyle.SystemModal)

        For i As Integer = 1 To endD.Day
            Dim bd As Date = CDate(start.Year & "/" & start.Month & "/" & i)
            Dim bd1 As CustomDate = New CustomDate(CDate(start.Year & "/" & start.Month & "/" & i), Color.Yellow, Color.Black, "入力済")
            Dim bd2 As String = Format(bd, "MM/dd")
            For r As Integer = 0 To Table.Rows.Count - 1
                If bd2 = Table.Rows(r)("発送日") Then
                    MonthCalendarM1.CustomDateAdd(bd1)
                    Exit For
                End If
            Next
        Next
        'MonthCalendar1.UpdateBoldedDates()
        'MonthCalendarM1.Update()

        '--------------------------------------------------------------
        Table.Dispose()
        Adapter.Dispose()
        SQLCm.Dispose()
        Cn.Dispose()
        '--------------------------------------------------------------
        '=======================================================
    End Sub

    '検索
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim dg As DataGridView = DataGridView2

        For r As Integer = 0 To dg.RowCount - 2
            For c As Integer = 0 To dg.ColumnCount - 1
                If InStr(dg.Item(c, r).Value, TextBox4.Text) > 0 Then
                    dg.Item(c, r).Selected = True
                    dg.CurrentCell = dg(c, r)
                    Exit For
                End If
            Next
        Next
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Form3.OptimizeMDB("Database2.accdb")
    End Sub
End Class