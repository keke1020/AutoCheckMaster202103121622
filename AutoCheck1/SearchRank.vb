Imports System.IO
Imports System.Data.OleDb
Imports System.Net

Public Class SearchRank
    Dim tenpo As String() = New String() {"http://item.rakuten.co.jp/patri/"}
    Dim tenpoName As String() = New String() {"通販の暁"}

    Private Sub SearchRank_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'ファイル読み取り
        Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\config\SearchList.csv"
        Dim csvRecords As ArrayList = CSV_READ(fName)
        For r As Integer = 0 To csvRecords.Count - 1
            Dim res As String() = Split(csvRecords(r), "|")
            DataGridView1.Rows.Add(res)
            Dim u As String = Replace(res(0), tenpo(0), "")
            u = Replace(u, "/", "")
            DataGridView2.Rows.Add(u)
        Next
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
            Dim doubleStr As Double = 0
            For i As Integer = 0 To fields.Length - 1
                If i = 0 Then
                    str = fields(i)
                Else
                    str &= "|" & fields(i)
                End If
            Next
            '保存
            csvRecords.Add(str)
        End While

        '後始末
        tfp.Close()
        Return csvRecords
    End Function

    Dim WebRefFlag As Boolean = True
    Dim nowR As Integer = 0

    Private Sub スタートToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles スタートToolStripMenuItem.Click
        For r As Integer = 0 To DataGridView1.RowCount - 2
            Dim code As String = DataGridView2.Item(0, r).Value
            Dim keyword As String = DataGridView1.Item(1, r).Value
            Dim url As String = "http://search.rakuten.co.jp/search/mall/" & keyword

            Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS")
            If ComboBox1.SelectedIndex = 1 Then
                enc = System.Text.Encoding.GetEncoding("EUC-JP")
            ElseIf ComboBox1.SelectedIndex = 0 Or ComboBox1.SelectedIndex = 2 Or ComboBox1.SelectedIndex = 3 Then
                enc = System.Text.Encoding.GetEncoding("utf-8")
            ElseIf ComboBox1.SelectedIndex = 4 Then
                enc = System.Text.Encoding.GetEncoding("Shift_JIS")
            Else
                enc = System.Text.Encoding.Default
            End If
            enc = System.Text.Encoding.GetEncoding("utf-8")

            Dim html As String = ""
            Dim req As WebRequest = WebRequest.Create(url)
            req.UseDefaultCredentials = False
            Dim res As WebResponse = req.GetResponse()
            Dim st As Stream = res.GetResponseStream()
            Dim sr As StreamReader = New StreamReader(st, enc)
            html = sr.ReadToEnd()
            sr.Close()
            st.Close()

            nowR = r
            Dim doc As HtmlAgilityPack.HtmlDocument = New HtmlAgilityPack.HtmlDocument
            doc.LoadHtml(html)
            HTMLGet(doc)
            Application.DoEvents()

            'WebRefFlag = False
            'nowR = r
            'WebBrowser1.Navigate(url)

            'Do
            '    Application.DoEvents()
            'Loop Until WebRefFlag = True

            System.Threading.Thread.Sleep(1000 * 1)
        Next

        MsgBox("検索順位チェック終了しました", MsgBoxStyle.SystemModal)
    End Sub

    Private Sub WebBrowser1_DocumentCompleted(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted
        If sender.Url.ToString <> e.Url.ToString Then
            Exit Sub
        End If

        'HTMLGet(WebBrowser1)
    End Sub

    Private Sub HTMLGet(ByVal doc As HtmlAgilityPack.HtmlDocument)
        Dim res As String = ""
        Dim rArray As String() = Nothing

        Dim selecter As String = ""
        selecter = "//div[@class='rsrDispTxtBoxRight']"
        Dim nodes As HtmlAgilityPack.HtmlNodeCollection = doc.DocumentNode.SelectNodes(selecter)
        rArray = Split(nodes(0).InnerText, "(全 ")
        rArray = Split(rArray(1), "件)")
        DataGridView2.Item(3, nowR).Value = rArray(0)
        Dim all As Integer = rArray(0)

        selecter = "//div[@class='rsrSResultItemTxt']"
        nodes = doc.DocumentNode.SelectNodes(selecter)

        Dim checkURL As String = DataGridView1.Item(0, nowR).Value
        Dim rank As Integer = 0
        For i As Integer = 0 To nodes.Count - 1
            rank += 1
            If InStr(nodes(i).InnerHtml, checkURL) > 0 Then
                DataGridView2.Item(1, nowR).Value = rank
                DataGridView2.Item(2, nowR).Value = Math.Round(100 - ((rank / all) * 100), 2, MidpointRounding.AwayFromZero)
                Exit For
            End If
        Next
    End Sub

    Public Sub HTMLGet2(ByVal wb As WebBrowser)
        Do
            Application.DoEvents()
        Loop Until wb.ReadyState = WebBrowserReadyState.Complete And wb.IsBusy = False

        Dim result As New ArrayList
        Dim allHtml As HtmlDocument = wb.Document

        Dim tag As String = ""
        Dim target As String = ""
        Dim targetStr As String = ""
        Dim res As String = ""
        Dim rArray As String() = Nothing

        tag = "div"
        target = "className"
        targetStr = "rsrDispTxtBoxRight"
        result = HTMLgetTxt(allHtml, tag, target, targetStr, 0)
        rArray = Split(result(0), "(全 ")
        rArray = Split(rArray(1), "件)")
        DataGridView2.Item(3, nowR).Value = rArray(0)
        Dim all As Integer = rArray(0)

        tag = "div"
        target = "className"
        targetStr = "rsrSResultItemTxt"
        result = HTMLgetTxt(allHtml, tag, target, targetStr, 1)
        Dim checkURL As String = DataGridView1.Item(0, nowR).Value
        Dim rank As Integer = 0
        For i As Integer = 0 To result.Count - 1
            rank += 1
            If InStr(result(i), checkURL) > 0 Then
                DataGridView2.Item(1, nowR).Value = rank
                DataGridView2.Item(2, nowR).Value = Math.Round(100 - ((rank / all) * 100), 2, MidpointRounding.AwayFromZero)
                Exit For
            End If
        Next

        WebRefFlag = True
    End Sub

    Public Function HTMLgetTxt(ByVal allHtml As HtmlDocument, ByVal tag As String, ByVal target As String, ByVal targetStr As String, ByVal mode As Integer) As ArrayList
        Dim result As New ArrayList

        For Each el As HtmlElement In allHtml.GetElementsByTagName(tag)
            If el.GetAttribute(target) = targetStr Then
                If mode = 0 Then
                    result.Add(el.InnerText)
                Else
                    result.Add(el.OuterHtml)
                End If
            End If
        Next

        Return result
    End Function

    '***************************************************************************************************
    'データベース関連
    '***************************************************************************************************
    Public Shared CnAccdb As String = ""
    Dim countNum As New ArrayList

    Private Sub 保存ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 保存ToolStripMenuItem.Click
        For r As Integer = 0 To DataGridView2.RowCount - 1

        Next

    End Sub

    Private Sub DB_save(ByVal str_line1 As String, ByVal str_line2 As String, ByVal inup As Integer, ByVal r As Integer)
        'データベース登録作業===================================
        CnAccdb = Path.GetDirectoryName(Form1.appPath) & "\db\Database3.accdb"
        '--------------------------------------------------------------
        Dim Cn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CnAccdb)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable
        Cn.Open()
        '--------------------------------------------------------------
        Dim List_Start As Date = CDate(Format(DateAdd(DateInterval.Day, -30, Now), "yyyy/MM/dd 00:00:00"))
        Dim List_End As Date = CDate(Format(DateAdd(DateInterval.Day, 1, Now), "yyyy/MM/dd 00:00:00"))
        Dim DateStr As String = " (日付 BETWEEN #" & CDate(List_Start) & "# AND #" & CDate(List_End) & "#)"
        Dim bango As String = DataGridView1.Item(0, r).Value
        Dim id As String = DataGridView1.Item(1, r).Value
        Dim whereCheck As String = " AND ([追跡番号] = '" & bango & "') AND ([注文ID] = '" & id & "')"
        SQLCm.CommandText = "SELECT * FROM 発送リスト WHERE" & DateStr & whereCheck
        Adapter.Fill(Table)

        If Table.Rows.Count > 0 Then
            inup = Table.Rows(0)("ID")
            SQLCm.CommandText = "UPDATE 発送リスト SET " & str_line2 & " WHERE ID = " & inup
            countNum(0) += 1
        Else
            If inup = 99999999 Then
                SQLCm.CommandText = "INSERT INTO 発送リスト (追跡番号, 注文ID, 届け先, 結果, 済, 発送日, 追跡日) VALUES (" & str_line1 & ")"
                countNum(1) += 1
            Else
                SQLCm.CommandText = "UPDATE 発送リスト SET " & str_line2 & " WHERE ID = " & inup
                countNum(0) += 1
            End If
            'SQLCm.CommandText = "INSERT INTO 発送リスト (追跡日, 追跡番号, 結果, 注文ID, 済, 発送日) VALUES (null,'402215210005',null,'1002500',null,null)"
        End If

        Try
            SQLCm.ExecuteNonQuery()
        Catch ex As Exception

        End Try

        '--------------------------------------------------------------
        Table.Dispose()
        Adapter.Dispose()
        SQLCm.Dispose()
        Cn.Dispose()
        '--------------------------------------------------------------
        '=======================================================
    End Sub


    '行番号を表示する
    Private Sub DataGridView1_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView1.RowPostPaint
        ' 行ヘッダのセル領域を、行番号を描画する長方形とする（ただし右端に4ドットのすき間を空ける）
        Dim rect As New Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, DataGridView1.RowHeadersWidth - 4, DataGridView1.Rows(e.RowIndex).Height)

        ' 上記の長方形内に行番号を縦方向中央＆右詰で描画する　フォントや色は行ヘッダの既定値を使用する
        TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), DataGridView1.RowHeadersDefaultCellStyle.Font, rect, DataGridView1.RowHeadersDefaultCellStyle.ForeColor, TextFormatFlags.VerticalCenter Or TextFormatFlags.Right)
    End Sub

    '行番号を表示する
    Private Sub DataGridView2_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView2.RowPostPaint
        ' 行ヘッダのセル領域を、行番号を描画する長方形とする（ただし右端に4ドットのすき間を空ける）
        Dim rect As New Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, DataGridView2.RowHeadersWidth - 4, DataGridView2.Rows(e.RowIndex).Height)

        ' 上記の長方形内に行番号を縦方向中央＆右詰で描画する　フォントや色は行ヘッダの既定値を使用する
        TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), DataGridView2.RowHeadersDefaultCellStyle.Font, rect, DataGridView2.RowHeadersDefaultCellStyle.ForeColor, TextFormatFlags.VerticalCenter Or TextFormatFlags.Right)
    End Sub

    Private Sub DataGridView1_Scroll(ByVal sender As Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles DataGridView1.Scroll
        If sender Is DataGridView1 Then
            DataGridView2.FirstDisplayedScrollingRowIndex = DataGridView1.FirstDisplayedScrollingRowIndex
            DataGridView2.FirstDisplayedScrollingColumnIndex = DataGridView1.FirstDisplayedScrollingColumnIndex
        ElseIf sender Is DataGridView2 Then
            DataGridView1.FirstDisplayedScrollingRowIndex = DataGridView2.FirstDisplayedScrollingRowIndex
            DataGridView1.FirstDisplayedScrollingColumnIndex = DataGridView2.FirstDisplayedScrollingColumnIndex
        End If
    End Sub
End Class