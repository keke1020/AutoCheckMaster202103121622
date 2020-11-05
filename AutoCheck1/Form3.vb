Imports System.Data.OleDb
Imports System.IO
Imports WiseupSoft.Win.Control
Imports System.Net

Public Class Form3
    'Dim home As String = "http://yahoo.co.jp/"
    Dim sagawa As String = "http://k2k.sagawa-exp.co.jp/p/web/okurijosearch.do?okurijoNo="
    'Dim sagawa As String = "http://k2k.sagawa-exp.co.jp/p/sagawa/web/okurijoinput.jsp"
    Dim yamato As String = "http://jizen.kuronekoyamato.co.jp/jizen/servlet/crjz.b.NQ0010?id="
    Dim yubin1 As String = "https://trackings.post.japanpost.jp/services/srv/search/?requestNo1="
    Dim yubin2 As String = "&search=追跡スタート"

    Private Sub Form3_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        AddHandler My.Settings.SettingChanging, AddressOf Settings_SettingChanging

        WebBrowser1.ScriptErrorsSuppressed = True
        'WebBrowser1.Navigate(home)
        DateTimePicker1.Value = DateAdd(DateInterval.Month, -1, Now)
        ComboBox1.SelectedIndex = 0

        'MonthCalendar1更新
        nowMonth = Today.Month
        Dim start As Date = CDate(Today.Year & "/" & Today.Month & "/1")
        DBdate(start)
    End Sub

    'フォームのサイズ最小化バグ修正
    Private Sub Settings_SettingChanging(ByVal sender As Object, ByVal e As System.Configuration.SettingChangingEventArgs)
        If Not Me.WindowState = FormWindowState.Normal Then
            If e.SettingName = "Form3Location" Then
                e.Cancel = True
            End If
        End If
    End Sub

    Dim haitatu As Integer = 0
    Dim compCount As Integer = 0
    Dim kekka As String = ""
    Dim webReadComplete As Boolean = False
    Dim numberArray As New ArrayList
    Dim kensakuTitle1 As String = ""
    Dim kensakuTitle2 As String = ""

    'STOP
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        TextBox2.Text = True
    End Sub

    '追跡開始
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\config\追跡番号.txt"
        Dim tArray As New ArrayList
        Using sr As New System.IO.StreamReader(fName, System.Text.Encoding.Default)
            Do While Not sr.EndOfStream
                Dim s As String = sr.ReadLine
                tArray.Add(s)
            Loop
        End Using

        'System.Threading.Thread.Sleep(1000)
        'WebBrowser1.Refresh()
        'WebBrowser1.Visible = False
        numberArray.Clear()

        Dim mae_tuiseki As String = ""
        Dim rechFlag As Boolean = False
        Dim rechNum As String = ""
        For r As Integer = 0 To DataGridView1.RowCount - 2
            If TextBox2.Text = True Then
                TextBox2.Text = False
                Exit For
            End If
RECHECK:
            ToolStripProgressBar1.Value = 0

            Dim sFlag As Boolean = False
            If CheckBox5.Checked = True Then
                For c As Integer = 0 To DataGridView1.ColumnCount - 1
                    If DataGridView1.Item(c, r).Selected = True Then
                        sFlag = True
                        Exit For
                    End If
                Next
            Else
                sFlag = True
            End If

            Dim cFlag As Boolean = True
            If CheckBox1.Checked = True And DataGridView1.Item(4, r).Value <> "" Then
                cFlag = False
            End If

            If sFlag = True And cFlag = True Then
                compCount = 0
                webReadComplete = False

                If r <> 0 Then
                    If DataGridView1.Item(3, r - 1).Value = "" Then
                        r -= 1
                    End If
                End If

                Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("utf-8")

                '追跡番号認識
                DataGridView1.Item(6, r).Value = Now.ToString
                Dim number As String = DataGridView1.Item(0, r).Value
                If mae_tuiseki <> number Then
                    Dim tuiseki_url As String = ""
                    If ComboBox1.SelectedIndex = 1 Then
                        tuiseki_url = sagawa & number
                        kensakuTitle1 = "佐川急便"
                        kensakuTitle2 = "【お荷物問い合わせサービス】"
                        haitatu = 1
                        enc = System.Text.Encoding.GetEncoding("csWindows31J")
                    ElseIf ComboBox1.SelectedIndex = 2 Then
                        tuiseki_url = yubin1 & number & yubin2
                        kensakuTitle1 = "検索結果"
                        kensakuTitle2 = "日本郵便"
                        haitatu = 2
                        enc = System.Text.Encoding.GetEncoding("utf-8")
                    ElseIf ComboBox1.SelectedIndex = 3 Then
                        tuiseki_url = yamato & number
                        kensakuTitle1 = "クロネコヤマト"
                        kensakuTitle2 = "荷物お問い合わせシステム"
                        haitatu = 3
                        enc = System.Text.Encoding.GetEncoding("csWindows31J")
                    Else
                        If rechFlag = False Then
                            For Each tStr As String In tArray
                                Dim n As String() = Split(tStr, "=")
                                If number.Substring(0, 2) = n(1) Then
                                    Select Case n(0)
                                        Case "佐川"
                                            tuiseki_url = sagawa & number
                                            kensakuTitle1 = "佐川急便"
                                            kensakuTitle2 = "【お荷物問い合わせサービス】"
                                            haitatu = 1
                                            enc = System.Text.Encoding.GetEncoding("csWindows31J")
                                        Case "ゆうメ"
                                            tuiseki_url = yubin1 & number & yubin2
                                            kensakuTitle1 = "検索結果"
                                            kensakuTitle2 = "日本郵便"
                                            haitatu = 2
                                            enc = System.Text.Encoding.GetEncoding("utf-8")
                                        Case Else
                                            tuiseki_url = yamato & number
                                            kensakuTitle1 = "クロネコヤマト"
                                            kensakuTitle2 = "荷物お問い合わせシステム"
                                            haitatu = 3
                                            enc = System.Text.Encoding.GetEncoding("csWindows31J")
                                    End Select

                                    Exit For
                                End If
                            Next
                        Else
                            If rechNum = 1 Then
                                tuiseki_url = yamato & number
                                kensakuTitle1 = "クロネコヤマト"
                                kensakuTitle2 = "荷物お問い合わせシステム"
                                haitatu = 3
                            Else
                                tuiseki_url = sagawa & number
                                kensakuTitle1 = "佐川急便"
                                kensakuTitle2 = "【お荷物問い合わせサービス】"
                                haitatu = 1
                            End If
                        End If
                    End If

                    ToolStripStatusLabel3.Text = r + 1
                    ToolStripStatusLabel1.Text = number
                    ToolStripStatusLabel2.Text = "追跡中"
                    ToolStripProgressBar1.Value = 10

                    Dim doc As HtmlAgilityPack.HtmlDocument = Nothing
                    If haitatu = 2 Then
                        Dim html As String = ""
                        Dim req As WebRequest = WebRequest.Create(tuiseki_url)
                        req.UseDefaultCredentials = False
                        Dim resW As WebResponse = req.GetResponse()
                        Dim st As Stream = resW.GetResponseStream()
                        Dim sr As StreamReader = New StreamReader(st, enc)
                        html = sr.ReadToEnd()
                        sr.Close()
                        st.Close()
                        doc = New HtmlAgilityPack.HtmlDocument
                        doc.LoadHtml(html)
                        TextBox3.Text = html
                        Application.DoEvents()
                    Else
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
                    End If

                    If numberArray.Contains(ToolStripStatusLabel1.Text) = False Then
                        Dim res = ""
                        Dim haiso As String = ""
                        If haitatu = 1 Then
                            res = HTMLgetTxt2("div", 8, 0, "")
                            If InStr(res, "該当なし") > 0 Or res = "error" Then
                                kekka = res
                            Else
                                kekka = res
                                haiso = "佐川"
                            End If
                        ElseIf haitatu = 2 Then
                            res = HTMLgetTxt(doc, "td", 3, 6, "郵便局")
                            If InStr(res, "見つかりません") > 0 Or res = "error" Then
                                kekka = "該当なし"
                            Else
                                kekka = res
                                haiso = "ゆうメ"
                            End If
                        ElseIf haitatu = 3 Then
                            res = HTMLgetTxt2("td", 15, 0, "")
                            If InStr(res, "未登録") > 0 Or res = "error" Then
                                kekka = res
                            Else
                                kekka = res
                                haiso = "ヤマト"
                            End If
                            'MsgBox(res)
                        End If

                        If haiso = "" And rechFlag = False And haitatu <> 3 Then
                            rechFlag = True
                            rechNum = haitatu
                            GoTo RECHECK
                        Else
                            rechFlag = False
                        End If

                        DataGridView1.Item(3, r).Value = kekka
                        DataGridView1.Item(4, r).Value = haiso
                        numberArray.Add(ToolStripStatusLabel1.Text)
                    End If
                Else
                    DataGridView1.Item(3, r).Value = DataGridView1.Item(3, r - 1).Value
                    DataGridView1.Item(4, r).Value = DataGridView1.Item(4, r - 1).Value
                    numberArray.Add(ToolStripStatusLabel1.Text)
                End If
                mae_tuiseki = number
                ToolStripProgressBar1.Value = 100

                System.Threading.Thread.Sleep(1000)
            End If

            'DataGridView1.Rows(r).Selected = True
            'If r > 0 Then
            '    DataGridView1.Rows(r - 1).Selected = False
            'End If
            If r > 10 Then
                DataGridView1.FirstDisplayedScrollingRowIndex = r - 9
            End If
        Next

        'WebBrowser1.Visible = True
        'For i As Integer = 0 To 2
        '    Application.DoEvents()
        '    System.Threading.Thread.Sleep(1000)
        'Next
        ToolStripProgressBar1.Value = 0
        MsgBox("終了しました", MsgBoxStyle.SystemModal)
    End Sub

    Private Function HTMLgetTxt(ByVal doc As HtmlAgilityPack.HtmlDocument, ByVal elName As String, ByVal elNum As Integer, ByVal elList As Integer, ByVal notList As String) As String
        Dim res As String = ""
        Dim resB As String = ""

        Dim selecter As String = ""
        selecter = "//" & elName
        Dim nodes As HtmlAgilityPack.HtmlNodeCollection = doc.DocumentNode.SelectNodes(selecter)

        If elList = 0 Then
            Try
                res = nodes(elNum).InnerText
            Catch ex As Exception
                res = "error"
            End Try
        Else
            For i As Integer = 0 To 100
                Try
                    'Dim resA As String = WebBrowser1.Document.GetElementsByTagName(elName)(elNum + (elList * i)).InnerText
                    Dim resA As String = nodes(elNum + (elList * i)).InnerText
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

    '名前,いくつめか,入れるString
    Private Function HTMLgetTxt2(ByVal elName As String, ByVal elNum As Integer, ByVal elList As Integer, ByVal notList As String) As String
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

    Private Sub WebBrowser1_DocumentCompleted(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted

    End Sub

    Private Sub WebBrowser1_DocumentTitleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles WebBrowser1.DocumentTitleChanged
        compCount += 1
        'ToolStripTextBox1.Text = WebBrowser1.Url.ToString

        'If compCount = 1 Then
        '    MsgBox(kensakuTitle1 & " " & kensakuTitle2 & vbCrLf & WebBrowser1.DocumentTitle, MsgBoxStyle.SystemModal)
        'End If
        If haitatu = 3 Then
            If InStr(WebBrowser1.DocumentTitle, kensakuTitle1) > 0 And InStr(WebBrowser1.DocumentTitle, kensakuTitle2) > 0 And numberArray.Contains(ToolStripStatusLabel1.Text) = False Then
                webReadComplete = True
                System.Threading.Thread.Sleep(2000)
            End If
        Else
            If compCount = 1 And numberArray.Contains(ToolStripStatusLabel1.Text) = False Then
                webReadComplete = True
                System.Threading.Thread.Sleep(2000)
            End If
        End If
    End Sub

    Private Sub DataGridView1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles DataGridView1.DragDrop
        If DataGridView1.RowCount > 1 Then
            Dim DR As DialogResult = MsgBox("追加しますか", MsgBoxStyle.YesNo Or MsgBoxStyle.SystemModal)
            If DR = Windows.Forms.DialogResult.No Then
                DataGridView1.Rows.Clear()
            End If
        End If

        For Each filename As String In e.Data.GetData(DataFormats.FileDrop)
            Dim csvRecords As New ArrayList()
            csvRecords = CSV_READ(filename)

            Dim cFlag As Integer = 0
            If InStr(csvRecords(0), "受注番号,受注ステータス") > 0 Then '楽天
                cFlag = 1
            ElseIf InStr(csvRecords(0), "OrderId,Id") > 0 Then 'Yahoo
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
                    If IsNumeric(sArray(85)) Then
                        Dim hDate As String = ""
                        If sArray(4) <> "" Then
                            hDate = Format(CDate(sArray(4)), "MM/dd")
                        End If
                        DataGridView1.Rows.Add(sArray(85), sArray(0), sArray(93) & " " & sArray(94), "", "", hDate)
                    End If
                ElseIf cFlag = 2 Then
                    If IsNumeric(sArray(45)) Then
                        Dim hDate As String = ""
                        If sArray(48) <> "" Then
                            hDate = Format(CDate(sArray(48)), "MM/dd")
                        End If
                        DataGridView1.Rows.Add(sArray(45), sArray(1), sArray(18), "", "", hDate)
                    End If
                End If
            Next

            'Dim fName As String = filename
            'Using sr As New System.IO.StreamReader(fName, System.Text.Encoding.Default)
            '    Do While Not sr.EndOfStream
            '        Dim s As String = sr.ReadLine
            '        s = Replace(s, vbCrLf, "")
            '        s = Replace(s, vbLf, "")
            '        s = Replace(s, """", "")
            '        Dim sArray As String() = Split(s, ",")

            '        If sArray.Length > 10 Then
            '            DataGridView1.Rows.Add(sArray)
            '        Else
            '            ReDim Preserve sArray(8)
            '            If IsNumeric(sArray(0)) Then
            '                DataGridView1.Rows.Add(sArray)
            '            End If
            '        End If
            '    Loop
            'End Using
        Next

        Me.TopMost = True
        Me.TopMost = False

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
                If i = 0 Then
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

    '一覧クリア
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        DataGridView1.Rows.Clear()
    End Sub

    '一括書込
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        For r As Integer = 0 To DataGridView1.RowCount - 2
            DataGridView1.Item(5, r).Value = Format(DateTimePicker3.Value, "MM/dd")
        Next
    End Sub

    Private Sub DataGridView1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DataGridView1.KeyDown
        Dim dgv As DataGridView = CType(sender, DataGridView)
        Dim x As Integer = dgv.CurrentCellAddress.X
        Dim y As Integer = dgv.CurrentCellAddress.Y

        If e.KeyCode = Keys.Delete Or e.KeyCode = Keys.Back Then
            ' セルの内容を消去
            Dim delR As New ArrayList
            For i As Integer = 0 To dgv.SelectedCells.Count - 1
                dgv.Item(dgv.SelectedCells(i).ColumnIndex, dgv.SelectedCells(i).RowIndex).Value = ""
            Next
        End If
    End Sub

    '行番号を表示する
    Private Sub DataGridView1_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView1.RowPostPaint
        ' 行ヘッダのセル領域を、行番号を描画する長方形とする（ただし右端に4ドットのすき間を空ける）
        Dim rect As New Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, DataGridView1.RowHeadersWidth - 4, DataGridView1.Rows(e.RowIndex).Height)

        ' 上記の長方形内に行番号を縦方向中央＆右詰で描画する　フォントや色は行ヘッダの既定値を使用する
        TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), DataGridView1.RowHeadersDefaultCellStyle.Font, rect, DataGridView1.RowHeadersDefaultCellStyle.ForeColor, TextFormatFlags.VerticalCenter Or TextFormatFlags.Right)
    End Sub

    '***************************************************************************************************
    'データベース関連
    '***************************************************************************************************
    Public Shared CnAccdb As String = ""
    Dim countNum As New ArrayList

    'DB保存
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

            If DataGridView1.Item(7, r).Value = "" Then 'INSERT
                inup = 99999999
            Else 'UPDATE
                inup = DataGridView1.Item(7, r).Value
            End If

            'INSERT用
            For i As Integer = 0 To 5
                If i = 0 Then
                    If DataGridView1.Item(i, r).Value = "" Then
                        str1 = "null"
                    Else
                        str1 = "'" & DataGridView1.Item(i, r).Value & "'"
                    End If
                Else
                    If DataGridView1.Item(i, r).Value = "" Then
                        str1 &= ",null"
                    Else
                        str1 &= ",'" & DataGridView1.Item(i, r).Value & "'"
                    End If
                End If
            Next
            For i As Integer = 6 To 6
                If DataGridView1.Item(i, r).Value = "" Then
                    str1 &= ",#" & Now & "#"
                Else
                    str1 &= ",#" & DataGridView1.Item(i, r).Value & "#"
                End If
            Next

            'UPDATE用
            str2 = "[追跡番号]='" & DataGridView1.Item(0, r).Value & "'"
            str2 &= ",[注文ID]='" & DataGridView1.Item(1, r).Value & "'"
            str2 &= ",[届け先]='" & DataGridView1.Item(2, r).Value & "'"
            str2 &= ",[結果]='" & DataGridView1.Item(3, r).Value & "'"
            str2 &= ",[済]='" & DataGridView1.Item(4, r).Value & "'"
            str2 &= ",[発送日]='" & DataGridView1.Item(5, r).Value & "'"
            If DataGridView1.Item(6, r).Value <> "" Then
                str2 &= ",[追跡日]=#" & DataGridView1.Item(6, r).Value & "#"
            Else
                str2 &= ",[追跡日]=#" & Now & "#"
            End If

            DB_save(str1, str2, inup, r)
            ToolStripProgressBar1.Value += 1
            Application.DoEvents()
        Next

        MsgBox("上書き " & countNum(0) & "件、新規 " & countNum(1) & "件 保存しました", MsgBoxStyle.SystemModal)
        ToolStripProgressBar1.Value = 0
        ToolStripProgressBar1.Maximum = 100
    End Sub

    Private Sub DB_save(ByVal str_line1 As String, ByVal str_line2 As String, ByVal inup As Integer, ByVal r As Integer)
        'データベース登録作業===================================
        CnAccdb = Path.GetDirectoryName(Form1.appPath) & "\db\Database.accdb"
        '--------------------------------------------------------------
        Dim Cn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CnAccdb)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable
        Cn.Open()
        '--------------------------------------------------------------
        Dim List_Start As Date = CDate(Format(DateAdd(DateInterval.Day, -30, Now), "yyyy/MM/dd 00:00:00"))
        Dim List_End As Date = CDate(Format(DateAdd(DateInterval.Day, 1, Now), "yyyy/MM/dd 00:00:00"))
        Dim DateStr As String = " (追跡日 BETWEEN #" & CDate(List_Start) & "# AND #" & CDate(List_End) & "#)"
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

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        DataGridView1.Rows.Clear()

        Dim List_Start As Date = CDate(Format(DateTimePicker1.Value, "yyyy/MM/dd 00:00:00"))
        Dim List_End As Date = CDate(Format(DateTimePicker2.Value, "yyyy/MM/dd 23:59:59"))
        Dim Send_Date As String = Format(DateTimePicker4.Value, "MM/dd")
        Dim DateStr As String = " (追跡日 BETWEEN #" & CDate(List_Start) & "# AND #" & CDate(List_End) & "#)"

        '=======================================================
        CnAccdb = Path.GetDirectoryName(Form1.appPath) & "\db\Database.accdb"
        '--------------------------------------------------------------
        Dim Cn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CnAccdb)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable
        Dim ds As New DataSet
        Cn.Open()
        '--------------------------------------------------------------
        Dim whereCheck As String = ""
        If CheckBox2.Checked = True Then
            whereCheck = " AND (([済] IS NULL) OR ([済] = ''))"
        End If
        Dim sendDate As String = ""
        If CheckBox6.Checked = True Then
            sendDate = " AND (([発送日] = '" & Send_Date & "'))"
        End If

        SQLCm.CommandText = "SELECT * FROM 発送リスト WHERE" & DateStr & whereCheck & sendDate
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
                DataGridView1.Item(0, rowNum).Value = Table.Rows(r)("追跡番号").ToString
                DataGridView1.Item(1, rowNum).Value = Table.Rows(r)("注文ID").ToString
                DataGridView1.Item(2, rowNum).Value = Table.Rows(r)("届け先").ToString
                DataGridView1.Item(3, rowNum).Value = Table.Rows(r)("結果").ToString
                DataGridView1.Item(4, rowNum).Value = Table.Rows(r)("済").ToString
                DataGridView1.Item(5, rowNum).Value = Table.Rows(r)("発送日").ToString
                DataGridView1.Item(6, rowNum).Value = Table.Rows(r)("追跡日").ToString
                DataGridView1.Item(7, rowNum).Value = Table.Rows(r)("ID").ToString
                rowNum += 1
            End If
        Next
        ToolStripStatusLabel3.Text = rowNum
        '--------------------------------------------------------------
        Table.Dispose()
        Adapter.Dispose()
        SQLCm.Dispose()
        Cn.Dispose()
        '--------------------------------------------------------------
        '=======================================================
    End Sub
    '***************************************************************************************************
    'データベース関連
    '***************************************************************************************************

    Dim nowMonth As Integer = 0

    Private Sub MonthCalendarM1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MonthCalendarM1.Load
        'nowMonth = Now.Month
        'Dim start As Date = CDate(Now.Year & "/" & Now.Month & "/1")
        'DBdate(start)
    End Sub

    Private Sub MonthCalendarM1_SelectedDateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MonthCalendarM1.SelectedDateChanged
        DateTimePicker4.Value = MonthCalendarM1.SelectedDate

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
        Dim DateStr As String = " (追跡日 BETWEEN #" & CDate(List_Start) & "# AND #" & CDate(List_End) & "#)"

        '=======================================================
        CnAccdb = Path.GetDirectoryName(Form1.appPath) & "\db\Database.accdb"
        '--------------------------------------------------------------
        Dim Cn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CnAccdb)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable
        Dim ds As New DataSet
        Cn.Open()
        '--------------------------------------------------------------

        Dim whereCheck As String = " GROUP BY [発送日]"

        SQLCm.CommandText = "SELECT 発送日 FROM 発送リスト " & " WHERE" & DateStr & whereCheck
        'SQLCm.CommandText = "SELECT * FROM 発送リスト WHERE" & DateStr
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

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        OptimizeMDB("Database.accdb")
    End Sub

    Public Sub OptimizeMDB(ByVal DBName As String)
        Dim jroEngin As JRO.JetEngine = Nothing
        Const CompactPath As String = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source={0}"
        CnAccdb = Path.GetDirectoryName(Form1.appPath) & "\" & DBName
        Dim MDB_DIR As String = Path.GetDirectoryName(CnAccdb)
        Dim MDB_NAME As String = Path.GetFileName(CnAccdb)

        Try
            'まず、バックアップファイルを作成
            Dim originalMDB As String = MDB_DIR & "\" & MDB_NAME
            Dim cloneMDB As String = originalMDB & ".Backup"
            System.IO.File.Copy(originalMDB, cloneMDB, True)

            '圧縮条件作成
            jroEngin = New JRO.JetEngine
            Dim beforeCompact As String = String.Format(CompactPath, originalMDB)
            Dim afterConpact As String = String.Format(CompactPath, originalMDB.Replace(MDB_NAME, "New_" & MDB_NAME))

            'New_foo.mdbが存在するとエラーが出るのであるなら削除
            If IO.File.Exists(originalMDB.Replace(MDB_NAME, "New_" & MDB_NAME)) Then
                IO.File.Delete(originalMDB.Replace(MDB_NAME, "New_" & MDB_NAME))
            End If

            '圧縮　
            jroEngin.CompactDatabase(beforeCompact, afterConpact)
            jroEngin = Nothing

            '圧縮ファイルに置き換える。
            System.IO.File.Replace(originalMDB.Replace(MDB_NAME, "New_" & MDB_NAME), originalMDB, originalMDB & ".bk")
        Catch ex As Exception
            Throw
        End Try

        MsgBox("最適化が終了しました", MsgBoxStyle.SystemModal)
    End Sub

End Class