Imports System.ComponentModel
Imports System.IO
Imports System.Net
Imports System.Text.RegularExpressions

Public Class HtmlDialog_F_TableCreate
    Dim CnAccdb_HD As String = ""

    Private miseArray As New ArrayList
    Private TableName As String = ""
    Private where As String = ""
    Private dH1 As New ArrayList

    Private serverMasterPath As String = "\\SERVER2-PC\orderA\dl_master\master.csv"
    Private localMasterPath As String = Path.GetDirectoryName(appPath) & "\htmltemplate\master.csv"
    Private MasterFile As New ArrayList
    Private mH As New ArrayList
    Private localCB3Path As String = Path.GetDirectoryName(appPath) & "\htmltemplate\各店舗設定.dat"
    Private localTB20Path As String = Path.GetDirectoryName(appPath) & "\htmltemplate\配送送料設定.dat"
    Private localCB6Path As String = Path.GetDirectoryName(appPath) & "\htmltemplate\文字素材.dat"
    Private cb6Array As String() = Nothing

    Private Sub HtmlDialog_F_TableCreate_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim di As New DirectoryInfo(Path.GetDirectoryName(Form1.appPath) & "\htmltemplate")
        Dim files As FileInfo() = di.GetFiles("*.txt", SearchOption.AllDirectories)
        ComboBox1.Items.Clear()
        ComboBox5.Items.Add("テンプレ")
        For Each f As FileInfo In files
            Dim fname As String = Path.GetFileNameWithoutExtension(f.Name)
            If InStr(fname, "テンプレ_") > 0 Then
                ComboBox5.Items.Add(fname)
            ElseIf InStr(fname, "自動") = 0 Then
                ComboBox1.Items.Add(fname)
            End If
        Next

        SplitContainer5.SplitterDistance = 90
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        ComboBox4.SelectedIndex = 0
        ComboBox5.SelectedIndex = 0
        ComboBox7.SelectedIndex = 0
        CnAccdb_HD = Path.GetDirectoryName(Form1.appPath) & "\db\miseDB.accdb"

        'Combobox3
        TextBox18.Text = File.ReadAllText(localCB3Path, encSJ)
        Dim cb3Array As String() = File.ReadAllLines(localCB3Path, encSJ)
        For i As Integer = 1 To cb3Array.Length - 1
            If cb3Array(i) <> "" Then
                miseArray.Add(cb3Array(i))
                Dim mA As String() = Split(miseArray(i - 1), "|=|")
                ComboBox3.Items.Add(mA(0))
            End If
        Next
        ComboBox3.SelectedItem = cb3Array(0)

        'Combobox6
        cb6Array = File.ReadAllLines(localCB6Path, enc64)
        For i As Integer = 0 To cb6Array.Length - 1
            If cb6Array(i) <> "" Then
                If Regex.IsMatch(cb6Array(i), "^//") Then
                    ComboBox6.Items.Add(Replace(cb6Array(i), "//", ""))
                End If
            End If
        Next
        ComboBox6.SelectedIndex = 0

        'Textbox20
        TextBox20.Text = File.ReadAllText(localTB20Path, encSJ)

        'DB
        CnAccdb = Path.GetDirectoryName(Form1.appPath) & "\db\miseDB.accdb"
        dH1 = TM_HEADER_GET(DGV1)

        '在庫数を調べるためのマスター
        Try
            Dim fs As FileInfo = New FileInfo(serverMasterPath)
            Dim fl As FileInfo = New FileInfo(localMasterPath)
            If fs.LastWriteTime > fl.LastWriteTime Then
                File.Copy(serverMasterPath, localMasterPath, True)
            End If
        Catch ex As Exception

        End Try
        MasterFile = TM_CSV_READ(localMasterPath)(0)
        Dim mHeader As String() = Split(MasterFile(0), "|=|")
        For Each header As String In mHeader
            mH.Add(header)
        Next
    End Sub

    Private Sub HtmlDialog_F_TableCreate_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        e.Cancel = True
        Me.Visible = False
    End Sub

    Private Sub DataGridView_RowPostPaint(ByVal sender As DataGridView, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles _
            DGV1.RowPostPaint
        ' 行ヘッダのセル領域を、行番号を描画する長方形とする（ただし右端に4ドットのすき間を空ける）
        Dim rect As New Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, sender.RowHeadersWidth - 4, sender.Rows(e.RowIndex).Height)
        ' 上記の長方形内に行番号を縦方向中央＆右詰で描画する　フォントや色は行ヘッダの既定値を使用する
        TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), sender.RowHeadersDefaultCellStyle.Font, rect, sender.RowHeadersDefaultCellStyle.ForeColor, TextFormatFlags.VerticalCenter Or TextFormatFlags.Right)
    End Sub

    Private Sub TextBox_KeyDown(sender As TextBox, e As KeyEventArgs) Handles _
            TextBox21.KeyDown
        If e.Modifiers = Keys.Control And e.KeyCode = Keys.A Then
            sender.SelectAll()
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Items.Count > 0 Then
            Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\htmltemplate\" & ComboBox1.SelectedItem & ".txt"
            Dim flag As Integer = 0
            Using sr As New System.IO.StreamReader(fName, System.Text.Encoding.Default)
                AzukiControl1.Text = sr.ReadToEnd
            End Using
        End If
    End Sub

    '取得
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim temp As String = AzukiControl1.Text
        Dim tb As TextBox() = New TextBox() {TextBox1, TextBox13, TextBox14, TextBox15, TextBox27, TextBox26, TextBox25, TextBox24}
        Dim kigo As String() = New String() {"A", "B", "C", "D", "E", "F", "G", "H"}

        For i As Integer = 0 To 7
            Dim code As String = tb(i).Text
            If code = "" Then
                Exit For
            Else
                If RadioButton1.Checked Then
                    SetTextDB(kigo(i), code)
                Else
                    SetTextWeb(kigo(i), code)
                End If
            End If
        Next

        AzukiControl1.Text = temp
        MsgBox("取得しました", MsgBoxStyle.SystemModal)
    End Sub

    Private Sub SetTextDB(ByVal num As String, ByVal code As String)
        If num <> "A" And CheckBox1.Checked Then
            AzukiControl1.Text = AzukiControl2.Text
            Application.DoEvents()
        End If

        Dim miseTABLE As String = "T_AKT_item"
        Dim miseURL As String = "patri"
        Select Case ComboBox2.Text
            Case "暁"
                miseTABLE = "T_AKT_item"
                miseURL = "patri"
            Case "アリス"
                miseTABLE = "T_ARS_item"
                miseURL = "alice-zk"
            Case "あかね楽天"
                miseTABLE = "T_AKR_item"
                miseURL = "ashop"
        End Select

        Dim table0 As DataTable = TM_DB_CONNECT_SELECT(CnAccdb_HD, miseTABLE, "[商品管理番号（商品URL）]='" & code & "'")
        If table0.Rows.Count > 0 Then
            Dim url As String = "https://item.rakuten.co.jp/" & miseURL & "/" & code
            AzukiControl2.Text = Replace(AzukiControl1.Text, "$" & num & "1", url)

            Dim imgStr1 As String = table0.Rows(0)("商品画像URL")
            Dim imgStr2 As String = Split(imgStr1, " ")(0)
            AzukiControl2.Text = Replace(AzukiControl2.Text, "$" & num & "2", imgStr2)

            Dim catchCopy As String = table0.Rows(0)("PC用キャッチコピー")
            AzukiControl2.Text = Replace(AzukiControl2.Text, "$" & num & "3", catchCopy)

            Dim kakaku As String = table0.Rows(0)("販売価格")
            AzukiControl2.Text = Replace(AzukiControl2.Text, "$" & num & "4", kakaku)

            Dim resStr As String = table0.Rows(0)("スマートフォン用商品説明文")
            Dim patternStr As String = "<.*?>"
            resStr = Regex.Replace(resStr, patternStr, String.Empty)
            resStr = Split(resStr, "\n")(0)
            resStr = Regex.Replace(resStr, "【.*】", "")
            'resStr = resStr.Substring(0, 30)
            AzukiControl2.Text = Replace(AzukiControl2.Text, "$" & num & "5", resStr)

            BrowserUpdate()
        End If
    End Sub

    Private Sub SetTextWeb(ByVal num As String, ByVal code As String)
        If num <> "A" And CheckBox1.Checked Then
            AzukiControl1.Text = AzukiControl2.Text
            Application.DoEvents()
        End If

        Dim miseURL As String = "patri"
        Select Case ComboBox2.Text
            Case "暁"
                miseURL = "patri"
            Case "アリス"
                miseURL = "alice-zk"
            Case "あかね楽天"
                miseURL = "ashop"
        End Select

        Dim url As String = "https://item.rakuten.co.jp/" & miseURL & "/" & code

        Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS")
        If InStr(url, "rakuten") > 0 Then
            enc = System.Text.Encoding.GetEncoding("EUC-JP")
        ElseIf InStr(url, "yahoo") > 0 Then
            enc = System.Text.Encoding.GetEncoding("utf-8")
        Else
            enc = System.Text.Encoding.Default
        End If

        Dim html As String = ""
        Try
            ServicePointManager.SecurityProtocol = DirectCast(3072, SecurityProtocolType)
            Dim req As WebRequest = WebRequest.Create(url)
            req.UseDefaultCredentials = False
            Dim res As WebResponse = req.GetResponse()
            Dim st As Stream = res.GetResponseStream()
            Dim sr As StreamReader = New StreamReader(st, enc)
            html = sr.ReadToEnd()
            sr.Close()
            st.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.SystemModal)
        End Try

        Dim doc As HtmlAgilityPack.HtmlDocument = New HtmlAgilityPack.HtmlDocument
        doc.LoadHtml(html)

        Dim strArray As String() = Nothing

        AzukiControl2.Text = Replace(AzukiControl1.Text, "$" & num & "1", url)

        '画像を取得
        Dim selecter As String = "//span[@class='sale_desc']"
        Dim resStr As String = HTMLGet(doc, selecter, 1)
        Dim doc2 As HtmlAgilityPack.HtmlDocument = New HtmlAgilityPack.HtmlDocument
        doc2.LoadHtml(resStr)
        selecter = "//img[1]"
        resStr = HTMLGet(doc2, selecter, 2)
        AzukiControl2.Text = Replace(AzukiControl2.Text, "$" & num & "2", resStr)

        'キャッチコピーを取得
        selecter = "//span[@class='catch_copy']"
        resStr = HTMLGet(doc, selecter, 0)
        resStr = Form1.StrConvNumeric(resStr)
        resStr = Form1.StrConvEnglish(resStr)
        resStr = Regex.Replace(resStr, "【.*】", "")
        AzukiControl2.Text = Replace(AzukiControl2.Text, "$" & num & "3", resStr)

        '価格を取得
        selecter = "//span[@class='price2']"
        resStr = HTMLGet(doc, selecter, 0)
        resStr = resStr.Replace(Chr(13), "").Replace(Chr(10), "")
        resStr = Replace(resStr, "円", "")
        AzukiControl2.Text = Replace(AzukiControl2.Text, "$" & num & "4", resStr)

        '説明文を取得
        selecter = "//span[@class='item_desc']"
        resStr = HTMLGet(doc, selecter, 0)
        resStr = Form1.StrConvNumeric(resStr)
        resStr = Form1.StrConvEnglish(resStr)
        resStr = resStr.Replace(Chr(13), "").Replace(Chr(10), "")
        resStr = Regex.Replace(resStr, "【.*】", "")
        resStr = resStr.Substring(0, 30)
        AzukiControl2.Text = Replace(AzukiControl2.Text, "$" & num & "5", resStr)

        BrowserUpdate()
        'MsgBox("取得しました", MsgBoxStyle.SystemModal)
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        TextBox1.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        TextBox15.Text = ""
    End Sub

    Private Function HTMLGet(ByVal doc As HtmlAgilityPack.HtmlDocument, ByVal selecter As String, ByVal inout As Integer)
        Dim num As Integer = 0

        Dim nodes As HtmlAgilityPack.HtmlNodeCollection = doc.DocumentNode.SelectNodes(selecter)
        'For Each node As HtmlAgilityPack.HtmlNode In nodes
        '    MsgBox(node.InnerHtml, MsgBoxStyle.SystemModal)
        'Next

        If Not nodes Is Nothing Then
            If inout = 0 Then
                Return nodes(num).InnerText
            ElseIf inout = 1 Then
                Return nodes(num).InnerHtml
            ElseIf inout = 2 Then
                Return nodes(num).GetAttributeValue("src", "")
            ElseIf inout = 3 Then
                Return nodes
            ElseIf inout = 4 Then
                Return nodes(num).GetAttributeValue("href", "")
            Else
                Return nodes(num).OuterHtml
            End If
        Else
            Return "error"
        End If
    End Function

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        If HTMLcreate.Visible Then
            AzukiControl2.SelectAll()
            AzukiControl2.Copy()
            HTMLcreate.az.Paste()
        End If
    End Sub

    Dim scX As Integer = 0
    Dim scY As Integer = 0
    Private Sub BrowserUpdate()
        Try
            scX = WebBrowser1.Document.Body.ScrollLeft
            scY = WebBrowser1.Document.Body.ScrollTop
        Catch ex As Exception

        End Try

        WebBrowser1.DocumentText = AzukiControl2.Text
        Application.DoEvents()
    End Sub

    Private Sub WebBrowser1_StatusTextChanged(sender As Object, e As EventArgs) Handles WebBrowser1.StatusTextChanged
        Try
            Me.ToolStripStatusLabel1.Text = CType(sender, WebBrowser).StatusText
        Catch ex As Exception

        End Try
    End Sub

    Private ScrollOn As Boolean = True
    Private Sub WebBrowser1_DocumentCompleted(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted
        WebBrowser1.Document.Body.ScrollLeft = scX
        WebBrowser1.Document.Body.ScrollTop = scY

        Zoom()
    End Sub

    Private Sub Zoom()
        Dim MyWeb As Object = WebBrowser1.ActiveXInstance
        MyWeb.ExecWB(63, 1, CType(TrackBar1.Value, Integer), IntPtr.Zero)
    End Sub

    Private Sub TrackBar1_Scroll(sender As Object, e As EventArgs) Handles TrackBar1.Scroll
        Dim MyWeb As Object = WebBrowser1.ActiveXInstance
        MyWeb.ExecWB(63, 1, CType(TrackBar1.Value, Integer), IntPtr.Zero)
        TextBox9.Text = TrackBar1.Value
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TrackBar1.Value = 50
        Dim MyWeb As Object = WebBrowser1.ActiveXInstance
        MyWeb.ExecWB(63, 1, CType(TrackBar1.Value, Integer), IntPtr.Zero)
        TextBox9.Text = TrackBar1.Value
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TrackBar1.Value = 100
        Dim MyWeb As Object = WebBrowser1.ActiveXInstance
        MyWeb.ExecWB(63, 1, CType(TrackBar1.Value, Integer), IntPtr.Zero)
        TextBox9.Text = TrackBar1.Value
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        TrackBar1.Value = 150
        Dim MyWeb As Object = WebBrowser1.ActiveXInstance
        MyWeb.ExecWB(63, 1, CType(TrackBar1.Value, Integer), IntPtr.Zero)
        TextBox9.Text = TrackBar1.Value
    End Sub




    '########################################################################
    '
    ' 自動生成
    '
    '########################################################################

    Private HtmlArray As New ArrayList

    Private Sub TabControl2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl2.SelectedIndexChanged
        If TabControl2.SelectedTab Is Nothing Then
            Exit Sub
        End If

        Select Case TabControl2.SelectedTab.Text
            Case "フリー"
                SplitContainer5.SplitterDistance = 90
            Case "バナー"
                SplitContainer5.SplitterDistance = 274
            Case "タイトル"
                SplitContainer5.SplitterDistance = 153
            Case "文字素材"
                SplitContainer5.SplitterDistance = 217
            Case "商品検索"
                SplitContainer5.SplitterDistance = 427
            Case "全体設定"
                SplitContainer5.SplitterDistance = 427
        End Select
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim listStr As String = ""
        If RadioButton3.Checked Then
            listStr = "テキスト1[|]"
        Else
            listStr = "テキスト2[|]"
        End If
        If Button6.Text = "テキスト追加" Then
            If RadioButton3.Checked Then
                listStr &= File.ReadAllText(text1_path, enc64)
            Else
                listStr &= File.ReadAllText(text2_path, enc64)
            End If
            HtmlArray.Add(listStr)
        Else
            listStr &= AzukiControl3.Text
            HtmlArray(ListBox1.SelectedIndex) = listStr
        End If
        LB1update()
    End Sub

    Private Sub RadioButton6_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton6.CheckedChanged
        TextBox5.Enabled = False
        TextBox8.Enabled = False
        TextBox16.Enabled = False
        TextBox23.Enabled = False
        TextBox28.Enabled = False
        TextBox22.Enabled = False
    End Sub

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged
        TextBox5.Enabled = True
        TextBox8.Enabled = True
        TextBox16.Enabled = True
        TextBox23.Enabled = False
        TextBox28.Enabled = False
        TextBox22.Enabled = False
    End Sub

    Private Sub RadioButton17_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton17.CheckedChanged
        TextBox5.Enabled = True
        TextBox8.Enabled = True
        TextBox16.Enabled = True
        TextBox23.Enabled = True
        TextBox28.Enabled = True
        TextBox22.Enabled = True
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Dim listStr As String = "<br>追加[|]<br>" & vbCrLf
        HtmlArray.Insert(ListBox1.SelectedIndex, listStr)
        LB1update()
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        Dim listStr As String = "<hr>追加[|]<hr>" & vbCrLf
        HtmlArray.Insert(ListBox1.SelectedIndex, listStr)
        LB1update()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim listStr As String = ""
        If RadioButton6.Checked Then
            listStr = "バナー1[|]画像1=" & TextBox4.Text & "|=|リンク1=" & TextBox6.Text & "|=|文字列1=" & TextBox12.Text & "|=|サイズ1=" & NumericUpDown5.Value
        ElseIf RadioButton5.Checked Then
            listStr = "バナー2[|]画像1=" & TextBox4.Text & "|=|リンク1=" & TextBox6.Text & "|=|文字列1=" & TextBox12.Text & "|=|サイズ1=" & NumericUpDown5.Value
            listStr &= "|=|画像2=" & TextBox5.Text & "|=|リンク2=" & TextBox8.Text & "|=|文字列2=" & TextBox16.Text & "|=|サイズ2=" & NumericUpDown6.Value
        Else
            listStr = "バナー3[|]画像1=" & TextBox4.Text & "|=|リンク1=" & TextBox6.Text & "|=|文字列1=" & TextBox12.Text & "|=|サイズ1=" & NumericUpDown5.Value
            listStr &= "|=|画像2=" & TextBox5.Text & "|=|リンク2=" & TextBox8.Text & "|=|文字列2=" & TextBox16.Text & "|=|サイズ2=" & NumericUpDown6.Value
            listStr &= "|=|画像3=" & TextBox22.Text & "|=|リンク3=" & TextBox28.Text & "|=|文字列3=" & TextBox23.Text & "|=|サイズ3=" & NumericUpDown8.Value
        End If
        If Button7.Text = "バナー追加" Then
            HtmlArray.Add(listStr)
        Else
            HtmlArray(ListBox1.SelectedIndex) = listStr
        End If
        LB1update()
    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        RadioButton6.Checked = True
        TextBox12.Text = ""
        TextBox6.Text = ""
        TextBox4.Text = ""
        NumericUpDown5.Value = 100
        TextBox16.Text = ""
        TextBox8.Text = ""
        TextBox5.Text = ""
        NumericUpDown6.Value = 100
        TextBox23.Text = ""
        TextBox28.Text = ""
        TextBox22.Text = ""
        NumericUpDown8.Value = 100
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        TextBox17.Text = Split(miseArray(ComboBox3.SelectedIndex), "|=|")(4)

        Dim str As String() = File.ReadAllLines(localCB3Path, encSJ)
        str(0) = ComboBox3.SelectedItem
        File.WriteAllLines(localCB3Path, str, encSJ)
        TextBox18.Text = File.ReadAllText(localCB3Path, encSJ)
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        Dim str As String = TextBox18.Text
        File.WriteAllText(localCB3Path, str, encSJ)
        MsgBox("保存しました")
    End Sub

    Private Sub DGV1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV1.CellClick
        TextBox19.Text = DGV1.Item(e.ColumnIndex, e.RowIndex).Value.ToString
    End Sub

    Private Sub TextBox19_TextChanged(sender As Object, e As EventArgs) Handles TextBox19.TextChanged
        If DGV1.SelectedCells.Count = 0 Then
            Exit Sub
        End If

        DGV1.SelectedCells(0).Value = TextBox19.Text
    End Sub

    Private Sub CheckBox10_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox10.CheckedChanged
        If CheckBox10.Checked Then
            TextBox29.Enabled = True
        Else
            TextBox29.Enabled = False
        End If
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        TextBox21.Text = Replace(TextBox7.Text, "|", vbCrLf)
        Panel5.Location = New Point(Button24.Location.X + 40, Button24.Location.Y + 40)
        Panel5.Visible = True
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        If TextBox21.Text = "" Then
            Panel5.Visible = False
        Else
            Dim str As String = ""
            Dim strLines As String() = Split(TextBox21.Text, vbCrLf)
            For i As Integer = 0 To strLines.Length - 1
                If strLines(i) <> "" Then
                    str &= strLines(i) & "|"
                End If
            Next
            str = str.TrimEnd("|")
            TextBox7.Text = str
            Panel5.Visible = False
        End If
    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        TextBox21.Text = ""
    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        Panel5.Visible = False
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If ComboBox3.SelectedIndex < 0 Then
            Exit Sub
        ElseIf TextBox7.Text = "" Then
            MsgBox("検索ワードは必須です")
            Exit Sub
        End If

        Dim mall As String = ""
        Dim s1 As String = ""
        Dim s2 As String = ""
        Dim catchCopy As String = ""
        Dim p1 As String = ""
        Dim p2 As String = ""
        Dim img1 As String = ""
        Dim setu1 As String = ""
        Dim souko1 As String = ""
        Dim haisou1 As String = ""
        Select Case Split(miseArray(ComboBox3.SelectedIndex), "|=|")(1)
            Case "楽天"
                mall = "楽天"
                s1 = "商品名"
                s2 = "商品管理番号（商品URL）"
                catchCopy = "PC用キャッチコピー"
                p1 = "販売価格"
                p2 = "表示価格"
                img1 = "商品画像URL"
                setu1 = "PC用商品説明文"
                souko1 = "倉庫指定"
                haisou1 = "配送方法セット管理番号"
            Case "Yahoo"
                mall = "Yahoo"
                s1 = "name"
                s2 = "code"
                catchCopy = "headline"
                p1 = "price"
                p2 = "sale-price"
                img1 = "caption"
                setu1 = "explanation"
                souko1 = "display"
                haisou1 = "postage-set"
            Case "Wowma"
                mall = "Wowma"

        End Select

        TableName = "T_" & Split(miseArray(ComboBox3.SelectedIndex), "|=|")(2) & "_item"
        If InStr(TextBox7.Text, "|") > 0 Then
            where = ""
            Dim tb7Array As String() = Split(TextBox7.Text, "|")
            For i As Integer = 0 To tb7Array.Length - 1
                If i <> 0 Then
                    where &= " or "
                End If
                where &= s1 & " like '%" & tb7Array(i) & "%'"
                where &= " or " & s2 & " like '%" & tb7Array(i) & "%'"
            Next
        Else
            where = s1 & " like '%" & TextBox7.Text & "%'"
            where &= " or " & s2 & " like '%" & TextBox7.Text & "%'"
        End If
        '=================================================
        TM_DB_CONNECT(CnAccdb)
        Dim Table1 As DataTable = TM_DB_GET_SELECT(TableName, where).Copy
        '=================================================

        If Table1.Rows.Count = 0 Then
            TM_DB_DISCONNECT()
            MsgBox("検索結果がありません")
            Exit Sub
        End If

        '一度配列に入れてランダムに並び替える
        Dim dt As String() = Nothing
        Dim dtStr As String = ""
        For r As Integer = 0 To Table1.Rows.Count - 1
            Dim dc As Object = ""
            For c As Integer = 0 To Table1.Columns.Count - 1
                dc &= Table1.Rows(r)(c).ToString & "|=|"
            Next
            If r = 0 Then
                dtStr = dc
            Else
                dtStr &= "[vblf]" & dc
            End If
        Next
        Dim list As String() = Split(dtStr, "[vblf]")
        list = list.OrderBy(Function(s) Guid.NewGuid()).ToArray()

        DGV1.Rows.Clear()
        Dim tH As New ArrayList
        Dim addRow As Integer = 0
        For r As Integer = 0 To Table1.Rows.Count - 1
            If r = 0 Then   'ヘッダー
                For c As Integer = 0 To Table1.Columns.Count - 1
                    tH.Add(Table1.Columns(c).ColumnName)
                Next
            End If
        Next
        '重複チェック
        '------------------------------------------------------------------------
        '商品検索2*5[|]
        'code|=|name|=|baika|=|tujouka|=|img|=|setumei|=|zaiko|=|souko|***|...[|]
        'checkbox|=|regexA|,|regexB|、|...|=|ComboBox3|=|TextBox7|=|TextBox29
        '------------------------------------------------------------------------
        Dim itemListArray As New ArrayList
        If CheckBox11.Checked Then
            For i As Integer = 0 To HtmlArray.Count - 1
                If InStr(HtmlArray(i), "商品検索") > 0 And i <> ListBox1.SelectedIndex Then
                    Dim list1 As String() = Split(HtmlArray(i), "[|]")
                    Dim list2 As String() = Split(list1(1), "|***|")
                    For k As Integer = 0 To list2.Length - 1
                        Dim list3 As String() = Split(list2(k), "|=|")
                        itemListArray.Add(list3(0))
                    Next
                End If
            Next
        End If
        For r As Integer = 0 To list.Length - 1
            Dim listArray As String() = Split(list(r), "|=|")
            Dim soukoA As String = "1"
            If mall = "Yahoo" Then soukoA = "0"
            If listArray(tH.IndexOf(souko1)).ToString = soukoA Then
                Continue For
            ElseIf itemListArray.Contains(listArray(tH.IndexOf(s2)).ToString) Then
                Continue For
            End If
            If Table1.Rows(r)(tH.IndexOf(s1)).ToString <> "" Then
                Dim targetStr As String = listArray(tH.IndexOf(s2)).ToString
                targetStr &= listArray(tH.IndexOf(s1)).ToString
                targetStr &= listArray(tH.IndexOf(catchCopy)).ToString
                If Regex.IsMatch(targetStr, TextBox7.Text) Then
                    DGV1.Rows.Add()
                    '商品コード
                    DGV1.Item(dH1.IndexOf("コード"), addRow).Value = listArray(tH.IndexOf(s2)).ToString
                    '商品名
                    Dim nameStr As String = listArray(tH.IndexOf(catchCopy)).ToString
                    If nameStr = "" Then
                        nameStr = listArray(tH.IndexOf(s1)).ToString
                    End If
                    nameStr = Regex.Replace(nameStr, "【.*?】\ {0,2}", "")
                    DGV1.Item(dH1.IndexOf("商品名"), addRow).Value = nameStr
                    '売価・通常価
                    Dim kakaku1 As String = listArray(tH.IndexOf(p1)).ToString
                    Dim kakaku2 As String = listArray(tH.IndexOf(p2)).ToString
                    If IsNumeric(kakaku1) And IsNumeric(kakaku2) Then
                        If CInt(kakaku1) < CInt(kakaku2) Then
                            DGV1.Item(dH1.IndexOf("売価"), addRow).Value = kakaku1
                            DGV1.Item(dH1.IndexOf("通常価"), addRow).Value = kakaku2
                        Else
                            DGV1.Item(dH1.IndexOf("売価"), addRow).Value = kakaku2
                            DGV1.Item(dH1.IndexOf("通常価"), addRow).Value = kakaku1
                        End If
                    Else
                        DGV1.Item(dH1.IndexOf("売価"), addRow).Value = listArray(tH.IndexOf(p1)).ToString
                        DGV1.Item(dH1.IndexOf("通常価"), addRow).Value = listArray(tH.IndexOf(p2)).ToString
                    End If
                    '画像URL
                    Dim imgStr As String = listArray(tH.IndexOf(img1)).ToString
                    Dim im As MatchCollection = Regex.Matches(imgStr, "http.*?jpg")
                    Dim imgM As String = ""
                    If im.Count > 0 Then
                        'imgM = im(0).Value
                        For m As Integer = 0 To im.Count - 1
                            If Regex.IsMatch(im(m).Value, "[\/|_][a-zA-Z]{1,2}[0-9]{2,4}") Then
                                imgM = im(m).Value
                                Exit For
                            End If
                        Next
                    End If
                    If InStr(imgM, ".jpg") = 0 And mall = "Yahoo" Then
                        imgM = Regex.Match(imgStr, "http.*?_[a-z]{1,2}[0-9]{1,3}").Value
                    ElseIf InStr(imgM, "\n") > 0 And mall = "Yahoo" Then
                        Dim im2 As MatchCollection = Regex.Matches(imgM, "http.*?_[a-z]{1,2}[0-9]{1,3}")
                        If im2.Count > 0 Then
                            imgM = im2(0).Value
                        Else
                            imgM = Regex.Match(imgM, "http.*?jpg").Value
                        End If
                    End If
                    imgM = Regex.Replace(imgM, " $", "")
                    DGV1.Item(dH1.IndexOf("画像"), addRow).Value = imgM
                    '紹介文
                    Dim pcStr As String = listArray(tH.IndexOf(setu1)).ToString
                    If pcStr <> "" Then
                        If RadioButton18.Checked Then
                            Dim pcM As MatchCollection = Regex.Matches(pcStr, "[■|◆|●|・].*?[\<|\\]")
                            Dim pNum As Integer = 0
                            Dim pcMstr As String = ""
                            For p As Integer = 0 To pcM.Count - 1
                                If pcM(p).Value <> "" Then
                                    pcMstr &= Regex.Replace(pcM(p).Value, "■|◆|●|・|\<", "")
                                    pcMstr &= "<br>"
                                End If
                                If pNum = 2 Then
                                    Exit For
                                End If
                                pNum += 1
                            Next
                            pcMstr = Replace(pcMstr, "\", "")     'Yahoo用
                            DGV1.Item(dH1.IndexOf("紹介文"), addRow).Value = pcMstr
                        Else
                            Dim pcM As String = Regex.Match(pcStr, "[■|◆|●|・].*?[\<|\\]").Value
                            pcM = Replace(pcM, "\", "")     'Yahoo用
                            If pcM <> "" Then
                                pcM = Regex.Replace(pcM, "■|◆|●|・|\<", "")
                            End If
                            DGV1.Item(dH1.IndexOf("紹介文"), addRow).Value = pcM
                        End If
                    End If
                    '配送方法
                    Dim haisouStr As String = listArray(tH.IndexOf(haisou1)).ToString
                    If haisouStr <> "" Then
                        DGV1.Item(dH1.IndexOf("配送"), addRow).Value = haisouStr
                    End If

                    Dim code As String = DGV1.Item(dH1.IndexOf("コード"), addRow).Value
                    For Each mLine As String In MasterFile
                        If mLine = "" Then
                            Continue For
                        End If
                        If Regex.IsMatch(mLine, "^" & code) Then
                            Dim mLineArray As String() = Split(mLine, "|=|")
                            Dim freeZaiko As String = mLineArray(mH.IndexOf("フリー在庫"))
                            Dim yoyakuFreeZaiko As String = mLineArray(mH.IndexOf("予約フリー"))
                            If IsNumeric(freeZaiko) And IsNumeric(yoyakuFreeZaiko) Then
                                DGV1.Item(dH1.IndexOf("在庫"), addRow).Value = CInt(freeZaiko) + CInt(yoyakuFreeZaiko)
                                If CInt(freeZaiko) > 0 And CInt(yoyakuFreeZaiko) > 0 Then
                                    DGV1.Item(dH1.IndexOf("在庫"), addRow).Style.BackColor = Color.LightPink
                                ElseIf CInt(freeZaiko) = 0 And CInt(yoyakuFreeZaiko) > 0 Then
                                    DGV1.Item(dH1.IndexOf("在庫"), addRow).Style.BackColor = Color.LightSalmon
                                End If
                            ElseIf IsNumeric(freeZaiko) And Not IsNumeric(yoyakuFreeZaiko) Then
                                DGV1.Item(dH1.IndexOf("在庫"), addRow).Value = CInt(freeZaiko)
                            ElseIf Not IsNumeric(freeZaiko) And IsNumeric(yoyakuFreeZaiko) Then
                                DGV1.Item(dH1.IndexOf("在庫"), addRow).Value = CInt(yoyakuFreeZaiko)
                                DGV1.Item(dH1.IndexOf("在庫"), addRow).Style.BackColor = Color.LightSalmon
                            Else
                                DGV1.Item(dH1.IndexOf("在庫"), addRow).Value = 0
                            End If
                            Exit For
                        End If
                    Next

                    If CInt(DGV1.Item(dH1.IndexOf("在庫"), addRow).Value) < NumericUpDown7.Value Then
                        DGV1.Rows.RemoveAt(addRow)
                    Else
                        addRow += 1
                    End If
                End If
            End If
        Next


        '文字列修正
        'A|,|B|、|C|,|D
        If TextBox17.Text <> "" Then
            Dim rSArray As String() = Split(TextBox17.Text, "|、|")
            For r As Integer = 0 To DGV1.RowCount - 1
                For c As Integer = 0 To DGV1.ColumnCount - 1
                    If Not DGV1.Item(c, r).Value Is Nothing Then
                        Dim str As String = DGV1.Item(c, r).Value
                        For i As Integer = 0 To rSArray.Length - 1
                            Dim rS As String() = Split(rSArray(i), "|,|")
                            str = Regex.Replace(str, rS(0), rS(1))
                        Next
                        DGV1.Item(c, r).Value = str
                    End If
                Next
            Next
        End If

        '=================================================
        TM_DB_DISCONNECT()
        '=================================================
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        If DGV1.RowCount = 0 Then
            Exit Sub
        End If

        Dim str As String = ""
        For r As Integer = 0 To DGV1.RowCount - 1
            If r <> 0 Then
                str &= "|"
            End If
            str &= DGV1.Item(dH1.IndexOf("コード"), r).Value
        Next
        TextBox7.Text = str
    End Sub


    Private syouhin1_path As String = Path.GetDirectoryName(appPath) & "\htmltemplate\自動_フリー_" & "商品検索1.txt"
    Private syouhin1toku_path As String = Path.GetDirectoryName(appPath) & "\htmltemplate\自動_フリー_" & "商品検索1特.txt"
    Private syouhin2_path As String = Path.GetDirectoryName(appPath) & "\htmltemplate\自動_フリー_" & "商品検索2.txt"
    Private syouhin3_path As String = Path.GetDirectoryName(appPath) & "\htmltemplate\自動_フリー_" & "商品検索3.txt"
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If DGV1.RowCount = 0 Then
            Exit Sub
        End If

        Dim listStr As String = ""
        Dim yoko As Integer = 0
        If RadioButton11.Checked Then
            listStr = "商品検索1*" & NumericUpDown1.Value & "[|]"
            yoko = 1
        ElseIf RadioButton10.Checked Then
            listStr = "商品検索2*" & NumericUpDown1.Value & "[|]"
            yoko = 2
        ElseIf RadioButton12.Checked Then
            listStr = "商品検索3*" & NumericUpDown1.Value & "[|]"
            yoko = 3
        ElseIf RadioButton18.Checked Then
            listStr = "商品検索1特*" & NumericUpDown1.Value & "[|]"
            yoko = 1
        End If

        Dim dgvStr As String = ""
        For r As Integer = 0 To yoko * NumericUpDown1.Value - 1
            If r > DGV1.RowCount - 1 Then
                Exit For
            End If

            If r <> 0 Then
                dgvStr &= "|***|"
            End If
            For c As Integer = 0 To DGV1.ColumnCount - 1
                If c <> 0 Then
                    dgvStr &= "|=|"
                End If
                If Not DGV1.Item(c, r).Value Is Nothing Then
                    dgvStr &= DGV1.Item(c, r).Value.ToString
                End If
            Next
        Next
        listStr &= dgvStr

        listStr &= "[|]"
        If CheckBox2.Checked Then listStr &= "1" Else listStr &= "0"
        If CheckBox3.Checked Then listStr &= "1" Else listStr &= "0"
        If CheckBox5.Checked Then listStr &= "1" Else listStr &= "0"
        If CheckBox6.Checked Then listStr &= "1" Else listStr &= "0"
        If CheckBox4.Checked Then listStr &= "1" Else listStr &= "0"
        If CheckBox7.Checked Then listStr &= "1" Else listStr &= "0"
        If CheckBox8.Checked Then listStr &= "1" Else listStr &= "0"
        If CheckBox9.Checked Then listStr &= "1" Else listStr &= "0"
        If CheckBox10.Checked Then listStr &= "1" Else listStr &= "0"
        listStr &= "|=|" & TextBox17.Text
        listStr &= "|=|" & ComboBox3.SelectedItem
        listStr &= "|=|" & TextBox7.Text
        listStr &= "|=|" & TextBox29.Text

        If Button9.Text = "商品欄追加" Then
            HtmlArray.Add(listStr)
        Else
            HtmlArray(ListBox1.SelectedIndex) = listStr
        End If
        LB1update()
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        If ColorDialog1.ShowDialog() = DialogResult.OK Then
            TextBox10.Text = String.Format("#{0:X2}{1:X2}{2:X2}", ColorDialog1.Color.R, ColorDialog1.Color.G, ColorDialog1.Color.B)
            TextBox10.BackColor = ColorDialog1.Color
        End If
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        If ColorDialog1.ShowDialog() = DialogResult.OK Then
            TextBox11.Text = String.Format("#{0:X2}{1:X2}{2:X2}", ColorDialog1.Color.R, ColorDialog1.Color.G, ColorDialog1.Color.B)
            TextBox11.BackColor = ColorDialog1.Color
        End If
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Dim listStr As String = "タイトル[|]"
        If RadioButton14.Checked Then
            listStr &= "ベタ背景"
        ElseIf RadioButton13.Checked Then
            listStr &= "アンダーライン"
        ElseIf RadioButton15.Checked Then
            listStr &= "左ライン"
        Else
            listStr &= "囲み"
        End If
        listStr &= "|=|" & TextBox10.Text & "|=|" & TextBox11.Text
        listStr &= "|=|" & TextBox3.Text
        listStr &= "|=|" & ComboBox4.SelectedItem
        If Button12.Text = "タイトル追加" Then
            HtmlArray.Add(listStr)
        Else
            HtmlArray(ListBox1.SelectedIndex) = listStr
        End If
        LB1update()
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        ListBox2.Items.Clear()
        Dim flag As Boolean = False
        For i As Integer = 0 To cb6Array.Length - 1
            If cb6Array(i) <> "" Then
                If Regex.IsMatch(cb6Array(i), "^//" & ComboBox6.SelectedItem) Then
                    flag = True
                ElseIf Regex.IsMatch(cb6Array(i), "^//") Then
                    If flag Then
                        Exit For
                    End If
                    flag = False
                ElseIf flag Then
                    Dim lb2Array As String() = Split(cb6Array(i), "|=|")
                    ListBox2.Items.Add(lb2Array(0))
                End If
            End If
        Next
    End Sub

    Private Sub LinkLabel3_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        If ColorDialog1.ShowDialog() = DialogResult.OK Then
            TextBox31.Text = String.Format("#{0:X2}{1:X2}{2:X2}", ColorDialog1.Color.R, ColorDialog1.Color.G, ColorDialog1.Color.B)
            TextBox31.BackColor = ColorDialog1.Color
        End If
    End Sub

    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged
        TextBox30.Text = ""
        For i As Integer = 0 To cb6Array.Length - 1
            If cb6Array(i) <> "" Then
                Dim lb2Array As String() = Split(cb6Array(i), "|=|")
                If ListBox2.SelectedItem = lb2Array(0) Then
                    Dim tb30txt As String = lb2Array(1)
                    TextBox30.Text = Replace(tb30txt, "<br>", vbCrLf)
                    AzukiControl3.Text = TextBox30.Text
                    Exit For
                End If
            End If
        Next
    End Sub

    Private Sub Button29_Click(sender As Object, e As EventArgs) Handles Button29.Click
        Dim listStr As String = "文字素材[|]"
        listStr &= AzukiControl3.Text
        listStr &= "[|]" & ComboBox6.SelectedItem & "|=|" & ListBox2.SelectedItem & "|=|" & TextBox31.Text & "|=|" & ComboBox7.SelectedItem & "|=|" & NumericUpDown9.Value
        If Button29.Text = "文字素材追加" Then
            HtmlArray.Add(listStr)
        Else
            HtmlArray(ListBox1.SelectedIndex) = listStr
        End If
        LB1update()
    End Sub


    Private Sub LB1update()
        Dim mae As Integer = ListBox1.SelectedIndex
        Try
            scX = WebBrowser1.Document.Body.ScrollLeft
            scY = WebBrowser1.Document.Body.ScrollTop
        Catch ex As Exception

        End Try
        ListBox1.Items.Clear()
        For Each hA As String In HtmlArray
            ListBox1.Items.Add(Split(hA, "[|]")(0))
        Next
        ListBox1.Items.Add("新規追加")
        HtmlUpdate()
        If mae > ListBox1.Items.Count - 1 Then
            ListBox1.SelectedIndex = 0
        Else
            ListBox1.SelectedIndex = mae
        End If
    End Sub

    Private text1_path As String = Path.GetDirectoryName(appPath) & "\htmltemplate\自動_フリー_" & "テキスト1.txt"
    Private text2_path As String = Path.GetDirectoryName(appPath) & "\htmltemplate\自動_フリー_" & "テキスト2.txt"
    Private banner1_path As String = Path.GetDirectoryName(appPath) & "\htmltemplate\自動_フリー_" & "バナー1.txt"
    Private banner2_path As String = Path.GetDirectoryName(appPath) & "\htmltemplate\自動_フリー_" & "バナー2.txt"
    Private banner3_path As String = Path.GetDirectoryName(appPath) & "\htmltemplate\自動_フリー_" & "バナー3.txt"
    Private mojisozai_path As String = Path.GetDirectoryName(appPath) & "\htmltemplate\自動_フリー_" & "文字素材.txt"
    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        Dim str As String = ""

        If ListBox1.SelectedIndex = ListBox1.Items.Count - 1 Or ListBox1.SelectedIndex = -1 Then
            GroupEnabled("全て")
            Exit Sub
        End If

        Dim lb1_selectArray As String() = Split(HtmlArray(ListBox1.SelectedIndex), "[|]")
        Dim lb1_pattern As String = ""
        Dim lb1_option As String = ""
        If lb1_selectArray.Length = 2 Then
            lb1_pattern = lb1_selectArray(1)
        ElseIf lb1_selectArray.Length = 3 Then
            lb1_pattern = lb1_selectArray(1)
            lb1_option = lb1_selectArray(2)
        End If

        Select Case True
            Case Regex.IsMatch(ListBox1.SelectedItem, "テキスト")
                TabControl2.SelectedTab = TabPage3
                If ListBox1.SelectedItem = "テキスト1" Then
                    RadioButton3.Checked = True
                Else
                    RadioButton4.Checked = True
                End If
                str = lb1_selectArray(1)
                Button6.Text = "テキスト修正"
            Case Regex.IsMatch(ListBox1.SelectedItem, "\<br\>|\<hr\>")
                TabControl2.SelectedTab = TabPage3
            Case Regex.IsMatch(ListBox1.SelectedItem, "バナー")
                TabControl2.SelectedTab = TabPage6
                If ListBox1.SelectedItem = "バナー1" Then
                    str = File.ReadAllText(banner1_path, enc64)
                    RadioButton6.Checked = True
                ElseIf ListBox1.SelectedItem = "バナー2" Then
                    str = File.ReadAllText(banner2_path, enc64)
                    RadioButton5.Checked = True
                Else
                    str = File.ReadAllText(banner3_path, enc64)
                    RadioButton17.Checked = True
                End If
                Dim lb1_patternArray As String() = Split(lb1_pattern, "|=|")
                For i As Integer = 0 To lb1_patternArray.Length - 1
                    Dim pA As String() = Split(lb1_patternArray(i), "=")
                    str = Replace(str, "[" & pA(0) & "]", pA(1))
                    If i = 0 Then TextBox4.Text = pA(1)
                    If i = 1 Then TextBox6.Text = pA(1)
                    If i = 2 Then TextBox12.Text = pA(1)
                    If i = 3 Then NumericUpDown5.Value = pA(1)
                    If i = 4 Then TextBox5.Text = pA(1)
                    If i = 5 Then TextBox8.Text = pA(1)
                    If i = 6 Then TextBox16.Text = pA(1)
                    If i = 7 Then NumericUpDown6.Value = pA(1)
                    If i = 8 Then TextBox22.Text = pA(1)
                    If i = 9 Then TextBox28.Text = pA(1)
                    If i = 10 Then TextBox23.Text = pA(1)
                    If i = 11 Then NumericUpDown8.Value = pA(1)
                Next
                Button7.Text = "バナー修正"
            Case Regex.IsMatch(ListBox1.SelectedItem, "タイトル")
                TabControl2.SelectedTab = TabPage7
                Dim lb1_patternArray As String() = Split(lb1_pattern, "|=|")
                If lb1_patternArray(0) = "ベタ背景" Then RadioButton14.Checked = True
                If lb1_patternArray(0) = "アンダーライン" Then RadioButton13.Checked = True
                If lb1_patternArray(0) = "左ライン" Then RadioButton15.Checked = True
                If lb1_patternArray(0) = "囲み" Then RadioButton16.Checked = True
                TextBox10.Text = lb1_patternArray(1)
                Dim colRest = New ColorConverter().ConvertFromString(lb1_patternArray(1))
                TextBox10.BackColor = colRest
                TextBox11.Text = lb1_patternArray(2)
                colRest = New ColorConverter().ConvertFromString(lb1_patternArray(2))
                TextBox11.BackColor = colRest
                TextBox3.Text = lb1_patternArray(3)
                ComboBox4.SelectedItem = lb1_patternArray(4)
                Button12.Text = "タイトル修正"
            Case Regex.IsMatch(ListBox1.SelectedItem, "文字素材")
                TabControl2.SelectedTab = TabPage8
                AzukiControl3.Text = lb1_pattern
                Dim lb1_optionArray As String() = Split(lb1_option, "|=|")
                ComboBox6.SelectedItem = lb1_optionArray(0)
                Application.DoEvents()
                ListBox2.SelectedItem = lb1_optionArray(1)
                TextBox31.Text = lb1_optionArray(2)
                Dim colRest = New ColorConverter().ConvertFromString(lb1_optionArray(2))
                TextBox31.BackColor = colRest
                ComboBox7.SelectedItem = lb1_optionArray(3)
                NumericUpDown9.Value = lb1_optionArray(4)
                Button29.Text = "文字素材修正"
            Case Regex.IsMatch(ListBox1.SelectedItem, "商品検索")
                TabControl2.SelectedTab = TabPage5
                Dim lb1 As String() = Split(ListBox1.SelectedItem, "*")
                If lb1(0) = "商品検索1" Then
                    RadioButton11.Checked = True
                ElseIf lb1(0) = "商品検索2" Then
                    RadioButton10.Checked = True
                ElseIf lb1(0) = "商品検索3" Then
                    RadioButton12.Checked = True
                ElseIf lb1(0) = "商品検索1特" Then
                    RadioButton18.Checked = True
                End If
                NumericUpDown1.Value = lb1(1)
                If lb1_option.Substring(0, 1) = "1" Then CheckBox2.Checked = True Else CheckBox2.Checked = False
                If lb1_option.Substring(1, 1) = "1" Then CheckBox3.Checked = True Else CheckBox3.Checked = False
                If lb1_option.Substring(2, 1) = "1" Then CheckBox5.Checked = True Else CheckBox5.Checked = False
                If lb1_option.Substring(3, 1) = "1" Then CheckBox6.Checked = True Else CheckBox6.Checked = False
                If lb1_option.Substring(4, 1) = "1" Then CheckBox4.Checked = True Else CheckBox4.Checked = False
                If lb1_option.Substring(5, 1) = "1" Then CheckBox7.Checked = True Else CheckBox7.Checked = False
                If lb1_option.Substring(6, 1) = "1" Then CheckBox8.Checked = True Else CheckBox8.Checked = False
                If lb1_option.Substring(7, 1) = "1" Then CheckBox9.Checked = True Else CheckBox9.Checked = False
                If lb1_option.Substring(8, 1) = "1" Then CheckBox10.Checked = True Else CheckBox10.Checked = False
                TextBox17.Text = Split(lb1_option, "|=|")(1)
                ComboBox3.SelectedItem = Split(lb1_option, "|=|")(2)
                Try
                    TextBox7.Text = Split(lb1_option, "|=|")(3)
                    TextBox29.Text = Split(lb1_option, "|=|")(4)
                Catch ex As Exception

                End Try
                Dim lb1_patternArray As String() = Split(lb1_pattern, "|***|")
                DGV1.Rows.Clear()
                For i As Integer = 0 To lb1_patternArray.Length - 1
                    Dim pA As String() = Split(lb1_patternArray(i), "|=|")
                    DGV1.Rows.Add(pA)
                Next
                Button9.Text = "商品欄修正"
            Case Else
                str = ""
        End Select

        Try
            ScrollOn = False
            Dim id As String = "program" & ListBox1.SelectedIndex
            Dim he As HtmlElement = WebBrowser1.Document.GetElementById(id)
            he.ScrollIntoView(True)
            scX = WebBrowser1.Document.Body.ScrollLeft
            scY = WebBrowser1.Document.Body.ScrollTop
            'he.SetAttribute("bgcolor", "lightyellow")
        Catch ex As Exception

        End Try

        Select Case False
            Case Regex.IsMatch(ListBox1.SelectedItem, "文字素材")
                AzukiControl3.Text = str
        End Select
        GroupEnabled(ListBox1.SelectedItem)
    End Sub

    Private Sub GroupEnabled(selStr As String)
        If selStr = "全て" Then
            AzukiControl3.Text = ""
            GroupBox2.Enabled = True
            RadioButton3.Checked = True
            Button6.Text = "テキスト追加"
            GroupBox7.Enabled = True
            GroupBox1.Enabled = True
            RadioButton6.Checked = True
            Button7.Text = "バナー追加"
            GroupBox10.Enabled = True
            Button29.Text = "文字素材追加"
            AzukiControl3.IsReadOnly = True
            AzukiControl3.BackColor = Color.LightGray
            GroupBox4.Enabled = True
            Button9.Text = "商品欄追加"
            GroupBox5.Enabled = True
            Button12.Text = "タイトル追加"
        Else
            GroupBox2.Enabled = False
            GroupBox7.Enabled = True
            GroupBox1.Enabled = False
            GroupBox5.Enabled = False
            GroupBox4.Enabled = False
            GroupBox10.Enabled = False
            AzukiControl3.IsReadOnly = False
            AzukiControl3.BackColor = Color.White

            Select Case True
                Case Regex.IsMatch(selStr, "テキスト")
                    GroupBox2.Enabled = True
                    AzukiControl3.IsReadOnly = False
                    AzukiControl3.BackColor = Color.White
                Case Regex.IsMatch(selStr, "バナー")
                    GroupBox1.Enabled = True
                    AzukiControl3.IsReadOnly = True
                    AzukiControl3.BackColor = Color.LightGray
                Case Regex.IsMatch(ListBox1.SelectedItem, "タイトル")
                    GroupBox5.Enabled = True
                    AzukiControl3.IsReadOnly = True
                    AzukiControl3.BackColor = Color.LightGray
                Case Regex.IsMatch(ListBox1.SelectedItem, "文字素材")
                    GroupBox10.Enabled = True
                    AzukiControl3.IsReadOnly = False
                    AzukiControl3.BackColor = Color.White
                Case Regex.IsMatch(selStr, "商品検索")
                    GroupBox4.Enabled = True
                    AzukiControl3.IsReadOnly = True
                    AzukiControl3.BackColor = Color.LightGray
                Case Regex.IsMatch(selStr, "\<br\>|\<hr\>")
                    GroupBox7.Enabled = True
                    AzukiControl3.IsReadOnly = True
                    AzukiControl3.BackColor = Color.LightGray
            End Select
            AzukiControl3.Refresh()
        End If
    End Sub

    Private Sub HtmlUpdate()
        Dim html As String = ""

        If RadioButton8.Checked Then
            html &= "<div align='left'>" & vbCrLf
        ElseIf RadioButton7.Checked Then
            html &= "<div align='center'>" & vbCrLf
        ElseIf RadioButton9.Checked Then
            html &= "<div align='right'>" & vbCrLf
        End If

        Dim TB20Array As String() = Nothing
        Dim tb20Array_temp As String() = Split(TextBox20.Text, vbCrLf)
        For Each ta As String In tb20Array_temp
            If ta <> "" Then
                If Regex.IsMatch(ta, "^" & ComboBox3.SelectedItem) Then
                    TB20Array = Split(Split(ta, "|=|")(1), ",")
                End If
            End If
        Next

        For k As Integer = 0 To HtmlArray.Count - 1
            Dim selectArray As String() = Split(HtmlArray(k), "[|]")
            If Regex.IsMatch(selectArray(0), "(非)") Then
                Continue For
            End If

            Dim str As String = ""
            Dim lb1_pattern As String = ""
            Dim lb1_option As String = ""
            If selectArray.Length = 2 Then
                lb1_pattern = selectArray(1)
            ElseIf selectArray.Length = 3 Then
                lb1_pattern = selectArray(1)
                lb1_option = selectArray(2)
            End If

            html &= "<span id='program" & k & "'>" & vbCrLf

            Select Case True
                Case Regex.IsMatch(selectArray(0), "テキスト")
                    str = selectArray(1)
                Case Regex.IsMatch(selectArray(0), "\<br\>|\<hr\>")
                    str = selectArray(1) & vbCrLf
                Case Regex.IsMatch(selectArray(0), "バナー")
                    If InStr(selectArray(0), "バナー1") > 0 Then
                        str = File.ReadAllText(banner1_path, enc64)
                    ElseIf InStr(selectArray(0), "バナー2") > 0 Then
                        str = File.ReadAllText(banner2_path, enc64)
                    Else
                        str = File.ReadAllText(banner3_path, enc64)
                    End If
                    Dim lb1_patternArray As String() = Split(lb1_pattern, "|=|")
                    For i As Integer = 0 To lb1_patternArray.Length - 1
                        Dim pA As String() = Split(lb1_patternArray(i), "=")
                        If InStr(pA(0), "文字列") > 0 Then
                            If pA(1) <> "" Then
                                Dim tStr As String = "<font size=+2><b>" & pA(1) & "</b></font><br>" & vbCrLf
                                str = Replace(str, "[" & pA(0) & "]", tStr)
                            Else
                                str = Replace(str, "[" & pA(0) & "]" & vbCrLf, "")
                            End If
                        ElseIf InStr(pA(0), "リンク") > 0 Then
                            If pA(1) <> "" Then
                                Dim tStr As String = "<a href='" & pA(1) & "' target='_blank'>" & vbCrLf
                                str = Replace(str, "[" & pA(0) & "]", tStr)
                                str = Replace(str, "[リンクE]", "</a>" & vbCrLf)
                            Else
                                str = Replace(str, "[" & pA(0) & "]" & vbCrLf, "")
                                str = Replace(str, "[リンクE]" & vbCrLf, "")
                            End If
                        Else
                            str = Replace(str, "[" & pA(0) & "]", pA(1))
                        End If
                    Next
                Case Regex.IsMatch(selectArray(0), "タイトル")
                    Dim lb1_patternArray As String() = Split(lb1_pattern, "|=|")
                    If lb1_patternArray(0) = "囲み" Then
                        str = "<table border=1 cellspacing=0 cellpadding=5 width=$WP%>" & vbCrLf
                    Else
                        str = "<table border=0 cellspacing=0 cellpadding=5 width=$WP%>" & vbCrLf
                    End If
                    str &= "<tr>" & vbCrLf
                    Select Case lb1_patternArray(0)
                        Case "ベタ背景", "囲み"
                            str &= "<td"
                            If lb1_patternArray(1) <> "" Then str &= " bgcolor='" & lb1_patternArray(1) & "'"
                            If lb1_patternArray(4) = "センター" Then str &= " align=center"
                            If lb1_patternArray(4) = "右寄せ" Then str &= " align=right"
                            str &= "><b><font size=+3"
                            If lb1_patternArray(2) <> "" Then str &= " color='" & lb1_patternArray(2) & "'>"
                            str &= "　" & lb1_patternArray(3)
                            str &= "</font></b></td>" & vbCrLf
                        Case "アンダーライン"
                            str &= "<td"
                            If lb1_patternArray(4) = "センター" Then str &= " align=center"
                            If lb1_patternArray(4) = "右寄せ" Then str &= " align=right"
                            str &= "><b><font size=+3"
                            If lb1_patternArray(2) <> "" Then str &= " color='" & lb1_patternArray(2) & "'>"
                            str &= "　" & lb1_patternArray(3)
                            str &= "</font></b></td>" & vbCrLf
                            str &= "</tr><tr>"
                            str &= "<td bgcolor='" & lb1_patternArray(1) & "' height=1></td>" & vbCrLf
                        Case "左ライン"
                            str &= "<td bgcolor='" & lb1_patternArray(1) & "' width=1></td>" & vbCrLf
                            str &= "<td"
                            If lb1_patternArray(4) = "センター" Then str &= " align=center"
                            If lb1_patternArray(4) = "右寄せ" Then str &= " align=right"
                            str &= "><b><font size=+3"
                            If lb1_patternArray(2) <> "" Then str &= " color='" & lb1_patternArray(2) & "'>"
                            str &= "　" & lb1_patternArray(3)
                            str &= "</font></b></td>" & vbCrLf
                    End Select
                    str &= "</tr>" & vbCrLf
                    If lb1_patternArray(0) = "囲み" Then
                        str &= "</table><table border=0>" & vbCrLf
                    End If
                    str &= "<tr><td colspan=3 height=$SH></td></tr>" & vbCrLf
                    str &= "</table>" & vbCrLf
                Case Regex.IsMatch(selectArray(0), "文字素材")
                    str = File.ReadAllText(mojisozai_path, enc64)
                    Dim lb1_optionArray As String() = Split(lb1_option, "|=|")
                    Dim alignStr As String = ""
                    If lb1_optionArray(3) = "センター" Then alignStr = "center"
                    If lb1_optionArray(3) = "左寄せ" Then alignStr = "left"
                    If lb1_optionArray(3) = "右寄せ" Then alignStr = "right"
                    str = Replace(str, "$ALIGN", alignStr)
                    Dim setStr As String = ""
                    Dim fSize As String = ""
                    If lb1_optionArray(4) <> "0" Then fSize = " size=+" & lb1_optionArray(4)
                    setStr &= "<font" & fSize & " color=" & lb1_optionArray(2) & ">" & vbCrLf
                    setStr &= Replace(lb1_pattern, vbCrLf, "<br>") & vbCrLf
                    setStr &= "</font>" & vbCrLf
                    str = Replace(str, "[文字素材]", setStr)
                Case Regex.IsMatch(selectArray(0), "商品検索")
                    Dim selList As String() = Split(selectArray(0), "*")
                    Dim templeStr As String = ""
                    Dim yoko As Integer = 0
                    Dim trHeight As String = ""
                    Dim cellPad As String = ""
                    If selList(0) = "商品検索1" Then
                        templeStr = File.ReadAllText(syouhin1_path, enc64)
                        yoko = 1
                    ElseIf selList(0) = "商品検索2" Then
                        templeStr = File.ReadAllText(syouhin2_path, enc64)
                        yoko = 2
                    ElseIf selList(0) = "商品検索3" Then
                        templeStr = File.ReadAllText(syouhin3_path, enc64)
                        yoko = 3
                    ElseIf selList(0) = "商品検索1特" Then
                        templeStr = File.ReadAllText(syouhin1toku_path, enc64)
                        yoko = 1
                        trHeight = " height=40"
                        cellPad = " cellpadding=3"
                    End If

                    Dim syouhinArray As String() = Split(lb1_pattern, "|***|")
                    For p As Integer = 0 To CInt(selList(1)) - 1
                        Dim rStr As String = templeStr
                        Dim yokoNum As Integer = 1
                        Dim start As Integer = p * yoko
                        For i As Integer = start To syouhinArray.Length - 1
                            Dim sA As String() = Split(syouhinArray(i), "|=|")
                            'リンク
                            If lb1_option.Substring(3, 1) = "1" Then
                                Dim mA As String() = Split(miseArray(ComboBox3.SelectedIndex), "|=|")
                                Dim linkPath1 As String = ""
                                Dim linkPath2 As String = ""
                                If mA(1) = "楽天" Then
                                    linkPath1 = "https://item.rakuten.co.jp/" & mA(3) & "/"
                                    linkPath2 = ""
                                ElseIf mA(1) = "Yahoo" Then
                                    linkPath1 = "https://store.shopping.yahoo.co.jp/" & mA(3) & "/"
                                    linkPath2 = ""
                                ElseIf mA(1) = "Wowma" Then
                                    linkPath1 = "https://wowma.jp/item/"
                                    linkPath2 = "?l=true&e=llA&e2=listing_flpro"
                                End If
                                rStr = Replace(rStr, "[リンク" & yokoNum & "]", "<a href='" & linkPath1 & sA(0) & linkPath2 & "' target='_blank'>")
                                rStr = Replace(rStr, "[リンクE]", "</a>")
                            Else
                                rStr = Replace(rStr, "[リンク" & yokoNum & "]" & vbCrLf, "")
                                rStr = Replace(rStr, "[リンクE]", "")
                            End If
                            '画像
                            rStr = Replace(rStr, "[画像" & yokoNum & "]", sA(4))
                            '商品名
                            If lb1_option.Substring(0, 1) = "1" And sA(1) <> "" Then
                                Dim sName As String = sA(1)
                                Dim tStr As String = "<font size=+2>" & sName & "</font><br>"
                                rStr = Replace(rStr, "[商品名" & yokoNum & "]", tStr)
                            Else
                                rStr = Replace(rStr, "[商品名" & yokoNum & "]" & vbCrLf, "")
                            End If
                            '通常価・売価
                            If sA(3) <> "" Or sA(2) <> "" Then
                                Dim kakakuStr As String = ""
                                kakakuStr = "<table border=0 width=100%" & cellPad & ">" & vbCrLf
                                kakakuStr &= "<tr><td align=right>"
                                If lb1_option.Substring(4, 1) = "1" And sA(3) <> "" Then
                                    kakakuStr &= "<font size=+1><s>" & sA(3) & "</s></font>円　"
                                End If
                                If lb1_option.Substring(1, 1) = "1" And sA(2) <> "" Then
                                    kakakuStr &= "<font size=+3 color=red><b>" & sA(2) & "</b></font>円（税込）"
                                End If
                                If lb1_option.Substring(5, 1) = "1" And sA(7) <> "" Then
                                    For Each tb20 As String In TB20Array
                                        If sA(7) = Split(tb20, "=")(0) Then
                                            If Split(tb20, "=")(2) = 0 Then
                                                kakakuStr &= "　<font size=+1 color=red><b>[送料無料]</b></font>"
                                            Else
                                                kakakuStr &= "+送料 <font size=+2><b>" & Split(tb20, "=")(2) & "</b></font>円"
                                            End If
                                            Exit For
                                        End If
                                    Next
                                End If
                                kakakuStr &= "</td></tr>" & vbCrLf
                                kakakuStr &= "</table>"
                                rStr = Replace(rStr, "[通常価・売価" & yokoNum & "]", kakakuStr)
                            End If
                            'バナー
                            If (lb1_option.Substring(6, 1) = "1" And sA(7) <> "") Or (lb1_option.Substring(7, 1) = "1") Or (lb1_option.Substring(8, 1) = "1") Then
                                Dim haisouStr As String = ""
                                Dim bgColor As String = ""
                                haisouStr = "<table border=0 width=100%" & cellPad & ">" & vbCrLf
                                haisouStr &= "<tr>"
                                Dim yokoC As Integer = 0
                                If lb1_option.Substring(6, 1) = "1" Then
                                    For Each tb20 As String In TB20Array
                                        If sA(7) = Split(tb20, "=")(0) Then
                                            If InStr(Split(tb20, "=")(1), "宅配便") > 0 Then bgColor = "blue"
                                            If InStr(Split(tb20, "=")(1), "メール便") > 0 Then bgColor = "red"
                                            If InStr(Split(tb20, "=")(1), "定形外") > 0 Then bgColor = "green"
                                            haisouStr &= "<td align=center bgcolor=" & bgColor & trHeight & ">" & vbCrLf
                                            haisouStr &= "<font size=+2 color=white>" & Split(tb20, "=")(1) & "</font>" & vbCrLf
                                            haisouStr &= "</td>" & vbCrLf
                                            yokoC += 1
                                            Exit For
                                        End If
                                    Next
                                End If
                                If lb1_option.Substring(7, 1) = "1" Then
                                    haisouStr &= "<td align=center bgcolor=orange" & trHeight & ">" & vbCrLf
                                    haisouStr &= "<font size=+2 color=white>セール</font>" & vbCrLf
                                    haisouStr &= "</td>" & vbCrLf
                                    yokoC += 1
                                End If
                                If lb1_option.Substring(8, 1) = "1" Then
                                    Dim bannerStr As String = Split(lb1_option, "|=|")(4)
                                    If bannerStr <> "" Then
                                        If sA.Length > 7 Then
                                            If sA(8) <> "" Then
                                                bannerStr = sA(8)
                                            End If
                                        End If
                                        haisouStr &= "<td align=center bgcolor=navy" & trHeight & ">" & vbCrLf
                                        haisouStr &= "<font size=+2 color=white>" & bannerStr & "</font>" & vbCrLf
                                        haisouStr &= "</td>" & vbCrLf
                                        yokoC += 1
                                    Else
                                        If sA.Length > 7 Then
                                            If sA(8) <> "" Then
                                                bannerStr = sA(8)
                                                haisouStr &= "<td align=center bgcolor=navy" & trHeight & ">" & vbCrLf
                                                haisouStr &= "<font size=+2 color=white>" & bannerStr & "</font>" & vbCrLf
                                                haisouStr &= "</td>" & vbCrLf
                                                yokoC += 1
                                            End If
                                        End If
                                    End If
                                End If
                                haisouStr &= "</tr>" & vbCrLf
                                If selList(0) = "商品検索1特" Then
                                    haisouStr &= "<tr><td colspan=" & yokoC & " height=5></td></tr>" & vbCrLf
                                End If
                                haisouStr &= "</table>" & vbCrLf
                                rStr = Replace(rStr, "[バナー" & yokoNum & "]", haisouStr)
                            Else
                                rStr = Replace(rStr, "[バナー" & yokoNum & "]" & vbCrLf, "")
                            End If
                            '紹介文
                            If lb1_option.Substring(2, 1) = "1" And sA(5) <> "" Then
                                Dim tStr As String = "<font size=+2>" & sA(5) & "</font><br>"
                                rStr = Replace(rStr, "[紹介文" & yokoNum & "]", tStr)
                            Else
                                rStr = Replace(rStr, "[紹介文" & yokoNum & "]" & vbCrLf, "")
                            End If
                            If yokoNum = selList(1) + 1 Then
                                Exit For
                            End If
                            yokoNum += 1
                        Next
                        str &= rStr & vbCrLf & vbCrLf
                    Next

                    ''文字列修正
                    ''A|,|B|、|C|,|D
                    'Dim repStr As String() = Split(lb1_option, "|=|")
                    'If str <> "" And repStr(1) <> "" Then
                    '    Dim rSArray As String() = Split(repStr(1), "|、|")
                    '    For r As Integer = 0 To rSArray.Length - 1
                    '        Dim rS As String() = Split(rSArray(r), "|,|")
                    '        str = Regex.Replace(str, rS(0), rS(1))
                    '    Next
                    'End If
                Case Else
                    str = ""
            End Select

            '全体設定
            If RadioButton8.Checked Then str = Replace(str, "$ALIGN", "left")
            If RadioButton7.Checked Then str = Replace(str, "$ALIGN", "center")
            If RadioButton9.Checked Then str = Replace(str, "$ALIGN", "right")
            str = Replace(str, "$WP", NumericUpDown3.Value)
            str = Replace(str, "$SW", NumericUpDown2.Value)
            str = Replace(str, "$SH", NumericUpDown4.Value)

            html &= str & vbCrLf
            html &= "</span>" & vbCrLf
        Next

        html &= "</div>" & vbCrLf

        WebBrowser1.DocumentText = html
    End Sub

    '----------------------------
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        If ListBox1.SelectedItem = "新規追加" Then
            Exit Sub
        End If

        Dim DR As DialogResult = MsgBox(ListBox1.SelectedItem & " を削除します", MsgBoxStyle.OkCancel)
        If DR = DialogResult.OK Then
            HtmlArray.RemoveAt(ListBox1.SelectedIndex)
            ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
            LB1update()
        Else
            Exit Sub
        End If
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        If ListBox1.SelectedItem = "新規追加" Then
            Exit Sub
        End If

        Dim selectArray As String() = Split(HtmlArray(ListBox1.SelectedIndex), "[|]")
        Dim hNew As String = ""
        For i As Integer = 0 To selectArray.Length - 1
            If i = 0 Then
                If InStr(selectArray(i), "(非)") > 0 Then
                    hNew &= Replace(selectArray(i), "(非)", "")
                Else
                    hNew &= "(非)" & selectArray(i)
                End If
            Else
                hNew &= "[|]" & selectArray(i)
            End If
        Next
        HtmlArray(ListBox1.SelectedIndex) = hNew
        LB1update()
    End Sub

    '----------------------------
    Private DragStartIdx As Integer = -1
    Private Sub ListBox1_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListBox1.MouseDown
        If e.Button = System.Windows.Forms.MouseButtons.Left Then
            ' マウスの左ボタンを押した時にListBoxの項目番号を保存
            DragStartIdx = sender.IndexFromPoint(New Point(e.X, e.Y))
        Else
            ' それ以外の時は何もしない
            DragStartIdx = -1
        End If
    End Sub

    Private Sub ListBox1_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListBox1.MouseUp
        ' マウスのボタンを上げた時は初期状態へ
        DragStartIdx = -1
        HtmlUpdate()
    End Sub

    Dim LB1Move As Boolean = True
    Private Sub ListBox1_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListBox1.MouseMove
        If LB1Move Then
            LB1Move = False
            If e.Button = System.Windows.Forms.MouseButtons.Left And DragStartIdx >= 0 Then
                'If (Control.ModifierKeys And Keys.Control) = Keys.Control Then
                '    ' Ctrlキーを押している時はドラッグ＆ドロップ
                '    ListBox1.DoDragDrop(ListBox1.Items(DragStartIdx).ToString(), DragDropEffects.All)
                'Else
                If ListBox1.SelectedItem = "新規追加" Then
                    Exit Sub
                End If

                ' Ctrlキーを押していない時はListBox内のドラッグ
                Dim DragEndIdx As Integer = ListBox1.IndexFromPoint(New Point(e.X, e.Y))

                If DragStartIdx >= 0 And DragEndIdx >= 0 And
                DragStartIdx <> DragEndIdx Then
                    If DragStartIdx < DragEndIdx Then
                        ListBox1.Items.Insert(DragEndIdx + 1, ListBox1.Items(DragStartIdx).ToString())
                        HtmlArray.Insert(DragEndIdx + 1, HtmlArray(DragStartIdx))
                        ListBox1.Items.RemoveAt(DragStartIdx)
                        HtmlArray.RemoveAt(DragStartIdx)
                        ListBox1.SelectedIndex = DragEndIdx
                    ElseIf DragStartIdx > DragEndIdx Then
                        ListBox1.Items.Insert(DragEndIdx, ListBox1.Items(DragStartIdx).ToString())
                        HtmlArray.Insert(DragEndIdx, HtmlArray(DragStartIdx))
                        ListBox1.Items.RemoveAt(DragStartIdx + 1)
                        HtmlArray.RemoveAt(DragStartIdx + 1)
                        ListBox1.SelectedIndex = DragEndIdx
                    End If
                    DragStartIdx = DragEndIdx
                End If
                'End If
            End If
            LB1Move = True
        End If
    End Sub

    Private Sub ListBox1_DragEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles ListBox1.DragEnter
        ' ドラッグ中にListBox内に入って来た時はドロップ許可
        e.Effect = DragDropEffects.Move
    End Sub

    Private Sub ListBox1_DragDrop(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles ListBox1.DragDrop
        ' ドラッグ中にマウスのボタンを上げた時はドロップ(最終行への追加)
        'sender.Items.Add(e.Data.GetData(DataFormats.Text, True))
    End Sub

    '----------------------------
    Private fromIndex As Integer
    Private dragIndex As Integer
    Private dragRect As Rectangle
    Private selectColumnMae As Integer = 0

    Private Sub DataGridView1_DragDrop(ByVal sender As Object, ByVal e As DragEventArgs) Handles DGV1.DragDrop
        Dim p As Point = DGV1.PointToClient(New Point(e.X, e.Y))
        dragIndex = DGV1.HitTest(p.X, p.Y).RowIndex
        If (e.Effect = DragDropEffects.Move) Then
            Dim dragRow As DataGridViewRow = e.Data.GetData(GetType(DataGridViewRow))
            DGV1.Rows.RemoveAt(fromIndex)
            DGV1.Rows.Insert(dragIndex, dragRow)
        End If
        DGV1.CurrentCell = DGV1(selectColumnMae, dragIndex)
    End Sub

    Private Sub DataGridView1_DragOver(ByVal sender As Object, ByVal e As DragEventArgs) Handles DGV1.DragOver
        e.Effect = DragDropEffects.Move
    End Sub

    Private Sub DataGridView1_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles DGV1.MouseDown
        selectColumnMae = DGV1.HitTest(e.X, e.Y).ColumnIndex
        If selectColumnMae < 0 Then
            selectColumnMae = 0
        End If
        fromIndex = DGV1.HitTest(e.X, e.Y).RowIndex
        If fromIndex > -1 Then
            Dim dragSize As Size = SystemInformation.DragSize
            dragRect = New Rectangle(New Point(e.X - (dragSize.Width / 2), e.Y - (dragSize.Height / 2)), dragSize)
        Else
            dragRect = Rectangle.Empty
        End If
    End Sub

    Private Sub DataGridView1_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles DGV1.MouseMove
        If (e.Button And MouseButtons.Left) = MouseButtons.Left Then
            If (dragRect <> Rectangle.Empty _
            AndAlso Not dragRect.Contains(e.X, e.Y)) Then
                DGV1.DoDragDrop(DGV1.Rows(fromIndex), DragDropEffects.Move)
            End If
        End If
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim sPath As String = ""
        Dim sName As String = ""
        If InStr(Me.Text, "メルマガ作成") > 0 Then
            sPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        Else
            sPath = Path.GetDirectoryName(Me.Text)
            sName = Path.GetFileNameWithoutExtension(Me.Text)
        End If
        Dim savePath As String = ""
        '.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
        Dim sfd As New SaveFileDialog With {
                .Filter = "txtファイル(*.txt)|*.txt|すべてのファイル(*.*)|*.*",
                .AutoUpgradeEnabled = False,
                .FilterIndex = 0,
                .Title = "保存先のファイルを選択してください",
                .InitialDirectory = sPath,
                .RestoreDirectory = True,
                .OverwritePrompt = True,
                .CheckPathExists = True,
                .FileName = sName
            }
        If sfd.ShowDialog(Me) = DialogResult.OK Then
            savePath = sfd.FileName
        Else
            Exit Sub
        End If

        Dim str As String = ""
        If RadioButton8.Checked Then str &= "left"
        If RadioButton7.Checked Then str &= "center"
        If RadioButton9.Checked Then str &= "right"
        str &= "," & NumericUpDown3.Value
        str &= "," & NumericUpDown2.Value
        str &= "," & NumericUpDown4.Value
        str &= "[txt改行]"
        For i As Integer = 0 To HtmlArray.Count - 1
            str &= HtmlArray(i) & "[txt改行]"
        Next

        File.WriteAllText(savePath, str, encSJ)
        MsgBox("保存しました")
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Dim ofd As New OpenFileDialog With {
            .InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
            .Filter = "txtファイル(*.txt)|*.txt|すべてのファイル(*.*)|*.*",
            .FilterIndex = 1,
            .Title = "開くファイルを選択してください",
            .RestoreDirectory = True,
            .CheckFileExists = True,
            .CheckPathExists = True
        }

        If ofd.ShowDialog() = DialogResult.OK Then
            TemplateRead(ofd.FileName)
        End If
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        If ComboBox5.SelectedIndex = 0 Then
            Exit Sub
        End If

        Dim templatePath As String = Path.GetDirectoryName(Form1.appPath) & "\htmltemplate\" & ComboBox5.SelectedItem & ".txt"
        Dim pattern As String = Regex.Replace(ComboBox5.SelectedItem, "テンプレ_|メルマガ", "")
        For i As Integer = 0 To ComboBox3.Items.Count - 1
            If InStr(ComboBox3.Items(i).ToString, pattern) > 0 Then
                ComboBox3.SelectedIndex = i
                Exit For
            End If
        Next
        TemplateRead(templatePath)
    End Sub

    Private Sub TemplateRead(templatePath)
        Me.Text = templatePath
        Dim str As String = File.ReadAllText(templatePath, encSJ)
        If InStr(str, "[txt改行]") = 0 Then
            MsgBox("テンプレートファイルではないようです。")
            Exit Sub
        End If
        Dim strArray As String() = Split(str, "[txt改行]")
        Dim aLine As String() = Split(strArray(0), ",")
        If aLine(0) = "left" Then RadioButton8.Checked = True
        If aLine(0) = "center" Then RadioButton7.Checked = True
        If aLine(0) = "right" Then RadioButton9.Checked = True
        NumericUpDown3.Value = aLine(1)
        NumericUpDown2.Value = aLine(2)
        NumericUpDown4.Value = aLine(3)
        HtmlArray.Clear()
        For i As Integer = 1 To strArray.Length - 1
            If strArray(i) <> "" Then
                HtmlArray.Add(strArray(i))
            End If
        Next
        LB1update()
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        Dim savePath As String = ""
        '.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
        Dim sfd As New SaveFileDialog With {
                .Filter = "txtファイル(*.txt)|*.txt|すべてのファイル(*.*)|*.*",
                .AutoUpgradeEnabled = False,
                .FilterIndex = 0,
                .Title = "保存先のファイルを選択してください",
                .InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                .RestoreDirectory = False,
                .OverwritePrompt = True,
                .CheckPathExists = True
            }
        If sfd.ShowDialog(Me) = DialogResult.OK Then
            savePath = sfd.FileName
        Else
            Exit Sub
        End If

        Dim html As String = WebBrowser1.DocumentText
        html = Regex.Replace(html, "\<[\/]{0,1}span.*?\>[\n|\r]", "")
        File.WriteAllText(savePath, html, encSJ)
        MsgBox("保存しました")
    End Sub

    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click
        Dim savePath As String = ""
        '.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
        Dim sfd As New SaveFileDialog With {
                .Filter = "txtファイル(*.txt)|*.txt|すべてのファイル(*.*)|*.*",
                .AutoUpgradeEnabled = False,
                .FilterIndex = 0,
                .Title = "保存先のファイルを選択してください",
                .InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                .RestoreDirectory = False,
                .OverwritePrompt = True,
                .CheckPathExists = True
            }
        If sfd.ShowDialog(Me) = DialogResult.OK Then
            savePath = sfd.FileName
        Else
            Exit Sub
        End If

        '------------------------------------------------------------------------
        '商品検索2*5[|]
        'code|=|name|=|baika|=|tujouka|=|img|=|setumei|=|zaiko|=|souko|***|...[|]
        'checkbox|=|regexA|,|regexB|、|...|=|ComboBox3|=|TextBox7|=|TextBox29
        '------------------------------------------------------------------------
        Dim str As String = ""
        For i As Integer = 0 To HtmlArray.Count - 1
            Dim list1 As String() = Split(HtmlArray(i), "[|]")
            If InStr(list1(0), "(非)") > 0 Then
                Continue For
            End If

            Dim res As String = ""
            Select Case True
                Case Regex.IsMatch(list1(0), "テキスト")
                    Dim list2 As String = list1(1)
                    res = Regex.Replace(list2, "<.*?>", "")
                Case Regex.IsMatch(list1(0), "商品検索")
                    Dim list2 As String() = Split(list1(1), "|***|")
                    Dim list3 As String = Split(list1(2), "|=|")(2)
                    Dim mA As String() = Split(miseArray(ComboBox3.Items.IndexOf(list3)), "|=|")
                    Dim linkPath1 As String = ""
                    Dim linkPath2 As String = ""
                    If mA(1) = "楽天" Then
                        linkPath1 = "https://item.rakuten.co.jp/" & mA(3) & "/"
                        linkPath2 = ""
                    ElseIf mA(1) = "Yahoo" Then
                        linkPath1 = "https://store.shopping.yahoo.co.jp/" & mA(3) & "/"
                        linkPath2 = ""
                    ElseIf mA(1) = "Wowma" Then
                        linkPath1 = "https://wowma.jp/item/"
                        linkPath2 = "?l=true&e=llA&e2=listing_flpro"
                    End If
                    For k As Integer = 0 To list2.Length - 1
                        Dim list2_2 As String() = Split(list2(k), "|=|")
                        res &= list2_2(0) & " " & list2_2(1) & vbCrLf
                        res &= linkPath1 & list2_2(0) & linkPath2 & vbCrLf
                    Next
                Case Regex.IsMatch(list1(0), "バナー")
                    Dim list2 As String() = Split(list1(1), "|=|")
                    Dim list3 As String() = Split(list2(2), "=")
                    If list3(1) <> "" Then
                        res &= list3(1) & vbCrLf
                    End If
                    Dim list4 As String() = Split(list2(1), "=")
                    If list4(1) <> "" Then
                        res &= list4(1) & vbCrLf
                    End If
                Case Regex.IsMatch(list1(0), "タイトル")
                    Dim list2 As String() = Split(list1(1), "|=|")
                    res &= vbCrLf & list2(3) & vbCrLf
            End Select

            str &= res & vbCrLf
        Next

        If str <> "" Then
            str = Regex.Replace(str, "[\r|\n]{5,99}", vbCrLf)
        End If

        File.WriteAllText(savePath, str, encSJ)
        MsgBox("保存しました")
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        Dim frm As Form = New HtmlDialog_F_TableCreate
        Me.Dispose()
        frm.Show()
    End Sub

    Private Sub 行削除ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 行削除ToolStripMenuItem.Click
        DGV1.Rows.RemoveAt(DGV1.SelectedCells(0).RowIndex)
    End Sub

    Private Sub ToolStripStatusLabel3_Click(sender As Object, e As EventArgs) Handles ToolStripStatusLabel3.Click
        Mall_main.Show()
        Application.DoEvents()
        Mall_main.CheckBox17.Checked = True
        Mall_main.CheckBox19.Checked = True
        Mall_main.CheckBox18.Checked = False
        Mall_main.Button23.PerformClick()
    End Sub

End Class