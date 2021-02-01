Imports System.Net
Imports System.IO
Imports EncodingOperation
Imports System.Text.RegularExpressions

Public Class HTMLdialog
    Dim secValue As String = ""
    'Dim tenpo As String() = New String() {"FKstyle", "Lucky9", "あかねYahoo", "暁", "アリス", "あかね楽天", "KuraNavi", "雑貨倉庫", "海東KT", "問よか", "Amazon(5%)", "Amazon(8%)", "Amazon(10%)"}
    Dim tenpo As String() = New String() {"FKstyle(以前)", "FKstyle", "Lucky9", "海東KT", "暁", "アリス", "あかね楽天", "KuraNavi", "雑貨倉庫", "あかねYahoo", "問よか", "Amazon(105%)", "Amazon(108%)", "Amazon(110%)"}
    Dim donya_vip As String() = New String() {"ブロンズ", "シルバー", "ゴールド", "プラチナ", "ダイヤ"}

    Private Sub HTMLdialog_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        secValue = "html_dialog"   'iniファイルセクション

        ServerListComb()

        '店舗選択ラジオボタン
        If CStr(Form1.inif.ReadINI(secValue, "tenpoCheck1", "")) <> "" Then
            RadioButton1.Checked = Form1.inif.ReadINI(secValue, "tenpoCheck1", "")
        End If
        If CStr(Form1.inif.ReadINI(secValue, "tenpoCheck2", "")) <> "" Then
            RadioButton2.Checked = Form1.inif.ReadINI(secValue, "tenpoCheck2", "")
        End If
        If CStr(Form1.inif.ReadINI(secValue, "tenpoCheck3", "")) <> "" Then
            RadioButton3.Checked = Form1.inif.ReadINI(secValue, "tenpoCheck3", "")
        End If
        If CStr(Form1.inif.ReadINI(secValue, "tenpoCheck4", "")) <> "" Then
            RadioButton4.Checked = Form1.inif.ReadINI(secValue, "tenpoCheck4", "")
        End If
        If CStr(Form1.inif.ReadINI(secValue, "tenpoCheck8", "")) <> "" Then
            RadioButton8.Checked = Form1.inif.ReadINI(secValue, "tenpoCheck8", "")
        End If
        '----------------------------------------

        For i As Integer = 0 To tenpo.Length - 1
            DataGridView1.Rows.Add(tenpo(i))
            ComboBox3.Items.Add(tenpo(i))
        Next
        ComboBox3.SelectedItem = "海東KT"
        TextBox38.Text = 0
        ComboBox11.SelectedIndex = 0

        ComboBox1.SelectedIndex = 0

        For i As Integer = 0 To donya_vip.Length - 1
            DataGridView2.Rows.Add(donya_vip(i))
        Next

        TextBox43.Text = "ブロンズ ＝ 定価" & vbCrLf & "シルバー ＝ 販売価格" & vbCrLf & "ゴールド、プラチナ、ダイヤ ＝ 販売価格 * 比率"
        TextBox44.Text = "価格 ＝ 売価 * 比率" & vbCrLf & "利益 ＝ 価格 - 原価" & vbCrLf & "利益率 ＝ 利益 / 原価"
    End Sub

    Private Sub HTMLdialog_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        ServerClose()
    End Sub

    Private Sub ServerListComb()
        Dim di2 As New System.IO.DirectoryInfo(Path.GetDirectoryName(Form1.appPath) & "\ftplist")
        Dim files2 As System.IO.FileInfo() = di2.GetFiles("*.txt", System.IO.SearchOption.AllDirectories)
        ComboBox4.Items.Clear()
        ComboBox5.Items.Clear()
        For Each f As System.IO.FileInfo In files2
            Dim fname As String = System.IO.Path.GetFileNameWithoutExtension(f.Name)
            ComboBox4.Items.Add(fname)
            ComboBox5.Items.Add(fname)
        Next
        ComboBox4.SelectedIndex = 0
        ComboBox5.SelectedIndex = 0
    End Sub

    Private Sub AzukiControl6_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles AzukiControl6.DragDrop
        If AzukiControl6.Text <> "" Then
            Dim DR As DialogResult = MsgBox("読み込んでよろしいですか？", MsgBoxStyle.YesNo Or MsgBoxStyle.SystemModal)
            If DR = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If
        End If

        For Each filename As String In e.Data.GetData(DataFormats.FileDrop)
            Dim e1 As EncodingJP
            Dim e2 As Boolean
            AzukiControl6.Text = TextFileEncoding.LoadFromFile(filename, e1, True, e2)
            ToolStripStatusLabel4.Text = EncodingName(e1)
            ToolStripStatusLabel5.Text = filename
            Exit For
        Next
    End Sub

    Private Sub AzukiControl6_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles AzukiControl6.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub AzukiControl7_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles AzukiControl7.DragDrop
        If AzukiControl7.Text <> "" Then
            Dim DR As DialogResult = MsgBox("読み込んでよろしいですか？", MsgBoxStyle.YesNo Or MsgBoxStyle.SystemModal)
            If DR = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If
        End If

        For Each filename As String In e.Data.GetData(DataFormats.FileDrop)
            Dim e1 As EncodingJP
            Dim e2 As Boolean
            AzukiControl7.Text = TextFileEncoding.LoadFromFile(filename, e1, True, e2)
            ToolStripStatusLabel6.Text = EncodingName(e1)
            ToolStripStatusLabel7.Text = filename
            Exit For
        Next
    End Sub

    Private Sub AzukiControl7_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles AzukiControl7.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        HTMLcreate.Button1_Click(sender, e)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Hide()
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

    '価格変更データ取得
    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Dim allArray As String() = Split(AzukiControl6.Text, vbCrLf)
        AzukiControl6.Text = ""

        ToolStripProgressBar1.Maximum = allArray.Length
        ToolStripProgressBar1.Value = 0

        Dim chengeNum As String = ""
        For i As Integer = 0 To allArray.Length - 1
            If allArray(i) <> "" Then
                Dim lineArray As String() = Split(allArray(i), ",")
                Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("EUC-JP")

                Dim url As String = lineArray(3)
                Dim html As String = ""
                Try
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

                '価格を取得
                Dim selecter As String = "//span[@class='price2']"
                Dim resStr As String = HTMLGet(doc, selecter, 0)
                resStr = resStr.Replace(Chr(13), "").Replace(Chr(10), "")
                resStr = Replace(resStr, "円", "")
                resStr = Replace(resStr, ",", "")

                '代入
                For k As Integer = 0 To lineArray.Length - 1
                    If k = lineArray.Length - 1 Then
                        AzukiControl6.Text &= lineArray(k) & vbCrLf
                    ElseIf k = 2 Then
                        If lineArray(k) <> resStr Then
                            chengeNum &= lineArray(0) & " " & lineArray(k) & "→" & resStr & vbCrLf
                        End If
                        AzukiControl6.Text &= resStr & ","
                    Else
                        AzukiControl6.Text &= lineArray(k) & ","
                    End If
                Next

                AzukiControl6.Document.SetCaretIndex(AzukiControl6.LineCount - 1, 0)
                AzukiControl6.ScrollToCaret()
                Application.DoEvents()
                ToolStripProgressBar1.Value += 1
                'System.Threading.Thread.Sleep(200)
            End If
        Next

        ToolStripProgressBar1.Value = 0
        MsgBox(chengeNum & "修正しました。", MsgBoxStyle.SystemModal)
    End Sub


    Private Sub 上書き保存ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 上書き保存ToolStripMenuItem.Click
        If AzukiControl6.Text <> "" Then
            Dim FileName As String = ToolStripStatusLabel5.Text
            SaveCsv(FileName, ToolStripStatusLabel4.Text)
            'MsgBox("上書き保存しました", MsgBoxStyle.SystemModal)
        Else
            SaveCsv("", ToolStripStatusLabel4.Text)
        End If
    End Sub

    Private Sub 名前を付けて保存ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 名前を付けて保存ToolStripMenuItem.Click
        Dim sfd As New SaveFileDialog With {
                    .AutoUpgradeEnabled = False,
            .FileName = "itemlist.csv"
        }
        Dim folder As String = Path.GetDirectoryName(Form1.appPath) & "\formchange"
        sfd.InitialDirectory = folder
        sfd.Filter = "テンプレートファイル(*.csv)|*.csv"
        sfd.FilterIndex = 0
        sfd.Title = "保存先のファイルを選択してください"
        sfd.RestoreDirectory = True
        sfd.OverwritePrompt = True
        sfd.CheckPathExists = True

        'ダイアログを表示する
        If sfd.ShowDialog(Me) = DialogResult.OK Then
            SaveCsv(sfd.FileName, ToolStripStatusLabel4.Text)
            'MsgBox(sfd.FileName & vbCrLf & "保存しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    Public Sub SaveCsv(ByVal fp As String, ByVal enc As String)
        If fp = "" Then
            Dim sfd As New SaveFileDialog()
            If ToolStripStatusLabel5.Text = "..." Then
                sfd.FileName = "新しいファイル.html"
            Else
                sfd.FileName = Path.GetFileName(ToolStripStatusLabel5.Text)
            End If
            sfd.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
            sfd.Filter = "HTMLファイル(*.html)|*.html|TXTファイル(*.txt)|*.txt|すべてのファイル(*.*)|*.*"
            sfd.AutoUpgradeEnabled = False
            sfd.FilterIndex = 0
            sfd.Title = "保存先のファイルを選択してください"
            sfd.RestoreDirectory = True
            sfd.OverwritePrompt = True
            sfd.CheckPathExists = True

            'ダイアログを表示する
            If sfd.ShowDialog(Me) = DialogResult.OK Then
                'SaveCsv(sfd.FileName, ToolStripStatusLabel5.Text)
                ' CSVファイルオープン
                Dim en As EncodingJP = DecodingName(enc)
                TextFileEncoding.SaveToFile(sfd.FileName, AzukiControl6.Text, en)
                'Dim sw As IO.StreamWriter = New IO.StreamWriter(sfd.FileName, False, System.Text.Encoding.GetEncoding(enc))
                'sw.Write(az.Text)
                'sw.Close()

                ToolStripStatusLabel5.Text = sfd.FileName
                MsgBox(sfd.FileName & vbCrLf & "保存しました", MsgBoxStyle.SystemModal)
            End If
        Else
            ' CSVファイルオープン
            Dim en As EncodingJP = DecodingName(enc)
            TextFileEncoding.SaveToFile(fp, AzukiControl6.Text, en)
            'Dim sw As IO.StreamWriter = New IO.StreamWriter(fp, False, System.Text.Encoding.GetEncoding(enc))
            'sw.Write(az.Text)
            'sw.Close()
            MsgBox("上書き保存しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    Private Function EncodingName(ByVal e As EncodingJP) As String
        Select Case e
            Case EncodingJP.ASCII
                Return "欧文"
            Case EncodingJP.SJIS
                Return "SJIS"
            Case EncodingJP.EUC
                Return "EUC"
            Case EncodingJP.JIS
                Return "JIS"
            Case EncodingJP.UTF16_LE
                Return "Unicode"
            Case EncodingJP.UTF16_BE
                Return "Unicode(Big Endian)"
            Case EncodingJP.UTF8
                Return "UTF-8"
            Case EncodingJP.UTF8_BOM
                Return "UTF-8(BOM)"
            Case EncodingJP.UTF7
                Return "UTF-7"
            Case Else
                Return ""
        End Select
    End Function

    Private Function DecodingName(ByVal e As String) As EncodingJP
        Select Case e
            Case "欧文"
                Return EncodingJP.ASCII
            Case "SJIS"
                Return EncodingJP.SJIS
            Case "EUC"
                Return EncodingJP.EUC
            Case "JIS"
                Return EncodingJP.JIS
            Case "Unicode"
                Return EncodingJP.UTF16_LE
            Case "Unicode(Big Endian)"
                Return EncodingJP.UTF16_BE
            Case "UTF-8"
                Return EncodingJP.UTF8
            Case "UTF-8(BOM)"
                Return EncodingJP.UTF8_BOM
            Case "UTF-7"
                Return EncodingJP.UTF7
            Case Else
                Return EncodingJP.NONE
        End Select
    End Function

    '新規登録
    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        If RadioButton8.Checked Then
            MsgBox("Yahooは対応していません")
            Exit Sub
        End If

        Dim allArray As String() = Split(AzukiControl6.Text, vbCrLf)
        For i As Integer = 0 To allArray.Length - 1
            Dim lineArray As String() = Split(allArray(i), ",")
            If lineArray(0) = TextBox5.Text Then
                MsgBox("既に登録したデータがあります", MsgBoxStyle.SystemModal)
                Exit Sub
            End If
        Next

        Dim allText As String = AzukiControl6.Text

        Try
            If RadioButton1.Checked = True Or RadioButton2.Checked = True Or RadioButton4.Checked = True Then
                Dim url As String = ""
                Dim cateUrl As String = ""
                If RadioButton1.Checked = True Then
                    url = "https://item.rakuten.co.jp/patri/" & TextBox5.Text
                    cateUrl = "https:item.rakuten.co.jppatric"
                ElseIf RadioButton4.Checked = True Then
                    url = "https://item.rakuten.co.jp/ashop/" & TextBox5.Text
                    cateUrl = "https:item.rakuten.co.jpashopc"
                Else
                    url = "https://item.rakuten.co.jp/alice-zk/" & TextBox5.Text
                    cateUrl = "https:item.rakuten.co.jpalice-zkc"
                End If

                Dim str As String = AzukiControl6.Text
                If CheckBox17.Checked = True Then
                    AzukiControl6.Text = ""
                End If

                Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("EUC-JP")

                Dim html As String = ""
                Dim req As WebRequest = WebRequest.Create(url)
                req.UseDefaultCredentials = False
                Dim res As WebResponse = req.GetResponse()
                Dim st As Stream = res.GetResponseStream()
                Dim sr As StreamReader = New StreamReader(st, enc)
                html = sr.ReadToEnd()
                sr.Close()
                st.Close()

                Dim doc As HtmlAgilityPack.HtmlDocument = New HtmlAgilityPack.HtmlDocument
                doc.LoadHtml(html)

                Dim strArray As String() = Nothing

                AzukiControl6.Text &= TextBox5.Text & ","

                'キャッチコピーを取得
                Dim selecter As String = "//span[@class='catch_copy']"
                Dim resStr As String = HTMLGet(doc, selecter, 0)
                resStr = Regex.Replace(resStr, "【.*】", "")
                resStr = Regex.Replace(resStr, "★.*★", "")
                AzukiControl6.Text &= resStr & ","

                '価格を取得
                selecter = "//span[@class='price2']"
                resStr = HTMLGet(doc, selecter, 0)
                resStr = resStr.Replace(Chr(13), "").Replace(Chr(10), "")
                resStr = Replace(resStr, "円", "")
                resStr = Replace(resStr, ",", "")
                AzukiControl6.Text &= resStr & ","

                AzukiControl6.Text &= url & ","

                '画像を取得
                selecter = "//span[@class='sale_desc']"
                resStr = HTMLGet(doc, selecter, 1)
                Dim doc2 As HtmlAgilityPack.HtmlDocument = New HtmlAgilityPack.HtmlDocument
                doc2.LoadHtml(resStr)
                selecter = "//img[1]"
                resStr = HTMLGet(doc2, selecter, 2)
                AzukiControl6.Text &= resStr & ","

                '登録日を入力
                If CheckBox17.Checked = True Then
                    AzukiControl6.Text &= Format(Now, "yyyyMMdd") & ","
                Else
                    AzukiControl6.Text &= "0,"
                End If

                'カテゴリを取得
                selecter = "//td[@class='sdtext']/a"
                Dim node As HtmlAgilityPack.HtmlNodeCollection = HTMLGet(doc, selecter, 3)
                Dim cateArray As ArrayList = New ArrayList
                For i As Integer = 0 To node.Count - 1
                    Dim oneurl As String = node(i).Attributes("href").Value.Trim()
                    oneurl = Replace(oneurl, "/", "")
                    If oneurl <> cateUrl Then
                        oneurl = Replace(oneurl, cateUrl, "")
                        If Not cateArray.Contains(oneurl) Then
                            cateArray.Add(oneurl)
                        End If
                    End If
                Next
                resStr = ""
                For i As Integer = 0 To cateArray.Count - 1
                    If i = 0 Then
                        resStr = cateArray(i)
                    Else
                        resStr &= "|" & cateArray(i)
                    End If
                Next
                AzukiControl6.Text &= resStr & ",0" & vbCrLf

                If CheckBox17.Checked = True Then
                    AzukiControl6.Text &= str
                    AzukiControl6.Document.SetCaretIndex(0, 0)
                    AzukiControl6.ScrollToCaret()
                Else
                    AzukiControl6.Document.SetCaretIndex(AzukiControl6.LineCount - 1, 0)
                    AzukiControl6.ScrollToCaret()
                End If
            ElseIf RadioButton3.Checked = True Then
                Dim url As String = "https://wowma.jp/bep/m/klist2?e_scope=O&user=38238702&keyword=[code]&submit=&clow=&chigh="
                url = Replace(url, "[code]", TextBox5.Text)

                Dim str As String = AzukiControl6.Text
                If CheckBox17.Checked = True Then
                    AzukiControl6.Text = ""
                End If

                WebBrowser1.Navigate(url)
                Do
                    Application.DoEvents()
                Loop Until WebBrowser1.ReadyState = WebBrowserReadyState.Complete And WebBrowser1.IsBusy = False

                '文字化け対策
                Dim buf As Byte() = New Byte(WebBrowser1.DocumentStream.Length) {}
                WebBrowser1.DocumentStream.Read(buf, 0, CInt(WebBrowser1.DocumentStream.Length))
                Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS")
                Dim s As String = enc.GetString(buf)

                Dim doc As HtmlAgilityPack.HtmlDocument = New HtmlAgilityPack.HtmlDocument
                doc.LoadHtml(s)

                Dim strArray As String() = Nothing

                '商品コード
                AzukiControl6.Text &= TextBox5.Text & ","

                '商品名を取得
                Dim selecter As String = "//span[@class='name']"
                Dim resStr As String = HTMLGet(doc, selecter, 0)
                resStr = resStr.Substring(0, 15)
                AzukiControl6.Text &= resStr & ","

                '価格を取得
                selecter = "//span[@class='price']"
                resStr = HTMLGet(doc, selecter, 0)
                resStr = Replace(resStr, vbLf, "")
                resStr = Replace(resStr, vbTab, "")
                resStr = Replace(resStr, ",", "")
                Dim array As String() = Split(resStr, "円")
                resStr = array(0)
                AzukiControl6.Text &= resStr & ","

                'URL
                selecter = "//a[@class='modImageColL']"
                resStr = HTMLGet(doc, selecter, 4)
                AzukiControl6.Text &= "https://wowma.jp" & resStr & ","

                '画像を取得
                selecter = "//div[@class='image']/p[@class='inner']/img[1]"
                resStr = HTMLGet(doc, selecter, 2)
                AzukiControl6.Text &= resStr & ","

                '登録日を入力
                If CheckBox17.Checked = True Then
                    AzukiControl6.Text &= Format(Now, "yyyyMMdd") & ","
                Else
                    AzukiControl6.Text &= "0,"
                End If

                AzukiControl6.Text &= vbCrLf

                If CheckBox17.Checked = True Then
                    AzukiControl6.Text &= str
                    AzukiControl6.Document.SetCaretIndex(0, 0)
                    AzukiControl6.ScrollToCaret()
                Else
                    AzukiControl6.Document.SetCaretIndex(AzukiControl6.LineCount - 1, 0)
                    AzukiControl6.ScrollToCaret()
                End If
            End If

            MsgBox("取得しました", MsgBoxStyle.SystemModal)
        Catch ex As Exception
            AzukiControl6.Text = allText
            MsgBox("商品が見つかりません", MsgBoxStyle.SystemModal)
        End Try
    End Sub


    'csvで価格更新
    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Dim ofd As New OpenFileDialog With {
            .FileName = "default.html",
            .InitialDirectory = Environment.SpecialFolder.DesktopDirectory,
            .Filter = "CSVファイル(*.csv)|*.csv|すべてのファイル(*.*)|*.*",
            .FilterIndex = 2,
            .Title = "開くファイルを選択してください",
            .RestoreDirectory = True,
            .CheckFileExists = True,
            .CheckPathExists = True
        }

        Dim changeNum As Integer = 0
        Dim changeStr As String = ""

        'ダイアログを表示する
        If ofd.ShowDialog() = DialogResult.OK Then
            Dim fName As String = ofd.FileName
            Dim itemList_web As ArrayList = CSV_READ(fName)

            Dim lines As String = AzukiControl6.Text
            lines = Replace(lines, vbCrLf, vbLf)
            lines = Replace(lines, vbCr, vbLf)
            Dim lineAZ As String() = Split(lines, vbLf)
            AzukiControl6.Text = ""

            ToolStripProgressBar1.Value = 0
            ToolStripProgressBar1.Maximum = lineAZ.Length

            '読み込みファイルの1行目からヘッダー検索して列番号を取得
            Dim iFnum1 As Integer = 0
            Dim iFnum2 As Integer = 0
            Dim iFnum3 As Integer = 0
            Dim itemListW2 As String() = Split(itemList_web(0), "|=|")
            For c As Integer = 0 To itemListW2.Length - 1
                Select Case itemListW2(c)
                    Case "商品番号"
                        iFnum1 = c
                    Case "販売価格"
                        iFnum2 = c
                    Case "倉庫指定"
                        iFnum3 = c
                End Select
            Next

            'Dim cateUrl As String = ""
            'If RadioButton1.Checked = True Then
            '    cateUrl = "https:item.rakuten.co.jppatri"
            'ElseIf RadioButton4.Checked = True Then
            '    cateUrl = "https:item.rakuten.co.jpashop"
            'Else
            '    cateUrl = "https:item.rakuten.co.jpalice-zk"
            'End If

            Dim cateUrl As String = "http.*/c/"
            Dim str As String = ""

            For r As Integer = 0 To lineAZ.Length - 1
                If lineAZ(r) <> "" Then
                    Dim lineSTR As String() = Split(lineAZ(r), ",") '商品コード,商品名,価格,商品URL,画像URL,更新日,カテゴリ,倉庫
                    'Dim code As String = Replace(lineSTR(3), "/", "")
                    'code = Replace(code, cateUrl, "")    'ディレクトリを分解する
                    Dim code As String = lineSTR(0)

                    Dim souko As String = ""
                    If lineSTR.Length >= 7 Then '倉庫の指定がある場合
                        souko = lineSTR(7)
                    End If

                    For i As Integer = 0 To itemList_web.Count - 1
                        Dim itemList As String() = Split(itemList_web(i), "|=|")
                        If code = itemList(iFnum1) Then
                            If lineSTR(2) = itemList(iFnum2) Then    '価格変更なし
                                If souko = "" Then
                                    str &= lineAZ(r)
                                Else
                                    For k As Integer = 0 To lineSTR.Length - 1
                                        If k = lineSTR.Length - 1 Then
                                            str &= itemList(iFnum3)
                                        Else
                                            str &= lineSTR(k) & ","
                                        End If
                                    Next
                                End If
                            Else
                                If souko = "" Then
                                    For k As Integer = 0 To lineSTR.Length - 1
                                        If k = lineSTR.Length - 1 Then
                                            str &= lineSTR(k)
                                        ElseIf k = 2 Then
                                            str &= itemList(iFnum2) & ","
                                        Else
                                            str &= lineSTR(k) & ","
                                        End If
                                    Next
                                    changeStr &= lineSTR(0) & " " & lineSTR(2) & " → " & itemList(iFnum2) & vbCrLf
                                    changeNum += 1
                                Else
                                    For k As Integer = 0 To lineSTR.Length - 1
                                        If k = lineSTR.Length - 1 Then
                                            str &= itemList(iFnum3)
                                        ElseIf k = 2 Then
                                            str &= itemList(iFnum2) & ","
                                        Else
                                            str &= lineSTR(k) & ","
                                        End If
                                    Next
                                    changeStr &= lineSTR(0) & " " & lineSTR(2) & " → " & itemList(iFnum2) & vbCrLf
                                    changeNum += 1
                                End If
                            End If
                            str &= vbCrLf
                            ToolStripStatusLabel1.Text = i
                            Exit For
                        End If
                    Next
                End If

                ToolStripProgressBar1.Value += 1
                Application.DoEvents()
            Next

            AzukiControl6.Text = str
        End If

        ToolStripProgressBar1.Value = 0
        MsgBox(changeNum & "件修正しました。" & vbCrLf & changeStr, MsgBoxStyle.SystemModal)
    End Sub

    'csvでカテゴリ更新
    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Dim ofd As New OpenFileDialog With {
            .FileName = "default.html",
            .InitialDirectory = Environment.SpecialFolder.DesktopDirectory
        }
        ofd.FileName = "dl-item-cat.csv"
        ofd.Filter = "CSVファイル(*.csv)|*.csv|すべてのファイル(*.*)|*.*"
        ofd.FilterIndex = 2
        ofd.Title = "開くファイルを選択してください"
        ofd.RestoreDirectory = True
        ofd.CheckFileExists = True
        ofd.CheckPathExists = True

        Dim cateStr As String = ""

        'ダイアログを表示する
        If ofd.ShowDialog() = DialogResult.OK Then
            Dim fName As String = ofd.FileName
            Dim itemList_web As ArrayList = CSV_READ(fName)

            Dim lines As String = AzukiControl6.Text
            lines = Replace(lines, vbCrLf, vbLf)
            lines = Replace(lines, vbCr, vbLf)
            Dim lineAZ As String() = Split(lines, vbLf)
            AzukiControl6.Text = ""

            ToolStripProgressBar1.Value = 0
            ToolStripProgressBar1.Maximum = lineAZ.Length

            Dim cateUrl As String = "http.*/c/"
            Dim str As String = ""

            For r As Integer = 0 To lineAZ.Length - 1
                If lineAZ(r) <> "" Then
                    Dim lineSTR As String() = Split(lineAZ(r), ",") '商品コード,商品名,価格,商品URL,画像URL,更新日,カテゴリ,倉庫
                    'Dim code As String = Replace(lineSTR(3), "/", "")
                    'code = Replace(code, cateUrl, "")    'ディレクトリを分解する
                    Dim code As String = lineSTR(0)

                    Dim yes As Integer = 0
                    Dim cates As String = ""
                    Dim delNum As ArrayList = New ArrayList
                    For i As Integer = itemList_web.Count - 1 To 0 Step -1
                        Dim itemList As String() = Split(itemList_web(i), "|=|")
                        If code = itemList(1) Then
                            Dim catew As String = Regex.Replace(itemList(5), cateUrl, "")
                            catew = Replace(catew, "/", "")
                            If cates = "" Then
                                cates = catew
                            Else
                                cates &= "|" & catew
                            End If
                            itemList_web.RemoveAt(i)
                            ToolStripStatusLabel1.Text = i
                        End If
                    Next

                    For k As Integer = 0 To lineSTR.Length - 1
                        If k = 6 Then
                            str &= cates & ","
                            If lineSTR(k) <> cates Then
                                cateStr &= lineSTR(0) & vbCrLf
                            End If
                        Else
                            str &= lineSTR(k) & ","
                        End If
                    Next
                    str &= vbCrLf
                End If

                ToolStripProgressBar1.Value += 1
                Application.DoEvents()
            Next

            AzukiControl6.Text = str
        End If

        ToolStripProgressBar1.Value = 0
        Dim cLines As String() = Split(cateStr, vbCrLf)
        If cLines.Length > 10 Then
            For i As Integer = 0 To 10
                cateStr &= cLines(i) & vbCrLf
            Next
            cateStr &= "..."
        End If
        MsgBox(cateStr & "修正しました。", MsgBoxStyle.SystemModal)
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



    Private Sub コピーToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles コピーToolStripMenuItem.Click
        Clipboard.SetText(DataGridView1.SelectedCells(0).Value)
    End Sub

    Private Sub TextBox38_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox38.Click
        TextBox38.SelectAll()
    End Sub

    Private Sub TextBox38_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox38.TextChanged
        PriceCalc2()
    End Sub

    Private Sub ComboBox11_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox11.SelectedIndexChanged
        PriceCalc2()
    End Sub

    Private Sub RadioButton5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton5.CheckedChanged
        ComboBox3.SelectedItem = "海東KT"
        PriceCalc2()
    End Sub

    Private Sub RadioButton9_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton9.CheckedChanged
        ComboBox3.SelectedItem = "あかね楽天"
        PriceCalc2()
    End Sub

    Private Sub RadioButton6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton6.CheckedChanged
        'ComboBox3.SelectedItem = "FKstyle"
        ComboBox3.SelectedItem = "FKstyle(以前)"
        PriceCalc2()
    End Sub

    Private Sub RadioButton7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ComboBox3.SelectedItem = "暁"
        PriceCalc2()
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        For i As Integer = 0 To DataGridView1.RowCount - 1
            If DataGridView1.Item(0, i).Value = ComboBox3.SelectedItem Then
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
            ElseIf DataGridView1.Item(0, i).Value = "FKstyle" Then
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.LightGreen
            ElseIf DataGridView1.Item(0, i).Value = "あかね楽天" And ComboBox3.SelectedItem = "海東KT" Then
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
            ElseIf DataGridView1.Item(0, i).Value = "海東KT" And ComboBox3.SelectedItem = "海東KT" Then
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
            Else
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Empty
            End If
        Next
        Label35.Text = ComboBox3.Text & "の価格"

        PriceCalc2()
    End Sub

    '元データ
    'Private SouryoMail As String() = New String() {"190", "0", "0", "190", "0", "0", "240", "240", "190", "240"}
    'Private SouryoTakuhai As String() = New String() {"750", "0", "0", "530", "0", "0", "770", "756", "530", "750"}
    'Private SouryoOhgata As String() = New String() {"990", "0", "0", "1260", "0", "0", "1250", "972", "1260", "1250"}
    'Private SouryoMail As String() = New String() {"190", "120", "0", "0", "190", "0", "0", "0", "240", "190", "0", "0", "0"}
    'Private SouryoTakuhai As String() = New String() {"750", "650", "0", "0", "530", "0", "0", "0", "756", "530", "0", "0", "0"}
    'Private SouryoOhgata As String() = New String() {"990", "890", "0", "0", "1260", "0", "0", "0", "972", "1260", "0", "0", "0"}

    Private SouryoMail As String() = New String() {"190", "190", "0", "0", "190", "0", "0", "0", "240", "0", "190", "0", "0", "0"}
    Private SouryoTakuhai As String() = New String() {"750", "750", "0", "0", "530", "0", "0", "0", "756", "0", "530", "0", "0", "0"}
    Private SouryoOhgata As String() = New String() {"990", "990", "0", "0", "1260", "0", "0", "0", "972", "0", "1260", "0", "0", "0"}
    Private Sub PriceCalc2()
        '"FKstyle", "Lucky9", "あかねYahoo", "暁", "アリス", "あかね楽天", "KuraNavi", "雑貨倉庫", "海東KT, "問よか""
        Select Case ComboBox11.SelectedItem
            Case "メール便"
                For i As Integer = 0 To tenpo.Length - 1
                    DataGridView1.Item(2, i).Value = SouryoMail(i)
                Next
            Case "宅配便"
                For i As Integer = 0 To tenpo.Length - 1
                    DataGridView1.Item(2, i).Value = SouryoTakuhai(i)
                Next
            Case "大型送料"
                For i As Integer = 0 To tenpo.Length - 1
                    DataGridView1.Item(2, i).Value = SouryoOhgata(i)
                Next
        End Select

        Dim price As Integer = 0
        Dim priceA As Integer = 0
        Dim priceFK As Integer = 0

        Select Case ComboBox3.Text
            Case "FKstyle(以前)", "FKstyle", "海東KT", "暁", "あかね楽天"
                TextBox7.Text = TextBox38.Text
            Case "Lucky9"
                If RadioButton5.Checked Or RadioButton9.Checked Then
                    TextBox7.Text = Math.Round(TextBox38.Text / 107 * 100)
                ElseIf RadioButton6.Checked Then
                    TextBox7.Text = "計算不能"
                End If
            Case "アリス"
                If RadioButton5.Checked Or RadioButton9.Checked Then
                    TextBox7.Text = Math.Round(TextBox38.Text / 110 * 100)
                ElseIf RadioButton6.Checked Then
                    TextBox7.Text = "計算不能"
                End If
            Case "KuraNavi"
                If RadioButton5.Checked Or RadioButton9.Checked Then
                    TextBox7.Text = "計算不能"
                ElseIf RadioButton6.Checked Then
                    TextBox7.Text = "計算不能"
                End If
            Case "問よか"
                If RadioButton5.Checked Or RadioButton9.Checked Then
                    TextBox7.Text = "計算不能"
                ElseIf RadioButton6.Checked Then
                    TextBox7.Text = "計算不能"
                End If
            Case "雑貨倉庫"
                If RadioButton5.Checked Or RadioButton9.Checked Then
                    TextBox7.Text = "計算不能"
                ElseIf RadioButton6.Checked Then
                    TextBox7.Text = "計算不能"
                End If
            Case "あかねYahoo"
                If RadioButton5.Checked Or RadioButton9.Checked Then
                    TextBox7.Text = "計算不能"
                ElseIf RadioButton6.Checked Then
                    TextBox7.Text = "計算不能"
                    'If ComboBox11.Text = "メール便" Then
                    '    If TextBox38.Text <= 600 Then
                    '        TextBox7.Text = TextBox38.Text + 30
                    '    Else
                    '        TextBox7.Text = TextBox38.Text + 40
                    '    End If
                    'ElseIf ComboBox11.Text = "宅配便" Then
                    '    TextBox7.Text = TextBox38.Text + 40
                    'ElseIf ComboBox11.Text = "大型送料" Then
                    '    TextBox7.Text = TextBox38.Text + 300
                    'End If
                End If
            Case Else
                TextBox7.Text = "計算不能"
        End Select

        If TextBox7.Text = "計算不能" Then
            TextBox7.BackColor = Color.Red
            TextBox7.ForeColor = Color.White
            Exit Sub
        Else
            TextBox7.BackColor = Color.Lavender
            TextBox7.ForeColor = Color.Black
        End If

        If TextBox7.Text = "" Then
            TextBox7.Text = 0
        ElseIf Not IsNumeric(TextBox7.Text) Then
            TextBox7.Text = 0
        End If

        Dim tenpoArray As New ArrayList(tenpo)
        Dim dH1 As ArrayList = TM_HEADER_GET(DataGridView1)

        For r As Integer = 0 To DataGridView1.RowCount - 1
            Select Case r
                Case tenpoArray.IndexOf("FKstyle(以前)")
                    If RadioButton5.Checked Or RadioButton9.Checked Then
                        priceA = TextBox7.Text - DataGridView1.Item(2, r).Value
                        If priceA <= 1999 Then
                            price = priceA + 90
                            TextBox40.Text = 90
                        Else
                            price = priceA + 150
                            TextBox40.Text = 150
                        End If
                        DataGridView1.Item(dH1.IndexOf("価格"), r).Value = price
                    ElseIf RadioButton6.Checked = True Then
                        DataGridView1.Item(dH1.IndexOf("価格"), r).Value = TextBox7.Text
                    Else
                        If ComboBox11.SelectedItem = "メール便" Then
                            DataGridView1.Item(dH1.IndexOf("価格"), r).Value = TextBox7.Text + 51
                        ElseIf ComboBox11.SelectedItem = "宅配便" Then
                            DataGridView1.Item(dH1.IndexOf("価格"), r).Value = TextBox7.Text + 10
                        Else
                            DataGridView1.Item(dH1.IndexOf("価格"), r).Value = TextBox7.Text + 10
                        End If
                    End If

                    If RadioButton6.Checked = True Then
                        priceFK = CInt(DataGridView1.Item(dH1.IndexOf("価格"), r).Value) + CInt(DataGridView1.Item(dH1.IndexOf("送料"), r).Value)
                        If priceFK <= 1999 Then
                            priceFK = priceFK - 90
                            TextBox40.Text = "-90"
                        Else
                            priceFK = priceFK - 150
                            TextBox40.Text = "-150"
                        End If
                    End If
                Case tenpoArray.IndexOf("FKstyle")
                    If RadioButton5.Checked Or RadioButton9.Checked Then
                        priceA = TextBox7.Text - DataGridView1.Item(2, r).Value
                        If priceA <= 1999 Then
                            price = priceA + 90
                            TextBox40.Text = 90
                        Else
                            price = priceA + 150
                            TextBox40.Text = 150
                        End If
                        DataGridView1.Item(dH1.IndexOf("価格"), r).Value = Math.Floor(price * 1.02)
                    ElseIf RadioButton6.Checked = True Then
                        DataGridView1.Item(dH1.IndexOf("価格"), r).Value = TextBox7.Text
                    Else
                        If ComboBox11.SelectedItem = "メール便" Then
                            DataGridView1.Item(dH1.IndexOf("価格"), r).Value = Math.Floor((TextBox7.Text + 51) * 1.02)
                        ElseIf ComboBox11.SelectedItem = "宅配便" Then
                            DataGridView1.Item(dH1.IndexOf("価格"), r).Value = Math.Floor((TextBox7.Text + 10) * 1.02)
                        Else
                            DataGridView1.Item(dH1.IndexOf("価格"), r).Value = Math.Floor((TextBox7.Text + 10) * 1.02)
                        End If
                    End If

                    'If RadioButton6.Checked = True Then
                    '    priceFK = CInt(DataGridView1.Item(dH1.IndexOf("価格"), r).Value) + CInt(DataGridView1.Item(dH1.IndexOf("送料"), r).Value)
                    '    If priceFK <= 1999 Then
                    '        priceFK = priceFK - 90
                    '        TextBox40.Text = "-90"
                    '    Else
                    '        priceFK = priceFK - 150
                    '        TextBox40.Text = "-150"
                    '    End If
                    'End If
                Case tenpoArray.IndexOf("暁")
                    If RadioButton5.Checked Or RadioButton9.Checked Then
                        price = TextBox7.Text - DataGridView1.Item(dH1.IndexOf("送料"), r).Value
                        DataGridView1.Item(dH1.IndexOf("価格"), r).Value = price
                    End If
            End Select
        Next

        '価格計算
        For r As Integer = 0 To DataGridView1.RowCount - 1
            Select Case r
                Case tenpoArray.IndexOf("Lucky9")
                    If RadioButton5.Checked Or RadioButton9.Checked Then
                        DataGridView1.Item(dH1.IndexOf("価格"), r).Value = Math.Floor(TextBox7.Text * 1.07)
                    Else
                        DataGridView1.Item(dH1.IndexOf("価格"), r).Value = Math.Floor(priceFK * 1.07)
                    End If
                Case tenpoArray.IndexOf("アリス")
                    If RadioButton5.Checked Or RadioButton9.Checked Then
                        DataGridView1.Item(dH1.IndexOf("価格"), r).Value = Math.Floor(TextBox7.Text * 1.1)
                    Else
                        DataGridView1.Item(dH1.IndexOf("価格"), r).Value = Math.Floor(priceFK * 1.1)
                    End If
                Case tenpoArray.IndexOf("海東KT"), tenpoArray.IndexOf("あかね楽天")
                    If RadioButton5.Checked Or RadioButton9.Checked Then
                        DataGridView1.Item(dH1.IndexOf("価格"), r).Value = TextBox7.Text
                    Else
                        DataGridView1.Item(dH1.IndexOf("価格"), r).Value = priceFK
                    End If
                Case tenpoArray.IndexOf("KuraNavi")
                    priceA = DataGridView1.Item(dH1.IndexOf("価格"), tenpoArray.IndexOf("FKstyle(以前)")).Value
                    If ComboBox11.SelectedItem = "メール便" Then
                        DataGridView1.Item(dH1.IndexOf("価格"), r).Value = priceA + 240
                    ElseIf ComboBox11.SelectedItem = "宅配便" Then
                        DataGridView1.Item(dH1.IndexOf("価格"), r).Value = priceA + 770 + 43
                    Else
                        DataGridView1.Item(dH1.IndexOf("価格"), r).Value = priceA + 1250 + 70
                    End If
                Case tenpoArray.IndexOf("問よか")
                    priceA = DataGridView1.Item(dH1.IndexOf("価格"), tenpoArray.IndexOf("FKstyle(以前)")).Value
                    price = Math.Floor(priceA * 0.945)
                    DataGridView1.Item(dH1.IndexOf("価格"), r).Value = price
                Case tenpoArray.IndexOf("雑貨倉庫")
                    priceA = DataGridView1.Item(dH1.IndexOf("価格"), tenpoArray.IndexOf("FKstyle(以前)")).Value
                    price = Math.Floor(priceA * 1.05)
                    DataGridView1.Item(dH1.IndexOf("価格"), r).Value = price
                Case tenpoArray.IndexOf("あかねYahoo")
                    'priceA = DataGridView1.Item(dH1.IndexOf("価格"), tenpoArray.IndexOf("海東KT")).Value
                    priceA = DataGridView1.Item(dH1.IndexOf("価格"), tenpoArray.IndexOf("海東KT")).Value
                    price = Math.Floor(priceA * 1.06)
                    DataGridView1.Item(dH1.IndexOf("価格"), r).Value = price
                Case tenpoArray.IndexOf("暁")
                    If RadioButton6.Checked Then
                        priceA = DataGridView1.Item(dH1.IndexOf("価格"), tenpoArray.IndexOf("海東KT")).Value
                        price = priceA - DataGridView1.Item(dH1.IndexOf("送料"), r).Value
                        DataGridView1.Item(dH1.IndexOf("価格"), r).Value = price
                    End If
                Case tenpoArray.IndexOf("Amazon(105%)")
                    price = Math.Round(TextBox38.Text * 1.05)
                    DataGridView1.Item(dH1.IndexOf("価格"), r).Value = price
                Case tenpoArray.IndexOf("Amazon(108%)")
                    price = Math.Round(TextBox38.Text * 1.08)
                    DataGridView1.Item(dH1.IndexOf("価格"), r).Value = price
                Case tenpoArray.IndexOf("Amazon(110%)")
                    price = Math.Round(TextBox38.Text * 1.1)
                    DataGridView1.Item(dH1.IndexOf("価格"), r).Value = price
            End Select
        Next

        For r As Integer = 0 To DataGridView1.RowCount - 1
            DataGridView1.Item(3, r).Value = CInt(DataGridView1.Item(1, r).Value) + CInt(DataGridView1.Item(2, r).Value)
        Next
    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        Dim content As String() = New String() {
             "FKstyle(以前)は送料込の価格から送料を引き、1999円以下は+90円、2000円以上は+150円補正",
             "FKstyleは送料込の価格から送料を引き、1999円以下は+90円、2000円以上は+150円補正",
             "Lucky9は基準店舗の送料込の価格+7%プラス",
             "海東KTは送料込の価格と同じ",
             "暁は基準店舗（あかね楽天）から送料を引く",
             "アリスは基準店舗の送料込の価格+10%プラス",
             "あかね楽天は送料込の価格と同じ",
             "KuraNaviはFK(以前)の価格+(メール便)240円、(宅配便)770+43円、(大型)1250+70円",
             "雑貨倉庫はFK(以前)の価格×1.05（切り捨て）",
             "あかねYahooは海東KT(以前)の価格+6%プラス",
             "問よかはFK(以前)の価格×0.945（切り捨て）",
             "Amazon(5%)は基本価格×1.05（四捨五入）",
             "Amazon(8%)は基本価格×1.08（四捨五入）",
             "Amazon(10%)は基本価格×1.1（四捨五入）"
       }
        '"KTはFK基準で、メール便600円以下-30円、600円以上-40円、宅配便-40円、大型-300円",
        Dim r As Integer = DataGridView1.SelectedCells(0).RowIndex
        TextBox39.Text = content(r) & vbCrLf
        TextBox39.Text &= "（宅）" & SouryoTakuhai(r) & " （大型）" & SouryoOhgata(r) & " （メール便）" & SouryoMail(r)
    End Sub

    'Private Sub PriceCalc4()
    '    '"FKstyle", "Lucky9", "あかねYahoo", "暁", "アリス", "あかね楽天", "KuraNavi", "問よか", "雑貨倉庫", "海東KT"
    '    Dim SouryoMail As String() = New String() {"189", "0", "0", "240", "0", "0", "240", "240", "240", "220"}
    '    Dim SouryoTakuhai As String() = New String() {"750", "0", "0", "760", "0", "0", "756", "750", "756", "760"}
    '    Dim SouryoOhgata As String() = New String() {"972", "0", "0", "1260", "0", "0", "1250", "1250", "972", "1080"}
    '    Select Case ComboBox11.SelectedItem
    '        Case "メール便"
    '            For i As Integer = 0 To tenpo.Length - 1
    '                DataGridView1.Item(2, i).Value = SouryoMail(i)
    '            Next
    '        Case "宅配便"
    '            For i As Integer = 0 To tenpo.Length - 1
    '                DataGridView1.Item(2, i).Value = SouryoTakuhai(i)
    '            Next
    '        Case "大型送料"
    '            For i As Integer = 0 To tenpo.Length - 1
    '                DataGridView1.Item(2, i).Value = SouryoOhgata(i)
    '            Next
    '    End Select

    '    Dim price As Integer = 0
    '    Dim priceA As Integer = 0
    '    Dim priceFK As Integer = 0

    '    Select Case ComboBox3.Text
    '        Case "FKstyle", "あかねYahoo", "暁", "あかね楽天"
    '            TextBox7.Text = TextBox38.Text
    '        Case "Lucky9"
    '            If RadioButton5.Checked Then
    '                TextBox7.Text = Math.Round(TextBox38.Text / 107 * 100)
    '            ElseIf RadioButton6.Checked Then
    '                TextBox7.Text = "計算不能"
    '            ElseIf RadioButton7.Checked Then
    '                TextBox7.Text = "計算不能"
    '            End If
    '        Case "アリス"
    '            If RadioButton5.Checked Then
    '                TextBox7.Text = Math.Round(TextBox38.Text / 110 * 100)
    '            ElseIf RadioButton6.Checked Then
    '                TextBox7.Text = "計算不能"
    '            ElseIf RadioButton7.Checked Then
    '                TextBox7.Text = "計算不能"
    '            End If
    '        Case "KuraNavi"
    '            If RadioButton5.Checked Then
    '                TextBox7.Text = "計算不能"
    '            ElseIf RadioButton6.Checked Then
    '                TextBox7.Text = "計算不能"
    '            ElseIf RadioButton7.Checked Then
    '                TextBox7.Text = TextBox38.Text
    '            End If
    '        Case "問よか"
    '            If RadioButton5.Checked Then
    '                TextBox7.Text = "計算不能"
    '            ElseIf RadioButton6.Checked Then
    '                TextBox7.Text = "計算不能"
    '            ElseIf RadioButton7.Checked Then
    '                If TextBox38.Text >= 4000 Then
    '                    TextBox7.Text = Math.Round(TextBox38.Text * 0.95)
    '                Else
    '                    TextBox7.Text = Math.Round(TextBox38.Text * 0.9)
    '                End If
    '            End If
    '        Case "雑貨倉庫"
    '            If RadioButton5.Checked Then
    '                TextBox7.Text = "計算不能"
    '            ElseIf RadioButton6.Checked Then
    '                TextBox7.Text = "計算不能"
    '            ElseIf RadioButton7.Checked Then
    '                TextBox7.Text = Math.Round(TextBox38.Text * 1.05)
    '            End If
    '        Case "海東KT"
    '            If RadioButton5.Checked Then
    '                TextBox7.Text = "計算不能"
    '            ElseIf RadioButton6.Checked Then
    '                If ComboBox11.Text = "メール便" Then
    '                    If TextBox38.Text <= 600 Then
    '                        TextBox7.Text = TextBox38.Text + 30
    '                    Else
    '                        TextBox7.Text = TextBox38.Text + 40
    '                    End If
    '                ElseIf ComboBox11.Text = "宅配便" Then
    '                    TextBox7.Text = TextBox38.Text + 40
    '                ElseIf ComboBox11.Text = "大型送料" Then
    '                    TextBox7.Text = TextBox38.Text + 300
    '                End If
    '            ElseIf RadioButton7.Checked Then
    '                TextBox7.Text = "計算不能"
    '            End If
    '    End Select

    '    If TextBox7.Text = "計算不能" Then
    '        TextBox7.BackColor = Color.Red
    '        TextBox7.ForeColor = Color.White
    '        Exit Sub
    '    Else
    '        TextBox7.BackColor = Color.Lavender
    '        TextBox7.ForeColor = Color.Black
    '    End If

    '    If TextBox7.Text = "" Then
    '        TextBox7.Text = 0
    '    ElseIf Not IsNumeric(TextBox7.Text) Then
    '        TextBox7.Text = 0
    '    End If

    '    Dim tenpoArray As New ArrayList(tenpo)
    '    Dim dH1 As ArrayList = TM_HEADER_GET(DataGridView1)

    '    For r As Integer = 0 To DataGridView1.RowCount - 1
    '        Select Case r
    '            Case tenpoArray.IndexOf("FKstyle")
    '                If RadioButton5.Checked = True Then
    '                    priceA = TextBox7.Text - DataGridView1.Item(2, r).Value
    '                    If priceA <= 1999 Then
    '                        price = priceA + 90
    '                        TextBox40.Text = 90
    '                    Else
    '                        price = priceA + 150
    '                        TextBox40.Text = 150
    '                    End If
    '                    DataGridView1.Item(dH1.IndexOf("価格"), r).Value = price
    '                ElseIf RadioButton6.Checked = True Then
    '                    DataGridView1.Item(dH1.IndexOf("価格"), r).Value = TextBox7.Text
    '                Else
    '                    If ComboBox11.SelectedItem = "メール便" Then
    '                        DataGridView1.Item(dH1.IndexOf("価格"), r).Value = TextBox7.Text + 51
    '                    ElseIf ComboBox11.SelectedItem = "宅配便" Then
    '                        DataGridView1.Item(dH1.IndexOf("価格"), r).Value = TextBox7.Text + 10
    '                    Else
    '                        DataGridView1.Item(dH1.IndexOf("価格"), r).Value = TextBox7.Text + 10
    '                    End If
    '                End If

    '                If RadioButton6.Checked = True Or RadioButton7.Checked = True Then
    '                    priceFK = CInt(DataGridView1.Item(dH1.IndexOf("価格"), r).Value) + CInt(DataGridView1.Item(dH1.IndexOf("送料"), r).Value)
    '                    If priceFK <= 1999 Then
    '                        priceFK = priceFK - 90
    '                        TextBox40.Text = "-90"
    '                    Else
    '                        priceFK = priceFK - 150
    '                        TextBox40.Text = "-150"
    '                    End If
    '                End If
    '            Case tenpoArray.IndexOf("暁")
    '                If RadioButton5.Checked = True Then
    '                    priceA = TextBox7.Text - DataGridView1.Item(dH1.IndexOf("送料"), r).Value
    '                    price = priceA + TextBox40.Text
    '                    DataGridView1.Item(dH1.IndexOf("価格"), r).Value = price
    '                ElseIf RadioButton6.Checked = True Then
    '                    If ComboBox11.SelectedItem = "メール便" Then
    '                        DataGridView1.Item(dH1.IndexOf("価格"), r).Value = TextBox7.Text - 51
    '                    ElseIf ComboBox11.SelectedItem = "宅配便" Then
    '                        DataGridView1.Item(dH1.IndexOf("価格"), r).Value = TextBox7.Text - 10
    '                    Else
    '                        DataGridView1.Item(dH1.IndexOf("価格"), r).Value = TextBox7.Text - 10
    '                    End If
    '                Else
    '                    DataGridView1.Item(dH1.IndexOf("価格"), r).Value = TextBox7.Text
    '                End If
    '        End Select
    '    Next

    '    '価格計算
    '    For r As Integer = 0 To DataGridView1.RowCount - 1
    '        Select Case r
    '            Case tenpoArray.IndexOf("Lucky9")
    '                If RadioButton5.Checked = True Then
    '                    DataGridView1.Item(dH1.IndexOf("価格"), r).Value = Math.Floor(TextBox7.Text * 1.07)
    '                Else
    '                    DataGridView1.Item(dH1.IndexOf("価格"), r).Value = Math.Floor(priceFK * 1.07)
    '                End If
    '            Case tenpoArray.IndexOf("アリス")
    '                If RadioButton5.Checked = True Then
    '                    DataGridView1.Item(dH1.IndexOf("価格"), r).Value = Math.Floor(TextBox7.Text * 1.1)
    '                Else
    '                    DataGridView1.Item(dH1.IndexOf("価格"), r).Value = Math.Floor(priceFK * 1.1)
    '                End If
    '            Case tenpoArray.IndexOf("あかねYahoo"), tenpoArray.IndexOf("あかね楽天")
    '                If RadioButton5.Checked = True Then
    '                    DataGridView1.Item(dH1.IndexOf("価格"), r).Value = TextBox7.Text
    '                Else
    '                    DataGridView1.Item(dH1.IndexOf("価格"), r).Value = priceFK
    '                End If
    '            Case tenpoArray.IndexOf("KuraNavi")
    '                DataGridView1.Item(dH1.IndexOf("価格"), r).Value = DataGridView1.Item(dH1.IndexOf("価格"), 3).Value
    '            Case tenpoArray.IndexOf("問よか")
    '                priceA = DataGridView1.Item(dH1.IndexOf("価格"), tenpoArray.IndexOf("暁")).Value
    '                If priceA >= 4000 Then
    '                    price = Math.Floor(priceA * 0.95)
    '                Else
    '                    price = Math.Floor(priceA * 0.9)
    '                End If
    '                DataGridView1.Item(dH1.IndexOf("価格"), r).Value = price
    '            Case tenpoArray.IndexOf("雑貨倉庫")
    '                priceA = DataGridView1.Item(dH1.IndexOf("価格"), tenpoArray.IndexOf("暁")).Value
    '                price = Math.Floor(priceA * 1.05)
    '                DataGridView1.Item(dH1.IndexOf("価格"), r).Value = price
    '            Case tenpoArray.IndexOf("海東KT")
    '                priceA = DataGridView1.Item(dH1.IndexOf("価格"), tenpoArray.IndexOf("FKstyle")).Value
    '                If ComboBox11.SelectedItem = "メール便" Then
    '                    If priceA <= 600 Then
    '                        DataGridView1.Item(dH1.IndexOf("価格"), r).Value = priceA - 30
    '                    Else
    '                        DataGridView1.Item(dH1.IndexOf("価格"), r).Value = priceA - 40
    '                    End If
    '                ElseIf ComboBox11.SelectedItem = "宅配便" Then
    '                    DataGridView1.Item(dH1.IndexOf("価格"), r).Value = priceA - 40
    '                Else
    '                    DataGridView1.Item(dH1.IndexOf("価格"), r).Value = priceA - 300
    '                End If
    '        End Select
    '    Next

    '    For r As Integer = 0 To DataGridView1.RowCount - 1
    '        DataGridView1.Item(3, r).Value = CInt(DataGridView1.Item(1, r).Value) + CInt(DataGridView1.Item(2, r).Value)
    '    Next
    'End Sub


    'Private Sub PriceCalc3()
    '    '"FKstyle", "Lucky9", "あかねYahoo", "暁", "アリス", "あかね楽天", "KuraNavi", "問よか", "雑貨倉庫", "海東KT"
    '    Dim SouryoMail As String() = New String() {"189", "0", "0", "240", "0", "0", "240", "240", "240", "220"}
    '    Dim SouryoTakuhai As String() = New String() {"750", "0", "0", "760", "0", "0", "756", "750", "756", "760"}
    '    Dim SouryoOhgata As String() = New String() {"972", "0", "0", "1260", "0", "0", "1250", "1250", "972", "1080"}
    '    Select Case ComboBox11.SelectedItem
    '        Case "メール便"
    '            For i As Integer = 0 To tenpo.Length - 1
    '                DataGridView1.Item(2, i).Value = SouryoMail(i)
    '            Next
    '        Case "宅配便"
    '            For i As Integer = 0 To tenpo.Length - 1
    '                DataGridView1.Item(2, i).Value = SouryoTakuhai(i)
    '            Next
    '        Case "大型送料"
    '            For i As Integer = 0 To tenpo.Length - 1
    '                DataGridView1.Item(2, i).Value = SouryoOhgata(i)
    '            Next
    '    End Select

    '    Dim kijunPrice As Integer = 0
    '    Select Case ComboBox3.SelectedItem
    '        Case "FKstyle", "暁", "KuraNavi"
    '            If TextBox38.Text <= 1999 + 90 Then
    '                kijunPrice = TextBox38.Text - 90
    '            Else
    '                kijunPrice = TextBox38.Text - 150
    '            End If
    '            kijunPrice = kijunPrice + DataGridView1.Item(2, ComboBox3.Items.IndexOf(ComboBox3.SelectedItem)).Value
    '            Label6.Text = DataGridView1.Item(2, ComboBox3.Items.IndexOf(ComboBox3.SelectedItem)).Value
    '        Case "あかねYahoo", "あかね楽天"
    '            kijunPrice = TextBox38.Text
    '            Label6.Text = 0
    '        Case "Lucky9"
    '            kijunPrice = Math.Round(TextBox38.Text / 107 * 100)
    '            Label6.Text = 0
    '        Case "アリス"
    '            kijunPrice = Math.Round(TextBox38.Text / 110 * 100)
    '            Label6.Text = 0
    '        Case "問よか"
    '            If TextBox38.Text <= 3999 Then
    '                kijunPrice = Math.Round(TextBox38.Text / 90 * 100)
    '            Else
    '                kijunPrice = Math.Round(TextBox38.Text / 95 * 100)
    '            End If
    '            kijunPrice = kijunPrice + DataGridView1.Item(2, ComboBox3.Items.IndexOf(ComboBox3.SelectedItem)).Value
    '            Label6.Text = DataGridView1.Item(2, ComboBox3.Items.IndexOf(ComboBox3.SelectedItem)).Value
    '        Case "雑貨倉庫"
    '            kijunPrice = Math.Round(TextBox38.Text / 105 * 100)
    '            If kijunPrice <= 1999 + 90 Then
    '                kijunPrice = kijunPrice - 90
    '            Else
    '                kijunPrice = kijunPrice - 150
    '            End If
    '            kijunPrice = kijunPrice + DataGridView1.Item(2, ComboBox3.Items.IndexOf("暁")).Value
    '            Label6.Text = DataGridView1.Item(2, ComboBox3.Items.IndexOf("暁")).Value
    '        Case "海東KT"
    '            If ComboBox11.SelectedItem = "メール便" Then
    '                If TextBox38.Text <= 600 + 30 Then
    '                    kijunPrice = TextBox38.Text + 30
    '                Else
    '                    kijunPrice = TextBox38.Text + 40
    '                End If
    '            ElseIf ComboBox11.SelectedItem = "宅配便" Then
    '                kijunPrice = TextBox38.Text + 40
    '            ElseIf ComboBox11.SelectedItem = "大型送料" Then
    '                kijunPrice = TextBox38.Text + 300
    '            End If
    '            If kijunPrice <= 1999 + 90 Then
    '                kijunPrice = kijunPrice - 90
    '            Else
    '                kijunPrice = kijunPrice - 150
    '            End If
    '            kijunPrice = kijunPrice + DataGridView1.Item(2, ComboBox3.Items.IndexOf("FKstyle")).Value
    '            Label6.Text = DataGridView1.Item(2, ComboBox3.Items.IndexOf("FKstyle")).Value
    '    End Select

    '    TextBox6.Text = kijunPrice

    '    For r As Integer = 0 To DataGridView1.RowCount - 1
    '        Select Case DataGridView1.Item(0, r).Value
    '            Case "FKstyle", "暁", "KuraNavi"
    '                Dim price As Integer = kijunPrice - DataGridView1.Item(2, r).Value
    '                If price <= 1999 Then
    '                    DataGridView1.Item(1, r).Value = price + 90
    '                Else
    '                    DataGridView1.Item(1, r).Value = price + 150
    '                End If
    '            Case "あかねYahoo", "あかね楽天"
    '                DataGridView1.Item(1, r).Value = kijunPrice
    '            Case "Lucky9"
    '                DataGridView1.Item(1, r).Value = Math.Round(kijunPrice * 1.07)
    '            Case "アリス"
    '                DataGridView1.Item(1, r).Value = Math.Round(kijunPrice * 1.1)
    '            Case "問よか"
    '                Dim price As Integer = kijunPrice
    '                If price <= 3999 Then
    '                    DataGridView1.Item(1, r).Value = Math.Round(price * 0.9)
    '                Else
    '                    DataGridView1.Item(1, r).Value = Math.Round(price * 0.95)
    '                End If
    '            Case "雑貨倉庫"
    '                Dim price As Integer = kijunPrice - DataGridView1.Item(2, r).Value
    '                If price <= 1999 Then
    '                    DataGridView1.Item(1, r).Value = Math.Round((price + 90) * 1.05)
    '                Else
    '                    DataGridView1.Item(1, r).Value = Math.Round((price + 150) * 1.05)
    '                End If
    '            Case "海東KT"
    '                Dim price As Integer = kijunPrice - DataGridView1.Item(2, r).Value
    '                If price <= 1999 Then
    '                    price = price + 90
    '                Else
    '                    price = price + 150
    '                End If
    '                If ComboBox11.SelectedItem = "メール便" Then
    '                    If price <= 600 Then
    '                        DataGridView1.Item(1, r).Value = price - 30
    '                    Else
    '                        DataGridView1.Item(1, r).Value = price - 40
    '                    End If
    '                ElseIf ComboBox11.SelectedItem = "宅配便" Then
    '                    DataGridView1.Item(1, r).Value = price - 40
    '                ElseIf ComboBox11.SelectedItem = "大型送料" Then
    '                    DataGridView1.Item(1, r).Value = price - 300
    '                End If
    '        End Select
    '    Next

    '    For r As Integer = 0 To DataGridView1.RowCount - 1
    '        DataGridView1.Item(3, r).Value = CInt(DataGridView1.Item(1, r).Value) + CInt(DataGridView1.Item(2, r).Value)
    '    Next
    'End Sub




    '***************************************************************************************************************
    ' FTP
    '***************************************************************************************************************
    Private FtpClient As TKFP.Net.FtpClient
    Private CurrentDirectory As TKFP.IO.DirectoryInfo

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        If Not (FtpClient Is Nothing) Then
            Return
        End If

        Try
            '接続用の情報作成
            Dim bp As New TKFP.Net.BasicFtpLogon(txtUserID.Text, txtPassword.Text)
            'POPクライアントクラスの定義
            Dim portNum As Integer = TextBox11.Text
            FtpClient = New TKFP.Net.FtpClient(bp, txtAddress.Text, portNum)
            'パッシブモードに設定
            If CheckBox3.Checked = True Then
                FtpClient.ConnectionMode = TKFP.Net.ConnectionModes.Passive
            Else
                FtpClient.ConnectionMode = TKFP.Net.ConnectionModes.Active
            End If
            'リストの取得コマンドを設定
            FtpClient.ListType = TKFP.Net.ListType.LIST
            '取得したリスト情報の有効時間を設定
            FtpClient.ListCacheValidityInterval = 60
            FtpClient.FileSystemCacheValidityInterval = 60

            'FTP over SSL/TLSを使用する場合は以下のように設定します
            'FtpClient.AuthenticationProtocol = TKFP.Net.AuthenticationProtocols.Explicit_TLS
            'AddHandler FtpClient.CertificateValidation, AddressOf FtpClient_CertificateValidation

            'LIST情報の解析クラスを指定
            FtpClient.ListDataLoader = New TKFP.IO.UnixListDataLoader        '標準のＦＴＰ（デフォルト）
            'FtpClient.ListDataLoader = New TKFP.IO.MsdosListDataLoader      'IISでディレクトリ表示スタイルをMS-DOSとした場合

            '日本語ファイル名の文字コードを指定
            'FtpClient.FileNameCharset = "sjis" 'シフトJIS（デフォルト）
            FtpClient.FileNameCharset = "utf-8" 'シフトJIS（デフォルト）

            'デバッグ用のログ出力ファイルを指定
            'FtpClient.DebugLogFileName = "d:\log.txt"


            'モニタ表示のために送受信イベントを登録
            AddHandler FtpClient.MessageReceive, AddressOf FtpClient_MessageReceive
            AddHandler FtpClient.MessageSend, AddressOf FtpClient_MessageSend
            '接続開始
            If Not FtpClient.Connect() Then
                ToolStripStatusLabel2.Text = "接続失敗"
                FtpClient.Close()
                FtpClient = Nothing
                Return
            Else
                ToolStripStatusLabel2.Text = "接続成功"
                If FtpClient.IsEncrypted Then
                    ToolStripStatusLabel2.Text += " (暗号化)"
                End If

                CurrentDirectory = New TKFP.IO.DirectoryInfo(FtpClient, txtDirectory.Text)
                ShowFolder()
            End If
        Catch ex As Exception
            RichTextBox1.AppendText(ex.Message)
            RichTextBox1.Focus()
            RichTextBox1.Refresh()
            RichTextBox1.ScrollToCaret()
        End Try
    End Sub 'button1_Click

    '/ サーバーの切断処理
    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        ServerClose()
    End Sub 'button2_Click

    Private Sub ServerClose()
        If FtpClient Is Nothing Then
            Return
        End If
        FtpClient.Close()
        FtpClient = Nothing
        Me.lviFolder.Items.Clear()
        ToolStripStatusLabel2.Text = "切断"
    End Sub

    '/ メッセージの受信イベント
    Private Sub FtpClient_MessageReceive(ByVal sender As Object, ByVal e As TKFP.Net.MessageArgs)
        RichTextBox1.SelectionColor = System.Drawing.Color.Blue
        RichTextBox1.AppendText(e.Message)
        RichTextBox1.Focus()
        RichTextBox1.Refresh()
        RichTextBox1.ScrollToCaret()
    End Sub 'FtpClient_MessageReceive

    '/ メッセージの送信イベント
    Private Sub FtpClient_MessageSend(ByVal sender As Object, ByVal e As TKFP.Net.MessageArgs)
        RichTextBox1.SelectionColor = System.Drawing.Color.Red
        RichTextBox1.AppendText(e.Message)
        RichTextBox1.Focus()
        RichTextBox1.Refresh()
        RichTextBox1.ScrollToCaret()
    End Sub 'FtpClient_MessageSend

    '/ フォルダの内容を表示する
    Private Sub ShowFolder()
        ToolStripStatusLabel3.Text = CurrentDirectory.FullName
        'If Not CurrentDirectory.Parent.FullName.StartsWith(txtDirectory.Text) Then
        '    Button3.Enabled = False
        'Else
        '    Button3.Enabled = True
        'End If
        lviFolder.Items.Clear()
        'サブフォルダの一覧取得
        Dim d As TKFP.IO.DirectoryInfo
        For Each d In CurrentDirectory.GetDirectories()
            lviFolder.Items.Add(New DirectoryItem(d))
        Next d
        'ファイル一覧取得
        Dim f As TKFP.IO.FileInfo
        For Each f In CurrentDirectory.GetFiles()
            lviFolder.Items.Add(New FileItem(f))
        Next f
    End Sub

    'ShowFolder
    '/ フォルダの選択処理
    Private Sub LviFolder_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lviFolder.DoubleClick
        If lviFolder.SelectedItems.Count = 0 Then
            Return
        End If
        Dim item As System.Windows.Forms.ListViewItem = lviFolder.SelectedItems(0)

        If TypeOf item Is DirectoryItem Then
            CurrentDirectory = CType(item, DirectoryItem).Directory
            ShowFolder()
        ElseIf TypeOf item Is FileItem Then
            Dim file As TKFP.IO.FileInfo = CType(item, FileItem).File
            Select Case file.Extension
                Case ".html", ".htm", ".css", ".txt", ".php", ".js"
                    Dim str As String = ""
                    Using fs As TKFP.IO.FileStream = file.OpenRead()
                        Dim b(1024 * 10) As Byte
                        Dim temp As System.Text.UTF8Encoding = New System.Text.UTF8Encoding(True)

                        Do While fs.Read(b, 0, b.Length) > 0
                            str &= temp.GetString(b)
                        Loop
                    End Using
                    HTMLcreate.az.Text = str
                Case Else
                    MsgBox("読み込めません", MsgBoxStyle.SystemModal)
            End Select
        End If
    End Sub 'lviFolder_DoubleClick

    'パーミッション取得
    Private Sub LviFolder_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lviFolder.Click
        If lviFolder.SelectedItems.Count = 0 Then
            Return
        End If
        Dim item As System.Windows.Forms.ListViewItem = lviFolder.SelectedItems(0)

        Dim permissionStr As String = ""
        If TypeOf item Is DirectoryItem Then
            CurrentDirectory = CType(item, DirectoryItem).Directory
            permissionStr = CurrentDirectory.Permission.PermissionCode
        ElseIf TypeOf item Is FileItem Then
            Dim file As TKFP.IO.FileInfo = CType(item, FileItem).File
            permissionStr = file.Permission.PermissionCode
        End If

        TextBox23.Text = permissionStr
    End Sub

    Private Sub TextBox23_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox23.TextChanged
        Dim permissionStr As String = TextBox23.Text

        Dim p As Integer = permissionStr.Substring(0, 1)
        If p >= 4 Then
            CheckBox8.Checked = True
        Else
            CheckBox8.Checked = False
        End If
        If 2 = p Or p = 3 Or p = 6 Or p = 7 Then
            CheckBox9.Checked = True
        Else
            CheckBox9.Checked = False
        End If
        If p Mod 2 = 1 Then
            CheckBox10.Checked = True
        Else
            CheckBox10.Checked = False
        End If
        p = permissionStr.Substring(1, 1)
        If p >= 4 Then
            CheckBox11.Checked = True
        Else
            CheckBox11.Checked = False
        End If
        If 2 = p Or p = 3 Or p = 6 Or p = 7 Then
            CheckBox13.Checked = True
        Else
            CheckBox13.Checked = False
        End If
        If p Mod 2 = 1 Then
            CheckBox12.Checked = True
        Else
            CheckBox12.Checked = False
        End If
        p = permissionStr.Substring(2, 1)
        If p >= 4 Then
            CheckBox14.Checked = True
        Else
            CheckBox14.Checked = False
        End If
        If 2 = p Or p = 3 Or p = 6 Or p = 7 Then
            CheckBox16.Checked = True
        Else
            CheckBox16.Checked = False
        End If
        If p Mod 2 = 1 Then
            CheckBox15.Checked = True
        Else
            CheckBox15.Checked = False
        End If
    End Sub

    Private Sub CheckBoxPerm_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
        CheckBox8.CheckedChanged, CheckBox9.CheckedChanged, CheckBox10.CheckedChanged,
        CheckBox11.CheckedChanged, CheckBox12.CheckedChanged, CheckBox13.CheckedChanged,
        CheckBox14.CheckedChanged, CheckBox15.CheckedChanged, CheckBox16.CheckedChanged

        Dim a As Integer = 0
        Dim b As Integer = 0
        Dim c As Integer = 0
        If CheckBox8.Checked = True Then a = 4 Else a = 0
        If CheckBox9.Checked = True Then a += 2 Else a += 0
        If CheckBox10.Checked = True Then a += 1 Else a += 0
        If CheckBox11.Checked = True Then b = 4 Else b = 0
        If CheckBox13.Checked = True Then b += 2 Else b += 0
        If CheckBox12.Checked = True Then b += 1 Else b += 0
        If CheckBox14.Checked = True Then c = 4 Else c = 0
        If CheckBox16.Checked = True Then c += 2 Else c += 0
        If CheckBox15.Checked = True Then c += 1 Else c += 0
        TextBox23.Text = a & b & c
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        If lviFolder.SelectedItems.Count = 0 Then
            Return
        End If
        Dim item As System.Windows.Forms.ListViewItem = lviFolder.SelectedItems(0)

        Dim a As Integer = TextBox23.Text.Substring(0, 1)
        Dim b As Integer = TextBox23.Text.Substring(0, 1)
        Dim c As Integer = TextBox23.Text.Substring(0, 1)
        Dim per As TKFP.IO.Permission = New TKFP.IO.Permission(a, b, c)

        Try
            Dim permissionStr As String = ""
            If TypeOf item Is DirectoryItem Then
                CurrentDirectory = CType(item, DirectoryItem).Directory
                CurrentDirectory.Permission = per
            ElseIf TypeOf item Is FileItem Then
                Dim file As TKFP.IO.FileInfo = CType(item, FileItem).File
                file.Permission = per
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.SystemModal)
        End Try
    End Sub


    '/ 親ディレクトリへ移動
    Private Sub ToolStripButton1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        If FtpClient Is Nothing Then
            Return
        End If
        CurrentDirectory = CurrentDirectory.Parent
        ShowFolder()
    End Sub 'button3_Click

    '/ 選択ファイルのダウンロード
    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        If FtpClient Is Nothing Then
            Return
        End If
        If lviFolder.SelectedItems.Count = 0 Then
            Return
        End If
        Dim item As System.Windows.Forms.ListViewItem = lviFolder.SelectedItems(0)

        If TypeOf item Is FileItem Then
            'サーバーファイル情報
            Dim file As TKFP.IO.FileInfo = CType(item, FileItem).File
            'デフォルトファイル名
            saveFileDialog1.FileName = file.Name
            Select Case saveFileDialog1.ShowDialog()
                Case System.Windows.Forms.DialogResult.OK
                    'ダウンロード
                    file.CopyTo(saveFileDialog1.FileName, True)
            End Select
        End If
    End Sub 'button4_Click

    '/ ファイルのアップロード
    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        If FtpClient Is Nothing Then
            Return
        End If
        Select Case openFileDialog1.ShowDialog()
            Case System.Windows.Forms.DialogResult.OK
                'ローカルのファイル情報
                Dim LocalFile As New System.IO.FileInfo(openFileDialog1.FileName)
                'サーバー用のファイル名作成
                Dim ServerFile As String = CurrentDirectory.FullName + LocalFile.Name
                'サーバー用ファイル情報クラス作成
                Dim file As New TKFP.IO.FileInfo(CurrentDirectory.FtpClient, ServerFile)
                'アップロード
                file.ReadFrom(LocalFile.FullName)
                'カレントフォルダの情報更新
                CurrentDirectory.Refresh()
                '表示情報更新
                ShowFolder()
        End Select
    End Sub 'button5_Click

    'SSLの証明書の確認イベント
    Private Sub FtpClient_CertificateValidation(ByVal sender As Object, ByVal e As TKFP.Net.CertificateValidationArgs)
        '全ての確認要求を許可します
        e.Cancel = False
    End Sub

    '/ ディレクトリ用のListViewItem
    Private Class DirectoryItem
        Inherits System.Windows.Forms.ListViewItem

        Public Directory As TKFP.IO.DirectoryInfo

        Public Sub New(ByVal Directory As TKFP.IO.DirectoryInfo)
            Me.Directory = Directory

            Me.Text = Directory.Name
            Me.SubItems.Add("<dir>")
            Me.SubItems.Add(Directory.Permission.PermissionCode)
            Me.ImageIndex = 0
        End Sub 'New
    End Class 'DirectoryItem

    '/ ファイル用のListViewItem
    Private Class FileItem
        Inherits System.Windows.Forms.ListViewItem
        Public File As TKFP.IO.FileInfo


        Public Sub New(ByVal File As TKFP.IO.FileInfo)
            Me.File = File

            Me.Text = File.Name
            Me.SubItems.Add(File.Length.ToString("#,##0"))
            Me.SubItems.Add(File.Permission.PermissionCode)
            Me.ImageIndex = 1
        End Sub 'New
    End Class 'FileItem 

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\ftplist\" & ComboBox4.SelectedItem & ".txt"
        TextBox12.Text = ComboBox4.SelectedItem
        ServerListChange(fName)
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox5.SelectedIndexChanged
        Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\ftplist\" & ComboBox5.SelectedItem & ".txt"
        TextBox12.Text = ComboBox5.SelectedItem
        ServerListChange(fName)
    End Sub

    Private Sub ServerListChange(ByVal fName As String)
        Using sr As New System.IO.StreamReader(fName, System.Text.Encoding.Default)
            Do While Not sr.EndOfStream
                Dim s As String = sr.ReadLine
                If InStr(s, "server=") > 0 Then
                    txtAddress.Text = Replace(s, "server=", "")
                ElseIf InStr(s, "id=") > 0 Then
                    txtUserID.Text = Replace(s, "id=", "")
                ElseIf InStr(s, "pass=") > 0 Then
                    txtPassword.Text = Replace(s, "pass=", "")
                ElseIf InStr(s, "dir=") > 0 Then
                    txtDirectory.Text = Replace(s, "dir=", "")
                ElseIf InStr(s, "port=") > 0 Then
                    TextBox11.Text = Replace(s, "port=", "")
                ElseIf InStr(s, "pasv=") > 0 Then
                    If Replace(s, "pasv=", "") = True Then
                        CheckBox3.Checked = True
                    Else
                        CheckBox3.Checked = False
                    End If
                Else

                End If
            Loop
        End Using
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        ServerListWrite()
    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        ServerListWrite()
    End Sub

    Private Sub ServerListWrite()
        If TextBox12.Text = "" Then
            MsgBox("名称が入っていません", MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        Dim str As String = ""
        str &= "server=" & txtAddress.Text & vbCrLf
        str &= "id=" & txtUserID.Text & vbCrLf
        str &= "pass=" & txtPassword.Text & vbCrLf
        str &= "dir=" & txtDirectory.Text & vbCrLf
        str &= "port=" & TextBox11.Text & vbCrLf
        If CheckBox3.Checked = True Then
            str &= "pasv=true" & vbCrLf
        Else
            str &= "pasv=false" & vbCrLf
        End If

        Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\ftplist\" & TextBox12.Text & ".txt"
        My.Computer.FileSystem.WriteAllText(fName, str, False, Form1.enc)   '初期設定の上書き保存

        ServerListComb()
        MsgBox(fName & vbCrLf & "保存しました", MsgBoxStyle.SystemModal)
    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\ftplist\" & TextBox12.Text & ".txt"
        If Not File.Exists(fName) Then
            MsgBox("削除ファイルがありません", MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        File.Delete(fName)

        ServerListComb()
        MsgBox(fName & vbCrLf & "ファイルを削除しました", MsgBoxStyle.SystemModal)
    End Sub

    '***************************************************************************************************************

    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            ComboBox7.Enabled = True
            TextBox17.Enabled = True
        Else
            ComboBox7.Enabled = False
            TextBox17.Enabled = False
        End If
        CheckBox6.Checked = False
        CheckBox7.Checked = False
    End Sub

    Private Sub CheckBox6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox6.CheckedChanged
        If CheckBox6.Checked = True Then
            ComboBox8.Enabled = True
            TextBox18.Enabled = True
        Else
            ComboBox8.Enabled = False
            TextBox18.Enabled = False
        End If
        If CheckBox7.Checked = True Then
            CheckBox7.Checked = False
        End If
    End Sub

    Private Sub CheckBox7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox7.CheckedChanged
        If CheckBox7.Checked = True Then
            ComboBox8.Enabled = True
            TextBox18.Enabled = True
        Else
            ComboBox8.Enabled = False
            TextBox18.Enabled = False
        End If
        If CheckBox6.Checked = True Then
            CheckBox6.Checked = False
        End If
    End Sub

    '抽出実行
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim colA As Integer = ComboBox6.SelectedIndex
        Dim searchA As String() = Split(TextBox16.Text, vbCrLf)
        Dim colB As Integer = ComboBox7.SelectedIndex
        Dim searchB As String() = Split(TextBox17.Text, vbCrLf)
        Dim colC As Integer = ComboBox8.SelectedIndex
        Dim searchC As String() = Split(TextBox18.Text, vbCrLf)

        Dim dgv As DataGridView = Nothing
        If ComboBox10.SelectedItem = "1:CSV1" Then
            dgv = Form1.CSV_FORMS(0).DataGridView1
            Form1.CSV_FORMS(0).Panel2.Visible = True
        ElseIf ComboBox10.SelectedItem = "2:CSV2" Then
            dgv = Form1.CSV_FORMS(1).DataGridView1
            Form1.CSV_FORMS(1).Panel2.Visible = True
        End If
        Application.DoEvents()

        Dim start As Integer = 0
        If CheckBox5.Checked = True Then
            If dgv.RowCount > 0 Then
                dgv.Rows(0).Visible = True
                start = 1
            End If
        End If

        '抽出1で表示になる行を保存する（抽出3用）
        Dim rowArray As ArrayList = New ArrayList

        '抽出1
        For r As Integer = start To dgv.RowCount - 1
            If dgv.Rows(r).IsNewRow Then
                Continue For
            End If

            Dim vFlag As Boolean = False
            If TextBox16.Text <> "" Then
                For i As Integer = 0 To searchA.Length - 1
                    If searchA(i) <> "" Then
                        If CheckBox18.Checked = True Then   'like
                            Dim regTxt As String = ".*" & searchA(i) & ".*"
                            Select Case ComboBox1.SelectedItem
                                Case "前方一致"
                                    regTxt = "^" & searchA(i)
                                Case "後方一致"
                                    regTxt = searchA(i) & "$"
                                Case "完全一致"
                                    regTxt = "^" & searchA(i) & "$"
                                Case "部分一致"

                                Case Else

                            End Select
                            If CheckBox19.Checked = False Then
                                If Regex.IsMatch(dgv.Item(colA, r).Value, regTxt) Then
                                    vFlag = True
                                    Exit For
                                End If
                            Else
                                If Not Regex.IsMatch(dgv.Item(colA, r).Value, regTxt) Then
                                    vFlag = True
                                    Exit For
                                End If
                            End If
                        Else
                            If CheckBox19.Checked = False Then  'not
                                If Regex.IsMatch(dgv.Item(colA, r).Value, "^" & searchA(i) & "$") Then
                                    vFlag = True
                                    Exit For
                                End If
                            Else
                                If Not Regex.IsMatch(dgv.Item(colA, r).Value, "^" & searchA(i) & "$") Then
                                    vFlag = True
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next
            Else
                vFlag = True
            End If

            '文字数抽出
            If CheckBox1.Checked And vFlag Then
                Dim strLen As Integer = dgv.Item(colA, r).Value.ToString.Length
                If strLen >= NumericUpDown1.Value And strLen <= NumericUpDown2.Value Then
                    If Not CheckBox19.Checked Then
                        vFlag = True
                    Else
                        vFlag = False
                    End If
                Else
                    If Not CheckBox19.Checked Then
                        vFlag = False
                    Else
                        vFlag = True
                    End If
                End If
            End If

            If vFlag Then
                rowArray.Add(r)
                dgv.Rows(r).Visible = True
            Else
                dgv.Rows(r).Visible = False
            End If
        Next

        '抽出2
        If CheckBox4.Checked = True Then
            For k As Integer = 0 To rowArray.Count - 1
                If dgv.Rows(rowArray(k)).Visible = True Then
                    For i As Integer = 0 To searchB.Length - 1
                        If searchB(i) <> "" Then
                            If CheckBox18.Checked = True Then
                                If Not Regex.IsMatch(dgv.Item(colB, rowArray(k)).Value, ".*" & searchB(i) & ".*") Then
                                    dgv.Rows(rowArray(k)).Visible = False
                                    Exit For
                                End If
                            Else
                                If Not Regex.IsMatch(dgv.Item(colB, rowArray(k)).Value, "^" & searchB(i) & "$") Then
                                    dgv.Rows(rowArray(k)).Visible = False
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                End If
            Next
        End If

        '抽出3
        If CheckBox6.Checked = True Then
            For k As Integer = 0 To rowArray.Count - 1
                If dgv.Rows(rowArray(k)).Visible = True Then
                    For i As Integer = 0 To searchC.Length - 1
                        If searchC(i) <> "" Then
                            If CheckBox18.Checked = True Then
                                If Not Regex.IsMatch(dgv.Item(colC, rowArray(k)).Value, ".*" & searchC(i) & ".*") Then
                                    dgv.Rows(rowArray(k)).Visible = False
                                    Exit For
                                End If
                            Else
                                If Not Regex.IsMatch(dgv.Item(colC, rowArray(k)).Value, "^" & searchC(i) & "$") Then
                                    dgv.Rows(rowArray(k)).Visible = False
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                End If
            Next
        ElseIf CheckBox7.Checked = True Then
            For k As Integer = 0 To rowArray.Count - 1
                If dgv.Rows(rowArray(k)).Visible = False Then
                    For i As Integer = 0 To searchC.Length - 1
                        If searchC(i) <> "" Then
                            If CheckBox18.Checked = True Then
                                If Regex.IsMatch(dgv.Item(colC, rowArray(k)).Value, ".*" & searchC(i) & ".*") Then
                                    dgv.Rows(rowArray(k)).Visible = True
                                    Exit For
                                End If
                            Else
                                If Regex.IsMatch(dgv.Item(colC, rowArray(k)).Value, "^" & searchC(i) & "$") Then
                                    dgv.Rows(rowArray(k)).Visible = True
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                End If
            Next
        End If

        If ComboBox10.SelectedItem = "1:CSV1" Then
            Form1.CSV_FORMS(0).Panel2.Visible = False
        ElseIf ComboBox10.SelectedItem = "2:CSV2" Then
            Form1.CSV_FORMS(1).Panel2.Visible = False
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Me.Hide()
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        If ComboBox10.SelectedItem = "1:CSV1" Then
            Form1.CSV_FORMS(0).抽出リストリセットToolStripMenuItem.PerformClick()
        ElseIf ComboBox10.SelectedItem = "2:CSV2" Then
            Form1.CSV_FORMS(1).抽出リストリセットToolStripMenuItem.PerformClick()
        End If
    End Sub

    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        If AzukiControl7.Text <> "" Then
            Dim FileName As String = ToolStripStatusLabel7.Text
            SaveCsv(FileName, ToolStripStatusLabel6.Text)
            'MsgBox("上書き保存しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    Public Sub SaveCsv2(ByVal fp As String, ByVal enc As String)
        If fp = "" Then
            Dim sfd As New SaveFileDialog()
            If ToolStripStatusLabel7.Text = "..." Then
                MsgBox("上書き保存できません", MsgBoxStyle.SystemModal)
                Exit Sub
            Else
                sfd.FileName = Path.GetFileName(ToolStripStatusLabel7.Text)
            End If
            sfd.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
            sfd.Filter = "HTMLファイル(*.html)|*.html|TXTファイル(*.txt)|*.txt|すべてのファイル(*.*)|*.*"
            sfd.FilterIndex = 0
            sfd.Title = "保存先のファイルを選択してください"
            sfd.RestoreDirectory = True
            sfd.OverwritePrompt = True
            sfd.CheckPathExists = True

            'ダイアログを表示する
            If sfd.ShowDialog(Me) = DialogResult.OK Then
                'SaveCsv(sfd.FileName, ToolStripStatusLabel5.Text)
                ' CSVファイルオープン
                Dim en As EncodingJP = DecodingName(enc)
                TextFileEncoding.SaveToFile(sfd.FileName, AzukiControl7.Text, en)

                ToolStripStatusLabel7.Text = sfd.FileName
                MsgBox(sfd.FileName & vbCrLf & "保存しました", MsgBoxStyle.SystemModal)
            End If
        Else
            Dim en As EncodingJP = DecodingName(enc)
            TextFileEncoding.SaveToFile(fp, AzukiControl7.Text, en)
            MsgBox("上書き保存しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        For i As Integer = 0 To DataGridView4.RowCount - 2
            Dim str As String = DataGridView4.Item(0, i).Value
            Dim num As String = String.Format("{0:D2}", i + 1)
            AzukiControl7.Text = Replace(AzukiControl7.Text, "$" & num, str)
        Next
    End Sub

    Private Sub 貼り付けToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 貼り付けToolStripMenuItem.Click
        Dim x As Integer = DataGridView4.CurrentCellAddress.X
        Dim y As Integer = DataGridView4.CurrentCellAddress.Y
        Dim clipText As String = Clipboard.GetText()
        clipText = clipText.Replace(vbCrLf, vbLf)
        clipText = clipText.Replace(vbCr, vbLf)
        Dim lines() As String = clipText.Split(vbLf)

        For i As Integer = 0 To lines.Length - 1
            If lines(i) <> "" Then
                If y + i >= DataGridView4.RowCount - 1 Then
                    DataGridView4.Rows.Add(1)
                End If
                DataGridView4.Item(x, y + i).Value = lines(i)
            End If
        Next
    End Sub

    Private Sub 行を削除ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 行を削除ToolStripMenuItem.Click
        Dim rArray As ArrayList = New ArrayList
        For i As Integer = 0 To DataGridView4.SelectedCells.Count - 1
            Dim selR As Integer = DataGridView4.SelectedCells(i).RowIndex
            If rArray.Contains(selR) = False Then
                rArray.Add(selR)
            End If
        Next
        rArray.Sort()
        rArray.Reverse()
        For i As Integer = 0 To rArray.Count - 1
            DataGridView4.Rows.RemoveAt(rArray(i))
        Next
    End Sub

    Private Sub リストクリアToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles リストクリアToolStripMenuItem.Click
        DataGridView4.Rows.Clear()
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        '*********************************************
        Form1.inif.WriteINI(secValue, "tenpoCheck1", RadioButton1.Checked)
        '*********************************************
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        '*********************************************
        Form1.inif.WriteINI(secValue, "tenpoCheck2", RadioButton2.Checked)
        '*********************************************
    End Sub

    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        '*********************************************
        Form1.inif.WriteINI(secValue, "tenpoCheck3", RadioButton3.Checked)
        '*********************************************
    End Sub

    Private Sub RadioButton4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton4.CheckedChanged
        '*********************************************
        Form1.inif.WriteINI(secValue, "tenpoCheck4", RadioButton4.Checked)
        '*********************************************
    End Sub

    Private Sub RadioButton8_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton8.CheckedChanged
        '*********************************************
        Form1.inif.WriteINI(secValue, "tenpoCheck8", RadioButton8.Checked)
        '*********************************************
    End Sub

    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        Dim fbd As New FolderBrowserDialog With {
            .Description = "フォルダを指定してください。",
            .RootFolder = Environment.SpecialFolder.Desktop,
            .SelectedPath = "",
            .ShowNewFolderButton = True
        }

        If fbd.ShowDialog(Me) = DialogResult.OK Then
            Dim di As New System.IO.DirectoryInfo(fbd.SelectedPath)
            TextBox28.Text = di.FullName
        End If

    End Sub

    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button25.Click
        If TextBox29.Text <> "" Then
            Dim di As New System.IO.DirectoryInfo(TextBox28.Text)
            Dim files As System.IO.FileInfo() = di.GetFiles("*.*", System.IO.SearchOption.AllDirectories)
            For Each f As System.IO.FileInfo In files
                Dim fname As String = System.IO.Path.GetFileNameWithoutExtension(f.Name)
                If InStr(fname, TextBox29.Text) > 0 Then
                    Dim dirname As String = System.IO.Path.GetDirectoryName(f.FullName)
                    Dim extname As String = System.IO.Path.GetExtension(f.FullName)
                    Dim nName As String = Replace(fname, TextBox29.Text, TextBox30.Text)
                    f.MoveTo(dirname & "\" & nName & extname)
                End If
            Next
            MsgBox("変更しました", MsgBoxStyle.SystemModal)
        Else
            MsgBox("置換前の文字列が入力されていません", MsgBoxStyle.SystemModal)
        End If
    End Sub

    Private Sub TextBox16_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox16.KeyDown
        If e.Modifiers = Keys.Control And e.KeyCode = Keys.A Then
            TextBox16.SelectAll()
        End If
    End Sub

    Private Sub TextBox17_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox17.KeyDown
        If e.Modifiers = Keys.Control And e.KeyCode = Keys.A Then
            TextBox17.SelectAll()
        End If
    End Sub

    Private Sub TextBox18_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox18.KeyDown
        If e.Modifiers = Keys.Control And e.KeyCode = Keys.A Then
            TextBox18.SelectAll()
        End If
    End Sub

    '価格変更
    Private Sub Button26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button26.Click
        Dim str As String = ""
        Dim lines As String() = Regex.Split(AzukiControl6.Text, "\n|\r|\n\r")
        For i As Integer = 0 To lines.Length - 1
            Dim line As String() = Split(lines(i), ",")
            If line.Length > 2 Then
                If line(0) = TextBox34.Text Then
                    Dim res As String = ""
                    For k As Integer = 0 To line.Length - 1
                        If res = "" Then
                            res = line(k)
                        Else
                            If k = 2 Then
                                res &= "," & TextBox37.Text
                            Else
                                res &= "," & line(k)
                            End If
                        End If
                    Next
                    res &= vbCrLf
                    str &= res
                Else
                    str &= lines(i) & vbCrLf
                End If
            End If
        Next

        AzukiControl6.Text = str
        MsgBox("修正しました", MsgBoxStyle.SystemModal)
    End Sub


    Private Sub Button27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button27.Click
        Dim ofd As New OpenFileDialog With {
            .FileName = "default.html",
            .InitialDirectory = Environment.SpecialFolder.DesktopDirectory,
            .Filter = "CSVファイル(*.csv)|*.csv|すべてのファイル(*.*)|*.*",
            .FilterIndex = 2,
            .Title = "開くファイルを選択してください",
            .RestoreDirectory = True,
            .CheckFileExists = True,
            .CheckPathExists = True
        }

        'ダイアログを表示する
        If ofd.ShowDialog() = DialogResult.OK Then
            Dim fName As String = ofd.FileName
            Dim itemList_web As ArrayList = CSV_READ(fName)

            'ヘッダーテキストをコレクションに取り込む
            Dim headerArray As String() = Split(itemList_web(0), "|=|")
            Dim DH1 As ArrayList = New ArrayList
            For c As Integer = 0 To headerArray.Length - 1
                DH1.Add(headerArray(c))
            Next c

            'azukitextをリスト化
            Dim lines As String = AzukiControl6.Text
            lines = Replace(lines, vbCrLf, vbLf)
            lines = Replace(lines, vbCr, vbLf)
            Dim lineAZ As String() = Split(lines, vbLf)
            AzukiControl6.Text = ""
            Application.DoEvents()

            Dim header_kakaku As String = "販売価格"
            Dim header_souko As String = "倉庫指定"
            Dim header_img As String = "商品画像URL"
            Dim codeNum As Integer = 1

            Dim url As String = ""
            Dim url2 As String = ""
            If RadioButton1.Checked Then
                url = "https://item.rakuten.co.jp/patri/"
            ElseIf RadioButton4.Checked Then
                url = "https://item.rakuten.co.jp/ashop/"
            ElseIf RadioButton8.Checked Then
                url = "https://store.shopping.yahoo.co.jp/kt-zkshop/"
                url2 = "https://item-shopping.c.yimg.jp/i/f/kt-zkshop_"
                header_kakaku = "price"
                header_souko = "display"
                header_img = "caption"
                codeNum = 2
            ElseIf RadioButton3.Checked Then    'wowma
                'wowmaのリンクアドレス＝＞https://wowma.jp/item/337207655?l=true&e=llA&e2=listing_flpro
                url = "https://wowma.jp/item/"
                url2 = "?l=true&e=llA&e2=listing_flpro"
                header_kakaku = "itemPrice"
                header_souko = "saleStatus"
                header_img = "imageUrl1"
                codeNum = 2
            Else
                url = "https://item.rakuten.co.jp/alice-zk/"
            End If

            'csvの商品コードをarraylistに入れる
            Dim csvCodeArray As ArrayList = New ArrayList
            For i As Integer = 0 To itemList_web.Count - 1
                Dim list As String() = Split(itemList_web(i), "|=|")
                csvCodeArray.Add(list(codeNum).ToLower)
            Next

            ToolStripProgressBar1.Value = 0
            ToolStripProgressBar1.Maximum = lineAZ.Length

            Dim DRstr As String = ""

            Dim str As String = ""
            For r As Integer = 0 To lineAZ.Length - 1
                Dim oneStr As String = ""
                If lineAZ(r) <> "" Then
                    Dim lineSTR As String() = Split(lineAZ(r), ",") '商品コード,商品名,価格,商品URL,画像URL,更新日,カテゴリ,倉庫
                    lineSTR(0) = lineSTR(0).ToLower
                    If csvCodeArray.Contains(lineSTR(0)) Then   'csvにある時
                        Dim csvNum As Integer = csvCodeArray.IndexOf(lineSTR(0))
                        Dim csvStr As String() = Split(itemList_web(csvNum), "|=|")

                        '商品コード
                        oneStr = lineSTR(0)

                        '名前
                        oneStr &= "," & lineSTR(1)

                        '価格
                        If csvStr(DH1.IndexOf(header_kakaku)) <> lineSTR(2) Then
                            DRstr &= "「" & lineSTR(0) & "」 " & lineSTR(2) & "→" & csvStr(DH1.IndexOf(header_kakaku)) & " 変更" & vbCrLf
                        End If
                        oneStr &= "," & csvStr(DH1.IndexOf(header_kakaku))

                        'アドレス
                        If header_souko = "saleStatus" Then     'wowma用
                            oneStr &= "," & url & csvStr(DH1.IndexOf("lotNumber")) & url2
                        Else
                            oneStr &= "," & url & lineSTR(0)
                        End If

                        '画像
                        If header_img = "商品画像URL" Then
                            If RadioButton4.Checked Then    '楽天あかねの画像は全て「mini」
                                oneStr &= "," & "https://image.rakuten.co.jp/ashop/cabinet/mini/" & lineSTR(0) & ".jpg"
                                'ElseIf RadioButton1.Checked Then    '暁の画像は全て「画像名＋a」
                                '    Dim imgArray As String() = Split(csvStr(DH1.IndexOf(header_img)), " ")
                                '    imgArray(0) = Replace(imgArray(0), ".jpg", "a.jpg")
                                '    imgArray(0) = Replace(imgArray(0), ".gif", "a.gif")
                                '    oneStr &= "," & imgArray(0)
                            Else
                                Dim imgArray As String() = Split(csvStr(DH1.IndexOf(header_img)), " ")
                                oneStr &= "," & imgArray(0)
                            End If
                        ElseIf header_img = "imageUrl1" Then    'wowma用
                            oneStr &= "," & csvStr(DH1.IndexOf(header_img))
                        Else    'Yahoo用
                            oneStr &= "," & url2 & csvStr(DH1.IndexOf("code"))
                        End If

                        '登録日
                        oneStr &= "," & lineSTR(5)

                        'カテゴリ
                        If Not RadioButton3.Checked Then   'wowmaはカテゴリ入れない
                            oneStr &= "," & lineSTR(6)
                        End If

                        '倉庫
                        oneStr &= "," & csvStr(DH1.IndexOf(header_souko))

                        If RadioButton3.Checked And csvStr(DH1.IndexOf(header_souko)) <> "2" Then
                            'wowmaは非表示商品はリストに入れない
                            str &= oneStr & vbLf
                        Else
                            str &= oneStr & vbLf
                        End If

                        ToolStripStatusLabel1.Text = csvNum

                        'csvリストから削除
                        csvCodeArray.RemoveAt(csvNum)
                        itemList_web.RemoveAt(csvNum)
                    Else
                        'csvに存在しないので削除対象
                        DRstr &= "「" & lineSTR(0) & "」削除" & vbCrLf
                    End If
                End If
            Next

            Dim newStr As String = ""
            If header_kakaku = "販売価格" Then
                For i As Integer = 0 To itemList_web.Count - 1
                    If InStr(itemList_web(i), "コントロールカラム") = 0 Then    '1行目は読み込まない
                        Dim oneStr As String = ""
                        Dim code As String = ""
                        Dim csvStr As String() = Split(itemList_web(i), "|=|")
                        oneStr = csvStr(DH1.IndexOf("商品管理番号（商品URL）"))
                        code = oneStr
                        Dim sName As String = csvStr(DH1.IndexOf("モバイル用キャッチコピー"))
                        sName = Regex.Replace(sName, "\n|\r|\r\n|★.*★|【.*】| |　|[a-zA-Z]{1,2}[0-9]{3}", "")
                        oneStr &= "," & sName
                        oneStr &= "," & csvStr(DH1.IndexOf("販売価格"))
                        oneStr &= "," & url & csvStr(DH1.IndexOf("商品管理番号（商品URL）"))
                        If RadioButton4.Checked Then    '楽天あかねの画像は全て「mini」
                            oneStr &= "," & "https://image.rakuten.co.jp/ashop/cabinet/mini/" & csvStr(DH1.IndexOf("商品管理番号（商品URL）")) & ".jpg"
                            'ElseIf RadioButton1.Checked Then
                            '    Dim imgArray As String() = Split(csvStr(DH1.IndexOf("商品画像URL")), " ")
                            '    imgArray(0) = Replace(imgArray(0), ".jpg", "a.jpg")
                            '    imgArray(0) = Replace(imgArray(0), ".gif", "a.gif")
                            '    oneStr &= "," & imgArray(0)
                        Else
                            Dim imgArray As String() = Split(csvStr(DH1.IndexOf("商品画像URL")), " ")
                            oneStr &= "," & imgArray(0)
                        End If
                        If CheckBox20.Checked = True Then
                            oneStr &= "," & Format(Now, "yyyyMMdd")
                        Else
                            oneStr &= ",0"
                        End If
                        oneStr &= "," & ""  'カテゴリーは別ファイルで登録
                        oneStr &= "," & csvStr(DH1.IndexOf(header_souko))
                        newStr &= oneStr & vbLf

                        DRstr &= "「" & code & "」追加" & vbCrLf
                    End If
                Next
            ElseIf header_kakaku = "itemPrice" Then             'wowma用
                For i As Integer = 0 To itemList_web.Count - 1
                    If InStr(itemList_web(i), "lotNumber") = 0 Then    '1行目は読み込まない
                        Dim oneStr As String = ""
                        Dim code As String = ""
                        Dim lotnum As String = ""
                        Dim csvStr As String() = Split(itemList_web(i), "|=|")
                        oneStr = csvStr(DH1.IndexOf("itemCode"))
                        code = oneStr
                        Dim sName As String = csvStr(DH1.IndexOf("itemName"))
                        sName = Regex.Replace(sName, "\n|\r|\r\n|★.*★|【.*】|[a-zA-Z]{1,2}[0-9]{3}", "")
                        Dim ssName As String = Split(sName, " ")(0)
                        oneStr &= "," & ssName
                        oneStr &= "," & csvStr(DH1.IndexOf("itemPrice"))
                        lotnum = csvStr(DH1.IndexOf("lotNumber"))
                        oneStr &= "," & url & lotnum & url2
                        oneStr &= "," & csvStr(DH1.IndexOf("imageUrl1"))
                        If CheckBox20.Checked = True Then
                            oneStr &= "," & Format(Now, "yyyyMMdd")
                        Else
                            oneStr &= ",0"
                        End If
                        oneStr &= "," & csvStr(DH1.IndexOf(header_souko))

                        If csvStr(DH1.IndexOf(header_souko)) <> "2" Then
                            newStr &= oneStr & vbLf
                            DRstr &= "「" & code & "」追加" & vbCrLf
                        End If
                    End If
                Next
            Else
                For i As Integer = 0 To itemList_web.Count - 1
                    If InStr(itemList_web(i), "path") = 0 Then    '1行目は読み込まない
                        Dim oneStr As String = ""
                        Dim code As String = ""
                        Dim csvStr As String() = Split(itemList_web(i), "|=|")
                        oneStr = csvStr(DH1.IndexOf("code"))
                        code = oneStr
                        Dim sName As String = csvStr(DH1.IndexOf("headline"))
                        sName = Regex.Replace(sName, "\n|\r|\r\n|★.*★|【.*】| |　|[a-zA-Z]{1,2}[0-9]{3}", "")
                        oneStr &= "," & sName
                        oneStr &= "," & csvStr(DH1.IndexOf("price"))
                        oneStr &= "," & url & csvStr(DH1.IndexOf("code"))
                        oneStr &= "," & url2 & csvStr(DH1.IndexOf("code"))
                        If CheckBox20.Checked = True Then
                            oneStr &= "," & Format(Now, "yyyyMMdd")
                        Else
                            oneStr &= ",0"
                        End If
                        Dim cate As String = csvStr(DH1.IndexOf("path"))
                        cate = Regex.Replace(cate, "\n|\r|\n\r", "|")
                        oneStr &= "," & cate  'カテゴリーは別ファイルで登録
                        oneStr &= "," & csvStr(DH1.IndexOf(header_souko))
                        newStr &= oneStr & vbLf

                        DRstr &= "「" & code & "」追加" & vbCrLf
                    End If
                Next
            End If

            If CheckBox20.Checked = True Then
                AzukiControl6.Text = newStr & str
            Else
                AzukiControl6.Text = str & newStr
            End If

            If DRstr = "" Then
                DRstr = "価格更新、商品削除・追加はありません。" & vbCrLf & "画像・商品URLは更新しました。"
            Else
                Dim DRlines As String() = Split(DRstr, vbCrLf)
                If DRlines.Length > 10 Then
                    For i As Integer = 0 To 10
                        DRstr &= DRlines(i) & vbCrLf
                    Next
                    DRstr &= "..."
                End If
                MsgBox(DRstr, MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
            End If
        End If
    End Sub

    Private Sub HTMLdialog_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        If Form1.CSV_FORMS(0).ToolStripStatusLabel5.Text = "1" Then
            ComboBox10.SelectedItem = "1:CSV1"
        ElseIf HTMLcreate.Label1.Text = "H" Then
            ComboBox10.SelectedItem = "1:CSV1"
        Else
            ComboBox10.SelectedItem = "2:CSV2"
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            NumericUpDown1.Enabled = True
            NumericUpDown2.Enabled = True
        Else
            NumericUpDown1.Enabled = False
            NumericUpDown2.Enabled = False
        End If
    End Sub

    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged
        If NumericUpDown2.Value < NumericUpDown1.Value Then
            NumericUpDown2.Value = NumericUpDown1.Value
        End If
    End Sub

    Private Sub NumericUpDown2_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown2.ValueChanged
        If NumericUpDown1.Value > NumericUpDown2.Value Then
            NumericUpDown1.Value = NumericUpDown2.Value
        End If
    End Sub

    Private Sub TextBox41_TextChanged(sender As Object, e As EventArgs) Handles TextBox41.TextChanged
        PriceCalc3("定価")
    End Sub

    Private Sub TextBox42_TextChanged(sender As Object, e As EventArgs) Handles TextBox42.TextChanged
        PriceCalc3("販売")
    End Sub

    Private Sub NumericUpDown8_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown8.ValueChanged
        PriceCalc3("税率")
    End Sub

    Private Sub NumericUpDown11_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown11.ValueChanged
        PriceCalc3("")
    End Sub

    Private Sub NumericUpDown12_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown12.ValueChanged
        PriceCalc3("")
    End Sub

    Private Sub NumericUpDown13_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown13.ValueChanged
        PriceCalc3("")
    End Sub

    Private Sub PriceCalc3(ByVal name As String)
        Dim action As Boolean = True

        Dim teika As Integer = 0
        If IsNumeric(TextBox41.Text) Then
            teika = TextBox41.Text
        Else
            action = False
        End If

        Dim hanbai As Integer = 0
        If IsNumeric(TextBox42.Text) Then
            hanbai = TextBox42.Text
        Else
            action = False
        End If

        Dim rate As Integer = NumericUpDown8.Value

        Select Case name
            Case "定価"
                hanbai = Math.Round(CInt(teika) / (1 + rate / 100) + 0.00000001)
                TextBox42.Text = hanbai
            Case "税率"
                'If IsNumeric(TextBox42.Text) Then
                '    teika = Math.Round(CInt(hanbai) * (1 + rate / 100) + 0.00000001)
                '   TextBox41.Text = teika
                'ElseIf IsNumeric(TextBox41.Text) Then
                '    teika = TextBox41.Text
                '    hanbai = Math.Round(CInt(teika) / (1 + rate / 100) + 0.00000001)
                '    TextBox42.Text = hanbai
                'Else
                '        TextBox41.Text = ""
                '    TextBox42.Text = ""
                '    Exit Sub
                'End If
                teika = Math.Round(CInt(hanbai) * (1 + rate / 100) + 0.00000001)
                TextBox41.Text = teika
            Case "販売"
                teika = Math.Round(CInt(hanbai) * (1 + rate / 100) + 0.00000001)
                TextBox41.Text = teika
        End Select

        If teika = 0 Or action = False Then
            For r As Integer = 0 To DataGridView2.RowCount - 1
                Select Case r
                    Case 1
                        DataGridView2.Item(1, r).Value = ""
                    Case 2
                        DataGridView2.Item(1, r).Value = ""
                    Case 3
                        DataGridView2.Item(1, r).Value = ""
                    Case 4
                        DataGridView2.Item(1, r).Value = ""
                    Case 5
                        DataGridView2.Item(1, r).Value = ""
                End Select
            Next
        Else
            Dim vip_price As Integer = 0
            For r As Integer = 0 To DataGridView2.RowCount - 1
                Select Case r
                    Case 0
                        If teika.ToString <> "" Then
                            DataGridView2.Item(1, r).Value = teika
                        Else
                            DataGridView2.Item(1, r).Value = ""
                        End If
                    Case 1
                        DataGridView2.Item(1, r).Value = hanbai
                    Case 2
                        vip_price = hanbai * NumericUpDown11.Value / 100
                        DataGridView2.Item(1, r).Value = vip_price
                    Case 3
                        vip_price = hanbai * NumericUpDown12.Value / 100
                        DataGridView2.Item(1, r).Value = vip_price
                    Case 4
                        vip_price = hanbai * NumericUpDown13.Value / 100
                        DataGridView2.Item(1, r).Value = vip_price
                End Select
            Next
        End If
    End Sub

    Private Sub keisan()
        Dim genka As Integer = NumericUpDown3.Value
        Dim baika As Integer = NumericUpDown4.Value
        Dim kakakuritsu1 As Integer = NumericUpDown5.Value
        Dim kakakuritsu2 As Integer = NumericUpDown6.Value
        Dim kakakuritsu3 As Integer = NumericUpDown7.Value

        'Dim dH3 As ArrayList = TM_HEADER_GET(DataGridView3)
        'DataGridView3.Rows.Clear()
        'DataGridView3.Rows.Add()

        Dim dt As DataTable = New DataTable()

        With dt.Columns
            .Add("原価")
            .Add("売価")
            .Add("価格1")
            .Add("価格2")
            .Add("価格3")
            .Add("利益")
            .Add("利益率")
        End With

        Dim kakaku1 As Integer = baika * kakakuritsu1 / 100
        Dim kakaku2 As Integer = baika * kakakuritsu2 / 100
        Dim kakaku3 As Integer = baika * kakakuritsu3 / 100

        Dim rieki1 As Integer = kakaku1 - genka
        Dim rieki2 As Integer = kakaku2 - genka
        Dim rieki3 As Integer = kakaku3 - genka

        dt.Rows.Add(genka, baika, kakaku1, kakaku2, kakaku3, rieki1 & " (価格1)", Math.Floor((rieki1 / genka) * 10000) / 100.0 & "% (価格1)")
        dt.Rows.Add(genka, baika, kakaku1, kakaku2, kakaku3, rieki2 & " (価格2)", Math.Floor((rieki2 / genka) * 10000) / 100.0 & "% (価格2)")
        dt.Rows.Add(genka, baika, kakaku1, kakaku2, kakaku3, rieki3 & " (価格3)", Math.Floor((rieki3 / genka) * 10000) / 100.0 & "% (価格3)")

        'DataGridView3.Item(dH3.IndexOf("原価"), 0).Value = genka
        'DataGridView3.Item(dH3.IndexOf("売価"), 0).Value = baika
        'DataGridView3.Item(dH3.IndexOf("価格1"), 0).Value = CInt(genka * kakakuritsu1 / 100)
        'DataGridView3.Item(dH3.IndexOf("価格2"), 0).Value = CInt(genka * kakakuritsu2 / 100)
        'DataGridView3.Item(dH3.IndexOf("価格3"), 0).Value = CInt(genka * kakakuritsu3 / 100)
        DataGridView3.DataSource = dt
    End Sub

    Private Sub NumericUpDown3_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown3.ValueChanged
        keisan()
    End Sub

    Private Sub NumericUpDown3_KeyUp(sender As Object, e As KeyEventArgs) Handles NumericUpDown3.KeyUp
        keisan()
    End Sub

    Private Sub NumericUpDown4_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown4.ValueChanged
        keisan()
    End Sub

    Private Sub NumericUpDown4_KeyUp(sender As Object, e As KeyEventArgs) Handles NumericUpDown4.KeyUp
        keisan()
    End Sub

    Private Sub NumericUpDown5_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown5.ValueChanged
        keisan()
    End Sub

    Private Sub NumericUpDown5_KeyUp(sender As Object, e As KeyEventArgs) Handles NumericUpDown5.KeyUp
        keisan()
    End Sub

    Private Sub NumericUpDown6_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown6.ValueChanged
        keisan()
    End Sub

    Private Sub NumericUpDown6_KeyUp(sender As Object, e As KeyEventArgs) Handles NumericUpDown6.KeyUp
        keisan()
    End Sub

    Private Sub NumericUpDown7_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown7.ValueChanged
        keisan()
    End Sub

    Private Sub NumericUpDown7_KeyUp(sender As Object, e As KeyEventArgs) Handles NumericUpDown7.KeyUp
        keisan()
    End Sub

    ' 指定したセルと1つ上のセルの値を比較
    Function IsTheSameCellValue(ByVal column As Integer, ByVal row As Integer, dgv As DataGridView) As Boolean

        Dim cell1 As DataGridViewCell = dgv(column, row)
        Dim cell2 As DataGridViewCell = dgv(column, row - 1)

        If cell1.Value = Nothing Or cell2.Value = Nothing Then
            Return False
        End If

        ' ここでは文字列としてセルの値を比較
        If cell1.Value.ToString() = cell2.Value.ToString() Then
            Return True
        Else
            Return False
        End If

    End Function

    ' DataGridViewのCellFormattingイベント・ハンドラ
    Private Sub DataGridView3_CellFormatting_1(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView3.CellFormatting
        ' 1行目については何もしない
        If e.RowIndex = 0 Then
            Return
        End If

        If IsTheSameCellValue(e.ColumnIndex, e.RowIndex, DataGridView3) Then
            e.Value = ""
            e.FormattingApplied = True ' 以降の書式設定は不要
        End If
    End Sub

    Private Sub DataGridView3_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs) Handles DataGridView3.CellPainting
        ' セルの下側の境界線を「境界線なし」に設定
        e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None

        ' 1行目や列ヘッダ、行ヘッダの場合は何もしない
        If e.RowIndex < 1 Or e.ColumnIndex < 0 Then
            Return
        End If

        If IsTheSameCellValue(e.ColumnIndex, e.RowIndex, DataGridView3) Then
            ' セルの上側の境界線を「境界線なし」に設定
            e.AdvancedBorderStyle.Top = DataGridViewAdvancedCellBorderStyle.None
        Else
            ' セルの上側の境界線を既定の境界線に設定
            e.AdvancedBorderStyle.Top = DataGridView3.AdvancedCellBorderStyle.Top
        End If
    End Sub
End Class