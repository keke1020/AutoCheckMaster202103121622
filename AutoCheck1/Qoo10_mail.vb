Imports System.Net
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports OpenQA.Selenium.Chrome
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Support
Imports System.Collections.ObjectModel
Imports OpenQA.Selenium.Interactions
Imports System.Threading
Imports Hnx8

Public Class Qoo10_mail
    Dim ENC_SJ As Encoding = Encoding.GetEncoding("SHIFT-JIS")
    Dim appPath As String = Path.GetDirectoryName(Application.ExecutablePath)
    Dim bw1_url As String = ""
    Dim wb_loading_ng As Boolean = True
    Dim template_mail As String = ""
    Dim splitter As String = ","    '区切り文字

    Private Delegate Sub CallDelegate()

    Private cr As ChromeDriver = Nothing
    Private sizeW As Integer = 500
    Private sizeH As Integer = 800

    Private Sub Qoo10_mail_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'template_mail = File.ReadAllText(appPath & "\config\template_Qoo10.txt", ENC_SJ)
        'TextBox2.Text = template_mail
        ToolStripTextBox1.Text = ""
    End Sub

    Dim td1 As Thread
    Dim td1_action As Boolean = False
    Dim noArray As New ArrayList
    Private Sub startThread()
        Dim listArray As String() = Split(TextBox1.Text, vbCrLf)
        Dim list_str As String = ""
        td1_action = True
        noArray.Clear()

        Array.Sort(listArray)
        For i As Integer = 0 To listArray.Count - 1
            If listArray(i) <> "" Then
                If listArray(i) <> Int(listArray(i)) Then
                    MsgBox("注文番号を数字で入力してください。", MsgBoxStyle.SystemModal)
                    Exit Sub
                End If
                list_str &= listArray(i) & vbCrLf
            End If
        Next

        '重複を取り除く
        Dim lines As String() = Split(list_str, vbCrLf)
        For i As Integer = 0 To lines.Length - 1
            If lines(i) = "" Then
                Continue For
            End If

            If Not noArray.Contains(lines(i)) Then
                noArray.Add(lines(i))
            End If
        Next

        If noArray.Count = 0 Then
            MsgBox("注文番号を入力してください。", MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        If cr Is Nothing Then
            ChromeStart()
        End If

        Invoke(New CallDelegate(AddressOf action1))

        td1.Abort()
    End Sub

    Private Sub action1()
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()

        'Dim title_arr As String() = "店舗伝票番号,受注日,受注郵便番号,受注住所１,受注住所２,受注名,受注名カナ,受注電話番号,受注メールアドレス,発送郵便番号,発送先住所１,発送先住所２,発送先名,発送先カナ,発送電話番号,支払方法,発送方法,商品計,税金,発送料,手数料,ポイント,その他費用,合計金額,ギフトフラグ,時間帯指定,日付指定,作業者欄,備考,商品名,商品コード,商品価格,受注数量,商品オプション,出荷済フラグ,顧客区分,顧客コード,消費税率（%）".Split(",")
        Dim title_arr As String() = "注文番号,結果".Split(",")

        For c As Integer = 0 To title_arr.Length - 1
            DataGridView1.Columns.Add(c, title_arr(c))
        Next

        Dim add_count As Integer = 0
        For i As Integer = 0 To noArray.Count - 1
            If td1_action = False Then
                Exit Sub
            End If

            Dim no As String = noArray(i)
            ToolStripTextBox1.Text = no
            Application.DoEvents()

            If IshaveDenpyo_ne_cr(no) Then
                DataGridView1.Rows.Add(1)
                DataGridView1.Item(0, i).Value = no
                DataGridView1.Item(1, i).Value = "○"
                DataGridView1.Item(0, i).Style.BackColor = Color.LightGreen
                Application.DoEvents()
            Else
                DataGridView1.Rows.Add(1)
                DataGridView1.Item(0, i).Value = no
                DataGridView1.Item(1, i).Value = "△"
                DataGridView1.Item(0, i).Style.BackColor = Color.Pink
                Application.DoEvents()
            End If
        Next
    End Sub

    '共通
    Private Function IsNetworkAvailable()
        Dim IsNetworkAvailable_Flag As Boolean = False
        If NetworkInformation.NetworkInterface.GetIsNetworkAvailable() Then
            ToolStripStatusLabel5.Text = "接続"
            ToolStripStatusLabel5.BackColor = Color.LightGreen
            IsNetworkAvailable_Flag = True
        Else
            ToolStripStatusLabel5.Text = "未接続"
            ToolStripStatusLabel5.BackColor = Color.Transparent
        End If
        Return IsNetworkAvailable_Flag
    End Function

    Private Sub WaitWebBrowser1Completed()
        Do While wb_loading_ng
            If wb_loading_ng = False Then
                Exit Do
            Else
                Application.DoEvents()
            End If
        Loop
        wb_loading_ng = True
        Application.DoEvents()
    End Sub

    Private Sub WebBrowser1_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs)
        wb_loading_ng = False
        bw1_url = e.Url.ToString
        Application.DoEvents()
    End Sub

    Private Function IshaveDenpyo_ne_wb(no As String)
        'Dim Ishave As Boolean = False
        'Dim url As String = "https://ne81.next-engine.com/Userjyuchu/jyuchuInp?kensaku_denpyo_no=" & no & "&jyuchu_meisai_order=jyuchu_meisai_gyo"
        'WebBrowser1.Navigate(New Uri(url))
        'WaitWebBrowser1Completed()
        'Application.DoEvents()

        'If bw1_url = "https://base.next-engine.org/users/sign_in/" Then
        '    WebBrowser1.Document.GetElementById("user_login_code").SetAttribute("value", "person2@yongshun006.com")
        '    WebBrowser1.Document.GetElementById("user_password").SetAttribute("value", "au_15u55")
        '    WebBrowser1.Document.GetElementsByTagName("input").GetElementsByName("commit")(0).InvokeMember("Click")
        '    WaitWebBrowser1Completed()

        '    Dim elesels As HtmlElementCollection = WebBrowser1.Document.GetElementsByTagName("a")
        '    For Each elesel As HtmlElement In elesels
        '        If elesel.GetAttribute("href") = "https://base.next-engine.org/apps/launch/?id=65459" Then
        '            elesel.InvokeMember("click")
        '            WaitWebBrowser1Completed()
        '            Exit For
        '        End If
        '    Next

        '    WebBrowser1.Navigate("https://ne81.next-engine.com/Usertop")
        '    WaitWebBrowser1Completed()

        '    WebBrowser1.Navigate(url)
        '    WaitWebBrowser1Completed()
        'End If

        'If InStr(WebBrowser1.Document.GetElementById("jyuchu_msg").GetElementsByTagName("b")(0).InnerText.ToString, "伝票番号の検索結果がありました") = 1 Then
        '    Ishave = True
        'Else
        '    Ishave = False
        'End If

        'Return Ishave
    End Function

    Private Function IshaveDenpyo_ne_cr(no As String)
        Dim Ishave As Boolean = False
        cr.Url = "https://ne81.next-engine.com/Userjyuchu/jyuchuInp?kensaku_denpyo_no=" & no & "&jyuchu_meisai_order=jyuchu_meisai_gyo"
        Dim result As String = cr.FindElementById("jyuchu_msg").Text

        If InStr(result, "伝票番号の検索結果がありました") = 1 Then
            Ishave = True
        Else
            Ishave = False
        End If

        Return Ishave
    End Function

    'Private Sub WebBrowser1_NewWindow(sender As Object, e As System.ComponentModel.CancelEventArgs)
    '    e.Cancel = True
    '    Dim url As String = sender.url.ToString
    '    WebBrowser1.Navigate(url)
    'End Sub

    '行番号を表示する
    Private Sub DataGridView_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView1.RowPostPaint
        Dim rect As New Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, sender.RowHeadersWidth - 4, sender.Rows(e.RowIndex).Height)
        TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), sender.RowHeadersDefaultCellStyle.Font, rect, sender.RowHeadersDefaultCellStyle.ForeColor, TextFormatFlags.VerticalCenter Or TextFormatFlags.Right)
    End Sub

    Private Sub ToolStripSplitButton1_ButtonClick(sender As Object, e As EventArgs) Handles ToolStripSplitButton1.ButtonClick
        Me.Dispose()
        Dim frm As Form = New Qoo10_mail
        frm.Show()
    End Sub


    'Chrome起動
    Private Sub ChromeStart()
        'chromedriver.exeをkill
        Dim hProcesses As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName("chromedriver")
        For Each hProcess As System.Diagnostics.Process In hProcesses
            hProcess.Kill()
        Next hProcess

        If InStr(appPath, "server2") > 0 Then
            'chrome.exeをkill
            Dim hProcesses2 As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName("chrome")
            For Each hProcess As System.Diagnostics.Process In hProcesses2
                hProcess.Kill()
            Next hProcess
        End If

        'conhost.exeをkill
        Dim hProcesses3 As System.Diagnostics.Process() = System.Diagnostics.Process.GetProcessesByName("conhost")
        For Each hProcess As System.Diagnostics.Process In hProcesses3
            Try
                hProcess.Kill()
            Catch ex As Exception

            End Try
        Next hProcess

        '---------------------------------
        '以下を回避するために「TimeSpan.FromSeconds(180)」「ChromeOptions.AddArguments("no-sandbox")」を追加
        'OpenQA.Selenium.WebDriverException : The HTTP request to the remote WebDriver server for URL http: //localhost:49192/session/f93057c8d4833f6a91026e82d339776f/url timed out after 60 seconds. ---> System.Net.WebException: 操作はタイムアウトになりました。
        '---------------------------------
        Dim DriverService = ChromeDriverService.CreateDefaultService()
        DriverService.HideCommandPromptWindow = True
        Dim ChromeOptions = New ChromeOptions()
        ChromeOptions.AddArguments("disable-infobars")
        ChromeOptions.AddArguments("no-sandbox")
        ChromeOptions.AddArguments("--disable-dev-shm-usage")
        'ChromeOptions.AddArgument("--headless")

        Dim nextenginePath As String = "https://ne107.next-engine.com/"
        cr = New ChromeDriver(DriverService, ChromeOptions, TimeSpan.FromSeconds(180)) With {
            .Url = nextenginePath
        }
        cr.Manage.Window.Size = New Size(sizeW, sizeH)
        cr.Manage.Window.Position = New Point(0, 0)
        'Me.Location = New Point(sizeW + 20, 0)

        'ネクストエンジンログイン
        cr.Url = nextenginePath
        'cr.FindElementById("user_login_code").SendKeys("jidou@yongshun006.com")
        'CR_SET_TEXT(cr.FindElementById("user_login_code"), "person1@yongshun006.com")
        cr.FindElementById("user_login_code").SendKeys("person1@yongshun006.com")
        cr.FindElementById("user_password").SendKeys("sdv_c28h")
        'CR_SET_TEXT(cr.FindElementById("user_password"), "sdv_c28h")
        cr.FindElementByName("commit").Submit()

        'トップページ
        cr.Url = "https://base.next-engine.org/apps/launch/?id=65459"

        'サーバーのホスト名が変わった時に取得できるように毎回解析する
        'Dim u As New Uri(cr.Url)
    End Sub

    Private Sub ChromeClose()
        cr.Quit()
    End Sub


    ''' <summary>
    ''' sendkeyで書き込むと遅いので、クリップボードを経由して書き込みする
    ''' </summary>
    ''' <param name="element">書き込むフォームエレメント</param>
    ''' <param name="writeStr">書き込む文字列</param>
    Private Sub CR_SET_TEXT(element As IWebElement, writeStr As String)
        element.Click()
        Clipboard.SetText(writeStr)
        Dim actionlist As Actions = New Actions(cr)
        actionlist.KeyDown(Keys.Control).SendKeys("v").KeyUp(Keys.Control).Perform()
    End Sub


    Public dataPathArray As New ArrayList
    Private Sub DataGridView1_DragDrop(sender As Object, e As DragEventArgs)
        'Dim tenpo As String
        'If ToolStripComboBox1.SelectedIndex = 0 Then
        '    tenpo = "Qoo10"
        'End If


        'dataPathArray.Clear()

        'Dim fCount As Integer = 0
        'For Each filename As String In e.Data.GetData(DataFormats.FileDrop)
        '    dataPathArray.Add(filename)
        '    fCount += 1
        'Next

        'If dataPathArray.Count > 0 Then
        '    DataGridView1.Rows.Clear()
        '    DataGridView1.Columns.Clear()

        '    For i As Integer = 0 To dataPathArray.Count - 1
        '        Dim csvRecords As ArrayList = TM_CSV_READ(dataPathArray(i))(0)

        'If tenpo = "Qoo10" Then
        '    If InStr(csvRecords(0), "配送状態|=|注文番号|=|カート番号|=|配送会社|=|送り状番号|=|発送日|=|発送予定日|=|商品名|=|数量|=|オプション情報|=|オプションコード|=|受取人名|=|販売者商品コード|=|外部広告|=|決済サイト") > 0 Then
        '        Dim title_arr As String() = "店舗伝票番号,受注日,受注郵便番号,受注住所１,受注住所２,受注名,受注名カナ,受注電話番号,受注メールアドレス,発送郵便番号,発送先住所１,発送先住所２,発送先名,発送先カナ,発送電話番号,支払方法,発送方法,商品計,税金,発送料,手数料,ポイント,その他費用,合計金額,ギフトフラグ,時間帯指定,日付指定,作業者欄,備考,商品名,商品コード,商品価格,受注数量,商品オプション,出荷済フラグ,顧客区分,顧客コード,消費税率（%）".Split(",")

        '        For c As Integer = 0 To title_arr.Length - 1
        '            DataGridView1.Columns.Add(c, title_arr(c))
        '        Next

        '        For r As Integer = 1 To csvRecords.Count - 1
        '            Dim lines As String() = Split(csvRecords(r), "|=|")
        '            DataGridView1.Rows.Add(1)
        '            '店舗伝票番号
        '            DataGridView1.Item(0, r - 1).Value = lines(1)
        '            '受注名
        '            DataGridView1.Item(5, r - 1).Value = lines(11)
        '            '支払方法
        '            DataGridView1.Item(15, r - 1).Value = "Qoo10"
        '            '発送方法
        '            If lines(3) = "佐川急便" Then
        '                DataGridView1.Item(16, r - 1).Value = "宅急便"
        '            ElseIf lines(3) = "ゆうパケット" Then
        '                DataGridView1.Item(16, r - 1).Value = "メール便"
        '            End If
        '            '商品名
        '            DataGridView1.Item(29, r - 1).Value = lines(7)
        '            '商品コード
        '            DataGridView1.Item(30, r - 1).Value = lines(12)
        '            '受注数量
        '            DataGridView1.Item(32, r - 1).Value = lines(8)
        '            '商品オプション
        '            If lines(10) <> "" Then
        '                DataGridView1.Item(33, r - 1).
        '                If InStr(lines(10), "=") > 0 Then
        '                    DataGridView1.Item(33, r - 1).Value = lines(10).Replace("=", "")
        '                Else
        '                    DataGridView1.Item(33, r - 1).Value = lines(10)
        '                End If  '出荷済フラグ
        '                                       '出荷済フラグe                             DataGridView1.Item(34, r - 1).Value = lines(8)                  Next
        '    Else
        '        MsgBox("Qoo10のフォーマットじゃないです。", MsgBoxStyle.SystemModal)
        '        Exit Sub
        '    End If
        'End If

        'Next
        'End If
    End Sub

    Private Sub DataGridView1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub スタートToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles スタートToolStripMenuItem.Click
        If IsNetworkAvailable() = False Then
            Exit Sub
        End If

        td1 = New Thread(AddressOf startThread)
        td1.Start()
    End Sub

    Private Sub ストップToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ストップToolStripMenuItem.Click
        td1.Abort()
        td1_action = False

        'ChromeClose()

        Application.DoEvents()

        'Do While (td1.ThreadState <> ThreadState.Aborted)
        '    Thread.Sleep(100)
        'Loop
    End Sub

    Private Sub ダウンロードToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ダウンロードToolStripMenuItem.Click
        If DataGridView1.Rows.Count = 0 Then
            MsgBox("データがないです。", MsgBoxStyle.SystemModal)
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
            sfd.FileName = "注文番号チェック.csv"
        End If

        'ダイアログを表示する
        If sfd.ShowDialog(Me) = DialogResult.OK Then
            SaveCsv(sfd.FileName, 0, DataGridView1, False)
            MsgBox(sfd.FileName & vbCrLf & "保存しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    ' DataGridViewをCSV出力する
    Public Sub SaveCsv(ByVal fp As String, ByVal mode As Integer, dgv As DataGridView, header_flag As Boolean)

        dgv.EndEdit()


        Dim CsvDelimiter As String = splitter
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
                                Case InStr(dt, splitter), InStr(dt, vbLf), InStr(dt, vbCr), InStr(dt, """") 'Regex.IsMatch(dt, ".*,.*|.*" & vbLf & ".*|.*" & vbCr & ".*|.*"".*")
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

    Private Sub クローズToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles クローズToolStripMenuItem1.Click
        ChromeClose()
    End Sub


    Private Sub DataGridView2_DragDrop(sender As Object, e As DragEventArgs) Handles DataGridView2.DragDrop
        dataPathArray.Clear()

        Dim fCount As Integer = 0
        For Each filename As String In e.Data.GetData(DataFormats.FileDrop)
            dataPathArray.Add(filename)
            fCount += 1
        Next

        If dataPathArray.Count > 1 Then
            MsgBox("一つファイルをいれてください。")
            Return
        Else
            DropRun(dataPathArray, fCount, True, True)
        End If
    End Sub

    Public Sub DropRun(dataPathArray As ArrayList, fCount As Integer, Optional headerAdd As Boolean = False, Optional addFlag As Boolean = True)
        If DataGridView2.RowCount > 0 Then
            DataGridView2.Rows.Clear()
            DataGridView2.Columns.Clear()
        End If

        For Each filename As String In dataPathArray
            Dim csvRecords As New ArrayList()
            csvRecords = CSV_READ(filename)

            Dim startRowNum As Integer = 0
            'If num = 0 And DataGridView2.RowCount > 0 Then
            '    Dim DR As DialogResult = MsgBox("追加ファイルのヘッダーを削除して良いですか？", MsgBoxStyle.YesNo Or MsgBoxStyle.SystemModal)
            '    If DR = DialogResult.Yes Then
            '        startRowNum = 1
            '    Else
            '        startRowNum = 0
            '    End If
            'ElseIf num > 0 And Not headerAdd Then
            '    startRowNum = 1
            'End If

            For r As Integer = startRowNum To csvRecords.Count - 1
                Dim sArray As String() = Split(csvRecords(r), "|=|")

                If r = 0 Then
                    If Regex.IsMatch(csvRecords(0), "配送状態|=|注文番号|=|カート番号|=|配送会社|=|送り状番号|=|発送日|=|注文日|=|入金日|=|お届け希望日|=|発送予定日|=|配送完了日|=|配送方法|=|商品番号|=|商品名|=|数量|=|オプション情報|=|オプションコード|=|おまけ|=|受取人名|=|受取人名(フリガナ)|=|受取人電話番号|=|受取人携帯電話番号|=|住所|=|郵便番号|=|国家|=|送料の決済|=|決済サイト|=|通貨|=|購入者決済金額|=|販売価格|=|割引額|=|注文金額の合計|=|供給原価の合計|=|購入者名|=|購入者名(フリガナ)|=|配送要請事項|=|購入者電話番号|=|購入者携帯電話番号|=|販売者商品コード|=|JANコード|=|規格番号|=|プレゼント贈り主|=|外部広告|=|素材") Then

                    Else
                        MsgBox("正しいファイルをいれてください。")
                        Return
                    End If
                End If

                'If DataGridView2.ColumnCount < sArray.Length - 1 Then
                For c As Integer = DataGridView2.ColumnCount To sArray.Length - 1
                    DataGridView2.Columns.Add(c, sArray(c))
                    DataGridView2.Columns(c).SortMode = DataGridViewColumnSortMode.NotSortable
                Next
                'End If

                If r > startRowNum Then
                    DataGridView2.Rows.Add(sArray)
                End If
            Next
        Next
    End Sub

    Private Sub DataGridView2_DragEnter(sender As Object, e As DragEventArgs) Handles DataGridView2.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub DataGridView2_RowPostPaint(sender As Object, e As DataGridViewRowPostPaintEventArgs) Handles DataGridView2.RowPostPaint, DataGridView3.RowPostPaint
        ' 行ヘッダのセル領域を、行番号を描画する長方形とする（ただし右端に4ドットのすき間を空ける）
        Dim rect As New Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, sender.RowHeadersWidth - 4, sender.Rows(e.RowIndex).Height)

        ' 上記の長方形内に行番号を縦方向中央＆右詰で描画する　フォントや色は行ヘッダの既定値を使用する
        TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), sender.RowHeadersDefaultCellStyle.Font, rect, sender.RowHeadersDefaultCellStyle.ForeColor, TextFormatFlags.VerticalCenter Or TextFormatFlags.Right)
    End Sub

    Public Function CSV_READ(ByRef path As String)
        Dim csvText As String = ""
        Dim file As FileInfo = New FileInfo(path)
        Dim fileEnc As String = ""
        Using reader As ReadJEnc.FileReader = New ReadJEnc.FileReader(file)
            Dim c As ReadJEnc.CharCode = reader.Read(file)
            fileEnc = c.Name
            csvText = reader.Text
        End Using
        If fileEnc = "ShiftJIS" Then
            fileEnc = "Shift-JIS"
        End If
        'ToolStripDropDownButton10.Text = fileEnc

        'Dim csvText As String = file.ReadAllText(path, System.Text.Encoding.Default)

        'カンマ、タブ等の区切り文字認識
        Dim header As String() = Regex.Split(csvText, vbCrLf & "|" & vbCr & "|" & vbLf)
        If InStr(header(0), vbTab) Then
            splitter = vbTab
            'ToolStripStatusLabel9.Text = "tab"
        Else
            splitter = ","
            'ToolStripStatusLabel9.Text = "comma"
        End If

        '前後の改行を削除しておく
        csvText = csvText.Trim(
            New Char() {ControlChars.Cr, ControlChars.Lf})

        Dim csvRecords As New System.Collections.ArrayList
        Dim csvFields As New System.Collections.ArrayList

        Dim csvTextLength As Integer = csvText.Length
        Dim startPos As Integer = 0
        Dim endPos As Integer = 0
        Dim field As String = ""

        While True
            '空白を飛ばす
            'While startPos < csvTextLength _
            '    AndAlso (csvText.Chars(startPos) = " "c OrElse csvText.Chars(startPos) = ControlChars.Tab)
            '    startPos += 1
            'End While

            'データの最後の位置を取得
            If startPos < csvTextLength _
                AndAlso csvText.Chars(startPos) = ControlChars.Quote Then
                '"で囲まれているとき
                '最後の"を探す
                endPos = startPos
                While True
                    endPos = csvText.IndexOf(ControlChars.Quote, endPos + 1)
                    If endPos < 0 Then
                        Throw New ApplicationException("""が不正")
                    End If
                    '"が2つ続かない時は終了
                    If endPos + 1 = csvTextLength OrElse csvText.Chars((endPos + 1)) <> ControlChars.Quote Then
                        Exit While
                    End If
                    '"が2つ続く
                    endPos += 1
                End While

                '一つのフィールドを取り出す
                field = csvText.Substring(startPos, endPos - startPos + 1)
                '""を"にする
                field = field.Substring(1, field.Length - 2).Replace("""""", """")

                endPos += 1
                '空白を飛ばす
                While endPos < csvTextLength AndAlso
                    csvText.Chars(endPos) <> splitter AndAlso
                    csvText.Chars(endPos) <> ControlChars.Lf
                    endPos += 1
                End While
            Else
                '"で囲まれていない
                'カンマか改行の位置
                endPos = startPos
                While endPos < csvTextLength AndAlso
                    csvText.Chars(endPos) <> splitter AndAlso
                    csvText.Chars(endPos) <> ControlChars.Lf
                    endPos += 1
                End While

                '一つのフィールドを取り出す
                field = csvText.Substring(startPos, endPos - startPos)
                '後の空白を削除
                field = field.TrimEnd()
            End If

            'フィールドの追加
            csvFields.Add(field)

            '行の終了か調べる
            If endPos >= csvTextLength OrElse
                csvText.Chars(endPos) = ControlChars.Lf Then
                '行の終了
                'レコードの追加
                csvFields.TrimToSize()
                Dim str As String = ""
                For i As Integer = 0 To csvFields.Count - 1
                    If i = 0 Then
                        str = csvFields(i)
                    Else
                        str &= "|=|" & csvFields(i)
                    End If
                Next
                'csvRecords.Add(csvFields)
                csvRecords.Add(str)
                csvFields = New System.Collections.ArrayList(csvFields.Count)

                If endPos >= csvTextLength Then
                    '終了
                    Exit While
                End If
            End If

            '次のデータの開始位置
            startPos = endPos + 1
        End While

        csvRecords.TrimToSize()

        For r As Integer = 0 To csvRecords.Count - 1
            If InStr(csvRecords(r), "E+") Then
                MsgBox("エクセルで保存し直した形跡があります。データを作成し直してください", MsgBoxStyle.SystemModal)
                Exit For
            End If
        Next

        'If Format(System.IO.File.GetCreationTime(path), "yyyy/MM/dd H:mm:ss") <> Format(System.IO.File.GetLastWriteTime(path), "yyyy/MM/dd H:mm:ss") Then
        '    MsgBox("エクセルで保存し直した形跡があります。データを作成し直してください", MsgBoxStyle.SystemModal)
        'End If

        Return csvRecords
    End Function

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        startAction()
    End Sub

    Private Sub startAction()
        If DataGridView2.RowCount = 0 Then
            MsgBox("ファイルをドロップしてください。")
            Return
        End If

        DataGridView3.Rows.Clear()
        DataGridView3.Columns.Clear()

        Dim header As String() = New String() {"店舗伝票番号", "受注日", "受注郵便番号", "受注住所１", "受注住所２", "受注名", "受注名カナ", "受注電話番号", "受注メールアドレス", "発送郵便番号", "発送先住所１", "発送先住所２", "発送先名", "発送先カナ", "発送電話番号", "支払方法", "発送方法", "商品計", "税金", "発送料", "手数料", "ポイント", "その他費用", "合計金額", "ギフトフラグ", "時間帯指定", "日付指定", "作業者欄", "備考", "商品名", "商品コード", "商品価格", "受注数量", "商品オプション", "出荷済フラグ"}

        For i As Integer = 0 To header.Length - 1
            Dim res As String() = Split(header(i), ",")
            DataGridView3.Columns.Add(i, res(0))
            DataGridView3.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        Dim dH2 As ArrayList = TM_HEADER_GET(DataGridView2)
        For r As Integer = 0 To DataGridView2.RowCount - 1
            Dim denpyo As String = DataGridView2.Item(dH2.IndexOf("注文番号"), r).Value
            Dim order_date As String = DataGridView2.Item(dH2.IndexOf("注文日"), r).Value
            Dim order_post As String = Replace(DataGridView2.Item(dH2.IndexOf("郵便番号"), r).Value, "'", "")
            Dim order_address As String = DataGridView2.Item(dH2.IndexOf("住所"), r).Value
            Dim order_name As String = DataGridView2.Item(dH2.IndexOf("受取人名"), r).Value
            Dim order_name2 As String = DataGridView2.Item(dH2.IndexOf("受取人名(フリガナ)"), r).Value
            Dim order_tel As String = ""
            If DataGridView2.Item(dH2.IndexOf("受取人携帯電話番号"), r).Value <> "-" And DataGridView2.Item(dH2.IndexOf("受取人携帯電話番号"), r).Value <> "" Then
                order_tel = DataGridView2.Item(dH2.IndexOf("受取人携帯電話番号"), r).Value
            End If
            Dim hasou As String = DataGridView2.Item(dH2.IndexOf("配送会社"), r).Value
            Dim price As String = Replace(DataGridView2.Item(dH2.IndexOf("購入者決済金額"), r).Value, ",", "")
            Dim price_total As String = Replace(DataGridView2.Item(dH2.IndexOf("注文金額の合計"), r).Value, ",", "")
            Dim good_name As String = DataGridView2.Item(dH2.IndexOf("商品名"), r).Value
            Dim code As String = Replace(DataGridView2.Item(dH2.IndexOf("販売者商品コード"), r).Value + DataGridView2.Item(dH2.IndexOf("オプションコード"), r).Value, "'", "")
            Dim price2 As String = DataGridView2.Item(dH2.IndexOf("販売価格"), r).Value
            Dim count As String = DataGridView2.Item(dH2.IndexOf("数量"), r).Value
            Dim op As String = DataGridView2.Item(dH2.IndexOf("オプション情報"), r).Value

            Dim add_row As String() = {denpyo, order_date, order_post, order_address, "", order_name, order_name2, order_tel, "", order_post, order_address, "", order_name, order_name2, order_tel, "請求書払い", hasou, price_total, 0, 0, 0, 0, 0, price, 0, "", "", "", "", good_name, code, price2, count, op, 0}
            DataGridView3.Rows.Add(add_row)

            If order_tel = "" Then
                DataGridView3(7, r).Style.BackColor = Color.Yellow
                DataGridView3(14, r).Style.BackColor = Color.Yellow
            End If
        Next
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        If DataGridView3.Rows.Count = 0 Then
            MsgBox("データがないです。", MsgBoxStyle.SystemModal)
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
            sfd.FileName = "hanyo-jyuchu_" + Format(System.DateTime.Now, "yyyyMMddHmmss") + ".csv"
        End If

        'ダイアログを表示する
        If sfd.ShowDialog(Me) = DialogResult.OK Then
            SaveCsv(sfd.FileName, 0, DataGridView3, True)
            MsgBox(sfd.FileName & vbCrLf & "保存しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    Private Sub ToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem3.Click
        If DataGridView3.Rows.Count = 0 Then
            MsgBox("データがないです。", MsgBoxStyle.SystemModal)
            Exit Sub
        End If


        Dim dH3 As ArrayList = TM_HEADER_GET(DataGridView3)
        For r As Integer = 0 To DataGridView3.RowCount - 1
            If DataGridView3.Item(dH3.IndexOf("受注電話番号"), r).Value = "" Then
                DataGridView3.Item(dH3.IndexOf("受注電話番号"), r).Value = "000-0000-0000"
            End If
            If DataGridView3.Item(dH3.IndexOf("発送電話番号"), r).Value = "" Then
                DataGridView3.Item(dH3.IndexOf("発送電話番号"), r).Value = "000-0000-0000"
            End If
        Next

    End Sub
End Class