Imports System.Net
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports OpenQA.Selenium.Chrome
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Interactions
Imports System.Threading
Imports Hnx8

Public Class SagawaSpare
    Private nextenginePath As String = "https://ne81.next-engine.com/"

    Dim splitter As String = ","    '区切り文字

    Dim sc As Integer = 0
    Dim sr As Integer = 0

    Private sizeW As Integer = 500
    Private sizeH As Integer = 800

    Private Delegate Sub CallDelegate()

    Dim cellChange As Boolean = True

    Public dataPathArray As New ArrayList
    Public headerAdd As Boolean = True
    Public addFlag As Boolean = True

    Private cr As ChromeDriver = Nothing

    Private sakiPath2 As String = ""

    Private Sub SagawaSpare_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        sakiPath2 = Form1.ロケーションToolStripMenuItem.Text
        DGV7_update()

        DGV10_update()

        DGV6_update()
    End Sub

    Private dgv6CodeArray As New ArrayList
    'Private dgv6DaihyoCodeArray As New ArrayList

    Private Sub DGV6_update()
        If DGV6.RowCount > 0 Then
            DGV6.Rows.Clear()
            DGV6.Columns.Clear()
        End If

        'ヘッダー行作成
        Dim header As String() = New String() {"商品コード", "商品名", "商品分類タグ", "代表商品コード", "ship-weight", "ロケーション", "梱包サイズ", "特殊", "sw2"}
        For c As Integer = 0 To header.Length - 1
            DGV6.Columns.Add(c, header(c))
            DGV6.Columns(c).SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        Dim fNameL As String = Path.GetDirectoryName(Form1.appPath) & "\location.csv"
        Dim csvRecords2 As New ArrayList()
        csvRecords2 = TM_CSV_READ(fNameL)(0)
        Dim codeArray2 As New ArrayList()
        For r As Integer = 0 To csvRecords2.Count - 1
            Dim sArray As String() = Split(csvRecords2(r), "|=|")
            codeArray2.Add(sArray(0))
            sArray(0) = Form1.StrConvToNarrow(sArray(0))    '商品コードを小文字で揃える
            DGV6.Rows.Add(sArray)
        Next

        'For r As Integer = 0 To csvRecords2.Count - 1
        '    Dim sArray As String() = Split(csvRecords2(r), "|=|")
        '    sArray(0) = Form1.StrConvToNarrow(sArray(0))    '商品コードを小文字で揃える
        '    DGV6.Rows.Add(sArray)
        'Next

        dgv6CodeArray.Clear()
        'dgv6DaihyoCodeArray.Clear()
        Dim dH6 As ArrayList = TM_HEADER_GET(DGV6)
        For r As Integer = 0 To DGV6.RowCount - 1
            dgv6CodeArray.Add(DGV6.Item(dH6.IndexOf("商品コード"), r).Value.ToString.ToLower)
            'dgv6DaihyoCodeArray.Add(DGV6.Item(dH6.IndexOf("代表商品コード"), r).Value.ToString.ToLower)
        Next
    End Sub

    Private Sub DGV7_update()
        If DGV7.RowCount > 0 Then
            DGV7.Rows.Clear()
            DGV7.Columns.Clear()
        End If

        'ヘッダー行作成
        Dim header As String() = New String() {"都道府県", "サイズ", "運賃"}
        For c As Integer = 0 To header.Length - 1
            DGV7.Columns.Add(c, header(c))
            DGV7.Columns(c).SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        Dim fNameL As String = Path.GetDirectoryName(Form1.appPath) & "\sagawa_fare.csv"
        Dim csvRecords2 As New ArrayList()
        csvRecords2 = TM_CSV_READ(fNameL)(0)
        Dim codeArray2 As New ArrayList()
        For r As Integer = 0 To csvRecords2.Count - 1
            Dim sArray As String() = Split(csvRecords2(r), "|=|")
            codeArray2.Add(sArray(0))
            sArray(0) = Form1.StrConvToNarrow(sArray(0))    '商品コードを小文字で揃える
            DGV7.Rows.Add(sArray)
        Next
    End Sub

    Private Sub DGV10_update()
        'サイズ容量読み込み
        DGV10.Rows.Clear()
        Dim csvRecords As New ArrayList
        csvRecords = TM_CSV_READ(Path.GetDirectoryName(Form1.appPath) & "\config\version2\サイズ容量.txt")(0)
        For r As Integer = 0 To csvRecords.Count - 1
            Dim sArray As String() = Split(csvRecords(r), "|=|")
            DGV10.Rows.Add(sArray)
        Next
    End Sub

    Private Sub DGV1_DragDrop(sender As Object, e As DragEventArgs) Handles DGV1.DragDrop, DGV11.DragDrop, DGV12.DragDrop
        dataPathArray.Clear()

        Dim fCount As Integer = 0
        Dim filname_ As String = ""

        For Each filename As String In e.Data.GetData(DataFormats.FileDrop)
            Dim type = filename.Split(".")
            If type(type.Length - 1) <> "csv" Then
                MsgBox("csv対応のみです。")
                Exit Sub
            End If

            dataPathArray.Add(filename)
            fCount += 1
            filname_ += filename + ","
        Next

        If dataPathArray.Count > 1 Then
            'CSV_F_files.DGV1.Rows.Clear()
            'For i As Integer = 0 To dataPathArray.Count - 1
            '    CSV_F_files.DGV1.Rows.Add(True, Path.GetFileName(dataPathArray(i)))
            'Next
            'CSV_F_files.dataPathArray = dataPathArray
            'Dim DR As DialogResult = CSV_F_files.ShowDialog()
            'If DR = DialogResult.OK Then
            DropRun(dataPathArray, sender, fCount, headerAdd, addFlag)
            'Else
            '    Exit Sub
            'End If
        Else
            DropRun(dataPathArray, sender, fCount, True, True)
        End If
    End Sub

    Public Sub DropRun(dataPathArray As ArrayList, dgv As DataGridView, fCount As Integer, Optional headerAdd As Boolean = False, Optional addFlag As Boolean = True)
        If dgv.RowCount > 0 Then
            If addFlag Then
                If dataPathArray.Count = 1 Then
                    Dim DR As DialogResult = MsgBox("追加しますか？", MsgBoxStyle.YesNoCancel Or MsgBoxStyle.SystemModal)
                    If DR = DialogResult.No Then
                        dgv.Columns.Clear()
                    ElseIf DR = DialogResult.Cancel Then
                        Exit Sub
                    End If
                End If
            Else
                dgv.Columns.Clear()
            End If
        End If

        PreDIV(dgv)

        Dim num As Integer = 0
        Dim colorArray As Color() = {Color.LightYellow, Color.LightCyan}
        For Each filename As String In dataPathArray
            Dim csvRecords As New ArrayList()
            csvRecords = TM_CSV_READ2(filename)

            Dim startRowNum As Integer = 0
            If num = 0 And dgv.RowCount > 0 Then
                Dim DR As DialogResult = MsgBox("追加ファイルのヘッダーを削除して良いですか？", MsgBoxStyle.YesNo Or MsgBoxStyle.SystemModal)
                If DR = DialogResult.Yes Then
                    startRowNum = 1
                Else
                    startRowNum = 0
                End If
            ElseIf num > 0 And Not headerAdd Then
                startRowNum = 1
            End If

            Dim insert_cob = 0
            For r As Integer = startRowNum To csvRecords.Count - 1
                Dim sArray As String() = Split(csvRecords(r), "|=|")
                If dgv.ColumnCount < sArray.Length - 1 Then
                    For c As Integer = dgv.ColumnCount To sArray.Length - 1
                        dgv.Columns.Add(c, sArray(c))
                        dgv.Columns(c).SortMode = DataGridViewColumnSortMode.NotSortable
                    Next
                End If

                If r > startRowNum Then
                    dgv.Rows.Add(sArray)
                End If

                If num > 0 Then
                    Dim cNum As Integer = num Mod 2
                    dgv.Item(0, dgv.RowCount - 2).Style.BackColor = colorArray(cNum) '追加した時だけ、一時的に色を付ける
                End If
            Next

            If num = 0 Then
                Me.Text = filename
            End If
            num += 1
        Next
    End Sub

    Public Sub PreDIV(dgv As DataGridView)
        '-----------------------------------------------------------------------
        'sc = DataGridView1.FirstDisplayedScrollingColumnIndex
        sc = dgv.HorizontalScrollingOffset
        sr = dgv.FirstDisplayedScrollingRowIndex
        'For r As Integer = 0 To DataGridView1.RowCount - 1
        '    If DataGridView1.Rows(r).Visible And Not DataGridView1.Rows(r).IsNewRow Then
        '        sr = DataGridView1.FirstDisplayedScrollingRowIndex
        '    End If
        'Next
        cellChange = False
        'DataGridView1.ScrollBars = ScrollBars.None
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
        dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        dgv.ScrollBars = ScrollBars.Both
        '-----------------------------------------------------------------------
    End Sub

    Private Sub DGV1_DragEnter(sender As Object, e As DragEventArgs) Handles DGV1.DragEnter, DGV8.DragEnter, DGV11.DragEnter, DGV12.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub DGV_RowPostPaint(ByVal sender As DataGridView, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DGV1.RowPostPaint, DGV2.RowPostPaint, DGV6.RowPostPaint, DGV7.RowPostPaint
        ' 行ヘッダのセル領域を、行番号を描画する長方形とする（ただし右端に4ドットのすき間を空ける）
        Dim rect As New Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, sender.RowHeadersWidth - 4, sender.Rows(e.RowIndex).Height)
        ' 上記の長方形内に行番号を縦方向中央＆右詰で描画する　フォントや色は行ヘッダの既定値を使用する
        TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), sender.RowHeadersDefaultCellStyle.Font, rect, sender.RowHeadersDefaultCellStyle.ForeColor, TextFormatFlags.VerticalCenter Or TextFormatFlags.Right)
    End Sub

    Private Sub フォームのリセットToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles フォームのリセットToolStripMenuItem.Click
        Me.Dispose()
        Dim frm As Form = New SagawaSpare
        frm.Show()
    End Sub

    Dim td1 As Thread

    Dim td1_action As Boolean = False
    Private Sub startThread1()
        If cr Is Nothing Then
            ChromeStart()
        End If

        td1_action = True

        Invoke(New CallDelegate(AddressOf action1))

        td1.Abort()
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
            MsgBox("未接続です。")
        End If
        Return IsNetworkAvailable_Flag
    End Function

    Dim homeURL As String = nextenginePath & "Userjyuchu/index?search_condi=17"
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

        cr = New ChromeDriver(DriverService, ChromeOptions, TimeSpan.FromSeconds(180)) With {
            .Url = homeURL
        }
        cr.Manage.Window.Size = New Size(sizeW, sizeH)
        cr.Manage.Window.Position = New Point(0, 0)
        Me.Location = New Point(sizeW + 20, 0)

        'ネクストエンジンログイン
        cr.Url = homeURL
        'cr.FindElementById("user_login_code").SendKeys("jidou@yongshun006.com")
        CR_SET_TEXT(cr.FindElementById("user_login_code"), "jidou@yongshun006.com")
        'cr.FindElementById("user_password").SendKeys("aa9922--")
        CR_SET_TEXT(cr.FindElementById("user_password"), "aa9922--")
        cr.FindElementByName("commit").Submit()

        'トップページ
        cr.Url = "https://base.next-engine.org/apps/launch/?id=65459"

        'サーバーのホスト名が変わった時に取得できるように毎回解析する
        Dim u As New Uri(cr.Url)
        nextenginePath = "https://" & u.Host & "/"
        homeURL = nextenginePath & "Userjyuchu/index?search_condi=17"
    End Sub

    Private Sub ChromeClose()
        cr.Quit()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If DGV1.Rows.Count = 0 Then
            MsgBox("データがないです。")
            Exit Sub
        End If

        Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)

        'For c As Integer = 0 To dH1.Count - 1
        '    Console.WriteLine(dH1(c))
        'Next
        If dH1.IndexOf("お客様管理№") = 0 Then
            MsgBox("「お客様管理№」フィールドがないです。")
            Exit Sub
        End If

        If IsNetworkAvailable() = False Then
            Exit Sub
        End If

        If cr Is Nothing Then
            ChromeStart()
        End If

        'action1()
        td1 = New Thread(AddressOf startThread1)
        td1.Start()

    End Sub

    Private Sub action1()
        DGV2.Rows.Clear()
        DGV2.Columns.Clear()

        Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
        For c As Integer = 0 To dH1.Count - 1
            DGV2.Columns.Add(c, dH1(c))
        Next

        'DGV2.Columns.Add(dH1.Count + 1, "種類")
        DGV2.Columns.Add(dH1.Count + 1, "商品(ne)")
        DGV2.Columns.Add(dH1.Count + 2, "サイズ(発注)")
        'DGV2.Columns.Add(dH1.Count + 4, "運賃(ne)")
        'DGV2.Columns.Add(dH1.Count + 3, "サイズ計算")
        'DGV2.Columns.Add(dH1.Count + 4, "サイズ計算説明")
        'DGV2.Columns.Add(dH1.Count + 6, "便数計算")
        DGV2.Columns.Add(dH1.Count + 3, "運賃(計算)")
        DGV2.Columns.Add(dH1.Count + 4, "運賃(計算説明)")
        DGV2.Columns.Add(dH1.Count + 5, "運賃差(計算 - 佐川)")

        Dim dH2 As ArrayList = TM_HEADER_GET(DGV2)

        'ToolStripProgressBar1.Maximum = DGV1.Rows.Count
        'ToolStripProgressBar1.Value = 0
        Label1.Text = 0
        Label3.Text = DGV1.Rows.Count

        For r As Integer = 0 To DGV1.Rows.Count - 1
            DGV2.Rows.Add(1)
            For c As Integer = 0 To dH1.Count - 1
                DGV2.Item(dH2.IndexOf(dH1(c)), r).Value = DGV1.Item(dH1.IndexOf(dH1(c)), r).Value
            Next
            Application.DoEvents()

            If td1_action = True Then
                Dim dH6 As ArrayList = TM_HEADER_GET(DGV6)
                Dim dH7 As ArrayList = TM_HEADER_GET(DGV7)
                Dim denpyono As String = DGV1.Item(dH1.IndexOf("お客様管理№"), r).Value
                Dim place As String = DGV1.Item(dH1.IndexOf("都道府県"), r).Value
                Dim fare_sg As Integer = DGV1.Item(dH1.IndexOf("運賃"), r).Value
                'Dim kosuu_sg As Integer = DGV1.Item(dH1.IndexOf("個数"), r).Value
                If denpyono <> "" Then
                    If denpyono = Int(denpyono) Then
                        Dim url As String = nextenginePath & "Userjyuchu/jyuchuInp?kensaku_denpyo_no=" & denpyono & "&jyuchu_meisai_order=jyuchu_meisai_gyo"
                        'cr.Url = "https://ne81.next-engine.com/Userjyuchu/jyuchuInp?kensaku_denpyo_no=" & denpyono & "&jyuchu_meisai_order=jyuchu_meisai_gyo"
                        'alertが残っているとエラーになるので対応
                        For i As Integer = 0 To 3
                            If IsAlertPresent() Then
                                Exit For
                            Else
                                Thread.Sleep(1000 * 1)
                                Application.DoEvents()
                            End If
                        Next
                        '伝票のURLを開く
                        If td1_action = True Then
                            cr.Url = url
                            Dim result As String = cr.FindElementById("jyuchu_msg").Text
                            If InStr(result, "伝票番号の検索結果がありました") = 1 Then
                                'Dim souryo_ne As String = cr.FindElementById("hasou_kin").GetAttribute("value") '発送代
                                'DGV2.Item(dH2.IndexOf("運賃(ne)"), r).Value = souryo_ne

                                'chromeからwebbrowserへ移行
                                Dim wb As WebBrowser = WebBrowser1
                                Dim html As String = cr.ExecuteScript("return document.body.outerHTML")
                                wb.ScriptErrorsSuppressed = True
                                wb.DocumentText = html
                                Application.DoEvents()

                                If InStr(wb.Document.Body.OuterHtml, "jyuchuMeisai_tablene_table") > 0 Then
                                    Dim tb As HtmlElementCollection = wb.Document.GetElementsByTagName("table")

                                    For Each he As HtmlElement In tb
                                        If InStr(he.Id, "jyuchuMeisai_tablene_table") > 0 Then
                                            Dim td_count As Integer = 0
                                            td_count = he.GetElementsByTagName("tr").Count
                                            Application.DoEvents()

                                            If td_count > 0 Then
                                                Dim cd_str = ""
                                                Dim weight_str = ""
                                                'Dim cd_temp As String = ""
                                                'Dim doukon As Boolean = False
                                                'Dim meisai_c As Integer = 0
                                                For c As Integer = 0 To td_count - 2
                                                    'Dim cd As String = cr.FindElementById("jyuchuMeisai_table" & c & "-3").GetAttribute("value")
                                                    Dim code_id As String = "jyuchuMeisai_table" & c & "-3"
                                                    Dim jyucyu_count_id As String = "jyuchuMeisai_table" & c & "-7" '引当数
                                                    If wb.Document.GetElementById(code_id) <> Nothing Then
                                                        Dim code As String = cr.FindElementById(code_id).Text.ToLower
                                                        'Dim count As String = MaruMojiConv(cr.FindElementById(jyucyu_count_id).Text)
                                                        Dim count As String = cr.FindElementById(jyucyu_count_id).Text
                                                        cd_str += code + "*" + count + ","

                                                        'If cd_temp = "" Then
                                                        '    cd_temp = code.ToLower
                                                        'Else
                                                        '    If cd_temp <> code Then
                                                        '        doukon = True
                                                        '    End If
                                                        'End If

                                                        If Not dgv6CodeArray.Contains(code) Then
                                                            weight_str += "計算不能" + "*" + count + ","
                                                        Else
                                                            Dim weight As String = DGV6.Item(dH6.IndexOf("ship-weight"), dgv6CodeArray.IndexOf(code)).Value
                                                            weight_str += weight + "*" + count + ","
                                                        End If
                                                        'meisai_c += 1
                                                    End If
                                                Next
                                                If cd_str <> "" Then
                                                    cd_str = cd_str.Substring(0, cd_str.Length - 1)
                                                    DGV2.Item(dH2.IndexOf("商品(ne)"), r).Value = cd_str
                                                End If

                                                If weight_str <> "" Then
                                                    weight_str = weight_str.Substring(0, weight_str.Length - 1)
                                                    DGV2.Item(dH2.IndexOf("サイズ(発注)"), r).Value = weight_str
                                                End If

                                                '同梱かどうか
                                                'If meisai_c > 1 Then
                                                '    DGV2.Item(dH2.IndexOf("種類"), r).Value = "同梱"
                                                'ElseIf meisai_c = 1 Then
                                                '    DGV2.Item(dH2.IndexOf("種類"), r).Value = "単品"
                                                'End If

                                                'Dim cdArray_ As String() = Split(cd_str, ",")
                                                Dim weightarray_ As String() = Split(weight_str, ",")
                                                'Dim checkcodejuchusu_ad228 As Integer = 0

                                                'For c7 As Integer = 0 To cdArray_.Length - 1
                                                '    Dim code_ As String = cdArray_(c7).Split("*")(0)
                                                '    Dim count_ As String = cdArray_(c7).Split("*")(1)

                                                '    If code_.ToLower = "ad228" Or code_.ToLower = "ad228-be" Or code_.ToLower = "ad228-bl" Or code_.ToLower = "ad228-co" Or code_.ToLower = "ad228-gr" Or code_.ToLower = "ad228-sb" Or code_.ToLower = "ad228-wa" Then
                                                '        If count_(1) = Int(count_(1)) Then
                                                '            checkcodejuchusu_ad228 = checkcodejuchusu_ad228 + count_
                                                '        End If
                                                '    End If
                                                'Next

                                                'Dim checkcodejuchusu_ad228_even = False
                                                'Dim checkcodejuchusu_ad228_even_count = 0
                                                'If checkcodejuchusu_ad228 > 1 Then
                                                '    If checkcodejuchusu_ad228 Mod 2 = 0 Then
                                                '        checkcodejuchusu_ad228_even = True
                                                '    Else
                                                '        checkcodejuchusu_ad228_even = False
                                                '        checkcodejuchusu_ad228_even_count = checkcodejuchusu_ad228 / 2
                                                '    End If
                                                'End If

                                                'Dim haisouSize As Double = 0
                                                'Dim haisouSize_info As String = ""
                                                'For w7 As Integer = 0 To weightArray_.Length - 1
                                                '    Dim haisouKind As String = ""
                                                '    Dim weight_str_ As String() = weightArray_(w7).Split("*")
                                                '    Dim code_ As String = cdArray_(w7).Split("*")(0).ToLower
                                                '    Dim count_ As String = cdArray_(w7).Split("*")(1)
                                                '    Dim count_maru As String = MaruMojiConv(count_)
                                                '    If weight_str_(0) <> "計算不能" Then
                                                '        'Dim weight_ As String() = weight_str_(0).Split("(")
                                                '        Dim sw As String = weight_str_(0)

                                                '        If Regex.IsMatch(sw, "P|p") Then
                                                '            haisouKind = "宅配便"
                                                '        ElseIf Regex.IsMatch(sw, "M|m") Then
                                                '            haisouKind = "メール便"
                                                '        ElseIf Regex.IsMatch(sw, "T|t") Then
                                                '            haisouKind = "定形外"
                                                '        End If

                                                '        Dim weight As String = Regex.Replace(sw, "P|p|M|m|T|t", "")
                                                '        weight = Regex.Replace(weight, "\(.*\)", "")
                                                '        Dim w2 As String = Regex.Match(weight, "\(.*\)").Value
                                                '        w2 = Regex.Replace(w2, "\(|\)", "")

                                                '        Dim henkouFlag As Boolean = False

                                                '        If Regex.IsMatch(sw, "M|m|T|t") Then
                                                '            henkouFlag = True
                                                '        End If

                                                '        Dim sp_check = True

                                                '        If haisouKind <> "" Then
                                                '            If code_ = "ny263-51" Or code_ = "ny306-51" Or code_ = "ny264-100-4000" Or code_ = "de055" Or code_ = "de055-01" Then
                                                '                weight = "200"
                                                '                sp_check = False
                                                '            End If

                                                '            If (code_ = "ad228" Or code_ = "ad228-be" Or code_ = "ad228-bl" Or code_ = "ad228-co" Or code_ = "ad228-gr" Or code_ = "ad228-sb" Or code_ = "ad228-wa") And checkcodejuchusu_ad228_even Then
                                                '                weight = "50"
                                                '                sp_check = False
                                                '            ElseIf (code_ = "ad228" Or code_ = "ad228-be" Or code_ = "ad228-bl" Or code_ = "ad228-co" Or code_ = "ad228-gr" Or code_ = "ad228-sb" Or code_ = "ad228-wa") And checkcodejuchusu_ad228_even = False Then
                                                '                If checkcodejuchusu_ad228 = 1 Then
                                                '                    weight = "100"
                                                '                    sp_check = False
                                                '                Else
                                                '                    checkcodejuchusu_ad228_even_count = checkcodejuchusu_ad228_even_count - 1
                                                '                    If checkcodejuchusu_ad228_even_count = 1 Then
                                                '                        weight = "100"
                                                '                        sp_check = False
                                                '                    Else
                                                '                        weight = "50"
                                                '                        sp_check = False
                                                '                    End If
                                                '                End If
                                                '            End If

                                                '            If code_ = "de072" Then
                                                '                weight = "100"
                                                '                sp_check = False
                                                '            End If

                                                '            '1個１便
                                                '            If (code_ = "zk101-a" Or code_ = "zk101-b" Or code_ = "zk101-c" Or code_ = "zk101-d" Or code_ = "zk101-e" Or code_ = "zk101-f") Then
                                                '                weight = "100"
                                                '                sp_check = False
                                                '            End If

                                                '            If code_ = "ad105" Then
                                                '                weight = "20"
                                                '                sp_check = False
                                                '            End If

                                                '            If code_ = "ny261" Or code_ = "ny261-01" Or code_ = "ny261-02" Or code_ = "ny261-02a" Or code_ = "ny261-03" Or code_ = "ny263" Or code_ = "ny263-00a" Or code_ = "ny264" Then
                                                '                weight = "1.66"
                                                '                sp_check = False
                                                '            End If


                                                '            If code_ = "mask01" Or code_ = "mask-wh" Or code_ = "mask05" Or code_ = "mask01-bl" Or code_ = "mask01-ko" Then
                                                '                weight = "2.5"
                                                '                sp_check = False
                                                '            End If

                                                '            '重さ・特殊の処理
                                                '            Dim sp As String = DGV6.Item(dH6.IndexOf("特殊"), dgv6CodeArray.IndexOf(code_)).Value
                                                '            Dim omotoku As Double = 0
                                                '            If InStr(sp, "重") > 0 And sp_check Then
                                                '                omotoku = OmosaCheck(code_, 1)
                                                '                If omotoku <> 0 Then
                                                '                    haisouSize += omotoku * CDbl(count_)
                                                '                    haisouSize_info += omotoku.ToString + "[特] *" + count_maru + "+"
                                                '                End If
                                                '            ElseIf InStr(sp, "特") > 0 And sp_check Then
                                                '                omotoku = TokusyuCheck(code_.ToLower, 1)
                                                '                If omotoku <> 0 Then
                                                '                    haisouSize += omotoku * CDbl(count_)
                                                '                    haisouSize_info += omotoku.ToString + "[特] *" + count_maru + "+"
                                                '                End If
                                                '            Else
                                                '                If Not IsNumeric(weight) Then
                                                '                    Exit For
                                                '                End If
                                                '                If henkouFlag Then
                                                '                    If haisouKind = "メール便" Then
                                                '                        haisouSize += (CDbl(weight) / 100) * CDbl(count_)
                                                '                        haisouSize_info += (CDbl(weight) / 100).ToString + "[重] *" + count_maru + "+"
                                                '                    ElseIf haisouKind = "定形外" Then
                                                '                        haisouSize += (CDbl(weight) / 75) * CDbl(count_)
                                                '                        haisouSize_info += (CDbl(weight) / 75).ToString + "[重] *" + count_maru + "+"
                                                '                    End If
                                                '                Else
                                                '                    If w2 <> "" Then
                                                '                        haisouSize += CDbl(w2) * CDbl(count_)
                                                '                        haisouSize_info += CDbl(w2) + "[パーセンテージ] *" + count_maru + "+"
                                                '                    Else
                                                '                        If sp_check Then
                                                '                            haisouSize += CDbl(TakuhaiPerConv(weight)) * CDbl(count_)
                                                '                            haisouSize_info += CDbl(TakuhaiPerConv(weight)).ToString + "[重] *" + count_maru + "+"
                                                '                        Else
                                                '                            haisouSize += CDbl(weight) * CDbl(count_)
                                                '                            haisouSize_info += CDbl(weight).ToString + "[重] *" + count_maru + "+"
                                                '                        End If
                                                '                    End If
                                                '                End If
                                                '            End If
                                                '        End If
                                                '    End If
                                                'Next
                                                'haisouSize = Math.Ceiling(haisouSize)
                                                '指定した数以上の数のうち、最小の整数値を返します
                                                'DGV2.Item(dH2.IndexOf("サイズ計算"), r).Value = haisouSize
                                                'DGV2.Item(dH2.IndexOf("サイズ計算説明"), r).Value = haisouSize_info.Substring(0, haisouSize_info.Length - 1)

                                                If weightarray_.Length > 0 And place <> "" Then
                                                    Dim fare As Integer = 0
                                                    Dim fare_ As String = ""
                                                    For wa As Integer = 0 To weightarray_.Length - 1
                                                        Dim weight_str_ As String() = weightarray_(wa).Split("*")
                                                        '数量
                                                        Dim count_str_ As Integer = weight_str_(1)

                                                        If weight_str_(0) <> "計算不能" Then
                                                            Dim sw As String = weight_str_(0)
                                                            If Regex.IsMatch(sw, "P|p") Then
                                                                Dim weight As String = Regex.Replace(sw, "P|p", "")
                                                                weight = Regex.Replace(weight, "\(.*\)", "")
                                                                For s7 As Integer = 0 To DGV7.Rows.Count - 1
                                                                    Dim weight2 As Integer = weight
                                                                    Dim sw3 As String = HaisouSizeCheck(weight2)
                                                                    If place = DGV7.Item(dH7.IndexOf("都道府県"), s7).Value And DGV7.Item(dH7.IndexOf("サイズ"), s7).Value = sw3 Then
                                                                        fare += DGV7.Item(dH7.IndexOf("運賃"), s7).Value * count_str_
                                                                        fare_ += DGV7.Item(dH7.IndexOf("運賃"), s7).Value + "*" + count_str_.ToString + "+"
                                                                        Exit For
                                                                    End If
                                                                Next
                                                            End If
                                                        End If
                                                        DGV2.Item(dH2.IndexOf("運賃(計算)"), r).Value = fare
                                                        DGV2.Item(dH2.IndexOf("運賃差(計算 - 佐川)"), r).Style.BackColor = Color.White
                                                        If fare_ <> "" Then
                                                            DGV2.Item(dH2.IndexOf("運賃(計算説明)"), r).Value = fare_.Substring(0, fare_.Length - 1)
                                                            Dim fare_satu As Integer = fare - fare_sg
                                                            DGV2.Item(dH2.IndexOf("運賃差(計算 - 佐川)"), r).Value = fare_satu
                                                            If fare_satu < 0 Then
                                                                DGV2.Item(dH2.IndexOf("運賃差(計算 - 佐川)"), r).Style.BackColor = Color.Yellow
                                                            End If
                                                        End If
                                                    Next
                                                End If

                                                'If haisouSize > 0 Then
                                                '    Dim binsuu As Integer = Math.Round(haisouSize / 100)

                                                '    If binsuu = 0 Then
                                                '        DGV2.Item(dH2.IndexOf("便数計算"), r).Value = 1
                                                '        Dim haisouSize_noko As Integer = HaisouSizeHoko(haisouSize)
                                                '        If place <> "" Then
                                                '            For c7 As Integer = 0 To DGV7.Rows.Count - 1
                                                '                If place.Contains(DGV7.Item(dH7.IndexOf("都道府県"), r).Value.ToString) And DGV7.Item(dH7.IndexOf("サイズ"), r).Value = haisouSize_noko.ToString Then
                                                '                    DGV2.Item(dH2.IndexOf("運賃(計算)"), r).Value = DGV7.Item(dH7.IndexOf("運賃"), r).Value
                                                '                    DGV2.Item(dH2.IndexOf("運賃(計算説明)"), r).Value = "/"
                                                '                    Exit For
                                                '                End If
                                                '            Next
                                                '        End If
                                                '    Else
                                                '        DGV2.Item(dH2.IndexOf("便数計算"), r).Value = binsuu
                                                '        Dim haisouSize_noko As Integer = HaisouSizeHoko(haisouSize - (binsuu * 100))
                                                '        If place <> "" Then
                                                '            Dim haisouSize_noko_1 As Integer = 0 '残りの運賃
                                                '            Dim haisouSize_noko_2 As Integer = 0 '1便の運賃

                                                '            If haisouSize_noko <> 0 Then
                                                '                For c7 As Integer = 0 To DGV7.Rows.Count - 1
                                                '                    If place.Contains(DGV7.Item(dH7.IndexOf("都道府県"), r).Value) And DGV7.Item(dH7.IndexOf("サイズ"), r).Value = haisouSize_noko Then
                                                '                        haisouSize_noko_1 = DGV7.Item(dH7.IndexOf("運賃"), r).Value
                                                '                        Exit For
                                                '                    End If
                                                '                Next
                                                '            End If

                                                '            For c7 As Integer = 0 To DGV7.Rows.Count - 1
                                                '                If place.Contains(DGV7.Item(dH7.IndexOf("都道府県"), r).Value) And DGV7.Item(dH7.IndexOf("サイズ"), r).Value = "170" Then
                                                '                    haisouSize_noko_2 = DGV7.Item(dH7.IndexOf("運賃"), r).Value
                                                '                    Exit For
                                                '                End If
                                                '            Next

                                                '            If haisouSize_noko_2 <> 0 Then
                                                '                If haisouSize_noko = 0 Then
                                                '                    DGV2.Item(dH2.IndexOf("運賃(計算)"), r).Value = haisouSize_noko_2 * binsuu
                                                '                    DGV2.Item(dH2.IndexOf("運賃(計算説明)"), r).Value = haisouSize_noko_2.ToString + "*" + binsuu.ToString
                                                '                Else
                                                '                    DGV2.Item(dH2.IndexOf("運賃(計算)"), r).Value = haisouSize_noko_1 + (haisouSize_noko_2 * binsuu)
                                                '                    DGV2.Item(dH2.IndexOf("運賃(計算説明)"), r).Value = haisouSize_noko_1.ToString + "+" + haisouSize_noko_2.ToString + "*" + binsuu.ToString
                                                '                End If
                                                '            End If
                                                '        End If
                                                '    End If
                                                'End If
                                                DGV2.CurrentCell = DGV2(dH2.IndexOf("運賃差(計算 - 佐川)"), r)
                                            End If
                                            Exit For
                                        End If
                                    Next
                                End If
                            End If
                        Else
                            Exit Sub
                        End If
                    End If
                End If
            Else
                Exit Sub
            End If

            'If ToolStripProgressBar1.Value < ToolStripProgressBar1.Maximum Then
            '    ToolStripProgressBar1.Value += 1
            'End If
            Label1.Text = r + 1
            Application.DoEvents()
        Next
        MsgBox("処理が終わりました。")
    End Sub

    Dim omosa As Double = 0
    Private Function OmosaCheck(code As String, juchusu As String)
        Dim dH6 As ArrayList = TM_HEADER_GET(DGV6)
        '重さは30kgまでで1便（NumericUpDown1.Value）
        Dim weight As String = ""
        'Dim dH6 As ArrayList = TM_HEADER_GET(dgv6)
        Dim sp As String = DGV6.Item(dH6.IndexOf("特殊"), dgv6CodeArray.IndexOf(code)).Value
        omosa = Replace(sp, "重", "")
        weight = (Math.Floor(100 / NumericUpDown1.Value) * omosa) * juchusu
        Return weight
    End Function

    Dim tokuGroup As New ArrayList
    Private Function TokusyuCheck(code As String, juchusu As String)
        Dim dH6 As ArrayList = TM_HEADER_GET(DGV6)
        '商品特定して便を決め打ちする場合（括弧内は特定グループ）
        '特(ad033)0-0-0-0-0-1　　　1個で160を1便
        '特(ad033)0-0-1-2-0-0  　　1個なら100、2個なら120
        '特(ad033)0-0-2-4-0-0　2個まで100、3～4個は120
        '箱が140までなら追加で同梱が可能な処理が必要！！！！！！！！！！！！！！！！！

        Dim weight As Double = 0
        'Dim dH6 As ArrayList = TM_HEADER_GET(dgv6)
        Dim sizeArray As String() = New String() {"60", "80", "100", "120", "140", "160"}
        Dim sp As String = DGV6.Item(dH6.IndexOf("特殊"), dgv6CodeArray.IndexOf(code)).Value

        Dim toku As String = Replace(sp, "特", "")
        Dim rr As New Regex("\(.*\) Then", RegexOptions.IgnoreCase)
        Dim mm As Match = rr.Match(toku)
        Dim tGroup As String = Regex.Replace(mm.Value, "\(|\)", "") 'グループを取得する

        toku = Replace(toku, mm.Value, "")  '取得した残り
        Dim tokuArray As String() = Split(toku, "-")    '特別設定の中身

        '一番大きい入る個数を調べる
        Dim maxToku As Integer = 0
        For i As Integer = 0 To tokuArray.Length - 1
            If maxToku < tokuArray(i) Then
                maxToku = tokuArray(i)
            End If
        Next

        '---------------------------------------------------------------------
        'グループの処理
        'グループの商品があれば、個数を再計算用に取り出し、Arrayから削除する
        '---------------------------------------------------------------------
        Dim gWeight As Integer = 0
        Dim gKosu As Integer = 0
        For i As Integer = tokuGroup.Count - 1 To 0 Step -1
            Dim gCode As String() = Split(tokuGroup(i), "=")
            If gCode(0) = tGroup Then
                gWeight += gCode(1)
                gKosu += gCode(2)
                tokuGroup.RemoveAt(i)
            End If
        Next

        Dim kosu As Integer = 0
        If gWeight > 0 Then
            kosu = juchusu + gKosu
        Else
            kosu = juchusu
        End If

        '特殊設定の計算（新計算 2019/01/31）
        Dim hako As Integer = Math.Floor(kosu / maxToku)
        For i As Integer = 0 To hako - 1
            tokuGroup.Add(tGroup & "=100=" & maxToku)
        Next
        Dim nokori As Integer = kosu - (hako * maxToku)
        If nokori > 0 Then
            weight += (100 / maxToku) * nokori
            tokuGroup.Add(tGroup & "=" & weight & "=" & nokori)
        End If
        Dim weight1 As Double = (hako * 100) + ((100 / maxToku) * nokori)

        hako = Math.Floor((kosu - juchusu) / maxToku)
        nokori = (kosu - juchusu) - (hako * maxToku)
        Dim weight2 As Double = (hako * 100) + ((100 / maxToku) * nokori)

        'グループの個数で計算－今回計算の追加を引いた分＝差
        weight = weight1 - weight2

        '旧計算
        'While kosu > 0
        '    For i As Integer = 0 To tokuArray.Length - 1
        '        If tokuArray(i) <> 0 And tokuArray(i) >= kosu Or tokuArray(i) = maxToku Then
        '            'weight = weight + TakuhaiPerConv(sizeArray(j))
        '            weight += (100 / maxToku) * tokuArray(i)
        '            tokuGroup.Add(tGroup & "=" & weight & "=" & kosu)
        '            kosu = kosu - tokuArray(i)
        '            If kosu <= 0 Then
        '                Exit While
        '            End If
        '        End If
        '    Next
        'End While

        'チェック用
        'MsgBox(haisouSize & vbCrLf & tokuGroup.Count & vbCrLf & tokuGroup(tokuGroup.Count - 1) & vbCrLf & "kosu" & kosu & vbCrLf & "weight:" & weight & vbCrLf & "gWeight:" & gWeight & vbCrLf & "gKosu:" & gKosu)

        '---------------------------------------------------------------------
        'グループの処理
        '前の商品がグループの商品なら、個数を追加し再計算して、差を引く
        '---------------------------------------------------------------------
        'weight = weight - gWeight

        Return weight
    End Function

    '宅配便のサイズ別容量計算を読み込む
    Private Function TakuhaiPerConv(sizeStr As String) As String
        Dim dH10 As ArrayList = TM_HEADER_GET(DGV10)
        Dim res As String = ""
        For r As Integer = 0 To DGV10.RowCount - 1
            If Regex.IsMatch(DGV10.Item(dH10.IndexOf("サイズ"), r).Value, "^" & sizeStr & "$") Then
                res = DGV10.Item(dH10.IndexOf("容量計算"), r).Value
                Exit For
            End If
        Next
        Return res
    End Function

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

    'alertが出ているか確認する
    Private Function IsAlertPresent()
        Dim presentFlag As Boolean = False
        Try
            Dim alert As IAlert = cr.SwitchTo.Alert
            presentFlag = True
            alert.Accept()
        Catch ex As Exception
            'Alert Not present
        End Try
        Return presentFlag
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        td1.Abort()
        td1_action = False

        ChromeClose()

        Application.DoEvents()
    End Sub

    Private Function MaruMojiConv(chumonsu As String)
        '丸付き数字に変換する
        'Select Case chumonsu
        '    Case 1
        '        chumonsu = chumonsu
        '    Case 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19
        '        chumonsu = Chr(Asc("①") + chumonsu - 1)
        '    Case Else
        '        chumonsu = "(" & chumonsu & ")"
        'End Select
        chumonsu = Chr(Asc("①") + chumonsu - 1)
        Return chumonsu
    End Function

    Private Function HaisouSizeCheck(haisouSize As Integer) As String
        Dim result As String = ""
        If haisouSize > 240 Then
            result = 260
        ElseIf haisouSize > 220 And haisouSize <= 240 Then
            result = 240
        ElseIf haisouSize > 200 And haisouSize <= 220 Then
            result = 220
        ElseIf haisouSize > 180 And haisouSize <= 200 Then
            result = 200
        ElseIf haisouSize > 170 And haisouSize <= 180 Then
            result = 180
        ElseIf haisouSize > 160 And haisouSize <= 170 Then
            result = 170
        ElseIf haisouSize > 140 And haisouSize <= 160 Then
            result = 160
        ElseIf haisouSize > 100 And haisouSize <= 140 Then
            result = 140
        ElseIf haisouSize > 80 And haisouSize <= 100 Then
            result = 100
        ElseIf haisouSize > 60 And haisouSize <= 80 Then
            result = 80
        ElseIf haisouSize > 0 And haisouSize <= 60 Then
            result = 60
        End If

        Return result
    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If DGV2.Rows.Count = 0 Then
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
            sfd.FileName = "佐川送料チェック.csv"
        End If

        'ダイアログを表示する
        If sfd.ShowDialog(Me) = DialogResult.OK Then
            SaveCsv(sfd.FileName, 0, DGV2, True)
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

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim ss As Integer = Format(Now, "ss")   '秒指定して動かす（未使用）

        ToolStripStatusLabel1.Text = Timer1.Interval / 1000
        'sakiPath = Form1.サーバーToolStripMenuItem.Text & "\update"
        sakiPath2 = Form1.ロケーションToolStripMenuItem.Text

        Dim files2 As FileInfo() = Nothing

        Dim fi1 As New FileInfo(Path.GetDirectoryName(Form1.appPath) & "\location.csv")
        TextBox24.Text = fi1.LastWriteTime.ToOADate

        'サーバーの接続を確認する
        If File.Exists(sakiPath2 & "\location.csv") Then
            'sakiPath = "Z:\中村作成データ\AutoCheck\update"
            'Dim di As New System.IO.DirectoryInfo(sakiPath & "\version2")
            'files2 = di.GetFiles("*.dat", System.IO.SearchOption.AllDirectories)
            Dim fi2 As New FileInfo(sakiPath2 & "\location.csv")
            Dim check = fi2.LastAccessTime
            'KryptonButton3.Enabled = True
        Else
            Exit Sub
        End If

        Dim flag As Integer = False

        Try
            Dim fi2 As New FileInfo(sakiPath2 & "\location.csv")
            TextBox25.Text = fi2.LastWriteTime.ToOADate
            If TextBox25.Text < 0 Or TextBox24.Text < 0 Then
                TextBox17.Text = "error"
                Timer1.Interval = 1000 * 5
                ToolStripStatusLabel1.Text = Timer1.Interval / 1000
            ElseIf TextBox25.Text - TextBox24.Text > 0 Then
                Form1.NotifyIcon1.BalloonTipTitle = "配送マスタのアップデートを感知しました"
                Form1.NotifyIcon1.BalloonTipText = "自動でアップデートを行ないます。再起動の必要はありません。"
                Form1.NotifyIcon1.ShowBalloonTip(5000)

                TextBox17.Text = "update"
                Dim sPath As String = sakiPath2 & "\location.csv"
                Dim mPath As String = Path.GetDirectoryName(Form1.appPath) & "\location.csv"
                File.Copy(sPath, mPath, True)
                Timer1.Interval = 1000 * 5
                ToolStripStatusLabel1.Text = Timer1.Interval / 1000

                'datagridview6再表示
                DGV6_update()
                Application.DoEvents()
            Else
                If ToolStripStatusLabel1.Text <> 180 Then
                    Form1.NotifyIcon1.BalloonTipTitle = "配送マスタ状態は最新です"
                    Form1.NotifyIcon1.ShowBalloonTip(5000)
                End If

                TextBox17.Text = "OK"
                Timer1.Interval = 1000 * 60 * 3
                ToolStripStatusLabel1.Text = Timer1.Interval / 1000
                flag = True

                'datagridview6再表示
                DGV6_update()
                Application.DoEvents()
            End If
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If TextBox6.Text = "" Then
            Exit Sub
        End If

        For r As Integer = 0 To DGV6.RowCount - 1
            For c As Integer = 0 To DGV6.ColumnCount - 1
                If CheckBox2.Checked = False Then
                    If TextBox6.Text = DGV6.Item(c, r).Value Then
                        DGV6.CurrentCell = DGV6(c, r)
                        Exit Sub
                    End If
                Else
                    If InStr(DGV6.Item(c, r).Value, TextBox6.Text) > 0 Then
                        DGV6.CurrentCell = DGV6(c, r)
                        Exit Sub
                    End If
                End If
            Next
        Next
        MsgBox("検索結果がありません。", MsgBoxStyle.SystemModal)
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        For r As Integer = 0 To DGV6.RowCount - 1
            If DGV6.Rows(r).Visible = False Then
                DGV6.Rows(r).Visible = True
            End If
        Next
    End Sub

    Private Sub 例を削除ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 例を削除ToolStripMenuItem.Click
        Dim dgv As DataGridView = DGV2
        Dim selCell = dgv.SelectedCells
        COLSCUT(dgv, selCell)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        'If DGV1.Rows.Count = 0 Then
        '    MsgBox("データがないです。")
        '    Exit Sub
        'End If

        'Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)

        ''For c As Integer = 0 To dH1.Count - 1
        ''    Console.WriteLine(dH1(c))
        ''Next
        'If dH1.IndexOf("お客様管理№") = 0 Then
        '    MsgBox("「お客様管理№」フィールドがないです。")
        '    Exit Sub
        'End If

        'If dH1.IndexOf("都道府県") = 0 Then
        '    MsgBox("「都道府県」フィールドがないです。")
        '    Exit Sub
        'End If

        'Dim dh7 As ArrayList = TM_HEADER_GET(DGV7)
        'Dim Lst_cities As New List(Of String)()
        'Dim Lst_Chongfu As New List(Of String)()
        'For r As Integer = 0 To DGV7.Rows.Count - 1
        '    Dim city As String = DGV7.Item(dh7.IndexOf("都道府県"), r).Value
        '    If Not Lst_Chongfu.Contains(city) Then
        '        Lst_cities.Add(city)
        '    End If
        '    Lst_Chongfu.Add(city)
        'Next

        'For r As Integer = 0 To DGV1.Rows.Count - 1
        '    Dim city As String = DGV1.Item(dH1.IndexOf("都道府県"), r).Value
        '    If city <> "" Then
        '        If Not Lst_cities.Contains(city) Then
        '            MsgBox(r & "行目の都道府県が正しくないです。")
        '            Exit Sub
        '        End If
        '    End If
        'Next






    End Sub

    Private Sub DGV_RowPostPaint(sender As Object, e As DataGridViewRowPostPaintEventArgs) Handles DGV7.RowPostPaint, DGV6.RowPostPaint, DGV2.RowPostPaint, DGV1.RowPostPaint, DGV8.RowPostPaint, DGV9.RowPostPaint, DGV11.RowPostPaint, DGV12.RowPostPaint, DGV13.RowPostPaint
        ' 行ヘッダのセル領域を、行番号を描画する長方形とする（ただし右端に4ドットのすき間を空ける）
        Dim rect As New Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, sender.RowHeadersWidth - 4, sender.Rows(e.RowIndex).Height)

        ' 上記の長方形内に行番号を縦方向中央＆右詰で描画する　フォントや色は行ヘッダの既定値を使用する
        TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), sender.RowHeadersDefaultCellStyle.Font, rect, sender.RowHeadersDefaultCellStyle.ForeColor, TextFormatFlags.VerticalCenter Or TextFormatFlags.Right)
    End Sub

    Private Sub DGV8_DragDrop(sender As Object, e As DragEventArgs) Handles DGV8.DragDrop
        dataPathArray.Clear()

        Dim fCount As Integer = 0
        Dim filname_ As String = ""

        For Each filename As String In e.Data.GetData(DataFormats.FileDrop)
            Dim type = filename.Split(".")
            If type(type.Length - 1) <> "csv" Then
                MsgBox("csv対応のみです。")
                Exit Sub
            End If

            dataPathArray.Add(filename)
            fCount += 1
            filname_ += filename + ","
        Next

        If dataPathArray.Count > 1 Then
            'CSV_F_files.DGV1.Rows.Clear()
            'For i As Integer = 0 To dataPathArray.Count - 1
            '    CSV_F_files.DGV1.Rows.Add(True, Path.GetFileName(dataPathArray(i)))
            'Next
            'CSV_F_files.dataPathArray = dataPathArray
            'Dim DR As DialogResult = CSV_F_files.ShowDialog()
            'If DR = DialogResult.OK Then
            DropRun(dataPathArray, DGV8, fCount, headerAdd, addFlag)
            'Else
            '    Exit Sub
            'End If
        Else
            DropRun(dataPathArray, DGV8, fCount, True, True)
        End If
    End Sub

    Dim td2 As Thread

    Dim td2_action As Boolean = False
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If DGV8.Rows.Count = 0 Then
            MsgBox("データがないです。")
            Exit Sub
        End If

        Dim dH8 As ArrayList = TM_HEADER_GET(DGV8)

        If dH8.IndexOf("問合せ番号") < 0 Then
            MsgBox("「問合せ番号」フィールドがないです。")
            Exit Sub
        End If

        For r As Integer = 0 To DGV8.RowCount - 1
            'Console.WriteLine(DGV8.Item(dH8.IndexOf("問合せ番号"), r).Value)
            If DGV8.Item(dH8.IndexOf("問合せ番号"), r).Value <> "" Then
                If InStr(DGV8.Item(dH8.IndexOf("問合せ番号"), r).Value, "E+") > 0 Then
                    MsgBox((r + 1) & "行目の「問合せ番号」を確認してください。")
                    Exit Sub
                End If

            End If
        Next

        If IsNetworkAvailable() = False Then
            Exit Sub
        End If

        If cr Is Nothing Then
            ChromeStart()
        End If

        'action1()
        td2 = New Thread(AddressOf startThread2)
        td2.Start()
    End Sub

    Private Sub startThread2()
        If cr Is Nothing Then
            ChromeStart()
        End If

        td2_action = True

        Invoke(New CallDelegate(AddressOf action2))

        td2.Abort()
    End Sub

    Private Sub action2()
        DGV9.Rows.Clear()
        DGV9.Columns.Clear()

        Dim dH8 As ArrayList = TM_HEADER_GET(DGV8)
        For c As Integer = 0 To dH8.Count - 1
            DGV9.Columns.Add(c, dH8(c))
        Next

        DGV9.Columns.Add(dH8.Count + 1, "商品(ne)")
        DGV9.Columns.Add(dH8.Count + 2, "サイズ(発注)")

        Label7.Text = 0
        Label5.Text = DGV8.Rows.Count

        Dim dH6 As ArrayList = TM_HEADER_GET(DGV6)
        Dim dH9 As ArrayList = TM_HEADER_GET(DGV9)
        For r As Integer = 0 To DGV8.Rows.Count - 1
            If td2_action = True Then
                DGV9.Rows.Add(1)
                For c As Integer = 0 To dH8.Count - 1
                    DGV9.Item(dH9.IndexOf(dH8(c)), r).Value = DGV8.Item(dH8.IndexOf(dH8(c)), r).Value
                Next
                Application.DoEvents()


                Dim bongo As String = DGV8.Item(dH8.IndexOf("問合せ番号"), r).Value
                If bongo <> "" Then
                    Dim url As String = nextenginePath & "Userjyuchu"
                    cr.Url = url
                    cr.FindElementById("jyuchu_dlg_open").Click()

                    Application.DoEvents()
                    CR_SET_TEXT(cr.FindElementById("sea_jyuchu_search_field10"), bongo)
                    cr.FindElementById("ne_dlg_btn3_searchJyuchuDlg").Click()

                    While True
                        If cr.FindElementById("jyuchu_result_base").Text <> "" Then
                            Exit While
                        End If
                    End While
                    Application.DoEvents()

                    If cr.FindElementById("jyuchu_result_base").Text <> "結果はありませんでした" Then
                        Dim wb As WebBrowser = WebBrowser2
                        Dim html As String = cr.ExecuteScript("return document.body.outerHTML")
                        wb.ScriptErrorsSuppressed = True
                        wb.DocumentText = html
                        Application.DoEvents()

                        If InStr(wb.Document.Body.OuterHtml, "searchJyuchu_tablene_table") > 0 Then
                            Dim tb As HtmlElementCollection = wb.Document.GetElementsByTagName("table")

                            Dim denpyo_list As New ArrayList()
                            For Each he As HtmlElement In tb
                                If InStr(he.Id, "searchJyuchu_tablene_table") > 0 Then
                                    Dim tr_count As Integer = 0
                                    tr_count = he.GetElementsByTagName("tr").Count
                                    Application.DoEvents()

                                    If tr_count > 0 Then
                                        For c As Integer = 0 To tr_count - 2
                                            Dim denpyo_id As String = "searchJyuchu_table" & c & "-0"
                                            If wb.Document.GetElementById(denpyo_id) <> Nothing Then
                                                Dim denpyo As String = cr.FindElementById(denpyo_id).Text
                                                denpyo_list.Add(denpyo)
                                            End If
                                        Next
                                    End If
                                End If
                            Next
                            If denpyo_list.Count > 0 Then
                                Dim cd_str As String = ""
                                Dim weight_str As String = ""
                                For l As Integer = 0 To denpyo_list.Count - 1
                                    Dim url2 As String = nextenginePath & "Userjyuchu/jyuchuInp?kensaku_denpyo_no=" & denpyo_list(l) & "&jyuchu_meisai_order=jyuchu_meisai_gyo"
                                    'alertが残っているとエラーになるので対応


                                    cr.Url = url2
                                    Dim result As String = cr.FindElementById("jyuchu_msg").Text
                                    If InStr(result, "伝票番号の検索結果がありました") = 1 Then
                                        Dim wb2 As WebBrowser = WebBrowser2
                                        Dim html2 As String = cr.ExecuteScript("return document.body.outerHTML")
                                        wb2.ScriptErrorsSuppressed = True
                                        wb2.DocumentText = html2
                                        Application.DoEvents()

                                        If InStr(wb2.Document.Body.OuterHtml, "jyuchuMeisai_tablene_table") > 0 Then
                                            Dim tb2 As HtmlElementCollection = wb2.Document.GetElementsByTagName("table")

                                            For Each he As HtmlElement In tb2
                                                If InStr(he.Id, "jyuchuMeisai_tablene_table") > 0 Then
                                                    Dim td_count As Integer = 0
                                                    td_count = he.GetElementsByTagName("tr").Count
                                                    Application.DoEvents()

                                                    If td_count > 0 Then
                                                        Dim cd_temp As String = ""
                                                        'Dim doukon As Boolean = False
                                                        'Dim meisai_c As Integer = 0
                                                        For c As Integer = 0 To td_count - 2
                                                            'Dim cd As String = cr.FindElementById("jyuchuMeisai_table" & c & "-3").GetAttribute("value")
                                                            Dim code_id As String = "jyuchuMeisai_table" & c & "-3"
                                                            Dim jyucyu_count_id As String = "jyuchuMeisai_table" & c & "-6"
                                                            If wb2.Document.GetElementById(code_id) <> Nothing Then
                                                                Dim code As String = cr.FindElementById(code_id).Text.ToLower
                                                                Dim count As String = cr.FindElementById(jyucyu_count_id).Text
                                                                cd_str += code + "*" + count + ","

                                                                If Not dgv6CodeArray.Contains(code) Then
                                                                    weight_str += "計算不能" + "*" + count + ","
                                                                Else
                                                                    Dim weight As String = DGV6.Item(dH6.IndexOf("ship-weight"), dgv6CodeArray.IndexOf(code)).Value
                                                                    weight_str += weight + "*" + count + ","
                                                                End If
                                                                'meisai_c += 1
                                                            End If
                                                        Next
                                                    End If
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                    End If
                                Next

                                If cd_str <> "" Then
                                    cd_str = cd_str.Substring(0, cd_str.Length - 1)
                                    DGV9.Item(dH9.IndexOf("商品(ne)"), r).Value = cd_str
                                End If

                                If weight_str <> "" Then
                                    weight_str = weight_str.Substring(0, weight_str.Length - 1)
                                    DGV9.Item(dH9.IndexOf("サイズ(発注)"), r).Value = weight_str
                                End If

                                Application.DoEvents()
                                DGV9.CurrentCell = DGV9(dH9.IndexOf("商品(ne)"), r)
                            End If
                        End If
                    End If
                End If
            End If

            Label7.Text = r + 1
            Application.DoEvents()
        Next
        MsgBox("処理が終わりました。")
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        td2.Abort()
        td2_action = False

        ChromeClose()

        Application.DoEvents()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If DGV9.Rows.Count = 0 Then
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
            sfd.FileName = "佐川商品チェック.csv"
        End If

        'ダイアログを表示する
        If sfd.ShowDialog(Me) = DialogResult.OK Then
            SaveCsv(sfd.FileName, 0, DGV9, True)
            MsgBox(sfd.FileName & vbCrLf & "保存しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        If DGV11.Rows.Count = 0 Then
            MsgBox("データがないです。")
            Exit Sub
        End If

        If DGV12.Rows.Count = 0 Then
            MsgBox("明細データがないです。")
            Exit Sub
        End If

        Dim dH11 As ArrayList = TM_HEADER_GET(DGV11)

        If dH11.IndexOf("お客様管理№") = 0 Then
            MsgBox("「お客様管理№」フィールドがないです。")
            Exit Sub
        End If

        DGV13.Rows.Clear()
        DGV13.Columns.Clear()

        Button13.Enabled = False

        For c As Integer = 0 To dH11.Count - 1
            DGV13.Columns.Add(c, dH11(c))
        Next

        DGV13.Columns.Add(dH11.Count + 1, "商品(ne)")
        DGV13.Columns.Add(dH11.Count + 2, "サイズ(発注)")
        DGV13.Columns.Add(dH11.Count + 3, "運賃(計算)")
        DGV13.Columns.Add(dH11.Count + 4, "運賃(計算説明)")
        DGV13.Columns.Add(dH11.Count + 5, "運賃差(計算 - 佐川)")
        DGV13.Columns.Add(dH11.Count + 6, "サイズ(佐川)")

        Dim dH12 As ArrayList = TM_HEADER_GET(DGV12)
        Dim dH13 As ArrayList = TM_HEADER_GET(DGV13)

        Label16.Text = 0
        Label12.Text = DGV11.Rows.Count

        For r As Integer = 0 To DGV11.Rows.Count - 1
            DGV13.Rows.Add(1)
            For c As Integer = 0 To dH11.Count - 1
                DGV13.Item(dH13.IndexOf(dH11(c)), r).Value = DGV11.Item(dH11.IndexOf(dH11(c)), r).Value
            Next
            Application.DoEvents()

            Dim dH6 As ArrayList = TM_HEADER_GET(DGV6)
            Dim dH7 As ArrayList = TM_HEADER_GET(DGV7)
            Dim denpyono As String = DGV11.Item(dH11.IndexOf("お客様管理№"), r).Value
            Dim place As String = DGV11.Item(dH11.IndexOf("都道府県"), r).Value
            Dim fare_sg As Integer = DGV11.Item(dH11.IndexOf("運賃"), r).Value

            If denpyono <> "" Then
                If denpyono = Int(denpyono) Then
                    Dim cd_str = ""
                    Dim weight_str = ""
                    For r2 As Integer = 0 To DGV12.Rows.Count - 1
                        If denpyono = DGV12.Item(dH12.IndexOf("伝票番号"), r2).Value Then
                            Dim goods_cd As String = DGV12.Item(dH12.IndexOf("商品ｺｰﾄﾞ"), r2).Value
                            Dim count As String = DGV12.Item(dH12.IndexOf("引当数"), r2).Value
                            cd_str += goods_cd + "*" + count + ","

                            If Not dgv6CodeArray.Contains(goods_cd) Then
                                weight_str += "計算不能" + "*" + count + ","
                            Else
                                Dim weight As String = DGV6.Item(dH6.IndexOf("ship-weight"), dgv6CodeArray.IndexOf(goods_cd)).Value
                                weight_str += weight + "*" + count + ","
                            End If
                        End If
                    Next

                    If cd_str <> "" Then
                        cd_str = cd_str.Substring(0, cd_str.Length - 1)
                        DGV13.Item(dH13.IndexOf("商品(ne)"), r).Value = cd_str
                    End If

                    If weight_str <> "" Then
                        weight_str = weight_str.Substring(0, weight_str.Length - 1)
                        DGV13.Item(dH13.IndexOf("サイズ(発注)"), r).Value = weight_str
                    End If

                    Dim weightarray_ As String() = Split(weight_str, ",")

                    If weightarray_.Length > 0 And place <> "" Then
                        Dim fare As Integer = 0
                        Dim fare_ As String = ""
                        For wa As Integer = 0 To weightarray_.Length - 1
                            Dim weight_str_ As String() = weightarray_(wa).Split("*")
                            If weight_str_.Length > 1 And weight_str_(0) <> "計算不能" And weight_str_(0) <> "" Then
                                '数量
                                Dim count_str_ As Integer = weight_str_(1)

                                Dim sw As String = weight_str_(0)
                                If Regex.IsMatch(sw, "P|p") Then
                                    Dim weight As String = Regex.Replace(sw, "P|p", "")
                                    weight = Regex.Replace(weight, "\(.*\)", "")
                                    For s7 As Integer = 0 To DGV7.Rows.Count - 1
                                        Dim weight2 As Integer = weight
                                        Dim sw3 As String = HaisouSizeCheck(weight2)
                                        If place = DGV7.Item(dH7.IndexOf("都道府県"), s7).Value And DGV7.Item(dH7.IndexOf("サイズ"), s7).Value = sw3 Then
                                            fare += DGV7.Item(dH7.IndexOf("運賃"), s7).Value * count_str_
                                            fare_ += DGV7.Item(dH7.IndexOf("運賃"), s7).Value + "*" + count_str_.ToString + "+"
                                            Exit For
                                        End If
                                    Next
                                End If
                                DGV13.Item(dH13.IndexOf("運賃(計算)"), r).Value = fare
                                DGV13.Item(dH13.IndexOf("運賃差(計算 - 佐川)"), r).Style.BackColor = Color.White
                                If fare_ <> "" Then
                                    DGV13.Item(dH13.IndexOf("運賃(計算説明)"), r).Value = fare_.Substring(0, fare_.Length - 1)
                                    Dim fare_satu As Integer = fare - fare_sg
                                    DGV13.Item(dH13.IndexOf("運賃差(計算 - 佐川)"), r).Value = fare_satu
                                    If fare_satu < 0 Then
                                        DGV13.Item(dH13.IndexOf("運賃差(計算 - 佐川)"), r).Style.BackColor = Color.Yellow
                                    End If
                                End If
                            End If
                        Next
                    End If
                    If weightarray_.Length = 1 And place <> "" Then
                        Dim weight_str_ As String() = weightarray_(0).Split("*")
                        If weight_str_.Length > 1 And weight_str_(0) <> "計算不能" And weight_str_(0) <> "" Then
                            '数量
                            Dim count_str_ As Integer = weight_str_(1)

                            Dim sw As String = weight_str_(0)
                            If Regex.IsMatch(sw, "P|p") And count_str_ = 1 Then
                                Dim size_sagawa As String = ""
                                Dim inDgv7 As Integer = 0
                                Dim sagawa_fare_r As String = ""

                                For s7 As Integer = 0 To DGV7.Rows.Count - 1
                                    If place = DGV7.Item(dH7.IndexOf("都道府県"), s7).Value And DGV7.Item(dH7.IndexOf("運賃"), s7).Value = fare_sg Then
                                        size_sagawa = DGV7.Item(dH7.IndexOf("サイズ"), s7).Value
                                        inDgv7 += 1
                                        Exit For
                                    End If
                                Next
                                If inDgv7 = 0 Then
                                    Dim spareDgv7_price As ArrayList = New ArrayList

                                    For s7 As Integer = 0 To DGV7.Rows.Count - 1
                                        If place = DGV7.Item(dH7.IndexOf("都道府県"), s7).Value Then
                                            spareDgv7_price.Add(DGV7.Item(dH7.IndexOf("運賃"), s7).Value)
                                        End If
                                    Next

                                    ' https: //blog.csdn.net/weixin_41563161/article/details/104762344
                                    If spareDgv7_price.Count > 0 Then
                                        CountSumToTarget_list_2.Clear()
                                        CountSumToTarget(spareDgv7_price, fare_sg)
                                        Dim sagawa_fare_r2 As String = ""
                                        Dim sagawa_fare_r3 As String = ""
                                        If CountSumToTarget_list_2.Count > 0 Then
                                            For i As Integer = 0 To CountSumToTarget_list_2.Count - 1
                                                Dim CountSumToTarget_list As String() = CountSumToTarget_list_2(i).Split("||")
                                                If CountSumToTarget_list.Count > 0 Then
                                                    For i2 As Integer = 0 To CountSumToTarget_list.Count - 1
                                                        If CountSumToTarget_list(i2) <> "" Then
                                                            Dim CountSumToTarget_r As String() = CountSumToTarget_list(i2).Split(",")
                                                            For i3 As Integer = 0 To CountSumToTarget_r.Count - 1
                                                                If CountSumToTarget_r(i3) <> "" Then
                                                                    For s7 As Integer = 0 To DGV7.Rows.Count - 1
                                                                        If place = DGV7.Item(dH7.IndexOf("都道府県"), s7).Value And DGV7.Item(dH7.IndexOf("運賃"), s7).Value = CountSumToTarget_r(i3) Then
                                                                            sagawa_fare_r2 += "P" & DGV7.Item(dH7.IndexOf("サイズ"), s7).Value & "(" & CountSumToTarget_r(i3) & ")" & "+"
                                                                            Exit For
                                                                        End If
                                                                    Next
                                                                End If
                                                            Next
                                                        End If
                                                        If sagawa_fare_r2 <> "" Then
                                                            sagawa_fare_r3 += sagawa_fare_r2.Substring(0, sagawa_fare_r2.Length - 1) & ","
                                                        End If
                                                        sagawa_fare_r2 = ""
                                                    Next
                                                End If
                                            Next
                                            If sagawa_fare_r3 <> "" Then
                                                sagawa_fare_r += sagawa_fare_r3
                                            End If
                                        End If
                                    End If
                                    If sagawa_fare_r <> "" Then
                                        DGV13.Item(dH13.IndexOf("サイズ(佐川)"), r).Value = sagawa_fare_r.Substring(0, sagawa_fare_r.Length - 1)
                                    End If
                                End If
                                If sagawa_fare_r = "" Then
                                    If size_sagawa <> "" Then
                                        DGV13.Item(dH13.IndexOf("サイズ(佐川)"), r).Value = "P" & size_sagawa
                                    End If
                                End If
                            End If
                        End If
                    End If
                    DGV13.CurrentCell = DGV13(dH13.IndexOf("サイズ(佐川)"), r)
                End If
            End If
            Label16.Text = r + 1
            Application.DoEvents()
        Next
        MsgBox("処理が終わりました。")
        Button13.Enabled = True
    End Sub

    Private CountSumToTarget_list_2 As New ArrayList
    Private Function CountSumToTarget(data As ArrayList, target As Integer)
        Dim str As String = ""

        If data.Count = 0 Or target < 0 Or target = 0 Then
            Return str
        End If

        Dim list_ As ArrayList = New ArrayList


        CountSumToTarget_(list_, data, target, 0)
    End Function

    Private Function CountSumToTarget_(list As ArrayList, data As ArrayList, target As Integer, start As Integer)
        If target < 0 Then
            Exit Function
        End If
        If target = 0 Then
            Dim result As String = ""
            For i As Integer = 0 To list.Count - 1
                result += list(i) & ","
            Next
            CountSumToTarget_list_2.Add(result.Substring(0, result.Length - 1) & "||")
        Else
            For i As Integer = start To data.Count - 1
                If i <> start Then
                    If data(i) = data(i - 1) Then
                        Continue For
                    Else
                        list.Add(data(i))
                        CountSumToTarget_(list, data, target - data(i), i + 1)
                        list.RemoveAt(list.Count - 1)
                    End If
                Else
                    list.Add(data(i))
                    CountSumToTarget_(list, data, target - data(i), i + 1)
                    list.RemoveAt(list.Count - 1)
                End If
            Next
        End If
    End Function

    'Private Function MinSum(data As Integer()) As Integer
    '    Dim total As Integer = 0
    '    For r As Integer = 0 To data.Count - 1
    '        If data(r) < 0 Then
    '            total += data(r)
    '        End If
    '    Next
    '    Return total
    'End Function

    'Private Function MaxSum(data As Integer()) As Integer
    '    Dim total As Integer = 0
    '    For r As Integer = 0 To data.Count - 1
    '        If data(r) > 0 Then
    '            total += data(r)
    '        End If
    '    Next
    '    Return total
    'End Function

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        If DGV13.Rows.Count = 0 Then
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
            sfd.FileName = "佐川送料チェック.csv"
        End If

        'ダイアログを表示する
        If sfd.ShowDialog(Me) = DialogResult.OK Then
            SaveCsv(sfd.FileName, 0, DGV13, True)
            MsgBox(sfd.FileName & vbCrLf & "保存しました", MsgBoxStyle.SystemModal)
        End If
    End Sub
End Class