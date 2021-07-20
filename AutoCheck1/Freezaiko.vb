Imports System.IO
Imports System.Xml
Imports System.Net
Imports System.Text.RegularExpressions
Imports System.Text
Imports OpenQA.Selenium.Chrome
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Support
Imports System.Collections.ObjectModel
Imports OpenQA.Selenium.Interactions
Imports System.Threading
Imports Hnx8
Imports MySql.Data.MySqlClient



Public Class Freezaiko
    Private cr As ChromeDriver = Nothing
    Private sizeW As Integer = 500
    Private sizeH As Integer = 800
    Private path_cr As String = ""

    Public dataPathArray As New ArrayList
    Public headerAdd As Boolean = True
    Public addFlag As Boolean = True

    Dim conn As New MySqlConnection

    Private Sub フォームのリセットToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles フォームのリセットToolStripMenuItem.Click
        Me.Dispose()
        Dim frm As Form = New Freezaiko
        frm.Show()
    End Sub

    Private Sub Freezaiko_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        NumericUpDown1.Value = 1
        Label10.Text = "1"
        TextBox2.Text = ""

        WebBrowser1.Url = New Uri("http://www.benri.com/calendar/")

        TextBox1.Text = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments).Replace("Documents", "Downloads")
        'testTb_load()
        CheckBox1.Checked = True
        CheckBox2.Checked = True
        CheckBox3.Checked = True
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        If IsNetworkAvailable() = False Then
            MsgBox("ネットに接続できない。")
            Exit Sub
        End If

        'ListBox1.Items.Add("ネットに接続します")

        If cr Is Nothing Then
            ChromeStart()
        End If

        cr.Url = "https://ne107.next-engine.com/User_Syohin_Search"

        cr.FindElementById("btn_search_item").Click()
        cr.FindElementById("btn_search_item_clear").Click()
        SelectChangeMulti("dlg_search_syohin_kbn", "0 : 通常", "False", False)

        Application.DoEvents()

        cr.FindElementById("btn_search_item_exec").Click()
        Application.DoEvents()

        Threading.Thread.Sleep(5000)

        CR_INVOKE_LINK(cr, 5, "//*[@id='search_syouhin_result']/div[5]/div[2]/div[2]/button", False, "")
        CR_INVOKE_LINK(cr, 5, "//*[@id='search_syouhin_result']/div[5]/div[2]/div[2]/div/div[2]/div/p/div/button", False, "")

        Application.DoEvents()

        Threading.Thread.Sleep(10000)

        Dim dlFolder As String = TextBox1.Text
        Dim files As String() = Directory.GetFiles(dlFolder, "*.csv", SearchOption.AllDirectories)


        'Dim fdatenum1 As Integer = 0
        'Dim fdatenum2 As Integer = 0
        'Dim fName As String = ""
        Dim fName_all As String = ""
        'For i As Integer = 0 To files.Length - 1
        '    Dim fName_ As String = Path.GetFileName(files(i))

        '    If fName_.Contains("data") And fName_.Length = 30 Then
        '        Dim fName_all_ As String = files(i)
        '        Dim fdatenum_ As String = fName_.Replace("data", "").Replace(".csv", "").Substring(0, 15)

        '        If IsNumeric(fdatenum_) Then
        '            Dim fdatenum3 As Integer = CInt(fdatenum_.Substring(0, 8))
        '            Dim fdatenum4 As Integer = CInt(fdatenum_.Substring(8, 7))

        '            If (fdatenum1 = 0 And fdatenum2 = 0) Or (fdatenum3 > fdatenum1 And fdatenum4 > fdatenum2) Then
        '                fdatenum1 = fdatenum3
        '                fdatenum2 = fdatenum4
        '                fName = fName_
        '                fName_all = fName_all_
        '            Else
        '                '同じ日
        '                If fdatenum3 = fdatenum1 Then
        '                    If fdatenum4 > fdatenum2 Then
        '                        fdatenum1 = fdatenum3
        '                        fdatenum2 = fdatenum4
        '                        fName = fName_
        '                        fName_all = fName_all_
        '                    End If
        '                ElseIf fdatenum3 > fdatenum1 Then
        '                    fdatenum1 = fdatenum3
        '                    fdatenum2 = fdatenum4
        '                    fName = fName_
        '                    fName_all = fName_all_
        '                End If
        '            End If
        '        End If
        '    End If
        'Next


        Dim newfile As List(Of String) = New List(Of String)()
        For index = 0 To files.Length - 1

            Dim cc = Path.GetFileName(files(index)).Length
            If files(index).Contains("data") And Path.GetFileName(files(index)).Length = 30 Then
                newfile.Add(files(index))
            End If
        Next

        Array.Sort(newfile.ToArray())

        fName_all = newfile(newfile.Count - 1)

        If fName_all <> "" Then
            TextBox2.Text = fName_all

            DGV1.Rows.Clear()
            DGV1.Columns.Clear()
            DropRun(fName_all, DGV1, True, True)
        Else
            MsgBox("ファイルが存在しないために処理が終わりです")
            Exit Sub
        End If

        DGV2.Rows.Clear()
        DGV2.Columns.Clear()
    End Sub

    '共通
    Private Function IsNetworkAvailable()
        Dim IsNetworkAvailable_Flag As Boolean = False
        If NetworkInformation.NetworkInterface.GetIsNetworkAvailable() Then
            ToolStripStatusLabel1.Text = "接続"
            ToolStripStatusLabel1.BackColor = Color.LightGreen
            IsNetworkAvailable_Flag = True
        Else
            ToolStripStatusLabel1.Text = "未接続"
            ToolStripStatusLabel1.BackColor = Color.Transparent
        End If
        Return IsNetworkAvailable_Flag
    End Function

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
        ChromeOptions.BinaryLocation = ""


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

    ''' <summary>
    ''' HTMLリンク・ボタンを押下
    ''' </summary>
    ''' <param name="cr">ChromeDriver</param>
    ''' <param name="tag">0=name、1=id、2=class、3=CssSelector、4=LinkText、5=XPath</param>
    ''' <param name="tagName">String</param>
    ''' <param name="dialogFlag">ダイアログを待機するか（defalt=false）</param>
    ''' <param name="alertDismiss">alertキャンセル</param>
    ''' <param name="alertAccept">alertOK</param>
    Public Sub CR_INVOKE_LINK(cr As ChromeDriver, tag As Integer, tagName As String, Optional dialogFlag As Boolean = False, Optional alertDismiss As String = "", Optional alertAccept As String = "")
        Try
            Dim element As IWebElement = Nothing
            Select Case tag
                Case "0"
                    element = cr.FindElementByName(tagName)
                Case "1"
                    element = cr.FindElementById(tagName)
                Case "2"
                    element = cr.FindElementByClassName(tagName)
                Case "3"
                    element = cr.FindElementByCssSelector(tagName)
                Case "4"
                    element = cr.FindElementByLinkText(tagName)
                Case "5"
                    element = cr.FindElementByXPath(tagName)
                Case Else
                    element = cr.FindElementByName(tagName)
            End Select
            element.Click()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        If dialogFlag Then
            For i As Integer = 0 To 3
                Try
                    Application.DoEvents()
                    Threading.Thread.Sleep(1000)
                    Dim html As String = cr.ExecuteScript("return document.body.outerHTML")
                Catch ex As UnhandledAlertException
                    Dim alert As IAlert = cr.SwitchTo.Alert
                    If Regex.IsMatch(alert.Text, alertDismiss) Then
                        alert.Dismiss()
                    ElseIf Regex.IsMatch(alert.Text, alertAccept) Then
                        alert.Accept()
                    Else
                        alert.Accept()
                    End If
                Catch ex As Exception

                End Try
            Next
        End If
    End Sub

    'selectを選択
    Function FindSelectElement(ByVal driver As IWebDriver, ByVal by As By) As Support.UI.SelectElement
        Dim e = driver.FindElement(by)
        Return New Support.UI.SelectElement(e)
    End Function

    ''' <summary>
    ''' 複数選択（select）
    ''' </summary>
    ''' <param name="id">セレクタ（id）</param>
    ''' <param name="values">選択肢テキスト「|区切り」</param>
    ''' <param name="op">True=一旦選択解除する</param>
    Private Sub SelectChangeMulti(id As String, values As String, op As String, Optional timerUse As Boolean = True)
        Dim se As UI.SelectElement = FindSelectElement(cr, By.Id(id))
        If op = "True" Then
            se.DeselectAll()
        End If
        Thread.Sleep(2000)
        Dim valueArray As String() = Split(values, "|")
        For i As Integer = 0 To valueArray.Length - 1
            se.SelectByText(valueArray(i))
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub

    Public Sub DropRun(fName As String, dgv As DataGridView, Optional headerAdd As Boolean = False, Optional addFlag As Boolean = True)
        Dim num As Integer = 0
        Dim colorArray As Color() = {Color.LightYellow, Color.LightCyan}

        Dim csvRecords As New ArrayList()
        csvRecords = TM_CSV_READ2(fName)

        For r As Integer = 0 To csvRecords.Count - 1
            Dim sArray As String() = Split(csvRecords(r), "|=|")
            If dgv.ColumnCount < sArray.Length - 1 Then
                For c As Integer = dgv.ColumnCount To sArray.Length - 1
                    dgv.Columns.Add(c, sArray(c))
                    dgv.Columns(c).SortMode = DataGridViewColumnSortMode.NotSortable
                Next
            End If

            If r > 0 Then
                dgv.Rows.Add(sArray)
            End If

            If num > 0 Then
                Dim cNum As Integer = num Mod 2
                dgv.Item(0, dgv.RowCount - 2).Style.BackColor = colorArray(cNum) '追加した時だけ、一時的に色を付ける
            End If
        Next
    End Sub

    Private Sub DGV_RowPostPaint(sender As Object, e As DataGridViewRowPostPaintEventArgs) Handles DGV1.RowPostPaint, DGV2.RowPostPaint, DGV3.RowPostPaint, DGV5.RowPostPaint, DGV6.RowPostPaint
        ' 行ヘッダのセル領域を、行番号を描画する長方形とする（ただし右端に4ドットのすき間を空ける）
        Dim rect As New Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, sender.RowHeadersWidth - 4, sender.Rows(e.RowIndex).Height)

        ' 上記の長方形内に行番号を縦方向中央＆右詰で描画する　フォントや色は行ヘッダの既定値を使用する
        TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), sender.RowHeadersDefaultCellStyle.Font, rect, sender.RowHeadersDefaultCellStyle.ForeColor, TextFormatFlags.VerticalCenter Or TextFormatFlags.Right)
    End Sub

    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs)
        Label10.Text = sender.value
    End Sub

    Private Sub 行を削除ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 行を削除ToolStripMenuItem1.Click
        Dim dgv As DataGridView = DGV2
        Dim selCell = dgv.SelectedCells
        ROWSCUT(dgv, selCell)
    End Sub

    Private Sub 処理ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 処理ToolStripMenuItem.Click
        Dim dgv As DataGridView = DGV1
        If dgv.Rows.Count = 0 Then
            MsgBox("データがないために処理されないです")
            Exit Sub
        End If

        DGV2.Rows.Clear()
        DGV2.Columns.Clear()

        'Dim rm_other As Integer()
        'Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
        'Dim del_lie As Integer = 0
        'For i As Integer = 0 To dH1.Count - 1
        '    If dH1(i) = "商品コード" Or dH1(i) = "商品名" Or dH1(i) = "引当数" Or dH1(i) = "フリー在庫" Then
        '    Else
        '        DGV1.Columns.RemoveAt(i - del_lie)
        '        del_lie = del_lie + 1
        '    End If
        'Next

        'dH1 = TM_HEADER_GET(DGV1)
        'Dim del_row As Integer = 0
        'For r As Integer = 0 To DGV1.RowCount - 1
        '    Dim rc As Integer = DGV1.RowCount - 1
        '    If r = rc Then
        '        Exit For
        '    End If

        '    If DGV1.Item(dH1.IndexOf("引当数"), r).Value <= DGV1.Item(dH1.IndexOf("フリー在庫"), r).Value Then
        '        DGV1.Rows.RemoveAt(r - del_row)
        '        del_lie = del_lie + 1
        '    End If
        'Next

        Dim dH As ArrayList = TM_HEADER_GET(dgv)

        Dim has_syohin As Boolean = False


        DGV2.Columns.Add(0, "商品コード")


        If dH.Contains("商品名") Then
            DGV2.Columns.Add(1, "商品名")
            has_syohin = True

            DGV2.Columns.Add(2, "引当数")
            DGV2.Columns.Add(3, "フリー在庫")
            DGV2.Columns.Add(4, "引当数*2")
            DGV2.Columns.Add(5, "引当数*3")
        Else
            DGV2.Columns.Add(1, "引当数")
            DGV2.Columns.Add(2, "フリー在庫")
            DGV2.Columns.Add(3, "引当数*2")
            DGV2.Columns.Add(4, "引当数*3")
        End If

        Dim dH2 As ArrayList = TM_HEADER_GET(DGV2)

        Dim add_row As Integer = 0
        For r As Integer = 0 To dgv.RowCount - 1
            If dgv.Item(dH.IndexOf("引当数"), r).Value = 0 Or dgv.Item(dH.IndexOf("引当数"), r).Value = "" Then
            Else
                If RadioButton1.Checked Then
                    If dgv.Item(dH.IndexOf("フリー在庫"), r).Value > (dgv.Item(dH.IndexOf("引当数"), r).Value * 2) Then
                        Continue For
                    End If
                End If
                If RadioButton2.Checked Then
                    If dgv.Item(dH.IndexOf("フリー在庫"), r).Value > (dgv.Item(dH.IndexOf("引当数"), r).Value * 3) Then
                        Continue For
                    End If
                End If
                If CheckBox2.Checked = True And CheckBox3.Checked = True Then
                    If (dgv.Item(dH.IndexOf("引当数"), r).Value <= dgv.Item(dH.IndexOf("フリー在庫"), r).Value) Or (dgv.Item(dH.IndexOf("フリー在庫"), r).Value < 10) Then
                        DGV2.Rows.Add(1)
                        DGV2.Item(dH2.IndexOf("商品コード"), add_row).Value = dgv.Item(dH.IndexOf("商品コード"), r).Value

                        Dim hikiate As Integer = CInt(dgv.Item(dH.IndexOf("引当数"), r).Value)
                        DGV2.Item(dH2.IndexOf("引当数"), add_row).Value = hikiate
                        DGV2.Item(dH2.IndexOf("フリー在庫"), add_row).Value = dgv.Item(dH.IndexOf("フリー在庫"), r).Value

                        If has_syohin Then
                            DGV2.Item(dH2.IndexOf("商品名"), add_row).Value = dgv.Item(dH.IndexOf("商品名"), r).Value
                        End If

                        DGV2.Item(dH2.IndexOf("引当数*2"), add_row).Value = hikiate * 2
                        DGV2.Item(dH2.IndexOf("引当数*3"), add_row).Value = hikiate * 3

                        add_row = add_row + 1
                    End If
                ElseIf CheckBox2.Checked = True And CheckBox3.Checked = False Then
                    If dgv.Item(dH.IndexOf("フリー在庫"), r).Value < 10 Then
                        DGV2.Rows.Add(1)
                        DGV2.Item(dH2.IndexOf("商品コード"), add_row).Value = dgv.Item(dH.IndexOf("商品コード"), r).Value

                        Dim hikiate As Integer = CInt(dgv.Item(dH.IndexOf("引当数"), r).Value)
                        DGV2.Item(dH2.IndexOf("引当数"), add_row).Value = hikiate
                        DGV2.Item(dH2.IndexOf("フリー在庫"), add_row).Value = dgv.Item(dH.IndexOf("フリー在庫"), r).Value

                        If has_syohin Then
                            DGV2.Item(dH2.IndexOf("商品名"), add_row).Value = dgv.Item(dH.IndexOf("商品名"), r).Value
                        End If

                        DGV2.Item(dH2.IndexOf("引当数*2"), add_row).Value = hikiate * 2
                        DGV2.Item(dH2.IndexOf("引当数*3"), add_row).Value = hikiate * 3

                        add_row = add_row + 1
                    End If
                ElseIf CheckBox2.Checked = False And CheckBox3.Checked = True Then
                    If dgv.Item(dH.IndexOf("引当数"), r).Value <= dgv.Item(dH.IndexOf("フリー在庫"), r).Value Then
                        DGV2.Rows.Add(1)
                        DGV2.Item(dH2.IndexOf("商品コード"), add_row).Value = dgv.Item(dH.IndexOf("商品コード"), r).Value

                        Dim hikiate As Integer = CInt(dgv.Item(dH.IndexOf("引当数"), r).Value)
                        DGV2.Item(dH2.IndexOf("引当数"), add_row).Value = hikiate
                        DGV2.Item(dH2.IndexOf("フリー在庫"), add_row).Value = dgv.Item(dH.IndexOf("フリー在庫"), r).Value

                        If has_syohin Then
                            DGV2.Item(dH2.IndexOf("商品名"), add_row).Value = dgv.Item(dH.IndexOf("商品名"), r).Value
                        End If

                        DGV2.Item(dH2.IndexOf("引当数*2"), add_row).Value = hikiate * 2
                        DGV2.Item(dH2.IndexOf("引当数*3"), add_row).Value = hikiate * 3

                        add_row = add_row + 1
                    End If
                ElseIf CheckBox2.Checked = False And CheckBox3.Checked = False Then
                    DGV2.Rows.Add(1)
                    DGV2.Item(dH2.IndexOf("商品コード"), add_row).Value = dgv.Item(dH.IndexOf("商品コード"), r).Value

                    Dim hikiate As Integer = CInt(dgv.Item(dH.IndexOf("引当数"), r).Value)
                    DGV2.Item(dH2.IndexOf("引当数"), add_row).Value = hikiate
                    DGV2.Item(dH2.IndexOf("フリー在庫"), add_row).Value = dgv.Item(dH.IndexOf("フリー在庫"), r).Value

                    If has_syohin Then
                        DGV2.Item(dH2.IndexOf("商品名"), add_row).Value = dgv.Item(dH.IndexOf("商品名"), r).Value
                    End If

                    DGV2.Item(dH2.IndexOf("引当数*2"), add_row).Value = hikiate * 2
                    DGV2.Item(dH2.IndexOf("引当数*3"), add_row).Value = hikiate * 3

                    add_row = add_row + 1
                End If
            End If
        Next

        MsgBox("処理しました")
    End Sub

    Private Sub testTb_load()
        Dim dlFolder As String = TextBox1.Text
        Dim files As String() = Directory.GetFiles(dlFolder, "*.csv", SearchOption.AllDirectories)

        Dim fdatenum1 As Integer = 0
        Dim fdatenum2 As Integer = 0
        Dim fName As String = ""
        Dim fName_all As String = ""
        For i As Integer = 0 To files.Length - 1
            Dim fName_ As String = Path.GetFileName(files(i))

            If fName_.Contains("data") And fName_.Length = 30 Then
                Dim fName_all_ As String = files(i)
                Dim fdatenum_ As String = fName_.Replace("data", "").Replace(".csv", "").Substring(0, 15)

                If IsNumeric(fdatenum_) Then
                    Dim fdatenum3 As Integer = CInt(fdatenum_.Substring(0, 8))
                    Dim fdatenum4 As Integer = CInt(fdatenum_.Substring(8, 7))

                    If (fdatenum1 = 0 And fdatenum2 = 0) Or (fdatenum3 > fdatenum1 And fdatenum4 > fdatenum2) Then
                        fdatenum1 = fdatenum3
                        fdatenum2 = fdatenum4
                        fName = fName_
                        fName_all = fName_all_
                    Else
                        '同じ日
                        If fdatenum3 = fdatenum1 Then
                            If fdatenum4 > fdatenum2 Then
                                fdatenum1 = fdatenum3
                                fdatenum2 = fdatenum4
                                fName = fName_
                                fName_all = fName_all_
                            End If
                        ElseIf fdatenum3 > fdatenum1 Then
                            fdatenum1 = fdatenum3
                            fdatenum2 = fdatenum4
                            fName = fName_
                            fName_all = fName_all_
                        End If
                    End If
                End If
            End If
        Next

        If fName_all <> "" Then
            TextBox2.Text = fName_all

            DGV1.Rows.Clear()
            DGV1.Columns.Clear()
            DropRun(fName_all, DGV3, True, True)
        Else
            MsgBox("ファイルが存在しないために処理が終わりです")
            Exit Sub
        End If
    End Sub

    Private Sub 列を削除ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 列を削除ToolStripMenuItem.Click
        Dim dgv As DataGridView = DGV2
        Dim selCell = dgv.SelectedCells
        COLSCUT(dgv, selCell)
    End Sub

    Private Sub エクスポートToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles エクスポートToolStripMenuItem.Click
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

        sfd.FileName = "予約変更" & Format(Now, "yyyyMMddHHmmss") & ".csv"

        'ダイアログを表示する
        If sfd.ShowDialog(Me) = DialogResult.OK Then
            SaveCsv(sfd.FileName, 0, DGV2, "SHIFT-JIS")
            MsgBox(sfd.FileName & vbCrLf & "保存しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    Private Sub DGV4_DragDrop(sender As Object, e As DragEventArgs) Handles DGV4.DragDrop
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
            DropRun(dataPathArray, sender, fCount, headerAdd, addFlag)
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

    Dim sc As Integer = 0
    Dim sr As Integer = 0

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
        'DataGridView1.ScrollBars = ScrollBars.None
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
        dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        dgv.ScrollBars = ScrollBars.Both
        '-----------------------------------------------------------------------
    End Sub

    Private Sub DGV4_DragEnter(sender As Object, e As DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        DGV5.Rows.Clear()
        DGV5.Columns.Clear()

        If DGV4.Rows.Count = 0 Then
            MsgBox("ファイルを入れてください。")
            Exit Sub
        End If

        Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
        If dH1.Contains("商品コード") Then
            MsgBox("元データの「商品コード」列がないです。")
            Exit Sub
        End If

        Dgv6_load()

        If DGV6.Rows.Count = 0 Then
            MsgBox("発注管理のデータがないために処理されないです。")
            Exit Sub
        End If

        Dim dH4 As ArrayList = TM_HEADER_GET(DGV4)
        For r As Integer = 0 To DGV4.Rows.Count - 1
            Dim code As String = DGV4.Item(dH4.IndexOf("商品コード"), r).Value
            Console.WriteLine(code)

        Next





    End Sub

    ' 执行查询的操作,    
    ' <param name="cmdText">需要执行语句,一般是Sql语句,也有存储过程</param>    
    ' <param name="cmdType">判断Sql语句的类型,一般都不是存储过程</param>    
    ' <returns>dataTable,查询到的表格</returns>
    Public Function ExecSelectNo(ByVal cmdText As String) As DataTable

        Dim sqlAdapter As New MySqlDataAdapter(cmdText, conn)
        Dim ds As New DataSet

        Try
            sqlAdapter.Fill(ds)           '用adapter将dataSet填充     
            Return ds.Tables(0)             'datatable为dataSet的第一个表   

        Catch ex As Exception
            Return Nothing
        Finally
            conn.Close()
        End Try
    End Function

    ' 执行增删改三个操作,
    ' <param name="cmdText">需要执行语句,一般是Sql语句,也有存储过程</param>    
    ' <param name="cmdType">判断Sql语句的类型,一般都不是存储过程</param>    
    Public Function ExecAddDelUpdate(ByVal cmdText As String, ByVal cmdType As CommandType) As Integer

        Dim myCommand As MySqlCommand = New MySqlCommand(cmdText, conn)
        Try
            conn.Open()                      '打开连接    
            Return myCommand.ExecuteNonQuery()
        Catch ex As Exception
            Return 0                         '如果出错,返回0    
        Finally
            conn.Close()
        End Try
    End Function


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dgv6_load()
    End Sub

    Private Sub Dgv6_load()
        ' 1.接続文字列を作成する
        Dim Builder = New MySqlConnectionStringBuilder()

        ' データベースに接続するために必要な情報をBuilderに与える
        Builder.Server = "192.168.12.178"
        Builder.Port = 3306
        Builder.UserID = "root"
        Builder.Password = ""
        Builder.Database = "sddb0040100537"
        Dim ConStr = Builder.ToString()

        ' 3.発行するSQL文を作成する
        conn.ConnectionString = ConStr
        conn.Open()

        Dim query As String = "SELECT ID,`code`,`sub-code`,`order-count`,`inspect-count`,`arrival-flag`,container FROM `list1` WHERE (`lockuser` NOT IN('del') OR `lockuser` IS NULL);"
        ' 4.データ取得のためのアダプタの設定
        Dim Adapter = New MySqlDataAdapter(query, conn)

        ' 5.データを取得し、セットする
        Dim Ds As New DataSet
        Adapter.Fill(Ds)
        ' 6.DataGridViewに取得したデータを表示させる
        DGV6.Rows.Clear()
        DGV6.Columns.Clear()

        DGV6.DataSource = Ds.Tables(0)
    End Sub

End Class