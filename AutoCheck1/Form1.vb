Imports System.Runtime.InteropServices
Imports System.Security.Permissions
Imports System.IO
Imports Microsoft.VisualBasic.FileIO
Imports System.Text.RegularExpressions
Imports System.IO.Compression
Imports System.Text
Imports System.Threading
Imports System.ComponentModel
Imports System.Net

Public Class Form1
    '"\HKEY_CURRENT_USER\SOFTWARE\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION"
    'AutoCheck1.exeとAutoCheck1.vshost.exeキーにDword値32bit
    '11001	0x2AF9	Internet Explorer 11, Edgeモード (最新のバージョンでレンダリング)
    '11000	0x2AF8	Internet Explorer 11
    '10001	0x2711	Internet Explorer 10, Standardsモード
    '10000	0x2710	Internet Explorer 10 (!DOCTYPE で指定がある場合は、Standardsモードになります。)
    '9999	0x270F	Internet Explorer 9, Standardsモード
    '9000	0x2710	Internet Explorer 9 (!DOCTYPE で指定がある場合は、Standardsモードになります。)
    '8888	0x22B8	Internet Explorer 8, Standardsモード
    '8000	0x1F40	Internet Explorer 8 (!DOCTYPE で指定がある場合は、Standardsモードになります。)
    '7000	0x1B58	Internet Explorer 7 (!DOCTYPE で指定がある場合は、Standardsモードになります。)
    '[HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_SCRIPTURL_MITIGATION]
    '名前 ChaRa.exe
    'DWORD 値(10進) 1

    '-------------------------------------------------------------------
    Public Shared AdminFlag As Boolean = False
    Dim secValue As String = ""
    Public Shared appPath As String = Reflection.Assembly.GetExecutingAssembly().Location
    Dim bookmarkURL As New ArrayList
    Public Shared navCount As Integer = 0

    Public Shared inif As inifile.inifileAccess
    Public Shared miniForm As Boolean = False

    Public WithEvents TabBrowser1 As New WebTabControl
    'Dim StatusStrip1 As New StatusStrip
    Dim ToolStripStatusLabel1 As New ToolStripStatusLabel
    'Dim ToolStripProgressBar1 As New ToolStripProgressBar
    Private checkItem() As CheckBox = New CheckBox() {}
    Private radioItemCheck() As RadioButton = New RadioButton() {}
    Public Shared enc As Encoding = Encoding.GetEncoding(932)

    Public Shared csv_f1 As Form = New Csv
    Public Shared csv_f2 As Form = New Csv
    Public Shared CSV_FORMS As Csv() = New Csv() {csv_f1, csv_f2}

    Dim checkTime As String = ""    'アップデートのチェック時間
    'Dim sakiPath As String = Path.GetDirectoryName(Form1.appPath) & "\AutoCheck1_bak.exe"

    Dim hotkeyID As Short



    '##################################################################
    ' Error処理
    '##################################################################
    Public Sub New()
        InitializeComponent()
        AddHandler Application.ThreadException, AddressOf Application_ThreadException
        AddHandler Threading.Thread.GetDomain().UnhandledException, AddressOf Application_UnhandledException

        '実装できない（未完成）
        'Dim pfc As New Text.PrivateFontCollection()
        'Dim appPath1 As String = Reflection.Assembly.GetExecutingAssembly().Location
        'pfc.AddFontFile(Path.GetDirectoryName(appPath1) & "\config\msgothic.ttc")
        'MyBase.Font = New System.Drawing.Font(pfc.Families(0), 9) ', GraphicsUnit.Point, 128)
        'pfc.Dispose()
    End Sub

    Public Shared Sub Application_ThreadException(ByVal sender As Object, ByVal e As ThreadExceptionEventArgs)
        Dim ex As Exception = CType(e.Exception, Exception)
        ShowErrorMessage(ex, "Application.ThreadExceptionによる例外通知です。")
    End Sub

    Public Shared Sub Application_UnhandledException(ByVal sender As Object, ByVal e As UnhandledExceptionEventArgs)
        Dim ex As Exception = CType(e.ExceptionObject, Exception)
        If Not ex Is Nothing Then
            ShowErrorMessage(ex, "Application.UnhandledExceptionによる例外通知です。")
        End If
    End Sub

    Public Shared Sub ShowErrorMessage(ByVal ex As Exception, ByVal extraMessage As String)
        CaptureDesktop()
        ErrorDialog.TextBox5.Text = ALIO_ERROR
        ErrorDialog.TextBox3.Text = "Application.ThreadExceptionによる例外通知です。"
        ErrorDialog.TextBox2.Text = ex.Message
        ErrorDialog.TextBox6.Text = ex.Message
        ErrorDialog.TextBox1.Text = ex.StackTrace
        ErrorDialog.TextBox4.Text = ex.Source
        Dim str As String = "" '"<InnerException>" & vbCrLf & ex.InnerException.ToString
        str &= vbCrLf & vbCrLf
        str &= "<StackTrace>" & vbCrLf & ex.StackTrace.ToString
        str &= vbCrLf & vbCrLf
        str &= "<TargetSite>" & vbCrLf & ex.TargetSite.ToString
        ErrorDialog.TextBox7.Text = str
        Application.DoEvents()
        If updateFlag = True Then
            ErrorDialog.MailSend()
        End If
        ErrorDialog.Height = 145
        ErrorDialog.ShowDialog()
    End Sub

    Private Shared Sub CaptureDesktop()
        Dim p As New Point(0, 0)
        Dim s As New Size(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height)
        Dim bmp As New Bitmap(s.Width, s.Height)
        Dim g As Graphics = Graphics.FromImage(bmp)
        g.CopyFromScreen(p, New Point(0, 0), s)
        g.DrawImage(bmp, 0, 0, Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height)
        g.Dispose()
        'Dim small As Bitmap = New Bitmap(bmp, CInt(s.Width / 4), CInt(s.Height / 4))
        bmp.Save(Path.GetDirectoryName(appPath) & "\error.jpg", Imaging.ImageFormat.Jpeg)
    End Sub
    '##################################################################

    Private Sub アップデート前処理ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles アップデート前処理ToolStripMenuItem.Click
        'exeのバックアップがあったら先に消す（180530ver）
        'sakiPath = Path.GetDirectoryName(Form1.appPath) & "\AutoCheck1_bak.exe"
        Dim bFiles As String() = Directory.GetFiles(Path.GetDirectoryName(appPath))
        Dim dCount As Integer = 0
        For Each bF As String In bFiles
            If File.Exists(bF) And Regex.IsMatch(bF, "AutoCheck1_bak") Then
                Try
                    My.Computer.FileSystem.DeleteFile(bF)
                    dCount += 1
                Catch ex As Exception

                End Try
            End If
        Next
        If dCount > 0 Then
            NotifyIcon1.BalloonTipText = "アップデート前処理完了しました。"
        Else
            NotifyIcon1.BalloonTipText = "アップデート可能です。"
        End If

        'If File.Exists(sakiPath) Then   'exeのバックアップがあったら先に消す
        '    File.Delete(sakiPath)
        '    MsgBox("アップデート前処理完了しました。", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
        'Else
        '    MsgBox("アップデート可能です。", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
        'End If
    End Sub

    '==================================================================
    'フォームがアクティブでない時にキーを取得
    ' Windows API functions and constants
    Private Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As IntPtr, ByVal id As Integer, ByVal fsModifiers As Integer, ByVal vk As Keys) As Integer
    Private Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As IntPtr, ByVal id As Integer) As Integer
    Private Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Short
    Private Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Short) As Short
    Private Const MOD_ALT As Integer = &H1
    Private Const MOD_CONTROL As Integer = &H2
    Private Const MOD_SHIFT As Integer = &H4
    Private Const MOD_WIN As Integer = &H8
    Private Const WM_HOTKEY As Integer = &H312
    '==================================================================

    '==================================================================
    Public Class DoubleBufferingDataGridView
        Inherits System.Windows.Forms.WebBrowser
        Public Sub New()
            Me.DoubleBuffered = True
        End Sub
    End Class
    '==================================================================


    ''' <summary>
    ''' Form1.Load
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'exeのバックアップがあったら先に消す（180530ver）
        'sakiPath = Path.GetDirectoryName(Form1.appPath) & "\AutoCheck1_bak.exe"
        Dim bFiles As String() = Directory.GetFiles(Path.GetDirectoryName(appPath))
        Dim dCount As Integer = 0
        For Each bF As String In bFiles
            If File.Exists(bF) And Regex.IsMatch(bF, "AutoCheck1_bak") Then
                Try
                    My.Computer.FileSystem.DeleteFile(bF)
                    dCount += 1
                Catch ex As Exception

                End Try
            End If
        Next
        If dCount > 0 Then
            NotifyIcon1.BalloonTipText = "アップデート前処理完了しました。"
        End If
        'If File.Exists(sakiPath) Then   'exeのバックアップがあったら先に消す
        '    File.Delete(sakiPath)
        'End If

        'アプリケーション終了イベント
        AddHandler Application.ApplicationExit, AddressOf Application_ApplicationExit

        '開発環境用
        If Regex.IsMatch(My.Computer.Name, "ABCD|tak|NAKA|PING|MAO|PING2|LIU", RegexOptions.IgnoreCase) And Regex.IsMatch(appPath, "debug", RegexOptions.IgnoreCase) Then
#If DEBUG Then
            管理ToolStripMenuItem.Enabled = True
            SplitContainer3.Panel1.BackColor = Color.Yellow
            TextBox6.BackColor = Color.MediumAquamarine
            TextBox6.ForeColor = Color.LightYellow
        Else
            管理ToolStripMenuItem.Enabled = False
            End If
#End If

        '自宅環境ではアップデートしない
        If InStr(Environment.MachineName, "TAKASHI") > 0 And InStr(Path.GetDirectoryName(appPath), "Debug") > 0 Then
            updateFlag = False
            AdminFlag = True
        ElseIf InStr(Environment.MachineName, "NAKA") > 0 And InStr(Path.GetDirectoryName(appPath), "Debug") > 0 Then
            updateFlag = False
            AdminFlag = True
        ElseIf InStr(Environment.MachineName, "MAO") > 0 And InStr(Path.GetDirectoryName(appPath), "Debug") > 0 Then
            updateFlag = False
            AdminFlag = True
        ElseIf InStr(Environment.MachineName, "PING") > 0 And InStr(Path.GetDirectoryName(appPath), "Debug") > 0 Then
            updateFlag = False
            AdminFlag = True
        ElseIf InStr(Environment.MachineName, "PING2") > 0 And InStr(Path.GetDirectoryName(appPath), "Debug") > 0 Then
            updateFlag = False
            AdminFlag = True
        ElseIf InStr(Environment.MachineName, "LIU2") > 0 And InStr(Path.GetDirectoryName(appPath), "Debug") > 0 Then
            updateFlag = False
            AdminFlag = True
        ElseIf InStr(Environment.MachineName, "NAKA") > 0 Or InStr(Environment.MachineName, "PING") > 0 Or InStr(Environment.MachineName, "LIU2") > 0 Or InStr(Environment.MachineName, "PING2") > 0 Then
            AdminFlag = True
        End If

        Me.Text = "くま Ver." & Application.ProductVersion
        くまVer2AutoCheckToolStripMenuItem.Text = "くま Ver." & Application.ProductVersion
        If AdminFlag Then
            くまVer2AutoCheckToolStripMenuItem.Text &= "（管理モード）"
            Me.Text &= "（管理モード）"
        End If
        secValue = "F1"   'iniファイルセクション

        '*********************************************
        inif = New inifile.inifileAccess(Path.GetDirectoryName(appPath) + "\config\config.ini")
        '*********************************************
        テンプレート自動更新ToolStripMenuItem.Checked = inif.ReadINI(secValue, "AutoUpdate", True)

        AddHandler My.Settings.SettingChanging, AddressOf Settings_SettingChanging

        '---------------------------------------------------------------
        'Form1.inif.WriteINI(secValue, "StartForm", "mini")
        'SplitContainer2.SplitterDistance = 168
        'Me.Size = New Size(183, 520)
        'Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedDialog
        'StatusStrip1.Visible = False
        '---------------------------------------------------------------

        'コントロール配列
        'checkItem = New CheckBox() {CheckBox1, CheckBox2, CheckBox3, CheckBox4, CheckBox5, CheckBox6, CheckBox7, CheckBox8}
        radioItemCheck = New RadioButton() {RadioButton1, RadioButton2, RadioButton3}

        SplitContainer4.SplitterDistance = 0

        ComboBox1.SelectedIndex = 0

        'ToolStripStatusLabel1
        ToolStripStatusLabel1.Spring = True
        ToolStripStatusLabel1.TextAlign = ContentAlignment.MiddleLeft
        'StatusStrip1
        StatusStrip1.Items.Add(ToolStripProgressBar1)
        StatusStrip1.Items.Add(ToolStripStatusLabel1)
        'Form1
        'Me.WindowState = FormWindowState.Maximized
        Controls.Add(StatusStrip1)
        'Me.Controls.Add(TabBrowser1)
        SplitContainer2.Panel2.Controls.Add(TabBrowser1)

        'SplitContainer2.Panel2.Hide()
        'SplitContainer2.SplitterDistance = ToolStrip1.Width + 100
        'SplitContainer2.Panel1.Controls.Add(TabBrowser1)
        TabBrowser1.BringToFront()
        TabBrowser1.Dock = DockStyle.Fill
        'http://www.freethy.cn/post-145.html
        TabBrowser1.SelectedTab.WebBrowser.GoHome()


        PC名ToolStripMenuItem.Text = Environment.MachineName
        Dim pcData As String() = File.ReadAllLines(Path.GetDirectoryName(appPath) & "\config\PC番号.dat")
        For Each pcLine As String In pcData
            If pcLine <> "" Then
                Dim lines As String() = Split(pcLine, ",")
                If Regex.IsMatch(lines(0), PC名ToolStripMenuItem.Text, RegexOptions.IgnoreCase) Then
                    ログイン名ToolStripMenuItem.Text = lines(1)
                    Exit For
                End If
            End If
        Next

        'ファイル読み取り
        Dim fName As String = Path.GetDirectoryName(appPath) & "\bookmark.txt"
        Using sr As New StreamReader(fName, System.Text.Encoding.Default)
            Do While Not sr.EndOfStream
                Dim s As String = sr.ReadLine
                If s = "---" Then
                    ToolStripComboBox1.Items.Add("------------")
                    bookmarkURL.Add("------------")
                Else
                    Dim sArray As String() = Split(s, ",")
                    ToolStripComboBox1.Items.Add(sArray(0))
                    bookmarkURL.Add(sArray(1))
                End If
            Loop
        End Using
        ToolStripComboBox1.SelectedIndex = 0

        'チェックリスト読込
        Dim CheckNum As String = ""
        For Each rb As RadioButton In radioItemCheck
            If rb.Checked Then
                CheckNum = rb.Tag
                Exit For
            End If
        Next
        CheckListRead(CheckNum)

        '強調リスト読込
        fName = Path.GetDirectoryName(appPath) & "\強調リスト.txt"
        Dim r As Integer = 0
        Using sr As New StreamReader(fName, Encoding.Default)
            Do While Not sr.EndOfStream
                Dim s As String = sr.ReadLine
                Dim sArray As String() = Split(s, ",")
                DataGridView3.Rows.Add(s)
                r += 1
            Loop
        End Using

        'チェックする時間をずらすためにランダムで生成する
        Dim cRandom As New System.Random()
        Dim randomBase As Integer = 20
        For i As Integer = 0 To 2
            Dim iResult As Integer = cRandom.Next(randomBase * i + 10, randomBase * i + 20)
            checkTime &= iResult & "|"
        Next
        checkTime = checkTime.TrimEnd("|")


        SplitContainer2.SplitterDistance = 168
        If Form1.inif.ReadINI(secValue, "StartForm", "") = "mini" And miniForm = False Then
            Me.Size = New Size(183, 500)
            StatusStrip1.Visible = False
        Else
            'Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedDialog
            StatusStrip1.Visible = True
            'My.Application.ApplicationContext.MainForm = frm
        End If

        'MsgBox(RegistryRead("Software\Microsoft\Windows\CurrentVersion\Run"))
        If RegistryRead("Software\Microsoft\Windows\CurrentVersion\Run") = "" Then
            OS起動時スタートToolStripMenuItem.Checked = False
        Else
            OS起動時スタートToolStripMenuItem.Checked = True
        End If

        TM_EnableDoubleBuffering(TabBrowser1)        'ダブルバッファをTrueにする（ちらつき防止）

        ' ホットキーのために唯一無二のIDを生成する
        hotkeyID = GlobalAddAtom("GlobalHotKey " & Me.GetHashCode().ToString())
        ' Ctrl+Aキーを登録する
        RegisterHotKey(Me.Handle, hotkeyID, MOD_SHIFT, Keys.F2)
    End Sub

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown

    End Sub

    Protected Overrides Sub WndProc(ByRef m As Message)
        MyBase.WndProc(m)
        If m.Msg = WM_HOTKEY Then
            ' ホットキーが押されたときの処理
            Form1_F_loginlist.Show()
        End If
    End Sub

    'Dim timer2Count As Integer = 0
    'Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
    '    If timer2Count < 10 Then
    '        timer2Count += 1
    '    Else
    '        If shiftOn = True Then
    '            shiftOn = False
    '        End If
    '    End If
    'End Sub


    Private Sub HTMLsetTxt(ByVal elName As String, ByVal elNum As Integer, ByVal targetStr As String)
        Try
            TabBrowser1.SelectedTab.WebBrowser.Document.All.GetElementsByName(elName)(elNum).InnerText = targetStr
        Catch ex As Exception
            MsgBox("Error:HTMLsetTxt", MsgBoxStyle.SystemModal)
        End Try
    End Sub

    'チェックリスト読込
    Private Sub CheckListRead(ByRef rm As String)
        Dim fName As String = Path.GetDirectoryName(appPath) & "\CheckList.txt"
        Dim tFlag As Boolean = False
        Using sr As New System.IO.StreamReader(fName, System.Text.Encoding.Default)
            Do While Not sr.EndOfStream
                Dim s As String = sr.ReadLine
                If InStr(s, "[") And InStr(s, "]") Then
                    If s = rm Then
                        tFlag = True
                    Else
                        tFlag = False
                    End If
                ElseIf s = "" Then

                Else
                    If tFlag = True Then
                        Dim sArray As String() = Split(s, ",")
                        DataGridView1.Rows.Add(False, 0, sArray(0), sArray(1), sArray(2))
                    End If
                End If
            Loop
        End Using
    End Sub

    'チェックリスト読込ボタン
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        DataGridView1.Rows.Clear()

        Dim CheckNum As String = ""
        For Each rb As Windows.Forms.RadioButton In radioItemCheck
            If rb.Checked = True Then
                CheckNum = rb.Tag
                Exit For
            End If
        Next
        CheckListRead(CheckNum)
    End Sub

    'チェックリスト保存ボタン
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim tenpo As String = ""
        For Each rb As Windows.Forms.RadioButton In radioItemCheck
            If rb.Checked Then
                tenpo = rb.Tag
                Exit For
            End If
        Next

        Dim newLine As String = ""
        Dim fName As String = Path.GetDirectoryName(appPath) & "\CheckList.txt"
        Dim tFlag As Integer = 0
        Using sr As New System.IO.StreamReader(fName, System.Text.Encoding.Default)
            Dim sFlag As Integer = 0
            Do While Not sr.EndOfStream
                Dim s As String = sr.ReadLine
                If InStr(s, "[") And InStr(s, "]") Then
                    newLine &= s & vbCrLf
                    If s = tenpo Then
                        tFlag = 1
                        Dim lines As String = ""
                        For r As Integer = 0 To DataGridView1.RowCount - 2
                            lines &= DataGridView1.Item(2, r).Value
                            lines &= "," & DataGridView1.Item(3, r).Value
                            lines &= "," & DataGridView1.Item(4, r).Value
                            lines &= vbCrLf
                        Next
                        newLine &= lines
                        newLine &= vbCrLf
                    Else
                        tFlag = 0
                    End If
                ElseIf s = "" Then
                    If tFlag = 0 Then
                        newLine &= vbCrLf
                    End If
                Else
                    If tFlag = 0 Then
                        newLine &= s & vbCrLf
                    Else

                    End If
                End If
            Loop
        End Using

        'newLine &= vbCrLf
        'newLine &= tenpo & vbCrLf

        'Dim line As String = ""
        'For r As Integer = 0 To DataGridView1.RowCount - 2
        '    line &= DataGridView1.Item(2, r).Value
        '    line &= "," & DataGridView1.Item(3, r).Value
        '    line &= "," & DataGridView1.Item(4, r).Value
        '    line &= vbCrLf
        'Next
        'newLine &= line

        MsgBox("保存しました" & vbCrLf & newLine, MsgBoxStyle.SystemModal)
    End Sub

    Private Sub TabBrowser_StatusTextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles TabBrowser1.StatusTextChanged
        Try
            Me.ToolStripStatusLabel1.Text = CType(sender, WebBrowser).StatusText
        Catch ex As Exception

        End Try
        Dim w As WebBrowser = CType(sender, WebBrowser)
        'ListBox3.Items.Add(CType(sender, WebBrowser).StatusText)
        'ListBox3.SelectedIndex = ListBox3.Items.Count - 1
    End Sub

    Private Sub TabBrowser_ProgressChanged(ByVal sender As Object, ByVal e As WebBrowserProgressChangedEventArgs) Handles TabBrowser1.ProgressChanged
        Me.ToolStripProgressBar1.Maximum = CInt(e.MaximumProgress)
        If e.CurrentProgress > 0 And e.CurrentProgress < Me.ToolStripProgressBar1.Maximum Then
            Me.ToolStripProgressBar1.Value = CInt(e.CurrentProgress)
            Application.DoEvents()
        End If
    End Sub
    '-------------------------------------------------------------------



    Public Shared DocumentText As String
    Public Shared home As String = "http://www.yahoo.co.jp"
    'Public Shared home As String = "http://www.google.co.jp"

    'home
    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        ToolStripTextBox1.Text = home
        WebTabPage.WebBrowser_Navigating(home)
        TabBrowser1.SelectedTab.WebBrowser.Navigate(New Uri(ToolStripTextBox1.Text))
    End Sub

    '戻る
    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        If TabBrowser1.SelectedTab.WebBrowser.CanGoBack = True Then
            TabBrowser1.SelectedTab.WebBrowser.GoBack()
        End If
    End Sub

    '進む
    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        If TabBrowser1.SelectedTab.WebBrowser.CanGoForward = True Then
            TabBrowser1.SelectedTab.WebBrowser.GoForward()
        End If
    End Sub

    'URL入力欄クリック
    Private Sub ToolStripTextBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.Click

    End Sub

    Private Sub ToolStripTextBox1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.DoubleClick
        ToolStripTextBox1.SelectAll()
    End Sub

    'Go
    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Try
            TabBrowser1.SelectedTab.WebBrowser.Focus()
            WebTabPage.WebBrowser_Navigating(ToolStripTextBox1.Text)
            TabBrowser1.SelectedTab.WebBrowser.Navigate(New Uri(ToolStripTextBox1.Text))
        Catch ex As System.UriFormatException
            Return
        End Try
    End Sub

    'ToolStripTextBox1のEnter入力
    Private Sub ToolStripTextBox1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ToolStripTextBox1.KeyUp
        If e.KeyCode = Keys.Enter Then
            Try
                WebTabPage.WebBrowser_Navigating(ToolStripTextBox1.Text)
                TabBrowser1.SelectedTab.WebBrowser.Navigate(New Uri(ToolStripTextBox1.Text))
            Catch ex As System.UriFormatException
                Return
            End Try
        End If
    End Sub

    'ストップ
    Private Sub ToolStripButton8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton8.Click
        TabBrowser1.SelectedTab.WebBrowser.Stop()
    End Sub

    'リフレッシュ
    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        navCount = 0
        TabBrowser1.SelectedTab.WebBrowser.Refresh()
    End Sub

    'ブックマーク移動
    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Dim path As String = bookmarkURL(ToolStripComboBox1.SelectedIndex)
        If path = "------------" Then
            Exit Sub
        End If

        TabBrowser1.SelectedTab.WebBrowser.Focus()
        WebTabPage.WebBrowser_Navigating(path)
        TabBrowser1.SelectedTab.WebBrowser.Navigate(New Uri(path))
    End Sub

    'お気に入り追加
    Private Sub ToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton7.Click

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        For Each Element As HtmlElement In TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")
            Select Case Element.GetAttribute("type")
                Case "text"
                    MsgBox("input:" & Element.GetAttribute("value"))
                Case "radio"
                    MsgBox("radio:" & Element.GetAttribute("value"))
            End Select
        Next
    End Sub

    'チェックボタン
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim str As String() = Nothing
        If InStr(TextBox1.Text, " ") Then
            str = Split(TextBox1.Text, " ")
        Else
            str = Split(TextBox1.Text, "")
        End If

        'For Each s As String In str
        'Find(TabBrowser1.SelectedTab.WebBrowser, s)
        'Next
        Find(TabBrowser1.SelectedTab.WebBrowser, str, 0)

        'データグリッドで調べる
        For r As Integer = 0 To DataGridView1.RowCount - 1
            If DataGridView1.Item(2, r).Value = "" Then

            Else
                Dim cs1 As String = DataGridView1.Item(4, r).Value
                Dim cs2 As String() = Split(cs1, "|")
                Dim cNum As Integer = 0
                Dim cFlag As Integer = 0
                'For Each s As String In cs2
                '    Dim RF As String = Find(TabBrowser1.SelectedTab.WebBrowser, s)
                '    cFlag += RF
                '    cNum += 1
                'Next
                cFlag = Find(TabBrowser1.SelectedTab.WebBrowser, cs2, 0)
                If cFlag > DataGridView1.Item(3, r).Value Then
                    DataGridView1.Item(0, r).Value = True
                Else
                    DataGridView1.Item(0, r).Value = False
                End If
                DataGridView1.Item(1, r).Value = cFlag
            End If
        Next r


        'Dim chNum As Integer = 0
        'str = New String() {"入荷予定", "入金待ち", "ゆうメール", "佐川急便", "ヤマト", "ゆうメール不可", "メール便不可", "宅配便に変更"}
        'For Each s As String In str
        '    Dim RF As String = Find(TabBrowser1.SelectedTab.WebBrowser, s)
        '    If RF > 0 Then
        '        checkItem(chNum).Checked = True
        '    Else
        '        checkItem(chNum).Checked = False
        '    End If
        '    chNum += 1
        'Next
    End Sub

    Private Sub CheckBox10_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox10.CheckedChanged
        Dim s(DataGridView3.RowCount - 2) As String
        For r As Integer = 0 To DataGridView3.RowCount - 2
            s(r) = DataGridView3.Item(0, r).Value
        Next
        Find(TabBrowser1.SelectedTab.WebBrowser, s, 0)
    End Sub

    Public Function Find(ByVal browser As WebBrowser, ByVal str As String(), ByRef num As Integer) As Integer
        Dim RF As Integer = 0
        Dim Doc As mshtml.HTMLDocument       'MSHTML.HTMLDocument 
        Dim Body As mshtml.HTMLBody      'MSHTML.HTMLBody 
        Dim Range As mshtml.IHTMLTxtRange  'MSHTML.IHTMLTxtRange 
        Dim BMK As String = ""

        '検索文字列を入れておいてください。 
        'If String.IsNullOrEmpty(str) Then
        If str.Length = 0 Then
            Return 0
        End If

        Doc = DirectCast(browser.Document.DomDocument, mshtml.IHTMLDocument2)
        Body = Doc.body
        Range = Body.createTextRange

        For r As Integer = 0 To DataGridView3.RowCount - 1
            DataGridView3.Item(0, r).Style.BackColor = Color.White
        Next

        '≫≫≫≫≫ 検索開始
        Dim Rnum As Integer = 0
        For Each f As String In str
            If f <> "" Then
                Do While Range.findText(f)
                    '最初に見つかった位置を保存しておきます。 
                    If String.IsNullOrEmpty(BMK) = 0 Then

                    End If
                    BMK = Range.getBookmark

                    '検索した語句を黄色く反転させる。 
                    Range.execCommand("BackColor", False, "Yellow")
                    RF += 1

                    DataGridView3.Item(0, Rnum).Style.BackColor = Color.Yellow
                    DataGridView3.CurrentCell = DataGridView3(0, Rnum)
                    DataGridView3.Item(0, Rnum).Selected = False

                    '論理カーソル位置を、検索した語句の末尾に移動させる。 
                    Range.collapse(False)

                    If num = 1 Then
                        Exit Do
                    End If
                Loop
            End If
            Rnum += 1
        Next
        '≪≪≪≪≪ 検索終了 

        'ついでに、最初に見つけた語句の位置までスクロールさせています。 
        If Not String.IsNullOrEmpty(BMK) Then
            Range.moveToBookmark(BMK)
            Range.scrollIntoView()
        End If

        '最後は一応、後始末を。 
        Range = Nothing
        Body = Nothing
        Doc = Nothing

        Return RF
    End Function

    '重複チェックボタン
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        DataGridView2.ClearSelection()

        Dim num As Integer = 0
        Dim mTxt As String = ""
        For r1 As Integer = 0 To DataGridView2.RowCount - 1
            If mTxt = "" Then
                mTxt = DataGridView2.Item(0, r1).Value
            Else
                If DataGridView2.Item(0, r1 - 1).Selected = False Then
                    For r2 As Integer = r1 To DataGridView2.RowCount - 1
                        If mTxt = DataGridView2.Item(0, r2).Value Then
                            DataGridView2.Item(0, r2).Selected = True
                            num += 1
                            If DataGridView2.Item(0, r1 - 1).Selected = False Then
                                DataGridView2.Item(0, r1 - 1).Selected = True
                                num += 1
                            End If
                        End If
                    Next
                Else

                End If
                mTxt = DataGridView2.Item(0, r1).Value
            End If
        Next

        TextBox2.Text = num
    End Sub

    '価格取得
    Private Sub TextBox4_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox4.KeyUp
        If e.KeyCode = Keys.Enter Then
            Button13_Click(sender, e)
        End If
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Dim code As String = TextBox4.Text

        Dim url As String = ""
        Dim pos As Integer = 0
        If ComboBox1.SelectedIndex = 0 Then
            url = "http://store.shopping.yahoo.co.jp/fkstyle/" & code & ".html"
        ElseIf ComboBox1.SelectedIndex = 1 Then
            url = "http://item.rakuten.co.jp/patri/" & code & "/"
        ElseIf ComboBox1.SelectedIndex = 2 Then
            url = "http://www.qoo10.jp/shop/yasuyasu?keyword=" & code
        ElseIf ComboBox1.SelectedIndex = 3 Then
            url = "http://www.qoo10.jp/shop/fukuoka?keyword=" & code
        End If

        TabBrowser1.SelectedTab.WebBrowser.Visible = False
        ToolStripTextBox1.Text = url
        WebTabPage.WebBrowser_Navigating(ToolStripTextBox1.Text)
        TabBrowser1.SelectedTab.WebBrowser.Navigate(New Uri(ToolStripTextBox1.Text))

        Try
            Dim res = KakakuGet(TabBrowser1.SelectedTab.WebBrowser, ComboBox1)
            'TabBrowser1.SelectedTab.WebBrowser.Document.Window.ScrollTo(0, pos)
            Dim str As String() = New String() {res}
            Find(TabBrowser1.SelectedTab.WebBrowser, str, 1)
            TextBox3.Text = res
        Catch ex As Exception
            MsgBox("見つかりません", MsgBoxStyle.SystemModal)
        End Try

        TabBrowser1.SelectedTab.WebBrowser.Visible = True
    End Sub

    Public Function KakakuGet(ByVal wb As WebBrowser, ByVal cb As ComboBox)
        Dim res As String = ""
        Try
            Do
                Application.DoEvents()
                'Loop Until Form1.TabBrowser1.SelectedTab.WebBrowser.ReadyState = WebBrowserReadyState.Complete And Form1.TabBrowser1.SelectedTab.WebBrowser.IsBusy = False
            Loop Until wb.ReadyState = WebBrowserReadyState.Complete And wb.IsBusy = False

            Dim result As New ArrayList
            Dim allHtml As HtmlDocument = wb.Document
            'Dim allHtml As HtmlDocument = WebBrowser1.Document

            Dim tag As String = ""
            Dim target As String = ""
            Dim targetStr As String = ""
            Dim num As Integer = 0
            If ComboBox1.SelectedIndex = 0 Then
                tag = "span"
                target = "className"
                targetStr = "elNum"
                num = 0
            ElseIf ComboBox1.SelectedIndex = 1 Then
                tag = "span"
                target = "className"
                targetStr = "price2"
                num = 0
            ElseIf ComboBox1.SelectedIndex = 2 Then
                tag = "strong"
                target = "title"
                targetStr = "クーポン利用価格"
                num = 0
            ElseIf ComboBox1.SelectedIndex = 3 Then
                tag = "strong"
                target = "title"
                targetStr = "クーポン利用価格"
                num = 0
            End If

            result = HTMLgetTxt(allHtml, tag, target, targetStr)

            res = result(num)
            'res = Replace(res, "円", "")
            'res = Replace(res, ",", "")
        Catch ex As Exception
            res = 0
        End Try

        Return res
    End Function

    Public Function HTMLgetTxt(ByVal allHtml As HtmlDocument, ByVal tag As String, ByVal target As String, ByVal targetStr As String) As ArrayList
        Dim result As New ArrayList

        For Each el As HtmlElement In allHtml.GetElementsByTagName(tag)
            If el.GetAttribute(target) = targetStr Then
                result.Add(el.InnerText)
            End If
        Next

        Return result
    End Function

    'DataGridView2への貼り付け
    Private _editingColumn As Integer
    Public ColumnChars() As String
    Private Sub DataGridView2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DataGridView2.KeyDown
        Dim dgv As DataGridView = CType(sender, DataGridView)
        Dim x As Integer = dgv.CurrentCellAddress.X
        Dim y As Integer = dgv.CurrentCellAddress.Y

        If e.KeyCode = Keys.Delete Or e.KeyCode = Keys.Back Then
            ' セルの内容を消去
            Dim delR As New ArrayList
            For r As Integer = 0 To dgv.SelectedCells.Count - 1
                delR.Add(dgv.SelectedCells(r).RowIndex)
            Next
            delR.Sort()
            delR.Reverse()
            For r As Integer = 0 To dgv.SelectedCells.Count - 1
                If delR(r) <> dgv.RowCount - 1 Then
                    dgv.Rows.RemoveAt(delR(r))
                End If
            Next
        ElseIf (e.Modifiers And Keys.Control) = Keys.Control And e.KeyCode = Keys.V Then
            ' クリップボードの内容を取得
            Dim clipText As String = Clipboard.GetText()
            ' 改行を変換
            clipText = clipText.Replace(vbCrLf, vbLf)
            clipText = clipText.Replace(vbCr, vbLf)
            ' 改行で分割
            Dim lines() As String = clipText.Split(vbLf)

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
                        DataGridView2_CellKeyPress(sender, tmpe)
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
                    ' ------------------------------------------------------
                    ' 次のセルへ
                    c2 = c2 + 1
                Next
            Next
        End If
    End Sub

    Private Sub DataGridView2_CellKeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs)
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



    '=================================================================================
    'リストボックス1
    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        Dim di As New System.IO.DirectoryInfo(Path.GetDirectoryName(appPath) & "\template")
        Dim files As System.IO.FileInfo() = di.GetFiles("*.txt", System.IO.SearchOption.AllDirectories)

        ListBox2.Items.Clear()

        'ListBox1に結果を表示する
        For Each f As System.IO.FileInfo In files
            Dim fname As String() = Split(System.IO.Path.GetFileNameWithoutExtension(f.Name), "、")
            If ListBox1.SelectedItem = fname(0) Then
                ListBox2.Items.Add(fname(1))
            End If
        Next
    End Sub

    'リストボックス2
    Private Sub ListBox2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox2.SelectedIndexChanged
        Dim dr As String = Path.GetDirectoryName(appPath) & "\template"
        Dim copyFile As String = dr & "\" & ListBox1.SelectedItem & "、" & ListBox2.SelectedItem & ".txt"
        Dim sr As New System.IO.StreamReader(copyFile, System.Text.Encoding.Default)
        Dim copyStr As String = sr.ReadToEnd
        sr.Close()

        ToolTip1.SetToolTip(ListBox2, copyStr)
        Clipboard.SetText(copyStr)
    End Sub







    'フォームのサイズ最小化バグ修正
    Private Sub Settings_SettingChanging(ByVal sender As Object, ByVal e As System.Configuration.SettingChangingEventArgs)
        If Not Me.WindowState = FormWindowState.Normal Then
            If e.SettingName = "MyClientSize" Or e.SettingName = "MyLocation" Then
                e.Cancel = True
            End If
        End If
    End Sub

    'DOM解析
    Public Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        FormFront(Form4)
        Form4.TreeView1.Nodes.Clear()
        Dim DomDocument As Object = TabBrowser1.SelectedTab.WebBrowser.Document.DomDocument
        Dim TreeNode As New TreeNode("<html>")
        Form4.TreeView1.BeginUpdate()
        Form4.TreeView1.Nodes.Add(TreeNode)
        Dim check As String() = New String() {"tr", "td", "th", "div", "a", "span", "font", "form", "input", "select", "option", "table", "p", "img", "ul", "li", "dd", "small"}
        For i As Integer = 0 To check.Length - 1
            HTMLDOM(TreeNode, check(i))
        Next
        'Form4.TreeView1.ExpandAll()
        Form4.TreeView1.EndUpdate()
        Marshal.ReleaseComObject(DomDocument)
    End Sub

    Private Sub HTMLDOM(ByRef TreeNode As TreeNode, ByRef name As String)
        Dim childNode As New TreeNode(name)
        TreeNode.Nodes.Add(childNode)

        Dim num As Integer = 0
        For Each Element As HtmlElement In TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName(name)
            Dim str As String = Replace(Element.OuterHtml, " ", "")
            str = Replace(Element.OuterHtml, "　", "")
            Dim childNodeLevel2 As New TreeNode(num & " : " & str)
            childNode.Nodes.Add(childNodeLevel2)
            num += 1
        Next
    End Sub


    'ブラウザ拡大縮小
    Private Sub ToolStripSplitButton1_DropDownItemClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ToolStripSplitButton1.DropDownItemClicked
        ToolStripSplitButton1.Text = e.ClickedItem.Text
        TabBrowser1.SelectedTab.WebBrowser.Document.Body.Style = ";Zoom:" & e.ClickedItem.Text
    End Sub

    Private Sub 終了ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 終了ToolStripMenuItem.Click
        CloseForm(0)
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If e.CloseReason = CloseReason.FormOwnerClosing Or e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
            Me.WindowState = FormWindowState.Minimized
            Try
                'If Me.ShowInTaskbar = True Then
                '    Me.ShowInTaskbar = False
                'End If
            Catch ex As Exception

            End Try
            'CloseForm()
        End If
    End Sub

    Private Sub CloseForm(mode As Integer)
        Dim ops As FormCollection = Application.OpenForms
        For i As Integer = ops.Count - 1 To 0 Step -1
            Try
                Select Case False
                    Case InStr(ops(i).Name, "Form1")
                        ops(i).Close()
                        ops(i).Dispose()
                End Select
            Catch ex As Exception

            End Try
        Next

        If NotifyIcon1.Visible And mode = 0 Then
            NotifyIcon1.Dispose()
        End If
    End Sub

    'アプリケーションの終了イベント
    Private Sub Application_ApplicationExit(ByVal sender As Object, ByVal e As EventArgs)
        ' ホットキーの登録を解除し、アトムを削除する
        Try
            UnregisterHotKey(Me.Handle, hotkeyID)
            GlobalDeleteAtom(hotkeyID)
        Catch ex As Exception

        End Try

        NotifyIcon1.Dispose()
        'ApplicationExitイベントハンドラを削除
        RemoveHandler Application.ApplicationExit, AddressOf Application_ApplicationExit
    End Sub

    Private Sub バージョン情報ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles バージョン情報ToolStripMenuItem.Click
        Dim str As String = ""
        ' バージョン名（AssemblyInformationalVersion属性）を取得
        str = "バージョン名：" & Application.ProductVersion & vbCrLf
        ' 製品名（AssemblyProduct属性）を取得
        str &= "製品名：" & Application.ProductName & vbCrLf
        ' 会社名（AssemblyCompany属性）を取得
        str &= "会社名：" & Application.CompanyName & vbCrLf
        MsgBox(str, MsgBoxStyle.OkOnly And MsgBoxStyle.SystemModal, "バージョン情報")
    End Sub

    Private Sub 設定ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 設定ToolStripMenuItem.Click
        FormFront(Form1_F_kanri)
    End Sub

    Private Sub 全てのタブを閉じるToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 全てのタブを閉じるToolStripMenuItem.Click
        MsgBox("未実装", MsgBoxStyle.SystemModal)
        Exit Sub

        If TabBrowser1.TabCount > 1 Then
            For i As Integer = 1 To TabBrowser1.TabPages.Count - 1
                TabBrowser1.TabPages(TabBrowser1.TabPages.Count - i).Dispose()
            Next
        End If
    End Sub

    Private Sub 荷物追跡ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 荷物追跡ToolStripMenuItem.Click
        FormFront(Form3)
    End Sub

    Private Sub 売行動向ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 売行動向ToolStripMenuItem.Click
        FormFront(Form7)
    End Sub

    Private Sub ランクチェックToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ランクチェックToolStripMenuItem.Click
        FormFront(SearchRank)
    End Sub


    Private Sub CSV変換ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 特殊処理ToolStripMenuItem.Click
        FormFront(Csv_change)
    End Sub

    Private Sub アップデートToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles アップデートToolStripMenuItem.Click
        'Dim Ret As Long = Shell(Path.GetDirectoryName(appPath) & "\ACupdate.exe", vbNormalFocus)
        'Dim ops As FormCollection = Application.OpenForms
        'For i As Integer = ops.Count - 1 To 0 Step -1
        '    Try
        '        Select Case False
        '            Case Regex.IsMatch(ops(i).Name, "UpdatePanel|Form1")
        '                ops(i).Close()
        '                ops(i).Dispose()
        '        End Select
        '    Catch ex As Exception

        '    End Try
        'Next

        UpdatePanel.ShowDialog()
        Application.DoEvents()
    End Sub

    Private Sub ログインToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ログインToolStripMenuItem.Click
        Dim browserUrl As String = TabBrowser1.SelectedTab.WebBrowser.Url.ToString()

        Dim fName As String = Path.GetDirectoryName(appPath) & "\config\ログイン.txt"
        Dim elName As String = ""
        Dim elNum As Integer = 0
        Dim targetStr As String = ""

        Dim Qoo1 As Integer = 0
        Dim Qoo2 As Integer = 0
        If InStr(browserUrl, "qsm.qoo10.jp") > 0 Then
            Dim DR As DialogResult = MsgBox("Qoo10 通販の雑貨倉庫なら「はい」、福岡通販堂なら「いいえ」", MsgBoxStyle.YesNo Or MsgBoxStyle.SystemModal)
            If DR = Windows.Forms.DialogResult.Yes Then
                Qoo1 = 0
                Qoo2 = 1
            Else
                Qoo1 = 2
                Qoo2 = 3
            End If
        End If

        Dim inPass As Integer = 0
        Using sr As New System.IO.StreamReader(fName, System.Text.Encoding.Default)
            Do While Not sr.EndOfStream
                Dim s As String = sr.ReadLine
                If s <> "" Then
                    Dim sArray As String() = Split(s, ",")
                    If InStr(browserUrl, sArray(0)) > 0 Then
                        If InStr(browserUrl, "qsm.qoo10.jp") > 0 Then
                            If Qoo1 <= inPass And Qoo2 >= inPass Then
                                elName = sArray(1)
                                elNum = sArray(2)
                                targetStr = sArray(3)

                                HTMLsetTxt(elName, elNum, targetStr)
                            End If
                            inPass += 1
                        Else
                            elName = sArray(1)
                            elNum = sArray(2)
                            targetStr = sArray(3)

                            HTMLsetTxt(elName, elNum, targetStr)
                        End If
                    End If
                End If
            Loop
        End Using
    End Sub

    Private Sub インストールフォルダを開くToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles インストールフォルダを開くToolStripMenuItem.Click
        System.Diagnostics.Process.Start(Path.GetDirectoryName(appPath))
    End Sub

    Private Sub 資格情報マネージャーToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 資格情報マネージャーToolStripMenuItem.Click
        System.Diagnostics.Process.Start("control.exe", "/name Microsoft.CredentialManager")
    End Sub

    Private Sub デバイスとプリンターの表示ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles デバイスとプリンターの表示ToolStripMenuItem.Click
        System.Diagnostics.Process.Start("control.exe", "/name Microsoft.DevicesAndPrinters")
    End Sub

    Private Sub フォルダーオプションToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles フォルダーオプションToolStripMenuItem.Click
        System.Diagnostics.Process.Start("control.exe", "/name Microsoft.FolderOptions")
    End Sub

    Private Sub ユーザーアカウントToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ユーザーアカウントToolStripMenuItem.Click
        System.Diagnostics.Process.Start("control.exe", "/name Microsoft.UserAccounts")
    End Sub

    Private Sub スプールフォルダToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles スプールフォルダToolStripMenuItem.Click
        Dim openPath As String = "C:\Windows\System32\spool\PRINTERS\"
        System.Diagnostics.Process.Start(Path.GetDirectoryName(openPath))
    End Sub

    Private Sub スプーラー再起動ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles スプーラー再起動ToolStripMenuItem.Click
        System.Diagnostics.Process.Start(Path.GetDirectoryName(appPath) & "\tools\clearprintspool.bat")
    End Sub

    Private Sub レジストリToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles レジストリToolStripMenuItem.Click
        FormFront(Registry)
    End Sub

    Private Sub Yahoo出荷日入力ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Yahoo出荷日入力ToolStripMenuItem.Click
        Dim inputText As String = InputBox("出荷日を入力してください" & vbCrLf & "例）2016/12/31", "出荷日入力", Format(Now, "yyyy/MM/dd"))
        If inputText <> "" Then
            Dim num As Integer = 0
            For i As Integer = 0 To 499
                Try
                    TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("shipDate" & i).SetAttribute("value", inputText)
                    num += 1
                Catch ex As Exception
                    Exit For
                End Try
            Next
            MsgBox(num & "件入力しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    Private Sub Yahoo追跡番号確認ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Yahoo追跡番号確認ToolStripMenuItem.Click
        Dialog.DataGridView1.Rows.Clear()

        Dim elements = TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")
        Dim r As Integer = 0
        For i As Integer = 0 To elements.Count - 1
            If InStr(elements(i).Name, "shipInvoiceNumber1_") > 0 Then
                Dialog.DataGridView1.Rows.Add(1)
                Dialog.DataGridView1.Item(1, r).Value = elements(i).GetAttribute("value")
                r += 1
            End If
        Next

        r = 0
        elements = TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")
        For i As Integer = 0 To elements.Count - 1
            If InStr(elements(i).Name, "shipUrl") > 0 Then
                Dialog.DataGridView1.Item(2, r).Value = elements(i).GetAttribute("value")
                r += 1
            End If
        Next

        r = 0
        elements = TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")
        For i As Integer = 0 To elements.Count - 1
            If InStr(elements(i).Name, "shipDate") > 0 Then
                Dialog.DataGridView1.Item(3, r).Value = elements(i).GetAttribute("value")
                r += 1
            End If
        Next

        r = 0
        elements = TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("a")
        For i As Integer = 0 To elements.Count - 1
            If InStr(elements(i).Name, "orderId") > 0 Then
                Dialog.DataGridView1.Item(0, r).Value = elements(i).InnerText
                r += 1
            End If
        Next

        FormFront(Dialog)
        Dialog.Size = New Size(700, 590)
    End Sub

    Private Sub 楽天追跡番号確認ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 楽天追跡番号確認ToolStripMenuItem.Click
        Dialog.DataGridView1.Rows.Clear()

        Dim elements = TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")
        Dim r As Integer = 0
        For i As Integer = 0 To elements.Count - 1
            If InStr(elements(i).Name, "shipping_number_") > 0 Then
                Dialog.DataGridView1.Rows.Add(1)
                Dialog.DataGridView1.Item(0, r).Value = elements(i).Name.Substring(16, 26)
                Dialog.DataGridView1.Item(1, r).Value = elements(i).GetAttribute("value")
                r += 1
            End If
        Next

        r = 0
        elements = TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("select")
        For i As Integer = 0 To elements.Count - 1
            If InStr(elements(i).Name, "deliver_company_") > 0 Then
                For k As Integer = 0 To elements(i).Children.Count - 1
                    If elements(i).Children.Item(k).GetAttribute("selected") = "True" Then
                        Dialog.DataGridView1.Item(2, r).Value = elements(i).Children.Item(k).InnerText
                    End If
                Next
                r += 1
            End If
        Next

        r = 0
        elements = TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")
        For i As Integer = 0 To elements.Count - 1
            If InStr(elements(i).Name, "shipping_date_") > 0 Then
                Dialog.DataGridView1.Item(3, r).Value = elements(i).GetAttribute("value")
                r += 1
            End If
        Next

        FormFront(Dialog)
        Dialog.Size = New Size(480, 590)
    End Sub

    Private Sub スマホ画面ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles スマホ画面ToolStripMenuItem.Click
        FormFront(iphone)
        iphone.Location = New Point(Me.Location.X + Me.Size.Width, Me.Location.Y + 100)
        iphone.Enabled = True
        'iphone.WebBrowser1.Navigate(TabBrowser1.SelectedTab.WebBrowser.Url)
    End Sub

    Private Sub Form1_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged
        If iphone.Enabled = True Then
            iphone.Location = New Point(Me.Location.X + Me.Size.Width, Me.Location.Y + 100)
        End If
        If Form1_F_Newitem.Enabled = True Then
            Form1_F_Newitem.Location = New Point(Me.Location.X + Me.Size.Width, Me.Location.Y)
        End If

        ToolStripTextBox1.Size = New Size(Me.Width - 536, 25)
    End Sub

    Private Sub Form1_Move(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Move
        If iphone.Enabled = True Then
            iphone.Location = New Point(Me.Location.X + Me.Size.Width, Me.Location.Y + 100)
        End If
        If Form1_F_Newitem.Enabled = True Then
            Form1_F_Newitem.Location = New Point(Me.Location.X + Me.Size.Width, Me.Location.Y)
        End If
    End Sub


#Region " 各種コンバート関数 "

    '############################################################
    '変換関数作成
    '数字全角→半角変換
    Public Function StrConvNumeric(ByVal s As String) As String
        Dim i As Integer
        For i = 0 To 9
            s = Replace(s, Chr(Asc("０") + i), Chr(Asc("0") + i))
        Next
        StrConvNumeric = s
    End Function

    '大文字→小文字
    Public Function StrConvToNarrow(ByVal s As String) As String
        s = s.ToLower()
        StrConvToNarrow = s
    End Function

    '英語全角→半角変換
    Public Function StrConvEnglish(ByVal s As String) As String
        Dim i As Integer
        For i = 0 To 25
            s = Replace(s, Chr(Asc("Ａ") + i), Chr(Asc("A") + i))
        Next
        For i = 0 To 25
            s = Replace(s, Chr(Asc("ａ") + i), Chr(Asc("a") + i))
        Next
        StrConvEnglish = s
    End Function

    '片仮名全角→半角変換
    Public Function StrConvKana(ByVal s As String) As String
        Dim i As Long
        Dim strTemp As String = Nothing
        Dim strKana As String = Nothing
        Dim chrKana As String = Nothing

        For i = 1& To Len(s)
            chrKana = Mid$(s, i, 1&)
            Select Case Asc(chrKana)
                Case 166 To 223
                    '半角が続いたら文字をつなぐ
                    strKana = strKana & chrKana
                Case Else
                    '全角文字になったら半角の未処理文字を全部全角
                    'に変換これにより濁点処理等が不要
                    If Len(strKana) > 0& Then
                        strTemp = strTemp & StrConv(strKana, vbWide)
                        strKana = vbNullString
                    End If
                    strTemp = strTemp & chrKana
            End Select
        Next i
        '最後の文字が半角の場合の処理
        If Len(strKana) > 0& Then
            strTemp = strTemp & StrConv(strKana, vbWide)
        End If
        StrConvKana = strTemp
    End Function

    '「"」を削除
    Public Function StrConvDoubleC(ByVal s As String) As String
        s = Replace(s, Chr(&H22), "")
        StrConvDoubleC = s
    End Function

    '全角「．」を半角
    Public Function StrConvComma(ByVal s As String) As String
        s = Replace(s, Chr(Asc("．")), Chr(Asc(".")))
        StrConvComma = s
    End Function

    '全角「＃」を半角
    Public Function StrConvSharp(ByVal s As String) As String
        s = Replace(s, Chr(Asc("＃")), Chr(Asc("#")))
        StrConvSharp = s
    End Function

    '全角「＆」を半角
    Public Function StrConvAnd(ByVal s As String) As String
        s = Replace(s, Chr(Asc("＆")), Chr(Asc("&")))
        StrConvAnd = s
    End Function

    '全角「－」を変換（平仮名・片仮名の後にある時）
    Public Function StrConvHyphen2(ByVal s As String) As String
        Dim i As Long
        Dim strTemp As String = ""
        Dim chrKana As String = ""
        Dim flag As Boolean = False

        For i = 1 To Len(s)
            chrKana = Mid$(s, i, 1&)
            Select Case Asc(chrKana)
                Case Asc("ぁ") To Asc("ん"), Asc("ァ") To Asc("ヶ")
                    strTemp &= chrKana
                    flag = True
                Case Asc("－")
                    If flag = True Then
                        strTemp &= "ー"
                    Else
                        strTemp &= chrKana
                    End If
                    flag = False
                Case Else
                    strTemp &= chrKana
                    flag = False
            End Select
        Next i

        StrConvHyphen2 = strTemp
    End Function

    '全角「－」を半角
    Public Function StrConvHyphen(ByVal s As String) As String
        s = Replace(s, Chr(Asc("－")), Chr(Asc("-")))
        StrConvHyphen = s
    End Function

    '半角「~」を全角
    Public Function StrConvNamisen(ByVal s As String) As String
        s = Replace(s, Chr(Asc("~")), Chr(Asc("～")))
        StrConvNamisen = s
    End Function

    '全角スペースを削除
    Public Function StrConvZSpace(ByVal s As String) As String
        s = Replace(s, "　", "")
        StrConvZSpace = s
    End Function

    '半角スペースを削除
    Public Function StrConvHSpace(ByVal s As String) As String
        s = Replace(s, " ", "")
        StrConvHSpace = s
    End Function

    '全角スペースを半角に
    Public Function StrConvZSpaceToH(ByVal s As String) As String
        s = Replace(s, "　", " ")
        StrConvZSpaceToH = s
    End Function

    '最初の00を消す
    Public Function StrConvDel00(ByVal s As String) As String
        Dim x As Integer = 0
        s = StrConvDoubleC(s)
        If IsNumeric(s) = True And s Mod 1 = 0 Then '文字列が数値形式か調べる
            If s < 9999999 Then     'JIS番号でないことを調べる
                x = Integer.Parse(s)
                s = x.ToString
            End If
        End If
        StrConvDel00 = s
    End Function

    '数値区切を追加
    Public Function StrConv3Keta(ByVal s As String) As String
        s = StrConvDoubleC(s)
        If IsNumeric(s) = True Then '文字列が数値形式か調べる
            If s < 9999999 Then     'JIS番号でないことを調べる
                s = String.Format("{0:#,0}", Integer.Parse(s))
                s = Replace(s, ",", "，") 'コンマはデータ上は全角にする
            End If
        End If
        s = Chr(&H22) & s & Chr(&H22)
        StrConv3Keta = s
    End Function
    '############################################################

#End Region

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox2.Checked = True Then
            CheckBox2.Checked = False
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox1.Checked = True Then
            CheckBox1.Checked = False
        End If
    End Sub

    Private Sub クリアToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles クリアToolStripMenuItem.Click
        DataGridView4.Rows.Clear()
    End Sub

    Private Sub 行削除ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 行削除ToolStripMenuItem.Click
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

    Private Sub NE連携チェックToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NE連携チェックToolStripMenuItem.Click
        FormFront(Dialog)
        Dialog.Size = New Size(620, 700)
        'Dialog.TabControl1.SelectedIndex = 1
        SetTabVisible(Dialog.TabControl1, "連携確認")
    End Sub

    Private Sub NEコードチェックToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NEコードチェックToolStripMenuItem.Click
        FormFront(Dialog)
        Dialog.Size = New Size(750, 700)
        'Dialog.TabControl1.SelectedIndex = 2
        SetTabVisible(Dialog.TabControl1, "コード確認")
    End Sub




    'フォーム自動化

    Private Sub 管理ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 管理ToolStripMenuItem.Click
        Kanri.Show()
        Kanri.BringToFront()
        'Kanri_F_login.ShowDialog()
        'Kanri_F_login.MaskedTextBox1.Focus()
        'Kanri_F_login.MaskedTextBox1.SelectAll()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            BackgroundWorker1.RunWorkerAsync()
        Catch ex As Exception

        End Try
    End Sub

    Private Delegate Sub CallDelegate()
    Private Shared textStr1 As String = ""
    Private Shared textStr2 As String = ""
    Private Shared st_textStr1 As String = ""
    Private Shared st_textStr2 As String = ""
    Private Shared st_textStr3 As String = ""
    Private Shared lb_text3 As String = ""
    Private Shared tsp_menu1 As String = ""
    Private Shared tsp_menu2 As String = ""

    Private Sub MeText()
        Me.Text = textStr1
    End Sub
    Private Sub ToolText()
        ToolStripStatusLabel2.Text = textStr2
    End Sub

    Private Sub TBText1()
        TextBox8.Text = st_textStr1
        TextBox7.Text = st_textStr2
    End Sub
    Private Sub TBText3()
        TextBox6.Text = st_textStr3
    End Sub
    Private Sub LBAdd1()
        ListBox3.Items.Add(lb_text3)
        ListBox3.SelectedIndex = ListBox3.Items.Count - 1
    End Sub
    Private Sub TSPMENU1()
        サーバーToolStripMenuItem.Text = tsp_menu1
        ロケーションToolStripMenuItem.Text = tsp_menu2
    End Sub
    Private Sub Button21_change()
        Button21.Enabled = True
        Button21.BackColor = Color.Yellow
    End Sub

    Public Shared updateFlag As Boolean = True
    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        If Not BackgroundWorker1.IsBusy Then
            Exit Sub
        ElseIf UpdatePanel.Visible Then
            Exit Sub
        End If

        'Try
        Dim severPath As String = ""      'サーバーのフォルダーパス
        Dim locationPath As String = "" 'ロケーションプログラムのフォルダーパス
        Dim mLoadFile As String = Path.GetDirectoryName(appPath) & "\config\version.txt"
        Dim sLoadFile As String = ""

        Dim dirLoadFile As String = Path.GetDirectoryName(appPath) & "\config\folder.txt"
        Dim flagTxt As String() = File.ReadAllLines(dirLoadFile, Encoding.GetEncoding("Shift_JIS"))
        Dim flag As Integer = 0
        For i As Integer = 0 To flagTxt.Length - 1
            flagTxt(i) = RTrim(flagTxt(i))
            If flagTxt(i) <> "" Then
                If InStr(flagTxt(i), "[Data Folder]") > 0 Then
                    flag = 1
                ElseIf InStr(flagTxt(i), "[Location Folder]") > 0 Then
                    flag = 2
                Else
                    If flag = 1 Then
                        If InStr(flagTxt(i), "SERVER") > 0 Then 'サーバーのアドレスを参照してるか確認
                            severPath = flagTxt(i)
                        Else
                            If updateFlag = True Then
                                FileRESTORE(dirLoadFile)
                                Exit Sub
                            End If
                        End If
                    ElseIf flag = 2 Then
                        locationPath = flagTxt(i)
                    End If
                End If
            End If
        Next

        tsp_menu1 = severPath
        tsp_menu2 = locationPath
        'If Regex.IsMatch(My.Computer.Name, "takashi", RegexOptions.IgnoreCase) Then
        If Regex.IsMatch(My.Computer.Name, "takashi", RegexOptions.IgnoreCase) Then
                'tsp_menu1 = SpecialDirectories.MyDocuments & "\test"
                tsp_menu1 = "E:\Documents\test"
                tsp_menu2 = SpecialDirectories.MyDocuments & "\test\logi"
                'ElseIf Regex.IsMatch(My.Computer.Name, "NAKA", RegexOptions.IgnoreCase) And Regex.IsMatch(appPath, "debug", RegexOptions.IgnoreCase) Then
                '    tsp_menu1 = SpecialDirectories.MyDocuments & "\test"
                '    tsp_menu2 = SpecialDirectories.MyDocuments & "\test\logi"
            End If
            Me.Invoke(New CallDelegate(AddressOf TSPMENU1))
        'サーバーToolStripMenuItem.Text = dirPath
        'ロケーションToolStripMenuItem.Text = locationPath

        Dim mVer As Integer = 0
        Dim sVer As Integer = 0

        If Regex.IsMatch(Format(Now, "ss"), checkTime) Then
            Dim VerTxt As String() = File.ReadAllLines(Path.GetDirectoryName(appPath) & "\config\version.txt", System.Text.Encoding.GetEncoding("Shift_JIS"))

            Dim fname As String = ""    'テンプレートのバージョン
            Dim fname2 As String = ""   'アプリのバージョン
            Dim fname3 As String = ""   '強制アップデートのバージョン
            Try
                Dim di As New DirectoryInfo(tsp_menu1 & "\update\version")
                Dim files As FileInfo() = di.GetFiles("*.txt", IO.SearchOption.AllDirectories)
                For Each f As FileInfo In files
                    fname = Path.GetFileNameWithoutExtension(f.Name)
                Next
                Dim files2 As FileInfo() = di.GetFiles("*.dat", IO.SearchOption.AllDirectories)
                For Each f As FileInfo In files2
                    fname2 = Path.GetFileNameWithoutExtension(f.Name)
                Next
                Dim files3 As FileInfo() = di.GetFiles("*.ini", IO.SearchOption.AllDirectories)
                For Each f As FileInfo In files3
                    fname3 = Path.GetFileNameWithoutExtension(f.Name)
                Next
            Catch ex As Exception

            End Try

            '強制アップデート
            If IsNumeric(fname3) And updateFlag = True Then
                Dim ap As String = Replace(Application.ProductVersion, ".", "")
                'If fname3 <> ap Then
                If fname3 > ap Then
                    NotifyIcon1.BalloonTipTitle = "強制アップデートのご連絡です"
                    NotifyIcon1.BalloonTipText = "約10秒後にアップデートを行ないます。"
                    NotifyIcon1.ShowBalloonTip(10000)
                    For i As Integer = 0 To 8
                        Threading.Thread.Sleep(1000)
                    Next
                    'MsgBox("強制アップデートです。再起動してください。", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
                    アップデートToolStripMenuItem.PerformClick()
                    Application.DoEvents()
                    Exit Sub
                End If
            End If

            Dim mv As String = Regex.Replace(VerTxt(0), "\[|\]", "")
            If IsNumeric(mv) = True Then
                mVer = mv
            End If
            If IsNumeric(fname) = True Then
                sVer = fname
            End If
            Try
                sLoadFile = tsp_menu1 & "\update\version.txt"
                'If mVer <> sVer And InStr(appPath, "Debug") = 0 Then
                If mVer < sVer And InStr(appPath, "Debug") = 0 Then
                    NotifyIcon1.BalloonTipTitle = "自動アップデートを感知しました"
                    NotifyIcon1.BalloonTipText = "テンプレート系のアップデートを行ないます。再起動の必要はありません。"
                    NotifyIcon1.ShowBalloonTip(5000)
                    lb_text3 = "> 自動アップデート"
                    Me.Invoke(New CallDelegate(AddressOf LBAdd1))
                    DataCopy(sLoadFile, tsp_menu1, Path.GetDirectoryName(appPath), updateFlag)
                    Dim UpdateTxt As String() = File.ReadAllLines(sLoadFile, Encoding.GetEncoding("Shift_JIS"))
                    Dim tempVer As String = UpdateTxt(0)
                    If updateFlag = True Then
                        File.WriteAllText(mLoadFile, tempVer)
                    End If
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            'TEMPファイル保存
            'If mVer < sVer Then
            If mVer <> sVer Then
                ''順番待ち
                'Dim serverPath As String = サーバーToolStripMenuItem.Text & "\update" & "\"
                'If File.Exists(serverPath & "順番待ち.txt") Then
                '    Dim junTxt As String = File.ReadAllText(serverPath & "順番待ち.txt", enc)
                '    Dim junArray As String() = Split(junTxt, vbCrLf)
                '    Dim junFlag As Boolean = False
                '    For i As Integer = 0 To junArray.Length - 1
                '        Dim jA As String() = Split(junArray(i), ",")
                '        If jA(0) = ログイン名ToolStripMenuItem.Text Then
                '            junFlag = True
                '            Exit For
                '        End If
                '    Next
                '    If Not junFlag Then
                '        If junArray.Length > 5 Then
                '            Exit Sub
                '        End If
                '    End If
                'Else
                '    Dim str As String = ログイン名ToolStripMenuItem.Text & "," & Now & vbCrLf
                '    File.WriteAllText(serverPath & "順番待ち.txt", str, enc)
                'End If

                Dim tempDir As String = SpecialDirectories.MyDocuments & "\taksoft\AutoCheck"
                If Not Directory.Exists(tempDir) Then
                    Directory.CreateDirectory(tempDir)
                End If
                Dim tempDirSub As String() = {"config", "config\save_template", "config\version2", "db", "formchange", "Image", "template", "version", "version2"}
                For i As Integer = 0 To tempDirSub.Length - 1
                    If Not Directory.Exists(tempDir & "\" & tempDirSub(i)) Then
                        Directory.CreateDirectory(tempDir & "\" & tempDirSub(i))
                    End If
                Next

                'uploadフォルダ番号（ランダム）
                'Dim rndStr As String = ""
                'Dim r As New System.Random()
                'Dim rndNum As Integer = r.Next(1, 4)
                'If rndNum = 1 Then
                '    rndStr = ""
                'Else
                '    rndStr = rndNum
                'End If

                sLoadFile = tsp_menu1 & "\update\version.txt"
                Dim UpdateTxt As String() = File.ReadAllLines(sLoadFile, Encoding.GetEncoding("Shift_JIS"))
                For i As Integer = 1 To UpdateTxt.Length - 1
                    If UpdateTxt(i) <> "" Then
                        If InStr(UpdateTxt(i), "[del]") > 0 Then
                            Dim delFile As String = Replace(UpdateTxt(i), "[del]", "")
                            Dim delDest As String = tempDir & "\" & delFile
                            If File.Exists(delDest) = True Then
                                File.Delete(delDest)
                                lb_text3 = "> [t_del]" & Path.GetFileName(delDest)
                                Me.Invoke(New CallDelegate(AddressOf LBAdd1))
                            End If
                        Else
                            Dim upFile As String = Replace(UpdateTxt(i), "[all]", "")
                            Dim copySrc As String = サーバーToolStripMenuItem.Text & "\update" & "\" & upFile
                            Dim copyDest As String = tempDir & "\" & upFile
                            If Not Directory.Exists(Path.GetDirectoryName(copyDest)) Then
                                Directory.CreateDirectory(Path.GetDirectoryName(copyDest))
                            End If
                            If File.GetLastWriteTime(copySrc) > File.GetLastWriteTime(copyDest) Then    '更新時間
                                File.Copy(copySrc, copyDest, True)
                                lb_text3 = "> [t]" & Path.GetFileName(copyDest)
                                Me.Invoke(New CallDelegate(AddressOf LBAdd1))
                            End If
                        End If
                    End If
                Next
                File.Copy(sLoadFile, tempDir & "\version.txt", True)
                '開発環境のみこっちでバージョン.txtをアップデートする
                If InStr(appPath, "Debug") > 0 Then
                    Dim tempVer As String = UpdateTxt(0)
                    File.WriteAllText(mLoadFile, tempVer)
                End If

                lb_text3 = "> tempコピー完了"
                Me.Invoke(New CallDelegate(AddressOf LBAdd1))
            End If

            Dim updateMes As String = ""
            If IsNumeric(fname2) = True Then
                Dim ap As String = Replace(Application.ProductVersion, ".", "")
                If fname2 - ap > 0 Then
                    updateMes = "（ソフトのアップデートがあります）"
                    'Button21.Enabled = True
                    'Button21.BackColor = Color.Yellow
                    Me.Invoke(New CallDelegate(AddressOf Button21_change))
                End If
            End If

            Dim adminMes As String = ""
            If AdminFlag Then
                adminMes = "（管理モード）"
            End If
            textStr1 = "くま Ver." & Application.ProductVersion & adminMes & "(" & fname2 & ")" & " - ap" & VerTxt(0) & "/sv[" & fname & "] " & updateMes
            Me.Invoke(New CallDelegate(AddressOf MeText))
            st_textStr1 = Application.ProductVersion & "(" & fname2 & ")"
            st_textStr2 = "ap" & VerTxt(0) & "/sv[" & fname & "] "
            Me.Invoke(New CallDelegate(AddressOf TBText1))
        End If

        st_textStr3 = Format(Now, "yyyy/MM/dd hh:mm:ss")
        Me.Invoke(New CallDelegate(AddressOf TBText3))
        textStr2 = Format(Now, "hh:mm:ss")
        Me.Invoke(New CallDelegate(AddressOf ToolText))
    End Sub


    Private Sub DataCopy(ByVal verPath As String, ByVal mDir As String, ByVal sDir As String, ByVal updateFlag As Boolean)
        Dim UpdateTxt As String() = File.ReadAllLines(verPath, System.Text.Encoding.GetEncoding("Shift_JIS"))
        For i As Integer = 0 To UpdateTxt.Length - 1
            Try
                If UpdateTxt(i) <> "" Then
                    Select Case True
                        Case Regex.IsMatch(UpdateTxt(i), "dat$|txt$|csv$|xlsx$|xls$|jpg$|html$|dll$|xml$")
                            If InStr(UpdateTxt(i), "[del]") > 0 Then
                                Dim delFile As String = Replace(UpdateTxt(i), "[del]", "")
                                Dim delDest As String = sDir & "\" & delFile
                                If updateFlag = True Then
                                    If File.Exists(delDest) = True Then
                                        File.Delete(delDest)
                                    End If
                                End If
                            Else
                                'UpdateTxt(i) = Replace(UpdateTxt(i), "[all]", "")
                                Dim copySrc As String = mDir & "\update\" & UpdateTxt(i)
                                Dim copyDest As String = sDir & "\" & UpdateTxt(i)
                                If File.GetLastWriteTime(copySrc) > File.GetLastWriteTime(copyDest) Then    '更新時間
                                    lb_text3 = "> " & Path.GetFileName(copyDest)
                                    Me.Invoke(New CallDelegate(AddressOf LBAdd1))
                                    If updateFlag = True Then
                                        If File.Exists(copyDest) Then
                                            File.Copy(copySrc, copyDest, True)
                                        Else
                                            My.Computer.FileSystem.CopyFile(copySrc, copyDest, True)
                                        End If
                                    End If
                                End If
                            End If
                        Case Else
                            If InStr(UpdateTxt(i), "[all]") > 0 Then
                                UpdateTxt(i) = Replace(UpdateTxt(i), "[all]", "")
                                Dim copySrc As String = mDir & "\update\" & UpdateTxt(i)
                                Dim copyDest As String = sDir & "\" & UpdateTxt(i)
                                If File.GetLastWriteTime(copySrc) > File.GetLastWriteTime(copyDest) Then    '更新時間
                                    lb_text3 = "> " & Path.GetFileName(copyDest)
                                    Me.Invoke(New CallDelegate(AddressOf LBAdd1))
                                    If updateFlag = True Then
                                        If File.Exists(copyDest) Then
                                            File.Copy(copySrc, copyDest, True)
                                        Else
                                            My.Computer.FileSystem.CopyFile(copySrc, copyDest, True)
                                        End If
                                    End If
                                End If
                            End If
                    End Select
                End If
            Catch ex As Exception

            End Try
        Next i

        lb_text3 = "> 自動更新終了しました"
        Me.Invoke(New CallDelegate(AddressOf LBAdd1))
        'ToolStripStatusLabel2.Text = ""
    End Sub

    Private Sub テンプレート自動更新ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles テンプレート自動更新ToolStripMenuItem.Click
        If テンプレート自動更新ToolStripMenuItem.Checked = True Then
            テンプレート自動更新ToolStripMenuItem.Checked = False
        Else
            テンプレート自動更新ToolStripMenuItem.Checked = True
        End If

        Form1.inif.WriteINI(secValue, "AutoUpdate", テンプレート自動更新ToolStripMenuItem.Checked)
    End Sub

    Private Sub テンプレートのみ更新ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles テンプレートのみ更新ToolStripMenuItem.Click
        Dim dirPath As String = ""  'サーバーのフォルダーパス
        Dim mLoadFile As String = Path.GetDirectoryName(appPath) & "\config\version.txt"
        Dim sLoadFile As String = ""

        'Dim mTS As DateTime
        'Dim sTS As DateTime

        Dim dirLoadFile As String = Path.GetDirectoryName(appPath) & "\config\folder.txt"
        Dim flagTxt As String() = File.ReadAllLines(dirLoadFile, System.Text.Encoding.GetEncoding("Shift_JIS"))
        Dim flag As Boolean = False
        For i As Integer = 0 To flagTxt.Length - 1
            If InStr(flagTxt(i), "[Data Folder]") > 0 Then
                flag = True
            Else
                If flag = True Then
                    dirPath = flagTxt(i)
                    Exit For
                End If
            End If
        Next

        sLoadFile = dirPath & "\update\version.txt"

        Try
            'If File.Exists(mLoadFile) Then
            '    mTS = File.GetLastWriteTime(mLoadFile)
            'End If
            'If File.Exists(sLoadFile) Then
            '    sTS = File.GetLastWriteTime(sLoadFile)
            'End If
            'If mTS < sTS Then   'インストールフォルダのversion.txtにバージョンを書き込む
            Dim UpdateTxt As String() = File.ReadAllLines(sLoadFile, System.Text.Encoding.GetEncoding("Shift_JIS"))
            Dim tempVer As String = UpdateTxt(0)
            File.WriteAllText(mLoadFile, tempVer)
            'End If

            Dim updateFlag As Boolean = True
            If InStr(Path.GetDirectoryName(appPath), "Debug") > 0 Then
                updateFlag = False
            End If
            DataCopy(sLoadFile, dirPath, Path.GetDirectoryName(appPath), updateFlag)
            MsgBox("更新終了", MsgBoxStyle.SystemModal)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.SystemModal)
        End Try
    End Sub

    Private Sub アップデータ更新ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles アップデータ更新ToolStripMenuItem.Click
        OneDataUpdate("\update\ACupdate.exe", "\ACupdate.exe")
    End Sub

    Private Sub MiseDBmdb更新ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MiseDBmdb更新ToolStripMenuItem.Click
        OneDataUpdate("\update\db\miseDB.accdb", "\db\miseDB.accdb")
    End Sub

    Private Sub AmazonToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AmazonToolStripMenuItem.Click
        OneDataUpdate("\update\db\AmazonNode.accdb", "\db\AmazonNode.accdb")
    End Sub

    Private Sub 住所データベース更新ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 住所データベース更新ToolStripMenuItem.Click
        MsgBox("住所データベースを更新します" & vbCrLf & "終わったらダイアログが出ます", MsgBoxStyle.SystemModal)
        OneDataUpdate("\update\db\ken_all.accdb", "\db\ken_all.accdb", 1)
        OneDataUpdate("\update\db\sagawa.accdb", "\db\sagawa.accdb", 1)
        OneDataUpdate("\update\db\Address.accdb", "\db\Address.accdb", 0)
    End Sub

    Private Sub OneDataUpdate(sPath As String, dPath As String, Optional endFlag As Integer = 0)
        Dim dirPath As String = ""  'サーバーのフォルダーパス
        Dim dirLoadFile As String = Path.GetDirectoryName(appPath) & "\config\folder.txt"
        Dim flagTxt As String() = File.ReadAllLines(dirLoadFile, System.Text.Encoding.GetEncoding("Shift_JIS"))
        Dim flag As Boolean = False
        For i As Integer = 0 To flagTxt.Length - 1
            If InStr(flagTxt(i), "[Data Folder]") > 0 Then
                flag = True
            Else
                If flag = True Then
                    dirPath = flagTxt(i)
                    Exit For
                End If
            End If
        Next

        Dim copySrc As String = dirPath & sPath
        Dim copyDest As String = Path.GetDirectoryName(appPath) & dPath
        File.Copy(copySrc, copyDest, True)

        If endFlag = 0 Then
            MsgBox("更新しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    Private Sub FileRESTORE(ByVal restorePath As String)
        Dim refileName As String = Path.GetFileNameWithoutExtension(restorePath)
        MsgBox("設定ファイルが正しくないため自己修復します。" & vbCrLf & refileName, MsgBoxStyle.Information Or MsgBoxStyle.SystemModal)
        Dim reServerSrc As String = "\\SERVER2-PC\Users\Public\program\autocheck\update\config\folder.txt"     'リストア用ファイルディレクトリは決め打ち
        Dim copyDest As String = Path.GetDirectoryName(appPath) & "\config\folder.txt"
        File.Copy(reServerSrc, copyDest, True)
    End Sub



    Private Sub Csv1ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Csv1ToolStripMenuItem.Click
        FormFront(Form1.CSV_FORMS(0))
    End Sub

    Private Sub Csv2ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Csv2ToolStripMenuItem.Click
        FormFront(Form1.CSV_FORMS(1))
    End Sub

    Private Sub ファイル名操作ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ファイル名操作ToolStripMenuItem.Click
        FormFront(HTMLdialog)
        HTMLdialog.Size = New Size(609, 187)
        SetTabVisible(HTMLdialog.TabControl2, "ファイル名操作")
    End Sub

    Private Sub ファイル名操作2ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ファイル名操作2ToolStripMenuItem.Click
        FormFront(Dialog)
        SetTabVisible(Dialog.TabControl1, "ファイル名操作")
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        FormFront(Dialog)
        SetTabVisible(Dialog.TabControl1, "ファイル名操作")
    End Sub

    Private Sub ToolStripComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        ToolStripButton4.PerformClick()
    End Sub

    Private Sub 新しいタブを開くToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 新しいタブを開くToolStripMenuItem.Click
        If TabBrowser1.SelectedTab.WebBrowser.ReadyState >= WebBrowserReadyState.Loading Then
            WebTabPage.WebBrowser_Navigating(TabBrowser1.SelectedTab.WebBrowser.Url.ToString)
            TabBrowser1.SelectedTab.WebBrowser.Navigate(TabBrowser1.SelectedTab.WebBrowser.Url, True)
        End If
    End Sub

    Private Sub プロパティToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles プロパティToolStripMenuItem.Click
        Shell("rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl")
    End Sub

    Private Sub ミニサイズToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ミニサイズToolStripMenuItem.Click
        'Me.Hide()
        'Start.Show()
        Form1.inif.WriteINI(secValue, "StartForm", "mini")
        SplitContainer2.SplitterDistance = 168
        Me.Size = New Size(183, 520)
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedDialog
        StatusStrip1.Visible = False
    End Sub

    Private Sub ToolStripButton9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton9.Click
        ミニサイズToolStripMenuItem.PerformClick()
    End Sub

    Public Sub FormFront(ByVal f As Form)
        f.Show()
        If f.WindowState = FormWindowState.Minimized Then
            f.WindowState = FormWindowState.Normal
        End If
        f.BringToFront()
    End Sub

    '**************************************************************************************************************
    Private Sub Button34_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button34.Click
        FormFront(Form2)
    End Sub

    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        WindowBig()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        FormFront(Csv_denpyo3)
    End Sub

    Private Sub Button33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button33.Click
        FormFront(CSV_FORMS(0))
    End Sub

    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        FormFront(CSV_FORMS(1))
    End Sub

    Private Sub Button31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button31.Click
        FormFront(Csv_change)
    End Sub

    Private Sub Button32_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button32.Click
        FormFront(HTMLcreate)
    End Sub

    Private Sub Button27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button27.Click
        'HTMLcreate.商品テーブル作成ToolStripMenuItem.PerformClick()
        HtmlDialog_F_TableCreate.Show()
    End Sub

    Private Sub Button28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button28.Click
        HTMLcreate.Itemlist更新ToolStripMenuItem.PerformClick()
    End Sub

    Private Sub Button29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        WindowBig()
        'newitem.Show()
        FormFront(Form1_F_Newitem)
        Form1_F_Newitem.Location = New Point(Me.Location.X + Me.Size.Width, Me.Location.Y)
        Form1_F_Newitem.TabControl1.SelectedTab = Form1_F_Newitem.TabControl1.TabPages("TabPage7")
    End Sub

    Private Sub Button26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button26.Click
        HTMLdialog.Show()
        'HTMLdialog.Size = New Size(439, 371)
        HTMLdialog.Size = New Size(550, 480)
        SetTabVisible(HTMLdialog.TabControl2, "計算機")
    End Sub

    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button25.Click
        WindowBig()
        'newitem.Show()
        FormFront(Form1_F_Newitem)
        Form1_F_Newitem.Location = New Point(Me.Location.X + Me.Size.Width, Me.Location.Y)
        Form1_F_Newitem.TabControl1.SelectedTab = Form1_F_Newitem.TabControl1.TabPages("TabPage1")
    End Sub

    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        WindowBig()
        'newitem.Show()
        FormFront(Form1_F_Newitem)
        Form1_F_Newitem.Location = New Point(Me.Location.X + Me.Size.Width, Me.Location.Y)
        Form1_F_Newitem.TabControl1.SelectedTab = Form1_F_Newitem.TabControl1.TabPages("TabPage5")
    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            Me.TopMost = True
            My.Settings("StartTopMost") = True
        Else
            Me.TopMost = False
            My.Settings("StartTopMost") = False
        End If
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        'Dim Ret As Long = Shell(Path.GetDirectoryName(appPath) & "\ACupdate.exe", vbNormalFocus)
        'End
        UpdatePanel.ShowDialog()
        Application.DoEvents()
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Form1.inif.WriteINI(secValue, "StartForm", "")
        WindowBig()
    End Sub

    Private Sub WindowBig()
        SplitContainer2.SplitterDistance = 168
        If Me.Size.Width - 950 < 0 Then
            Me.Size = New Size(950, 650)
        ElseIf Me.Size.Height - 650 < 0 Then
            Me.Size = New Size(950, 650)
        End If
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
        StatusStrip1.Visible = True
    End Sub

    Private Sub モールチェックToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles モールチェックToolStripMenuItem.Click
        FormFront(Dialog)
        'Dialog.TabControl1.SelectTab("TabPage6")
        SetTabVisible(Dialog.TabControl1, "モールチェック")
        Dialog.Size = New Size(818, 569)
    End Sub

    Private Sub サーバーToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles サーバーToolStripMenuItem.Click
        MsgBox(サーバーToolStripMenuItem.Text)
    End Sub

    Private Sub ロケーションToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ロケーションToolStripMenuItem.Click
        MsgBox(ロケーションToolStripMenuItem.Text)
    End Sub

    Private Sub 再起動ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 再起動ToolStripMenuItem.Click
        CloseForm(1)
        'Application.Restart()
        Process.Start(Path.GetDirectoryName(appPath) & "\AutoCheck3.exe")
        Application.Exit()
        'End
    End Sub

    Private Sub 終了ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 終了ToolStripMenuItem1.Click
        Me.Close()
        Me.Dispose()
        Application.Exit()
        'End
    End Sub

    Private Sub ウィンドウを開くToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ウィンドウを開くToolStripMenuItem.Click
        Me.ShowInTaskbar = True
        Me.WindowState = FormWindowState.Normal
    End Sub

    Private Sub NotifyIcon1_DoubleClick(sender As Object, e As EventArgs) Handles NotifyIcon1.DoubleClick
        Me.ShowInTaskbar = True
        Me.WindowState = FormWindowState.Normal
    End Sub

    Private Sub ToolStripDropDownButton6_DropDownOpening(sender As Object, e As EventArgs) Handles ToolStripDropDownButton6.DropDownOpening
        '各種ソフト一覧読み込み
        Try
            '\\SERVER2-PC\files\ツール\soft.txt
            Dim serverPath As String = サーバーToolStripMenuItem.Text
            serverPath = Replace(serverPath, "\Users\Public\program\autocheck", "\files\ツール\soft.txt")
            If updateFlag = False Then
                serverPath = Path.GetDirectoryName(appPath) & "\soft.txt"
            End If
            ToolStripDropDownButton6.DropDownItems.Clear()
            Using sr As New StreamReader(serverPath, Encoding.Default)
                Do While Not sr.EndOfStream
                    Dim s As String = sr.ReadLine
                    Dim sArray As String() = Split(s, ",")
                    If sArray(0) = "---" Then
                        ToolStripDropDownButton6.DropDownItems.Add("-")
                    Else
                        ToolStripDropDownButton6.DropDownItems.Add(sArray(1))
                        Dim num As Integer = ToolStripDropDownButton6.DropDownItems.Count - 1
                        ToolStripDropDownButton6.DropDownItems(num).Tag = s
                        ToolStripDropDownButton6.DropDownItems(num).ToolTipText = sArray(3)
                        AddHandler ToolStripDropDownButton6.DropDownItems(num).Click, AddressOf TSDD_DDO_CLICK
                    End If
                Loop
            End Using
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TSDD_DDO_CLICK(sender As Object, e As EventArgs)
        Dim tsdd As ToolStripDropDownItem = sender
        Dim sArray As String() = Split(tsdd.Tag, ",")
        Dim DR As DialogResult = MsgBox(sArray(1) & "をインストールまたはコピーします", MsgBoxStyle.OkCancel)
        If DR = DialogResult.Cancel Then
            Exit Sub
        End If

        Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        Dim serverPath As String = サーバーToolStripMenuItem.Text
        serverPath = Replace(serverPath, "\Users\Public\program\autocheck", "\files\ツール")
        'If updateFlag = False Then
        '    serverPath = desktopPath & "\デスクトップ"
        'End If

        If InStr(sArray(0), "copy") > 0 Then    '指定フォルダにコピーするだけ
            Dim copyDestPath As String = ""
            Dim place As String() = Split(sArray(0), "=")
            Select Case place(1)
                Case "mydocument"
                    copyDestPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal)
                Case "desktop"
                    copyDestPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
            End Select
            Dim copySrcPath As String = serverPath & "\" & sArray(2)
            copyDestPath = copyDestPath & "\" & sArray(2)

            If File.Exists(copyDestPath) Then
                DR = MsgBox("既にインストールされています。上書きしますか？", MsgBoxStyle.OkCancel)
                If DR = DialogResult.Cancel Then
                    Exit Sub
                End If
                File.Copy(copySrcPath, copyDestPath)
            ElseIf Directory.Exists(copyDestPath) Then
                DR = MsgBox("既にインストールされています。上書きしますか？", MsgBoxStyle.OkCancel)
                If DR = DialogResult.Cancel Then
                    Exit Sub
                End If
                My.Computer.FileSystem.CopyDirectory(copySrcPath, copyDestPath, True)
            Else
                My.Computer.FileSystem.CopyDirectory(copySrcPath, copyDestPath, True)
            End If

            If InStr(sArray(4), "shortcut") > 0 Then
                Dim scSrcPath As String() = Split(sArray(4), "=")
                Dim scSrcName As String = Path.GetFileNameWithoutExtension(scSrcPath(1))
                Dim objShell As Object = CreateObject("WScript.Shell")
                Dim objLink As Object
                objLink = objShell.CreateShortcut(desktopPath & "\" & scSrcName & ".lnk")
                With objLink
                    .targetpath = copyDestPath & "\" & scSrcPath(1)
                    .description = sArray(3)
                    .iconlocation = copyDestPath & "\" & scSrcPath(1)   'Path.GetFileName(scSrcPath(1))
                    .workingdirectory = copyDestPath
                    .save()
                End With
            End If
        ElseIf InStr(sArray(0), "install") > 0 Then     'インストーラー開始
            If GetInstallList(sArray(4)) Then
                DR = MsgBox("既にインストールしていますが、上書きインストールしますか？", MsgBoxStyle.OkCancel Or MsgBoxStyle.SystemModal)
                If DR = DialogResult.Cancel Then
                    Exit Sub
                End If
            End If
            Try
                Dim installDestPath As String = serverPath & "\" & sArray(2)
                Process.Start(installDestPath)
            Catch ex As Exception
                Exit Sub
            End Try
        End If

        'MsgBox("インストール完了しました。" & vbCrLf & sArray(5))
    End Sub


    Private Function GetInstallList(softName As String) As Boolean
        Dim wkRegKey As Microsoft.Win32.RegistryKey
        Dim wkKeyName As String
        Dim strKeyNames() As String
        Dim strKeyName As String
        Dim rKey As Microsoft.Win32.RegistryKey
        Dim displayName As String

        wkKeyName = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
        wkRegKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(wkKeyName, False)
        strKeyNames = wkRegKey.GetSubKeyNames()

        For Each strKeyName In strKeyNames
            rKey = wkRegKey.OpenSubKey(strKeyName)
            'インストールソフトウェア名取得
            If Not rKey.GetValue("DisplayName") Is Nothing Then
                displayName = rKey.GetValue("DisplayName").ToString()
                If InStr(displayName.ToLower, softName.ToLower) > 0 Then
                    Return True
                    Exit Function
                End If
            End If
        Next

        Return False
    End Function


    'TabColtrol表示切替
    Public Sub SetTabVisible(ByVal oTabControl As TabControl, ByVal tText As String)
        Dim nameArray As ArrayList = New ArrayList
        If oTabControl.Tag Is Nothing Then
            For i As Integer = 0 To oTabControl.TabPages.Count - 1
                nameArray.Add(oTabControl.TabPages(i).Text)
            Next
        Else
            Dim oAllTabPages As List(Of KeyValuePair(Of TabPage, Boolean)) = CType(oTabControl.Tag, List(Of KeyValuePair(Of TabPage, Boolean)))
            For Each oTabAndVisible As KeyValuePair(Of TabPage, Boolean) In oAllTabPages
                nameArray.Add(oTabAndVisible.Key.Text)
            Next
        End If

        For i As Integer = 0 To nameArray.Count - 1
            If Regex.IsMatch(nameArray(i), tText) Then
                SetTabVisible2(oTabControl, i, True)
            Else
                SetTabVisible2(oTabControl, i, False)
            End If
        Next
    End Sub

    Public Sub SetTabVisible2(ByVal oTabControl As TabControl, ByVal nIndex As Integer, ByVal bVisible As Boolean)
        Dim oAllTabPages As List(Of KeyValuePair(Of TabPage, Boolean))
        If oTabControl.Tag Is Nothing Then
            ' 全タブと、その表示状態を保持
            oAllTabPages = New List(Of KeyValuePair(Of TabPage, Boolean))
            For Each oTabPage As TabPage In oTabControl.TabPages
                Dim oTabAndVisible As New KeyValuePair(Of TabPage, Boolean)(oTabPage, True)
                oAllTabPages.Add(oTabAndVisible)
            Next
            oTabControl.Tag = oAllTabPages
        Else
            ' 全タブと、その表示状態を取得
            oAllTabPages = CType(oTabControl.Tag, List(Of KeyValuePair(Of TabPage, Boolean)))
        End If

        ' タブの表示状態を設定
        oAllTabPages(nIndex) = New KeyValuePair(Of TabPage, Boolean)(oAllTabPages(nIndex).Key, bVisible)

        oTabControl.TabPages.Clear()
        For Each oTabAndVisible As KeyValuePair(Of TabPage, Boolean) In oAllTabPages
            If oTabAndVisible.Value = True Then
                oTabControl.TabPages.Add(oTabAndVisible.Key)
            End If
        Next
    End Sub

    Private Sub 伝票変換ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 伝票変換ToolStripMenuItem.Click
        FormFront(Csv_denpyo3)
    End Sub

    Private Sub CSV編集ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CSV編集ToolStripMenuItem.Click
        FormFront(Csv)
    End Sub

    Private Sub 特殊処理ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 特殊処理ToolStripMenuItem1.Click
        FormFront(Csv_change)
    End Sub

    Private Sub 楽くまDBToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 楽くまDBToolStripMenuItem.Click
        FormFront(Mall_main)
    End Sub

    Private Sub 電話対応ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 電話対応ToolStripMenuItem.Click
        FormFront(Form2)
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        FormFront(Mall_main)
    End Sub

    Private Sub OS起動時スタートToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OS起動時スタートToolStripMenuItem.Click
        If OS起動時スタートToolStripMenuItem.Checked Then
            OS起動時スタートToolStripMenuItem.Checked = False
            DelCurrentVersionRun()
        Else
            OS起動時スタートToolStripMenuItem.Checked = True
            SetCurrentVersionRun()
        End If
    End Sub

    Private Sub テストToolStripMenuItem_Click(sender As Object, e As EventArgs)
        FormFront(Mall_main)
    End Sub

    Private Sub ソースToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ソースToolStripMenuItem.Click
        FormFront(HTMLcreate)
        HTMLcreate.AzukiControl1.Text = Me.TabBrowser1.SelectedTab.WebBrowser.Document.Body.OuterHtml
    End Sub

    Private Sub レンダリングバージョン書き込みToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles レンダリングバージョン書き込みToolStripMenuItem.Click
        Dim strRegPath As String = "Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION"
        Dim regKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(strRegPath, True)
        Dim strProcessName As String = Process.GetCurrentProcess().ProcessName + ".exe"
        regKey.SetValue(strProcessName, 11001, Microsoft.Win32.RegistryValueKind.DWord)
        Me.Dispose()
        Dim frm As Form = New Form1
        frm.Show()
    End Sub

    Private Sub 全てToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 全てToolStripMenuItem.Click
        Dim saveFolder As String = ""
        Dim dlg = VistaFolderBrowserDialog1
        dlg.RootFolder = Environment.SpecialFolder.Desktop
        dlg.Description = "画像を保存するフォルダを選択してください"
        If dlg.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
            saveFolder = dlg.SelectedPath
        End If
        If saveFolder = "" Then
            Exit Sub
        End If

        TM_IMGDL_LOGIC("")
    End Sub

    Private Sub 型番指定ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 型番指定ToolStripMenuItem.Click
        Dim c As String = Path.GetFileNameWithoutExtension(ToolStripTextBox1.Text)
        Dim code As String = InputBox("ダウンロードする画像の型番を指定してください", "画像ダウンロード", c)
        If code = "" Then
            Exit Sub
        End If


        TM_IMGDL_LOGIC(code)
    End Sub

    Private Sub TM_IMGDL_LOGIC(Optional code As String = "")
        Dim saveFolder As String = ""
        Dim dlg = VistaFolderBrowserDialog1
        dlg.RootFolder = Environment.SpecialFolder.Desktop
        dlg.Description = "画像を保存するフォルダを選択してください"
        If dlg.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
            saveFolder = dlg.SelectedPath
        End If
        If saveFolder = "" Then
            Exit Sub
        End If

        Dim dlCount As Integer = 0
        'Dim pattern As String = "<img src=""(?<text>.*?)"".*>|<img src='(?<text>.*?)'.*>" '|<img src=(?<text>.*?) ?.*>"
        Dim pattern As String = "(?<text>http.*?(\.jpg|\.jpeg|\.gif|\.png|\.tif|\.bmp|""|'))"
        Dim source As String = Me.TabBrowser1.SelectedTab.WebBrowser.Document.Body.OuterHtml
        Dim sourceArray As String() = Regex.Split(source, "\n|\r|\n\r")
        For Each s As String In sourceArray
            If InStr(s, code) > 0 And (InStr(s, ".jpg") > 0 Or InStr(s, ".gif") > 0) Then
                Dim m As MatchCollection = Regex.Matches(s, pattern, RegexOptions.IgnoreCase Or RegexOptions.Singleline)
                For i As Integer = 0 To m.Count - 1
                    Dim res As String = m(i).Groups("text").Value
                    If InStr(res, code) > 0 And (InStr(s, ".jpg") > 0 Or InStr(s, ".gif") > 0) Then
                        If TM_IMGDL(saveFolder, res) = True Then
                            dlCount += 1
                        End If
                    End If
                Next
            End If
        Next

        MsgBox("ダウンロードしました")
    End Sub

    Private Sub HTML操作ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HTML操作ToolStripMenuItem.Click
        WindowBig()
        FormFront(Form1_F_sousa)
        Form1_F_sousa.Location = New Point(Me.Location.X + Me.Size.Width, Me.Location.Y)
    End Sub

    Private Sub テスト中ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoginListToolStripMenuItem.Click
        Form1_F_loginlist.Show()
    End Sub

    Private Sub 在庫チェックリストToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 在庫チェックリストToolStripMenuItem.Click
        FormFront(ZaikoList)
    End Sub

    Private Sub Form1_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        If Me.ShowInTaskbar = False Then
            Me.ShowInTaskbar = True
        End If

        TabBrowser1.Enabled = True
    End Sub

    Private Sub Form1_Deactivate(sender As Object, e As EventArgs) Handles Me.Deactivate
        TabBrowser1.Enabled = False
    End Sub

    'UserAgent変更
    Private Sub 自動ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 自動ToolStripMenuItem.Click
        自動ToolStripMenuItem.Checked = True
        IEToolStripMenuItem.Checked = False
        ChromeToolStripMenuItem.Checked = False
        TabBrowser1.SelectedTab.WebBrowser.Refresh()
    End Sub

    Private Sub IEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles IEToolStripMenuItem.Click
        自動ToolStripMenuItem.Checked = False
        IEToolStripMenuItem.Checked = True
        ChromeToolStripMenuItem.Checked = False
        TabBrowser1.SelectedTab.WebBrowser.Refresh()
    End Sub

    Private Sub ChromeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ChromeToolStripMenuItem.Click
        自動ToolStripMenuItem.Checked = False
        IEToolStripMenuItem.Checked = False
        ChromeToolStripMenuItem.Checked = True
        TabBrowser1.SelectedTab.WebBrowser.Refresh()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        WindowBig()
        FormFront(Form1_F_sousa)
        Form1_F_sousa.Location = New Point(Me.Location.X + Me.Size.Width, Me.Location.Y)
    End Sub

    Private Sub 荷物追跡ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 荷物追跡ToolStripMenuItem1.Click
        FormFront(Tuiseki)
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        FormFront(Tuiseki)
    End Sub

    Private Sub ショートカットの修理ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ショートカットの修理ToolStripMenuItem.Click
        UpdatePanel.ShortcutCheck()
    End Sub

    Private Sub ファイルから開くToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ファイルから開くToolStripMenuItem.Click
        Dim ofd As New OpenFileDialog With {
            .InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
            .Filter = "htmlファイル(*.html)|*.html|すべてのファイル(*.*)|*.*",
            .FilterIndex = 1,
            .Title = "開くファイルを選択してください",
            .CheckFileExists = True,
            .CheckPathExists = True
        }

        If ofd.ShowDialog() = DialogResult.OK Then
            TabBrowser1.SelectedTab.WebBrowser.Url = New Uri(ofd.FileName)
        End If
    End Sub

    Private Sub ToolStripMenuItem20_Click(sender As Object, e As EventArgs)
        MailManager.Show()
    End Sub

    Private Sub Qoo10送信ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Qoo10チェックToolStripMenuItem.Click
        Qoo10_mail.Show()
    End Sub

    Private Sub 佐川送料比較ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 佐川送料比較ToolStripMenuItem.Click
        SagawaSpare.Show()
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        If Me.Timer1.Enabled Then
            Me.Timer1.Enabled = False
            Button14.BackColor = Color.LightPink
        Else
            Me.Timer1.Enabled = True
            Button14.BackColor = Color.PaleGreen
        End If
    End Sub

    Private Sub フリー在庫不足ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles フリー在庫不足ToolStripMenuItem.Click
        Freezaiko.Show()
    End Sub


    'Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
    'Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hwndParent As Integer, ByVal hwndChildAfter As Integer, ByVal lpszClass As String, ByVal lpszWindow As String) As Integer
    'Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Integer, ByVal MSG As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
End Class


'#############################################################################################
' WebTabControl
'#############################################################################################
Public Class WebTabControl
    Inherits CustomTabControl

    '--------------------------------
    'ちらつき防止ダブルバッファリング
    Private WS_EX_COMPOSITED As Integer = &H2000000
    Protected Overrides ReadOnly Property CreateParams() As System.Windows.Forms.CreateParams
        Get
            Dim c As CreateParams
            c = MyBase.CreateParams
            c.ExStyle = c.ExStyle Or Me.WS_EX_COMPOSITED
            Return c
        End Get
    End Property
    '--------------------------------

    Public Event DocumentCompleted As WebBrowserDocumentCompletedEventHandler
    Public Event ProgressChanged As WebBrowserProgressChangedEventHandler
    Public Event StatusTextChanged As EventHandler

    Sub New()
        MyBase.New()
        AddNewWebTabPage()
        Me.DisplayStyle = TabStyle.Chrome
        'TM_EnableDoubleBuffering(Me)        'ダブルバッファをTrueにする（ちらつき防止）
    End Sub

    Private Sub WebTabPage_NewWindow3(ByVal sender As Object, ByVal e As WebBrowserNewWindow3EventArgs)
        Dim WebTabPage As WebTabPage = AddNewWebTabPage()
        e.PpDisp = WebTabPage.WebBrowser.Application
        WebTabPage.WebBrowser.RegisterAsBrowser = True
        'TM_EnableDoubleBuffering(WebTabPage)        'ダブルバッファをTrueにする（ちらつき防止）
    End Sub

    Private Sub WebTabPage_WindowClosing(ByVal sender As Object, ByVal e As EventArgs)
        If Me.SelectedIndex > 0 Then 'タブページが残り１つの場合は閉じない
            WebTabPageClose()
        Else
            Exit Sub
        End If
    End Sub

    Private Sub WebTabPage_TabClosing(ByVal sender As Object, ByVal e As TabControlCancelEventArgs) Handles MyBase.TabClosing
        e.Cancel = True

        If Me.SelectedIndex > 0 Then 'タブページが残り１つの場合は閉じない
            WebTabPageClose()
        Else
            e.Cancel = True
        End If
    End Sub

    Private Sub WebTabPageClose()
        'Dim SelectedTab As TabPage = Me.SelectedTab
        'Me.SelectedIndex -= 1
        'Me.Controls.Remove(SelectedTab)
        'SelectedTab.Dispose()
        'Me.SelectTab(Me.TabCount - 1)

        'Dim SelectedTab As TabPage = Me.SelectedTab
        'If SelectedTab.Text = "伝票番号: 受注番号" Then
        '    Dim dialogResult As String = MsgBox("閉じずに再表示しますか？" & vbCrLf & "新しい伝票を作成する時は「はい」", MsgBoxStyle.YesNoCancel)
        '    If dialogResult = MsgBoxResult.Yes Then
        '        Me.SelectedTab.WebBrowser.Navigate(Me.SelectedTab.WebBrowser.Url)
        '        'Me.SelectedTab.WebBrowser.Refresh()
        '        Exit Sub
        '    ElseIf dialogResult = MsgBoxResult.Cancel Then
        '        Exit Sub
        '    End If
        'End If

        'Me.SelectedTab.Text = ""
        'Me.SelectedTab.Visible = False
        'Me.SelectedTab.WebBrowser.Visible = False
        'Me.SelectedTab.Dispose()

        'Me.SelectedIndex -= 1
        'SelectedTab.Controls.Remove(Me.SelectedTab.WebBrowser)
        'SelectedTab.Controls.Remove(Me.SelectedTab._WebBrowser)
        'Me.Controls.Remove(SelectedTab)

        If Regex.IsMatch(Me.SelectedTab.Text, "^伝票番号:.*") Then
            Dim rmWebTab As TabPage = Me.SelectedTab
            Me.SelectedTab._WebBrowser.Dispose()
            rmWebTab.Dispose()
            Me.SelectedTab._WebBrowser.Refresh()
        Else
            Me.SelectedTab.Dispose()
        End If


        'Dim w As WebBrowser = SelectedTab._WebBrowser
        'TabPages.Remove(SelectedTab)
        'w.Dispose()

        'SelectedTab.Hide()
        'TabPages.RemoveAt(TabPages.Count - 1)

        'Me.SelectedTab.Dispose()
        'Me.SelectedTab.Hide()
        Me.SelectTab(Me.TabCount - 1)
    End Sub

    Private Sub WebTabPage_DocumentCompleted(ByVal sender As Object, ByVal e As WebBrowserDocumentCompletedEventArgs)
        RaiseEvent DocumentCompleted(sender, e)
    End Sub

    Private Sub WebTabPage_StatusTextChanged(ByVal sender As Object, ByVal e As EventArgs)
        RaiseEvent StatusTextChanged(sender, e)
    End Sub

    Private Sub WebTabPage_ProgressChanged(ByVal sender As Object, ByVal e As WebBrowserProgressChangedEventArgs)
        RaiseEvent ProgressChanged(sender, e)
    End Sub

    Dim wbtArray As ArrayList = New ArrayList
    Public Function AddNewWebTabPage() As WebTabPage
        Dim WebTabPage As New WebTabPage
        AddHandler WebTabPage.NewWindow3, AddressOf WebTabPage_NewWindow3
        AddHandler WebTabPage.DocumentCompleted, AddressOf WebTabPage_DocumentCompleted
        AddHandler WebTabPage.WindowClosing, AddressOf WebTabPage_WindowClosing
        AddHandler WebTabPage.StatusTextChanged, AddressOf WebTabPage_StatusTextChanged
        AddHandler WebTabPage.ProgressChanged, AddressOf WebTabPage_ProgressChanged
        Me.Controls.Add(WebTabPage)
        WebTabPage.Tag = WebTabPage
        Me.SelectedTab = WebTabPage
        wbtArray.Add(WebTabPage._WebBrowser)
        Return WebTabPage
    End Function

    Shadows Property SelectedTab() As WebTabPage
        Get
            Return DirectCast(MyBase.SelectedTab, WebTabPage)
        End Get
        Set(ByVal value As WebTabPage)
            MyBase.SelectedTab = DirectCast(value, TabPage)
        End Set
    End Property

End Class

'*************************************************************************************
Public Class WebTabPage
    Inherits TabPage

    Public WithEvents _WebBrowser As New ExWebBrowser2
    Private ReadOnly ToolStrip As New ToolStrip
    Private ReadOnly lblAddress As New ToolStripLabel
    Private WithEvents BtnClose As New ToolStripButton

    Public Event DocumentCompleted As WebBrowserDocumentCompletedEventHandler
    Public Event NewWindow3 As WebBrowserNewWindow3EventHandler
    Public Event ProgressChanged As WebBrowserProgressChangedEventHandler
    Public Event StatusTextChanged As EventHandler
    Public Event WindowClosing As EventHandler

    Overloads Property Text() As String 'タブブラウザーのタイトル表示（10文字で省略する）
        Get
            Return MyBase.Text
        End Get
        Set(ByVal value As String)
            If IsNothing(value) Then
                MyBase.Text = String.Empty
            Else
                MyBase.ToolTipText = value
                If value.Length > 10 Then
                    MyBase.Text = String.Format(value, "{0,-10}").Substring(0, 10)
                Else
                    MyBase.Text = value
                End If
            End If
        End Set
    End Property

    ReadOnly Property WebBrowser() As ExWebBrowser2
        Get
            Return Me._WebBrowser
        End Get
    End Property

    <DllImport("urlmon.dll", CharSet:=CharSet.Ansi)>
    Private Shared Function UrlMkSetSessionOption(ByVal dwOption As Integer, ByVal str As String, ByVal nLength As Integer, ByVal dwReserved As Integer) As Integer
    End Function
    Sub New()
        MyBase.New()
        Me.SuspendLayout()

        'Toolstrip1
        Me.Controls.Add(ToolStrip)
        With ToolStrip
            .BackColor = Color.White
            .GripStyle = ToolStripGripStyle.Hidden
            .RenderMode = ToolStripRenderMode.System
        End With
        'lblTitle
        Me.lblAddress.Overflow = ToolStripItemOverflow.Never
        Me.ToolStrip.Items.Add(lblAddress)
        'btnClose
        Me.ToolStrip.Items.Add(BtnClose)
        With BtnClose
            .Alignment = ToolStripItemAlignment.Right
            .Overflow = ToolStripItemOverflow.Never
            .Font = New Font("Marlett", 9)
            .Text = "r"
        End With
        'WebBrowser1
        Me.Controls.Add(_WebBrowser)
        _WebBrowser.Dock = DockStyle.Fill
        _WebBrowser.BringToFront()

        'userAgent
        'Const URLMON_OPTION_USERAGENT As Integer = &H10000001
        'ここに設定したいUser-Agentをセットする。 
        'Dim UserAgent As String = "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.86 Safari/537.36"
        'UrlMkSetSessionOption(URLMON_OPTION_USERAGENT, UserAgent, UserAgent.Length, 0)

        'If Form1.CheckBox9.Checked = True Then
        _WebBrowser.ScriptErrorsSuppressed = True
        'End If

        Me.ResumeLayout()
    End Sub

    Public Shared Sub WebBrowser_Navigating(url As String)
        'userAgent
        Const URLMON_OPTION_USERAGENT As Integer = &H10000001
        Dim UserAgent As String = ""

        If Form1.自動ToolStripMenuItem.Checked Then
            'If InStr(url, "store.yahoo.co.jp") > 0 Then
            '    UserAgent = "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36"
            'Else
            UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64; Trident/7.0; rv:11.0) like Gecko"
            'End If
        Else
            If Form1.IEToolStripMenuItem.Checked Then
                UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64; Trident/7.0; rv:11.0) like Gecko"
            ElseIf Form1.ChromeToolStripMenuItem.Checked Then
                UserAgent = "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36"
            End If
        End If

        UrlMkSetSessionOption(URLMON_OPTION_USERAGENT, UserAgent, UserAgent.Length, 0)
    End Sub

    Private Sub WebBrowser_Navigated(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserNavigatedEventArgs) Handles _WebBrowser.Navigated
        Me.lblAddress.Text = e.Url.ToString 'webbrowser上のステータス文字
        Form1.ToolStripLabel2.Text = 0
    End Sub

    Private Sub BtnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles BtnClose.Click
        RaiseEvent WindowClosing(sender, e)
    End Sub

    'Private Sub WebBrowser_DocumentTitleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles _WebBrowser.DocumentTitleChanged
    Private Sub WebBrowser_DocumentCompleted(ByVal sender As WebBrowser, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles _WebBrowser.DocumentCompleted
        If sender.Url.ToString <> e.Url.ToString Then
            Exit Sub
        End If

        Me.Text = CType(sender, WebBrowser).DocumentTitle
        Form1.ToolStripTextBox1.Text = CType(sender, WebBrowser).Url.ToString
        Form1.ToolStripLabel2.Text += 1

        If Form1.CheckBox10.Checked = True Then
            Dim s(Form1.DataGridView3.RowCount - 2) As String
            For r As Integer = 0 To Form1.DataGridView3.RowCount - 2
                s(r) = Form1.DataGridView3.Item(0, r).Value
            Next
            Form1.Find(sender, s, 0)
        End If

        Try
            'ログイン情報取得
            Dim q = From Tag As HtmlElement In _WebBrowser.Document.Body.GetElementsByTagName("INPUT")
                    Select Tag, Type = Tag.GetAttribute("type")
                    Where Type = "button" OrElse Type = "submit"
                    Select Tag

            For Each button In q
                button.AttachEventHandler("onclick", Function() LoginSave())
            Next

            q = From Tag As HtmlElement In _WebBrowser.Document.Body.GetElementsByTagName("BUTTON")
            For Each button In q
                button.AttachEventHandler("onclick", Function() LoginSave())
            Next
        Catch ex As Exception
            'MsgBox(ex.Message)
        End Try

        LoginLoad()

        '楽天ログインのみ補正
        'If InStr(_WebBrowser.Url.ToString, "sp_id") > 0 Then
        '    loginRakutenNum = loginRakutenNum * 2 - 1
        'End If
        Dim loginRakuten As String = Form1_F_loginlist.ComboBox1.Items(loginRakutenNum)
        Dim loginYahoo As String = Form1_F_loginlist.ComboBox2.Items(loginYahooNum)

        'loginlistの番号
        Dim tenpoColNum As Integer = 1
        Dim idColNum As Integer = 2
        Dim idClassColNum As Integer = 6
        Dim passColNum As Integer = 3
        Dim passClassColNum As Integer = 7
        Dim urlColNum As Integer = 5

        'ログインフォーム書き込み
        Dim num As Integer = 0  'ログイン店舗
        For i As Integer = 0 To loginArray.Count - 1
            Dim lArray As String() = Split(loginArray(i), ",")
            If lArray.Length >= 7 Then
                Form1.ToolStripLabel1.Text = InStr(_WebBrowser.Url.ToString, lArray(urlColNum))
                If InStr(_WebBrowser.Url.ToString, lArray(urlColNum)) > 0 Then
                    Try
                        num += 1

                        If InStr(_WebBrowser.Url.ToString, "yahoo") > 0 Then
                            If InStr(lArray(tenpoColNum), loginYahoo) = 0 Then
                                Continue For
                            End If
                        ElseIf InStr(_WebBrowser.Url.ToString, "rakuten") > 0 Then
                            If InStr(lArray(tenpoColNum), loginRakuten) = 0 Then
                                Continue For
                            End If
                        End If

                        If _WebBrowser.Document.GetElementsByTagName("input")(lArray(idClassColNum)) <> Nothing Then
                            If _WebBrowser.Document.GetElementsByTagName("input")(lArray(idClassColNum)).Enabled Then
                                _WebBrowser.Document.GetElementsByTagName("input")(lArray(idClassColNum)).SetAttribute("value", lArray(idColNum))
                            End If
                        End If

                        If _WebBrowser.Document.GetElementsByTagName("input")(lArray(passClassColNum)) <> Nothing Then
                            If _WebBrowser.Document.GetElementsByTagName("input")(lArray(passClassColNum)).Enabled Then
                                _WebBrowser.Document.GetElementsByTagName("input")(lArray(passClassColNum)).SetAttribute("value", lArray(passColNum))
                                Exit For
                            End If
                        End If
                    Catch ex As Exception

                    End Try
                End If
            End If
        Next
    End Sub

    Dim base64 As New MyBase64str("UTF-8")
    Dim ENC_UTF8 As Encoding = Encoding.UTF8
    Dim ENC_SJ As Encoding = Encoding.GetEncoding("SHIFT-JIS")
    Private Sub LoginLoad()
        Dim svFile As String = Form1.サーバーToolStripMenuItem.Text & "\update\loginlist4.dat"
        Dim llFile As String = Path.GetDirectoryName(Form1.appPath) & "\loginlist4.dat"
        If Not File.Exists(llFile) Then
            File.Copy(svFile, llFile)
        End If

        If File.Exists(llFile) Then
            loginArray.Clear()
            Dim llArray As String() = Split(base64.Decode(File.ReadAllText(llFile, ENC_UTF8)), vbCrLf)
            loginArray.AddRange(llArray)

            For i As Integer = 0 To loginArray.Count - 1
                If loginArray(i) <> "" Then
                    If Regex.IsMatch(loginArray(i), "^楽天=") Then
                        loginRakutenNum = Split(loginArray(i), "=")(1)
                    ElseIf Regex.IsMatch(loginArray(i), "^Yahoo=") Then
                        loginYahooNum = Split(loginArray(i), "=")(1)
                    End If
                End If
            Next
        End If

        'Dim fArray As String() = File.ReadAllLines(LoadFile, ENC_SJ)
        'loginArray.Clear()
        'Dim LoadFile As String = Path.GetDirectoryName(Form1.appPath) & "\login2.txt"
        'If File.Exists(LoadFile) Then
        '    Dim fArray As String() = File.ReadAllLines(LoadFile, ENC_SJ)
        '    For i As Integer = 0 To fArray.Length - 1
        '        If InStr(fArray(i), "|") > 0 Then
        '            loginArray.Add(fArray(i))
        '        Else
        '            If InStr(fArray(i), ",") > 0 Then
        '                Dim numArray As String() = Split(fArray(i), ",")
        '                loginRakutenNum = numArray(0)
        '                loginYahooNum = numArray(1)
        '            End If
        '        End If
        '    Next

        '    'データ更新
        '    Dim updateFlag As Boolean = False
        '    For i As Integer = 0 To loginArray.Count - 1
        '        Dim lineArray As String() = Split(loginArray(i), "|")
        '        For k As Integer = 0 To llArray.Length - 1
        '            Dim kArray As String() = Split(llArray(k), ",")
        '            If kArray.Length > 3 And InStr(llArray(k), "●") > 0 Then
        '                If lineArray(4) = kArray(1) Then
        '                    If kArray(2) <> "" Then
        '                        Dim lA As String() = Split(lineArray(1), ",")
        '                        If lA(1) <> "" And lA(1) <> kArray(2) Then
        '                            loginArray(i) = lineArray(0) & "|" & lA(0) & "," & kArray(2) & "|" & lineArray(2) & "|" & lineArray(3) & "|" & lineArray(4)
        '                            updateFlag = True
        '                        End If
        '                    End If
        '                    If kArray(3) <> "" Then
        '                        Dim lA As String() = Split(lineArray(2), ",")
        '                        If lA(1) <> "" And lA(1) <> kArray(3) Then
        '                            loginArray(i) = lineArray(0) & "|" & lineArray(1) & "|" & lA(0) & "," & kArray(3) & "|" & lineArray(3) & "|" & lineArray(4)
        '                            updateFlag = True
        '                        End If
        '                    End If
        '                    Exit For
        '                End If
        '            End If
        '        Next
        '    Next

        '    If updateFlag Then
        '        Dim wStr As String = ""
        '        For Each line As String In loginArray
        '            wStr &= line & vbCrLf
        '        Next
        '        File.WriteAllText(Path.GetDirectoryName(Form1.appPath) & "\login2.txt", wStr, ENC_SJ)
        '    End If
        'End If

        loginArray2.Clear()
        Dim LoadFile2 As String = Path.GetDirectoryName(Form1.appPath) & "\login3.txt"
        If File.Exists(LoadFile2) Then
            Dim fArray As String() = File.ReadAllLines(LoadFile2, ENC_SJ)
            loginArray2.AddRange(fArray)
        End If
    End Sub

    Dim submitFlag As Integer = 0
    Dim loginArray As New ArrayList
    Dim loginRakutenNum As Integer = 1
    Dim loginYahooNum As Integer = 1
    Dim loginArray2 As New ArrayList
    Dim loginArrayNum As Integer = 0
    Private Function LoginSave()
        'ログイン登録済が調べる
        Dim checkA As String = ""
        For i As Integer = 0 To loginArray.Count - 1
            Dim lArray As String() = Split(loginArray(i), "|")
            If InStr(_WebBrowser.Url.ToString, lArray(0)) > 0 Then
                checkA = loginArray(i)
                loginArrayNum = i
                submitFlag = 1
            End If
        Next

        'フォームの取得用キーワード
        'Dim id As String() = New String() {"name", "user_name", "username", "id", "rlogin-username-ja", "rlogin-username-2-ja"}
        'Dim pass As String() = New String() {"pass", "password", "passwd", "psd", "rlogin-password-ja", "rlogin-password-2-ja"}
        Dim id As String() = Nothing
        Dim pass As String() = Nothing
        Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\config\loginform.txt"
        Dim q As Integer = 0
        Using sr As New System.IO.StreamReader(fName, System.Text.Encoding.Default)
            Do While Not sr.EndOfStream
                Dim s As String = sr.ReadLine
                If q = 0 Then
                    id = Split(s, ",")
                ElseIf q = 1 Then
                    pass = Split(s, ",")
                Else
                    Exit Do
                End If
                q += 1
            Loop
        End Using

        'フォームのvalue取得
        Dim idNew As String = ""
        Dim passNew As String = ""
        For i As Integer = 0 To id.Length - 1
            Dim res As String = ""
            Try
                If _WebBrowser.Document.GetElementsByTagName("INPUT")(id(i)) <> Nothing Then
                    res = _WebBrowser.Document.GetElementsByTagName("INPUT")(id(i)).GetAttribute("value")
                    If res <> "" Then
                        idNew = id(i) & "," & res
                        Exit For
                    End If
                End If
            Catch ex As Exception

            End Try
        Next
        For i As Integer = 0 To pass.Length - 1
            Dim res As String = ""
            Try
                If _WebBrowser.Document.GetElementsByTagName("INPUT")(pass(i)) <> Nothing Then
                    res = _WebBrowser.Document.GetElementsByTagName("INPUT")(pass(i)).GetAttribute("value")
                    If res <> "" Then
                        passNew = pass(i) & "," & res
                        Exit For
                    End If
                End If
            Catch ex As Exception

            End Try
        Next

        If idNew = "" And passNew = "" Then
            Return False
            Exit Function
        End If

        'URLの切り取り
        Dim u As New Uri(Form1.ToolStripTextBox1.Text)
        Dim uPath As String = u.GetLeftPart(UriPartial.Path)
        Dim uQueryStr As String = ""
        If u.Query = "" Then
            uQueryStr = ""
        Else
            If u.Query.Length > 20 Then
                uQueryStr = u.Query.Substring(0, 20)
            Else
                uQueryStr = u.Query
            End If
        End If
        Dim uStr As String = uPath & uQueryStr

        If submitFlag = 0 Then  '新規
            '情報の保存
            If idNew <> "" Then
                Dim DR As DialogResult = MsgBox("ログイン情報を保存しますか？", MsgBoxStyle.OkCancel Or MsgBoxStyle.SystemModal)
                If DR = Windows.Forms.DialogResult.OK Then
                    loginArray.Add(uStr & "|" & idNew & "|" & passNew)
                    Dim str As String = ""
                    For Each res As String In loginArray
                        str &= res & vbCrLf
                    Next
                    Dim dirLoadFile As String = Path.GetDirectoryName(Form1.appPath) & "\login3.txt"
                    My.Computer.FileSystem.WriteAllText(dirLoadFile, str, False, Form1.enc)
                End If
            End If
        Else    '更新
            'Dim checkB As String = uStr & "|" & idNew & "|" & passNew
            'If checkA <> checkB Then    '前に登録したのと同じだったら何もしない
            '    Dim DR As DialogResult = MsgBox("ログイン情報を更新しますか？", MsgBoxStyle.OkCancel Or MsgBoxStyle.SystemModal)
            '    If DR = Windows.Forms.DialogResult.OK Then
            '        Dim cA As String() = Split(checkA, "|")
            '        If idNew = "" Then
            '            idNew = cA(1)
            '        End If
            '        If passNew = "" Then
            '            passNew = cA(2)
            '        End If
            '        loginArray(loginArrayNum) = uStr & "|" & idNew & "|" & passNew
            '        Dim str As String = ""
            '        For Each resB As String In loginArray
            '            str &= resB & vbCrLf
            '        Next
            '        Dim dirLoadFile As String = Path.GetDirectoryName(Form1.appPath) & "\login.txt"
            '        My.Computer.FileSystem.WriteAllText(dirLoadFile, str, False, Form1.enc)
            '    End If
            'End If
        End If

        Return True
    End Function

    'Private Sub WebBrowser_DocumentCompleted(ByVal sender As Object, ByVal e As WebBrowserDocumentCompletedEventArgs) Handles _WebBrowser.DocumentCompleted
    '    RaiseEvent DocumentCompleted(sender, e)
    'End Sub

    'Private Sub _WebBrowser_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles _WebBrowser.GotFocus
    '    'ページ読込の最後を感知
    '    'If WebBrowser.Url = DirectCast(sender, WebBrowser).Url Then
    '    'Form1.ToolStripTextBox1.Text = _WebBrowser.Url.ToString

    '    If Form1.CheckBox10.Checked = True Then
    '        Dim s(Form1.DataGridView3.RowCount - 2) As String
    '        For r As Integer = 0 To Form1.DataGridView3.RowCount - 2
    '            s(r) = Form1.DataGridView3.Item(0, r).Value
    '        Next
    '        Form1.Find(sender, s, 0)
    '    End If
    '    'End If
    'End Sub

    Private Sub WebBrowser_StatusTextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles _WebBrowser.StatusTextChanged
        RaiseEvent StatusTextChanged(sender, e)
    End Sub

    Private Sub WebBrowser_ProgressChanged(ByVal sender As Object, ByVal e As WebBrowserProgressChangedEventArgs) Handles _WebBrowser.ProgressChanged
        RaiseEvent ProgressChanged(sender, e)
    End Sub

    Private Sub WebBrowser_NewWindow3(ByVal sender As Object, ByVal e As WebBrowserNewWindow3EventArgs) Handles _WebBrowser.NewWindow3
        RaiseEvent NewWindow3(sender, e)
    End Sub

    Private Sub WebBrowser_WindowClosing(ByVal sender As Object, ByVal e As EventArgs) Handles _WebBrowser.WindowClosing
        RaiseEvent WindowClosing(sender, e)
    End Sub


End Class

'*************************************************************************************
Public Class ExWebBrowser2
    Inherits ExWebBrowser

    Sub New()
        MyBase.New()
    End Sub

    'WindowClosingイベントの拡張
    Enum GETWINDOWCMD
        GW_HWNDFIRST = 0
        GW_HWNDLAST = 1
        GW_HWNDNEXT = 2
        GW_HWNDPREV = 3
        GW_OWNER = 4
        GW_CHILD = 5
        GW_ENABLEDPOPUP = 6
    End Enum

    <DllImport("user32.dll")> Private Shared Function GetWindow(ByVal hWnd As IntPtr, ByVal uCmd As GETWINDOWCMD) As IntPtr
    End Function

    Public Event WindowClosing As EventHandler


    Protected Overridable Sub OnWindowClosing(ByVal e As EventArgs)
        RaiseEvent WindowClosing(Me, e)
    End Sub

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        Const WM_PARENTNOTIFY = &H210
        Const WM_DESTROY = &H2
        If m.Msg = WM_PARENTNOTIFY Then
            If m.WParam.ToInt32 = WM_DESTROY Then
                If m.LParam = GetWindow(Me.Handle, GETWINDOWCMD.GW_CHILD) Then
                    Dim e As New EventArgs
                    OnWindowClosing(e)
                    Return
                End If
            End If
        End If
        MyBase.WndProc(m)
    End Sub

    'NewWindow2イベントの拡張
    Private cookie As AxHost.ConnectionPointCookie
    Private helper As WebBrowser2EventHelper

    Public Event NewWindow3 As WebBrowserNewWindow3EventHandler

    <System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)>
    <System.Runtime.InteropServices.DispIdAttribute(200)>
    Public ReadOnly Property Application() As Object
        Get
            If IsNothing(Me.ActiveXInstance) Then
                Throw New AxHost.InvalidActiveXStateException("Application", AxHost.ActiveXInvokeKind.PropertyGet)
            End If
            Return Me.ActiveXInstance.Application
        End Get
    End Property

    <System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)>
    <System.Runtime.InteropServices.DispIdAttribute(552)>
    Public Property RegisterAsBrowser() As Boolean
        Get
            If IsNothing(Me.ActiveXInstance) Then
                Throw New AxHost.InvalidActiveXStateException("RegisterAsBrowser", AxHost.ActiveXInvokeKind.PropertyGet)
            End If
            Return Me.ActiveXInstance.RegisterAsBrowser
        End Get
        Set(ByVal value As Boolean)
            If IsNothing(Me.ActiveXInstance) Then
                Throw New AxHost.InvalidActiveXStateException("RegisterAsBrowser", AxHost.ActiveXInvokeKind.PropertySet)
            End If
            Me.ActiveXInstance.RegisterAsBrowser = value
        End Set
    End Property

    <PermissionSetAttribute(SecurityAction.LinkDemand, Name:="FullTrust")>
    Protected Overrides Sub CreateSink()
        MyBase.CreateSink()
        helper = New WebBrowser2EventHelper(Me)
        cookie = New AxHost.ConnectionPointCookie(Me.ActiveXInstance, helper, GetType(IDWebBrowserEvents2))
    End Sub

    <PermissionSetAttribute(SecurityAction.LinkDemand, Name:="FullTrust")>
    Protected Overrides Sub DetachSink()
        If cookie IsNot Nothing Then
            cookie.Disconnect()
            cookie = Nothing
        End If
        MyBase.DetachSink()
    End Sub

    Protected Overridable Sub OnNewWindow3(ByVal e As WebBrowserNewWindow3EventArgs)
        RaiseEvent NewWindow3(Me, e)
    End Sub

    Private Class WebBrowser2EventHelper
        Inherits StandardOleMarshalObject
        Implements IDWebBrowserEvents2

        Private parent As ExWebBrowser2

        Public Sub New(ByVal obj As ExWebBrowser2)
            parent = obj
        End Sub

        '*	@param	[i/o]	ppDisp			現在のウインドウのナビゲーションに使用するWebBrowserオブジェクト(?).
        '*	@param	[i/o]	cancel			True = デフォルトのブラウザ(IE)で新しいページを開く.
        '*	@param	[in]	dwFlags			NWMF値(新しいウインドウのポップアップ・ウインドウに関する値).
        '*	@param	[in]	bstrUrlContext	現在のウインドウに表示されているページのURL.
        '*	@param	[in]	bstrUrl			新しいウインドウに表示するページのURL.
        Public Sub NewWindow3(ByRef ppDisp As Object, ByRef cancel As Boolean, ByVal dwFlags As UInt32,
                            ByVal bstrUrlContext As String, ByVal bstrUrl As String) Implements IDWebBrowserEvents2.NewWindow3
            Dim e As New WebBrowserNewWindow3EventArgs(ppDisp, bstrUrlContext, bstrUrl)
            parent.OnNewWindow3(e)
            ppDisp = e.PpDisp
            cancel = e.Cancel
        End Sub
    End Class

End Class

'*************************************************************************************
'NewWindow2イベント
Public Delegate Sub WebBrowserNewWindow3EventHandler(ByVal sender As Object, ByVal e As WebBrowserNewWindow3EventArgs)

'*************************************************************************************
Public Class WebBrowserNewWindow3EventArgs
    Inherits System.ComponentModel.CancelEventArgs

    Private ppDispValue As Object
    Private bstrUrlContextValue As String
    Private bstrUrlValue As String

    Public Sub New(ByRef ppDisp As Object, ByVal bstrUrlContext As String, ByVal bstrUrl As String)
        ppDispValue = ppDisp
        bstrUrlContextValue = bstrUrlContext
        bstrUrlValue = bstrUrl
    End Sub

    Public Property PpDisp() As Object
        Get
            Return ppDispValue
        End Get
        Set(ByVal value As Object)
            ppDispValue = value
        End Set
    End Property

    Public Property BstrUrlContext() As String
        Get
            Return bstrUrlContextValue
        End Get
        Set(ByVal value As String)
            bstrUrlContextValue = value
        End Set
    End Property

    Public Property BstrUrl() As String
        Get
            Return bstrUrlValue
        End Get
        Set(ByVal value As String)
            bstrUrlValue = value
        End Set
    End Property

End Class

'*************************************************************************************
<ComImport(), Guid("34A715A0-6587-11D0-924A-0020AFC7AC4D"),
InterfaceType(ComInterfaceType.InterfaceIsIDispatch),
TypeLibType(TypeLibTypeFlags.FHidden)>
Public Interface IDWebBrowserEvents2
    Enum DISPID
        NEWWINDOW3 = 273    '251
    End Enum

    <DispId(DISPID.NEWWINDOW3)> Sub NewWindow3(
      <InAttribute(), OutAttribute(), MarshalAs(UnmanagedType.IDispatch)> ByRef pDisp As Object,
      <InAttribute(), OutAttribute()> ByRef cancel As Boolean,
      <InAttribute()> ByVal dwFlags As UInt32,
      <InAttribute()> ByVal bstrUrlContext As String,
      <InAttribute()> ByVal bstrUrl As String)

End Interface

'*************************************************************************************
Public Class ExWebBrowser
    Inherits WebBrowser

    Sub New()
        MyBase.New()
    End Sub

    'キーボードショートカットの処理を再定義
    Public Overrides Function PreProcessMessage(ByRef msg As System.Windows.Forms.Message) As Boolean
        Const WM_KEYDOWN As Integer = &H100
        If msg.Msg = WM_KEYDOWN Then
            Dim keyCode As Keys = CType(msg.WParam, Keys) And Keys.KeyCode
            If My.Computer.Keyboard.CtrlKeyDown Then
                Select Case keyCode
                    Case Keys.C
                        If Form1.CheckBox1.Checked = True Then
                            Me.Document.ExecCommand("Copy", False, Nothing)
                            Form1.DataGridView4.Rows.Add(1)
                            Form1.DataGridView4.Item(0, Form1.DataGridView4.RowCount - 2).Value = Clipboard.GetText()
                            Dim r As Integer = Form1.DataGridView4.RowCount - 1
                            Form1.DataGridView4.ClearSelection()
                            Form1.DataGridView4.Item(0, r).Selected = True
                            Form1.DataGridView4.CurrentCell = Form1.DataGridView4(0, r)
                        Else
                            Me.Document.ExecCommand("Copy", False, Nothing)
                        End If
                        Return True
                    Case Keys.V
                        Me.Document.ExecCommand("Paste", False, Nothing)
                        Return True
                    Case Keys.N
                        If Me.ReadyState >= WebBrowserReadyState.Loading Then
                            WebTabPage.WebBrowser_Navigating(Me.Url.ToString)
                            Me.Navigate(Me.Url, True)
                        End If
                        Return True
                    Case Keys.P
                        Me.ShowPrintPreviewDialog()
                        Return True
                    Case Keys.W
                        Me.Dispose()
                        Return True
                End Select
            ElseIf My.Computer.Keyboard.ShiftKeyDown Then
                If Form1.CheckBox1.Checked = True Then
                    Me.Document.ExecCommand("Copy", False, Nothing)
                    Form1.DataGridView4.Rows.Add(1)
                    Form1.DataGridView4.Item(0, Form1.DataGridView4.RowCount - 2).Value = Clipboard.GetText()
                    Dim r As Integer = Form1.DataGridView4.RowCount - 1
                    Form1.DataGridView4.ClearSelection()
                    Form1.DataGridView4.Item(0, r).Selected = True
                    Form1.DataGridView4.CurrentCell = Form1.DataGridView4(0, r)
                    Return True
                ElseIf Form1.CheckBox2.Checked = True Then
                    If Form1.DataGridView4.SelectedCells(0).Value <> "" Then    'nullは貼り付けられないので回避
                        Clipboard.SetText(Form1.DataGridView4.SelectedCells(0).Value)
                        Me.Document.ExecCommand("Paste", False, vbNull)
                    End If
                    Dim r As Integer = Form1.DataGridView4.SelectedCells(0).RowIndex + 1
                    If r > Form1.DataGridView4.RowCount - 2 Then
                        Form1.DataGridView4.ClearSelection()
                        Form1.DataGridView4.Item(0, 0).Selected = True
                        Form1.DataGridView4.CurrentCell = Form1.DataGridView4(0, 0)
                        Return True
                    Else
                        Form1.DataGridView4.ClearSelection()
                        Form1.DataGridView4.Item(0, r).Selected = True
                        Form1.DataGridView4.CurrentCell = Form1.DataGridView4(0, r)
                        Return True
                    End If
                End If
            End If
        End If
        Return MyBase.PreProcessMessage(msg)
    End Function

    'WM_APPCOMMANDメッセージに対応
    'WinUser.h
    Enum APPCOMMAND
        BROWSER_BACKWARD = 1
        BROWSER_FORWARD = 2
        BROWSER_REFRESH = 3
        BROWSER_STOP = 4
        BROWSER_SEARCH = 5
        BROWSER_FAVORITES = 6
        BROWSER_HOME = 7
    End Enum

    'WinUser.h
    '#define GET_APPCOMMAND_LPARAM(lParam) ((short)(HIWORD(lParam) & ~FAPPCOMMAND_MASK))
    Private Function GET_APPCOMMAND_LPARAM(ByVal lParam As IntPtr) As Short
        Const FAPPCOMMAND_MASK As UInt16 = &HF000
        Return CShort(((CType(lParam, Integer) And &HFFFF0000) >> 16) And (Not FAPPCOMMAND_MASK))
    End Function

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        Const WM_APPCOMMAND = &H319
        If m.Msg = WM_APPCOMMAND Then
            Select Case GET_APPCOMMAND_LPARAM(m.LParam)
                Case APPCOMMAND.BROWSER_BACKWARD
                    Me.GoBack()
                    Return
                Case APPCOMMAND.BROWSER_FORWARD
                    Me.GoForward()
                    Return
                Case APPCOMMAND.BROWSER_REFRESH
                    Me.Refresh()
                    Return
                Case APPCOMMAND.BROWSER_STOP
                    Me.Stop()
                    Return
                Case APPCOMMAND.BROWSER_SEARCH
                    Me.GoSearch()
                    Return
                Case APPCOMMAND.BROWSER_HOME
                    Me.GoHome()
                    Return
            End Select
        End If
        MyBase.WndProc(m)
    End Sub


End Class





