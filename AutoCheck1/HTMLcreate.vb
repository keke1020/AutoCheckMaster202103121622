Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security.Permissions
Imports System.Text.RegularExpressions
Imports Sgry.Azuki
Imports Sgry.Azuki.Highlighter
Imports EncodingOperation
Imports SHDocVw
Imports HtmlAgilityPack
Imports ZetaHtmlTidy
Imports Sgry.Azuki.WinForms

Public Class HTMLcreate
    ' ole32.dllのOleDrawを使用する(Const定義)
    Const DVASPECT_CONTENT As Integer = 1

    ' ole32.dllのOleDrawを使用する(DllImport定義)
    <System.Runtime.InteropServices.DllImport("ole32.dll")>
    Public Shared Function OleDraw(
        ByVal pUnk As IntPtr,
        ByVal dwAspect As Integer,
        ByVal hdcDraw As IntPtr,
        ByRef lprcBounds As Rectangle) As Integer
    End Function

    Dim x As Integer = 0
    Dim y As Integer = 0
    Dim picSize As Size
    Dim inteliList As New ArrayList
    Public Shared az As WinForms.AzukiControl

    Dim azArray As WinForms.AzukiControl()
    Dim tbArray As ToolStripStatusLabel()

    Private Sub Form6_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ListBox1.Hide()

        AddHandler My.Settings.SettingChanging, AddressOf Settings_SettingChanging

        ' WebBrowserを非表示にする
        'Me.WebBrowser1.Visible = False

        ' WebBrowserのサイズを自由に変更出来るようにする
        'Me.WebBrowser1.Dock = DockStyle.None

        ' WebBrowserのスクロールバーを無効にする
        'Me.WebBrowser1.ScrollBarsEnabled = False


        ' WEBページのURLを指定
        Me.WebBrowser1.Navigate("")

        'ファイル読み取り
        Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\config\intelisense.txt"
        Using sr As New System.IO.StreamReader(fName, System.Text.Encoding.Default)
            Do While Not sr.EndOfStream
                Dim s As String = sr.ReadLine
                If s <> "" Then
                    inteliList.Add(s)
                End If
            Loop
        End Using

        'シンタックスハイライター
        Dim kwdFile As String() = New String() {"base", "html", "html5", "CSSporp", "CSSprop_d", "CSSvalue", "JavaScript"}
        Dim key2 As New ArrayList
        For i As Integer = 0 To kwdFile.Length - 1
            fName = Path.GetDirectoryName(Form1.appPath) & "\config\" & kwdFile(i) & ".kwd"
            Using sr As New System.IO.StreamReader(fName, System.Text.Encoding.Default)
                Do While Not sr.EndOfStream
                    Dim s As String = sr.ReadLine
                    If InStr(s, "//") = 0 And s <> "" Then
                        If key2.Contains(s) = 0 Then
                            key2.Add(Trim(s))
                        End If
                    End If
                Loop
            End Using
        Next
        Dim key3 As String() = Nothing
        key3 = DirectCast(key2.ToArray(GetType(String)), String())
        key2.Clear()
        Array.Sort(key3)

        inteliList.AddRange(key3)

        Dim keys As KeywordHighlighter = New KeywordHighlighter
        'Dim key1 As String() = New String() {"</a", "</br", "</div", "</font", "</head", _
        '    "</strong", "</table", "</tbody", "</td", "</tr", _
        '    "<a", "<br", "<center", "<div", "<font", "<head", "<html", "<img", _
        '    "<strong", "<table", "<tbody", "<td", "<title", "<tr", _
        '    ">", "align=", "alt=", "bgColor=", "border=", "cellSpacing=", "color=", _
        '    "face=", "height=", "href=", "size=", "src=", "style=", "valign=", "width="}
        'Array.Sort(key1)
        keys.AddEnclosure("""", """", CharClass.Comment, True)
        keys.AddKeywordSet(key3, CharClass.Keyword, False)
        keys.AddEnclosure("<!--", "-->", CharClass.Comment, True)

        az = AzukiControl1
        AzukiControl1.Highlighter = keys
        AzukiControl2.Highlighter = keys
    End Sub

    'フォームのサイズ最小化バグ修正
    Private Sub Settings_SettingChanging(ByVal sender As Object, ByVal e As System.Configuration.SettingChangingEventArgs)
        If Not Me.WindowState = FormWindowState.Normal Then
            If e.SettingName = "Form6Size" Or e.SettingName = "Form6Location" Then
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub Form6_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If e.CloseReason = CloseReason.FormOwnerClosing Or e.CloseReason = CloseReason.UserClosing Then
            Me.Visible = False
            e.Cancel = True
        End If
    End Sub

    Private Sub AzukiControl1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles AzukiControl1.DragDrop
        If AzukiControl1.Text <> "" Then
            Dim DR As DialogResult = MsgBox("読み込んでよろしいですか？", MsgBoxStyle.YesNo Or MsgBoxStyle.SystemModal)
            If DR = Windows.Forms.DialogResult.No Then
                AzukiControl1.Text = ""
            End If
        End If

        For Each filename As String In e.Data.GetData(DataFormats.FileDrop)
            Dim e1 As EncodingJP
            Dim e2 As Boolean
            AzukiControl1.Text = TextFileEncoding.LoadFromFile(filename, e1, True, e2)
            ToolStripStatusLabel5.Text = EncodingName(e1)

            'Dim sr As New System.IO.StreamReader(filename)
            'AzukiControl1.Text = sr.ReadToEnd()
            'sr.Close()
            ToolStripStatusLabel1.Text = filename
            Exit For
        Next
    End Sub

    Private Sub AzukiControl1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles AzukiControl1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
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

    Private Sub ブラウザ更新ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ブラウザ更新ToolStripMenuItem.Click
        BrouserUpdate()
    End Sub

    Dim aSorce As String = ""
    Dim bSorce As String = ""
    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        If ToolStripButton1.Text = "自動更新OFF" Then
            Exit Sub
        End If

        If AzukiControl1.TextLength = 0 Then
            Exit Sub
        Else
            If ToolStripStatusLabel1.Text = "..." Then
                If aSorce = "" Then
                    BrouserUpdate()
                    aSorce = AzukiControl1.Text
                Else
                    bSorce = AzukiControl1.Text
                    If aSorce <> bSorce Then
                        BrouserUpdate()
                    End If
                    aSorce = AzukiControl1.Text
                End If
            End If
        End If
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        If ToolStripButton1.Text = "自動更新OFF" Then
            Timer2.Enabled = True
            ToolStripButton1.Text = "自動更新ON"
        Else
            Timer2.Enabled = False
            ToolStripButton1.Text = "自動更新OFF"
        End If
    End Sub

    Private Sub BrouserUpdate()
        TabControl2.SelectedIndex = 0
        scX = WebBrowser1.Document.Body.ScrollLeft
        scY = WebBrowser1.Document.Body.ScrollTop

        Dim flag As Boolean = False
        If ToolStripStatusLabel1.Text <> "..." Then
            If File.Exists(ToolStripStatusLabel1.Text) Then
                WebBrowser1.Navigate(ToolStripStatusLabel1.Text)
                flag = False
            Else
                WebBrowser1.DocumentText = az.Text
            End If
        Else
            WebBrowser1.DocumentText = az.Text
        End If

        'If flag = True Then
        'WebBrowser1.DocumentText = az.Text
        'WebBrowser1.Navigate(New System.Uri("about:blank"))
        'WebBrowser1.Document.DomDocument.write(az.Text)
        'WebBrowser1.DocumentText = az.Text

        'WebBrowser1.Url = New Uri(ToolStripStatusLabel1.Text)
        'WebBrowser1.DocumentText = az.Text
        'Dim fName As String = "G:\楽天サイト新規\index_pc.html"
        ''Dim fName As String = "http://localhost/index_pc.html"
        'Dim u As New Uri(fName)
        'ToolStripTextBox1.Text = u.ToString
        ''WebBrowser1.Navigate(New System.Uri(fName))
        'WebBrowser1.Navigate(u)
        'End If

        Application.DoEvents()

        'IMG_GET()
    End Sub

    Private Sub ブラウザ表示ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ブラウザ表示ToolStripMenuItem.Click
        If ブラウザ表示ToolStripMenuItem.Checked = True Then
            ブラウザ表示ToolStripMenuItem.Checked = False
            SplitContainer1.Panel2Collapsed = True
        Else
            ブラウザ表示ToolStripMenuItem.Checked = True
            SplitContainer1.Panel2Collapsed = False
        End If
    End Sub

    Private Sub HTMLチェッカーToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HTMLチェッカーToolStripMenuItem.Click
        Dim jsStr As String = "javascript:var w=window.open('','_blank','width=800,height=500,scrollbars=yes');var s=document.createElement('script');s.charset='Shift_JIS';s.id='tagcheck-script';s.src='http://tkr-net.tk/tagcheck/tagcheck2.js?'+Math.random();document.body.appendChild(s);void(0);"

        'Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\tagcheck2.js"
        'Dim u As New Uri(fName)
        'fName = Replace(fName, "\", "/")
        'Dim jsStr As String = "javascript:var w=window.open('','_blank','width=800,height=500,scrollbars=yes');var s=document.createElement('script');s.charset='Shift_JIS';s.id='tagcheck-script';s.src='" & u.AbsoluteUri.ToString & "?'+Math.random();document.body.appendChild(s);void(0);"

        'Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\tagcheck2.js"
        'Dim r As Integer = 0
        'Using sr As New System.IO.StreamReader(fName, System.Text.Encoding.Default)
        '    jsStr = sr.ReadToEnd
        'End Using

        'Dim RetVal As Object = Shell(u.AbsoluteUri)
        'WebBrowser1.Url = New Uri(jsStr)
        WebBrowser1.Navigate(jsStr)
    End Sub

    Private Sub 整形のみToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 整形のみToolStripMenuItem.Click
        Seikei()
    End Sub

    Private Sub 行頭SPTAB除去ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 行頭SPTAB除去ToolStripMenuItem.Click
        Sptabdel()
        Seikei()
    End Sub

    Private Sub Sptabdel()
        ToolStripProgressBar1.Maximum = az.LineCount
        ToolStripProgressBar1.Value = 0
        For r As Integer = 0 To az.LineCount - 1
            Dim start As Integer = az.Document.GetLineHeadIndex(r)
            Dim ends As Integer = start + az.Document.GetLineLength(r)
            Dim str As String = az.Document.GetLineContent(r)
            str = LTrim(str)
            str = Replace(str, "  <", "<")
            str = Replace(str, " <", "<")
            str = Replace(str, vbTab, "")
            If Not str Is Nothing Then
                az.Document.Replace(str, start, ends)
            End If
            ToolStripProgressBar1.Value += 1
        Next
    End Sub

    Private Sub Seikei()
        ToolStripProgressBar1.Maximum = az.Document.Length * 5
        ToolStripProgressBar1.Value = 0

        Dim azl As Integer = az.LineCount * 100
        For linenum As Integer = 0 To azl - 1
            If linenum > az.LineCount - 1 Then
                Exit For
            End If
            ToolStripProgressBar1.Maximum = azl + 1

            Dim word As String = ""
            For charanum As Integer = az.GetLineHeadIndex(linenum) To az.GetLineHeadIndex(linenum) + az.GetLineLength(linenum) - 1
                Dim c As String = az.Document.GetCharAt(charanum)
                c = c.ToLower
                word &= c
                'MsgBox("line:" & linenum & " chara:" & c & vbCrLf & "word:" & word & vbCrLf & "maxLines:" & azl, MsgBoxStyle.SystemModal)
                Select Case word
                    Case vbCr, vbLf, vbCrLf
                        Exit For
                    Case Else
                        Select Case True
                            Case Regex.IsMatch(word, "<br>.|<t.>.|</t.*>.|</d.*>.|/a>.|-->.|ter>.|ter"">.|blank>.|0>.|gif'>.|%>.")
                                az.Document.Replace(vbCrLf, charanum, charanum)
                                azl += 1
                                Exit For
                        End Select
                End Select
            Next

            ToolStripProgressBar1.Value += 1
        Next

        ToolStripProgressBar1.Value = 0
    End Sub

    Private Sub Br追加ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Br追加ToolStripMenuItem.Click
        az.Text = Replace(az.Text, vbCrLf, "<br>" & vbCrLf)
    End Sub

    Private Sub Br削除ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Br削除ToolStripMenuItem.Click
        az.Text = Replace(az.Text, "<br>", "")
    End Sub

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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TrackBar1.Value = 50
        Dim MyWeb As Object = WebBrowser1.ActiveXInstance
        MyWeb.ExecWB(63, 1, CType(TrackBar1.Value, Integer), IntPtr.Zero)
        TextBox9.Text = TrackBar1.Value
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
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

    Private Sub WebBrowser1_StatusTextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles WebBrowser1.StatusTextChanged
        ToolStripStatusLabel6.Text = WebBrowser1.StatusText
    End Sub


    'HP横幅
    Private Sub ToolStripDropDownButton6_DropDownItemClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ToolStripDropDownButton6.DropDownItemClicked
        Dim per As String = e.ClickedItem.Text
        ToolStripDropDownButton6.Text = per
        Dim perNum As String = Replace(per, "px", "")
        WebBrowser1.Width = perNum
    End Sub

    Dim kadoFlag As Boolean = True
    Private Sub IMG_GET()
        If kadoFlag = False Then
            Exit Sub
        End If

        For i As Integer = 1 To 100
            If WebBrowser1.IsBusy = False Then Exit For
            System.Threading.Thread.Sleep(1000)
            System.Windows.Forms.Application.DoEvents()
        Next

        kadoFlag = False
        Try
            WebBrowser1.Width = WebBrowser1.Document.Body.ScrollRectangle.Width
            WebBrowser1.Height = WebBrowser1.Document.Body.ScrollRectangle.Height
            Dim bmp As New Bitmap(WebBrowser1.Width, WebBrowser1.Height)
            Dim gra As Graphics = Graphics.FromImage(bmp)
            Dim hdc As IntPtr = gra.GetHdc
            Dim web As IntPtr = System.Runtime.InteropServices.Marshal.GetIUnknownForObject(WebBrowser1.ActiveXInstance)
            Dim rect As Rectangle = New Rectangle(0, 0, bmp.Width, bmp.Height)
            OleDraw(web, DVASPECT_CONTENT, hdc, rect)
            System.Runtime.InteropServices.Marshal.Release(web)
            gra.Dispose()

            Dim iName As String = Path.GetDirectoryName(Form1.appPath) & "\webbrowser.bmp"
            bmp.Save(iName)

            PictureBox1.Size = New Size(bmp.Width, bmp.Height)
            picSize = New Size(bmp.Width, bmp.Height)
            PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
            PictureBox1.ImageLocation = iName

            'Dim perNum As String = Replace(ToolStripSplitButton1.Text, "%", "")
            'PictureBox1.Size = New Size((picSize.Width / 100) * perNum, (picSize.Height / 100) * perNum)

            If ONToolStripMenuItem.Checked = True Then
                '    SplitContainer1.SplitterDistance = Me.Width - ((bmp.Width / 90) * perNum)
            End If

            Panel1.AutoScrollPosition = New Point(-scX, -scY)
            PictureBox1.Location = New Point(posX, posY)
        Catch ex As Exception

        End Try
        kadoFlag = True
    End Sub

    Private Sub 保存パスのリセットToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 保存パスのリセットToolStripMenuItem.Click
        ToolStripStatusLabel1.Text = "..."
        Application.DoEvents()

        IMG_GET()
    End Sub

    Private Sub フォームのリセットToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles フォームのリセットToolStripMenuItem.Click
        Dim frm As Form = New HTMLcreate
        Me.Dispose()
        frm.Show()
    End Sub

    Dim scX As Integer = 0
    Dim scY As Integer = 0
    Dim posX As Integer = 0
    Dim posY As Integer = 0
    Private Sub Panel1_Scroll(ByVal sender As Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles Panel1.Scroll
        scX = Panel1.AutoScrollPosition.X
        scY = Panel1.AutoScrollPosition.Y
        posX = PictureBox1.Location.X
        posY = PictureBox1.Location.Y
    End Sub

    Private Sub 切り取りToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 切り取りToolStripMenuItem.Click
        az.Cut()
    End Sub

    Private Sub コピーToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles コピーToolStripMenuItem.Click
        az.Copy()
    End Sub

    Private Sub 貼り付けToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 貼り付けToolStripMenuItem.Click
        az.Paste()
    End Sub

    Private Sub 全て選択ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 全て選択ToolStripMenuItem.Click
        az.SelectAll()
    End Sub

    'コンテキストメニュー1
    Private Sub 切り取りToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 切り取りToolStripMenuItem2.Click
        az.Cut()
    End Sub

    Private Sub コピーToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles コピーToolStripMenuItem2.Click
        az.Copy()
    End Sub

    Private Sub 貼り付けToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 貼り付けToolStripMenuItem2.Click
        az.Paste()
    End Sub

    Private Sub 上書き保存ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 上書き保存ToolStripMenuItem.Click
        If az.Text <> "" Then
            Dim FileName As String = ToolStripStatusLabel1.Text
            SaveCsv(FileName, ToolStripStatusLabel5.Text)
            'MsgBox("上書き保存しました", MsgBoxStyle.SystemModal)
        Else
            SaveCsv("", ToolStripStatusLabel5.Text)
        End If
    End Sub

    Private Sub 名前を付けて保存ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 名前を付けて保存ToolStripMenuItem.Click
        SaveCsv("", ToolStripStatusLabel5.Text)
    End Sub

    Public Sub SaveCsv(ByVal fp As String, ByVal enc As String)
        If fp = "" Or fp = "..." Then
            Dim sfd As New SaveFileDialog With {
                .InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                .Filter = "HTMLファイル(*.html)|*.html|TXTファイル(*.txt)|*.txt|すべてのファイル(*.*)|*.*",
                    .AutoUpgradeEnabled = False,
                .FilterIndex = 0,
                .Title = "保存先のファイルを選択してください",
                .RestoreDirectory = True,
                .OverwritePrompt = True,
                .CheckPathExists = True
            }
            If ToolStripStatusLabel1.Text = "" Then
                sfd.FileName = "新しいファイル.html"
            Else
                Dim pName As String = "新しいファイル"
                If ToolStripStatusLabel1.Text <> "..." Then
                    pName = ToolStripStatusLabel1.Text
                End If
                sfd.FileName = Path.GetFileName(pName)
            End If

            'ダイアログを表示する
            If sfd.ShowDialog(Me) = DialogResult.OK Then
                'SaveCsv(sfd.FileName, ToolStripStatusLabel5.Text)
                ' CSVファイルオープン
                Dim en As EncodingJP = DecodingName(enc)
                TextFileEncoding.SaveToFile(sfd.FileName, az.Text, en)
                'Dim sw As IO.StreamWriter = New IO.StreamWriter(sfd.FileName, False, System.Text.Encoding.GetEncoding(enc))
                'sw.Write(az.Text)
                'sw.Close()

                ToolStripStatusLabel1.Text = sfd.FileName
                MsgBox(sfd.FileName & vbCrLf & "保存しました", MsgBoxStyle.SystemModal)
            End If
        Else
            ' CSVファイルオープン
            Dim en As EncodingJP = DecodingName(enc)
            TextFileEncoding.SaveToFile(fp, az.Text, en)
            'Dim sw As IO.StreamWriter = New IO.StreamWriter(fp, False, System.Text.Encoding.GetEncoding(enc))
            'sw.Write(az.Text)
            'sw.Close()
            MsgBox("上書き保存しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    Private Sub SHIFTJISToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SHIFTJISToolStripMenuItem.Click
        SaveCsv("", "SJIS")
    End Sub

    Private Sub JISToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles JISToolStripMenuItem.Click
        SaveCsv("", "JIS")
    End Sub

    Private Sub EUCToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EUCToolStripMenuItem.Click
        SaveCsv("", "EUC")
    End Sub

    Private Sub UnicodeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UnicodeToolStripMenuItem.Click
        SaveCsv("", "Unicode")
    End Sub

    Private Sub UTF8ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UTF8ToolStripMenuItem.Click
        SaveCsv("", "UTF-8")
    End Sub

    Private Sub UTF8BOMToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UTF8BOMToolStripMenuItem.Click
        SaveCsv("", "UTF-8(BOM)")
    End Sub

    Private Sub UTF7ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UTF7ToolStripMenuItem.Click
        SaveCsv("", "UTF-7")
    End Sub

    Private Sub UnicodeBEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UnicodeBEToolStripMenuItem.Click
        SaveCsv("", "Unicode(Big Endian)")
    End Sub

    Private Sub ASCIIToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ASCIIToolStripMenuItem.Click
        SaveCsv("", "欧文")
    End Sub

    Private Sub 折り返し表示ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 折り返し表示ToolStripMenuItem.Click
        If 折り返し表示ToolStripMenuItem.Checked = False Then
            az.ViewType = Sgry.Azuki.ViewType.WrappedProportional
            az.ViewWidth = AzukiControl1.Width - 30
            折り返し表示ToolStripMenuItem.Checked = True
        Else
            az.ViewType = Sgry.Azuki.ViewType.Proportional
            折り返し表示ToolStripMenuItem.Checked = False
        End If
    End Sub

    Private Sub AzukiControl1_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AzukiControl1.SizeChanged
        AzukiControl1.ViewWidth = AzukiControl1.Width - 30
        AzukiControl2.ViewWidth = AzukiControl2.Width - 30
    End Sub

    'ツール
    Private Sub 片仮名全角半角ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 片仮名全角半角ToolStripMenuItem.Click
        az.Text = WideKanaToNarrow(az.Text)
    End Sub

    Private Function WideKanaToNarrow(ByVal inString As String) As String
        Dim r As Regex
        Dim mev As MatchEvaluator
        '半角にしたい文字を正規表現で指定
        r = New Regex("[ァ-ヶー]")
        mev = New MatchEvaluator(AddressOf ToNarrow)
        WideKanaToNarrow = r.Replace(inString, mev)
    End Function

    '*******************************************************************************
    Private Sub 片仮名半角全角ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 片仮名半角全角ToolStripMenuItem.Click
        az.Text = FKanaHan2Zen(az.Text)
    End Sub

    Private Function FKanaHan2Zen(ByVal myString As String) As String
        Dim i As Integer, chrKana As String, strKana As String = ""
        Dim strTemp As String = ""
        If myString Like "*[｡-ﾟ]*" Then
            For i = 0 To myString.Length - 1
                chrKana = myString.Chars(i)
                Select Case Asc(chrKana)
                    Case 161 To 223
                        strKana &= chrKana  '半角が続いたら文字をつなぐ
                    Case Else  '全角文字になったら半角の文字を全角に変換
                        If strKana.Length > 0 Then
                            strTemp &= StrConv(strKana, VbStrConv.Wide)
                            strKana = ""
                        End If
                        strTemp &= chrKana
                End Select
            Next i
            If strKana.Length > 0 Then  '最後の文字が半角の場合の処理
                strTemp &= StrConv(strKana, VbStrConv.Wide)
            End If
        Else
            Return myString
        End If
        Return strTemp
    End Function
    '*******************************************************************************

    Private Sub 英数大文字小文字ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 英数大文字小文字ToolStripMenuItem.Click
        az.Text = StrConv(az.Text, VbStrConv.Lowercase, &H411)
    End Sub

    '*******************************************************************************
    Private Sub 英数全角半角ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 英数全角半角ToolStripMenuItem.Click
        az.Text = WideAlphaNumToNarrow(az.Text)
    End Sub

    Private Function WideAlphaNumToNarrow(ByVal inString As String) As String
        Dim r As Regex
        Dim mev As MatchEvaluator
        '半角にしたい文字を正規表現で指定
        r = New Regex("[０-９]+|[ａ-ｚ]+|[Ａ-Ｚ]+|，|．")
        mev = New MatchEvaluator(AddressOf ToNarrow)
        WideAlphaNumToNarrow = r.Replace(inString, mev)
    End Function

    Private Shared Function ToNarrow(ByVal m As Match) As String
        ToNarrow = StrConv(m.ToString, VbStrConv.Narrow)
    End Function
    '*******************************************************************************

    Private Sub 検索置換ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 検索置換ToolStripMenuItem.Click
        Search.Size = New Size(412, 454)
        Search.TabControl1.SelectedIndex = 0
        Search.Show()
    End Sub

    Private Sub 行目表示ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 行目表示ToolStripMenuItem.Click
        az.Document.SetCaretIndex(0, 0)
        az.ScrollToCaret()
    End Sub

    Private Sub 最終行表示ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 最終行表示ToolStripMenuItem.Click
        Dim r As Integer = az.Document.LineCount
        az.Document.SetCaretIndex(r - 1, 0)
        az.ScrollToCaret()
    End Sub


    '***************************************************************************************************
    'インテリセンス関連
    '***************************************************************************************************
    Dim keywordA As String = ""
    Dim spaceCount As String = ""
    Private Sub AzukiControl1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles AzukiControl1.KeyPress
        Select Case e.KeyChar
            Case ControlChars.Back
                If ToolStripStatusLabel2.Text.Length > 0 Then
                    If spaceCount.Length > 0 Then
                        spaceCount = spaceCount.Substring(0, spaceCount.Length - 1)
                    Else
                        ToolStripStatusLabel2.Text = ToolStripStatusLabel2.Text.Substring(0, ToolStripStatusLabel2.Text.Length - 1)
                    End If
                End If
            Case " "
                spaceCount &= " "
                InteliShow()
            Case Else
                ToolStripStatusLabel2.Text &= e.KeyChar
        End Select
    End Sub

    Private Sub AzukiControl1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles AzukiControl1.KeyUp
        Select Case e.KeyCode
            Case Keys.Enter
                InteliClose()
            Case Keys.Delete, Keys.Back
                InteliShow()
            Case Keys.Space

        End Select

    End Sub

    Private Sub AzukiControl1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles AzukiControl1.KeyDown
        If e.Control And e.KeyCode = Keys.B Then
            PasteImgUrl()
        End If

    End Sub

    Private Function PasteImgUrl() As Boolean
        Dim lineNum As Integer = AzukiControl1.GetLineIndexFromCharIndex(AzukiControl1.CaretIndex)
        Dim startS As Integer = AzukiControl1.GetLineHeadIndex(lineNum)
        Dim endS As Integer = startS + AzukiControl1.GetLineLength(lineNum)
        Dim str As String = AzukiControl1.Document.GetLineContent(AzukiControl1.GetLineIndexFromCharIndex(AzukiControl1.CaretIndex))
        Dim clipStr As String = Clipboard.GetText()
        Dim setStr As String = ""

        If InStr(clipStr, ".jpg") > 0 Or InStr(clipStr, ".gif") > 0 Or InStr(clipStr, ".png") > 0 Then
            Dim sArray1 As String() = Split(str, """")
            If sArray1.Length > 0 Then
                For i As Integer = 0 To sArray1.Length - 1
                    If InStr(sArray1(i), ".jpg") > 0 Or InStr(sArray1(i), ".gif") > 0 Or InStr(sArray1(i), ".png") > 0 Then
                        Dim DR As DialogResult = MsgBox(sArray1(i) & vbCrLf & "→ " & clipStr & vbCrLf & "変更します。絶対パスはyes、相対パスはno。", MsgBoxStyle.YesNoCancel Or MsgBoxStyle.SystemModal)
                        If DR = Windows.Forms.DialogResult.Yes Then
                            If setStr <> "" Then
                                setStr &= """"
                            End If
                            setStr &= clipStr
                        ElseIf DR = Windows.Forms.DialogResult.No Then
                            If setStr <> "" Then
                                setStr &= """"
                            End If
                            setStr &= Path.GetFileName(clipStr)
                        Else
                            If setStr <> "" Then
                                setStr &= """"
                            End If
                            setStr &= sArray1(i)
                        End If
                    Else
                        If setStr <> "" Then
                            setStr &= """"
                        End If
                        setStr &= sArray1(i)
                    End If
                Next
                AzukiControl1.Document.Replace(setStr, startS, endS)
            End If
        ElseIf InStr(clipStr, ".html") > 0 Or InStr(clipStr, ".htm") > 0 Or InStr(clipStr, "http") > 0 Or InStr(clipStr, "https") > 0 Then
            Dim sArray1 As String() = Split(str, """")
            If sArray1.Length > 0 Then
                For i As Integer = 0 To sArray1.Length - 1
                    If InStr(sArray1(i), ".html") > 0 Or InStr(sArray1(i), ".htm") > 0 Or InStr(sArray1(i), "http") > 0 Or InStr(sArray1(i), "https") > 0 Then
                        Dim DR As DialogResult = MsgBox(sArray1(i) & vbCrLf & "→ " & clipStr & vbCrLf & "変更します。絶対パスはyes、相対パスはno。", MsgBoxStyle.YesNoCancel Or MsgBoxStyle.SystemModal)
                        If DR = Windows.Forms.DialogResult.Yes Then
                            If setStr <> "" Then
                                setStr &= """"
                            End If
                            setStr &= clipStr
                        ElseIf DR = Windows.Forms.DialogResult.No Then
                            If setStr <> "" Then
                                setStr &= """"
                            End If
                            setStr &= Path.GetFileName(clipStr)
                        Else
                            If setStr <> "" Then
                                setStr &= """"
                            End If
                            setStr &= sArray1(i)
                        End If
                    Else
                        If setStr <> "" Then
                            setStr &= """"
                        End If
                        setStr &= sArray1(i)
                    End If
                Next
                AzukiControl1.Document.Replace(setStr, startS, endS)
            End If
        End If
    End Function

    Private Sub AzukiControl1_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles AzukiControl1.MouseClick
        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                InteliClose()
        End Select
    End Sub

    Private Sub InteliShow()
        If keywordA.Length = 0 Then
            InteliClose()
            Exit Sub
        End If

        ListBox1.Items.Clear()

        Dim num As Integer = 0
        Dim tagFlag As Boolean = False
        For i As Integer = 0 To inteliList.Count - 1
            If inteliList(i).ToString.Substring(0, 1) <> "," Then
                If keywordA.Length < inteliList(i).ToString.Length Then
                    If keywordA = inteliList(i).ToString.Substring(0, keywordA.Length) Then
                        ListBox1.Items.Add(inteliList(i))
                        tagFlag = True
                        num += 1
                    Else
                        'ToolStripStatusLabel1.Text = num
                        tagFlag = False
                    End If
                Else
                    tagFlag = False
                End If
            Else
                Dim keywordB As String = Replace(inteliList(i).ToString, ",", "")
                If tagFlag = True Then
                    ListBox1.Items.Add(keywordB)
                    num += 1
                Else
                    If InStr(keywordB, keywordA) > 0 Then
                        ListBox1.Items.Add(keywordB)
                        num += 1
                    End If
                End If
            End If
        Next

        If num = 0 Then
            InteliClose()
        Else
            If num > 20 Then
                num = 20
            End If
            ListBox1.Size = New Size(ListBox1.Size.Width, 12 * num + 12)
            Dim p As Point = AzukiControl1.GetPositionFromIndex(AzukiControl1.Document.CaretIndex)
            ListBox1.Location = New Point(p.X + 2, p.Y + 22)
            ListBox1.Show()
            Me.ListBox1.BringToFront()
            Timer1.Start()
        End If
    End Sub

    Private Sub InteliClose()
        Timer1.Stop()
        timers = 0
        keywordA = ""
        spaceCount = ""
        ListBox1.Hide()

        ToolStripStatusLabel2.Text = keywordA
    End Sub

    Dim kakomi As String = ""
    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        AZ_SETTEXT("<b>", "</b>")
    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        AZ_SETTEXT("<i>", "</i>")
    End Sub

    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles ToolStripButton5.Click
        AZ_SETTEXT("<s>", "</s>")
    End Sub

    Private Sub ToolStripButton6_Click(sender As Object, e As EventArgs) Handles ToolStripButton6.Click
        AZ_SETTEXT("<a href=" & kakomi & kakomi & ">", "</a>")
    End Sub

    Private Sub ToolStripButton7_Click(sender As Object, e As EventArgs) Handles ToolStripButton7.Click

    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        AZ_SETTEXT("<img src=" & kakomi & kakomi & " border=0>", "")
    End Sub

    Private Sub ToolStripButton8_Click(sender As Object, e As EventArgs) Handles ToolStripButton8.Click
        Dim str As String = "<table border=" & kakomi & "0" & kakomi & " cellpadding=" & kakomi & "0" & kakomi & " cellspacing=" & kakomi & "0" & kakomi & ">"
        str &= vbLf & "<tr><td>" & vbLf & "</td></tr>" & vbLf & "</table>"
        AZ_SETTEXT(str, "")
    End Sub

    Private Sub ToolStripButton9_Click(sender As Object, e As EventArgs) Handles ToolStripButton9.Click
        AZ_SETTEXT("<div align=" & kakomi & "left" & kakomi & ">", "</div>")
    End Sub

    Private Sub ToolStripButton10_Click(sender As Object, e As EventArgs) Handles ToolStripButton10.Click
        AZ_SETTEXT("<div align=" & kakomi & "right" & kakomi & ">", "</div>")
    End Sub

    Private Sub ToolStripButton11_Click(sender As Object, e As EventArgs) Handles ToolStripButton11.Click
        AZ_SETTEXT("<div align=" & kakomi & "center" & kakomi & ">", "</center>")
    End Sub

    Private Sub ToolStripButton12_Click(sender As Object, e As EventArgs) Handles ToolStripButton12.Click
        AZ_SETTEXT("<br>", "")
    End Sub

    Private Sub ToolStripButton13_Click(sender As Object, e As EventArgs) Handles ToolStripButton13.Click
        AZ_SETTEXT("<font size=" & kakomi & "3" & kakomi & ">", "</font>")
    End Sub

    Private Sub ToolStripButton14_Click(sender As Object, e As EventArgs) Handles ToolStripButton14.Click
        AZ_SETTEXT("<font color=" & kakomi & "red" & kakomi & ">", "</font>")
    End Sub

    Private Sub ToolStripButton15_Click(sender As Object, e As EventArgs) Handles ToolStripButton15.Click
        Dim str As String = "<iframe src=" & kakomi & kakomi
        str &= " width=" & kakomi & kakomi
        str &= " height=" & kakomi & kakomi
        str &= " scrolling=" & kakomi & "no" & kakomi
        str &= " frameborder=" & kakomi & "no" & kakomi
        str &= ">"
        AZ_SETTEXT(str, "</iframe>")
    End Sub

    Private Sub ToolStripButton16_Click(sender As Object, e As EventArgs) Handles ToolStripButton16.Click
        If ToolStripButton16.Text = "'" Then
            ToolStripButton16.Text = """"
            kakomi = """"
        ElseIf ToolStripButton16.Text = "×" Then
            ToolStripButton16.Text = "'"
            kakomi = "'"
        Else
            ToolStripButton16.Text = "×"
            kakomi = ""
        End If
    End Sub

    Private Sub AZ_SETTEXT(str1, str2)
        Dim az As AzukiControl = AzukiControl1
        If TabControl1.SelectedTab.Text = "HTML2" Then
            az = AzukiControl2
        End If

        Dim startS As Integer = az.CaretIndex
        Dim endS As Integer = startS + az.GetSelectedText.Length
        Dim selLength As Integer = az.GetSelectedText.Length - 1

        az.Document.Replace(str1, startS, startS)
        Dim endS2 As Integer = endS + str1.length
        If str2 <> "" Then
            az.Document.Replace(str2, endS2, endS2)
        End If

        'AzukiControl1.SetSelection(startS, endS2 + selLength)
    End Sub

    Dim timers As Integer = 0
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If timers > 100 Then
            InteliClose()
        Else
            timers += 1
        End If
    End Sub

    Private Sub ListBox1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.DoubleClick
        AzukiControl1.Focus()

        Dim start As Integer = AzukiControl1.CaretIndex - keywordA.Length + spaceCount.Length
        AzukiControl1.Document.Replace(ListBox1.SelectedItem, start, AzukiControl1.CaretIndex)
        keywordA = ""
        spaceCount = ""

        'If InStr(ListBox1.SelectedItem, "<") > 0 Then
        '    Dim start As Integer = AzukiControl1.CaretIndex - keywordA.Length
        '    AzukiControl1.Document.Replace(ListBox1.SelectedItem, start, AzukiControl1.CaretIndex)
        '    inteliClose()
        'Else
        '    Dim start As Integer = AzukiControl1.CaretIndex - spaceCount.Length
        '    AzukiControl1.Document.Replace(ListBox1.SelectedItem, start, AzukiControl1.CaretIndex)
        '    spaceCount = ""
        'End If
    End Sub

    Private Sub ToolStripStatusLabel2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripStatusLabel2.TextChanged
        keywordA = ToolStripStatusLabel2.Text
        InteliShow()
    End Sub
    '***************************************************************************************************
    'インテリセンス関連
    '***************************************************************************************************

    Private Sub RGB入力ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RGB入力ToolStripMenuItem.Click
        If ColorDialog1.ShowDialog() = DialogResult.OK Then
            '選択された色の取得
            'Dim start As Integer = az.CaretIndex - az.GetSelectedText.Length
            'Dim endS As Integer = az.CaretIndex

            Dim start As Integer = 0
            Dim endS As Integer = 0
            If az.Document.AnchorIndex < az.CaretIndex Then
                start = az.CaretIndex - az.GetSelectedText.Length
                endS = az.CaretIndex
            Else
                start = az.CaretIndex
                endS = az.CaretIndex + az.GetSelectedText.Length
            End If
            Dim clR As String = Convert.ToString(ColorDialog1.Color.R, 16)
            Dim clG As String = Convert.ToString(ColorDialog1.Color.G, 16)
            Dim clB As String = Convert.ToString(ColorDialog1.Color.B, 16)
            az.Document.Replace("#" & clR.PadLeft(2, "0"c) & clG.PadLeft(2, "0"c) & clB.PadLeft(2, "0"c), start, endS)
        End If
    End Sub

    Private Sub 行左スペース除去ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 行左スペース除去ToolStripMenuItem.Click
        For r As Integer = 0 To az.LineCount - 1
            Dim start As Integer = az.Document.GetLineHeadIndex(r)
            Dim ends As Integer = start + az.Document.GetLineLength(r)
            Dim str As String = az.Document.GetLineContent(r)
            str = LTrim(str)
            az.Document.Replace(str, start, ends)
        Next
    End Sub

    Private Sub 行右スペース除去ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 行右スペース除去ToolStripMenuItem.Click
        For r As Integer = 0 To az.LineCount - 1
            Dim start As Integer = az.Document.GetLineHeadIndex(r)
            Dim ends As Integer = start + az.Document.GetLineLength(r)
            Dim str As String = az.Document.GetLineContent(r)
            str = RTrim(str)
            az.Document.Replace(str, start, ends)
        Next
    End Sub


    '***************************************************************************************************
    '右クリックコンテクストメニュー
    '***************************************************************************************************
    Private Sub 画像ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 画像ToolStripMenuItem1.Click
        SelFileName(1)
    End Sub

    Private Sub HTMLToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HTMLToolStripMenuItem1.Click
        SelFileName(2)
    End Sub

    Private Sub 画像ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 画像ToolStripMenuItem.Click
        SelFileName(3)
    End Sub

    Private Sub HTMLToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HTMLToolStripMenuItem.Click
        SelFileName(4)
    End Sub

    Private Sub SelFileName(ByVal a As Integer)
        Dim num As Integer = az.Document.CaretIndex
        Dim LineNum As Integer = az.Document.GetLineIndexFromCharIndex(num)
        Dim start As Integer = az.Document.GetLineHeadIndex(LineNum)
        Dim ends As Integer = start + az.Document.GetLineLength(LineNum)
        Dim str As String = az.Document.GetLineContent(LineNum)
        Dim sArray As String() = Split(str, """")

        If a = 3 Then
            For i = 0 To sArray.Length - 1
                Select Case True
                    Case Regex.IsMatch(sArray(i), ".*jpg.*|.*jpeg.*|.*gif.*|.*png.*")
                        Dim SR As Sgry.Azuki.SearchResult = az.Document.FindNext(sArray(i), start)
                        az.Document.SetSelection(SR.Begin, SR.End)
                        az.ScrollToCaret()
                        Exit For
                End Select
            Next
        ElseIf a = 4 Then
            For i = 0 To sArray.Length - 1
                Select Case True
                    Case Regex.IsMatch(sArray(i), ".*http:.*|.*html.*|.*htm.*|.*php.*")
                        Dim SR As Sgry.Azuki.SearchResult = az.Document.FindNext(sArray(i), start)
                        az.Document.SetSelection(SR.Begin, SR.End)
                        az.ScrollToCaret()
                        Exit For
                End Select
            Next
        Else
            For i = 0 To sArray.Length - 1
                Dim flag As Boolean = False
                Dim len As Integer = 0

                If a = 1 And (InStr(sArray(i), ".jpg") > 0 Or InStr(sArray(i), ".gif") > 0 Or InStr(sArray(i), ".png") > 0) Then
                    len = 4
                    flag = True
                ElseIf a = 2 And (InStr(sArray(i), ".html") > 0) Then
                    len = 5
                    flag = True
                ElseIf a = 2 And (InStr(sArray(i), ".htm") > 0 Or InStr(sArray(i), ".php") > 0) Then
                    len = 4
                    flag = True
                End If

                If flag = True Then
                    Dim sName As String = Path.GetFileName(sArray(i))
                    Dim SR As Sgry.Azuki.SearchResult = az.Document.FindNext(sName, start)
                    If Not SR Is Nothing Then
                        az.Document.SetSelection(SR.Begin, SR.End - len)
                        az.ScrollToCaret()
                        Exit For
                    End If
                End If
            Next
        End If
    End Sub

    Private Sub 価格画像入力ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 価格画像入力ToolStripMenuItem.Click
        Dim IB As String = InputBox("価格")
        If IsNumeric(IB) Then
            IB = Format(CInt(IB), "#,###")
            Dim mstr As String = "<img src='http://shopping.c.yimg.jp/lib/fkstyle/mail-price-NN.jpg'>"
            Dim str As String = ""

            str = Replace(mstr, "NN", "en")
            For i As Integer = 0 To IB.Length - 1
                If IB(i) = "," Then
                    str &= Replace(mstr, "NN", "ten")
                Else
                    str &= Replace(mstr, "NN", IB(i))
                End If
            Next

            If az.Document.AnchorIndex < az.CaretIndex Then
                Dim start As Integer = az.CaretIndex - az.GetSelectedText.Length
                Dim endS As Integer = az.CaretIndex
                az.Document.Replace(str, start, endS)
            Else
                Dim start As Integer = az.CaretIndex
                Dim endS As Integer = az.CaretIndex + az.GetSelectedText.Length
                az.Document.Replace(str, endS, start)
            End If
        End If
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        If TabControl1.SelectedIndex = 0 Then
            az = AzukiControl1
            ToolStrip1.Enabled = True
            If ブラウザ表示ToolStripMenuItem.Checked = True Then
                SplitContainer1.Panel2Collapsed = False
            Else
                SplitContainer1.Panel2Collapsed = True
            End If
        ElseIf TabControl1.SelectedIndex = 1 Then
            az = AzukiControl2
            ToolStrip1.Enabled = True
            If ブラウザ表示ToolStripMenuItem.Checked = True Then
                SplitContainer1.Panel2Collapsed = False
            Else
                SplitContainer1.Panel2Collapsed = True
            End If
        Else
            ToolStrip1.Enabled = False
            SplitContainer1.Panel2Collapsed = True
        End If
    End Sub

    Private Sub ToolStripMenuItem16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem16.Click
        Dim num As Integer = az.Document.CaretIndex
        Dim LineNum As Integer = az.Document.GetLineIndexFromCharIndex(num)
        Dim start As Integer = az.Document.GetLineHeadIndex(LineNum)
        Dim ends As Integer = start + az.Document.GetLineLength(LineNum)
        'Dim str As String = az.Document.GetLineContent(LineNum)
        'Dim sArray As String() = Split(str, """")
        az.Document.SetSelection(start, ends)
        az.ScrollToCaret()
    End Sub

    Private Sub AzukiControl1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AzukiControl1.TextChanged
        ToolStripStatusLabel3.Text = Format(AzukiControl1.TextLength, "00000")
        HtmlCheck()
    End Sub

    Private Sub AzukiControl2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AzukiControl2.TextChanged
        ToolStripStatusLabel4.Text = Format(AzukiControl2.TextLength, "00000")
        HtmlCheck()
    End Sub

    Dim checkTag As String = ""
    Private Sub HtmlCheck()
        Dim startArray As String() = New String() {"<table", "<tr", "<td", "<p", "<div", "<span", "<ul", "<li", "<dl", "<iframe", "<a "}
        Dim endArray As String() = New String() {"</table>", "</tr>", "</td>", "</p>", "</div>", "</span>", "</ul>", "</li>", "</dl>", "</iframe>", "</a>"}

        checkTag = ""
        Dim str As String = ""
        For i As Integer = 0 To startArray.Length - 1
            Dim a As Integer = CountChar(az.Text, startArray(i))
            Dim b As Integer = CountChar(az.Text, endArray(i))
            If a < b Then
                If str = "" Then
                    str &= "タグ不足→"
                End If
                str &= "[" & endArray(i) & "]"
                If checkTag = "" Then
                    checkTag = startArray(i) & "," & endArray(i)
                End If
            ElseIf a > b Then
                If str = "" Then
                    str &= "タグ不足→"
                End If
                str &= "[" & startArray(i) & "]"
                If checkTag = "" Then
                    checkTag = startArray(i) & "," & endArray(i)
                End If
            End If
        Next

        If str <> "" Then
            TextBox1.BackColor = Color.Red
            TextBox1.ForeColor = Color.White
            TextBox1.Text = str
        Else
            TextBox1.BackColor = Color.White
            TextBox1.ForeColor = Color.Black
            TextBox1.Text = "HTML簡易チェックOK"
        End If
    End Sub

    '閉じタグの無い行を調べる
    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If checkTag <> "" Then
            Dim cArray As String() = Split(checkTag, ",")
            Dim sArray As String() = Split(az.Text, cArray(0))

            For i As Integer = 0 To sArray.Length - 1
                If i > 0 Then
                    If InStr(sArray(i), cArray(1)) = 0 Then
                        Dim lines As String() = Split(sArray(i), vbCrLf)
                        TagSel(lines(0))
                        Exit For
                    End If
                End If
            Next

        End If
    End Sub

    Public Shared Function CountChar(ByVal s As String, ByVal c As String) As Integer
        If c = "<li" Then   '<liと<linkの対策
            s = s.Replace("<link", "<xxxx")
        End If

        Dim d As Integer = s.Length - s.Replace(c, "").Length
        Return d / c.Length
    End Function

    Private Sub TagSel(ByVal comment As String)
        Dim startLine As Integer = 0
        Dim endLine As Integer = 0

        Dim lines As String() = Split(az.Text, vbCrLf)
        For i As Integer = 0 To lines.Length - 1
            If InStr(lines(i), comment) > 0 Then
                startLine = az.GetLineHeadIndex(i)
                endLine = az.GetLineHeadIndex(i) + az.GetLineLength(i)
                Exit For
            End If
        Next

        az.SetSelection(startLine, endLine)
        az.ScrollToCaret()
    End Sub

    Private Sub ONToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ONToolStripMenuItem.Click
        ONToolStripMenuItem.Checked = True
        OFFToolStripMenuItem.Checked = False
        IMG_GET()
    End Sub

    Private Sub OFFToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OFFToolStripMenuItem.Click
        ONToolStripMenuItem.Checked = False
        OFFToolStripMenuItem.Checked = True
    End Sub

    'HTML更新
    Private TextBoxNewSize As Size  '移動位置の保存用変数

    'Private Sub TabControl2_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TabControl2.MouseDown
    '    If e.Button = System.Windows.Forms.MouseButtons.Left Then
    '        'ドラッグ開始時点の位置を取得
    '        TextBoxNewSize = New Size(e.X, e.Y)
    '    End If
    'End Sub

    'Private Sub TabControl2_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TabControl2.MouseMove
    '    If e.Button = System.Windows.Forms.MouseButtons.Left Then
    '        'ドラッグ中の位置情報を取得して、その位置に表示
    '        TabControl2.Location = Point.op_Subtraction( _
    '              Me.PointToClient(System.Windows.Forms.Cursor.Position), TextBoxNewSize)
    '    End If
    'End Sub

    Dim mode As Integer = 0
    Private Sub 解析ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 解析ToolStripMenuItem.Click
        Dim list As String() = Split(az.GetSelectedText, vbCrLf)

        Dim comment1 As String = "newitems.gif"
        Dim comment2 As String = "recommenditem.gif"

        For i As Integer = 0 To list.Length - 1
            If InStr(list(i), comment1) > 0 Or InStr(list(i), comment2) > 0 Then
                mode = 1
                Exit For
            End If
        Next

        Dim str As String = ""
        Select Case mode
            Case 1
                'az.Enabled = False
                'HTMLdialog.TabControl2.SelectTab("TabPage3")
                Form1.SetTabVisible(HTMLdialog.TabControl2, "新商品")
                Dim lFlag As Boolean = False
                For i As Integer = 0 To list.Length - 1
                    If InStr(list(i), comment1) > 0 Or InStr(list(i), comment2) > 0 Then
                        lFlag = True
                    ElseIf lFlag = True And (InStr(list(i), "<figure>") > 0 Or InStr(list(i), "</figure>") > 0) Then
                        str &= list(i) & vbCrLf
                    ElseIf InStr(list(i), "</ul>") > 0 Then
                        mode = 0
                        Exit For
                    End If
                Next
                'az = AzukiControl3
            Case Else
                Exit Sub
        End Select

        'AzukiControl3.Text = str
        'TabControl2.Size = New Size(704, 600)
        'TabControl2.Visible = True
        HTMLdialog.AzukiControl3.Text = str
        HTMLdialog.Size = New Size(818, 569)
        HTMLdialog.Show()
    End Sub

    '更新
    Public Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If TabControl1.SelectedTab.Text = "HTML1" Then
            az = AzukiControl1
        ElseIf TabControl1.SelectedTab.Text = "HTML2" Then
            az = AzukiControl2
        End If

        Dim list1 As String() = Split(az.GetSelectedText, vbCrLf)
        Dim list2 As String() = Split(HTMLdialog.AzukiControl3.Text, vbCrLf)

        Dim str As String = ""
        Dim k As Integer = 0
        For i As Integer = 0 To list1.Length - 1
            If InStr(list1(i), "<figure>") > 0 Or InStr(list1(i), "</figure>") > 0 Then
                str &= list2(k) & vbCrLf
                k += 1
            Else
                str &= list1(i) & vbCrLf
            End If
        Next

        If az.Document.AnchorIndex < az.CaretIndex Then
            Dim start As Integer = az.CaretIndex - az.GetSelectedText.Length
            Dim endS As Integer = az.CaretIndex
            az.Document.Replace(str, start, endS)
        Else
            Dim start As Integer = az.CaretIndex
            Dim endS As Integer = az.CaretIndex + az.GetSelectedText.Length
            az.Document.Replace(str, endS, start)
        End If

        az.Enabled = True
        'TabControl2.Visible = False
        HTMLdialog.Close()
    End Sub

    Private Sub NewitemsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewitemsToolStripMenuItem.Click
        SeleAz("<!-- newitems -->")
    End Sub

    Private Sub RecommenditemsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RecommenditemsToolStripMenuItem.Click
        SeleAz("<!-- recommenditems -->")
    End Sub

    Private Sub SeleAz(ByVal comment As String)
        Dim startLine As Integer = 0
        Dim endLine As Integer = 0

        Dim lines As String() = Split(az.Text, vbCrLf)
        For i As Integer = 0 To lines.Length - 1
            If InStr(lines(i), comment) > 0 Then
                If startLine = 0 Then
                    startLine = az.GetLineHeadIndex(i)
                Else
                    endLine = az.GetLineHeadIndex(i) + az.GetLineLength(i)
                End If
            End If
        Next

        az.SetSelection(startLine, endLine)
        az.ScrollToCaret()
    End Sub

    Private Sub 商品テーブル作成ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 商品テーブル作成ToolStripMenuItem.Click
        'HTMLdialog.TabControl2.SelectTab("TabPage4")
        'Form1.SetTabVisible(HTMLdialog.TabControl2, "商品テーブル作成")
        'HTMLdialog.Size = New Size(818, 569)
        'HTMLdialog.Show()
        HtmlDialog_F_TableCreate.Show()
    End Sub

    Private Sub Itemlist更新ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Itemlist更新ToolStripMenuItem.Click
        'HTMLdialog.TabControl2.SelectTab("TabPage1")
        Form1.SetTabVisible(HTMLdialog.TabControl2, "itemlist更新")
        HTMLdialog.Size = New Size(818, 569)
        HTMLdialog.Show()
    End Sub

    Private Sub リスト更新ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles リスト更新ToolStripMenuItem.Click
        'HTMLdialog.TabControl2.SelectTab("TabPage9")
        Form1.SetTabVisible(HTMLdialog.TabControl2, "リスト更新")
        HTMLdialog.Size = New Size(818, 569)
        HTMLdialog.Show()
    End Sub

    Private Sub ToolStripDropDownButton8_DropDownItemClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ToolStripDropDownButton8.DropDownItemClicked
        ToolStripDropDownButton8.Text = e.ClickedItem.Text
        Dim Sizes As Size = WebBrowser1.Document.Body.ScrollRectangle.Size
        Dim w As Integer = Replace(ToolStripDropDownButton6.Text, "px", "")
        Dim p As Integer = Replace(ToolStripDropDownButton8.Text, "%", "")
        Dim per As Double = Math.Round(Panel1.Width / w * 100)

        WebBrowser1.Document.Body.ScrollLeft = scX
        WebBrowser1.Document.Body.ScrollTop = scY

        If ToolStripDropDownButton8.Text = "幅指定" Then
            TrackBar1.Value = per
            Zoom()
        Else
            TrackBar1.Value = p
            Zoom()
        End If
    End Sub

    Private Sub 計算機ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 計算機ToolStripMenuItem.Click
        HTMLdialog.Show()
        HTMLdialog.Size = New Size(439, 371)
        'HTMLdialog.TabControl2.SelectedIndex = 3
        Form1.SetTabVisible(HTMLdialog.TabControl2, "計算機")
    End Sub

    Private Sub FTPToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FTPToolStripMenuItem.Click
        HTMLdialog.Show()
        HTMLdialog.Size = New Size(818, 569)
        'HTMLdialog.TabControl2.SelectedIndex = 4
        Form1.SetTabVisible(HTMLdialog.TabControl2, "FTP")
    End Sub

    Private Sub HTMLチェッカー2ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HTMLチェッカー2ToolStripMenuItem.Click
        DataGridView1.Rows.Clear()
        TabControl2.SelectedIndex = 2
        Dim check As String() = New String() {"html", "head", "body", "table", "tr", "td", "div", "a", "span", "font", "form", "input", "select", "option", "p", "img", "ul", "li", "small", "iframe"}
        For i As Integer = 0 To check.Length - 1
            HTMLDOM(check(i))
        Next
    End Sub

    Private Sub HTMLDOM(ByRef name As String)
        DataGridView1.Rows.Add(name, "")
        DataGridView1.Item(0, DataGridView1.RowCount - 1).Style.BackColor = Color.Aquamarine

        Dim num As Integer = 1
        For Each Element As HtmlElement In WebBrowser1.Document.GetElementsByTagName(name)
            'Dim str As String = Replace(Element.OuterHtml, " ", "")
            'str = Replace(Element.OuterHtml, "　", "")
            Dim str As String = Element.OuterHtml
            DataGridView1.Rows.Add(1)
            DataGridView1.Item(0, DataGridView1.RowCount - 1).Value = num
            DataGridView1.Item(1, DataGridView1.RowCount - 1).Value = str
            num += 1

            Dim stTagCount As Integer = 0
            Dim enTagCount As Integer = 0
            Select Case name
                Case "html", "head", "body", "table", "tr", "td", "div", "a", "span", "font", "form", "select", "option", "p", "ul", "li", "small", "iframe"
                    Dim strLen As Integer = str.Length
                    Dim tag As String = "<" & name
                    Dim stTagLen As Integer = tag.Length
                    str = Replace(str, "<" & name, "")
                    stTagCount = (strLen - str.Length) / stTagLen
                    strLen = str.Length
                    tag = "</" & name & ">"
                    Dim enTagLen As Integer = tag.Length
                    str = Replace(str, tag, "")
                    enTagCount = (strLen - str.Length) / enTagLen
                Case "img", "option"

            End Select

            If stTagCount <> enTagCount Then
                DataGridView1.Item(0, DataGridView1.RowCount - 1).Style.BackColor = Color.Yellow
            End If
        Next
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        TextBox2.Text = DataGridView1.SelectedCells(0).Value
    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        Dim SearchChar As String = DataGridView1.SelectedCells(0).Value
        'Dim SR As Sgry.Azuki.SearchResult = HTMLcreate.az.Document.FindNext(SearchChar, HTMLcreate.az.Document.CaretIndex)

        Dim searchLenArray As Integer() = New Integer() {SearchChar.Length, 40, 30, 20, 10, 5, 4, 3, 2}
        For i As Integer = 0 To searchLenArray.Length - 1
            Dim SR As Sgry.Azuki.SearchResult = HTMLcreate.az.Document.FindNext(SearchChar.Substring(0, searchLenArray(i)), 0)
            If Not SR Is Nothing Then
                HTMLcreate.az.Document.SetSelection(SR.Begin, SR.End)
                HTMLcreate.az.ScrollToCaret()
                Exit For
            End If
        Next
    End Sub

    '複数ファイル処理
    Dim azFocus As WinForms.AzukiControl

    Dim azFilePath As String() = New String(12) {}


    Dim az_change As Boolean = False

    Private Sub AzTextChange()
        Dim changeLine As Integer
        Dim changeColIndex As Integer
        azFocus.Document.GetCaretIndex(changeLine, changeColIndex)
        Dim str As String = azFocus.Document.GetLineContent(changeLine)
        For i As Integer = 0 To azArray.Length - 1
            If Not azArray(i) Is azFocus Then
                If azArray(i).Text <> "" Then
                    Dim cIndex As Integer = azArray(i).Document.GetLineHeadIndex(changeLine)
                    Dim lStr As String = azArray(i).Document.GetLineContent(changeLine)
                    Dim lLen As Integer = lStr.Length
                    azArray(i).Document.Replace(str, cIndex, cIndex + lLen)
                End If
            End If
        Next
    End Sub




    Private Sub 切り取りToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 切り取りToolStripMenuItem1.Click
        azFocus.Cut()
    End Sub

    Private Sub コピーToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles コピーToolStripMenuItem1.Click
        azFocus.Copy()
    End Sub

    Private Sub 貼り付けToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 貼り付けToolStripMenuItem1.Click
        azFocus.Paste()
    End Sub

    Private Sub 全て選択ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 全て選択ToolStripMenuItem1.Click
        azFocus.SelectAll()
    End Sub

    Private Sub まとめて開くToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            For i As Integer = 0 To 8
                If azArray(i).Text <> "" Then
                    azArray(i).Text = ""
                    tbArray(i).Text = ""
                    azFilePath(i) = ""
                End If
            Next
            Dim num As Integer = 0
            For Each strFilePath As String In OpenFileDialog1.FileNames
                Dim e1 As EncodingJP
                Dim e2 As Boolean
                azArray(num).Text = TextFileEncoding.LoadFromFile(strFilePath, e1, True, e2)
                azFilePath(num) = strFilePath
                Dim strFileName As String = IO.Path.GetFileName(strFilePath)
                tbArray(num).Text = strFileName
                num += 1
                If num > 8 Then
                    Exit For
                End If
            Next
        End If
    End Sub

    Private Sub TsdbMulti_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim Num As String = Replace(sender.Name, "tsdb", "")
        azArray(Num - 1).Text = ""
        tbArray(Num - 1).Text = ""
    End Sub










    Private Sub TemplateOpen(ByVal key As String)
        For i As Integer = 0 To 8
            If azArray(i).Text <> "" Then
                azArray(i).Text = ""
                tbArray(i).Text = ""
                azFilePath(i) = ""
            End If
        Next

        Dim di As New System.IO.DirectoryInfo(Path.GetDirectoryName(Form1.appPath) & "\template")
        Dim files As System.IO.FileInfo() = di.GetFiles("*.dat", System.IO.SearchOption.AllDirectories)

        Dim sFilePath As ArrayList = New ArrayList
        For Each strFilePath As FileInfo In files
            If InStr(strFilePath.FullName, key) > 0 And InStr(strFilePath.FullName, "★") = 0 Then
                sFilePath.Add(strFilePath.FullName)
            End If
        Next

        Dim num As Integer = 0
        For Each strFilePath As String In sFilePath
            Dim e1 As EncodingJP
            Dim e2 As Boolean
            azArray(num).Text = TextFileEncoding.LoadFromFile(strFilePath, e1, True, e2)
            azFilePath(num) = strFilePath
            Dim strFileName As String = IO.Path.GetFileName(strFilePath)
            tbArray(num).Text = strFileName
            num += 1
        Next
    End Sub




    'コードフォーマット
    Private Function Array_Create(ByVal s As String) As ArrayList
        s = Regex.Replace(s, ">", ">" & vbCrLf)
        Dim array As String() = Split(s, vbCrLf)

        For i As Integer = 0 To array.Length - 1
            array(i) = array(i).TrimStart()
        Next

        Dim str As New ArrayList(array)
        Return str
    End Function

    Private Function BR_ADD(ByVal htmlRecords As ArrayList) As ArrayList
        Dim str As ArrayList = New ArrayList

        Dim tagBR As String() = New String() {"<center", "<table", "</table>", "<td ", "</td>", "</tr>", "<img ", "<a ", "<b>", "<font", "br>", "<div"}

        For i As Integer = 0 To htmlRecords.Count - 1
            Dim flag As Boolean = False
            For c As Integer = 0 To tagBR.Length - 1
                If InStr(htmlRecords(i), tagBR(c)) > 0 Then
                    str.Add(htmlRecords(i) & vbCrLf)
                    flag = True
                    Exit For
                End If
            Next
            If flag = False Then
                str.Add(htmlRecords(i))
            End If
        Next

        Return str
    End Function

    Private Function HTMLTIDY(ByVal htmlstr As String)
        Dim htmlRecords As ArrayList = New ArrayList

        'ZetaHtmlTidy
        Dim html As String = htmlstr
        Dim s As String = ""
        Using tidy As HtmlTidy = New HtmlTidy()
            s = tidy.CleanHtml(html, HtmlTidyOptions.ConvertToXhtml)
        End Using

        s = Replace(s, vbCrLf, "")
        s = Replace(s, vbLf, "")
        s = Replace(s, vbCr, "")
        s = Regex.Replace(s, "<!DOCTYPE h.+HTML Tidy.+head><body>", "")
        's = Regex.Replace(s, "<!DOCTYPE h.+HTML Tidy.+/>", "")
        's = Regex.Replace(s, "<meta.+HTML Tidy.+head><body>", "")
        s = Replace(s, "</body></html>", "")
        s = Replace(s, " />", ">")

        htmlRecords = Array_Create(s)
        htmlRecords = BR_ADD(htmlRecords)

        Dim str As String = ""
        For k As Integer = 0 To htmlRecords.Count - 1
            str &= htmlRecords(k)
        Next

        Return str
    End Function

    Private Sub 全てToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 全てToolStripMenuItem.Click
        az.Text = HTMLTIDY(az.Text)
        MsgBox("終了しました", MsgBoxStyle.SystemModal)
    End Sub

    Private Sub 選択範囲のみToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 選択範囲のみToolStripMenuItem.Click
        Dim str As String = HTMLTIDY(az.GetSelectedText)

        Dim start As Integer = 0
        Dim endS As Integer = 0
        If az.Document.AnchorIndex < az.CaretIndex Then
            start = az.CaretIndex - az.GetSelectedText.Length
            endS = az.CaretIndex
        Else
            start = az.CaretIndex
            endS = az.CaretIndex + az.GetSelectedText.Length
        End If
        az.Document.Replace(str)
        az.Document.SetSelection(start, str.Length)

        MsgBox("終了しました", MsgBoxStyle.SystemModal)
    End Sub


    '行番号を表示する





    Private Sub HTMLcreate_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        Label1.Text = "H"
        Form1.CSV_FORMS(0).ToolStripStatusLabel5.Text = "H"
        Form1.CSV_FORMS(1).ToolStripStatusLabel5.Text = "H"
    End Sub
End Class



