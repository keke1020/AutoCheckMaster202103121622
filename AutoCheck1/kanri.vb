Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports Hnx8

Public Class Kanri
    Dim secValue As String = ""

    Private CurrentDirectoryM As DirectoryInfo
    Private CurrentDirectoryS As DirectoryInfo
    Dim motoPath As String = ""
    Dim sakiPath As String = ""

    Dim localTimePath As String = appPathDir & "\db\sagawa.dat"
    Dim localPath As String = appPathDir & "\db\sagawa.accdb"
    Dim serverTimePath As String = Form1.サーバーToolStripMenuItem.Text & "\update\db\sagawa.dat"
    Dim serverPath As String = Form1.サーバーToolStripMenuItem.Text & "\update\db\sagawa.accdb"

    Dim txtPathArray As New ArrayList

    Private Sub Kanri_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        secValue = "kanri"   'iniファイルセクション

        txtPathArray.Add("伝票用住所,伝票用住所.txt")
        txtPathArray.Add("会社情報,会社情報.txt")
        txtPathArray.Add("PC番号,config\PC番号.dat")
        txtPathArray.Add("bookmark,bookmark.txt")
        For Each tA As String In txtPathArray
            Dim tAstr As String = Split(tA, ",")(0)
            ListBox3.Items.Add(tAstr)
        Next
        ListBox3.SelectedIndex = 0

        motoPath = Path.GetDirectoryName(Form1.appPath)
        Connect()

        Form1.Timer1.Enabled = False

        'sagawa.accdb
        TextBox6.Text = File.ReadAllText(localTimePath, encSJ)
        TextBox7.Text = File.ReadAllText(serverTimePath, encSJ)
    End Sub

    Private Sub Kanri_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Form1.Timer1.Enabled = True
    End Sub

    Private Sub MFolder()
        DGV1.Rows.Clear()

        Dim CD1 As DirectoryInfo = New DirectoryInfo(motoPath)
        Dim d As DirectoryInfo
        For Each d In CD1.GetDirectories()
            DGV1.Rows.Add("[" & d.Name & "]", d.LastWriteTime.ToString)
            FileDGVinsert(DGV1, motoPath & "\" & d.Name)
        Next d
        Dim f As FileInfo
        For Each f In CD1.GetFiles()
            DGV1.Rows.Add(">" & f.Name, f.LastWriteTime.ToString)
        Next f
    End Sub

    Private Sub FileDGVinsert(DGVX As DataGridView, xPath As String)
        Dim CD2 As DirectoryInfo = New DirectoryInfo(xPath)
        Dim d As DirectoryInfo
        For Each d In CD2.GetDirectories()
            DGVX.Rows.Add(">[" & d.Name & "]", d.LastWriteTime.ToString)
        Next d
        Dim f As FileInfo
        For Each f In CD2.GetFiles()
            DGVX.Rows.Add(">>" & f.Name, f.LastWriteTime.ToString)
        Next f
    End Sub

    Private Sub SFolder()
        DGV2.Rows.Clear()

        Dim CD1 As DirectoryInfo = New DirectoryInfo(sakiPath)
        'Dim CD1 As DirectoryInfo = New DirectoryInfo("E:\Documents\test\update")
        Dim d As DirectoryInfo
        For Each d In CD1.GetDirectories()
            DGV2.Rows.Add("[" & d.Name & "]", d.LastWriteTime.ToString)
            FileDGVinsert(DGV2, sakiPath & "\" & d.Name)
        Next d
        Dim f As FileInfo
        For Each f In CD1.GetFiles()
            DGV2.Rows.Add(">" & f.Name, f.LastWriteTime.ToString)
        Next f
    End Sub

    '/ ディレクトリ用のListViewItem
    Private Class DirectoryItem
        Inherits System.Windows.Forms.ListViewItem
        Public Directory As DirectoryInfo
        Public Sub New(ByVal Directory As DirectoryInfo)
            Me.Directory = Directory
            Me.Text = Directory.Name
            Me.SubItems.Add("<dir>")
            Me.ImageIndex = 0
        End Sub 'New
    End Class 'DirectoryItem

    '/ ファイル用のListViewItem
    Private Class FileItem
        Inherits System.Windows.Forms.ListViewItem
        Public File As FileInfo
        Public Sub New(ByVal File As FileInfo)
            Me.File = File
            Me.Text = File.Name
            Me.SubItems.Add(File.Length.ToString("#,##0"))
            Me.ImageIndex = 1
        End Sub 'New
    End Class 'FileItem 

    Private Sub ToolStripDropDownButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripDropDownButton1.Click
        If AzukiControl3.Document.LineCount > 1 Then
            Dim fName As String = TextBox1.Text
            SaveTxt(fName)

            'ファイルサーバー\update\version\000010.txt ファイル名を更新する
            'ファイルサーバーでは更新日時を取得できないため
            Dim newName As String = ""
            newName = AzukiControl3.Document.GetLineContent(0)
            newName = Replace(newName, "[", "")
            newName = Replace(newName, "]", "")
            Dim di As New DirectoryInfo(sakiPath & "\version")
            Dim files As FileInfo() = di.GetFiles("*.txt", SearchOption.AllDirectories)
            For Each f As FileInfo In files
                FileSystem.Rename(f.FullName, sakiPath & "\version\" & newName & ".txt")
            Next

            MsgBox("保存しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    Public Sub SaveTxt(ByVal fp As String)
        If fp <> "" Then
            Dim sw As StreamWriter = New StreamWriter(fp, False, System.Text.Encoding.GetEncoding("SHIFT-JIS"))
            sw.Write(AzukiControl3.Text)
            sw.Close()
        End If
    End Sub

    Private Sub Mdb以外ファイル更新ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Mdb以外ファイル更新ToolStripMenuItem.Click
        UpdateFile(2)
    End Sub

    Private Sub ファイル更新ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ファイル更新ToolStripMenuItem.Click
        UpdateFile(0)
    End Sub

    Private Sub 強制ファイル更新ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 強制ファイル更新ToolStripMenuItem.Click
        UpdateFile(1)
    End Sub

    Private Sub UpdateFile(ByVal mode As Integer)
        Dim UpdateTxt As String() = Split(AzukiControl3.Text, vbLf)

        Dim copySrc As String = ""
        Dim copyDest As String = ""
        For i As Integer = 0 To UpdateTxt.Length - 1
            ToolStripProgressBar1.Value = 0
            ToolStripProgressBar1.Maximum = 4

            UpdateTxt(i) = Replace(UpdateTxt(i), vbCrLf, "")
            UpdateTxt(i) = Replace(UpdateTxt(i), vbCr, "")
            UpdateTxt(i) = Replace(UpdateTxt(i), vbLf, "")
            UpdateTxt(i) = Replace(UpdateTxt(i), "[all]", "")
            If UpdateTxt(i) <> "" And InStr(UpdateTxt(i), "[") = 0 Then
                ListBox1.Items.Add("「 " & UpdateTxt(i) & " 」をコピーします。")
                ListBox1.SelectedIndex = ListBox1.Items.Count - 1
                copySrc = motoPath & "\" & UpdateTxt(i)

                'サーバーにデータ保存
                copyDest = sakiPath & "\" & UpdateTxt(i)
                ListBox1.SelectedIndex = ListBox1.Items.Count - 1
                If File.Exists(copyDest) Then
                    If InStr(copyDest, "AmazonNode") > 0 Then
                        Dim a = ""
                    End If

                    If Regex.IsMatch(Path.GetExtension(copySrc), "exe$|pdb$|dll$|vb$|resx$|config$|ini$|xml$") Then
                        File.Copy(copySrc, copyDest, True)
                    Else
                        If mode = 2 Then
                            If Not Regex.IsMatch(Path.GetExtension(copySrc), "mdb$|accdb$") Then
                                File.Copy(copySrc, copyDest, True)
                            End If
                        ElseIf mode = 1 Then
                            File.Copy(copySrc, copyDest, True)
                        Else
                            If Regex.IsMatch(Path.GetExtension(copySrc), "dat$|txt$|csv$|xlsx$|xls$|pdf$|html$|bat$|jpg$") Then
                                File.Copy(copySrc, copyDest, True)
                            End If
                        End If
                    End If
                Else
                    My.Computer.FileSystem.CopyFile(copySrc, copyDest, True)
                End If

                ToolStripProgressBar1.Value += 1
                Application.DoEvents()

                ListBox1.Items.Add("「 " & UpdateTxt(i) & " 」をコピーしました。")
                ListBox1.Items.Add("・・・")
                ListBox1.SelectedIndex = ListBox1.Items.Count - 1
                If i <= 2 Then
                    Application.DoEvents()
                End If
            End If
        Next i
        SFolder()

        ToolStripProgressBar1.Value = 0
        ListBox1.Items.Add("アップデートが終了しました。")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
    End Sub

    Private Sub ファイルサーバー用ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ファイルサーバー用ToolStripMenuItem.Click
        ファイルサーバー用ToolStripMenuItem.Checked = True
        デバッグ用ToolStripMenuItem.Checked = False
        'sakiPath = "Z:\中村作成データ\AutoCheck\update"
        'sakiPath = "\\SERVER-PC\Users\Public\program\autocheck\update"
        sakiPath = Form1.サーバーToolStripMenuItem.Text & "\update"
        Connect()
    End Sub

    Private Sub デバッグ用ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles デバッグ用ToolStripMenuItem.Click
        ファイルサーバー用ToolStripMenuItem.Checked = False
        デバッグ用ToolStripMenuItem.Checked = True
        sakiPath = "E:\ライブラリ\マイ ドキュメント\test\update"
        Connect()
    End Sub

    Private Sub Connect()
        'Try
        'ファイル読み取り
        Dim fName As String = sakiPath & "\version.txt"
            TextBox1.Text = fName
            Dim sr As New StreamReader(fName, System.Text.Encoding.GetEncoding("shift_jis"))
            AzukiControl3.Text = sr.ReadToEnd()
            sr.Close()

            MFolder()
            SFolder()
            TextBox4.Text = motoPath
            TextBox5.Text = sakiPath

            '時間比較
            For r1 As Integer = 0 To DGV1.RowCount - 1
                If InStr(DGV1.Item(0, r1).Value, "[") = 0 Then
                    For r2 As Integer = 0 To DGV2.RowCount - 1
                        If DGV1.Item(0, r1).Value = DGV2.Item(0, r2).Value Then
                            If CDate(DGV1.Item(1, r1).Value) > CDate(DGV2.Item(1, r2).Value) Then
                                DGV1.Item(1, r1).Style.BackColor = Color.Yellow
                            Else
                                DGV2.Item(1, r2).Style.BackColor = Color.Yellow
                            End If
                            Exit For
                        End If
                    Next
                End If
            Next

            Dim ver As String = AzukiControl3.Document.GetLineContent(0)
            ver = Regex.Replace(ver, "\[|\]", "")
            ToolStripTextBox3.Text = ver

            Dim di As New DirectoryInfo(sakiPath & "\version")
            Dim files2 As FileInfo() = di.GetFiles("*.dat", SearchOption.AllDirectories)
            For Each f As FileInfo In files2
                ToolStripStatusLabel2.Text = Path.GetFileNameWithoutExtension(f.Name)
            Next
            Dim files3 As FileInfo() = di.GetFiles("*.ini", SearchOption.AllDirectories)
            For Each f As FileInfo In files3
                ToolStripStatusLabel4.Text = Replace(Path.GetFileNameWithoutExtension(f.Name), ".", "")
            Next

            '更新ファイルの更新時間チェック
            Dim aStr As String() = Split(AzukiControl3.Text, vbCrLf)
            For i As Integer = 0 To aStr.Length - 1
                If aStr(i) <> "" Then
                    For r1 As Integer = 0 To DGV1.RowCount - 1
                        Dim cc As Color = Color.LightBlue
                        Dim target As String = ""
                        If InStr(aStr(i), "[del]") > 0 Then
                            aStr(i) = Replace(aStr(i), "[del]", "")
                            cc = Color.Orange
                        ElseIf InStr(aStr(i), "[all]") > 0 Then
                            aStr(i) = Replace(aStr(i), "[all]", "")
                            cc = Color.Yellow
                        End If
                        If InStr(aStr(i), "\") > 0 Then
                            Dim ar As String() = Split(aStr(i), "\")
                            target = ar(ar.Length - 1)
                        Else
                            target = aStr(i)
                        End If
                        If InStr(DGV1.Item(0, r1).Value, target) > 0 Then
                            DGV1.Item(0, r1).Style.BackColor = cc
                        End If
                    Next
                End If
            Next

            ToolStripTextBox1.Text = Replace(Application.ProductVersion, ".", "")
        'Catch ex As Exception

        'End Try
    End Sub

    Private Sub ToolStripTextBox2_Click(sender As Object, e As EventArgs) Handles ToolStripTextBox2.Click
        If ToolStripTextBox2.Checked = True Then
            ToolStripTextBox2.Text = "OFF"
        Else
            ToolStripTextBox2.Text = "ON"
        End If
    End Sub

    Private Sub ファイルサーバー接続ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ファイルサーバー接続ToolStripMenuItem.Click
        '*********************************************
        Dim iniStr1 As String = Form1.inif.ReadINI(secValue, "lpRemoteName", "")
        Dim iniStr2 As String = Form1.inif.ReadINI(secValue, "lpUser", "")
        Dim iniStr3 As String = Form1.inif.ReadINI(secValue, "lpPass", "")
        Dim iniStr4 As String = Form1.inif.ReadINI(secValue, "lpLocalName", "")
        '*********************************************
        If iniStr1 = "" Then
            Dim IR As String = InputBox("WevDavのパスを入力してください")
            iniStr1 = IR
            '*********************************************
            Form1.inif.WriteINI(secValue, "lpRemoteName", IR)
            '*********************************************
        End If
        If iniStr2 = "" Then
            Dim IR As String = InputBox("ユーザー名を入力してください")
            iniStr2 = IR
            '*********************************************
            Form1.inif.WriteINI(secValue, "lpUser", IR)
            '*********************************************
        End If
        If iniStr3 = "" Then
            Dim IR As String = InputBox("パスワードを入力してください")
            iniStr3 = IR
            '*********************************************
            Form1.inif.WriteINI(secValue, "lpPass", IR)
            '*********************************************
        End If
        If iniStr4 = "" Then
            Dim IR As String = InputBox("ローカルのドライブ名を入力してください。例 Z")
            iniStr4 = IR
            '*********************************************
            Form1.inif.WriteINI(secValue, "lpLocalName", IR)
            '*********************************************
        End If

        Dim result As Integer
        Dim myResource As NETRESOURCE

        myResource.dwScope = 2
        myResource.dwType = 1
        myResource.dwDisplayType = 0
        myResource.dwUsage = Nothing
        myResource.lpComment = Nothing
        myResource.lpLocalName = iniStr4 & ":" 'ネットワークドライブ名
        myResource.lpProvider = Nothing
        myResource.lpRemoteName = iniStr1

        result = WNetAddConnection2(myResource, iniStr2, iniStr3, 0)
        MsgBox(result.ToString(), MsgBoxStyle.SystemModal)
    End Sub

    Private Sub 切断ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 切断ToolStripMenuItem.Click
        Dim result As Boolean
        result = WNetCancelConnection("Z:", False)
        MsgBox(result.ToString, MsgBoxStyle.SystemModal)
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        'ファイルサーバー\update\version\23660.dat ファイル名を更新する
        'ファイルサーバーでは更新日時を取得できないため
        'Dim newName As String = ""
        'newName = Application.ProductVersion
        'newName = Replace(newName, ".", "")
        Dim di As New DirectoryInfo(sakiPath & "\version")
        Dim files As FileInfo() = di.GetFiles("*.dat", SearchOption.AllDirectories)
        For Each f As FileInfo In files
            FileSystem.Rename(f.FullName, sakiPath & "\version\" & ToolStripTextBox1.Text & ".dat")
        Next

        If ToolStripTextBox2.Text = "ON" Then
            Dim files3 As FileInfo() = di.GetFiles("*.ini", SearchOption.AllDirectories)
            For Each f As FileInfo In files3
                FileSystem.Rename(f.FullName, sakiPath & "\version\" & ToolStripTextBox1.Text & ".ini")
            Next
        End If

        MsgBox("保存しました", MsgBoxStyle.SystemModal)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Test.Show()
        'FormFront(PrintData)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim mes As String = ""

        Dim str2 As String = "abcd180510235930"
        Dim base As Integer = 32
        mes &= Chr(18 + base) & Chr(5 + base) & Chr(10 + base)
        mes &= Chr(23 + base) & Chr(59 + base) & Chr(30 + base) & vbCrLf

        mes &= vbCrLf

        Dim str As String = "abcd180510152230"
        mes &= str & vbCrLf
        mes &= str.Length & vbCrLf
        Dim c1 As String = ""
        For i As Integer = 0 To str.Length - 1
            c1 &= Asc(str(i))
            mes &= str(i) & ","
        Next
        mes &= vbCrLf & c1.Length & vbCrLf
        mes &= c1 & vbCrLf
        Dim c2 As String = ""
        For i As Integer = 0 To c1.Length / 4 - 1
            Dim c3 As Integer = c1.Substring(i * 4, 4)
            c2 &= RadixConvert.ToString(c3, 32, True)
            mes &= c3 & ","
        Next
        mes &= vbCrLf & c2.Length & vbCrLf
        mes &= c2 & vbCrLf
        MsgBox(mes)
    End Sub

    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged
        Select Case ListBox2.SelectedItem
            Case "taskkill"
                TextBox3.Text = "taskkill -f -t /pid 400"
                Clipboard.SetText(TextBox3.Text)
                TextBox3.Text &= vbCrLf & "taskkill -f -t /im excel.exe"
            Case "netstat"
                TextBox3.Text = "netstat -nao"
                Clipboard.SetText(TextBox3.Text)
        End Select
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ErrorDialog.Height = 145
        Throw New Exception
        'ErrorDialog.Show()
    End Sub

    Private Sub DGV1_SelectionChanged(sender As Object, e As EventArgs) Handles DGV1.SelectionChanged
        If DGV1.SelectedCells.Count = 0 Or DGV1.RowCount = 0 Or DGV2.RowCount = 0 Then
            Exit Sub
        End If

        For r As Integer = 0 To DGV2.RowCount - 1
            If DGV1.Item(0, DGV1.SelectedCells(0).RowIndex).Value = DGV2.Item(0, r).Value Then
                DGV2.CurrentCell = DGV2(0, r)
                DGV2.Item(0, r).Selected = True
            End If
        Next
    End Sub

    Private Sub DGV2_SelectionChanged(sender As Object, e As EventArgs) Handles DGV2.SelectionChanged
        If DGV1.RowCount = 0 Or DGV2.SelectedCells.Count = 0 Or DGV2.RowCount = 0 Then
            Exit Sub
        End If

        For r As Integer = 0 To DGV1.RowCount - 1
            If DGV2.Item(0, DGV2.SelectedCells(0).RowIndex).Value = DGV1.Item(0, r).Value Then
                DGV1.CurrentCell = DGV1(0, r)
                DGV1.Item(0, r).Selected = True
            End If
        Next
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Try
            Button5.Text = "更新中..."
            Button5.Enabled = False
            Application.DoEvents()
            File.Copy(localPath, serverPath, True)
            File.WriteAllText(serverTimePath, Format(Now, "yyyy/MM/dd HH:mm:00"))
            Button5.Text = "更 新"
            Button5.Enabled = True
            Application.DoEvents()
        Catch ex As Exception
            Button5.Text = "更 新"
            Button5.Enabled = True
            Application.DoEvents()
            MsgBox(ex.Message)
            Exit Sub
        End Try
        MsgBox("データベースを更新しました")
        TextBox6.Text = File.ReadAllText(localTimePath, encSJ)
        TextBox7.Text = File.ReadAllText(serverTimePath, encSJ)
    End Sub

    Private Sub ListBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox3.SelectedIndexChanged
        TextBox8.Text = Split(txtPathArray(ListBox3.SelectedIndex), ",")(1)
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim fName As String = Split(txtPathArray(ListBox3.SelectedIndex), ",")(1)
        Dim readPath As String = Path.GetDirectoryName(appPath) & "\" & fName

        Dim files As FileInfo = New FileInfo(readPath)
        Dim fileEnc As String = ""
        Using reader As ReadJEnc.FileReader = New ReadJEnc.FileReader(files)
            Dim c As ReadJEnc.CharCode = reader.Read(files)
            fileEnc = c.Name
        End Using
        If fileEnc = "ShiftJIS" Then
            fileEnc = "Shift-JIS"
        ElseIf fileEnc = "UTF-8N" Then
            fileEnc = "utf-8"
        End If
        ToolStripStatusLabel5.Text = fileEnc

        AzukiControl1.Text = File.ReadAllText(readPath, Encoding.GetEncoding(fileEnc))
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If AzukiControl1.Text = "" Then
            Exit Sub
        End If

        Dim fName As String = Split(txtPathArray(ListBox3.SelectedIndex), ",")(1)
        Dim savePath As String = Path.GetDirectoryName(appPath) & "\" & fName
        File.WriteAllText(savePath, AzukiControl1.Text, Encoding.GetEncoding(ToolStripStatusLabel5.Text))
        MsgBox("保存しました。サーバーには登録していません。")
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        Dim ver As String = AzukiControl3.Document.GetLineContent(0)
        Dim moto As String = ver
        ver = Regex.Replace(ver, "\[|\]", "")
        Dim verNew As String = CInt(ver) + 1
        verNew = verNew.PadLeft(6, "0")
        ToolStripTextBox3.Text = verNew
        AzukiControl3.Text = Replace(AzukiControl3.Text, moto, "[" & verNew & "]")
    End Sub
End Class

