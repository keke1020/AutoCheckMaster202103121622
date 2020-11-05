Option Explicit On

Imports Microsoft.VisualBasic.FileIO
Imports System.IO
Imports System.IO.Path
Imports System.Reflection
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading

Public Class UpdatePanel
    Public Shared Assembly_name As String = "AutoCheckUpdate"
    Public Shared dirPath As String         '設定保存用ディレクトリ
    Public Shared DataDirPath As String     'データ保存用ディレクトリ
    Public Shared TempDirPath As String     'マイドキュメント保存ディレクトリ
    Public Shared UpdateTXT As String() = Nothing

    Dim motoPath As String = Path.GetDirectoryName(Form1.appPath) & "\AutoCheck1.exe"
    'Dim sakiPath As String = Path.GetDirectoryName(Form1.appPath) & "\AutoCheck1_bak.exe"

    Private Sub Update_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Text = "アップデートをチェックします Ver." & Application.ProductVersion

        'プログラムのディレクトリを取得する
        Dim asm As Assembly = Assembly.GetEntryAssembly()
        Dim fullpass As String = asm.Location
        dirPath = Path.GetDirectoryName(fullpass)

        Dim dirLoadFile As String = dirPath & "\config\folder.txt"
        Dim flag As Integer = 0
        If File.Exists(dirLoadFile) Then
            Using parser As New TextFieldParser(dirLoadFile, Encoding.GetEncoding("Shift_JIS"))
                While Not parser.EndOfData
                    Dim fields As String = parser.ReadLine
                    If fields = "[Data Folder]" Then
                        flag = 1
                    ElseIf flag = 1 And fields <> "" Then
                        DataDirPath = fields
                        flag = 2
                    End If
                End While
            End Using
        Else
            MessageBox.Show(dirLoadFile & vbCrLf & "設定ディレクトリ保存ファイルがありません", Assembly_name, MessageBoxButtons.OK)
            End
        End If

        If Regex.IsMatch(My.Computer.Name, "takashi", RegexOptions.IgnoreCase) Then
            DataDirPath = SpecialDirectories.MyDocuments & "\test"
        End If
        'DataDirPath = Form1.サーバーToolStripMenuItem.Text
        TempDirPath = SpecialDirectories.MyDocuments

        Dim UpdatePath As String = DataDirPath & "\update\version.txt"
        If File.Exists(UpdatePath) Then
            UpdateTXT = File.ReadAllLines(UpdatePath, Encoding.GetEncoding("Shift_JIS"))
        Else
            MessageBox.Show(UpdatePath & vbCrLf & "バージョンファイルがありません", Assembly_name, MessageBoxButtons.OK)
            End
        End If

        Timer1.Enabled = True
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False
        Application.DoEvents()

        Dim serverVer As String() = File.ReadAllLines(DataDirPath & "\update\version.txt", Encoding.GetEncoding("Shift_JIS"))
        Label1.Text = Regex.Replace(serverVer(0), "\[|\]", "")
        Dim tempVer As String() = File.ReadAllLines(SpecialDirectories.MyDocuments & "\taksoft\AutoCheck\version.txt", Encoding.GetEncoding("Shift_JIS"))
        Label2.Text = Regex.Replace(tempVer(0), "\[|\]", "")
        Dim localVer As String() = File.ReadAllLines(Path.GetDirectoryName(Form1.appPath) & "\config\version.txt", Encoding.GetEncoding("Shift_JIS"))
        Label3.Text = Regex.Replace(localVer(0), "\[|\]", "")

        If Label1.Text <> "" And Label2.Text <> "" Then
            If Label1.Text = Label2.Text Then
                Label2.BackColor = Color.DarkOrange
            Else
                Label1.BackColor = Color.DarkOrange
            End If
        Else
            Label1.BackColor = Color.DarkOrange
        End If

        'exeのバックアップがあったら先に消す（180530ver）
        'sakiPath = Path.GetDirectoryName(Form1.appPath) & "\AutoCheck1_bak.exe"
        Dim bFiles As String() = Directory.GetFiles(Path.GetDirectoryName(Form1.appPath))
        For Each bF As String In bFiles
            If File.Exists(bF) And Regex.IsMatch(bF, "AutoCheck1_bak") Then
                Try
                    My.Computer.FileSystem.DeleteFile(bF)
                Catch ex As Exception

                End Try
            End If
        Next

        Dim newPath As String = dirPath & "\AutoCheck1_bak_" & Format(Now(), "yyMMddHHmmss") & ".exe"
        File.Move(motoPath, newPath)   'exeファイルをリネームしておく

        ListBox1.Items.Add("アップデート処理開始")
        If Label1.BackColor = Color.DarkOrange Then
            ListBox1.Items.Add("コピー元：" & DataDirPath)
        Else
            ListBox1.Items.Add("コピー元：" & TempDirPath)
        End If
        ListBox1.Items.Add("コピー先：" & dirPath)
        ListBox1.Items.Add("コピーを開始します。")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        Application.DoEvents()

        Threading.Thread.Sleep(1000)

        ProgressBar1.Maximum = UpdateTXT.Length
        ProgressBar1.Value = 0
        ProgressBar2.Maximum = 1

        Dim copySrc As String = ""
        Dim copyDest As String = ""
        For i As Integer = 0 To UpdateTXT.Length - 1
            If Regex.IsMatch(UpdateTXT(i), "\[[0-9]+\]") Then
                'コピー対象ではない行
            ElseIf InStr(UpdateTXT(i), "[del]") > 0 Then
                Dim findPath As String = Replace(UpdateTXT(i), "[del]", "")
                findPath = dirPath & "\" & findPath
                ProgressBar2.Value = 0
                If File.Exists(findPath) Then
                    ListBox1.Items.Add("「 " & UpdateTXT(i) & " 」を削除します。")
                    ListBox1.SelectedIndex = ListBox1.Items.Count - 1
                    File.Delete(findPath)
                    ListBox1.Items.Add("「 " & UpdateTXT(i) & " 」を削除しました。")
                    ListBox1.Items.Add("・・・")
                    ListBox1.SelectedIndex = ListBox1.Items.Count - 1
                    ProgressBar1.Value += 1
                    ProgressBar2.Value = 1
                End If
            Else
                ProgressBar2.Value = 0
                UpdateTXT(i) = Replace(UpdateTXT(i), vbCrLf, "")
                UpdateTXT(i) = Replace(UpdateTXT(i), vbCr, "")
                UpdateTXT(i) = Replace(UpdateTXT(i), vbLf, "")
                ListBox1.Items.Add("「 " & UpdateTXT(i) & " 」をコピーします。")
                ListBox1.SelectedIndex = ListBox1.Items.Count - 1
                If Label1.BackColor = Color.DarkOrange Then
                    copySrc = DataDirPath & "\update\" & UpdateTXT(i)
                Else
                    copySrc = TempDirPath & "\taksoft\AutoCheck\" & UpdateTXT(i)
                End If
                copyDest = dirPath & "\" & UpdateTXT(i)
                copySrc = Replace(copySrc, "[all]", "")
                copyDest = Replace(copyDest, "[all]", "")
                ListBox1.SelectedIndex = ListBox1.Items.Count - 1

                Try
                    'mdb等を更新させる時は[all]を付ける
                    Select Case True
                        Case Regex.IsMatch(Path.GetExtension(copyDest), "mdb|accdb|dll")
                            If InStr(UpdateTXT(i), "[all]") > 0 Then
                                If File.Exists(copyDest) = True Then
                                    File.Copy(copySrc, copyDest, True)
                                Else
                                    My.Computer.FileSystem.CopyFile(copySrc, copyDest, True)
                                End If
                            End If
                        Case Else
                            If File.Exists(copyDest) = True Then
                                File.Copy(copySrc, copyDest, True)
                            Else
                                My.Computer.FileSystem.CopyFile(copySrc, copyDest, True)
                            End If
                    End Select
                Catch ex As Exception

                End Try

                'IO.File.Copy(copySrc, copyDest, True)
                'My.Computer.FileSystem.CopyFile(copySrc, copyDest, True)
                ListBox1.Items.Add("「 " & UpdateTXT(i) & " 」をコピーしました。")
                ListBox1.Items.Add("・・・")
                ListBox1.SelectedIndex = ListBox1.Items.Count - 1
                ProgressBar1.Value += 1
                ProgressBar2.Value = 1
            End If

            Application.DoEvents()
        Next i

        ProgressBar1.Value = 0
        ProgressBar2.Value = 0

        'If File.Exists(sakiPath) Then   'exeのバックアップを消す
        '    File.Delete(sakiPath)
        'End If

        ListBox1.Items.Add("ショートカットの確認中")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        Application.DoEvents()

        ShortcutCheck()

        ListBox1.Items.Add("ショートカットの確認終了")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        Application.DoEvents()

        ListBox1.Items.Add("アップデートが終了しました。")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        Button1.Enabled = True

        Application.DoEvents()
        Threading.Thread.Sleep(1000)
        'CloseForm()

        Dim appPath As String = Reflection.Assembly.GetExecutingAssembly().Location
        Process.Start(Path.GetDirectoryName(appPath) & "\AutoCheck3.exe")

        'If Form1.NotifyIcon1.Visible Then
        '    Form1.NotifyIcon1.Dispose()
        'End If

        'MsgBox("test")

        Me.Close()
        Me.Dispose()
        Application.Exit()
        'End

    End Sub

    Public Sub ShortcutCheck()
        Dim WshShell As Object = CreateObject("WScript.Shell")
        Dim appPath As String = Form1.appPath
        Dim sp As String() = New String() {Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"}

        Dim sCount As Integer = 0
        For i As Integer = 0 To 1
            Dim spPath As String = sp(i)
            For Each stFilePath As String In Directory.GetFiles(spPath, "*.lnk")
                If InStr(stFilePath, "AutoCheck1") > 0 Then
                    Dim ShellLink As Object = WshShell.CreateShortcut(stFilePath)
                    Dim appLink As String = ShellLink.TargetPath
                    If appLink <> appPath Then
                        ShellLink.TargetPath = """" & appPath & """"
                        ShellLink.WorkingDirectory = """" & Path.GetDirectoryName(appPath) & """"
                        ShellLink.Save

                        ListBox1.Items.Add("ショートカットを修理しました")
                        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
                        sCount += 1
                        Application.DoEvents()
                    End If
                End If
            Next stFilePath
        Next

        If sCount = 0 Then
            ListBox1.Items.Add("ショートカットを修理：0個")
            ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        Else
            ListBox1.Items.Add("ショートカットを修理：" & sCount & "個")
            ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Dim Ret As Long = Shell(dirPath & "\AutoCheck1.exe", vbNormalFocus)
        'End
        If Form1.NotifyIcon1.Visible Then
            Form1.NotifyIcon1.Dispose()
        End If

        End


        'Form1.再起動ToolStripMenuItem.PerformClick()
    End Sub

    Private Sub CloseForm()
        If Form1.NotifyIcon1.Visible Then
            Form1.NotifyIcon1.Dispose()
        End If

        'Dim frm As Form = New Form1
        'Me.Dispose()
        'frm.Show()

        'Application.Restart()

        Dim appPath As String = Reflection.Assembly.GetExecutingAssembly().Location
        Process.Start(Path.GetDirectoryName(appPath) & "\AutoCheck3.exe")

        'MsgBox("test")

        Application.Exit()
        'End


        'For i As Integer = Form1.TabBrowser1.TabPages.Count - 1 To 0
        '    Form1.TabBrowser1.TabPages.RemoveAt(i)
        'Next

        'Dim ops As FormCollection = Application.OpenForms
        'For i As Integer = ops.Count - 1 To 0 Step -1
        '    Try
        '        Select Case False
        '            Case InStr(ops(i).Name, "Form1")
        '                ops(i).Close()
        '                ops(i).Dispose()
        '        End Select
        '    Catch ex As Exception

        '    End Try
        'Next

        'If Form1.NotifyIcon1.Visible Then
        '    Form1.NotifyIcon1.Dispose()
        'End If

        'Application.Restart()
    End Sub

End Class
