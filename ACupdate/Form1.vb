Option Explicit On

Imports Microsoft.VisualBasic.FileIO
Imports System.IO
Imports System.IO.Path
Imports System.Reflection
Imports System.Text.RegularExpressions

Public Class Form1
    Public Shared Assembly_name As String = "AutoCheckUpdate"
    Public Shared dirPath As String         '設定保存用ディレクトリ
    Public Shared DataDirPath As String     'データ保存用ディレクトリ
    Public Shared UpdateTXT As String() = Nothing

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Text = "アップデートをチェックします Ver." & Application.ProductVersion

        'プログラムのディレクトリを取得する
        Dim asm As Assembly = Assembly.GetEntryAssembly()
        Dim fullpass As String = asm.Location
        dirPath = Path.GetDirectoryName(fullpass)

        Dim dirLoadFile As String = dirPath & "\config\folder.txt"
        Dim flag As Integer = 0
        If File.Exists(dirLoadFile) Then
            Using parser As New TextFieldParser(dirLoadFile, System.Text.Encoding.GetEncoding("Shift_JIS"))
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

        Dim UpdatePath As String = DataDirPath & "\update\version.txt"
        If File.Exists(UpdatePath) Then
            UpdateTXT = File.ReadAllLines(UpdatePath, System.Text.Encoding.GetEncoding("Shift_JIS"))
        Else
            MessageBox.Show(UpdatePath & vbCrLf & "バージョンファイルがありません", Assembly_name, MessageBoxButtons.OK)
            End
        End If

        ListBox1.Items.Add("アップデート処理開始")
        ListBox1.Items.Add("コピー元：" & DataDirPath)
        ListBox1.Items.Add("コピー先：" & dirPath)
        ListBox1.Items.Add("コピーを開始します。")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1

        'Timer1.Enabled = True
        Run()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        'If ProgressBar1.Value <> 5 Then
        '    System.Threading.Thread.Sleep(500)
        '    ProgressBar1.Value += 1
        'Else
        '    Run()
        'End If
    End Sub

    Private Sub Run()
        'Timer1.Enabled = False

        ProgressBar1.Maximum = UpdateTXT.Length
        ProgressBar1.Value = 0
        ProgressBar2.Maximum = 1

        Dim copySrc As String = ""
        Dim copyDest As String = ""
        For i As Integer = 0 To UpdateTXT.Length - 1
            If InStr(UpdateTXT(i), "[") = 0 Then
                ProgressBar2.Value = 0
                ListBox1.Items.Add("「 " & UpdateTXT(i) & " 」をコピーします。")
                ListBox1.SelectedIndex = ListBox1.Items.Count - 1
                copySrc = DataDirPath & "\update\" & UpdateTXT(i)
                copyDest = dirPath & "\" & UpdateTXT(i)
                ListBox1.SelectedIndex = ListBox1.Items.Count - 1
                If File.Exists(copyDest) = True Then
                    IO.File.Copy(copySrc, copyDest, True)
                Else
                    My.Computer.FileSystem.CopyFile(copySrc, copyDest, True)
                End If
                'IO.File.Copy(copySrc, copyDest, True)
                'My.Computer.FileSystem.CopyFile(copySrc, copyDest, True)
                ListBox1.Items.Add("「 " & UpdateTXT(i) & " 」をコピーしました。")
                ListBox1.Items.Add("・・・")
                ListBox1.SelectedIndex = ListBox1.Items.Count - 1
                ProgressBar1.Value += 1
                ProgressBar2.Value = 1
            ElseIf instr(UpdateTXT(i), "[del]") > 0 Then
                Dim findPath As String = Replace(UpdateTXT(i), "[del]", "")
                ProgressBar2.Value = 0
                If Not File.Exists(findPath) Then
                    ListBox1.Items.Add("「 " & UpdateTXT(i) & " 」を削除します。")
                    ListBox1.SelectedIndex = ListBox1.Items.Count - 1
                    File.Delete(findPath)
                    ListBox1.Items.Add("「 " & UpdateTXT(i) & " 」を削除しました。")
                    ListBox1.Items.Add("・・・")
                    ListBox1.SelectedIndex = ListBox1.Items.Count - 1
                    ProgressBar1.Value += 1
                    ProgressBar2.Value = 1
                End If
            End If
        Next i

        ProgressBar1.Value = 0
        ProgressBar2.Value = 0

        ListBox1.Items.Add("アップデートが終了しました。")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        Button1.Enabled = True
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Ret As Long = Shell(dirPath & "\AutoCheck1.exe", vbNormalFocus)
        End
    End Sub

End Class
