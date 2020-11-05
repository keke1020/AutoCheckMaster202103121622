Imports System.ComponentModel
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

Public Class Form1_F_loginlist
    Dim ENC_UTF8 As Encoding = Encoding.UTF8
    Dim appPath As String = Path.GetDirectoryName(Application.ExecutablePath)
    Dim serverFile As String = Form1.サーバーToolStripMenuItem.Text & "\update\loginlist4.dat"
    Dim localFile As String = appPath & "\loginlist4.dat"
    Dim serverArray As String() = Nothing
    Dim localArray As String() = Nothing
    Dim pcNoArray As New ArrayList
    Dim maildealerArray As New ArrayList
    Dim nextEngineArray As New ArrayList
    Dim SaveUpdateArray As New ArrayList
    Dim browserLoginArray As New ArrayList

    Dim base64 As New MyBase64str("UTF-8")
    'Dim cnvStr As String = base64.Encode("Hello World こんにちは")
    'Dim rstStr As String = base64.Decode(cnvStr)

    Private Sub Form1_loginlist_Load(sender As Object, e As EventArgs) Handles Me.Load
        For c As Integer = 0 To DGV1.ColumnCount - 1
            DGV1.Columns(c).SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        If Not Form1.AdminFlag Then
            管理用出力ToolStripMenuItem.Visible = False
            管理用出力サーバーToolStripMenuItem.Visible = False
            管理用取込ToolStripMenuItem.Visible = False
        End If

        If Not File.Exists(serverFile) Then
            MsgBox("ファイルサーバーが接続されている時に起動してください")
            Exit Sub
        End If

        'ファイル読み込み
        'serverArray = File.ReadAllLines(serverFile, ENC)
        serverArray = Split(base64.Decode(File.ReadAllText(serverFile, ENC_UTF8)), vbCrLf)
        Label6.Text = "S:" & serverArray(0)

        'ファイル確認
        If Not File.Exists(localFile) Then
            File.Copy(serverFile, localFile, True)
        End If

        'localArray = File.ReadAllLines(localFile, ENC_UTF8)
        localArray = Split(base64.Decode(File.ReadAllText(localFile, ENC_UTF8)), vbCrLf)
        Label1.Text = "L:" & localArray(0)

        '更新確認
        Dim updateFlag As Boolean = False
        Dim readArray As String() = Nothing
        If serverArray.Length > 0 Then
            If CDate(serverArray(0)) > CDate(localArray(0)) Then
                readArray = serverArray
                updateFlag = True
                SaveUpdateArray.Add(Now.ToString)
            Else
                readArray = localArray
            End If
        Else
            readArray = localArray
        End If

        'リスト表示
        Dim listNo As Integer = 0
        Dim hkRowNo As Integer = 0  '発注管理用パスワード行
        Dim mdRowNo As Integer = 0  'メールディーラーパスワード行
        Dim neRowNo As Integer = 0  'ネクストエンジンパスワード行
        Dim brRowNo As Integer = 0  'ブラウザログイン行
        For i As Integer = 1 To readArray.Length - 1
            If readArray(i) = "" Then
                '*****************
                If updateFlag Then
                    SaveUpdateArray.Add("")
                End If
                '*****************
                Continue For
            End If

            Dim pass As String = ""
            If Regex.IsMatch(readArray(i), "^//") Then
                '*****************
                If updateFlag Then
                    SaveUpdateArray.Add(readArray(i))
                End If
                '*****************
                listNo += 1
            ElseIf Regex.IsMatch(readArray(i), "^\-") Then
                Dim line As String = Regex.Replace(readArray(i), "^\-", "")
                Dim lineArray As String() = Split(line, ",")
                If InStr(lineArray(0), "●") > 0 Then
                    '*****************
                    If updateFlag Then
                        SaveUpdateArray.Add(readArray(i))
                    End If
                    '*****************
                Else
                    '*****************
                    If updateFlag Then
                        '●が無い行はアップデートしない
                        line = Regex.Replace(localArray(i), "^\-", "")
                        lineArray = Split(line, ",")
                        SaveUpdateArray.Add(localArray(i))
                    End If
                    '*****************
                End If
                If lineArray.Length >= 3 Then
                    pass = lineArray(3)
                    'lineArray(3) = Regex.Replace(lineArray(3), ".{1}", "*")
                End If
                DGV1.Rows.Add(lineArray)
                DGV1.Rows(DGV1.RowCount - 1).DefaultCellStyle.BackColor = Color.LightCyan
                DGV1.Rows(DGV1.RowCount - 1).Visible = False

                If lineArray(1) = "発注管理個人別" Then
                    hkRowNo = DGV1.RowCount - 1
                ElseIf lineArray(1) = "メールディーラー" Then
                    mdRowNo = DGV1.RowCount - 1
                ElseIf lineArray(1) = "ネクストエンジン" Then
                    neRowNo = DGV1.RowCount - 1
                ElseIf lineArray(1) = "ブラウザログイン" Then
                    brRowNo = DGV1.RowCount - 1
                End If
            Else
                '*****************
                If updateFlag Then
                    SaveUpdateArray.Add(readArray(i))
                End If
                '*****************
                If listNo = 0 Then
                    Dim lineArray As String() = Split(readArray(i), ",")
                    DGV1.Rows.Add(lineArray)
                    DGV1.Rows(DGV1.RowCount - 1).DefaultCellStyle.BackColor = Color.LightBlue
                    DGV1.Rows(DGV1.RowCount - 1).ReadOnly = True
                ElseIf listNo = 1 Then
                    pcNoArray.Add(readArray(i))
                ElseIf listNo = 2 Then
                    maildealerArray.Add(readArray(i))
                ElseIf listNo = 3 Then
                    nextEngineArray.Add(readArray(i))
                ElseIf listNo = 4 Then
                    browserLoginArray.Add(readArray(i))
                End If
            End If
            If listNo = 0 Then
                DGV1.Item(3, DGV1.RowCount - 1).ToolTipText = pass
            End If
        Next

        Dim pcName As String = Form1.PC名ToolStripMenuItem.Text.ToLower

        'PC番号から発注管理読み出し
        For i As Integer = 0 To pcNoArray.Count - 1
            Dim pA As String = Split(pcNoArray(i), ",")(0).ToLower
            If pcName = pA Then
                DGV1.Item(2, hkRowNo).Value = Split(pcNoArray(i), ",")(1) & "さん"
                Dim pass As String = Split(pcNoArray(i), ",")(2)
                'DGV1.Item(3, hkRowNo).Value = Regex.Replace(pass, ".{1}", "*")
                DGV1.Item(3, hkRowNo).Value = pass
                Exit For
            End If
        Next

        'PC番号からネクストエンジン読み出し
        For i As Integer = 0 To nextEngineArray.Count - 1
            Dim pA As String = Split(nextEngineArray(i), ",")(0).ToLower
            If pcName = pA Then
                DGV1.Item(2, neRowNo).Value = Split(nextEngineArray(i), ",")(1)
                DGV1.Item(3, neRowNo).Value = Split(nextEngineArray(i), ",")(2)
                Exit For
            End If
        Next

        'PC番号からメールディーラー読み出し
        For i As Integer = 0 To maildealerArray.Count - 1
            Dim pA As String = Split(maildealerArray(i), ",")(0).ToLower
            If pcName = pA Then
                DGV1.Item(2, mdRowNo).Value = Split(maildealerArray(i), ",")(1)
                Dim pass As String = Split(maildealerArray(i), ",")(2)
                'DGV1.Item(3, mdRowNo).Value = Regex.Replace(pass, ".{1}", "*")
                DGV1.Item(3, mdRowNo).Value = pass
                Exit For
            End If
        Next

        'ブラウザログインから読み出し
        For i As Integer = 0 To browserLoginArray.Count - 1
            Dim br As String() = Split(browserLoginArray(i), "=")
            Select Case br(0)
                Case "楽天"
                    If ComboBox1.Items.Count > br(1) Then
                        ComboBox1.SelectedIndex = br(1)
                    Else
                        ComboBox1.SelectedIndex = 0
                    End If
                Case "Yahoo"
                    If ComboBox2.Items.Count > br(1) Then
                        ComboBox2.SelectedIndex = br(1)
                    Else
                        ComboBox2.SelectedIndex = 0
                    End If
            End Select
        Next

    End Sub

    Private Sub Form1_F_loginlist_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        'アップデートある時
        If SaveUpdateArray.Count > 0 Then
            Dim sArray As String() = DirectCast(SaveUpdateArray.ToArray(GetType(String)), String())
            File.WriteAllLines(localFile, sArray, ENC_UTF8)

            Label2.Text = "更新 " & Format(Now, "MM/dd HH:mm:ss")
            Label2.Visible = True
        End If
    End Sub

    Private Sub DGV1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV1.CellDoubleClick
        If DGV1.Item(0, e.RowIndex).Value = ">" Then
            For r As Integer = 0 To e.RowIndex
                If DGV1.Item(0, r).Value <> ">" Then
                    DGV1.Rows(r).Visible = False
                End If
            Next
            Dim listNo As Integer = 0
            For r As Integer = e.RowIndex + 1 To DGV1.RowCount - 1  'リストを開く
                If DGV1.Item(0, r).Value = ">" Then
                    listNo += 1
                Else
                    If listNo = 0 Then
                        If DGV1.Rows(r).Visible Then
                            DGV1.Rows(r).Visible = False
                        Else
                            DGV1.Rows(r).Visible = True
                        End If
                    Else
                        DGV1.Rows(r).Visible = False
                    End If
                End If
            Next
        Else
            'If e.ColumnIndex >= 2 Then
            '    If DGV1.Item(e.ColumnIndex, e.RowIndex).Value <> "" Then
            '        Dim str As String = DGV1.Item(e.ColumnIndex, e.RowIndex).Value
            '        Clipboard.SetDataObject(str, True)
            '        Me.Close()
            '    End If
            'End If
        End If
    End Sub

    '---------------------------------------------------
    'フォームを動かせるようにする
    'マウスのクリック位置を記憶
    Private mousePoint As Point

    'Form1のMouseDownイベントハンドラ
    'マウスのボタンが押されたとき
    Private Sub Form1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles _
            MyBase.MouseDown, DGV1.MouseDown, Panel1.MouseDown, Panel3.MouseDown, Label2.MouseDown
        If (e.Button And MouseButtons.Left) = MouseButtons.Left Then
            '位置を記憶する
            mousePoint = New Point(e.X, e.Y)
        End If
    End Sub

    'Form1のMouseMoveイベントハンドラ
    'マウスが動いたとき
    Private Sub Form1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles _
            MyBase.MouseMove, DGV1.MouseMove, Panel1.MouseMove, Panel3.MouseDown, Label2.MouseMove
        If (e.Button And MouseButtons.Left) = MouseButtons.Left Then
            Me.Left += e.X - mousePoint.X
            Me.Top += e.Y - mousePoint.Y
            'または、つぎのようにする
            'Me.Location = New Point( _
            '    Me.Location.X + e.X - mousePoint.X, _
            '    Me.Location.Y + e.Y - mousePoint.Y)
        End If
    End Sub
    '---------------------------------------------------

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Me.Close()
    End Sub

    'Private Sub 入力ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 入力ToolStripMenuItem.Click
    '    If DGV1.SelectedCells.Count > 0 Then
    '        Dim selCol As Integer = DGV1.SelectedCells(0).ColumnIndex
    '        Dim selRow As Integer = DGV1.SelectedCells(0).RowIndex
    '        If DGV1.Item(0, selRow).Value = ">" Then
    '            Exit Sub
    '        ElseIf selCol <= 1 Then
    '            Exit Sub
    '        ElseIf DGV1.Item(1, selrow).Value = "発注管理個人別" Then
    '            Exit Sub
    '        End If

    '        If selCol = 2 Then
    '            Dim txt As String = InputBox("IDまたはユーザー名入力", "入力")
    '            If txt <> "" Then
    '                DGV1.SelectedCells(0).Value = txt
    '                DGV1.SelectedCells(0).Style.BackColor = Color.Yellow
    '            End If
    '        ElseIf selCol = 3 Then
    '            Dim txt As String = InputBox("パスワード入力", "入力", DGV1.Item(3, selRow).Value)
    '            If txt <> "" Then
    '                'DGV1.Item(3, selRow).Value = Regex.Replace(txt, ".{1}", "*")
    '                DGV1.Item(3, selRow).Style.BackColor = Color.Yellow
    '                DGV1.Item(3, selRow).Value = txt
    '            End If
    '        End If
    '    End If
    'End Sub

    'Private Sub 削除ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 削除ToolStripMenuItem.Click
    '    If DGV1.SelectedCells.Count > 0 Then
    '        Dim selCol As Integer = DGV1.SelectedCells(0).ColumnIndex
    '        Dim selRow As Integer = DGV1.SelectedCells(0).RowIndex
    '        If DGV1.Item(0, selRow).Value = ">" Then
    '            Exit Sub
    '        ElseIf selCol <= 1 Then
    '            Exit Sub
    '        ElseIf DGV1.Item(1, selRow).Value = "発注管理個人別" Then
    '            Exit Sub
    '        End If

    '        DGV1.SelectedCells(0).Value = ""
    '        DGV1.SelectedCells(0).Style.BackColor = Color.Empty
    '    End If
    'End Sub

    'Private Sub パスワードを見るToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles パスワードを見るToolStripMenuItem.Click
    '    If DGV1.SelectedCells.Count > 0 Then
    '        Dim selCol As Integer = DGV1.SelectedCells(0).ColumnIndex
    '        Dim selRow As Integer = DGV1.SelectedCells(0).RowIndex
    '        If DGV1.Item(0, selRow).Value = ">" Then
    '            Exit Sub
    '        ElseIf selCol <= 1 Then
    '            Exit Sub
    '        ElseIf DGV1.Item(1, selRow).Value = "発注管理個人別" Then
    '            Exit Sub
    '        End If

    '        If selCol = 3 And DGV1.Item(5, selRow).Value <> "" Then
    '            MsgBox(DGV1.Item(5, selRow).Value)
    '        End If
    '    End If
    'End Sub

    Private Sub 保存ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 保存ToolStripMenuItem.Click
        listSave()
    End Sub

    Private Sub listSave()
        Dim savePath As New ArrayList
        savePath.Add(serverFile)
        savePath.Add(localFile)
        Dim serverArray As String() = Split(base64.Decode(File.ReadAllText(serverFile, ENC_UTF8)), vbCrLf)
        Dim localArray As String() = Split(base64.Decode(File.ReadAllText(localFile, ENC_UTF8)), vbCrLf)

        Dim nowTime As String = Now.ToString
        For i As Integer = 0 To savePath.Count - 1
            Dim saveStr As String = ""
            Dim readArray As String() = Split(base64.Decode(File.ReadAllText(savePath(i), ENC_UTF8)), vbCrLf)
            For k As Integer = 0 To readArray.Length - 1
                If k = 0 Then
                    saveStr = nowTime & vbCrLf
                Else
                    If readArray(k) <> "" Then
                        If Regex.IsMatch(readArray(k), "^-") Then
                            Dim dgvStr As String = ""
                            Dim dgvNokori As String = ""
                            Dim readStr As String = ""
                            Dim ra() As String = Split(readArray(k), ",")
                            For c As Integer = 0 To DGV1.ColumnCount - 1
                                If c = 0 Then
                                    dgvStr &= "-" & DGV1.Item(c, k - 1).Value & ","
                                ElseIf c <= 3 Then
                                    dgvStr &= DGV1.Item(c, k - 1).Value & ","
                                End If
                                If c >= 5 Then
                                    dgvNokori &= DGV1.Item(c, k - 1).Value & ","
                                End If
                                If c <= 3 Then
                                    readStr &= ra(c) & ","
                                End If
                            Next
                            'dgvStr = dgvStr.TrimEnd(",")
                            'readStr = readStr.TrimEnd(",")
                            'dgvNokori = dgvNokori.TrimEnd(",")

                            If Not Regex.IsMatch(readArray(k), "ネクストエンジン|メールディーラー|発注管理個人別") Then
                                Dim rowNum As Integer = k - 1
                                If i = 0 Then   'サーバー
                                    If InStr(DGV1.Item(0, rowNum).Value, "●") > 0 Then
                                        If dgvStr = readStr Then
                                            saveStr &= readArray(k) & vbCrLf
                                        Else
                                            saveStr &= dgvStr & nowTime & "," & dgvNokori & vbCrLf
                                        End If
                                    Else
                                        saveStr &= readArray(k) & vbCrLf
                                    End If
                                Else            'ローカル
                                    If dgvStr = readStr Then
                                        saveStr &= readArray(k) & vbCrLf
                                    Else
                                        saveStr &= dgvStr & nowTime & "," & dgvNokori & vbCrLf
                                    End If
                                End If
                            Else
                                If i = 0 Then   'サーバー
                                    saveStr &= readArray(k) & vbCrLf
                                Else            'ローカル
                                    saveStr &= dgvStr & nowTime & "," & dgvNokori & vbCrLf
                                End If
                            End If
                        Else
                            saveStr &= readArray(k) & vbCrLf
                        End If
                    Else
                        saveStr &= readArray(k) & vbCrLf
                    End If

                End If
            Next

            Dim cnvStr As String = base64.Encode(saveStr)
            File.WriteAllText(savePath(i), cnvStr, ENC_UTF8)
            'File.WriteAllText("D:\Desktop\test" & i & ".txt", saveStr, ENC_UTF8)
        Next

        MsgBox("保存しました")
    End Sub

    'Private Sub 保存ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 保存ToolStripMenuItem.Click
    '    Dim selRow As Integer = DGV1.SelectedCells(0).RowIndex
    '    If DGV1.Item(0, selRow).Value = ">" Then
    '        Exit Sub
    '    End If

    '    Dim DR As DialogResult = MsgBox("選択行のみ保存します", MsgBoxStyle.OkCancel Or MsgBoxStyle.SystemModal)
    '    Dim savePath As New ArrayList
    '    Dim serverFlag As Boolean = False
    '    If DR = DialogResult.OK Then
    '        savePath.Add(serverFile)
    '        savePath.Add(localFile)
    '    Else
    '        Exit Sub
    '    End If

    '    Dim nowTime As String = Now.ToString
    '    For k As Integer = 0 To savePath.Count - 1
    '        Dim saveArray As New ArrayList
    '        For c As Integer = 0 To DGV1.ColumnCount - 1
    '            saveArray.Add(DGV1.Item(c, selRow).Value)
    '        Next

    '        '共有しない行チェック
    '        If saveArray(0) <> "●" Then
    '            If k = 0 Then
    '                Continue For
    '            End If
    '        End If

    '        Dim saveStr As String = ""
    '        'Dim readArray As String() = File.ReadAllLines(savePath(0), ENC_UTF8)
    '        Dim readArray As String() = Split(base64.Decode(File.ReadAllText(savePath(0), ENC_UTF8)), vbCrLf)
    '        For i As Integer = 0 To readArray.Length - 1
    '            If i = 0 Then
    '                If saveArray(0) <> "●" Then
    '                    saveStr = readArray(i) & vbCrLf
    '                Else
    '                    saveStr = nowTime & vbCrLf
    '                End If
    '            Else
    '                Dim rA As String() = Split(readArray(i), ",")
    '                If rA.Length > 1 Then
    '                    If rA(1) = saveArray(1) Then
    '                        Dim str As String = "-" & saveArray(0) & ","
    '                        str &= saveArray(1) & "," & saveArray(2) & ","
    '                        str &= saveArray(5) & "," & nowTime
    '                        saveStr &= str & vbCrLf
    '                    Else
    '                        saveStr &= readArray(i) & vbCrLf
    '                    End If
    '                Else
    '                    saveStr &= readArray(i) & vbCrLf
    '                End If
    '            End If
    '        Next


    '        Dim cnvStr As String = base64.Encode(saveStr)
    '        File.WriteAllText(savePath(k), cnvStr, ENC_UTF8)
    '    Next

    '    MsgBox("保存しました")
    'End Sub

    Private Sub 閉じるToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 閉じるToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub ContextMenuStrip1_Opening(sender As Object, e As CancelEventArgs) Handles ContextMenuStrip1.Opening
        If Form1.AdminFlag Then
            ToolStripSeparator1.Visible = True
            管理用出力ToolStripMenuItem.Visible = True
            管理用出力サーバーToolStripMenuItem.Visible = True
            管理用取込ToolStripMenuItem.Visible = True
            管理用取込サーバーToolStripMenuItem.Visible = True
        Else
            ToolStripSeparator1.Visible = False
            管理用出力ToolStripMenuItem.Visible = False
            管理用出力サーバーToolStripMenuItem.Visible = False
            管理用取込ToolStripMenuItem.Visible = False
            管理用取込サーバーToolStripMenuItem.Visible = False
        End If
    End Sub

    Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
    Private Sub 管理用出力ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 管理用出力ToolStripMenuItem.Click
        Dim readArray As String() = Split(base64.Decode(File.ReadAllText(localFile, ENC_UTF8)), vbCrLf)
        File.WriteAllLines(desktopPath & "\loginlist4.dat", readArray, ENC_UTF8)
    End Sub

    Private Sub 管理用出力サーバーToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 管理用出力サーバーToolStripMenuItem.Click
        Dim readArray As String() = Split(base64.Decode(File.ReadAllText(serverFile, ENC_UTF8)), vbCrLf)
        File.WriteAllLines(desktopPath & "\loginlist4.dat", readArray, ENC_UTF8)
    End Sub

    Private Sub 管理用取込ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 管理用取込ToolStripMenuItem.Click
        Torikomi(localFile, 0)
    End Sub

    Private Sub 管理用取込サーバーToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 管理用取込サーバーToolStripMenuItem.Click
        Torikomi(serverFile, 1)
    End Sub

    Private Sub Torikomi(SavePath As String, mode As Integer)
        Dim ofd As New OpenFileDialog With {
            .FileName = "",
            .InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
            .Filter = "DATファイル(*.dat)|*.dat|すべてのファイル(*.*)|*.*",
                    .AutoUpgradeEnabled = False,
            .FilterIndex = 1,
            .Title = "開くファイルを選択してください",
            .RestoreDirectory = True,
            .CheckFileExists = True,
            .CheckPathExists = True
        }
        If ofd.ShowDialog() = DialogResult.OK Then
            Dim openStr As String = File.ReadAllText(ofd.FileName, ENC_UTF8)
            Dim cnvStr As String = base64.Encode(openStr)
            File.WriteAllText(SavePath, cnvStr, ENC_UTF8)
            If mode = 0 Then
                MsgBox("ローカルに保存しました")
            Else
                MsgBox("サーバーに保存しました")
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        localArray = Split(base64.Decode(File.ReadAllText(localFile, ENC_UTF8)), vbCrLf)
        For i As Integer = 0 To localArray.Length - 1
            If localArray(i) <> "" Then
                If Regex.IsMatch(localArray(i), "^楽天=") Then
                    localArray(i) = "楽天=" & ComboBox1.SelectedIndex
                ElseIf Regex.IsMatch(localArray(i), "^Yahoo=") Then
                    localArray(i) = "Yahoo=" & ComboBox2.SelectedIndex
                End If
            End If
        Next

        Dim saveStr As String = ""
        Dim brFlag As Integer = 0
        For i As Integer = 0 To localArray.Length - 1
            If localArray(i) = "" Then  '改行が2行以上は削除
                If brFlag >= 1 Then
                    Continue For
                Else
                    brFlag = 1
                End If
            Else
                brFlag = 0
            End If
            saveStr &= localArray(i) & vbCrLf
        Next

        Dim cnvStr As String = base64.Encode(saveStr)
        File.WriteAllText(localFile, cnvStr, ENC_UTF8)
        MsgBox("ブラウザログイン店舗を更新しました")
        Me.Close()
    End Sub

End Class


'********************************************************************
'Base64エンコード
Public Class MyBase64str

    Private enc As Encoding

    Public Sub New(ByVal encStr As String)
        enc = Encoding.GetEncoding(encStr)
    End Sub

    Public Function Encode(ByVal str As String) As String
        Return Convert.ToBase64String(enc.GetBytes(str))
    End Function

    Public Function Decode(ByVal str As String) As String
        Try
            Return enc.GetString(Convert.FromBase64String(str))
        Catch ex As Exception
            Return str
        End Try
    End Function

End Class
'********************************************************************


