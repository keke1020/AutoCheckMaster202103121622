Imports System.ComponentModel
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text

Public Class Csv_denpyo3_F_dgv
    Public CVD3_mode As String = "change"
    Private serverDir As String = Form1.サーバーToolStripMenuItem.Text & "\denpyoLog\"
    Private lockPath As String = serverDir & "lock.txt"
    Private lockUser As String = ""

    Private Sub Csv_denpyo3_F_dgv_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim todayCsvPath As String = serverDir & Format(Now, "yyyyMMdd") & ".csv"

        If CVD3_mode = "change" Then
            If File.Exists(lockPath) Then
                Dim lockstr As String = File.ReadAllText(lockPath, encSJ)
                lockUser = lockstr
                If lockstr <> "" And lockstr <> Form1.ログイン名ToolStripMenuItem.Text Then
                    MsgBox(lockstr & "さんが編集中です" & vbCrLf & vbCrLf & "強制的にロック解除する場合は" & vbCrLf & lockPath & vbCrLf & "ファイルを削除してください")
                    Me.Close()
                    Exit Sub
                End If
            End If
            File.WriteAllText(lockPath, Form1.ログイン名ToolStripMenuItem.Text, encSJ)
            lockUser = Form1.ログイン名ToolStripMenuItem.Text

            GroupBox3.Visible = True
            GroupBox2.Visible = False
        Else
            todayCsvPath = serverDir & Format(DateTimePicker1.Value, "yyyyMMdd") & ".csv"

            GroupBox3.Visible = False
            GroupBox2.Visible = True
        End If

        If File.Exists(todayCsvPath) Then
            DGV1_read(todayCsvPath)
        End If
    End Sub

    Private Sub DGV1_read(todayCsvPath As String)
        DGV1.Rows.Clear()
        DGV1.Columns.Clear()

        Dim cArray As ArrayList = TM_CSV_READ(todayCsvPath)(0)
        Dim cH As String() = Split(cArray(0), "|=|")

        For c As Integer = 0 To cH.Count - 1
            DGV1.Columns.Add(c, cH(c))
            DGV1.Columns(DGV1.ColumnCount - 1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        For r As Integer = 1 To cArray.Count - 1
            DGV1.Rows.Add()
            Dim line As String() = Split(cArray(r), "|=|")
            For c As Integer = 0 To line.Length - 1
                If line(c) <> "" Then
                    DGV1.Item(c, r - 1).Value = line(c)
                End If
            Next
        Next

        DGV1_show_count()
        DGV1_count()
    End Sub

    Private Sub DGV1_show_count()
        Dim showCount As Integer = 0
        For r As Integer = 0 To DGV1.RowCount - 1
            If DGV1.Rows(r).Visible Then
                showCount += 1
            End If
        Next
        Label9.Text = showCount & " / " & DGV1.RowCount
    End Sub

    Private Sub DGV1_count()
        DGV3.Rows.Clear()

        Dim binsu As Integer() = New Integer() {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}
        Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
        For r As Integer = 0 To DGV1.RowCount - 1
            Dim checkStr As String = DGV1.Item(dH1.IndexOf("倉庫"), r).Value & DGV1.Item(dH1.IndexOf("便種"), r).Value & DGV1.Item(dH1.IndexOf("伝票ソフト"), r).Value
            Select Case checkStr
                Case "太宰府陸便BIZlogi", "太宰府陸便e飛伝2"
                    binsu(0) += CInt(DGV1.Item(dH1.IndexOf("マスタ便数"), r).Value)
                Case "太宰府航空便BIZlogi", "太宰府航空便e飛伝2"
                    binsu(1) += CInt(DGV1.Item(dH1.IndexOf("マスタ便数"), r).Value)
                Case "太宰府陸便メール便", "太宰府航空便メール便"
                    binsu(2) += CInt(DGV1.Item(dH1.IndexOf("マスタ便数"), r).Value)
                Case "太宰府陸便定形外", "太宰府航空便定形外"
                    binsu(3) += CInt(DGV1.Item(dH1.IndexOf("マスタ便数"), r).Value)
                Case "井相田陸便BIZlogi", "井相田陸便e飛伝2"
                    binsu(4) += CInt(DGV1.Item(dH1.IndexOf("マスタ便数"), r).Value)
                Case "井相田航空便BIZlogi", "井相田航空便e飛伝2"
                    binsu(5) += CInt(DGV1.Item(dH1.IndexOf("マスタ便数"), r).Value)
                Case "井相田陸便メール便", "井相田航空便メール便"
                    binsu(6) += CInt(DGV1.Item(dH1.IndexOf("マスタ便数"), r).Value)
                Case "井相田陸便定形外", "井相田航空便定形外"
                    binsu(7) += CInt(DGV1.Item(dH1.IndexOf("マスタ便数"), r).Value)
                Case "名古屋陸便BIZlogi", "名古屋陸便e飛伝2"
                    binsu(8) += CInt(DGV1.Item(dH1.IndexOf("マスタ便数"), r).Value)
                Case "名古屋航空便BIZlogi", "名古屋航空便e飛伝2"
                    binsu(9) += CInt(DGV1.Item(dH1.IndexOf("マスタ便数"), r).Value)
                Case "名古屋陸便メール便", "名古屋航空便メール便"
                    binsu(10) += CInt(DGV1.Item(dH1.IndexOf("マスタ便数"), r).Value)
                Case "名古屋陸便定形外", "名古屋航空便定形外"
                    binsu(11) += CInt(DGV1.Item(dH1.IndexOf("マスタ便数"), r).Value)
                Case "太宰府陸便ヤマト"
                    binsu(12) += CInt(DGV1.Item(dH1.IndexOf("マスタ便数"), r).Value)
                Case "井相田陸便ヤマト"
                    binsu(13) += CInt(DGV1.Item(dH1.IndexOf("マスタ便数"), r).Value)
                Case "太宰府船便ヤマト"
                    binsu(16) += CInt(DGV1.Item(dH1.IndexOf("マスタ便数"), r).Value)
                Case "井相田船便ヤマト"
                    binsu(17) += CInt(DGV1.Item(dH1.IndexOf("マスタ便数"), r).Value)
                Case "複数倉庫陸便", "複数倉庫航空便"
                    binsu(15) += CInt(DGV1.Item(dH1.IndexOf("マスタ便数"), r).Value)
                Case Else
                    binsu(14) += CInt(DGV1.Item(dH1.IndexOf("マスタ便数"), r).Value)
            End Select
        Next

        DGV3.Rows.Add("(太)宅配便", binsu(0))
        DGV3.Rows.Add("(太)航空便", binsu(1))
        DGV3.Rows.Add("(太)ゆパケ", binsu(2))
        DGV3.Rows.Add("(太)定形外", binsu(3))

        DGV3.Item(0, 0).Style.BackColor = Color.FromArgb(240, 248, 255)
        DGV3.Item(0, 1).Style.BackColor = Color.FromArgb(240, 248, 255)
        DGV3.Item(0, 2).Style.BackColor = Color.FromArgb(240, 248, 255)
        DGV3.Item(0, 3).Style.BackColor = Color.FromArgb(240, 248, 255)

        DGV3.Rows.Add("(井)宅配便", binsu(4))
        DGV3.Rows.Add("(井)航空便", binsu(5))
        DGV3.Rows.Add("(井)ゆパケ", binsu(6))
        DGV3.Rows.Add("(井)定形外", binsu(7))

        DGV3.Item(0, 4).Style.BackColor = Color.FromArgb(225, 255, 255)
        DGV3.Item(0, 5).Style.BackColor = Color.FromArgb(225, 255, 255)
        DGV3.Item(0, 6).Style.BackColor = Color.FromArgb(225, 255, 255)
        DGV3.Item(0, 7).Style.BackColor = Color.FromArgb(225, 255, 255)


        DGV3.Rows.Add("(名)宅配便", binsu(8))
        DGV3.Rows.Add("(名)航空便", binsu(9))
        DGV3.Rows.Add("(名)ゆパケ", binsu(10))
        DGV3.Rows.Add("(名)定形外", binsu(11))

        DGV3.Item(0, 8).Style.BackColor = Color.FromArgb(212, 242, 231)
        DGV3.Item(0, 9).Style.BackColor = Color.FromArgb(212, 242, 231)
        DGV3.Item(0, 10).Style.BackColor = Color.FromArgb(212, 242, 231)
        DGV3.Item(0, 11).Style.BackColor = Color.FromArgb(212, 242, 231)

        DGV3.Rows.Add("(太)ヤマト_陸便", binsu(12))
        DGV3.Rows.Add("(太)ヤマト_船便", binsu(16))
        DGV3.Rows.Add("(井)ヤマト_陸便", binsu(13))
        DGV3.Rows.Add("(井)ヤマト_船便", binsu(17))

        DGV3.Item(0, 12).Style.BackColor = Color.FromArgb(240, 248, 255)
        DGV3.Item(0, 13).Style.BackColor = Color.FromArgb(240, 248, 255)
        DGV3.Item(0, 14).Style.BackColor = Color.FromArgb(240, 248, 255)
        DGV3.Item(0, 15).Style.BackColor = Color.FromArgb(240, 248, 255)

        DGV3.Rows.Add("複数倉庫", binsu(15))
        DGV3.Rows.Add("不明", binsu(14))
    End Sub

    '行番号を表示する
    Private Sub DataGridView1_RowPostPaint(ByVal sender As DataGridView, ByVal e As DataGridViewRowPostPaintEventArgs) Handles DGV1.RowPostPaint
        ' 行ヘッダのセル領域を、行番号を描画する長方形とする（ただし右端に4ドットのすき間を空ける）
        Dim rect As New Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, sender.RowHeadersWidth - 4, sender.Rows(e.RowIndex).Height)
        ' 上記の長方形内に行番号を縦方向中央＆右詰で描画する　フォントや色は行ヘッダの既定値を使用する
        TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), sender.RowHeadersDefaultCellStyle.Font, rect, sender.RowHeadersDefaultCellStyle.ForeColor, TextFormatFlags.VerticalCenter Or TextFormatFlags.Right)
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        For r As Integer = 0 To DGV1.RowCount - 1
            Dim flag As Boolean = False
            For c As Integer = 0 To DGV1.ColumnCount - 1
                If DGV1.Item(c, r).Value <> "" Then
                    If Regex.IsMatch(DGV1.Item(c, r).Value, TextBox33.Text) Then
                        DGV1.Rows(r).Visible = True
                        flag = True
                        Exit For
                    End If
                End If
            Next

            If Not flag Then
                DGV1.Rows(r).Visible = False
            End If
        Next
        DGV1_show_count()

        MsgBox("抽出しました")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox33.Text = ""
        For r As Integer = 0 To DGV1.RowCount - 1
            DGV1.Rows(r).Visible = True
        Next
    End Sub

    Private Sub 行を削除ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 行を削除ToolStripMenuItem.Click
        Dim selRow As New ArrayList
        For i As Integer = 0 To DGV1.SelectedCells.Count - 1
            If Not selRow.Contains(DGV1.SelectedCells(i).RowIndex) Then
                selRow.Add(DGV1.SelectedCells(i).RowIndex)
            End If
        Next

        selRow.Sort()
        For i As Integer = selRow.Count - 1 To 0 Step -1
            DGV1.Rows.RemoveAt(selRow(i))
        Next
    End Sub

    Private Sub Csv_denpyo3_F_dgv_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If CVD3_mode = "change" Then
            If lockUser = Form1.ログイン名ToolStripMenuItem.Text Then
                File.Delete(lockPath)
            End If
        End If
        Me.Dispose()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DGV1_count()

        Dim str As String = ""
        For c As Integer = 0 To DGV1.ColumnCount - 1
            If c <> 0 Then
                str &= ","
            End If
            str &= DGV1.Columns(c).HeaderText
        Next
        str &= vbCrLf

        For r As Integer = 0 To DGV1.RowCount - 1
            For c As Integer = 0 To DGV1.ColumnCount - 1
                If c <> 0 Then
                    str &= ","
                End If
                str &= """" & DGV1.Item(c, r).Value & """"
            Next
            str &= vbCrLf
        Next
        Dim todayCsvPath As String = serverDir & Format(Now, "yyyyMMdd") & ".csv"
        File.WriteAllText(todayCsvPath, str, encSJ)

        Dim str2 As String = ""
        For r As Integer = 0 To DGV3.RowCount - 1
            str2 &= DGV3.Item(1, r).Value & ","
        Next
        Dim todayTxtPath As String = serverDir & Format(Now, "yyyyMMdd") & ".txt"
        File.WriteAllText(todayTxtPath, str2, encSJ)

        File.Delete(lockPath)

        MsgBox("保存しました")
        Me.Close()
    End Sub

    Private Sub DGV1_SelectionChanged(sender As Object, e As EventArgs) Handles DGV1.SelectionChanged
        If DGV1.SelectedCells.Count > 0 Then
            Dim dCell As DataGridViewSelectedCellCollection = DGV1.SelectedCells
            For i As Integer = dCell.Count - 1 To 0 Step -1
                If DGV1.Rows(dCell(i).RowIndex).Visible = False Then
                    dCell(i).Selected = False
                End If
            Next
            ToolStripStatusLabel1.Text = DGV1.SelectedCells.Count
            ToolStripStatusLabel2.Text = DGV1.SelectedCells(0).Value
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim todayCsvPath As String = serverDir & Format(DateTimePicker1.Value, "yyyyMMdd") & ".csv"
        If File.Exists(todayCsvPath) Then
            DGV1_read(todayCsvPath)
        Else
            MsgBox("データがありません")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim DR As DialogResult = MsgBox("ロックを強制解除します。実績数に影響が出る場合があります。伝票担当に確認の上作業してください。", MsgBoxStyle.OkCancel)

        If DR = DialogResult.OK Then
            If File.Exists(lockPath) Then
                File.Delete(lockPath)
                MsgBox("ロック解除しました")
            Else
                MsgBox("ロックファイルがありませんでした")
            End If
        Else
            MsgBox("ロック解除をキャンセルしました")
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If DGV1.RowCount > 0 Then
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
                sfd.FileName = sPath & Format(Now(), "_yyyymmddhhmmss") & ".csv"
            Else
                sfd.FileName = "denpyoList" & Format(Now(), "_yyyymmddhhmmss") & ".csv"
            End If


            If sfd.ShowDialog(Me) = DialogResult.OK Then
                Dim strArray As New ArrayList
                strArray.Add("商品コード,出荷数,倉庫")

                'Dim dgv6CodeArray As New ArrayList
                Dim dH6 As ArrayList = TM_HEADER_GET(Csv_denpyo3.DGV6)
                'If Csv_denpyo3.DGV6.RowCount > 0 Then
                '    For r As Integer = 0 To Csv_denpyo3.DGV6.RowCount - 1
                '        dgv6CodeArray.Add(Csv_denpyo3.DGV6.Item(dH6.IndexOf("商品コード"), r).Value.ToString.ToLower)
                '    Next
                'End If

                Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
                Dim temp_code As New ArrayList
                Dim souko_arr As String() = New String() {"太宰府", "井相田", "名古屋"}


                For s As Integer = 0 To souko_arr.Count - 1
                    Dim souko As String = souko_arr(s)
                    For c As Integer = 0 To DGV1.RowCount - 1
                        If Regex.IsMatch(DGV1.Item(dH1.IndexOf("倉庫"), c).Value, souko) Then
                            Dim cArray As String() = Split(DGV1.Item(dH1.IndexOf("商品マスタ"), c).Value, "、")
                            If cArray.Count > 0 Then
                                For k As Integer = 0 To cArray.Length - 1
                                    Dim code As String() = Split(cArray(k), "*")
                                    If code.Count = 2 Then
                                        'If dgv6CodeArray.Count > 0 Then
                                        Dim code_name As String = code(0).ToLower
                                        Dim code_count As String = code(1).Trim.Replace("・", "")
                                        'If dgv6CodeArray.Contains(code_name) And code_count = Int(code_count) Then
                                        If IsNumeric(code_count) Then
                                            If temp_code.Contains(code_name) Then
                                                Dim hasCount As Integer = 0
                                                For r As Integer = 0 To strArray.Count - 1
                                                    Dim compare_str As String() = Split(strArray(r), ",")
                                                    'If Regex.IsMatch(compare_str(0), code_name) And Regex.IsMatch(compare_str(2), souko) Then
                                                    If (compare_str(0) = code_name) And (compare_str(2) = souko) Then
                                                        strArray(r) = code_name & "," & (Int(compare_str(1)) + Int(code_count)) & "," & souko
                                                        hasCount = hasCount + 1
                                                        Exit For
                                                    End If
                                                Next
                                                If hasCount = 0 Then
                                                    temp_code.Add(code_name)
                                                    strArray.Add(code_name & "," & code_count & "," & souko)
                                                End If
                                            Else
                                                temp_code.Add(code_name)
                                                strArray.Add(code_name & "," & code_count & "," & souko)
                                            End If
                                        End If
                                        'End If
                                    End If
                                Next
                            End If
                        End If
                    Next
                Next

                Dim Values() As String = DirectCast(strArray.ToArray(GetType(String)), String())
                File.WriteAllLines(sfd.FileName, Values, Encoding.GetEncoding("SHIFT-JIS"))
            End If
        Else
            MsgBox("データがない為にダウンロードできません")
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If DGV1.RowCount > 0 Then
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
                sfd.FileName = sPath & Format(Now(), "_yyyymmddhhmmss") & ".csv"
            Else
                sfd.FileName = "denpyo_data" & Format(Now(), "_yyyymmddhhmmss") & ".csv"
            End If

            If sfd.ShowDialog(Me) = DialogResult.OK Then
                DGV_TO_CSV_SAVE(sfd.FileName, DGV1,, True)
            End If
        End If
    End Sub
End Class