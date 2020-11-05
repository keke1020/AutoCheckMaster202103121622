Imports System.ComponentModel
Imports System.IO
Imports System.Text.RegularExpressions

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

        Dim binsu As Integer() = New Integer() {0, 0, 0, 0, 0, 0, 0, 0, 0}
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
                Case Else
                    binsu(8) += CInt(DGV1.Item(dH1.IndexOf("マスタ便数"), r).Value)
            End Select
        Next

        DGV3.Rows.Add("(太)宅配便", binsu(0))
        DGV3.Rows.Add("(太)航空便", binsu(1))
        DGV3.Rows.Add("(太)ゆパケ", binsu(2))
        DGV3.Rows.Add("(太)定形外", binsu(3))
        DGV3.Rows.Add("(井)宅配便", binsu(4))
        DGV3.Rows.Add("(井)航空便", binsu(5))
        DGV3.Rows.Add("(井)ゆパケ", binsu(6))
        DGV3.Rows.Add("(井)定形外", binsu(7))
        DGV3.Rows.Add("不明", binsu(8))
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
End Class