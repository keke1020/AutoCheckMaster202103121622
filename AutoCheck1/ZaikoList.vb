Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

Public Class ZaikoList
    Dim appPath As String = Path.GetDirectoryName(Reflection.Assembly.GetExecutingAssembly().Location)
    Dim jogaiPath As String = appPath & "\config\除外コード.txt"
    Dim ENC_SJ As Encoding = Encoding.GetEncoding("SHIFT-JIS")
    Dim dH1 As New ArrayList


    Private Sub ZaikoList_Load(sender As Object, e As EventArgs) Handles Me.Load
        dH1 = TM_HEADER_GET(DGV1)
        Label4.Text = Format(File.GetLastWriteTime(appPath & "\" & TextBox1.Text), "yyyy/MM/dd HH:mm:ss")

        If File.Exists(jogaiPath) Then
            TextBox3.Text = File.ReadAllText(jogaiPath, ENC_SJ)
        End If

        ComboBox1.Items.Add("")
        ComboBox2.Items.Add("")
        For i As Integer = 0 To dH1.Count - 1
            ComboBox1.Items.Add(dH1(i))
            ComboBox2.Items.Add(dH1(i))
        Next
        ComboBox1.SelectedItem = "棚番"
        ComboBox2.SelectedItem = "商品コード"
    End Sub

    '行番号を表示する
    Private Sub DataGridView_RowPostPaint(ByVal sender As DataGridView, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles _
        DGV1.RowPostPaint
        ' 行ヘッダのセル領域を、行番号を描画する長方形とする（ただし右端に4ドットのすき間を空ける）
        Dim rect As New Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, sender.RowHeadersWidth - 4, sender.Rows(e.RowIndex).Height)
        ' 上記の長方形内に行番号を縦方向中央＆右詰で描画する　フォントや色は行ヘッダの既定値を使用する
        TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), sender.RowHeadersDefaultCellStyle.Font, rect, sender.RowHeadersDefaultCellStyle.ForeColor, TextFormatFlags.VerticalCenter Or TextFormatFlags.Right)
    End Sub

    Private Sub DataGridView1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles DGV1.DragDrop
        For Each filename As String In e.Data.GetData(DataFormats.FileDrop)
            readData(filename)
            Exit For
        Next

        Dim rowCount As Integer = 0
        For r As Integer = 0 To DGV1.RowCount - 1
            If DGV1.Rows(r).Visible Then
                rowCount += 1
            End If
        Next
        TextBox2.Text = rowCount

        ToolStripProgressBar1.Value = 0
        MsgBox("読込終了")
    End Sub

    Private Sub DataGridView1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles DGV1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim ofd As New OpenFileDialog With {
            .InitialDirectory = Environment.SpecialFolder.Desktop,
            .Filter = "CSVファイル(*.csv)|*.csv|すべてのファイル(*.*)|*.*",
            .FilterIndex = 1,
            .Title = "開くファイルを選択してください",
            .RestoreDirectory = True,
            .CheckFileExists = True,
            .CheckPathExists = True
        }

        If ofd.ShowDialog() = DialogResult.OK Then
            readData(ofd.FileName)
        End If

        Dim rowCount As Integer = 0
        For r As Integer = 0 To DGV1.RowCount - 1
            If DGV1.Rows(r).Visible Then
                rowCount += 1
            End If
        Next
        TextBox2.Text = rowCount

        ToolStripProgressBar1.Value = 0
        MsgBox("読込終了")
    End Sub

    '再読込
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        readData(TextBox40.Text)

        Dim rowCount As Integer = 0
        For r As Integer = 0 To DGV1.RowCount - 1
            If DGV1.Rows(r).Visible Then
                rowCount += 1
            End If
        Next
        TextBox2.Text = rowCount

        ToolStripProgressBar1.Value = 0
        MsgBox("読込終了")
    End Sub

    Private Function readData(filename)
        'TextBox40.Text = Path.GetFileName(filename)
        TextBox40.Text = filename
        DGV1.Rows.Clear()
        Dim csvH As New Generic.List(Of String)
        Dim csvRecords1 As ArrayList = TM_CSV_READ(filename)(0)

        ToolStripProgressBar1.Value = 0
        ToolStripProgressBar1.Maximum = csvRecords1.Count
        For r As Integer = 0 To csvRecords1.Count - 1
            Dim sArray As String() = Split(csvRecords1(r), "|=|")
            If r = 0 Then
                csvH = sArray.ToList()
            Else
                DGV1.Rows.Add()
                For c As Integer = 0 To DGV1.ColumnCount - 1
                    Dim header As String = DGV1.Columns(c).HeaderText
                    If csvH.Contains(header) Then
                        DGV1.Item(c, r - 1).Value = sArray(csvH.IndexOf(header))
                    End If
                Next
            End If
            ToolStripProgressBar1.Value += 1
        Next
        ToolStripProgressBar1.Value = 0
        ToolStripProgressBar1.Maximum = DGV1.Rows.Count

        'location.csv合成
        Dim csvRecords2 As ArrayList = TM_CSV_READ(appPath & "\" & TextBox1.Text)(0)
        Dim csvH2 As String() = Split(csvRecords2(0), "|=|")
        If csvH2.Contains("ロケーション") Then
            Dim csvCodeList As New ArrayList
            For i As Integer = 0 To csvRecords2.Count - 1
                Dim csvH3 As String() = Split(csvRecords2(i), "|=|")
                csvCodeList.Add(csvH3(Array.IndexOf(csvH2, "商品コード")))
            Next
            For r As Integer = 0 To DGV1.RowCount - 1
                Dim code As String = DGV1.Item(dH1.IndexOf("商品コード"), r).Value.ToString.ToLower
                If csvCodeList.Contains(code) Then
                    Dim csvH4 As String() = Split(csvRecords2(csvCodeList.IndexOf(code)), "|=|")
                    DGV1.Item(dH1.IndexOf("棚番"), r).Value = csvH4(Array.IndexOf(csvH2, "ロケーション"))
                End If
                ToolStripProgressBar1.Value += 1
            Next
        End If

        Return 1
    End Function

    'location.csv更新
    Private Sub Button36_Click(sender As Object, e As EventArgs) Handles Button36.Click
        If Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable() Then
            Dim srcPath As String = "\\SERVER2-PC\ordery\logi\location.csv"
            Dim dstPath As String = appPath & "\" & TextBox1.Text
            If File.Exists(srcPath) Then
                File.Copy(srcPath, dstPath, True)
                Label4.Text = File.GetLastWriteTime(dstPath)
                MsgBox("location.csvを更新しました")
            Else
                MsgBox("location.csvが見つかりません")
            End If
        Else
            MsgBox("LAN接続がありません")
        End If
    End Sub

    Dim maereg As String = ""
    Private Sub CheckedChanged(sender As Object, e As EventArgs) Handles _
            RadioButton1.CheckedChanged, RadioButton2.CheckedChanged, RadioButton3.CheckedChanged,
            RadioButton4.CheckedChanged, RadioButton5.CheckedChanged
        If DGV1.RowCount = 0 Then
            Exit Sub
        End If

        Dim reg As String = ""
        If RadioButton2.Checked Then
            reg = "3[A-Za-z]"
        ElseIf RadioButton3.Checked Then
            reg = "1[A-Za-z]"
        ElseIf RadioButton4.Checked Then
            reg = "2[A-Za-z]"
        ElseIf RadioButton5.Checked Then
            reg = "SD"
        End If

        If maereg = reg Then
            Exit Sub
        Else
            maereg = reg
        End If

        If Kirikae(reg) = 1 Then
            MsgBox("表示切り替えました")
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles _
        CheckBox1.CheckedChanged, CheckBox2.CheckedChanged, CheckBox3.CheckedChanged
        If DGV1.RowCount = 0 Then
            Exit Sub
        End If

        Dim reg As String = ""
        If RadioButton2.Checked Then
            reg = "3[A-Za-z]"
        ElseIf RadioButton3.Checked Then
            reg = "1[A-Za-z]"
        ElseIf RadioButton4.Checked Then
            reg = "2[A-Za-z]"
        End If

        If Kirikae(reg) = 1 Then
            MsgBox("表示切り替えました")
        End If
    End Sub

    Private Function Kirikae(reg As String)
        ToolStripProgressBar1.Value = 0
        ToolStripProgressBar1.Maximum = DGV1.Rows.Count

        '除外コード
        Dim jogaiArray As New ArrayList
        If TextBox3.Text <> "" Then
            Dim jogaiLines As String() = Split(TextBox3.Text, vbCrLf)
            For i As Integer = 0 To jogaiLines.Length - 1
                If jogaiLines(i) <> "" Then
                    jogaiArray.Add(jogaiLines(i))
                End If
            Next
            TextBox3.Text = ""
            For i As Integer = 0 To jogaiArray.Count - 1
                TextBox3.Text &= jogaiArray(i) & vbCrLf
            Next
        End If

        Dim rowCount As Integer = 0
        For r As Integer = 0 To DGV1.RowCount - 1
            Dim visibleFlag As Boolean = True

            '取扱中
            If CheckBox1.Checked Then
                Dim tori As String = DGV1.Item(dH1.IndexOf("取扱区分"), r).Value
                If Not tori = "取扱中" Then
                    visibleFlag = False
                End If
            End If

            '予約除外
            If visibleFlag And CheckBox2.Checked Then
                Dim yoyaku As String = DGV1.Item(dH1.IndexOf("商品区分"), r).Value
                If yoyaku = "予約" Then
                    visibleFlag = False
                End If
            End If

            '棚番
            If visibleFlag Then
                Dim tana As String = DGV1.Item(dH1.IndexOf("棚番"), r).Value
                If reg = "" Then
                    If Not DGV1.Rows(r).Visible Then
                        visibleFlag = True
                    End If
                ElseIf reg = "2" Then
                    If tana = "" Then
                        visibleFlag = True
                    ElseIf Regex.IsMatch(tana, "2[A-Za-z]") Then
                        visibleFlag = True
                    Else
                        visibleFlag = False
                    End If
                ElseIf tana = "" Then
                    visibleFlag = False
                    visibleFlag = False
                ElseIf Regex.IsMatch(tana, "^" & reg) Then
                    visibleFlag = True
                Else
                    visibleFlag = False
                End If
            End If

            '在庫数
            If visibleFlag And CheckBox3.Checked Then
                Dim zaiko As String = DGV1.Item(dH1.IndexOf("在庫数"), r).Value
                If IsNumeric(zaiko) Then
                    If CInt(zaiko) >= NumericUpDown1.Value And CInt(zaiko) <= NumericUpDown2.Value Then
                        visibleFlag = True
                    Else
                        visibleFlag = False
                    End If
                Else
                    visibleFlag = False
                End If
            End If

            '除外コード
            If visibleFlag And TextBox3.Text <> "" Then
                Dim code As String = DGV1.Item(dH1.IndexOf("商品コード"), r).Value
                Dim jogaiFlag As Boolean = False
                For i As Integer = 0 To jogaiArray.Count - 1
                    If Regex.IsMatch(code, "^" & jogaiArray(i)) Then
                        jogaiFlag = True
                        Exit For
                    End If
                Next
                If jogaiFlag Then
                    visibleFlag = False
                Else
                    visibleFlag = True
                End If
            End If

            '表示
            If visibleFlag Then
                DGV1.Rows(r).Visible = True
                rowCount += 1

                'JANチェック
                Dim cJan As String = ""
                If RadioButton2.Checked Then
                    cJan = RadioButton2.Text & "/"
                ElseIf RadioButton3.Checked Then
                    cJan = RadioButton3.Text & "/"
                ElseIf RadioButton4.Checked Then
                    cJan = RadioButton4.Text & "/"
                ElseIf RadioButton5.Checked Then
                    cJan = RadioButton5.Text & "/"
                Else
                    cJan = "/"
                End If
                If InStr(DGV1.Item(dH1.IndexOf("JANコード"), r).Value, cJan) > 0 Then
                    DGV1.Item(dH1.IndexOf("JANコード"), r).Style.BackColor = Color.Empty
                Else
                    DGV1.Item(dH1.IndexOf("JANコード"), r).Style.BackColor = Color.Yellow
                End If
            Else
                DGV1.Rows(r).Visible = False
            End If

            ToolStripProgressBar1.Value += 1
        Next

        TextBox2.Text = rowCount
        ToolStripProgressBar1.Value = 0

        Return 1
    End Function

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        NameChangeSave()
    End Sub

    Private Sub NameChangeSave()
        If DGV1.RowCount = 0 Then
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
            sfd.FileName = "新しいファイル.csv"
        End If

        'ダイアログを表示する
        If sfd.ShowDialog(Me) = DialogResult.OK Then
            SaveCsv(sfd.FileName)
            MsgBox(sfd.FileName & vbCrLf & "保存しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    ' DataGridViewをCSV出力する
    Public Sub SaveCsv(ByVal fp As String)
        DGV1.EndEdit()

        Dim CsvDelimiter As String = ","

        ToolStripProgressBar1.Maximum = DGV1.Rows.Count
        ToolStripProgressBar1.Value = 0
        ' CSVファイルオープン
        Dim str As String = ""
        Dim sw As StreamWriter = New StreamWriter(fp, False, ENC_SJ)

        Dim header As String = ""
        For c As Integer = 0 To DGV1.ColumnCount - 1
            header &= DGV1.Columns(c).HeaderText & CsvDelimiter
        Next
        header = header.TrimEnd(CsvDelimiter) & vbLf
        sw.Write(header)

        For r As Integer = 0 To DGV1.Rows.Count - 1
            If DGV1.Rows(r).Visible Then
                For c As Integer = 0 To DGV1.Columns.Count - 1
                    ' DataGridViewのセルのデータ取得
                    Dim dt As String = ""
                    If DGV1.Rows(r).Cells(c).Value Is Nothing = False Then
                        dt = DGV1.Rows(r).Cells(c).Value.ToString()
                        dt = Replace(dt, vbCrLf, vbLf)
                        'dt = Replace(dt, vbLf, "")
                        'If mode = 1 Then
                        '    dt = """" & dt & """"
                        'Else
                        If Not dt Is Nothing Then
                            dt = Replace(dt, """", """""")
                            Select Case True
                                Case InStr(dt, CsvDelimiter), InStr(dt, vbLf), InStr(dt, vbCr), InStr(dt, """") 'Regex.IsMatch(dt, ".*,.*|.*" & vbLf & ".*|.*" & vbCr & ".*|.*"".*")
                                    dt = """" & dt & """"
                            End Select
                        End If
                        'End If
                    End If
                    If c < DGV1.Columns.Count - 1 Then
                        dt = dt & CsvDelimiter
                    End If
                    ' CSVファイル書込
                    sw.Write(dt)
                Next
                sw.Write(vbLf)
            End If
            ToolStripProgressBar1.Value += 1
            If ToolStripProgressBar1.Value Mod 500 = 0 Then
                Application.DoEvents()
            End If
        Next

        ' CSVファイルクローズ
        sw.Close()

        ToolStripProgressBar1.Value = 0
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        File.WriteAllText(jogaiPath, TextBox3.Text, ENC_SJ)
        MsgBox("保存しました")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If DGV1.RowCount = 0 Then
            Exit Sub
        ElseIf ComboBox1.SelectedItem = "" And ComboBox2.SelectedItem = "" Then
            Exit Sub
        End If

        Dim locationSet As New DataSet
        Dim locationTable As New DataTable("main")
        Dim sortStr As String = ""
        If ComboBox1.SelectedItem <> "" Then
            locationTable.Columns.Add(ComboBox1.SelectedItem)
            sortStr = ComboBox1.SelectedItem & " ASC"
        End If
        If ComboBox2.SelectedItem <> "" Then
            locationTable.Columns.Add(ComboBox2.SelectedItem)
            If sortStr <> "" Then
                sortStr &= ", "
            End If
            sortStr &= ComboBox1.SelectedItem & " ASC"
        End If
        locationTable.Columns.Add("STR")
        locationSet.Tables.Add(locationTable)

        For r As Integer = 0 To DGV1.RowCount - 1
            If DGV1.Rows(r).Visible Then
                Dim m1 As String = ""
                Dim m2 As String = ""
                If ComboBox1.SelectedItem <> "" Then
                    m1 = DGV1.Item(dH1.IndexOf(ComboBox1.SelectedItem), r).Value
                End If
                If ComboBox2.SelectedItem <> "" Then
                    m2 = DGV1.Item(dH1.IndexOf(ComboBox2.SelectedItem), r).Value
                End If
                Dim line As String = ""
                For c As Integer = 0 To DGV1.ColumnCount - 1
                    line &= DGV1.Item(c, r).Value & ","
                Next
                line = line.TrimEnd(",")

                locationSet.Tables("main").Rows.Add(m1, m2, line)
            End If
        Next

        Dim tb As DataTable = locationSet.Tables("main")
        If tb.Rows.Count > 0 Then
            Dim view As New DataView(tb) With {
                            .Sort = sortStr
                        }
            DGV1.Rows.Clear()
            For Each row As DataRowView In view
                Dim lineArray As String() = Split(row("STR"), ",")
                DGV1.Rows.Add(lineArray)
            Next
        End If

        MsgBox("並べ替えました")
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        For r As Integer = 0 To DGV1.RowCount - 1
            If DGV1.Rows(r).IsNewRow Then
                Continue For
            ElseIf DGV1.Rows(r).Visible = False Then
                Continue For
            End If

            Dim flag As Boolean = False
            For c As Integer = 0 To DGV1.ColumnCount - 1
                If DGV1.Item(c, r).Style.BackColor = Color.Yellow Then
                    DGV1.Rows(r).Visible = True
                    flag = True
                    Exit For
                End If
            Next

            If Not flag Then
                DGV1.Rows(r).Visible = False
            End If
        Next
    End Sub
End Class