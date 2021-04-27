Imports System.IO
Imports System.Text.RegularExpressions

Public Class Search
    Dim mode1 As String = "1:CSV1"
    Dim mode2 As String = "2:CSV2"
    Dim modeH As String = "0:HTML1"
    Dim start As Integer = 0
    Dim sC As Integer = 0
    Dim sR As Integer = 0
    Dim sPath As String = ""

    Dim dg As DataGridView = Nothing

    Private Sub Search_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        sPath = Path.GetDirectoryName(Form1.appPath) & "\config\search.dat"

        ComboBox18.SelectedIndex = 0
        ComboBox2.SelectAll()
    End Sub

    Private Sub Search_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        If Form1.CSV_FORMS(0).ToolStripStatusLabel5.Text = "1" Then
            ComboBox4.SelectedItem = mode1
        ElseIf HTMLcreate.Label1.Text = "H" Then
            ComboBox4.SelectedItem = modeH
        Else
            ComboBox4.SelectedItem = mode2
        End If
    End Sub

    Private Sub Search_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        CB1_READ()
    End Sub

    Private Sub Search_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.GotFocus
        If ComboBox3.Focused = False Then
            ComboBox2.SelectAll()
        End If
    End Sub

    '検索
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ComboBox2.Text = "" Then
            MsgBox("検索する言葉を入力してください", MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        Dim SearchChar As String = SearchList(0, ComboBox2.Text)

        'Try
        If ComboBox4.SelectedItem = "0:HTML1" Then
            Dim startC As Integer = HTMLcreate.az.Document.CaretIndex
            Dim endC As Integer = HTMLcreate.az.Document.Length
            If startC >= endC Then
                startC = 0
            End If
            Dim SR As Sgry.Azuki.SearchResult = HTMLcreate.az.Document.FindNext(SearchChar, startC, endC, True)
            If Not SR Is Nothing Then
                HTMLcreate.az.Document.SetSelection(SR.Begin, SR.End)
                HTMLcreate.az.ScrollToCaret()
            Else
                MsgBox("指定の文字は見つかりませんでした", MsgBoxStyle.SystemModal)
            End If
        Else
            If ComboBox4.SelectedItem = "1:CSV1" Then
                dg = Form1.CSV_FORMS(0).DataGridView1
            Else
                dg = Form1.CSV_FORMS(1).DataGridView1
            End If

            Dim preSearch As String = SearchList(1, ComboBox2.Text)
            If CheckBox3.Checked Then
                SearchChar = SearchList(1, ComboBox2.Text)
            End If
            Select Case True
                Case RadioButton6.Checked
                    SearchChar = ".*" & SearchChar & ".*"
                Case RadioButton7.Checked
                    SearchChar = "^" & SearchChar & ".*"
                Case RadioButton9.Checked
                    SearchChar = "^" & SearchChar & "$"
                Case RadioButton8.Checked
                    SearchChar = ".*" & SearchChar & "$"
            End Select

            Dim num As Integer = 0
            If CheckBox1.Checked = False Then
                dg.ClearSelection()
                For r = sR To dg.RowCount - 1
                    For c = 0 To dg.ColumnCount - 1
                        If (r = sR And c > sC) Or r > sR Then
                            If InStr(dg.Item(c, r).Value, preSearch) > 0 Then  'regex速度が遅いので、先にinstrで検索
                                If Regex.IsMatch(dg.Item(c, r).Value, SearchChar) Then
                                    dg.Item(c, r).Selected = True
                                    dg.CurrentCell = dg(c, r)
                                    If r = dg.RowCount - 1 And c = dg.ColumnCount - 1 Then
                                        sC = 0
                                        sR = 0
                                        Exit Sub
                                    Else
                                        sC = c
                                        sR = r
                                        Exit Sub
                                    End If
                                End If
                            End If
                            If r = dg.RowCount - 1 And c = dg.ColumnCount - 1 Then
                                sC = 0
                                sR = 0
                                MsgBox("指定の文字は見つかりませんでした", MsgBoxStyle.SystemModal)
                            End If
                        End If
                    Next
                Next
            Else
                Dim dgselCell As DataGridViewSelectedCellCollection = dg.SelectedCells
                dg.ClearSelection()
                Dim searchNum As Integer = 0
                For i As Integer = 0 To dgselCell.Count - 1
                    If dgselCell(i).Value <> "" Then
                        If Regex.IsMatch(dgselCell(i).Value, SearchChar) Then
                            dgselCell(i).Selected = True
                            If CheckBox2.Checked = True Then
                                dgselCell(i).Style.BackColor = Color.LightBlue
                            End If

                            If searchNum = 0 Then
                                dg.CurrentCell = dg(dgselCell(i).ColumnIndex, dgselCell(i).RowIndex)
                                'dg.FirstDisplayedScrollingRowIndex = dgselCell(i).RowIndex
                                'dg.FirstDisplayedScrollingColumnIndex = dgselCell(i).ColumnIndex
                            End If
                            searchNum += 1
                        End If
                        If CheckBox4.Checked = True Then
                            If IsGroup(dgselCell(i).Value) Then

                                dg.Rows(dgselCell(i).RowIndex).DefaultCellStyle.BackColor = Color.Red
                            End If
                        End If
                    End If
                Next
                MsgBox(searchNum & "件見つかりました", MsgBoxStyle.SystemModal)
            End If
        End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.SystemModal)
        'End Try
    End Sub



    Private Sub IsGroupStyly()
        Dim dgselRows As DataGridViewSelectedRowCollection = dg.SelectedRows
        For index = 0 To dgselRows.Count - 1
            If IsGroup(dgselRows(index).Cells(2)) Or IsGroup(dgselRows(index).Cells(3)) Or IsGroup(dgselRows(index).Cells(5)) Then
                dgselRows(index).DefaultCellStyle.BackColor = Color.Red
            End If
        Next
    End Sub

    Dim Group As String() = {"医院", "組合", "会社", "機構", "法人", "薬局", "センター", "(株)", "（株）", "商店", "店", "支社", "(有)"， "(合)"}
    Private Function IsGroup(obj As Object) As Boolean
        For index = 0 To Group.Count - 1
            If InStr(obj, Group(index)) Then
                Return True
            End If
        Next
        Return False
    End Function

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If ComboBox2.Text = "" Then
            MsgBox("検索する言葉を入力してください", MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        Dim SearchChar As String = SearchList(0, ComboBox2.Text)

        If ComboBox4.SelectedItem = "0:HTML1" Then
            Dim SR As Sgry.Azuki.SearchResult = HTMLcreate.az.Document.FindNext(SearchChar, 0)
            If Not SR Is Nothing Then
                HTMLcreate.az.Document.SetSelection(SR.Begin, SR.End)
                HTMLcreate.az.ScrollToCaret()
            Else
                MsgBox("指定の文字は見つかりませんでした", MsgBoxStyle.SystemModal)
            End If
        Else
            sC = 0
            sR = 0
            Button1.PerformClick()
        End If
    End Sub

    '置換
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If ComboBox2.Text = "" And TextBox3.Text = "" Then
            MsgBox("検索する言葉を入力してください", MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        Dim SearchChar As String = ""
        If RadioButton13.Checked = True Then
            SearchChar = SearchList(0, ComboBox2.Text)
        Else
            SearchChar = TextBox3.Text
        End If
        SearchChar = Replace(SearchChar, "[\t]", vbTab)
        SearchChar = Replace(SearchChar, "[\n]", vbCrLf)
        SearchChar = Replace(SearchChar, "\n", vbCr)
        SearchChar = Replace(SearchChar, "\r", vbLf)
        SearchChar = Replace(SearchChar, "\\", "\")

        Dim replaceChar As String = ""
        If RadioButton10.Checked = True Then
            replaceChar = SearchList(1, ComboBox3.Text)
        ElseIf RadioButton15.Checked = True Then
            replaceChar = TextBox5.Text
        End If
        replaceChar = Replace(replaceChar, "[\t]", vbTab)
        replaceChar = Replace(replaceChar, "[\n]", vbCrLf)
        replaceChar = Replace(replaceChar, "\n", vbCr)
        replaceChar = Replace(replaceChar, "\r", vbLf)
        replaceChar = Replace(replaceChar, "\\", "\")

        If ComboBox4.SelectedItem = "0:HTML1" Then
            Dim SR As Sgry.Azuki.SearchResult = HTMLcreate.az.Document.FindNext(SearchChar, HTMLcreate.az.Document.CaretIndex)
            If Not SR Is Nothing Then
                HTMLcreate.az.Document.Replace(replaceChar, SR.Begin, SR.End)
                HTMLcreate.az.Document.SetSelection(SR.Begin, SR.Begin + replaceChar.Length)
                HTMLcreate.az.ScrollToCaret()
                Button2.Text = "次を置換"
                LinkLabel1.Enabled = True
            Else
                Button2.Text = "置　換"
                LinkLabel1.Enabled = False
                MsgBox("指定の文字は見つかりませんでした", MsgBoxStyle.SystemModal)
            End If
        Else
            If ComboBox4.SelectedItem = "1:CSV1" Then
                dg = Form1.CSV_FORMS(0).DataGridView1
            Else
                dg = Form1.CSV_FORMS(1).DataGridView1
            End If
            'For i As Integer = 0 To dg.SelectedCells.Count - 1
            '    dg.SelectedCells.Item(i).Selected = False
            'Next

            'ヘッダーを取得
            Dim dgHeader As ArrayList = New ArrayList
            For c As Integer = 0 To dg.ColumnCount - 1
                If RadioButton11.Checked = True Then
                    dgHeader.Add(dg.Columns(c).HeaderText)
                ElseIf RadioButton12.Checked = True Then
                    dgHeader.Add(dg.Item(c, 0).Value)
                End If
            Next

            Dim op As RegexOptions
            Select Case ComboBox18.SelectedItem
                Case "複数行"
                    op = RegexOptions.Multiline
                Case "単一行"
                    op = RegexOptions.Singleline
                Case "小文字"
                    op = RegexOptions.IgnoreCase
                Case Else
                    op = RegexOptions.None
            End Select

            If CheckBox1.Checked = False Then
                For r = sR To dg.RowCount - 1
                    If RadioButton11.Checked = True Then
                        replaceChar = dg.Item(dgHeader.IndexOf(TextBox1.Text), r).Value
                    ElseIf RadioButton12.Checked = True Then
                        replaceChar = dg.Item(dgHeader.IndexOf(TextBox2.Text), r).Value
                    End If

                    For c = 0 To dg.ColumnCount - 1
                        If (r = sR And c > sC) Or r > sR Then
                            If CheckBox3.Checked = True Then
                                If Regex.IsMatch(dg.Item(c, r).Value, SearchChar, op) Then
                                    dg.Item(c, r).Value = Regex.Replace(dg.Item(c, r).Value, SearchChar, replaceChar, op)
                                    dg.Item(c, r).Selected = True
                                    dg.CurrentCell = dg(c, r)
                                    If r = dg.RowCount - 1 And c = dg.ColumnCount - 1 Then
                                        sC = 0
                                        sR = 0
                                        Button2.Text = "置　換"
                                        LinkLabel1.Enabled = False
                                        MsgBox("指定の文字は見つかりませんでした", MsgBoxStyle.SystemModal)
                                        Exit Sub
                                    Else
                                        sC = c
                                        sR = r
                                        Button2.Text = "次を置換"
                                        LinkLabel1.Enabled = True
                                        Exit Sub
                                    End If
                                End If
                            Else
                                If InStr(dg.Item(c, r).Value, SearchChar) > 0 Then
                                    dg.Item(c, r).Value = Replace(dg.Item(c, r).Value, SearchChar, replaceChar, op)
                                    dg.Item(c, r).Selected = True
                                    dg.CurrentCell = dg(c, r)
                                    If r = dg.RowCount - 1 And c = dg.ColumnCount - 1 Then
                                        sC = 0
                                        sR = 0
                                        Button2.Text = "置　換"
                                        LinkLabel1.Enabled = False
                                        MsgBox("指定の文字は見つかりませんでした", MsgBoxStyle.SystemModal)
                                        Exit Sub
                                    Else
                                        sC = c
                                        sR = r
                                        Button2.Text = "次を置換"
                                        LinkLabel1.Enabled = True
                                        Exit Sub
                                    End If
                                End If
                            End If
                            If r = dg.RowCount - 1 And c = dg.ColumnCount - 1 Then
                                sC = 0
                                sR = 0
                                Button2.Text = "置　換"
                                LinkLabel1.Enabled = False
                                MsgBox("指定の文字は見つかりませんでした", MsgBoxStyle.SystemModal)
                            End If
                        End If
                    Next
                Next
            Else
                Dim num As Integer = 0
                For i As Integer = 0 To dg.SelectedCells.Count - 1
                    If RadioButton11.Checked = True Then
                        replaceChar = dg.Item(dgHeader.IndexOf(TextBox1.Text), dg.SelectedCells(i).RowIndex).Value
                    ElseIf RadioButton12.Checked = True Then
                        replaceChar = dg.Item(dgHeader.IndexOf(TextBox2.Text), dg.SelectedCells(i).RowIndex).Value
                    End If

                    If CheckBox3.Checked = True Then
                        If Regex.IsMatch(dg.SelectedCells(i).Value, SearchChar, op) Then
                            dg.SelectedCells(i).Value = Regex.Replace(dg.SelectedCells(i).Value, SearchChar, replaceChar, op)
                            num += 1
                        End If
                    Else
                        If InStr(dg.SelectedCells(i).Value, SearchChar) > 0 Then
                            dg.SelectedCells(i).Value = Replace(dg.SelectedCells(i).Value, SearchChar, replaceChar, op)
                            num += 1
                        End If
                    End If
                Next
                sC = 0
                sR = 0
                If num <> 0 Then
                    MsgBox(num & "件置換ました", MsgBoxStyle.SystemModal)
                Else
                    MsgBox("指定の文字は見つかりませんでした", MsgBoxStyle.SystemModal)
                End If
            End If
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        sC = 0
        sR = 0
        Button2.Text = "置　換"
        LinkLabel1.Enabled = False
    End Sub

    '全て置換
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If ComboBox2.Text = "" And TextBox3.Text = "" Then
            MsgBox("検索する言葉を入力してください", MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        Dim SearchChar As String = ""
        If RadioButton13.Checked = True Then
            SearchChar = SearchList(0, ComboBox2.Text)
        Else
            SearchChar = TextBox3.Text
        End If
        SearchChar = Replace(SearchChar, "[\t]", vbTab)
        SearchChar = Replace(SearchChar, "[\n]", vbCrLf)
        SearchChar = Replace(SearchChar, "\n", vbCr)
        SearchChar = Replace(SearchChar, "\r", vbLf)
        SearchChar = Replace(SearchChar, "\\", "\")

        Dim replaceChar As String = ""
        If RadioButton10.Checked = True Then
            replaceChar = SearchList(1, ComboBox3.Text)
        ElseIf RadioButton15.Checked = True Then
            replaceChar = TextBox5.Text
        End If
        replaceChar = Replace(replaceChar, "[\t]", vbTab)
        replaceChar = Replace(replaceChar, "[\n]", vbCrLf)
        replaceChar = Replace(replaceChar, "\n", vbCr)
        replaceChar = Replace(replaceChar, "\r", vbLf)
        replaceChar = Replace(replaceChar, "\\", "\")

        If replaceChar Is Nothing Then
            replaceChar = ""
        End If

        If ComboBox4.SelectedItem = "0:HTML1" Then
            Dim SR As Sgry.Azuki.SearchResult = HTMLcreate.az.Document.FindNext(SearchChar, 0)
            If Not SR Is Nothing Then
                Dim num As Integer = 0
                Do While Not SR Is Nothing
                    HTMLcreate.az.Document.Replace(replaceChar, SR.Begin, SR.End)
                    HTMLcreate.az.Document.SetSelection(SR.Begin, SR.Begin + replaceChar.Length)
                    HTMLcreate.az.ScrollToCaret()
                    SR = HTMLcreate.az.Document.FindNext(SearchChar, HTMLcreate.az.Document.CaretIndex)
                    num += 1
                Loop
                MsgBox(num & "件置換しました", MsgBoxStyle.SystemModal)
            Else
                MsgBox("指定の文字は見つかりませんでした", MsgBoxStyle.SystemModal)
            End If
            Me.Close()
        Else
            If ComboBox4.SelectedItem = "1:CSV1" Then
                dg = Form1.CSV_FORMS(0).DataGridView1
            Else
                dg = Form1.CSV_FORMS(1).DataGridView1
            End If
            'For i As Integer = 0 To dg.SelectedCells.Count - 1
            '    dg.SelectedCells.Item(i).Selected = False
            'Next
            Dim num As Integer = 0
            If CheckBox1.Checked = False Then
                For r = sR To dg.RowCount - 1
                    For c = 0 To dg.ColumnCount - 1
                        If dg.Item(c, r).Value <> "" Then
                            If CheckBox3.Checked = True Then
                                If Regex.IsMatch(dg.Item(c, r).Value, SearchChar) Then
                                    dg.Item(c, r).Value = Regex.Replace(dg.Item(c, r).Value, SearchChar, replaceChar)
                                    dg.Item(c, r).Selected = True
                                    dg.CurrentCell = dg(c, r)
                                    num += 1
                                End If
                            Else
                                If InStr(dg.Item(c, r).Value, SearchChar) > 0 Then
                                    dg.Item(c, r).Value = Replace(dg.Item(c, r).Value, SearchChar, replaceChar)
                                    dg.Item(c, r).Selected = True
                                    dg.CurrentCell = dg(c, r)
                                    num += 1
                                End If
                            End If
                        End If
                    Next
                Next
            Else
                For i As Integer = 0 To dg.SelectedCells.Count - 1
                    Try
                        If CheckBox3.Checked = True Then
                            If Regex.IsMatch(dg.SelectedCells(i).Value, SearchChar) Then
                                dg.SelectedCells(i).Value = Regex.Replace(dg.SelectedCells(i).Value, SearchChar, replaceChar)
                                num += 1
                            End If
                        Else
                            If InStr(dg.SelectedCells(i).Value, SearchChar) > 0 Then
                                dg.SelectedCells(i).Value = Replace(dg.SelectedCells(i).Value, SearchChar, replaceChar)
                                num += 1
                            End If
                        End If
                    Catch ex As Exception

                    End Try
                Next
            End If
            sC = 0
            sR = 0
            If num <> 0 Then
                MsgBox(num & "件置換しました", MsgBoxStyle.SystemModal)
            Else
                MsgBox("指定の文字は見つかりませんでした", MsgBoxStyle.SystemModal)
            End If
        End If
    End Sub

    Private Function SearchList(ByVal mode As Integer, ByVal str As String)
        If mode = 1 Then
            str = Replace(str, "\\", "\")
        ElseIf mode = 0 Then
            str = Replace(str, "\\", "[\]")
            str = Replace(str, "\", "\\")
            str = Replace(str, "[\]", "\\")
        End If
        str = Replace(str, "[\t]", vbTab)
        str = Replace(str, "[\n]", vbCrLf)
        str = Replace(str, "\path", "^(.*/).*$")

        If str Is Nothing Then
            str = ""
        End If
        Return str
    End Function

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        start = 0
        Button1.Text = "検　索"
        Me.Visible = False
    End Sub

    Private Sub Search_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If e.CloseReason = CloseReason.FormOwnerClosing Or e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
            start = 0
            Button1.Text = "検　索"
            Me.Visible = False
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim div As String = ""
        If RadioButton1.Checked = True Then
            div = vbTab
        ElseIf RadioButton2.Checked = True Then
            div = ";"
        ElseIf RadioButton3.Checked = True Then
            div = ","
        ElseIf RadioButton4.Checked = True Then
            div = " "
        ElseIf RadioButton5.Checked = True Then
            If TextBox4.Text = "" Then
                MsgBox("区切り文字が入力されていません", MsgBoxStyle.SystemModal)
                Exit Sub
            End If
            div = TextBox4.Text
        End If

        For i As Integer = 0 To dg.SelectedCells.Count - 1
            If dg.SelectedCells(i).Value <> "" Then
                Dim array As String() = Split(dg.SelectedCells(i).Value, div)
                Dim colplus As Integer = (dg.SelectedCells(i).ColumnIndex + array.Length) - dg.ColumnCount
                If colplus > 0 Then
                    For c As Integer = 0 To colplus - 1
                        dg.Columns.Add("column" & dg.Columns.Count, dg.Columns.Count)
                    Next
                End If
                Dim cc As Integer = dg.SelectedCells(i).ColumnIndex
                Dim rr As Integer = dg.SelectedCells(i).RowIndex
                For k As Integer = 0 To array.Length - 1
                    dg.Item(cc + k, rr).Value = array(k)
                Next
            End If
        Next
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Me.Visible = False
    End Sub

    Private Sub TextBox1_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        ComboBox2.SelectAll()
    End Sub

    Private Sub TextBox2_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        ComboBox3.SelectAll()
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        Dim mode As String = ComboBox4.SelectedItem
        If InStr(mode, "HTML") > 0 Then
            RadioButton11.Enabled = False
            RadioButton12.Enabled = False
            RadioButton10.Checked = True
        Else
            RadioButton11.Enabled = True
            RadioButton12.Enabled = True
        End If
    End Sub

    '検索保存
    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim sArray As ArrayList = New ArrayList
        Using sr As New StreamReader(sPath, System.Text.Encoding.Default)
            Do While Not sr.EndOfStream
                Dim s As String = sr.ReadLine
                sArray.Add(s)
            Loop
        End Using

        Dim str As String = ComboBox2.Text & "{,}" & ComboBox3.Text
        If Not sArray.Contains(str) Then
            sArray.Add(str)
            Dim res As String() = DirectCast(sArray.ToArray(GetType(String)), String())
            File.WriteAllLines(sPath, res)
        End If

        CB1_READ()
        MsgBox("保存しました", MsgBoxStyle.SystemModal)
    End Sub

    Private Sub CB1_READ()
        ComboBox1.Items.Clear()
        ComboBox1.Items.Add("検索保存リスト")
        Using sr As New StreamReader(sPath, System.Text.Encoding.Default)
            Do While Not sr.EndOfStream
                Dim s As String = sr.ReadLine
                ComboBox1.Items.Add(s)
            Loop
        End Using
        ComboBox1.SelectedIndex = 0
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim strArray As String() = Split(ComboBox1.SelectedItem, "{,}")
        If strArray.Length > 1 Then
            ComboBox2.Text = strArray(0)
            ComboBox3.Text = strArray(1)
        End If
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Dim sArray As ArrayList = New ArrayList
        Using sr As New StreamReader(sPath, System.Text.Encoding.Default)
            Do While Not sr.EndOfStream
                Dim s As String = sr.ReadLine
                If ComboBox1.SelectedItem <> s Then
                    sArray.Add(s)
                End If
            Loop
        End Using
        Dim res As String() = DirectCast(sArray.ToArray(GetType(String)), String())
        File.WriteAllLines(sPath, res)

        CB1_READ()
        MsgBox("削除しました", MsgBoxStyle.SystemModal)
    End Sub


    Private Sub RadioButton13_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton13.CheckedChanged
        If RadioButton13.Checked = True Then
            ComboBox2.Enabled = True
            TextBox3.Enabled = False
        End If
    End Sub

    Private Sub RadioButton14_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton14.CheckedChanged
        If RadioButton14.Checked = True Then
            ComboBox2.Enabled = False
            TextBox3.Enabled = True
        End If
    End Sub

    Private Sub RadioButton10_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton10.CheckedChanged
        If RadioButton10.Checked = True Then
            ComboBox3.Enabled = True
            TextBox5.Enabled = False
            TextBox1.Enabled = False
            TextBox2.Enabled = False
        End If
    End Sub

    Private Sub RadioButton15_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton15.CheckedChanged
        If RadioButton15.Checked = True Then
            ComboBox3.Enabled = False
            TextBox5.Enabled = True
            TextBox1.Enabled = False
            TextBox2.Enabled = False
        End If
    End Sub

    Private Sub RadioButton11_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton11.CheckedChanged
        If RadioButton11.Checked = True Then
            ComboBox3.Enabled = False
            TextBox5.Enabled = False
            TextBox1.Enabled = True
            TextBox2.Enabled = False
        End If
    End Sub

    Private Sub RadioButton12_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton12.CheckedChanged
        If RadioButton12.Checked = True Then
            ComboBox3.Enabled = False
            TextBox5.Enabled = False
            TextBox1.Enabled = False
            TextBox2.Enabled = True
        End If
    End Sub

    Private Sub TextBox_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles _
            TextBox3.KeyUp, TextBox5.KeyUp
        Dim tb As TextBox = sender
        If e.Modifiers = Keys.Control And e.KeyCode = Keys.A Then
            tb.SelectAll()
        End If
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim mode As String = ComboBox4.SelectedItem
        If mode = "1:CSV1" Then
            dg = Form1.CSV_FORMS(0).DataGridView1
        Else
            dg = Form1.CSV_FORMS(1).DataGridView1
        End If

        ListBox1.Items.Clear()
        For c As Integer = 0 To dg.ColumnCount - 1
            ListBox1.Items.Add(dg.Item(c, 0).Value)
        Next
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Dim mode As String = ComboBox4.SelectedItem
        If mode = "1:CSV1" Then
            dg = Form1.CSV_FORMS(0).DataGridView1
        Else
            dg = Form1.CSV_FORMS(1).DataGridView1
        End If
        Dim dH As ArrayList = TM_HEADER_1ROW_GET(dg)

        For i As Integer = 0 To dg.SelectedCells.Count - 1
            For k As Integer = 0 To ListBox1.SelectedItems.Count - 1
                Dim str As String = TextBox6.Text
                Dim koumoku As String = dg.Item(TM_ArIndexof(dH, ListBox1.SelectedItems(k)), dg.SelectedCells(i).RowIndex).Value
                If koumoku <> "" Then
                    Str = Replace(Str, "_AA_", koumoku)
                    dg.SelectedCells(i).Value = Replace(dg.SelectedCells(i).Value, TextBox7.Text, str & TextBox7.Text)
                End If
            Next
        Next
    End Sub


End Class