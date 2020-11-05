Imports System.IO
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic.FileIO

Public Class Form1_F_sousa


    Private Sub Form1_F_sousa_Load(sender As Object, e As EventArgs) Handles Me.Load
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0

        Dim dgvArray As DataGridView() = {DGV1}
        TM_DGV_CHIRATUKI(dgvArray)


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DGV1.Rows.Clear()

        Dim pattern As String = ""
        Dim sitei As String = TextBox1.Text
        If ComboBox1.SelectedItem = "id" Then
            pattern = "id=.{0,1}" & sitei & ".*?[ |>]"
        ElseIf ComboBox1.SelectedItem = "class" Then
            pattern = "class=.{0,1}" & sitei & ".*?[ |>]"
        Else
            pattern = "name=.{0,1}" & sitei & ".*?[ |>]"
        End If

        Dim html As String = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.Body.OuterHtml
        Dim mc As MatchCollection = Regex.Matches(html, pattern, RegexOptions.IgnoreCase)
        For Each Element As Match In mc
            If Element.Value <> "" Then
                Dim res As String = Element.Value
                res = Replace(res, """", "")
                res = Replace(res, "'", "")
                res = Regex.Replace(res, "<|>", "")
                Dim resArray As String() = Split(res, "=")
                If resArray.Length > 1 Then
                    If Regex.IsMatch(resArray(1), "\s|　") Then
                        Dim res2Array As String() = Regex.Split(resArray(1), "\s|　")
                        For i As Integer = 0 To res2Array.Length - 1
                            If InStr(res2Array(i), sitei) > 0 Then
                                DGV1.Rows.Add(res2Array(i))
                                Exit For
                            End If
                        Next
                    Else
                        DGV1.Rows.Add(resArray(1))
                    End If
                End If
            End If
        Next
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            'Dim doc As HtmlAgilityPack.HtmlDocument = New HtmlAgilityPack.HtmlDocument
            'doc.LoadHtml(Form1.TabBrowser1.SelectedTab.WebBrowser.Document.Body.OuterHtml)

            For r As Integer = 0 To DGV1.RowCount - 1
                Dim form As String = ComboBox1.SelectedItem
                Dim sitei As String = DGV1.Item(0, r).Value
                Dim v As String = DGV1.Item(1, r).Value
                If v Is Nothing Then
                    v = ""
                End If

                If form = "name" Then
                    Form1.TabBrowser1.SelectedTab.WebBrowser.Document.All.GetElementsByName(sitei)(0).SetAttribute("value", v)
                ElseIf form = "id" Then
                    Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementById(sitei).SetAttribute("value", v)
                ElseIf form = "class" Then
                    Dim hc As HtmlElementCollection = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName(ComboBox2.SelectedItem)
                    For Each h As HtmlElement In hc
                        If InStr(h.InnerHtml, sitei) > 0 Then
                            h.SetAttribute("value", v)
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        If LinkLabel1.Text = "目印で調べる" Then
            SplitContainer1.SplitterDistance = 205
            LinkLabel1.Text = "隠す"
        Else
            SplitContainer1.SplitterDistance = 329
            LinkLabel1.Text = "目印で調べる"
        End If
    End Sub

    '調べる
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim pattern1 As String = "id=.{0,1}(.*?)[\s|>|\""|']"
        Dim pattern2 As String = "name=.{0,1}(.*?)[\s|>|\""|']"
        Dim pattern3 As String = "class=.{0,1}(.*?)[\s|>|\""|']"
        Dim marker As String = TextBox2.Text
        'Dim html As String = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.Body.OuterHtml
        'Dim mc As MatchCollection = Regex.Matches(html, marker)

        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""

        Dim elementArray As New ArrayList
        For Each Element As HtmlElement In Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")
            elementArray.Add(Element.OuterHtml)
        Next

        For i As Integer = 0 To elementArray.Count - 1
            Dim res As String = elementArray(i).ToString
            If InStr(res, marker) > 0 Then
                TextBox6.Text = res
                Dim pm1 As Match = Regex.Match(res, pattern1)
                While pm1.Success
                    TextBox3.Text = pm1.Groups(1).Value
                    Exit While
                End While
                Dim pm2 As Match = Regex.Match(res, pattern2)
                While pm2.Success
                    TextBox4.Text = pm2.Groups(1).Value
                    Exit While
                End While
                Dim pm3 As Match = Regex.Match(res, pattern3)
                While pm3.Success
                    TextBox5.Text = pm3.Groups(1).Value
                    Exit While
                End While
                Exit For
            End If
        Next
    End Sub

    'ヤフオク画面検索
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        DGV2.Rows.Clear()

        Dim trArray As New ArrayList
        For Each Element As HtmlElement In Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("tr")
            trArray.Add(Element)
        Next
        Dim tdArray As New ArrayList
        For Each Element As HtmlElement In Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("td")
            tdArray.Add(Element.InnerText)
        Next

        For Each Element As HtmlElement In trArray
            Dim html As String = Element.OuterHtml
            If html <> "" Then
                html = Regex.Replace(html, "\n|\r|""""|'", "")
                Dim pArray As String() = Split(TextBox8.Text, ",")
                Dim flag As Boolean = True
                For i As Integer = 0 To pArray.Length - 1
                    If InStr(html, pArray(i)) = 0 Then
                        flag = False
                        Exit For
                    End If
                Next
                If flag = True Then
                    If InStr(html, TextBox10.Text) = 0 Then
                        Dim searchArray As String() = Split(TextBox9.Text, ",")
                        Dim searchStr As String = ""
                        Dim m As MatchCollection = Regex.Matches(html, "<small>.*?<\/small>")
                        For i As Integer = 0 To searchArray.Length - 1
                            Dim smFlag As Boolean = False
                            For k As Integer = 0 To m.Count - 1
                                Dim mStr As String = m(k).Value
                                If InStr(mStr, searchArray(i)) > 0 Then
                                    If smFlag = False Then
                                        smFlag = True
                                    End If
                                Else
                                    If smFlag = True Then
                                        If mStr <> "" Then
                                            mStr = Regex.Replace(mStr, "（.*?）", "")
                                            mStr = Regex.Replace(mStr, "<.*?>", "")
                                            mStr = Regex.Replace(mStr, "\s|　", "")
                                            mStr = Replace(mStr, "&nbsp;", "")
                                        Else
                                            mStr = ""
                                        End If
                                        If i = 0 Then
                                            searchStr = mStr
                                        Else
                                            searchStr &= "," & mStr
                                        End If
                                        Exit For
                                    End If
                                End If
                            Next
                        Next

                        DGV2.Rows.Add(searchStr, html)
                    End If
                End If
            End If
        Next

        TextBox16_Change()
    End Sub

    'チェック選択
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox7.Text = "" Then
            Exit Sub
        End If

        Dim cCount As Integer = 0
        Dim noAr As String() = Split(TextBox7.Text, vbCrLf)
        Dim noArray As New ArrayList(noAr)
        Dim endArray As New ArrayList
        For i As Integer = noArray.Count - 1 To 0 Step -1
            If noArray(i) <> "" Then
                Dim aucID As String = noArray(i)
                If InStr(noArray(i), "-") > 0 Then
                    Dim aucID1 As String = Split(aucID, "-")(1)
                    Dim aucID2 As String = Split(aucID, "-")(1)
                    For r As Integer = 0 To DGV2.RowCount - 1
                        If InStr(DGV2.Item(0, r).Value, aucID1) > 0 And InStr(DGV2.Item(0, r).Value, aucID2) > 0 Then
                            DGV2.Item(0, r).Style.BackColor = Color.Yellow
                            Dim hc As HtmlElementCollection = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.All.GetElementsByName("item")
                            hc(r).SetAttribute("CHECKED", "1")
                            Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementById("td1_" & r).SetAttribute("bgcolor", "yellow")
                            Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementById("td2_" & r).SetAttribute("bgcolor", "yellow")
                            endArray.Add(noArray(i))
                            noArray.RemoveAt(i)
                            cCount += 1
                            Exit For
                        End If
                    Next
                Else
                    For r As Integer = 0 To DGV2.RowCount - 1
                        If InStr(DGV2.Item(0, r).Value, aucID) > 0 Then
                            DGV2.Item(0, r).Style.BackColor = Color.Yellow
                            Dim hc As HtmlElementCollection = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.All.GetElementsByName("item")
                            hc(r).SetAttribute("CHECKED", "1")
                            Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementById("td1_" & r).SetAttribute("bgcolor", "yellow")
                            Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementById("td2_" & r).SetAttribute("bgcolor", "yellow")
                            endArray.Add(noArray(i))
                            noArray.RemoveAt(i)
                            cCount += 1
                            Exit For
                        End If
                    Next
                End If
            End If
        Next

        TextBox7.Text = ""
        For Each str As String In noArray
            TextBox7.Text &= str & vbCrLf
        Next
        For Each str As String In endArray
            TextBox11.Text &= str & vbCrLf
        Next

        TextBox16_Change()
        MsgBox(cCount & "件、選択完了しました")
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        TextBox7.Text = ""
        TextBox11.Text = ""
        TextBox16_Change()
    End Sub

    'メールディーラーの評価検索
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        DGV3.Rows.Clear()

        Dim divArray As New ArrayList
        Dim hw As HtmlWindowCollection = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.Window.Frames
        For i As Integer = 0 To hw.Count - 2
            For Each Element As HtmlElement In hw(i).Document.GetElementsByTagName("td")
                If InStr(Element.InnerHtml, TextBox14.Text) > 0 Then
                    If Regex.IsMatch(Element.InnerHtml, "^<div") Then
                        divArray.Add(Element)
                    End If
                End If
                'MsgBox(Element.InnerText)
            Next
        Next

        For Each Element As HtmlElement In divArray
            Dim html As String = Element.InnerText
            If html <> "" Then
                DGV3.Rows.Add()
                Dim newNum As Integer = DGV3.RowCount - 1

                html = Regex.Replace(html, "\n|\r|""""|'|　| ", "")

                Dim id As String = ""
                Dim m As MatchCollection = Regex.Matches(html, TextBox13.Text)
                If m.Count > 0 Then
                    id = m(0).Groups(1).Value
                End If

                m = Regex.Matches(html, TextBox12.Text)
                If m.Count > 0 Then
                    DGV3.Item(0, newNum).Value = id & "-" & m(0).Groups(1).Value
                End If

                m = Regex.Matches(html, TextBox15.Text)
                If m.Count > 0 Then
                    DGV3.Item(1, newNum).Value = m(0).Groups(1).Value
                End If
            End If
        Next
    End Sub

    Private Sub TextBox7_KeyDown(sender As TextBox, e As KeyEventArgs) Handles TextBox7.KeyDown, TextBox11.KeyDown
        If e.Modifiers = Keys.Control And e.KeyCode = Keys.A Then
            sender.SelectAll()
        End If
    End Sub

    Private Sub TextBox16_Change()
        Dim str As String = ""
        Dim dgv2Count As Integer = 0
        For r As Integer = 0 To DGV2.RowCount - 1
            If DGV2.Item(0, r).Style.BackColor = Color.Yellow Then
                dgv2Count += 1
            End If
        Next
        Dim txt7Count As Integer = 0
        Dim txt7Lines As String() = Split(TextBox7.Text, vbCrLf)
        For i As Integer = 0 To txt7Lines.Length - 1
            If txt7Lines(i) <> "" Then
                txt7Count += 1
            End If
        Next
        Dim txt11Count As Integer = 0
        Dim txt11Lines As String() = Split(TextBox11.Text, vbCrLf)
        For i As Integer = 0 To txt11Lines.Length - 1
            If txt11Lines(i) <> "" Then
                txt11Count += 1
            End If
        Next

        str = "ヤフオク画面行：" & DGV2.RowCount & " / 選択中：" & dgv2Count & vbCrLf
        str &= "取引番号：" & txt7Count & " (A)" & vbCrLf
        str &= "選択終了分：" & txt11Count & " (B)"
        str &= " (A)+(B)：" & txt7Count + txt11Count & vbCrLf

        TextBox16.Text = str
    End Sub


#Region "DGV規定"
    '####################################################################################
    '####################################################################################
    '
    'DGV規定
    '
    '####################################################################################
    '####################################################################################
    '行番号を表示する
    Private Sub DataGridView_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles _
            DGV1.RowPostPaint
        Dim rect As New Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, sender.RowHeadersWidth - 4, sender.Rows(e.RowIndex).Height)
        TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), sender.RowHeadersDefaultCellStyle.Font, rect, sender.RowHeadersDefaultCellStyle.ForeColor, TextFormatFlags.VerticalCenter Or TextFormatFlags.Right)
    End Sub

    Public FocusDGV As DataGridView = DGV1
    Private Sub DataGridView_GotFocus(sender As DataGridView, e As EventArgs) Handles _
            DGV1.GotFocus, DGV3.GotFocus
        If sender Is Nothing Then
            FocusDGV = DGV1
        Else
            FocusDGV = sender
        End If
    End Sub

    '修正履歴取得
    Private CellErrFlg As Boolean = False
    Private CellValue
    Private Sub DataGridView_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles _
        DGV1.DataError, DGV3.DataError
        CellErrFlg = True
    End Sub

    Private Sub DataGridView_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles _
            DGV1.CellBeginEdit, DGV3.CellBeginEdit
        Dim dgv As DataGridView = sender
        CellErrFlg = False
        CellValue = dgv.CurrentCell.Value   '該当セルの値を格納しておく
    End Sub

    Private Sub DataGridView_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles _
            DGV1.CellEndEdit, DGV3.CellEndEdit
        Dim dgv As DataGridView = sender

        '変更したセルに色をつける
        dgv.CurrentCell.Style.BackColor = Color.Yellow
    End Sub

    Dim selChangeFlag As Boolean = True

    '****************************************************************************************************
    ' DataGridView選択
    '****************************************************************************************************
    Private Sub 切り取りToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles _
        切り取りToolStripMenuItem1.Click
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        CUTS(FocusDGV)
    End Sub

    Private Sub コピーToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles _
        コピーToolStripMenuItem1.Click
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        COPYS(FocusDGV, 0)
    End Sub

    '「"」で囲んでコピー
    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles _
        ToolStripMenuItem1.Click
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        COPYS(FocusDGV, 1)
    End Sub

    Private Sub 貼り付けToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles _
        貼り付けToolStripMenuItem1.Click
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        PASTES(FocusDGV, selCell)
    End Sub

    '子項目をクリックした時、SourceControlがnullになるので先に取得
    Private contextMenuSourceControl As Control = Nothing
    Private Sub ContextMenuStrip1_Opened(sender As Object, e As EventArgs) Handles _
            ContextMenuStrip1.Opened
        Dim menu As ContextMenuStrip = DirectCast(sender, ContextMenuStrip)
        'SourceControlプロパティの内容を覚えておく
        contextMenuSourceControl = menu.SourceControl
    End Sub

    Private Sub 上に挿入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles _
        上に挿入ToolStripMenuItem.Click
        If Not contextMenuSourceControl Is Nothing Then
            FocusDGV = contextMenuSourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        ROWSADD(FocusDGV, selCell, 1)
    End Sub

    Private Sub 下に挿入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles _
        下に挿入ToolStripMenuItem.Click
        If Not contextMenuSourceControl Is Nothing Then
            FocusDGV = contextMenuSourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        ROWSADD(FocusDGV, selCell, 0)
    End Sub

    Private Sub 行を選択直下に複製ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles _
        行を選択直下に複製ToolStripMenuItem.Click
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim dgRow As ArrayList = New ArrayList
        For i As Integer = 0 To FocusDGV.SelectedCells.Count - 1
            If dgRow.Contains(FocusDGV.SelectedCells(i).RowIndex) = False Then
                dgRow.Add(FocusDGV.SelectedCells(i).RowIndex)
            End If
        Next
        dgRow.Sort()
        dgRow.Reverse()

        For r As Integer = 0 To dgRow.Count - 1
            Dim addRow As Integer = dgRow(0) + r + 1
            FocusDGV.Rows.InsertCopy(dgRow(r), addRow)
            For c As Integer = 0 To FocusDGV.ColumnCount - 1
                If FocusDGV.Item(c, dgRow(r)).Value <> "" Then
                    FocusDGV.Item(c, addRow).Value = FocusDGV.Item(c, dgRow(r)).Value
                End If
            Next
        Next
    End Sub

    Private Sub 右に挿入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles _
        右に挿入ToolStripMenuItem.Click
        If Not contextMenuSourceControl Is Nothing Then
            FocusDGV = contextMenuSourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        COLSADD(FocusDGV, selCell, 0)
    End Sub

    Private Sub 左に挿入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles _
        左に挿入ToolStripMenuItem.Click
        If Not contextMenuSourceControl Is Nothing Then
            FocusDGV = contextMenuSourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        COLSADD(FocusDGV, selCell, 1)
    End Sub

    Private Sub 削除ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles _
        削除ToolStripMenuItem1.Click
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        DELS(FocusDGV, selCell)
    End Sub

    Private Sub 背景色を消すToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles _
        背景色を消すToolStripMenuItem.Click
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        For i As Integer = 0 To FocusDGV.SelectedCells.Count - 1
            FocusDGV.SelectedCells(i).Style.BackColor = DefaultBackColor
        Next
    End Sub

    Private Sub 列選択ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles _
        列選択ToolStripMenuItem1.Click
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        selChangeFlag = False
        ColSelect(FocusDGV, 0)
        selChangeFlag = True
    End Sub

    Private Sub 行を削除ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles _
        行を削除ToolStripMenuItem1.Click
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        ROWSCUT(FocusDGV, selCell)
    End Sub

    Private Sub 列を削除ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles _
        列を削除ToolStripMenuItem1.Click
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        COLSCUT(FocusDGV, selCell)
    End Sub

    ' ----------------------------------------------------------------------------------
    ' DataGridViewでCtrl-Vキー押下時にクリップボードから貼り付けるための実装↓↓↓↓↓↓
    ' DataGridViewでDelやBackspaceキー押下時にセルの内容を消去するための実装↓↓↓↓↓↓
    ' ----------------------------------------------------------------------------------
    Public Shared ColumnChars() As String
    Private _editingColumn As Integer   '編集中のカラム番号
    Private _editingCtrl As DataGridViewTextBoxEditingControl
    Public Shared innerTextBox As TextBox

    Private Sub DataGridView_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles _
        DGV1.EditingControlShowing, DGV3.EditingControlShowing
        If TypeOf e.Control Is DataGridViewTextBoxEditingControl Then
            Dim dgv As DataGridView = CType(sender, DataGridView)
            innerTextBox = CType(e.Control, TextBox)    '編集のために表示されているコントロールを取得
        End If
        ' 編集中のカラム番号を保存
        _editingColumn = CType(sender, DataGridView).CurrentCellAddress.X
        Try
            ' 編集中のTextBoxEditingControlにKeyPressイベント設定
            _editingCtrl = CType(e.Control, DataGridViewTextBoxEditingControl)
            AddHandler _editingCtrl.KeyPress, AddressOf DataGridView_CellKeyPress
        Catch
        End Try
    End Sub

    Private Sub DataGridView_CellKeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs)
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

    Private Sub DataGridView_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles _
        DGV1.KeyUp, DGV3.KeyUp
        Dim dgv As DataGridView = CType(sender, DataGridView)
        Dim selCell = dgv.SelectedCells

        If e.KeyCode = Keys.Back Then    ' セルの内容を消去
            DELS(dgv, selCell)
        ElseIf e.Modifiers = Keys.Control And e.KeyCode = Keys.Enter Then '複数セル一括入力
            Dim str As String = dgv.CurrentCell.Value
            For i As Integer = 0 To selCell.Count - 1
                selCell(i).Value = str
            Next
        End If
    End Sub
    ' ----------------------------------------------------------------------------------
    ' DataGridViewでCtrl-Vキー押下時にクリップボードから貼り付けるための実装↑↑↑↑↑↑
    ' DataGridViewでDelやBackspaceキー押下時にセルの内容を消去するための実装↑↑↑↑↑↑
    ' ----------------------------------------------------------------------------------

    Private Sub DELS(ByVal dgv As DataGridView, ByVal selCell As Object)
        For i As Integer = 0 To selCell.Count - 1
            dgv(selCell(i).ColumnIndex, selCell(i).RowIndex).Value = ""
        Next
    End Sub

    ''' <summary>
    ''' 切り取り（汎用版）
    ''' </summary>
    ''' <param name="dgv"></param>
    Private Sub CUTS(ByVal dgv As DataGridView)
        If dgv.GetCellCount(DataGridViewElementStates.Selected) > 0 Then
            If dgv.CurrentCell.IsInEditMode = True Then
                If innerTextBox IsNot Nothing Then
                    innerTextBox.Cut()
                End If
            Else
                Try
                    Clipboard.SetText(dgv.GetClipboardContent.GetText)

                    '切り取りなので、選択セルの内容をクリアする
                    For i As Integer = 0 To dgv.SelectedCells.Count - 1
                        Dim c As Integer = dgv.SelectedCells.Item(i).ColumnIndex
                        Dim r As Integer = dgv.SelectedCells.Item(i).RowIndex
                        dgv.Item(c, r).Value = ""
                    Next i
                Catch ex As System.Runtime.InteropServices.ExternalException
                    MessageBox.Show("切り取りできませんでした。")
                End Try
            End If
        End If
    End Sub

    ''' <summary>
    ''' コピー（汎用版）
    ''' </summary>
    ''' <param name="dgv"></param>
    ''' <param name="mode"></param>
    Private Sub COPYS(ByVal dgv As DataGridView, ByVal mode As Integer)
        If dgv.GetCellCount(DataGridViewElementStates.Selected) > 0 Then
            If dgv.CurrentCell.IsInEditMode = True Then
                If innerTextBox IsNot Nothing Then
                    innerTextBox.Copy()
                End If
            Else
                Try
                    If dgv.SelectedCells.Count > 0 Then
                        Dim lstStrBuf As New List(Of String)
                        For Each cell As DataGridViewCell In dgv.SelectedCells
                            If IsDBNull(cell.Value) Then
                                lstStrBuf.Add(Nothing)
                            Else
                                lstStrBuf.Add(cell.Value)
                                'cell.Value = (cell.Value.ToString()).Replace(vbNewLine, vbLf)
                                cell.Value = Replace(cell.Value, vbCrLf, vbLf)
                                If mode = 1 Then    '通常コピーはダブルコーテーションで囲む
                                    cell.Value = Replace(cell.Value, """", """""")
                                    cell.Value = """" & cell.Value & """"
                                End If
                            End If
                        Next

                        'Clipboard.SetText(dgv.GetClipboardContent.GetText)
                        Dim returnValue As DataObject
                        returnValue = dgv.GetClipboardContent()
                        Dim strTab = returnValue.GetText()
                        Clipboard.SetDataObject(strTab, True)

                        Dim i As Integer = 0
                        For Each cell As DataGridViewCell In dgv.SelectedCells
                            cell.Value = lstStrBuf(i)
                            i += 1
                        Next
                    End If
                Catch ex As System.Runtime.InteropServices.ExternalException
                    MessageBox.Show("コピーできませんでした。")
                End Try
            End If
        End If
    End Sub

    ''' <summary>
    ''' 貼り付け（汎用版）
    ''' </summary>
    ''' <param name="dgv"></param>
    ''' <param name="selCell"></param>
    Private Sub PASTES(ByVal dgv As DataGridView, ByVal selCell As Object)
        Dim x As Integer = dgv.CurrentCellAddress.X
        Dim y As Integer = dgv.CurrentCellAddress.Y
        ' クリップボードの内容を取得
        Dim clipText As String = Clipboard.GetText()

        ' 改行を変換
        'clipText = clipText.Replace(vbCrLf, vbLf)
        'clipText = clipText.Replace(vbCr, vbLf)
        ' 改行で分割
        'Dim lines() As String = clipText.Split(vbNewLine)

        '================================================================
        '改行を含むテキストを取得するためTextFieldParserで解析
        Dim csvRecords As New ArrayList()
        'Dim clipText2 As String = Replace(clipText, vbLf, "|=|")
        'Dim clipCSV As String() = Split(clipText2, vbCrLf)

        'For i As Integer = 0 To clipCSV.Length - 1
        '    Dim rs As New System.IO.StringReader(clipText)
        '    Dim tfp As New TextFieldParser(rs) With {
        '            .TextFieldType = FileIO.FieldType.Delimited,
        '            .Delimiters = New String() {vbTab, ","},
        '            .HasFieldsEnclosedInQuotes = True
        '        }

        '    While Not tfp.EndOfData
        '        Dim str As String = ""
        '        Dim fields As String() = tfp.ReadFields()   'フィールドを読み込む

        '        For k As Integer = 0 To fields.Length - 1
        '            If k = 0 Then
        '                str = fields(k)
        '            Else
        '                str &= vbTab & fields(k)
        '            End If
        '        Next
        '        csvRecords.Add(str)
        '    End While

        '    tfp.Close()
        '    rs.Close()
        'Next

        clipText = Replace(clipText, vbCrLf, "<<CRLF>>")
        clipText = Replace(clipText, vbCr, "<<CR>>")
        clipText = Replace(clipText, vbLf, "<<LF>>")
        Dim clipTextArray As String() = Split(clipText, "<<CRLF>>")
        'Dim clipTextArray As String() = Split(clipText, vbCrLf)

        For i As Integer = 0 To clipTextArray.Length - 1
            Dim flag As Boolean = False
            Dim rs As New StringReader(clipTextArray(i))
            Using tfp As New TextFieldParser(rs)
                tfp.TextFieldType = FieldType.Delimited
                tfp.Delimiters = New String() {vbTab} ', ","}
                tfp.HasFieldsEnclosedInQuotes = True
                tfp.TrimWhiteSpace = False

                While Not tfp.EndOfData
                    Dim str As String = ""
                    Dim fields As String() = tfp.ReadFields()   'フィールドを読み込む

                    For k As Integer = 0 To fields.Length - 1
                        If k = 0 Then
                            str = fields(k)
                        Else
                            str &= vbTab & fields(k)
                        End If
                    Next

                    str = Replace(str, "<<CR>>", vbCr)
                    str = Replace(str, "<<LF>>", vbLf)
                    csvRecords.Add("" & str)
                    flag = True
                End While
            End Using
            rs.Close()

            'データが無い行をチェックし追加する
            If flag = False Then
                csvRecords.Add("")
            End If
        Next

        '================================================================

        Dim lines() As String = Nothing
        lines = DirectCast(csvRecords.ToArray(GetType(String)), String())

        Dim oneFlag As Boolean = False
        If lines.Length = 1 Then
            Dim vals() As String = lines(0).Split(vbTab)
            If vals.Length = 1 Then
                oneFlag = True
            End If
        End If

        If oneFlag = True Then
            If dgv.IsCurrentCellInEditMode = False Then
                For i As Integer = 0 To selCell.count - 1
                    dgv(selCell(i).ColumnIndex, selCell(i).RowIndex).Value = lines(0)
                    dgv(selCell(i).ColumnIndex, selCell(i).RowIndex).style.backcolor = Color.Yellow
                Next
            Else
                'datagridviewのセルをテキストボックスとして扱い
                '選択文字列を置き換えればpasteになる
                Dim TextEditCtrl As DataGridViewTextBoxEditingControl = DirectCast(dgv.EditingControl, DataGridViewTextBoxEditingControl)
                TextEditCtrl.SelectedText = lines(0)
                'dgv(selCell(0).ColumnIndex, selCell(0).RowIndex).Value = Replace(dgv(selCell(0).ColumnIndex, selCell(0).RowIndex).Value, TextEditCtrl.SelectedText, lines(0))
                'dgv.EndEdit()
            End If
        Else
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
                        DataGridView_CellKeyPress(dgv, tmpe)
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
                    dgv(x + c2, y + r).Style.BackColor = Color.Yellow
                    ' ------------------------------------------------------
                    ' 次のセルへ
                    c2 = c2 + 1
                Next
            Next
        End If

    End Sub

    Private Sub ROWSADD(ByVal dgv As DataGridView, ByVal selCell As Object, ByVal mode As Integer)
        Dim delR As New ArrayList
        For i As Integer = 0 To selCell.Count - 1
            If Not delR.Contains(selCell(i).RowIndex) Then
                delR.Add(selCell(i).RowIndex)
            End If
        Next
        delR.Sort()
        If mode = 1 Then
            dgv.Rows.Insert(delR(0), delR.Count)
        Else
            delR.Reverse()
            If delR(0) + 1 >= dgv.RowCount - 1 Then
                dgv.Rows.Add(delR.Count)
            Else
                delR(0) += 1
                dgv.Rows.Insert(delR(0), delR.Count)
            End If
        End If
    End Sub

    Private Sub COLSADD(ByVal dgv As DataGridView, ByVal selCell As Object, ByVal mode As Integer)
        Dim delR As New ArrayList
        For i As Integer = 0 To selCell.Count - 1
            If Not delR.Contains(selCell(i).ColumnIndex) Then
                delR.Add(selCell(i).ColumnIndex)
            End If
        Next
        delR.Sort()
        Dim num As String = dgv.ColumnCount
        Dim textColumn As New DataGridViewTextBoxColumn With {
            .Name = "Column" & num,
            .HeaderText = num
        }
        If mode = 1 Then
            dgv.Columns.Insert(delR(0), textColumn)
        Else
            delR.Reverse()
            delR(0) += 1
            If delR(0) + 1 <= dgv.ColumnCount - 1 Then
                dgv.Columns.Insert(delR(0), textColumn)
            Else
                dgv.Columns.Add(textColumn)
            End If
        End If
    End Sub

    ''' <summary>
    ''' 列選択（汎用版）
    ''' </summary>
    ''' <param name="dgv"></param>
    ''' <param name="mode">mode=0 1行目を選択する、mode=1 1行目を選択除外</param>
    Private Sub ColSelect(dgv As DataGridView, mode As Integer)
        Dim selCell = dgv.SelectedCells
        Dim selCol As ArrayList = New ArrayList
        For i As Integer = 0 To selCell.Count - 1
            If Not selCol.Contains(selCell(i).ColumnIndex) Then
                selCol.Add(selCell(i).ColumnIndex)
            End If
        Next

        If mode = 0 Then
            For i As Integer = 0 To selCol.Count - 1
                For r As Integer = 0 To dgv.RowCount - 1
                    If Not dgv.Rows(r).IsNewRow Then
                        dgv.Item(selCol(i), r).selected = True
                    End If
                Next
            Next
        Else
            For i As Integer = 0 To selCol.Count - 1
                For r As Integer = 1 To dgv.RowCount - 1
                    If Not dgv.Rows(r).IsNewRow Then
                        dgv.Item(selCol(i), r).selected = True
                    End If
                Next
            Next
        End If
    End Sub

    Private Sub ROWSCUT(ByVal dgv As DataGridView, ByVal selCell As Object)
        Dim delR As New ArrayList
        For i As Integer = 0 To selCell.Count - 1
            If Not delR.Contains(selCell(i).RowIndex) Then
                delR.Add(selCell(i).RowIndex)
            End If
        Next
        delR.Sort()
        'delR.Reverse()

        'データ処理を軽くするため画面を一時的に消す
        If delR.Count > 100 Then
            dgv.Visible = False
            Application.DoEvents()
        End If

        For r As Integer = delR.Count - 1 To 0 Step -1
            If Not dgv.Rows(delR(r)).IsNewRow Then
                dgv.Rows.RemoveAt(delR(r))
            End If
        Next
        If delR.Count > 100 Then
            dgv.Visible = True
        End If
    End Sub

    Private Sub COLSCUT(ByVal dgv As DataGridView, ByVal selCell As Object)
        Dim delC As New ArrayList
        For i As Integer = 0 To selCell.Count - 1
            If Not delC.Contains(selCell(i).ColumnIndex) Then
                delC.Add(selCell(i).ColumnIndex)
            End If
        Next
        delC.Sort()
        delC.Reverse()
        For c As Integer = 0 To delC.Count - 1
            If delC(c) < dgv.ColumnCount Then
                dgv.Columns.RemoveAt(delC(c))
            End If
        Next
    End Sub

    '####################################################################################
    '####################################################################################
    '
    'DGV規定ここまで
    '
    '####################################################################################
    '####################################################################################
#End Region


End Class