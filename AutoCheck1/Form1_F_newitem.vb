Imports System.Data.OleDb
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports Microsoft.International.Converters
Imports ZetaHtmlTidy
Imports System.Text

Public Class Form1_F_Newitem

    Dim ENC_SJ As Encoding = Encoding.GetEncoding("SHIFT-JIS")

    Private Sub Newitem_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        TemplateFolder()

        '配送予定日_Y_（納期情報選択肢）
        Dim cb3List As String() = File.ReadAllLines(Path.GetDirectoryName(Form1.appPath) & "\config\newitem_cb3.txt", ENC_SJ)
        For i As Integer = 0 To cb3List.Length - 1
            If cb3List(i) <> "" Then
                ComboBox3.Items.Add(cb3List(i))
            End If
        Next
    End Sub

    Private Sub Newitem_FormClosing(ByVal sender As Object, ByVal e As Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If e.CloseReason = CloseReason.FormOwnerClosing Or e.CloseReason = CloseReason.UserClosing Then
            Me.Visible = False
            e.Cancel = True
        End If
    End Sub

    Private Sub TemplateFolder()
        Dim di As New DirectoryInfo(Path.GetDirectoryName(Form1.appPath) & "\formchange")
        Dim files As FileInfo() = di.GetFiles("*.txt", SearchOption.AllDirectories)

        ToolStripComboBox1.Items.Clear()
        ToolStripComboBox1.Items.Add("選択してください")

        For Each f As FileInfo In files
            If Regex.IsMatch(f.ToString, "yahoo-|楽天-") Then
                ToolStripComboBox1.Items.Add(Path.GetFileNameWithoutExtension(f.ToString))
            End If
        Next

        ToolStripComboBox1.SelectedIndex = 0
    End Sub

    Private Sub ToolStripComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        If ToolStripComboBox1.SelectedItem <> "選択してください" Then
            Dim folder As String = Path.GetDirectoryName(Form1.appPath) & "\formchange"
            Dim p As String = folder & "\" & ToolStripComboBox1.SelectedItem & ".txt"
            Dim csvRecords As New ArrayList()
            csvRecords = CSV_READ(p)

            DataGridView1.Rows.Clear()
            For r As Integer = 0 To csvRecords.Count - 1
                If Regex.IsMatch(csvRecords(r), "^//") Then
                    Continue For
                End If

                Dim lines As String() = Split(csvRecords(r), "|=|")
                DataGridView1.Rows.Add(lines)
            Next

            If InStr(ToolStripComboBox1.SelectedItem, "yahoo") > 0 Then
                TextBox22.Enabled = False
            Else
                TextBox22.Enabled = True
            End If

            If InStr(ToolStripComboBox1.SelectedItem, "暁") > 0 Then
                'TextBox30.BackColor = Color.PaleGreen
                'TextBox30.Enabled = True
                'TextBox13.Enabled = False
                TextBox2.Enabled = False
                ComboBox4.Enabled = False
                ComboBox4.BackColor = Color.DimGray
            Else
                'TextBox30.BackColor = Color.DimGray
                'TextBox30.Enabled = False
                'TextBox13.Enabled = True
                TextBox2.Enabled = True
                ComboBox4.Enabled = True
                ComboBox4.BackColor = Color.White
            End If
        End If
    End Sub

    Private Function CSV_READ(ByRef path As String)
        Dim csvRecords As New ArrayList()

        'CSVファイル名
        Dim csvFileName As String = path

        'Shift JISで読み込む
        Dim tfp As New FileIO.TextFieldParser(csvFileName, System.Text.Encoding.GetEncoding(932)) With {
            .TextFieldType = FileIO.FieldType.Delimited,
            .Delimiters = New String() {","},
            .HasFieldsEnclosedInQuotes = True,
            .TrimWhiteSpace = True
        }

        While Not tfp.EndOfData
            'フィールドを読み込む
            Dim fields As String() = tfp.ReadFields()
            Dim str As String = ""
            For i As Integer = 0 To fields.Length - 1
                If i = 0 Then
                    str = fields(i)
                Else
                    str &= "|=|" & fields(i)
                End If
            Next
            '保存
            csvRecords.Add(str)
        End While

        '後始末
        tfp.Close()
        Return csvRecords
    End Function

    '全角チェック
    Private Sub DataGridView3_CellValueChanged(ByVal sender As Object, ByVal e As Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellValueChanged
        If e.ColumnIndex >= 0 And e.RowIndex >= 0 Then
            Dim checkStr As String = DataGridView3.Item(e.ColumnIndex, e.RowIndex).Value
            If checkStr <> "" Then
                checkStr = Regex.Replace(checkStr, "[^0-9A-Za-z０-９Ａ-Ｚａ-ｚ]", "")
                If Not (IsOneByteChar(checkStr)) Then
                    DataGridView3.Item(e.ColumnIndex, e.RowIndex).Style.BackColor = Color.Red
                    MsgBox("枝番号（コード）に全角文字が入力されています", MsgBoxStyle.SystemModal)
                Else
                    DataGridView3.Item(e.ColumnIndex, e.RowIndex).Style.BackColor = Color.White
                End If
            Else
                DataGridView3.Item(e.ColumnIndex, e.RowIndex).Style.BackColor = Color.White
            End If
        End If
    End Sub

    '1バイト文字で構成された文字列であるか判定
    Private Function IsOneByteChar(ByVal str As String) As Boolean
        Dim byte_data As Byte() = System.Text.Encoding.GetEncoding(932).GetBytes(str)
        If (byte_data.Length = str.Length) Then
            Return True
        Else
            Return False
        End If
    End Function


    '行番号を表示する
    Private Sub DataGridView1_RowPostPaint(ByVal sender As DataGridView, ByVal e As Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles _
            DataGridView1.RowPostPaint, DataGridView2.RowPostPaint, DataGridView3.RowPostPaint, DGV6.RowPostPaint
        ' 行ヘッダのセル領域を、行番号を描画する長方形とする（ただし右端に4ドットのすき間を空ける）
        Dim rect As New Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, sender.RowHeadersWidth - 4, sender.Rows(e.RowIndex).Height)
        ' 上記の長方形内に行番号を縦方向中央＆右詰で描画する　フォントや色は行ヘッダの既定値を使用する
        TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), sender.RowHeadersDefaultCellStyle.Font, rect, sender.RowHeadersDefaultCellStyle.ForeColor, TextFormatFlags.VerticalCenter Or TextFormatFlags.Right)
    End Sub

    Private Sub 上書き保存ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 上書き保存ToolStripMenuItem.Click
        Dim folder As String = Path.GetDirectoryName(Form1.appPath) & "\formchange"
        Dim filename As String = folder & "\" & ToolStripComboBox1.SelectedItem & ".txt"
        SaveCsv(filename, DataGridView1)
        MsgBox(filename & vbCrLf & "保存しました", MsgBoxStyle.SystemModal)
    End Sub

    Private Sub 名前を付けて保存ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 名前を付けて保存ToolStripMenuItem.Click
        Dim sfd As New SaveFileDialog With {
            .FileName = "テンプレート.txt"
        }
        Dim folder As String = Path.GetDirectoryName(Form1.appPath) & "\formchange"
        sfd.InitialDirectory = folder
        sfd.Filter = "テンプレートファイル(*.txt)|*.txt"
        sfd.FilterIndex = 0
        sfd.Title = "保存先のファイルを選択してください"
        sfd.RestoreDirectory = True
        sfd.OverwritePrompt = True
        sfd.CheckPathExists = True

        'ダイアログを表示する
        If sfd.ShowDialog(Me) = DialogResult.OK Then
            SaveCsv(sfd.FileName, DataGridView1)
            MsgBox(sfd.FileName & vbCrLf & "保存しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    Public Sub SaveCsv(ByVal fp As String, ByVal dg As DataGridView)
        ' CSVファイルオープン
        Dim sw As IO.StreamWriter = New IO.StreamWriter(fp, False, System.Text.Encoding.GetEncoding("SHIFT-JIS"))
        Dim minusRow As Integer = 1
        If dg.AllowUserToAddRows = True Then
            minusRow = 2
        End If

        Dim writeStr As String = ""
        For r As Integer = 0 To dg.Rows.Count - minusRow
            Dim dl As String = ""
            For c As Integer = 0 To dg.Columns.Count - 1
                ' DataGridViewのセルのデータ取得
                Dim dt As String = ""
                dt = dg.Rows(r).Cells(c).Value.ToString()
                dt = Replace(dt, vbCrLf, vbLf)
                'dt = Replace(dt, vbLf, "")
                If Not dt Is Nothing Then
                    dt = Replace(dt, """", """""")
                    Select Case True
                        Case Regex.IsMatch(dt, ".*,.*|.*" & vbLf & ".*|.*" & vbCr & ".*|.*"".*")
                            dt = """" & dt & """"
                    End Select
                End If
                If c < dg.Columns.Count - 1 Then
                    dt = dt & ","
                End If
                dl &= dt
            Next
            writeStr &= dl & vbLf
        Next
        sw.Write(writeStr)
        ' CSVファイルクローズ
        sw.Close()
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Dim errorMsg As String = ""
        Dim searchArray As ArrayList = New ArrayList

        For r As Integer = 0 To DataGridView1.RowCount - 1
            If DataGridView1.Item(0, r).Value <> "" Then
                Dim commandLine As String = ""
                For c As Integer = 0 To DataGridView1.ColumnCount - 1
                    commandLine &= DataGridView1.Item(c, r).Value & "|=|"
                Next

                'Try
                Dim command As String() = Split(commandLine, "|=|")
                Dim form As String = command(1)
                Dim name As String = command(2)
                Dim value As String = command(3)
                Select Case True
                    Case Regex.IsMatch(value, "_H_")
                        Dim selStr As String() = Split(ComboBox1.SelectedItem, ":")
                        value = Replace(value, "_H_", selStr(0))

                        '暁
                        Dim haiso_no As Integer = 0
                        If ToolStripComboBox1.SelectedItem = "楽天-通販の暁用" Then
                            If InStr(selStr(1), "宅配のみ") > 0 Then
                                haiso_no = 2
                                '単品配送設定
                                Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementById("r50").SetAttribute("Checked", "True")
                                Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementById("r51").SetAttribute("Checked", "")
                            ElseIf InStr(selStr(1), "大型のみ") > 0 Then
                                haiso_no = 3
                                '単品配送設定
                                'Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input").GetElementsByName("postage")(0).SetAttribute("text", "1260")

                                'Dim jqueryCode As String = "$('[id='r51']').removeAttr('disabled'); $('label[for='r51']').css({color:'#000000'}); $('[name='single_item_shipping_reason']').removeAttr('disabled');"
                                'Form1.TabBrowser1.SelectedTab.WebBrowser.Document.InvokeScript("eval", New Object() {jqueryCode})
                                'Dim jqueryCode As String = "alert(123);"
                                'Form1.TabBrowser1.SelectedTab.WebBrowser.Document.InvokeScript(jqueryCode)
                                'Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input").GetElementsByName("postage")(0).SetAttribute("text", "123")
                                'Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input").GetElementsByName("postage")(0).RaiseEvent("OnMouseDown")

                                Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementById("r43").InvokeMember("Click")
                                Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementById("r42").InvokeMember("Click")
                                'Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementById("r50").SetAttribute("Checked", "")
                                'Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementById("r51").SetAttribute("disabled", "")
                                'Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementById("r51").SetAttribute("Checked", "True")
                                Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("Select").GetElementsByName("single_item_shipping_reason")(0).SetAttribute("value", 3)
                            ElseIf InStr(selStr(1), "メールのみ") > 0 Then
                                haiso_no = 4
                                '単品配送設定
                                Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementById("r50").SetAttribute("Checked", "True")
                                Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementById("r51").SetAttribute("Checked", "")
                            ElseIf InStr(selStr(1), "定形外郵便") > 0 Then
                                haiso_no = 5
                                '単品配送設定
                                Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementById("r50").SetAttribute("Checked", "True")
                                Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementById("r51").SetAttribute("Checked", "")
                            End If
                            value = Replace(value, "_H_", haiso_no)
                        End If

                    Case Regex.IsMatch(value, "_S_")
                        value = Replace(value, "_S_", TextBox2.Text)
                    Case Regex.IsMatch(value, "_SA_")   '_SA_=0     'searchArrayの呼び出し
                        Dim selStr As String() = Split(value, "=")
                        value = searchArray(selStr(1))
                    Case Regex.IsMatch(value, "_P1_")
                        value = Replace(value, "_P1_", TextBox13.Text)
                    Case Regex.IsMatch(value, "_Y_")
                        Dim selStr As String() = Split(ComboBox3.SelectedItem, ":")
                        value = Replace(value, "_Y_", selStr(0))
                    Case Regex.IsMatch(value, "_W_")
                        Dim selStr As String() = Split(ComboBox4.SelectedItem, ":")
                        value = Replace(value, "_W_", selStr(0))
                        If InStr(value, "/*/") > 0 Then   '暁選択肢用
                            'Dim vA As String() = Split(value, "/*/")
                            'value = CInt(vA(0)) + CInt(vA(1))
                            Exit Select
                        End If
                End Select

                Select Case command(0)
                    Case "置換"
                        If name = "smart_caption" Then
                            Console.WriteLine(123)
                        End If

                        Dim replaceRes As String() = Split(value, "「=」")
                        Dim str As String = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName(form)(name).GetAttribute("value")
                        Select Case True
                            Case Regex.IsMatch(replaceRes(1), "_link_") '画像リンクを置換
                                Dim value2 As String = ""
                                Dim widthNum As String = Replace(replaceRes(1), "_link_", "")
                                Dim Num As String() = Split(widthNum, "-")
                                Dim rowNum As Integer = Num(0)  '開始番号

                                'If InStr(ToolStripComboBox1.Text, "楽天") > 0 Then
                                '    For Each m As Match In Regex.Matches(str, replaceRes(0))
                                '        If rowNum <= Num(1) Then
                                '            If DataGridView2.Item(0, rowNum).Value <> "" Then
                                '                value2 = DataGridView2.Item(0, rowNum).Value
                                '                str = Regex.Replace(str, m.Value, value2)
                                '                rowNum += 1
                                '            End If
                                '        Else
                                '            str = ""
                                '            rowNum += 1
                                '        End If
                                '    Next
                                'Else
                                '新変更ロジック（とりあえずYahoo用）
                                Dim lineArray As String() = Regex.Split(str, vbCrLf & "|" & vbCr & "|" & vbLf)
                                str = ""
                                For Each lA As String In lineArray
                                    If Regex.IsMatch(lA, replaceRes(0)) Then
                                        If rowNum = Num(0) Then
                                            For rr As Integer = 0 To DataGridView2.RowCount - 1
                                                If DataGridView2.Item(0, rr).Value <> "" Then
                                                    value2 = DataGridView2.Item(0, rr).Value
                                                    str &= Regex.Replace(lA, replaceRes(0), value2) & vbCrLf
                                                End If
                                                rowNum += 1
                                            Next
                                        Else
                                            '行削除
                                        End If
                                    Else
                                        str &= lA & vbCrLf
                                    End If
                                Next
                                'End If
                            Case Regex.IsMatch(replaceRes(1), "_splink_")   '楽天スマホ画面用
                                Dim value1 As String = ""
                                Dim widthNum As String = Replace(replaceRes(1), "_splink_", "")
                                Dim Num As String() = Split(widthNum, "-")
                                Dim rowNum As Integer = Num(0)  '開始番号
                                value1 &= "<!--imgS-->" & vbLf
                                value1 &= "<p>" & vbLf
                                For i As Integer = rowNum To DataGridView2.RowCount - 2
                                    If rowNum <= Num(1) Then
                                        value1 &= "<img src='"
                                        value1 &= DataGridView2.Item(0, i).Value
                                        'value1 &= "' alt=' width=' 500'='' width='100%' '=''>" & vbLf
                                        value1 &= " width='100%'>" & vbLf
                                    End If
                                Next
                                value1 &= "</p>" & vbLf & "<br>" & vbLf
                                value1 &= "<!--imgE-->" & vbLf
                                str = Regex.Replace(str, replaceRes(0), value1)
                            Case Else   '文字の置換
                                str = Regex.Replace(str, replaceRes(0), replaceRes(1))
                        End Select
                        Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName(form)(name).SetAttribute("value", str)
                    Case "入力"
                        Dim value2 As String = ""
                        '暁　大型　個別送料
                        'If ToolStripComboBox1.SelectedItem = "楽天-通販の暁用" And ComboBox4.SelectedIndex = 3 And name = "postage" Then
                        '    value2 = 1260
                        'Else
                        If value <> "" Then
                            value = Okikae(value)
                            Select Case True
                                Case Regex.IsMatch(value, "_link_") '画像リンクを取得
                                    Dim rowNum As Integer = Replace(value, "_link_", "")
                                    If rowNum < DataGridView2.RowCount - 1 Then
                                        If DataGridView2.Item(0, rowNum).Value <> "" Then
                                            value2 = DataGridView2.Item(0, rowNum).Value
                                        Else
                                            value2 = ""
                                        End If
                                    End If
                                Case Regex.IsMatch(value, "_code_") 'コード+番号
                                    Dim rowNum As Integer = Replace(value, "_code_", "")
                                    If rowNum < DataGridView2.RowCount - 1 Then
                                        If DataGridView2.Item(0, rowNum).Value <> "" Then
                                            If rowNum = 0 Then
                                                value2 = TextBox3.Text
                                            Else
                                                value2 = TextBox3.Text & TextBox9.Text & rowNum
                                            End If
                                        Else
                                            value2 = ""
                                        End If
                                    End If
                                Case Else
                                    If value = "none" Then
                                        value2 = ""
                                    Else
                                        value2 = value
                                    End If
                            End Select
                        Else
                            value2 = value
                        End If
                        'End If

                        Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName(form)(name).SetAttribute("value", value2)
                        If name = "postage" And ToolStripComboBox1.SelectedItem = "楽天-通販の暁用" And ComboBox1.SelectedItem = "3:大型のみ" Then
                            Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementById("r43").InvokeMember("Click")
                            Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementById("r42").InvokeMember("Click")
                            Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementById("r51").InvokeMember("Click")
                        End If
                    Case "整形"
                        If value = "0" Then     'HTML TIDY
                            Dim str As String = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName(form)(name).GetAttribute("value")
                            Dim res As String = SeikeiHTML(str)
                            Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName(form)(name).SetAttribute("value", res)
                        ElseIf value = "1" Or value = "2" Then '改行チェック
                            Dim str As String = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName(form)(name).GetAttribute("value")
                            str = Replace(str, vbCrLf, vbLf)
                            str = Replace(str, vbCr, vbLf)
                            Dim strArray As String() = Split(str, vbLf)
                            Dim flag As Integer = 0
                            Dim brFlag As Boolean = False
                            Dim res As String = ""
                            For i As Integer = 0 To strArray.Length - 1
                                If value = "2" And strArray(i) = "" Then
                                    If brFlag = False Then
                                        flag += 1
                                        brFlag = True
                                        res &= strArray(i) & vbLf
                                    End If
                                Else
                                    Select Case True
                                        Case Regex.IsMatch(strArray(i), "◆|■")
                                            flag = 1
                                            brFlag = False
                                            res &= strArray(i) & vbLf
                                        Case Regex.IsMatch(strArray(i), "【|○")
                                            If flag = 1 And Regex.IsMatch(strArray(i), "【") Then
                                                If value = "1" Then
                                                    strArray(i) = "<br>" & vbLf & strArray(i)
                                                Else
                                                    strArray(i) = vbLf & strArray(i)
                                                End If
                                                flag += 1
                                            ElseIf flag = 2 And Regex.IsMatch(strArray(i), "○") Then
                                                If value = "1" Then
                                                    strArray(i) = "<br>" & vbLf & strArray(i)
                                                Else
                                                    strArray(i) = vbLf & strArray(i)
                                                End If
                                                flag += 1
                                            End If
                                            brFlag = False
                                            res &= strArray(i) & vbLf
                                        Case Regex.IsMatch(strArray(i), "^<br>$")
                                            If brFlag = False Then
                                                flag += 1
                                                brFlag = True
                                                res &= strArray(i) & vbLf
                                            End If
                                        Case Else
                                            res &= strArray(i) & vbLf
                                    End Select
                                End If
                            Next
                            Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName(form)(name).SetAttribute("value", res)
                        End If
                    Case "選択"
                        If form = "radio" Then
                            Dim hCollection As HtmlElementCollection = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")
                            For Each element As HtmlElement In hCollection
                                If element.GetAttribute("name") = name And element.GetAttribute("value") = value Then
                                    element.SetAttribute("Checked", "True")
                                    Exit For
                                End If
                            Next
                        ElseIf form = "checkbox" Then
                            If value = "True" Then
                                Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input").Item(name).SetAttribute("Checked", value)
                            Else
                                Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input").Item(name).SetAttribute("Checked", "")
                            End If
                        ElseIf form = "select" Then
                            Dim hCollection As HtmlElementCollection = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("select")
                            If IsNumeric(value) Then
                                If ToolStripComboBox1.SelectedItem = "楽天-通販の暁用" And name = "soryo_kbn1" Then
                                    If ComboBox4.SelectedIndex = 3 Then
                                        value = 0
                                    Else
                                        If value = 0 Or value = 3 Then
                                            value = 1
                                        End If
                                    End If
                                End If
                                Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("select")(name).SetAttribute("selectedIndex", value)
                            Else
                                Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("select")(name).SetAttribute("selectedItem", value)
                            End If
                        ElseIf form = "textarea" Then
                            '選択,textarea,フォーム名,_W_|$$|value1||value2
                            Dim hCollection As HtmlElementCollection = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("textarea")
                            Dim valueArray As String() = Split(value, "|$$|")
                            If IsNumeric(valueArray(0)) Then
                                Dim v As String() = Split(valueArray(valueArray(0)), "||")  '_W_の選択している番号取得
                                For Each element As HtmlElement In hCollection
                                    If element.GetAttribute("name") = name Then
                                        Dim str As String = element.GetAttribute("value")
                                        str = Replace(str, v(0), v(1))
                                        element.SetAttribute("value", str)
                                    End If
                                Next
                            End If
                        End If
                    Case "追加"
                        Dim str As String = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName(form)(name).GetAttribute("value")
                        value = Okikae(value)
                        If Regex.IsMatch(value, "_code_") = True Then
                            value = Replace(value, "_code_", TextBox3.Text)
                        End If
                        If InStr(str, "_前_") > 0 Then
                            str = Replace(str, "_前_", "")
                            str = str & value
                        Else
                            str = value & str
                        End If
                        Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName(form)(name).SetAttribute("value", str)
                    Case "処理"
                        If command(3) = "全半" Then
                            Dim str As String = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName(form)(name).GetAttribute("value")
                            str = Form1.StrConvNumeric(str)
                            str = Form1.StrConvEnglish(str)
                            str = Form1.StrConvComma(str)
                            str = Form1.StrConvNumeric(str)
                            str = Form1.StrConvNamisen(str)
                            Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName(form)(name).SetAttribute("value", str)
                        ElseIf command(3) = "ランダム置換" Then
                            Dim str As String = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName(form)(name).GetAttribute("value")
                            str = Replace(str, "　", " ")
                            str = Regex.Replace(str, "\s{1,5}", " ")
                            Dim mae As New ArrayList
                            Dim mc As MatchCollection = Regex.Matches(str, "【.*】\s{0,2}")
                            For Each m As Match In mc
                                mae.Add(m.Value)
                            Next
                            Dim ato As New ArrayList
                            mc = Regex.Matches(str, "\s{0,2}[a-zA-z]{1,2}[0-9]{3}$")
                            For Each m As Match In mc
                                ato.Add(m.Value)
                            Next
                            For k As Integer = 0 To mae.Count - 1
                                str = Replace(str, mae(k), "")
                            Next
                            For k As Integer = 0 To ato.Count - 1
                                str = Replace(str, ato(k), "")
                            Next

                            'ランダム並び替え
                            Dim strArray As String() = Split(str, " ")
                            Dim strArray2 As String() = strArray.OrderBy(Function(s) Guid.NewGuid()).ToArray()
                            For k As Integer = 0 To strArray2.Length - 1
                                If k = 0 Then
                                    str = strArray2(k)
                                Else
                                    str &= " " & strArray2(k)
                                End If
                            Next

                            For k As Integer = mae.Count - 1 To 0 Step -1
                                str = mae(k) & str
                            Next
                            For k As Integer = ato.Count - 1 To 0 Step -1
                                str = str & ato(k)
                            Next
                            Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName(form)(name).SetAttribute("value", str)
                        End If
                    Case "保持"
                        Dim va As String() = Split(value, "=")
                        If form = "radio" Then
                            Dim hCollection As HtmlElementCollection = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")
                            Dim num As Integer = 0
                            For Each element As HtmlElement In hCollection
                                If element.GetAttribute("name") = name Then
                                    If element.GetAttribute("checked") Then
                                        searchArray.Add(va(num))
                                        Exit For
                                    End If
                                    num += 1
                                End If
                            Next
                        ElseIf form = "checkbox" Then
                            If Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input").Item(name).GetAttribute("value") = va(0) Then
                                searchArray.Add(va(1))
                            Else
                                searchArray.Add(va(2))
                            End If
                        ElseIf form = "select" Then
                            'Dim hCollection As HtmlElementCollection = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("select")
                            If Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("select")(name).GetAttribute("selectedIndex") = va(0) Then
                                searchArray.Add(va(1))
                            Else
                                searchArray.Add(va(2))
                            End If
                        ElseIf form = "input" Then
                            Dim str As String = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName(form)(name).GetAttribute("value")
                            If InStr(str, va(0)) > 0 Then
                                searchArray.Add(va(1))
                            Else
                                searchArray.Add(va(2))
                            End If
                        End If
                    Case "押下"   'ex) 押下,a,link,追加表示情報   ex) 押下,a,getid,checkbutton
                        If form = "a" And name = "link" Then
                            Dim links As HtmlElementCollection = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("a")
                            For Each link As HtmlElement In links
                                If InStr(link.InnerText, value) > 0 Then
                                    link.InvokeMember("Click")
                                    Exit For
                                End If
                            Next
                        ElseIf name = "getid" Then
                            Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementById(value).InvokeMember("Click")
                        End If

                End Select
                DataGridView1.Item(1, r).Style.BackColor = Color.Yellow
                'Catch ex As Exception
                '    errorMsg &= r + 1 & "行目：" & ex.Message & vbCrLf
                'End Try
            End If
        Next

        MsgBox("各種修正が終了しました。" & vbCrLf & vbCrLf & errorMsg, MsgBoxStyle.SystemModal)
    End Sub

    Private Function Okikae(ByVal value As String)
        Select Case True
            Case Regex.IsMatch(value, "_sp_")
                value = Replace(value, "_sp_", " ")
            Case Regex.IsMatch(value, "_SP_")
                value = Replace(value, "_SP_", "　")
            Case Regex.IsMatch(value, "_n_")
                value = Replace(value, "_n_", vbCrLf)
            Case Regex.IsMatch(value, "_t_")
                value = Replace(value, "_t_", vbTab)
        End Select

        value = Replace(value, "_year_", Now.ToString("yyyy"))
        value = Replace(value, "_m+0_", Month(Now))
        value = Replace(value, "_m-2_", Month(DateAdd(DateInterval.Month, -2, Now)))
        value = Replace(value, "_m-1_", Month(DateAdd(DateInterval.Month, -1, Now)))
        value = Replace(value, "_m+1_", Month(DateAdd(DateInterval.Month, 1, Now)))
        value = Replace(value, "_m+2_", Month(DateAdd(DateInterval.Month, 2, Now)))

        Return value
    End Function

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
        Dim tagBR As String() = New String() {"<center", "</center", "<table", "</table>", "<td ", "</td>", "</tr>", "<img ", "<a ", "<b>", "<font", "br>", "<div", "-->"}

        For i As Integer = 0 To htmlRecords.Count - 1
            Dim flag As Boolean = False
            If InStr(htmlRecords(i), "<table") > 0 Then
                htmlRecords(i) = Replace(htmlRecords(i), "<table", vbCrLf & "<table")
            End If
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

    Public Function SeikeiHTML(html As String) As String
        If html = "" Then
            Return ""
        Else
            Dim htmlRecords As ArrayList = New ArrayList

            'ZetaHtmlTidy
            Dim s As String = ""
            Using tidy As HtmlTidy = New HtmlTidy()
                s = tidy.CleanHtml(html, HtmlTidyOptions.ConvertToXhtml)
            End Using

            s = Replace(s, vbCrLf, "")
            s = Replace(s, vbLf, "")
            s = Replace(s, vbCr, "")
            s = Regex.Replace(s, "<!DOCTYPE h.+HTML Tidy.+head><body>", "")
            s = Regex.Replace(s, "<html.+xmlns.+head><body>", "")
            s = Replace(s, "</body></html>", "")
            s = Replace(s, " />", ">")
            s = LTrim(s)
            s = Replace(s, "  <", "<")
            s = Replace(s, " <", "<")
            s = Replace(s, vbTab, "")
            s = Replace(s, "</font></p><br>", "</font></p =''><br>")  '特殊対応（楽天 smart_caption）

            htmlRecords = Array_Create(s)
            htmlRecords = BR_ADD(htmlRecords)

            Dim str As String = ""
            For k As Integer = 0 To htmlRecords.Count - 1
                str &= htmlRecords(k)
            Next

            Return str
        End If
    End Function

    '1行目からコード生成
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ToolStripComboBox1.SelectedIndex = 0 Then
            Dim bUrl As String = Form1.TabBrowser1.SelectedTab.WebBrowser.Url.PathAndQuery
            Select Case True
                Case Regex.IsMatch(bUrl, "351386")
                    ToolStripComboBox1.SelectedItem = "楽天-あかね用"
                Case Regex.IsMatch(bUrl, "321434")
                    ToolStripComboBox1.SelectedItem = "楽天-通販の暁用"
                Case Regex.IsMatch(bUrl, "313483")
                    ToolStripComboBox1.SelectedItem = "楽天-アリス用"
                Case Regex.IsMatch(bUrl, "lucky9")
                    ToolStripComboBox1.SelectedItem = "yahoo-ラキナイ用"
                Case Regex.IsMatch(bUrl, "kt-zkshop")
                    ToolStripComboBox1.SelectedItem = "yahoo-海東用"
                Case Regex.IsMatch(bUrl, "akaneashop")
                    ToolStripComboBox1.SelectedItem = "yahoo-あかね用"
                Case Regex.IsMatch(bUrl, "fkstyle")
                    ToolStripComboBox1.SelectedItem = "yahoo-FK用"
                Case Else
                    MsgBox("テンプレート店舗を先に選択してください")
                    Exit Sub
            End Select
        End If

        If DataGridView2.RowCount <= 1 Then
            Exit Sub
        End If

        Dim gazo As String = DataGridView2.Item(0, 0).Value
        DataGridView2.AllowUserToAddRows = False

        If DataGridView2.RowCount > 2 Then
            DataGridView2.Rows.Clear()
            DataGridView2.Rows.Add(gazo)
        End If

        If gazo = "" And TextBox3.Text <> "" Then
            Dim tenpo As String = ToolStripComboBox1.SelectedItem
            Dim code As String = TextBox3.Text
            Select Case True
                Case Regex.IsMatch(tenpo, "楽天-あかね|通販の暁")
                    Dim moji2 As String = code.Substring(0, 2)
                    TextBox9.Text = "-"
                    Select Case moji2
                        Case "ad", "ap"
                            TextBox22.Text = moji2
                        Case "od"
                            If Regex.IsMatch(tenpo, "楽天-あかね") Then
                                TextBox22.Text = moji2
                            ElseIf Regex.IsMatch(tenpo, "通販の暁") Then
                                TextBox22.Text = code.Substring(0, 1)
                            End If
                        Case Else
                            TextBox22.Text = code.Substring(0, 1)
                    End Select
                    If Regex.IsMatch(tenpo, "楽天-あかね") Then
                        gazo = "https://image.rakuten.co.jp/ashop/cabinet/" & TextBox22.Text & "/" & code & ".jpg"
                    ElseIf Regex.IsMatch(tenpo, "通販の暁") Then
                        gazo = "https://image.rakuten.co.jp/patri/cabinet/" & TextBox22.Text & "/" & code & ".jpg"
                    End If
                Case Regex.IsMatch(tenpo, "アリス")
                    TextBox22.Text = "0000"
                    TextBox9.Text = "-"
                    gazo = "https://image.rakuten.co.jp/alice-zk/cabinet/" & TextBox22.Text & "/" & code & ".jpg"
                Case Regex.IsMatch(tenpo, "海東|ラキナイ")
                    TextBox22.Text = ""
                    TextBox9.Text = "_"
                    If Regex.IsMatch(tenpo, "海東") Then
                        gazo = "https://item-shopping.c.yimg.jp/i/f/kt-zkshop_" & code
                    ElseIf Regex.IsMatch(tenpo, "ラキナイ") Then
                        gazo = "https://item-shopping.c.yimg.jp/i/f/lucky9_" & code
                    End If
                Case Regex.IsMatch(tenpo, "FK|yahoo-あかね")
                    TextBox22.Text = ""
                    TextBox9.Text = "-"
                    If Regex.IsMatch(tenpo, "FK") Then
                        gazo = "https://shopping.c.yimg.jp/lib/fkstyle/" & code & ".jpg"
                    ElseIf Regex.IsMatch(tenpo, "yahoo-あかね") Then
                        gazo = "https://shopping.c.yimg.jp/lib/akaneashop/" & code & ".jpg"
                    End If
            End Select

            DataGridView2.Rows.Add(gazo)
        End If

        gazo = DataGridView2.Item(0, 0).Value

        Dim gazoDir As String = Path.GetDirectoryName(gazo)
        gazoDir = Replace(gazoDir, "https:\", "https://")
        gazoDir = Replace(gazoDir, "http:\", "http://")
        gazoDir = Replace(gazoDir, "\", "/")
        Dim gazoName As String = Path.GetFileNameWithoutExtension(gazo)
        'gazoName = Regex.Replace(gazoName, "[fk|kt|lu|ak].*?_", "")
        TextBox3.Text = gazoName
        Dim gazoExt As String = Path.GetExtension(gazo)
        For i As Integer = NumericUpDown2.Value To NumericUpDown1.Value
            DataGridView2.Rows.Add(gazoDir & "/" & gazoName & TextBox9.Text & i & gazoExt)
        Next
        DataGridView2.AllowUserToAddRows = True
    End Sub

    '正規表現チェック
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If TextBox6.Text <> "" And TextBox7.Text <> "" Then
            Try
                Dim pattern As New Regex(TextBox7.Text)
                Dim m As System.Text.RegularExpressions.Match = pattern.Match(TextBox6.Text)
                TextBox8.Text = m.Value
                MsgBox("チェック完了", MsgBoxStyle.SystemModal)
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        Else
            MsgBox("文字列がありません", MsgBoxStyle.SystemModal)
        End If
    End Sub


    '項目選択肢を取得
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        DataGridView3.Rows.Clear()
        DataGridView3.Columns.Clear()
        DataGridView3.Columns.Add(0, "行/列")
        DataGridView3.Columns(DataGridView3.ColumnCount - 1).SortMode = DataGridViewColumnSortMode.NotSortable

        'webbrowserのdocumentText文字化け対策
        Dim buf As Byte() = New Byte(Form1.TabBrowser1.SelectedTab.WebBrowser.DocumentStream.Length) {}
        Form1.TabBrowser1.SelectedTab.WebBrowser.DocumentStream.Read(buf, 0, CInt(Form1.TabBrowser1.SelectedTab.WebBrowser.DocumentStream.Length))
        Dim ec As System.Text.Encoding = System.Text.Encoding.GetEncoding(Form1.TabBrowser1.SelectedTab.WebBrowser.Document.Encoding)
        Dim htmls As String = ec.GetString(buf)

        Dim doc As HtmlAgilityPack.HtmlDocument = New HtmlAgilityPack.HtmlDocument
        doc.LoadHtml(htmls)
        Dim doc2 As HtmlAgilityPack.HtmlDocument = New HtmlAgilityPack.HtmlDocument

        Dim s As String = ""
        Dim selecter As String = "//font"
        Dim selecter2 As String = "//option"
        Dim nodes As HtmlAgilityPack.HtmlNodeCollection = doc.DocumentNode.SelectNodes(selecter)
        DataGridView3.Rows.Add(1)
        Dim zaikoNum As Integer = 0

        '旧版本
        'For Each node As HtmlAgilityPack.HtmlNode In nodes
        '    Dim str As String = node.InnerText
        '    str = Regex.Replace(str, "\r|\n", " ")
        '    str = Replace(str, " ", "")
        '    If InStr(str, "同行一括") > 0 Or InStr(str, "/-") > 0 Or InStr(str, "在庫数:") > 0 Then
        '        If InStr(str, "同行一括") > 0 Then
        '            DataGridView3.Rows.Add(1)
        '            str = Replace(str, "同行一括", "")
        '            DataGridView3.Item(0, DataGridView3.RowCount - 2).Value = str
        '            zaikoNum = 0
        '        ElseIf InStr(str, "同列一括") > 0 Or InStr(str, "/-") > 0 Then
        '            DataGridView3.Columns.Add(1, "")
        '            DataGridView3.Columns(DataGridView3.ColumnCount - 1).SortMode = DataGridViewColumnSortMode.NotSortable
        '            str = Replace(str, "同列一括", "")
        '            DataGridView3.Item(DataGridView3.ColumnCount - 1, 0).Value = str
        '        ElseIf InStr(str, "在庫数:") > 0 Then
        '            str = Replace(str, "在庫数:", "")
        '            DataGridView3.Item(zaikoNum + 1, DataGridView3.RowCount - 2).Value = str
        '            '予約---------------------------------------------
        '            doc2.LoadHtml(node.ParentNode.InnerHtml)
        '            Dim nodes2 As HtmlAgilityPack.HtmlNodeCollection = doc2.DocumentNode.SelectNodes(selecter2)
        '            If nodes2 IsNot Nothing Then
        '                For Each node2 As HtmlAgilityPack.HtmlNode In nodes2
        '                    If InStr(node2.OuterHtml, "selected") > 0 Then
        '                        DataGridView3.Item(zaikoNum + 1, DataGridView3.RowCount - 2).Value &= "," & node2.GetAttributeValue("value", "")
        '                    End If
        '                Next
        '            End If
        '            '予約---------------------------------------------
        '            zaikoNum += 1
        '        End If
        '        s &= str & vbCrLf
        '    End If
        'Next

        '新增 加了タグID和画像
        For Each node As HtmlAgilityPack.HtmlNode In nodes
            Dim str As String = node.InnerText
            str = Regex.Replace(str, "\r|\n", " ")
            str = Replace(str, " ", "")
            If InStr(str, "同行一括") > 0 Or InStr(str, "/-") > 0 Or InStr(str, "在庫数:") > 0 Then
                If InStr(str, "同行一括") > 0 Then
                    DataGridView3.Rows.Add(1)
                    str = Replace(str, "同行一括", "")
                    DataGridView3.Item(0, DataGridView3.RowCount - 2).Value = str
                    zaikoNum = 0
                ElseIf InStr(str, "同列一括") > 0 Or InStr(str, "/-") > 0 Then
                    DataGridView3.Columns.Add(1, "")
                    DataGridView3.Columns(DataGridView3.ColumnCount - 1).SortMode = DataGridViewColumnSortMode.NotSortable
                    str = Replace(str, "同列一括", "")
                    DataGridView3.Item(DataGridView3.ColumnCount - 1, 0).Value = str
                ElseIf InStr(str, "在庫数:") > 0 Then
                    str = Replace(str, "在庫数:", "")
                    DataGridView3.Item(zaikoNum + 1, DataGridView3.RowCount - 2).Value = str
                    '予約---------------------------------------------
                    doc2.LoadHtml(node.ParentNode.InnerHtml)
                    Dim nodes2 As HtmlAgilityPack.HtmlNodeCollection = doc2.DocumentNode.SelectNodes(selecter2)
                    If nodes2 IsNot Nothing Then
                        For Each node2 As HtmlAgilityPack.HtmlNode In nodes2
                            If InStr(node2.OuterHtml, "selected") > 0 Then
                                DataGridView3.Item(zaikoNum + 1, DataGridView3.RowCount - 2).Value &= "," & node2.GetAttributeValue("value", "")
                            End If
                        Next
                    End If
                    '予約---------------------------------------------
                    zaikoNum += 1
                End If
                s &= str & vbCrLf
            End If
        Next

        Dim ontaguid As Boolean = False 'ontaguid是true的时候，后一次循环取得id
        Dim ongazou As Boolean = False
        Dim index As Integer = 0

        If DataGridView3.ColumnCount = 2 Then '如果是两列
            DataGridView3.Columns.Add(2, "タグID")
            DataGridView3.Columns.Add(3, "画像")
            index = 1 '从第二行开始加
            For Each node As HtmlAgilityPack.HtmlNode In nodes
                Dim str As String = node.InnerText
                str = Regex.Replace(str, "\r|\n", " ")
                str = Replace(str, " ", "")
                If InStr(str, "タグID") > 0 Then
                    ontaguid = True
                    Continue For
                ElseIf InStr(str, "画像:") > 0 Then
                    ongazou = True
                    Continue For
                ElseIf ontaguid = True Then
                    DataGridView3.Item(2, index).Value = str
                    ontaguid = False
                ElseIf ongazou = True Then
                    DataGridView3.Item(3, index).Value = str
                    index += 1
                    ongazou = False
                End If
            Next
        ElseIf DataGridView3.ColumnCount > 2 And DataGridView3.RowCount > 2 Then '如果是多列多行
            Dim column_count As Integer = DataGridView3.ColumnCount - 1

            Dim addTaguid_arr As New ArrayList
            Dim addTaguid As String '要添加的タグID
            Dim addTaguid_count As Integer = 0 '加了多少次

            Dim addGazou_arr As New ArrayList
            Dim addGazou As String '要添加的タグID
            Dim addGazou_count As Integer = 0 '加了多少次

            DataGridView3.Columns.Add(column_count, "タグID")
            DataGridView3.Columns.Add(column_count + 1, "画像")

            For Each node As HtmlAgilityPack.HtmlNode In nodes
                Dim str As String = node.InnerText
                str = Regex.Replace(str, "\r|\n", " ")
                str = Replace(str, " ", "")
                If InStr(str, "タグID") > 0 Then
                    ontaguid = True
                    Continue For
                ElseIf InStr(str, "画像:") > 0 Then
                    ongazou = True
                    Continue For
                ElseIf ontaguid = True Then
                    If addTaguid_count < column_count And addTaguid_count <> (column_count - 1) Then
                        If InStr(addTaguid, "|") > 0 Then
                            addTaguid = addTaguid & str & "|"
                        Else
                            addTaguid = str & "|"
                        End If
                        addTaguid_count += 1
                    ElseIf addTaguid_count = (column_count - 1) Then
                        addTaguid = addTaguid & str
                        addTaguid_arr.Add(addTaguid)
                        addTaguid = "" 'clear
                        addTaguid_count = 0
                    End If
                    ontaguid = False
                ElseIf ongazou = True Then
                    If addGazou_count < column_count And addGazou_count <> (column_count - 1) Then
                        If InStr(addGazou, "|") > 0 Then
                            addGazou = addGazou & str & "|"
                        Else
                            addGazou = str & "|"
                        End If
                        addGazou_count += 1
                    ElseIf addGazou_count = (column_count - 1) Then
                        addGazou = addGazou & str
                        addGazou_arr.Add(addGazou)
                        addGazou = "" 'clear
                        addGazou_count = 0
                    End If
                    ongazou = False
                End If
            Next

            If addTaguid_arr.Count > 0 Then
                If addTaguid_arr.Count = DataGridView3.RowCount - 2 Then
                    For r As Integer = 1 To DataGridView3.RowCount - 2
                        DataGridView3.Item(DataGridView3.ColumnCount - 2, r).Value = addTaguid_arr(r - 1)
                        DataGridView3.Item(DataGridView3.ColumnCount - 1, r).Value = addGazou_arr(r - 1)
                    Next
                End If
            End If
        End If

        TextBox10.Text = s
    End Sub

    Private Function HTMLGet(ByVal doc As HtmlAgilityPack.HtmlDocument, ByVal selecter As String, ByVal inout As Integer, ByVal num As Integer)
        Dim nodes As HtmlAgilityPack.HtmlNodeCollection = doc.DocumentNode.SelectNodes(selecter)
        'For Each node As HtmlAgilityPack.HtmlNode In nodes
        '    MsgBox(node.InnerHtml, MsgBoxStyle.SystemModal)
        'Next

        If Not nodes Is Nothing Then
            If inout = 0 Then
                Return nodes(num).InnerText
            ElseIf inout = 1 Then
                Return nodes(num).InnerHtml
            ElseIf inout = 2 Then
                Return nodes(num).GetAttributeValue("src", "")
            ElseIf inout = 3 Then
                Return nodes
            Else
                Return nodes(num).OuterHtml
            End If
        Else
            Return "error"
        End If
    End Function

    '行列入替
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim dArray As ArrayList = New ArrayList
        For r As Integer = 0 To DataGridView3.RowCount - 2
            Dim lineStr As String = ""
            For c As Integer = 0 To DataGridView3.ColumnCount - 1
                If c = 0 Then
                    lineStr = DataGridView3.Item(c, r).Value
                Else
                    lineStr &= "|=|" & DataGridView3.Item(c, r).Value
                End If
            Next
            dArray.Add(lineStr)
        Next

        DataGridView3.Rows.Clear()
        DataGridView3.Columns.Clear()

        For i As Integer = 0 To dArray.Count - 1
            If i = 0 Then
                DataGridView3.Columns.Add(0, "行/列")
            Else
                DataGridView3.Columns.Add(1, "")
            End If
        Next

        For i As Integer = 0 To dArray.Count - 1
            Dim rArray As String() = Split(dArray(i), "|=|")
            If i = 0 Then
                DataGridView3.Rows.Add(rArray.Length)
            End If
            For r As Integer = 0 To rArray.Length - 1
                DataGridView3.Item(i, r).Value = rArray(r)
            Next
        Next
    End Sub

    '項目選択肢を書き込み
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim dH3 As ArrayList = TM_HEADER_GET(DataGridView3)
        Dim tagid_index As Integer = 0
        If dH3.Contains("タグID") Then
            tagid_index = dH3.IndexOf("タグID")
        End If
        For r As Integer = 0 To 20
            If r = 0 Then
                For c As Integer = 0 To 20
                    If c <> tagid_index Then
                        If c <= DataGridView3.ColumnCount - 1 Then
                            If DataGridView3.Item(c, r).Value <> "" Then
                                If InStr(DataGridView3.Item(c, r).Value, "/-") = 0 Then
                                    MsgBox("「" & DataGridView3.Item(c, r).Value & "」書式が異なります", MsgBoxStyle.SystemModal)
                                Else
                                    Try
                                        Dim strArray As String() = Split(DataGridView3.Item(c, r).Value, "/-")
                                        Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("choice_name_hor_" & c).SetAttribute("value", strArray(0))
                                        Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("child_no_hor_" & c).SetAttribute("value", strArray(1))
                                    Catch ex As Exception
                                        MsgBox("書き込みできません。" & vbCrLf & "ページが正しいか確認してください。", MsgBoxStyle.SystemModal)
                                        Exit Sub
                                    End Try
                                End If
                            End If
                        Else
                            Try
                                Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("choice_name_hor_" & c).SetAttribute("value", "")
                                Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("child_no_hor_" & c).SetAttribute("value", "")
                            Catch ex As Exception

                            End Try
                        End If
                    End If
                Next
            Else
                If r <= DataGridView3.RowCount - 1 Then
                    If DataGridView3.Item(0, r).Value <> "" Then
                        Try
                            Dim strArray As String() = Split(DataGridView3.Item(0, r).Value, "/-")
                            Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("choice_name_ver_" & r).SetAttribute("value", strArray(0))
                            Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("child_no_ver_" & r).SetAttribute("value", strArray(1))
                        Catch ex As Exception
                            MsgBox("書き込みできません。" & vbCrLf & "ページが正しいか確認してください。", MsgBoxStyle.SystemModal)
                            Exit Sub
                        End Try
                    End If
                Else
                    Try
                        Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("choice_name_ver_" & r).SetAttribute("value", "")
                        Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("child_no_ver_" & r).SetAttribute("value", "")
                    Catch ex As Exception

                    End Try
                End If
            End If
        Next
    End Sub

    '在庫書き込み
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        ZaikoKakikomi(0)
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        ZaikoKakikomi(1)
    End Sub

    Private Sub ZaikoKakikomi(mode As Integer)

        'old  SfB
        'For r As Integer = 1 To DataGridView3.RowCount - 1
        '    For c As Integer = 1 To DataGridView3.ColumnCount - 1
        '        If DataGridView3.Item(c, r).Value <> "" Then
        '            Try
        '                Dim name1 As String = "inventory_" & c & "_" & r
        '                Dim name2_1 As String = "inventory_" & c & "_" & r - mode
        '                Dim name2_2 As String = "normal_delvdate_id_" & c & "_" & r - mode
        '                Dim name2_3 As String = "backorder_delvdate_id_" & c & "_" & r - mode
        '                Dim value As String = DataGridView3.Item(c, r).Value

        '                If InStr(value, ",") > 0 Then
        '                    Dim valueArray As String() = Split(value, ",")
        '                    Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")(name2_1).SetAttribute("value", valueArray(0))
        '                    If valueArray.Length > 1 Then
        '                        If valueArray(1) <> "" Then
        '                            Try
        '                                Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("select")(name2_2).SetAttribute("value", valueArray(1))
        '                            Catch ex As Exception

        '                            End Try
        '                        End If
        '                    End If
        '                    If valueArray.Length > 2 Then
        '                        If valueArray(2) <> "" Then
        '                            Try
        '                                Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("select")(name2_3).SetAttribute("value", valueArray(2))
        '                            Catch ex As Exception

        '                            End Try
        '                        End If
        '                    End If
        '                    'Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementById("normal_delvdate_id_1_0").RaiseEvent("onchange")
        '                Else
        '                    Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")(name1).SetAttribute("value", value)
        '                End If
        '            Catch ex As Exception
        '                MsgBox("書き込みできません。" & vbCrLf & "ページが正しいか確認してください。", MsgBoxStyle.SystemModal)
        '                Exit Sub
        '            End Try
        '        End If
        '    Next
        'Next

        Dim dH3 As ArrayList = TM_HEADER_GET(DataGridView3)
        Dim tagid_index As Integer = 0
        If dH3.Contains("タグID") Then
            tagid_index = dH3.IndexOf("タグID")
        End If

        Dim gazou_index As Integer = 0
        If dH3.Contains("画像") Then
            gazou_index = dH3.IndexOf("画像")
        End If

        Dim tagid_index_left As Integer = 0
        Dim tagid_index_right As Integer = 0
        For r As Integer = 1 To DataGridView3.RowCount - 1
            For c As Integer = 1 To DataGridView3.ColumnCount - 1
                If c = tagid_index Or c = gazou_index Then
                    Continue For
                End If
                If DataGridView3.Item(c, r).Value <> "" Then
                    Try
                        Dim name1 As String = "inventory_" & c & "_" & r
                        Dim name2_1 As String = "inventory_" & c & "_" & r - mode
                        Dim name2_2 As String = "normal_delvdate_id_" & c & "_" & r - mode
                        Dim name2_3 As String = "backorder_delvdate_id_" & c & "_" & r - mode
                        Dim value As String = DataGridView3.Item(c, r).Value

                        If InStr(value, ",") > 0 Then
                            Dim valueArray As String() = Split(value, ",")
                            Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")(name2_1).SetAttribute("value", valueArray(0))
                            If valueArray.Length > 1 Then
                                If valueArray(1) <> "" Then
                                    Try
                                        Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("select")(name2_2).SetAttribute("value", valueArray(1))
                                    Catch ex As Exception

                                    End Try
                                End If
                            End If
                            If valueArray.Length > 2 Then
                                If valueArray(2) <> "" Then
                                    Try
                                        Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("select")(name2_3).SetAttribute("value", valueArray(2))
                                    Catch ex As Exception

                                    End Try
                                End If
                            End If
                            'Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementById("normal_delvdate_id_1_0").RaiseEvent("onchange")
                        Else
                            Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")(name1).SetAttribute("value", value)
                            tagid_index_left = c
                            tagid_index_right = r

                            If tagid_index <> 0 Then
                                Dim tagid As String = "inventory_tag_id_" & tagid_index_left & "_" & tagid_index_right
                                Dim tagid_val As String() = DataGridView3.Item(tagid_index, r).Value.ToString.Split("|")
                                Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementById(tagid).InnerText = tagid_val(c - 1)
                            End If
                            If gazou_index <> 0 Then
                                Dim gazou As String = "inventory_image_url_" & tagid_index_left & "_" & tagid_index_right
                                Dim gazou_val As String() = DataGridView3.Item(gazou_index, r).Value.ToString.Split("|")
                                Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementById(gazou).InnerText = gazou_val(c - 1)
                            End If
                        End If
                    Catch ex As Exception
                        MsgBox("書き込みできません。" & vbCrLf & "ページが正しいか確認してください。", MsgBoxStyle.SystemModal)
                        Exit Sub
                    End Try
                End If
            Next
        Next
    End Sub

    Private Sub DataGridView1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.GotFocus
        dgv = DataGridView1
    End Sub

    Private Sub DataGridView2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView2.GotFocus
        dgv = DataGridView2
    End Sub

    Private Sub DataGridView3_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView3.GotFocus
        dgv = DataGridView3
    End Sub

    Private Sub DataGridView4_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView4.GotFocus
        dgv = DataGridView4
    End Sub

    Private Sub DataGridView5_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView5.GotFocus
        dgv = DataGridView5
    End Sub

    Private Sub DataGridView6_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles DGV6.GotFocus
        dgv = DGV6
    End Sub

    '*******************************************************************************
    'ContextMenuStrip
    '*******************************************************************************
    Dim dgv As DataGridView = DataGridView1
    Private Sub 切り取りToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 切り取りToolStripMenuItem1.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim selCell = dgv.SelectedCells
        CUTS(dgv)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub コピーToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles コピーToolStripMenuItem1.Click
        Dim selCell = dgv.SelectedCells
        COPYS(dgv)
    End Sub

    Private Sub 貼り付けToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 貼り付けToolStripMenuItem1.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim selCell = dgv.SelectedCells
        PASTES(dgv, selCell)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 上に挿入ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 上に挿入ToolStripMenuItem.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim selCell = dgv.SelectedCells
        ROWSADD(dgv, selCell, 1)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 下に挿入ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 下に挿入ToolStripMenuItem.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim selCell = dgv.SelectedCells
        ROWSADD(dgv, selCell, 0)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 右に挿入ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 右に挿入ToolStripMenuItem.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim selCell = dgv.SelectedCells
        COLSADD(dgv, selCell, 0)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 左に挿入ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 左に挿入ToolStripMenuItem.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim selCell = dgv.SelectedCells
        COLSADD(dgv, selCell, 1)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 行を選択直下に複製ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 行を選択直下に複製ToolStripMenuItem.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim selR As Integer = dgv.SelectedCells(0).RowIndex
        dgv.Rows.InsertCopy(selR, selR + 1)
        For c As Integer = 0 To dgv.ColumnCount - 1
            If dgv.Item(c, selR).Value <> "" Then
                dgv.Item(c, selR + 1).Value = dgv.Item(c, selR).Value
            End If
        Next
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 画像番号を1増やすToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 画像番号を1増やすToolStripMenuItem.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim selR As Integer = dgv.SelectedCells(0).RowIndex
        dgv.Rows.InsertCopy(selR, selR + 1)
        For c As Integer = 0 To dgv.ColumnCount - 1
            If dgv.Item(c, selR).Value <> "" Then   'alt欄の修正（番号付きなら1増やす）
                If c = 0 Then
                    If InStr(dgv.Item(c, selR).Value, "-") > 0 Then
                        Dim strArray As String() = Split(dgv.Item(c, selR).Value, "-")
                        If IsNumeric(strArray(1)) Then
                            dgv.Item(c, selR + 1).Value = strArray(0) & "-" & CInt(strArray(1)) + 1
                        Else
                            dgv.Item(c, selR + 1).Value = dgv.Item(c, selR).Value
                        End If
                    Else
                        dgv.Item(c, selR + 1).Value = dgv.Item(c, selR).Value
                    End If
                Else
                    Dim imgStr As String = dgv.Item(c, selR).Value
                    Dim filePath As String = Path.GetDirectoryName(imgStr)
                    filePath = Replace(filePath, "https:\", "https://")
                    filePath = Replace(filePath, "\", "/")
                    If InStr(imgStr, "_") = 0 Then
                        Dim fileName As String() = Split(Path.GetFileNameWithoutExtension(imgStr), "-")
                        Dim fileExt As String = Path.GetExtension(imgStr)
                        If fileName.Length >= 2 Then
                            If IsNumeric(fileName(1)) Then
                                imgStr = filePath & "/" & fileName(0) & "-" & CInt(fileName(1)) + 1 & "" & fileExt
                                dgv.Item(c, selR + 1).Value = imgStr
                            Else
                                dgv.Item(c, selR + 1).Value = dgv.Item(c, selR).Value
                            End If
                        End If
                    Else
                        Dim fileName As String() = Regex.Split(Path.GetFileNameWithoutExtension(imgStr), "\-|_")
                        Dim fileExt As String = Path.GetExtension(imgStr)
                        Dim fnNum As Integer = fileName.Length - 1
                        If fileName.Length >= 2 Then
                            If IsNumeric(fileName(fnNum)) Then
                                Dim fn As String = ""
                                For i As Integer = 0 To fileName.Length - 2
                                    If i = 0 Then
                                        fn &= fileName(i)
                                    Else
                                        fn &= "_" & fileName(i)
                                    End If
                                Next
                                imgStr = filePath & "/" & fn & "_" & CInt(fileName(fnNum)) + 1 & "" & fileExt
                                dgv.Item(c, selR + 1).Value = imgStr
                            Else
                                dgv.Item(c, selR + 1).Value = dgv.Item(c, selR).Value
                            End If
                        End If
                    End If
                End If
            End If
        Next
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 削除ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 削除ToolStripMenuItem1.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim selCell = dgv.SelectedCells
        DELS(dgv, selCell)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 行を削除ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 行を削除ToolStripMenuItem1.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim selCell = dgv.SelectedCells
        ROWSCUT(dgv, selCell)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    Private Sub 列を削除ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 列を削除ToolStripMenuItem1.Click
        '-----------------------------------------------------------------------
        PreDIV()
        '-----------------------------------------------------------------------
        Dim selCell = dgv.SelectedCells
        COLSCUT(dgv, selCell)
        '-----------------------------------------------------------------------
        RetDIV()
        '-----------------------------------------------------------------------
    End Sub

    '↑↑=======================================================================↑↑
    'ContextMenuStripここまで
    '===============================================================================

    ' ----------------------------------------------------------------------------------
    ' DataGridViewでCtrl-Vキー押下時にクリップボードから貼り付けるための実装↓↓↓↓↓↓
    ' DataGridViewでDelやBackspaceキー押下時にセルの内容を消去するための実装↓↓↓↓↓↓
    ' ----------------------------------------------------------------------------------
    Public ColumnChars() As String
    Private _editingColumn As Integer   '編集中のカラム番号
    Private _editingCtrl As DataGridViewTextBoxEditingControl

    Private Sub DataGridView_EditingControlShowing(
        ByVal sender As Object, ByVal e As DataGridViewEditingControlShowingEventArgs) _
        Handles DataGridView1.EditingControlShowing
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

    Private Sub DataGridView_KeyUp(ByVal sender As Object, ByVal e As Windows.Forms.KeyEventArgs) Handles DataGridView1.KeyUp
        Dim dgv As DataGridView = CType(sender, DataGridView)
        Dim selCell = dgv.SelectedCells

        If e.KeyCode = Keys.Back Then    ' セルの内容を消去
            If DataGridView1.IsCurrentCellInEditMode Then

            Else
                DELS(dgv, selCell)
            End If
        ElseIf e.Modifiers = Keys.Control And e.KeyCode = Keys.Enter Then '複数セル一括入力
            '-----------------------------------------------------------------------
            PreDIV()
            '-----------------------------------------------------------------------
            Dim str As String = selCell(0).Value
            For i As Integer = 0 To selCell.Count - 1
                selCell(i).Value = str
            Next
            '-----------------------------------------------------------------------
            RetDIV()
            '-----------------------------------------------------------------------
        End If
    End Sub
    ' ----------------------------------------------------------------------------------
    ' DataGridViewでCtrl-Vキー押下時にクリップボードから貼り付けるための実装↑↑↑↑↑↑
    ' DataGridViewでDelやBackspaceキー押下時にセルの内容を消去するための実装↑↑↑↑↑↑
    ' ----------------------------------------------------------------------------------

    '*******************************************************************************
    '切り取り、コピー、貼り付けなど
    '*******************************************************************************
    Private Sub DELS(ByVal dgv As DataGridView, ByVal selCell As Object)
        'Dim dgv As DataGridView = DataGridView1
        'Dim selCell = dgv.SelectedCells
        For i As Integer = 0 To selCell.Count - 1
            'dgv(x, y).Value = ""
            dgv(selCell(i).ColumnIndex, selCell(i).RowIndex).Value = ""
        Next
    End Sub

    Public Shared DGV1tb As TextBox
    Private Sub DataGridView1_EditingControlShowing(ByVal sender As Object, ByVal e As Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles _
            DataGridView1.EditingControlShowing, DataGridView3.EditingControlShowing
        If TypeOf e.Control Is DataGridViewTextBoxEditingControl Then
            '編集のために表示されているコントロールを取得
            DGV1tb = CType(e.Control, TextBox)
        End If
    End Sub

    Private Sub COPYS(ByVal dgv As DataGridView)
        If dgv.GetCellCount(DataGridViewElementStates.Selected) > 0 Then
            If dgv.CurrentCell.IsInEditMode = True Then
                DGV1tb.Copy()
            Else
                Try
                    'Clipboard.SetText(dgv.GetClipboardContent.GetText)
                    Dim returnValue As DataObject
                    returnValue = dgv.GetClipboardContent()
                    Dim strTab = returnValue.GetText()
                    Clipboard.SetDataObject(strTab, True)
                Catch ex As System.Runtime.InteropServices.ExternalException
                    MessageBox.Show("コピーできませんでした。")
                End Try
            End If
        End If
    End Sub

    Private Sub CUTS(ByVal dgv As DataGridView)
        If dgv.GetCellCount(DataGridViewElementStates.Selected) > 0 Then
            If dgv.CurrentCell.IsInEditMode = True Then
                DGV1tb.Cut()
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
                    MessageBox.Show("コピーできませんでした。")
                End Try
            End If
        End If
    End Sub

    Private Sub PASTES(ByVal dgv As DataGridView, ByVal selCell As Object)
        Dim x As Integer = dgv.CurrentCellAddress.X
        Dim y As Integer = dgv.CurrentCellAddress.Y
        ' クリップボードの内容を取得
        Dim clipText As String = Clipboard.GetText()
        ' 改行を変換
        clipText = clipText.Replace(vbCrLf, vbLf)
        clipText = clipText.Replace(vbCr, vbLf)
        ' 改行で分割
        Dim lines() As String = clipText.Split(vbLf)

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
                    ' ------------------------------------------------------
                    ' 次のセルへ
                    c2 = c2 + 1
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
        delR.Reverse()
        For r As Integer = 0 To delR.Count - 1
            If delR(r) < dgv.RowCount - 1 Then
                dgv.Rows.RemoveAt(delR(r))
            End If
        Next
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
    'Private Sub ROWSADD(ByVal dgv As DataGridView, ByVal selCell As Object)
    '    Dim delR As New ArrayList
    '    For i As Integer = 0 To selCell.Count - 1
    '        If Not delR.Contains(selCell(i).RowIndex) Then
    '            delR.Add(selCell(i).RowIndex)
    '        End If
    '    Next
    '    delR.Sort()
    '    If delR(0) <> 0 Then
    '        dgv.Rows.Insert(delR(0), delR.Count)
    '    End If
    'End Sub

    'Private Sub COLSADD(ByVal dgv As DataGridView, ByVal selCell As Object)
    '    Dim delR As New ArrayList
    '    For i As Integer = 0 To selCell.Count - 1
    '        If Not delR.Contains(selCell(i).ColumnIndex) Then
    '            delR.Add(selCell(i).ColumnIndex)
    '        End If
    '    Next
    '    delR.Sort()
    '    If delR(0) <> 0 Then
    '        Dim num As String = dgv.ColumnCount
    '        Dim textColumn As New DataGridViewTextBoxColumn()
    '        textColumn.Name = "Column" & num
    '        textColumn.HeaderText = num
    '        dgv.Columns.Insert(delR(0), textColumn)
    '    End If
    'End Sub

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

    '↑↑=======================================================================↑↑
    'ContextMenuStripここまで
    '===============================================================================

    Dim cellChange As Boolean = True
    Dim sc As Integer = 0
    Dim sr As Integer = 0
    Public Sub PreDIV()
        '-----------------------------------------------------------------------
        sc = DataGridView1.HorizontalScrollingOffset
        sr = DataGridView1.FirstDisplayedScrollingRowIndex
        cellChange = False
        DataGridView1.ScrollBars = ScrollBars.None
        'DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
        DataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        DataGridView1.ScrollBars = ScrollBars.Both
        '-----------------------------------------------------------------------
    End Sub

    Public Sub RetDIV()
        '-----------------------------------------------------------------------
        cellChange = True
        '***************************************************
        DataGridView1.HorizontalScrollingOffset = sc
        '***************************************************
        If sr > DataGridView1.RowCount - 2 Then
            sr = DataGridView1.RowCount - 2
        End If
        If sr < 0 Then
            sr = 0
        End If
        DataGridView1.FirstDisplayedScrollingRowIndex = sr
        '-----------------------------------------------------------------------
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        If TabControl1.SelectedIndex = 0 Then
            dgv = DataGridView1
        ElseIf TabControl1.SelectedIndex = 1 Then
            dgv = DataGridView3
        End If

        If TabControl1.SelectedTab.Text = "住所変換" Then
            Timer1.Enabled = True
        Else
            Timer1.Enabled = False
        End If
    End Sub

    'クリア
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        DataGridView2.Rows.Clear()
        TextBox3.Text = ""
        'ComboBox1.SelectedIndex = 0    '配送方法クリアしない
        TextBox2.Text = ""
        TextBox13.Text = ""
        'ComboBox3.SelectedIndex = 0    '通常・予約クリアしない
    End Sub

    '楽天カテゴリー情報取得
    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        DataGridView4.Rows.Clear()

        'webbrowserのdocumentText文字化け対策
        Dim buf As Byte() = New Byte(Form1.TabBrowser1.SelectedTab.WebBrowser.DocumentStream.Length) {}
        Form1.TabBrowser1.SelectedTab.WebBrowser.DocumentStream.Read(buf, 0, CInt(Form1.TabBrowser1.SelectedTab.WebBrowser.DocumentStream.Length))
        Dim ec As System.Text.Encoding = System.Text.Encoding.GetEncoding(Form1.TabBrowser1.SelectedTab.WebBrowser.Document.Encoding)
        Dim htmls As String = ec.GetString(buf)

        Dim doc As HtmlAgilityPack.HtmlDocument = New HtmlAgilityPack.HtmlDocument
        doc.LoadHtml(htmls)

        Dim s As String = ""
        Dim selecter As String = "//font"
        Dim nodes As HtmlAgilityPack.HtmlNodeCollection = doc.DocumentNode.SelectNodes(selecter)
        Dim cate As String = ""
        Dim cateParent As Boolean = True
        Dim catePNum As Integer = 0
        For Each node As HtmlAgilityPack.HtmlNode In nodes
            If Regex.IsMatch(node.InnerText, "(\d{1,})") = True Then
                If Regex.IsMatch(node.InnerText, "下位カテゴリ|階層目|複数商品形式|Rakuten") = False Then
                    If cateParent = True Then
                        Dim parent As String = node.InnerText
                        parent = Replace(parent, "amp;", "")
                        DataGridView4.Rows.Add(parent)
                        cate = parent
                        Dim numStr As String = Regex.Replace(parent, ".*\(", "")
                        numStr = Regex.Replace(numStr, "\)", "")
                        catePNum = numStr
                    Else
                        Dim child As String = cate & "\" & node.InnerText
                        child = Replace(child, "amp;", "")
                        DataGridView4.Rows.Add(child)
                        catePNum -= 1
                        If catePNum = 0 Then
                            cateParent = True
                        Else
                            cateParent = False
                        End If
                    End If
                End If
            Else
                If Regex.IsMatch(node.InnerText, "[+]") And cateParent = 0 Then
                    cateParent = True
                Else
                    cateParent = False
                End If
            End If
        Next
    End Sub

    '情報取得
    Dim Dmode As Integer = 0
    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Try
            'webbrowserのdocumentText文字化け対策
            Dim buf As Byte() = New Byte(Form1.TabBrowser1.SelectedTab.WebBrowser.DocumentStream.Length) {}
            Form1.TabBrowser1.SelectedTab.WebBrowser.DocumentStream.Read(buf, 0, CInt(Form1.TabBrowser1.SelectedTab.WebBrowser.DocumentStream.Length))
            Dim ec As System.Text.Encoding = System.Text.Encoding.GetEncoding(Form1.TabBrowser1.SelectedTab.WebBrowser.Document.Encoding)
            Dim htmls As String = ec.GetString(buf)

            Dim doc As HtmlAgilityPack.HtmlDocument = New HtmlAgilityPack.HtmlDocument
            doc.LoadHtml(htmls)

            DataGridView5.Rows.Clear()

            Dim str As String = ""
            Dim sArray As String() = Nothing

            str = HTMLGet(doc, "//span[@class='transact-info-table-cell']", 0, 5)
            If InStr(str, "商品ID") > 0 Then    'メルカリ
                Dmode = 1
                str = HTMLGet(doc, "//ul[@class='transact-info-table-cell']", 0, 5)
                str = StrSeikei(str)
                DataGridView5.Rows.Add("商品ID", str)
                str = HTMLGet(doc, "//ul[@class='transact-info-table-cell']", 0, 2)
                str = StrSeikei(str)
                DataGridView5.Rows.Add("販売手数料", str)
                str = HTMLGet(doc, "//ul[@class='transact-info-table-cell']", 0, 3)
                str = StrSeikei(str)
                DataGridView5.Rows.Add("販売利益", str)
                str = HTMLGet(doc, "//ul[@class='transact-info-table-cell']", 0, 4)
                str = StrSeikei(str)
                Dim d1 As String() = Split(str, "日")
                Dim d2 As String() = Split(d1(0), "月")
                str = Year(Now) & "/" & String.Format("{0:D2}", CInt(d2(0))) & "/" & String.Format("{0:D2}", CInt(d2(1)))
                DataGridView5.Rows.Add("購入日時", str)
                str = HTMLGet(doc, "//span[@class='buyer-name']", 0, 0)
                str = StrSeikei(str)
                Dim namae As String = str
                DataGridView5.Rows.Add("名前", str)
                str = HTMLGet(doc, "//ul[@class='transact-info-table-cell']", 0, 6)
                str = StrSeikei(str)
                str = Replace(str, namae, "")
                str = Replace(str, "様", "")
                str = Replace(str, "-", "")
                str = str.Substring(0, 9)
                DataGridView5.Rows.Add("郵便番号", str)
                str = HTMLGet(doc, "//ul[@class='transact-info-table-cell']", 0, 6)
                str = StrSeikei(str)
                str = Replace(str, namae, "")
                str = Replace(str, "様", "")
                str = str.Substring(11)
                DataGridView5.Rows.Add("住所", str)
                str = Form1.TabBrowser1.SelectedTab.WebBrowser.Url.ToString
                DataGridView5.Rows.Add("URL", str)
            Else    'ヤフオク
                Dmode = 2
                str = HTMLGet(doc, "//dd[@class='decItmID']", 0, 0)
                str = Replace(str, "オークションID： ", "")
                DataGridView5.Rows.Add("オークションID", str)
                str = HTMLGet(doc, "//dd[@class='decPrice']", 0, 0)
                str = Replace(str, "円", "")
                str = Replace(str, ",", "")
                sArray = Split(str, "： ")
                DataGridView5.Rows.Add("落札価格", sArray(2))
                str = HTMLGet(doc, "//div[@class='decCnfWr']", 0, 5)
                str = Replace(str, "円）", "")
                str = Replace(str, ",", "")
                sArray = Split(str, "：")
                DataGridView5.Rows.Add("送料", sArray(1))
                str = HTMLGet(doc, "//dd[@class='decMDT']", 0, 0)
                str = Replace(str, "日", "")
                sArray = Split(str, " ")
                sArray = Split(sArray(1), "月")
                str = Year(Now) & "/" & String.Format("{0:D2}", CInt(sArray(0))) & "/" & String.Format("{0:D2}", CInt(sArray(1)))
                DataGridView5.Rows.Add("終了日時", str)
                str = HTMLGet(doc, "//div[@class='decCnfWr']", 0, 1)
                str = Replace(str, " ", "")
                str = Replace(str, "　", "")
                DataGridView5.Rows.Add("氏名", str)
                str = HTMLGet(doc, "//div[@class='decCnfWr']", 0, 2)
                str = Replace(str, "〒", "")
                DataGridView5.Rows.Add("郵便番号", str)
                str = HTMLGet(doc, "//div[@class='decCnfWr']", 0, 3)
                str = Replace(str, " ", "")
                str = Replace(str, "　", "")
                DataGridView5.Rows.Add("住所", str)
                str = HTMLGet(doc, "//div[@class='decCnfWr']", 0, 4)
                DataGridView5.Rows.Add("電話番号", str)
                str = HTMLGet(doc, "//p[@class='fnt12 decTxtNtc']", 0, 0)
                str = Replace(str, "通知先：", "")
                str = Replace(str, " [ 変更する ]", "")
                DataGridView5.Rows.Add("通知先", str)
                str = HTMLGet(doc, "//dd[@class='decBuyerID']", 0, 0)
                str = Replace(str, "落札者：", "")
                str = Regex.Replace(str, "（.*）", "")
                str = Replace(str, " ", "")
                DataGridView5.Rows.Add("落札者ID", str)
                str = Form1.TabBrowser1.SelectedTab.WebBrowser.Url.ToString
                DataGridView5.Rows.Add("URL", str)
            End If
        Catch ex As Exception
            MsgBox("取引画面では無い可能性があります", MsgBoxStyle.SystemModal)
        End Try
    End Sub

    Private Function StrSeikei(ByVal str As String)
        str = Replace(str, " ", "")
        str = Replace(str, "　", "")
        str = Replace(str, "¥", "")
        str = Replace(str, ",", "")
        str = Replace(str, "〒", "")

        Return str
    End Function

    'NE伝票入力
    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        If DataGridView5.RowCount > 0 Then
            '全角修正
            For r As Integer = 0 To DataGridView5.RowCount - 1
                Dim res As String = DataGridView5.Item(1, r).Value
                res = Form1.StrConvNumeric(res)
                res = Form1.StrConvEnglish(res)
                res = Form1.StrConvHyphen(res)
                DataGridView5.Item(1, r).Value = res
            Next

            Try
                Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("tenpo_denpyo_no").SetAttribute("value", DataGridView5.Item(1, 0).Value)
                Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("jyuchu_bi").SetAttribute("value", DataGridView5.Item(1, 3).Value)
                Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("jyuchu_name").SetAttribute("value", DataGridView5.Item(1, 4).Value)
                Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("jyuchu_yubin_bangou").SetAttribute("value", DataGridView5.Item(1, 5).Value)
                Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("jyuchu_jyusyo1").SetAttribute("value", DataGridView5.Item(1, 6).Value)
                Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("hasou_name").SetAttribute("value", DataGridView5.Item(1, 4).Value)
                Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("hasou_yubin_bangou").SetAttribute("value", DataGridView5.Item(1, 5).Value)
                Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("hasou_jyusyo1").SetAttribute("value", DataGridView5.Item(1, 6).Value)
            Catch ex As Exception
                MsgBox("伝票起票画面に入力できませんでした", MsgBoxStyle.SystemModal)
                Exit Sub
            End Try

            If Dmode = 1 Then
                Try
                    Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("jyuchu_denwa").SetAttribute("value", DataGridView5.Item(1, 0).Value)
                    Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("goukei_kin").SetAttribute("value", DataGridView5.Item(1, 2).Value)
                    Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("tesuryo_kin").SetAttribute("value", "-" & DataGridView5.Item(1, 1).Value)
                    Dim syohin_kin As Integer = CInt(DataGridView5.Item(1, 2).Value) + CInt(DataGridView5.Item(1, 1).Value)
                    Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("syohin_kin").SetAttribute("value", syohin_kin)
                    Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("textarea")("sagyosya_ran").SetAttribute("value", DataGridView5.Item(1, 7).Value)

                    Dim hCollection As HtmlElementCollection = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("select")
                    Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("select")("tenpo_code").SetAttribute("selectedIndex", "9")
                    Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("select")("hasou_kbn").SetAttribute("selectedIndex", "1")
                    Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("select")("gift_flg").SetAttribute("selectedIndex", "0")
                    Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("select")("siharai_kbn").SetAttribute("selectedIndex", "35")
                Catch ex As Exception
                    MsgBox("メルカリの情報を伝票起票画面に入力できませんでした", MsgBoxStyle.SystemModal)
                    Exit Sub
                End Try
            ElseIf Dmode = 2 Then
                Try
                    Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("jyuchu_yubin_bangou").SetAttribute("value", DataGridView5.Item(1, 5).Value)
                    Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("jyuchu_denwa").SetAttribute("value", DataGridView5.Item(1, 7).Value)
                    Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("hasou_denwa").SetAttribute("value", DataGridView5.Item(1, 7).Value)
                    Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("syohin_kin").SetAttribute("value", DataGridView5.Item(1, 1).Value)
                    Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("hasou_kin").SetAttribute("value", DataGridView5.Item(1, 2).Value)
                    Dim goukei_kin As Integer = CInt(DataGridView5.Item(1, 2).Value) + CInt(DataGridView5.Item(1, 1).Value)
                    Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("goukei_kin").SetAttribute("value", goukei_kin)
                    If DataGridView5.Item(1, 8).Value <> "error" Then
                        Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("jyuchu_mail_adr").SetAttribute("value", DataGridView5.Item(1, 8).Value)
                    End If
                    Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("jyuchu_kana").SetAttribute("value", DataGridView5.Item(1, 9).Value)
                    Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("textarea")("sagyosya_ran").SetAttribute("value", DataGridView5.Item(1, 10).Value)

                    Dim hCollection As HtmlElementCollection = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("select")
                    Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("select")("tenpo_code").SetAttribute("selectedIndex", "10")
                    Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("select")("hasou_kbn").SetAttribute("selectedIndex", "1")
                    Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("select")("gift_flg").SetAttribute("selectedIndex", "0")
                    Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("select")("siharai_kbn").SetAttribute("selectedIndex", "35")
                Catch ex As Exception
                    MsgBox("ヤフオクの情報を伝票起票画面に入力できませんでした", MsgBoxStyle.SystemModal)
                    Exit Sub
                End Try
            Else
                MsgBox("情報取得内容が整合していません", MsgBoxStyle.SystemModal)
                Exit Sub
            End If
        End If
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        DGV6.Rows.Clear()

        'webbrowserのdocumentText文字化け対策
        Dim buf As Byte() = New Byte(Form1.TabBrowser1.SelectedTab.WebBrowser.DocumentStream.Length) {}
        Form1.TabBrowser1.SelectedTab.WebBrowser.DocumentStream.Read(buf, 0, CInt(Form1.TabBrowser1.SelectedTab.WebBrowser.DocumentStream.Length))
        Dim ec As System.Text.Encoding = System.Text.Encoding.GetEncoding(Form1.TabBrowser1.SelectedTab.WebBrowser.Document.Encoding)
        Dim htmls As String = ec.GetString(buf)

        Dim doc As HtmlAgilityPack.HtmlDocument = New HtmlAgilityPack.HtmlDocument
        doc.LoadHtml(htmls)

        If InStr(Form1.TabBrowser1.SelectedTab.WebBrowser.Url.Host, "rakuten") > 0 Then
            For i As Integer = 0 To 19
                DGV6.Rows.Add(1)
                Try
                    Dim gazo As String = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("image_alt" & i + 1).GetAttribute("value")
                    Dim url As String = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("image_url" & i + 1).GetAttribute("value")
                    DGV6.Item(0, DGV6.RowCount - 2).Value = gazo
                    DGV6.Item(1, DGV6.RowCount - 2).Value = url
                Catch ex As Exception

                End Try
            Next
        ElseIf InStr(Form1.TabBrowser1.SelectedTab.WebBrowser.Url.Host, "yahoo") > 0 Then
            Dim pattern As String = "http.*?(item\-shopping).*?[a-z]{1,2}[0-9_\-]{3,6}"
            Dim sc As String = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("textarea")("__submit__caption").GetAttribute("value")
            Dim mc As MatchCollection = Regex.Matches(sc, pattern)
            For Each m As Match In mc
                DGV6.Rows.Add(1)
                DGV6.Item(1, DGV6.RowCount - 2).Value = m.Value.ToString
            Next
        Else
            MsgBox("取得ページが間違っているようです", MsgBoxStyle.SystemModal)
        End If
    End Sub

    '選択行？追加
    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        Dim selRowArray As New ArrayList
        For i As Integer = 0 To DGV6.SelectedCells.Count - 1
            If Not selRowArray.Contains(DGV6.SelectedCells(i).RowIndex) Then
                selRowArray.Add(DGV6.SelectedCells(i).RowIndex)
            End If
        Next
        For i As Integer = 0 To selRowArray.Count - 1
            Dim imgPath As String = DGV6.Item(1, selRowArray(i)).Value
            If imgPath = "" Then
                Continue For
            End If

            Dim newPath As String = imgPath & "?" & Format(Now, "MMddHHmmss")
            If InStr(imgPath, "?") > 0 Then
                Dim iP As String() = Split(imgPath, "?")
                newPath = iP(0) & "?" & Format(Now, "MMddHHmmss")
            End If
            DGV6.Item(1, selRowArray(i)).Value = newPath
        Next

        MsgBox("更新文字を追加・更新しました")
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        Dim selRowArray As New ArrayList
        For i As Integer = 0 To DGV6.SelectedCells.Count - 1
            If Not selRowArray.Contains(DGV6.SelectedCells(i).RowIndex) Then
                selRowArray.Add(DGV6.SelectedCells(i).RowIndex)
            End If
        Next
        For i As Integer = 0 To selRowArray.Count - 1
            Dim imgPath As String = DGV6.Item(1, selRowArray(i)).Value
            If imgPath = "" Then
                Continue For
            End If

            If InStr(imgPath, "?") > 0 Then
                Dim iP As String() = Split(imgPath, "?")
                DGV6.Item(1, selRowArray(i)).Value = iP(0)
            End If
        Next

        MsgBox("更新文字を削除しました")
    End Sub


    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        If InStr(Form1.TabBrowser1.SelectedTab.WebBrowser.Url.Host, "rakuten") > 0 Then
            For r As Integer = 0 To 19
                If DGV6.Item(0, r).Value <> "" Then
                    Try '商品画像(1)～(20)
                        If DGV6.Item(0, r).Value <> "" Then
                            Dim value1 As String = DGV6.Item(0, r).Value
                            Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("image_alt" & r + 1).SetAttribute("value", value1)
                        Else
                            Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("image_alt" & r + 1).SetAttribute("value", "")
                        End If
                        If DGV6.Item(1, r).Value <> "" Then
                            Dim value2 As String = DGV6.Item(1, r).Value
                            Dim value2array As String() = Split(value2, "?")
                            Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("image_url" & r + 1).SetAttribute("value", value2array(0))
                        Else
                            Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("image_url" & r + 1).SetAttribute("value", "")
                        End If
                    Catch ex As Exception
                        MsgBox("情報を書き込めませんでした", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
                    End Try
                End If
            Next

            'スマホ・PC用商品説明文
            'Dim template As String() = New String() {"<img src='_url_' alt=' width=' 500'='' width='100%' '=''>", "<img src='_url_' width=100%>"}
            Dim template As String() = New String() {"<img src='_url_' width='100%'>", "<img src='_url_' width=100%>"}
            Dim tags As String() = New String() {"smart_caption", "display_caption"}
            Dim wrightStr As String() = New String() {"", ""}
            wrightStr(0) &= "<!--imgS-->" & vbCrLf & "<p>"
            wrightStr(0) &= "_IMG_"
            wrightStr(0) &= "</p><br>" & vbLf & "<!--imgE-->"
            wrightStr(1) &= "<!--img-->" & vbCrLf
            wrightStr(1) &= "_IMG_"
            wrightStr(1) &= "<!--img-->"
            Dim pattern As String() = New String() {"<!--imgS.*imgE-->", "<!--img-->.*<!--img-->"}

            For i As Integer = 0 To 1
                Dim readStr As String = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("textarea")(tags(i)).InnerText
                Dim ws As String = ""
                For r As Integer = 0 To 19
                    If DGV6.Item(0, r).Value <> "" Then
                        If Not (i = 0 And DGV6.Item(0, r).Value = "info") Then
                            ws &= Replace(template(i), "_url_", DGV6.Item(1, r).Value) & vbCrLf
                        End If
                    End If
                Next
                'If Regex.IsMatch(readStr, "<!--imgS.*imgE-->", RegexOptions.Singleline) Then
                '    MsgBox("true")
                'Else
                '    MsgBox("false")
                'End If
                Dim str As String = Replace(wrightStr(i), "_IMG_", ws)
                Dim r2wStr As String = Regex.Replace(readStr, pattern(i), str, RegexOptions.Singleline)
                Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("textarea")(tags(i)).InnerText = r2wStr
            Next
            MsgBox("書き込みました", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
        ElseIf InStr(Form1.TabBrowser1.SelectedTab.WebBrowser.Url.Host, "yahoo") > 0 Then
            Dim scNameArray As String() = {"__submit__caption", "__submit__sp_additional"}
            For k As Integer = 0 To scNameArray.Length - 1
                Dim pattern As String = "http.*?(item\-shopping).*?[a-z]{1,2}[0-9_\-]{3,6}"
                Dim sc As String = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("textarea")(scNameArray(k)).GetAttribute("value")
                sc = SeikeiHTML(sc)
                Dim scArray As String() = Split(sc, vbCrLf)
                'Dim mc As MatchCollection = Regex.Matches(sc, pattern)
                Dim str As String = ""
                Dim addNum As Integer = 0
                For i As Integer = 0 To scArray.Length - 1
                    If Regex.IsMatch(scArray(i), pattern) Then
                        If addNum = 0 Then
                            For r As Integer = 0 To DGV6.RowCount - 2
                                str &= Regex.Replace(scArray(i), pattern, DGV6.Item(1, r).Value.ToString) & vbCrLf
                            Next
                        Else
                            '行削除
                        End If
                        addNum += 1
                    Else
                        str &= scArray(i) & vbCrLf
                    End If
                Next
                Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("textarea")(scNameArray(k)).SetAttribute("value", str)
            Next
            MsgBox("書き込みました", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
        Else
            MsgBox("書き込みするページが間違っているようです", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
        End If
    End Sub


    Public Shared CnAccdb As String = ""
    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        If TextBox14.Text = "" Then
            Exit Sub
        End If

        TextBox14.Text = Regex.Replace(TextBox14.Text, "-|\s|　", "")

        '=======================================================
        CnAccdb = Path.GetDirectoryName(Form1.appPath) & "\db\ken_all.accdb"
        '--------------------------------------------------------------
        Dim Cn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CnAccdb)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable
        Dim ds As New DataSet
        Cn.Open()
        '--------------------------------------------------------------

        Dim whereCheck As String = ""
        If TextBox14.Text <> "" Then
            whereCheck = " ([郵便番号] = '" & TextBox14.Text & "')"
        End If

        SQLCm.CommandText = "SELECT * FROM [X-ken-all2] WHERE" & whereCheck
        Adapter.Fill(Table)

        Dim addr As String = ""
        Dim addr_kana As String = ""
        If Table.Rows.Count <> 0 Then
            addr &= Table.Rows(0)("県名").ToString
            addr &= Table.Rows(0)("市町村名").ToString
            addr &= Table.Rows(0)("町名").ToString
            addr_kana &= Table.Rows(0)("県名読み").ToString
            addr_kana &= Table.Rows(0)("市町村名読み").ToString
            addr_kana &= Table.Rows(0)("町名読み").ToString
            TextBox27.Text = KanaConverter.HalfwidthKatakanaToKatakana(Table.Rows(0)("県名読み").ToString)
            TextBox28.Text = KanaConverter.HalfwidthKatakanaToKatakana(Table.Rows(0)("市町村名読み").ToString)
            TextBox29.Text = KanaConverter.HalfwidthKatakanaToKatakana(Table.Rows(0)("町名読み").ToString)
        Else
            MsgBox("見つかりません", MsgBoxStyle.SystemModal)
        End If

        '--------------------------------------------------------------
        Table.Dispose()
        Adapter.Dispose()
        SQLCm.Dispose()
        Cn.Dispose()
        '--------------------------------------------------------------
        '=======================================================

        TextBox15.Text = addr
        TextBox25.Text = KanaConverter.HalfwidthKatakanaToKatakana(addr_kana)
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        If TextBox16.Text = "" Then
            Exit Sub
        End If

        TextBox17.Text = AddrHenkan()
        TextBox26.Text = TextBox15.Text & AddrHenkan2()
        TxtCheck()
        If Not CheckBox1.Checked Then
            TextBox19.Text = Addr2Post()
        End If
        PostCheck()
        TextBox17.Text = Regex.Replace(TextBox17.Text, "\s|　", "")
    End Sub

    Private Sub TxtCheck()
        '変換の状態を調べチェックボタンを変更する
        If Regex.IsMatch(TextBox17.Text,
            "[\p{IsCJKUnifiedIdeographs}" &
            "\p{IsCJKCompatibilityIdeographs}" &
            "\p{IsCJKUnifiedIdeographsExtensionA}]|" &
            "[\uD840-\uD869][\uDC00-\uDFFF]|\uD869[\uDC00-\uDEDF]") Then
            CheckBox1.Checked = False
        Else
            If Regex.IsMatch(TextBox15.Text,
                "[\p{IsCJKUnifiedIdeographs}" &
                "\p{IsCJKCompatibilityIdeographs}" &
                "\p{IsCJKUnifiedIdeographsExtensionA}]|" &
                "[\uD840-\uD869][\uDC00-\uDFFF]|\uD869[\uDC00-\uDEDF]") Then
                CheckBox1.Checked = True
            Else
                CheckBox1.Checked = False
            End If
        End If
        If Regex.IsMatch(TextBox21.Text, ".*[a-zA-z].*") Then
            CheckBox6.Checked = False
        Else
            CheckBox6.Checked = True
        End If
    End Sub

    Private Sub PostCheck()
        If IsNumeric(TextBox19.Text) And TextBox19.Text <> TextBox14.Text Then
            CheckBox2.Checked = True
        Else
            CheckBox2.Checked = False
        End If
    End Sub

    Private Function AddrHenkan() As String
        'TextBox16.Text = TextBox16.Text.ToLower
        TextBox16.Text = Replace(TextBox16.Text, "　", " ")
        TextBox16.Text = Regex.Replace(TextBox16.Text, "-ken", "ken ", RegexOptions.IgnoreCase)
        TextBox16.Text = Regex.Replace(TextBox16.Text, "yo-to", "yoto ", RegexOptions.IgnoreCase)
        TextBox16.Text = Regex.Replace(TextBox16.Text, "-fu", "fu ", RegexOptions.IgnoreCase)
        TextBox16.Text = Regex.Replace(TextBox16.Text, "-hu", "fu ", RegexOptions.IgnoreCase)
        TextBox16.Text = Regex.Replace(TextBox16.Text, "-ku", "ku ", RegexOptions.IgnoreCase)
        TextBox16.Text = Regex.Replace(TextBox16.Text, "-city", "shi ", RegexOptions.IgnoreCase)
        TextBox16.Text = Regex.Replace(TextBox16.Text, "-gun ", "gun", RegexOptions.IgnoreCase)
        TextBox16.Text = Regex.Replace(TextBox16.Text, "hokkaido", "Hokkaidou", RegexOptions.IgnoreCase)
        TextBox16.Text = Regex.Replace(TextBox16.Text, "tokyoto", "Toukyouto", RegexOptions.IgnoreCase)
        TextBox16.Text = Regex.Replace(TextBox16.Text, "kyotofu", "Kyoutofu", RegexOptions.IgnoreCase)
        TextBox16.Text = Regex.Replace(TextBox16.Text, "osakafu", "Ousakafu", RegexOptions.IgnoreCase)
        TextBox16.Text = Regex.Replace(TextBox16.Text, "hyogoken", "Hyougoken", RegexOptions.IgnoreCase)
        TextBox16.Text = Regex.Replace(TextBox16.Text, "kochiken", "Kouchiken", RegexOptions.IgnoreCase)
        TextBox16.Text = Regex.Replace(TextBox16.Text, "ohitaken", "Ouitaken", RegexOptions.IgnoreCase)
        TextBox16.Text = Regex.Replace(TextBox16.Text, "oitaken", "Ouitaken", RegexOptions.IgnoreCase)
        TextBox16.Text = Regex.Replace(TextBox16.Text, "([A-Za-z])\-", "$1ー")

        TextBox18.Text = KanaConverter.RomajiToHiragana(TextBox16.Text)
        TextBox18.Text = KanaConverter.HiraganaToKatakana(TextBox18.Text)

        ''=======================================================
        'CnAccdb = Path.GetDirectoryName(Form1.appPath) & "\db\Address.accdb"
        ''--------------------------------------------------------------
        'Dim Cn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CnAccdb)
        'Dim SQLCm As OleDbCommand = Cn.CreateCommand
        'Dim Adapter As New OleDbDataAdapter(SQLCm)
        'Dim Table As New DataTable
        'Dim ds As New DataSet
        'Cn.Open()
        ''--------------------------------------------------------------

        'Dim nameArray As ArrayList = New ArrayList From {
        '    "都道府県カナ,都道府県",
        '    "市区町村カナ,市区町村",
        '    "町域カナ,町域",
        '    "字丁目カナ,字丁目"
        '}

        'Dim adrrs As String() = Split(TextBox18.Text, " ")
        'Dim adrrArray As ArrayList = New ArrayList(adrrs)
        'Dim whereCheck As String = ""
        Dim addr As String = ""

        'For i As Integer = 0 To 2
        '    Dim header As String() = Split(nameArray(i), ",")
        '    If whereCheck = "" Then
        '        whereCheck = " ([" & header(0) & "] = '" & adrrArray(i) & "')"
        '    Else
        '        whereCheck &= " AND ([" & header(0) & "] = '" & adrrArray(i) & "')"
        '    End If
        'Next
        'SQLCm.CommandText = "SELECT * FROM [AD_住所] WHERE" & whereCheck
        'Adapter.Fill(Table)
        'If Table.Rows.Count > 0 Then
        '    For i As Integer = 0 To 2
        '        Dim header As String() = Split(nameArray(i), ",")
        '        addr &= Table.Rows(0)(header(1)).ToString & " "
        '    Next
        'End If
        'Table.Clear()

        'For i As Integer = 3 To adrrArray.Count - 1
        '    Dim ad As String = adrrArray(i)
        '    For k As Integer = 0 To kanaH.Length - 1
        '        ad = Regex.Replace(ad, kanaH(k), kanjiH(k))
        '    Next
        '    addr &= ad
        'Next

        'For k As Integer = 0 To kanaG.Length - 1
        '    Dim kG As String() = Split(kanaG(k), "|")
        '    Dim m As Match = Regex.Match(addr, kG(0))
        '    If m.Value <> "" Then
        '        addr = Regex.Replace(addr, kG(1), kanjiG(k))
        '    End If
        'Next

        ''--------------------------------------------------------------
        'Table.Dispose()
        'Adapter.Dispose()
        'SQLCm.Dispose()
        'Cn.Dispose()
        ''--------------------------------------------------------------
        ''=======================================================

        Return addr
    End Function

    Dim romaF As String() = New String() {"Room"}
    Dim kanaF As String() = New String() {"ルーム"}
    Dim kanaH As String() = New String() {"チョウメ", "バンチ", "カブシキガイシャ", "ゴウシツ", "ダンチ", "シエイジュウタク", "ケンエイジュウタク|ケネイジュウタク", "^アザ$"}
    Dim kanjiH As String() = New String() {"丁目", "番地", "株式会社", "号室", "団地", "市営住宅", "県営住宅", "字"}
    Dim kanaG As String() = New String() {"([0-9]){1}ゴウ", "([0-9]){1}バン", "([ァ-ー]){1}\-"}
    Dim kanjiG As String() = New String() {"$1号", "$1番", "$1ー"}
    Private Function AddrHenkan2()
        Dim addr As String = TextBox18.Text
        If addr = "" Then
            Return addr
            Exit Function
        End If

        addr = Replace(addr, " ", "")

        'カタカナの揺らぎに対応
        Dim TB As TextBox() = {TextBox27, TextBox28, TextBox29}
        Dim patternArray1 As String() = New String() {"ウ", "ズ", "ヅ", "オオ", "オウ", "マチ", "チョウ", "ハタ", "ハタケ", "ンニャ", "シ", "イチ", "ツ", "ッ"}
        Dim patternArray2 As String() = New String() {"", "ヅ", "ズ", "オウ", "オオ", "チョウ", "マチ", "ハタケ", "ハタ", "ンヤ", "イチ", "シ", "ッ", "ツ"}
        For k As Integer = 0 To TB.Length - 1
            If addr <> "" Then
                Dim pattern1 As String = TB(k).Text
                If Regex.IsMatch(addr, "^" & pattern1) Then
                    addr = Regex.Replace(addr, "^" & pattern1, "")
                Else
                    For i As Integer = 0 To patternArray1.Length - 1
                        Dim pattern2 As String = Replace(pattern1, patternArray1(i), patternArray2(i))
                        If Regex.IsMatch(addr, "^" & pattern2) Then
                            addr = Regex.Replace(addr, "^" & pattern2, "")
                            Exit For
                        End If
                    Next
                End If
            End If
        Next

        For k As Integer = 0 To kanaH.Length - 1
            addr = Regex.Replace(addr, kanaH(k), kanjiH(k))
        Next

        For k As Integer = 0 To kanaG.Length - 1
            Dim m As Match = Regex.Match(addr, kanaG(k))
            If m.Value <> "" Then
                addr = Regex.Replace(addr, kanaG(k), kanjiG(k))
            End If
        Next

        Return addr
    End Function

    Private Function Addr2Post() As String
        Dim adrrs As String() = Split(TextBox17.Text, " ")
        Dim adrrArray As ArrayList = New ArrayList(adrrs)

        '=======================================================
        CnAccdb = Path.GetDirectoryName(Form1.appPath) & "\db\ken_all.accdb"
        '--------------------------------------------------------------
        Dim Cn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CnAccdb)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable
        Dim ds As New DataSet
        Cn.Open()
        '--------------------------------------------------------------

        Dim post As String = ""
        Dim whereCheck As String = ""
        If adrrs.Count > 2 Then
            whereCheck = " ([県名] = '" & adrrs(0) & "')"
            whereCheck &= " AND ([市町村名] = '" & adrrs(1) & "')"
            whereCheck &= " AND ([町名] = '" & adrrs(2) & "')"

            SQLCm.CommandText = "SELECT * FROM [X-ken-all2] WHERE" & whereCheck
            Adapter.Fill(Table)

            If Table.Rows.Count > 0 Then
                post = Table.Rows(0)("郵便番号").ToString
            End If
        Else
            post = ""
        End If

        '--------------------------------------------------------------
        Table.Dispose()
        Adapter.Dispose()
        SQLCm.Dispose()
        Cn.Dispose()
        '--------------------------------------------------------------
        '=======================================================

        Return post
    End Function

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        'webbrowserのdocumentText文字化け対策
        Dim buf As Byte() = New Byte(Form1.TabBrowser1.SelectedTab.WebBrowser.DocumentStream.Length) {}
        Form1.TabBrowser1.SelectedTab.WebBrowser.DocumentStream.Read(buf, 0, CInt(Form1.TabBrowser1.SelectedTab.WebBrowser.DocumentStream.Length))
        Dim wbEnc As String = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.Encoding
        If wbEnc = "_autodetect" Then
            wbEnc = "utf-8"
        End If
        Dim ec As System.Text.Encoding = System.Text.Encoding.GetEncoding(wbEnc)
        Dim htmls As String = ec.GetString(buf)

        Dim doc As HtmlAgilityPack.HtmlDocument = New HtmlAgilityPack.HtmlDocument
        doc.LoadHtml(htmls)

        TextBox24.Text = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("hasou_jyusyo1").GetAttribute("value")
        TextBox23.Text = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("hasou_jyusyo2").GetAttribute("value")

        '漢字が入っていたら中止
        For i As Integer = 0 To TextBox24.Text.Length - 1
            If IsKanji(TextBox24.Text(i)) Then
                Exit Sub
            End If
        Next

        TextBox16.Text = ""
        For i As Integer = 0 To 19
            DGV6.Rows.Add(1)
            Try
                Dim hYubin As String = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("hasou_yubin_bangou").GetAttribute("value")
                TextBox14.Text = hYubin
                Dim hAddr1 As String = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("hasou_jyusyo1").GetAttribute("value")
                Dim hAddr2 As String = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("hasou_jyusyo2").GetAttribute("value")
                If TextBox16.Text = "" Then
                    TextBox16.Text = hAddr1
                Else
                    '住所1最後数字と住所2最初数字の場合「-」追加
                    If IsNumeric(hAddr1.Substring(hAddr1.Length - 1, 1)) And IsNumeric(hAddr2.Substring(0, 1)) Then
                        TextBox16.Text = hAddr1 & "-" & hAddr2
                    Else
                        TextBox16.Text = hAddr1 & " " & hAddr2
                    End If
                End If
            Catch ex As Exception

            End Try
        Next

        Button17.PerformClick()
        Button18.PerformClick()

        Dim hNamae As String = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.GetElementsByTagName("input")("hasou_name").GetAttribute("value")
        hNamae = Regex.Replace(hNamae, "\s|　", "")
        TextBox20.Text = hNamae
        TextBox21.Text = KanaConverter.RomajiToHiragana(hNamae)
        TextBox21.Text = KanaConverter.HiraganaToKatakana(TextBox21.Text)
    End Sub

    ''' <summary>
    ''' 漢字かどうかを調べる
    ''' </summary>
    ''' <param name="c">Char</param>
    ''' <returns>Boolean</returns>
    Public Shared Function IsKanji(ByVal c As Char) As Boolean
        'CJK統合漢字、CJK互換漢字、CJK統合漢字拡張Aの範囲にあるか調べる
        Return (ChrW(&H4E00) <= c AndAlso c <= ChrW(&H9FCF)) OrElse
        (ChrW(&HF900) <= c AndAlso c <= ChrW(&HFAFF)) OrElse
        (ChrW(&H3400) <= c AndAlso c <= ChrW(&H4DBF))
    End Function

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        Dim hd As HtmlDocument = Form1.TabBrowser1.SelectedTab.WebBrowser.Document

        Dim addr As String = ""
        If CheckBox1.Checked Then
            addr = TextBox15.Text
        End If
        If CheckBox3.Checked Then
            addr &= TextBox17.Text
        ElseIf CheckBox5.Checked Then
            addr = TextBox26.Text
        End If
        hd.GetElementsByTagName("input")("hasou_jyusyo1").SetAttribute("value", addr)
        hd.GetElementsByTagName("input")("hasou_jyusyo2").SetAttribute("value", "")

        If CheckBox4.Checked Then
            hd.GetElementsByTagName("input")("jyuchu_jyusyo1").SetAttribute("value", addr)
            hd.GetElementsByTagName("input")("jyuchu_jyusyo2").SetAttribute("value", "")
        End If

        If CheckBox2.Checked Then
            hd.GetElementsByTagName("input")("hasou_yubin_bangou").SetAttribute("value", TextBox19.Text)
        End If

        Dim namae As String = TextBox21.Text
        If CheckBox6.Checked Then
            hd.GetElementsByTagName("input")("hasou_kana").SetAttribute("value", TextBox20.Text)
            hd.GetElementsByTagName("input")("hasou_name").SetAttribute("value", namae)
            'namae = KanaConverter.HiraganaToKatakana(namae)
            'hd.GetElementsByTagName("input")("hasou_kana").SetAttribute("value", namae)
        End If
    End Sub

    Dim wbURL As String = ""
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Try
            Dim nowURL As String = Form1.TabBrowser1.SelectedTab.WebBrowser.Url.ToString
            If InStr(nowURL, "ne81.next-engine.com/Userjyuchu/jyuchuInp?kensaku_denpyo_no=") > 0 Then
                If InStr(Form1.TabBrowser1.SelectedTab.WebBrowser.DocumentText, "</html>") > 0 Then
                    If nowURL <> wbURL Then
                        wbURL = nowURL
                        Button15.PerformClick()
                    End If
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ToolStripButton1_MouseEnter(sender As Object, e As EventArgs) Handles ToolStripButton1.MouseEnter
        Me.Activate()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Select Case ComboBox1.SelectedItem
            Case "1:メール便", "4:メールのみ", "7:メールのみ"
                If ToolStripComboBox1.SelectedItem = "楽天-通販の暁用" Then
                    ComboBox4.SelectedIndex = -1
                Else
                    ComboBox4.SelectedItem = "1:メール便"
                End If
                TextBox2.Text = ""
            Case "1:宅配メール", "2:宅配のみ", "2:宅配便", "6:宅配のみ"
                If ToolStripComboBox1.SelectedItem = "楽天-通販の暁用" Then
                    ComboBox4.SelectedIndex = -1
                Else
                    ComboBox4.SelectedItem = "2:宅配便"
                End If
                TextBox2.Text = ""
            Case "3:大型のみ", "2:宅配便（大型）", "3:使用不可（大型）"
                ComboBox4.SelectedItem = "3:大型宅配便"
                If ToolStripComboBox1.SelectedItem = "楽天-通販の暁用" Then
                    TextBox2.Text = 1260
                Else
                    TextBox2.Text = ""
                End If
                ComboBox4.SelectedIndex = -1
            Case "5:定形外郵便", "4:定形外郵便", "8:定形外郵便"
                If ToolStripComboBox1.SelectedItem = "楽天-通販の暁用" Then
                    ComboBox4.SelectedIndex = -1
                Else
                    ComboBox4.SelectedItem = "4:定形外郵便"
                End If
                TextBox2.Text = ""
        End Select
    End Sub

    Private Sub TextBox30_TextChanged(sender As Object, e As EventArgs) Handles TextBox30.TextChanged
        akatuki_price()
    End Sub

    Private Sub akatuki_price()
        Dim price As String = TextBox30.Text.Trim
        Dim price_rs As Integer = 0
        If price <> "" And ComboBox1.SelectedIndex <> -1 Then
            If IsNumeric(price) Then
                If price = Int(price) Then
                    Select Case ComboBox1.SelectedItem.ToString
                        Case "2:宅配のみ"
                            If price > 3980 Then
                                TextBox2.Text = ""
                                'price_rs = Math.Round((price + 400) * 1.04)
                            Else
                                TextBox2.Text = ""
                                'price_rs = Math.Round((price - 400) * 1.04)
                            End If
                        Case "3:大型のみ"
                            TextBox2.Text = 1260
                            'price_rs = Math.Round((price - 1260) * 1.04)
                        Case "4:メールのみ", "5:定形外郵便"
                            TextBox2.Text = ""
                            'price_rs = Math.Round((price - 80) * 1.04)
                    End Select
                    TextBox13.Text = price_rs
                    If price_rs < 0 Or price_rs = 0 Then
                        TextBox13.BackColor = Color.Red
                    Else
                        TextBox13.BackColor = Color.Yellow
                    End If
                End If
            End If
        Else
            TextBox13.Text = ""
            TextBox2.Text = ""
            TextBox13.BackColor = Color.Yellow
        End If


    End Sub

    Private Sub ComboBox1_TextChanged(sender As Object, e As EventArgs) Handles ComboBox1.TextChanged
        akatuki_price()
    End Sub

    'Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
    '    If ToolStripComboBox1.SelectedIndex = 7 Then
    '        If ComboBox4.SelectedIndex = 3 Then
    '            TextBox2.Text = 1260
    '        Else
    '            TextBox2.Text = ""
    '        End If
    '    End If
    'End Sub

    Private Sub resultEventHandler(sender As Object, e As EventArgs)
        MessageBox.Show("Loaded")
    End Sub
End Class

