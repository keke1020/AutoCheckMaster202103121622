Imports System.IO
Imports System.Data.OleDb

Public Class OtherDialog
    Dim eigyousyo As ArrayList = New ArrayList
    Dim tempoList As ArrayList = New ArrayList

    Private Sub OtherDialog_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\config\佐川営業所一覧.csv"
        eigyousyo.Clear()
        Using sr As New System.IO.StreamReader(fName, System.Text.Encoding.Default)
            Do While Not sr.EndOfStream
                Dim s As String = sr.ReadLine
                If s <> "" Then
                    eigyousyo.Add(s)
                End If
            Loop
        End Using

        'fName = Path.GetDirectoryName(Form1.appPath) & "\会社情報.txt"
        'Dim flag As Boolean = False
        'Using sr As New System.IO.StreamReader(fName, System.Text.Encoding.Default)
        '    Do While Not sr.EndOfStream
        '        Dim s As String = sr.ReadLine
        '        If s = "//店舗リスト" Then
        '            flag = True
        '        ElseIf s = "" Then

        '        ElseIf flag = True Then
        '            tempoList.Add(s)
        '            Dim tArray As String() = Split(s, "=")
        '            ComboBox6.Items.Add(tArray(0))
        '        End If
        '    Loop
        'End Using
        'If ComboBox6.Items.Count > 0 Then
        '    ComboBox6.SelectedIndex = 0
        'End If

        ComboBox2.SelectedIndex = 0
        ComboBox3.SelectedIndex = 26

        ComboBox11.SelectedIndex = 0
        ComboBox12.SelectedIndex = 0
        ComboBox7.SelectedIndex = 0
        ComboBox13.SelectedIndex = 0
        ComboBox14.SelectedIndex = 0
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Visible = False
    End Sub

    Public Shared qoo10mode As Integer = 0
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Dim csvFNum As Integer = Form1.CSV_FORMS(0).ToolStripStatusLabel5.Text - 1
        'Dim csvForm As Csv = Form1.CSV_FORMS(csvFNum)
        If RadioButton1.Checked = True Then     '佐川急便
            qoo10mode = 0
        ElseIf RadioButton2.Checked = True Then 'ネクストエンジン（追跡・発送予定）
            qoo10mode = 3
        ElseIf RadioButton3.Checked = True Then '日本通運
            qoo10mode = 1
        ElseIf RadioButton5.Checked = True Then 'ヤマト運輸
            qoo10mode = 4
        ElseIf RadioButton4.Checked = True Then 'ネクストエンジン（同梱キャンセル分）
            qoo10mode = 5
        ElseIf RadioButton6.Checked = True Then 'ネクストエンジン（予約）
            qoo10mode = 6
        ElseIf RadioButton7.Checked = True Then 'ネクストエンジン（確認用）
            qoo10mode = 7
        End If
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Visible = False
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        ComboBox1.Items.Clear()
        Dim flag As Boolean = False
        For i As Integer = 0 To eigyousyo.Count - 1
            If InStr(eigyousyo(i), ",") > 0 Then
                If flag = True Then
                    Dim array As String() = Split(eigyousyo(i), ",")
                    ComboBox1.Items.Add(array(0))
                End If
            Else
                If ComboBox3.SelectedItem = eigyousyo(i) Then
                    flag = True
                Else
                    flag = False
                End If
            End If
        Next
        ComboBox1.SelectedIndex = 0
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        For i As Integer = 0 To eigyousyo.Count - 1
            If InStr(eigyousyo(i), ",") > 0 Then
                Dim array As String() = Split(eigyousyo(i), ",")
                If ComboBox1.SelectedItem = array(0) Then
                    TextBox1.Text = array(2)
                    Exit For
                End If
            End If
        Next
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.Visible = False
    End Sub

    '一時保存
    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        TextBox9.Text = TextBox2.Text
        TextBox8.Text = TextBox3.Text
        ComboBox5.SelectedItem = ComboBox4.SelectedItem
        Dim str As String = Replace(TextBox4.Text, ComboBox5.SelectedItem, "")
        For i As Integer = 0 To ComboBox15.Items.Count - 1
            If InStr(str, ComboBox15.Items(i)) > 0 Then
                ComboBox15.SelectedIndex = i
            End If
        Next
        str = Replace(str, ComboBox15.SelectedItem, "")
        For i As Integer = 0 To ComboBox16.Items.Count - 1
            If InStr(str, ComboBox16.Items(i)) > 0 Then
                ComboBox16.SelectedIndex = i
            End If
        Next
        str = Replace(str, ComboBox16.SelectedItem, "")
        TextBox7.Text = str
        TextBox6.Text = TextBox5.Text
        ComboBox17.SelectedItem = ComboBox14.SelectedItem
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            DateTimePicker1.Enabled = True
        Else
            DateTimePicker1.Enabled = False
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            TextBox19.Enabled = True
        Else
            TextBox19.Enabled = False
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            ComboBox9.Enabled = True
            ComboBox8.Enabled = True
            TextBox20.Enabled = True
        Else
            ComboBox9.Enabled = False
            ComboBox8.Enabled = False
            TextBox20.Enabled = False
        End If
    End Sub

    Private Sub ComboBox9_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox9.SelectedIndexChanged
        ComboBox8.Items.Clear()
        Dim flag As Boolean = False
        For i As Integer = 0 To eigyousyo.Count - 1
            If InStr(eigyousyo(i), ",") > 0 Then
                If flag = True Then
                    Dim array As String() = Split(eigyousyo(i), ",")
                    ComboBox8.Items.Add(array(0))
                End If
            Else
                If ComboBox9.SelectedItem = eigyousyo(i) Then
                    flag = True
                Else
                    flag = False
                End If
            End If
        Next
        ComboBox8.SelectedIndex = 0
    End Sub

    Private Sub ComboBox8_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox8.SelectedIndexChanged
        For i As Integer = 0 To eigyousyo.Count - 1
            If InStr(eigyousyo(i), ",") > 0 Then
                Dim array As String() = Split(eigyousyo(i), ",")
                If ComboBox8.SelectedItem = array(0) Then
                    TextBox20.Text = array(2)
                    Exit For
                End If
            End If
        Next
    End Sub

    Public Sub CsvToDenpyo(ByVal dgv As DataGridView)
        If dgv.RowCount = 0 Then
            Exit Sub
        End If

        dgvWrite = dgv

        'datagridviewのヘッダーテキストをコレクションに取り込む
        Dim DGVheaderCheck As ArrayList = New ArrayList
        Dim headerStr As String = ""
        For c As Integer = 0 To dgv.Columns.Count - 1
            DGVheaderCheck.Add(dgv.Item(c, 0).Value)
            headerStr &= dgv.Item(c, 0).Value
        Next c

        Dim selR As Integer = dgvWrite.SelectedCells(0).RowIndex

        'Try
        '佐川急便------------------------------------------------------------------------------------
        If InStr(headerStr, "便種（スピードを選択）便種（クール便指定）") > 0 And dgvWrite.RowCount > 1 Then
            ComboBox11.SelectedItem = "佐川急便ビズロジ"

            TextBox2.Text = dgvWrite.Item(DGVheaderCheck.IndexOf("お届け先電話番号"), selR).Value
            TextBox3.Text = dgvWrite.Item(DGVheaderCheck.IndexOf("お届け先郵便番号"), selR).Value
            For i As Integer = 0 To ComboBox4.Items.Count - 1
                If InStr(dgvWrite.Item(DGVheaderCheck.IndexOf("お届け先住所１"), selR).Value, ComboBox4.Items(i)) > 0 Then
                    ComboBox4.SelectedIndex = i
                    Exit For
                End If
            Next
            TextBox4.Text = Replace(dgvWrite.Item(DGVheaderCheck.IndexOf("お届け先住所１"), selR).Value, ComboBox4.SelectedItem, "")
            TextBox4.Text = LTrim(TextBox4.Text)
            TextBox4.Text &= dgvWrite.Item(DGVheaderCheck.IndexOf("お届け先住所２"), selR).Value
            TextBox4.Text &= dgvWrite.Item(DGVheaderCheck.IndexOf("お届け先住所３"), selR).Value
            TextBox5.Text = Replace(dgvWrite.Item(DGVheaderCheck.IndexOf("お届け先名称１"), selR).Value, " 様", "")
            TextBox5.Text = Replace(TextBox5.Text, " 御中", "")
            TextBox5.Text = Replace(TextBox5.Text, " 宛", "")
            TextBox22.Text = dgvWrite.Item(DGVheaderCheck.IndexOf("お届け先名称２"), selR).Value
            For i As Integer = 0 To ComboBox6.Items.Count - 1
                If InStr(dgvWrite.Item(DGVheaderCheck.IndexOf("ご依頼主名称１"), selR).Value, ComboBox6.Items(i)) > 0 Then
                    ComboBox6.SelectedIndex = i
                    Exit For
                End If
            Next
            TextBox14.Text = dgvWrite.Item(DGVheaderCheck.IndexOf("品名１"), selR).Value
            TextBox15.Text = dgvWrite.Item(DGVheaderCheck.IndexOf("品名２"), selR).Value
            TextBox21.Text = dgvWrite.Item(DGVheaderCheck.IndexOf("品名２"), selR).Value
            TextBox16.Text = dgvWrite.Item(DGVheaderCheck.IndexOf("品名３"), selR).Value
            TextBox17.Text = dgvWrite.Item(DGVheaderCheck.IndexOf("品名４"), selR).Value
            TextBox18.Text = dgvWrite.Item(DGVheaderCheck.IndexOf("品名５"), selR).Value
            NumericUpDown1.Value = dgvWrite.Item(DGVheaderCheck.IndexOf("出荷個数"), selR).Value
            ComboBox13.SelectedItem = dgvWrite.Item(DGVheaderCheck.IndexOf("便種（スピードを選択）"), selR).Value
            If dgvWrite.Item(DGVheaderCheck.IndexOf("配達日"), selR).Value <> "" Then
                CheckBox1.Checked = True
                DateTimePicker1.Value = dgvWrite.Item(DGVheaderCheck.IndexOf("配達日"), selR).Value
            Else
                CheckBox1.Checked = False
            End If
            Select Case dgvWrite.Item(DGVheaderCheck.IndexOf("配達指定時間帯"), selR).Value
                Case "01"
                    ComboBox7.SelectedItem = "午前中"
                Case "12"
                    ComboBox7.SelectedItem = "12時-14時"
                Case "14"
                    ComboBox7.SelectedItem = "14時-16時"
                Case "16"
                    ComboBox7.SelectedItem = "16時-18時"
                Case "04"
                    ComboBox7.SelectedItem = "18時-21時"
                Case Else
                    ComboBox7.SelectedItem = "指定なし"
            End Select
            If dgvWrite.Item(DGVheaderCheck.IndexOf("代引金額"), selR).Value <> "" Then
                If dgvWrite.Item(DGVheaderCheck.IndexOf("代引金額"), selR).Value > 0 Then
                    CheckBox3.Checked = True
                    TextBox19.Text = dgvWrite.Item(DGVheaderCheck.IndexOf("代引金額"), selR).Value
                Else
                    CheckBox3.Checked = False
                    TextBox19.Text = 0
                End If
            Else
                CheckBox3.Checked = False
                TextBox19.Text = "0"
            End If
            If dgvWrite.Item(DGVheaderCheck.IndexOf("営業所止めフラグ"), selR).Value <> "" And dgvWrite.Item(DGVheaderCheck.IndexOf("営業所止めフラグ"), selR).Value <> "0" Then
                CheckBox2.Checked = True
                TextBox20.Text = dgvWrite.Item(DGVheaderCheck.IndexOf("営業店コード"), selR).Value
            Else
                CheckBox2.Checked = False
            End If
        ElseIf (InStr(headerStr, "品名記事名1記事名2注文日") > 0 Or InStr(headerStr, "コレクト代金引換額(税込)") > 0) And dgvWrite.RowCount > 1 Then
            If InStr(headerStr, "コレクト代金引換額(税込)") = 0 Then
                ComboBox11.SelectedItem = "ゆうパック元払い"
            Else
                ComboBox11.SelectedItem = "ゆうパック代引き"
            End If

            TextBox2.Text = dgvWrite.Item(DGVheaderCheck.IndexOf("お届け先　電話番号"), selR).Value
            TextBox3.Text = dgvWrite.Item(DGVheaderCheck.IndexOf("お届け先　郵便番号"), selR).Value
            For i As Integer = 0 To ComboBox4.Items.Count - 1
                If InStr(dgvWrite.Item(DGVheaderCheck.IndexOf("お届け先　住所1"), selR).Value, ComboBox4.Items(i)) > 0 Then
                    ComboBox4.SelectedIndex = i
                    Exit For
                End If
            Next
            TextBox25.Text = dgvWrite.Item(DGVheaderCheck.IndexOf("お届け先　住所2"), selR).Value
            TextBox4.Text = Replace(dgvWrite.Item(DGVheaderCheck.IndexOf("お届け先　住所1"), selR).Value, ComboBox4.SelectedItem, "")
            TextBox4.Text = LTrim(TextBox4.Text)
            TextBox5.Text = Replace(dgvWrite.Item(DGVheaderCheck.IndexOf("お届け先　名称"), selR).Value, " 様", "")
            TextBox5.Text = Replace(TextBox5.Text, " 御中", "")
            TextBox5.Text = Replace(TextBox5.Text, " 宛", "")
            TextBox22.Text = dgvWrite.Item(DGVheaderCheck.IndexOf("お届け先　名称２"), selR).Value
            For i As Integer = 0 To ComboBox6.Items.Count - 1
                If InStr(dgvWrite.Item(DGVheaderCheck.IndexOf("ご依頼主　名称１"), selR).Value, ComboBox6.Items(i)) > 0 Then
                    ComboBox6.SelectedIndex = i
                    Exit For
                End If
            Next
            If dgvWrite.Item(DGVheaderCheck.IndexOf("配送希望日"), selR).Value <> "" Then
                CheckBox1.Checked = True
                Dim d As String = dgvWrite.Item(DGVheaderCheck.IndexOf("配送希望日"), selR).Value
                Dim ddate As String = d.Substring(0, 4) & "/" & d.Substring(4, 2) & "/" & d.Substring(6, 2)
                DateTimePicker1.Value = ddate
            Else
                CheckBox1.Checked = False
            End If
            Select Case dgvWrite.Item(DGVheaderCheck.IndexOf("配送希望時間帯"), selR).Value
                Case "51"
                    ComboBox7.SelectedItem = "午前中"
                Case "52"
                    ComboBox7.SelectedItem = "12時-14時"
                Case "53"
                    ComboBox7.SelectedItem = "14時-16時"
                Case "54"
                    ComboBox7.SelectedItem = "16時-18時"
                Case "55"
                    ComboBox7.SelectedItem = "18時-21時"
                Case Else
                    ComboBox7.SelectedItem = "指定なし"
            End Select
            TextBox14.Text = dgvWrite.Item(DGVheaderCheck.IndexOf("品名"), selR).Value
            TextBox15.Text = dgvWrite.Item(DGVheaderCheck.IndexOf("記事名1"), selR).Value
            NumericUpDown1.Value = dgvWrite.Item(DGVheaderCheck.IndexOf("複数個口数"), selR).Value
            TextBox21.Text = dgvWrite.Item(DGVheaderCheck.IndexOf("注文日"), selR).Value
            If InStr(headerStr, "コレクト代金引換額(税込)") > 0 Then
                CheckBox3.Checked = True
                TextBox19.Text = dgvWrite.Item(DGVheaderCheck.IndexOf("コレクト代金引換額(税込)"), selR).Value
            End If
            'ゆうパケット------------------------------------------------------------------------------------
        ElseIf InStr(headerStr, "お届け先　郵便番号") > 0 And dgvWrite.RowCount > 1 Then
            ComboBox11.SelectedItem = "ゆうパケット"

            TextBox2.Text = dgvWrite.Item(DGVheaderCheck.IndexOf("お届け先　電話番号"), selR).Value
            TextBox3.Text = dgvWrite.Item(DGVheaderCheck.IndexOf("お届け先　郵便番号"), selR).Value
            For i As Integer = 0 To ComboBox4.Items.Count - 1
                If InStr(dgvWrite.Item(DGVheaderCheck.IndexOf("お届け先　住所1"), selR).Value, ComboBox4.Items(i)) > 0 Then
                    ComboBox4.SelectedIndex = i
                    Exit For
                End If
            Next
            TextBox4.Text = Replace(dgvWrite.Item(DGVheaderCheck.IndexOf("お届け先　住所1"), selR).Value, ComboBox4.SelectedItem, "")
            TextBox4.Text = LTrim(TextBox4.Text)
            TextBox5.Text = Replace(dgvWrite.Item(DGVheaderCheck.IndexOf("お届け先　名称"), selR).Value, " 様", "")
            TextBox5.Text = Replace(TextBox5.Text, " 御中", "")
            TextBox5.Text = Replace(TextBox5.Text, " 宛", "")
            TextBox22.Text = dgvWrite.Item(DGVheaderCheck.IndexOf("お届け先　名称２"), selR).Value
            TextBox21.Text = dgvWrite.Item(DGVheaderCheck.IndexOf("お届け先　メールアドレス１"), selR).Value
            For i As Integer = 0 To ComboBox6.Items.Count - 1
                If InStr(dgvWrite.Item(DGVheaderCheck.IndexOf("ご依頼主　名称１"), selR).Value, ComboBox6.Items(i)) > 0 Then
                    ComboBox6.SelectedIndex = i
                    Exit For
                End If
            Next



        End If
        'Catch ex As Exception
        '    MsgBox("選択データを開くことができません", MsgBoxStyle.SystemModal)
        'End Try
    End Sub

    Public Shared dgvWrite As DataGridView
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If dgvWrite.RowCount = 0 Then
            MsgBox("更新データがありません", MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        Try
            Dim str As String = DenpyoWrite()
            Dim strArray As String() = Split(str, ",")
            Dim r As Integer = dgvWrite.SelectedCells(0).RowIndex
            For c As Integer = 0 To strArray.Length - 1
                dgvWrite.Item(c, r).Value = strArray(c)
            Next
            MsgBox("更新しました", MsgBoxStyle.SystemModal)
        Catch ex As Exception
            MsgBox("データ列が異なるため更新できません。配送会社の選択を確認してください。", MsgBoxStyle.SystemModal)
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If dgvWrite.RowCount = 0 Then
            Select Case ComboBox11.SelectedItem
                Case "佐川急便ビズロジ", "佐川急便e飛伝2"
                    Csv_change.ComboBox1.SelectedItem = "FKstyle配（佐川）変換"
                Case "ゆうパケット"
                    Csv_change.ComboBox1.SelectedItem = "FKstyle配（ゆパケ）変換"
                Case "ゆうパック元払い"
                    Csv_change.ComboBox1.SelectedItem = "FKstyle配（ゆパック元）変換"
                Case "ゆうパック代引き"
                    Csv_change.ComboBox1.SelectedItem = "FKstyle配（ゆパック代引）変換"
            End Select
            Csv_change.Button6.PerformClick()
        End If

        Try
            Dim str As String = DenpyoWrite()
            Dim strArray As String() = Split(str, ",")
            dgvWrite.Rows.Add(strArray)
            MsgBox("追加しました", MsgBoxStyle.SystemModal)
        Catch ex As Exception
            MsgBox("データ列が異なるため追加できません。配送会社の選択を確認してください。", MsgBoxStyle.SystemModal)
        End Try
    End Sub

    Private Function DenpyoWrite()
        Dim str As String = ""
        Select Case ComboBox11.SelectedItem
            Case "佐川急便ビズロジ"
                str = ""                             '住所録コード
                str &= "," & TextBox2.Text           'お届け先電話番号
                str &= "," & TextBox3.Text           'お届け先郵便番号
                str &= "," & ComboBox4.SelectedItem & " " & TextBox4.Text  'お届け先住所１
                str &= "," & ""                      'お届け先住所２
                str &= "," & ""                      'お届け先住所３
                If InStr(str, " 様") > 0 Or InStr(str, " 御中") > 0 Or InStr(str, " 宛") > 0 Then
                    str &= "," & TextBox5.Text       'お届け先名称１
                Else
                    str &= "," & TextBox5.Text & " " & ComboBox14.SelectedItem
                End If
                str &= "," & TextBox22.Text          'お届け先名称２
                str &= "," & TextBox21.Text          'お客様管理ナンバー
                str &= "," & TextBox13.Text          'お客様コード
                str &= "," & ""                      '部署・担当者
                str &= "," & ""                      '荷送人電話番号
                str &= "," & TextBox12.Text          'ご依頼主電話番号
                str &= "," & TextBox10.Text          'ご依頼主郵便番号
                str &= "," & TextBox11.Text          'ご依頼主住所１
                str &= "," & TextBox23.Text          'ご依頼主住所２
                Dim cArray As String() = Split(tempoList(ComboBox6.SelectedIndex), "=")
                str &= "," & cArray(2)               'ご依頼主名称１
                str &= "," & ""                      'ご依頼主名称２
                str &= "," & ""                      '荷姿コード
                str &= "," & TextBox14.Text          '品名１
                str &= "," & TextBox21.Text          '品名２
                str &= "," & TextBox16.Text          '品名３
                str &= "," & TextBox17.Text          '品名４
                str &= "," & TextBox18.Text          '品名５
                str &= "," & NumericUpDown1.Value    '出荷個数
                If ComboBox13.SelectedItem = "飛脚宅配便" Then  '便種(スピードを選択)
                    str &= "," & "000"
                ElseIf ComboBox13.SelectedItem = "飛脚スーパー便" Then
                    str &= "," & "001"
                ElseIf ComboBox13.SelectedItem = "飛脚即日配達便" Then
                    str &= "," & "002"
                ElseIf ComboBox13.SelectedItem = "飛脚航空便(翌日中配達)" Then
                    str &= "," & "003"
                ElseIf ComboBox13.SelectedItem = "飛脚航空便(翌日午前中配達)" Then
                    str &= "," & "004"
                ElseIf ComboBox13.SelectedItem = "飛脚ジャストタイム便" Then
                    str &= "," & "005"
                End If
                str &= "," & ""                      '便種(クール便指定)
                If CheckBox1.Checked = True Then
                    str &= "," & Format(DateTimePicker1.Value, "yyyyMMdd")   '配達日
                Else
                    str &= "," & ""
                End If
                If ComboBox7.SelectedItem = "指定なし" Then '配達指定時間帯
                    str &= "," & ""
                ElseIf ComboBox7.SelectedItem = "午前中" Then
                    str &= "," & "01"
                ElseIf ComboBox7.SelectedItem = "12時-14時" Then
                    str &= "," & "12"
                ElseIf ComboBox7.SelectedItem = "14時-16時" Then
                    str &= "," & "14"
                ElseIf ComboBox7.SelectedItem = "16時-18時" Then
                    str &= "," & "16"
                ElseIf ComboBox7.SelectedItem = "18時-21時" Then
                    str &= "," & "04"
                End If
                str &= "," & ""                       '配達指定時間(時分)
                If CheckBox3.Checked = True Then      '代引金額
                    str &= "," & TextBox19.Text
                Else
                    str &= "," & ""
                End If
                str &= "," & "0"                      '消費税
                str &= "," & ""                       '保険金額
                str &= "," & "011"                    '指定シール１
                str &= "," & ""                       '指定シール２
                str &= "," & ""                       '指定シール３
                If CheckBox2.Checked = True Then      '営業所止めフラグ
                    str &= "," & "1"
                Else
                    str &= "," & "0"
                End If
                str &= "," & TextBox20.Text           '営業店コード
                str &= "," & ""                       '元着区分
                str &= "," & TextBox21.Text           '注文日
            Case "ゆうパケット", "ゆうパック元払い", "ゆうパック代引き"
                '共通---------
                str = TextBox3.Text           'お届け先　郵便番号
                str &= "," & ComboBox4.SelectedItem & " " & TextBox4.Text  'お届け先　住所1
                str &= "," & TextBox25.Text          'お届け先　住所2
                str &= ","                           'お届け先　住所3
                If InStr(str, " 様") > 0 Or InStr(str, " 御中") > 0 Or InStr(str, " 宛") > 0 Then
                    str &= "," & TextBox5.Text       'お届け先　名称
                Else
                    str &= "," & TextBox5.Text '& " " & ComboBox14.SelectedItem 敬称はゆうプリで入れる
                End If
                str &= "," & TextBox22.Text          'お届け先　名称２
                str &= "," & TextBox2.Text           'お届け先　電話番号
                str &= "," & TextBox10.Text          'ご依頼主　郵便番号
                str &= "," & TextBox11.Text & TextBox23.Text          'ご依頼主　住所１
                Dim cArray As String() = Split(tempoList(ComboBox6.SelectedIndex), "=")
                str &= "," & cArray(2)               'ご依頼主　名称１
                str &= "," & TextBox12.Text          'ご依頼主　電話番号
                '共通---------

                If ComboBox11.SelectedItem = "ゆうパケット" Then
                    str &= "," & "3"                      '商品サイズ/厚さ区分
                    str &= "," & TextBox21.Text           'お届け先　メールアドレス１
                    str &= "," & TextBox21.Text           '注文日
                ElseIf ComboBox11.SelectedItem = "ゆうパック元払い" Then
                    str &= "," & NumericUpDown1.Value     '複数個口数
                    str &= "," & TextBox21.Text           'お届け先　メールアドレス１
                    If CheckBox1.Checked = True Then
                        str &= "," & Format(DateTimePicker1.Value, "yyyyMMdd")   '配達日
                    Else
                        str &= "," & ""
                    End If
                    If ComboBox7.SelectedItem = "指定なし" Then '配達指定時間帯
                        str &= "," & ""
                    ElseIf ComboBox7.SelectedItem = "午前中" Then
                        str &= "," & "51"
                    ElseIf ComboBox7.SelectedItem = "12時-14時" Then
                        str &= "," & "52"
                    ElseIf ComboBox7.SelectedItem = "14時-16時" Then
                        str &= "," & "53"
                    ElseIf ComboBox7.SelectedItem = "16時-18時" Then
                        str &= "," & "54"
                    ElseIf ComboBox7.SelectedItem = "18時-21時" Then
                        str &= "," & "55"
                    End If
                    str &= "," & TextBox14.Text           '品名
                    str &= "," & TextBox15.Text           '記事名1
                    str &= "," & TextBox16.Text           '記事名2
                    str &= "," & TextBox21.Text           '注文日
                ElseIf ComboBox11.SelectedItem = "ゆうパック代引き" Then
                    str &= "," & NumericUpDown1.Value    '複数個口数
                    If CheckBox3.Checked = True Then      '代引金額
                        str &= "," & TextBox19.Text
                    Else
                        str &= "," & ""
                    End If
                    str &= "," & TextBox21.Text           'お届け先　メールアドレス１
                    If CheckBox1.Checked = True Then
                        str &= "," & Format(DateTimePicker1.Value, "yyyyMMdd")   '配達日
                    Else
                        str &= "," & ""
                    End If
                    If ComboBox7.SelectedItem = "指定なし" Then '配達指定時間帯
                        str &= "," & ""
                    ElseIf ComboBox7.SelectedItem = "午前中" Then
                        str &= "," & "51"
                    ElseIf ComboBox7.SelectedItem = "12時-14時" Then
                        str &= "," & "52"
                    ElseIf ComboBox7.SelectedItem = "14時-16時" Then
                        str &= "," & "53"
                    ElseIf ComboBox7.SelectedItem = "16時-18時" Then
                        str &= "," & "54"
                    ElseIf ComboBox7.SelectedItem = "18時-21時" Then
                        str &= "," & "55"
                    End If
                    str &= "," & TextBox14.Text           '品名
                    str &= "," & TextBox15.Text           '記事名1
                    str &= "," & TextBox16.Text           '記事名2
                    str &= "," & TextBox21.Text           '注文日
                End If
        End Select

        Return str
    End Function

    '都道府県選択
    Private Sub ComboBox5_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox5.SelectedIndexChanged
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
        If ComboBox5.SelectedItem <> "" Then
            whereCheck = " ([県名] = '" & ComboBox5.SelectedItem & "')"
        End If

        SQLCm.CommandText = "SELECT * FROM [X-ken-all2] WHERE" & whereCheck
        Adapter.Fill(Table)

        ComboBox15.Items.Clear()
        For r As Integer = 0 To Table.Rows.Count - 1
            Dim str As String = Table.Rows(r)("市町村名").ToString
            If str <> "" Then
                If Not ComboBox15.Items.Contains(str) Then
                    ComboBox15.Items.Add(Table.Rows(r)("市町村名").ToString)
                End If
            End If
        Next

        ComboBox15.SelectedIndex = 0

        '--------------------------------------------------------------
        Table.Dispose()
        Adapter.Dispose()
        SQLCm.Dispose()
        Cn.Dispose()
        '--------------------------------------------------------------
        '=======================================================
    End Sub

    '市町村選択
    Private Sub ComboBox15_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox15.SelectedIndexChanged
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
        If ComboBox5.SelectedItem <> "" Then
            whereCheck = " ([県名] = '" & ComboBox5.SelectedItem & "')"
        End If
        If ComboBox15.SelectedItem <> "" Then
            whereCheck &= " AND ([市町村名] = '" & ComboBox15.SelectedItem & "')"
        End If

        SQLCm.CommandText = "SELECT * FROM [X-ken-all2] WHERE" & whereCheck
        Adapter.Fill(Table)

        ComboBox16.Items.Clear()
        For r As Integer = 0 To Table.Rows.Count - 1
            Dim str As String = Table.Rows(r)("町名").ToString
            If str <> "" Then
                If Not ComboBox16.Items.Contains(str) Then
                    ComboBox16.Items.Add(Table.Rows(r)("町名").ToString)
                End If
            End If
        Next

        ComboBox16.SelectedIndex = 0

        '--------------------------------------------------------------
        Table.Dispose()
        Adapter.Dispose()
        SQLCm.Dispose()
        Cn.Dispose()
        '--------------------------------------------------------------
        '=======================================================
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox6.SelectedIndexChanged
        Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\会社情報.txt"
        Dim flag As Integer = 0
        Using sr As New System.IO.StreamReader(fName, System.Text.Encoding.Default)
            Do While Not sr.EndOfStream
                Dim s As String = sr.ReadLine
                If InStr(s, "[") > 0 And InStr(s, "]") > 0 Then
                    If flag = 2 Then
                        flag = 3
                    Else
                        If InStr(s, ComboBox6.SelectedItem) > 0 Then
                            flag = 1
                        ElseIf flag = 1 And InStr(s, "[銀行口座]") > 0 Then
                            Exit Do
                        End If
                    End If
                Else
                    If InStr(s, "郵便=") > 0 Then
                        s = Replace(s, "郵便=〒", "")
                        TextBox10.Text = s
                    ElseIf InStr(s, "住所=") > 0 Then
                        s = Replace(s, "住所=", "")
                        Dim sArray As String() = Split(s, " ")
                        TextBox11.Text = sArray(0)
                        TextBox23.Text = sArray(1)
                    ElseIf InStr(s, "電話=") > 0 Then
                        s = Replace(s, "電話=受付電話 : ", "")
                        TextBox12.Text = s
                    End If
                End If
            Loop
        End Using
        For i As Integer = 0 To tempoList.Count - 1
            If InStr(tempoList(i), ComboBox6.SelectedItem) > 0 Then
                Dim cArray As String() = Split(tempoList(i), "=")
                TextBox13.Text = cArray(1)
                Exit For
            End If
        Next
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        YubinToAdress(0, TextBox3)
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        YubinToAdress(1, TextBox8)
    End Sub

    Public Shared CnAccdb As String = ""
    Private Sub YubinToAdress(ByVal mode As Integer, ByVal TB1 As TextBox)
        Dim str As String() = New String() {"-", " ", "　"}
        For i As Integer = 0 To str.Length - 1
            TB1.Text = Replace(TB1.Text, str(i), "")
        Next

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
        If TB1.Text <> "" Then
            whereCheck = " ([郵便番号] = '" & TB1.Text & "')"
        End If

        SQLCm.CommandText = "SELECT * FROM [X-ken-all2] WHERE" & whereCheck
        Adapter.Fill(Table)

        If Table.Rows.Count <> 0 Then
            If mode = 0 Then
                ComboBox4.SelectedItem = Table.Rows(0)("県名").ToString
                TextBox4.Text = Table.Rows(0)("市町村名").ToString & Table.Rows(0)("町名").ToString
            ElseIf mode = 1 Then
                ComboBox5.SelectedItem = Table.Rows(0)("県名").ToString
                ComboBox15.SelectedItem = Table.Rows(0)("市町村名").ToString
                ComboBox16.SelectedItem = Table.Rows(0)("町名").ToString
            End If
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

        Ritou(mode)
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
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
        If ComboBox5.SelectedItem <> "" Then
            whereCheck = " ([県名] = '" & ComboBox5.SelectedItem & "')"
        End If
        If ComboBox15.SelectedItem <> "" Then
            whereCheck &= " AND ([市町村名] = '" & ComboBox15.SelectedItem & "')"
        End If
        If ComboBox16.SelectedItem <> "" Then
            whereCheck &= " AND ([町名] = '" & ComboBox16.SelectedItem & "')"
        End If

        SQLCm.CommandText = "SELECT * FROM [X-ken-all2] WHERE" & whereCheck
        Adapter.Fill(Table)

        If Table.Rows.Count <> 0 Then
            TextBox8.Text = Table.Rows(0)("郵便番号").ToString
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

        Ritou(1)
    End Sub

    '離島アラート
    Private Sub Ritou(ByVal mode As Integer)
        Dim tFlag As Integer = 0
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
        If mode = 0 Then
            If TextBox3.Text <> "" Then
                whereCheck = " ([郵便番号] = '" & TextBox3.Text & "')"
            End If
        ElseIf mode = 1 Then
            If TextBox8.Text <> "" Then
                whereCheck = " ([郵便番号] = '" & TextBox8.Text & "')"
            End If
            If ComboBox5.SelectedItem = "沖縄県" Then
                whereCheck = " ([住所] = '" & ComboBox5.SelectedItem & "')"
            End If
        End If

        SQLCm.CommandText = "SELECT * FROM [離島一覧] WHERE" & whereCheck
        Adapter.Fill(Table)

        If Table.Rows.Count <> 0 Then
            TextBox24.Text = "離島"
            TextBox24.BackColor = Color.Yellow
            tFlag = 1
        Else
            tFlag = 0
        End If

        '--------------------------------------------------------------
        Table.Dispose()
        Adapter.Dispose()
        SQLCm.Dispose()
        Cn.Dispose()
        '--------------------------------------------------------------
        '=======================================================

        If tFlag = 0 Then
            TextBox24.Text = ""
            TextBox24.BackColor = Color.White
        End If
    End Sub

    Private Sub ComboBox11_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox11.SelectedIndexChanged
        Select Case ComboBox11.SelectedItem
            Case "佐川急便ビズロジ", "佐川急便e飛伝2"
                GroupBox5.Enabled = True        '指定・その他グループ
                GroupBox4.Enabled = True        '品名グループ
                NumericUpDown1.Enabled = True   '出荷個数
                ComboBox13.Enabled = True       '便種
                CheckBox3.Enabled = True        '代引き
                CheckBox2.Enabled = True        '営業所留め
                TextBox16.Enabled = True        '品名3
                TextBox17.Enabled = True        '品名4
                TextBox18.Enabled = True        '品名5
            Case "ゆうパケット"
                GroupBox5.Enabled = False
                GroupBox4.Enabled = False
                NumericUpDown1.Enabled = False
                ComboBox13.Enabled = False
                CheckBox3.Checked = False
                CheckBox3.Enabled = False
                CheckBox2.Checked = False
                CheckBox2.Enabled = False
                TextBox16.Enabled = False
                TextBox17.Enabled = False
                TextBox18.Enabled = False
            Case "ゆうパック元払い", "ゆうパック代引き"
                GroupBox5.Enabled = True
                GroupBox4.Enabled = True
                NumericUpDown1.Enabled = True
                ComboBox13.Enabled = False
                If ComboBox11.SelectedItem = "ゆうパック元払い" Then
                    CheckBox3.Checked = False
                    CheckBox3.Enabled = False
                Else
                    CheckBox3.Checked = True
                    CheckBox3.Enabled = True
                End If
                CheckBox2.Checked = False
                CheckBox2.Enabled = False
                TextBox16.Enabled = False
                TextBox17.Enabled = False
                TextBox18.Enabled = False

        End Select
    End Sub

End Class