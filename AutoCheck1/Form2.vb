Imports System.Runtime.InteropServices
Imports System.Security.Permissions
Imports System.IO
Imports System.Data.OleDb
Imports System.Net
Imports System.Text.RegularExpressions
Imports System.Text
Imports NPOI.SS.UserModel

Public Class Form2

    Dim appPath As String = System.Reflection.Assembly.GetExecutingAssembly().Location
    Dim souryo As New ArrayList
    Dim otodoke As ArrayList = New ArrayList

    Dim CnAccdb_mise As String = ""

    Private Sub Form2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ComboBox1.SelectedIndex = 0
        ComboBox3.SelectedIndex = 26
        ComboBox2.SelectedIndex = 0

        AddHandler My.Settings.SettingChanging, AddressOf Settings_SettingChanging

        CnAccdb_mise = Path.GetDirectoryName(appPath) & "\db\miseDB.accdb"
        Kaisya("錦商事株式会社 Yahoo FKstyle")

        Dim fName2 As String = Path.GetDirectoryName(Form1.appPath) & "\config\サービスレベル.csv"
        otodoke = CSV_READ(fName2)

        Timer1.Enabled = True

    End Sub

    Private Function CSV_READ(ByRef path As String)
        Dim csvText As String = System.IO.File.ReadAllText(path, System.Text.Encoding.Default)

        '前後の改行を削除しておく
        csvText = csvText.Trim(
            New Char() {ControlChars.Cr, ControlChars.Lf})

        Dim csvRecords As New System.Collections.ArrayList
        Dim csvFields As New System.Collections.ArrayList

        Dim csvTextLength As Integer = csvText.Length
        Dim startPos As Integer = 0
        Dim endPos As Integer = 0
        Dim field As String = ""

        While True
            '空白を飛ばす
            While startPos < csvTextLength _
                AndAlso (csvText.Chars(startPos) = " "c OrElse csvText.Chars(startPos) = ControlChars.Tab)
                startPos += 1
            End While

            'データの最後の位置を取得
            If startPos < csvTextLength _
                AndAlso csvText.Chars(startPos) = ControlChars.Quote Then
                '"で囲まれているとき
                '最後の"を探す
                endPos = startPos
                While True
                    endPos = csvText.IndexOf(ControlChars.Quote, endPos + 1)
                    If endPos < 0 Then
                        Throw New ApplicationException("""が不正")
                    End If
                    '"が2つ続かない時は終了
                    If endPos + 1 = csvTextLength OrElse csvText.Chars((endPos + 1)) <> ControlChars.Quote Then
                        Exit While
                    End If
                    '"が2つ続く
                    endPos += 1
                End While

                '一つのフィールドを取り出す
                field = csvText.Substring(startPos, endPos - startPos + 1)
                '""を"にする
                field = field.Substring(1, field.Length - 2).Replace("""""", """")

                endPos += 1
                '空白を飛ばす
                While endPos < csvTextLength AndAlso
                    csvText.Chars(endPos) <> ","c AndAlso
                    csvText.Chars(endPos) <> ControlChars.Lf
                    endPos += 1
                End While
            Else
                '"で囲まれていない
                'カンマか改行の位置
                endPos = startPos
                While endPos < csvTextLength AndAlso
                    csvText.Chars(endPos) <> ","c AndAlso
                    csvText.Chars(endPos) <> ControlChars.Lf
                    endPos += 1
                End While

                '一つのフィールドを取り出す
                field = csvText.Substring(startPos, endPos - startPos)
                '後の空白を削除
                field = field.TrimEnd()
            End If

            'フィールドの追加
            csvFields.Add(field)

            '行の終了か調べる
            If endPos >= csvTextLength OrElse
                csvText.Chars(endPos) = ControlChars.Lf Then
                '行の終了
                'レコードの追加
                csvFields.TrimToSize()
                Dim str As String = ""
                For i As Integer = 0 To csvFields.Count - 1
                    If i = 0 Then
                        str = csvFields(i)
                    Else
                        str &= "|=|" & csvFields(i)
                    End If
                Next
                'csvRecords.Add(csvFields)
                csvRecords.Add(str)
                csvFields = New System.Collections.ArrayList(csvFields.Count)

                If endPos >= csvTextLength Then
                    '終了
                    Exit While
                End If
            End If

            '次のデータの開始位置
            startPos = endPos + 1
        End While

        csvRecords.TrimToSize()

        For r As Integer = 0 To csvRecords.Count - 1
            If InStr(csvRecords(r), "E+") Then
                MsgBox("エクセルで保存し直した形跡があります。データを作成し直してください", MsgBoxStyle.SystemModal)
                Exit For
            End If
        Next

        Return csvRecords
    End Function

    Private Sub Kaisya(ByVal kStr As String)
        TextBox22.Text = ""
        TextBox23.Text = ""

        Dim fName As String = Path.GetDirectoryName(appPath) & "\会社情報.txt"
        Dim flag As Integer = 0
        Using sr As New StreamReader(fName, Encoding.Default)
            Do While Not sr.EndOfStream
                Dim s As String = sr.ReadLine
                If InStr(s, "[") > 0 And InStr(s, "]") > 0 Then
                    If flag = 2 Then
                        flag = 3
                    Else
                        If InStr(s, "[" & kStr & "]") > 0 Then
                            flag = 1
                        ElseIf flag = 1 And InStr(s, "[銀行口座]") > 0 Then
                            flag = 2
                        End If
                    End If
                ElseIf InStr(s, "---") > 0 Then

                Else
                    If flag > 0 Then
                        s = Replace(s, "メール=", "")
                        s = Replace(s, "郵便=", "")
                        s = Replace(s, "住所=", "")
                        s = Replace(s, "電話=", "")
                        If flag = 1 Then
                            TextBox22.Text &= s & vbCrLf
                        ElseIf flag = 2 Then
                            TextBox23.Text &= s & vbCrLf
                        End If
                    End If
                End If
            Loop
        End Using
    End Sub

    'フォームのサイズ最小化バグ修正
    Private Sub Settings_SettingChanging(ByVal sender As Object, ByVal e As System.Configuration.SettingChangingEventArgs)
        If Not Me.WindowState = FormWindowState.Normal Then
            If e.SettingName = "Form2Location" Then
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub Form2_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If e.CloseReason = CloseReason.FormOwnerClosing Or e.CloseReason = CloseReason.UserClosing Then
            Me.Visible = False
            e.Cancel = True
        End If
    End Sub

    Private Sub TabPage1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles TabPage1.DragDrop
        Dim DR As DialogResult = MsgBox("読み込んでよろしいですか？", MsgBoxStyle.YesNo Or MsgBoxStyle.SystemModal)
        If DR = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If

        Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
        For Each filename As String In e.Data.GetData(DataFormats.FileDrop)
            Using sr As New System.IO.StreamReader(filename, System.Text.Encoding.UTF8)
                Dim num As Integer = 0
                DGV1.Rows.Clear()
                Dim DGrow As Integer = 0
                Do While Not sr.EndOfStream
                    Dim s As String = sr.ReadLine
                    'Dim sArray As String() = Split(s, ",")
                    If num = 1 Then
                        Dim sArray As String() = Split(s, " ")
                        TextBox2.Text = sArray(1)
                        Button2_Click(sender, e)
                    ElseIf num = 2 Then
                        s = Replace(s, ComboBox3.SelectedItem, "")
                        s = Replace(s, ComboBox4.SelectedItem, "")
                        s = Replace(s, ComboBox5.SelectedItem, "")
                        TextBox1.Text = s
                    ElseIf num = 3 Then
                        TextBox3.Text = Replace(s, "TEL：", "")
                    ElseIf num = 4 Then
                        TextBox8.Text = Replace(s, "名前：", "")
                    ElseIf num = 5 Then
                        s = Replace(s, "日時指定：", "")
                        s = Replace(s, "時間指定：", "")
                        Dim sArray As String() = Split(s, " ")
                        DateTimePicker1.Value = sArray(0)
                        ComboBox2.SelectedItem = sArray(1)
                    ElseIf num >= 6 Then
                        If InStr(s, ":") > 0 And InStr(s, "＝") > 0 Then
                            DGV1.Rows.Add(1)
                            Dim sArray1 As String() = Split(s, ":")
                            Dim sArray2 As String() = Split(sArray1(1), "×")
                            DGV1.Item(dH1.IndexOf("商品コード"), DGrow).Value = sArray2(0)
                            Dim sArray3 As String() = Split(sArray2(1), "＝")
                            DGV1.Item(dH1.IndexOf("個数"), DGrow).Value = sArray3(0)
                            Dim sArray4 As String() = Split(sArray3(1), " ")
                            DGV1.Item(dH1.IndexOf("金額"), DGrow).Value = sArray4(0)
                            'sArray5(1) = Replace(sArray5(1), "）", "")
                            Dim sArray5 As String() = Split(sArray4(1), "（")
                            Dim sArray6 As String() = Split(sArray5(1), "）")
                            DGV1.Item(dH1.IndexOf("名称"), DGrow).Value = sArray6(0)
                            DGV1.Item(dH1.IndexOf("オプション"), DGrow).Value = sArray6(1)
                            DGrow += 1
                        ElseIf InStr(s, "メモ：") > 0 Then
                            s = Replace(s, "メモ：", "")
                            TextBox13.Text = s
                        End If
                    End If
                    num += 1
                Loop
            End Using
        Next
    End Sub

    Private Sub TabPage1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles TabPage1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Dim DB_Table_name As String = ""
    Dim DB_Code_Header As String = ""
    Dim DB_Price_Header As String = ""
    Dim DB_Name_Header As String = ""
    Dim DB_Option_Header As String = ""

    'ストア名変更
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim fName As String = Path.GetDirectoryName(appPath) & "\config\送料.txt"
        Dim tFlag As Boolean = False
        souryo.Clear()
        Dim listArray As New ArrayList
        Using sr As New StreamReader(fName, System.Text.Encoding.Default)
            Do While Not sr.EndOfStream
                Dim s As String = sr.ReadLine
                If InStr(s, "[") > 0 And InStr(s, "]") > 0 Then
                    If InStr(s, ComboBox1.SelectedItem) > 0 Then
                        tFlag = True
                    Else
                        tFlag = False
                    End If
                ElseIf s = "" Then

                Else
                    If tFlag = True Then
                        Dim sArray As String() = Split(s, "=")
                        If sArray(0) = "メール便" Then
                            souryo.Add(sArray(1))
                        ElseIf sArray(0) = "送料" Then
                            TextBox5.Text = sArray(1)
                            souryo.Add(sArray(1))
                        ElseIf sArray(0) = "大型" Then
                            souryo.Add(sArray(1))
                        ElseIf sArray(0) = "離島・沖縄" Then
                            souryo.Add(sArray(1))
                        ElseIf sArray(0) = "北海道" Then
                            souryo.Add(sArray(1))
                        ElseIf sArray(0) = "代金引換" Then
                            TextBox6.Text = sArray(1)
                            souryo.Add(sArray(1))
                            tFlag = False
                        End If
                    End If
                End If
            Loop
        End Using

        Select Case ComboBox1.SelectedItem
            Case "FKstyle"
                DB_Table_name = "T_FKY_item"
            Case "通販の暁"
                DB_Table_name = "T_AKT_item"
            Case "通販の雑貨倉庫"
                DB_Table_name = "T_ZSQ_item"
            Case "福岡通販堂"
                DB_Table_name = "T_MASTER_master"
            Case "KuraNavi"
                DB_Table_name = "T_KNW_item"
            Case "問屋よかろうもん"
                DB_Table_name = "T_YKM_item"
            Case "雑貨の国のアリス"
                DB_Table_name = "T_ARS_item"
            Case "Lucky9"
                DB_Table_name = "T_LCY_item"
            Case "あかねY"
                DB_Table_name = "T_AKY_item"
            Case "あかね楽天"
                DB_Table_name = "T_AKR_item"
            Case "ヤフーKT"
                DB_Table_name = "T_KTY_item"
            Case "ヤフオクKT"
                DB_Table_name = "T_MASTER_master"
        End Select

        Select Case ComboBox1.SelectedItem
            Case "FKstyle", "Lucky9", "あかねY", "ヤフーKT"
                DB_Code_Header = "code"
                DB_Price_Header = "price"
                DB_Name_Header = "name"
                DB_Option_Header = "options"
            Case "通販の暁", "雑貨の国のアリス", "あかね楽天"
                DB_Code_Header = "商品管理番号（商品URL）"
                DB_Price_Header = "販売価格"
                DB_Name_Header = "商品名"
                DB_Option_Header = ""
            Case "通販の雑貨倉庫"
                DB_Code_Header = "Seller Code"
                DB_Price_Header = "Sell Price"
                DB_Name_Header = "Item Name"
                DB_Option_Header = "Option Info"
            Case "問屋よかろうもん"
                DB_Code_Header = "独自商品コード"
                DB_Price_Header = "販売価格"
                DB_Name_Header = "商品名"
                DB_Option_Header = "商品別特殊表示"
            Case "KuraNavi"
                DB_Code_Header = "itemCode"
                DB_Price_Header = "itemPrice"
                DB_Name_Header = "itemName"
                DB_Option_Header = "itemOption1,itemOption2,itemOption3"
            Case "福岡通販堂", "ヤフオクKT"
                DB_Code_Header = "商品コード"
                DB_Price_Header = "売価"
                DB_Name_Header = "商品名"
                DB_Option_Header = "JANコード"
        End Select
    End Sub

    '合計計算
    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged
        Goukei()
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged
        Goukei()
    End Sub

    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox6.TextChanged
        Goukei()
    End Sub

    Dim souryoKingaku As Integer = 0
    Private Sub SouryoKeisan()
        souryoKingaku = 0

        If CheckBox4.Checked = True And CheckBox5.Checked = False Then
            souryoKingaku = souryo(2)
        ElseIf CheckBox5.Checked = True Then
            souryoKingaku = 0
        Else
            souryoKingaku = souryo(1)
        End If
        Label42.Text = "送料：" & souryoKingaku & "円"

        If CheckBox6.Checked = True Then
            souryoKingaku += souryo(4)
            Label42.Text &= vbCrLf & "地域・その他：" & souryo(4) & "円"
        ElseIf CheckBox7.Checked = True Or CheckBox3.Checked = True Then
            souryoKingaku += souryo(3)
            Label42.Text &= vbCrLf & "地域・その他：" & souryo(3) & "円"
        Else
            Label42.Text &= vbCrLf & "地域・その他：0円"
        End If

        TextBox5.Text = souryoKingaku
    End Sub

    '送料離島
    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        SouryoKeisan()
    End Sub

    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged
        SouryoKeisan()
    End Sub

    Private Sub CheckBox5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox5.CheckedChanged
        SouryoKeisan()
    End Sub

    Private Sub CheckBox6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox6.CheckedChanged
        If CheckBox6.Checked = True Then
            CheckBox7.Checked = False
        End If
        SouryoKeisan()
    End Sub

    Private Sub CheckBox7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox7.CheckedChanged
        If CheckBox7.Checked = True Then
            CheckBox6.Checked = False
        End If
        SouryoKeisan()
    End Sub

    Private Sub Goukei()
        Dim result As Integer
        If Integer.TryParse(TextBox4.Text, result) And Integer.TryParse(TextBox5.Text, result) And Integer.TryParse(TextBox6.Text, result) Then
            Dim numA As Integer = Integer.Parse(TextBox4.Text)
            Dim numB As Integer = Integer.Parse(TextBox5.Text)
            Dim numC As Integer = Integer.Parse(TextBox6.Text)
            TextBox7.Text = numA + numB + numC
        End If
    End Sub

    '***************************************************************************************************
    'データベース関連
    '***************************************************************************************************
    Public Shared CnAccdb As String = ""
    Dim countNum As New ArrayList
    Dim sichoYomi As New ArrayList
    Dim choYomi As New ArrayList

    '郵便番号データ処理
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim fName As String = Path.GetDirectoryName(appPath) & "\x-ken-all.csv"
        Dim sName As String = Path.GetDirectoryName(appPath) & "\x-ken-all2.csv"
        Dim kenName As String = Path.GetDirectoryName(appPath) & "\ken-index.csv"
        Dim postName As String = Path.GetDirectoryName(appPath) & "\post-index.csv"

        Dim linesNum As Integer = 0
        Dim kenList As New ArrayList

        Using sr As New System.IO.StreamReader(fName, System.Text.Encoding.Default)
            Using sw As New System.IO.StreamWriter(sName, False, System.Text.Encoding.Default)
                Do While Not sr.EndOfStream
                    Dim s As String = sr.ReadLine
                    If s <> "" Then
                        s = Replace(s, """", "")
                        Dim sArray As String() = Split(s, ",")  '未作成

                    Else
                        s = ""
                    End If

                    sw.WriteLine(s)
                Loop
            End Using
            linesNum += 1
        End Using

        'Using sw As New System.IO.StreamWriter(kenName, False, System.Text.Encoding.Default)
        '    For Each kstr As String In kenList
        '        sw.WriteLine(kstr)
        '    Next
        'End Using

        MsgBox("終了しました", MsgBoxStyle.SystemModal)
    End Sub

    '都道府県
    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        'Dim fName As String = Path.GetDirectoryName(appPath) & "\x-ken-all2.csv"

        'ComboBox4.Items.Clear()

        'Dim tFlag As Integer = 0
        'Using sr As New System.IO.StreamReader(fName, System.Text.Encoding.Default)
        '    Do While Not sr.EndOfStream
        '        Dim s As String = sr.ReadLine
        '        If s <> "" Then
        '            Dim sArray As String() = Split(s, ",")
        '            If ComboBox3.SelectedItem = sArray(6) Then
        '                If ComboBox4.Items.Contains(sArray(7)) = False Then
        '                    ComboBox4.Items.Add(sArray(7))
        '                End If
        '                tFlag = 1
        '            End If
        '            If tFlag = 1 And ComboBox3.SelectedItem <> sArray(6) Then
        '                ComboBox4.SelectedIndex = 0
        '                Exit Do
        '            End If
        '        End If
        '    Loop
        'End Using

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
        If ComboBox3.SelectedItem <> "" Then
            whereCheck = " ([県名] = '" & ComboBox3.SelectedItem & "')"
        End If

        SQLCm.CommandText = "SELECT * FROM [X-ken-all2] WHERE" & whereCheck
        Adapter.Fill(Table)

        ComboBox4.Items.Clear()
        sichoYomi.Clear()
        For r As Integer = 0 To Table.Rows.Count - 1
            Dim str As String = Table.Rows(r)("市町村名").ToString
            If str <> "" Then
                If Not ComboBox4.Items.Contains(str) Then
                    ComboBox4.Items.Add(Table.Rows(r)("市町村名").ToString)
                    sichoYomi.Add(Table.Rows(r)("市町村名読み").ToString)
                End If
            End If
        Next

        TextBox25.Text = sichoYomi.Item(0)
        ComboBox4.SelectedIndex = 0

        '--------------------------------------------------------------
        Table.Dispose()
        Adapter.Dispose()
        SQLCm.Dispose()
        Cn.Dispose()
        '--------------------------------------------------------------
        '=======================================================

        '北海道・沖縄
        If ComboBox3.SelectedItem = "北海道" Then
            CheckBox6.Checked = True
            CheckBox7.Checked = False
        ElseIf ComboBox3.SelectedItem = "沖縄県" Then
            CheckBox6.Checked = False
            CheckBox7.Checked = True
        Else
            CheckBox6.Checked = False
            CheckBox7.Checked = False
        End If

        TextBox24.Text = ComboBox3.Text & ComboBox4.Text & ComboBox5.Text & TextBox1.Text
    End Sub

    '市町村
    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        'Dim fName As String = Path.GetDirectoryName(appPath) & "\x-ken-all2.csv"

        'ComboBox5.Items.Clear()

        'Dim tFlag As Integer = 0
        'Using sr As New System.IO.StreamReader(fName, System.Text.Encoding.Default)
        '    Do While Not sr.EndOfStream
        '        Dim s As String = sr.ReadLine
        '        If s <> "" Then
        '            Dim sArray As String() = Split(s, ",")
        '            If ComboBox4.SelectedItem = sArray(7) Then
        '                If ComboBox5.Items.Contains(sArray(8)) = False Then
        '                    ComboBox5.Items.Add(sArray(8))
        '                End If
        '                tFlag = 1
        '            End If
        '            If tFlag = 1 And ComboBox4.SelectedItem <> sArray(7) Then
        '                ComboBox5.SelectedIndex = 0
        '                Exit Do
        '            End If
        '        End If
        '    Loop
        'End Using

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
        If ComboBox3.SelectedItem <> "" Then
            whereCheck = " ([県名] = '" & ComboBox3.SelectedItem & "')"
        End If
        If ComboBox4.SelectedItem <> "" Then
            whereCheck &= " AND ([市町村名] = '" & ComboBox4.SelectedItem & "')"
        End If

        SQLCm.CommandText = "SELECT * FROM [X-ken-all2] WHERE" & whereCheck
        Adapter.Fill(Table)

        ComboBox5.Items.Clear()
        choYomi.Clear()
        For r As Integer = 0 To Table.Rows.Count - 1
            Dim str As String = Table.Rows(r)("町名").ToString
            If str <> "" Then
                If Not ComboBox5.Items.Contains(str) Then
                    ComboBox5.Items.Add(Table.Rows(r)("町名").ToString)
                    choYomi.Add(Table.Rows(r)("町名読み").ToString)
                End If
            End If
        Next

        TextBox25.Text = sichoYomi.Item(ComboBox4.SelectedIndex)
        If choYomi.Count > 0 Then
            TextBox26.Text = choYomi.Item(0)
            ComboBox5.SelectedIndex = 0
        Else
            TextBox26.Text = ""
            ComboBox5.Text = ""
        End If

        '--------------------------------------------------------------
        Table.Dispose()
        Adapter.Dispose()
        SQLCm.Dispose()
        Cn.Dispose()
        '--------------------------------------------------------------
        '=======================================================

        TextBox24.Text = ComboBox3.Text & ComboBox4.Text & ComboBox5.Text & TextBox1.Text

        'お届け目安検索
        TextBox27.Text = ""
        For i As Integer = 0 To otodoke.Count - 1
            If InStr(otodoke(i), "|=||=|") = 0 Then
                Dim otoLine As String() = Split(otodoke(i), "|=|")
                If Regex.IsMatch(ComboBox3.SelectedItem, "^" & otoLine(0) & ".*") Then
                    If Regex.IsMatch(ComboBox4.SelectedItem, ".*" & otoLine(1) & ".*") Then
                        TextBox27.Text = otoLine(2)
                        Exit For
                    End If
                End If
            End If
        Next
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox5.SelectedIndexChanged
        TextBox26.Text = choYomi.Item(ComboBox5.SelectedIndex)
        TextBox24.Text = ComboBox3.Text & ComboBox4.Text & ComboBox5.Text & TextBox1.Text
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        TextBox24.Text = ComboBox3.Text & ComboBox4.Text & ComboBox5.Text & TextBox1.Text
    End Sub

    '郵便番号検索
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        'Dim fName As String = Path.GetDirectoryName(appPath) & "\x-ken-all2.csv"

        'Dim tFlag As Integer = 0
        'Using sr As New System.IO.StreamReader(fName, System.Text.Encoding.Default)
        '    Do While Not sr.EndOfStream
        '        Dim s As String = sr.ReadLine
        '        If s <> "" Then
        '            Dim sArray As String() = Split(s, ",")
        '            If sArray(6) = ComboBox3.SelectedItem And sArray(7) = ComboBox4.SelectedItem And sArray(8) = ComboBox5.SelectedItem Then
        '                TextBox2.Text = sArray(2)
        '                Exit Do
        '            End If
        '        End If
        '    Loop
        'End Using

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
        If ComboBox3.SelectedItem <> "" Then
            whereCheck = " ([県名] = '" & ComboBox3.SelectedItem & "')"
        End If
        If ComboBox4.SelectedItem <> "" Then
            whereCheck &= " AND ([市町村名] = '" & ComboBox4.SelectedItem & "')"
        End If
        If ComboBox5.SelectedItem <> "" Then
            whereCheck &= " AND ([町名] = '" & ComboBox5.SelectedItem & "')"
        End If

        SQLCm.CommandText = "SELECT * FROM [X-ken-all2] WHERE" & whereCheck
        Adapter.Fill(Table)

        If Table.Rows.Count <> 0 Then
            TextBox2.Text = Table.Rows(0)("郵便番号").ToString
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

        Ritou()

        TextBox24.Text = ComboBox3.Text & ComboBox4.Text & ComboBox5.Text & TextBox1.Text
    End Sub

    Private Sub TextBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            If TextBox2.Text <> "" Then
                Button2.PerformClick()
            End If
        End If
    End Sub

    '住所検索
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'Dim fName As String = Path.GetDirectoryName(appPath) & "\x-ken-all2.csv"

        'Dim str As String() = New String() {"-", " ", "　"}
        'For i As Integer = 0 To str.Length - 1
        '    TextBox2.Text = Replace(TextBox2.Text, str(i), "")
        'Next

        'Using sr As New System.IO.StreamReader(fName, System.Text.Encoding.Default)
        '    Do While Not sr.EndOfStream
        '        Dim s As String = sr.ReadLine
        '        If s <> "" Then
        '            Dim sArray As String() = Split(s, ",")
        '            If TextBox2.Text = sArray(2) Then
        '                ComboBox3.SelectedItem = sArray(6)
        '                ComboBox4.SelectedItem = sArray(7)
        '                ComboBox5.SelectedItem = sArray(8)
        '                Exit Do
        '            End If
        '        End If
        '    Loop
        'End Using

        'Dim str As String() = New String() {"-", " ", "　"}
        'For i As Integer = 0 To str.Length - 1
        '    TextBox2.Text = Replace(TextBox2.Text, str(i), "")
        'Next
        TextBox2.Text = Regex.Replace(TextBox2.Text, "-|\s|　", "")

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
        If TextBox2.Text <> "" Then
            whereCheck = " ([郵便番号] = '" & TextBox2.Text & "')"
        End If

        SQLCm.CommandText = "SELECT * FROM [X-ken-all2] WHERE" & whereCheck
        Adapter.Fill(Table)

        If Table.Rows.Count <> 0 Then
            ComboBox3.SelectedItem = Table.Rows(0)("県名").ToString
            ComboBox4.SelectedItem = Table.Rows(0)("市町村名").ToString
            ComboBox5.SelectedItem = Table.Rows(0)("町名").ToString
            TextBox25.Text = Table.Rows(0)("市町村名読み").ToString
            TextBox26.Text = Table.Rows(0)("町名読み").ToString
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

        Ritou()

        TextBox24.Text = ComboBox3.Text & ComboBox4.Text & ComboBox5.Text & TextBox1.Text
    End Sub

    '離島アラート
    Private Sub Ritou()
        If TextBox2.Text = "" Then
            Exit Sub
        End If

        Dim tFlag As Integer = 0
        'If TextBox2.Text >= 9000000 Then
        '    CheckBox1.Checked = True
        '    TextBox21.Text = "沖縄県全て離島扱い"
        '    tFlag = 1
        'Else
        '    Dim fName As String = Path.GetDirectoryName(appPath) & "\config\離島一覧.csv"

        '    Using sr As New System.IO.StreamReader(fName, System.Text.Encoding.Default)
        '        Do While Not sr.EndOfStream
        '            Dim s As String = sr.ReadLine
        '            If s <> "" Then
        '                Dim sArray As String() = Split(s, ",")
        '                If TextBox2.Text = sArray(1) Then
        '                    CheckBox1.Checked = True
        '                    TextBox21.Text = sArray(0)
        '                    TextBox21.BackColor = Color.Yellow
        '                    tFlag = 1
        '                    Exit Do
        '                End If
        '            End If
        '        Loop
        '    End Using
        'End If

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
        If TextBox2.Text <> "" Then
            whereCheck = " ([郵便番号] = '" & TextBox2.Text & "')"
        End If
        If ComboBox3.SelectedItem = "沖縄県" Then
            whereCheck = " ([住所] = '" & ComboBox3.SelectedItem & "')"
        End If

        SQLCm.CommandText = "SELECT * FROM [離島一覧] WHERE" & whereCheck
        Adapter.Fill(Table)

        If Table.Rows.Count <> 0 Then
            CheckBox1.Checked = True
            TextBox21.Text = Table.Rows(0)("住所").ToString
            TextBox21.BackColor = Color.Yellow
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
            CheckBox1.Checked = False
            TextBox21.Text = ""
            TextBox21.BackColor = Color.White
        End If
    End Sub

    '日本郵便営業所検索
    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Dim url As String = "http://www.post.japanpost.jp/shiten_search/index.html"
        Form1.ToolStripTextBox1.Text = url
        Form1.TabBrowser1.SelectedTab.WebBrowser.Navigate(New Uri(Form1.ToolStripTextBox1.Text))
    End Sub

    '佐川営業所検索
    Private Sub LinkLabel3_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        Dim url As String = "http://www.sagawa-exp.co.jp/send/branch_search/tanto/"
        Form1.ToolStripTextBox1.Text = url
        Form1.TabBrowser1.SelectedTab.WebBrowser.Navigate(New Uri(Form1.ToolStripTextBox1.Text))
    End Sub

    'ヤマト営業所検索
    Private Sub LinkLabel4_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel4.LinkClicked
        Dim url As String = "http://www.e-map.ne.jp/smt/yamato01/"
        Form1.ToolStripTextBox1.Text = url
        Form1.TabBrowser1.SelectedTab.WebBrowser.Navigate(New Uri(Form1.ToolStripTextBox1.Text))
    End Sub


    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim all As HtmlElementCollection = Form1.TabBrowser1.SelectedTab.WebBrowser.Document.All
        'Dim forms As HtmlElementCollection = all.GetElementsByName("formSearchWord")
        Dim forms As HtmlElementCollection = all.GetElementsByName("")
        forms(1).InnerText = "東京"
    End Sub

    '価格取得
    Private Sub DataGridView1_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGV1.CellContentClick
        If DGV1.Columns(e.ColumnIndex).HeaderText = "取得" Then
            Panel3.Visible = True
            DGV1.Enabled = False

            Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
            Dim selnum As Integer = DGV1.SelectedCells(0).RowIndex
            Dim code As String = DGV1.Item(dH1.IndexOf("商品コード"), selnum).Value
            If DGV1.Item(dH1.IndexOf("代表コード"), selnum).Value <> "" Then
                code = DGV1.Item(dH1.IndexOf("代表コード"), selnum).Value
            End If

            Dim url As String = ""
            Dim pos As Integer = 0
            If ComboBox1.SelectedItem = "FKstyle" Then
                url = "https://store.shopping.yahoo.co.jp/fkstyle/" & code & ".html"
                pos = 4300
            ElseIf ComboBox1.SelectedItem = "通販の暁" Then
                url = "https://item.rakuten.co.jp/patri/" & code & "/"
                pos = 10000
            ElseIf ComboBox1.SelectedItem = "通販の雑貨倉庫" Then
                url = "http://www.qoo10.jp/shop/yasuyasu?keyword=" & code
            ElseIf ComboBox1.SelectedItem = "福岡通販堂" Then
                url = "http://www.qoo10.jp/shop/fukuoka?keyword=" & code
            ElseIf ComboBox1.SelectedItem = "KuraNavi" Then
                url = "http://www.dena-ec.com/bep/m/klist2?e_scope=O&user=38238702&keyword=" & code & "&clow=&chigh=&submit.x=40&submit.y=6"
            ElseIf ComboBox1.SelectedItem = "問屋よかろうもん" Then
                url = "https://www.yokaro.shop/shop/shopbrand.html?search=" & code
            ElseIf ComboBox1.SelectedItem = "雑貨の国のアリス" Then
                url = "https://item.rakuten.co.jp/alice-zk/" & code & "/"
            ElseIf ComboBox1.SelectedItem = "Lucky9" Then
                url = "https://store.shopping.yahoo.co.jp/lucky9/" & code & ".html"
            ElseIf ComboBox1.SelectedItem = "あかねY" Then
                url = "https://store.shopping.yahoo.co.jp/akaneashop/" & code & ".html"
            ElseIf ComboBox1.SelectedItem = "あかね楽天" Then
                url = "https://item.rakuten.co.jp/ashop/" & code & "/"
                pos = 10000
            ElseIf ComboBox1.SelectedItem = "ヤフーKT" Then
                url = "https://store.shopping.yahoo.co.jp/kt-zkshop/" & code & ".html"
            End If

            Form1.TabBrowser1.AddNewWebTabPage()
            Me.Focus()
            Application.DoEvents()

            Dim str As String() = Nothing

            Try
                '価格取得
                If CheckBox2.Checked = True Then
                    If ComboBox1.SelectedIndex <= 7 Then
                        Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS")
                        Select Case ComboBox1.SelectedItem
                            Case "通販の暁", "問屋よかろうもん", "雑貨の国のアリス"
                                enc = System.Text.Encoding.GetEncoding("EUC-JP")
                            Case "FKstyle", "通販の雑貨倉庫", "福岡通販堂", "Lucky9", "あかねY", "ヤフーKT"
                                enc = System.Text.Encoding.GetEncoding("utf-8")
                            Case "KuraNavi"
                                enc = System.Text.Encoding.GetEncoding("Shift_JIS")
                            Case Else
                                MsgBox("価格取得できないサイトです。直接店舗を開いて検索してください。")
                                Exit Sub
                                'enc = System.Text.Encoding.Default
                        End Select

                        Dim html As String = ""
                        Try
                            ServicePointManager.SecurityProtocol = DirectCast(3072, SecurityProtocolType)
                            Dim req As WebRequest = WebRequest.Create(url)
                            req.UseDefaultCredentials = False
                            Dim res As WebResponse = req.GetResponse()
                            Dim st As Stream = res.GetResponseStream()
                            Dim sr As StreamReader = New StreamReader(st, enc)
                            html = sr.ReadToEnd()
                            sr.Close()
                            st.Close()
                        Catch ex As Exception
                            MsgBox(ex.Message, MsgBoxStyle.SystemModal)
                        End Try

                        If DGV1.Item(dH1.IndexOf("個数"), selnum).Value Is Nothing Then
                            DGV1.Item(dH1.IndexOf("個数"), selnum).Value = 1
                        End If

                        Dim doc As HtmlAgilityPack.HtmlDocument = New HtmlAgilityPack.HtmlDocument
                        doc.LoadHtml(html)
                        Dim resStr As String = KakakuGet(doc, ComboBox1)
                        resStr = Replace(resStr, "円", "")
                        resStr = Replace(resStr, ",", "")
                        resStr = Replace(resStr, "(税込)", "")
                        DGV1.Item(dH1.IndexOf("金額"), selnum).Value = resStr

                        If DGV1.Item(dH1.IndexOf("名称"), selnum).Value = "" Then
                            resStr = NameGet(doc, ComboBox1)
                            'resStr = Replace(resStr, "</strong>", "")
                            'resStr = Replace(resStr, "<strong>", "")
                            If resStr.Length > 20 Then
                                resStr = resStr.Substring(0, 20)
                            End If
                            If resStr <> "" Then
                                DGV1.Item(dH1.IndexOf("名称"), selnum).Value = resStr
                            Else
                                DGV1.Item(dH1.IndexOf("名称"), selnum).Value = ""
                            End If
                        End If

                        str = New String() {resStr}   '強調する
                    End If
                End If
            Catch ex As Exception
                DGV1.Item(dH1.IndexOf("個数"), selnum).Value = 0
                DGV1.Item(dH1.IndexOf("金額"), selnum).Value = 0
                DGV1.Item(dH1.IndexOf("名称"), selnum).Value = ""
            End Try

            DGV1.Enabled = True
            Panel3.Visible = False

            Form1.ToolStripTextBox1.Text = url
            Dim selTab As AutoCheck1.WebTabPage = Form1.TabBrowser1.SelectedTab
            selTab.WebBrowser.Visible = False
            selTab.WebBrowser.Navigate(New Uri(Form1.ToolStripTextBox1.Text))
            Do
                Application.DoEvents()
            Loop Until selTab.WebBrowser.ReadyState = WebBrowserReadyState.Complete And selTab.WebBrowser.IsBusy = False
            If Not str Is Nothing Then
                Form1.Find(selTab.WebBrowser, str, 1)   '強調する
            End If
            selTab.WebBrowser.Visible = True

            If CheckBox2.Checked = True Then
                If ComboBox1.SelectedIndex = 4 Then
                    Dim res = KakakuGet2(selTab.WebBrowser, ComboBox1)
                    Dim str2 As String() = New String() {res}
                    Form1.Find(selTab.WebBrowser, str2, 1)
                    res = Replace(res, "円", "")
                    res = Replace(res, ",", "")
                    res = Replace(res, "(税込)", "")

                    If DGV1.Item(dH1.IndexOf("個数"), selnum).Value Is Nothing Then
                        DGV1.Item(dH1.IndexOf("個数"), selnum).Value = 1
                    End If
                    DGV1.Item(dH1.IndexOf("金額"), selnum).Value = res

                    If DGV1.Item(dH1.IndexOf("名称"), selnum).Value = "" Then
                        res = NameGet2(selTab.WebBrowser, ComboBox1)
                        If res <> "" Then
                            DGV1.Item(dH1.IndexOf("名称"), selnum).Value = res
                        Else
                            DGV1.Item(dH1.IndexOf("名称"), selnum).Value = ""
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Function KakakuGet(ByVal doc As HtmlAgilityPack.HtmlDocument, ByVal cb As ComboBox)
        Dim num As Integer = 0

        Dim selecter As String = ""
        Select Case ComboBox1.SelectedItem
            Case "FKstyle", "Lucky9", "あかねY"
                selecter = "//span[@class='elNum']"
                num = 0
            Case "通販の暁", "雑貨の国のアリス"
                selecter = "//span[@class='price2']"
                num = 0
            Case "通販の雑貨倉庫", "福岡通販堂"
                selecter = "//strong[@title='クーポン利用価格']"
                num = 0
            Case "KuraNavi"
                selecter = "//span[@class='price']"
                num = 0
            Case "問屋よかろうもん"
                selecter = "//span[@class='price']" '未作成
                num = 0
        End Select

        Dim nodes As HtmlAgilityPack.HtmlNodeCollection = doc.DocumentNode.SelectNodes(selecter)
        'For Each node As HtmlAgilityPack.HtmlNode In nodes
        '    MsgBox(node.InnerHtml, MsgBoxStyle.SystemModal)
        'Next

        Return nodes(num).InnerText
    End Function

    Private Function NameGet(ByVal doc As HtmlAgilityPack.HtmlDocument, ByVal cb As ComboBox)
        Dim num As Integer = 0

        Dim selecter As String = ""
        Select Case ComboBox1.SelectedItem
            Case "FKstyle", "Lucky9", "あかねY"
                selecter = "//p[@class='lead']"
                num = 0
            Case "通販の暁", "雑貨の国のアリス"
                selecter = "//span[@class='catch_copy']"
                num = 0
            Case "通販の雑貨倉庫", "福岡通販堂"
                selecter = "//a[@class='tt']"
                num = 0
            Case "KuraNavi"
                selecter = "//span[@class='name']"
                num = 0
            Case "問屋よかろうもん"
                selecter = "//span[@class='name']"  '未作成
                num = 0
        End Select

        Dim nodes As HtmlAgilityPack.HtmlNodeCollection = doc.DocumentNode.SelectNodes(selecter)
        'For Each node As HtmlAgilityPack.HtmlNode In nodes
        '    MsgBox(node.InnerHtml, MsgBoxStyle.SystemModal)
        'Next

        Return nodes(num).InnerText
    End Function

    Public Function KakakuGet2(ByVal wb As WebBrowser, ByVal cb As ComboBox)
        Do
            Application.DoEvents()
            'Loop Until Form1.TabBrowser1.SelectedTab.WebBrowser.ReadyState = WebBrowserReadyState.Complete And Form1.TabBrowser1.SelectedTab.WebBrowser.IsBusy = False
        Loop Until wb.ReadyState = WebBrowserReadyState.Complete And wb.IsBusy = False

        Dim result As New ArrayList
        Dim allHtml As HtmlDocument = wb.Document
        'Dim allHtml As HtmlDocument = WebBrowser1.Document

        Dim tag As String = ""
        Dim target As String = ""
        Dim targetStr As String = ""
        Dim num As Integer = 0
        Select Case ComboBox1.SelectedItem
            Case "FKstyle", "Lucky9", "あかねY"
                tag = "span"
                target = "className"
                targetStr = "elNum"
                num = 0
            Case "通販の暁", "雑貨の国のアリス"
                tag = "span"
                target = "className"
                targetStr = "price2"
                num = 0
            Case "通販の雑貨倉庫", "福岡通販堂"
                tag = "strong"
                target = "title"
                targetStr = "クーポン利用価格"
                num = 0
            Case "KuraNavi"
                tag = "span"
                target = "className"
                targetStr = "price"
                num = 0
            Case "問屋よかろうもん"
                tag = "span"
                target = "className"
                targetStr = "price"
                num = 0
        End Select

        result = HTMLgetTxt(allHtml, tag, target, targetStr)
        Dim res As String = ""
        If result.Count > 0 Then
            res = result(num)
        Else
            res = ""
        End If

        Return res
    End Function

    Public Function HTMLgetTxt(ByVal allHtml As HtmlDocument, ByVal tag As String, ByVal target As String, ByVal targetStr As String) As ArrayList
        Dim result As New ArrayList

        For Each el As HtmlElement In allHtml.GetElementsByTagName(tag)
            If el.GetAttribute(target) = targetStr Then
                result.Add(el.InnerText)
            End If
        Next

        Return result
    End Function

    Public Function NameGet2(ByVal wb As WebBrowser, ByVal cb As ComboBox)
        Do
            Application.DoEvents()
            'Loop Until Form1.TabBrowser1.SelectedTab.WebBrowser.ReadyState = WebBrowserReadyState.Complete And Form1.TabBrowser1.SelectedTab.WebBrowser.IsBusy = False
        Loop Until wb.ReadyState = WebBrowserReadyState.Complete And wb.IsBusy = False

        Dim result As New ArrayList
        Dim allHtml As HtmlDocument = wb.Document
        'Dim allHtml As HtmlDocument = WebBrowser1.Document

        Dim tag As String = ""
        Dim target As String = ""
        Dim targetStr As String = ""
        Dim num As Integer = 0
        Select Case ComboBox1.SelectedItem
            Case "FKstyle", "Lucky9", "あかねY"
                tag = "p"
                target = "className"
                targetStr = "lead"
                num = 0
            Case "通販の暁", "雑貨の国のアリス"
                tag = "span"
                target = "className"
                targetStr = "catch_copy"
                num = 0
            Case "通販の雑貨倉庫", "福岡通販堂"
                tag = "a"
                target = "className"
                targetStr = "tt"
                num = 0
            Case "KuraNavi"
                tag = "span"
                target = "className"
                targetStr = "name"
                num = 0
            Case "問屋よかろうもん"
                tag = "span"
                target = "className"
                targetStr = "name"
                num = 0
        End Select

        result = HTMLgetTxt(allHtml, tag, target, targetStr)
        Dim res As String = ""
        If result.Count > 0 Then
            res = result(num)
        Else
            res = ""
        End If
        If res.Length > 20 Then
            res = res.Substring(0, 20)
        End If

        Return res
    End Function

    Private Sub DataGridView1_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGV1.CellEndEdit
        If e.ColumnIndex = 0 Then
            Dim code As String = DGV1.Item(e.ColumnIndex, e.RowIndex).Value
            If code <> "" Then
                Dim table As DataTable = TM_DB_CONNECT_SELECT(CnAccdb_mise, DB_Table_name, "[" & DB_Code_Header & "]='" & code.ToLower & "'")

                Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
                For r As Integer = 0 To table.Rows.Count - 1
                    DGV1.Item(dH1.IndexOf("金額"), e.RowIndex).Value = table.Rows(0)(DB_Price_Header)
                    DGV1.Item(dH1.IndexOf("名称"), e.RowIndex).Value = table.Rows(0)(DB_Name_Header)
                    If DB_Option_Header <> "" Then
                        DGV1.Item(dH1.IndexOf("オプション"), e.RowIndex).Value = table.Rows(0)(DB_Option_Header)
                    End If
                    Exit For
                Next
            End If
        End If
    End Sub

    'DataGridView1セルの変更時
    Private Sub DataGridView1_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGV1.CellValueChanged
        If DGV1.RowCount > 1 Then
            Dim result As Integer = 0
            Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
            For r As Integer = 0 To DGV1.RowCount - 1
                'If Not DataGridView1.Item(1, r).Value Is DBNull.Value And Not DataGridView1.Item(2, r).Value Is DBNull.Value Then
                result = result + (DGV1.Item(dH1.IndexOf("金額"), r).Value * DGV1.Item(dH1.IndexOf("個数"), r).Value)
                'End If
            Next
            TextBox4.Text = result
        End If
    End Sub

    'DataGridView1への貼り付け
    Private _editingColumn As Integer
    Public ColumnChars() As String

    Private Sub DataGridView1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DGV1.KeyDown
        Dim dgv As DataGridView = CType(sender, DataGridView)
        Dim x As Integer = dgv.CurrentCellAddress.X
        Dim y As Integer = dgv.CurrentCellAddress.Y

        If e.KeyCode = Keys.Delete Or e.KeyCode = Keys.Back Then
            ' セルの内容を消去
            Dim delR As New ArrayList
            For r As Integer = 0 To dgv.SelectedCells.Count - 1
                delR.Add(dgv.SelectedCells(r).RowIndex)
            Next
            delR.Sort()
            delR.Reverse()
            For r As Integer = 0 To dgv.SelectedCells.Count - 1
                If delR(r) <> dgv.RowCount - 1 Then
                    dgv.Rows.RemoveAt(delR(r))
                End If
            Next
        ElseIf (e.Modifiers And Keys.Control) = Keys.Control And e.KeyCode = Keys.V Then
            ' クリップボードの内容を取得
            Dim clipText As String = Clipboard.GetText()
            ' 改行を変換
            clipText = clipText.Replace(vbCrLf, vbLf)
            clipText = clipText.Replace(vbCr, vbLf)
            ' 改行で分割
            Dim lines() As String = clipText.Split(vbLf)

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
                        DataGridView1_CellKeyPress(sender, tmpe)
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

    Private Sub DataGridView1_CellKeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs)
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

    '管理ページ表示
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim url As String = ""
        'If ComboBox1.SelectedIndex = 0 Then
        '    url = "https://pro.store.yahoo.co.jp/protop"
        'ElseIf ComboBox1.SelectedIndex = 1 Then
        '    url = "https://mainmenu.rms.rakuten.co.jp/rms/"
        'ElseIf ComboBox1.SelectedIndex = 2 Then
        '    url = "https://qsm.qoo10.jp/gmkt.inc.gsm.web/Login.aspx"
        'ElseIf ComboBox1.SelectedIndex = 3 Then
        '    url = "https://qsm.qoo10.jp/gmkt.inc.gsm.web/Login.aspx"
        'ElseIf ComboBox1.SelectedIndex = 4 Then
        '    MsgBox("DeNAは未対応", MsgBoxStyle.SystemModal)
        'End If
        url = "https://ne01.next-engine.com/Userjyuchu/jyuchuInp?kensaku_denpyo_no=" & TextBox9.Text

        Form1.ToolStripTextBox1.Text = url
        Form1.TabBrowser1.SelectedTab.WebBrowser.Navigate(New Uri(Form1.ToolStripTextBox1.Text))
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Dim url As String = ""
        url = "https://ne01.next-engine.com/Userjyuchu/jyuchuInp?kensaku_denpyo_no=" & TextBox19.Text
        Form1.ToolStripTextBox1.Text = url
        Form1.TabBrowser1.SelectedTab.WebBrowser.Navigate(New Uri(Form1.ToolStripTextBox1.Text))
    End Sub

    '返品交換用商品ページ表示
    Private Sub Button6_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim code As String = TextBox11.Text

        Dim url As String = ""
        If ComboBox1.SelectedItem = "FKstyle" Then
            url = "https://store.shopping.yahoo.co.jp/fkstyle/" & code & ".html"
        ElseIf ComboBox1.SelectedItem = "通販の暁" Then
            url = "https://item.rakuten.co.jp/patri/" & code & "/"
        ElseIf ComboBox1.SelectedItem = "通販の雑貨倉庫" Then
            url = "http://www.qoo10.jp/shop/yasuyasu?keyword=" & code
        ElseIf ComboBox1.SelectedItem = "福岡通販堂" Then
            url = "http://www.qoo10.jp/shop/fukuoka?keyword=" & code
        ElseIf ComboBox1.SelectedItem = "KuraNavi" Then
            url = "https://wowma.jp/bep/m/klist2?e_scope=O&user=38238702&keyword=" & code & "&clow=&chigh=&submit.x=40&submit.y=6"
        ElseIf ComboBox1.SelectedItem = "雑貨の国のアリス" Then
            url = "https://item.rakuten.co.jp/alice-zk/" & code & "/"
        ElseIf ComboBox1.SelectedItem = "Lucky9" Then
            url = "https://store.shopping.yahoo.co.jp/lucky9/" & code & ".html"
        ElseIf ComboBox1.SelectedItem = "あかねY" Then
            url = "https://store.shopping.yahoo.co.jp/akaneashop/" & code & ".html"
        ElseIf ComboBox1.SelectedItem = "問屋よかろうもん" Then
            url = "https://www.yokaro.shop/shop/shopbrand.html?search=" & code
        End If

        Form1.ToolStripTextBox1.Text = url
        Form1.TabBrowser1.SelectedTab.WebBrowser.Navigate(New Uri(Form1.ToolStripTextBox1.Text))
    End Sub

    '佐川お届け日数検索
    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        If CheckBox1.Checked = True Then
            Label34.Text = 4
        Else
            Dim fName As String = Path.GetDirectoryName(appPath) & "\config\お届け日数.txt"

            Dim tFlag As Integer = 0
            Using sr As New System.IO.StreamReader(fName, System.Text.Encoding.Default)
                Do While Not sr.EndOfStream
                    Dim s As String = sr.ReadLine
                    If s <> "" Then
                        Dim sArray As String() = Split(s, "=")
                        If ComboBox3.SelectedItem = sArray(0) Then
                            Label34.Text = sArray(1)
                            Dim d As Date = DateAdd(DateInterval.Day, Integer.Parse(Label34.Text), DateTimePicker2.Value)
                            TextBox20.Text = Format(d, "yyyy年MM月dd日")
                            Exit Do
                        End If
                    End If
                Loop
            End Using
        End If

        MsgBox("検索しました", MsgBoxStyle.SystemModal)
    End Sub

    '発送日変更
    Private Sub DateTimePicker2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker2.ValueChanged
        If DateTimePicker2.Value.DayOfWeek = DayOfWeek.Saturday Then
            DateTimePicker2.Value = DateAdd(DateInterval.Day, 2, DateTimePicker2.Value)
        ElseIf DateTimePicker2.Value.DayOfWeek = DayOfWeek.Sunday Then
            DateTimePicker2.Value = DateAdd(DateInterval.Day, 1, DateTimePicker2.Value)
        End If

        Dim fName As String = Path.GetDirectoryName(appPath) & "\Holiday.dat"
        Dim tFlag As Integer = 0
        Using sr As New System.IO.StreamReader(fName, System.Text.Encoding.Default)
            Do While Not sr.EndOfStream
                Dim s As String = sr.ReadLine
                If s <> "" Then
                    If InStr(s, "=") > 0 Then
                        Dim sArray As String() = Split(s, " = ")
                        If sArray(0) = Format(DateTimePicker2.Value, "yyyyMMdd") Then
                            If DateTimePicker2.Value.DayOfWeek = DayOfWeek.Saturday Then
                                DateTimePicker2.Value = DateAdd(DateInterval.Day, 2, DateTimePicker2.Value)
                            ElseIf DateTimePicker2.Value.DayOfWeek = DayOfWeek.Sunday Then
                                DateTimePicker2.Value = DateAdd(DateInterval.Day, 1, DateTimePicker2.Value)
                            Else
                                DateTimePicker2.Value = DateAdd(DateInterval.Day, 1, DateTimePicker2.Value)
                            End If
                            Exit Do
                        End If
                    End If
                End If
            Loop
        End Using

        Dim d As Date = DateAdd(DateInterval.Day, Integer.Parse(Label34.Text), DateTimePicker2.Value)
        TextBox20.Text = Format(d, "yyyy年MM月dd日")
    End Sub

    '保存
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim sfd As New SaveFileDialog()
        If ToolStripStatusLabel1.Text = "..." Then
            sfd.FileName = "電話対応" & Format(Now, "-yyyyMMdd-HHmm") & ".txt"
            If TextBox9.Text <> "" And TextBox12.Text <> "" Then
                sfd.FileName = TextBox9.Text & " " & TextBox12.Text & ".txt"
            ElseIf TextBox19.Text <> "" And TextBox18.Text <> "" Then
                sfd.FileName = TextBox19.Text & " " & TextBox18.Text & ".txt"
            End If
            ToolStripStatusLabel1.Text = sfd.FileName
        Else
            sfd.FileName = ToolStripStatusLabel1.Text
        End If

        sfd.AutoUpgradeEnabled = False
        sfd.Title = "保存先のファイルを選択してください"
        sfd.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        sfd.Filter = "TXTファイル(*.txt)|*.*"
        sfd.RestoreDirectory = True
        sfd.OverwritePrompt = True

        If sfd.ShowDialog(Me) = DialogResult.OK Then
            Dim stream As System.IO.Stream
            stream = sfd.OpenFile()
            Dim wStr = SaveFile()

            If Not (stream Is Nothing) Then
                'ファイルに書き込む
                Dim sw As New System.IO.StreamWriter(stream)
                sw.Write(wStr)
                '閉じる
                sw.Close()
                stream.Close()
            End If
        End If
    End Sub

    Private Function SaveFile()
        Dim wStr As String = ""

        If TabControl1.SelectedIndex = 0 Then
            Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
            wStr = Now & vbCrLf
            wStr &= "〒 " & TextBox2.Text & vbCrLf
            wStr &= ComboBox3.SelectedItem & ComboBox4.SelectedItem & ComboBox5.SelectedItem & TextBox1.Text & vbCrLf
            wStr &= "TEL：" & TextBox3.Text & vbCrLf
            wStr &= "名前：" & TextBox8.Text & vbCrLf
            Dim a As String = Format(Now, "yyyy年MM月dd日")
            Dim b As String = Format(DateTimePicker1.Value, "yyyy年MM月dd日")
            If Format(Now, "yyyy年MM月dd日") = Format(DateTimePicker1.Value, "yyyy年MM月dd日") Then
                wStr &= "日時指定：なし "
            Else
                wStr &= "日時指定：" & Format(DateTimePicker1.Value, "yyyy年MM月dd日") & " "
            End If
            If ComboBox2.SelectedIndex = 0 Then
                wStr &= "時間指定：なし" & vbCrLf
            Else
                wStr &= "時間指定：" & ComboBox2.SelectedItem & vbCrLf
            End If
            Dim naiyou As String = ""
            For r As Integer = 0 To DGV1.RowCount - 2
                naiyou &= r + 1 & ":" & DGV1.Item(dH1.IndexOf("商品コード"), r).Value & ""
                naiyou &= "×" & DGV1.Item(dH1.IndexOf("個数"), r).Value & ""
                naiyou &= "＝" & DGV1.Item(dH1.IndexOf("金額"), r).Value & " "
                naiyou &= "（" & DGV1.Item(dH1.IndexOf("名称"), r).Value & "）"
                naiyou &= "" & DGV1.Item(dH1.IndexOf("オプション"), r).Value
                naiyou &= vbCrLf
            Next
            wStr &= naiyou & vbCrLf
            wStr &= "金額：商品代金 " & TextBox4.Text & "円＋送料 " & TextBox5.Text & "円＋代引き " & TextBox6.Text & "円＝合計" & TextBox7.Text & "円" & vbCrLf
            wStr &= "メモ：" & TextBox13.Text & vbCrLf
        ElseIf TabControl1.SelectedIndex = 1 Then
            wStr = Now & vbCrLf
            wStr &= "注文番号：" & TextBox9.Text & vbCrLf
            wStr &= "名前：" & TextBox12.Text & vbCrLf
            wStr &= "〒 " & TextBox2.Text & vbCrLf
            wStr &= ComboBox3.SelectedItem & ComboBox4.SelectedItem & ComboBox5.SelectedItem & TextBox1.Text & vbCrLf
            wStr &= "電話：" & TextBox14.Text & vbCrLf
            wStr &= "メール：" & TextBox16.Text & vbCrLf
            wStr &= "商品番号：" & TextBox11.Text & vbCrLf
            wStr &= "クレーム内容：" & vbCrLf & TextBox10.Text & vbCrLf
        Else
            Return wStr
        End If

        Return wStr
    End Function

    'CSV（ネクストエンジン起票用）
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sfd As New SaveFileDialog()
        If ToolStripStatusLabel1.Text = "..." Then
            sfd.FileName = "電話対応" & Format(Now, "-yyyyMMdd-HHmm") & ".csv"
            If TextBox9.Text <> "" And TextBox12.Text <> "" Then
                sfd.FileName = TextBox9.Text & " " & TextBox12.Text & ".csv"
            ElseIf TextBox19.Text <> "" And TextBox18.Text <> "" Then
                sfd.FileName = TextBox19.Text & " " & TextBox18.Text & ".csv"
            End If
            ToolStripStatusLabel1.Text = sfd.FileName
        Else
            sfd.FileName = ToolStripStatusLabel1.Text
        End If

        sfd.AutoUpgradeEnabled = False
        sfd.Title = "保存先のファイルを選択してください"
        sfd.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        sfd.Filter = "CSVファイル(*.csv)|*.*"
        sfd.RestoreDirectory = True
        sfd.OverwritePrompt = True

        If sfd.ShowDialog(Me) = DialogResult.OK Then
            Dim str As String = "店舗伝票番号,受注日,受注郵便番号,受注住所１,受注住所２,受注名,受注名カナ,受注電話番号,受注メールアドレス,発送郵便番号,発送先住所１,発送先住所２,発送先名,発送先カナ,発送電話番号,支払方法,発送方法,商品計,税金,発送料,手数料,ポイント,その他費用,合計金額,ギフトフラグ,時間帯指定,日付指定,作業者欄,備考,商品名,商品コード,商品価格,受注数量,商品オプション,出荷済フラグ"
            str &= vbCrLf

            Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
            For r As Integer = 0 To DGV1.RowCount - 1
                If DGV1.Rows(r).IsNewRow Then
                    Continue For
                End If

                str &= Format(Now, "yyyyMMddHHmmssfff") & ","   '店舗伝票番号
                str &= Format(Now, "yyyy/MM/dd HH:mm:ss") & "," '受注日
                str &= Replace(TextBox2.Text, "-", "") & ","    '受注郵便番号
                str &= TextBox24.Text & ","                     '受注住所１
                str &= "" & ","                                 '受注住所２
                str &= TextBox8.Text & ","                      '受注名
                str &= "" & ","                                 '受注名カナ
                str &= TextBox3.Text & ","                      '受注電話番号
                str &= TextBox29.Text & ","                     '受注メールアドレス
                str &= Replace(TextBox2.Text, "-", "") & ","    '発送郵便番号
                str &= TextBox24.Text & ","                     '発送先住所１
                str &= "" & ","                                 '発送先住所２
                str &= TextBox8.Text & ","                      '発送先名
                str &= "" & ","                                 '発送先カナ
                str &= TextBox3.Text & ","                      '発送電話番号
                If RadioButton10.Checked Then
                    str &= "代金引換" & ","                     '支払方法
                ElseIf RadioButton11.Checked Then
                    str &= "銀行振込" & ","
                Else
                    str &= "その他" & ","
                End If
                If RadioButton13.Checked Then
                    str &= "宅配便" & ","                       '発送方法
                ElseIf RadioButton14.Checked Then
                    str &= "メール便" & ","
                Else
                    str &= "定形外郵便" & ","
                End If
                str &= TextBox4.Text & ","                      '商品計
                str &= "" & ","                                 '税金
                str &= TextBox5.Text & ","                      '発送料
                str &= TextBox6.Text & ","                      '手数料
                str &= "" & ","                                 'ポイント
                str &= "" & ","                                 'その他費用
                str &= TextBox7.Text & ","                      '合計金額
                str &= "" & ","                                 'ギフトフラグ
                If ComboBox2.Text = "" Then
                    str &= "" & ","                             '時間帯指定
                Else
                    str &= "時間帯指定[" & ComboBox2.Text & "]" & ","
                End If
                If DateDiff(DateInterval.Day, DateTimePicker1.Value, Now) > 0 Then
                    str &= Format(DateTimePicker1.Value, "yyyy/MM/dd") & ","  '日付指定
                Else
                    str &= "" & ","
                End If
                str &= "" & ","                                 '作業者欄
                str &= TextBox13.Text & ","                     '備考

                str &= DGV1.Item(dH1.IndexOf("名称"), r).Value & ","
                str &= DGV1.Item(dH1.IndexOf("商品コード"), r).Value & ","
                str &= DGV1.Item(dH1.IndexOf("金額"), r).Value & ","
                str &= DGV1.Item(dH1.IndexOf("個数"), r).Value & ","
                str &= DGV1.Item(dH1.IndexOf("オプション"), r).Value & ","

                str &= ""                                       '出荷済フラグ
                str &= vbCrLf
            Next

            File.WriteAllText(sfd.FileName, str, Encoding.GetEncoding("SHIFT-JIS"))
            MsgBox("csvデータを出力しました")
        End If
    End Sub


    Private Sub フォームのリセットToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles フォームのリセットToolStripMenuItem.Click
        Dim frm As Form = New Form2
        Me.Dispose()
        frm.Show()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        ToolStripStatusLabel2.Text = Format(Now, "HH:mm:ss")

        If Now.Second Mod 30 = 0 Then
            Dim fFolder As String = Path.GetDirectoryName(Form1.appPath) & "\PhoneBackUp\"
            If Not Directory.Exists(fFolder) Then
                Directory.CreateDirectory(fFolder)
            End If
            Dim fPath As String = fFolder & "PhoneBackUp_" & Format(Now, "yyMMddHHmm") & ".txt"
            Dim wStr = SaveFile()
            Dim sw As StreamWriter = New StreamWriter(fPath, False, System.Text.Encoding.Default)
            sw.Write(wStr)
            sw.Close()
            'My.Computer.FileSystem.WriteAllText(fPath, "", False, Form1.enc)
        End If
    End Sub

    Private Sub 行を削除ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 行を削除ToolStripMenuItem1.Click
        Dim dgv As DataGridView = DGV1
        Dim selCell = dgv.SelectedCells
        ROWSCUT(dgv, selCell)
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

    '納品書
    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Dim desk As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        Dim wb As IWorkbook = WorkbookFactory.Create(Path.GetDirectoryName(Form1.appPath) & "\納品書.xlsx")
        Dim ws As ISheet = wb.GetSheetAt(wb.GetSheetIndex("0"))

        Dim fName As String = Path.GetDirectoryName(appPath) & "\会社情報.txt"
        Dim flag As Integer = 0
        Dim Kshopname As String = ""
        Dim Kyubin As String = ""
        Dim Kjusyo As String = ""
        Dim Kdenwa As String = ""
        Dim Kmail As String = ""
        Using sr As New StreamReader(fName, Encoding.Default)
            Do While Not sr.EndOfStream
                Dim s As String = sr.ReadLine
                If Regex.IsMatch(s, ComboBox1.SelectedItem) Then
                    flag = 1
                ElseIf flag = 1 Then
                    If Regex.IsMatch(s, "^ショップ名=") Then
                        Dim array As String() = Split(s, "=")
                        Kshopname = array(1)
                    ElseIf Regex.IsMatch(s, "^郵便=") Then
                        Dim array As String() = Split(s, "=")
                        Kyubin = array(1)
                    ElseIf Regex.IsMatch(s, "^メール=") Then
                        Dim array As String() = Split(s, "=")
                        Kmail = array(1)
                    ElseIf Regex.IsMatch(s, "^住所=") Then
                        Dim array As String() = Split(s, "=")
                        Kjusyo = array(1)
                    ElseIf Regex.IsMatch(s, "^電話=受付電話") Then
                        Dim array As String() = Split(s, "=受付")
                        Kdenwa = array(1)
                    ElseIf Regex.IsMatch(s, "^----") Then
                        Exit Do
                    End If
                End If
            Loop
        End Using

        For r As Integer = 0 To 50
            If r < 16 Then
                For c As Integer = 0 To 6
                    Dim str As String = ""
                    Select Case ws.GetRow(r).GetCell(c).StringCellValue
                        Case "[名称]"
                            str = TextBox8.Text & " 様"
                        Case "[住所]"
                            str = "〒" & TextBox2.Text & " " & TextBox24.Text
                        Case "[電話]"
                            str = "TEL:" & TextBox3.Text
                        Case "[ショップ名]"
                            str = Kshopname
                        Case "[会社郵便]"
                            str = Kyubin
                        Case "[会社住所]"
                            str = Kjusyo
                        Case "[会社電話]"
                            str = Kdenwa
                        Case "[会社メール]"
                            str = Kmail
                    End Select
                    If str <> "" Then
                        ws.GetRow(r).GetCell(c).SetCellValue(str)
                    End If
                Next
            ElseIf r > 38 Then
                For c As Integer = 0 To 6
                    Select Case ws.GetRow(r).GetCell(c).StringCellValue
                        Case "[送料]"
                            ws.GetRow(r).GetCell(c).SetCellValue(CInt(TextBox5.Text))
                        Case "[代引料]"
                            ws.GetRow(r).GetCell(c).SetCellValue(CInt(TextBox6.Text))
                    End Select
                Next
            Else
                Dim numR As Integer = r - 16
                Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
                If numR < DGV1.RowCount - 1 Then
                    For c As Integer = 0 To 6
                        Select Case ws.GetRow(r).GetCell(c).StringCellValue
                            Case "[a" & numR + 1 & "]"
                                If DGV1.Item(dH1.IndexOf("名称"), numR).Value = "" Then
                                    Dim sStr As String = DGV1.Item(dH1.IndexOf("商品コード"), numR).Value
                                    ws.GetRow(r).GetCell(c).SetCellValue(sStr)
                                Else
                                    Dim sStr As String = DGV1.Item(dH1.IndexOf("商品コード"), numR).Value & "(" & DGV1.Item(dH1.IndexOf("名称"), numR).Value & ")"
                                    ws.GetRow(r).GetCell(c).SetCellValue(sStr)
                                End If
                            Case "[b" & numR + 1 & "]"
                                ws.GetRow(r).GetCell(c).SetCellValue(CInt(DGV1.Item(dH1.IndexOf("個数"), numR).Value))
                            Case "[c" & numR + 1 & "]"
                                ws.GetRow(r).GetCell(c).SetCellValue(CInt(DGV1.Item(dH1.IndexOf("金額"), numR).Value))
                            Case "[d" & numR + 1 & "]"
                                Dim sStr As String = DGV1.Item(dH1.IndexOf("オプション"), numR).Value
                                ws.GetRow(r).GetCell(c).SetCellValue(sStr)
                        End Select
                    Next
                Else
                    For c As Integer = 0 To 6
                        If InStr(ws.GetRow(r).GetCell(c).StringCellValue, "[") > 0 Then
                            ws.GetRow(r).GetCell(c).SetCellValue("")
                        End If
                    Next
                End If
            End If
        Next

        Using wfs = File.Create(desk & "\納品書" & Format(Now, "yyMMdd") & ".xlsx")
            wb.Write(wfs)
        End Using
        ws = Nothing
        wb = Nothing
        MsgBox("納品書出力しました", MsgBoxStyle.SystemModal)
    End Sub

    Private Sub RadioButtonA_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.Click,
        RadioButton2.Click, RadioButton3.Click, RadioButton4.Click, RadioButton5.Click,
        RadioButton6.Click, RadioButton7.Click, RadioButton8.Click, RadioButton9.Click

        Dim rb As RadioButton = sender
        Kaisya(rb.Text)
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim fFolder As String = Path.GetDirectoryName(Form1.appPath) & "\PhoneBackUp\"
        If Not Directory.Exists(fFolder) Then
            Directory.CreateDirectory(fFolder)
        End If
        System.Diagnostics.Process.Start(fFolder)
    End Sub

    Private Sub DGV1_SelectionChanged(sender As Object, e As EventArgs) Handles DGV1.SelectionChanged
        If DGV1.RowCount > 0 And DGV1.SelectedCells.Count > 0 Then
            TextBox28.Text = DGV1.SelectedCells(0).Value

            Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
            Dim selRow As Integer = DGV1.SelectedCells(0).RowIndex
            If CStr(DGV1.Item(dH1.IndexOf("商品コード"), selRow).Value) <> "" Then
                Dim code As String = DGV1.Item(dH1.IndexOf("商品コード"), selRow).Value.ToString.ToLower
                If DGV1.Item(dH1.IndexOf("代表コード"), selRow).Value <> "" Then
                    code = DGV1.Item(dH1.IndexOf("代表コード"), selRow).Value.ToString.ToLower
                End If
                Dim dH2 As ArrayList = TM_HEADER_GET(DGV2)
                Dim header2 As String() = {"商品コード", "商品名", "商品区分", "JANコード", "商品分類タグ", "代表商品コード", "在庫数", "引当数", "フリー在庫", "予約在庫", "予約引当", "予約フリー"}
                If DGV1.Item(dH1.IndexOf("商品コード"), selRow).Value <> "" Then
                    DGV2.Rows.Clear()
                    Dim table As DataTable = TM_DB_CONNECT_SELECT(CnAccdb_mise, "T_MASTER_master", "[商品コード]like'%" & code & "%'")
                    For i As Integer = 0 To table.Rows.Count - 1
                        If Regex.IsMatch(table.Rows(0)("商品コード"), "^" & code) Then
                            DGV2.Rows.Add()
                            For k As Integer = 0 To header2.Length - 1
                                DGV2.Item(dH2.IndexOf(header2(k)), DGV2.RowCount - 1).Value = table.Rows(i)(header2(k))
                            Next
                        End If
                    Next
                End If
            End If
        End If
    End Sub

    Private Sub DGV2_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV2.CellContentDoubleClick
        If DGV1.RowCount > 0 And DGV1.SelectedCells.Count > 0 Then
            Dim selRow1 As Integer = DGV1.SelectedCells(0).RowIndex
            Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
            Dim selRow2 As Integer = DGV2.SelectedCells(0).RowIndex
            Dim dH2 As ArrayList = TM_HEADER_GET(DGV2)

            If DGV1.Item(dH1.IndexOf("代表コード"), selRow1).Value = "" Then
                DGV1.Item(dH1.IndexOf("代表コード"), selRow1).Value = DGV1.Item(dH1.IndexOf("商品コード"), selRow1).Value
            End If
            DGV1.Item(dH1.IndexOf("商品コード"), selRow1).Value = DGV2.Item(dH2.IndexOf("商品コード"), selRow2).Value
            Dim sName As String = DGV2.Item(dH2.IndexOf("商品名"), selRow2).Value
            Dim sNameArray As String() = Split(sName, "@")
            DGV1.Item(dH1.IndexOf("名称"), selRow1).Value = sNameArray(0)
            If sNameArray.Length >= 2 Then
                DGV1.Item(dH1.IndexOf("オプション"), selRow1).Value = sNameArray(1)
            ElseIf sNameArray.Length >= 3 Then
                DGV1.Item(dH1.IndexOf("オプション"), selRow1).Value = sNameArray(1) & sNameArray(2)
            Else
                DGV1.Item(dH1.IndexOf("オプション"), selRow1).Value = ""
            End If
        End If
    End Sub

End Class