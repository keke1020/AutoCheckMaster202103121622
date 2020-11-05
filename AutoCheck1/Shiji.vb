Imports System.Net
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Public Class Shiji
    'Private cr As ChromeDriver = Nothing

    Private sizeW As Integer = 500
    Private sizeH As Integer = 800

    '再起動用
    Dim ChromeStartFlag As Boolean = False
    Dim ChromeRunFlag As Boolean = False

    Dim base64 As New MyBase64str("UTF-8")
    Dim ENC_UTF8 As Encoding = Encoding.UTF8
    Dim ENC_SJ As Encoding = Encoding.GetEncoding("SHIFT-JIS")

    Dim wb_loading_ng As Boolean = True
    Private locationPath As String = ""

    Dim loginArray As New ArrayList

    Private Sub Shiji_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label6.Text = "00"
        Label8.Text = "00"
        Label5.Text = ""

        ComboBox1.Items.Clear()
        ComboBox2.Items.Clear()
        ComboBox3.Items.Clear()

        ComboBox1.Items.Add("あかね")
        'ComboBox1.Items.Add("アリス")
        ComboBox1.Items.Add("通販の暁")
        ComboBox1.SelectedIndex = -1

        ComboBox2.Items.Add("予約解除")
        ComboBox2.Items.Add("予約に変更")

        ComboBox4.SelectedIndex = 1

        CheckBox1.Checked = False
        CheckBox3.Checked = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If IsNetworkAvailable() = False Then
            Exit Sub
        End If

        If ComboBox1.SelectedIndex = -1 Then
            MsgBox("店舗を選択してください。", MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        ComboBox3.Items.Clear()

        Dim svFile As String = Form1.サーバーToolStripMenuItem.Text & "\update\loginlist4.dat"
        Dim llFile As String = Path.GetDirectoryName(Form1.appPath) & "\loginlist4.dat"
        If Not File.Exists(llFile) Then
            File.Copy(svFile, llFile)
        End If

        Dim rakuten_login As String = ""
        Dim shop_bid As Integer
        If File.Exists(llFile) Then
            loginArray.Clear()
            Dim llArray As String() = Split(base64.Decode(File.ReadAllText(llFile, ENC_UTF8)), vbCrLf)
            loginArray.AddRange(llArray)

            If ComboBox1.SelectedIndex = 0 Then
                shop_bid = 351386

                'あかね楽、あかね楽個人
                For i As Integer = 0 To loginArray.Count - 1
                    If loginArray(i) <> "" Then
                        If Regex.IsMatch(loginArray(i), "あかね楽") And Regex.IsMatch(loginArray(i), "https://glogin.rms.rakuten.co.jp") And Regex.IsMatch(loginArray(i), "login_id") Then
                            Dim lArray As String() = Split(loginArray(i), ",")
                            rakuten_login &= "," & lArray(2)
                            rakuten_login &= "," & lArray(3)
                        ElseIf Regex.IsMatch(loginArray(i), "あかね楽個人") And Regex.IsMatch(loginArray(i), "https://glogin.rms.rakuten.co.jp") And Regex.IsMatch(loginArray(i), "rlogin-username-2-ja") Then
                            Dim lArray As String() = Split(loginArray(i), ",")
                            rakuten_login &= "," & lArray(2)
                            rakuten_login &= "," & lArray(3)
                        End If
                    End If
                Next
                'ElseIf ComboBox1.SelectedIndex = 1 Then
                '    shop_bid = 313483

                '    For i As Integer = 0 To loginArray.Count - 1
                '        If loginArray(i) <> "" Then
                '            If Regex.IsMatch(loginArray(i), "アリス") And Regex.IsMatch(loginArray(i), "https://glogin.rms.rakuten.co.jp") And Regex.IsMatch(loginArray(i), "login_id") Then
                '                Dim lArray As String() = Split(loginArray(i), ",")
                '                rakuten_login &= "," & lArray(2)
                '                rakuten_login &= "," & lArray(3)
                '            ElseIf Regex.IsMatch(loginArray(i), "アリス個人") And Regex.IsMatch(loginArray(i), "https://glogin.rms.rakuten.co.jp") And Regex.IsMatch(loginArray(i), "rlogin-username-2-ja") Then
                '                Dim lArray As String() = Split(loginArray(i), ",")
                '                rakuten_login &= "," & lArray(2)
                '                rakuten_login &= "," & lArray(3)
                '            End If
                '        End If
                '    Next
            ElseIf ComboBox1.SelectedIndex = 1 Then
                shop_bid = 321434
                '暁個人、暁
                For i As Integer = 0 To loginArray.Count - 1
                    If loginArray(i) <> "" Then
                        If Regex.IsMatch(loginArray(i), "暁個人") And Regex.IsMatch(loginArray(i), "https://glogin.rms.rakuten.co.jp") Then
                            Dim lArray As String() = Split(loginArray(i), ",")
                            rakuten_login &= "," & lArray(2)
                            rakuten_login &= "," & lArray(3)
                        ElseIf Regex.IsMatch(loginArray(i), "暁") And Regex.IsMatch(loginArray(i), "https://glogin.rms.rakuten.co.jp") Then
                            Dim lArray As String() = Split(loginArray(i), ",")
                            rakuten_login &= "," & lArray(2)
                            rakuten_login &= "," & lArray(3)
                        End If
                    End If
                Next
            End If
        End If

        rakuten_login = rakuten_login.Substring(1, rakuten_login.Length - 1)

        If wb1_rtlogin(rakuten_login) Then
            WebBrowser1.Navigate(New Uri("https://master.rms.rakuten.co.jp/rms/mall/rsf/shop/vc?__event=RS00_001_002&shop_bid=" & shop_bid))
            WaitWebBrowser1Completed()

            Dim count As Integer = 1
            For Each Element As HtmlElement In WebBrowser1.Document.GetElementsByTagName("table")
                Dim str As String = Replace(Element.OuterHtml, " ", "")
                str = Replace(Element.OuterHtml, "　", "")

                If Regex.IsMatch(str, "納期管理番号") Then
                    Dim delivery As String() = Split(str, New Char() {"#ffffff"})

                    Dim delivery_count As Integer = 0
                    For i As Integer = 0 To delivery.Count - 1
                        If delivery(i).Contains("font size") Then
                            If delivery_count = 0 Then
                                delivery_count += 1
                                Continue For
                            End If

                            Dim delivery_ As String() = Split(delivery(i), New Char() {"<td>"})
                            'Console.WriteLine(delivery_(6))

                            ComboBox3.Items.Add(count & ":" & delivery_(6).Substring(15, delivery_(6).Length - 15))
                            count += 1
                        End If
                    Next
                    Exit For
                End If
            Next
            MsgBox("項目を取得しました。", MsgBoxStyle.SystemModal)
        End If
    End Sub

    '共通
    Private Function IsNetworkAvailable()
        Dim IsNetworkAvailable_Flag As Boolean = False
        If NetworkInformation.NetworkInterface.GetIsNetworkAvailable() Then
            Label9.Text = "接続"
            Label9.BackColor = Color.LightGreen
            IsNetworkAvailable_Flag = True
        Else
            Label9.Text = "未接続"
            Label9.BackColor = Color.Transparent
        End If
        Return IsNetworkAvailable_Flag
    End Function


    Private Sub WaitWebBrowser1Completed()
        Do While wb_loading_ng
            If wb_loading_ng = False Then
                Exit Do
            Else
                Application.DoEvents()
            End If
        Loop
        wb_loading_ng = True
    End Sub

    Private Sub WebBrowser1_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted
        wb_loading_ng = False
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ComboBox2.SelectedIndex = -1 Then
            MsgBox("指示を選択してください。", MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        If ComboBox2.SelectedIndex = 1 Then
            If ComboBox3.SelectedIndex = -1 Then
                MsgBox("項目を選択してください。", MsgBoxStyle.SystemModal)
                Exit Sub
            End If
        End If

        Button2.BackColor = Color.White

        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add(0, "コード")

        If ComboBox2.SelectedIndex = 0 Then
            DataGridView1.Columns.Add(1, "指示")
            DataGridView1.Columns.Add(2, "結果")
            DataGridView1.Columns(0).ReadOnly = True
            DataGridView1.Columns(1).ReadOnly = True
            DataGridView1.Columns(2).ReadOnly = True
            DataGridView1.Columns(0).Width = "115"
            DataGridView1.Columns(1).Width = "200"
            DataGridView1.Columns(2).Width = "115"

            TextBox1.Text = Regex.Replace(TextBox1.Text, "　| ", "")
            Dim listArray As String() = Split(TextBox1.Text, vbCrLf)
            Array.Sort(listArray)
            Dim list_str As String = ""
            Dim codeArray As New ArrayList
            For i As Integer = 0 To listArray.Count - 1
                If listArray(i) <> "" Then
                    list_str &= listArray(i) & vbCrLf
                End If
            Next

            '去重
            Dim lines As String() = Split(list_str, vbCrLf)
            Dim res As String = ""
            Dim add_count = 1
            For i As Integer = 0 To lines.Length - 1
                If lines(i) = "" Then
                    Continue For
                End If

                If Not codeArray.Contains(lines(i)) Then
                    codeArray.Add(lines(i))
                    DataGridView1.Rows.Add(1)
                    DataGridView1.Item(0, add_count - 1).Value = lines(i)
                    DataGridView1.Item(1, add_count - 1).Value = "予約解除"
                    add_count += 1
                End If
            Next
        ElseIf ComboBox2.SelectedIndex = 1 Then
            Dim cmbCol As New DataGridViewComboBoxColumn()
            cmbCol.HeaderText = "項目"

            For i As Integer = 0 To ComboBox3.Items.Count - 1
                cmbCol.Items.Add(ComboBox3.Items(i))
            Next

            DataGridView1.Columns.Insert(1, cmbCol)

            DataGridView1.Columns.Add(2, "結果")

            DataGridView1.Columns(0).ReadOnly = True
            DataGridView1.Columns(2).ReadOnly = True

            DataGridView1.Columns(0).Width = "110"
            DataGridView1.Columns(1).Width = "210"
            DataGridView1.Columns(2).Width = "110"

            TextBox1.Text = Regex.Replace(TextBox1.Text, "　| ", "")
            Dim listArray As String() = Split(TextBox1.Text, vbCrLf)
            Array.Sort(listArray)
            Dim list_str As String = ""
            Dim codeArray As New ArrayList
            For i As Integer = 0 To listArray.Count - 1
                If listArray(i) <> "" Then
                    list_str &= listArray(i) & vbCrLf
                End If
            Next

            '去重
            Dim lines As String() = Split(list_str, vbCrLf)
            Dim res As String = ""
            Dim add_count = 1
            For i As Integer = 0 To lines.Length - 1
                If lines(i) = "" Then
                    Continue For
                End If

                If Not codeArray.Contains(lines(i)) Then
                    codeArray.Add(lines(i))
                    DataGridView1.Rows.Add(1)
                    DataGridView1.Item(0, add_count - 1).Value = lines(i)
                    DataGridView1.Item(1, add_count - 1).Value = ComboBox3.SelectedItem
                    add_count += 1
                End If
            Next
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If IsNetworkAvailable() = False Then
            Exit Sub
        End If

        If ComboBox2.SelectedIndex = -1 Then
            MsgBox("指示を選択してください。", MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        If ComboBox2.SelectedIndex = 1 Then
            If ComboBox3.SelectedIndex = -1 Then
                MsgBox("項目を選択してください。", MsgBoxStyle.SystemModal)
                Exit Sub
            End If
        End If

        If DataGridView1.RowCount = 0 Then
            MsgBox("抽出してください。", MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        If ComboBox4.SelectedIndex = -1 Then
            MsgBox("実行タイプを選択してください。", MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        If Button2.BackColor = Color.Yellow Then
            MsgBox("もう一度抽出してください。", MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        If ComboBox4.SelectedIndex = 0 Then
            Label5.Text = "テスト"
        ElseIf ComboBox4.SelectedIndex = 1 Then
            Label5.Text = "実行"
        End If

        TextBox2.Text = TextBox2.Text & vbCrLf & "--------------------- " & (Format(DateTime.Now, "yyyy/MM/dd HH:mm:ss")).ToString & " start ---------------------"

        Label6.Text = "0"
        Label8.Text = "0"

        Label8.Text = DataGridView1.RowCount.ToString

        Dim rakuten_login As String = ""
        Dim shop_bid As Integer
        Dim shop As String
        If ComboBox1.SelectedIndex = 0 Then
            shop = "akane"
            'ElseIf ComboBox1.SelectedIndex = 1 Then
            '    shop = "alice"
        ElseIf ComboBox1.SelectedIndex = 1 Then
            shop = "akatsuki"
        End If

        Dim svFile As String = Form1.サーバーToolStripMenuItem.Text & "\update\loginlist4.dat"
        Dim llFile As String = Path.GetDirectoryName(Form1.appPath) & "\loginlist4.dat"

        If Not File.Exists(llFile) Then
            File.Copy(svFile, llFile)
        End If

        If File.Exists(llFile) Then
            loginArray.Clear()
            Dim llArray As String() = Split(base64.Decode(File.ReadAllText(llFile, ENC_UTF8)), vbCrLf)
            loginArray.AddRange(llArray)

            If ComboBox1.SelectedIndex = 0 Then
                shop_bid = 351386

                'あかね楽、あかね楽個人
                For i As Integer = 0 To loginArray.Count - 1
                    If loginArray(i) <> "" Then
                        If Regex.IsMatch(loginArray(i), "あかね楽") And Regex.IsMatch(loginArray(i), "https://glogin.rms.rakuten.co.jp") And Regex.IsMatch(loginArray(i), "login_id") Then
                            Dim lArray As String() = Split(loginArray(i), ",")
                            rakuten_login &= "," & lArray(2)
                            rakuten_login &= "," & lArray(3)
                        ElseIf Regex.IsMatch(loginArray(i), "あかね楽個人") And Regex.IsMatch(loginArray(i), "https://glogin.rms.rakuten.co.jp") And Regex.IsMatch(loginArray(i), "rlogin-username-2-ja") Then
                            Dim lArray As String() = Split(loginArray(i), ",")
                            rakuten_login &= "," & lArray(2)
                            rakuten_login &= "," & lArray(3)
                        End If
                    End If
                Next
                'ElseIf ComboBox1.SelectedIndex = 1 Then
                '    shop_bid = 313483

                '    For i As Integer = 0 To loginArray.Count - 1
                '        If loginArray(i) <> "" Then
                '            If Regex.IsMatch(loginArray(i), "アリス") And Regex.IsMatch(loginArray(i), "https://glogin.rms.rakuten.co.jp") And Regex.IsMatch(loginArray(i), "login_id") Then
                '                Dim lArray As String() = Split(loginArray(i), ",")
                '                rakuten_login &= "," & lArray(2)
                '                rakuten_login &= "," & lArray(3)
                '            ElseIf Regex.IsMatch(loginArray(i), "アリス個人") And Regex.IsMatch(loginArray(i), "https://glogin.rms.rakuten.co.jp") And Regex.IsMatch(loginArray(i), "rlogin-username-2-ja") Then
                '                Dim lArray As String() = Split(loginArray(i), ",")
                '                rakuten_login &= "," & lArray(2)
                '                rakuten_login &= "," & lArray(3)
                '            End If
                '        End If
                '    Next
            ElseIf ComboBox1.SelectedIndex = 1 Then
                shop_bid = 321434
                '暁個人、暁
                For i As Integer = 0 To loginArray.Count - 1
                    If loginArray(i) <> "" Then
                        If Regex.IsMatch(loginArray(i), "暁個人") And Regex.IsMatch(loginArray(i), "https://glogin.rms.rakuten.co.jp") Then
                            Dim lArray As String() = Split(loginArray(i), ",")
                            rakuten_login &= "," & lArray(2)
                            rakuten_login &= "," & lArray(3)
                        ElseIf Regex.IsMatch(loginArray(i), "暁") And Regex.IsMatch(loginArray(i), "https://glogin.rms.rakuten.co.jp") Then
                            Dim lArray As String() = Split(loginArray(i), ",")
                            rakuten_login &= "," & lArray(2)
                            rakuten_login &= "," & lArray(3)
                        End If
                    End If
                Next
            End If
        End If

        rakuten_login = rakuten_login.Substring(1, rakuten_login.Length - 1)

        If wb1_rtlogin(rakuten_login) = False Then
            Exit Sub
        End If
        Application.DoEvents()

        Dim dH6 As ArrayList = TM_HEADER_GET(DGV6)
        If ComboBox2.SelectedIndex = 0 Then
            Dim shilyori_c As Integer = 1
            For r As Integer = 0 To DataGridView1.RowCount - 1
                Label6.Text = shilyori_c
                Dim success As Boolean = True
                Dim delteyoyaku As Boolean = False
                Dim code_ As String = DataGridView1.Item(0, r).Value
                Dim code_m As String() = Split(code_, "-")
                Dim search_code As String = code_m(0)

                Dim isDaihyo As Boolean = False
                If code_m.Count = 2 Then
                    For r2 As Integer = 0 To DGV6.RowCount - 1
                        If code_ = DGV6.Item(dH6.IndexOf("代表商品コード"), r2).Value.ToString.ToLower Then
                            isDaihyo = True
                            Exit For
                        ElseIf code_ = DGV6.Item(dH6.IndexOf("商品コード"), r2).Value.ToString.ToLower And DGV6.Item(dH6.IndexOf("代表商品コード"), r2).Value.trim = "" Then
                            isDaihyo = True
                            Exit For
                        End If
                    Next
                End If

                If code_m.Count = 2 And isDaihyo Then
                    search_code = code_
                End If

                If code_m.Count = 3 Then
                    For r2 As Integer = 0 To DGV6.RowCount - 1
                        If code_m(0) & "-" & code_m(1) = DGV6.Item(dH6.IndexOf("代表商品コード"), r2).Value.ToString.ToLower Then
                            isDaihyo = True
                            Exit For
                        End If
                    Next
                End If

                If code_m.Count = 3 And isDaihyo Then
                    search_code = code_m(0) & "-" & code_m(1)
                End If

                WebBrowser1.Navigate(New Uri("https://item.rms.rakuten.co.jp/rms/mall/rsf/item/vc?__event=RI03_001_002&shop_bid=" & shop_bid & "&mng_number=" & search_code))
                WaitWebBrowser1Completed()

                If WebBrowser1.Document.Body.InnerText.Contains("以下のエラーがあります") Then
                    DataGridView1.Item(2, r).Value = "×"
                    DataGridView1.Item(1, r).Style.BackColor = Color.Pink
                    success = False
                    dgvResult(DataGridView1, r, False, True, "")
                    Continue For
                Else
                    Dim hasou As Integer = -1
                    '配送方法セット
                    If shop = "akane" Then
                        If InStr(findSelectById("delivery_set_id"), "宅配便") > 0 Then
                            hasou = 1
                        ElseIf InStr(findSelectById("delivery_set_id"), "メール便") > 0 Or InStr(findSelectById("delivery_set_id"), "定形外") > 0 Then
                            hasou = 4
                        End If
                    ElseIf shop = "alice" Then
                        hasou = 1
                    ElseIf shop = "akatsuki" Then
                        If InStr(findSelectById("delivery_set_id"), "宅配便[特定送料](佐川急便)") > 0 Then
                            hasou = 1
                        ElseIf InStr(findSelectById("delivery_set_id"), "追跡可能メール便[特定送料](日本郵便)") > 0 Or InStr(findSelectById("delivery_set_id"), "定形外郵便[特定送料](日本郵便)") > 0 Then
                            hasou = 4
                        End If
                    End If

                    If hasou = -1 Then
                        MsgBox("発送方法は不明です。", MsgBoxStyle.SystemModal)
                        Exit Sub
                    End If

                    '枝番がある
                    If WebBrowser1.Document.GetElementById("r31").GetAttribute("checked") = "True" Then
                        'r31 項目選択肢別
                        WebBrowser1.Navigate(New Uri("https://item.rms.rakuten.co.jp/rms/mall/rsf/item/vc?__event=RI03_001_005&shop_bid=" & shop_bid & "&mng_number=" & search_code & "&delvdate_register_type=0"))
                        WaitWebBrowser1Completed()

                        Dim buf As Byte() = New Byte(WebBrowser1.DocumentStream.Length) {}
                        WebBrowser1.DocumentStream.Read(buf, 0, CInt(WebBrowser1.DocumentStream.Length))
                        Dim ec As System.Text.Encoding = System.Text.Encoding.GetEncoding(WebBrowser1.Document.Encoding)
                        Dim htmls As String = ec.GetString(buf)

                        Dim doc As HtmlAgilityPack.HtmlDocument = New HtmlAgilityPack.HtmlDocument
                        doc.LoadHtml(htmls)

                        '例: sh004-xl
                        '① 行 Sサイズ/--s Mサイズ/--m
                        '列 無し

                        '例: ee270-bk-xl
                        '② 行 ブラック/--bk シルバー/--si
                        '列 S/--s L/--l 

                        '例: zk095-02
                        '③ 行 種類/-
                        '列 01ブルー/--01 02羽根/--02

                        '例: ad228-bl
                        '④ 行 選択/-
                        '列 ブルー/--bl

                        '例: ny264-200 - 50kn
                        Dim selecter As String = "//font"
                        Dim nodes As HtmlAgilityPack.HtmlNodeCollection = doc.DocumentNode.SelectNodes(selecter)

                        Dim node_x As Integer = -1
                        Dim node_y As Integer = -1
                        Dim node_x_c As Integer = 0 '総個数
                        Dim node_y_c As Integer = 0
                        Dim node_x_ｓ As Integer = 0 '同じ個数
                        Dim node_y_ｓ As Integer = 0
                        Dim eda_x As String = ""

                        If code_m.Count = 2 Or (code_m.Count = 3 And isDaihyo) Then
                            eda_x = "--" & code_m(1)

                            For Each node As HtmlAgilityPack.HtmlNode In nodes
                                Dim str As String = node.InnerText
                                str = Regex.Replace(str, "\r|\n", " ")
                                str = Replace(str, " ", "")
                                'Console.WriteLine(str)

                                If InStr(str, "同行一括") > 0 Or InStr(str, "/-") > 0 Then '列
                                    If InStr(str, "同行一括") > 0 Then
                                        str = Replace(str, "同行一括", "")
                                        node_y_c += 1

                                        If InStr(str, "/-") Then
                                            Dim str_sp As String() = Split(str, "/")

                                            If str_sp.Length = 3 Then
                                                If str_sp(2) = eda_x Then
                                                    node_y = node_y_c
                                                    node_y_ｓ += 1
                                                End If
                                            Else
                                                If str_sp(1) = eda_x Then
                                                    node_y = node_y_c
                                                    node_y_ｓ += 1
                                                End If
                                            End If
                                        End If
                                    ElseIf InStr(str, "同列一括") > 0 Or InStr(str, "/-") > 0 Then '行 
                                        str = Replace(str, "同列一括", "")
                                        node_x_c += 1

                                        If InStr(str, "/-") Then
                                            Dim str_sp As String() = Split(str, "/")
                                            If str_sp.Length = 3 Then
                                                If str_sp(2) = eda_x Then
                                                    node_x = node_x_c
                                                    node_x_ｓ += 1
                                                End If
                                            Else
                                                If str_sp(1) = eda_x Then
                                                    node_x = node_x_c
                                                    node_x_ｓ += 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next

                            If node_x = -1 And node_y = -1 Then '無し
                                dgvResult(DataGridView1, r, False, True, "")
                                Continue For
                            End If

                            If node_x_c = 0 Or node_y_c = 0 Then '無し
                                dgvResult(DataGridView1, r, False, True, "")
                                Continue For
                            End If

                            '匹配到行，列只有1个
                            If node_x <> -1 And node_y = 1 Then '無し
                                dgvResult(DataGridView1, r, False, True, "")
                                Continue For
                            End If

                            If node_y_ｓ > 1 Or node_y_ｓ > 1 Then '同じ枝個数が多い
                                dgvResult(DataGridView1, r, False, True, "")
                                Continue For
                            End If

                            If node_y_c = 1 And node_x_c = 1 Then  'サンプル③
                                node_y = 1
                            End If

                            'ad038
                            If node_x <> -1 And node_y = -1 Then '
                                node_y = 0
                            End If

                            If node_x <> -1 And node_y_c = 1 Then 'サンプル①
                                Dim result As String = selectShijiBySelect_index(hasou, "normal_delvdate_id_" & node_x & "_0")
                                Dim result_ As String() = Split(result, "||")
                                If result_(1) = "true" Then
                                    If ComboBox4.SelectedIndex = 1 Then
                                        Dim click_flag As Boolean = False
                                        Dim eleSels As HtmlElementCollection = WebBrowser1.Document.GetElementsByTagName("input")
                                        For Each eleSel As HtmlElement In eleSels
                                            If eleSel.GetAttribute("value").Trim = "納期情報および在庫数を変更する" Then
                                                eleSel.InvokeMember("Click")
                                                WaitWebBrowser1Completed()

                                                If WebBrowser1.Document.Body.InnerText.Contains("R-Storefront編集禁止") <> "False" Then
                                                    MsgBox("現在、商品一括編集処理中のため、R-Storefrontでの編集作業はできません。", MsgBoxStyle.SystemModal)
                                                    Exit Sub
                                                End If

                                                click_flag = True
                                                Exit For
                                            End If
                                        Next
                                        If click_flag Then
                                            dgvResult(DataGridView1, r, True, False, TextBox2.Text & vbCrLf & code_ & ":" & result_(0) & " --> 項目選択肢[" & hasou & "]")
                                        Else
                                            dgvResult(DataGridView1, r, False, False, "")
                                            success = False
                                            Continue For
                                        End If
                                    Else
                                        dgvResult(DataGridView1, r, True, False, TextBox2.Text & vbCrLf & code_ & ":" & result_(0) & " --> 項目選択肢[" & hasou & "]")
                                    End If
                                Else
                                    dgvResult(DataGridView1, r, False, False, "")
                                    success = False
                                    Continue For
                                End If
                            ElseIf node_x_c = 1 And node_y <> -1 Then 'サンプル③
                                Dim result As String = ""
                                If node_y = 1 And node_y_c = 1 Then
                                    result = selectShijiBySelect_index(hasou, "normal_delvdate_id_1_0")
                                Else
                                    result = selectShijiBySelect_index(hasou, "normal_delvdate_id_1_" & node_y)
                                End If
                                Dim result_ As String() = Split(result, "||")
                                If result_(1) = "true" Then
                                    If ComboBox4.SelectedIndex = 1 Then
                                        Dim click_flag As Boolean = False
                                        Dim eleSels As HtmlElementCollection = WebBrowser1.Document.GetElementsByTagName("input")
                                        For Each eleSel As HtmlElement In eleSels
                                            If eleSel.GetAttribute("value").Trim = "納期情報および在庫数を変更する" Then
                                                eleSel.InvokeMember("Click")
                                                WaitWebBrowser1Completed()

                                                If WebBrowser1.Document.Body.InnerText.Contains("R-Storefront編集禁止") <> "False" Then
                                                    MsgBox("現在、商品一括編集処理中のため、R-Storefrontでの編集作業はできません。", MsgBoxStyle.SystemModal)
                                                    Exit Sub
                                                End If

                                                click_flag = True
                                                Exit For
                                            End If
                                        Next
                                        If click_flag Then
                                            dgvResult(DataGridView1, r, True, False, TextBox2.Text & vbCrLf & code_ & ":" & result_(0) & " --> 項目選択肢[" & hasou & "]")
                                        Else
                                            dgvResult(DataGridView1, r, False, True, "")
                                            success = False
                                            Continue For
                                        End If
                                    Else
                                        dgvResult(DataGridView1, r, True, False, TextBox2.Text & vbCrLf & code_ & ":" & result_(0) & " --> 項目選択肢[" & hasou & "]")
                                    End If
                                Else
                                    dgvResult(DataGridView1, r, False, False, "")
                                    success = False
                                    Continue For
                                End If
                            End If
                        ElseIf code_m.Count = 3 And isDaihyo = False Then
                            eda_x = "--" & code_m(1)
                            Dim eda_y As String = code_m(2)

                            For Each node As HtmlAgilityPack.HtmlNode In nodes
                                Dim str As String = node.InnerText
                                str = Regex.Replace(str, "\r|\n", " ")
                                str = Replace(str, " ", "")
                                'Console.WriteLine(str)

                                'mb052-gou-a
                                If InStr(str, "同行一括") > 0 Or InStr(str, "/-") > 0 Then '列
                                    If InStr(str, "同行一括") > 0 Then
                                        str = Replace(str, "同行一括", "")
                                        node_y_c += 1

                                        If InStr(str, "/-") Then
                                            Dim str_sp As String() = Split(str, "/")
                                            If str_sp(1) = "--" & eda_y Or str_sp(1) = "-" & eda_y Then
                                                node_y = node_y_c
                                                node_y_ｓ += 1
                                            End If
                                        End If
                                    ElseIf InStr(str, "同列一括") > 0 Or InStr(str, "/-") > 0 Then '行 
                                        str = Replace(str, "同列一括", "")
                                        node_x_c += 1

                                        If InStr(str, "/-") Then
                                            Dim str_sp As String() = Split(str, "/")
                                            If str_sp.Length = 3 Then
                                                If str_sp(2) = eda_x Then
                                                    node_x = node_x_c
                                                    node_x_ｓ += 1
                                                End If
                                            ElseIf str_sp.Length = 2 Then
                                                If str_sp(1) = eda_x Then
                                                    node_x = node_x_c
                                                    node_x_ｓ += 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next

                            If node_x_c = 0 Or node_y_c = 0 Then '無し
                                dgvResult(DataGridView1, r, False, True, "")
                                success = False
                                Continue For
                            End If

                            If node_y_ｓ > 1 Or node_y_ｓ > 1 Then '同じ枝個数が多い
                                dgvResult(DataGridView1, r, False, True, "")
                                success = False
                                Continue For
                            End If

                            Dim result As String = selectShijiBySelect_index(hasou, "normal_delvdate_id_" & node_x & "_" & node_y)
                            Dim result_ As String() = Split(result, "||")
                            If result_(1) = "true" Then
                                If ComboBox4.SelectedIndex = 1 Then
                                    Dim click_flag As Boolean = False
                                    Dim eleSels As HtmlElementCollection = WebBrowser1.Document.GetElementsByTagName("input")
                                    For Each eleSel As HtmlElement In eleSels
                                        If eleSel.GetAttribute("value").Trim = "納期情報および在庫数を変更する" Then
                                            eleSel.InvokeMember("Click")
                                            WaitWebBrowser1Completed()

                                            If WebBrowser1.Document.Body.InnerText.Contains("R-Storefront編集禁止") <> "False" Then
                                                MsgBox("現在、商品一括編集処理中のため、R-Storefrontでの編集作業はできません。", MsgBoxStyle.SystemModal)
                                                Exit Sub
                                            End If

                                            click_flag = True
                                            Exit For
                                        End If
                                    Next

                                    If click_flag Then
                                        dgvResult(DataGridView1, r, True, False, TextBox2.Text & vbCrLf & code_ & ":" & result_(0) & " --> 項目選択肢[" & hasou & "]")
                                    Else
                                        dgvResult(DataGridView1, r, False, True, "")
                                        success = False
                                        Continue For
                                    End If
                                Else
                                    dgvResult(DataGridView1, r, True, False, TextBox2.Text & vbCrLf & code_ & ":" & result_(0) & " --> 項目選択肢[" & hasou & "]")
                                End If
                            Else
                                dgvResult(DataGridView1, r, False, False, "")
                                success = False
                                Continue For
                            End If
                        End If

                        If success Then
                            WebBrowser1.Navigate(New Uri("https://item.rms.rakuten.co.jp/rms/mall/rsf/item/vc?__event=RI03_003_003&shop_bid=" & shop_bid & "&mng_number=" & search_code))
                            WaitWebBrowser1Completed()
                            If WebBrowser1.Document.GetElementById("normal_delvdate_id").GetAttribute("selectedIndex") = 1 Or WebBrowser1.Document.GetElementById("normal_delvdate_id").GetAttribute("selectedIndex") = 4 Then
                                delteyoyaku = True
                            Else
                                delteyoyaku = False
                            End If

                            '商品名
                            If CheckBox1.Checked And delteyoyaku Then
                                WebBrowser1.Navigate(New Uri("https://item.rms.rakuten.co.jp/rms/mall/rsf/item/vc?__event=RI03_001_002&shop_bid=" & shop_bid & "&mng_number=" & search_code))
                                WaitWebBrowser1Completed()

                                If InStr(WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).GetAttribute("value"), "＜一部商品予約＞") > 0 Then
                                    WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).SetAttribute("value", WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).GetAttribute("value").Replace("＜一部商品予約＞", ""))
                                ElseIf InStr(WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).GetAttribute("value"), "＜一部予約＞") > 0 Then
                                    WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).SetAttribute("value", WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).GetAttribute("value").Replace("＜一部予約＞", ""))
                                ElseIf InStr(WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).GetAttribute("value"), "＜予約＞") > 0 Then
                                    WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).SetAttribute("value", WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).GetAttribute("value").Replace("＜予約＞", ""))
                                End If

                                If ComboBox4.SelectedIndex = 1 Then
                                    WebBrowser1.Document.GetElementById("submitButton").InvokeMember("Click")
                                    WaitWebBrowser1Completed()

                                    If WebBrowser1.Document.Body.InnerText.Contains("R-Storefront編集禁止") <> "False" Then
                                        dgvResult(DataGridView1, r, False, False, "")
                                        MsgBox("現在、商品一括編集処理中のため、R-Storefrontでの編集作業はできません。", MsgBoxStyle.SystemModal)
                                        Exit Sub
                                    End If
                                End If
                            End If

                            '項目選択肢を確認
                            If CheckBox3.Checked Then
                                WebBrowser1.Navigate(New Uri("https://item.rms.rakuten.co.jp/rms/mall/rsf/item/vc?__event=RI03_003_002&shop_bid=" & shop_bid & "&mng_number=" & search_code))
                                WaitWebBrowser1Completed()

                                If ComboBox4.SelectedIndex = 1 Then
                                    komoku_delete1(shop_bid, search_code)
                                End If
                            End If

                            dgvResult(DataGridView1, r, True, True, "")
                        End If
                    Else
                        '項目選択肢別を選択 normal_delvdate_id
                        If InStr(findSelectById("delivery_set_id"), "宅配便") > 0 Then
                            WebBrowser1.Document.GetElementById("normal_delvdate_id").GetElementsByTagName("option").Item(1).SetAttribute("selected", "selected")
                        ElseIf InStr(findSelectById("delivery_set_id"), "メール便") > 0 Or InStr(findSelectById("delivery_set_id"), "定形外") > 0 Then
                            WebBrowser1.Document.GetElementById("normal_delvdate_id").GetElementsByTagName("option").Item(4).SetAttribute("selected", "selected")
                        End If

                        ''商品名を確認
                        If CheckBox1.Checked Then
                            If InStr(WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).GetAttribute("value"), "＜一部商品予約＞") > 0 Then
                                WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).SetAttribute("value", WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).GetAttribute("value").Replace("＜一部商品予約＞", ""))
                            ElseIf InStr(WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).GetAttribute("value"), "＜一部予約＞") > 0 Then
                                WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).SetAttribute("value", WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).GetAttribute("value").Replace("＜一部予約＞", ""))
                            ElseIf InStr(WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).GetAttribute("value"), "＜予約＞") > 0 Then
                                WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).SetAttribute("value", WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).GetAttribute("value").Replace("＜予約＞", ""))
                            End If
                        End If

                        If ComboBox4.SelectedIndex = 1 Then
                            WebBrowser1.Document.GetElementById("submitButton").InvokeMember("Click")
                            WaitWebBrowser1Completed()

                            If WebBrowser1.Document.Body.InnerText.Contains("R-Storefront編集禁止") <> "False" Then
                                dgvResult(DataGridView1, r, False, False, "")
                                MsgBox("現在、商品一括編集処理中のため、R-Storefrontでの編集作業はできません。", MsgBoxStyle.SystemModal)
                                Exit Sub
                            End If
                        End If


                        '項目選択肢を確認
                        If CheckBox3.Checked And ComboBox4.SelectedIndex = 1 Then
                            komoku_delete1(shop_bid, search_code)
                        End If

                        dgvResult(DataGridView1, r, True, True, TextBox2.Text & vbCrLf & code_ & " --> 予約解除")
                    End If
                End If
                shilyori_c += 1
            Next
        ElseIf ComboBox2.SelectedIndex = 1 Then
            Dim shilyori_c As Integer = 1
            For r As Integer = 0 To DataGridView1.RowCount - 1
                Label6.Text = shilyori_c
                Dim success As Boolean = True
                Dim addyoyaku As Boolean = False
                Dim code_ As String = DataGridView1.Item(0, r).Value
                Dim code_m As String() = Split(code_, "-")
                Dim search_code As String = code_m(0)

                Dim shizistr As String = DataGridView1.Item(1, r).Value
                Dim shizistr_ As String() = Split(shizistr, ":")
                shizistr = shizistr_(1)

                Dim isDaihyo As Boolean = False
                If code_m.Count = 2 Then
                    For r2 As Integer = 0 To DGV6.RowCount - 1
                        If code_ = DGV6.Item(dH6.IndexOf("代表商品コード"), r2).Value.ToString.ToLower Then
                            isDaihyo = True
                            Exit For
                        ElseIf code_ = DGV6.Item(dH6.IndexOf("商品コード"), r2).Value.ToString.ToLower And DGV6.Item(dH6.IndexOf("代表商品コード"), r2).Value.trim = "" Then
                            isDaihyo = True
                            Exit For
                        End If
                    Next
                End If

                If code_m.Count = 2 And isDaihyo Then
                    search_code = code_
                End If

                If code_m.Count = 3 Then
                    For r2 As Integer = 0 To DGV6.RowCount - 1
                        If code_m(0) & "-" & code_m(1) = DGV6.Item(dH6.IndexOf("代表商品コード"), r2).Value.ToString.ToLower Then
                            isDaihyo = True
                            Exit For
                        End If
                    Next
                End If

                If code_m.Count = 3 And isDaihyo Then
                    search_code = code_m(0) & "-" & code_m(1)
                End If

                WebBrowser1.Navigate(New Uri("https://item.rms.rakuten.co.jp/rms/mall/rsf/item/vc?__event=RI03_001_002&shop_bid=" & shop_bid & "&mng_number=" & search_code))
                WaitWebBrowser1Completed()

                If WebBrowser1.Document.Body.InnerText.Contains("以下のエラーがあります") Then
                    DataGridView1.Item(2, r).Value = "×"
                    DataGridView1.Item(1, r).Style.BackColor = Color.Pink
                    success = False
                    dgvResult(DataGridView1, r, False, True, "")
                    Continue For
                Else
                    '枝番がある
                    If WebBrowser1.Document.GetElementById("r31").GetAttribute("checked") = "True" Then
                        WebBrowser1.Navigate(New Uri("https://item.rms.rakuten.co.jp/rms/mall/rsf/item/vc?__event=RI03_001_005&shop_bid=" & shop_bid & "&mng_number=" & search_code & "&delvdate_register_type=0"))
                        WaitWebBrowser1Completed()

                        Dim buf As Byte() = New Byte(WebBrowser1.DocumentStream.Length) {}
                        WebBrowser1.DocumentStream.Read(buf, 0, CInt(WebBrowser1.DocumentStream.Length))
                        Dim ec As System.Text.Encoding = System.Text.Encoding.GetEncoding(WebBrowser1.Document.Encoding)
                        Dim htmls As String = ec.GetString(buf)

                        Dim doc As HtmlAgilityPack.HtmlDocument = New HtmlAgilityPack.HtmlDocument
                        doc.LoadHtml(htmls)

                        '例: sh004-xl
                        '① 行 Sサイズ/--s Mサイズ/--m
                        '列 無し

                        '例: ee270-bk-xl
                        '② 行 ブラック/--bk シルバー/--si
                        '列 S/--s L/--l 

                        '例: zk095-02
                        '③ 行 種類/-
                        '列 01ブルー/--01 02羽根/--02

                        '例: ny264-200 - 50kn
                        Dim selecter As String = "//font"
                        Dim nodes As HtmlAgilityPack.HtmlNodeCollection = doc.DocumentNode.SelectNodes(selecter)

                        Dim node_x As Integer = -1
                        Dim node_y As Integer = -1
                        Dim node_x_c As Integer = 0 '総個数
                        Dim node_y_c As Integer = 0
                        Dim node_x_ｓ As Integer = 0 '同じ個数
                        Dim node_y_ｓ As Integer = 0
                        Dim eda_x As String = ""
                        Dim eda_name As String = ""

                        If code_m.Count = 2 Or (code_m.Count = 3 And isDaihyo) Then
                            eda_x = "--" & code_m(1)

                            For Each node As HtmlAgilityPack.HtmlNode In nodes
                                Dim str As String = node.InnerText
                                str = Regex.Replace(str, "\r|\n", " ")
                                str = Replace(str, " ", "")
                                'Console.WriteLine(str)

                                If InStr(str, "同行一括") > 0 Or InStr(str, "/-") > 0 Then '列
                                    If InStr(str, "同行一括") > 0 Then
                                        str = Replace(str, "同行一括", "")
                                        node_y_c += 1

                                        If InStr(str, "/--") Then
                                            Dim str_sp As String() = Split(str, "/")

                                            If str_sp.Length = 3 Then
                                                If str_sp(2) = eda_x Then
                                                    node_y = node_y_c
                                                    node_y_ｓ += 1
                                                    eda_name = str_sp(0)
                                                End If
                                            Else
                                                If str_sp(1) = eda_x Then
                                                    node_y = node_y_c
                                                    node_y_ｓ += 1
                                                    eda_name = str_sp(0)
                                                End If
                                            End If
                                        End If
                                    ElseIf InStr(str, "同列一括") > 0 Or InStr(str, "/-") > 0 Then '行 
                                        str = Replace(str, "同列一括", "")
                                        node_x_c += 1

                                        If InStr(str, "/--") Then
                                            Dim str_sp As String() = Split(str, "/")
                                            If str_sp.Length = 3 Then
                                                If str_sp(2) = eda_x Then
                                                    node_x = node_x_c
                                                    node_x_ｓ += 1
                                                    eda_name = str_sp(0)
                                                End If
                                            Else
                                                If str_sp(1) = eda_x Then
                                                    node_x = node_x_c
                                                    node_x_ｓ += 1
                                                    eda_name = str_sp(0)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next

                            If node_x = -1 And node_y = -1 Then '無し
                                dgvResult(DataGridView1, r, False, True, "")
                                Continue For
                            End If

                            If node_x_c = 0 Or node_y_c = 0 Then '無し
                                dgvResult(DataGridView1, r, False, True, "")
                                Continue For
                            End If

                            '匹配到行，列只有1个
                            If node_x <> -1 And node_y = 1 Then '無し
                                dgvResult(DataGridView1, r, False, True, "")
                                Continue For
                            End If

                            If node_y_ｓ > 1 Or node_y_ｓ > 1 Then '同じ枝個数が多い
                                dgvResult(DataGridView1, r, False, True, "")
                                Continue For
                            End If

                            If node_y_c = 1 And node_x_c = 1 Then  'サンプル③
                                node_y = 1
                            End If

                            If node_x <> -1 And node_y = -1 Then
                                node_y = 0
                            End If

                            If node_x <> -1 And node_y_c = 1 Then 'サンプル①
                                Dim result As String = ""
                                If shop = "alice" Then
                                    result = selectShijiBySelect_index(3, "normal_delvdate_id_" & node_x & "_0")
                                Else
                                    result = selectShijiBySelect(shizistr, "normal_delvdate_id_" & node_x & "_0")
                                End If

                                Dim result_ As String() = Split(result, "||")
                                If result_(1) = "true" Then
                                    If ComboBox4.SelectedIndex = 1 Then
                                        Dim click_flag As Boolean = False
                                        Dim eleSels As HtmlElementCollection = WebBrowser1.Document.GetElementsByTagName("input")
                                        For Each eleSel As HtmlElement In eleSels
                                            If eleSel.GetAttribute("value") = "　　納期情報および在庫数を変更する　" Then
                                                eleSel.InvokeMember("Click")
                                                WaitWebBrowser1Completed()

                                                If WebBrowser1.Document.Body.InnerText.Contains("R-Storefront編集禁止") <> "False" Then
                                                    MsgBox("現在、商品一括編集処理中のため、R-Storefrontでの編集作業はできません。", MsgBoxStyle.SystemModal)
                                                    Exit Sub
                                                End If

                                                click_flag = True
                                                Exit For
                                            End If
                                        Next
                                        If click_flag Then
                                            dgvResult(DataGridView1, r, True, True, TextBox2.Text & vbCrLf & code_ & ":" & result_(0) & " --> " & shizistr)
                                        Else
                                            dgvResult(DataGridView1, r, False, False, "")
                                            success = False
                                            Continue For
                                        End If
                                    Else
                                        dgvResult(DataGridView1, r, True, True, TextBox2.Text & vbCrLf & code_ & ":" & result_(0) & " --> " & shizistr)
                                    End If
                                Else
                                    dgvResult(DataGridView1, r, False, False, "")
                                    success = False
                                    Continue For
                                End If
                            ElseIf node_x_c = 1 And node_y <> -1 Then 'サンプル③
                                Dim result As String = ""
                                'If node_y = 1 Then
                                '    If shop = "alice" Then
                                '        result = selectShijiBySelect_index(3, "normal_delvdate_id_1_0")
                                '    Else
                                '        result = selectShijiBySelect(shizistr, "normal_delvdate_id_1_0")
                                '    End If
                                'Else
                                '    If shop = "alice" Then
                                '        result = selectShijiBySelect_index(3, "normal_delvdate_id_1_" & node_y)
                                '    Else
                                '        result = selectShijiBySelect(shizistr, "normal_delvdate_id_1_" & node_y)
                                '    End If
                                'End If
                                If node_y = 1 And node_y_c = 1 Then
                                    result = selectShijiBySelect(shizistr, "normal_delvdate_id_1_0")
                                Else
                                    result = selectShijiBySelect(shizistr, "normal_delvdate_id_1_" & node_y)
                                End If
                                Dim result_ As String() = Split(result, "||")
                                If result_(1) = "true" Then
                                    If ComboBox4.SelectedIndex = 1 Then
                                        Dim click_flag As Boolean = False
                                        Dim eleSels As HtmlElementCollection = WebBrowser1.Document.GetElementsByTagName("input")
                                        For Each eleSel As HtmlElement In eleSels
                                            If eleSel.GetAttribute("value").Trim = "納期情報および在庫数を変更する" Then
                                                eleSel.InvokeMember("Click")
                                                WaitWebBrowser1Completed()

                                                If WebBrowser1.Document.Body.InnerText.Contains("R-Storefront編集禁止") <> "False" Then
                                                    MsgBox("現在、商品一括編集処理中のため、R-Storefrontでの編集作業はできません。", MsgBoxStyle.SystemModal)
                                                    Exit Sub
                                                End If

                                                click_flag = True
                                                Exit For
                                            End If
                                        Next
                                        If click_flag Then
                                            dgvResult(DataGridView1, r, True, True, TextBox2.Text & vbCrLf & code_ & ":" & result_(0) & " --> " & shizistr)
                                        Else
                                            dgvResult(DataGridView1, r, False, True, "")
                                            success = False
                                            Continue For
                                        End If
                                    Else
                                        dgvResult(DataGridView1, r, True, True, TextBox2.Text & vbCrLf & code_ & ":" & result_(0) & " --> " & shizistr)
                                    End If
                                Else
                                    dgvResult(DataGridView1, r, False, True, "")
                                    success = False
                                    Continue For
                                End If
                            End If
                        ElseIf code_m.Count = 3 And isDaihyo = False Then
                            eda_x = "--" & code_m(1)
                            Dim eda_y As String = code_m(2)

                            For Each node As HtmlAgilityPack.HtmlNode In nodes
                                Dim str As String = node.InnerText
                                str = Regex.Replace(str, "\r|\n", " ")
                                str = Replace(str, " ", "")
                                'Console.WriteLine(str)

                                If InStr(str, "同行一括") > 0 Or InStr(str, "/-") > 0 Then '列
                                    If InStr(str, "同行一括") > 0 Then
                                        str = Replace(str, "同行一括", "")
                                        node_y_c += 1

                                        If InStr(str, "/-") Then
                                            Dim str_sp As String() = Split(str, "/")
                                            If str_sp(1) = "--" & eda_y Or str_sp(1) = "-" & eda_y Then
                                                node_y = node_y_c
                                                node_y_ｓ += 1
                                                If eda_name <> "" Then
                                                    eda_name = eda_name & "/" & str_sp(0)
                                                Else
                                                    eda_name = str_sp(0)
                                                End If
                                            End If
                                        End If
                                    ElseIf InStr(str, "同列一括") > 0 Or InStr(str, "/-") > 0 Then '行 
                                        str = Replace(str, "同列一括", "")
                                        node_x_c += 1

                                        If InStr(str, "/-") Then
                                            Dim str_sp As String() = Split(str, "/")
                                            If str_sp.Length = 3 Then
                                                If str_sp(2) = eda_x Then
                                                    node_x = node_x_c
                                                    node_x_ｓ += 1
                                                    If eda_name <> "" Then
                                                        eda_name = eda_name & "/" & str_sp(0)
                                                    Else
                                                        eda_name = str_sp(0)
                                                    End If
                                                End If
                                            ElseIf str_sp.Length = 2 Then
                                                If str_sp(1) = eda_x Then
                                                    node_x = node_x_c
                                                    node_x_ｓ += 1
                                                    If eda_name <> "" Then
                                                        eda_name = eda_name & "/" & str_sp(0)
                                                    Else
                                                        eda_name = str_sp(0)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next

                            If node_x_c = 0 Or node_y_c = 0 Then '無し
                                dgvResult(DataGridView1, r, False, True, "")
                                success = False
                                Continue For
                            End If

                            If node_y_ｓ > 1 Or node_y_ｓ > 1 Then '同じ枝個数が多い
                                dgvResult(DataGridView1, r, False, True, "")
                                success = False
                                Continue For
                            End If

                            Dim result As String = ""
                            If shop = "alice" Then
                                result = selectShijiBySelect_index(3, "normal_delvdate_id_" & node_x & "_" & node_y)
                            Else
                                result = selectShijiBySelect(shizistr, "normal_delvdate_id_" & node_x & "_" & node_y)
                            End If

                            Dim result_ As String() = Split(result, "||")
                            If result_(1) = "true" Then
                                If ComboBox4.SelectedIndex = 1 Then
                                    Dim click_flag As Boolean = False
                                    Dim eleSels As HtmlElementCollection = WebBrowser1.Document.GetElementsByTagName("input")
                                    For Each eleSel As HtmlElement In eleSels
                                        If eleSel.GetAttribute("value").Trim = "納期情報および在庫数を変更する" Then
                                            eleSel.InvokeMember("Click")
                                            WaitWebBrowser1Completed()

                                            If WebBrowser1.Document.Body.InnerText.Contains("R-Storefront編集禁止") <> "False" Then
                                                MsgBox("現在、商品一括編集処理中のため、R-Storefrontでの編集作業はできません。", MsgBoxStyle.SystemModal)
                                                Exit Sub
                                            End If

                                            click_flag = True
                                            Exit For
                                        End If
                                    Next

                                    If click_flag Then
                                        dgvResult(DataGridView1, r, True, True, TextBox2.Text & vbCrLf & code_ & ":" & result_(0) & " --> " & shizistr)
                                    Else
                                        dgvResult(DataGridView1, r, False, True, "")
                                        success = False
                                        Continue For
                                    End If
                                Else
                                    dgvResult(DataGridView1, r, True, True, TextBox2.Text & vbCrLf & code_ & ":" & result_(0) & " --> " & shizistr)
                                End If
                            Else
                                dgvResult(DataGridView1, r, False, True, "")
                                success = False
                                Continue For
                            End If
                        End If

                        If success Then
                            WebBrowser1.Navigate(New Uri("https://item.rms.rakuten.co.jp/rms/mall/rsf/item/vc?__event=RI03_003_003&shop_bid=" & shop_bid & "&mng_number=" & search_code))
                            WaitWebBrowser1Completed()
                            If WebBrowser1.Document.GetElementById("normal_delvdate_id").GetAttribute("selectedIndex") = 0 Then
                                addyoyaku = False
                            Else
                                addyoyaku = True
                            End If

                            If ComboBox4.SelectedIndex = 1 Then
                                If shop = "akane" Then
                                    '商品名
                                    WebBrowser1.Navigate(New Uri("https://item.rms.rakuten.co.jp/rms/mall/rsf/item/vc?__event=RI03_001_002&shop_bid=" & shop_bid & "&mng_number=" & search_code))
                                    WaitWebBrowser1Completed()

                                    If InStr(WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).GetAttribute("value"), "＜一部商品予約＞") > 0 Then
                                        WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).SetAttribute("value", WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).GetAttribute("value").Replace("＜一部商品予約＞", ""))
                                    ElseIf InStr(WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).GetAttribute("value"), "＜一部予約＞") > 0 Then
                                        WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).SetAttribute("value", WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).GetAttribute("value").Replace("＜一部予約＞", ""))
                                    ElseIf InStr(WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).GetAttribute("value"), "＜予約＞") > 0 Then
                                        WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).SetAttribute("value", WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).GetAttribute("value").Replace("＜予約＞", ""))
                                    End If

                                    If addyoyaku Then
                                        WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).SetAttribute("value", "＜予約＞" & WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).GetAttribute("value"))
                                    Else
                                        WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).SetAttribute("value", "＜一部商品予約＞" & WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).GetAttribute("value"))
                                    End If

                                    WebBrowser1.Document.GetElementById("submitButton").InvokeMember("Click")
                                    WaitWebBrowser1Completed()

                                    If WebBrowser1.Document.Body.InnerText.Contains("R-Storefront編集禁止") <> "False" Then
                                        dgvResult(DataGridView1, r, False, False, "")
                                        MsgBox("現在、商品一括編集処理中のため、R-Storefrontでの編集作業はできません。", MsgBoxStyle.SystemModal)
                                        Exit Sub
                                    End If
                                End If
                            End If

                            '項目選択肢
                            If shop = "akane" Then
                                WebBrowser1.Navigate(New Uri("https://item.rms.rakuten.co.jp/rms/mall/rsf/item/vc?__event=RI03_003_002&shop_bid=" & shop_bid & "&mng_number=" & search_code))
                                WaitWebBrowser1Completed()

                                If ComboBox4.SelectedIndex = 1 Then
                                    Dim eleSels As HtmlElementCollection = WebBrowser1.Document.GetElementsByTagName("input")
                                    For Each eleSel As HtmlElement In eleSels
                                        If eleSel.GetAttribute("value").Trim = "項目選択肢を追加する" Then
                                            eleSel.InvokeMember("Click")
                                            WaitWebBrowser1Completed()

                                            If WebBrowser1.Document.Body.InnerText.Contains("R-Storefront編集禁止") <> "False" Then
                                                MsgBox("現在、商品一括編集処理中のため、R-Storefrontでの編集作業はできません。", MsgBoxStyle.SystemModal)
                                                Exit Sub
                                            End If
                                            Exit For
                                        End If
                                    Next
                                    If shizistr.Contains("入荷待ち(") Then
                                        Dim nichiji As String() = shizistr.Split("(")
                                        Dim nichiji_ As String = nichiji(1).Replace(")", "")
                                        WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("choice_name")(0).SetAttribute("value", "☆予約商品（" & eda_name & "）" & nichiji_)
                                    Else
                                        WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("choice_name")(0).SetAttribute("value", "☆予約商品（" & eda_name & "）" & shizistr)
                                    End If

                                    WebBrowser1.Document.GetElementsByTagName("input").GetElementsByName("value_1")(0).SetAttribute("value", "了承済み！")

                                    eleSels = WebBrowser1.Document.GetElementsByTagName("input")
                                    For Each eleSel As HtmlElement In eleSels
                                        If eleSel.GetAttribute("value").Trim = "項目選択肢を登録する" Then
                                            eleSel.InvokeMember("Click")
                                            WaitWebBrowser1Completed()

                                            If WebBrowser1.Document.Body.InnerText.Contains("R-Storefront編集禁止") <> "False" Then
                                                MsgBox("現在、商品一括編集処理中のため、R-Storefrontでの編集作業はできません。", MsgBoxStyle.SystemModal)
                                                Exit Sub
                                            End If
                                            Exit For
                                        End If
                                    Next
                                    dgvResult(DataGridView1, r, True, True, "")
                                End If
                            End If
                        End If
                    Else
                        '枝番なし
                        Dim result As String = selectShijiBySelect(shizistr, "normal_delvdate_id")
                        Dim result_ As String() = Split(result, "||")
                        If result_(1) = "true" Then
                            If shop = "akane" Then
                                '商品名
                                If InStr(WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).GetAttribute("value"), "＜一部商品予約＞") > 0 Then
                                    WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).SetAttribute("value", WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).GetAttribute("value").Replace("＜一部商品予約＞", ""))
                                ElseIf InStr(WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).GetAttribute("value"), "＜一部予約＞") > 0 Then
                                    WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).SetAttribute("value", WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).GetAttribute("value").Replace("＜一部予約＞", ""))
                                ElseIf InStr(WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).GetAttribute("value"), "＜予約＞") > 0 Then
                                    WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).SetAttribute("value", WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).GetAttribute("value").Replace("＜予約＞", ""))
                                End If

                                WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).SetAttribute("value", "＜予約＞" & WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("item_name")(0).GetAttribute("value"))
                            End If

                            If ComboBox4.SelectedIndex = 1 Then
                                WebBrowser1.Document.GetElementById("submitButton").InvokeMember("Click")
                                WaitWebBrowser1Completed()

                                If WebBrowser1.Document.Body.InnerText.Contains("R-Storefront編集禁止") <> "False" Then
                                    dgvResult(DataGridView1, r, False, False, "")
                                    MsgBox("現在、商品一括編集処理中のため、R-Storefrontでの編集作業はできません。", MsgBoxStyle.SystemModal)
                                    Exit Sub
                                End If

                                If shop = "akane" Then
                                    komoku_delete1(shop_bid, search_code)

                                    WebBrowser1.Navigate(New Uri("https://item.rms.rakuten.co.jp/rms/mall/rsf/item/vc?__event=RI03_003_002&shop_bid=" & shop_bid & "&mng_number=" & search_code))
                                    WaitWebBrowser1Completed()

                                    Dim eleSels As HtmlElementCollection = WebBrowser1.Document.GetElementsByTagName("input")
                                    For Each eleSel As HtmlElement In eleSels
                                        If eleSel.GetAttribute("value").Trim = "項目選択肢を追加する" Then
                                            eleSel.InvokeMember("Click")
                                            WaitWebBrowser1Completed()

                                            If WebBrowser1.Document.Body.InnerText.Contains("R-Storefront編集禁止") <> "False" Then
                                                MsgBox("現在、商品一括編集処理中のため、R-Storefrontでの編集作業はできません。", MsgBoxStyle.SystemModal)
                                                Exit Sub
                                            End If
                                            Exit For
                                        End If
                                    Next
                                    WebBrowser1.Document.GetElementsByTagName("textarea").GetElementsByName("choice_name")(0).SetAttribute("value", "☆予約商品" & shizistr)

                                    WebBrowser1.Document.GetElementsByTagName("input").GetElementsByName("value_1")(0).SetAttribute("value", "了承済み！")

                                    eleSels = WebBrowser1.Document.GetElementsByTagName("input")
                                    For Each eleSel As HtmlElement In eleSels
                                        If eleSel.GetAttribute("value").Trim = "項目選択肢を登録する" Then
                                            eleSel.InvokeMember("Click")
                                            WaitWebBrowser1Completed()

                                            If WebBrowser1.Document.Body.InnerText.Contains("R-Storefront編集禁止") <> "False" Then
                                                MsgBox("現在、商品一括編集処理中のため、R-Storefrontでの編集作業はできません。", MsgBoxStyle.SystemModal)
                                                Exit Sub
                                            End If
                                            Exit For
                                        End If
                                    Next
                                    dgvResult(DataGridView1, r, True, True, TextBox2.Text & vbCrLf & code_ & ":" & result_(0) & " --> " & shizistr)
                                Else
                                    dgvResult(DataGridView1, r, True, True, TextBox2.Text & vbCrLf & code_ & ":" & result_(0) & " --> " & shizistr)
                                End If
                            Else
                                dgvResult(DataGridView1, r, True, True, TextBox2.Text & vbCrLf & code_ & ":" & result_(0) & " --> " & shizistr)
                            End If
                        Else
                            dgvResult(DataGridView1, r, False, True, "")
                        End If
                    End If
                End If
                shilyori_c += 1
            Next
        End If
        TextBox2.Text = TextBox2.Text & vbCrLf & "--------------------- " & (Format(DateTime.Now, "yyyy/MM/dd HH:mm:ss")).ToString & "  end  ---------------------"
    End Sub

    Private Sub komoku_delete1(shop_bid As String, search_code As String)
        For r As Integer = 0 To 4
            WebBrowser1.Navigate(New Uri("https://item.rms.rakuten.co.jp/rms/mall/rsf/item/vc?__event=RI03_003_002&shop_bid=" & shop_bid & "&mng_number=" & search_code))
            WaitWebBrowser1Completed()

            Dim click_flag As Boolean = False
            Dim eleSels As HtmlElementCollection = WebBrowser1.Document.GetElementsByTagName("table")(22).GetElementsByTagName("font")
            For i As Integer = 0 To eleSels.Count - 1
                If eleSels(i).GetElementsByTagName("input").Count > 0 Then
                    If eleSels(i).GetElementsByTagName("input")(0).GetAttribute("value").Trim = "削除する" Then
                        If InStr(eleSels(i).Parent.Parent.Parent.GetElementsByTagName("font")(0).InnerText, "予約") > 0 Then
                            eleSels(i).GetElementsByTagName("input")(0).InvokeMember("click")
                            WaitWebBrowser1Completed()
                            click_flag = True
                            Exit For
                        End If
                    End If
                End If
            Next

            '削除した場合
            If click_flag Then
                eleSels = WebBrowser1.Document.GetElementsByTagName("input")
                For Each eleSel As HtmlElement In eleSels
                    If eleSel.GetAttribute("value").Trim = "項目選択肢を削除する" Then
                        eleSel.InvokeMember("Click")
                        WaitWebBrowser1Completed()

                        If WebBrowser1.Document.Body.InnerText.Contains("R-Storefront編集禁止") <> "False" Then
                            MsgBox("現在、商品一括編集処理中のため、R-Storefrontでの編集作業はできません。", MsgBoxStyle.SystemModal)
                            Exit Sub
                        End If
                        Exit For
                    End If
                Next
            Else
                Exit For
            End If
        Next
    End Sub

    Private Function wb1_rtlogin(rakuten_login As String) As Boolean
        Dim rakuten_login_ As String() = Split(rakuten_login, ",")

        If rakuten_login_.Count = 4 Then
            WebBrowser1.ScriptErrorsSuppressed = True
            WebBrowser1.Navigate(New Uri("https://glogin.rms.rakuten.co.jp/?sp_id=1"))
            WaitWebBrowser1Completed()
            Application.DoEvents()
            '--------------------------- ログイン start ---------------------------

            'Dim objDoc As HtmlDocument = WebBrowser1.Document

            WebBrowser1.Document.GetElementById("rlogin-username-ja").SetAttribute("value", rakuten_login_(0))
            WebBrowser1.Document.GetElementById("rlogin-password-ja").SetAttribute("value", rakuten_login_(1))
            WebBrowser1.Document.Body.All.GetElementsByName("submit")(0).InvokeMember("Click")
            WaitWebBrowser1Completed()

            If WebBrowser1.Document.GetElementById("rlogin-username-2-ja") Is Nothing Then
                MsgBox("パスワードを確認してください。", MsgBoxStyle.SystemModal)
                Return False
            End If

            WebBrowser1.Document.GetElementById("rlogin-username-2-ja").SetAttribute("value", rakuten_login_(2))
            WebBrowser1.Document.GetElementById("rlogin-password-2-ja").SetAttribute("value", rakuten_login_(3))
            WebBrowser1.Document.Body.All.GetElementsByName("submit")(0).InvokeMember("Click")
            WaitWebBrowser1Completed()

            If WebBrowser1.Document.Body.All.GetElementsByName("submit") Is Nothing Then
            Else
                WebBrowser1.Document.Body.All.GetElementsByName("submit")(0).InvokeMember("Click")
                WaitWebBrowser1Completed()
            End If

            WebBrowser1.Document.GetElementById("confirm").InvokeMember("Submit")
            WaitWebBrowser1Completed()

            Dim eleSels As HtmlElementCollection = WebBrowser1.Document.GetElementsByTagName("button")
            For Each eleSel As HtmlElement In eleSels
                If eleSel.InnerText IsNot Nothing Then
                    If eleSel.InnerText.Trim = "RMSメインメニューへ進む" Then
                        eleSel.InvokeMember("Click")
                        WaitWebBrowser1Completed()
                        Exit For
                    End If
                End If
            Next

            '--------------------------- ログイン end ---------------------------
            Return True
        Else
            MsgBox("ログイン情報を確認してください。", MsgBoxStyle.SystemModal)
            Return False
        End If
    End Function

    '行番号を表示する
    Private Sub DataGridView_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles _
            DataGridView1.RowPostPaint, DGV6.RowPostPaint
        Dim rect As New Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, sender.RowHeadersWidth - 4, sender.Rows(e.RowIndex).Height)
        TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), sender.RowHeadersDefaultCellStyle.Font, rect, sender.RowHeadersDefaultCellStyle.ForeColor, TextFormatFlags.VerticalCenter Or TextFormatFlags.Right)
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Button2.BackColor = Color.Yellow
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ToolStripStatusLabel1.Text = Timer1.Interval / 1000
        locationPath = Form1.ロケーションToolStripMenuItem.Text

        Dim fi1 As New FileInfo(Path.GetDirectoryName(Form1.appPath) & "\location.csv")
        TextBox24.Text = fi1.LastWriteTime.ToOADate

        'サーバーの接続を確認する
        If File.Exists(locationPath & "\location.csv") Then
            Dim fi2 As New FileInfo(locationPath & "\location.csv")
            Dim check = fi2.LastAccessTime
            ToolStripStatusLabel2.Text = "接続OK"
            ToolStripStatusLabel2.ForeColor = Color.Black
            ToolStripStatusLabel2.BackColor = Color.Empty
        Else
            ToolStripStatusLabel2.Text = "未接続"
            ToolStripStatusLabel2.ForeColor = Color.White
            ToolStripStatusLabel2.BackColor = Color.Red

            Exit Sub
        End If

        Try
            Dim fi2 As New FileInfo(locationPath & "\location.csv")
            TextBox25.Text = fi2.LastWriteTime.ToOADate
            If TextBox25.Text < 0 Or TextBox24.Text < 0 Then
                TextBox17.Text = "error"
                Timer1.Interval = 1000 * 5
                ToolStripStatusLabel1.Text = Timer1.Interval / 1000
            ElseIf TextBox25.Text - TextBox24.Text > 0 Then
                Form1.NotifyIcon1.BalloonTipTitle = "配送マスタのアップデートを感知しました"
                Form1.NotifyIcon1.BalloonTipText = "自動でアップデートを行ないます。再起動の必要はありません。"
                Form1.NotifyIcon1.ShowBalloonTip(1000)

                TextBox17.Text = "update"
                Dim sPath As String = locationPath & "\location.csv"
                Dim mPath As String = Path.GetDirectoryName(Form1.appPath) & "\location.csv"
                File.Copy(sPath, mPath, True)
                Timer1.Interval = 1000 * 5
                ToolStripStatusLabel1.Text = Timer1.Interval / 1000

                'datagridview6再表示
                DGV6_update()
                Application.DoEvents()
            Else
                TextBox17.Text = "OK"
                Timer1.Interval = 1000 * 60 * 3
                ToolStripStatusLabel1.Text = Timer1.Interval / 1000

                'datagridview6再表示
                DGV6_update()
                Application.DoEvents()
            End If
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub

    Private Sub DGV6_update()
        'datagridview6再表示
        If DGV6.RowCount > 0 Then
            DGV6.Rows.Clear()
            DGV6.Columns.Clear()
        End If

        'ヘッダー行作成
        Dim header As String() = New String() {"商品コード", "商品名", "商品分類タグ", "代表商品コード", "ship-weight", "ロケーション", "梱包サイズ", "特殊", "sw2"}
        For c As Integer = 0 To header.Length - 1
            DGV6.Columns.Add(c, header(c))
            DGV6.Columns(c).SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        Dim fNameL As String = Path.GetDirectoryName(Form1.appPath) & "\location.csv"
        Dim csvRecords2 As New ArrayList()
        csvRecords2 = TM_CSV_READ(fNameL)(0)
        Dim codeArray2 As New ArrayList()
        For r As Integer = 0 To csvRecords2.Count - 1
            Dim sArray As String() = Split(csvRecords2(r), "|=|")
            codeArray2.Add(sArray(0))
        Next

        For r As Integer = 0 To csvRecords2.Count - 1
            Dim sArray As String() = Split(csvRecords2(r), "|=|")
            sArray(0) = Form1.StrConvToNarrow(sArray(0))    '商品コードを小文字で揃える
            DGV6.Rows.Add(sArray)
        Next
    End Sub

    Private Sub フォームのリセットToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles フォームのリセットToolStripMenuItem.Click
        Me.Dispose()
        Dim frm As Form = New Shiji
        frm.Show()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        DGV6_update()
    End Sub

    Private Function findSelectById(element As String) As String
        Dim eleSels As HtmlElementCollection = WebBrowser1.Document.GetElementsByTagName("select")
        Dim item As String = ""
        For Each eleSel As HtmlElement In eleSels
            If eleSel.Name = element Then
                For i As Integer = 0 To eleSel.GetElementsByTagName("option").Count - 1
                    If eleSel.GetElementsByTagName("option").Item(i).GetAttribute("selected") = "True" Then
                        item = eleSel.GetElementsByTagName("option").Item(i).InnerText
                    End If
                Next
            End If
        Next
        Return item
    End Function

    Private Function selectShijiBySelect(element As String) As String
        'Dim str As String = ""
        'Dim action As Boolean = False
        'Dim moto As String = ""
        'Dim eleSels As HtmlElementCollection = WebBrowser1.Document.GetElementsByTagName("select")
        'For Each eleSel As HtmlElement In eleSels
        '    If eleSel.Name = element Then
        '        Dim eleOpts As HtmlElementCollection = eleSel.GetElementsByTagName("option")
        '        'eleSel.GetElementsByTagName("option").Item(1).OuterHtml
        '        For i As Integer = 0 To eleSel.GetElementsByTagName("option").Count - 1
        '            If eleSel.GetElementsByTagName("option").Item(i).GetAttribute("selected") = "True" Then
        '                moto = eleSel.GetElementsByTagName("option").Item(i).InnerText
        '            End If
        '        Next
        '        For i As Integer = 0 To eleSel.GetElementsByTagName("option").Count - 1
        '            If eleSel.GetElementsByTagName("option").Item(i).OuterHtml.Contains(shizi) Then
        '                eleSel.GetElementsByTagName("option").Item(i).SetAttribute("selected", "selected")
        '                action = True
        '            End If
        '        Next
        '    End If
        'Next
        'Application.DoEvents()
        'If action Then
        '    str = moto & "||true"
        'Else
        '    str = moto & "||false"
        'End If
        'Return str
    End Function

    Private Function selectShijiBySelect_index(index As Integer, element As String) As String
        Dim str As String = ""
        Dim action As Boolean = False
        Dim moto As String = ""
        Dim eleSels As HtmlElementCollection = WebBrowser1.Document.GetElementsByTagName("select")
        For Each eleSel As HtmlElement In eleSels
            If eleSel.Name = element Then
                Dim eleOpts As HtmlElementCollection = eleSel.GetElementsByTagName("option")
                'eleSel.GetElementsByTagName("option").Item(1).OuterHtml
                For i As Integer = 0 To eleSel.GetElementsByTagName("option").Count - 1
                    If eleSel.GetElementsByTagName("option").Item(i).GetAttribute("selected") = "True" Then
                        moto = eleSel.GetElementsByTagName("option").Item(i).InnerText
                        eleSel.GetElementsByTagName("option").Item(index).SetAttribute("selected", "selected")
                        action = True
                        Exit For
                    End If
                Next
                Exit For
            End If
        Next
        Application.DoEvents()
        If action Then
            str = moto & "||true"
        Else
            str = moto & "||false"
        End If
        Return str
    End Function

    Private Function selectShijiBySelect(shizi As String, element As String) As String
        Dim str As String = ""
        Dim action As Boolean = False
        Dim moto As String = ""
        Dim eleSels As HtmlElementCollection = WebBrowser1.Document.GetElementsByTagName("select")
        For Each eleSel As HtmlElement In eleSels
            If eleSel.Name = element Then
                Dim eleOpts As HtmlElementCollection = eleSel.GetElementsByTagName("option")
                'eleSel.GetElementsByTagName("option").Item(1).OuterHtml
                For i As Integer = 0 To eleSel.GetElementsByTagName("option").Count - 1
                    If eleSel.GetElementsByTagName("option").Item(i).GetAttribute("selected") = "True" Then
                        moto = eleSel.GetElementsByTagName("option").Item(i).InnerText
                        Exit For
                    End If
                Next
                For i As Integer = 0 To eleSel.GetElementsByTagName("option").Count - 1
                    If eleSel.GetElementsByTagName("option").Item(i).OuterHtml.Contains(shizi) Then
                        eleSel.GetElementsByTagName("option").Item(i).SetAttribute("selected", "selected")
                        action = True
                    End If
                Next
                Exit For
            End If
        Next
        Application.DoEvents()
        If action Then
            str = moto & "||true"
        Else
            str = moto & "||false"
        End If
        Return str
    End Function

    Private Sub dgvResult(dgv As DataGridView, r As Integer, flag As Boolean, flag2 As Boolean, log As String)
        If flag2 Then
            If flag Then
                dgv.Item(2, r).Value = "○"
                dgv.Item(0, r).Style.BackColor = Color.LightGreen
            Else
                dgv.Item(2, r).Value = "×"
                dgv.Item(0, r).Style.BackColor = Color.Pink
            End If
        End If

        If log <> "" Then
            TextBox2.Text = log
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Select Case ComboBox1.SelectedIndex
            Case 0
                CheckBox1.Checked = True
                CheckBox3.Checked = True
            Case 1
                CheckBox1.Checked = False
                CheckBox3.Checked = False
            Case 2
                CheckBox1.Checked = False
                CheckBox3.Checked = False
        End Select
        Button2.BackColor = Color.Yellow
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Button2.BackColor = Color.Yellow
    End Sub
End Class