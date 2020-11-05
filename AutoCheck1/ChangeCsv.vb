Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports Microsoft.International.Converters
Imports Microsoft.VisualBasic.FileIO
Imports Ookii.Dialogs
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel
Imports PdfGeneratorNetFree
Imports PdfGeneratorNetFree.PdfContentItem
Imports PdfGeneratorNetFree.PgnStyle
Imports System.Windows.Forms

Public Class ChangeCsv
    ReadOnly ENC_SJ As Encoding = Encoding.GetEncoding("shift-jis")
    Private dH1 As New ArrayList
    Private dH3 As New ArrayList
    Private dH4 As New ArrayList

    Private Sub NeCsv_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DGV1.Rows.Clear()
        DGV2.Rows.Clear()
        DGV3.Rows.Clear()
        DGV4.Rows.Clear()

        dH1 = TM_HEADER_GET(DGV1)
        dH3 = TM_HEADER_GET(DGV3)
        dH4 = TM_HEADER_GET(DGV4)
    End Sub

    Private Sub DataGridView_DragDrop(ByVal sender As DataGridView, ByVal e As System.Windows.Forms.DragEventArgs) Handles _
             DGV1.DragDrop

        DGV1.Rows.Clear()
        DGV2.Rows.Clear()
        DGV3.Rows.Clear()
        DGV4.Rows.Clear()

        Dim dataPathArray As New ArrayList
        For Each filename As String In e.Data.GetData(DataFormats.FileDrop)
            dataPathArray.Add(filename)
        Next

        If dataPathArray.Count > 0 Then
            For i As Integer = 0 To dataPathArray.Count - 1
                DGV2.Rows.Add(Path.GetFileName(dataPathArray(i)))
            Next
        End If

        '######################################################################################################################
        FileDataRead(dataPathArray)
        '######################################################################################################################
    End Sub

    Private Sub DataGridView_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles _
            DGV1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    '行番号を表示する
    Private Sub DataGridView_RowPostPaint(ByVal sender As DataGridView, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles _
            DGV1.RowPostPaint, DGV2.RowPostPaint, DGV3.RowPostPaint, DGV4.RowPostPaint
        ' 行ヘッダのセル領域を、行番号を描画する長方形とする（ただし右端に4ドットのすき間を空ける）
        Dim rect As New Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, sender.RowHeadersWidth - 4, sender.Rows(e.RowIndex).Height)
        ' 上記の長方形内に行番号を縦方向中央＆右詰で描画する　フォントや色は行ヘッダの既定値を使用する
        TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), sender.RowHeadersDefaultCellStyle.Font, rect, sender.RowHeadersDefaultCellStyle.ForeColor, TextFormatFlags.VerticalCenter Or TextFormatFlags.Right)
    End Sub

    Public Sub FileDataRead(dataPathArray As ArrayList)
        Dim dgv As DataGridView = DGV1
        Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)

        Dim allAdd As Boolean = False   '読み込んだファイルを全て追加するか
        Dim dgv1NoArray As New ArrayList    '重複確認用Array
        Dim dgv1ErrorArray As New ArrayList

        'If DGV1.RowCount > 0 Then
        '    For r As Integer = 0 To DGV1.RowCount - 1
        '        Dim no As String = DGV1.Item(dH1.IndexOf("伝票番号"), r).Value & "," & DGV1.Item(dH1.IndexOf("受注番号"), r).Value
        '        dgv1NoArray.Add(no)
        '    Next
        'End If
        'If DGV3.RowCount > 0 Then
        '    Dim dH3 As ArrayList = TM_HEADER_GET(DGV3)
        '    For r As Integer = 0 To DGV3.RowCount - 1
        '        Dim no As String = DGV3.Item(dH3.IndexOf("伝票番号"), r).Value & "," & DGV3.Item(dH3.IndexOf("明細行"), r).Value
        '        dgv3NoArray.Add(no)
        '    Next
        'End If

        'テンプレート分析用
        'Dim fBunsekiArray As String() = File.ReadAllLines(Path.GetDirectoryName(Form1.appPath) & "\template\☆共通-ファイル分析.dat", ENC_SJ)
        'Dim fBnameArray As New ArrayList
        'Dim fBArray As New ArrayList
        'Dim fBHeader As String() = Split(fBunsekiArray(0), ",")
        'For i As Integer = 0 To fBHeader.Length - 1
        '    Dim str As String = ""
        '    For k As Integer = 0 To fBunsekiArray.Length - 1
        '        Dim fBline As String() = Split(fBunsekiArray(k), ",")
        '        If k = 0 Then
        '            fBnameArray.Add(fBline(i))
        '        Else
        '            If fBline(i) <> "" Then
        '                str &= fBline(i) & ","
        '            Else
        '                Exit For
        '            End If
        '        End If
        '    Next
        '    str = str.TrimEnd(",")
        '    fBArray.Add(str)
        'Next

        'dataPathArray.Sort()
        'dataPathArray.Reverse()
        'For Each filename As String In dataPathArray
        '    Dim csvRecords As New ArrayList()
        '    csvRecords = TM_CSV_READ(filename)(0)

        '    'ファイル認識で分別する
        '    Dim cHeader As String = Replace(csvRecords(0), "|=|", ",")
        '    Dim fB As String = ""
        '    For i As Integer = 0 To fBArray.Count - 1
        '        If fBArray(i) = cHeader Then
        '            fB = fBnameArray(i)
        '            Exit For
        '        End If
        '    Next

        '    Select Case fB
        '        'Case "明細"
        '        '    dgv = DGV3
        '        Case "画面"
        '            dgv = DGV1
        '        Case Else
        '            MsgBox(filename & vbCrLf & "ファイルの認識ができません。")
        '            Exit Sub
        '    End Select

        '    For r As Integer = 0 To csvRecords.Count - 1
        '        Dim sArray As String() = Split(csvRecords(r), "|=|")
        '        If dgv.ColumnCount < sArray.Length - 1 Then
        '            For c As Integer = dgv.ColumnCount To sArray.Length - 1
        '                dgv.Columns.Add(c, sArray(c))
        '                dgv.Columns(c).SortMode = DataGridViewColumnSortMode.NotSortable
        '            Next
        '        End If

        '        If r > 0 Then
        '            dgv.Rows.Add(sArray)
        '        End If
        '    Next
        'Next

        Dim check_header1 As String = "お届け先コード,お届け先郵便番号,お届け先住所１,お届け先住所２,お届け先住所３,お届け先名１,お届け先名２,お届け先電話,お届け先メールアドレス,代行ご依頼主郵便番号,代行ご依頼主住所１,代行ご依頼主住所２,代行ご依頼主名１,代行ご依頼主名２,代行ご依頼主電話,代行ご依頼主メールアドレス,佐川急便顧客コード,問い合せ№,発送日,個数,顧客管理番号,便種コード,配達指定日,時間帯コード,代引金額,代引消費税,保険金額,元着区分（未使用）,営止め区分,営止精算店コード,営止精算店ローカルコード,記事欄01,記事欄02,記事欄03,記事欄04,記事欄05,記事欄06,記事欄07,記事欄08,記事欄09,記事欄10,記事欄11,記事欄12,マスタコード,マスタ配送,マスタ個口,マスタ備考,マーク,処理用,処理用2"
        For Each filename As String In dataPathArray
            Dim csvRecords As New ArrayList()
            csvRecords = TM_CSV_READ(filename)(0)
            Dim cHeader As String = Replace(csvRecords(0), "|=|", ",")

            If check_header1 = cHeader Then
                dgv = DGV1
            Else
                MsgBox(filename & vbCrLf & "ファイルの認識ができません。")
                Exit Sub
            End If

            For r As Integer = 0 To csvRecords.Count - 1
                Dim sArray As String() = Split(csvRecords(r), "|=|")
                If dgv.ColumnCount < sArray.Length - 1 Then
                    For c As Integer = dgv.ColumnCount To sArray.Length - 1
                        dgv.Columns.Add(c, sArray(c))
                        dgv.Columns(c).SortMode = DataGridViewColumnSortMode.NotSortable
                    Next
                End If
                If r > 0 Then
                    dgv.Rows.Add(sArray)
                End If
            Next
        Next
    End Sub

    Private Sub フォームのリセットToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles フォームのリセットToolStripMenuItem.Click
        Me.Dispose()
        Dim frm As Form = New ChangeCsv
        frm.Show()
    End Sub

    Private Sub KryptonButton1_Click_1(sender As Object, e As EventArgs) Handles KryptonButton1.Click
        If DGV1.RowCount = 0 Then
            MsgBox("ファイルを入れてください。", MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        DGV3.Rows.Clear()
        DGV4.Rows.Clear()

        Dim dgvIndex As Integer = -1
        For r As Integer = 0 To DGV1.RowCount - 1
            Dim master_code As String = DGV1.Item(dH1.IndexOf("マスタコード"), r).Value
            Dim masterCodeArray As String() = Split(master_code, "、")

            dgvIndex = dgvIndex + 1
            DGV3.Rows.Add()

            '同梱
            If masterCodeArray.Length > 1 Then
                DGV3.Item(dH3.IndexOf("受注数"), dgvIndex).Value = "同梱"
                DGV3.Item(dH3.IndexOf("受注数"), dgvIndex).Style.BackColor = Color.Lavender
            Else
                Dim code As String() = Split(masterCodeArray(0), "*")
                If code Is Nothing Then
                    DGV3.Item(dH3.IndexOf("受注数"), dgvIndex).Style.BackColor = Color.Red
                Else
                    If code.Length = 1 Then
                        DGV3.Item(dH3.IndexOf("受注数"), dgvIndex).Style.BackColor = Color.Red
                    Else
                        If IsNumeric(code(1)) Then
                            DGV3.Item(dH3.IndexOf("受注数"), dgvIndex).Value = code(1)
                        Else
                            DGV3.Item(dH3.IndexOf("受注数"), dgvIndex).Style.BackColor = Color.Red
                        End If
                    End If
                End If
            End If

            dgv3AddRowFromDgv1("お届け先コード", r, dgvIndex)
            dgv3AddRowFromDgv1("お届け先コード", r, dgvIndex)
            dgv3AddRowFromDgv1("お届け先郵便番号", r, dgvIndex)
            dgv3AddRowFromDgv1("お届け先住所１", r, dgvIndex)
            dgv3AddRowFromDgv1("お届け先住所２", r, dgvIndex)
            dgv3AddRowFromDgv1("お届け先住所３", r, dgvIndex)
            dgv3AddRowFromDgv1("お届け先名１", r, dgvIndex)
            dgv3AddRowFromDgv1("お届け先名２", r, dgvIndex)
            dgv3AddRowFromDgv1("お届け先名２", r, dgvIndex)
            dgv3AddRowFromDgv1("お届け先電話", r, dgvIndex)
            dgv3AddRowFromDgv1("お届け先メールアドレス", r, dgvIndex)
            dgv3AddRowFromDgv1("代行ご依頼主郵便番号", r, dgvIndex)
            dgv3AddRowFromDgv1("代行ご依頼主住所１", r, dgvIndex)
            dgv3AddRowFromDgv1("代行ご依頼主住所２", r, dgvIndex)
            dgv3AddRowFromDgv1("代行ご依頼主名１", r, dgvIndex)
            dgv3AddRowFromDgv1("代行ご依頼主名２", r, dgvIndex)
            dgv3AddRowFromDgv1("代行ご依頼主電話", r, dgvIndex)
            dgv3AddRowFromDgv1("代行ご依頼主メールアドレス", r, dgvIndex)
            dgv3AddRowFromDgv1("佐川急便顧客コード", r, dgvIndex)
            dgv3AddRowFromDgv1("問い合せ№", r, dgvIndex)
            dgv3AddRowFromDgv1("発送日", r, dgvIndex)
            dgv3AddRowFromDgv1("個数", r, dgvIndex)
            dgv3AddRowFromDgv1("顧客管理番号", r, dgvIndex)
            dgv3AddRowFromDgv1("便種コード", r, dgvIndex)
            dgv3AddRowFromDgv1("配達指定日", r, dgvIndex)
            dgv3AddRowFromDgv1("時間帯コード", r, dgvIndex)
            dgv3AddRowFromDgv1("代引金額", r, dgvIndex)
            dgv3AddRowFromDgv1("代引消費税", r, dgvIndex)
            dgv3AddRowFromDgv1("保険金額", r, dgvIndex)
            dgv3AddRowFromDgv1("元着区分（未使用）", r, dgvIndex)
            dgv3AddRowFromDgv1("営止め区分", r, dgvIndex)
            dgv3AddRowFromDgv1("営止精算店コード", r, dgvIndex)
            dgv3AddRowFromDgv1("営止精算店ローカルコード", r, dgvIndex)
            dgv3AddRowFromDgv1("記事欄01", r, dgvIndex)
            dgv3AddRowFromDgv1("記事欄02", r, dgvIndex)
            dgv3AddRowFromDgv1("記事欄03", r, dgvIndex)
            dgv3AddRowFromDgv1("記事欄04", r, dgvIndex)
            dgv3AddRowFromDgv1("記事欄05", r, dgvIndex)
            dgv3AddRowFromDgv1("記事欄06", r, dgvIndex)
            dgv3AddRowFromDgv1("記事欄07", r, dgvIndex)
            dgv3AddRowFromDgv1("記事欄08", r, dgvIndex)
            dgv3AddRowFromDgv1("記事欄09", r, dgvIndex)
            dgv3AddRowFromDgv1("記事欄10", r, dgvIndex)
            dgv3AddRowFromDgv1("記事欄11", r, dgvIndex)
            dgv3AddRowFromDgv1("記事欄12", r, dgvIndex)
            dgv3AddRowFromDgv1("マスタコード", r, dgvIndex)
            dgv3AddRowFromDgv1("マスタ配送", r, dgvIndex)
            dgv3AddRowFromDgv1("マスタ個口", r, dgvIndex)
            dgv3AddRowFromDgv1("マスタ備考", r, dgvIndex)
            dgv3AddRowFromDgv1("マーク", r, dgvIndex)
            dgv3AddRowFromDgv1("処理用", r, dgvIndex)
            dgv3AddRowFromDgv1("処理用2", r, dgvIndex)
        Next

        '佐川
        Dim errorCount As Integer = 0
        dgvIndex = -1
        For r As Integer = 0 To DGV1.RowCount - 1
            dgvIndex = dgvIndex + 1
            DGV4.Rows.Add()

            Dim master_code As String = DGV1.Item(dH1.IndexOf("マスタコード"), r).Value
            Dim masterCodeArray As String() = Split(master_code, "、")

            '同梱
            If masterCodeArray.Length > 1 Then
                DGV4.Item(dH3.IndexOf("受注数"), dgvIndex).Value = "同梱"
                DGV4.Item(dH3.IndexOf("受注数"), dgvIndex).Style.BackColor = Color.Lavender
            Else
                Dim code As String() = Split(masterCodeArray(0), "*")
                If code Is Nothing Then
                    DGV4.Item(dH4.IndexOf("受注数"), dgvIndex).Style.BackColor = Color.Red
                Else
                    If code.Length = 1 Then
                        DGV4.Item(dH4.IndexOf("受注数"), dgvIndex).Style.BackColor = Color.Red
                    Else
                        If IsNumeric(code(1)) Then
                            DGV4.Item(dH4.IndexOf("受注数"), dgvIndex).Value = code(1)
                        Else
                            errorCount = errorCount + 1
                            DGV4.Item(dH4.IndexOf("受注数"), dgvIndex).Style.BackColor = Color.Red
                        End If
                    End If
                End If
            End If

            DGV4.Item(dH4.IndexOf("顧客コード"), dgvIndex).Value = "957769750045"
            DGV4.Item(dH4.IndexOf("電話番号"), dgvIndex).Value = DGV1.Item(dH1.IndexOf("お届け先電話"), r).Value
            DGV4.Item(dH4.IndexOf("郵便番号"), dgvIndex).Value = DGV1.Item(dH1.IndexOf("お届け先郵便番号"), r).Value
            DGV4.Item(dH4.IndexOf("住所1"), dgvIndex).Value = DGV1.Item(dH1.IndexOf("お届け先住所１"), r).Value
            DGV4.Item(dH4.IndexOf("住所2"), dgvIndex).Value = DGV1.Item(dH1.IndexOf("お届け先住所２"), r).Value
            DGV4.Item(dH4.IndexOf("住所3"), dgvIndex).Value = DGV1.Item(dH1.IndexOf("お届け先住所３"), r).Value
            DGV4.Item(dH4.IndexOf("代行住所1"), dgvIndex).Value = DGV1.Item(dH1.IndexOf("代行ご依頼主住所１"), r).Value
            DGV4.Item(dH4.IndexOf("代行住所2"), dgvIndex).Value = DGV1.Item(dH1.IndexOf("代行ご依頼主住所２"), r).Value
            DGV4.Item(dH4.IndexOf("代行電話番号"), dgvIndex).Value = DGV1.Item(dH1.IndexOf("代行ご依頼主電話"), r).Value
            DGV4.Item(dH4.IndexOf("代行名称1"), dgvIndex).Value = DGV1.Item(dH1.IndexOf("代行ご依頼主名１"), r).Value

            Dim zikantai As String = DGV1.Item(dH1.IndexOf("時間帯コード"), dgvIndex).Value
            Select Case zikantai
                Case "指定無し"
                    DGV4.Item(dH4.IndexOf("配達指定時間帯"), dgvIndex).Value = "00"
                Case "午前中"
                    DGV4.Item(dH4.IndexOf("配達指定時間帯"), dgvIndex).Value = "01"
                Case "12時-14時"
                    DGV4.Item(dH4.IndexOf("配達指定時間帯"), dgvIndex).Value = "12"
                Case "14時-16時"
                    DGV4.Item(dH4.IndexOf("配達指定時間帯"), dgvIndex).Value = "14"
                Case "16時-18時"
                    DGV4.Item(dH4.IndexOf("配達指定時間帯"), dgvIndex).Value = "16"
                Case "18時-20時"
                    DGV4.Item(dH4.IndexOf("配達指定時間帯"), dgvIndex).Value = "18"
                Case "19時-21時"
                    DGV4.Item(dH4.IndexOf("配達指定時間帯"), dgvIndex).Value = "19"
                Case "18時-21時"
                    DGV4.Item(dH4.IndexOf("配達指定時間帯"), dgvIndex).Value = "04"
            End Select

            DGV4.Item(dH4.IndexOf("代引金額"), dgvIndex).Value = DGV1.Item(dH1.IndexOf("代引金額"), r).Value

            'Dim name As String = DGV1.Item(dH1.IndexOf("お届け先名１"), r).Value
            'Dim name_sp As String() = Split(name, ")")
            'If name_sp.Length > 1 Then
            '    DGV4.Item(dH4.IndexOf("名称1"), dgvIndex).Value = name_sp(1)
            'Else
            '    DGV4.Item(dH4.IndexOf("名称1"), dgvIndex).Value = name
            'End If
            DGV4.Item(dH4.IndexOf("名称1"), dgvIndex).Value = DGV1.Item(dH1.IndexOf("お届け先名１"), r).Value
            DGV4.Item(dH4.IndexOf("記事１"), dgvIndex).Value = DGV1.Item(dH1.IndexOf("記事欄01"), r).Value
            DGV4.Item(dH4.IndexOf("記事３"), dgvIndex).Value = DGV1.Item(dH1.IndexOf("記事欄06"), r).Value
            DGV4.Item(dH4.IndexOf("記事４"), dgvIndex).Value = DGV1.Item(dH1.IndexOf("記事欄07"), r).Value
            DGV4.Item(dH4.IndexOf("個数"), dgvIndex).Value = DGV1.Item(dH1.IndexOf("個数"), r).Value
            DGV4.Item(dH4.IndexOf("便種"), dgvIndex).Value = "000"
            DGV4.Item(dH4.IndexOf("顧客管理番号"), dgvIndex).Value = DGV1.Item(dH1.IndexOf("顧客管理番号"), r).Value
        Next
        MsgBox("処理しました。", MsgBoxStyle.SystemModal)
        If errorCount > 0 Then
            MsgBox("受注数ないデータがあります。", MsgBoxStyle.SystemModal)
        End If
    End Sub


    Private Sub dgv3AddRowFromDgv1(key As String, r As Integer, r2 As Integer)
        DGV3.Item(dH3.IndexOf(key), r2).Value = DGV1.Item(dH1.IndexOf(key), r).Value
    End Sub

    '伝票番号に対応する商品コードを「商品マスタ」に読み込む
    'Private Sub MasterInput()
    '    Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
    '    Dim dH3 As ArrayList = TM_HEADER_GET(DGV3)
    '    Dim error_count As Integer = 0

    '    For r1 As Integer = 0 To DGV1.RowCount - 1
    '        Dim dNum As String = DGV1.Item(dH1.IndexOf("伝票番号"), r1).Value
    '        Dim str As String = ""

    '        For r2 As Integer = 0 To DGV3.RowCount - 1
    '            If dNum = DGV3.Item(dH3.IndexOf("伝票番号"), r2).Value Then
    '                DGV3.Item(dH3.IndexOf("伝票番号"), r2).Style.BackColor = Color.LemonChiffon
    '                Dim codeHeader As String = "商品ｺｰﾄﾞ"
    '                Dim code As String = ""

    '                If InStr(DGV3.Item(dH3.IndexOf("商品ｺｰﾄﾞ"), r2).Value, "wake") > 0 Then    'ヤフオク用
    '                    If Regex.IsMatch(DGV3.Item(dH3.IndexOf(codeHeader), r2).Value, "^[a-zA-Z0-9\-\/\=]+$") Then
    '                        codeHeader = "商品名"
    '                        code = Regex.Split(DGV3.Item(dH3.IndexOf(codeHeader), r2).Value, "\/|\=")(0)
    '                        DGV1.Item(dH1.IndexOf("配送備考"), r1).Value = "[[" & DGV3.Item(dH3.IndexOf(codeHeader), r2).Value & "]]"
    '                    Else
    '                        code = DGV3.Item(dH3.IndexOf(codeHeader), r2).Value
    '                    End If
    '                Else
    '                    code = DGV3.Item(dH3.IndexOf(codeHeader), r2).Value
    '                End If
    '                str &= code & "*" & DGV3.Item(dH3.IndexOf("受注数"), r2).Value & "、"
    '            End If
    '        Next

    '        str = str.TrimEnd("、")
    '        If str <> "" Then
    '            DGV1.Item(dH1.IndexOf("商品マスタ"), r1).Value = str
    '            DGV1.Item(dH1.IndexOf("商品マスタ"), r1).Style.BackColor = Color.Empty
    '        Else
    '            DGV1.Item(dH1.IndexOf("商品マスタ"), r1).Style.BackColor = Color.Red
    '            'MsgBox("配送データと明細データが合っていない可能性があります")
    '            DGV1.CurrentCell = DGV1(dH1.IndexOf("商品マスタ"), r1)
    '            'Return False
    '            'Exit Function
    '            error_count = error_count + 1
    '        End If
    '    Next

    '    If error_count > 0 Then
    '        MsgBox("配送データと明細データが合っていない可能性があります")
    '    End If
    '    'Return True
    'End Sub

    Private Sub KryptonButton3_Click(sender As Object, e As EventArgs) Handles KryptonButton3.Click
        If DGV3.RowCount = 0 Or DGV4.RowCount = 0 Then
            MsgBox("処理してください。", MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        If DownloadCheck() <> True Then
            Exit Sub
        End If

        Dim errorCount As Integer = 0

        'フォルダ選択
        Dim dlg = New VistaFolderBrowserDialog With {
            .RootFolder = Environment.SpecialFolder.Desktop,
            .Description = "フォルダを選択してください"
        }
        Dim saveDir As String = ""
        If dlg.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
            saveDir = dlg.SelectedPath

            'ヘッダー行作成
            'Dim header As String = ""
            'For c As Integer = 0 To DGV1.ColumnCount - 1
            '    If CStr(header) = "" Then
            '        header = """" & DGV1.Columns(c).HeaderText & """"
            '    Else
            '        header &= "," & """" & DGV1.Columns(c).HeaderText & """"
            '    End If
            'Next

            'Dim str As String = header & vbCrLf
            'For r As Integer = 0 To DGV3.RowCount - 1
            '    If DGV3.Item(dH3.IndexOf("受注数"), r).Value = "同梱" Then
            '        str &= DGV3.Item(dH3.IndexOf("お届け先コード"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("お届け先郵便番号"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("お届け先住所１"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("お届け先住所２"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("お届け先住所３"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("お届け先名１"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("お届け先名２"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("お届け先電話"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("お届け先メールアドレス"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("代行ご依頼主郵便番号"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("代行ご依頼主住所１"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("代行ご依頼主住所２"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("代行ご依頼主名１"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("代行ご依頼主名２"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("代行ご依頼主電話"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("代行ご依頼主メールアドレス"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("佐川急便顧客コード"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("問い合せ№"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("発送日"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("個数"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("顧客管理番号"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("便種コード"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("配達指定日"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("時間帯コード"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("代引金額"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("代引消費税"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("保険金額"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("元着区分（未使用）"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("営止め区分"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("営止精算店コード"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("営止精算店ローカルコード"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("記事欄01"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("記事欄02"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("記事欄03"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("記事欄04"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("記事欄05"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("記事欄06"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("記事欄07"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("記事欄08"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("記事欄09"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("記事欄10"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("記事欄11"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("記事欄12"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("マスタコード"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("マスタ配送"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("マスタ個口"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("マスタ備考"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("マーク"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("処理用"), r).Value & ","
            '        str &= DGV3.Item(dH3.IndexOf("処理用2"), r).Value & ","
            '        str &= vbLf
            '    End If
            'Next

            Dim saveName As String = saveDir & "\" & Format(Now, "yyyyMMddHHmmss") & ".xlsx"
            'File.WriteAllText(saveName, str, ENC_SJ)

            Dim xWorkbook As IWorkbook = New XSSFWorkbook()

            '佐川
            Dim sheetSaga As ISheet = xWorkbook.CreateSheet("佐川(全)")
            Dim inRow As Integer = 0
            sheetSaga.CreateRow(inRow)
            For c As Integer = 1 To DGV4.ColumnCount - 1
                Dim v As String = DGV4.Columns(c).HeaderText
                If IsNumeric(v) Then
                    sheetSaga.GetRow(inRow).CreateCell(c - 1).SetCellValue(CDbl(v))
                Else
                    sheetSaga.GetRow(inRow).CreateCell(c - 1).SetCellValue(v)
                End If
            Next
            For r As Integer = 0 To DGV4.RowCount - 1
                inRow = inRow + 1
                sheetSaga.CreateRow(inRow)
                For c As Integer = 1 To DGV4.ColumnCount - 1
                    Dim v As String = DGV4.Item(c, r).Value
                    sheetSaga.GetRow(inRow).CreateCell(c - 1).SetCellValue(v)
                Next
            Next

            'test
            Dim count As Integer = 0
            For r As Integer = 0 To DGV4.RowCount - 1
                If DGV4.Item(dH4.IndexOf("受注数"), r).Value = "1" Then
                    count = count + 1
                End If
            Next
            If count > 0 Then
                Dim sheetSaga1 As ISheet = xWorkbook.CreateSheet("佐川①")
                inRow = 0
                sheetSaga1.CreateRow(inRow)
                For c As Integer = 1 To DGV4.ColumnCount - 1
                    Dim v As String = DGV4.Columns(c).HeaderText
                    If IsNumeric(v) Then
                        sheetSaga1.GetRow(inRow).CreateCell(c - 1).SetCellValue(CDbl(v))
                    Else
                        sheetSaga1.GetRow(inRow).CreateCell(c - 1).SetCellValue(v)
                    End If
                Next
                For r As Integer = 0 To DGV4.RowCount - 1
                    If DGV4.Item(dH4.IndexOf("受注数"), r).Value = "1" Then
                        inRow = inRow + 1
                        sheetSaga1.CreateRow(inRow)
                        For c As Integer = 1 To DGV4.ColumnCount - 1
                            Dim v As String = DGV4.Item(c, r).Value
                            sheetSaga1.GetRow(inRow).CreateCell(c - 1).SetCellValue(v)
                        Next
                    End If
                Next
            End If

            'test
            count = 0
            For r As Integer = 0 To DGV4.RowCount - 1
                If DGV4.Item(dH4.IndexOf("受注数"), r).Value = "2" Then
                    count = count + 1
                End If
            Next
            If count > 0 Then
                Dim sheetSaga2 As ISheet = xWorkbook.CreateSheet("佐川②")
                inRow = 0
                sheetSaga2.CreateRow(inRow)
                For c As Integer = 1 To DGV4.ColumnCount - 1
                    Dim v As String = DGV4.Columns(c).HeaderText
                    If IsNumeric(v) Then
                        sheetSaga2.GetRow(inRow).CreateCell(c - 1).SetCellValue(CDbl(v))
                    Else
                        sheetSaga2.GetRow(inRow).CreateCell(c - 1).SetCellValue(v)
                    End If
                Next
                For r As Integer = 0 To DGV4.RowCount - 1
                    If DGV4.Item(dH4.IndexOf("受注数"), r).Value = "2" Then
                        inRow = inRow + 1
                        sheetSaga2.CreateRow(inRow)
                        For c As Integer = 1 To DGV4.ColumnCount - 1
                            Dim v As String = DGV4.Item(c, r).Value
                            sheetSaga2.GetRow(inRow).CreateCell(c - 1).SetCellValue(v)
                        Next
                    End If
                Next
            End If

            'test
            count = 0
            For r As Integer = 0 To DGV4.RowCount - 1
                If DGV4.Item(dH4.IndexOf("受注数"), r).Value = "3" Then
                    count = count + 1
                End If
            Next
            If count > 0 Then
                Dim sheetSaga3 As ISheet = xWorkbook.CreateSheet("佐川③")
                inRow = 0
                sheetSaga3.CreateRow(inRow)
                For c As Integer = 1 To DGV4.ColumnCount - 1
                    Dim v As String = DGV4.Columns(c).HeaderText
                    If IsNumeric(v) Then
                        sheetSaga3.GetRow(inRow).CreateCell(c - 1).SetCellValue(CDbl(v))
                    Else
                        sheetSaga3.GetRow(inRow).CreateCell(c - 1).SetCellValue(v)
                    End If
                Next
                For r As Integer = 1 To DGV4.RowCount - 1
                    If DGV4.Item(dH4.IndexOf("受注数"), r).Value = "3" Then
                        inRow = inRow + 1
                        sheetSaga3.CreateRow(inRow)
                        For c As Integer = 1 To DGV4.ColumnCount - 1
                            Dim v As String = DGV4.Item(c, r).Value
                            sheetSaga3.GetRow(inRow).CreateCell(c - 1).SetCellValue(v)
                        Next
                    End If
                Next
            End If

            Dim sheetall As ISheet = xWorkbook.CreateSheet("元データ(全)")
            inRow = 0
            sheetall.CreateRow(inRow)
            For c As Integer = 0 To DGV1.ColumnCount - 1
                Dim v As String = DGV1.Columns(c).HeaderText
                'If IsNumeric(v) Then
                '    sheetall.GetRow(inRow).CreateCell(c).SetCellValue(CDbl(v))
                'Else
                sheetall.GetRow(inRow).CreateCell(c).SetCellValue(v)
                'End If
            Next
            For r As Integer = 0 To DGV3.RowCount - 1
                inRow = inRow + 1
                sheetall.CreateRow(inRow)
                If DGV3.Item(dH3.IndexOf("受注数"), r).Value = "" Then
                    errorCount = errorCount + 1
                End If
                For c As Integer = 1 To DGV3.ColumnCount - 1
                    Dim v As String = DGV3.Item(c, r).Value
                    If IsNumeric(v) Then
                        sheetall.GetRow(inRow).CreateCell(c - 1).SetCellValue(CDbl(v))
                    Else
                        sheetall.GetRow(inRow).CreateCell(c - 1).SetCellValue(v)
                    End If
                Next
            Next

            'test
            count = 0
            For r As Integer = 0 To DGV3.RowCount - 1
                If DGV3.Item(dH3.IndexOf("受注数"), r).Value <> "1" And DGV3.Item(dH3.IndexOf("受注数"), r).Value <> "2" And DGV3.Item(dH3.IndexOf("受注数"), r).Value <> "3" Then
                    count = count + 1
                End If
            Next
            If count > 0 Then
                Dim sheet999 As ISheet = xWorkbook.CreateSheet("元データ(外)")
                inRow = 0
                sheet999.CreateRow(inRow)
                For c As Integer = 0 To DGV1.ColumnCount - 1
                    Dim v As String = DGV1.Columns(c).HeaderText
                    If IsNumeric(v) Then
                        sheet999.GetRow(inRow).CreateCell(c).SetCellValue(CDbl(v))
                    Else
                        sheet999.GetRow(inRow).CreateCell(c).SetCellValue(v)
                    End If
                Next
                For r As Integer = 0 To DGV3.RowCount - 1
                    If DGV3.Item(dH3.IndexOf("受注数"), r).Value <> "1" And DGV3.Item(dH3.IndexOf("受注数"), r).Value <> "2" And DGV3.Item(dH3.IndexOf("受注数"), r).Value <> "3" Then
                        inRow = inRow + 1
                        sheet999.CreateRow(inRow)
                        For c As Integer = 1 To DGV3.ColumnCount - 1
                            Dim v As String = DGV3.Item(c, r).Value
                            'If IsNumeric(v) Then
                            '    sheet999.GetRow(inRow).CreateCell(c - 1).SetCellValue(CDbl(v))
                            'Else
                            sheet999.GetRow(inRow).CreateCell(c - 1).SetCellValue(v)
                            'End If
                        Next
                    End If
                Next
            End If

            'test
            count = 0
            For r As Integer = 0 To DGV3.RowCount - 1
                If DGV3.Item(dH3.IndexOf("受注数"), r).Value = "1" Then
                    count = count + 1
                End If
            Next
            If count > 0 Then
                inRow = 0
                Dim sheet1 As ISheet = xWorkbook.CreateSheet("元データ①")
                sheet1.CreateRow(inRow)
                For c As Integer = 0 To DGV1.ColumnCount - 1
                    Dim v As String = DGV1.Columns(c).HeaderText
                    If IsNumeric(v) Then
                        sheet1.GetRow(inRow).CreateCell(c).SetCellValue(CDbl(v))
                    Else
                        sheet1.GetRow(inRow).CreateCell(c).SetCellValue(v)
                    End If
                Next
                For r As Integer = 0 To DGV3.RowCount - 1
                    If DGV3.Item(dH3.IndexOf("受注数"), r).Value = "1" Then
                        inRow = inRow + 1
                        sheet1.CreateRow(inRow)
                        For c As Integer = 1 To DGV3.ColumnCount - 1
                            Dim v As String = DGV3.Item(c, r).Value
                            'If IsNumeric(v) Then
                            '    sheet1.GetRow(inRow).CreateCell(c - 1).SetCellValue(CDbl(v))
                            'Else
                            sheet1.GetRow(inRow).CreateCell(c - 1).SetCellValue(v)
                            'End If
                        Next
                    End If
                Next
            End If

            'test
            count = 0
            For r As Integer = 0 To DGV3.RowCount - 1
                If DGV3.Item(dH3.IndexOf("受注数"), r).Value = "2" Then
                    count = count + 1
                End If
            Next
            If count > 0 Then
                inRow = 0
                Dim sheet2 As ISheet = xWorkbook.CreateSheet("元データ②")
                sheet2.CreateRow(inRow)
                For c As Integer = 0 To DGV1.ColumnCount - 1
                    Dim v As String = DGV1.Columns(c).HeaderText
                    If IsNumeric(v) Then
                        sheet2.GetRow(inRow).CreateCell(c).SetCellValue(CDbl(v))
                    Else
                        sheet2.GetRow(inRow).CreateCell(c).SetCellValue(v)
                    End If
                Next
                For r As Integer = 0 To DGV3.RowCount - 1
                    If DGV3.Item(dH3.IndexOf("受注数"), r).Value = "2" Then
                        inRow = inRow + 1
                        sheet2.CreateRow(inRow)
                        For c As Integer = 1 To DGV3.ColumnCount - 1
                            Dim v As String = DGV3.Item(c, r).Value
                            'If IsNumeric(v) Then
                            '    sheet2.GetRow(inRow).CreateCell(c - 1).SetCellValue(CDbl(v))
                            'Else
                            sheet2.GetRow(inRow).CreateCell(c - 1).SetCellValue(v)
                            'End If
                        Next
                    End If
                Next
            End If

            'test
            count = 0
            For r As Integer = 0 To DGV3.RowCount - 1
                If DGV3.Item(dH3.IndexOf("受注数"), r).Value = "3" Then
                    count = count + 1
                End If
            Next
            If count > 0 Then
                inRow = 0
                Dim sheet3 As ISheet = xWorkbook.CreateSheet("元データ③")
                sheet3.CreateRow(inRow)
                For c As Integer = 0 To DGV1.ColumnCount - 1
                    Dim v As String = DGV1.Columns(c).HeaderText
                    If IsNumeric(v) Then
                        sheet3.GetRow(inRow).CreateCell(c).SetCellValue(CDbl(v))
                    Else
                        sheet3.GetRow(inRow).CreateCell(c).SetCellValue(v)
                    End If
                Next
                For r As Integer = 0 To DGV3.RowCount - 1
                    If DGV3.Item(dH3.IndexOf("受注数"), r).Value = "3" Then
                        inRow = inRow + 1
                        sheet3.CreateRow(inRow)
                        For c As Integer = 1 To DGV3.ColumnCount - 1
                            Dim v As String = DGV3.Item(c, r).Value
                            'If IsNumeric(v) Then
                            '    sheet3.GetRow(inRow).CreateCell(c - 1).SetCellValue(CDbl(v))
                            'Else
                            sheet3.GetRow(inRow).CreateCell(c - 1).SetCellValue(v)
                            'End If
                        Next
                    End If
                Next
            End If

            'test
            count = 0
            For r As Integer = 0 To DGV3.RowCount - 1
                If DGV3.Item(dH3.IndexOf("受注数"), r).Value = "" Then
                    count = count + 1
                End If
            Next
            If count > 0 Then
                inRow = 0
                Dim sheetError As ISheet = xWorkbook.CreateSheet("元データ(エラー)")
                sheetError.CreateRow(inRow)
                For c As Integer = 0 To DGV1.ColumnCount - 1
                    Dim v As String = DGV1.Columns(c).HeaderText
                    If IsNumeric(v) Then
                        sheetError.GetRow(inRow).CreateCell(c).SetCellValue(CDbl(v))
                    Else
                        sheetError.GetRow(inRow).CreateCell(c).SetCellValue(v)
                    End If
                Next
                For r As Integer = 0 To DGV3.RowCount - 1
                    If DGV3.Item(dH3.IndexOf("受注数"), r).Value = "" Then
                        inRow = inRow + 1
                        sheetError.CreateRow(inRow)
                        For c As Integer = 1 To DGV3.ColumnCount - 1
                            Dim v As String = DGV3.Item(c, r).Value
                            'If IsNumeric(v) Then
                            'sheetError.GetRow(inRow).CreateCell(c - 1).SetCellValue(CDbl(v))
                            'Else
                            sheetError.GetRow(inRow).CreateCell(c - 1).SetCellValue(v)
                            'End If
                        Next
                    End If
                Next
            End If

            inRow = 0

            Using wfs = File.Create(saveName)
                xWorkbook.Write(wfs)
            End Using
            xWorkbook = Nothing
            MsgBox(saveName & vbCrLf & "保存しました", MsgBoxStyle.SystemModal)
            If errorCount > 0 Then
                MsgBox("受注数ないデータがあります。", MsgBoxStyle.SystemModal)
            End If
        Else
            Exit Sub
        End If
    End Sub

    Private Function DownloadCheck()
        Dim flag As Boolean = True
        'Dim indexes As Integer() = System.Globalization.StringInfo.ParseCombiningCharacters(DGV4.Item(dH4.IndexOf("代行住所2"), r).Value)
        'Dim len As Integer = indexes.Length
        For r As Integer = 0 To DGV4.RowCount - 1
            If DGV4.Item(dH4.IndexOf("代行名称1"), r).Value <> "" Then
                Dim len As Integer = DGV4.Item(dH4.IndexOf("代行名称1"), r).Value.Length
                If len > 16 Then
                    DGV4.Item(dH4.IndexOf("代行名称1"), r).Style.BackColor = Color.Red
                    DGV4.CurrentCell = DGV4(dH4.IndexOf("代行名称1"), r)
                    flag = False
                End If
            End If

            '16
            If DGV4.Item(dH4.IndexOf("代行住所2"), r).Value <> "" Then

                Dim len As Integer = DGV4.Item(dH4.IndexOf("代行住所2"), r).Value.Length
                If len > 16 Then
                    DGV4.Item(dH4.IndexOf("代行住所2"), r).Style.BackColor = Color.Red
                    DGV4.CurrentCell = DGV4(dH4.IndexOf("代行住所2"), r)
                    flag = False
                End If
            End If

            'If DGV4.Item(dH4.IndexOf("住所1"), r).Value <> "" Then
            '    Dim len As Integer = DGV4.Item(dH4.IndexOf("住所1"), r).Value.Length
            '    If len > 16 Then
            '        DGV4.Item(dH4.IndexOf("住所1"), r).Style.BackColor = Color.DarkOrange
            '        DGV4.CurrentCell = DGV4(dH4.IndexOf("住所1"), r)
            '        flag = False
            '    End If
            'End If

            'If DGV4.Item(dH4.IndexOf("住所2"), r).Value <> "" Then
            '    Dim len As Integer = DGV4.Item(dH4.IndexOf("住所2"), r).Value.Length
            '    If len > 16 Then
            '        DGV4.Item(dH4.IndexOf("住所2"), r).Style.BackColor = Color.Yellow
            '        DGV4.CurrentCell = DGV4(dH4.IndexOf("住所2"), r)
            '        flag = False
            '    End If
            'End If

            'If DGV4.Item(dH4.IndexOf("住所3"), r).Value <> "" Then
            '    Dim len As Integer = DGV4.Item(dH4.IndexOf("住所3"), r).Value.Length
            '    If len > 16 Then
            '        DGV4.Item(dH4.IndexOf("住所3"), r).Style.BackColor = Color.Yellow
            '        DGV4.CurrentCell = DGV4(dH4.IndexOf("住所3"), r)
            '        flag = False
            '    End If
            'End If

            'If DGV4.Item(dH4.IndexOf("名称1"), r).Value <> "" Then
            '    Dim len As Integer = DGV4.Item(dH4.IndexOf("名称1"), r).Value.Length
            '    If len > 16 Then
            '        DGV4.Item(dH4.IndexOf("名称1"), r).Style.BackColor = Color.Yellow
            '        DGV4.CurrentCell = DGV4(dH4.IndexOf("名称1"), r)
            '        flag = False
            '    End If
            'End If
        Next
        TabControl2.SelectedIndex = 0
        Return flag
    End Function

    Dim BeforeColor As Color() = New Color() {Color.Empty, Color.Empty, Color.Empty, Color.Empty, Color.Empty}
    Private Sub DGV1_SelectionChanged(sender As Object, e As EventArgs) Handles DGV1.SelectionChanged
        If sender.Rows.Count = 0 Then
            Exit Sub
        End If

        Dim dHSender As ArrayList = TM_HEADER_GET(sender)
        Dim dgv As DataGridView

        Dim num As Integer = 0
        Select Case True
            Case sender Is DGV1
                'num = 0
                dgv = DGV1
                AzukiControl1.Text = dgv.SelectedCells(0).Value
        End Select

        '選択行に色を付ける
        For r As Integer = 0 To sender.RowCount - 1
            If sender.Rows(r).DefaultCellStyle.BackColor <> BeforeColor(num) Then
                sender.Rows(r).DefaultCellStyle.BackColor = BeforeColor(num)
            End If
        Next
        Dim selRow As ArrayList = New ArrayList
        For i As Integer = 0 To sender.SelectedCells.Count - 1
            If Not selRow.Contains(sender.SelectedCells(i).RowIndex) Then
                selRow.Add(sender.SelectedCells(i).RowIndex)
            End If
        Next
        For i As Integer = 0 To selRow.Count - 1
            BeforeColor(num) = sender.Rows(selRow(i)).DefaultCellStyle.BackColor
            sender.Rows(selRow(i)).DefaultCellStyle.BackColor = Color.LightCyan
        Next


        'ヘッダー取得
        'Dim dH3 As ArrayList = TM_HEADER_GET(DGV3)

        'If sender.RowCount > 0 Then
        '    If sender.SelectedCells.Count > 0 And DGV3.RowCount > 0 Then
        '        Dim dNum As String = sender.Item(dHSender.IndexOf("伝票番号"), sender.SelectedCells(0).RowIndex).Value
        '        If IsNumeric(dNum) Then '伝票番号
        '            For r As Integer = 0 To DGV3.RowCount - 1
        '                If DGV3.Item(dH3.IndexOf("伝票番号"), r).Value = dNum Then
        '                    DGV3.Item(dH3.IndexOf("伝票番号"), r).Selected = True
        '                    DGV3.CurrentCell = DGV3(dH3.IndexOf("伝票番号"), r)
        '                End If
        '            Next
        '        End If
        '    End If
        'End If
    End Sub

    Private Sub DGV3_SelectionChanged(sender As Object, e As EventArgs) Handles DGV3.SelectionChanged
        If sender.Rows.Count = 0 Then
            Exit Sub
        End If

        Dim dHSender As ArrayList = TM_HEADER_GET(sender)
        Dim dgv As DataGridView

        Dim num As Integer = 0
        Select Case True
            Case sender Is DGV3
                'num = 0
                dgv = DGV3
                AzukiControl3.Text = dgv.SelectedCells(0).Value
        End Select

        '選択行に色を付ける
        For r As Integer = 0 To sender.RowCount - 1
            If sender.Rows(r).DefaultCellStyle.BackColor <> BeforeColor(num) Then
                sender.Rows(r).DefaultCellStyle.BackColor = BeforeColor(num)
            End If
        Next
        Dim selRow As ArrayList = New ArrayList
        For i As Integer = 0 To sender.SelectedCells.Count - 1
            If Not selRow.Contains(sender.SelectedCells(i).RowIndex) Then
                selRow.Add(sender.SelectedCells(i).RowIndex)
            End If
        Next
        For i As Integer = 0 To selRow.Count - 1
            BeforeColor(num) = sender.Rows(selRow(i)).DefaultCellStyle.BackColor
            sender.Rows(selRow(i)).DefaultCellStyle.BackColor = Color.LightCyan
        Next
    End Sub

    Private Sub DGV4_SelectionChanged(sender As Object, e As EventArgs) Handles DGV4.SelectionChanged
        If sender.Rows.Count = 0 Then
            Exit Sub
        End If

        Dim dHSender As ArrayList = TM_HEADER_GET(sender)
        Dim dgv As DataGridView

        Dim num As Integer = 0
        Select Case True
            Case sender Is DGV4
                'num = 0
                dgv = DGV4
                AzukiControl4.Text = dgv.SelectedCells(0).Value
        End Select

        '選択行に色を付ける
        For r As Integer = 0 To sender.RowCount - 1
            If sender.Rows(r).DefaultCellStyle.BackColor <> BeforeColor(num) Then
                sender.Rows(r).DefaultCellStyle.BackColor = BeforeColor(num)
            End If
        Next
        Dim selRow As ArrayList = New ArrayList
        For i As Integer = 0 To sender.SelectedCells.Count - 1
            If Not selRow.Contains(sender.SelectedCells(i).RowIndex) Then
                selRow.Add(sender.SelectedCells(i).RowIndex)
            End If
        Next
        For i As Integer = 0 To selRow.Count - 1
            BeforeColor(num) = sender.Rows(selRow(i)).DefaultCellStyle.BackColor
            sender.Rows(selRow(i)).DefaultCellStyle.BackColor = Color.LightCyan
        Next
    End Sub

    Private Sub AzukiControl1_TextChanged(sender As Object, e As EventArgs) Handles AzukiControl1.TextChanged
        If DGV1.SelectedCells.Count > 0 Then
            DGV1.SelectedCells(0).Value = AzukiControl1.Text
        End If
    End Sub

    Private Sub AzukiControl4_TextChanged(sender As Object, e As EventArgs) Handles AzukiControl4.TextChanged
        If DGV4.SelectedCells.Count > 0 Then
            DGV4.SelectedCells(0).Value = AzukiControl4.Text
        End If
    End Sub

    Private Sub AzukiControl3_TextChanged(sender As Object, e As EventArgs) Handles AzukiControl3.TextChanged
        If DGV3.SelectedCells.Count > 0 Then
            DGV3.SelectedCells(0).Value = AzukiControl3.Text
        End If
    End Sub

    Private Sub 切り取りToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 切り取りToolStripMenuItem1.Click
        Dim dgv As DataGridView = DGV1
        Dim selCell = dgv.SelectedCells
        CUTS(dgv)
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        Dim dgv As DataGridView = DGV3
        Dim selCell = dgv.SelectedCells
        CUTS(dgv)
    End Sub

    Private Sub ToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem4.Click
        Dim dgv As DataGridView = DGV4
        Dim selCell = dgv.SelectedCells
        CUTS(dgv)
    End Sub

    Private Sub コピーToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles コピーToolStripMenuItem1.Click
        Dim dgv As DataGridView = DGV1
        Dim selCell = dgv.SelectedCells
        COPYS(dgv, 0)
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        Dim dgv As DataGridView = DGV3
        Dim selCell = dgv.SelectedCells
        COPYS(dgv, 0)
    End Sub

    Private Sub ToolStripMenuItem5_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem5.Click
        Dim dgv As DataGridView = DGV4
        Dim selCell = dgv.SelectedCells
        COPYS(dgv, 0)
    End Sub

    Private Sub 貼り付けToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 貼り付けToolStripMenuItem1.Click
        Dim dgv As DataGridView = DGV1
        Dim selCell = dgv.SelectedCells
        PASTES(dgv, selCell)
    End Sub

    Private Sub ToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem3.Click
        Dim dgv As DataGridView = DGV3
        Dim selCell = dgv.SelectedCells
        PASTES(dgv, selCell)
    End Sub

    Private Sub ToolStripMenuItem6_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem6.Click
        Dim dgv As DataGridView = DGV4
        Dim selCell = dgv.SelectedCells
        PASTES(dgv, selCell)
    End Sub
End Class