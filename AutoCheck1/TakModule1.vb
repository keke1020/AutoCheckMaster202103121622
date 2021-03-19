Imports System.IO
Imports System.Net
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Xml
Imports Hnx8
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel

'================================================
'------------------------------------------------
'
'汎用プログラムモジュール
'
'------------------------------------------------
'================================================

Module TakModule1
    Public Const Delimiter_comma As String = ","
    Public Const Delimiter_csvREAD As String = "|=|"

    Public encSJ As Encoding = Encoding.GetEncoding("SHIFT-JIS")
    Public enc64 As Encoding = Encoding.UTF8

    Public appPath As String = System.Reflection.Assembly.GetExecutingAssembly().Location
    Public appPathDir As String = Path.GetDirectoryName(appPath)
    Public desktopPath As String = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)


    '********************************************
    ''' <summary>ヘッダー取得</summary>
    ''' <param name="dgv">データグリッドビュー</param>
    Public Function TM_HEADER_GET(dgv As DataGridView) As ArrayList
        Dim DGVheaderCheck As ArrayList = New ArrayList
        For c As Integer = 0 To dgv.Columns.Count - 1
            DGVheaderCheck.Add(dgv.Columns(c).HeaderText)
        Next c
        Return DGVheaderCheck
    End Function

    '********************************************
    ''' <summary>1行目ヘッダー取得</summary>
    ''' <param name="dgv">データグリッドビュー</param>
    Public Function TM_HEADER_1ROW_GET(dgv As DataGridView, Optional row As Integer = 0) As ArrayList
        Dim DGVheaderCheck As ArrayList = New ArrayList
        For c As Integer = 0 To dgv.Columns.Count - 1
            DGVheaderCheck.Add(dgv.Item(c, row).Value)
        Next c
        Return DGVheaderCheck
    End Function

    '********************************************
    Public ALIO_ERROR As String = ""
    ''' <summary>
    ''' ArrayのIndexof検索で例外しか出ないので、内容を取得する
    ''' </summary>
    ''' <param name="arrays">配列でもArraylistでも可</param>
    ''' <param name="str">検索文字</param>
    ''' <returns></returns>
    Public Function TM_ArIndexof(arrays As Object, str As String) As Integer
        Dim res As Integer = 0
        If arrays.GetType.IsArray Then
            Dim ar As String() = arrays
            If ar.Contains(str) Then
                res = Array.IndexOf(ar, str)
            Else
                res = -1
            End If
        Else
            Dim ar As ArrayList = arrays
            If ar.Contains(str) Then
                res = ar.IndexOf(str)
            Else
                res = -1
            End If
        End If
        If res = -1 Then
            ALIO_ERROR = "「" & str & "」が処理するリストに存在しません" & vbCrLf & "Error:" & arrays.name.ToString
        End If
        Return res
    End Function

    '********************************************
    ''' <summary>
    ''' 楽天API読み書き
    ''' </summary>
    Public Class APIlist
        Public Shared _url As String = ""
        Public Shared _postget As String = ""
        Public Shared _getStr As String = ""
        Public Shared _postKeys As String() = Nothing
        Public Shared _postItem As String() = Nothing
    End Class

    '********************************************
    ''' <summary>
    ''' 楽天API実行
    ''' </summary>
    ''' <param name="authkey"></param>
    ''' <returns></returns>
    Public Function API_RUN(ByVal authkey As String) As String
        Dim url As String = APIlist._url
        Dim postget As String = ""
        Dim postDataBytes As Byte() = Nothing

        If APIlist._postget = "GET" Then
            postget = "GET"
            url &= "?" & APIlist._getStr
            'url &= " HTTP/1.1"
        Else
            postget = "POST"
            Dim stream As Stream = New MemoryStream
            Dim xmlWtr As XmlWriter = XmlWriter.Create(stream)
            xmlWtr.WriteStartDocument()
            'POSTするXMLを作成
            Dim keys As String() = APIlist._postKeys
            Dim itemkeys As String() = APIlist._postItem
            For i As Integer = 0 To keys.Length - 1
                xmlWtr.WriteStartElement(keys(i))
                If i = keys.Length - 1 Then
                    For k As Integer = 0 To itemkeys.Length - 1
                        Dim ik As String() = Split(itemkeys(k), Delimiter_comma)
                        xmlWtr.WriteStartElement(ik(0))
                        xmlWtr.WriteString(ik(1))
                        xmlWtr.WriteEndElement()
                    Next
                End If
            Next
            For i As Integer = 0 To keys.Length - 1
                xmlWtr.WriteEndElement()
            Next
            xmlWtr.Flush()
            xmlWtr.Close()
            stream.Position = 0
            Dim data As String = New StreamReader(stream).ReadToEnd()
            stream.Close()
            postDataBytes = Encoding.ASCII.GetBytes(data)
        End If

        Dim req As HttpWebRequest = CType(WebRequest.Create(url), HttpWebRequest)
        req.Method = postget
        req.Headers.Add("Authorization", authkey)
        req.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/29.0.1547.76 Safari/537.36"
        req.ContentType = "application/xml"
        If APIlist._postget = "post" Then
            req.ContentLength = postDataBytes.Length
            Dim reqStream As Stream = req.GetRequestStream()
            reqStream.Write(postDataBytes, 0, postDataBytes.Length)
            reqStream.Close()
        End If

        Dim res As WebResponse = req.GetResponse()
        Dim resStream As Stream = res.GetResponseStream()
        Dim sr As StreamReader = New StreamReader(resStream, enc64)
        Dim html As String = sr.ReadToEnd()
        sr.Close()
        resStream.Close()

        'File.WriteAllText(desktopPath & "\test.txt", html)
        'MsgBox(html)

        Return html
    End Function

    '********************************************
    Dim resArray As ArrayList = New ArrayList
    '********************************************
    ''' <summary>
    ''' XMLパース
    ''' </summary>
    ''' <param name="html">テキスト</param>
    ''' <param name="pnodeName">対象ノード</param>
    ''' <returns>ArrayList</returns>
    Public Function XML_READ(ByVal html As String, ByVal pnodeName As String) As ArrayList
        Dim xmlDoc As New XmlDocument()
        xmlDoc.LoadXml(html)

        Dim str As String = ""

        'XMLを再帰処理を使用してパース
        Dim nodes As XmlNodeList = xmlDoc.ChildNodes
        For i As Integer = 0 To nodes.Count - 1
            If nodes(i).NodeType = XmlNodeType.Element Then
                GetRecursiveXML(nodes(i), "")
            End If
        Next

        Return resArray
    End Function

    '********************************************
    ''' <summary>
    ''' XML再帰処理
    ''' </summary>
    ''' <param name="node"></param>
    ''' <param name="mark"></param>
    Private Sub GetRecursiveXML(ByVal node As XmlNode, ByVal mark As String)
        Dim cNodes As XmlNodeList = node.ChildNodes
        For i As Integer = 0 To cNodes.Count - 1
            If cNodes(i).HasChildNodes Then
                Dim flag As Boolean = False
                For Each ccNode As XmlNode In cNodes(i).ChildNodes
                    If ccNode.NodeType = XmlNodeType.Element Then
                        flag = True
                        Exit For
                    End If
                Next
                Dim mark2 As String = mark & ">"
                If flag = False Then
                    resArray.Add(mark2 & cNodes(i).Name & Delimiter_comma & cNodes(i).InnerText)
                Else
                    resArray.Add(mark2 & cNodes(i).Name & Delimiter_comma)
                    GetRecursiveXML(cNodes(i), mark2)
                End If
            End If
        Next
    End Sub

    '********************************************
    '--------------------------------------------
    ' XML作成
    ' keyArray.add("root")
    ' keyArray.add("item,xxx")
    ' => <root><item>xxx</item></root>
    '--------------------------------------------
    '********************************************
    Public Function XML_CREATE(ByVal keyArray As ArrayList) As String
        Dim stream As Stream = New MemoryStream
        Dim st As XmlWriterSettings = New XmlWriterSettings With {
            .Indent = True,
            .IndentChars = vbTab
        }
        Dim xmlWtr As XmlWriter = XmlWriter.Create(stream, st)

        xmlWtr.WriteStartDocument()

        Dim keyStart As ArrayList = New ArrayList
        For i As Integer = 0 To keyArray.Count - 1
            Dim kStr As String() = Split(keyArray(i), Delimiter_comma)
            If kStr.Length = 1 Then
                For j As Integer = 0 To keyStart.Count - 1
                    If keyStart.Contains(kStr(0)) Then
                        xmlWtr.WriteEndElement()
                        keyStart.RemoveAt(keyStart.Count - 1)
                    Else
                        Exit For
                    End If
                Next
                xmlWtr.WriteStartElement(kStr(0))
                keyStart.Add(kStr(0))
            Else
                xmlWtr.WriteStartElement(kStr(0))
                xmlWtr.WriteString(kStr(1))
                xmlWtr.WriteEndElement()
            End If
        Next

        For i As Integer = 0 To keyStart.Count - 1
            xmlWtr.WriteEndElement()
        Next

        xmlWtr.Flush()
        xmlWtr.Close()

        stream.Position = 0
        Dim res As String = New StreamReader(stream).ReadToEnd()
        stream.Close()

        Return res
    End Function

    '********************************************
    '--------------------------------------------
    ' XML読み出し（要素指定）
    ' str = テキスト
    ' node = 対象ノード
    ' => ArrayList
    '--------------------------------------------
    '********************************************
    Public Function XML_READ_A(ByVal str As String, ByVal node As String) As ArrayList
        Dim xmlDoc As New XmlDocument
        xmlDoc.LoadXml(str)
        Dim nodelist As XmlNodeList = xmlDoc.GetElementsByTagName(node)
        Dim resArray As ArrayList = New ArrayList
        For Each nodes As XmlNode In nodelist
            resArray.Add(nodes.InnerText)
        Next

        Return resArray
    End Function

    '********************************************
    '--------------------------------------------
    ' XML変更
    ' str = テキスト
    ' num = ノードリストの順番号
    ' node = 対象ノード
    ' newStr = 変更文字列
    ' => string
    '--------------------------------------------
    '********************************************
    Public Function XML_REPLACE(ByVal str As String, ByVal num As Integer, ByVal node As String, ByVal newStr As String) As String
        Dim xmlDoc As New XmlDocument
        xmlDoc.LoadXml(str)
        Dim nodelist As XmlNodeList = xmlDoc.GetElementsByTagName(node)
        If nodelist.Count > 0 Then
            nodelist(num).InnerText = newStr
        End If

        Return str
    End Function

    '********************************************
    '--------------------------------------------
    ' 64Bitエンコード・デコードclass
    '--------------------------------------------
    '********************************************
    Public Class MyBase64str
        Private enc As Encoding
        Public Sub New(ByVal encStr As String)
            enc = Encoding.GetEncoding(encStr)
        End Sub
        Public Function Encode(ByVal str As String) As String
            Return Convert.ToBase64String(enc.GetBytes(str))
        End Function
        Public Function Decode(ByVal str As String) As String
            Return enc.GetString(Convert.FromBase64String(str))
        End Function
    End Class

    '********************************************
    '--------------------------------------------
    ' csv読込
    '--------------------------------------------
    '********************************************
    ''' <summary>
    ''' CSV読込（汎用版）
    ''' </summary>
    ''' <param name="path">読み込むファイルのパス</param>
    ''' <returns>戻り値：ArrayList,エンコード名(string)</returns>
    Public Function TM_CSV_READ(ByRef path As String) As Object()
        Dim csvText As String = ""
        Dim fileEnc As String = ""
        Try
            Dim file As FileInfo = New FileInfo(path)
            Using reader As ReadJEnc.FileReader = New ReadJEnc.FileReader(file)
                Dim c As ReadJEnc.CharCode = reader.Read(file)
                fileEnc = c.Name
                csvText = reader.Text
            End Using
            'csvtiiaosText = File.ReadAllText(path, System.Text.Encoding.Default)
        Catch ex As Exception
            MsgBox("ファイルを開けません。" & vbNewLine & "元ファイルをエクセル等で開いていませんか？", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation Or MsgBoxStyle.SystemModal)
            Return Nothing
            Exit Function
        End Try

        '前後の改行を削除しておく
        csvText = csvText.Trim(New Char() {ControlChars.Cr, ControlChars.Lf})

        'カンマ、タブ等の区切り文字認識
        Dim Splitter As String = ","
        Dim header As String() = Regex.Split(csvText, vbCrLf & "|" & vbCr & "|" & vbLf)
        'Dim header As String() = Split(csvText, vbCrLf)
        If InStr(header(0), vbTab) Then
            Splitter = vbTab
        End If

        Dim csvRecords As New ArrayList
        Dim csvFields As New ArrayList

        Dim csvTextLength As Integer = csvText.Length
        Dim startPos As Integer = 0
        Dim endPos As Integer = 0
        Dim field As String = ""

        While True
            '空白を飛ばす
            'While startPos < csvTextLength _
            '    AndAlso (csvText.Chars(startPos) = " "c OrElse csvText.Chars(startPos) = ControlChars.Tab)
            '    startPos += 1
            'End While

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
                    csvText.Chars(endPos) <> Splitter AndAlso
                    csvText.Chars(endPos) <> ControlChars.Lf
                    endPos += 1
                End While
            Else
                '"で囲まれていない
                'カンマか改行の位置
                endPos = startPos
                While endPos < csvTextLength AndAlso
                    csvText.Chars(endPos) <> Splitter AndAlso
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
                csvFields = New ArrayList(csvFields.Count)

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

        Return New Object() {csvRecords, fileEnc}
    End Function

    Public Function TM_CSV_READ2(ByRef path As String)
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

    '********************************************
    '--------------------------------------------
    ' excel読込
    '--------------------------------------------
    '********************************************
    Public Function TM_XLS_READ(path As String, Optional sheetName As String = "")
        Dim xlsRecords As New ArrayList

        'Try
        '    Dim wb As New C1XLBook With {
        '    .CompatibilityMode = CompatibilityMode.NoLimits
        '}
        '    Dim sheet1 As XLSheet = Nothing
        '    wb.Load(path)
        '    If sheetName = "" Then
        '        sheet1 = wb.Sheets(0)
        '    Else
        '        sheet1 = wb.Sheets(sheetName)
        '    End If

        '    Dim charsToTrim() As Char = {"|"c, "="c}
        '    Dim row As Integer, col As Integer
        '    For row = 0 To sheet1.Rows.Count - 1
        '        Dim data As String = ""
        '        For col = 0 To sheet1.Columns.Count - 1
        '            data &= sheet1(row, col).Value & "|=|"
        '        Next col
        '        data = data.TrimEnd(charsToTrim)
        '        xlsRecords.Add(data)
        '    Next row
        'Catch ex As Exception

        'xlsxに対応するために、NPOIで読み込み
        Dim wb As IWorkbook = WorkbookFactory.Create(path)
        Dim ws As ISheet = wb.GetSheetAt(wb.GetSheetIndex(sheetName))
        Dim iLastRow As Integer = ws.LastRowNum   'シートの最終行数取得

        '最終行まで読込み
        Dim charsToTrim() As Char = {"|"c, "="c}
        Dim LastColumnsNum As Integer = 0
        For iCount As Integer = 0 To iLastRow
            If LastColumnsNum < ws.GetRow(iCount).LastCellNum Then
                LastColumnsNum = ws.GetRow(iCount).LastCellNum
            End If
        Next
        For iCount As Integer = 0 To iLastRow
            Dim getRow As IRow = ws.GetRow(iCount) '行取得
            Dim data As String = ""
            For c As Integer = 0 To LastColumnsNum - 1
                If getRow.GetCell(c) IsNot Nothing Then
                    data &= getRow.GetCell(c).ToString & "|=|"
                Else
                    data &= "|=|"
                End If
            Next
            data = data.TrimEnd(charsToTrim)
            xlsRecords.Add(data)
        Next

        'End Try

        Return xlsRecords
    End Function

    '********************************************
    '--------------------------------------------
    ' listbox（Process）表示
    '--------------------------------------------
    '********************************************
    Public Sub LIST_VIEW(ByVal LB As ListBox, ByRef mes As String, ByRef se As String)
        Dim seStr As String = ""
        Select Case se
            Case "start"
                seStr = "START"
                LB.Items.Add(mes & ">>" & seStr)
            Case "end"
                seStr = "END"
                LB.Items.Add("..." & seStr & ":" & mes)
            Case "error"
                seStr = "ERROR"
                LB.Items.Add("[" & seStr & "]:" & mes)
            Case Else
                seStr = se
                LB.Items.Add("  =>" & mes & "/" & seStr)
        End Select

        LB.SelectedIndex = LB.Items.Count - 1
        Application.DoEvents()
    End Sub

    '********************************************
    '--------------------------------------------
    ' FormのBringtoFrontと最小化状態復帰
    '--------------------------------------------
    '********************************************
    Public Sub FormFront(ByVal f As Form)
        f.Show()
        If f.WindowState = FormWindowState.Minimized Then
            f.WindowState = FormWindowState.Normal
        End If
        f.BringToFront()
    End Sub

    '********************************************
    '--------------------------------------------
    ' アプリケーションに関連づけられたアイコンを取得するためのAPI関数
    '--------------------------------------------
    '********************************************
    Public Function Icon_GET(ByVal icoPath As String) As Icon
        Dim strExeFilePath As String = ""
        'fileNameに関連付けられた実行ファイルのパスを取得する
        'Dim a = FindExecutable(icoPath, Nothing, exePath)
        If FindExecutable(icoPath, 0, strExeFilePath) <= 32 Then
            Return Nothing
        End If
        FindExecutable(icoPath, 0, strExeFilePath)
        'Dim ico As Icon = Icon.ExtractAssociatedIcon("C:\Program Files\Microsoft Office\Office14\EXCEL.EXE")
        Dim ico As Icon = Nothing 'Icon.ExtractAssociatedIcon("EXCEL.EXE")
        Return ico
    End Function

    Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long

    '********************************************
    '--------------------------------------------
    ' 機種依存文字チェック
    '--------------------------------------------
    '********************************************
    Public Function CheckIzonMoji(s) As Boolean
        Dim i As Integer
        If IsNumeric(s) Then
            CheckIzonMoji = False
            Exit Function
        End If
        For i = 1 To Len(s)
            Select Case Asc(Mid(s, i, 1))
                Case &H8540 To &H889E, &HEB40 To &HEFFC, &HF040 To &HFFFF, &HA0 To &HFF
                    CheckIzonMoji = True
                    Exit Function
            End Select
        Next
        CheckIzonMoji = False
    End Function

    '********************************************
    '--------------------------------------------
    ' CSV禁則文字チェック
    '--------------------------------------------
    '********************************************
    ''' <summary>
    ''' CSV禁則文字チェック
    ''' </summary>
    ''' <param name="s">object</param>
    ''' <returns>string</returns>
    Public Function CheckCSVkinsoku(s) As String
        s = Replace(s, """", "")
        s = Replace(s, "'", "")
        s = Replace(s, "|", "｜")
        s = Replace(s, ",", "，")
        s = Replace(s, "?", "？")
        CheckCSVkinsoku = s
    End Function

    '********************************************
    '--------------------------------------------
    ' TabColtrol表示切替
    '--------------------------------------------
    '********************************************
    Public Sub SetTabVisible(ByVal oTabControl As TabControl, ByVal tText As String)
        Dim nameArray As ArrayList = New ArrayList
        For i As Integer = 0 To oTabControl.TabPages.Count - 1
            nameArray.Add(oTabControl.TabPages(i).Text)
        Next

        For i As Integer = 0 To nameArray.Count - 1
            If Regex.IsMatch(nameArray(i), tText) Then
                SetTabVisible2(oTabControl, i, True)
            Else
                SetTabVisible2(oTabControl, i, False)
            End If
        Next
    End Sub

    Public Sub SetTabVisible2(ByVal oTabControl As TabControl, ByVal nIndex As Integer, ByVal bVisible As Boolean)
        Dim oAllTabPages As List(Of KeyValuePair(Of TabPage, Boolean))
        If oTabControl.Tag Is Nothing Then
            ' 全タブと、その表示状態を保持
            oAllTabPages = New List(Of KeyValuePair(Of TabPage, Boolean))
            For Each oTabPage As TabPage In oTabControl.TabPages
                Dim oTabAndVisible As New KeyValuePair(Of TabPage, Boolean)(oTabPage, True)
                oAllTabPages.Add(oTabAndVisible)
            Next
            oTabControl.Tag = oAllTabPages
        Else
            ' 全タブと、その表示状態を取得
            oAllTabPages = CType(oTabControl.Tag, List(Of KeyValuePair(Of TabPage, Boolean)))
        End If

        ' タブの表示状態を設定
        oAllTabPages(nIndex) = New KeyValuePair(Of TabPage, Boolean)(oAllTabPages(nIndex).Key, bVisible)

        oTabControl.TabPages.Clear()
        For Each oTabAndVisible As KeyValuePair(Of TabPage, Boolean) In oAllTabPages
            If oTabAndVisible.Value = True Then
                oTabControl.TabPages.Add(oTabAndVisible.Key)
            End If
        Next
    End Sub

    '********************************************
    '--------------------------------------------
    ' 名前を付けて保存（CSV）
    '--------------------------------------------
    '********************************************
    Public Sub NameChangeSave(mode As Integer, defaultPath As String, dgv As DataGridView)
        Dim sfd As New SaveFileDialog With {
            .Filter = "CSVファイル(*.csv)|*.csv|すべてのファイル(*.*)|*.*",
                    .AutoUpgradeEnabled = False,
            .FilterIndex = 0,
            .Title = "保存先のファイルを選択してください",
            .RestoreDirectory = True,
            .OverwritePrompt = True,
            .CheckPathExists = True
        }
        '.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),

        If InStr(defaultPath, "\") > 0 Then
            Dim sPath As String = Path.GetFileNameWithoutExtension(defaultPath)
            sfd.FileName = sPath & ".csv"
        Else
            sfd.FileName = "新しいファイル.csv"
        End If

        'ダイアログを表示する
        If sfd.ShowDialog() = DialogResult.OK Then
            SaveCsv(sfd.FileName, mode, dgv)
            MsgBox(sfd.FileName & vbCrLf & "保存しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    ''' <summary>
    ''' （汎用版）DataGridViewをcsvに保存
    ''' </summary>
    ''' <param name="filepath">保存するパス</param>
    ''' <param name="dgv">保存するDataGridView</param>
    ''' <param name="ENC">エンコード（初期：SHIFT-JIS）</param>
    ''' <param name="kakomi">コーテーションで囲むか（初期：False=囲まない）</param>
    ''' <param name="visibleFlag">表示行のみ保存（初期：False=全て）</param>
    ''' <param name="CsvDelimiter">区切り文字（初期：「,」）</param>
    Public Sub DGV_TO_CSV_SAVE(filepath As String, dgv As DataGridView, Optional ENC As String = "SHIFT-JIS", Optional kakomi As Boolean = False, Optional visibleFlag As Boolean = False, Optional CsvDelimiter As String = ",")
        If dgv.RowCount = 0 Then
            Exit Sub
        End If

        dgv.EndEdit()
        Dim strArray As New ArrayList

        Dim kakomiStr As String = ""
        If kakomi Then
            kakomiStr = """"
        End If

        Dim header As String = ""
        For c As Integer = 0 To dgv.ColumnCount - 1
            header &= kakomiStr & dgv.Columns(c).HeaderText & kakomiStr & CsvDelimiter
        Next
        header = header.TrimEnd(CsvDelimiter)
        strArray.Add(header)

        For r As Integer = 0 To dgv.RowCount - 1
            If visibleFlag = False Or (visibleFlag And dgv.Rows(r).Visible) Then
                Dim line As String = ""
                For c As Integer = 0 To dgv.ColumnCount - 1
                    line &= kakomiStr & dgv.Item(c, r).Value & kakomiStr & CsvDelimiter
                Next
                line = line.TrimEnd(CsvDelimiter)
                strArray.Add(line)
            End If
        Next

        Dim Values() As String = DirectCast(strArray.ToArray(GetType(String)), String())
        File.WriteAllLines(filepath, Values, Encoding.GetEncoding(ENC))
    End Sub

    ''' <summary>
    ''' SaveCsv
    ''' </summary>
    ''' <param name="fp">String</param>
    ''' <param name="mode">1=「"」囲み、2=vbTab</param>
    ''' <param name="dgv">DataGridView</param>
    ''' <param name="ENC">optional SHIFT-JIS</param>
    Public Sub SaveCsv(fp As String, mode As Integer, dgv As DataGridView, Optional ENC As String = "SHIFT-JIS")
        Dim visibleFlag As Boolean = False
        For r As Integer = 0 To dgv.Rows.Count - 1
            If dgv.Rows(r).Visible = False Then
                Dim DR As DialogResult = MsgBox("表示行のみ保存しますか？", MsgBoxStyle.YesNo Or MsgBoxStyle.SystemModal)
                If DR = Windows.Forms.DialogResult.Yes Then
                    visibleFlag = True
                    Exit For
                Else
                    Exit For
                End If
            End If
        Next

        ' CSVファイルオープン
        Dim str As String = ""
        Dim sw As StreamWriter = New StreamWriter(fp, False, Encoding.GetEncoding(ENC))

        '区切り
        Dim delimiter As String = ","
        If mode = 2 Then
            delimiter = vbTab
        End If

        'ヘッダー行
        Dim dt As String = ""
        For c As Integer = 0 To dgv.Columns.Count - 1
            dt &= dgv.Columns(c).HeaderText & delimiter
        Next
        dt = dt.TrimEnd(delimiter) & vbLf
        sw.Write(dt)

        For r As Integer = 0 To dgv.Rows.Count - 1
            If visibleFlag = False Or (visibleFlag = True And dgv.Rows(r).Visible = True) Then
                For c As Integer = 0 To dgv.Columns.Count - 1
                    ' DataGridViewのセルのデータ取得
                    dt = ""
                    If dgv.Rows(r).Cells(c).Value Is Nothing = False Then
                        dt = dgv.Rows(r).Cells(c).Value.ToString()
                        dt = Replace(dt, vbCrLf, vbLf)
                        'dt = Replace(dt, vbLf, "")
                        If mode = 1 Then
                            dt = """" & dt & """"
                        Else
                            If Not dt Is Nothing Then
                                dt = Replace(dt, """", """""")
                                Select Case True
                                    Case InStr(dt, ","), InStr(dt, vbLf), InStr(dt, vbCr), InStr(dt, """") 'Regex.IsMatch(dt, ".*,.*|.*" & vbLf & ".*|.*" & vbCr & ".*|.*"".*")
                                        dt = """" & dt & """"
                                End Select
                            End If
                        End If
                    End If
                    If c < dgv.Columns.Count - 1 Then
                        dt = dt & delimiter
                    End If
                    ' CSVファイル書込
                    sw.Write(dt)
                Next
                sw.Write(vbLf)
            End If
        Next

        ' CSVファイルクローズ
        sw.Close()
    End Sub

    ''' <summary>
    ''' CurrentUserのRunにアプリケーションの実行ファイルパスを登録する
    ''' </summary>
    Public Sub SetCurrentVersionRun()
        'Runキーを開く
        Dim regkey As Microsoft.Win32.RegistryKey =
            Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
            "Software\Microsoft\Windows\CurrentVersion\Run", True)
        '値の名前に製品名、値のデータに実行ファイルのパスを指定し、書き込む
        regkey.SetValue(Application.ProductName, Application.ExecutablePath)
        '閉じる
        regkey.Close()
        MsgBox("OS起動時スタート登録しました", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
    End Sub

    ''' <summary>
    ''' CurrentUserのRunにアプリケーションの実行ファイルパスを登録する
    ''' </summary>
    Public Sub DelCurrentVersionRun()
        'Runキーを開く
        Dim regkey As Microsoft.Win32.RegistryKey =
            Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
            "Software\Microsoft\Windows\CurrentVersion\Run", True)
        '値の名前に製品名、値のデータに実行ファイルのパスを指定し、書き込む
        regkey.DeleteValue(Application.ProductName, False)
        '閉じる
        regkey.Close()
        MsgBox("OS起動時スタート削除しました", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
    End Sub

    Public Function RegistryRead(key As String)
        'キー（HKEY_CURRENT_USER\Software\test\sub）を読み取り専用で開く
        Dim regkey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(key, False)
        'キーが存在しないときは null が返される
        If regkey Is Nothing Then
            Return vbNull
        End If

        '文字列を読み込む（指定した名前の値が存在しないときは null が返される）
        Dim stringValue As String = DirectCast(regkey.GetValue("AutoCheck2"), String)
        'キーに値が存在しないときに指定した値を返すようにするには、次のようにする
        '（ここでは"default"を返す）
        'Dim stringValue As String = DirectCast(regkey.GetValue("string", "default"), String)

        ''整数（REG_DWORD）を読み込む
        'Dim intValue As Integer = CInt(regkey.GetValue("int"))
        ''整数（REG_QWORD）を読み込む
        'Dim longVal As Long = CLng(regkey.GetValue("QWord"))
        ''文字列配列を読み込む
        'Dim strings As String() = DirectCast(regkey.GetValue("StringArray"), String())
        ''バイト配列を読み込む
        'Dim bytes As Byte() = DirectCast(regkey.GetValue("Bytes"), Byte())

        '閉じる
        regkey.Close()
        Return stringValue
    End Function


    ''' <summary>
    ''' 文字列からバイト数を指定して部分文字列を取得する。
    ''' </summary>
    ''' <param name="value">対象文字列。</param>
    ''' <param name="startIndex">開始位置。（バイト数）</param>
    ''' <param name="length">長さ。（バイト数）</param>
    ''' <returns>部分文字列。</returns>
    ''' <remarks>文字列は <c>Shift_JIS</c> でエンコーディングして処理を行います。</remarks>
    Public Function SubstringByte(ByVal value As String, ByVal startIndex As Integer, ByVal length As Integer) As String
        Dim sjisEnc As Encoding = Encoding.GetEncoding("Shift_JIS")
        Dim byteArray() As Byte = sjisEnc.GetBytes(value)

        If byteArray.Length < startIndex + 1 Then
            Return ""
        End If

        If byteArray.Length < startIndex + length Then
            length = byteArray.Length - startIndex
        End If

        Dim cut As String = sjisEnc.GetString(byteArray, startIndex, length)

        ' 最初の文字が全角の途中で切れていた場合はカット
        Dim left As String = sjisEnc.GetString(byteArray, 0, startIndex + 1)
        Dim first As Char = value(left.Length - 1)
        If 0 < cut.Length AndAlso Not first = cut(0) Then
            cut = cut.Substring(1)
        End If

        ' 最後の文字が全角の途中で切れていた場合はカット
        left = sjisEnc.GetString(byteArray, 0, startIndex + length)

        Dim last As Char = value(left.Length - 1)
        If 0 < cut.Length AndAlso Not last = cut(cut.Length - 1) Then
            cut = cut.Substring(0, cut.Length - 1)
        End If

        Return cut
    End Function

    ''' <summary>
    ''' 指定された文字列を含むウィンドウタイトルを持つプロセスを取得します。
    ''' </summary>
    ''' <param name="windowTitle">ウィンドウタイトルに含む文字列。</param>
    ''' <returns>該当するプロセスの配列。</returns>
    Public Function GetProcessesByWindowTitle(windowTitle As String) As _
            System.Diagnostics.Process()
        Dim list As New System.Collections.ArrayList()

        'すべてのプロセスを列挙する
        Dim p As System.Diagnostics.Process
        For Each p In System.Diagnostics.Process.GetProcesses()
            '指定された文字列がメインウィンドウのタイトルに含まれているか調べる
            If 0 <= p.MainWindowTitle.IndexOf(windowTitle) Then
                '含まれていたら、コレクションに追加
                list.Add(p)
            End If
        Next

        'コレクションを配列にして返す
        Return DirectCast(list.ToArray(GetType(System.Diagnostics.Process)),
            System.Diagnostics.Process())
    End Function


    '********************************************
    '--------------------------------------------
    ' JANチェックデジット計算
    '--------------------------------------------
    '********************************************
    Public Function Calc_Check_Digit(strCheck_Digit_Value) As String
        Dim intValue(12) As Integer      '値を1桁ずつ格納する配列
        Dim i As Integer
        Dim intEven As Integer      '偶数
        Dim intOdds As Integer      '奇数
        Dim strValue_13 As String       '13桁のJANコード

        '12桁及び7桁の場合があり、左埋めにする。
        strValue_13 = strCheck_Digit_Value.PadRight(13, "0"c)   '左埋め

        intEven = 0
        intOdds = 0

        '右から2番目から13番目の値について、偶数位置、奇数位置の値をそれぞれ足していく。
        For i = 2 To Len(strValue_13)
            If (i Mod 2) <> 0 Then
                intOdds = intOdds + CInt(Mid(strValue_13, 14 - i, 1))
            Else
                intEven = intEven + CInt(Mid(strValue_13, 14 - i, 1))
            End If
        Next i

        '偶数位置の合計を3倍する
        intEven = intEven * 3

        '偶数位置の合計の3倍の値と奇数位置の合計の値を足し、
        '1の位の値が0の場合は0を返す。その他の場合は10から1の位の値を引いた値を返す。
        Dim rightA As String = CStr(intEven + intOdds)
        If CInt(rightA.Substring(rightA.Length - 1, 1)) = 0 Then
            Calc_Check_Digit = "0"
        Else
            Calc_Check_Digit = CStr(10 - CInt(rightA.Substring(rightA.Length - 1, 1)))
        End If
    End Function
    '-------------------------------------------------------------------------------------------------------------

    ''' <summary>
    ''' DataTableの内容をCSVファイルに保存する
    ''' </summary>
    ''' <param name="dt">CSVに変換するDataTable</param>
    ''' <param name="csvPath">保存先のCSVファイルのパス</param>
    ''' <param name="writeHeader">ヘッダを書き込む時はtrue。</param>
    Public Sub ConvertDataTableToCsv(dt As DataTable, csvPath As String, writeHeader As Boolean)
        'CSVファイルに書き込むときに使うEncoding
        Dim enc As System.Text.Encoding =
            System.Text.Encoding.GetEncoding("Shift_JIS")

        '書き込むファイルを開く
        Dim sr As New System.IO.StreamWriter(csvPath, False, enc)

        Dim colCount As Integer = dt.Columns.Count
        Dim lastColIndex As Integer = colCount - 1
        Dim i As Integer

        'ヘッダを書き込む
        If writeHeader Then
            For i = 0 To colCount - 1
                'ヘッダの取得
                Dim field As String = dt.Columns(i).Caption
                '"で囲む
                field = EncloseDoubleQuotesIfNeed(field)
                'フィールドを書き込む
                sr.Write(field)
                'カンマを書き込む
                If lastColIndex > i Then
                    sr.Write(","c)
                End If
            Next
            '改行する
            sr.Write(vbCrLf)
        End If

        'レコードを書き込む
        Dim row As DataRow
        For Each row In dt.Rows
            For i = 0 To colCount - 1
                'フィールドの取得
                Dim field As String = row(i).ToString()
                '"で囲む
                field = EncloseDoubleQuotesIfNeed(field)
                'フィールドを書き込む
                sr.Write(field)
                'カンマを書き込む
                If lastColIndex > i Then
                    sr.Write(","c)
                End If
            Next
            '改行する
            sr.Write(vbCrLf)
        Next

        '閉じる
        sr.Close()
    End Sub

    ''' <summary>
    ''' 必要ならば、文字列をダブルクォートで囲む
    ''' </summary>
    Private Function EncloseDoubleQuotesIfNeed(field As String) As String
        If NeedEncloseDoubleQuotes(field) Then
            Return EncloseDoubleQuotes(field)
        End If
        Return field
    End Function

    ''' <summary>
    ''' 文字列をダブルクォートで囲む
    ''' </summary>
    Private Function EncloseDoubleQuotes(field As String) As String
        If field.IndexOf(""""c) > -1 Then
            '"を""とする
            field = field.Replace("""", """""")
        End If
        Return """" & field & """"
    End Function

    ''' <summary>
    ''' 文字列をダブルクォートで囲む必要があるか調べる
    ''' </summary>
    Private Function NeedEncloseDoubleQuotes(field As String) As Boolean
        Return field.IndexOf(""""c) > -1 OrElse
            field.IndexOf(","c) > -1 OrElse
            field.IndexOf(ControlChars.Cr) > -1 OrElse
            field.IndexOf(ControlChars.Lf) > -1 OrElse
            field.StartsWith(" ") OrElse
            field.StartsWith(vbTab) OrElse
            field.EndsWith(" ") OrElse
            field.EndsWith(vbTab)
    End Function

    '-------------------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' 画像ダウンロード
    ''' </summary>
    ''' <param name="fol">ダウンロード先フォルダ</param>
    ''' <param name="url">ダウンロードし解析するURL</param>
    ''' <returns></returns>
    Public Function TM_IMGDL(ByVal fol As String, ByVal url As String)
        If url <> "" Then
            url = Replace(url, """", "")
            url = Replace(url, "'", "")

            Dim flag As Boolean = True
            Dim fName As String = Path.GetFileName(url)
            Dim newPath As String = fol & "\" & fName

            Try
                ServicePointManager.SecurityProtocol = DirectCast(3072, SecurityProtocolType)
                Dim wc As New System.Net.WebClient()
                wc.DownloadFile(url, newPath)
                wc.Dispose()
            Catch ex As Exception
                flag = False
            End Try

            If flag Then
                Try
                    'Dim a = Path.GetExtension(fName)
                    If Path.GetExtension(url) = "" Then
                        Dim img As Image = Image.FromFile(newPath)
                        If img.RawFormat.Equals(Imaging.ImageFormat.Gif) Then
                            img.Dispose()
                            If Not File.Exists(newPath & ".gif") Then
                                File.Move(newPath, newPath & ".gif")
                            End If
                        ElseIf img.RawFormat.Equals(Imaging.ImageFormat.Jpeg) Then
                            img.Dispose()
                            If Not File.Exists(newPath & ".jpg") Then
                                File.Move(newPath, newPath & ".jpg")
                            End If
                        ElseIf img.RawFormat.Equals(Imaging.ImageFormat.Png) Then
                            img.Dispose()
                            If Not File.Exists(newPath & ".png") Then
                                File.Move(newPath, newPath & ".png")
                            End If
                        Else
                            img.Dispose()
                            If Not File.Exists(newPath & ".jpg") Then
                                File.Move(newPath, newPath & ".jpg")
                            End If
                        End If
                    End If
                Catch ex As Exception

                End Try
            End If
        End If

        Return True
    End Function

    ''' <summary>
    ''' コントロールのDoubleBufferedプロパティをTrueにする
    ''' </summary>
    ''' <param name="control">対象のコントロール</param>
    Public Sub TM_EnableDoubleBuffering(control As Control)
        control.GetType().InvokeMember(
            "DoubleBuffered",
            BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.SetProperty,
            Nothing,
            control,
            New Object() {True})
    End Sub
End Module
