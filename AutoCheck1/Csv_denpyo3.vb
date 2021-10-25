Imports System.ComponentModel
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports Microsoft.International.Converters
Imports Microsoft.VisualBasic.FileIO
Imports NPOI.SS.UserModel
Imports Ookii.Dialogs
Imports PdfGeneratorNetFree
Imports PdfGeneratorNetFree.PdfContentItem
Imports PdfGeneratorNetFree.PgnStyle
Imports Microsoft.VisualBasic

Public Class Csv_denpyo3
    'Private Declare Function HideCaret Lib "user32.dll" (ByVal hwnd As IntPtr) As Boolean
    'Private dgv1dNoArray As New ArrayList
    Private dgv1dNoArray As New ArrayList
    Private dgv6CodeArray As New ArrayList
    Private dgv6DaihyoCodeArray As New ArrayList
    Private dH6 As New ArrayList


    Private sakiPath As String = ""
    Private sakiPath2 As String = ""

    ReadOnly ENC_SJ As Encoding = Encoding.GetEncoding("shift-jis")
    Public Shared CnAccdb As String = ""

    Private koumoku As New Hashtable
    Private helpArray As New ArrayList
    Private helpMessage As New ArrayList

    Private true_str As String = "true"
    Private false_str As String = "false"

    Private nagoya_str As String = "名古屋"
    Private dazaifu_str As String = "太宰府"
    Private yamato_str As String = "ヤマト"
    Private fukususouko_str As String = "複数倉庫"
    'やまとへんかん
    'Dim isYamatoGood As Boolean = False
    'Dim isYamatoGood_fukumu As Boolean = False

    Public YU2FlagLoad = False

    '邮局 适用P60发送地域
    '大阪以下
    Private checkaddress_oosakika As String() = New String() {"熊本県", "宮崎県", "鹿児島県", "福岡県", "佐賀県", "長崎県", "大分県", "徳島県", "香川県", "愛媛県", "高知県", "鳥取県", "島根県", "岡山県", "広島県", "山口県", "滋賀県", "京都府", "大阪府", "兵庫県", "奈良県", "和歌山県"}
    '大阪以上
    Private checkaddress_oosakizyou As String() = New String() {"富山県", "石川県", "福井県", "岐阜県", "静岡県", "愛知県", "三重県", "新潟県", "長野県", "茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県", "山梨県", "宮城県", "山形県", "福島県", "青森県", "岩手県", "秋田県", "北海道"}
    Private yamato_goods As String() = New String() {"ny264-50", "ny306-51", "ny331-50-306", "ny331-50-flpi", "ny331-50-flwh", "ny331-50-393bl", "ny331-50-393blbl", "ny331-50-393blhu", "ny331-50-393hu", "ny331-50-393im", "ny331-50-393iv", "ny331-50-393kf", "ny331-50-393kh", "ny331-50-393kobl", "ny331-50-393kohu", "ny331-50-393koice", "ny331-50-393kolash", "ny331-50-393kosaku", "ny331-50-393koyegr", "ny331-50-393lash", "ny331-50-393milk", "ny331-50-393ne", "ny331-50-393rose", "ny331-50-393saku", "ny331-50-393yegr",
    "ny331-50-pi", "ny331-50-pa", "ny331-50-dapi", "ny331-50-bk", "ny331-50-hu", "ny331-50-wh", "ny331-50-be", "ny331-50-co", "ny331-50-lgr",
    "ny331-50-lor", "ny331-50-rose", "ny263-50-c", "ny341-50-40", "ny263-306-42", "ny331-50-ye", "ny344", "ny373-hu", "ny373-pi",
    "n  y373-bk", "ny373-kobk", "ny373-kowh", "ny373-wh", "ny385-a", "ny385-b", "ny385-c", "ny263-51",
        "ny385-d", "ny100-bk", "ny100-gr", "ny100-wh", "ee247", "zk200",
        "ad205-bk", "ad205-bl", "ad205-gl", "ad205-gr", "ad205-pi", "mb077-bk", "mb077-pi",
     "ny331-50-kobk", "ny331-50-kohu", "ny331-50-koor", "ny331-50-kopi",
     "ad147-gr", "ad147-or", "ad269-bl", "ad269-pi", "ad269-bk", "de112", "e094", "od352-bk", "od352-bl", "od352-re",
     "sl050-bl", "sl050-gr", "sl050-ka", "sl050-re", "sl058",
     "zk216", "mb137", "sl065-bl", "sl065-pa", "sl065-pi", "sl065-ra", "sl065-wh", "sl066-bl", "sl066-pa", "sl066-pi", "sl066-ra", "sl066-wh",
      "sl067-bl", "sl067-pa", "sl067-pi", "sl067-ra", "sl067-wh", "ny171", "ny305", "ny233", "od304", "od343", "ee247",
    "rt015-bk", "rt015-bl", "zk213", "ny236", "ad267-l", "rt003-bk", "rt003-wh",
    "e103", "kp005", "ad194", "mb082-bk", "mb082-re",
    "ad229-bk", "ad229-re", "ad267-m", "ad267-l", "ee204", "mb064", "sl047", "zk192"}



    '增山发的需求 变成yamato
    ' ad096    "zk200",
    'ad194   ヤマト
    'ad229-bk
    'ad229-re
    'ad267-m
    'ap039-Or　
    'ee204
    'ee270-bk-xl
    'ee270-bl-l
    'ee270-bl-xl
    'ee270-si-xl
    'kp005
    'mb064
    'mb137
    'sl047
    'sl066-bl
    'sl066-pi
    'sl067-bl
    'zk192
    'zk213
    'zk216
    'zk218-01
    'zk253-bk
    'zk280


    'ヤマト

    '仓库需求去掉部分发yamato的货号 20210207
    'zk221  zk253  od322 e098 ad096 zk280



    Dim yamato_title As String = "お客様管理番号,送り状種類,クール区分,伝票番号,出荷予定日,お届け予定（指定）日,配達時間帯,お届け先コード,お届け先電話番号,お届け先電話番号枝番,お届け先郵便番号,お届け先住所,お届け先住所（アパートマンション名）,お届け先会社・部門名１,お届け先会社・部門名２,お届け先名,お届け先名略称カナ,敬称,ご依頼主コード,ご依頼主電話番号,ご依頼主電話番号枝番,ご依頼主郵便番号,ご依頼主住所,ご依頼主住所（アパートマンション名）,ご依頼主名,ご依頼主略称カナ,品名コード１,品名１,品名コード２,品名２,荷扱い１,荷扱い２,記事,コレクト代金引換額（税込）,コレクト内消費税額等,営業所止置き,営業所コード,発行枚数,個数口枠の印字,ご請求先顧客コード,ご請求先分類コード,運賃管理番号,クロネコwebコレクトデータ登録,クロネコwebコレクト加盟店番号,クロネコwebコレクト申込受付番号１,クロネコwebコレクト申込受付番号２,クロネコwebコレクト申込受付番号３,お届け予定ｅメール利用区分,お届け予定ｅメールe-mailアドレス,入力機種,お届け予定eメールメッセージ,お届け完了eメール利用区分,お届け完了ｅメールe-mailアドレス,お届け完了ｅメールメッセージ,クロネコ収納代行利用区分,収納代行決済ＱＲコード印刷,収納代行請求金額(税込),収納代行内消費税額等,収納代行請求先郵便番号,収納代行請求先住所,収納代行請求先住所（アパートマンション名）,収納代行請求先会社・部門名１,収納代行請求先会社・部門名２,収納代行請求先名(漢字),収納代行請求先名(カナ),収納代行問合せ先名(漢字),収納代行問合せ先郵便番号,収納代行問合せ先住所,収納代行問合せ先住所（アパートマンション名）,収納代行問合せ先電話番号,収納代行管理番号,収納代行品名,収納代行備考,複数口くくりキー,検索キータイトル１,検索キー１,検索キータイトル２,検索キー２,検索キータイトル３,検索キー３,検索キータイトル４,検索キー４,検索キータイトル５,検索キー５,予備,予備,投函予定メール利用区分,投函予定メールe-mailアドレス,投函予定メールメッセージ,投函完了メール（お届け先宛）利用区分,投函完了メール（お届け先宛）e-mailアドレス,投函完了メール（お届け先宛）メールメッセージ,投函完了メール（ご依頼主宛）利用区分,投函完了メール（ご依頼主宛）e-mailアドレス,投函完了メール（ご依頼主宛）メールメッセージ,連携管理番号,通知メールアドレス"
    Private yamato_title_sp As String() = yamato_title.Split(",")



    Private yupaku_addressArr As String() = New String() {"TEST"}
    'Private yupaku_addressArrPro As String() = New String() {"福島県", "山口県", "兵庫県", "大阪", "東京"}
    Private yupaku_addressArrPro As String() = New String() {"test県"}
    'Private yupaku_goods As String() = New String() {"ad081-bk", "ny263-51", "ny331-50-bk", "ad009-bk", "ad109", "ny098-mint"}
    Private yupaku_goods As String() = New String() {""}

    Public yupakucheck As Boolean = False
    Private yupaku_str As String = "ゆう2"
    'Private yupaku_OutlyingislandsArr As String() = New String() {"沖縄"}
    Private yupaku_OutlyingislandsArr As String() = New String() {"test"}
    Dim isyupakuGoodBool As Boolean = False
    Dim YU2Denpyus As ArrayList = New ArrayList
    Dim denpyounoS As New ArrayList
    Dim CurDenpyoNo As String
    Dim siharaiYU2 As String = ""
    Dim YU2tenpuArr As String() = New String() {"test"}


    Public Function isyupakutenpu(tenpo As String) As Boolean
        If tenpo = "" Then
            Return False
        Else
            For Each ten In YU2tenpuArr
                If InStr(tenpo, ten) Then
                    Return True
                End If
            Next
        End If
        Return False
    End Function

    Public Function isyupakuGoodsbyCode(code As String) As Boolean
        If code = "" Then
            Return False
        Else
            For index = 0 To yupaku_goods.Count - 1
                If yupaku_goods(index) = code Then
                    Return True
                End If
            Next
        End If
        Return False
    End Function



    Public Function isyupakubyAddress(address As String) As Boolean
        If address = "" Then
            Return False
        Else
            For index = 0 To yupaku_addressArrPro.Count - 1
                'If  InStr ()  yupaku_addressArrPro(index) = address Then
                If InStr(address, yupaku_addressArrPro(index)) Then
                    Return True
                End If
            Next
        End If
        Return False
    End Function

    Public Function IsOutlyingislandsbyAddress(address As String) As Boolean
        If address = "" Then
            Return False
        Else
            For index = 0 To yupaku_OutlyingislandsArr.Count - 1
                If InStr(address, yupaku_OutlyingislandsArr(index)) Then
                    Return True
                End If
            Next
        End If
        Return False
    End Function


    'MaxMumAndCurMunSagawaYupakuPro
    Dim yupakufName As String = Path.GetDirectoryName(Form1.appPath) & "\config\version2\MaxMumAndCurMunSagawaYupakuPro.txt"
    Public Function isyupakuGoodsOverMax()

        'Dim fName As String = ""
        Dim csvRecords As New ArrayList
        Dim max As String
        Dim cur As String
        Dim returndata As New ArrayList

        'fName = Path.GetDirectoryName(Form1.appPath) & "\config\version2\MaxMumAndCurMunSagawaYupakuPro.txt"
        csvRecords = TM_CSV_READ(yupakufName)(0)
        Debug.WriteLine(csvRecords)
        For r As Integer = 0 To csvRecords.Count - 1
            Dim sArray As String() = Split(csvRecords(r), "|=|")
            'Dim sArraytemp As String() = Split(sArray(0), ",")
            Debug.WriteLine("ceshiyong")
            Debug.WriteLine(sArray(0))
            returndata.Add(sArray(0))
        Next



        If IsNothing(returndata) Then
            Return -1
        End If

        Dim ss As String = Format(Now(), "yyyy-MM-dd")

        If returndata(2) <> ss Then
            WriteFile(yupakufName, returndata(0), 0)
            Return returndata(0)
        End If

        'If returndata(0) = "" Or returndata(1) = "" Then
        '    Return -1
        'End If

        If returndata(0) > returndata(1) Then
            Return returndata(0) - returndata(1)
        Else
            'WriteFile(yupakufName, returndata(0), returndata(0))
            Return -1
        End If
        Return -1
    End Function

    Public Sub WriteFile(fileName As String, max As String, cur As String)

        'If fileName = "" Or data = "" Then
        '    Exit Sub
        'End If

        Try
            'Dim bytes() As Byte = File.ReadAllBytes(fileName)



            'Dim writeFile As TextWriter = New StreamWriter("c:\textwriter.txt")
            Dim writeFile As TextWriter = New StreamWriter(fileName)

            'writeFile.WriteLine("890", "350")
            Dim ss As String = Format(Now(), "yyyy-MM-dd")

            MsgBox("开始" + ss)
            writeFile.WriteLine(max & vbCrLf & cur & vbCrLf & ss)

            writeFile.Flush()
            writeFile.Close()
            writeFile = Nothing



        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


        MsgBox("完成")


    End Sub





    Dim managertool As ManagerTools

    Public Sub test()
        managertool = New ManagerTools()
        managertool.ManagerFun("tesssss")


    End Sub

    Private Sub UpdataManagerHands_yupaku()
        'ManagerHand_yupaku = New ManagerTools()


    End Sub



    Private Sub Csv_denpyo3_Load(sender As Object, e As EventArgs) Handles Me.Load
        sakiPath = Form1.サーバーToolStripMenuItem.Text & "\update"
        sakiPath2 = Form1.ロケーションToolStripMenuItem.Text
        CnAccdb = Path.GetDirectoryName(Form1.appPath) & "\db\sagawa.accdb"



        'コントロール設定
        'DGV1    'メインcsv読み込み
        'DGV2    '商品表示
        'DGV3    '明細csv読み込み
        'DGV4    '伝票確認
        'DGV5    '商品検索
        'DGV6    '配送マスタ（location.csv）
        'DGV7    'e飛伝2
        'DGV8    'BIZ-logi
        'DGV9    'メール便
        'DGV13   '定形外
        'DGV10   '宅配便サイズ別容量変換リスト
        'DGV11   '未定
        'DGV12   'プレビュー
        'DGV14   '履歴
        'DGV15   '元データエラー
        'DGV16   'ピッキングリスト
        'DGV17   '別紙PDF
        'DGV18   'ファイル一覧

        '******************************************
        'dgv画面チラつき防止
        '******************************************
        'DataGirdViewのTypeを取得
        Dim dgvtype As Type = GetType(DataGridView)
        'プロパティ設定の取得
        Dim dgvPropertyInfo As Reflection.PropertyInfo = dgvtype.GetProperty("DoubleBuffered", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)

        '対象のDataGridViewにtrueをセットする
        Dim dgvArray As DataGridView() = {DGV1, DGV2, DGV3, DGV4, DGV5, DGV6, DGV7, DGV8, DGV9, DGV10, DGV12, DGV13, DGV14, DGV15, DGV16, DGV17, DGV18}
        For i As Integer = 0 To dgvArray.Length - 1
            dgvPropertyInfo.SetValue(dgvArray(i), True, Nothing)
        Next

        '******************************************
        '初期化
        '******************************************
        Dim fName As String = ""
        Dim csvRecords As New ArrayList


        'Dim Managertool As New ManagerTools
        'Dim cc = Managertool.ManagerFun("zaizheli")
        'PrintLine("nihao" + cc)

        '******************************************
        '各種設定等読み込み
        '******************************************
        fName = Path.GetDirectoryName(Form1.appPath) & "\config\csvd3config.txt"
        Dim kLines As String() = File.ReadAllLines(fName, ENC_SJ)
        For i As Integer = 0 To kLines.Length - 1


            Dim kl As String() = Split(kLines(i), "=")
            Select Case kl(0)
                Case "元データ規定" & HS1.Text
                    If kl(1) = "true" Then CheckBox1.Checked = True _
                    Else CheckBox1.Checked = False
                Case "元データ規定" & HS2.Text
                    If kl(1) = "true" Then CheckBox4.Checked = True _
                    Else CheckBox4.Checked = False
                'Case "元データ規定" & HS3.Text
                '    If kl(1) = "true" Then CheckBox3.Checked = True _
                '    Else CheckBox3.Checked = False
                Case "元データ規定" & HS4.Text
                    If kl(1) = "true" Then CheckBox32.Checked = True _
                    Else CheckBox32.Checked = False
                Case "元データ規定" & HS1.Text & "航空"
                    If kl(1) = "true" Then CheckBox17.Checked = True _
                        Else CheckBox17.Checked = False
                Case "元データ規定" & HS2.Text & "航空"
                    If kl(1) = "true" Then CheckBox18.Checked = True _
                        Else CheckBox18.Checked = False
                Case "元データ規定" & HS4.Text & "航空"
                    If kl(1) = "true" Then CheckBox3.Checked = True _
                        Else CheckBox3.Checked = False
            End Select
        Next

        helpMessage.Add("NE検索画面のcsvを追加すると、ピッキングの情報を追加でき、修正しやすくなります")
        helpMessage.Add("定形外は2便まで、ゆうパケットは3便まで、それ以上は宅配便。ゆうパケットと定形外は宅配便です")
        helpMessage.Add("宅配便はPでサイズ表記、ゆうパケットはMで100％表記、定形外はTでグラム表記(250gまで)")

        Dim memo As String = "2020/03/31" & vbCrLf & "    50個1便: ny261,ny263"
        memo &= vbCrLf & "2020/04/02" & vbCrLf & "    50個1便: ny264"
        memo &= vbCrLf & "2020/04/11" & vbCrLf & "    50個1便: mask01,mask01-bl,mask01-ko"
        memo &= vbCrLf & "2020/04/14" & vbCrLf & "    50個1便(以前全部) -> 60個1便"
        memo &= vbCrLf & "2020/05/01" & vbCrLf & "    40個1便: mask01,mask-wh,mask05,mask01-bl,mask01-ko"
        memo &= vbCrLf & "2020/05/13" & vbCrLf & "    2個までメール便: ny263-51"
        memo &= vbCrLf & "2020/06/08" & vbCrLf & "    4個以上宅配便: ny275-pi,ny275-bk"
        memo &= vbCrLf & "2020/06/16" & vbCrLf & "    航空便不可: ny221,zk121"
        memo &= vbCrLf & "2020/06/18" & vbCrLf & "    航空便不可: ny177,od363"
        memo &= vbCrLf & "2020/06/25" & vbCrLf & "    航空便不可: ad201"
        memo &= vbCrLf & "2020/07/30" & vbCrLf & "    44個1便(宅配便): ny263-51"
        memo &= vbCrLf & "2020/08/05" & vbCrLf & "    2個までメール便: ny306-51"
        memo &= vbCrLf & "2020/08/28" & vbCrLf & "    50個1便(宅配便): ny263-51"
        memo &= vbCrLf & "2020/09/30" & vbCrLf & "    1個2便(宅配便): ny264-100-4000,de055"
        memo &= vbCrLf & "2020/10/01" & vbCrLf & "    同梱しない、2個一便になる: ad228"
        memo &= vbCrLf & "2020/10/05" & vbCrLf & "    同梱しない: de072"
        memo &= vbCrLf & "2020/10/07" & vbCrLf & "    2個までメール便: ny263-306-51"
        memo &= vbCrLf & "2020/10/14" & vbCrLf & "    1個1便: zk101 5個まで1便:ad105"
        memo &= vbCrLf & "2020/11/04" & vbCrLf & "    1個2便: ny328"
        memo &= vbCrLf & "2020/05/13" & vbCrLf & "    2個以上(2個を含む)宅配便: ny263-51"
        memo &= vbCrLf & "2020/12/07" & vbCrLf & "    3個以上(3個を含む)宅配便: ny263-51"
        memo &= vbCrLf & "2021/01/14" & vbCrLf & "    40個1便(宅配便): ny263-51"
        memo &= vbCrLf & "2021/01/21" & vbCrLf & "    6個まで1便(宅配便): ny185"
        memo &= vbCrLf & "2021/01/28" & vbCrLf & "    2個以上(2個を含む)宅配便: ny331-50-flwh"
        memo &= vbCrLf & "2021/02/01" & vbCrLf & "    3個以上(3個を含む)宅配便: ny331-50-flwh"
        'memo &= vbCrLf & "2020/12/25" & vbCrLf & "    2個以上(2個を含む)（冲绳北海道除外） 宅配便: sl065 sl066 sl067"
        memo &= vbCrLf & "2021/02/01" & vbCrLf & "    40個1便: ny264"
        memo &= vbCrLf & "2021/02/03" & vbCrLf & "    3個以上(3個を含む)宅配便: ny263-306-42,ny331-50マスク"
        memo &= vbCrLf & "2021/02/04" & vbCrLf & "    6000枚1便: ny263-51"
        memo &= vbCrLf & "2021/02/18" & vbCrLf & "    8個1便: ad009 4個1便: ny263-50-c(ヤマト)"
        memo &= vbCrLf & "2021/02/24" & vbCrLf & "    50個1便: ny263-306-42"
        memo &= vbCrLf & "2021/02/24" & vbCrLf & "    40個1便: ny263-331-50"
        TextBox11.Text = memo

        Dim memo2 As String = "※複数倉庫の場合はneで伝票を分割してください"

        'memo2 &= vbCrLf & "変更日: 2020/02/05"
        'memo2 &= vbCrLf
        'memo2 &= vbCrLf & "太宰府:"
        'memo2 &= vbCrLf & "ny261-1000-a(数量: 1, 3 ,5[奇数]...)"

        'memo2 &= vbCrLf & vbCrLf & "名古屋:"
        'memo2 &= vbCrLf & "ny261-2000-a(数量: 2, 4 ,6[偶数]...)"
        'memo2 &= vbCrLf
        'memo2 &= vbCrLf & "ny263-51(40倍数)"
        'memo2 &= vbCrLf
        'memo2 &= vbCrLf & "ny264(40倍数)"
        'memo2 &= vbCrLf & "ny264-100(20倍数)"
        'memo2 &= vbCrLf & "ny264-200(10倍数)"
        'memo2 &= vbCrLf & "ny264-500(4倍数)"

        TextBox13.Text = memo2

        '******************************************
        '配送マスタ、ロケーションマスタの読み込み
        DGV6_update()
        dH6 = TM_HEADER_GET(DGV6)
        '******************************************

        '各ヘッダー読み込み
        For k As Integer = 0 To 5
            Dim fStrArray As String() = Nothing
            Dim dgv As DataGridView = Nothing
            If k = 0 Then
                fStrArray = File.ReadAllLines(Path.GetDirectoryName(Form1.appPath) & "\template\" & TextBox8.Text & ".dat", Encoding.GetEncoding("shift-jis"))
                dgv = DGV1
            ElseIf k = 1 Then
                fStrArray = File.ReadAllLines(Path.GetDirectoryName(Form1.appPath) & "\template\" & TextBox2.Text & ".dat", Encoding.GetEncoding("shift-jis"))
                dgv = DGV7
            ElseIf k = 2 Then
                fStrArray = File.ReadAllLines(Path.GetDirectoryName(Form1.appPath) & "\template\" & TextBox3.Text & ".dat", Encoding.GetEncoding("shift-jis"))
                dgv = DGV8
            ElseIf k = 3 Then
                fStrArray = File.ReadAllLines(Path.GetDirectoryName(Form1.appPath) & "\template\" & TextBox1.Text & ".dat", Encoding.GetEncoding("shift-jis"))
                dgv = DGV9
            ElseIf k = 4 Then
                fStrArray = File.ReadAllLines(Path.GetDirectoryName(Form1.appPath) & "\template\" & TextBox42.Text & ".dat", Encoding.GetEncoding("shift-jis"))
                dgv = DGV13
            ElseIf k = 5 Then
                fStrArray = File.ReadAllLines(Path.GetDirectoryName(Form1.appPath) & "\template\" & TextBox14.Text & ".dat", Encoding.GetEncoding("shift-jis"))
                dgv = TMSDGV
            End If

            For i As Integer = 0 To fStrArray.Length - 1
                Dim res As String() = Split(fStrArray(i), ",")
                dgv.Columns.Add(i, res(0))
                dgv.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
            Next
        Next

        'サイズ容量読み込み
        DGV10.Rows.Clear()
        fName = Path.GetDirectoryName(Form1.appPath) & "\config\version2\サイズ容量.txt"
        csvRecords = TM_CSV_READ(fName)(0)
        For r As Integer = 0 To csvRecords.Count - 1
            Dim sArray As String() = Split(csvRecords(r), "|=|")
            DGV10.Rows.Add(sArray)
        Next

        fName = Path.GetDirectoryName(Form1.appPath) & "\config\denpyoTaihi.csv"
        'kind,基本,e飛伝2,佐川BIZ,日本郵便,定形外,e飛伝pro,TMS
        Dim koumokuLines As String() = File.ReadAllLines(fName, ENC_SJ)
        For i As Integer = 0 To koumokuLines.Length - 1
            Dim kl As String() = Split(koumokuLines(i), ",")
            koumoku(kl(0)) = {kl(1), kl(2), kl(3), kl(4), kl(5), kl(6)}
        Next
        Dim cc = koumoku
        fName = Path.GetDirectoryName(Form1.appPath) & "\config\ヘルプdenpyo3.dat"
        Dim helpLines As String() = File.ReadAllLines(fName, ENC_SJ)
        For i As Integer = 0 To helpLines.Length - 1
            helpArray.Add(helpLines(i))
        Next

        If Regex.IsMatch(My.Computer.Name, "takashi", RegexOptions.IgnoreCase) Then
            CheckBox16.Checked = True
        End If

        '******************************************
        '初期設定
        '******************************************
        KryptonComboBox1.SelectedIndex = 0

        Application.DoEvents()

        Dim localUpdateTime As String = Nothing
        Dim serverUpdateTime As String = Nothing
        '佐川DBデータベース更新確認
        If File.Exists(localTimePath) Then
            localUpdateTime = File.ReadAllText(localTimePath, encSJ)
            佐川DBアップデートToolStripMenuItem.Text = "佐川DBアップデート(" & Format(CDate(localUpdateTime), "yyyy/MM/dd HH:mm") & ")"
            If File.Exists(serverTimePath) Then
                serverUpdateTime = File.ReadAllText(serverTimePath, encSJ)
                If CDate(localUpdateTime) < CDate(serverUpdateTime) Then
                    MsgBox("佐川DBデータベースの更新をします。" & vbCrLf & "しばらくお待ちください。")
                    SagawaUpdate(serverUpdateTime)
                End If
            End If
        Else
            佐川DBアップデートToolStripMenuItem.Text = "佐川DBアップデート(sagawa.dat)"
        End If

        '住所DBデータベース更新確認
        If File.Exists(kenallLocalTimePath) Then
            localUpdateTime = File.ReadAllText(kenallLocalTimePath, encSJ)
            住所DBアップデートToolStripMenuItem.Text = "住所DBアップデート(" & Format(CDate(localUpdateTime), "yyyy/MM/dd HH:mm") & ")"
            If File.Exists(serverTimePath) Then
                serverUpdateTime = File.ReadAllText(kenallServerTimePath, encSJ)
                If CDate(localUpdateTime) < CDate(serverUpdateTime) Then
                    MsgBox("住所DBデータベースの更新をします。" & vbCrLf & "しばらくお待ちください。")
                    AddressUpdate(serverUpdateTime)
                End If
            End If
        Else
            住所DBアップデートToolStripMenuItem.Text = "佐川DBアップデート(ken_all.dat)"
        End If




    End Sub


    '行番号を表示する
    Private Sub DataGridView_RowPostPaint(ByVal sender As DataGridView, ByVal e As Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles _
            DGV1.RowPostPaint, DGV3.RowPostPaint, DGV4.RowPostPaint,
            DGV6.RowPostPaint, DGV7.RowPostPaint, DGV8.RowPostPaint, DGV9.RowPostPaint,
            DGV12.RowPostPaint, DGV13.RowPostPaint, DGV14.RowPostPaint, DGV15.RowPostPaint,
            DGV16.RowPostPaint, DGV17.RowPostPaint, DGV18.RowPostPaint
        ' 行ヘッダのセル領域を、行番号を描画する長方形とする（ただし右端に4ドットのすき間を空ける）
        Dim rect As New Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, sender.RowHeadersWidth - 4, sender.Rows(e.RowIndex).Height)
        ' 上記の長方形内に行番号を縦方向中央＆右詰で描画する　フォントや色は行ヘッダの既定値を使用する
        TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), sender.RowHeadersDefaultCellStyle.Font, rect, sender.RowHeadersDefaultCellStyle.ForeColor, TextFormatFlags.VerticalCenter Or TextFormatFlags.Right)
    End Sub

    Public FocusDGV As DataGridView = DGV1
    Private Sub DataGridView_GotFocus(sender As Object, e As EventArgs) Handles _
            DGV1.GotFocus, DGV3.GotFocus, DGV4.GotFocus,
            DGV6.GotFocus, DGV7.GotFocus, DGV8.GotFocus, DGV9.GotFocus,
            DGV12.GotFocus, DGV13.GotFocus, DGV15.GotFocus,
            DGV16.GotFocus, DGV17.GotFocus, DGV18.GotFocus
        FocusDGV = sender
    End Sub

    'ファイルのドラッグドロップ
    Private Sub DataGridView_DragDrop(ByVal sender As DataGridView, ByVal e As Windows.Forms.DragEventArgs) Handles _
             DGV1.DragDrop, DGV7.DragDrop, DGV8.DragDrop, DGV9.DragDrop,
             DGV3.DragDrop, DGV12.DragDrop, DGV13.DragDrop

        Dim dataPathArray As New ArrayList
        For Each filename As String In e.Data.GetData(DataFormats.FileDrop)
            dataPathArray.Add(filename)
        Next

        '-------------------------------------------
        'ファイルダイアログのテスト中
        '-------------------------------------------
        'If dataPathArray.Count = 0 Then
        '    Exit Sub
        'ElseIf dataPathArray.Count = 1 Then
        '    FileDataRead(dataPathArray)
        'Else
        '    Csv_denpyo3_F_files.DGV1.Rows.Clear()
        '    Csv_denpyo3_F_files.dataPathArray = dataPathArray
        '    For i As Integer = 0 To dataPathArray.Count - 1
        '        Csv_denpyo3_F_files.DGV1.Rows.Add("", Path.GetFileName(dataPathArray(i)))
        '    Next
        '    Dim DR As DialogResult = Csv_denpyo3_F_files.ShowDialog
        '    If DR = DialogResult.OK Then
        '        FileDataRead(dataPathArray)
        '    Else
        '        Exit Sub
        '    End If
        'End If
        '-------------------------------------------

        If dataPathArray.Count > 0 Then
            For i As Integer = 0 To dataPathArray.Count - 1
                DGV18.Rows.Add("", Path.GetFileName(dataPathArray(i)))
            Next
            'TabControl3.SelectTab("TabPage22")
            'TabControl9.SelectTab("TabPage33")
        End If

        '######################################################################################################################
        FileDataRead(dataPathArray)
        '######################################################################################################################
    End Sub

    Private Sub DataGridView_DragEnter(ByVal sender As Object, ByVal e As Windows.Forms.DragEventArgs) Handles _
            DGV1.DragEnter, DGV3.DragEnter, DGV7.DragEnter, DGV8.DragEnter, DGV9.DragEnter,
            DGV12.DragEnter, DGV13.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    '現在のデータで再読込
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Dim bkDir As String = Path.GetDirectoryName(appPath)
        Dim dgv1path As String = bkDir & "\PhoneBackUp\dgv1.csv"
        If File.Exists(dgv1path) Then
            File.Delete(dgv1path)
        End If
        Dim dgv3path As String = bkDir & "\PhoneBackUp\dgv3.csv"
        If File.Exists(dgv3path) Then
            File.Delete(dgv3path)
        End If

        Dim dataPathArray As New ArrayList
        If DGV1.RowCount > 0 Then
            DGV_TO_CSV_SAVE(dgv1path, DGV1, "SHIFT-JIS", True)
            DGV1.Rows.Clear()
            dataPathArray.Add(dgv1path)
        End If
        If DGV3.RowCount > 0 Then
            DGV_TO_CSV_SAVE(dgv3path, DGV3, "SHIFT-JIS", True)
            DGV3.Rows.Clear()
            dataPathArray.Add(dgv3path)
        End If

        FileDataRead(dataPathArray)
    End Sub

    'ドロップデータの認識
    Dim fileHaisou As String = ""
    Public Sub FileDataRead(dataPathArray As ArrayList)
        Me.Activate()
        Label9.Visible = True
        ToolStripStatusLabel2.Text = "[情報]"
        ListBox1.Items.Clear()
        ListBox3.Items.Clear()
        DGV17.Rows.Clear()
        KryptonButton1.Text = "調べる"

        Dim dgv As DataGridView = DGV1
        Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)

        Dim allAdd As Boolean = False   '読み込んだファイルを全て追加するか
        Dim dgv1NoArray As New ArrayList    '重複確認用Array
        Dim dgv1ErrorArray As New ArrayList
        Dim dgv3NoArray As New ArrayList
        Dim dgv3ErrorArray As New ArrayList
        If DGV1.RowCount > 0 Then
            For r As Integer = 0 To DGV1.RowCount - 1
                Dim no As String = DGV1.Item(dH1.IndexOf("伝票番号"), r).Value & "," & DGV1.Item(dH1.IndexOf("受注番号"), r).Value
                dgv1NoArray.Add(no)
            Next
        End If
        If DGV3.RowCount > 0 Then
            Dim dH3 As ArrayList = TM_HEADER_GET(DGV3)
            For r As Integer = 0 To DGV3.RowCount - 1
                Dim no As String = DGV3.Item(dH3.IndexOf("伝票番号"), r).Value & "," & DGV3.Item(dH3.IndexOf("明細行"), r).Value
                dgv3NoArray.Add(no)
            Next
        End If

        'テンプレート分析用
        Dim fBunsekiArray As String() = File.ReadAllLines(Path.GetDirectoryName(Form1.appPath) & "\template\" & TextBox41.Text & ".dat", ENC_SJ)
        Dim fBnameArray As New ArrayList
        Dim fBArray As New ArrayList
        Dim fBHeader As String() = Split(fBunsekiArray(0), ",")
        For i As Integer = 0 To fBHeader.Length - 1
            Dim str As String = ""
            For k As Integer = 0 To fBunsekiArray.Length - 1
                Dim fBline As String() = Split(fBunsekiArray(k), ",")
                If k = 0 Then
                    fBnameArray.Add(fBline(i))
                Else
                    If fBline(i) <> "" Then
                        str &= fBline(i) & ","
                    Else
                        Exit For
                    End If
                End If
            Next
            str = str.TrimEnd(",")
            fBArray.Add(str)
        Next

        Dim addFlag As Boolean = False
        dataPathArray.Sort()
        dataPathArray.Reverse()
        For Each filename As String In dataPathArray
            Dim fn As String = Path.GetFileName(filename)
            ToolStripMenuItem2.DropDownItems.Add(filename)

            Dim csvRecords As New ArrayList()
            csvRecords = TM_CSV_READ(filename)(0)

            'ファイル認識で分別する
            Dim cHeader As String = Replace(csvRecords(0), "|=|", ",")
            Dim fB As String = ""
            For i As Integer = 0 To fBArray.Count - 1
                If fBArray(i) = cHeader Then
                    fB = fBnameArray(i)
                    Exit For
                End If
            Next

            Dim gamen1 As String = "画面"
            Dim gamen2 As String = "画面not"
            If fB = "画面" Then
                Dim DR As DialogResult = MsgBox(fn & vbCrLf & "情報の追加データなら「はい」" & vbNewLine & "元データとして読込なら「いいえ」", MsgBoxStyle.YesNo Or MsgBoxStyle.SystemModal)
                If DR = DialogResult.Yes Then
                    gamen1 = "画面not"
                    gamen2 = "画面"
                End If
            End If

            Select Case fB
                Case "元データ", "佐川", "日本郵便", "定形外", "ヤマト", gamen1
                    addFlag = False
                    dgv = DGV1
                    TabControl2.SelectedIndex = 0
                    DGV7.Rows.Clear()
                    DGV8.Rows.Clear()
                    DGV9.Rows.Clear()
                    DGV13.Rows.Clear()
                Case "明細"
                    addFlag = False
                    dgv = DGV3
                    TabControl5.SelectedIndex = 0
                    DGV12.Rows.Clear()
                    DGV12.Columns.Clear()
                Case "追加", "出荷指示", "カスタムデータ追加", gamen2
                    addFlag = True
                    dgv = DGV1
                    TabControl2.SelectedIndex = 0
                Case Else
                    MsgBox(filename & vbCrLf & "ファイルの認識ができません。")
                    Continue For
            End Select

            '伝票確認用
            If fB = "画面" Then
                Kakunin(filename)
            End If

            Dim dH As ArrayList = TM_HEADER_GET(dgv)

            '追加または削除
            If allAdd = False And addFlag = False Then
                If dgv.RowCount > 0 Then
                    Dim DR As DialogResult = MsgBox("ファイルを全て追加しますか", MsgBoxStyle.YesNo Or MsgBoxStyle.SystemModal)
                    If DR = Windows.Forms.DialogResult.No Then
                        dgv.Rows.Clear()
                        dgv1NoArray.Clear()
                        dgv3NoArray.Clear()
                    ElseIf DR = Windows.Forms.DialogResult.Yes Then
                        allAdd = True
                    End If
                End If
            End If

            '------------------------------------------------------------------
            '根据文件名字决定发送方式
            Select Case True
                Case Regex.IsMatch(filename, "sagawaMail|yupacketCSV|ymailCSV|yamatonekopos", RegexOptions.IgnoreCase)
                    fileHaisou = "メール便"
                Case Regex.IsMatch(filename, "yucreCSV", RegexOptions.IgnoreCase)
                    fileHaisou = "定形外"
                Case Regex.IsMatch(filename, "sagawa_ePro")
                    fileHaisou = "宅配便"
                    'Case Regex.IsMatch(filename, "ymailCSV")
                    '    fileHaisou = "ヤマト"

                Case Else
                    fileHaisou = "宅配便"
            End Select


            ToolStripProgressBar1.Maximum += csvRecords.Count

            Dim fStrArray As String() = Nothing
            If dgv Is DGV1 Then
                Dim cc = Path.GetDirectoryName(Form1.appPath) & "\template\" & TextBox8.Text
                fStrArray = File.ReadAllLines(Path.GetDirectoryName(Form1.appPath) & "\template\" & TextBox8.Text & ".dat", Encoding.GetEncoding("shift-jis"))
            Else
                fStrArray = Split(csvRecords(0), "|=|")
            End If

            'dgv1は対応する列のみ読み込む
            Dim header As String() = Split(csvRecords(0), "|=|")
            Dim dNo As String = ""

            For r As Integer = 1 To csvRecords.Count - 1
                Dim sArray As String() = Split(csvRecords(r), "|=|")
                Dim changeRowNo As Integer = -1     '追加修正行

                If addFlag = False Then
                    dgv.Rows.Add()
                Else
                    dNo = sArray(Array.IndexOf(header, "伝票番号"))
                    'dgv1の修正行を探す
                    For r1 As Integer = 0 To DGV1.RowCount - 1
                        If DGV1.Item(dH.IndexOf("伝票番号"), r1).Value = dNo Then
                            changeRowNo = r1
                            Exit For
                        End If
                    Next
                    If changeRowNo = -1 Then
                        Continue For
                    End If
                End If

                Dim siharai As String = ""
                Dim goukei As String = ""
                For c As Integer = 0 To dgv.ColumnCount - 1
                    'fStrArrayを「,」で区切り、ヘッダーを探す
                    Dim hA As String() = Split(fStrArray(c), ",")   '共通-基本データから読み出し
                    For i As Integer = 0 To hA.Length - 1
                        If hA(i) <> "" Then
                            If header.Contains(hA(i)) Then
                                Select Case hA(i)
                                    Case "受注者電話番号", "発送先電話番号", "お届け先電話", "届け先電話番号"
                                        Dim str As String = sArray(Array.IndexOf(header, hA(i)))
                                        str = Replace(str, "+", "")
                                        dgv.Item(c, dgv.RowCount - 1).Value = str
                                    Case "配送日", "配達希望日", "配達指定日"
                                        Dim str As String = sArray(Array.IndexOf(header, hA(i)))
                                        If InStr(sArray(Array.IndexOf(header, hA(i))), ":") > 0 Then
                                            str = Format(CDate(str), "yyyyMMdd")
                                        End If
                                        If addFlag Then '追加
                                            If CStr(dgv.Item(c, changeRowNo).Value) = "" Then
                                                dgv.Item(c, changeRowNo).Value = str
                                            End If
                                        Else
                                            dgv.Item(c, dgv.RowCount - 1).Value = str
                                        End If
                                    Case "発送方法"
                                        Dim str As String = sArray(Array.IndexOf(header, hA(i)))
                                        Select Case True
                                            Case Regex.IsMatch(str, "佐川")
                                                str = "宅配便"
                                            Case Regex.IsMatch(str, "ゆうパケット|メール便")
                                                str = "メール便"
                                            Case Regex.IsMatch(str, "定形外")
                                                str = "定形外"
                                            Case Regex.IsMatch(str, "ヤマト(DM便)|ヤマト(ネコポス)")
                                                'str = "ヤマト"
                                                str = "メール便"
                                        End Select
                                        If addFlag Then '追加
                                            If CStr(dgv.Item(c, changeRowNo).Value) = "" Then
                                                dgv.Item(c, changeRowNo).Value = str
                                            End If
                                        Else
                                            dgv.Item(c, dgv.RowCount - 1).Value = str
                                        End If
                                    Case "支払方法"
                                        siharai = sArray(Array.IndexOf(header, hA(i)))
                                        siharaiYU2 = siharai
                                        If addFlag Then '追加
                                            If CStr(dgv.Item(c, changeRowNo).Value) = "" Then
                                                dgv.Item(c, changeRowNo).Value = siharai
                                            End If
                                        Else
                                            dgv.Item(c, dgv.RowCount - 1).Value = siharai
                                        End If

                                        If header.Contains("支払方法") Then
                                            Dim str As String = ""
                                            If header.Contains("合計金額") Then
                                                str = sArray(Array.IndexOf(header, "合計金額"))
                                            Else
                                                str = sArray(Array.IndexOf(header, "総合計"))
                                            End If
                                            goukei = CStr(CInt(str))
                                        End If
                                    Case "合計金額"
                                        If addFlag Then '追加
                                            If CStr(dgv.Item(c, changeRowNo).Value) = "" Then
                                                If goukei <> "" Then
                                                    dgv.Item(c, changeRowNo).Value = goukei
                                                Else
                                                    dgv.Item(c, changeRowNo).Value = sArray(Array.IndexOf(header, hA(i)))
                                                End If
                                            End If
                                        Else
                                            If goukei <> "" Then
                                                dgv.Item(c, dgv.RowCount - 1).Value = goukei
                                            Else
                                                dgv.Item(c, dgv.RowCount - 1).Value = sArray(Array.IndexOf(header, hA(i)))
                                            End If
                                        End If
                                    Case "総合計"  '画面csv用
                                        If addFlag Then '追加
                                            If CStr(dgv.Item(c, changeRowNo).Value) = "" Then
                                                If goukei <> "" And siharai = "代金引換" Then
                                                    dgv.Item(c, changeRowNo).Value = goukei
                                                Else
                                                    dgv.Item(c, changeRowNo).Value = "0"
                                                End If
                                            End If
                                        Else
                                            If goukei <> "" And siharai = "代金引換" Then
                                                dgv.Item(c, dgv.RowCount - 1).Value = goukei
                                            Else
                                                dgv.Item(c, dgv.RowCount - 1).Value = "0"
                                            End If
                                        End If
                                    Case "配送時間帯", "時間指定", "配達時間帯区分"
                                        Dim str As String = sArray(Array.IndexOf(header, hA(i)))
                                        Select Case str
                                            Case "01"
                                                str = "午前中"
                                            Case "12", "１２時～１４時"
                                                str = "12時-14時"
                                            Case "14", "１４時～１６時"
                                                str = "14時-16時"
                                            Case "16", "１６時～１８時"
                                                str = "16時-18時"
                                            Case "18", "１８時～２０時"
                                                str = "18時-20時"
                                            Case "19", "１９時～２１時"
                                                str = "19時-21時"
                                            Case "04", "１８時～２１時"
                                                str = "18時-21時"
                                        End Select
                                        If addFlag Then '追加
                                            If dgv.Item(c, changeRowNo).Value = "" Then
                                                dgv.Item(c, changeRowNo).Value = str
                                            End If
                                        Else
                                            dgv.Item(c, dgv.RowCount - 1).Value = str
                                        End If
                                    Case "発送先住所1", "送り先住所1"
                                        Dim addr1 As String = "発送先住所1"
                                        Dim addr2 As String = "発送先住所2"
                                        If hA(i) = "送り先住所1" Then
                                            addr1 = "送り先住所1"
                                            addr2 = "送り先住所2"
                                        End If
                                        If addFlag Then '追加
                                            If CStr(dgv.Item(c, changeRowNo).Value) = "" Then
                                                dgv.Item(c, changeRowNo).Value = sArray(Array.IndexOf(header, addr1))
                                                dgv.Item(c, changeRowNo).Value &= sArray(Array.IndexOf(header, addr2))
                                            End If
                                        Else
                                            dgv.Item(c, dgv.RowCount - 1).Value = sArray(Array.IndexOf(header, addr1))
                                            dgv.Item(c, dgv.RowCount - 1).Value &= sArray(Array.IndexOf(header, addr2))
                                        End If
                                    Case "営業所止め", "営業所止", "営業所止置き"
                                        Dim str As String = "0"
                                        Dim eidome As String = "営業所止め"
                                        If hA(i) = "営業所止" Then
                                            eidome = "営業所止"
                                        End If
                                        If hA(i) = "営業所止置き" Then
                                            eidome = "営業所止置き"
                                        End If
                                        If sArray(Array.IndexOf(header, eidome)) = "" Then
                                            str = "0"
                                        Else
                                            str = "1"
                                        End If
                                        If addFlag Then '追加
                                            If CStr(dgv.Item(c, changeRowNo).Value) = "" Then
                                                dgv.Item(c, changeRowNo).Value = str
                                            End If
                                        Else
                                            dgv.Item(c, dgv.RowCount - 1).Value = str
                                        End If
                                    Case "不足"
                                        dgv.Item(c, dgv.RowCount - 1).Style.BackColor = Color.Gray
                                    Case Else
                                        If addFlag Then '追加
                                            If CStr(dgv.Item(c, changeRowNo).Value) = "" Then
                                                Dim str As String = sArray(Array.IndexOf(header, hA(i)))
                                                If str <> "" Then
                                                    dgv.Item(c, changeRowNo).Style.BackColor = Color.Empty
                                                End If
                                                dgv.Item(c, changeRowNo).Value = sArray(Array.IndexOf(header, hA(i)))
                                            End If
                                        Else
                                            dgv.Item(c, dgv.RowCount - 1).Value = sArray(Array.IndexOf(header, hA(i)))
                                        End If
                                End Select
                            End If
                        End If
                    Next
                Next

                If dgv Is DGV1 Then
                    dgv.Item(dH.IndexOf("データ"), dgv.RowCount - 1).Value = fB
                    If CStr(dgv.Item(dH.IndexOf("発送方法"), dgv.RowCount - 1).Value) = "" Then
                        dgv.Item(dH.IndexOf("発送方法"), dgv.RowCount - 1).Value = fileHaisou
                    End If
                End If

                '伝票番号で重複検索
                If dgv Is DGV1 And addFlag = False Then
                    '伝票番号と受注番号が既にあるなら読み込まない
                    Dim no As String = dgv.Item(dH.IndexOf("伝票番号"), dgv.RowCount - 1).Value & "," & dgv.Item(dH.IndexOf("受注番号"), dgv.RowCount - 1).Value
                    If dgv1NoArray.Contains(no) Then
                        dgv.Item(0, dgv.RowCount - 1).Style.BackColor = Color.Orange
                        dgv1ErrorArray.Add(no)
                    Else
                        dgv1NoArray.Add(no)
                    End If
                ElseIf dgv Is DGV3 Then
                    '伝票番号と明細行が既にあるなら読み込まない
                    Dim no As String = dgv.Item(dH.IndexOf("伝票番号"), dgv.RowCount - 1).Value & "," & dgv.Item(dH.IndexOf("明細行"), dgv.RowCount - 1).Value
                    If dgv3NoArray.Contains(no) Then
                        dgv.Item(0, dgv.RowCount - 1).Style.BackColor = Color.Orange
                        dgv3ErrorArray.Add(no)
                    Else
                        dgv3NoArray.Add(no)
                    End If
                End If

                ToolStripProgressBar1.Value += 1
            Next

            If dgv1ErrorArray.Count > 0 Or dgv3ErrorArray.Count > 0 Then
                Dim dgv1ErrorCode = ""
                If dgv1ErrorArray.Count > 0 Then
                    For r As Integer = 0 To dgv1ErrorArray.Count - 1
                        Dim dgv1ErrorArray_ As String() = Split(dgv1ErrorArray(r), ",")
                        If r = 0 Then
                            dgv1ErrorCode = dgv1ErrorArray_(0)
                        Else
                            dgv1ErrorCode = dgv1ErrorCode & "," & dgv1ErrorArray_(0)
                        End If
                    Next
                End If

                Dim dgv3ErrorCode = ""
                If dgv3ErrorArray.Count > 0 Then
                    For r As Integer = 0 To dgv3ErrorArray.Count - 1
                        Dim dgv3ErrorArray_ As String() = Split(dgv3ErrorArray(r), ",")
                        If r = 0 Then
                            dgv3ErrorCode = dgv3ErrorArray_(0)
                        Else
                            dgv3ErrorCode = dgv3ErrorCode & "," & dgv3ErrorArray_(0)
                        End If
                    Next
                    'dgv3ErrorCode = dgv3ErrorCode.Remove(dgv3ErrorCode.Length - 1)
                End If

                Dim mes As String = ""

                mes &= "元データ：" & dgv1ErrorArray.Count & "件" & vbCrLf
                If dgv1ErrorArray.Count > 0 Then
                    mes &= "重複伝票番号:" & dgv1ErrorCode & vbCrLf
                End If

                mes &= "明細データ：" & dgv3ErrorArray.Count & "件" & vbCrLf
                If dgv3ErrorArray.Count > 0 Then
                    mes &= "重複伝票番号:" & dgv3ErrorCode & vbCrLf
                Else

                End If

                mes &= "重複しています。" & vbCrLf & "重複発送や商品を余計に発送する可能性があるので、データを取り直してください。"
                LockFlag = True
                MsgBox(mes, MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation Or MsgBoxStyle.SystemModal)
            End If
            '------------------------------------------------------------------

            'ヘッダーチェック
            HeaderCheckCSV(dgv, dH)
        Next

        Application.DoEvents()

        '背景赤のセルを選択する
        Do
            For r As Integer = 0 To DGV1.RowCount - 1
                For c As Integer = 0 To DGV1.ColumnCount - 1
                    If DGV1.Item(c, r).Style.BackColor = Color.Red Then
                        dgv.CurrentCell = dgv(c, r)
                        Exit Do
                    End If
                Next
            Next
            Exit Do
        Loop

        If Panel8.Visible = False And Panel3.Visible = False And Panel7.Visible = False Then
            Label9.Visible = True
            Application.DoEvents()

            If MasterInput() Then
                '商品マスターに存在するかチェック（location.csvに問題が無いか調べる）
                For r1 As Integer = 0 To DGV1.RowCount - 1
                    Dim mCode As String = DGV1.Item(dH1.IndexOf("商品マスタ"), r1).Value
                    Dim mCodeArray As String() = Split(mCode, "、")
                    For i As Integer = 0 To mCodeArray.Length - 1
                        Dim code As String() = Split(mCodeArray(i), "*")
                        If Not dgv6CodeArray.Contains(code(0).ToLower) Then
                            MsgBox("倉庫管理の登録データに問題があります。商品担当に連絡してください。" & vbCrLf & code(0).ToLower)
                            For c As Integer = 0 To DGV1.ColumnCount - 1
                                DGV1.Item(c, r1).Style.BackColor = Color.Red
                            Next
                            DGV1.CurrentCell = DGV1(dH1.IndexOf("商品マスタ"), r1)
                            'Panel4.Visible = False
                            'Exit Sub
                        End If
                    Next
                Next

                Label78.Text = "上記内代引き人数 " & DaibikiCount(DGV1, "all") & "人"      '代引きカウント
                '修改入口
                MasterHaisouKeisan(0, -1, DGV1, 0, 0, "")  'マスタ配送の計算

                SagawaKoukuuBin(0, 0)   '佐川航空便チェック
                SagyoranGet()           '作業用欄取得
                YoyakuGet()             '予約状態取得



                ''やまとへんかん
                'Dim isYamatoGood As Boolean = True
                'Dim isYamatoGood_fukumu As Boolean = False '包含yamato的货


                For r1 As Integer = 0 To DGV1.RowCount - 1
                    Dim temp = DGV1.Item(dH1.IndexOf("マスタ配送"), r1).Value
                    Dim temp2 = DGV1.Item(dH1.IndexOf("発送方法"), r1).Value

                    Dim c = (DGV1.Item(dH1.IndexOf("発送倉庫"), r1).Value = dazaifu_str Or DGV1.Item(dH1.IndexOf("発送倉庫"), r1).Value = "井相田")

                    If temp = "ヤマト" And temp2 = "ヤマト" And c Then
                        DGV1.Item(dH1.IndexOf("マスタ配送"), r1).Value = "ヤマト(陸便)"
                        DGV1.Item(dH1.IndexOf("発送方法"), r1).Value = yamato_str
                        DGV1.Item(dH1.IndexOf("データ"), r1).Value = "画面"
                        If DGV1.Item(dH1.IndexOf("元便種"), r1).Value <> "陸便" Then
                            DGV1.Item(dH1.IndexOf("マスタ配送"), r1).Value = "ヤマト(船便)"
                            DGV1.Item(dH1.IndexOf("便種"), r1).Value = "船便"
                        End If
                    End If
                Next


                If False Then
                    For r1 As Integer = 0 To DGV1.RowCount - 1
                        'has_ny263_50_c = False
                        'ny263_50_c_count = 0
                        'DGV1.Item(dH1.IndexOf("元便種"), r1).Value = "陸便"
                        Dim mCode As String = DGV1.Item(dH1.IndexOf("商品マスタ"), r1).Value
                        Dim mCodeArray As String() = Split(mCode, "、")


                        Dim isYamatoGood As Boolean = False
                        Dim isYamatoGood_fukumu As Boolean = False
                        For checkcodei As Integer = 0 To mCodeArray.Length - 1
                            Dim checkcode As String() = Split(mCodeArray(checkcodei), "*")
                            'If Not (yamato_goods.Contains(checkcode(0).ToLower)) Then
                            '    isYamatoGood = False
                            'End If

                            'If Not (yamato_goods.Contains(checkcode(0).ToLower) And Not InStr(checkcode(0).ToLower, "ny393-50-") And Not InStr(checkcode(0).ToLower, "ap005") And Not InStr(checkcode(0).ToLower, "ap039")) Then
                            '    isYamatoGood = False
                            'End If





                            '这里是yamato旧的逻辑
                            If yamato_goods.Contains(checkcode(0).ToLower) Or InStr(checkcode(0).ToLower, "ny393-50-") Or InStr(checkcode(0).ToLower, "ap005") Or InStr(checkcode(0).ToLower, "ap039") Or InStr(checkcode(0).ToLower, "ee270") Or InStr(checkcode(0).ToLower, "zk218") Then
                                isYamatoGood = True
                            End If

                            If yamato_goods.Contains(checkcode(0).ToLower) Then
                                If isYamatoGood_fukumu = False Then
                                    isYamatoGood_fukumu = True
                                End If
                            End If


                            'yamato的最后判断
                            'If DGV1.Item(dH1.IndexOf("マスタ配送"), r1).Value = "ヤマト" Then
                            '    isYamatoGood = True
                            'End If



                        Next
                        For checkcodei As Integer = 0 To mCodeArray.Length - 1
                            Dim checkcode As String() = Split(mCodeArray(checkcodei), "*")
                            'isyupakuGoodBool = isyupakuGoodsbyCode(checkcode(0))
                        Next
                        Dim haisouSaki As String = DGV1.Item(dH1.IndexOf("発送先住所"), r1).Value
                        If isYamatoGood And (DGV1.Item(dH1.IndexOf("発送倉庫"), r1).Value = dazaifu_str Or DGV1.Item(dH1.IndexOf("発送倉庫"), r1).Value = "井相田") And (Regex.IsMatch(DGV1.Item(dH1.IndexOf("発送方法"), r1).Value, "メール便") Or Regex.IsMatch(DGV1.Item(dH1.IndexOf("発送方法"), r1).Value, "ヤマト")) Then
                            'If DGV1.Item(dH1.IndexOf("sw"), r1).Value <> Nothing And ny263_50_c_sw <> Nothing Then
                            '    If DGV1.Item(dH1.IndexOf("マスタ配送"), r1).Value = "宅配便" Then
                            '        DGV1.Item(dH1.IndexOf("sw"), r1).Value = Math.Ceiling((DGV1.Item(dH1.IndexOf("sw"), r1).Value * 100) - (ny263_50_c_sw * ny263_50_c_count) + (25 * ny263_50_c_count))
                            '    Else
                            '        DGV1.Item(dH1.IndexOf("sw"), r1).Value = Math.Ceiling(DGV1.Item(dH1.IndexOf("sw"), r1).Value - (ny263_50_c_sw * ny263_50_c_count) + (25 * ny263_50_c_count))
                            '    End If
                            '    DGV1.Item(dH1.IndexOf("マスタ便数"), r1).Value = Math.Ceiling(DGV1.Item(dH1.IndexOf("sw"), r1).Value / 100)
                            'End If

                            DGV1.Item(dH1.IndexOf("マスタ配送"), r1).Value = "ヤマト(陸便)"
                            DGV1.Item(dH1.IndexOf("発送方法"), r1).Value = yamato_str
                            DGV1.Item(dH1.IndexOf("データ"), r1).Value = "画面"

                            If DGV1.Item(dH1.IndexOf("元便種"), r1).Value <> "陸便" Then
                                DGV1.Item(dH1.IndexOf("マスタ配送"), r1).Value = "ヤマト(船便)"
                                DGV1.Item(dH1.IndexOf("便種"), r1).Value = "船便"
                            End If
                            'ElseIf Regex.IsMatch(DGV1.Item(dH1.IndexOf("発送方法"), r1).Value, "宅配便") Then
                            '    If DGV1.Item(dH1.IndexOf("sw"), r1).Value <> Nothing And ny263_50_c_sw <> Nothing Then
                            '        DGV1.Item(dH1.IndexOf("sw"), r1).Value = Math.Ceiling((DGV1.Item(dH1.IndexOf("sw"), r1).Value * 100) - (ny263_50_c_sw * ny263_50_c_count) + (2.5 * ny263_50_c_count))
                            '        DGV1.Item(dH1.IndexOf("マスタ便数"), r1).Value = Math.Ceiling(DGV1.Item(dH1.IndexOf("sw"), r1).Value / 100)
                            '    End If
                        ElseIf isYamatoGood_fukumu = True And isYamatoGood = False And Regex.IsMatch(DGV1.Item(dH1.IndexOf("発送方法"), r1).Value, "メール便") Then 'ヤマト和メール一起的时候改成佐川
                            ' isYamatoGood_fukumu = True And isYamatoGood = False 不是全是yamato的货，但是包含
                            DGV1.Item(dH1.IndexOf("発送方法"), r1).Value = "宅配便"
                            DGV1.Item(dH1.IndexOf("マスタ配送"), r1).Value = "宅配便"
                            DGV1.Item(dH1.IndexOf("データ"), r1).Value = "佐川"

                            If IsNumeric(DGV1.Item(dH1.IndexOf("sw"), r1).Value) Then
                                DGV1.Item(dH1.IndexOf("マスタ便数"), r1).Value = Math.Ceiling(DGV1.Item(dH1.IndexOf("sw"), r1).Value / 10000)
                            End If


                            'Dim haisouSaki As String = DGV1.Item(dH1.IndexOf("発送先住所"), r1).Value

                        ElseIf CheckBox34.Checked And
                                isyupakubyAddress(haisouSaki) And isyupakuGoodBool And Regex.IsMatch(DGV1.Item(dH1.IndexOf("発送方法"), r1).Value, "宅配便") Then

                            'DGV1.Item(dH1.IndexOf("発送方法"), r1).Value = "メール便"
                            'DGV1.Item(dH1.IndexOf("マスタ配送"), r1).Value = "メール便"
                            'DGV1.Item(dH1.IndexOf("データ"), r1).Value = "ゆう200"
                            ''一定  
                            'YU2Flag = True
                        End If

                    Next
                End If
                MaisuKeisanLeft()       '枚数計算メイン左

            End If
            Label9.Visible = False
        End If

        'dgv1の伝票番号を取得保存
        dgv1dNoArray.Clear()
        For r As Integer = 0 To DGV1.RowCount - 1
            dgv1dNoArray.Add(DGV1.Item(dH1.IndexOf("伝票番号"), r).Value)
        Next
        TabPage4.Text = "元データ(" & dgv1dNoArray.Count & ")"

        ToolStripProgressBar1.Value = 0
        selChangeFlag = True

        Label9.Visible = False
    End Sub

    '伝票確認データ読込
    Private Sub Kakunin(fileName As String)
        Dim csvRecords As ArrayList = TM_CSV_READ(fileName)(0)
        Dim header As String() = Split(csvRecords(0), "|=|")
        Dim dH4 As ArrayList = TM_HEADER_GET(DGV4)

        For r As Integer = 1 To csvRecords.Count - 1
            Dim sArray As String() = Split(csvRecords(r), "|=|")
            DGV4.Rows.Add()
            Dim rownum As Integer = DGV4.RowCount - 1
            DGV4.Item(dH4.IndexOf("伝票番号"), rownum).Value = sArray(Array.IndexOf(header, "伝票番号"))
            DGV4.Item(dH4.IndexOf("店舗"), rownum).Value = sArray(Array.IndexOf(header, "店舗"))
            DGV4.Item(dH4.IndexOf("送り先名"), rownum).Value = sArray(Array.IndexOf(header, "送り先名"))
            DGV4.Item(dH4.IndexOf("送り先住所"), rownum).Value = sArray(Array.IndexOf(header, "送り先住所"))
        Next
    End Sub

    Private Sub HeaderCheckCSV(dgv As DataGridView, dH As ArrayList)
        'ヘッダーチェック
        Dim checkHeader As String() = Nothing
        Dim checkHeader2 As String() = Nothing
        Dim panelShow As Panel = Nothing
        If dgv Is DGV1 Then
            checkHeader = New String() {"伝票番号", "受注番号", "購入者名", "発送先名", "発送先郵便番号", "発送先住所"}
            checkHeader2 = New String() {"購入者電話番号", "発送先電話番号"}
            panelShow = Panel8
        Else
            checkHeader = New String() {"店舗", "購入者名", "受注番号", "伝票番号", "商品ｺｰﾄﾞ", "商品区分"}
            panelShow = Panel3
        End If

        Dim flag As Boolean = True
        Dim flag2 As Boolean = True
        For r As Integer = 0 To dgv.RowCount - 1
            For c As Integer = 0 To dgv.ColumnCount - 1
                For i As Integer = 0 To checkHeader.Length - 1
                    Try
                        If dgv.Item(dH.IndexOf(checkHeader(i)), r).Value = "" Then
                            dgv.Item(dH.IndexOf(checkHeader(i)), r).Style.BackColor = Color.Red
                            'dgv.CurrentCell = dgv(dH.IndexOf(checkHeader(i)), r)
                            flag = False
                            Exit For
                        End If
                    Catch ex As Exception
                        flag = False
                        Exit For
                    End Try
                Next
                If checkHeader2 IsNot Nothing Then
                    For i As Integer = 0 To checkHeader2.Length - 1
                        Try
                            If CStr(dgv.Item(dH.IndexOf(checkHeader2(i)), r).Value) = "" Then
                                dgv.Item(dH.IndexOf(checkHeader2(i)), r).Style.BackColor = Color.Red
                                'dgv.CurrentCell = dgv(dH.IndexOf(checkHeader2(i)), r)
                                flag2 = False
                            End If
                        Catch ex As Exception
                        End Try
                    Next
                End If
            Next
            If flag = False Then
                Exit For
            End If
        Next

        If flag Then
            panelShow.Visible = False

            If dgv Is DGV1 Then
                If flag2 Then
                    Panel7.Visible = False
                Else
                    Panel7.Visible = True
                End If
            End If
        Else
            panelShow.Visible = True
        End If

    End Sub

    '代引きカウント
    Private Function DaibikiCount(dgv As DataGridView, souko As String)
        Dim num As Integer = 0
        If dgv Is DGV1 Then
            num = 0
        ElseIf dgv Is DGV7 Then
            num = 1
        ElseIf dgv Is DGV8 Then
            num = 2
        ElseIf dgv Is TMSDGV Then
            num = 6
        Else
            Return 0
            Exit Function
        End If
        Dim dH As ArrayList = TM_HEADER_GET(dgv)
        Dim counter As Integer = 0
        For r As Integer = 0 To dgv.RowCount - 1
            If dgv.Item(dH.IndexOf(koumoku("goukei")(num)), r).Value > 0 Then
                If souko = "all" Then
                    counter += 1
                Else
                    If dgv.Item(dH.IndexOf(koumoku("syori2")(num)), r).Value = souko Then
                        counter += 1
                    End If
                End If
            End If
        Next
        Return counter
    End Function

    '予約状態取得
    Private Sub YoyakuGet()
        Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
        Dim dH3 As ArrayList = TM_HEADER_GET(DGV3)

        ToolStripProgressBar1.Maximum = DGV1.RowCount
        ToolStripProgressBar1.Value = 0
        Application.DoEvents()

        For r As Integer = 0 To DGV1.RowCount - 1
            Dim dNum As String = DGV1.Item(dH1.IndexOf("伝票番号"), r).Value

            '明細から取得
            For r3 As Integer = DGV3.RowCount - 1 To 0 Step -1
                If dNum = DGV3.Item(dH3.IndexOf("伝票番号"), r3).Value Then
                    Dim str As String = DGV3.Item(dH3.IndexOf("商品ｵﾌﾟｼｮﾝ"), r3).Value
                    str = Regex.Replace(str, "\n|\r", "")
                    Dim matches As MatchCollection = Regex.Matches(str, "予約|出荷予定")
                    If matches.Count > 0 Then
                        DGV1.Item(dH1.IndexOf("予約"), r).Value = "予約"
                    End If
                    If DGV3.Item(dH3.IndexOf("明細行"), r3).Value = 1 Then
                        Exit For
                    End If
                End If
            Next

            ToolStripProgressBar1.Value += 1
        Next
    End Sub

    '備考取得
    Private Sub SagyoranGet()
        Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
        Dim dH3 As ArrayList = TM_HEADER_GET(DGV3)

        ToolStripProgressBar1.Maximum = DGV1.RowCount
        ToolStripProgressBar1.Value = 0
        Application.DoEvents()

        For r As Integer = 0 To DGV1.RowCount - 1
            Dim dNum As String = DGV1.Item(dH1.IndexOf("伝票番号"), r).Value

            '明細からは作業者欄が取得できる
            For r3 As Integer = 0 To DGV3.RowCount - 1
                If dNum = DGV3.Item(dH3.IndexOf("伝票番号"), r3).Value Then
                    Dim str As String = DGV3.Item(dH3.IndexOf("作業者欄"), r3).Value
                    str = Regex.Replace(str, "\n|\r", "")
                    Dim matches As MatchCollection = Regex.Matches(str, "\[\[.*?\]\]")
                    For Each m In matches
                        DGV1.Item(dH1.IndexOf("作業者欄"), r).Value &= m.ToString
                    Next
                    Exit For
                End If
            Next

            'マーク・マスタ備考生成
            Dim mark As String = ""
            Dim sagyo As String = DGV1.Item(dH1.IndexOf("作業者欄"), r).Value
            Dim Bikou1 As String = DGV1.Item(dH1.IndexOf("配送備考"), r).Value
            Dim Bikou2 As String = DGV1.Item(dH1.IndexOf("備考"), r).Value
            Dim Bikou3 As String = DGV1.Item(dH1.IndexOf("納品書備考"), r).Value
            Dim bikouAll As String = sagyo & Bikou1 & Bikou2 & Bikou3

            '★|♪|●
            If InStr(bikouAll, "★") > 0 Then    '別紙あり
                mark &= "★"
            End If
            If InStr(bikouAll, "♪") > 0 Then    'プレゼント
                mark &= "♪"
            End If
            If InStr(bikouAll, "●") > 0 Then    '急ぎ
                mark &= "●"
            End If
            If mark <> "" Then
                If DGV1.Item(dH1.IndexOf("マーク"), r).Value = "" Then
                    DGV1.Item(dH1.IndexOf("マーク"), r).Value = mark
                Else
                    If InStr(DGV1.Item(dH1.IndexOf("マーク"), r).Value, "★") > 0 And InStr(mark, "★") > 0 Then
                        DGV1.Item(dH1.IndexOf("マーク"), r).Value &= mark
                    End If
                End If
            End If

            'マスタ備考
            Dim masterBikouArray As New ArrayList
            Dim masterBikou As String = ""
            Dim matchesB As MatchCollection = Regex.Matches(bikouAll, "\[\[.*?\]\]")
            For Each m In matchesB
                Dim ms As String = Regex.Replace(m.ToString, "\[\[|\]\]", "")
                If Not masterBikouArray.Contains(ms) Then
                    masterBikouArray.Add(ms)
                End If
            Next
            For Each mb In masterBikouArray
                masterBikou &= mb & "、"
            Next
            masterBikou = masterBikou.TrimEnd("、")
            If masterBikou <> "" Then
                If DGV1.Item(dH1.IndexOf("マスタ備考"), r).Value = "" Then
                    DGV1.Item(dH1.IndexOf("マスタ備考"), r).Value = masterBikou
                Else
                    DGV1.Item(dH1.IndexOf("マスタ備考"), r).Value &= "、" & masterBikou
                End If
            End If

            ToolStripProgressBar1.Value += 1
        Next
    End Sub





    Dim codesidIsChangeYTT As String() = {"ny331-50", "ny263", "ny405", "ny429", "ny385", "ny411-60", "ny439", "ny373", "ny393-50", "ny417-30"}
    Public Function isChangeYTT(codeid As String) As Boolean
        If codesidIsChangeYTT.Contains(codeid) Then
            Return True
        End If
        Return False
    End Function

    'true 变佐川  false 不变佐川
    Public Function IsNotChangeCode(codeid As String) As Boolean
        If codeid.Contains("ny411") Then

            If Not codeid.Contains("ny411-60-") Then
                Return False
            Else
                Return True
            End If
        End If

        If codeid.Contains("ny393") Then

            'If Not codeid.Contains("ny393-50") And Not codeid.Contains("ny393-500") Then
            '    Return False
            'Else
            '    Return True
            'End If


            If codeid.Contains("ny393-50") And Not codeid.Contains("ny393-500") Then
                Return True
            Else
                Return False

            End If




        End If

        If codeid.Contains("ny331") Then

            If Not codeid.Contains("ny331-50-") Then
                Return False
            Else
                Return True
            End If
        End If

        If codeid.Contains("ny417") Then

            If Not codeid.Contains("ny417-30-") Then
                Return False
            Else
                Return True
            End If
        End If
        Return True
    End Function









    '---------------------------------------
    'BIZ-logi 便種コード
    '000 陸便
    '030 航空便R1
    '033 ジャスト便
    '140 クール便
    '141 クールR
    '150 クール凍
    '151 クールR凍
    '---------------------------------------

    '佐川航空便
    Private Sub SagawaKoukuuBin(ByVal mode As Integer, ByVal num As Integer)
        'ファイル読み取り
        Dim chiikiArray As ArrayList = New ArrayList
        Dim pmArray As ArrayList = New ArrayList
        Dim sCodeArray As ArrayList = New ArrayList
        Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\config\version2\航空便可不可.txt"
        Dim flag As Integer = 0
        Using sr As New StreamReader(fName, System.Text.Encoding.Default)
            Do While Not sr.EndOfStream
                Dim s As String = sr.ReadLine
                If s = "" Then
                    '空行
                ElseIf s = "[地域]" Then
                    flag = 1
                ElseIf s = "[PM値]" Then
                    flag = 2
                ElseIf s = "[商品コード不可指定]" Then
                    flag = 3
                Else
                    If flag = 1 Then
                        chiikiArray.Add(s)
                    ElseIf flag = 2 Then
                        pmArray.Add(s)
                    ElseIf flag = 3 Then
                        sCodeArray.Add(s)
                    End If
                End If
            Loop
        End Using

        'datagridviewのヘッダーテキストをコレクションに取り込む
        Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
        'Dim dH6 As ArrayList = TM_HEADER_GET(dgv6)

        'マスタの商品コードをコレクションに取り込む
        Dim DGVCodeArray As New ArrayList
        For r As Integer = 0 To DGV6.RowCount - 1
            DGVCodeArray.Add(DGV6.Item(dH6.IndexOf("商品コード"), r).Value)
        Next

        'モードにより調べる行数を変更
        Dim rowArray As New ArrayList
        If mode = 0 Then
            For r As Integer = 0 To DGV1.RowCount - 1
                rowArray.Add(r)
            Next
        ElseIf mode = 1 Then
            For i As Integer = 0 To DGV1.SelectedCells.Count - 1
                If Not rowArray.Contains(DGV1.SelectedCells(i).RowIndex) Then
                    rowArray.Add(DGV1.SelectedCells(i).RowIndex)
                End If
            Next
        Else
            rowArray.Add(num)
        End If

        Dim chiiki As String = ""
        For i As Integer = 0 To chiikiArray.Count - 1
            If chiiki = "" Then
                chiiki = chiikiArray(i)
            Else
                chiiki &= "|" & chiikiArray(i)
            End If
        Next

        Dim pm As String = ""
        For i As Integer = 0 To pmArray.Count - 1
            If InStr(pmArray(i), "×") > 0 Then
                Dim pA As String() = Split(pmArray(i), "=")
                If pm = "" Then
                    pm = "p" & pA(0) & "|" & "P" & pA(0)
                Else
                    pm &= "|p" & pA(0) & "|" & "P" & pA(0)
                End If
            End If
        Next

        For i As Integer = 0 To rowArray.Count - 1
            'エラー除く
            If DGV1.Item(dH1.IndexOf("商品マスタ"), i).Style.BackColor = Color.Red Then
                Continue For
            End If

            Dim koukuuFlag As Boolean = True

            '商品コードで検索
            Dim mCode As String = DGV1.Item(TM_ArIndexof(dH1, "商品マスタ"), rowArray(i)).Value
            For k As Integer = 0 To sCodeArray.Count - 1
                If InStr(mCode, sCodeArray(k)) > 0 Then
                    koukuuFlag = False
                    Exit For
                End If
            Next

            'PM値で検索（マスタコード分解）
            If koukuuFlag = True Then
                Dim mCodeArray As String() = Split(mCode, "、")
                For k As Integer = 0 To mCodeArray.Length - 1
                    Dim codeA As String() = Split(mCodeArray(k), "*")
                    codeA(0) = codeA(0).ToLower
                    Dim sw As String = DGV6.Item(TM_ArIndexof(dH6, "ship-weight"), TM_ArIndexof(DGVCodeArray, codeA(0))).Value
                    If Regex.IsMatch(sw, pm) Then
                        koukuuFlag = False
                        Exit For
                    End If
                Next
            End If

            '地域で検索
            If koukuuFlag = True Then
                Dim addr As String = DGV1.Item(dH1.IndexOf("発送先住所"), rowArray(i)).Value
                addr = Regex.Replace(addr, "(\s|　)", "")    '空白を削除

                If Regex.IsMatch(addr.Substring(0, 5), chiiki) Then       '北海道|沖縄
                    'Dim SakiYubin As String = DGV1.Item(dH1.IndexOf("発送先郵便番号"), rowArray(i)).Value
                    'Dim whereCheck As String = " ([YUBIN] = '" & SakiYubin & "')"
                    'Dim sql As String = "SELECT * FROM [TOWN_T] WHERE" & whereCheck
                    'Dim table As DataTable = TM_DB_CONNECT_SELECT_SQL(CnAccdb, sql)

                    'If table.Rows.Count <> 0 Then
                    '    Dim TRANSIT_KBN As String = table.Rows(0)("TRANSIT_KBN").ToString   '離島検索（0だけ検索で良いかは不明）
                    '    If TRANSIT_KBN <> "0" Then   '離島フラグが立っていれば航空便不可
                    '        koukuuFlag = False
                    '    End If
                    'End If

                    'table.Clear()
                    koukuuFlag = True
                Else
                    koukuuFlag = False
                End If
            End If

            'まとめ
            If koukuuFlag = True Then
                '宅配便で、物流倉庫以外航空便OK
                'If dgv1.Item(TM_ArIndexof(dH1, "マスタ配送"), rowArray(i)).Value = "メール便" Or dgv1.Item(TM_ArIndexof(dH1, "発送倉庫"), rowArray(i)).Value = "物流倉庫" Then
                '宅配便で、物流倉庫も航空便許可（2019/01/28）
                DGV1.Item(dH1.IndexOf("元便種"), rowArray(i)).Value = "航空便"
                If DGV1.Item(TM_ArIndexof(dH1, "マスタ配送"), rowArray(i)).Value = "メール便" Then
                    DGV1.Item(dH1.IndexOf("便種"), rowArray(i)).Value = "陸便"
                Else
                    DGV1.Item(dH1.IndexOf("便種"), rowArray(i)).Value = "航空便"
                End If
            Else
                DGV1.Item(dH1.IndexOf("便種"), rowArray(i)).Value = "陸便"
                DGV1.Item(dH1.IndexOf("元便種"), rowArray(i)).Value = "陸便"
            End If
        Next
    End Sub


    '伝票番号に対応する商品コードを「商品マスタ」に読み込む
    Private Function MasterInput()
        Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
        Dim dH3 As ArrayList = TM_HEADER_GET(DGV3)

        For r1 As Integer = 0 To DGV1.RowCount - 1
            Dim dNum As String = DGV1.Item(dH1.IndexOf("伝票番号"), r1).Value

            Dim str As String = ""
            Dim tenpo As String = ""
            For r2 As Integer = 0 To DGV3.RowCount - 1

                If dNum = DGV3.Item(dH3.IndexOf("伝票番号"), r2).Value Then

                    DGV3.Item(dH3.IndexOf("伝票番号"), r2).Style.BackColor = Color.LemonChiffon

                    Dim codeHeader As String = "商品ｺｰﾄﾞ"
                    Dim code As String = ""
                    If InStr(DGV3.Item(dH3.IndexOf("商品ｺｰﾄﾞ"), r2).Value, "addfee") > 0 Then     '送料追加コードを除外
                        DGV1.Item(dH1.IndexOf("データ"), r1).Value &= "/addfee有り"
                        DGV1.Item(dH1.IndexOf("マスタ備考"), r1).Value &= "宅配送料済"
                        DGV1.Item(dH1.IndexOf("マーク"), r1).Value &= "●"
                        Continue For
                    End If
                    If CheckBox11.Checked And InStr(DGV3.Item(dH3.IndexOf("商品ｺｰﾄﾞ"), r2).Value, "wake") > 0 Then    'ヤフオク用
                        If Regex.IsMatch(DGV3.Item(dH3.IndexOf(codeHeader), r2).Value, "^[a-zA-Z0-9\-\/\=]+$") Then
                            codeHeader = "商品名"
                            code = Regex.Split(DGV3.Item(dH3.IndexOf(codeHeader), r2).Value, "\/|\=")(0)
                            DGV1.Item(dH1.IndexOf("配送備考"), r1).Value = "[[" & DGV3.Item(dH3.IndexOf(codeHeader), r2).Value & "]]"
                        Else
                            code = DGV3.Item(dH3.IndexOf(codeHeader), r2).Value
                        End If
                    Else
                        code = DGV3.Item(dH3.IndexOf(codeHeader), r2).Value
                    End If
                    str &= code & "*" & DGV3.Item(dH3.IndexOf("受注数"), r2).Value & "、"
                    If tenpo = "" Then
                        tenpo = DGV3.Item(dH3.IndexOf("店舗"), r2).Value
                    End If
                End If
            Next

            '購入者電話番号に「/」があったら卸用の依頼主情報を入れる
            Dim iraiDenwa As String = DGV1.Item(dH1.IndexOf("購入者電話番号"), r1).Value
            If InStr(iraiDenwa, "/") > 0 Then
                Dim iraiDenwa1 As String() = Split(iraiDenwa, ",")
                Dim iraiDenwa2 As String() = Split(iraiDenwa1(0), "/")
                'If InStr(iraiDenwa2(1), "select") > 0 Or InStr(iraiDenwa2(1), "Select") > 0 Or InStr(iraiDenwa2(1), "SELECT") > 0 Then
                '    tenpo = "Selecting"
                'Else
                '    tenpo = "卸直送" & iraiDenwa2(1)
                'End If

                If InStr(iraiDenwa2(1), "TL") > 0 Then
                    tenpo = "TL"
                Else
                    tenpo = "卸直送" & iraiDenwa2(1)
                End If
            End If

            str = str.TrimEnd("、")
            If str <> "" Then
                DGV1.Item(dH1.IndexOf("商品マスタ"), r1).Value = str
                DGV1.Item(dH1.IndexOf("商品マスタ"), r1).Style.BackColor = Color.Empty
                DGV1.Item(dH1.IndexOf("店舗"), r1).Value = tenpo
            Else
                Label9.Visible = False
                Panel3.Visible = True
                DGV1.Item(dH1.IndexOf("商品マスタ"), r1).Style.BackColor = Color.Red
                MsgBox("配送データと明細データが合っていない可能性があります")
                DGV1.CurrentCell = DGV1(dH1.IndexOf("商品マスタ"), r1)
                Return False
                Exit Function
            End If
        Next
        Return True
    End Function

    Private tag_decide As New ArrayList
    ''' <summary>
    ''' マスタ配送の計算を行う
    ''' </summary>
    ''' <param name="type">0: ファイルをいれる場合　1:処理する場合</param>
    ''' <param name="targetRow">対象行を指定、-1なら全行対象</param>
    Private Sub MasterHaisouKeisan(type As Integer, targetRow As Integer, dgvT As DataGridView, gyono As Integer, dgvTNum As Integer, filed As String)
        Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
        Dim dH3 As ArrayList = TM_HEADER_GET(DGV3)
        'Dim dH6 As ArrayList = TM_HEADER_GET(dgv6)

        Dim doukonArray As String() = File.ReadAllLines(appPathDir & "\config\version2\同梱特殊.txt", ENC_SJ)
        Dim ny331_50_codes As String() = New String() {"ny331-50-306"， "ny331-50-be"， "ny331-50-bk"， "ny331-50-co"， "ny331-50-dapi"， "ny331-50-flpi"， "ny331-50-flwh"， "ny331-50-hu"， "ny331-50-pa"， "ny331-50-pi"， "ny331-50-wh", "ny331-50-ye", "ny331-50-rose", "ny331-50-lor", "ny331-50-lgr", "ny331-50-kobk", "ny331-50-kohu", "ny331-50-koor", "ny331-50-kopi", "ny331-50-ot306"}
        'pa084-5
        'Dim masuku_zyogai As String() = New String() {"ny261-1000-a"， "ny261-2000-a"， "ny261-1000-ye"， "ny261-2000-ye"， "ny264-100-4000", "ny264-100-4000wh", "ny264-100-4000ye", "ny263-51", "ny264-100", "ny264-200", "ny264", "ny264-500", "ny264-3000a"}
        Dim masuku_zyogai As String() = New String() {"ny261-1000-a"， "ny261-1000-ye"， "ny261-2000-ye"， "ny264-100-4000", "ny264-100-4000wh", "ny264-100-4000ye", "ny264-100", "ny264-200", "ny264", "ny264-500", "ny264-3000a"}

        'Dim ny331_2500_codes As String() = New String() {"ny331-2500-be"}
        '扁盒口罩
        'Dim masuku_50codesPro As String() = New String() {"TEST"}

        'Dim masuku_50codesPro As String() = New String() {"ny341-10", "ny263-A", "ny263-B", "ny263-C", "ny344", "ny373-30-bk", "ny373-30-pi", "ny373-30-wh", "ny373-30-kobk", "ny373-30-kowh", "ny385-40-a", "ny385-40-b", "ny385-40-c", "ny385-40-d", "ny331-50-306"， "ny331-50-be"， "ny331-50-bk"， "ny331-50-co"， "ny331-50-dapi"， "ny331-50-flpi"， "ny331-50-flwh"， "ny331-50-hu"， "ny331-50-pa"， "ny331-50-pi"， "ny331-50-wh", "ny331-50-ye"}

        Dim masuku_50codesPro As String() = New String() {"ny341-10", "ny344"}



        tag_decide.Clear()

        If type = 0 Then
            ToolStripProgressBar1.Maximum = DGV1.RowCount
            ToolStripProgressBar1.Value = 0
            Application.DoEvents()

            For r1 As Integer = 0 To DGV1.RowCount - 1
                If targetRow <> -1 And targetRow <> r1 Then
                    Continue For
                End If

                'エラー除く
                If DGV1.Item(dH1.IndexOf("商品マスタ"), r1).Style.BackColor = Color.Red Then
                    Continue For
                End If

                '重さ、特殊の初期化
                omosa = 0
                tokuGroup.Clear()

                Dim haisouKind_moto As String = DGV1.Item(dH1.IndexOf("発送方法"), r1).Value
                'If DGV1.Item(dH1.IndexOf("伝票番号"), r1).Value = "3449455" Then
                '    Console.WriteLine(123)
                'End If

                Dim maeKind As String = ""
                Dim haisouKind As String = ""
                Dim haisouSize As Double = 0
                Dim sizeName As String = ""
                Dim haisouSizeArray As New ArrayList    'バラで保存

                'addfee対応
                If InStr(DGV1.Item(dH1.IndexOf("データ"), r1).Value, "addfee有り") > 0 Then
                    haisouKind = "宅配便"
                End If

                '同梱特殊対応
                Dim mCode As String = DGV1.Item(dH1.IndexOf("商品マスタ"), r1).Value
                Dim mCodeArray As String() = Split(mCode, "、")
                Dim mTag As String = ""
                Dim mItemName As String = ""

                'ny263-51 1個メール便、他には宅配便のため
                Dim special_taku As Boolean = False
                'ny263-51を含む、宅配便になる可能
                Dim special_taku2 As Boolean = False

                Dim special_takumasukuPro As Boolean = False
                'yamato转成宅配便
                Dim special_takuyamoto As Boolean = False
                'メール便になる
                Dim special_mail As Boolean = False

                Dim haisouSaki As String = DGV1.Item(dH1.IndexOf("発送先住所"), r1).Value

                Dim checkcodejuchusu_ny263_51 As Integer = 0
                Dim checkcodejuchusu_od492 As Integer = 0

                Dim checkcodejuchusu_ny263_306_42 As Integer = 0
                Dim checkcodejuchusu_ny263_306_51 As Integer = 0
                Dim checkcodejuchusu_ny275 As Integer = 0
                Dim checkcodejuchusu_ny306 As Integer = 0

                Dim checkcodejuchusu_ny373 As Integer = 0

                Dim checkcodejuchusu_ny385 As Integer = 0

                'Dim checkcodejuchusu_ad228 As Integer = 0
                'Dim checkcodejuchusu_sl065 As Integer = 0
                'Dim checkcodejuchusu_sl066 As Integer = 0
                'Dim checkcodejuchusu_sl067 As Integer = 0
                Dim checkcodejuchusu_ny331_50_flwh As Integer = 0


                Dim checkcodejuchusu_ny261 As Integer = 0
                Dim checkcodejuchusu_ny261_1000 As Integer = 0
                Dim checkcodejuchusu_ny261_2000 As Integer = 0


                Dim checkcodejuchusu_ny261ye As Integer = 0
                Dim checkcodejuchusu_ny261_1000ye As Integer = 0
                Dim checkcodejuchusu_ny261_2000ye As Integer = 0

                Dim checkcodejuchusu_ny264 As Integer = 0
                Dim checkcodejuchusu_ny264_100 As Integer = 0
                Dim checkcodejuchusu_ny264_200 As Integer = 0
                Dim checkcodejuchusu_ny264_500 As Integer = 0
                Dim checkcodejuchusu_ny264_3000 As Integer = 0
                'Dim checkcodejuchusu_pa084_5 As Integer = 0
                'Dim checkcodejuchusu_pa084_7 As Integer = 0

                'Dim checkcodejuchusu_pa084_ho As Integer = 0
                'Dim checkcodejuchusu_od433_co As Integer = 0
                Dim checkcodejuchusu_od433_wa As Integer = 0
                Dim checkcodejuchusu_ny331_50 As Integer = 0 'ny331_50シリーズ
                Dim checkcodejuchusu_ny331_2500 As Integer = 0 'ny331_2500シリーズ
                'Dim isYamatoGood As Boolean = True


                '强制更改标记
                Dim souko_check = False

                Dim masuku_50codesProjuchusu As Integer = 0 ' 


                For checkcodei As Integer = 0 To mCodeArray.Length - 1
                    Dim checkcode As String() = Split(mCodeArray(checkcodei), "*")
                    Dim checkcode_ As String() = Nothing
                    If InStr(checkcode(0), "-") > 0 Then
                        checkcode_ = Split(checkcode(0), "-")
                    End If

                    'If Not yamato_goods.Contains(checkcode(0).ToLower) Then
                    '    isYamatoGood = False
                    'End If

                    'If checkcode(0).ToLower = "ad228" Or checkcode(0).ToLower = "ad228-be" Or checkcode(0).ToLower = "ad228-bl" Or checkcode(0).ToLower = "ad228-co" Or checkcode(0).ToLower = "ad228-gr" Or checkcode(0).ToLower = "ad228-sb" Or checkcode(0).ToLower = "ad228-wa" Then
                    '    If checkcode(1) = Int(checkcode(1)) Then
                    '        checkcodejuchusu_ad2   = checkcodejuchusu_ad228 + checkcode(1)
                    '    End If
                    'End If

                    'If Not checkcode_ Is Nothing And Regex.IsMatch(haisouSaki, "沖縄") = False And Regex.IsMatch(haisouSaki, "北海道") = False Then
                    Dim tagYU2 As String = MasterTag(checkcode(0).ToLower)
                    Dim YU2tenpo As String = DGV1.Item(dH1.IndexOf("店舗"), r1).Value
                    Dim YU2Send As String = DGV1.Item(dH1.IndexOf("発送方法"), r1).Value
                    'YU2Flag = False
                    'Dim haisousoku As String = DGV1.Item(dH1.IndexOf("発送倉庫"), r1).Value

                    ''最初这里把所有的yupaku变成yamato
                    'If InStr(YU2Send, "メール便") Then
                    '    DGV1.Item(dH1.IndexOf("発送方法"), r1).Value = "ヤマト"
                    '    isYamatoGood = True
                    'End If

                    'If yamato_goods.Contains(checkcode(0).ToLower) Or InStr(checkcode(0).ToLower, "ny393-50-") Or InStr(checkcode(0).ToLower, "ap005") Or InStr(checkcode(0).ToLower, "ap039") Or InStr(checkcode(0).ToLower, "ee270") Or InStr(checkcode(0).ToLower, "zk218") Then
                    '    DGV1.Item(dH1.IndexOf("発送方法"), r1).Value = "ヤマト"
                    '    DGV1.Item(dH1.IndexOf("マスタ配送"), r1).Value = "ヤマト"
                    'End If

                    If YU2Send = "メール便" Then
                        DGV1.Item(dH1.IndexOf("発送方法"), r1).Value = "ヤマト"
                        DGV1.Item(dH1.IndexOf("マスタ配送"), r1).Value = "ヤマト"
                    End If


                    '这里是判断是否YU2的最初的判断
                    If CheckBox34.Checked And isyupakutenpu(YU2tenpo) = False Then
                        If InStr(tagYU2, "太宰府") And siharaiYU2 <> "代金引換" And InStr(YU2Send, "宅配便") Then
                            'If CheckBox34.Checked And InStr(tagYU2, "太宰府") Then
                            'If CheckBox34.Checked And InStr(tagYU2, "") Then
                            'siharai <> "代金引換" 
                            If (isyupakubyAddress(haisouSaki) And isyupakuGoodsbyCode(checkcode(0))) Or IsOutlyingislandsbyAddress(haisouSaki) Then
                                DGV1.Item(dH1.IndexOf("発送方法"), r1).Value = "メール便"
                                'DGV1.Item(dH1.IndexOf("マスタ配送"), r1).Value = "メール便"
                                DGV1.Item(dH1.IndexOf("マスタ配送"), r1).Value = "ゆう200"
                                DGV1.Item(dH1.IndexOf("データ"), r1).Value = "ゆう200"
                                DGV1.Item(dH1.IndexOf("ピック指示内容"), r1).Value = "ゆうパック元払"
                                'ピック指示内容,,,ピック指示内容,ﾋﾟｯｷﾝｸﾞ指示内容,ピッキング指示,
                                haisouKind_moto = "メール便"
                                YU2FlagLoad = True
                            End If
                        End If
                    End If
                    If checkcode(0).ToLower = "ny263-51" Then
                        '個数は整数
                        If checkcode(1) = Int(checkcode(1)) Then
                            checkcodejuchusu_ny263_51 = checkcodejuchusu_ny263_51 + checkcode(1)
                        End If
                    End If

                    If checkcode(0).ToLower = "od492" Then
                        '個数は整数
                        If checkcode(1) = Int(checkcode(1)) Then
                            checkcodejuchusu_od492 = checkcodejuchusu_od492 + checkcode(1)
                        End If
                    End If


                    If checkcode(0).ToLower = "ny275-bk" Or checkcode(0).ToLower = "ny275-pi" Then
                        If checkcode(1) = Int(checkcode(1)) Then
                            checkcodejuchusu_ny275 = checkcodejuchusu_ny275 + checkcode(1)
                        End If
                    End If

                    If checkcode(0).ToLower = "ny263-306-42" Then
                        '個数は整数
                        If checkcode(1) = Int(checkcode(1)) Then
                            checkcodejuchusu_ny263_306_42 = checkcodejuchusu_ny263_306_42 + checkcode(1)
                        End If
                    End If

                    If checkcode(0).ToLower = "ny263-306-51" Then
                        '個数は整数
                        If checkcode(1) = Int(checkcode(1)) Then
                            checkcodejuchusu_ny263_306_51 = checkcodejuchusu_ny263_306_51 + checkcode(1)
                        End If
                    End If

                    If checkcode(0).ToLower = "ny306-51" Then
                        If checkcode(1) = Int(checkcode(1)) Then
                            checkcodejuchusu_ny306 = checkcodejuchusu_ny306 + checkcode(1)
                        End If
                    End If

                    If checkcode(0).ToLower = "ny331-50-flwh" Then
                        If checkcode(1) = Int(checkcode(1)) Then
                            checkcodejuchusu_ny331_50_flwh = checkcodejuchusu_ny331_50_flwh + checkcode(1)
                        End If
                    End If

                    'If checkcode(0).ToLower = "ny261-1000-a" Or checkcode(0).ToLower = "ny261-2000-a" Or checkcode(0).ToLower = "ny264-100-4000" Or checkcode(0).ToLower = "ny264-100-4000wh" Then
                    If checkcode(0).ToLower = "ny261-1000-a" Or checkcode(0).ToLower = "ny264-100-4000" Or checkcode(0).ToLower = "ny264-100-4000wh" Then

                        If checkcode(1) = Int(checkcode(1)) Then
                            If checkcode(0).ToLower = "ny261-1000-a" Then
                                checkcodejuchusu_ny261 = checkcodejuchusu_ny261 + checkcode(1)
                                checkcodejuchusu_ny261_1000 = checkcodejuchusu_ny261_1000 + checkcode(1)
                            ElseIf checkcode(0).ToLower = "ny261-2000-a" Then
                                checkcodejuchusu_ny261 = checkcodejuchusu_ny261 + (checkcode(1) * 2)
                                checkcodejuchusu_ny261_2000 = checkcodejuchusu_ny261_2000 + checkcode(1)
                            ElseIf checkcode(0).ToLower = "ny264-100-4000" Or checkcode(0).ToLower = "ny264-100-4000wh" Then
                                checkcodejuchusu_ny261 = checkcodejuchusu_ny261 + (checkcode(1) * 4)
                            End If
                        End If
                    End If


                    If checkcode(0).ToLower = "ny261-1000-ye" Or checkcode(0).ToLower = "ny261-2000-ye" Or checkcode(0).ToLower = "ny264-100-4000ye" Then
                        If checkcode(1) = Int(checkcode(1)) Then
                            If checkcode(0).ToLower = "ny261-1000-ye" Then
                                checkcodejuchusu_ny261ye = checkcodejuchusu_ny261ye + checkcode(1)
                                checkcodejuchusu_ny261_1000ye = checkcodejuchusu_ny261_1000ye + checkcode(1)
                            ElseIf checkcode(0).ToLower = "ny261-2000-ye" Then
                                checkcodejuchusu_ny261ye = checkcodejuchusu_ny261ye + (checkcode(1) * 2)
                                checkcodejuchusu_ny261_2000ye = checkcodejuchusu_ny261_2000ye + checkcode(1)
                            ElseIf checkcode(0).ToLower = "ny264-100-4000ye" Then
                                checkcodejuchusu_ny261ye = checkcodejuchusu_ny261ye + (checkcode(1) * 4)
                            End If
                        End If
                    End If




                    If checkcode(0).ToLower = "ny264" Then
                        If checkcode(1) = Int(checkcode(1)) Then
                            checkcodejuchusu_ny264 = checkcodejuchusu_ny264 + checkcode(1)
                        End If
                    End If

                    If checkcode(0).ToLower = "ny264-100" Then
                        If checkcode(1) = Int(checkcode(1)) Then
                            checkcodejuchusu_ny264_100 = checkcodejuchusu_ny264_100 + checkcode(1)
                        End If
                    End If

                    If checkcode(0).ToLower = "ny264-200" Then
                        If checkcode(1) = Int(checkcode(1)) Then
                            checkcodejuchusu_ny264_200 = checkcodejuchusu_ny264_200 + checkcode(1)
                        End If
                    End If

                    If checkcode(0).ToLower = "ny264-500" Then
                        If checkcode(1) = Int(checkcode(1)) Then
                            checkcodejuchusu_ny264_500 = checkcodejuchusu_ny264_500 + checkcode(1)
                        End If
                    End If


                    'If ny331_50_codes.Contains(checkcode(0).ToLower) Then
                    '    If checkcode(1) = Int(checkcode(1)) Then
                    '        checkcodejuchusu_ny331_50 = checkcodejuchusu_ny331_50 + checkcode(1)
                    '    End If
                    'End If


                    If masuku_50codesPro.Contains(checkcode(0).ToLower) Then
                        If checkcode(1) = Int(checkcode(1)) Then
                            masuku_50codesProjuchusu = masuku_50codesProjuchusu + checkcode(1)
                        End If
                    End If




                    If checkcode(0).ToLower = "ny264-3000a" Then
                        If checkcode(1) = Int(checkcode(1)) Then
                            checkcodejuchusu_ny264_3000 = checkcodejuchusu_ny264_3000 + checkcode(1)
                        End If
                    End If



                    'If checkcode(0).ToLower = "pa084-5" Then
                    '    If checkcode(1) = Int(checkcode(1)) Then
                    '        checkcodejuchusu_pa084_5 = checkcodejuchusu_pa084_5 + checkcode(1)
                    '    End If
                    'End If



                    'If checkcode(0).ToLower = "pa084-7" Then
                    '    If checkcode(1) = Int(checkcode(1)) Then
                    '        checkcodejuchusu_pa084_7 = checkcodejuchusu_pa084_7 + checkcode(1)
                    '    End If
                    'End If


                    'If checkcode(0).ToLower = "pa084-ho" Then
                    '    If checkcode(1) = Int(checkcode(1)) Then
                    '        checkcodejuchusu_pa084_ho = checkcodejuchusu_pa084_ho + checkcode(1)
                    '    End If
                    'End If


                    'If checkcode(0).ToLower = "od433-co" Then
                    '    If checkcode(1) = Int(checkcode(1)) Then
                    '        checkcodejuchusu_od433_co = checkcodejuchusu_od433_co + checkcode(1)
                    '    End If
                    'End If

                    If checkcode(0).ToLower = "od433-wa" Then
                        If checkcode(1) = Int(checkcode(1)) Then
                            checkcodejuchusu_od433_wa = checkcodejuchusu_od433_wa + checkcode(1)
                        End If
                    End If


                    'ny373受注数
                    If InStr(checkcode(0).ToLower, "ny373") Then
                        If checkcode(1) = Int(checkcode(1)) Then
                            checkcodejuchusu_ny373 = checkcodejuchusu_ny373 + checkcode(1)
                        End If
                    End If





                    'ny385受注数
                    If InStr(checkcode(0).ToLower, "ny385") Then
                        If checkcode(1) = Int(checkcode(1)) Then
                            checkcodejuchusu_ny385 = checkcodejuchusu_ny385 + checkcode(1)
                        End If
                    End If




                    ''ny331_2500
                    'If ny331_2500_codes.Contains(checkcode(0).ToLower) Then
                    '    If checkcode(1) = Int(checkcode(1)) Then
                    '        checkcodejuchusu_ny331_2500 = checkcodejuchusu_ny331_2500 + checkcode(1)
                    '    End If
                    'End If







                Next

                If checkcodejuchusu_ny331_50 >= 3 And InStr(haisouSaki, "沖縄") = False Then
                    special_taku = True 'special_taku: 強制的に宅配便にする

                End If




                'If InStr(haisouSaki, "冲绳") Then
                '    If checkcodejuchusu_ny331_50 >= 5 Then
                '        special_taku = True
                '    End If
                'Else
                '    If checkcodejuchusu_ny331_50 >= 3 Then
                '        special_taku = True
                '    End If
                'End If
                If checkcodejuchusu_ny331_50 > 0 And InStr(haisouSaki, "沖縄") = False Then
                    special_taku2 = True 'special_taku: 強制的に宅配便にする(可能)
                End If



                special_takumasukuPro = False
                special_takuyamoto = False
                '发往冲绳的货物标记
                Dim sentToOkinawa = False
                If InStr(haisouSaki, "沖縄") Then

                    If checkcodejuchusu_ny331_50 > 5 Or masuku_50codesProjuchusu > 5 Then
                        'special_taku = True
                        special_takumasukuPro = True


                    End If

                    sentToOkinawa = True


                Else
                    'If checkcodejuchusu_ny331_50 > 3 Or masuku_50codesProjuchusu > 3 Then
                    '    special_taku = True
                    'End If
                End If


                If checkcodejuchusu_ny263_51 >= 3 Then
                    special_taku = True 'special_taku: 強制的に宅配便にする
                End If
                If checkcodejuchusu_ny263_51 > 0 Then
                    special_taku2 = True 'special_taku: 強制的に宅配便にする(可能)
                End If


                If checkcodejuchusu_ny263_306_42 >= 3 Then
                    special_taku = True
                End If
                If checkcodejuchusu_ny263_306_42 > 0 Then
                    special_taku2 = True
                End If

                If checkcodejuchusu_ny263_306_51 > 2 Then
                    special_taku = True
                End If
                If checkcodejuchusu_ny263_306_51 > 0 Then
                    special_taku2 = True
                End If

                If checkcodejuchusu_ny306 > 2 Then
                    special_taku = True
                End If
                If checkcodejuchusu_ny306 > 0 Then
                    special_taku2 = True
                End If

                If checkcodejuchusu_ny275 > 4 Then
                    special_taku = True
                End If
                If checkcodejuchusu_ny275 > 0 Then
                    special_taku2 = True
                End If


                'If checkcodejuchusu_sl065 >= 2 Then
                '    special_taku = True
                'End If
                'If checkcodejuchusu_sl065 > 0 Then
                '    special_taku2 = True
                'End If

                'If checkcodejuchusu_sl066 >= 2 Then
                '    special_taku = True
                'End If
                'If checkcodejuchusu_sl066 > 0 Then
                '    special_taku2 = True
                'End If

                'If checkcodejuchusu_sl067 >= 2 Then
                '    special_taku = True
                'End If
                'If checkcodejuchusu_sl067 > 0 Then
                '    special_taku2 = True
                'End If

                If checkcodejuchusu_ny331_50_flwh >= 3 Then
                    special_taku = True
                End If
                If checkcodejuchusu_ny331_50_flwh > 0 Then
                    special_taku2 = True
                End If


                'ny373修改
                If checkcodejuchusu_ny373 >= 5 Then
                    special_taku = True
                End If


                If checkcodejuchusu_ny385 >= 3 Then
                    special_taku = True
                End If




                'If checkcodejuchusu_ny373 > 0 Then
                '    special_taku2 = True
                'End If


                'Dim checkcodejuchusu_ad228_even = False
                'Dim checkcodejuchusu_ad228_even_count = 0
                'If checkcodejuchusu_ad228 > 1 Then
                '    If checkcodejuchusu_ad228 Mod 2 = 0 Then
                '        checkcodejuchusu_ad228_even = True
                '    Else
                '        checkcodejuchusu_ad228_even = False
                '        checkcodejuchusu_ad228_even_count = checkcodejuchusu_ad228 / 2
                '    End If
                'End If

                'Dim ny261_isnagoya As Boolean = False
                'If checkcodejuchusu_ny261 > 0 Then
                'If checkcodejuchusu_ny261 <= 3 Then
                '    If checkcodejuchusu_ny261_1000 > 0 And checkcodejuchusu_ny261_2000 = 0 Then
                '        If checkcodejuchusu_ny261_1000 = 1 Or checkcodejuchusu_ny261_1000 = 3 Then
                '            ny261_isnagoya = False '1000
                '        Else
                '            ny261_isnagoya = True '2000
                '        End If
                '    ElseIf checkcodejuchusu_ny261_1000 = 0 And checkcodejuchusu_ny261_2000 > 0 Then
                '        ny261_isnagoya = True
                '    ElseIf checkcodejuchusu_ny261_1000 > 0 And checkcodejuchusu_ny261_2000 > 0 Then
                '        ny261_isnagoya = False '3000
                '    End If
                'Else
                '    If checkcodejuchusu_ny261 Mod 2 = 1 Then
                '        ny261_isnagoya = False
                '    Else
                '        ny261_isnagoya = True
                '    End If
                'End If


                'If checkcodejuchusu_ny261 <= 2 Then
                '        If checkaddress_oosakizyou.Contains(haisouSaki) Then
                '            ny261_isnagoya = True
                '        Else
                '            ny261_isnagoya = False
                '        End If
                '    Else
                '        ny261_isnagoya = True
                '    End If
                'End If

                '这里的执行过程， 满足条件的情况下会把仓库加到tag_decide
                Dim ny261_isnagoya As Boolean = Nothing
                If checkcodejuchusu_ny261 > 0 Then
                    ny261_isnagoya = checkSouko_DaOrNa(tag_decide, "ny261", checkcodejuchusu_ny261, haisouSaki)
                End If
                'ny261_isnagoya = True
                Dim od492_isnagoyaortaizafu As Boolean = Nothing
                If checkcodejuchusu_od492 > 0 Then
                    od492_isnagoyaortaizafu = checkSouko_DaOrNa(tag_decide, "od492", checkcodejuchusu_od492, haisouSaki)
                End If


                If checkcodejuchusu_ny261ye > 0 Then
                    ny261_isnagoya = checkSouko_DaOrNa(tag_decide, "ny261ye", checkcodejuchusu_ny261ye, haisouSaki)
                End If


                '这里
                Dim ny263_51_isnagoya As Boolean = Nothing
                If checkcodejuchusu_ny263_51 > 0 And checkcodejuchusu_ny263_51 < 20 Then
                    ny263_51_isnagoya = checkSouko_DaOrNa(tag_decide, "ny263-51", checkcodejuchusu_ny263_51, haisouSaki)
                End If

                Dim ny264_isnagoya As Boolean = Nothing
                If checkcodejuchusu_ny264 > 0 Then
                    ny264_isnagoya = checkSouko_DaOrNa(tag_decide, "ny264", checkcodejuchusu_ny264, haisouSaki)
                End If

                Dim ny264_100_isnagoya As Boolean = Nothing
                If checkcodejuchusu_ny264_100 > 0 Then
                    ny264_100_isnagoya = checkSouko_DaOrNa(tag_decide, "ny264-100", checkcodejuchusu_ny264_100, haisouSaki)
                End If

                Dim ny264_200_isnagoya As Boolean = Nothing
                If checkcodejuchusu_ny264_200 > 0 Then
                    ny264_200_isnagoya = checkSouko_DaOrNa(tag_decide, "ny264-200", checkcodejuchusu_ny264_200, haisouSaki)
                End If

                Dim ny264_500_isnagoya As Boolean = Nothing
                If checkcodejuchusu_ny264_500 > 0 Then
                    ny264_500_isnagoya = checkSouko_DaOrNa(tag_decide, "ny264-500", checkcodejuchusu_ny264_500, haisouSaki)
                End If

                Dim ny264_3000_isnagoya As Boolean = Nothing
                If checkcodejuchusu_ny264_3000 > 0 Then
                    ny264_3000_isnagoya = checkSouko_DaOrNa(tag_decide, "ny264-3000a", checkcodejuchusu_ny264_3000, haisouSaki)
                End If

                'Dim pa084_5_isnagoya As Boolean = Nothing
                'If checkcodejuchusu_pa084_5 > 0 Then
                '    pa084_5_isnagoya = checkSouko_DaOrNa(tag_decide, "pa084-5", checkcodejuchusu_pa084_5, haisouSaki)
                'End If


                'Dim pa084_7_isnagoya As Boolean = Nothing
                'If checkcodejuchusu_pa084_7 > 0 Then
                '    pa084_7_isnagoya = checkSouko_DaOrNa(tag_decide, "pa084-7", checkcodejuchusu_pa084_7, haisouSaki)
                'End If


                'Dim pa084_ho_isnagoya As Boolean = Nothing
                'If checkcodejuchusu_pa084_ho > 0 Then
                '    pa084_ho_isnagoya = checkSouko_DaOrNa(tag_decide, "pa084-ho", checkcodejuchusu_pa084_ho, haisouSaki)
                'End If

                'Dim od433_co_isnagoya As Boolean = Nothing
                'If checkcodejuchusu_od433_co > 0 Then
                '    od433_co_isnagoya = checkSouko_DaOrNa(tag_decide, "od433-co", checkcodejuchusu_od433_co, haisouSaki)
                'End If

                Dim od433_wa_isnagoya As Boolean = Nothing
                If checkcodejuchusu_od433_wa > 0 Then
                    od433_wa_isnagoya = checkSouko_DaOrNa(tag_decide, "od433-wa", checkcodejuchusu_od433_wa, haisouSaki)
                End If


                Dim fukusuSoukoFlag As Boolean = False
                For i As Integer = 0 To mCodeArray.Length - 1
                    Dim code As String() = Split(mCodeArray(i), "*")
                    Dim sw As String = MasterWeight(code(0).ToLower)

                    sizeName &= sw & "/"
                    Dim tag As String = MasterTag(code(0).ToLower)
                    Dim itemName As String = MasterItemName(code(0).ToLower)
                    Dim juchusu As String = code(1)

                    souko_check = False


                    If haisouKind = "" Then
                        If Regex.IsMatch(sw, "P|p") Then
                            haisouKind = "宅配便"
                        ElseIf Regex.IsMatch(sw, "M|m") Then
                            haisouKind = "メール便"
                        ElseIf Regex.IsMatch(sw, "T|t") Then
                            haisouKind = "定形外"
                        End If
                    End If

                    Dim weight As String = Regex.Replace(sw, "P|p|M|m|T|t", "")
                    Dim w2 As String = Regex.Match(weight, "\(.*\)").Value
                    w2 = Regex.Replace(w2, "\(|\)", "")
                    weight = Regex.Replace(weight, "\(.*\)", "")

                    Dim sp_check = True

                    'ny261-1000 1, 2 >= 太宰府
                    'ny261-2000 1 >= 名古屋
                    'ny261 > 3000場合 奇数=>太宰府 偶数=>名古屋

                    'If checkcodejuchusu_ny261 > 0 Then
                    '    If checkcodejuchusu_ny261 <= 3 Then
                    '        If checkcodejuchusu_ny261_1000 > 0 And checkcodejuchusu_ny261_2000 = 0 Then
                    '            ny261_isnagoya = False
                    '        ElseIf checkcodejuchusu_ny261_1000 = 0 And checkcodejuchusu_ny261_2000 > 0 Then
                    '            ny261_isnagoya = True
                    '        ElseIf checkcodejuchusu_ny261_1000 > 0 And checkcodejuchusu_ny261_2000 > 0 Then
                    '            ny261_isnagoya = False
                    '        End If
                    '    Else
                    '        If checkcodejuchusu_ny261 Mod 2 = 1 Then
                    '            ny261_isnagoya = False
                    '        Else
                    '            ny261_isnagoya = True
                    '        End If
                    '    End If
                    'End If




                    If ny261_isnagoya Then
                        If haisouKind = "宅配便" And ((code(0).ToLower = "ny261-1000-a") Or (code(0).ToLower = "ny261-1000-ye")) Then
                            'weight = "16.66"
                            weight = "25"
                            sp_check = False
                        End If

                        If haisouKind = "宅配便" And ((code(0).ToLower = "ny261-2000-a") Or (code(0).ToLower = "ny261-2000-ye")) Then
                            'weight = "33.32"
                            weight = "50"
                            sp_check = False
                        End If

                        If haisouKind = "宅配便" And ((code(0).ToLower = "ny264-100-4000") Or (code(0).ToLower = "ny264-100-4000wh") Or (code(0).ToLower = "ny264-100-4000ye")) Then
                            'weight = "66.64"
                            weight = "100"
                            sp_check = False
                        End If


                    Else
                        If haisouKind = "宅配便" And ((code(0).ToLower = "ny261-1000-a") Or (code(0).ToLower = "ny261-1000-ye")) Then
                            weight = "50"
                            sp_check = False
                        End If

                        If haisouKind = "宅配便" And ((code(0).ToLower = "ny261-2000-a") Or (code(0).ToLower = "ny261-2000-ye")) Then
                            weight = "100"
                            sp_check = False
                        End If

                        If haisouKind = "宅配便" And ((code(0).ToLower = "ny264-100-4000") Or (code(0).ToLower = "ny264-100-4000wh") Or (code(0).ToLower = "ny264-100-4000ye")) Then
                            weight = "200"
                            sp_check = False
                        End If

                    End If

                    If (haisouKind = "メール便" And (code(0).ToLower = "ny263-51") And ny263_51_isnagoya And special_taku) Or (haisouKind_moto = "宅配便" And haisouKind = "宅配便" And (code(0).ToLower = "ny263-51") And ny263_51_isnagoya) Then
                        'weight = "83" '6000枚一箱 ny263-51是50枚 1/120 后面会除两次100 10000/120 
                        weight = "125" '10000/80
                        sp_check = False
                        'ElseIf (haisouKind = "メール便" And (code(0).ToLower = "ny263-51") And ny263_51_isnagoya = False And special_taku) Or (haisouKind_moto = "宅配便" And haisouKind = "宅配便" And (code(0).ToLower = "ny263-51") And ny263_51_isnagoya = False) Then
                        '    weight = "0.0083"
                        '    sp_check = False
                    ElseIf special_taku = True And haisouKind = "メール便" And code(0).ToLower = "ny263-51" And ny263_51_isnagoya = False Then
                        weight = "250" '40個2便
                        sp_check = False
                        'ElseIf InStr(code(0).ToLower, "ny263-331-50") Then
                        '    weight = "250" '40   1便
                        '    sp_check = False
                    End If

                    'If haisouKind = "宅配便" And (code(0).ToLower = "ny264") Then
                    '    weight = "2.5" '100/40
                    '    sp_check = False
                    'End If

                    'ny264 50枚
                    If haisouKind = "宅配便" And code(0).ToLower = "ny264" And ny264_isnagoya Then
                        weight = "1.25"
                        sp_check = False
                    ElseIf haisouKind = "宅配便" And code(0).ToLower = "ny264" And ny264_isnagoya = False Then
                        weight = "2.5" '100/40
                        sp_check = False
                    End If

                    'ny264-100 100枚
                    If haisouKind = "宅配便" And code(0).ToLower = "ny264-100" And ny264_100_isnagoya Then
                        weight = "2.5"
                        sp_check = False
                    End If

                    'ny264-200
                    If haisouKind = "宅配便" And code(0).ToLower = "ny264-200" And ny264_200_isnagoya Then
                        weight = "5"
                        sp_check = False
                    End If

                    'ny264-500
                    If haisouKind = "宅配便" And code(0).ToLower = "ny264-500" And ny264_500_isnagoya Then
                        weight = "12.5"
                        sp_check = False
                    End If

                    If haisouKind = "宅配便" And InStr(code(0).ToLower, "ap069") Then
                        weight = "12.5"
                        sp_check = False
                    End If




                    If haisouKind = "宅配便" And InStr(code(0).ToLower, "ny373-150-") Then
                        weight = "12.5"
                        sp_check = False
                    End If



                    'If haisouKind = "宅配便" And code(0).ToLower = "pa084-5" And pa084_5_isnagoya Then
                    '    weight = "12.5"
                    '    sp_check = False
                    'End If

                    If haisouKind = "宅配便" And code(0).ToLower = "pa084-7" Then
                        weight = "50"
                        sp_check = False
                    End If

                    If haisouKind = "宅配便" And code(0).ToLower = "pa084-5" Then
                        weight = "33"
                        sp_check = False
                    End If



                    If haisouKind = "宅配便" And code(0).ToLower = "pa084-ho" Then
                        weight = "17"
                        sp_check = False
                    End If


                    'ny264-3000
                    If haisouKind = "宅配便" And code(0).ToLower = "ny264-3000a" And ny264_3000_isnagoya Then
                        weight = "100" '3000枚一个口(100) => 4000枚(75) 又改回一个一口 100
                        sp_check = False
                    End If

                    If special_taku = True And haisouKind = "メール便" And code(0).ToLower = "ny306-51" Then
                        weight = "250"
                        sp_check = False
                    End If

                    'If haisouKind = "メール便" And (code(0).ToLower = "ny373-wh" Or code(0).ToLower = "ny373-hu" Or code(0).ToLower = "ny373-pi" Or code(0).ToLower = "ny373-bk" Or code(0).ToLower = "ny373-kob" Or code(0).ToLower = "ny373-kowh" Or code(0).ToLower = "ny373-kobk") Then
                    If haisouKind = "メール便" And InStr(code(0).ToLower, "ny373") And Not InStr(code(0).ToLower, "ny373-600") And Not InStr(code(0).ToLower, "ny373-150") Then

                        If special_taku Then
                            'weight = "178"
                            weight = "250"
                            sp_check = False
                        Else
                            'weight = "100"
                        End If
                    End If


                    'If isyupakubyAddress(haisouSaki) And isyupakuGoodBool And Regex.IsMatch(DGV1.Item(dH1.IndexOf("発送方法"), r1).Value, "宅配便") Then
                    '    weight = "100"
                    '    sp_check = False

                    'End If



                    'If special_taku = True And haisouKind = "メール便" And InStr(code(0).ToLower, "ny373") Then
                    '    weight = "178"
                    '    sp_check = False
                    'End If



                    'If special_taku = True And haisouKind = "メール便" And InStr(code(0).ToLower, "ny373") Then

                    'End If

                    'If InStr(code(0).ToLower, "ny373") And haisouKind = "メール便" Then
                    '    If special_taku Then
                    '        '56一个便
                    '        weight = "178"
                    '        'Else
                    '        '    'weight = "2.5"
                    '        '    weight = "100"
                    '    End If
                    'End If



                    'ny385 的 mlb 3个以上发佐川 佐川40个一便
                    If InStr(code(0).ToLower, "ny385") And haisouKind = "メール便" Then
                        If special_taku Then
                            weight = "250"
                        End If
                    End If


                    If haisouKind = "宅配便" And code(0).ToLower = "ap092" Then
                        weight = "2.5" '40一个口
                        sp_check = False
                    End If

                    If (haisouKind_moto = "宅配便" And code(0).ToLower = "ny263-306-42") Or (code(0).ToLower = "ny263-306-42" And special_taku) Then
                        weight = "200" '50個1便
                        sp_check = False
                    End If

                    If (haisouKind_moto = "宅配便" And haisouKind = "メール便" And ny331_50_codes.Contains(code(0).ToLower) Or (haisouKind = "メール便" And special_taku And ny331_50_codes.Contains(code(0).ToLower))) Then
                        weight = "228"
                        sp_check = False
                    End If

                    '扁盒口罩
                    If (haisouKind_moto = "宅配便" And haisouKind = "メール便" And masuku_50codesPro.Contains(code(0).ToLower) Or (haisouKind = "メール便" And special_takumasukuPro And masuku_50codesPro.Contains(code(0).ToLower))) Then
                        weight = "228"
                        sp_check = False
                    End If

                    'Or (code(0).ToLower = "ny405-ne")
                    If haisouKind_moto = "宅配便" And (code(0).ToLower = "ny405-bk" Or (code(0).ToLower = "ny405-bl") Or (code(0).ToLower = "ny405-bl") Or
                       (code(0).ToLower = "ny405-blhu") Or (code(0).ToLower = "ny405-ne") Or (code(0).ToLower = "ny405-hu") Or (code(0).ToLower = "ny405-iv") Or (code(0).ToLower = "ny405-lash") Or (code(0).ToLower = "ny405-milk") Or (code(0).ToLower = "ny405-pi") Or (code(0).ToLower = "ny405-rose") Or (code(0).ToLower = "ny405-wh") Or (code(0).ToLower = "ny405-ye")) Then
                        weight = "250"
                        sp_check = False
                    End If



                    'If haisouKind = "メール便" And (code(0).ToLower = "ny405-bk" Or (code(0).ToLower = "ny405-bl") Or (code(0).ToLower = "ny405-bl") Or
                    '   (code(0).ToLower = "ny405-blhu") Or (code(0).ToLower = "ny405-ne") Or (code(0).ToLower = "ny405-hu") Or (code(0).ToLower = "ny405-iv") Or (code(0).ToLower = "ny405-lash") Or (code(0).ToLower = "ny405-milk") Or (code(0).ToLower = "ny405-pi") Or (code(0).ToLower = "ny405-rose") Or (code(0).ToLower = "ny405-wh") Or (code(0).ToLower = "ny405-ye")) Then
                    '    weight = "250"
                    '    sp_check = False
                    'End If



                    If (haisouKind_moto = "宅配便" And code(0).ToLower = "ny263-306-51") Or (code(0).ToLower = "ny263-306-51" And special_taku) Then
                        weight = "300"
                        sp_check = False
                    End If

                    If haisouKind = "宅配便" And (code(0).ToLower = "de055" Or code(0).ToLower = "de055-01" Or code(0).ToLower = "od437-wa" Or code(0).ToLower = "od437-gr" Or code(0).ToLower = "od437-co" Or code(0).ToLower = "od437-hu") Then
                        weight = "130"
                        sp_check = False
                    End If

                    If haisouKind = "宅配便" And (code(0).ToLower = "ny348") Then
                        weight = "16.66" '100/6
                        sp_check = False
                    End If

                    If haisouKind = "宅配便" And ((code(0).ToLower = "ny261-450-c") Or (code(0).ToLower = "ny261-450-d")) Then
                        weight = "16.66" '100/6
                        sp_check = False
                    End If

                    'ad228 2個なら1便 
                    'If haisouKind = "宅配便" And (code(0).ToLower = "ad228" Or code(0).ToLower = "ad228-be" Or code(0).ToLower = "ad228-bl" Or code(0).ToLower = "ad228-co" Or code(0).ToLower = "ad228-gr" Or code(0).ToLower = "ad228-sb" Or code(0).ToLower = "ad228-wa") And checkcodejuchusu_ad228_even Then
                    '    weight = "50"
                    '    sp_check = False
                    'ElseIf haisouKind = "宅配便" And (code(0).ToLower = "ad228" Or code(0).ToLower = "ad228-be" Or code(0).ToLower = "ad228-bl" Or code(0).ToLower = "ad228-co" Or code(0).ToLower = "ad228-gr" Or code(0).ToLower = "ad228-sb" Or code(0).ToLower = "ad228-wa") And checkcodejuchusu_ad228_even = False Then
                    '    If checkcodejuchusu_ad228 = 1 Then
                    '        weight = "100"
                    '        sp_check = False
                    '    Else
                    '        checkcodejuchusu_ad228_even_count = checkcodejuchusu_ad228_even_count - 1
                    '        If checkcodejuchusu_ad228_even_count = 1 Then
                    '            weight = "100"
                    '            sp_check = False
                    '        Else
                    '            weight = "50"
                    '            sp_check = False
                    '        End If
                    '    End If
                    'End If

                    If haisouKind = "宅配便" And code(0).ToLower = "de072" Then
                        weight = "100"
                        sp_check = False
                    End If

                    If haisouKind = "宅配便" And code(0).ToLower = "ny328" Then
                        weight = "200"
                        sp_check = False
                    End If

                    If haisouKind = "宅配便" And code(0).ToLower = "ny185" Then
                        weight = "16.66"
                        sp_check = False
                    End If

                    '1個１便
                    If haisouKind = "宅配便" And (code(0).ToLower = "zk101-a" Or code(0).ToLower = "zk101-b" Or code(0).ToLower = "zk101-c" Or code(0).ToLower = "zk101-d" Or code(0).ToLower = "zk101-e" Or code(0).ToLower = "zk101-f" Or code(0).ToLower = "zk101-g" Or code(0).ToLower = "zk101-h" Or code(0).ToLower = "zk101-i") Then
                        weight = "100"
                        sp_check = False
                    End If

                    '5個まで１便
                    If haisouKind = "宅配便" And (code(0).ToLower = "ad105") Then
                        weight = "20"
                        sp_check = False
                    End If


                    If haisouKind = "宅配便" And (code(0).ToLower = "ee225") Then
                        weight = "16.66"
                        sp_check = False
                    End If

                    If haisouKind = "宅配便" And (code(0).ToLower = "ee227") Then
                        weight = "16.66"
                        sp_check = False
                    End If

                    If haisouKind = "宅配便" And InStr(code(0).ToLower, "ny331-2500") And InStr(code(0).ToLower, "2000") = False Then

                        'If haisouKind = "宅配便" And InStr(code(0).ToLower, "ny331-2500") Then

                        weight = "125"
                        sp_check = False
                    End If
                    'If haisouKind = "宅配便" And (code(0).ToLower = "ny365-xl" Or code(0).ToLower = "ny365-xxl" Or code(0).ToLower = "ny365-xxxl") Then
                    '    weight = "20"
                    '    sp_check = False
                    'End If


                    If haisouKind = "宅配便" And code(0).ToLower = "ad149" Then
                        weight = "16.5"
                        sp_check = False
                    End If

                    If haisouKind = "宅配便" And code(0).ToLower = "ny084" Then
                        weight = "12.5"
                        sp_check = False
                    End If

                    If haisouKind = "宅配便" And code(0).ToLower = "ny019" Then
                        weight = "0.67"
                        sp_check = False
                    End If
                    If haisouKind = "宅配便" And InStr(code(0).ToLower, "ad077") Then
                        weight = "25"
                        sp_check = False
                    End If


                    If code(0).ToLower = "ny263-51" Then
                        souko_check = True
                    End If

                    'isyupakuGoodBool = isyupakuGoodsbyCode(code(0))
                    'If CheckBox34.Checked And isyupakubyAddress(haisouSaki) And isyupakuGoodBool And Regex.IsMatch(DGV1.Item(dH1.IndexOf("発送方法"), r1).Value, "宅配便") Then
                    '    'weight = "100"

                    '    DGV1.Item(dH1.IndexOf("発送方法"), r1).Value = "メール便"
                    '    DGV1.Item(dH1.IndexOf("マスタ配送"), r1).Value = "メール便"
                    '    DGV1.Item(dH1.IndexOf("データ"), r1).Value = "ゆう200"
                    '    '一定  
                    '    'DGV1.Item(dH1.IndexOf("マスタ便数"), r1).Value = hakoArray.Count
                    '    YU2Flag = True
                    '    haisouKind = "メール便"
                    '    haisouKind_moto = DGV1.Item(dH1.IndexOf("発送方法"), r1).Value
                    '    sw = "M100"
                    'End If

                    'デフォルトがメール便で、宅配便があったら、haisouKindを宅配便にする
                    'メール便と定形外の場合は、宅配便にする
                    If (haisouKind = "メール便" And Regex.IsMatch(sw, "P|p|T|t")) Or (haisouKind = "定形外" And Regex.IsMatch(sw, "P|p|M|m")) Then
                        maeKind = haisouKind
                        haisouKind = "宅配便"
                        Select Case maeKind
                            Case "メール便"
                                haisouSize = haisouSize / 100   'メール便サイズ計算を宅配便に変更
                            Case "定形外"
                                haisouSize = haisouSize / 75   '定形外サイズ計算を宅配便に変更
                        End Select
                    End If

                    'haisouKindが宅配便の時、メール便を1/100にする
                    Dim henkouFlag As Boolean = False
                    If haisouKind = "宅配便" And Regex.IsMatch(sw, "M|m|T|t") Then
                        henkouFlag = True
                    ElseIf special_taku = True And Regex.IsMatch(sw, "M|m|T|t") Then
                        henkouFlag = True
                    End If

                    'バラで保存（とりあえず定形外用で使用）
                    For j As Integer = 1 To juchusu
                        haisouSizeArray.Add(weight)
                    Next

                    '重さ・特殊の処理
                    Dim sp As String = DGV6.Item(dH6.IndexOf("特殊"), dgv6CodeArray.IndexOf(code(0).ToLower)).Value

                    If InStr(sp, "重") > 0 And sp_check Then
                        Dim omotoku As Double = OmosaCheck(code(0).ToLower, juchusu, haisouSize)
                        If omotoku <> 0 Then
                            haisouSize = haisouSize + omotoku
                        End If
                    ElseIf InStr(sp, "特") > 0 And sp_check Then
                        Dim omotoku As Double = TokusyuCheck(code(0).ToLower, juchusu, haisouSize)
                        If omotoku <> 0 Then
                            haisouSize = haisouSize + omotoku
                        End If
                    Else
                        '-------
                        'weight = 20
                        If Not IsNumeric(weight) Then
                            DGV1.Item(dH1.IndexOf("マスタ配送"), r1).Value = "計算不能"
                            DGV1.Item(dH1.IndexOf("マスタ配送"), r1).Style.BackColor = Color.Red
                            Exit For
                        End If
                        If haisouKind = "メール便" And special_taku = False Then
                            haisouSize = haisouSize + (CDbl(weight) * CDbl(juchusu))
                        ElseIf haisouKind = "定形外" And special_taku = False Then
                            haisouSize = haisouSize + (CDbl(weight) * CDbl(juchusu))
                        Else
                            If henkouFlag Then
                                'If ny331_50_codes.Contains(code(0).ToLower) Or code(0).ToLower.Contains("ny263-331-50") Or code(0).ToLower.Contains("ny393-50-") Then

                                If (code(0).ToLower.Contains("ny331-50") And Not code(0).ToLower.Contains("ny331-500")) Or code(0).ToLower.Contains("ny263-331-50") Or (code(0).ToLower.Contains("ny393-50") And Not code(0).ToLower.Contains("ny393-500")) Then

                                    '1 7 9 2 3 9 12 1 2
                                    haisouSize = haisouSize + (2.5 * CDbl(juchusu))
                                    'ElseIf code(0).ToLower.Contains("ny439-") Then
                                    '    haisouSize = haisouSize + (2 * CDbl(juchusu))
                                Else
                                    haisouSize = haisouSize + ((CDbl(weight) / 100) * CDbl(juchusu))
                                End If


                            Else
                                If w2 <> "" Then    'P160(98)、P100(3.3)は1個のパーセンテージ（優先）
                                    haisouSize = haisouSize + (CDbl(w2) * CDbl(juchusu))
                                Else
                                    '50個一便
                                    If code(0).ToLower = "ny261" Or code(0).ToLower = "ny261-01" Or code(0).ToLower = "ny261-02" Or code(0).ToLower = "ny261-02a" Or code(0).ToLower = "ny261-03" Or code(0).ToLower = "ny263" Or code(0).ToLower = "ny263-00a" Then
                                        haisouSize = haisouSize + (1.66 * CDbl(juchusu))
                                    ElseIf code(0).ToLower = "mask01" Or code(0).ToLower = "mask-wh" Or code(0).ToLower = "mask05" Or code(0).ToLower = "mask01-bl" Or code(0).ToLower = "mask01-ko" Then
                                        haisouSize = haisouSize + (2.5 * CDbl(juchusu))
                                        'ElseIf code(0).ToLower = "ny185" Then
                                        '    '6個まで１便、以上２便
                                        '    haisouSize = haisouSize + (16.65 * CDbl(juchusu))
                                    ElseIf Regex.IsMatch(code(0).ToLower, "ad009") Then
                                        'If code(0).ToLower = "ad009-ne" Or code(0).ToLower = "ad009-wa" Then
                                        '    haisouSize = haisouSize + (20 * CDbl(juchusu)) '100/5
                                        'Else
                                        '    haisouSize = haisouSize + (12.5 * CDbl(juchusu)) '100/8
                                        'End If
                                        haisouSize = haisouSize + (10 * CDbl(juchusu)) '100/10

                                    Else
                                        If sp_check Then

                                            haisouSize = haisouSize + (CDbl(TakuhaiPerConv(weight)) * CDbl(juchusu))
                                        Else

                                            haisouSize = haisouSize + (CDbl(weight) * CDbl(juchusu))
                                        End If


                                        'If YU2Flag Then
                                        '    'haisouSize = haisouSize + (CDbl(weight) * CDbl(juchusu))
                                        '    haisouSize = haisouSize + (CDbl(TakuhaiPerConv(weight)) * CDbl(juchusu))
                                        'End If

                                    End If
                                End If
                            End If
                        End If
                        '-------
                    End If

                    'If DGV1.Item(dH1.IndexOf("伝票番号"), r1).Value = "3275468" Then
                    '    Console.WriteLine(123)
                    'End If



                    'yamato三个以上变成佐川
                    If special_taku = False Then
                        Dim cc = DGV1.Item(dH1.IndexOf("マスタ配送"), r1).Value
                        If cc = "ヤマト" Then
                            '只要不是扁盒口罩  yamato三个以上变成佐川



                            'If haisouSize >= NumericUpDown4.Value * 100 And Not masuku_50codesPro.Contains(code(0).ToLower) And Not code(0).Contains("ny373") And Not (code(0).Contains("ny417")) And Not code(0).Contains("ny411") And Not sentToOkinawa Then



                            'If haisouSize >= NumericUpDown4.Value * 100 And Not masuku_50codesPro.Contains(code(0).ToLower) And Not code(0).Contains("ny411-60") Then

                            '    special_taku = True
                            '    special_takuyamoto = True
                            '    haisouKind = "宅配便"
                            '    haisouSize = haisouSize / 100
                            'ElseIf haisouSize >= NumericUpDown4.Value * 100 And (code(0).Contains("ny417-30")) And Not sentToOkinawa Then
                            '    'c213 'haisouKind = "ヤマト"
                            '    special_taku = True
                            '    special_takuyamoto = True
                            '    haisouKind = "宅配便"
                            '    haisouSize = haisouSize / 100
                            'End If



                            If haisouSize >= NumericUpDown4.Value * 100 And IsNotChangeCode(code(0)) Then
                                special_taku = True
                                special_takuyamoto = True
                                haisouKind = "宅配便"
                                haisouSize = haisouSize / 100
                            End If







                            'If haisouSize >= 5 * 100 And (masuku_50codesPro.Contains(code(0).ToLower) Or code(0).Contains("ny373") Or (code(0).Contains("ny417") Or code(0).Contains("ny411"))) And sentToOkinawa Then

                            '    special_taku = True
                            '    special_takuyamoto = True
                            '    haisouKind = "宅配便"
                            '    haisouSize = haisouSize / 100
                            'Else
                            '    'c213 'haisouKind = "ヤマト"
                            'End If



                            'If isChangeYTT(code(0)) And Not sentToOkinawa Then
                            '    special_taku = True
                            '    special_takuyamoto = True
                            '    haisouKind = "宅配便"
                            '    haisouSize = haisouSize / 100
                            'End If





                        End If

                    End If


                    'メール便指定数以上は宅配便に変更   邮局变成佐川
                    If special_taku = False Then


                        If haisouKind = "メール便" Then
                            If haisouSize > NumericUpDown4.Value * 100 Then
                                If code(0).ToLower = "ny275-bk" Or code(0).ToLower = "ny275-pi" Then
                                    If haisouSize > 400 Then
                                        haisouKind = "宅配便"
                                        haisouSize = haisouSize / 100   'メール便サイズ計算を宅配便に変更
                                    Else
                                        special_mail = True
                                    End If
                                ElseIf masuku_50codesPro.Contains(code(0).ToLower) Then

                                    Debug.WriteLine("masuku_50codesPro")
                                Else
                                    haisouKind = "宅配便"
                                    haisouSize = haisouSize / 100   'メール便サイズ計算を宅配便に変更
                                End If
                            End If

                        ElseIf haisouKind = "定形外" Then
                            If haisouSize > NumericUpDown5.Value * NumericUpDown6.Value Then
                                haisouKind = "宅配便"
                                haisouSize = haisouSize / 75   '定形外サイズ計算を宅配便に変更
                            End If
                        End If
                    End If

                    '宅配便があれば、宅配便商品の商品名を使用する
                    If mItemName = "" Then
                        mItemName = itemName
                    Else
                        If Regex.IsMatch(sw, "P|p") Then
                            mItemName = itemName
                        End If
                    End If

                    'ny261-1000 - a除外して計算
                    'If (code(0).ToLower <> "ny261-1000-a" And code(0).ToLower <> "ny261-2000-a" And code(0).ToLower <> "ny264-100-4000") Then
                    If masuku_zyogai.Contains(code(0).ToLower) = False Then
                        '商品分類タグから発送倉庫を調べる
                        Dim tag_arr As String() = tag.Split("]")

                        '[太宰府][井相田発送][新宮][名古屋] 複数
                        Dim tag_check_arr As String() = New String() {HS1.Text, HS2.Text, HS3.Text, HS4.Text}
                        Dim tenmp_tag_check As String = ""
                        If tag_arr.Count > 0 Then
                            For k As Integer = 0 To tag_arr.Length - 1
                                If Regex.IsMatch(tag_arr(k), HS1.Text) Or Regex.IsMatch(tag_arr(k), HS2.Text) Or Regex.IsMatch(tag_arr(k), HS3.Text) Or Regex.IsMatch(tag_arr(k), HS4.Text) Then
                                    If tenmp_tag_check = "" Then
                                        tenmp_tag_check = tag_arr(k).Replace("[", "").Replace("]", "")
                                    Else
                                        If tag_arr(k).Replace("[", "").Replace("]", "") <> tenmp_tag_check Then
                                            fukusuSoukoFlag = True
                                            Exit For
                                        End If
                                    End If
                                End If
                            Next
                        Else
                        End If

                        If fukusuSoukoFlag Then
                            mTag = "複数倉庫"
                        Else
                            Select Case True
                                Case Regex.IsMatch(tag, HS2.Text)
                                    tag = HS2.Text
                                Case Regex.IsMatch(tag, HS3.Text)
                                    tag = HS3.Text
                                Case Regex.IsMatch(tag, HS1.Text)
                                    tag = HS1.Text
                                Case Regex.IsMatch(tag, HS4.Text)
                                    tag = HS4.Text
                                Case Else
                                    tag = "不明"
                            End Select

                            If mTag = "" Then
                                mTag = tag
                            ElseIf mTag = "不明" Or mTag = "複数倉庫" Then
                                mTag = mTag     '変更無し
                            ElseIf mTag = tag Then
                                mTag = tag      '変更無し
                            Else
                                mTag = "複数倉庫"
                            End If
                        End If

                        'MsgBox(maeKind & "/" & haisouKind & "/" & sw & "/" & haisouSize)
                    End If
                Next

                '同梱特殊処理
                If mTag = "複数倉庫" Then
                    Dim hassouArray As New ArrayList
                    For i As Integer = 0 To mCodeArray.Length - 1
                        Dim dFlag As Boolean = False
                        Dim code As String() = Split(mCodeArray(i), "*")
                        'ny261-1000 - a除外して計算
                        If masuku_zyogai.Contains(code(0).ToLower) = False Then
                            For Each dou As String In doukonArray
                                Dim douA As String() = Split(dou, ",")
                                If code(0).ToLower = douA(0) Then
                                    If Not hassouArray.Contains(douA(1)) Then
                                        hassouArray.Add(douA(1))
                                    End If
                                    dFlag = True
                                    Exit For
                                End If
                            Next
                            If dFlag = False Then
                                Dim tag As String = MasterTag(code(0).ToLower)
                                Select Case True
                                    Case Regex.IsMatch(tag, HS2.Text)
                                        tag = HS2.Text
                                    Case Regex.IsMatch(tag, HS3.Text)
                                        tag = HS3.Text
                                    Case Regex.IsMatch(tag, HS1.Text)
                                        tag = HS1.Text
                                    Case Regex.IsMatch(tag, HS4.Text)
                                        tag = HS4.Text
                                    Case Else
                                        tag = "不明"
                                End Select
                                If Not hassouArray.Contains(tag) Then
                                    hassouArray.Add(tag)
                                End If
                            End If
                        End If
                    Next

                    If hassouArray.Count = 1 And fukusuSoukoFlag = False Then
                        mTag = hassouArray(0)
                    Else
                        mTag = "複数倉庫"
                    End If
                End If

                'また他の商品で倉庫を判断する
                If mTag <> "複数倉庫" Then


                    If tag_decide.Count > 0 Then
                        Dim tag_decide_tem As String = ""
                        If tag_decide.Count = 1 Then
                            tag_decide_tem = tag_decide(0)
                        Else
                            For i As Integer = 0 To tag_decide.Count - 1
                                If tag_decide_tem = "" Then
                                    tag_decide_tem = tag_decide(i)
                                Else
                                    If tag_decide(i) <> tag_decide_tem Then
                                        mTag = fukususouko_str
                                        Exit For
                                    End If
                                End If
                            Next
                        End If

                        If mTag <> fukususouko_str Then
                            If tag_decide_tem = nagoya_str Then
                                If mTag = "" Then
                                    mTag = nagoya_str
                                ElseIf mTag = nagoya_str Then


                                Else
                                    mTag = fukususouko_str
                                End If
                            Else
                                If mTag = "" Then
                                    mTag = dazaifu_str
                                ElseIf mTag = nagoya_str Then
                                    mTag = fukususouko_str
                                Else

                                End If
                            End If
                        End If

                        If souko_check Then
                            For index = 0 To tag_decide.Count - 1
                                mTag = tag_decide(index)
                            Next
                        End If
                    End If

                    'ny261_isnagoya
                    'If checkcodejuchusu_ny261 > 0 Then 'ny261がある場合、ny261で判断した「ny261_isnagoya」は「mTag」を比較する
                    '    If ny261_isnagoya Then
                    '        If mTag = "" Then
                    '            mTag = "名古屋"
                    '        ElseIf mTag = "名古屋" Then

                    '        Else
                    '            mTag = "複数倉庫"
                    '        End If
                    '    Else
                    '        If mTag = "" Then
                    '            mTag = "太宰府"
                    '        ElseIf mTag = "名古屋" Then
                    '            mTag = "複数倉庫"
                    '        Else

                    '        End If
                    '    End If
                    'End If
                    tag_decide.Clear()
                End If

                If DGV1.Item(dH1.IndexOf("マスタ配送"), r1).Value <> "計算不能" Then
                    If special_taku = True Then
                        haisouKind = "宅配便"
                    End If

                    'If Regex.IsMatch(haisouKind, "メール便|定形外") And DGV1.Item(dH1.IndexOf("合計金額"), r1).Value <> 0 And mCodeArray.Contains("ny263-51") = False And mCodeArray.Contains("ny275-pi") = False And mCodeArray.Contains("ny275-bk") = False Then
                    If Regex.IsMatch(haisouKind, "メール便|定形外") And DGV1.Item(dH1.IndexOf("合計金額"), r1).Value <> 0 Then
                        haisouKind = "宅配便"
                        haisouSize = haisouSize * 0.01
                    End If

                    '便数計算（メール便と宅配便は100％で1便、定形外は250gで1便）
                    '箱計算（100％の箱を用意し、入る箱に順次入れ、個数を計算する）

                    ''ヤマト便数计算  新写的
                    'If Regex.IsMatch(haisouKind, "ヤマト") Then
                    '    DGV1.Item(dH1.IndexOf("マスタ便数"), r1).Value = Math.Ceiling(haisouSize / 100)
                    'End If


                    If Regex.IsMatch(haisouKind, "宅配便") Then
                        DGV1.Item(dH1.IndexOf("マスタ便数"), r1).Value = Math.Ceiling(haisouSize / 100)
                    Else


                        Dim hakoArray As New ArrayList
                        hakoArray.Add(0)
                        For i As Integer = 0 To haisouSizeArray.Count - 1
                            Dim hakoFlag As Boolean = False
                            For k As Integer = 0 To hakoArray.Count - 1
                                Dim per As Double = haisouSizeArray(i)
                                If Regex.IsMatch(haisouKind, "定形外") Then
                                    per = (haisouSizeArray(i) / 250) * 100
                                End If
                                If per + hakoArray(k) <= 100 Then
                                    hakoArray(k) = per + hakoArray(k)
                                    hakoFlag = True
                                    Exit For
                                End If
                            Next
                            If Not hakoFlag Then
                                hakoArray.Add(haisouSizeArray(i))
                            End If
                        Next


                        'If special_takumasukuPro Then
                        If Regex.IsMatch(haisouKind, "メール便") Then
                            If hakoArray.Count > NumericUpDown4.Value Then

                                If special_mail = False Then
                                    haisouKind = "宅配便"
                                    haisouSize = haisouSize * 0.01
                                    DGV1.Item(dH1.IndexOf("マスタ便数"), r1).Value = Math.Ceiling(haisouSize / 100)
                                Else
                                    DGV1.Item(dH1.IndexOf("マスタ便数"), r1).Value = hakoArray.Count
                                End If
                            ElseIf haisouKind_moto = "宅配便" Then
                                haisouSize = haisouSize * 0.01
                                DGV1.Item(dH1.IndexOf("マスタ便数"), r1).Value = Math.Ceiling(haisouSize / 100)
                            Else
                                DGV1.Item(dH1.IndexOf("マスタ便数"), r1).Value = hakoArray.Count
                            End If
                        Else
                            If hakoArray.Count > NumericUpDown5.Value Then
                                haisouKind = "宅配便"
                                haisouSize = haisouSize * 0.01
                                DGV1.Item(dH1.IndexOf("マスタ便数"), r1).Value = Math.Ceiling(haisouSize / 100)
                            ElseIf haisouKind_moto = "宅配便" Then
                                haisouSize = haisouSize * 0.01
                                DGV1.Item(dH1.IndexOf("マスタ便数"), r1).Value = Math.Ceiling(haisouSize / 100)
                            Else
                                DGV1.Item(dH1.IndexOf("マスタ便数"), r1).Value = hakoArray.Count
                            End If
                        End If

                        'End If
                        ''「最大重量 / 2 > 最大数」なら宅配便
                        'Dim counter As Integer = 0
                        'For i As Integer = 0 To haisouSizeArray.Count - 1
                        '    If haisouSizeArray(i) > (NumericUpDown6.Value / 2) Then
                        '        counter += 1
                        '    End If
                        'Next
                        'If counter >= NumericUpDown5.Value And haisouSizeArray.Count > NumericUpDown5.Value Then  '定形外最大数
                        '    haisouKind = "宅配便"
                        '    haisouSize = haisouSize * 0.01
                        '    DGV1.Item(dH1.IndexOf("マスタ便数"), r1).Value = Math.Ceiling(haisouSize / 100)
                        'Else
                        '    DGV1.Item(dH1.IndexOf("マスタ便数"), r1).Value = Math.Ceiling(haisouSize / NumericUpDown6.Value)
                        'End If
                    End If




                    Dim cc = DGV1.Item(dH1.IndexOf("マスタ配送"), r1).Value
                    Dim special_takuyamoto2 = False
                    If cc = "ヤマト" Then
                        special_takuyamoto2 = True
                    End If


                    DGV1.Item(dH1.IndexOf("マスタ配送"), r1).Value = haisouKind

                    If special_taku = True And special_takumasukuPro = False Then
                        DGV1.Item(dH1.IndexOf("発送方法"), r1).Value = "宅配便"
                        DGV1.Item(dH1.IndexOf("データ"), r1).Value = "佐川"
                    End If
                    If haisouKind = "宅配便" And special_taku2 = True And special_takumasukuPro = False Then
                        DGV1.Item(dH1.IndexOf("発送方法"), r1).Value = "宅配便"
                        DGV1.Item(dH1.IndexOf("データ"), r1).Value = "佐川"
                    End If
                    'special_takuyamoto  yamato强制变成宅配便
                    If special_taku And haisouKind = "宅配便" And special_takuyamoto Then
                        DGV1.Item(dH1.IndexOf("発送方法"), r1).Value = "宅配便"
                        DGV1.Item(dH1.IndexOf("データ"), r1).Value = "佐川"
                    End If

                    If haisouKind = "メール便" And special_takuyamoto2 Then
                        DGV1.Item(dH1.IndexOf("マスタ配送"), r1).Value = "ヤマト"
                        DGV1.Item(dH1.IndexOf("発送方法"), r1).Value = yamato_str
                        DGV1.Item(dH1.IndexOf("データ"), r1).Value = "画面"

                    End If




                    If haisouKind = "宅配便" And special_takumasukuPro = True Then
                        DGV1.Item(dH1.IndexOf("発送方法"), r1).Value = "宅配便"
                        DGV1.Item(dH1.IndexOf("データ"), r1).Value = "佐川"
                    End If





                    'If isyupakuGoodBool Then

                    'End If


                    'If haisouKind = "メール便" And special_taku = True Then
                    '    DGV1.Item(dH1.IndexOf("発送方法"), r1).Value = "宅配便"
                    '    DGV1.Item(dH1.IndexOf("データ"), r1).Value = "佐川"
                    'End If





                    'If (mTag = dazaifu_str Or mTag = "井相田") And (Regex.IsMatch(DGV1.Item(dH1.IndexOf("発送方法"), r1).Value, "メール便") Or Regex.IsMatch(DGV1.Item(dH1.IndexOf("発送方法"), r1).Value, "ヤマト")) And DGV1.Item(dH1.IndexOf("便種"), r1).Value = "陸便" And isYamatoGood And special_taku = False Then
                    '    DGV1.Item(dH1.IndexOf("マスタ配送"), r1).Value = yamato_str
                    '    DGV1.Item(dH1.IndexOf("発送方法"), r1).Value = yamato_str
                    '    DGV1.Item(dH1.IndexOf("データ"), r1).Value = "画面"
                    'End If


                    DGV1.Item(dH1.IndexOf("sw"), r1).Value = Math.Ceiling(haisouSize)

                    DGV1.Item(dH1.IndexOf("サイズ"), r1).Value = sizeName.TrimEnd("/")

                    DGV1.Item(dH1.IndexOf("発送倉庫"), r1).Value = mTag





                    If mTag = "複数倉庫" Then
                        DGV1.Item(dH1.IndexOf("発送倉庫"), r1).Style.BackColor = Color.Red
                    End If

                    DGV1.Item(dH1.IndexOf("品名"), r1).Value = mItemName
                End If

                ToolStripProgressBar1.Value += 1
            Next
        ElseIf type = 1 Then
            Dim c = 0
            'If dgvT Is DGV7 Or dgvT Is DGV8 Then
            '    Dim dHTemp As ArrayList = TM_HEADER_GET(dgvT)
            '    Dim haisouSize_ As Double = 0
            '    Dim haisouSizeArray_ As New ArrayList    'バラで保存

            '    If DGV1.Item(dH1.IndexOf("商品マスタ"), gyono).Style.BackColor <> Color.Red Then
            '        Dim mCode_ As String = DGV1.Item(dH1.IndexOf("商品マスタ"), gyono).Value
            '        Dim mCodeArray_ As String() = Split(mCode_, "、")

            '        For i As Integer = 0 To mCodeArray_.Length - 1
            '            Dim code_ As String() = Split(mCodeArray_(i), "*")
            '            Dim sw_ As String = MasterWeight(code_(0).ToLower)
            '            Dim tag_ As String = MasterTag(code_(0).ToLower)
            '            Dim juchusu_ As String = code_(1)
            '            Dim haisouKind_ As String = ""
            '            Dim maeKind_ As String = ""

            '            If Regex.IsMatch(sw_, "P|p") Then
            '                haisouKind_ = "宅配便"
            '            ElseIf Regex.IsMatch(sw_, "M|m") Then
            '                haisouKind_ = "メール便"
            '            ElseIf Regex.IsMatch(sw_, "T|t") Then
            '                haisouKind_ = "定形外"
            '            End If

            '            Dim weight_ As String = Regex.Replace(sw_, "P|p|M|m|T|t", "")
            '            Dim w2_ As String = Regex.Match(weight_, "\(.*\)").Value
            '            w2_ = Regex.Replace(w2_, "\(|\)", "")
            '            weight_ = Regex.Replace(weight_, "\(.*\)", "")

            '            '(ny263-51)3便以上宅配便
            '            If code_(0).ToLower = "ny263-51" Then
            '                weight_ = "60"
            '            End If

            '            'デフォルトがメール便で、宅配便があったら、haisouKindを宅配便にする
            '            'メール便と定形外の場合は、宅配便にする
            '            If (haisouKind_ = "メール便" And Regex.IsMatch(sw_, "P|p|T|t")) Or (haisouKind_ = "定形外" And Regex.IsMatch(sw_, "P|p|M|m")) Then
            '                maeKind_ = haisouKind_
            '                Select Case maeKind_
            '                    Case "メール便"
            '                        haisouSize_ = haisouSize_ / 100   'メール便サイズ計算を宅配便に変更
            '                    Case "定形外"
            '                        haisouSize_ = haisouSize_ / 75   '定形外サイズ計算を宅配便に変更
            '                End Select
            '            End If

            '            'haisouKindが宅配便の時、メール便を1/100にする
            '            Dim henkouFlag_ As Boolean = False
            '            If Regex.IsMatch(sw_, "M|m|T|t") Then
            '                henkouFlag_ = True
            '            End If

            '            'バラで保存（とりあえず定形外用で使用）
            '            For j As Integer = 1 To juchusu_
            '                haisouSizeArray_.Add(weight_)
            '            Next

            '            '重さ・特殊の処理
            '            Dim sp As String = DGV6.Item(dH6.IndexOf("特殊"), dgv6CodeArray.IndexOf(code_(0).ToLower)).Value
            '            If InStr(sp, "重") > 0 Then
            '                Dim omotoku As Double = OmosaCheck(code_(0).ToLower, juchusu_, haisouSize_)
            '                If omotoku <> 0 Then
            '                    haisouSize_ = haisouSize_ + omotoku
            '                End If
            '            ElseIf InStr(sp, "特") > 0 Then
            '                Dim omotoku As Double = TokusyuCheck(code_(0).ToLower, juchusu_, haisouSize_)
            '                If omotoku <> 0 Then
            '                    haisouSize_ = haisouSize_ + omotoku
            '                End If
            '            Else
            '                '-------
            '                If Not IsNumeric(weight_) Then
            '                    DGV1.Item(dH1.IndexOf("マスタ配送"), gyono).Value = "計算不能"
            '                    DGV1.Item(dH1.IndexOf("マスタ配送"), gyono).Style.BackColor = Color.Red
            '                    Exit Sub
            '                End If
            '                If henkouFlag_ Then
            '                    haisouSize_ = haisouSize_ + ((CDbl(weight_) / 100) * CDbl(juchusu_))
            '                Else
            '                    If w2_ <> "" Then    'P160(98)、P100(3.3)は1個のパーセンテージ（優先）
            '                        haisouSize_ = haisouSize_ + (CDbl(w2_) * CDbl(juchusu_))
            '                    Else
            '                        '50個一便
            '                        If code_(0).ToLower = "ny261" Or code_(0).ToLower = "ny261-01" Or code_(0).ToLower = "ny261-02" Or code_(0).ToLower = "ny261-02a" Or code_(0).ToLower = "ny261-03" Or code_(0).ToLower = "ny263" Or code_(0).ToLower = "ny263-00a" Or code_(0).ToLower = "ny264" Then
            '                            haisouSize_ = haisouSize_ + (1.66 * CDbl(juchusu_))
            '                        ElseIf code_(0).ToLower = "mask01" Or code_(0).ToLower = "mask-wh" Or code_(0).ToLower = "mask05" Or code_(0).ToLower = "mask01-bl" Or code_(0).ToLower = "mask01-ko" Then
            '                            haisouSize_ = haisouSize_ + (2.5 * CDbl(juchusu_))
            '                        Else
            '                            haisouSize_ = haisouSize_ + (CDbl(TakuhaiPerConv(weight_)) * CDbl(juchusu_))
            '                        End If
            '                    End If
            '                End If
            '            End If
            '        Next

            '        If DGV1.Item(dH1.IndexOf("マスタ配送"), gyono).Value <> "計算不能" Then
            '            '便数計算（メール便と宅配便は100％で1便、定形外は250gで1便）
            '            '箱計算（100％の箱を用意し、入る箱に順次入れ、個数を計算する）
            '            'DGV1.Item(dH1.IndexOf("マスタ便数"), r1).Value = Math.Ceiling(haisouSize_ / 100)
            '            dgvT.Item(dHTemp.IndexOf(filed), dgvTNum).Value = Math.Ceiling(haisouSize_ / 100)
            '        End If

            '        DGV1.Item(dH1.IndexOf("sw"), gyono).Value = Math.Ceiling(haisouSize_)
            '    End If
            'ElseIf dgvT Is DGV9 Or dgvT Is DGV13 Then
            '    Dim dHTemp As ArrayList = TM_HEADER_GET(dgvT)
            '    Dim haisouSize_ As Double = 0
            '    Dim haisouSizeArray_ As New ArrayList    'バラで保存

            '    Dim codesArray As New ArrayList
            '    Dim swArray As New ArrayList
            '    Dim tokuArray As New ArrayList

            '    If DGV1.Item(dH1.IndexOf("商品マスタ"), gyono).Style.BackColor <> Color.Red Then
            '        Dim mCode_ As String = DGV1.Item(dH1.IndexOf("商品マスタ"), gyono).Value
            '        Dim mCodeArray_ As String() = Split(mCode_, "、")

            '        For i As Integer = 0 To mCodeArray_.Length - 1
            '            Dim code_ As String() = Split(mCodeArray_(i), "*")
            '            Dim sw_ As String = MasterWeight(code_(0).ToLower)
            '            Dim tag_ As String = MasterTag(code_(0).ToLower)
            '            Dim toku_ As String = MasterToku(code_(0).ToLower)
            '            Dim juchusu_ As String = code_(1)

            '            If IsNumeric(juchusu_) Then
            '                For r As Integer = 0 To juchusu_ - 1
            '                    codesArray.Add(code_)
            '                    swArray.Add(sw_)
            '                    tokuArray.Add(toku_)
            '                Next
            '            End If
            '        Next

            '        Dim res As String = ""
            '        If dgvT Is DGV9 Then
            '            res = Keisan("メール便", codesArray, swArray, tokuArray)
            '        ElseIf dgvT Is DGV13 Then
            '            res = Keisan("定形外", codesArray, swArray, tokuArray)
            '        End If
            '        'Console.WriteLine(res)
            '        'TextBox10.Text = Split(res, ",")(0) 捆包数
            '        'TextBox9.Text = Split(res, ",")(2) weight
            '        If res <> "" Then
            '            dgvT.Item(dHTemp.IndexOf(filed), dgvTNum).Value = Split(res, ",")(0)
            '            DGV1.Item(dH1.IndexOf("sw"), gyono).Value = Split(res, ",")(2)
            '        End If
            '    End If
            'End If
        End If



        '搬迁
        'yamato()




    End Sub

    Dim omosa As Double = 0
    Private Function OmosaCheck(code As String, juchusu As String, haisouSize As String)
        '重さは30kgまでで1便（NumericUpDown1.Value）
        Dim weight As String = ""
        'Dim dH6 As ArrayList = TM_HEADER_GET(dgv6)
        Dim sp As String = DGV6.Item(dH6.IndexOf("特殊"), dgv6CodeArray.IndexOf(code)).Value
        omosa = Replace(sp, "重", "")
        weight = (Math.Floor(100 / NumericUpDown1.Value) * omosa) * juchusu
        Return weight
    End Function

    Dim tokuGroup As New ArrayList
    Private Function TokusyuCheck(code As String, juchusu As String, haisouSize As String)
        '商品特定して便を決め打ちする場合（括弧内は特定グループ）
        '特(ad033)0-0-0-0-0-1　　　1個で160を1便
        '特(ad033)0-0-1-2-0-0  　　1個なら100、2個なら120
        '特(ad033)0-0-2-4-0-0　2個まで100、3～4個は120
        '箱が140までなら追加で同梱が可能な処理が必要！！！！！！！！！！！！！！！！！

        Dim weight As Double = 0
        'Dim dH6 As ArrayList = TM_HEADER_GET(dgv6)
        Dim sizeArray As String() = New String() {"60", "80", "100", "120", "140", "160"}
        Dim sp As String = DGV6.Item(dH6.IndexOf("特殊"), dgv6CodeArray.IndexOf(code)).Value

        Dim toku As String = Replace(sp, "特", "")
        Dim rr As New Regex("\(.*\)", RegexOptions.IgnoreCase)
        Dim mm As Match = rr.Match(toku)
        Dim tGroup As String = Regex.Replace(mm.Value, "\(|\)", "") 'グループを取得する

        toku = Replace(toku, mm.Value, "")  '取得した残り
        Dim tokuArray As String() = Split(toku, "-")    '特別設定の中身

        '一番大きい入る個数を調べる
        Dim maxToku As Integer = 0
        For i As Integer = 0 To tokuArray.Length - 1
            If maxToku < tokuArray(i) Then
                maxToku = tokuArray(i)
            End If
        Next

        '---------------------------------------------------------------------
        'グループの処理
        'グループの商品があれば、個数を再計算用に取り出し、Arrayから削除する
        '---------------------------------------------------------------------
        Dim gWeight As Integer = 0
        Dim gKosu As Integer = 0
        For i As Integer = tokuGroup.Count - 1 To 0 Step -1
            Dim gCode As String() = Split(tokuGroup(i), "=")
            If gCode(0) = tGroup Then
                gWeight += gCode(1)
                gKosu += gCode(2)
                tokuGroup.RemoveAt(i)
            End If
        Next

        Dim kosu As Integer = 0
        If gWeight > 0 Then
            kosu = juchusu + gKosu
        Else
            kosu = juchusu
        End If

        '特殊設定の計算（新計算 2019/01/31）
        Dim hako As Integer = Math.Floor(kosu / maxToku)
        For i As Integer = 0 To hako - 1
            tokuGroup.Add(tGroup & "=100=" & maxToku)
        Next
        Dim nokori As Integer = kosu - (hako * maxToku)
        If nokori > 0 Then
            weight += (100 / maxToku) * nokori
            tokuGroup.Add(tGroup & "=" & weight & "=" & nokori)
        End If
        Dim weight1 As Double = (hako * 100) + ((100 / maxToku) * nokori)

        hako = Math.Floor((kosu - juchusu) / maxToku)
        nokori = (kosu - juchusu) - (hako * maxToku)
        Dim weight2 As Double = (hako * 100) + ((100 / maxToku) * nokori)

        'グループの個数で計算－今回計算の追加を引いた分＝差
        weight = weight1 - weight2

        '旧計算
        'While kosu > 0
        '    For i As Integer = 0 To tokuArray.Length - 1
        '        If tokuArray(i) <> 0 And tokuArray(i) >= kosu Or tokuArray(i) = maxToku Then
        '            'weight = weight + TakuhaiPerConv(sizeArray(j))
        '            weight += (100 / maxToku) * tokuArray(i)
        '            tokuGroup.Add(tGroup & "=" & weight & "=" & kosu)
        '            kosu = kosu - tokuArray(i)
        '            If kosu <= 0 Then
        '                Exit While
        '            End If
        '        End If
        '    Next
        'End While

        'チェック用
        'MsgBox(haisouSize & vbCrLf & tokuGroup.Count & vbCrLf & tokuGroup(tokuGroup.Count - 1) & vbCrLf & "kosu:" & kosu & vbCrLf & "weight:" & weight & vbCrLf & "gWeight:" & gWeight & vbCrLf & "gKosu:" & gKosu)

        '---------------------------------------------------------------------
        'グループの処理
        '前の商品がグループの商品なら、個数を追加し再計算して、差を引く
        '---------------------------------------------------------------------
        'weight = weight - gWeight

        Return weight
    End Function


    'YU2FlagMasterWeight 重新计算
    Public Function YU2FlagMasterWeight(code As String) As String
        Dim res As String = ""
        res = "M100"
        Return res
    End Function

    'マスタのweightを読み込む
    Private Function MasterWeight(code As String) As String
        Dim res As String = ""
        res = DGV6.Item(dH6.IndexOf("ship-weight"), dgv6CodeArray.IndexOf(code.ToLower)).Value


        'If CheckBox34.Checked And (YU2FlagLoad Or YU2Denpyus.Contains(CurDenpyoNo)) Then
        '    res = YU2FlagMasterWeight(code)
        'End If

        Return res
    End Function

    'マスタのタグを読み込む
    Private Function MasterTag(code As String) As String
        Dim res As String = ""
        res = DGV6.Item(dH6.IndexOf("商品分類タグ"), dgv6CodeArray.IndexOf(code.ToLower)).Value
        Return res
    End Function

    'マスタの商品名を読み込む
    Private Function MasterItemName(code As String) As String
        Dim res As String = ""
        res = DGV6.Item(dH6.IndexOf("商品名"), dgv6CodeArray.IndexOf(code.ToLower)).Value
        Return res
    End Function

    'マスタの商品名を読み込む
    Private Function MasterToku(code As String) As String
        Dim res As String = ""
        res = DGV6.Item(dH6.IndexOf("特殊"), dgv6CodeArray.IndexOf(code.ToLower)).Value
        Return res
    End Function

    '元データより枚数ベースで各伝票枚数の計算をする
    Private Sub MaisuKeisanLeft()
        LinkLabel7.Text = 0
        LinkLabel10.Text = 0
        LinkLabel28.Text = 0

        LinkLabel8.Text = 0
        LinkLabel11.Text = 0
        LinkLabel27.Text = 0
        DGV15.Rows.Clear()

        LinkLabel9.Text = 0
        LinkLabel12.Text = 0
        LinkLabel13.Text = 0

        LinkLabel14.Text = 0


        'ヤ(路) 太宰府
        LinkLabel31.Text = 0
        'ヤ(路) 井相田
        LinkLabel32.Text = 0

        'ゆう2(陸) 太宰府
        LinkLabel39.Text = 0
        'ゆう2(船) 太宰府
        LinkLabel40.Text = 0
        'ゆう2(陸) 井相田
        LinkLabel41.Text = 0
        'ゆう2(船) 井相田
        LinkLabel42.Text = 0


        Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)

        ToolStripProgressBar1.Maximum = DGV1.RowCount
        ToolStripProgressBar1.Value = 0
        Application.DoEvents()

        Dim tStr As String = ""
        For r As Integer = 0 To DGV1.RowCount - 1
            Dim dNo As String = DGV1.Item(dH1.IndexOf("伝票番号"), r).Value
            Dim hassou As String = DGV1.Item(dH1.IndexOf("発送方法"), r).Value
            Dim masutahaiso As String = DGV1.Item(dH1.IndexOf("マスタ配送"), r).Value
            Dim souko As String = DGV1.Item(dH1.IndexOf("発送倉庫"), r).Value
            Dim binsu As String = DGV1.Item(dH1.IndexOf("マスタ便数"), r).Value
            Dim YU2 As String = DGV1.Item(dH1.IndexOf("データ"), r).Value

            If IsNumeric(binsu) Then
                Select Case True
                    '太宰府
                    Case hassou = "宅配便" And souko = HS1.Text And YU2 <> "ゆう200"
                        LinkLabel7.Text += 1 'CInt(binsu) 
                        '井相田
                    Case hassou = "宅配便" And souko = HS2.Text And YU2 <> "ゆう200"
                        LinkLabel10.Text += 1 'CInt(binsu)
                        '名古屋
                    Case hassou = "宅配便" And souko = HS4.Text And YU2 <> "ゆう200"
                        LinkLabel28.Text += 1 'CInt(binsu)
                    'Case hassou = "宅配便" And souko = HS3.Text
                    '    LinkLabel13.Text += 1 'CInt(binsu)
                    Case hassou = "メール便" And souko = HS1.Text And YU2 <> "ゆう200"
                        LinkLabel8.Text += 1 'CInt(binsu)
                    Case hassou = "メール便" And souko = HS2.Text And YU2 <> "ゆう200"
                        LinkLabel11.Text += 1 'CInt(binsu)
                    'Case hassou = "メール便" And souko = HS3.Text
                    '    LinkLabel14.Text += 1 'CInt(binsu)
                    '    DGV15.Rows.Add(dNo, "メール便が" & HS3.Text)
                    Case hassou = "メール便" And souko = HS4.Text And YU2 <> "ゆう200"
                        LinkLabel27.Text += 1 'CInt(binsu)
                    Case hassou = "定形外" And souko = HS1.Text And YU2 <> "ゆう200"
                        LinkLabel9.Text += 1 'CInt(binsu)
                    Case hassou = "定形外" And souko = HS2.Text And YU2 <> "ゆう200"
                        LinkLabel12.Text += 1 'CInt(binsu)
                    Case hassou = "定形外" And souko = HS4.Text And YU2 <> "ゆう200"
                        LinkLabel13.Text += 1 'CInt(binsu)
                    Case masutahaiso = "ヤマト(陸便)" And souko = HS1.Text And YU2 <> "ゆう200"
                        LinkLabel31.Text += 1 'CInt(binsu)
                    Case masutahaiso = "ヤマト(船便)" And souko = HS1.Text And YU2 <> "ゆう200"
                        LinkLabel33.Text += 1 'CInt(binsu)
                    Case masutahaiso = "ヤマト(陸便)" And souko = HS2.Text And YU2 <> "ゆう200"
                        LinkLabel32.Text += 1 'CInt(binsu)
                    Case masutahaiso = "ヤマト(船便)" And souko = HS2.Text And YU2 <> "ゆう200"
                        LinkLabel36.Text += 1 'CInt(binsu)
                        '太宰府路便
                    Case YU2 = "ゆう200" And souko = HS1.Text
                        LinkLabel39.Text += 1 'CInt(binsu) 
                        '太宰府船便
                    Case YU2 = "ゆう200" And souko = HS1.Text
                        LinkLabel40.Text += 1 'CInt(binsu)
                        '井相田路便
                    Case YU2 = "ゆう200" And souko = HS2.Text
                        LinkLabel41.Text += 1 'CInt(binsu)
                        '井相田船便
                    Case YU2 = "ゆう200" And souko = HS2.Text
                        LinkLabel42.Text += 1 'CInt(binsu)
                    Case Else
                        LinkLabel14.Text += 1 'CInt(binsu)
                        DGV15.Rows.Add(dNo, "複数倉庫エラー")
                End Select
            Else
                LinkLabel14.Text += 1
                DGV15.Rows.Add(dNo, "便数文字エラー")
                tStr &= dNo & " "
            End If

            ToolStripProgressBar1.Value += 1
        Next

        ToolTip1.SetToolTip(LinkLabel14, tStr)
    End Sub

    '宅配便のサイズ別容量計算を読み込む
    Private Function TakuhaiPerConv(sizeStr As String) As String
        Dim dH10 As ArrayList = TM_HEADER_GET(DGV10)
        Dim res As String = ""


        For r As Integer = 0 To DGV10.RowCount - 1

            Dim DD = DGV10.Item(dH10.IndexOf("サイズ"), r).Value
            If Regex.IsMatch(DGV10.Item(dH10.IndexOf("サイズ"), r).Value, "^" & sizeStr & "$") Then
                res = DGV10.Item(dH10.IndexOf("容量計算"), r).Value
                Exit For
            End If
        Next
        Return res
    End Function

    Private Sub TextBox_KeyDown(sender As TextBox, e As KeyEventArgs) Handles _
            TextBox6.KeyDown, TextBox12.KeyDown,
            TextBox28.KeyDown, TextBox30.KeyDown, TextBox33.KeyDown, TextBox34.KeyDown, TextBox36.KeyDown
        If e.Modifiers = Keys.Control And e.KeyCode = Keys.A Then
            sender.SelectAll()
        End If
    End Sub

    Private Sub Csv_denpyo_FormClosing(ByVal sender As Object, ByVal e As Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If e.CloseReason = CloseReason.FormOwnerClosing Or e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
            Me.Visible = False
        End If
    End Sub

    Private Sub LIST4VIEW(ByRef mes As String, ByRef se As String)
        Dim seStr As String = ""
        Select Case se
            Case "start", "s"
                seStr = "START"
                ListBox4.Items.Add(mes & ">>" & seStr)
            Case "end", "e"
                seStr = "END"
                ListBox4.Items.Add("..." & seStr & ":" & mes)
            Case "error", "r"
                seStr = "ERROR"
                ListBox4.Items.Add("[" & seStr & "]:" & mes)
            Case Else
                seStr = se
                ListBox4.Items.Add("  =>" & mes & "/" & seStr)
        End Select

        ListBox4.SelectedIndex = ListBox4.Items.Count - 1
        Application.DoEvents()
    End Sub

    Private Sub フォームのリセットToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles フォームのリセットToolStripMenuItem.Click
        Me.Dispose()
        Dim frm As Form = New Csv_denpyo3
        frm.Show()

        Csv_denpyo3_F_count.LB11.Text = 0
        Csv_denpyo3_F_count.LB12.Text = 0
        Csv_denpyo3_F_count.LB13.Text = 0
        Csv_denpyo3_F_count.LB14.Text = 0
        Csv_denpyo3_F_count.LB35.Text = 0
        Csv_denpyo3_F_count.LB39.Text = 0

        Csv_denpyo3_F_count.LB15.Text = 0
        Csv_denpyo3_F_count.LB16.Text = 0
        Csv_denpyo3_F_count.LB17.Text = 0
        Csv_denpyo3_F_count.LB18.Text = 0
        Csv_denpyo3_F_count.LB36.Text = 0
        Csv_denpyo3_F_count.LB40.Text = 0

        Csv_denpyo3_F_count.LB31.Text = 0
        Csv_denpyo3_F_count.LB32.Text = 0
        Csv_denpyo3_F_count.LB33.Text = 0
        Csv_denpyo3_F_count.LB34.Text = 0

        Csv_denpyo3_F_count.LB47.Text = 0
        Csv_denpyo3_F_count.LB48.Text = 0
    End Sub

    Private Sub 配送情報ダウンロードのリセットToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 配送情報csvのリセットToolStripMenuItem.Click
        DGV1.Rows.Clear()
        DGV7.Rows.Clear()
        DGV8.Rows.Clear()
        DGV9.Rows.Clear()
        DGV13.Rows.Clear()
        Panel8.Show()
    End Sub

    Private Sub 受注明細データのリセットToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 受注明細csvのリセットToolStripMenuItem.Click
        DGV3.Rows.Clear()
        Panel3.Show()
    End Sub



    '============================================
    '--------------------------------------------
    'DataGridView コピー貼付けなど
    '--------------------------------------------
    '============================================
    '修正履歴取得
    Private CellErrFlg As Boolean = False
    Private CellValue
    Private Sub DataGridView_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles _
        DGV1.DataError, DGV7.DataError, DGV8.DataError, DGV9.DataError, DGV13.DataError
        CellErrFlg = True
    End Sub

    Private Sub DataGridView_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles _
        DGV1.CellBeginEdit, DGV7.CellBeginEdit, DGV8.CellBeginEdit, DGV9.CellBeginEdit, DGV13.CellBeginEdit
        Dim dgv As DataGridView = sender
        CellErrFlg = False
        CellValue = dgv.CurrentCell.Value   '該当セルの値を格納しておく
    End Sub

    Private Sub DataGridView_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles _
        DGV1.CellEndEdit, DGV7.CellEndEdit, DGV8.CellEndEdit, DGV9.CellEndEdit, DGV13.CellEndEdit
        Dim dgv As DataGridView = sender

        '変更したセルに色をつける
        If CellValue <> dgv.CurrentCell.Value Then
            dgv.CurrentCell.Style.BackColor = Color.Yellow
            If dgv Is DGV1 Then
                ValidateChangeSave(DGV1, e.ColumnIndex, e.RowIndex)
                TextBox39.Text = "変更"
                TextBox39.BackColor = Color.Yellow
            End If
        End If
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
        DGV1.EditingControlShowing, DGV7.EditingControlShowing, DGV8.EditingControlShowing, DGV9.EditingControlShowing, DGV13.EditingControlShowing
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

    Private Sub DataGridView_KeyUp(ByVal sender As Object, ByVal e As Windows.Forms.KeyEventArgs) Handles _
        DGV1.KeyUp, DGV7.KeyUp, DGV8.KeyUp, DGV9.KeyUp, DGV13.KeyUp
        Dim dgv As DataGridView = CType(sender, DataGridView)
        Dim selCell = dgv.SelectedCells

        If e.KeyCode = Keys.Back Then    ' セルの内容を消去
            If DGV1.IsCurrentCellInEditMode Then

            Else
                DELS(dgv, selCell)
            End If
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
                innerTextBox.Cut()
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
                innerTextBox.Copy()
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

        Dim clipTextArray As String() = Split(clipText, vbCrLf)
        For i As Integer = 0 To clipTextArray.Length - 1
            Dim flag As Boolean = False
            Dim rs As New StringReader(clipTextArray(i))
            Using tfp As New TextFieldParser(rs)
                tfp.TextFieldType = FieldType.Delimited
                tfp.Delimiters = New String() {vbTab, ","}
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
                    If dgv.Item(selCell(i).ColumnIndex, selCell(i).RowIndex).visible Then
                        dgv(selCell(i).ColumnIndex, selCell(i).RowIndex).Value = lines(0)
                        dgv(selCell(i).ColumnIndex, selCell(i).RowIndex).style.backcolor = Color.Yellow
                    End If
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
                    'If y + r = dgv.rowCount - 1 And _
                    '   dgv.AllowUserToAddRows = True Then
                    '    dgv.rowCount = dgv.rowCount + 1
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
                    If y + r = dgv.RowCount - 1 And dgv.AllowUserToAddRows = True Then
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

        Dim start As Integer = 0
        If mode <> 0 Then
            start = 1
        End If

        For i As Integer = 0 To selCol.Count - 1
            For r As Integer = start To dgv.RowCount - 1
                If Not dgv.Rows(r).IsNewRow Then
                    If dgv.Rows(r).Visible Then
                        dgv.Item(selCol(i), r).selected = True
                    End If
                End If
            Next
        Next
    End Sub

    Private Sub ROWSCUT(ByVal dgv As DataGridView, ByVal selCell As Object)
        Dim delR As New ArrayList
        For i As Integer = 0 To selCell.Count - 1
            If Not delR.Contains(selCell(i).RowIndex) Then
                If dgv.Rows(selCell(i).rowindex).Visible Then
                    delR.Add(selCell(i).RowIndex)
                End If
            End If
        Next
        delR.Sort()
        'delR.Reverse()

        '伝票番号を調べて履歴を残す
        If dgv Is DGV1 Then
            Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
            Dim dH14 As ArrayList = TM_HEADER_GET(DGV14)
            Dim delCount As Integer = 0
            For i As Integer = 0 To delR.Count - 1
                Dim dNo As String = DGV1.Item(dH1.IndexOf("伝票番号"), delR(i)).Value
                For r14 As Integer = 0 To DGV14.RowCount - 1
                    If DGV14.Item(dH14.IndexOf("伝票番号"), r14).Value = dNo Then
                        DGV14.Item(dH14.IndexOf("処理"), r14).Value = "行削除"
                        delCount += 1

                        '別紙PDFも調べる
                        For k As Integer = ListBox3.Items.Count - 1 To 0 Step -1
                            If ListBox3.Items(k) = dNo Then
                                ListBox3.Items.RemoveAt(k)
                                DGV17.Rows.RemoveAt(k)
                            End If
                        Next
                    End If
                Next
                If delCount = 0 Then
                    DGV14.Rows.Add(dNo, "", "手動", "行削除")
                End If
            Next
        End If

        'データ処理を軽くするため画面を一時的に消す
        If delR.Count > 100 Then
            dgv.Visible = False
            Application.DoEvents()
        End If

        dgv.ClearSelection()

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

    '============================================
    '--------------------------------------------
    'コンテキストメニュー
    '--------------------------------------------
    '============================================
    '子項目をクリックした時、SourceControlがnullになるので先に取得
    Private contextMenuSourceControl As Control = Nothing
    Private Sub ContextMenuStrip1_Opened(sender As Object, e As EventArgs) Handles _
            ContextMenuStrip1.Opened
        Dim menu As ContextMenuStrip = DirectCast(sender, ContextMenuStrip)
        'SourceControlプロパティの内容を覚えておく
        contextMenuSourceControl = menu.SourceControl
    End Sub

    Private Sub 切り取りToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 切り取りToolStripMenuItem1.Click
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        CUTS(FocusDGV)
    End Sub

    Private Sub コピーToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles コピーToolStripMenuItem1.Click
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        COPYS(FocusDGV, 0)
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        COPYS(FocusDGV, 1)
    End Sub

    Private Sub 貼り付けToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 貼り付けToolStripMenuItem1.Click
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        PASTES(FocusDGV, selCell)
    End Sub

    Private Sub 上に挿入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 上に挿入ToolStripMenuItem.Click
        If Not contextMenuSourceControl Is Nothing Then
            FocusDGV = contextMenuSourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        ROWSADD(FocusDGV, selCell, 1)
    End Sub

    Private Sub 下に挿入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 下に挿入ToolStripMenuItem.Click
        If Not contextMenuSourceControl Is Nothing Then
            FocusDGV = contextMenuSourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        ROWSADD(FocusDGV, selCell, 0)
    End Sub

    Private Sub 行を選択直下に複製ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 行を選択直下に複製ToolStripMenuItem.Click
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
                If CStr(FocusDGV.Item(c, dgRow(r)).Value) <> "" Then
                    FocusDGV.Item(c, addRow).Value = FocusDGV.Item(c, dgRow(r)).Value
                End If
            Next
        Next
    End Sub

    Private Sub 右に挿入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 右に挿入ToolStripMenuItem.Click
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        COLSADD(FocusDGV, selCell, 0)
    End Sub

    Private Sub 左に挿入ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 左に挿入ToolStripMenuItem.Click
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        COLSADD(FocusDGV, selCell, 1)
    End Sub

    Private Sub 削除ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 削除ToolStripMenuItem1.Click
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        DELS(FocusDGV, selCell)
    End Sub

    Private Sub 背景色を消すToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 背景色を消すToolStripMenuItem.Click
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        For i As Integer = 0 To FocusDGV.SelectedCells.Count - 1
            FocusDGV.SelectedCells(i).Style.BackColor = DefaultBackColor
        Next
    End Sub

    Private Sub 列選択ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 列選択ToolStripMenuItem1.Click
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        selChangeFlag = False
        ColSelect(FocusDGV, 0)
        selChangeFlag = True
    End Sub

    Private Sub 行を削除ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 行を削除ToolStripMenuItem1.Click
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        ROWSCUT(FocusDGV, selCell)
    End Sub

    Private Sub 列を削除ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 列を削除ToolStripMenuItem1.Click
        If Not sender.Owner.SourceControl Is Nothing Then
            FocusDGV = sender.Owner.SourceControl
        End If
        Dim selCell = FocusDGV.SelectedCells
        COLSCUT(FocusDGV, selCell)
    End Sub



    '============================================
    '--------------------------------------------
    '配送マスタ
    '--------------------------------------------
    '============================================
    'dgv6加载
    Private Sub DGV6_update()
        'datagridview6再表示
        If DGV6.RowCount > 0 Then
            DGV6.Rows.Clear()
            DGV6.Columns.Clear()
        End If

        'ヘッダー行作成
        Dim header As String() = New String() {"商品コード", "商品名", "商品分類タグ", "代表商品コード", "ship-weight", "ロケーション", "ロケーション(名古屋)", "梱包サイズ", "特殊", "sw2"}
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
            sArray(0) = RTrim(sArray(0))
            sArray(0) = Form1.StrConvToNarrow(sArray(0))    '商品コードを小文字で揃える


            DGV6.Rows.Add(sArray)
        Next

        dgv6CodeArray.Clear()
        dgv6DaihyoCodeArray.Clear()
        Dim dH6 As ArrayList = TM_HEADER_GET(DGV6)
        For r As Integer = 0 To DGV6.RowCount - 1
            Dim tempcode = DGV6.Item(dH6.IndexOf("商品コード"), r).Value.ToString.ToLower
            dgv6CodeArray.Add(tempcode)
            dgv6DaihyoCodeArray.Add(DGV6.Item(dH6.IndexOf("代表商品コード"), r).Value.ToString.ToLower)
        Next
    End Sub

    Dim helpNumber As Integer = 0
    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        Dim ss As Integer = Format(Now, "ss")   '秒指定して動かす（未使用）
        If CInt(ss) Mod 10 = 0 Then
            ToolStripLabel2.Text = helpMessage(helpNumber)
            If helpNumber = helpMessage.Count - 1 Then
                helpNumber = 0
            Else
                helpNumber += 1
            End If
        End If

        ToolStripStatusLabel3.Text = Format(Now, "HH:mm:ss")
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim ss As Integer = Format(Now, "ss")   '秒指定して動かす（未使用）

        ToolStripStatusLabel1.Text = Timer1.Interval / 1000
        'sakiPath = Form1.サーバーToolStripMenuItem.Text & "\update"
        sakiPath2 = Form1.ロケーションToolStripMenuItem.Text

        Dim files2 As FileInfo() = Nothing

        Dim fi1 As New FileInfo(Path.GetDirectoryName(Form1.appPath) & "\location.csv")
        TextBox24.Text = fi1.LastWriteTime.ToOADate


        'サーバーの接続を確認する
        If File.Exists(sakiPath2 & "\location.csv") Then
            'sakiPath = "Z:\中村作成データ\AutoCheck\update"
            'Dim di As New System.IO.DirectoryInfo(sakiPath & "\version2")
            'files2 = di.GetFiles("*.dat", System.IO.SearchOption.AllDirectories)
            Dim fi2 As New FileInfo(sakiPath2 & "\location.csv")
            Dim check = fi2.LastAccessTime
            ToolStripLabel1.Text = "接続OK"
            ToolStripLabel1.ForeColor = Color.Black
            ToolStripLabel1.BackColor = Color.Empty
            'KryptonButton3.Enabled = True
            Panel6.Enabled = True
        Else
            ToolStripLabel1.Text = "未接続"
            ToolStripLabel1.ForeColor = Color.White
            ToolStripLabel1.BackColor = Color.Red
            If LockFlag Then    'ロック中は処理できないようにする
                KryptonButton3.Enabled = False
                KryptonButton1.Enabled = False
            End If
            Panel6.Enabled = False

            Exit Sub
        End If

        Dim flag As Integer = False

        Try
            Dim fi2 As New FileInfo(sakiPath2 & "\location.csv")
            TextBox25.Text = fi2.LastWriteTime.ToOADate
            If TextBox25.Text < 0 Or TextBox24.Text < 0 Then
                TextBox17.Text = "error"
                Timer1.Interval = 1000 * 5
                ToolStripStatusLabel1.Text = Timer1.Interval / 1000
            ElseIf TextBox25.Text - TextBox24.Text > 0 Then
                Form1.NotifyIcon1.BalloonTipTitle = "配送マスタのアップデートを感知しました"
                Form1.NotifyIcon1.BalloonTipText = "自動でアップデートを行ないます。再起動の必要はありません。"
                Form1.NotifyIcon1.ShowBalloonTip(5000)

                TextBox17.Text = "update"
                Dim sPath As String = sakiPath2 & "\location.csv"
                Dim mPath As String = Path.GetDirectoryName(Form1.appPath) & "\location.csv"
                File.Copy(sPath, mPath, True)
                Timer1.Interval = 1000 * 5
                ToolStripStatusLabel1.Text = Timer1.Interval / 1000

                'datagridview6再表示
                DGV6_update()
                Application.DoEvents()
            Else
                If ToolStripStatusLabel1.Text <> 180 Then
                    Form1.NotifyIcon1.BalloonTipTitle = "配送マスタ状態は最新です"
                    Form1.NotifyIcon1.BalloonTipText = "伝票処理の作業が可能です。"
                    Form1.NotifyIcon1.ShowBalloonTip(5000)
                End If

                TextBox17.Text = "OK"
                Timer1.Interval = 1000 * 60 * 3
                ToolStripStatusLabel1.Text = Timer1.Interval / 1000
                flag = True

                'datagridview6再表示
                DGV6_update()
                Application.DoEvents()
            End If
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub

    'マスタ検索
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox6.Text = "" Then
            Exit Sub
        End If

        For r As Integer = 0 To DGV6.RowCount - 1
            For c As Integer = 0 To DGV6.ColumnCount - 1
                If CheckBox2.Checked = False Then
                    If TextBox6.Text = DGV6.Item(c, r).Value Then
                        DGV6.CurrentCell = DGV6(c, r)
                        Exit Sub
                    End If
                Else
                    If InStr(DGV6.Item(c, r).Value, TextBox6.Text) > 0 Then
                        DGV6.CurrentCell = DGV6(c, r)
                        Exit Sub
                    End If
                End If
            Next
        Next
        MsgBox("検索結果がありません。", MsgBoxStyle.SystemModal)
    End Sub

    Private Sub TextBox6_KeyDown(ByVal sender As Object, ByVal e As Windows.Forms.KeyEventArgs) Handles TextBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button3.PerformClick()
        End If
    End Sub

    'リセット
    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        For r As Integer = 0 To DGV6.RowCount - 1
            If DGV6.Rows(r).Visible = False Then
                DGV6.Rows(r).Visible = True
            End If
        Next
    End Sub

    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button25.Click
        DGV6_update()
    End Sub

    Private Sub DataGridView6_SelectionChanged(sender As Object, e As EventArgs) Handles DGV6.SelectionChanged
        If DGV6.RowCount = 0 Or DGV6.SelectedCells.Count = 0 Then
            Exit Sub
        End If

        TextBox26.Text = DGV6.SelectedCells(0).Value

        Dim res As String = DGV6.Item(0, DGV6.SelectedCells(0).RowIndex).Value
        Dim code As String() = Split(res, "-")
        Dim url As String = ""
        Dim url1 As String = ""
        'url1 = "http://gigaplus.makeshop.jp/manpou/item/"
        url1 = "https://shopping.c.yimg.jp/lib/fkstyle/"
        'url1 = "https://shopping.c.yimg.jp/lib/lucky9/"
        'url1 = "https://image.rakuten.co.jp/patri/cabinet/pc/"
        'url1 = "https://image.rakuten.co.jp/alice-zk/cabinet/pc/"
        url = url1 & code(0) & ".jpg"
        PictureBox1.ImageLocation = url
    End Sub





    '============================================
    '--------------------------------------------
    'メイン左
    '--------------------------------------------
    '============================================
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles _
            CheckBox1.CheckedChanged, CheckBox3.CheckedChanged, CheckBox4.CheckedChanged,
            CheckBox17.CheckedChanged, CheckBox18.CheckedChanged, CheckBox32.CheckedChanged
        CheckDenpyoChange(sender)
    End Sub

    Private Function CheckDenpyoChange(sender As CheckBox)
        If sender.Checked Then
            sender.Text = "e飛伝2"
        Else
            sender.Text = "BIZlogi"
        End If

        Return True
    End Function

    Private Sub Button36_Click(sender As Object, e As EventArgs) Handles Button36.Click
        Dim dgv As DataGridView = DGV1
        Dim headerStr As String = ""
        Dim headerStr2 As String = ""
        Dim headerStr3 As String = ""
        Dim headerStr4 As String = ""
        Select Case TabControl2.SelectedTab.Text
            Case "佐川e飛伝2"
                dgv = DGV7
                headerStr = koumoku("denpyoNo")(1)
                headerStr2 = koumoku("syori2")(1)
                headerStr3 = koumoku("binsyu")(1)
                headerStr4 = koumoku("sakiname")(1)
            Case "佐川BIZlogi"
                dgv = DGV8
                headerStr = koumoku("denpyoNo")(2)
                headerStr2 = koumoku("syori2")(2)
                headerStr3 = koumoku("binsyu")(2)
                headerStr4 = koumoku("sakiname")(2)
            Case "メール便"
                dgv = DGV9
                headerStr = koumoku("denpyoNo")(3)
                headerStr2 = koumoku("syori2")(3)
                headerStr3 = koumoku("binsyu")(3)
                headerStr4 = koumoku("sakiname")(3)
            Case "定形外"
                dgv = DGV13
                headerStr = koumoku("denpyoNo")(4)
                headerStr2 = koumoku("syori2")(4)
                headerStr3 = koumoku("binsyu")(4)
                headerStr4 = koumoku("sakiname")(4)
            Case "TMS"
                dgv = TMSDGV
                headerStr = koumoku("denpyoNo")(6)
                headerStr2 = koumoku("syori2")(6)
                headerStr3 = koumoku("binsyu")(6)
                headerStr4 = koumoku("sakiname")(6)
            Case Else
                dgv = DGV1
                headerStr = koumoku("denpyoNo")(0)
                headerStr2 = koumoku("syori2")(0)
                headerStr3 = koumoku("binsyu")(0)
                headerStr4 = koumoku("sakiname")(0)
        End Select

        If dgv.SelectedCells.Count > 0 Then
            Dim selRowArray As ArrayList = New ArrayList
            For i As Integer = 0 To dgv.SelectedCells.Count - 1
                If Not selRowArray.Contains(dgv.SelectedCells(i).RowIndex) Then
                    selRowArray.Add(dgv.SelectedCells(i).RowIndex)
                End If
            Next

            'datagridviewのヘッダーテキストをコレクションに取り込む
            Dim dH As ArrayList = TM_HEADER_GET(dgv)

            For i As Integer = 0 To selRowArray.Count - 1
                Dim dNo As String = dgv.Item(TM_ArIndexof(dH, headerStr), selRowArray(i)).Value
                Dim hmei As String = dgv.Item(TM_ArIndexof(dH, headerStr4), selRowArray(i)).Value
                Dim souko As String = dgv.Item(TM_ArIndexof(dH, headerStr2), selRowArray(i)).Value
                Dim binsyu As String = dgv.Item(TM_ArIndexof(dH, headerStr3), selRowArray(i)).Value
                If dNo <> "" Then
                    ListBox3.Items.Add(dNo)
                    DGV17.Rows.Add(dNo, hmei, souko, binsyu)
                End If
            Next

            MsgBox("PDFリストに伝票番号を追加しました", MsgBoxStyle.Information Or MsgBoxStyle.SystemModal)
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ListBox3.SelectedItems.Count > 0 Then
            Dim r As Integer = ListBox3.SelectedIndex
            ListBox3.Items.RemoveAt(r)
            DGV17.Rows.RemoveAt(r)
        End If
    End Sub

    Private Sub 別紙PDFのみ出力ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 別紙PDFのみ出力ToolStripMenuItem.Click
        Dim path As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) & "\"
        PDFsave(path)
    End Sub

    Private Sub PDFsave(path_ As String, Optional mode As Boolean = True)
        If ListBox3.Items.Count = 0 Then
            If mode = True Then
                MsgBox("出力可能なリストがありません", MsgBoxStyle.Information Or MsgBoxStyle.SystemModal)
            End If
            Exit Sub
        End If



        ListBox3.Items.Clear()

        For Each item As String In denpyounoS
            ListBox3.Items.Add(item)
        Next

        Dim soukoName As String() = New String() {HS1.Text, HS1.Text & "航空", HS2.Text, HS2.Text & "航空", HS3.Text, HS3.Text & "航空", HS4.Text, HS4.Text & "航空", "倉庫指定無し"}
        Dim printStrArray1 As New ArrayList     '太宰府
        Dim printStrArray2 As New ArrayList     '太宰府航空
        Dim printStrArray3 As New ArrayList     '井相田
        Dim printStrArray4 As New ArrayList     '井相田航空
        Dim printStrArray5 As New ArrayList     '物流
        Dim printStrArray6 As New ArrayList     '物流航空
        Dim printStrArray8 As New ArrayList     '名古屋
        Dim printStrArray9 As New ArrayList     '名古屋航空
        Dim printStrArray7 As New ArrayList     '倉庫指定無し
        Dim printHinmeiArray1 As New ArrayList     '太宰府
        Dim printHinmeiArray2 As New ArrayList     '太宰府航空
        Dim printHinmeiArray3 As New ArrayList     '井相田
        Dim printHinmeiArray4 As New ArrayList     '井相田航空
        Dim printHinmeiArray5 As New ArrayList     '物流
        Dim printHinmeiArray6 As New ArrayList     '物流航空
        Dim printHinmeiArray8 As New ArrayList     '名古屋
        Dim printHinmeiArray9 As New ArrayList     '名古屋航空
        Dim printHinmeiArray7 As New ArrayList     '倉庫指定無し
        For i As Integer = 0 To ListBox3.Items.Count - 1
            'datagridviewのヘッダーテキストをコレクションに取り込む
            Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)

            For r As Integer = 0 To DGV1.Rows.Count - 1
                If ListBox3.Items(i) = DGV1.Item(TM_ArIndexof(dH1, koumoku("denpyoNo")(0)), r).Value Then
                    Dim str As String = ""
                    'str &= DGV1.Item(TM_ArIndexof(dH1, "商品コード1"), r).Value & vbCrLf
                    'str &= vbCrLf
                    str &= "店舗／" & DGV1.Item(TM_ArIndexof(dH1, koumoku("irai3")(0)), r).Value & vbCrLf
                    str &= vbCrLf
                    str &= "伝票番号／" & DGV1.Item(TM_ArIndexof(dH1, koumoku("denpyoNo")(0)), r).Value & vbCrLf

                    Dim str2 As String = ""
                    Dim hinmeiStr As String = ""
                    Dim cArray As String() = Split(DGV1.Item(TM_ArIndexof(dH1, "商品マスタ"), r).Value, "、")
                    Dim souko As String = DGV1.Item(TM_ArIndexof(dH1, "発送倉庫"), r).Value
                    Array.Sort(cArray)
                    Dim kosu As Integer = 0
                    For k As Integer = 0 To cArray.Length - 1
                        Dim code As String() = Split(cArray(k), "*")

                        '商品コードから品名取得
                        Dim hinmei As String = GET_Hinmei(code(0))

                        kosu += CInt(code(1))
                        Dim localCode As String = ValiationAdd1(code(0), souko, code(1), 1)
                        Dim localCode_ As String() = Split(localCode, "|")
                        localCode = Replace(localCode_(0), "*", " * ")
                        str2 &= "〈" & k + 1 & "〉 " & localCode & vbCrLf
                        hinmeiStr &= " 【" & hinmei & "】" & vbCrLf
                    Next

                    str &= "お届け先名／" & DGV1.Item(TM_ArIndexof(dH1, koumoku("sakiname")(0)), r).Value & "　（注文個数計／" & kosu & "）" & vbCrLf
                    str &= "お届け先住所／" & DGV1.Item(TM_ArIndexof(dH1, koumoku("sakiAddr1")(0)), r).Value & vbCrLf

                    If DGV1.Item(TM_ArIndexof(dH1, "便種"), r).Value = "航空便" Then
                        souko &= "航空"
                    End If
                    str &= "発送元倉庫／" & souko & "　【 発送方法／" & DGV1.Item(TM_ArIndexof(dH1, "発送方法"), r).Value & " 】" & vbCrLf
                    str &= vbCrLf

                    str &= str2
                    str &= vbCrLf
                    str &= DGV1.Item(TM_ArIndexof(dH1, "納品書備考"), r).Value & vbCrLf

                    If souko = HS1.Text Then
                        printStrArray1.Add(str)
                        printHinmeiArray1.Add(hinmeiStr)
                    ElseIf souko = HS1.Text & "航空" Then
                        printStrArray2.Add(str)
                        printHinmeiArray2.Add(hinmeiStr)
                    ElseIf souko = HS2.Text Then
                        printStrArray3.Add(str)
                        printHinmeiArray3.Add(hinmeiStr)
                    ElseIf souko = HS2.Text & "航空" Then
                        printStrArray4.Add(str)
                        printHinmeiArray4.Add(hinmeiStr)
                    ElseIf souko = HS3.Text Then
                        printStrArray5.Add(str)
                        printHinmeiArray5.Add(hinmeiStr)
                    ElseIf souko = HS3.Text & "航空" Then
                        printStrArray6.Add(str)
                        printHinmeiArray6.Add(hinmeiStr)
                    ElseIf souko = HS4.Text Then
                        printStrArray8.Add(str)
                        printHinmeiArray8.Add(hinmeiStr)
                    ElseIf souko = HS4.Text & "航空" Then
                        printStrArray9.Add(str)
                        printHinmeiArray9.Add(hinmeiStr)
                    Else
                        printStrArray7.Add(str)
                        printHinmeiArray7.Add(hinmeiStr)
                    End If
                    Exit For
                End If
            Next
        Next

        ''デバッグ用メモリ描画
        'For lineNum As Integer = 0 To 20
        '    Dim textStyle1 = New TextStyle("メイリオ", 7)
        '    Dim content1() As AbstractPdfContentItem = {New Text(0, 50 * lineNum, 10, 10, (50 * lineNum) - 842, textStyle1)}
        '    docPDF.AddContent("page" & pageNum, content1)
        'Next

        'ログイン名前取得
        Dim loginName As String = "(" & Replace(Form1.ログイン名ToolStripMenuItem.Text, "さん", "") & ")"
        Dim fname As String = path_ & "●別紙"

        Dim newName As String = ""
        Dim psArray As Object() = New Object() {printStrArray1, printStrArray2, printStrArray3, printStrArray4, printStrArray5, printStrArray6, printStrArray8, printStrArray9, printStrArray7}
        Dim phArray As Object() = New Object() {printHinmeiArray1, printHinmeiArray2, printHinmeiArray3, printHinmeiArray4, printHinmeiArray5, printHinmeiArray6, printHinmeiArray8, printHinmeiArray9, printHinmeiArray7}
        Dim A4width As Double = 595
        Dim A4Height As Double = 842
        'For k As Integer = 0 To 4
        For k As Integer = 0 To 8
            If psArray(k).count > 0 Then
                Dim docPDF As New PdfDocument
                Dim windir = Environment.ExpandEnvironmentVariables("%windir%")
                docPDF.AddFont("メイリオ", windir + "\Fonts\meiryo.ttc", 0)
                docPDF.AddPage("page1")
                Dim pageNum As Integer = 1

                For i As Integer = 0 To psArray(k).Count - 1
                    '行数を調べる
                    Dim gyou As String() = Split(phArray(k)(i), vbCrLf)     '1ページ31行
                    Dim hPos1 As Integer = A4Height - 100
                    Dim hPos2 As Integer = A4Height - 259
                    Dim hPos3 As Integer = A4Height - 252
                    Dim hHeight As Double = 22.5
                    Dim hMax As Integer = 22
                    Dim yokohaba As Double = 540

                    If gyou.Length >= 34 Then
                        hPos2 = A4Height - 195
                        hPos3 = A4Height - 186
                        hHeight = 13.5
                        hMax = 44
                    ElseIf gyou.Length >= 24 Then
                        hPos2 = A4Height - 215
                        hPos3 = A4Height - 208
                        hHeight = 16.55
                        hMax = 33
                    End If

                    '横線
                    For lineNum As Integer = 0 To hMax
                        Dim content1() As AbstractPdfContentItem = {New Line(50, hPos3 - (hHeight * lineNum), yokohaba, hPos3 - (hHeight * lineNum))}
                        docPDF.AddContent("page" & pageNum, content1)
                    Next

                    '情報～商品
                    Dim doc As String() = Split(psArray(k)(i), vbCrLf)
                    For lineNum As Integer = 0 To doc.Length - 1
                        Dim textStyle = New TextStyle("メイリオ", 11)
                        Dim content2() As AbstractPdfContentItem = {New Text(50, hPos1 - (hHeight * lineNum), yokohaba, 30, doc(lineNum), textStyle)}
                        docPDF.AddContent("page" & pageNum, content2)
                    Next

                    '品名表示
                    Dim doc2 As String = phArray(k)(i)
                    Dim doc2Array As String() = Split(doc2, vbCrLf)
                    For lineNum As Integer = 0 To doc2Array.Length - 1
                        Dim textStyle2 = New TextStyle("メイリオ", 9)
                        Dim content3() As AbstractPdfContentItem = {New Text(400, hPos2 - (hHeight * lineNum), yokohaba, 30, doc2Array(lineNum), textStyle2)}
                        docPDF.AddContent("page" & pageNum, content3)
                    Next

                    If i <> psArray(k).Count - 1 Then
                        pageNum += 1
                        docPDF.AddPage("page" & pageNum)
                        ''デバッグ用メモリ描画
                        'For lineNum As Integer = 0 To 20
                        '    Dim textStyle1 = New TextStyle("メイリオ", 7)
                        '    Dim content1() As AbstractPdfContentItem = {New Text(0, 50 * lineNum, 10, 10, (50 * lineNum) - 842, textStyle1)}
                        '    docPDF.AddContent("page" & pageNum, content1)
                        'Next
                    End If
                Next

                Try
                    newName = fname & "(" & soukoName(k) & ")" & loginName & "_" & Format(Now, "yyyyMMddHHmmss") & ".pdf"
                    File.WriteAllBytes(newName, docPDF.GetBinary())
                    docPDF.Clear()
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                    Exit Sub
                End Try

                docPDF.Clear()
            End If
        Next

        If mode = True Then
            Dim ret As Integer = MessageBox.Show("作成したPDFを表示しますか？", "PDFの表示", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If ret = DialogResult.Yes Then
                Process.Start(newName)
            Else
                MessageBox.Show("PDFファイルを保存しました")
            End If
        End If
    End Sub

    Private Sub Excelsave_hacyu(saveList As ArrayList, dir As String, Optional mode As Boolean = True)
        If ListBox3.Items.Count = 0 Then
            If mode = True Then
                MsgBox("出力可能なリストがありません", MsgBoxStyle.Information Or MsgBoxStyle.SystemModal)
            End If
            Exit Sub
        End If

        Dim soukoName As String() = New String() {HS1.Text, HS1.Text & "航空", HS2.Text, HS2.Text & "航空", HS3.Text, HS3.Text & "航空", "倉庫指定無し"}
        Dim printStrArray1 As New ArrayList     '新宮
        Dim printStrArray1_list As New ArrayList
        Dim printStrArray2 As New ArrayList     '新宮航空
        Dim printStrArray2_list As New ArrayList
        Dim printStrArray3 As New ArrayList     '井相田
        Dim printStrArray3_list As New ArrayList
        Dim printStrArray4 As New ArrayList     '井相田航空
        Dim printStrArray4_list As New ArrayList
        Dim printStrArray5 As New ArrayList     '物流
        Dim printStrArray5_list As New ArrayList
        Dim printStrArray6 As New ArrayList     '物流航空
        Dim printStrArray6_list As New ArrayList
        Dim printStrArray7 As New ArrayList     '倉庫指定無し
        Dim printStrArray7_list As New ArrayList
        Dim printHinmeiArray1 As New ArrayList     '新宮
        Dim printHinmeiArray2 As New ArrayList     '新宮航空
        Dim printHinmeiArray3 As New ArrayList     '井相田
        Dim printHinmeiArray4 As New ArrayList     '井相田航空
        Dim printHinmeiArray5 As New ArrayList     '物流
        Dim printHinmeiArray6 As New ArrayList     '物流航空
        Dim printHinmeiArray7 As New ArrayList     '倉庫指定無し

        For i As Integer = 0 To ListBox3.Items.Count - 1
            'datagridviewのヘッダーテキストをコレクションに取り込む
            Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)

            For r As Integer = 0 To DGV1.Rows.Count - 1
                'If DGV1.Item(dH1.IndexOf("店舗"), r).Value = "とんよか卸" Then

                If DGV1.Item(dH1.IndexOf("店舗"), r).Value = "卸" Then
                    If ListBox3.Items(i) = DGV1.Item(TM_ArIndexof(dH1, koumoku("denpyoNo")(0)), r).Value Then
                        Dim printStrArray_info As New ArrayList
                        'Dim str As String = ""
                        '伝票番号
                        printStrArray_info.Add(DGV1.Item(TM_ArIndexof(dH1, koumoku("denpyoNo")(0)), r).Value)

                        Dim str2 As String = ""
                        Dim str2_ As New ArrayList
                        Dim hinmeiStr As String = ""
                        Dim hinmeiStr_ As New ArrayList
                        Dim cArray As String() = Split(DGV1.Item(TM_ArIndexof(dH1, "商品マスタ"), r).Value, "、")
                        Array.Sort(cArray)
                        Dim kosu As Integer = 0
                        Dim souko As String = DGV1.Item(TM_ArIndexof(dH1, "発送倉庫"), r).Value
                        For k As Integer = 0 To cArray.Length - 1
                            Dim code As String() = Split(cArray(k), "*")

                            '商品コードから品名取得
                            Dim hinmei As String = GET_Hinmei(code(0))

                            kosu += CInt(code(1))
                            Dim localCode As String = ValiationAdd1(code(0), souko, code(1), 1)
                            Dim localCode_ As String() = Split(localCode, "|")
                            localCode = Replace(localCode_(0), "*", " * ")
                            str2 = "〈" & k + 1 & "〉 " & localCode
                            str2_.Add(str2)

                            hinmeiStr_.Add(" 【" & hinmei & "】")
                        Next

                        'お届け先名
                        printStrArray_info.Add(DGV1.Item(TM_ArIndexof(dH1, koumoku("sakiname")(0)), r).Value & "　（注文個数計／" & kosu & "）")

                        'お届け先住所
                        printStrArray_info.Add(DGV1.Item(TM_ArIndexof(dH1, koumoku("sakiAddr1")(0)), r).Value)


                        If DGV1.Item(TM_ArIndexof(dH1, "便種"), r).Value = "航空便" Then
                            souko &= "航空"
                        End If

                        'お届け先住所
                        printStrArray_info.Add(souko & "　【 発送方法／" & DGV1.Item(TM_ArIndexof(dH1, "発送方法"), r).Value & " 】")

                        'str &= str2
                        'str &= vbCrLf
                        'str &= DGV1.Item(TM_ArIndexof(dH1, "納品書備考"), r).Value & vbCrLf

                        If souko = HS1.Text Then
                            printStrArray1.Add(printStrArray_info)
                            printStrArray1_list.Add(str2_)
                            printHinmeiArray1.Add(hinmeiStr_)
                        ElseIf souko = HS1.Text & "航空" Then
                            printStrArray2.Add(printStrArray_info)
                            printStrArray2_list.Add(str2_)
                            printHinmeiArray2.Add(hinmeiStr_)
                        ElseIf souko = HS2.Text Then
                            printStrArray3.Add(printStrArray_info)
                            printStrArray3_list.Add(str2_)
                            printHinmeiArray3.Add(hinmeiStr_)
                        ElseIf souko = HS2.Text & "航空" Then
                            printStrArray4.Add(printStrArray_info)
                            printStrArray4_list.Add(str2_)
                            printHinmeiArray4.Add(hinmeiStr_)
                        ElseIf souko = HS3.Text Then
                            printStrArray5.Add(printStrArray_info)
                            printStrArray5_list.Add(str2_)
                            printHinmeiArray5.Add(hinmeiStr_)
                        ElseIf souko = HS3.Text & "航空" Then
                            printStrArray6.Add(printStrArray_info)
                            printStrArray6_list.Add(str2_)
                            printHinmeiArray6.Add(hinmeiStr_)
                        Else
                            printStrArray7.Add(printStrArray_info)
                            printStrArray7_list.Add(str2_)
                            printHinmeiArray7.Add(hinmeiStr_)
                        End If
                    End If
                End If
            Next
        Next

        ''デバッグ用メモリ描画
        'For lineNum As Integer = 0 To 20
        '    Dim textStyle1 = New TextStyle("メイリオ", 7)
        '    Dim content1() As AbstractPdfContentItem = {New Text(0, 50 * lineNum, 10, 10, (50 * lineNum) - 842, textStyle1)}
        '    docPDF.AddContent("page" & pageNum, content1)
        'Next

        'ログイン名前取得
        Dim loginName As String = "(" & Replace(Form1.ログイン名ToolStripMenuItem.Text, "さん", "") & ")"
        'Dim fname As String = path_ & "●別紙"

        Dim newName As String = ""
        Dim psArray As Object() = New Object() {printStrArray1, printStrArray2, printStrArray3, printStrArray4, printStrArray5, printStrArray6, printStrArray7}
        Dim phArray As Object() = New Object() {printHinmeiArray1, printHinmeiArray2, printHinmeiArray3, printHinmeiArray4, printHinmeiArray5, printHinmeiArray6, printHinmeiArray7}
        'Dim A4width As Double = 595
        'Dim A4Height As Double = 842
        Dim data_count As Integer = 0

        For k As Integer = 0 To 3
            If psArray(k).count > 0 Then
                'For i As Integer = 0 To psArray(k).Count - 1
                Try
                    For itemNum As Integer = 0 To psArray(k).count - 1
                        Dim wb As IWorkbook = WorkbookFactory.Create(Path.GetDirectoryName(Form1.appPath) & "\config\発注書.xlsx")
                        Dim ws As ISheet = wb.GetSheetAt(wb.GetSheetIndex("Sheet1"))

                        Dim Font As IFont = ws.Workbook.CreateFont
                        'Font.Underline = FontUnderline.SINGLE.ByteValue
                        Font.FontHeightInPoints = 12
                        Font.FontName = "ＭＳ Ｐゴシック"

                        'ws.GetRow(3).GetCell(1).SetCellValue("とんよか卸")
                        ws.GetRow(3).GetCell(1).SetCellValue("卸")
                        ws.GetRow(6).GetCell(3).SetCellValue(psArray(k)(itemNum)(0).ToString)
                        ws.GetRow(7).GetCell(3).SetCellValue(psArray(k)(itemNum)(1).ToString)
                        ws.GetRow(8).GetCell(3).SetCellValue(psArray(k)(itemNum)(2).ToString)
                        ws.GetRow(9).GetCell(3).SetCellValue(psArray(k)(itemNum)(3).ToString)

                        Dim createRow As Integer = 12
                        Dim createRow_ As Integer = 14

                        If k = 0 Then
                            For listNum As Integer = 0 To printStrArray1_list(itemNum).count - 1
                                ws.CreateRow(createRow + listNum).Height = 530
                                ws.GetRow(createRow + listNum).CreateCell(1)
                                Dim ce = ws.GetRow(createRow + listNum).GetCell(1)
                                ce.SetCellValue(printStrArray1_list(itemNum)(listNum).ToString)
                                ce.CellStyle.SetFont(Font)
                                'ce.CellStyle.BorderBottom = BorderStyle.Thin

                                ws.GetRow(createRow + listNum).CreateCell(6)
                                Dim ce2 = ws.GetRow(createRow + listNum).GetCell(6)
                                ce2.SetCellValue(printHinmeiArray1(itemNum)(listNum).ToString)
                                ce2.CellStyle.SetFont(Font)

                                createRow_ += 1
                            Next
                        ElseIf k = 1 Then
                            For listNum As Integer = 0 To printStrArray2_list(itemNum).count - 1
                                ws.CreateRow(createRow + listNum).Height = 530
                                ws.GetRow(createRow + listNum).CreateCell(1)
                                Dim ce = ws.GetRow(createRow + listNum).GetCell(1)
                                ce.SetCellValue(printStrArray2_list(itemNum)(listNum).ToString)
                                ce.CellStyle.SetFont(Font)

                                ws.GetRow(createRow + listNum).CreateCell(6)
                                Dim ce2 = ws.GetRow(createRow + listNum).GetCell(6)
                                ce2.SetCellValue(printHinmeiArray2(itemNum)(listNum).ToString)
                                ce2.CellStyle.SetFont(Font)

                                createRow_ += 1
                            Next
                        ElseIf k = 2 Then
                            For listNum As Integer = 0 To printStrArray3_list(itemNum).count - 1
                                ws.CreateRow(createRow + listNum).Height = 530
                                ws.GetRow(createRow + listNum).CreateCell(1)
                                Dim ce = ws.GetRow(createRow + listNum).GetCell(1)
                                ce.SetCellValue(printStrArray3_list(itemNum)(listNum).ToString)
                                ce.CellStyle.SetFont(Font)

                                ws.GetRow(createRow + listNum).CreateCell(6)
                                Dim ce2 = ws.GetRow(createRow + listNum).GetCell(6)
                                ce2.SetCellValue(printHinmeiArray3(itemNum)(listNum).ToString)
                                ce2.CellStyle.SetFont(Font)

                                createRow_ += 1
                            Next
                        ElseIf k = 3 Then
                            For listNum As Integer = 0 To printStrArray4_list(itemNum).count - 1
                                ws.CreateRow(createRow + listNum).Height = 530
                                ws.GetRow(createRow + listNum).CreateCell(1)
                                Dim ce = ws.GetRow(createRow + listNum).GetCell(1)
                                ce.SetCellValue(printStrArray4_list(itemNum)(listNum).ToString)
                                ce.CellStyle.SetFont(Font)

                                ws.GetRow(createRow + listNum).CreateCell(6)
                                Dim ce2 = ws.GetRow(createRow + listNum).GetCell(6)
                                ce2.SetCellValue(printHinmeiArray4(itemNum)(listNum).ToString)
                                ce2.CellStyle.SetFont(Font)

                                createRow_ += 1
                            Next
                        End If

                        ws.CreateRow(createRow_).Height = 530
                        ws.GetRow(createRow_).CreateCell(6)
                        Dim ce3 = ws.GetRow(createRow_).GetCell(6)
                        ce3.SetCellValue("担当者")
                        ce3.CellStyle.SetFont(Font)

                        ws.GetRow(createRow_).CreateCell(7)
                        Dim ce4 = ws.GetRow(createRow_).GetCell(7)
                        ce4.SetCellValue(Replace(Form1.ログイン名ToolStripMenuItem.Text, "さん", ""))
                        ce4.CellStyle.SetFont(Font)

                        Using wfs = File.Create(dir & "▲発注書" & Format(Now, "_yyyyMMddHHmmss") & createRow_ & loginName & ".xlsx")
                            wb.Write(wfs)
                        End Using
                        ws = Nothing
                        wb = Nothing
                        data_count += 1
                    Next
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                    Exit Sub
                End Try
                'Next
            End If
        Next

        If data_count > 0 Then
            saveList.Add("▲発注書.xlsx")
        End If
    End Sub

    '調べる
    Private Sub KryptonButton1_Click(sender As Object, e As EventArgs) Handles KryptonButton1.Click
        CheckList()
    End Sub
    Private Sub CheckList(Optional mode As Integer = 0)
        If DGV1.RowCount = 0 Then
            MsgBox("データが読み込まれていません", MsgBoxStyle.Exclamation Or MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        If Panel8.Visible = False And Panel3.Visible = False And Panel7.Visible = False Then
            LIST4VIEW("項目チェック", "start")
            Label9.Visible = True

            '初期化
            CheckBox8.Checked = False
            If mode = 0 Then
                ListBox1.Items.Clear()
            End If

            Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)

            'マスタ配送と発送方法が異なる時
            For r As Integer = 0 To DGV1.RowCount - 1
                Dim mes As String = "配送（" & DGV1(TM_ArIndexof(dH1, "マスタ配送"), r).Value & "が" & DGV1(TM_ArIndexof(dH1, "発送方法"), r).Value & "）確認"
                'If dgv1(TM_ArIndexof(dH1, "発送方法"), r).Value = "メール便" Then
                '    mes = "配送（宅配便がメール便）確認"
                'ElseIf dgv1(TM_ArIndexof(dH1, "発送方法"), r).Value = "定形外" Then
                '    mes = "配送（定形外がメール便）確認"
                'Else
                '    mes = "配送（メール便が宅配便）確認"
                'End If
                If DGV1(TM_ArIndexof(dH1, "マスタ配送"), r).Value <> "ヤマト(陸便)" And DGV1(TM_ArIndexof(dH1, "マスタ配送"), r).Value <> "ヤマト(船便)" Then
                    If DGV1(TM_ArIndexof(dH1, "マスタ配送"), r).Value <> DGV1(TM_ArIndexof(dH1, "発送方法"), r).Value Then
                        DGV1(TM_ArIndexof(dH1, "マスタ配送"), r).Style.BackColor = Color.Yellow
                        DGV1(TM_ArIndexof(dH1, "発送方法"), r).Style.BackColor = Color.Yellow
                        Dim mesline As String = TM_ArIndexof(dH1, "発送方法") & "/" & r + 1 & " : " & mes
                        ListBox1.Items.Add(mesline)
                        '---------------
                        Dim dNo As String = DGV1(TM_ArIndexof(dH1, "伝票番号"), r).Value
                        ValidateRowAdd(dNo, mesline)
                        '---------------
                    End If
                End If
            Next

            '複数倉庫エラー
            For r As Integer = 0 To DGV1.RowCount - 1
                If DGV1(TM_ArIndexof(dH1, "発送倉庫"), r).Value = "複数倉庫" Then
                    DGV1(TM_ArIndexof(dH1, "発送倉庫"), r).Style.BackColor = Color.Yellow
                    Dim mesline As String = TM_ArIndexof(dH1, "発送倉庫") & "/" & r + 1 & " : 複数倉庫確認"
                    ListBox1.Items.Add(mesline)
                    '---------------
                    Dim dNo As String = DGV1(TM_ArIndexof(dH1, "伝票番号"), r).Value
                    ValidateRowAdd(dNo, mesline)
                    '---------------
                End If
            Next

            'ヤマトで名古屋の場合エラー
            For r As Integer = 0 To DGV1.RowCount - 1
                If DGV1(TM_ArIndexof(dH1, "発送方法"), r).Value = "ヤマト" And DGV1(TM_ArIndexof(dH1, "発送倉庫"), r).Value = "名古屋" Then
                    DGV1(TM_ArIndexof(dH1, "発送方法"), r).Style.BackColor = Color.Yellow
                    Dim mesline As String = TM_ArIndexof(dH1, "発送倉庫") & "/" & r + 1 & " : ヤマト名古屋確認"
                    ListBox1.Items.Add(mesline)
                    '---------------
                    Dim dNo As String = DGV1(TM_ArIndexof(dH1, "伝票番号"), r).Value
                    ValidateRowAdd(dNo, mesline)
                    '---------------
                End If
            Next

            '使用不可文字チェック
            For r As Integer = 0 To DGV1.RowCount - 1
                '文字数MAX
                Dim max As Integer = 16 * 3
                Select Case DGV1(TM_ArIndexof(dH1, "伝票ソフト"), r).Value
                    Case "e飛伝2"
                        max = 16 * 3
                    Case "BIZlogi"
                        max = 32 * 3
                    Case "メール便", "定型外"
                        max = 25 * 2
                    Case Else
                        max = 16 * 3
                End Select

                '住所の文字数チェック
                Dim tAddr As String = DGV1(TM_ArIndexof(dH1, "発送先住所"), r).Value
                If tAddr.Length > max Then
                    DGV1(TM_ArIndexof(dH1, "発送先住所"), r).Style.BackColor = Color.Yellow
                    Dim mesline As String = TM_ArIndexof(dH1, "発送先住所") & "/" & r + 1 & " : 住所文字数 桁溢れ"
                    ListBox1.Items.Add(mesline)
                    '---------------
                    Dim dNo As String = DGV1(TM_ArIndexof(dH1, "伝票番号"), r).Value
                    ValidateRowAdd(dNo, mesline)
                    '---------------
                End If

                '名称・備考の文字数チェック
                Dim tName1 As String = DGV1(TM_ArIndexof(dH1, "発送先名"), r).Value
                If tName1.Length > 16 Then
                    DGV1(TM_ArIndexof(dH1, "発送先名"), r).Style.BackColor = Color.Yellow
                    Dim mesline As String = TM_ArIndexof(dH1, "発送先名") & "/" & r + 1 & " : 発送先名 桁溢れ"
                    ListBox1.Items.Add(mesline)
                    '---------------
                    Dim dNo As String = DGV1(TM_ArIndexof(dH1, "伝票番号"), r).Value
                    ValidateRowAdd(dNo, mesline)
                    '---------------
                End If

                '電話番号の桁数チェック
                Dim denwa As String = DGV1(TM_ArIndexof(dH1, "発送先電話番号"), r).Value
                If denwa.Length > 14 Then
                    DGV1(TM_ArIndexof(dH1, "発送先電話番号"), r).Style.BackColor = Color.Yellow
                    Dim mesline As String = TM_ArIndexof(dH1, "発送先電話番号") & "/" & r + 1 & " : 電話番号 桁溢れ"
                    ListBox1.Items.Add(mesline)
                    '---------------
                    Dim dNo As String = DGV1(TM_ArIndexof(dH1, "伝票番号"), r).Value
                    ValidateRowAdd(dNo, mesline)
                    '---------------
                End If

                '機種依存など変な文字をチェック
                If Regex.IsMatch(tAddr, "\'|""|\||\,|\?|#|〓") Then
                    DGV1(TM_ArIndexof(dH1, "発送先住所"), r).Style.BackColor = Color.Yellow
                    Dim mesline As String = TM_ArIndexof(dH1, "発送先住所") & "/" & r + 1 & " : 使用不可文字使用"
                    ListBox1.Items.Add(mesline)
                    '---------------
                    Dim dNo As String = DGV1(TM_ArIndexof(dH1, "伝票番号"), r).Value
                    ValidateRowAdd(dNo, mesline)
                    '---------------
                End If

                If Regex.IsMatch(tName1, "\'|""|\||\,|\?|#|〓") Then
                    DGV1(TM_ArIndexof(dH1, "発送先名"), r).Style.BackColor = Color.Yellow
                    Dim mesline As String = TM_ArIndexof(dH1, "発送先名") & "/" & r + 1 & " : 使用不可文字使用"
                    ListBox1.Items.Add(mesline)
                    '---------------
                    Dim dNo As String = DGV1(TM_ArIndexof(dH1, "伝票番号"), r).Value
                    ValidateRowAdd(dNo, mesline)
                    '---------------
                End If

                '支払方法チェック
                'If DGV1.Item(dH1.IndexOf("支払方法"), r).Value = "代金引換" And DGV1.Item(dH1.IndexOf("発送倉庫"), r).Value = "名古屋" Then
                'DGV1.Item(dH1.IndexOf("支払方法"), r).Style.BackColor = Color.Yellow
                'DGV1.Item(dH1.IndexOf("商品マスタ"), r).Style.BackColor = Color.Red

                'Dim mesline As String = TM_ArIndexof(dH1, "支払方法") & "/" & r + 1 & " : 名古屋代金引換不可"
                'ListBox1.Items.Add(mesline)
                ''---------------
                'Dim dNo As String = DGV1(TM_ArIndexof(dH1, "伝票番号"), r).Value
                'ValidateRowAdd(dNo, mesline)
                '---------------
                'End If

            Next

            '佐川エラーチェック
            SagawaCheck()
            SagawaCheck2()

            '番地未入力
            For r As Integer = 0 To DGV1.RowCount - 1
                Dim tAddr As String = DGV1(TM_ArIndexof(dH1, "発送先住所"), r).Value
                Dim addressCheck As Boolean = False
                Dim work As String
                Dim i As Integer
                For i = 0 To tAddr.Length - 1 '住所の文字数分繰り返す
                    work = tAddr.Substring(i, 1)
                    If IsNumeric(work) Then
                        addressCheck = True
                        Exit For
                    End If
                Next i
                If addressCheck = False Then
                    DGV1(TM_ArIndexof(dH1, "発送先住所"), r).Style.BackColor = Color.Yellow
                    Dim mesline As String = TM_ArIndexof(dH1, "発送先住所") & "/" & r + 1 & " : 番地未入力"
                    ListBox1.Items.Add(mesline)
                    '---------------
                    Dim dNo As String = DGV1(TM_ArIndexof(dH1, "伝票番号"), r).Value
                    ValidateRowAdd(dNo, mesline)
                    '---------------
                End If
            Next

            '日付指定チェック
            If dH1.Contains("配送日") = True Then
                For r As Integer = 0 To DGV1.RowCount - 1
                    Dim listDate As String = DGV1.Item(TM_ArIndexof(dH1, "配送日"), r).Value
                    If listDate <> "配送日" And listDate <> "" Then
                        Dim nowDate As String = Format(Now, "yyyy/MM/dd 00:00:00")
                        listDate = listDate.Substring(0, 4) & "/" & listDate.Substring(4, 2) & "/" & listDate.Substring(6, 2)
                        If DateDiff(DateInterval.Day, CDate(nowDate), CDate(listDate & " 00:00:00")) <= 0 Then
                            DGV1.Item(TM_ArIndexof(dH1, "配送日"), r).Style.BackColor = Color.Orange
                            DGV1(TM_ArIndexof(dH1, "配送日"), r).Style.BackColor = Color.Yellow
                            Dim mesline As String = TM_ArIndexof(dH1, "配送日") & "/" & r + 1 & " : 日付指定が現在日より前"
                            ListBox1.Items.Add(mesline)
                            '---------------
                            Dim dNo As String = DGV1(TM_ArIndexof(dH1, "伝票番号"), r).Value
                            ValidateRowAdd(dNo, mesline)
                            '---------------
                        End If
                    End If
                Next
            End If

            'その他黄色（エラー）を残す
            For r As Integer = 0 To DGV1.RowCount - 1
                For c As Integer = 0 To DGV1.ColumnCount - 1
                    If DGV1.Item(c, r).Style.BackColor = Color.Yellow Then
                        DGV1(TM_ArIndexof(dH1, "配送日"), r).Style.BackColor = Color.Yellow
                    End If
                Next
            Next

            For i As Integer = 0 To ListBox1.Items.Count - 1
                If InStr(ListBox1.Items(i), "中継") = 0 Then
                    TextBox48.Text = "修正"
                    TextBox48.BackColor = Color.Yellow
                End If
            Next

            Label9.Visible = False
            LIST4VIEW("項目チェック", "end")
        Else
            MsgBox("データがチェックできません", MsgBoxStyle.Exclamation Or MsgBoxStyle.SystemModal)
            Exit Sub
        End If
    End Sub

    Private Sub SagawaCheck()
        LIST4VIEW("佐川チェック", "start")
        Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)

        TM_DB_CONNECT(CnAccdb)

        For r As Integer = 0 To DGV1.Rows.Count - 1
            If DGV1(TM_ArIndexof(dH1, "発送方法"), r).Value = "宅配便" Then
                Dim SakiYubin As String = DGV1(TM_ArIndexof(dH1, "発送先郵便番号"), r).Value

                Dim whereCheck As String = " WHERE ([YUBIN] = '" & SakiYubin & "')"
                Dim table As DataTable = TM_DB_GET_SQL("TOWN_T", whereCheck)

                If table.Rows.Count <> 0 Then
                    Dim DAIBIKI_KBN As String = table.Rows(0)("DAIBIKI_KBN").ToString
                    Dim TIME_KBN As String = table.Rows(0)("TIME_KBN").ToString
                    Dim CHAKU_KBN As String = table.Rows(0)("CHAKU_KBN").ToString
                    Dim TRANSIT_KBN As String = table.Rows(0)("TRANSIT_KBN").ToString
                    Dim OGB_KBN As String = table.Rows(0)("OGB_KBN").ToString

                    If CStr(DGV1(TM_ArIndexof(dH1, "配送時間帯"), r).Value) <> "" Then
                        If TIME_KBN <> "0" Then
                            DGV1.Item(TM_ArIndexof(dH1, "配送時間帯"), r).Style.BackColor = Color.Yellow
                            Dim mesline As String = TM_ArIndexof(dH1, "配送時間帯") & "/" & r + 1 & " : 時間指定-不可地域"
                            ListBox1.Items.Add(mesline)
                            '---------------
                            Dim dNo As String = DGV1(TM_ArIndexof(dH1, "伝票番号"), r).Value
                            ValidateRowAdd(dNo, mesline)
                            '---------------
                        End If
                    End If
                    If TRANSIT_KBN <> "0" Then
                        DGV1.Item(TM_ArIndexof(dH1, "発送先郵便番号"), r).Style.BackColor = Color.LemonChiffon
                        DGV1.Item(TM_ArIndexof(dH1, "発送先住所"), r).Style.BackColor = Color.LemonChiffon
                        Dim mesline As String = TM_ArIndexof(dH1, "発送先郵便番号") & "/" & r + 1 & " : 中継料金-追加地域"
                        ListBox1.Items.Add(mesline)
                        '---------------
                        Dim dNo As String = DGV1(TM_ArIndexof(dH1, "伝票番号"), r).Value
                        ValidateRowAdd(dNo, mesline)
                        '---------------
                    End If
                    If IsNumeric(DGV1(TM_ArIndexof(dH1, "合計金額"), r).Value) Then
                        If DGV1(TM_ArIndexof(dH1, "合計金額"), r).Value > 0 And DAIBIKI_KBN <> "0" Then
                            DGV1.Item(TM_ArIndexof(dH1, "合計金額"), r).Style.BackColor = Color.Yellow
                            Dim mesline As String = TM_ArIndexof(dH1, "合計金額") & "/" & r + 1 & " : 代金引換-不可地域"
                            ListBox1.Items.Add(mesline)
                            '---------------
                            Dim dNo As String = DGV1(TM_ArIndexof(dH1, "伝票番号"), r).Value
                            ValidateRowAdd(dNo, mesline)
                            '---------------
                        End If
                    End If
                Else
                    '見つからなければスルー
                End If

                table.Clear()

                ''文字数MAX
                'Dim max As Integer = 16 * 3
                'Select Case dgv1(TM_ArIndexof(dH1, "伝票ソフト"), r).Value
                '    Case "e飛伝2"
                '        max = 16 * 3
                '    Case "BIZlogi"
                '        max = 32 * 3
                '    Case "メール便", "定型外"
                '        max = 25 * 2
                '    Case Else
                '        max = 16 * 3
                'End Select

                ''住所の文字数チェック
                'Dim tAddr As String = dgv1(TM_ArIndexof(dH1, "発送先住所"), r).Value
                'If tAddr.Length > max Then
                '    DataGridView1(TM_ArIndexof(dH1, "発送先住所"), r).Style.BackColor = Color.Yellow
                '    Dim mesline As String = TM_ArIndexof(dH1, "発送先住所") & "/" & r + 1 & " : 住所文字数 桁溢れ"
                '    ListBox1.Items.Add(mesline)
                '    '---------------
                '    Dim dNo As String = dgv1(TM_ArIndexof(dH1, "伝票番号"), r).Value
                '    ValidateRowAdd(dNo, mesline)
                '    '---------------
                'End If

                ''名称・備考の文字数チェック
                'Dim tName1 As String = DataGridView1(TM_ArIndexof(dH1, "発送先名"), r).Value
                'If tName1.Length > 16 Then
                '    DataGridView1(TM_ArIndexof(dH1, "発送先名"), r).Style.BackColor = Color.Yellow
                '    Dim mesline As String = TM_ArIndexof(dH1, "発送先名") & "/" & r + 1 & " : 発送先名 桁溢れ"
                '    ListBox1.Items.Add(mesline)
                '    '---------------
                '    Dim dNo As String = dgv1(TM_ArIndexof(dH1, "伝票番号"), r).Value
                '    ValidateRowAdd(dNo, mesline)
                '    '---------------
                'End If

                ''機種依存など変な文字をチェック
                'If Regex.IsMatch(tAddr, "\'|""|\||\,|\?|#") Then
                '    DataGridView1(TM_ArIndexof(dH1, "発送先住所"), r).Style.BackColor = Color.Yellow
                '    Dim mesline As String = TM_ArIndexof(dH1, "発送先住所") & "/" & r + 1 & " : 使用不可文字使用"
                '    ListBox1.Items.Add(mesline)
                '    '---------------
                '    Dim dNo As String = dgv1(TM_ArIndexof(dH1, "伝票番号"), r).Value
                '    ValidateRowAdd(dNo, mesline)
                '    '---------------
                'End If

                'If Regex.IsMatch(tName1, "\'|""|\||\,|\?|#") Then
                '    DataGridView1(TM_ArIndexof(dH1, "発送先名"), r).Style.BackColor = Color.Yellow
                '    Dim mesline As String = TM_ArIndexof(dH1, "発送先名") & "/" & r + 1 & " : 使用不可文字使用"
                '    ListBox1.Items.Add(mesline)
                '    '---------------
                '    Dim dNo As String = dgv1(TM_ArIndexof(dH1, "伝票番号"), r).Value
                '    ValidateRowAdd(dNo, mesline)
                '    '---------------
                'End If

                If r Mod 50 = 0 Then
                    LIST4VIEW("チェック中 ～" & r, "...")
                End If

            End If
        Next

        TM_DB_DISCONNECT()
    End Sub

    Private Sub SagawaCheck2()
        LIST4VIEW("住所チェック", "start")
        Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)

        TM_DB_CONNECT(CnAccdb)

        For r As Integer = 0 To DGV1.Rows.Count - 1
            LIST4VIEW(r, "startSagawaCheck2")
            If DGV1(TM_ArIndexof(dH1, "発送方法"), r).Value = "宅配便" Then
                Dim SakiYubin As String = DGV1(TM_ArIndexof(dH1, "発送先郵便番号"), r).Value

                Dim whereCheck As String = " WHERE ([YUBIN] = '" & SakiYubin & "')"
                Dim table1 As DataTable = TM_DB_GET_SQL("TOWN_T", whereCheck)
                '郵便番号同じ住所がここで複数出てくる
                If table1.Rows.Count = 0 Then
                    Continue For
                End If

                Dim check_jusyo As String = ""
                Dim check_jusyo2 As String = ""
                For i As Integer = 0 To table1.Rows.Count - 1
                    Dim JIS_5 As String = table1.Rows(i)("JIS_5").ToString

                    whereCheck = " WHERE ([JIS_5] = '" & JIS_5 & "')"
                    Dim table2 As DataTable = TM_DB_GET_SQL("CITY_T", whereCheck)
                    If table2.Rows.Count = 0 Then
                        Continue For
                    End If
                    Dim SHI_NAME As String = table2.Rows(0)("SHI_NAME").ToString
                    Dim JIS_2 As String = table2.Rows(0)("JIS_2").ToString
                    table2.Clear()

                    whereCheck = " WHERE ([JIS_2] = '" & JIS_2 & "')"
                    Dim table3 As DataTable = TM_DB_GET_SQL("KEN_T", whereCheck)
                    If table3.Rows.Count = 0 Then
                        Continue For
                    End If
                    Dim TODOFUKEN As String = table3.Rows(0)("TODOFUKEN").ToString
                    table3.Clear()

                    TODOFUKEN = Regex.Replace(TODOFUKEN, "　| ", "")
                    SHI_NAME = Regex.Replace(SHI_NAME, "　| ", "")
                    SHI_NAME = Replace(SHI_NAME, "ケ", "ヶ")
                    If check_jusyo <> "" Then
                        check_jusyo &= "|"
                    End If
                    check_jusyo &= "^" & TODOFUKEN & SHI_NAME & "|^" & TODOFUKEN & " " & SHI_NAME & "|^" & TODOFUKEN & "　" & SHI_NAME
                    check_jusyo2 = TODOFUKEN & SHI_NAME
                Next
                table1.Clear()

                check_jusyo = "(" & check_jusyo & ")"
                Dim SakiJusyo As String = DGV1(TM_ArIndexof(dH1, "発送先住所"), r).Value
                SakiJusyo = Replace(SakiJusyo, "ケ", "ヶ")
                SakiJusyo = Regex.Replace(SakiJusyo, " |　", "")
                If Not Regex.IsMatch(SakiJusyo, check_jusyo) Then
                    DGV1.Item(TM_ArIndexof(dH1, "発送先住所"), r).Style.BackColor = Color.Yellow
                    Dim mesline As String = TM_ArIndexof(dH1, "発送先住所") & "/" & r + 1 & " : 住所(" & check_jusyo2 & ")"
                    ListBox1.Items.Add(mesline)
                    '---------------
                    Dim dNo As String = DGV1(TM_ArIndexof(dH1, "伝票番号"), r).Value
                    ValidateRowAdd(dNo, mesline)
                    '---------------
                End If
            Else
                '見つからなければスルー
            End If
        Next

        TM_DB_DISCONNECT()
        LIST4VIEW("住所チェック", "end")
    End Sub

    Private Sub ListBox1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.DoubleClick
        If ListBox1.Items.Count > 0 And ListBox1.SelectedItem <> Nothing Then
            Dim res1 As String() = Split(ListBox1.SelectedItem, " : ")
            Dim res2 As String() = Split(res1(0), "/")
            Dim c As Integer = res2(0)
            Dim r As Integer = res2(1) - 1
            If DGV1.Rows(r).Visible Then
                TabControl2.SelectedIndex = 0
                DGV1.CurrentCell = DGV1(c, r)
            Else
                MsgBox("行がありません。行削除などしていませんか？")
            End If
        End If
    End Sub

    'ヘルプ表示
    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        Try
            If ListBox1.Items.Count > 0 Then
                For i As Integer = 0 To helpArray.Count - 1
                    Dim help As String() = Split(helpArray(i), ",")
                    If Regex.IsMatch(ListBox1.SelectedItems(0), help(0)) Then
                        ToolStripLabel2.Text = help(1)
                        Exit For
                    End If
                Next
            Else
                ToolStripLabel2.Text = "ヘルプ"
            End If
        Catch ex As Exception
            ToolStripLabel2.Text = "ヘルプ"
        End Try
    End Sub

    '抽出
    Private Sub CheckBox8_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox8.CheckedChanged
        If DGV1.RowCount > 0 Then
            LB1_visible(CheckBox8.Checked, CheckBox14.Checked)
        End If
    End Sub

    Private Sub LB1_visible(visiFlag As Boolean, chukeiFlag As Boolean)
        If ListBox1.Items.Count = 0 Or ListBox1.SelectedItems.Count = 0 Then
            Exit Sub
        End If

        Dim lbArray As New ArrayList
        If visiFlag Then
            For i As Integer = 0 To ListBox1.Items.Count - 1
                If chukeiFlag And InStr(ListBox1.Items(i), "中継") > 0 Then
                    Continue For
                End If

                Dim res1 As String() = Split(ListBox1.Items(i), " : ")
                Dim flag As Boolean = True

                If CheckBox9.Checked Then
                    If res1(1) = Split(ListBox1.SelectedItem, " : ")(1) Then
                        flag = True
                    Else
                        flag = False
                    End If
                Else
                    flag = True
                End If

                If flag Then
                    Dim res2 As String() = Split(res1(0), "/")
                    Dim r As Integer = res2(1) - 1
                    If Not lbArray.Contains(r) Then
                        lbArray.Add(r)
                    End If
                End If
            Next
            For r As Integer = 0 To DGV1.RowCount - 1
                If lbArray.Contains(r) Then
                    DGV1.Rows(r).Visible = True
                Else
                    DGV1.Rows(r).Visible = False
                End If
            Next
        Else
            For r As Integer = 0 To DGV1.RowCount - 1
                DGV1.Rows(r).Visible = True
            Next
        End If
    End Sub

    Private Sub ListBox3_DoubleClick(sender As Object, e As EventArgs) Handles ListBox3.DoubleClick
        If ListBox3.Items.Count > 0 And ListBox3.SelectedItem <> Nothing And DGV1.Rows.Count > 0 Then
            Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
            For r As Integer = 0 To DGV1.RowCount - 1
                If ListBox3.SelectedItem = DGV1.Item(dH1.IndexOf("伝票番号"), r).Value Then
                    DGV1.CurrentCell = DGV1(dH1.IndexOf("伝票番号"), r)
                    Exit For
                End If
            Next
        End If
    End Sub

    '太宰府或いは名古屋
    Private Function checkSouko_DaOrNa(tag_decide As ArrayList, code As String, count As Integer, haisouSaki As String) As Boolean
        Dim bl As Boolean = Nothing 'true: 名古屋 false: 太宰府

        If code = "ny263-51" Then
            'If count Mod 20 = 0 Then
            '    If count = 20 Or count = 40 Then '20，40個
            '        Dim isoosakaizyo As Boolean = checkHaisosaki_DaOrNa(haisouSaki)
            '        If checkHaisosaki_DaOrNa(haisouSaki) Then
            '            bl = True 'true: 名古屋
            '        Else
            '            bl = False 'false: 太宰府
            '        End If
            '    Else
            '        bl = True '60以上(60倍数)
            '    End If
            'Else
            '    bl = False
            'End If
            bl = False
        ElseIf code = "ny264" Then
            If count Mod 20 = 0 Then
                If count = 20 Or count = 40 Then '20，40個
                    Dim isoosakaizyo As Boolean = checkHaisosaki_DaOrNa(haisouSaki)
                    If checkHaisosaki_DaOrNa(haisouSaki) Then
                        bl = True 'true: 名古屋
                    Else
                        bl = False 'false: 太宰府
                    End If
                Else
                    bl = True '60以上(60倍数)
                End If
            Else
                bl = False
            End If
        ElseIf code = "ny261" Then
            If count <= 2 Then '1000枚和2000枚
                If checkHaisosaki_DaOrNa(haisouSaki) Then
                    bl = True
                Else
                    bl = False
                End If
                'If count = 1 Then
                '        'If checkHaisosaki_DaOrNa(haisouSaki) Then
                '        '    bl = True 'true: 名古屋
                '        'Else
                '        bl = False 'false: 太宰府
                '        'End If
                '    Else
                '        bl = True 'true: 名古屋
                '    End If
            Else
                '    If count Mod 2 = 1 Then
                '    bl = False 'false: 太宰府
                'Else
                '    bl = True 'true: 名古屋
                'End If
                bl = True
            End If
        ElseIf code = "ny261ye" Then

            bl = False 'false: 太宰府

        ElseIf code = "ny264" Then
            If count Mod 20 = 0 Then
                If count = 20 Or count = 40 Then '20，40個
                    Dim isoosakaizyo As Boolean = checkHaisosaki_DaOrNa(haisouSaki)
                    If checkHaisosaki_DaOrNa(haisouSaki) Then
                        bl = True 'true: 名古屋
                    Else
                        bl = False 'false: 太宰府
                    End If
                Else
                    bl = True '40以上(40倍数)
                End If
            Else
                bl = False
            End If
        ElseIf code = "ny264-100" Then
            If count Mod 10 = 0 Then
                If count = 10 Or count = 20 Then '10，20個
                    bl = False 'false: 太宰府
                Else
                    bl = True 'true: 名古屋
                End If
            Else
                bl = False
            End If
        ElseIf code = "ny264-200" Then
            If count Mod 5 = 0 Then
                If count = 5 Or count = 10 Then '10，20個
                    bl = False 'false: 太宰府
                Else
                    bl = True 'true: 名古屋
                End If
            Else
                bl = False
            End If
        ElseIf code = "ny264-500" Then
            If count Mod 2 = 0 Then
                If count = 2 Or count = 4 Then '10，20個
                    bl = False 'false: 太宰府
                Else
                    bl = True 'true: 名古屋
                End If
            Else
                bl = False
            End If
        ElseIf code = "ny264-3000a" Then
            Dim isoosakaizyo As Boolean = checkHaisosaki_DaOrNa(haisouSaki)
            If checkHaisosaki_DaOrNa(haisouSaki) Then
                bl = True 'true: 名古屋
            Else
                bl = False 'false: 太宰府
            End If
            bl = True 'true: 名古屋
            'ElseIf code = "pa084-5" Then
            '    'If count >= 3 Then
            '    '    bl = True
            '    'Else
            '    '    bl = False
            '    'End If
            '    bl = True
            'ElseIf code = "pa084-ho" Then
            '    bl = True
            'ElseIf code = "pa084-7" Then
            '    If count >= 2 Then
            '        bl = True
            '    Else
            '        bl = False
            '    End If
            'ElseIf code = "od433-co" Then
            '    If checkHaisosaki_DaOrNa(haisouSaki) Then
            '        bl = True 'true: 名古屋
            '    Else
            '        bl = False 'false: 太宰府
            '    End If
        ElseIf code = "od433-wa" Then
            'If checkHaisosaki_DaOrNa(haisouSaki) Then
            '    bl = True 'true: 名古屋
            'Else
            '    bl = False 'false: 太宰府
            'End If

            bl = True 'true: 名古屋
        ElseIf code = "od492" Then
            If checkHaisosaki_DaOrNa(haisouSaki) Then
                bl = True 'true: 名古屋
            Else
                bl = False 'false: 太宰府
            End If


        End If

        If bl Then
            tag_decide.Add(nagoya_str)
        Else
            tag_decide.Add(dazaifu_str)
        End If

        Return bl
    End Function

    Private Function checkHaisosaki_DaOrNa(haisouSaki As String) As Boolean 'true: 大阪以上 'false: 大阪以降 
        Dim bl As Boolean = False
        haisouSaki = haisouSaki.Substring(0, 5)
        For r As Integer = 0 To checkaddress_oosakizyou.Count - 1
            If Regex.IsMatch(haisouSaki, checkaddress_oosakizyou(r)) Then
                bl = True
                Exit For
            End If
        Next
        Return bl
    End Function

    '============================================
    '--------------------------------------------
    'メイン中央
    '--------------------------------------------
    '============================================
    Dim selChangeFlag As Boolean = True
    Dim BeforeColor As Color() = New Color() {Color.Empty, Color.Empty, Color.Empty, Color.Empty, Color.Empty}
    Private Sub DataGridView_SelectionChanged(sender As DataGridView, e As EventArgs) Handles _
            DGV1.SelectionChanged, DGV7.SelectionChanged, DGV8.SelectionChanged,
            DGV9.SelectionChanged, DGV13.SelectionChanged
        If selChangeFlag = False Then
            Exit Sub
        ElseIf sender.Rows.Count = 0 Then
            Exit Sub
        End If

        Dim dHSender As ArrayList = TM_HEADER_GET(sender)

        Dim num As Integer = 0
        Select Case True
            Case sender Is DGV1
                num = 0
            Case sender Is DGV7
                num = 1
            Case sender Is DGV8
                num = 2
            Case sender Is DGV9
                num = 3
            Case sender Is DGV13
                num = 4
            Case sender Is TMSDGV
                num = 5
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
        Dim dH2 As ArrayList = TM_HEADER_GET(DGV2)
        Dim dH3 As ArrayList = TM_HEADER_GET(DGV3)
        'Dim dH6 As ArrayList = TM_HEADER_GET(dgv6)

        If sender.RowCount > 0 Then
            If sender.SelectedCells.Count > 0 And DGV3.RowCount > 0 Then
                '選択セルの中身を表示する
                TextBox12.Text = sender.SelectedCells(0).Value
                If TextBox12.Text <> "" Then
                    TextBox12.Text = Regex.Replace(TextBox12.Text, "\n|\r|\n\r", "")
                End If
                Dim dNum As String = sender.Item(dHSender.IndexOf(koumoku("denpyoNo")(num)), sender.SelectedCells(0).RowIndex).Value
                If IsNumeric(dNum) Then '伝票番号
                    Dim tenpo As String = ""
                    For r As Integer = 0 To DGV3.RowCount - 1
                        If DGV3.Item(dH3.IndexOf("伝票番号"), r).Value = dNum Then
                            DGV3.Item(dH3.IndexOf("伝票番号"), r).Selected = True
                            DGV3.CurrentCell = DGV3(dH3.IndexOf("伝票番号"), r)
                            '店舗
                            If tenpo = "" Then
                                tenpo = DGV3.Item(dH3.IndexOf("店舗"), r).Value
                                ToolStripLabel3.Text = "店舗：" & tenpo
                            End If
                        End If
                    Next

                    'Dim header As String() = New String() {"商品コード", "商品名", "商品分類タグ", "代表商品コード", "ship-weight", "ロケーション", "梱包サイズ", "特殊", "sw2"}
                    '商品の情報を検索表示する
                    DGV2.Rows.Clear()
                    Dim codeArray As String() = Split(sender.Item(dHSender.IndexOf(koumoku("masterCode")(num)), sender.SelectedCells(0).RowIndex).Value, "、")
                    Dim chumonsu As Integer = 0
                    If codeArray(0) <> "" Then
                        For i As Integer = 0 To codeArray.Length - 1
                            Dim code As String() = Split(codeArray(i), "*")
                            chumonsu += CInt(code(1))
                            Dim codeA As New ArrayList(code)    'arraylistに移行して追加し、配列に戻す
                            If dgv6CodeArray.Contains(code(0).ToLower) Then
                                codeA.Add(DGV6.Item(dH6.IndexOf("ship-weight"), dgv6CodeArray.IndexOf(code(0).ToLower)).Value)
                                codeA.Add(DGV6.Item(dH6.IndexOf("商品分類タグ"), dgv6CodeArray.IndexOf(code(0).ToLower)).Value)
                                codeA.Add(DGV6.Item(dH6.IndexOf("商品名"), dgv6CodeArray.IndexOf(code(0).ToLower)).Value)
                                codeA.Add(DGV6.Item(dH6.IndexOf("ロケーション"), dgv6CodeArray.IndexOf(code(0).ToLower)).Value)
                                codeA.Add(DGV6.Item(dH6.IndexOf("ロケーション(名古屋)"), dgv6CodeArray.IndexOf(code(0).ToLower)).Value)
                                codeA.Add(DGV6.Item(dH6.IndexOf("梱包サイズ"), dgv6CodeArray.IndexOf(code(0).ToLower)).Value)
                                codeA.Add(DGV6.Item(dH6.IndexOf("特殊"), dgv6CodeArray.IndexOf(code(0).ToLower)).Value)
                                codeA.Add(DGV6.Item(dH6.IndexOf("sw2"), dgv6CodeArray.IndexOf(code(0).ToLower)).Value)
                            Else

                            End If
                            code = CType(codeA.ToArray(Type.GetType("System.String")), String())    'arraylistを配列にする
                            DGV2.Rows.Add(code)
                        Next
                        TabPage10.Text = "商品コード(" & chumonsu & ")"
                    End If
                End If

                '++++++++++++++++++++++++++++++++++++++++++++++
                '伝票操作画面に表示する
                DenpyoSousa(sender, num, sender.SelectedCells(0).RowIndex)
                '++++++++++++++++++++++++++++++++++++++++++++++

                'タブの自動表示変更
                Dim headerStr As String = sender.Columns(sender.SelectedCells(0).ColumnIndex).HeaderText
                Select Case True
                    Case Regex.IsMatch(headerStr, "郵便|住所|発送先名|電話")
                        TabControl3.SelectedIndex = 0
                        TabControl4.SelectedIndex = 1
                    Case Regex.IsMatch(headerStr, "営業所")
                        TabControl3.SelectedIndex = 1
                    Case Regex.IsMatch(headerStr, "商品コード")
                        '動かさない
                    Case Else
                        TabControl3.SelectedIndex = 0
                        TabControl4.SelectedIndex = 0
                End Select
            End If
        End If


    End Sub

    'ダブルクリックで元データに飛ぶ
    Private Sub DataGridView_CellDoubleClick(sender As DataGridView, e As DataGridViewCellEventArgs) Handles _
            DGV7.CellDoubleClick, DGV8.CellDoubleClick, DGV9.CellDoubleClick, DGV13.CellDoubleClick, TMSDGV.CellDoubleClick
        Dim num As Integer = 0
        Select Case True
            Case sender Is DGV7
                num = 1
            Case sender Is DGV8
                num = 2
            Case sender Is DGV9
                num = 3
            Case sender Is DGV13
                num = 4
            Case sender Is TMSDGV
                num = 5
        End Select

        Dim dH As ArrayList = TM_HEADER_GET(sender)
        Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
        Dim selRow As Integer = sender.SelectedCells(0).RowIndex
        Dim dNo As String = sender.Item(dH.IndexOf(koumoku("denpyoNo")(num)), selRow).Value
        For r As Integer = 0 To DGV1.RowCount - 1
            If DGV1.Item(dH1.IndexOf("伝票番号"), r).Value = dNo Then
                TabControl2.SelectedIndex = 0
                DGV1.CurrentCell = DGV1(dH1.IndexOf("伝票番号"), r)
                Exit For
            End If
        Next
    End Sub

    Private Sub 処理後データに移動ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 処理後データに移動ToolStripMenuItem.Click
        If InStr(TabControl2.SelectedTab.Text, "元データ") = 0 Then
            Exit Sub
        ElseIf DGV1.RowCount = 0 Or DGV1.SelectedCells.Count = 0 Then
            Exit Sub
        End If

        Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
        Dim dNo As String = DGV1.Item(dH1.IndexOf("伝票番号"), DGV1.SelectedCells(0).RowIndex).Value

        Dim dgvSel As DataGridView() = {DGV7, DGV8, DGV9, DGV13}
        Do
            For i As Integer = 0 To dgvSel.Length - 1
                For r As Integer = 0 To dgvSel(i).RowCount - 1
                    For c As Integer = 0 To dgvSel(i).ColumnCount - 1
                        If dgvSel(i).Item(c, r).Value = dNo Then
                            TabControl2.SelectedIndex = i + 1
                            dgvSel(i).CurrentCell = dgvSel(i)(c, r)
                            Exit Do
                        End If
                    Next
                Next
            Next

            MsgBox("データが見つかりません", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
            Exit Do
        Loop
    End Sub

    'データの出荷個数を調べる
    Private Sub DataGridView_KosuCount(sender As DataGridView)
        '初期化
        'Dim lkArray As LinkLabel() = {LinkLabel1, LinkLabel2, LinkLabel3, LinkLabel4, LinkLabel5, LinkLabel6, LinkLabel15, LinkLabel16, LinkLabel17, LinkLabel18, LinkLabel19, LinkLabel20, LinkLabel21, LinkLabel22}
        Dim lkArray As LinkLabel() = {LinkLabel1, LinkLabel2, LinkLabel3, LinkLabel4, LinkLabel5, LinkLabel6, LinkLabel15, LinkLabel16, LinkLabel17, LinkLabel18, LinkLabel25, LinkLabel26, LinkLabel23, LinkLabel24, LinkLabel20, LinkLabel19, LinkLabel29, LinkLabel30}

        Dim dHSender As ArrayList = TM_HEADER_GET(sender)
        Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
        Dim dH3 As ArrayList = TM_HEADER_GET(DGV3)

        'Dim kosu As Integer() = New Integer() {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}
        '名古屋は10から ヤマト(太)是第14 ヤマト(井)是第15      ゆう2(太路)是 18  ゆう2(太船)是19   ゆう2(井路)是 20  ゆう2(井船)是21    (22 23  24 TMS)
        Dim kosu As Integer() = New Integer() {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}
        'Dim kosuKouku As Integer() = New Integer() {0, 0, 0, 0, 0}
        '名古屋は9から
        Dim kosuKouku As Integer() = New Integer() {0, 0, 0, 0, 0, 0, 0}
        'Dim tenpoList As String() = New String() {"", "", "", "", "", "", "", "", ""}        '出力確認票用に店舗を入れる
        '名古屋は9から ヤマト是第13 ヤマト(井)是第14      ゆう2太宰府 是第17 ゆう2(井)是第18     index= 18  19  20 TMS
        Dim tenpoList As String() = New String() {"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""}        '出力確認票用に店舗を入れる
        'Dim tenpoListKouku As String() = New String() {"", "", "", "", ""}
        '名古屋は5から
        Dim tenpoListKouku As String() = New String() {"", "", "", "", "", "", ""}
        Dim KoukuMeisai As String = ""
        For i As Integer = 1 To koumoku("koguchi").Length - 1
            If dHSender.Contains(koumoku("koguchi")(i)) Then
                For r As Integer = 0 To sender.RowCount - 1
                    '店舗を取得
                    Dim dNo As String = sender.Item(dHSender.IndexOf(koumoku("denpyoNo")(i)), r).Value
                    Dim tenpo As String = ""
                    For k As Integer = 0 To DGV3.RowCount - 1
                        If dNo = DGV3.Item(dH3.IndexOf("伝票番号"), k).Value Then
                            tenpo = DGV3.Item(dH3.IndexOf("店舗"), k).Value
                            Exit For
                        End If
                    Next

                    Dim KoukuFlag As Boolean = False
                    If sender Is DGV7 Then
                        If sender.Item(dHSender.IndexOf(koumoku("binsyu")(i)), r).Value = "003" Then
                            Select Case sender.Item(dHSender.IndexOf(koumoku("syori2")(i)), r).Value
                                Case HS1.Text
                                    kosuKouku(0) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                                Case HS2.Text
                                    kosuKouku(1) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                                Case HS3.Text
                                    kosuKouku(2) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                                Case HS4.Text
                                    kosuKouku(5) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                            End Select
                            KoukuFlag = True
                        Else
                            Select Case sender.Item(dHSender.IndexOf(koumoku("syori2")(i)), r).Value
                                Case HS1.Text
                                    kosu(0) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                                Case HS2.Text
                                    kosu(1) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                                Case HS3.Text
                                    kosu(2) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                                Case HS4.Text
                                    kosu(10) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                            End Select
                        End If

                        '-----
                        If KoukuFlag = False Then
                            Dim tl As String = ""
                            Select Case sender.Item(dHSender.IndexOf(koumoku("syori2")(i)), r).Value
                                Case HS1.Text
                                    tl = tenpoList(0)
                                Case HS2.Text
                                    tl = tenpoList(1)
                                Case HS3.Text
                                    tl = tenpoList(2)
                                Case HS4.Text
                                    tl = tenpoList(9)
                            End Select
                            If InStr(tl, tenpo) = 0 Then
                                If tl = "" Then
                                    tl = tenpo
                                Else
                                    tl &= "、" & tenpo
                                End If
                            End If
                            Select Case sender.Item(dHSender.IndexOf(koumoku("syori2")(i)), r).Value
                                Case HS1.Text
                                    tenpoList(0) = tl
                                Case HS2.Text
                                    tenpoList(1) = tl
                                Case HS3.Text
                                    tenpoList(2) = tl
                                Case HS4.Text
                                    tenpoList(9) = tl
                            End Select
                        Else
                            KoukuMeisai &= dNo
                            KoukuMeisai &= "," & tenpo & "," & "航空便"
                            KoukuMeisai &= "," & sender.Item(dHSender.IndexOf(koumoku("sakiAddr1")(i)), r).Value
                            KoukuMeisai &= "," & sender.Item(dHSender.IndexOf(koumoku("hinmei")(i)), r).Value
                            KoukuMeisai &= "," & sender.Item(dHSender.IndexOf(koumoku("sakiname")(i)), r).Value
                            KoukuMeisai &= "," & sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value
                            KoukuMeisai &= "," & sender.Item(dHSender.IndexOf(koumoku("masterCode")(i)), r).Value
                            KoukuMeisai &= "," & DGV1.Item(dH1.IndexOf("sw"), dgv1dNoArray.IndexOf(dNo)).Value
                            KoukuMeisai &= "," & DGV1.Item(dH1.IndexOf("サイズ"), dgv1dNoArray.IndexOf(dNo)).Value
                            KoukuMeisai &= "," & DGV1.Item(dH1.IndexOf("発送倉庫"), dgv1dNoArray.IndexOf(dNo)).Value
                            KoukuMeisai &= vbNewLine

                            Dim tl As String = ""
                            Select Case sender.Item(dHSender.IndexOf(koumoku("syori2")(i)), r).Value
                                Case HS1.Text
                                    tl = tenpoListKouku(0)
                                Case HS2.Text
                                    tl = tenpoListKouku(1)
                                Case HS3.Text
                                    tl = tenpoListKouku(2)
                                Case HS4.Text
                                    tl = tenpoListKouku(5)
                            End Select
                            If InStr(tl, tenpo) = 0 Then
                                If tl = "" Then
                                    tl = tenpo
                                Else
                                    tl &= "、" & tenpo
                                End If
                            End If
                            Select Case sender.Item(dHSender.IndexOf(koumoku("syori2")(i)), r).Value
                                Case HS1.Text
                                    tenpoListKouku(0) = tl
                                Case HS2.Text
                                    tenpoListKouku(1) = tl
                                Case HS3.Text
                                    tenpoListKouku(2) = tl
                                Case HS4.Text
                                    tenpoListKouku(5) = tl
                            End Select
                        End If
                        '-----

                    ElseIf sender Is DGV8 Then
                        '航空
                        If sender.Item(dHSender.IndexOf(koumoku("binsyu")(i)), r).Value = "030" Then
                            Select Case sender.Item(dHSender.IndexOf(koumoku("syori2")(i)), r).Value
                                Case HS1.Text
                                    kosuKouku(4) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                                    If InStr(tenpoListKouku(4), tenpo) = 0 Then
                                        If tenpoListKouku(4) = "" Then
                                            tenpoListKouku(4) = tenpo
                                        Else
                                            tenpoListKouku(4) &= "、" & tenpo
                                        End If
                                    End If
                                Case HS2.Text
                                    kosuKouku(3) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                                    If InStr(tenpoListKouku(3), tenpo) = 0 Then
                                        If tenpoListKouku(3) = "" Then
                                            tenpoListKouku(3) = tenpo
                                        Else
                                            tenpoListKouku(3) &= "、" & tenpo
                                        End If
                                    End If
                                Case HS3.Text
                                    kosuKouku(2) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                                    If InStr(tenpoListKouku(2), tenpo) = 0 Then
                                        If tenpoListKouku(2) = "" Then
                                            tenpoListKouku(2) = tenpo
                                        Else
                                            tenpoListKouku(2) &= "、" & tenpo
                                        End If
                                    End If
                                Case HS4.Text
                                    kosuKouku(6) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                                    If InStr(tenpoListKouku(6), tenpo) = 0 Then
                                        If tenpoListKouku(6) = "" Then
                                            tenpoListKouku(6) = tenpo
                                        Else
                                            tenpoListKouku(6) &= "、" & tenpo
                                        End If
                                    End If
                            End Select
                            'kosuKouku(3) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                        Else
                            Select Case sender.Item(dHSender.IndexOf(koumoku("syori2")(i)), r).Value
                                Case HS1.Text
                                    kosu(8) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                                    If InStr(tenpoList(8), tenpo) = 0 Then
                                        If tenpoList(8) = "" Then
                                            tenpoList(8) = tenpo
                                        Else
                                            tenpoList(8) &= "、" & tenpo
                                        End If
                                    End If
                                Case HS2.Text
                                    kosu(3) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                                    If InStr(tenpoList(3), tenpo) = 0 Then
                                        If tenpoList(3) = "" Then
                                            tenpoList(3) = tenpo
                                        Else
                                            tenpoList(3) &= "、" & tenpo
                                        End If
                                    End If
                                Case HS4.Text
                                    kosu(11) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                                    If InStr(tenpoList(10), tenpo) = 0 Then
                                        If tenpoList(10) = "" Then
                                            tenpoList(10) = tenpo
                                        Else
                                            tenpoList(10) &= "、" & tenpo
                                        End If
                                    End If
                                Case HS3.Text
                                    kosu(2) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                                    If InStr(tenpoList(2), tenpo) = 0 Then
                                        If tenpoList(2) = "" Then
                                            tenpoList(2) = tenpo
                                        Else
                                            tenpoList(2) &= "、" & tenpo
                                        End If
                                    End If
                            End Select
                            'kosu(3) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                        End If
                    ElseIf sender Is DGV9 Then
                        Dim isYamato As Boolean = False
                        Dim isMail As Boolean = False
                        If sender.Item(dHSender.IndexOf("マスタ配送"), r).Value = "メール便" Or sender.Item(dHSender.IndexOf("マスタ配送"), r).Value = "メール便" Then
                            isMail = True
                        End If

                        If sender.Item(dHSender.IndexOf("マスタ配送"), r).Value = "ヤマト(陸便)" Or sender.Item(dHSender.IndexOf("マスタ配送"), r).Value = "ヤマト(船便)" Then
                            isYamato = True
                        End If
                        Dim isyupakuGoodBoolTemp As Boolean = False
                        Dim denpyoNum = sender.Item(dHSender.IndexOf("お客様側管理番号"), r).Value

                        'If sender.Item(dHSender.IndexOf("処理用"), r).Value = "ゆう200" Then
                        '    isyupakuGoodBoolTemp = True
                        'End If


                        Select Case sender.Item(dHSender.IndexOf(koumoku("syori2")(i)), r).Value
                            Case HS1.Text
                                If isYamato Then
                                    If sender.Item(dHSender.IndexOf("マスタ配送"), r).Value = "ヤマト(陸便)" Then
                                        kosu(14) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                                    ElseIf sender.Item(dHSender.IndexOf("マスタ配送"), r).Value = "ヤマト(船便)" Then
                                        kosu(16) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                                    End If


                                ElseIf isyupakuGoodBoolTemp Then
                                    '太宰府 yu 2  的路便
                                    kosu(18) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                                    kosu(17) += 0
                                ElseIf isMail Then
                                    kosu(4) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                                Else
                                    Label118.Text += 1
                                    'kosu(4) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)

                                End If

                            Case HS2.Text
                                If isYamato Then
                                    If sender.Item(dHSender.IndexOf("マスタ配送"), r).Value = "ヤマト(陸便)" Then
                                        kosu(15) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                                    ElseIf sender.Item(dHSender.IndexOf("マスタ配送"), r).Value = "ヤマト(船便)" Then
                                        kosu(17) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                                    End If

                                ElseIf isyupakuGoodBoolTemp Then
                                    kosu(20) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                                ElseIf isMail Then

                                    kosu(6) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                                Else
                                    Label119.Text += 1
                                    'kosu(6) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)


                                End If

                                '名古屋
                            Case HS4.Text
                                kosu(12) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                        End Select

                        Dim tl As String = ""
                        Select Case sender.Item(dHSender.IndexOf(koumoku("syori2")(i)), r).Value
                            Case HS1.Text
                                If isYamato Then
                                    If sender.Item(dHSender.IndexOf("マスタ配送"), r).Value = "ヤマト(陸便)" Then
                                        tl = tenpoList(13)
                                    ElseIf sender.Item(dHSender.IndexOf("マスタ配送"), r).Value = "ヤマト(船便)" Then
                                        tl = tenpoList(15)
                                    End If

                                ElseIf isyupakuGoodBoolTemp Then

                                    '太宰府yu2路便   暂定无船便
                                    tl = tenpoList(17)
                                Else
                                    tl = tenpoList(4)
                                End If
                            Case HS2.Text
                                If isYamato Then
                                    If sender.Item(dHSender.IndexOf("マスタ配送"), r).Value = "ヤマト(陸便)" Then
                                        tl = tenpoList(14)
                                    ElseIf sender.Item(dHSender.IndexOf("マスタ配送"), r).Value = "ヤマト(船便)" Then
                                        tl = tenpoList(16)
                                    End If
                                ElseIf isyupakuGoodBoolTemp Then
                                    tl = tenpoList(18)
                                Else
                                    tl = tenpoList(6)
                                End If
                        End Select
                        If InStr(tl, tenpo) = 0 Then
                            If tl = "" Then
                                tl = tenpo
                            Else
                                tl &= "、" & tenpo
                            End If
                        End If
                        Select Case sender.Item(dHSender.IndexOf(koumoku("syori2")(i)), r).Value
                            Case HS1.Text
                                If isYamato Then
                                    tenpoList(4) = tl
                                Else
                                    If sender.Item(dHSender.IndexOf("マスタ配送"), r).Value = "ヤマト(陸便)" Then
                                        tenpoList(13) = tl
                                    ElseIf sender.Item(dHSender.IndexOf("マスタ配送"), r).Value = "ヤマト(船便)" Then
                                        tenpoList(15) = tl
                                    End If
                                End If



                                If isyupakuGoodBoolTemp Then
                                    tenpoList(4) = tl

                                Else
                                    If sender.Item(dHSender.IndexOf("マスタ配送"), r).Value = "ヤマト(陸便)" Then
                                        tenpoList(13) = tl
                                    ElseIf sender.Item(dHSender.IndexOf("マスタ配送"), r).Value = "ヤマト(船便)" Then
                                        tenpoList(15) = tl
                                    End If
                                End If



                            Case HS2.Text
                                If isYamato Then
                                    If sender.Item(dHSender.IndexOf("マスタ配送"), r).Value = "ヤマト(陸便)" Then
                                        tenpoList(14) = tl
                                    ElseIf sender.Item(dHSender.IndexOf("マスタ配送"), r).Value = "ヤマト(船便)" Then
                                        tenpoList(16) = tl
                                    End If
                                Else
                                    tenpoList(6) = tl
                                End If
                            Case HS4.Text
                                tenpoList(11) = tl
                        End Select
                    ElseIf sender Is DGV13 Then
                        Select Case sender.Item(dHSender.IndexOf(koumoku("syori2")(i)), r).Value
                            Case HS1.Text
                                kosu(5) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                            Case HS2.Text
                                kosu(7) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                            Case HS4.Text
                                kosu(13) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                        End Select

                        Dim tl As String = ""
                        Select Case sender.Item(dHSender.IndexOf(koumoku("syori2")(i)), r).Value
                            Case HS1.Text
                                tl = tenpoList(5)
                            Case HS2.Text
                                tl = tenpoList(7)
                            Case HS4.Text
                                tl = tenpoList(12)
                        End Select
                        If InStr(tl, tenpo) = 0 Then
                            If tl = "" Then
                                tl = tenpo
                            Else
                                tl &= "、" & tenpo
                            End If
                        End If
                        Select Case sender.Item(dHSender.IndexOf(koumoku("syori2")(i)), r).Value
                            Case HS1.Text
                                tenpoList(5) = tl
                            Case HS2.Text
                                tenpoList(7) = tl
                            Case HS4.Text
                                tenpoList(12) = tl
                        End Select


                    ElseIf sender Is TMSDGV Then

                        Select Case sender.Item(dHSender.IndexOf(koumoku("syori2")(i)), r).Value
                            '太宰府
                            Case HS1.Text
                                tenpoList(18) = CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                                '井相田
                            Case HS2.Text
                                tenpoList(19) = CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                                '名古屋
                            Case HS4.Text
                                tenpoList(20) = CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                        End Select




                        Select Case sender.Item(dHSender.IndexOf(koumoku("syori2")(i)), r).Value
                            Case HS1.Text
                                kosu(22) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                            Case HS2.Text
                                kosu(23) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                            Case HS4.Text
                                kosu(24) += CInt(sender.Item(dHSender.IndexOf(koumoku("koguchi")(i)), r).Value)
                        End Select








                    End If
                Next
                Exit For
            End If
        Next

        '個数入力
        Select Case True
            Case sender Is DGV7
                LinkLabel1.Text = kosu(0)
                ToolTip1.SetToolTip(LinkLabel1, tenpoList(0))
                LinkLabel2.Text = kosuKouku(0)
                ToolTip1.SetToolTip(LinkLabel2, tenpoListKouku(0))

                LinkLabel16.Text = kosu(1)
                ToolTip1.SetToolTip(LinkLabel16, tenpoList(1))
                LinkLabel15.Text = kosuKouku(1)
                ToolTip1.SetToolTip(LinkLabel15, tenpoListKouku(1))

                LinkLabel25.Text = kosu(10)
                ToolTip1.SetToolTip(LinkLabel25, tenpoList(9))
                LinkLabel26.Text = kosuKouku(5)
                ToolTip1.SetToolTip(LinkLabel26, tenpoListKouku(5))

                If KoukuMeisai <> "" Then
                    TextBox4.Text = KoukuMeisai
                End If
            Case sender Is DGV8
                LinkLabel3.Text = kosu(3)
                ToolTip1.SetToolTip(LinkLabel3, tenpoList(3))
                LinkLabel4.Text = kosuKouku(3)
                ToolTip1.SetToolTip(LinkLabel4, tenpoListKouku(3))

                LinkLabel21.Text = kosu(8)
                ToolTip1.SetToolTip(LinkLabel21, tenpoList(8))
                LinkLabel22.Text = kosuKouku(4)
                ToolTip1.SetToolTip(LinkLabel22, tenpoListKouku(4))

                'LinkLabel19.Text = kosu(2)
                'ToolTip1.SetToolTip(LinkLabel19, tenpoList(2))
                'LinkLabel20.Text = kosuKouku(2)
                'ToolTip1.SetToolTip(LinkLabel20, tenpoListKouku(2))

                LinkLabel23.Text = kosu(11)

                ToolTip1.SetToolTip(LinkLabel23, tenpoList(10))
                LinkLabel24.Text = kosuKouku(6)
                ToolTip1.SetToolTip(LinkLabel24, tenpoListKouku(6))


                TextBox5.Text = KoukuMeisai
            Case sender Is DGV9
                LinkLabel5.Text = kosu(4)
                ToolTip1.SetToolTip(LinkLabel5, tenpoList(4))

                LinkLabel17.Text = kosu(6)
                ToolTip1.SetToolTip(LinkLabel7, tenpoList(6))

                LinkLabel20.Text = kosu(12)
                ToolTip1.SetToolTip(LinkLabel20, tenpoList(11))

                LinkLabel29.Text = kosu(14)
                ToolTip1.SetToolTip(LinkLabel29, tenpoList(13))

                LinkLabel30.Text = kosu(15)
                ToolTip1.SetToolTip(LinkLabel30, tenpoList(14))

                LinkLabel34.Text = kosu(16)
                ToolTip1.SetToolTip(LinkLabel34, tenpoList(15))


                '太宰府的yu2路边
                LinkLabel38.Text = kosu(18)
                ToolTip1.SetToolTip(LinkLabel38, tenpoList(17))
                Console.WriteLine(">>>" + tenpoList(16))

                '太宰府的yu2船边  暂时不会有船便
                LinkLabel37.Text = "0"
                ToolTip1.SetToolTip(LinkLabel37, "0")


            Case sender Is DGV13
                LinkLabel6.Text = kosu(5)
                ToolTip1.SetToolTip(LinkLabel6, tenpoList(5))
                LinkLabel18.Text = kosu(7)
                ToolTip1.SetToolTip(LinkLabel18, tenpoList(7))

                LinkLabel19.Text = kosu(13)
                ToolTip1.SetToolTip(LinkLabel19, tenpoList(12))





            Case sender Is TMSDGV
                '太宰府TMS
                Label128.Text = kosu(22)
                ToolTip1.SetToolTip(Label128, tenpoList(18))


                Label124.Text = kosu(23)
                ToolTip1.SetToolTip(Label124, tenpoList(19))


                Label122.Text = kosu(24)

                ToolTip1.SetToolTip(Label122, tenpoList(20))


        End Select

        'フォント変更
        For i As Integer = 0 To lkArray.Length - 1
            Dim oldFont As Font = lkArray(i).Font
            If lkArray(i).Text = 0 Then
                lkArray(i).Font = New Font(oldFont, oldFont.Style Or FontStyle.Regular)
                lkArray(i).LinkColor = SystemColors.GradientActiveCaption
                lkArray(i).VisitedLinkColor = SystemColors.GradientActiveCaption
            Else
                lkArray(i).Font = New Font(oldFont, FontStyle.Bold)
                lkArray(i).LinkColor = Color.Yellow
                lkArray(i).VisitedLinkColor = Color.Yellow
            End If
        Next
        Dim cc = LinkLabel38.Text
        '実績表示（2019/11/13）
        Csv_denpyo3_F_count.LB11.Text = "+" & (CInt(LinkLabel1.Text) + CInt(LinkLabel21.Text))
        Csv_denpyo3_F_count.LB12.Text = "+" & (CInt(LinkLabel2.Text) + CInt(LinkLabel22.Text))
        Csv_denpyo3_F_count.LB13.Text = "+" & LinkLabel5.Text
        Csv_denpyo3_F_count.LB14.Text = "+" & LinkLabel6.Text
        Csv_denpyo3_F_count.LB35.Text = "+" & LinkLabel29.Text
        Csv_denpyo3_F_count.LB39.Text = "+" & LinkLabel34.Text
        '太宰府的yu2路便
        'Csv_denpyo3_F_count.LB32.Text = "+" & LinkLabel38.Text
        Csv_denpyo3_F_count.LB60.Text = "+" & LinkLabel38.Text
        '太宰府的yu2船便
        Csv_denpyo3_F_count.LB61.Text = "+" & "0"
        Csv_denpyo3_F_count.LB40.Text = "+" & LinkLabel35.Text
        Csv_denpyo3_F_count.LB36.Text = "+" & LinkLabel30.Text
        Csv_denpyo3_F_count.LB15.Text = "+" & (CInt(LinkLabel16.Text) + CInt(LinkLabel3.Text))
        Csv_denpyo3_F_count.LB16.Text = "+" & (CInt(LinkLabel15.Text) + CInt(LinkLabel4.Text))
        Csv_denpyo3_F_count.LB17.Text = "+" & LinkLabel17.Text
        Csv_denpyo3_F_count.LB18.Text = "+" & LinkLabel18.Text

        Csv_denpyo3_F_count.LB31.Text = "+" & (CInt(LinkLabel23.Text) + CInt(LinkLabel25.Text))
        Csv_denpyo3_F_count.LB32.Text = "+" & (CInt(LinkLabel24.Text) + CInt(LinkLabel26.Text))
        Csv_denpyo3_F_count.LB33.Text = "+" & LinkLabel20.Text
        Csv_denpyo3_F_count.LB34.Text = "+" & LinkLabel19.Text



        Csv_denpyo3_F_count.LB47.Text = "+" & Label128.Text
        Csv_denpyo3_F_count.LB48.Text = "+" & Label124.Text
    End Sub

    Private Sub TabControl5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl5.SelectedIndexChanged
        Select Case TabControl5.SelectedTab.Text
            Case "プレビュー"
                If SplitContainer4.Panel2Collapsed = False Then
                    SplitContainer4.Panel2Collapsed = True
                End If
            Case Else
                If SplitContainer4.Panel2Collapsed Then
                    SplitContainer4.Panel2Collapsed = False
                End If
        End Select
    End Sub

    '計算パネル
    Private Sub Label9_VisibleChanged(sender As Object, e As EventArgs) Handles Label9.VisibleChanged
        Panel_Visible()
    End Sub

    Private Sub Panel3_VisibleChanged(sender As Object, e As EventArgs) Handles Panel3.VisibleChanged
        Panel_Visible()
    End Sub

    Private Sub Panel7_VisibleChanged(sender As Object, e As EventArgs) Handles Panel7.VisibleChanged
        Panel_Visible()
    End Sub

    Private Sub Panel8_VisibleChanged(sender As Object, e As EventArgs) Handles Panel8.VisibleChanged
        Panel_Visible()
    End Sub

    Private Sub Panel_Visible()
        If Not Panel3.Visible And Not Panel7.Visible And Not Panel8.Visible And Not Label9.Visible Then
            KryptonButton3.Enabled = True
            KryptonButton1.Enabled = True
            If LockFlag Then    'ロック中は処理できないようにする
                KryptonButton3.Enabled = False
                KryptonButton1.Enabled = False
            End If
        Else
            KryptonButton3.Enabled = False
            KryptonButton1.Enabled = False
        End If
    End Sub

    '検索
    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        If ToolStripTextBox1.Text <> "" Then
            EasySearch()
        End If
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        If ToolStripTextBox1.Text <> "" Then
            EasySearch(True)
        End If
    End Sub

    Private Sub ToolStripTextBox1_TextChanged(sender As Object, e As EventArgs) Handles ToolStripTextBox1.TextChanged
        If ToolStripTextBox1.Text <> "" Then
            EasySearch()
            ToolStripTextBox1.Focus()
        End If
    End Sub

    Private Sub EasySearch(Optional startFlag As Boolean = False)
        Dim dgvS As DataGridView = Nothing
        Select Case TabControl2.SelectedTab.Text
            Case "佐川e飛伝2"
                dgvS = DGV7
            Case "佐川BIZlogi"
                dgvS = DGV8
            Case "メール便"
                dgvS = DGV9
            Case "定形外"
                dgvS = DGV13
            Case "TMS"
                dgvS = TMSDGV
            Case Else
                dgvS = DGV1
        End Select

        If dgvS.RowCount = 0 Then
            Exit Sub
        End If

        Dim search As String = Regex.Replace(ToolStripTextBox1.Text, "\s|　", "|")

        If startFlag = True Then
            dgvS.CurrentCell = dgvS(0, 0)
        ElseIf dgvS.SelectedCells.Count = 0 Then
            dgvS.CurrentCell = dgvS(0, 0)
        End If

        Dim num As Integer = dgvS.CurrentCell.RowIndex * (dgvS.ColumnCount - 1)
        If dgvS.CurrentCell Is Nothing Then
            num = 0
        End If
        num += dgvS.CurrentCell.ColumnIndex
        For r As Integer = 0 To dgvS.Rows.Count - 1
            For c As Integer = 0 To dgvS.Columns.Count - 1
                Dim nows As Integer = (r * (dgvS.ColumnCount - 1)) + c
                If num < nows Then
                    If CStr(dgvS.Item(c, r).Value) <> "" Then
                        If Regex.IsMatch(dgvS.Item(c, r).Value, search) Then
                            dgvS.CurrentCell = dgvS(c, r)
                            Exit Sub
                        End If
                    End If
                End If
            Next
        Next

        MsgBox("検索結果がありませんでした", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
    End Sub

    Private Sub TabControl2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl2.SelectedIndexChanged

        'test()
        If InStr(TabControl2.SelectedTab.Text, "元データ") > 0 Then
            TextBox12.ReadOnly = False
            切り取りToolStripMenuItem1.Enabled = True
            貼り付けToolStripMenuItem1.Enabled = True
            挿入ToolStripMenuItem.Enabled = True
            行を選択直下に複製ToolStripMenuItem.Enabled = True
            削除ToolStripMenuItem1.Enabled = True
            背景色を消すToolStripMenuItem.Enabled = True
            行を削除ToolStripMenuItem1.Enabled = True
        Else
            TextBox12.ReadOnly = True
            切り取りToolStripMenuItem1.Enabled = False
            貼り付けToolStripMenuItem1.Enabled = False
            挿入ToolStripMenuItem.Enabled = False
            行を選択直下に複製ToolStripMenuItem.Enabled = False
            削除ToolStripMenuItem1.Enabled = False
            背景色を消すToolStripMenuItem.Enabled = False
            行を削除ToolStripMenuItem1.Enabled = False
        End If
    End Sub

    Private Sub TextBox12_Validated(sender As Object, e As EventArgs) Handles TextBox12.Validated
        DGV1.SelectedCells(0).Value = TextBox12.Text

        '履歴保存
        If CellValue <> DGV1.CurrentCell.Value Then
            DGV1.CurrentCell.Style.BackColor = Color.Yellow
            ValidateChangeSave(DGV1, DGV1.SelectedCells(0).ColumnIndex, DGV1.SelectedCells(0).RowIndex)
        End If
    End Sub

    Private Sub TextBox12_KeyUp(sender As Object, e As KeyEventArgs) Handles TextBox12.KeyUp
        Select Case e.KeyCode
            Case Keys.Delete, Keys.Back, Keys.Enter
                DGV1.SelectedCells(0).Value = TextBox12.Text
        End Select
    End Sub


    '*****************************************************************************************
    'undo-redo(extensions.vb)
    Dim undoStack As New Stack(Of Object()())
    Dim redoStack As New Stack(Of Object()())
    Dim ignore As Boolean = False
    Dim cellChange As Boolean = True
    Dim cellChangeArray As ArrayList = New ArrayList

    Private Sub DataGridView1_CellValidated(ByVal sender As Object, ByVal e As Windows.Forms.DataGridViewCellEventArgs) Handles DGV1.CellValidated
        If ignore Then Return
        If undoStack.NeedsItem(DGV1) Then
            undoStack.Push(DGV1.Rows.Cast(Of DataGridViewRow).Where(Function(r) Not r.IsNewRow).Select(Function(r) r.Cells.Cast(Of DataGridViewCell).Select(Function(c) c.Value).ToArray).ToArray)
        End If
        元に戻すToolStripMenuItem.Enabled = undoStack.Count > 1
    End Sub

    Private Sub 元に戻すToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 元に戻すToolStripMenuItem.Click
        Dim dgBack As ArrayList = New ArrayList
        For r As Integer = 0 To DGV1.RowCount - 1
            For c As Integer = 0 To DGV1.ColumnCount - 1
                Select Case DGV1.Item(c, r).Style.BackColor.Name
                    Case 0, Color.White.Name

                    Case Else
                        dgBack.Add(c & "," & r & "," & DGV1.Item(c, r).Style.BackColor.Name)
                End Select
            Next
        Next

        If redoStack.Count = 0 OrElse redoStack.NeedsItem(DGV1) Then
            redoStack.Push(DGV1.Rows.Cast(Of DataGridViewRow).Where(Function(r) Not r.IsNewRow).Select(Function(r) r.Cells.Cast(Of DataGridViewCell).Select(Function(c) c.Value).ToArray).ToArray)
        End If
        Dim rows()() As Object = undoStack.Pop
        While rows.ItemEquals(DGV1.Rows.Cast(Of DataGridViewRow).Where(Function(r) Not r.IsNewRow).ToArray)
            rows = undoStack.Pop
        End While
        ignore = True
        DGV1.Rows.Clear()
        For x As Integer = 0 To rows.GetUpperBound(0)
            DGV1.Rows.Add(rows(x))
        Next
        ignore = False
        元に戻すToolStripMenuItem.Enabled = undoStack.Count > 0
        やり直しToolStripMenuItem.Enabled = redoStack.Count > 0

        For i As Integer = 0 To dgBack.Count - 1
            Dim cr As String() = Split(dgBack(i), ",")
            Dim colorName = New ColorConverter().ConvertFromString(cr(2))   'カラーコンバーター
            DGV1.Item(CInt(cr(0)), CInt(cr(1))).Style.BackColor = colorName
        Next
    End Sub

    Private Sub やり直しToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles やり直しToolStripMenuItem.Click
        If undoStack.Count = 0 OrElse undoStack.NeedsItem(DGV1) Then
            undoStack.Push(DGV1.Rows.Cast(Of DataGridViewRow).Where(Function(r) Not r.IsNewRow).Select(Function(r) r.Cells.Cast(Of DataGridViewCell).Select(Function(c) c.Value).ToArray).ToArray)
        End If
        Dim rows()() As Object = redoStack.Pop
        While rows.ItemEquals(DGV1.Rows.Cast(Of DataGridViewRow).Where(Function(r) Not r.IsNewRow).ToArray)
            rows = redoStack.Pop
        End While
        ignore = True
        DGV1.Rows.Clear()
        For x As Integer = 0 To rows.GetUpperBound(0)
            DGV1.Rows.Add(rows(x))
        Next
        ignore = False
        やり直しToolStripMenuItem.Enabled = redoStack.Count > 0
        元に戻すToolStripMenuItem.Enabled = undoStack.Count > 0
    End Sub
    '*****************************************************************************************

    'CSV移行
    Private Sub Csv1ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Csv1ToolStripMenuItem.Click
        CsvIkou(Form1.CSV_FORMS(0), Form1.CSV_FORMS(0).DataGridView1)
    End Sub

    Private Sub Csv2ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Csv2ToolStripMenuItem.Click
        CsvIkou(Form1.CSV_FORMS(1), Form1.CSV_FORMS(1).DataGridView1)
    End Sub

    Private Sub CsvIkou(ByVal frm As Object, ByVal csvDg As DataGridView)
        If csvDg.RowCount > 0 Then
            csvDg.Rows.Clear()
            csvDg.Columns.Clear()
        End If

        Dim dgv As DataGridView = DGV1
        Select Case TabControl2.SelectedTab.Text
            Case "佐川e飛伝2"
                dgv = DGV7
            Case "佐川BIZlogi"
                dgv = DGV8
            Case "メール便"
                dgv = DGV9
            Case "定形外"
                dgv = DGV13
            Case Else
                dgv = DGV1
        End Select

        '-----------------------------------------------------------------------
        frm.preDIV()
        '-----------------------------------------------------------------------
        For c As Integer = 0 To dgv.ColumnCount - 1
            csvDg.Columns.Add("column" & c + 1, c)
        Next
        csvDg.Rows.Add()
        For c As Integer = 0 To dgv.ColumnCount - 1
            csvDg.Item(c, 0).Value = dgv.Columns(c).HeaderText
        Next

        Dim flag As Boolean = True
        For r As Integer = 0 To dgv.RowCount - 1
            If dgv.Rows(r).IsNewRow = False Then
                csvDg.Rows.Add()
                For c As Integer = 0 To dgv.ColumnCount - 1
                    csvDg.Item(c, r + 1).Value = dgv.Item(c, r).Value
                Next
            End If
        Next
        '-----------------------------------------------------------------------
        frm.retDIV()
        '-----------------------------------------------------------------------

        FormFront(frm)
        MsgBox("移行完了", MsgBoxStyle.SystemModal)
    End Sub

    Private Sub 行抽出を戻すToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 行抽出を戻すToolStripMenuItem.Click
        Dim dgv As DataGridView() = {DGV1, DGV7, DGV8, DGV9, DGV13}
        For i As Integer = 0 To dgv.Length - 1
            For r As Integer = 0 To dgv(i).RowCount - 1
                dgv(i).Rows(r).Visible = True
            Next
        Next
        MsgBox("行抽出を戻しました", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information Or MsgBoxStyle.SystemModal)
    End Sub

    Private Sub 行抽出黄色ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 行抽出黄色ToolStripMenuItem.Click
        MsgBox("未完成", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information Or MsgBoxStyle.SystemModal)
    End Sub

    Private Sub 行抽出青ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 行抽出青ToolStripMenuItem.Click
        MsgBox("未完成", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information Or MsgBoxStyle.SystemModal)
    End Sub

    Private Sub 行抽出オレンジToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 行抽出オレンジToolStripMenuItem.Click
        MsgBox("未完成", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information Or MsgBoxStyle.SystemModal)
    End Sub

    Private Sub 行抽出赤ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 行抽出赤ToolStripMenuItem.Click
        MsgBox("未完成", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information Or MsgBoxStyle.SystemModal)
    End Sub

    Private Sub DataGridView3_SelectionChanged(sender As Object, e As EventArgs) Handles DGV3.SelectionChanged
        If DGV3.SelectedCells.Count > 0 Then
            ToolStripStatusLabel2.Text = DGV3.SelectedCells(0).Value
            If ToolStripStatusLabel2.Text <> "" Then
                ToolStripStatusLabel2.Text = Regex.Replace(ToolStripStatusLabel2.Text, "\n|\r|\n\r", "")
            End If
        Else
            ToolStripStatusLabel2.Text = "[情報]"
        End If
    End Sub






    '============================================
    '--------------------------------------------
    'メイン右
    '--------------------------------------------
    '============================================
    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        If ToolStripButton2.Text = ">>>" Then
            ToolStripButton2.Text = "<<<"
            SplitContainer2.Panel2Collapsed = True
        Else
            ToolStripButton2.Text = ">>>"
            SplitContainer2.Panel2Collapsed = False
        End If
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        '20180316
        TextBox22.Text = Format(DateTimePicker1.Value, "yyyyMMdd")
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        TextBox22.Text = ""
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        ComboBox5.SelectedIndex = 0
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ComboBox2.SelectedItem = ComboBox1.SelectedItem
        If ComboBox1.SelectedItem = "ヤマト" Then
            ComboBox7.SelectedItem = "ヤマト(陸便)"
        End If
        TextBox7.Text = ComboBox1.SelectedItem
    End Sub

    Dim CB7txt As String = ""
    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        Select Case CB7txt
            Case "メール便", "定形外"
                If ComboBox7.SelectedItem = "宅配便" Then
                    NumericUpDown2.Value = 1
                End If
        End Select
    End Sub

    '伝票操作画面に表示する
    Private Sub DenpyoSousa(sender As DataGridView, num As Integer, selRow As Integer)
        Dim dH As ArrayList = TM_HEADER_GET(sender)
        Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)

        TextBox19.Text = sender.Item(dH.IndexOf(koumoku("denpyoNo")(num)), selRow).Value
        TextBox20.Text = sender.Item(dH.IndexOf(koumoku("sakiname")(num)), selRow).Value

        'dgv1のデータを検索する
        Dim targetRow As Integer = -1
        Select Case True
            Case sender Is DGV1
                targetRow = selRow
            Case Else
                For r As Integer = 0 To DGV1.RowCount - 1
                    If TextBox19.Text = DGV1.Item(dH1.IndexOf(koumoku("denpyoNo")(0)), r).Value Then
                        targetRow = r
                        Exit For
                    End If
                Next
        End Select
        If targetRow = -1 Then
            Exit Sub
        End If

        'dgv1からデータを表示する
        Dim hassouSouko As String = DGV1.Item(dH1.IndexOf("発送倉庫"), targetRow).Value
        Dim hassou As String = DGV1.Item(dH1.IndexOf("発送方法"), targetRow).Value
        Dim dSoft As String = DGV1.Item(dH1.IndexOf("伝票ソフト"), targetRow).Value
        If dSoft = "" Then
            If hassou = "宅配便" Then
                If hassouSouko = HS1.Text Then
                    ComboBox1.SelectedItem = CheckBox1.Text
                ElseIf hassouSouko = HS2.Text Then
                    ComboBox1.SelectedItem = CheckBox4.Text
                    'ElseIf hassouSouko = HS3.Text Then
                    '    ComboBox1.SelectedItem = CheckBox4.Text
                ElseIf hassouSouko = HS4.Text Then
                    ComboBox1.SelectedItem = CheckBox32.Text
                End If
            ElseIf hassou = "メール便" Then
                ComboBox1.SelectedItem = "メール便"
            ElseIf hassou = "定形外" Then
                ComboBox1.SelectedItem = "定形外"
            ElseIf hassou = "ヤマト" Then
                ComboBox1.SelectedItem = "ヤマト"
            Else
                ComboBox1.SelectedItem = ""
            End If
        Else
            ComboBox1.SelectedItem = dSoft
        End If

        TextBox22.Text = DGV1.Item(dH1.IndexOf(koumoku("siteiBi")(0)), targetRow).Value
        ComboBox7.SelectedItem = DGV1.Item(dH1.IndexOf(koumoku("masterHaiso")(0)), targetRow).Value
        CB7txt = ComboBox7.SelectedItem     'メール便、定形外から宅配便に変更した時、便数を1にするため、前の値を保存

        ComboBox2.SelectedItem = DGV1.Item(dH1.IndexOf("伝票ソフト"), targetRow).Value
        ComboBox3.SelectedItem = DGV1.Item(dH1.IndexOf("発送倉庫"), targetRow).Value

        Dim jikan As String = DGV1.Item(dH1.IndexOf(koumoku("siteiJikan")(0)), targetRow).Value
        ComboBox5.SelectedItem = JikanToStr(jikan)

        Dim goukei As String = DGV1.Item(dH1.IndexOf("合計金額"), targetRow).Value
        TextBox21.Text = goukei
        If goukei = "" Or goukei = 0 Then
            ComboBox4.SelectedItem = "元払い"
        Else
            ComboBox4.SelectedItem = "代引き"
        End If

        Dim binsyu As String = DGV1.Item(dH1.IndexOf("便種"), targetRow).Value
        Select Case binsyu
            Case "", "陸便", "000"
                RadioButton9.Checked = True
            Case Else
                RadioButton10.Checked = True
        End Select

        NumericUpDown2.Value = DGV1.Item(dH1.IndexOf("マスタ便数"), targetRow).Value
        If InStr(DGV1.Item(dH1.IndexOf("購入者電話番号"), targetRow).Value, ",") > 0 Then
            Label47.ForeColor = Color.Red
            Label47.Text = "便(NE:" & Split(DGV1.Item(dH1.IndexOf("購入者電話番号"), targetRow).Value, ",")(1) & "便)"
        Else
            Label47.ForeColor = Color.Black
            Label47.Text = "便"
        End If

        TextBox34.Text = DGV1.Item(dH1.IndexOf("マーク"), targetRow).Value
        TextBox33.Text = DGV1.Item(dH1.IndexOf("マスタ備考"), targetRow).Value

        '----------------------

        TextBox44.Text = sender.Item(dH.IndexOf(koumoku("denpyoNo")(num)), selRow).Value

        TextBox43.Text = DGV1.Item(dH1.IndexOf("発送先郵便番号"), targetRow).Value
        TextBox45.Text = DGV1.Item(dH1.IndexOf("発送先住所"), targetRow).Value
        TextBox46.Text = DGV1.Item(dH1.IndexOf("発送先名"), targetRow).Value
        TextBox47.Text = DGV1.Item(dH1.IndexOf("発送先電話番号"), targetRow).Value

        If TextBox45.Text <> "" Then
            Dim res As String = KanaConverter.RomajiToHiragana(TextBox45.Text)
            res = KanaConverter.HiraganaToKatakana(res)
            res = Regex.Replace(res, "\s|　|\'|""|\||\,|\?|#|〓", "")
            If CStr(res) <> "" Then
                If Regex.IsMatch(res, "&#.*?\;") Then   '機種依存文字を修正する
                    Dim m As MatchCollection = Regex.Matches(res, "&#.*?\;")
                    For Each mStr In m
                        Dim mm1 As String = Replace(mStr.ToString, "&#", "&H")
                        Dim mm2 As String = Replace(mm1, ";", "")
                        res = Regex.Replace(res, mStr.ToString, Chr(mm2))
                    Next
                End If
            End If
            TextBox51.Text = res
        Else
            TextBox51.Text = ""
        End If

        If TextBox46.Text <> "" Then
            TextBox52.Text = Regex.Replace(TextBox46.Text, "\s|　", "")
            TextBox50.Text = KanaConverter.RomajiToHiragana(TextBox46.Text)
            TextBox50.Text = KanaConverter.HiraganaToKatakana(TextBox50.Text)
            TextBox50.Text = Regex.Replace(TextBox50.Text, "\s|　", "")
            Dim res As String = TextBox50.Text
            If Regex.IsMatch(res, "&#.*?\;") Then   '機種依存文字を修正する
                Dim m As MatchCollection = Regex.Matches(res, "&#.*?\;")
                For Each mStr In m
                    'Dim mm1 As String = Replace(mStr.ToString, "&#", "&H")
                    Dim mm1 As String = Replace(mStr.ToString, "&#", "")
                    Dim mm2 As String = Replace(mm1, ";", "")
                    res = Regex.Replace(res, mStr.ToString, ChrW(mm2))
                Next
            End If
            res = Replace(res, "株式会社", "（株）")
            res = Replace(res, "有限会社", "（有）")
            res = Replace(res, "合同会社", "（合）")
            TextBox50.Text = res
        Else
            TextBox52.Text = ""
            TextBox50.Text = ""
        End If

        '----------------------

        If DGV1.Item(dH1.IndexOf("営業所留め"), targetRow).Value = "1" Then
            Label67.Visible = True
        Else
            Label67.Visible = False
        End If
        TextBox28.Text = DGV1.Item(dH1.IndexOf("営業所コード"), targetRow).Value
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim str As String = TextBox45.Text
        TextBox45.Text = TextBox51.Text
        TextBox51.Text = str
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim str As String = TextBox46.Text
        TextBox46.Text = TextBox52.Text
        TextBox52.Text = str
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim str As String = TextBox46.Text
        TextBox46.Text = TextBox50.Text
        TextBox50.Text = str
    End Sub


    Private Function JikanToStr(jikan As String)
        Dim res As String = ""
        Select Case jikan
            Case "01"
                res = "午前中"
            Case "12"
                res = "12時-14時"
            Case "14"
                res = "14時-16時"
            Case "16"
                res = "16時-18時"
            Case "18"
                res = "18時-20時"
            Case "19"
                res = "19時-21時"
            Case "04"
                res = "18時-21時"
            Case Else
                res = jikan
        End Select
        Return res
    End Function

    Private Function StrToJikan(jikan As String)
        Dim res As String = ""
        Select Case jikan
            Case "午前中"
                res = "01"
            Case "12時-14時"
                res = "12"
            Case "14時-16時"
                res = "14"
            Case "16時-18時"
                res = "16"
            Case "18時-20時"
                res = "18"
            Case "19時-21時"
                res = "19"
            Case "18時-21時"
                res = "04"
            Case Else
                res = jikan
        End Select
        Return res
    End Function

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        If ComboBox4.SelectedItem = "代引き" Then
            TextBox21.Enabled = True
        Else
            TextBox21.Enabled = False
        End If
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        Clipboard.SetText(TextBox19.Text)
    End Sub

    '適用ボタン
    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        'エラーチェック
        If DGV1.RowCount = 0 Or TextBox19.Text = "" Then
            Exit Sub
        End If
        If ComboBox1.SelectedItem = "" Or ComboBox7.SelectedItem = "" Then
            MsgBox("発送方法と倉庫を選択してください")
            Exit Sub
        End If

        If ComboBox1.SelectedItem = "メール便" And ComboBox7.SelectedItem <> "メール便" Then

        ElseIf Regex.IsMatch(ComboBox1.SelectedItem, "e飛伝2|BIZlogi") And ComboBox7.SelectedItem <> "宅配便" Then

        End If

        TabPage13_check(tp13ArrayA)
        Dim dNo As String = TextBox19.Text

        'dgv1の選択行を伝票番号から取得
        Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
        Dim selRow As Integer = 0
        For r As Integer = 0 To DGV1.RowCount - 1
            If DGV1.Item(dH1.IndexOf("伝票番号"), r).Value = dNo Then
                selRow = r
                Exit For
            End If
        Next

        '修正した箇所に色を付ける
        Dim cell As DataGridViewCell = Nothing

        cell = DGV1.Item(dH1.IndexOf("発送方法"), selRow)
        If ComboBox1.SelectedItem = "メール便" Then
            If cell.Value <> "メール便" Then
                cell.Style.BackColor = Color.Yellow
                cell.Value = "メール便"
            End If
        ElseIf ComboBox1.SelectedItem = "定形外" Then
            If cell.Value <> "定形外" Then
                cell.Style.BackColor = Color.Yellow
                cell.Value = "定形外"
            End If
        ElseIf ComboBox1.SelectedItem = "ヤマト" Then
            If cell.Value <> "ヤマト" Then
                cell.Style.BackColor = Color.Yellow
                cell.Value = "ヤマト"
            End If
        Else
            If cell.Value <> "宅配便" Then
                cell.Style.BackColor = Color.Yellow
                cell.Value = "宅配便"
            End If
        End If

        cell = DGV1.Item(dH1.IndexOf("マスタ配送"), selRow)
        If cell.Value <> ComboBox7.SelectedItem Then
            cell.Style.BackColor = Color.Yellow
            cell.Value = ComboBox7.SelectedItem
        End If

        cell = DGV1.Item(dH1.IndexOf("伝票ソフト"), selRow)
        If cell.Value <> ComboBox2.SelectedItem Then
            cell.Style.BackColor = Color.Yellow
            cell.Value = ComboBox2.SelectedItem
        End If

        cell = DGV1.Item(dH1.IndexOf("発送倉庫"), selRow)
        If cell.Value <> ComboBox3.SelectedItem Then
            cell.Style.BackColor = Color.Yellow
            cell.Value = ComboBox3.SelectedItem
        End If

        cell = DGV1.Item(dH1.IndexOf("配送日"), selRow)
        If cell.Value <> TextBox22.Text Then
            cell.Style.BackColor = Color.Yellow
            cell.Value = TextBox22.Text
        End If

        cell = DGV1.Item(dH1.IndexOf("配送時間帯"), selRow)
        If cell.Value <> StrToJikan(ComboBox5.SelectedItem) Then
            cell.Style.BackColor = Color.Yellow
            cell.Value = StrToJikan(ComboBox5.SelectedItem)
        End If

        cell = DGV1.Item(dH1.IndexOf("便種"), selRow)
        If RadioButton9.Checked Then
            If cell.Value <> RadioButton9.Text Then
                cell.Style.BackColor = Color.Yellow
                cell.Value = RadioButton9.Text
            End If
        ElseIf RadioButton10.Checked Then
            If cell.Value <> RadioButton10.Text Then
                cell.Style.BackColor = Color.Yellow
                cell.Value = RadioButton10.Text
            End If
        End If

        cell = DGV1.Item(dH1.IndexOf("マスタ便数"), selRow)
        If cell.Value <> NumericUpDown2.Value Then
            cell.Style.BackColor = Color.Yellow
            cell.Value = NumericUpDown2.Value
        End If

        cell = DGV1.Item(dH1.IndexOf("マーク"), selRow)
        If cell.Value <> TextBox34.Text Then
            cell.Style.BackColor = Color.Yellow
            cell.Value = TextBox34.Text
        End If
        cell = DGV1.Item(dH1.IndexOf("マスタ備考"), selRow)
        If cell.Value <> TextBox33.Text Then
            cell.Style.BackColor = Color.Yellow
            cell.Value = TextBox33.Text
        End If

        'タブを変え、修正箇所へ飛ぶ
        If TabControl2.SelectedIndex <> 0 Then
            For c As Integer = 0 To DGV1.ColumnCount - 1
                If DGV1.Item(c, selRow).Style.BackColor = Color.Yellow Then
                    DGV1.CurrentCell = DGV1(c, selRow)
                End If
            Next
            TabControl2.SelectedIndex = 0
        End If

        '再処理
        'KryptonButton3.PerformClick()
        'Application.DoEvents()

        'プレビュー
        'SplitContainer4.Panel2Collapsed = True
        'TabControl5.SelectedIndex = 1
        '編集になる伝票を取得
        Dim motoRow As New ArrayList
        Dim dgvM As DataGridView() = {DGV7, DGV8, DGV9, DGV13}
        Dim dHMoto As New ArrayList
        Do
            For i As Integer = 0 To dgvM.Length - 1
                dHMoto = TM_HEADER_GET(dgvM(i))
                For r As Integer = 0 To dgvM(i).RowCount - 1
                    If dNo = dgvM(i).Item(dHMoto.IndexOf(koumoku("denpyoNo")(i + 1)), r).Value Then
                        motoRow.Add(i & "," & r)
                        Exit Do
                    End If
                Next
            Next
            Exit Do
        Loop

        DGV12.Rows.Clear()
        DGV12.Columns.Clear()
        If motoRow.Count > 0 Then
            Dim header_mR As String() = Split(motoRow(0), ",")      '保存した1行目のdgv番号からヘッダー行を作成する
            Dim dHm As ArrayList = TM_HEADER_GET(dgvM(header_mR(0)))
            For c As Integer = 0 To dHm.Count - 1
                DGV12.Columns.Add(c, dHm(c))
            Next
            Dim changeFlag As Boolean = False
            For i As Integer = 0 To motoRow.Count - 1               'プレビュー作成
                Dim mR As String() = Split(motoRow(i), ",")
                DGV12.Rows.Add()
                For c As Integer = 0 To dgvM(mR(0)).ColumnCount - 1 '（難読）dgv12=新規行←input(元dgvの「motoRow(1)=mr(1)」)を反映
                    DGV12.Item(c, DGV12.RowCount - 1).Value = dgvM(mR(0)).Item(c, CInt(mR(1))).Value
                    DGV12.Item(c, DGV12.RowCount - 1).Style.BackColor = dgvM(mR(0)).Item(c, CInt(mR(1))).Style.BackColor
                    If changeFlag = False And dgvM(mR(0)).Item(c, CInt(mR(1))).Style.BackColor = Color.Yellow Then  '修正のある所は黄色背景なので、最初の所に移動
                        DGV12.CurrentCell = DGV12(c, DGV12.RowCount - 1)
                        DGV12.ClearSelection()
                        changeFlag = True
                    End If
                Next
            Next
        End If

        '履歴保存
        If CellValue <> DGV1.CurrentCell.Value Then
            DGV1.CurrentCell.Style.BackColor = Color.Yellow
            ValidateChangeSave(DGV1, DGV1.SelectedCells(0).ColumnIndex, DGV1.SelectedCells(0).RowIndex)
        End If

    End Sub

    '------------------------
    '修正内容を保存するか確認
    Dim tp13ArrayA As New ArrayList
    Dim tp13ArrayB As New ArrayList
    Private Sub TabPage13_Enter(sender As Object, e As EventArgs) Handles TabPage13.Enter
        TabPage13_check(tp13ArrayA)
    End Sub

    Private Sub TabPage13_Leave(sender As Object, e As EventArgs) Handles TabPage13.Leave
        '伝票番号が同じ時だけ
        If tp13ArrayA.Count > 0 Then
            If tp13ArrayA(0) <> TextBox19.Text Then
                Exit Sub
            End If
        End If

        TabPage13_check(tp13ArrayB)
        Dim flag As Boolean = False
        For i As Integer = 0 To tp13ArrayA.Count - 1
            If tp13ArrayA(i) <> tp13ArrayB(i) Then
                flag = True
                Exit For
            End If
        Next
        If flag Then
            Dim DR As DialogResult = MsgBox("修正内容を適用していません。適用しますか？", MsgBoxStyle.OkCancel Or MsgBoxStyle.SystemModal)
            If DR = DialogResult.OK Then
                Button21.PerformClick()
            End If
        End If
    End Sub

    Private Sub TabPage13_check(a As ArrayList)
        a.Clear()
        a.Add(TextBox19.Text)
        a.Add(ComboBox1.SelectedItem)
        a.Add(TextBox22.Text)
        a.Add(ComboBox7.SelectedItem)
        a.Add(ComboBox2.SelectedItem)
        a.Add(ComboBox3.SelectedItem)
        a.Add(ComboBox5.SelectedItem)
        a.Add(ComboBox6.SelectedItem)
        a.Add(ComboBox4.SelectedItem)
        a.Add(TextBox21.Text)
        a.Add(RadioButton9.Checked)
        a.Add(NumericUpDown2.Value)
        a.Add(TextBox34.Text)
        a.Add(TextBox33.Text)
    End Sub
    '------------------------

    Private Sub TextBox45_TextChanged(sender As Object, e As EventArgs) Handles TextBox45.TextChanged
        Label27.Text = TextBox45.Text.Length
        Dim max As Integer = 16 * 3
        Select Case TextBox7.Text
            Case "e飛伝2"
                max = 16 * 3
            Case "BIZlogi"
                max = 32 * 3
            Case "メール便", "定型外"
                max = 25 * 2
            Case Else
                max = 16 * 3
        End Select
        If TextBox45.Text.Length > max Then
            TextBox45.BackColor = Color.Yellow
        Else
            TextBox45.BackColor = Color.White
        End If
    End Sub

    Private Sub TextBox51_TextChanged(sender As Object, e As EventArgs) Handles TextBox51.TextChanged
        Label64.Text = TextBox51.Text.Length
    End Sub

    Private Sub TextBox46_TextChanged(sender As Object, e As EventArgs) Handles TextBox46.TextChanged
        Label26.Text = TextBox46.Text.Length
        Dim max As Integer = 16
        Select Case TextBox7.Text
            Case "e飛伝2"
                max = 16
            Case "BIZlogi"
                max = 32
            Case "メール便", "定型文"
                max = 25
            Case Else
                max = 16
        End Select
        If TextBox46.Text.Length > max Then
            TextBox46.BackColor = Color.Yellow
        Else
            TextBox46.BackColor = Color.White
        End If
    End Sub

    Private Sub TextBox52_TextChanged(sender As Object, e As EventArgs) Handles TextBox52.TextChanged
        Label65.Text = TextBox52.Text.Length
    End Sub

    Private Sub TextBox50_TextChanged(sender As Object, e As EventArgs) Handles TextBox50.TextChanged
        Label66.Text = TextBox50.Text.Length
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TabPage19_check(tp19ArrayA)
        Dim dNo As String = TextBox44.Text

        '編集になる伝票を取得（処理後のリスト）
        Dim motoRow As New ArrayList
        Dim dgvM As DataGridView() = {DGV7, DGV8, DGV9, DGV13}
        Dim dHMoto As New ArrayList
        Do
            For i As Integer = 0 To dgvM.Length - 1
                dHMoto = TM_HEADER_GET(dgvM(i))
                For r As Integer = 0 To dgvM(i).RowCount - 1
                    If dNo = dgvM(i).Item(dHMoto.IndexOf(koumoku("denpyoNo")(i + 1)), r).Value Then
                        motoRow.Add(i & "," & r)
                        Exit Do
                    End If
                Next
            Next
            Exit Do
        Loop

        'dgv1の選択行を伝票番号から取得
        Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
        Dim selRow As Integer = 0
        For r As Integer = 0 To DGV1.RowCount - 1
            If DGV1.Item(dH1.IndexOf("伝票番号"), r).Value = dNo Then
                selRow = r
                Exit For
            End If
        Next

        '修正した箇所に色を付ける
        Dim cell As DataGridViewCell = Nothing

        cell = DGV1.Item(dH1.IndexOf("発送先郵便番号"), selRow)
        If cell.Value <> TextBox43.Text Then
            cell.Style.BackColor = Color.Yellow
            cell.Value = TextBox43.Text
        End If

        cell = DGV1.Item(dH1.IndexOf("発送先住所"), selRow)
        If cell.Value <> TextBox45.Text Then
            cell.Style.BackColor = Color.Yellow
            Dim addr As String = TextBox45.Text
            addr = Regex.Replace(addr, "\r|\n|\r\n", "")
            cell.Value = addr
        End If

        cell = DGV1.Item(dH1.IndexOf("発送先名"), selRow)
        If cell.Value <> TextBox46.Text Then
            cell.Style.BackColor = Color.Yellow
            cell.Value = TextBox46.Text
        End If

        cell = DGV1.Item(dH1.IndexOf("発送先電話番号"), selRow)
        If cell.Value <> TextBox47.Text Then
            cell.Style.BackColor = Color.Yellow
            cell.Value = TextBox47.Text
        End If

        'タブを変え、修正箇所へ飛ぶ
        If TabControl2.SelectedIndex <> 0 Then
            TabControl2.SelectedIndex = 0
            For c As Integer = 0 To DGV1.ColumnCount - 1
                If DGV1.Item(c, selRow).Style.BackColor = Color.Yellow Then
                    DGV1.CurrentCell = DGV1(c, selRow)
                End If
            Next
        End If

        'プレビュー
        If motoRow.Count > 0 Then
            'SplitContainer4.Panel2Collapsed = True
            'TabControl5.SelectedIndex = 1
            DGV12.Rows.Clear()
            DGV12.Columns.Clear()
            Dim header_mR As String() = Split(motoRow(0), ",")      '保存した1行目のdgv番号からヘッダー行を作成する
            Dim dHm As ArrayList = TM_HEADER_GET(dgvM(header_mR(0)))
            For c As Integer = 0 To dHm.Count - 1
                DGV12.Columns.Add(c, dHm(c))
            Next
            Dim changeFlag As Boolean = False
            For i As Integer = 0 To motoRow.Count - 1               'プレビュー作成
                Dim mR As String() = Split(motoRow(i), ",")
                DGV12.Rows.Add()
                For c As Integer = 0 To dgvM(mR(0)).ColumnCount - 1 '（難読）dgv12=新規行←input(元dgvの「motoRow(1)=mr(1)」)を反映
                    DGV12.Item(c, DGV12.RowCount - 1).Value = dgvM(mR(0)).Item(c, CInt(mR(1))).Value
                    DGV12.Item(c, DGV12.RowCount - 1).Style.BackColor = dgvM(mR(0)).Item(c, CInt(mR(1))).Style.BackColor
                    If changeFlag = False And dgvM(mR(0)).Item(c, CInt(mR(1))).Style.BackColor = Color.Yellow Then  '修正のある所は黄色背景なので、最初の所に移動
                        DGV12.CurrentCell = DGV12(c, DGV12.RowCount - 1)
                        DGV12.ClearSelection()
                        changeFlag = True
                    End If
                Next
            Next
        End If

        '履歴保存
        If CellValue <> DGV1.CurrentCell.Value Then
            DGV1.CurrentCell.Style.BackColor = Color.Yellow
            ValidateChangeSave(DGV1, DGV1.SelectedCells(0).ColumnIndex, DGV1.SelectedCells(0).RowIndex)
        End If

    End Sub

    '------------------------
    '修正内容を保存するか確認
    Dim tp19ArrayA As New ArrayList
    Dim tp19ArrayB As New ArrayList
    Private Sub TabPage19_Enter(sender As Object, e As EventArgs) Handles TabPage19.Enter
        TabPage19_check(tp19ArrayA)
    End Sub

    Private Sub TabPage19_Leave(sender As Object, e As EventArgs) Handles TabPage19.Leave
        '伝票番号が同じ時だけ
        If tp19ArrayA.Count > 0 Then
            If tp19ArrayA(0) <> TextBox44.Text Then
                Exit Sub
            End If
        End If

        TabPage19_check(tp19ArrayB)
        Dim flag As Boolean = False
        For i As Integer = 0 To tp19ArrayA.Count - 1
            If tp19ArrayA(i) <> tp19ArrayB(i) Then
                flag = True
                Exit For
            End If
        Next
        If flag Then
            Dim DR As DialogResult = MsgBox("修正内容を適用していません。適用しますか？", MsgBoxStyle.OkCancel Or MsgBoxStyle.SystemModal)
            If DR = DialogResult.OK Then
                Button1.PerformClick()
            End If
        End If
    End Sub

    Private Sub TabPage19_check(a As ArrayList)
        a.Clear()
        a.Add(TextBox44.Text)
        a.Add(TextBox43.Text)
        a.Add(TextBox45.Text)
        a.Add(TextBox46.Text)
        a.Add(TextBox47.Text)
    End Sub
    '------------------------


    '商品の検索
    Private Sub Button33_Click(sender As Object, e As EventArgs) Handles Button33.Click
        Shohin_Search()
    End Sub

    Private Sub TextBox36_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox36.KeyDown
        If e.KeyCode = Keys.Enter Then
            Shohin_Search()
        End If
    End Sub

    Private Sub Shohin_Search()
        DGV5.Rows.Clear()

        'datagridviewのヘッダーテキストをコレクションに取り込む
        'Dim dH6 As ArrayList = TM_HEADER_GET(dgv6)

        Dim listFlag As Boolean = False
        For r As Integer = 0 To DGV6.RowCount - 1
            If InStr(DGV6.Item(TM_ArIndexof(dH6, "商品コード"), r).Value, TextBox36.Text) > 0 Then
                Dim str1 As String = ValiationAdd1(DGV6.Item(TM_ArIndexof(dH6, "商品コード"), r).Value, "", NumericUpDown3.Value)
                Dim str1_ As String() = Split(str1, "|")
                Dim str2 As String = DGV6.Item(TM_ArIndexof(dH6, "商品名"), r).Value
                DGV5.Rows.Add(str1_(0), str2)
            End If
        Next
    End Sub

    Private Sub DataGridView5_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV5.CellDoubleClick
        If DGV1.RowCount > 0 Then
            Dim header As String() = New String() {"商品コード1", "商品コード2", "商品コード3", "商品コード4"}
            Dim inputStr As String = DGV5.SelectedCells(0).Value
            Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
            If dH1.Contains(header(0)) Then
                Dim selRow As Integer = DGV1.SelectedCells(0).RowIndex
                For i As Integer = 0 To 3
                    If CStr(DGV1.Item(dH1.IndexOf(header(i)), selRow).Value) = "" Then
                        DGV1.Item(dH1.IndexOf(header(i)), selRow).Value = inputStr
                        DGV1.CurrentCell = DGV1(dH1.IndexOf(header(i)), selRow)
                        TabControl2.SelectedIndex = 0
                        Exit For
                    End If
                Next
            End If
        End If
    End Sub

    '********************************************
    '営業所留め
    '********************************************
    Dim CnAccdb_ken_all As String = Path.GetDirectoryName(Form1.appPath) & "\db\ken_all.accdb"
    Dim countNum As New ArrayList
    Dim sichoYomi As New ArrayList
    Dim choYomi As New ArrayList

    Private Sub Button31_Click(sender As Object, e As EventArgs) Handles Button31.Click
        If DGV1.RowCount > 0 And DGV1.SelectedCells.Count > 0 Then
            'datagridviewのヘッダーテキストをコレクションに取り込む
            Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)

            Dim selRow As Integer = DGV1.SelectedCells(0).RowIndex
            Dim address As String = DGV1.Item(TM_ArIndexof(dH1, "発送先住所"), selRow).Value
            address = Replace(address, " ", "")
            address = Replace(address, "　", "")

            For i As Integer = 0 To ComboBox10.Items.Count - 1
                If InStr(address, ComboBox10.Items(i)) > 0 Then
                    ComboBox10.SelectedIndex = i
                End If
            Next
            address = Replace(address, ComboBox10.SelectedItem, "")

            For i As Integer = 0 To ComboBox9.Items.Count - 1
                If InStr(address, ComboBox9.Items(i)) > 0 Then
                    ComboBox9.SelectedIndex = i
                End If
            Next
            address = Replace(address, ComboBox9.SelectedItem, "")

            For i As Integer = 0 To ComboBox8.Items.Count - 1
                If InStr(address, ComboBox8.Items(i)) > 0 Then
                    ComboBox8.SelectedIndex = i
                End If
            Next
        End If
    End Sub

    '郵便番号データ処理
    Private Sub Button26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button26.Click
        If TextBox30.Text = "" Then
            Exit Sub
        End If

        Dim str As String() = New String() {"-", " ", "　"}
        For i As Integer = 0 To str.Length - 1
            TextBox30.Text = Replace(TextBox30.Text, str(i), "")
        Next

        Dim whereCheck As String = ""
        If TextBox30.Text <> "" Then
            whereCheck = " ([郵便番号] = '" & TextBox30.Text & "')"
        End If

        '=======================================================
        Dim sql As String = "SELECT * FROM [X-ken-all2] WHERE" & whereCheck
        Dim table As DataTable = TM_DB_CONNECT_SELECT_SQL(CnAccdb_ken_all, sql)
        '=======================================================

        If table.Rows.Count <> 0 Then
            ComboBox10.SelectedItem = table.Rows(0)("県名").ToString
            ComboBox9.SelectedItem = table.Rows(0)("市町村名").ToString
            ComboBox8.SelectedItem = table.Rows(0)("町名").ToString
            Label48.Text = table.Rows(0)("市町村名読み").ToString
            Label49.Text = table.Rows(0)("町名読み").ToString
        Else
            MsgBox("見つかりません", MsgBoxStyle.SystemModal)
        End If
    End Sub


    Private Sub ComboBox10_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox10.SelectedIndexChanged
        Dim whereCheck As String = ""
        If ComboBox10.SelectedItem <> "" Then
            whereCheck = " ([県名] = '" & ComboBox10.SelectedItem & "')"
        End If

        '=======================================================
        Dim sql As String = "SELECT * FROM [X-ken-all2] WHERE" & whereCheck
        Dim table As DataTable = TM_DB_CONNECT_SELECT_SQL(CnAccdb_ken_all, sql)
        '=======================================================

        ComboBox9.Items.Clear()
        sichoYomi.Clear()
        For r As Integer = 0 To table.Rows.Count - 1
            Dim str As String = table.Rows(r)("市町村名").ToString
            If str <> "" Then
                If Not ComboBox9.Items.Contains(str) Then
                    ComboBox9.Items.Add(table.Rows(r)("市町村名").ToString)
                    sichoYomi.Add(table.Rows(r)("市町村名読み").ToString)
                End If
            End If
        Next

        ComboBox9.SelectedIndex = 0
        Label48.Text = sichoYomi.Item(ComboBox9.SelectedIndex)
    End Sub

    Private Sub ComboBox9_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox9.SelectedIndexChanged
        Dim whereCheck As String = ""
        If ComboBox10.SelectedItem <> "" Then
            whereCheck = " ([県名] = '" & ComboBox10.SelectedItem & "')"
        End If
        If ComboBox9.SelectedItem <> "" Then
            whereCheck &= " AND ([市町村名] = '" & ComboBox9.SelectedItem & "')"
        End If

        '=======================================================
        Dim sql As String = "SELECT * FROM [X-ken-all2] WHERE" & whereCheck
        Dim table As DataTable = TM_DB_CONNECT_SELECT_SQL(CnAccdb_ken_all, sql)
        '=======================================================

        ComboBox8.Items.Clear()
        choYomi.Clear()
        For r As Integer = 0 To table.Rows.Count - 1
            Dim str As String = table.Rows(r)("町名").ToString
            If str <> "" Then
                If Not ComboBox8.Items.Contains(str) Then
                    ComboBox8.Items.Add(table.Rows(r)("町名").ToString)
                    choYomi.Add(table.Rows(r)("町名読み").ToString)
                End If
            End If
        Next

        Label48.Text = sichoYomi.Item(ComboBox9.SelectedIndex)
        If choYomi.Count > 0 Then
            Label49.Text = choYomi.Item(0)
            ComboBox8.SelectedIndex = 0
        Else
            Label49.Text = ""
            ComboBox8.Text = ""
        End If
    End Sub

    Private Sub ComboBox8_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox8.SelectedIndexChanged
        Label49.Text = choYomi.Item(ComboBox8.SelectedIndex)
    End Sub

    Private Sub Button34_Click(sender As Object, e As EventArgs) Handles Button34.Click
        EigyousyoSearch(1)
    End Sub

    Private Sub Button29_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button29.Click
        EigyousyoSearch(0)
    End Sub

    Dim kyokuArray As ArrayList = New ArrayList
    Private Sub EigyousyoSearch(ByVal mode As Integer)
        Dim tableName As String = ""
        Dim headerKyoku As String = ""
        Select Case mode
            Case 0
                tableName = "郵便局一覧"
                headerKyoku = "局名"
            Case 1
                tableName = "佐川営業所一覧"
                headerKyoku = "営業所名"
        End Select

        Dim searchStr As String = ""
        Dim whereCheck As String = ""

        If TextBox28.Text = "" Then
            If CheckBox19.Checked = True Then
                searchStr = ComboBox9.SelectedItem & ComboBox8.SelectedItem
            Else
                searchStr = ComboBox9.SelectedItem
            End If
            whereCheck = " ([住所] like '%" & searchStr & "%')"
        Else
            searchStr = TextBox28.Text
            whereCheck = " ([" & headerKyoku & "] like '%" & searchStr & "%')"
        End If

        '=======================================================
        TextBox23.Text = "SELECT * FROM [" & tableName & "] WHERE" & whereCheck
        Dim sql As String = TextBox23.Text
        Dim table As DataTable = TM_DB_CONNECT_SELECT_SQL(CnAccdb_ken_all, sql)
        '=======================================================

        ListBox2.Items.Clear()
        kyokuArray.Clear()
        For r As Integer = 0 To table.Rows.Count - 1
            Dim str As String = table.Rows(r)(headerKyoku).ToString
            If str <> "" Then
                If Not ListBox2.Items.Contains(str) Then
                    ListBox2.Items.Add(table.Rows(r)(headerKyoku).ToString)
                    Select Case mode
                        Case 0
                            kyokuArray.Add(table.Rows(r)(headerKyoku) & "," & table.Rows(r)("郵便番号") & "," & table.Rows(r)("住所") & "," & table.Rows(r)("電話番号"))
                        Case 1
                            kyokuArray.Add(table.Rows(r)(headerKyoku) & "," & table.Rows(r)("郵便番号") & "," & table.Rows(r)("住所") & "," & table.Rows(r)("電話番号") & "," & table.Rows(r)("番号") & "," & table.Rows(r)("ローカルコード"))
                    End Select
                End If
            End If
        Next
    End Sub

    Private Sub CheckBox19_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox19.CheckedChanged
        If CheckBox19.Checked = True Then
            ComboBox8.Enabled = True
        Else
            ComboBox8.Enabled = False
        End If
    End Sub

    Private Sub ListBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox2.SelectedIndexChanged
        If kyokuArray.Count > 0 Then
            For i As Integer = 0 To kyokuArray.Count - 1
                Dim res As String() = Split(kyokuArray(i), ",")
                If res(0) = ListBox2.SelectedItem Then
                    TextBox29.Text = res(1)
                    TextBox31.Text = res(2)
                    TextBox32.Text = res(3)
                    If res.Length > 4 Then
                        TextBox35.Text = res(4)
                        TextBox53.Text = res(5)
                    Else
                        TextBox35.Text = ""
                        TextBox53.Text = ""
                    End If
                    Exit Sub
                End If
            Next
        End If

        TextBox29.Text = ""
        TextBox31.Text = ""
        TextBox32.Text = ""
        TextBox35.Text = ""
        TextBox53.Text = ""
    End Sub

    '局留め適用
    '郵便局
    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click
        If DGV1.RowCount > 0 And DGV1.SelectedCells.Count > 0 Then
            'datagridviewのヘッダーテキストをコレクションに取り込む
            Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)

            For r As Integer = 0 To DGV1.SelectedCells.Count - 1
                Dim selRow As Integer = DGV1.SelectedCells(r).RowIndex
                DGV1.Item(dH1.IndexOf("営業所留め"), selRow).Value = "1"
                If ListBox2.SelectedItems.Count > 0 Then
                    If InStr(ListBox2.SelectedItem, "郵便局") = 0 Then
                        DGV1.Item(dH1.IndexOf("営業所コード"), selRow).Value = ListBox2.SelectedItem & "郵便局"
                    Else
                        DGV1.Item(dH1.IndexOf("営業所コード"), selRow).Value = ListBox2.SelectedItem
                    End If
                    DGV1.Item(dH1.IndexOf("ローカルコード"), selRow).Value = Replace(TextBox29.Text, "-", "")
                Else
                    DGV1.Item(dH1.IndexOf("営業所コード"), selRow).Value = TextBox35.Text
                End If
            Next

            '履歴保存
            If CellValue <> DGV1.CurrentCell.Value Then
                DGV1.CurrentCell.Style.BackColor = Color.Yellow
                ValidateChangeSave(DGV1, DGV1.SelectedCells(0).ColumnIndex, DGV1.SelectedCells(0).RowIndex)
            End If

            MsgBox("登録しました", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
        End If
    End Sub

    '佐川
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If DGV1.RowCount > 0 And DGV1.SelectedCells.Count > 0 Then
            'datagridviewのヘッダーテキストをコレクションに取り込む
            Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)

            For r As Integer = 0 To DGV1.SelectedCells.Count - 1
                Dim selRow As Integer = DGV1.SelectedCells(r).RowIndex
                DGV1.Item(dH1.IndexOf("営業所留め"), selRow).Value = "1"
                If ListBox2.SelectedItems.Count > 0 Then
                    DGV1.Item(dH1.IndexOf("営業所コード"), selRow).Value = TextBox35.Text
                    DGV1.Item(dH1.IndexOf("ローカルコード"), selRow).Value = TextBox53.Text
                Else
                    MsgBox("営業所を選択してください")
                    Exit Sub
                End If
            Next

            '履歴保存
            If CellValue <> DGV1.CurrentCell.Value Then
                DGV1.CurrentCell.Style.BackColor = Color.Yellow
                ValidateChangeSave(DGV1, DGV1.SelectedCells(0).ColumnIndex, DGV1.SelectedCells(0).RowIndex)
            End If

            MsgBox("登録しました", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
        End If
    End Sub


    Private Sub Button30_Click(sender As Object, e As EventArgs) Handles Button30.Click
        If DGV1.RowCount > 0 And DGV1.SelectedCells.Count > 0 Then
            'datagridviewのヘッダーテキストをコレクションに取り込む
            Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)

            For r As Integer = 0 To DGV1.SelectedCells.Count - 1
                Dim selRow As Integer = DGV1.SelectedCells(r).RowIndex
                DGV1.Item(dH1.IndexOf("営業所留め"), selRow).Value = "0"
                DGV1.Item(dH1.IndexOf("営業所コード"), selRow).Value = ""
                DGV1.Item(dH1.IndexOf("ローカルコード"), selRow).Value = ""
            Next

            '履歴保存
            If CellValue <> DGV1.CurrentCell.Value Then
                DGV1.CurrentCell.Style.BackColor = Color.Yellow
                ValidateChangeSave(DGV1, DGV1.SelectedCells(0).ColumnIndex, DGV1.SelectedCells(0).RowIndex)
            End If

        End If
    End Sub


    '********************************************
    '梱包計算
    '********************************************
    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        If TextBox18.Text = "" Or ComboBox11.Text = "" Then
            Exit Sub
        End If

        DGV2.Rows.Clear()
        Dim dH2 As ArrayList = TM_HEADER_GET(DGV2)
        Dim codeArray As String() = Split(TextBox18.Text, vbCrLf)
        For i As Integer = 0 To codeArray.Length - 1
            If codeArray(i) = "" Then
                Continue For
            End If

            Dim line As String() = Split(codeArray(i), "*")
            If line.Length = 1 Then
                MsgBox("コードまたは注文数の書き方が間違っています")
                Exit Sub
            End If

            Dim checkCode As String = ""
            Dim codeRow As Integer = 0

            If dgv6CodeArray.Contains(line(0)) Then
                codeRow = dgv6CodeArray.IndexOf(line(0))
                checkCode = DGV6.Item(TM_ArIndexof(dH6, "商品コード"), codeRow).Value
            ElseIf dgv6DaihyoCodeArray.Contains(line(0)) Then
                '代表コードだけで指定している時は、最初の商品コードに変換する
                codeRow = dgv6DaihyoCodeArray.IndexOf(line(0))
                checkCode = DGV6.Item(TM_ArIndexof(dH6, "商品コード"), codeRow).Value
            Else
                Continue For
            End If

            DGV2.Rows.Add(1)
            Dim newRow As Integer = DGV2.RowCount - 1
            DGV2.Item(dH2.IndexOf("コード"), newRow).Value = checkCode
            DGV2.Item(dH2.IndexOf("注"), newRow).Value = line(1)
            DGV2.Item(dH2.IndexOf("SW"), newRow).Value = DGV6.Item(TM_ArIndexof(dH6, "ship-weight"), codeRow).Value
            DGV2.Item(dH2.IndexOf("タグ"), newRow).Value = DGV6.Item(TM_ArIndexof(dH6, "商品分類タグ"), codeRow).Value
            DGV2.Item(dH2.IndexOf("商品名"), newRow).Value = DGV6.Item(TM_ArIndexof(dH6, "商品名"), codeRow).Value
            DGV2.Item(dH2.IndexOf("ロケーション"), newRow).Value = DGV6.Item(TM_ArIndexof(dH6, "ロケーション"), codeRow).Value
            DGV2.Item(dH2.IndexOf("梱包サイズ"), newRow).Value = DGV6.Item(TM_ArIndexof(dH6, "梱包サイズ"), codeRow).Value
            DGV2.Item(dH2.IndexOf("特殊"), newRow).Value = DGV6.Item(TM_ArIndexof(dH6, "特殊"), codeRow).Value
            DGV2.Item(dH2.IndexOf("sw2"), newRow).Value = DGV6.Item(TM_ArIndexof(dH6, "sw2"), codeRow).Value

            'Exit For
        Next

        Dim codesArray As New ArrayList
        Dim swArray As New ArrayList
        Dim tokuArray As New ArrayList
        For r As Integer = 0 To DGV2.RowCount - 1
            For i As Integer = 1 To DGV2.Item(dH2.IndexOf("注"), r).Value
                codesArray.Add(DGV2.Item(dH2.IndexOf("コード"), r).Value)
                swArray.Add(DGV2.Item(dH2.IndexOf("SW"), r).Value)
                tokuArray.Add(DGV2.Item(dH2.IndexOf("特殊"), r).Value)
            Next
        Next
        Dim res As String = Keisan(ComboBox11.SelectedItem, codesArray, swArray, tokuArray)
        ComboBox11.SelectedItem = Split(res, ",")(1)
        TextBox10.Text = Split(res, ",")(0)
        TextBox9.Text = Split(res, ",")(2)
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        ComboBox11.SelectedIndex = 0
        TextBox10.Text = ""
        TextBox9.Text = ""
        TextBox18.Text = ""
    End Sub

    Private Function Keisan(haisouKind As String, codesArray As ArrayList, swArray As ArrayList, tokuArray As ArrayList)
        '重さ、特殊の初期化
        omosa = 0
        tokuGroup.Clear()

        Dim maeKind As String = ""
        Dim haisouSize As Double = 0
        Dim sizeName As String = ""
        Dim haisouSizeArray As New ArrayList    'バラで保存

        Dim mTag As String = ""
        Dim mItemName As String = ""
        For i As Integer = 0 To swArray.Count - 1
            If haisouKind = "" Then
                If Regex.IsMatch(swArray(i), "P|p") Then
                    haisouKind = "宅配便"
                ElseIf Regex.IsMatch(swArray(i), "M|m") Then
                    haisouKind = "メール便"
                ElseIf Regex.IsMatch(swArray(i), "T|t") Then
                    haisouKind = "定形外"
                End If
            End If

            'デフォルトがメール便で、宅配便があったら、haisouKindを宅配便にする
            'メール便と定形外の場合は、宅配便にする
            If (haisouKind = "メール便" And Regex.IsMatch(swArray(i), "P|p|T|t")) Or (haisouKind = "定形外" And Regex.IsMatch(swArray(i), "P|p")) Then
                maeKind = haisouKind
                haisouKind = "宅配便"
                Select Case maeKind
                    Case "メール便"
                        haisouSize = haisouSize / 100   'メール便サイズ計算を宅配便に変更
                    Case "定形外"
                        haisouSize = haisouSize / 75   '定形外サイズ計算を宅配便に変更
                End Select
            End If

            'haisouKindが宅配便の時、メール便を1/100にする
            Dim henkouFlag As Boolean = False
            If haisouKind = "宅配便" And Regex.IsMatch(swArray(i), "M|m|T|t") Then
                henkouFlag = True
            End If

            Dim weight As String = Regex.Replace(swArray(i), "P|p|M|m|T|t", "")
            Dim w2 As String = Regex.Match(weight, "\(.*\)").Value
            w2 = Regex.Replace(w2, "\(|\)", "")
            weight = Regex.Replace(weight, "\(.*\)", "")

            'バラで保存（とりあえず定形外用で使用）
            haisouSizeArray.Add(weight)

            '重さ・特殊の処理
            If InStr(tokuArray(i), "重") > 0 Then
                Dim omotoku As Double = OmosaCheck(codesArray(i).ToLower, 1, haisouSize)
                If omotoku <> 0 Then
                    haisouSize = haisouSize + omotoku
                End If
            ElseIf InStr(tokuArray(i), "特") > 0 Then
                Dim omotoku As Double = TokusyuCheck(codesArray(i).ToLower, 1, haisouSize)
                If omotoku <> 0 Then
                    haisouSize = haisouSize + omotoku
                End If
            Else
                '-------
                If haisouKind = "メール便" Then
                    haisouSize = haisouSize + (CDbl(weight) * CDbl(1))
                ElseIf haisouKind = "定形外" Then
                    haisouSize = haisouSize + (CDbl(weight) * CDbl(1))
                Else
                    If henkouFlag Then
                        haisouSize = haisouSize + ((CDbl(weight) / 100) * CDbl(1))
                    Else
                        If w2 <> "" Then    'P160(98)、P100(3.3)は1個のパーセンテージ（優先）
                            haisouSize = haisouSize + (CDbl(w2) * CDbl(1))
                        Else
                            haisouSize = haisouSize + (CDbl(TakuhaiPerConv(weight)) * CDbl(1))
                        End If
                    End If
                End If
                '-------
            End If

            'メール便指定数以上は宅配便に変更
            If haisouKind = "メール便" Then
                If haisouSize > NumericUpDown4.Value * 100 Then
                    haisouKind = "宅配便"
                    haisouSize = haisouSize / 100   'メール便サイズ計算を宅配便に変更
                End If
            ElseIf haisouKind = "定形外" Then
                If haisouSize > NumericUpDown5.Value * NumericUpDown6.Value Then
                    haisouKind = "宅配便"
                    haisouSize = haisouSize / 75   'メール便サイズ計算を宅配便に変更
                End If
            End If
        Next

        Dim res As String = ""

        '便数計算（メール便と宅配便は100％で1便、定形外は250gで1便）
        If Regex.IsMatch(haisouKind, "メール便|宅配便") Then
            res = Math.Ceiling(haisouSize / 100)
        Else
            '「最大重量 / 2 > 最大数」なら宅配便
            Dim counter As Integer = 0
            For i As Integer = 0 To haisouSizeArray.Count - 1
                If haisouSizeArray(i) > (NumericUpDown6.Value / 2) Then
                    counter += 1
                End If
            Next
            If counter > NumericUpDown5.Value Then
                haisouKind = "宅配便"
                haisouSize = haisouSize * 0.01
                res = Math.Ceiling(haisouSize / 100)
            Else
                res = Math.Ceiling(haisouSize / NumericUpDown6.Value)
            End If
        End If

        res &= "," & haisouKind
        res &= "," & Math.Ceiling(haisouSize)

        Return res
    End Function


    '*********************
    'エラー処理
    '*********************
    Private Sub KryptonButton2_Click(sender As Object, e As EventArgs) Handles KryptonButton2.Click
        Dim counter As Integer = 0
        For i As Integer = 0 To ListBox1.Items.Count - 1
            Dim res1 As String() = Split(ListBox1.Items(i), " : ")
            Dim res2 As String() = Split(res1(0), "/")
            Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
            If res1(1) = "時間指定-不可地域" Then
                DGV1.Item(dH1.IndexOf("配送日"), CInt(res2(1)) - 1).Value = ""
                DGV1.Item(dH1.IndexOf("配送日"), CInt(res2(1)) - 1).Style.BackColor = Color.Aqua
                ValidateChangeSave(DGV1, dH1.IndexOf("配送日"), CInt(res2(1)) - 1)
                DGV1.Item(dH1.IndexOf("配送時間帯"), CInt(res2(1)) - 1).Value = ""
                DGV1.Item(dH1.IndexOf("配送時間帯"), CInt(res2(1)) - 1).Style.BackColor = Color.Aqua
                ValidateChangeSave(DGV1, dH1.IndexOf("配送時間帯"), CInt(res2(1)) - 1)
                counter += 1
            End If
        Next
        MsgBox(counter & "件修正しました")
    End Sub

    Private Sub KryptonButton6_Click(sender As Object, e As EventArgs) Handles KryptonButton6.Click
        Dim counter As Integer = 0
        For i As Integer = 0 To ListBox1.Items.Count - 1
            Dim res1 As String() = Split(ListBox1.Items(i), " : ")
            Dim res2 As String() = Split(res1(0), "/")
            Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
            If res1(1) = "日付指定が現在日より前" Then
                DGV1.Item(dH1.IndexOf("配送日"), CInt(res2(1)) - 1).Value = ""
                DGV1.Item(dH1.IndexOf("配送日"), CInt(res2(1)) - 1).Style.BackColor = Color.Aqua
                ValidateChangeSave(DGV1, dH1.IndexOf("配送日"), CInt(res2(1)) - 1)
                counter += 1
            End If
        Next
        MsgBox(counter & "件修正しました")
    End Sub

    Private Sub KryptonButton4_Click(sender As Object, e As EventArgs) Handles KryptonButton4.Click
        Dim counter As Integer = 0
        For i As Integer = 0 To ListBox1.Items.Count - 1
            Dim res1 As String() = Split(ListBox1.Items(i), " : ")
            Dim res2 As String() = Split(res1(0), "/")
            Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
            If res1(1) = "配送（メール便が宅配便）確認" Or res1(1) = "配送（宅配便がメール便）確認" Then
                DGV1.Item(dH1.IndexOf("マスタ配送"), CInt(res2(1)) - 1).Value = DGV1.Item(dH1.IndexOf("発送方法"), CInt(res2(1)) - 1).Value
                DGV1.Item(dH1.IndexOf("マスタ配送"), CInt(res2(1)) - 1).Style.BackColor = Color.Aqua
                DGV1.Item(dH1.IndexOf("発送方法"), CInt(res2(1)) - 1).Style.BackColor = Color.Aqua
                ValidateChangeSave(DGV1, dH1.IndexOf("マスタ配送"), CInt(res2(1)) - 1)
                counter += 1
            End If
        Next
        MsgBox(counter & "件修正しました")
    End Sub

    Private Sub KryptonButton5_Click(sender As Object, e As EventArgs) Handles KryptonButton5.Click
        Dim counter As Integer = 0
        For i As Integer = 0 To ListBox1.Items.Count - 1
            Dim res1 As String() = Split(ListBox1.Items(i), " : ")
            Dim res2 As String() = Split(res1(0), "/")
            Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
            If res1(1) = "配送（メール便が宅配便）確認" Or res1(1) = "配送（宅配便がメール便）確認" Then
                DGV1.Item(dH1.IndexOf("発送方法"), CInt(res2(1)) - 1).Value = DGV1.Item(dH1.IndexOf("マスタ配送"), CInt(res2(1)) - 1).Value
                DGV1.Item(dH1.IndexOf("マスタ配送"), CInt(res2(1)) - 1).Style.BackColor = Color.Aqua
                DGV1.Item(dH1.IndexOf("発送方法"), CInt(res2(1)) - 1).Style.BackColor = Color.Aqua
                ValidateChangeSave(DGV1, dH1.IndexOf("発送方法"), CInt(res2(1)) - 1)
                counter += 1
            End If
        Next
        MsgBox(counter & "件修正しました")
    End Sub

    Private Sub DataGridView1_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles DGV1.CellValidating
        Dim dgv As DataGridView = DirectCast(sender, DataGridView)
        '新しい行のセルでなく、セルの内容が変更されている時だけ検証する 
        If e.RowIndex = dgv.NewRowIndex OrElse Not dgv.IsCurrentCellDirty Then
            Exit Sub
        End If

        Dim col As Integer = e.ColumnIndex
        Dim row As Integer = e.RowIndex

        ValidateChangeSave(dgv, col, row)
    End Sub

    '履歴
    Private Sub ValidateRowAdd(dNo As String, mesline As String)
        Dim mesArray As String() = Split(mesline, " : ")
        For r As Integer = 0 To DGV14.RowCount - 1
            If dNo = DGV14.Item(0, r).Value And mesArray(1) = DGV14.Item(2, r).Value Then
                Exit Sub
            End If
        Next
        If InStr(mesline, " : ") > 0 Then
            Dim mes As String() = Split(mesline, " : ")
            DGV14.Rows.Add(dNo, mes(0), mes(1))
        Else
            DGV14.Rows.Add(dNo, "", mesline)
        End If
    End Sub

    Private Sub ValidateChangeSave(dgv As DataGridView, col As Integer, row As Integer)
        Dim counter As Integer = 0
        Dim str As String = ""
        If CStr(dgv.Item(col, row).Value) = "" Then
            str = "NULL"
        Else
            str = dgv.Item(col, row).Value
        End If

        TextBox39.Text = "変更"
        TextBox39.BackColor = Color.Yellow

        For r As Integer = 0 To DGV14.RowCount - 1
            If DGV14.Item(1, r).Value <> "" Then
                Dim res2 As String() = Split(DGV14.Item(1, r).Value, "/")
                If col = res2(0) And row = res2(1) - 1 Then
                    DGV14.Item(2, r).Value = str
                    counter += 1
                End If
            End If
        Next

        If counter = 0 Then
            Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
            Dim mes1 As String = col & "/" & row
            Dim mes2 As String = "手動変更"
            DGV14.Rows.Add(DGV1.Item(dH1.IndexOf("伝票番号"), row).Value, mes1, mes2, str)
        End If
    End Sub

    Private Sub DataGridView14_SelectionChanged(sender As Object, e As EventArgs) Handles DGV14.SelectionChanged
        If DGV14.RowCount > 0 And DGV14.SelectedCells.Count > 0 Then
            TextBox49.Text = DGV14.SelectedCells(0).Value
        End If
    End Sub


    '*********************
    '伝票確認
    '*********************
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        If DGV1.RowCount = 0 Then
            Exit Sub
        End If

        CheckBox12.Checked = False
        DGV4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
        Dim dH4 As ArrayList = TM_HEADER_GET(DGV4)
        Dim dgv1dArray As New ArrayList
        For r As Integer = 0 To DGV1.RowCount - 1
            Dim d1No As String = DGV1.Item(dH1.IndexOf("伝票番号"), r).Value
            dgv1dArray.Add(d1No)
        Next

        '重複チェック
        Dim dgv4dArray As New ArrayList
        For r As Integer = 0 To DGV4.RowCount - 1
            Dim d4No As String = DGV4.Item(dH4.IndexOf("伝票番号"), r).Value
            If dgv4dArray.Contains(d4No) Then
                DGV4.Item(dH4.IndexOf("check"), r).Value = "重"
                DGV4.Item(dH4.IndexOf("check"), r).Style.BackColor = Color.Orange
            Else
                dgv4dArray.Add(d4No)
            End If
        Next

        '存在チェック
        For r As Integer = 0 To DGV4.RowCount - 1
            If DGV4.Item(dH4.IndexOf("check"), r).Value <> "" Then
                Continue For
            End If

            DGV4.CurrentCell = DGV4(0, r)
            DGV4.Item(dH4.IndexOf("check"), r).Value = ""
            DGV4.Item(dH4.IndexOf("check"), r).Style.BackColor = Color.Empty

            If r Mod 5 = 0 Then
                Application.DoEvents()
            End If

            Dim d4No As String = DGV4.Item(dH4.IndexOf("伝票番号"), r).Value
            If dgv1dArray.Contains(d4No) Then
                DGV4.Item(dH4.IndexOf("check"), r).Value = "○"
            End If
            If DGV4.Item(dH4.IndexOf("check"), r).Value <> "○" Then
                DGV4.Item(dH4.IndexOf("check"), r).Value = "×"
                DGV4.Item(dH4.IndexOf("check"), r).Style.BackColor = Color.Yellow
            End If
        Next

        DGV4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Application.DoEvents()
        MsgBox("検索終了しました")
    End Sub

    Private Sub CheckBox12_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox12.CheckedChanged
        DGV4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        If CheckBox12.Checked Then
            Dim dH4 As ArrayList = TM_HEADER_GET(DGV4)
            For r As Integer = 0 To DGV4.RowCount - 1
                If Not Regex.IsMatch(DGV4.Item(dH4.IndexOf("check"), r).Value, "×|重") Then
                    DGV4.Rows(r).Visible = False
                End If
            Next
        Else
            For r As Integer = 0 To DGV4.RowCount - 1
                DGV4.Rows(r).Visible = True
            Next
        End If
        DGV4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
    End Sub





    '============================================
    '--------------------------------------------
    '件数表示リンクラベル
    '--------------------------------------------
    '============================================
    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim checkArray As String() = New String() {"便種（スピードを選択）=000", "処理用2=" & HS1.Text}
        FilterDGV(DGV7, checkArray, "佐川e飛伝2")
    End Sub

    ''' <summary>
    ''' DataGridViewフィルタ
    ''' </summary>
    ''' <param name="dgv">DataGridView</param>
    ''' <param name="checkArray">配列</param>
    Private Sub FilterDGV(dgv As DataGridView, checkArray As String(), tabName As String)
        selChangeFlag = False
        Dim dH As ArrayList = TM_HEADER_GET(dgv)

        If tabName = "ヤマト(陸)(太宰府)" Then
            Dim flag As Boolean = True
            For r As Integer = 0 To dgv.RowCount - 1
                If dgv.Item(dH.IndexOf("マスタ配送"), r).Value = "ヤマト(陸便)" And dgv.Item(dH.IndexOf("処理用2"), r).Value = "太宰府" Then
                    flag = True
                Else
                    flag = False
                End If
                dgv.Rows(r).Visible = flag
            Next
            TabControl2.SelectedIndex = 3
        ElseIf tabName = "ヤマト(陸)(井相田)" Then
            Dim flag As Boolean = True
            For r As Integer = 0 To dgv.RowCount - 1
                If dgv.Item(dH.IndexOf("マスタ配送"), r).Value = "ヤマト(陸便)" And dgv.Item(dH.IndexOf("処理用2"), r).Value = "井相田" Then
                    flag = True
                Else
                    flag = False
                End If
                dgv.Rows(r).Visible = flag
            Next
            TabControl2.SelectedIndex = 3
        ElseIf tabName = "ヤマト(船)(太宰府)" Then
            Dim flag As Boolean = True
            For r As Integer = 0 To dgv.RowCount - 1
                If dgv.Item(dH.IndexOf("マスタ配送"), r).Value = "ヤマト(船便)" And dgv.Item(dH.IndexOf("処理用2"), r).Value = "太宰府" Then
                    flag = True
                Else
                    flag = False
                End If
                dgv.Rows(r).Visible = flag
            Next
            TabControl2.SelectedIndex = 3
        ElseIf tabName = "ヤマト(船)(井相田)" Then
            Dim flag As Boolean = True
            For r As Integer = 0 To dgv.RowCount - 1
                If dgv.Item(dH.IndexOf("マスタ配送"), r).Value = "ヤマト(船便)" And dgv.Item(dH.IndexOf("処理用2"), r).Value = "井相田" Then
                    flag = True
                Else
                    flag = False
                End If
                dgv.Rows(r).Visible = flag
            Next
            TabControl2.SelectedIndex = 3
        Else
            For r As Integer = 0 To dgv.RowCount - 1
                Dim flag As Boolean = True
                For i As Integer = 0 To checkArray.Length - 1
                    Dim ca As String() = Split(checkArray(i), "=")
                    If dgv.Item(dH.IndexOf(ca(0)), r).Value = ca(1) Then
                        flag = True
                    Else
                        flag = False
                        Exit For
                    End If
                Next
                dgv.Rows(r).Visible = flag
            Next
            For i As Integer = 0 To TabControl2.TabPages.Count - 1
                If InStr(TabControl2.TabPages(i).Text, tabName) > 0 Then
                    TabControl2.SelectedIndex = i
                    Exit For
                End If
            Next
        End If

        selChangeFlag = True
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Dim checkArray As String() = New String() {"便種（スピードを選択）=003", "処理用2=" & HS1.Text}
        FilterDGV(DGV7, checkArray, "佐川e飛伝2")
    End Sub

    Private Sub LinkLabel5_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel5.LinkClicked
        Dim checkArray As String() = New String() {"処理用2=" & HS1.Text}
        FilterDGV(DGV9, checkArray, "メール便")
    End Sub

    Private Sub LinkLabel6_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel6.LinkClicked
        Dim checkArray As String() = New String() {"処理用2=" & HS1.Text}
        FilterDGV(DGV13, checkArray, "定形外")
    End Sub

    Private Sub LinkLabel29_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel29.LinkClicked
        Dim checkArray As String() = New String() {"処理用2=" & HS1.Text}
        FilterDGV(DGV9, checkArray, "ヤマト(陸)(太宰府)")
    End Sub

    Private Sub LinkLabel30_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel30.LinkClicked
        Dim checkArray As String() = New String() {"処理用2=" & HS2.Text}
        FilterDGV(DGV9, checkArray, "ヤマト(陸)(井相田)")
    End Sub

    Private Sub LinkLabel34_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel34.LinkClicked
        Dim checkArray As String() = New String() {"処理用2=" & HS1.Text}
        FilterDGV(DGV9, checkArray, "ヤマト(船)(太宰府)")
    End Sub

    Private Sub LinkLabel35_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel35.LinkClicked
        Dim checkArray As String() = New String() {"処理用2=" & HS2.Text}
        FilterDGV(DGV9, checkArray, "ヤマト(船)(井相田)")
    End Sub

    Private Sub LinkLabel16_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel16.LinkClicked
        Dim checkArray As String() = New String() {"便種（スピードを選択）=000", "処理用2=" & HS2.Text}
        FilterDGV(DGV7, checkArray, "佐川e飛伝2")
    End Sub

    Private Sub LinkLabel15_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel15.LinkClicked
        Dim checkArray As String() = New String() {"便種（スピードを選択）=003", "処理用2=" & HS2.Text}
        FilterDGV(DGV7, checkArray, "佐川e飛伝2")
    End Sub

    Private Sub LinkLabel3_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        Dim checkArray As String() = New String() {"便種コード=000", "処理用2=" & HS2.Text}
        FilterDGV(DGV8, checkArray, "BIZlogi")
    End Sub

    Private Sub LinkLabel4_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel4.LinkClicked
        Dim checkArray As String() = New String() {"便種コード=030", "処理用2=" & HS2.Text}
        FilterDGV(DGV8, checkArray, "BIZlogi")
    End Sub

    Private Sub Linklabel17_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel17.LinkClicked
        Dim checkArray As String() = New String() {"処理用2=" & HS2.Text}
        FilterDGV(DGV9, checkArray, "メール便")
    End Sub

    Private Sub Linklabel18_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel18.LinkClicked
        Dim checkArray As String() = New String() {"処理用2=" & HS2.Text}
        FilterDGV(DGV13, checkArray, "定形外")
    End Sub

    Private Sub Linklabel19_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
        Dim checkArray As String() = New String() {"便種コード=000", "処理用2=" & HS3.Text}
        FilterDGV(DGV8, checkArray, "BIZlogi")
    End Sub

    Private Sub LinkLabel20_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
        Dim checkArray As String() = New String() {"便種コード=030", "処理用2=" & HS3.Text}
        FilterDGV(DGV8, checkArray, "BIZlogi")
    End Sub

    Private Sub LinkLabel21_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel21.LinkClicked
        Dim checkArray As String() = New String() {"便種コード=000", "処理用2=" & HS1.Text}
        FilterDGV(DGV8, checkArray, "BIZlogi")
    End Sub

    Private Sub LinkLabel22_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel22.LinkClicked
        Dim checkArray As String() = New String() {"便種コード=030", "処理用2=" & HS1.Text}
        FilterDGV(DGV8, checkArray, "BIZlogi")
    End Sub

    Private Sub LinkLabel25_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel25.LinkClicked
        Dim checkArray As String() = New String() {"便種（スピードを選択）=000", "処理用2=" & HS4.Text}
        FilterDGV(DGV7, checkArray, "佐川e飛伝2")
    End Sub

    Private Sub LinkLabel26_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel26.LinkClicked
        Dim checkArray As String() = New String() {"便種（スピードを選択）=003", "処理用2=" & HS4.Text}
        FilterDGV(DGV7, checkArray, "佐川e飛伝2")
    End Sub

    Private Sub LinkLabel23_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel23.LinkClicked
        Dim checkArray As String() = New String() {"便種コード=000", "処理用2=" & HS4.Text}
        FilterDGV(DGV8, checkArray, "BIZlogi")
    End Sub

    Private Sub LinkLabel24_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel24.LinkClicked
        Dim checkArray As String() = New String() {"便種コード=030", "処理用2=" & HS4.Text}
        FilterDGV(DGV8, checkArray, "BIZlogi")
    End Sub

    Private Sub LinkLabel20_LinkClicked_1(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel20.LinkClicked
        Dim checkArray As String() = New String() {"処理用2=" & HS4.Text}
        FilterDGV(DGV9, checkArray, "メール便")
    End Sub

    Private Sub LinkLabel19_LinkClicked_1(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel19.LinkClicked
        Dim checkArray As String() = New String() {"処理用2=" & HS4.Text}
        FilterDGV(DGV13, checkArray, "定形外")
    End Sub

    'Error
    Private Sub LinkLabel14_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel14.LinkClicked
        For r As Integer = 0 To DGV1.RowCount - 1
            Dim flag As Boolean = False
            For c As Integer = 0 To DGV1.ColumnCount - 1
                If DGV1.Item(c, r).Style.BackColor = Color.Red Then
                    DGV1.Rows(r).Visible = True
                    flag = True
                    Exit For
                End If
            Next
            If flag = False Then
                DGV1.Rows(r).Visible = False
            End If
        Next

        TabControl3.SelectTab("TabPage22")
        TabControl9.SelectTab("TabPage30")
    End Sub





    '============================================
    '--------------------------------------------
    '設定
    '--------------------------------------------
    '============================================
    Private Sub CheckBox16_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox16.CheckedChanged
        If CheckBox16.Checked = True Then
            KryptonButton3.Enabled = True
            KryptonButton1.Enabled = True
            If LockFlag Then    'ロック中は処理できないようにする
                KryptonButton3.Enabled = False
                KryptonButton1.Enabled = False
            End If
            Timer1.Enabled = False
        Else
            KryptonButton3.Enabled = False
            KryptonButton1.Enabled = False
            Timer1.Enabled = True
        End If
    End Sub

    Dim readPath As String = Path.GetDirectoryName(Form1.appPath)
    Dim readEnc As Encoding = ENC_SJ
    Private Sub KryptonComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles KryptonComboBox1.SelectedIndexChanged
        Select Case KryptonComboBox1.SelectedItem
            Case "航空便可不可"
                readPath = Path.GetDirectoryName(Form1.appPath) & "\config\version2\" & KryptonComboBox1.SelectedItem & ".txt"
                readEnc = ENC_SJ
            Case "同梱特殊"
                readPath = Path.GetDirectoryName(Form1.appPath) & "\config\version2\" & KryptonComboBox1.SelectedItem & ".txt"
                readEnc = ENC_SJ
        End Select
        AzukiControl2.Text = File.ReadAllText(readPath, readEnc)
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        File.WriteAllText(readPath, AzukiControl2.Text, readEnc)
        MsgBox("保存しました", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
    End Sub




    '============================================
    '--------------------------------------------
    '変換処理
    '--------------------------------------------
    '============================================
    'ヘッダー+"-"+ヘッダー,空欄（"は'でもOK）
    '[固定],文字列
    '[規定],今日、明日、日付=20170220,フォーマット（省略時 yyyyMMdd）
    '[変換]ヘッダー,search=word|search=word
    '[検索],ヘッダー=search/ヘッダー
    '[検索],ヘッダー=search/[固定]=TrueWord|FalseWord
    '　　　↑search 文字、>0
    '[分割]ヘッダー,splitword=number,
    '　　　↑number=last最後
    '[選択],ヘッダー1|ヘッダー2|ヘッダー3,notWord
    '[特殊]ヘッダー,[Qoo10日]
    '[特殊]ヘッダー,[Qoo10時]
    '[特殊]ヘッダー,[住所]1or2
    '[特殊]ヘッダー,[発送備考欄個口]
    '[特殊]企業名,[依頼主]
    '[特殊]企業名,[電話]000-000-0000
    '[合成]商品,

    Dim adPatternArray As New ArrayList
    Dim adPatternStr As String = ""

    Dim Template1 As String() = Nothing
    Dim Template2 As String() = Nothing
    Dim Template3 As String() = Nothing
    Dim Template4 As String() = Nothing
    Dim Template5 As String() = Nothing
    Private Sub KryptonButton3_Click(sender As Object, e As EventArgs) Handles KryptonButton3.Click
        If DGV6.RowCount = 0 Then
            MsgBox("マスタ確認中です。ファイルサーバーが繋がっていない場合は接続してください。", MsgBoxStyle.SystemModal)
            Exit Sub
        ElseIf Panel8.Visible = True Then
            MsgBox("配送情報ダウンロードのデータを入れないと動作しません。", MsgBoxStyle.SystemModal)
            Exit Sub
        ElseIf Panel3.Visible = True Then  '明細必須化
            MsgBox("明細を入れないと動作しません。", MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        'Dim ret As Integer = MessageBox.Show("作成したPDFを表示しますか？", "PDFの表示", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        'If ret = DialogResult.Yes Then
        '    Process.Start(newName)
        'Else
        '    MessageBox.Show("PDFファイルを保存しました")
        'End If

        'ヘッダー
        Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
        Dim dH3 As ArrayList = TM_HEADER_GET(DGV3)
        Dim dH9 As ArrayList = TM_HEADER_GET(DGV9)

        Label118.Text = 0
        Label119.Text = 0

        Dim fourtyup As Integer = 0
        For r As Integer = 0 To DGV1.RowCount - 1
            If DGV1.Item(dH1.IndexOf("マスタ便数"), r).Value > 40 Then
                DGV1.Item(dH1.IndexOf("マスタ便数"), r).Style.BackColor = Color.Red
                fourtyup += 1
            End If

            ''新任务
            'If DGV1.Item(dH1.IndexOf("マスタ便数"), r).Value > 4 Then
            '    DGV1.Item(dH1.IndexOf("処理用"), r).Value = "TMS"
            'End If

            'If DGV1.Item(dH1.IndexOf("支払方法"), r).Value = "代金引換" And DGV1.Item(dH1.IndexOf("発送倉庫"), r).Value = "名古屋" Then
            '    DGV1.Item(dH1.IndexOf("支払方法"), r).Style.BackColor = Color.Red
            'End If
        Next

        If fourtyup > 0 Then
            Dim ret As Integer = MessageBox.Show("口数が40以上のデータがあります、確認しましたか？", "確認しました", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If ret = DialogResult.Yes Then

            Else
                MessageBox.Show("処理を停止しました")
                Return
            End If
        End If

        'dgv1の伝票番号を取得保存
        dgv1dNoArray.Clear()
        Dim dgv1dNoArray_ch As New ArrayList
        For r As Integer = 0 To DGV1.RowCount - 1
            dgv1dNoArray.Add(DGV1.Item(dH1.IndexOf("伝票番号"), r).Value)
            If Not dgv1dNoArray_ch.Contains(DGV1.Item(dH1.IndexOf("伝票番号"), r).Value) Then
                dgv1dNoArray_ch.Add((DGV1.Item(dH1.IndexOf("伝票番号"), r).Value))
            End If
        Next

        If dgv1dNoArray.Count > dgv1dNoArray_ch.Count Then
            MsgBox("重複な伝票番号があります。", MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        For r As Integer = 0 To DGV1.RowCount - 1
            DGV1.Rows(r).Visible = True
        Next

        Label9.Visible = True
        TextBox39.Text = ""
        TextBox39.BackColor = Color.Empty
        TextBox48.Text = ""
        TextBox48.BackColor = Color.Empty


        '住所分割用パターン
        adPatternArray.Add("[0-9]{1,5}-[0-9]{1,5}-[0-9]{1,5}-[0-9]{1,5}")
        adPatternArray.Add("[0-9]{1,5}-[0-9]{1,5}-[0-9]{1,5}")
        adPatternArray.Add("[0-9]{1,5}-[0-9]{1,5}")
        adPatternArray.Add("-[0-9]{1,5}号")
        adPatternArray.Add("[0-9]{1,5}丁目[0-9]{1,5}[番|番地|-][0-9]{1,5}号*")
        adPatternArray.Add("[0-9]{1,5}丁目[0-9]{1,5}番*")
        adPatternArray.Add("[0-9]{1,5}丁目")
        adPatternArray.Add("[0-9]{1,5}番地")
        adPatternArray.Add("[0-9]{1,5}号室")
        adPatternArray.Add("[0-9]{3,5}$")
        For Each ap As String In adPatternArray
            adPatternStr &= ap & "|"
        Next
        adPatternStr = adPatternStr.TrimEnd("|")

        '初期化
        DGV7.Rows.Clear()
        DGV8.Rows.Clear()
        DGV9.Rows.Clear()
        DGV13.Rows.Clear()
        TMSDGV.Rows.Clear()
        ListBox3.Items.Clear()
        DGV17.Rows.Clear()
        ListBox1.Items.Clear()
        CheckBox8.Checked = False
        AzukiControl1.Text = ""
        AzukiControl3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""

        'テンプレート読み込み
        Template1 = File.ReadAllLines(Path.GetDirectoryName(appPath) & "\template\" & TextBox2.Text & ".dat", ENC_SJ)
        Template2 = File.ReadAllLines(Path.GetDirectoryName(appPath) & "\template\" & TextBox3.Text & ".dat", ENC_SJ)
        Template3 = File.ReadAllLines(Path.GetDirectoryName(appPath) & "\template\" & TextBox1.Text & ".dat", ENC_SJ)
        Template4 = File.ReadAllLines(Path.GetDirectoryName(appPath) & "\template\" & TextBox42.Text & ".dat", ENC_SJ)
        Template5 = File.ReadAllLines(Path.GetDirectoryName(appPath) & "\template\" & TextBox14.Text & ".dat", ENC_SJ)
        '伝票設定用ファイル読み取り
        Dim denpyoConfig As String() = File.ReadAllLines(Path.GetDirectoryName(appPath) & "\config\denpyoConfig.txt", ENC_SJ)
        Dim regStr As String = ""
        Dim regIrai As String = ""
        Dim regHakoCode As String = ""
        Dim dFlag As Integer = 0
        For Each s As String In denpyoConfig
            If s = "[伝票依頼主用]" Then
                dFlag = 1
            ElseIf s = "[依頼主変換]" Then
                dFlag = 2
            ElseIf s = "[箱毎コード記入数]" Then
                dFlag = 3
            ElseIf dFlag = 1 Then
                regStr = s
            ElseIf dFlag = 2 Then
                regIrai = s
            ElseIf dFlag = 3 Then
                regHakoCode = s
            End If
        Next

        '伝票ソフト列の値を削除する
        For r As Integer = 0 To DGV1.RowCount - 1
            If DGV1.Item(dH1.IndexOf("伝票ソフト"), r).Style.BackColor <> Color.Yellow Then
                DGV1.Item(dH1.IndexOf("伝票ソフト"), r).Value = ""
            End If
        Next

        ToolStripProgressBar1.Maximum = DGV1.RowCount
        ToolStripProgressBar1.Value = 0
        Application.DoEvents()

        '变换处理
        LIST4VIEW("変換処理", "start")
        For r As Integer = 0 To DGV1.RowCount - 1

            Dim hassou As String = DGV1.Item(dH1.IndexOf("発送方法"), r).Value
            Dim souko As String = DGV1.Item(dH1.IndexOf("発送倉庫"), r).Value
            Dim binsyu As String = DGV1.Item(dH1.IndexOf("便種"), r).Value
            Dim bincount As String = DGV1.Item(dH1.IndexOf("マスタ便数"), r).Value
            Dim Purchaser As String = DGV1.Item(dH1.IndexOf("購入者名"), r).Value
            Dim PurAddress As String = DGV1.Item(dH1.IndexOf("発送先住所"), r).Value
            Dim SenderName As String = DGV1.Item(dH1.IndexOf("発送先名"), r).Value

            Dim codecc As String = DGV1.Item(dH1.IndexOf("商品マスタ"), r).Value

            Dim YU2flagTemp As String = DGV1.Item(dH1.IndexOf("データ"), r).Value

            If YU2flagTemp = "ゆう200" Then
                YU2Denpyus.Add(DGV1.Item(dH1.IndexOf("伝票番号"), r).Value)
            End If

            '这里去判断是不是法人票

            Dim isGroupOrderFlag = CheckIsGroupOrder(SenderName, PurAddress)



            '-----------------------------------------------------请删除-
            'binsyu = "航空便"
            '-----------------------------------------------------请删除-
            '発送倉庫にエラーがある行は処理しない
            Dim soukoError As Boolean = False
            If souko = "" Then
                soukoError = True
                'ElseIf Not Regex.IsMatch(souko, HS1.Text & "|" & HS2.Text & "|" & HS3.Text & "|" & HS4.Text) Then
            ElseIf Not Regex.IsMatch(souko, HS1.Text & "|" & HS2.Text & "|" & HS4.Text) Then
                soukoError = True
            End If
            If soukoError Then
                DGV1.Item(dH1.IndexOf("商品マスタ"), r).Style.BackColor = Color.Orange
                DGV1.CurrentCell = DGV1(dH1.IndexOf("発送倉庫"), r)
                If DGV1(dH1.IndexOf("発送倉庫"), r).Value = "" Then
                    Dim mesline As String = TM_ArIndexof(dH1, "商品マスタ") & "/" & r + 1 & " : マスタ無しエラー（商品担当へ）"
                    ListBox1.Items.Add(mesline)
                    '---------------
                    Dim dNo As String = DGV1(TM_ArIndexof(dH1, "伝票番号"), r).Value
                    ValidateRowAdd(dNo, mesline)
                    '---------------
                Else
                    Dim mesline As String = TM_ArIndexof(dH1, "商品マスタ") & "/" & r + 1 & " : " & DGV1(dH1.IndexOf("発送倉庫"), r).Value & "エラー（商品担当へ）"
                    ListBox1.Items.Add(mesline)
                    '---------------
                    Dim dNo As String = DGV1(TM_ArIndexof(dH1, "伝票番号"), r).Value
                    ValidateRowAdd(dNo, mesline)
                    '---------------
                End If
                Continue For
            End If

            '行毎のテンプレート選択
            Dim TemplateUse As String() = Nothing
            Dim dSoft As String = DGV1.Item(dH1.IndexOf("伝票ソフト"), r).Value
            'If DGV1.Item(dH1.IndexOf("伝票番号"), r).Value = "1869138" Then
            '    Console.WriteLine(123)
            'End If

            If dSoft = "" Then
                TemplateUse = TemplateSet(hassou, souko, binsyu, bincount, isGroupOrderFlag)    'dgv決定、伝票ソフト取得   法人票标记
                dSoft = denpyoSoft
            Else
                Select Case dSoft
                    Case "e飛伝2"
                        TemplateUse = Template1
                        dgvT = DGV7
                        dSoft = "e飛伝2"
                    Case "BIZlogi"
                        TemplateUse = Template2
                        dgvT = DGV8
                        dSoft = "BIZlogi"
                    Case "メール便"
                        TemplateUse = Template3
                        dgvT = DGV9
                        dSoft = "メール便"
                    Case "ヤマト"
                        TemplateUse = Template3
                        dgvT = DGV9
                        dSoft = "ヤマト"
                    Case "定形外"
                        TemplateUse = Template4
                        dgvT = DGV13
                        dSoft = "定形外"
                    Case "TMS"
                        TemplateUse = Template5
                        dgvT = TMSDGV
                        dSoft = "TMS"
                End Select
            End If
            DGV1.Item(dH1.IndexOf("伝票ソフト"), r).Value = dSoft

            Dim dHTemp As ArrayList = TM_HEADER_GET(dgvT)
            dgvT.Rows.Add()
            Dim dgvTNum As Integer = dgvT.RowCount - 1

            '処理
            Dim koguchi As ArrayList = New ArrayList    'メール便用個口リスト
            For i As Integer = 0 To TemplateUse.Length - 1
                If i = 21 Then
                    Console.WriteLine(21)
                End If
                If TemplateUse(i) = "" Or Regex.IsMatch(TemplateUse(i), "^//") Then
                    Continue For
                End If

                Dim TemplateLine As String() = Split(TemplateUse(i), ",")

                'If dSoft = "メール便" Then
                '    Console.WriteLine(123)
                'End If

                'If DGV1.Item(dH1.IndexOf("伝票番号"), r).Value = "1915684" And TemplateLine(0) = "複数個口数" Then
                '    Console.WriteLine(123)
                'End If

                '入力列,[固定],文字
                '入力列,[規定],今日|明日|日付,フォーマット(yyyyMMdd)

                '分岐
                Select Case True
                    '==================================================
                    '元データの背景カラー移行無し
                    Case Regex.IsMatch(TemplateLine(1), "\[固定\]")
                        Try
                            dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = TemplateLine(2)
                        Catch ex As Exception
                            dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = "Error:固定"
                            dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Style.BackColor = Color.Orange
                        End Try
                    '==================================================
                    '元データの背景カラー移行無し
                    Case Regex.IsMatch(TemplateLine(1), "\[規定\]")
                        Try
                            Dim formatStr As String = TemplateLine(3)
                            If formatStr = "" Then
                                formatStr = "yyyyMMdd"
                            End If
                            Select Case True
                                Case Regex.IsMatch(TemplateLine(2), "今日")
                                    dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = Format(Today, formatStr)
                                Case Regex.IsMatch(TemplateLine(2), "明日")
                                    dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = Format(DateAdd(DateInterval.Day, 1, Today), formatStr)
                                Case Regex.IsMatch(TemplateLine(2), "日付")
                                    Dim dStr As String = Replace(TemplateLine(2), "日付=", "")
                                    dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = Format(dStr, formatStr)
                            End Select
                        Catch ex As Exception
                            dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = "Error:規定"
                            dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Style.BackColor = Color.Orange
                        End Try
                    '==================================================
                    '元データの背景移行あり
                    Case Regex.IsMatch(TemplateLine(1), "\[変換\]")
                        Dim henkanC As String = Replace(TemplateLine(1), "[変換]", "")
                        Dim sArray1 As String() = Split(TemplateLine(2), "|")
                        For k As Integer = 0 To sArray1.Length - 1
                            Dim sArray2 As String() = Split(sArray1(k), "=")
                            If sArray2(0) = "" Then
                                If CStr(DGV1.Item(dH1.IndexOf(henkanC), r).Value) = "" Then
                                    dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = sArray2(1)
                                    Exit For
                                End If
                            Else
                                If InStr(DGV1.Item(dH1.IndexOf(henkanC), r).Value, sArray2(0)) > 0 Then
                                    If souko = HS4.Text And dSoft = "BIZlogi" And i = 23 Then '23:時間帯
                                        If sArray2(0) = "午前中" Then
                                            dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = "01"
                                        Else
                                            dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = sArray2(0)
                                        End If
                                    Else
                                        dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = sArray2(1)
                                    End If
                                    Exit For
                                End If
                            End If
                        Next
                        dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Style.BackColor = MotoColor(TemplateLine(1), r)
                    '==================================================
                    '元データの背景移行あり
                    Case Regex.IsMatch(TemplateLine(1), "\[分割\]")
                        Dim henkanC As String = Replace(TemplateLine(1), "[分割]", "")
                        Dim splitword As String = TemplateLine(2)
                        splitword = Replace(splitword, "\s", " ")
                        Dim sArray1 As String() = Split(splitword, "=")
                        Dim str As String = DGV1.Item(dH1.IndexOf(henkanC), r).Value
                        If str <> "" Then
                            Dim sArray2 As String() = Regex.Split(str, sArray1(0))
                            If sArray1(1) = "last" Then '最後のワードを取得
                                dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = sArray2(sArray2.Length - 1)
                            ElseIf sArray2.Length >= sArray1(1) Then
                                dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = sArray2(sArray1(1) - 1)
                            Else
                                dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = ""
                            End If
                        Else
                            dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = ""
                        End If
                        dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Style.BackColor = MotoColor(TemplateLine(1), r)
                    '==================================================
                    '元データの背景移行あり
                    Case Regex.IsMatch(TemplateLine(1), "\[検索\]")
                        Dim henkanC As String = TemplateLine(2)
                        Dim sArray1 As String() = Split(henkanC, "/")       '合計金額=>0/008|   '商品番号=ad/True|False     '商品番号=ad/送料
                        Dim sArray2 As String() = Split(sArray1(0), "=")    '合計金額=>0        '商品番号=ad
                        If InStr(sArray2(1), ">") > 0 Then
                            Dim num As String = Replace(InStr(sArray2(1), ">"), ">", "")
                            If DGV1.Item(dH1.IndexOf(sArray2(0)), r).Value > num Then
                                Dim sArray4 As String() = Split(sArray1(1), "|")
                                dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = sArray4(0)
                            Else
                                Dim sArray4 As String() = Split(sArray1(1), "|")
                                dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = sArray4(1)
                            End If
                        ElseIf InStr(DGV1.Item(dH1.IndexOf(sArray2(0)), r).Value, sArray2(1)) > 0 Then   'TrueWord
                            If InStr(sArray1(1), "[固定]") <> 0 Then
                                Dim henkanD As String = Replace(TemplateLine(2), "[固定]=", "")
                                Dim sArray4 As String() = Split(henkanD, "|")
                                dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = sArray4(0)
                            Else
                                Dim sArray4 As String() = Split(sArray1(1), "|")
                                If InStr(sArray4(0), "+") > 0 Then                      '「+」で連結
                                    Dim s As String() = Split(sArray4(0), "+")
                                    Dim res As String = ""
                                    For k As Integer = 0 To s.Length - 1
                                        res &= DGV1.Item(dH1.IndexOf(s(k)), r).Value
                                    Next
                                    dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = res
                                Else
                                    dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = DGV1.Item(dH1.IndexOf(sArray4(0)), r).Value
                                End If
                            End If
                        Else                                                                 'FalseWord
                            If InStr(sArray1(1), "[固定]") <> 0 Then
                                Dim henkanD As String = Replace(TemplateLine(2), "[固定]=", "")
                                Dim sArray4 As String() = Split(henkanD, "|")
                                dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = sArray4(1)
                            Else
                                Dim sArray4 As String() = Split(sArray1(1), "|")
                                If InStr(sArray4(1), "+") > 0 Then                      '「+」で連結
                                    Dim s As String() = Split(sArray4(1), "+")
                                    Dim res As String = ""
                                    For k As Integer = 0 To s.Length - 1
                                        res &= DGV1.Item(dH1.IndexOf(s(k)), r).Value
                                    Next
                                    dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = res
                                Else
                                    dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = DGV1.Item(dH1.IndexOf(sArray4(1)), r).Value
                                End If
                            End If
                        End If
                        dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Style.BackColor = MotoColor(TemplateLine(2), r)
                     '==================================================
                    '元データの背景移行 未確認
                    Case Regex.IsMatch(TemplateLine(1), "\[選択\]")
                        Dim sentakuC As String = TemplateLine(2)
                        If InStr(sentakuC, "|") > 0 Then
                            Dim arrayA As String() = Split(sentakuC, "|")
                            Dim res As String = ""
                            For k As Integer = 0 To arrayA.Length - 1
                                If CStr(DGV1.Item(dH1.IndexOf(arrayA(k)), r).Value) <> "" Then
                                    res = DGV1.Item(dH1.IndexOf(arrayA(k)), r).Value
                                    Exit For
                                End If
                            Next
                            dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = res
                        Else
                            dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = DGV1.Item(dH1.IndexOf(sentakuC), r).Value
                        End If
                    '==================================================
                    '元データの背景移行 未確認
                    Case Regex.IsMatch(TemplateLine(1), "\[計算\]")
                        Dim keisan As String = TemplateLine(2)
                        Dim parts As String() = Regex.Split(keisan, "(?<1>\*|\/|\-|\+)")
                        Dim exp As String = ""
                        For k As Integer = 0 To parts.Length - 1
                            If Not IsNumeric(parts(k)) And Not Regex.IsMatch(parts(i), "[\*|\/|\-|\+]{1,1}") Then
                                exp &= DGV1.Item(dH1.IndexOf(parts(k)), r).Value
                            Else
                                exp &= parts(k)
                            End If
                        Next
                        '文字列を計算する
                        Dim dt As New System.Data.DataTable()
                        Dim result As Double = CDbl(dt.Compute(exp, ""))
                        dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = result
                    '==================================================
                    Case Regex.IsMatch(TemplateLine(1), "\[特殊\]")
                        '==================================================
                        '元データの背景移行あり
                        If Regex.IsMatch(TemplateLine(2), "\[住所\]") Then
                            Try
                                Dim tokuC As String = Replace(TemplateLine(1), "[特殊]", "")
                                Dim mode As String = Replace(TemplateLine(2), "[住所]", "")

                                '「都道府県/市区町村/それ以降」の3つに分けるなら
                                '(...??[都道府県])((?:旭川|伊達|石狩|盛岡|奥州|田村|南相馬|那須塩原|東村山|武蔵村山|羽村|十日町|上越|富山|野々市|大町|蒲郡|四日市|姫路|大和郡山|廿日市|下松|岩国|田川|大村)市|.+?郡(?:玉村|大町|.+?)[町村]|.+?市.+?区|.+?[市区町村])(.+)

                                '番地で住所を2つに分ける
                                '「福岡県福岡市中央区薬院1-2-3ライオンズマンション薬院第1-111」これが区分できない
                                Dim address As String = DGV1.Item(dH1.IndexOf(tokuC), r).Value
                                address = Regex.Replace(address, "\s|　", "")
                                address = Regex.Replace(address, "－", "-")
                                address = Regex.Replace(address, "([0-9])[ー|－|一]([0-9])", "$1-$2")
                                Dim ad As MatchCollection = Regex.Matches(address, adPatternStr)
                                If ad.Count > 1 Then
                                    Dim m As Match = ad(0)
                                    Dim index As Integer = m.Index + m.Length
                                    If mode = 1 Then
                                        dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = address.Substring(0, index)
                                    ElseIf mode = 2 Then
                                        dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = address.Substring(index)
                                    End If
                                ElseIf ad.Count = 1 Then
                                    Dim m As Match = ad(ad.Count - 1)
                                    Dim index As Integer = m.Index + m.Length
                                    If address.Length = index Then
                                        index = m.Index
                                        If mode = 1 Then
                                            dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = address.Substring(0, index)
                                        ElseIf mode = 2 Then
                                            dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = address.Substring(index)
                                        End If
                                    Else
                                        If mode = 1 Then
                                            dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = address.Substring(0, index)
                                        ElseIf mode = 2 Then
                                            dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = address.Substring(index)
                                        End If
                                    End If
                                Else
                                    If mode = 1 Then
                                        dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = address
                                    Else
                                        dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = ""
                                    End If
                                End If
                                If mode = 2 Then
                                    Dim addr1 As String = dgvT.Item(dHTemp.IndexOf(TemplateLine(0)) - 1, dgvTNum).Value
                                    Dim addr2 As String = dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value
                                    If addr1.Length > 16 And addr2.Length < 10 Then
                                        Dim addr3 As String = address.Substring(0, address.Length / 2)
                                        Dim a = addr3.Substring(addr3.Length - 1, 1)
                                        If IsNumeric(addr3.Substring(addr3.Length - 1, 1)) Then
                                            Dim addr4 As String = address.Substring(address.Length / 2)
                                            Dim num As Integer = 1
                                            For n = 0 To 10
                                                If Not IsNumeric(addr4.Substring(n, 1)) Then
                                                    num = n
                                                    Exit For
                                                End If
                                            Next
                                            addr1 = address.Substring(0, (address.Length / 2) + num)
                                            addr2 = address.Substring((address.Length / 2) + num)
                                        Else
                                            addr1 = address.Substring(0, address.Length / 2)
                                            addr2 = address.Substring(address.Length / 2)
                                        End If
                                        dgvT.Item(dHTemp.IndexOf(TemplateLine(0)) - 1, dgvTNum).Value = addr1
                                        dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = addr2
                                    End If
                                End If

                                dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Style.BackColor = MotoColor(TemplateLine(1), r)

                                '番地未入力チェック
                                If address <> "" Then
                                    Dim addressCheck As Boolean = False
                                    Dim work As String
                                    Dim k As Integer
                                    For k = 0 To address.Length - 1 '住所の文字数分繰り返す
                                        work = address.Substring(k, 1)
                                        If IsNumeric(work) Then
                                            addressCheck = True
                                            Exit For
                                        End If
                                    Next k
                                    If mode = 1 Then
                                        If addressCheck = False Then
                                            dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Style.BackColor = Color.Yellow
                                        End If
                                    End If
                                End If
                            Catch ex As Exception
                                dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = "Error:住所変換"
                                dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Style.BackColor = Color.Orange
                            End Try
                            '==================================================
                        ElseIf Regex.IsMatch(TemplateLine(2), "\[個口\]") Then
                            Dim tokuC As String = Replace(TemplateLine(1), "[特殊]", "")
                            Dim denwa As String = DGV1.Item(dH1.IndexOf(tokuC), r).Value
                            'ネクストエンジンの購入者電話番号で複数個口を指定している時は優先して入れる
                            If InStr(denwa, ",") > 0 Then
                                Dim kosuu As String() = Split(denwa, ",")
                                If IsNumeric(kosuu(1)) Then '個口が数字で、1個以上の場合のみ
                                    If kosuu(1) > 1 Then
                                        dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = kosuu(1)
                                        koguchi.Add(r & "," & kosuu(1))
                                    End If
                                End If
                            Else    '元データで計算したマスタ便数を入れる
                                'TemplateLine(0) e飛伝:出荷個数 bizlogi:個数 メール定形外:複数個口数

                                'If dgvT Is DGV7 Or dgvT Is DGV8 Then
                                '    MasterHaisouKeisan(1, -1, dgvT, r, dgvTNum, TemplateLine(0))
                                'Else
                                '    dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = DGV1.Item(dH1.IndexOf("マスタ便数"), r).Value
                                'End If

                                dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = DGV1.Item(dH1.IndexOf("マスタ便数"), r).Value
                                koguchi.Add(r & "," & 1)
                            End If
                            '==================================================
                        ElseIf Regex.IsMatch(TemplateLine(2), "\[電話\]") Then
                            Dim kounyuMei As String = Regex.Replace(DGV1.Item(dH1.IndexOf("購入者名"), r).Value, "\s|　", "")
                            Dim HassouMei As String = Regex.Replace(DGV1.Item(dH1.IndexOf("発送先名"), r).Value, "\s|　", "")
                            Dim kounyuTel As String = DGV1.Item(dH1.IndexOf("購入者電話番号"), r).Value
                            Dim HassouTel As String = DGV1.Item(dH1.IndexOf("発送先電話番号"), r).Value
                            If kounyuTel <> "" Then
                                kounyuTel = Regex.Replace(kounyuTel, "\s|　|-", "")
                            End If
                            If HassouTel <> "" Then
                                HassouTel = Regex.Replace(HassouTel, "\s|　|-", "")
                            End If

                            If kounyuMei <> HassouMei And kounyuTel <> HassouTel Then
                                dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = kounyuTel
                            Else
                                dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = ""
                            End If
                            '==================================================
                        ElseIf Regex.IsMatch(TemplateLine(2), "\[依頼主\]") Then
                            Dim kounyuMei As String = Regex.Replace(DGV1.Item(dH1.IndexOf("購入者名"), r).Value, "\s|　", "")
                            Dim HassouMei As String = Regex.Replace(DGV1.Item(dH1.IndexOf("発送先名"), r).Value, "\s|　", "")
                            Dim kounyuTel As String = Replace(DGV1.Item(dH1.IndexOf("購入者電話番号"), r).Value, "-", "")
                            Dim kounyuYubin As String = Replace(DGV1.Item(dH1.IndexOf("購入者郵便番号"), r).Value, "-", "")
                            Dim HassouTel As String = Replace(DGV1.Item(dH1.IndexOf("発送先電話番号"), r).Value, "-", "")
                            Dim HassouYubin As String = Replace(DGV1.Item(dH1.IndexOf("発送先郵便番号"), r).Value, "-", "")
                            If kounyuTel <> "" Then
                                kounyuTel = Regex.Replace(kounyuTel, "\s|　|-", "")
                            End If
                            If HassouTel <> "" Then
                                HassouTel = Regex.Replace(HassouTel, "\s|　|-", "")
                            End If

                            If Regex.IsMatch(kounyuMei, "^[a-zA-Z0-9\.\-_\%\&\+\*\/]+$") Then
                                dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = ""
                            ElseIf kounyuMei <> HassouMei Then
                                dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = "(依頼)" & kounyuMei & "様"    '& kounyuTel
                            Else
                                dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = ""
                            End If
                        End If
                    '==================================================
                    '元データの背景移行必要か未確認
                    Case Regex.IsMatch(TemplateLine(1), "\[合成\]")
                        Dim denC As String = Replace(TemplateLine(1), "[合成]", "")    'マスタコードに入れるのはココ
                        Dim denpyoNum As String = DGV1.Item(dH1.IndexOf(denC), r).Value

                        Dim res As String = ""
                        For r2 As Integer = 0 To DGV3.RowCount - 1
                            Dim sDenpyoNum As String = DGV3.Item(dH3.IndexOf("伝票番号"), r2).Value    '伝票番号
                            If denpyoNum = sDenpyoNum Then
                                Dim sCode As String = DGV3.Item(dH3.IndexOf("商品ｺｰﾄﾞ"), r2).Value     '商品ｺｰﾄﾞ
                                Dim sCount As String = DGV3.Item(dH3.IndexOf("受注数"), r2).Value      '受注数
                                Dim sOp As String = DGV3.Item(dH3.IndexOf("商品ｵﾌﾟｼｮﾝ"), r2).Value           '商品ｵﾌﾟｼｮﾝ

                                Dim anchor As String = ".*(?<op>\(\(.*\)\)).*"
                                Dim sOp_re As Object = New Regex(anchor, RegexOptions.IgnoreCase Or RegexOptions.Singleline)
                                Dim m As Match = sOp_re.Match(sOp)
                                Dim opMark As String = ""
                                While m.Success
                                    opMark = m.Groups("op").Value
                                    opMark = Regex.Replace(opMark, "\(\(|\)\)", "")
                                    Exit While
                                End While

                                If res <> "" Then
                                    res &= "、"
                                End If
                                If CheckBox21.Checked = False Then
                                    res &= sCode & "*" & sCount
                                Else
                                    res &= sCode & "(" & opMark & ")" & "*" & sCount
                                End If
                            End If
                        Next
                        dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = res
                        dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Style.BackColor = MotoColor(TemplateLine(1), r)
                        '==================================================
                        '元データの背景移行あり
                    Case Else
                        If InStr(TemplateLine(1), "+") > 0 Then
                            Dim arrayA As String() = Split(TemplateLine(1), "+")
                            Dim res As String = ""
                            For k As Integer = 0 To arrayA.Length - 1
                                If InStr(arrayA(k), Chr(34)) = 0 And InStr(arrayA(k), Chr(39)) = 0 Then
                                    res &= DGV3.Item(arrayA(k), r).Value
                                Else
                                    Dim resA As String = Replace(arrayA(k), Chr(34), "")
                                    resA = Replace(arrayA(k), Chr(39), "")
                                    res &= resA
                                End If
                            Next
                            dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = res
                            dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Style.BackColor = MotoColor(TemplateLine(1), r)
                        Else
                            If dH1.Contains(TemplateLine(1)) Then
                                'If TemplateLine(0) = "マスタ個口" And dgvT Is DGV7 Then
                                '    dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = dgvT.Item(dHTemp.IndexOf("出荷個数"), dgvTNum).Value
                                '    dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Style.BackColor = MotoColor(TemplateLine(1), r)
                                'ElseIf TemplateLine(0) = "マスタ個口" And dgvT Is DGV8 Then
                                '    dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = dgvT.Item(dHTemp.IndexOf("個数"), dgvTNum).Value
                                '    dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Style.BackColor = MotoColor(TemplateLine(1), r)
                                'Else
                                '    dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = DGV1.Item(dH1.IndexOf(TemplateLine(1)), r).Value
                                '    dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Style.BackColor = MotoColor(TemplateLine(1), r)
                                'End If

                                dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Value = DGV1.Item(dH1.IndexOf(TemplateLine(1)), r).Value
                                dgvT.Item(dHTemp.IndexOf(TemplateLine(0)), dgvTNum).Style.BackColor = MotoColor(TemplateLine(1), r)
                                'Console.WriteLine(TemplateLine(0) & " : " & DGV1.Item(dH1.IndexOf(TemplateLine(1)), r).Value)
                            End If
                        End If
                End Select
            Next    'テンプレート行毎

            ToolStripProgressBar1.Value += 1
            Application.DoEvents()
        Next        'グリッドの行毎

        'dgv 新任务
        Dim dgv As DataGridView() = {DGV7, DGV8, DGV9, DGV13, TMSDGV}

        '依頼主合成（グリッド毎）
        LIST4VIEW("依頼主合成", "...")
        For i As Integer = 0 To 3
            MeisaiIrainushi(dgv(i))
        Next

        '注文数を名前の前に表示する
        If CheckBox10.Checked Then
            For i As Integer = 0 To dgv.Length - 1
                ChumonsuName(dgv(i))
            Next
        End If

        'メール便は行複製（箱番号追加）
        If DGV9.RowCount > 0 Then
            LIST4VIEW("メール便行複製", "start")
            DenpyoRowIncrease(DGV9)
        End If
        If DGV13.RowCount > 0 Then
            LIST4VIEW("定形外行複製", "start")
            DenpyoRowIncrease(DGV13)
        End If

        '箱毎コード記入（グリッド毎）
        LIST4VIEW("箱毎コード記入", "start")
        For i As Integer = 0 To dgv.Length - 1
            HakoCodeInput(dgv(i), regHakoCode)
        Next

        '同じ番号をを除く
        If ListBox3.Items.Count > 0 Then
            Dim listbox3_ As New List(Of String)
            For Each s As String In ListBox3.Items
                listbox3_.Add(s)
            Next
            ListBox3.Items.Clear()

            For Each s2 As String In listbox3_.Distinct()
                ListBox3.Items.Add(s2)
            Next
        End If


        Dim dH17 As ArrayList = TM_HEADER_GET(DGV17)
        Dim dgv17_d As New List(Of String)
        Dim dgv17_r As New List(Of Integer)
        For r_17 As Integer = 0 To DGV17.RowCount - 1
            Dim denpyono = DGV17.Item(dH17.IndexOf("伝票番号"), r_17).Value
            Dim deleteFalg = False
            If r_17 = 0 Then
                dgv17_d.Add(denpyono)
            Else
                For Each r3 As String In dgv17_d.Distinct()
                    If r3 = denpyono Then
                        deleteFalg = True
                        Exit For
                    End If
                Next
                If deleteFalg Then
                    dgv17_r.Add(r_17)
                    'DGV17.Rows.RemoveAt(r_17)
                Else
                    dgv17_d.Add(denpyono)
                End If
            End If
        Next

        Dim d_17_index = 0
        For Each d_17 As String In dgv17_r
            If d_17_index = 0 Then
                DGV17.Rows.RemoveAt(d_17)
            Else
                DGV17.Rows.RemoveAt(d_17 - d_17_index)
            End If
            d_17_index += 1
        Next


        'メール便商品複数別紙の対応
        HakoCodeCheck(DGV9, regHakoCode)
        HakoCodeCheck(DGV13, regHakoCode)

        'マークを品名に記入
        For i As Integer = 0 To dgv.Length - 1
            MarkOnHinmei(dgv(i))
        Next

        'シール指定（e飛伝2のみ）
        If DGV7.RowCount > 0 Then
            LIST4VIEW("シール指定（e飛伝2）", "start")
            SealSitei(0)
        End If

        'データの出荷個数を調べる（グリッド毎）
        LIST4VIEW("出荷数チェック", "start")
        For i As Integer = 0 To dgv.Length - 1
            DataGridView_KosuCount(dgv(i))
        Next

        '代引きカウント
        Dim singu As Integer = CInt(DaibikiCount(DGV7, HS1.Text)) + CInt(DaibikiCount(DGV8, HS1.Text))
        GroupBox2.Text = HS1.Text & "（代引き " & singu & "）"
        Dim isouda As Integer = CInt(DaibikiCount(DGV7, HS2.Text)) + CInt(DaibikiCount(DGV8, HS2.Text))
        GroupBox6.Text = HS2.Text & "（代引き " & isouda & "）"
        Dim nagoya As Integer = CInt(DaibikiCount(DGV7, HS4.Text)) + CInt(DaibikiCount(DGV8, HS4.Text))
        GroupBox7.Text = HS4.Text & "（代引き " & nagoya & "）"
        'Dim buturyu As Integer = CInt(DaibikiCount(DGV7, HS3.Text)) + CInt(DaibikiCount(DGV8, HS3.Text))
        'GroupBox7.Text = HS3.Text & "（代引き " & buturyu & "）"

        ToolStripProgressBar1.Value = 0
        Application.DoEvents()
        CheckList(1)

        'ピッキングリスト作成
        If CheckBox13.Checked Then
            DGV16.Rows.Clear()
            Dim dH16 As ArrayList = TM_HEADER_GET(DGV16)
            Dim dgv16Array As New ArrayList
            For r1 As Integer = 0 To DGV1.RowCount - 1
                For r3 As Integer = 0 To DGV3.RowCount - 1
                    If DGV1.Item(dH1.IndexOf("伝票番号"), r1).Value = DGV3.Item(dH3.IndexOf("伝票番号"), r3).Value Then
                        Dim souko As String = DGV1.Item(dH1.IndexOf("発送倉庫"), r1).Value
                        Dim hassou As String = DGV1.Item(dH1.IndexOf("発送方法"), r1).Value
                        Dim code As String = DGV3.Item(dH3.IndexOf("商品ｺｰﾄﾞ"), r3).Value
                        Dim check As String = code & souko & hassou
                        If Not dgv16Array.Contains(check) Then
                            DGV16.Rows.Add()
                            dgv16Array.Add(check)
                            For i As Integer = 0 To dH16.Count - 1
                                If dH16(i) = "商品名" Then
                                    DGV16.Item(i, DGV16.RowCount - 1).Value = DGV6.Item(dH6.IndexOf("商品名"), dgv6CodeArray.IndexOf(code.ToLower)).Value
                                ElseIf dH16(i) = "発送方法" Then
                                    DGV16.Item(i, DGV16.RowCount - 1).Value = hassou
                                ElseIf dH16(i) = "発送倉庫" Then
                                    DGV16.Item(i, DGV16.RowCount - 1).Value = souko
                                Else
                                    DGV16.Item(i, DGV16.RowCount - 1).Value = DGV3.Item(dH3.IndexOf(dH16(i)), r3).Value
                                End If
                            Next
                        Else
                            DGV16.Item(dH16.IndexOf("受注数"), dgv16Array.IndexOf(check)).Value += CInt(DGV3.Item(dH3.IndexOf("受注数"), r3).Value)
                        End If
                    End If
                Next
            Next
        End If

        'For r1 As Integer = 0 To DGV9.RowCount - 1
        '    Dim mCode As String = DGV9.Item(dH9.IndexOf("マスタコード"), r1).Value
        '    Dim mCodeArray As String() = Split(mCode, "、")
        '    If mCodeArray.Length = 1 Then
        '        Dim code As String() = Split(mCodeArray(0), "*")
        '        If code(0).ToLower = "ny263-51" Then
        '            If code(1) = "1" Or code(1) = "2" Then
        '                If code(1) = "1" Then
        '                    DGV9(dH9.IndexOf("重量"), r1).Value = "100"
        '                ElseIf code(1) = "2" Then
        '                    DGV9(dH9.IndexOf("重量"), r1).Value = "200"
        '                End If
        '            End If
        '        End If
        '    ElseIf mCodeArray.Length = 2 Then
        '        Dim code1 As String() = Split(mCodeArray(0), "*")
        '        Dim code2 As String() = Split(mCodeArray(1), "*")
        '        If code1(0) = "ny263-51" And code1(1) = "1" And code2(0) = "ny263-51" And code2(1) = "1" Then
        '            DGV9(dH9.IndexOf("重量"), r1).Value = "200"
        '        End If
        '    End If
        'Next

        '強制処理
        Kyoseisyori()

        Label9.Visible = False
        MsgBox("終了しました", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)

        'If HakoCodeInput_err <> "" Then
        '    HakoCodeInput_err = "佐川BIZlogiの型番がない伝票は下記をご参照ください。" & vbCrLf & HakoCodeInput_err
        '    MsgBox(HakoCodeInput_err, MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
        'End If
    End Sub

    ''' <summary>
    ''' dgv1の修正セル背景カラーを移行する
    ''' </summary>
    ''' <param name="headerStr">ヘッダー文字列（[.*]はここで削除）</param>
    ''' <param name="r">作業行</param>
    ''' <returns>セル背景カラー</returns>
    Private Function MotoColor(headerStr As String, r As Integer) As Color
        Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
        headerStr = Regex.Replace(headerStr, "\[.*\]", "")
        If dH1.Contains(headerStr) Then
            MotoColor = DGV1.Item(dH1.IndexOf(headerStr), r).Style.BackColor
        Else
            MotoColor = Color.Empty
        End If
    End Function

    ''' <summary>
    ''' シール指定（e飛伝2のみ）
    ''' </summary>
    ''' <param name="mode">0=全行、1=選択行、2=行指定</param>
    ''' <param name="num">mode2の時の指定行番号</param>
    Private Sub SealSitei(ByVal mode As Integer, Optional num As Integer = 0)
        Dim dgv As DataGridView = DGV7

        'datagridviewのヘッダーテキストをコレクションに取り込む
        Dim dHSel As ArrayList = TM_HEADER_GET(dgv)

        '********************************************************************
        'グリッド複数または単独モード指定
        Dim rowArray As ArrayList = New ArrayList
        If mode = 0 Then
            For r As Integer = 0 To dgv.RowCount - 1
                rowArray.Add(r)
            Next
        ElseIf mode = 1 Then
            For i As Integer = 0 To dgv.SelectedCells.Count - 1
                If Not rowArray.Contains(dgv.SelectedCells(i).RowIndex) Then
                    rowArray.Add(dgv.SelectedCells(i).RowIndex)
                End If
            Next
        Else
            rowArray.Add(num)
        End If
        '********************************************************************

        For i As Integer = 0 To rowArray.Count - 1
            If dgv.Item(dHSel.IndexOf("代引金額"), rowArray(i)).Value > 0 Then
                '003 航空便
                If dgv.Item(dHSel.IndexOf("便種（スピードを選択）"), rowArray(i)).Value = "003" Then
                    dgv.Item(dHSel.IndexOf("指定シール１"), rowArray(i)).Value = ""
                Else
                    dgv.Item(dHSel.IndexOf("指定シール１"), rowArray(i)).Value = "008"
                End If
            Else
                dgv.Item(dHSel.IndexOf("指定シール１"), rowArray(i)).Value = ""
            End If
            If dgv.Item(dHSel.IndexOf("営業所止めフラグ"), rowArray(i)).Value = "1" Then
                '003 航空便
                If dgv.Item(dHSel.IndexOf("便種（スピードを選択）"), rowArray(i)).Value = "003" Then
                    dgv.Item(dHSel.IndexOf("指定シール２"), rowArray(i)).Value = ""
                Else
                    dgv.Item(dHSel.IndexOf("指定シール２"), rowArray(i)).Value = "004"
                End If
            ElseIf CStr(dgv.Item(dHSel.IndexOf("配達指定時間帯"), rowArray(i)).Value) <> "" Then
                '003 航空便じゃない
                If dgv.Item(dHSel.IndexOf("便種（スピードを選択）"), rowArray(i)).Value <> "003" Then
                    If dgv.Item(dHSel.IndexOf("配達指定時間帯"), rowArray(i)).Value = "01" Then
                        dgv.Item(dHSel.IndexOf("指定シール２"), rowArray(i)).Value = "016"
                    Else
                        dgv.Item(dHSel.IndexOf("指定シール２"), rowArray(i)).Value = "007"
                    End If
                End If
            Else
                '003 航空便じゃない
                If dgv.Item(dHSel.IndexOf("便種（スピードを選択）"), rowArray(i)).Value <> "003" Then
                    If IsNumeric(dgv.Item(dHSel.IndexOf("配達日"), rowArray(i)).Value) Then
                        dgv.Item(dHSel.IndexOf("指定シール２"), rowArray(i)).Value = "005"
                    Else
                        dgv.Item(dHSel.IndexOf("指定シール２"), rowArray(i)).Value = ""
                    End If
                End If
            End If
            '003 航空便
            If dgv.Item(dHSel.IndexOf("便種（スピードを選択）"), rowArray(i)).Value = "003" Then
                dgv.Item(dHSel.IndexOf("指定シール３"), rowArray(i)).Value = "017"
            Else
                If dgv.Item(dHSel.IndexOf("指定シール３"), rowArray(i)).Value = "" Then
                    dgv.Item(dHSel.IndexOf("指定シール３"), rowArray(i)).Value = "011"
                End If
            End If

            '航空便シール（空いている所に入れる）
            If dgv.Item(dHSel.IndexOf("便種（スピードを選択）"), rowArray(i)).Value = "003" Then
                'If dgv.Item(dHSel.IndexOf("指定シール１"), rowArray(i)).Value = "" Then
                '    dgv.Item(dHSel.IndexOf("指定シール１"), rowArray(i)).Value = "017"
                'ElseIf dgv.Item(dHSel.IndexOf("指定シール２"), rowArray(i)).Value = "" Then
                '    dgv.Item(dHSel.IndexOf("指定シール２"), rowArray(i)).Value = "017"
                'Else
                '    dgv.Item(dHSel.IndexOf("指定シール３"), rowArray(i)).Value = "017"
                'End If

                If dgv.Item(dHSel.IndexOf("指定シール２"), rowArray(i)).Value <> "017" And dgv.Item(dHSel.IndexOf("指定シール３"), rowArray(i)).Value <> "017" Then
                    dgv.Item(dHSel.IndexOf("指定シール１"), rowArray(i)).Value = "017"
                ElseIf dgv.Item(dHSel.IndexOf("指定シール１"), rowArray(i)).Value <> "017" And dgv.Item(dHSel.IndexOf("指定シール３"), rowArray(i)).Value <> "017" Then
                    dgv.Item(dHSel.IndexOf("指定シール２"), rowArray(i)).Value = "017"
                ElseIf dgv.Item(dHSel.IndexOf("指定シール１"), rowArray(i)).Value <> "017" And dgv.Item(dHSel.IndexOf("指定シール２"), rowArray(i)).Value <> "017" Then
                    dgv.Item(dHSel.IndexOf("指定シール３"), rowArray(i)).Value = "017"
                End If
            End If
        Next
    End Sub

    Private Sub ChumonsuName(dgv As DataGridView)
        Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
        Dim dHSel As ArrayList = TM_HEADER_GET(dgv)
        Dim num As Integer = 0
        If dgv Is DGV7 Then
            num = 1
        ElseIf dgv Is DGV8 Then
            num = 2
        ElseIf dgv Is DGV9 Then
            num = 3
        ElseIf dgv Is DGV13 Then
            num = 4
        ElseIf dgv Is TMSDGV Then
            num = 5
        End If
        For r As Integer = 0 To dgv.RowCount - 1
            Dim sakiName As String = dgv.Item(dHSel.IndexOf(koumoku("sakiname")(num)), r).Value
            Dim masterCodeArray As String() = Split(dgv.Item(dHSel.IndexOf(koumoku("masterCode")(num)), r).Value, "、")
            Dim ChumonSu As Integer = 0
            For k As Integer = 0 To masterCodeArray.Length - 1
                Dim resArray2 As String() = Split(masterCodeArray(k), "*")
                If IsNumeric(resArray2(1)) Then
                    ChumonSu += resArray2(1)
                End If
            Next

            Dim yoyaku As String = ""
            Dim dNo As String = dgv.Item(dHSel.IndexOf(koumoku("denpyoNo")(num)), r).Value
            For r1 As Integer = 0 To DGV1.RowCount - 1
                If DGV1.Item(dH1.IndexOf("伝票番号"), r1).Value = dNo Then
                    If InStr(DGV1.Item(dH1.IndexOf("予約"), r1).Value, "予約") > 0 Then
                        yoyaku = "予"
                        Exit For
                    End If
                End If
            Next

            dgv.Item(TM_ArIndexof(dHSel, koumoku("sakiname")(num)), r).Value = sakiName.Insert(0, "(" & yoyaku & ChumonSu & ")")
        Next
    End Sub

    '箱毎コード記入
    Dim HakoCodeInput_err As String = ""
    Private Sub HakoCodeInput(dgv As DataGridView, regHakoCode As String)
        Dim rh As String() = Split(regHakoCode, ",")
        Dim hSettei As String() = Nothing
        Dim headerDN As String = ""
        Dim soukoDN As String = ""
        Dim binsyuDN As String = ""
        Dim hmeiDN As String = ""
        Dim headerKS As String = ""
        Dim num As Integer = 0
        Dim marumode As Integer = 0

        If dgv Is DGV7 Then
            hSettei = Split(rh(1), "/")
            headerDN = koumoku("denpyoNo")(1)
            soukoDN = koumoku("syori2")(1)
            binsyuDN = koumoku("binsyu")(1)
            hmeiDN = koumoku("sakiname")(1)
            headerKS = koumoku("masterKoguchi")(1)
            marumode = 1
            num = 1
        ElseIf dgv Is DGV8 Then
            hSettei = Split(rh(2), "/")
            headerDN = koumoku("denpyoNo")(2)
            soukoDN = koumoku("syori2")(2)
            binsyuDN = koumoku("binsyu")(2)
            hmeiDN = koumoku("sakiname")(2)
            headerKS = koumoku("masterKoguchi")(2)
            marumode = 0
            num = 2
        ElseIf dgv Is DGV9 Then
            hSettei = Split(rh(3), "/")
            headerDN = koumoku("denpyoNo")(3)
            soukoDN = koumoku("syori2")(3)
            binsyuDN = koumoku("binsyu")(3)
            hmeiDN = koumoku("sakiname")(3)
            headerKS = koumoku("masterKoguchi")(3)
            num = 3
        ElseIf dgv Is DGV13 Then
            hSettei = Split(rh(4), "/")
            headerDN = koumoku("denpyoNo")(4)
            soukoDN = koumoku("syori2")(4)
            binsyuDN = koumoku("binsyu")(4)
            hmeiDN = koumoku("sakiname")(4)
            headerKS = koumoku("masterKoguchi")(4)
            num = 4
        ElseIf dgv Is TMSDGV Then
            hSettei = Split(rh(5), "/")
            headerDN = koumoku("denpyoNo")(5)
            soukoDN = koumoku("syori2")(5)
            binsyuDN = koumoku("binsyu")(5)
            hmeiDN = koumoku("sakiname")(5)
            headerKS = koumoku("masterKoguchi")(5)
            marumode = 1
            num = 5
        End If
        Dim hSetteiHeader As String() = Split(hSettei(2), "|")

        Dim dHSel As ArrayList = TM_HEADER_GET(dgv)
        Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
        'Dim dH6 As ArrayList = TM_HEADER_GET(dgv6)
        Dim sCodeHeader As String() = New String() {"商品コード1", "商品コード2", "商品コード3", "商品コード4"}

        Dim mcCodeArray As New ArrayList
        Dim mcPMArray As New ArrayList
        Dim mcNameArray As New ArrayList    '表示が1つだけの時商品名を入れる

        Dim dH3 As ArrayList = TM_HEADER_GET(DGV3)
        LIST4VIEW("HakoCodeInput", "start")
        For r As Integer = 0 To dgv.RowCount - 1
            '***************************************************************************
            'dgv1に反映するため、伝票番号取得
            '***************************************************************************
            Dim selDenpyoNo As String = dgv.Item(dHSel.IndexOf(koumoku("denpyoNo")(num)), r).Value  '伝票番号取得
            CurDenpyoNo = selDenpyoNo
            Dim dgv1Row As Integer = 0      'dgv1の伝票番号列を取得する
            For r1 As Integer = 0 To DGV1.RowCount - 1
                If selDenpyoNo = DGV1.Item(dH1.IndexOf("伝票番号"), r1).Value Then
                    dgv1Row = r1
                    Exit For
                End If
            Next
            '***************************************************************************

            Dim masterCode As String() = Split(dgv.Item(dHSel.IndexOf("マスタコード"), r).Value, "、")
            Dim binsu As String = dgv.Item(dHSel.IndexOf(headerKS), r).Value

            'メール便************************************************
            If dgv Is DGV9 Or dgv Is DGV13 Then    'メール便は複製した行に商品コードを振り分ける
                'PM値を計算して商品コードを振り分ける
                If CStr(dgv.Item(dHSel.IndexOf("箱番号"), r).Value) <> "" Then
                    If Regex.IsMatch(dgv.Item(dHSel.IndexOf("箱番号"), r).Value, "^1/") Then
                        mcCodeArray.Clear()
                        mcPMArray.Clear()
                        mcNameArray.Clear()
                        'マスタコードを1商品ずつに分ける
                        For i As Integer = 0 To masterCode.Length - 1
                            Dim code As String() = Split(masterCode(i), "*")
                            For k As Integer = 0 To code(1) - 1
                                mcCodeArray.Add(code(0))
                                Dim mStr As String = MasterWeight(code(0))
                                mcPMArray.Add(Regex.Replace(MasterWeight(code(0)), "M|m|P|p|T|t", ""))
                                mcNameArray.Add(MasterItemName(code(0)))
                            Next
                        Next
                    End If
                Else
                    mcCodeArray.Clear()
                    mcPMArray.Clear()
                    mcNameArray.Clear()
                    'マスタコードを1商品ずつに分ける
                    For i As Integer = 0 To masterCode.Length - 1
                        Dim code As String() = Split(masterCode(i), "*")
                        For k As Integer = 0 To code(1) - 1
                            mcCodeArray.Add(code(0))


                            Dim mStr As String = MasterWeight(code(0))
                            mcPMArray.Add(Regex.Replace(MasterWeight(code(0)), "M|m|P|p|T|t", ""))
                            mcNameArray.Add(MasterItemName(code(0)))
                        Next
                    Next
                End If

                '1便のMAX

                Dim sizeMax As Integer = 100
                If dgv Is DGV13 Then
                    sizeMax = NumericUpDown6.Value
                    'sizeMax = 300
                End If
                '这里
                'sizeMax = 300
                '無理やり変換したときの商品表示（未完成）
                'どうするか・・・

                mcCodeArray.Reverse()
                mcPMArray.Reverse()
                mcNameArray.Reverse()
                Dim kei As Double = 0
                Dim dhCol As Integer = 0    '入力するcol列
                Dim colArray(4) As String
                'If DGV13.Item(dHSel.IndexOf("お客様側管理番号"), r).Value = "3479882" Then
                '    Console.WriteLine(123)
                'End If

                'test   FALSE 一定记得删除
                'sizeMax = 300

                'LIST4VIEW("dgv RowCount1", "system")

                For i As Integer = mcCodeArray.Count - 1 To 0 Step -1




                    'LIST4VIEW("dgv RowCount2", "system")
                    Dim hSouko As String = DGV1.Item(dH1.IndexOf("発送倉庫"), dgv1Row).Value
                    If kei + mcPMArray(i) <= sizeMax Then   '100％（定形外250）までのサイズ計算
                        Dim res As String = mcCodeArray(i) & "*1"
                        Dim valiCode As String = Split(res, "*")(0)
                        Dim inputStr As String = ""
                        'hSettei(1) = 2
                        '删除
                        If dhCol >= 0 And dhCol < hSettei(1) Then
                            If dhCol = 0 Then
                                'inputStr = valiCode & "*1"
                                'dgv.Item(dHSel.IndexOf(hSetteiHeader(dhCol)), r).Value = res
                                inputStr = res
                                Dim val_ As String() = Split(ValiationAdd1(valiCode, hSouko, 1, 0, True), "|")
                                dgv.Item(dHSel.IndexOf(hSetteiHeader(dhCol)), r).Value = val_(0)
                            Else
                                Dim mae As String() = Split(colArray(dhCol - 1), "*")
                                Dim ato As String() = Split(res, "*")
                                If mae(0) = ato(0) Then
                                    If dhCol > 0 Then
                                        dhCol = dhCol - 1
                                    End If
                                    Dim mArray As String() = Split(mae(1), " ")
                                    If mArray.Length > 1 Then
                                        Dim chumonsu As Integer = CInt(mArray(0)) + 1
                                        'inputStr = mae(0) & "*" & chumonsu
                                        'dgv.Item(dHSel.IndexOf(hSetteiHeader(dhCol)), r).Value = mae(0) & "*" & MaruMojiConv(chumonsu) & " " & mArray(1)
                                        inputStr = mae(0) & "*" & chumonsu
                                        Dim val_ As String() = Split(ValiationAdd1(mae(0), hSouko, chumonsu, 0, True), "|")
                                        dgv.Item(dHSel.IndexOf(hSetteiHeader(dhCol)), r).Value = val_(0)
                                    Else
                                        Dim chumonsu As Integer = CInt(MaruMojiModoshi(mae(1))) + 1
                                        'inputStr = mae(0) & "*" & chumonsu
                                        'dgv.Item(dHSel.IndexOf(hSetteiHeader(dhCol)), r).Value = mae(0) & "*" & MaruMojiConv(chumonsu)
                                        inputStr = mae(0) & "*" & chumonsu
                                        Dim val_ As String() = Split(ValiationAdd1(mae(0), hSouko, chumonsu, 0, True), "|")
                                        dgv.Item(dHSel.IndexOf(hSetteiHeader(dhCol)), r).Value = val_(0)
                                    End If
                                Else
                                    'inputStr = valiCode & "*1"
                                    'dgv.Item(dHSel.IndexOf(hSetteiHeader(dhCol)), r).Value = res
                                    inputStr = res
                                    Dim val_ As String() = Split(ValiationAdd1(valiCode, hSouko, 1, 0, True), "|")
                                    dgv.Item(dHSel.IndexOf(hSetteiHeader(dhCol)), r).Value = val_(0)
                                End If
                            End If

                            colArray(dhCol) = inputStr
                            'dgv1の商品コード1～4に入力する
                            DGV1.Item(dH1.IndexOf(sCodeHeader(dhCol)), dgv1Row).Value = inputStr

                            kei += mcPMArray(i)
                            dgv.Item(dHSel.IndexOf("重量"), r).Value = Math.Floor(kei)

                            mcCodeArray.RemoveAt(i)
                            mcPMArray.RemoveAt(i)
                            dhCol += 1
                        Else
                            '指定数以上は入らない   加入别纸票
                            ListBox3.Items.Add(dgv.Item(dHSel.IndexOf(headerDN), r).Value)
                            If binsyuDN <> "" Then
                                DGV17.Rows.Add(dgv.Item(dHSel.IndexOf(headerDN), r).Value, dgv.Item(dHSel.IndexOf(hmeiDN), r).Value, dgv.Item(dHSel.IndexOf(soukoDN), r).Value, dgv.Item(dHSel.IndexOf(binsyuDN), r).Value)
                            Else
                                DGV17.Rows.Add(dgv.Item(dHSel.IndexOf(headerDN), r).Value, dgv.Item(dHSel.IndexOf(hmeiDN), r).Value, dgv.Item(dHSel.IndexOf(soukoDN), r).Value, "")
                            End If
                            dgv.Item(dHSel.IndexOf(hSetteiHeader(0)), r).Value = "★商品複数別処理"
                            If InStr(dgv.Item(dHSel.IndexOf(koumoku("bikou2")(num)), r).Value, "★") = 0 Then
                                dgv.Item(dHSel.IndexOf(koumoku("bikou2")(num)), r).Value = "★"
                            End If
                            For k As Integer = 1 To 3
                                dgv.Item(dHSel.IndexOf(hSetteiHeader(k)), r).Value = ""
                            Next
                            DGV1.Item(dH1.IndexOf(sCodeHeader(0)), dgv1Row).Value = "★商品複数別処理"
                            If InStr(DGV1.Item(dH1.IndexOf(koumoku("bikou2")(0)), dgv1Row).Value, "★") = 0 Then
                                DGV1.Item(dH1.IndexOf(koumoku("bikou2")(0)), dgv1Row).Value = "★"
                            End If
                            Exit For
                        End If
                    End If
                Next

                'If dgv.Item(dHSel.IndexOf("お客様側管理番号"), r).Value = "1866352" Then
                '    Console.WriteLine(123)
                'End If
                '複数伝票にした時に商品が漏れないよう、最後の伝票でチェックする
                If CStr(dgv.Item(dHSel.IndexOf("箱番号"), r).Value) <> "" Then
                    Dim hb As String() = Split(dgv.Item(dHSel.IndexOf("箱番号"), r).Value, "/")
                    If hb(0) = hb(1) Then               '箱番号が最後 
                        If mcCodeArray.Count > 0 Then   'コードが残っていないか
                            Dim dNo As String = dgv.Item(dHSel.IndexOf(headerDN), r).Value
                            For r3 As Integer = 0 To dgv.RowCount - 1
                                If dNo = dgv.Item(dHSel.IndexOf(headerDN), r3).Value Then
                                    dgv.Item(dHSel.IndexOf(hSetteiHeader(0)), r3).Value = "無理変換★計算不能"
                                End If
                            Next
                        End If
                    End If
                    'Else    '箱番号が無い時にコードが残っていた場合
                    '    If mcCodeArray.Count > 0 Then
                    '        Dim dNo As String = dgv.Item(dHSel.IndexOf(headerDN), r).Value
                    '        For r3 As Integer = 0 To dgv.RowCount - 1
                    '            If dNo = dgv.Item(dHSel.IndexOf(headerDN), r3).Value Then
                    '                dgv.Item(dHSel.IndexOf(hSetteiHeader(0)), r3).Value = "無理変換★コード残り"
                    '            End If
                    '        Next
                    '    End If
                End If
                '--------------------------------------------------------------

                If colArray(0) = "" Then    '無理やり変換した時ここで弾く
                    ListBox3.Items.Add(dgv.Item(dHSel.IndexOf(headerDN), r).Value)
                    Dim binsyu As String = ""
                    If binsyuDN <> "" Then
                        binsyu = dgv.Item(dHSel.IndexOf(binsyuDN), r).Value
                    End If
                    DGV17.Rows.Add(dgv.Item(dHSel.IndexOf(headerDN), r).Value, dgv.Item(dHSel.IndexOf(hmeiDN), r).Value, dgv.Item(dHSel.IndexOf(soukoDN), r).Value, binsyu)
                    dgv.Item(dHSel.IndexOf(hSetteiHeader(0)), r).Value = "★計算不能（伝票戻して）"
                    dgv.Item(dHSel.IndexOf(hSetteiHeader(0)), r).Style.BackColor = Color.Red
                    If InStr(dgv.Item(dHSel.IndexOf(koumoku("bikou2")(num)), r).Value, "★") = 0 Then
                        dgv.Item(dHSel.IndexOf(koumoku("bikou2")(num)), r).Value = "★"
                    End If
                    For k As Integer = 1 To 3
                        dgv.Item(dHSel.IndexOf(hSetteiHeader(k)), r).Value = ""
                    Next
                    DGV1.Item(dH1.IndexOf(sCodeHeader(0)), dgv1Row).Value = "★計算不能（伝票戻して）"
                    DGV1.Item(dH1.IndexOf(sCodeHeader(0)), dgv1Row).Style.BackColor = Color.Red
                    If InStr(DGV1.Item(dH1.IndexOf(koumoku("bikou2")(0)), dgv1Row).Value, "★") = 0 Then
                        DGV1.Item(dH1.IndexOf(koumoku("bikou2")(0)), dgv1Row).Value = "★"
                    End If
                    'Else
                    '    For i As Integer = 0 To 3
                    '        Dim str As String = colArray(i)
                    '        If InStr(str, "*") > 0 Then
                    '            Dim strArray As String() = Split(str, " (")
                    '            Dim code As String() = Split(strArray(0), "*")
                    '            Dim res As String = ValiationAdd1(code(0), MaruMojiModoshi(code(1)), 0, True)
                    '            'dgv.Item(dHSel.IndexOf(hSetteiHeader(i)), r).Value = strArray(0) & "*" & MaruMojiConv(CInt(strArray(1)))
                    '            dgv.Item(dHSel.IndexOf(hSetteiHeader(i)), r).Value = res
                    '        End If
                    '    Next
                End If

                If IsNumeric(dgv.Item(dHSel.IndexOf(headerDN), r).Value) Then
                    For r2 As Integer = 0 To DGV3.RowCount - 1
                        If dgv.Item(dHSel.IndexOf(headerDN), r).Value = DGV3.Item(dH3.IndexOf("伝票番号"), r2).Value Then
                            If DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "卸" Then
                                ListBox3.Items.Add(dgv.Item(dHSel.IndexOf(headerDN), r).Value)
                                If binsyuDN <> "" Then
                                    DGV17.Rows.Add(dgv.Item(dHSel.IndexOf(headerDN), r).Value, dgv.Item(dHSel.IndexOf(hmeiDN), r).Value, dgv.Item(dHSel.IndexOf(soukoDN), r).Value, dgv.Item(dHSel.IndexOf(binsyuDN), r).Value)
                                Else
                                    DGV17.Rows.Add(dgv.Item(dHSel.IndexOf(headerDN), r).Value, dgv.Item(dHSel.IndexOf(hmeiDN), r).Value, dgv.Item(dHSel.IndexOf(soukoDN), r).Value, "")
                                End If
                                DGV1.Item(dH1.IndexOf(sCodeHeader(0)), dgv1Row).Value = "★商品複数別処理"
                                If InStr(DGV1.Item(dH1.IndexOf(koumoku("bikou2")(0)), dgv1Row).Value, "★") = 0 Then
                                    DGV1.Item(dH1.IndexOf(koumoku("bikou2")(0)), dgv1Row).Value = "★"
                                End If
                            End If
                        End If
                    Next
                End If

                'ゆうパケット・定形外商品名補助
                If CheckBox6.Checked Then
                    If CStr(dgv.Item(dHSel.IndexOf(hSetteiHeader(1)), r).Value) = "" Then
                        Try
                            '品名1のコードから品名を計算
                            Dim code1 As String = Split(dgv.Item(dHSel.IndexOf(hSetteiHeader(0)), r).Value, "*")(0)
                            code1 = Regex.Replace(code1, "[0123][A-Za-z][0-9A-Za-z][0-9]{0,3}-|BS[0-9]{0,3}-|SD[0-9]{0,3}-|XX[0-9Xx]{0,3}-", "")
                            Dim codeName As String = DGV6.Item(dH6.IndexOf("商品名"), dgv6CodeArray.IndexOf(code1.ToLower)).Value
                            Dim nameStr As String() = Regex.Split(codeName, "\s|　")
                            If dgv IsNot DGV9 Then
                                dgv.Item(dHSel.IndexOf(hSetteiHeader(1)), r).Value = "※" & nameStr(0)
                            End If
                            'dgv.Item(dHSel.IndexOf(hSetteiHeader(1)), r).Value = "※" & nameStr(0)
                        Catch ex As Exception

                        End Try
                    End If
                End If
            Else    '宅配便************************************************


                '一定删除  true
                If masterCode.Length > hSettei(1) Then
                    '这里有别纸

                    'dgv.Item(dHSel.IndexOf("処理用2"), r).Value
                    '一定删除
                    'DGV1.Item(dH1.IndexOf(sCodeHeader(0)), dgv1Row).Value = "别纸删除"



                    ListBox3.Items.Add(dgv.Item(dHSel.IndexOf(headerDN), r).Value)
                    DGV17.Rows.Add(dgv.Item(dHSel.IndexOf(headerDN), r).Value, dgv.Item(dHSel.IndexOf(hmeiDN), r).Value, dgv.Item(dHSel.IndexOf(soukoDN), r).Value, dgv.Item(dHSel.IndexOf(binsyuDN), r).Value)
                    dgv.Item(dHSel.IndexOf(hSetteiHeader(0)), r).Value = "★商品複数別処理"
                    If InStr(dgv.Item(dHSel.IndexOf(koumoku("bikou2")(num)), r).Value, "★") = 0 Then
                        dgv.Item(dHSel.IndexOf(koumoku("bikou2")(num)), r).Value = "★"
                    End If
                    DGV1.Item(dH1.IndexOf(sCodeHeader(0)), dgv1Row).Value = "★商品複数別処理"
                    If InStr(DGV1.Item(dH1.IndexOf(koumoku("bikou2")(0)), dgv1Row).Value, "★") = 0 Then
                        DGV1.Item(dH1.IndexOf(koumoku("bikou2")(0)), dgv1Row).Value = "★"
                    End If
                Else
                    For i As Integer = 0 To masterCode.Length - 1
                        Dim code As String() = Split(masterCode(i), "*")
                        Dim locationAdd As Boolean = True
                        Dim res As String = ""
                        Dim hSouko As String = DGV1.Item(dH1.IndexOf("発送倉庫"), dgv1Row).Value
                        'If hSouko = HS3.Text Then
                        '    If CheckBox7.Checked Then '物流ロケーション追加フラグ
                        '        locationAdd = False
                        '    End If
                        '    Dim val_ As String() = Split(ValiationAdd1(code(0), code(1), 1, locationAdd), "|")  '物流は丸付き文字を解除する
                        '    res = val_(0)
                        '    If val_(1) = "ng" Then
                        '        HakoCodeInput_err &= "行no: " & r + 1 & " , コード: " & code(0) & " , " & hSetteiHeader(i) & ":" & res & vbCrLf
                        '    End If
                        If hSouko = HS2.Text Then
                            If CheckBox31.Checked Then '井相田ロケーション追加フラグ
                                locationAdd = False
                            End If
                            Dim val_ As String() = Split(ValiationAdd1(code(0), hSouko, code(1), marumode, locationAdd), "|")
                            res = val_(0)
                            If val_(1) = "ng" Then
                                HakoCodeInput_err &= "行no: " & r + 1 & " , コード: " & code(0) & " , " & hSetteiHeader(i) & ":" & res & vbCrLf
                            End If
                        Else
                            Dim val_ As String() = Split(ValiationAdd1(code(0), hSouko, code(1), marumode, locationAdd), "|")
                            res = val_(0)
                            If val_(1) = "ng" Then
                                HakoCodeInput_err &= "行no: " & r + 1 & " , コード: " & code(0) & " , " & hSetteiHeader(i) & ":" & res & vbCrLf
                            End If
                        End If

                        If dgv Is DGV8 Then
                            Dim souko As String = dgv.Item(dHSel.IndexOf("処理用2"), r).Value
                            If (hSetteiHeader(i) = "記事欄06" Or hSetteiHeader(i) = "記事欄07" Or hSetteiHeader(i) = "記事欄08" Or hSetteiHeader(i) = "記事欄09") And souko = HS4.Text Then
                                If hSetteiHeader(i) = "記事欄06" Then
                                    dgv.Item(dHSel.IndexOf("記事欄02"), r).Value = res
                                ElseIf hSetteiHeader(i) = "記事欄07" Then
                                    dgv.Item(dHSel.IndexOf("記事欄03"), r).Value = res
                                ElseIf hSetteiHeader(i) = "記事欄08" Then
                                    dgv.Item(dHSel.IndexOf("記事欄04"), r).Value = res
                                ElseIf hSetteiHeader(i) = "記事欄09" Then
                                    dgv.Item(dHSel.IndexOf("記事欄05"), r).Value = res
                                End If
                                dgv.Item(dHSel.IndexOf(hSetteiHeader(i)), r).Value = ""
                            Else
                                dgv.Item(dHSel.IndexOf(hSetteiHeader(i)), r).Value = res
                            End If
                        Else
                            dgv.Item(dHSel.IndexOf(hSetteiHeader(i)), r).Value = res
                        End If

                        DGV1.Item(dH1.IndexOf(sCodeHeader(i)), dgv1Row).Value = masterCode(i)
                    Next

                    If IsNumeric(dgv.Item(dHSel.IndexOf(headerDN), r).Value) Then
                        For r2 As Integer = 0 To DGV3.RowCount - 1
                            If dgv.Item(dHSel.IndexOf(headerDN), r).Value = DGV3.Item(dH3.IndexOf("伝票番号"), r2).Value Then
                                If DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "卸" Then
                                    ListBox3.Items.Add(dgv.Item(dHSel.IndexOf(headerDN), r).Value)
                                    DGV1.Item(dH1.IndexOf(sCodeHeader(0)), dgv1Row).Value = "★商品複数別処理"
                                    If InStr(DGV1.Item(dH1.IndexOf(koumoku("bikou2")(0)), dgv1Row).Value, "★") = 0 Then
                                        DGV1.Item(dH1.IndexOf(koumoku("bikou2")(0)), dgv1Row).Value = "★"
                                    End If
                                End If
                                If binsyuDN <> "" Then
                                    DGV17.Rows.Add(dgv.Item(dHSel.IndexOf(headerDN), r).Value, dgv.Item(dHSel.IndexOf(hmeiDN), r).Value, dgv.Item(dHSel.IndexOf(soukoDN), r).Value, dgv.Item(dHSel.IndexOf(binsyuDN), r).Value)
                                Else
                                    DGV17.Rows.Add(dgv.Item(dHSel.IndexOf(headerDN), r).Value, dgv.Item(dHSel.IndexOf(hmeiDN), r).Value, dgv.Item(dHSel.IndexOf(soukoDN), r).Value, "")
                                End If
                            End If
                        Next
                    End If
                End If

            End If

            ''品名（e飛伝2だけ「ご依頼主住所２」に表示する）
            'Dim mark As String = "品名"
            'Dim hinmeiStr As String() = Regex.Split(DGV1.Item(dH1.IndexOf("品名"), dgv1Row).Value, "\s|　")
            'If DGV1.Item(dH1.IndexOf("マーク"), dgv1Row).Value <> "" Then
            '    mark = DGV1.Item(dH1.IndexOf("マーク"), dgv1Row).Value
            '    If dgv Is DGV7 Then
            '        dgv.Item(dHSel.IndexOf(koumoku("hinmei")(1)), r).Value = "＜" & mark & "＞" & hinmeiStr(0)
            '    ElseIf dgv Is DGV8 Then
            '        dgv.Item(dHSel.IndexOf(koumoku("hinmei")(2)), r).Value = "＜" & mark & "＞" & hinmeiStr(0)
            '    ElseIf dgv Is DGV9 Then
            '        dgv.Item(dHSel.IndexOf(koumoku("hinmei")(3)), r).Value = "＜" & mark & "＞" & hinmeiStr(0)
            '    ElseIf dgv Is DGV13 Then
            '        dgv.Item(dHSel.IndexOf(koumoku("hinmei")(4)), r).Value = "＜" & mark & "＞" & hinmeiStr(0)
            '    End If
            'Else
            '    If dgv Is DGV7 Then
            '        dgv.Item(dHSel.IndexOf(koumoku("hinmei")(1)), r).Value = "＜" & mark & "＞" & hinmeiStr(0)
            '    ElseIf dgv Is DGV8 Then
            '        dgv.Item(dHSel.IndexOf(koumoku("hinmei")(2)), r).Value = "＜" & mark & "＞" & hinmeiStr(0)
            '    End If
            'End If
        Next
    End Sub

    'マークを品名の前に付ける
    Private Sub MarkOnHinmei(dgv As DataGridView)
        Dim dHSel As ArrayList = TM_HEADER_GET(dgv)
        Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
        Dim dgvNum As Integer = 0

        If dgv Is DGV7 Then
            dgvNum = 1
        ElseIf dgv Is DGV8 Then
            dgvNum = 2
        ElseIf dgv Is DGV9 Then
            dgvNum = 3
        ElseIf dgv Is DGV13 Then
            dgvNum = 4
        End If

        If CheckBox24.Checked And (dgv Is DGV7 Or dgv Is DGV8) Then
            '複数個口マーク（宅配便）
            For r As Integer = 0 To dgv.RowCount - 1
                Dim mark As String = ""
                Dim binsu As Integer = dgv.Item(dHSel.IndexOf(koumoku("koguchi")(dgvNum)), r).Value
                If binsu > 1 Then
                    'dgv.Item(dHSel.IndexOf(koumoku("bikou2")(dgvNum)), r).Value &= "＃" & binsu & "枚"
                    mark = dgv.Item(dHSel.IndexOf(koumoku("bikou2")(dgvNum)), r).Value & "＃" & binsu & "枚"
                End If
                HinmeiMaeMark(dgv, dHSel, dH1, dgvNum, r, mark)

                'マークに追加すると伝票が先にまとめられて出力になる
                If CheckBox27.Checked Then
                    dgv.Item(dHSel.IndexOf(koumoku("bikou2")(dgvNum)), r).Value = mark
                End If
            Next
        ElseIf CheckBox26.Checked And (dgv Is DGV9 Or dgv Is DGV13) Then
            '複数個口マーク（メール便）
            For r As Integer = 0 To dgv.RowCount - 1
                Dim mark As String = ""
                Dim binsu As Integer = dgv.Item(dHSel.IndexOf("マスタ個口"), r).Value
                If binsu > 1 Then
                    'dgv.Item(dHSel.IndexOf(koumoku("bikou2")(dgvNum)), r).Value &= "＃" & binsu & "枚"
                    mark = dgv.Item(dHSel.IndexOf(koumoku("bikou2")(dgvNum)), r).Value & "＃" & binsu & "枚"
                    HinmeiMaeMark(dgv, dHSel, dH1, dgvNum, r, mark)
                End If

                'マークに追加すると伝票が先にまとめられて出力になる
                If CheckBox28.Checked Then
                    dgv.Item(dHSel.IndexOf(koumoku("bikou2")(dgvNum)), r).Value = mark
                End If
            Next
        ElseIf dgv Is DGV7 Then
            For r As Integer = 0 To dgv.RowCount - 1
                '***************************************************************************
                'dgv1に反映するため、伝票番号取得
                '***************************************************************************
                Dim selDenpyoNo As String = dgv.Item(dHSel.IndexOf(koumoku("denpyoNo")(dgvNum)), r).Value  '伝票番号取得
                Dim dgv1Row As Integer = 0      'dgv1の伝票番号列を取得する
                For r1 As Integer = 0 To DGV1.RowCount - 1
                    If selDenpyoNo = DGV1.Item(dH1.IndexOf("伝票番号"), r1).Value Then
                        dgv1Row = r1
                        Exit For
                    End If
                Next
                '複数個口マークを使用しない時、品名（e飛伝2だけ「ご依頼主住所２」に表示する）
                Dim hinmeiStr As String() = Regex.Split(DGV1.Item(dH1.IndexOf("品名"), dgv1Row).Value, "\s|　")
                dgv.Item(dHSel.IndexOf(koumoku("hinmei")(1)), r).Value = "＜品＞" & hinmeiStr(0)
            Next
        End If

    End Sub

    'マークを品名の前に転記
    Private Sub HinmeiMaeMark(dgv As DataGridView, dHSel As ArrayList, dH1 As ArrayList, dgvNum As Integer, r As Integer, mark As String)
        Dim markStr As String = mark
        If (dgv Is DGV9 Or dgv Is DGV13) And InStr(markStr, "＃") = 0 Then
            Exit Sub
        End If
        If (dgv Is DGV9 Or dgv Is DGV13) And markStr <> "" Then
            markStr = Regex.Replace(markStr, "★●♪", "")
        End If

        Dim denpyoNo As String = dgv.Item(dHSel.IndexOf(koumoku("denpyoNo")(dgvNum)), r).Value
        Dim hinmeiStr As String() = Nothing
        Dim siharai As String = ""


        'dgv1から品名を取得
        For r1 As Integer = 0 To DGV1.RowCount - 1
            If DGV1.Item(dH1.IndexOf("伝票番号"), r1).Value = denpyoNo Then
                hinmeiStr = Regex.Split(DGV1.Item(dH1.IndexOf("品名"), r1).Value, "\s|　")

                '再発送決済分の時マークを付ける
                If DGV1.Item(dH1.IndexOf("支払方法"), r1).Value <> "" Then
                    If Regex.IsMatch(DGV1.Item(dH1.IndexOf("支払方法"), r1).Value, "再発送") Then
                        siharai = "特"
                    End If
                End If
                Exit For
            End If
        Next

        Dim markDefault As String = "品"
        If hinmeiStr(0) <> "" Then
            If markStr <> "" Then
                markDefault = markStr
                dgv.Item(dHSel.IndexOf(koumoku("hinmei")(dgvNum)), r).Value = "＜" & markDefault & siharai & "＞" & hinmeiStr(0)
            Else
                If siharai <> "" Then
                    markDefault = siharai
                End If
                dgv.Item(dHSel.IndexOf(koumoku("hinmei")(dgvNum)), r).Value = "＜" & markDefault & "＞" & hinmeiStr(0)
            End If
        End If
    End Sub

    'メール便で無理変換した際に商品複数別紙がある時は、対象伝票の項目を修正する
    Private Sub HakoCodeCheck(dgv As DataGridView, regHakoCode As String)
        Dim rh As String() = Split(regHakoCode, ",")
        Dim hSettei As String() = Nothing
        Dim num As Integer = 0

        If dgv Is DGV7 Then
            hSettei = Split(rh(1), "/")
            num = 1
        ElseIf dgv Is DGV8 Then
            hSettei = Split(rh(2), "/")
            num = 2
        ElseIf dgv Is DGV9 Then
            hSettei = Split(rh(3), "/")
            num = 3
        ElseIf dgv Is DGV13 Then
            hSettei = Split(rh(4), "/")
            num = 4
        End If
        Dim hSetteiHeader As String() = Split(hSettei(2), "|")

        Dim dHSel As ArrayList = TM_HEADER_GET(dgv)
        Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
        For r As Integer = 0 To dgv.RowCount - 1
            If dgv.Item(dHSel.IndexOf(hSetteiHeader(0)), r).Value = "★商品複数別処理" Then
                If InStr(dgv.Item(dHSel.IndexOf("箱番号"), r).Value, "/") > 0 Then
                    Dim selDenpyoNo As String = dgv.Item(dHSel.IndexOf(koumoku("denpyoNo")(num)), r).Value  '伝票番号取得
                    For r2 As Integer = 0 To dgv.RowCount - 1
                        If dgv.Item(dHSel.IndexOf(koumoku("denpyoNo")(num)), r2).Value = selDenpyoNo Then
                            dgv.Item(dHSel.IndexOf(hSetteiHeader(0)), r2).Value = "★商品複数別処理"
                            dgv.Item(dHSel.IndexOf(hSetteiHeader(1)), r2).Value = ""
                            dgv.Item(dHSel.IndexOf(hSetteiHeader(2)), r2).Value = ""
                            dgv.Item(dHSel.IndexOf(hSetteiHeader(3)), r2).Value = ""
                            dgv.Item(dHSel.IndexOf(koumoku("bikou2")(num)), r2).Value = "★"

                            For r3 As Integer = 0 To DGV1.RowCount - 1
                                If DGV1.Item(dH1.IndexOf(koumoku("denpyoNo")(0)), r3).Value = selDenpyoNo Then
                                    If DGV1.Item(dH1.IndexOf(koumoku("free2")(0)), r3).Value <> "★商品複数別処理" Then
                                        DGV1.Item(dH1.IndexOf(koumoku("free2")(0)), r3).Value = "★商品複数別処理"
                                        DGV1.Item(dH1.IndexOf(koumoku("free3")(0)), r3).Value = ""
                                        DGV1.Item(dH1.IndexOf(koumoku("free4")(0)), r3).Value = ""
                                        DGV1.Item(dH1.IndexOf(koumoku("free5")(0)), r3).Value = ""
                                    End If
                                End If
                            Next
                        End If
                    Next
                End If
            End If
        Next
    End Sub

    Public dataB As String() = {"医院", "組合", "会社", "機構", "法人", "薬局", "センター", "(株)", "（株）", "商店", "店", "支社", "(有)"， "(合)"}
    Private Function IsCompanyOrIndividual_7(r As Integer, dhM As ArrayList)

        '148067700195
        Dim data1， data2， data3, data4
        data1 = DGV7.Item(dhM.IndexOf("お届け先住所１"), r).Value  'お届け先住所
        data2 = DGV7.Item(dhM.IndexOf("お届け先住所２"), r).Value 'お届け先住所（アパートマンション名）
        data3 = DGV7.Item(dhM.IndexOf("お届け先名称１"), r).Value 'お届け先名称１
        data4 = DGV7.Item(dhM.IndexOf("お届け先名称２"), r).Value 'お届け先名称2

        Dim dataC = {data1， data2， data3, data4}


        For i As Integer = 0 To dataC.Length - 1
            If Not IsNothing(dataC(i)) Then
                For j As Integer = 0 To dataB.Length - 1
                    If InStr(dataC(i), dataB(j)) Then
                        Return True
                    End If
                Next
            End If
        Next
        Return False
    End Function



    Private Function IsCompanyOrIndividual_8(r As Integer, dhM As ArrayList)


        '148067700195
        Dim data1， data2， data3, data4
        data1 = DGV8.Item(dhM.IndexOf("お届け先住所１"), r).Value  'お届け先住所
        data2 = DGV8.Item(dhM.IndexOf("お届け先住所２"), r).Value 'お届け先住所（アパートマンション名）
        data3 = DGV8.Item(dhM.IndexOf("お届け先名１"), r).Value 'お届け先名称１
        data4 = DGV8.Item(dhM.IndexOf("お届け先名２"), r).Value 'お届け先名称2

        Dim dataC = {data1， data2， data3, data4}

        For i As Integer = 0 To dataC.Length - 1
            If Not IsNothing(dataC(i)) Then
                For j As Integer = 0 To dataB.Length - 1
                    Dim d = dataC(i)
                    Dim C = dataB(j)
                    Debug.WriteLine(dataC(i))
                    Debug.WriteLine(dataB(j))
                    If InStr(dataC(i), dataB(j)) Then
                        Return True
                    End If
                Next
            End If
        Next
        Return False
    End Function



    Private Function IsCompanyOrIndividual_tms(r As Integer, dhM As ArrayList)


        '148067700195
        Dim data1， data2， data3, data4
        data1 = TMSDGV.Item(dhM.IndexOf("お届け先住所１"), r).Value  'お届け先住所
        data2 = TMSDGV.Item(dhM.IndexOf("お届け先住所２"), r).Value 'お届け先住所（アパートマンション名）
        data3 = TMSDGV.Item(dhM.IndexOf("お届け先名称１"), r).Value 'お届け先名称１
        data4 = TMSDGV.Item(dhM.IndexOf("お届け先名称２"), r).Value 'お届け先名称2

        Dim dataC = {data1， data2， data3, data4}

        For i As Integer = 0 To dataC.Length - 1
            If Not IsNothing(dataC(i)) Then
                For j As Integer = 0 To dataB.Length - 1
                    Dim d = dataC(i)
                    Dim C = dataB(j)
                    Debug.WriteLine(dataC(i))
                    Debug.WriteLine(dataB(j))
                    If InStr(dataC(i), dataB(j)) Then
                        Return True
                    End If
                Next
            End If
        Next
        Return False
    End Function





    Private Sub Kyoseisyori()
        Dim dH3 As ArrayList = TM_HEADER_GET(DGV3)
        Dim dH7 As ArrayList = TM_HEADER_GET(DGV7)
        '佐川bizlog
        Dim dH8 As ArrayList = TM_HEADER_GET(DGV8)
        Dim dH9 As ArrayList = TM_HEADER_GET(DGV9)
        Dim dH13 As ArrayList = TM_HEADER_GET(DGV13)
        Dim tms_dgv As ArrayList = TM_HEADER_GET(TMSDGV)
        'とんよか卸　型番出ない
        If DGV7.RowCount > 0 Then
            For r As Integer = 0 To DGV7.RowCount - 1
                Dim CC = DGV7.Item(dH7.IndexOf("処理用2"), r).Value
                If DGV7.Item(dH7.IndexOf("処理用2"), r).Value = "名古屋" Then
                    'DGV7.Item(dH7.IndexOf("お客様コード"), r).Value = "148067700005"

                    '法人
                    If IsCompanyOrIndividual_7(r, dH7) Then
                        DGV7.Item(dH7.IndexOf("お客様コード"), r).Value = "148067700196"
                    Else
                        DGV7.Item(dH7.IndexOf("お客様コード"), r).Value = "148067700005"
                    End If
                ElseIf DGV7.Item(dH7.IndexOf("処理用2"), r).Value = "太宰府" Then
                    '957769750002  个人     957769750072  法人

                    If IsCompanyOrIndividual_7(r, dH7) Then
                        'DGV7.Item(dH7.IndexOf("お客様コード"), r).Value = "957769750072"
                        DGV7.Item(dH7.IndexOf("お客様コード"), r).Value = "957769750070"
                    Else
                        DGV7.Item(dH7.IndexOf("お客様コード"), r).Value = "957769750002"
                    End If

                Else
                End If
                Dim denpyocount = DGV7.Item(dH7.IndexOf("マスタ個口"), r).Value

                If denpyocount > 4 Then
                    DGV7.Item(dH7.IndexOf("処理用"), r).Value = "TMS"
                End If







                Dim denpyoNum = DGV7.Item(dH7.IndexOf("お客様管理ナンバー"), r).Value
                For r2 As Integer = 0 To DGV3.RowCount - 1
                    If denpyoNum = DGV3.Item(dH3.IndexOf("伝票番号"), r2).Value Then
                        If DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "卸" Then
                            If InStr(DGV7.Item(dH7.IndexOf("お届け先名称１"), r).Value, "夢みつけ隊株式会社") > 0 Then
                                DGV7.Item(dH7.IndexOf("品名１"), r).Value = "★発注書有り"
                            Else
                                'DGV7.Item(dH7.IndexOf("品名１"), r).Value = ""
                            End If
                            DGV7.Item(dH7.IndexOf("品名２"), r).Value = ""
                            DGV7.Item(dH7.IndexOf("品名３"), r).Value = ""
                            DGV7.Item(dH7.IndexOf("品名４"), r).Value = ""
                            DGV7.Item(dH7.IndexOf("品名５"), r).Value = ""
                            DGV7.Item(dH7.IndexOf("ご依頼主電話番号"), r).Value = "092-980-1866"
                            DGV7.Item(dH7.IndexOf("ご依頼主郵便番号"), r).Value = "812-0881"
                            DGV7.Item(dH7.IndexOf("ご依頼主住所１"), r).Value = "福岡県福岡市博多区井相田１－８－"
                            DGV7.Item(dH7.IndexOf("ご依頼主住所２"), r).Value = "33"
                            DGV7.Item(dH7.IndexOf("ご依頼主名称１"), r).Value = "万方商事株式会社"
                            DGV7.Item(dH7.IndexOf("ご依頼主名称２"), r).Value = ""
                            Exit For


                        ElseIf DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "AZFK" Then
                            'If InStr(DGV7.Item(dH7.IndexOf("お届け先名称１"), r).Value, "夢みつけ隊株式会社") > 0 Then
                            '    DGV7.Item(dH7.IndexOf("品名１"), r).Value = "★発注書有り"
                            'Else
                            '    'DGV7.Item(dH7.IndexOf("品名１"), r).Value = ""
                            'End If




                            'DGV7.Item(dH7.IndexOf("品名２"), r).Value = ""
                            'DGV7.Item(dH7.IndexOf("品名３"), r).Value = ""
                            'DGV7.Item(dH7.IndexOf("品名４"), r).Value = ""
                            'DGV7.Item(dH7.IndexOf("品名５"), r).Value = ""





                            DGV7.Item(dH7.IndexOf("ご依頼主電話番号"), r).Value = "092-586-6853"
                            'ご依頼主　郵便番号

                            'ご依頼主郵便番号
                            DGV7.Item(dH7.IndexOf("ご依頼主郵便番号"), r).Value = "812-0881"
                            DGV7.Item(dH7.IndexOf("ご依頼主住所１"), r).Value = "福岡県福岡市博多区井相田2丁目3番43 102号"

                            'DGV7.Item(dH7.IndexOf("ご依頼主住所２"), r).Value = ""
                            DGV7.Item(dH7.IndexOf("ご依頼主名称１"), r).Value = "Amazon FK"
                            'DGV7.Item(dH7.IndexOf("ご依頼主　名称２"), r).Value = ""
                            Exit For


                        ElseIf DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "AZ海東" Then
                            'If InStr(DGV7.Item(dH7.IndexOf("お届け先名称１"), r).Value, "夢みつけ隊株式会社") > 0 Then
                            '    DGV7.Item(dH7.IndexOf("品名１"), r).Value = "★発注書有り"
                            'Else
                            '    'DGV7.Item(dH7.IndexOf("品名１"), r).Value = ""
                            'End If


                            'DGV7.Item(dH7.IndexOf("品名２"), r).Value = ""
                            'DGV7.Item(dH7.IndexOf("品名３"), r).Value = ""
                            'DGV7.Item(dH7.IndexOf("品名４"), r).Value = ""
                            'DGV7.Item(dH7.IndexOf("品名５"), r).Value = ""


                            DGV7.Item(dH7.IndexOf("ご依頼主電話番号"), r).Value = "092-986-5538"
                            'ご依頼主　郵便番号

                            'ご依頼主郵便番号
                            DGV7.Item(dH7.IndexOf("ご依頼主郵便番号"), r).Value = "811-0123"
                            DGV7.Item(dH7.IndexOf("ご依頼主住所１"), r).Value = "福岡県糟屋郡新宮町上府北3-6-3"

                            'DGV7.Item(dH7.IndexOf("ご依頼主住所２"), r).Value = ""
                            DGV7.Item(dH7.IndexOf("ご依頼主名称１"), r).Value = "Amazon 海東"
                            'DGV7.Item(dH7.IndexOf("ご依頼主名称２"), r).Value = ""
                            Exit For




                        ElseIf DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "問屋よか" And (Replace(DGV3.Item(dH3.IndexOf("購入者名"), r2).Value, " ", "") = "山田善晴A" Or Replace(DGV3.Item(dH3.IndexOf("購入者名"), r2).Value, " ", "") = "山田善晴B" Or Replace(DGV3.Item(dH3.IndexOf("購入者名"), r2).Value, " ", "") = "山田善晴C") Then
                            DGV7.Item(dH7.IndexOf("ご依頼主電話番号"), r).Value = "092-577-9205"
                            DGV7.Item(dH7.IndexOf("ご依頼主郵便番号"), r).Value = ""
                            DGV7.Item(dH7.IndexOf("ご依頼主住所１"), r).Value = "福岡県春日市紅葉ケ丘東 3-43"
                            DGV7.Item(dH7.IndexOf("ご依頼主住所２"), r).Value = "SUN紅葉ヶ丘"
                            DGV7.Item(dH7.IndexOf("ご依頼主名称１"), r).Value = "株式会社ハイハイ"
                            DGV7.Item(dH7.IndexOf("ご依頼主名称２"), r).Value = ""

                            If Replace(DGV3.Item(dH3.IndexOf("購入者名"), r2).Value, " ", "") = "山田善晴A" Then
                                DGV7.Item(dH7.IndexOf("品名１"), r).Value = "株式会社　マイウェイ様ご依頼分"
                                DGV7.Item(dH7.IndexOf("品名２"), r).Value = ""
                                DGV7.Item(dH7.IndexOf("品名３"), r).Value = ""
                                DGV7.Item(dH7.IndexOf("品名４"), r).Value = ""
                                DGV7.Item(dH7.IndexOf("品名５"), r).Value = ""
                            ElseIf Replace(DGV3.Item(dH3.IndexOf("購入者名"), r2).Value, " ", "") = "山田善晴B" Then
                                DGV7.Item(dH7.IndexOf("品名１"), r).Value = "（有）ビーサイレンス様ご依頼分"
                                DGV7.Item(dH7.IndexOf("品名２"), r).Value = ""
                                DGV7.Item(dH7.IndexOf("品名３"), r).Value = ""
                                DGV7.Item(dH7.IndexOf("品名４"), r).Value = ""
                                DGV7.Item(dH7.IndexOf("品名５"), r).Value = ""
                            ElseIf Replace(DGV3.Item(dH3.IndexOf("購入者名"), r2).Value, " ", "") = "山田善晴C" Then
                                DGV7.Item(dH7.IndexOf("品名１"), r).Value = "（株）向島自動車用品製作所様ご依"
                                DGV7.Item(dH7.IndexOf("品名２"), r).Value = "頼分"
                                DGV7.Item(dH7.IndexOf("品名３"), r).Value = ""
                                DGV7.Item(dH7.IndexOf("品名４"), r).Value = ""
                                DGV7.Item(dH7.IndexOf("品名５"), r).Value = ""
                            End If
                            Exit For
                        End If
                    End If
                Next
            Next
        End If

        '新任务
        If DGV8.RowCount > 0 Then
            For r As Integer = 0 To DGV8.RowCount - 1
                Dim cc = DGV8.Item(dH8.IndexOf("処理用2"), r).Value
                If DGV8.Item(dH8.IndexOf("処理用2"), r).Value = "名古屋" Then
                    'DGV8.Item(dH8.IndexOf("佐川急便顧客コード"), r).Value = "148067700005"
                    ''148067700196
                    If IsCompanyOrIndividual_8(r, dH8) Then
                        DGV8.Item(dH8.IndexOf("佐川急便顧客コード"), r).Value = "148067700196"
                    Else
                        DGV8.Item(dH8.IndexOf("佐川急便顧客コード"), r).Value = "148067700005"
                    End If
                ElseIf DGV8.Item(dH8.IndexOf("処理用2"), r).Value = "太宰府" Then

                    If IsCompanyOrIndividual_8(r, dH8) Then
                        '957769750002   957769750072
                        'DGV8.Item(dH8.IndexOf("佐川急便顧客コード"), r).Value = "957769750072"
                        DGV8.Item(dH8.IndexOf("佐川急便顧客コード"), r).Value = "957769750070"
                    Else
                        DGV8.Item(dH8.IndexOf("佐川急便顧客コード"), r).Value = "957769750002"
                    End If

                Else
                End If




                Dim denpyoNum = DGV8.Item(dH8.IndexOf("顧客管理番号"), r).Value
                For r2 As Integer = 0 To DGV3.RowCount - 1
                    If denpyoNum = DGV3.Item(dH3.IndexOf("伝票番号"), r2).Value Then
                        If DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "卸" Then
                            DGV8.Item(dH8.IndexOf("代行ご依頼主電話"), r).Value = "092-980-1866"
                            DGV8.Item(dH8.IndexOf("代行ご依頼主郵便番号"), r).Value = "812-0881"
                            DGV8.Item(dH8.IndexOf("代行ご依頼主住所１"), r).Value = "福岡県福岡市博多区井相田１－８－"
                            DGV8.Item(dH8.IndexOf("代行ご依頼主住所２"), r).Value = "33"
                            DGV8.Item(dH8.IndexOf("代行ご依頼主名１"), r).Value = "万方商事株式会社"
                            DGV8.Item(dH8.IndexOf("代行ご依頼主名２"), r).Value = ""
                            DGV8.Item(dH8.IndexOf("代行ご依頼主メールアドレス"), r).Value = ""
                            Exit For


                        ElseIf DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "AZFK" Then
                            'If InStr(DGV7.Item(dH7.IndexOf("お届け先名称１"), r).Value, "夢みつけ隊株式会社") > 0 Then
                            '    DGV7.Item(dH7.IndexOf("品名１"), r).Value = "★発注書有り"
                            'Else
                            '    'DGV7.Item(dH7.IndexOf("品名１"), r).Value = ""
                            'End If




                            'DGV7.Item(dH7.IndexOf("品名２"), r).Value = ""
                            'DGV7.Item(dH7.IndexOf("品名３"), r).Value = ""
                            'DGV7.Item(dH7.IndexOf("品名４"), r).Value = ""
                            'DGV7.Item(dH7.IndexOf("品名５"), r).Value = ""





                            DGV8.Item(dH8.IndexOf("代行ご依頼主電話"), r).Value = "092-586-6853"
                            'ご依頼主　郵便番号

                            'ご依頼主郵便番号
                            DGV8.Item(dH8.IndexOf("代行ご依頼主郵便番号"), r).Value = "812-0881"
                            DGV8.Item(dH8.IndexOf("代行ご依頼主住所１"), r).Value = "福岡県福岡市博多区井相田2丁目3番43 102号"

                            'DGV8.Item(dH8.IndexOf("代行ご依頼主住所２"), r).Value = ""
                            DGV8.Item(dH8.IndexOf("代行ご依頼主名１"), r).Value = "Amazon FK"
                            'DGV8.Item(dH8.IndexOf("代行ご依頼主名２"), r).Value = ""
                            Exit For


                        ElseIf DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "AZ海東" Then
                            'If InStr(DGV7.Item(dH7.IndexOf("お届け先名称１"), r).Value, "夢みつけ隊株式会社") > 0 Then
                            '    DGV7.Item(dH7.IndexOf("品名１"), r).Value = "★発注書有り"
                            'Else
                            '    'DGV7.Item(dH7.IndexOf("品名１"), r).Value = ""
                            'End If


                            'DGV7.Item(dH7.IndexOf("品名２"), r).Value = ""
                            'DGV7.Item(dH7.IndexOf("品名３"), r).Value = ""
                            'DGV7.Item(dH7.IndexOf("品名４"), r).Value = ""
                            'DGV7.Item(dH7.IndexOf("品名５"), r).Value = ""

                            Debug.WriteLine("")

                            DGV8.Item(dH8.IndexOf("代行ご依頼主電話"), r).Value = "092-986-5538"
                            'ご依頼主　郵便番号

                            'ご依頼主郵便番号
                            DGV8.Item(dH8.IndexOf("代行ご依頼主郵便番号"), r).Value = "811-0123"
                            DGV8.Item(dH8.IndexOf("代行ご依頼主住所１"), r).Value = "福岡県糟屋郡新宮町上府北3-6-3"

                            'DGV8.Item(dH8.IndexOf("代行ご依頼主住所２"), r).Value = ""
                            DGV8.Item(dH8.IndexOf("代行ご依頼主名１"), r).Value = "Amazon 海東"
                            'DGV8.Item(dH8.IndexOf("代行ご依頼主名２"), r).Value = ""
                            Exit For


                        ElseIf DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "問屋よか" And (Replace(DGV3.Item(dH3.IndexOf("購入者名"), r2).Value, " ", "") = "山田善晴A" Or Replace(DGV3.Item(dH3.IndexOf("購入者名"), r2).Value, " ", "") = "山田善晴B" Or Replace(DGV3.Item(dH3.IndexOf("購入者名"), r2).Value, " ", "") = "山田善晴C") Then
                            DGV8.Item(dH8.IndexOf("代行ご依頼主電話"), r).Value = "092-577-9205"
                            DGV8.Item(dH8.IndexOf("代行ご依頼主郵便番号"), r).Value = ""
                            DGV8.Item(dH8.IndexOf("代行ご依頼主住所１"), r).Value = "福岡県春日市紅葉ケ丘東 3-43"
                            DGV8.Item(dH8.IndexOf("代行ご依頼主住所２"), r).Value = "SUN紅葉ヶ丘"
                            DGV8.Item(dH8.IndexOf("代行ご依頼主名１"), r).Value = "株式会社ハイハイ"
                            DGV8.Item(dH8.IndexOf("代行ご依頼主名２"), r).Value = ""
                            DGV8.Item(dH8.IndexOf("代行ご依頼主メールアドレス"), r).Value = ""

                            If Replace(DGV3.Item(dH3.IndexOf("購入者名"), r2).Value, " ", "") = "山田善晴A" Then
                                DGV8.Item(dH8.IndexOf("記事欄11"), r).Value = "株式会社　マイウェイ様ご依頼分"
                            ElseIf Replace(DGV3.Item(dH3.IndexOf("購入者名"), r2).Value, " ", "") = "山田善晴B" Then
                                DGV8.Item(dH8.IndexOf("記事欄11"), r).Value = "（有）ビーサイレンス様ご依頼分"
                            ElseIf Replace(DGV3.Item(dH3.IndexOf("購入者名"), r2).Value, " ", "") = "山田善晴C" Then
                                DGV8.Item(dH8.IndexOf("記事欄11"), r).Value = "（株）向島自動車用品製作所様ご依頼分"
                            End If
                            Exit For
                        End If
                    End If
                Next
            Next
        End If

        If DGV9.RowCount > 0 Then
            For r As Integer = 0 To DGV9.RowCount - 1

                Dim denpyoNum = DGV9.Item(dH9.IndexOf("お客様側管理番号"), r).Value
                'If YU2Denpyus.Contains(denpyoNum)    Then
                '    DGV9.Item(dH9.IndexOf("処理用"), r).Value = "ゆう200"
                'End If
                For r2 As Integer = 0 To DGV3.RowCount - 1
                    Dim YU2tenpo = DGV3.Item(dH3.IndexOf("店舗"), r2).Value
                    If YU2Denpyus.Contains(denpyoNum) Then
                        DGV9.Item(dH9.IndexOf("処理用"), r).Value = "ゆう200"
                        DGV9.Item(dH9.IndexOf("お客様指定配送種類"), r).Value = "ゆうパック元払"
                        DGV9.Item(dH9.IndexOf("商品サイズ／厚さ区分"), r).Value = "60"
                    End If

                    If denpyoNum = DGV3.Item(dH3.IndexOf("伝票番号"), r2).Value Then
                        If DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "卸" Then
                            DGV9.Item(dH9.IndexOf("ご依頼主　郵便番号"), r).Value = "812-0881"
                            DGV9.Item(dH9.IndexOf("ご依頼主　電話番号"), r).Value = "092-980-1866"
                            DGV9.Item(dH9.IndexOf("ご依頼主　住所１"), r).Value = "福岡県福岡市博多区井相田１－８－"
                            DGV9.Item(dH9.IndexOf("ご依頼主　住所２"), r).Value = "33"
                            DGV9.Item(dH9.IndexOf("ご依頼主　名称１"), r).Value = "万方商事株式会社"
                            DGV9.Item(dH9.IndexOf("ご依頼主　名称２"), r).Value = ""
                            Exit For


                        ElseIf DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "AZFK" Then
                            'If InStr(DGV7.Item(dH7.IndexOf("お届け先名称１"), r).Value, "夢みつけ隊株式会社") > 0 Then
                            '    DGV7.Item(dH7.IndexOf("品名１"), r).Value = "★発注書有り"
                            'Else
                            '    'DGV7.Item(dH7.IndexOf("品名１"), r).Value = ""
                            'End If




                            'DGV7.Item(dH7.IndexOf("品名２"), r).Value = ""
                            'DGV7.Item(dH7.IndexOf("品名３"), r).Value = ""
                            'DGV7.Item(dH7.IndexOf("品名４"), r).Value = ""
                            'DGV7.Item(dH7.IndexOf("品名５"), r).Value = ""





                            DGV9.Item(dH9.IndexOf("ご依頼主　電話番号"), r).Value = "092-586-6853"
                            'ご依頼主　郵便番号

                            'ご依頼主郵便番号
                            DGV9.Item(dH9.IndexOf("ご依頼主　郵便番号"), r).Value = "812-0881"
                            DGV9.Item(dH9.IndexOf("ご依頼主　住所１"), r).Value = "福岡県福岡市博多区井相田2丁目3番43 102号"

                            'DGV9.Item(dH9.IndexOf("ご依頼主　住所２"), r).Value = ""
                            DGV9.Item(dH9.IndexOf("ご依頼主　名称１"), r).Value = "Amazon FK"
                            'DGV9.Item(dH9.IndexOf("ご依頼主　名称２"), r).Value = ""
                            Exit For

                        ElseIf DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "AZ海東" Then
                            'If InStr(DGV7.Item(dH7.IndexOf("お届け先名称１"), r).Value, "夢みつけ隊株式会社") > 0 Then
                            '    DGV7.Item(dH7.IndexOf("品名１"), r).Value = "★発注書有り"
                            'Else
                            '    'DGV7.Item(dH7.IndexOf("品名１"), r).Value = ""
                            'End If


                            'DGV7.Item(dH7.IndexOf("品名２"), r).Value = ""
                            'DGV7.Item(dH7.IndexOf("品名３"), r).Value = ""
                            'DGV7.Item(dH7.IndexOf("品名４"), r).Value = ""
                            'DGV7.Item(dH7.IndexOf("品名５"), r).Value = ""


                            DGV9.Item(dH9.IndexOf("ご依頼主　電話番号"), r).Value = "092-986-5538"
                            'ご依頼主　郵便番号

                            'ご依頼主郵便番号
                            DGV9.Item(dH9.IndexOf("ご依頼主　郵便番号"), r).Value = "811-0123"
                            DGV9.Item(dH9.IndexOf("ご依頼主　住所１"), r).Value = "福岡県糟屋郡新宮町上府北3-6-3"

                            'DGV9.Item(dH9.IndexOf("ご依頼主　住所２"), r).Value = ""
                            DGV9.Item(dH9.IndexOf("ご依頼主　名称１"), r).Value = "Amazon 海東"
                            'DGV9.Item(dH9.IndexOf("ご依頼主　名称２"), r).Value = ""
                            Exit For






                        ElseIf DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "問屋よか" And (Replace(DGV3.Item(dH3.IndexOf("購入者名"), r2).Value, " ", "") = "山田善晴A" Or Replace(DGV3.Item(dH3.IndexOf("購入者名"), r2).Value, " ", "") = "山田善晴B" Or Replace(DGV3.Item(dH3.IndexOf("購入者名"), r2).Value, " ", "") = "山田善晴C") Then
                            DGV9.Item(dH9.IndexOf("ご依頼主　郵便番号"), r).Value = ""
                            DGV9.Item(dH9.IndexOf("ご依頼主　電話番号"), r).Value = "092-577-9205"
                            DGV9.Item(dH9.IndexOf("ご依頼主　住所１"), r).Value = "福岡県春日市紅葉ケ丘東 3-43"
                            DGV9.Item(dH9.IndexOf("ご依頼主　住所２"), r).Value = "SUN紅葉ヶ丘"
                            DGV9.Item(dH9.IndexOf("ご依頼主　名称１"), r).Value = "株式会社ハイハイ"
                            DGV9.Item(dH9.IndexOf("ご依頼主　名称２"), r).Value = ""
                            Exit For
                        End If
                    End If
                Next
            Next
        End If


        If DGV13.RowCount > 0 Then
            For r As Integer = 0 To DGV13.RowCount - 1
                Dim denpyoNum = DGV13.Item(dH13.IndexOf("お客様側管理番号"), r).Value

                For r2 As Integer = 0 To DGV3.RowCount - 1
                    If denpyoNum = DGV3.Item(dH3.IndexOf("伝票番号"), r2).Value Then
                        If DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "卸" Then
                            DGV13.Item(dH13.IndexOf("ご依頼主　郵便番号"), r).Value = "812-0881"
                            DGV13.Item(dH13.IndexOf("ご依頼主　電話番号"), r).Value = "092-980-1866"
                            DGV13.Item(dH13.IndexOf("ご依頼主　住所１"), r).Value = "福岡県福岡市博多区井相田１－８－"
                            DGV13.Item(dH13.IndexOf("ご依頼主　住所２"), r).Value = "33"
                            DGV13.Item(dH13.IndexOf("ご依頼主　名称１"), r).Value = "万方商事株式会社"
                            DGV13.Item(dH13.IndexOf("ご依頼主　名称２"), r).Value = ""
                            Exit For
                        ElseIf DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "AZFK" Then
                            'If InStr(DGV7.Item(dH7.IndexOf("お届け先名称１"), r).Value, "夢みつけ隊株式会社") > 0 Then
                            '    DGV7.Item(dH7.IndexOf("品名１"), r).Value = "★発注書有り"
                            'Else
                            '    'DGV7.Item(dH7.IndexOf("品名１"), r).Value = ""
                            'End If
                            'DGV7.Item(dH7.IndexOf("品名２"), r).Value = ""
                            'DGV7.Item(dH7.IndexOf("品名３"), r).Value = ""
                            'DGV7.Item(dH7.IndexOf("品名４"), r).Value = ""
                            'DGV7.Item(dH7.IndexOf("品名５"), r).Value = ""
                            DGV13.Item(dH13.IndexOf("ご依頼主　電話番号"), r).Value = "092-586-6853"
                            'ご依頼主　郵便番号

                            'ご依頼主郵便番号
                            DGV13.Item(dH13.IndexOf("ご依頼主　郵便番号"), r).Value = "812-0881"
                            DGV13.Item(dH13.IndexOf("ご依頼主　住所１"), r).Value = "福岡県福岡市博多区井相田2丁目3番43 102号"

                            'DGV13.Item(dH13.IndexOf("ご依頼主　住所２"), r).Value = ""
                            DGV13.Item(dH13.IndexOf("ご依頼主　名称１"), r).Value = "Amazon FK"
                            'DGV13.Item(dH13.IndexOf("ご依頼主　名称２"), r).Value = ""
                            Exit For


                        ElseIf DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "AZ海東" Then
                            'If InStr(DGV7.Item(dH7.IndexOf("お届け先名称１"), r).Value, "夢みつけ隊株式会社") > 0 Then
                            '    DGV7.Item(dH7.IndexOf("品名１"), r).Value = "★発注書有り"
                            'Else
                            '    'DGV7.Item(dH7.IndexOf("品名１"), r).Value = ""
                            'End If


                            'DGV7.Item(dH7.IndexOf("品名２"), r).Value = ""
                            'DGV7.Item(dH7.IndexOf("品名３"), r).Value = ""
                            'DGV7.Item(dH7.IndexOf("品名４"), r).Value = ""
                            'DGV7.Item(dH7.IndexOf("品名５"), r).Value = ""


                            DGV13.Item(dH13.IndexOf("ご依頼主　電話番号"), r).Value = "092-986-5538"
                            'ご依頼主　郵便番号

                            'ご依頼主郵便番号
                            DGV13.Item(dH13.IndexOf("ご依頼主　郵便番号"), r).Value = "811-0123"
                            DGV13.Item(dH13.IndexOf("ご依頼主　住所１"), r).Value = "福岡県糟屋郡新宮町上府北3-6-3"

                            'DGV13.Item(dH13.IndexOf("ご依頼主　住所２"), r).Value = ""
                            DGV13.Item(dH13.IndexOf("ご依頼主　名称１"), r).Value = "Amazon 海東"
                            'DGV13.Item(dH13.IndexOf("ご依頼主　名称２"), r).Value = ""
                            Exit For
                        ElseIf DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "問屋よか" And (Replace(DGV3.Item(dH3.IndexOf("購入者名"), r2).Value, " ", "") = "山田善晴A" Or Replace(DGV3.Item(dH3.IndexOf("購入者名"), r2).Value, " ", "") = "山田善晴B" Or Replace(DGV3.Item(dH3.IndexOf("購入者名"), r2).Value, " ", "") = "山田善晴C") Then
                            DGV13.Item(dH13.IndexOf("ご依頼主　郵便番号"), r).Value = ""
                            DGV13.Item(dH13.IndexOf("ご依頼主　電話番号"), r).Value = "092-577-9205"
                            DGV13.Item(dH13.IndexOf("ご依頼主　住所１"), r).Value = "福岡県春日市紅葉ケ丘東 3-43"
                            DGV13.Item(dH13.IndexOf("ご依頼主　住所２"), r).Value = "SUN紅葉ヶ丘"
                            DGV13.Item(dH13.IndexOf("ご依頼主　名称１"), r).Value = "株式会社ハイハイ"
                            DGV13.Item(dH13.IndexOf("ご依頼主　名称２"), r).Value = ""
                            Exit For
                        End If
                    End If
                Next
            Next
        End If

        '
        '新任务TMS
        If TMSDGV.RowCount > 0 Then



            For r As Integer = 0 To TMSDGV.RowCount - 1
                If TMSDGV.Item(tms_dgv.IndexOf("処理用2"), r).Value = "名古屋" Then
                    'DGV7.Item(dH7.IndexOf("お客様コード"), r).Value = "148067700005"

                    '法人
                    If IsCompanyOrIndividual_tms(r, tms_dgv) Then
                        TMSDGV.Item(tms_dgv.IndexOf("お客様コード"), r).Value = "148067700196"
                    Else
                        TMSDGV.Item(tms_dgv.IndexOf("お客様コード"), r).Value = "148067700005"
                    End If
                ElseIf TMSDGV.Item(tms_dgv.IndexOf("処理用2"), r).Value = "太宰府" Then
                    '957769750002  个人     957769750072  法人
                    If IsCompanyOrIndividual_tms(r, tms_dgv) Then
                        'TMSDGV.Item(tms_dgv.IndexOf("お客様コード"), r).Value = "957769750072"
                        TMSDGV.Item(tms_dgv.IndexOf("お客様コード"), r).Value = "957769750070"

                    Else
                        TMSDGV.Item(tms_dgv.IndexOf("お客様コード"), r).Value = "957769750002"
                    End If

                Else
                End If
                Dim denpyocount = TMSDGV.Item(tms_dgv.IndexOf("マスタ個口"), r).Value

                'If denpyocount > 4 Then
                TMSDGV.Item(tms_dgv.IndexOf("処理用"), r).Value = "TMS"
                'End If
                Dim denpyoNum = TMSDGV.Item(tms_dgv.IndexOf("お客様管理ナンバー"), r).Value
                For r2 As Integer = 0 To DGV3.RowCount - 1

                    Dim cc = DGV3.Item(dH3.IndexOf("店舗"), r2).Value
                    If denpyoNum = DGV3.Item(dH3.IndexOf("伝票番号"), r2).Value Then
                        If DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "卸" Then
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主電話番号"), r).Value = "092-980-1866"
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主郵便番号"), r).Value = "812-0881"
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主住所１"), r).Value = "福岡県福岡市博多区井相田１－８－"
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主住所２"), r).Value = "33"
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主名称１"), r).Value = "万方商事株式会社"
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主名称２"), r).Value = ""
                            Exit For

                        ElseIf DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "AZFK" Then
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主電話番号"), r).Value = "092-586-6853"
                            'ご依頼主　郵便番号
                            'ご依頼主郵便番号
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主郵便番号"), r).Value = "812-0881"
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主住所１"), r).Value = "福岡県福岡市博多区井相田2丁目3番43 102号"
                            'DGV7.Item(dH7.IndexOf("ご依頼主住所２"), r).Value = ""
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主名称１"), r).Value = "Amazon FK"
                            'DGV7.Item(dH7.IndexOf("ご依頼主　名称２"), r).Value = ""
                            Exit For
                        ElseIf DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "AZ海東" Then
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主電話番号"), r).Value = "092-986-5538"
                            'ご依頼主　郵便番号
                            'ご依頼主郵便番号
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主郵便番号"), r).Value = "811-0123"
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主住所１"), r).Value = "福岡県糟屋郡新宮町上府北3-6-3"
                            'DGV7.Item(dH7.IndexOf("ご依頼主住所２"), r).Value = ""
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主名称１"), r).Value = "Amazon 海東"
                            'DGV7.Item(dH7.IndexOf("ご依頼主名称２"), r).Value = ""
                            Exit For


                        ElseIf DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "ラキナイ" Then
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主電話番号"), r).Value = "092-985-0275"
                            'ご依頼主　郵便番号
                            'ご依頼主郵便番号
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主郵便番号"), r).Value = "811-1361"
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主住所１"), r).Value = "福岡県福岡市南区西長住1丁目12番41号501"
                            'DGV7.Item(dH7.IndexOf("ご依頼主住所２"), r).Value = ""
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主名称１"), r).Value = "Yahoo!Lucky9"
                            'DGV7.Item(dH7.IndexOf("ご依頼主名称２"), r).Value = ""
                            Exit For

                        ElseIf DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "FK" Then
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主電話番号"), r).Value = "092-586-6853"
                            'ご依頼主　郵便番号
                            'ご依頼主郵便番号
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主郵便番号"), r).Value = "816-0911"
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主住所１"), r).Value = "福岡県大野城市大城4-13-15"
                            'DGV7.Item(dH7.IndexOf("ご依頼主住所２"), r).Value = ""
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主名称１"), r).Value = "Yahoo!Fkstyle"
                            'DGV7.Item(dH7.IndexOf("ご依頼主名称２"), r).Value = ""
                            Exit For


                        ElseIf DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "ヤフーKT" Then
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主電話番号"), r).Value = "092-986-5538"
                            'ご依頼主　郵便番号
                            'ご依頼主郵便番号
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主郵便番号"), r).Value = "812-0123"
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主住所１"), r).Value = "福岡県糟屋郡新宮町上府北3-6-3"
                            'DGV7.Item(dH7.IndexOf("ご依頼主住所２"), r).Value = ""
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主名称１"), r).Value = "雑貨ショップＫＴ 海東"
                            'DGV7.Item(dH7.IndexOf("ご依頼主名称２"), r).Value = ""
                            Exit For
                        ElseIf DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "あかねY" Then
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主電話番号"), r).Value = "092-980-1866"
                            'ご依頼主　郵便番号
                            'ご依頼主郵便番号
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主郵便番号"), r).Value = "816-0922"
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主住所１"), r).Value = "福岡県大野城市山田2-2-35"
                            'DGV7.Item(dH7.IndexOf("ご依頼主住所２"), r).Value = ""
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主名称１"), r).Value = "Yahoo!あかねAshop"
                            'DGV7.Item(dH7.IndexOf("ご依頼主名称２"), r).Value = ""
                            Exit For


                        ElseIf DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "あかね楽天" Then
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主電話番号"), r).Value = "092-985-0295"
                            'ご依頼主　郵便番号
                            'ご依頼主郵便番号
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主郵便番号"), r).Value = "816-0922"
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主住所１"), r).Value = "福岡県大野城市山田2-2-35"
                            'DGV7.Item(dH7.IndexOf("ご依頼主住所２"), r).Value = ""
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主名称１"), r).Value = "楽天 あかねAshop"
                            'DGV7.Item(dH7.IndexOf("ご依頼主名称２"), r).Value = ""
                            Exit For

                        ElseIf DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "アリス" Then
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主電話番号"), r).Value = "092-985-2056"
                            'ご依頼主　郵便番号
                            'ご依頼主郵便番号
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主郵便番号"), r).Value = "816-0901"
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主住所１"), r).Value = "福岡県大野城市乙金東１丁目2-52-1"
                            'DGV7.Item(dH7.IndexOf("ご依頼主住所２"), r).Value = ""
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主名称１"), r).Value = "楽天　雑貨の国のアリス"
                            'DGV7.Item(dH7.IndexOf("ご依頼主名称２"), r).Value = ""
                            Exit For
                        ElseIf DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "暁" Then
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主電話番号"), r).Value = "092-986-1116"
                            'ご依頼主　郵便番号
                            'ご依頼主郵便番号
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主郵便番号"), r).Value = "812-0881"
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主住所１"), r).Value = "福岡県福岡市博多区井相田1-8-33-203"
                            'DGV7.Item(dH7.IndexOf("ご依頼主住所２"), r).Value = ""
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主名称１"), r).Value = "楽天　通販の暁"
                            'DGV7.Item(dH7.IndexOf("ご依頼主名称２"), r).Value = ""
                            Exit For

                        ElseIf DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "雑貨倉庫" Then
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主電話番号"), r).Value = "012-069-9991"
                            'ご依頼主　郵便番号
                            'ご依頼主郵便番号
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主郵便番号"), r).Value = "812-0881"
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主住所１"), r).Value = "福岡県福岡市博多区井相田1-8-33"
                            'DGV7.Item(dH7.IndexOf("ご依頼主住所２"), r).Value = ""
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主名称１"), r).Value = "Qoo10通販の雑貨倉庫"
                            'DGV7.Item(dH7.IndexOf("ご依頼主名称２"), r).Value = ""
                            Exit For
                        ElseIf DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "クラナビ" Then
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主電話番号"), r).Value = "092-980-1144"
                            'ご依頼主　郵便番号
                            'ご依頼主郵便番号
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主郵便番号"), r).Value = "812-0881"
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主住所１"), r).Value = "福岡県福岡市博多区井相田1-8-33"
                            'DGV7.Item(dH7.IndexOf("ご依頼主住所２"), r).Value = ""
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主名称１"), r).Value = "auPAYマーケット KuraNavi"
                            'DGV7.Item(dH7.IndexOf("ご依頼主名称２"), r).Value = ""
                            Exit For


                        ElseIf DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "YオクKT" Then
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主電話番号"), r).Value = "092-986-5538"
                            'ご依頼主　郵便番号
                            'ご依頼主郵便番号
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主郵便番号"), r).Value = "811-0123"
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主住所１"), r).Value = "福岡県糟屋郡新宮町上府北3丁目6番地3号"
                            'DGV7.Item(dH7.IndexOf("ご依頼主住所２"), r).Value = ""
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主名称１"), r).Value = "雑貨ＫＴ海東（ヤフオク）"
                            'DGV7.Item(dH7.IndexOf("ご依頼主名称２"), r).Value = ""
                            Exit For
                        ElseIf DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "ヤフオク付" Then
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主電話番号"), r).Value = "092-980-1866"
                            'ご依頼主　郵便番号
                            'ご依頼主郵便番号
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主郵便番号"), r).Value = "812-0881"
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主住所１"), r).Value = "福岡県福岡市博多区 井相田1-8-33"
                            'DGV7.Item(dH7.IndexOf("ご依頼主住所２"), r).Value = ""
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主名称１"), r).Value = "べんけい"
                            'DGV7.Item(dH7.IndexOf("ご依頼主名称２"), r).Value = ""
                            Exit For

                        ElseIf DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "トココ" Then
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主電話番号"), r).Value = "092-986-3343"
                            'ご依頼主　郵便番号
                            'ご依頼主郵便番号
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主郵便番号"), r).Value = "811-0123"
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主住所１"), r).Value = "福岡県糟屋郡新宮町上府北3丁目794番23-3F"
                            'DGV7.Item(dH7.IndexOf("ご依頼主住所２"), r).Value = ""
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主名称１"), r).Value = "Amazon通販のトココ"
                            'DGV7.Item(dH7.IndexOf("ご依頼主名称２"), r).Value = ""
                            Exit For

                        ElseIf DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "サラダ" Then
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主電話番号"), r).Value = "18072377678"
                            'ご依頼主　郵便番号
                            'ご依頼主郵便番号
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主郵便番号"), r).Value = "265100"
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主住所１"), r).Value = "Haiyang shi Hushanjie Youzhengju nan 171hao"
                            'DGV7.Item(dH7.IndexOf("ご依頼主住所２"), r).Value = ""
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主名称１"), r).Value = "Amazonサラダ"
                            'DGV7.Item(dH7.IndexOf("ご依頼主名称２"), r).Value = ""
                            Exit For
                        ElseIf DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "フリット" Then
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主電話番号"), r).Value = "13454959238"
                            'ご依頼主　郵便番号
                            'ご依頼主郵便番号
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主郵便番号"), r).Value = "265106"
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主住所１"), r).Value = "Shangdong Sheng　Yantai Shi,Haiyang Shi Chuangxinchuangyedasha Haibing zhonglu 163hao"
                            'DGV7.Item(dH7.IndexOf("ご依頼主住所２"), r).Value = ""
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主名称１"), r).Value = "Amazon Hewflit"
                            'DGV7.Item(dH7.IndexOf("ご依頼主名称２"), r).Value = ""
                            Exit For
                        ElseIf DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "AZアリス" Then
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主電話番号"), r).Value = "092-985-2056"
                            'ご依頼主　郵便番号
                            'ご依頼主郵便番号
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主郵便番号"), r).Value = "8180114"
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主住所１"), r).Value = "福岡県太宰府市大字北谷960番地"
                            'DGV7.Item(dH7.IndexOf("ご依頼主住所２"), r).Value = ""
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主名称１"), r).Value = "Amazon雑貨の国のアリス"
                            'DGV7.Item(dH7.IndexOf("ご依頼主名称２"), r).Value = ""
                            Exit For
                        ElseIf DGV3.Item(dH3.IndexOf("店舗"), r2).Value = "問屋よか" And (Replace(DGV3.Item(dH3.IndexOf("購入者名"), r2).Value, " ", "") = "山田善晴A" Or Replace(DGV3.Item(dH3.IndexOf("購入者名"), r2).Value, " ", "") = "山田善晴B" Or Replace(DGV3.Item(dH3.IndexOf("購入者名"), r2).Value, " ", "") = "山田善晴C") Then
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主電話番号"), r).Value = "092-577-9205"
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主郵便番号"), r).Value = ""
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主住所１"), r).Value = "福岡県春日市紅葉ケ丘東 3-43"
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主住所２"), r).Value = "SUN紅葉ヶ丘"
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主名称１"), r).Value = "株式会社ハイハイ"
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主名称２"), r).Value = ""

                            If Replace(DGV3.Item(dH3.IndexOf("購入者名"), r2).Value, " ", "") = "山田善晴A" Then
                                TMSDGV.Item(tms_dgv.IndexOf("品名１"), r).Value = "株式会社　マイウェイ様ご依頼分"
                                TMSDGV.Item(tms_dgv.IndexOf("品名２"), r).Value = ""
                                TMSDGV.Item(tms_dgv.IndexOf("品名３"), r).Value = ""
                                TMSDGV.Item(tms_dgv.IndexOf("品名４"), r).Value = ""
                                TMSDGV.Item(tms_dgv.IndexOf("品名５"), r).Value = ""
                            ElseIf Replace(DGV3.Item(dH3.IndexOf("購入者名"), r2).Value, " ", "") = "山田善晴B" Then
                                TMSDGV.Item(tms_dgv.IndexOf("品名１"), r).Value = "（有）ビーサイレンス様ご依頼分"
                                TMSDGV.Item(tms_dgv.IndexOf("品名２"), r).Value = ""
                                TMSDGV.Item(tms_dgv.IndexOf("品名３"), r).Value = ""
                                TMSDGV.Item(tms_dgv.IndexOf("品名４"), r).Value = ""
                                TMSDGV.Item(tms_dgv.IndexOf("品名５"), r).Value = ""
                            ElseIf Replace(DGV3.Item(dH3.IndexOf("購入者名"), r2).Value, " ", "") = "山田善晴C" Then
                                TMSDGV.Item(tms_dgv.IndexOf("品名１"), r).Value = "（株）向島自動車用品製作所様ご依"
                                TMSDGV.Item(tms_dgv.IndexOf("品名２"), r).Value = "頼分"
                                TMSDGV.Item(tms_dgv.IndexOf("品名３"), r).Value = ""
                                TMSDGV.Item(tms_dgv.IndexOf("品名４"), r).Value = ""
                                TMSDGV.Item(tms_dgv.IndexOf("品名５"), r).Value = ""
                            End If
                            Exit For
                        Else

                            'TMSDGV.Item(tms_dgv.IndexOf("品名２"), r).Value = ""
                            'TMSDGV.Item(tms_dgv.IndexOf("品名３"), r).Value = ""
                            'TMSDGV.Item(tms_dgv.IndexOf("品名４"), r).Value = ""
                            'TMSDGV.Item(tms_dgv.IndexOf("品名５"), r).Value = ""
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主電話番号"), r).Value = "092-980-1866"
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主郵便番号"), r).Value = "812-0881"
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主住所１"), r).Value = "福岡県福岡市博多区井相田１－８－"
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主住所２"), r).Value = "33"
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主名称１"), r).Value = "万方商事株式会社"
                            TMSDGV.Item(tms_dgv.IndexOf("ご依頼主名称２"), r).Value = ""


                        End If
                    End If
                Next
            Next
        End If


        If False Then
            '別紙により並べ替え
            'If ListBox3.Items.Count > 0 Then
            '    If DGV7.RowCount > 0 Then
            '        For k As Integer = ListBox3.Items.Count - 1 To 0 Step -1
            '            If ListBox3.Items(k) <> "" Then
            '                For r As Integer = 0 To DGV7.RowCount - 1
            '                    If DGV7.Item(dH7.IndexOf("お客様管理ナンバー"), r).Value = ListBox3.Items(k) Then
            '                        If r <> 0 Then
            '                            DGV7.Rows.Insert(0)
            '                            For c As Integer = 0 To DGV7.ColumnCount - 1
            '                                DGV7.Item(c, 0).Value = DGV7.Item(c, r + 1).Value
            '                            Next
            '                            DGV7.Rows.RemoveAt(r + 1)
            '                        End If
            '                        Exit For
            '                    End If
            '                Next
            '            End If
            '        Next
            '    End If
            '    If DGV8.RowCount > 0 Then
            '        For k As Integer = ListBox3.Items.Count - 1 To 0 Step -1
            '            If ListBox3.Items(k) <> "" Then
            '                For r As Integer = 0 To DGV8.RowCount - 1
            '                    If DGV8.Item(dH8.IndexOf("顧客管理番号"), r).Value = ListBox3.Items(k) Then
            '                        If r <> 0 Then
            '                            DGV8.Rows.Insert(0)
            '                            For c As Integer = 0 To DGV8.ColumnCount - 1
            '                                DGV8.Item(c, 0).Value = DGV8.Item(c, r + 1).Value
            '                            Next
            '                            DGV8.Rows.RemoveAt(r + 1)
            '                        End If
            '                        Exit For
            '                    End If
            '                Next
            '            End If
            '        Next
            '    End If
            '    If DGV9.RowCount > 0 Then
            '        For k As Integer = ListBox3.Items.Count - 1 To 0 Step -1
            '            Dim search_code As String = ListBox3.Items(k)
            '            If search_code <> "" Then
            '                Dim count As Integer = 0
            '                For r As Integer = DGV9.RowCount - 1 To 0 Step -1
            '                    If DGV9.Item(dH9.IndexOf("お客様側管理番号"), r).Value = search_code Then
            '                        count += 1
            '                    End If
            '                Next
            '                If count > 0 Then
            '                    If count = 1 Then
            '                        For r As Integer = DGV9.RowCount - 1 To 0 Step -1
            '                            If DGV9.Item(dH9.IndexOf("お客様側管理番号"), r).Value = search_code Then
            '                                If r <> 0 Then
            '                                    DGV9.Rows.Insert(0)
            '                                    For c As Integer = 0 To DGV9.ColumnCount - 1
            '                                        DGV9.Item(c, 0).Value = DGV9.Item(c, r + 1).Value
            '                                    Next
            '                                    DGV9.Rows.RemoveAt(r + 1)
            '                                End If
            '                                Exit For
            '                            End If
            '                        Next
            '                    Else
            '                        For cc As Integer = 0 To count - 1
            '                            For r As Integer = DGV9.RowCount - 1 To 0 Step -1
            '                                If DGV9.Item(dH9.IndexOf("お客様側管理番号"), r).Value = search_code Then
            '                                    If r <> 0 Then
            '                                        DGV9.Rows.Insert(0)
            '                                        For c As Integer = 0 To DGV9.ColumnCount - 1
            '                                            DGV9.Item(c, 0).Value = DGV9.Item(c, r + 1).Value
            '                                        Next
            '                                        DGV9.Rows.RemoveAt(r + 1)
            '                                    End If
            '                                    Exit For
            '                                End If
            '                            Next
            '                        Next
            '                    End If
            '                End If

            '            End If
            '        Next
            '    End If
            '    If DGV13.RowCount > 0 Then
            '        For k As Integer = ListBox3.Items.Count - 1 To 0 Step -1
            '            Dim search_code As String = ListBox3.Items(k)
            '            If search_code <> "" Then
            '                Dim count As Integer = 0
            '                For r As Integer = DGV13.RowCount - 1 To 0 Step -1
            '                    If DGV13.Item(dH13.IndexOf("お客様側管理番号"), r).Value = search_code Then
            '                        count += 1
            '                    End If
            '                Next
            '                If count > 0 Then
            '                    If count = 1 Then
            '                        For r As Integer = DGV13.RowCount - 1 To 0 Step -1
            '                            If DGV13.Item(dH13.IndexOf("お客様側管理番号"), r).Value = search_code Then
            '                                If r <> 0 Then
            '                                    DGV13.Rows.Insert(0)
            '                                    For c As Integer = 0 To DGV13.ColumnCount - 1
            '                                        DGV13.Item(c, 0).Value = DGV13.Item(c, r + 1).Value
            '                                    Next
            '                                    DGV13.Rows.RemoveAt(r + 1)
            '                                End If
            '                                Exit For
            '                            End If
            '                        Next
            '                    Else
            '                        For cc As Integer = 0 To count - 1
            '                            For r As Integer = DGV13.RowCount - 1 To 0 Step -1
            '                                If DGV13.Item(dH13.IndexOf("お客様側管理番号"), r).Value = search_code Then
            '                                    If r <> 0 Then
            '                                        DGV13.Rows.Insert(0)
            '                                        For c As Integer = 0 To DGV13.ColumnCount - 1
            '                                            DGV13.Item(c, 0).Value = DGV13.Item(c, r + 1).Value
            '                                        Next
            '                                        DGV13.Rows.RemoveAt(r + 1)
            '                                    End If
            '                                    Exit For
            '                                End If
            '                            Next
            '                        Next
            '                    End If
            '                End If

            '            End If
            '        Next
            '    End If
            'End If
        End If
    End Sub

    ''' <summary>
    ''' ロケーション・バリエーション追加
    ''' </summary>
    ''' <param name="shouhinCode">商品コード</param>
    '''  ''' <param name="soukoname">倉庫</param>
    ''' <param name="chumonsu">受注数</param>
    ''' <param name="mode">0=丸付き文字にする、1=そのまま</param>
    ''' <param name="locationAdd">ロケーション追加</param>
    ''' <returns></returns>
    ''' 
    Private Function ValiationAdd1(ByVal shouhinCode As String, ByVal soukoname As String, Optional chumonsu As Integer = 1, Optional mode As Integer = 0, Optional locationAdd As Boolean = True)
        Dim res As String = ""
        Dim locationStr_ng As Boolean = False

        'Dim dH6 As ArrayList = TM_HEADER_GET(dgv6)

        'ロケーションコード・バリエーションを検索して入れる
        Dim location As String = ""
        Dim valiation As String = ""
        If CheckBox15.Checked Then
            If CheckBox21.Checked Then
                shouhinCode = Regex.Replace(shouhinCode, "\(\(.*\)\)", "")
            End If
            If dgv6CodeArray.Contains(shouhinCode.ToLower) Then
                Dim k As Integer = dgv6CodeArray.IndexOf(shouhinCode.ToLower)
                Dim locationStr As String = ""
                If soukoname = "名古屋" Then
                    locationStr = DGV6.Item(dH6.IndexOf("ロケーション(名古屋)"), k).Value
                Else
                    locationStr = DGV6.Item(dH6.IndexOf("ロケーション"), k).Value
                End If
                If locationStr = "" Then
                    locationStr_ng = True
                End If
                If locationAdd Then
                    If CheckBox7.Checked And InStr(locationStr, "BS") > 0 Then
                        location = "-"
                    Else
                        location = locationStr & "-"
                    End If
                Else
                    location = "-"
                End If

                If CheckBox22.Checked Then
                    Dim vali As String() = Split(DGV6.Item(dH6.IndexOf("商品名"), k).Value, " @")
                    If vali.Length = 2 Then
                        valiation = " (" & vali(1) & ")"
                    ElseIf vali.Length = 3 Then
                        valiation = " (" & vali(1) & "" & vali(2) & ")"
                    Else
                        valiation = ""
                    End If
                End If
            End If
            If location = "-" Then
                location = ""
            End If
        End If

        '丸付き数字に変換する
        Dim chumonStr As String = chumonsu
        If mode = 0 Then
            chumonStr = MaruMojiConv(chumonsu)
        End If

        If CheckBox22.Checked = False Then
            res = location & shouhinCode & "*" & chumonStr
        Else
            Dim codeStr As String = location & shouhinCode & "*" & chumonStr & valiation
            Dim sjisEnc As Encoding = Encoding.GetEncoding("Shift_JIS") '文字バイト数をカウント
            Dim bb = sjisEnc.GetByteCount(codeStr)
            If sjisEnc.GetByteCount(codeStr) > 35 Then  'ゆうパケット、佐川で変化させる
                codeStr = SubstringByte(codeStr, 0, 35)
            End If
            res = codeStr
        End If

        res = StrConv(res, VbStrConv.Katakana Or VbStrConv.Narrow, &H411)

        'サブコードを大文字にする
        If CheckBox23.Checked = True Then
            res = SubCodeUpper(res, shouhinCode)
        End If

        If locationStr_ng = True Then
            res = res + "|ng"
        Else
            res = res + "|ok"
        End If

        Return res
    End Function

    Private Function GET_Hinmei(shouhinCode As String)
        Dim res As String = ""

        If dgv6CodeArray.Contains(shouhinCode.ToLower) Then
            Dim k As Integer = dgv6CodeArray.IndexOf(shouhinCode.ToLower)
            Dim hinmei As String() = Split(DGV6.Item(dH6.IndexOf("商品名"), k).Value, " ")
            res = hinmei(0)
        End If

        Return res
    End Function

    Private Function MaruMojiConv(chumonsu As String)
        '丸付き数字に変換する
        Select Case chumonsu
            Case 1
                chumonsu = chumonsu
            Case 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19
                chumonsu = Chr(Asc("①") + chumonsu - 1)
                'chumonsu = chumonsu
            Case Else
                chumonsu = "(" & chumonsu & ")"
                'chumonsu = chumonsu
        End Select
        Return chumonsu
    End Function

    Private Function MaruMojiModoshi(chumonsu As String)
        '丸付き数字を通常に戻す
        Select Case chumonsu
            Case 1
                chumonsu = chumonsu
            Case "②"
                chumonsu = 2
            Case "③"
                chumonsu = 3
            Case "④"
                chumonsu = 4
            Case "⑤"
                chumonsu = 5
            Case "⑥"
                chumonsu = 6
            Case "⑦"
                chumonsu = 7
            Case "⑧"
                chumonsu = 8
            Case "⑨"
                chumonsu = 9
            Case "⑩"
                chumonsu = 10
            Case "⑪"
                chumonsu = 11
            Case "⑫"
                chumonsu = 12
            Case "⑬"
                chumonsu = 13
            Case "⑭"
                chumonsu = 14
            Case "⑮"
                chumonsu = 15
            Case "⑯"
                chumonsu = 16
            Case "⑰"
                chumonsu = 17
            Case "⑱"
                chumonsu = 18
            Case "⑲"
                chumonsu = 19
            Case Else
                chumonsu = Regex.Replace(chumonsu, "\(|\)", "")
                'chumonsu = chumonsu
        End Select
        Return chumonsu
    End Function

    'サブコードを大文字にする
    Private Function SubCodeUpper(location As String, code As String) As String
        Dim codeArray As String() = Split(code, "-")
        If codeArray.Length > 1 Then
            Dim concatCode As String = codeArray(0)
            For i As Integer = 1 To codeArray.Length - 1
                concatCode &= "-" & codeArray(i).ToUpper
            Next
            If InStr(code, "*") > 0 Then
                Dim c1 As String() = Split(code, "*")
                code = c1(0)
                Dim c2 As String() = Split(concatCode, "*")
                concatCode = c2(0)
            End If
            location = Replace(location, code, concatCode)
        End If
        SubCodeUpper = location
    End Function

    'メール便用行複製
    Private Sub DenpyoRowIncrease(dgv As DataGridView)
        Dim dHSel As ArrayList = TM_HEADER_GET(dgv)
        For r As Integer = dgv.RowCount - 1 To 0 Step -1
            Dim koguchi As String = dgv.Item(dHSel.IndexOf("複数個口数"), r).Value
            If CInt(koguchi) > 1 Then
                For i As Integer = 0 To CInt(koguchi) - 1
                    If i > 0 Then   'コピー元行は飛ばす
                        dgv.Rows.InsertCopy(r, r + 1)
                        For c As Integer = 0 To dgv.ColumnCount - 1
                            dgv.Item(c, r + 1).Value = dgv.Item(c, r).Value
                        Next
                    End If
                    dgv.Item(dHSel.IndexOf("複数個口数"), r).Value = "1"
                    dgv.Item(dHSel.IndexOf("箱番号"), r).Value = CInt(koguchi) - i & "/" & CInt(koguchi)
                Next
            End If
        Next
    End Sub

    Private Sub MeisaiIrainushi(dgv As DataGridView)
        'datagridviewのヘッダーテキストをコレクションに取り込む
        Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)
        Dim dH3 As ArrayList = TM_HEADER_GET(DGV3)

        'グリッド毎の設定
        Dim flagMoji As String = ""
        Dim headerDenpyoNo As String = ""
        Dim headerIraiArray As String() = Nothing
        Select Case True
            Case dgv Is DGV7
                flagMoji = "//伝票用リスト2"
                headerDenpyoNo = "お客様管理ナンバー"
                headerIraiArray = New String() {"ご依頼主郵便番号", "ご依頼主住所１", "ご依頼主名称１", "ご依頼主住所２", "ご依頼主電話番号"}
            Case dgv Is DGV8
                flagMoji = "//伝票用リスト2"
                headerDenpyoNo = "顧客管理番号"
                headerIraiArray = New String() {"代行ご依頼主郵便番号", "代行ご依頼主住所１", "代行ご依頼主名１", "代行ご依頼主住所２", "代行ご依頼主電話"}
            Case dgv Is DGV9
                flagMoji = "//伝票用リスト"
                headerDenpyoNo = "お客様側管理番号"
                headerIraiArray = New String() {"ご依頼主　郵便番号", "ご依頼主　住所１", "ご依頼主　名称１", "ご依頼主　住所２", "ご依頼主　電話番号"}
            Case dgv Is DGV13
                flagMoji = "//伝票用リスト"
                headerDenpyoNo = "お客様側管理番号"
                headerIraiArray = New String() {"ご依頼主　郵便番号", "ご依頼主　住所１", "ご依頼主　名称１", "ご依頼主　住所２", "ご依頼主　電話番号"}
        End Select
        Dim dHSel As ArrayList = TM_HEADER_GET(dgv)

        '配送用の店舗リスト読み込み
        Dim KaishaArray As New ArrayList
        Dim KaishaLines As String() = File.ReadAllLines(Path.GetDirectoryName(Form1.appPath) & "\伝票用住所.txt", ENC_SJ)
        Dim flag As Integer = 0
        For Each s As String In KaishaLines
            If s = flagMoji Then
                flag = 1
            ElseIf flag = 1 Then
                If s <> "" And InStr(s, ",") > 0 Then
                    KaishaArray.Add(s)
                End If
            End If
        Next

        '購入者電話番号の情報を調べるためにDGV1の伝票番号リストを作成
        Dim dgv1dNoArray As New ArrayList
        Dim dgv1tenpoArray As New ArrayList
        For r As Integer = 0 To DGV1.RowCount - 1
            dgv1dNoArray.Add(DGV1.Item(dH1.IndexOf("伝票番号"), r).Value)
            dgv1tenpoArray.Add(DGV1.Item(dH1.IndexOf("店舗"), r).Value)
        Next

        For r1 As Integer = 0 To dgv.RowCount - 1
            Dim denpyoNum As String = dgv.Item(dHSel.IndexOf(headerDenpyoNo), r1).Value
            Dim orositenpo As String = dgv1tenpoArray(dgv1dNoArray.IndexOf(denpyoNum))  'DGV1の店舗
            If InStr(orositenpo, "卸直送") > 0 Then
                '卸の場合、依頼主情報をここで入れる
                For i As Integer = 0 To KaishaArray.Count - 1                   '会社情報の設定より
                    Dim kArray As String() = Split(KaishaArray(i), ",")
                    If kArray(0) = orositenpo Then
                        dgv.Item(dHSel.IndexOf(headerIraiArray(0)), r1).Value = kArray(1)
                        dgv.Item(dHSel.IndexOf(headerIraiArray(1)), r1).Value = kArray(2)
                        If kArray(3) = "卸直送TL" Then
                            dgv.Item(dHSel.IndexOf(headerIraiArray(2)), r1).Value = "TL"
                        Else
                            dgv.Item(dHSel.IndexOf(headerIraiArray(2)), r1).Value = kArray(3)
                        End If

                        dgv.Item(dHSel.IndexOf(headerIraiArray(4)), r1).Value = kArray(4)
                    End If
                Next
            Else
                For r3 As Integer = 0 To DGV3.RowCount - 1
                    If denpyoNum = DGV3.Item(dH3.IndexOf("伝票番号"), r3).Value Then
                        Dim tenpo As String = DGV3.Item(dH3.IndexOf("店舗"), r3).Value  '店舗
                        For i As Integer = 0 To KaishaArray.Count - 1                   '会社情報の設定より
                            Dim kArray As String() = Split(KaishaArray(i), ",")
                            If kArray(0) = tenpo Then
                                dgv.Item(dHSel.IndexOf(headerIraiArray(0)), r1).Value = kArray(1)
                                dgv.Item(dHSel.IndexOf(headerIraiArray(1)), r1).Value = kArray(2)
                                dgv.Item(dHSel.IndexOf(headerIraiArray(2)), r1).Value = kArray(3)
                                If CStr(dgv.Item(dHSel.IndexOf(headerIraiArray(3)), r1).Value) = "" Then
                                    dgv.Item(dHSel.IndexOf(headerIraiArray(4)), r1).Value = kArray(4)
                                End If
                            End If
                        Next
                    End If
                Next
            End If
        Next
    End Sub


    'テンプレート選択
    Dim dgvT As DataGridView = Nothing
    Dim denpyoSoft As String = ""
    Private Function TemplateSet(hassou As String, souko As String, binsyu As String, bincount As String， isGroupOrderFlag As Boolean) As String()
        Dim TenmlateUse As String() = Nothing
        Dim mode As String = ""
        Select Case True
            Case hassou = "宅配便"
                If souko = HS1.Text Then
                    If binsyu = "航空便" Then
                        mode = CheckBox17.Text
                    Else
                        mode = CheckBox1.Text
                    End If
                    '井相田 只有ehiden
                ElseIf souko = HS2.Text Then
                    If binsyu = "航空便" Then
                        mode = CheckBox18.Text
                    Else
                        mode = CheckBox4.Text
                    End If
                    '名古屋
                ElseIf souko = HS4.Text Then
                    If binsyu = "航空便" Then
                        'ehiden2
                        mode = CheckBox3.Text
                    Else
                        'ehiden2
                        mode = CheckBox32.Text
                    End If
                    'ElseIf souko = HS3.Text Then
                    '        mode = CheckBox3.Text
                End If

            Case hassou = "メール便"
                mode = "メール便"
            Case hassou = "ヤマト"
                mode = "メール便"
            Case hassou = "定形外"
                mode = "定形外"
        End Select


        '新任务TMS@  这里暂时取消掉TMS
        'If bincount > 4 And hassou = "宅配便" And souko = HS1.Text And isGroupOrderFlag Then
        '    mode = "TMS"
        'End If

        If mode = "BIZlogi" Then
            TenmlateUse = Template2
            dgvT = DGV8
            denpyoSoft = "BIZlogi"
        ElseIf mode = "メール便" Then
            TenmlateUse = Template3
            dgvT = DGV9
            If hassou = "メール便" Then
                denpyoSoft = "メール便"
            ElseIf hassou = "ヤマト" Then
                denpyoSoft = "ヤマト"
            End If
        ElseIf mode = "定形外" Then
            TenmlateUse = Template4
            dgvT = DGV13
            denpyoSoft = "定形外"
        ElseIf mode = "TMS" Then
            TenmlateUse = Template5
            dgvT = TMSDGV
            denpyoSoft = "TMS"
        Else
            TenmlateUse = Template1
            dgvT = DGV7
            denpyoSoft = "e飛伝2"
        End If

        Return TenmlateUse
    End Function




    '============================================
    '--------------------------------------------
    '保存  出力
    '保存系
    '--------------------------------------------
    '============================================
    Private serverDir As String = Form1.サーバーToolStripMenuItem.Text & "\denpyoLog\"
    Private lockPath As String = serverDir & "lock.txt"
    Private lockUser As String = ""
    Dim qtm As String = """"  '双引号
    Dim qtm2 As String = """" & ","  '双引号加逗号
    Private Sub ゆうプリ用ロケ順保存ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ゆうプリ用ロケ順保存ToolStripMenuItem.Click
        If DGV1.RowCount = 0 And DGV7.RowCount = 0 And DGV8.RowCount = 0 And DGV9.RowCount = 0 And DGV13.RowCount = 0 Then
            Exit Sub
        ElseIf (TextBox48.Text <> "" And TextBox48.Text <> "中継") Or TextBox39.Text <> "" Then
            Dim DR As DialogResult = MsgBox("変更が処理されていませんがそのまま保存しますか？", MsgBoxStyle.Exclamation Or MsgBoxStyle.OkCancel Or MsgBoxStyle.SystemModal)
            If DR = DialogResult.Cancel Then
                Exit Sub
            End If
        End If


        '実績ロックされているか確認
        If File.Exists(lockPath) Then
            Dim lockstr As String = File.ReadAllText(lockPath, encSJ)
            lockUser = lockstr
            If lockstr <> "" And lockstr <> Form1.ログイン名ToolStripMenuItem.Text Then
                MsgBox(lockstr & "さんが編集中です。しばらくお待ちください。")
                Exit Sub
            End If
        End If
        'If Csv_denpyo3_F_count.Button2.BackColor <> Color.Yellow Then
        '    File.WriteAllText(lockPath, Form1.ログイン名ToolStripMenuItem.Text, encSJ)
        'End If

        Label9.Visible = True

        '初期化
        CheckBox8.Checked = False
        For r As Integer = 0 To DGV1.RowCount - 1
            DGV1.Rows(r).Visible = True
        Next
        Application.DoEvents()

        Dim dgvM As DataGridView() = {DGV1, DGV7, DGV8, DGV9, DGV13， TMSDGV}

        'visibleを戻す
        For i As Integer = 0 To dgvM.Length - 1
            For r As Integer = 0 To dgvM(i).RowCount - 1
                dgvM(i).Rows(r).Visible = True
            Next
        Next

        LIST4VIEW("DGV EndEdit", "system")
        For i As Integer = 0 To dgvM.Length - 1
            dgvM(i).EndEdit()
        Next

        'フォルダ選択
        Dim dlg = New VistaFolderBrowserDialog With {
            .RootFolder = Environment.SpecialFolder.Desktop,
            .Description = "フォルダを選択してください"
        }
        Dim saveDir As String = ""
        If dlg.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
            saveDir = dlg.SelectedPath
        Else
            Exit Sub
        End If

        Dim nameArray As String() = New String() {"GU", "BS", "SD", "NA", "GUK", "BSK", "SDK", "NAK", "YMD", "YMI", "YPK2T", "YPK2J"}
        'Dim nameArray As String() = New String() {"GU", "BS", "SD"}
        Dim fNameArray As String() = New String() {HS1.Text, HS3.Text & "倉庫", HS2.Text, HS4.Text, HS1.Text & "航空便", HS3.Text & "倉庫航空便", HS2.Text & "航空便", HS4.Text & "航空便", HS1.Text & "ヤマト便", HS2.Text & "ヤマト便", HS1.Text & "ゆう2", HS2.Text & "ゆう2"}
        'Dim fNameArray As String() = New String() {HS1.Text, HS3.Text & "倉庫", HS2.Text, HS1.Text & "航空便", HS3.Text & "倉庫航空便", HS2.Text & "航空便"}
        'Dim fNameArray As String() = New String() {HS1.Text, HS3.Text & "倉庫", HS2.Text, HS1.Text & "航空便", HS3.Text & "倉庫航空便", HS2.Text & "航空便"}
        'Dim fNameArray As String() = New String() {HS1.Text, HS3.Text & "倉庫", HS2.Text}

        'ログイン名前取得
        Dim loginName As String = "(" & Replace(Form1.ログイン名ToolStripMenuItem.Text, "さん", "") & ")"

        LIST4VIEW("データ分割", "system")
        Dim FileName As String = ""
        Dim dHfiles As New ArrayList
        Dim saveList As New ArrayList
        Dim arr_yamato_dazaifu As New ArrayList
        Dim arr_yamato_isouda As New ArrayList

        Dim dH7 As ArrayList = TM_HEADER_GET(DGV7)
        Dim dH8 As ArrayList = TM_HEADER_GET(DGV8)
        Dim dH9 As ArrayList = TM_HEADER_GET(DGV9)
        Dim dH13 As ArrayList = TM_HEADER_GET(DGV13)
        Dim tms_dh As ArrayList = TM_HEADER_GET(TMSDGV)

        Dim okurizyo_type = "7" '送り状種類 ネコポス
        Dim syukayoteibi As String = Format(Now, "yyyy/MM/dd")
        Dim cool_kubun As String = "0"
        Dim seikyukyakucode As String = CStr("0929801866")

        For i As Integer = 0 To dgvM.Length - 1
            If dgvM(i).RowCount > 0 Then
                Dim locationSet As New DataSet
                For k As Integer = 0 To nameArray.Length - 1
                    Dim locationTable As New DataTable(nameArray(k))
                    locationTable.Columns.Add("KIJI2")
                    locationTable.Columns.Add("CODE")
                    locationTable.Columns.Add("NAME")
                    locationTable.Columns.Add("STR")
                    locationTable.Columns.Add("KIJI3") '名古屋の順序
                    locationTable.Columns.Add("denpyoNo") '名古屋の順序
                    locationSet.Tables.Add(locationTable)
                Next

                Select Case True
                    Case dgvM(i) Is DGV1
                        FileName = "(印刷しない)元データ"
                        dHfiles = TM_HEADER_GET(DGV1)
                    Case dgvM(i) Is DGV7
                        FileName = "e飛伝2" & loginName
                        dHfiles = TM_HEADER_GET(DGV7)
                    Case dgvM(i) Is DGV8
                        FileName = "BIZlogi" & loginName
                        dHfiles = TM_HEADER_GET(DGV8)
                    Case dgvM(i) Is DGV9
                        FileName = "ゆうパケット" & loginName
                        dHfiles = TM_HEADER_GET(DGV9)
                    Case dgvM(i) Is DGV13
                        FileName = "定形外" & loginName
                        dHfiles = TM_HEADER_GET(DGV13)
                    Case dgvM(i) Is TMSDGV
                        FileName = "TMS" & loginName
                        dHfiles = TM_HEADER_GET(TMSDGV)
                End Select

                'ヘッダー行作成
                Dim header As String = ""
                For c As Integer = 0 To dgvM(i).ColumnCount - 1
                    If CStr(header) = "" Then
                        header = """" & dgvM(i).Columns(c).HeaderText & """"
                    Else
                        header &= "," & """" & dgvM(i).Columns(c).HeaderText & """"
                    End If
                Next

                For r As Integer = 0 To dgvM(i).RowCount - 1



                    '並び替え用データ作成
                    Dim kiji2 As String = ""
                    Dim code As String = ""
                    Dim master As String = ""
                    Dim name As String = ""
                    Dim kiji3 As String = ""
                    'denpyoNo 太宰府别纸 按受注票号顺序出  已经废弃 别纸顺序按照其他方式实现了
                    Dim denpyoNo As String = ""
                    If koumoku("bikou2")(i) <> "" Then
                        kiji2 = dgvM(i).Item(TM_ArIndexof(dHfiles, koumoku("bikou2")(i)), r).Value
                    End If
                    If koumoku("free2")(i) <> "" Then
                        code = dgvM(i).Item(TM_ArIndexof(dHfiles, koumoku("free2")(i)), r).Value
                    End If
                    If koumoku("masterCode")(i) <> "" Then
                        master = dgvM(i).Item(TM_ArIndexof(dHfiles, koumoku("masterCode")(i)), r).Value
                    End If
                    If koumoku("sakiname")(i) <> "" Then
                        name = dgvM(i).Item(TM_ArIndexof(dHfiles, koumoku("sakiname")(i)), r).Value
                    End If
                    If koumoku("denpyoNo")(i) <> "" Then
                        denpyoNo = dgvM(i).Item(TM_ArIndexof(dHfiles, koumoku("denpyoNo")(i)), r).Value
                    End If

                    If dgvM(i) Is DGV8 Then 'bizlogi
                        kiji3 = dgvM(i).Item(dH8.IndexOf("記事欄02"), r).Value
                    End If

                    '元データはチェックかけない
                    Dim line As String = ""
                    Dim PlaceFlag As String = ""
                    Dim KoukuuFlag As String = ""

                    '禁則文字チェック
                    Dim sStr As String = ""
                    Dim sFlag As Boolean = False
                    For c As Integer = 0 To dgvM(i).ColumnCount - 1
                        Dim cellStr As String = dgvM(i).Item(c, r).Value
                        If cellStr <> "" Then
                            cellStr = cellStr.Replace(Chr(13), "").Replace(Chr(10), "")     'cell内改行削除
                        End If

                        If CStr(line) = "" Then
                            If CStr(cellStr) = "" Then
                                line = """" & "" & """"
                            Else
                                line = """" & CheckCSVkinsoku(cellStr) & """"  'CSV禁則文字チェック
                            End If
                        Else
                            Select Case c
                                Case TM_ArIndexof(dHfiles, koumoku("masterCode")(i)), TM_ArIndexof(dHfiles, koumoku("masterBikou")(i))
                                    If cellStr <> "" Then
                                        If CheckBox29.Checked Then
                                            'マスタコードを50bytで削除する
                                            '但し、物流倉庫などで出荷する場合にマスタコードを削除すると
                                            'ピッキングリストが出せない。
                                            '2019/07/04時点、ad022だけなら50byt超えないかも
                                            'cellStr = VBStrings.MidB(cellStr, 1, 50)
                                        End If
                                    End If
                                Case TM_ArIndexof(dHfiles, koumoku("free2")(i)), TM_ArIndexof(dHfiles, koumoku("free3")(i)), TM_ArIndexof(dHfiles, koumoku("free4")(i)), TM_ArIndexof(dHfiles, koumoku("free5")(i))
                                    '物流BIZlogiで記事欄32byt設定
                                    Dim souko As String = dgvM(i).Item(TM_ArIndexof(dHfiles, koumoku("syori2")(i)), r).Value
                                    If InStr(souko, HS3.Text) > 0 Then
                                        If cellStr <> "" Then
                                            cellStr = VBStrings.MidB(cellStr, 1, 32)
                                        End If
                                    End If
                                Case TM_ArIndexof(dHfiles, koumoku("siteiJikan")(i))
                                    If cellStr <> "" Then   '物流BIZlogiの時間指定は01、04等の指定
                                        If CheckBox30.Checked Then
                                            Dim souko As String = dgvM(i).Item(TM_ArIndexof(dHfiles, koumoku("syori2")(i)), r).Value
                                            If InStr(souko, HS3.Text) > 0 Then
                                                cellStr = StrToJikan(cellStr)
                                            End If
                                        End If
                                    End If
                                    'Case TM_ArIndexof(dHfiles, koumoku("siteiJikan")(i))

                            End Select
                            If CStr(cellStr) = "" Then
                                line &= "," & """" & "" & """"
                            Else
                                line &= "," & """" & CheckCSVkinsoku(cellStr) & """"    'CSV禁則文字チェック
                            End If
                        End If
                    Next

                    If dgvM(i) Is DGV1 Then
                        PlaceFlag = "GU"
                    ElseIf dgvM(i) IsNot DGV1 Then
                        '新宮・物流・井相田分割
                        If CheckBox20.Checked Then
                            Dim one As String() = Split(master, "、")
                            For k As Integer = 0 To one.Length - 1
                                If one(k) <> "" And InStr(one(k), "*") > 0 Then
                                    'Dim onecode As String() = Split(one(k), "*")
                                    'Dim checkCode As String = ValiationAdd1(onecode(0),,, False)
                                    Dim souko As String = dgvM(i).Item(TM_ArIndexof(dHfiles, "処理用2"), r).Value
                                    Dim isyamato As Boolean = False
                                    If dgvM(i) Is DGV9 Then
                                        If dgvM(i).Item(TM_ArIndexof(dHfiles, "マスタ配送"), r).Value = "ヤマト(陸便)" Or dgvM(i).Item(TM_ArIndexof(dHfiles, "マスタ配送"), r).Value = "ヤマト(船便)" Then
                                            isyamato = True
                                        End If
                                    End If

                                    If isyamato Then
                                        If InStr(souko, HS1.Text) > 0 Then
                                            PlaceFlag = "YMD"
                                        ElseIf InStr(souko, HS2.Text) > 0 Then
                                            PlaceFlag = "YMI"
                                        End If
                                        Exit For
                                    Else
                                        If InStr(souko, HS3.Text) > 0 Then
                                            PlaceFlag = "BS"
                                            Exit For
                                        ElseIf InStr(souko, HS2.Text) > 0 Then
                                            PlaceFlag = "SD"
                                            Exit For
                                        ElseIf InStr(souko, HS4.Text) > 0 Then
                                            PlaceFlag = "NA"
                                            Exit For
                                        Else
                                            PlaceFlag = "GU"
                                        End If
                                    End If
                                End If
                            Next
                        End If

                        '当用户需要处理
                        If CheckBox34.Checked Then
                            Dim one As String() = Split(master, "、")
                            For k As Integer = 0 To one.Length - 1
                                If one(k) <> "" And InStr(one(k), "*") > 0 Then
                                    'Dim onecode As String() = Split(one(k), "*")
                                    'Dim checkCode As String = ValiationAdd1(onecode(0),,, False)
                                    Dim souko As String = dgvM(i).Item(TM_ArIndexof(dHfiles, "処理用2"), r).Value
                                    Dim YU2FlagPro As String = dgvM(i).Item(TM_ArIndexof(dHfiles, "処理用"), r).Value
                                    Dim YU2Flag = False
                                    If YU2FlagPro = "ゆう200" Then
                                        YU2Flag = True
                                    End If

                                    'If isyupakuGoodBool And YU2Flag Then
                                    If YU2Flag Then
                                        '目前只需要处理太宰府  HS1.Tex 太宰府
                                        If InStr(souko, HS1.Text) > 0 Then
                                            PlaceFlag = "YPK2T"
                                            'HS2.Tex 井相田
                                        ElseIf InStr(souko, HS2.Text) > 0 Then
                                            PlaceFlag = "YPK2J"
                                        Else
                                            PlaceFlag = "YPK2T"
                                        End If
                                        Exit For
                                        'Else
                                        '    If InStr(souko, HS3.Text) > 0 Then
                                        '        PlaceFlag = "BS"
                                        '        Exit For
                                        '    ElseIf InStr(souko, HS2.Text) > 0 Then
                                        '        PlaceFlag = "SD"
                                        '        Exit For
                                        '    ElseIf InStr(souko, HS4.Text) > 0 Then
                                        '        PlaceFlag = "NA"
                                        '        Exit For
                                        '    Else
                                        '        PlaceFlag = "GU"
                                        '    End If
                                    End If
                                End If
                            Next
                        End If


                        '航空便分割
                        If CheckBox25.Checked Then
                            Select Case True
                                Case Regex.IsMatch(FileName, "e飛伝2|BIZlogi")
                                    Dim binsyu As String = dgvM(i).Item(TM_ArIndexof(dHfiles, koumoku("binsyu")(i)), r).Value
                                    Select Case binsyu
                                        Case "003", "030"
                                            KoukuuFlag = "K"
                                        Case Else
                                            KoukuuFlag = ""
                                    End Select
                                Case Else
                                    '日本郵便は使用していない
                                    KoukuuFlag = ""
                            End Select
                        End If
                    End If

                    'データ分割
                    If CheckBox20.Checked = True Or CheckBox25.Checked = True Then
                        Dim flagConcat As String = PlaceFlag & KoukuuFlag
                        locationSet.Tables(flagConcat).Rows.Add(kiji2, code, name, line, kiji3, denpyoNo)
                    Else
                        locationSet.Tables(nameArray(0)).Rows.Add(kiji2, code, name, line, kiji3, denpyoNo)  'GU
                    End If


                    Debug.WriteLine(line)

                Next

                Dim saveName As String = ""
                'ロケーション順変更し、上書き防止チェック保存
                For k As Integer = 0 To nameArray.Length - 1
                    Dim tb As DataTable = locationSet.Tables(nameArray(k))
                    If tb.Rows.Count > 0 Then
                        'KIJI2, CODE, NAME, STR, KIJI3
                        'Dim view As New DataView(tb) With {
                        '    .Sort = "KIJI2 DESC, CODE ASC, NAME ASC"
                        '}
                        Dim view As New DataView(tb)

                        If nameArray(k) = "NA" And dgvM(i) Is DGV8 Then
                            view.Sort = "KIJI3"
                        Else
                            view.Sort = "KIJI2 DESC, CODE ASC, NAME ASC"
                        End If

                        Dim str As String = ""

                        Dim yamato_header As String = ""





                        '是佐川的情况下 且 没有别纸的  后期确认一下是不是需要加一个仓库条件   false 注释掉
                        If (dgvM(i) Is DGV8 Or dgvM(i) Is DGV7) And ListBox3.Items.Count = 0 And False Then


                            str = header & vbCrLf
                            For Each row As DataRowView In view
                                str &= row("STR") & vbCrLf
                            Next


                            Dim dstName As String = saveDir & "\★" & FileName & ".csv"
                            If InStr(dstName, "印刷しない") > 0 Then
                                dstName = Replace(dstName, "★", "")
                            End If
                            saveName = dstName
                            If dgvM(i) IsNot DGV1 Then
                                saveName = Replace(dstName, ".csv", "_" & fNameArray(k) & "_" & Format(Now, "yyyyMMddHHmmss") & ".csv")
                            End If
                            If File.Exists(saveName) Then
                                Dim sl As String = Path.GetFileName(saveName) & "→"
                                saveName = Replace(dstName, ".csv", "_" & Format(Now(), "HHmmss") & ".csv")
                                sl &= Path.GetFileName(saveName)
                                saveList.Add(sl)
                            Else
                                saveList.Add(Path.GetFileName(saveName))
                            End If

                            File.WriteAllText(saveName, str, ENC_SJ)

                            'false 注释掉
                        ElseIf (dgvM(i) Is DGV8 Or dgvM(i) Is DGV7) And ListBox3.Items.Count > 0 And False Then



                            For index As Integer = 0 To 1
                                str = header & vbCrLf
                                Dim count As Integer '计数

                                If index = 0 Then
                                    count = 0
                                    For Each row As DataRowView In view
                                        Dim dataRow As String() = row("STR").ToString.Split(",")
                                        For c As Integer = 0 To dataRow.Count - 1
                                            dataRow(c) = dataRow(c).Replace("""", "")
                                        Next

                                        Dim denpyouno As String = ""
                                        If dgvM(i) Is DGV7 Then
                                            denpyouno = dataRow(dH7.IndexOf("お客様管理ナンバー"))
                                        ElseIf dgvM(i) Is DGV8 Then
                                            denpyouno = dataRow(dH8.IndexOf("顧客管理番号"))
                                        ElseIf dgvM(i) Is DGV9 Then
                                            denpyouno = dataRow(dH9.IndexOf("お客様側管理番号"))
                                        ElseIf dgvM(i) Is DGV13 Then
                                            denpyouno = dataRow(dH13.IndexOf("お客様側管理番号"))
                                        ElseIf dgvM(i) Is TMSDGV Then
                                            denpyouno = dataRow(tms_dh.IndexOf("お客様管理ナンバー"))
                                        End If

                                        If (ListBox3.Items.Contains(denpyouno) = False) Or (denpyouno = "") Then '不包含或者可能空的话
                                            str &= row("STR") & vbCrLf
                                            count = count + 1
                                        End If
                                    Next
                                    If count > 0 Then '有数据就出力
                                        Dim dstName As String = saveDir & "\★" & FileName & ".csv"
                                        saveName = dstName
                                        saveName = Replace(dstName, ".csv", "_" & fNameArray(k) & "_" & Format(Now, "yyyyMMddHHmmss") & ".csv")

                                        If File.Exists(saveName) Then
                                            Dim sl As String = Path.GetFileName(saveName) & "→"
                                            saveName = Replace(dstName, ".csv", "_" & Format(Now(), "HHmmss") & ".csv")
                                            sl &= Path.GetFileName(saveName)
                                            saveList.Add(sl)
                                        Else
                                            saveList.Add(Path.GetFileName(saveName))
                                        End If
                                        File.WriteAllText(saveName, str, ENC_SJ)
                                    End If
                                Else
                                    count = 0
                                    For Each row As DataRowView In view
                                        Dim datarow As String() = row("str").ToString.Split(",")
                                        For c As Integer = 0 To datarow.Count - 1
                                            datarow(c) = datarow(c).Replace("""", "")
                                        Next

                                        Dim denpyouno As String = ""
                                        If dgvM(i) Is DGV7 Then
                                            denpyouno = datarow(dH7.IndexOf("お客様管理ナンバー"))
                                        ElseIf dgvM(i) Is DGV8 Then
                                            denpyouno = datarow(dH8.IndexOf("顧客管理番号"))
                                        ElseIf dgvM(i) Is DGV9 Then
                                            denpyouno = datarow(dH9.IndexOf("お客様側管理番号"))
                                        ElseIf dgvM(i) Is DGV13 Then
                                            denpyouno = datarow(dH13.IndexOf("お客様側管理番号"))
                                        ElseIf dgvM(i) Is TMSDGV Then
                                            denpyouno = dataRow(tms_dh.IndexOf("お客様管理ナンバー"))
                                        End If

                                        If ListBox3.Items.Contains(denpyouno) = True Then '不包含或者可能空的话
                                            str &= row("str") & vbCrLf
                                            count = count + 1
                                        End If
                                    Next
                                    If count > 0 Then '有数据就出力
                                        Dim dstname As String = saveDir & "\★" & FileName & ".csv"
                                        saveName = dstname
                                        saveName = Replace(dstname, ".csv", "_" & fNameArray(k) & "(新別紙)" & "_" & Format(Now, "yyyymmddhhmmss") & ".csv")

                                        If File.Exists(saveName) Then
                                            Dim sl As String = Path.GetFileName(saveName) & "→"
                                            saveName = Replace(dstname, ".csv", "_" & Format(Now(), "hhmmss") & ".csv")
                                            sl &= Path.GetFileName(saveName)
                                            saveList.Add(sl)
                                        Else
                                            saveList.Add(Path.GetFileName(saveName))
                                        End If
                                        File.WriteAllText(saveName, str, ENC_SJ)
                                    End If
                                End If
                            Next
                        End If






                        'ヤマト 和 yu2
                        If nameArray(k) = "YMD" Or nameArray(k) = "YMI" Then
                            For c As Integer = 0 To yamato_title_sp.Count - 1
                                If CStr(yamato_header) = "" Then
                                    yamato_header = """" & yamato_title_sp(c) & """"
                                Else
                                    yamato_header &= "," & """" & yamato_title_sp(c) & """"
                                End If
                            Next
                            str = yamato_header & vbCrLf
                            Dim line_yamato As String = ""
                            For Each row As DataRowView In view
                                Dim strtemp As String = ""
                                Dim dataRow As String() = row("STR").ToString.Split(",")

                                For c As Integer = 0 To dataRow.Count - 1
                                    dataRow(c) = dataRow(c).Replace("""", "")
                                Next
                                line_yamato = qtm & dataRow(dH9.IndexOf("お客様側管理番号")) & qtm & "," 'お客様管理番号

                                line_yamato &= qtm & okurizyo_type & qtm2
                                line_yamato &= qtm & cool_kubun & qtm2  'クール区分
                                line_yamato &= qtm & qtm2 '伝票番号
                                line_yamato &= qtm & syukayoteibi & qtm2 '出荷予定日

                                'お届け予定（指定）日
                                If dataRow(dH9.IndexOf("配送希望日")) = "" Then
                                    line_yamato &= qtm & qtm2
                                Else
                                    If IsNumeric(dataRow(dH9.IndexOf("配送希望日"))) Then
                                        line_yamato &= qtm & dataRow(dH9.IndexOf("配送希望日")).Substring(0, 4) & "/" & dataRow(dH9.IndexOf("配送希望日")).Substring(4, 2) & "/" & dataRow(dH9.IndexOf("配送希望日")).Substring(6, 2) & qtm2
                                    Else
                                        line_yamato &= qtm & qtm2
                                    End If
                                End If

                                '配達時間帯
                                '0812: 午前中
                                '1416: 14～16時
                                '1618: 16～18時
                                '1820: 18～20時
                                '1921: 19～21時
                                If dataRow(dH9.IndexOf("配送希望時間帯")) = "午前中" Then
                                    line_yamato &= qtm & "0812" & qtm2
                                ElseIf dataRow(dH9.IndexOf("配送希望時間帯")) = "12～14時" Or dataRow(dH9.IndexOf("配送希望時間帯")) = "14～16時" Then
                                    line_yamato &= qtm & "1416" & qtm2
                                ElseIf dataRow(dH9.IndexOf("配送希望時間帯")) = "16～18時" Then
                                    line_yamato &= qtm & "1618" & qtm2
                                ElseIf dataRow(dH9.IndexOf("配送希望時間帯")) = "18～20時" Then
                                    line_yamato &= qtm & "1820" & qtm2
                                ElseIf dataRow(dH9.IndexOf("配送希望時間帯")) = "19～21時" Or dataRow(dH9.IndexOf("配送希望時間帯")) = "20～21時" Then
                                    line_yamato &= qtm & "1921" & qtm2
                                Else
                                    line_yamato &= qtm & qtm2
                                End If

                                line_yamato &= qtm & qtm2 'お届け先コード
                                line_yamato &= qtm & dataRow(dH9.IndexOf("お届け先　電話番号")) & qtm2 'お届け先電話番号
                                line_yamato &= qtm & qtm2 'お届け先電話番号枝番
                                line_yamato &= qtm & dataRow(dH9.IndexOf("お届け先　郵便番号")) & qtm2 'お届け先郵便番号
                                line_yamato &= qtm & dataRow(dH9.IndexOf("お届け先　住所1")) & qtm2 'お届け先住所
                                Dim bbbyamato = dataRow(dH9.IndexOf("お届け先　住所2"))
                                line_yamato &= qtm & dataRow(dH9.IndexOf("お届け先　住所2")) & qtm2 'お届け先住所（アパートマンション名）
                                line_yamato &= qtm & qtm2 'お届け先会社・部門名１
                                line_yamato &= qtm & qtm2 'お届け先会社・部門名２

                                'お届け先名
                                If InStr(dataRow(dH9.IndexOf("お届け先　名称")), ")") Then


                                    Dim todokesakimei As String() = Split(dataRow(dH9.IndexOf("お届け先　名称")), ")")

                                    If todokesakimei.Length > 2 Then

                                        Dim startIndex = dataRow(dH9.IndexOf("お届け先　名称")).IndexOf(")")
                                        Dim ccstr = dataRow(dH9.IndexOf("お届け先　名称")).Substring(0, startIndex + 1)

                                        Dim Strtempc = Split(dataRow(dH9.IndexOf("お届け先　名称")), ccstr)
                                        Strtempc(1).TrimStart(")")
                                        line_yamato &= qtm & Strtempc(1) & qtm2
                                    Else
                                        line_yamato &= qtm & todokesakimei(1) & qtm2
                                    End If

                                Else
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("お届け先　名称")) & qtm2
                                End If

                                line_yamato &= qtm & qtm2 'お届け先名略称カナ
                                line_yamato &= qtm & "様" & qtm2 '敬称
                                Dim cc = dataRow(dH9.IndexOf("ご依頼主　名称１"))

                                If dataRow(dH9.IndexOf("ご依頼主　名称１")) = "auPAYマーケット KuraNavi" Then
                                    line_yamato &= qtm & "au_KuraNavi" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "092-980-1144" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "812-0881" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県福岡市博多区井相田1-8-33-102" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & "au PAY マーケット" & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "Yahoo!FKstyle" Then
                                    line_yamato &= qtm & "Yahoo_Fk" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "092-586-6853" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & "2" & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "816-0911" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県大野城市大城4-13-15" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "楽天 通販の暁" Then
                                    line_yamato &= qtm & "Ra_Akatuki" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "092-986-1116" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "812-0881" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県福岡市博多区井相田1-8-33-203" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "Amazon通販のトココ" Then
                                    line_yamato &= qtm & "Ama_tokoko" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "000-0000-0000" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "816-0922" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県大野城市山田2-2-35" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    'line_yamato &= qtm & "," 'ご依頼主略称カナ
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ

                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "楽天 雑貨の国のアリス" Then
                                    line_yamato &= qtm & "Ra_Alice" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "092-985-2056" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & "1" & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "816-0901" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県大野城市乙金東１－２－５２－１" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "Qoo10通販の雑貨倉庫" Then
                                    line_yamato &= qtm & "Qo_zakka" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "0120-699-991" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "812-0881" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県福岡市博多区井相田1-8-33-203" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "Qoo10福岡通販堂" Then
                                    line_yamato &= qtm & "Qo_tuhan_do" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "092-586-6853" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "812-0881" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県福岡市博多区井相田1-8-33-205" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "万方商事株式会社" Then
                                    line_yamato &= qtm & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "092-980-1866" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "812-0881" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県福岡市博多区井相田1-8-33-101" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "Yahoo!Lucky9" Then
                                    line_yamato &= qtm & "Yahoo_Lucky9" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "092-985-0275" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "812-0881" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県福岡市博多区井相田1-8-33-202" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "べんけい（soing1702）" Then
                                    line_yamato &= qtm & "Yaho_oku" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "000-0000-0000" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "812-0881" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県福岡市博多区井相田２－３－４３－１０２" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "サラダ" Or dataRow(dH9.IndexOf("ご依頼主　名称１")) = "Amazon Sarada" Then
                                    line_yamato &= qtm & "Ama_Sarada" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "000-0000-0000" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "812-0881" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県福岡市博多区井相田1-8-33-2F" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "Amazon Charyee" Then
                                    line_yamato &= qtm & "Ama_Charyee" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "000-0000-0000" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "812-0881" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県福岡市博多区井相田1-8-33-103" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "Yahoo!あかねAshop" Then
                                    line_yamato &= qtm & "Yahoo_Akane" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "092-985-0302" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "816-0922" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県大野城市山田2-2-35" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "楽天 あかねAshop" Then
                                    line_yamato &= qtm & "Ra_Akane" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "092-985-0295" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "816-0922" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県大野城市山田2-2-35" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "Amazon Hewflit" Then
                                    line_yamato &= qtm & "Ama_Hewflit" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "000-0000-0000" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "816-0901" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県大野城市乙金東１－２－５２－１" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "雑貨ショップKT 海東" Then
                                    line_yamato &= qtm & "Yahoo_KT" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "092-986-5538" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "811-0123" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県糟屋郡新宮町上府北3-6-3" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "Amazon 雑貨の国のアリス" Then
                                    line_yamato &= qtm & "a_Alice" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "000-0000-0000" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "816-0901" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県大野城市乙金東１－２－５２－１" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "雑貨KT海東（ヤフオク）" Then
                                    line_yamato &= qtm & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "092-986-5538" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "811-0123" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県糟屋郡新宮町上府北3-6-3" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "Amazon 海東" Then

                                    line_yamato &= qtm & "Ama_kaitou" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "092-986-5538" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "811-0123" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県糟屋郡新宮町上府北3-6-3" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ


                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "Amazon FK" Then

                                    line_yamato &= qtm & "AZFK" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "092-586-6853" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "812-0881" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県福岡市博多区井相田2丁目3番43 102号" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "seiyishi" Then
                                    line_yamato &= qtm & "seiyishi" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "000-0000-0000" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "812-0881" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県福岡市博多区井相田1-8-33-101" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                Else



                                    line_yamato &= qtm & "seiyishi" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "000-0000-0000" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "812-0881" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県福岡市博多区井相田1-8-33-101" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ




                                    'line_yamato &= qtm & "," 'ご依頼主コード
                                    'line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　電話番号")) & qtm2 'ご依頼主電話番号
                                    'line_yamato &= qtm & "," 'ご依頼主電話番号枝番
                                    'line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　郵便番号")) & qtm2 'ご依頼主郵便番号
                                    'line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　住所１")) & qtm2 'ご依頼主住所
                                    'line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　住所２")) & qtm2 'ご依頼主住所（アパートマンション名）
                                    'line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    'line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                End If


                                Dim temp2 = dataRow(dH9.IndexOf("フリー項目２"))
                                Dim temp3 = dataRow(dH9.IndexOf("フリー項目３"))
                                Dim temp4 = dataRow(dH9.IndexOf("フリー項目４"))
                                Dim temp5 = dataRow(dH9.IndexOf("フリー項目５"))
                                Dim n = dataRow(dH9.IndexOf("箱番号"))
                                Dim mastrcode As String = dataRow(dH9.IndexOf("マスタコード"))



                                Dim mastrStr As String() = mastrcode.Split("、")
                                'Dim fushuflag = False

                                'If InStr(temp2, "★") Then
                                '    fushuflag = True
                                'End If

                                'Dim temp2ss As String = Regex.Replace(temp2, "\([^\(]*\)", "")
                                'Dim temp3ss As String = Regex.Replace(temp3, "\([^\(]*\)", "")
                                'Dim temp4ss As String = Regex.Replace(temp4, "\([^\(]*\)", "")
                                'Dim temp5ss As String = Regex.Replace(temp5, "\([^\(]*\)", "")

                                'If fushuflag Then
                                'line_yamato &= qtm & "" & qtm2  '品名コード１
                                'Else
                                line_yamato &= qtm & changeMarumozi(dataRow(dH9.IndexOf("フリー項目２"))) & qtm2  '品名コード１
                                'End If




                                'If fushuflag Then
                                'line_yamato &= qtm & "★全★ " & mastrStr(0) & " " & mastrStr(1) & " " & mastrStr(2) & qtm2
                                'Else
                                line_yamato &= qtm & temp2 & "  " & temp3 & qtm2  '品名１
                                'End If



                                'line_yamato &= qtm & dataRow(dH9.IndexOf("マスタコード")) & qtm2  '品名１

                                line_yamato &= qtm & qtm2  '品名コード２
                                'line_yamato &= qtm & dataRow(dH9.IndexOf("品名")) & qtm2  '品名２

                                'If fushuflag Then
                                '    Dim temp = ""
                                '    For index = 3 To mastrStr.Length - 1
                                '        temp &= mastrStr(index) + "  "
                                '    Next

                                '    line_yamato &= qtm & temp & qtm2
                                'Else
                                line_yamato &= qtm & temp4 & "  " & temp5 & qtm2  '品名２
                                'End If


                                line_yamato &= qtm & dataRow(dH9.IndexOf("フリー項目１")) & "__" & dataRow(dH9.IndexOf("箱番号")) & qtm2  '荷扱い１
                                line_yamato &= qtm & qtm2  '荷扱い２

                                'If dataRow(dH9.IndexOf("マスタ配送")) = "ヤマト(陸便)" Then
                                '    'line_yamato &= qtm & qtm2  '記事
                                '    line_yamato &= qtm & dataRow(dH9.IndexOf("フリー項目１")) & "__" & dataRow(dH9.IndexOf("箱番号")) & qtm2  '記事
                                'Else
                                '    line_yamato &= qtm & dataRow(dH9.IndexOf("フリー項目１")) & "船便" & "__" & dataRow(dH9.IndexOf("箱番号")) & qtm2  '記事
                                'End If
                                line_yamato &= qtm & dataRow(dH9.IndexOf("品名")) & qtm2  '記事

                                line_yamato &= qtm & qtm2  'コレクト代金引換額（税込）
                                line_yamato &= qtm & qtm2  'コレクト内消費税額等
                                line_yamato &= qtm & "0" & qtm2 '営業所止置き
                                line_yamato &= qtm & qtm2  '営業所コード
                                line_yamato &= qtm & "1" & qtm2 '発行枚数
                                line_yamato &= qtm & qtm2  '個数口枠の印字
                                line_yamato &= qtm & seikyukyakucode & qtm2  'ご請求先顧客コード
                                line_yamato &= qtm & qtm2  'ご請求先分類コード
                                line_yamato &= qtm & "01" & qtm2 '運賃管理番号
                                line_yamato &= qtm & "0" & qtm2 'クロネコwebコレクトデータ登録
                                line_yamato &= qtm & qtm2  'クロネコwebコレクト加盟店番号
                                line_yamato &= qtm & qtm2  'クロネコwebコレクト申込受付番号１
                                line_yamato &= qtm & qtm2  'クロネコwebコレクト申込受付番号２
                                line_yamato &= qtm & qtm2  'クロネコwebコレクト申込受付番号３
                                line_yamato &= qtm & "0" & qtm2 'お届け予定ｅメール利用区分
                                line_yamato &= qtm & qtm2  'お届け予定ｅメールe-mailアドレス
                                line_yamato &= qtm & qtm2  '入力機種
                                line_yamato &= qtm & qtm2  'お届け予定eメールメッセージ
                                line_yamato &= qtm & "0" & qtm2 'お届け完了eメール利用区分
                                line_yamato &= qtm & qtm2  'お届け完了ｅメールe-mailアドレス
                                line_yamato &= qtm & qtm2  'お届け完了ｅメールメッセージ
                                line_yamato &= qtm & "0" & qtm2 'クロネコ収納代行利用区分
                                line_yamato &= qtm & qtm2  '収納代行決済ＱＲコード印刷
                                line_yamato &= qtm & qtm2  '収納代行請求金額(税込)
                                line_yamato &= qtm & qtm2  '収納代行内消費税額等
                                line_yamato &= qtm & qtm2  '収納代行請求先郵便番号
                                line_yamato &= qtm & qtm2  '収納代行請求先住所
                                line_yamato &= qtm & qtm2  '収納代行請求先住所（アパートマンション名）
                                line_yamato &= qtm & qtm2  '収納代行請求先会社・部門名１
                                line_yamato &= qtm & qtm2  '収納代行請求先会社・部門名２
                                line_yamato &= qtm & qtm2  '収納代行請求先名(漢字)
                                line_yamato &= qtm & qtm2  '収納代行請求先名(カナ)
                                line_yamato &= qtm & qtm2  '収納代行問合せ先名(漢字)
                                line_yamato &= qtm & qtm2  '収納代行問合せ先郵便番号
                                line_yamato &= qtm & qtm2  '収納代行問合せ先住所
                                line_yamato &= qtm & qtm2  '収納代行問合せ先住所（アパートマンション名）
                                line_yamato &= qtm & qtm2  '収納代行問合せ先電話番号
                                line_yamato &= qtm & qtm2  '収納代行管理番号
                                line_yamato &= qtm & qtm2  '収納代行品名
                                line_yamato &= qtm & qtm2  '収納代行備考
                                line_yamato &= qtm & qtm2  '複数口くくりキー
                                line_yamato &= qtm & qtm2  '検索キータイトル１
                                line_yamato &= qtm & qtm2  '検索キー１
                                line_yamato &= qtm & qtm2  '検索キータイトル２
                                line_yamato &= qtm & qtm2  '検索キー２
                                line_yamato &= qtm & qtm2  '検索キータイトル３
                                line_yamato &= qtm & qtm2  '検索キー３
                                line_yamato &= qtm & qtm2  '検索キータイトル４
                                line_yamato &= qtm & qtm2  '検索キー４
                                line_yamato &= qtm & qtm2  '検索キータイトル５
                                line_yamato &= qtm & qtm2  '検索キー５
                                line_yamato &= qtm & qtm2  '予備
                                line_yamato &= qtm & qtm2  '予備
                                line_yamato &= qtm & "0" & qtm2 '投函予定メール利用区分
                                line_yamato &= qtm & qtm2  '投函予定メールe-mailアドレス
                                line_yamato &= qtm & qtm2  '投函予定メールメッセージ
                                line_yamato &= qtm & "0" & qtm2 '投函完了メール（お届け先宛）利用区分
                                line_yamato &= qtm & qtm2  '投函完了メール（お届け先宛）e-mailアドレス
                                line_yamato &= qtm & qtm2  '投函完了メール（お届け先宛）メールメッセージ
                                line_yamato &= qtm & "0" & qtm2 '投函完了メール（ご依頼主宛）利用区分
                                line_yamato &= qtm & qtm2  '投函完了メール（ご依頼主宛）e-mailアドレス
                                line_yamato &= qtm & qtm2  '投函完了メール（ご依頼主宛）メールメッセージ
                                line_yamato &= qtm & qtm2  '連携管理番号
                                line_yamato &= qtm & qtm  '通知メールアドレス
                                line_yamato &= vbCrLf

                                str &= line_yamato

                                Debug.WriteLine(str)                              

                            Next


                            If nameArray(k) = "YMD" Then
                                    saveName = saveDir & "\★ヤマト" & loginName & "_" & HS1.Text & "_" & Format(Now, "yyyyMMddHHmmss") & ".csv"
                                Else
                                    saveName = saveDir & "\★ヤマト" & loginName & "_ " & HS2.Text & "_" & Format(Now, "yyyyMMddHHmmss") & ".csv"
                                End If

                                saveList.Add(Path.GetFileName(saveName))

                                File.WriteAllText(saveName, str, ENC_SJ) '-------------------- 20210312 add ----------------
                                'ElseIf nameArray(k) = "YPK2T" Or nameArray(k) = "YPK2J" And isyupakubyAddress(haisouSaki) And isyupakuGoodBool Then


                                '暂时注释  走yu
                            ElseIf (nameArray(k) = "YPK2T" Or nameArray(k) = "YPK2J") And False Then

                                LIST4VIEW("YPK2T", "start")

                                For c As Integer = 0 To yamato_title_sp.Count - 1
                                    If CStr(yamato_header) = "" Then
                                        yamato_header = """" & yamato_title_sp(c) & """"
                                    Else
                                        yamato_header &= "," & """" & yamato_title_sp(c) & """"
                                    End If
                                Next
                                str = yamato_header & vbCrLf
                                Dim line_yamato As String = ""
                                For Each row As DataRowView In view
                                    Dim dataRow As String() = row("STR").ToString.Split(",")

                                    For c As Integer = 0 To dataRow.Count - 1
                                        dataRow(c) = dataRow(c).Replace("""", "")
                                    Next
                                    line_yamato = qtm & dataRow(dH9.IndexOf("お客様側管理番号")) & qtm & "," 'お客様管理番号

                                    line_yamato &= qtm & okurizyo_type & qtm2
                                    line_yamato &= qtm & cool_kubun & qtm2  'クール区分
                                    line_yamato &= qtm & qtm2 '伝票番号
                                    line_yamato &= qtm & syukayoteibi & qtm2 '出荷予定日

                                    'お届け予定（指定）日
                                    If dataRow(dH9.IndexOf("配送希望日")) = "" Then
                                        line_yamato &= qtm & qtm2
                                    Else
                                        If IsNumeric(dataRow(dH9.IndexOf("配送希望日"))) Then

                                            Dim strnnnn = dataRow(dH9.IndexOf("配送希望日"))


                                            line_yamato &= qtm & dataRow(dH9.IndexOf("配送希望日")).Substring(0, 4) & "/" & dataRow(dH9.IndexOf("配送希望日")).Substring(4, 2) & "/" & dataRow(dH9.IndexOf("配送希望日")).Substring(6, 2) & qtm2
                                        Else
                                            line_yamato &= qtm & qtm2
                                        End If
                                    End If

                                    '配達時間帯
                                    '0812: 午前中
                                    '1416: 14～16時
                                    '1618: 16～18時
                                    '1820: 18～20時
                                    '1921: 19～21時
                                    If dataRow(dH9.IndexOf("配送希望時間帯")) = "午前中" Then
                                        line_yamato &= qtm & "0812" & qtm2
                                    ElseIf dataRow(dH9.IndexOf("配送希望時間帯")) = "12～14時" Or dataRow(dH9.IndexOf("配送希望時間帯")) = "14～16時" Then
                                        line_yamato &= qtm & "1416" & qtm2
                                    ElseIf dataRow(dH9.IndexOf("配送希望時間帯")) = "16～18時" Then
                                        line_yamato &= qtm & "1618" & qtm2
                                    ElseIf dataRow(dH9.IndexOf("配送希望時間帯")) = "18～20時" Then
                                        line_yamato &= qtm & "1820" & qtm2
                                    ElseIf dataRow(dH9.IndexOf("配送希望時間帯")) = "19～21時" Or dataRow(dH9.IndexOf("配送希望時間帯")) = "20～21時" Then
                                        line_yamato &= qtm & "1921" & qtm2
                                    Else
                                        line_yamato &= qtm & qtm2
                                    End If

                                    line_yamato &= qtm & qtm2 'お届け先コード
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("お届け先　電話番号")) & qtm2 'お届け先電話番号
                                    line_yamato &= qtm & qtm2 'お届け先電話番号枝番
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("お届け先　郵便番号")) & qtm2 'お届け先郵便番号
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("お届け先　住所1")) & qtm2 'お届け先住所
                                    Dim caaaaa = dataRow(dH9.IndexOf("お届け先　住所2"))
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("お届け先　住所2")) & qtm2 'お届け先住所（アパートマンション名）
                                    line_yamato &= qtm & qtm2 'お届け先会社・部門名１
                                    line_yamato &= qtm & qtm2 'お届け先会社・部門名２

                                    'お届け先名
                                    If InStr(dataRow(dH9.IndexOf("お届け先　名称")), ")") Then
                                    Dim todokesakimei As String() = Split(dataRow(dH9.IndexOf("お届け先　名称")), ")")

                                    Dim temp = ""
                                    For intemp = 1 To todokesakimei.Length - 1
                                        temp &= todokesakimei(intemp)
                                    Next
                                    line_yamato &= qtm & temp & qtm2
                                Else
                                        line_yamato &= qtm & dataRow(dH9.IndexOf("お届け先　名称")) & qtm2
                                    End If

                                    line_yamato &= qtm & qtm2 'お届け先名略称カナ
                                    line_yamato &= qtm & "様" & qtm2 '敬称

                                    If dataRow(dH9.IndexOf("ご依頼主　名称１")) = "auPAYマーケット KuraNavi" Then
                                        line_yamato &= qtm & "au_KuraNavi" & qtm2 'ご依頼主コード
                                        line_yamato &= qtm & "092-980-1144" & qtm2 'ご依頼主電話番号
                                        line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                        line_yamato &= qtm & "812-0881" & qtm2 'ご依頼主郵便番号
                                        line_yamato &= qtm & "福岡県福岡市博多区井相田1-8-33-102" & qtm2 'ご依頼主住所
                                        line_yamato &= qtm & "au PAY マーケット" & qtm2 'ご依頼主住所（アパートマンション名）
                                        line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                        line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                    ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "Yahoo!FKstyle" Then
                                        line_yamato &= qtm & "Yahoo_Fk" & qtm2 'ご依頼主コード
                                        line_yamato &= qtm & "092-586-6853" & qtm2 'ご依頼主電話番号
                                        line_yamato &= qtm & "2" & qtm2 'ご依頼主電話番号枝番
                                        line_yamato &= qtm & "816-0911" & qtm2 'ご依頼主郵便番号
                                        line_yamato &= qtm & "福岡県大野城市大城4-13-15" & qtm2 'ご依頼主住所
                                        line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                        line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                        line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                    ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "楽天 通販の暁" Then
                                        line_yamato &= qtm & "Ra_Akatuki" & qtm2 'ご依頼主コード
                                        line_yamato &= qtm & "092-986-1116" & qtm2 'ご依頼主電話番号
                                        line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                        line_yamato &= qtm & "812-0881" & qtm2 'ご依頼主郵便番号
                                        line_yamato &= qtm & "福岡県福岡市博多区井相田1-8-33-203" & qtm2 'ご依頼主住所
                                        line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                        line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                        line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                    ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "Amazon通販のトココ" Then
                                        line_yamato &= qtm & "Ama_tokoko" & qtm2 'ご依頼主コード
                                        line_yamato &= qtm & "000-0000-0000" & qtm2 'ご依頼主電話番号
                                        line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                        line_yamato &= qtm & "816-0922" & qtm2 'ご依頼主郵便番号
                                        line_yamato &= qtm & "福岡県大野城市山田2-2-35" & qtm2 'ご依頼主住所
                                        line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                        line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                        line_yamato &= qtm & "," 'ご依頼主略称カナ
                                    ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "楽天 雑貨の国のアリス" Then
                                        line_yamato &= qtm & "Ra_Alice" & qtm2 'ご依頼主コード
                                        line_yamato &= qtm & "092-985-2056" & qtm2 'ご依頼主電話番号
                                        line_yamato &= qtm & "1" & qtm2 'ご依頼主電話番号枝番
                                        line_yamato &= qtm & "816-0901" & qtm2 'ご依頼主郵便番号
                                        line_yamato &= qtm & "福岡県大野城市乙金東１－２－５２－１" & qtm2 'ご依頼主住所
                                        line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                        line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                        line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                    ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "Qoo10通販の雑貨倉庫" Then
                                        line_yamato &= qtm & "Qo_zakka" & qtm2 'ご依頼主コード
                                        line_yamato &= qtm & "0120-699-991" & qtm2 'ご依頼主電話番号
                                        line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                        line_yamato &= qtm & "812-0881" & qtm2 'ご依頼主郵便番号
                                        line_yamato &= qtm & "福岡県福岡市博多区井相田1-8-33-203" & qtm2 'ご依頼主住所
                                        line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                        line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                        line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                    ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "Qoo10福岡通販堂" Then
                                        line_yamato &= qtm & "Qo_tuhan_do" & qtm2 'ご依頼主コード
                                        line_yamato &= qtm & "092-586-6853" & qtm2 'ご依頼主電話番号
                                        line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                        line_yamato &= qtm & "812-0881" & qtm2 'ご依頼主郵便番号
                                        line_yamato &= qtm & "福岡県福岡市博多区井相田1-8-33-205" & qtm2 'ご依頼主住所
                                        line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                        line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                        line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                    ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "万方商事株式会社" Then
                                        line_yamato &= qtm & qtm2 'ご依頼主コード
                                        line_yamato &= qtm & "092-980-1866" & qtm2 'ご依頼主電話番号
                                        line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                        line_yamato &= qtm & "812-0881" & qtm2 'ご依頼主郵便番号
                                        line_yamato &= qtm & "福岡県福岡市博多区井相田1-8-33-101" & qtm2 'ご依頼主住所
                                        line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                        line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                        line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                    ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "Yahoo!Lucky9" Then
                                        line_yamato &= qtm & "Yahoo_Lucky9" & qtm2 'ご依頼主コード
                                        line_yamato &= qtm & "092-985-0275" & qtm2 'ご依頼主電話番号
                                        line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                        line_yamato &= qtm & "812-0881" & qtm2 'ご依頼主郵便番号
                                        line_yamato &= qtm & "福岡県福岡市博多区井相田1-8-33-202" & qtm2 'ご依頼主住所
                                        line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                        line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                        line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                    ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "べんけい（soing1702）" Then
                                        line_yamato &= qtm & "Yaho_oku" & qtm2 'ご依頼主コード
                                        line_yamato &= qtm & "000-0000-0000" & qtm2 'ご依頼主電話番号
                                        line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                        line_yamato &= qtm & "812-0881" & qtm2 'ご依頼主郵便番号
                                        line_yamato &= qtm & "福岡県福岡市博多区井相田２－３－４３－１０２" & qtm2 'ご依頼主住所
                                        line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                        line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                        line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                    ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "サラダ" Or dataRow(dH9.IndexOf("ご依頼主　名称１")) = "Amazon Sarada" Then
                                        line_yamato &= qtm & "Ama_Sarada" & qtm2 'ご依頼主コード
                                        line_yamato &= qtm & "000-0000-0000" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "812-0881" & qtm2 'ご依頼主郵便番号
                                        line_yamato &= qtm & "福岡県福岡市博多区井相田1-8-33-2F" & qtm2 'ご依頼主住所
                                        line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                        line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                        line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                    ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "Amazon Charyee" Then
                                        line_yamato &= qtm & "Ama_Charyee" & qtm2 'ご依頼主コード
                                        line_yamato &= qtm & "000-0000-0000" & qtm2 'ご依頼主電話番号
                                        line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                        line_yamato &= qtm & "812-0881" & qtm2 'ご依頼主郵便番号
                                        line_yamato &= qtm & "福岡県福岡市博多区井相田1-8-33-103" & qtm2 'ご依頼主住所
                                        line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                        line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                        line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                    ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "Yahoo!あかねAshop" Then
                                        line_yamato &= qtm & "Yahoo_Akane" & qtm2 'ご依頼主コード
                                        line_yamato &= qtm & "092-985-0302" & qtm2 'ご依頼主電話番号
                                        line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                        line_yamato &= qtm & "816-0922" & qtm2 'ご依頼主郵便番号
                                        line_yamato &= qtm & "福岡県大野城市山田2-2-35" & qtm2 'ご依頼主住所
                                        line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                        line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                        line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                    ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "楽天 あかねAshop" Then
                                        line_yamato &= qtm & "Ra_Akane" & qtm2 'ご依頼主コード
                                        line_yamato &= qtm & "092-985-0295" & qtm2 'ご依頼主電話番号
                                        line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                        line_yamato &= qtm & "816-0922" & qtm2 'ご依頼主郵便番号
                                        line_yamato &= qtm & "福岡県大野城市山田2-2-35" & qtm2 'ご依頼主住所
                                        line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                        line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                        line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                    ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "Amazon Hewflit" Then
                                        line_yamato &= qtm & "Ama_Hewflit" & qtm2 'ご依頼主コード
                                        line_yamato &= qtm & "000-0000-0000" & qtm2 'ご依頼主電話番号
                                        line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                        line_yamato &= qtm & "816-0901" & qtm2 'ご依頼主郵便番号
                                        line_yamato &= qtm & "福岡県大野城市乙金東１－２－５２－１" & qtm2 'ご依頼主住所
                                        line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                        line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                        line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                    ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "雑貨ショップKT 海東" Then
                                        line_yamato &= qtm & "Yahoo_KT" & qtm2 'ご依頼主コード
                                        line_yamato &= qtm & "092-986-5538" & qtm2 'ご依頼主電話番号
                                        line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                        line_yamato &= qtm & "811-0123" & qtm2 'ご依頼主郵便番号
                                        line_yamato &= qtm & "福岡県糟屋郡新宮町上府北3-6-3" & qtm2 'ご依頼主住所
                                        line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                        line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                        line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                    ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "Amazon 雑貨の国のアリス" Then
                                        line_yamato &= qtm & "a_Alice" & qtm2 'ご依頼主コード
                                        line_yamato &= qtm & "000-0000-0000" & qtm2 'ご依頼主電話番号
                                        line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                        line_yamato &= qtm & "816-0901" & qtm2 'ご依頼主郵便番号
                                        line_yamato &= qtm & "福岡県大野城市乙金東１－２－５２－１" & qtm2 'ご依頼主住所
                                        line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                        line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                        line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                    ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "雑貨KT海東（ヤフオク）" Then
                                        line_yamato &= qtm & qtm2 'ご依頼主コード
                                        line_yamato &= qtm & "092-986-5538" & qtm2 'ご依頼主電話番号
                                        line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                        line_yamato &= qtm & "811-0123" & qtm2 'ご依頼主郵便番号
                                        line_yamato &= qtm & "福岡県糟屋郡新宮町上府北3-6-3" & qtm2 'ご依頼主住所
                                        line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                        line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                        line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                    Else
                                        line_yamato &= qtm & "," 'ご依頼主コード
                                        line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　電話番号")) & qtm2 'ご依頼主電話番号
                                        line_yamato &= qtm & "," 'ご依頼主電話番号枝番
                                        line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　郵便番号")) & qtm2 'ご依頼主郵便番号
                                        line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　住所１")) & qtm2 'ご依頼主住所
                                        line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　住所２")) & qtm2 'ご依頼主住所（アパートマンション名）
                                        line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                        line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                    End If

                                    line_yamato &= qtm & changeMarumozi(dataRow(dH9.IndexOf("フリー項目２"))) & qtm2  '品名コード１
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("フリー項目２")) & "(" & dataRow(dH9.IndexOf("フリー項目３")) & ")" & qtm2  '品名１

                                    line_yamato &= qtm & qtm2  '品名コード２
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("品名")) & qtm2  '品名２
                                    line_yamato &= qtm & qtm2  '荷扱い１
                                    line_yamato &= qtm & qtm2  '荷扱い２

                                    If dataRow(dH9.IndexOf("マスタ配送")) = "ヤマト(陸便)" Then
                                        line_yamato &= qtm & qtm2  '記事
                                    Else
                                        line_yamato &= qtm & "船便" & qtm2  '記事
                                    End If

                                    line_yamato &= qtm & qtm2  'コレクト代金引換額（税込）
                                    line_yamato &= qtm & qtm2  'コレクト内消費税額等
                                    line_yamato &= qtm & "0" & qtm2 '営業所止置き
                                    line_yamato &= qtm & qtm2  '営業所コード
                                    line_yamato &= qtm & "1" & qtm2 '発行枚数
                                    line_yamato &= qtm & qtm2  '個数口枠の印字
                                    line_yamato &= qtm & seikyukyakucode & qtm2  'ご請求先顧客コード
                                    line_yamato &= qtm & qtm2  'ご請求先分類コード
                                    line_yamato &= qtm & "01" & qtm2 '運賃管理番号
                                    line_yamato &= qtm & "0" & qtm2 'クロネコwebコレクトデータ登録
                                    line_yamato &= qtm & qtm2  'クロネコwebコレクト加盟店番号
                                    line_yamato &= qtm & qtm2  'クロネコwebコレクト申込受付番号１
                                    line_yamato &= qtm & qtm2  'クロネコwebコレクト申込受付番号２
                                    line_yamato &= qtm & qtm2  'クロネコwebコレクト申込受付番号３
                                    line_yamato &= qtm & "0" & qtm2 'お届け予定ｅメール利用区分
                                    line_yamato &= qtm & qtm2  'お届け予定ｅメールe-mailアドレス
                                    line_yamato &= qtm & qtm2  '入力機種
                                    line_yamato &= qtm & qtm2  'お届け予定eメールメッセージ
                                    line_yamato &= qtm & "0" & qtm2 'お届け完了eメール利用区分
                                    line_yamato &= qtm & qtm2  'お届け完了ｅメールe-mailアドレス
                                    line_yamato &= qtm & qtm2  'お届け完了ｅメールメッセージ
                                    line_yamato &= qtm & "0" & qtm2 'クロネコ収納代行利用区分
                                    line_yamato &= qtm & qtm2  '収納代行決済ＱＲコード印刷
                                    line_yamato &= qtm & qtm2  '収納代行請求金額(税込)
                                    line_yamato &= qtm & qtm2  '収納代行内消費税額等
                                    line_yamato &= qtm & qtm2  '収納代行請求先郵便番号
                                    line_yamato &= qtm & qtm2  '収納代行請求先住所
                                    line_yamato &= qtm & qtm2  '収納代行請求先住所（アパートマンション名）
                                    line_yamato &= qtm & qtm2  '収納代行請求先会社・部門名１
                                    line_yamato &= qtm & qtm2  '収納代行請求先会社・部門名２
                                    line_yamato &= qtm & qtm2  '収納代行請求先名(漢字)
                                    line_yamato &= qtm & qtm2  '収納代行請求先名(カナ)
                                    line_yamato &= qtm & qtm2  '収納代行問合せ先名(漢字)
                                    line_yamato &= qtm & qtm2  '収納代行問合せ先郵便番号
                                    line_yamato &= qtm & qtm2  '収納代行問合せ先住所
                                    line_yamato &= qtm & qtm2  '収納代行問合せ先住所（アパートマンション名）
                                    line_yamato &= qtm & qtm2  '収納代行問合せ先電話番号
                                    line_yamato &= qtm & qtm2  '収納代行管理番号
                                    line_yamato &= qtm & qtm2  '収納代行品名
                                    line_yamato &= qtm & qtm2  '収納代行備考
                                    line_yamato &= qtm & qtm2  '複数口くくりキー
                                    line_yamato &= qtm & qtm2  '検索キータイトル１
                                    line_yamato &= qtm & qtm2  '検索キー１
                                    line_yamato &= qtm & qtm2  '検索キータイトル２
                                    line_yamato &= qtm & qtm2  '検索キー２
                                    line_yamato &= qtm & qtm2  '検索キータイトル３
                                    line_yamato &= qtm & qtm2  '検索キー３
                                    line_yamato &= qtm & qtm2  '検索キータイトル４
                                    line_yamato &= qtm & qtm2  '検索キー４
                                    line_yamato &= qtm & qtm2  '検索キータイトル５
                                    line_yamato &= qtm & qtm2  '検索キー５
                                    line_yamato &= qtm & qtm2  '予備
                                    line_yamato &= qtm & qtm2  '予備
                                    line_yamato &= qtm & "0" & qtm2 '投函予定メール利用区分
                                    line_yamato &= qtm & qtm2  '投函予定メールe-mailアドレス
                                    line_yamato &= qtm & qtm2  '投函予定メールメッセージ
                                    line_yamato &= qtm & "0" & qtm2 '投函完了メール（お届け先宛）利用区分
                                    line_yamato &= qtm & qtm2  '投函完了メール（お届け先宛）e-mailアドレス
                                    line_yamato &= qtm & qtm2  '投函完了メール（お届け先宛）メールメッセージ
                                    line_yamato &= qtm & "0" & qtm2 '投函完了メール（ご依頼主宛）利用区分
                                    line_yamato &= qtm & qtm2  '投函完了メール（ご依頼主宛）e-mailアドレス
                                    line_yamato &= qtm & qtm2  '投函完了メール（ご依頼主宛）メールメッセージ
                                    line_yamato &= qtm & qtm2  '連携管理番号
                                    line_yamato &= qtm & qtm  '通知メールアドレス
                                    line_yamato &= vbCrLf

                                    str &= line_yamato
                                Next
                                If nameArray(k) = "YPK2J" Then
                                    saveName = saveDir & "\★ゆう2" & loginName & "_" & HS2.Text & "_" & Format(Now, "yyyyMMddHHmmss") & ".csv"
                                Else
                                    saveName = saveDir & "\★ゆう2" & loginName & "_ " & HS1.Text & "_" & Format(Now, "yyyyMMddHHmmss") & ".csv"
                                End If
                                saveList.Add(Path.GetFileName(saveName))
                                File.WriteAllText(saveName, str, ENC_SJ)
                                'End If
                                '这里
                            Else
                                '-------------------- 20210312 改修前 start ----------------
                                'str = header & vbCrLf
                                'For Each row As DataRowView In view
                                '    str &= row("STR") & vbCrLf
                                'Next

                                'Dim dstName As String = saveDir & "\★" & FileName & ".csv"
                                'If InStr(dstName, "印刷しない") > 0 Then
                                '    dstName = Replace(dstName, "★", "")
                                'End If
                                'saveName = dstName
                                'If dgvM(i) IsNot DGV1 Then
                                '    saveName = Replace(dstName, ".csv", "_" & fNameArray(k) & "_" & Format(Now, "yyyyMMddHHmmss") & ".csv")
                                'End If
                                'If File.Exists(saveName) Then
                                '    Dim sl As String = Path.GetFileName(saveName) & "→"
                                '    saveName = Replace(dstName, ".csv", "_" & Format(Now(), "HHmmss") & ".csv")
                                '    sl &= Path.GetFileName(saveName)
                                '    saveList.Add(sl)
                                'Else
                                '    saveList.Add(Path.GetFileName(saveName))
                                'End If
                                '-------------------- 20210312 改修前 end ----------------

                                '-------------------- 20210312 改修后 start ----------------
                                If dgvM(i) Is DGV1 Or (dgvM(i) IsNot DGV1 And CheckBox33.Checked = False) Or (dgvM(i) IsNot DGV1 And CheckBox33.Checked And ListBox3.Items.Count = 0) Then
                                    str = header & vbCrLf
                                For Each row As DataRowView In view
                                    str &= row("STR") & vbCrLf
                                Next
                                Dim dstName As String = saveDir & "\★" & FileName & ".csv"
                                    If InStr(dstName, "印刷しない") > 0 Then
                                        dstName = Replace(dstName, "★", "")
                                    End If
                                    saveName = dstName
                                If dgvM(i) IsNot DGV1 Then
                                    'If (nameArray(k) = "YPK2J" Or nameArray(k) = "YPK2T") Then
                                    '    saveName = Replace(dstName, "ゆうパケット", "ゆうパック")
                                    '    saveName = Replace(saveName, ".csv", "_" & fNameArray(k) & "_" & Format(Now, "yyyyMMddHHmmss") & ".csv")
                                    'Else
                                    saveName = Replace(dstName, ".csv", "_" & fNameArray(k) & "_" & Format(Now, "yyyyMMddHHmmss") & ".csv")
                                    'End If

                                End If

                                    If File.Exists(saveName) Then
                                    Dim sl As String = Path.GetFileName(saveName) & "→"
                                    saveName = Replace(dstName, ".csv", "_" & Format(Now(), "HHmmss") & ".csv")
                                    sl &= Path.GetFileName(saveName)
                                    saveList.Add(sl)
                                Else
                                    saveList.Add(Path.GetFileName(saveName))
                                End If

                                File.WriteAllText(saveName, str, ENC_SJ)
                                ElseIf dgvM(i) IsNot DGV1 And CheckBox33.Checked And ListBox3.Items.Count > 0 Then
                                '循环两次 第一次是出力没有Listbox3，第二次出力Listbox3的数据

                                For index As Integer = 0 To 1
                                    str = header & vbCrLf
                                    Dim count As Integer '计数



                                    'denpyounoS.Clear()
                                    If index = 0 Then
                                        count = 0
                                        For Each row As DataRowView In view
                                            Dim dataRow As String() = row("STR").ToString.Split(",")
                                            For c As Integer = 0 To dataRow.Count - 1
                                                dataRow(c) = dataRow(c).Replace("""", "")
                                            Next

                                            Dim denpyouno As String = ""
                                            If dgvM(i) Is DGV7 Then
                                                denpyouno = dataRow(dH7.IndexOf("お客様管理ナンバー"))
                                            ElseIf dgvM(i) Is DGV8 Then
                                                denpyouno = dataRow(dH8.IndexOf("顧客管理番号"))
                                            ElseIf dgvM(i) Is DGV9 Then
                                                denpyouno = dataRow(dH9.IndexOf("お客様側管理番号"))
                                            ElseIf dgvM(i) Is DGV13 Then
                                                denpyouno = dataRow(dH13.IndexOf("お客様側管理番号"))
                                            ElseIf dgvM(i) Is TMSDGV Then
                                                denpyouno = dataRow(tms_dh.IndexOf("お客様管理ナンバー"))
                                            End If

                                            If (ListBox3.Items.Contains(denpyouno) = False) Or (denpyouno = "") Then '不包含或者可能空的话
                                                str &= row("STR") & vbCrLf
                                                count = count + 1
                                            End If
                                        Next
                                        If count > 0 Then '有数据就出力
                                            Dim dstName As String = saveDir & "\★" & FileName & ".csv"
                                            saveName = dstName
                                            saveName = Replace(dstName, ".csv", "_" & fNameArray(k) & "_" & Format(Now, "yyyyMMddHHmmss") & ".csv")

                                            If File.Exists(saveName) Then
                                                Dim sl As String = Path.GetFileName(saveName) & "→"
                                                saveName = Replace(dstName, ".csv", "_" & Format(Now(), "HHmmss") & ".csv")
                                                sl &= Path.GetFileName(saveName)
                                                saveList.Add(sl)
                                            Else
                                                saveList.Add(Path.GetFileName(saveName))
                                            End If
                                            File.WriteAllText(saveName, str, ENC_SJ)
                                        End If
                                    Else
                                        count = 0
                                        For Each row As DataRowView In view
                                            Dim dataRow As String() = row("STR").ToString.Split(",")
                                            For c As Integer = 0 To dataRow.Count - 1
                                                dataRow(c) = dataRow(c).Replace("""", "")
                                            Next

                                            Dim denpyouno As String = ""
                                            If dgvM(i) Is DGV7 Then
                                                denpyouno = dataRow(dH7.IndexOf("お客様管理ナンバー"))
                                            ElseIf dgvM(i) Is DGV8 Then
                                                denpyouno = dataRow(dH8.IndexOf("顧客管理番号"))
                                            ElseIf dgvM(i) Is DGV9 Then
                                                denpyouno = dataRow(dH9.IndexOf("お客様側管理番号"))
                                            ElseIf dgvM(i) Is DGV13 Then
                                                denpyouno = dataRow(dH13.IndexOf("お客様側管理番号"))
                                            ElseIf dgvM(i) Is TMSDGV Then
                                                denpyouno = dataRow(tms_dh.IndexOf("お客様管理ナンバー"))
                                            End If

                                            If ListBox3.Items.Contains(denpyouno) = True Then '不包含或者可能空的话
                                                str &= row("STR") & vbCrLf
                                                count = count + 1

                                                denpyounoS.Add(denpyouno)
                                            End If


                                            'BIAOJI
                                            'For Each item As String In ListBox3.Items
                                            '    Debug.WriteLine(item)
                                            '    If item = denpyouno Then
                                            '        str &= row("STR") & vbCrLf
                                            '        count = count + 1
                                            '    End If

                                            'Next
                                        Next






                                        If count > 0 Then '有数据就出力
                                            Dim dstName As String = saveDir & "\★" & FileName & ".csv"
                                            saveName = dstName
                                            saveName = Replace(dstName, ".csv", "_" & fNameArray(k) & "(別紙)" & "_" & Format(Now, "yyyyMMddHHmmss") & ".csv")

                                            If File.Exists(saveName) Then
                                                Dim sl As String = Path.GetFileName(saveName) & "→"
                                                saveName = Replace(dstName, ".csv", "_" & Format(Now(), "HHmmss") & ".csv")
                                                sl &= Path.GetFileName(saveName)
                                                saveList.Add(sl)
                                            Else
                                                saveList.Add(Path.GetFileName(saveName))
                                            End If

                                            File.WriteAllText(saveName, str, ENC_SJ)
                                        End If
                                    End If
                                Next

                            End If

                            '-------------------- 20210312 改修后 end ----------------


                        End If



                        'ListBox3.Items.Clear()

                        'For Each item As String In denpyounoS
                        '    ListBox3.Items.Add(item)
                        'Next



                        'ゆう2 删除 zhushidiao
                        If (nameArray(k) = "YPK2T" Or nameArray(k) = "YPK2J") And False Then

                            LIST4VIEW("YPK2T", "start")

                            For c As Integer = 0 To yamato_title_sp.Count - 1
                                If CStr(yamato_header) = "" Then
                                    yamato_header = """" & yamato_title_sp(c) & """"
                                Else
                                    yamato_header &= "," & """" & yamato_title_sp(c) & """"
                                End If
                            Next
                            str = yamato_header & vbCrLf
                            Dim line_yamato As String = ""
                            For Each row As DataRowView In view
                                Dim dataRow As String() = row("STR").ToString.Split(",")

                                For c As Integer = 0 To dataRow.Count - 1
                                    dataRow(c) = dataRow(c).Replace("""", "")
                                Next
                                line_yamato = qtm & dataRow(dH9.IndexOf("お客様側管理番号")) & qtm & "," 'お客様管理番号

                                line_yamato &= qtm & okurizyo_type & qtm2
                                line_yamato &= qtm & cool_kubun & qtm2  'クール区分
                                line_yamato &= qtm & qtm2 '伝票番号
                                line_yamato &= qtm & syukayoteibi & qtm2 '出荷予定日

                                'お届け予定（指定）日
                                If dataRow(dH9.IndexOf("配送希望日")) = "" Then
                                    line_yamato &= qtm & qtm2
                                Else
                                    If IsNumeric(dataRow(dH9.IndexOf("配送希望日"))) Then
                                        line_yamato &= qtm & dataRow(dH9.IndexOf("配送希望日")).Substring(0, 4) & "/" & dataRow(dH9.IndexOf("配送希望日")).Substring(4, 2) & "/" & dataRow(dH9.IndexOf("配送希望日")).Substring(6, 2) & qtm2
                                    Else
                                        line_yamato &= qtm & qtm2
                                    End If
                                End If

                                '配達時間帯
                                '0812: 午前中
                                '1416: 14～16時
                                '1618: 16～18時
                                '1820: 18～20時
                                '1921: 19～21時
                                If dataRow(dH9.IndexOf("配送希望時間帯")) = "午前中" Then
                                    line_yamato &= qtm & "0812" & qtm2
                                ElseIf dataRow(dH9.IndexOf("配送希望時間帯")) = "12～14時" Or dataRow(dH9.IndexOf("配送希望時間帯")) = "14～16時" Then
                                    line_yamato &= qtm & "1416" & qtm2
                                ElseIf dataRow(dH9.IndexOf("配送希望時間帯")) = "16～18時" Then
                                    line_yamato &= qtm & "1618" & qtm2
                                ElseIf dataRow(dH9.IndexOf("配送希望時間帯")) = "18～20時" Then
                                    line_yamato &= qtm & "1820" & qtm2
                                ElseIf dataRow(dH9.IndexOf("配送希望時間帯")) = "19～21時" Or dataRow(dH9.IndexOf("配送希望時間帯")) = "20～21時" Then
                                    line_yamato &= qtm & "1921" & qtm2
                                Else
                                    line_yamato &= qtm & qtm2
                                End If

                                line_yamato &= qtm & qtm2 'お届け先コード
                                line_yamato &= qtm & dataRow(dH9.IndexOf("お届け先　電話番号")) & qtm2 'お届け先電話番号
                                line_yamato &= qtm & qtm2 'お届け先電話番号枝番
                                line_yamato &= qtm & dataRow(dH9.IndexOf("お届け先　郵便番号")) & qtm2 'お届け先郵便番号
                                line_yamato &= qtm & dataRow(dH9.IndexOf("お届け先　住所1")) & qtm2 'お届け先住所
                                line_yamato &= qtm & dataRow(dH9.IndexOf("お届け先　住所2")) & qtm2 'お届け先住所（アパートマンション名）
                                line_yamato &= qtm & qtm2 'お届け先会社・部門名１
                                line_yamato &= qtm & qtm2 'お届け先会社・部門名２

                                'お届け先名
                                If InStr(dataRow(dH9.IndexOf("お届け先　名称")), ")") Then
                                    Dim todokesakimei As String() = Split(dataRow(dH9.IndexOf("お届け先　名称")), ")")
                                    line_yamato &= qtm & todokesakimei(1) & qtm2
                                Else
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("お届け先　名称")) & qtm2
                                End If

                                line_yamato &= qtm & qtm2 'お届け先名略称カナ
                                line_yamato &= qtm & "様" & qtm2 '敬称

                                If dataRow(dH9.IndexOf("ご依頼主　名称１")) = "auPAYマーケット KuraNavi" Then
                                    line_yamato &= qtm & "au_KuraNavi" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "092-980-1144" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "812-0881" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県福岡市博多区井相田1-8-33-102" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & "au PAY マーケット" & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "Yahoo!FKstyle" Then
                                    line_yamato &= qtm & "Yahoo_Fk" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "092-586-6853" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & "2" & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "816-0911" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県大野城市大城4-13-15" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "楽天 通販の暁" Then
                                    line_yamato &= qtm & "Ra_Akatuki" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "092-986-1116" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "812-0881" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県福岡市博多区井相田1-8-33-203" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "Amazon通販のトココ" Then
                                    line_yamato &= qtm & "Ama_tokoko" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "000-0000-0000" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "816-0922" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県大野城市山田2-2-35" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & "," 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "楽天 雑貨の国のアリス" Then
                                    line_yamato &= qtm & "Ra_Alice" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "092-985-2056" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & "1" & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "816-0901" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県大野城市乙金東１－２－５２－１" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "Qoo10通販の雑貨倉庫" Then
                                    line_yamato &= qtm & "Qo_zakka" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "0120-699-991" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "812-0881" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県福岡市博多区井相田1-8-33-203" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "Qoo10福岡通販堂" Then
                                    line_yamato &= qtm & "Qo_tuhan_do" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "092-586-6853" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "812-0881" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県福岡市博多区井相田1-8-33-205" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "万方商事株式会社" Then
                                    line_yamato &= qtm & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "092-980-1866" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "812-0881" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県福岡市博多区井相田1-8-33-101" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "Yahoo!Lucky9" Then
                                    line_yamato &= qtm & "Yahoo_Lucky9" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "092-985-0275" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "812-0881" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県福岡市博多区井相田1-8-33-202" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "べんけい（soing1702）" Then
                                    line_yamato &= qtm & "Yaho_oku" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "000-0000-0000" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "812-0881" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県福岡市博多区井相田２－３－４３－１０２" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "サラダ" Or dataRow(dH9.IndexOf("ご依頼主　名称１")) = "Amazon Sarada" Then
                                    line_yamato &= qtm & "Ama_Sarada" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "000-0000-0000" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "812-0881" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県福岡市博多区井相田1-8-33-2F" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "Amazon Charyee" Then
                                    line_yamato &= qtm & "Ama_Charyee" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "000-0000-0000" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "812-0881" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県福岡市博多区井相田1-8-33-103" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "Yahoo!あかねAshop" Then
                                    line_yamato &= qtm & "Yahoo_Akane" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "092-985-0302" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "816-0922" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県大野城市山田2-2-35" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "楽天 あかねAshop" Then
                                    line_yamato &= qtm & "Ra_Akane" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "092-985-0295" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "816-0922" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県大野城市山田2-2-35" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "Amazon Hewflit" Then
                                    line_yamato &= qtm & "Ama_Hewflit" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "000-0000-0000" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "816-0901" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県大野城市乙金東１－２－５２－１" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "雑貨ショップKT 海東" Then
                                    line_yamato &= qtm & "Yahoo_KT" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "092-986-5538" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "811-0123" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県糟屋郡新宮町上府北3-6-3" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "Amazon 雑貨の国のアリス" Then
                                    line_yamato &= qtm & "a_Alice" & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "000-0000-0000" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "816-0901" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県大野城市乙金東１－２－５２－１" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                ElseIf dataRow(dH9.IndexOf("ご依頼主　名称１")) = "雑貨KT海東（ヤフオク）" Then
                                    line_yamato &= qtm & qtm2 'ご依頼主コード
                                    line_yamato &= qtm & "092-986-5538" & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & qtm2 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & "811-0123" & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & "福岡県糟屋郡新宮町上府北3-6-3" & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                Else
                                    line_yamato &= qtm & "," 'ご依頼主コード
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　電話番号")) & qtm2 'ご依頼主電話番号
                                    line_yamato &= qtm & "," 'ご依頼主電話番号枝番
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　郵便番号")) & qtm2 'ご依頼主郵便番号
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　住所１")) & qtm2 'ご依頼主住所
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　住所２")) & qtm2 'ご依頼主住所（アパートマンション名）
                                    line_yamato &= qtm & dataRow(dH9.IndexOf("ご依頼主　名称１")) & qtm2 'ご依頼主名
                                    line_yamato &= qtm & qtm2 'ご依頼主略称カナ
                                End If

                                line_yamato &= qtm & changeMarumozi(dataRow(dH9.IndexOf("フリー項目２"))) & qtm2  '品名コード１
                                line_yamato &= qtm & dataRow(dH9.IndexOf("フリー項目２")) & "(" & dataRow(dH9.IndexOf("フリー項目３")) & ")" & qtm2  '品名１

                                line_yamato &= qtm & qtm2  '品名コード２
                                line_yamato &= qtm & dataRow(dH9.IndexOf("品名")) & qtm2  '品名２
                                line_yamato &= qtm & qtm2  '荷扱い１
                                line_yamato &= qtm & qtm2  '荷扱い２

                                If dataRow(dH9.IndexOf("マスタ配送")) = "ヤマト(陸便)" Then
                                    line_yamato &= qtm & qtm2  '記事
                                Else
                                    line_yamato &= qtm & "船便" & qtm2  '記事
                                End If

                                line_yamato &= qtm & qtm2  'コレクト代金引換額（税込）
                                line_yamato &= qtm & qtm2  'コレクト内消費税額等
                                line_yamato &= qtm & "0" & qtm2 '営業所止置き
                                line_yamato &= qtm & qtm2  '営業所コード
                                line_yamato &= qtm & "1" & qtm2 '発行枚数
                                line_yamato &= qtm & qtm2  '個数口枠の印字
                                line_yamato &= qtm & seikyukyakucode & qtm2  'ご請求先顧客コード
                                line_yamato &= qtm & qtm2  'ご請求先分類コード
                                line_yamato &= qtm & "01" & qtm2 '運賃管理番号
                                line_yamato &= qtm & "0" & qtm2 'クロネコwebコレクトデータ登録
                                line_yamato &= qtm & qtm2  'クロネコwebコレクト加盟店番号
                                line_yamato &= qtm & qtm2  'クロネコwebコレクト申込受付番号１
                                line_yamato &= qtm & qtm2  'クロネコwebコレクト申込受付番号２
                                line_yamato &= qtm & qtm2  'クロネコwebコレクト申込受付番号３
                                line_yamato &= qtm & "0" & qtm2 'お届け予定ｅメール利用区分
                                line_yamato &= qtm & qtm2  'お届け予定ｅメールe-mailアドレス
                                line_yamato &= qtm & qtm2  '入力機種
                                line_yamato &= qtm & qtm2  'お届け予定eメールメッセージ
                                line_yamato &= qtm & "0" & qtm2 'お届け完了eメール利用区分
                                line_yamato &= qtm & qtm2  'お届け完了ｅメールe-mailアドレス
                                line_yamato &= qtm & qtm2  'お届け完了ｅメールメッセージ
                                line_yamato &= qtm & "0" & qtm2 'クロネコ収納代行利用区分
                                line_yamato &= qtm & qtm2  '収納代行決済ＱＲコード印刷
                                line_yamato &= qtm & qtm2  '収納代行請求金額(税込)
                                line_yamato &= qtm & qtm2  '収納代行内消費税額等
                                line_yamato &= qtm & qtm2  '収納代行請求先郵便番号
                                line_yamato &= qtm & qtm2  '収納代行請求先住所
                                line_yamato &= qtm & qtm2  '収納代行請求先住所（アパートマンション名）
                                line_yamato &= qtm & qtm2  '収納代行請求先会社・部門名１
                                line_yamato &= qtm & qtm2  '収納代行請求先会社・部門名２
                                line_yamato &= qtm & qtm2  '収納代行請求先名(漢字)
                                line_yamato &= qtm & qtm2  '収納代行請求先名(カナ)
                                line_yamato &= qtm & qtm2  '収納代行問合せ先名(漢字)
                                line_yamato &= qtm & qtm2  '収納代行問合せ先郵便番号
                                line_yamato &= qtm & qtm2  '収納代行問合せ先住所
                                line_yamato &= qtm & qtm2  '収納代行問合せ先住所（アパートマンション名）
                                line_yamato &= qtm & qtm2  '収納代行問合せ先電話番号
                                line_yamato &= qtm & qtm2  '収納代行管理番号
                                line_yamato &= qtm & qtm2  '収納代行品名
                                line_yamato &= qtm & qtm2  '収納代行備考
                                line_yamato &= qtm & qtm2  '複数口くくりキー
                                line_yamato &= qtm & qtm2  '検索キータイトル１
                                line_yamato &= qtm & qtm2  '検索キー１
                                line_yamato &= qtm & qtm2  '検索キータイトル２
                                line_yamato &= qtm & qtm2  '検索キー２
                                line_yamato &= qtm & qtm2  '検索キータイトル３
                                line_yamato &= qtm & qtm2  '検索キー３
                                line_yamato &= qtm & qtm2  '検索キータイトル４
                                line_yamato &= qtm & qtm2  '検索キー４
                                line_yamato &= qtm & qtm2  '検索キータイトル５
                                line_yamato &= qtm & qtm2  '検索キー５
                                line_yamato &= qtm & qtm2  '予備
                                line_yamato &= qtm & qtm2  '予備
                                line_yamato &= qtm & "0" & qtm2 '投函予定メール利用区分
                                line_yamato &= qtm & qtm2  '投函予定メールe-mailアドレス
                                line_yamato &= qtm & qtm2  '投函予定メールメッセージ
                                line_yamato &= qtm & "0" & qtm2 '投函完了メール（お届け先宛）利用区分
                                line_yamato &= qtm & qtm2  '投函完了メール（お届け先宛）e-mailアドレス
                                line_yamato &= qtm & qtm2  '投函完了メール（お届け先宛）メールメッセージ
                                line_yamato &= qtm & "0" & qtm2 '投函完了メール（ご依頼主宛）利用区分
                                line_yamato &= qtm & qtm2  '投函完了メール（ご依頼主宛）e-mailアドレス
                                line_yamato &= qtm & qtm2  '投函完了メール（ご依頼主宛）メールメッセージ
                                line_yamato &= qtm & qtm2  '連携管理番号
                                line_yamato &= qtm & qtm  '通知メールアドレス
                                line_yamato &= vbCrLf

                                str &= line_yamato
                                Console.WriteLine(">> " + str)
                            Next

                            If nameArray(k) = "YPK2T" Then
                                saveName = saveDir & "\★ゆう2" & loginName & "_" & HS1.Text & "_" & Format(Now, "yyyyMMddHHmmss") & ".csv"
                            ElseIf nameArray(k) = "YPK2J" Then
                                saveName = saveDir & "\★ゆう2" & loginName & "_ " & HS2.Text & "_" & Format(Now, "yyyyMMddHHmmss") & ".csv"
                            Else
                                saveName = saveDir & "\★ゆう2" & loginName & "_ " & "_ " & "_" & Format(Now, "yyyyMMddHHmmss") & ".csv"
                            End If
                            saveList.Add(Path.GetFileName(saveName))
                            File.WriteAllText(saveName, str, ENC_SJ)
                        End If
                        'File.WriteAllText(saveName, str, ENC_SJ) '-------------------- 20210312 改修前 这里不要了 ----------------

                        '--------------
                        '実績
                        'If InStr(FileName, "(印刷しない)元データ") > 0 Then

                        If InStr(FileName, "(印刷しない)元データ") > 0 And Csv_denpyo3_F_count.Button2.BackColor <> Color.Yellow Then
                                'If InStr(FileName, "(印刷しない)元データ") > 0 Then
                                LIST4VIEW("実績処理開始", "START")
                                Dim serverDir As String = Form1.サーバーToolStripMenuItem.Text & "\denpyoLog\"
                                Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)

                                Dim todayTxtPath As String = ""
                                Dim todayCsvPath As String = ""
                            '这里
                            If Regex.IsMatch(appPath, "debug", RegexOptions.IgnoreCase) Then
                                todayTxtPath = desktopPath & "\" & Format(Now, "yyyyMMdd") & ".txt"
                                todayCsvPath = desktopPath & "\" & Format(Now, "yyyyMMdd") & ".csv"
                            Else
                                todayTxtPath = serverDir & Format(Now, "yyyyMMdd") & ".txt"
                                todayCsvPath = serverDir & Format(Now, "yyyyMMdd") & ".csv"
                            End If

                            Dim checkArray As New ArrayList
                            Dim mArray As New ArrayList
                            '18.19.20 >> TMS
                            Dim binsu As Integer() = New Integer() {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}
                            If File.Exists(todayCsvPath) Then
                                    mArray = TM_CSV_READ(todayCsvPath)(0)
                                    Dim mH As String() = Split(mArray(0), "|=|")
                                    For p As Integer = 1 To mArray.Count - 1
                                        Dim mLine As String() = Split(mArray(p), "|=|")

                                        checkArray.Add(mLine(Array.IndexOf(mH, "伝票番号")) & "," & mLine(Array.IndexOf(mH, "受注番号")))
                                    Next
                                End If
                                Dim jStr As String = ""
                                Dim cArray As ArrayList = TM_CSV_READ(saveName)(0)
                                Dim cH As String() = Split(cArray(0), "|=|")
                                Dim nTime As String = Format(Now(), "HHmmss")
                                For p As Integer = 1 To cArray.Count - 1
                                    Dim cLine As String() = Split(cArray(p), "|=|")

                                    Dim a = cLine(Array.IndexOf(cH, "伝票番号"))

                                    Dim B = cLine(Array.IndexOf(cH, "受注番号"))
                                    Dim K3 = cLine(Array.IndexOf(cH, "マスタ便数"))
                                    Dim check As String = cLine(Array.IndexOf(cH, "伝票番号")) & "," & cLine(Array.IndexOf(cH, "受注番号"))
                                    '重複しない
                                    If checkArray.Contains(check) Then
                                        Dim cNum As Integer = checkArray.IndexOf(check)
                                        checkArray.RemoveAt(cNum)
                                        mArray.RemoveAt(cNum + 1)
                                    End If
                                    jStr &= """" & cLine(Array.IndexOf(cH, "伝票ソフト")) & ""","
                                    jStr &= """" & cLine(Array.IndexOf(cH, "伝票番号")) & ""","
                                    jStr &= """" & cLine(Array.IndexOf(cH, "店舗")) & ""","
                                    jStr &= """" & cLine(Array.IndexOf(cH, "受注番号")) & ""","
                                    jStr &= """" & cLine(Array.IndexOf(cH, "購入者名")) & ""","
                                    jStr &= """" & cLine(Array.IndexOf(cH, "発送先名")) & ""","
                                    jStr &= """" & cLine(Array.IndexOf(cH, "便種")) & ""","
                                    jStr &= """" & cLine(Array.IndexOf(cH, "マスタ便数")) & ""","
                                    jStr &= """" & cLine(Array.IndexOf(cH, "商品マスタ")) & ""","
                                    jStr &= """" & nTime & ""","
                                    jStr &= """" & Form1.ログイン名ToolStripMenuItem.Text & ""","
                                    jStr &= """" & cLine(Array.IndexOf(cH, "発送倉庫")) & ""","
                                    jStr &= """" & saveName & """" & vbCrLf
                                    Select Case cLine(Array.IndexOf(cH, "発送倉庫")) & cLine(Array.IndexOf(cH, "便種")) & cLine(Array.IndexOf(cH, "伝票ソフト"))
                                        Case "太宰府陸便BIZlogi", "太宰府陸便e飛伝2"
                                            binsu(0) += CInt(cLine(Array.IndexOf(cH, "マスタ便数")))
                                        Case "太宰府航空便BIZlogi", "太宰府航空便e飛伝2"
                                            binsu(1) += CInt(cLine(Array.IndexOf(cH, "マスタ便数")))
                                        Case "太宰府陸便メール便", "太宰府航空便メール便"
                                            binsu(2) += CInt(cLine(Array.IndexOf(cH, "マスタ便数")))
                                        Case "太宰府陸便定形外", "太宰府航空便定形外"
                                            binsu(3) += CInt(cLine(Array.IndexOf(cH, "マスタ便数")))
                                        Case "井相田陸便BIZlogi", "井相田陸便e飛伝2"
                                            binsu(4) += CInt(cLine(Array.IndexOf(cH, "マスタ便数")))
                                        Case "井相田航空便BIZlogi", "井相田航空便e飛伝2"
                                            binsu(5) += CInt(cLine(Array.IndexOf(cH, "マスタ便数")))
                                        Case "井相田陸便メール便", "井相田航空便メール便"
                                            binsu(6) += CInt(cLine(Array.IndexOf(cH, "マスタ便数")))
                                        Case "井相田陸便定形外", "井相田航空便定形外"
                                            binsu(7) += CInt(cLine(Array.IndexOf(cH, "マスタ便数")))
                                        Case "名古屋陸便BIZlogi", "名古屋陸便e飛伝2"
                                            binsu(8) += CInt(cLine(Array.IndexOf(cH, "マスタ便数")))
                                        Case "名古屋航空便BIZlogi", "名古屋航空便e飛伝2"
                                            binsu(9) += CInt(cLine(Array.IndexOf(cH, "マスタ便数")))
                                        Case "名古屋陸便メール便", "名古屋航空便メール便"
                                            binsu(10) += CInt(cLine(Array.IndexOf(cH, "マスタ便数")))
                                        Case "名古屋陸便定形外", "名古屋航空便定形外"
                                            binsu(11) += CInt(cLine(Array.IndexOf(cH, "マスタ便数")))
                                        Case "太宰府陸便ヤマト"
                                            binsu(12) += CInt(cLine(Array.IndexOf(cH, "マスタ便数")))
                                        Case "井相田陸便ヤマト"
                                            binsu(13) += CInt(cLine(Array.IndexOf(cH, "マスタ便数")))
                                        Case "太宰府船便ヤマト"
                                            binsu(14) += CInt(cLine(Array.IndexOf(cH, "マスタ便数")))
                                        Case "井相田船便ヤマト"
                                            binsu(15) += CInt(cLine(Array.IndexOf(cH, "マスタ便数")))
                                        Case "太宰府陸便ゆう2"
                                            binsu(16) += CInt(cLine(Array.IndexOf(cH, "マスタ便数")))
                                        Case "井相田陸便ゆう2"
                                        binsu(17) += CInt(cLine(Array.IndexOf(cH, "マスタ便数")))
                                    Case "太宰府陸便TMS"
                                        binsu(18) += CInt(cLine(Array.IndexOf(cH, "マスタ便数")))
                                    Case "井相田陸便TMS"
                                        binsu(19) += CInt(cLine(Array.IndexOf(cH, "マスタ便数")))
                                    Case "名古屋陸便TMS"
                                        binsu(20) += CInt(cLine(Array.IndexOf(cH, "マスタ便数")))
                                        'Case Else
                                        '    binsu(8) += CInt(cLine(Array.IndexOf(cH, "マスタ便数")))
                                End Select

                                    If p Mod 500 = 0 Then
                                        LIST4VIEW("確認(1)..." & p, "system")
                                    End If
                                Next


                                Dim mStrArray As New ArrayList
                                If mArray.Count > 1 Then
                                    For m As Integer = 1 To mArray.Count - 1
                                        Dim sStr As String = ""
                                        Dim mLine As String() = Split(mArray(m), "|=|")
                                        Dim mH As String() = Split(mArray(0), "|=|")
                                        For Each mL As String In mLine
                                            sStr &= """" & mL & ""","
                                        Next
                                        mStrArray.Add(sStr)
                                        'sStr &= vbCrLf
                                        Select Case mLine(Array.IndexOf(mH, "倉庫")) & mLine(Array.IndexOf(mH, "便種")) & mLine(Array.IndexOf(mH, "伝票ソフト"))

                                        'Case "太宰府陸便BIZlogi", "太宰府陸便e飛伝2"
                                        '    binsu(0) += CInt(mLine(Array.IndexOf(mH, "マスタ便数")))
                                        'Case "太宰府航空便BIZlogi", "太宰府航空便e飛伝2"
                                        '    binsu(1) += CInt(mLine(Array.IndexOf(mH, "マスタ便数")))
                                        'Case "太宰府陸便メール便", "太宰府航空便メール便"
                                        '    binsu(2) += CInt(mLine(Array.IndexOf(mH, "マスタ便数")))
                                        'Case "太宰府陸便定形外", "太宰府航空便定形外"
                                        '    binsu(3) += CInt(mLine(Array.IndexOf(mH, "マスタ便数")))
                                        'Case "井相田陸便BIZlogi", "井相田陸便e飛伝2"
                                        '    binsu(4) += CInt(mLine(Array.IndexOf(mH, "マスタ便数")))
                                        'Case "井相田航空便BIZlogi", "井相田航空便e飛伝2"
                                        '    binsu(5) += CInt(mLine(Array.IndexOf(mH, "マスタ便数")))
                                        'Case "井相田陸便メール便", "井相田航空便メール便"
                                        '    binsu(6) += CInt(mLine(Array.IndexOf(mH, "マスタ便数")))
                                        'Case "井相田陸便定形外", "井相田航空便定形外"
                                        '    binsu(7) += CInt(mLine(Array.IndexOf(mH, "マスタ便数")))
                                        'Case Else
                                        '    binsu(8) += CInt(mLine(Array.IndexOf(mH, "マスタ便数")))
                                            Case "太宰府陸便BIZlogi", "太宰府陸便e飛伝2"

                                                binsu(0) += CInt(mLine(Array.IndexOf(mH, "マスタ便数")))
                                            Case "太宰府航空便BIZlogi", "太宰府航空便e飛伝2"
                                                binsu(1) += CInt(mLine(Array.IndexOf(mH, "マスタ便数")))
                                            Case "太宰府陸便メール便", "太宰府航空便メール便"
                                                binsu(2) += CInt(mLine(Array.IndexOf(mH, "マスタ便数")))
                                            Case "太宰府陸便定形外", "太宰府航空便定形外"
                                                binsu(3) += CInt(mLine(Array.IndexOf(mH, "マスタ便数")))
                                            Case "井相田陸便BIZlogi", "井相田陸便e飛伝2"
                                                binsu(4) += CInt(mLine(Array.IndexOf(mH, "マスタ便数")))
                                            Case "井相田航空便BIZlogi", "井相田航空便e飛伝2"
                                                binsu(5) += CInt(mLine(Array.IndexOf(mH, "マスタ便数")))
                                            Case "井相田陸便メール便", "井相田航空便メール便"
                                                binsu(6) += CInt(mLine(Array.IndexOf(mH, "マスタ便数")))
                                            Case "井相田陸便定形外", "井相田航空便定形外"
                                                binsu(7) += CInt(mLine(Array.IndexOf(mH, "マスタ便数")))
                                            Case "名古屋陸便BIZlogi", "名古屋陸便e飛伝2"
                                                binsu(8) += CInt(mLine(Array.IndexOf(mH, "マスタ便数")))
                                            Case "名古屋航空便BIZlogi", "名古屋航空便e飛伝2"
                                                binsu(9) += CInt(mLine(Array.IndexOf(mH, "マスタ便数")))
                                            Case "名古屋陸便メール便", "名古屋航空便メール便"
                                                binsu(10) += CInt(mLine(Array.IndexOf(mH, "マスタ便数")))
                                            Case "名古屋陸便定形外", "名古屋航空便定形外"
                                                binsu(11) += CInt(mLine(Array.IndexOf(mH, "マスタ便数")))
                                            Case "太宰府陸便ヤマト"
                                                binsu(12) += CInt(mLine(Array.IndexOf(mH, "マスタ便数")))
                                            Case "井相田陸便ヤマト"
                                                binsu(13) += CInt(mLine(Array.IndexOf(mH, "マスタ便数")))
                                            Case "太宰府船便ヤマト"
                                                binsu(14) += CInt(mLine(Array.IndexOf(mH, "マスタ便数")))
                                            Case "井相田船便ヤマト"
                                                binsu(15) += CInt(mLine(Array.IndexOf(mH, "マスタ便数")))

                                            Case "太宰府陸便ゆう2"
                                            binsu(16) += CInt(mLine(Array.IndexOf(mH, "マスタ便数")))
                                        Case "井相田陸便ゆう2"
                                            binsu(17) += CInt(mLine(Array.IndexOf(mH, "マスタ便数")))
                                            '这里
                                        Case "太宰府陸便TMS"
                                            binsu(18) += CInt(mLine(Array.IndexOf(mH, "マスタ便数")))
                                        Case "井相田陸便TMS"
                                            binsu(19) += CInt(mLine(Array.IndexOf(mH, "マスタ便数")))
                                        Case "名古屋陸便TMS"
                                            binsu(20) += CInt(mLine(Array.IndexOf(mH, "マスタ便数")))









                                    End Select

                                        If m Mod 500 = 0 Then
                                            LIST4VIEW("確認(2)..." & m, "system")
                                        End If
                                    Next
                                End If
                                'jStr = "伝票ソフト,伝票番号,店舗,受注番号,購入者名,発送先名,便種,マスタ便数,商品マスタ,時間,担当,倉庫,ファイル名" & vbCrLf & jStr
                                If File.Exists(todayCsvPath) Then
                                    mStrArray.Insert(0, "伝票ソフト,伝票番号,店舗,受注番号,購入者名,発送先名,便種,マスタ便数,商品マスタ,時間,担当,倉庫,ファイル名")
                                    Dim jArray As String() = Split(jStr, vbCrLf)
                                    For j As Integer = 0 To jArray.Length - 1
                                        If jArray(j) <> "" Then
                                            mStrArray.Insert(j + 1, jArray(j))
                                        End If
                                    Next
                                    Dim writeArray As String() = CType(mStrArray.ToArray(Type.GetType("System.String")), String())


                                    File.WriteAllLines(todayCsvPath, writeArray, ENC_SJ)

                                Else
                                    jStr = "伝票ソフト,伝票番号,店舗,受注番号,購入者名,発送先名,便種,マスタ便数,商品マスタ,時間,担当,倉庫,ファイル名" & vbCrLf & jStr


                                    File.WriteAllText(todayCsvPath, jStr, ENC_SJ)
                                End If
                                Dim countStr As String = ""
                            For m As Integer = 0 To binsu.Length - 1
                                countStr &= binsu(m) & ","
                            Next





                            LIST4VIEW("実績処理......", "system")
                            '测试结束后修改
                            File.WriteAllText(todayTxtPath, countStr, ENC_SJ)
                            If File.Exists(lockPath) Then
                                File.Delete(lockPath)

                            End If

                            LIST4VIEW("実績処理終了", "END")
                            End If
                            '--------------
                        End If
                        LIST4VIEW("Save " & nameArray(k), "system")
                Next

                LIST4VIEW("Save " & nameArray(i), "system")

            End If
        Next

        'ヤマト
        'Dim yamato_title_header As String = ""

        'For c As Integer = 0 To yamato_title_sp.Count - 1
        '    If CStr(yamato_title_header) = "" Then
        '        yamato_title_header = """" & yamato_title_sp(c) & """"
        '    Else
        '        yamato_title_header &= "," & """" & yamato_title_sp(c) & """"
        '    End If
        'Next

        'If arr_yamato_dazaifu.Count > 0 Then
        '    Dim str_yamato As String = yamato_title_header & vbCrLf

        '    For ya As Integer = 0 To arr_yamato_dazaifu.Count - 1
        '        str_yamato &= arr_yamato_dazaifu(ya) & vbCrLf
        '    Next
        '    Dim saveName_yamato As String = "★ヤマト" & loginName & "_太宰府_" & Format(Now, "yyyyMMddHHmmss") & ".csv"
        '    File.WriteAllText(saveDir & "\" & saveName_yamato, str_yamato, ENC_SJ)
        '    saveList.Add(saveName_yamato)
        'End If

        'If arr_yamato_isouda.Count > 0 Then
        '    Dim str_yamato As String = yamato_title_header & vbCrLf

        '    For ya As Integer = 0 To arr_yamato_isouda.Count - 1
        '        str_yamato &= arr_yamato_isouda(ya) & vbCrLf
        '    Next
        '    Dim saveName_yamato As String = "★ヤマト" & loginName & "_井相田_" & Format(Now, "yyyyMMddHHmmss") & ".csv"
        '    File.WriteAllText(saveDir & "\" & saveName_yamato, str_yamato, ENC_SJ)
        '    saveList.Add(saveName_yamato)
        'End If

        'DGV3保存
        If DGV3.RowCount > 0 Then
            FileName = saveDir & "\" & "(印刷しない)明細データ.csv"
            DGV_TO_CSV_SAVE(FileName, DGV3,, True)
            saveList.Add("(印刷しない)明細データ.csv")
        End If

        'DGV14保存
        If DGV14.RowCount > 0 Then
            FileName = saveDir & "\" & "(印刷しない)履歴データ.csv"
            DGV_TO_CSV_SAVE(FileName, DGV14,, True)
            saveList.Add("(印刷しない)履歴データ.csv")
        End If

        '別紙PDF保存
        If ListBox3.Items.Count > 0 Then
            PDFsave(saveDir & "\", False)
            saveList.Add("●別紙PDF.pdf")

            'とんよか卸だけ出力
            Excelsave_hacyu(saveList, saveDir & "\", False)
        End If

        If CheckBox5.Checked Then
            EXCELsave(saveDir & "\")
            saveList.Add("▲" & TextBox40.Text)
        End If

        Label9.Visible = False
        saveList.Sort()
        MsgBox(TM_ARRAYLIST_TO_STR(saveList, vbCrLf) & vbCrLf & "以上のファイルを保存しました", MsgBoxStyle.OkOnly Or MsgBoxStyle.SystemModal)
        LIST4VIEW("Save Logic", "e")
    End Sub
    Private Function changeMarumozi(str As String)
        Dim marumozi As String() = New String() {"②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩", "⑪", "⑫", "⑬", "⑭", "⑮", "⑯", "⑰", "⑱", "⑲"}
        Dim marumozi_ As String() = New String() {"2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19"}

        If str <> "" Then
            For i As Integer = 0 To marumozi.Length - 1
                If InStr(str, marumozi(i)) > 0 Then
                    str = str.Replace(marumozi(i), marumozi_(i))
                End If
            Next
        End If

        Return str
    End Function

    Private Sub 出荷指示書のみ出力ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 出荷指示書のみ出力ToolStripMenuItem.Click
        Dim saveDir As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        EXCELsave(saveDir & "\")
    End Sub

    Private Sub EXCELsave(dir As String)
        'Dim sheet As Integer = 0

        '店舗名を取得する
        'Dim dH1 As ArrayList = TM_HEADER_GET(dgv1)
        'Dim tenpoArray As New ArrayList
        'For r As Integer = 0 To dgv1.RowCount - 1
        '    Dim tenpo As String = dgv1.Item(dH1.IndexOf("店舗"), r).Value
        '    If Not tenpoArray.Contains(tenpo) Then
        '        tenpoArray.Add(tenpo)
        '    End If
        'Next
        'Dim tenpoStr As String = ""
        'For Each tp As String In tenpoArray
        '    tenpoStr &= tp & "、"
        'Next
        'tenpoStr = tenpoStr.TrimEnd("、")

        '店舗名を取得する（tooltipに保存した店舗名を利用）
        Dim tenpoStr As String = ""
        Dim lkArray As LinkLabel() = {LinkLabel5, LinkLabel6, LinkLabel1, LinkLabel21}
        For i As Integer = 0 To lkArray.Length - 1
            Dim tip As String() = Split(ToolTip1.GetToolTip(lkArray(i)), "、")
            For k As Integer = 0 To tip.Length - 1
                If InStr(tenpoStr, tip(k)) = 0 Then
                    If tenpoStr = "" Then
                        tenpoStr = tip(k)
                    Else
                        tenpoStr &= "/" & tip(k)
                    End If
                End If
            Next
        Next

        Dim dH1 As ArrayList = TM_HEADER_GET(DGV1)

        '別紙枚数の計算
        Dim bessiArray() As Integer = New Integer() {0, 0, 0, 0, 0, 0}
        For i As Integer = 0 To ListBox3.Items.Count - 1
            For r As Integer = 0 To DGV1.Rows.Count - 1
                If ListBox3.Items(i) = DGV1.Item(TM_ArIndexof(dH1, koumoku("denpyoNo")(0)), r).Value Then
                    Select Case DGV1.Item(TM_ArIndexof(dH1, koumoku("syori2")(0)), r).Value
                        Case HS1.Text
                            If DGV1.Item(TM_ArIndexof(dH1, koumoku("binsyu")(0)), r).Value = "陸便" Then
                                bessiArray(0) += 1
                            Else
                                bessiArray(1) += 1
                            End If
                        Case HS2.Text
                            If DGV1.Item(TM_ArIndexof(dH1, koumoku("binsyu")(0)), r).Value = "陸便" Then
                                bessiArray(2) += 1
                            Else
                                bessiArray(3) += 1
                            End If
                        Case HS3.Text
                            If DGV1.Item(TM_ArIndexof(dH1, koumoku("binsyu")(0)), r).Value = "陸便" Then
                                bessiArray(4) += 1
                            Else
                                bessiArray(5) += 1
                            End If
                    End Select
                    Exit For
                End If
            Next
        Next


        'ログイン名前取得
        Dim loginName As String = "(" & Replace(Form1.ログイン名ToolStripMenuItem.Text, "さん", "") & ")"

        'テンプレートに記入
        If LinkLabel1.Text > 0 Or LinkLabel21.Text > 0 Or LinkLabel5.Text > 0 Or LinkLabel6.Text > 0 Then
            Dim wb As IWorkbook = WorkbookFactory.Create(Path.GetDirectoryName(Form1.appPath) & "\config\" & TextBox40.Text)
            Dim ws As ISheet = wb.GetSheetAt(wb.GetSheetIndex("Sheet1"))
            ws.GetRow(3).GetCell(1).SetCellValue(tenpoStr)
            ws.GetRow(7).GetCell(5).SetCellValue(CInt(Math.Ceiling(CInt(LinkLabel5.Text) / 10)))    'ゆうパケット
            ws.GetRow(7).GetCell(7).SetCellValue(LinkLabel5.Text)
            ws.GetRow(9).GetCell(5).SetCellValue(CInt(Math.Ceiling(CInt(LinkLabel6.Text) / 10)))    '定形外
            ws.GetRow(9).GetCell(7).SetCellValue(CInt(LinkLabel6.Text))
            ws.GetRow(11).GetCell(5).SetCellValue(CInt(LinkLabel1.Text) + CInt(LinkLabel21.Text))   'e飛伝2+BIZlogi
            ws.GetRow(13).GetCell(5).SetCellValue(bessiArray(0))                                    '別紙
            ws.GetRow(17).GetCell(0).SetCellValue(AzukiControl1.Text)
            ws.GetRow(15).GetCell(7).SetCellValue(Replace(Form1.ログイン名ToolStripMenuItem.Text, "さん", ""))
            Using wfs = File.Create(dir & "▲出力確認票" & Format(Now, "_yyyyMMddHHmmss") & loginName & ".xlsx")
                wb.Write(wfs)
            End Using
            ws = Nothing
            wb = Nothing
        End If

        '航空便
        If LinkLabel2.Text > 0 Or LinkLabel22.Text > 0 Then
            'tenpoStr = ToolTip1.GetToolTip(LinkLabel2)
            tenpoStr = ""
            Dim lkArray2 As LinkLabel() = {LinkLabel2, LinkLabel22}
            For i As Integer = 0 To lkArray2.Length - 1
                Dim tip As String() = Split(ToolTip1.GetToolTip(lkArray2(i)), "、")
                For k As Integer = 0 To tip.Length - 1
                    If InStr(tenpoStr, tip(k)) = 0 Then
                        If tenpoStr = "" Then
                            tenpoStr = tip(k)
                        Else
                            tenpoStr &= "/" & tip(k)
                        End If
                    End If
                Next
            Next
            Dim wb As IWorkbook = WorkbookFactory.Create(Path.GetDirectoryName(Form1.appPath) & "\config\" & Path.GetFileNameWithoutExtension(TextBox40.Text) & "(航空便).xlsx")
            Dim ws As ISheet = wb.GetSheetAt(wb.GetSheetIndex("Sheet1"))
            ws.GetRow(3).GetCell(1).SetCellValue(tenpoStr)
            ws.GetRow(8).GetCell(5).SetCellValue(CInt(LinkLabel2.Text) + CInt(LinkLabel22.Text))    '出荷個数
            ws.GetRow(10).GetCell(5).SetCellValue(bessiArray(1))                                    '別紙
            ws.GetRow(14).GetCell(0).SetCellValue(AzukiControl3.Text)
            ws.GetRow(12).GetCell(7).SetCellValue(Replace(Form1.ログイン名ToolStripMenuItem.Text, "さん", ""))
            Using wfs = File.Create(dir & "▲出力確認票(航空便)" & Format(Now, "_yyyyMMddHHmmss") & loginName & ".xlsx")
                wb.Write(wfs)
            End Using
            ws = Nothing
            wb = Nothing
            File.WriteAllText(dir & "(印刷しない)航空便明細(e飛伝2).txt", TextBox4.Text)
        End If
        If LinkLabel4.Text > 0 Then
            File.WriteAllText(dir & "(印刷しない)航空便明細(BIZlogi).txt", TextBox5.Text)
        End If
    End Sub

    Dim LockFlag As Boolean = False
    Private Sub 処理済のロック解除ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 処理済のロック解除ToolStripMenuItem.Click
        Dim DR As DialogResult = MsgBox("ロック解除後の処理は行なえません。解除しますか？" & vbCrLf & "※リセットするまでロック解除されます。", MsgBoxStyle.OkCancel Or MsgBoxStyle.SystemModal)
        If DR = DialogResult.Cancel Then
            Exit Sub
        End If

        KryptonButton3.Enabled = False

        DGV7.ReadOnly = False
        DGV8.ReadOnly = False
        DGV9.ReadOnly = False
        DGV13.ReadOnly = False
    End Sub

    'ピッキングリストcsv出力
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        If DGV16.RowCount = 0 Then
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
            .CheckPathExists = True,
            .FileName = "ピッキングリスト_" & Format(Now, "yyyyMMddHHmmss") & ".csv"
        }

        'ダイアログを表示する
        If sfd.ShowDialog(Me) = DialogResult.OK Then
            Dim kindName As String = ""
            Dim dH16 As ArrayList = TM_HEADER_GET(DGV16)
            Dim LB5 As New ArrayList
            For i As Integer = 0 To ListBox5.SelectedItems.Count - 1
                LB5.Add(ListBox5.SelectedItems(i))
                kindName &= ListBox5.SelectedItems(i)
            Next
            Dim LB6 As New ArrayList
            For i As Integer = 0 To ListBox6.SelectedItems.Count - 1
                LB6.Add(ListBox6.SelectedItems(i))
                kindName &= ListBox6.SelectedItems(i)
            Next
            For r As Integer = DGV16.RowCount - 1 To 0 Step -1
                If LB5.Contains(DGV16.Item(dH16.IndexOf("発送倉庫"), r).Value) And LB6.Contains(DGV16.Item(dH16.IndexOf("発送方法"), r).Value) Then
                    DGV16.Rows(r).Visible = True
                Else
                    DGV16.Rows(r).Visible = False
                End If
            Next

            Dim fName As String = Replace(sfd.FileName, "ピッキングリスト_", "ピッキングリスト(" & kindName & ")_")
            SaveCsv(fName, 1, DGV16)

            For r As Integer = DGV16.RowCount - 1 To 0 Step -1
                DGV16.Rows(r).Visible = True
            Next

            MsgBox(sfd.FileName & vbCrLf & "保存しました", MsgBoxStyle.SystemModal)
        End If
    End Sub

    Private localTimePath As String = appPathDir & "\db\sagawa.dat"
    Private localPath As String = appPathDir & "\db\sagawa.accdb"
    Private serverTimePath As String = Form1.サーバーToolStripMenuItem.Text & "\update\db\sagawa.dat"
    Private serverPath As String = Form1.サーバーToolStripMenuItem.Text & "\update\db\sagawa.accdb"

    Private kenallLocalTimePath As String = appPathDir & "\db\ken_all.dat"
    Private kenallLocalPath As String = appPathDir & "\db\ken_all.accdb"
    Private kenallServerTimePath As String = Form1.サーバーToolStripMenuItem.Text & "\update\db\ken_all.dat"
    Private kenallServerPath As String = Form1.サーバーToolStripMenuItem.Text & "\update\db\ken_all.accdb"

    Private Sub 佐川DBアップデートToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 佐川DBアップデートToolStripMenuItem.Click
        Dim sagawaDBflag As Boolean = False

        If Not File.Exists(localTimePath) Then
            File.WriteAllText(localTimePath, "")
        End If

        Dim localUpdateTime As String = File.ReadAllText(localTimePath, encSJ)
        Dim serverUpdateTime As String = File.ReadAllText(serverTimePath, encSJ)
        If File.Exists(localTimePath) Then
            If localUpdateTime = serverUpdateTime Then
                MsgBox("佐川DBデータベースは最新版です")
                sagawaDBflag = True
            End If
        End If
        If Not sagawaDBflag Then
            SagawaUpdate(serverUpdateTime)
        End If
    End Sub

    Private Function SagawaUpdate(serverUpdateTime As String)
        Try
            Label9.Text = "コピー中"
            Label9.Show()
            Application.DoEvents()
            File.Copy(serverPath, localPath, True)
            File.WriteAllText(localTimePath, serverUpdateTime)
            Label9.Text = "計算中"
            Label9.Visible = False
            Application.DoEvents()
        Catch ex As Exception
            Label9.Text = "計算中"
            Label9.Visible = False
            Application.DoEvents()
            MsgBox(ex.Message)
            Return False
            Exit Function
        End Try
        MsgBox("佐川DBデータベースを更新しました")
        Return True
    End Function

    Private Sub 住所DBアップデートToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 住所DBアップデートToolStripMenuItem.Click
        Dim addressDBflag As Boolean = False

        If Not File.Exists(kenallLocalTimePath) Then
            File.WriteAllText(kenallLocalTimePath, "")
        End If

        Dim localUpdateTime As String = File.ReadAllText(kenallLocalTimePath, encSJ)
        Dim serverUpdateTime As String = File.ReadAllText(kenallServerTimePath, encSJ)
        If File.Exists(kenallLocalTimePath) Then
            If localUpdateTime = serverUpdateTime Then
                MsgBox("住所DBデータベースは最新版です")
                addressDBflag = True
            End If
        End If
        If Not addressDBflag Then
            AddressUpdate(serverUpdateTime)
        End If
    End Sub

    Private Function AddressUpdate(serverUpdateTime As String)
        Try
            Label9.Text = "コピー中"
            Label9.Show()
            Application.DoEvents()
            File.Copy(kenallServerPath, kenallLocalPath, True)
            File.WriteAllText(kenallLocalTimePath, serverUpdateTime)
            Label9.Text = "計算中"
            Label9.Visible = False
            Application.DoEvents()
        Catch ex As Exception
            Label9.Text = "計算中"
            Label9.Visible = False
            Application.DoEvents()
            MsgBox(ex.Message)
            Return False
            Exit Function
        End Try
        MsgBox("住所DBデータベースを更新しました")
        Return True
    End Function

    Private Sub TextBox_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox6.KeyDown, TextBox36.KeyDown, TextBox34.KeyDown, TextBox33.KeyDown, TextBox30.KeyDown, TextBox28.KeyDown, TextBox12.KeyDown

    End Sub

    Private Sub Csv_denpyo3_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged
        Csv_denpyo3_F_count.Location = New Point(Me.Location.X, Me.Location.Y + Me.Size.Height)
    End Sub

    Private Sub Csv_denpyo3_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        Csv_denpyo3_F_count.Location = New Point(Me.Location.X, Me.Location.Y + Me.Size.Height)
    End Sub

    'Private Sub Csv_denpyo3_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
    '    Csv_denpyo3_F_count.Close()
    'End Sub

    Private Sub Csv_denpyo3_VisibleChanged(sender As Object, e As EventArgs) Handles Me.VisibleChanged
        If Me.Visible Then
            Csv_denpyo3_F_count.Show()
            Csv_denpyo3_F_count.Location = New Point(Me.Location.X, Me.Location.Y + Me.Size.Height)
        Else
            Csv_denpyo3_F_count.Close()
        End If
    End Sub

    Private Sub Csv_denpyo3_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        Csv_denpyo3_F_count.TopMost = True
        Csv_denpyo3_F_count.TopMost = False
        Me.TopMost = True
        Me.TopMost = False
    End Sub

    Private Sub CsvNeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CsvSagawaToolStripMenuItem.Click
        Csv_denpyo3_F_count.Close()
        Me.Dispose()
        Dim frm As Form = New ChangeCsv
        frm.Show()
    End Sub

    Private Sub 伝票出すToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 伝票出すToolStripMenuItem.Click
        Csv_denpyo3_F_count.Close()
        Me.Dispose()
        Dim frm As Form = New SagawaDenpyo
        frm.Show()
    End Sub

    Private Sub ネコポス出力ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ネコポス出力ToolStripMenuItem.Click
        If KryptonButton3.Enabled = False Then
            MsgBox("ファイルをいれてください。")
            Exit Sub
        End If

        If DGV7.RowCount = 0 And DGV8.RowCount = 0 And DGV9.RowCount = 0 And DGV13.RowCount = 0 Then
            MsgBox("処理してください。")
            Exit Sub
        End If

        'フォルダ選択
        Dim dlg = New VistaFolderBrowserDialog With {
            .RootFolder = Environment.SpecialFolder.Desktop,
            .Description = "フォルダを選択してください"
        }
        Dim saveDir As String = ""
        If dlg.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
            saveDir = dlg.SelectedPath

            Dim fileName As String = "\b2web_" & Format(Now, "yyyyMMddHHmmss") & ".csv"
            Dim dstName As String = saveDir & fileName

            Dim Data_title As String = "お客様管理番号,送り状種類,クール区分,伝票番号,出荷予定日,お届け予定（指定）日,配達時間帯,お届け先コード,お届け先電話番号,お届け先電話番号枝番,お届け先郵便番号,お届け先住所,お届け先住所（アパートマンション名）,お届け先会社・部門名１,お届け先会社・部門名２,お届け先名,お届け先名略称カナ,敬称,ご依頼主コード,ご依頼主電話番号,ご依頼主電話番号枝番,ご依頼主郵便番号,ご依頼主住所,ご依頼主住所（アパートマンション名）,ご依頼主名,ご依頼主略称カナ,品名コード１,品名１,品名コード２,品名２,荷扱い１,荷扱い２,記事,コレクト代金引換額（税込）,コレクト内消費税額等,営業所止置き,営業所コード,発行枚数,個数口枠の印字,ご請求先顧客コード,ご請求先分類コード,運賃管理番号,クロネコwebコレクトデータ登録,クロネコwebコレクト加盟店番号,クロネコwebコレクト申込受付番号１,クロネコwebコレクト申込受付番号２,クロネコwebコレクト申込受付番号３,お届け予定ｅメール利用区分,お届け予定ｅメールe-mailアドレス,入力機種,お届け予定eメールメッセージ,お届け完了eメール利用区分,お届け完了ｅメールe-mailアドレス,お届け完了ｅメールメッセージ,クロネコ収納代行利用区分,収納代行決済ＱＲコード印刷,収納代行請求金額(税込),収納代行内消費税額等,収納代行請求先郵便番号,収納代行請求先住所,収納代行請求先住所（アパートマンション名）,収納代行請求先会社・部門名１,収納代行請求先会社・部門名２,収納代行請求先名(漢字),収納代行請求先名(カナ),収納代行問合せ先名(漢字),収納代行問合せ先郵便番号,収納代行問合せ先住所,収納代行問合せ先住所（アパートマンション名）,収納代行問合せ先電話番号,収納代行管理番号,収納代行品名,収納代行備考,複数口くくりキー,検索キータイトル１,検索キー１,検索キータイトル２,検索キー２,検索キータイトル３,検索キー３,検索キータイトル４,検索キー４,検索キータイトル５,検索キー５,予備,予備,投函予定メール利用区分,投函予定メールe-mailアドレス,投函予定メールメッセージ,投函完了メール（お届け先宛）利用区分,投函完了メール（お届け先宛）e-mailアドレス,投函完了メール（お届け先宛）メールメッセージ,投函完了メール（ご依頼主宛）利用区分,投函完了メール（ご依頼主宛）e-mailアドレス,投函完了メール（ご依頼主宛）メールメッセージ,連携管理番号,通知メールアドレス"

            Dim sw As StreamWriter = New StreamWriter(dstName, False, ENC_SJ)
            sw.Write(Data_title)
            sw.Write(vbLf)

            Dim dH7 As ArrayList = TM_HEADER_GET(DGV7)
            Dim dH8 As ArrayList = TM_HEADER_GET(DGV8)
            Dim dH9 As ArrayList = TM_HEADER_GET(DGV9)
            Dim dH13 As ArrayList = TM_HEADER_GET(DGV13)

            Dim okurizyo_type = "7" '送り状種類 ネコポス
            Dim syukayoteibi As String = Format(Now, "yyyy/MM/dd")
            Dim cool_kubun As String = "0"
            Dim seikyukyakucode As String = "0929801866"

            If DGV7.RowCount > 0 Then
                For r As Integer = 0 To DGV7.RowCount - 1
                    Dim data As String = DGV7.Item(dH7.IndexOf("お客様管理ナンバー"), r).Value & "," 'お客様管理番号
                    data &= okurizyo_type & ","
                    data &= cool_kubun & "," 'クール区分
                    data &= "," '伝票番号
                    data &= syukayoteibi & "," '出荷予定日

                    'お届け予定（指定）日
                    If DGV7.Item(dH7.IndexOf("配達日"), r).Value = "" Then
                        data &= ","
                    Else
                        If IsNumeric(DGV7.Item(dH7.IndexOf("配達日"), r).Value) Then
                            data &= DGV7.Item(dH7.IndexOf("配達日"), r).Value.Substring(0, 4) & "/" & DGV7.Item(dH7.IndexOf("配達日"), r).Value.Substring(4, 2) & "/" & DGV7.Item(dH7.IndexOf("配達日"), r).Value.Substring(6, 2) & ","
                        Else
                            data &= ","
                        End If
                    End If


                    '配達時間帯
                    '0812: 午前中
                    '1416: 14～16時
                    '1618: 16～18時
                    '1820: 18～20時
                    '1921: 19～21時
                    If DGV7.Item(dH7.IndexOf("配達指定時間帯"), r).Value = "01" Then
                        data &= "0812,"
                    ElseIf DGV7.Item(dH7.IndexOf("配達指定時間帯"), r).Value = "12" Or DGV7.Item(dH7.IndexOf("お客様管理ナンバー"), r).Value = "14" Then
                        data &= "1416,"
                    ElseIf DGV7.Item(dH7.IndexOf("配達指定時間帯"), r).Value = "16" Then
                        data &= "1618,"
                    ElseIf DGV7.Item(dH7.IndexOf("配達指定時間帯"), r).Value = "04" Or DGV7.Item(dH7.IndexOf("お客様管理ナンバー"), r).Value = "18" Then
                        data &= "1820,"
                    ElseIf DGV7.Item(dH7.IndexOf("配達指定時間帯"), r).Value = "19" Then
                        data &= "1921,"
                    Else
                        data &= ","
                    End If

                    data &= "," 'お届け先コード
                    data &= DGV7.Item(dH7.IndexOf("お届け先電話番号"), r).Value & "," 'お届け先電話番号
                    data &= "," 'お届け先電話番号枝番
                    data &= DGV7.Item(dH7.IndexOf("お届け先郵便番号"), r).Value & "," 'お届け先郵便番号
                    data &= DGV7.Item(dH7.IndexOf("お届け先住所１"), r).Value & "," 'お届け先住所
                    data &= DGV7.Item(dH7.IndexOf("お届け先住所２"), r).Value & "," 'お届け先住所（アパートマンション名）
                    data &= "," 'お届け先会社・部門名１
                    data &= "," 'お届け先会社・部門名２

                    'お届け先名
                    If InStr(DGV7.Item(dH7.IndexOf("お届け先名称１"), r).Value, ")") Then
                        Dim todokesakimei As String() = Split(DGV7.Item(dH7.IndexOf("お届け先名称１"), r).Value, ")")
                        data &= todokesakimei(1) & ","
                    Else
                        data &= DGV7.Item(dH7.IndexOf("お届け先名称１"), r).Value & ","
                    End If

                    data &= "," 'お届け先名略称カナ
                    data &= "様," '敬称

                    'ご依頼主名称１(e飛伝)
                    'ご依頼主コード(ネコポス)
                    'ご依頼主名(ネコポス)
                    '         ↓

                    '壱
                    'auPAYマーケット KuraNavi
                    'au_KuraNavi
                    'ＫｕｒａＮａｖｉ

                    '弐
                    'Yahoo!FKstyle
                    'Yahoo_Fk
                    'Ｙａｈｏｏ!Ｆｋｓｔｙｌｅ

                    '三
                    '楽天通販の暁
                    'Ra_Akatuki
                    '楽天 通販の暁

                    '四
                    'Amazon通販のトココ
                    'Ama_tokoko
                    'Ａｍａｚｏｎ通販のトココ

                    '伍
                    '楽天 雑貨の国のアリス
                    'Ra_Alice
                    '楽天雑貨の国のアリス

                    '六
                    'Qoo10通販の雑貨倉庫
                    'Qo_zakka
                    'Ｑｏｏ１０通販の雑貨倉庫

                    '七
                    'とんよか卸 -> 万方商事株式会社
                    '万方商事(株)

                    '八
                    'Yahoo!Lucky9
                    'Yahoo_Lucky9
                    'Ｙａｈｏｏ!Ｌｕｃｋｙ９

                    '九
                    'ヤフオク付 -> べんけい（soing1702）

                    '十
                    'サラダ
                    'Ama_Sarada
                    'ＡｍａｚｏｎＳａｒａｄａ

                    '十一
                    'Yahoo!あかねAshop
                    'Yahoo_Akane
                    'Ｙａｈｏｏ!あかねＡｓｈｏｐ

                    '十弐
                    '楽天 あかねAshop
                    'Ra_Akane
                    '楽天あかねＡｓｈｏｐ

                    '壱拾参
                    'フリット - > Amazon Hewflit
                    'Ama_Hewflit
                    'ＡｍａｚｏｎＨｅｗｆｌｉｔ

                    '壱拾弐
                    '雑貨ショップKT 海東
                    'Yahoo_KT
                    'Ｙａｈｏｏ!雑貨ショップＫ・Ｔ

                    '壱拾四
                    'Amazon 雑貨の国のアリス
                    'a_Alice
                    'ａｍａｚｏｎ雑貨の国のアリス

                    If DGV7.Item(dH7.IndexOf("ご依頼主名称１"), r).Value = "auPAYマーケット KuraNavi" Then
                        data &= "au_KuraNavi," 'ご依頼主コード
                        data &= "092-980-1144," 'ご依頼主電話番号
                        data &= "," 'ご依頼主電話番号枝番
                        data &= "812-0881," 'ご依頼主郵便番号
                        data &= "福岡県福岡市博多区井相田１－８－３３," 'ご依頼主住所
                        data &= "au PAY マーケット," 'ご依頼主住所（アパートマンション名）
                        data &= "ＫｕｒａＮａｖｉ," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV7.Item(dH7.IndexOf("ご依頼主名称１"), r).Value = "Yahoo!FKstyle" Then
                        data &= "Yahoo_Fk," 'ご依頼主コード
                        data &= "092-586-6853," 'ご依頼主電話番号
                        data &= "2," 'ご依頼主電話番号枝番
                        data &= "816-0911," 'ご依頼主郵便番号
                        data &= "福岡県大野城市大城4-13-15," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "Ｙａｈｏｏ！Ｆｋｓｔｙｌｅ," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV7.Item(dH7.IndexOf("ご依頼主名称１"), r).Value = "楽天 通販の暁" Then
                        data &= "Ra_Akatuki," 'ご依頼主コード
                        data &= "092-986-1116," 'ご依頼主電話番号
                        data &= "," 'ご依頼主電話番号枝番
                        data &= "812-0881," 'ご依頼主郵便番号
                        data &= "福岡県福岡市博多区井相田," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "楽天　通販の暁," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV7.Item(dH7.IndexOf("ご依頼主名称１"), r).Value = "Amazon通販のトココ" Then
                        data &= "Ama_tokoko," 'ご依頼主コード
                        data &= "0929801866," 'ご依頼主電話番号
                        data &= "3," 'ご依頼主電話番号枝番
                        data &= "816-0921," 'ご依頼主郵便番号
                        data &= "福岡県大野城市仲畑２丁目６－１６－Ｂ２０２," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "Ａｍａｚｏｎ通販のトココ," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV7.Item(dH7.IndexOf("ご依頼主名称１"), r).Value = "楽天 雑貨の国のアリス" Then
                        data &= "Ra_Alice," 'ご依頼主コード
                        data &= "092-985-2056," 'ご依頼主電話番号
                        data &= "1," 'ご依頼主電話番号枝番
                        data &= "816-0901," 'ご依頼主郵便番号
                        data &= "福岡県大野城市乙金東１－２－５２－１," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "楽天　雑貨の国のアリス," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV7.Item(dH7.IndexOf("ご依頼主名称１"), r).Value = "Qoo10通販の雑貨倉庫" Then
                        data &= "Qo_zakka," 'ご依頼主コード
                        data &= "0929801866," 'ご依頼主電話番号
                        data &= "," 'ご依頼主電話番号枝番
                        data &= "812-0881," 'ご依頼主郵便番号
                        data &= "福岡県福岡市博多区井相田1-8-33," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "Ｑｏｏ１０通販の雑貨倉庫," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV7.Item(dH7.IndexOf("ご依頼主名称１"), r).Value = "万方商事株式会社" Then
                        data &= "," 'ご依頼主コード
                        data &= "080-3987-8690," 'ご依頼主電話番号
                        data &= "," 'ご依頼主電話番号枝番
                        data &= "812-0881," 'ご依頼主郵便番号
                        data &= "福岡県福岡市博多区井相田1-8-33," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "万方商事(株)," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV7.Item(dH7.IndexOf("ご依頼主名称１"), r).Value = "Yahoo!Lucky9" Then
                        data &= "Yahoo_Lucky9," 'ご依頼主コード
                        data &= "092-985-0275," 'ご依頼主電話番号
                        data &= "," 'ご依頼主電話番号枝番
                        data &= "812-0881," 'ご依頼主郵便番号
                        data &= "福岡県福岡市博多区井相田１－８－３３," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "Ｙａｈｏｏ！Ｌｕｃｋｙ９," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV7.Item(dH7.IndexOf("ご依頼主名称１"), r).Value = "べんけい（soing1702）" Then
                        data &= "Yaho_oku," 'ご依頼主コード
                        data &= "0929801866," 'ご依頼主電話番号
                        data &= "6," 'ご依頼主電話番号枝番
                        data &= "812-0881," 'ご依頼主郵便番号
                        data &= "福岡県福岡市博多区井相田２－３－４３－１０２," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "べんけい（ｓｏｉｎｇ１７０２）," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV7.Item(dH7.IndexOf("ご依頼主名称１"), r).Value = "サラダ" Then
                        data &= "Ama_Sarada," 'ご依頼主コード
                        data &= "0929801866," 'ご依頼主電話番号
                        data &= "2," 'ご依頼主電話番号枝番
                        data &= "812-0863," 'ご依頼主郵便番号
                        data &= "福岡県福岡市博多区金の隈３－１６－６２４," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "Ａｍａｚｏｎ　Ｓａｒａｄａ," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV7.Item(dH7.IndexOf("ご依頼主名称１"), r).Value = "Yahoo!あかねAshop" Then
                        data &= "Yahoo_Akane," 'ご依頼主コード
                        data &= "0929850302," 'ご依頼主電話番号
                        data &= "7," 'ご依頼主電話番号枝番
                        data &= "816-0921," 'ご依頼主郵便番号
                        data &= "福岡県大野城市仲畑２丁目６－１６－Ｂ２０２," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "Ｙａｈｏｏ！あかねＡｓｈｏｐ," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV7.Item(dH7.IndexOf("ご依頼主名称１"), r).Value = "楽天 あかねAshop" Then
                        data &= "Ra_Akane," 'ご依頼主コード
                        data &= "0929850295," 'ご依頼主電話番号
                        data &= "5," 'ご依頼主電話番号枝番
                        data &= "812-0881," 'ご依頼主郵便番号
                        data &= "福岡県福岡市博多区井相田１－８－３３－２０１," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "楽天　あかねＡｓｈｏｐ," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV7.Item(dH7.IndexOf("ご依頼主名称１"), r).Value = "Amazon Hewflit" Then
                        data &= "Ama_Hewflit," 'ご依頼主コード
                        data &= "0929801866," 'ご依頼主電話番号
                        data &= "4," 'ご依頼主電話番号枝番
                        data &= "816-0901," 'ご依頼主郵便番号
                        data &= "福岡県大野城市乙金東１－２－５２－１," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "Ａｍａｚｏｎ　Ｈｅｗｆｌｉｔ," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV7.Item(dH7.IndexOf("ご依頼主名称１"), r).Value = "雑貨ショップKT 海東" Then
                        data &= "Yahoo_KT," 'ご依頼主コード
                        data &= "092-986-5538," 'ご依頼主電話番号
                        data &= "," 'ご依頼主電話番号枝番
                        data &= "811-0123," 'ご依頼主郵便番号
                        data &= "福岡県糟屋郡新宮町上府北３丁目６－３," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "Ｙａｈｏｏ！雑貨ショップＫ・Ｔ," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV7.Item(dH7.IndexOf("ご依頼主名称１"), r).Value = "Amazon 雑貨の国のアリス" Then
                        data &= "a_Alice," 'ご依頼主コード
                        data &= "092-985-2056," 'ご依頼主電話番号
                        data &= "2," 'ご依頼主電話番号枝番
                        data &= "816-0901," 'ご依頼主郵便番号
                        data &= "福岡県大野城市乙金東１－２－５２－１," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "ａｍａｚｏｎ雑貨の国のアリス," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    Else
                        data &= "," 'ご依頼主コード
                        data &= DGV7.Item(dH7.IndexOf("ご依頼主電話番号"), r).Value & "," 'ご依頼主電話番号
                        data &= "," 'ご依頼主電話番号枝番
                        data &= DGV7.Item(dH7.IndexOf("ご依頼主郵便番号"), r).Value & "," 'ご依頼主郵便番号
                        data &= DGV7.Item(dH7.IndexOf("ご依頼主住所１"), r).Value & "," 'ご依頼主住所
                        data &= DGV7.Item(dH7.IndexOf("ご依頼主住所２"), r).Value & "," 'ご依頼主住所（アパートマンション名）
                        data &= DGV7.Item(dH7.IndexOf("ご依頼主名称１"), r).Value & "," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    End If

                    data &= "," '品名コード１
                    data &= DGV7.Item(dH7.IndexOf("品名１"), r).Value & "," '品名１
                    data &= "," '品名コード２
                    data &= "," '品名２
                    data &= "," '荷扱い１
                    data &= "," '荷扱い２
                    data &= "," '記事
                    data &= "," 'コレクト代金引換額（税込）
                    data &= "," 'コレクト内消費税額等
                    data &= "0," '営業所止置き
                    data &= "," '営業所コード
                    data &= "1," '発行枚数
                    data &= "," '個数口枠の印字
                    data &= seikyukyakucode & "," 'ご請求先顧客コード
                    data &= "," 'ご請求先分類コード
                    data &= "01," '運賃管理番号
                    data &= "0," 'クロネコwebコレクトデータ登録
                    data &= "," 'クロネコwebコレクト加盟店番号
                    data &= "," 'クロネコwebコレクト申込受付番号１
                    data &= "," 'クロネコwebコレクト申込受付番号２
                    data &= "," 'クロネコwebコレクト申込受付番号３
                    data &= "0," 'お届け予定ｅメール利用区分
                    data &= "," 'お届け予定ｅメールe-mailアドレス
                    data &= "," '入力機種
                    data &= "," 'お届け予定eメールメッセージ
                    data &= "0," 'お届け完了eメール利用区分
                    data &= "," 'お届け完了ｅメールe-mailアドレス
                    data &= "," 'お届け完了ｅメールメッセージ
                    data &= "0," 'クロネコ収納代行利用区分
                    data &= "," '収納代行決済ＱＲコード印刷
                    data &= "," '収納代行請求金額(税込)
                    data &= "," '収納代行内消費税額等
                    data &= "," '収納代行請求先郵便番号
                    data &= "," '収納代行請求先住所
                    data &= "," '収納代行請求先住所（アパートマンション名）
                    data &= "," '収納代行請求先会社・部門名１
                    data &= "," '収納代行請求先会社・部門名２
                    data &= "," '収納代行請求先名(漢字)
                    data &= "," '収納代行請求先名(カナ)
                    data &= "," '収納代行問合せ先名(漢字)
                    data &= "," '収納代行問合せ先郵便番号
                    data &= "," '収納代行問合せ先住所
                    data &= "," '収納代行問合せ先住所（アパートマンション名）
                    data &= "," '収納代行問合せ先電話番号
                    data &= "," '収納代行管理番号
                    data &= "," '収納代行品名
                    data &= "," '収納代行備考
                    data &= "," '複数口くくりキー
                    data &= "," '検索キータイトル１
                    data &= "," '検索キー１
                    data &= "," '検索キータイトル２
                    data &= "," '検索キー２
                    data &= "," '検索キータイトル３
                    data &= "," '検索キー３
                    data &= "," '検索キータイトル４
                    data &= "," '検索キー４
                    data &= "," '検索キータイトル５
                    data &= "," '検索キー５
                    data &= "," '予備
                    data &= "," '予備
                    data &= "0," '投函予定メール利用区分
                    data &= "," '投函予定メールe-mailアドレス
                    data &= "," '投函予定メールメッセージ
                    data &= "0," '投函完了メール（お届け先宛）利用区分
                    data &= "," '投函完了メール（お届け先宛）e-mailアドレス
                    data &= "," '投函完了メール（お届け先宛）メールメッセージ
                    data &= "0," '投函完了メール（ご依頼主宛）利用区分
                    data &= "," '投函完了メール（ご依頼主宛）e-mailアドレス
                    data &= "," '投函完了メール（ご依頼主宛）メールメッセージ
                    data &= "," '連携管理番号
                    data &= "" '通知メールアドレス
                    sw.Write(data)
                    sw.Write(vbLf)
                Next
            End If

            If DGV8.RowCount > 0 Then
                For r As Integer = 0 To DGV8.RowCount - 1
                    Dim data As String = DGV8.Item(dH8.IndexOf("顧客管理番号"), r).Value & "," 'お客様管理番号
                    data &= okurizyo_type & ","
                    data &= cool_kubun & "," 'クール区分
                    data &= "," '伝票番号
                    data &= syukayoteibi & "," '出荷予定日

                    'お届け予定（指定）日
                    If DGV8.Item(dH8.IndexOf("配達指定日"), r).Value = "" Then
                        data &= ","
                    Else
                        If IsNumeric(DGV8.Item(dH8.IndexOf("配達指定日"), r).Value) Then
                            data &= DGV8.Item(dH8.IndexOf("配達指定日"), r).Value.Substring(0, 4) & "/" & DGV8.Item(dH8.IndexOf("配達指定日"), r).Value.Substring(4, 2) & "/" & DGV8.Item(dH8.IndexOf("配達指定日"), r).Value.Substring(6, 2) & ","
                        Else
                            data &= ","
                        End If
                    End If


                    '配達時間帯
                    '0812: 午前中
                    '1416: 14～16時
                    '1618: 16～18時
                    '1820: 18～20時
                    '1921: 19～21時
                    If DGV8.Item(dH8.IndexOf("時間帯コード"), r).Value = "午前中" Then
                        data &= "0812,"
                    ElseIf DGV8.Item(dH8.IndexOf("時間帯コード"), r).Value = "12時-14時" Or DGV8.Item(dH8.IndexOf("時間帯コード"), r).Value = "14時-16時" Then
                        data &= "1416,"
                    ElseIf DGV8.Item(dH8.IndexOf("時間帯コード"), r).Value = "16時-18時" Then
                        data &= "1618,"
                    ElseIf DGV8.Item(dH8.IndexOf("時間帯コード"), r).Value = "18時-21時" Or DGV8.Item(dH8.IndexOf("時間帯コード"), r).Value = "18時-20時" Then
                        data &= "1820,"
                    ElseIf DGV8.Item(dH8.IndexOf("時間帯コード"), r).Value = "19時-21時" Then
                        data &= "1921,"
                    Else
                        data &= ","
                    End If

                    data &= "," 'お届け先コード
                    data &= DGV8.Item(dH8.IndexOf("お届け先電話"), r).Value & "," 'お届け先電話番号
                    data &= "," 'お届け先電話番号枝番
                    data &= DGV8.Item(dH8.IndexOf("お届け先郵便番号"), r).Value & "," 'お届け先郵便番号
                    data &= DGV8.Item(dH8.IndexOf("お届け先住所１"), r).Value & "," 'お届け先住所
                    data &= DGV8.Item(dH8.IndexOf("お届け先住所２"), r).Value & "," 'お届け先住所（アパートマンション名）
                    data &= "," 'お届け先会社・部門名１
                    data &= "," 'お届け先会社・部門名２

                    'お届け先名
                    If InStr(DGV8.Item(dH8.IndexOf("お届け先名１"), r).Value, ")") Then
                        Dim todokesakimei As String() = Split(DGV8.Item(dH8.IndexOf("お届け先名１"), r).Value, ")")
                        data &= todokesakimei(1) & ","
                    Else
                        data &= DGV8.Item(dH8.IndexOf("お届け先名１"), r).Value & ","
                    End If

                    data &= "," 'お届け先名略称カナ
                    data &= "様," '敬称

                    If DGV8.Item(dH8.IndexOf("代行ご依頼主名１"), r).Value = "auPAYマーケット KuraNavi" Then
                        data &= "au_KuraNavi," 'ご依頼主コード
                        data &= "092-980-1144," 'ご依頼主電話番号
                        data &= "," 'ご依頼主電話番号枝番
                        data &= "812-0881," 'ご依頼主郵便番号
                        data &= "福岡県福岡市博多区井相田１－８－３３," 'ご依頼主住所
                        data &= "au PAY マーケット," 'ご依頼主住所（アパートマンション名）
                        data &= "ＫｕｒａＮａｖｉ," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV8.Item(dH8.IndexOf("代行ご依頼主名１"), r).Value = "Yahoo!FKstyle" Then
                        data &= "Yahoo_Fk," 'ご依頼主コード
                        data &= "092-586-6853," 'ご依頼主電話番号
                        data &= "2," 'ご依頼主電話番号枝番
                        data &= "816-0911," 'ご依頼主郵便番号
                        data &= "福岡県大野城市大城4-13-15," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "Ｙａｈｏｏ！Ｆｋｓｔｙｌｅ," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV8.Item(dH8.IndexOf("代行ご依頼主名１"), r).Value = "楽天 通販の暁" Then
                        data &= "Ra_Akatuki," 'ご依頼主コード
                        data &= "092-986-1116," 'ご依頼主電話番号
                        data &= "," 'ご依頼主電話番号枝番
                        data &= "812-0881," 'ご依頼主郵便番号
                        data &= "福岡県福岡市博多区井相田," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "楽天　通販の暁," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV8.Item(dH8.IndexOf("代行ご依頼主名１"), r).Value = "Amazon通販のトココ" Then
                        data &= "Ama_tokoko," 'ご依頼主コード
                        data &= "0929801866," 'ご依頼主電話番号
                        data &= "3," 'ご依頼主電話番号枝番
                        data &= "816-0921," 'ご依頼主郵便番号
                        data &= "福岡県大野城市仲畑２丁目６－１６－Ｂ２０２," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "Ａｍａｚｏｎ通販のトココ," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV8.Item(dH8.IndexOf("代行ご依頼主名１"), r).Value = "楽天 雑貨の国のアリス" Then
                        data &= "Ra_Alice," 'ご依頼主コード
                        data &= "092-985-2056," 'ご依頼主電話番号
                        data &= "1," 'ご依頼主電話番号枝番
                        data &= "816-0901," 'ご依頼主郵便番号
                        data &= "福岡県大野城市乙金東１－２－５２－１," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "楽天　雑貨の国のアリス," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV8.Item(dH8.IndexOf("代行ご依頼主名１"), r).Value = "Qoo10通販の雑貨倉庫" Then
                        data &= "Qo_zakka," 'ご依頼主コード
                        data &= "0929801866," 'ご依頼主電話番号
                        data &= "," 'ご依頼主電話番号枝番
                        data &= "812-0881," 'ご依頼主郵便番号
                        data &= "福岡県福岡市博多区井相田1-8-33," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "Ｑｏｏ１０通販の雑貨倉庫," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV8.Item(dH8.IndexOf("代行ご依頼主名１"), r).Value = "万方商事株式会社" Then
                        data &= "," 'ご依頼主コード
                        data &= "080-3987-8690," 'ご依頼主電話番号
                        data &= "," 'ご依頼主電話番号枝番
                        data &= "812-0881," 'ご依頼主郵便番号
                        data &= "福岡県福岡市博多区井相田1-8-33," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "万方商事(株)," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV8.Item(dH8.IndexOf("代行ご依頼主名１"), r).Value = "Yahoo!Lucky9" Then
                        data &= "Yahoo_Lucky9," 'ご依頼主コード
                        data &= "092-985-0275," 'ご依頼主電話番号
                        data &= "," 'ご依頼主電話番号枝番
                        data &= "812-0881," 'ご依頼主郵便番号
                        data &= "福岡県福岡市博多区井相田１－８－３３," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "Ｙａｈｏｏ！Ｌｕｃｋｙ９," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV8.Item(dH8.IndexOf("代行ご依頼主名１"), r).Value = "べんけい（soing1702）" Then
                        data &= "Yaho_oku," 'ご依頼主コード
                        data &= "0929801866," 'ご依頼主電話番号
                        data &= "6," 'ご依頼主電話番号枝番
                        data &= "812-0881," 'ご依頼主郵便番号
                        data &= "福岡県福岡市博多区井相田２－３－４３－１０２," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "べんけい（ｓｏｉｎｇ１７０２）," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV8.Item(dH8.IndexOf("代行ご依頼主名１"), r).Value = "サラダ" Then
                        data &= "Ama_Sarada," 'ご依頼主コード
                        data &= "0929801866," 'ご依頼主電話番号
                        data &= "2," 'ご依頼主電話番号枝番
                        data &= "812-0863," 'ご依頼主郵便番号
                        data &= "福岡県福岡市博多区金の隈３－１６－６２４," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "Ａｍａｚｏｎ　Ｓａｒａｄａ," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV8.Item(dH8.IndexOf("代行ご依頼主名１"), r).Value = "Yahoo!あかねAshop" Then
                        data &= "Yahoo_Akane," 'ご依頼主コード
                        data &= "0929850302," 'ご依頼主電話番号
                        data &= "7," 'ご依頼主電話番号枝番
                        data &= "816-0921," 'ご依頼主郵便番号
                        data &= "福岡県大野城市仲畑２丁目６－１６－Ｂ２０２," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "Ｙａｈｏｏ！あかねＡｓｈｏｐ," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV8.Item(dH8.IndexOf("代行ご依頼主名１"), r).Value = "楽天 あかねAshop" Then
                        data &= "Ra_Akane," 'ご依頼主コード
                        data &= "0929850295," 'ご依頼主電話番号
                        data &= "5," 'ご依頼主電話番号枝番
                        data &= "812-0881," 'ご依頼主郵便番号
                        data &= "福岡県福岡市博多区井相田１－８－３３－２０１," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "楽天　あかねＡｓｈｏｐ," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV8.Item(dH8.IndexOf("代行ご依頼主名１"), r).Value = "Amazon Hewflit" Then
                        data &= "Ama_Hewflit," 'ご依頼主コード
                        data &= "0929801866," 'ご依頼主電話番号
                        data &= "4," 'ご依頼主電話番号枝番
                        data &= "816-0901," 'ご依頼主郵便番号
                        data &= "福岡県大野城市乙金東１－２－５２－１," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "Ａｍａｚｏｎ　Ｈｅｗｆｌｉｔ," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV8.Item(dH8.IndexOf("代行ご依頼主名１"), r).Value = "雑貨ショップKT 海東" Then
                        data &= "Yahoo_KT," 'ご依頼主コード
                        data &= "092-986-5538," 'ご依頼主電話番号
                        data &= "," 'ご依頼主電話番号枝番
                        data &= "811-0123," 'ご依頼主郵便番号
                        data &= "福岡県糟屋郡新宮町上府北３丁目６－３," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "Ｙａｈｏｏ！雑貨ショップＫ・Ｔ," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV8.Item(dH8.IndexOf("代行ご依頼主名１"), r).Value = "Amazon 雑貨の国のアリス" Then
                        data &= "a_Alice," 'ご依頼主コード
                        data &= "092-985-2056," 'ご依頼主電話番号
                        data &= "2," 'ご依頼主電話番号枝番
                        data &= "816-0901," 'ご依頼主郵便番号
                        data &= "福岡県大野城市乙金東１－２－５２－１," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "ａｍａｚｏｎ雑貨の国のアリス," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    Else
                        data &= "," 'ご依頼主コード
                        data &= DGV8.Item(dH8.IndexOf("代行ご依頼主電話"), r).Value & "," 'ご依頼主電話番号
                        data &= "," 'ご依頼主電話番号枝番
                        data &= DGV8.Item(dH8.IndexOf("代行ご依頼主郵便番号"), r).Value & "," 'ご依頼主郵便番号
                        data &= DGV8.Item(dH8.IndexOf("代行ご依頼主住所１"), r).Value & "," 'ご依頼主住所
                        data &= DGV8.Item(dH8.IndexOf("代行ご依頼主住所２"), r).Value & "," 'ご依頼主住所（アパートマンション名）
                        data &= DGV8.Item(dH8.IndexOf("代行ご依頼主名１"), r).Value & "," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    End If

                    data &= "," '品名コード１
                    data &= DGV8.Item(dH8.IndexOf("記事欄01"), r).Value & "," '品名１
                    data &= "," '品名コード２
                    data &= "," '品名２
                    data &= "," '荷扱い１
                    data &= "," '荷扱い２
                    data &= "," '記事
                    data &= "," 'コレクト代金引換額（税込）
                    data &= "," 'コレクト内消費税額等
                    data &= "0," '営業所止置き
                    data &= "," '営業所コード
                    data &= "1," '発行枚数
                    data &= "," '個数口枠の印字
                    data &= seikyukyakucode & "," 'ご請求先顧客コード
                    data &= "," 'ご請求先分類コード
                    data &= "01," '運賃管理番号
                    data &= "0," 'クロネコwebコレクトデータ登録
                    data &= "," 'クロネコwebコレクト加盟店番号
                    data &= "," 'クロネコwebコレクト申込受付番号１
                    data &= "," 'クロネコwebコレクト申込受付番号２
                    data &= "," 'クロネコwebコレクト申込受付番号３
                    data &= "0," 'お届け予定ｅメール利用区分
                    data &= "," 'お届け予定ｅメールe-mailアドレス
                    data &= "," '入力機種
                    data &= "," 'お届け予定eメールメッセージ
                    data &= "0," 'お届け完了eメール利用区分
                    data &= "," 'お届け完了ｅメールe-mailアドレス
                    data &= "," 'お届け完了ｅメールメッセージ
                    data &= "0," 'クロネコ収納代行利用区分
                    data &= "," '収納代行決済ＱＲコード印刷
                    data &= "," '収納代行請求金額(税込)
                    data &= "," '収納代行内消費税額等
                    data &= "," '収納代行請求先郵便番号
                    data &= "," '収納代行請求先住所
                    data &= "," '収納代行請求先住所（アパートマンション名）
                    data &= "," '収納代行請求先会社・部門名１
                    data &= "," '収納代行請求先会社・部門名２
                    data &= "," '収納代行請求先名(漢字)
                    data &= "," '収納代行請求先名(カナ)
                    data &= "," '収納代行問合せ先名(漢字)
                    data &= "," '収納代行問合せ先郵便番号
                    data &= "," '収納代行問合せ先住所
                    data &= "," '収納代行問合せ先住所（アパートマンション名）
                    data &= "," '収納代行問合せ先電話番号
                    data &= "," '収納代行管理番号
                    data &= "," '収納代行品名
                    data &= "," '収納代行備考
                    data &= "," '複数口くくりキー
                    data &= "," '検索キータイトル１
                    data &= "," '検索キー１
                    data &= "," '検索キータイトル２
                    data &= "," '検索キー２
                    data &= "," '検索キータイトル３
                    data &= "," '検索キー３
                    data &= "," '検索キータイトル４
                    data &= "," '検索キー４
                    data &= "," '検索キータイトル５
                    data &= "," '検索キー５
                    data &= "," '予備
                    data &= "," '予備
                    data &= "0," '投函予定メール利用区分
                    data &= "," '投函予定メールe-mailアドレス
                    data &= "," '投函予定メールメッセージ
                    data &= "0," '投函完了メール（お届け先宛）利用区分
                    data &= "," '投函完了メール（お届け先宛）e-mailアドレス
                    data &= "," '投函完了メール（お届け先宛）メールメッセージ
                    data &= "0," '投函完了メール（ご依頼主宛）利用区分
                    data &= "," '投函完了メール（ご依頼主宛）e-mailアドレス
                    data &= "," '投函完了メール（ご依頼主宛）メールメッセージ
                    data &= "," '連携管理番号
                    data &= "" '通知メールアドレス
                    sw.Write(data)
                    sw.Write(vbLf)
                Next
            End If

            If DGV9.RowCount > 0 Then
                For r As Integer = 0 To DGV9.RowCount - 1
                    Dim data As String = DGV9.Item(dH9.IndexOf("お客様側管理番号"), r).Value & "," 'お客様管理番号
                    data &= okurizyo_type & ","
                    data &= cool_kubun & "," 'クール区分
                    data &= "," '伝票番号
                    data &= syukayoteibi & "," '出荷予定日

                    'お届け予定（指定）日
                    If DGV9.Item(dH9.IndexOf("配送希望日"), r).Value = "" Then
                        data &= ","
                    Else
                        If IsNumeric(DGV9.Item(dH9.IndexOf("配送希望日"), r).Value) Then
                            data &= DGV9.Item(dH9.IndexOf("配送希望日"), r).Value.Substring(0, 4) & "/" & DGV9.Item(dH9.IndexOf("配送希望日"), r).Value.Substring(4, 2) & "/" & DGV9.Item(dH9.IndexOf("配送希望日"), r).Value.Substring(6, 2) & ","
                        Else
                            data &= ","
                        End If
                    End If


                    '配達時間帯
                    '0812: 午前中
                    '1416: 14～16時
                    '1618: 16～18時
                    '1820: 18～20時
                    '1921: 19～21時
                    If DGV9.Item(dH9.IndexOf("配送希望時間帯"), r).Value = "午前中" Then
                        data &= "0812,"
                    ElseIf DGV9.Item(dH9.IndexOf("配送希望時間帯"), r).Value = "12～14時" Or DGV9.Item(dH9.IndexOf("配送希望時間帯"), r).Value = "14～16時" Then
                        data &= "1416,"
                    ElseIf DGV9.Item(dH9.IndexOf("配送希望時間帯"), r).Value = "16～18時" Then
                        data &= "1618,"
                    ElseIf DGV9.Item(dH9.IndexOf("配送希望時間帯"), r).Value = "18～20時" Then
                        data &= "1820,"
                    ElseIf DGV9.Item(dH9.IndexOf("配送希望時間帯"), r).Value = "19～21時" Or DGV9.Item(dH9.IndexOf("配送希望時間帯"), r).Value = "20～21時" Then
                        data &= "1921,"
                    Else
                        data &= ","
                    End If

                    data &= "," 'お届け先コード
                    data &= DGV9.Item(dH9.IndexOf("お届け先　電話番号"), r).Value & "," 'お届け先電話番号
                    data &= "," 'お届け先電話番号枝番
                    data &= DGV9.Item(dH9.IndexOf("お届け先　郵便番号"), r).Value & "," 'お届け先郵便番号
                    data &= DGV9.Item(dH9.IndexOf("お届け先　住所1"), r).Value & "," 'お届け先住所
                    data &= DGV9.Item(dH9.IndexOf("お届け先　住所2"), r).Value & "," 'お届け先住所（アパートマンション名）
                    data &= "," 'お届け先会社・部門名１
                    data &= "," 'お届け先会社・部門名２

                    'お届け先名
                    If InStr(DGV9.Item(dH9.IndexOf("お届け先　名称"), r).Value, ")") Then
                        Dim todokesakimei As String() = Split(DGV9.Item(dH9.IndexOf("お届け先　名称"), r).Value, ")")
                        data &= todokesakimei(1) & ","
                    Else
                        data &= DGV9.Item(dH9.IndexOf("お届け先　名称"), r).Value & ","
                    End If

                    data &= "," 'お届け先名略称カナ
                    data &= "様," '敬称

                    If DGV9.Item(dH9.IndexOf("ご依頼主　名称１"), r).Value = "auPAYマーケット KuraNavi" Then
                        data &= "au_KuraNavi," 'ご依頼主コード
                        data &= "092-980-1144," 'ご依頼主電話番号
                        data &= "," 'ご依頼主電話番号枝番
                        data &= "812-0881," 'ご依頼主郵便番号
                        data &= "福岡県福岡市博多区井相田１－８－３３," 'ご依頼主住所
                        data &= "au PAY マーケット," 'ご依頼主住所（アパートマンション名）
                        data &= "ＫｕｒａＮａｖｉ," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV9.Item(dH9.IndexOf("ご依頼主　名称１"), r).Value = "Yahoo!FKstyle" Then
                        data &= "Yahoo_Fk," 'ご依頼主コード
                        data &= "092-586-6853," 'ご依頼主電話番号
                        data &= "2," 'ご依頼主電話番号枝番
                        data &= "816-0911," 'ご依頼主郵便番号
                        data &= "福岡県大野城市大城4-13-15," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "Ｙａｈｏｏ！Ｆｋｓｔｙｌｅ," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV9.Item(dH9.IndexOf("ご依頼主　名称１"), r).Value = "楽天 通販の暁" Then
                        data &= "Ra_Akatuki," 'ご依頼主コード
                        data &= "092-986-1116," 'ご依頼主電話番号
                        data &= "," 'ご依頼主電話番号枝番
                        data &= "812-0881," 'ご依頼主郵便番号
                        data &= "福岡県福岡市博多区井相田," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "楽天　通販の暁," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV9.Item(dH9.IndexOf("ご依頼主　名称１"), r).Value = "Amazon通販のトココ" Then
                        data &= "Ama_tokoko," 'ご依頼主コード
                        data &= "0929801866," 'ご依頼主電話番号
                        data &= "3," 'ご依頼主電話番号枝番
                        data &= "816-0921," 'ご依頼主郵便番号
                        data &= "福岡県大野城市仲畑２丁目６－１６－Ｂ２０２," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "Ａｍａｚｏｎ通販のトココ," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV9.Item(dH9.IndexOf("ご依頼主　名称１"), r).Value = "楽天 雑貨の国のアリス" Then
                        data &= "Ra_Alice," 'ご依頼主コード
                        data &= "092-985-2056," 'ご依頼主電話番号
                        data &= "1," 'ご依頼主電話番号枝番
                        data &= "816-0901," 'ご依頼主郵便番号
                        data &= "福岡県大野城市乙金東１－２－５２－１," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "楽天　雑貨の国のアリス," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV9.Item(dH9.IndexOf("ご依頼主　名称１"), r).Value = "Qoo10通販の雑貨倉庫" Then
                        data &= "Qo_zakka," 'ご依頼主コード
                        data &= "0929801866," 'ご依頼主電話番号
                        data &= "," 'ご依頼主電話番号枝番
                        data &= "812-0881," 'ご依頼主郵便番号
                        data &= "福岡県福岡市博多区井相田1-8-33," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "Ｑｏｏ１０通販の雑貨倉庫," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV9.Item(dH9.IndexOf("ご依頼主　名称１"), r).Value = "万方商事株式会社" Then
                        data &= "," 'ご依頼主コード
                        data &= "080-3987-8690," 'ご依頼主電話番号
                        data &= "," 'ご依頼主電話番号枝番
                        data &= "812-0881," 'ご依頼主郵便番号
                        data &= "福岡県福岡市博多区井相田1-8-33," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "万方商事(株)," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV9.Item(dH9.IndexOf("ご依頼主　名称１"), r).Value = "Yahoo!Lucky9" Then
                        data &= "Yahoo_Lucky9," 'ご依頼主コード
                        data &= "092-985-0275," 'ご依頼主電話番号
                        data &= "," 'ご依頼主電話番号枝番
                        data &= "812-0881," 'ご依頼主郵便番号
                        data &= "福岡県福岡市博多区井相田１－８－３３," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "Ｙａｈｏｏ！Ｌｕｃｋｙ９," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV9.Item(dH9.IndexOf("ご依頼主　名称１"), r).Value = "べんけい（soing1702）" Then
                        data &= "Yaho_oku," 'ご依頼主コード
                        data &= "0929801866," 'ご依頼主電話番号
                        data &= "6," 'ご依頼主電話番号枝番
                        data &= "812-0881," 'ご依頼主郵便番号
                        data &= "福岡県福岡市博多区井相田２－３－４３－１０２," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "べんけい（ｓｏｉｎｇ１７０２）," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV9.Item(dH9.IndexOf("ご依頼主　名称１"), r).Value = "サラダ" Then
                        data &= "Ama_Sarada," 'ご依頼主コード
                        data &= "0929801866," 'ご依頼主電話番号
                        data &= "2," 'ご依頼主電話番号枝番
                        data &= "812-0863," 'ご依頼主郵便番号
                        data &= "福岡県福岡市博多区金の隈３－１６－６２４," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "Ａｍａｚｏｎ　Ｓａｒａｄａ," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV9.Item(dH9.IndexOf("ご依頼主　名称１"), r).Value = "Yahoo!あかねAshop" Then
                        data &= "Yahoo_Akane," 'ご依頼主コード
                        data &= "0929850302," 'ご依頼主電話番号
                        data &= "7," 'ご依頼主電話番号枝番
                        data &= "816-0921," 'ご依頼主郵便番号
                        data &= "福岡県大野城市仲畑２丁目６－１６－Ｂ２０２," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "Ｙａｈｏｏ！あかねＡｓｈｏｐ," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV9.Item(dH9.IndexOf("ご依頼主　名称１"), r).Value = "楽天 あかねAshop" Then
                        data &= "Ra_Akane," 'ご依頼主コード
                        data &= "0929850295," 'ご依頼主電話番号
                        data &= "5," 'ご依頼主電話番号枝番
                        data &= "812-0881," 'ご依頼主郵便番号
                        data &= "福岡県福岡市博多区井相田１－８－３３－２０１," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "楽天　あかねＡｓｈｏｐ," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV9.Item(dH9.IndexOf("ご依頼主　名称１"), r).Value = "Amazon Hewflit" Then
                        data &= "Ama_Hewflit," 'ご依頼主コード
                        data &= "0929801866," 'ご依頼主電話番号
                        data &= "4," 'ご依頼主電話番号枝番
                        data &= "816-0901," 'ご依頼主郵便番号
                        data &= "福岡県大野城市乙金東１－２－５２－１," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "Ａｍａｚｏｎ　Ｈｅｗｆｌｉｔ," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV9.Item(dH9.IndexOf("ご依頼主　名称１"), r).Value = "雑貨ショップKT 海東" Then
                        data &= "Yahoo_KT," 'ご依頼主コード
                        data &= "092-986-5538," 'ご依頼主電話番号
                        data &= "," 'ご依頼主電話番号枝番
                        data &= "811-0123," 'ご依頼主郵便番号
                        data &= "福岡県糟屋郡新宮町上府北３丁目６－３," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "Ｙａｈｏｏ！雑貨ショップＫ・Ｔ," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV9.Item(dH9.IndexOf("ご依頼主　名称１"), r).Value = "Amazon 雑貨の国のアリス" Then
                        data &= "a_Alice," 'ご依頼主コード
                        data &= "092-985-2056," 'ご依頼主電話番号
                        data &= "2," 'ご依頼主電話番号枝番
                        data &= "816-0901," 'ご依頼主郵便番号
                        data &= "福岡県大野城市乙金東１－２－５２－１," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "ａｍａｚｏｎ雑貨の国のアリス," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ


                    ElseIf DGV9.Item(dH9.IndexOf("ご依頼主　名称１"), r).Value = "AZFK" Then

                        Debug.WriteLine(data)



                    ElseIf DGV9.Item(dH9.IndexOf("ご依頼主　名称１"), r).Value = "AZ海東" Then

                        Debug.WriteLine(data)

                    ElseIf DGV9.Item(dH9.IndexOf("ご依頼主　名称１"), r).Value = "Amazon 雑貨の国のアリス" Then

                    Else
                        data &= "," 'ご依頼主コード
                        data &= DGV9.Item(dH9.IndexOf("ご依頼主　電話番号"), r).Value & "," 'ご依頼主電話番号
                        data &= "," 'ご依頼主電話番号枝番
                        data &= DGV9.Item(dH9.IndexOf("ご依頼主　郵便番号"), r).Value & "," 'ご依頼主郵便番号
                        data &= DGV9.Item(dH9.IndexOf("ご依頼主　住所１"), r).Value & "," 'ご依頼主住所
                        data &= DGV9.Item(dH9.IndexOf("ご依頼主　住所２"), r).Value & "," 'ご依頼主住所（アパートマンション名）
                        data &= DGV9.Item(dH9.IndexOf("ご依頼主　名称１"), r).Value & "," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    End If

                    data &= "," '品名コード１
                    data &= DGV9.Item(dH9.IndexOf("品名"), r).Value & "," '品名１
                    data &= "," '品名コード２
                    data &= "," '品名２
                    data &= "," '荷扱い１
                    data &= "," '荷扱い２
                    data &= "," '記事
                    data &= "," 'コレクト代金引換額（税込）
                    data &= "," 'コレクト内消費税額等
                    data &= "0," '営業所止置き
                    data &= "," '営業所コード
                    data &= "1," '発行枚数
                    data &= "," '個数口枠の印字
                    data &= seikyukyakucode & "," 'ご請求先顧客コード
                    data &= "," 'ご請求先分類コード
                    data &= "01," '運賃管理番号
                    data &= "0," 'クロネコwebコレクトデータ登録
                    data &= "," 'クロネコwebコレクト加盟店番号
                    data &= "," 'クロネコwebコレクト申込受付番号１
                    data &= "," 'クロネコwebコレクト申込受付番号２
                    data &= "," 'クロネコwebコレクト申込受付番号３
                    data &= "0," 'お届け予定ｅメール利用区分
                    data &= "," 'お届け予定ｅメールe-mailアドレス
                    data &= "," '入力機種
                    data &= "," 'お届け予定eメールメッセージ
                    data &= "0," 'お届け完了eメール利用区分
                    data &= "," 'お届け完了ｅメールe-mailアドレス
                    data &= "," 'お届け完了ｅメールメッセージ
                    data &= "0," 'クロネコ収納代行利用区分
                    data &= "," '収納代行決済ＱＲコード印刷
                    data &= "," '収納代行請求金額(税込)
                    data &= "," '収納代行内消費税額等
                    data &= "," '収納代行請求先郵便番号
                    data &= "," '収納代行請求先住所
                    data &= "," '収納代行請求先住所（アパートマンション名）
                    data &= "," '収納代行請求先会社・部門名１
                    data &= "," '収納代行請求先会社・部門名２
                    data &= "," '収納代行請求先名(漢字)
                    data &= "," '収納代行請求先名(カナ)
                    data &= "," '収納代行問合せ先名(漢字)
                    data &= "," '収納代行問合せ先郵便番号
                    data &= "," '収納代行問合せ先住所
                    data &= "," '収納代行問合せ先住所（アパートマンション名）
                    data &= "," '収納代行問合せ先電話番号
                    data &= "," '収納代行管理番号
                    data &= "," '収納代行品名
                    data &= "," '収納代行備考
                    data &= "," '複数口くくりキー
                    data &= "," '検索キータイトル１
                    data &= "," '検索キー１
                    data &= "," '検索キータイトル２
                    data &= "," '検索キー２
                    data &= "," '検索キータイトル３
                    data &= "," '検索キー３
                    data &= "," '検索キータイトル４
                    data &= "," '検索キー４
                    data &= "," '検索キータイトル５
                    data &= "," '検索キー５
                    data &= "," '予備
                    data &= "," '予備
                    data &= "0," '投函予定メール利用区分
                    data &= "," '投函予定メールe-mailアドレス
                    data &= "," '投函予定メールメッセージ
                    data &= "0," '投函完了メール（お届け先宛）利用区分
                    data &= "," '投函完了メール（お届け先宛）e-mailアドレス
                    data &= "," '投函完了メール（お届け先宛）メールメッセージ
                    data &= "0," '投函完了メール（ご依頼主宛）利用区分
                    data &= "," '投函完了メール（ご依頼主宛）e-mailアドレス
                    data &= "," '投函完了メール（ご依頼主宛）メールメッセージ
                    data &= "," '連携管理番号
                    data &= "" '通知メールアドレス
                    sw.Write(data)
                    sw.Write(vbLf)
                Next
            End If

            If DGV13.RowCount > 0 Then
                For r As Integer = 0 To DGV13.RowCount - 1
                    Dim data As String = DGV13.Item(dH13.IndexOf("お客様側管理番号"), r).Value & "," 'お客様管理番号
                    data &= okurizyo_type & ","
                    data &= cool_kubun & "," 'クール区分
                    data &= "," '伝票番号
                    data &= syukayoteibi & "," '出荷予定日

                    'お届け予定（指定）日
                    If DGV13.Item(dH13.IndexOf("配送希望日"), r).Value = "" Then
                        data &= ","
                    Else
                        If IsNumeric(DGV13.Item(dH13.IndexOf("配送希望日"), r).Value) Then
                            data &= DGV13.Item(dH13.IndexOf("配送希望日"), r).Value.Substring(0, 4) & "/" & DGV13.Item(dH13.IndexOf("配送希望日"), r).Value.Substring(4, 2) & "/" & DGV13.Item(dH13.IndexOf("配送希望日"), r).Value.Substring(6, 2) & ","
                        Else
                            data &= ","
                        End If
                    End If


                    '配達時間帯
                    '0812: 午前中
                    '1416: 14～16時
                    '1618: 16～18時
                    '1820: 18～20時
                    '1921: 19～21時
                    If DGV13.Item(dH13.IndexOf("配送希望時間帯"), r).Value = "午前中" Then
                        data &= "0812,"
                    ElseIf DGV13.Item(dH13.IndexOf("配送希望時間帯"), r).Value = "12～14時" Or DGV13.Item(dH13.IndexOf("配送希望時間帯"), r).Value = "14～16時" Then
                        data &= "1416,"
                    ElseIf DGV13.Item(dH13.IndexOf("配送希望時間帯"), r).Value = "16～18時" Then
                        data &= "1618,"
                    ElseIf DGV13.Item(dH13.IndexOf("配送希望時間帯"), r).Value = "18～20時" Then
                        data &= "1820,"
                    ElseIf DGV13.Item(dH13.IndexOf("配送希望時間帯"), r).Value = "19～21時" Or DGV13.Item(dH13.IndexOf("配送希望時間帯"), r).Value = "20～21時" Then
                        data &= "1921,"
                    Else
                        data &= ","
                    End If

                    data &= "," 'お届け先コード
                    data &= DGV13.Item(dH13.IndexOf("お届け先　電話番号"), r).Value & "," 'お届け先電話番号
                    data &= "," 'お届け先電話番号枝番
                    data &= DGV13.Item(dH13.IndexOf("お届け先　郵便番号"), r).Value & "," 'お届け先郵便番号
                    data &= DGV13.Item(dH13.IndexOf("お届け先　住所1"), r).Value & "," 'お届け先住所
                    data &= DGV13.Item(dH13.IndexOf("お届け先　住所2"), r).Value & "," 'お届け先住所（アパートマンション名）
                    data &= "," 'お届け先会社・部門名１
                    data &= "," 'お届け先会社・部門名２

                    'お届け先名
                    If InStr(DGV13.Item(dH13.IndexOf("お届け先　名称"), r).Value, ")") Then
                        Dim todokesakimei As String() = Split(DGV13.Item(dH13.IndexOf("お届け先　名称"), r).Value, ")")
                        data &= todokesakimei(1) & ","
                    Else
                        data &= DGV13.Item(dH13.IndexOf("お届け先　名称"), r).Value & ","
                    End If

                    data &= "," 'お届け先名略称カナ
                    data &= "様," '敬称

                    If DGV13.Item(dH13.IndexOf("ご依頼主　名称１"), r).Value = "auPAYマーケット KuraNavi" Then
                        data &= "au_KuraNavi," 'ご依頼主コード
                        data &= "092-980-1144," 'ご依頼主電話番号
                        data &= "," 'ご依頼主電話番号枝番
                        data &= "812-0881," 'ご依頼主郵便番号
                        data &= "福岡県福岡市博多区井相田１－８－３３," 'ご依頼主住所
                        data &= "au PAY マーケット," 'ご依頼主住所（アパートマンション名）
                        data &= "ＫｕｒａＮａｖｉ," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV13.Item(dH13.IndexOf("ご依頼主　名称１"), r).Value = "Yahoo!FKstyle" Then
                        data &= "Yahoo_Fk," 'ご依頼主コード
                        data &= "092-586-6853," 'ご依頼主電話番号
                        data &= "2," 'ご依頼主電話番号枝番
                        data &= "816-0911," 'ご依頼主郵便番号
                        data &= "福岡県大野城市大城4-13-15," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "Ｙａｈｏｏ！Ｆｋｓｔｙｌｅ," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV13.Item(dH13.IndexOf("ご依頼主　名称１"), r).Value = "楽天 通販の暁" Then
                        data &= "Ra_Akatuki," 'ご依頼主コード
                        data &= "092-986-1116," 'ご依頼主電話番号
                        data &= "," 'ご依頼主電話番号枝番
                        data &= "812-0881," 'ご依頼主郵便番号
                        data &= "福岡県福岡市博多区井相田," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "楽天　通販の暁," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV13.Item(dH13.IndexOf("ご依頼主　名称１"), r).Value = "Amazon通販のトココ" Then
                        data &= "Ama_tokoko," 'ご依頼主コード
                        data &= "0929801866," 'ご依頼主電話番号
                        data &= "3," 'ご依頼主電話番号枝番
                        data &= "816-0921," 'ご依頼主郵便番号
                        data &= "福岡県大野城市仲畑２丁目６－１６－Ｂ２０２," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "Ａｍａｚｏｎ通販のトココ," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV13.Item(dH13.IndexOf("ご依頼主　名称１"), r).Value = "楽天 雑貨の国のアリス" Then
                        data &= "Ra_Alice," 'ご依頼主コード
                        data &= "092-985-2056," 'ご依頼主電話番号
                        data &= "1," 'ご依頼主電話番号枝番
                        data &= "816-0901," 'ご依頼主郵便番号
                        data &= "福岡県大野城市乙金東１－２－５２－１," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "楽天　雑貨の国のアリス," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV13.Item(dH13.IndexOf("ご依頼主　名称１"), r).Value = "Qoo10通販の雑貨倉庫" Then
                        data &= "Qo_zakka," 'ご依頼主コード
                        data &= "0929801866," 'ご依頼主電話番号
                        data &= "," 'ご依頼主電話番号枝番
                        data &= "812-0881," 'ご依頼主郵便番号
                        data &= "福岡県福岡市博多区井相田1-8-33," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "Ｑｏｏ１０通販の雑貨倉庫," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV13.Item(dH13.IndexOf("ご依頼主　名称１"), r).Value = "万方商事株式会社" Then
                        data &= "," 'ご依頼主コード
                        data &= "080-3987-8690," 'ご依頼主電話番号
                        data &= "," 'ご依頼主電話番号枝番
                        data &= "812-0881," 'ご依頼主郵便番号
                        data &= "福岡県福岡市博多区井相田1-8-33," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "万方商事(株)," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV13.Item(dH13.IndexOf("ご依頼主　名称１"), r).Value = "Yahoo!Lucky9" Then
                        data &= "Yahoo_Lucky9," 'ご依頼主コード
                        data &= "092-985-0275," 'ご依頼主電話番号
                        data &= "," 'ご依頼主電話番号枝番
                        data &= "812-0881," 'ご依頼主郵便番号
                        data &= "福岡県福岡市博多区井相田１－８－３３," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "Ｙａｈｏｏ！Ｌｕｃｋｙ９," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV13.Item(dH13.IndexOf("ご依頼主　名称１"), r).Value = "べんけい（soing1702）" Then
                        data &= "Yaho_oku," 'ご依頼主コード
                        data &= "0929801866," 'ご依頼主電話番号
                        data &= "6," 'ご依頼主電話番号枝番
                        data &= "812-0881," 'ご依頼主郵便番号
                        data &= "福岡県福岡市博多区井相田２－３－４３－１０２," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "べんけい（ｓｏｉｎｇ１７０２）," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV13.Item(dH13.IndexOf("ご依頼主　名称１"), r).Value = "サラダ" Then
                        data &= "Ama_Sarada," 'ご依頼主コード
                        data &= "0929801866," 'ご依頼主電話番号
                        data &= "2," 'ご依頼主電話番号枝番
                        data &= "812-0863," 'ご依頼主郵便番号
                        data &= "福岡県福岡市博多区金の隈３－１６－６２４," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "Ａｍａｚｏｎ　Ｓａｒａｄａ," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV13.Item(dH13.IndexOf("ご依頼主　名称１"), r).Value = "Yahoo!あかねAshop" Then
                        data &= "Yahoo_Akane," 'ご依頼主コード
                        data &= "0929801866," 'ご依頼主電話番号
                        data &= "7," 'ご依頼主電話番号枝番
                        data &= "816-0921," 'ご依頼主郵便番号
                        data &= "福岡県大野城市仲畑２丁目６－１６－Ｂ２０２," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "Ｙａｈｏｏ！あかねＡｓｈｏｐ," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV13.Item(dH13.IndexOf("ご依頼主　名称１"), r).Value = "楽天 あかねAshop" Then
                        data &= "Ra_Akane," 'ご依頼主コード
                        data &= "0929801866," 'ご依頼主電話番号
                        data &= "5," 'ご依頼主電話番号枝番
                        data &= "812-0881," 'ご依頼主郵便番号
                        data &= "福岡県福岡市博多区井相田１－８－３３－２０１," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "楽天　あかねＡｓｈｏｐ," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV13.Item(dH13.IndexOf("ご依頼主　名称１"), r).Value = "Amazon Hewflit" Then
                        data &= "Ama_Hewflit," 'ご依頼主コード
                        data &= "0929801866," 'ご依頼主電話番号
                        data &= "4," 'ご依頼主電話番号枝番
                        data &= "816-0901," 'ご依頼主郵便番号
                        data &= "福岡県大野城市乙金東１－２－５２－１," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "Ａｍａｚｏｎ　Ｈｅｗｆｌｉｔ," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV13.Item(dH13.IndexOf("ご依頼主　名称１"), r).Value = "雑貨ショップKT 海東" Then
                        data &= "Yahoo_KT," 'ご依頼主コード
                        data &= "092-986-5538," 'ご依頼主電話番号
                        data &= "," 'ご依頼主電話番号枝番
                        data &= "811-0123," 'ご依頼主郵便番号
                        data &= "福岡県糟屋郡新宮町上府北３丁目６－３," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "Ｙａｈｏｏ！雑貨ショップＫ・Ｔ," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    ElseIf DGV13.Item(dH13.IndexOf("ご依頼主　名称１"), r).Value = "Amazon 雑貨の国のアリス" Then
                        data &= "a_Alice," 'ご依頼主コード
                        data &= "092-985-2056," 'ご依頼主電話番号
                        data &= "2," 'ご依頼主電話番号枝番
                        data &= "816-0901," 'ご依頼主郵便番号
                        data &= "福岡県大野城市乙金東１－２－５２－１," 'ご依頼主住所
                        data &= "," 'ご依頼主住所（アパートマンション名）
                        data &= "ａｍａｚｏｎ雑貨の国のアリス," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    Else
                        data &= "," 'ご依頼主コード
                        data &= DGV13.Item(dH13.IndexOf("ご依頼主　電話番号"), r).Value & "," 'ご依頼主電話番号
                        data &= "," 'ご依頼主電話番号枝番
                        data &= DGV13.Item(dH13.IndexOf("ご依頼主　郵便番号"), r).Value & "," 'ご依頼主郵便番号
                        data &= DGV13.Item(dH13.IndexOf("ご依頼主　住所１"), r).Value & "," 'ご依頼主住所
                        data &= DGV13.Item(dH13.IndexOf("ご依頼主　住所２"), r).Value & "," 'ご依頼主住所（アパートマンション名）
                        data &= DGV13.Item(dH13.IndexOf("ご依頼主　名称１"), r).Value & "," 'ご依頼主名
                        data &= "," 'ご依頼主略称カナ
                    End If

                    data &= "," '品名コード１
                    data &= DGV13.Item(dH13.IndexOf("品名"), r).Value & "," '品名１
                    data &= "," '品名コード２
                    data &= "," '品名２
                    data &= "," '荷扱い１
                    data &= "," '荷扱い２
                    data &= "," '記事
                    data &= "," 'コレクト代金引換額（税込）
                    data &= "," 'コレクト内消費税額等
                    data &= "0," '営業所止置き
                    data &= "," '営業所コード
                    data &= "1," '発行枚数
                    data &= "," '個数口枠の印字
                    data &= seikyukyakucode & "," 'ご請求先顧客コード
                    data &= "," 'ご請求先分類コード
                    data &= "01," '運賃管理番号
                    data &= "0," 'クロネコwebコレクトデータ登録
                    data &= "," 'クロネコwebコレクト加盟店番号
                    data &= "," 'クロネコwebコレクト申込受付番号１
                    data &= "," 'クロネコwebコレクト申込受付番号２
                    data &= "," 'クロネコwebコレクト申込受付番号３
                    data &= "0," 'お届け予定ｅメール利用区分
                    data &= "," 'お届け予定ｅメールe-mailアドレス
                    data &= "," '入力機種
                    data &= "," 'お届け予定eメールメッセージ
                    data &= "0," 'お届け完了eメール利用区分
                    data &= "," 'お届け完了ｅメールe-mailアドレス
                    data &= "," 'お届け完了ｅメールメッセージ
                    data &= "0," 'クロネコ収納代行利用区分
                    data &= "," '収納代行決済ＱＲコード印刷
                    data &= "," '収納代行請求金額(税込)
                    data &= "," '収納代行内消費税額等
                    data &= "," '収納代行請求先郵便番号
                    data &= "," '収納代行請求先住所
                    data &= "," '収納代行請求先住所（アパートマンション名）
                    data &= "," '収納代行請求先会社・部門名１
                    data &= "," '収納代行請求先会社・部門名２
                    data &= "," '収納代行請求先名(漢字)
                    data &= "," '収納代行請求先名(カナ)
                    data &= "," '収納代行問合せ先名(漢字)
                    data &= "," '収納代行問合せ先郵便番号
                    data &= "," '収納代行問合せ先住所
                    data &= "," '収納代行問合せ先住所（アパートマンション名）
                    data &= "," '収納代行問合せ先電話番号
                    data &= "," '収納代行管理番号
                    data &= "," '収納代行品名
                    data &= "," '収納代行備考
                    data &= "," '複数口くくりキー
                    data &= "," '検索キータイトル１
                    data &= "," '検索キー１
                    data &= "," '検索キータイトル２
                    data &= "," '検索キー２
                    data &= "," '検索キータイトル３
                    data &= "," '検索キー３
                    data &= "," '検索キータイトル４
                    data &= "," '検索キー４
                    data &= "," '検索キータイトル５
                    data &= "," '検索キー５
                    data &= "," '予備
                    data &= "," '予備
                    data &= "0," '投函予定メール利用区分
                    data &= "," '投函予定メールe-mailアドレス
                    data &= "," '投函予定メールメッセージ
                    data &= "0," '投函完了メール（お届け先宛）利用区分
                    data &= "," '投函完了メール（お届け先宛）e-mailアドレス
                    data &= "," '投函完了メール（お届け先宛）メールメッセージ
                    data &= "0," '投函完了メール（ご依頼主宛）利用区分
                    data &= "," '投函完了メール（ご依頼主宛）e-mailアドレス
                    data &= "," '投函完了メール（ご依頼主宛）メールメッセージ
                    data &= "," '連携管理番号
                    data &= "" '通知メールアドレス
                    sw.Write(data)
                    sw.Write(vbLf)
                Next
            End If

            ' CSVファイルクローズ
            sw.Close()
            MsgBox("ファイル(" & fileName & ")を出力しました。")
        Else
            Exit Sub
        End If
    End Sub

    Private Sub 佐川送料比較ToolStripMenuItem_Click(sender As Object, e As EventArgs)
        SagawaSpare.Show()
    End Sub


    Private Sub CheckBox34_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox34.CheckedChanged
        '暂时注释掉

        'yupakucheck = CheckBox34.Checked
        'If (CheckBox34.Checked) Then
        '    yupakucheck = True
        '    MessageBox.Show("You are in the CheckBox.CheckedChanged event.")
        '    isyupakuGoodsOverMax()
        'Else
        '    WriteFile(yupakufName, "350", "300")
        '    yupakucheck = False
        '    MessageBox.Show("写入成功")
        'End If


    End Sub


    '判断是不是法人票
    Public Function CheckIsGroupOrder(purchser As String, address As String) As Boolean


        If purchser = "" Or address = "" Then
            Return False
        End If
        For index = 0 To dataB.Length - 1

            If InStr(purchser, dataB(index)) > 0 Or InStr(address, dataB(index)) Then
                Return True
            End If
        Next
        Return False

    End Function
End Class