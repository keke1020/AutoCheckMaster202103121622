Imports System.Net
Imports System.IO
Imports System.Text.RegularExpressions

Public Class dialog
    Private _TextBox As New List(Of TextBox)
    Dim tempoCodeArray As ArrayList = New ArrayList
    Dim tenpo As String() = New String() {"FKstyle", "Lucky9", "あかねYahoo", "暁", "アリス", "あかね楽天", "KuraNavi", "問よか", "通販の雑貨倉庫"}


    Private Sub dialog_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        _TextBox.Add(TextBox3)
        _TextBox.Add(TextBox5)
        _TextBox.Add(TextBox6)
        _TextBox.Add(TextBox4)
        _TextBox.Add(TextBox8)
        _TextBox.Add(TextBox7)

        'ファイル読み取り
        Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\config\CodeCheck.txt"
        Using sr As New System.IO.StreamReader(fName, System.Text.Encoding.Default)
            Do While Not sr.EndOfStream
                Dim s As String = sr.ReadLine
                tempoCodeArray.Add(s)
            Loop
        End Using

        RadioButton3.Checked = True
        ComboBox6.SelectedIndex = 0

        'tenpoフォルダチェック
        If Not Directory.Exists(Form1.appPath & "\tenpo") Then
            Dim di2 As System.IO.DirectoryInfo = System.IO.Directory.CreateDirectory(Path.GetDirectoryName(Form1.appPath) & "\tenpo")
        End If

        DataGridView8.Rows.Add("マス�