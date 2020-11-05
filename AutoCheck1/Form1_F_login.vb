Option Explicit On
Imports System.ComponentModel
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports EncodingOperation

Public Class Form1_F_login
    Dim ENC_UTF8 As Encoding = Encoding.UTF8
    Dim loginList1 As ArrayList = New ArrayList     'loginlist.txt
    Dim loginList2 As ArrayList = New ArrayList     'ログイン2.txt

    Dim CotrolArray1 As TextBox()
    Dim CotrolArray2 As TextBox()
    Dim CotrolArray3 As TextBox()
    Dim CotrolArray4 As TextBox()
    Dim CotrolArray5 As TextBox()
    Dim CotrolArray6 As TextBox()
    Dim CotrolArray7 As TextBox()
    Dim CotrolArray8 As TextBox()
    Dim CotrolArray9 As TextBox()
    Dim CotrolArray10 As TextBox()
    Dim CotrolArray11 As TextBox()

    Dim base64 As New MyBase64str("UTF-8")

    Private Sub Opt_F_login_Load(sender As Object, e As EventArgs) Handles Me.Load
        'ファイル読み取り
        'Dim fName1 As String = Path.GetDirectoryName(Form1.appPath) & "\loginlist.dat"
        Dim serverFile As String = Form1.サーバーToolStripMenuItem.Text & "\update\loginlist.dat"
        Try
            Dim readArray As String() = Split(base64.Decode(File.ReadAllText(serverFile, ENC_UTF8)), vbCrLf)
            For Each rA As String In readArray
                If rA <> "" Then
                    loginList1.Add(rA)
                End If
            Next
        Catch ex As Exception

        End Try

        Dim fName2 As String = Path.GetDirectoryName(Form1.appPath) & "\login2.txt"
        Dim lArray2 As String() = File.ReadAllLines(fName2, Encoding.Default)
        For Each lA As String In lArray2
            If lA <> "" Then
                loginList2.Add(lA)
            End If
        Next

        LoginLoad2()
    End Sub

    Dim passArray As New ArrayList
    Private Sub LoginLoad2()
        Dim saveFlag As Boolean = False
        Dim dH3 As ArrayList = TM_HEADER_GET(DGV3)
        For i As Integer = 0 To loginList2.Count - 1
            If InStr(loginList2(i), "|") > 0 Then
                Dim loginLine As String() = Split(loginList2(i), "|")
                DGV3.Rows.Add()
                Dim newRow As Integer = DGV3.RowCount - 1
                Dim loginUser As String() = Split(loginLine(1), ",")
                Dim loginPass As String() = Split(loginLine(2), ",")

                For k As Integer = 0 To loginList1.Count - 1
                    Dim aArray As String() = Split(loginList1(k), ",")
                    If Regex.IsMatch(loginList1(k), "^\-●") Then
                        If loginLine(4) = aArray(1) Then
                            If loginUser(1) <> aArray(2) Then
                                loginUser(1) = aArray(2)
                                saveFlag = True
                            End If
                            If loginPass(1) <> aArray(3) Then
                                loginPass(1) = aArray(3)
                                saveFlag = True
                            End If
                            Exit For
                        End If
                    End If
                Next

                DGV3.Item(dH3.IndexOf("ID"), newRow).Value = loginUser(1)
                DGV3.Item(dH3.IndexOf("パス"), newRow).Value = Regex.Replace(loginPass(1), ".{1}", "*")
                passArray.Add(loginPass(1))
                DGV3.Item(dH3.IndexOf("アドレス"), newRow).Value = loginLine(0)
                DGV3.Item(dH3.IndexOf("IDclass"), newRow).Value = loginUser(0)
                DGV3.Item(dH3.IndexOf("パスclass"), newRow).Value = loginPass(0)
                DGV3.Item(dH3.IndexOf("サイト"), newRow).Value = loginLine(3)

                DGV3.Item(dH3.IndexOf("サイト"), newRow).Style.BackColor = Color.LightCyan
                DGV3.Item(dH3.IndexOf("アドレス"), newRow).Style.BackColor = Color.LightGray
                DGV3.Item(dH3.IndexOf("IDclass"), newRow).Style.BackColor = Color.LightGray
                DGV3.Item(dH3.IndexOf("パスclass"), newRow).Style.BackColor = Color.LightGray
            Else
                Dim loginDefault As String() = Split(loginList2(i), ",")
                Dim search1 As String = ""
                Dim search2 As String = ""
                Select Case loginDefault(0)
                    Case 1
                        RadioButton6.Checked = True
                        search1 = RadioButton6.Text
                    Case 2
                        RadioButton5.Checked = True
                        search1 = RadioButton5.Text
                    Case Else
                        RadioButton4.Checked = True
                        search1 = RadioButton4.Text
                End Select
                Select Case loginDefault(1)
                    Case 1
                        RadioButton7.Checked = True
                        search2 = RadioButton7.Text
                    Case 2
                        RadioButton9.Checked = True
                        search2 = RadioButton9.Text
                    Case 3
                        RadioButton8.Checked = True
                        search2 = RadioButton8.Text
                    Case 4
                        RadioButton10.Checked = True
                        search2 = RadioButton10.Text
                    Case Else
                        RadioButton11.Checked = True
                        search2 = RadioButton11.Text
                End Select
                For r As Integer = 0 To DGV3.RowCount - 1
                    If Regex.IsMatch(DGV3.Item(dH3.IndexOf("サイト"), r).Value, search1 & "|" & search2) Then
                        DGV3.Item(dH3.IndexOf("ID"), r).Style.BackColor = Color.Yellow
                        DGV3.Item(dH3.IndexOf("パス"), r).Style.BackColor = Color.Yellow
                    End If
                Next
            End If
        Next

        If saveFlag Then
            DataSave(0)
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DataSave(1)
    End Sub

    Private Sub DataSave(mode As Integer)
        Dim saveStr As String = ""
        Dim dH3 As ArrayList = TM_HEADER_GET(DGV3)

        For r As Integer = 0 To DGV3.RowCount - 1
            Dim lineStr As String = DGV3.Item(dH3.IndexOf("アドレス"), r).Value
            lineStr &= "|" & DGV3.Item(dH3.IndexOf("IDclass"), r).Value & "," & DGV3.Item(dH3.IndexOf("ID"), r).Value
            lineStr &= "|" & DGV3.Item(dH3.IndexOf("パスclass"), r).Value & "," & passArray(r)
            lineStr &= "|" & DGV3.Item(dH3.IndexOf("サイト"), r).Value
            lineStr &= "|" & TenpoHenkan(DGV3.Item(dH3.IndexOf("サイト"), r).Value)
            lineStr &= vbCrLf
            saveStr &= lineStr
        Next

        Select Case True
            Case RadioButton6.Checked
                saveStr &= "1"
            Case RadioButton5.Checked
                saveStr &= "2"
            Case RadioButton4.Checked
                saveStr &= "3"
        End Select
        Select Case True
            Case RadioButton7.Checked
                saveStr &= "," & "1"
            Case RadioButton9.Checked
                saveStr &= "," & "2"
            Case RadioButton8.Checked
                saveStr &= "," & "3"
            Case RadioButton10.Checked
                saveStr &= "," & "4"
            Case RadioButton11.Checked
                saveStr &= "," & "5"
        End Select

        Dim fName As String = Path.GetDirectoryName(Form1.appPath) & "\login2.txt"
        File.WriteAllText(fName, saveStr, Encoding.GetEncoding("SHIFT-JIS"))

        If mode = 1 Then
            MsgBox("保存しました。現在のページを再読込します。", MsgBoxStyle.SystemModal)
            Dim path As String = Form1.ToolStripTextBox1.Text
            Form1.TabBrowser1.SelectedTab.WebBrowser.Focus()
            Form1.TabBrowser1.SelectedTab.WebBrowser.Navigate(New Uri(path))
            Me.Close()
        End If
    End Sub

    Private Function TenpoHenkan(tenpo As String)
        Dim res As String = ""
        Select Case tenpo
            Case "あかねYahoo(Yahooビジネスマネージャー)"
                res = "Yあかね BM"
            Case "あかねYahoo(Yahooストアクリエイター)"
                res = "Yあかね YL"
            Case "FKstyle(Yahooビジネスマネージャー)"
                res = "FKstyle BM"
            Case "FKstyle(Yahooストアクリエイター)"
                res = "FKstyle YL"
            Case "ラキナイ(Yahooビジネスマネージャー)"
                res = "Lucky9 BM"
            Case "ラキナイ(Yahooストアクリエイター)"
                res = "Lucky9 YL"
            Case "海東ショップ(Yahooビジネスマネージャー)"
                res = "海東KT BM"
            Case "海東ショップ(Yahooストアクリエイター)"
                res = "海東KT YL"
            Case "海東ヤフオク(Yahooビジネスマネージャー)"
                res = "海東ヤフオク BM"
            Case "海東ヤフオク(Yahooログイン)"
                res = "海東ヤフオク YL"
            Case "ネクストエンジン"
                res = "ネクストエンジン"
            Case "メルカリ"
                res = "メルカリ"
            Case "あかね楽天(楽天RMS1)"
                res = "あかね楽"
            Case "あかね楽天(楽天RMS2)"
                res = "あかね楽個人"
            Case "暁(楽天RMS1)"
                res = "暁"
            Case "暁(楽天RMS2)"
                res = "暁個人"
            Case "アリス(楽天RMS1)"
                res = "アリス"
            Case "アリス(楽天RMS2)"
                res = "アリス個人"
            Case "メールディーラー"
                res = "メールディーラー"
        End Select

        Return res
    End Function

    Private Sub RadioButton6_CheckedChanged(sender As Object, e As EventArgs) Handles _
            RadioButton6.CheckedChanged, RadioButton5.CheckedChanged, RadioButton4.CheckedChanged,
            RadioButton7.CheckedChanged, RadioButton9.CheckedChanged, RadioButton8.CheckedChanged, RadioButton10.CheckedChanged, RadioButton11.CheckedChanged

        Dim dH3 As ArrayList = TM_HEADER_GET(DGV3)

        For r As Integer = 0 To DGV3.RowCount - 1
            DGV3.Item(dH3.IndexOf("ID"), r).Style.BackColor = Color.Empty
            DGV3.Item(dH3.IndexOf("パス"), r).Style.BackColor = Color.Empty
        Next

        Dim search1 As String = ""
        Dim search2 As String = ""
        Select Case True
            Case RadioButton6.Checked
                search1 = RadioButton6.Text
            Case RadioButton5.Checked
                search1 = RadioButton5.Text
            Case RadioButton4.Checked
                search1 = RadioButton4.Text
        End Select
        Select Case True
            Case RadioButton7.Checked
                search2 = RadioButton7.Text
            Case RadioButton9.Checked
                search2 = RadioButton9.Text
            Case RadioButton8.Checked
                search2 = RadioButton8.Text
            Case RadioButton10.Checked
                search2 = RadioButton10.Text
            Case RadioButton11.Checked
                search2 = RadioButton11.Text
        End Select

        For r As Integer = 0 To DGV3.RowCount - 1
            If Regex.IsMatch(DGV3.Item(dH3.IndexOf("サイト"), r).Value, search1 & "|" & search2) Then
                DGV3.Item(dH3.IndexOf("ID"), r).Style.BackColor = Color.Yellow
                DGV3.Item(dH3.IndexOf("パス"), r).Style.BackColor = Color.Yellow
            End If
        Next

    End Sub

    Private Sub ContextMenuStrip1_Opening(sender As Object, e As CancelEventArgs) Handles ContextMenuStrip1.Opening
        Select Case DGV3.SelectedCells(0).ColumnIndex
            Case 1, 2
                'コンテキストメニュー開いてOK
            Case Else
                e.Cancel = True
        End Select
    End Sub

    Private Sub 入力ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 入力ToolStripMenuItem.Click
        Dim selCol As Integer = DGV3.SelectedCells(0).ColumnIndex
        Dim selRow As Integer = DGV3.SelectedCells(0).RowIndex
        If selCol = 1 Then
            Dim txt As String = InputBox("IDまたはユーザー名入力", "入力")
            If txt <> "" Then
                DGV3.SelectedCells(0).Value = txt
            End If
        ElseIf selCol = 2 Then
            Dim txt As String = InputBox("パスワード入力", "入力")
            If txt <> "" Then
                DGV3.Item(2, selRow).Value = Regex.Replace(txt, ".{1}", "*")
                passArray(selRow) = txt
            End If
        End If
    End Sub

    Private Sub 削除ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 削除ToolStripMenuItem.Click
        If DGV3.SelectedCells.Count > 0 Then
            Dim selCol As Integer = DGV3.SelectedCells(0).ColumnIndex
            Dim selRow As Integer = DGV3.SelectedCells(0).RowIndex
            DGV3.SelectedCells(0).Value = ""
            DGV3.SelectedCells(0).Style.BackColor = Color.Empty
        End If
    End Sub

    Private Sub パスワードを見るToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles パスワードを見るToolStripMenuItem.Click
        If DGV3.SelectedCells.Count > 0 Then
            Dim selCol As Integer = DGV3.SelectedCells(0).ColumnIndex
            Dim selRow As Integer = DGV3.SelectedCells(0).RowIndex
            If selCol <> 2 Then
                Exit Sub
            End If

            If selCol = 2 And passArray(selRow) <> "" Then
                MsgBox(passArray(selRow))
            End If
        End If
    End Sub
End Class


