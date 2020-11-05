Imports System.IO

Public Class Csv_denpyo3_F_count
    Private serverDir As String = Form1.サーバーToolStripMenuItem.Text & "\denpyoLog\"

    Private Sub Csv_denpyo3_F_count_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.Size = New Size(1033, 88)
        If Form1.AdminFlag Then
            Button2.BackColor = Color.Yellow
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Button2.BackColor = Color.Gainsboro Then
            Button2.BackColor = Color.Yellow
        Else
            Button2.BackColor = Color.Gainsboro
        End If
    End Sub

    Dim timer1Count As Integer = 0
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False

        Dim flag As Integer = 0
        Try
            If Directory.Exists(serverDir) Then
                Dim todayTxtPath As String = serverDir & Format(DateTimePicker1.Value, "yyyyMMdd") & ".txt"
                Dim todayCsvPath As String = serverDir & Format(DateTimePicker1.Value, "yyyyMMdd") & ".csv"
                If File.Exists(todayCsvPath) Then
                    If timer1Count > 50 Or timer1Count = 0 Then
                        Dim line As String = File.ReadLines(todayTxtPath, encSJ)(0)
                        Dim lineArray As String() = Split(line, ",")
                        LB1.Text = lineArray(0)
                        LB2.Text = lineArray(1)
                        LB3.Text = lineArray(2)
                        LB4.Text = lineArray(3)
                        LB5.Text = lineArray(4)
                        LB6.Text = lineArray(5)
                        LB7.Text = lineArray(6)
                        LB8.Text = lineArray(7)
                        LB_d.Text = CInt(LB1.Text) + CInt(LB2.Text) + CInt(LB3.Text) + CInt(LB4.Text)
                        LB_i.Text = CInt(LB5.Text) + CInt(LB6.Text) + CInt(LB7.Text) + CInt(LB8.Text)
                        timer1Count = 0
                    End If

                    flag = 0
                Else
                    flag = 1
                End If
            Else
                flag = 2
            End If
        Catch ex As Exception
            flag = 2
        End Try

        If flag = 0 Then
            If Label18.BackColor = Color.Silver Then
                Label18.BackColor = Color.Yellow
            Else
                Label18.BackColor = Color.Silver
            End If
        ElseIf flag = 1 Then
            If Label18.BackColor = Color.Silver Then
                Label18.BackColor = Color.Orange
            Else
                Label18.BackColor = Color.Silver
            End If
        Else
            If Label18.BackColor = Color.Silver Then
                Label18.BackColor = Color.Red
            Else
                Label18.BackColor = Color.Silver
            End If
        End If

        If flag >= 1 Then
            LB1.Text = 0
            LB2.Text = 0
            LB3.Text = 0
            LB4.Text = 0
            LB5.Text = 0
            LB6.Text = 0
            LB7.Text = 0
            LB8.Text = 0
            LB_d.Text = 0
            LB_i.Text = 0
        End If

        timer1Count += 1
        Timer1.Enabled = True
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        DateTimePicker1.Value = DateAdd(DateInterval.Day, -1, DateTimePicker1.Value)
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        DateTimePicker1.Value = DateAdd(DateInterval.Day, 1, DateTimePicker1.Value)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim savePath As String = ""
        '.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
        Dim sfd As New SaveFileDialog With {
            .InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                .Filter = "CSVファイル(*.csv)|*.csv|すべてのファイル(*.*)|*.*",
                .AutoUpgradeEnabled = False,
                .FilterIndex = 0,
                .Title = "保存先のファイルを選択してください",
                .RestoreDirectory = True,
                .OverwritePrompt = True,
                .CheckPathExists = True
            }

        sfd.FileName = "出力枚数実績_" & Format(Now, "yyyyMMdd") & ".csv"

        'ダイアログを表示する
        If sfd.ShowDialog(Me) = DialogResult.OK Then
            savePath = sfd.FileName
            Dim savePath2 As String = Path.GetDirectoryName(savePath) & "\" & "出力一覧_" & Format(Now, "yyyyMMdd") & ".csv"
            Dim todayCsvPath As String = serverDir & Format(DateTimePicker1.Value, "yyyyMMdd") & ".csv"
            If File.Exists(todayCsvPath) Then
                File.Copy(todayCsvPath, savePath2, True)
            End If

            Dim str As String = ""
            str &= "倉庫,宅配便,航空便,ゆうパケット,定形外" & vbCrLf
            str &= Label8.Text & "," & LB1.Text & "," & LB2.Text & "," & LB3.Text & "," & LB4.Text & "," & vbCrLf
            str &= Label9.Text & "," & LB5.Text & "," & LB6.Text & "," & LB7.Text & "," & LB8.Text & "," & vbCrLf

            File.WriteAllText(savePath, str, encSJ)

            MsgBox("保存しました")
        Else
            Exit Sub
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Csv_denpyo3_F_dgv.CVD3_mode = "change"
        Csv_denpyo3_F_dgv.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Csv_denpyo3_F_dgv.CVD3_mode = "search"
        Csv_denpyo3_F_dgv.Show()
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Csv_denpyo3_F_dgv.DateTimePicker1.Value = DateTimePicker1.Value
    End Sub

End Class
