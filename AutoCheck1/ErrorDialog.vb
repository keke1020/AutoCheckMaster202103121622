Imports System.IO
Imports System.Windows.Forms

Public Class ErrorDialog

    Private Sub ErrorDialog_Load(sender As Object, e As EventArgs) Handles Me.Load


    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        MailSend()
    End Sub

    Public Sub MailSend()
        Try
            Dim errorMes As String = ""
            Dim smtpserver As String = "210.172.183.37"
            Dim port As Integer = "587"
            Dim userid As String = "pr3650@yongshun006.com"
            Dim pass As String = "$6A2XQ6J"
            Dim address As System.Net.IPAddress = System.Net.Dns.GetHostEntry(smtpserver).AddressList(0)

            Dim logon As TKMP.Net.ISmtpLogon
            'logon = New TKMP.Net.AuthCramMd5(userid, pass) 'AUTH CRAM-MD5でログオンを行ないます
            'logon = New TKMP.Net.AuthLogin(userid, pass)   'AUTH LOGINでログオンを行ないます
            'logon = New TKMP.Net.AuthPlain(userid, pass)   'AUTH PLAINでログオンを行ないます
            logon = New TKMP.Net.AuthAuto(userid, pass)     'CRAM-MD5 PLAIN LOGINの順で利用可能なものを優先してログオンを行ないます
            'logon = New TKMP.Net.PopBeforeSMTP(popclient)  'POP Before SMTPでログオンを行ないます。使用するにはPOPへの接続情報が必要です

            Dim smtp As New TKMP.Net.SmtpClient(address, port, logon)
            If Not smtp.Connect() Then
                errorMes = "logon dismis"
            Else
                Dim sendAddress As String = "ping@yongshun006.com"
                Dim writer As New TKMP.Writer.MailWriter With {
                    .FromAddress = userid
                }
                writer.Headers.Add("From", Environment.MachineName & " <" & userid & ">")
                writer.ToAddressList.Add(sendAddress)
                'writer.ToAddressList.Add(TextBox11.Text)
                writer.Headers.Add("To", sendAddress & " <" & sendAddress & ">")
                writer.Headers.Add("Subject", "エラー:" & Form1.ログイン名ToolStripMenuItem.Text & "/" & Environment.UserName & "（" & Format(Now(), "ymdhmi") & "）")

                Dim str As String = ""
                str &= Form1.ログイン名ToolStripMenuItem.Text & vbCrLf
                str &= Environment.UserName & vbCrLf
                str &= Environment.MachineName & vbCrLf
                str &= Environment.OSVersion.VersionString & vbCrLf
                str &= Environment.UserDomainName & vbCrLf
                str &= vbCrLf
                str &= Application.ProductName & vbCrLf
                str &= Application.ProductVersion & vbCrLf
                For Each formName As Form In Application.OpenForms
                    str &= "Form:" & formName.Name & ">>" & formName.Text & vbCrLf
                Next
                str &= "OpenForms:" & Application.OpenForms.Count & vbCrLf
                str &= vbCrLf
                str &= "[1] " & TextBox5.Text & vbCrLf
                str &= "[2] " & TextBox4.Text & vbCrLf
                str &= "[3] " & TextBox3.Text & vbCrLf
                str &= "[4] " & TextBox2.Text & vbCrLf
                str &= "[5] " & TextBox1.Text & vbCrLf
                str &= vbCrLf
                'str &= "[6] " & TextBox7.Text & vbCrLf

                Dim capturePath As String = Path.GetDirectoryName(Form1.appPath) & "\error.jpg"
                If File.Exists(capturePath) Then
                    Dim part1 As New TKMP.Writer.TextPart(str)
                    Dim part2 As New TKMP.Writer.FilePart(capturePath)
                    Dim mainpart As New TKMP.Writer.MultiPart(part1, part2)
                    writer.MainPart = mainpart
                Else
                    writer.MainPart = New TKMP.Writer.TextPart(str)
                End If

                'Select Case TempFileList.Count
                '    Case 0
                '        writer.MainPart = New TKMP.Writer.TextPart(str & vbCrLf & TextBox10.Text)
                '    Case 1
                '        Dim part1 As New TKMP.Writer.TextPart(str & vbCrLf & TextBox10.Text)
                '        Dim part2 As New TKMP.Writer.FilePart(TempFileList(0))
                '        Dim mainpart As New TKMP.Writer.MultiPart(part1, part2)
                '        writer.MainPart = mainpart
                '    Case 2
                '        Dim part1 As New TKMP.Writer.TextPart(str & vbCrLf & TextBox10.Text)
                '        Dim part2 As New TKMP.Writer.FilePart(TempFileList(0))
                '        Dim part3 As New TKMP.Writer.FilePart(TempFileList(1))
                '        Dim mainpart As New TKMP.Writer.MultiPart(part1, part2, part3)
                '        writer.MainPart = mainpart
                '    Case Else
                '        MsgBox("現在、添付ファイルは2ファイルまでです")
                '        Exit Sub
                'End Select

                smtp.SendMail(writer)
                smtp.Close()
            End If

        Catch ex As Exception

        End Try

        'Me.Close()
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        If InStr(Button1.Text, ">") > 0 Then
            Me.Height = 485
            Button1.Text = "< 詳細"
        Else
            Me.Height = 145
            Button1.Text = "詳細 >"
        End If
    End Sub
End Class
