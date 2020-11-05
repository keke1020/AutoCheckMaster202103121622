Option Explicit On

Public Class Kanri_F_login
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If MaskedTextBox1.Text = "takashi" Then
            Me.Close()
            Kanri.Show()
            Kanri.BringToFront()
        Else
            Me.Close()
        End If
    End Sub

    Private Sub MaskedTextBox1_KeyUp(sender As Object, e As KeyEventArgs) Handles MaskedTextBox1.KeyUp
        If e.KeyData = Keys.Enter Then
            Button8.PerformClick()
        End If
    End Sub

End Class