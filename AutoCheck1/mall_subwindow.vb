Option Explicit On

Public Class Mall_subwindow


    '行番号を表示する
    Private Sub DataGridView_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles _
        DataGridView2.RowPostPaint
        Dim rect As New Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, sender.RowHeadersWidth - 4, sender.Rows(e.RowIndex).Height)
        TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), sender.RowHeadersDefaultCellStyle.Font, rect, sender.RowHeadersDefaultCellStyle.ForeColor, TextFormatFlags.VerticalCenter Or TextFormatFlags.Right)
    End Sub

    Private Sub WebBrowser1_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted
        Dim MyWeb As Object = WebBrowser1.ActiveXInstance
        'Dim per As Integer = (WebBrowser1.Width / 500) * 100
        Dim per As Integer = (WebBrowser1.Width / 500) * TextBox9.Text
        MyWeb.ExecWB(63, 1, CType(per, Integer), IntPtr.Zero)
    End Sub

    Private Sub TrackBar1_Scroll(sender As Object, e As EventArgs) Handles TrackBar1.Scroll
        Dim MyWeb As Object = WebBrowser1.ActiveXInstance
        MyWeb.ExecWB(63, 1, CType(TrackBar1.Value, Integer), IntPtr.Zero)
        TextBox9.Text = TrackBar1.Value
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TrackBar1.Value = 50
        Dim MyWeb As Object = WebBrowser1.ActiveXInstance
        MyWeb.ExecWB(63, 1, CType(TrackBar1.Value, Integer), IntPtr.Zero)
        TextBox9.Text = TrackBar1.Value
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TrackBar1.Value = 100
        Dim MyWeb As Object = WebBrowser1.ActiveXInstance
        MyWeb.ExecWB(63, 1, CType(TrackBar1.Value, Integer), IntPtr.Zero)
        TextBox9.Text = TrackBar1.Value
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        TrackBar1.Value = 150
        Dim MyWeb As Object = WebBrowser1.ActiveXInstance
        MyWeb.ExecWB(63, 1, CType(TrackBar1.Value, Integer), IntPtr.Zero)
        TextBox9.Text = TrackBar1.Value
    End Sub
End Class