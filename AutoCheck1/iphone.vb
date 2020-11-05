Imports System.Runtime.InteropServices
Imports SHDocVw

Public Class iphone
    <DllImport("urlmon.dll", CharSet:=CharSet.Ansi)> _
        Private Shared Function UrlMkSetSessionOption(ByVal dwOption As Integer, ByVal pBuffer As String, ByVal dwBufferLength As Integer, ByVal dwReserved As Integer) As Integer
    End Function
    Const URLMON_OPTION_USERAGENT As Integer = &H10000001
    Public Function ChangeUserAgent(ByVal Agent As String)
        UrlMkSetSessionOption(URLMON_OPTION_USERAGENT, Agent, Agent.Length, 0)
        Return &H10000001
    End Function


    Public Shared UserAgent As String = ""


    Private Sub iphone_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'フォームの境界線をなくす
        Me.FormBorderStyle = FormBorderStyle.None
        '大きさを適当に変更
        Me.Size = New Size(313, 634)
        '透明を指定する
        Me.TransparencyKey = Color.FromArgb(255, 255, 254)

        UserAgent = String.Format("User-Agent: {0}", "Mozilla/5.0 (iPhone; U; CPU like Mac OS X; en)")
        WebBrowser1.Navigate(Form1.TabBrowser1.SelectedTab.WebBrowser.Url, "_self", Nothing, UserAgent)


    End Sub


    Private Sub zoom(ByVal zoomvalue As Integer)
        Try
            Dim MyWeb As Object = Me.WebBrowser1.ActiveXInstance
            MyWeb = Me.WebBrowser1.ActiveXInstance
            MyWeb.ExecWB(OLECMDID.OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT.OLECMDEXECOPT_DONTPROMPTUSER, CObj(zoomvalue), CObj(IntPtr.Zero))
            MyWeb = Nothing
        Catch ex As Exception
            'MessageBox.Show("Error:" & ex.Message)
        End Try
    End Sub

    Private Sub WebBrowser1_DocumentCompleted(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted
        Dim Sizes As Size = WebBrowser1.Document.Body.ScrollRectangle.Size
        Dim per As Double = Math.Round(WebBrowser1.Width / Sizes.Width * 100)
        If InStr(WebBrowser1.Url.ToString, "yahoo") = 0 Then
            per = 70
        End If
        Label1.Text = per & "%"
        zoom(per)
    End Sub

    'Dim isNavigating As Boolean = False

    'Private Sub WebBrowser1_Navigating(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserNavigatingEventArgs) Handles WebBrowser1.Navigating
    '    If (isNavigating = False) Then
    '        Dim url2 As String = WebBrowser1.Url.ToString
    '        isNavigating = True
    '        e.Cancel = True
    '        Navigate(url2)
    '    End If
    'End Sub

    'Private Sub Navigate(ByVal url As String)
    '    WebBrowser1.Navigate(url, "_self", Nothing, UserAgent)
    'End Sub

    'Private Sub WebBrowser1_Navigated(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserNavigatedEventArgs) Handles WebBrowser1.Navigated
    '    isNavigating = False
    'End Sub

    'Closeボタン
    Private Sub Panel2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Panel2.Click
        Me.Close()
    End Sub

    '移行ボタン
    Private Sub Panel3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Panel3.Click
        Dim url As String = Form1.TabBrowser1.SelectedTab.WebBrowser.Url.ToString
        WebBrowser1.Navigate(url, "_self", Nothing, UserAgent)
    End Sub

    'Backボタン
    Private Sub Panel4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Panel4.Click

    End Sub

    'Refreshボタン
    Private Sub Panel5_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Panel5.Click
        Dim url As String = WebBrowser1.Url.ToString
        WebBrowser1.Navigate(url, "_self", Nothing, UserAgent)
    End Sub

    'Followボタン
    Private Sub Panel6_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Panel6.Click

    End Sub
End Class