Option Explicit On
Imports System.IO

Public Class Form1
    Dim timerCount As Integer = 0

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        Timer1.Enabled = True
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If timerCount > 3 Then
            Timer1.Enabled = False
            'MsgBox(Path.GetDirectoryName(Reflection.Assembly.GetExecutingAssembly().Location) & "\AutoCheck1.exe")
            Process.Start(Path.GetDirectoryName(Reflection.Assembly.GetExecutingAssembly().Location) & "\AutoCheck1.exe")
            End
        Else
            timerCount += 1
        End If
    End Sub
End Class
