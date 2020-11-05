Imports System.IO

Public Class kanri
    Dim secValue As String = ""

    Private CurrentDirectoryM As DirectoryInfo
    Private CurrentDirectoryS As DirectoryInfo
    Dim motoPath As String = ""
    Dim sakiPath As String = ""

    Private Sub kanri_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        secValue = "kanri"   'iniファイルセクション

        motoPath = Path.GetDirectoryName(Form1.appPath)
        connect()

        Form1.Timer1.Enabled = False
    End Sub

    Private Sub kanri_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Form1.Timer1.Enabled = True
    End Sub

    Private Sub mFolder()
        ListView1.Items.Clear()
        CurrentDirectoryM = New DirectoryInfo(motoPath)

        Dim d As IO.DirectoryInfo
        For Each d In CurrentDirectoryM.GetDirectories()
            ListView1.Items.Add(New DirectoryItem(d))
        Next d
        Dim f As IO.FileInfo
        For Each f In CurrentDirectoryM.GetFiles()
            ListView1.Items.Add(New FileItem(f))
        Next f
    End Sub

    Private Sub sFolder()
        ListView2.Items.Clear()
        CurrentDirectoryS = New DirectoryInfo(sakiPath)

        Dim d As IO.DirectoryInfo
        Dim f As IO.FileInfo
        For Each d In CurrentDirectoryS.GetDirectories()
            