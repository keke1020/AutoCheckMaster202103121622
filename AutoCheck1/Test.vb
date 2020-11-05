Imports System.ComponentModel
Imports System.Runtime.InteropServices

Public Class Test
    Private Sub Test_Load(sender As Object, e As EventArgs) Handles Me.Load
        ''プリンタ名
        ''Dim printerName As String = "SHARP MX-3650FN SPDL2-c"
        'Dim printerName As String = "Adobe PDF"
        ''プリンタ情報を取得する
        'Dim pinfo As PRINTER_INFO_2 = GetPrinterInfo(printerName)
        ''ポートを表示する
        'MsgBox("Port:" & pinfo.pPortName & vbCrLf &
        '    "Status:" & pinfo.Status.ToString() & vbCrLf &
        '    "Comment:" & pinfo.pComment)



        Dim strPrintServer As String, WMIObject As String, PrinterSet As Object, Printer As Object
        strPrintServer = "localhost"
        'strPrintServer = "\\.\root\CIMV2"
        WMIObject = "winmgmts://" & strPrintServer

        PrinterSet = GetObject(WMIObject).InstancesOf("win32_Printer")
        Dim str As String = ""
        For Each Printer In PrinterSet
            str &= Printer.Name & ": " & Printer.PrinterStatus & "/" & Printer.JobCountSinceLastReset & vbCrLf
        Next Printer

    End Sub



    <DllImport("winspool.drv", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function OpenPrinter(
    ByVal pPrinterName As String, ByRef hPrinter As IntPtr,
    ByVal pDefault As IntPtr) As Boolean
    End Function

    <DllImport("winspool.drv", SetLastError:=True)>
    Private Shared Function ClosePrinter(
    ByVal hPrinter As IntPtr) As Boolean
    End Function

    <DllImport("winspool.drv", SetLastError:=True)>
    Private Shared Function GetPrinter(
    ByVal hPrinter As IntPtr, ByVal dwLevel As Integer,
    ByVal pPrinter As IntPtr, ByVal cbBuf As Integer,
    ByRef pcbNeeded As Integer) As Boolean
    End Function

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Public Structure PRINTER_INFO_2
        Public pServerName As String
        Public pPrinterName As String
        Public pShareName As String
        Public pPortName As String
        Public pDriverName As String
        Public pComment As String
        Public pLocation As String
        Public pDevMode As IntPtr
        Public pSepFile As String
        Public pPrintProcessor As String
        Public pDatatype As String
        Public pParameters As String
        Public pSecurityDescriptor As IntPtr
        Public Attributes As System.UInt32
        Public Priority As System.UInt32
        Public DefaultPriority As System.UInt32
        Public StartTime As System.UInt32
        Public UntilTime As System.UInt32
        Public Status As System.UInt32
        Public cJobs As System.UInt32
        Public AveragePPM As System.UInt32
    End Structure

    ''' <summary>
    ''' プリンタの情報をPRINTER_INFO_2で取得する
    ''' </summary>
    ''' <param name="printerName">プリンタ名</param>
    ''' <returns>プリンタの情報</returns>
    Public Shared Function GetPrinterInfo(
    ByVal printerName As String) As PRINTER_INFO_2
        'プリンタのハンドルを取得する
        Dim hPrinter As IntPtr
        If Not OpenPrinter(printerName, hPrinter, IntPtr.Zero) Then
            Throw New Win32Exception(Marshal.GetLastWin32Error())
        End If

        Dim pPrinterInfo As IntPtr = IntPtr.Zero
        Try
            '必要なバイト数を取得する
            Dim needed As Integer
            GetPrinter(hPrinter, 2, IntPtr.Zero, 0, needed)
            If needed <= 0 Then
                Throw New Exception("失敗しました。")
            End If
            'メモリを割り当てる
            pPrinterInfo = Marshal.AllocHGlobal(needed)

            'プリンタ情報を取得する
            Dim temp As Integer
            If Not GetPrinter(hPrinter, 2, pPrinterInfo, needed, temp) Then
                Throw New Win32Exception(Marshal.GetLastWin32Error())
            End If

            'PRINTER_INFO_2型にマーシャリングする
            Dim printerInfo As PRINTER_INFO_2 =
            CType(Marshal.PtrToStructure(
            pPrinterInfo, GetType(PRINTER_INFO_2)), PRINTER_INFO_2)

            '結果を返す
            Return printerInfo
        Finally
            '後始末をする
            ClosePrinter(hPrinter)
            Marshal.FreeHGlobal(pPrinterInfo)
        End Try
    End Function



    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If Now.Second Mod 2 = 0 Then
            BackgroundWorker1.RunWorkerAsync()
            ToolStripStatusLabel1.Text = ToolStripStatusLabel1.Text + 1
        End If
    End Sub

    Private Delegate Sub CallDelegate()
    Private Shared st_textStr1 As String = ""
    Private Sub TBText1()
        TextBox1.Text = TextBox1.Text & vbCrLf & "---" & st_textStr1
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim strPrintServer As String, WMIObject As String, PrinterSet As Object, Printer As Object
        strPrintServer = "localhost"
        'strPrintServer = "\\.\root\CIMV2"
        WMIObject = "winmgmts://" & strPrintServer

        Dim str As String = ""
        PrinterSet = GetObject(WMIObject).InstancesOf("Win32_PrintJob")
        For Each Printer In PrinterSet
            str &= Printer.Name & ": " & Printer.Document & "/" & Printer.Owner & vbCrLf
        Next Printer

        st_textStr1 = str
        Me.Invoke(New CallDelegate(AddressOf TBText1))
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Timer1.Enabled = True
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Timer1.Enabled = False
    End Sub
End Class

