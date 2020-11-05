Module Win32API
    'ネットワークドライブ接続用
    Public Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" _
    (ByRef lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Integer) As Integer

    'ネットワークドライブ切断用
    Public Declare Function WNetCancelConnection Lib "mpr.dll" Alias "WNetCancelConnectionA" _
    (ByVal lpszName As String, ByVal bForce As Long) As Long


    Public Structure NETRESOURCE

        Public dwScope As Integer
        Public dwType As Integer
        Public dwDisplayType As Integer
        Public dwUsage As Integer
        Public lpLocalName As String
        Public lpRemoteName As String
        Public lpComment As String
        Public lpProvider As String

    End Structure
End Module
