Public Class Registry


    Private Sub Regread(ByVal resStr As String)
        Dim stringValue As String = DirectCast(Microsoft.Win32.Registry.GetValue(resStr, "string", "default"), String)
    End Sub

    Dim keyStr As String = ""
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        TextBox1.Text = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\i8042prt\Parameters"
        If TextBox1.Text = "" Then
            Exit Sub
        End If

        Dim wkReg As Microsoft.Win32.RegistryKey = HKEY(TextBox1.Text)

        'キーを読み取り専用で開く
        Dim regkey As Microsoft.Win32.RegistryKey = wkReg.OpenSubKey(keyStr, False)

        If regkey Is Nothing Then
            MsgBox("レジストリがありません", MsgBoxStyle.SystemModal)
            Exit Sub
        End If

        Dim keyNames() As String = regkey.GetSubKeyNames()
        Dim valueNames() As String = regkey.GetValueNames()
        For i As Integer = 0 To regkey.ValueCount - 1
            DataGridView1.Rows.Add(1)
            DataGridView1.Item(2, i).Value = valueNames(i)
        Next

        regkey.Close()  '閉じる

        'DataGridView1.Rows.Add(1)
        'Dim dType As String = ""
        'Select Case regkey.GetValueKind("String")
        '    Case Microsoft.Win32.RegistryValueKind.String
        '        dType = "REG_SZ"
        '        Dim stringValue As String = DirectCast(regkey.GetValue("Parameters"), String)
        '        DataGridView1.Item(2, 0).Value = stringValue
        '    Case Microsoft.Win32.RegistryValueKind.Binary
        '        dType = "REG_BINARY"
        '    Case Microsoft.Win32.RegistryValueKind.DWord
        '        dType = "REG_DWORD"
        '        Dim intValue As Integer = CInt(regkey.GetValue("Parameters"))
        '        DataGridView1.Item(2, 0).Value = intValue
        '    Case Microsoft.Win32.RegistryValueKind.ExpandString
        '        dType = "REG_EXPAND_SZ"
        '    Case Microsoft.Win32.RegistryValueKind.MultiString
        '        dType = "REG_MULTI_SZ"
        '    Case Microsoft.Win32.RegistryValueKind.QWord
        '        dType = "REG_QWORD"
        '        Dim longVal As Long = CLng(regkey.GetValue("Parameters"))
        '        DataGridView1.Item(2, 0).Value = longVal
        '    Case Microsoft.Win32.RegistryValueKind.Unknown
        '        dType = "UNKNOWN"
        '    Case Else
        '        dType = "NOT"
        'End Select

        'DataGridView1.Item(1, 0).Value = dType

    End Sub

    Private Function HKEY(ByVal str As String) As Microsoft.Win32.RegistryKey
        Dim wkReg As Microsoft.Win32.RegistryKey

        If InStr(str, "HKEY_CLASSES_ROOT") > 0 Then
            wkReg = Microsoft.Win32.Registry.ClassesRoot
            keyStr = Replace(str, "HKEY_CLASSES_ROOT\", "")
        ElseIf InStr(str, "HKEY_CURRENT_USER") > 0 Then
            wkReg = Microsoft.Win32.Registry.Users
            keyStr = Replace(str, "HKEY_CURRENT_USER\", "")
        ElseIf InStr(str, "HKEY_LOCAL_MACHINE") > 0 Then
            wkReg = Microsoft.Win32.Registry.LocalMachine
            keyStr = Replace(str, "HKEY_LOCAL_MACHINE\", "")
        ElseIf InStr(str, "HKEY_USERS") > 0 Then
            wkReg = Microsoft.Win32.Registry.CurrentUser
            keyStr = Replace(str, "HKEY_USERS", "")
        ElseIf InStr(str, "HKEY_CURRENT_CONFIG\") > 0 Then
            wkReg = Microsoft.Win32.Registry.CurrentConfig
            keyStr = Replace(str, "HKEY_CURRENT_CONFIG\", "")
            '--------------------------------------------
            '以下は.net4.5以降廃止
            'ElseIf InStr(str, "HKEY_DYN_DATA") > 0 Then
            '    wkReg = Microsoft.Win32.Registry.DynData
            '    keyStr = Replace(str, "HKEY_DYN_DATA\", "")
            '--------------------------------------------
        ElseIf InStr(str, "HKEY_PERFORMANCE_DATA") > 0 Then
            wkReg = Microsoft.Win32.Registry.PerformanceData
            keyStr = Replace(str, "HKEY_PERFORMANCE_DATA\", "")
        Else
            wkReg = Nothing
            keyStr = ""
        End If

        Return wkReg
    End Function
End Class