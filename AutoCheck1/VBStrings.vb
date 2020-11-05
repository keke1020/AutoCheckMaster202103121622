Option Strict On

Public Class VBStrings

#Region "　LeftB メソッド　"

    ' -----------------------------------------------------------------------------------------
    ' <summary>
    '     文字列の左端から指定したバイト数分の文字列を返します。</summary>
    ' <param name="stTarget">
    '     取り出す元になる文字列。<param>
    ' <param name="iByteSize">
    '     取り出すバイト数。</param>
    ' <returns>
    '     左端から指定されたバイト数分の文字列。</returns>
    ' -----------------------------------------------------------------------------------------
    Public Shared Function LeftB(ByVal stTarget As String, ByVal iByteSize As Integer) As String
        Return MidB(stTarget, 1, iByteSize)
    End Function

#End Region

#Region "　MidB メソッド (+1)　"

    ' -----------------------------------------------------------------------------------------
    ' <summary>
    '     文字列の指定されたバイト位置以降のすべての文字列を返します。</summary>
    ' <param name="stTarget">
    '     取り出す元になる文字列。</param>
    ' <param name="iStart">
    '     取り出しを開始する位置。</param>
    ' <returns>
    '     指定されたバイト位置以降のすべての文字列。</returns>
    ' -----------------------------------------------------------------------------------------
    Public Shared Function MidB(ByVal stTarget As String, ByVal iStart As Integer) As String
        Dim hEncoding As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS")
        Dim bBytes As Byte() = hEncoding.GetBytes(stTarget)

        Return hEncoding.GetString(bBytes, iStart - 1, bBytes.Length - iStart + 1)
    End Function


    ' -----------------------------------------------------------------------------------------
    ' <summary>
    '     文字列の指定されたバイト位置から、指定されたバイト数分の文字列を返します。</summary>
    ' <param name="stTarget">
    '     取り出す元になる文字列。</param>
    ' <param name="iStart">
    '     取り出しを開始する位置。</param>
    ' <param name="iByteSize">
    '     取り出すバイト数。</param>
    ' <returns>
    '     指定されたバイト位置から指定されたバイト数分の文字列。</returns>
    ' -----------------------------------------------------------------------------------------
    Public Shared Function MidB _
    (ByVal stTarget As String, ByVal iStart As Integer, ByVal iByteSize As Integer) As String
        Dim hEncoding As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS")
        Dim bBytes As Byte() = hEncoding.GetBytes(stTarget)

        If iByteSize > bBytes.Length - iStart + 1 Then
            iByteSize = bBytes.Length - iStart + 1
        End If
        Return hEncoding.GetString(bBytes, iStart - 1, iByteSize)
    End Function

#End Region

#Region "　RightB メソッド　"

    ' -----------------------------------------------------------------------------------------
    ' <summary>
    '     文字列の右端から指定されたバイト数分の文字列を返します。</summary>
    ' <param name="stTarget">
    '     取り出す元になる文字列。</param>
    ' <param name="iByteSize">
    '     取り出すバイト数。</param>
    ' <returns>
    '     右端から指定されたバイト数分の文字列。</returns>
    ' -----------------------------------------------------------------------------------------
    Public Shared Function RightB(ByVal stTarget As String, ByVal iByteSize As Integer) As String
        Dim hEncoding As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS")
        Dim bBytes As Byte() = hEncoding.GetBytes(stTarget)

        Return hEncoding.GetString(bBytes, bBytes.Length - iByteSize, iByteSize)
    End Function

#End Region

End Class
