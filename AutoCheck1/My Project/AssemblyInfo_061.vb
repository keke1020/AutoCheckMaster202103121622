Imports System.Runtime.InteropServices
Imports System.Security.Permissions
Imports System.IO
Imports Microsoft.VisualBasic.FileIO
Imports System.Text.RegularExpressions

Public Class Form1
    '"\HKEY_CURRENT_USER\SOFTWARE\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION"
    'AutoCheck1.exeとAutoCheck1.vshost.exeキーにDword値32bit
    '11001	0x2AF9	Internet Explorer 11, Edgeモード (最新のバージョンでレンダリング)
    '11000	0x2AF8	Internet Explorer 11
    '10001	0x2711	Internet Explorer 10, Standardsモード
    '10000	0x2710	Internet Explorer 10 (!DOCTYPE で指定がある場合は、Standardsモードになります。)
    '9999	0x270F	Internet Explorer 9, Standardsモード
    '9000	0x2710	Internet Explorer 9 (!DOCTYPE で指定がある場合は、Standardsモードになります。)
    '8888	0x22B8	Internet Explorer 8, Standardsモード
    '8000	0x1F40	Internet Explorer 8 (!DOCTYPE で指定がある場合は、Standardsモードになります。)
    '7000	0x1B58	Internet Explorer 7 (!DOCTYPE で指定がある場合は、Standardsモードになります。)
    '[HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_SCRIPTURL_MITIGATION]
    '名前 ChaRa.exe
    'DWORD 値(10進) 1

    '-------------------------------------------------------------------
    Dim secValue As String