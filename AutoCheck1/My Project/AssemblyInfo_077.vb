Size,ÚIðÊÝÉp¡²Ið,
Color,ÚIðÊÝÉpc²Ið
,,
Price(Essential),[Åè],0
Qty.(Essential),[Åè],0
Option Code(Optional),[¹]ÚIðÊÝÉp¡²IðqÔ,ÚIðÊÝÉpc²IðqÔ
HS Code(Optional),,
Q-inventory Code(Optional),,
                                                                                                                                                                                                                                                   ç¾å¨ã®ã¹ã¬ããã® CurrentUICulture ãã­ããã£ããªã¼ãã¼ã©ã¤ããã¾ãã
</summary>
</member><member name="T:AutoCheck1.My.Resources.Resources">
	<summary>
  ã­ã¼ã«ã©ã¤ãºãããæå­åãªã©ãæ¤ç´¢ããããã®ãå³å¯ã«åæå®ããããªã½ã¼ã¹ ã¯ã©ã¹ã§ãã
</summary>
</member>
</members>
</doc>                                                                                                                                                                  End Using

        RadioButton3.Checked = True
        ComboBox6.SelectedIndex = 0

        'tenpoãã©ã«ããã§ãã¯
        If Not Directory.Exists(Form1.appPath & "\tenpo") Then
            Dim di2 As System.IO.DirectoryInfo = System.IO.Directory.CreateDirectory(Path.GetDirectoryName(Form1.appPath) & "\tenpo")
        End If

        For i As Integer = 0 To tenpo.Length - 1