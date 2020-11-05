Imports System.IO
Imports System.Text.RegularExpressions

Public Class Form4

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Form1.Button11_Click(sender, e)
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        TreeView1.CollapseAll()
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        For i As Integer = 0 To TreeView1.Nodes.Count - 1
        Next
    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        Dim str As String = ""
        For Each Node1 As TreeNode In TreeView1.Nodes
            str &= "(" & Node1.Index & ") " & Node1.FullPath & vbNewLine
            If Node1.GetNodeCount(False) > 0 Then
                For Each node2 As TreeNode In Node1.Nodes
                    str &= "　[" & node2.Index & "] " & node2.FullPath & vbNewLine
                    If node2.GetNodeCount(False) > 0 Then
                        Dim num As Integer = 0
                        For Each node3 As TreeNode In node2.Nodes
                            Dim n As String = Regex.Replace(node3.FullPath, "\t|" & vbCr & "|" & vbLf & "|　", "")
                            n = Regex.Replace(n, " {2,100}", " ")
                            str &= "　　<" & num & "> " & n & vbNewLine
                            num += 1
                        Next
                    End If
                Next
            End If
        Next

        Dim sfd As New SaveFileDialog With {
                .InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                .Filter = "TXTファイル(*.txt)|*.txt|すべてのファイル(*.*)|*.*",
                .AutoUpgradeEnabled = False,
                .FilterIndex = 0,
                .Title = "保存先のファイルを選択してください",
                .RestoreDirectory = True,
                .OverwritePrompt = True,
                .CheckPathExists = True,
                .FileName = "DOM.txt"
            }

        'ダイアログを表示する
        If sfd.ShowDialog(Me) = DialogResult.OK Then
            File.WriteAllText(sfd.FileName, str)
        End If
    End Sub
End Class