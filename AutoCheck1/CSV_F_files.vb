Public Class CSV_F_files
    Public Shared dataPathArray As New ArrayList

    '読み込み
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If DGV1.RowCount > 0 Then
            Dim csvNum As Integer = 0
            If Form1.CSV_FORMS(0).ToolStripStatusLabel5.Text = 2 Then
                csvNum = 1
            Else
                csvNum = 0
            End If

            Dim headerFlag As Boolean = True
            If CheckBox1.Checked Then
                headerFlag = False
            End If

            Dim addFlag As Boolean = False
            If CheckBox2.Checked Then
                addFlag = True
            End If

            Form1.CSV_FORMS(csvNum).dataPathArray = dataPathArray
            Form1.CSV_FORMS(csvNum).headerAdd = headerFlag
            Form1.CSV_FORMS(csvNum).addFlag = addFlag
            'Form1.CSV_FORMS(csvNum).DropRun(dataPathArray, dataPathArray.Count - 1, headerFlag, addFlag)
        End If

        Me.Close()
    End Sub

    '選択行削除
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim delRowArray As New ArrayList
        For i As Integer = 0 To DGV1.SelectedCells.Count - 1
            If Not delRowArray.Contains(DGV1.SelectedCells(i).RowIndex) Then
                delRowArray.Add(DGV1.SelectedCells(i).RowIndex)
            End If
        Next
        delRowArray.Sort()
        For i As Integer = 0 To delRowArray.Count - 1
            dataPathArray.RemoveAt(delRowArray(i))
            DGV1.Rows.RemoveAt(delRowArray(i))
        Next
    End Sub
End Class