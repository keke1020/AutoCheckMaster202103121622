Public Class Csv_denpyo3_F_files
    Public Shared dataPathArray As New ArrayList

    Private Sub Csv_denpyo3_F_files_Load(sender As Object, e As EventArgs) Handles Me.Load

    End Sub

    '読み込み
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'If DGV1.RowCount > 0 Then
        '    Csv_denpyo3.FileDataRead(dataPathArray)
        'End If

        'Csv_denpyo3.FileDataRead(dataPathArray)
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