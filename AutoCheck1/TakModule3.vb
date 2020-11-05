Module TakModule3

    ''' <summary>
    ''' DataGridViewチラつき防止
    ''' </summary>
    ''' <param name="dgvArray">DataGridViewArray</param>
    Public Sub TM_DGV_CHIRATUKI(dgvArray As DataGridView())
        '******************************************
        'dgv画面チラつき防止
        '******************************************
        'DataGirdViewのTypeを取得
        Dim dgvtype As Type = GetType(DataGridView)
        'プロパティ設定の取得
        Dim dgvPropertyInfo As Reflection.PropertyInfo = dgvtype.GetProperty("DoubleBuffered", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)

        '対象のDataGridViewにtrueをセットする
        For i As Integer = 0 To dgvArray.Length - 1
            dgvPropertyInfo.SetValue(dgvArray(i), True, Nothing)
        Next
    End Sub

    ''' <summary>
    ''' DataGridView列固定
    ''' </summary>
    ''' <param name="dgv">DataGridView</param>
    Private Sub TM_DGV_KOTEI(dgv As DataGridView)
        Dim selCell = dgv.SelectedCells

        For c As Integer = 0 To dgv.ColumnCount - 1
            dgv.Columns(c).Frozen = False
            dgv.Columns(c).DividerWidth = 0
        Next

        Dim koteiR As New ArrayList
        For i As Integer = 0 To selCell.Count - 1
            If Not koteiR.Contains(selCell(i).ColumnIndex) Then
                koteiR.Add(selCell(i).ColumnIndex)
            End If
        Next
        koteiR.Sort()
        Dim fdsc As Integer = dgv.FirstDisplayedScrollingColumnIndex
        If koteiR(0) > 0 Then
            For i As Integer = 0 To dgv.Columns.Count - 1
                If i < fdsc Then
                    dgv.Columns(i).Visible = False
                Else
                    If koteiR(0) >= fdsc Then
                        dgv.Columns(koteiR(0) - 1).Frozen = True
                        dgv.Columns(koteiR(0) - 1).DividerWidth = 3
                        Exit For
                    End If
                End If
            Next
        End If
    End Sub

    ''' <summary>
    ''' DataGridView列固定解除
    ''' </summary>
    ''' <param name="dgv">DataGridView</param>
    Private Sub TM_DGV_KOTEIDEL(dgv As DataGridView)
        For c As Integer = 0 To dgv.Columns.Count - 1
            dgv.Columns(c).Visible = True
        Next
        For c As Integer = 0 To dgv.ColumnCount - 1
            If dgv.Columns(c).Frozen = True Then
                dgv.Columns(c).Frozen = False
            Else
                dgv.Columns(c).DividerWidth = 0
            End If
        Next
    End Sub

End Module
