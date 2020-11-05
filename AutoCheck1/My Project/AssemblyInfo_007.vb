
                            End If
                        End If
                    End If
                Next
            End If
        Next

        HaisoSizeAutoKeisan(1)
        HukuSuGyoCopy(1, gArray) '複数個口の場合行を複製する
        NouhinSyoList() '納品書を出力する必要があるリスト
    End Sub

    'リストから自動計算
    Private Sub HaisoSizeAutoKeisan(ByVal mode As Integer)
        SplitContainer7.SplitterDistance = SplitContainer7.Width - 290
        TabControl3.SelectTab("TabPage10")
        Panel1.Visible = True
        Application.DoEvents()

        DataGridView11.Rows.Clear()
        'Dim selRowNum As Integer = DataGridView1.SelectedCells(0).RowIndex

        '-----------
        'datagridviewのヘッダーテキストをコレクションに取り込む
        Dim DGVheaderCheck As ArrayList = New ArrayList
        For c As Integer = 0 To DataGridView1.Columns.Count - 1
            DGVheaderCheck.Add(DataGridView1.Item(c, 0).Value)
        Next c
        Dim DGVheaderCheck2 As ArrayList = New ArrayList
        For c As Integer = 0 To DataGridView6.Columns.Count - 1
            DGVheaderCheck2.Add(DataGridView6.Item(c, 0).Value)
        Next c
        '-----------

        'mode：0=DataGridView1全行、1=選択行のみ
        Dim rowStart As Integer = 0
        Dim rowMax As Integer = 0
        Dim selR