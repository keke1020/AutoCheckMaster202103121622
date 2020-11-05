<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PrintData
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows フォーム デザイナで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
    'Windows フォーム デザイナを使用して変更できます。  
    'コード エディタを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.ReoGridControl1 = New unvell.ReoGrid.ReoGridControl()
        Me.SuspendLayout()
        '
        'ReoGridControl1
        '
        Me.ReoGridControl1.BackColor = System.Drawing.Color.White
        Me.ReoGridControl1.ColumnHeaderContextMenuStrip = Nothing
        Me.ReoGridControl1.LeadHeaderContextMenuStrip = Nothing
        Me.ReoGridControl1.Location = New System.Drawing.Point(37, 59)
        Me.ReoGridControl1.Name = "ReoGridControl1"
        Me.ReoGridControl1.RowHeaderContextMenuStrip = Nothing
        Me.ReoGridControl1.Script = Nothing
        Me.ReoGridControl1.SheetTabContextMenuStrip = Nothing
        Me.ReoGridControl1.SheetTabNewButtonVisible = True
        Me.ReoGridControl1.SheetTabVisible = True
        Me.ReoGridControl1.SheetTabWidth = 400
        Me.ReoGridControl1.ShowScrollEndSpacing = True
        Me.ReoGridControl1.Size = New System.Drawing.Size(509, 324)
        Me.ReoGridControl1.TabIndex = 0
        Me.ReoGridControl1.Text = "ReoGridControl1"
        '
        'PrintData
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(807, 547)
        Me.Controls.Add(Me.ReoGridControl1)
        Me.Name = "PrintData"
        Me.Text = "帳票作成"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents ReoGridControl1 As unvell.ReoGrid.ReoGridControl
End Class
