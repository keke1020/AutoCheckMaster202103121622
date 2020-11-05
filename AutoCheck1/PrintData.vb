Option Explicit On

Imports unvell.ReoGrid


Public Class PrintData
    Private Sub PrintData_Load(sender As Object, e As EventArgs) Handles Me.Load
        ReoGridControl1.Cursor = Cursors.Arrow
        ReoGridControl1.CellsSelectionCursor = Cursors.Arrow



    End Sub


End Class