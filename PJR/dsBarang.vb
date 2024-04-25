

Partial Public Class dsBarang
    Partial Class dtPeriodeReturDataTable

        Private Sub dtPeriodeReturDataTable_ColumnChanging(ByVal sender As System.Object, ByVal e As System.Data.DataColumnChangeEventArgs) Handles Me.ColumnChanging
            If (e.Column.ColumnName = Me.PLUColumn.ColumnName) Then
                'Add user code here
            End If

        End Sub

    End Class

End Class
