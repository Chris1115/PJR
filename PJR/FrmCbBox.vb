Public Class FrmCbBox
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        FormMain.cbHariBukaToko = ComboBox1.Text

        Me.Close()
    End Sub

    Private Sub FrmCbBox_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.SelectedIndex = -1
        If FormMain.cbHariBukaToko = "" Then
            Button1.Enabled = False
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged


        If ComboBox1.Text <> "" Then
            Button1.Enabled = True

        End If
    End Sub
End Class