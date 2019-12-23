Public Class ПереводОрганиз
    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick
        Перевод.ComboBox1.Text = ListBox1.SelectedItem.ToString
        Me.Close()
    End Sub

    Private Sub ПереводОрганиз_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub
End Class