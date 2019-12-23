Public Class ПереводНовОтдел
    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub

    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick
        Перевод.ComboBox3.Text = ListBox1.SelectedItem.ToString
        Me.Close()
    End Sub
End Class