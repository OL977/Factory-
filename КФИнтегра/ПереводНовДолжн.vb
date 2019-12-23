Public Class ПереводНовДолж
    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick
        Перевод.ComboBox4.Text = ListBox1.SelectedItem.ToString
        Me.Close()
    End Sub
End Class