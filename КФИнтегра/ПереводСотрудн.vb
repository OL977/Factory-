Public Class ПереводСотрудн

    Private Sub ПереводСотрудн_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub

    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick
        Перевод.ComboBox2.Text = ListBox1.SelectedItem.ToString
        Me.Close()
    End Sub
End Class