Public Class ПереводРазр
    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick
        Перевод.ComboBox7.Text = ListBox1.SelectedItem.ToString
        Me.Close()
    End Sub

    Private Sub ПереводРазр_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim fd As New List(Of String)() From {"-", "1", "2", "3", "4", "5", "6"}
        ListBox1.Items.Clear()
        For i As Integer = 0 To fd.Count - 1
            ListBox1.Items.Add(fd(i))
        Next
    End Sub
End Class