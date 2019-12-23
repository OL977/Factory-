Public Class ПереводСрокКонтр
    Private Sub ПереводСрокКонтр_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim fd As New List(Of String)() From {"1", "2", "3", "4", "5"}
        ListBox1.Items.Clear()
        For i As Integer = 0 To fd.Count - 1
            ListBox1.Items.Add(fd(i))
        Next
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub

    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick
        If Not ListBox1.SelectedIndex = -1 Then
            Перевод.ComboBox8.Text = ListBox1.SelectedItem.ToString
            Me.Close()
        End If
    End Sub
End Class