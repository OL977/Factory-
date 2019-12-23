Public Class ПереводВремНач
    Private Sub ПереводВремНач_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim fd As New List(Of String)() From {"8.30", "9.00", "10.00", "10.30", "11.00", "12.00"}
        ListBox1.Items.Clear()
        For i As Integer = 0 To fd.Count - 1
            ListBox1.Items.Add(fd(i))
        Next
    End Sub

    Private Sub ПереводВремНач_DoubleClick(sender As Object, e As EventArgs) Handles MyBase.DoubleClick

    End Sub

    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick
        If Not ListBox1.SelectedIndex = -1 Then
            Перевод.ComboBox10.Text = ListBox1.SelectedItem.ToString
            Me.Close()
        End If
    End Sub
End Class