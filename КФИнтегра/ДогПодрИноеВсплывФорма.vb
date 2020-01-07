Public Class ДогПодрИноеВсплывФорма
    Public Property ЕдИзм() As String
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        If TextBox1.Text <> "" Then
            ЕдИзм = TextBox1.Text
        Else
            MessageBox.Show("Заполните поле 'Единица измерения!'", Рик)
            Exit Sub
        End If
        Me.Close()
    End Sub
End Class