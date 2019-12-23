Public Class ПарольНаСтатистику
    Dim login As String
    Dim Password As String
    Dim l As Integer

    Private Sub ПарольНаСтатистику_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Me.MdiParent = MDIParent1
        login = "vika@a"
        Password = "1389925vika"

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        l += 1
        Статистика("Ввод логина - " & TextBox1.Text, "Ввод пароля - " & TextBox2.Text, "Попытка входа № " & l)

        If TextBox1.Text = login Then
            If TextBox2.Text = Password Then
                Me.Close()
                Статистикаc.ShowDialog()
            Else
                MessageBox.Show("Не правильно введён пароль!", Рик)
            End If
        Else
            MessageBox.Show("Не правильно введён логин!", Рик)

        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class