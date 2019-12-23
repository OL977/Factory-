Option Explicit On
Imports System.ComponentModel
Imports System.Data.OleDb
Public Class КонтрДопЛемел
    Dim d As Integer
    Dim Flag As Boolean = True
    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            TextBox1.Focus()
        End If
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            d = 0
            If IsNumeric(TextBox1.Text) = True Then
                TextBox2.Text = CType(CType(TextBox1.Text, Integer) - 1, String)
            Else
                MessageBox.Show("В поле можно ввести только целые числа!", Рик)
                d = 1
                Exit Sub
            End If
            Button1.Focus()
        End If
    End Sub

    Private Sub КонтрДопЛемел_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim f() As String = {"без предварительного испытания", "с предварительным испытанием 1 месяц", " предварительным испытанием 2 месяца", " предварительным испытанием 3 месяца"}

        ComboBox1.Items.Clear()
        ComboBox1.Items.AddRange(f)

        If Прием.CheckBox5.Checked = True Then
            ComboBox1.Text = vstavContr
            TextBox1.Text = vstavContr1
        Else
            ComboBox1.Text = f.First
            TextBox1.Text = "25"
            TextBox2.Text = "24"
        End If


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If d = 1 Then
            MessageBox.Show("Заполните правильно поле 'Трудовой отпуск (дней)'!", Рик)
            Exit Sub
        End If

        If ComboBox1.Text = "" Then
            MessageBox.Show("Выберите условия испытательного срока!", Рик)
            Exit Sub
        End If
        ЛемелТрОтп = TextBox1.Text
        ЛемелИспытСрок = ComboBox1.Text
        TextBox1.Text = ""
        ComboBox1.Text = ""
        Flag = False
        Me.Close()
    End Sub

    Private Sub КонтрДопЛемел_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing


        If Flag = False Then
            e.Cancel = Flag
        Else
            e.Cancel = Flag
            MessageBox.Show("Обязательно принять решение по испытатальному периоду и отпуску!", Рик)
            'Me.ShowDialog()

        End If
        'Me.WindowState = FormWindowState.Minimized
    End Sub
End Class