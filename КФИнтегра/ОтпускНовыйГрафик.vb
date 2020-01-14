Public Class ОтпускНовыйГрафик
    Dim StrSql As String
    Dim dt As DataTable


    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            CheckBox2.Checked = False

        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            CheckBox1.Checked = False

        End If
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            MaskedTextBox1.Focus()
        End If
    End Sub

    Private Sub MaskedTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            DomainUpDown1.Focus()
        End If
    End Sub

    Private Sub DomainUpDown1_KeyDown(sender As Object, e As KeyEventArgs) Handles DomainUpDown1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            ComboBox1.Focus()
        End If
    End Sub

    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            MaskedTextBox2.Focus()
        End If
    End Sub

    Private Sub MaskedTextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Button1.Focus()
        End If
    End Sub



    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ЗапГрид1()
    End Sub
    Private Sub ЗапГрид1()

        Dim нагод As String
        If CheckBox1.Checked = True Then
            нагод = DomainUpDown1.Text
        Else
            нагод = ComboBox1.Text
        End If

        StrSql = "INSERT INTO Отпуск(Орг,НаГод,Номер,Составлен,Утвержден)  VALUES('" & Отпуск1.ComboBox2.Text & "',
'" & нагод & "','" & TextBox1.Text & "','" & MaskedTextBox1.Text & "','" & MaskedTextBox2.Text & "')"
        Updates(StrSql)

        StrSql = ""

        StrSql = "SELECT НаГод, Номер, Составлен, Утвержден FROM Отпуск WHERE Орг='" & Отпуск1.ComboBox2.Text & "'"
        dt = Selects(StrSql)
        Отпуск1.Grid1.DataSource = dt
        Статистика1("Нет", "Создание нового графика отпусков", Отпуск1.ComboBox2.Text)
        Me.Close()
    End Sub

    Private Sub ОтпускНовыйГрафик_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckBox1.Checked = True
    End Sub
End Class