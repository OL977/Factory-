Public Class НовОргДопФорма
    Dim МРаботы As String
    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            CheckBox1.Checked = False
            CheckBox2.Checked = False
            TextBox1.Visible = True
        Else
            TextBox1.Visible = False
            TextBox1.Text = ""
        End If

    End Sub

    Private Sub НовОргДопФорма_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Visible = False
        CheckBox1.Checked = True
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            CheckBox2.Checked = False
            CheckBox3.Checked = False

        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            CheckBox1.Checked = False
            CheckBox3.Checked = False

        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim fg As String
        If CheckBox1.Checked = True Then
            fg = "Юридический адрес"
        End If

        If CheckBox2.Checked = True Then
            fg = "Фактический адрес"
        End If

        If CheckBox3.Checked = True Then
            fg = TextBox1.Text
        End If

        If MessageBox.Show("Выбран " & fg & "! Вы подтверждаете выбор?", Рик, MessageBoxButtons.YesNo) = DialogResult.No Then Exit Sub
        Me.Close()
            Контрагент.CheckBox7.Checked = True
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        If MessageBox.Show("Сохранить данные?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then Exit Sub

        If CheckBox1.Checked = True Then
            МРаботы = Контрагент.TextBox7.Text
        End If

        If CheckBox2.Checked = True Then
            МРаботы = Контрагент.TextBox12.Text
        End If

        If CheckBox3.Checked = True Then
            МРаботы = TextBox1.Text
        End If

        Dim strsql As String
            strsql = "INSERT INTO ОбъектОбщепита(НазвОрг,АдресОбъекта) VALUES('" & TextBox1.Text & "', '" & МРаботы & "')"
            Updates(strsql)
        MessageBox.Show("Данные внесены!", Рик)
        Me.Close()
        Контрагент.CheckBox7.Checked = True
    End Sub
End Class