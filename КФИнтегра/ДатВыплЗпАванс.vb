Public Class ДатВыплЗпАванс
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Соед(0)

        If MessageBox.Show("Внести изменения?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.Cancel Then
            Exit Sub
        End If
        Dim ir, ir2 As Integer

        If IsNumeric(TextBox1.Text) = False And TextBox1.Text <> "" Then
            MessageBox.Show("Внесите числовое значение или оставьте поле пустым!", Рик)
            Exit Sub
        End If
        If IsNumeric(TextBox2.Text) = False And TextBox2.Text <> "" Then
            MessageBox.Show("Внесите числовое значение или оставьте поле пустым!", Рик)
            Exit Sub
        End If

        If Not TextBox1.Text = "" Then
            If CType(TextBox1.Text, Integer) > 31 Then
                MessageBox.Show("Внесите корректные данные!", Рик)
                Exit Sub
            End If
        End If
        If Not TextBox2.Text = "" Then
            If CType(TextBox2.Text, Integer) > 31 Then
                MessageBox.Show("Внесите корректные данные!", Рик)
                Exit Sub
            End If
        End If

        'Dim gf As String = ""
        'If TextBox1.Text <> "" And Not TextBox2.Text <> "" Then
        '    Try
        '        ir = CType(TextBox1.Text, Integer)

        '    Catch ex As Exception
        '        MessageBox.Show("Внесите корректные данные!", Рик)
        '        Exit Sub
        '    End Try

        '    If ir > 31 Then
        '        MessageBox.Show("Внесите корректные данные!", Рик)
        '        Exit Sub
        '    End If

        Dim list As New Dictionary(Of String, Object)
        list.Add("@Прием", Прием.ComboBox1.Text)



        Updates(stroka:="UPDATE КарточкаСотрудника 
SET ДатаЗарплаты='" & TextBox1.Text & "', ДатаАванса='" & TextBox2.Text & "'
from  Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр
WHERE Сотрудники.НазвОрганиз =@Прием", list)


        '        ElseIf Not TextBox1.Text <> "" And TextBox2.Text <> "" Then

        '            Try
        '                ir2 = CType(TextBox2.Text, Integer)
        '            Catch ex As Exception
        '                MessageBox.Show("Внесите корректные данные!", Рик)
        '                Exit Sub
        '            End Try

        '            If ir2 > 31 Then
        '                MessageBox.Show("Внесите корректные данные!", Рик)
        '                Exit Sub
        '            End If

        '            Dim sqlstr As String = "UPDATE КарточкаСотрудника 
        'INNER JOIN Сотрудники ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр
        'SET ДатаАванса='" & TextBox2.Text & "',ДатаЗарплаты='" & gf & "'
        'WHERE Сотрудники.НазвОрганиз ='" & Прием.ComboBox1.Text & "'"
        '            Updates(sqlstr)

        '        Else

        '            Try
        '                ir = CType(TextBox1.Text, Integer)
        '                ir2 = CType(TextBox2.Text, Integer)
        '            Catch ex As Exception
        '                MessageBox.Show("Внесите корректные данные!", Рик)
        '                Exit Sub
        '            End Try

        '            If ir > 31 Or ir2 > 31 Then
        '                MessageBox.Show("Внесите корректные данные!", Рик)
        '                Exit Sub
        '            End If

        '            Dim sqlstr As String = "UPDATE КарточкаСотрудника
        'INNER JOIN Сотрудники ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр
        'SET ДатаЗарплаты='" & TextBox1.Text & "', ДатаАванса='" & TextBox2.Text & "'
        'WHERE Сотрудники.НазвОрганиз ='" & Прием.ComboBox1.Text & "'"
        '            Updates(sqlstr)
        '        End If

        MessageBox.Show("Данные изменены!", Рик)
        Статистика("Нет", "Изменение даты выплаты аванса или зарплаты", Прием.ComboBox1.Text)
        Me.Close()
        Прием.Com1sel()
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox2.Focus()
        End If
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.Button1.Focus()
        End If
    End Sub


End Class