Option Explicit On
Imports System.Data.OleDb
Public Class Примечание

    Private Sub Примечание_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Прим = 1 Then
            If Прием.TextBox1.Text = "" Then
                TextBox2.Text = ""
            Else
                TextBox2.Text = Trim(Прием.TextBox1.Text) & " " & Trim(Прием.TextBox2.Text) & " " & Trim(Прием.TextBox3.Text)
            End If

        Else

        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ДоговорПодрядаСписки.clb2 = True
        TextBox2.Text = ""
        Me.Close()
    End Sub
    Private Sub savep()
        Dim list As New Dictionary(Of String, Object)
        list.Add("@IDСотр", CType(Label2.Text, Integer))
        Updates(stroka:="UPDATE КарточкаСотрудника SET Примечание='" & RichTextBox1.Text & "' WHERE IDСотр=@IDСотр", list, "КарточкаСотрудника")

    End Sub
    Private Sub ДогПодр()
        Dim list As New Dictionary(Of String, Object)
        list.Add("@IDСотр", CType(Label2.Text, Integer))
        Updates(stroka:="UPDATE ДогПодряда SET Примечание='" & RichTextBox1.Text & "' WHERE ID=@IDСотр", list, "ДогПодряда")

    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox2.Text = "" Then
            MessageBox.Show("Нет данных для сохранения!", Рик)
            Exit Sub
        End If

        If Прим = 1 Then
            Прием.Примечани = ""
            Прием.Примечани = RichTextBox1.Text
            MessageBox.Show("Данные приняты!", Рик)
            Me.Close()
        ElseIf Прим = 3 Then
            ДогПодр()
            MessageBox.Show("Данные приняты!", Рик)
            Me.Close()
        Else
            savep()
            MessageBox.Show("Данные внесены!", Рик)
            'НеподпДокументы.ПоискПоСотр()
            Me.Close()
        End If
        Статистика(TextBox2.Text, "Примечение при приеме", "Примечение - " & RichTextBox1.Text)
    End Sub
End Class