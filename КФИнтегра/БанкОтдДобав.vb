Imports System.Data.OleDb
Public Class БанкОтдДобав
    Private Sub БанкОтдДобав_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim strsql As String = "SELECT КорНазБанк FROM БанкКор ORDER BY КорНазБанк"
        Dim ds As DataTable = Selects(strsql)
        ComboBox1.Items.Clear()
        For Each r As DataRow In ds.Rows
            ComboBox1.Items.Add(r(0).ToString)
        Next

    End Sub
    Private Function пров()
        If ComboBox1.Text = "" Then
            MessageBox.Show("Выберите банк!", Рик)
            Return 1
        End If
        If RichTextBox1.Text = "" Then
            MessageBox.Show("Заполните раздел отделения банка!", Рик)
            Return 1
        End If
        If RichTextBox2.Text = "" Then
            MessageBox.Show("Заполните раздел БИК банка!", Рик)
            Return 1
        End If
        Return 0


    End Function
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If пров() = 1 Then Exit Sub
        Dim strsql As String = "INSERT INTO Банк(Наименование,БИК) VALUES('" & RichTextBox1.Text & " (" & ComboBox1.Text & ")" & "','" & RichTextBox2.Text & "')"
        Updates(strsql)

        MessageBox.Show("Данные вневены!", Рик)
        Me.Close()
        Контрагент.refr()
    End Sub
End Class