Imports System.Data.OleDb
Public Class Банк

    Private Sub Банк_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Me.MdiParent = MDIParent1
        'Соед(0)
    End Sub

    Private Sub Банк_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        'Контрагент.CheckBox3.Checked = False
        'Соед(1)
    End Sub
    Private Function Проверка()
        If RichTextBox1.Text = "" Then
            MessageBox.Show("Заполните короткое название банка", Рик)
            Return 1
        End If
        If RichTextBox2.Text = "" Then
            MessageBox.Show("Заполните полное название банка", Рик)
            Return 1
        End If
        If RichTextBox3.Text = "" Then
            MessageBox.Show("Заполните БИК банка", Рик)
            Return 1
        End If

        Return 0
    End Function
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Проверка() = 1 Then Exit Sub

        If MessageBox.Show("Создать новый банк?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = DialogResult.Cancel Then Exit Sub


        Dim StrSql As String = "INSERT INTO БанкКор(КорНазБанк) VALUES ('" & RichTextBox1.Text & "')"
        Updates(StrSql)


        Dim StrSql1 As String = "INSERT INTO Банк(Наименование,БИК) VALUES ('" & RichTextBox2.Text & "','" & RichTextBox3.Text & "')"
        Updates(StrSql1)

        MessageBox.Show("Данные сохранены", Рик)

        Dim StrSql8 As String = "SELECT КорНазБанк FROM БанкКор ORDER BY КорНазБанк " 'WHERE Наименование LIKE '%" & Контрагент.ComboBox3.Text & "%'
        Dim ds8 As DataTable
        ds8 = Selects(StrSql8)
        Контрагент.ComboBox3.Items.Clear()
        For Each r As DataRow In ds8.Rows
            Контрагент.ComboBox3.Items.Add(r(0).ToString)
        Next

        Статистика("Нет", "Добавление банка", "Нет")
        RichTextBox1.Text = ""
        RichTextBox2.Text = ""
        RichTextBox3.Text = ""
        Me.Close()
    End Sub



End Class