Imports System.Data.OleDb
Public Class ИзменБанк

    Dim Код As String
    Private Sub ИзменБанк_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Me.MdiParent = MDIParent1


        'conn = New OleDbConnection
        'conn.ConnectionString = ConString
        'Try
        '    conn.Open()
        'Catch ex As Exception
        '    MessageBox.Show("Не подключен диск U")
        'End Try

        Dim StrSql8 As String = "SELECT DISTINCT КорНазБанк FROM БанкКор ORDER BY КорНазБанк"
        Dim ds As DataTable = Selects(StrSql8)

        Me.ComboBox1.Items.Clear()
        For Each r As DataRow In ds.Rows
            Me.ComboBox1.Items.Add(r(0).ToString)
        Next

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        Dim StrSql8 As String = "SELECT Наименование FROM Банк WHERE Наименование LIKE '%" & ComboBox1.Text & "%' ORDER BY Наименование "
        Dim ds As DataTable = Selects(StrSql8)
        Me.ComboBox2.Items.Clear()
        For Each r As DataRow In ds.Rows
            Me.ComboBox2.Items.Add(r(0).ToString)
        Next
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim ds As DataTable = Selects(StrSql:="Select БИК,КодГлав From Банк Where Наименование='" & ComboBox2.Text & "'")

        Try
            TextBox18.Text = ds.Rows(0).Item(0).ToString
        Catch ex As Exception
            MessageBox.Show("В базе нет данных по БИК(у) отделения", Рик, MessageBoxButtons.OK)
        End Try
        Код = ds.Rows(0).Item(1).ToString
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim StrSql As String
        StrSql = "UPDATE Банк SET Наименование='" & ComboBox2.Text & "', БИК='" & TextBox18.Text & "' WHERE КодГлав=" & CType(Код, Integer) & ""
        Updates(StrSql)


        MessageBox.Show("Изменения внесены в базу!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Information)
        'thisForm.Close()
        ComboBox2.Text = ""
        TextBox18.Text = ""
        ComboBox1.Text = ""

        Me.Close()
        Контрагент.refr()

    End Sub
End Class