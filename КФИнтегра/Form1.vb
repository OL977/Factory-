Imports System.Data.OleDb

Public Class Form1
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        refreshgrid()
    End Sub
    Private Sub refreshgrid()
        Dim StrSql As String
        StrSql = "SELECT * FROM Сотрудники"
        Dim c As New OleDbCommand With {
            .Connection = conn,
            .CommandText = StrSql
        }
        Dim ds As New DataSet
        Dim da As New OleDbDataAdapter(c)
        da.Fill(ds, "Сотрудники")
        Grid1.DataSource = ds
        Grid1.DataMember = "Сотрудники"

    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim dBaseCommand As New OleDbCommand
        Dim dBaseConnection As New OleDbConnection(ConString)
        Dim adapter As OleDbDataAdapter
        dBaseCommand.Connection = dBaseConnection
        dBaseCommand.CommandType = CommandType.Text
        adapter = New OleDbDataAdapter(dBaseCommand)
        Dim dt_tt As New DataTable
        dBaseCommand.CommandText = "select ФИОСборное,КодСотрудники from Сотрудники" 'имя столбца и таблицы в запросе ваши... 
        'ну надеюсь понятно написал... тут выберутся все значения... но можно использовать и более сложный запрос...
        adapter.Fill(dt_tt) 'заполняем созданную таблицу данными из запроса
        'можно и тут сделать чего-нибудь с табличкой если надо...сортировать, добавить/удалить что нужно, сделать выборку...
        Me.ComboBox1.Items.Clear() 'очистили комбобокс от предыдущих значений

        For Each r As DataRow In dt_tt.Rows
            Me.ComboBox1.Items.Add(r(0).ToString)
            'Me.ComboBox3.Items.Add(r(1).ToString) 'заполняем комбобокс значениями единственного столбца нашей таблички
        Next
    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        conn = New OleDbConnection With {
            .ConnectionString = ConString
        }
        conn.Open()

        Dim dBaseCommand As New OleDbCommand
        Dim dBaseConnection As New OleDbConnection(ConString)
        Dim adapter As OleDbDataAdapter
        dBaseCommand.Connection = dBaseConnection
        dBaseCommand.CommandType = CommandType.Text
        adapter = New OleDbDataAdapter(dBaseCommand)
        Dim dt_tt As New DataTable
        dBaseCommand.CommandText = "select НазвОрг from Клиент" 'имя столбца и таблицы в запросе ваши... 
        'ну надеюсь понятно написал... тут выберутся все значения... но можно использовать и более сложный запрос...
        adapter.Fill(dt_tt) 'заполняем созданную таблицу данными из запроса
        'можно и тут сделать чего-нибудь с табличкой если надо...сортировать, добавить/удалить что нужно, сделать выборку...
        Me.ComboBox2.Items.Clear() 'очистили комбобокс от предыдущих значений

        For Each r As DataRow In dt_tt.Rows
            Me.ComboBox2.Items.Add(r(0).ToString) 'заполняем комбобокс значениями единственного столбца нашей таблички
        Next
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim StrSql, a, b, d, we As String
        a = TextBox1.Text
        b = TextBox2.Text
        d = TextBox3.Text
        we = ComboBox2.Text

        StrSql = "INSERT INTO Сотрудники(НазвОрганиз,Фамилия,Имя,Отчество) VALUES ('" & we & "','" & a & "','" & b & "','" & d & "')"
        'Dim c As New OleDbCommand With {
        '    .Connection = conn,
        '    .CommandText = StrSql
        '}
        'c.ExecuteNonQuery()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click 'insert в базу
        Dim StrSql, a, b, d, we As String
        a = TextBox6.Text
        b = TextBox5.Text
        d = TextBox4.Text
        we = ComboBox1.Text

        StrSql = "INSERT INTO ДогСотрудн(Сотрудник,Контракт,ДатаКонтракта,Приказ) VALUES ('" & we & "','" & a & "','" & b & "','" & d & "')"
        Dim c As New OleDbCommand With {
            .Connection = conn,
            .CommandText = StrSql
        }
        c.ExecuteNonQuery()
    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged 'выбираем и вставляем в текстбокс
        Dim StrSql, a As String
        a = ComboBox1.Text
        StrSql = "SELECT Сотрудники.КодСотрудники, Сотрудники.ФИОСборное, ДогСотрудн.Контракт, ДогСотрудн.ДатаКонтракта FROM Сотрудники INNER JOIN ДогСотрудн ON Сотрудники.КодСотрудники = ДогСотрудн.IDСотр WHERE Сотрудники.ФИОСборное='" & a & "'"
        Dim c As New OleDbCommand
        c.Connection = conn
        c.CommandText = StrSql
        Dim ds As New DataSet
        Dim da As New OleDbDataAdapter(c)
        da.Fill(ds, "Сотрудники")
        'Grid1.DataSource = ds


        Try
            TextBox9.Text = ds.Tables("Сотрудники").Rows(0).Item(0)
            TextBox7.Text = ds.Tables("Сотрудники").Rows(0).Item(2)
            TextBox8.Text = ds.Tables("Сотрудники").Rows(0).Item(3)
        Catch ex As Exception
            MessageBox.Show("Нет данных в базе!!!")
        End Try

    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click 'редактируем
        Dim StrSql, a, b, d As String
        Dim f As Integer
        a = ComboBox1.Text
        b = TextBox7.Text
        d = TextBox8.Text
        f = TextBox9.Text
        StrSql = "UPDATE ДогСотрудн SET Контракт ='" & b & "', ДатаКонтракта ='" & d & "' WHERE " & f & " = ДогСотрудн.IDСотр"
        Dim c As New OleDbCommand
        c.Connection = conn
        c.CommandText = StrSql
        c.ExecuteNonQuery()
        'Dim ds As New DataSet
        'Dim da As New OleDbDataAdapter(c)
        'da.Fill(ds, "Сотрудники")
        ''Grid1.DataSource = ds


        'Try
        '    TextBox7.Text = ds.Tables("Сотрудники").Rows(0).Item(0)
        '    TextBox8.Text = ds.Tables("Сотрудники").Rows(0).Item(1)
        'Catch ex As Exception
        '    MessageBox.Show("Нет данных в базе!!!")
        'End Try
    End Sub

    Private Sub Grid1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellContentClick

    End Sub
End Class
