Option Explicit On
Imports System.Data.OleDb

Public Class СправкаЗпКонстанта
    Dim strsql As String
    Dim ds, ds1 As DataTable
    Private Sub СправкаЗпКонстанта_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Me.ComboBox1.Items.Clear()
        Dim ut() As Object = {Now.Year - 3, Now.Year - 2, Now.Year - 1, Now.Year}
        ComboBox1.Items.AddRange(ut)
        ComboBox1.Text = Now.Year
        strsql = "SELECT * FROM КонстантаПоВычетамЗп WHERE Год=" & ComboBox1.SelectedItem & ""
        ds = Selects(strsql)
        Try
            TextBox1.Text = ds.Rows(0).Item(1).ToString
            TextBox2.Text = ds.Rows(0).Item(2).ToString
            TextBox3.Text = ds.Rows(0).Item(3).ToString
            TextBox4.Text = ds.Rows(0).Item(4).ToString
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'If MessageBox.Show("Заменить старые данные?", Рик, MessageBoxButtons.YesNo) = DialogResult.No Then
        '    Exit Sub
        'End If

        errds = 0
        Dim strsql3 As String = "SELECT МаксДоход FROM КонстантаПоВычетамЗп WHERE Год=" & CType(ComboBox1.Text, Integer) & ""
        Dim dsx As DataTable = Selects(strsql3)

        If errds = 0 Then

            Dim strsql2 As String = "UPDATE КонстантаПоВычетамЗп SET МаксДоход='" & TextBox1.Text & "',СуммаВычета= '" & TextBox2.Text & "',ВычетНаОднРеб= '" & TextBox3.Text & "',
ВычетНаДваИБолДет='" & TextBox4.Text & "'
WHERE Год=" & CType(ComboBox1.Text, Integer) & ""
            Updates(strsql2)
        Else
            strsql = "INSERT INTO КонстантаПоВычетамЗп(МаксДоход, СуммаВычета, ВычетНаОднРеб, ВычетНаДваИБолДет,Год)
VALUES('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "'," & ComboBox1.Text & ")"
            Updates(strsql)

        End If
        MessageBox.Show("Данные изменены!", Рик)
        Me.Close()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        strsql = "SELECT * FROM КонстантаПоВычетамЗп WHERE Год=" & CType(ComboBox1.SelectedItem, Integer) & ""
        ds = Selects(strsql)
        Try
            TextBox1.Text = ds.Rows(0).Item(1).ToString
            TextBox2.Text = ds.Rows(0).Item(2).ToString
            TextBox3.Text = ds.Rows(0).Item(3).ToString
            TextBox4.Text = ds.Rows(0).Item(4).ToString
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        Me.Close()
    End Sub

    Private Sub СправкаЗпКонстанта_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
    End Sub
End Class