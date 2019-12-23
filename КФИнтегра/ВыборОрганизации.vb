Option Explicit On
Imports System.Data.OleDb
Public Class ВыборОрганизации
    Private Sub ВыборОрганизации_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MdiParent = MDIParent1
        'conn = New OleDbConnection
        'conn.ConnectionString = ConString
        'Try
        '    conn.Open()
        'Catch ex As Exception
        '    MessageBox.Show("Не подключен диск U")
        'End Try


        Me.ComboBox1.Items.Clear()
        For Each r As DataRow In СписокКлиентовОсновной.Rows
            Me.ComboBox1.Items.Add(r(0).ToString)
        Next



    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Прием.ComboBox1.Text = Me.ComboBox1.Text
        Прием.CheckBox7.Checked = True
        Me.Close()
        Прием.Show()
    End Sub
End Class