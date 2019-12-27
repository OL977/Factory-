
Option Explicit On
Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.IO
Public Class ОтменУвол

    Public отмувсм As Integer

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ОтменаУв()

    End Sub
    Private Sub ОтменаУв()
        Me.Cursor = Cursors.WaitCursor
        Dim strsql, fd As String
        Dim dx As Date = CDate("02/02/1990")
        'Dim dx2 As String = Format(dx, "MM\/dd\/yyyy")
        Dim dx2 As String = Replace(Format(dx, "yyyy\/MM\/dd"), "/", "")

        fd = ""
        If CheckBox1.Checked = True Then
            strsql = "UPDATE КарточкаСотрудника SET ДатаУвольнения=Null, ПриказОбУвольн=Null WHERE IDСотр=" & CType(Label6.Text, Integer) & ""
            Updates(strsql)
            Parallel.Invoke(Sub() RunMoving6())
        End If

        If CheckBox2.Checked = True Then

            'strsql = "SELECT Фамилия FROM Сотрудники WHERE КодСотрудники=" & CType(Label6.Text, Integer) & ""
            'Dim ds As DataTable = Selects(strsql)
            Dim ds = dtSotrudnikiAll.Select("КодСотрудники=" & CType(Label6.Text, Integer) & "")
            Dim files() As String
            Try
                files = IO.Directory.GetFiles(OnePath & Уволенные.ComboBox2.Text, "*" & ds(0).Item("Фамилия").ToString & " уволен" & "*.doc", IO.SearchOption.AllDirectories)
            Catch ex As Exception
                MessageBox.Show("Нет такого файла", Рик)
                Me.Cursor = Cursors.Default
                отмувсм = 0
                Exit Sub
            End Try

            Try
                IO.File.Delete(files(0))
            Catch ex As Exception
                MessageBox.Show("Нет такого файла", Рик)
                Me.Cursor = Cursors.Default
                отмувсм = 0
                Exit Sub
            End Try



        End If

        'Dim strsql4 As String = "SELECT НазвОрганиз FROM Сотрудники WHERE КодСотрудники=" & CType(Label6.Text, Integer) & ""
        Уволенные.Идент = 1

        Dim ds3 = dtSotrudnikiAll.Select("КодСотрудники=" & CType(Label6.Text, Integer) & "")
        Статистика1(TextBox2.Text, "Отмена увольнения", ds3(0).Item("НазвОрганиз").ToString)
        Me.Cursor = Cursors.Default


        Me.Close()
    End Sub

    Private Sub ОтменУвол_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub ОтменУвол_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing


    End Sub
End Class