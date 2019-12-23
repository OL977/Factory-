Option Explicit On
Imports System.Data.OleDb
Public Class ПереводИстория
    Dim ds As DataTable
    Private Sub ПереводИстория_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim StrSql As String = "SELECT НазвОрг FROM Клиент ORDER BY НазвОрг"
        Dim ds As DataTable = Selects(StrSql)
        Me.ComboBox1.AutoCompleteCustomSource.Clear()
        Me.ComboBox1.Items.Clear()
        For Each r As DataRow In ds.Rows
            Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox1.Items.Add(r(0).ToString)
            ПереводОрганиз.ListBox1.Items.Add(r(0).ToString)
        Next
    End Sub
    Private Sub refreshgrid(ByVal d As Integer)
        Dim strsql As String
        If d = 1 Then
            strsql = "SELECT ДолжСтар as [Старая должность], ДолжНов as [Новая должность], 
ДатаДолжНов as [Дата перевода],  РазрСтар as [Старый разряд], РазрНов as [Новый разряд],
ТарифСтар as [Старая ставка], ТарифНов as [Новая ставка], ФИОСотр
FROM Перевод
WHERE Организация='" & ComboBox1.Text & "'"
        Else
            strsql = "SELECT ДолжСтар as [Старая должность], ДолжНов as [Новая должность], 
ДатаДолжНов as [Дата перевода],  РазрСтар as [Старый разряд], РазрНов as [Новый разряд],
ТарифСтар as [Старая ставка], ТарифНов as [Новая ставка], ФИОСотр
FROM Перевод
WHERE Организация='" & ComboBox1.Text & "' and ФИОСотр='" & ComboBox2.Text & "'"

        End If
        Try
            ds.Clear()
        Catch ex As Exception

        End Try

        ds = Selects(strsql)
        If errds = 1 Then Exit Sub
        Grid1.DataSource = ds
        Dim dc As DataGridViewColumn = Grid1.Columns(7)
        dc.DisplayIndex = 0 ' Индекс  для отображения
        Grid1.Columns(0).Width = 250
        Grid1.Columns(1).Width = 250
        Grid1.Columns(7).Width = 250

        'Grid1.Columns(0).Visible = False
        'Grid1.Columns(1).Visible = False


        Grid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect ' выделяет всю строку в grid1
        Grid1.MultiSelect = False


    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim StrSql As String = "SELECT ФИОСборное FROM Сотрудники WHERE НазвОрганиз='" & ComboBox1.Text & "' ORDER BY ФИОСборное "
        Dim ds As DataTable = Selects(StrSql)
        Me.ComboBox2.AutoCompleteCustomSource.Clear()
        Me.ComboBox2.Items.Clear()
        For Each r As DataRow In ds.Rows
            Me.ComboBox2.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox2.Items.Add(r(0).ToString)
            ПереводСотрудн.ListBox1.Items.Add(r(0).ToString)
        Next
        ComboBox2.Text = ""
        refreshgrid(1)
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        refreshgrid(2)
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        ds.Clear()
        ComboBox1.Items.Clear()
        ComboBox2.Items.Clear()
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        Me.Close()
    End Sub

    Private Sub ПереводИстория_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        Try
            ComboBox1.Items.Clear()
            ComboBox2.Items.Clear()
            ComboBox1.Text = ""
            ComboBox2.Text = ""
            ds.Clear()
        Catch ex As Exception

        End Try
    End Sub
End Class