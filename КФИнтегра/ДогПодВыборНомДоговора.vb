Option Explicit On
Imports System.Data.OleDb
Public Class ДогПодВыборНомДоговора
    Friend f As Integer
    Public Flag As Boolean = True
    Private Sub ДогПодВыборНомДоговора_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ВстДанВДогВЫб(ДогПодномДогПод)
    End Sub
    Public Sub ВстДанВДогВЫб(ByVal d As Integer)
        Dim strsql As String = "SELECT DISTINCT НомерДогПодр FROM ДогПодряда WHERE ID=" & d & ""
        Dim ds As DataTable = Selects(strsql)
        Try
            Me.ListBox1.Items.Clear()
        Catch ex As Exception

        End Try

        For Each r As DataRow In ds.Rows
            Me.ListBox1.Items.Add(r(0).ToString)
        Next
        Label2.Text = Прием.ComboBox19.Text

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ListBox1.SelectedItem = "" Then
            MessageBox.Show("Выберите договор!", Рик)
            Exit Sub
        End If
        ДогПодномДогПодНомДог = ListBox1.SelectedItem.ToString
        ДогПодномДогПодСтДог = ListBox1.SelectedItem.ToString
        Label2.Text = ""
        Flag = False
        Me.Close()
    End Sub

    Private Sub ДогПодВыборНомДоговора_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        e.Cancel = Flag
        If Flag = True Then
            MessageBox.Show("Выберите номер договора!", Рик)
        End If
    End Sub
End Class