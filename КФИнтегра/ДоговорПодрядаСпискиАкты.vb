Option Explicit On
Imports System.Data.OleDb
Imports System.Threading
Imports MySql.Data.MySqlClient
Imports System.Management
Imports System.ComponentModel
Imports System.IO

Public Class ДоговорПодрядаСпискиАкты
    Private Sub ДоговорПодрядаСпискиАкты_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim ds As DataTable = ДоговорПодрядаСписки.dsart

        Label3.Text = ds.Rows(0).Item(9).ToString
        Label4.Text = ds.Rows(0).Item(1).ToString
        Grid1.DataSource = ds
        Grid1.Columns(0).Visible = False
        Grid1.Columns(1).Visible = False
        Grid1.Columns(9).Visible = False
        Grid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Grid1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter 'по центру
        Grid1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub

    Private Sub ДоговорПодрядаСпискиАкты_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Grid1.Rows.Count = 0 Then
            Exit Sub
        End If
        Grid1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        Grid1.SelectAll()

        Clipboard.SetDataObject(Grid1.GetClipboardContent())


        'Dim j As String = "C:\Users\Public\Documents\dgv.txt"
        'If IO.File.Exists(j) = False Then
        '    IO.File.Copy(OnePath & "\ОБЩДОКИ\General\dgv.txt", j)
        '    путь1 = j
        'Else
        '    путь1 = j
        'End If

        Начало("dgv.txt")
        Dim путь1 = firthtPath & "\dgv.txt"
        'Записываем текст из буфера обмена в файл
        Using writer As New StreamWriter(путь1, False, System.Text.Encoding.Unicode)
            writer.Write(Clipboard.GetText())
        End Using

        'Process.Start(путь3, Chr(34) & путь1 & Chr(34))
        Process.Start("excel.exe", Chr(34) & путь1 & Chr(34))
    End Sub


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Grid1.Rows.Count = 0 Then
            Exit Sub
        End If

        Grid1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        Grid1.SelectAll()
        Clipboard.SetDataObject(Grid1.GetClipboardContent())

        Начало("dgv.html")
        Dim путь = firthtPath & "\dgv.html"
        'Dim j As String = "C:\Users\Public\Documents\dgv.html"
        'If IO.File.Exists(j) = False Then
        '    IO.File.Copy(OnePath & "\ОБЩДОКИ\General\dgv.html", j)
        '    путь = j
        'Else
        '    путь = j
        'End If


        Using writer As New StreamWriter(путь, False, System.Text.Encoding.Unicode)
            writer.Write(Clipboard.GetText(TextDataFormat.Html))
        End Using
        Process.Start(путь)



    End Sub
End Class