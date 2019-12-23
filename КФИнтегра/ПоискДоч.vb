Imports System.Data.OleDb
Imports System.IO
Public Class ПоискДоч

    Private Sub ПоискДоч_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub ListBox1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ListBox1.MouseDoubleClick
        If ListBox1.SelectedIndex = -1 Then
            MessageBox.Show("Выберите документ для просмотра!", Рик, MessageBoxButtons.OK)
            Exit Sub
        End If

        Dim i As Integer = ListBox1.SelectedIndex
        ВыгрузкаФайловНаЛокалыныйКомп(FilesList27(i), PathVremyanka & ListBox1.SelectedItem)
        Dim proc As Process = Process.Start(PathVremyanka & ListBox1.SelectedItem)
        proc.WaitForExit()
        proc.Close()
        ЗагрНаСерверИУдаление(PathVremyanka & ListBox1.SelectedItem, FilesList27(i), PathVremyanka & ListBox1.SelectedItem)


    End Sub
End Class