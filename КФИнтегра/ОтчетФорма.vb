Public Class ОтчетФорма
    Private Sub ОтчетФорма_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: данная строка кода позволяет загрузить данные в таблицу "ИнтеграDataSet.ДогПодОсобен". При необходимости она может быть перемещена или удалена.
        Me.ДогПодОсобенTableAdapter.Fill(Me.ИнтеграDataSet.ДогПодОсобен)

        Me.ReportViewer1.RefreshReport()
    End Sub
End Class