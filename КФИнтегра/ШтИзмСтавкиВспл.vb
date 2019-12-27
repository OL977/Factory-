Imports System.Data.OleDb
Imports System.Threading

Public Class ШтИзмСтавкиВспл
    Dim dst As DataTable
    Dim b1 As ШтИзмСтавкиВсплКласс
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub ШтИзмСтавкиВспл_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Loading()
    End Sub
    Private Sub Loading()
        b1 = New ШтИзмСтавкиВсплКласс(mast2(0), mast2(1)) 'работа с классом
        b1.таблица()
        'dst = b1.ds
        'Dim f As Integer = dst.Rows.Count
        Grid1.DataSource = b1.ds
        Label2.Text = mast2(1)

        If Grid1.Rows.Count = 1 Then
            Button3.Enabled = False
        Else
            Button3.Enabled = True
        End If

        Grid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        If Grid1.Rows.Count = 1 Then
            Dim strsql As String = "UPDATE ШтСвод SET ТарифнаяСтавка = '" & Replace(Grid1.Rows(0).Cells(4).Value, ".", ",") & "'
WHERE КодШтСвод=" & Grid1.Rows(0).Cells(0).Value & ""
            Updates(strsql)
        Else
            For i As Integer = 0 To Grid1.Rows.Count - 1
                Try
                    CDate(Grid1.Rows(i).Cells(4).Value.ToString).ToShortDateString()
                Catch ex As Exception
                    If MessageBox.Show("Введите правильно дату или нажмите отмена", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.None) = DialogResult.Cancel Then
                        Loading()
                    Else
                        Exit Sub
                    End If

                End Try
                Dim strsql As String = "UPDATE ШтСводИзмСтавка SET Дата='" & CDate(Grid1.Rows(i).Cells(4).Value.ToString).ToShortDateString & "', Ставка = '" & Replace(Grid1.Rows(i).Cells(5).Value, ".", ",") & "'
WHERE Код=" & Grid1.Rows(i).Cells(0).Value & ""
                Updates(strsql)
            Next
        End If
        MessageBox.Show("Данные изменены!", Рик)
        ШтатноеКласс1.ВыборСтавкиПоДате()
    End Sub

    Private Sub Grid1_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles Grid1.DataError

        Dim dg = b1.ds.Select("Код=" & Grid1.CurrentRow.Cells(0).Value & "")

        If MessageBox.Show("Введите правильно дату или нажмите отмена!", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.None) = DialogResult.Cancel Then
            Grid1.CurrentRow.Cells(4).Value = ""

            Grid1.CurrentRow.Cells(4).Value = Strings.Left(dg(0).Item(4).ToString, 10)
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If MessageBox.Show("Удалить выбранную строку?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Exit Sub
        End If
        Dim strsql As String = "DELETE FROM ШтСводИзмСтавка WHERE Код=" & Grid1.CurrentRow.Cells(0).Value & ""
        Updates(strsql)
        MessageBox.Show("Данные удалены!", Рик)
        ШтатноеКласс1.ВыборСтавкиПоДате()
    End Sub
End Class