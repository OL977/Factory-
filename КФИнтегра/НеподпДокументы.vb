Option Explicit On

Imports System.Data
Imports System.Data.OleDb
Imports System.Threading
Public Class НеподпДокументы
    Dim ИДСотр As Integer
    Dim btprint As Boolean = False
    Private Delegate Sub comb2()
    Private Sub Обход()
        If ComboBox1.InvokeRequired Then
            Me.Invoke(New comb2(AddressOf Обход))
        Else
            Me.ComboBox2.AutoCompleteCustomSource.Clear()
            Me.ComboBox2.Items.Clear()
            For Each r As DataRow In СписокКлиентовОсновной.Rows
                Me.ComboBox2.AutoCompleteCustomSource.Add(r.Item(0).ToString())
                Me.ComboBox2.Items.Add(r(0).ToString)
            Next
        End If
    End Sub
    Private Sub НеподпДокументы_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MdiParent = MDIParent1
        Dim df As New Thread(AddressOf Обход)
        df.IsBackground = True
        df.Start()
        MaskedTextBox2.Text = Now.Date
        'Me.TextBox1.Text = DateTime.Now.ToString("dd.MM.yyyy")
        ComboBox1.Enabled = False

    End Sub
    Private Sub ОтборПоОрганизПрим()
        Dim strsql As String
        Dim df As String

        Dim list As New Dictionary(Of String, Object)
        list.Add("@НазвОрганиз", ComboBox2.Text)


        Dim ds = Selects(StrSql:="Select Сотрудники.КодСотрудники, Сотрудники.НазвОрганиз as [Организация], Сотрудники.ФИОСборное as [ФИО],
КарточкаСотрудника.ДатаПриема as [Дата приема], Штатное.Отдел, Штатное.Должность, ДогСотрудн.Контракт, ДогСотрудн.Приказ, КарточкаСотрудника.ПриказОбУвольн as [Приказ об увольнении], КарточкаСотрудника.ПриказПродлКонтр as [Приказ о продл_контр],
КарточкаСотрудника.НомУведИзмСрокЗарп as [Номер уведомления изменения срока зарплаты], КарточкаСотрудника.Примечание
FROM((Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр) INNER JOIN Штатное On Сотрудники.КодСотрудники = Штатное.ИДСотр) INNER JOIN ДогСотрудн On Сотрудники.КодСотрудники = ДогСотрудн.IDСотр
WHERE Сотрудники.НазвОрганиз=@НазвОрганиз and КарточкаСотрудника.Примечание Is Not Null ORDER BY Сотрудники.ФИОСборное", list)
        Grid1.DataSource = ds
        GridView(Grid1)
        Grid1.Columns(0).Visible = False
        Grid1.Columns(1).Width = 150
        Grid1.Columns(2).Width = 300
    End Sub
    Private Sub ОтборПоОрганиз()
        Dim list As New Dictionary(Of String, Object)
        list.Add("@НазвОрганиз", ComboBox2.Text)


        Dim ds = Selects(StrSql:="Select Сотрудники.КодСотрудники, Сотрудники.НазвОрганиз as [Организация], Сотрудники.ФИОСборное as [ФИО],
КарточкаСотрудника.ДатаПриема as [Дата приема], Штатное.Отдел, Штатное.Должность, ДогСотрудн.Контракт, ДогСотрудн.Приказ, КарточкаСотрудника.ПриказОбУвольн as [Приказ об увольнении], КарточкаСотрудника.ПриказПродлКонтр as [Приказ о продл_контр],
КарточкаСотрудника.НомУведИзмСрокЗарп as [Номер уведомления изменения срока зарплаты], КарточкаСотрудника.Примечание
FROM((Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр) INNER JOIN Штатное On Сотрудники.КодСотрудники = Штатное.ИДСотр) INNER JOIN ДогСотрудн On Сотрудники.КодСотрудники = ДогСотрудн.IDСотр
WHERE Сотрудники.НазвОрганиз  =@НазвОрганиз ORDER BY Сотрудники.ФИОСборное", list)

        Grid1.DataSource = ds
        GridView(Grid1)
        Grid1.Columns(0).Visible = False
        Grid1.Columns(1).Width = 150
        Grid1.Columns(2).Width = 300

    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If CheckBox1.Checked = False Then
            ОтборПоОрганизПрим()
        Else
            ОтборПоОрганиз()
        End If

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If ComboBox2.Text = "" Then
            MessageBox.Show("Выберите организацию", Рик)
            CheckBox1.Checked = False
            Exit Sub
        End If
        If CheckBox1.Checked = True Then
            ОтборПоОрганиз()
        Else
            If CheckBox2.Checked = False Then
                ОтборПоОрганизПрим()
            End If

        End If
    End Sub

    Private Sub ЗагрСотр()
        Dim ds = From x In dtSotrudnikiAll Order By x.Item("ФИОСборное") Select x
        'StrSql = "SELECT ФИОСборное,КодСотрудники FROM Сотрудники ORDER BY ФИОСборное"
        'Dim ds As DataTable
        'ds = Selects(StrSql)
        Me.ComboBox1.AutoCompleteCustomSource.Clear()
        Me.ComboBox1.Items.Clear()
        For Each r In ds
            Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item("ФИОСборное").ToString())
            Me.ComboBox1.Items.Add(r.Item("ФИОСборное").ToString())
        Next

        Me.ComboBox3.Items.Clear()
        For Each r In ds
            Me.ComboBox3.Items.Add(r.Item("КодСотрудники").ToString())
        Next


        ИДСотр = ds.First.Item("КодСотрудники")
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            ComboBox1.Enabled = True
            ComboBox2.Text = ""
            ComboBox2.Enabled = False
            CheckBox1.Enabled = False
            ЗагрСотр()

        Else
            ComboBox1.Enabled = False
            ComboBox1.Text = ""
            ComboBox2.Enabled = True
            CheckBox1.Enabled = True
        End If
    End Sub
    Public Sub ПоискПоСотр()
        Label3.Text = ComboBox3.Items.Item(ComboBox1.SelectedIndex)

        Dim list As New Dictionary(Of String, Object)
        list.Add("@КодСотрудники", CType(Label3.Text, Integer))

        Dim ds = Selects(StrSql:= "Select Сотрудники.КодСотрудники, Сотрудники.НазвОрганиз as [Организация], Сотрудники.ФИОСборное as [ФИО],
КарточкаСотрудника.ДатаПриема as [Дата приема], Штатное.Отдел, Штатное.Должность, ДогСотрудн.Контракт, ДогСотрудн.Приказ, КарточкаСотрудника.ПриказОбУвольн as [Приказ об увольнении], КарточкаСотрудника.ПриказПродлКонтр as [Приказ о продл_контр],
КарточкаСотрудника.НомУведИзмСрокЗарп as [Номер уведомления изменения срока зарплаты], КарточкаСотрудника.Примечание
FROM((Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр) INNER JOIN Штатное On Сотрудники.КодСотрудники = Штатное.ИДСотр) INNER JOIN ДогСотрудн On Сотрудники.КодСотрудники = ДогСотрудн.IDСотр
WHERE Сотрудники.КодСотрудники=@КодСотрудники ORDER BY Сотрудники.ФИОСборное", list)

        Grid1.DataSource = ds
        GridView(Grid1)
        Grid1.Columns(0).Visible = False
        Grid1.Columns(1).Width = 150
        Grid1.Columns(2).Width = 300

    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ПоискПоСотр()
    End Sub

    Private Sub Grid1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellDoubleClick
        Прим = Nothing
        Прим = 2
        Примечание.Label2.Text = Grid1.CurrentRow.Cells(0).Value.ToString
        Примечание.RichTextBox1.Text = Grid1.CurrentRow.Cells(11).Value.ToString
        Примечание.TextBox2.Text = Grid1.CurrentRow.Cells(2).Value.ToString
        Примечание.ShowDialog()

        If CheckBox2.Checked = False And CheckBox1.Checked = False Then
            ОтборПоОрганизПрим()
        ElseIf CheckBox2.Checked = False And CheckBox1.Checked = True Then
            ОтборПоОрганиз()
        Else
            ПоискПоСотр()
        End If

    End Sub
    Private Sub ExtrExcel()

        If ComboBox2.Text = "" Then
            MessageBox.Show("Выберите организацию!")
            Exit Sub
        End If
        If Grid1.Rows.Count < 1 Then
            MessageBox.Show("Нет данных для экспорта!")
            Exit Sub
        End If

        If ComboBox2.Text = "" And ComboBox1.Text = "" Then
            MessageBox.Show("Выберите сотрудника!")
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor

        'Try
        '    If IO.File.Exists("C:\Users\Public\Downloads\Dokumenty.xlsx") Then
        '        IO.File.Delete("C:\Users\Public\Downloads\Dokumenty.xlsx")
        '    End If
        'Catch ex As Exception

        'End Try
        'Dim время As String = Format(TextBox1.Text, "MM\/dd\/yyyy")


        Dim i, j As Integer 'сохранение в эксель
        Dim xlapp As Microsoft.Office.Interop.Excel.Application
        Dim xlworkbook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlworksheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim misvalue As Object = Reflection.Missing.Value
        xlapp = New Microsoft.Office.Interop.Excel.Application
        Начало("\Dokumenty.xlsx")
        xlworkbook = xlapp.Workbooks.Add(firthtPath & "\Dokumenty.xlsx")
        xlworksheet = xlworkbook.Sheets("Лист1")



        xlworksheet.Cells(1, 2) = Grid1(1, 0).Value.ToString
        xlworksheet.Cells(1, 6) = Now.ToShortDateString

        For i = 0 To Grid1.Rows.Count - 1
            For j = 2 To Grid1.ColumnCount - 1
                'For k As Integer = 1 To Grid1.Columns.Count
                'xlworksheet.Cells(1, k) = Grid1.Columns(k - 1).HeaderText
                'Grid1(7, i).Value = Strings.Left(Grid1(7, i).Value.ToString, 10)
                If j = 3 Then
                    xlworksheet.Cells(i + 4, j) = Strings.Left(Grid1(j, i).Value.ToString, 10)
                Else
                    xlworksheet.Cells(i + 4, j) = Grid1(j, i).Value.ToString
                End If


                'Next
            Next
        Next



        'xlworksheet.Cells(2, 6) = времянач
        'xlworksheet.Cells(2, 7) = времякон


        Try
            xlworksheet.SaveAs(firthtPath & "\Dokumenty.xlsx")
            If btprint = True Then
                xlworksheet.PrintOutEx()
                xlworkbook.Close()
                xlapp.Quit()

                releaseobject(xlapp)
                releaseobject(xlworkbook)
                releaseobject(xlworksheet)
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show("Закройте эксель файл 'Таблица' и повторите попытку")
            Me.Cursor = Cursors.Default
            Exit Sub
        End Try



        Dim Res As DialogResult = MessageBox.Show("Экспорт завершен!.При нажатии Да будет открыт сгенерированный файл, при нажатии Нет будет предложено сохранить файл.", Рик, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
        If Res = DialogResult.Yes Then
            xlworkbook.Close()
            xlapp.Quit()

            releaseobject(xlapp)
            releaseobject(xlworkbook)
            releaseobject(xlworksheet)

            Process.Start(firthtPath & "\Dokumenty.xlsx")
            Me.Cursor = Cursors.Default
            Exit Sub

        ElseIf Res = DialogResult.No Then


            Dim Filename As String = ""
            SaveFileDialog1.FileName = "Уволенные сотрудники предприятия_ " & ComboBox2.Text
            SaveFileDialog1.Filter = "xls files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
            SaveFileDialog1.FilterIndex = 1
            SaveFileDialog1.RestoreDirectory = True

            If SaveFileDialog1.ShowDialog = DialogResult.OK Then
                Filename = SaveFileDialog1.FileName
                xlworkbook.SaveAs(SaveFileDialog1.FileName)

            End If


        ElseIf Res = DialogResult.Cancel Then
            MessageBox.Show("Сохранение результатов экспорта отменено!")

        End If
        xlworkbook.Close()
        xlapp.Quit()

        releaseobject(xlapp)
        releaseobject(xlworkbook)
        releaseobject(xlworksheet)
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs)
        ExtrExcel()
    End Sub
    Private Sub releaseobject(ByVal obj As Object)
        Try
            Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        btprint = True
        ExtrExcel()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ExtrExcel()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If Grid1.CurrentRow Is Nothing Then
            MessageBox.Show("Выберите сотрудника!", Рик)
            Exit Sub
        End If


        Прим = Nothing
        Прим = 2
        Примечание.Label2.Text = Grid1.CurrentRow.Cells(0).Value.ToString
        Примечание.RichTextBox1.Text = Grid1.CurrentRow.Cells(11).Value.ToString
        Примечание.TextBox2.Text = Grid1.CurrentRow.Cells(2).Value.ToString
        Примечание.ShowDialog()

        If CheckBox2.Checked = False And CheckBox1.Checked = False Then
            ОтборПоОрганизПрим()
        ElseIf CheckBox2.Checked = False And CheckBox1.Checked = True Then
            ОтборПоОрганиз()
        Else
            ПоискПоСотр()
        End If
    End Sub
End Class