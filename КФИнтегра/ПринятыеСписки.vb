Imports System.Data
Imports System.Data.OleDb

Public Class ПринятыеСписки
    Public Da As New OleDbDataAdapter 'Адаптер
    Dim tbl As New DataTable
    Dim ds As New DataTable
    Dim cb As OleDb.OleDbCommandBuilder
    Dim btprint As Boolean = False
    Dim sd As Boolean
    Dim ДатаНач, ДатаОконч, Организ, idClient, год2, год, КорИмя, Коротч, времянач, времякон, времянач1, времякон1 As String
    Dim idsotr As Integer

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If Grid1.Rows.Count = 0 Then Exit Sub
        If idsotr = 0 Then
            MessageBox.Show("Выберите сотрудника!", Рик)
            Exit Sub
        End If
        Dim vb = From x In dtPutiDokumentovAll Where x.Item("IDСотрудник") = idsotr _
                                                   And x.Item("ДокМесто").ToString.Contains("Прием-Приказ") Select x

        If vb.ElementAtOrDefault(0) IsNot Nothing Then
            ВыгрузкаФайловНаЛокалыныйКомп(vb(0).Item("ПолныйПуть"), PathVremyanka & "/" & vb(0).Item("ИмяФайла"))
            Dim proc As Process = Process.Start(PathVremyanka & "/" & vb(0).Item("ИмяФайла"))
            proc.WaitForExit()
            proc.Close()

            ЗагрНаСерверИУдаление(PathVremyanka & "/" & vb(0).Item("ИмяФайла"), vb(0).Item("ПолныйПуть"), PathVremyanka & "/" & vb(0).Item("ИмяФайла"))
        End If

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If Grid1.Rows.Count = 0 Then Exit Sub
        If idsotr = 0 Then
            MessageBox.Show("Выберите сотрудника!", Рик)
            Exit Sub
        End If
        Dim vb = From x In dtPutiDokumentovAll Where x.Item("IDСотрудник") = idsotr _
                                                   And x.Item("ДокМесто").ToString.Contains("Прием-Зявление") Select x

        If vb.ElementAtOrDefault(0) IsNot Nothing Then
            ВыгрузкаФайловНаЛокалыныйКомп(vb(0).Item("ПолныйПуть"), PathVremyanka & "/" & vb(0).Item("ИмяФайла"))
            Dim proc As Process = Process.Start(PathVremyanka & "/" & vb(0).Item("ИмяФайла"))
            proc.WaitForExit()
            proc.Close()
            ЗагрНаСерверИУдаление(PathVremyanka & "/" & vb(0).Item("ИмяФайла"), vb(0).Item("ПолныйПуть"), PathVremyanka & "/" & vb(0).Item("ИмяФайла"))
        End If

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If Grid1.Rows.Count = 0 Then Exit Sub
        If idsotr = 0 Then
            MessageBox.Show("Выберите сотрудника!", Рик)
            Exit Sub
        End If
        Dim vb = From x In dtPutiDokumentovAll Where x.Item("IDСотрудник") = idsotr _
                                                   And x.Item("ДокМесто").ToString.Contains("Прием-Контракт") Select x


        If vb.ElementAtOrDefault(0) IsNot Nothing Then
            ВыгрузкаФайловНаЛокалыныйКомп(vb(0).Item("ПолныйПуть"), PathVremyanka & "/" & vb(0).Item("ИмяФайла"))
            Dim proc As Process = Process.Start(PathVremyanka & "/" & vb(0).Item("ИмяФайла"))
            proc.WaitForExit()
            proc.Close()
            ЗагрНаСерверИУдаление(PathVremyanka & "/" & vb(0).Item("ИмяФайла"), vb(0).Item("ПолныйПуть"), PathVremyanka & "/" & vb(0).Item("ИмяФайла"))
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        btprint = False
        ExtrExcel()
    End Sub

    Private Sub Grid1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellClick
        idsotr = Grid1.CurrentRow.Cells("КодСотрудники").Value
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        refreshgrid()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        btprint = False
        ExtrExcel()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        btprint = True
        ExtrExcel()
    End Sub

    Private Sub DateTimePicker3_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker3.ValueChanged
        'refreshgrid()
    End Sub

    Private Sub DateTimePicker4_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker4.ValueChanged
        'refreshgrid()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        'refreshgrid()
    End Sub

    Private Sub ПринятыеСписки_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MdiParent = MDIParent1
        год2 = Year(Now)
        год = Year(Now)
        'If Me.Прием_Load = vbTrue Then Form1.Load = False

        Me.ComboBox2.Items.Clear()
        For Each r As DataRow In СписокКлиентовОсновной.Rows
            Me.ComboBox2.Items.Add(r(0).ToString)
        Next
        MaskedTextBox1.Text = Now.Date

        Try
            If IO.File.Exists(firthtPath & "\Таблица2.xlsx") Then
                IO.File.Delete(firthtPath & "\Таблица2.xlsx")
            End If
        Catch ex As Exception

        End Try





        DateTimePicker4.Format = DateTimePickerFormat.Short
        DateTimePicker3.Format = DateTimePickerFormat.Short
    End Sub
    Private Sub refreshgrid()
        Организ = ComboBox2.Text


        'времянач = Format(DateTimePicker4.Value, "MM\/dd\/yyyy")
        'времякон = Format(DateTimePicker3.Value, "MM\/dd\/yyyy")

        'времянач1 = Replace(Format(DateTimePicker4.Value, "yyyy\/MM\/dd"), "/", "")
        'времякон1 = Replace(Format(DateTimePicker3.Value, "yyyy\/MM\/dd"), "/", "")

        Dim list As New Dictionary(Of String, Object)
        list.Add("@НазвОрганиз", Организ)
        list.Add("@начало", DateTimePicker4.Value)
        list.Add("@конец", DateTimePicker3.Value)


        '        StrSql = "SELECT Сотрудники.КодСотрудники as [ID], Сотрудники.НазвОрганиз as Наименование, Сотрудники.ФИОСборное as [ФИО Сотрудника], ДогСотрудн.ДатаКонтракта as [Дата приема сотрудника], 
        'ДогСотрудн.Контракт as [Контракт],КарточкаСотрудника.СрокКонтракта as [Продолжительность контракта, лет], 
        'ДогСотрудн.СрокОкончКонтр as [Дата окончания контракта]
        'From (Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр) INNER JOIN ДогСотрудн ON Сотрудники.КодСотрудники = ДогСотрудн.IDСотр
        'Where Сотрудники.НазвОрганиз = '" & Организ & "' AND ((ДогСотрудн.ДатаКонтракта) Between #" & времянач & "# And #" & времякон & "#) ORDER BY Сотрудники.ФИОСборное"
        Dim ds = Selects(StrSql:="SELECT Сотрудники.НазвОрганиз, Сотрудники.КодСотрудники, Штатное.Должность as Должность, Сотрудники.ФИОСборное as ФИО,
Штатное.РасчДолжностнОклад as [Расчетно должностной оклад], ДогСотрудн.Контракт as [Номер контракта], ДогСотрудн.ДатаКонтракта as [Дата контракта],
КарточкаСотрудника.СрокКонтракта as [Период контракта], ДогСотрудн.СрокОкончКонтр as [Дата окончания контракта], КарточкаСотрудника.ДатаУвольнения as [Дата увольнения], ДогСотрудн.Перевод
FROM ((Сотрудники INNER JOIN ДогСотрудн ON Сотрудники.КодСотрудники = ДогСотрудн.IDСотр) INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр) INNER JOIN Штатное ON Сотрудники.КодСотрудники = Штатное.ИДСотр
Where Сотрудники.НазвОрганиз =@НазвОрганиз AND ((ДогСотрудн.ДатаКонтракта) Between @начало And @конец) ORDER BY Сотрудники.ФИОСборное", list) 'AND ((ДогСотрудн.ДатаКонтракта) Between '" & времянач1 & "' And '" & времякон1 & "') ORDER BY Сотрудники.ФИОСборное" 

        'ds = Selects(StrSql)

        Grid1.DataSource = ds
        GridView(Grid1)
        Grid1.Columns(0).Visible = False
        Grid1.Columns(1).Visible = False
        Grid1.Columns(3).Width = 250
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

        Me.Cursor = Cursors.WaitCursor
        'Try
        '    If IO.File.Exists("C:\Users\Public\Downloads\Prinyatye.xlsx") Then
        '        IO.File.Delete("C:\Users\Public\Downloads\Prinyatye.xlsx")
        '    End If
        '    IO.File.Copy(OnePath & "\ОБЩДОКИ\General\Prinyatye.xlsx", "C:\Users\Public\Downloads\Prinyatye.xlsx")
        'Catch ex As Exception

        'End Try

        'Try
        '    If IO.File.Exists("C:\Users\Public\Downloads\Таблица2.xlsx") Then
        '        IO.File.Delete("C:\Users\Public\Downloads\Таблица2.xlsx")
        '    End If
        'Catch ex As Exception

        'End Try


        Dim i, j As Integer 'сохранение в эксель
        Dim xlapp As Microsoft.Office.Interop.Excel.Application
        Dim xlworkbook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlworksheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim misvalue As Object = Reflection.Missing.Value
        xlapp = New Microsoft.Office.Interop.Excel.Application
        Начало("Prinyatye.xlsx")
        xlworkbook = xlapp.Workbooks.Add(firthtPath & "\Prinyatye.xlsx")
        xlworksheet = xlworkbook.Sheets("Лист1")


        xlworksheet.Cells(1, 3) = Grid1(0, 0).Value.ToString

        For i = 0 To Grid1.Rows.Count - 1
            For j = 2 To Grid1.ColumnCount - 1
                If j = 6 Or j = 9 Then
                    xlworksheet.Cells(i + 4, j) = Strings.Left(Grid1(j, i).Value.ToString, 10)
                Else
                    xlworksheet.Cells(i + 4, j) = Grid1(j, i).Value.ToString
                End If
            Next
        Next

        xlworksheet.Cells(2, 6) = времянач
        xlworksheet.Cells(2, 7) = времякон

        Try

            '    xlworksheet.Range("I4:I39").Select()
            '    xlworksheet.Selection
            '    xlworksheet.Range("E4").Select()
            '    xlworksheet.Selection.PasteSpecial(Paste:=xlworksheet.xlPasteValues, Operation:=xlworksheet.xlNone, SkipBlanks _
            ':=False, Transpose:=False)
            '    xlworksheet.Range("I3").Select()
            '    xlworksheet.Application.CutCopyMode = False
            '    xlworksheet.ActiveCell.FormulaR1C1 = ""
            '    xlworksheet.Range("K3").Select()






            xlworksheet.SaveAs(firthtPath & "\Таблица2.xlsx")
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
            releaseobject(xlapp)
            releaseobject(xlworkbook)
            releaseobject(xlworksheet)
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
            IO.File.Delete(firthtPath & "\Prinyatye.xlsx")

            Using proc As Process = Process.Start(firthtPath & "\Таблица2.xlsx")

            End Using
            'proc.Start(firthtPath & "\Таблица3.xlsx")
            'Proc.WaitForExit()
            'IO.File.Delete(firthtPath & "\Таблица2.xlsx")
            Me.Cursor = Cursors.Default
            Exit Sub

        ElseIf Res = DialogResult.No Then

            Dim времянач1 As String = времянач.Replace("/", ".")
            Dim времякон1 As String = времякон.Replace("/", ".")

            Dim Filename As String = ""
            SaveFileDialog1.FileName = "Принятые сотрудники предприятия_ " & ComboBox2.Text & "(" & " c " & времянач1 & " по " & времякон1 & ")"
            SaveFileDialog1.Filter = "xls files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
            SaveFileDialog1.FilterIndex = 1
            SaveFileDialog1.RestoreDirectory = True

            If SaveFileDialog1.ShowDialog = DialogResult.OK Then
                Filename = SaveFileDialog1.FileName
                xlworkbook.SaveAs(SaveFileDialog1.FileName)
                MessageBox.Show("Данные сохранены!", Рик)
            Else
                xlworkbook.Close()
                xlapp.Quit()

                releaseobject(xlapp)
                releaseobject(xlworkbook)
                releaseobject(xlworksheet)
                Me.Cursor = Cursors.Default
                IO.File.Delete(firthtPath & "\Prinyatye.xlsx")
                Exit Sub
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
        IO.File.Delete(firthtPath & "\Prinyatye.xlsx")
        'If MessageBox.Show("Открыть файл?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
        '    Process.Start("C:\Users\Public\Downloads\Таблица.xlsx")
        'End If
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
End Class