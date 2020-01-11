Option Explicit On

Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Public Class Уволенные
    Public Da As New OleDbDataAdapter 'Адаптер
    Dim ds As New DataTable
    Dim tbl As New DataTable
    Dim cb As OleDb.OleDbCommandBuilder

    Dim sd As Boolean
    Dim ДатаНач, ДатаОконч, Организ, idClient, год2, год, КорИмя, Коротч, ПровПродКонтр, ПродлКонтрС, СрокОкончКонтр, ПродлКонтрПо As String
    Dim mRow As Integer = 0
    Dim newpage As Boolean = True
    Dim btprint As Boolean = False
    Public ФСодр As String
    Dim ФормаСобстПолн, ЭлАдрес, Банк, БИК, АдресБанка, РасСчет, ЮрАдрес, УНП, ДолжРуков, ФИОРукРодПад,
        ОснованиеДейств, МестоРаб, ФИОКор, ФормаСобствКор, СборноеРеквПолн, ДолжРуковРодПад, ДолжРуковВинПад, КонтТелефон As String
    Dim proc2 As Process

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Public Property Идент() As Integer = 0

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If Grid1.Rows.Count = 0 Then Exit Sub
        If idsotr = 0 Then
            MessageBox.Show("Выберите сотрудника!", Рик)
            Exit Sub
        End If

        Dim fd As Integer = Grid1.CurrentRow.Cells("КодСотрудники").Value
        ОтменУвол.TextBox1.Text = Nothing
        ОтменУвол.TextBox2.Text = Nothing
        ОтменУвол.TextBox3.Text = Nothing
        ОтменУвол.CheckBox1.Checked = False
        ОтменУвол.CheckBox2.Checked = False



        ОтменУвол.TextBox1.Text = Grid1.CurrentRow.Cells("Должность").Value
        ОтменУвол.TextBox2.Text = Grid1.CurrentRow.Cells("ФИО").Value
        ОтменУвол.TextBox3.Text = Grid1.CurrentRow.Cells("Дата увольнения").Value
        ОтменУвол.Label6.Text = fd

        ОтменУвол.ShowDialog()


        If Идент = 1 Then
            refreshgrid()
            MessageBox.Show("Сотрудник восстановлен!", Рик)
        End If
    End Sub

    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click

        If Grid1.Rows.Count = 0 Then Exit Sub
        If idsotr = 0 Then
            MessageBox.Show("Выберите сотрудника!", Рик)
            Exit Sub
        End If
        'Dim vb = From x In dtPutiDokumentovAll Where x.Item("IDСотрудник") = idsotr _
        '                                          And x.Item("ДокМесто").ToString.Contains("Заявление-Увольнение") Select x

        Using dbcx As New DbAllDataContext
            Dim var = (From x In dbcx.ПутиДокументов.AsEnumerable
                       Where x.IDСотрудник = idsotr And x.ДокМесто.Contains("Заявление-Увольнение")
                       Select x).ToList

            If var.Count > 0 Then
                ВыгрузкаФайловНаЛокалыныйКомп(var(0).ПолныйПуть, PathVremyanka & "/" & var(0).ИмяФайла)
                Dim proc As Process = Process.Start(PathVremyanka & "/" & var(0).ИмяФайла)
                proc.WaitForExit()
                proc.Close()
                ЗагрНаСерверИУдаление(PathVremyanka & "/" & var(0).ИмяФайла, var(0).ПолныйПуть, PathVremyanka & "/" & var(0).ИмяФайла)
            End If



        End Using



    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        If Grid1.Rows.Count = 0 Then Exit Sub
        If idsotr = 0 Then
            MessageBox.Show("Выберите сотрудника!", Рик)
            Exit Sub
        End If
        'Dim vb = From x In dtPutiDokumentovAll Where x.Item("IDСотрудник") = idsotr _
        '                                           And x.Item("ДокМесто").ToString.Contains("Приказ-Увольнение") Select x


        Using dbcx As New DbAllDataContext
            Dim var = (From x In dbcx.ПутиДокументов.AsEnumerable
                       Where x.IDСотрудник = idsotr And x.ДокМесто.Contains("Приказ-Увольнение")
                       Select x).ToList


            If var.Count > 0 Then
                ВыгрузкаФайловНаЛокалыныйКомп(var(0).ПолныйПуть, PathVremyanka & "/" & var(0).ИмяФайла)
                Dim proc As Process = Process.Start(PathVremyanka & "/" & var(0).ИмяФайла)
                proc.WaitForExit()
                proc.Close()
                ЗагрНаСерверИУдаление(PathVremyanka & "/" & var(0).ИмяФайла, var(0).ПолныйПуть, PathVremyanka & "/" & var(0).ИмяФайла)
            End If


        End Using






    End Sub

    Dim idsotr As Integer
    Private Sub Grid1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellClick
        idsotr = Grid1.CurrentRow.Cells("КодСотрудники").Value
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        refreshgrid()
    End Sub

    Private Sub PrintDocument2_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument2.PrintPage
        Dim bm As New Bitmap(Me.Grid1.Width, Me.Grid1.Height)
        Grid1.DrawToBitmap(bm, New Rectangle(0, 0, Me.Grid1.Width, Me.Grid1.Height))
        e.Graphics.DrawImage(bm, 0, 0)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)
        PrintDocument2.DefaultPageSettings.Landscape = True
        PrintDocument2.Print()
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs)
        'btprint = True
        'ExtrExcel()

    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        With Grid1
            Dim fmt As StringFormat = New StringFormat(StringFormatFlags.LineLimit)
            fmt.LineAlignment = StringAlignment.Center
            fmt.Trimming = StringTrimming.EllipsisCharacter
            Dim y As Single = e.MarginBounds.Top
            Do While mRow < .RowCount
                Dim row As DataGridViewRow = .Rows(mRow)
                Dim x As Single = e.MarginBounds.Left
                Dim h As Single = 0
                For Each cell As DataGridViewCell In row.Cells
                    Dim rc As RectangleF = New RectangleF(x, y, cell.Size.Width, cell.Size.Height)
                    e.Graphics.DrawRectangle(Pens.Black, rc.Left, rc.Top, rc.Width, rc.Height)
                    If (newpage) Then
                        e.Graphics.DrawString(Grid1.Columns(cell.ColumnIndex).HeaderText, .Font, Brushes.Black, rc, fmt)
                    Else
                        e.Graphics.DrawString(Grid1.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString(), .Font, Brushes.Black, rc, fmt)
                    End If
                    x += rc.Width
                    h = Math.Max(h, rc.Height)
                Next
                newpage = False
                y += h
                mRow += 1
                If y + h > e.MarginBounds.Bottom Then
                    e.HasMorePages = True
                    mRow -= 1
                    newpage = True
                    Exit Sub
                End If
            Loop
            mRow = 0
        End With
    End Sub


    Private Sub Button4_Click(sender As Object, e As EventArgs)
        PrintPreviewDialog1.Document = PrintDocument1
        PrintPreviewDialog1.ShowDialog()
    End Sub

    Private Sub SaveFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles SaveFileDialog1.FileOk

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
        Dim времянач As String = Format(DateTimePicker4.Value, "MM\/dd\/yyyy")
        Dim времякон As String = Format(DateTimePicker3.Value, "MM\/dd\/yyyy")
        Me.Cursor = Cursors.WaitCursor

        'Try
        '    If IO.File.Exists("C:\Users\Public\Downloads\Таблица3.xlsx") Then
        '        IO.File.Delete("C:\Users\Public\Downloads\Таблица3.xlsx")
        '    End If
        'Catch ex As Exception

        'End Try

        'Try
        '    If IO.File.Exists("C:\Users\Public\Downloads\Uvolennye.xlsx") Then
        '        IO.File.Delete("C:\Users\Public\Downloads\Uvolennye.xlsx")
        '    End If
        '    IO.File.Copy(OnePath & "\ОБЩДОКИ\General\Uvolennye.xlsx", "C:\Users\Public\Downloads\Uvolennye.xlsx")
        'Catch ex As Exception

        'End Try


        Dim i, j As Integer 'сохранение в эксель
        Dim xlapp As Microsoft.Office.Interop.Excel.Application
        Dim xlworkbook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlworksheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim misvalue As Object = Reflection.Missing.Value
        xlapp = New Microsoft.Office.Interop.Excel.Application
        Начало("Uvolennye.xlsx")
        xlworkbook = xlapp.Workbooks.Add(firthtPath & "\Uvolennye.xlsx")
        xlworksheet = xlworkbook.Sheets("Лист1")



        xlworksheet.Cells(1, 3) = Grid1(0, 0).Value.ToString

        For i = 0 To Grid1.Rows.Count - 1
            For j = 2 To Grid1.ColumnCount - 1

                If j = 6 Or j = 11 Or j = 8 Then
                    xlworksheet.Cells(i + 4, j) = Strings.Left(Grid1(j, i).Value.ToString, 10)
                Else
                    xlworksheet.Cells(i + 4, j) = Grid1(j, i).Value.ToString
                End If



            Next
        Next


        xlworksheet.Cells(2, 6) = DateTimePicker4.Value.ToShortDateString
        xlworksheet.Cells(2, 7) = DateTimePicker3.Value.ToShortDateString


        Try
            xlworksheet.SaveAs(firthtPath & "\Таблица3.xlsx")
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
            IO.File.Delete(firthtPath & "\Uvolennye.xlsx")


            Using proc2 = Process.Start(firthtPath & "\Таблица3.xlsx")

            End Using




            'proc.Start(firthtPath & "\Таблица3.xlsx")

            'IO.File.Delete(firthtPath & "\Таблица3.xlsx")
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
        IO.File.Delete(firthtPath & "\Uvolennye.xlsx")
        'If MessageBox.Show("Открыть файл?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
        '    Process.Start("C:\Users\Public\Downloads\Таблица.xlsx")
        'End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

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

    Private Sub DateTimePicker3_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker3.ValueChanged
        'refreshgrid()
    End Sub

    Private Sub DateTimePicker4_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker4.ValueChanged
        'refreshgrid()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        'refreshgrid()
    End Sub


    Private Sub Уволенные_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MdiParent = MDIParent1
        'Me.WindowState = FormWindowState.Maximized
        Parallel.Invoke(Sub() dtPutiDokumentov())
        год2 = Year(Now)

        год = Year(Now)
        'If Me.Прием_Load = vbTrue Then Form1.Load = False

        Me.ComboBox2.Items.Clear()
        For Each r As DataRow In СписокКлиентовОсновной.Rows
            Me.ComboBox2.Items.Add(r(0).ToString)

        Next
        MaskedTextBox1.Text = Now.Date

        DateTimePicker4.Format = DateTimePickerFormat.Short
        DateTimePicker3.Format = DateTimePickerFormat.Short

        Try
            If IO.File.Exists(firthtPath & "\Таблица3.xlsx") Then
                IO.File.Delete(firthtPath & "\Таблица3.xlsx")
            End If
        Catch ex As Exception

        End Try





        Dim cor = New Tuple(Of String, String)("flvbycr", "flvbycr") 'кортеж
        Dim f(СписокКлиентовОсновной.Rows.Count - 1) As Tuple(Of String, String, String)
        'Dim cor2(СписокКлиентовОсновной.Rows.Count - 1)
        Dim cor1(СписокКлиентовОсновной.Rows.Count - 1) As Tuple(Of String, String, String)
        Dim cor2(СписокКлиентовОсновной.Rows.Count - 1)
        For i As Integer = 0 To СписокКлиентовОсновной.Rows.Count - 1 Step 3
            Try
                cor2(i) = New Tuple(Of String, String, String)(СписокКлиентовОсновной.Rows(i).Item(0), СписокКлиентовОсновной.Rows(i + 1).Item(0), СписокКлиентовОсновной.Rows(i + 2).Item(0))
                cor1(i) = cor2(i)
            Catch ex As Exception

            End Try
        Next


        'Dim cor3 As String = cor1(3).Item2

        Dim holiday = (#07/04/2017#, Вова:="Independence Day", Буля:=True) 'кортеж имнованый 2 и 3 элементы
        Dim ff4 As String = holiday.Вова
        Dim ff5 As Boolean = holiday.Буля

    End Sub

    Private Sub refreshgrid()
        'Организ = ComboBox2.Text
        ''Dim времянач As String = Format(DateTimePicker4.Value, "MM\/dd\/yyyy")
        ''Dim времякон As String = Format(DateTimePicker3.Value, "MM\/dd\/yyyy")
        'Dim времянач As String = Replace(Format(DateTimePicker4.Value, "yyyy\/MM\/dd"), "/", "")
        'Dim времякон As String = Replace(Format(DateTimePicker3.Value, "yyyy\/MM\/dd"), "/", "")
        'tbl.Clear()

        Dim list As New Dictionary(Of String, Object)
        list.Add("@НазвОрганиз", ComboBox2.Text)
        list.Add("@начало", DateTimePicker4.Value)
        list.Add("@конец", DateTimePicker3.Value)



        Dim ds = Selects(StrSql:="SELECT Сотрудники.НазвОрганиз, Сотрудники.КодСотрудники, Штатное.Должность as Должность, Сотрудники.ФИОСборное as ФИО,
Штатное.РасчДолжностнОклад as [Расчетно должностной оклад], ДогСотрудн.Контракт as [Номер контракта], ДогСотрудн.ДатаКонтракта as [Дата контракта],
КарточкаСотрудника.ПриказОбУвольн as [Приказ об увольнении], КарточкаСотрудника.ДатаПриказаОбУвольн as [Дата приказа об увольнении], 
КарточкаСотрудника.СрокКонтракта as [Период контракта], ДогСотрудн.СрокОкончКонтр as [Дата окончания контракта], КарточкаСотрудника.ДатаУвольнения as [Дата увольнения]
FROM ((Сотрудники INNER JOIN ДогСотрудн ON Сотрудники.КодСотрудники = ДогСотрудн.IDСотр) INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр) INNER JOIN Штатное ON Сотрудники.КодСотрудники = Штатное.ИДСотр
Where Сотрудники.НазвОрганиз =@НазвОрганиз AND ((КарточкаСотрудника.ДатаУвольнения) Between @начало And @конец) ORDER BY Сотрудники.ФИОСборное", list)

        Grid1.DataSource = ds
        GridView(Grid1)
        Grid1.Columns(0).Visible = False
        Grid1.Columns(1).Visible = False
        Grid1.Columns(3).Width = 350
        Grid1.Columns(2).Width = 350
        Grid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    End Sub

    Private Sub Grid1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellDoubleClick

        Dim fd As Integer = Grid1.CurrentRow.Cells("КодСотрудники").Value
        ОтменУвол.TextBox1.Text = Nothing
        ОтменУвол.TextBox2.Text = Nothing
        ОтменУвол.TextBox3.Text = Nothing
        ОтменУвол.CheckBox1.Checked = False
        ОтменУвол.CheckBox2.Checked = False



        ОтменУвол.TextBox1.Text = Grid1.CurrentRow.Cells("Должность").Value
        ОтменУвол.TextBox2.Text = Grid1.CurrentRow.Cells("ФИО").Value
        ОтменУвол.TextBox3.Text = Grid1.CurrentRow.Cells("Дата увольнения").Value
        ОтменУвол.Label6.Text = fd

        ОтменУвол.ShowDialog()


        If ОтменУвол.отмувсм = 1 Then
            refreshgrid()
            MessageBox.Show("Сотрудник восстановлен!", Рик)
        End If



    End Sub
End Class