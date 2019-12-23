Option Explicit On

Imports System.Data
Imports System.Data.OleDb
Imports System.Threading
Imports System.IO
Public Class ОтпускСписки
    Dim Norg, Dat As String
    Dim ds1, ds2 As DataTable
    Dim ds1a, ds2a As DataTable
    Dim xlapp As Microsoft.Office.Interop.Excel.Application
    Dim xlworkbook As Microsoft.Office.Interop.Excel.Workbook
    Dim xlworksheet As Microsoft.Office.Interop.Excel.Worksheet
    Dim xlapp1 As Microsoft.Office.Interop.Excel.Application
    Dim xlworkbook1 As Microsoft.Office.Interop.Excel.Workbook
    Dim xlworksheet1 As Microsoft.Office.Interop.Excel.Worksheet
    Dim misvalue As Object = Reflection.Missing.Value
    Dim ПотокЭксел, ПотокСборка As Thread
    Dim strAdres, strAdres1 As String
    Dim com1, com2, com3 As String


    Private Sub ОтпускСписки_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MdiParent = MDIParent1
        Dim fd As New Thread(AddressOf Обход1)
        fd.IsBackground = True
        fd.Start()

        ComboBox4.Items.Clear()

        For i As Integer = 1 To 12
            ComboBox2.Items.Add(MonthName(i))
            ComboBox4.Items.Add(i)
        Next


    End Sub
    Private Delegate Sub comb1()
    Private Sub Обход1()
        If ComboBox1.InvokeRequired Then
            Me.Invoke(New comb1(AddressOf Обход1))
        Else
            ComboBox1.AutoCompleteCustomSource.Clear()
            ComboBox1.Items.Clear()
            For Each r As DataRow In СписокКлиентовОсновной.Rows
                Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
                Me.ComboBox1.Items.Add(r(0).ToString)
            Next
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ComboBox2.Text = ""

        refreshgrid()


    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            ComboBox2.Text = ""
            Grid1.DataSource = ds1a
            Grid2.DataSource = ds2a
        End If
    End Sub

    Private Function ВыборкаGrid1(ByVal d As String) As List(Of Integer)
        Dim M As Integer = Month(CDate("1 " & d))
        Dim index As New List(Of Integer)(ds1a.Rows.Count)

        For Each row As DataRow In ds1a.Rows
            If row.Field(Of String)(5) = "" Then Continue For

            Dim f As Date = CDate(row.Field(Of String)(5))
            Dim Mf As Integer = Month(f)


            Dim f1 As Date = CDate(row.Field(Of String)(7))
            Dim Mf1 As Integer = Month(f1)
            Dim Mf2, Mf3 As Integer
            Try
                Dim f2 As Date = CDate(row.Field(Of String)(8))
                Mf2 = Month(f2)

                Dim f3 As Date = CDate(row.Field(Of String)(10))
                Mf3 = Month(f3)
            Catch ex As Exception
                Mf2 = 13
                Mf3 = 13
            End Try


            If Mf = M Or Mf1 = M Or Mf2 = M Or Mf3 = M Then
                index.Add(row.Field(Of Integer)(0))
            End If
        Next

        index.Sort()
        Return index
    End Function
    Private Function ВыборкаGrid2(ByVal d As String) As List(Of Integer)
        Dim M As Integer = Month(CDate("1 " & d))
        Dim index As New List(Of Integer)(ds2a.Rows.Count)

        For Each row As DataRow In ds2a.Rows
            If row.Field(Of String)(7) = "" Then Continue For

            Dim f As Date = CDate(row.Field(Of String)(7))
            Dim Mf As Integer = Month(f)


            Dim f1 As Date = CDate(row.Field(Of String)(8))
            Dim Mf1 As Integer = Month(f1)

            If Mf = M Or Mf1 = M Then
                index.Add(row.Field(Of Integer)(0))
            End If
        Next

        index.Sort()
        Return index
    End Function
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged



        If ComboBox1.Text = "" Then
            MessageBox.Show("Выберите организацию!")
            Exit Sub
        End If
        ComboBox3.Text = ""
        CheckBox1.Checked = False

        ''If ds1.Rows.Count >= 1 Then

        'Dim k As List(Of Integer) = ВыборкаGrid1(ComboBox2.Text) 'трудовой отпуск

        'Dim k As List(Of Integer) 'трудовой отпуск
        'k.Add(ds1a.Rows.Count)

        'If k.Count > 0 Then
        '    For Each row As DataRow In ds1a.Rows
        '        Dim l As New List(Of Boolean)(k.Count - 1)
        '        For i As Integer = 0 To k.Count - 1
        '            If Not row.Field(Of Integer)(0) = k(i) Then
        '                l.Add(False)
        '            Else
        '                l.Add(True)

        '            End If
        '        Next
        '        If l.Contains("true") Then
        '            l.Clear()
        '            Continue For
        '        Else
        '            row.Delete()
        '            l.Clear()
        '        End If

        '    Next
        '    ds1 = ds1a.Copy
        '    ds1a.RejectChanges()
        '    Grid1.DataSource = ds1
        'Else
        '    ds1.Clear()
        '    Grid1.DataSource = ds1
        '    ds1 = ds1a.Copy
        'End If

        'End If



        Dim li = ComboBox4.Items.Item(ComboBox2.SelectedIndex) 'linq to datagridview
        Dim dtrf = From x In ds1a.AsEnumerable() Where Month(CDate(x.Item("Начало первой части отпуска"))) = li Select x
        Grid1.DataSource = dtrf.AsDataView()


        ''If ds2.Rows.Count >= 1 Then

        'Dim k1 As List(Of Integer) = ВыборкаGrid2(ComboBox2.Text) 'социальный отпуск

        'If k1.Count > 0 Then
        '    For Each row As DataRow In ds2a.Rows
        '        Dim l As New List(Of Boolean)(k1.Count - 1)
        '        For i As Integer = 0 To k1.Count - 1
        '            If Not row.Field(Of Integer)(0) = k1(i) Then
        '                l.Add(False)
        '            Else
        '                l.Add(True)

        '            End If
        '        Next
        '        If l.Contains("true") Then
        '            l.Clear()
        '            Continue For
        '        Else
        '            row.Delete()
        '            l.Clear()
        '        End If

        '    Next


        '    ds2 = ds2a.Copy
        '    ds2a.RejectChanges()
        '    Grid2.DataSource = ds2
        'Else
        '    ds2.Clear()
        '    Grid2.DataSource = ds2
        '    ds2 = ds2a.Copy
        'End If

        Dim li2 = ComboBox4.Items.Item(ComboBox2.SelectedIndex) 'linq to datagridview
        Dim dtr = From x In ds2a.AsEnumerable() Where Month(CDate(x.Item("ПериодС"))) = li2 Select x
        Grid2.DataSource = dtr.AsDataView()

        'End If



    End Sub
    Private Function ВыборкаGrid1Год() As List(Of Integer)

        Dim index As New List(Of Integer)(ds1a.Rows.Count)

        For Each row As DataRow In ds1a.Rows
            If row.Field(Of String)(3) = "" Then Continue For

            If row.Field(Of String)(3) = ComboBox3.Text Then
                index.Add(row.Field(Of Integer)(0))
            End If
        Next

        index.Sort()
        Return index
    End Function
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        If ComboBox1.Text = "" Or ComboBox2.Text = "" Then
            MessageBox.Show("Выберите месяц и организацию!")
            ComboBox3.Text = ""
            Exit Sub
        End If

        Dim k As List(Of Integer) = ВыборкаGrid1Год()

        If k.Count > 0 Then
            For Each row As DataRow In ds1a.Rows
                Dim l As New List(Of Boolean)(k.Count - 1)
                For i As Integer = 0 To k.Count - 1
                    If Not row.Field(Of Integer)(0) = k(i) Then
                        l.Add(False)
                    Else
                        l.Add(True)

                    End If
                Next
                If l.Contains("true") Then
                    l.Clear()
                    Continue For
                Else
                    row.Delete()
                    l.Clear()
                End If

            Next
            ds1 = ds1a.Copy
            ds1a.RejectChanges()
            Grid1.DataSource = ds1
        Else
            ds1.Clear()
            Grid1.DataSource = ds1
            ds1 = ds1a.Copy
        End If


    End Sub
    Private Sub grid2excel()
        Me.Cursor = Cursors.WaitCursor
        If Grid2.Rows.Count = 0 Then
            Exit Sub
        End If

        xlapp1 = New Microsoft.Office.Interop.Excel.Application
        'xlworkbook = xlapp.Workbooks.Add(misvalue)
        Начало("OtpuskSocSpisok.xlsx")
        xlworkbook1 = xlapp1.Workbooks.Add(firthtPath & "\OtpuskSocSpisok.xlsx")
        xlworksheet1 = xlworkbook1.Sheets("dgv1")

        For ddh As Integer = 0 To Grid2.Rows.Count - 1 'основа для вставки в эксель
            With xlworksheet1
                .Cells(ddh + 5, 1) = Grid2.Rows(ddh).Cells(2).Value
                .Cells(ddh + 5, 2) = Grid2.Rows(ddh).Cells(3).Value
                .Cells(ddh + 5, 3) = Grid2.Rows(ddh).Cells(4).Value
                .Cells(ddh + 5, 4) = Grid2.Rows(ddh).Cells(5).Value
                .Cells(ddh + 5, 5) = Grid2.Rows(ddh).Cells(6).Value
                .Cells(ddh + 5, 6) = Grid2.Rows(ddh).Cells(7).Value
                .Cells(ddh + 5, 7) = Grid2.Rows(ddh).Cells(8).Value
            End With
        Next
        xlworksheet1.Range(Cell1:="A" & (5), Cell2:="G" & (Grid2.Rows.Count)).HorizontalAlignment = -4108
        xlworksheet1.Range(Cell1:="A" & (5), Cell2:="G" & (Grid2.Rows.Count + 4)).Cells.Borders.LineStyle = 1
        xlworksheet1.Cells(2, 1) = com1
        xlworksheet1.Cells(2, 4) = com2

        Dim Name As String = "Отпуск социальный " & com2 & ".xlsx"
        xlworksheet1.SaveAs(PathVremyanka & Name)
        xlworkbook1.Close()
        xlapp1.Quit()
        releaseobject(xlapp1)
        releaseobject(xlworkbook1)
        releaseobject(xlworksheet1)




        Конец(ComboBox1.Text & "\ОтпускСписки\" & Now.Year, Name, ComboBox1.Text, "\OtpuskSocSpisok.xlsx", "Отпуск социальный списки")
        ВыгрузкаФайловНаЛокалыныйКомп(FTPString & ComboBox1.Text & "\ОтпускСписки\" & Now.Year & "\Отпуск социальный " & com2 & ".xlsx", PathVremyanka & Name)
        strAdres1 = PathVremyanka & Name
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        com1 = ComboBox1.Text
        com2 = ComboBox2.Text
        strAdres1 = ""
        strAdres = ""
        If Grid1.Rows.Count = 0 Then
            If Grid2.Rows.Count = 0 Then
                Exit Sub
            Else
                grid2excel()
                Using proc As Process = Process.Start(strAdres1)

                End Using
                Exit Sub
            End If
        End If

        grid2excel()

        Me.Cursor = Cursors.WaitCursor
        xlapp = New Microsoft.Office.Interop.Excel.Application
        'xlworkbook = xlapp.Workbooks.Add(misvalue)
        Начало("OtpuskTrudSpisok.xlsx")
        xlworkbook = xlapp.Workbooks.Add(firthtPath & "\OtpuskTrudSpisok.xlsx")
        xlworksheet = xlworkbook.Sheets("dgv")


        For ddh As Integer = 0 To Grid1.Rows.Count - 1 'основа для вставки в эксель
            With xlworksheet
                .Cells(ddh + 5, 1) = Grid1.Rows(ddh).Cells(0).Value
                .Cells(ddh + 5, 2) = Grid1.Rows(ddh).Cells(1).Value
                .Cells(ddh + 5, 3) = Grid1.Rows(ddh).Cells(2).Value
                .Cells(ddh + 5, 4) = Grid1.Rows(ddh).Cells(3).Value
                .Cells(ddh + 5, 5) = Grid1.Rows(ddh).Cells(4).Value
                .Cells(ddh + 5, 6) = Grid1.Rows(ddh).Cells(5).Value
                .Cells(ddh + 5, 7) = Grid1.Rows(ddh).Cells(6).Value
                .Cells(ddh + 5, 8) = Grid1.Rows(ddh).Cells(7).Value
                .Cells(ddh + 5, 9) = Grid1.Rows(ddh).Cells(8).Value
                .Cells(ddh + 5, 10) = Grid1.Rows(ddh).Cells(9).Value
                .Cells(ddh + 5, 11) = Grid1.Rows(ddh).Cells(10).Value
                .Cells(ddh + 5, 12) = Grid1.Rows(ddh).Cells(11).Value
                .Cells(ddh + 5, 13) = Grid1.Rows(ddh).Cells(12).Value
                .Cells(ddh + 5, 14) = Grid1.Rows(ddh).Cells(13).Value
                .Cells(ddh + 5, 15) = Grid1.Rows(ddh).Cells(14).Value
                .Cells(ddh + 5, 16) = Grid1.Rows(ddh).Cells(15).Value


                '.Cells(ddh + i, 3) = ФИО(ddh).ToString
                '.Cells(ddh + i, 2) = ds2.Rows(ddh).Item(0).ToString

                '.Cells(ddh + i, 4).NumberFormat = "0.00"
                '.Cells(ddh + i, 4) = ds2.Rows(ddh).Item(2).ToString
                '.Cells(ddh + i, 5).NumberFormat = "0.00"
                '.Cells(ddh + i, 5) = ds2.Rows(ddh).Item(3).ToString
                ''If ДлПроц(ds2.Rows(ddh).Item(4).ToString) = 0 Then 'поверка на наличие десятичных знаков в процентах
                '.Cells(ddh + i, 6).NumberFormat = "0.00"
                ''End If
                '.Cells(ddh + i, 6) = ds2.Rows(ddh).Item(4).ToString
                '.Cells(ddh + i, 7).NumberFormat = "0.00"
                '.Cells(ddh + i, 7) = "=ROUND(RC[-2]*RC[-1]/100,2)"
                '.Cells(ddh + i, 8).NumberFormat = "0.00"
                '.Cells(ddh + i, 8) = "=ROUND(RC[-3]+RC[-1],2)"
                '.Cells(ddh + i, 9).NumberFormat = "0.00"
                '.Cells(ddh + i, 9) = "=ROUND(RC[-5]*RC[-1],2)"
                '.Cells(ddh + i, 10).NumberFormat = "0.00"
                '.Cells(ddh + i, 10) = "=ROUND(RC[-2]/167.3,2)"
            End With
        Next
        xlworksheet.Range(Cell1:="A" & (5), Cell2:="P" & (Grid1.Rows.Count)).HorizontalAlignment = -4108
        xlworksheet.Range(Cell1:="A" & (5), Cell2:="P" & (Grid1.Rows.Count + 4)).Cells.Borders.LineStyle = 1
        xlworksheet.Cells(2, 2) = ComboBox1.Text
        xlworksheet.Cells(2, 4) = ComboBox2.Text


        Dim Name As String = "Отпуск трудовой " & ComboBox2.Text & ".xlsx"
        xlworksheet.SaveAs(PathVremyanka & Name)
        xlworkbook.Close()
        xlapp.Quit()

        releaseobject(xlapp)
        releaseobject(xlworkbook)
        releaseobject(xlworksheet)

        Конец(ComboBox1.Text & "\ОтпускСписки\" & Now.Year, Name, ComboBox1.Text, "\OtpuskTrudSpisok.xlsx", "Отпуск трудовой списки")
        ВыгрузкаФайловНаЛокалыныйКомп(FTPString & ComboBox1.Text & "\ОтпускСписки\" & Now.Year & "\Отпуск трудовой " & ComboBox2.Text & ".xlsx", PathVremyanka & Name)
        strAdres = PathVremyanka & Name




        If MessageBox.Show("Экспорт завершен! Открыть файл?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then

            Process.Start(strAdres)
            If strAdres1 <> "" Then
                Process.Start(strAdres1)

            End If
            Me.Cursor = Cursors.Default
            Exit Sub
        Else
            Try
                IO.File.Delete(strAdres)
                IO.File.Delete(strAdres1)
            Catch ex As Exception

            End Try
            Me.Cursor = Cursors.Default
            Exit Sub
        End If




    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)
        com1 = ComboBox1.Text
        com2 = ComboBox2.Text
        strAdres1 = ""
        strAdres = ""
        If Grid1.Rows.Count = 0 Then
            If Grid2.Rows.Count = 0 Then
                Exit Sub
            Else
                grid2excel()
                Process.Start(strAdres1)
                Exit Sub
            End If
        End If

        Dim gh As New Thread(AddressOf grid2excel)
        gh.IsBackground = True
        gh.SetApartmentState(ApartmentState.STA)
        gh.Start()


        Me.Cursor = Cursors.WaitCursor
        xlapp = New Microsoft.Office.Interop.Excel.Application
        'xlworkbook = xlapp.Workbooks.Add(misvalue)
        xlworkbook = xlapp.Workbooks.Add(OnePath & "\ОБЩДОКИ\General\OtpuskTrudSpisok.xlsx")
        xlworksheet = xlworkbook.Sheets("dgv")


        For ddh As Integer = 0 To Grid1.Rows.Count - 1 'основа для вставки в эксель
            With xlworksheet
                .Cells(ddh + 5, 1) = Grid1.Rows(ddh).Cells(0).Value
                .Cells(ddh + 5, 2) = Grid1.Rows(ddh).Cells(1).Value
                .Cells(ddh + 5, 3) = Grid1.Rows(ddh).Cells(2).Value
                .Cells(ddh + 5, 4) = Grid1.Rows(ddh).Cells(3).Value
                .Cells(ddh + 5, 5) = Grid1.Rows(ddh).Cells(4).Value
                .Cells(ddh + 5, 6) = Grid1.Rows(ddh).Cells(5).Value
                .Cells(ddh + 5, 7) = Grid1.Rows(ddh).Cells(6).Value
                .Cells(ddh + 5, 8) = Grid1.Rows(ddh).Cells(7).Value
                .Cells(ddh + 5, 9) = Grid1.Rows(ddh).Cells(8).Value
                .Cells(ddh + 5, 10) = Grid1.Rows(ddh).Cells(9).Value
                .Cells(ddh + 5, 11) = Grid1.Rows(ddh).Cells(10).Value
                .Cells(ddh + 5, 12) = Grid1.Rows(ddh).Cells(11).Value
                .Cells(ddh + 5, 13) = Grid1.Rows(ddh).Cells(12).Value
                .Cells(ddh + 5, 14) = Grid1.Rows(ddh).Cells(13).Value
                .Cells(ddh + 5, 15) = Grid1.Rows(ddh).Cells(14).Value
                .Cells(ddh + 5, 16) = Grid1.Rows(ddh).Cells(15).Value


                '.Cells(ddh + i, 3) = ФИО(ddh).ToString
                '.Cells(ddh + i, 2) = ds2.Rows(ddh).Item(0).ToString

                '.Cells(ddh + i, 4).NumberFormat = "0.00"
                '.Cells(ddh + i, 4) = ds2.Rows(ddh).Item(2).ToString
                '.Cells(ddh + i, 5).NumberFormat = "0.00"
                '.Cells(ddh + i, 5) = ds2.Rows(ddh).Item(3).ToString
                ''If ДлПроц(ds2.Rows(ddh).Item(4).ToString) = 0 Then 'поверка на наличие десятичных знаков в процентах
                '.Cells(ddh + i, 6).NumberFormat = "0.00"
                ''End If
                '.Cells(ddh + i, 6) = ds2.Rows(ddh).Item(4).ToString
                '.Cells(ddh + i, 7).NumberFormat = "0.00"
                '.Cells(ddh + i, 7) = "=ROUND(RC[-2]*RC[-1]/100,2)"
                '.Cells(ddh + i, 8).NumberFormat = "0.00"
                '.Cells(ddh + i, 8) = "=ROUND(RC[-3]+RC[-1],2)"
                '.Cells(ddh + i, 9).NumberFormat = "0.00"
                '.Cells(ddh + i, 9) = "=ROUND(RC[-5]*RC[-1],2)"
                '.Cells(ddh + i, 10).NumberFormat = "0.00"
                '.Cells(ddh + i, 10) = "=ROUND(RC[-2]/167.3,2)"
            End With
        Next
        xlworksheet.Range(Cell1:="A" & (5), Cell2:="P" & (Grid1.Rows.Count)).HorizontalAlignment = -4108
        xlworksheet.Range(Cell1:="A" & (5), Cell2:="P" & (Grid1.Rows.Count + 4)).Cells.Borders.LineStyle = 1
        xlworksheet.Cells(2, 2) = ComboBox1.Text
        xlworksheet.Cells(2, 4) = ComboBox2.Text




        If Not IO.Directory.Exists(OnePath & ComboBox1.Text & "\ОтпускСписки\" & Now.Year) Then
            IO.Directory.CreateDirectory(OnePath & ComboBox1.Text & "\ОтпускСписки\" & Now.Year)
        End If
        Dim d As String = OnePath & ComboBox1.Text & "\ОтпускСписки\" & Now.Year & "\Отпуск трудовой " & ComboBox2.Text & ".xlsx"
        Try
            xlworksheet.SaveAs(d)
        Catch ex As Exception
            'KillExcel()
            'If MessageBox.Show("Такой файл уже существует! Заменить старый файл новым?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
            IO.File.Delete(d)
            xlworksheet.SaveAs(d)
            'End If
        End Try
        strAdres = d




        Dim Res As DialogResult = MessageBox.Show("Экспорт завершен!. При нажатии Да будет открыт сгенерированный файл, при нажатии Нет будет предложено сохранить файл.", Рик, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
        If Res = DialogResult.Yes Then
            xlworkbook.Close()
            xlapp.Quit()

            releaseobject(xlapp)
            releaseobject(xlworkbook)
            releaseobject(xlworksheet)

            Process.Start(strAdres)

            If strAdres1 <> "" Then
                Process.Start(strAdres1)
            End If

            Me.Cursor = Cursors.Default
            Exit Sub

        ElseIf Res = DialogResult.No Then

            'Dim времянач1 As String = времянач.Replace("/", ".")
            'Dim времякон1 As String = времякон.Replace("/", ".")

            Dim Filename As String = ""
            'SaveFileDialog1.FileName = "Принятые сотрудники предприятия_ " & ComboBox1.Text & "(" & " c " & времянач1 & " по " & времякон1 & ")"
            SaveFileDialog1.Filter = "xls files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
            SaveFileDialog1.FilterIndex = 1
            SaveFileDialog1.RestoreDirectory = True

            If SaveFileDialog1.ShowDialog = DialogResult.OK Then
                Filename = SaveFileDialog1.FileName
                xlworkbook.SaveAs(SaveFileDialog1.FileName)
                MessageBox.Show("Данные сохранены!", Рик)
                If strAdres1 <> "" Then
                    Process.Start(strAdres1)
                End If
            Else
                xlworkbook.Close()
                xlapp.Quit()

                releaseobject(xlapp)
                releaseobject(xlworkbook)
                releaseobject(xlworksheet)
                Me.Cursor = Cursors.Default
                Exit Sub
            End If


        ElseIf Res = DialogResult.Cancel Then
            MessageBox.Show("Сохранение результатов экспорта отменено!")
            Me.Cursor = Cursors.Default
        End If
        xlworkbook.Close()
        xlapp.Quit()

        releaseobject(xlapp)
        releaseobject(xlworkbook)
        releaseobject(xlworksheet)





        'Grid1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText

        'Grid1.SelectAll()
        'Clipboard.SetDataObject(Grid1.GetClipboardContent())

        'Dim путь1

        'If IO.File.Exists("C:\Users\Public\Documents\dgv.txt") = False Then
        '    IO.File.Copy(OnePath & "\ОБЩДОКИ\General\dgv.txt", "C:\Users\Public\Documents\dgv.txt")
        '    путь1 = "C:\Users\Public\Documents\dgv.txt"
        'Else
        '    путь1 = "C:\Users\Public\Documents\dgv.txt"
        'End If

        ''Записываем текст из буфера обмена в файл
        'Using writer1 As New StreamWriter(путь1, False, System.Text.Encoding.Unicode)
        '    writer1.Write(Clipboard.GetText())
        'End Using

        ''Process.Start(путь3, Chr(34) & путь1 & Chr(34))
        'Process.Start("excel.exe", Chr(34) & путь1 & Chr(34))
        'grid2excel()


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
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Grid1.Rows.Count = 0 Then
            Exit Sub
        End If

        Grid1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        Grid1.SelectAll()
        Clipboard.SetDataObject(Grid1.GetClipboardContent())

        Dim путь

        If IO.File.Exists("C:\Users\Public\Documents\dgv.html") = False Then
            IO.File.Copy(OnePath & "\ОБЩДОКИ\General\dgv.html", "C:\Users\Public\Documents\dgv.html")
            путь = "C:\Users\Public\Documents\dgv.html"
        Else
            путь = "C:\Users\Public\Documents\dgv.html"
        End If


        Using writer As New StreamWriter(путь, False, System.Text.Encoding.Unicode)
            writer.Write(Clipboard.GetText(TextDataFormat.Html))
        End Using
        Process.Start(путь)

        grid2html()
    End Sub
    Private Sub grid2html()
        If Grid2.Rows.Count = 0 Then
            Exit Sub
        End If

        Grid2.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        Grid2.SelectAll()
        Clipboard.SetDataObject(Grid2.GetClipboardContent())

        Dim путь

        If IO.File.Exists("C:\Users\Public\Documents\dgv1.html") = False Then
            IO.File.Copy(OnePath & "\ОБЩДОКИ\General\dgv1.html", "C:\Users\Public\Documents\dgv1.html")
            путь = "C:\Users\Public\Documents\dgv1.html"
        Else
            путь = "C:\Users\Public\Documents\dgv1.html"
        End If


        Using writer1 As New StreamWriter(путь, False, System.Text.Encoding.Unicode)
            writer1.Write(Clipboard.GetText(TextDataFormat.Html))
        End Using
        Process.Start(путь)
    End Sub
    Private Sub refreshgrid()
        Norg = ComboBox1.Text
        Dat = ComboBox2.Text
        Dim ds As New DataTable
        Dim list As New Dictionary(Of String, Object)
        list.Add("@Орг", Norg)
        list.Add("@Организация", Norg)

        If Not ds1 Is Nothing Then
            ds1.Clear()
        End If


        ds1 = Selects(StrSql:="SELECT ОтпускСотрудники.Код, ОтпускСотрудники.ФИО,
ОтпускСотрудники.КолДнейОтпуска as [Дней отпуска], ОтпускСотрудники.ПериодС as [Период С], ОтпускСотрудники.ПериодПо as [Период По],
ОтпускСотрудники.ДатаНач1 as [Начало первой части отпуска], ОтпускСотрудники.Продолж1 as [Продолж-ть], ОтпускСотрудники.ДатаОконч1 as [Дата окончания первой части отпуска],
ОтпускСотрудники.ДатаНач2 as [Начало второй части отпуска], ОтпускСотрудники.Продолж2 as [Продолж-ть2], ОтпускСотрудники.ДатаОконч2 as [Дата окончания второй части отпуска],
ОтпускСотрудники.ПоложеноЗаГод as [Положено за год дней отпуска], ОтпускСотрудники.Израсходовано as [Использовано], ОтпускСотрудники.ОсталосьЭтотГод as [Остаток за этот год],
ОтпускСотрудники.ОсталосьПрошлГод as [Остаток за прошлый год], ОтпускСотрудники.Итого
FROM Отпуск INNER JOIN ОтпускСотрудники ON Отпуск.Код = ОтпускСотрудники.IDОтпуск
WHERE Отпуск.Орг=@Орг ORDER BY ОтпускСотрудники.ФИО", list)




        If Not ds2 Is Nothing Then
            ds2.Clear()
        End If
        ds2 = Selects(StrSql:="SELECT * FROM ОтпускСоц WHERE Организация=@Организация ORDER BY Сотрудник", list)


        ds1a = ds1.Copy
        Grid1.DataSource = ds1
        Grid1.Columns(0).Width = 50
        Grid1.Columns(1).Width = 250
        Grid1.Columns(6).Width = 100
        Grid1.Columns(9).Width = 100
        Grid1.Columns(12).Width = 100
        Grid1.EnableHeadersVisualStyles = False
        Grid1.ColumnHeadersDefaultCellStyle.BackColor = Color.LightGreen
        Grid1.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Raised
        Grid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Grid1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        ds2a = ds2.Copy
        Grid2.DataSource = ds2
        Grid2.Columns(0).Visible = False
        Grid2.Columns(1).Visible = False
        Grid2.Columns(2).Width = 250
        Grid2.EnableHeadersVisualStyles = False
        Grid2.ColumnHeadersDefaultCellStyle.BackColor = Color.LightGreen
        Grid2.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Raised
        Grid2.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Grid2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter


        Dim cm1 As Task = New Task(AddressOf Год)
        cm1.Start()
    End Sub

    Private Delegate Sub CombxDel1()
    Private Sub Год()

        If ComboBox3.InvokeRequired Then
            Me.Invoke(New CombxDel1(AddressOf Год))
        Else
            Dim strsql As String = "SELECT DISTINCT ПериодС FROM ОтпускСотрудники"
            Dim ds As DataTable = Selects(strsql)
            Me.ComboBox3.AutoCompleteCustomSource.Clear()
            Me.ComboBox3.Items.Clear()
            For Each r As DataRow In ds.Rows
                Me.ComboBox3.AutoCompleteCustomSource.Add(r.Item(0).ToString())
                Me.ComboBox3.Items.Add(r(0).ToString)
            Next
            ComboBox3.Text = ""
        End If

    End Sub
End Class