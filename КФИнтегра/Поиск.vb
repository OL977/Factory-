Imports System.Data.OleDb
Imports System.IO
Public Class Поиск
    Public ds, ds1, ds2, ds3, ds4 As DataTable
    Dim StrSql, lbs, фио As String
    Dim k As Integer
    Private Delegate Sub раб()
    Private Delegate Sub орг()
    Private Sub Работники()

        If ComboBox2.InvokeRequired Then
            Me.Invoke(New раб(AddressOf Работники))
        Else
            Dim ds = From x In dtSotrudnikiAll Order By x.Item("ФИОСборное") Select x.Item("ФИОСборное")
            'StrSql = "SELECT ФИОСборное FROM Сотрудники ORDER BY ФИОСборное"
            'ds = Selects(StrSql)
            Me.ComboBox2.AutoCompleteCustomSource.Clear()
            Me.ComboBox2.Items.Clear()
            For Each r In ds
                Me.ComboBox2.AutoCompleteCustomSource.Add(r.ToString())
                Me.ComboBox2.Items.Add(r.ToString)
            Next
            ComboBox2.Enabled = False
        End If


    End Sub
    Private Sub cm2()

        Dim ds = From x In dtSotrudnikiAll Order By x.Item("ФИОСборное") Select x.Item("ФИОСборное")
        Me.ComboBox2.AutoCompleteCustomSource.Clear()
        Me.ComboBox2.Items.Clear()
        For Each r In ds
            Me.ComboBox2.AutoCompleteCustomSource.Add(r.ToString())
            Me.ComboBox2.Items.Add(r.ToString)
        Next
        ComboBox2.Enabled = False
    End Sub

    Private Async Sub cm2async()
        Await Task.Run(Sub() cm2())
    End Sub
    Private Sub cm1()
        Me.ComboBox1.AutoCompleteCustomSource.Clear()
        Me.ComboBox1.Items.Clear()
        For Each r As DataRow In СписокКлиентовОсновной.Rows
            Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox1.Items.Add(r(0).ToString)
        Next
    End Sub
    Private Async Sub Организ()
        If ComboBox1.InvokeRequired Then
            Me.Invoke(New орг(AddressOf Организ))
        Else
            Me.ComboBox1.AutoCompleteCustomSource.Clear()
            Me.ComboBox1.Items.Clear()
            For Each r As DataRow In СписокКлиентовОсновной.Rows
                Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
                Me.ComboBox1.Items.Add(r(0).ToString)
            Next
        End If
        Await Task.Delay(0)
    End Sub


    Private Sub Поиск_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Dim dуеrs As Task = New Task(AddressOf Работники)
        'Dim dуеrs1 As Task = New Task(AddressOf Организ)
        Parallel.Invoke(Sub() cm2())
        Me.MdiParent = MDIParent1
        'Me.WindowState = FormWindowState.Maximized
        'dуеrs.Start()
        'dуеrs1.Start()
        'If Me.Прием_Load = vbTrue Then Form1.Load = False
        'Parallel.Invoke(Sub() Организ())
        cm1()



    End Sub
    Private Sub Чист()
        Try
            StrSql = ""
            ds.Clear()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub ЛистБокс1()
        'Чист()


        ListBox1.Items.Clear()
        Dim ds = (From x In dtShtatnoeOtdelyAll.AsEnumerable
                  Join y In dtClientAll.AsEnumerable On x.Field(Of String)("Клиент") Equals
                     y.Field(Of String)("НазвОрг")
                  Where y.Field(Of String)("НазвОрг") = ComboBox1.Text
                  Order By x.Field(Of String)("Отделы")
                  Select x.Field(Of String)("Отделы") Distinct)


        '        StrSql = "SELECT ШтОтделы.Отделы FROM Клиент INNER JOIN ШтОтделы ON Клиент.НазвОрг = ШтОтделы.Клиент
        'WHERE Клиент.НазвОрг='" & ComboBox1.Text & "' ORDER BY ШтОтделы.Отделы "
        '        ds = Selects(StrSql)
        For i As Integer = 0 To ds.Count - 1
            ListBox1.Items.Add(ds(i).ToString)
        Next

        If CheckBox1.Checked = True And lbs <> "" Then
            ListBox1.SelectedItem = lbs
        End If

        Dim ds1 = (From x In dtSotrudnikiAll.AsEnumerable
                   Join y In dtShtatnoeAll.AsEnumerable On x.Field(Of Integer)("КодСотрудники") Equals
                     y.Field(Of Integer)("ИДСотр")
                   Where x.Field(Of String)("НазвОрганиз") = ComboBox1.Text
                   Select New With {.Должность = y.Field(Of String)("Должность"), .ФИО = x.Item("ФИОСборное"), .Разряд = y.Item("Разряд"),
                     .Тарифная_ставка = y.Item("ТарифнаяСтавка"), .ПовышениеОкладаПроцент = y.Item("ПовышОклПроц"),
                     .ПовышениеОкладаРуб = y.Item("ПовышОклРуб"), .РасчДолжностнОклад = y.Item("РасчДолжностнОклад"),
                     .ФонОплатыТруда = y.Item("ФонОплатыТруда"), .ЧасоваяТарифСтавка = y.Item("ЧасоваяТарифСтавка"),
                     .КодСотрудники = x.Item("КодСотрудники")}).ToList()

        Grid1.DataSource = ds1
        Grid1.Columns("ПовышениеОкладаРуб").HeaderText = "Повышение оклада, руб"
        Grid1.Columns("Тарифная_ставка").HeaderText = "Тарифная ставка"
        Grid1.Columns("ПовышениеОкладаПроцент").HeaderText = "Повышение оклада, %"
        Grid1.Columns("РасчДолжностнОклад").HeaderText = "Расчетно-должностной оклад, руб"
        Grid1.Columns("ФонОплатыТруда").HeaderText = "Фонд оплаты труда,"
        Grid1.Columns("ЧасоваяТарифСтавка").HeaderText = "Часовая тарифная ставка, руб"



        Grid1.Font = New Font("Microsoft Sans Serif", 8, FontStyle.Regular)
        Grid1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Try
            Grid1.Rows(0).Cells(0).Style.WrapMode = DataGridViewTriState.True
        Catch ex As Exception

        End Try
        'Grid1.Columns(1).Width = 220
        Grid1.Columns(9).Visible = False
        Grid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect ' выделяет всю строку в grid1
        Grid1.MultiSelect = False
        GridView(Grid1)






        'Grid1.Dispose()
        ПоискДоч.Hide()


    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        ЛистБокс1()

    End Sub
    Private Sub КликЛист()

        If CheckBox1.Checked = False Then

            If ComboBox1.Text = "" Then
                MessageBox.Show("Выберите организацию", Рик)
                Exit Sub
            End If

            If ListBox1.SelectedIndex = -1 Then
                MessageBox.Show("Выберите отдел!", Рик, MessageBoxButtons.OK)
                Exit Sub
            End If
            Try
                ds1.Clear()
                ds2.Clear()
            Catch ex As Exception

            End Try
            lbs = ""
            lbs = ListBox1.SelectedItem.ToString
        Else
            Try
                ds1.Clear()
                ds2.Clear()
            Catch ex As Exception

            End Try

        End If


        'Чист()
        Dim ds = (From x In dtSotrudnikiAll.AsEnumerable
                  Join y In dtShtatnoeAll.AsEnumerable On x.Field(Of Integer)("КодСотрудники") Equals
                     y.Field(Of Integer)("ИДСотр")
                  Where x.Field(Of String)("НазвОрганиз") = ComboBox1.Text And y.Field(Of String)("Отдел") = lbs
                  Select New With {.Должность = y.Field(Of String)("Должность"), .ФИО = x.Item("ФИОСборное"), .Разряд = y.Item("Разряд"),
                     .Тарифная_ставка = y.Item("ТарифнаяСтавка"), .ПовышениеОкладаПроцент = y.Item("ПовышОклПроц"),
                     .ПовышениеОкладаРуб = y.Item("ПовышОклРуб"), .РасчДолжностнОклад = y.Item("РасчДолжностнОклад"),
                     .ФонОплатыТруда = y.Item("ФонОплатыТруда"), .ЧасоваяТарифСтавка = y.Item("ЧасоваяТарифСтавка"),
                     .КодСотрудники = x.Item("КодСотрудники")}).ToList()


        '        StrSql = "Select Штатное.Должность, Сотрудники.ФИОСборное as [ФИО], Штатное.Разряд, Штатное.ТарифнаяСтавка as [Тарифная ставка],
        'Штатное.ПовышОклПроц as [Повышение оклада, %], Штатное.ПовышОклРуб as [Повышение оклада, руб] , Штатное.РасчДолжностнОклад as [Расчетно должностной оклад],
        'Штатное.ФонОплатыТруда as [ФОТ], Штатное.ЧасоваяТарифСтавка as [Часовая тарифная ставка], Сотрудники.КодСотрудники
        'From Сотрудники INNER Join Штатное On Сотрудники.КодСотрудники = Штатное.ИДСотр
        'Where Сотрудники.НазвОрганиз = '" & ComboBox1.Text & "' And Штатное.Отдел = '" & lbs & "' ORDER BY Сотрудники.ФИОСборное"
        '        ds = Selects(StrSql)


        Grid1.DataSource = ds
        Grid1.Columns("ПовышениеОкладаРуб").HeaderText = "Повышение оклада, руб"
        Grid1.Columns("Тарифная_ставка").HeaderText = "Тарифная ставка"
        Grid1.Columns("ПовышениеОкладаПроцент").HeaderText = "Повышение оклада, %"
        Grid1.Columns("РасчДолжностнОклад").HeaderText = "Расчетно-должностной оклад, руб"
        Grid1.Columns("ФонОплатыТруда").HeaderText = "Фонд оплаты труда,"
        Grid1.Columns("ЧасоваяТарифСтавка").HeaderText = "Часовая тарифная ставка, руб"



        Grid1.Font = New Font("Microsoft Sans Serif", 8, FontStyle.Regular)
        Grid1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Try
            Grid1.Rows(0).Cells(0).Style.WrapMode = DataGridViewTriState.True
        Catch ex As Exception

        End Try
        'Grid1.Columns(1).Width = 220
        Grid1.Columns(9).Visible = False
        Grid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect ' выделяет всю строку в grid1
        Grid1.MultiSelect = False
        GridView(Grid1)

        Dim fg As DataTable
        Grid3.DataSource = fg
        Grid2.DataSource = fg

        ПоискДоч.Hide()
    End Sub


    Private Sub ListBox1_Click(sender As Object, e As EventArgs) Handles ListBox1.Click
        КликЛист()

    End Sub
    Private Sub ЗапГрид2()

        k = Grid1.CurrentRow.Cells("КодСотрудники").Value

        StrSql = "SELECT Сотрудники.ПаспортСерия as [Серия], Сотрудники.ПаспортНомер as [Номер],Сотрудники.ПаспортКогдаВыдан as [Дата выдачи],
Сотрудники.ПаспортКемВыдан as [Кем выдан], Сотрудники.ИДНомер as [Идент_номер], Сотрудники.Регистрация,
Сотрудники.МестоПрожив as [Прописка], Сотрудники.КонтТелефон as [Телефон], Сотрудники.СтраховойПолис as [Полис], Сотрудники.Пол
FROM Сотрудники WHERE Сотрудники.КодСотрудники=" & k & ""
        ds1 = Selects(StrSql)
        Grid2.DataSource = ds1

        Grid2.Font = New Font("Microsoft Sans Serif", 8, FontStyle.Regular)
        Grid2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Try
            Grid2.Rows(0).Cells(0).Style.WrapMode = DataGridViewTriState.True
        Catch ex As Exception

        End Try
        GridView(Grid2)
    End Sub

    Private Sub Grid1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellClick
        ПоискДоч.Hide()
        ЗапГрид2()
        ЗапГрид3()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            ComboBox2.Enabled = True
            ComboBox1.Enabled = False
            ComboBox1.Text = String.Empty
            ListBox1.Items.Clear()
            Try
                ds.Clear()
                ds1.Clear()
                ds2.Clear()
            Catch ex As Exception

            End Try

        Else
            Try

                ComboBox2.Enabled = False
                ComboBox1.Enabled = True
                ListBox1.Items.Clear()
                ComboBox1.Text = String.Empty
                ComboBox2.Text = String.Empty
                ds.Clear()
                ds1.Clear()
                ds2.Clear()
            Catch ex As Exception

            End Try

        End If
        ПоискДоч.Hide()
    End Sub
    Private Sub ВыбПоСотр()
        Dim ds5 As DataTable
        '        StrSql = ""
        '        StrSql = "SELECT Сотрудники.КодСотрудники, Сотрудники.НазвОрганиз
        'FROM Сотрудники Where ФИОСборное = '" & ComboBox2.Text & "'"
        '        StrSql = "SELECT Сотрудники.КодСотрудники, Сотрудники.НазвОрганиз, Штатное.Отдел
        'FROM Сотрудники INNER JOIN Штатное ON Сотрудники.КодСотрудники = Штатное.ИДСотр
        'Where Сотрудники.ФИОСборное = '" & ComboBox2.Text & "'"
        'Try
        '    ds3.Clear()

        'Catch ex As Exception

        'End Try

        'ds3 = Selects(StrSql)

        Dim tdy As DataTable
        Grid3.DataSource = tdy
        Dim ds3 = dtSotrudnikiAll.Select("ФИОСборное = '" & ComboBox2.Text & "'")

        Dim strsql3 As String = "SELECT Штатное.Отдел FROM Штатное WHERE ИДСотр=" & ds3(0).Item("КодСотрудники") & ""
        ds5 = Selects(strsql3)
        lbs = ""
        If errds = 1 Then
            MessageBox.Show("С данным сотрудником заключен договор подрядка!", Рик)
        Else
            lbs = ds5.Rows(0).Item(0)
        End If
        k = 0
        k = ds3(0).Item("КодСотрудники")
        ComboBox1.Text = ds3(0).Item("НазвОрганиз").ToString
        ЛистБокс1()
        КликЛист()

        Grid1.ClearSelection() 'Поиск в Grid1
        For Each row As DataGridViewRow In Grid1.Rows
            For Each cell As DataGridViewCell In row.Cells
                If (cell.FormattedValue).Contains(ComboBox2.Text) Then
                    row.Selected = True
                    Grid1.FirstDisplayedScrollingRowIndex = row.Index
                End If
            Next
        Next

        If lbs = "" Then Exit Sub
        'Dim strsql2 As String
        'strsql2 = "SELECT Примечание FROM КарточкаСотрудника WHERE IDСотр=" & k & ""
        'Dim ds4 As DataTable
        Dim ds4 = dtKartochkaSotrudnikaAll.Select("IDСотр=" & k & "")
        'ds4 = Selects(strsql2)
        RichTextBox1.Text = ds4(0).Item("Примечание").ToString

        'ЗапГрид2()
        'ЗапГрид3()


    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        ВыбПоСотр()
    End Sub

    Private Sub ЗапГрид3()

        k = Grid1.CurrentRow.Cells("КодСотрудники").Value

        Dim ds2 = (From x In dtKartochkaSotrudnikaAll.AsEnumerable Where x.Item("IDСотр") = k
                   Select New With {.Прием = x.Item("ДатаПриема"), .Увольнение = x.Item("ДатаУвольнения"),
                   .Прод_контр = x.Item("СрокКонтракта"), .Тип_работы = x.Item("ТипРаботы"),
                      .Ставка = x.Item("Ставка"), .Дата_ПрикаУвольн = x.Item("ДатаПриказаОбУвольн"),
                      .ОснованиеУвольн = x.Item("ОснованиеУвольн"), .ДатаУведомлПродКонтр = x.Item("ДатаУведомлПродКонтр"),
                      .НомерУведомлПродКонтр = x.Item("НомерУведомлПродКонтр"), .СрокПродлКонтракта = x.Item("СрокПродлКонтракта"),
                      .ПродлКонтрС = x.Item("ПродлКонтрС"), .ПродлКонтрПо = x.Item("ПродлКонтрПо")}).ToList()


        '        StrSql = ""
        '        StrSql = "SELECT КарточкаСотрудника.ДатаПриема as [Прием], КарточкаСотрудника.ДатаУвольнения as [Увольнение],
        'КарточкаСотрудника.СрокКонтракта as [Прод_контр], КарточкаСотрудника.ТипРаботы as [Тип работы], КарточкаСотрудника.Ставка,
        'КарточкаСотрудника.ДатаПриказаОбУвольн as [Дата ПрикаУвольн], КарточкаСотрудника.ОснованиеУвольн as [Основание увольн],
        'КарточкаСотрудника.ДатаУведомлПродКонтр as [Дата УведПродл Контракта], КарточкаСотрудника.НомерУведомлПродКонтр as [Номер увед],
        'КарточкаСотрудника.СрокПродлКонтракта as [Продл_контр], КарточкаСотрудника.ПродлКонтрС as [Продл_C], КарточкаСотрудника.ПродлКонтрПо as [Продл_По]
        'FROM КарточкаСотрудника WHERE КарточкаСотрудника.IDСотр=" & k & ""
        '        ds2 = Selects(StrSql)
        Grid3.DataSource = ds2
        Grid3.Columns("Прод_контр").HeaderText = "Срок контракта, лет"
        Grid3.Columns("Тип_работы").HeaderText = "Тип работы"
        Grid3.Columns("Дата_ПрикаУвольн").HeaderText = "Дата приказа увольнения"
        Grid3.Columns("ОснованиеУвольн").HeaderText = "Основание увольнения"
        Grid3.Columns("ДатаУведомлПродКонтр").HeaderText = "Дата уведомления о продлении контракта"
        Grid3.Columns("НомерУведомлПродКонтр").HeaderText = "Номер уведомления о продлении контракта"
        Grid3.Columns("СрокПродлКонтракта").HeaderText = "Продление контракта"
        Grid3.Columns("ПродлКонтрС").HeaderText = "Дата начала продления контракта"
        Grid3.Columns("ПродлКонтрПо").HeaderText = "Дата окончания контракта"


        Grid3.Font = New Font("Microsoft Sans Serif", 8, FontStyle.Regular)
        Grid3.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Try
            Grid3.Rows(0).Cells(0).Style.WrapMode = DataGridViewTriState.True
        Catch ex As Exception

        End Try
        GridView(Grid3)
    End Sub
    Private Sub Grid3_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid3.CellClick
        'ЗапГрид3()

    End Sub
    Private Sub Доки()
        ПоискДоч.Hide()
        фио = ""
        фио = Grid1.CurrentRow.Cells("ФИО").Value
        Dim ds4 = dtSotrudnikiAll.Select("ФИОСборное='" & фио & "'")

        Dim ds5 = From d In dtPutiDokumentovAll Where Not IsDBNull(d.Item("IDСотрудник")) Select d

        Dim list2 = From x In ds5.AsEnumerable Where x.Item("IDСотрудник") = ds4(0).Item("КодСотрудники")
                    Select (x.Item("ИмяФайла"), x.Item("ПолныйПуть"))


        'StrSql = "SELECT Фамилия FROM Сотрудники WHERE ФИОСборное='" & фио & "'"
        'Try
        '    ds4.Clear()
        'Catch ex As Exception

        'End Try
        'ds4 = Selects(StrSql)

        'FilesList27 = IO.Directory.GetFiles(OnePath & ComboBox1.Text, "*" & ds4.Rows(0).Item(0).ToString & "*.doc", IO.SearchOption.AllDirectories)
        'Dim gth4 As String

        'Dim file2() As String = IO.Directory.GetFiles(OnePath & ComboBox1.Text, "*" & ds4.Rows(0).Item(0).ToString & "*.doc", IO.SearchOption.AllDirectories)


        'For n As Integer = 0 To FilesList27.Length - 1
        '    gth4 = ""
        '    gth4 = IO.Path.GetFileName(file2(n))
        '    file2(n) = gth4
        '    'TextBox44.Text &= gth + vbCrLf
        'Next

        ''ListBox2.Items.Add(Files2)

        ПоискДоч.Label1.Text = ComboBox1.Text & ", " & Grid1.CurrentRow.Cells("Должность").Value & ", " & фио & "."
        ПоискДоч.ListBox1.Items.Clear()
        FilesList27.Clear()
        For Each r In list2
            ПоискДоч.ListBox1.Items.Add(r.Item1)
            FilesList27.Add(r.Item2)
        Next



        'For i = 0 To file2.Length - 1 ' Распечатываем весь получившийся массив
        '    ПоискДоч.ListBox1.Items.Add(file2(i)) ' На ListBox2
        '    FilesList27
        'Next
        ПоискДоч.Show()

    End Sub
    Private Sub Grid1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellDoubleClick

        Доки()

    End Sub
End Class