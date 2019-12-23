Option Explicit On
Imports System.Data.OleDb
Public Class ДоговорПодрядаАктИное
    Dim КодАкт As Integer
    Dim ПослНомАкт As Integer
    Dim strsql, strsql1 As String
    Dim ds1, ds2 As DataTable
    Dim file2() As String
    Dim FilesList() As String
    Dim СохрЗак As String
    Dim НадоБновл As Boolean
    Private Delegate Sub CombxDel1()
    Dim v As Integer
    Private Delegate Sub Orgd(ByVal d As String)
    'Dim a1 As String() = {TextBox1.Text, TextBox6.Text, TextBox7.Text, TextBox8.Text, TextBox9.Text}
    'Dim a2 As String() = {TextBox11.Text, TextBox10.Text, TextBox5.Text, TextBox3.Text, TextBox2.Text}
    'Dim a3 As String() = {TextBox16.Text, TextBox15.Text, TextBox14.Text, TextBox13.Text, TextBox12.Text}
    'Dim a4 As String() = {TextBox21.Text, TextBox20.Text, TextBox19.Text, TextBox18.Text, TextBox17.Text}
    Dim a0()
    Dim Код As Integer
    Dim datrow As DataTable

    Private Async Sub Com1()

        If ComboBox1.InvokeRequired Then

            Me.Invoke(New CombxDel1(AddressOf Com1))
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

    Private Sub ДоговорПодрядаАктИное_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dtPutiDokumentov() 'обновляем пути документов

        Dim task As Task = New Task(AddressOf Com1)
        task.Start()

        MaskedTextBox1.Enabled = False

    End Sub
    Public Sub Очистка2() 'очистка контролов
        For Each groupboxControl In Me.Controls.OfType(Of GroupBox)() 'очистка контролов внутри гроупбоксов
            For Each txt In groupboxControl.Controls.OfType(Of TextBox)()
                txt.Text = ""
            Next
            'For Each cbo In groupboxControl.Controls.OfType(Of ComboBox)()
            '    cbo.SelectedIndex = -1
            'Next
            For Each mas In groupboxControl.Controls.OfType(Of MaskedTextBox)()
                mas.Text = ""
            Next
            'For Each txt In groupboxControl.Controls.OfType(Of TextBox)()
            '    txt.Text = ""
            'Next
        Next

        ComboBox5.Text = ""

        Dim dt As New DataTable
        Grid1.DataSource = dt
    End Sub
    Public Sub Очистка() 'очистка контролов

        'For Each F_Control As Control In F.Controls
        '    Dim _control As Object = F.Controls(F_Control.Name)
        '    If TypeOf _control Is TextBox Then
        '        _control.Text = ""
        '    ElseIf TypeOf _control Is ListBox Then
        '        _control.items.clear()
        '        ElseIf TypeOf _control Is ComboBox Then
        '            _control.selectedindex = -1
        '        ElseIf TypeOf _control Is RichTextBox Then
        '            _control.text = ""
        '        ElseIf TypeOf _control Is MaskedTextBox Then
        '            _control.text = ""
        '        End If
        'Next F_Control

        'For Each F_Control As Object In F.Controls
        '    If TypeOf F.Controls(F_Control.Name) Is TextBox Then
        '        F.Controls(F_Control.Name).Text = "-"
        '    ElseIf TypeOf F.Controls(F_Control.Name) Is ListBox Then
        '        F.Controls(F_Control.Name).items.clear()
        '        'ElseIf TypeOf F.Controls(F_Control.Name) Is ComboBox Then
        '        '    F.Controls(F_Control.Name).selectedindex = -1
        '    End If
        'Next F_Control

        For Each groupboxControl In Me.Controls.OfType(Of GroupBox)() 'очистка контролов внутри гроупбоксов
            For Each txt In groupboxControl.Controls.OfType(Of TextBox)()
                txt.Text = ""
            Next
            'For Each cbo In groupboxControl.Controls.OfType(Of ComboBox)()
            '    cbo.SelectedIndex = -1
            'Next
            For Each mas In groupboxControl.Controls.OfType(Of MaskedTextBox)()
                mas.Text = ""
            Next
            For Each txt In Me.Controls.OfType(Of TextBox)()
                txt.Text = ""
            Next
        Next

        ComboBox3.Items.Clear()
        ComboBox3.Text = ""
        ComboBox5.Text = ""
        ComboBox5.Items.Clear()
        ListBox1.Items.Clear()
        Dim dt As New DataTable
        Grid1.DataSource = dt

    End Sub
    Private Sub ComboBox19_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox19.SelectedIndexChanged
        Очистка()

        Com1Sel9()

    End Sub
    Private Sub Com1Sel()
        Dim list As New Dictionary(Of String, Object)
        list.Add("@НазвОрганиз", ComboBox1.SelectedItem)

        Dim ds = Selects(StrSql:="SELECT DISTINCT Сотрудники.ФИОСборное, ДогПодряда.ID
FROM Сотрудники INNER JOIN ДогПодряда ON Сотрудники.КодСотрудники = ДогПодряда.ID
WHERE Сотрудники.НазвОрганиз=@НазвОрганиз AND СтоимРуб1 IS NOT Null ORDER BY Сотрудники.ФИОСборное", list)

        If ComboBox19.Text <> "" Then 'проверка стоит ли чистить все поля
            Очистка()
        End If


        Label96.Text = "N"
        ComboBox2.Items.Clear()
        Me.ComboBox19.AutoCompleteCustomSource.Clear()
        Me.ComboBox19.Items.Clear()
        For Each r As DataRow In ds.Rows
            Me.ComboBox19.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox19.Items.Add(r(0).ToString)
            Me.ComboBox2.Items.Add(r(1).ToString)
        Next



    End Sub
    Private Sub refreshList2()

        Dim ds2 = From x In dtSotrudnikiAll Where Not IsDBNull(x.Item("ФИОСборное")) = ComboBox19.Text Select x.Item("Фамилия")
        Dim id
        If IsNumeric(CType(Label96.Text, Integer)) Then
            id = CType(Label96.Text, Integer)
        Else
            Throw New System.Exception("Нет идентификатора сотрудника!")

        End If

        'Dim StrSql2 As String = "Select Фамилия From Сотрудники Where ФИОСборное ='" & ComboBox19.Text & "'"
        'Dim c2 As New OleDbCommand With {
        '    .Connection = conn,
        '    .CommandText = StrSql2
        '}
        'Dim ds2 As New DataSet
        'Dim da2 As New OleDbDataAdapter(c2)
        'da2.Fill(ds2, "Ставка2")

        Dim list = From x In (From x In dtPutiDokumentovAll.AsEnumerable Where Not IsDBNull(x.Item("IDСотрудник")) Select x) Where x.Item("IDСотрудник") = id And x.Item("ДокМесто").ToString.Contains("Акт договор подряда иное") Select x

        ListBox1.Items.Clear()
        ComboBox4.Items.Clear()

        For Each f In list ' Распечатываем весь получившийся массив
            ListBox1.Items.Add(f.Item("ИмяФайла").ToString) ' На ListBox2
            ComboBox4.Items.Add(f.Item("ПолныйПуть"))
        Next


        'FilesList = Nothing
        'file2 = Nothing
        'Dim gth4 As String
        'Try
        '    FilesList = IO.Directory.GetFiles(OnePath & ComboBox1.Text, "*" & ds2(0).Item(0) & "*.doc*", IO.SearchOption.AllDirectories)
        '    file2 = IO.Directory.GetFiles(OnePath & ComboBox1.Text, "*" & ds2(0).Item(0) & "*.doc*", IO.SearchOption.AllDirectories)
        '    For n As Integer = 0 To FilesList.Length - 1
        '        gth4 = ""
        '        gth4 = IO.Path.GetFileName(file2(n))
        '        file2(n) = gth4
        '        'TextBox44.Text &= gth + vbCrLf
        '    Next
        '    ListBox1.Items.Clear()

        '    For i = 0 To file2.Length - 1 ' Распечатываем весь получившийся массив
        '        ListBox1.Items.Add(file2(i)) ' На ListBox2
        '    Next
        'Catch ex As Exception

        'End Try



        'ListBox2.Items.Add(Files2)



    End Sub
    Private Sub refreshList2(ByVal name As String)

        'Dim ds2 = From x In dtSotrudnikiAll Where x.Item("ФИОСборное") = ComboBox19.Text Select x.Item("Фамилия")
        Dim id
        If IsNumeric(CType(Label96.Text, Integer)) Then
            id = CType(Label96.Text, Integer)
        Else
            Throw New System.Exception("Нет идентификатора сотрудника!")

        End If


        Dim list = From x In (From x1 In dtPutiDokumentovAll.AsEnumerable Where Not IsDBNull(x1.Item("IDСотрудник")) Select x1)
                   Where x.Item("IDСотрудник") = id _
                                                 And x.Item("ДокМесто").ToString.Contains("Акт договор подряда иное") _
                                                 And x.Item("ИмяФайла").ToString.Contains(name) _
                                                 And x.Item("ИмяФайла").ToString.Contains(ComboBox5.Text) Select x

        '(From x1 In dtPutiDokumentovAll.AsEnumerable Where Not IsDBNull(x1.Item("IDСотрудник")) Select x1)


        ListBox1.Items.Clear()
        ComboBox4.Items.Clear()

        For Each f In list ' Распечатываем весь получившийся массив
            ListBox1.Items.Add(f.Item("ИмяФайла").ToString) ' На ListBox2
            ComboBox4.Items.Add(f.Item("ПолныйПуть"))
        Next



    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            If MaskedTextBox2.Text <> "" Then
                MaskedTextBox3.Text = MaskedTextBox2.Text
                MaskedTextBox6.Text = MaskedTextBox2.Text
            End If
        End If
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Com1Sel()
    End Sub


    Private Sub MaskedTextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            MaskedTextBox3.Focus()
        End If
    End Sub

    Private Sub MaskedTextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            MaskedTextBox6.Focus()
        End If
    End Sub




    Private Sub txt1(ByVal txt As String)
        'Dim f As Double = CDbl(txt)
        'f = Math.Round(f / CDbl(TextBox2.Text), 2)
        'f = Replace(f, ".", ",")
        'If СправкаПоЗарплате.bool(f) = True Then
        '    TextBox1.Text = f & ",00"
        'Else
        '    TextBox1.Text = f
        '    If СправкаПоЗарплате.Count(f) = 1 Then
        '        TextBox1.Text = f & "0"
        '    End If
        'End If

    End Sub

    Private Sub ОформлTxt(ByVal ds As DataTable)

        'ComboBox3.Text = ds(0).Item(2).ToString
        datrow = ds.Copy()



        If ds.Rows.Count > 1 Then
            'Dim m = From x In ds

            'ds = ds.AsEnumerable().Distinct(DataRowComparer.Default).ToArray()
            ComboBox5.Items.Clear()
            For Each r In ds.Rows
                If ComboBox5.Items.Contains(r.Item("ПорНомерАктаИное").ToString) Then
                    Continue For
                End If
                ComboBox5.Items.Add(r.Item("ПорНомерАктаИное").ToString)

            Next
        Else

            MessageBox.Show("Нет готовых актов для данного договора!", Рик)


            'If ds(0).Item(19).ToString.Length = 1 Then
            '    ComboBox5.Text = "00" & ds(0).Item(19).ToString
            'ElseIf ds(0).Item(19).ToString.Length = 2 Then
            '    ComboBox5.Text = "0" & ds(0).Item(19).ToString
            'Else
            '    ComboBox5.Text = ds(0).Item(19).ToString
            'End If
            ''TextBox4.Text = ds.Rows(0).Item(19).ToString
            'MaskedTextBox1.Text = ds(0).Item(23).ToString
            'MaskedTextBox2.Text = ds(0).Item(20).ToString
            'MaskedTextBox3.Text = ds(0).Item(21).ToString
            'MaskedTextBox6.Text = ds(0).Item(22).ToString

        End If




    End Sub

    Private Sub ОбщДанПоДог2()
        Dim j = CType(Label96.Text, Integer)
        'Dim ds = From x In dtDogovorPadriadaAll Where x.Item("ID") = j And x.Item("НомерДогПодр") = ComboBox3.Text Select x
        Dim ds
        Try
            ds = (From x In dtDogovorPadriadaAll.AsEnumerable
                  Join y In dtDogPodrAktInoeAll.AsEnumerable On x.Field(Of Integer)("Код") Equals
                          y.Field(Of Integer)("IDДогПодряда")
                  Where Not x.IsNull("ID") AndAlso x.Field(Of Integer)("ID") = j AndAlso Not x.IsNull("НомерДогПодр") _
                      AndAlso x.Field(Of String)("НомерДогПодр") = ComboBox3.Text
                  Select y).CopyToDataTable()
        Catch ex As Exception
            MessageBox.Show("Нет готовых актов для данного договора!", Рик)
            ComboBox5.Items.Clear()
            Exit Sub
        End Try




        'Dim ds = dtDogovorPadriadaAll.Select("ID=" & j & " And НомерДогПодр=" & ComboBox3.Text & "")



        ОформлTxt(ds)


        Parallel.Invoke(Sub() refreshList2(ComboBox3.Text))
    End Sub
    Private Sub ОбщДанПоДог()
        Dim j
        If IsNumeric(CType(Label96.Text, Integer)) Then
            j = CType(Label96.Text, Integer)
        Else
            Throw New Exception("Не определен идентификатор сотрудника!")
        End If
        'Dim ds = From x In dtDogovorPadriadaAll Where x.Item("ID") = j Select x
        Dim ds = dtDogovorPadriadaAll.Select("ID=" & j & "")

        'Dim strsql As String = "SELECT * FROM ДогПодряда WHERE ID=" & CType(Label96.Text, Integer) & ""
        'Dim ds As DataTable = Selects(strsql)

        'ОформлTxt(ds)

    End Sub
    Private Sub ОтборДоговоров()

        Dim j = CType(Label96.Text, Integer)
        Dim ds = From x In dtDogovorPadriadaAll Where Not x.IsNull("ID") AndAlso x.Item("ID") = j Select x.Item("НомерДогПодр") Distinct


        'Dim strsql As String = "SELECT DISTINCT НомерДогПодр FROM ДогПодряда WHERE ID=" & CType(Label96.Text, Integer) & ""
        'Dim ds As DataTable = Selects(strsql)

        'Dim df As Integer = ds.Count
        'If df = 1 Then
        '    ОбщДанПоДог()
        'Else

        ComboBox3.AutoCompleteCustomSource.Clear()
        ComboBox3.Items.Clear()
        For Each r In ds
            Me.ComboBox3.AutoCompleteCustomSource.Add(r)
            Me.ComboBox3.Items.Add(r)
        Next
    End Sub
    Private Sub Com1Sel9()

        Label96.Text = ComboBox2.Items.Item(ComboBox19.SelectedIndex)
        'If ОтборДоговоров() = 0 Then
        '    refreshList2()
        'End If

        ОтборДоговоров()



    End Sub
    Private Function Проверка()
        If ComboBox1.Text = "" Or ComboBox19.Text = "" Then
            MessageBox.Show("Выберите организацию и сотрудника!", Рик)
            Return 1
        End If

        If ComboBox5.Text = "" Then
            MessageBox.Show("Заполните номер акта!", Рик)
            Return 1
        End If

        If MaskedTextBox2.MaskCompleted = False Or MaskedTextBox3.MaskCompleted = False Then
            MessageBox.Show("Заполните раздел 'период'!", Рик)
            Return 1
        End If

        If MaskedTextBox6.MaskCompleted = False Then
            MessageBox.Show("Заполните дату акта!", Рик)
            Return 1
        End If

        If MaskedTextBox1.MaskCompleted = False Then
            MessageBox.Show("Выберите дату оплаты работ!", Рик)
            Return 1
        End If

        'If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Then
        '    MessageBox.Show("Заполните раздел 'Отработанное время и начисленная сумма'!", Рик)
        '    Return 1
        'End If

        Return 0
    End Function
    Private Function ПровДубл()
        НадоБновл = False
        Dim df = From x In dtDogovorPadriadaAll Where x.Item("ID") = CType(Label96.Text, Integer) And x.Item("НомерДогПодр").ToString = ComboBox3.Text And x.Item("ПорНомерАктаИное").ToString = ComboBox5.Text Select x

        'Dim strsql As String = "SELECT ПорНомерАктаИное FROM ДогПодряда WHERE ID=" & CType(Label96.Text, Integer) & " and НомерДогПодр='" & ComboBox3.Text & "' AND ПорНомерАктаИное='" & TextBox4.Text & "'"
        'Dim df As DataTable = Selects(strsql)

        If df.Count > 0 Then
            If MessageBox.Show("Заменить старые данные акта №" & CType(ComboBox5.Text, Integer) & " -новыми?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                Return 1
            End If
            НадоБновл = True
        End If
        Return 0
    End Function
    Private Sub НовСтрока(ByVal strsql2 As String, ByVal код As Integer)

        Dim c As New OleDbCommand
        c.Connection = conn
        c.CommandText = strsql2
        Dim ds As New DataSet
        Dim da As New OleDbDataAdapter(c)
        da.Fill(ds, "Сохранение")

        Dim cb As New OleDbCommandBuilder(da)
        Dim dsNewRow As DataRow

        dsNewRow = ds.Tables("Сохранение").NewRow()
        dsNewRow.Item("IDДогПодр") = код
        dsNewRow.Item("ПорНомерАкта") = ComboBox5.Text
        dsNewRow.Item("ЗаПериодС") = MaskedTextBox2.Text
        dsNewRow.Item("ЗаПериодПо") = MaskedTextBox3.Text
        'dsNewRow.Item("ВремяРабот") = TextBox1.Text
        'dsNewRow.Item("СтоимЧаса") = TextBox2.Text
        'dsNewRow.Item("СтоимРабот") = TextBox3.Text
        dsNewRow.Item("ДатаАкта") = MaskedTextBox6.Text
        dsNewRow.Item("ДатаОплатыРабот") = MaskedTextBox1.Text
        ds.Tables("Сохранение").Rows.Add(dsNewRow)

        'ds.Tables("Сохранение").Rows(0).Item(0) = a
        'ds.Tables("Сохранение").Rows(0).Item(1) = Me.TextBox1.Text
        da.Update(ds, "Сохранение")

    End Sub
    Private Sub СохрВБазу()
        Dim strsql As String
        Dim ds3 As DataTable
        Dim objlist As New ArrayList()

        Dim list As New Dictionary(Of String, Object)
        list.Add("@ID", CType(Label96.Text, Integer))
        list.Add("@НомерДогПодр", ComboBox3.Text)
        list.Add("@ПорНомерАктаИное", ComboBox5.Text)


        Updates(stroka:="UPDATE ДогПодряда SET ЗаПериодСИное='" & MaskedTextBox2.Text & "',ЗаПериодПоИное='" & MaskedTextBox3.Text & "',
ДатаАктаИное='" & MaskedTextBox6.Text & "',ДатаОплатыРаботИное='" & MaskedTextBox1.Text & "'
WHERE ID=@ID AND НомерДогПодр=@НомерДогПодр AND ПорНомерАктаИное=@ПорНомерАктаИное", list, "ДогПодряда")










        '        If НадоБновл = True Then 'изменение существующего акта
        '            For i As Integer = 0 To v - 1
        '                Dim list As New Dictionary(Of String, Object)
        '                list.Add("@ID", CType(Label96.Text, Integer))
        '                list.Add("@НомерДогПодр", ComboBox3.Text)
        '                list.Add("@ПорНомерАктаИное", ComboBox5.Text)
        '                list.Add("@ВыпРаб1", a0(i)(0))

        '                Updates(stroka:="UPDATE ДогПодряда SET ОбъемВыпРаботАктИное='" & a0(i)(3) & "',ЕдИзмерАктИное='" & a0(i)(1) & "',
        'ОбщСтоимРаботАктИное='" & a0(i)(4) & "',ЗаПериодСИное='" & MaskedTextBox2.Text & "',ЗаПериодПоИное='" & MaskedTextBox3.Text & "',
        'ДатаАктаИное='" & MaskedTextBox6.Text & "',ДатаОплатыРаботИное='" & MaskedTextBox1.Text & "'
        'WHERE ID=@ID AND НомерДогПодр=@НомерДогПодр AND ПорНомерАктаИное=@ПорНомерАктаИное AND ВыпРаб1=@ВыпРаб1", List, "ДогПодряда")
        '                list.Clear()
        '            Next
        '            MessageBox.Show("Данные изменены!", Рик)
        '        Else
        '            For i As Integer = 0 To v - 1
        '                Dim list2 As New Dictionary(Of String, Object)
        '                list2.Add("@ID", CType(Label96.Text, Integer))
        '                list2.Add("@НомерДогПодр", ComboBox3.Text)
        '                list2.Add("@ПорНомерАктаИное", ComboBox5.Text) 'создание нового акта
        '                list2.Add("@ОбъемВыпРаботАктИное", a0(i)(3))
        '                list2.Add("@ЕдИзмерАктИное", a0(i)(1))
        '                list2.Add("@СтоимЕдРаботыАктИное", a0(i)(2))
        '                list2.Add("@ОбщСтоимРаботАктИное", a0(i)(4))
        '                list2.Add("@ЗаПериодСИное", MaskedTextBox2.Text)
        '                list2.Add("@ЗаПериодПоИное", MaskedTextBox3.Text)
        '                list2.Add("@ДатаАктаИное", MaskedTextBox6.Text)
        '                list2.Add("@ДатаОплатыРаботИное", MaskedTextBox1.Text)
        '                list2.Add("@ВыпРаб1", a0(i)(0))




        '                Updates(stroka:="INSERT INTO ДогПодряда(ОбъемВыпРаботАктИное,ЕдИзмерАктИное,СтоимЕдРаботыАктИное,
        'ОбщСтоимРаботАктИное, ЗаПериодСИное, ЗаПериодПоИное, ДатаАктаИное,ДатаОплатыРаботИное, ПорНомерАктаИное, iD, НомерДогПодр,ВыпРаб1)
        'VALUES(@ОбъемВыпРаботАктИное,@ЕдИзмерАктИное,@СтоимЕдРаботыАктИное,@ОбщСтоимРаботАктИное,@ЗаПериодСИное, @ЗаПериодПоИное,
        '@ДатаАктаИное,@ДатаОплатыРаботИное,@ПорНомерАктаИное,@ID,@НомерДогПодр,@ВыпРаб1)", list2, "ДогПодряда")






        '                list2.Clear()
        '            Next
        '            MessageBox.Show("Данные добавлены в базу!", Рик)

        '        End If





    End Sub
    Private Sub Обновление()
        '        Dim strsql As String = "UPDATE ДогПодрядаАкт SET ПорНомерАкта=" & CType(TextBox4.Text, Integer) & ",ЗаПериодС='" & MaskedTextBox2.Text & "',ЗаПериодПо='" & MaskedTextBox3.Text & "',
        'ВремяРабот='" & TextBox1.Text & "',СтоимЧаса='" & TextBox2.Text & "', СтоимРабот='" & TextBox3.Text & "', ДатаАкта='" & MaskedTextBox6.Text & "', ДатаОплатыРабот='" & MaskedTextBox1.Text & "'
        '        WHERE ДогПодрядаАкт.Код=" & КодАкт & ""
        '        Updates(strsql)
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'a0 = {({TextBox1.Text, TextBox6.Text, TextBox7.Text, TextBox8.Text, TextBox9.Text}), ({TextBox11.Text, TextBox10.Text, TextBox5.Text, TextBox3.Text, TextBox2.Text}),
        '    ({TextBox16.Text, TextBox15.Text, TextBox14.Text, TextBox13.Text, TextBox12.Text}), ({TextBox21.Text, TextBox20.Text, TextBox19.Text, TextBox18.Text, TextBox17.Text})}
        If Проверка() = 1 Then Exit Sub
        'If ПровДубл() = 1 Then Exit Sub
        СохрВБазу()
        If CheckBox4.Checked = False Then
            Доки()
        End If
        'Me.Close()
    End Sub
    Private Function ДогПодНом()
        'Dim strsql As String = "SELECT DISTINCT ДатаДогПодр FROM ДогПодряда WHERE ID=" & CType(Label96.Text, Integer) & " AND НомерДогПодр='" & ComboBox3.Text & "'"

        'Dim df As DataTable = Selects(strsql)

        Dim df = From x In dtDogovorPadriadaAll Where x.Item("ID") = CType(Label96.Text, Integer) And x.Item("НомерДогПодр") = ComboBox3.Text Select x.Item("ДатаДогПодр") Distinct
        Return df
    End Function
    Private Function ТекстРаботИменПадеж(ByVal d As String) As String
        'Dim strsql As String = "SELECT ТесктИменПад FROM ДогПодОсобен WHERE Текст='" & d & "'"
        'Dim ds As DataTable = Selects(strsql)

        Dim ds = dtPodriadaOsobenAll.Select("Текст='" & d & "'")
        Dim df As String = ds(0).Item("ТесктИменПад").ToString
        Return df

    End Function
    Private Sub Доки()

        'Dim oWord As Word.Application
        'Dim oDoc As Word.Document

        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document
        Me.Cursor = Cursors.WaitCursor

        oWord = CreateObject("Word.Application")
        oWord.Visible = False
        Dim delstring As String
        'Select Case v
        '    Case 1
        '        Начало("ActPodriadaInoe1.doc")
        '        oWordDoc = oWord.Documents.Add(firthtPath & "\ActPodriadaInoe1.doc")
        '        delstring = "\ActPodriadaInoe1.doc"
        '    Case 2
        '        Начало("ActPodriadaInoe2.doc")
        '        oWordDoc = oWord.Documents.Add(firthtPath & "\ActPodriadaInoe2.doc")
        '        delstring = "\ActPodriadaInoe2.doc"
        '    Case 3
        '        Начало("ActPodriadaInoe3.doc")
        '        oWordDoc = oWord.Documents.Add(firthtPath & "\ActPodriadaInoe3.doc")
        '        delstring = "\ActPodriadaInoe3.doc"
        '    Case 4
        '        Начало("ActPodriadaInoe4.doc")
        '        oWordDoc = oWord.Documents.Add(firthtPath & "\ActPodriadaInoe4.doc")
        '        delstring = "\ActPodriadaInoe4.doc"
        '    Case Else
        '        Начало("ActPodriadaInoe1.doc")
        '        delstring = "\ActPodriadaInoe1.doc"
        '        oWordDoc = oWord.Documents.Add(firthtPath & "\ActPodriadaInoe1.doc")
        'End Select





        'Try
        '    IO.File.Copy(strcon, "C:\Users\Public\Documents\Рик\ActPodriada.doc")
        'Catch ex As Exception
        '    If ex.Message.Contains("уже существует") Then
        '        Try
        '            IO.File.Delete("C:\Users\Public\Documents\Рик\ActPodriada.doc")
        '            IO.File.Copy(strcon, "C:\Users\Public\Documents\Рик\ActPodriada.doc")
        '        Catch e As System.IO.IOException
        '            If e.Message.Contains("используется другим процессом") Then
        '                For Each p As Process In Process.GetProcessesByName("winword")
        '                    p.Kill()
        '                    p.WaitForExit()
        '                Next
        '            End If
        '        End Try
        '    End If
        '    IO.File.Delete("C:\Users\Public\Documents\Рик\ActPodriada.doc")
        '    IO.File.Copy(strcon, "C:\Users\Public\Documents\Рик\ActPodriada.doc")
        'End Try



        Начало("ActPodriadaInoe5.doc")
        oWordDoc = oWord.Documents.Add(firthtPath & "\ActPodriadaInoe5.doc")

        Dim Организация = Org(ComboBox1.Text)
        Dim Сотрудник = Sotrudnic(CType(Label96.Text, Integer))

        Dim ДП As IEnumerable(Of Object) = ДогПодНом()

        Dim tbl = oWordDoc.Tables(2)
        tbl.Rows.AllowBreakAcrossPages = False
        'Dim row = tbl.Rows.Add()
        Dim allstoim As Double

        For x As Integer = 0 To Grid1.Rows.Count - 1 'заполняем таблицу по новому
            Dim row = tbl.Rows.Add()
            With row
                .Cells(1).Range.Text = Grid1.Rows(x).Cells(0).Value
                .Cells(2).Range.Text = Grid1.Rows(x).Cells(1).Value
                .Cells(3).Range.Text = Grid1.Rows(x).Cells(3).Value
                .Cells(4).Range.Text = Grid1.Rows(x).Cells(2).Value
                .Cells(5).Range.Text = Grid1.Rows(x).Cells(4).Value
            End With
            allstoim += CType(Replace(Grid1.Rows(x).Cells(4).Value, ".", ","), Double)
        Next

        With oWordDoc.Bookmarks
            .Item("АктПодр1").Range.Text = ComboBox5.Text
            .Item("АктПодр2").Range.Text = ДП(0)
            .Item("АктПодр3").Range.Text = ComboBox3.Text & " - " & Strings.Right(ДП(0), 4)
            .Item("АктПодр4").Range.Text = MaskedTextBox2.Text
            .Item("АктПодр5").Range.Text = MaskedTextBox3.Text
            If Организация(0).Item(1).ToString = "Индивидуальный предприниматель" Then
                .Item("АктПодр6").Range.Text = Организация(0).Item(1).ToString & " " & Организация(0).Item(0).ToString
                .Item("АктПодр21").Range.Text = Организация(0).Item(1).ToString & " " & Организация(0).Item(0).ToString
            Else
                .Item("АктПодр6").Range.Text = Организация(0).Item(1).ToString & " «" & Организация(0).Item(0).ToString & "» "
                .Item("АктПодр21").Range.Text = Организация(0).Item(1).ToString & " «" & Организация(0).Item(0).ToString & "» "
            End If
            If Организация(0).Item(0).ToString = "Итал Гэлэри Плюс" Then
                .Item("АктПодр7").Range.Text = ДобОконч(Организация(0).Item(18).ToString) & " " & Организация(0).Item(29).ToString
                .Item("АктПодр8").Range.Text = ""
            ElseIf Организация(0).Item(1).ToString = "Индивидуальный предприниматель" Then
                .Item("АктПодр7").Range.Text = Организация(0).Item(29).ToString
                .Item("АктПодр8").Range.Text = "действующего на основании " & Организация(0).Item(20).ToString
            Else
                .Item("АктПодр7").Range.Text = ДобОконч(Организация(0).Item(18).ToString) & " " & Организация(0).Item(29).ToString
                .Item("АктПодр8").Range.Text = "действующего на основании " & Организация(0).Item(20).ToString
            End If
            .Item("АктПодр9").Range.Text = ComboBox19.Text
            .Item("АктПодр10").Range.Text = "N " & ComboBox3.Text & " - " & Strings.Right(ДП(0), 4)
            .Item("АктПодр11").Range.Text = ДП(0)











            'For i As Integer = 0 To 0
            '    .Item("АктПодр12").Range.Text = ТекстРаботИменПадеж(a0(0)(0))
            '    .Item("АктПодр13").Range.Text = (a0(0)(3))
            '    .Item("АктПодр14").Range.Text = (a0(0)(2))
            '    .Item("АктПодр15").Range.Text = (a0(0)(4))
            '    .Item("АктПодр36").Range.Text = (a0(0)(1))
            '    If v = 1 Then Exit For
            '    .Item("АктПодр37").Range.Text = ТекстРаботИменПадеж(a0(1)(0))
            '    .Item("АктПодр38").Range.Text = (a0(1)(1))
            '    .Item("АктПодр39").Range.Text = (a0(1)(3))
            '    .Item("АктПодр40").Range.Text = (a0(1)(2))
            '    .Item("АктПодр41").Range.Text = (a0(1)(4))
            '    If v = 2 Then Exit For
            '    .Item("АктПодр42").Range.Text = ТекстРаботИменПадеж(a0(2)(0))
            '    .Item("АктПодр43").Range.Text = (a0(2)(1))
            '    .Item("АктПодр44").Range.Text = (a0(2)(3))
            '    .Item("АктПодр45").Range.Text = (a0(2)(2))
            '    .Item("АктПодр46").Range.Text = (a0(2)(4))
            '    If v = 3 Then Exit For
            '    .Item("АктПодр47").Range.Text = ТекстРаботИменПадеж(a0(3)(0))
            '    .Item("АктПодр48").Range.Text = (a0(3)(1))
            '    .Item("АктПодр49").Range.Text = (a0(3)(3))
            '    .Item("АктПодр50").Range.Text = (a0(3)(2))
            '    .Item("АктПодр51").Range.Text = (a0(3)(4))
            'Next

            'If v = 1 Then
            'allstoim = ДобРазрядности(TextBox9.Text)
            'ElseIf v = 2 Then
            '    If IsNumeric(TextBox2.Text) And IsNumeric(TextBox9.Text) Then
            '        allstoim = ДобРазрядности(CType(Math.Round(CDbl(TextBox9.Text) + CDbl(TextBox2.Text), 2), String))
            '    Else
            '        Throw New Exception("Не все строки заполены!")
            '    End If
            'ElseIf v = 3 Then
            '        allstoim = ДобРазрядности(CType(Math.Round(CDbl(TextBox9.Text) + CDbl(TextBox2.Text) + CDbl(TextBox12.Text), 2), String))
            '    Else
            '        allstoim = ДобРазрядности(CType(Math.Round(CDbl(TextBox9.Text) + CDbl(TextBox2.Text) + CDbl(TextBox12.Text) + CDbl(TextBox17.Text), 2), String))
            'End If



            allstoim = Math.Round(allstoim, 2)

            Dim allststring As String
            If (allstoim = Math.Truncate(allstoim)) Then
                allststring = CType(allstoim, String) & ",00"
            Else
                allststring = CType(allstoim, String)

            End If


            .Item("АктПодр16").Range.Text = allststring
            .Item("АктПодр17").Range.Text = ЧислоПрописДляСправки(allstoim)
            Dim mObj As Object = Подоходный(allstoim)
            .Item("АктПодр18").Range.Text = mObj(0) & " руб."
            .Item("АктПодр19").Range.Text = mObj(1) & " руб."
            .Item("АктПодр20").Range.Text = MaskedTextBox1.Text
            .Item("АктПодр22").Range.Text = Организация(0).Item(4).ToString
            .Item("АктПодр23").Range.Text = Организация(0).Item(2).ToString
            .Item("АктПодр24").Range.Text = Организация(0).Item(14).ToString
            .Item("АктПодр25").Range.Text = Организация(0).Item(12).ToString
            .Item("АктПодр26").Range.Text = Организация(0).Item(11).ToString
            If Организация(0).Item(31) = True And Not Организация(0).Item(1).ToString = "Индивидуальный предприниматель" Then
                .Item("АктПодр27").Range.Text = ФИОКорРук(Организация(0).Item(19).ToString, True)
            ElseIf Организация(0).Item(31) = True And Организация(0).Item(1).ToString = "Индивидуальный предприниматель" Then
                .Item("АктПодр27").Range.Text = ФИОКорРук(Организация(0).Item(19).ToString, True)
            ElseIf Организация(0).Item(31) = False And Организация(0).Item(1).ToString = "Индивидуальный предприниматель" Then
                .Item("АктПодр27").Range.Text = ФИОКорРук(Организация(0).Item(19).ToString, True)
            Else
                .Item("АктПодр27").Range.Text = ФИОКорРук(Организация(0).Item(19).ToString, False)
            End If

            .Item("АктПодр28").Range.Text = ComboBox19.Text
            .Item("АктПодр29").Range.Text = Сотрудник(0).Item(10).ToString & Сотрудник(0).Item(11).ToString
            .Item("АктПодр30").Range.Text = Сотрудник(0).Item(12).ToString
            .Item("АктПодр31").Range.Text = Сотрудник(0).Item(14).ToString
            .Item("АктПодр32").Range.Text = Сотрудник(0).Item(16).ToString
            .Item("АктПодр33").Range.Text = Сотрудник(0).Item(15).ToString
            .Item("АктПодр34").Range.Text = ФИОКорРук(ComboBox19.Text, False)
            .Item("АктПодр35").Range.Text = MaskedTextBox6.Text


        End With

        Dim Name As String = ComboBox5.Text & " " & ФИОКорРук(ComboBox19.Text, False) & " от " & MaskedTextBox6.Text & " (Акт договор подряда иное)(Договор № " & ComboBox3.Text & ")" & ".doc"
        Dim СохрЗак2 As New List(Of String)
        СохрЗак2.AddRange(New String() {ComboBox1.Text & "\Договор подряда\" & Now.Year & "\", Name})
        oWordDoc.SaveAs2(PathVremyanka & Name,,,,,, False)
        oWordDoc.Close(True)
        oWord.Quit(True)
        Конец(ComboBox1.Text & "\Договор подряда\" & Now.Year, Name, CType(Label96.Text, Integer), ComboBox1.Text, delstring, "Акт договор подряда иное")
        Dim massFTP3 As New ArrayList
        massFTP3.Add(СохрЗак2)
        Parallel.Invoke(Sub() RunMoving4())






        'If Not IO.Directory.Exists(OnePath & ComboBox1.Text & "\Договор подряда\" & Now.Year) Then
        '    IO.Directory.CreateDirectory(OnePath & ComboBox1.Text & "\Договор подряда\" & Now.Year)
        'End If

        'oWordDoc.SaveAs2("C:\Users\Public\Documents\Рик\" & TextBox4.Text & " " & ФИОКорРук(ComboBox19.Text, False) & " от " & MaskedTextBox6.Text & " (Акт договор подряда иное)" & ".doc",,,,,, False)

        ''oWordDoc.SaveAs2("U: \Офис\Финансовый\6. Бух.услуги\Кадры\" & Клиент & "\Заявление\" & Год & "\" & Заявление(9) & " (заявление)" & ".doc",,,,,, False)
        'Try
        '    IO.File.Copy("C:\Users\Public\Documents\Рик\" & TextBox4.Text & " " & ФИОКорРук(ComboBox19.Text, False) & " от " & MaskedTextBox6.Text & " (Акт договор подряда иное)" & ".doc", OnePath & ComboBox1.Text & "\Договор подряда\" & Now.Year & "\" & TextBox4.Text & " " & ФИОКорРук(ComboBox19.Text, False) & " от " & MaskedTextBox6.Text & " (Акт договор подряда)" & ".doc")
        'Catch ex As Exception
        '    If MessageBox.Show("Акт договора подряда с сотрудником " & ФИОКорРук(ComboBox19.Text, False) & " существует. Заменить старый документ новым?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then
        '        Try
        '            IO.File.Delete(OnePath & ComboBox1.Text & "\Договор подряда\" & Now.Year & "\" & TextBox4.Text & " " & ФИОКорРук(ComboBox19.Text, False) & " от " & MaskedTextBox6.Text & " (Акт договор подряда иное)" & ".doc")
        '        Catch ex1 As Exception
        '            MessageBox.Show("Закройте файл!", Рик)
        '        End Try


        '        IO.File.Copy("C:\Users\Public\Documents\Рик\" & TextBox4.Text & " " & ФИОКорРук(ComboBox19.Text, False) & " от " & MaskedTextBox6.Text & " (Акт договор подряда)" & ".doc", OnePath & ComboBox1.Text & "\Договор подряда\" & Now.Year & "\" & TextBox4.Text & " " & ФИОКорРук(ComboBox19.Text, False) & " от " & MaskedTextBox6.Text & " (Акт договор подряда)" & ".doc")
        '    End If
        'End Try
        'СохрЗак = OnePath & ComboBox1.Text & "\Договор подряда\" & Now.Year & "\" & TextBox4.Text & " " & ФИОКорРук(ComboBox19.Text, False) & " от " & MaskedTextBox6.Text & " (Акт договор подряда)" & ".doc"

        'oWordDoc.Close(True)
        'oWord.Quit(True)

        'Dim mass() As String
        If MessageBox.Show("Акт договора подряда с сотрудником " & vbCrLf & ФИОКорРук(ComboBox19.Text, False) & " сформирован успешно!" & vbCrLf & "Распечатать документ!", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.None) = DialogResult.OK Then
            ПечатьДоковFTP(massFTP3, 2)
        End If
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick
        If ListBox1.SelectedIndex = -1 Then
            MessageBox.Show("Выберите документ для просмотра!", Рик, MessageBoxButtons.OK)
            Exit Sub
        End If

        Dim l As String = ComboBox4.Items.Item(ListBox1.SelectedIndex)

        ВыгрузкаФайловНаЛокалыныйКомп(l, PathVremyanka & ListBox1.SelectedItem)

        Dim proc As Process = Process.Start(PathVremyanka & ListBox1.SelectedItem)
        proc.WaitForExit()
        proc.Close()

        ЗагрНаСерверИУдаление(PathVremyanka & ListBox1.SelectedItem, l, ListBox1.SelectedItem)
    End Sub

















    Private Sub MaskedTextBox2_TextChanged(sender As Object, e As EventArgs) Handles MaskedTextBox2.TextChanged
        If CheckBox2.Checked = True Then
            If MaskedTextBox2.Text <> "" Then
                MaskedTextBox3.Text = MaskedTextBox2.Text
                MaskedTextBox6.Text = MaskedTextBox2.Text
            End If
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            MaskedTextBox1.Enabled = True
        Else
            MaskedTextBox1.Enabled = False
        End If
    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            MaskedTextBox2.Focus()

            Dim pl As String
            If ComboBox5.Text <> "" Then
                Try
                    Dim i As Integer = CInt(ComboBox5.Text)
                    Select Case i
                        Case < 10
                            pl = Str(i)
                            ComboBox5.Text = "00" & i

                        Case 10 To 99
                            pl = Str(i)
                            ComboBox5.Text = "0" & i
                    End Select
                Catch ex As Exception

                End Try
            Else
                ComboBox5.Text = "бн"
            End If

        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        If ComboBox3.Text <> "" Then
            Очистка2()
            ОбщДанПоДог2()
        End If
    End Sub

    Private Sub ListBox1_DrawItem(sender As Object, e As DrawItemEventArgs) Handles ListBox1.DrawItem

        If e.Index < 0 Then Exit Sub

        e.DrawBackground()

        Dim bcolor As Color = If(Equals(e.BackColor, ListBox1.BackColor),
                                  ListBox1.BackColor,
                                  Color.LightGreen)




        e.Graphics.FillRectangle(New SolidBrush(bcolor), e.Bounds)

        e.Graphics.DrawString(ListBox1.Items.Item(e.Index).ToString,
                              ListBox1.Font,
                              New SolidBrush(ListBox1.ForeColor),
                              e.Bounds,
                              New StringFormat With {.Alignment = StringAlignment.Near,
                                                     .LineAlignment = StringAlignment.Center,
                                                     .Trimming = StringTrimming.None,
                                                     .FormatFlags = StringFormatFlags.NoWrap})

    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        If ComboBox3.Text = "" Then
            MessageBox.Show("Выберите номер договора!", Рик)
            Exit Sub
        End If




        ОбновлGrid()

        'Dim d = ds.Count

        'With Me
        '    .TextBox1.Text = ds(0).Item(11).ToString
        '    .TextBox6.Text = ds(0).Item(14).ToString
        '    If ds(0).Item(17).ToString = "" Then
        '        .TextBox7.Text = ds(0).Item(12).ToString & "," & ds(0).Item(13).ToString
        '    Else
        '        .TextBox7.Text = ds(0).Item(17).ToString
        '    End If
        '    .TextBox8.Text = ds(0).Item(15).ToString
        '    .TextBox9.Text = ds(0).Item(18).ToString

        '    If d = 1 Then Exit Sub
        '    GroupBox4.Enabled = True
        '    .TextBox11.Text = ds(1).Item(11).ToString
        '    .TextBox10.Text = ds(1).Item(14).ToString
        '    If ds(1).Item(17).ToString = "" Then
        '        .TextBox5.Text = ds(1).Item(12).ToString & "," & ds(1).Item(13).ToString
        '    Else
        '        .TextBox5.Text = ds(1).Item(17).ToString
        '    End If
        '    .TextBox3.Text = ds(1).Item(15).ToString
        '    .TextBox2.Text = ds(1).Item(18).ToString

        '    If d = 2 Then Exit Sub
        '    GroupBox5.Enabled = True
        '    .TextBox16.Text = ds(2).Item(11).ToString
        '    .TextBox15.Text = ds(2).Item(14).ToString
        '    If ds(2).Item(17).ToString = "" Then
        '        .TextBox14.Text = ds(2).Item(12).ToString & "," & ds(2).Item(13).ToString
        '    Else
        '        .TextBox14.Text = ds(2).Item(17).ToString
        '    End If
        '    .TextBox13.Text = ds(2).Item(15).ToString
        '    .TextBox12.Text = ds(2).Item(18).ToString

        '    If d = 3 Then Exit Sub
        '    GroupBox6.Enabled = True
        '    .TextBox21.Text = ds(3).Item(11).ToString
        '    .TextBox20.Text = ds(3).Item(14).ToString
        '    If ds(3).Item(17).ToString = "" Then
        '        .TextBox19.Text = ds(3).Item(12).ToString & "," & ds(3).Item(13).ToString
        '    Else
        '        .TextBox19.Text = ds(3).Item(17).ToString
        '    End If
        '    .TextBox18.Text = ds(3).Item(15).ToString
        '    .TextBox17.Text = ds(3).Item(18).ToString

        'End With


    End Sub
    Private Sub ОбновлGrid()
        Dim j = CType(Label96.Text, Integer)
        Dim ds1

        ds1 = dtDogovorPadriadaAll.Select("ID=" & j & " And НомерДогПодр=" & ComboBox3.Text & " And ПорНомерАктаИное=" & ComboBox5.Text & "")





        Dim ds = From x In dtDogovorPadriadaAll.AsEnumerable() Where Not x.IsNull("ID") AndAlso x.Item("ID") = j _
                                                               AndAlso Not x.IsNull("НомерДогПодр") AndAlso x.Item("НомерДогПодр") = ComboBox3.Text _
                                                               AndAlso Not x.IsNull("ПорНомерАктаИное") AndAlso x.Item("ПорНомерАктаИное") = ComboBox5.Text
                 Select New With {.Наименование = x.Item("ВыпРаб1"), .Единица = x.Item("ЕдИзмерАктИное"), .Стоимость2 = x.Item("СтоимЕдРаботыАктИное"),
                     .Объем = x.Item("ОбъемВыпРаботАктИное"), .Стоимость = x.Item("ОбщСтоимРаботАктИное"), .Код = x.Item("Код")}



        If ds1(0).Item(19).ToString.Length = 1 Then
            ComboBox5.Text = "00" & ds1(0).Item(19).ToString
        ElseIf ds1(0).Item(19).ToString.Length = 2 Then
            ComboBox5.Text = "0" & ds1(0).Item(19).ToString

        Else
            ComboBox5.Text = ds1(0).Item(19).ToString
        End If


        Grid1.DataSource = ds.ToList()
        Grid1.Columns(1).HeaderCell.Value = "Единица измерения"
        Grid1.Columns(2).HeaderCell.Value = "Цена"
        Grid1.Columns(3).HeaderCell.Value = "Объем работ"
        Grid1.Columns("Код").Visible = False
        GridView(Grid1)




        MaskedTextBox1.Text = ds1(0).Item(23).ToString
        MaskedTextBox2.Text = ds1(0).Item(20).ToString
        MaskedTextBox3.Text = ds1(0).Item(21).ToString
        MaskedTextBox6.Text = ds1(0).Item(22).ToString
    End Sub

    Private Sub Grid1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellClick
        RichTextBox1.Text = Grid1.CurrentRow.Cells.Item(0).Value
        TextBox4.Text = Grid1.CurrentRow.Cells.Item(1).Value
        TextBox22.Text = Grid1.CurrentRow.Cells.Item(2).Value
        TextBox23.Text = Grid1.CurrentRow.Cells.Item(3).Value
        TextBox24.Text = Grid1.CurrentRow.Cells.Item(4).Value
        Код = Grid1.CurrentRow.Cells.Item(5).Value

    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox23_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox23.KeyDown
        If e.KeyValue = Keys.Enter Then
            e.SuppressKeyPress = True
            Dim f = Replace(TextBox23.Text, ".", ",")


            If TextBox22.Text <> "" And IsNumeric(TextBox22.Text) And IsNumeric(f) Then
                TextBox24.Text = CType(Math.Round(CDbl(f) * CDbl(TextBox22.Text), 2), String)
                TextBox24.Text = ДобРазрядности(TextBox24.Text)
                TextBox23.Text = ДобРазрядности(f)
            End If
        End If

    End Sub

    Private Sub ЗаполнДанн(ByVal ds As DataTable)

        КодАкт = Nothing
        КодАкт = ds.Rows(0).Item(7)
        MaskedTextBox2.Text = ds.Rows(0).Item(2).ToString
        If ds.Rows(0).Item(1).ToString.Length = 1 Then
            ComboBox5.Text = "00" & ds.Rows(0).Item(1).ToString
        ElseIf ds.Rows(0).Item(1).ToString.Length = 2 Then
            ComboBox5.Text = "0" & ds.Rows(0).Item(1).ToString
        Else
            ComboBox5.Text = ds.Rows(0).Item(1).ToString
        End If

        MaskedTextBox3.Text = ds.Rows(0).Item(3).ToString
        'TextBox1.Text = ds.Rows(0).Item(4).ToString
        'TextBox2.Text = ds.Rows(0).Item(5).ToString
        'TextBox3.Text = ds.Rows(0).Item(6).ToString
        MaskedTextBox6.Text = ds.Rows(0).Item(8).ToString
        MaskedTextBox1.Text = ds.Rows(0).Item(9).ToString
        ПослНомАкт = Nothing
        ПослНомАкт = ds.Rows(0).Item(1)

    End Sub

    Private Sub TextBox24_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            If Not IsNumeric(TextBox24.Text) And Not IsNumeric(TextBox23.Text) Then Exit Sub
            Dim f As Double
            f = Replace(TextBox24.Text, ".", ",")
            If СправкаПоЗарплате.bool(f) = True Then
                TextBox24.Text = f & ",00"
            Else
                TextBox24.Text = f
                If СправкаПоЗарплате.Count(f) = 1 Then
                    TextBox24.Text = f & "0"
                End If
            End If

            Dim txt7 As Double = Replace(TextBox22.Text, ".", ",")
            Dim txt9 As Double = Replace(TextBox24.Text, ".", ",")

            If IsNumeric(txt7) And IsNumeric(txt9) Then
                TextBox23.Text = CType(Math.Round(txt9 / txt7, 2), String)
                Dim fd As Double
                fd = Replace(TextBox23.Text, ".", ",")
                If СправкаПоЗарплате.bool(fd) = True Then
                    TextBox23.Text = fd & ",00"
                Else
                    TextBox23.Text = fd
                    If СправкаПоЗарплате.Count(fd) = 1 Then
                        TextBox23.Text = fd & "0"
                    End If
                End If

            End If

        End If
    End Sub

    Private Sub ВычДатВыплат(ByVal dt As Date)
        Dim int As Integer = dt.Month
        Dim год As Integer = dt.Year

        If int = 12 Then
            Dim dp As Date = CDate("01." & "01." & CType(год + 1, String))
            MaskedTextBox1.Text = dp.AddDays(10)
        Else
            int = int + 1
            Dim st As String = CType(int, String)
            If int < 10 Then
                st = CType("0" & int, String)
            End If
            Dim dp As Date = CDate("01." & st & "." & CType(год, String))
            MaskedTextBox1.Text = dp.AddDays(9)
        End If


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If IsNumeric(TextBox23.Text) = False Then
            MessageBox.Show("Введите числовое значение!", Рик)
            Exit Sub
        End If

        If TextBox4.Text = "" Or TextBox22.Text = "" Then
            MessageBox.Show("Выберите обьект для изменения!", Рик)
            Exit Sub
        End If



        Dim list As New Dictionary(Of String, Object)
        list.Add("@Код", Код)
        list.Add("@ОбъемВыпРаботАктИное", TextBox23.Text)
        list.Add("@ОбщСтоимРаботАктИное", TextBox24.Text)

        'list.Add("@ЗаПериодСИное", MaskedTextBox2.Text)
        'list.Add("@ЗаПериодПоИное", MaskedTextBox3.Text)
        'list.Add("@ДатаАктаИное", MaskedTextBox6.Text)
        'list.Add("@ДатаОплатыРаботИное", MaskedTextBox1.Text)


        Updates(stroka:="UPDATE ДогПодряда SET ОбъемВыпРаботАктИное=@ОбъемВыпРаботАктИное, ОбщСтоимРаботАктИное=@ОбщСтоимРаботАктИное 
WHERE Код=@Код", list)

        'ЗаПериодСИное =@ЗаПериодСИное, ЗаПериодПоИное=@ЗаПериодПоИное, ДатаАктаИное=@ДатаАктаИное, ДатаОплатыРаботИное=@ДатаОплатыРаботИное
        Parallel.Invoke(Sub() RunMoving9())
        MessageBox.Show("Данные договора подряда изменены!", Рик)

        ОбновлGrid()
    End Sub

    Private Sub TextBox23_LostFocus(sender As Object, e As EventArgs) Handles TextBox23.LostFocus
        Dim f = Replace(TextBox23.Text, ".", ",")
        If TextBox22.Text <> "" And IsNumeric(TextBox22.Text) And IsNumeric(f) Then
            TextBox24.Text = CType(Math.Round(CDbl(f) * CDbl(TextBox22.Text), 2), String)
            TextBox24.Text = ДобРазрядности(TextBox24.Text)
            TextBox23.Text = ДобРазрядности(f)
        End If
    End Sub

    Private Sub TextBox24_LostFocus(sender As Object, e As EventArgs) Handles TextBox24.LostFocus

        'If Not IsNumeric(TextBox24.Text) And Not IsNumeric(TextBox23.Text) Then Exit Sub
        '    Dim f As Double
        '    f = Replace(TextBox24.Text, ".", ",")
        '    If СправкаПоЗарплате.bool(f) = True Then
        '        TextBox24.Text = f & ",00"
        '    Else
        '        TextBox24.Text = f
        '        If СправкаПоЗарплате.Count(f) = 1 Then
        '            TextBox24.Text = f & "0"
        '        End If
        '    End If

        '    Dim txt7 As Double = Replace(TextBox22.Text, ".", ",")
        '    Dim txt9 As Double = Replace(TextBox24.Text, ".", ",")

        '    If IsNumeric(txt7) And IsNumeric(txt9) Then
        '        TextBox23.Text = CType(Math.Round(txt9 / txt7, 2), String)
        '        Dim fd As Double
        '        fd = Replace(TextBox23.Text, ".", ",")
        '        If СправкаПоЗарплате.bool(fd) = True Then
        '            TextBox23.Text = fd & ",00"
        '        Else
        '            TextBox23.Text = fd
        '            If СправкаПоЗарплате.Count(fd) = 1 Then
        '                TextBox23.Text = fd & "0"
        '            End If
        '        End If

        '    End If


    End Sub

    Private Sub ComboBox5_LostFocus(sender As Object, e As EventArgs) Handles ComboBox5.LostFocus
        Dim dt As New List(Of String)
        If ComboBox5.Text = "" Then Exit Sub
        For Each x In datrow.Rows
            If Not x.Item("ПорНомерАктаИное").ToString = ComboBox5.Text Then
                Select Case ComboBox5.Text.Length
                    Case 1
                        ComboBox5.Text = "00" & ComboBox5.Text
                        Exit For
                    Case 2
                        ComboBox5.Text = "0" & ComboBox5.Text
                        Exit For
                End Select
            End If
        Next
    End Sub
End Class
