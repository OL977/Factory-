Option Explicit On
Imports System.Data.OleDb
Public Class ДогПодрядаАктИноеСоздать
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
    Dim datrow As DataRow()
    Dim DtGr2 As New DataTable
    Dim ls As Boolean = False
    Dim _index As Integer

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


    Private Sub ДогПодрядаАктИноеСоздать_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        TextBox1.Text = ""
        ListBox1.Items.Clear()
        Dim dt As New DataTable
        Grid1.DataSource = dt
        Grid2.DataSource = dt

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

        Dim ds2 = From x In dtSotrudnikiAll Where x.Item("ФИОСборное") = ComboBox19.Text Select x.Item("Фамилия")
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
                                                 And x.Item("ИмяФайла").ToString.Contains(name) Select x

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
            If MaskedTextBox3.MaskCompleted = True And MaskedTextBox2.MaskCompleted = True And MaskedTextBox6.MaskCompleted = True Then
                ВычДатВыплат(MaskedTextBox6.Text)
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


    Private Sub ОбщДанПоДог2()

        Dim j = CType(Label96.Text, Integer)
        'Dim ds = dtDogovorPadriadaAll.Select("ID=" & j & " And НомерДогПодр=" & ComboBox3.Text & "")
        ОбновлGrid()
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

        ОбновлGrid()

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
        If ComboBox3.Text = "" Then
            MessageBox.Show("Выберите номер договора!", Рик)
            Return 1
        End If

        If TextBox1.Text = "" Or Not IsNumeric(TextBox1.Text) Then
            MessageBox.Show("Заполните номер акта!", Рик)
            Return 1
        End If

        If Not IsNumeric(TextBox1.Text) Then
            MessageBox.Show("Поле номер акта должно быть целочисленным!", Рик)
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

        If IsDBNull(DtGr2) Or DtGr2.Rows.Count = 0 Then
            MessageBox.Show("Сформируйте список работ для добавления в акт!", Рик)
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
        Dim df = From x In dtDogovorPadriadaAll Where Not x.IsNull("ID") AndAlso x.Item("ID") = CType(Label96.Text, Integer) _
                                                AndAlso Not x.IsNull("НомерДогПодр") AndAlso x.Item("НомерДогПодр").ToString = ComboBox3.Text Select x

        'Dim strsql As String = "SELECT ПорНомерАктаИное FROM ДогПодряда WHERE ID=" & CType(Label96.Text, Integer) & " and НомерДогПодр='" & ComboBox3.Text & "' AND ПорНомерАктаИное='" & TextBox4.Text & "'"
        'Dim df As DataTable = Selects(strsql)

        If df.Count > 0 Then
            If MessageBox.Show("Заменить старые данные акта №" & CType(TextBox1.Text, Integer) & " -новыми?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                Return 1
            End If
            НадоБновл = True
        End If
        Return 0
    End Function
    Private Sub НовСтрока(ByVal strsql2 As String, ByVal код As Integer)

        'Dim c As New OleDbCommand
        'c.Connection = conn
        'c.CommandText = strsql2
        'Dim ds As New DataSet
        'Dim da As New OleDbDataAdapter(c)
        'da.Fill(ds, "Сохранение")

        'Dim cb As New OleDbCommandBuilder(da)
        'Dim dsNewRow As DataRow

        'dsNewRow = ds.Tables("Сохранение").NewRow()
        'dsNewRow.Item("IDДогПодр") = код
        'dsNewRow.Item("ПорНомерАкта") = ComboBox5.Text
        'dsNewRow.Item("ЗаПериодС") = MaskedTextBox2.Text
        'dsNewRow.Item("ЗаПериодПо") = MaskedTextBox3.Text
        ''dsNewRow.Item("ВремяРабот") = TextBox1.Text
        ''dsNewRow.Item("СтоимЧаса") = TextBox2.Text
        ''dsNewRow.Item("СтоимРабот") = TextBox3.Text
        'dsNewRow.Item("ДатаАкта") = MaskedTextBox6.Text
        'dsNewRow.Item("ДатаОплатыРабот") = MaskedTextBox1.Text
        'ds.Tables("Сохранение").Rows.Add(dsNewRow)

        ''ds.Tables("Сохранение").Rows(0).Item(0) = a
        ''ds.Tables("Сохранение").Rows(0).Item(1) = Me.TextBox1.Text
        'da.Update(ds, "Сохранение")

    End Sub
    Private Sub СохрВБазу()

        Dim b = From x In dtDogovorPadriadaAll.AsEnumerable Where Not x.IsNull("ID") AndAlso x.Item("ID") = CType(Label96.Text, Integer) _
                                                            AndAlso Not x.IsNull("НомерДогПодр") AndAlso x.Item("НомерДогПодр") = ComboBox3.Text
                Select x.Item("Код")
        Dim b1 As Integer = CType(b(0).ToString, Integer)

        Dim objlist As New ArrayList()

        Dim list As New Dictionary(Of String, Object)
        list.Add("@IDДогПодряда", b1)
        list.Add("@ПорНомерАктаИное", TextBox1.Text)
        list.Add("@ЗаПериодСИное", MaskedTextBox2.Text)
        list.Add("@ЗаПериодПоИное", MaskedTextBox3.Text)
        list.Add("@ДатаАктаИное", MaskedTextBox6.Text)
        list.Add("@ДатаОплатыРаботИное", MaskedTextBox1.Text)

        For Each r As DataRow In DtGr2.Rows
            list.Add("@ОбъемВыпРаботАктИное", r.Item("Объем работ").ToString)
            list.Add("@ЕдИзмерАктИное", r.Item("Единица измерения").ToString)
            list.Add("@СтоимЕдРаботыАктИное", r.Item("Цена").ToString)
            list.Add("@ОбщСтоимРаботАктИное", r.Item("Стоимость").ToString)
            list.Add("@ВыпРаб1", r.Item("Наименование").ToString)

            '            Updates(stroka:="INSERT INTO ДогПодряда(ОбъемВыпРаботАктИное,ЕдИзмерАктИное,СтоимЕдРаботыАктИное,
            'ОбщСтоимРаботАктИное, ЗаПериодСИное, ЗаПериодПоИное, ДатаАктаИное, ДатаОплатыРаботИное, ПорНомерАктаИное, iD, НомерДогПодр,ВыпРаб1)
            'VALUES(@ОбъемВыпРаботАктИное,@ЕдИзмерАктИное,@СтоимЕдРаботыАктИное,@ОбщСтоимРаботАктИное,@ЗаПериодСИное, @ЗаПериодПоИное,
            '@ДатаАктаИное,@ДатаОплатыРаботИное,@ПорНомерАктаИное,@ID,@НомерДогПодр,@ВыпРаб1)", list, "ДогПодряда")

            Updates(stroka:="INSERT INTO ДогПодрядаАктИное(
ОбъемВыпРаботАктИное,
ЕдИзмерАктИное,
СтоимЕдРаботыАктИное,
ОбщСтоимРаботАктИное,
ЗаПериодСИное,
ЗаПериодПоИное,
ДатаАктаИное,
ДатаОплатыРаботИное,
ПорНомерАктаИное,
IDДогПодряда,
ВыпРаб1)

VALUES(
@ОбъемВыпРаботАктИное,
@ЕдИзмерАктИное,
@СтоимЕдРаботыАктИное,
@ОбщСтоимРаботАктИное,
@ЗаПериодСИное,
@ЗаПериодПоИное,
@ДатаАктаИное,
@ДатаОплатыРаботИное,
@ПорНомерАктаИное,
@IDДогПодряда,
@ВыпРаб1)", list, "ДогПодряда")

            list.Remove("@ОбъемВыпРаботАктИное")
            list.Remove("@ЕдИзмерАктИное")
            list.Remove("@СтоимЕдРаботыАктИное")
            list.Remove("@ОбщСтоимРаботАктИное")
            list.Remove("@ВыпРаб1")

        Next




        '        Updates(stroka:="UPDATE ДогПодряда SET ЗаПериодСИное='" & MaskedTextBox2.Text & "',ЗаПериодПоИное='" & MaskedTextBox3.Text & "',
        'ДатаАктаИное='" & MaskedTextBox6.Text & "',ДатаОплатыРаботИное='" & MaskedTextBox1.Text & "'
        'WHERE ID=@ID AND НомерДогПодр=@НомерДогПодр AND ПорНомерАктаИное=@ПорНомерАктаИное", list, "ДогПодряда")


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

        If TextBox1.Text.Length = 1 Then
            TextBox1.Text = "00" & TextBox1.Text
        ElseIf TextBox1.Text.Length = 2 Then
            TextBox1.Text = "0" & TextBox1.Text
        End If

        Dim f = From x In ListBox1.Items Select Strings.Left(x.ToString, 3)

        For Each r In f
            If r = TextBox1.Text Then
                MessageBox.Show("Такой номер акта уже существует!" & vbCrLf & "Выбериет другой", Рик)
                Exit Sub
            End If
        Next




        СохрВБазу()
        If CheckBox4.Checked = False Then
            Доки()
        End If
        Очистка()
        ComboBox19.Text = ""
        RichTextBox1.Text = ""
        CheckBox2.Checked = False
        Com1Sel9()
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


        Начало("ActPodriadaInoe5.doc")
        oWordDoc = oWord.Documents.Add(firthtPath & "\ActPodriadaInoe5.doc")

        Dim Организация = Org(ComboBox1.Text)
        Dim Сотрудник = Sotrudnic(CType(Label96.Text, Integer))

        Dim ДП As IEnumerable(Of Object) = ДогПодНом()

        Dim tbl = oWordDoc.Tables(2)
        tbl.Rows.AllowBreakAcrossPages = False
        'Dim row = tbl.Rows.Add()
        Dim allstoim As Double

        For x As Integer = 0 To Grid2.Rows.Count - 1 'заполняем таблицу по новому
            Dim row = tbl.Rows.Add()
            With row
                .Cells(1).Range.Text = Grid2.Rows(x).Cells(0).Value
                .Cells(2).Range.Text = Grid2.Rows(x).Cells(1).Value
                .Cells(3).Range.Text = Grid2.Rows(x).Cells(3).Value
                .Cells(4).Range.Text = Grid2.Rows(x).Cells(2).Value
                .Cells(5).Range.Text = Grid2.Rows(x).Cells(4).Value
            End With
            allstoim += CType(Replace(Grid2.Rows(x).Cells(4).Value, ".", ","), Double)
        Next

        With oWordDoc.Bookmarks
            .Item("АктПодр1").Range.Text = TextBox1.Text
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

        Dim Name As String = TextBox1.Text & " " & ФИОКорРук(ComboBox19.Text, False) & " от " & MaskedTextBox6.Text & " (Акт договор подряда иное)(Договор № " & ComboBox3.Text & ")" & ".doc"
        Dim СохрЗак2 As New List(Of String)
        СохрЗак2.AddRange(New String() {ComboBox1.Text & "\Договор подряда\" & Now.Year & "\", Name})
        oWordDoc.SaveAs2(PathVremyanka & Name,,,,,, False)
        oWordDoc.Close(True)
        oWord.Quit(True)
        Конец(ComboBox1.Text & "\Договор подряда\" & Now.Year, Name, CType(Label96.Text, Integer), ComboBox1.Text, delstring, "Акт договор подряда иное")
        Dim massFTP3 As New ArrayList
        massFTP3.Add(СохрЗак2)
        Parallel.Invoke(Sub() RunMoving4())


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

    Private Sub ОбновлGrid()
        Dim j = CType(Label96.Text, Integer)
        'Dim ds1

        'ds1 = dtDogovorPadriadaAll.Select("ID=" & j & " And НомерДогПодр=" & ComboBox3.Text & "")

        'Dim ds = (From x In dtDogovorPadriadaAll.AsEnumerable() Where Not x.IsNull("ID") AndAlso x.Item("ID") = j _
        '                                                       AndAlso Not x.IsNull("НомерДогПодр") AndAlso x.Item("НомерДогПодр") = ComboBox3.Text
        '          Select New With {.Наименование = x.Item("ВыпРаб1"), .Единица = x.Item("ЕдИзмерАктИное"), .Стоимость2 = x.Item("СтоимЕдРаботыАктИное"),
        '             .Объем = x.Item("ОбъемВыпРаботАктИное"), .Стоимость = x.Item("ОбщСтоимРаботАктИное"), .Код = x.Item("Код")}).ToList


        Dim ds = (From x In dtDogovorPadriadaAll.AsEnumerable()
                  Join y In dtDogPodrRabotyInoeAll.AsEnumerable On x.Field(Of Integer)("Код") Equals
                     y.Field(Of Integer)("IDДогПодряда")
                  Where Not x.IsNull("ID") AndAlso x.Item("ID") = j AndAlso Not x.IsNull("НомерДогПодр") AndAlso x.Item("НомерДогПодр") = ComboBox3.Text
                  Select New With
                     {.Наименование = y.Item("ВыпРаб1"),
                     .Единица = y.Item("ВидИзм"),
                     .Стоимость = (y.Field(Of String)("СтоимРуб1") & "," & y.Field(Of String)("СтоимКоп1")),
                      .Код = x.Item("Код"),
                      .IDtbRabotInoe = y.Item("ID")}).ToList

        '.Стоимость2 = x.Item("СтоимЕдРаботыАктИное"),
        '.Объем = x.Item("ОбъемВыпРаботАктИное"),



        'If ds1(0).Item(19).ToString.Length = 1 Then
        '    ComboBox5.Text = "00" & ds1(0).Item(19).ToString
        'ElseIf ds1(0).Item(19).ToString.Length = 2 Then
        '    ComboBox5.Text = "0" & ds1(0).Item(19).ToString

        'Else
        '    ComboBox5.Text = ds1(0).Item(19).ToString
        'End If


        Grid1.DataSource = ds
        Grid1.Columns(1).HeaderCell.Value = "Единица измерения"
        Grid1.Columns(2).HeaderCell.Value = "Цена"
        Grid1.Columns("Код").Visible = False
        Grid1.Columns("IDtbRabotInoe").Visible = False
        GridView(Grid1)



        If Not IsDBNull(DtGr2) Then
            DtGr2.Clear()
        End If

        Try
            DtGr2.Columns.Add("Наименование")
            DtGr2.Columns.Add("Единица измерения")
            DtGr2.Columns.Add("Цена") '8
            DtGr2.Columns.Add("Объем работ")
            DtGr2.Columns.Add("Стоимость")
            DtGr2.Columns.Add("Код")
        Catch ex As Exception

        End Try


        Grid2.DataSource = DtGr2
        Grid2.Columns("Код").Visible = False
        GridViewRed(Grid2)
        'grid2Создание(Grid1)


    End Sub
    Private Sub grid2Создание(ByVal d As DataGridView)

        Grid2.Columns.Add("С1", "Наименование")
        Grid2.Columns.Add("С2", "Единица измерения")
        Grid2.Columns.Add("С3", "Цена")
        Grid2.Columns.Add("С4", "Объем работ")
        Grid2.Columns.Add("С5", "Стоимость")
        Grid2.Columns.Add("С6", "Код")
        Grid2.Columns("С6").Visible = False
        GridView(Grid2)


        'Grid2.DataSource.rows.clear()
        'For x As Integer = 1 To Grid2.Rows.Count
        '    Grid2.Rows.RemoveAt(x)
        'Next

    End Sub
    Private Sub TextBox23_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox23.KeyDown
        'If e.KeyValue = Keys.Enter Then
        '    e.SuppressKeyPress = True
        '    Dim f = Replace(TextBox23.Text, ".", ",")


        '    If TextBox22.Text <> "" And IsNumeric(TextBox22.Text) And IsNumeric(f) Then
        '        TextBox24.Text = CType(Math.Round(CDbl(f) * CDbl(TextBox22.Text), 2), String)
        TextBox24.Text = ДобРазрядности(TextBox24.Text)
        '        TextBox23.Text = ДобРазрядности(f)
        '    End If
        'End If

    End Sub

    Private Sub ЗаполнДанн(ByVal ds As DataTable)

        '    КодАкт = Nothing
        '    КодАкт = ds.Rows(0).Item(7)
        '    MaskedTextBox2.Text = ds.Rows(0).Item(2).ToString
        '    If ds.Rows(0).Item(1).ToString.Length = 1 Then
        '        ComboBox5.Text = "00" & ds.Rows(0).Item(1).ToString
        '    ElseIf ds.Rows(0).Item(1).ToString.Length = 2 Then
        '        ComboBox5.Text = "0" & ds.Rows(0).Item(1).ToString
        '    Else
        '        ComboBox5.Text = ds.Rows(0).Item(1).ToString
        '    End If

        '    MaskedTextBox3.Text = ds.Rows(0).Item(3).ToString
        '    'TextBox1.Text = ds.Rows(0).Item(4).ToString
        '    'TextBox2.Text = ds.Rows(0).Item(5).ToString
        '    'TextBox3.Text = ds.Rows(0).Item(6).ToString
        '    MaskedTextBox6.Text = ds.Rows(0).Item(8).ToString
        '    MaskedTextBox1.Text = ds.Rows(0).Item(9).ToString
        '    ПослНомАкт = Nothing
        '    ПослНомАкт = ds.Rows(0).Item(1)

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
    Private Sub btn2Click()
        DtGr2.Rows(_index).Item("Цена") = TextBox22.Text
        DtGr2.Rows(_index).Item("Объем работ") = TextBox23.Text
        DtGr2.Rows(_index).Item("Стоимость") = TextBox24.Text
        Grid2.DataSource = DtGr2
        Grid2.Columns("Код").Visible = False
        GridViewRed(Grid2)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If MessageBox.Show("Сохранить изменения?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
            btn2Click()
        End If



        '        If IsNumeric(TextBox23.Text) = False Then
        '            MessageBox.Show("Введите числовое значение!", Рик)
        '            Exit Sub
        '        End If

        '        If TextBox4.Text = "" Or TextBox22.Text = "" Then
        '            MessageBox.Show("Выберите обьект для изменения!", Рик)
        '            Exit Sub
        '        End If



        '        Dim list As New Dictionary(Of String, Object)
        '        list.Add("@Код", Код)
        '        list.Add("@ОбъемВыпРаботАктИное", TextBox23.Text)
        '        list.Add("@ОбщСтоимРаботАктИное", TextBox24.Text)

        '        'list.Add("@ЗаПериодСИное", MaskedTextBox2.Text)
        '        'list.Add("@ЗаПериодПоИное", MaskedTextBox3.Text)
        '        'list.Add("@ДатаАктаИное", MaskedTextBox6.Text)
        '        'list.Add("@ДатаОплатыРаботИное", MaskedTextBox1.Text)


        '        Updates(stroka:="UPDATE ДогПодряда SET ОбъемВыпРаботАктИное=@ОбъемВыпРаботАктИное, ОбщСтоимРаботАктИное=@ОбщСтоимРаботАктИное 
        'WHERE Код=@Код", list)

        '        'ЗаПериодСИное =@ЗаПериодСИное, ЗаПериодПоИное=@ЗаПериодПоИное, ДатаАктаИное=@ДатаАктаИное, ДатаОплатыРаботИное=@ДатаОплатыРаботИное
        '        Parallel.Invoke(Sub() RunMoving9())
        '        MessageBox.Show("Данные договора подряда изменены!", Рик)

        '        ОбновлGrid()
    End Sub

    Private Sub TextBox23_LostFocus(sender As Object, e As EventArgs) Handles TextBox23.LostFocus
        Dim f = Replace(TextBox23.Text, ".", ",")
        If TextBox22.Text <> "" And IsNumeric(TextBox22.Text) And IsNumeric(f) Then
            TextBox24.Text = CType(Math.Round(CDbl(f) * CDbl(TextBox22.Text), 2), String)
            TextBox24.Text = ДобРазрядности(TextBox24.Text)
            TextBox23.Text = ДобРазрядности(f)
        End If
    End Sub

    Private Sub Grid2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid2.CellClick
        RichTextBox1.Text = Grid2.CurrentRow.Cells.Item(0).Value
        TextBox4.Text = Grid2.CurrentRow.Cells.Item(1).Value
        TextBox22.Text = Grid2.CurrentRow.Cells.Item(2).Value
        TextBox23.Text = Grid2.CurrentRow.Cells.Item(3).Value
        TextBox24.Text = Grid2.CurrentRow.Cells.Item(4).Value
        Код = Grid2.CurrentRow.Cells.Item(5).Value
        Button2.Enabled = False
        _index = Grid2.CurrentRow.Index
    End Sub

    Private Sub ПереносДанных()
        If Grid1.Rows.Count = 0 Then Exit Sub


        Dim s As Integer = Grid1.CurrentRow.Cells(4).Value

        If Grid2.Rows.Count = 0 Then
            Dim rw = dtDogPodrRabotyInoeAll.Select("ID=" & Grid1.CurrentRow.Cells(4).Value & "")

            Dim row As DataRow = DtGr2.NewRow
            row("Наименование") = rw(0).Item("ВыпРаб1").ToString
            row("Единица измерения") = rw(0).Item("ВидИзм").ToString
            row("Цена") = rw(0).Item("СтоимРуб1").ToString & "," & rw(0).Item("СтоимКоп1").ToString
            row("Объем работ") = ""
            row("Стоимость") = ""
            row("Код") = rw(0).Item("ID").ToString
            DtGr2.Rows.Add(row)



            Grid2.DataSource = DtGr2
            Grid2.Columns("Код").Visible = False
            GridViewRed(Grid2)
        Else


            Dim g = From x In DtGr2 Select x.Item("Код")


            If Not g.Contains(CType(s, String)) And Grid1.Rows.Count > DtGr2.Rows.Count Then

                Dim rw = dtDogPodrRabotyInoeAll.Select("ID=" & s & "")
                Dim row As DataRow = DtGr2.NewRow

                row("Наименование") = rw(0).Item("ВыпРаб1").ToString
                row("Единица измерения") = rw(0).Item("ВидИзм").ToString
                row("Цена") = rw(0).Item("СтоимРуб1").ToString & "," & rw(0).Item("СтоимКоп1").ToString
                row("Объем работ") = ""
                row("Стоимость") = ""
                row("Код") = rw(0).Item("ID").ToString
                DtGr2.Rows.Add(row)



                Grid2.DataSource = DtGr2
                Grid2.Columns("Код").Visible = False
                GridViewRed(Grid2)
            End If


        End If
    End Sub

    Private Sub Grid1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellDoubleClick
        If ls = True Then
            ПереносДанных()
        End If

    End Sub
    Private Sub grid2Вставка(ByVal d As DataGridViewRow)

        Dim f As New DataGridViewRow
        f.Cells(0).Value = d.Cells(0).Value
        f.Cells(1).Value = d.Cells(1).Value
        f.Cells(2).Value = d.Cells(2).Value
        f.Cells(3).Value = d.Cells(3).Value
        f.Cells(4).Value = d.Cells(4).Value
        f.Cells(5).Value = d.Cells(5).Value
        Grid2.Rows.Add(f)

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        If Grid2.Rows.Count = 0 Then
            Exit Sub
        End If

        If Код = Nothing Then
            MessageBox.Show("Выберите объект для удаления!", Рик)
            Exit Sub
        End If

        If MessageBox.Show("Удалить '" & Grid2.CurrentRow.Cells.Item("Наименование").Value & "'" & vbCrLf & "из списка отобранных работ?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then

            DtGr2.Rows.RemoveAt(Grid2.CurrentRow.Index)
            Grid2.DataSource = DtGr2
            GridViewRed(Grid2)
            Код = 0
            RichTextBox1.Text = ""
            TextBox4.Text = ""
            TextBox22.Text = ""
            TextBox23.Text = ""
            TextBox24.Text = ""


        End If




    End Sub

    Private Sub Grid1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellClick
        ls = False
        If e.RowIndex >= 0 Then
            ls = True
        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If ls = True Then
            ПереносДанных()
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        TextBox22.Text = Replace(TextBox22.Text, ".", ",")
        TextBox23.Text = Replace(TextBox23.Text, ".", ",")

        If RichTextBox1.Text = "" Then Exit Sub


        If (TextBox22.Text = "" Or Not IsNumeric(TextBox22.Text)) Then
            MessageBox.Show("Заполните поле 'Цена'!", Рик)
            Exit Sub
        End If
        If (TextBox23.Text = "" Or Not IsNumeric(TextBox23.Text)) Then
            MessageBox.Show("Заполните поле 'Объем работ'!", Рик)
            Exit Sub
        End If


        TextBox22.Text = Math.Round(CType(TextBox22.Text, Double), 2)

        TextBox22.Text = ДобРазрядности(TextBox22.Text)


        Dim f = Replace(TextBox23.Text, ".", ",")
        TextBox24.Text = CType(Math.Round(CDbl(f) * CDbl(TextBox22.Text), 2), String)
        TextBox24.Text = ДобРазрядности(TextBox24.Text)
        Button2.Enabled = True

        If MessageBox.Show("Сохранить изменения?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            btn2Click()
        End If







        'Dim txt22 As Double = Math.Round(CType(TextBox22.Text, Double), 2)



        'Dim g As Double = Math.Round((txt22 * CType(TextBox23.Text, Double)), 2)


        ''проверка на double и добавление нулей
        'If g = Math.Truncate(g) Then

        '    Dim k As Integer = Math.Truncate(Math.Log10(g)) + 1
        '    Dim p As String = CType(g, String)
        '    k = p.Length - k
        '    If k = 1 Then
        '        TextBox24.Text = CType(g, String) & "0"
        '    End If
        'Else
        '    TextBox24.Text = CType(g, String) & "00"

        'End If



        'If (CType(txt22, String).Length - Math.Truncate(Math.Log10(txt22)) + 1) = 1 Then
        '    TextBox22.Text = CType(txt22, String) & "0"
        'End If



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

    Private Sub MaskedTextBox6_LostFocus(sender As Object, e As EventArgs) Handles MaskedTextBox6.LostFocus
        If MaskedTextBox6.MaskCompleted = True Then
            ВычДатВыплат(MaskedTextBox6.Text)
        End If
    End Sub



    'Private Sub ComboBox5_LostFocus(sender As Object, e As EventArgs)
    '    Dim dt As New List(Of String)
    '    If ComboBox5.Text = "" Then Exit Sub
    '    For Each x In datrow
    '        If Not x.Item("ПорНомерАктаИное").ToString = ComboBox5.Text Then
    '            Select Case ComboBox5.Text.Length
    '                Case 1
    '                    ComboBox5.Text = "00" & ComboBox5.Text
    '                    Exit For
    '                Case 2
    '                    ComboBox5.Text = "0" & ComboBox5.Text
    '                    Exit For
    '            End Select
    '        End If
    '    Next
    'End Sub
End Class