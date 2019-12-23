Public Class ОтчетСотрудники
    Dim f As Boolean = False
    Private Sub ОтчетСотрудники_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MdiParent = MDIParent1

        Me.ComboBox1.AutoCompleteCustomSource.Clear()
        Me.ComboBox1.Items.Clear()
        For Each r As DataRow In СписокКлиентовОсновной.Rows
            Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox1.Items.Add(r(0).ToString)
        Next
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Com1Sel()
    End Sub
    Private Sub Com1Sel()

        Dim ds = From x In dtSotrudnikiAll Where x.Item("НазвОрганиз") = ComboBox1.SelectedItem
                 Order By x.Item("ФИОСборное") Select (x.Item("ФИОСборное"), x.Item("КодСотрудники"))





        '        Dim strsql As String = "SELECT DISTINCT Сотрудники.ФИОСборное, КодСотрудники FROM Сотрудники
        'WHERE Сотрудники.НазвОрганиз='" & ComboBox1.SelectedItem & "' ORDER BY Сотрудники.ФИОСборное"
        '        Dim ds As DataTable = Selects(strsql)
        Label96.Text = "N"
        ComboBox2.Items.Clear()
        Me.ComboBox19.AutoCompleteCustomSource.Clear()
        Me.ComboBox19.Items.Clear()
        For Each r In ds
            Me.ComboBox19.AutoCompleteCustomSource.Add(r.Item1.ToString())
            Me.ComboBox19.Items.Add(r.Item1.ToString)
            Me.ComboBox2.Items.Add(r.Item2.ToString)
        Next
        ComboBox19.Text = ""
        Чист()

    End Sub
    Private Sub Чист()
        For Each F_Control As Control In Me.Controls
            Dim _control As Object = Me.Controls(F_Control.Name)
            If TypeOf _control Is TextBox Then
                _control.Text = ""
                'ElseIf TypeOf _control Is ListBox Then
                '    _control.items.clear()
                'ElseIf TypeOf _control Is ComboBox Then
                '    _control.selectedindex = -1
                'ElseIf TypeOf _control Is RichTextBox Then
                '    _control.text = ""
            End If
        Next F_Control
    End Sub

    Private Sub ComboBox19_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox19.SelectedIndexChanged
        Com1Sel9()
        If f = False Then
            But1()
        End If
        f = False
    End Sub
    Private Sub Com1Sel9()
        If ComboBox1.Text = "" Then
            MessageBox.Show("Выберите организацию!", Рик)
            f = True
            Exit Sub
        End If
        Label96.Text = ComboBox2.Items.Item(ComboBox19.SelectedIndex)
        Чист()
    End Sub


    Private Sub But1()
        Dim strsql As String
        '        strsql = "SELECT Сотрудники*, КарточкаСотрудника*,ДогСотрудн*
        'FROM (Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр) INNER JOIN ДогСотрудн ON Сотрудники.КодСотрудники = ДогСотрудн.IDСотр
        'WHERE Сотрудники.КодСотрудники=" & CType(Label96.Text, Integer) & ""

        Dim ds = dtSotrudnikiAll.Select("КодСотрудники=" & CType(Label96.Text, Integer) & "")

        '        strsql = "SELECT Сотрудники.ФИОСборное, Сотрудники.ПаспортНомер,
        'Сотрудники.ПаспортСерия, Сотрудники.ПаспортКогдаВыдан,
        'Сотрудники.ДоКакогоДейств, Сотрудники.ПаспортКемВыдан, Сотрудники.ИДНомер,
        'Сотрудники.МестоПрожив, Сотрудники.СтраховойПолис, Сотрудники.НаличеДогПодряда,
        'Сотрудники.Регистрация,  Сотрудники.КонтТелГор, Сотрудники.КонтТелефон,
        'Сотрудники.Пол, Сотрудники.ДатаРожд, Сотрудники.Гражданин, Сотрудники.Иностранец,
        'Сотрудники.ФИОСборноеСтар, Сотрудники.ДатаИзменения
        'FROM Сотрудники
        'WHERE Сотрудники.КодСотрудники=" & CType(Label96.Text, Integer) & ""

        '        Dim ds As DataTable = Selects(strsql)

        TextBox1.Text = ds(0).Item("ФИОСборное").ToString
        TextBox2.Text = ds(0).Item("ПаспортНомер").ToString
        TextBox3.Text = ds(0).Item("ПаспортСерия").ToString
        TextBox4.Text = ds(0).Item("ПаспортКогдаВыдан").ToString 'дата выдачи паспорта
        TextBox5.Text = ds(0).Item("ДоКакогоДейств").ToString
        TextBox6.Text = ds(0).Item("ИДНомер").ToString 'ид паспорта
        TextBox7.Text = ds(0).Item("ПаспортКемВыдан").ToString 'кем выдан
        TextBox8.Text = ds(0).Item("СтраховойПолис").ToString 'полис
        TextBox9.Text = ds(0).Item("МестоПрожив").ToString 'прожив
        TextBox10.Text = ds(0).Item("Регистрация").ToString 'регистр
        TextBox11.Text = ds(0).Item("КонтТелГор").ToString 'тел гор
        TextBox12.Text = ds(0).Item("КонтТелефон").ToString 'тел моб
        TextBox13.Text = ds(0).Item("Пол").ToString 'пол
        TextBox14.Text = ds(0).Item("ДатаРожд").ToString 'дата рожд
        TextBox15.Text = ds(0).Item("Гражданин").ToString 'гражданство
        TextBox16.Text = ds(0).Item("ФИОСборноеСтар").ToString 'старое фио
        TextBox17.Text = Strings.Left(ds(0).Item("ДатаИзменения").ToString, 10) 'дата измен фио
        If ds(0).Item("Иностранец") = False Then
            TextBox18.Text = "Нет" 'иностранец
            TextBox18.ForeColor = Color.Black
        Else
            TextBox18.Text = "Да" 'иностранец
            TextBox18.ForeColor = Color.Red
        End If

        TextBox19.Text = ds(0).Item("НаличеДогПодряда").ToString 'наличие договора подряда
        'TextBox1.Text = ds.Rows(0).Item(0).ToString



        'Grid1.DataSource = ds
        'GridView(Grid1)


    End Sub



End Class