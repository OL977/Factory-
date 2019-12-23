Option Explicit On
Imports System.Data.OleDb

Public Class Перевод
    Public ds As DataTable
    Dim StrSql, СтПосле, ПроцПосле, ПродлКонтр, ПоСовмест, СуммирУчет, Тип, Название,
        ФормаСобстПолн, ДолжРуков, ФИОРукРодПад, ОснованиеДейств, МестоРаб, ФИОКор, ФИОКорРукДат,
        УНП, КонтТелефон, ЮрАдрес, РасСчет, Банк, БИК, АдресБанка, ЭлАдрес, ФормаСобствКор, СборноеРеквПолн,
        ДолжРуковВинПад, СтавкаНов, СклонГод, ПоСовмИлиОсн, ПоСовмПриказ, ДолжРуковРодПад, ДолжСотрВинПад,
        ДолжСотр, ФИОСторДляЗаявл, СохрЗак, ФамилияСотр, СотрФИОРод, СотрАдрес, СотрПасп, СотрПаспВыд, Пол, СохрКонтр, СтарРазряд, СохрПрик As String
    Dim massFTP2, massFTP3 As New ArrayList()
    Dim sf As Double
    Dim КодСотр, pr As Integer
    Private Sub Перевод_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Me.ComboBox1.AutoCompleteCustomSource.Clear()
        Me.ComboBox1.Items.Clear()
        For Each r As DataRow In СписокКлиентовОсновной.Rows
            Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox1.Items.Add(r(0).ToString)
            ПереводОрганиз.ListBox1.Items.Add(r(0).ToString)
        Next

    End Sub
    Private Sub Cb2()

        Dim ds = From x In dtSotrudnikiAll Where x.Item("НазвОрганиз") = ComboBox1.Text Order By x.Item("ФИОСборное") Ascending Select x.Item("ФИОСборное")

        '"SELECT ФИОСборное FROM Сотрудники WHERE НазвОрганиз='" & ComboBox1.Text & "' ORDER BY ФИОСборное "
        'Try
        '    ds.Clear()
        'Catch ex As Exception

        'End Try

        'ds = Selects(StrSql)
        Me.ComboBox2.AutoCompleteCustomSource.Clear()
        Me.ComboBox2.Items.Clear()
        For Each r In ds
            Me.ComboBox2.AutoCompleteCustomSource.Add(r.ToString())
            Me.ComboBox2.Items.Add(r.ToString)
            ПереводСотрудн.ListBox1.Items.Add(r.ToString)
        Next

        Dim ds1 = From v In dtShtatnoeOtdelyAll Where v.Item("Клиент") = ComboBox1.Text Select v.Item("Отделы") Distinct
        StrSql = "SELECT DISTINCT Отделы FROM ШтОтделы WHERE Клиент='" & ComboBox1.Text & "'"
        'ds.Clear()
        'ds = Selects(StrSql)
        Me.ComboBox3.AutoCompleteCustomSource.Clear()
        Me.ComboBox3.Items.Clear()
        For Each r In ds1
            Me.ComboBox3.AutoCompleteCustomSource.Add(r.ToString())
            Me.ComboBox3.Items.Add(r.ToString)
            ПереводНовОтдел.ListBox1.Items.Add(r.ToString)
        Next


        Dim list As New Dictionary(Of String, Object)
        list.Add("@НазвОрг", ComboBox1.Text)
        list.Add("@НазвОрганиз", ComboBox1.Text)


        Dim ds2 = Selects(StrSql:="SELECT DISTINCT КарточкаСотрудника.ДатаЗарплаты, КарточкаСотрудника.ДатаАванса
FROM (Клиент INNER JOIN Сотрудники ON Клиент.НазвОрг = Сотрудники.НазвОрганиз) INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр
WHERE Клиент.НазвОрг=@НазвОрг AND Сотрудники.НазвОрганиз =@НазвОрганиз", list)
        'ds.Clear()
        'ds = Selects(StrSql)
        Try
            TextBox9.Text = ds2.Rows(0).Item(0)
            TextBox10.Text = ds2.Rows(0).Item(1)

        Catch ex As Exception

        End Try



        Dim ds3 = From b In dtObjectObshepitaAll Where b.Item("НазвОрг") = ComboBox1.Text Select b.Item("АдресОбъекта")
        'StrSql = ""
        'StrSql = "SELECT АдресОбъекта FROM ОбъектОбщепита WHERE НазвОрг='" & ComboBox1.Text & "'"

        'ds = Selects(StrSql)
        Me.ComboBox6.Items.Clear()
        For Each r In ds3
            Me.ComboBox6.Items.Add(r.ToString)
        Next
        ComboBox6.Text = ds3.First.ToString

        'Dim Folders() As String
        'Try
        '    Folders = IO.Directory.GetDirectories(OnePath & ComboBox1.Text & "\Приказ", "*", IO.SearchOption.TopDirectoryOnly)
        'Catch ex As Exception

        'End Try

        'Dim gth4 As String
        'Try
        '    For n As Integer = 0 To Folders.Length - 1
        '        gth4 = ""
        '        gth4 = IO.Path.GetFileName(Folders(n))
        '        Folders(n) = gth4
        '        'TextBox44.Text &= gth + vbCrLf
        '    Next

        'Catch ex As Exception
        '    MessageBox.Show("У данной организации нет папки приказы!", Рик)
        '    Exit Sub
        'End Try


        Dim list2 = listFluentFTP(ComboBox1.Text & "\Приказ\")
        For Each f In list2
            ComboBox13.Items.Add(f.ToString)
        Next

        'ComboBox13.Text = Now.Year

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        clearall()
        Cb2()
    End Sub
    Private Sub clearall()
        ComboBox2.Text = String.Empty
        ComboBox3.Text = String.Empty
        ComboBox4.Text = String.Empty
        ComboBox5.Text = String.Empty
        ComboBox7.Text = String.Empty
        ComboBox8.Text = String.Empty
        ComboBox9.Text = String.Empty
        ComboBox10.Text = String.Empty
        ComboBox11.Text = String.Empty
        TextBox1.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        MaskedTextBox1.Text = ""
        MaskedTextBox3.Text = ""
        MaskedTextBox2.Text = ""
        GroupBox23.Visible = True
        GroupBox24.Visible = True
        ComboBox13.Items.Clear()
        ComboBox13.Text = String.Empty

    End Sub
    Private Function ПроверкаЗаполн()
        If ComboBox1.Text = "" Then
            MessageBox.Show("Выберите организацию!", Рик)
            Return 1
            Exit Function
        End If
        If ComboBox2.Text = "" Then
            MessageBox.Show("Выберите сотрудника!", Рик)
            Return 1
            Exit Function
        End If
        If MaskedTextBox1.MaskCompleted = False Or MaskedTextBox2.MaskCompleted = False Or MaskedTextBox3.MaskCompleted = False Then
            MessageBox.Show("Выберите дату приказа или заявления!", Рик)
            Return 1
            Exit Function
        End If
        If ComboBox3.Text = "" Then
            MessageBox.Show("Выберите новый отдел!", Рик)
            Return 1
            Exit Function
        End If
        If ComboBox4.Text = "" Then
            MessageBox.Show("Выберите новую должность!", Рик)
            Return 1
            Exit Function
        End If
        If TextBox3.Text = "" And ComboBox3.Text <> "" And ComboBox4.Text <> "" And ComboBox7.Enabled = True Then
            MessageBox.Show("Выберите разряд!", Рик)
            Return 1
            Exit Function
        End If

        If ComboBox5.Text = "" Then
            MessageBox.Show("Выберите ставку!", Рик)
            Return 1
            Exit Function
        End If

        If TextBox7.Text = "" Then
            MessageBox.Show("Выберите номер контракта!", Рик)
            Return 1
            Exit Function
        End If
        If TextBox6.Text = "" Then
            MessageBox.Show("Выберите номер приказа!", Рик)
            Return 1
            Exit Function
        End If
        If ComboBox8.Text = "" Or Not IsNumeric(ComboBox8.Text) Then
            MessageBox.Show("Выберите продолжительность контракта!", Рик)
            Return 1
            Exit Function
        End If

        If ComboBox9.Text = "" Then
            MessageBox.Show("Выберите тип работы!", Рик)
            Return 1
            Exit Function
        End If

        If GroupBox23.Visible = True And ComboBox9.Text = "" Then
            MessageBox.Show("Выберите время начала работы!", Рик)
            Return 1
            Exit Function
        End If
        If GroupBox23.Visible = True And ComboBox10.Text = "" Then
            MessageBox.Show("Выберите время начала работы!", Рик)
            Return 1
            Exit Function
        End If

        If GroupBox24.Visible = True And ComboBox11.Text = "" Then
            MessageBox.Show("Выберите продолжительность рабочего дня!", Рик)
            Return 1
            Exit Function
        End If
        If ComboBox6.Text = "" Then
            MessageBox.Show("Выберите объект общепита!", Рик)
            Return 1
            Exit Function
        End If


        Return 0
    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ПереводОрганиз.ShowDialog()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ПереводСотрудн.ShowDialog()

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

        Dim list As New Dictionary(Of String, Object)
        list.Add("@НазвОрганиз", ComboBox1.Text)
        list.Add("@ФИОСборное", ComboBox2.Text)


        Dim ds = Selects(StrSql:="SELECT Штатное.Отдел, Штатное.Должность, Штатное.Разряд
From Сотрудники INNER JOIN Штатное ON Сотрудники.КодСотрудники = Штатное.ИДСотр
WHERE Сотрудники.НазвОрганиз=@НазвОрганиз AND Сотрудники.ФИОСборное=@ФИОСборное", list)

        Try
            If ds.Rows(0).Item(2).ToString <> "" Then
                TextBox1.Text = ComboBox2.Text & ", Отдел - (" & ds.Rows(0).Item(0).ToString & "), Должность - (" & ds.Rows(0).Item(1).ToString & "), Разряд - (" & ds.Rows(0).Item(2).ToString & ")."
            Else
                TextBox1.Text = ComboBox2.Text & ", Отдел - (" & ds.Rows(0).Item(0).ToString & "), Должность - (" & ds.Rows(0).Item(1).ToString & ")."
            End If

        Catch ex As Exception
            MessageBox.Show("У данного сотрудника нет должности!", Рик)
            Exit Sub
        End Try

        ДолжСотр = ""
        ДолжСотр = ds.Rows(0).Item(1).ToString
        СтарРазряд = ds.Rows(0).Item(2).ToString


        Dim ds2 = dtSotrudnikiAll.Select("НазвОрганиз='" & ComboBox1.Text & "' AND ФИОСборное='" & ComboBox2.Text & "'")

        '        StrSql = "SELECT ФамилияДляЗаявления, ИмяДляЗаявления, ОтчествоДляЗаявления, Фамилия, ФИОРодПод, ПаспортСерия, ПаспортНомер, ПаспортКогдаВыдан,
        'ПаспортКемВыдан, ИДНомер, Регистрация, Пол FROM Сотрудники 
        'WHERE Сотрудники.НазвОрганиз='" & ComboBox1.Text & "' AND Сотрудники.ФИОСборное='" & ComboBox2.Text & "'"
        '        ds.Clear()
        '        ds = Selects(StrSql)
        ФИОСторДляЗаявл = ds2(0).Item("ФамилияДляЗаявления").ToString & " " & ds2(0).Item("ИмяДляЗаявления").ToString & " " & ds2(0).Item("ОтчествоДляЗаявления").ToString
        ФамилияСотр = ds2(0).Item("Фамилия").ToString
        СотрФИОРод = ds2(0).Item("ФИОРодПод").ToString
        СотрАдрес = ds2(0).Item("Регистрация").ToString
        СотрПасп = ds2(0).Item("ПаспортСерия").ToString & " № " & ds2(0).Item("ПаспортНомер").ToString
        СотрПаспВыд = ds2(0).Item("ПаспортКемВыдан").ToString & " " & ds2(0).Item("ПаспортКогдаВыдан").ToString & " л/н " & ds2(0).Item("ИДНомер").ToString
        Пол = ds2(0).Item("Пол").ToString


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ПереводНовОтдел.ShowDialog()
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        Me.ComboBox5.Text = String.Empty
        Me.ComboBox4.Text = String.Empty
        Me.ComboBox7.Text = String.Empty
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""

        Dim list As New Dictionary(Of String, Object)
        list.Add("@Клиент", ComboBox1.Text)
        list.Add("@Отделы", ComboBox3.Text)

        Dim ds = Selects(StrSql:="SELECT DISTINCT ШтСвод.Должность
FROM ШтОтделы INNER JOIN ШтСвод ON ШтОтделы.Код = ШтСвод.Отдел
WHERE ШтОтделы.Клиент=@Клиент AND ШтОтделы.Отделы=@Отделы", list)

        Me.ComboBox4.AutoCompleteCustomSource.Clear()
        Me.ComboBox4.Items.Clear()
        ПереводНовДолж.ListBox1.Items.Clear()

        For Each r As DataRow In ds.Rows
            Me.ComboBox4.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox4.Items.Add(r(0).ToString)
            ПереводНовДолж.ListBox1.Items.Add(r(0).ToString)
        Next
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ПереводНовДолж.ShowDialog()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ПереводСтавка.ShowDialog()
    End Sub
    Private Sub зачистка()
        ComboBox2.Text = ""
        TextBox1.Text = ""
        MaskedTextBox1.Text = ""
        MaskedTextBox2.Text = ""
        MaskedTextBox3.Text = ""
        ComboBox3.Text = ""
        ComboBox4.Text = ""
        ComboBox7.Text = ""
        ComboBox5.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        ComboBox8.Text = ""
        ComboBox9.Text = ""
        ComboBox10.Text = ""
        ComboBox11.Text = ""

    End Sub
    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged



        If CheckBox3.Checked = True Then
            зачистка()
            GroupBox5.Enabled = False
            GroupBox7.Enabled = False
            GroupBox11.Enabled = False
            GroupBox12.Enabled = False
            GroupBox13.Enabled = False
            GroupBox14.Enabled = False
            GroupBox15.Enabled = False
            GroupBox17.Enabled = False
            GroupBox18.Enabled = False
            GroupBox19.Enabled = False
            GroupBox20.Enabled = False
            GroupBox23.Enabled = False
            GroupBox24.Enabled = False
            GroupBox25.Enabled = False
            GroupBox26.Enabled = False
            CheckBox4.Enabled = False
            CheckBox2.Enabled = False


        Else

            GroupBox5.Enabled = True
            GroupBox7.Enabled = True
            GroupBox11.Enabled = True
            GroupBox12.Enabled = True
            GroupBox13.Enabled = True
            GroupBox14.Enabled = True
            GroupBox15.Enabled = True
            GroupBox17.Enabled = True
            GroupBox18.Enabled = True
            GroupBox19.Enabled = True
            GroupBox20.Enabled = True
            GroupBox23.Enabled = True
            GroupBox24.Enabled = True
            GroupBox25.Enabled = True
            GroupBox26.Enabled = True
            CheckBox4.Enabled = True
            CheckBox2.Enabled = True
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        ПереводРазр.ShowDialog()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        ПереводСрокКонтр.ShowDialog()
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked = True Then
            MaskedTextBox2.Text = MaskedTextBox1.Text
            MaskedTextBox3.Text = MaskedTextBox1.Text
        Else
            MaskedTextBox2.Clear()
            MaskedTextBox3.Clear()
        End If

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        ПереводВремНач.ShowDialog()
    End Sub

    Private Sub ComboBox13_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox13.SelectedIndexChanged
        ComboBox12.Items.Clear()
        ComboBox12.Text = String.Empty
        ComboBox14.Items.Clear()
        ComboBox14.Text = String.Empty

        Dim Files3(), Files4()

        Dim f = listFluentFTP(ComboBox1.Text & "\Приказ\" & ComboBox13.Text)
        Dim d = listFluentFTP(ComboBox1.Text & "\Контракт\" & ComboBox13.Text)

        If f.Count > 0 Then
            For Each x In f
                ComboBox12.Items.Add(x.ToString)
            Next
        End If

        If d.Count > 0 Then
            For Each x In d
                ComboBox14.Items.Add(x.ToString)
            Next
        End If


        'Files3 = (IO.Directory.GetFiles(OnePath & ComboBox1.Text & "\Приказ\" & ComboBox13.Text, "*.doc", IO.SearchOption.TopDirectoryOnly))
        'Files4 = (IO.Directory.GetFiles(OnePath & ComboBox1.Text & "\Контракт\" & ComboBox13.Text, "*.doc", IO.SearchOption.TopDirectoryOnly))
        'Dim gth As String
        'For n As Integer = 0 To Files3.Length - 1
        '    gth = ""
        '    gth = IO.Path.GetFileName(Files3(n))
        '    Files3(n) = gth
        '    'TextBox44.Text &= gth + vbCrLf
        'Next
        'ComboBox12.Items.AddRange(Files3)
        'Try
        '    ComboBox12.Text = Files3.Last.ToString
        'Catch ex As Exception
        '    MessageBox.Show("Нет файлов в папке!", Рик)
        'End Try

        'Dim gth1 As String
        'For n1 As Integer = 0 To Files4.Length - 1
        '    gth1 = ""
        '    gth1 = IO.Path.GetFileName(Files4(n1))
        '    Files4(n1) = gth1
        '    'TextBox44.Text &= gth + vbCrLf
        'Next
        'ComboBox14.Items.AddRange(Files4)
        'Try
        '    ComboBox14.Text = Files4.Last.ToString
        'Catch ex As Exception
        '    MessageBox.Show("Нет файлов в папке!", Рик)
        'End Try


    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        ПереводГрафик.ShowDialog()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        ПереводПродлРабДня.ShowDialog()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        ПереводИстория.ShowDialog()
    End Sub

    Private Sub ComboBox11_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox11.SelectedIndexChanged
        ПереводПродлРабДня.РасчПер()
    End Sub

    Private Sub ComboBox10_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox10.SelectedIndexChanged
        ПереводПродлРабДня.РасчПер()
    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        ВыбСРазрядом()

    End Sub
    Private Sub ВыбСРазрядом()
        If ComboBox3.Text <> "" And ComboBox4.Text <> "" And ComboBox7.Text <> "" Then
            '            StrSql = "Select ШтСвод.ТарифнаяСтавка,ШтСвод.ПовышениеПроц
            'From ШтОтделы INNER Join ШтСвод On ШтОтделы.Код = ШтСвод.Отдел
            'Where ШтОтделы.Отделы ='" & Отдел & "' AND ШтСвод.Должность='" & Должность & "' AND ШтСвод.Разряд='" & Разряд & "'"
            'Соед(0)
            Dim list As New Dictionary(Of String, Object)
            list.Add("@Отделы", ComboBox3.Text)
            list.Add("@Должность", ComboBox4.Text)
            list.Add("@Разряд", ComboBox7.Text)
            list.Add("@Клиент", ComboBox1.Text)


            Dim ds = Selects(StrSql:="Select  ШтСвод.ТарифнаяСтавка, ШтСвод.ПовышениеПроц
From ШтОтделы INNER Join ШтСвод On ШтОтделы.Код = ШтСвод.Отдел
Where ШтОтделы.Отделы =@Отделы AND ШтСвод.Должность =@Должность AND ШтСвод.Разряд=@Разряд AND ШтОтделы.Клиент =@Клиент", list)


            Try
                Me.TextBox3.Text = ""
                Me.TextBox4.Text = ""
                Me.TextBox5.Text = ""

                Dim dstbl As String = ds.Rows(0).Item(0).ToString

                If dstbl <> "." Then dstbl = Replace(dstbl, ".", ",")
                If dstbl <> "," Then
                    sf = Nothing
                    sf = CType(dstbl, Double)
                    Dim sfd As String = CType(sf, String)
                    Dim ДлНач As Integer = sfd.Length
                    TextBox3.Text = Math.Floor(sf)
                    Dim Дл As Integer = TextBox3.TextLength
                    ДлНач -= Дл
                    Dim vm As String

                    If ДлНач = 3 Then
                        vm = Strings.Right(Math.Round(sf - Math.Floor(sf), 2), 2)
                    ElseIf ДлНач = 2 Then
                        vm = Strings.Right(Math.Round(sf - Math.Floor(sf), 2), 1)
                    Else
                        vm = 0
                    End If
                    'Dim vm2 As String = Math.Round(sf - Math.Floor(sf), 2)

                    Dim vmn As String = CType(vm, Integer)
                    If vmn = "0" Then vm = Strings.Right(vm, 1) & "0"
                    If dstbl.Length > sfd.Length Then vm = vm & "0"

                    TextBox3.Text = TextBox3.Text & "," & vm
                    TextBox3.ForeColor = Color.Green
                Else
                    TextBox3.Text = ds.Rows(0).Item(0).ToString
                End If

                TextBox5.Text = ds.Rows(0).Item(1).ToString
            Catch ex As Exception

                MessageBox.Show("Нет данных в базе, относительно разряда!!!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            End Try

            Dim s As Double = Replace(TextBox3.Text, ".", ",")
            Dim f As Double = Replace(TextBox5.Text, ".", ",")

            TextBox4.Text = Math.Round(s + (s * f / 100), 2)


        End If
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        Me.ComboBox7.Enabled = True
        Me.ComboBox7.Text = String.Empty
        TextBox5.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""

        ВыбРазр()
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged

    End Sub

    Public Sub ВыбРазр()

        Dim list As New Dictionary(Of String, Object)
        list.Add("@Отделы", ComboBox3.Text)
        list.Add("@Должность", ComboBox4.Text)
        list.Add("@Клиент", ComboBox1.Text)


        Dim ds = Selects(StrSql:="SELECT ШтОтделы.Отделы, ШтСвод.Должность, ШтСвод.Разряд, ШтСвод.ТарифнаяСтавка,
ШтСвод.ПовышениеПроц, ШтСвод.ТарСтПослеИспСрока, ПовПроцПослеИспСрока
FROM ШтОтделы INNER JOIN ШтСвод ON ШтОтделы.Код = ШтСвод.Отдел
WHERE ШтОтделы.Отделы=@Отделы AND ШтСвод.Должность=@Должность And ШтОтделы.Клиент=@Клиент", list)

        СтПосле = ds.Rows(0).Item(5).ToString
        ПроцПосле = ds.Rows(0).Item(6).ToString

        Dim s, f As Double
        Dim ghfd As String = ds.Rows(0).Item(2).ToString
        'Dim ghfd1 As String = ds.Rows(1).Item(2).ToString

        If ds.Rows(0).Item(1) <> "" And ghfd <> "" Then
            ComboBox7.Items.Clear()
            For Each r As DataRow In ds.Rows
                ComboBox7.Items.Add(r(2).ToString)
            Next

            Dim vb As String = ds.Rows(0).Item(4).ToString()

            'TextBox5.Text = ds.Rows(0).Item(4).ToString()
            'TextBox3.Text = ds.Rows(0).Item(3).ToString()

            's = Replace(TextBox3.Text, ".", ",")
            'f = Replace(TextBox5.Text, ".", ",")
            'TextBox4.Text = Math.Round(s + (s * f / 100), 2)
        Else
            Me.ComboBox7.Enabled = False
            TextBox5.Text = ds.Rows(0).Item(4).ToString()

            Dim dstbl As String = ds.Rows(0).Item(3).ToString

            If dstbl <> "." Then dstbl = Replace(dstbl, ".", ",")
            If dstbl <> "," Then
                sf = Nothing
                sf = CType(dstbl, Double)
                TextBox3.Text = Math.Floor(sf)
                Dim vm As String = Strings.Right(Math.Round(sf - Math.Floor(sf), 2), 2)
                Dim fd As Boolean = vm.Contains(",")
                If fd = True Then
                    vm = Strings.Right(vm, 1) & "0"
                End If
                Dim vmn As String = CType(vm, Integer)
                If vmn = "0" Then vm = Strings.Right(vm, 1) & "0"
                TextBox3.Text = TextBox3.Text & "," & vm
            Else
                TextBox3.Text = ds.Rows(0).Item(0).ToString
            End If

            s = Replace(TextBox3.Text, ".", ",")
            f = Replace(TextBox5.Text, ".", ",")

            TextBox4.Text = Math.Round(s + (s * f / 100), 2)

        End If

    End Sub

    Private Sub MaskedTextBox1_MaskInputRejected(sender As Object, e As MaskInputRejectedEventArgs) Handles MaskedTextBox1.MaskInputRejected

    End Sub

    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown

    End Sub

    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox1.Focus()
        End If
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.MaskedTextBox1.Focus()
        End If
    End Sub

    Private Sub MaskedTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.MaskedTextBox3.Focus()
        End If
    End Sub
    Private Sub MaskedTextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.MaskedTextBox2.Focus()
        End If
    End Sub
    Private Sub MaskedTextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox3.Focus()
        End If
    End Sub

    Private Sub ComboBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox4.Focus()
        End If
    End Sub

    Private Sub ComboBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox7.Focus()
        End If
    End Sub

    Private Sub ComboBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox7.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox5.Focus()
        End If
    End Sub

    Private Sub ComboBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox7.Focus()
        End If
    End Sub

    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox3.Focus()
        End If
    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox4.Focus()
        End If
    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox7.Focus()
        End If
    End Sub

    Private Sub ComboBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox8.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox6.Focus()
        End If
    End Sub

    Private Sub TextBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox7.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox8.Focus()
        End If
    End Sub

    Private Sub TextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox8.Focus()
        End If
    End Sub

    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox9.Focus()
        End If
    End Sub

    Private Sub ComboBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox9.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox10.Focus()
        End If
    End Sub

    Private Sub ComboBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox10.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox11.Focus()
        End If
    End Sub

    Private Sub ComboBox11_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox11.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.Button12.Focus()
        End If
    End Sub

    Private Sub ComboBox9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox9.SelectedIndexChanged
        If ComboBox9.Text = "График" Or ComboBox9.Text = "ПВТР" Then
            GroupBox23.Visible = False
            GroupBox24.Visible = False

        Else
            GroupBox23.Visible = True
            GroupBox24.Visible = True

        End If
        TextBox12.Text = ""
        TextBox11.Text = ""
    End Sub
    Private Sub Вычис()
        Dim dad As Date = CDate(MaskedTextBox1.Text)
        Select Case ComboBox8.Text
            Case "1"
                dad = dad.AddMonths(12)
                ПродлКонтр = dad.AddDays(-1)
            Case "2"
                dad = dad.AddMonths(24)
                ПродлКонтр = dad.AddDays(-1)
            Case "3"
                dad = dad.AddMonths(36)
                ПродлКонтр = dad.AddDays(-1)
            Case "4"
                dad = dad.AddMonths(48)
                ПродлКонтр = dad.AddDays(-1)
            Case "5"
                dad = dad.AddMonths(60)
                ПродлКонтр = dad.AddDays(-1)

        End Select
    End Sub
    Private Sub Обновление()
        'StrSql = "SELECT Сотрудники.КодСотрудники From Сотрудники WHERE ФИОСборное='" & ComboBox2.Text & "' and НазвОрганиз='" & ComboBox1.Text & "'"
        'ds.Clear()
        'ds = Selects(StrSql)
        'КодСотр = ds.Rows(0).Item(0)
        Вычис()

        Dim tm As Date
        Try
            tm = CDate(MaskedTextBox1.Text).AddYears(CType(ComboBox8.Text, Integer))
            tm.AddMonths(-1)
            tm.AddDays(-1)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Dim ПоСовмест, СуммирУчет As String

        If CheckBox2.Checked = True Then
            ПоСовмест = "по совместительству"
        Else
            ПоСовмест = ""
        End If

        If CheckBox4.Checked = True Then
            СуммирУчет = "Да"
        Else
            СуммирУчет = ""
        End If

        If TextBox8.Text <> "" Then
            TextBox6.Text = TextBox6.Text & " к-" & TextBox8.Text
        End If

        Dim f = dtKartochkaSotrudnikaAll.Select("IDСотр=" & КодСотр & "")

        If f.Count > 0 Then
            Updates(stroka:="DELETE FROM КарточкаСотрудника WHERE IDСотр=" & КодСотр & "")
        End If


        Dim list2 As New Dictionary(Of String, Object)
        list2.Add("@IDСотр", КодСотр)
        list2.Add("@ДатаПриема", MaskedTextBox1.Text)
        list2.Add("@ДатаУведомлПродКонтр", tm)

        Updates(stroka:="INSERT INTO КарточкаСотрудника(Ставка,ДатаПеревода,ДатаЗаявленияПеревода,ДатаПриказаПеревода,НомерПриказаПеревода,
ДатаПриема,СрокКонтракта,ТипРаботы,ВремяНачРаботы,Обед,ОкончРабДня,ДатаУведомлПродКонтр,НеПродлениеКонтр,АдресОбъектаОбщепита,
ДатаЗарплаты,ДатаАванса,ПоСовмест,СуммирУчет,IDСотр)
VALUES('" & ComboBox5.Text & "','" & MaskedTextBox3.Text & "','" & MaskedTextBox2.Text & "','" & MaskedTextBox1.Text & "','" & TextBox6.Text & "',
@ДатаПриема,'" & ComboBox8.Text & "','" & ComboBox9.Text & "','" & ComboBox10.Text & "','" & TextBox11.Text & "',
'" & TextBox12.Text & "',@ДатаУведомлПродКонтр,'False','" & f(0).Item("АдресОбъектаОбщепита").ToString & "','" & f(0).Item("ДатаЗарплаты").ToString & "',
'" & f(0).Item("ДатаАванса").ToString & "','" & ПоСовмест & "','" & СуммирУчет & "', @IDСотр)", list2, "КарточкаСотрудника")


        '        StrSql = ""
        '        StrSql = "UPDATE КарточкаСотрудника SET ДатаПриема='" & MaskedTextBox1.Text & "', СрокКонтракта='" & ComboBox8.Text & "', ТипРаботы='" & ComboBox9.Text & "', 
        'Ставка='" & ComboBox5.Text & "', ВремяНачРаботы='" & ComboBox10.Text & "', ПродолРабДня='" & ComboBox11.Text & "', Обед='" & TextBox11.Text & "',
        'ОкончРабДня='" & TextBox12.Text & "', ДатаУведомлПродКонтр='" & ПродлКонтр & "', ДатаПеревода='" & MaskedTextBox3.Text & "',
        'ПоСовмест='" & ПоСовмест & "', СуммирУчет='" & СуммирУчет & "', ДатаЗаявленияПеревода='" & MaskedTextBox2.Text & "', ДатаПриказаПеревода='" & MaskedTextBox1.Text & "',
        'НомерПриказаПеревода= '" & TextBox6.Text & "' WHERE IDСотр=" & КодСотр & ""
        '        Updates(StrSql)


        Updates(stroka:= "UPDATE ДогСотрудн SET Контракт='" & TextBox7.Text & "', ДатаКонтракта='" & MaskedTextBox1.Text & "', 
СрокОкончКонтр='" & ПродлКонтр & "', Приказ='" & TextBox6.Text & "', Датаприказа='" & MaskedTextBox1.Text & "', Перевод='Переведен на новую должность' 
        WHERE IDСотр=@IDСотр", list2)



        Dim fr As Double
        Dim dx As Double = Replace(ComboBox5.Text, ".", ",")
        fr = Math.Round(CType(TextBox4.Text, Double) * dx, 2)
        Dim list As New Dictionary(Of String, Object)
        list.Add("@Должность", ComboBox4.Text)
        list.Add("@Разряд", ComboBox7.Text)
        list.Add("@ТарифнаяСтавка", Math.Round(CType(TextBox3.Text, Double), 2))
        list.Add("@ПовышОклПроц", Math.Round(CType(TextBox5.Text, Double), 2))
        list.Add("@РасчДолжностнОклад", Math.Round(CType(TextBox4.Text, Double), 2))
        list.Add("@ФонОплатыТруда", fr)
        list.Add("@ИДСотр", КодСотр)






        Updates(stroka:="UPDATE Штатное SET Отдел='" & ComboBox3.Text & "', Должность=@Должность, 
Разряд=@Разряд, ТарифнаяСтавка=@ТарифнаяСтавка, ПовышОклПроц=@ПовышОклПроц,
РасчДолжностнОклад=@РасчДолжностнОклад, ФонОплатыТруда=@ФонОплатыТруда
        WHERE ИДСотр=@ИДСотр", list, "Штатное")


        Статистика(ComboBox2.Text, "Перевод сотрудника на другую должность", ComboBox1.Text)

    End Sub
    Function Проверка() As Integer

        If ComboBox1.Text = "" Then
            MessageBox.Show("Выберите организацию!", Рик)
            Return 1
        End If
        If ComboBox2.Text = "" Then
            MessageBox.Show("Выберите сотрудника!", Рик)
            Return 1
        End If
        If MaskedTextBox1.MaskCompleted = False Then
            MessageBox.Show("Выберите дату приказа!", Рик)
            Return 1
        End If
        If MaskedTextBox2.MaskCompleted = False Then
            MessageBox.Show("Выберите дату заявления!", Рик)
            Return 1
        End If
        If MaskedTextBox3.MaskCompleted = False Then
            MessageBox.Show("Выберите дату перевода!", Рик)
            Return 1
        End If
        If ComboBox5.Text = "" Then
            MessageBox.Show("Выберите ставку!", Рик)
            Return 1
        End If
        If TextBox6.Text = "" Then
            MessageBox.Show("Выберите номер приказа!", Рик)
            Return 1
        End If
        If ComboBox6.Text = "" Then
            MessageBox.Show("Выберите обьект общепита!", Рик)
            Return 1
        End If



        Return 0
    End Function
    Private Sub СохрВБазу()
        'StrSql = ""
        'StrSql = "SELECT Сотрудники.КодСотрудники From Сотрудники WHERE ФИОСборное='" & ComboBox2.Text & "' and НазвОрганиз='" & ComboBox1.Text & "'"
        'ds.Clear()
        'ds = Selects(StrSql)
        'КодСотр = ds.Rows(0).Item(0)

        Dim f = dtKartochkaSotrudnikaAll.Select("IDСотр=" & КодСотр & "")
        'If f.Count > 0 Then
        '    Updates(stroka:="DELETE FROM КарточкаСотрудника WHERE IDСотр=" & КодСотр & "")
        'End If

        'Dim tm As Date
        'Try
        '    tm = CDate(MaskedTextBox1.Text).AddYears(CType(ComboBox8.Text, Integer))
        '    tm.AddMonths(-1)
        '    tm.AddDays(-1)
        'Catch ex As Exception
        '    MessageBox.Show(ex.Message)
        'End Try
        'Dim ПоСовмест, СуммирУчет As String

        'If CheckBox2.Checked = True Then
        '    ПоСовмест = "по совместительству"
        'Else
        '    ПоСовмест = ""
        'End If

        'If CheckBox4.Checked = True Then
        '    СуммирУчет = "Да"
        'Else
        '    СуммирУчет = ""
        'End If

        Dim s As String
        If TextBox8.Text <> "" Then
            s = TextBox6.Text & " к " & TextBox8.Text
        Else
            s = TextBox6.Text
        End If



        Dim list2 As New Dictionary(Of String, Object)
        list2.Add("@IDСотр", КодСотр)
        'list2.Add("@ДатаПриема", MaskedTextBox1.Text)
        'list2.Add("@ДатаУведомлПродКонтр", tm)



        Updates(stroka:="UPDATE КарточкаСотрудника SET Ставка='" & ComboBox5.Text & "',ДатаПеревода='" & MaskedTextBox3.Text & "',
ДатаЗаявленияПеревода='" & MaskedTextBox2.Text & "',ДатаПриказаПеревода='" & MaskedTextBox1.Text & "',НомерПриказаПеревода='" & s & "'
WHERE IDСотр=@IDСотр", list2)

        '        Updates(stroka:="INSERT INTO КарточкаСотрудника(Ставка,ДатаПеревода,ДатаЗаявленияПеревода,ДатаПриказаПеревода,НомерПриказаПеревода,
        'ДатаПриема,СрокКонтракта,ТипРаботы,ВремяНачРаботы,Обед,ОкончРабДня,ДатаУведомлПродКонтр,НеПродлениеКонтр,АдресОбъектаОбщепита,
        'ДатаЗарплаты,ДатаАванса,ПоСовмест,СуммирУчет,IDСотр)
        'VALUES('" & ComboBox5.Text & "','" & MaskedTextBox3.Text & "','" & MaskedTextBox2.Text & "','" & MaskedTextBox1.Text & "','" & TextBox6.Text & "',
        '@ДатаПриема,'" & ComboBox8.Text & "','" & ComboBox9.Text & "','" & ComboBox10.Text & "','" & TextBox11.Text & "',
        ''" & TextBox12.Text & "',@ДатаУведомлПродКонтр,'False','" & f(0).Item("АдресОбъектаОбщепита").ToString & "','" & f(0).Item("ДатаЗарплаты").ToString & "',
        ''" & f(0).Item("ДатаАванса").ToString & "','" & ПоСовмест & "','" & СуммирУчет & "', @IDСотр)", list2, "КарточкаСотрудника")


        Dim ds = dtSotrudnikiAll.Select("КодСотрудники=" & КодСотр & "")
        '        StrSql = "SELECT ФамилияДляЗаявления, ИмяДляЗаявления, ОтчествоДляЗаявления, Фамилия, ФИОРодПод FROM Сотрудники 
        'WHERE Сотрудники.КодСотрудники=" & КодСотр & ""
        '        ds.Clear()
        '        ds = Selects(StrSql)

        ФИОСторДляЗаявл = ds(0).Item("ФамилияДляЗаявления").ToString & " " & ds(0).Item("ИмяДляЗаявления").ToString & " " & ds(0).Item("ОтчествоДляЗаявления").ToString
        ФамилияСотр = ds(0).Item("Фамилия").ToString
        СотрФИОРод = ds(0).Item("ФИОРодПод").ToString


        Статистика(ComboBox2.Text, "Перевод сотрудника на другую ставку", ComboBox1.Text)



    End Sub
    Private Sub ДокиПереводаставка()

        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document
        oWord = CreateObject("Word.Application")
        oWord.Visible = False

        Начало("ZayavleniePerevodStavka.docx")
        oWordDoc = oWord.Documents.Add(firthtPath & "\ZayavleniePerevodStavka.docx")

        ДолжСотрВинПад = ДобОконч(ДолжСотр)
        СтавкаНов = Склонение(ComboBox5.Text) 'склонение ставки
        ДолжРуковРодПад = ДолжРодПадежФункц(ДолжРуков)

        With oWordDoc.Bookmarks
            .Item("ЗАКЛпер0").Range.Text = MaskedTextBox2.Text
            If Not ComboBox1.Text = "Итал Гэлэри Плюс" Then
                If ДолжРуковРодПад = "Индивидуальному предпринимателю" Or ФормаСобствКор = "ИП" Then
                    .Item("ЗАКЛпер1").Range.Text = ДолжРуковРодПад
                Else
                    .Item("ЗАКЛпер1").Range.Text = ДолжРуковРодПад & " """ & ФормаСобствКор & """ " & ComboBox1.Text
                End If
                .Item("ЗАКЛпер2").Range.Text = ФИОКорРукДат
            End If
            If СтарРазряд <> "" And Not СтарРазряд = "-" Then
                .Item("ЗАКЛпер3").Range.Text = ДолжСотрВинПад & " " & разрядстрока(CType(СтарРазряд, Integer))
            Else
                .Item("ЗАКЛпер3").Range.Text = ДолжСотрВинПад
            End If

            .Item("ЗАКЛпер4").Range.Text = ФИОСторДляЗаявл
            .Item("ЗАКЛпер6").Range.Text = MaskedTextBox3.Text
            .Item("ЗАКЛпер7").Range.Text = MaskedTextBox2.Text
            .Item("ЗАКЛпер8").Range.Text = ФИОКорРук(ComboBox2.Text, False)
            .Item("ЗАКЛпер9").Range.Text = Replace(ComboBox5.Text, ".", ",") & " " & СтавкаНов

        End With

        Dim Name As String = ФамилияСотр & " (Заявление_Перевод_Ставка)" & ".docx"
        Dim СохрЗак As New List(Of String)
        СохрЗак.AddRange(New String() {ComboBox1.Text & "\Заявление\" & Now.Year, Name})
        oWordDoc.SaveAs2(PathVremyanka & Name,,,,,, False)
        oWordDoc.Close(True)
        oWord.Quit(True)
        Конец(ComboBox1.Text & "\Заявление\" & Now.Year, Name, КодСотр, ComboBox1.Text, "\ZayavleniePerevodStavka.docx", "Заявление_Перевод_Ставка")
        massFTP3.Add(СохрЗак)


        Dim oWord1 As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc1 As Microsoft.Office.Interop.Word.Document
        oWord1 = CreateObject("Word.Application")
        oWord1.Visible = False

        Начало("PrikazNaPerevodStavka.doc")
        oWordDoc1 = oWord1.Documents.Add(firthtPath & "\PrikazNaPerevodStavka.doc")

        With oWordDoc1.Bookmarks
            .Item("П1").Range.Text = MaskedTextBox1.Text
            If TextBox8.Text <> "" Then
                .Item("П2").Range.Text = TextBox6.Text & " к-" & TextBox8.Text
            Else
                .Item("П2").Range.Text = TextBox6.Text & "-к"
            End If

            .Item("П3").Range.Text = ФИОКорРук(ФИОСторДляЗаявл, False)
            .Item("П6").Range.Text = СотрФИОРод

            Select Case СтарРазряд
                Case = ""
                    .Item("П9").Range.Text = Strings.LCase(ДобОконч(ДолжСотр))
                Case <> ""
                    .Item("П9").Range.Text = Strings.LCase(ДобОконч(ДолжСотр)) & " " & СтарРазряд & "-го разряда"
            End Select
            .Item("П13").Range.Text = MaskedTextBox3.Text
            .Item("П17").Range.Text = ФИОКорРук(ФИОСторДляЗаявл, False)
            .Item("П22").Range.Text = ФИОКорРук(ComboBox2.Text, False)
            .Item("П25").Range.Text = ФормаСобстПолн
            If ФормаСобстПолн = "Индивидуальный предприниматель" Then
                .Item("П26").Range.Text = ComboBox1.Text
            Else
                .Item("П26").Range.Text = " «" & ComboBox1.Text & "» "
            End If

            .Item("П27").Range.Text = ЮрАдрес
            .Item("П28").Range.Text = УНП
            .Item("П29").Range.Text = РасСчет
            .Item("П30").Range.Text = АдресБанка
            .Item("П31").Range.Text = БИК
            .Item("П33").Range.Text = ЭлАдрес
            .Item("П34").Range.Text = КонтТелефон
            '.Item("П35").Range.Text = МестоРаб

            If Not ComboBox1.Text = "Итал Гэлэри Плюс" Then
                If ДолжРуков = "Индивидуальный предприниматель" Then
                    .Item("П36").Range.Text = ДолжРуков
                    .Item("П37").Range.Text = ""
                    .Item("П38").Range.Text = ФИОКор
                Else
                    .Item("П36").Range.Text = ДолжРуков & " " & ФормаСобствКор
                    .Item("П37").Range.Text = "«" & ComboBox1.Text & "»"
                    .Item("П38").Range.Text = ФИОКор
                End If

            End If
            .Item("П39").Range.Text = Replace(ComboBox5.Text, ".", ",")
            .Item("П40").Range.Text = Склонение(ComboBox5.Text)
        End With


        Dim Name1 As String = TextBox6.Text & " перевод " & ФамилияСотр & " от " & Me.MaskedTextBox1.Text & " (ПриказПереводСтавка)" & " .doc"
        Dim СохрПрик As New List(Of String)
        СохрПрик.AddRange(New String() {ComboBox1.Text & "\ПриказПеревод\" & Now.Year, Name1})
        oWordDoc1.SaveAs2(PathVremyanka & Name1,,,,,, False)
        oWordDoc1.Close(True)
        oWord1.Quit(True)
        Конец(ComboBox1.Text & "\ПриказПеревод\" & Now.Year, Name1, КодСотр, ComboBox1.Text, "\PrikazNaPerevodStavka.doc", "ПриказПереводСтавка")
        massFTP3.Add(СохрПрик)


    End Sub
    Private Sub УчетИзм()
        Dim list As New Dictionary(Of String, Object)
        list.Add("@КодСотрудники", КодСотр)



        Dim g As Boolean = False
        Dim ds2 = Selects(StrSql:="SELECT КарточкаСотрудника.ДатаПриема, Штатное.Должность, Штатное.Разряд, КарточкаСотрудника.Ставка, Штатное.Отдел
FROM (Сотрудники INNER JOIN Штатное ON Сотрудники.КодСотрудники = Штатное.ИДСотр) INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр
WHERE Сотрудники.КодСотрудники=@КодСотрудники", list)
        'Dim ds2 As DataTable = Selects(strsql2)
        Dim ds = From x In dtPerevodAll Where x.Item("IDСотр") = КодСотр Select x
        'Dim strsql As String = "SELECT * FROM Перевод WHERE IDСотр=" & КодСотр & ""
        'Dim ds As DataTable = Selects(strsql)

        Dim dscoun As Integer = ds.Count
        Dim dscyrcl As Integer
        If ds.Count = 0 Then 'если вообще нет строк (т.е первая)
            If CheckBox3.Checked = False Then
                dscyrcl += 1
                Updates(stroka:="INSERT INTO Перевод(ДатаДолжСтарС,ДолжСтар,ДолжНов,ДатаДолжНов,ДатаРазрСтарС,РазрСтар,РазрНов,ДатаРазрНов,
ДатаТарифСтарс,ТарифСтар,ТарифНов,ДатаТарифНов,IDСотр,ДатаОтделСтар,ОтделСтар,ОтделНов,ДатаОтдНов,Организация,ФИОСотр,ИзменСтавка) VALUES( '" & ds2.Rows(0).Item(0).ToString & "','" & ds2.Rows(0).Item(1).ToString & "','" & ComboBox4.Text & "',
'" & MaskedTextBox3.Text & "','" & ds2.Rows(0).Item(0).ToString & "','" & ds2.Rows(0).Item(2).ToString & "', '" & ComboBox7.Text & "','" & MaskedTextBox3.Text & "',
'" & ds2.Rows(0).Item(0).ToString & "','" & ds2.Rows(0).Item(3).ToString & "','" & ComboBox5.Text & "','" & MaskedTextBox3.Text & "', " & КодСотр & ",'" & ds2.Rows(0).Item(0).ToString & "',
'" & ds2.Rows(0).Item(4).ToString & "','" & ComboBox3.Text & "', '" & MaskedTextBox3.Text & "','" & ComboBox1.Text & "','" & ComboBox2.Text & "','False')", list, "Перевод")

            Else
                Updates(stroka:="INSERT INTO Перевод(ДатаДолжСтарС,ДолжСтар,ДолжНов,ДатаДолжНов,ДатаРазрСтарС,РазрСтар,РазрНов,ДатаРазрНов,
ДатаТарифСтарс,ТарифСтар,ТарифНов,ДатаТарифНов,IDСотр,ИзменСтавка,ДатаОтделСтар,ОтделСтар,ОтделНов,ДатаОтдНов,Организация,ФИОСотр) VALUES( '" & ds2.Rows(0).Item(0).ToString & "','" & ds2.Rows(0).Item(1).ToString & "','" & ds2.Rows(0).Item(1).ToString & "',
'" & MaskedTextBox3.Text & "','" & ds2.Rows(0).Item(0).ToString & "','" & ds2.Rows(0).Item(2).ToString & "', '" & ds2.Rows(0).Item(2).ToString & "','" & ds2.Rows(0).Item(0).ToString & "',
'" & ds2.Rows(0).Item(0).ToString & "','" & ds2.Rows(0).Item(3).ToString & "','" & ComboBox5.Text & "','" & MaskedTextBox3.Text & "', " & КодСотр & ", 'True','" & ds2.Rows(0).Item(0).ToString & "','" & ds2.Rows(0).Item(4).ToString & "',
'" & ds2.Rows(0).Item(4).ToString & "','" & ds2.Rows(0).Item(0).ToString & "','" & ComboBox1.Text & "','" & ComboBox2.Text & "')", list, "Перевод")
                dscyrcl += 1

            End If
        Else
            If ds.Count = 1 Then 'если одна строка существует сотрудника

                If CDate(ds(0).Item(5).ToString) = CDate(MaskedTextBox3.Text) Or CDate(ds(0).Item(9).ToString) = CDate(MaskedTextBox3.Text) Or CDate(ds(0).Item(13).ToString) = CDate(MaskedTextBox3.Text) Then
                    If CheckBox3.Checked = True Then
                        Updates(stroka:="UPDATE Перевод SET ТарифНов='" & ComboBox5.Text & "',ИзменСтавка='True'
WHERE IDСотр=@КодСотрудники", list, "Перевод")

                    Else

                        Updates(stroka:="UPDATE Перевод SET ДолжСтар='" & ds2.Rows(0).Item(1).ToString & "', ДолжНов='" & ComboBox4.Text & "',
РазрСтар='" & ds2.Rows(0).Item(2).ToString & "',РазрНов='" & ComboBox7.Text & "',ОтделСтар='" & ds2.Rows(0).Item(4).ToString & "',
ОтделНов='" & ComboBox3.Text & "'
WHERE IDСотр=@КодСотрудники", list, "Перевод")

                    End If

                Else
                    If CheckBox3.Checked = True Then
                        Updates(stroka:="INSERT INTO Перевод(ДатаДолжСтарС,ДолжСтар,ДолжНов,ДатаДолжНов,ДатаРазрСтарС,РазрСтар,РазрНов,ДатаРазрНов,
ДатаТарифСтарс,ТарифСтар,ТарифНов,ДатаТарифНов,IDСотр,ИзменСтавка,ДатаОтделСтар,ОтделСтар,ОтделНов,ДатаОтдНов,Организация,ФИОСотр) VALUES( '" & ds(0).Item(5).ToString & "','" & ds(0).Item(4).ToString & "','" & ds(0).Item(4).ToString & "',
'" & MaskedTextBox3.Text & "','" & ds(0).Item(6).ToString & "','" & ds(0).Item(8).ToString & "', '" & ds(0).Item(8).ToString & "','" & ds(0).Item(9).ToString & "',
'" & ds(0).Item(13).ToString & "','" & ds(0).Item(12).ToString & "','" & ComboBox5.Text & "','" & MaskedTextBox3.Text & "', " & КодСотр & ", 'True',
'" & ds(0).Item(15).ToString & "', '" & ds(0).Item(17).ToString & "','" & ds(0).Item(17).ToString & "','" & ds(0).Item(18).ToString & "','" & ComboBox1.Text & "','" & ComboBox2.Text & "')", list, "Перевод")


                    Else
                        Updates(stroka:="INSERT INTO Перевод(ДатаДолжСтарС,ДолжСтар,ДолжНов,ДатаДолжНов,ДатаРазрСтарС,РазрСтар,РазрНов,ДатаРазрНов,
ДатаТарифСтарс,ТарифСтар,ТарифНов,ДатаТарифНов,IDСотр,ДатаОтделСтар,ОтделСтар,ОтделНов,ДатаОтдНов,Организация,ФИОСотр,ИзменСтавка) VALUES( '" & ds(0).Item(5).ToString & "','" & ds(0).Item(4).ToString & "','" & ComboBox4.Text & "',
'" & MaskedTextBox3.Text & "','" & ds(0).Item(9).ToString & "','" & ds(0).Item(8).ToString & "', '" & ComboBox7.Text & "','" & MaskedTextBox3.Text & "',
'" & ds(0).Item(13).ToString & "','" & ds(0).Item(12).ToString & "','" & ComboBox5.Text & "','" & MaskedTextBox1.Text & "', " & КодСотр & ", '" & ds(0).Item(18).ToString & "',
'" & ds(0).Item(17).ToString & "','" & ComboBox3.Text & "', '" & MaskedTextBox3.Text & "','" & ComboBox1.Text & "','" & ComboBox2.Text & "','False')", list, "Перевод")

                    End If

                End If
            Else 'если несколько строк существует сотрудника

                For Each r In ds 'если изменяем данные по дате
                    dscyrcl += 1
                    If CDate(r.Item(13).ToString) = CDate(MaskedTextBox3.Text) Then
                        Dim j As String = Format(r.Item(13).ToString, "MM\/dd\/yyyy")
                        Dim list2 As New Dictionary(Of String, Object)
                        list2.Add("@IDСотр", r.Item(1))
                        list2.Add("@ДатаТарифНов", r.Item(13).ToString)

                        If CheckBox3.Checked = True Then

                            Updates(stroka:="UPDATE Перевод
SET ТарифНов='" & ComboBox5.Text & "', ИзменСтавка='True' 
WHERE IDСотр=@IDСотр AND ДатаТарифНов=@ДатаТарифНов", list2, "Перевод")

                            Exit Sub
                        End If
                    Else
                        Dim list3 As New Dictionary(Of String, Object)
                        list3.Add("@IDСотр", r.Item(1))

                        If CheckBox3.Checked = True Then 'если добавляем новую тарифн ставку
                            If dscoun - dscyrcl = 0 Then
                                g = False

                                Updates(stroka:="INSERT INTO Перевод(ДатаТарифСтарс,ТарифСтар,ТарифНов,ДатаТарифНов,IDСотр,
ДатаДолжСтарС, ДолжСтар, ДолжНов, ДатаДолжНов,ДатаРазрСтарС,РазрСтар,РазрНов,ДатаРазрНов,ИзменСтавка,ДатаОтделСтар,ОтделСтар,ОтделНов,ДатаОтдНов,Организация,ФИОСотр)
VALUES( '" & r.Item(13).ToString & "','" & r.Item(12).ToString & "', '" & ComboBox5.Text & "','" & MaskedTextBox1.Text & "'," & r.Item(1) & ",
'" & CDate(r.Item(5).ToString) & "', '" & r.Item(4).ToString & "','" & r.Item(4).ToString & "',
'" & MaskedTextBox3.Text & "','" & CDate(r.Item(6).ToString) & "','" & r.Item(8).ToString & "','" & r.Item(8).ToString & "',
'" & CDate(r.Item(9).ToString) & "', 'True','" & r.Item(15).ToString & "','" & r.Item(17).ToString & "','" & r.Item(17).ToString & "',
'" & MaskedTextBox3.Text & "','" & ComboBox1.Text & "','" & ComboBox2.Text & "')", list3, "Перевод")

                                g = True
                            End If
                        End If
                    End If

                    If CDate(r.Item(5).ToString) = CDate(MaskedTextBox3.Text) Then 'если дата ДатаДолжНов совпадает с датой перевода
                        g = False
                        Dim df As String = Format(r.Item(5), "MM\/dd\/yyyy")

                        Dim list4 As New Dictionary(Of String, Object)
                        list4.Add("@IDСотр", r.Item(1))
                        list4.Add("@ДатаДолжНов", df)

                        Updates(stroka:="UPDATE Перевод SET ДолжНов='" & ComboBox4.Text & "', РазрНов='" & ComboBox7.Text & "',
ТарифНов='" & ComboBox5.Text & "', ОтделНов='" & ComboBox3.Text & "' 
WHERE IDСотр=@IDСотр AND ДатаДолжНов=@ДатаДолжНов", list, "Перевод")
                        g = True

                    Else

                        If dscoun - dscyrcl = 0 And g = False Then
                            Updates(stroka:="INSERT INTO Перевод(ДатаДолжСтарС, ДолжСтар, ДолжНов, ДатаДолжНов,ДатаРазрСтарС,
РазрСтар,РазрНов,ДатаРазрНов,IDСотр,ДатаТарифСтарс,ТарифСтар,ТарифНов,ДатаТарифНов,ДатаОтделСтар,ОтделСтар,ОтделНов,ДатаОтдНов,Организация,ФИОСотр,ИзменСтавка)
VALUES('" & CDate(r.Item(5).ToString) & "', '" & r.Item(4).ToString & "','" & ComboBox4.Text & "', '" & MaskedTextBox3.Text & "',
'" & CDate(r.Item(9).ToString) & "', '" & r.Item(8).ToString & "','" & ComboBox7.Text & "', '" & MaskedTextBox3.Text & "'," & r.Item(1) & ", '" & CDate(r.Item(13).ToString) & "',
'" & r.Item(12).ToString & "','" & ComboBox5.Text & "','" & MaskedTextBox3.Text & "','" & r.Item(18).ToString & "','" & r.Item(17).ToString & "',
'" & ComboBox3.Text & "','" & MaskedTextBox3.Text & "','" & ComboBox1.Text & "','" & ComboBox2.Text & "','False')", list, "Перевод")

                        End If
                    End If
                Next
            End If
        End If

    End Sub
    Private Sub КодСотрУд()

        Dim f = dtSotrudnikiAll.Select("ФИОСборное='" & ComboBox2.Text & "' and НазвОрганиз='" & ComboBox1.Text & "'")

        'StrSql = "SELECT Сотрудники.КодСотрудники From Сотрудники WHERE ФИОСборное='" & ComboBox2.Text & "' and НазвОрганиз='" & ComboBox1.Text & "'"
        'ds.Clear()
        'ds = Selects(StrSql)
        КодСотр = f(0).Item("КодСотрудники")
    End Sub
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click

        Dim mass() As String
        Me.Cursor = Cursors.WaitCursor
        КодСотрУд()

        If CheckBox3.Checked = True Then
            pr = Проверка()
            If pr = 1 Then
                Exit Sub
            End If


            УчетИзм()
            СохрВБазу()

            If CheckBox1.Checked = False Then
                СборДанОрганиз()
                ДокиПереводаставка()
                If MessageBox.Show("Все данные внесены и оформлены! Распечатать пакет документов? ", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.None) = DialogResult.OK Then
                    ПечатьДоковFTP(massFTP3)
                End If
            Else
                MessageBox.Show("Данные внесены в базу!", Рик)
            End If

            Me.Cursor = Cursors.Default
        Me.Close()
        Me.Dispose()
        Exit Sub
        End If






        Dim sd As Integer = ПроверкаЗаполн()
        If sd = 1 Then
            Me.Cursor = Cursors.Default
            Exit Sub
        End If

        УчетИзм()
        Обновление() 'сохранение в базу
        If CheckBox1.Checked = False Then
            СборДанОрганиз()
            Доки()
            If MessageBox.Show("Все данные внесены и оформлены! Распечатать пакет документов? ", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.None) = DialogResult.OK Then
                ПечатьДоковFTP(massFTP2)
            End If
            Me.Cursor = Cursors.Default
            Me.Hide()
            Me.Close()
            Exit Sub
        End If


        Me.Cursor = Cursors.Default
        MessageBox.Show("Данные внесены в базу!", Рик)
        Me.Close()
        Me.Dispose()

    End Sub
    Private Sub Доки()

        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document
        oWord = CreateObject("Word.Application")
        oWord.Visible = False
        Начало("ZayavleniePerevod.docx")

        oWordDoc = oWord.Documents.Add(firthtPath & "\ZayavleniePerevod.docx")
        'Dim d As String = Заявление(6)
        ДолжСотрВинПад = ДобОконч(ДолжСотр)

        СтавкаНов = Склонение(ComboBox5.Text) 'склонение ставки
        СклонГод = Склонение2(ComboBox8.Text) ' склонение год

        If CheckBox2.Checked = True Then 'галочка по осн или по совместительству
            ПоСовмИлиОсн = "совместительству"
            ПоСовмПриказ = "по совместительству"
        Else
            ПоСовмИлиОсн = "основной работе"
            ПоСовмПриказ = "основное место работы"
        End If
        ДолжРуковРодПад = ДолжРодПадежФункц(ДолжРуков)
        'MsgBox(ДолжСОконч)
        With oWordDoc.Bookmarks
            .Item("ЗАКЛпер0").Range.Text = MaskedTextBox2.Text
            If ДолжРуковРодПад = "Индивидуальному предпринимателю" Or ФормаСобствКор = "ИП" Then
                .Item("ЗАКЛпер1").Range.Text = ДолжРуковРодПад
            Else
                .Item("ЗАКЛпер1").Range.Text = ДолжРуковРодПад & " """ & ФормаСобствКор & """ " & ComboBox1.Text
            End If
            .Item("ЗАКЛпер2").Range.Text = ФИОКорРукДат

            If СтарРазряд <> "" And Not СтарРазряд = "-" Then
                .Item("ЗАКЛпер3").Range.Text = ДолжСотрВинПад & " " & разрядстрока(CType(СтарРазряд, Integer))
            Else
                .Item("ЗАКЛпер3").Range.Text = ДолжСотрВинПад
            End If
            .Item("ЗАКЛпер4").Range.Text = ФИОСторДляЗаявл

            If ComboBox7.Text <> "" And Not ComboBox7.Text = "-" Then
                .Item("ЗАКЛпер5").Range.Text = Strings.LCase(ДобОконч(ComboBox4.Text)) & " " & разрядстрока(CType(ComboBox7.Text, Integer))
            Else
                .Item("ЗАКЛпер5").Range.Text = Strings.LCase(ДобОконч(ComboBox4.Text))
            End If

            If CheckBox2.Checked = True Then
                .Item("ЗАКЛпер6").Range.Text = MaskedTextBox3.Text & "г." & " на " & ComboBox5.Text & " " & Склонение(ComboBox5.Text) & " по совместительству."
            Else
                .Item("ЗАКЛпер6").Range.Text = MaskedTextBox3.Text & "г." & " на " & ComboBox5.Text & " " & Склонение(ComboBox5.Text) & "."
            End If

            .Item("ЗАКЛпер7").Range.Text = MaskedTextBox2.Text
            .Item("ЗАКЛпер8").Range.Text = ФИОКорРук(ComboBox2.Text, False)

        End With

        Dim Name As String = ФамилияСотр & "(Заявление_Перевод)" & ".docx"
        Dim СохрЗак As New List(Of String)
        СохрЗак.AddRange(New String() {ComboBox1.Text & "\Заявление\" & Now.Year, Name})
        oWordDoc.SaveAs2(PathVremyanka & Name,,,,,, False)
        oWordDoc.Close(True)
        oWord.Quit(True)
        Конец(ComboBox1.Text & "\Заявление\" & Now.Year, Name, КодСотр, ComboBox1.Text, "\ZayavleniePerevod.docx", "Заявление_Перевод")
        massFTP2.Add(СохрЗак)

        Dim diskU As String

        Dim oWord1 As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc1 As Microsoft.Office.Interop.Word.Document
        oWord1 = CreateObject("Word.Application")
        oWord1.Visible = False

        Начало("KontraktPerevod.doc")
        oWordDoc1 = oWord1.Documents.Add(firthtPath & "\KontraktPerevod.doc")
        With oWordDoc1.Bookmarks
            .Item("К0").Range.Text = TextBox7.Text
            .Item("К1").Range.Text = MaskedTextBox1.Text
            .Item("К2").Range.Text = ComboBox2.Text
            .Item("К5").Range.Text = СотрФИОРод
            If ComboBox7.Text = "-" Then
                .Item("К8").Range.Text = Strings.LCase(ДобОконч(ComboBox4.Text))
            ElseIf ComboBox7.Text = "1" Or ComboBox7.Text = "2" Or ComboBox7.Text = "3" Or ComboBox7.Text = "4" Or ComboBox7.Text = "5" Or ComboBox7.Text = "6" Then
                .Item("К8").Range.Text = LCase(ДобОконч(ComboBox4.Text)) & " " & ComboBox7.Text & " разряда"
            Else
                .Item("К8").Range.Text = Strings.LCase(ДобОконч(ComboBox4.Text))
            End If
            .Item("К9").Range.Text = ComboBox5.Text & " " & СтавкаНов
            .Item("К10").Range.Text = ComboBox8.Text & " (" & ЧислПроп(ComboBox8.Text) & ") " & СклонГод
            .Item("К11").Range.Text = MaskedTextBox1.Text
            .Item("К12").Range.Text = ПродлКонтр
            .Item("К13").Range.Text = ComboBox2.Text
            .Item("К16").Range.Text = СотрАдрес
            .Item("К17").Range.Text = СотрПасп
            .Item("К18").Range.Text = СотрПаспВыд
            .Item("К23").Range.Text = ФИОКорРук(ComboBox2.Text, False)
            .Item("К25").Range.Text = ФИОКорРук(ComboBox2.Text, False)
            .Item("К26").Range.Text = TextBox3.Text
            .Item("К27").Range.Text = ЧислоПропис(TextBox3.Text)
            .Item("К28").Range.Text = ПоСовмИлиОсн
            .Item("К29").Range.Text = TextBox5.Text
            .Item("К30").Range.Text = TextBox4.Text
            .Item("К32").Range.Text = ЧислоПропис(TextBox4.Text)
            Select Case ComboBox3.Text
                Case "Руководители"
                    .Item("К38").Range.Text = "должностной инструкции"
                Case "Специалисты"
                    .Item("К38").Range.Text = "должностной инструкции"
            End Select

            Select Case ComboBox9.Text
                Case "График"
                    .Item("К33").Range.Text = "согласно графику работ"
                    .Item("К34").Range.Text = "согласно графику работ"
                    .Item("К35").Range.Text = "согласно графику работ"
                    Select Case CheckBox4.Checked
                        Case False
                            .Item("К36").Range.Text = "Суббота, Воскресенье"
                        Case True
                            .Item("К36").Range.Text = "согласно графику работ"
                            .Item("К37").Range.Text = "11.5. работнику устанавливается суммированный учет рабочего времени с учетным периодом - год."
                    End Select

                Case "ПВТР"
                    .Item("К33").Range.Text = "согласно правил внутреннего трудового распорядка"
                    .Item("К34").Range.Text = "согласно правил внутреннего трудового распорядка"
                    .Item("К35").Range.Text = "согласно правил внутреннего трудового распорядка"
                    Select Case CheckBox4.Checked
                        Case False
                            .Item("К36").Range.Text = "согласно графику работ"
                        Case True
                            .Item("К36").Range.Text = "согласно графику работ"
                            .Item("К37").Range.Text = "11.5. работнику устанавливается суммированный учет рабочего времени с учетным периодом - год."
                    End Select

                Case "Задать"
                    .Item("К33").Range.Text = ComboBox10.Text
                    .Item("К34").Range.Text = TextBox11.Text
                    .Item("К35").Range.Text = TextBox12.Text

                    Select Case CheckBox4.Checked
                        Case False
                            .Item("К36").Range.Text = "Суббота, Воскресенье"
                        Case True
                            .Item("К36").Range.Text = "согласно графику работ"
                            .Item("К37").Range.Text = "11.5. работнику устанавливается суммированный учет рабочего времени с учетным периодом - год."
                    End Select
            End Select

            .Item("К39").Range.Text = ФормаСобстПолн
            .Item("К40").Range.Text = ComboBox1.Text
            .Item("К41").Range.Text = ДобОконч(ДолжРуков)
            .Item("К42").Range.Text = ФИОРукРодПад

            If Not ComboBox1.Text = "Итал Гэлэри Плюс" Then
                .Item("К43").Range.Text = ОснованиеДейств
            Else
                .Item("К51").Range.Text = ""
            End If
            .Item("К44").Range.Text = МестоРаб
            .Item("К45").Range.Text = ФИОКор
            .Item("К46").Range.Text = СборноеРеквПолн
            .Item("К47").Range.Text = Year(Now).ToString
            .Item("К48").Range.Text = TextBox9.Text
            If TextBox10.Text = "" Or TextBox10.Text = "НЕТ" Then
                .Item("К49").Range.Text = ""
            Else
                .Item("К49").Range.Text = "и " & TextBox10.Text & "-го (аванс) "
            End If

            'If ComboBox10.Text = "1.0" Then
            .Item("К50").Range.Text = "1 ставка"
            'Else
            '    .Item("К50").Range.Text = ComboBox10.Text & " ставки"
            'End If
            Select Case Пол
                Case "М"
                    .Item("К52").Range.Text = "ним"
                Case "Ж"
                    .Item("К52").Range.Text = "ней"
            End Select

        End With


        Dim Name1 As String = TextBox7.Text & " " & ФамилияСотр & " (контракт_перевод)" & ".doc"
        Dim СохрКонтр As New List(Of String)
        СохрКонтр.AddRange(New String() {ComboBox1.Text & "\Контракт\" & Now.Year, Name1})
        oWordDoc1.SaveAs2(PathVremyanka & Name1,,,,,, False)
        oWordDoc1.Close(True)
        oWord1.Quit(True)
        Конец(ComboBox1.Text & "\Контракт\" & Now.Year, Name1, КодСотр, ComboBox1.Text, "\KontraktPerevod.doc", "контракт_перевод")
        massFTP2.Add(СохрКонтр)


        Dim oWord2 As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc2 As Microsoft.Office.Interop.Word.Document
        oWord2 = CreateObject("Word.Application")
        oWord2.Visible = False

        Начало("PrikazNaPerevod.doc")
        oWordDoc2 = oWord2.Documents.Add(firthtPath & "\PrikazNaPerevod.doc")

        With oWordDoc2.Bookmarks
            .Item("П1").Range.Text = MaskedTextBox1.Text

            If TextBox8.Text <> "" Then
                .Item("П2").Range.Text = TextBox6.Text & "-к-" & TextBox8.Text
            Else
                .Item("П2").Range.Text = TextBox6.Text & "-к"
            End If

            .Item("П3").Range.Text = ФИОКорРук(ФИОСторДляЗаявл, False)
            .Item("П6").Range.Text = СотрФИОРод
            If ComboBox7.Text = "-" Then
                .Item("П10").Range.Text = Strings.LCase(ДобОконч(ComboBox4.Text))
            ElseIf ComboBox7.Text = "1" Or ComboBox7.Text = "2" Or ComboBox7.Text = "3" Or ComboBox7.Text = "4" Or ComboBox7.Text = "5" Or ComboBox7.Text = "6" Then
                .Item("П10").Range.Text = Strings.LCase(ДобОконч(ComboBox4.Text)) & " " & ComboBox7.Text & " разряда"
            Else
                .Item("П10").Range.Text = Strings.LCase(ДобОконч(ComboBox4.Text))
            End If

            Select Case СтарРазряд
                Case = ""
                    .Item("П9").Range.Text = Strings.LCase(ДобОконч(ДолжСотр))
                Case = "-"
                    .Item("П9").Range.Text = Strings.LCase(ДобОконч(ДолжСотр))
                Case <> ""
                    .Item("П9").Range.Text = Strings.LCase(ДобОконч(ДолжСотр)) & " " & СтарРазряд & " разряда"
            End Select

            If CheckBox2.Checked = True Then
                .Item("П11").Range.Text = MaskedTextBox3.Text & "г. на " & ComboBox5.Text & " " & Склонение(ComboBox5.Text) & " по совместительству"
            Else
                .Item("П11").Range.Text = MaskedTextBox3.Text & "г. на " & ComboBox5.Text & " " & Склонение(ComboBox5.Text)
            End If


            .Item("П12").Range.Text = ComboBox8.Text & " " & Склонение2(ComboBox8.Text)
            .Item("П13").Range.Text = MaskedTextBox1.Text
            .Item("П14").Range.Text = ПродлКонтр
            .Item("П17").Range.Text = ФИОКорРук(ФИОСторДляЗаявл, False)
            .Item("П20").Range.Text = TextBox7.Text
            .Item("П21").Range.Text = MaskedTextBox1.Text
            .Item("П22").Range.Text = ФИОКорРук(ComboBox2.Text, False)
            .Item("П25").Range.Text = ФормаСобстПолн
            If ФормаСобстПолн = "Индивидуальный предприниматель" Then
                .Item("П26").Range.Text = ComboBox1.Text
            Else
                .Item("П26").Range.Text = " «" & ComboBox1.Text & "» "
            End If

            .Item("П27").Range.Text = ЮрАдрес
            .Item("П28").Range.Text = УНП
            .Item("П29").Range.Text = РасСчет
            .Item("П30").Range.Text = АдресБанка
            .Item("П31").Range.Text = БИК
            .Item("П33").Range.Text = ЭлАдрес
            .Item("П34").Range.Text = КонтТелефон
            '.Item("П35").Range.Text = МестоРаб
            If Not ComboBox1.Text = "Итал Гэлэри Плюс" Then
                If ДолжРуков = "Индивидуальный предприниматель" Then
                    .Item("П36").Range.Text = ДолжРуков
                    .Item("П37").Range.Text = ""
                Else
                    .Item("П36").Range.Text = ДолжРуков & " " & ФормаСобствКор
                    .Item("П37").Range.Text = "«" & ComboBox1.Text & "»"
                End If
            End If
            .Item("П38").Range.Text = ФИОКор
        End With

        Dim Name2 As String = TextBox7.Text & " " & ФамилияСотр & " (контракт_перевод)" & ".doc"
        Dim СохрПрик As New List(Of String)
        СохрПрик.AddRange(New String() {ComboBox1.Text & "\ПриказПеревод\" & Now.Year, Name2})
        oWordDoc2.SaveAs2(PathVremyanka & Name2,,,,,, False)
        oWordDoc2.Close(True)
        oWord2.Quit(True)
        Конец(ComboBox1.Text & "\ПриказПеревод\" & Now.Year, Name2, КодСотр, ComboBox1.Text, "\PrikazNaPerevod.doc", "ПриказПеревод")
        massFTP2.Add(СохрПрик)
    End Sub
    Private Sub Чист()
        StrSql = ""
        ds.Clear()

    End Sub
    Private Sub СборДанОрганиз()
        Dim ds = From x In dtObjectObshepitaAll Where x.Item("АдресОбъекта") = ComboBox6.Text And x.Item("НазвОрг") = ComboBox1.Text Select x
        'StrSql = "SELECT ТипОбъекта, НазОбъекта From [ОбъектОбщепита] Where АдресОбъекта = '" & ComboBox6.Text & "' AND НазвОрг= '" & ComboBox1.Text & "'"
        'ds = Selects(StrSql)

        Тип = Strings.Trim(Strings.LCase(ds(0).Item("ТипОбъекта").ToString))
        Название = """" & ds(0).Item("НазОбъекта").ToString & ""","

        'сборка данных для доков со стороны руководства

        Dim ds1 = From c In dtClientAll Where c.Item("НазвОрг") = ComboBox1.Text Select c

        '        StrSql = "Select ФормаСобств, УНП, ЮрАдрес, 
        'Банк, БИКБанка, АдресБанка, Отделение,РасчСчетРубли, 
        'ДолжнРуководителя, ФИОРуководителя, ОснованиеДейств, ФИОРукРодПадеж,
        'КонтТелефон, ЭлАдрес, ФИОРукДатПадеж, РукИП
        'From Клиент
        'Where НазвОрг = '" & ComboBox1.Text & "'"
        '        ds = Selects(StrSql)
        Dim РуковИП As String
        If ds1(0).Item("РукИП") = "True" Then
            РуковИП = "ИП "
        Else
            РуковИП = ""
        End If

        ФормаСобстПолн = ds1(0).Item("ФормаСобств").ToString
        ДолжРуков = ds1(0).Item("ДолжнРуководителя").ToString
        ФИОРукРодПад = РуковИП & ds1(0).Item("ФИОРукРодПадеж").ToString
        ОснованиеДейств = ds1(0).Item("ОснованиеДейств").ToString
        МестоРаб = Тип & " " & Название & " " & ComboBox6.Text

        'короткое фио 
        Dim nm As String = ds1(0).Item("ФИОРуководителя").ToString
        Dim nm0 As Integer = Len(ds1(0).Item("ФИОРуководителя").ToString)
        Dim nm1 As String = Strings.Left(nm, InStr(nm, " "))
        Dim nm2 As Integer = Len(nm1)
        Dim nm3 As String = Strings.Right(nm, (nm0 - nm2))
        Dim nm31 As Integer = Len(nm3)
        Dim nm4 As String = Strings.UCase(Strings.Left(Strings.Left(nm3, InStr(nm3, " ")), 1))
        Dim nm41 As Integer = Len(Strings.Left(nm3, InStr(nm3, " ")))
        Dim nm5 As String = Strings.UCase(Strings.Left(Strings.Right(nm3, nm31 - nm41), 1))
        Dim nm6 = Strings.Left(ds1(0).Item("ФИОРукДатПадеж").ToString, InStr(ds1(0).Item("ФИОРукДатПадеж").ToString, " "))



        ФИОКор = РуковИП & nm1 & " " & nm4 & "." & nm5 & "."
        ФИОКорРукДат = РуковИП & nm6 & " " & nm4 & "." & nm5 & "."
        УНП = ds1(0).Item("УНП").ToString
        КонтТелефон = ds1(0).Item("КонтТелефон").ToString
        ЮрАдрес = ds1(0).Item("ЮрАдрес").ToString
        РасСчет = ds1(0).Item("РасчСчетРубли").ToString
        Банк = ds1(0).Item("Банк").ToString
        БИК = ds1(0).Item("БИКБанка").ToString
        АдресБанка = ds1(0).Item("АдресБанка").ToString
        ЭлАдрес = ds1(0).Item("ЭлАдрес").ToString


        'сокращенное название орг
        Dim h = dtformft.Select("ПолноеНазвание = '" & ds1(0).Item("ФормаСобств").ToString & "'")

        'Dim StrSql9 As String = "Select Сокращенное From ФормаСобств Where ПолноеНазвание = '" & ds1(0).Item(0).ToString & "'"
        'Dim c9 As New OleDbCommand With {
        '        .Connection = conn,
        '        .CommandText = StrSql9
        '    }
        'Dim ds9 As New DataSet
        'Dim da9 As New OleDbDataAdapter(c9)
        'da9.Fill(ds9, "Ст")

        ФормаСобствКор = h(0).Item("Сокращенное").ToString

        СборноеРеквПолн = ФормаСобствКор & " """ & ComboBox1.Text & """ " & ds1(0).Item("ЮрАдрес").ToString & " IBAN " _
        & ds1(0).Item("РасчСчетРубли").ToString & " в " & ds1(0).Item("АдресБанка").ToString & " " & " БИК " _
        & ds1(0).Item("БИКБанка").ToString & " УНП " & ds1(0).Item("УНП").ToString
    End Sub
    Private Sub TextBox7_LostFocus(sender As Object, e As EventArgs) Handles TextBox7.LostFocus
        Dim pl As String
        If TextBox7.Text <> "" Then
            Try
                Dim i As Integer = CInt(TextBox7.Text)
                Select Case i
                    Case < 10
                        pl = Str(i)
                        TextBox7.Text = "00" & i

                    Case 10 To 99
                        pl = Str(i)
                        TextBox7.Text = "0" & i
                End Select
            Catch ex As Exception
                TextBox7.Text = "б.н"
            End Try
        Else
            TextBox7.Text = "б.н"
        End If
    End Sub

    Private Sub TextBox6_LostFocus(sender As Object, e As EventArgs) Handles TextBox6.LostFocus
        Dim pl As String
        If TextBox6.Text <> "" Then
            Try
                Dim i As Integer = CInt(TextBox6.Text)
                Select Case i
                    Case < 10
                        pl = Str(i)
                        TextBox6.Text = "00" & i

                    Case 10 To 99
                        pl = Str(i)
                        TextBox6.Text = "0" & i
                End Select
            Catch ex As Exception
                TextBox6.Text = "б.н"
            End Try
        Else
            TextBox6.Text = "б.н"
        End If
    End Sub

    Private Sub Перевод_Closed(sender As Object, e As EventArgs) Handles MyBase.Closed
        Me.Close()
        Me.Dispose()
    End Sub
End Class