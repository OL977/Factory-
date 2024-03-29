﻿Option Explicit On
Imports System.Data.OleDb
Imports System.Threading
Imports MySql.Data.MySqlClient
Imports System.Management
Imports System.Data.SqlClient
Imports System.Linq
Imports System.Linq.Dynamic
Imports System.IO
Imports System.Globalization
Imports System.ComponentModel
Imports Zidium


'Imports unvell.ReoGrid
'Imports System.IO
'Imports Microsoft.Office.Interop.Word
Public Class Прием

    Public ds, dsGeneral As DataTable
    Public Примечани As String
    Dim StrSql As String
    Dim Должность, a, n, w, Разряд, Клиент, CorName, CorOtch, rub, Год, СрокКонтр, Ставка, РДОкопейки, ОргдляДокум, but2cl, результат, СохрЛемел, СохрПинфуд As String
    Dim Заявление() As String, Контракт() As String, Прием() As String, Курьер() As String,
        МатОтвет() As String, MassДолжн() As String, Приказ() As String, MassДогПодрОбяз() As Integer
    Dim arrtbox As New Dictionary(Of String, String)
    Dim arrtcom As New Dictionary(Of String, String)
    Dim arrtmask As New Dictionary(Of String, String)
    Dim dad As Date
    Dim РДОрубли, sf, ФондОТ As Double
    Dim mass() As String
    Dim massFTP As New ArrayList()
    Dim СохрЗакFTP As New List(Of String)()
    Dim СохрКонтрFTP As New List(Of String)()
    Dim СохрПрикFTP As New List(Of String)()
    Dim ИнстрFTP As New List(Of String)()
    Dim СохрДогПодрFTP As New List(Of String)()
    Dim ПровВходаCom19 As Boolean = False, ПровВходаCom8 As Boolean = False
    Dim СохрЗак, СохрПрик, СохрКонтр, НПриказа, surName, surNameAll, Знач, СписОбязан, ПовышениеПроц, ТарифнаяСт, Отделы, Инстр, ИнстрП, ДогПодНомСтарДог As String
    Dim fx, КодСотрудника, очПоля, ПрКонт, ПрПодр, rz, IDLДогПодрОбяз, hscol, hscol2, hscol22, dfe, ПровИнстр As Integer

    Dim ФормаСобстПолн, ЭлАдрес, ФИОКорРукДат, Банк, БИК, АдресБанка, РасСчет,
        ЮрАдрес, УНП, ДолжРуков, ФИОРукРодПад, ОснованиеДейств, МестоРаб, ФИОКор,
        ФормаСобствКор, СборноеРеквПолн, ДолжРуковРодПад, ДолжРуковВинПад, КонтТелефон, ДПодНом, inp, СтПосле, ПроцПосле, ДатРожд, СохрАмасейл, СохрПрикЛемел As String
    Dim CombBox7 As Integer = 0, mlk As Integer = 0, fgm As Integer = 0
    Dim IDsot1 As Integer
    Public v As Boolean = False, tabcon As Boolean = False
    Dim IDso, txtbx46l As Integer
    Dim mo As Object
    Dim массивДогПодр() As String, массив2() As String
    Dim ДолжСОконч, СтавкаНов, СклонГод, СрКонтПроп, ПоСовмИлиОсн, ПоСовмПриказ As String
    Dim combx1, combx28, combx8, combx9, combx7, combx10, combx11, combx19, combx16, combx12, combx15, combx18, combx14, combx13, combx3, combx4, combx5, combx6 As String
    Friend cmb8, cmb19, cmb18, cmb26, cmb28, txtbxD46, combxS19, txtbx38, txtbx44, txtbx47, txtbx48, txtbx6, txtbx49, txtbx50, mskbx3, txt1, txt2, txt3 As String
    Private Shared Applications As List(Of Microsoft.Office.Interop.Word.Application) = New List(Of Microsoft.Office.Interop.Word.Application)
    Dim f, f1, f2, f3 As Task
    Dim TskList As List(Of Task)
    Dim TskArr() As Task
    'Private Declare Auto Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
    Private Delegate Sub CombxDel1()
    Private Delegate Sub Txtb46()
    Private Delegate Sub txtbx1()
    Private Delegate Sub txtb38()
    Private Delegate Sub comb38()
    Private Delegate Sub comb18()
    Private Delegate Sub comb8()
    Private Delegate Sub comb19()
    Private Delegate Sub txtbx46()
    Dim Dtxt46 As Double
    Dim ОбязН As String
    Dim ВсплПриЗагрНов As Thread
    Dim ДокКонтрПерем As String
    Dim К33, К34, К35, К36, К37 As String
    Dim Поток As New Thread(AddressOf ДанныеКлиентаДогПодряда)
    Dim Поток1 As New Thread(AddressOf НалогиИОбязанДогПодряда)
    Public РазрИзменКонтр
    Dim Решение As String
    Dim idДолжность, idОбязанность As Integer



    Private Sub ДанИзБазы()
        If ComboBox20.InvokeRequired Or ComboBox21.InvokeRequired Then
            Me.Invoke(New comb38(AddressOf ДанИзБазы))
        Else
            Me.ComboBox20.Items.Clear()
            Me.ComboBox21.Items.Clear()
            Dim ut() As Object = {Now.Year - 2, Now.Year - 1, Now.Year}
            ComboBox20.Items.AddRange(ut)
            ComboBox21.Items.AddRange(ut)
        End If
    End Sub

    Private Sub Com1()
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

    End Sub
    Private Sub Прием_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If ИмяКомп = "OLEGLAPTOP" Then
            Button11.Visible = True
        Else
            Button11.Visible = False
        End If


        КонтрПровИндивид = {}
        КонтрПровИндивид = {"Амасейлс", "ЛемеЛ Лабс"}
        Год = Year(Now)

        Parallel.Invoke(Sub() ДанИзБазы())
        Parallel.Invoke(Sub() Com1())
        'Dim cm1 As Task = New Task(AddressOf Com1)
        'cm1.Start()

        GroupBox26.Visible = False
        GroupBox27.Visible = False
        ComboBox26.Visible = False

        MaskedTextBox3.Text = Now
        MaskedTextBox4.Text = Now
        Dim dad As Date = CDate(MaskedTextBox4.Text)

        MaskedTextBox5.Text = dad.AddMonths(12)
        Dim dad2 As Date = CDate(MaskedTextBox5.Text)
        MaskedTextBox5.Text = dad2.AddDays(-1)


        TabControl1.TabPages.Remove(TabPage3)
        TabControl1.TabPages.Remove(TabPage4)
        'Com1()

        CheckBox26.Visible = False
        CheckBox23.Visible = False
        Me.Label56.Visible = False
        Me.Label55.Visible = False
        CheckBox27.Enabled = False

        TextBox42.Text = Now.ToShortDateString & "г."


        Me.TextBox43.Text = ""
        Me.TextBox47.Text = ""
        ComboBox19.Enabled = False

        dtShtatnoeOtdely()

        'изменяем вкладки в Tabcontrol
        TabControl1.DrawMode = TabDrawMode.OwnerDrawFixed
        For Each tg As TabPage In TabControl1.TabPages
            tg.BackColor = Color.Gainsboro
        Next



    End Sub

    Private Sub ТарифнаяСтавка()
        'Соед(0)
        Me.ComboBox7.Enabled = True
        Label47.Enabled = True
        Label79.Enabled = True

        StrSql = ""
        StrSql = "SELECT ШтОтделы.Отделы, ШтСвод.Должность, ШтСвод.Разряд, ШтСвод.ТарифнаяСтавка,
ШтСвод.ПовышениеПроц, ШтСвод.ТарСтПослеИспСрока, ПовПроцПослеИспСрока
FROM ШтОтделы INNER JOIN ШтСвод ON ШтОтделы.Код = ШтСвод.Отдел
WHERE ШтОтделы.Отделы='" & Отдел & "' AND ШтСвод.Должность='" & Должность & "' AND ШтОтделы.Клиент='" & Клиент & "'"
        ds.Clear()
        ds = Selects(StrSql)

        'Соед(0)

        Отделы = ds.Rows(0).Item(0).ToString
        ТарифнаяСт = ds.Rows(0).Item(3).ToString
        ПовышениеПроц = ds.Rows(0).Item(4).ToString()

        СтПосле = ds.Rows(0).Item(5).ToString
        ПроцПосле = ds.Rows(0).Item(6).ToString



        If Должность = "Кладовщик" Then Me.ComboBox7.Enabled = True

        Dim ghfd(ds.Rows.Count - 1) As String
        Dim ghfr As Integer
        For i As Integer = 0 To ds.Rows.Count - 1
            ghfd(i) = ds.Rows(i).Item(2).ToString
            If ghfd(i) = "1" Or ghfd(i) = "2" Or ghfd(i) = "3" Or ghfd(i) = "4" Or ghfd(i) = "5" Or ghfd(i) = "-" Or ghfd(i) = "6" Then
                ghfr = 1

            End If
        Next




        'Dim ghfd1 As String = ds.Rows(1).Item(2).ToString

        If ds.Rows(0).Item(1) <> "" And ghfr = 1 Then
            ПовышениеПроц = ds.Rows(0).Item(4).ToString()
        Else
            СвертывРазр(ds)
        End If
    End Sub
    Private Sub Очистка()
        TextBox33.Text = ""
        TextBox44.Text = ""
        TextBox43.Text = ""
        TextBox46.Text = ""
        TextBox48.Text = ""
        TextBox47.Text = ""
    End Sub
    Private Sub СвертывРазр(ByVal ds As DataTable)
        Очистка()
        Me.ComboBox7.Enabled = False
        'Label47.Enabled = False
        'Label79.Enabled = False

        TextBox46.Text = ds.Rows(0).Item(4).ToString()

        Dim dstbl As String = ds.Rows(0).Item(3).ToString

        If dstbl <> "." Then dstbl = Replace(dstbl, ".", ",")
        If dstbl <> "," Then
            sf = Nothing
            sf = CType(dstbl, Double)
            Dim sfd As String = CType(sf, String)
            Dim ДлНач As Integer = sfd.Length
            TextBox33.Text = Math.Floor(sf)
            Dim Дл As Integer = TextBox33.TextLength
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

            TextBox44.Text = vm
        Else
            TextBox33.Text = ds.Rows(0).Item(0).ToString
        End If
    End Sub
    Private Sub com19collection()

        If ComboBox19.InvokeRequired Or ComboBox26.InvokeRequired Then
            Me.Invoke(New comb19(AddressOf Ускорен))
        Else
            If ПровВходаCom19 = False Then
                ПровВходаCom19 = True
                ComboBox19.AutoCompleteCustomSource.Clear()
                ComboBox19.Items.Clear()
                ComboBox26.Items.Clear()


                'dtSotrudnikiAll.DefaultView.Sort = "ФИОСборное" & " ASC"            'по возрастанию
                dtSotrudnikiAll.Select("", "ФИОСборное")
                'Parallel.ForEach(Of DataRow, ds.AsEnumerable(), Fun(Of ()()()()
                Dim var1 = From x In dtSotrudnikiAll.Rows Where x.Item("НазвОрганиз") = Клиент Select x 'рабочий linq для заполнения комбобоксов
                Dim var = From x In dtSotrudnikiAll.Rows Where x.Item("НазвОрганиз") = Клиент Order By "ФИОСборное" Select x   'рабочий linq для заполнения комбобоксов  и order by
                'Dim var3 = dtSotrudnikiAll.Rows.AsQueryable.Where("НазвОрганиз" = Клиент).OrderBy()
                'var3 = var3.AsQueryable.Where()
                'var1 = var1.OrderBy(Function(c) c.ФИОСборное)
                For Each r As DataRow In var1
                    ComboBox19.AutoCompleteCustomSource.Add(r.Item(0).ToString())
                    ComboBox19.Items.Add(r("ФИОСборное").ToString)
                    'Me.ComboBox19.Items.Add(r(1).ToString)
                    ComboBox26.Items.Add(r("КодСотрудники").ToString)
                Next
                ComboBox19.Text = ""
            End If
        End If
    End Sub

    Private Sub Ускорен()

        If ComboBox19.InvokeRequired Or ComboBox26.InvokeRequired Then
            Me.Invoke(New comb19(AddressOf Ускорен))
        Else
            If ПровВходаCom19 = False Then
                ПровВходаCom19 = True
                ComboBox19.AutoCompleteCustomSource.Clear()
                ComboBox19.Items.Clear()
                ComboBox26.Items.Clear()
                Dim var As List(Of Сотрудники)
                Using dbcx As New DbAll1DataContext
                    var = (From x In dbcx.Сотрудники.AsEnumerable
                           Where x.НазвОрганиз = ComboBox1.Text
                           Order By x.ФИОСборное
                           Select x).ToList()
                End Using


                'Dim var1 = From x In dtSotrudnikiAll.Rows Where x.Item("НазвОрганиз") = Клиент Select x 'рабочий linq для заполнения комбобоксов
                'Dim var = From x In dtSotrudnikiAll.Rows Where x.Item("НазвОрганиз") = Клиент Order By x.Item("ФИОСборное") Select x   'рабочий linq для заполнения комбобоксов  и order by

                For Each r In var
                    ComboBox19.AutoCompleteCustomSource.Add(r.ФИОСборное)
                    ComboBox19.Items.Add(Trim(r.ФИОСборное & "" & r.ТипОтношения))
                    ComboBox26.Items.Add(r.КодСотрудники)
                Next


                'оригинал до 24.12.19
                'For Each r As DataRow In var
                '    ComboBox19.AutoCompleteCustomSource.Add(r.Item(0).ToString())
                '    ComboBox19.Items.Add(r("ФИОСборное").ToString)
                '    ComboBox26.Items.Add(r("КодСотрудники").ToString)
                'Next






                ComboBox19.Text = ""
            End If


        End If




        If ComboBox8.InvokeRequired Then
            Me.Invoke(New comb8(AddressOf Ускорен))
        Else
            If ПровВходаCom8 = False Then
                ПровВходаCom8 = True
                Dim StrSql As String = "SELECT DISTINCT ШтОтделы.Отделы FROM Клиент INNER JOIN ШтОтделы ON Клиент.НазвОрг = ШтОтделы.Клиент WHERE Клиент.НазвОрг='" & Клиент & "'"
                ds = Selects(StrSql)
                ComboBox8.AutoCompleteCustomSource.Clear()
                ComboBox8.Items.Clear()
                For Each r As DataRow In ds.Rows
                    Me.ComboBox8.AutoCompleteCustomSource.Add(r.Item(0).ToString())
                    Me.ComboBox8.Items.Add(r(0).ToString)
                Next
                ComboBox8.Text = ""
            End If

        End If


        'StrSql = ""
        'StrSql = "SELECT ФИОСборное, КодСотрудники FROM Сотрудники WHERE НазвОрганиз='" & Клиент & "' ORDER BY ФИОСборное "
        'dsGeneral = Selects(StrSql)

        'Dim m = dtSotrudnikiAll.Rows.Count



        'Dim var() = dtSotrudnikiAll.Select("НазвОрганиз='" & Клиент & "'")



        'Dim var = From x In dtSotrudnikiAll.Rows Where x.Item("НазвОрганиз") = Клиент Select x 'рабочий linq для заполнения комбобоксов
        'If ComboBox19.InvokeRequired Or ComboBox26.InvokeRequired Then
        '    Me.Invoke(New comb19(AddressOf Ускорен))
        'Else
        '    ComboBox19.AutoCompleteCustomSource.Clear()
        '    ComboBox19.Items.Clear()
        '    ComboBox26.Items.Clear()

        '    'Parallel.ForEach(Of DataRow, ds.AsEnumerable(), Fun(Of ()()()()

        '    For Each r As DataRow In var
        '        ComboBox19.AutoCompleteCustomSource.Add(r.Item(0).ToString())
        '        ComboBox19.Items.Add(r("ФИОСборное").ToString)
        '        'Me.ComboBox19.Items.Add(r(1).ToString)
        '        ComboBox26.Items.Add(r("КодСотрудники").ToString)
        '    Next
        '    ComboBox19.Text = ""
        'End If
    End Sub


    Public Sub Com1sel()

        Клиент = ""
        Клиент = ComboBox1.Text
        'Dim f As Boolean = Await Ускорен()

        Dim df1 As New Thread(AddressOf Ускорен) 'асинхронно
        df1.IsBackground = True
        df1.Start()


        Dim df As New Thread(AddressOf Ускор1)
        df.IsBackground = True
        df.Start()

        If ПровИндивидКонтр(ComboBox1.Text) = True Then
            GroupBox26.Visible = True
            GroupBox27.Visible = True
        Else
            GroupBox26.Visible = False
            GroupBox27.Visible = False
        End If


    End Sub
    Private Sub Ускор1()
        Dim StrSql As String = "SELECT DISTINCT КарточкаСотрудника.ДатаЗарплаты, КарточкаСотрудника.ДатаАванса
FROM (Клиент INNER JOIN Сотрудники ON Клиент.НазвОрг = Сотрудники.НазвОрганиз) INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр
WHERE Клиент.НазвОрг='" & Клиент & "' AND Сотрудники.НазвОрганиз ='" & Клиент & "'"
        Dim ds As DataTable = Selects(StrSql)
        Try
            TextBox40.Text = ds.Rows(0).Item(0)
            TextBox56.Text = ds.Rows(0).Item(1)

        Catch ex As Exception

        End Try
    End Sub
    Private Sub Обьект()




        If ComboBox18.InvokeRequired Then
            Me.Invoke(New comb18(AddressOf Обьект))
        Else

            Dim ds As DataTable = Selects(StrSql:="SELECT АдресОбъекта FROM ОбъектОбщепита WHERE НазвОрг='" & Клиент & "'")
            ComboBox18.Items.Clear()
            For Each r As DataRow In ds.Rows
                ComboBox18.Items.Add(r(0).ToString)
            Next

            Try
                combx18 = ds.Rows(0).Item(0).ToString

            Catch ex As Exception
                MessageBox.Show("При регистрации организации не создан объект общепита!" & vbCrLf & "Выберите объект заново или поставьте галочку отметив текст 'Нет объекта общепита'", Рик)

                Exit Sub
            End Try
        End If

    End Sub
    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        ОчисткаМаяковAsync()
        ClAll()

        Клиент = ComboBox1.Text
        combx1 = ComboBox1.Text
        ДогПодрВклЧекбокс5 = False
        'ComboBox18.Text = combx18
        Dim f As New Thread(AddressOf com1selcombx1) With {
            .IsBackground = True
        }
        f.SetApartmentState(ApartmentState.STA)
        f.Start()

        Dim f1 As New Thread(AddressOf Обьект) With {
            .IsBackground = True
        }
        f1.SetApartmentState(ApartmentState.STA)
        f1.Start()

        'Dim go As New Thread(AddressOf СборДаннОрганиз) 'сбор данных организации при выборе организации
        'go.IsBackground = True
        'go.SetApartmentState(ApartmentState.STA)
        'go.Start()

        Parallel.Invoke(Sub() СборДаннОрганиз()) 'сбор данных организации при выборе организации




        Com1sel()
        'Обьект()



    End Sub
    Private Sub com1selcombx1()

        'If mlk = 0 Then
        '    Dim l As New Thread(AddressOf ClAll)
        '    l.IsBackground = True
        '    l.SetApartmentState(ApartmentState.STA)
        '    l.Start()
        'End If


        Dim df1 As New Thread(AddressOf Ускорен) 'асинхронно
        df1.IsBackground = True
        df1.SetApartmentState(ApartmentState.STA)
        df1.Start()


        Dim df As New Thread(AddressOf Ускор1)
        df.IsBackground = True
        df.SetApartmentState(ApartmentState.STA)
        df.Start()











    End Sub
    'Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs)

    '    If ComboBox1.Text = "" Then
    '        MsgBox("Выберите организацию",, "ООО РикКонсалтинг")
    '        Me.ComboBox1.Focus()

    '        Exit Sub
    '    End If

    '    'If ComboBox1.Text <> "" And CheckBox5.Checked = True And ComboBox19.SelectedItem = "" Then
    '    '    CheckBox5.Checked = False
    '    'End If


    '    sender.text = StrConv(sender.text, VbStrConv.ProperCase)
    '    sender.SelectionStart = sender.text.Length
    '    TextBox6.Text = TextBox1.text

    'End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        sender.text = StrConv(sender.text, VbStrConv.ProperCase)
        sender.SelectionStart = sender.text.Length
        Me.TextBox5.Text = Me.TextBox2.Text

    End Sub


    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        sender.text = StrConv(sender.text, VbStrConv.ProperCase)
        sender.SelectionStart = sender.text.Length
        Me.TextBox4.Text = Me.TextBox3.Text

    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        sender.text = StrConv(sender.text, VbStrConv.ProperCase)
        sender.SelectionStart = sender.text.Length
        Me.TextBox34.Text = Me.TextBox6.Text
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        sender.text = StrConv(sender.text, VbStrConv.ProperCase)
        sender.SelectionStart = sender.text.Length
        Me.TextBox11.Text = Me.TextBox5.Text
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        sender.text = StrConv(sender.text, VbStrConv.ProperCase)
        sender.SelectionStart = sender.text.Length
        Me.TextBox10.Text = Me.TextBox4.Text
    End Sub

    Private Sub TextBox24_TextChanged(sender As Object, e As EventArgs) Handles TextBox24.TextChanged
        sender.text = StrConv(sender.text, VbStrConv.ProperCase)
        sender.SelectionStart = sender.text.Length
    End Sub

    Private Sub TextBox25_TextChanged(sender As Object, e As EventArgs) Handles TextBox25.TextChanged
        sender.text = StrConv(sender.text, VbStrConv.ProperCase)
        sender.SelectionStart = sender.text.Length
    End Sub

    Private Sub TextBox27_TextChanged(sender As Object, e As EventArgs) Handles TextBox27.TextChanged
        sender.text = StrConv(sender.text, VbStrConv.ProperCase)
        sender.SelectionStart = sender.text.Length
    End Sub

    Private Sub TextBox30_TextChanged(sender As Object, e As EventArgs) Handles TextBox30.TextChanged
        sender.text = StrConv(sender.text, VbStrConv.ProperCase)
        sender.SelectionStart = sender.text.Length
    End Sub

    Private Sub TextBox32_TextChanged(sender As Object, e As EventArgs) Handles TextBox32.TextChanged
        sender.text = StrConv(sender.text, VbStrConv.ProperCase)
        sender.SelectionStart = sender.text.Length
    End Sub

    Private Sub TextBox36_TextChanged(sender As Object, e As EventArgs) Handles TextBox36.TextChanged
        sender.text = StrConv(sender.text, VbStrConv.ProperCase)
        sender.SelectionStart = sender.text.Length
    End Sub

    Private Sub TextBox33_TextChanged(sender As Object, e As EventArgs) Handles TextBox33.TextChanged

        'TextBox43.Text = ""

        'If Me.TextBox33.Text <> "" Then
        '    rub = Пропись(Me.TextBox33.Text)
        '    TextBox43.Text = rub & "бел.руб. 00 копеек"
        'ПропОклад()
        'ElseIf TextBox44.Text = "" Then
        '    TextBox43.Text = rub & "бел.рублей"

        'End If

    End Sub
    Public Sub ПропОклад()

        'Await Task.Delay(0)

        If TextBox33.Text = "" Or TextBox33.Text = "0" Then

            Exit Sub
        End If
        Dim sfd As String
        Dim valr2, valr As Double
        If sf Then
            valr2 = sf
        Else
            sfd = TextBox33.Text & "," & TextBox44.Text
            valr2 = sfd.Replace(".", ",")
        End If

        If TextBox46.Text = "" And CheckBox5.Checked = False Then
            Dim StrSql As String = "Select  ШтСвод.ПовышениеПроц
From ШтОтделы INNER Join ШтСвод On ШтОтделы.Код = ШтСвод.Отдел
Where ШтОтделы.Отделы ='" & ComboBox8.Text & "' AND ШтСвод.Должность = '" & ComboBox9.Text & "' AND ШтСвод.Разряд='" & ComboBox7.Text & "' AND ШтОтделы.Клиент = '" & ComboBox1.Text & "'"
            Dim ds As DataTable = Selects(StrSql)
            If ds.Rows(0).Item(0).ToString = "" Then
                valr = Math.Round((valr2 + (valr2 * 0 / 100)), 2)
            Else
                valr = Math.Round((valr2 + (valr2 * CType(ds.Rows(0).Item(0), Double) / 100)), 2)
            End If

        Else
            valr = Math.Round((valr2 + (valr2 * CType(TextBox46.Text.Replace(".", ","), Double) / 100)), 2)
        End If



        Select Case ComboBox10.Text
            Case ""
                ФондОТ = valr
                Exit Select
            Case "0.25"
                ФондОТ = valr * 0.25
            Case "0.5"
                ФондОТ = valr * 0.5
            Case "0.75"
                ФондОТ = valr * 0.75
            Case "1.0"
                ФондОТ = valr
        End Select

        'valr = valr * Val(ComboBox10.Text)
        РДОрубли = Math.Floor(valr)
        РДОкопейки = System.Math.Round(valr - Math.Floor(valr), 2)
        РДОкопейки = Mid(РДОкопейки, InStr(1, РДОкопейки, ",") + 1)
        If Len(РДОкопейки) = 1 Then
            РДОкопейки = РДОкопейки + "0"
        End If
        valr = System.Math.Round(valr, 2)
        Me.TextBox48.Text = Str(valr)
        Dim ПрРуб As String = Пропись(РДОрубли)
        Dim ПрКоп As String = Пропись(РДОкопейки)
        If valr <> Fix(valr) Then
            Me.TextBox47.Text = ПрРуб & "бел.руб, " & Strings.LCase(ПрКоп) & "копеек"
        Else
            Me.TextBox47.Text = ПрРуб & "бел.руб, 00 копеек"
        End If

        If TextBox33.Text <> "" And TextBox44.Text = "" Or TextBox33.Text <> "" And TextBox44.Text = "00" Then
            rub = Пропись(TextBox33.Text)
            TextBox43.Text = rub & "бел.руб. 00 копеек"
        ElseIf TextBox33.Text <> "" And TextBox44.Text <> "" Then

            TextBox43.Text = Пропись(TextBox33.Text) & "бел.руб, " & Strings.LCase(Пропись(TextBox44.Text)) & " копеек."

        End If

    End Sub

    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged

        TextBox12.Text = TextBox12.Text.ToUpper()
        TextBox12.Select(TextBox12.Text.Length, 0)
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        TextBox8.Text = TextBox8.Text.ToUpper()
        TextBox8.Select(TextBox8.Text.Length, 0)
        TextBox45.Text = TextBox8.Text

        Label76.Text = TextBox8.Text.Length
        If TextBox8.Text.Length = 14 Then
            Label76.ForeColor = Color.Green
            Label77.ForeColor = Color.Green
            Label77.Text = "OK"
        Else
            Label76.ForeColor = Color.Red
            Label77.ForeColor = Color.Red
            Label77.Text = "NO"
        End If

        'Select Case TextBox8.TextLength
        '    Case 3
        '        TextBox8.Text &= " "
        'End Select
    End Sub


    Private Sub TextBox46_TextChanged(sender As Object, e As EventArgs) Handles TextBox46.TextChanged
        'If ComboBox9.Text <> "" Or CheckBox5.Checked = True Then
        '    ПропОклад()
        'End If
    End Sub

    Private Sub ComboBox15_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox15.SelectedIndexChanged
        If ComboBox15.Text <> "" Then
            Label88.ForeColor = Color.Green
            Label88.Text = "OK"
        Else
            Label88.ForeColor = Color.Red
            Label88.Text = "NO"

        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs)
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox12.Text = ""
        TextBox45.Text = ""
        If CheckBox1.Checked = True Then 'иностранец
            TextBox7.MaxLength = 25
            TextBox12.MaxLength = 10
            TextBox8.MaxLength = 25
        ElseIf CheckBox1.Checked = False Then
            TextBox7.MaxLength = 7
            TextBox12.MaxLength = 2
            TextBox8.MaxLength = 14
            TextBox7.Text = ""
            TextBox12.Text = ""
            TextBox8.Text = ""
            TextBox45.Text = ""
        End If
    End Sub





    Private Sub TextBox45_TextChanged(sender As Object, e As EventArgs) Handles TextBox45.TextChanged
        TextBox45.Text = TextBox45.Text.ToUpper()
        TextBox45.Select(TextBox45.Text.Length, 0)
    End Sub
    Private Sub Clr(ByVal ct As Control) 'функция чистит все текстбоксы
        For Each c As Control In ct.Controls
            If TypeOf c Is TextBox Then
                c.Text = ""
            Else
                Clr(c)
            End If
        Next
    End Sub
    Private Sub ClAll()

        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox20.Text = ""
        TextBox21.Text = ""
        TextBox40.Text = ""
        MaskedTextBox10.Text = ""
        TextBox37.Text = ""
        TextBox24.Text = ""
        TextBox23.Text = ""
        TextBox19.Text = ""
        TextBox25.Text = ""
        TextBox27.Text = ""
        TextBox30.Text = ""
        TextBox32.Text = ""
        TextBox36.Text = ""
        ComboBox14.Text = ""
        ComboBox14.Text = "Нет"
        TextBox29.Text = ""
        TextBox26.Text = ""
        TextBox28.Text = ""
        TextBox31.Text = ""
        TextBox35.Text = ""
        MaskedTextBox1.Text = ""
        'MaskedTextBox1.Text = Format(Now, "dd.MM.yyyy")
        MaskedTextBox2.Text = ""
        'Dim dft As Date = CDate(MaskedTextBox1.Text)
        'MaskedTextBox2.Text = dft.AddYears(10)
        TextBox12.Text = ""
        TextBox7.Text = ""
        TextBox9.Text = ""
        TextBox8.Text = ""
        TextBox45.Text = ""
        TextBox44.Text = ""
        Label98.Text = ""
        CheckBox28.Checked = False
        If ПровИндивидКонтр(ComboBox1.Text) = True Then
            MaskedTextBox9.Text = ""
            TextBox51.Text = ""
        End If

        'лист2
        ComboBox7.Text = ""
        ComboBox8.Text = ""
        ComboBox9.Text = ""
        ComboBox10.Text = ""
        ComboBox18.Text = ""
        ComboBox12.Text = ""
        ComboBox15.Text = ""
        ComboBox16.Text = ""
        If CheckBox5.Checked = False Then
            ComboBox19.Text = ""
        End If
        'CheckBox5.Checked = False
        CheckBox7.Checked = False
        ComboBox11.Text = ""
        TextBox33.Text = ""
        TextBox43.Text = ""
        TextBox46.Text = ""
        TextBox47.Text = ""
        TextBox48.Text = ""
        TextBox38.Text = ""
        TextBox41.Text = ""
        TextBox49.Text = ""
        TextBox50.Text = ""
        CheckBox2.Checked = False
        CheckBox4.Checked = False
        TextBox40.Text = String.Empty
        TextBox56.Text = String.Empty
        Label88.Text = "NO"
        Label88.ForeColor = Color.Red
        Label89.Text = "NO"
        Label89.ForeColor = Color.Red
        Label90.Text = "NO"
        Label90.ForeColor = Color.Red
        Label85.Text = "NO"
        Label85.ForeColor = Color.Red
        MaskedTextBox3.Text = Now.ToShortDateString
        MaskedTextBox4.Text = Now.ToShortDateString
        MaskedTextBox5.Text = Now.ToShortDateString
        Label48.Text = ""
    End Sub
    Private Sub ДанПрошлГод()
        Dim Прошл, Сегод As String
        Прошл = Now.Year - 1
        Сегод = Now.Year
        Dim Files(), Files4(), Files3(), Files2() As String

        Try

            Files3 = (IO.Directory.GetFiles(OnePath & ComboBox1.Text & "\Контракт\" & Сегод, "*.doc", IO.SearchOption.TopDirectoryOnly))
            Files2 = (IO.Directory.GetFiles(OnePath & ComboBox1.Text & "\Приказ\" & Сегод, "*.doc", IO.SearchOption.TopDirectoryOnly))
        Catch ex As Exception

        End Try


        Try
            Files = (IO.Directory.GetFiles(OnePath & ComboBox1.Text & "\Контракт\" & Прошл, "*.doc", IO.SearchOption.TopDirectoryOnly))
            Files4 = (IO.Directory.GetFiles(OnePath & ComboBox1.Text & "\Приказ\" & Прошл, "*.doc", IO.SearchOption.TopDirectoryOnly))
        Catch ex As Exception

        End Try


        Me.ComboBox2.Items.Clear()
        Me.ComboBox17.Items.Clear()


        Dim gth, gth2, gth3, gth4 As String

        Try
            For n As Integer = 0 To Files2.Length - 1
                gth = ""
                gth = IO.Path.GetFileName(Files2(n))
                Files2(n) = gth
                'TextBox44.Text &= gth + vbCrLf
            Next

            For n As Integer = 0 To Files3.Length - 1
                gth3 = ""
                gth3 = IO.Path.GetFileName(Files3(n))
                Files3(n) = gth3
                'TextBox44.Text &= gth + vbCrLf
            Next

            Array.Sort(Files2)
            Array.Sort(Files3)

            'ComboBox21.Items.AddRange(Files2)
            ComboBox2.Items.AddRange(Files3)


        Catch ex As Exception
            MessageBox.Show("Это будет первый контракт!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
        End Try


        Try
            For n As Integer = 0 To Files.Length - 1
                gth2 = ""
                gth2 = IO.Path.GetFileName(Files(n))
                Files(n) = gth2
                'TextBox44.Text &= gth + vbCrLf
            Next


            For n As Integer = 0 To Files4.Length - 1
                gth4 = ""
                gth4 = IO.Path.GetFileName(Files4(n))
                Files4(n) = gth4
                'TextBox44.Text &= gth + vbCrLf
            Next

            Array.Sort(Files)
            Array.Sort(Files4)
            ComboBox17.Items.AddRange(Files4)
            'ComboBox20.Items.AddRange(Files)
        Catch ex As Exception
            'MessageBox.Show("Нет документов за прошлый год!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
        End Try

        Me.ComboBox9.Items.Clear()

    End Sub

    Private Sub ComboBox11_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox11.SelectedIndexChanged
        If ComboBox11.Text = "" Then
            Label81.ForeColor = Color.Red
            Label81.Text = "NO"
        Else
            Label81.ForeColor = Color.Green
            Label81.Text = "OK"
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged_1(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then 'иностранец
            TextBox7.MaxLength = 25
            TextBox12.MaxLength = 10
            TextBox8.MaxLength = 25
        ElseIf CheckBox1.Checked = False Then
            TextBox7.MaxLength = 7
            TextBox12.MaxLength = 2
            TextBox8.MaxLength = 14
        End If
    End Sub

    Private Sub TextBox21_TextChanged(sender As Object, e As EventArgs) Handles TextBox21.TextChanged
        If TextBox21.Text <> "" Then

            Label84.ForeColor = Color.Green
            Label84.Text = "OK"
        Else

            Label84.ForeColor = Color.Red
            Label84.Text = "NO"
        End If





        TextBox20.Text = TextBox21.Text
    End Sub

    Private Sub TextBox44_TextChanged(sender As Object, e As EventArgs) Handles TextBox44.TextChanged
        'TextBox43.Text = ""

        'If Me.TextBox44.Text <> "" Then
        '    Dim kop As String = Пропись(Me.TextBox44.Text)
        '    kop = Strings.LCase(kop)
        '    TextBox43.Text = rub & "бел.рублей " & kop & "копеек."
        'ElseIf Me.TextBox44.Text = "" And TextBox33.Text <> "" Then
        '    TextBox43.Text = rub & "бел.рублей 00 копеек"
        'End If
        'If Me.TextBox44.Text = "00" Or Me.TextBox44.Text = ToString(0) And TextBox33.Text <> "" Then
        '    TextBox43.Text = rub & "бел.рублей 00 копеек"
        'End If
        'If TextBox33.Text <> "" Then
        '    ПропОклад()
        'End If

    End Sub

    Private Sub TextBox20_TextChanged(sender As Object, e As EventArgs) Handles TextBox20.TextChanged

    End Sub
    Private Function ПроверкаУвольнения(ByVal ИДНомер As String) As String

        Dim ds7 = Selects(StrSql:="SELECT КарточкаСотрудника.ДатаУвольнения
FROM Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр
WHERE Сотрудники.ИДНомер ='" & ИДНомер & "'") 'ищем по ИДПаспорта, уволен ли сотрудник
        Dim s As String
        Try
            s = ds7.Rows(0).Item(0).ToString
        Catch ex As Exception
            Exit Function
        End Try

        Select Case s
            Case <> ""
                Dim sp As String = MsgBox("Сотрудник ранее работал на этом предприятии " & vbCrLf & "но был уволен " & s & " продолжить далее?", vbOKCancel, Рик)
                If sp = 2 Then
                    Return 1
                    Exit Function
                    Return 2
                End If

            Case ""
                Dim sp As String = MsgBox("Сотрудник работает на этом предприятии " & vbCrLf & "и пока числится в штате " & vbCrLf & "Продолжить далее оформление?", vbOKCancel, Рик)
                If sp = 2 Then
                    Return 1
                    Exit Function
                    Return 2
                End If
        End Select
    End Function
    Private Sub Чист()
        Try
            StrSql = ""
            ds.Clear()
        Catch ex As Exception

        End Try

    End Sub
    Private Function МестоРаботы()
        Dim StrSql As String = "SELECT * From ОбъектОбщепита Where АдресОбъекта = '" & combx18 & "' AND НазвОрг= '" & combx1 & "'"
        Dim ds As DataTable = Selects(StrSql)
        Dim Тип, Название As String

        Try
            If ds.Rows(0).Item(3).ToString = "" Then
                Название = ""
            Else
                Название = """" & ds.Rows(0).Item(3).ToString & ""","
            End If
        Catch ex As Exception
            MessageBox.Show("Выберите другой объект общепита!", Рик)
            Return 1
        End Try


        If ds.Rows(0).Item(2).ToString = "" Then
            Тип = ""
        Else
            Тип = Strings.Trim(Strings.LCase(ds.Rows(0).Item(2).ToString))
        End If

        If Название = "" And Тип = "" Then
            МестоРаб = combx18
        ElseIf Название <> "" And Тип = "" Then
            МестоРаб = Название & " " & combx18
        ElseIf Название = "" And Тип <> "" Then
            МестоРаб = Тип & " " & combx18
        Else
            МестоРаб = Тип & " " & Название & " " & combx18
        End If
        Return 0


    End Function
    Public Class NegativeNumberException
        Inherits Exception
        Sub New()
            MyBase.New("В базе дданных нет информаци!")
        End Sub
    End Class

    Private Sub СборДаннОрганиз()

        'сборка данных для доков со стороны руководства, новый вариант


        Dim ds() As DataRow = ВыборкаСтрокиИзТаблицы(combx1, dtClientAll, "НазвОрг")
        'Dim ds1 = dtClientAll.Select("НазвОрг='" & Клиент & "'")

        Dim РуковИП As String

        If ds(0).Item("РукИП") = "True" Then
            РуковИП = "ИП"
        Else
            РуковИП = ""
        End If

        ФормаСобстПолн = ds(0).Item("ФормаСобств").ToString
        ДолжРуков = ds(0).Item("ДолжнРуководителя").ToString
        ФИОРукРодПад = РуковИП & ds(0).Item("ФИОРукРодПадеж").ToString
        ОснованиеДейств = ds(0).Item("ОснованиеДейств").ToString



        'короткое фио 
        Dim nm As String = ds(0).Item("ФИОРуководителя").ToString
        Dim nm0 As Integer = Len(ds(0).Item("ФИОРуководителя").ToString)
        Dim nm1 As String = Strings.Left(nm, InStr(nm, " "))
        Dim nm2 As Integer = Len(nm1)
        Dim nm3 As String = Strings.Right(nm, (nm0 - nm2))
        Dim nm31 As Integer = Len(nm3)
        Dim nm4 As String = Strings.UCase(Strings.Left(Strings.Left(nm3, InStr(nm3, " ")), 1))
        Dim nm41 As Integer = Len(Strings.Left(nm3, InStr(nm3, " ")))
        Dim nm5 As String = Strings.UCase(Strings.Left(Strings.Right(nm3, nm31 - nm41), 1))
        Dim nm6 = Strings.Left(ds(0).Item("ФИОРукДатПадеж").ToString, InStr(ds(0).Item("ФИОРукДатПадеж").ToString, " "))



        ФИОКор = РуковИП & nm1 & " " & nm4 & "." & nm5 & "."
        ФИОКорРукДат = РуковИП & nm6 & " " & nm4 & "." & nm5 & "."
        УНП = ds(0).Item("УНП").ToString
        КонтТелефон = ds(0).Item("КонтТелефон").ToString
        ЮрАдрес = ds(0).Item("ЮрАдрес").ToString
        РасСчет = ds(0).Item("РасчСчетРубли").ToString
        Банк = ds(0).Item("Банк").ToString
        БИК = ds(0).Item("БИКБанка").ToString
        АдресБанка = ds(0).Item("АдресБанка").ToString
        ЭлАдрес = ds(0).Item("ЭлАдрес").ToString


        ''сокращенное название орг
        'Dim ds9 As DataTable = Selects(StrSql:="Select Сокращенное From ФормаСобств Where ПолноеНазвание = '" & ds(0).Item("ФормаСобств").ToString & "'")



        Dim ds9() As DataRow = ВыборкаСтрокиИзТаблицы(ds(0).Item("ФормаСобств").ToString, dtformft, "ПолноеНазвание")


        If ds9.Length = 0 Then Throw New NegativeNumberException()
        If ds9.Length >= 1 Then
            ФормаСобствКор = ds9(0).Item(2).ToString
        End If


        Dim кл2 As String

        If ФормаСобствКор = "ИП" Then
            кл2 = " " & Клиент & " "
        Else
            кл2 = " """ & Клиент & """ "
        End If

        'Dim strsql5 As String = "SELECT * From Клиент Where Клиент.НазвОрг = '" & Клиент & "'"
        'Dim ds3 As DataTable = Selects(strsql5)


        СборноеРеквПолн = ФормаСобствКор & кл2 & vbCrLf & ds(0).Item(4).ToString & " IBAN " _
        & ds(0).Item(14).ToString & " в " & ds(0).Item(13).ToString & vbCrLf & " БИК " _
        & ds(0).Item(11).ToString & vbCrLf & " УНП " & ds(0).Item(2).ToString


    End Sub
    Private Sub ИзменФамил(ByVal id As Integer)

        Dim ds = dtSotrudnikiAll.Select("КодСотрудники=" & id & "")
        'Dim strsql As String = "SELECT * FROM Сотрудники WHERE КодСотрудники=" & id & ""
        'Dim ds As DataTable = Selects(strsql)
        Dim list As New Dictionary(Of String, Object)
        list.Add("@КодСотрудники", id)


        Updates(stroka:="UPDATE Сотрудники SET ФамилияСтар='" & ds(0).Item(2).ToString & "',ИмяСтар='" & ds(0).Item(3).ToString & "',
ОтчествоСтар='" & ds(0).Item(4).ToString & "',ФИОСборноеСтар='" & Trim(ds(0).Item(2).ToString & " " & ds(0).Item(3).ToString & " " & ds(0).Item(4).ToString) & "',
ФамилияРодПадСтар='" & ds(0).Item(6).ToString & "',ИмяРодПадСтар='" & ds(0).Item(7).ToString & "',
ОтчествоРодПадСтар='" & ds(0).Item(8).ToString & "', ДатаИзменения='" & Now.ToShortDateString & "', ФИОРодПодСтар='" & Trim(ds(0).Item(6).ToString & " " & ds(0).Item(7).ToString & " " & ds(0).Item(8).ToString) & "'
WHERE КодСотрудники=@КодСотрудники", list, "Сотрудники")

    End Sub
    Private Async Sub ОбновлСотрудника()




        Await Task.Delay(0)
        Dim IDСотрудника As Integer
        Try
            IDСотрудника = CType(Label96.Text, Integer)
        Catch ex As Exception
            MessageBox.Show("Сотрудника нет в базе!", Рик)
            Exit Sub
        End Try

        Dim list As New Dictionary(Of String, Object)
        list.Add("@КодСотрудники", IDСотрудника)
        list.Add("@IDСотр", IDСотрудника)


        If CheckBox28.Checked = True Then
            ИзменФамил(IDСотрудника)
        End If
        Dim inostan As String
        If CheckBox1.Checked = True Then
            inostan = "True"
        Else
            inostan = "False"
        End If

        'Обновляем таблицу сотрудники данными и обновляем саму таблицу.
        Updates(stroka:="UPDATE Сотрудники SET Сотрудники.Фамилия='" & Trim(TextBox1.Text) & "', Сотрудники.Имя='" & Trim(TextBox2.Text) & "', Сотрудники.Отчество='" & Trim(TextBox3.Text) & "', 
Сотрудники.ФамилияРодПад='" & Trim(TextBox6.Text) & "', Сотрудники.ИмяРодПад='" & Trim(TextBox5.Text) & "', Сотрудники.ОтчествоРодПад='" & Trim(TextBox4.Text) & "', 
Сотрудники.ПаспортСерия='" & TextBox12.Text & "', Сотрудники.ПаспортНомер='" & TextBox7.Text & "', Сотрудники.ПаспортКогдаВыдан='" & MaskedTextBox1.Text & "',
Сотрудники.ДоКакогоДейств='" & MaskedTextBox2.Text & "', Сотрудники.ПаспортКемВыдан='" & TextBox9.Text & "', Сотрудники.ИДНомер='" & TextBox8.Text & "',
Сотрудники.Регистрация='" & TextBox21.Text & "', Сотрудники.МестоПрожив='" & TextBox20.Text & "', Сотрудники.КонтТелГор='" & TextBox37.Text & "',
Сотрудники.КонтТелефон='" & MaskedTextBox10.Text & "', Сотрудники.СтраховойПолис='" & TextBox45.Text & "', Сотрудники.ФамилияДляЗаявления='" & Trim(TextBox34.Text) & "',
Сотрудники.ИмяДляЗаявления='" & Trim(TextBox11.Text) & "', Сотрудники.ОтчествоДляЗаявления='" & Trim(TextBox10.Text) & "', Сотрудники.Пол='" & cmb28 & "', Сотрудники.ДатаРожд='" & MaskedTextBox9.Text & "',
Сотрудники.Гражданин='" & TextBox51.Text & "', Сотрудники.Иностранец='" & inostan & "',
ФИОСборное='" & Trim(TextBox1.Text) & " " & Trim(TextBox2.Text) & " " & Trim(TextBox3.Text) & "', ФИОРодПод='" & Trim(TextBox6.Text) & " " & Trim(TextBox5.Text) & " " & Trim(TextBox4.Text) & "'
        WHERE Сотрудники.КодСотрудники=@КодСотрудники", list, "Сотрудники")



        Dim ds = Selects(StrSql:="SELECT СоставСемьи.ФИО FROM СоставСемьи WHERE IDСотр=@IDСотр", list)


        Select Case errds
            Case 1
                Updates(stroka:="INSERT INTO СоставСемьи(IDСотр, КолДетей, ФИО, МестоРаботы, Телефон, ДетиПол1, ФИО1, ДатаРождения1, ДетиПол2, ФИО2, ДатаРождения2, ДетиПол3, ФИО3, ДатаРождения3, ДетиПол4, ФИО4, ДатаРождения4, ДетиПол5, ФИО5, ДатаРождения5)
VALUES(" & IDСотрудника & ",'" & combx14 & "','" & TextBox24.Text & "','" & TextBox23.Text & "',
'" & TextBox19.Text & "','" & combx3 & "', '" & TextBox25.Text & "','" & TextBox29.Text & "','" & combx4 & "', '" & TextBox27.Text & "',
'" & TextBox26.Text & "','" & combx5 & "', '" & TextBox30.Text & "','" & TextBox28.Text & "','" & combx6 & "', '" & TextBox32.Text & "',
'" & TextBox31.Text & "','" & combx13 & "',' " & TextBox36.Text & "', '" & TextBox35.Text & "')", list, "СоставСемьи")

            Case 0

                Updates(stroka:="UPDATE СоставСемьи SET СоставСемьи.КолДетей='" & combx14 & "', СоставСемьи.ФИО='" & TextBox24.Text & "', СоставСемьи.МестоРаботы='" & TextBox23.Text & "',
СоставСемьи.Телефон='" & TextBox19.Text & "', СоставСемьи.ДетиПол1='" & combx3 & "', СоставСемьи.ФИО1='" & TextBox25.Text & "', СоставСемьи.ДатаРождения1='" & TextBox29.Text & "',
СоставСемьи.ДетиПол2='" & combx4 & "', СоставСемьи.ФИО2='" & TextBox27.Text & "', СоставСемьи.ДатаРождения2='" & TextBox26.Text & "', СоставСемьи.ДетиПол3='" & combx5 & "',
СоставСемьи.ФИО3='" & TextBox30.Text & "', СоставСемьи.ДатаРождения3='" & TextBox28.Text & "', СоставСемьи.ДетиПол4='" & combx6 & "', СоставСемьи.ФИО4='" & TextBox32.Text & "',
СоставСемьи.ДатаРождения4='" & TextBox31.Text & "', СоставСемьи.ДетиПол5='" & combx13 & "', СоставСемьи.ФИО5='" & TextBox36.Text & "', СоставСемьи.ДатаРождения5='" & TextBox35.Text & "'
WHERE СоставСемьи.IDСотр =@IDСотр", list, "СоставСемьи")

        End Select



        Dim посм As String
        If CheckBox2.Checked = True Then
            посм = "по совместительству"
        Else
            посм = ""
        End If


        ДатаУведомл(combx11, MaskedTextBox4.Text)

        Dim adf As String

        If CheckBox4.Checked = True Then
            adf = "Да"
        Else
            adf = ""
        End If


        Dim ds2 = Selects(StrSql:="SELECT ДатаПриема,ПродлКонтрС FROM КарточкаСотрудника WHERE IDСотр=" & IDСотрудника & "")


        Select Case errds
            Case 1

                Updates(stroka:="INSERT INTO КарточкаСотрудника(IDСотр,ДатаПриема,СрокКонтракта,ТипРаботы,Ставка,ВремяНачРаботы,ПродолРабДня,АдресОбъектаОбщепита,ПоСовмест,
ДатаЗарплаты,ДатаАванса,ДатаУведомлПродКонтр,СуммирУчет,Примечание) VALUES(" & IDСотрудника & ",'" & MaskedTextBox4.Text & "','" & combx11 & "','" & combx15 & "',
'" & combx10 & "','" & combx12 & "', '" & combx16 & "','" & combx18 & "','" & посм & "','" & TextBox40.Text & "',
'" & TextBox56.Text & "','" & ДатаУведомл(combx11, MaskedTextBox4.Text) & "', '" & adf & "', '" & Примечани & "')", list, "КарточкаСотрудника")
            Case 0


                If ds2.Rows(0).Item(1).ToString <> "" Then
                    If MessageBox.Show("С данным сотрудником продлен контракт" & vbCrLf & "Если вы не меняли!" & vbCrLf & "1)Дату(приказа,контракта)" & vbCrLf & "2)Период контракта!" & vbCrLf & "Нажмите 'Да'" & vbCrLf & "Если были изменения нажмите 'Нет'", Рик, MessageBoxButtons.YesNo) = DialogResult.No Then
                        If MessageBox.Show("Будет внесены следующие изменения!" & vbCrLf & "1)Заменены старые даты приема, контракта, приказа" & vbCrLf & "2)Изменена дата уведомления о продлении контракта" & vbCrLf & "3)Удалены все даты продлений контракта", Рик, MessageBoxButtons.YesNo) = DialogResult.Yes Then
                            Dim bn As String = ""
                            Dim bi As Integer = Nothing
                            'чистим данные из карточки сотрудника
                            Updates(stroka:="UPDATE КарточкаСотрудника SET
КарточкаСотрудника.НомерУведомлПродКонтр='" & bn & "', КарточкаСотрудника.СрокПродлКонтракта='" & bn & "',
КарточкаСотрудника.ПродлКонтрС = Null, КарточкаСотрудника.ПродлКонтрПо = Null, КарточкаСотрудника.ПриказПродлКонтр='" & bn & "'
WHERE КарточкаСотрудника.IDСотр=@IDСотр", list, "КарточкаСотрудника")


                            'чистим данные из таблицы продление контракта
                            Updates(stroka:="UPDATE ПродлКонтракта SET
ПродлКонтракта.ДатаПриема='" & bn & "', ПродлКонтракта.ДатаОкончания='" & bn & "', ПродлКонтракта.СрокКонтракта='" & bn & "', 
ПродлКонтракта.НомерУвед='" & bn & "', ПродлКонтракта.ПервоеПродлениеС='" & bn & "', ПродлКонтракта.ПервоеПродлениеПо='" & bn & "',
ПродлКонтракта.ПервоеПродлениеСрок='" & bn & "', ПродлКонтракта.НомерУвед1='" & bn & "', ПродлКонтракта.ВтороеПродлениеС='" & bn & "',
ПродлКонтракта.ВтороеПродлениеПо='" & bn & "', ПродлКонтракта.ВтороеПродлениеСрок='" & bn & "', ПродлКонтракта.НомерУвед2='" & bn & "',
ПродлКонтракта.ТретьеПродлениеС='" & bn & "', ПродлКонтракта.ТретьеПродлениеПо='" & bn & "', ПродлКонтракта.ТретьеПродлениеСрок='" & bn & "',
ПродлКонтракта.НомерУвед3='" & bn & "', ПродлКонтракта.ЧетвертоеПродлениеС='" & bn & "', ПродлКонтракта.ЧетвертоеПродлениеПо='" & bn & "',
ПродлКонтракта.ЧетвертоеПродлениеСрок='" & bn & "', ПродлКонтракта.НомерУвед4='" & bn & "'
WHERE ПродлКонтракта.IDСотр=@IDСотр", list, "ПродлКонтракта")



                            'вносим данные в таблицу карточка сотрудника
                            Updates(stroka:="UPDATE КарточкаСотрудника SET КарточкаСотрудника.ДатаПриема='" & MaskedTextBox4.Text & "',
КарточкаСотрудника.СрокКонтракта='" & combx11 & "', КарточкаСотрудника.ТипРаботы='" & combx15 & "',
КарточкаСотрудника.Ставка='" & combx10 & "', КарточкаСотрудника.ВремяНачРаботы='" & combx12 & "',
КарточкаСотрудника.ПродолРабДня='" & combx16 & "',КарточкаСотрудника.АдресОбъектаОбщепита='" & combx18 & "',
КарточкаСотрудника.ПоСовмест='" & посм & "',КарточкаСотрудника.ДатаЗарплаты='" & TextBox40.Text & "',
КарточкаСотрудника.ДатаАванса='" & TextBox56.Text & "', 
КарточкаСотрудника.ДатаУведомлПродКонтр='" & ДатаУведомл(combx11, MaskedTextBox4.Text) & "', КарточкаСотрудника.СуммирУчет= '" & adf & "',
КарточкаСотрудника.Примечание= '" & Примечани & "'
WHERE КарточкаСотрудника.IDСотр=@IDСотр", list, "КарточкаСотрудника")



                            TextBox38.Text = Replace(TextBox38.Text, "\", ".")
                            TextBox38.Text = Replace(TextBox38.Text, "/", ".")


                            'вносим данные в таблицу продление контракта
                            Updates(stroka:="UPDATE ПродлКонтракта SET ДатаПриема='" & MaskedTextBox4.Text & "', ДатаОкончания='" & MaskedTextBox5.Text & "',
СрокКонтракта='" & combx11 & "', НомерУвед='" & TextBox38.Text & "'
WHERE ПродлКонтракта.IDСотр=@IDСотр", list, "ПродлКонтракта")


                        End If
                        '                    Else
                        '                        Чист()
                        '                        StrSql = "UPDATE КарточкаСотрудника SET 
                        'КарточкаСотрудника.СрокКонтракта='" & combx11 & "', КарточкаСотрудника.ТипРаботы='" & combx15 & "',
                        'КарточкаСотрудника.Ставка='" & combx10 & "', КарточкаСотрудника.ВремяНачРаботы='" & combx12 & "',
                        'КарточкаСотрудника.ПродолРабДня='" & combx16 & "',КарточкаСотрудника.АдресОбъектаОбщепита='" & combx18 & "',
                        'КарточкаСотрудника.ПоСовмест='" & посм & "',КарточкаСотрудника.ДатаЗарплаты='" & TextBox40.Text & "',
                        'КарточкаСотрудника.ДатаАванса='" & TextBox56.Text & "',  КарточкаСотрудника.СуммирУчет= '" & adf & "',
                        'КарточкаСотрудника.Примечание= '" & Примечани & "'
                        'WHERE КарточкаСотрудника.IDСотр= " & IDСотрудника & ""
                        '                        Updates(StrSql)
                    End If
                    Updates(stroka:="UPDATE КарточкаСотрудника SET 
КарточкаСотрудника.СрокКонтракта='" & combx11 & "', КарточкаСотрудника.ТипРаботы='" & combx15 & "',
КарточкаСотрудника.Ставка='" & combx10 & "', КарточкаСотрудника.ВремяНачРаботы='" & combx12 & "',
КарточкаСотрудника.ПродолРабДня='" & combx16 & "',КарточкаСотрудника.АдресОбъектаОбщепита='" & combx18 & "',
КарточкаСотрудника.ПоСовмест='" & посм & "',КарточкаСотрудника.ДатаЗарплаты='" & TextBox40.Text & "',
КарточкаСотрудника.ДатаАванса='" & TextBox56.Text & "', КарточкаСотрудника.СуммирУчет= '" & adf & "',
КарточкаСотрудника.Примечание= '" & Примечани & "'
WHERE КарточкаСотрудника.IDСотр=@IDСотр", list, "КарточкаСотрудника")

                Else
                    Updates(stroka:="UPDATE КарточкаСотрудника SET КарточкаСотрудника.ДатаПриема='" & MaskedTextBox4.Text & "',
КарточкаСотрудника.СрокКонтракта='" & combx11 & "', КарточкаСотрудника.ТипРаботы='" & combx15 & "',
КарточкаСотрудника.Ставка='" & combx10 & "', КарточкаСотрудника.ВремяНачРаботы='" & combx12 & "',
КарточкаСотрудника.ПродолРабДня='" & combx16 & "',КарточкаСотрудника.АдресОбъектаОбщепита='" & combx18 & "',
КарточкаСотрудника.ПоСовмест='" & посм & "',КарточкаСотрудника.ДатаЗарплаты='" & TextBox40.Text & "',
КарточкаСотрудника.ДатаАванса='" & TextBox56.Text & "', 
КарточкаСотрудника.ДатаУведомлПродКонтр='" & ДатаУведомл(combx11, MaskedTextBox4.Text) & "', КарточкаСотрудника.СуммирУчет= '" & adf & "',
КарточкаСотрудника.Примечание= '" & Примечани & "'
WHERE КарточкаСотрудника.IDСотр=@IDСотр", list, "КарточкаСотрудника")

                    'вносим данные в таблицу продление контракта
                    Updates(stroka:="UPDATE ПродлКонтракта SET ДатаПриема='" & MaskedTextBox4.Text & "', ДатаОкончания='" & MaskedTextBox5.Text & "',
СрокКонтракта='" & combx11 & "', НомерУвед='" & TextBox38.Text & "'
WHERE ПродлКонтракта.IDСотр=@IDСотр", list, "ПродлКонтракта")

                End If



        End Select


        Dim рдо, рдо2 As Double
        Dim Сум, РДОс As String

        рдо = Replace(TextBox48.Text, ".", ",")
        рдо2 = Replace(combx10, ".", ",")
        рдо = Math.Round((рдо * рдо2), 2)
        РДОс = CType(рдо, String)

        If TextBox44.Text = "00" Then
            Сум = TextBox33.Text
        Else
            Сум = TextBox33.Text & "," & TextBox44.Text
        End If

        If combx7 = "нет" Then combx7 = ""

        StrSql = ""
        If CheckBox26.Checked = True Then

            Чист()

            StrSql = "SELECT Штатное.ИДСотр FROM Штатное WHERE Штатное.ИДСотр = " & IDСотрудника & ""
            ds = Selects(StrSql)
            If errds = 1 Then
                errds = 0
                Dim fg As Double
                If txtbxD46 = "" Then
                    fg = Nothing
                ElseIf txtbxD46.Length > 2 Then
                    fg = CType(txtbxD46, Double)
                Else
                    fg = CType(txtbxD46, Integer)
                End If

                Dim ПовоклРуб As Double = Math.Round(((CDbl(Сум) * fg) / 100), 2)
                Dim ЧТС As Double = Math.Round(CDbl(Replace(TextBox48.Text, ".", ",")) / 168, 2)

                Updates(stroka:="INSERT INTO Штатное(ИДСотр,Отдел,Должность,Разряд,ТарифнаяСтавка,ПовышОклПроц,
РасчДолжностнОклад,ФонОплатыТруда,ПовышОклРуб,ЧасоваяТарифСтавка)
VALUES(" & IDСотрудника & ",'" & combx8 & "', '" & combx9 & "', '" & combx7 & "', " & Replace(CDbl(Сум), ",", ".") & ",
" & Replace(fg, ",", ".") & ", " & Replace(CDbl(Replace(TextBox48.Text, ".", ",")), ",", ".") & ", " & Replace(CDbl(РДОс), ",", ".") & ",
" & Replace(CDbl(ПовоклРуб), ",", ".") & "," & Replace(CDbl(ЧТС), ",", ".") & ")", list, "Штатное")
            Else

                Чист()
                Dim Dfg As Double
                If txtbxD46 = "" Then
                    Dfg = Nothing
                ElseIf txtbxD46.Length > 2 Then
                    Dfg = CType(Replace(txtbxD46, ".", ","), Double)
                Else
                    Dfg = CType(txtbxD46, Integer)
                End If



                '                StrSql = "UPDATE Штатное SET Штатное.Отдел='" & combx8 & "', Штатное.Должность='" & combx9 & "',
                'Штатное.Разряд='" & combx7 & "', Штатное.ТарифнаяСтавка='" & CDbl(Сум) & "', Штатное.ПовышОклПроц='" & Dfg & "',
                'Штатное.РасчДолжностнОклад='" & CDbl(Replace(TextBox48.Text, ".", ",")) & "', Штатное.ФонОплатыТруда='" & CDbl(РДОс) & "'
                'WHERE Штатное.ИДСотр = " & IDСотрудника & ""

                Dim rsch As Double = CDbl(Replace(TextBox48.Text, ".", ","))

                Dim ПовоклРуб As Double = Math.Round(((CDbl(Сум) * Dfg) / 100), 2)
                Dim ЧТС As Double = Math.Round(rsch / 168, 2)


                Updates(stroka:="UPDATE Штатное SET Штатное.Отдел='" & combx8 & "', Штатное.Должность='" & combx9 & "',
Штатное.Разряд='" & combx7 & "', Штатное.ТарифнаяСтавка=" & Replace(CDbl(Сум), ",", ".") & ", Штатное.ПовышОклПроц=" & Replace(Dfg, ",", ".") & ",
Штатное.РасчДолжностнОклад=" & Replace(rsch, ",", ".") & ", Штатное.ФонОплатыТруда=" & Replace(CDbl(РДОс), ",", ".") & ",
ПовышОклРуб= " & Replace(ПовоклРуб, ",", ".") & ", ЧасоваяТарифСтавка= " & Replace(ЧТС, ",", ".") & "
WHERE Штатное.ИДСотр =@IDСотр", list, "Штатное")

            End If
        Else
            Dim k = РазрИзменКонтр
            If (Not k = arrtcom("ComboBox10")) And CheckBox5.Checked = True Then
                Dim list1 As New Dictionary(Of String, Object)
                list1.Add("@IDСотр", IDСотрудника)

                Dim f = (From x In dtShtatnoeAll Where x.Item("ИДСотр") = IDСотрудника Select x.Item("РасчДолжностнОклад")).FirstOrDefault
                Dim f1 As Double
                Try
                    f1 = CType(f, Double)
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
                If IsNumeric(f1) Then
                    f1 = Math.Round(f1 * CDbl(Replace(arrtcom("ComboBox10"), ".", ",")), 2)
                    Dim f2 As Double = Math.Round(f1 / 168, 2)
                    list1.Add("@ФонОплатыТруда", f1)
                    list1.Add("@ЧасоваяТарифСтавка", f2)


                    If f1 > 0 Then
                        Updates(stroka:="UPDATE Штатное SET ФонОплатыТруда=@ФонОплатыТруда, ЧасоваяТарифСтавка=@ЧасоваяТарифСтавка
WHERE ИДСотр=@IDСотр", list1, "Штатное")

                    End If
                End If



            End If

        End If


        Чист()
        StrSql = "SELECT Контракт FROM ДогСотрудн WHERE IDСотр=" & IDСотрудника & ""
        ds = Selects(StrSql)
        Select Case errds
            Case 1
                Updates(stroka:="INSERT INTO ДогСотрудн(IDСотр,Контракт,ДатаКонтракта,СрокОкончКонтр,Приказ,Датаприказа) VALUES(" & IDСотрудника & ",
'" & TextBox38.Text & "','" & MaskedTextBox3.Text & "','" & MaskedTextBox5.Text & "','" & TextBox41.Text & " - " & TextBox58.Text & TextBox57.Text & "',
'" & MaskedTextBox3.Text & "')", list, "ДогСотрудн")

            Case 0

                Updates(stroka:="UPDATE ДогСотрудн Set Контракт='" & TextBox38.Text & "', ДатаКонтракта='" & MaskedTextBox3.Text & "',
        СрокОкончКонтр ='" & MaskedTextBox5.Text & "', Приказ='" & TextBox41.Text & " - " & TextBox58.Text & TextBox57.Text & "', Датаприказа='" & MaskedTextBox3.Text & "'
        WHERE IDСотр =@IDСотр", list, "ДогСотрудн")

        End Select

        'RunMoving2()
        'Обнов1()


        Статистика1(combx19, "Изменение данных сотрудника", combx1)

    End Sub
    Private Sub Обнов1()
        dtDogovorPadriada()
        dtProdlenieKontrakta()
        dtSotrudniki() '
        dtKartochkaSotrudnika()
        dtShtatnoe()
        dtDogovorSotrudnik()
    End Sub
    Private Function Налоги(ByVal d As String, ByVal f As String) As List(Of String)

        Dim Копейки2 As String
        Dim Копейки As Integer = CType(f, Integer)
        If Копейки < 10 Then
            Копейки2 = 0 & f
        Else
            Копейки2 = f
        End If
        Dim ПДЦелаяЧасть, ПДДробнаяЧасть, БГСЦелаяЧасть, БГСДробнаяЧасть As String
        Dim СтоимЧаса As Double = CType((d & "," & Копейки2), Double) '.Replace(".", ",")
        Dim ПодохНалог As Double = Math.Round((СтоимЧаса * 13 / 100), 2)
        Dim БГС As Double = Math.Round((СтоимЧаса * 1 / 100), 2)

        If ПодохНалог < 1 Then
            ПДЦелаяЧасть = "0"
            Dim x As String = CType(ПодохНалог, String)
            Select Case x.Length
                Case 3
                    ПДДробнаяЧасть = Strings.Right(CType(ПодохНалог, String), 1)
                    ПДДробнаяЧасть = ПДДробнаяЧасть & "0"
                Case 4
                    ПДДробнаяЧасть = Strings.Right(CType(ПодохНалог, String), 2)
            End Select

            'Dim fg As Integer = CType(ПДДробнаяЧасть, Integer)
            'ПДДробнаяЧасть = CType(fg, String)
        Else
            Dim vz As String = CType(ПодохНалог, String)
            Dim vx As Integer = Strings.Len(vz) - 3
            ПДЦелаяЧасть = Strings.Left(CType(ПодохНалог, String), vx)
            ПДДробнаяЧасть = Strings.Right(CType(ПодохНалог, String), 2)
            Dim fg As Integer = CType(ПДДробнаяЧасть, Integer)
            ПДДробнаяЧасть = CType(fg, String)
        End If

        If БГС < 1 Then

            БГСЦелаяЧасть = "0"
            Dim y As String = CType(БГС, String)
            Select Case y.Length
                Case 3
                    БГСДробнаяЧасть = Strings.Right(CType(БГС, String), 1)
                    БГСДробнаяЧасть = БГСДробнаяЧасть & "0"
                Case 4
                    БГСДробнаяЧасть = Strings.Right(CType(БГС, String), 2)

            End Select

        Else

            Dim vz As String = CType(БГС, String)
            Dim flei As Integer = vz.Length

            If flei = 1 Then
                БГСЦелаяЧасть = vz
                БГСДробнаяЧасть = "0"
            ElseIf vz.Contains(",") Or vz.Contains(".") Then

                Replace(vz, ".", ",")
                Dim целое As Integer = InStr(vz, ",") - 1
                БГСЦелаяЧасть = Strings.Left(vz, целое)
                Dim дробь As Integer = InStrRev(vz, ",") - 1

                If дробь = 1 Then
                    БГСДробнаяЧасть = Strings.Right(vz, целое) & "0"
                Else
                    БГСДробнаяЧасть = Strings.Right(vz, целое)
                End If
            End If

            '    Dim vx As Integer = Strings.Len(vz) - 3
            'БГСЦелаяЧасть = Strings.Left(CType(БГС, String), vx)
            'БГСДробнаяЧасть = Strings.Right(CType(БГС, String), 2)
            'Dim fg As Integer = CType(БГСДробнаяЧасть, Integer)
            'БГСДробнаяЧасть = CType(fg, String)
        End If


        Dim strValues As String() = New String() {ОбязН, ПДЦелаяЧасть, ПДДробнаяЧасть, БГСЦелаяЧасть, БГСДробнаяЧасть} 'из массива в лист оф очень класная штука
        Dim Лист As List(Of String) = strValues.ToList()
        Return Лист


    End Function

    Private Sub ДанныеКлиентаДогПодряда()

        'Данные по клиенту

        Dim StrSql As String = "SELECT Клиент.ФормаСобств, Клиент.ДолжнРуководителя,
Клиент.ФИОРукРодПадеж, Клиент.ОснованиеДейств, Клиент.УНП, Клиент.ЮрАдрес, Клиент.РасчСчетРубли,
Клиент.АдресБанка, Клиент.БИКБанка, Клиент.ФИОРуководителя, Клиент.РукИП
FROM Клиент
WHERE Клиент.НазвОрг='" & arrtcom("ComboBox1") & "'"
        Dim ds As DataTable = Selects(StrSql)

        ReDim массивДогПодр(ds.Columns.Count - 1)

        For i As Integer = 0 To ds.Columns.Count - 1
            массивДогПодр(i) = ds.Rows(0).Item(i)
        Next

        ФИОКор = ""
        Dim РуковИП As String

        If ds.Rows(0).Item(10) = True Then
            ФИОКор = ФИОКорРук(ds.Rows(0).Item(9).ToString, True)
            РуковИП = "ИП "
        Else
            РуковИП = ""
            ФИОКор = ФИОКорРук(ds.Rows(0).Item(9).ToString, False)
        End If

        ФормаСобстПолн = ""
        ФормаСобстПолн = ds.Rows(0).Item(0)

        ФормаСобствКор = ""
        ФормаСобствКор = ФормСобствКор(ds.Rows(0).Item(0).ToString)

        ФИОРукРодПад = ""
        ФИОРукРодПад = РуковИП & ds.Rows(0).Item(2).ToString

        ДолжРуковВинПад = ДобОконч(ds.Rows(0).Item(1).ToString)


        'данные по объекту общепита
        Dim StrSql4 As String = "SELECT ОбъектОбщепита.ТипОбъекта, ОбъектОбщепита.НазОбъекта, ОбъектОбщепита.АдресОбъекта
FROM ОбъектОбщепита
WHERE ОбъектОбщепита.АдресОбъекта='" & arrtcom("ComboBox25") & "'"
        Dim ds4 As DataTable = Selects(StrSql4)
        МестоРаб = ""
        Try
            If ds4.Rows(0).Item(0).ToString = "" And ds4.Rows(0).Item(1).ToString = "" Then
                МестоРаб = ds4.Rows(0).Item(2).ToString
            Else
                МестоРаб = Strings.LCase(ds4.Rows(0).Item(0).ToString) & " «" & ds4.Rows(0).Item(1).ToString & "» " & ds4.Rows(0).Item(2).ToString
            End If
        Catch ex As Exception
            MessageBox.Show("Объект общепита у данного сотрудника не найден!", Рик)
        End Try





    End Sub
    Private Function ДогПодрядаПроверка() As Integer
        If CheckBox5.Checked = False Then
            If MessageBox.Show("Сформировать договор подряда?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.Cancel Then
                Return 1
            End If
        End If

        If but2cl = 1 Then
            MessageBox.Show("Проверьте номер договора подряда!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Hand)
            Return 1
        End If

        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Then
            MessageBox.Show("Заполните ФИО сотрудника!", Рик, MessageBoxButtons.OK)
            Return 1
        End If
        If TextBox20.Text = "" Or TextBox21.Text = "" Then
            MessageBox.Show("Заполните адрес сотрудника!", Рик, MessageBoxButtons.OK)
            Return 1
        End If

        If Примечани = "" Then
            If MessageBox.Show("Вы НЕ заполнили примечание!" & vbCrLf & "Выберите OK - если хотите продолжить, или ОТМЕНА - если хотите изменить", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.Cancel Then
                Return 1
            End If

        End If

        If TextBox7.Text = "" Or TextBox8.Text = "" Or TextBox9.Text = "" Or MaskedTextBox1.Text = "" Or MaskedTextBox2.Text = "" Or TextBox12.Text = "" Or TextBox45.Text = "" And CheckBox1.Checked = False Then
            MessageBox.Show("Заполните паспортные данные сотрудника!", Рик, MessageBoxButtons.OK)
            Return 1
        End If
        If TextBox55.Text = "" Or MaskedTextBox6.MaskCompleted = False Or MaskedTextBox7.MaskCompleted = False Or MaskedTextBox8.MaskCompleted = False Or ComboBox22.Text = "" Then
            MessageBox.Show("Заполните все поля для Договора подряда!", Рик, MessageBoxButtons.OK)
            Return 1
        End If

        If ComboBox25.Text = "" Then
            MessageBox.Show("Выберите объект общепита!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            Return 1
        End If

        If ComboBox19.Text = "" And CheckBox5.Checked = True And CheckBox7.Checked = True Then

        End If






        Return 0
    End Function
    Private Sub ДогПодряда()

        If ДогПодрядаПроверка() = 1 Then Exit Sub





        Me.Cursor = Cursors.WaitCursor
        Статистика1(Trim(TextBox1.Text) & " " & Trim(TextBox2.Text) & " " & Trim(TextBox3.Text), "Принятие договор подряда", ComboBox1.Text)
        Dim IDСотрудника As Integer




        'добавляем новый договор
        If (CheckBox5.Checked = False Or Решение = "Подряд") And CheckBox7.Checked = True Then

            Dim list As New Dictionary(Of String, Object)
            list.Add("@НазвОрганиз", ComboBox1.Text)
            list.Add("@Фамилия", Trim(TextBox1.Text))
            list.Add("@Имя", Trim(TextBox2.Text))
            list.Add("@Отчество", Trim(TextBox3.Text))
            list.Add("@ФамилияРодПад", Trim(TextBox6.Text))
            list.Add("@ИмяРодПад", Trim(TextBox5.Text))
            list.Add("@ОтчествоРодПад", Trim(TextBox4.Text))
            list.Add("@ПаспортСерия", TextBox12.Text)
            list.Add("@ПаспортНомер", TextBox7.Text)
            list.Add("@ПаспортКогдаВыдан", MaskedTextBox1.Text)
            list.Add("@ДоКакогоДейств", MaskedTextBox2.Text)
            list.Add("@ПаспортКемВыдан", TextBox9.Text)
            list.Add("@ИДНомер", TextBox8.Text)
            list.Add("@Регистрация", TextBox21.Text)
            list.Add("@МестоПрожив", TextBox20.Text)
            list.Add("@КонтТелГор", TextBox37.Text)
            list.Add("@КонтТелефон", MaskedTextBox10.Text)
            list.Add("@СтраховойПолис", TextBox45.Text)
            list.Add("@НаличеДогПодряда", "Да")
            list.Add("@Пол", ComboBox28.Text)

            If CheckBox1.Checked = True Then
                list.Add("@Иностранец", "True")
            Else
                list.Add("@Иностранец", "False")
            End If

            list.Add("@ФамилияДляЗаявления", Trim(TextBox34.Text))
            list.Add("@ИмяДляЗаявления", Trim(TextBox11.Text))
            list.Add("@ОтчествоДляЗаявления", Trim(TextBox10.Text))
            list.Add("@ФИОСборное", Trim(TextBox1.Text) & " " & Trim(TextBox2.Text) & " " & Trim(TextBox3.Text))
            list.Add("@ФИОРодПод", Trim(TextBox6.Text) & " " & Trim(TextBox5.Text) & " " & Trim(TextBox4.Text))
            list.Add("@ТипОтношения", "(дп)")


            IDСотрудника = Updates(stroka:="INSERT INTO Сотрудники(НазвОрганиз,Фамилия,Имя,Отчество,ФамилияРодПад,ИмяРодПад,ОтчествоРодПад,ПаспортСерия,
ПаспортНомер,ПаспортКогдаВыдан,ДоКакогоДейств,ПаспортКемВыдан,ИДНомер,Регистрация,МестоПрожив,КонтТелГор,КонтТелефон,СтраховойПолис,
НаличеДогПодряда,Пол,Иностранец,ФамилияДляЗаявления,ИмяДляЗаявления,ОтчествоДляЗаявления,ФИОСборное,ФИОРодПод,ТипОтношения)
            VALUES(@НазвОрганиз,@Фамилия,@Имя,@Отчество,@ФамилияРодПад,@ИмяРодПад,@ОтчествоРодПад,@ПаспортСерия,@ПаспортНомер,@ПаспортКогдаВыдан,
@ДоКакогоДейств,@ПаспортКемВыдан,@ИДНомер,@Регистрация,@МестоПрожив,@КонтТелГор,@КонтТелефон,@СтраховойПолис,@НаличеДогПодряда,@Пол,
@Иностранец,@ФамилияДляЗаявления,@ИмяДляЗаявления,@ОтчествоДляЗаявления,@ФИОСборное,@ФИОРодПод,@ТипОтношения); SELECT SCOPE_IDENTITY()", list, "Сотрудники", 1) 'возврат ID




            '            Dim strsql1 As String = "SELECT КодСотрудники From Сотрудники WHERE НазвОрганиз='" & ComboBox1.Text & "' AND Фамилия='" & Trim(TextBox1.Text) & "' AND 
            'Имя='" & Trim(TextBox2.Text) & "' AND Отчество='" & Trim(TextBox3.Text) & "'"
            '                Dim ds3 As DataTable = Selects(strsql1)

            '                IDСотрудника = ds3.Rows(0).Item(0)

            If TextBox39.Text <> "" Then
                ДПодНом = Me.TextBox55.Text & "." & TextBox39.Text
            Else
                ДПодНом = Me.TextBox55.Text
            End If

        ElseIf CheckBox7.Checked = True And CheckBox5.Checked = True Then
            Try
                IDСотрудника = CType(Label96.Text, Integer)
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            If TextBox39.Text <> "" Then
                ДПодНом = Me.TextBox55.Text & "." & TextBox39.Text
            Else
                ДПодНом = Me.TextBox55.Text
            End If
            Dim _ВидДоговора As String = "час"
            If MessageBox.Show("Если изменить действующий договор подряда" & vbCrLf & "выберите - Да'" & vbCrLf & "Если создать новый" & vbCrLf & "выберите - Нет'", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                ОбнДогПодр()
            ElseIf ComboBox27.Text = "час" Then
                Чист()
                StrSql = "INSERT INTO ДогПодряда(ID,НомерДогПодр,ДатаДогПодр,Должность,ДатаНачала,ДатаОконч,СтоимЧасаРуб,СтоимЧасаКоп,ОбъекОбщепита,Примечание,ВидДоговора)
VALUES(" & IDСотрудника & ",'" & ДПодНом & "','" & Me.MaskedTextBox6.Text & "','" & Me.ComboBox22.Text & "','" & Me.MaskedTextBox7.Text & "',
'" & Me.MaskedTextBox8.Text & "','" & Me.TextBox61.Text & "','" & Me.TextBox62.Text & "','" & Me.ComboBox25.Text & "','" & Примечани & "','" & _ВидДоговора & "')"
                Updates(StrSql)

            ElseIf ComboBox27.Text = "иное" Then

                For i As Integer = 0 To ДогПодрВыпРаб.Count - 1
                    Dim mn As Object
                    mn = ДогПодЦиклРабот(i)
                    Чист()
                    StrSql = "INSERT INTO ДогПодряда(ID,НомерДогПодр,ДатаДогПодр,Должность,ДатаНачала,ДатаОконч,СтоимРуб1,СтоимКоп1,ОбъекОбщепита,Примечание,ВыпРаб1,ВидИзм)
            VALUES(" & IDСотрудника & ",'" & ДПодНом & "','" & Me.MaskedTextBox6.Text & "','" & Me.ComboBox22.Text & "','" & Me.MaskedTextBox7.Text & "',
            '" & Me.MaskedTextBox8.Text & "','" & ДогПодрВыпРабСтР(i) & "','" & ДогПодрВыпРабСтК(i) & "','" & Me.ComboBox25.Text & "','" & Примечани & "','" & ДогПодрВыпРаб(i) & "','" & ДогПодрВыпРабСтОб(i) & "')"
                    Updates(StrSql)
                Next
            End If

        End If

        If (CheckBox5.Checked = False Or Решение = "Подряд") And ComboBox27.Text = "час" Then
            Dim _ВидДоговора As String = "час"
            Dim strsql = "INSERT INTO ДогПодряда(ID,НомерДогПодр,ДатаДогПодр,Должность,ДатаНачала,ДатаОконч,СтоимЧасаРуб,СтоимЧасаКоп,ОбъекОбщепита,Примечание,ВидДоговора)
VALUES(" & IDСотрудника & ",'" & ДПодНом & "','" & Me.MaskedTextBox6.Text & "','" & Me.ComboBox22.Text & "','" & Me.MaskedTextBox7.Text & "',
'" & Me.MaskedTextBox8.Text & "','" & Me.TextBox61.Text & "','" & Me.TextBox62.Text & "','" & Me.ComboBox25.Text & "','" & Примечани & "','" & _ВидДоговора & "')"
            Updates(strsql)
        ElseIf CheckBox5.Checked = False And ComboBox27.Text = "иное" Then
            For i As Integer = 0 To ДогПодрВыпРаб.Count - 1
                Dim mn As Object
                mn = ДогПодЦиклРабот(i)
                Чист()
                StrSql = "INSERT INTO ДогПодряда(ID,НомерДогПодр,ДатаДогПодр,Должность,ДатаНачала,ДатаОконч,СтоимРуб1,СтоимКоп1,ОбъекОбщепита,Примечание,ВыпРаб1,ВидИзм)
            VALUES(" & IDСотрудника & ",'" & ДПодНом & "','" & Me.MaskedTextBox6.Text & "','" & Me.ComboBox22.Text & "','" & Me.MaskedTextBox7.Text & "',
            '" & Me.MaskedTextBox8.Text & "','" & ДогПодрВыпРабСтР(i) & "','" & ДогПодрВыпРабСтК(i) & "','" & Me.ComboBox25.Text & "','" & Примечани & "','" & ДогПодрВыпРаб(i) & "','" & ДогПодрВыпРабСтОб(i) & "')"
                Updates(StrSql)
            Next
        End If

        If Поток.IsAlive Or Поток1.IsAlive Then
            Поток.Join()
            Поток1.Join()

        End If

        RunMoving2()

        ComboBox19.AutoCompleteCustomSource.Clear()
        ComboBox19.Items.Clear()
        ComboBox26.Items.Clear()


        dtSotrudnikiAll.DefaultView.Sort = "ФИОСборное" & " ASC"            'по возрастанию
        dtSotrudnikiAll.Select("", "ФИОСборное")

        Dim var = From x In dtSotrudnikiAll.AsEnumerable Where Not x.IsNull("НазвОрганиз") AndAlso x.Item("НазвОрганиз") = Клиент Select x 'рабочий linq для заполнения комбобоксов

        Dim var1 = From x In var.AsEnumerable Order By x.Item("ФИОСборное")   'рабочий linq для заполнения комбобоксов  и order by



        For Each r As DataRow In var1
            ComboBox19.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            ComboBox19.Items.Add(Trim(r("ФИОСборное").ToString & "" & r("ТипОтношения").ToString))
            'Me.ComboBox19.Items.Add(r(1).ToString)
            ComboBox26.Items.Add(r("КодСотрудники").ToString)
        Next
        ComboBox19.Text = ""






        'Parallel.Invoke(Sub() com19collection())
        ДокиПодряда(массивДогПодр, массив2)

        Me.Cursor = Cursors.Default
        CheckBox6.Checked = True
        CheckBox6.Checked = False
        Обнов1()
        Решение = ""
    End Sub
    Private Sub НалогиИОбязанДогПодряда()

        'Выгружаем с базы обязанннсоти по должности переменную
        Dim StrSql9 As String = "SELECT ДогПодрОбязан.Обязанности
FROM ДогПодДолжн INNER JOIN ДогПодрОбязан ON ДогПодДолжн.Код = ДогПодрОбязан.ID
WHERE ДогПодДолжн.Клиент='" & arrtcom("ComboBox1") & "' AND ДогПодДолжн.Должность='" & arrtcom("ComboBox22") & "'"
        Dim ds9 As DataTable = Selects(StrSql9)

        ОбязН = ""

        For Each rd As DataRow In ds9.Rows
            ОбязН &= "● " & rd(0).ToString & ";" & vbCrLf
        Next


        If arrtcom("ComboBox27") = "час" Then
            массив2 = Налоги(arrtbox("TextBox61"), arrtbox("TextBox62")).ToArray
        Else
            массив2 = {ОбязН}
        End If
    End Sub
    Private Function ДогПодЦиклРабот(ByVal a As Integer) As Object
        Select Case a
            Case 0
                Return {"СтоимРуб1", "СтоимКоп1", "ВыпРаб1"}
            Case 1
                Return {"СтоимРуб2", "СтоимКоп2", "ВыпРаб2"}
            Case 2
                Return {"СтоимРуб3", "СтоимКоп3", "ВыпРаб3"}
            Case 3
                Return {"СтоимРуб4", "СтоимКоп4", "ВыпРаб4"}
        End Select

    End Function
    Private Sub СортДогПод(ByVal d As String)
        Дпод1 = ""
        Дпод2 = ""
        Select Case d
            Case "час"
                Дпод1 = "потраченных часов"
            Case "иное"
                Дпод1 = "выполненных работ"
                ВидыРаботДогПодряда.ShowDialog()
        End Select
    End Sub
    Public Sub ДокиПодряда(ByVal массив() As String, ByVal массив2() As String)
        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document
        'Dim oWordPara As Microsoft.Office.Interop.Word.Paragraph

        'KillProc()

        oWord = CreateObject("Word.Application")
        oWord.Visible = False


        ВыгрузкаФайловНаЛокалыныйКомп(FTPStringAllDOC & "DPodriada.doc", firthtPath & "\DPodriada.doc")

        oWordDoc = oWord.Documents.Add(firthtPath & "\DPodriada.doc")

        With oWordDoc.Bookmarks
            .Item("ДП1").Range.Text = ДПодНом & " - " & Now.Year
            .Item("ДП2").Range.Text = arrtmask("MaskedTextBox6")
            .Item("ДП3").Range.Text = ФормаСобстПолн

            If ФормаСобстПолн = "Индивидуальный предприниматель" Then
                .Item("ДП4").Range.Text = arrtcom("ComboBox1")
                .Item("ДП5").Range.Text = ФИОРукРодПад
                .Item("ДП16").Range.Text = Strings.LCase(массив(1)) & " " & arrtcom("ComboBox1") & " " 'индивидуальный предпринматель 
                .Item("ДП27").Range.Text = arrtcom("ComboBox1")
            Else
                .Item("ДП4").Range.Text = " «" & arrtcom("ComboBox1") & "» "
                If arrtcom("ComboBox1") = "Итал Гэлэри Плюс" Then
                    Dim l As String = Strings.Left(ДолжРуковВинПад, 1)
                    l = Strings.LCase(l)
                    Dim d2 As Integer = ДолжРуковВинПад.Length - 1
                    ДолжРуковВинПад = l & Strings.Right(ДолжРуковВинПад, d2)
                    .Item("ДП5").Range.Text = ДолжРуковВинПад & " " & ФИОРукРодПад
                    Dim f As String = Strings.Left(массив(1), 1)
                    f = Strings.LCase(f)
                    Dim d1 As Integer = массив(1).Length - 1
                    массив(1) = f & Strings.Right(массив(1), d1)
                    .Item("ДП16").Range.Text = массив(1) & " " & ФормаСобствКор & " «" & arrtcom("ComboBox1") & "» "
                    'директор ООО "назв орг" ds.Rows(0).Item(1).ToString
                    .Item("ДП27").Range.Text = " «" & arrtcom("ComboBox1") & "» "
                Else
                    .Item("ДП5").Range.Text = Strings.LCase(ДолжРуковВинПад) & " " & ФИОРукРодПад
                    .Item("ДП16").Range.Text = Strings.LCase(массив(1)) & " " & ФормаСобствКор & " «" & arrtcom("ComboBox1") & "» "
                    'директор ООО "назв орг" ds.Rows(0).Item(1).ToString
                    .Item("ДП27").Range.Text = " «" & arrtcom("ComboBox1") & "» "
                End If
            End If



            If ComboBox1.Text = "Итал Гэлэри Плюс" Then
                .Item("ДП6").Range.Text = ""
                .Item("ДП41").Range.Text = ""
            Else
                .Item("ДП6").Range.Text = массив(3)
            End If
            'ds.Rows(0).Item(3).ToString
            .Item("ДП7").Range.Text = arrtbox("TextBox1") & " " & arrtbox("TextBox2") & " " & arrtbox("TextBox3")
            .Item("ДП8").Range.Text = МестоРаб
            'For ipd As Integer = LBound(arr2) To UBound(arr2)
            .Item("ДП9").Range.Text = массив2(0) 'ОбязН
            'Next


            .Item("ДП17").Range.Text = arrtmask("MaskedTextBox7")
            .Item("ДП18").Range.Text = arrtmask("MaskedTextBox8")
            '.Item("ДП19").Range.Text = TextBox61.Text
            '.Item("ДП20").Range.Text = TextBox62.Text
            '.Item("ДП21").Range.Text = массив2(1) 'ПДЦелаяЧасть
            '.Item("ДП22").Range.Text = массив2(2) 'ПДДробнаяЧасть
            '.Item("ДП23").Range.Text = массив2(3) 'БГСЦелаяЧасть
            '.Item("ДП24").Range.Text = массив2(4) 'БГСДробнаяЧасть
            .Item("ДП25").Range.Text = МестоРаб
            .Item("ДП26").Range.Text = ФормаСобствКор

            .Item("ДП28").Range.Text = массив(5) 'ds.Rows(0).Item(5).ToString 'ЮрАдрес
            .Item("ДП29").Range.Text = массив(4) 'ds.Rows(0).Item(4).ToString 'унп
            .Item("ДП30").Range.Text = массив(6) 'ds.Rows(0).Item(6).ToString
            .Item("ДП31").Range.Text = массив(7) 'ds.Rows(0).Item(7).ToString
            .Item("ДП32").Range.Text = массив(8) 'ds.Rows(0).Item(8).ToString
            .Item("ДП33").Range.Text = arrtbox("TextBox1") & " " & arrtbox("TextBox2") & " " & arrtbox("TextBox3")
            .Item("ДП34").Range.Text = arrtbox("TextBox12") & " " & arrtbox("TextBox7")
            .Item("ДП35").Range.Text = arrtmask("MaskedTextBox1")
            .Item("ДП36").Range.Text = arrtbox("TextBox9")
            .Item("ДП37").Range.Text = arrtbox("TextBox21")
            .Item("ДП38").Range.Text = arrtbox("TextBox8")
            .Item("ДП39").Range.Text = ФИОКор
            .Item("ДП40").Range.Text = arrtbox("TextBox1") & " " & Strings.Left(arrtbox("TextBox2"), 1) & "." & Strings.Left(arrtbox("TextBox3"), 1) & "."
            .Item("ДП42").Range.Text = Дпод1

            If ComboBox27.Text = "час" Then
                .Item("ДП43").Range.Text = "Стоимость часа работы – " & arrtbox("TextBox61") & "р " & arrtbox("TextBox62") & "коп, в том числе: подоходный налог – " & массив2(1) & "р " & массив2(2) & "коп.; отчисления в пенсионный фонд – " & массив2(3) & "р " & массив2(4) & "коп."
            Else
                .Item("ДП43").Range.Text = Дпод2
            End If

        End With


        Dim dirstring As String = Клиент & "/Договор подряда/" & Now.Year & "/" 'место сохранения файла

        dirstring = СозданиепапкиНаСервере(dirstring) 'полный путь на сервер(кроме имени и разрешения файла)


        Dim put, Name As String
        Name = ДПодНом & " " & arrtbox("TextBox1") & " от " & arrtmask("MaskedTextBox6") & "(Договор подряда)" & ".doc"
        put = PathVremyanka & Name 'место в корне программы

        ВыборкаИзагрНаСервер(dirstring, Name, "Прием-Дог Подряд")

        oWordDoc.SaveAs2(put,,,,,, False)
        oWordDoc.Close(True)
        oWord.Quit(True)

        СохрДогПодрFTP.AddRange(New String() {dirstring, Name})
        dirstring += Name


        ЗагрНаСерверИУдаление(put, dirstring, put)

        ВременнаяПапкаУдалениеФайла(firthtPath & "\DPodriada.doc")




    End Sub
    Private Sub ВыборкаИзагрНаСервер(ByVal dirstring As String, ByVal Name As String, ByVal НазДок As String)

        If CheckBox5.Checked = True Then
            Parallel.Invoke(Sub() ЗагрВБазуПутиДоков2(CType(Label96.Text, Integer), dirstring, Name, НазДок, arrtcom("ComboBox1")))
        Else
            Dim b = dtSotrudnikiAll.Select("Фамилия='" & txt1 & "' and ПаспортНомер='" & arrtbox("TextBox7") & "' and ИДНомер='" & arrtbox("TextBox8") & "'") 'выбираем данные по сотруднику
            Dim kd As Integer = CType(b(0).Item("КодСотрудники").ToString, Integer) 'находим ИД сотрудника
            Parallel.Invoke(Sub() ЗагрВБазуПутиДоков2(kd, dirstring, Name, НазДок, arrtcom("ComboBox1"))) 'заполняем данные путей и назв файла
        End If



    End Sub
    Public Function ПровДляКонтр()
        Me.Cursor = Cursors.WaitCursor

        If CheckBox5.Checked = True And ComboBox19.Text = "" Then
            MessageBox.Show("Выберите сотрудника для изменения!", Рик)
            Return 1
        End If





        If CheckBox1.Checked = False And MaskedTextBox1.Text = "" Or MaskedTextBox2.Text = "" Then
            MessageBox.Show("Вы не выбрали дату выдачи или дату срока действия паспорта!", Рик)
            Return 1
        End If

        If TextBox56.Text = "" Then
            If MessageBox.Show("Вы не выбрали дату выплаты аванса! Выбрать?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Information) = DialogResult.Yes Then
                Return 1
            End If
        End If
        If TextBox40.Text = "" Then
            MessageBox.Show("Выберите дату выплаты зарплаты!", Рик)
            Return 1
        End If

        If CheckBox1.Checked = False Then
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Then
                MessageBox.Show("Заполните данные сотрудника ФИО!", Рик)
                Return 1
            End If
        End If

        If CheckBox5.Checked = False Then
            If Not Примечани <> "" Then
                If MessageBox.Show("Вы НЕ заполнили примечание!" & vbCrLf & "Выберите OK - если хотите продолжить, или ОТМЕНА - если хотите изменить", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.Cancel Then
                    Return 1
                End If

            End If
        End If


        If Not ComboBox18.Text <> "" Then
            MessageBox.Show("Выберите объект общепита!", Рик)
            Return 1
        End If
        If CheckBox26.Checked = True Then
            If Not ComboBox8.Text <> "" Then
                MessageBox.Show("Выберите отдел!", Рик)
                Return 1
            End If
            If Not ComboBox9.Text <> "" Then
                MessageBox.Show("Выберите должность!", Рик)
                Return 1
            End If

            If ComboBox7.Items.Count = 1 Then
                If ComboBox7.Enabled = True And Not ComboBox7.Text <> "" Then
                    MessageBox.Show("Выберите разряд!", Рик)
                    Return 1
                End If
            End If


        End If

        If Not ComboBox10.Text <> "" Then
            MessageBox.Show("Выберите ставку!", Рик)
            Return 1
        End If
        If Not TextBox7.TextLength = 7 And CheckBox1.Checked = False Then 'проверяем заполненность поля номер паспорта кол-во цифр
            MessageBox.Show("Проверьте поле номер паспорта!", Рик, MessageBoxButtons.OK)
            Return 1
        End If
        If Not TextBox8.TextLength = 14 And CheckBox1.Checked = False Then
            MessageBox.Show("Неправильно заполнено поле 'Идентификационный номер'", Рик)
            Return 1
        End If

        If TextBox57.Text <> "" Then
            TextBox57.Text = " - " & TextBox57.Text
        End If
        If TextBox57.Text <> "" Then
            TextBox57.Text = " - " & TextBox57.Text
        End If

        If ПровИндивидКонтр(ComboBox1.Text) = True Then
            If MaskedTextBox9.MaskCompleted = False Then
                MessageBox.Show("Неправильно заполнено поле 'Дата рождения'", Рик)
                Return 1
            End If
            If TextBox51.Text = "" Then
                MessageBox.Show("Заполните поле 'Сотрудник - гражданин какой страны?'", Рик)
                Return 1
            End If
        End If

        If IsNumeric(TextBox41.Text) Then
            НПриказа = TextBox41.Text & " - " & TextBox58.Text & TextBox57.Text
        Else
            НПриказа = TextBox41.Text
        End If
        If Not ComboBox28.Text <> "" Then
            MessageBox.Show("Выберите пол сотрудника!", Рик)
            Return 1
        End If
        Dim парк As String = ПроверкаЗаполненности(Должность)
        If парк = Nothing Then
            Return 1
        End If
        Dim от1, дол1, разр1 As String

        If CheckBox5.Checked = True And CheckBox26.Checked = False Then
            Dim strsql85 As String = "SELECT Отдел,Должность,Разряд FROM Штатное WHERE ИДСотр=" & CType(Label96.Text, Integer) & ""
            Dim hk As DataTable = Selects(strsql85)
            от1 = hk.Rows(0).Item(0).ToString
            дол1 = hk.Rows(0).Item(1).ToString
            разр1 = hk.Rows(0).Item(2).ToString
        Else
            от1 = ComboBox8.Text
            дол1 = ComboBox9.Text
            разр1 = ComboBox7.Text
        End If

        ПровИнстр = Nothing
        StrSql = "SELECT ШтСвод.ДолжИнструкция FROM ШтОтделы INNER JOIN ШтСвод ON ШтОтделы.Код = ШтСвод.Отдел
WHERE ШтОтделы.Клиент='" & ComboBox1.Text & "' AND ШтОтделы.Отделы='" & от1 & "' AND ШтСвод.Должность='" & дол1 & "'
AND ШтСвод.Разряд='" & разр1 & "'"
        ds = Selects(StrSql)
        Try
            If ds.Rows(0).Item(0).ToString = "False" Then
                If MessageBox.Show("Для данной должности не сформирована должностная инструкция!" & vbCrLf & "Оформить инструкцию?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                    ПровИнстр = 1
                    Return 0
                Else
                    v = False
                    ДолжИнстр.ShowDialog()
                    If ДолжИнстр.текст = "" Or ДолжИнстр.Ном = "" Then
                        If MessageBox.Show("Вы не заполнили номер или текст инструкции!" & vbCrLf & "Все равно продолжить?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.Cancel Then
                            v = False
                            ДолжИнстр.ShowDialog()
                        End If
                    End If

                    If v = False Then
                        ПровИнстр = 1
                        Return 0
                    Else
                        ДокиИнструкция()
                        ПровИнстр = Nothing
                        Return 0
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            MessageBox.Show("Проверьте разряд сотрудника, он не совпадает с данными в штатном расписании!", Рик)
            Return 1
        End Try


        Me.Cursor = Cursors.Default
        Return 0
    End Function
    Private Sub ДокиИнструкцияНовыйПуть()

        Me.Cursor = Cursors.WaitCursor

        Dim ds As Клиент
        Using dbcx As New DbAll1DataContext
            ds = (From x In dbcx.Клиент.AsEnumerable
                  Where x.НазвОрг = ComboBox1.Text
                  Select x).FirstOrDefault()
        End Using

        'начало создание документа
        Dim oWord As New Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document
        Начало("Instrukciya.docx")

        oWordDoc = oWord.Documents.Add(firthtPath & "\Instrukciya.docx")

        With oWordDoc.Bookmarks

            If ds.ФормаСобств = "Индивидуальный предприниматель" Then

                .Item("Инстр1").Range.Text = ФормСобствКор(ds.ФормаСобств) & " " & ComboBox1.Text
            Else
                .Item("Инстр1").Range.Text = ФормСобствКор(ds.ФормаСобств) & " «" & ComboBox1.Text & "»"
            End If

            .Item("Инстр2").Range.Text = ДолжИнстр.Дат
            .Item("Инстр3").Range.Text = ДолжИнстр.Ном

            'руководитель ИП
            If ds.РукИП = "True" Then
                Dim ФИО As String = ФИОКорРук(ds.ФИОРуководителя, True)
                .Item("Инстр4").Range.Text = ds.ДолжнРуководителя & " " & ФИО
                .Item("Инстр5").Range.Text = ФИО
                .Item("Инстр9").Range.Text = ФИО

            Else
                'руководитель не ИП
                Dim ФИО1 As String = ФИОКорРук(ds.ФИОРуководителя, False)
                If ds.ФормаСобств = "Индивидуальный предприниматель" Then
                    .Item("Инстр4").Range.Text = ds.ДолжнРуководителя
                Else
                    .Item("Инстр4").Range.Text = ds.ДолжнРуководителя & " " & ФИО1
                End If
                .Item("Инстр5").Range.Text = ФИО1
                .Item("Инстр9").Range.Text = ФИО1

            End If

            .Item("Инстр6").Range.Text = ДолжИнстр.Дат
            .Item("Инстр7").Range.Text = ДолжИнстр.текст
            .Item("Инстр8").Range.Text = ds.ДолжнРуководителя
            .Item("Инстр10").Range.Text = Now.Year
            .Item("Инстр11").Range.Text = Now.Year

        End With

        Dim dirstring As String = ComboBox1.Text & "/Должностные инструкции/" 'место сохранения файла

        dirstring = СозданиепапкиНаСервере(dirstring) 'полный путь на сервер(кроме имени и разрешения файла)


        Dim put, Name As String

        'If (ComboBox7.Text = "" Or ComboBox7.Text = "-") And CheckBox26.Checked = False And CheckBox5.Checked = True Then
        '    Dim dr = dtShtatnoeAll.Select("ИДСотр=" & CType(Label96.Text, Integer) & "")

        '    Name = ДолжИнстр.Ном & " " & dr(0).Item("Отдел").ToString & " " & dr(0).Item("Должность").ToString & " " & dr(0).Item("Разряд").ToString & ".doc"

        'ElseIf (ComboBox7.Text = "" Or ComboBox7.Text = "-") And CheckBox26.Checked = True And CheckBox5.Checked = True Then

        '    Name = ДолжИнстр.Ном & " " & Trim(ComboBox8.Text) & " " & Trim(ComboBox9.Text) & ".doc"

        'ElseIf (ComboBox7.Text = "" Or ComboBox7.Text = "-") And CheckBox5.Checked = False Then
        '    Name = ДолжИнстр.Ном & " " & Trim(ComboBox8.Text) & " " & Trim(ComboBox9.Text) & ".doc"
        'Else


        '    Name = ДолжИнстр.Ном & " " & Trim(ComboBox8.Text) & " " & Trim(ComboBox9.Text) & " " & Trim(ComboBox7.Text) & ".doc"

        'End If



        Name = ДолжИнстр.Ном & " " & Trim(ComboBox8.Text) & " " & Trim(ComboBox9.Text) & " " & Trim(ComboBox7.Text) & ".doc"
        put = PathVremyanka & Name 'место в корне программы

        'ВыборкаИзагрНаСервер(dirstring, Name, "Прием-Инструкция")

        'Dim b = dtSotrudnikiAll.Select("ФИОСборное='" & combx19 & "'") 'выбираем данные по сотруднику
        'Dim kd As Integer = CType(b(0).Item("КодСотрудники").ToString, Integer) 'находим ИД сотрудника
        'ЗагрВБазуПутиДоковAsync(kd, dirstring, Name, "Прием-Зявление") 'заполняем данные путей и назв файла

        oWordDoc.SaveAs2(put,,,,,, False)
        dirstring += Name

        oWordDoc.Close(True)
        oWord.Quit(True)

        ЗагрНаСерверИУдаление(put, dirstring, put)

        'конец создание документа

        'выбираем из Штсвод номер и обновляем данные
        Dim ШтСв As ШтСвод
        Using dbc As New DbAll1DataContext
            ШтСв = (From x In dbc.Клиент.AsEnumerable
                    Join y In dbc.ШтОтделы.AsEnumerable On x.НазвОрг Equals y.Клиент
                    Join z In dbc.ШтСвод.AsEnumerable On y.Код Equals z.Отдел
                    Where x.НазвОрг = ComboBox1.Text And y.Отделы = ComboBox8.Text _
                          And z.Должность = ComboBox9.Text And z.Разряд = ComboBox7.Text
                    Select z).FirstOrDefault()


            If ШтСв IsNot Nothing Then
                ШтСв.ДолжИнструкция = "True"

                'если разряд число, то соединяем.
                If ComboBox7.Text <> "" And IsNumeric(ComboBox7.Text) Then
                    ШтСв.НомерДолжИнстр = ДолжИнстр.Ном & " " & Trim(ComboBox8.Text) & " " & Trim(ComboBox9.Text) & " " & Trim(ComboBox7.Text)
                Else
                    ШтСв.НомерДолжИнстр = ДолжИнстр.Ном & " " & Trim(ComboBox8.Text) & " " & Trim(ComboBox9.Text)
                End If

                ШтСв.ТекстИнструкции = ДолжИнстр.текст
                ШтСв.ДатаИнструкции = ДолжИнстр.Дат
                dbc.SubmitChanges()
            End If
        End Using

        Me.Cursor = Cursors.Default

    End Sub



    Private Sub ДокиИнструкция()

        Dim ds = dtClientAll.Select("НазвОрг='" & arrtcom("ComboBox1") & "'")

        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document
        'Dim oWordPara As Microsoft.Office.Interop.Word.Paragraph

        'KillProc()

        oWord = CreateObject("Word.Application")
        oWord.Visible = False
        Начало("Instrukciya.docx")

        oWordDoc = oWord.Documents.Add(firthtPath & "\Instrukciya.docx")

        With oWordDoc.Bookmarks

            If ds(0).Item(1).ToString = "Индивидуальный предприниматель" Then

                .Item("Инстр1").Range.Text = ФормСобствКор(ds(0).Item(1).ToString) & " " & ComboBox1.Text
            Else
                .Item("Инстр1").Range.Text = ФормСобствКор(ds(0).Item(1).ToString) & " «" & ComboBox1.Text & "»"
            End If

            .Item("Инстр2").Range.Text = ДолжИнстр.Дат
            .Item("Инстр3").Range.Text = ДолжИнстр.Ном

            If ds(0).Item(31) = "True" Then
                .Item("Инстр4").Range.Text = ds(0).Item(18).ToString & " " & ФИОКорРук(ds(0).Item(19).ToString, True)
                .Item("Инстр5").Range.Text = ФИОКорРук(ds(0).Item(19).ToString, True)
                .Item("Инстр9").Range.Text = ФИОКорРук(ds(0).Item(19).ToString, True)
            Else
                If ds(0).Item(1).ToString = "Индивидуальный предприниматель" Then
                    .Item("Инстр4").Range.Text = ds(0).Item(18).ToString
                Else
                    .Item("Инстр4").Range.Text = ds(0).Item(18).ToString & " " & ФИОКорРук(ds(0).Item(19).ToString, False)
                End If

                .Item("Инстр5").Range.Text = ФИОКорРук(ds(0).Item(19).ToString, False)
                .Item("Инстр9").Range.Text = ФИОКорРук(ds(0).Item(19).ToString, False)
            End If

            .Item("Инстр6").Range.Text = ДолжИнстр.Дат
            .Item("Инстр7").Range.Text = ДолжИнстр.текст
            .Item("Инстр8").Range.Text = ds(0).Item(18).ToString
            .Item("Инстр10").Range.Text = Now.Year
            .Item("Инстр11").Range.Text = Now.Year

        End With

        Dim dirstring As String = arrtcom("ComboBox1") & "/Должностные инструкции/" 'место сохранения файла

        dirstring = СозданиепапкиНаСервере(dirstring) 'полный путь на сервер(кроме имени и разрешения файла)


        Dim put, Name As String

        If (ComboBox7.Text = "" Or ComboBox7.Text = "-") And CheckBox26.Checked = False And CheckBox5.Checked = True Then
            Dim dr = dtShtatnoeAll.Select("ИДСотр=" & CType(Label96.Text, Integer) & "")

            Name = ДолжИнстр.Ном & " " & dr(0).Item("Отдел").ToString & " " & dr(0).Item("Должность").ToString & " " & dr(0).Item("Разряд").ToString & ".doc"

        ElseIf (ComboBox7.Text = "" Or ComboBox7.Text = "-") And CheckBox26.Checked = True And CheckBox5.Checked = True Then

            Name = ДолжИнстр.Ном & " " & Trim(ComboBox8.Text) & " " & Trim(ComboBox9.Text) & ".doc"

        ElseIf (ComboBox7.Text = "" Or ComboBox7.Text = "-") And CheckBox5.Checked = False Then
            Name = ДолжИнстр.Ном & " " & Trim(ComboBox8.Text) & " " & Trim(ComboBox9.Text) & ".doc"
        Else


            Name = ДолжИнстр.Ном & " " & Trim(ComboBox8.Text) & " " & Trim(ComboBox9.Text) & " " & Trim(ComboBox7.Text) & ".doc"

        End If

        put = PathVremyanka & Name 'место в корне программы

        'ВыборкаИзагрНаСервер(dirstring, Name, "Прием-Инструкция")

        'Dim b = dtSotrudnikiAll.Select("ФИОСборное='" & combx19 & "'") 'выбираем данные по сотруднику
        'Dim kd As Integer = CType(b(0).Item("КодСотрудники").ToString, Integer) 'находим ИД сотрудника
        'ЗагрВБазуПутиДоковAsync(kd, dirstring, Name, "Прием-Зявление") 'заполняем данные путей и назв файла

        oWordDoc.SaveAs2(put,,,,,, False)
        dirstring += Name

        oWordDoc.Close(True)
        oWord.Quit(True)

        ЗагрНаСерверИУдаление(put, dirstring, put)



        dtPutiDokumentov()

        Dim gf As String = "True"

        If CheckBox26.Checked = False And CheckBox5.Checked = True Then
            Dim dr1 = dtShtatnoeAll.Select("ИДСотр=" & CType(Label96.Text, Integer) & "")
            Dim dsj As DataTable = Selects(StrSql:="SELECT ШтСвод.КодШтСвод
FROM ШтОтделы INNER JOIN ШтСвод ON ШтОтделы.Код = ШтСвод.Отдел
WHERE ШтОтделы.Клиент='" & arrtcom("ComboBox1") & "' AND ШтОтделы.Отделы= '" & dr1(0).Item("Отдел").ToString & "' AND
ШтСвод.Должность= '" & dr1(0).Item("Должность").ToString & "' AND ШтСвод.Разряд= '" & dr1(0).Item("Разряд").ToString & "'")

            Updates(stroka:="UPDATE ШтСвод SET ДолжИнструкция='" & gf & "',
НомерДолжИнстр='" & ДолжИнстр.Ном & " " & dr1(0).Item("Отдел").ToString & " " & dr1(0).Item("Должность").ToString & " " & dr1(0).Item("Разряд").ToString & "',
ТекстИнструкции='" & ДолжИнстр.текст & "', ДатаИнструкции='" & ДолжИнстр.Дат & "'
WHERE КодШтСвод=" & dsj.Rows(0).Item(0) & "")

        ElseIf ComboBox7.Text = "" Or ComboBox7.Text = "-" And CheckBox26.Checked = True And CheckBox5.Checked = True Then

            Dim dtv As DataTable = Selects(StrSql:="SELECT ШтСвод.КодШтСвод
FROM(Клиент INNER JOIN ШтОтделы On Клиент.НазвОрг = ШтОтделы.Клиент) INNER JOIN ШтСвод On ШтОтделы.Код = ШтСвод.Отдел
            WHERE Клиент.НазвОрг ='" & ComboBox1.Text & "' AND ШтОтделы.Отделы='" & ComboBox8.Text & "'
AND ШтСвод.Должность='" & ComboBox9.Text & "' AND ШтСвод.Разряд='" & ComboBox7.Text & "'")

            Updates(stroka:="UPDATE ШтСвод SET ДолжИнструкция='" & gf & "',
НомерДолжИнстр='" & ДолжИнстр.Ном & " " & Trim(ComboBox8.Text) & " " & Trim(ComboBox9.Text) & "',
ТекстИнструкции='" & ДолжИнстр.текст & "', ДатаИнструкции='" & ДолжИнстр.Дат & "'
WHERE КодШтСвод=" & dtv.Rows(0).Item(0) & "")

        ElseIf ComboBox7.Text = "" Or ComboBox7.Text = "-" And CheckBox5.Checked = False Then

            Dim dtv As DataTable = Selects(StrSql:="SELECT ШтСвод.КодШтСвод
FROM(Клиент INNER JOIN ШтОтделы On Клиент.НазвОрг = ШтОтделы.Клиент) INNER JOIN ШтСвод On ШтОтделы.Код = ШтСвод.Отдел
            WHERE Клиент.НазвОрг ='" & ComboBox1.Text & "' AND ШтОтделы.Отделы='" & ComboBox8.Text & "'
AND ШтСвод.Должность='" & ComboBox9.Text & "' AND ШтСвод.Разряд='" & ComboBox7.Text & "'")

            Updates(stroka:="UPDATE ШтСвод SET ДолжИнструкция='" & gf & "',
НомерДолжИнстр='" & ДолжИнстр.Ном & " " & Trim(ComboBox8.Text) & " " & Trim(ComboBox9.Text) & "',
ТекстИнструкции='" & ДолжИнстр.текст & "', ДатаИнструкции='" & ДолжИнстр.Дат & "'
WHERE КодШтСвод=" & dtv.Rows(0).Item(0) & "")

        ElseIf ComboBox7.Text <> "" And CheckBox5.Checked = False Then

            Dim dtv As DataTable = Selects(StrSql:="SELECT ШтСвод.КодШтСвод
FROM(Клиент INNER JOIN ШтОтделы On Клиент.НазвОрг = ШтОтделы.Клиент) INNER JOIN ШтСвод On ШтОтделы.Код = ШтСвод.Отдел
            WHERE Клиент.НазвОрг ='" & ComboBox1.Text & "' AND ШтОтделы.Отделы='" & ComboBox8.Text & "'
AND ШтСвод.Должность='" & ComboBox9.Text & "' AND ШтСвод.Разряд='" & ComboBox7.Text & "'")

            Updates(stroka:="UPDATE ШтСвод SET ДолжИнструкция='" & gf & "',
НомерДолжИнстр='" & ДолжИнстр.Ном & " " & Trim(ComboBox8.Text) & " " & Trim(ComboBox9.Text) & " " & Trim(ComboBox7.Text) & "',
ТекстИнструкции='" & ДолжИнстр.текст & "', ДатаИнструкции='" & ДолжИнстр.Дат & "'
WHERE КодШтСвод=" & dtv.Rows(0).Item(0) & "")

        ElseIf ComboBox7.Text <> "" And CheckBox5.Checked = True And CheckBox26.Checked = True Then

            Dim dtv As DataTable = Selects(StrSql:="SELECT ШтСвод.КодШтСвод
FROM(Клиент INNER JOIN ШтОтделы On Клиент.НазвОрг = ШтОтделы.Клиент) INNER JOIN ШтСвод On ШтОтделы.Код = ШтСвод.Отдел
            WHERE Клиент.НазвОрг ='" & ComboBox1.Text & "' AND ШтОтделы.Отделы='" & ComboBox8.Text & "'
AND ШтСвод.Должность='" & ComboBox9.Text & "' AND ШтСвод.Разряд='" & ComboBox7.Text & "'")

            Updates(stroka:="UPDATE ШтСвод SET ДолжИнструкция='" & gf & "',
НомерДолжИнстр='" & ДолжИнстр.Ном & " " & Trim(ComboBox8.Text) & " " & Trim(ComboBox9.Text) & "',
ТекстИнструкции='" & ДолжИнстр.текст & "', ДатаИнструкции='" & ДолжИнстр.Дат & "'
WHERE КодШтСвод=" & dtv.Rows(0).Item(0) & "")

        End If



    End Sub
    Public Function ПровДляПодряда()
        If CheckBox1.Checked = False And MaskedTextBox1.Text = "" Or MaskedTextBox2.Text = "" Then
            MessageBox.Show("Вы не выбрали дату выдачи или дату  срока действия паспорта!", Рик)
            Return 1
        End If

        If CheckBox1.Checked = False And TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Then
            MessageBox.Show("Заполните данные сотрудника ФИО!", Рик)
            Return 1
        End If
        If Not ComboBox25.Text <> "" Then
            MessageBox.Show("Выберите объект общепита!", Рик)
            Return 1
        End If

        If Not ComboBox28.Text <> "" Then
            MessageBox.Show("Выберите пол сотрудника!", Рик)
            Return 1
        End If






        If Not TextBox7.TextLength = 7 And CheckBox1.Checked = False Then 'проверяем заполненность поля номер паспорта кол-во цифр
            MessageBox.Show("Проверьте поле номер паспорта!", Рик, MessageBoxButtons.OK)
            Return 1
        End If
        If Not TextBox8.TextLength = 14 And CheckBox1.Checked = False Then
            MessageBox.Show("Неправильно заполнено поле 'Идентификационный номер'", Рик)

            Return 1
        End If
        If IsNumeric(TextBox55.Text) = False Then
            MessageBox.Show("Номер договора-подряда должен быть целочисленным!", Рик)
            Return 1
        End If




        Return 0
    End Function
    Private Sub ДопЛемеЛ()
        КонтрДопЛемел.ShowDialog()
        Dim strsql As String = "UPDATE КарточкаСотрудника SET НаличиеИспытСрока='" & ЛемелИспытСрок & "', ПериодОтпДляКонтр='" & ЛемелТрОтп & "' WHERE IDСотр=" & IDsot1 & ""
        Updates(strsql)
    End Sub
    Private Function ДопЛемеЛКонтр() As Object

        If CheckBox5.Checked = False Then
            Dim strsql As String = "SELECT НаличиеИспытСрока,ПериодОтпДляКонтр FROM КарточкаСотрудника WHERE IDСотр=" & IDsot1 & ""
            Dim ds As DataTable = Selects(strsql)
            Return {ds.Rows(0).Item(0).ToString, ds.Rows(0).Item(1).ToString}
        Else
            Dim strsql As String = "SELECT НаличиеИспытСрока,ПериодОтпДляКонтр FROM КарточкаСотрудника WHERE IDСотр=" & CType(Label96.Text, Integer) & ""
            Dim ds As DataTable = Selects(strsql)
            Return {ds.Rows(0).Item(0).ToString, ds.Rows(0).Item(1).ToString}
        End If

    End Function
    Private Async Sub ДокПрикаЛемел()
        Await Task.Delay(0)
        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document
        'KillProc()


        oWord = CreateObject("Word.Application")
        oWord.Visible = False


        Try 'проверка если есть в С: папке файл Приказ его удаляем и создаем новый

            IO.File.Copy(OnePath & "\ОБЩДОКИ\Лемел лабс\Prikaz.docx", "C:\Users\Public\Documents\Рик\Prikaz.docx")
        Catch ex As Exception
            If ex.Message.Contains("уже существует") Then
                Try
                    IO.File.Delete("C:\Users\Public\Documents\Рик\Prikaz.docx")
                    IO.File.Copy(OnePath & "\ОБЩДОКИ\Лемел лабс\Prikaz.docx", "C:\Users\Public\Documents\Рик\Prikaz.docx")
                Catch e As System.IO.IOException
                    If e.Message.Contains("используется другим процессом") Then
                        ПрверкаАсинхрПотоков(Task.CurrentId)
                    End If
                End Try
            End If
            IO.File.Delete("C:\Users\Public\Documents\Рик\Prikaz.docx")
            IO.File.Copy(OnePath & "\ОБЩДОКИ\Лемел лабс\Prikaz.docx", "C:\Users\Public\Documents\Рик\Prikaz.docx")
        End Try


        oWordDoc = oWord.Documents.Add("C:\Users\Public\Documents\Рик\Prikaz.docx")

        With oWordDoc.Bookmarks
            .Item("П1").Range.Text = Приказ(5)
            .Item("П2").Range.Text = НПриказа
            .Item("П3").Range.Text = TextBox6.Text
            .Item("П4").Range.Text = CorName
            .Item("П5").Range.Text = CorOtch
            .Item("П6").Range.Text = TextBox6.Text
            .Item("П7").Range.Text = Приказ(3)
            .Item("П8").Range.Text = Приказ(4)

            If combx7 = "-" Then
                .Item("П9").Range.Text = Strings.LCase(ДолжСОконч)
            ElseIf combx7 = "1" Or combx7 = "2" Or combx7 = "3" Or combx7 = "4" Or combx7 = "5" Or combx7 = "6" Then
                .Item("П9").Range.Text = Strings.LCase(ДолжСОконч) & " " & combx7 & " разряда"
            Else
                .Item("П9").Range.Text = Strings.LCase(ДолжСОконч)
            End If


            .Item("П10").Range.Text = Приказ(6)
            .Item("П11").Range.Text = Ставка
            .Item("П12").Range.Text = СтавкаНов
            .Item("П13").Range.Text = СрокКонтр
            .Item("П14").Range.Text = СклонГод
            .Item("П15").Range.Text = Приказ(6)
            .Item("П16").Range.Text = Приказ(7)
            .Item("П17").Range.Text = Приказ(2)
            .Item("П18").Range.Text = CorName
            .Item("П19").Range.Text = CorOtch
            .Item("П20").Range.Text = Приказ(8)
            .Item("П21").Range.Text = Приказ(5)
            .Item("П22").Range.Text = Приказ(9)
            .Item("П23").Range.Text = CorName
            .Item("П24").Range.Text = CorOtch
            .Item("П25").Range.Text = ФормаСобстПолн

            If ФормаСобстПолн = "Индивидуальный предприниматель" Then
                .Item("П26").Range.Text = Клиент
            Else
                .Item("П26").Range.Text = " «" & Клиент & "» "
            End If

            .Item("П27").Range.Text = ЮрАдрес
            .Item("П28").Range.Text = УНП
            .Item("П29").Range.Text = РасСчет
            .Item("П30").Range.Text = АдресБанка
            .Item("П31").Range.Text = БИК
            .Item("П33").Range.Text = ЭлАдрес
            .Item("П34").Range.Text = КонтТелефон
            .Item("П35").Range.Text = МестоРаб

            If ДолжРуков = "Индивидуальный предприниматель" Then
                .Item("П36").Range.Text = ДолжРуков
                .Item("П37").Range.Text = ""
            Else
                .Item("П36").Range.Text = ДолжРуков & " " & ФормаСобствКор
                .Item("П37").Range.Text = " «" & Клиент & "» "
            End If


            .Item("П38").Range.Text = ФИОКор
            .Item("П39").Range.Text = ПоСовмПриказ
            .Item("П40").Range.Text = mo(0)
        End With
        If Not IO.Directory.Exists(OnePath & Клиент & "\Приказ\" & Год) Then
            IO.Directory.CreateDirectory(OnePath & Клиент & "\Приказ\" & Год)
        End If

        oWordDoc.SaveAs2("C:\Users\Public\Documents\Рик\" & НПриказа & " прием " & Приказ(9) & " от " & Me.MaskedTextBox3.Text & " (приказ)" & " - " & IDso & " .docx",,,,,, False)

        Try
            IO.File.Copy("C:\Users\Public\Documents\Рик\" & НПриказа & " прием " & Приказ(9) & " от " & Me.MaskedTextBox3.Text & " (приказ)" & " - " & IDso & " .docx", OnePath & Клиент & "\Приказ\" & Год & "\" & НПриказа & " прием " & Приказ(9) & " от " & Me.MaskedTextBox3.Text & " (приказ)" & " - " & IDso & " .docx")
        Catch ex As Exception
            'If MessageBox.Show("Приказ № " & НПриказа & " прием " & Приказ(9) & " от " & Me.MaskedTextBox3.Text & " (приказ)" & " - " & IDso & " уже существует. Заменить старый документ новым?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then
            IO.File.Delete(OnePath & Клиент & "\Приказ\" & Год & "\" & НПриказа & " прием " & Приказ(9) & " от " & Me.MaskedTextBox3.Text & " (приказ)" & " - " & IDso & " .docx")
            IO.File.Copy("C:\Users\Public\Documents\Рик\" & НПриказа & " прием " & Приказ(9) & " от " & Me.MaskedTextBox3.Text & " (приказ)" & " - " & IDso & " .docx", OnePath & Клиент & "\Приказ\" & Год & "\" & НПриказа & " прием " & Приказ(9) & " от " & Me.MaskedTextBox3.Text & " (приказ)" & " - " & IDso & " .docx")
            'End If
        End Try
        СохрПрикЛемел = OnePath & Клиент & "\Приказ\" & Год & "\" & НПриказа & " прием " & Приказ(9) & " от " & Me.MaskedTextBox3.Text & " (приказ)" & " - " & IDso & " .docx"


        oWordDoc.Close(True)
        oWord.Quit(True)

        IO.File.Delete("C:\Users\Public\Documents\Рик\Prikaz.docx")
        IO.File.Delete("C:\Users\Public\Documents\Рик\" & НПриказа & " прием " & Приказ(9) & " от " & Me.MaskedTextBox3.Text & " (приказ)" & " - " & IDso & " .docx")


    End Sub
    'Private Sub ПинфудСервис()
    '    inp = ""
    '    ДатРожд = ""
    '    inp = TextBox51.Text
    '    ДатРожд = MaskedTextBox9.Text
    '    Me.Cursor = Cursors.WaitCursor
    '    If CheckBox5.Checked = False Then
    '        ДобавлНовогоСотрудника()
    '        Доки("Пинфуд Сервис")

    '        MessageBox.Show("Сотрудник добавлен!", Рик)
    '        If MessageBox.Show("Контракт № " & txtbx38 & " от " & MaskedTextBox3.Text & vbCrLf & "Приказ № " & НПриказа &
    '          TextBox57.Text & " от " & MaskedTextBox3.Text & vbCrLf & "Заявление от " & MaskedTextBox3.Text & vbCrLf &
    '          "С сотрудником " & vbCrLf & TextBox1.Text & " " & TextBox2.Text & " " & TextBox3.Text & vbCrLf & "Инструкция " & ИнстрП & vbCrLf & "Сформированы!" & vbCrLf & "Распечатать Документы?",
    '          Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then

    '            Task.WaitAll(TskArr)
    '            ПечатьДоковКол(СохрЗак, 1)
    '            ПечатьДоковКол(СохрПрик, 1)
    '            If Not ПровИнстр = 1 Then
    '                ПечатьДоковКол(Инстр, 2)
    '            End If
    '            ПечатьДоковКол(СохрПинфуд, 2)
    '        End If
    '    End If

    '    If CheckBox5.Checked = True And CheckBox7.Checked = False Then

    '        Dim r1 As Task = New Task(AddressOf ОбновлСотрудника)
    '        r1.Start()
    '        MessageBox.Show("Все данные сотрудника " & TextBox6.Text & " " & TextBox5.Text & " " & TextBox4.Text & vbCrLf & " удачно внесены в базу!", Рик, MessageBoxButtons.OK, MessageBoxIcon.None)
    '        If CheckBox23.Checked = True Then
    '            r1.Wait()
    '            Доки("Пинфуд Сервис")

    '            If MessageBox.Show("Контракт № " & txtbx38 & " от " & MaskedTextBox3.Text & vbCrLf & "Приказ № " & НПриказа &
    '              TextBox57.Text & " от " & MaskedTextBox3.Text & vbCrLf & "Заявление от " & MaskedTextBox3.Text & vbCrLf &
    '              "С сотрудником " & vbCrLf & TextBox1.Text & " " & TextBox2.Text & " " & TextBox3.Text & vbCrLf & "Инструкция " & ИнстрП & vbCrLf & "Сформированы!" & vbCrLf & "Распечатать Документы?",
    '              Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then

    '                Task.WaitAll(TskArr)

    '                ПечатьДоковКол(СохрЗак, 1)
    '                ПечатьДоковКол(СохрПрик, 1)
    '                If Not ПровИнстр = 1 Then
    '                    ПечатьДоковКол(Инстр, 2)
    '                End If
    '                ПечатьДоковКол(СохрПинфуд, 2)
    '            End If


    '        End If
    '    End If

    '    Me.Cursor = Cursors.Default

    'End Sub
    'Private Sub ЛемеЛ()
    '    inp = ""
    '    ДатРожд = ""
    '    inp = TextBox51.Text
    '    ДатРожд = MaskedTextBox9.Text
    '    Me.Cursor = Cursors.WaitCursor
    '    If CheckBox5.Checked = False Then
    '        ДобавлНовогоСотрудника()
    '        Доки("ЛемеЛ")

    '        MessageBox.Show("Сотрудник добавлен!", Рик)
    '        If MessageBox.Show("Контракт № " & txtbx38 & " от " & MaskedTextBox3.Text & vbCrLf & "Приказ № " & НПриказа &
    '          TextBox57.Text & " от " & MaskedTextBox3.Text & vbCrLf & "Заявление от " & MaskedTextBox3.Text & vbCrLf &
    '          "С сотрудником " & vbCrLf & TextBox1.Text & " " & TextBox2.Text & " " & TextBox3.Text & vbCrLf & "Инструкция " & ИнстрП & vbCrLf & "Сформированы!" & vbCrLf & "Распечатать Документы?",
    '          Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then

    '            Task.WaitAll(TskArr)
    '            ПечатьДоковКол(СохрЗак, 1)
    '            ПечатьДоковКол(СохрПрикЛемел, 1)
    '            If Not ПровИнстр = 1 Then
    '                ПечатьДоковКол(Инстр, 2)
    '            End If
    '            ПечатьДоковКол(СохрЛемел, 2)
    '        End If
    '    End If

    '    If CheckBox5.Checked = True And CheckBox7.Checked = False Then

    '        Dim r1 As Task = New Task(AddressOf ОбновлСотрудника)
    '        r1.Start()
    '        MessageBox.Show("Все данные сотрудника " & TextBox6.Text & " " & TextBox5.Text & " " & TextBox4.Text & vbCrLf & " удачно внесены в базу!", Рик, MessageBoxButtons.OK, MessageBoxIcon.None)
    '        If CheckBox23.Checked = True Then
    '            r1.Wait()
    '            Доки("ЛемеЛ")

    '            If MessageBox.Show("Контракт № " & txtbx38 & " от " & MaskedTextBox3.Text & vbCrLf & "Приказ № " & НПриказа &
    '              TextBox57.Text & " от " & MaskedTextBox3.Text & vbCrLf & "Заявление от " & MaskedTextBox3.Text & vbCrLf &
    '              "С сотрудником " & vbCrLf & TextBox1.Text & " " & TextBox2.Text & " " & TextBox3.Text & vbCrLf & "Инструкция " & ИнстрП & vbCrLf & "Сформированы!" & vbCrLf & "Распечатать Документы?",
    '              Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then

    '                Task.WaitAll(TskArr)

    '                ПечатьДоковКол(СохрЗак, 1)
    '                ПечатьДоковКол(СохрПрикЛемел, 1)
    '                If Not ПровИнстр = 1 Then
    '                    ПечатьДоковКол(Инстр, 2)
    '                End If
    '                ПечатьДоковКол(СохрЛемел, 2)
    '            End If


    '        End If
    '    End If

    '    Me.Cursor = Cursors.Default

    'End Sub
    Private Sub ДокиПинфуд()
        'Await Task.Delay(0)
        'KillProc()
        'Me.Cursor = Cursors.WaitCursor
        Dim s As New Thread(AddressOf КонтрРазряд) 'поток1
        Dim combx15Th As New Thread(AddressOf Combx15Контракт) 'поток 2
        combx15Th.Start()

        Dim diskU As String = OnePath & "\ОБЩДОКИ\Пинфудсервис\Kontrakt.doc"
        Dim diskC As String = "C:\Users\Public\Documents\Рик\Kontrakt.doc"
        Try 'проверка если есть в С: папке файл Контракт его удаляем и создаем новый

            IO.File.Copy(diskU, diskC)
        Catch ex As Exception
            If ex.Message.Contains("уже существует") Then
                Try
                    IO.File.Delete(diskC)
                    IO.File.Copy(diskU, diskC)
                Catch e As System.IO.IOException
                    If e.Message.Contains("используется другим процессом") Then
                        ПрверкаАсинхрПотоков(Task.CurrentId)
                    End If
                End Try
                'Dim mdoc As Object
                'mdoc = GetObject(, "Word.Application")
                'For Each mdoc In mdoc.Documents
                '    If mdoc.name = "Kontrakt.doc" Then
                '        mdoc.close()
                '    End If
                'Next
                'mdoc.Close(True)
                'mdoc.Quit(True)

                'IO.File.Delete("C:\Users\Public\Documents\Рик\Kontrakt.doc")
                'IO.File.Copy(OnePath & "\ОБЩДОКИ\General\Kontrakt.doc", "C:\Users\Public\Documents\Рик\Kontrakt.doc")
            End If
        End Try
        s.Start()


        Dim oWord2 As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc2 As Microsoft.Office.Interop.Word.Document
        oWord2 = CreateObject("Word.Application")
        oWord2.Visible = False

        oWordDoc2 = oWord2.Documents.Add("C:\Users\Public\Documents\Рик\Kontrakt.doc")
        With oWordDoc2.Bookmarks
            .Item("К0").Range.Text = Контракт(0)
            .Item("К1").Range.Text = mskbx3
            .Item("К2").Range.Text = Заявление(9)
            .Item("К3").Range.Text = Заявление(10)
            .Item("К4").Range.Text = Заявление(11)
            .Item("К5").Range.Text = Заявление(1)
            .Item("К6").Range.Text = Заявление(2)
            .Item("К7").Range.Text = Заявление(3)
            s.Join()
            .Item("К8").Range.Text = ДокКонтрПерем
            'If combx7 = "-" Then
            '    .Item("К8").Range.Text = Strings.LCase(ДолжСОконч)
            'ElseIf combx7 = "1" Or combx7 = "2" Or combx7 = "3" Or combx7 = "4" Or combx7 = "5" Or combx7 = "6" Then
            '    .Item("К8").Range.Text = LCase(ДолжСОконч) & " " & combx7 & " разряда"
            'Else
            '    .Item("К8").Range.Text = Strings.LCase(ДолжСОконч)
            'End If
            .Item("К9").Range.Text = Заявление(8) & " " & СтавкаНов
            .Item("К10").Range.Text = Контракт(2) & " (" & СрКонтПроп & ") " & СклонГод
            .Item("К11").Range.Text = Контракт(1)
            .Item("К12").Range.Text = Контракт(3)
            .Item("К13").Range.Text = Заявление(9)
            .Item("К14").Range.Text = Заявление(10)
            .Item("К15").Range.Text = Заявление(11)
            .Item("К16").Range.Text = Заявление(4)
            .Item("К17").Range.Text = Контракт(5)
            .Item("К18").Range.Text = Контракт(6)
            .Item("К19").Range.Text = Контракт(8)
            .Item("К20").Range.Text = Контракт(7)
            .Item("К21").Range.Text = Контракт(9)
            .Item("К22").Range.Text = Заявление(9)
            .Item("К23").Range.Text = CorName
            .Item("К24").Range.Text = CorOtch
            .Item("К25").Range.Text = Заявление(9) & " " & CorName & CorOtch
            .Item("К26").Range.Text = Контракт(4) & "," & txtbx44
            .Item("К27").Range.Text = Контракт(10)
            .Item("К28").Range.Text = ПоСовмИлиОсн
            'If TextBox46.InvokeRequired Then
            '    Me.Invoke(New txtbx46(AddressOf ДокКонтракт))
            'Else
            .Item("К29").Range.Text = txtbxD46
            'End If
            .Item("К30").Range.Text = РДОрубли
            .Item("К31").Range.Text = РДОкопейки
            .Item("К32").Range.Text = txtbx47
            Select Case combx8
                Case "Руководители"
                    .Item("К38").Range.Text = "должностной инструкции"
                Case "Специалисты"
                    .Item("К38").Range.Text = "должностной инструкции"
            End Select


            combx15Th.Join()
            .Item("К33").Range.Text = К33
            .Item("К34").Range.Text = К34
            .Item("К35").Range.Text = К35
            .Item("К36").Range.Text = К36
            .Item("К37").Range.Text = К37


            .Item("К39").Range.Text = ФормаСобстПолн

            If ФормаСобстПолн = "Индивидуальный предприниматель" Then
                .Item("К40").Range.Text = Клиент
                .Item("К41").Range.Text = ""
            Else
                .Item("К40").Range.Text = " «" & Клиент & "» "
                .Item("К41").Range.Text = ДолжРуковВинПад
            End If

            .Item("К42").Range.Text = ФИОРукРодПад

            If Not combx1 = "Итал Гэлэри Плюс" Then
                .Item("К43").Range.Text = ОснованиеДейств
            Else
                .Item("К51").Range.Text = ""
            End If
            .Item("К44").Range.Text = МестоРаб
            .Item("К45").Range.Text = ФИОКор
            .Item("К46").Range.Text = СборноеРеквПолн
            .Item("К47").Range.Text = Year(Now).ToString
            .Item("К48").Range.Text = TextBox40.Text
            If TextBox56.Text = "" Or TextBox56.Text = "НЕТ" Then
                .Item("К49").Range.Text = ""
            Else
                .Item("К49").Range.Text = "и " & TextBox56.Text & "-го (аванс) "
            End If
            'If ComboBox10.Text = "1.0" Then
            .Item("К50").Range.Text = "1 ставка"
            'Else
            '    .Item("К50").Range.Text = ComboBox10.Text & " ставки"
            'End If
            Select Case combx28
                Case "М"
                    .Item("К52").Range.Text = "ним"
                Case "Ж"
                    .Item("К52").Range.Text = "ней"
            End Select

        End With

        Dim PathContract As String = OnePath & Клиент & "\Контракт\" & Год
        If Not IO.Directory.Exists(PathContract) Then
            IO.Directory.CreateDirectory(PathContract)
        End If

        Dim diskUSave = OnePath & Клиент & "\Контракт\" & Год & "\" & txtbx38 & " " & Заявление(9) & " (контракт)" & " - " & IDso & ".doc"
        Try
            oWordDoc2.SaveAs2(diskUSave,,,,,, False)
        Catch ex As Exception
            If ex.Message.Contains("уже существует") Then
                IO.File.Delete(diskUSave)
                oWordDoc2.SaveAs2(diskUSave,,,,,, False)
            End If
            oWordDoc2.SaveAs2("C:\Users\Public\Documents\Рик\" & txtbx38 & " " & Заявление(9) & " (контракт)" & " - " & IDso & ".doc",,,,,, False)
            IO.File.Copy("C:\Users\Public\Documents\Рик\" & txtbx38 & " " & Заявление(9) & " (контракт)" & " - " & IDso & ".doc", diskUSave)
        End Try
        СохрПинфуд = diskUSave
        oWordDoc2.Close(True)
        oWord2.Quit(True)

        УдалениеСтарыхФайловВПапкеРик("C:\Users\Public\Documents\Рик\" & txtbx38 & " " & Заявление(9) & " (контракт)" & " - " & IDso & ".doc")
        УдалениеСтарыхФайловВПапкеРик("C:\Users\Public\Documents\Рик\Kontrakt.doc")
    End Sub
    Private Async Sub ДокиЛемеЛ()
        Await Task.Delay(0)
        Dim ДолжСОконч, СтавкаНов, СклонГод As String

        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document
        'Dim oWordPara As Microsoft.Office.Interop.Word.Paragraph

        'KillProc()

        oWord = CreateObject("Word.Application")
        oWord.Visible = False


        'ДолжРуковВинПад = ДобОконч(ДолжРуков)
        ДолжСОконч = ДобОконч(Должность)

        СтавкаНов = Склонение(Ставка) 'склонение ставки
        СклонГод = Склонение2(СрокКонтр) ' склонение год
        'СрКонтПроп = ЧислПроп(ComboBox11.Text)
        'ДолжРуковРодПад = ДолжРодПадежФункц(ДолжРуков)
        mo = ДопЛемеЛКонтр() 'испытательный срок и отпуск запрос
        Dim d As String = Replace(Format(CDate(MaskedTextBox3.Text), "dd.MMMM.yyyy"), ".", " ")
        Dim bll As Boolean = txtbx48.Contains(",")
        Dim txt48 As String


        If bll = True Then
            txt48 = Replace(txtbx48, ",", ".")
        Else
            txt48 = txtbx48 & ".00"
        End If

        Try
            IO.File.Copy(OnePath & "\ОБЩДОКИ\Лемел лабс\Kontrakt.doc", "C:\Users\Public\Documents\Рик\Контракт Лемел.doc")
        Catch ex As Exception
            If ex.Message.Contains("уже существует") Then
                Try
                    IO.File.Delete("C:\Users\Public\Documents\Рик\Контракт Лемел.doc")
                    IO.File.Copy(OnePath & "\ОБЩДОКИ\Лемел лабс\Kontrakt.doc", "C:\Users\Public\Documents\Рик\Контракт Лемел.doc")
                Catch e As System.IO.IOException
                    If e.Message.Contains("используется другим процессом") Then

                        ПрверкаАсинхрПотоков(Task.CurrentId)

                    End If
                End Try
            End If
            IO.File.Delete("C:\Users\Public\Documents\Рик\Контракт Лемел.doc")
            IO.File.Copy(OnePath & "\ОБЩДОКИ\Лемел лабс\Kontrakt.doc", "C:\Users\Public\Documents\Рик\Контракт Лемел.doc")
        End Try


        oWordDoc = oWord.Documents.Add("C:\Users\Public\Documents\Рик\Контракт Лемел.doc")

        With oWordDoc.Bookmarks
            .Item("Тк1").Range.Text = Trim(txtbx38)
            .Item("Тк2").Range.Text = d
            If combx28 = "М" Then
                .Item("Тк3").Range.Text = "гражданин"
            Else
                .Item("Тк3").Range.Text = "гражданка"
            End If

            .Item("Тк4").Range.Text = Trim(TextBox1.Text) & " " & Trim(TextBox2.Text) & " " & Trim(TextBox3.Text)
            .Item("Тк5").Range.Text = Strings.LCase(ДолжСОконч)
            If combx10 = "1.0" Then
                .Item("Тк6").Range.Text = "полную ставку"
            Else
                .Item("Тк6").Range.Text = Ставка & " " & СтавкаНов
            End If

            If CheckBox2.Checked = False Then
                .Item("Тк7").Range.Text = "основным местом работы"
            Else
                .Item("Тк7").Range.Text = "работой по совместительству"
            End If
            .Item("Тк8").Range.Text = combx11 & Склонение2(combx11)

            .Item("Тк9").Range.Text = Replace(Format(CDate(Заявление(7)), "dd.MMMM.yyyy"), ".", " ")
            .Item("Тк10").Range.Text = Replace(Format(CDate(Приказ(7)), "dd.MMMM.yyyy"), ".", " ")
            .Item("Тк11").Range.Text = mo(0)
            .Item("Тк12").Range.Text = Strings.LCase(ДолжСОконч)

            Select Case combx15
                Case "ПВТР"
                    .Item("Тк13").Range.Text = "9 часов 00 минут"
                    .Item("Тк14").Range.Text = "18 часов 00 минут"
                    .Item("Тк15").Range.Text = "с 13 часов 00 минут до 14 часов 00 минут"
                    .Item("Тк17").Range.Text = "суббота и воскресенье"
                Case "График"
                    .Item("Тк13").Range.Text = "9 часов 00 минут"
                    .Item("Тк14").Range.Text = "18 часов 00 минут"
                    .Item("Тк15").Range.Text = "с 13 часов 00 минут до 14 часов 00 минут"
                    .Item("Тк17").Range.Text = "суббота и воскресенье"
                Case "Задать"
                    .Item("Тк13").Range.Text = ВремяНач(combx12)
                    .Item("Тк14").Range.Text = ВремяНач(txtbx50)
                    .Item("Тк15").Range.Text = txtbx49
                    .Item("Тк17").Range.Text = "суббота и воскресенье"
            End Select

            .Item("Тк18").Range.Text = mo(1)
            .Item("Тк19").Range.Text = arrtbox("TextBox33") & "." & txtbx44 & " (" & Replace(Replace(arrtbox("TextBox43"), "копеек", "коп."), "бел.рублей", "бел.руб,") & ") "
            .Item("Тк20").Range.Text = txtbxD46
            .Item("Тк21").Range.Text = txt48 & " (" & Replace(arrtbox("TextBox47"), "копеек", "коп.") & ") "
            .Item("Тк22").Range.Text = Trim(arrtbox("TextBox51"))
            .Item("Тк23").Range.Text = Trim(arrtbox("TextBox1")) & " " & Trim(arrtbox("TextBox2")) & " " & Trim(arrtbox("TextBox3"))
            .Item("Тк24").Range.Text = Replace(Format(CDate(arrtmask("MaskedTextBox9")), "dd.MMMM.yyyy"), ".", " ") & " года рождения"
            .Item("Тк25").Range.Text = arrtbox("TextBox12") & " " & arrtbox("TextBox7")
            .Item("Тк26").Range.Text = Replace(Format(CDate(arrtmask("MaskedTextBox1")), "dd.MMMM.yyyy"), ".", " ")
            .Item("Тк27").Range.Text = Trim(arrtbox("TextBox9"))
            .Item("Тк28").Range.Text = Strings.Left(Trim(arrtbox("TextBox2")), 1) & "." & Strings.Left(Trim(arrtbox("TextBox3")), 1) & "." & Trim(arrtbox("TextBox1"))
            .Item("Тк29").Range.Text = d
            .Item("Тк30").Range.Text = d
            .Item("Тк31").Range.Text = d
            .Item("Тк32").Range.Text = d
            .Item("Тк33").Range.Text = Trim(arrtbox("TextBox40"))
            If TextBox56.Text <> "" Then
                .Item("Тк34").Range.Text = ", а за вторую половину месяца " & Trim(arrtbox("TextBox56")) & " числа"
            Else
                .Item("Тк34").Range.Text = ""
            End If

            .Item("Тк35").Range.Text = CType(mo(1), Integer) - 1
            .Item("Тк36").Range.Text = d

        End With

        If Not IO.Directory.Exists(OnePath & Клиент & "\Контракт\" & Год) Then
            IO.Directory.CreateDirectory(OnePath & Клиент & "\Контракт\" & Год)
        End If

        oWordDoc.SaveAs2("C:\Users\Public\Documents\Рик\" & txtbx38 & " " & Заявление(9) & " (контракт)" & ".doc",,,,,, False)
        'oWordDoc.SaveAs2(OnePath & Клиент & "\Контракт\" & Год & "\" & Me.TextBox38.Text & " " & Заявление(9) & " (контракт)" & ".doc",,,,,, False)
        Try
            IO.File.Copy("C:\Users\Public\Documents\Рик\" & txtbx38 & " " & Заявление(9) & " (контракт)" & ".doc", OnePath & Клиент & "\Контракт\" & Now.Year & "\" & txtbx38 & " " & Заявление(9) & " (контракт)" & ".doc")
        Catch ex As Exception
            'If MessageBox.Show("Контракт с сотрудником " & Заявление(9) & " существует. Заменить старый документ новым?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then
            Try
                IO.File.Delete(OnePath & Клиент & "\Контракт\" & Now.Year & "\" & txtbx38 & " " & Заявление(9) & " (контракт)" & ".doc")
            Catch ex1 As Exception
                'KillProc()
                MessageBox.Show("Закройте файл!", Рик)
            End Try

            IO.File.Copy("C:\Users\Public\Documents\Рик\" & txtbx38 & " " & Заявление(9) & " (контракт)" & ".doc", OnePath & Клиент & "\Контракт\" & Now.Year & "\" & txtbx38 & " " & Заявление(9) & " (контракт)" & ".doc")
            'End If
        End Try

        СохрЛемел = OnePath & Клиент & "\Контракт\" & Now.Year & "\" & txtbx38 & " " & Заявление(9) & " (контракт)" & ".doc"

        oWordDoc.Close(True)
        oWord.Quit(True)

        IO.File.Delete("C:\Users\Public\Documents\Рик\Контракт Лемел.doc")
        IO.File.Delete("C:\Users\Public\Documents\Рик\" & txtbx38 & " " & Заявление(9) & " (контракт)" & ".doc")



    End Sub
    Public Async Sub УдаляемФонПроцессы()
        Await Task.Delay(0)

        For Each p As Process In Process.GetProcessesByName("winword")
            p.Kill()
            p.WaitForExit()
        Next

    End Sub
    Private Async Sub Комбы()


        combx18 = ""
        combx1 = ComboBox1.Text
        combx28 = ComboBox28.Text
        combx8 = ComboBox8.Text
        combx9 = ComboBox9.Text
        combx7 = ComboBox7.Text
        combx15 = ComboBox15.Text
        combx10 = ComboBox10.Text
        combx12 = ComboBox12.Text
        combx18 = ComboBox18.Text
        combx19 = ComboBox19.Text
        combx11 = ComboBox11.Text
        cmb28 = ComboBox28.Text
        combx14 = ComboBox14.Text
        combx3 = ComboBox3.Text
        combx13 = ComboBox13.Text
        combx4 = ComboBox4.Text
        combx5 = ComboBox5.Text
        combx6 = ComboBox6.Text
        combx16 = ComboBox16.Text
        txtbxD46 = TextBox46.Text
        txtbx38 = TextBox38.Text
        txtbx46l = TextBox46.TextLength
        txtbx44 = TextBox44.Text
        txtbx48 = TextBox48.Text
        txtbx49 = TextBox49.Text
        txtbx47 = TextBox47.Text
        txtbx50 = TextBox50.Text
        txtbx6 = TextBox6.Text
        mskbx3 = MaskedTextBox3.Text
        txt1 = TextBox1.Text
        txt2 = TextBox2.Text
        txt3 = TextBox3.Text


    End Sub
    Private Sub ЗаполнМассВТабах()

        For Each gp In TabPage1.Controls.OfType(Of GroupBox) 'таб1

            For Each tx In gp.Controls.OfType(Of TextBox)
                arrtbox.Add(tx.Name, tx.Text)
            Next
            For Each tx In gp.Controls.OfType(Of ComboBox)
                arrtcom.Add(tx.Name, tx.Text)
            Next

            For Each ts In gp.Controls.OfType(Of MaskedTextBox)
                arrtmask.Add(ts.Name, ts.Text)
            Next

            For Each tx1 In gp.Controls.OfType(Of GroupBox)

                For Each tx In tx1.Controls.OfType(Of ComboBox)
                    arrtcom.Add(tx.Name, tx.Text)
                Next
                For Each ts In tx1.Controls.OfType(Of MaskedTextBox)
                    arrtmask.Add(ts.Name, ts.Text)
                Next
                For Each tx In tx1.Controls.OfType(Of TextBox)
                    arrtbox.Add(tx.Name, tx.Text)
                Next
            Next

        Next

        If TabControl1.TabPages.ContainsKey("TabPage2") Then 'перебор табов 

            For Each gp In TabPage2.Controls.OfType(Of GroupBox) 'таб2 

                For Each tf In gp.Controls.OfType(Of TextBox)
                    arrtbox.Add(tf.Name, tf.Text)
                Next
                For Each tx In gp.Controls.OfType(Of ComboBox)
                    arrtcom.Add(tx.Name, tx.Text)
                Next
                For Each ts In gp.Controls.OfType(Of MaskedTextBox)
                    arrtmask.Add(ts.Name, ts.Text)
                Next

                For Each tx1 In gp.Controls.OfType(Of GroupBox)

                    For Each tx In tx1.Controls.OfType(Of ComboBox)
                        arrtcom.Add(tx.Name, tx.Text)
                    Next
                    For Each ts In tx1.Controls.OfType(Of MaskedTextBox)
                        arrtmask.Add(ts.Name, ts.Text)
                    Next
                    For Each tf In tx1.Controls.OfType(Of TextBox)
                        arrtbox.Add(tf.Name, tf.Text)
                    Next

                Next

            Next

        ElseIf TabControl1.TabPages.ContainsKey("TabPage3") Then 'перебор табов по Договору подряда

            For Each gp In TabPage3.Controls.OfType(Of GroupBox) 'таб3

                For Each tx In gp.Controls.OfType(Of TextBox)
                    arrtbox.Add(tx.Name, tx.Text)
                Next
                For Each tx In gp.Controls.OfType(Of ComboBox)
                    arrtcom.Add(tx.Name, tx.Text)
                Next
                For Each ts In gp.Controls.OfType(Of MaskedTextBox)
                    arrtmask.Add(ts.Name, ts.Text)
                Next

                For Each tx1 In gp.Controls.OfType(Of GroupBox)

                    For Each tx In tx1.Controls.OfType(Of ComboBox)
                        arrtcom.Add(tx.Name, tx.Text)
                    Next
                    For Each ts In tx1.Controls.OfType(Of MaskedTextBox)
                        arrtmask.Add(ts.Name, ts.Text)
                    Next
                    For Each tf In tx1.Controls.OfType(Of TextBox)
                        arrtbox.Add(tf.Name, tf.Text)
                    Next

                Next
            Next
        End If

    End Sub
    Private Sub ЗаполнМассВнеТабах()

        Dim Ctrl As Control
        Dim Ctrl1 As Control
        Dim Ctrl2 As Control

        For Each Ctrl In Me.Controls 'перебираем текстбоксы вне tabcontrol и groupbox
            If TypeName(Ctrl) = "TextBox" Then
                arrtbox.Add(Ctrl.Name, Ctrl.Text)
                'Ctrl.Value = "бла-бла-бла"
            End If
        Next

        For Each Ctrl1 In Me.Controls 'перебираем combobox вне tabcontrol и groupbox
            If TypeName(Ctrl1) = "ComboBox" Then
                arrtcom.Add(Ctrl1.Name, Ctrl1.Text)
                'Ctrl.Value = "бла-бла-бла"
            End If
        Next

        For Each Ctrl2 In Me.Controls 'перебираем maskedbox вне tabcontrol и groupbox
            If TypeName(Ctrl2) = "MaskedTextBox" Then
                arrtmask.Add(Ctrl2.Name, Ctrl2.Text)
                'Ctrl.Value = "бла-бла-бла"
            End If
        Next

        For Each gh In Me.Controls.OfType(Of GroupBox) 'перебираем combobox вне tabcontrol но в groupbox

            For Each tx In gh.Controls.OfType(Of ComboBox)
                arrtcom.Add(tx.Name, tx.Text)
            Next

            For Each ts In gh.Controls.OfType(Of MaskedTextBox)
                arrtmask.Add(ts.Name, ts.Text)
            Next
            For Each tf In gh.Controls.OfType(Of TextBox)
                arrtbox.Add(tf.Name, tf.Text)
            Next
        Next

    End Sub
    Private Async Sub ОчисткаМаяковAsync()
        Await Task.Run(Sub() ОчисткаМаяков())
    End Sub
    Private Function ПроверкаКонтрактИлиПодрядДобавляем()
        'проверяем заполненность
        If CheckBox5.Checked = False Or ComboBox19.Text = "" Or IsNumeric(Label96.Text) = False Or CheckBox27.Checked = True Then Return 0
        'проверяем стоит ли обрабатывать дальше или это обновление существующего
        Dim f3 = dtKartochkaSotrudnikaAll.Select("IDСотр=" & CType(Label96.Text, Integer) & "")
        Dim dp3 = dtDogovorPadriadaAll.Select("ID=" & CType(Label96.Text, Integer) & "")

        If CheckBox7.Checked = False And f3.Length > 0 Then Return 0
        If CheckBox7.Checked = True And dp3.Length > 0 Then Return 0

        'Проверка (существует ли в комбобоксе сотрудник с дп или к или р(что бы не дублировать.


        Dim dpod As Boolean
        Dim kont As Boolean
        Dim tdog As Boolean
        'If ComboBox19.Items.Contains("(дп)") Then
        '    pl = True
        '    If ComboBox19.Items.Contains(RTrim(Replace(ComboBox19.SelectedItem, "дп", ""))) Or ComboBox19.Items.Contains(RTrim(Replace(ComboBox19.SelectedItem, "дп", "к"))) Then
        '        sl = True
        '    End If
        'End If

        Dim m = Strings.Right(ComboBox19.SelectedItem, 4)
        Select Case m
            Case "(дп)"
                dpod = True
                For Each r In ComboBox19.Items
                    If RTrim(Replace(ComboBox19.Text, "(дп)", "")) = r Or RTrim(Replace(ComboBox19.Text, "(дп)", "(кт)")) = r Then
                        kont = True
                    End If
                    If RTrim(Replace(ComboBox19.Text, "(дп)", "(тд)")) = r Then
                        tdog = True
                    End If
                Next
            Case "(тд)"
                tdog = True
                For Each r In ComboBox19.Items
                    If RTrim(Replace(ComboBox19.Text, "(тд)", "")) = r Or RTrim(Replace(ComboBox19.Text, "(тд)", "(кт)")) = r Then
                        kont = True
                    End If
                    If RTrim(Replace(ComboBox19.Text, "(тд)", "(дп)")) = r Then
                        dpod = True
                    End If
                Next
            Case Else
                kont = True
                Dim mp As String
                Dim mp1 As String
                If ComboBox19.Text.Contains("(кт)") Then
                    mp = RTrim(Replace(ComboBox19.Text, "(кт)", "(тд)"))
                    mp1 = RTrim(Replace(ComboBox19.Text, "(кт)", "(дп)"))
                Else
                    mp = ComboBox19.Text & "(тд)"
                    mp1 = ComboBox19.Text & "(дп)"
                End If

                For Each r In ComboBox19.Items

                    If mp1 = r Then
                        dpod = True
                    End If

                    If mp = r Then
                        tdog = True
                    End If

                Next
        End Select

        'ищем по фио контракт или труд дог


        'ищем по фио договор подряда или труд дог


        If CheckBox5.Checked = True And ComboBox19.Text <> "" And IsNumeric(Label96.Text) And CheckBox7.Checked = False Then
            'Dim f = dtKartochkaSotrudnikaAll.Select("IDСотр=" & CType(Label96.Text, Integer) & "")
            'Dim dp = dtDogovorPadriadaAll.Select("ID=" & CType(Label96.Text, Integer) & "")

            If kont = True And dpod = True Then
                MessageBox.Show("У данного сотрудника уже заключен контракт!", Рик)
                Return 1
            End If

            If dp3.Length > 0 Then
                If MessageBox.Show("Создать контракт?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                    Решение = "Контракт"
                End If
            End If

        End If



        If CheckBox5.Checked = True And ComboBox19.Text <> "" And IsNumeric(Label96.Text) And CheckBox7.Checked = True Then
            'Dim f = dtKartochkaSotrudnikaAll.Select("IDСотр=" & CType(Label96.Text, Integer) & "")
            'Dim dp = dtDogovorPadriadaAll.Select("ID=" & CType(Label96.Text, Integer) & "")

            If kont = True And dpod = True Then
                MessageBox.Show("У данного сотрудника уже заключен договор-подряда!", Рик)
                Return 1
            End If


            If f3.Length > 0 Then



                If MessageBox.Show("Создать договор-подряда?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then

                    Решение = "Подряд"
                End If
            End If



        End If


        Return 0


    End Function




    Private Sub ОчисткаМаяков()
        ПровВходаCom8 = False
        ПровВходаCom19 = False
    End Sub
    Private Sub НовыйПутьДПЦена()

        'Проверяем ДПЦена

        If ПроверкаЗаполненностиВкладкиКонтрактНовыйПуть() = 1 Then
            Exit Sub
        End If
        СохраняемКонтролыВСписки(TabPage3)








    End Sub
    Private Sub НовыйПутьДПИное()

        'Проверяем ДПИное

        If ПроверкаЗаполненностиВкладкиКонтрактНовыйПуть() = 1 Then
            Exit Sub
        End If
        СохраняемКонтролыВСписки(TabPage4)






    End Sub
    Private Function ПроверкаЗаполненностиВкладкиКонтрактНовыйПуть()
        If TextBox56.Text = "" Then
            If MessageBox.Show("Вы не выбрали дату выплаты аванса! Выбрать?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Information) = DialogResult.Yes Then
                Return 1
            End If
        End If
        If TextBox40.Text = "" Then
            MessageBox.Show("Выберите дату выплаты зарплаты!", Рик)
            Return 1
        End If


        If Not Примечани <> "" Then
            If MessageBox.Show("Вы НЕ заполнили примечание!" & vbCrLf & "Выберите OK - если хотите продолжить, или ОТМЕНА - если хотите изменить", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.Cancel Then
                Return 1
            End If

        End If



        If Not ComboBox18.Text <> "" Then
            MessageBox.Show("Выберите объект общепита!", Рик)
            Return 1
        End If

        If CheckBox26.Checked = True Then
            If Not ComboBox8.Text <> "" Then
                MessageBox.Show("Выберите отдел!", Рик)
                Return 1
            End If
            If Not ComboBox9.Text <> "" Then
                MessageBox.Show("Выберите должность!", Рик)
                Return 1
            End If

            If ComboBox7.Items.Count = 1 Then
                If ComboBox7.Enabled = True And Not ComboBox7.Text <> "" Then
                    MessageBox.Show("Выберите разряд!", Рик)
                    Return 1
                End If
            End If


        End If

        If Not ComboBox10.Text <> "" Then
            MessageBox.Show("Выберите ставку!", Рик)
            Return 1
        End If

        If TextBox57.Text <> "" Then
            TextBox57.Text = " - " & TextBox57.Text
        End If

        'проверка номера приказа на число
        If IsNumeric(TextBox41.Text) Then
            НПриказа = TextBox41.Text & " - " & TextBox58.Text & TextBox57.Text
        Else
            НПриказа = TextBox41.Text
        End If

        If ComboBox8.Text = "" Or ComboBox9.Text = "" Or ComboBox10.Text = "" Or TextBox33.Text = "" Then
            MessageBox.Show("Заполните раздел подразделение!", Рик)
            Return 1
        End If
        If MaskedTextBox3.Text = "" Or TextBox38.Text = "" Or TextBox41.Text = "" Or MaskedTextBox4.Text = "" Or MaskedTextBox5.Text = "" Or ComboBox15.Text = "" Or ComboBox11.Text = "" Then
            MessageBox.Show("Заполните раздел контракт и приказ!", Рик)
            Return 1
        End If


        Dim от1, дол1, разр1 As String

        'If CheckBox5.Checked = True And CheckBox26.Checked = False Then
        '    Dim strsql85 As String = "SELECT Отдел,Должность,Разряд FROM Штатное WHERE ИДСотр=" & CType(Label96.Text, Integer) & ""
        '    Dim hk As DataTable = Selects(strsql85)
        '    от1 = hk.Rows(0).Item(0).ToString
        '    дол1 = hk.Rows(0).Item(1).ToString
        '    разр1 = hk.Rows(0).Item(2).ToString
        'Else
        от1 = ComboBox8.Text
        дол1 = ComboBox9.Text
        разр1 = ComboBox7.Text
        'End If

        ПровИнстр = Nothing

        'выбираем инструкцию, если есть.
        Dim ds
        Using dbcx As New DbAll1DataContext
            ds = (From x In dbcx.ШтОтделы.AsEnumerable
                  Join y In dbcx.ШтСвод.AsEnumerable On x.Код Equals y.Отдел
                  Where x.Клиент = ComboBox1.Text And x.Отделы = от1 And y.Должность = дол1 And y.Разряд = разр1
                  Select y.ДолжИнструкция).FirstOrDefault()
        End Using


        '        Dim ds = Selects(StrSql:="SELECT ШтСвод.ДолжИнструкция FROM ШтОтделы INNER JOIN ШтСвод ON ШтОтделы.Код = ШтСвод.Отдел
        'WHERE ШтОтделы.Клиент='" & ComboBox1.Text & "' AND ШтОтделы.Отделы='" & от1 & "' AND ШтСвод.Должность='" & дол1 & "'
        'AND ШтСвод.Разряд='" & разр1 & "'")
        Try
            If ds = "False" Then
                If MessageBox.Show("Для данной должности не сформирована должностная инструкция!" & vbCrLf & "Оформить инструкцию?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                    ПровИнстр = 1
                    Return 0
                Else
                    v = False
                    ДолжИнстр.ShowDialog()
                    If ДолжИнстр.текст = "" Or ДолжИнстр.Ном = "" Then
                        If MessageBox.Show("Вы не заполнили номер или текст инструкции!" & vbCrLf & "Все равно продолжить?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.Cancel Then
                            v = False
                            ДолжИнстр.ShowDialog()
                        End If
                    End If

                    If v = False Then
                        ПровИнстр = 1
                        Return 0
                    Else
                        ДокиИнструкцияНовыйПуть()
                        ПровИнстр = Nothing
                        Return 0
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            MessageBox.Show("Проверьте разряд сотрудника, он не совпадает с данными в штатном расписании!", Рик)
            Return 1
        End Try

        Return 0
    End Function
    Private Sub ДокПредварДаннНовыйПуть()
        Try
            IDso = CType(Label96.Text, Integer)
        Catch ex As Exception
            IDso = IDsot1
        End Try

        ДолжРуковВинПад = ДобОконч(ДолжРуков)  'ok

        'если изменяем сотрудника и поле должность пустое то подтягиваем должность из базы
        If Должность = "" And CheckBox5.Checked = True And CheckBox23.Checked = True Then
            ДолжСОконч = ДобОконч(ДолжПриИзменСотр())
        Else
            ДолжСОконч = ДобОконч(Replace(Должность, ".", ""))
        End If

        СтавкаНов = Склонение(arrtcom("ComboBox10")) 'СтавкаНов = Склонение(Ставка) 'склонение ставки   'ok
        СклонГод = Склонение2(arrtcom("ComboBox11")) 'СклонГод = Склонение2(СрокКонтр) ' склонение год  'ok
        СрКонтПроп = ЧислПроп(arrtcom("ComboBox11")) 'СрКонтПроп = ЧислПроп(ComboBox11.Text) 'ok

        If CheckBox2.Checked = True Then 'галочка по осн или по совместительству
            ПоСовмИлиОсн = "совместительству"
            ПоСовмПриказ = "по совместительству"
        Else
            ПоСовмИлиОсн = "основной работе"
            ПоСовмПриказ = "основное место работы"
        End If
        ДолжРуковРодПад = ДолжРодПадежФункц(ДолжРуков)  'ok
    End Sub
    Private Function ДолжнИразрядДокЗаявлениеновыйПуть(ByVal d As Integer)
        Dim s As String = ""
        Using dbcx As New DbAll1DataContext
            Dim var = (From x In dbcx.Штатное.AsEnumerable
                       Where x.ИДСотр = d
                       Select x).FirstOrDefault()

            If IsNumeric(var.Разряд) Then
                s = " " & var.Разряд & " разряда"
            End If
        End Using

        Return s


    End Function
    Private Sub ДокЗаявлениеНовыйПуть(ByVal Сотрудник As IEnumerable(Of Сотрудники))

        Dim oWord1 As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc1 As Microsoft.Office.Interop.Word.Document
        oWord1 = CreateObject("Word.Application")
        oWord1.Visible = False

        ВыгрузкаФайловНаЛокалыныйКомп(FTPStringAllDOC & "Zayavlenie.doc", firthtPath & "\Zayavlenie.doc")
        oWordDoc1 = oWord1.Documents.Add(firthtPath & "\Zayavlenie.doc")


        With oWordDoc1.Bookmarks
            .Item("ЗАКЛ0").Range.Text = arrtmask("MaskedTextBox3")
            .Item("ЗАКЛ1").Range.Text = Сотрудник(0).ФамилияДляЗаявления 'Trim(Приказ(2))
            .Item("ЗАКЛ2").Range.Text = Сотрудник(0).ИмяДляЗаявления 'Trim(Заявление(13))
            .Item("ЗАКЛ3").Range.Text = Сотрудник(0).ОтчествоДляЗаявления 'Trim(Заявление(14))
            .Item("ЗАКЛ4").Range.Text = Сотрудник(0).Регистрация 'Заявление(4)
            .Item("ЗАКЛ5").Range.Text = Сотрудник(0).КонтТелефон 'Заявление(5)
            .Item("ЗАКЛ6").Range.Text = LCase(ДолжСОконч) & ДолжнИразрядДокЗаявлениеновыйПуть(Сотрудник(0).КодСотрудники)
            .Item("ЗАКЛ7").Range.Text = arrtmask("MaskedTextBox4") 'Заявление(7)MaskedTextBox4.Text
            .Item("ЗАКЛ8").Range.Text = arrtcom("ComboBox10") 'Заявление(8)
            .Item("ЗАКЛ9").Range.Text = Mid(Сотрудник(0).Имя, 1, 1) & "." 'CorName
            .Item("ЗАКЛ10").Range.Text = Mid(Сотрудник(0).Отчество, 1, 1) & "." 'CorOtch
            .Item("ЗАКЛ11").Range.Text = Сотрудник(0).Фамилия 'Заявление(9)
            .Item("ЗАКЛ12").Range.Text = СтавкаНов 'OK
            .Item("ЗАКЛ13").Range.Text = ДолжРуковРодПад  'OK

            If ДолжРуковРодПад = "Индивидуальному предпринимателю" Or ФормаСобствКор = "ИП" Then 'OK
                .Item("ЗАКЛ14").Range.Text = ""
                .Item("ЗАКЛ18").Range.Text = "по месту нахождения"
            Else
                .Item("ЗАКЛ14").Range.Text = ФормаСобствКор & " """ & Клиент & """ "
                .Item("ЗАКЛ18").Range.Text = "в"
            End If

            .Item("ЗАКЛ15").Range.Text = ФИОКорРукДат 'OK
            .Item("ЗАКЛ16").Range.Text = МестоРаб 'OK
            .Item("ЗАКЛ17").Range.Text = arrtmask("MaskedTextBox3")


        End With


        Dim dirstring As String = Клиент & "/Заявление/" & Now.Year & "/" 'место сохранения файла

        dirstring = СозданиепапкиНаСервере(dirstring) 'полный путь на сервер(кроме имени и разрешения файла)


        Dim put, Name As String
        Name = Сотрудник(0).Фамилия & " (заявление)" & " - " & IDso & ".doc"
        put = PathVremyanka & Name 'место в корне программы


        ВыборкаИзагрНаСервер(dirstring, Name, "Прием-Зявление")


        oWordDoc1.SaveAs2(put,,,,,, False)
        oWordDoc1.Close(True)
        СохрЗакFTP.AddRange(New String() {dirstring, Name})
        oWord1.Quit(True)
        dirstring += Name
        ЗагрНаСерверИУдаление(put, dirstring, put)



        ВременнаяПапкаУдалениеФайла(firthtPath & "\Zayavlenie.doc")
    End Sub
    Private Sub ГрафикСортНовыйПуть()
        К33 = ""
        К34 = ""
        К35 = ""
        К36 = ""
        К37 = ""


        Select Case arrtcom("ComboBox15")
            Case "График"
                К33 = "согласно графику работ"
                К34 = "согласно графику работ"
                К35 = "согласно графику работ"
                Select Case CheckBox4.Checked
                    Case False
                        К36 = "Суббота, Воскресенье"
                    Case True
                        К36 = "согласно графику работ"
                        К37 = "11.5. работнику устанавливается суммированный учет рабочего времени с учетным периодом - год."
                End Select

            Case "ПВТР"
                К33 = "согласно правил внутреннего трудового распорядка"
                К34 = "согласно правил внутреннего трудового распорядка"
                К35 = "согласно правил внутреннего трудового распорядка"

                Select Case CheckBox4.Checked
                    Case False
                        К36 = "согласно графику работ"
                    Case True
                        К36 = "согласно графику работ"
                        К37 = "11.5. работнику устанавливается суммированный учет рабочего времени с учетным периодом - год."
                End Select

            Case "Задать"
                К33 = arrtcom("ComboBox12") 'combx12
                К34 = arrtbox("TextBox49") 'TextBox49.Text
                К35 = arrtbox("TextBox50") 'TextBox50.Text

                Select Case CheckBox4.Checked
                    Case False
                        К36 = "Суббота, Воскресенье"
                    Case True
                        К36 = "согласно графику работ"
                        К37 = "11.5. работнику устанавливается суммированный учет рабочего времени с учетным периодом - год."
                End Select
        End Select
    End Sub
    Private Sub ДКонтрактНовыйПуть(ByVal Сотрудник As IEnumerable(Of Сотрудники))

        Dim oWord2 As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc2 As Microsoft.Office.Interop.Word.Document
        oWord2 = CreateObject("Word.Application")
        oWord2.Visible = False

        ВыгрузкаФайловНаЛокалыныйКомп(FTPStringAllDOC & "Kontrakt.doc", firthtPath & "\Kontrakt.doc")

        oWordDoc2 = oWord2.Documents.Add(firthtPath & "\Kontrakt.doc")



        With oWordDoc2.Bookmarks
            .Item("К0").Range.Text = arrtbox("TextBox38") 'Контракт(0) 'ok
            .Item("К1").Range.Text = arrtmask("MaskedTextBox3") 'mskbx3 'ok
            .Item("К2").Range.Text = Сотрудник(0).Фамилия 'Заявление(9) 'ok
            .Item("К3").Range.Text = Сотрудник(0).Имя 'Заявление(10) 'ok
            .Item("К4").Range.Text = Сотрудник(0).Отчество 'Заявление(11) 'ok
            .Item("К5").Range.Text = Сотрудник(0).ФамилияРодПад 'Заявление(1)  'ok
            .Item("К6").Range.Text = Сотрудник(0).ИмяРодПад 'Заявление(2) 'ok
            .Item("К7").Range.Text = Сотрудник(0).ОтчествоРодПад 'Заявление(3) 'ok

            'Await Task.Run(Sub() КонтрРазряд()) 'ok

            .Item("К8").Range.Text = LCase(ДолжСОконч) & ДолжнИразрядДокЗаявлениеновыйПуть(Сотрудник(0).КодСотрудники) ' ДокКонтрПерем 'ok

            .Item("К9").Range.Text = arrtcom("ComboBox10") & " " & СтавкаНов 'Заявление(8) 'ok
            .Item("К10").Range.Text = arrtcom("ComboBox11") & " (" & СрКонтПроп & ") " & СклонГод 'Контракт(2) & " (" & СрКонтПроп & ") " & СклонГод  'ok
            .Item("К11").Range.Text = arrtmask("MaskedTextBox4") 'Контракт(1) 'ok
            .Item("К12").Range.Text = arrtmask("MaskedTextBox5") 'Контракт(3) 'ok
            .Item("К13").Range.Text = Сотрудник(0).Фамилия 'Заявление(9)'ok
            .Item("К14").Range.Text = Сотрудник(0).Имя 'Заявление(10) 'ok
            .Item("К15").Range.Text = Сотрудник(0).Отчество 'Заявление(11)  'ok
            .Item("К16").Range.Text = Сотрудник(0).Регистрация 'Заявление(4)  'ok
            .Item("К17").Range.Text = Сотрудник(0).ПаспортСерия 'Контракт(5) 'ok
            .Item("К18").Range.Text = Сотрудник(0).ПаспортНомер 'Контракт(6) 'ok
            .Item("К19").Range.Text = Сотрудник(0).ПаспортКемВыдан 'Контракт(8) 'ok
            .Item("К20").Range.Text = Сотрудник(0).ПаспортКогдаВыдан 'Контракт(7) 'ok
            .Item("К21").Range.Text = Сотрудник(0).ИДНомер 'Контракт(9) 'ok
            .Item("К22").Range.Text = Сотрудник(0).Фамилия 'Заявление(9)'ok
            .Item("К23").Range.Text = Mid(Сотрудник(0).Имя, 1, 1) & "." 'CorName 'ok
            .Item("К24").Range.Text = Mid(Сотрудник(0).Отчество, 1, 1) & "." 'CorOtch 'ok
            .Item("К25").Range.Text = Сотрудник(0).Фамилия & " " & Mid(Сотрудник(0).Имя, 1, 1) & "." & Mid(Сотрудник(0).Отчество, 1, 1) & "." 'Заявление(9) & " " & CorName & CorOtch 'ok
            .Item("К26").Range.Text = arrtbox("TextBox33") & "," & arrtbox("TextBox44") 'Контракт(4) & "," & txtbx44  'ok
            .Item("К27").Range.Text = arrtbox("TextBox43") 'Контракт(10) 'ok
            .Item("К28").Range.Text = ПоСовмИлиОсн 'ok
            'If TextBox46.InvokeRequired Then
            '    Me.Invoke(New txtbx46(AddressOf ДокКонтракт))
            'Else
            .Item("К29").Range.Text = arrtbox("TextBox46") 'txtbxD46 'ok
            'End If
            .Item("К30").Range.Text = РДОрубли 'ok
            .Item("К31").Range.Text = РДОкопейки  'ok
            .Item("К32").Range.Text = arrtbox("TextBox47") 'txtbx47 'ok
            Select Case arrtcom("ComboBox8")'combx8 'ok
                Case "Руководители"
                    .Item("К38").Range.Text = "должностной инструкции"
                Case "Специалисты"
                    .Item("К38").Range.Text = "должностной инструкции"
            End Select

            ГрафикСортНовыйПуть() 'Await Task.Run(Sub() Combx15Контракт()) 'ok

            .Item("К33").Range.Text = К33 'ok
            .Item("К34").Range.Text = К34  'ok
            .Item("К35").Range.Text = К35  'ok
            .Item("К36").Range.Text = К36  'ok
            .Item("К37").Range.Text = К37  'ok


            .Item("К39").Range.Text = ФормаСобстПолн 'ok

            If ФормаСобстПолн = "Индивидуальный предприниматель" Then
                .Item("К40").Range.Text = Клиент  'ok
                .Item("К41").Range.Text = ""  'ok
            Else
                .Item("К40").Range.Text = " «" & Клиент & "» "  'ok
                .Item("К41").Range.Text = ДолжРуковВинПад 'ok
            End If

            .Item("К42").Range.Text = ФИОРукРодПад 'ok

            If Not arrtcom("ComboBox1") = "Итал Гэлэри Плюс" Then 'combx1  'ok
                .Item("К43").Range.Text = ОснованиеДейств 'ok
            Else
                .Item("К51").Range.Text = ""  'ok
            End If
            .Item("К44").Range.Text = МестоРаб 'ok
            .Item("К45").Range.Text = ФИОКор  'ok
            .Item("К46").Range.Text = СборноеРеквПолн 'ok
            .Item("К47").Range.Text = Year(Now).ToString  'ok
            .Item("К48").Range.Text = arrtbox("TextBox40") 'TextBox40.Text 'ok

            If arrtbox("TextBox56") = "" Or arrtbox("TextBox56") = "НЕТ" Then  'If TextBox56.Text = "" Or TextBox56.Text = "НЕТ" Then 'ok
                .Item("К49").Range.Text = "" 'ok
            Else
                .Item("К49").Range.Text = "и " & arrtbox("TextBox56") & "-го (аванс) "  'ok
            End If
            'If ComboBox10.Text = "1.0" Then
            .Item("К50").Range.Text = "1 ставка"
            'Else
            '    .Item("К50").Range.Text = ComboBox10.Text & " ставки"
            'End If
            Select Case Сотрудник(0).Пол'combx28
                Case "М"
                    .Item("К52").Range.Text = "ним"
                Case "Ж"
                    .Item("К52").Range.Text = "ней"
            End Select

        End With


        Dim dirstring As String = Клиент & "/Контракт/" & Now.Year & "/" 'место сохранения файла
        dirstring = СозданиепапкиНаСервере(dirstring) 'полный путь на сервер(кроме имени и разрешения файла)


        Dim put, Name As String
        Name = arrtbox("TextBox38") & " " & Сотрудник(0).Фамилия & " (контракт)" & " - " & Сотрудник(0).КодСотрудники & ".doc" 'txtbx38 & " " & Заявление(9) & " (контракт)" & " - " & IDso & ".doc"
        put = PathVremyanka & Name 'место в корне программы

        ВыборкаИзагрНаСервер(dirstring, Name, "Прием-Контракт")

        oWordDoc2.SaveAs2(put,,,,,, False)


        oWordDoc2.Close(True)
        oWord2.Quit(True)
        СохрКонтрFTP.AddRange(New String() {dirstring, Name})
        dirstring += Name

        ЗагрНаСерверИУдаление(put, dirstring, put)


        ВременнаяПапкаУдалениеФайла(firthtPath & "\Kontrakt.doc")



    End Sub
    Private Sub ДокПриказНовыйПуть(ByVal Сотрудник As IEnumerable(Of Сотрудники))
        Dim oWord3 As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc3 As Microsoft.Office.Interop.Word.Document
        oWord3 = CreateObject("Word.Application")
        oWord3.Visible = False

        ВыгрузкаФайловНаЛокалыныйКомп(FTPStringAllDOC & "Prikaz.doc", firthtPath & "\Prikaz.doc")

        oWordDoc3 = oWord3.Documents.Add(firthtPath & "\Prikaz.doc")


        With oWordDoc3.Bookmarks
            .Item("П1").Range.Text = arrtmask("MaskedTextBox3") 'Приказ(5) 'ok
            .Item("П2").Range.Text = НПриказа 'ok
            .Item("П3").Range.Text = Сотрудник(0).ФамилияРодПад 'txtbx6  'ok
            .Item("П4").Range.Text = Mid(Сотрудник(0).Имя, 1, 1) & "." 'CorName  'ok
            .Item("П5").Range.Text = Mid(Сотрудник(0).Отчество, 1, 1) & "." 'CorOtch  'ok
            .Item("П6").Range.Text = Сотрудник(0).ФамилияРодПад 'txtbx6 'ok
            .Item("П7").Range.Text = Сотрудник(0).ИмяРодПад 'Приказ(3) 'ok
            .Item("П8").Range.Text = Сотрудник(0).ОтчествоРодПад 'Приказ(4)  'ok
            .Item("П9").Range.Text = LCase(ДолжСОконч) & ДолжнИразрядДокЗаявлениеновыйПуть(Сотрудник(0).КодСотрудники) 'Strings.LCase(ДолжСОконч) & ДолжнИразрядДокЗаявление()  'ok
            .Item("П10").Range.Text = arrtmask("MaskedTextBox4") 'Приказ(6)'ok
            .Item("П11").Range.Text = arrtcom("ComboBox10") 'Ставка  'ok
            .Item("П12").Range.Text = СтавкаНов  'ok
            .Item("П13").Range.Text = arrtcom("ComboBox11") 'СрокКонтр 'ok
            .Item("П14").Range.Text = СклонГод 'ok
            .Item("П15").Range.Text = arrtmask("MaskedTextBox4") 'Приказ(6)'ok
            .Item("П16").Range.Text = arrtmask("MaskedTextBox5") 'Приказ(7)
            .Item("П17").Range.Text = Сотрудник(0).ФамилияДляЗаявления 'Приказ(2)
            .Item("П18").Range.Text = Mid(Сотрудник(0).Имя, 1, 1) & "." 'CorName  'ok
            .Item("П19").Range.Text = Mid(Сотрудник(0).Отчество, 1, 1) & "." 'CorOtch  'ok
            .Item("П20").Range.Text = arrtbox("TextBox38") 'Приказ(8) 'ok
            .Item("П21").Range.Text = arrtmask("MaskedTextBox3") 'Приказ(5) 'ok
            .Item("П22").Range.Text = Сотрудник(0).Фамилия 'Приказ(9)  'ok
            .Item("П23").Range.Text = Mid(Сотрудник(0).Имя, 1, 1) & "." 'CorName  'ok
            .Item("П24").Range.Text = Mid(Сотрудник(0).Отчество, 1, 1) & "." 'CorOtch  'ok
            .Item("П25").Range.Text = ФормаСобстПолн 'ok

            If ФормаСобстПолн = "Индивидуальный предприниматель" Then
                .Item("П26").Range.Text = Клиент   'ok
            Else
                .Item("П26").Range.Text = " «" & Клиент & "» "  'ok
            End If

            .Item("П27").Range.Text = ЮрАдрес   'ok
            .Item("П28").Range.Text = УНП  'ok
            .Item("П29").Range.Text = РасСчет  'ok
            .Item("П30").Range.Text = АдресБанка 'ok
            .Item("П31").Range.Text = БИК 'ok
            .Item("П33").Range.Text = ЭлАдрес 'ok
            .Item("П34").Range.Text = КонтТелефон  'ok
            .Item("П35").Range.Text = МестоРаб   'ok

            If ДолжРуков = "Индивидуальный предприниматель" Then    'ok
                .Item("П36").Range.Text = ДолжРуков   'ok
                .Item("П37").Range.Text = ""   'ok
            Else
                .Item("П36").Range.Text = ДолжРуков & " " & ФормаСобствКор  'ok
                .Item("П37").Range.Text = " «" & Клиент & "» "  'ok
            End If


            .Item("П38").Range.Text = ФИОКор  'ok
            .Item("П39").Range.Text = ПоСовмПриказ  'ok

        End With


        Dim dirstring As String = Клиент & "/Приказ/" & Now.Year & "/" 'место сохранения файла
        dirstring = СозданиепапкиНаСервере(dirstring) 'полный путь на сервер(кроме имени и разрешения файла)


        Dim put, Name As String
        Name = НПриказа & " прием " & Сотрудник(0).Фамилия & " от " & arrtmask("MaskedTextBox3") & " (приказ)" & " - " & Сотрудник(0).КодСотрудники & " .doc" 'НПриказа & " прием " & Приказ(9) & " от " & mskbx3 & " (приказ)" & " - " & IDso & " .doc"
        put = PathVremyanka & Name 'место в корне программы

        ВыборкаИзагрНаСервер(dirstring, Name, "Прием-Приказ")

        'Dim b = dtSotrudnikiAll.Select("ФИОСборное='" & combx19 & "'") 'выбираем данные по сотруднику
        'Dim kd As Integer = CType(b(0).Item("КодСотрудники").ToString, Integer) 'находим ИД сотрудника
        'ЗагрВБазуПутиДоковAsync(kd, dirstring, Name, "Прием-Приказ") 'заполняем данные путей и назв файла

        oWordDoc3.SaveAs2(put,,,,,, False)


        oWordDoc3.Close(True)
        oWord3.Quit(True)

        СохрПрикFTP.AddRange(New String() {dirstring, Name})
        dirstring += Name

        ЗагрНаСерверИУдаление(put, dirstring, put)

        ВременнаяПапкаУдалениеФайла(firthtPath & "\Prikaz.doc")
    End Sub
    Private Sub ДокИнструкцНовыйПуть(Сотрудник As IEnumerable(Of Сотрудники))


        Dim hk As DataTable
        Dim list As New Dictionary(Of String, Object)()        '
        list.Add("@Клиент", arrtcom("ComboBox1"))
        list.Add("@Отделы", arrtcom("ComboBox8"))
        list.Add("@Должность", arrtcom("ComboBox9"))
        list.Add("@Разряд", arrtcom("ComboBox7"))

        'list.Add("@Клиент", combx1)
        'list.Add("@Отделы", combx8)
        'list.Add("@Должность", combx9)
        'list.Add("@Разряд", combx7)



        list.Add("@ID", Сотрудник(0).КодСотрудники)

        ИнстрFTP.Clear()

        Try
            If Not ПровИнстр = 1 Then
                'Формируем инструкцию 
                If CheckBox23.Checked = False Then
                    Dim dg = Selects(StrSql:="Select ШтСвод.НомерДолжИнстр FROM ШтОтделы INNER JOIN ШтСвод On ШтОтделы.Код = ШтСвод.Отдел
        WHERE ШтОтделы.Клиент=@Клиент AND ШтОтделы.Отделы=@Отделы AND ШтСвод.Должность=@Должность AND ШтСвод.Разряд=@Разряд AND ШтСвод.ДолжИнструкция='True'", list)
                    If errds = 0 Then
                        ИнстрП = dg.Rows(0).Item(0).ToString
                        ИнстрFTP.AddRange(New String() {FTPString & arrtcom("ComboBox1") & "/Должностные инструкции/", dg.Rows(0).Item(0).ToString & ".doc"})
                    End If

                Else
                    If CheckBox26.Checked = True Then
                        Dim dg = Selects(StrSql:="SELECT ШтСвод.НомерДолжИнстр FROM ШтОтделы INNER JOIN ШтСвод ON ШтОтделы.Код = ШтСвод.Отдел
WHERE ШтОтделы.Клиент=@Клиент AND ШтОтделы.Отделы=@Отделы AND ШтСвод.Должность=@Должность
AND ШтСвод.Разряд=@Разряд AND ШтСвод.ДолжИнструкция='True'", list)

                        If errds = 0 Then
                            ИнстрП = dg.Rows(0).Item(0).ToString
                            'Инстр = OnePath & combx1 & "\Должностные инструкции\" & dg.Rows(0).Item(0).ToString & ".doc"
                            ИнстрFTP.AddRange(New String() {FTPString & arrtcom("ComboBox1") & "/Должностные инструкции/", dg.Rows(0).Item(0).ToString & ".doc"})
                        End If

                    Else
                        If Not hk Is Nothing Then hk.Clear()
                        hk = Selects(StrSql:="SELECT Отдел,Должность,Разряд FROM Штатное WHERE ИДСотр=@ID", list)

                        If Not hk Is Nothing Then
                            list.Add("@Отделы2", hk.Rows(0).Item(0).ToString)
                            list.Add("@Должность2", hk.Rows(0).Item(1).ToString)
                            list.Add("@Разряд2", hk.Rows(0).Item(2).ToString)
                        End If

                        If errds = 0 Then
                            Dim dg = Selects(StrSql:="SELECT ШтСвод.НомерДолжИнстр FROM ШтОтделы INNER JOIN ШтСвод ON ШтОтделы.Код = ШтСвод.Отдел
WHERE ШтОтделы.Клиент=@Клиент AND ШтОтделы.Отделы=@Отделы2 AND ШтСвод.Должность=@Должность2
AND ШтСвод.Разряд=@Разряд2 AND ШтСвод.ДолжИнструкция='True'", list)

                            Try
                                ИнстрП = dg.Rows(0).Item(0).ToString
                                ИнстрFTP.AddRange(New String() {FTPString & arrtcom("ComboBox1") & "/Должностные инструкции/", dg.Rows(0).Item(0).ToString & ".doc"})
                            Catch ex As Exception

                            End Try

                            'Инстр = OnePath & combx1 & "\Должностные инструкции\" & dg.Rows(0).Item(0).ToString & ".doc"

                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Sub
    Private Sub ДокиКонтрактНовыйПуть()

        'выбираем данные из таблицы Сотрудники по idсотрудника
        Dim Сотрудник As IEnumerable(Of Сотрудники)
        Using dbcx As New DbAll1DataContext
            Сотрудник = (From x In dbcx.Сотрудники.AsEnumerable
                         Where x.КодСотрудники = CType(Label96.Text, Integer)
                         Select x).ToList()
        End Using

        ДокПредварДаннНовыйПуть()

        'если доки надо оформить
        If CheckBox23.Checked = True Then
            ДокЗаявлениеНовыйПуть(Сотрудник)
            ДКонтрактНовыйПуть(Сотрудник)
            ДокПриказНовыйПуть(Сотрудник)
            ДокИнструкцНовыйПуть(Сотрудник)
            If MessageBox.Show("Контракт № " & arrtbox("TextBox38") & " от " & arrtmask("MaskedTextBox3") & vbCrLf & "Приказ № " & НПриказа &
 arrtbox("TextBox57") & " от " & arrtmask("MaskedTextBox3") & vbCrLf & "Заявление от " & arrtmask("MaskedTextBox3") & vbCrLf &
 "С сотрудником " & vbCrLf & Сотрудник(0).Фамилия & " " & Сотрудник(0).Имя & " " & Сотрудник(0).Отчество & vbCrLf & "Инструкция " & ИнстрП & vbCrLf & "Сформированы!" & vbCrLf & "Распечатать Документы?",
                               Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then

                ПечатьКонтрактНовыйПуть()
            End If
        Else
            MessageBox.Show("Сотрудник добавлен в базу!", Рик)
        End If



    End Sub
    Private Sub ПечатьКонтрактНовыйПуть()
        'If MessageBox.Show("Напишите количество копий для контракта! Укажите цифру 1 или 2!", Рик, ) Then
        'End If
        Do
            rz = InputBox("Напишите количество копий для контракта! Укажите цифру 1 или 2!", "1 или 2")
        Loop Until rz = 1 Or rz = 2

        'Task.WaitAll(TskArr)

        Select Case rz
            Case 1
                If ПровИнстр = 1 Then
                    massFTP.Add(СохрЗакFTP)
                    massFTP.Add(СохрКонтрFTP)
                    massFTP.Add(СохрПрикFTP)
                Else
                    massFTP.Add(СохрЗакFTP)
                    massFTP.Add(СохрКонтрFTP)
                    massFTP.Add(СохрПрикFTP)
                    massFTP.Add(ИнстрFTP)
                    massFTP.Add(ИнстрFTP)
                End If
            Case 2
                If ПровИнстр = 1 Then
                    massFTP.Add(СохрЗакFTP)
                    massFTP.Add(СохрКонтрFTP)
                    massFTP.Add(СохрПрикFTP)
                    massFTP.Add(СохрКонтрFTP)
                Else
                    massFTP.Add(СохрЗакFTP)
                    massFTP.Add(СохрКонтрFTP)
                    massFTP.Add(СохрПрикFTP)
                    massFTP.Add(СохрКонтрFTP)
                    massFTP.Add(ИнстрFTP)
                    massFTP.Add(ИнстрFTP)
                End If
        End Select
        ПечатьДоковFTP(massFTP)


        'Task.WaitAll(TskArr)


    End Sub
    Private Function МестоРаботыНовыйПуть()
        Dim Тип As String = "", Название As String = ""
        Using dbcx As New DbAll1DataContext
            Dim var = (From x In dbcx.ОбъектОбщепита.AsEnumerable
                       Where x.АдресОбъекта = arrtcom("ComboBox18") And x.НазвОрг = arrtcom("ComboBox1")
                       Select x).FirstOrDefault()

            If var IsNot Nothing Then

                If var.НазОбъекта <> "" Then
                    Название = """" & var.НазОбъекта & ""","
                End If

                If var.ТипОбъекта <> "" Then
                    Тип = Strings.Trim(Strings.LCase(var.ТипОбъекта))
                End If

            Else

                MessageBox.Show("Выберите другой объект общепита!", Рик)
                Return 1
            End If

        End Using



        If Название = "" And Тип = "" Then
            МестоРаб = arrtcom("ComboBox18")
        ElseIf Название <> "" And Тип = "" Then
            МестоРаб = Название & " " & arrtcom("ComboBox18")
        ElseIf Название = "" And Тип <> "" Then
            МестоРаб = Тип & " " & arrtcom("ComboBox18")
        Else
            МестоРаб = Тип & " " & Название & " " & arrtcom("ComboBox18")
        End If

        Return 0

    End Function
    Private Sub ОбновлениеСотрудникаКонтрактНовыйПуть(ByVal IDSotr As Integer)

        Dim ФИО As String
        'обновляем данные во всех таблицах по Контракту
        Using dbcx As New DbAll1DataContext
            Dim var = (From x In dbcx.Сотрудники.AsEnumerable
                       Join y In dbcx.Штатное.AsEnumerable On x.КодСотрудники Equals y.ИДСотр
                       Join z In dbcx.ПродлКонтракта.AsEnumerable On x.КодСотрудники Equals z.IDСотр
                       Join u In dbcx.КарточкаСотрудника.AsEnumerable On x.КодСотрудники Equals u.IDСотр
                       Join s In dbcx.ДогСотрудн.AsEnumerable On x.КодСотрудники Equals s.IDСотр
                       Where x.КодСотрудники = IDSotr
                       Select x, y, z, u, s).FirstOrDefault
            If var Is Nothing Then
                Exit Sub
            End If

            ФИО = var.x.ФИОСборное

            If arrtbox("TextBox46") = "" Then
                Dtxt46 = Nothing
            ElseIf arrtbox("TextBox46").Length > 2 Then
                Dtxt46 = CType(Replace(arrtbox("TextBox46"), ".", ","), Double)
            Else
                Dtxt46 = CType(arrtbox("TextBox46"), Integer)
            End If

            Dim ФОТ2 As Double = Replace(arrtbox("TextBox48"), ".", ",")
            Dim ФОТ3 As Double = Replace(arrtcom("ComboBox10"), ".", ",")
            ФОТ2 = ФОТ2 * ФОТ3

            Dim dcx As Double = Replace(arrtbox("TextBox48"), ".", ",")
            Dim fgd As Double = CType(arrtbox("TextBox33") & "," & arrtbox("TextBox44"), Double)

            'обновляем данные в таблице Штатное
            Try
                var.y.ПовышОклРуб = Math.Round(fgd * Replace(Dtxt46, ",", ".") / 100, 2)
            Catch ex As Exception
                var.y.ПовышОклРуб = Math.Round(fgd * Replace(Dtxt46, ".", ",") / 100, 2)
            End Try

            Try
                var.y.ЧасоваяТарифСтавка = Math.Round(Replace(dcx, ",", ".") / 168, 2)
            Catch ex As Exception
                var.y.ЧасоваяТарифСтавка = Math.Round(Replace(dcx, ".", ",") / 168, 2)
            End Try

            If arrtcom("ComboBox7") = "" Then
                var.y.Разряд = ""
            Else
                var.y.Разряд = arrtcom("ComboBox7")
            End If
            var.y.Должность = arrtcom("ComboBox9")

            var.y.ТарифнаяСтавка = Math.Round(CType(arrtbox("TextBox33") & "," & arrtbox("TextBox44"), Double), 2)
            var.y.ПовышОклПроц = Replace(Dtxt46, ",", ".")

            Try
                var.y.РасчДолжностнОклад = Replace(dcx, ".", ",")
            Catch ex As Exception
                var.y.РасчДолжностнОклад = Replace(dcx, ",", ".")
            End Try

            var.y.Отдел = arrtcom("ComboBox8")

            Try
                var.y.ФонОплатыТруда = Replace(ФОТ2, ".", ",")
            Catch ex As Exception
                var.y.ФонОплатыТруда = Replace(ФОТ2, ",", ".")
            End Try

            'Обновляем таблицу Продление Контракта.
            If var.z.ПервоеПродлениеС = "" Then
                var.z.ДатаПриема = arrtmask("MaskedTextBox4")
                var.z.ДатаОкончания = arrtmask("MaskedTextBox5")
                var.z.СрокКонтракта = arrtcom("ComboBox11")
                var.z.НомерУвед = arrtbox("TextBox38")
            Else
                If Not var.z.ДатаПриема = arrtmask("MaskedTextBox4") Or Not var.z.ДатаОкончания = arrtmask("MaskedTextBox5") Or Not var.z.СрокКонтракта = arrtcom("ComboBox11") Then
                    If MessageBox.Show("С данным сотрудником был продлен контракт" & vbCrLf & "Если вы продолжите изменения то данные о продлении будут удалены!" & vbCrLf & "Продолжить?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                        var.z.ДатаПриема = arrtmask("MaskedTextBox4")
                        var.z.ДатаОкончания = arrtmask("MaskedTextBox5")
                        var.z.СрокКонтракта = arrtcom("ComboBox11")
                        var.z.НомерУвед = arrtbox("TextBox38")

                        var.z.ПервоеПродлениеС = ""
                        var.z.ПервоеПродлениеПо = ""
                        var.z.ПервоеПродлениеСрок = ""

                        var.z.ВтороеПродлениеПо = ""
                        var.z.ВтороеПродлениеС = ""
                        var.z.ВтороеПродлениеСрок = ""

                        var.z.ТретьеПродлениеПо = ""
                        var.z.ТретьеПродлениеС = ""
                        var.z.ТретьеПродлениеСрок = ""

                        var.z.ЧетвертоеПродлениеПо = ""
                        var.z.ЧетвертоеПродлениеС = ""
                        var.z.ЧетвертоеПродлениеСрок = ""

                        var.z.НомерУвед1 = ""
                        var.z.НомерУвед2 = ""
                        var.z.НомерУвед3 = ""
                        var.z.НомерУвед4 = ""

                        var.z.Итого = ""
                    End If
                End If
            End If


            Dim _ПоСовмест, _СуммирУчет As String
            If CheckBox2.Checked = True Then
                _ПоСовмест = "по совместительству"
            Else
                _ПоСовмест = ""
            End If
            If CheckBox4.Checked = True Then
                _СуммирУчет = "Да"
            Else
                _СуммирУчет = ""
            End If


            'Обновляем таблицу КарточкаСотрудника данные контракта и обновляем таблицу.
            With var.u
                .ДатаПриема = arrtmask("MaskedTextBox4")
                .СрокКонтракта = arrtcom("ComboBox11")
                .ТипРаботы = arrtcom("ComboBox15")
                .Ставка = arrtcom("ComboBox10")
                .ВремяНачРаботы = arrtcom("ComboBox12")
                .ПродолРабДня = arrtcom("ComboBox16")
                .Обед = arrtbox("TextBox49")
                .ОкончРабДня = arrtbox("TextBox50")
                .ДатаУведомлПродКонтр = ДатаУведомл(arrtcom("ComboBox11"), arrtmask("MaskedTextBox4"))
                .АдресОбъектаОбщепита = arrtcom("ComboBox18")
                .ДатаЗарплаты = arrtbox("TextBox40")
                .ДатаАванса = arrtbox("TextBox56")
                .ПоСовмест = _ПоСовмест
                .СуммирУчет = _СуммирУчет
            End With

            If Примечани = "" Or Примечани Is Nothing Then
                var.u.Примечание = ""
            Else
                var.u.Примечание = Примечани
            End If

            'Вставляем в таблицу ДогСотрудн данные контракта и обновляем таблицу.
            With var.s
                .Контракт = arrtbox("TextBox38")
                .ДатаКонтракта = arrtmask("MaskedTextBox3")
                .СрокОкончКонтр = arrtmask("MaskedTextBox5")
                .Приказ = НПриказа
                .Датаприказа = arrtmask("MaskedTextBox3")
            End With

            dbcx.SubmitChanges()

        End Using

        Статистика1(ФИО, "Обновление данных контракта сотрудника", arrtcom("ComboBox1"))


    End Sub

    Private Sub НовыйПутьКонтракт()

        'Проверяем Контракт 
        If ПроверкаЗаполненностиВкладкиКонтрактНовыйПуть() = 1 Then
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor

        СохраняемКонтролыВСписки(TabPage2)

        'оформление места работы 
        If МестоРаботыНовыйПуть() = 1 Then Exit Sub

        If IsNumeric(Label96.Text) = False Then
            MessageBox.Show("Нет идентификатора сотрудника!", Рик)
            Exit Sub
        End If

        Dim IDSotr As Integer = CType(Label96.Text, Integer)
        'добавляем сотрудника в базу или обновляем данные
        Using dbcx As New DbAll1DataContext
            Dim var = (From x In dbcx.Сотрудники.AsEnumerable
                       Where x.ДанныеИзСправочника = "True" And x.ТипОтношения = "(кт)" And x.КодСотрудники = IDSotr
                       Select x).FirstOrDefault
            If var Is Nothing Then
                ДобавлениеСотрудникаНовыйПуть()
            Else
                ОбновлениеСотрудникаКонтрактНовыйПуть(IDSotr)
            End If
        End Using


        If CheckBox23.Checked = False Then
            If MessageBox.Show("Оформить пакет документов?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                MessageBox.Show("Данные изменены!", Рик)
                ОчисткаАктивнойВкладкиНовыйПуть()
                Me.Cursor = Cursors.Default
                Exit Sub
            Else
                CheckBox23.Checked = True
            End If
        End If


        'контракт оформление документов
        ДокиКонтрактНовыйПуть()

        ОчисткаАктивнойВкладкиНовыйПуть()
        Me.Cursor = Cursors.Default

    End Sub
    Private Sub ДобавлениеСотрудникаНовыйПуть()
        Dim IDSotr As Integer = CType(Label96.Text, Integer)

        'добавляем данные в таблицу Сотрудники
        Using dbcx As New DbAll1DataContext
            Dim var = (From x In dbcx.Сотрудники.AsEnumerable
                       Where x.КодСотрудники = IDSotr
                       Select x).FirstOrDefault
            If var IsNot Nothing Then
                var.ПровДатыКонтр = Trim(arrtmask("MaskedTextBox3"))
                var.ТипОтношения = "(кт)"
                dbcx.SubmitChanges()
            End If
        End Using

        If arrtbox("TextBox46") = "" Then
            Dtxt46 = Nothing
        ElseIf arrtbox("TextBox46").Length > 2 Then
            Dtxt46 = CType(Replace(arrtbox("TextBox46"), ".", ","), Double)
        Else
            Dtxt46 = CType(arrtbox("TextBox46"), Integer)
        End If

        Dim ФОТ2 As Double = Replace(arrtbox("TextBox48"), ".", ",")
        Dim ФОТ3 As Double = Replace(arrtcom("ComboBox10"), ".", ",")
        ФОТ2 = ФОТ2 * ФОТ3

        'Dim ФОТ2 As Double = Math.Round(TextBox48.Text * ComboBox10.Text, 2)

        IDsot1 = Nothing
        IDsot1 = IDSotr

        Dim dcx As Double = Replace(arrtbox("TextBox48"), ".", ",")
        Dim fgd As Double = CType(arrtbox("TextBox33") & "," & arrtbox("TextBox44"), Double)

        'вставляем данные в таблицу штатное
        Dim idШтатное As Integer
        Using dbcx As New DbAll1DataContext
            Dim f As New Штатное()

            Try
                f.ПовышОклРуб = Math.Round(fgd * Replace(Dtxt46, ",", ".") / 100, 2)
            Catch ex As Exception
                f.ПовышОклРуб = Math.Round(fgd * Replace(Dtxt46, ".", ",") / 100, 2)
            End Try

            Try
                f.ЧасоваяТарифСтавка = Math.Round(Replace(dcx, ",", ".") / 168, 2)
            Catch ex As Exception
                f.ЧасоваяТарифСтавка = Math.Round(Replace(dcx, ".", ",") / 168, 2)
            End Try

            If arrtcom("ComboBox7") = "" Then
                f.Разряд = ""
            Else
                f.Разряд = arrtcom("ComboBox7")
            End If

            f.ИДСотр = IDSotr
            f.Должность = arrtcom("ComboBox9")

            f.ТарифнаяСтавка = Math.Round(CType(arrtbox("TextBox33") & "," & arrtbox("TextBox44"), Double), 2)
            f.ПовышОклПроц = Replace(Dtxt46, ",", ".")

            Try
                f.РасчДолжностнОклад = Replace(dcx, ".", ",")
            Catch ex As Exception
                f.РасчДолжностнОклад = Replace(dcx, ",", ".")
            End Try

            f.Отдел = arrtcom("ComboBox8")

            Try
                f.ФонОплатыТруда = Replace(ФОТ2, ".", ",")
            Catch ex As Exception
                f.ФонОплатыТруда = Replace(ФОТ2, ",", ".")
            End Try


            dbcx.Штатное.InsertOnSubmit(f)
            dbcx.SubmitChanges()
            idШтатное = f.Код
        End Using

        'Вставляем в таблицу продление контракта.
        Dim sot
        Using dbcx As New DbAll1DataContext

            sot = (From x In dbcx.Сотрудники.AsEnumerable
                   Where x.КодСотрудники = IDSotr
                   Select x).FirstOrDefault



            Dim var As New ПродлКонтракта()
            var.IDСотр = IDSotr
            var.ФИО = sot.ФИОСборное
            var.ДатаПриема = arrtmask("MaskedTextBox4")
            var.ДатаОкончания = arrtmask("MaskedTextBox5")
            var.СрокКонтракта = arrtcom("ComboBox11")
            var.НомерУвед = arrtbox("TextBox38")

            dbcx.ПродлКонтракта.InsertOnSubmit(var)
            dbcx.SubmitChanges()

        End Using


        Dim _ПоСовмест, _СуммирУчет As String
        If CheckBox2.Checked = True Then
            _ПоСовмест = "по совместительству"
        Else
            _ПоСовмест = ""
        End If
        If CheckBox4.Checked = True Then
            _СуммирУчет = "Да"
        Else
            _СуммирУчет = ""
        End If

        'Вставляем в таблицу Карточкасотрудника данные контракта и обновляем таблицу.
        Using dbcx As New DbAll1DataContext
            Dim var As New КарточкаСотрудника With {
                .IDСотр = IDSotr,
                .ДатаПриема = arrtmask("MaskedTextBox4"),
                .СрокКонтракта = arrtcom("ComboBox11"),
                .ТипРаботы = arrtcom("ComboBox15"),
                .Ставка = arrtcom("ComboBox10"),
                .ВремяНачРаботы = arrtcom("ComboBox12"),
                .ПродолРабДня = arrtcom("ComboBox16"),
                .Обед = arrtbox("TextBox49"),
                .ОкончРабДня = arrtbox("TextBox50"),
                .ДатаУведомлПродКонтр = ДатаУведомл(arrtcom("ComboBox11"), arrtmask("MaskedTextBox4")),
                .АдресОбъектаОбщепита = arrtcom("ComboBox18"),
                .ДатаЗарплаты = arrtbox("TextBox40"),
                .ДатаАванса = arrtbox("TextBox56"),
                .ПоСовмест = _ПоСовмест,
                .СуммирУчет = _СуммирУчет
            }
            If Примечани = "" Or Примечани Is Nothing Then
                var.Примечание = ""
            Else
                var.Примечание = Примечани
            End If

            dbcx.КарточкаСотрудника.InsertOnSubmit(var)
            dbcx.SubmitChanges()

        End Using

        'Вставляем в таблицу ДогСотрудн данные контракта и обновляем таблицу.
        Using dbcx As New DbAll1DataContext
            Dim var As New ДогСотрудн() With {
                .IDСотр = IDSotr,
                .Контракт = arrtbox("TextBox38"),
                .ДатаКонтракта = arrtmask("MaskedTextBox3"),
                .СрокОкончКонтр = arrtmask("MaskedTextBox5"),
                .Приказ = НПриказа,
                .Датаприказа = arrtmask("MaskedTextBox3")
            }

            dbcx.ДогСотрудн.InsertOnSubmit(var)
            dbcx.SubmitChanges()
        End Using


        Статистика1(sot.ФИОСборное, "Добавление нового сотрудника", arrtcom("ComboBox1"))

    End Sub
    Private Sub ОчисткаАктивнойВкладкиНовыйПуть()


        'перебираем все контролы в гроупбоксах

        For Each gp In TabControl1.SelectedTab.Controls.OfType(Of GroupBox) 'таб2 

            For Each tf In gp.Controls.OfType(Of TextBox)
                tf.Text = ""
            Next

            For Each tx In gp.Controls.OfType(Of ComboBox)
                tx.Text = ""
            Next

            For Each ts In gp.Controls.OfType(Of MaskedTextBox)
                ts.Text = ""   'arrtmask.Add(ts.Name, ts.Text)
            Next

            For Each tx1 In gp.Controls.OfType(Of GroupBox)

                For Each tx In tx1.Controls.OfType(Of ComboBox)
                    tx.Text = ""   'arrtcom.Add(tx.Name, tx.Text)
                Next
                For Each ts In tx1.Controls.OfType(Of MaskedTextBox)
                    ts.Text = ""  'arrtmask.Add(ts.Name, ts.Text)
                Next
                For Each tf In tx1.Controls.OfType(Of TextBox)
                    tf.Text = "" 'arrtbox.Add(tf.Name, tf.Text)
                Next

            Next

        Next
    End Sub
    Private Sub СохраняемКонтролыВСписки(ByVal TabPageName As TabPage)

        Parallel.Invoke(Sub() ЗаполнМассВнеТабах())


        'перебираем все контролы в гроупбоксах

        For Each gp In TabPageName.Controls.OfType(Of GroupBox) 'таб2 

            For Each tf In gp.Controls.OfType(Of TextBox)
                arrtbox.Add(tf.Name, tf.Text)
            Next

            For Each tx In gp.Controls.OfType(Of ComboBox)
                arrtcom.Add(tx.Name, tx.Text)
            Next

            For Each ts In gp.Controls.OfType(Of MaskedTextBox)
                arrtmask.Add(ts.Name, ts.Text)
            Next

            For Each tx1 In gp.Controls.OfType(Of GroupBox)

                For Each tx In tx1.Controls.OfType(Of ComboBox)
                    arrtcom.Add(tx.Name, tx.Text)
                Next
                For Each ts In tx1.Controls.OfType(Of MaskedTextBox)
                    arrtmask.Add(ts.Name, ts.Text)
                Next
                For Each tf In tx1.Controls.OfType(Of TextBox)
                    arrtbox.Add(tf.Name, tf.Text)
                Next

            Next

        Next


    End Sub
    Private Sub УдалениеСотрНовыйПуть()

        Me.Cursor = Cursors.WaitCursor

        Dim idc As Integer = CType(Label96.Text, Integer)

        'Проверяем есть в папке путиДокументов данные, если есть копируем и удаляем доки на сервере и в таблице ПутиДокументов
        Using dbcx As New DbAll1DataContext
            Dim f = (From x In dbcx.ПутиДокументов.AsEnumerable
                     Where x.IDСотрудник = idc
                     Select x).ToList
            If f.Count > 0 Then
                Dim vart2 = From x In dtPutiDokumentovAll.AsEnumerable Where Not IsDBNull(x.Item("IDСотрудник")) Select x
                Dim vart = From x1 In vart2 Where x1.Item("IDСотрудник") = idc Select x1.Item("Путь") & x1.Item("ИмяФайла")
                For b As Integer = 0 To vart.Count - 1
                    DeleteFluentFTP(vart(b).ToString)
                Next
            End If
        End Using



        Using dbcx As New DbAll1DataContext

            'удалем из таблицы Сотрудники
            Dim var = (From x In dbcx.Сотрудники.AsEnumerable
                       Where x.КодСотрудники = idc
                       Select x).SingleOrDefault
            If var IsNot Nothing Then
                dbcx.Сотрудники.DeleteOnSubmit(var)
                dbcx.SubmitChanges()
            End If

            ''удалем из таблицы КарточкаСотрудника
            'Dim var1 = (From x In dbcx.КарточкаСотрудника.AsEnumerable
            '            Where x.IDСотр = idc
            '            Select x).SingleOrDefault
            'If var1 IsNot Nothing Then
            '    dbcx.КарточкаСотрудника.DeleteOnSubmit(var1)
            '    dbcx.SubmitChanges()
            'End If

            ''удалем из таблицы ДогСотрудн
            'Dim var2 = (From x In dbcx.ДогСотрудн.AsEnumerable
            '            Where x.IDСотр = idc
            '            Select x).SingleOrDefault
            'If var2 IsNot Nothing Then
            '    dbcx.ДогСотрудн.DeleteOnSubmit(var2)
            '    dbcx.SubmitChanges()
            'End If

            ''удалем из таблицы Штатное
            'Dim var3 = (From x In dbcx.Штатное.AsEnumerable
            '            Where x.ИДСотр = idc
            '            Select x).SingleOrDefault
            'If var3 IsNot Nothing Then
            '    dbcx.Штатное.DeleteOnSubmit(var3)
            '    dbcx.SubmitChanges()
            'End If

            ''удалем из таблицы ПутиДокументов
            'Dim var4 = (From x In dbcx.ПутиДокументов.AsEnumerable
            '                Where x.IDСотрудник = idc
            '                Select x).SingleOrDefault
            'If var4 IsNot Nothing Then
            '    dbcx.ПутиДокументов.DeleteOnSubmit(var4)
            '    dbcx.SubmitChanges()
            'End If


        End Using

        Статистика1(ComboBox19.Text, "Удаление сотрудника", ComboBox1.Text)
        MessageBox.Show("Сотрудник удален из базы!", Рик)
        CheckBox27.Checked = False


        Me.Cursor = Cursors.Default

        ОчисткаАктивнойВкладкиНовыйПуть()

        Button23.PerformClick()

        ComboBox19.Text = ""
        Label96.Text = ""

    End Sub
    Private Sub Удаление()
        If IsNumeric(Label96.Text) Then
            Using dbcx As New DbAll1DataContext  'определяем, оформлен ли сотрудник через справочник
                Dim var = (From x In dbcx.Сотрудники.AsEnumerable
                           Where x.КодСотрудники = CType(Label96.Text, Integer)
                           Select x.ДанныеИзСправочника).FirstOrDefault()
                If var = "True" Then
                    УдалениеСотрНовыйПуть()
                Else
                    УдалениеСотр()

                End If
            End Using

        End If




    End Sub

    Public Async Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click



        If arrtbox.Any Then
            arrtbox.Clear()
        End If

        If arrtmask.Any Then
            arrtmask.Clear()
        End If

        If arrtcom.Any Then
            arrtcom.Clear()
        End If

        If CheckBox27.Checked = True Then
            Удаление()

            Me.Cursor = Cursors.Default
            Exit Sub
        End If




        'прверяем по новому пути идти или по старому
        If CheckBox9.Checked = True And CheckBox9.Enabled = True And CheckBox5.Checked = True Then

            НовыйПутьДПИное()
            Exit Sub
        ElseIf CheckBox8.Checked = True And CheckBox8.Enabled = True And CheckBox5.Checked = True Then
            НовыйПутьДПЦена()
            Exit Sub
        ElseIf CheckBox10.Checked = True And CheckBox10.Enabled = True And CheckBox5.Checked = True Then

            НовыйПутьКонтракт()
            Exit Sub
        End If





        'For i = 1 To 77
        '    arrtbox.Add(TextBox1.Text)
        '    'arrtbox.Add(Me.Controls("TextBox" & i).Text)
        'Next i
        ЗаполнМассВТабах()
        ЗаполнМассВнеТабах()
        ОчисткаМаяковAsync()
        dtPutiDokumentov()

        If ПроверкаКонтрактИлиПодрядДобавляем() = 1 Then
            Exit Sub
        End If



        Dim PrintPapie As Integer = 0
        If CheckBox5.Checked = False Then
            If MessageBox.Show("Сформировать пакет документов?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                PrintPapie = 1
            End If
        End If

        Комбы()
        'Dim СборДанОрг As New Thread(AddressOf СборДаннОрганиз)
        'СборДанОрг.Start()

        'Dim d As Integer = УдалениеСотр() 'удаление сотрудника
        'If d = 1 Then
        '    Me.Cursor = Cursors.WaitCursor

        '    Me.Cursor = Cursors.Default
        '    Exit Sub
        'End If

        СрокКонтр = ComboBox11.Text
        Ставка = ComboBox10.Text
        CorName = Mid(TextBox2.Text, 1, 1) & "."
        CorOtch = Mid(TextBox3.Text, 1, 1) & "."

        ReDim Заявление(-1)
        ReDim Контракт(-1)
        ReDim Приказ(-1)



        Заявление = {MaskedTextBox3.Text, Trim(TextBox6.Text), Trim(TextBox5.Text), Trim(TextBox4.Text),
            TextBox21.Text, MaskedTextBox10.Text, ComboBox9.Text, MaskedTextBox4.Text, ComboBox10.Text,
            Trim(TextBox1.Text), Trim(TextBox2.Text), Trim(TextBox3.Text), MaskedTextBox3.Text, Trim(TextBox11.Text), Trim(TextBox10.Text)}
        'For index = 0 To Заявление.GetUpperBound(0) 'перебор массива
        '    Debug.WriteLine(Заявление(index))
        'Next

        Контракт = {TextBox38.Text, MaskedTextBox4.Text, ComboBox11.Text, MaskedTextBox5.Text, TextBox33.Text,
            TextBox12.Text, TextBox7.Text, MaskedTextBox1.Text, TextBox9.Text, TextBox8.Text, TextBox43.Text}

        Приказ = {TextBox42.Text, НПриказа, Trim(TextBox34.Text), Trim(TextBox5.Text), Trim(TextBox4.Text),
            MaskedTextBox3.Text, Me.MaskedTextBox4.Text, MaskedTextBox5.Text,
            TextBox38.Text, Trim(TextBox1.Text)}




        If CheckBox7.Checked = False Then

            If МестоРаботы() = 1 Then
                Exit Sub
            End If
        End If



        If CheckBox7.Checked = False Then




            ПрКонт = ПровДляКонтр()
            If ПрКонт = 1 Then
                'Соед(0)
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
        Else
            ПрПодр = ПровДляПодряда()
            If ПрПодр = 1 Then
                'Соед(0)

                Exit Sub
            End If
        End If


        If ПровИндивидКонтр(ComboBox1.Text) = False And CheckBox5.Checked = True And CheckBox7.Checked = False And Решение = "" Then 'обновление данных

            ОбщОбновл()

            очПоля = 1
            CheckBox6.Checked = True
            CheckBox6.Checked = False
            ComboBox20.Text = ""
            ComboBox2.Text = ""
            ComboBox21.Text = ""
            TextBox40.Text = ""
            TextBox56.Text = ""
            ComboBox17.Text = ""
            MaskedTextBox3.Text = DateTime.Now.ToString("dd.MM.yyyy")
            MaskedTextBox4.Text = Format(Now, "dd.MM.yyyy")
            Label85.Text = "NO"
            Label89.Text = "NO"
            Label90.Text = "NO"
            Parallel.Invoke(Sub() Com1sel())
            Me.Cursor = Cursors.WaitCursor

            Me.Cursor = Cursors.Default
            ALLALL()
            Exit Sub
        End If


        a = ComboBox1.Text 'проверка есть ли уже такой сотрудник в базе?
        surName = Trim(TextBox1.Text)
        surNameAll = Trim(surName) & " " & Trim(Me.TextBox2.Text) & " " & Trim(Me.TextBox3.Text)


        If PrintPapie = 0 Then
            Dim tfd As String = MsgBox("Сохранить данные?", vbOKCancel, Рик)
            If tfd = "2" Then
                'Соед(0)

                Exit Sub
            End If
        End If

        Me.Cursor = Cursors.WaitCursor

        'If ComboBox1.Text = "Амасейлс" Then
        '    амасейлс()
        '    Me.Cursor = Cursors.Default
        '    Exit Sub
        'End If

        'If ComboBox1.Text = "ЛемеЛ Лабс" Then
        '    ЛемеЛ()
        '    Me.Cursor = Cursors.Default
        '    Exit Sub
        'End If

        'If ComboBox1.Text = "Пинфуд Сервис" And CheckBox7.Checked = False Then
        '    ПинфудСервис()
        '    Me.Cursor = Cursors.Default
        '    Exit Sub
        'End If




        If CheckBox7.Checked = True And Not ComboBox1.Text = "Амасейлс" Then  'договор подряда
            Try
                Поток.IsBackground = True
                Поток1.IsBackground = True
                Поток.Start()
                Поток1.Start()
            Catch ex As Exception
                If ex.ToString.Contains("Поток не существует; нельзя получить доступ к данным о состоянии.") Then
                    Поток = New Thread(AddressOf ДанныеКлиентаДогПодряда)
                    Поток1 = New Thread(AddressOf НалогиИОбязанДогПодряда)
                    Поток.IsBackground = True
                    Поток1.IsBackground = True
                    Поток.Start()
                    Поток1.Start()
                End If

            End Try

            ДогПодряда()
            ALLALL()

            If MessageBox.Show("Распечатать договор-подряда?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                Me.Cursor = Cursors.WaitCursor
                massFTP.Add(СохрДогПодрFTP)
                massFTP.Add(СохрДогПодрFTP)
                ПечатьДоковFTP(massFTP)
            End If
            Me.Cursor = Cursors.WaitCursor

            Me.Cursor = Cursors.Default
            Exit Sub
        End If

        Dim Про As Integer
        Await Task.Run(Sub() Про = ПровДублСотр())
        If Про = 1 Then
            Me.Cursor = Cursors.Default
            Exit Sub
        End If

        Dim ДобСотТаск As Task = New Task(AddressOf ДобавлНовогоСотрудника)
        ДобСотТаск.Start()
        'Parallel.Invoke(Sub() ДобавлНовогоСотрудника())

        If PrintPapie = 0 Then
            MessageBox.Show("Сотрудник добавлен!", Рик)
        End If

        If PrintPapie = 1 Then 'основной модуль по оформлению документов
            ДобСотТаск.Wait()
            Доки("общ")

            Me.Cursor = Cursors.Default
        End If

        If PrintPapie = 1 Then

            If MessageBox.Show("Контракт № " & TextBox38.Text & " от " & MaskedTextBox3.Text & vbCrLf & "Приказ № " & НПриказа &
          TextBox57.Text & " от " & MaskedTextBox3.Text & vbCrLf & "Заявление от " & MaskedTextBox3.Text & vbCrLf &
          "С сотрудником " & vbCrLf & Trim(TextBox1.Text) & " " & Trim(TextBox2.Text) & " " & Trim(TextBox3.Text) & vbCrLf & "Инструкция " & ИнстрП & vbCrLf & "Сформированы!" & vbCrLf & "Распечатать Документы?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then

                Do
                    rz = InputBox("Напишите количество копий для контракта! Укажите цифру 1 или 2!", "1 или 2")
                Loop Until rz = 1 Or rz = 2

                Task.WaitAll(TskArr)

                Select Case rz
                    Case 1
                        If ПровИнстр = 1 Then
                            massFTP.Add(СохрЗакFTP)
                            massFTP.Add(СохрКонтрFTP)
                            massFTP.Add(СохрПрикFTP)
                        Else
                            massFTP.Add(СохрЗакFTP)
                            massFTP.Add(СохрКонтрFTP)
                            massFTP.Add(СохрПрикFTP)
                            massFTP.Add(ИнстрFTP)
                            massFTP.Add(ИнстрFTP)
                        End If
                    Case 2
                        If ПровИнстр = 1 Then
                            massFTP.Add(СохрЗакFTP)
                            massFTP.Add(СохрКонтрFTP)
                            massFTP.Add(СохрПрикFTP)
                            massFTP.Add(СохрКонтрFTP)
                        Else
                            massFTP.Add(СохрЗакFTP)
                            massFTP.Add(СохрКонтрFTP)
                            massFTP.Add(СохрПрикFTP)
                            massFTP.Add(СохрКонтрFTP)
                            massFTP.Add(ИнстрFTP)
                            massFTP.Add(ИнстрFTP)
                        End If
                End Select
                ПечатьДоковFTP(massFTP)

            Else
                Task.WaitAll(TskArr)

            End If
        Else


        End If

        Me.Cursor = Cursors.Default


        'MessageBox.Show("ok")
        'sw.Stop()
        очПоля = 1
        CheckBox6.Checked = True
        CheckBox6.Checked = False
        ComboBox20.Text = ""
        ComboBox2.Text = ""
        ComboBox21.Text = ""
        TextBox40.Text = ""
        TextBox56.Text = ""
        ComboBox17.Text = ""
        MaskedTextBox3.Text = DateTime.Now.ToString("dd.MM.yyyy")
        MaskedTextBox4.Text = Format(Now, "dd.MM.yyyy")
        Label85.Text = "NO"
        Label89.Text = "NO"
        Label90.Text = "NO"

        Com1sel()

    End Sub

    Private Sub ОбщОбновл()
        If CheckBox5.Checked = True And CheckBox7.Checked = False Then 'если надо внезти изменения и распечатать

            Dim rf As Task = New Task(AddressOf ОбновлСотрудника)
            rf.Start()
            MessageBox.Show("Все данные сотрудника " & TextBox6.Text & " " & TextBox5.Text & " " & TextBox4.Text & vbCrLf & " удачно внесены в базу!", Рик, MessageBoxButtons.OK, MessageBoxIcon.None)
            If CheckBox23.Checked = True Then
                'СборДаннОрганиз()
                rf.Wait()
                Доки("общ")

                If MessageBox.Show("Все данные изменены. Документы оформлены!" & vbCrLf & " Распечатать? ", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.None) = DialogResult.OK Then
                    Do
                        Try
                            rz = CType((InputBox("Напишите количество копий для контракта! Укажите цифру 1 или 2!", "1 или 2")), Integer)
                        Catch ex As Exception
                            rz = 0
                        End Try

                    Loop Until rz = 1 Or rz = 2
                    Me.Cursor = Cursors.WaitCursor

                    Task.WaitAll(TskArr)

                    'Select Case rz
                    '    Case 1
                    '        If ПровИнстр = 1 Then
                    '            mass = {СохрЗак, СохрКонтр, СохрПрик}
                    '        Else
                    '            mass = {СохрЗак, СохрКонтр, СохрПрик, Инстр, Инстр}
                    '        End If
                    '    Case 2
                    '        If ПровИнстр = 1 Then
                    '            mass = {СохрЗак, СохрКонтр, СохрПрик, СохрКонтр}
                    '        Else
                    '            mass = {СохрЗак, СохрКонтр, СохрПрик, СохрКонтр, Инстр, Инстр}
                    '        End If
                    'End Select
                    Select Case rz
                        Case 1
                            If ПровИнстр = 1 Then
                                massFTP.Add(СохрЗакFTP)
                                massFTP.Add(СохрКонтрFTP)
                                massFTP.Add(СохрПрикFTP)
                            Else
                                massFTP.Add(СохрЗакFTP)
                                massFTP.Add(СохрКонтрFTP)
                                massFTP.Add(СохрПрикFTP)
                                massFTP.Add(ИнстрFTP)
                                massFTP.Add(ИнстрFTP)

                            End If
                        Case 2
                            If ПровИнстр = 1 Then
                                massFTP.Add(СохрЗакFTP)
                                massFTP.Add(СохрКонтрFTP)
                                massFTP.Add(СохрПрикFTP)
                                massFTP.Add(СохрКонтрFTP)


                            Else
                                massFTP.Add(СохрЗакFTP)
                                massFTP.Add(СохрКонтрFTP)
                                massFTP.Add(СохрПрикFTP)
                                massFTP.Add(СохрКонтрFTP)
                                massFTP.Add(ИнстрFTP)
                                massFTP.Add(ИнстрFTP)

                            End If

                    End Select
                    ПечатьДоковFTP(massFTP)


                End If
            End If
            Me.Cursor = Cursors.WaitCursor
            If TskArr IsNot Nothing Then
                Task.WaitAll(TskArr)
            End If


            Me.Cursor = Cursors.Default

            If CheckBox26.Checked = True Then
                CheckBox26.Checked = False
            End If

        End If
    End Sub
    Private Sub амасейлс()

        inp = TextBox51.Text
        ДатРожд = MaskedTextBox9.Text

        If CheckBox5.Checked = False Then

            ДобавлНовогоСотрудника()
            If MessageBox.Show("Сотрудник добавлен в базу! Оформить Документы?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                Доки("амасейлс")

                MessageBox.Show("Сотрудник добавлен!", Рик)
                If MessageBox.Show("Контракт № " & TextBox38.Text & " от " & MaskedTextBox3.Text & vbCrLf & "Приказ № " & НПриказа &
                  TextBox57.Text & " от " & MaskedTextBox3.Text & vbCrLf & "Заявление от " & MaskedTextBox3.Text & vbCrLf &
                  "С сотрудником " & vbCrLf & TextBox1.Text & " " & TextBox2.Text & " " & TextBox3.Text & vbCrLf & "Инструкция " & ИнстрП & vbCrLf & "Сформированы!" & vbCrLf & "Распечатать Документы?",
                  Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then

                    Task.WaitAll(TskArr)
                    ПечатьДоковКол(СохрАмасейл, 2)
                    ПечатьДоковКол(СохрЗак, 1)
                    ПечатьДоковКол(СохрПрик, 1)

                    If Not ПровИнстр = 1 Then
                        ПечатьДоковКол(Инстр, 2)
                    End If
                End If

            End If
            'СборДаннОрганиз()


        End If

        If CheckBox5.Checked = True And CheckBox7.Checked = False Then
            'KillProc()
            Dim rf As Task = New Task(AddressOf ОбновлСотрудника)
            rf.Start()
            MessageBox.Show("Все данные сотрудника " & TextBox6.Text & " " & TextBox5.Text & " " & TextBox4.Text & vbCrLf & " обновлены!", Рик, MessageBoxButtons.OK, MessageBoxIcon.None)
            If CheckBox23.Checked = True Then
                rf.Wait()
                Доки("амасейлс")

                MessageBox.Show("Данные изменены!", Рик)
                If MessageBox.Show("Контракт № " & TextBox38.Text & " от " & MaskedTextBox3.Text & vbCrLf & "Приказ № " & НПриказа &
                  TextBox57.Text & " от " & MaskedTextBox3.Text & vbCrLf & "Заявление от " & MaskedTextBox3.Text & vbCrLf &
                  "С сотрудником " & vbCrLf & TextBox1.Text & " " & TextBox2.Text & " " & TextBox3.Text & vbCrLf & "Инструкция " & ИнстрП & vbCrLf & "Сформированы!" & vbCrLf & "Распечатать Документы?",
                  Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then

                    Task.WaitAll(TskArr)
                    ПечатьДоковКол(СохрАмасейл, 2)
                    ПечатьДоковКол(СохрЗак, 1)
                    ПечатьДоковКол(СохрПрик, 1)
                    If Not ПровИнстр = 1 Then
                        ПечатьДоковКол(Инстр, 2)
                    End If
                End If
            End If

        End If



        очПоля = 1
        CheckBox6.Checked = True
        CheckBox6.Checked = False
        ComboBox20.Text = ""
        ComboBox2.Text = ""
        ComboBox21.Text = ""
        TextBox40.Text = ""
        TextBox56.Text = ""
        ComboBox17.Text = ""
        MaskedTextBox3.Text = DateTime.Now.ToString("dd.MM.yyyy")
        MaskedTextBox4.Text = Format(Now, "dd.MM.yyyy")
        Label85.Text = "NO"
        Label89.Text = "NO"
        Label90.Text = "NO"

        Com1sel()

    End Sub

    Private Async Sub ДокиАмасейл()
        Await Task.Delay(0)
        Dim СтавкаНов, СклонГод, ПоСовмИлиОсн As String

        Try 'проверка если есть в С: папке файл Контракт его удаляем и создаем новый
            IO.File.Copy(OnePath & "\ОБЩДОКИ\Амасейлс\Контракт Амасейлс.docx", "C:\Users\Public\Documents\Рик\Контракт Амасейлс.docx")
        Catch ex As Exception
            If ex.Message.Contains("уже существует") Then
                Try
                    IO.File.Delete("C:\Users\Public\Documents\Рик\Контракт Амасейлс.docx")
                    IO.File.Copy(OnePath & "\ОБЩДОКИ\Амасейлс\Контракт Амасейлс.docx", "C:\Users\Public\Documents\Рик\Контракт Амасейлс.docx")
                Catch e As System.IO.IOException
                    If e.Message.Contains("используется другим процессом") Then
                        ПрверкаАсинхрПотоков(Task.CurrentId)
                    End If
                End Try
            End If
            If "Контракт Амасейлс.docx" <> "" Then IO.File.Delete("C:\Users\Public\Documents\Рик\Контракт Амасейлс.docx")
            IO.File.Copy(OnePath & "\ОБЩДОКИ\Амасейлс\Контракт Амасейлс.docx", "C:\Users\Public\Documents\Рик\Контракт Амасейлс.docx")
        End Try



        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document
        oWord = CreateObject("Word.Application")
        oWord.Visible = False

        'ДолжСОконч = ДобОконч(Должность)
        СтавкаНов = Склонение(Ставка) 'склонение ставки
        СклонГод = Склонение2(СрокКонтр) ' склонение год

        If CheckBox2.Checked = True Then 'галочка по осн или по совместительству
            ПоСовмИлиОсн = "совместительству"
        Else
            ПоСовмИлиОсн = "основному месту"

        End If



        oWordDoc = oWord.Documents.Add("C:\Users\Public\Documents\Рик\Контракт Амасейлс.docx")
        With oWordDoc.Bookmarks
            .Item("КСам1").Range.Text = Trim(txtbx38)
            .Item("КСам2").Range.Text = arrtmask("MaskedTextBox3")
            .Item("КСам3").Range.Text = Trim(arrtbox("TextBox1")) & " " & Trim(arrtbox("TextBox2")) & " " & Trim(arrtbox("TextBox3"))
            If combx28 = "М" Then
                .Item("КСам4").Range.Text = "ин"
            Else
                .Item("КСам4").Range.Text = "ка"
            End If
            .Item("КСам5").Range.Text = StrConv(Trim(inp), VbStrConv.ProperCase)
            .Item("КСам6").Range.Text = Trim(arrtbox("TextBox6")) & " " & Trim(arrtbox("TextBox5")) & " " & Trim(arrtbox("TextBox4"))
            .Item("КСам7").Range.Text = Strings.LCase(ДолжСОконч)
            .Item("КСам8").Range.Text = СрокКонтр & " " & СклонГод
            .Item("КСам9").Range.Text = arrtmask("MaskedTextBox4") & " по " & arrtmask("MaskedTextBox5")
            .Item("КСам10").Range.Text = arrtbox("TextBox1") & " " & CorName & CorOtch
            .Item("КСам11").Range.Text = arrtbox("TextBox33") & "," & arrtbox("TextBox44") & " (" & arrtbox("TextBox43") & ") "
            .Item("КСам12").Range.Text = arrtbox("TextBox46")
            .Item("КСам13").Range.Text = РДОрубли & "," & РДОкопейки & " (" & arrtbox("TextBox47") & ") "

            'If СтПосле <> "" Then
            '    '.Item("КСам14").Range.Text = СтПосле & " (" & ЧислоПропис(СтПосле) & ") "
            '    '.Item("КСам24").Range.Text = ПроцПосле
            '    Dim общ As String = CType(Math.Round(CType(СтПосле, Double) + (CType(СтПосле, Double) * CType(ПроцПосле, Integer) / 100), 2), String)
            '    '.Item("КСам25").Range.Text = общ & " (" & ЧислоПропис(общ) & ") "
            'Else
            '    '.Item("КСам14").Range.Text = arrtbox("TextBox33") & "," & arrtbox("TextBox44") & " (" & arrtbox("TextBox43") & ") "
            '    '.Item("КСам24").Range.Text = arrtbox("TextBox46")
            '    '.Item("КСам25").Range.Text = РДОрубли & "," & РДОкопейки & " (" & arrtbox("TextBox47") & ") "
            'End If
            .Item("КСам15").Range.Text = Trim(arrtbox("TextBox1")) & " " & Trim(arrtbox("TextBox2")) & " " & Trim(arrtbox("TextBox3"))

            If combx28 = "М" Then

                .Item("КСам16").Range.Text = "Гражданин " & StrConv(Trim(inp), VbStrConv.ProperCase)
            Else
                .Item("КСам16").Range.Text = "Гражданка " & StrConv(Trim(inp), VbStrConv.ProperCase)
            End If

            .Item("КСам17").Range.Text = ДатРожд & " г.р."
            .Item("КСам18").Range.Text = arrtbox("TextBox12") & " " & arrtbox("TextBox7")
            .Item("КСам19").Range.Text = arrtbox("TextBox8")
            .Item("КСам20").Range.Text = arrtmask("MaskedTextBox1")
            .Item("КСам21").Range.Text = Trim(arrtbox("TextBox9"))
            .Item("КСам22").Range.Text = arrtbox("TextBox21")
            .Item("КСам23").Range.Text = arrtbox("TextBox1") & " " & CorName & CorOtch
            .Item("КСам26").Range.Text = ПоСовмИлиОсн
            .Item("КСам27").Range.Text = combx10 & " " & Склонение(combx10)
            .Item("КСам28").Range.Text = combx12
            .Item("КСам29").Range.Text = arrtbox("TextBox50")
            .Item("КСам30").Range.Text = arrtbox("TextBox49")
            .Item("КСам31").Range.Text = ФИОРукРодПад
            .Item("КСам32").Range.Text = ФИОКор
            .Item("КСам33").Range.Text = ФИОКор
            .Item("КСам34").Range.Text = Strings.LCase(ДолжРуковВинПад)




        End With
        If Not IO.Directory.Exists(OnePath & Клиент & "\Контракт\" & Год) Then
            IO.Directory.CreateDirectory(OnePath & Клиент & "\Контракт\" & Год)
        End If

        oWordDoc.SaveAs2("C:\Users\Public\Documents\Рик\" & txtbx38 & " " & Заявление(9) & " (контракт)" & ".docx",,,,,, False)
        'oWordDoc.SaveAs2(OnePath & Клиент & "\Контракт\" & Год & "\" & Me.TextBox38.Text & " " & Заявление(9) & " (контракт)" & ".doc",,,,,, False)
        Try
            IO.File.Copy("C:\Users\Public\Documents\Рик\" & txtbx38 & " " & Заявление(9) & " (контракт)" & ".docx", OnePath & Клиент & "\Контракт\" & Год & "\" & txtbx38 & " " & Заявление(9) & " (контракт)" & ".docx")
        Catch ex As Exception
            'If MessageBox.Show("Контракт с сотрудником " & Заявление(9) & " существует. Заменить старый документ новым?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then
            Try
                IO.File.Delete(OnePath & Клиент & "\Контракт\" & Год & "\" & txtbx38 & " " & Заявление(9) & " (контракт)" & ".docx")
            Catch ex1 As Exception
                MessageBox.Show("Закройте файл!", Рик)
            End Try
            IO.File.Copy("C:\Users\Public\Documents\Рик\" & txtbx38 & " " & Заявление(9) & " (контракт)" & ".docx", OnePath & Клиент & "\Контракт\" & Год & "\" & txtbx38 & " " & Заявление(9) & " (контракт)" & ".docx")
            'End If
        End Try
        СохрАмасейл = OnePath & Клиент & "\Контракт\" & Год & "\" & txtbx38 & " " & Заявление(9) & " (контракт)" & ".docx"

        oWordDoc.Close(True)
        oWord.Quit(True)

        IO.File.Delete("C:\Users\Public\Documents\Рик\Контракт Амасейлс.docx")
        IO.File.Delete("C:\Users\Public\Documents\Рик\" & txtbx38 & " " & Заявление(9) & " (контракт)" & ".docx")
    End Sub
    Private Sub ОбнДогПодр()

        'Чист()
        'StrSql = "SELECT КодСотрудники FROM Сотрудники WHERE Фамилия='" & Me.TextBox1.Text & "' AND Имя='" & Me.TextBox2.Text & "' AND Отчество='" & Me.TextBox3.Text & "'"
        'ds = Selects(StrSql)
        Dim idСотруд As Integer
        Try
            idСотруд = CType(Label96.Text, Integer)
        Catch ex As Exception
            MessageBox.Show("Вы пытаетесь изменить сотрудника! Но его еще нет в базе", Рик)
            Exit Sub
        End Try

        Dim list As New Dictionary(Of String, Object)
        list.Add("@КодСотрудники", idСотруд)

        Updates(stroka:="UPDATE Сотрудники SET Сотрудники.Фамилия='" & Trim(TextBox1.Text) & "', Сотрудники.Имя='" & Trim(TextBox2.Text) & "', Сотрудники.Отчество='" & Trim(TextBox3.Text) & "', 
Сотрудники.ФамилияРодПад='" & Trim(TextBox6.Text) & "', Сотрудники.ИмяРодПад='" & Trim(TextBox5.Text) & "', Сотрудники.ОтчествоРодПад='" & Trim(TextBox4.Text) & "', 
Сотрудники.ПаспортСерия='" & TextBox12.Text & "', Сотрудники.ПаспортНомер='" & TextBox7.Text & "', Сотрудники.ПаспортКогдаВыдан='" & MaskedTextBox1.Text & "',
Сотрудники.ДоКакогоДейств='" & MaskedTextBox2.Text & "', Сотрудники.ПаспортКемВыдан='" & TextBox9.Text & "', Сотрудники.ИДНомер='" & TextBox8.Text & "',
Сотрудники.Регистрация='" & TextBox21.Text & "', Сотрудники.МестоПрожив='" & TextBox20.Text & "', Сотрудники.КонтТелГор='" & TextBox37.Text & "',
Сотрудники.КонтТелефон='" & MaskedTextBox10.Text & "', Сотрудники.СтраховойПолис='" & TextBox45.Text & "', Сотрудники.ФамилияДляЗаявления='" & Trim(TextBox34.Text) & "',
Сотрудники.ИмяДляЗаявления='" & Trim(TextBox11.Text) & "', Сотрудники.ОтчествоДляЗаявления='" & Trim(TextBox10.Text) & "', Сотрудники.Пол='" & ComboBox28.Text & "',
ФИОСборное='" & Trim(TextBox1.Text) & " " & Trim(TextBox2.Text) & " " & Trim(TextBox3.Text) & "', ФИОРодПод='" & Trim(TextBox6.Text) & " " & Trim(TextBox5.Text) & " " & Trim(TextBox4.Text) & "'
        WHERE Сотрудники.КодСотрудники=@КодСотрудники", list, "Сотрудники")






        If TextBox61.Text = "" Then
            Dim dog As String
            If ДогПодрНомНовы = 0 Then
                УдалСтар(idСотруд, ДогПодНомерСтар)
                dog = ДогПодНомерСтар
            Else
                dog = ДПодНом
            End If
            For i As Integer = 0 To ДогПодрВыпРаб.Count - 1
                Dim StrSql As String = "INSERT INTO ДогПодряда(ID,НомерДогПодр,ДатаДогПодр,Должность,ДатаНачала,ДатаОконч,СтоимРуб1,СтоимКоп1,ОбъекОбщепита,Примечание,ВыпРаб1,ВидИзм)
            VALUES(" & idСотруд & ",'" & dog & "','" & MaskedTextBox6.Text & "','" & ComboBox22.Text & "','" & Me.MaskedTextBox7.Text & "',
            '" & Me.MaskedTextBox8.Text & "','" & ДогПодрВыпРабСтР(i) & "','" & ДогПодрВыпРабСтК(i) & "','" & Me.ComboBox25.Text & "','" & Примечани & "','" & ДогПодрВыпРаб(i) & "','" & ДогПодрВыпРабСтОб(i) & "')"
                Updates(StrSql)
            Next
        Else
            Dim strsql2 As String = "SELECT Код FROM ДогПодряда WHERE ID=" & idСотруд & ""
            Dim ds5 As DataTable = Selects(strsql2)
            If errds = 1 Then
                Dim StrSql7 As String = "INSERT INTO ДогПодряда(НомерДогПодр,ДатаДогПодр,Должность,ДатаНачала,ДатаОконч,СтоимЧасаРуб,СтоимЧасаКоп,ОбъекОбщепита,Примечание,ID)
VALUES('" & ДПодНом & "', '" & Me.MaskedTextBox6.Text & "','" & Me.ComboBox22.Text & "','" & Me.MaskedTextBox7.Text & "','" & Me.MaskedTextBox8.Text & "',
'" & Me.TextBox61.Text & "','" & Me.TextBox62.Text & "','" & Me.ComboBox25.Text & "','" & Примечани & "'," & idСотруд & ")"
                Updates(StrSql7)
            Else
                Dim StrSql1 As String = "UPDATE ДогПодряда SET ДатаДогПодр ='" & Me.MaskedTextBox6.Text & "', Должность='" & Me.ComboBox22.Text & "',ДатаНачала='" & Me.MaskedTextBox7.Text & "',
ДатаОконч='" & Me.MaskedTextBox8.Text & "',СтоимЧасаРуб='" & Me.TextBox61.Text & "',СтоимЧасаКоп='" & Me.TextBox62.Text & "',ОбъекОбщепита='" & Me.ComboBox25.Text & "', Примечание='" & Примечани & "'
WHERE ДогПодряда.ID=" & idСотруд & " AND НомерДогПодр='" & ДПодНом & "'"
                Updates(StrSql1)
            End If
        End If




    End Sub
    Private Sub УдалСтар(ByVal d As Integer, ByVal s As String)
        Dim strsql As String = "delete FROM ДогПодряда WHERE ID=" & d & " AND НомерДогПодр='" & s & "'"
        Updates(strsql)

    End Sub
    Private Sub ComboBox9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox9.SelectedIndexChanged

        ComboBox7.Text = ""
        Должность = ""
        Должность = ComboBox9.Text

        Dim list As New Dictionary(Of String, Object)()        '
        list.Add("@Клиент", Клиент)
        list.Add("@Отделы", Отдел)
        list.Add("@Должность", Должность)
        'list.Add("@Разряд", combx7)   

        Dim ds1 = Selects(StrSql:="SELECT ШтСвод.Разряд
FROM ШтОтделы INNER JOIN ШтСвод ON ШтОтделы.Код = ШтСвод.Отдел
WHERE ШтОтделы.Отделы=@Отделы AND ШтСвод.Должность=@Должность AND ШтОтделы.Клиент=@Клиент ORDER BY ШтСвод.Разряд", list)

        Dim bvn As Integer = ds1.Rows.Count

        If bvn = 1 And ds1.Rows(0).Item(0).ToString = "" Then
            Dim ds2 = Selects(StrSql:="SELECT ШтОтделы.Отделы, ШтСвод.Должность, ШтСвод.Разряд, ШтСвод.ТарифнаяСтавка,
ШтСвод.ПовышениеПроц, ШтСвод.ТарСтПослеИспСрока, ПовПроцПослеИспСрока
FROM ШтОтделы INNER JOIN ШтСвод ON ШтОтделы.Код = ШтСвод.Отдел
WHERE ШтОтделы.Отделы=@Отделы AND ШтСвод.Должность=@Должность AND ШтОтделы.Клиент=@Клиент", list)


            Dim ds32 As DataTable = ПроверкаИзмененияТарифнойСтавки(ds2, MaskedTextBox3.Text)

            СвертывРазр(ds32)
            Отделы = ds32.Rows(0).Item(0).ToString() 'это сам добавил
            ТарифнаяСт = ds32.Rows(0).Item(3).ToString() 'это сам добавил
            ПовышениеПроц = ds32.Rows(0).Item(4).ToString() 'это сам добавил

            ПропОклад() 'оклад прописью

            Exit Sub
        ElseIf (bvn >= 1 And ds1.Rows(0).Item(0).ToString <> "") Or (bvn > 1 And ds1.Rows(0).Item(0).ToString = "") Then
            Очистка()

            'If IsDBNull(ПроверкаИзмененияТарифнойСтавки()) = True Then
            Dim ds3 = Selects(StrSql:="SELECT ШтОтделы.Отделы, ШтСвод.Должность, ШтСвод.Разряд, ШтСвод.ТарифнаяСтавка,
ШтСвод.ПовышениеПроц, ШтСвод.ТарСтПослеИспСрока, ПовПроцПослеИспСрока
FROM ШтОтделы INNER JOIN ШтСвод ON ШтОтделы.Код = ШтСвод.Отдел
WHERE ШтОтделы.Отделы=@Отделы AND ШтСвод.Должность=@Должность AND ШтОтделы.Клиент=@Клиент", list)

            Отделы = ds3.Rows(0).Item(0).ToString
            ТарифнаяСт = ds3.Rows(0).Item(3).ToString
            ПовышениеПроц = ds3.Rows(0).Item(4).ToString()

            Me.ComboBox7.AutoCompleteCustomSource.Clear()
            Me.ComboBox7.Items.Clear()
            For Each r As DataRow In ds1.Rows
                Me.ComboBox7.AutoCompleteCustomSource.Add(r.Item(0).ToString())
                Me.ComboBox7.Items.Add(r(0).ToString)
            Next

            Label79.ForeColor = Color.Red
            Label79.Text = "NO"

            If ComboBox7.Enabled = False Then
                ComboBox7.Enabled = True
            End If
        End If



    End Sub
    Private Function ПроверкаИзмененияТарифнойСтавки(ByVal dsin As DataTable, ByVal datex As String) As DataTable
        Dim list As New Dictionary(Of String, Object)()        '
        list.Add("@Клиент", Клиент)
        list.Add("@Отделы", Отдел)
        list.Add("@Должность", Должность)
        'list.Add("@Разряд", combx7)



        Dim ds As DataTable = Selects(StrSql:="SELECT ШтСвод.КодШтСвод
FROM ШтОтделы INNER JOIN ШтСвод ON ШтОтделы.Код = ШтСвод.Отдел
WHERE ШтОтделы.Отделы=@Отделы AND ШтСвод.Должность=@Должность AND ШтОтделы.Клиент=@Клиент", list)

        Dim ds1 As DataTable = Selects(StrSql:="SELECT * FROM ШтСводИзмСтавка
WHERE IDКодШтСвод=" & CType(ds.Rows(0).Item(0).ToString, Integer) & "")
        Dim DateEx As Date = CDate(datex)
        If Not errds = 1 Then
            ds1.DefaultView.Sort = "Дата DESC"
            ds1 = ds1.DefaultView.ToTable()

            For x As Integer = 0 To ds1.Rows.Count - 1
                If DateEx >= ds1.Rows(x).Item(2) Then
                    dsin.Rows(0).Item(3) = ds1.Rows(x).Item(3)
                    Return dsin
                End If
            Next
        End If
        Return dsin
    End Function
    Private Function ДолжПриИзменСотр()

        Dim list As New Dictionary(Of String, Object)()        '
        list.Add("@НазвОрганиз", ComboBox1.Text)
        list.Add("@ФИОСборное", ComboBox19.Text)


        StrSql = ""
        Dim ds = Selects(StrSql:="SELECT Штатное.Должность
FROM Сотрудники INNER JOIN Штатное ON Сотрудники.КодСотрудники = Штатное.ИДСотр
WHERE Сотрудники.НазвОрганиз=@НазвОрганиз AND Сотрудники.ФИОСборное=@ФИОСборное", list)

        ДолжПриИзменСотр = ds.Rows(0).Item(0).ToString
        Return Replace(ДолжПриИзменСотр, ".", "")


    End Function
    Private Sub ПрверкаАсинхрПотоков(ByVal d As Integer)
        Dim dx As Task
        For i As Integer = 0 To TskList.Count - 1
            If TskList(i).Id = d Then
                dx = TskList(i)
            End If
        Next

        'Dim bn As Task = Tasks.
        TskList.Remove(dx)
        Dim fc As Integer = TskList.Count
        Dim arrTask(fc - 1) As Task
        For i As Integer = 0 To TskList.Count - 1
            arrTask(i) = TskList.Item(i)
        Next
        Task.WaitAll(arrTask)


        'If d.Equals(f) Then
        '    'Dim arrTask() As Task = {f, f1, f2, f3}



        '    If f1.Status = TaskStatus.Running Then
        '        f1.Wait()
        '    End If
        '    If f2.Status = TaskStatus.Running Then
        '        f2.Wait()
        '    End If
        '    If f3.Status = TaskStatus.Running Then
        '        f3.Wait()
        '    End If
        'End If

        'If f.Status = TaskStatus.Running Then
        '        f.Wait()
        '    End If
        '    If f2.Status = TaskStatus.Running Then
        '        f2.Wait()
        '    End If
        '    If f3.Status = TaskStatus.Running Then
        '        f3.Wait()
        '    End If
        Dim tsk As Task = New Task(AddressOf УдаляемФонПроцессы)
        tsk.Start()
        tsk.Wait()
    End Sub

    Private Sub КонтрРазряд()

        ДокКонтрПерем = ""
        If CheckBox5.Checked = False Then

            If combx7 = "-" Then
                ДокКонтрПерем = Strings.LCase(ДолжСОконч)
            ElseIf combx7 = "1" Or combx7 = "2" Or combx7 = "3" Or combx7 = "4" Or combx7 = "5" Or combx7 = "6" Then
                ДокКонтрПерем = LCase(ДолжСОконч) & " " & combx7 & " разряда"
            Else
                ДокКонтрПерем = Strings.LCase(ДолжСОконч)
            End If
        Else

            Dim row = dtShtatnoeAll.Select("ИДСотр=" & CType(Label96.Text, Integer) & "")
            Dim f As String = row(0).Item("Разряд").ToString

            If f = "1" Or f = "2" Or f = "3" Or f = "4" Or f = "5" Or f = "6" Then
                ДокКонтрПерем = Strings.LCase(ДолжСОконч) & " " & f & " разряда"
            Else
                ДокКонтрПерем = Strings.LCase(ДолжСОконч)
            End If

        End If

    End Sub
    Private Sub Combx15Контракт()
        К33 = ""
        К34 = ""
        К35 = ""
        К36 = ""
        К37 = ""


        Select Case combx15
            Case "График"
                К33 = "согласно графику работ"
                К34 = "согласно графику работ"
                К35 = "согласно графику работ"
                Select Case CheckBox4.Checked
                    Case False
                        К36 = "Суббота, Воскресенье"
                    Case True
                        К36 = "согласно графику работ"
                        К37 = "11.5. работнику устанавливается суммированный учет рабочего времени с учетным периодом - год."
                End Select

            Case "ПВТР"
                К33 = "согласно правил внутреннего трудового распорядка"
                К34 = "согласно правил внутреннего трудового распорядка"
                К35 = "согласно правил внутреннего трудового распорядка"

                Select Case CheckBox4.Checked
                    Case False
                        К36 = "согласно графику работ"
                    Case True
                        К36 = "согласно графику работ"
                        К37 = "11.5. работнику устанавливается суммированный учет рабочего времени с учетным периодом - год."
                End Select

            Case "Задать"
                К33 = combx12
                К34 = TextBox49.Text
                К35 = TextBox50.Text

                Select Case CheckBox4.Checked
                    Case False
                        К36 = "Суббота, Воскресенье"
                    Case True
                        К36 = "согласно графику работ"
                        К37 = "11.5. работнику устанавливается суммированный учет рабочего времени с учетным периодом - год."
                End Select
        End Select
    End Sub

    Private Async Sub ДокКонтракт()
        'Await Task.Delay(0)
        'KillProc()
        'Me.Cursor = Cursors.WaitCursor
        'Dim s As New Thread(AddressOf КонтрРазряд) 'поток1

        'Dim combx15Th As New Thread(AddressOf Combx15Контракт) 'поток 2
        'combx15Th.Start()


        's.Start()


        Dim oWord2 As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc2 As Microsoft.Office.Interop.Word.Document
        oWord2 = CreateObject("Word.Application")
        oWord2.Visible = False

        ВыгрузкаФайловНаЛокалыныйКомп(FTPStringAllDOC & "Kontrakt.doc", firthtPath & "\Kontrakt.doc")

        oWordDoc2 = oWord2.Documents.Add(firthtPath & "\Kontrakt.doc")

        With oWordDoc2.Bookmarks
            .Item("К0").Range.Text = Контракт(0)
            .Item("К1").Range.Text = mskbx3
            .Item("К2").Range.Text = Заявление(9)
            .Item("К3").Range.Text = Заявление(10)
            .Item("К4").Range.Text = Заявление(11)
            .Item("К5").Range.Text = Заявление(1)
            .Item("К6").Range.Text = Заявление(2)
            .Item("К7").Range.Text = Заявление(3)

            Await Task.Run(Sub() КонтрРазряд())
            .Item("К8").Range.Text = ДокКонтрПерем
            'If combx7 = "-" Then
            '    .Item("К8").Range.Text = Strings.LCase(ДолжСОконч)
            'ElseIf combx7 = "1" Or combx7 = "2" Or combx7 = "3" Or combx7 = "4" Or combx7 = "5" Or combx7 = "6" Then
            '    .Item("К8").Range.Text = LCase(ДолжСОконч) & " " & combx7 & " разряда"
            'Else
            '    .Item("К8").Range.Text = Strings.LCase(ДолжСОконч)
            'End If
            .Item("К9").Range.Text = Заявление(8) & " " & СтавкаНов
            .Item("К10").Range.Text = Контракт(2) & " (" & СрКонтПроп & ") " & СклонГод
            .Item("К11").Range.Text = Контракт(1)
            .Item("К12").Range.Text = Контракт(3)
            .Item("К13").Range.Text = Заявление(9)
            .Item("К14").Range.Text = Заявление(10)
            .Item("К15").Range.Text = Заявление(11)
            .Item("К16").Range.Text = Заявление(4)
            .Item("К17").Range.Text = Контракт(5)
            .Item("К18").Range.Text = Контракт(6)
            .Item("К19").Range.Text = Контракт(8)
            .Item("К20").Range.Text = Контракт(7)
            .Item("К21").Range.Text = Контракт(9)
            .Item("К22").Range.Text = Заявление(9)
            .Item("К23").Range.Text = CorName
            .Item("К24").Range.Text = CorOtch
            .Item("К25").Range.Text = Заявление(9) & " " & CorName & CorOtch
            .Item("К26").Range.Text = Контракт(4) & "," & txtbx44
            .Item("К27").Range.Text = Контракт(10)
            .Item("К28").Range.Text = ПоСовмИлиОсн
            'If TextBox46.InvokeRequired Then
            '    Me.Invoke(New txtbx46(AddressOf ДокКонтракт))
            'Else
            .Item("К29").Range.Text = txtbxD46
            'End If
            .Item("К30").Range.Text = РДОрубли
            .Item("К31").Range.Text = РДОкопейки
            .Item("К32").Range.Text = txtbx47
            Select Case combx8
                Case "Руководители"
                    .Item("К38").Range.Text = "должностной инструкции"
                Case "Специалисты"
                    .Item("К38").Range.Text = "должностной инструкции"
            End Select

            Await Task.Run(Sub() Combx15Контракт())
            'combx15Th.Join()
            .Item("К33").Range.Text = К33
            .Item("К34").Range.Text = К34
            .Item("К35").Range.Text = К35
            .Item("К36").Range.Text = К36
            .Item("К37").Range.Text = К37


            .Item("К39").Range.Text = ФормаСобстПолн

            If ФормаСобстПолн = "Индивидуальный предприниматель" Then
                .Item("К40").Range.Text = Клиент
                .Item("К41").Range.Text = ""
            Else
                .Item("К40").Range.Text = " «" & Клиент & "» "
                .Item("К41").Range.Text = ДолжРуковВинПад
            End If

            .Item("К42").Range.Text = ФИОРукРодПад

            If Not combx1 = "Итал Гэлэри Плюс" Then
                .Item("К43").Range.Text = ОснованиеДейств
            Else
                .Item("К51").Range.Text = ""
            End If
            .Item("К44").Range.Text = МестоРаб
            .Item("К45").Range.Text = ФИОКор
            .Item("К46").Range.Text = СборноеРеквПолн
            .Item("К47").Range.Text = Year(Now).ToString
            .Item("К48").Range.Text = TextBox40.Text
            If TextBox56.Text = "" Or TextBox56.Text = "НЕТ" Then
                .Item("К49").Range.Text = ""
            Else
                .Item("К49").Range.Text = "и " & TextBox56.Text & "-го (аванс) "
            End If
            'If ComboBox10.Text = "1.0" Then
            .Item("К50").Range.Text = "1 ставка"
            'Else
            '    .Item("К50").Range.Text = ComboBox10.Text & " ставки"
            'End If
            Select Case combx28
                Case "М"
                    .Item("К52").Range.Text = "ним"
                Case "Ж"
                    .Item("К52").Range.Text = "ней"
            End Select

        End With


        Dim dirstring As String = Клиент & "/Контракт/" & Now.Year & "/" 'место сохранения файла
        dirstring = СозданиепапкиНаСервере(dirstring) 'полный путь на сервер(кроме имени и разрешения файла)


        Dim put, Name As String
        Name = txtbx38 & " " & Заявление(9) & " (контракт)" & " - " & IDso & ".doc"
        put = PathVremyanka & Name 'место в корне программы

        ВыборкаИзагрНаСервер(dirstring, Name, "Прием-Контракт")

        'Dim b = dtSotrudnikiAll.Select("ФИОСборное='" & combx19 & "'") 'выбираем данные по сотруднику and НазвОрганиз='" & combx1 & " & txt1 & "" & txt2 & "" & txt3 &
        'Dim kd As Integer = CType(b(0).Item("КодСотрудники").ToString, Integer) 'находим ИД сотрудника
        'ЗагрВБазуПутиДоковAsync(kd, dirstring, Name, "Прием-Контракт") 'заполняем данные путей и назв файла

        oWordDoc2.SaveAs2(put,,,,,, False)


        oWordDoc2.Close(True)
        oWord2.Quit(True)
        СохрКонтрFTP.AddRange(New String() {dirstring, Name})
        dirstring += Name

        ЗагрНаСерверИУдаление(put, dirstring, put)


        ВременнаяПапкаУдалениеФайла(firthtPath & "\Kontrakt.doc")




        'If Not IO.Directory.Exists(OnePath & Клиент & "\Контракт\" & Год) Then
        '    IO.Directory.CreateDirectory(OnePath & Клиент & "\Контракт\" & Год)
        'End If
        'Try
        '    oWordDoc2.SaveAs2(OnePath & Клиент & "\Контракт\" & Год & "\" & txtbx38 & " " & Заявление(9) & " (контракт)" & " - " & IDso & ".doc",,,,,, False)
        'Catch ex As Exception
        '    If ex.Message.Contains("уже существует") Then
        '        IO.File.Delete(OnePath & Клиент & "\Контракт\" & Год & "\" & txtbx38 & " " & Заявление(9) & " (контракт)" & " - " & IDso & ".doc")
        '        oWordDoc2.SaveAs2(OnePath & Клиент & "\Контракт\" & Год & "\" & txtbx38 & " " & Заявление(9) & " (контракт)" & " - " & IDso & ".doc",,,,,, False)
        '    End If
        '    oWordDoc2.SaveAs2("C:\Users\Public\Documents\Рик\" & txtbx38 & " " & Заявление(9) & " (контракт)" & " - " & IDso & ".doc",,,,,, False)
        '    IO.File.Copy("C:\Users\Public\Documents\Рик\" & txtbx38 & " " & Заявление(9) & " (контракт)" & " - " & IDso & ".doc", OnePath & Клиент & "\Контракт\" & Год & "\" & txtbx38 & " " & Заявление(9) & " (контракт)" & " - " & IDso & ".doc")
        'End Try

        'oWordDoc2.Close(True)
        'oWord2.Quit(True)

        'УдалениеСтарыхФайловВПапкеРик("C:\Users\Public\Documents\Рик\" & txtbx38 & " " & Заявление(9) & " (контракт)" & " - " & IDso & ".doc")
        'УдалениеСтарыхФайловВПапкеРик("C:\Users\Public\Documents\Рик\Kontrakt.doc")

        'MessageBox.Show("Контракт закончен")
    End Sub
    Public Sub УдалениеСтарыхФайловВПапкеРик(ByVal d As String)
        If IO.File.Exists(d) Then
            IO.File.Delete(d)
        End If
    End Sub
    Private Async Sub ДокПриказ()



        Dim oWord3 As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc3 As Microsoft.Office.Interop.Word.Document
        oWord3 = CreateObject("Word.Application")
        oWord3.Visible = False

        ВыгрузкаФайловНаЛокалыныйКомп(FTPStringAllDOC & "Prikaz.doc", firthtPath & "\Prikaz.doc")

        oWordDoc3 = oWord3.Documents.Add(firthtPath & "\Prikaz.doc")

        With oWordDoc3.Bookmarks
            .Item("П1").Range.Text = Приказ(5)
            .Item("П2").Range.Text = НПриказа
            .Item("П3").Range.Text = txtbx6
            .Item("П4").Range.Text = CorName
            .Item("П5").Range.Text = CorOtch
            .Item("П6").Range.Text = txtbx6
            .Item("П7").Range.Text = Приказ(3)
            .Item("П8").Range.Text = Приказ(4)
            .Item("П9").Range.Text = Strings.LCase(ДолжСОконч) & ДолжнИразрядДокЗаявление()
            .Item("П10").Range.Text = Приказ(6)
            .Item("П11").Range.Text = Ставка
            .Item("П12").Range.Text = СтавкаНов
            .Item("П13").Range.Text = СрокКонтр
            .Item("П14").Range.Text = СклонГод
            .Item("П15").Range.Text = Приказ(6)
            .Item("П16").Range.Text = Приказ(7)
            .Item("П17").Range.Text = Приказ(2)
            .Item("П18").Range.Text = CorName
            .Item("П19").Range.Text = CorOtch
            .Item("П20").Range.Text = Приказ(8)
            .Item("П21").Range.Text = Приказ(5)
            .Item("П22").Range.Text = Приказ(9)
            .Item("П23").Range.Text = CorName
            .Item("П24").Range.Text = CorOtch
            .Item("П25").Range.Text = ФормаСобстПолн

            If ФормаСобстПолн = "Индивидуальный предприниматель" Then
                .Item("П26").Range.Text = Клиент
            Else
                .Item("П26").Range.Text = " «" & Клиент & "» "
            End If

            .Item("П27").Range.Text = ЮрАдрес
            .Item("П28").Range.Text = УНП
            .Item("П29").Range.Text = РасСчет
            .Item("П30").Range.Text = АдресБанка
            .Item("П31").Range.Text = БИК
            .Item("П33").Range.Text = ЭлАдрес
            .Item("П34").Range.Text = КонтТелефон
            .Item("П35").Range.Text = МестоРаб

            If ДолжРуков = "Индивидуальный предприниматель" Then
                .Item("П36").Range.Text = ДолжРуков
                .Item("П37").Range.Text = ""
            Else
                .Item("П36").Range.Text = ДолжРуков & " " & ФормаСобствКор
                .Item("П37").Range.Text = " «" & Клиент & "» "
            End If


            .Item("П38").Range.Text = ФИОКор
            .Item("П39").Range.Text = ПоСовмПриказ

        End With


        Dim dirstring As String = Клиент & "/Приказ/" & Now.Year & "/" 'место сохранения файла
        dirstring = СозданиепапкиНаСервере(dirstring) 'полный путь на сервер(кроме имени и разрешения файла)


        Dim put, Name As String
        Name = НПриказа & " прием " & Приказ(9) & " от " & mskbx3 & " (приказ)" & " - " & IDso & " .doc"
        put = PathVremyanka & Name 'место в корне программы

        ВыборкаИзагрНаСервер(dirstring, Name, "Прием-Приказ")

        'Dim b = dtSotrudnikiAll.Select("ФИОСборное='" & combx19 & "'") 'выбираем данные по сотруднику
        'Dim kd As Integer = CType(b(0).Item("КодСотрудники").ToString, Integer) 'находим ИД сотрудника
        'ЗагрВБазуПутиДоковAsync(kd, dirstring, Name, "Прием-Приказ") 'заполняем данные путей и назв файла

        oWordDoc3.SaveAs2(put,,,,,, False)


        oWordDoc3.Close(True)
        oWord3.Quit(True)

        СохрПрикFTP.AddRange(New String() {dirstring, Name})
        dirstring += Name

        ЗагрНаСерверИУдаление(put, dirstring, put)

        ВременнаяПапкаУдалениеФайла(firthtPath & "\Prikaz.doc")

        'MessageBox.Show("Приказ закончен")
    End Sub

    Private Function ДолжнИразрядДокЗаявление()

        If CheckBox5.Checked = False Then
            If combx7 = "1" Or combx7 = "2" Or combx7 = "3" Or combx7 = "4" Or combx7 = "5" Or combx7 = "6" Then
                Return " " & combx7 & " разряда"
            End If
        Else
            Dim row = dtShtatnoeAll.Select("ИДСотр=" & CType(Label96.Text, Integer) & "")
            Dim f As String = row(0).Item("Разряд").ToString

            If f = "1" Or f = "2" Or f = "3" Or f = "4" Or f = "5" Or f = "6" Then
                Return " " & f & " разряда"
            End If
        End If
        Return ""
    End Function
    Private Async Sub ДокЗаявление()


        Dim oWord1 As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc1 As Microsoft.Office.Interop.Word.Document
        oWord1 = CreateObject("Word.Application")


        oWord1.Visible = False


        ВыгрузкаФайловНаЛокалыныйКомп(FTPStringAllDOC & "Zayavlenie.doc", firthtPath & "\Zayavlenie.doc")

        oWordDoc1 = oWord1.Documents.Add(firthtPath & "\Zayavlenie.doc")


        With oWordDoc1.Bookmarks
            .Item("ЗАКЛ0").Range.Text = Заявление(0)
            .Item("ЗАКЛ1").Range.Text = Trim(Приказ(2))
            .Item("ЗАКЛ2").Range.Text = Trim(Заявление(13))
            .Item("ЗАКЛ3").Range.Text = Trim(Заявление(14))
            .Item("ЗАКЛ4").Range.Text = Заявление(4)
            .Item("ЗАКЛ5").Range.Text = Заявление(5)
            .Item("ЗАКЛ6").Range.Text = LCase(ДолжСОконч) & ДолжнИразрядДокЗаявление()
            .Item("ЗАКЛ7").Range.Text = arrtmask("MaskedTextBox4") 'Заявление(7)MaskedTextBox4.Text
            .Item("ЗАКЛ8").Range.Text = Заявление(8)
            .Item("ЗАКЛ9").Range.Text = CorName
            .Item("ЗАКЛ10").Range.Text = CorOtch
            .Item("ЗАКЛ11").Range.Text = Заявление(9)
            .Item("ЗАКЛ12").Range.Text = СтавкаНов
            .Item("ЗАКЛ13").Range.Text = ДолжРуковРодПад

            If ДолжРуковРодПад = "Индивидуальному предпринимателю" Or ФормаСобствКор = "ИП" Then
                .Item("ЗАКЛ14").Range.Text = ""
                .Item("ЗАКЛ18").Range.Text = "по месту нахождения"
            Else
                .Item("ЗАКЛ14").Range.Text = ФормаСобствКор & " """ & Клиент & """ "
                .Item("ЗАКЛ18").Range.Text = "в"
            End If

            .Item("ЗАКЛ15").Range.Text = ФИОКорРукДат
            .Item("ЗАКЛ16").Range.Text = МестоРаб
            .Item("ЗАКЛ17").Range.Text = Заявление(0)


        End With


        Dim dirstring As String = Клиент & "/Заявление/" & Now.Year & "/" 'место сохранения файла

        dirstring = СозданиепапкиНаСервере(dirstring) 'полный путь на сервер(кроме имени и разрешения файла)


        Dim put, Name As String
        Name = Заявление(9) & " (заявление)" & " - " & IDso & ".doc"
        put = PathVremyanka & Name 'место в корне программы


        ВыборкаИзагрНаСервер(dirstring, Name, "Прием-Зявление")


        oWordDoc1.SaveAs2(put,,,,,, False)
        oWordDoc1.Close(True)
        СохрЗакFTP.AddRange(New String() {dirstring, Name})
        oWord1.Quit(True)
        dirstring += Name
        ЗагрНаСерверИУдаление(put, dirstring, put)



        ВременнаяПапкаУдалениеФайла(firthtPath & "\Zayavlenie.doc")


        'MessageBox.Show("Заявление окончен")

    End Sub


    Private Sub ДокИнструкц()
        Dim hk As DataTable
        Dim list As New Dictionary(Of String, Object)()        '
        list.Add("@Клиент", combx1)
        list.Add("@Отделы", combx8)
        list.Add("@Должность", combx9)
        list.Add("@Разряд", combx7)

        If Label96.Text = "№" Or Label96.Text = "" Then
            list.Add("@ID", 0)
        Else
            Try
                list.Add("@ID", CType(Label96.Text, Integer))
            Catch ex As Exception
                list.Add("@ID", 0)
            End Try

        End If


        ИнстрFTP.Clear()

        Try
            If Not ПровИнстр = 1 Then
                'Формируем инструкцию 
                If CheckBox23.Checked = False Then
                    Dim dg = Selects(StrSql:="Select ШтСвод.НомерДолжИнстр FROM ШтОтделы INNER JOIN ШтСвод On ШтОтделы.Код = ШтСвод.Отдел
        WHERE ШтОтделы.Клиент=@Клиент AND ШтОтделы.Отделы=@Отделы AND ШтСвод.Должность=@Должность AND ШтСвод.Разряд=@Разряд AND ШтСвод.ДолжИнструкция='True'", list)
                    If errds = 0 Then
                        ИнстрП = dg.Rows(0).Item(0).ToString
                        ИнстрFTP.AddRange(New String() {FTPString & combx1 & "/Должностные инструкции/", dg.Rows(0).Item(0).ToString & ".doc"})
                    End If

                Else
                    If CheckBox26.Checked = True Then
                        Dim dg = Selects(StrSql:="SELECT ШтСвод.НомерДолжИнстр FROM ШтОтделы INNER JOIN ШтСвод ON ШтОтделы.Код = ШтСвод.Отдел
WHERE ШтОтделы.Клиент=@Клиент AND ШтОтделы.Отделы=@Отделы AND ШтСвод.Должность=@Должность
AND ШтСвод.Разряд=@Разряд AND ШтСвод.ДолжИнструкция='True'", list)

                        If errds = 0 Then
                            ИнстрП = dg.Rows(0).Item(0).ToString
                            'Инстр = OnePath & combx1 & "\Должностные инструкции\" & dg.Rows(0).Item(0).ToString & ".doc"
                            ИнстрFTP.AddRange(New String() {FTPString & combx1 & "/Должностные инструкции/", dg.Rows(0).Item(0).ToString & ".doc"})
                        End If

                    Else
                        If Not hk Is Nothing Then hk.Clear()
                        hk = Selects(StrSql:="SELECT Отдел,Должность,Разряд FROM Штатное WHERE ИДСотр=@ID", list)

                        If Not hk Is Nothing Then
                            list.Add("@Отделы2", hk.Rows(0).Item(0).ToString)
                            list.Add("@Должность2", hk.Rows(0).Item(1).ToString)
                            list.Add("@Разряд2", hk.Rows(0).Item(2).ToString)
                        End If

                        If errds = 0 Then
                            Dim dg = Selects(StrSql:="SELECT ШтСвод.НомерДолжИнстр FROM ШтОтделы INNER JOIN ШтСвод ON ШтОтделы.Код = ШтСвод.Отдел
WHERE ШтОтделы.Клиент=@Клиент AND ШтОтделы.Отделы=@Отделы2 AND ШтСвод.Должность=@Должность2
AND ШтСвод.Разряд=@Разряд2 AND ШтСвод.ДолжИнструкция='True'", list)

                            Try
                                ИнстрП = dg.Rows(0).Item(0).ToString
                                ИнстрFTP.AddRange(New String() {FTPString & combx1 & "/Должностные инструкции/", dg.Rows(0).Item(0).ToString & ".doc"})
                            Catch ex As Exception

                            End Try

                            'Инстр = OnePath & combx1 & "\Должностные инструкции\" & dg.Rows(0).Item(0).ToString & ".doc"

                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        'MessageBox.Show("Инструкция закончена")
    End Sub
    Private Sub ДокПредварДанн()
        Try
            IDso = CType(Label96.Text, Integer)
        Catch ex As Exception
            IDso = IDsot1
        End Try

        ДолжРуковВинПад = ДобОконч(ДолжРуков)

        If Должность = "" And CheckBox5.Checked = True And CheckBox23.Checked = True Then 'если изменяем сотрудника и поле должность пустое то подтягиваем должность из базы
            ДолжСОконч = ДобОконч(ДолжПриИзменСотр())
        Else
            ДолжСОконч = ДобОконч(Replace(Должность, ".", ""))
        End If

        СтавкаНов = Склонение(Ставка) 'склонение ставки
        СклонГод = Склонение2(СрокКонтр) ' склонение год
        СрКонтПроп = ЧислПроп(ComboBox11.Text)

        If CheckBox2.Checked = True Then 'галочка по осн или по совместительству
            ПоСовмИлиОсн = "совместительству"
            ПоСовмПриказ = "по совместительству"
        Else
            ПоСовмИлиОсн = "основной работе"
            ПоСовмПриказ = "основное место работы"
        End If
        ДолжРуковРодПад = ДолжРодПадежФункц(ДолжРуков)

    End Sub

    Private Async Sub Доки(ByVal all As String)
        Await Task.Delay(0)
        Me.Cursor = Cursors.WaitCursor
        Erase TskArr
        ДокПредварДанн()


        Select Case all
            Case "общ"


                f = New Task(AddressOf ДокЗаявление) 'асинхрон
                f1 = New Task(AddressOf ДокКонтракт) 'асинхрон
                f2 = New Task(AddressOf ДокПриказ) 'асинхрон
                f3 = New Task(AddressOf ДокИнструкц) 'асинхрон


                TskArr = {f, f1, f2, f3}
                For Each r As Task In TskArr
                    r.Start()
                Next
                'TskList.Exists(TskList)
                Try
                    TskList.Clear()
                    TskList = {f, f1, f2, f3}.ToList
                Catch ex As Exception
                    TskList = {f, f1, f2, f3}.ToList
                End Try

            Case "амасейлс"

                f = New Task(AddressOf ДокЗаявление) 'асинхрон
                f1 = New Task(AddressOf ДокиАмасейл) 'асинхрон
                f2 = New Task(AddressOf ДокИнструкц) 'асинхрон
                f3 = New Task(AddressOf ДокПриказ) 'асинхрон


                TskArr = {f, f1, f2, f3}
                For Each r As Task In TskArr
                    r.Start()
                Next

                Try
                    TskList.Clear()
                    TskList = {f, f1, f2, f3}.ToList
                Catch ex As Exception
                    TskList = {f, f1, f2, f3}.ToList
                End Try


            Case "ЛемеЛ"
                ДопЛемеЛ()

                f = New Task(AddressOf ДокЗаявление) 'асинхрон
                f1 = New Task(AddressOf ДокиЛемеЛ) 'асинхрон
                f2 = New Task(AddressOf ДокИнструкц) 'асинхрон
                f3 = New Task(AddressOf ДокПрикаЛемел) 'асинхрон


                TskArr = {f, f1, f2, f3}
                For Each r As Task In TskArr
                    r.Start()
                Next

                Try
                    TskList.Clear()
                    TskList = {f, f1, f2, f3}.ToList
                Catch ex As Exception
                    TskList = {f, f1, f2, f3}.ToList
                End Try
            Case "Пинфуд Сервис"

                f = New Task(AddressOf ДокЗаявление) 'асинхрон
                f1 = New Task(AddressOf ДокиПинфуд) 'асинхрон
                f2 = New Task(AddressOf ДокИнструкц) 'асинхрон
                f3 = New Task(AddressOf ДокПриказ) 'асинхрон


                TskArr = {f, f1, f2, f3}
                For Each r As Task In TskArr
                    r.Start()
                Next

                Try
                    TskList.Clear()
                    TskList = {f, f1, f2, f3}.ToList
                Catch ex As Exception
                    TskList = {f, f1, f2, f3}.ToList
                End Try





        End Select
        dtShtatnoe()
        Me.Cursor = Cursors.Default
    End Sub


    Function Проверка(ByVal a As String, ByVal IDPass As String) As VariantType

        Dim ds = Selects(StrSql:="Select КодСотрудники From Сотрудники 
WHERE ФИОСборное='" & a & "' and ИДНомер='" & IDPass & "' and НазвОрганиз='" & ComboBox1.Text & "'")

        Dim xv As String
        Try
            xv = ds.Rows(0).Item(0).ToString
        Catch ex As Exception
            Return 1
            Exit Function
        End Try
        Return 0

    End Function
    Private Function ПровДублСотр() As Integer

        Dim dhg As DataTable = Selects(StrSql:="SELECT Сотрудники.НазвОрганиз, Сотрудники.ФИОСборное, Сотрудники.ИДНомер, КарточкаСотрудника.ДатаУвольнения
FROM Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр
WHERE КарточкаСотрудника.ДатаУвольнения Is Null AND Сотрудники.НазвОрганиз='" & combx1 & "' AND Сотрудники.ФИОСборное='" & Trim(TextBox1.Text) & " " & Trim(TextBox2.Text) & " " & Trim(TextBox3.Text) & "'
AND Сотрудники.ИДНомер='" & TextBox8.Text & "'")

        If errds = 0 Then
            If MessageBox.Show("Данный сотрудник еще работает в этой компании и пока не уволен!" & vbCrLf & "Всё равно создать нового сотрудника?", Рик, MessageBoxButtons.YesNo) = DialogResult.No Then
                Return 1
            End If
        End If
        Return 0

    End Function


    Private Sub ДобавлНовогоСотрудника()


        Dim list2 As New Dictionary(Of String, Object)
        With list2
            .Add("@НазвОрганиз", a)
            .Add("@Фамилия", Trim(surName))
            .Add("@Имя", Trim(arrtbox("TextBox2")))
            .Add("@Отчество", Trim(arrtbox("TextBox3")))
            .Add("@ФамилияРодПад", Trim(arrtbox("TextBox6")))
            .Add("@ИмяРодПад", Trim(arrtbox("TextBox5")))
            .Add("@ОтчествоРодПад", Trim(arrtbox("TextBox4")))

            .Add("@ПаспортСерия", arrtbox("TextBox12"))
            .Add("@ПаспортНомер", arrtbox("TextBox7"))
            .Add("@ПаспортКогдаВыдан", arrtmask("MaskedTextBox1"))
            .Add("@ДоКакогоДейств", arrtmask("MaskedTextBox2"))
            .Add("@ПаспортКемВыдан", arrtbox("TextBox9"))
            .Add("@ИДНомер", arrtbox("TextBox8"))
            .Add("@Регистрация", arrtbox("TextBox21"))
            .Add("@МестоПрожив", arrtbox("TextBox20"))
            .Add("@КонтТелГор", arrtbox("TextBox37"))

            .Add("@КонтТелефон", arrtmask("MaskedTextBox10"))
            .Add("@СтраховойПолис", arrtbox("TextBox45"))
            .Add("@НаличеДогПодряда", "Нет")
            .Add("@ФамилияДляЗаявления", Trim(arrtbox("TextBox34")))
            .Add("@ИмяДляЗаявления", Trim(arrtbox("TextBox11")))
            .Add("@ОтчествоДляЗаявления", Trim(arrtbox("TextBox10")))
            .Add("@Пол", combx28)
            .Add("@ДатаРожд", "no")
            .Add("@Гражданин", "no")
            .Add("@ПровДатыКонтр", Trim(arrtmask("MaskedTextBox3")))
            If CheckBox1.Checked = True Then
                .Add("@Иностранец", "True")
            Else
                .Add("@Иностранец", "False")
            End If
            .Add("@ФИОСборное", Trim(surName) & " " & Trim(arrtbox("TextBox2")) & " " & Trim(arrtbox("TextBox3")))
            .Add("@ФИОРодПод", Trim(arrtbox("TextBox6")) & " " & Trim(arrtbox("TextBox5")) & " " & Trim(arrtbox("TextBox4")))
            .Add("@ТипОтношения", "(кт)")
        End With


        Dim Newid As Integer

        Newid = Updates(stroka:="INSERT INTO Сотрудники(НазвОрганиз, Фамилия, Имя, Отчество, ФамилияРодПад, ИмяРодПад, ОтчествоРодПад,
ПаспортСерия, ПаспортНомер, ПаспортКогдаВыдан, ДоКакогоДейств, ПаспортКемВыдан, ИДНомер, Регистрация, МестоПрожив, КонтТелГор, 
КонтТелефон, СтраховойПолис, НаличеДогПодряда,ФамилияДляЗаявления, ИмяДляЗаявления, ОтчествоДляЗаявления, Пол, ДатаРожд,  
Гражданин, ПровДатыКонтр, Иностранец, ФИОСборное, ФИОРодПод, ТипОтношения)

VALUES(@НазвОрганиз, @Фамилия, @Имя, @Отчество,@ФамилияРодПад, @ИмяРодПад,@ОтчествоРодПад,
@ПаспортСерия, @ПаспортНомер,@ПаспортКогдаВыдан, @ДоКакогоДейств, @ПаспортКемВыдан, @ИДНомер, @Регистрация, @МестоПрожив, @КонтТелГор,
@КонтТелефон, @СтраховойПолис, @НаличеДогПодряда,@ФамилияДляЗаявления, @ИмяДляЗаявления, @ОтчествоДляЗаявления, @Пол, @ДатаРожд,
@Гражданин, @ПровДатыКонтр, @Иностранец, @ФИОСборное, @ФИОРодПод, @ТипОтношения); SELECT SCOPE_IDENTITY()", list2, "Сотрудники", 1) 'возвращает ID


        'txt46Delegat() 'запуск заранее для получения данных для StrSql33

        If TextBox46.Text = "" Then
            Dtxt46 = Nothing
        ElseIf txtbx46l > 2 Then
            Dtxt46 = CType(Replace(arrtbox("TextBox46"), ".", ","), Double)
        Else
            Dtxt46 = CType(arrtbox("TextBox46"), Integer)
        End If

        '        Dim StrSql2 As String = "SELECT КодСотрудники FROM Сотрудники
        'WHERE ФИОСборное='" & surNameAll & "' And НазвОрганиз = '" & a & "' and ИДНомер='" & arrtbox("TextBox8") & "' AND ПровДатыКонтр='" & arrtmask("MaskedTextBox3") & "'"
        '        Dim ds25 As DataTable = Selects(StrSql2)


        '        If errds = 1 Then
        '            Try
        '                Dim fvxz As String = ds25.Rows(0).Item(0).ToString()
        '            Catch ex As Exception
        '                MessageBox.Show(ex.ToString)

        '            End Try
        '        End If

        Dim ФОТ2 As Double = Replace(arrtbox("TextBox48"), ".", ",")
        Dim ФОТ3 As Double = Replace(combx10, ".", ",")
        ФОТ2 = ФОТ2 * ФОТ3

        'Dim ФОТ2 As Double = Math.Round(TextBox48.Text * ComboBox10.Text, 2)
        Dim idClient As Integer

        idClient = Newid
        IDsot1 = Nothing
        IDsot1 = Newid
        Dim cbx7 As String
        If combx7 = "нет" Then
            cbx7 = ""
        Else
            cbx7 = combx7
        End If

        Dim dcx As Double = Replace(arrtbox("TextBox48"), ".", ",")
        Dim fgd As Double = CType(arrtbox("TextBox33") & "," & arrtbox("TextBox44"), Double)

        Dim list As New Dictionary(Of String, Object)

        With list
            Try
                .Add("@ПовышОклРуб", Math.Round(fgd * Replace(Dtxt46, ",", ".") / 100, 2))
            Catch ex As Exception
                .Add("@ПовышОклРуб", Math.Round(fgd * Replace(Dtxt46, ".", ",") / 100, 2))
                MDIParent1.ОбработкаОшибок(ex, "Form2, 4865")
            End Try

            Try
                .Add("@ЧасоваяТарифСтавка", Math.Round(Replace(dcx, ",", ".") / 168, 2))
            Catch ex As Exception
                .Add("@ЧасоваяТарифСтавка", Math.Round(Replace(dcx, ".", ",") / 168, 2))
                MDIParent1.ОбработкаОшибок(ex, "Form2, 4873")

            End Try

            .Add("@ИДСотр", idClient)
            .Add("@Должность", combx9)
            .Add("@Разряд", cbx7)
            .Add("@ТарифнаяСтавка", arrtbox("TextBox33") & "." & arrtbox("TextBox44"))
            .Add("@ПовышОклПроц", Replace(Dtxt46, ",", "."))
            .Add("@РасчДолжностнОклад", Replace(dcx, ",", "."))
            .Add("@Отдел", combx8)
        End With

        Updates(stroka:="INSERT INTO Штатное(ИДСотр, Должность, Разряд, ТарифнаяСтавка, ПовышОклПроц, РасчДолжностнОклад,
Отдел,ПовышОклРуб,ЧасоваяТарифСтавка)
VALUES(@ИДСотр, @Должность, @Разряд, @ТарифнаяСтавка, @ПовышОклПроц,
@РасчДолжностнОклад, @Отдел, @ПовышОклРуб, @ЧасоваяТарифСтавка)", list)

        'Добавляем ФОТ и обновляем таблицу штатное
        Updates(stroka:="UPDATE Штатное SET ФонОплатыТруда=" & Replace(ФОТ2, ",", ".") & " WHERE ИДСотр=@ИДСотр", list, "Штатное")


        'Вставляем в таблицу продление контракта и обновляем таблицу.
        Updates(stroka:="INSERT INTO ПродлКонтракта(IDСотр,ФИО,ДатаПриема,ДатаОкончания,СрокКонтракта,НомерУвед)
VALUES(@ИДСотр,'" & surNameAll & "','" & arrtmask("MaskedTextBox4") & "','" & arrtmask("MaskedTextBox5") & "',
'" & combx11 & "','" & arrtbox("TextBox38") & "')", list, "ПродлКонтракта")



        Dim _ПоСовмест, _СуммирУчет As String
        If CheckBox2.Checked = True Then
            _ПоСовмест = "по совместительству"
        Else
            _ПоСовмест = ""
        End If
        If CheckBox4.Checked = True Then
            _СуммирУчет = "Да"
        Else
            _СуммирУчет = ""
        End If


        Dim list4 As New Dictionary(Of String, Object)
        list4.Add("@IDСотр", idClient)
        list4.Add("@ДатаПриема", arrtmask("MaskedTextBox4"))
        list4.Add("@СрокКонтракта", combx11)
        list4.Add("@ТипРаботы", combx15)
        list4.Add("@Ставка", combx10)
        list4.Add("@ВремяНачРаботы", combx12)
        list4.Add("@ПродолРабДня", combx16)
        list4.Add("@Обед", arrtbox("TextBox49"))
        list4.Add("@ОкончРабДня", arrtbox("TextBox50"))
        list4.Add("@ДатаУведомлПродКонтр", ДатаУведомл(combx11, arrtmask("MaskedTextBox4")))
        list4.Add("@АдресОбъектаОбщепита", combx18)
        list4.Add("@ДатаЗарплаты", arrtbox("TextBox40"))
        list4.Add("@ДатаАванса", arrtbox("TextBox56"))
        list4.Add("@ПоСовмест", _ПоСовмест)
        list4.Add("@СуммирУчет", _СуммирУчет)
        If Примечани = "" Or Примечани Is Nothing Then
            list4.Add("@Примечание", "")
        Else
            list4.Add("@Примечание", Примечани)
        End If



        'Вставляем в таблицу Карточкасотрудника данные контракта и обновляем таблицу.
        Updates(stroka:="INSERT INTO КарточкаСотрудника(IDСотр,ДатаПриема,СрокКонтракта,ТипРаботы,
Ставка,ВремяНачРаботы,ПродолРабДня,Обед,
ОкончРабДня,ДатаУведомлПродКонтр,АдресОбъектаОбщепита,ДатаЗарплаты,
ДатаАванса,ПоСовмест,СуммирУчет,Примечание)
VALUES(@IDСотр,@ДатаПриема,@СрокКонтракта,@ТипРаботы,
@Ставка,@ВремяНачРаботы,@ПродолРабДня,@Обед,
@ОкончРабДня,@ДатаУведомлПродКонтр,@АдресОбъектаОбщепита,@ДатаЗарплаты,
@ДатаАванса,@ПоСовмест,@СуммирУчет,@Примечание)", list4, "КарточкаСотрудника")

        'Вставляем в таблицу ДогСотрудн данные контракта и обновляем таблицу.
        Dim list5 As New Dictionary(Of String, Object)
        list5.Add("@IDСотр", idClient)
        list5.Add("@Контракт", arrtbox("TextBox38"))
        list5.Add("@ДатаКонтракта", arrtmask("MaskedTextBox3"))
        list5.Add("@СрокОкончКонтр", arrtmask("MaskedTextBox5"))
        list5.Add("@Приказ", НПриказа)
        list5.Add("@Датаприказа", arrtmask("MaskedTextBox3"))

        Updates(stroka:="INSERT INTO ДогСотрудн(IDСотр,Контракт,ДатаКонтракта,СрокОкончКонтр,Приказ,Датаприказа)
VALUES(@IDСотр,@Контракт,@ДатаКонтракта,@СрокОкончКонтр,@Приказ,@Датаприказа)", list5, "ДогСотрудн")


        If arrtbox("TextBox25") <> "" Then
            дети(idClient)
        End If

        Статистика1(Trim(arrtbox("TextBox1")) & " " & Trim(arrtbox("TextBox2")) & " " & Trim(arrtbox("TextBox3")), "Добавление нового сотрудника", combx1)

    End Sub
    Private Sub дети(ByVal idClient As Integer)

        Dim StrSql5 As String = "SELECT КолДетей, IDСотр, ФИО, МестоРаботы, Телефон, 
ДетиПол1, ФИО1, ДатаРождения1, 
ДетиПол2, ФИО2, ДатаРождения2, 
ДетиПол3, ФИО3, ДатаРождения3, 
ДетиПол4, ФИО4, ДатаРождения4, 
ДетиПол5, ФИО5, ДатаРождения5
FROM СоставСемьи"

        Dim da As SqlDataAdapter = Доработчик(StrSql5)
        Dim ds5 As New DataSet
        da.Fill(ds5, "Семья")
        Dim cb5 As New SqlCommandBuilder(da)
        Dim dsNewRow5 As DataRow
        dsNewRow5 = ds5.Tables("Семья").NewRow()

        With dsNewRow5
            .Item("IDСотр") = idClient
            'dsNewRow1.Item("Фамилия") = Me.TextBox1.Text
            .Item("КолДетей") = combx14
            .Item("ФИО") = arrtbox("TextBox24")
            .Item("МестоРаботы") = arrtbox("TextBox23")
            .Item("Телефон") = arrtbox("TextBox19")
            .Item("ДетиПол1") = combx3
            .Item("ФИО1") = arrtbox("TextBox25")
            .Item("ДатаРождения1") = arrtbox("TextBox29")
            .Item("ДетиПол2") = combx4
            .Item("ФИО2") = arrtbox("TextBox27")
            .Item("ДатаРождения2") = arrtbox("TextBox26")
            .Item("ДетиПол3") = combx5
            .Item("ФИО3") = arrtbox("TextBox30")
            .Item("ДатаРождения3") = arrtbox("TextBox28")
            .Item("ДетиПол4") = combx6
            .Item("ФИО4") = arrtbox("TextBox32")
            .Item("ДатаРождения4") = arrtbox("TextBox31")
            .Item("ДетиПол5") = combx13
            .Item("ФИО5") = arrtbox("TextBox36")
            .Item("ДатаРождения5") = arrtbox("TextBox35")
        End With
        ds5.Tables("Семья").Rows.Add(dsNewRow5)
        da.Update(ds5, "Семья")

        If connДоработчик.State = ConnectionState.Open Then
            connДоработчик.Close()
        End If

    End Sub

    Private Function УдалениеСотр() As Integer

        If CheckBox5.Checked = True And CheckBox27.Checked = True Then
            Me.Cursor = Cursors.WaitCursor
            'StrSql = "SELECT КодСотрудники FROM Сотрудники WHERE НазвОрганиз='" & ComboBox1.Text & "' and ФИОСборное='" & ComboBox19.Text & "'"
            'ds = Selects(StrSql)
            Dim idc As Integer = CType(Label96.Text, Integer)
            Dim list As New Dictionary(Of String, Object)
            list.Add("@КодСотрудники", idc)

            Updates(stroka:="DELETE FROM Сотрудники WHERE КодСотрудники=@КодСотрудники", list)

            Updates(stroka:="DELETE FROM КарточкаСотрудника WHERE IDСотр=@КодСотрудники", list)

            Updates(stroka:="DELETE FROM ДогСотрудн WHERE IDСотр=@КодСотрудники", list)

            Updates(stroka:="DELETE FROM Штатное WHERE ИДСотр=@КодСотрудники", list)



            Dim dRow = dtPutiDokumentovAll.Select("IDСотрудник=" & idc & "")

            If dRow.Count > 0 Then
                Dim var2 = From x In dtPutiDokumentovAll.AsEnumerable Where Not IsDBNull(x.Item("IDСотрудник")) Select x
                Dim var = From x1 In var2 Where x1.Item("IDСотрудник") = idc Select x1.Item("Путь") & x1.Item("ИмяФайла")
                For b As Integer = 0 To var.Count - 1
                    DeleteFluentFTP(var(b).ToString)
                Next


                'For Each f.In var
                '    _DeleteFluentFTP(f)
                'Next
                Updates(stroka:="DELETE FROM ПутиДокументов WHERE IDСотрудник=" & idc & "")
            End If


            Parallel.Invoke(Sub() RunMoving2())
            Parallel.Invoke(Sub() RunMoving4())
            Статистика1(ComboBox19.Text, "Удаление сотрудника", ComboBox1.Text)
            MessageBox.Show("Сотрудник удален из базы!", Рик)
            CheckBox27.Checked = False

            Me.Cursor = Cursors.Default
            очПоля = 1
            CheckBox6.Checked = True
            CheckBox6.Checked = False
            ComboBox20.Text = ""
            ComboBox2.Text = ""
            ComboBox21.Text = ""
            TextBox40.Text = ""
            TextBox56.Text = ""
            ComboBox17.Text = ""
            MaskedTextBox3.Text = DateTime.Now.ToString("dd.MM.yyyy")
            MaskedTextBox4.Text = DateTime.Now.ToString("dd.MM.yyyy")
            Label85.Text = "NO"
            Label89.Text = "NO"
            Label90.Text = "NO"
            Com1sel()
            ComboBox19.Text = ""
            Label96.Text = ""


            Return 1
        End If

        Return 0

    End Function
    'Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs)
    '    If e.KeyCode = Keys.Enter Then
    '        e.SuppressKeyPress = True
    '        Me.TextBox2.Focus()
    '    End If

    'End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox3.Select()
        End If

    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox6.Select()
        End If

    End Sub

    Private Sub ComboBox15_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox15.SelectedValueChanged

        If ComboBox15.Text = "График" Or ComboBox15.Text = "ПВТР" Then
            ComboBox12.Visible = False
            ComboBox16.Visible = False
            Label54.Visible = False
            Label66.Visible = False
        Else
            ComboBox12.Visible = True
            ComboBox16.Visible = True
            Label54.Visible = True
            Label66.Visible = True
        End If
    End Sub

    Private Sub ComboBox11_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox11.SelectedValueChanged
        Dim MyCultureInfo As New CultureInfo("ru-RU")
        'Dim MyString As String = "12 Juni 2008"
        'Dim MyDateTime As DateTime = DateTime.Parse(MyString, MyCultureInfo)

        Dim pattern As String = "dd.MM.yyyy"
        Dim parsedDate As Date

        Dim dad As New Date
        dad = Date.ParseExact(MaskedTextBox4.Text, pattern, MyCultureInfo)


        Try
            'Dim dad As DateTime = Date.ParseExact(MaskedTextBox4.Text, pattern, MyCultureInfo)
            'Dim dadis As String = MaskedTextBox4.Text
            'Dim dadr As DateTime = DateTime.Parse(dadis, MyCultureInfo)

            'Dim dad3 = Format(dadis, "dd\/.\/MMMM.\/yyyy")
        Catch ex As Exception




            If ex.Message.Contains("Приведение строки ") Then
                Exit Sub
            End If

            'Dim MyCultureInfo1 As New CultureInfo("en-EN")
            'Dim dadis As String = MaskedTextBox4.Text.ToString
            'Dim dadr As DateTime = DateTime.Parse(dadis, MyCultureInfo)

        End Try

        Select Case ComboBox11.Text
            Case "1"
                MaskedTextBox5.Text = dad.AddMonths(12)
                Dim dad2 As Date = CDate(MaskedTextBox5.Text)
                MaskedTextBox5.Text = dad2.AddDays(-1)
            Case "2"
                MaskedTextBox5.Text = dad.AddMonths(24)
                Dim dad2 As Date = CDate(MaskedTextBox5.Text)
                MaskedTextBox5.Text = dad2.AddDays(-1)
            Case "3"
                MaskedTextBox5.Text = dad.AddMonths(36)
                Dim dad2 As Date = CDate(MaskedTextBox5.Text)
                MaskedTextBox5.Text = dad2.AddDays(-1)
            Case "4"
                MaskedTextBox5.Text = dad.AddMonths(48)
                Dim dad2 As Date = CDate(MaskedTextBox5.Text)
                MaskedTextBox5.Text = dad2.AddDays(-1)
            Case "5"
                MaskedTextBox5.Text = dad.AddMonths(60)
                Dim dad2 As Date = CDate(MaskedTextBox5.Text)
                MaskedTextBox5.Text = dad2.AddDays(-1)
            Case Else
                MaskedTextBox5.Text = Now.Date.ToShortDateString
        End Select
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        If CheckBox1.Checked = True Then
            Label80.ForeColor = Color.Green
            Label80.Text = "OK"
        End If
        If TextBox7.Text.Length = 7 And CheckBox1.Checked = False Then

            Label80.ForeColor = Color.Green
            Label80.Text = "OK"
        Else

            Label80.ForeColor = Color.Red
            Label80.Text = "NO"
        End If
    End Sub

    Private Sub TextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox5.Focus()
        End If
    End Sub

    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox4.Focus()
        End If
    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox34.Focus()
        End If
    End Sub

    Private Sub TextBox21_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox21.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox20.Focus()
        End If
    End Sub

    Private Sub TextBox20_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox20.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox37.Focus()
        End If
    End Sub

    Private Sub TextBox37_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox37.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.MaskedTextBox10.Focus()
        End If
    End Sub

    'Private Sub MaskedTextBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox10.KeyDown
    '    If e.KeyCode = Keys.Enter Then
    '        e.SuppressKeyPress = True
    '        Me.TextBox12.Focus()
    '    End If
    'End Sub

    Private Sub TextBox24_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox24.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox23.Focus()
        End If
    End Sub
    Private Sub TextBox23_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox23.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox19.Focus()
        End If

    End Sub
    Private Sub TextBox19_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox19.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox14.Focus()
        End If
    End Sub

    Private Sub TextBox25_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox25.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox27.Focus()
        End If
    End Sub

    Private Sub TextBox27_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox27.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox30.Focus()
        End If
    End Sub
    Private Sub TextBox30_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox30.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox32.Focus()
        End If
    End Sub

    Private Sub TextBox32_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox32.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox36.Focus()
        End If
    End Sub

    Private Sub TextBox36_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox36.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox12.Focus()
        End If
    End Sub

    Private Sub TextBox12_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox12.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox7.Focus()
        End If
    End Sub
    Private Sub ЧастЧист(ByVal d As Integer)
        Label96.Text = ""
        ComboBox19.Enabled = False
        Label48.Text = ""
        ComboBox7.Enabled = True
        ComboBox9.Enabled = True
        ComboBox8.Enabled = True
        ComboBox19.Enabled = False
        ComboBox1.Enabled = True

        CheckBox6.Checked = True
        CheckBox26.Checked = False
        'GroupBox14.Enabled = True
        ComboBox19.Text = String.Empty
        ComboBox19.Text = ""
        CheckBox23.Visible = False
        CheckBox23.Checked = False
        CheckBox26.Visible = False
        'Соед(0)
        Com1sel()
        'Соед(0)
        ComboBox19.Enabled = False
        CheckBox27.Enabled = False
        MaskedTextBox3.Text = Now.ToShortDateString
        MaskedTextBox4.Text = Now.ToShortDateString
        MaskedTextBox5.Text = Now.ToShortDateString
        mlk = d
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged




        If ComboBox1.Text = "" Then
            MessageBox.Show("Выберите организацию!", Рик)
            CheckBox5.Checked = False
            Exit Sub
        End If


        If CheckBox5.Checked = True Then

            CheckBox28.Enabled = True
            ComboBox7.Enabled = False
            ComboBox9.Enabled = False
            ComboBox8.Enabled = False
            ComboBox19.Enabled = True
            ComboBox1.Enabled = False

            'GroupBox14.Enabled = False
            CheckBox23.Visible = True
            CheckBox26.Visible = True
            CheckBox27.Enabled = True

        Else
            ДогПодрВклЧекбокс5 = False
            CheckBox28.Enabled = False
            'If MessageBox.Show("Очистить все поля?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            '    ЧастЧист(0)




            'Else
            ЧастЧист(1)
            Dim var1 = From x In dtShtatnoeOtdelyAll.Rows Where Not IsDBNull(x.Item("Клиент")) Select x
            Dim var = From x In var1 Order By x.Item("Отделы") Ascending Where x.Item("Клиент") = ComboBox1.Text Select x.Item("Отделы") Distinct.ToList 'рабочий linq для заполнения комбобоксов  и order by

            For Each r In var
                ComboBox8.AutoCompleteCustomSource.Add(r.ToString())
                ComboBox8.Items.Add(r.ToString)
                'Me.ComboBox19.Items.Add(r(1).ToString)

            Next
            ComboBox8.Text = ""
            'End If

            'закрываем чекбоксы
            CheckBox8.Checked = False
            CheckBox9.Checked = False
            CheckBox10.Checked = False
            CheckBox8.Enabled = False
            CheckBox9.Enabled = False
            CheckBox10.Enabled = False

            If CheckBox7.Enabled = False Then
                CheckBox7.Enabled = True
            End If


            TabControl1.TabPages.Remove(TabPage1)
            TabControl1.TabPages.Remove(TabPage2)
            TabControl1.TabPages.Remove(TabPage3)
            TabControl1.TabPages.Remove(TabPage4)
            TabControl1.TabPages.Add(TabPage1)
            TabControl1.TabPages.Add(TabPage2)






        End If
        CheckBox6.Checked = False
    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox20.Text = ""
        TextBox21.Text = ""
        MaskedTextBox10.Text = ""
        TextBox37.Text = ""
        TextBox24.Text = ""
        TextBox23.Text = ""
        TextBox19.Text = ""
        TextBox25.Text = ""
        TextBox27.Text = ""
        TextBox30.Text = ""
        TextBox32.Text = ""
        TextBox36.Text = ""
        ComboBox14.Text = String.Empty
        TextBox29.Text = ""
        TextBox26.Text = ""
        TextBox28.Text = ""
        TextBox31.Text = ""
        TextBox35.Text = ""
        MaskedTextBox1.Text = ""
        MaskedTextBox2.Text = ""
        TextBox12.Text = ""
        TextBox7.Text = ""
        TextBox9.Text = ""
        TextBox8.Text = ""
        TextBox45.Text = ""

        If CheckBox7.Checked = False Then

            'лист2
            Label79.Text = "NO"
            Label79.ForeColor = Color.Red
            Label85.Text = "NO"
            Label85.ForeColor = Color.Red

            Label88.Text = "NO"
            Label88.ForeColor = Color.Red
            Label89.Text = "NO"
            Label89.ForeColor = Color.Red
            Label90.Text = "NO"
            Label90.ForeColor = Color.Red

            TextBox40.Text = String.Empty
            TextBox56.Text = String.Empty
            ComboBox7.Text = String.Empty
            ComboBox8.Text = String.Empty
            ComboBox9.Text = String.Empty
            ComboBox10.Text = String.Empty
            ComboBox18.Text = String.Empty
            ComboBox12.Text = String.Empty
            ComboBox15.Text = String.Empty
            ComboBox16.Text = String.Empty
            ComboBox11.Text = String.Empty
            TextBox33.Text = ""
            TextBox43.Text = ""
            TextBox46.Text = ""
            TextBox47.Text = ""
            TextBox48.Text = ""
            TextBox38.Text = ""
            TextBox41.Text = ""
            TextBox49.Text = ""
            TextBox50.Text = ""
            TextBox44.Text = ""
            CheckBox6.Checked = False

            If очПоля = 1 Then
                очПоля = 0
            Else
                Me.ComboBox9.Items.Clear()
                Me.ComboBox8.Items.Clear()
            End If


        End If

        Try
            If CheckBox7.Checked = True Then
                ComboBox22.Text = String.Empty
                ComboBox23.Text = String.Empty
                ComboBox25.Text = String.Empty
                ComboBox24.Text = String.Empty
                TextBox55.Text = ""
                MaskedTextBox6.Text = ""
                MaskedTextBox7.Text = ""
                MaskedTextBox8.Text = ""
                TextBox61.Text = ""
                TextBox62.Text = ""
                ListBox1.Items.Clear()
                TextBox39.Text = ""
                TextBox63.Text = ""
            End If
        Catch ex As Exception

        End Try




    End Sub
    Private Sub ВстСтарФамилию(ByVal ds As DataRow())
        If ds(0).Item(32).ToString <> "" Then
            Label98.Text = "Старая фамилия ( " & ds(0).Item(32).ToString & " ) была до " & Strings.Left(ds(0).Item(40).ToString, 10)
        End If
    End Sub
    Private Sub Comb22Update()

        dbcx = New DbAll1DataContext()
        Dim ds1 = From x In dbcx.ДогПодДолжн
                  Where x.Клиент = ComboBox1.Text
                  Order By x.Должность
                  Select x.Должность, x.Код

        'Dim ds1 = dtDogPodrDoljnostAll.Select("Клиент='" & ComboBox1.Text & "'")




        'ComboBox22.AutoCompleteCustomSource.Clear()
        'ComboBox22.DataSource.Clear()

        'ComboBox22.AutoCompleteMode = AutoCompleteMode.Suggest
        'ComboBox22.AutoCompleteCustomSource = AutoCompleteSource.ListItems(ds1)
        ComboBox22.DataSource = ds1
        ComboBox22.DisplayMember = "Должность"
        ComboBox22.ValueMember = "Код"


        'For Each r In ds1
        '    Me.ComboBox22.AutoCompleteCustomSource.Add(r.Должность.ToString())
        '    Me.ComboBox22.Items.Add(r.Должность.ToString)
        'Next
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        If RichTextBox2.Text = "" Then
            MessageBox.Show("Введите должность!", Рик)
            Exit Sub
        End If

        Using dbcx = New DbAll1DataContext
            Dim ds = From x In dbcx.ДогПодДолжн.AsEnumerable Where x.Клиент = ComboBox1.Text And x.Должность = RichTextBox2.Text Select x
            If ds.Count > 0 Then
                MessageBox.Show("Должность " & RichTextBox2.Text & " уже существует!", Рик)
                Exit Sub
            End If
        End Using

        'Dim ds = dtDogPodrDoljnostAll.Select("Клиент='" & ComboBox1.Text & "' And Должность ='" & RichTextBox2.Text & "'")
        'If ds.Length > 0 Then
        '    If MessageBox.Show("Должность " & RichTextBox2.Text & " уже существует!" & vbCrLf & "Содать новую?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
        '        Exit Sub
        '    End If
        'End If
        Dim f2 As String = RichTextBox2.Text
        Dim db As New DbAll1DataContext() 'добавление в базу datacontext
        Dim f As New ДогПодДолжн()
        f.Клиент = ComboBox1.Text
        f.Должность = RichTextBox2.Text
        db.GetTable(Of ДогПодДолжн).InsertOnSubmit(f)
        db.SubmitChanges()
        idДолжность = f.Код

        '        Dim list As New Dictionary(Of String, Object)
        '        list.Add("@Клиент", ComboBox1.Text)
        '        list.Add("@Должность", RichTextBox2.Text)
        '        'Содаем новую должность
        '        idДолжность = Updates(stroka:="INSERT INTO ДогПодДолжн(Клиент,Должность)
        'VALUES(@Клиент,@Должность);SELECT SCOPE_IDENTITY()", list, "ДогПодДолжн", 1)
        Comb22Update()
        checkbx24()
        MessageBox.Show("Должность добавлена!", Рик)
        ComboBox22.SelectedItem = f2

    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        If RichTextBox2.Text = "" Then
            MessageBox.Show("Выберите должность для изменения!", Рик)
            Exit Sub
        End If




        Using dbcx = New DbAll1DataContext()
            'Dim var = dbcx.ДогПодДолжн.Single(Function(x) x.Код = ComboBox22.SelectedValue)
            Dim var = (From x In dbcx.ДогПодДолжн.AsEnumerable Where x.Код = ComboBox22.SelectedValue Select x).Single()
            If var IsNot Nothing Then
                var.Должность = RichTextBox2.Text
                dbcx.SubmitChanges()
            End If
        End Using

        Dim f As String = RichTextBox2.Text

        'Dim list As New Dictionary(Of String, Object)
        'list.Add("@Код", idДолжность)
        'list.Add("@Должность", RichTextBox2.Text)
        ''Содаем новую должность
        'Updates(stroka:="UPDATE ДогПодДолжн SET Должность=@Должность WHERE Код=@Код", list, "ДогПодДолжн")

        Parallel.Invoke(Sub() Comb22Update())
        Parallel.Invoke(Sub() checkbx24())
        ComboBox22.SelectedItem = f


        MessageBox.Show("Должность изменена!", Рик)

    End Sub
    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        If RichTextBox2.Text = "" Then
            MessageBox.Show("Выберите должность для удаления!", Рик)
            Exit Sub
        End If

        If MessageBox.Show("Удалить должность " & RichTextBox2.Text & " и её обязанности?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Exit Sub
        End If


        Using dbcx = New DbAll1DataContext() 'мой delete
            'Dim var = dbcx.ДогПодДолжн.Single(Function(x) x.Код = ComboBox22.SelectedValue)
            Dim var = (From x In dbcx.ДогПодДолжн.AsEnumerable Where x.Код = ComboBox22.SelectedValue Select x).FirstOrDefault
            If var IsNot Nothing Then
                dbcx.ДогПодДолжн.DeleteOnSubmit(var)
                dbcx.SubmitChanges()
            End If
        End Using


        Parallel.Invoke(Sub() Comb22Update())
        Parallel.Invoke(Sub() checkbx24())



        MessageBox.Show("Должность удалена!", Рик)



        'Dim list As New Dictionary(Of String, Object)
        'list.Add("@Код", idДолжность)

        ''Содаем новую должность
        'Updates(stroka:="DELETE ДогПодДолжн WHERE Код=@Код", list, "ДогПодДолжн")

    End Sub

    Private Sub Grid1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellClick
        If IsDBNull(Grid1) Or Grid1.Rows.Count = 0 Then
            Exit Sub
        End If

        RichTextBox2.Text = Grid1.CurrentRow.Cells(2).Value
        idДолжность = Grid1.CurrentRow.Cells(0).Value
        RichTextBox1.Text = Grid1.CurrentRow.Cells(3).Value
        idОбязAsync(ComboBox22.SelectedValue)
    End Sub
    Private Sub idОбяз(ByVal d As Integer)
        Using dbcx = New DbAll1DataContext
            idОбязанность = (From x In dbcx.ДогПодрОбязан.AsEnumerable
                             Join y In dbcx.ДогПодДолжн On x.ID Equals y.Код
                             Where x.Обязанности = Grid1.CurrentRow.Cells(3).Value And y.Код = d
                             Select x.Код).FirstOrDefault()
        End Using
    End Sub
    Private Async Sub idОбязAsync(ByVal d As Integer)
        Await Task.Run(Sub() idОбяз(d))
    End Sub

    Private Sub УскорИзменСотр()

        Dim ds = dtSotrudnikiAll.Select("КодСотрудники= " & КодСотрудника & "")


        'Dim StrSql As String = "Select * From Сотрудники Where КодСотрудники= " & КодСотрудника & ""
        'Dim ds As DataTable = Selects(StrSql)

        ВстСтарФамилию(ds) 'проверка девечьей фамилии


        With Me
            'If TextBox1.InvokeRequired Then
            '    Me.Invoke(New txtbx1(AddressOf УскорИзменСотр))
            'Else
            .TextBox1.Text = ds(0).Item(2).ToString
            'End If
            .TextBox2.Text = ds(0).Item(3).ToString
            .TextBox3.Text = ds(0).Item(4).ToString
            .TextBox6.Text = ds(0).Item(6).ToString
            .TextBox5.Text = ds(0).Item(7).ToString
            .TextBox4.Text = ds(0).Item(8).ToString
            .TextBox21.Text = ds(0).Item(16).ToString 'прописка
            .TextBox20.Text = ds(0).Item(17).ToString
            .TextBox37.Text = ds(0).Item(18).ToString 'телгор
            .MaskedTextBox10.Text = ds(0).Item(19).ToString ' телмоб
            .TextBox7.Text = ds(0).Item(11).ToString 'папсорт номер
            .TextBox12.Text = ds(0).Item(10).ToString 'серия
            .MaskedTextBox1.Text = ds(0).Item(12).ToString 'когда выдан
            .MaskedTextBox2.Text = ds(0).Item(13).ToString 'по какое
            .TextBox9.Text = ds(0).Item(14).ToString 'кем выдан
            .TextBox8.Text = ds(0).Item(15).ToString ' ID паспорта
            .TextBox45.Text = ds(0).Item(21).ToString 'номер свидет
            .TextBox34.Text = ds(0).Item(23).ToString
            .TextBox11.Text = ds(0).Item(24).ToString
            .TextBox10.Text = ds(0).Item(25).ToString
            .combx28 = ds(0).Item(26).ToString
            Select Case combx1
                Case "Амасейлс"
                    .MaskedTextBox9.Text = ds(0).Item(28).ToString
                    .TextBox51.Text = ds(0).Item(29).ToString
                Case "ЛемеЛ Лабс"
                    .MaskedTextBox9.Text = ds(0).Item(28).ToString
                    .TextBox51.Text = ds(0).Item(29).ToString
            End Select
            CheckBox1.Checked = CType(ds(0).Item(31).ToString, Boolean)
        End With

        'Зполняем семью

        Dim ds1 = dtSostavSemyiAll.Select("IDСотр = " & КодСотрудника & "")

        '        Dim StrSql1 As String = "Select СоставСемьи.КолДетей, СоставСемьи.ФИО, СоставСемьи.МестоРаботы, 
        'СоставСемьи.Телефон, СоставСемьи.ДетиПол1, СоставСемьи.ФИО1, СоставСемьи.ДатаРождения1, СоставСемьи.ДетиПол2,
        'СоставСемьи.ФИО2, СоставСемьи.ДатаРождения2, СоставСемьи.ДетиПол3, СоставСемьи.ФИО3, СоставСемьи.ДатаРождения3,
        'СоставСемьи.ДетиПол4, СоставСемьи.ФИО4, СоставСемьи.ДатаРождения4, СоставСемьи.ДетиПол5, СоставСемьи.ФИО5, СоставСемьи.ДатаРождения5
        'From СоставСемьи
        'Where СоставСемьи.IDСотр = " & КодСотрудника & ""
        '        Dim ds1 As DataTable = Selects(StrSql1)
        Try
            With Me
                .ComboBox14.Text = ds1(0).Item("КолДетей").ToString
                .TextBox24.Text = ds1(0).Item("ФИО").ToString
                .TextBox23.Text = ds1(0).Item("МестоРаботы").ToString
                .TextBox19.Text = ds1(0).Item("Телефон").ToString 'телефон
                .ComboBox3.Text = ds1(0).Item("ДетиПол1").ToString 'пол
                .TextBox25.Text = ds1(0).Item("ФИО1").ToString 'фио
                .TextBox29.Text = ds1(0).Item("ДатаРождения1").ToString 'дата рож1
                .ComboBox4.Text = ds1(0).Item("ДетиПол2").ToString 'пол
                .TextBox27.Text = ds1(0).Item("ФИО2").ToString 'фио
                .TextBox26.Text = ds1(0).Item("ДатаРождения2").ToString 'дата рож2
                .ComboBox5.Text = ds1(0).Item("ДетиПол3").ToString 'пол
                .TextBox30.Text = ds1(0).Item("ФИО3").ToString 'фио
                .TextBox28.Text = ds1(0).Item("ДатаРождения3").ToString 'дата рож3
                .ComboBox6.Text = ds1(0).Item("ДетиПол4").ToString
                .TextBox32.Text = ds1(0).Item("ФИО4").ToString
                .TextBox31.Text = ds1(0).Item("ДатаРождения4").ToString 'дата рож4
                .ComboBox13.Text = ds1(0).Item("ДетиПол5").ToString
                .TextBox36.Text = ds1(0).Item("ФИО5").ToString
                .TextBox35.Text = ds1(0).Item("ДатаРождения5").ToString 'дата рож5

            End With
        Catch ex As Exception
            'MessageBox.Show("Некоторые данные не зарегистрированы в системе!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
        End Try

    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        If RichTextBox1.Text = "" Then
            MessageBox.Show("Выберите объект для изменения!", Рик)
            Exit Sub
        End If

        Using dbcx = New DbAll1DataContext
            Dim v = (From x In dbcx.ДогПодрОбязан.AsEnumerable Where x.ID = ComboBox22.SelectedValue Select x.Обязанности).ToList
            If v.Contains(RichTextBox1.Text) Then
                MessageBox.Show("Внесите изменения!", Рик)
                Exit Sub
            End If
        End Using



        Using dbcx = New DbAll1DataContext() 'мой update
            Dim idob = (From x In dbcx.ДогПодрОбязан.AsEnumerable
                        Join y In dbcx.ДогПодДолжн.AsEnumerable On x.ID Equals y.Код
                        Where x.Обязанности = Grid1.CurrentRow.Cells(3).Value And y.Код = ComboBox22.SelectedValue
                        Select x.Код).FirstOrDefault()

            Dim var = (From x In dbcx.ДогПодрОбязан.AsEnumerable Where x.Код = idob Select x).SingleOrDefault
            If var IsNot Nothing Then
                var.Обязанности = RichTextBox1.Text
                dbcx.SubmitChanges()
            End If
        End Using


        Parallel.Invoke(Sub() checkbx24())


        ListBox1.DataSource = list1Update(ComboBox1.Text, ComboBox22.Text)

        MessageBox.Show("Обязанность изменена!", Рик)

    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        If RichTextBox2.Text = "" Then
            MessageBox.Show("Выберите объект для удаления!", Рик)
            Exit Sub
        End If

        If MessageBox.Show("Удалить обязанность?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.Cancel Then
            Exit Sub
        End If

        Using dbcx = New DbAll1DataContext() 'мой удаление
            Dim idob = (From x In dbcx.ДогПодрОбязан.AsEnumerable
                        Join y In dbcx.ДогПодДолжн.AsEnumerable On x.ID Equals y.Код
                        Where x.Обязанности = Grid1.CurrentRow.Cells(3).Value And y.Код = ComboBox22.SelectedValue
                        Select x.Код).FirstOrDefault()

            Dim var = (From x In dbcx.ДогПодрОбязан.AsEnumerable Where x.Код = idob Select x).Single
            If var IsNot Nothing Then
                dbcx.ДогПодрОбязан.DeleteOnSubmit(var)
                dbcx.SubmitChanges()
            End If
        End Using
        Parallel.Invoke(Sub() checkbx24())
        RichTextBox1.Text = ""

        ListBox1.DataSource = list1Update(ComboBox1.Text, ComboBox22.Text)
        MessageBox.Show("Данные удалены!", Рик)

    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        If RichTextBox1.Text = "" Then
            MessageBox.Show("Заполните обязанность!", Рик)
            Exit Sub
        End If


        Using dbcx = New DbAll1DataContext
            Dim v = (From x In dbcx.ДогПодрОбязан.AsEnumerable Where x.ID = ComboBox22.SelectedValue Select x.Обязанности).ToList
            If v.Contains(RichTextBox1.Text) Then
                MessageBox.Show("Обязанность '" & RichTextBox1.Text & "'" & vbCrLf & "уже создана!", Рик)
                Exit Sub
            End If
        End Using

        'мой insert
        Using dbcx = New DbAll1DataContext
            Dim f As New ДогПодрОбязан()
            f.Обязанности = RichTextBox1.Text
            f.ID = ComboBox22.SelectedValue
            dbcx.ДогПодрОбязан.InsertOnSubmit(f)
            dbcx.SubmitChanges()
            idДолжность = f.Код
        End Using



        Parallel.Invoke(Sub() checkbx24())
        ListBox1.DataSource = list1Update(ComboBox1.Text, ComboBox22.Text)
        'listbx1Saync(ComboBox1.Text, ComboBox22.Text)

        MessageBox.Show("Обязанность добавлена!", Рик)

    End Sub

    Private Sub УскорИзменСотрКарт()
        'Await Task.Delay(0)

        'Зполняем КарточкуСотрудника

        '        Dim StrSql As String = "Select КарточкаСотрудника.ДатаПриема, КарточкаСотрудника.СрокКонтракта, КарточкаСотрудника.ТипРаботы, 
        'КарточкаСотрудника.Ставка, КарточкаСотрудника.ВремяНачРаботы, КарточкаСотрудника.ПродолРабДня, КарточкаСотрудника.АдресОбъектаОбщепита,
        'КарточкаСотрудника.ДатаЗарплаты, КарточкаСотрудника.ДатаАванса, КарточкаСотрудника.ПоСовмест, КарточкаСотрудника.СуммирУчет, НаличиеИспытСрока, ПериодОтпДляКонтр
        'From КарточкаСотрудника
        'Where КарточкаСотрудника.IDСотр =  " & КодСотрудника & ""
        '        Dim ds As DataTable = Selects(StrSql)


        Dim ds = dtKartochkaSotrudnikaAll.Select("IDСотр =  " & КодСотрудника & "")


        If ds.Length > 0 Then
            With Me
                .MaskedTextBox4.Text = ds(0).Item("ДатаПриема").ToString
                .ComboBox11.Text = ds(0).Item("СрокКонтракта").ToString
                .ComboBox15.Text = ds(0).Item("ТипРаботы").ToString
                Dim s As String
                If ds(0).Item("Ставка") = "1" Then
                    ComboBox10.Text = ds(0).Item("Ставка").ToString & ".0"
                Else
                    s = ds(0).Item("Ставка").ToString
                    s = Replace(s, ",", ".")
                    ComboBox10.Text = s.ToString
                End If
                '
                .ComboBox12.Text = ds(0).Item("ВремяНачРаботы").ToString
                .ComboBox16.Text = ds(0).Item("ПродолРабДня").ToString
                .ComboBox18.Text = ds(0).Item("АдресОбъектаОбщепита").ToString
                .TextBox40.Text = ds(0).Item("ДатаЗарплаты").ToString
                .TextBox56.Text = ds(0).Item("ДатаАванса").ToString
                If ds(0).Item("ПоСовмест").ToString = "по совместительству" Then
                    CheckBox2.Checked = True
                End If
                If ds(0).Item("СуммирУчет").ToString = "Да" Then
                    CheckBox4.Checked = True
                End If

            End With
        Else

            'Dim StrSql1 As String = "Select НаличеДогПодряда From Сотрудники Where КодСотрудники = " & КодСотрудника & ""
            'Dim ds1 As DataTable = Selects(StrSql1)


            Dim ds1 = dtSotrudnikiAll.Select("КодСотрудники = " & КодСотрудника & "")
            If ds1(0).Item("НаличеДогПодряда") = "Да" Then
                'MessageBox.Show("С " & ComboBox19.SelectedItem & " оформлен договор подряда!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Information)
                CheckBox7.Checked = True

                'ДогПодрЗаполн()
                ДанДогПодр(КодСотрудника)

                Exit Sub
            End If





        End If



        'Try
        '    With Me
        '        .MaskedTextBox4.Text = ds(0).Item(0).ToString
        '        .ComboBox11.Text = ds(0).Item(1).ToString
        '        .ComboBox15.Text = ds(0).Item(2).ToString
        '        Dim s As String
        '        If ds(0).Item(3) = "1" Then
        '            ComboBox10.Text = ds(0).Item(3).ToString & ".0"
        '        Else
        '            s = ds(0).Item(3).ToString
        '            s = Replace(s, ",", ".")
        '            ComboBox10.Text = s.ToString
        '        End If
        '        '
        '        .ComboBox12.Text = ds(0).Item(4).ToString
        '        .ComboBox16.Text = ds(0).Item(5).ToString
        '        .ComboBox18.Text = ds(0).Item(6).ToString
        '        .TextBox40.Text = ds(0).Item(7).ToString
        '        .TextBox56.Text = ds(0).Item(8).ToString
        '        If ds(0).Item(9).ToString = "по совместительству" Then
        '            CheckBox2.Checked = True
        '        End If
        '        If ds(0).Item(10).ToString = "Да" Then
        '            CheckBox4.Checked = True
        '        End If
        '    End With
        'Catch ex As Exception

        '    Dim StrSql1 As String = "Select НаличеДогПодряда From Сотрудники Where КодСотрудники = " & КодСотрудника & ""
        '    Dim ds1 As DataTable = Selects(StrSql1)

        '    If ds1.Rows(0).Item(0) = "Да" Then
        '        'MessageBox.Show("С " & ComboBox19.SelectedItem & " оформлен договор подряда!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Information)
        '        CheckBox7.Checked = True

        '        'ДогПодрЗаполн()
        '        ДанДогПодр(КодСотрудника)

        '        Exit Sub
        '    End If
        'End Try


        If combx1 = "ЛемеЛ Лабс" Then
            vstavContr = ds(0).Item("НаличиеИспытСрока").ToString
            vstavContr1 = ds(0).Item("ПериодОтпДляКонтр").ToString
        End If


    End Sub
    Private Sub dtn2Click()
        If GroupBox19.Visible = True Then
            GroupBox19.Visible = False
            For Each gh In GroupBox19.Controls.OfType(Of RichTextBox)
                gh.Text = ""
            Next
            Dim dt As New DataTable
            Grid1.DataSource = dt
        End If
    End Sub
    Private Async Sub dtn2ClickAsyns()
        Await Task.Run(Sub() dtn2Click())
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        dtn2Click()
    End Sub

    Private Function УскорИзменСотрДог() As Integer


        'Заполняем из ДогСотрудн

        Dim StrSql As String = "Select ДогСотрудн.Контракт, ДогСотрудн.СрокОкончКонтр, ДогСотрудн.Приказ, ДогСотрудн.ДатаКонтракта
From ДогСотрудн
Where ДогСотрудн.IDСотр = " & КодСотрудника & ""
        Dim ds As DataTable = Selects(StrSql)
        Try
            'If TextBox38.InvokeRequired Then
            '    Me.Invoke(New txtb38(AddressOf УскорИзменСотрДог))
            'Else
            Me.TextBox38.Text = ds.Rows(0).Item(0).ToString
            'End If

            Me.MaskedTextBox5.Text = ds.Rows(0).Item(1).ToString
            Me.TextBox41.Text = Strings.Left(ds.Rows(0).Item(2).ToString, 3)
            Me.MaskedTextBox3.Text = ds.Rows(0).Item(3).ToString

            If (TabControl1.TabPages.Contains(TabPage1) = False) Then
                TabControl1.TabPages.Insert(0, TabPage1)
                TabControl1.SelectTab(TabPage1)
            End If

        Catch ex As Exception
            Dim strsql4 As String = "SELECT Должность FROM ДогПодряда WHERE ID=" & КодСотрудника & ""
            Dim ds4 As DataTable = Selects(strsql4)
            If errds = 1 Then

                MessageBox.Show("Сотрудник не зарегистрирован в системе!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                Return 1
            End If
        End Try



        'Заполняем из Штатное

        Dim StrSql1 As String = "Select Штатное.Отдел, Штатное.Должность, Штатное.Разряд, Штатное.ТарифнаяСтавка, Штатное.ПовышОклПроц, Штатное.РасчДолжностнОклад
        From Штатное Where Штатное.ИДСотр = " & КодСотрудника & ""
        Dim ds1 As DataTable = Selects(StrSql1)

        Try
            'ComboBox8.Text = ds.Rows(0).Item(0).ToString
            'ComboBox9.Text = ds.Rows(0).Item(1).ToString
            Label48.Text = ""
            If ds1.Rows(0).Item(2) <> "" Then
                Label48.Text = ds1.Rows(0).Item(1).ToString & " " & combxS19 & " работает в отделе " & ds1.Rows(0).Item(0).ToString & ", разряд " & ds1.Rows(0).Item(2).ToString
            Else
                Label48.Text = ds1.Rows(0).Item(1).ToString & " " & combxS19 & " работает в отделе " & ds1.Rows(0).Item(0).ToString
            End If

            'Должность = ds1.Rows(0).Item(1).ToString

            'CombBox7 = 0
            'CombBox7 = 1 'проверка для Комбобокс 7
            combx7 = ds1.Rows(0).Item(2).ToString ' перепроверить
            Dim ВхДан As String = ds1.Rows(0).Item(3).ToString
            Dim ВхданКол As Integer = ВхДан.Length


            Dim cela As Double = ds1.Rows(0).Item(3).ToString
            Dim cel As Double = Math.Floor(cela)
            TextBox33.Text = CType(cel, String)

            If cela - cel = 0 Then
                TextBox44.Text = "00"
            Else
                cela = Math.Round((cela - cel), 2)
                Dim Окон As String = CType(cela, String)
                If Окон.Length > 3 Then
                    TextBox44.Text = Strings.Right(Окон, 2)
                Else
                    Окон = Strings.Right(Окон, 1)
                    TextBox44.Text = Окон & "0"
                End If

            End If

            TextBox46.Text = ds1.Rows(0).Item(4).ToString
            Dim proc As Double = Replace(ds1.Rows(0).Item(4), ".", ",")

            If ds1.Rows(0).Item(5).ToString = "" Then
                TextBox48.Text = CType(Math.Round((cela + (cela * proc / 100)), 2), String)
            Else
                TextBox48.Text = ds1.Rows(0).Item(5).ToString
            End If

        Catch ex As Exception
            Dim strsql5 As String = "SELECT Должность FROM ДогПодряда WHERE ID=" & КодСотрудника & ""
            Dim ds5 As DataTable = Selects(strsql5)
            If errds = 1 Then
                MessageBox.Show("Сотрудник не зарегистрирован в системе!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Asterisk)

            End If

        End Try


        Dim StrSql2 As String = "Select АдресОбъектаОбщепита From КарточкаСотрудника Where IDСотр = " & КодСотрудника & ""
        Dim ds2 As DataTable = Selects(StrSql2)
        Try
            combx18 = ds2.Rows(0).Item(0).ToString
        Catch ex As Exception

        End Try

        Return 0

    End Function
    Private Function ПроверкаДоговорИлиПодрядЧерезСправочник(ByVal _КодСотр As Integer) As String

        Dim var As List(Of ДогСотрудн)
        Using dbcx As New DbAll1DataContext

            var = (From x In dbcx.ДогСотрудн.AsEnumerable
                   Where x.IDСотр = _КодСотр
                   Select x).ToList
        End Using

        If var.Count > 0 Then  'вставляем данные во вкладку контракт
            'ДогСотрудн.Контракт, ДогСотрудн.СрокОкончКонтр, ДогСотрудн.Приказ, ДогСотрудн.ДатаКонтракта
            TextBox38.Text = var(0).Контракт
            MaskedTextBox5.Text = var(0).СрокОкончКонтр
            TextBox41.Text = Strings.Left(var(0).Приказ, 3)
            MaskedTextBox3.Text = var(0).ДатаКонтракта


            Dim var1 As Штатное
            Using dbcx1 As New DbAll1DataContext 'штатное
                var1 = (From x In dbcx1.Штатное.AsEnumerable
                        Where x.ИДСотр = _КодСотр
                        Select x).FirstOrDefault()
            End Using

            If var1 IsNot Nothing Then
                'Штатное.Отдел, Штатное.Должность, Штатное.Разряд, Штатное.ТарифнаяСтавка, Штатное.ПовышОклПроц, Штатное.РасчДолжностнОклад
                Label48.Text = ""
                CheckBox26.Checked = True
                If var1.Разряд <> "" Then
                    'Label48.Text = var1.Должность & " " & combxS19 & " работает в отделе " & var1.Отдел & ", разряд " & var1.Разряд
                    ComboBox8.Text = var1.Отдел
                    ComboBox9.Text = var1.Должность
                    ComboBox7.Text = var1.Разряд
                Else
                    'Label48.Text = var1.Должность & " " & combxS19 & " работает в отделе " & var1.Отдел
                    ComboBox8.Text = var1.Отдел
                    ComboBox9.Text = var1.Должность
                End If


                combx7 = var1.Разряд ' перепроверить
                Dim ВхДан As String = var1.ТарифнаяСтавка
                Dim ВхданКол As Integer = ВхДан.Length


                Dim cela As Double = Replace(var1.ТарифнаяСтавка, ".", ",")
                Dim cel As Double = Math.Floor(cela)
                TextBox33.Text = CType(cel, String)

                If cela - cel = 0 Then
                    TextBox44.Text = "00"
                Else
                    cela = Math.Round((cela - cel), 2)
                    Dim Окон As String = CType(cela, String)
                    If Окон.Length > 3 Then
                        TextBox44.Text = Strings.Right(Окон, 2)
                    Else
                        Окон = Strings.Right(Окон, 1)
                        TextBox44.Text = Окон & "0"
                    End If

                End If

                TextBox46.Text = var1.ПовышОклПроц
                Dim proc As Double = Replace(var1.ПовышОклПроц, ".", ",")

                If CType(var1.РасчДолжностнОклад, String) = "" Then
                    TextBox48.Text = CType(Math.Round((cela + (cela * proc / 100)), 2), String)
                Else
                    TextBox48.Text = var1.РасчДолжностнОклад
                End If

            End If



            Dim var2 As КарточкаСотрудника
            Using dbc As New DbAll1DataContext 'адрес обьекта общепита 
                var2 = (From x In dbc.КарточкаСотрудника.AsEnumerable
                        Where x.IDСотр = _КодСотр
                        Select x).FirstOrDefault()
            End Using

            If var2 IsNot Nothing Then

                combx18 = var2.АдресОбъектаОбщепита

                With Me
                    .MaskedTextBox4.Text = var2.ДатаПриема
                    .ComboBox11.Text = var2.СрокКонтракта
                    .ComboBox15.Text = var2.ТипРаботы
                    Dim s As String
                    If var2.Ставка = "1" Then
                        ComboBox10.Text = var2.Ставка & ".0"
                    Else
                        s = var2.Ставка
                        s = Replace(s, ",", ".")
                        ComboBox10.Text = s.ToString
                    End If

                    .ComboBox12.Text = var2.ВремяНачРаботы
                    .ComboBox16.Text = var2.ПродолРабДня
                    .ComboBox18.Text = var2.АдресОбъектаОбщепита
                    .TextBox40.Text = var2.ДатаЗарплаты
                    .TextBox56.Text = var2.ДатаАванса
                    If var2.ПоСовмест = "по совместительству" Then
                        CheckBox2.Checked = True
                    End If
                    If var2.СуммирУчет = "Да" Then
                        CheckBox4.Checked = True
                    End If

                End With

            End If

            Return "Контракт"
        End If


        'проверяем на договор подряда час//

        Dim var3 As List(Of ДогПодряда)
        Using dbcx As New DbAll1DataContext

            var3 = (From x In dbcx.ДогПодряда.AsEnumerable
                    Join y In dbcx.Сотрудники.AsEnumerable On y.КодСотрудники Equals x.ID
                    Where x.ID = _КодСотр And x.ВидДоговора = "час" And y.НаличеДогПодряда = "Да"
                    Select x).ToList()
        End Using

        If var3.Count > 0 Then
            Dim f As New ДогПодВыборНомДоговора
            f.var3 = var3
            f.Flag = True
            f.Label2.Text = ComboBox19.Text
            f.ListBox1.DataSource = (From x In var3.AsEnumerable Select x.НомерДогПодр Distinct).ToList()
            f.ShowDialog()

            'выбарнный договор
            Dim lis = f.ВыборНомера

            'выбираем данные по выбранному договору из таблицы догподряда
            Dim var4 = (From x In var3.AsEnumerable Where x.НомерДогПодр = f.ВыборНомера Select x).FirstOrDefault()

            'выбираем недостающие данные для заполнения данных во вкладку договор подряда (час)
            Dim var5
            Using dbcxc As New DbAll1DataContext
                var5 = (From x In dbcxc.Клиент.AsEnumerable
                        Join y In dbcxc.ДогПодДолжн.AsEnumerable On x.НазвОрг Equals y.Клиент
                        Join c In dbcxc.ДогПодрОбязан.AsEnumerable On y.Код Equals c.ID
                        Where x.НазвОрг = ComboBox1.Text And y.Должность = var4.Должность
                        Select c).ToList()
            End Using

            'Вставляем данные в форму договор подряда (час)

            With Me

                Dim n, d As String
                If var4.НомерДогПодр.Contains(".") Then
                    'n = Mid(var4.НомерДогПодр, 1, InStr(var4.НомерДогПодр, ".")) 'значение слева до точки
                    n = var4.НомерДогПодр.Substring(0, var4.НомерДогПодр.IndexOf(".")) 'значение слева до точки
                    d = Strings.Right(var4.НомерДогПодр, var4.НомерДогПодр.Length - (n.Length + 1)) 'значение после точки
                Else
                    n = var4.НомерДогПодр
                    d = ""
                End If
                .TextBox55.Text = n
                .TextBox39.Text = d
                .ComboBox22.Text = var4.Должность
                .MaskedTextBox6.Text = var4.ДатаДогПодр
                .MaskedTextBox7.Text = var4.ДатаНачала
                .MaskedTextBox8.Text = var4.ДатаОконч
                .ComboBox27.Text = "час"
                .TextBox61.Text = var4.СтоимЧасаРуб
                .TextBox62.Text = var4.СтоимЧасаКоп
                .ListBox1.DataSource = var5
            End With








            'ДогПодВыборНомДоговора.var3 = var3
            'ДогПодВыборНомДоговора.Flag = True
            'ДогПодВыборНомДоговора.ShowDialog()

            'Чист()
            'StrSql = "Select * From ДогПодряда Where ID = " & ID & " and НомерДогПодр='" & ДогПодномДогПодНомДог & "'"
            'ds = Selects(StrSql)


            'подумать как заполнить форму договороа подряда час


            Return "Подряд(час)"
        End If









        Return "No"









    End Function

    Private Sub ДоговорПодрядаИноеЗаполенениеФормы()
        'года 
        Dim com32 = listFluentFTP(ComboBox1.Text & "/Договор подряда/")
        ComboBox32.DataSource = com32
        'Объект общепита, единицы измерения из списка
        Dim list3
        Using dbcx As New DbAll1DataContext
            ComboBox31.DataSource = (From x In dbcx.ОбъектОбщепита.AsEnumerable
                                     Where x.НазвОрг = ComboBox1.Text
                                     Select x.АдресОбъекта).ToList()

            list3 = (From x In dbcx.ДогПодОсобен.AsEnumerable
                     Where x.Организация = ComboBox1.Text
                     Select x.ЕденицаИзм Distinct).ToList()
        End Using
        'даты
        MaskedTextBox11.Text = Now.ToShortDateString
        MaskedTextBox12.Text = Now.ToShortDateString
        MaskedTextBox13.Text = Now.ToShortDateString

        'заполняем единицы измерения
        ComboBox29.Items.Clear()
        For Each item In list3
            ComboBox29.Items.Add(item)
        Next

        Dim _grid5 As List(Of ДогПодОсобен)
        Using dbcx As New DbAll1DataContext
            _grid5 = (From x In dbcx.ДогПодОсобен.AsEnumerable
                      Where x.Организация = ComboBox1.Text
                      Select x).ToList()
        End Using
        If _grid5.Count > 0 Then
            Grid5.DataSource = _grid5
            Grid5.Columns(1).HeaderText = "Единица измерения"
            Grid5.Columns(2).HeaderText = "Выполняемая работа"
            Grid5.Columns(3).HeaderText = "Выполняемая работа (Именительный падеж)"
            Grid5.Columns(0).Visible = False
            Grid5.Columns(4).Visible = False
            GridView(Grid5)
        Else
            Dim dt As New DataTable
            dt.Columns.Add("Единица измерения")
            dt.Columns.Add("Выполняемая работа")
            dt.Columns.Add("Выполняемая работа (Именительный падеж)")
            Grid5.DataSource = dt
            GridView(Grid5)
        End If


        Dim dts As New DataTable
        dts.Columns.Add("Руб.")
        dts.Columns.Add("Коп.")
        dts.Columns.Add("Ед.изм")
        dts.Columns.Add("Вып.работа")
        dts.Columns.Add("Вып.работа (Имен.падеж)")
        Grid6.DataSource = dts
        'GridView(Grid6)
        Grid6.Columns(2).Width = 60
        Grid6.Columns(0).Width = 60
        Grid6.Columns(1).Width = 60
        Grid6.EnableHeadersVisualStyles = False
        Grid6.ColumnHeadersDefaultCellStyle.Font = New Font(Grid6.ColumnHeadersDefaultCellStyle.Font.FontFamily, 10, FontStyle.Bold)



    End Sub


    Private Sub CheckBox9_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox9.CheckedChanged

        If CheckBox9.Checked = False Then Exit Sub 'временная мера
        If CheckBox9.Checked = True Then
            CheckBox8.Checked = False
            CheckBox10.Checked = False
        End If
        If ComboBox1.Text = "" Then
            MessageBox.Show("Выберите организацию!")
            CheckBox9.Checked = False
            Exit Sub
        End If




        If CheckBox9.Checked = True Then

            TabControl1.TabPages.Remove(TabPage1)
            TabControl1.TabPages.Remove(TabPage2)
            TabControl1.TabPages.Remove(TabPage3)
            TabControl1.TabPages.Remove(TabPage4)
            TabControl1.TabPages.Add(TabPage4)
            ДоговорПодрядаИноеЗаполенениеФормы()

        Else
            TabControl1.TabPages.Remove(TabPage4)
            TabControl1.TabPages.Add(TabPage1)
            TabControl1.TabPages.Add(TabPage2)
        End If

    End Sub

    Private Sub ComboBox32_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox32.SelectedIndexChanged
        ComboBox33.Items.Clear()
        ComboBox33.Text = ""

        Dim listCombo3 As Object = listFluentFTP(ComboBox1.Text & "/Договор подряда/" & ComboBox32.Text & "/")

        For Each item In listCombo3
            ComboBox33.Items.Add(Replace(item, FTPString & ComboBox1.Text & "/Договор подряда/" & ComboBox32.Text & "/", ""))
        Next
    End Sub

    Private Sub RichTextBox3_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox3.TextChanged
        RichTextBox6.Text = RichTextBox3.Text
    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        Dim f As New ДогПодрИноеВсплывФорма
        f.ShowDialog()
        Dim fl = f.ЕдИзм
        If fl <> "" Then
            ComboBox29.Items.Add(f.ЕдИзм)
        End If


    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        If ComboBox29.SelectedIndex = -1 Then
            MessageBox.Show("Выберите элемент для удаления!", Рик)
            Exit Sub
        End If

        ComboBox29.Items.RemoveAt(ComboBox29.SelectedIndex)


    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        If ComboBox29.SelectedIndex = -1 Then
            MessageBox.Show("Отметьте необходимую единицу измерения!", Рик)
            Exit Sub
        End If

        If RichTextBox3.Text = "" Then
            MessageBox.Show("Заполните поле 'Выполняемая работа'!", Рик)
            Exit Sub
        End If
        If RichTextBox6.Text = "" Then
            MessageBox.Show("Заполните поле 'Выполняемая работа (в именительном падеже)'!", Рик)
            Exit Sub
        End If
        'сохраняем данные в базу ДогПодОсобен
        Using dbcx As New DbAll1DataContext
            Dim f As New ДогПодОсобен()
            f.ЕденицаИзм = ComboBox29.SelectedItem
            f.Текст = RichTextBox3.Text
            f.ТесктИменПад = RichTextBox6.Text
            f.Организация = ComboBox1.Text
            dbcx.ДогПодОсобен.InsertOnSubmit(f)
            dbcx.SubmitChanges()
        End Using

        'Выбираем из базы необходимые данные
        Dim _grid5 As List(Of ДогПодОсобен)
        Using dbcx As New DbAll1DataContext
            _grid5 = (From x In dbcx.ДогПодОсобен.AsEnumerable
                      Where x.Организация = ComboBox1.Text
                      Select x).ToList()
        End Using

        'Вставляем данные в Grid5
        If _grid5.Count > 0 Then
            Grid5.DataSource = _grid5
            Grid5.Columns(1).Name = "Единица измерения"
            Grid5.Columns(2).Name = "Выполняемая работа"
            Grid5.Columns(3).Name = "Выполняемая работа (Именительный падеж)"
            Grid5.Columns(0).Visible = False

            Grid5.Columns(4).Visible = False
            GridView(Grid5)
        Else
            Dim dt As New DataTable
            dt.Columns.Add("Единица измерения")
            dt.Columns.Add("Выполняемая работа")
            dt.Columns.Add("Выполняемая работа (Именительный падеж)")
            Grid5.DataSource = dt
            GridView(Grid5)
        End If

        MessageBox.Show("Данные добавлены!", Рик)

    End Sub

    Private Sub Grid5_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid5.CellClick

        If Grid5.CurrentRow Is Nothing Then
            Exit Sub
        End If

        If Grid5.CurrentRow.Index = -1 Then
            Exit Sub
        End If


        If Not ComboBox29.Items.Contains(Grid5.CurrentRow.Cells(1).Value) Then
            ComboBox29.Items.Add(Grid5.CurrentRow.Cells(1).Value)
            ComboBox29.SelectedItem = Grid5.CurrentRow.Cells(1).Value
        Else
            ComboBox29.SelectedItem = Grid5.CurrentRow.Cells(1).Value
        End If

        RichTextBox3.Text = Grid5.CurrentRow.Cells(2).Value
        RichTextBox6.Text = Grid5.CurrentRow.Cells(3).Value
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        If ComboBox29.Text = "" Then
            MessageBox.Show("Выберите единицу измерения!")
            Exit Sub
        End If
        If RichTextBox3.Text = "" Then
            MessageBox.Show("Заполните поле 'Выполненная работа'!", Рик)
            Exit Sub
        End If
        If RichTextBox6.Text = "" Then
            MessageBox.Show("Заполните поле 'Выполняемая работа (в именительном падеже)'!", Рик)
            Exit Sub
        End If

        If Grid5.CurrentRow Is Nothing Then
            MessageBox.Show("Выберите в таблице строку для изменения!", Рик)
            Exit Sub
        End If

        Using dbcx = New DbAll1DataContext() 'мой update2
            Dim var = (From x In dbcx.ДогПодОсобен.AsEnumerable
                       Where x.Код = Grid5.CurrentRow.Cells(0).Value
                       Select x).SingleOrDefault
            If var IsNot Nothing Then
                var.ЕденицаИзм = ComboBox29.Text
                var.Текст = RichTextBox3.Text
                var.ТесктИменПад = RichTextBox6.Text
                dbcx.SubmitChanges()
            End If
        End Using

        MessageBox.Show("Данные изменены!", Рик)

        UpgradeGrid5()
    End Sub

    Private Sub UpgradeGrid5()
        Dim list3
        Using dbcx As New DbAll1DataContext
            ComboBox31.DataSource = (From x In dbcx.ОбъектОбщепита.AsEnumerable
                                     Where x.НазвОрг = ComboBox1.Text
                                     Select x.АдресОбъекта).ToList()

            list3 = (From x In dbcx.ДогПодОсобен.AsEnumerable
                     Where x.Организация = ComboBox1.Text
                     Select x.ЕденицаИзм Distinct).ToList()
        End Using
        'даты
        MaskedTextBox11.Text = Now.ToShortDateString
        MaskedTextBox12.Text = Now.ToShortDateString
        MaskedTextBox13.Text = Now.ToShortDateString

        'заполняем единицы измерения
        ComboBox29.Items.Clear()
        For Each item In list3
            ComboBox29.Items.Add(item)
        Next

        Dim _grid5 As List(Of ДогПодОсобен)
        Using dbcx As New DbAll1DataContext
            _grid5 = (From x In dbcx.ДогПодОсобен.AsEnumerable
                      Where x.Организация = ComboBox1.Text
                      Select x).ToList()
        End Using
        If _grid5.Count > 0 Then
            Grid5.DataSource = _grid5
            Grid5.Columns(1).Name = "Единица измерения"
            Grid5.Columns(2).Name = "Выполняемая работа"
            Grid5.Columns(3).Name = "Выполняемая работа (Именительный падеж)"
            Grid5.Columns(0).Visible = False
            Grid5.Columns(4).Visible = False
            GridView(Grid5)
        Else
            Dim dt As New DataTable
            dt.Columns.Add("Единица измерения")
            dt.Columns.Add("Выполняемая работа")
            dt.Columns.Add("Выполняемая работа (Именительный падеж)")
            Grid5.DataSource = dt
            GridView(Grid5)
        End If
        Me.Grid5.FirstDisplayedCell = Me.Grid5.CurrentCell
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

        If Grid5.CurrentRow Is Nothing Then
            MessageBox.Show("Выберите в таблице строку для удаления!", Рик)
            Exit Sub
        End If

        If MessageBox.Show("Удалить выполняемую руботу " & vbCrLf & Grid5.CurrentRow.Cells(2).Value & "?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.Cancel Then
            Exit Sub
        End If


        Using dbcx = New DbAll1DataContext() 'мой delete
            Dim var = (From x In dbcx.ДогПодОсобен.AsEnumerable
                       Where x.Код = Grid5.CurrentRow.Cells(0).Value
                       Select x).SingleOrDefault
            If var IsNot Nothing Then
                dbcx.ДогПодОсобен.DeleteOnSubmit(var)
                dbcx.SubmitChanges()

            End If
        End Using
        UpgradeGrid5()


    End Sub

    Private Sub Grid5_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid5.CellDoubleClick
        If Grid5.CurrentRow Is Nothing Or Grid5.CurrentRow.Index = -1 Then
            MessageBox.Show("Выберите строку!", Рик)
            Exit Sub
        End If
        Grid6Insert(Grid5.CurrentRow)

    End Sub

    Private Sub Grid6Insert(ByVal d As DataGridViewRow)
        Dim dt As New DataTable
        dt.Columns.Add("Руб.")
        dt.Columns.Add("Коп.")
        dt.Columns.Add("Ед.изм")
        dt.Columns.Add("Вып.работа")
        dt.Columns.Add("Вып.работа (Имен.падеж)")

        dt = Grid6.DataSource

        Dim row As DataRow = dt.NewRow
        row("Ед.изм") = d.Cells(1).Value
        row("Вып.работа") = d.Cells(2).Value
        row("Вып.работа (Имен.падеж)") = d.Cells(3).Value
        dt.Rows.Add(row)

        Grid6.DataSource = dt
        GridView(Grid6)
        Grid6.Columns(2).Width = 60
        Grid6.Columns(0).Width = 60
        Grid6.Columns(1).Width = 60
        Grid6.EnableHeadersVisualStyles = False
        Grid6.ColumnHeadersDefaultCellStyle.Font = New Font(Grid6.ColumnHeadersDefaultCellStyle.Font.FontFamily, 10, FontStyle.Bold)



    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        If TextBox22.Text = "" Then
            MessageBox.Show("Заполните поле рубли!", Рик)
            Exit Sub
        End If

        If Grid6.CurrentRow Is Nothing Then
            MessageBox.Show("Выберите строку для добавления стоимости работ!", Рик)
            Exit Sub
        End If

        If Grid6.CurrentRow.Index = -1 Then
            MessageBox.Show("Выберите строку для добавления стоимости работ!", Рик)
            Exit Sub
        End If

        If MessageBox.Show("Добавить данные?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.Cancel Then
            Exit Sub
        End If



        'вставляем полученные данные в таблицу
        Dim dt As New DataTable
        dt.Columns.Add("Руб.")
        dt.Columns.Add("Коп.")
        dt.Columns.Add("Ед.изм")
        dt.Columns.Add("Вып.работа")
        dt.Columns.Add("Вып.работа (Имен.падеж)")

        dt = Grid6.DataSource

        dt.Rows(Grid6.CurrentRow.Index).Item(0) = TextBox22.Text
        If TextBox52.Text = "" Then
            dt.Rows(Grid6.CurrentRow.Index).Item(1) = "00"
        Else
            If TextBox52.Text.Length < 2 And TextBox52.Text.Length > 0 Then
                dt.Rows(Grid6.CurrentRow.Index).Item(1) = "0" & TextBox52.Text
            Else
                dt.Rows(Grid6.CurrentRow.Index).Item(1) = TextBox52.Text
            End If

        End If


        Grid6.DataSource = dt
        GridView(Grid6)
        Grid6.Columns(2).Width = 60
        Grid6.Columns(0).Width = 60
        Grid6.Columns(1).Width = 60
        TextBox52.Text = ""
        TextBox22.Text = ""



    End Sub

    Private Sub Grid6_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid6.CellClick
        If Grid6.CurrentRow Is Nothing Then
            Exit Sub
        End If

        If Grid6.CurrentRow.Index = -1 Then
            Exit Sub
        End If




        If IsDBNull(Grid6.CurrentRow.Cells(0).Value) Then
            Exit Sub
        End If

        If IsDBNull(Grid6.CurrentRow.Cells(1).Value) Then
            Exit Sub
        End If

        TextBox22.Text = Grid6.CurrentRow.Cells(0).Value
        TextBox52.Text = Grid6.CurrentRow.Cells(1).Value

    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        If TextBox22.Text = "" Then
            MessageBox.Show("Заполните поле рубли!", Рик)
            Exit Sub
        End If
        If Grid6.CurrentRow Is Nothing Then
            MessageBox.Show("Выберите строку для изменения стоимости работ!", Рик)
            Exit Sub
        End If

        If Grid6.CurrentRow.Index = -1 Then
            MessageBox.Show("Выберите строку для изменения стоимости работ!", Рик)
            Exit Sub
        End If






        If MessageBox.Show("Изменить данные?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.Cancel Then
            Exit Sub
        End If

        Dim dt As New DataTable
        dt.Columns.Add("Руб.")
        dt.Columns.Add("Коп.")
        dt.Columns.Add("Ед.изм")
        dt.Columns.Add("Вып.работа")
        dt.Columns.Add("Вып.работа (Имен.падеж)")

        dt = Grid6.DataSource

        dt.Rows(Grid6.CurrentRow.Index).Item(0) = TextBox22.Text



        If TextBox52.Text = "" Then
            dt.Rows(Grid6.CurrentRow.Index).Item(1) = "00"
        Else
            If TextBox52.Text.Length < 2 And TextBox52.Text.Length > 0 Then
                dt.Rows(Grid6.CurrentRow.Index).Item(1) = "0" & TextBox52.Text
            Else
                dt.Rows(Grid6.CurrentRow.Index).Item(1) = TextBox52.Text
            End If

        End If


        Grid6.DataSource = dt
        GridView(Grid6)
        Grid6.Columns(2).Width = 60
        Grid6.Columns(0).Width = 60
        Grid6.Columns(1).Width = 60
        TextBox52.Text = ""
        TextBox22.Text = ""





    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        If Grid6.CurrentRow Is Nothing Then
            MessageBox.Show("Выберите строку для удаления!", Рик)
            Exit Sub
        End If

        If Grid6.CurrentRow.Index = -1 Then
            MessageBox.Show("Выберите строку для удаления!", Рик)
            Exit Sub
        End If

        If MessageBox.Show("Удалить выбранную строку?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.Cancel Then
            Exit Sub
        End If


        Dim dt As New DataTable
        dt.Columns.Add("Руб.")
        dt.Columns.Add("Коп.")
        dt.Columns.Add("Ед.изм")
        dt.Columns.Add("Вып.работа")
        dt.Columns.Add("Вып.работа (Имен.падеж)")

        dt = Grid6.DataSource

        dt.Rows.RemoveAt(Grid6.CurrentRow.Index)

        If Grid6.CurrentRow Is Nothing Then
            Exit Sub
        End If

        If TextBox52.Text = "" Then
            dt.Rows(Grid6.CurrentRow.Index).Item(1) = "00"
        Else
            If TextBox52.Text.Length < 2 And TextBox52.Text.Length > 0 Then
                dt.Rows(Grid6.CurrentRow.Index).Item(1) = "0" & TextBox52.Text
            Else
                dt.Rows(Grid6.CurrentRow.Index).Item(1) = TextBox52.Text
            End If

        End If


        Grid6.DataSource = dt
        GridView(Grid6)
        Grid6.Columns(2).Width = 60
        Grid6.Columns(0).Width = 60
        Grid6.Columns(1).Width = 60
        TextBox52.Text = ""
        TextBox22.Text = ""

    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click

        If ComboBox1.Text = "" Then Exit Sub

        If ComboBox19.InvokeRequired Or ComboBox26.InvokeRequired Then
            Me.Invoke(New comb19(AddressOf Ускорен))
        Else

            ComboBox19.AutoCompleteCustomSource.Clear()
            ComboBox19.Items.Clear()
            ComboBox26.Items.Clear()
            Dim var As List(Of Сотрудники)
            Using dbcx As New DbAll1DataContext
                var = (From x In dbcx.Сотрудники.AsEnumerable
                       Where x.НазвОрганиз = ComboBox1.Text
                       Order By x.ФИОСборное
                       Select x).ToList()
            End Using


            For Each r In var
                ComboBox19.AutoCompleteCustomSource.Add(r.ФИОСборное)
                ComboBox19.Items.Add(Trim(r.ФИОСборное & "" & r.ТипОтношения))
                ComboBox26.Items.Add(r.КодСотрудники)
            Next

            ComboBox19.Text = ""
        End If
    End Sub

    Private Sub CheckBox8_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox8.CheckedChanged

        If CheckBox8.Checked = False Then
            Exit Sub
        End If

        If CheckBox8.Checked = True Then
            CheckBox9.Checked = False
            CheckBox10.Checked = False
        End If




        If (TabControl1.TabPages.Contains(TabPage3) = True) Then
            TabControl1.TabPages.Remove(TabPage1)
            TabControl1.TabPages.Remove(TabPage2)
            TabControl1.TabPages.Remove(TabPage4)
            TabControl1.SelectTab(TabPage3)
        Else
            TabControl1.TabPages.Add(TabPage3)
            TabControl1.TabPages.Remove(TabPage1)
            TabControl1.TabPages.Remove(TabPage2)
            TabControl1.TabPages.Remove(TabPage4)
        End If
    End Sub

    Private Sub CheckBox10_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox10.CheckedChanged
        If CheckBox10.Checked = False Then
            Exit Sub
        End If

        If CheckBox10.Checked = True Then
            CheckBox9.Checked = False
            CheckBox8.Checked = False
        End If

        If (TabControl1.TabPages.Contains(TabPage2) = True) Then
            TabControl1.TabPages.Remove(TabPage1)
            TabControl1.TabPages.Remove(TabPage3)
            TabControl1.TabPages.Remove(TabPage4)
            TabControl1.SelectTab(TabPage2)
        Else
            TabControl1.TabPages.Add(TabPage2)
            TabControl1.TabPages.Remove(TabPage1)
            TabControl1.TabPages.Remove(TabPage3)
            TabControl1.TabPages.Remove(TabPage4)
        End If
    End Sub


    Private Function ПроверкаОформенСотрудникЧерезСправочник(ByVal _КодСотр As Integer) As Boolean
        Me.Cursor = Cursors.WaitCursor

        Using dbcx As New DbAll1DataContext  'определяем, оформлен ли сотрудник через справочник

            Dim var = (From x In dbcx.Сотрудники.AsEnumerable
                       Where x.КодСотрудники = _КодСотр
                       Select x.ДанныеИзСправочника).FirstOrDefault()
            If Not var = "True" Then
                Me.Cursor = Cursors.Default
                Return False
            End If
        End Using

        Dim f As String = ПроверкаДоговорИлиПодрядЧерезСправочник(_КодСотр)
        If f = "Контракт" Then
            TabControl1.TabPages.Remove(TabPage3)
            TabControl1.TabPages.Remove(TabPage1)
            TabControl1.TabPages.Remove(TabPage2)
            TabControl1.TabPages.Remove(TabPage4)
            TabControl1.TabPages.Add(TabPage2)
            TabControl1.SelectTab(TabPage2)

            CheckBox9.Enabled = True
            CheckBox10.Enabled = True
            CheckBox10.Checked = True
            CheckBox8.Enabled = True
            CheckBox7.Enabled = False

            Me.Cursor = Cursors.Default
            Return True
        ElseIf f = "Подряд(час)" Then
            TabControl1.TabPages.Remove(TabPage2)
            TabControl1.TabPages.Remove(TabPage1)
            TabControl1.TabPages.Remove(TabPage3)
            TabControl1.TabPages.Remove(TabPage4)
            TabControl1.TabPages.Add(TabPage3)
            TabControl1.SelectTab(TabPage3)
            CheckBox8.Checked = True

            Me.Cursor = Cursors.Default
            Return True
        Else
            If (TabControl1.TabPages.Contains(TabPage2) = True) Then
                TabControl1.TabPages.Remove(TabPage1)
                TabControl1.TabPages.Remove(TabPage3)
                TabControl1.TabPages.Remove(TabPage4)
                TabControl1.SelectTab(TabPage2)
            Else

                TabControl1.TabPages.Add(TabPage2)
                TabControl1.SelectTab(TabPage2)
                TabControl1.TabPages.Remove(TabPage1)
                TabControl1.TabPages.Remove(TabPage3)
                TabControl1.TabPages.Remove(TabPage4)
            End If
            CheckBox9.Enabled = True
            CheckBox10.Enabled = True
            CheckBox10.Checked = True
            CheckBox8.Enabled = True
            CheckBox7.Enabled = False

            Me.Cursor = Cursors.Default
            Return True
        End If



    End Function


    Private Sub Com19sel()
        ClAll()





        ComboBox26.Visible = False
        Label96.Text = ComboBox26.Items.Item(ComboBox19.SelectedIndex)
        'Label96.Text = ComboBox19.Items.Item(ComboBox19.SelectedIndex)
        combxS19 = ComboBox19.SelectedItem.ToString
        CheckBox26.Checked = False

        'проверка контракт или дог подряда по справочнику
        If ПроверкаОформенСотрудникЧерезСправочник(ComboBox26.Items.Item(ComboBox19.SelectedIndex)) = True Then
            Exit Sub
        End If

        'Заполняем сотрудника, паспорт и прописку
        загрПрил()
        Try
            КодСотрудника = CType(Label96.Text, Integer)
        Catch ex As Exception
            MessageBox.Show("Нет в базе идентификатора данного сотрудника", Рик)
            Exit Sub
        End Try




        Dim РазрИзменКонтр = (From x In dtKartochkaSotrudnikaAll.AsEnumerable Where Not IsDBNull(x.Item("IDСотр")) _
                                                                             AndAlso x.Item("IDСотр") = КодСотрудника
                              Select x.Item("Ставка")).LastOrDefault
        'РазрИзменКонтр = (From x In РазрИзменКонтр1.AsEnumerable Where x.Item("IDСотр") = КодСотрудника Select x.Item("Ставка")).LastOrDefault




        Dim f As Integer = УскорИзменСотрДог()
        If f = 1 Then Exit Sub
        УскорИзменСотр()
        УскорИзменСотрКарт() 'асинхрон
        ПропОклад()

        ''Соед(0)

    End Sub
    Public Sub ВстДанных(ByVal ds As DataTable)

        Dim ut() As Object = {"м2", "м3", "м.п."}

        Try
            ВидыРаботДогПодряда.ComboBox1.Items.Clear()
            ВидыРаботДогПодряда.ComboBox2.Items.Clear()
            ВидыРаботДогПодряда.ComboBox5.Items.Clear()
            ВидыРаботДогПодряда.ComboBox7.Items.Clear()
            ВидыРаботДогПодряда.ComboBox9.Items.Clear()
        Catch ex As Exception

        End Try

        ВидыРаботДогПодряда.GroupBox3.Enabled = False
        ВидыРаботДогПодряда.GroupBox4.Enabled = False
        ВидыРаботДогПодряда.GroupBox5.Enabled = False
        ВидыРаботДогПодряда.GroupBox6.Enabled = False

        ВидыРаботДогПодряда.ComboBox1.Items.AddRange(ut)
        ВидыРаботДогПодряда.ComboBox2.Items.AddRange(ut)
        ВидыРаботДогПодряда.ComboBox5.Items.AddRange(ut)
        ВидыРаботДогПодряда.ComboBox7.Items.AddRange(ut)
        ВидыРаботДогПодряда.ComboBox9.Items.AddRange(ut)

        'Dim strsql As String = "SELECT * FROM ДогПодОсобен"
        'Dim ds1 As DataTable = Selects(strsql)
        For i As Integer = 0 To ds.Rows.Count - 1
            Select Case i
                Case 0
                    ВидыРаботДогПодряда.CheckBox2.Checked = True
                    ВидыРаботДогПодряда.ComboBox2.Text = ds.Rows(i).Item(14).ToString
                    ВидыРаботДогПодряда.ComboBox3.Text = ds.Rows(i).Item(11).ToString
                    ВидыРаботДогПодряда.TextBox1.Text = ds.Rows(i).Item(12).ToString
                    ВидыРаботДогПодряда.TextBox2.Text = ds.Rows(i).Item(13).ToString
                    ВидыРаботДогПодряда.GroupBox3.Enabled = True

                Case 1
                    ВидыРаботДогПодряда.CheckBox1.Checked = True
                    ВидыРаботДогПодряда.ComboBox5.Text = ds.Rows(i).Item(14).ToString
                    ВидыРаботДогПодряда.ComboBox4.Text = ds.Rows(i).Item(11).ToString
                    ВидыРаботДогПодряда.TextBox4.Text = ds.Rows(i).Item(12).ToString
                    ВидыРаботДогПодряда.TextBox3.Text = ds.Rows(i).Item(13).ToString
                    ВидыРаботДогПодряда.GroupBox4.Enabled = True

                Case 2
                    ВидыРаботДогПодряда.CheckBox3.Checked = True
                    ВидыРаботДогПодряда.ComboBox7.Text = ds.Rows(i).Item(14).ToString
                    ВидыРаботДогПодряда.ComboBox6.Text = ds.Rows(i).Item(11).ToString
                    ВидыРаботДогПодряда.TextBox6.Text = ds.Rows(i).Item(12).ToString
                    ВидыРаботДогПодряда.TextBox5.Text = ds.Rows(i).Item(13).ToString
                    ВидыРаботДогПодряда.GroupBox5.Enabled = True
                Case 3
                    ВидыРаботДогПодряда.CheckBox4.Checked = True
                    ВидыРаботДогПодряда.ComboBox9.Text = ds.Rows(i).Item(14).ToString
                    ВидыРаботДогПодряда.ComboBox8.Text = ds.Rows(i).Item(11).ToString
                    ВидыРаботДогПодряда.TextBox8.Text = ds.Rows(i).Item(12).ToString
                    ВидыРаботДогПодряда.TextBox7.Text = ds.Rows(i).Item(13).ToString
                    ВидыРаботДогПодряда.GroupBox6.Enabled = True
            End Select

        Next

        Дпод1 = ""
        ВидыРаботДогПодряда.ShowDialog()








    End Sub
    Private Sub ЗакрытиеДляДогПодряда(ByVal d As Boolean)

        If d = True Then
            TextBox61.Visible = True
            TextBox62.Visible = True
            Label43.Visible = True
            Label97.Visible = True

        Else
            TextBox61.Visible = False
            TextBox62.Visible = False
            Label43.Visible = False
            Label97.Visible = False

        End If

    End Sub

    Private Sub ДанДогПодр(ByVal ID As Integer)
        Чист()
        StrSql = "Select DISTINCT НомерДогПодр From ДогПодряда Where ID = " & ID & ""
        Dim ds1 As DataTable = Selects(StrSql)
        Dim ds As DataTable

        If ds1.Rows.Count > 1 Then
            ДогПодномДогПод = ID
            ДогПодВыборНомДоговора.Flag = True
            ДогПодВыборНомДоговора.ShowDialog()
            Чист()
            StrSql = "Select * From ДогПодряда Where ID = " & ID & " and НомерДогПодр='" & ДогПодномДогПодНомДог & "'"
            ds = Selects(StrSql)
        Else
            Чист()
            StrSql = "Select * From ДогПодряда Where ID = " & ID & ""
            ds = Selects(StrSql)
        End If

        Try
            ДогПодНомерСтар = ds.Rows(0).Item(2).ToString
            ComboBox25.Text = ds.Rows(0).Item(9).ToString

            Dim leh As String = ds.Rows(0).Item(2).ToString
            Dim leh2 As String

            If leh.Length > 3 Then
                leh2 = Strings.Left(leh, 3)
                If leh.Length = 5 Then
                    TextBox55.Text = leh2
                    TextBox39.Text = Strings.Right(leh, 1)
                ElseIf leh.Length = 6 Then
                    TextBox55.Text = leh2
                    TextBox39.Text = Strings.Right(leh, 2)
                End If
            Else
                TextBox55.Text = ds.Rows(0).Item(2).ToString
                TextBox39.Text = ""
            End If

            MaskedTextBox6.Text = ds.Rows(0).Item(3).ToString
            ComboBox22.Text = ds.Rows(0).Item(4).ToString
            MaskedTextBox7.Text = ds.Rows(0).Item(5).ToString
            MaskedTextBox8.Text = ds.Rows(0).Item(6).ToString

            If ds.Rows(0).Item(7).ToString = "" Then
                TabControl1.SelectTab(TabPage3)
                ЗакрытиеДляДогПодряда(False)

                ВстДанных(ds)

                If CheckBox5.Checked = True And КрестикНажатиеДогПодряда = False Then
                    TextBox61.Text = ""
                    'Dim g As New Thread(Sub() Button1.PerformClick())
                    'g.SetApartmentState(ApartmentState.STA)
                    'g.IsBackground = True
                    'g.Start()
                    Button1.PerformClick()
                Else

                End If




                'Dim bdf As Boolean = Await Обход(ds)
                'Обход(ds)
            Else
                TextBox61.Text = ds.Rows(0).Item(7).ToString
                TextBox62.Text = ds.Rows(0).Item(8).ToString
                ComboBox27.Text = "час"
                ЗакрытиеДляДогПодряда(True)

            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try
        Try
            Примечание.RichTextBox1.Text = ds.Rows(0).Item(10).ToString
            Примечание.TextBox2.Text = Trim(TextBox1.Text) & " " & Trim(TextBox2.Text) & " " & Trim(TextBox3.Text)
            Примечани = ""
            Примечани = ds.Rows(0).Item(10).ToString
        Catch ex As Exception
            MessageBox.Show("Переоформите с этим сотрудников данные договора подряда", Рик)
        End Try

    End Sub
    Private Async Sub ДогПодВстВДопФорму(ByVal ds As DataTable)
        Dim ut() As Object = {"м2", "м3", "м.п."}
        Try
            ВидыРаботДогПодряда.ComboBox1.Items.Clear()
            ВидыРаботДогПодряда.ComboBox2.Items.Clear()
            ВидыРаботДогПодряда.ComboBox5.Items.Clear()
            ВидыРаботДогПодряда.ComboBox7.Items.Clear()
            ВидыРаботДогПодряда.ComboBox9.Items.Clear()
        Catch ex As Exception

        End Try

        ВидыРаботДогПодряда.GroupBox3.Enabled = False
        ВидыРаботДогПодряда.GroupBox4.Enabled = False
        ВидыРаботДогПодряда.GroupBox5.Enabled = False
        ВидыРаботДогПодряда.GroupBox6.Enabled = False



        ВидыРаботДогПодряда.ComboBox1.Items.AddRange(ut)
        ВидыРаботДогПодряда.ComboBox2.Items.AddRange(ut)
        ВидыРаботДогПодряда.ComboBox5.Items.AddRange(ut)
        ВидыРаботДогПодряда.ComboBox7.Items.AddRange(ut)
        ВидыРаботДогПодряда.ComboBox9.Items.AddRange(ut)

        'Dim strsql As String = "SELECT * FROM ДогПодОсобен"
        'Dim ds1 As DataTable = Selects(strsql)
        For i As Integer = 0 To ds.Rows.Count - 1
            Select Case i
                Case 0
                    ВидыРаботДогПодряда.CheckBox2.Checked = True
                    ВидыРаботДогПодряда.ComboBox3.Text = ds.Rows(i).Item(11).ToString
                    ВидыРаботДогПодряда.ComboBox2.Text = ds.Rows(i).Item(14).ToString
                    ВидыРаботДогПодряда.TextBox1.Text = ds.Rows(i).Item(12).ToString
                    ВидыРаботДогПодряда.TextBox2.Text = ds.Rows(i).Item(13).ToString
                    ВидыРаботДогПодряда.GroupBox3.Enabled = True

                Case 1
                    ВидыРаботДогПодряда.CheckBox1.Checked = True
                    ВидыРаботДогПодряда.ComboBox4.Text = ds.Rows(i).Item(11).ToString
                    ВидыРаботДогПодряда.ComboBox5.Text = ds.Rows(i).Item(14).ToString
                    ВидыРаботДогПодряда.TextBox4.Text = ds.Rows(i).Item(12).ToString
                    ВидыРаботДогПодряда.TextBox3.Text = ds.Rows(i).Item(13).ToString
                    ВидыРаботДогПодряда.GroupBox4.Enabled = True

                Case 2
                    ВидыРаботДогПодряда.CheckBox3.Checked = True
                    ВидыРаботДогПодряда.ComboBox6.Text = ds.Rows(i).Item(11).ToString
                    ВидыРаботДогПодряда.ComboBox7.Text = ds.Rows(i).Item(14).ToString
                    ВидыРаботДогПодряда.TextBox6.Text = ds.Rows(i).Item(12).ToString
                    ВидыРаботДогПодряда.TextBox5.Text = ds.Rows(i).Item(13).ToString
                    ВидыРаботДогПодряда.GroupBox5.Enabled = True
                Case 3
                    ВидыРаботДогПодряда.CheckBox4.Checked = True
                    ВидыРаботДогПодряда.ComboBox8.Text = ds.Rows(i).Item(11).ToString
                    ВидыРаботДогПодряда.ComboBox9.Text = ds.Rows(i).Item(14).ToString
                    ВидыРаботДогПодряда.TextBox8.Text = ds.Rows(i).Item(12).ToString
                    ВидыРаботДогПодряда.TextBox7.Text = ds.Rows(i).Item(13).ToString
                    ВидыРаботДогПодряда.GroupBox6.Enabled = True
            End Select
        Next
        'ComboBox27.Text = "иное"
        ЗакрытиеДляДогПодряда(False)


        Await Task.Delay(1000)
    End Sub


    Private Sub TextBox38_TextChanged(sender As Object, e As EventArgs) Handles TextBox38.TextChanged
        If TextBox38.Text = "" Then
            Label82.ForeColor = Color.Red
            Label82.Text = "NO"
        Else
            Label82.ForeColor = Color.Green
            Label82.Text = "OK"
        End If
    End Sub

    Private Sub ComboBox20_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox20.SelectedIndexChanged
        'Контракты(ComboBox20.Text)
        Dim l = listFluentFTP("/" & ComboBox1.Text & "/" & "Контракт" & "/" & ComboBox20.Text & "/")

        ComboBox2.Items.Clear()
        ComboBox2.Text = ""
        For x As Integer = 0 To l.Count - 1
            ComboBox2.Items.Add(l(x).ToString)
        Next



    End Sub

    Private Sub TextBox41_TextChanged(sender As Object, e As EventArgs) Handles TextBox41.TextChanged
        If TextBox41.Text = "" Then
            Label83.ForeColor = Color.Red
            Label83.Text = "NO"
        Else
            Label83.ForeColor = Color.Green
            Label83.Text = "OK"
        End If
    End Sub
    Private Sub ДогПодрЗаполн()
        If ComboBox1.Text = "" Then
            MessageBox.Show("Выберите организацию!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            CheckBox7.Checked = False
            Exit Sub
        End If
        If CheckBox7.Checked = True Then
            CheckBox7.Enabled = True
            CheckBox23.Enabled = False
            CheckBox8.Checked = False
            CheckBox9.Checked = False
            CheckBox10.Checked = False
            CheckBox8.Enabled = False
            CheckBox9.Enabled = False
            CheckBox10.Enabled = False
        Else
            CheckBox23.Enabled = True
            GroupBox19.Visible = False
            dtn2ClickAsyns()
        End If
        'Соед(0)

        'Чист()
        'StrSql = "SELECT Должность FROM ДогПодДолжн WHERE Клиент='" & ComboBox1.Text & "'"
        'ds = Selects(StrSql)
        Dim ds1
        dbcx = New DbAll1DataContext
        ds1 = From x In dbcx.ДогПодДолжн Where x.Клиент = ComboBox1.Text
              Order By x.Должность
              Select x.Должность, x.Код

        ComboBox22.DataSource = ds1
        ComboBox22.DisplayMember = "Должность"
        ComboBox22.ValueMember = "Код"
        'Dim ds1 = dtDogPodrDoljnostAll.Select("Клиент='" & ComboBox1.Text & "'")

        'Me.ComboBox22.AutoCompleteCustomSource.Clear()
        'Me.ComboBox22.Items.Clear()
        'For Each r As DataRow In ds1
        '    Me.ComboBox22.AutoCompleteCustomSource.Add(r.Item("Должность").ToString())
        '    Me.ComboBox22.Items.Add(r("Должность").ToString)
        'Next

        'Чист()
        'StrSql = "SELECT АдресОбъекта FROM ОбъектОбщепита WHERE НазвОрг='" & ComboBox1.Text & "'"
        'ds = Selects(StrSql)

        'Соед(0)
        Dim ds = dtObjectObshepitaAll.Select("НазвОрг='" & ComboBox1.Text & "'")

        Me.ComboBox25.Items.Clear()
        For Each r As DataRow In ds
            Me.ComboBox25.Items.Add(r("АдресОбъекта").ToString)
        Next
        ComboBox25.Text = ds(0).Item("АдресОбъекта").ToString

        If CheckBox5.Checked = False Then
            MaskedTextBox6.Text = Now
            MaskedTextBox7.Text = Now
            MaskedTextBox8.Text = Now
        End If

        If CheckBox7.Checked = True Then


            TabControl1.TabPages.Add(TabPage3)
            TabControl1.TabPages.Remove(TabPage2)
            'TabControl1.SelectedTab = TabControl1.TabPages("Договор подряда")

            If (TabControl1.TabPages.Contains(TabPage1) = False) Then
                TabControl1.TabPages.Insert(0, TabPage1)
                TabControl1.SelectTab(TabPage1)
            End If
            'TabControl1.SelectTab(TabPage1) Then


            СозданиепапкиНаСервере(ComboBox1.Text & "/Договор подряда/" & Now.Year & "/")   'создание папки (если вдруг нет)
            Dim listCombo2 As List(Of String) = listFluentFTP(ComboBox1.Text & "/Договор подряда/" & Now.Year & "/")

            If listCombo2.Count = 0 Then
                TabControl1.SelectTab(TabPage1)
                ComboBox24.Text = ""
                ComboBox24.Text = "В базе нет договоров подряда организации " & ComboBox1.Text
                ComboBox24.Enabled = False
                ComboBox23.Enabled = False

                Dim ut3() As Object = {"час", "иное"}
                Try
                    ComboBox27.Items.Clear()
                Catch ex2 As Exception

                End Try
                ComboBox27.Items.AddRange(ut3)
                Exit Sub
            End If
            'приостановлено 01.07.19 (пока настраивается сервер)
            Dim listCombo3 As Object = listFluentFTP(ComboBox1.Text & "/Договор подряда/")
            ComboBox23.Items.Clear()  'года
            For Each item In listCombo3
                ComboBox23.Items.Add(item.ToString)
            Next
            'ComboBox23.Text = Now.Year


            ComboBox24.Items.Clear()
            Dim m As String = FTPString & ComboBox1.Text & "/Договор подряда/" & Now.Year & "/"

            For x As Integer = 0 To listCombo2.Count - 1
                ComboBox24.Items.Add(Replace(listCombo2(x).ToString, m, ""))
            Next


        Else
            TabControl1.TabPages.Remove(TabPage3)
            TabControl1.TabPages.Add(TabPage2)
            ComboBox24.Text = ""
        End If

        ComboBox24.Enabled = True
        ComboBox23.Enabled = True

        Dim ut() As Object = {"час", "иное"}
        Try
            ComboBox27.Items.Clear()
        Catch ex As Exception

        End Try
        ComboBox27.Items.AddRange(ut)



        If CheckBox7.Checked = True And CheckBox5.Checked = False Then
            ЗакрытиеДляДогПодряда(False)
        End If
    End Sub
    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged


        ДогПодрЗаполн()
        'End If
    End Sub

    Private Sub TextBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox7.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.MaskedTextBox1.Focus()
        End If
    End Sub

    Private Sub textbox55_TextChanged(sender As Object, e As EventArgs) Handles TextBox55.TextChanged
        If TextBox55.ForeColor = Color.Red Then
            TextBox55.ForeColor = Color.Black
        End If

    End Sub

    Private Sub TextBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox9.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True

            Dim sd As String = Strings.UCase(Strings.Left(TextBox9.Text, 1))
            Try
                TextBox9.Text = sd & Strings.Right(TextBox9.Text, (TextBox9.TextLength - 1))
                Me.TextBox8.Focus()
            Catch ex As Exception
                MessageBox.Show("Введите правильно данные!", Рик)
                TextBox9.Focus()

            End Try


        End If
    End Sub
    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown

        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True

            If Not TextBox8.TextLength = 14 And CheckBox1.Checked = False Then
                MessageBox.Show("Неправильно заполнен идентификационный номер паспорта!", Рик)
                TextBox8.Focus()
                Exit Sub
            End If
            TextBox45.Focus()
        End If
    End Sub

    Private Sub TextBox19_TextChanged(sender As Object, e As EventArgs) Handles TextBox19.TextChanged

    End Sub

    Private Sub ComboBox12_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox12.SelectedIndexChanged
        РасчПер()

    End Sub

    Private Sub TextBox45_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox45.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            TabControl1.SelectedTab = TabPage2
            ComboBox8.Focus()
        End If





    End Sub
    Private Sub ComboBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox8.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox9.Focus()
        End If
    End Sub
    Private Sub ComboBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox9.KeyDown
        If ComboBox7.Visible = True Then
            If e.KeyCode = Keys.Enter Then
                e.SuppressKeyPress = True
                Me.ComboBox7.Focus()
            ElseIf e.KeyCode = Keys.Enter Then
                e.SuppressKeyPress = True
                Me.ComboBox10.Focus()
            End If
        End If
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        GroupBox19.Visible = True
        Dim db As New DbAll1DataContext

        Dim dt = (From x In db.ДогПодДолжн
                  From y In db.ДогПодрОбязан
                  Where x.Должность = ComboBox22.Text And x.Клиент = ComboBox1.Text And y.ID = x.Код
                  Select New With {.Идентификатор = x.Код, .Организация = x.Клиент, x.Должность, y.Обязанности}).ToList()
        Grid1.DataSource = dt
        GridView(Grid1)

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        If ComboBox1.Text = "" Then Exit Sub
        Dim r As New Random()
        Dim Mass As Integer() = Enumerable.Range(0, 50).OrderBy(Function(n) r.Next).Take(25).ToArray()
        TextBox1.Text = Path.GetRandomFileName()
        TextBox2.Text = "text1"
        TextBox3.Text = "text2"
        TextBox21.Text = "Прописка"
        TextBox12.Text = "vv"
        TextBox7.Text = 7777777
        MaskedTextBox1.Text = Now.Date
        MaskedTextBox2.Text = Now.Date
        TextBox9.Text = Path.GetRandomFileName()
        TextBox8.Text = CType(Mass(0) & Mass(1) & Mass(2) & Mass(3) & Mass(4) & Mass(5) & Mass(6), String)
        TextBox45.Text = CType(Mass(0) & Mass(1) & Mass(2) & Mass(3) & Mass(4) & Mass(5) & Mass(6), String)
        ComboBox11.SelectedIndex = 1
        ComboBox15.SelectedIndex = 1
        ComboBox18.SelectedIndex = 0
        ComboBox8.SelectedIndex = 0
        ComboBox9.SelectedIndex = 0
        ComboBox10.SelectedIndex = 1
        TextBox38.Text = 15
        TextBox41.Text = 25
        TextBox40.Text = 15
        CheckBox4.Checked = True


        If TextBox8.Text.Length < 14 Then
            Dim y = 14 - TextBox8.Text.Length
            If y = 1 Then
                TextBox8.Text &= "6"
            ElseIf 2 Then
                TextBox8.Text &= "21"
            Else
                TextBox8.Text &= "216"
            End If
        End If

        If TextBox45.Text.Length < 14 Then
            Dim y = 14 - TextBox8.Text.Length
            If y = 1 Then
                TextBox45.Text &= "6"
            ElseIf 2 Then
                TextBox45.Text &= "21"
            Else
                TextBox45.Text &= "216"
            End If
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ComboBox16_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox16.SelectedIndexChanged
        РасчПер()

    End Sub
    Private Sub ComboBox19_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox19.SelectedIndexChanged


        Комбы()

        Com19ForДогПодр = ComboBox19.SelectedItem
        ДогПодрВклЧекбокс5 = True

        If ВидыРаботДогПодряда.ComboBox3.Text <> "" Or ВидыРаботДогПодряда.TextBox1.Text <> "" Then
            'ВидыРаботДогПодряда.ОчисВидыРаботДогПодряда()
            'Dim ftask As New Thread(AddressOf ВидыРаботДогПодряда.ОчисВидыРаботДогПодряда)
            'ftask.IsBackground = True
            'ftask.SetApartmentState(ApartmentState.STA)
            'ftask.Start()

            ВидыРаботДогПодряда.CheckBox2.Checked = True
            ВидыРаботДогПодряда.CheckBox1.Checked = True
            ВидыРаботДогПодряда.CheckBox3.Checked = True
            ВидыРаботДогПодряда.CheckBox4.Checked = True
            ВидыРаботДогПодряда.CheckBox2.Checked = False
            ВидыРаботДогПодряда.CheckBox1.Checked = False
            ВидыРаботДогПодряда.CheckBox3.Checked = False
            ВидыРаботДогПодряда.CheckBox4.Checked = False
            ВидыРаботДогПодряда.Очистка(ВидыРаботДогПодряда)


        End If

        Com19sel()

    End Sub
    Private Sub ComboBox27_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox27.SelectedIndexChanged

        СортДогПод(ComboBox27.Text)

        If ComboBox27.Text = "час" And CheckBox5.Checked = False Then
            ЗакрытиеДляДогПодряда(True)
            TextBox61.Text = ""
            TextBox62.Text = ""
        ElseIf ComboBox27.Text = "иное" And CheckBox5.Checked = False Then
            ЗакрытиеДляДогПодряда(False)
            TextBox61.Text = ""
            TextBox62.Text = ""
        ElseIf ComboBox27.Text = "час" And CheckBox5.Checked = True And ПрЗакрВидыРаб = ComboBox19.SelectedItem Then
            ЗакрытиеДляДогПодряда(True)
            TextBox61.Text = ""
            TextBox62.Text = ""
        ElseIf ComboBox27.Text = "иное" And CheckBox5.Checked = True And ПрЗакрВидыРаб = ComboBox19.SelectedItem Then
            Dim номдог As String
            If TextBox39.Text <> "" Then
                номдог = TextBox55.Text & "." & TextBox39.Text
            Else
                номдог = TextBox55.Text
            End If

            If Not номдог = ДогПодНомерСтар And КрестикНажатиеДогПодряда = False Then
                ЗакрытиеДляДогПодряда(False)
                TextBox61.Text = ""
                TextBox62.Text = ""
                Button1.PerformClick()
            Else
                ЗакрытиеДляДогПодряда(False)
                TextBox61.Text = ""
                TextBox62.Text = ""

            End If

        End If

    End Sub
    Private Sub TextBox34_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox34.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox11.Focus()
        End If
    End Sub

    Private Sub TextBox11_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox11.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox10.Focus()
        End If
    End Sub

    Private Sub TextBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox10.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox21.Focus()
        End If
    End Sub

    Private Sub ComboBox19_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox19.KeyDown

    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        If CheckBox5.Checked = False Then
            comb7()
        ElseIf CheckBox5.Checked = True And CheckBox26.Checked = True Then
            comb7()
        End If

        If ComboBox7.Text <> "" Then
            Label79.ForeColor = Color.Green
            Label79.Text = "OK"
        Else
            Label79.ForeColor = Color.Red
            Label79.Text = "NO"
        End If

    End Sub



    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If ComboBox1.Text = "" Then
            MsgBox("Выберите организацию",, "ООО РикКонсалтинг")
            Me.ComboBox1.Focus()

            Exit Sub
        End If

        'If ComboBox1.Text <> "" And CheckBox5.Checked = True And ComboBox19.SelectedItem = "" Then
        '    CheckBox5.Checked = False
        'End If


        sender.text = StrConv(sender.text, VbStrConv.ProperCase)
        sender.SelectionStart = sender.text.Length
        TextBox6.Text = TextBox1.Text
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox2.Focus()
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim strsql As String
        Dim ds As DataTable

        If CheckBox5.Checked = True Then
            strsql = "SELECT * FROM ПродлКонтракта WHERE IDСотр=" & CType(Label96.Text, Integer) & ""
            ds = Selects(strsql)
            proverka = 1
            Try
                УведомлениеФорма.TextBox1.Text = ds.Rows(0).Item(2).ToString
                УведомлениеФорма.TextBox2.Text = ds.Rows(0).Item(3).ToString
                УведомлениеФорма.TextBox3.Text = ds.Rows(0).Item(4).ToString
                УведомлениеФорма.TextBox4.Text = ds.Rows(0).Item(5).ToString
                УведомлениеФорма.TextBox5.Text = ds.Rows(0).Item(9).ToString
                УведомлениеФорма.TextBox6.Text = ds.Rows(0).Item(8).ToString
                УведомлениеФорма.TextBox7.Text = ds.Rows(0).Item(7).ToString
                УведомлениеФорма.TextBox8.Text = ds.Rows(0).Item(10).ToString
                УведомлениеФорма.TextBox9.Text = ds.Rows(0).Item(14).ToString
                УведомлениеФорма.TextBox10.Text = ds.Rows(0).Item(13).ToString
                УведомлениеФорма.TextBox11.Text = ds.Rows(0).Item(12).ToString
                УведомлениеФорма.TextBox12.Text = ds.Rows(0).Item(11).ToString
                УведомлениеФорма.TextBox13.Text = ds.Rows(0).Item(18).ToString
                УведомлениеФорма.TextBox14.Text = ds.Rows(0).Item(17).ToString
                УведомлениеФорма.TextBox15.Text = ds.Rows(0).Item(16).ToString
                УведомлениеФорма.TextBox16.Text = ds.Rows(0).Item(15).ToString
                УведомлениеФорма.TextBox17.Text = ds.Rows(0).Item(22).ToString
                УведомлениеФорма.TextBox18.Text = ds.Rows(0).Item(21).ToString
                УведомлениеФорма.TextBox19.Text = ds.Rows(0).Item(20).ToString
                УведомлениеФорма.TextBox20.Text = ds.Rows(0).Item(19).ToString
                УведомлениеФорма.TextBox21.Text = TextBox38.Text
                УведомлениеФорма.ShowDialog()
            Catch ex As Exception
                MessageBox.Show("Нет данных в базе!", Рик)
                Exit Sub
            End Try

        Else
            MessageBox.Show("Данные доступны только при 'Изменении данных сотрудника!'", Рик)
        End If
    End Sub

    Private Sub Данные()
        Dim StrSql4 As String = "SELECT ШтОтделы.Отделы, ШтСвод.Должность, ШтСвод.Разряд, ШтСвод.ТарифнаяСтавка,
ШтСвод.ПовышениеПроц, ШтСвод.ТарСтПослеИспСрока, ПовПроцПослеИспСрока, КодШтСвод
FROM ШтОтделы INNER JOIN ШтСвод ON ШтОтделы.Код = ШтСвод.Отдел
WHERE ШтОтделы.Отделы='" & Отдел & "' AND ШтСвод.Должность='" & Должность & "' AND ШтОтделы.Клиент='" & Клиент & "' AND ШтСвод.Разряд='" & ComboBox7.Text & "'"
        Dim ds5 As DataTable = Selects(StrSql4)

        If errds = 1 Then Exit Sub

        ds5 = ПроверкаИзмененияТарифнойСтавки2(ds5, MaskedTextBox3.Text)


        TextBox46.Text = ds5.Rows(0).Item(4).ToString()

        Dim dstbl As String = ds5.Rows(0).Item(3).ToString

        If dstbl <> "." Then dstbl = Replace(dstbl, ".", ",")
        If dstbl <> "," Then
            sf = Nothing
            sf = CType(dstbl, Double)
            Dim sfd As String = CType(sf, String)
            Dim ДлНач As Integer = sfd.Length
            TextBox33.Text = Math.Floor(sf)
            Dim Дл As Integer = TextBox33.TextLength
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

            TextBox44.Text = vm
        Else
            TextBox33.Text = ds.Rows(0).Item(0).ToString
        End If

        ПропОклад()
    End Sub
    Private Function ПроверкаИзмененияТарифнойСтавки2(ByVal dsin As DataTable, ByVal datex As String) As DataTable


        Dim ds1 As DataTable = Selects(StrSql:="SELECT * FROM ШтСводИзмСтавка WHERE IDКодШтСвод=" & CType(dsin.Rows(0).Item(7).ToString, Integer) & "")
        Dim DateEx As Date = CDate(datex)
        If Not errds = 1 Then
            ds1.DefaultView.Sort = "Дата DESC"
            ds1 = ds1.DefaultView.ToTable()

            For x As Integer = 0 To ds1.Rows.Count - 1
                If DateEx >= ds1.Rows(x).Item(2) Then
                    dsin.Rows(0).Item(3) = ds1.Rows(x).Item(3)
                    Return dsin
                End If
            Next
        End If
        Return dsin
    End Function
    Private Sub comb7()

        'If CombBox7 = 1 Then
        '    CombBox7 = 0
        '    Exit Sub
        'End If

        Данные()
        Exit Sub

        If Not ComboBox7.Text = "" Then
            Разряд = ComboBox7.Text
        End If

        If Отдел <> "" And Должность <> "" And ComboBox7.Text <> "" Then
            '            StrSql = "Select ШтСвод.ТарифнаяСтавка,ШтСвод.ПовышениеПроц
            'From ШтОтделы INNER Join ШтСвод On ШтОтделы.Код = ШтСвод.Отдел
            'Where ШтОтделы.Отделы ='" & Отдел & "' AND ШтСвод.Должность='" & Должность & "' AND ШтСвод.Разряд='" & Разряд & "'"
            'Соед(0)


            Dim ds As DataTable = Selects(StrSql:="Select  ШтСвод.ТарифнаяСтавка, ШтСвод.ПовышениеПроц, КодШтСвод
From ШтОтделы INNER Join ШтСвод On ШтОтделы.Код = ШтСвод.Отдел
Where ШтОтделы.Отделы ='" & Отдел & "' AND ШтСвод.Должность = '" & Должность & "' AND ШтСвод.Разряд='" & Разряд & "' AND ШтОтделы.Клиент = '" & Клиент & "'")

            ds = ПроверкаИзмененияТарифнойСтавки2(ds, MaskedTextBox3.Text)

            Label79.ForeColor = Color.Green
            Label79.Text = "OK"
            Try
                Очистка()
                Dim dstbl As String = ds.Rows(0).Item(0).ToString

                If dstbl <> "." Then dstbl = Replace(dstbl, ".", ",")
                If dstbl <> "," Then
                    sf = Nothing
                    sf = CType(dstbl, Double)
                    Dim sfd As String = CType(sf, String)
                    Dim ДлНач As Integer = sfd.Length
                    TextBox33.Text = Math.Floor(sf)
                    Dim Дл As Integer = TextBox33.TextLength
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

                    TextBox44.Text = vm
                Else
                    TextBox33.Text = ds.Rows(0).Item(0).ToString
                End If

                TextBox46.Text = ds.Rows(0).Item(1).ToString
            Catch ex As Exception
                Label79.ForeColor = Color.Red
                Label79.Text = "NO"
                MessageBox.Show("Нет данных в базе, относительно разряда!!!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            End Try
        Else
            If ComboBox7.Items.Count > 1 Then
                Me.TextBox33.Text = ""
                Me.TextBox43.Text = ""
                Me.TextBox46.Text = ""
                Me.TextBox48.Text = ""
                Me.TextBox47.Text = ""
                Me.TextBox44.Text = ""
                Dim раз As String = ""

                Dim StrSql2 As String = "Select  ШтСвод.ТарифнаяСтавка, ШтСвод.ПовышениеПроц, КодШтСвод
From ШтОтделы INNER Join ШтСвод On ШтОтделы.Код = ШтСвод.Отдел
Where ШтОтделы.Отделы ='" & Отдел & "' AND ШтСвод.Должность = '" & Должность & "' AND ШтСвод.Разряд='" & раз & "' AND ШтОтделы.Клиент = '" & Клиент & "'"
                Dim ds As DataTable = Selects(StrSql2)

                ds = ПроверкаИзмененияТарифнойСтавки2(ds, MaskedTextBox3.Text)

                Label79.ForeColor = Color.Green
                Label79.Text = "OK"
                Try
                    Me.TextBox33.Text = ""
                    Me.TextBox43.Text = ""
                    Me.TextBox46.Text = ""
                    Me.TextBox48.Text = ""
                    Me.TextBox47.Text = ""
                    Me.TextBox44.Text = ""
                    Dim dstbl As String = ds.Rows(0).Item(0).ToString

                    If dstbl <> "." Then dstbl = Replace(dstbl, ".", ",")
                    If dstbl <> "," Then
                        sf = Nothing
                        sf = CType(dstbl, Double)
                        Dim sfd As String = CType(sf, String)
                        Dim ДлНач As Integer = sfd.Length
                        TextBox33.Text = Math.Floor(sf)
                        Dim Дл As Integer = TextBox33.TextLength
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

                        TextBox44.Text = vm
                    Else
                        TextBox33.Text = ds.Rows(0).Item(0).ToString
                    End If

                    TextBox46.Text = ds.Rows(0).Item(1).ToString
                Catch ex As Exception
                    Label79.ForeColor = Color.Red
                    Label79.Text = "NO"
                    MessageBox.Show("Нет данных в базе, относительно разряда!!!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                End Try

            Else
                Me.ComboBox7.Enabled = False
                Label47.Enabled = False
                Label79.Enabled = False

                TextBox46.Text = ПовышениеПроц

                Dim dstbl As String = ТарифнаяСт

                If dstbl <> "." Then dstbl = Replace(dstbl, ".", ",")
                If dstbl <> "," Then
                    sf = Nothing
                    sf = CType(dstbl, Double)
                    TextBox33.Text = Math.Floor(sf)
                    Dim vm As String = Strings.Right(Math.Round(sf - Math.Floor(sf), 2), 2)
                    Dim vmn As String = CType(vm, Integer)
                    If vmn = "0" Then vm = Strings.Right(vm, 1) & "0"
                    TextBox44.Text = vm
                Else
                    TextBox33.Text = Отделы
                End If

            End If


        End If
    End Sub
    Private Sub загрПрил()
        StrSql = ""
        StrSql = "SELECT Примечание FROM КарточкаСотрудника WHERE IDСотр=" & CType(Label96.Text, Integer) & ""
        Dim ds7 As DataTable
        ds7 = Selects(StrSql)
        Примечание.RichTextBox1.Text = ""
        Try
            Примечание.RichTextBox1.Text = ds7.Rows(0).Item(0).ToString
            Примечани = ""
            Примечани = ds7.Rows(0).Item(0).ToString
        Catch ex As Exception

        End Try


    End Sub
    Private Sub Button5_Click_1(sender As Object, e As EventArgs) Handles Button5.Click
        If CheckBox5.Checked = False Then
            Прим = 1
            Примечание.ShowDialog()
        Else

            Прим = 0
            Примечание.ShowDialog()
        End If

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        If CheckBox5.Checked = False And CheckBox7.Checked = True Then
            Прим = 1
            Примечание.ShowDialog()
        ElseIf CheckBox5.Checked = True And CheckBox7.Checked = True Then
            Прим = 1
            Примечание.ShowDialog()
        End If
    End Sub



    Private Sub CheckBox25_CheckedChanged(sender As Object, e As EventArgs)
        'Dim IDLДогПодрОбяз2 As Integer
        'If CheckBox25.Checked = True Then

        '    'Соед(0)

        '    Dim StrSql As String = "SELECT Код FROM ДогПодДолжн WHERE Клиент='" & ComboBox1.Text & "' AND Должность= '" & ComboBox22.Text & "'"

        '    Dim ds As New DataSet
        '    Dim da As SqlDataAdapter = Доработчик(StrSql)
        '    Try
        '        da.Fill(ds, "Cn")
        '        IDLДогПодрОбяз2 = ds.Tables("cn").Rows(0).Item(0)
        '    Catch ex As Exception
        '        MessageBox.Show("Выберите должность!", Рик)
        '        'Соед(0)

        '        Exit Sub
        '    End Try

        '    If MessageBox.Show("Вы уверены что хотите удалить должность " & ComboBox22.Text & " ?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Hand) = DialogResult.OK Then

        '        Updates(stroka:="delete FROM ДогПодДолжн WHERE Код=" & IDLДогПодрОбяз2 & "")
        '        MessageBox.Show("Должность удалена!", Рик)
        '        refrdoljn()
        '    Else
        '        'Соед(0)

        '        Exit Sub

        '    End If

        'End If

        'Соед(0)

    End Sub
    Private Sub Com8sel()

        Me.ComboBox7.Visible = True
        Label47.Visible = True
        Label79.Visible = True
        Me.ComboBox7.Text = ""

        Me.TextBox33.Clear()
        Me.TextBox43.Clear()
        Me.TextBox44.Clear()
        Me.TextBox47.Clear()
        Me.TextBox48.Clear()
        Me.TextBox46.Clear()
        '''Соед(0)
        'Клиент
        Отдел = ComboBox8.Text




        Dim ds As DataTable = Selects(StrSql:="SELECT DISTINCT ШтСвод.Должность FROM ШтОтделы INNER JOIN ШтСвод ON ШтОтделы.Код = ШтСвод.Отдел
WHERE ШтОтделы.Клиент='" & Клиент & "' AND ШтОтделы.Отделы ='" & Отдел & "'")

        Me.ComboBox9.AutoCompleteCustomSource.Clear()
        Me.ComboBox9.Items.Clear()
        For Each r As DataRow In ds.Rows
            Me.ComboBox9.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox9.Items.Add(r(0).ToString)
        Next
        Me.ComboBox9.Text = ""
        Me.ComboBox9.Text = String.Empty
        Me.ComboBox7.Visible = True
        '''Соед(0)
    End Sub

    Private Sub ComboBox8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox8.SelectedIndexChanged
        'If conn.State = ConnectionState.Closed Then
        '    'Соед(0)
        'End If

        Com8sel()

    End Sub

    Private Sub CheckBox26_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox26.CheckedChanged
        If CheckBox26.Checked = True Then
            ComboBox9.Enabled = True
            ComboBox8.Enabled = True
            ComboBox7.Enabled = True
        Else
            ComboBox9.Enabled = False
            ComboBox8.Enabled = False
            ComboBox7.Enabled = False
            ComboBox9.Text = String.Empty
            ComboBox8.Text = String.Empty
            ComboBox7.Text = String.Empty
            'Соед(0)
            'Com19sel()
            'Соед(0)
        End If
    End Sub

    Private Sub ComboBox28_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox28.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            If GroupBox27.Visible = True Then
                MaskedTextBox9.Focus()
            Else
                TextBox9.Focus()
            End If

        End If
    End Sub

    Private Sub TextBox51_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox51.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            TextBox12.Focus()
        End If
    End Sub

    Private Sub ComboBox21_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox21.SelectedIndexChanged
        Dim l = listFluentFTP("/" & ComboBox1.Text & "/" & "Приказ" & "/" & ComboBox21.Text & "/")

        ComboBox17.Items.Clear()
        ComboBox17.Text = ""
        For x As Integer = 0 To l.Count - 1
            ComboBox17.Items.Add(l(x).ToString)
        Next
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ДатВыплЗпАванс.TextBox1.Text = Me.TextBox40.Text
        ДатВыплЗпАванс.TextBox2.Text = Me.TextBox56.Text
        ДатВыплЗпАванс.Show()
    End Sub



    'Private Sub TextBox40_SelectedValueChanged(sender As Object, e As EventArgs) Handles TextBox40.SelectedValueChanged
    '    If TextBox40.Text <> "" Then
    '        Label89.ForeColor = Color.Green
    '        Label89.Text = "OK"
    '    Else
    '        Label89.ForeColor = Color.Red
    '        Label89.Text = "NO"

    '    End If
    'End Sub

    Private Sub TextBox40_TextChanged(sender As Object, e As EventArgs) Handles TextBox40.TextChanged

        If TextBox40.Text <> "" Then
            Label89.ForeColor = Color.Green
            Label89.Text = "OK"
        Else
            Label89.ForeColor = Color.Red
            Label89.Text = "NO"

        End If


    End Sub

    Private Sub TextBox56_TextChanged(sender As Object, e As EventArgs) Handles TextBox56.TextChanged

        If TextBox56.Text <> "" Then
            Label90.ForeColor = Color.Green
            Label90.Text = "OK"
        Else
            Label90.ForeColor = Color.Red
            Label90.Text = "NO"

        End If


    End Sub

    Private Sub ComboBox18_TextChanged(sender As Object, e As EventArgs) Handles ComboBox18.TextChanged
        If CheckBox5.Checked = True Then
            If ComboBox18.Text = "" Then
                Label85.ForeColor = Color.Red
                Label85.Text = "NO"
            Else
                Label85.ForeColor = Color.Green
                Label85.Text = "OK"
            End If
        End If
    End Sub


    Private Sub РасчПер()
        Dim часы As Decimal = Val(ComboBox16.Text) 'расчет времени обеда и конца рабочего дня
        Dim ВрНач As Decimal = Val(ComboBox12.Text)
        'часы = Math.Floor(часы)
        Select Case часы
            Case 9
                If Not (ВрНач = 8.3 Or ВрНач = 10.3) Then
                    TextBox50.Text = Str(часы + ВрНач) & ".00"
                    Dim с As String = Str(4 + ВрНач)
                    Dim по As String = Str(4 + ВрНач + 1)
                    TextBox49.Text = "с" & с & ".00 до" & по & ".00"
                Else
                    TextBox50.Text = Str(часы + ВрНач) & "0"
                    Dim с As String = Str(4 + ВрНач)
                    Dim по As String = Str(4 + ВрНач + 1)
                    TextBox49.Text = "с" & с & "0 до" & по & "0"
                End If

            Case 10
                If Not (ВрНач = 8.3 Or ВрНач = 10.3) Then
                    TextBox50.Text = Str(часы + ВрНач) & ".00"
                    Dim с As String = Str(5 + ВрНач)
                    Dim по As String = Str(5 + ВрНач + 1)
                    TextBox49.Text = "с" & с & ".00 до" & по & ".00"
                Else
                    TextBox50.Text = Str(часы + ВрНач) & "0"
                    Dim с As String = Str(5 + ВрНач)
                    Dim по As String = Str(5 + ВрНач + 1)
                    TextBox49.Text = "с" & с & "0 до" & по & "0"
                End If

            Case 11
                If Not (ВрНач = 8.3 Or ВрНач = 10.3) Then
                    TextBox50.Text = Str(часы + ВрНач) & ".00"
                    Dim с As String = Str(5 + ВрНач)
                    Dim по As String = Str(5 + ВрНач + 1)
                    TextBox49.Text = "с" & с & ".00 до" & по & ".00"
                Else
                    TextBox50.Text = Str(часы + ВрНач) & "0"
                    Dim с As String = Str(5 + ВрНач)
                    Dim по As String = Str(5 + ВрНач + 1)
                    TextBox49.Text = "с" & с & "0 до" & по & "0"
                End If
            Case 12
                If Not (ВрНач = 8.3 Or ВрНач = 10.3) Then
                    TextBox50.Text = Str(часы + ВрНач) & ".00"
                    Dim с As String = Str(6 + ВрНач)
                    Dim по As String = Str(6 + ВрНач + 1)
                    TextBox49.Text = "с" & с & ".00 до" & по & ".00"
                Else
                    TextBox50.Text = Str(часы + ВрНач) & "0"
                    Dim с As String = Str(6 + ВрНач)
                    Dim по As String = Str(6 + ВрНач + 1)
                    TextBox49.Text = "с" & с & "0 до" & по & "0"
                End If
            Case 4.3
                If Not (ВрНач = 8.3 Or ВрНач = 10.3) Then
                    TextBox50.Text = Str(часы + ВрНач) & "0"
                    Dim с As String = Str(2 + ВрНач)
                    Dim по As String = Str(2 + ВрНач + 0.3)
                    TextBox49.Text = "с" & с & ".00 до" & по & "0"
                Else
                    Select Case ВрНач
                        Case 8.3
                            TextBox50.Text = "13.00"
                            TextBox49.Text = "с 10.30 по 11.00"
                        Case 10.3
                            TextBox50.Text = "15.00"
                            TextBox49.Text = "с 12.30 по 13.00"
                    End Select

                End If
            Case 2.15
                If Not (ВрНач = 8.3 Or ВрНач = 10.3) Then
                    TextBox50.Text = Str(часы + ВрНач)
                    Dim с As String = Str(1 + ВрНач)
                    Dim по As String = Str(1 + ВрНач + 0.15)
                    TextBox49.Text = "с" & с & ".00 до " & по
                Else
                    Select Case ВрНач
                        Case 8.3
                            TextBox50.Text = "10.45"
                            TextBox49.Text = "с 9.30 по 9.45"
                        Case 10.3
                            TextBox50.Text = "12.45"
                            TextBox49.Text = "с 11.30 по 11.45"
                    End Select
                End If

        End Select
    End Sub
    Private Sub refrdoljn()
        GroupBox19.Visible = False


        'If CheckBox25.Checked = True Then
        '    ComboBox22.Text = ""
        'End If




        'CheckBox8.Checked = False
        'CheckBox25.Checked = False


        Dim ds8 As DataTable = Selects(StrSql:="SELECT Должность FROM ДогПодДолжн
WHERE Клиент='" & ComboBox1.Text & "'")

        Me.ComboBox22.AutoCompleteCustomSource.Clear()
        Me.ComboBox22.Items.Clear()
        For Each r As DataRow In ds8.Rows
            Me.ComboBox22.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox22.Items.Add(r(0).ToString)
        Next

        Dim ds As DataTable = Selects(StrSql:="SELECT ДогПодрОбязан.Обязанности
FROM ДогПодДолжн INNER JOIN ДогПодрОбязан ON ДогПодДолжн.Код = ДогПодрОбязан.ID
WHERE ДогПодДолжн.Клиент='" & ComboBox1.Text & "' AND ДогПодДолжн.Должность= '" & ComboBox22.Text & "'")

        Me.ListBox1.Items.Clear()
        For Each r As DataRow In ds.Rows
            Me.ListBox1.Items.Add(r(0).ToString)
        Next


    End Sub
    Private Function list1Update(ByVal _Клиент As String, ByVal _Должность As String) As List(Of String)
        Dim ds As List(Of String)
        Using db As New DbAll1DataContext
            ds = (From x In db.ДогПодДолжн.AsEnumerable
                  Join y In db.ДогПодрОбязан.AsEnumerable On x.Код Equals y.ID
                  Where x.Клиент = _Клиент And x.Должность = _Должность
                  Select y.Обязанности).ToList()
        End Using
        Return ds


    End Function
    Private Sub listbx1(ByVal t As String, ByVal d As String)
        Dim ds
        Dim db As New DbAll1DataContext
        ds = (From x In db.ДогПодДолжн.AsEnumerable
              Join y In db.ДогПодрОбязан.AsEnumerable On x.Код Equals y.ID
              Where x.Клиент = t And x.Должность = d
              Order By y.Обязанности
              Select y.Обязанности).ToList()

        '        Dim ds = Selects(StrSql:="SELECT ДогПодрОбязан.Обязанности
        'FROM ДогПодДолжн INNER JOIN ДогПодрОбязан ON ДогПодДолжн.Код = ДогПодрОбязан.ID
        'WHERE ДогПодДолжн.Клиент='" & ComboBox1.Text & "' AND ДогПодДолжн.Должность= '" & ComboBox22.Text & "'")

        ListBox1.DataSource = ds
    End Sub
    Private Async Sub listbx1Saync(ByVal t As String, ByVal d As String)
        Await Task.Run(Sub() listbx1(t, d))
    End Sub

    Private Sub ComboBox22_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox22.SelectedIndexChanged

        If GroupBox19.Visible = True Then
            listbx1(ComboBox1.Text, ComboBox22.Text)
            checkbx24()
        Else
            'listbx1Saync(ComboBox1.Text, ComboBox22.Text)
            listbx1(ComboBox1.Text, ComboBox22.Text)
        End If

        'Dim db As New DbAll1DataContext
        'Dim ds = (From x In db.ДогПодДолжн.AsEnumerable
        '          Join y In db.ДогПодрОбязан.AsEnumerable On x.Код Equals y.ID
        '          Where x.Клиент = ComboBox1.Text And x.Должность = ComboBox22.Text
        '          Order By y.Обязанности
        '          Select y.Обязанности).ToList()

        ''        Dim ds = Selects(StrSql:="SELECT ДогПодрОбязан.Обязанности
        ''FROM ДогПодДолжн INNER JOIN ДогПодрОбязан ON ДогПодДолжн.Код = ДогПодрОбязан.ID
        ''WHERE ДогПодДолжн.Клиент='" & ComboBox1.Text & "' AND ДогПодДолжн.Должность= '" & ComboBox22.Text & "'")

        'ListBox1.DataSource = ds


        'ListBox1.Items.Clear()
        'For Each r In ds
        '    ListBox1.Items.Add(r.ToString)
        'Next

        RichTextBox1.Text = ""
        RichTextBox2.Text = ComboBox22.Text
    End Sub

    Private Sub checkbx24()
        Using dbcx = New DbAll1DataContext()
            Dim dt = From x In dbcx.ДогПодДолжн.AsEnumerable
                     Join y In dbcx.ДогПодрОбязан.AsEnumerable On x.Код Equals y.ID
                     Where x.Код = ComboBox22.SelectedValue And x.Клиент = ComboBox1.Text
                     Select New With {.Идентификатор = x.Код, .Организация = x.Клиент, x.Должность, y.Обязанности}

            Grid1.DataSource = dt.ToList()
            'Grid1.Columns("Код").DividerWidth = 80

        End Using
        GridView(Grid1)

    End Sub

    Private Sub CheckBox24_CheckedChanged(sender As Object, e As EventArgs)

        'Старые данные до 27.12.19                                     

        Exit Sub
        'Соед(0)

        Dim ds As DataTable = Selects(StrSql:="SELECT ДогПодрОбязан.Обязанности, ДогПодрОбязан.ID, ДогПодрОбязан.Код
FROM ДогПодДолжн INNER JOIN ДогПодрОбязан ON ДогПодДолжн.Код = ДогПодрОбязан.ID
WHERE ДогПодДолжн.Клиент='" & ComboBox1.Text & "' AND ДогПодДолжн.Должность= '" & ComboBox22.Text & "'")



        'Соед(0)




        'hscol = HS1.LongCount
        'For ia As Integer = 0 To HS.Count - 1
        '    Console.WriteLine(a(ia))
        'Next
        'hscol2 = hscol + 0
        Dim ms() As String
        'Dim ms() As String = {TextBox64.Text, TextBox65.Text, TextBox66.Text, TextBox67.Text, TextBox73.Text, TextBox72.Text, TextBox71.Text, TextBox70.Text,
        '    TextBox69.Text, TextBox68.Text, TextBox75.Text, TextBox74.Text, TextBox77.Text}

        For i As Integer = 0 To ms.Length - 1
            If ms(i) = "" Then
                hscol = hscol + 1
            End If
        Next
        hscol22 = ms.Length - hscol
        Try
            IDLДогПодрОбяз = ds.Rows(0).Item(1)
            MassДогПодрОбяз = New Integer(hscol22 - 1) {}
        Catch ex As Exception

        End Try


        For df As Integer = 0 To hscol22 - 1
            MassДогПодрОбяз(df) = ds.Rows(df).Item(2)
        Next


        'СписОбязан

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)

    End Sub



    Private Sub TextBox34_TextChanged(sender As Object, e As EventArgs) Handles TextBox34.TextChanged
        sender.text = StrConv(sender.text, VbStrConv.ProperCase)
        sender.SelectionStart = sender.text.Length
    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged
        sender.text = StrConv(sender.text, VbStrConv.ProperCase)
        sender.SelectionStart = sender.text.Length
    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged
        sender.text = StrConv(sender.text, VbStrConv.ProperCase)
        sender.SelectionStart = sender.text.Length
    End Sub

    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            TextBox1.Focus()
        End If
    End Sub

    'Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
    '    but2cl = 1
    '    Button1.PerformClick()
    'End Sub



    Private Sub ComboBox10_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox10.SelectedIndexChanged
        If ComboBox10.Text = "" Then
            Label78.ForeColor = Color.Red
            Label78.Text = "NO"
        Else
            Label78.ForeColor = Color.Green
            Label78.Text = "OK"
        End If
    End Sub

    Private Sub TextBox56_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox56.KeyDown
        If ComboBox15.Text = "Задать" Or ComboBox15.Text = "" Then

            If e.KeyCode = Keys.Enter Then
                e.SuppressKeyPress = True
                Me.ComboBox12.Focus()
            Else
                Me.Button1.Focus()
            End If


        End If
    End Sub

    Private Sub ComboBox18_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox18.SelectedIndexChanged
        If ComboBox18.Text = "" Then
            Label85.ForeColor = Color.Red
            Label85.Text = "NO"
        Else
            Label85.ForeColor = Color.Green
            Label85.Text = "OK"
        End If
    End Sub

    Private Sub ComboBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox7.KeyDown
        If e.KeyCode = Keys.Enter And Label79.Text = "OK" Then
            e.SuppressKeyPress = True
            Me.ComboBox10.Focus()
        End If
    End Sub

    Private Sub ComboBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox10.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox2.Focus()
        End If
    End Sub

    Private Sub TextBox33_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox33.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox44.Focus()
        End If
    End Sub
    Private Sub TextBox44_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox44.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox46.Focus()
        End If
    End Sub
    Private Sub TextBox46_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox46.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox2.Focus()
        End If
    End Sub
    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox17.Focus()
        End If
    End Sub
    Private Sub ComboBox17_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox17.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox18.Focus()
        End If
    End Sub
    Private Sub MaskedTextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            MaskedTextBox4.Focus()

        End If
    End Sub
    Private Sub ComboBox11_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox11.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox38.Focus()
        End If
    End Sub
    Private Sub TextBox38_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox38.KeyDown


        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox41.Focus()

            Dim pl As String
            If TextBox38.Text <> "" Then
                Try
                    Dim i As Integer = CInt(TextBox38.Text)
                    Select Case i
                        Case < 10
                            pl = Str(i)
                            TextBox38.Text = "00" & i

                        Case 10 To 99
                            pl = Str(i)
                            TextBox38.Text = "0" & i
                    End Select
                Catch ex As Exception
                    TextBox38.Text = Replace(TextBox38.Text, "/", ".")
                    TextBox38.Text = Replace(TextBox38.Text, "\", ".")
                    TextBox38.Text = "б.н"
                End Try
            Else
                TextBox38.Text = "б.н"
            End If

        End If


    End Sub
    Private Sub TextBox41_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox41.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            ComboBox15.Focus()

            Dim pl As String
            If TextBox41.Text <> "" Then
                Dim i As Integer = CInt(TextBox41.Text)
                Select Case i

                    Case < 10
                        pl = Str(i)
                        TextBox41.Text = "00" & i

                    Case 10 To 99
                        pl = Str(i)
                        TextBox41.Text = "0" & i
                End Select
            End If
        End If
    End Sub
    Private Sub MaskedTextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox4.KeyDown

        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox11.Focus()
            РасчСрокаКонтр()
        End If

    End Sub
    Private Sub MaskedTextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox5.KeyDown

        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox15.Focus()
        End If
    End Sub
    Private Sub ComboBox15_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox15.KeyDown

        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox40.Focus()


        End If

    End Sub

    Private Sub ComboBox12_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox12.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox16.Focus()
        End If

    End Sub
    Private Sub ComboBox16_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox16.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Button1.Focus()
        End If

    End Sub
    Public Function ПроверкаЗаполненности(ByVal sw As String) As String
        Dim s As String = "ООО РикКонсалтинг"
        If ComboBox1.Text = "" Then
            MsgBox("Выберите организацию!",, s)
            Return Nothing

        End If

        If CheckBox1.Checked = False Then
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Then
                MsgBox("Заполните реквизиты сотрудника",, s)
                Return Nothing
            End If
        End If


        If TextBox20.Text = "" Or TextBox21.Text = "" Then
            MsgBox("Заполните адрес сотрудника!",, s)
            Return Nothing
        End If


        If CheckBox1.Checked = False Then
            If TextBox7.Text = "" Or TextBox8.Text = "" Or TextBox9.Text = "" Or TextBox12.Text = "" Or TextBox45.Text = "" Then
                MsgBox("Заполните паспортные данные сотрудника!",, s)
                Return Nothing
            End If
        End If


        If CheckBox7.Checked = False And CheckBox5.Checked = False Then
            If ComboBox8.Text = "" Or ComboBox9.Text = "" Or ComboBox10.Text = "" Or TextBox33.Text = "" Then
                MsgBox("Заполните раздел подразделение!",, s)
                Return Nothing
            End If

            If MaskedTextBox3.Text = "" Or TextBox38.Text = "" Or TextBox41.Text = "" Or MaskedTextBox4.Text = "" Or MaskedTextBox5.Text = "" Or ComboBox15.Text = "" Or ComboBox11.Text = "" Then
                MsgBox("Заполните раздел контракт и приказ!",, s)
                Return Nothing
            End If
        End If
        ПроверкаЗаполненности = 1
        Return ПроверкаЗаполненности
    End Function

    Private Sub ComboBox12_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox12.SelectedValueChanged
        Dim часы As Decimal = Val(ComboBox16.Text) 'расчет времени обеда и конца рабочего дня
        Dim ВрНач As Decimal = Val(ComboBox12.Text)
        'часы = Math.Floor(часы)
        Select Case часы
            Case 9
                TextBox50.Text = Str(часы + ВрНач) & ".00"
                Dim с As String = Str(4 + ВрНач)
                Dim по As String = Str(4 + ВрНач + 1)
                TextBox49.Text = "с " & с & ".00  до " & по & ".00"
            Case 10
                TextBox50.Text = Str(часы + ВрНач) & ".00"
                Dim с As String = Str(5 + ВрНач)
                Dim по As String = Str(5 + ВрНач + 1)
                TextBox49.Text = "с " & с & ".00  до " & по & ".00"
            Case 11
                TextBox50.Text = Str(часы + ВрНач) & ".00"
                Dim с As String = Str(5 + ВрНач)
                Dim по As String = Str(5 + ВрНач + 1)
                TextBox49.Text = "с " & с & ".00  до " & по & ".00"
            Case 12
                TextBox50.Text = Str(часы + ВрНач) & ".00"
                Dim с As String = Str(6 + ВрНач)
                Dim по As String = Str(6 + ВрНач + 1)
                TextBox49.Text = "с " & с & ".00  до " & по & ".00"
        End Select
    End Sub
    Private Sub РасчСрокаКонтр()
        Try
            dad = CDate(MaskedTextBox4.Text)
        Catch ex As Exception
            MessageBox.Show("Введите правильно формат даты!", Рик)
            Exit Sub
        End Try

        Select Case ComboBox11.Text
            Case "1"
                MaskedTextBox5.Text = dad.AddMonths(12)
                Dim dad2 As Date = CDate(MaskedTextBox5.Text)
                MaskedTextBox5.Text = dad2.AddDays(-1)
            Case "2"
                MaskedTextBox5.Text = dad.AddMonths(24)
                Dim dad2 As Date = CDate(MaskedTextBox5.Text)
                MaskedTextBox5.Text = dad2.AddDays(-1)
            Case "3"
                MaskedTextBox5.Text = dad.AddMonths(36)
                Dim dad2 As Date = CDate(MaskedTextBox5.Text)
                MaskedTextBox5.Text = dad2.AddDays(-1)
            Case "4"
                MaskedTextBox5.Text = dad.AddMonths(48)
                Dim dad2 As Date = CDate(MaskedTextBox5.Text)
                MaskedTextBox5.Text = dad2.AddDays(-1)
            Case "5"
                MaskedTextBox5.Text = dad.AddMonths(60)
                Dim dad2 As Date = CDate(MaskedTextBox5.Text)
                MaskedTextBox5.Text = dad2.AddDays(-1)

        End Select
    End Sub

    Private Async Sub ComboBox10_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox10.SelectedValueChanged
        ПропОклад()
    End Sub

    Private Sub ComboBox14_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox14.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            TabControl1.SelectedTab = TabControl1.TabPages("TabPage2")
            Me.ComboBox8.Focus()
        End If
    End Sub

    Private Sub MaskedTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Dim dad As Date
            Try
                dad = CDate(MaskedTextBox1.Text)
                MaskedTextBox2.Text = dad.AddYears(10)
            Catch ex As Exception
                MessageBox.Show("Это поле даты! Проверьте введенные значения", Рик)
                MaskedTextBox1.Focus()
                MaskedTextBox1.Text = ""
                Exit Sub
            End Try

            Me.MaskedTextBox2.Focus()
        End If
    End Sub

    Private Sub MaskedTextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox28.Focus()
        End If
    End Sub

    Private Sub ComboBox18_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox18.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.MaskedTextBox3.Focus()
        End If
    End Sub

    Private Sub ComboBox23_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox23.SelectedValueChanged
        ComboBox24.Items.Clear()
        ComboBox24.Text = ""
        'Dim Files3()

        'Files3 = (IO.Directory.GetFiles(OnePath & ComboBox1.Text & "\Договор подряда\" & ComboBox23.Text, "*.doc", IO.SearchOption.TopDirectoryOnly))

        'Dim gth As String
        'For n As Integer = 0 To Files3.Length - 1
        '    gth = ""
        '    gth = IO.Path.GetFileName(Files3(n))
        '    Files3(n) = gth
        '    'TextBox44.Text &= gth + vbCrLf
        'Next
        Dim listCombo3 As Object = listFluentFTP(ComboBox1.Text & "/Договор подряда/" & ComboBox23.Text & "/")

        For Each item In listCombo3
            ComboBox24.Items.Add(Replace(item, FTPString & ComboBox1.Text & "/Договор подряда/" & ComboBox23.Text & "/", ""))
        Next

    End Sub

    Private Sub textbox55_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox55.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.MaskedTextBox6.Focus()
            Try
                Dim pl As String
                If TextBox55.Text <> "" And IsNumeric(TextBox55.Text) Then
                    Dim i As Integer = CInt(TextBox55.Text)
                    Select Case i

                        Case < 10
                            pl = Str(i)
                            TextBox55.Text = "00" & i
                            but2cl = 0
                        Case 10 To 99
                            pl = Str(i)
                            TextBox55.Text = "0" & i
                            but2cl = 0
                        Case > 100
                            but2cl = 0
                    End Select
                End If
            Catch ex As Exception
                MessageBox.Show("Номером договора подряда может быть только целочисленное значение!", Рик, MessageBoxButtons.OK)
                but2cl = 1
            End Try

        End If
    End Sub

    Private Sub ComboBox23_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox23.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox24.Focus()
        End If
    End Sub

    Private Sub ComboBox24_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox24.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox55.Focus()
        End If
    End Sub

    Private Sub MaskedTextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox22.Focus()
        End If
    End Sub

    Private Sub ComboBox22_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox22.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.MaskedTextBox7.Focus()
        End If
    End Sub

    Private Sub MaskedTextBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox7.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.MaskedTextBox8.Focus()
        End If
    End Sub

    Private Sub MaskedTextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox8.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox61.Focus()
        End If
    End Sub

    Private Sub TextBox61_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox61.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox62.Focus()
        End If
    End Sub

    Private Sub MaskedTextBox1_LostFocus(sender As Object, e As EventArgs) Handles MaskedTextBox1.LostFocus
        'Try
        '    Dim dad As Date = CDate(MaskedTextBox1.Text)
        '    MaskedTextBox2.Text = dad.AddYears(10)
        'Catch ex As Exception

        'End Try

    End Sub


    Private Sub TextBox41_LostFocus(sender As Object, e As EventArgs) Handles TextBox41.LostFocus
        Dim pl As String
        If TextBox41.Text <> "" Then

            Dim i As Integer = CInt(TextBox41.Text)
            Select Case i

                Case < 10
                    pl = Str(i)
                    TextBox41.Text = "00" & i

                Case 10 To 99
                    pl = Str(i)
                    TextBox41.Text = "0" & i
            End Select
        End If
    End Sub

    Private Sub TextBox38_LostFocus(sender As Object, e As EventArgs) Handles TextBox38.LostFocus
        Dim pl As String
        If TextBox38.Text <> "" Then
            Try
                Dim i As Integer = CInt(TextBox38.Text)
                Select Case i

                    Case < 10
                        pl = Str(i)
                        TextBox38.Text = "00" & i

                    Case 10 To 99
                        pl = Str(i)
                        TextBox38.Text = "0" & i
                End Select
            Catch ex As Exception
                TextBox38.Text = "б.н"
            End Try
        Else
            TextBox38.Text = "б.н"

        End If
    End Sub

    Private Sub TextBox40_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox40.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox56.Focus()


        End If
    End Sub

    Private Sub TextBox7_LostFocus(sender As Object, e As EventArgs) Handles TextBox7.LostFocus
        'Dim f As Integer
        'If CheckBox1.Checked = False Then
        '    Try
        '        f = CType(TextBox7.Text, Integer)
        '    Catch ex As Exception
        '        MessageBox.Show("Это поле должно содеражть только цифры!")
        '        TextBox7.Focus()
        '        TextBox7.Text = ""
        '    End Try
        'End If

    End Sub

    Private Sub MaskedTextBox3_LostFocus(sender As Object, e As EventArgs) Handles MaskedTextBox3.LostFocus
        'MaskedTextBox1.Text = Format(Now, "dd.MM.yyyy")
        Dim s As Date = CDate("01.01.1900")
        Dim s2 As Date = CDate("01.01.2050")
        Dim s3 As Date
        Try
            s3 = CDate(MaskedTextBox3.Text)
        Catch ex As Exception
            MessageBox.Show("Заполните поле дата!", Рик)
            MaskedTextBox3.Focus()
            Exit Sub
        End Try


        If s3 > s2 Or s3 < s Then
            MessageBox.Show("Проверьте дату!", Рик)
            MaskedTextBox3.Focus()
            MaskedTextBox3.Text = ""
            Exit Sub
        End If
        If ComboBox8.Text <> "" And ComboBox9.Text <> "" Then
            Данные()
        End If

    End Sub

    Private Sub TextBox55_LostFocus(sender As Object, e As EventArgs) Handles TextBox55.LostFocus
        Dim pl As String
        If TextBox55.Text <> "" And IsNumeric(TextBox55.Text) Then
            Dim i As Integer = CInt(TextBox55.Text)
            Select Case i

                Case < 10
                    pl = Str(i)
                    TextBox55.Text = "00" & i

                Case 10 To 99
                    pl = Str(i)
                    TextBox55.Text = "0" & i
            End Select
        Else
            MessageBox.Show("Номер договор-подряда должен быть целочисленным!", Рик)
            TextBox55.ForeColor = Color.Red
        End If
    End Sub



    Private Sub MaskedTextBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox10.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True

            If GroupBox26.Visible = True Then
                TextBox51.Focus()
            Else
                TextBox12.Focus()
            End If

        End If
    End Sub

    Private Sub ComboBox1_MouseWheel(sender As Object, e As MouseEventArgs) Handles ComboBox1.MouseWheel
        Dim mwe As HandledMouseEventArgs = DirectCast(e, HandledMouseEventArgs)
        mwe.Handled = True
    End Sub

    Private Sub ComboBox19_MouseWheel(sender As Object, e As MouseEventArgs) Handles ComboBox19.MouseWheel
        Dim mwe As HandledMouseEventArgs = DirectCast(e, HandledMouseEventArgs)
        mwe.Handled = True

    End Sub

    Private Sub ComboBox8_MouseWheel(sender As Object, e As MouseEventArgs) Handles ComboBox8.MouseWheel
        Dim mwe As HandledMouseEventArgs = DirectCast(e, HandledMouseEventArgs)
        mwe.Handled = True
    End Sub
    Private Sub ComboBox9_MouseWheel(sender As Object, e As MouseEventArgs) Handles ComboBox9.MouseWheel
        Dim mwe As HandledMouseEventArgs = DirectCast(e, HandledMouseEventArgs)
        mwe.Handled = True
    End Sub

    Private Sub ComboBox11_MouseWheel(sender As Object, e As MouseEventArgs) Handles ComboBox11.MouseWheel
        Dim mwe As HandledMouseEventArgs = DirectCast(e, HandledMouseEventArgs)
        mwe.Handled = True
    End Sub

    Private Sub ComboBox15_MouseWheel(sender As Object, e As MouseEventArgs) Handles ComboBox15.MouseWheel
        Dim mwe As HandledMouseEventArgs = DirectCast(e, HandledMouseEventArgs)
        mwe.Handled = True
    End Sub

    Private Sub ComboBox12_MouseWheel(sender As Object, e As MouseEventArgs) Handles ComboBox12.MouseWheel
        Dim mwe As HandledMouseEventArgs = DirectCast(e, HandledMouseEventArgs)
        mwe.Handled = True
    End Sub

    Private Sub ComboBox16_MouseWheel(sender As Object, e As MouseEventArgs) Handles ComboBox16.MouseWheel
        Dim mwe As HandledMouseEventArgs = DirectCast(e, HandledMouseEventArgs)
        mwe.Handled = True
    End Sub
    Private Sub ComboBox7_MouseWheel(sender As Object, e As MouseEventArgs) Handles ComboBox7.MouseWheel
        Dim mwe As HandledMouseEventArgs = DirectCast(e, HandledMouseEventArgs)
        mwe.Handled = True
    End Sub

    Private Sub ComboBox10_MouseWheel(sender As Object, e As MouseEventArgs) Handles ComboBox10.MouseWheel
        Dim mwe As HandledMouseEventArgs = DirectCast(e, HandledMouseEventArgs)
        mwe.Handled = True
    End Sub
    Private Sub ComboBox18_MouseWheel(sender As Object, e As MouseEventArgs) Handles ComboBox18.MouseWheel
        Dim mwe As HandledMouseEventArgs = DirectCast(e, HandledMouseEventArgs)
        mwe.Handled = True
    End Sub
    Private Sub ComboBox22_MouseWheel(sender As Object, e As MouseEventArgs) Handles ComboBox22.MouseWheel
        Dim mwe As HandledMouseEventArgs = DirectCast(e, HandledMouseEventArgs)
        mwe.Handled = True
    End Sub
    Private Sub ComboBox25_MouseWheel(sender As Object, e As MouseEventArgs) Handles ComboBox25.MouseWheel
        Dim mwe As HandledMouseEventArgs = DirectCast(e, HandledMouseEventArgs)
        mwe.Handled = True
    End Sub
    Private Sub ComboBox2_MouseWheel(sender As Object, e As MouseEventArgs) Handles ComboBox2.MouseWheel
        Dim mwe As HandledMouseEventArgs = DirectCast(e, HandledMouseEventArgs)
        mwe.Handled = True
    End Sub
    Private Sub ComboBox20_MouseWheel(sender As Object, e As MouseEventArgs) Handles ComboBox20.MouseWheel
        Dim mwe As HandledMouseEventArgs = DirectCast(e, HandledMouseEventArgs)
        mwe.Handled = True
    End Sub
    Private Sub ComboBox21_MouseWheel(sender As Object, e As MouseEventArgs) Handles ComboBox21.MouseWheel
        Dim mwe As HandledMouseEventArgs = DirectCast(e, HandledMouseEventArgs)
        mwe.Handled = True
    End Sub
    Private Sub ComboBox17_MouseWheel(sender As Object, e As MouseEventArgs) Handles ComboBox17.MouseWheel
        Dim mwe As HandledMouseEventArgs = DirectCast(e, HandledMouseEventArgs)
        mwe.Handled = True
    End Sub
    Private Sub MaskedTextBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox9.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            TextBox9.Focus()
        End If
    End Sub

    Private Sub MaskedTextBox3_TextChanged(sender As Object, e As EventArgs) Handles MaskedTextBox3.TextChanged



    End Sub

    Private Sub TabControl1_DrawItem(sender As Object, e As DrawItemEventArgs) Handles TabControl1.DrawItem

        Dim g As Graphics = e.Graphics
        Dim tp As TabPage
        Try
            tp = TabControl1.TabPages(e.Index)
        Catch ex As Exception
            Exit Sub
        End Try

        Dim br As Brush
        Dim sf As New StringFormat

        Dim r As New RectangleF(e.Bounds.X, e.Bounds.Y + 2, e.Bounds.Width, e.Bounds.Height - 2)

        sf.Alignment = StringAlignment.Center

        Dim strTitle As String = tp.Text

        'If the current index is the Selected Index, change the color 
        If TabControl1.SelectedIndex = e.Index Then

            'this is the background color of the tabpage header
            br = New SolidBrush(Color.Gainsboro) ' chnge to your choice
            g.FillRectangle(br, e.Bounds)

            'this is the foreground color of the text in the tab header
            br = New SolidBrush(Color.Black) ' change to your choice
            g.DrawString(strTitle, TabControl1.Font, br, r, sf)

        Else

            'these are the colors for the unselected tab pages 
            br = New SolidBrush(Color.LightSteelBlue) ' Change this to your preference
            g.FillRectangle(br, e.Bounds)
            br = New SolidBrush(Color.Black)
            g.DrawString(strTitle, TabControl1.Font, br, r, sf)

        End If

    End Sub
End Class