Option Explicit On
Imports System.Data.Linq
Imports System.Data.Linq.Mapping
Imports System.Data.OleDb
Imports System.Threading

Public Class КонтрДопИзменСтавка
    Public ds As DataTable
    Dim StrSql, ФИОРукРодПад, ФИОПолнРук, ФИОрКОР, ФормСобсКоротко, ДолжДирСОконч,
              inp, ФормСобПолн, ДолжРук, ОснДейств, ДатОкон As String

    Dim ДолжОконСотр1 As New Dictionary(Of Integer, String)()
    Dim ФИОСотрКор1 As New Dictionary(Of Integer, String)()
    Dim Разряд1 As New Dictionary(Of Integer, String)()
    Dim ДатаКонтр1 As New Dictionary(Of Integer, String)()
    Dim НомерКонтр1 As New Dictionary(Of Integer, String)()
    Dim Пол1 As New Dictionary(Of Integer, String)()
    Dim НомерУведИзмОклада1 As New Dictionary(Of Integer, Integer)()
    Dim ДатаУведИзмОклада1 As New Dictionary(Of Integer, String)()
    Dim СотрИКод As New Dictionary(Of Integer, String)()




    Dim oWord() As Microsoft.Office.Interop.Word.Application
    Dim oWordDoc() As Microsoft.Office.Interop.Word.Document
    Dim sotkol As Integer
    Dim thprov As Thread
    Dim combx1, mskbx1, rich1, rich2 As String

    Dim СохрЗак1() As String, СохрЗак2() As String, СохрЗак3() As String, СохрЗак4() As String, СохрЗак5() As String, СохрЗак6() As String



    Private Sub RichTextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.RichTextBox1.Focus()
        End If
    End Sub

    Dim cl, IDСотр, errs As Integer
    Private Sub КонтрДопИзменСтавка_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.ComboBox1.AutoCompleteCustomSource.Clear()
        Me.ComboBox1.Items.Clear()
        For Each r As DataRow In СписокКлиентовОсновной.Rows
            Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox1.Items.Add(r(0).ToString)
        Next
    End Sub

    '    Private Sub Button3_Click(sender As Object, e As EventArgs)
    '        For i = 0 To ListBox2.SelectedItems.Count - 1
    '            Чист()
    '            StrSql = "SELECT Сотрудники.КодСотрудники FROM Сотрудники 
    'WHERE Сотрудники.НазвОрганиз='" & ComboBox1.Text & "' AND Сотрудники.ФИОСборное='" & ListBox2.SelectedItems(i) & "'"
    '            ds = Selects(StrSql)
    '            IDСотр = Nothing
    '            IDСотр = ds.Rows(0).Item(0)


    '            Чист()
    '            StrSql = "UPDATE КарточкаСотрудника SET НомУведИзмСрокЗарп=" & 0 & ", ДатаСогласияНаИзмен='', ДатаВсуплСоглаш='', ДатаУведом=''
    'WHERE КарточкаСотрудника.IDСотр=" & IDСотр & ""
    '            Updates(StrSql)
    '        Next
    '        MessageBox.Show("Данные согласно Номера уведомления, даты согласования очищены удачно!", Рик)
    '    End Sub

    Private Sub MaskedTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.RichTextBox2.Focus()
        End If
    End Sub



    Private Sub Очист()
        RichTextBox1.Text = ""

        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        MaskedTextBox1.Text = ""

    End Sub
    Private Sub refreshes()
        СотрИКод.Clear()
        Чист()
        Dim list As New Dictionary(Of String, Object)()        '
        list.Add("@combx1", combx1)

        Dim ds = Selects(StrSql:="SELECT DISTINCT ФИОСборное, КодСотрудники
FROM Сотрудники
WHERE НазвОрганиз=@combx1 AND НаличеДогПодряда='Нет' ORDER BY ФИОСборное", list)


        For Each r As DataRow In ds.Rows
            ListBox1.Items.Add(r.Item(0).ToString())
            ListBox3.Items.Add(r.Item(1).ToString())
            'dict.Add(r.Item(0).ToString, r.Item(1))
            СотрИКод.Add(r.Item(1), r.Item(0).ToString)
        Next


        'sotkol = ds.Rows.Count
    End Sub
    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        Очист()
        Label7.Text = ComboBox1.Text
        combx1 = ComboBox1.Text

        ДанСотрОбщ()
        refreshes()
        'thprov = New Thread(AddressOf proverka) With {
        '    .IsBackground = True}
        'thprov.Start(0)
        ОбновлПапкиУведомление()
        Выборка() 'данные по организации)


    End Sub
    Private Sub ОбновлПапкиУведомление()

        СозданиепапкиНаСервере(combx1 & " / Уведомление / " & Now.Year)

        Dim list = listFluentFTP(combx1 & " / Уведомление / " & Now.Year & "/")


        ListBox2.Items.Clear()
        For Each x In list
            ListBox2.Items.Add(list.ToString)
        Next



        'For i = 0 To file2.Length - 1 ' Распечатываем весь получившийся массив
        '    ListBox2.Items.Add(file2.Reverse(i)) ' На ListBox2
        'Next
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ListBox1.Items.Clear()
    End Sub
    Private Function ПровЗаполн()
        If ComboBox1.Text = "" Then
            MessageBox.Show("Выберите организацию", Рик)
            Return 1
        End If

        If RichTextBox1.Text = "" Then
            MessageBox.Show("Заполните текст поля, в который внсятся изменения!", Рик)
            Return 1
        End If

        If RichTextBox2.Text = "" Then
            MessageBox.Show("Заполните поле "" Тема уведомления ""!", Рик)
            Return 1
        End If

        If ListBox1.SelectedIndex = -1 Then
            MessageBox.Show("Выберите сотрудников!", Рик)
            Return 1
        End If
        If MaskedTextBox1.Text = "" Then
            MessageBox.Show("Выберите дату уведомления!", Рик)
            Return 1
        End If

        Return 0
    End Function
    Private Sub Чист()
        StrSql = ""
        If Not ds Is Nothing Then
            ds.Clear()
        End If

    End Sub
    Private Sub база()
        Dim datsog, datsogl As String
        Dim datuv As Date = CDate(MaskedTextBox1.Text)
        datsog = Strings.Left(datuv.AddDays(15).ToString, 10)
        datuv = datuv.AddMonths(1)
        datuv = datuv.AddDays(2)
        datsogl = Strings.Left(datuv.ToString, 10)
        datuv = datuv.AddDays(-1)
        ДатОкон = Strings.Left(datuv.ToString, 10)


        If ДатаУведИзмОклада1.Count >= 1 Then
            ДатаУведИзмОклада1.Clear()
            НомерУведИзмОклада1.Clear()
        End If

        For i = 0 To ListBox1.SelectedItems.Count - 1

            Dim ds = dtSotrudnikiAll.Select("НазвОрганиз='" & ComboBox1.Text & "' AND ФИОСборное='" & ListBox1.SelectedItems(i) & "'")
            IDСотр = Nothing
            IDСотр = ds(0).Item(0)

            Dim cld As String = ""
            cl = cl + 1

            'Dim db As New DataContext(ConString)

            'Dim КарточкаСотр As Table(Of КарточкаСотрудника)

            'КарточкаСотр = db.GetTable(Of КарточкаСотрудника)()

            'Dim КартСот As КартСотр = New КартСотр()
            'КартСот = db.GetTable(Of КартСотр)().FirstOrDefault
            'КартСот.НомерУведИзмОклада = cl
            'КартСот.ДатаУведомИзмОклада = MaskedTextBox1.Text
            'db.SubmitChanges()

            Dim list As New Dictionary(Of String, Object)
            list.Add("@IDСотр", IDСотр)

            Updates(stroka:="UPDATE КарточкаСотрудника SET НомерУведИзмОклада='" & cl & "', ДатаУведомИзмОклада='" & MaskedTextBox1.Text & "'
            WHERE КарточкаСотрудника.IDСотр=@IDСотр", list)


            ДатаУведИзмОклада1.Add(ListBox3.SelectedItems(i), MaskedTextBox1.Text)
            НомерУведИзмОклада1.Add(ListBox3.SelectedItems(i), cl)
        Next

        dtKartochkaSotrudnika()
    End Sub
    Private Sub Выборка()
        '        Чист() 'выборка данных организации


        '        StrSql = "Select Клиент.ФормаСобств, Клиент.ДолжнРуководителя, Клиент.ФИОРуководителя,
        'Клиент.ФИОРукРодПадеж, Клиент.РукИП, [ОснованиеДейств]
        'From Клиент Where Клиент.НазвОрг = '" & combx1 & "'"

        '        ds = Selects(StrSql)
        Dim ds = dtClientAll.Select("НазвОрг = '" & combx1 & "'")

        Dim РуковИП As String
        If ds(0).Item("РукИП") = "True" Then
            РуковИП = "ИП "
        Else
            РуковИП = ""
        End If

        ФормСобПолн = ""
        ФИОПолнРук = ""
        ФИОРукРодПад = ""
        ФИОрКОР = ""
        ФормСобсКоротко = ""
        ДолжДирСОконч = ""
        ДолжРук = ""
        ОснДейств = ""

        ФормСобПолн = ds(0).Item("ФормаСобств").ToString
        ФИОПолнРук = ds(0).Item("ФИОРуководителя").ToString
        ФИОРукРодПад = РуковИП & ds(0).Item("ФИОРукРодПадеж").ToString
        ФИОрКОР = ФИОКорРук(ФИОПолнРук, ds(0).Item("РукИП"))
        ФормСобсКоротко = ФормСобствКор(ds(0).Item("ФормаСобств").ToString)
        ДолжДирСОконч = ДобОконч(ds(0).Item("ДолжнРуководителя").ToString)
        ДолжРук = ds(0).Item("ДолжнРуководителя").ToString
        ОснДейств = ds(0).Item("ОснованиеДейств").ToString
    End Sub
    Private Async Sub ДанСотрОбщ()
        Await Task.Run(Sub() ДанСотрОбщ1())
    End Sub
    Private Sub ДанСотрОбщ1()
        If Разряд1.Count > 1 Then
            Разряд1.Clear()
            ДатаКонтр1.Clear()
            НомерКонтр1.Clear()
            ДолжОконСотр1.Clear()
            ФИОСотрКор1.Clear()
            Пол1.Clear()
        End If


        Dim list As New Dictionary(Of String, Object)()        '
        list.Add("@combx1", combx1)

        Dim ds = Selects(StrSql:="SELECT Сотрудники.ФИОСборное, Сотрудники.КодСотрудники, Штатное.Должность, Штатное.Разряд,
ДогСотрудн.ДатаКонтракта, ДогСотрудн.Контракт, Сотрудники.НазвОрганиз, Сотрудники.Пол
FROM ((Сотрудники INNER JOIN ДогСотрудн ON Сотрудники.КодСотрудники = ДогСотрудн.IDСотр) INNER JOIN Штатное ON Сотрудники.КодСотрудники = Штатное.ИДСотр)
INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр
WHERE Сотрудники.НазвОрганиз=@combx1", list)

        For i As Integer = 0 To ds.Rows.Count - 1

            If ds.Rows(i).Item(3).ToString <> "" And Not ds.Rows(i).Item(3).ToString = "-" Then
                Try

                    Разряд1.Add(ds.Rows(i).Item(1), разрядстрока(CType(ds.Rows(i).Item(3), Integer)))

                Catch ex1 As Exception
                    Разряд1.Add(ds.Rows(i).Item(1), "")
                End Try
            Else
                Разряд1.Add(ds.Rows(i).Item(1), "")
            End If

            ДатаКонтр1.Add(ds.Rows(i).Item(1), Strings.Left(ds.Rows(i).Item(4), 10))
            НомерКонтр1.Add(ds.Rows(i).Item(1), ds.Rows(i).Item(5).ToString)
            ДолжОконСотр1.Add(ds.Rows(i).Item(1), ds.Rows(i).Item(2).ToString)
            ФИОСотрКор1.Add(ds.Rows(i).Item(1), ФИОКорРук(ds.Rows(i).Item(0).ToString, False))
            Пол1.Add(ds.Rows(i).Item(1), ds.Rows(i).Item(7).ToString)

        Next

        'Dim ДолжОконСотр1 As New Dictionary(Of String, String)()
        'Dim ФИОСотрКор1 As New Dictionary(Of String, String)()
        'Dim Разряд1 As New Dictionary(Of String, String)()
        'Dim ДатаКонтр1 As New Dictionary(Of String, String)()
        'Dim НомерКонтр1 As New Dictionary(Of String, String)()


    End Sub



    '    Private Sub ДанСотр(ByVal IDсот As Integer)

    '            Чист() ' выбираем данные по сотруднику
    '            StrSql = "SELECT Штатное.Должность, Штатное.Разряд, ДогСотрудн.ДатаКонтракта, ДогСотрудн.Контракт, Сотрудники.ФИОСборное
    'FROM (Сотрудники INNER JOIN ДогСотрудн ON Сотрудники.КодСотрудники = ДогСотрудн.IDСотр) INNER JOIN Штатное ON Сотрудники.КодСотрудники = Штатное.ИДСотр
    'Where Штатное.ИДСотр = " & IDсот & " And ДогСотрудн.IDСотр = " & IDсот & ""
    '            ds = Selects(StrSql)
    '            Try
    '                'ДолжОконСотр = ДолжРодПадежФункц(ds.Rows(0).Item(0).ToString)
    '                ДолжОконСотр = ds.Rows(0).Item(0).ToString
    '            Catch ex As Exception
    '                MessageBox.Show("У данного сотрудника нет должности!", Рик)
    '                'errs = 1
    '                Exit Sub

    '            End Try


    '            ФИОСотрКор = ФИОКорРук(ds.Rows(0).Item(4).ToString, False)

    '            If ds.Rows(0).Item(1) <> "" Then
    '                Try
    '                    Разряд = разрядстрока(CType(ds.Rows(0).Item(1), Integer))
    '                Catch ex As Exception
    '                    Разряд = ""
    '                End Try
    '            Else
    '                Разряд = ""
    '            End If

    '            ДатаКонтр = Strings.Left(ds.Rows(0).Item(2), 10)
    '            НомерКонтр = ds.Rows(0).Item(3).ToString


    '            'Dim StrSql2 As String = "SELECT ФИОДатПадКому FROM Сотрудники Where Сотрудники.КодСотрудники = " & IDсот & ""
    '            'Dim ds2 As DataTable = Selects(StrSql2)

    '            'If Not ds2.Rows(0).Item(0).ToString <> "" Then
    '            '    inp = InputBox("Введите ФИО сотрудника " & ds.Rows(0).Item(4).ToString & " в Дательном падеже (Уведомление)'Кому?'", Рик)

    '            '    Do Until inp <> ""
    '            '        MessageBox.Show("Повторите ввод ФИО!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            '        inp = InputBox("Введите ФИО сотрудника " & ds.Rows(0).Item(4).ToString & " в Дательном падеже (Уведомление)'Кому?'", Рик)

    '            '    Loop
    '            '    Чист()
    '            '    StrSql = "UPDATE Сотрудники SET ФИОДатПадКому='" & inp & "' Where Сотрудники.КодСотрудники = " & IDсот & ""
    '            '    Updates(StrSql)
    '            'Else
    '            '    inp = ds2.Rows(0).Item(0).ToString
    '            'End If


    '        End Sub
    Private Sub доки1(ByVal Поток1 As List(Of Integer))
        'Dim СохрДП(ListBox1.SelectedItems.Count - 1) As String
        ReDim СохрЗак1(Поток1.Count - 1)

        Try
            If IO.Directory.Exists("c:\Users\Public\Documents\Рик") Then
                'IO.Directory.Delete("c:\Users\Public\Documents\Рик", True)
                'IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рик")
            Else
                IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рик")
            End If
        Catch ex As Exception

        End Try

        If Not IO.Directory.Exists(OnePath & combx1 & "\Уведомление\" & Year(Now)) Then
            IO.Directory.CreateDirectory(OnePath & combx1 & "\Уведомление\" & Year(Now))
        End If


        'thprov.Join()

        Try
            IO.File.Copy(OnePath & "\ОБЩДОКИ\General\UvedObIzmenOklada.doc", "C:\Users\Public\Documents\Рик\УведОбИзменОклада1.doc")
        Catch ex As Exception
            IO.File.Delete("C:\Users\Public\Documents\Рик\УведОбИзменОклада1.doc")
            IO.File.Copy(OnePath & "\ОБЩДОКИ\General\UvedObIzmenOklada.doc", "C:\Users\Public\Documents\Рик\УведОбИзменОклада1.doc")
        End Try

        СозданиепапкиНаСервере(combx1 & "\Уведомление\" & Year(Now))

        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document


        For i = 0 To Поток1.Count - 1
            Начало("\UvedObIzmenOklada.doc")

            oWord = CreateObject("Word.Application")
            oWord.Visible = False
            oWordDoc = oWord.Documents.Add(firthtPath & "\UvedObIzmenOklada.doc")

            '            Dim StrSql As String = "SELECT КарточкаСотрудника.НомерУведИзмОклада, КарточкаСотрудника.ДатаУведомИзмОклада
            'FROM КарточкаСотрудника
            '            Where КарточкаСотрудника.IDСотр=" & IDСотр & ""
            '            Dim ds As DataTable = Selects(StrSql)

            With oWordDoc.Bookmarks
                .Item("УвЗП1").Range.Text = ФормСобПолн & " «" & combx1 & "» "
                .Item("УвЗП2").Range.Text = ДолжОконСотр1(Поток1.Item(i)) & " " & Разряд1(Поток1.Item(i))
                .Item("УвЗП3").Range.Text = СотрИКод(Поток1.Item(i))
                .Item("УвЗП4").Range.Text = НомерУведИзмОклада1(Поток1.Item(i))
                .Item("УвЗП5").Range.Text = ДатаУведИзмОклада1(Поток1.Item(i)) & "г."
                .Item("УвЗП6").Range.Text = Trim(rich1)
                '.Item("УвЗП7").Range.Text = ds.Rows(0).Item(2).ToString
                '.Item("УвЗП8").Range.Text = TextBox1.Text
                '.Item("УвЗП9").Range.Text = TextBox2.Text
                '.Item("УвЗП10").Range.Text = ds.Rows(0).Item(1).ToString
                .Item("УвЗП11").Range.Text = ДолжРук
                If ДолжРук = "Индивидуальный предприниматель" Then
                    .Item("УвЗП12").Range.Text = ""
                Else
                    .Item("УвЗП12").Range.Text = ФормСобсКоротко & " «" & combx1 & "» "
                End If
                .Item("УвЗП13").Range.Text = ФИОрКОР
                .Item("УвЗП14").Range.Text = ФИОСотрКор1(Поток1.Item(i))
                .Item("УвЗП15").Range.Text = ДатаУведИзмОклада1(Поток1.Item(i)) & "г."
                .Item("УвЗП16").Range.Text = ФИОСотрКор1(Поток1.Item(i))
                .Item("УвЗП17").Range.Text = ДатаУведИзмОклада1(Поток1.Item(i)) & "г."
                .Item("УвЗП18").Range.Text = СотрИКод(Поток1.Item(i))
                .Item("УвЗП19").Range.Text = Trim(rich2)

                If Пол1(Поток1.Item(i)) = "М" Then
                    .Item("УвЗП20").Range.Text = "ый"
                ElseIf Пол1(Поток1.Item(i)) = "Ж" Then
                    .Item("УвЗП20").Range.Text = "ая"
                Else
                    .Item("УвЗП20").Range.Text = "ый"

                End If
            End With

            'oWordDoc.SaveAs2("Увед. " & ds.Rows(0).Item(0).ToString & " от " & ds.Rows(0).Item(1).ToString & " " & ФИОСотрКор & "(" & Trim(rich2) & ")" & ".doc",,,,,, False)
            Dim d As String = OnePath & combx1 & "\Уведомление\" & Year(Now) & "\Уведомление " & НомерУведИзмОклада1(Поток1.Item(i)) & " от " & ДатаУведИзмОклада1(Поток1.Item(i)) & " " & ФИОСотрКор1(Поток1.Item(i)) & "(" & Trim(rich2) & ")" & ".doc"
            Try
                oWordDoc.SaveAs(d,,,,,, False)
            Catch ex As Exception
                If MessageBox.Show("Уведомление с сотрудником " & ФИОСотрКор1(Поток1.Item(i)) & " существует. Заменить старый документ новым?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then
                    Try
                        IO.File.Delete(d)
                    Catch ex1 As Exception
                        MessageBox.Show("Закройте файл!", Рик)
                    End Try
                    oWordDoc.SaveAs(d,,,,,, False)
                End If
            End Try
            СохрЗак1(i) = d

            oWordDoc.Close(True)
            oWord.Quit(True)


        Next
        IO.File.Delete("C:\Users\Public\Documents\Рик\УведОбИзменОклада1.doc")
        'ОбновлПапкиУведомление()




    End Sub
    Private Sub доки2(ByVal Поток1 As List(Of Integer))


        'Dim СохрДП(ListBox1.SelectedItems.Count - 1) As String
        ReDim СохрЗак2(Поток1.Count - 1)

        Try
            If IO.Directory.Exists("c:\Users\Public\Documents\Рик") Then
                'IO.Directory.Delete("c:\Users\Public\Documents\Рик", True)
                'IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рик")
            Else
                IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рик")
            End If
        Catch ex As Exception

        End Try

        If Not IO.Directory.Exists(OnePath & combx1 & "\Уведомление\" & Year(Now)) Then
            IO.Directory.CreateDirectory(OnePath & combx1 & "\Уведомление\" & Year(Now))
        End If


        'thprov.Join()

        Try
            IO.File.Copy(OnePath & "\ОБЩДОКИ\General\UvedObIzmenOklada.doc", "C:\Users\Public\Documents\Рик\УведОбИзменОклада2.doc")
        Catch ex As Exception
            IO.File.Delete("C:\Users\Public\Documents\Рик\УведОбИзменОклада2.doc")
            IO.File.Copy(OnePath & "\ОБЩДОКИ\General\UvedObIzmenOklada.doc", "C:\Users\Public\Documents\Рик\УведОбИзменОклада2.doc")
        End Try


        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document


        For i = 0 To Поток1.Count - 1

            oWord = CreateObject("Word.Application")
            oWord.Visible = False
            oWordDoc = oWord.Documents.Add("C:\Users\Public\Documents\Рик\УведОбИзменОклада2.doc")

            With oWordDoc.Bookmarks
                .Item("УвЗП1").Range.Text = ФормСобПолн & " «" & combx1 & "» "
                .Item("УвЗП2").Range.Text = ДолжОконСотр1(Поток1.Item(i)) & " " & Разряд1(Поток1.Item(i))
                .Item("УвЗП3").Range.Text = СотрИКод(Поток1.Item(i))
                .Item("УвЗП4").Range.Text = НомерУведИзмОклада1(Поток1.Item(i))
                .Item("УвЗП5").Range.Text = ДатаУведИзмОклада1(Поток1.Item(i)) & "г."
                .Item("УвЗП6").Range.Text = Trim(rich1)
                '.Item("УвЗП7").Range.Text = ds.Rows(0).Item(2).ToString
                '.Item("УвЗП8").Range.Text = TextBox1.Text
                '.Item("УвЗП9").Range.Text = TextBox2.Text
                '.Item("УвЗП10").Range.Text = ds.Rows(0).Item(1).ToString
                .Item("УвЗП11").Range.Text = ДолжРук
                If ДолжРук = "Индивидуальный предприниматель" Then
                    .Item("УвЗП12").Range.Text = ""
                Else
                    .Item("УвЗП12").Range.Text = ФормСобсКоротко & " «" & combx1 & "» "
                End If
                .Item("УвЗП13").Range.Text = ФИОрКОР
                .Item("УвЗП14").Range.Text = ФИОСотрКор1(Поток1.Item(i))
                .Item("УвЗП15").Range.Text = ДатаУведИзмОклада1(Поток1.Item(i)) & "г."
                .Item("УвЗП16").Range.Text = ФИОСотрКор1(Поток1.Item(i))
                .Item("УвЗП17").Range.Text = ДатаУведИзмОклада1(Поток1.Item(i)) & "г."
                .Item("УвЗП18").Range.Text = СотрИКод(Поток1.Item(i))
                .Item("УвЗП19").Range.Text = Trim(rich2)

                If Пол1(Поток1.Item(i)) = "М" Then
                    .Item("УвЗП20").Range.Text = "ый"
                ElseIf Пол1(Поток1.Item(i)) = "Ж" Then
                    .Item("УвЗП20").Range.Text = "ая"
                Else
                    .Item("УвЗП20").Range.Text = "ый"

                End If
            End With

            'oWordDoc.SaveAs2("Увед. " & ds.Rows(0).Item(0).ToString & " от " & ds.Rows(0).Item(1).ToString & " " & ФИОСотрКор & "(" & Trim(rich2) & ")" & ".doc",,,,,, False)
            Dim d As String = OnePath & combx1 & "\Уведомление\" & Year(Now) & "\Уведомление " & НомерУведИзмОклада1(Поток1.Item(i)) & " от " & ДатаУведИзмОклада1(Поток1.Item(i)) & " " & ФИОСотрКор1(Поток1.Item(i)) & "(" & Trim(rich2) & ")" & ".doc"
            Try
                oWordDoc.SaveAs(d,,,,,, False)
            Catch ex As Exception
                If MessageBox.Show("Уведомление с сотрудником " & ФИОСотрКор1(Поток1.Item(i)) & " существует. Заменить старый документ новым?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then
                    Try
                        IO.File.Delete(d)
                    Catch ex1 As Exception
                        MessageBox.Show("Закройте файл!", Рик)
                    End Try
                    oWordDoc.SaveAs(d,,,,,, False)
                End If
            End Try
            СохрЗак2(i) = d

            oWordDoc.Close(True)
            oWord.Quit(True)

        Next
        IO.File.Delete("C:\Users\Public\Documents\Рик\УведОбИзменОклада2.doc")
        'ОбновлПапкиУведомление()

    End Sub

    Private Sub доки3(ByVal Поток1 As List(Of Integer))

        'Dim СохрДП(ListBox1.SelectedItems.Count - 1) As String
        ReDim СохрЗак3(Поток1.Count - 1)

        Try
            If IO.Directory.Exists("c:\Users\Public\Documents\Рик") Then
                'IO.Directory.Delete("c:\Users\Public\Documents\Рик", True)
                'IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рик")
            Else
                IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рик")
            End If
        Catch ex As Exception

        End Try

        If Not IO.Directory.Exists(OnePath & combx1 & "\Уведомление\" & Year(Now)) Then
            IO.Directory.CreateDirectory(OnePath & combx1 & "\Уведомление\" & Year(Now))
        End If


        'thprov.Join()

        Try
            IO.File.Copy(OnePath & "\ОБЩДОКИ\General\UvedObIzmenOklada.doc", "C:\Users\Public\Documents\Рик\УведОбИзменОклада3.doc")
        Catch ex As Exception
            IO.File.Delete("C:\Users\Public\Documents\Рик\УведОбИзменОклада3.doc")
            IO.File.Copy(OnePath & "\ОБЩДОКИ\General\UvedObIzmenOklada.doc", "C:\Users\Public\Documents\Рик\УведОбИзменОклада3.doc")
        End Try


        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document


        For i = 0 To Поток1.Count - 1

            oWord = CreateObject("Word.Application")
            oWord.Visible = False
            oWordDoc = oWord.Documents.Add("C:\Users\Public\Documents\Рик\УведОбИзменОклада3.doc")


            With oWordDoc.Bookmarks
                .Item("УвЗП1").Range.Text = ФормСобПолн & " «" & combx1 & "» "
                .Item("УвЗП2").Range.Text = ДолжОконСотр1(Поток1.Item(i)) & " " & Разряд1(Поток1.Item(i))
                .Item("УвЗП3").Range.Text = СотрИКод(Поток1.Item(i))
                .Item("УвЗП4").Range.Text = НомерУведИзмОклада1(Поток1.Item(i))
                .Item("УвЗП5").Range.Text = ДатаУведИзмОклада1(Поток1.Item(i)) & "г."
                .Item("УвЗП6").Range.Text = Trim(rich1)
                '.Item("УвЗП7").Range.Text = ds.Rows(0).Item(2).ToString
                '.Item("УвЗП8").Range.Text = TextBox1.Text
                '.Item("УвЗП9").Range.Text = TextBox2.Text
                '.Item("УвЗП10").Range.Text = ds.Rows(0).Item(1).ToString
                .Item("УвЗП11").Range.Text = ДолжРук
                If ДолжРук = "Индивидуальный предприниматель" Then
                    .Item("УвЗП12").Range.Text = ""
                Else
                    .Item("УвЗП12").Range.Text = ФормСобсКоротко & " «" & combx1 & "» "
                End If
                .Item("УвЗП13").Range.Text = ФИОрКОР
                .Item("УвЗП14").Range.Text = ФИОСотрКор1(Поток1.Item(i))
                .Item("УвЗП15").Range.Text = ДатаУведИзмОклада1(Поток1.Item(i)) & "г."
                .Item("УвЗП16").Range.Text = ФИОСотрКор1(Поток1.Item(i))
                .Item("УвЗП17").Range.Text = ДатаУведИзмОклада1(Поток1.Item(i)) & "г."
                .Item("УвЗП18").Range.Text = СотрИКод(Поток1.Item(i))
                .Item("УвЗП19").Range.Text = Trim(rich2)

                If Пол1(Поток1.Item(i)) = "М" Then
                    .Item("УвЗП20").Range.Text = "ый"
                ElseIf Пол1(Поток1.Item(i)) = "Ж" Then
                    .Item("УвЗП20").Range.Text = "ая"
                Else
                    .Item("УвЗП20").Range.Text = "ый"

                End If
            End With

            'oWordDoc.SaveAs2("Увед. " & ds.Rows(0).Item(0).ToString & " от " & ds.Rows(0).Item(1).ToString & " " & ФИОСотрКор & "(" & Trim(rich2) & ")" & ".doc",,,,,, False)
            Dim d As String = OnePath & combx1 & "\Уведомление\" & Year(Now) & "\Уведомление " & НомерУведИзмОклада1(Поток1.Item(i)) & " от " & ДатаУведИзмОклада1(Поток1.Item(i)) & " " & ФИОСотрКор1(Поток1.Item(i)) & "(" & Trim(rich2) & ")" & ".doc"
            Try
                oWordDoc.SaveAs(d,,,,,, False)
            Catch ex As Exception
                If MessageBox.Show("Уведомление с сотрудником " & ФИОСотрКор1(Поток1.Item(i)) & " существует. Заменить старый документ новым?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then
                    Try
                        IO.File.Delete(d)
                    Catch ex1 As Exception
                        MessageBox.Show("Закройте файл!", Рик)
                    End Try
                    oWordDoc.SaveAs(d,,,,,, False)
                End If
            End Try
            СохрЗак3(i) = d

            oWordDoc.Close(True)
            oWord.Quit(True)

        Next
        IO.File.Delete("C:\Users\Public\Documents\Рик\УведОбИзменОклада3.doc")





    End Sub

    Private Sub доки5(ByVal Поток1 As List(Of Integer))

        'Dim СохрДП(ListBox1.SelectedItems.Count - 1) As String
        ReDim СохрЗак5(Поток1.Count - 1)

        Try
            If IO.Directory.Exists("c:\Users\Public\Documents\Рик") Then
                'IO.Directory.Delete("c:\Users\Public\Documents\Рик", True)
                'IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рик")
            Else
                IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рик")
            End If
        Catch ex As Exception

        End Try

        If Not IO.Directory.Exists(OnePath & combx1 & "\Уведомление\" & Year(Now)) Then
            IO.Directory.CreateDirectory(OnePath & combx1 & "\Уведомление\" & Year(Now))
        End If


        'thprov.Join()

        Try
            IO.File.Copy(OnePath & "\ОБЩДОКИ\General\UvedObIzmenOklada.doc", "C:\Users\Public\Documents\Рик\УведОбИзменОклада5.doc")
        Catch ex As Exception
            IO.File.Delete("C:\Users\Public\Documents\Рик\УведОбИзменОклада5.doc")
            IO.File.Copy(OnePath & "\ОБЩДОКИ\General\UvedObIzmenOklada.doc", "C:\Users\Public\Documents\Рик\УведОбИзменОклада5.doc")
        End Try


        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document


        For i = 0 To Поток1.Count - 1

            oWord = CreateObject("Word.Application")
            oWord.Visible = False
            oWordDoc = oWord.Documents.Add("C:\Users\Public\Documents\Рик\УведОбИзменОклада5.doc")


            With oWordDoc.Bookmarks
                .Item("УвЗП1").Range.Text = ФормСобПолн & " «" & combx1 & "» "
                .Item("УвЗП2").Range.Text = ДолжОконСотр1(Поток1.Item(i)) & " " & Разряд1(Поток1.Item(i))
                .Item("УвЗП3").Range.Text = СотрИКод(Поток1.Item(i))
                .Item("УвЗП4").Range.Text = НомерУведИзмОклада1(Поток1.Item(i))
                .Item("УвЗП5").Range.Text = ДатаУведИзмОклада1(Поток1.Item(i)) & "г."
                .Item("УвЗП6").Range.Text = Trim(rich1)
                '.Item("УвЗП7").Range.Text = ds.Rows(0).Item(2).ToString
                '.Item("УвЗП8").Range.Text = TextBox1.Text
                '.Item("УвЗП9").Range.Text = TextBox2.Text
                '.Item("УвЗП10").Range.Text = ds.Rows(0).Item(1).ToString
                .Item("УвЗП11").Range.Text = ДолжРук
                If ДолжРук = "Индивидуальный предприниматель" Then
                    .Item("УвЗП12").Range.Text = ""
                Else
                    .Item("УвЗП12").Range.Text = ФормСобсКоротко & " «" & combx1 & "» "
                End If
                .Item("УвЗП13").Range.Text = ФИОрКОР
                .Item("УвЗП14").Range.Text = ФИОСотрКор1(Поток1.Item(i))
                .Item("УвЗП15").Range.Text = ДатаУведИзмОклада1(Поток1.Item(i)) & "г."
                .Item("УвЗП16").Range.Text = ФИОСотрКор1(Поток1.Item(i))
                .Item("УвЗП17").Range.Text = ДатаУведИзмОклада1(Поток1.Item(i)) & "г."
                .Item("УвЗП18").Range.Text = СотрИКод(Поток1.Item(i))
                .Item("УвЗП19").Range.Text = Trim(rich2)

                If Пол1(Поток1.Item(i)) = "М" Then
                    .Item("УвЗП20").Range.Text = "ый"
                ElseIf Пол1(Поток1.Item(i)) = "Ж" Then
                    .Item("УвЗП20").Range.Text = "ая"
                Else
                    .Item("УвЗП20").Range.Text = "ый"

                End If
            End With

            'oWordDoc.SaveAs2("Увед. " & ds.Rows(0).Item(0).ToString & " от " & ds.Rows(0).Item(1).ToString & " " & ФИОСотрКор & "(" & Trim(rich2) & ")" & ".doc",,,,,, False)
            Dim d As String = OnePath & combx1 & "\Уведомление\" & Year(Now) & "\Уведомление " & НомерУведИзмОклада1(Поток1.Item(i)) & " от " & ДатаУведИзмОклада1(Поток1.Item(i)) & " " & ФИОСотрКор1(Поток1.Item(i)) & "(" & Trim(rich2) & ")" & ".doc"
            Try
                oWordDoc.SaveAs(d,,,,,, False)
            Catch ex As Exception
                If MessageBox.Show("Уведомление с сотрудником " & ФИОСотрКор1(Поток1.Item(i)) & " существует. Заменить старый документ новым?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then
                    Try
                        IO.File.Delete(d)
                    Catch ex1 As Exception
                        MessageBox.Show("Закройте файл!", Рик)
                    End Try
                    oWordDoc.SaveAs(d,,,,,, False)
                End If
            End Try
            СохрЗак5(i) = d

            oWordDoc.Close(True)
            oWord.Quit(True)

        Next
        IO.File.Delete("C:\Users\Public\Documents\Рик\УведОбИзменОклада5.doc")





    End Sub
    Private Sub доки6(ByVal Поток1 As List(Of Integer))

        'Dim СохрДП(ListBox1.SelectedItems.Count - 1) As String
        ReDim СохрЗак6(Поток1.Count - 1)

        Try
            If IO.Directory.Exists("c:\Users\Public\Documents\Рик") Then
                'IO.Directory.Delete("c:\Users\Public\Documents\Рик", True)
                'IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рик")
            Else
                IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рик")
            End If
        Catch ex As Exception

        End Try

        If Not IO.Directory.Exists(OnePath & combx1 & "\Уведомление\" & Year(Now)) Then
            IO.Directory.CreateDirectory(OnePath & combx1 & "\Уведомление\" & Year(Now))
        End If


        'thprov.Join()

        Try
            IO.File.Copy(OnePath & "\ОБЩДОКИ\General\UvedObIzmenOklada.doc", "C:\Users\Public\Documents\Рик\УведОбИзменОклада6.doc")
        Catch ex As Exception
            IO.File.Delete("C:\Users\Public\Documents\Рик\УведОбИзменОклада6.doc")
            IO.File.Copy(OnePath & "\ОБЩДОКИ\General\UvedObIzmenOklada.doc", "C:\Users\Public\Documents\Рик\УведОбИзменОклада6.doc")
        End Try


        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document


        For i = 0 To Поток1.Count - 1

            oWord = CreateObject("Word.Application")
            oWord.Visible = False
            oWordDoc = oWord.Documents.Add("C:\Users\Public\Documents\Рик\УведОбИзменОклада6.doc")


            With oWordDoc.Bookmarks
                .Item("УвЗП1").Range.Text = ФормСобПолн & " «" & combx1 & "» "
                .Item("УвЗП2").Range.Text = ДолжОконСотр1(Поток1.Item(i)) & " " & Разряд1(Поток1.Item(i))
                .Item("УвЗП3").Range.Text = СотрИКод(Поток1.Item(i))
                .Item("УвЗП4").Range.Text = НомерУведИзмОклада1(Поток1.Item(i))
                .Item("УвЗП5").Range.Text = ДатаУведИзмОклада1(Поток1.Item(i)) & "г."
                .Item("УвЗП6").Range.Text = Trim(rich1)
                '.Item("УвЗП7").Range.Text = ds.Rows(0).Item(2).ToString
                '.Item("УвЗП8").Range.Text = TextBox1.Text
                '.Item("УвЗП9").Range.Text = TextBox2.Text
                '.Item("УвЗП10").Range.Text = ds.Rows(0).Item(1).ToString
                .Item("УвЗП11").Range.Text = ДолжРук
                If ДолжРук = "Индивидуальный предприниматель" Then
                    .Item("УвЗП12").Range.Text = ""
                Else
                    .Item("УвЗП12").Range.Text = ФормСобсКоротко & " «" & combx1 & "» "
                End If
                .Item("УвЗП13").Range.Text = ФИОрКОР
                .Item("УвЗП14").Range.Text = ФИОСотрКор1(Поток1.Item(i))
                .Item("УвЗП15").Range.Text = ДатаУведИзмОклада1(Поток1.Item(i)) & "г."
                .Item("УвЗП16").Range.Text = ФИОСотрКор1(Поток1.Item(i))
                .Item("УвЗП17").Range.Text = ДатаУведИзмОклада1(Поток1.Item(i)) & "г."
                .Item("УвЗП18").Range.Text = СотрИКод(Поток1.Item(i))
                .Item("УвЗП19").Range.Text = Trim(rich2)

                If Пол1(Поток1.Item(i)) = "М" Then
                    .Item("УвЗП20").Range.Text = "ый"
                ElseIf Пол1(Поток1.Item(i)) = "Ж" Then
                    .Item("УвЗП20").Range.Text = "ая"
                Else
                    .Item("УвЗП20").Range.Text = "ый"

                End If
            End With

            'oWordDoc.SaveAs2("Увед. " & ds.Rows(0).Item(0).ToString & " от " & ds.Rows(0).Item(1).ToString & " " & ФИОСотрКор & "(" & Trim(rich2) & ")" & ".doc",,,,,, False)
            Dim d As String = OnePath & combx1 & "\Уведомление\" & Year(Now) & "\Уведомление " & НомерУведИзмОклада1(Поток1.Item(i)) & " от " & ДатаУведИзмОклада1(Поток1.Item(i)) & " " & ФИОСотрКор1(Поток1.Item(i)) & "(" & Trim(rich2) & ")" & ".doc"
            Try
                oWordDoc.SaveAs(d,,,,,, False)
            Catch ex As Exception
                If MessageBox.Show("Уведомление с сотрудником " & ФИОСотрКор1(Поток1.Item(i)) & " существует. Заменить старый документ новым?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then
                    Try
                        IO.File.Delete(d)
                    Catch ex1 As Exception
                        MessageBox.Show("Закройте файл!", Рик)
                    End Try
                    oWordDoc.SaveAs(d,,,,,, False)
                End If
            End Try
            СохрЗак6(i) = d

            oWordDoc.Close(True)
            oWord.Quit(True)

        Next
        IO.File.Delete("C:\Users\Public\Documents\Рик\УведОбИзменОклада6.doc")





    End Sub
    Private Sub доки()

        'Dim СохрДП(ListBox1.SelectedItems.Count - 1) As String
        'ReDim СохрЗак4(Поток1.Count - 1)

        'Try
        '    If IO.Directory.Exists("c:\Users\Public\Documents\Рик") Then
        '        'IO.Directory.Delete("c:\Users\Public\Documents\Рик", True)
        '        'IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рик")
        '    Else
        '        IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рик")
        '    End If
        'Catch ex As Exception

        'End Try

        'If Not IO.Directory.Exists(OnePath & combx1 & "\Уведомление\" & Year(Now)) Then
        '    IO.Directory.CreateDirectory(OnePath & combx1 & "\Уведомление\" & Year(Now))
        'End If


        ''thprov.Join()

        'Try
        '    IO.File.Copy(OnePath & "\ОБЩДОКИ\General\UvedObIzmenOklada.doc", "C:\Users\Public\Documents\Рик\УведОбИзменОклада4.doc")
        'Catch ex As Exception
        '    IO.File.Delete("C:\Users\Public\Documents\Рик\УведОбИзменОклада4.doc")
        '    IO.File.Copy(OnePath & "\ОБЩДОКИ\General\UvedObIzmenOklada.doc", "C:\Users\Public\Documents\Рик\УведОбИзменОклада4.doc")
        'End Try
        Dim massFTP As New ArrayList()

        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document


        For i = 0 To ListBox1.SelectedItems.Count - 1

            oWord = CreateObject("Word.Application")
            oWord.Visible = False
            Начало("\UvedObIzmenOklada.doc")
            oWordDoc = oWord.Documents.Add(firthtPath & "\UvedObIzmenOklada.doc")

            'Dim db As New DataContext(ConString)
            ''Dim r = db.GetTable(Of Сотрудники)().Where(Function(f) f.ФИОСборное = ListBox1.SelectedItem).Select(Function(v) v.КодСотрудники)
            'Dim r = db.GetTable(Of Сотрудники)().ElementAtOrDefault()
            Dim r = dtSotrudnikiAll.Select("ФИОСборное='" & ListBox1.SelectedItems(i) & "'")

            Dim id = CType(r(0).Item("КодСотрудники").ToString, Integer)

            With oWordDoc.Bookmarks
                .Item("УвЗП1").Range.Text = ФормСобПолн & " «" & combx1 & "» "
                .Item("УвЗП2").Range.Text = Replace(ДолжОконСотр1(id), ".", "") & " " & Разряд1(id)
                .Item("УвЗП3").Range.Text = СотрИКод(id)
                .Item("УвЗП4").Range.Text = НомерУведИзмОклада1(id)
                .Item("УвЗП5").Range.Text = ДатаУведИзмОклада1(id) & "г."
                .Item("УвЗП6").Range.Text = Trim(rich1)
                '.Item("УвЗП7").Range.Text = ds.Rows(0).Item(2).ToString
                '.Item("УвЗП8").Range.Text = TextBox1.Text
                '.Item("УвЗП9").Range.Text = TextBox2.Text
                '.Item("УвЗП10").Range.Text = ds.Rows(0).Item(1).ToString
                .Item("УвЗП11").Range.Text = ДолжРук
                If ДолжРук = "Индивидуальный предприниматель" Then
                    .Item("УвЗП12").Range.Text = ""
                Else
                    .Item("УвЗП12").Range.Text = ФормСобсКоротко & " «" & combx1 & "» "
                End If
                .Item("УвЗП13").Range.Text = ФИОрКОР
                .Item("УвЗП14").Range.Text = ФИОСотрКор1(id)
                .Item("УвЗП15").Range.Text = ДатаУведИзмОклада1(id) & "г."
                .Item("УвЗП16").Range.Text = ФИОСотрКор1(id)
                .Item("УвЗП17").Range.Text = ДатаУведИзмОклада1(id) & "г."
                .Item("УвЗП18").Range.Text = СотрИКод(id)
                .Item("УвЗП19").Range.Text = Trim(rich2)

                If Пол1(id) = "М" Then
                    .Item("УвЗП20").Range.Text = "ый"
                ElseIf Пол1(id) = "Ж" Then
                    .Item("УвЗП20").Range.Text = "ая"
                Else
                    .Item("УвЗП20").Range.Text = "ый"

                End If
            End With

            Dim name As String = "Уведомление " & НомерУведИзмОклада1(id) & " от " & ДатаУведИзмОклада1(id) & " " & ФИОСотрКор1(id) & "(" & Trim(rich2) & ")" & ".doc"




            Dim fthprint As New List(Of String)
            fthprint.AddRange(New String() {combx1 & "/Уведомление/" & Now.Year & "/", name})
            massFTP.Add(fthprint)

            oWordDoc.SaveAs2(PathVremyanka & name,,,,,, False)
            oWordDoc.Close(True)
            oWord.Quit(True)
            Конец(combx1 & "\Уведомление\" & Now.Year, name, id, combx1, "\UvedObIzmenOklada.doc", "УведомлениеОбИзменинииДанныхВКонтракте")


            'oWordDoc.SaveAs2("Увед. " & ds.Rows(0).Item(0).ToString & " от " & ds.Rows(0).Item(1).ToString & " " & ФИОСотрКор & "(" & Trim(rich2) & ")" & ".doc",,,,,, False)
            'Dim d As String = OnePath & combx1 & "\Уведомление\" & Year(Now) & "\Уведомление " & НомерУведИзмОклада1(id & " от " & ДатаУведИзмОклада1(Поток1.Item(i)) & " " & ФИОСотрКор1(Поток1.Item(i)) & "(" & Trim(rich2) & ")" & ".doc"
            'Try
            '    oWordDoc.SaveAs(d,,,,,, False)
            'Catch ex As Exception
            '    If MessageBox.Show("Уведомление с сотрудником " & ФИОСотрКор1(Поток1.Item(i)) & " существует. Заменить старый документ новым?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then
            '        Try
            '            IO.File.Delete(d)
            '        Catch ex1 As Exception
            '            MessageBox.Show("Закройте файл!", Рик)
            '        End Try
            '        oWordDoc.SaveAs(d,,,,,, False)
            '    End If
            'End Try
            'СохрЗак4(i) = d

            'oWordDoc.Close(True)
            'oWord.Quit(True)

        Next


        If MessageBox.Show("Распечатать Документы?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            ПечатьДоковFTP(massFTP)
        End If


    End Sub



    Private Sub доки4(ByVal Поток1 As List(Of Integer))

        'Dim СохрДП(ListBox1.SelectedItems.Count - 1) As String
        ReDim СохрЗак4(Поток1.Count - 1)

        Try
            If IO.Directory.Exists("c:\Users\Public\Documents\Рик") Then
                'IO.Directory.Delete("c:\Users\Public\Documents\Рик", True)
                'IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рик")
            Else
                IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рик")
            End If
        Catch ex As Exception

        End Try

        If Not IO.Directory.Exists(OnePath & combx1 & "\Уведомление\" & Year(Now)) Then
            IO.Directory.CreateDirectory(OnePath & combx1 & "\Уведомление\" & Year(Now))
        End If


        'thprov.Join()

        Try
            IO.File.Copy(OnePath & "\ОБЩДОКИ\General\UvedObIzmenOklada.doc", "C:\Users\Public\Documents\Рик\УведОбИзменОклада4.doc")
        Catch ex As Exception
            IO.File.Delete("C:\Users\Public\Documents\Рик\УведОбИзменОклада4.doc")
            IO.File.Copy(OnePath & "\ОБЩДОКИ\General\UvedObIzmenOklada.doc", "C:\Users\Public\Documents\Рик\УведОбИзменОклада4.doc")
        End Try


        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document


        For i = 0 To Поток1.Count - 1

            oWord = CreateObject("Word.Application")
            oWord.Visible = False
            oWordDoc = oWord.Documents.Add("C:\Users\Public\Documents\Рик\УведОбИзменОклада4.doc")


            With oWordDoc.Bookmarks
                .Item("УвЗП1").Range.Text = ФормСобПолн & " «" & combx1 & "» "
                .Item("УвЗП2").Range.Text = ДолжОконСотр1(Поток1.Item(i)) & " " & Разряд1(Поток1.Item(i))
                .Item("УвЗП3").Range.Text = СотрИКод(Поток1.Item(i))
                .Item("УвЗП4").Range.Text = НомерУведИзмОклада1(Поток1.Item(i))
                .Item("УвЗП5").Range.Text = ДатаУведИзмОклада1(Поток1.Item(i)) & "г."
                .Item("УвЗП6").Range.Text = Trim(rich1)
                '.Item("УвЗП7").Range.Text = ds.Rows(0).Item(2).ToString
                '.Item("УвЗП8").Range.Text = TextBox1.Text
                '.Item("УвЗП9").Range.Text = TextBox2.Text
                '.Item("УвЗП10").Range.Text = ds.Rows(0).Item(1).ToString
                .Item("УвЗП11").Range.Text = ДолжРук
                If ДолжРук = "Индивидуальный предприниматель" Then
                    .Item("УвЗП12").Range.Text = ""
                Else
                    .Item("УвЗП12").Range.Text = ФормСобсКоротко & " «" & combx1 & "» "
                End If
                .Item("УвЗП13").Range.Text = ФИОрКОР
                .Item("УвЗП14").Range.Text = ФИОСотрКор1(Поток1.Item(i))
                .Item("УвЗП15").Range.Text = ДатаУведИзмОклада1(Поток1.Item(i)) & "г."
                .Item("УвЗП16").Range.Text = ФИОСотрКор1(Поток1.Item(i))
                .Item("УвЗП17").Range.Text = ДатаУведИзмОклада1(Поток1.Item(i)) & "г."
                .Item("УвЗП18").Range.Text = СотрИКод(Поток1.Item(i))
                .Item("УвЗП19").Range.Text = Trim(rich2)

                If Пол1(Поток1.Item(i)) = "М" Then
                    .Item("УвЗП20").Range.Text = "ый"
                ElseIf Пол1(Поток1.Item(i)) = "Ж" Then
                    .Item("УвЗП20").Range.Text = "ая"
                Else
                    .Item("УвЗП20").Range.Text = "ый"

                End If
            End With

            'oWordDoc.SaveAs2("Увед. " & ds.Rows(0).Item(0).ToString & " от " & ds.Rows(0).Item(1).ToString & " " & ФИОСотрКор & "(" & Trim(rich2) & ")" & ".doc",,,,,, False)
            Dim d As String = OnePath & combx1 & "\Уведомление\" & Year(Now) & "\Уведомление " & НомерУведИзмОклада1(Поток1.Item(i)) & " от " & ДатаУведИзмОклада1(Поток1.Item(i)) & " " & ФИОСотрКор1(Поток1.Item(i)) & "(" & Trim(rich2) & ")" & ".doc"
            Try
                oWordDoc.SaveAs(d,,,,,, False)
            Catch ex As Exception
                If MessageBox.Show("Уведомление с сотрудником " & ФИОСотрКор1(Поток1.Item(i)) & " существует. Заменить старый документ новым?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then
                    Try
                        IO.File.Delete(d)
                    Catch ex1 As Exception
                        MessageBox.Show("Закройте файл!", Рик)
                    End Try
                    oWordDoc.SaveAs(d,,,,,, False)
                End If
            End Try
            СохрЗак4(i) = d

            oWordDoc.Close(True)
            oWord.Quit(True)

        Next
        IO.File.Delete("C:\Users\Public\Documents\Рик\УведОбИзменОклада4.doc")

        If MessageBox.Show("Распечатать Документы?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Dim thend As New Threading.Thread(AddressOf ПечатьДоков2)
            thend.IsBackground = True
            thend.Start(СохрЗак4)
        End If


    End Sub

    Private Sub proverka(ByVal gf As Integer)
        'oWord = oWord(ListBox1.SelectedItems.Count - 1)
        'Dim oWordDoc(ListBox1.SelectedItems.Count) As Microsoft.Office.Interop.Word.Document
        'oWord = CreateObject("Word.Application")
        'oWord.Visible = False
        If gf = 0 Then
            ReDim oWord(sotkol)
            ReDim oWordDoc(sotkol)

            For i As Integer = 0 To sotkol - 1
                oWord(i) = CreateObject("Word.Application")
                oWord(i).Visible = False
                oWordDoc(i) = oWord(i).Documents.Add("C:\Users\Public\Documents\Рик\UvedObIzmenOklada.doc")
            Next
        End If

        If gf = 1 Then
            If oWord.Length = ListBox1.SelectedItems.Count Then
                Exit Sub
            Else
                thprov.Join()
                'ReDim oWord(ListBox1.SelectedItems.Count)
                'ReDim oWordDoc(ListBox1.SelectedItems.Count)

                For i As Integer = (ListBox1.SelectedItems.Count) To oWord.Length
                    oWordDoc(i).Close(True)
                    oWord(i).Quit(True)
                Next
            End If

        End If

        If gf = 3 Then
            ReDim oWord(ListBox1.SelectedItems.Count)
            ReDim oWordDoc(ListBox1.SelectedItems.Count)

            For i As Integer = 0 To ListBox1.SelectedItems.Count - 1
                oWord(i) = CreateObject("Word.Application")
                oWord(i).Visible = False
                oWordDoc(i) = oWord(i).Documents.Add("C:\Users\Public\Documents\Рик\UvedObIzmenOklada.doc")
            Next
        End If






    End Sub
    Private Function Сортировка() As ArrayList
        Dim Поток1 As Integer
        Dim Поток2 As Integer
        Dim Поток3 As Integer
        Dim Поток4 As Integer
        Dim Поток5 As Integer
        Dim Сотр1 As New List(Of Integer)()
        Dim Сотр2 As New List(Of Integer)()
        Dim Сотр3 As New List(Of Integer)()
        Dim Сотр4 As New List(Of Integer)()
        Dim Сотр5 As New List(Of Integer)()
        Dim СотрОбщ As New ArrayList()

        Dim df As Integer = ListBox3.SelectedItems.Count
        If df < 5 Then
            For i As Integer = 0 To ListBox3.SelectedItems.Count - 1
                Сотр1.Add(ListBox3.SelectedItems(i))
            Next
            СотрОбщ.Add(Сотр1)
            Return СотрОбщ

        ElseIf df Mod 5 > 0 Then 'есть остаток
            Dim dh As Integer = Math.Floor(df / 5) 'округлем до ближайщего меньшего целого
            Dim da As Integer = df - (dh * 5)
            Сотр1.Clear()
            Поток1 = dh + da
            Поток2 = dh
            Поток3 = dh
            Поток4 = dh
            Поток5 = dh

            For i As Integer = 0 To ListBox3.SelectedItems.Count - 1
                If i < Поток1 Then
                    Сотр1.Add(ListBox3.SelectedItems(i))

                ElseIf i >= Поток1 And i < (Поток2 + Поток1) Then
                    Сотр2.Add(ListBox3.SelectedItems(i))

                ElseIf i >= (Поток2 + Поток1) And i <= Поток2 + Поток1 + Поток3 Then
                    Сотр3.Add(ListBox3.SelectedItems(i))

                ElseIf i >= Поток2 + Поток1 + Поток3 And i <= Поток2 + Поток1 + Поток3 + Поток4 Then
                    Сотр4.Add(ListBox3.SelectedItems(i))

                ElseIf i >= Поток2 + Поток1 + Поток3 + Поток4 And i <= ListBox3.SelectedItems.Count - 1 Then
                    Сотр5.Add(ListBox3.SelectedItems(i))
                End If
            Next

            СотрОбщ.AddRange({Сотр1, Сотр2, Сотр3, Сотр4, Сотр5})
            Return СотрОбщ
        Else
            Сотр1.Clear()
            Поток1 = df / 5
            Поток2 = df / 5
            Поток3 = df / 5
            Поток4 = df / 5
            Поток5 = df / 5

            For i As Integer = 0 To ListBox3.SelectedItems.Count - 1
                If i < Поток1 Then
                    Сотр1.Add(ListBox3.SelectedItems(i))

                ElseIf i >= Поток1 And i < (Поток2 + Поток1) Then
                    Сотр2.Add(ListBox3.SelectedItems(i))

                ElseIf i >= (Поток2 + Поток1) And i <= Поток2 + Поток1 + Поток3 Then
                    Сотр3.Add(ListBox3.SelectedItems(i))

                ElseIf i >= Поток2 + Поток1 + Поток3 And i <= Поток2 + Поток1 + Поток3 + Поток4 Then
                    Сотр4.Add(ListBox3.SelectedItems(i))

                ElseIf i >= Поток2 + Поток1 + Поток3 + Поток4 And i <= ListBox3.SelectedItems.Count - 1 Then
                    Сотр5.Add(ListBox3.SelectedItems(i))
                End If
            Next
            СотрОбщ.AddRange({Сотр1, Сотр2, Сотр3, Сотр4, Сотр5})
            Return СотрОбщ
        End If


    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'If oWord.Count = 0 Then
        '    proverka(3)
        'End If

        'proverka(1)

        If MessageBox.Show("Сформировать уведомления?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Exit Sub
        End If
        Me.Cursor = Cursors.WaitCursor
        combx1 = ComboBox1.Text
        mskbx1 = MaskedTextBox1.Text
        rich1 = RichTextBox1.Text
        rich2 = RichTextBox2.Text


        Dim sd As Integer = ПровЗаполн()


        If sd = 1 Then
            Exit Sub
        End If







        'Dim obj As New ArrayList()
        'obj = Сортировка()
        'Dim Поток1 As New List(Of Integer)()
        'Dim Поток2 As New List(Of Integer)()
        'Dim Поток3 As New List(Of Integer)()
        'Dim Поток4 As New List(Of Integer)()
        'Dim Поток5 As New List(Of Integer)()

        'Поток1 = obj.Item(0)
        'If obj.Count = 1 Then
        '    Поток1 = obj.Item(0)
        'Else
        '    Поток1 = obj.Item(0)
        '    Поток2 = obj.Item(1)
        '    Поток3 = obj.Item(2)
        '    Поток4 = obj.Item(3)
        '    Поток5 = obj.Item(4)
        'End If



        'Dim ПотокНов1 As New Thread(AddressOf доки1)
        'ПотокНов1.IsBackground = True
        'Dim ПотокНов2 As New Thread(AddressOf доки2)
        'ПотокНов2.IsBackground = True
        'Dim ПотокНов3 As New Thread(AddressOf доки3)
        'ПотокНов3.IsBackground = True
        'Dim ПотокНов4 As New Thread(AddressOf доки5)
        'ПотокНов4.IsBackground = True
        'Dim ПотокНов5 As New Thread(AddressOf доки6)
        'ПотокНов5.IsBackground = True







        Dim list As New Dictionary(Of String, Object)()        '
        list.Add("@НазвОрганиз", ComboBox1.Text)

        Dim ds = Selects(StrSql:="Select COUNT (КарточкаСотрудника.НомерУведИзмОклада)
        FROM Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр
        WHERE Сотрудники.НазвОрганиз=@НазвОрганиз and Сотрудники.НаличеДогПодряда='Нет'", list)

        cl = ds.Rows(0).Item(0)

        база()


        доки()
        Me.Cursor = Cursors.Default
        ОбновлПапкиУведомление()

        Exit Sub







        'If obj.Count = 1 Then
        '    доки4(Поток1)
        '    Me.Cursor = Cursors.Default
        '    ОбновлПапкиУведомление()
        '    Exit Sub
        'Else
        '    ПотокНов1.Start(Поток1)
        '    ПотокНов2.Start(Поток2)
        '    ПотокНов3.Start(Поток3)
        '    ПотокНов4.Start(Поток4)
        '    ПотокНов5.Start(Поток5)

        'End If

        Me.Cursor = Cursors.Default

        'ПотокНов1.Join()
        'ПотокНов2.Join()
        'ПотокНов3.Join()
        'ПотокНов4.Join()
        'ПотокНов5.Join()
        ОбновлПапкиУведомление()
        If MessageBox.Show("Распечатать Документы?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            'Dim s1 As Integer = СохрЗак1.Count - 1
            'Dim s2 As Integer = СохрЗак2.Count - 1
            'Dim s3 As Integer = СохрЗак3.Count - 1
            'Dim s4 As Integer = СохрЗак5.Count - 1
            'Dim s5 As Integer = СохрЗак6.Count - 1


            'ПечатьДоковКол(СохрЗак1(s1), 1)
            'ПечатьДоковКол(СохрЗак2(s2), 1)
            'ПечатьДоковКол(СохрЗак3(s3), 1)
            'ПечатьДоковКол(СохрЗак5(s4), 1)
            'ПечатьДоковКол(СохрЗак6(s5), 1)
        End If
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub ListBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedValueChanged
        ListBox3.SelectedIndices.Clear()
        ListBox3.SelectedIndex = ListBox1.SelectedIndex

        If ListBox1.SelectedItems.Count > 1 Then
            For Each r In ListBox1.SelectedIndices
                ListBox3.SetSelected(r, True)
            Next
        End If

    End Sub
End Class
