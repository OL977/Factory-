Option Explicit On
Imports System.Data.OleDb

Public Class ДопПоСрокамОплаты
    Public ds As DataTable
    Dim StrSql, ФИОРукРодПад, ФИОПолнРук, ФИОрКОР, ФИОСотрКор, ФормСобсКоротко, ДолжДирСОконч,
        ДолжОконСотр, Разряд, inp, ДатаКонтр, НомерКонтр, ФормСобПолн, ДолжРук, ОснДейств, ДатОкон As String
    Dim cl, IDСотр, errs As Integer
    Dim СохрFTP As New List(Of String)

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        For i = 0 To ListBox2.SelectedItems.Count - 1
            Чист()
            StrSql = "SELECT Сотрудники.КодСотрудники FROM Сотрудники 
WHERE Сотрудники.НазвОрганиз='" & ComboBox1.Text & "' AND Сотрудники.ФИОСборное='" & ListBox2.SelectedItems(i) & "'"
            ds = Selects(StrSql)
            IDСотр = Nothing
            IDСотр = ds.Rows(0).Item(0)


            Чист()
            StrSql = "UPDATE КарточкаСотрудника SET НомУведИзмСрокЗарп=" & 0 & ", ДатаСогласияНаИзмен='', ДатаВсуплСоглаш='', ДатаУведом=''
WHERE КарточкаСотрудника.IDСотр=" & IDСотр & ""
            Updates(StrSql)
        Next
        MessageBox.Show("Данные согласно Номера уведомления, даты согласования очищены удачно!", Рик)
    End Sub

    Private Sub MaskedTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox1.Focus()
        End If
    End Sub

    Private Sub Заработная_плата_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MdiParent = MDIParent1
        Me.WindowState = FormWindowState.Maximized

        Me.ComboBox1.AutoCompleteCustomSource.Clear()
        Me.ComboBox1.Items.Clear()
        For Each r As DataRow In СписокКлиентовОсновной.Rows
            Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox1.Items.Add(r(0).ToString)
        Next



    End Sub



    Private Sub Очист()

        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        MaskedTextBox1.Text = ""

    End Sub
    Private Sub refreshes()
        'Чист()
        'StrSql = "SELECT ФИОСборное FROM Сотрудники WHERE НазвОрганиз='" & ComboBox1.Text & "' ORDER BY ФИОСборное"
        'ds = Selects(StrSql)
        Dim ds = From x In dtSotrudnikiAll Where x.Item("НазвОрганиз") = ComboBox1.Text Order By x.Item("ФИОСборное") Select x
        For Each r As DataRow In ds
            ListBox1.Items.Add(r.Item("ФИОСборное").ToString())
            ListBox2.Items.Add(r.Item("ФИОСборное").ToString())
        Next
    End Sub
    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        Очист()
        refreshes()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ListBox1.Items.Clear()
    End Sub
    Private Function ПровЗаполн()
        If ComboBox1.Text = "" Then
            MessageBox.Show("Выберите организацию", Рик)
            Return 1
        End If

        If TextBox1.Text = "" Then
            If MessageBox.Show("Выберите дату аванса!" & vbCrLf & "Если дата аванса ненужна выберите ОТМЕНА!", Рик, MessageBoxButtons.OKCancel) = DialogResult.OK Then
                Return 1
            End If
        End If

        If TextBox2.Text = "" Then
            MessageBox.Show("Выберите дату выплаты зарплаты!", Рик)
            Return 1
        End If

        If TextBox4.Text = "" Then
            MessageBox.Show("Выберите номер пункта договора, в который внсятся изменения!", Рик)
            Return 1
        End If
        If TextBox3.Text = "" Then
            MessageBox.Show("Заполните текст поля, в который внсятся изменения!", Рик)
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


        For i = 0 To ListBox1.SelectedItems.Count - 1

            Dim f = dtSotrudnikiAll.Select("НазвОрганиз='" & ComboBox1.Text & "' AND ФИОСборное='" & ListBox1.SelectedItems(i) & "'")

            '            Чист()
            '            StrSql = "SELECT Сотрудники.КодСотрудники FROM Сотрудники 
            'WHERE Сотрудники.НазвОрганиз='" & ComboBox1.Text & "' AND Сотрудники.ФИОСборное='" & ListBox1.SelectedItems(i) & "'"
            '            ds = Selects(StrSql)
            IDСотр = Nothing
            IDСотр = CType(f(0).Item("КодСотрудники").ToString, Integer)

            Dim cld As String = ""
            cl = cl + 1
            Select Case cl
                Case < 10
                    cld = "00" & CType(cl, String)
                Case 10 To 99
                    cld = "0" & CType(cl, String)
                Case > 99
                    cld = CType(cl, String)
            End Select

            Dim list As New Dictionary(Of String, Object)
            list.Add("@IDСотр", IDСотр)


            Updates(stroka:="UPDATE КарточкаСотрудника SET ДатаЗарплаты='" & TextBox2.Text & "', ДатаАванса='" & TextBox1.Text & "', 
НомУведИзмСрокЗарп='" & cld & "',ДатаСогласияНаИзмен='" & datsog & "',ДатаВсуплСоглаш='" & datsogl & "', ДатаУведом='" & MaskedTextBox1.Text & "'
WHERE КарточкаСотрудника.IDСотр=@IDСотр", list, "КарточкаСотрудника")
        Next

    End Sub
    Private Sub Выборка()
        'Dim list As New Dictionary(Of String, Object)
        'list.Add("@НазвОрг", ComboBox1.Text)

        Dim ds = dtClientAll.Select("НазвОрг='" & ComboBox1.Text & "'")


        '        Dim ds = Selects(StrSql:="Select Клиент.ФормаСобств, Клиент.ДолжнРуководителя,
        'Клиент.ФИОРуководителя, Клиент.ФИОРукРодПадеж, Клиент.РукИП, Клиент.ОснованиеДейств
        'From Клиент Where Клиент.НазвОрг =@НазвОрг", list)


        Dim РуковИП As String
        If ds(0).Item("РукИП") = "True" Then
            РуковИП = "ИП "
        Else
            РуковИП = ""
        End If
        ФормСобПолн = ds(0).Item("ФормаСобств").ToString
        ФИОПолнРук = ds(0).Item("ФИОРуководителя").ToString
        ФИОРукРодПад = РуковИП & ds(0).Item("ФИОРукРодПадеж").ToString
        ФИОрКОР = ФИОКорРук(ФИОПолнРук, ds(0).Item("РукИП"))
        ФормСобсКоротко = ФормСобствКор(ds(0).Item("ФормаСобств").ToString)
        ДолжДирСОконч = ДобОконч(ds(0).Item("ДолжнРуководителя").ToString)
        ДолжРук = ds(0).Item("ДолжнРуководителя").ToString
        ОснДейств = ds(0).Item("ОснованиеДейств").ToString
    End Sub
    Private Sub ДанСотр(ByVal IDсот As Integer)

        ' выбираем данные по сотруднику
        Dim list As New Dictionary(Of String, Object)
        list.Add("@IDСотр", IDсот)
        list.Add("@ИДСотр", IDсот)


        Dim ds = Selects(StrSql:="SELECT Штатное.Должность, Штатное.Разряд, ДогСотрудн.ДатаКонтракта, ДогСотрудн.Контракт, Сотрудники.ФИОСборное
FROM (Сотрудники INNER JOIN ДогСотрудн ON Сотрудники.КодСотрудники = ДогСотрудн.IDСотр) INNER JOIN Штатное ON Сотрудники.КодСотрудники = Штатное.ИДСотр
Where Штатное.ИДСотр =ИДСотр And ДогСотрудн.IDСотр =@IDСотр", list)

        Dim f = dtSotrudnikiAll.Select("КодСотрудники=" & IDсот & "")

        Try
            'ДолжОконСотр = ДолжРодПадежФункц(ds.Rows(0).Item(0).ToString)
            ДолжОконСотр = ds.Rows(0).Item(0).ToString
        Catch ex As Exception
            MessageBox.Show("У сотрудника " & f(0).Item("ФИОСборное").ToString & " нет должности!", Рик)
            'errs = 1
            Exit Sub

        End Try


        ФИОСотрКор = ФИОКорРук(ds.Rows(0).Item(4).ToString, False)

        If ds.Rows(0).Item(1) <> "" Then
            Try
                Разряд = разрядстрока(CType(ds.Rows(0).Item(1), Integer))
            Catch ex As Exception
                Разряд = ""
            End Try
        Else
            Разряд = ""
        End If

        ДатаКонтр = Strings.Left(ds.Rows(0).Item(2), 10)
        НомерКонтр = ds.Rows(0).Item(3).ToString


        'Dim StrSql2 As String = "SELECT ФИОДатПадКому FROM Сотрудники Where Сотрудники.КодСотрудники = " & IDсот & ""
        'Dim ds2 As DataTable = Selects(StrSql2)

        'If Not ds2.Rows(0).Item(0).ToString <> "" Then
        '    inp = InputBox("Введите ФИО сотрудника " & ds.Rows(0).Item(4).ToString & " в Дательном падеже (Уведомление)'Кому?'", Рик)

        '    Do Until inp <> ""
        '        MessageBox.Show("Повторите ввод ФИО!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Error)
        '        inp = InputBox("Введите ФИО сотрудника " & ds.Rows(0).Item(4).ToString & " в Дательном падеже (Уведомление)'Кому?'", Рик)

        '    Loop
        '    Чист()
        '    StrSql = "UPDATE Сотрудники SET ФИОДатПадКому='" & inp & "' Where Сотрудники.КодСотрудники = " & IDсот & ""
        '    Updates(StrSql)
        'Else
        '    inp = ds2.Rows(0).Item(0).ToString
        'End If


    End Sub
    Private Sub доки()



        Выборка()
        Dim СохрДП(ListBox1.SelectedItems.Count - 1) As String
        Dim СохрЗак(ListBox1.SelectedItems.Count - 1) As String
        Dim massFTP As New ArrayList()



        For i = 0 To ListBox1.SelectedItems.Count - 1

            Dim ds1 = dtSotrudnikiAll.Select("НазвОрганиз='" & ComboBox1.Text & "' AND ФИОСборное='" & ListBox1.SelectedItems(i) & "'")
            '            StrSql = "SELECT КодСотрудники, Пол 
            'FROM Сотрудники WHERE Сотрудники.НазвОрганиз='" & ComboBox1.Text & "' AND Сотрудники.ФИОСборное='" & ListBox1.SelectedItems(i) & "'"
            '            ds = Selects(StrSql)
            IDСотр = Nothing
            IDСотр = CType(ds1(0).Item("КодСотрудники").ToString, Integer)
            Dim Пол As String = ds1(0).Item("Пол").ToString

            ДанСотр(IDСотр)
            If errs = 1 Then Exit Sub


            Dim ds = dtKartochkaSotrudnikaAll.Select("IDСотр=" & IDСотр & "")
            If ds.Length = 0 Then Continue For
            'StrSql = "SELECT КарточкаСотрудника.НомУведИзмСрокЗарп, КарточкаСотрудника.ДатаСогласияНаИзмен,
            '    КарточкаСотрудника.ДатаВсуплСоглаш, КарточкаСотрудника.ДатаУведом FROM КарточкаСотрудника
            'Where КарточкаСотрудника.IDСотр=" & IDСотр & ""
            'ds = Selects(StrSql)

            Dim СохрFTP As New List(Of String)

            Dim oWord As Microsoft.Office.Interop.Word.Application
            Dim oWord1 As Microsoft.Office.Interop.Word.Application
            Dim oWordDoc As Microsoft.Office.Interop.Word.Document
            Dim oWordDoc1 As Microsoft.Office.Interop.Word.Document

            oWord = CreateObject("Word.Application")
            oWord.Visible = False

            Начало("UvedObIzmenSrokovVyplZarpliAvansa.doc")

            oWordDoc = oWord.Documents.Add(firthtPath & "\UvedObIzmenSrokovVyplZarpliAvansa.doc")

            With oWordDoc.Bookmarks
                .Item("УвЗП1").Range.Text = ФормСобПолн & " «" & ComboBox1.Text & "» "
                .Item("УвЗП2").Range.Text = ДолжОконСотр & " " & Разряд
                .Item("УвЗП3").Range.Text = ListBox1.SelectedItems(i).ToString
                .Item("УвЗП4").Range.Text = ds(0).Item("НомУведИзмСрокЗарп")
                .Item("УвЗП5").Range.Text = ds(0).Item("ДатаУведом").ToString
                .Item("УвЗП6").Range.Text = ФормСобсКоротко & " «" & ComboBox1.Text & "» "
                .Item("УвЗП7").Range.Text = ds(0).Item("ДатаВсуплСоглаш").ToString
                .Item("УвЗП8").Range.Text = TextBox1.Text
                .Item("УвЗП9").Range.Text = TextBox2.Text
                .Item("УвЗП10").Range.Text = ds(0).Item("ДатаСогласияНаИзмен").ToString
                .Item("УвЗП11").Range.Text = ДолжРук
                .Item("УвЗП12").Range.Text = ФормСобсКоротко & " «" & ComboBox1.Text & "» "
                .Item("УвЗП13").Range.Text = ФИОрКОР
                .Item("УвЗП14").Range.Text = ФИОСотрКор
                .Item("УвЗП15").Range.Text = ds(0).Item("ДатаУведом").ToString
                .Item("УвЗП16").Range.Text = ФИОСотрКор
                .Item("УвЗП17").Range.Text = ds(0).Item("ДатаУведом").ToString
                .Item("УвЗП18").Range.Text = ListBox1.SelectedItems(i).ToString
                .Item("УвЗП19").Range.Text = ДатОкон
                If Пол = "М" Then
                    .Item("УвЗП20").Range.Text = "ый"
                ElseIf Пол = "Ж" Then
                    .Item("УвЗП20").Range.Text = "ая"
                Else
                    .Item("УвЗП20").Range.Text = "ый"

                End If
            End With



            'If Not IO.Directory.Exists(OnePath & ComboBox1.Text & "\Уведомление\" & Year(Now)) Then
            '    IO.Directory.CreateDirectory(OnePath & ComboBox1.Text & "\Уведомление\" & Year(Now))
            'End If

            Dim name As String = ds(0).Item("НомУведИзмСрокЗарп").ToString & " от " & ds(0).Item("ДатаУведом").ToString & " " & ФИОСотрКор & "(Изм.Сроков.Вып.Зп.)" & ".doc"

            СохрFTP.AddRange(New String() {ComboBox1.Text & "\Уведомление\" & Now.Year, name})
            massFTP.Add(СохрFTP)
            oWordDoc.SaveAs2(PathVremyanka & name,,,,,, False)
            Конец(ComboBox1.Text & "\Уведомление\" & Now.Year, name, IDСотр, ComboBox1.Text, "\UvedObIzmenSrokovVyplZarpliAvansa.doc", "Изм.Сроков.Вып.Зп.")
            oWordDoc.Close(True)
            oWord.Quit(True)

            'СохрЗак(i) = "C:\Users\Public\Documents\Рик\Уведомление " & ds.Rows(0).Item(0).ToString & " от " & ds.Rows(0).Item(3).ToString & " " & ФИОСотрКор & "(Изм.Сроков.Вып.Зп.)" & ".doc"
            ''oWordDoc.SaveAs2("U: \Офис\Финансовый\6. Бух.услуги\Кадры\" & Клиент & "\Заявление\" & Год & "\" & Заявление(9) & " (заявление)" & ".doc",,,,,, False)
            'Try
            '    IO.File.Copy("C:\Users\Public\Documents\Рик\Уведомление " & ds.Rows(0).Item(0).ToString & " от " & ds.Rows(0).Item(3).ToString & " " & ФИОСотрКор & "(Изм.Сроков.Вып.Зп.)" & ".doc", OnePath & ComboBox1.Text & "\Уведомление\" & Year(Now) & "\Уведомление " & ds.Rows(0).Item(0).ToString & " от " & ds.Rows(0).Item(3).ToString & " " & ФИОСотрКор & "(Изм.Сроков.Вып.Зп.)" & ".doc")
            'Catch ex As Exception
            '    If MessageBox.Show("Уведомление с сотрудником " & ФИОСотрКор & " существует. Заменить старый документ новым?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then
            '        Try
            '            IO.File.Delete(OnePath & ComboBox1.Text & "\Уведомление\" & Year(Now) & "\Уведомление " & ds.Rows(0).Item(0).ToString & " от " & ds.Rows(0).Item(3).ToString & " " & ФИОСотрКор & "(Изм.Сроков.Вып.Зп.)" & ".doc")
            '        Catch ex1 As Exception
            '            MessageBox.Show("Закройте файл!", Рик)
            '        End Try
            '        IO.File.Copy("C:\Users\Public\Documents\Рик\Уведомление " & ds.Rows(0).Item(0).ToString & " от " & ds.Rows(0).Item(3).ToString & " " & ФИОСотрКор & "(Изм.Сроков.Вып.Зп.)" & ".doc", OnePath & ComboBox1.Text & "\Уведомление\" & Year(Now) & "\Уведомление " & ds.Rows(0).Item(0).ToString & " от " & ds.Rows(0).Item(3).ToString & " " & ФИОСотрКор & "(Изм.Сроков.Вып.Зп.)" & ".doc")
            '    End If
            'End Try
            'oWordDoc.Close(True)

            'Try
            '    IO.File.Copy(OnePath & "\ОБЩДОКИ\General\DopSiglObIzmenSrokovVyplZarplaty.doc", "C:\Users\Public\Documents\Рик\DopSiglObIzmenSrokovVyplZarplaty.doc")
            'Catch ex As Exception
            '    If "DopSiglObIzmenSrokovVyplZarplaty.doc" <> "" Then IO.File.Delete("C:\Users\Public\Documents\Рик\DopSiglObIzmenSrokovVyplZarplaty.doc")
            '    IO.File.Copy(OnePath & "\ОБЩДОКИ\General\DopSiglObIzmenSrokovVyplZarplaty.doc", "C:\Users\Public\Documents\Рик\DopSiglObIzmenSrokovVyplZarplaty.doc")
            'End Try


            oWord1 = CreateObject("Word.Application")
            oWord1.Visible = False


            Начало("DopSiglObIzmenSrokovVyplZarplaty.doc")
            oWordDoc1 = oWord1.Documents.Add(firthtPath & "\DopSiglObIzmenSrokovVyplZarplaty.doc")
            With oWordDoc1.Bookmarks
                .Item("ДпЗарпИзм1").Range.Text = ДатаКонтр
                .Item("ДпЗарпИзм2").Range.Text = НомерКонтр
                .Item("ДпЗарпИзм3").Range.Text = ds(0).Item("ДатаСогласияНаИзмен").ToString
                .Item("ДпЗарпИзм4").Range.Text = ФормСобПолн & " «" & ComboBox1.Text & "» "
                .Item("ДпЗарпИзм5").Range.Text = ДолжДирСОконч & " " & ФИОРукРодПад
                .Item("ДпЗарпИзм6").Range.Text = ОснДейств
                .Item("ДпЗарпИзм7").Range.Text = ListBox1.SelectedItems(i).ToString
                .Item("ДпЗарпИзм8").Range.Text = TextBox4.Text
                .Item("ДпЗарпИзм9").Range.Text = TextBox3.Text
                .Item("ДпЗарпИзм10").Range.Text = ds(0).Item("ДатаВсуплСоглаш").ToString
                .Item("ДпЗарпИзм11").Range.Text = ФормСобсКоротко
                .Item("ДпЗарпИзм12").Range.Text = ComboBox1.Text
                .Item("ДпЗарпИзм13").Range.Text = ФИОрКОР
                .Item("ДпЗарпИзм14").Range.Text = ФИОСотрКор

            End With

            СохрFTP.Clear()

            Dim name1 As String = "ДопСогл " & ds(0).Item("НомУведИзмСрокЗарп").ToString & " от " & ds(0).Item("ДатаУведом").ToString & " " & ФИОСотрКор & "(Изм.Сроков.Вып.Зп.)" & ".doc"

            СохрFTP.AddRange(New String() {ComboBox1.Text & "\DopSoglashenie\" & Now.Year, name1})
            massFTP.Add(СохрFTP)
            oWordDoc1.SaveAs2(PathVremyanka & name1,,,,,, False)
            Конец(ComboBox1.Text & "\Уведомление\" & Now.Year, name1, IDСотр, ComboBox1.Text, "\DopSiglObIzmenSrokovVyplZarplaty.doc", "Уведомление.Изм.Сроков.Вып.Зп.")
            oWordDoc1.Close(True)
            oWord1.Quit(True)



            'If Not IO.Directory.Exists(OnePath & ComboBox1.Text & "\DopSoglashenie\" & Year(Now)) Then
            '    IO.Directory.CreateDirectory(OnePath & ComboBox1.Text & "\DopSoglashenie\" & Year(Now))
            'End If
            'oWordDoc.SaveAs2("C:\Users\Public\Documents\Рик\ДопСогл " & ds(0).Item("НомУведИзмСрокЗарп").ToString & " от " & ds.Rows(0).Item(3).ToString & " " & ФИОСотрКор & "(Изм.Сроков.Вып.Зп.)" & ".doc",,,,,, False)
            'СохрДП(i) = "C:\Users\Public\Documents\Рик\ДопСогл " & ds(0).Item("НомУведИзмСрокЗарп").ToString & " от " & ds.Rows(0).Item(3).ToString & " " & ФИОСотрКор & "(Изм.Сроков.Вып.Зп.)" & ".doc"
            ''oWordDoc.SaveAs2("U: \Офис\Финансовый\6. Бух.услуги\Кадры\" & Клиент & "\Заявление\" & Год & "\" & Заявление(9) & " (заявление)" & ".doc",,,,,, False)
            'Try
            '    IO.File.Copy("C:\Users\Public\Documents\Рик\ДопСогл " & ds.Rows(0).Item(0).ToString & " от " & ds.Rows(0).Item(3).ToString & " " & ФИОСотрКор & "(Изм.Сроков.Вып.Зп.)" & ".doc", OnePath & ComboBox1.Text & "\DopSoglashenie\" & Year(Now) & "\ДопСогл " & ds.Rows(0).Item(0).ToString & " от " & ds.Rows(0).Item(3).ToString & " " & ФИОСотрКор & "(Изм.Сроков.Вып.Зп.)" & ".doc")
            'Catch ex As Exception
            '    If MessageBox.Show("DopSoglashenie с сотрудником " & ФИОСотрКор & " существует. Заменить старый документ новым?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then
            '        Try
            '            IO.File.Delete(OnePath & ComboBox1.Text & "\DopSoglashenie\" & Year(Now) & "\ДопСогл " & ds(0).Item("НомУведИзмСрокЗарп").ToString & " от " & ds.Rows(0).Item(3).ToString & " " & ФИОСотрКор & "(Изм.Сроков.Вып.Зп.)" & ".doc")
            '        Catch ex1 As Exception
            '            MessageBox.Show("Закройте файл!", Рик)
            '        End Try
            '        IO.File.Copy("C:\Users\Public\Documents\Рик\ДопСогл " & ds.Rows(0).Item(0).ToString & " от " & ds.Rows(0).Item(3).ToString & " " & ФИОСотрКор & "(Изм.Сроков.Вып.Зп.)" & ".doc", OnePath & ComboBox1.Text & "\DopSoglashenie\" & Year(Now) & "\ДопСогл " & ds.Rows(0).Item(0).ToString & " от " & ds.Rows(0).Item(3).ToString & " " & ФИОСотрКор & "(Изм.Сроков.Вып.Зп.)" & ".doc")
            '    End If
            ''End Try
            'oWordDoc.Close(True)
            'oWord.Quit(True)

            'If MessageBox.Show("Данные по сотруднику " & ФИОСотрКор & " внесены! Распечатать Документы? ", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.None) = DialogResult.OK Then
            '    Dim mass() As String = {СохрДП, СохрЗак}
            '    ПечатьДоков(mass)
            'End If
            'MessageBox.Show(ФИОСотрКор & " - OK!", Рик)
        Next

        'Dim mass() As String = {СохрДП(ListBox1.SelectedItems.Count - 1), СохрЗак(ListBox1.SelectedItems.Count - 1)}
        'ПечатьДоков2(СохрДП, СохрЗак)



        ПечатьДоковFTP(massFTP)


    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim sd As Integer = ПровЗаполн()
        If sd = 1 Then
            Exit Sub
        End If
        Me.Cursor = Cursors.WaitCursor

        Dim list As New Dictionary(Of String, Object)
        list.Add("@НазвОрганиз", ComboBox1.Text)


        Dim ds = Selects(StrSql:= "Select COUNT (КарточкаСотрудника.НомУведИзмСрокЗарп)
FROM Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр
WHERE Сотрудники.НазвОрганиз=@НазвОрганиз", list)

        cl = ds.Rows(0).Item(0)
            база()
            доки()

            Me.Cursor = Cursors.Default




    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox2.Focus()
        End If
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox4.Focus()
        End If
    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox3.Focus()
        End If
    End Sub


End Class