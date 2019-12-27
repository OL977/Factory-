Option Explicit On
Imports System.Data.OleDb
Imports System.Threading
Imports System.Data.SqlClient

Public Class Увольнение
    Public Da As New OleDbDataAdapter 'Адаптер
    Public Ds As New DataSet 'Пустой набор записей
    Dim tbl As New DataTable
    Dim tbl2 As New DataTable
    Dim cb As OleDb.OleDbCommandBuilder
    Dim dc9 As DataTable
    Dim ds4(), ds6() As DataRow
    Dim Сотруд As String
    Dim ID, IDСотр As Integer
    Dim Org, Год, Организ, ФамСотрРодПад, ПрикУвольн, СохрЗак As String
    Dim dsДанСотр As DataTable
    Dim arrtbox As New Dictionary(Of String, String)
    Dim arrtcom As New Dictionary(Of String, String)
    Dim arrtmask As New Dictionary(Of String, String)
    Dim ПрикУвольнFTP As New List(Of String)
    Dim СохрЗакFTP As New List(Of String)
    Dim idsotrudnika As Integer
    Private Sub Form5_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MdiParent = MDIParent1

        Dim Год As String = Year(Now)


        ComboBox1.AutoCompleteCustomSource.Clear()
        Me.ComboBox1.Items.Clear()
        For Each r As DataRow In СписокКлиентовОсновной.Rows
            Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox1.Items.Add(r(0).ToString)
        Next

        Me.TextBox1.Text = DateTime.Now.ToString("dd.MM.yyyy")
        Me.MaskedTextBox2.Text = DateTime.Now.ToString("dd.MM.yyyy")
        Me.MaskedTextBox1.Text = DateTime.Now.ToString("dd.MM.yyyy")
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Org = ComboBox1.Text

        'Dim d As New Thread(AddressOf СборДанОрг)
        'd.IsBackground = True
        'd.Start()

        Parallel.Invoke(Sub() СборДанОрг())

        'If Me.Прием_Load = vbTrue Then Form1.Load = False
        Dim bg As String = "Нет"
        Dim StrSql As String
        StrSql = "SELECT ФИОСборное, КодСотрудники FROM Сотрудники INNER Join КарточкаСотрудника On Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр
WHERE Сотрудники.НазвОрганиз='" & Org & "' AND Сотрудники.НаличеДогПодряда='" & bg & "' And КарточкаСотрудника.ДатаУвольнения Is Null ORDER BY ФИОСборное"
        Dim ds As DataTable = Selects(StrSql)

        Refreshgrid()

        Me.MaskedTextBox2.Text = DateTime.Now.ToString("dd.MM.yyyy")
        Me.MaskedTextBox1.Text = DateTime.Now.ToString("dd.MM.yyyy")


        Dim _list As List(Of String) = listFluentFTP(ComboBox1.Text & "/Приказ/")
        'Dim var = From x In dtPutiDokumentovAll.Rows Where x.item("Предприятие") = Org And x.item("ИмяФайла") = "* Приказ *" Select x
        'Dim var2 = From r In var Select r.item("ИмяФайла") = "* Приказ *"



        ComboBox2.Items.Clear()
        For x As Integer = 0 To _list.Count - 1
            ComboBox2.AutoCompleteCustomSource.Add(_list.Item(x).ToString)
            ComboBox2.Items.Add(_list.Item(x).ToString)
        Next





        'Dim Folders() As String
        'Try
        '    Folders = IO.Directory.GetDirectories(OnePath & ComboBox1.Text & "\Приказ", "*", IO.SearchOption.TopDirectoryOnly)
        'Catch ex As Exception

        'End Try

        'Dim gth4 As String
        'For n As Integer = 0 To Folders.Length - 1
        '    gth4 = ""
        '    gth4 = IO.Path.GetFileName(Folders(n))
        '    Folders(n) = gth4
        '    'TextBox44.Text &= gth + vbCrLf
        'Next

        TextBox5.Text = ""

        ComboBox4.Items.Clear()
        ComboBox4.Text = ""
        'ComboBox2.Items.AddRange(Folders)

        ComboBox5.AutoCompleteCustomSource.Clear()
        ComboBox5.Items.Clear()
        For Each r As DataRow In ds.Rows
            ComboBox5.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            ComboBox5.Items.Add(r(0).ToString)
        Next

        ComboBox6.Items.Clear()
        For Each r As DataRow In ds.Rows
            ComboBox6.Items.Add(r(1).ToString)
        Next
        ComboBox5.Text = ""
        TextBox5.Text = ""
    End Sub
    Private Sub СборДанОрг()
        'If Not ds4 Is Nothing Then
        '    ds4.Clear()
        'ElseIf Not ds6 Is Nothing Then
        '    ds6.Clear()
        'End If
        'выборка данных организации
        ds4 = Nothing
        ds4 = dtClientAll.Select("НазвОрг='" & Org & "'")

        '        Dim StrSql4 As String = "SELECT Клиент.ФормаСобств, Клиент.УНП, Клиент.ЮрАдрес, Клиент.КонтТелефон, 
        'Клиент.ЭлАдрес, Клиент.Банк, Клиент.БИКБанка, Клиент.АдресБанка, Клиент.РасчСчетРубли, Клиент.ДолжнРуководителя, 
        'Клиент.ФИОРуководителя, Клиент.НазвОрг, Клиент.РукИП
        'FROM Клиент
        'WHERE Клиент.НазвОрг='" & Org & "'"
        '        'ds4 = Selects(StrSql4)
        ds6 = Nothing
        ds6 = dtformft.Select("ПолноеНазвание='" & ds4(0).Item("ФормаСобств") & "'")
        'выборка короткого названия формы собственности
        'Dim StrSql6 As String = "SELECT Сокращенное FROM ФормаСобств WHERE ПолноеНазвание='" & ds4(0).Item("ФормаСобств") & "'"
        'ds6 = Selects(StrSql6)
    End Sub
    Private Sub ДокиУволПриказ(ByVal inp As String)

        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document
        'Dim oWordPara As Microsoft.Office.Interop.Word.Paragraph

        oWord = CreateObject("Word.Application")
        oWord.Visible = False


        'Try
        '    IO.File.Copy(OnePath & "\ОБЩДОКИ\General\PrikazNaUvolnenie.doc", "C:\Users\Public\Documents\PrikazNaUvolnenie.doc")
        'Catch ex As Exception
        '    If "PrikazNaUvolnenie.doc" <> "" Then IO.File.Delete("C:\Users\Public\Documents\PrikazNaUvolnenie.doc")
        '    IO.File.Copy(OnePath & "\ОБЩДОКИ\General\PrikazNaUvolnenie.doc", "C:\Users\Public\Documents\PrikazNaUvolnenie.doc")
        'End Try
        ВыгрузкаФайловНаЛокалыныйКомп(FTPStringAllDOC & "PrikazNaUvolnenie.doc", firthtPath & "\PrikazNaUvolnenie.doc")

        oWordDoc = oWord.Documents.Add(firthtPath & "\PrikazNaUvolnenie.doc")
        'Dim d As String = Заявление(6)


        'MsgBox(ДолжСОконч)
        With oWordDoc.Bookmarks
            .Item("Увольн1").Range.Text = arrtmask("MaskedTextBox1") & "г."
            .Item("Увольн2").Range.Text = arrtbox("TextBox4")
            .Item("Увольн5").Range.Text = dc9.Rows(0).Item(0).ToString
            If dc9.Rows(0).Item(6).ToString <> "" And Not dc9.Rows(0).Item(6).ToString = "-" Then
                .Item("Увольн6").Range.Text = LCase(ДобОконч(dc9.Rows(0).Item(1))) & " " & разрядстрока(CType(dc9.Rows(0).Item(6).ToString, Integer))
            Else
                .Item("Увольн6").Range.Text = LCase(ДобОконч(dc9.Rows(0).Item(1)))
            End If

            .Item("Увольн7").Range.Text = arrtmask("MaskedTextBox2") & "г." 'дата увольнения
            .Item("Увольн8").Range.Text = ФИОКорРук(inp, False) 'фио в вин падеже

            If ds4(0).Item("ФормаСобств").ToString = "Индивидуальный предприниматель" Then
                .Item("Увольн9").Range.Text = ds4(0).Item("ДолжнРуководителя") 'должность
                .Item("Увольн10").Range.Text = ""
            Else
                .Item("Увольн9").Range.Text = ds4(0).Item("ДолжнРуководителя") & " " & ds6(0).Item("Сокращенное") 'должность и кор форм.собств
                .Item("Увольн10").Range.Text = "«" & arrtcom("ComboBox1") & "»"
            End If

            .Item("Увольн11").Range.Text = ФИОКорРук(ds4(0).Item("ФИОРуководителя"), ds4(0).Item("РукИП")) 'функция по сокращению ФИО руководителя
            .Item("Увольн12").Range.Text = dc9.Rows(0).Item(2) & " " & UCase(Strings.Left(dc9.Rows(0).Item(3), 1)) & "." & UCase(Strings.Left(dc9.Rows(0).Item(4), 1)) & "."
            .Item("Увольн13").Range.Text = dc9.Rows(0).Item(2) & " " & UCase(Strings.Left(dc9.Rows(0).Item(3), 1)) & "." & UCase(Strings.Left(dc9.Rows(0).Item(4), 1)) & "."
            .Item("Увольн14").Range.Text = ds4(0).Item("ФормаСобств") ' форма собственности

            If ds4(0).Item("ФормаСобств").ToString = "Индивидуальный предприниматель" Then
                .Item("Увольн15").Range.Text = arrtcom("ComboBox1")
            Else
                .Item("Увольн15").Range.Text = "«" & arrtcom("ComboBox1") & "»"
            End If

            .Item("Увольн16").Range.Text = ds4(0).Item("ЮрАдрес") 'адрес юр
            .Item("Увольн17").Range.Text = ds4(0).Item("УНП") ' унп
            .Item("Увольн18").Range.Text = ds4(0).Item("РасчСчетРубли") ' р/с
            .Item("Увольн19").Range.Text = ds4(0).Item("Банк") ' банк
            .Item("Увольн20").Range.Text = ds4(0).Item("БИКБанка") ' бик банк
            .Item("Увольн21").Range.Text = ds4(0).Item("АдресБанка") ' адрес банка
            .Item("Увольн22").Range.Text = ds4(0).Item("ЭлАдрес") ' эл.адрес
            .Item("Увольн23").Range.Text = ds4(0).Item("КонтТелефон") ' телефон
            .Item("Увольн24").Range.Text = ОсновУвольн(arrtcom("ComboBox3"))
            .Item("Увольн25").Range.Text = Now.Year.ToString & "г."

        End With


        Dim dirstring As String = arrtcom("ComboBox1") & "/Приказ/" & Now.Year & "/" 'место сохранения файла
        dirstring = СозданиепапкиНаСервере(dirstring) 'полный путь на сервер(кроме имени и разрешения файла)


        Dim put, Name As String
        Name = arrtbox("TextBox4") & " " & dc9.Rows(0).Item(2) & " уволен " & arrtmask("MaskedTextBox2") & ".doc"
        put = PathVremyanka & Name 'место в корне программы


        Parallel.Invoke(Sub() ЗагрВБазуПутиДоков2(idsotrudnika, dirstring, Name, "Приказ-Увольнение", arrtcom("ComboBox1"))) 'заполняем данные путей и назв файла


        oWordDoc.SaveAs2(put,,,,,, False)
        dirstring += Name

        oWordDoc.Close(True)
        oWord.Quit(True)

        ПрикУвольнFTP.AddRange(New String() {dirstring, Name})

        ЗагрНаСерверИУдаление(put, dirstring, put)

        ВременнаяПапкаУдалениеФайла(firthtPath & "\PrikazNaUvolnenie.doc")


    End Sub
    Private Sub УдалениеСтарыхФайловВПапкеРик(ByVal d As String)
        If IO.File.Exists(d) Then
            IO.File.Delete(d)
        End If
    End Sub

    Private Sub Увольн()
        Dim inp As String
        If MessageBox.Show(ComboBox5.Text & " будет уволен.", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button1) = DialogResult.No Then Exit Sub

        IDСотр = CType(Label10.Text, Integer)
        'Dim strsql49 As String = "SELECT ФамилияДляЗаявления, ИмяДляЗаявления, ОтчествоДляЗаявления FROM Сотрудники WHERE КодСотрудники=" & IDСотр & ""
        'Dim dte As DataTable = Selects(strsql49)

        Dim dte() = dtSotrudnikiAll.Select("КодСотрудники=" & IDСотр & "")

        If dte(0).Item("ФамилияДляЗаявления").ToString <> "" Then
            inp = dte(0).Item("ФамилияДляЗаявления").ToString & " " & dte(0).Item("ИмяДляЗаявления").ToString & " " & dte(0).Item("ОтчествоДляЗаявления").ToString
        Else
            inp = InputBox("Введите ФИО сотрудника " & ComboBox5.Text & vbCrLf & " в Винительном падеже 'Заявление от Кого?'", Рик, ComboBox5.Text)
        End If

        Do Until inp <> ""
            MessageBox.Show("Повторите ввод данных!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Error)
            inp = InputBox("Введите ФИО сотрудника " & ComboBox5.Text & " в Винительном падеже 'Заявление от Кого?'", Рик, ComboBox5.Text)
        Loop
        Me.Cursor = Cursors.WaitCursor
        'ФамСотрРодПад = combobox5.Text
        Сотруд = ComboBox5.Text



        Dim list As New Dictionary(Of String, Object)
        list.Add("@КодСотрудники", IDСотр)

        If inp <> "" Then 'добавляем в базу фамилию в род падеже
            Updates(stroka:= "UPDATE Сотрудники SET ФамилияДляУвольнения= '" & inp & "'
            Where КодСотрудники =@КодСотрудники", list, "Сотрудники")

        End If
        'Dim DateEx As String = Format(CDate(MaskedTextBox2.Text), "MM\/dd\/yyyy")
        'Dim DateПр As String = Format(CDate(MaskedTextBox1.Text), "MM\/dd\/yyyy")

        Dim DateEx As String = Replace(Format(CDate(MaskedTextBox2.Text), "yyyy\/MM\/dd"), "/", "")
        Dim DateПр As String = Replace(Format(CDate(MaskedTextBox1.Text), "yyyy\/MM\/dd"), "/", "")

        'обновляем номер приказа, дата приказа и дата увольнения
        Updates(stroka:= "UPDATE КарточкаСотрудника
SET ДатаУвольнения= '" & DateEx & "' , ПриказОбУвольн='" & TextBox4.Text & "', ДатаПриказаОбУвольн= '" & DateПр & "'
Where КарточкаСотрудника.IDСотр =@КодСотрудники", list, "КарточкаСотрудника")


        Dim pu As New Thread(AddressOf ДокиУволПриказ) 'доки приказ
        pu.IsBackground = True
        pu.Start(inp)


        Статистика1(ComboBox5.Text, "Увольнение сотрудника", ComboBox1.Text)

        ВносИзмен2()

        'Dim pt As New Thread(AddressOf ДокиЗаяв) 'доки заявление
        'pt.IsBackground = True
        'pt.Start()

        ДокиЗаяв()
        Me.Cursor = Cursors.Default

        If MessageBox.Show("Распечатать Документы?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
            'pt.Join()
            pu.Join()
            Dim mass As New ArrayList
            mass.Add(ПрикУвольнFTP)
            mass.Add(СохрЗакFTP)

            If ComboBox3.Text = "По истечению срока контракта" Then

                ПечатьДоковFTP(mass(0))
            Else
                ПечатьДоковFTP(mass)
            End If

            'Dim print As New Thread(AddressOf ПечатьДоков)
            'print.IsBackground = True
            'print.Start(mass)

        End If


        Refreshgrid()
        ComboBox5.Text = ""
        TextBox5.Text = ""
        MaskedTextBox1.Text = Now.Date
        TextBox4.Text = ""
        MaskedTextBox2.Text = Now.Date
        'TextBox7.Text = ""
        ComboBox3.Text = ""
        ComboBox2.Text = ""
        ComboBox4.Text = ""

        Dim bg As String = "Нет"
        Dim StrSql As String
        StrSql = "SELECT ФИОСборное FROM Сотрудники INNER Join КарточкаСотрудника On Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр
WHERE Сотрудники.НазвОрганиз='" & ComboBox1.Text & "' AND Сотрудники.НаличеДогПодряда='" & bg & "' And КарточкаСотрудника.ДатаУвольнения Is Null ORDER BY ФИОСборное"
        Dim dfs As DataTable = Selects(StrSql)

        ComboBox5.AutoCompleteCustomSource.Clear()
        ComboBox5.Items.Clear()
        For Each r As DataRow In dfs.Rows
            ComboBox5.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            ComboBox5.Items.Add(r(0).ToString)
        Next
        ComboBox5.Text = ""




    End Sub
    Private Sub ДокиЗаяв()

        If arrtcom("ComboBox3") = "По истечению срока контракта" Then Exit Sub

        'Dim strsql As String = "SELECT * FROM Клиент WHERE НазвОрг='" & ComboBox1.Text & "'"
        'Dim ds As DataTable = Selects(strsql)

        'Dim strsql1 As String = "SELECT * FROM Сотрудники WHERE КодСотрудники=" & ID & ""
        'Dim ds1 As DataTable = Selects(strsql1)


        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document
        'Dim oWordPara As Microsoft.Office.Interop.Word.Paragraph

        oWord = CreateObject("Word.Application")
        oWord.Visible = False

        'Try
        '    IO.File.Copy(OnePath & "\ОБЩДОКИ\General\ZayavlenieUvolnenie.doc", "C:\Users\Public\Documents\Рик\ZayavlenieUvolnenie.doc")
        'Catch ex As Exception
        '    If ex.Message.Contains("уже существует") Then
        '        IO.File.Delete("C:\Users\Public\Documents\Рик\ZayavlenieUvolnenie.doc")
        '        IO.File.Copy(OnePath & "\ОБЩДОКИ\General\ZayavlenieUvolnenie.doc", "C:\Users\Public\Documents\Рик\ZayavlenieUvolnenie.doc")
        '    Else
        '        KillProc()
        '    End If
        'End Try
        ВыгрузкаФайловНаЛокалыныйКомп(FTPStringAllDOC & "ZayavlenieUvolnenie.doc", firthtPath & "\ZayavlenieUvolnenie.doc")

        oWordDoc = oWord.Documents.Add(firthtPath & "\ZayavlenieUvolnenie.doc")

        Dim dsДанКл() = dtClientAll.Select("НазвОрг= '" & Org & "'") 'выбрали данные по организации

        With oWordDoc.Bookmarks
            .Item("ЗаяУв1").Range.Text = arrtmask("MaskedTextBox1")
            If dsДанКл(0).Item(1).ToString = "Индивидуальный предприниматель" Then
                .Item("ЗаяУв2").Range.Text = ДолжРодПадежФункц(dsДанКл(0).Item(1).ToString)
            Else
                .Item("ЗаяУв2").Range.Text = ДолжРодПадежФункц(dsДанКл(0).Item(18).ToString) & " " & ФормСобствКор(dsДанКл(0).Item(1).ToString) & " «" & arrtcom("ComboBox1") & "» "
            End If
            If dsДанКл(0).Item(31) = True Then
                .Item("ЗаяУв3").Range.Text = ФИОКорРук(dsДанКл(0).Item(19).ToString, True)
            Else
                .Item("ЗаяУв3").Range.Text = ФИОКорРук(dsДанКл(0).Item(19).ToString, False)
            End If

            .Item("ЗаяУв4").Range.Text = Trim(dsДанСотр.Rows(0).Item(23).ToString) & " " & Trim(dsДанСотр.Rows(0).Item(24).ToString) & " " & Trim(dsДанСотр.Rows(0).Item(25).ToString)
            .Item("ЗаяУв5").Range.Text = dsДанСотр.Rows(0).Item(16).ToString
            .Item("ЗаяУв6").Range.Text = dsДанСотр.Rows(0).Item(19).ToString
            .Item("ЗаяУв7").Range.Text = arrtmask("MaskedTextBox2")
            .Item("ЗаяУв8").Range.Text = arrtmask("MaskedTextBox1")
            .Item("ЗаяУв9").Range.Text = ФИОКорРук(dsДанСотр.Rows(0).Item(5).ToString, False)

        End With

        Dim dirstring As String = arrtcom("ComboBox1") & "/Заявление/" & Now.Year & "/" 'место сохранения файла
        dirstring = СозданиепапкиНаСервере(dirstring) 'полный путь на сервер(кроме имени и разрешения файла)


        Dim put, Name As String
        Name = ФИОКорРук(dsДанСотр.Rows(0).Item(5).ToString, False) & " (Заявление увольнение)" & ".doc"
        put = PathVremyanka & Name 'место в корне программы

        Parallel.Invoke(Sub() ЗагрВБазуПутиДоков2(idsotrudnika, dirstring, Name, "Заявление-Увольнение", arrtcom("ComboBox1"))) 'заполняем данные путей и назв файла

        oWordDoc.SaveAs2(put,,,,,, False)
        dirstring += Name

        oWordDoc.Close(True)
        oWord.Quit(True)

        СохрЗакFTP.AddRange(New String() {dirstring, Name})

        ЗагрНаСерверИУдаление(put, dirstring, put)

        ВременнаяПапкаУдалениеФайла(firthtPath & "\ZayavlenieUvolnenie.doc")


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


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If TextBox9.Text <> "" Then
            TextBox4.Text = TextBox4.Text & " - " & TextBox8.Text & " - " & TextBox9.Text
        Else
            TextBox4.Text = TextBox4.Text & " - " & TextBox8.Text
        End If

        If arrtbox.Any Then
            arrtbox.Clear()
        End If

        If arrtmask.Any Then
            arrtmask.Clear()
        End If

        If arrtcom.Any Then
            arrtcom.Clear()
        End If
        idsotrudnika = Nothing
        idsotrudnika = CType(Label10.Text, Integer)
        ЗаполнМассВнеТабах()
        Увольн()
    End Sub
    Private Sub ВносИзмен2()
        Dim StrSql As String
        Dim f, t, y, u As String
        f = Me.MaskedTextBox2.Text
        t = Me.TextBox4.Text
        y = Me.MaskedTextBox1.Text
        u = Me.ComboBox3.Text

        Dim list As New Dictionary(Of String, Object)
        list.Add("@IDСотр", IDСотр)

        Updates(stroka:="UPDATE КарточкаСотрудника SET ДатаУвольнения = '" & f & "', ПриказОбУвольн = '" & t & "', ДатаПриказаОбУвольн = '" & y & "',
ОснованиеУвольн = '" & u & "', НеПродлениеКонтр='True' WHERE IDСотр =@IDСотр", list, "КарточкаСотрудника")
        MsgBox("Сотрудник " & Сотруд & " был уволен " & f,, Рик)
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        ComboBox4.Items.Clear()
        ComboBox4.Text = ""

        Dim _list As List(Of String) = listFluentFTP(ComboBox1.Text & "/Приказ/" & ComboBox2.Text & "/")


        ComboBox4.Items.Clear()
        For x As Integer = 0 To _list.Count - 1
            ComboBox4.AutoCompleteCustomSource.Add(_list.Item(x).ToString)
            ComboBox4.Items.Add(_list.Item(x).ToString)
        Next

        'Files3 = (IO.Directory.GetFiles(OnePath & ComboBox1.Text & "\Приказ\" & ComboBox2.Text, "*.doc", IO.SearchOption.TopDirectoryOnly))

        'Dim gth As String
        'For n As Integer = 0 To Files3.Length - 1
        '    gth = ""
        '    gth = IO.Path.GetFileName(Files3(n))
        '    Files3(n) = gth
        '    'TextBox44.Text &= gth + vbCrLf
        'Next

        'ComboBox4.Items.AddRange(Files3)


    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            Refreshgrid()
        Else
            Refreshgrid()
        End If
    End Sub

    Private Sub СборданСотр()
        If Not dc9 Is Nothing Then
            dc9.Clear()
        End If
        ' выборка данных для сотрудника
        Dim StrSql1 As String = "SELECT Сотрудники.ФИОРодПод, Штатное.Должность, Сотрудники.Фамилия, Сотрудники.Имя,
Сотрудники.Отчество, Сотрудники.ФамилияДляУвольнения, Штатное.Разряд
        FROM Сотрудники INNER JOIN Штатное ON Сотрудники.КодСотрудники = Штатное.ИДСотр
        WHERE Сотрудники.КодСотрудники = " & IDСотр & ""
        dc9 = Selects(StrSql1)
        Try
            Dim fd As String = dc9.Rows(0).Item(0).ToString
        Catch ex As Exception
            MessageBox.Show("Поверьте и исправьте, есть ли у сотрудника должность и отдел. Или возможно он принят по Договор-Подряду", Рик)
            Me.Cursor = Cursors.Default

        End Try
    End Sub
    Private Sub dsданСотруд(ByVal ID As Integer)
        If Not dsДанСотр Is Nothing Then
            dsДанСотр.Clear()
        End If
        Dim strsql1 As String = "SELECT * FROM Сотрудники WHERE КодСотрудники=" & ID & ""
        dsДанСотр = Selects(strsql1)
    End Sub
    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged

        Grid1.ClearSelection()
        Dim ind As Boolean = False
        For Each row As DataGridViewRow In Grid1.Rows
            For Each cell As DataGridViewCell In row.Cells
                If (cell.FormattedValue).Contains(Me.ComboBox5.Text) Then
                    row.Selected = True
                    Grid1.FirstDisplayedScrollingRowIndex = row.Index
                End If
            Next
            'If ind = False Then row.Visible = False
            'ind = False
        Next

        'Dim strsql As String = "SELECT ИДНомер FROM Сотрудники WHERE НазвОрганиз='" & ComboBox1.Text & "' AND ФИОСборное='" & ComboBox5.Text & "'"
        'Dim dr As DataTable = Selects(strsql)

        Dim dr = dtSotrudnikiAll.Select("НазвОрганиз='" & ComboBox1.Text & "' AND ФИОСборное='" & ComboBox5.Text & "'")

        TextBox5.Text = ""
        TextBox5.Text = dr(0).Item("ИДНомер").ToString

        Label10.Text = ComboBox6.Items.Item(ComboBox5.SelectedIndex)

        IDСотр = Nothing
        IDСотр = ComboBox6.Items.Item(ComboBox5.SelectedIndex)

        Dim g As New Thread(AddressOf СборданСотр)
        g.IsBackground = True
        g.Start()

        Dim g1 As New Thread(AddressOf dsданСотруд) 'данные по сотруднику для заявления
        g1.IsBackground = True
        g1.Start(IDСотр)


    End Sub

    Private Sub Refreshgrid()

        Dim StrSql1, StrSql2 As String
        If Not tbl Is Nothing Then
            tbl.Clear()
        ElseIf Not tbl2 Is Nothing Then
            tbl2.Clear()
        End If


        If CheckBox1.Checked = False Then

            StrSql1 = "SELECT Сотрудники.НазвОрганиз as [Организация], Сотрудники.ФИОСборное as [Сотрудник], Штатное.Должность, 
        КарточкаСотрудника.ДатаПриема as [Принят на работу], ДогСотрудн.СрокОкончКонтр as [Окончание контракта], КарточкаСотрудника.СрокПродлКонтракта as [Продление контракта, лет],
        КарточкаСотрудника.ПродлКонтрС as [Контракт продлен, с], КарточкаСотрудника.ПродлКонтрПо as [Контракт продлен, по], Сотрудники.КодСотрудники 
        FROM ((Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр) INNER JOIN Штатное ON Сотрудники.КодСотрудники = Штатное.ИДСотр) INNER JOIN ДогСотрудн ON Сотрудники.КодСотрудники = ДогСотрудн.IDСотр
        WHERE Сотрудники.НазвОрганиз='" & Org & "' AND КарточкаСотрудника.ДатаУвольнения Is Null AND Сотрудники.НаличеДогПодряда='Нет' ORDER BY Сотрудники.ФИОСборное " 'КарточкаСотрудника.НеПродлениеКонтр=No (заменил на - КарточкаСотрудника.ДатаУвольнения Is Null)
            tbl = Selects(StrSql1)
            Grid1.DataSource = tbl
            Grid1.Columns(8).Visible = False
            Grid1.Columns(0).Visible = False
            Grid1.Columns(1).Width = 300
            Grid1.Columns(2).Width = 300
            GridView(Grid1)

        Else
            'AND КарточкаСотрудника.ПриказОбУвольн='" & ged & "'
            StrSql2 = "SELECT Сотрудники.НазвОрганиз as [Организация], Сотрудники.ФИОСборное as [Сотрудник], Штатное.Должность, 
        КарточкаСотрудника.ДатаПриема as [Принят на работу], ДогСотрудн.СрокОкончКонтр as [Окончание контракта], КарточкаСотрудника.ДатаУведомлПродКонтр as [Дата уведомления о не продлении контракта],
        КарточкаСотрудника.НомерУведомлПродКонтр as [Номер уведомления о непродлении контракта], Сотрудники.КодСотрудники 
        FROM ((Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр) INNER JOIN Штатное ON Сотрудники.КодСотрудники = Штатное.ИДСотр) INNER JOIN ДогСотрудн ON Сотрудники.КодСотрудники = ДогСотрудн.IDСотр
        WHERE Сотрудники.НазвОрганиз='" & Org & "' AND КарточкаСотрудника.НеПродлениеКонтр='True' AND КарточкаСотрудника.ПриказОбУвольн Is Null ORDER BY Сотрудники.ФИОСборное "

            tbl2 = Selects(StrSql2)

            Grid1.DataSource = tbl2
            Grid1.Columns(7).Visible = False
            Grid1.Columns(0).Visible = False
            GridView(Grid1)
        End If



    End Sub

    Private Sub MaskedTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            MaskedTextBox2.Focus()
        End If
    End Sub

    Private Sub MaskedTextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            TextBox4.Focus()
        End If
    End Sub

    Private Sub TextBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox9.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            ComboBox3.Focus()
        End If
    End Sub

    Private Sub ComboBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Button2.Focus()
        End If
    End Sub

    Private Sub Grid1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellClick
        'If e.RowIndex = -1 Then Exit Sub

        'ComboBox5.Text = Grid1.CurrentRow.Cells(1).Value.ToString

        'Dim StrSql As String
        'StrSql = "SELECT ИДНомер,КодСотрудники FROM Сотрудники WHERE НазвОрганиз='" & Org & "' AND ФИОСборное='" & ComboBox5.Text & "'"
        'Dim c As New OleDbCommand With {
        '    .Connection = conn,
        '    .CommandText = StrSql
        '}
        'Dim ds As New DataSet
        'Dim da As New OleDbDataAdapter(c)
        'da.Fill(ds, "КонтРед")

        'TextBox5.Text = ds.Tables("КонтРед").Rows(0).Item(0).ToString
        'IDСотр = ds.Tables("КонтРед").Rows(0).Item(1)

        ComboBox5.SelectedItem = Grid1.CurrentRow.Cells(1).Value



    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True

            Dim i As Integer
            Dim pl As String
            If TextBox4.Text <> "" Then
                Try
                    i = CInt(TextBox4.Text)
                Catch ex As Exception
                    MessageBox.Show("Это поле только для цифр!", Рик)
                    TextBox4.Text = ""
                    Exit Sub
                End Try

                Select Case i
                    Case < 10
                        pl = Str(i)
                        TextBox4.Text = "00" & i

                    Case 10 To 99
                        pl = Str(i)
                        TextBox4.Text = "0" & i
                End Select

            End If
            TextBox9.Focus()
        End If
    End Sub



    Private Sub Grid1_ColumnHeaderCellChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles Grid1.ColumnHeaderCellChanged

    End Sub

    Private Sub Grid1_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles Grid1.ColumnHeaderMouseClick
        Grid1.Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
        Grid1.Columns(4).SortMode = DataGridViewColumnSortMode.NotSortable
        Grid1.Columns(5).SortMode = DataGridViewColumnSortMode.NotSortable
        Grid1.Columns(6).SortMode = DataGridViewColumnSortMode.NotSortable
        Grid1.Columns(7).SortMode = DataGridViewColumnSortMode.NotSortable
        Grid1.Columns(8).SortMode = DataGridViewColumnSortMode.NotSortable
    End Sub

    Private Sub TextBox4_LostFocus(sender As Object, e As EventArgs) Handles TextBox4.LostFocus
        Dim pl As String
        Dim i As Integer
        If TextBox4.Text <> "" Then
            Try
                i = CInt(TextBox4.Text)
            Catch ex As Exception
                MessageBox.Show("Это поле только для цифр!", Рик)
                TextBox4.Text = ""
                Exit Sub
            End Try

            Select Case i
                Case < 10
                    pl = Str(i)
                    TextBox4.Text = "00" & i

                Case 10 To 99
                    pl = Str(i)
                    TextBox4.Text = "0" & i
            End Select
        End If
    End Sub
End Class