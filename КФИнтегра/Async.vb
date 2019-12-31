Imports System.Data.OleDb
Imports MySql.Data.MySqlClient
Imports System.Management
Imports System.IO
Imports System.Security.Policy
Imports System.Data.SqlClient

Module Async
    Public Готовый As New List(Of String)
    Public DirAll As New Dictionary(Of String, String)()

    Public dtformft, dtClientAll, dtSotrudnikiAll, dtShtatnoeAll, dtPutiDokumentovAll,
        dtShtatnoeOtdelyAll, dtKartochkaSotrudnikaAll, dtObjectObshepitaAll, dtStatistikaAll,
        dtPerevodAll, dtProdlenieKontraktaAll, dtShtatnoeSvodnoeIzmenStavkaAll, dtSostavSemyiAll,
        dtDogovorPadriadaAll, dtShtatnoeSvodnoeAll, dtDogovorSotrudnikAll, dtOtpuskAll, dtOtpuskSotrudnikiAll,
    dtOtpuskSocAll, dtDogPodryadaAktAll, dtOkonchanieAll, dtPodriadaOsobenAll, dtStatnoeRaspisanieAll,
    dtDogPodrAktInoeAll, dtDogPodrDoljnostAll, dtDogPodrRabotyInoeAll As DataTable 'таблицы с полным списком всех полей для выборок
    Public Sub RunMoving()
        dtformft = Selects(StrSql:="SELECT * FROM ФормаСобств")
    End Sub
    Public Async Sub dtformftAsync() 'ФормаСобственности
        Await Task.Run((Sub() RunMoving()))
    End Sub
    Public Sub RunMoving1()
        dtClientAll = Selects(StrSql:="SELECT * FROM Клиент")
    End Sub
    Public Async Sub dtClient() 'Все клиенты
        Await Task.Run((Sub() RunMoving1()))
    End Sub
    Public Function ВыборкаСтрокиИзТаблицы(ByVal переменная As String, ByVal ds As DataTable, ByVal столбец As String) As DataRow()
        Dim r As String = столбец & "='" & переменная & "'"
        Dim f = ds.Select(r)
        Return f
    End Function
    Public Sub RunMoving2()
        dtSotrudnikiAll = Selects(StrSql:="SELECT * FROM Сотрудники")
    End Sub
    Public Async Sub dtSotrudniki() 'Все сотрудники
        Await Task.Run((Sub() RunMoving2()))
    End Sub
    Public Async Sub ВсплывФормапередЗагрузкой() 'асинхронно проверка уведомления о продлении контракта
        Await Task.Run(Sub() ВсплывФормаПриЗагр.ВсплывФорма())
    End Sub
    Public Function ComboSotrudnikiLinqS(ByVal ИмяСтолбца As String, ByVal ИскомоеЗначение As String)
        Dim var = From x In dtSotrudnikiAll.Rows Where x.Item(ИмяСтолбца) = ИскомоеЗначение Select x 'рабочий linq для заполнения комбобоксов
        Return var
    End Function

    Public Sub ЗагрВБазуПутиДоков2(ByVal id As Integer, ByVal Path As String, ByVal Name As String, ByVal Other As String, ByVal Org As String)
        Dim ds = dtPutiDokumentovAll.Select("IDСотрудник=" & id & " and Путь='" & Path & "' and ИмяФайла='" & Name & "' and ДокМесто='" & Other & "'")

        Try
            If ds.Length = 0 Then
                '                Updates(stroka:="INSERT INTO ПутиДокументов(IDСотрудник,Путь,ИмяФайла,ДокМесто,Предприятие) 
                'VALUES(" & id & ",'" & Path & "', '" & Name & "', '" & Other & "','" & Org & "')")
                Dim m As String = Path & Name
                Updates(stroka:="INSERT INTO ПутиДокументов(IDСотрудник,Путь,ИмяФайла,ДокМесто,Предприятие,ПолныйПуть) 
VALUES(" & id & ",'" & Path & "', '" & Name & "', '" & Other & "','" & Org & "','" & m & "')")

                '            Else
                '                Updates(stroka:="UPDATE ПутиДокументов SET Путь='" & Path & "',[ИмяФайла]='" & Name & "', ДокМесто='" & Other & "', Предприятие='" & Org & "'
                'WHERE IDСотрудник=" & id & "") and 
                '                '                Updates(stroka:="UPDATE ПутиДокументов SET Путь='" & Path & "', ИмяФайла='" & Name & "', ДокМесто='" & Other & "', Предприятие='" & Org & "'
                '                'WHERE IDСотрудник=" & id & "")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message & "Строка 54 ASync")
        End Try


    End Sub
    Public Async Sub ЗагрВБазуПутиДоковAsync(ByVal id As Integer, ByVal Path As String, ByVal Name As String, ByVal Other As String, ByVal Org As String) 'Все сотрудники
        Await Task.Run((Sub() ЗагрВБазуПутиДоков2(id, Path, Name, Other, Org)))
    End Sub

    Public Sub RunMoving3()
        dtShtatnoeAll = Selects(StrSql:="SELECT * FROM Штатное")
    End Sub
    Public Async Sub dtShtatnoe() 'Штатное
        Await Task.Run((Sub() RunMoving3()))
    End Sub
    Public Sub RunMoving4()
        dtPutiDokumentovAll = Selects(StrSql:="SELECT * FROM ПутиДокументов")
    End Sub
    Public Async Sub dtPutiDokumentov() 'Все данные из пути документов
        Await Task.Run((Sub() RunMoving4()))
    End Sub

    Public Sub RunMoving5()
        dtShtatnoeOtdelyAll = Selects(StrSql:="SELECT * FROM ШтОтделы")
    End Sub
    Public Async Sub dtShtatnoeOtdely() 'Штатное отделы
        Await Task.Run((Sub() RunMoving5()))
    End Sub
    Public Sub RunMoving6()
        dtKartochkaSotrudnikaAll = Selects(StrSql:="SELECT * FROM КарточкаСотрудника")
    End Sub
    Public Async Sub dtKartochkaSotrudnika() 'Карточка сотрудника
        Await Task.Run((Sub() RunMoving6()))
    End Sub
    Public Sub RunMoving7()
        dtProdlenieKontraktaAll = Selects(StrSql:="SELECT * FROM ПродлКонтракта")
    End Sub
    Public Async Sub dtProdlenieKontrakta() 'ПродлКонтракта
        Await Task.Run((Sub() RunMoving7()))
    End Sub

    Public Sub RunMoving8()
        dtDogovorSotrudnikAll = Selects(StrSql:="SELECT * FROM ДогСотрудн")

    End Sub
    Public Async Sub dtDogovorSotrudnik() 'Договор сотрудник
        Await Task.Run(Sub() RunMoving8())
    End Sub
    Public Async Sub dtDogovorPadriada() 'Договор подряда
        Await Task.Run((Sub() RunMoving9()))
    End Sub
    Public Sub RunMoving9()
        dtDogovorPadriadaAll = Selects(StrSql:="SELECT * FROM ДогПодряда")
    End Sub


    Public Sub RunMoving10()
        dtShtatnoeSvodnoeAll = Selects(StrSql:="SELECT * FROM ШтСвод")
    End Sub
    Public Async Sub dtShtatnoeSvodnoe() 'ШтСводИзмСтавка
        Await Task.Run((Sub() RunMoving10()))
    End Sub
    Public Async Sub dtShtatnoeSvodnoeIzmenStavka() 'Договор сотрудник
        Await Task.Run((Sub() RunMoving11()))
    End Sub
    Public Sub RunMoving11()
        dtShtatnoeSvodnoeIzmenStavkaAll = Selects(StrSql:="SELECT * FROM ШтСводИзмСтавка")
    End Sub

    Public Async Sub dtSostavSemyi() 'Состав семьи
        Await Task.Run((Sub() RunMoving12()))
    End Sub
    Public Sub RunMoving12()
        dtSostavSemyiAll = Selects(StrSql:="SELECT * FROM СоставСемьи")
    End Sub
    Public Async Sub dtStatistika() 'Состав семьи
        Await Task.Run((Sub() RunMoving13()))
    End Sub
    Public Sub RunMoving13()
        dtStatistikaAll = Selects(StrSql:="SELECT * FROM Статистика")
    End Sub
    Public Sub ALLALL()
        dtStatistika()
        dtSostavSemyi()
        dtShtatnoeSvodnoeIzmenStavka()
        dtShtatnoeSvodnoe()
        dtDogovorPadriada()
        dtDogovorSotrudnik()
        dtProdlenieKontrakta()
        dtKartochkaSotrudnika()
        dtShtatnoeOtdely()
        dtPutiDokumentov()
        dtShtatnoe()
        dtSotrudniki()
        dtClient()
        dtformftAsync()
        dtPerevod()
        dtObjectObshepita()
        dtOtpusk()
        dtOtpuskSotrudniki()
        dtOtpuskSoc()
        dtDogPodryadaAkt()
        dtOkonchanie()
        dtPodriadaOsoben()
        dtStatnoeRaspisanie()
        dtDogPodrDoljnost()
        dtDogPodrAktInoe()
        dtDogPodrRabotyInoe()
    End Sub

    Public Async Sub dtPerevod() 'Перевод
        Await Task.Run((Sub() RunMoving14()))
    End Sub
    Public Sub RunMoving14()
        dtPerevodAll = Selects(StrSql:="SELECT * FROM Перевод")
    End Sub

    Public Async Sub dtObjectObshepita() 'Перевод
        Await Task.Run((Sub() RunMoving15()))
    End Sub
    Public Sub RunMoving15()
        dtObjectObshepitaAll = Selects(StrSql:="SELECT * FROM ОбъектОбщепита")
    End Sub

    Public Async Sub dtOtpusk() 'Отпуск
        Await Task.Run((Sub() RunMoving16()))
    End Sub
    Public Sub RunMoving16()
        dtOtpuskAll = Selects(StrSql:="SELECT * FROM Отпуск")
    End Sub
    Public Async Sub dtOtpuskSotrudniki() 'ОтпускСотрудники
        Await Task.Run((Sub() RunMoving17()))
    End Sub
    Public Sub RunMoving17()
        dtOtpuskSotrudnikiAll = Selects(StrSql:="SELECT * FROM ОтпускСотрудники")
    End Sub
    Public Async Sub dtOtpuskSoc() 'Отпуск социальный
        Await Task.Run((Sub() RunMoving18()))
    End Sub
    Public Sub RunMoving18()
        dtOtpuskSocAll = Selects(StrSql:="SELECT * FROM ОтпускСоц")
    End Sub
    Public Async Sub dtDogPodryadaAkt() 'ДогПодрядаАкт
        Await Task.Run((Sub() RunMoving19()))
    End Sub
    Public Sub RunMoving19()
        dtDogPodryadaAktAll = Selects(StrSql:="SELECT * FROM ДогПодрядаАкт")
    End Sub
    Public Async Sub dtOkonchanie() 'Окончание
        Await Task.Run((Sub() RunMoving20()))
    End Sub
    Public Sub RunMoving20()
        dtOkonchanieAll = Selects(StrSql:="SELECT * FROM Окончание")
    End Sub
    Public Async Sub dtPodriadaOsoben() 'ДогПодОсобен
        Await Task.Run((Sub() RunMoving21()))
    End Sub
    Public Sub RunMoving21()
        dtPodriadaOsobenAll = Selects(StrSql:="SELECT * FROM ДогПодОсобен")
    End Sub
    Public Async Sub dtStatnoeRaspisanie() 'ШтатРаспис
        Await Task.Run((Sub() RunMoving22()))
    End Sub
    Public Sub RunMoving22()
        dtStatnoeRaspisanieAll = Selects(StrSql:="SELECT * FROM ШтатРаспис")
    End Sub

    Public Async Sub dtDogPodrDoljnost() 'ДогПодДолжн
        Await Task.Run((Sub() RunMoving23()))
    End Sub
    Public Sub RunMoving23()
        dtDogPodrDoljnostAll = Selects(StrSql:="SELECT * FROM ДогПодДолжн")
    End Sub
    Public Async Sub dtDogPodrAktInoe() 'ДогПодрядаАктИное
        Await Task.Run((Sub() RunMoving24()))
    End Sub
    Public Sub RunMoving24()
        dtDogPodrAktInoeAll = Selects(StrSql:="SELECT * FROM ДогПодрядаАктИное")
    End Sub
    Public Async Sub dtDogPodrRabotyInoe() 'ДогПодрядаРаботыИное
        Await Task.Run((Sub() RunMoving25()))
    End Sub
    Public Sub RunMoving25()
        dtDogPodrRabotyInoeAll = Selects(StrSql:="SELECT * FROM ДогПодрядаРаботыИное")
    End Sub


    Private Sub allFiles() 'выборка всех файлов в папках
        Dim sw As New Stopwatch 'вычисление выполнения метода
        sw.Start()

        DirAll.Clear()

        Dim listPath = listFluentFTP("/")


        Dim listPath2 As New List(Of String)
        Dim listYear As New List(Of String)
        Dim Промеж As New List(Of String)
        Dim listPath2prom As New List(Of String)

        For Each item In listPath
            If Not item = "ALLINALLDATABASE" Then
                listPath2.AddRange(listFluentFTP(item & "/"))
            End If
        Next
        listPath2 = listPath2.Distinct().ToList

        For Each item In listPath
            For Each item1 In listPath2
                If Not item1.Contains("инструкции") Then
                    listYear.AddRange(listFluentFTP(item & "/" & item1 & "/"))
                End If
            Next
        Next
        listYear = listYear.Distinct().ToList

        For Each item1 In listPath
            For Each item2 In listPath2
                If item1.Contains("инструкции") Then
                    Промеж = listFluentFTP(item1 & "/" & item2 & "/")
                    For Each item4 In Промеж
                        Try
                            DirAll.Add(item4, item1 & "/" & item2 & "/")
                        Catch ex As Exception
                            MessageBox.Show(ex.Message)
                        End Try

                    Next
                Else
                    For Each item3 In listYear
                        Промеж = listFluentFTP(item1 & "/" & item2 & "/" & item3 & "/")
                        For Each item4 In Промеж
                            Try
                                DirAll.Add(item4, item1 & "/" & item2 & "/" & item3 & "/")
                            Catch ex As Exception
                                MessageBox.Show(ex.Message)
                            End Try

                        Next
                    Next
                End If
            Next
        Next

        'Dim f = From x In DirAll.Keys Where x.Contains("инструкции") Select x



        sw.Stop()
        MessageBox.Show((sw.ElapsedMilliseconds / 100.0).ToString())
    End Sub
    Private Sub allFiles2() 'выборка всех файлов в папках

        'Dim sw As New Stopwatch 'вычисление выполнения метода время
        'sw.Start()

        DirAll.Clear()

        listFluentFTP2("/")

        'sw.Stop()
        'MessageBox.Show((sw.ElapsedMilliseconds / 100.0).ToString())
    End Sub

    Public Sub Начало(ByVal Файл As String)
        If Файл.Contains("\") Then
            Файл = Strings.Right(Файл, Файл.Length - 1)
            ВыгрузкаФайловНаЛокалыныйКомп(FTPStringAllDOC & Файл, firthtPath & "\" & Файл)
        Else
            ВыгрузкаФайловНаЛокалыныйКомп(FTPStringAllDOC & Файл, firthtPath & "\" & Файл)
        End If

    End Sub

    Public Sub Конец(ByVal МестоСохр As String, ByVal ИмяФайла As String, ByVal ID As Integer,
                     ByVal Организация As String, ByVal УдалениеНачало As String, ByVal Пояснение As String)

        If Not Strings.Right(МестоСохр, 1) = "/" Then
            МестоСохр &= "/"
        Else


        End If


        МестоСохр = СозданиепапкиНаСервере(МестоСохр) 'полный путь на сервер(кроме имени и разрешения файла)


        Dim put, Name As String
        'Name = ДПодНом & " " & arrtbox("TextBox1") & " от " & arrtmask("MaskedTextBox6") & "(Договор подряда)" & ".doc"
        put = PathVremyanka & ИмяФайла 'место в корне программы

        ВыборкаИзагрНаСервер2(МестоСохр, ИмяФайла, Пояснение, ID, Организация)
        МестоСохр += ИмяФайла

        'СохрFTP.AddRange(New String() {МестоСохр, ИмяФайла})



        ЗагрНаСерверИУдаление(put, МестоСохр, put)

        ВременнаяПапкаУдалениеФайла(firthtPath & УдалениеНачало)

        'Образец заполнения!!!

        'Dim name As String = "ДопСогл " & ПровДанн & "_" & ВремяС & " " & Ds.Rows(0).Item(1).ToString & " " & КорИмя & " " & Коротч & "(Доп.Продл.Контр)" & ".doc"

        'ДопСоглFTP.AddRange(New String() {ComboBox1.Text & "\Дополнительное солгашение\" & Now.Year, Name})

        'oWordDoc2.SaveAs2(PathVremyanka & Name,,,,,, False)
        'oWordDoc2.Close(True)
        'oWord2.Quit(True)
        'Конец(ComboBox1.Text & "\Дополнительное солгашение\" & Now.Year, Name, Mass, ComboBox1.Text, "\DopSoglashenie.doc", "ДопCолгашПродлКонтракта")


        'Печать образец
        ''massFTP(это arrya list)'пояснение
        'massFTP.Add(УведомлFTP)
        ' ПечатьДоковFTP(massFTP)

        'Dim list As New Dictionary(Of String, Object)()        '
        'List.Add("@УведомлПродл", УведомлПродл)



    End Sub
    Public Sub Конец(ByVal МестоСохр As String, ByVal ИмяФайла As String, ByVal Организация As String, ByVal УдалениеНачало As String, ByVal Пояснение As String)

        If Not Strings.Right(МестоСохр, 1) = "/" Then
            МестоСохр &= "/"
        Else
        End If


        МестоСохр = СозданиепапкиНаСервере(МестоСохр) 'полный путь на сервер(кроме имени и разрешения файла)

        Dim put As String

        put = PathVremyanka & ИмяФайла 'место в корне программы

        Dim ds = dtPutiDokumentovAll.Select("Путь='" & МестоСохр & "' and ИмяФайла='" & ИмяФайла & "' and ДокМесто='" & Пояснение & "'")
        Dim list As New Dictionary(Of String, Object)
        'list.Add("@Код", ds(0).Item("Код"))



        Dim m As String = МестоСохр & ИмяФайла
        If ds.Length = 0 Then
            '                Updates(stroka:="INSERT INTO ПутиДокументов(IDСотрудник,Путь,ИмяФайла,ДокМесто,Предприятие) 
            'VALUES(" & id & ",'" & Path & "', '" & Name & "', '" & Other & "','" & Org & "')")

            Updates(stroka:="INSERT INTO ПутиДокументов(Путь,ИмяФайла,ДокМесто,Предприятие,ПолныйПуть) 
VALUES('" & МестоСохр & "', '" & ИмяФайла & "', '" & Пояснение & "','" & Организация & "','" & m & "')", list, "ПутиДокументов")
        Else
            list.Add("@Код", ds(0).Item("Код"))
            Updates(stroka:="UPDATE ПутиДокументов SET Путь='" & МестоСохр & "', ИмяФайла='" & ИмяФайла & "', ДокМесто='" & Пояснение & "',
Предприятие='" & Организация & "',ПолныйПуть='" & m & "' WHERE Код=@Код", list, "ПутиДокументов")

        End If


        'ВыборкаИзагрНаСервер2(МестоСохр, ИмяФайла, Пояснение, Организация)
        МестоСохр += ИмяФайла

        'СохрFTP.AddRange(New String() {МестоСохр, ИмяФайла})



        ЗагрузкаФайловНаСервер(put, МестоСохр)
        ВременнаяПапкаУдалениеФайла(put)





    End Sub

    Public Sub ВыборкаИзагрНаСервер2(ByVal _МестоСохр As String, ByVal _ИмяФайла As String,
                                     ByVal _Форма As String, ByVal ID As Integer, ByVal Организация As String)

        ЗагрВБазуПутиДоков2(ID, _МестоСохр, _ИмяФайла, _Форма, Организация) 'заполняем данные путей и назв файла

    End Sub







End Module
