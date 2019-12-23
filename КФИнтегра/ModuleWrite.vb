Imports System.Data.OleDb
Module ModuleWrite
    Dim n As Integer
    Public КонтрПровИндивид() As String
    Dim WithEvents Proc As Process
    Public vstavContr, vstavContr1 As String

    Public firthtPath As String 'папка где храняться файлы из General
    Public PathVremyanka As String 'папка времянка в General

    Function Пропись(ByVal число As Integer) As String

        Static triad(4) As Integer, numb1(0 To 19) As String, numb2(0 To 9) As String, numb3(0 To 9) As String
        If число = 0 Then
            Пропись = ""
            Exit Function
        End If
        Dim ss As Decimal = число
        triad(1) = ss@ - Int(ss@ / 1000) * 1000
        ss@ = Int(ss@ / 1000)
        triad(2) = ss@ - Int(ss@ / 1000) * 1000
        ss@ = Int(ss@ / 1000)
        triad(3) = ss@ - Int(ss@ / 1000) * 1000
        ss@ = Int(ss@ / 1000)
        triad(4) = ss@ - Int(ss@ / 1000) * 1000
        ss@ = Int(ss@ / 1000)
        numb1(0) = ""
        numb1(1) = "один "
        numb1(2) = "два "
        numb1(3) = "три "
        numb1(4) = "четыре "
        numb1(5) = "пять "
        numb1(6) = "шесть "
        numb1(7) = "семь "
        numb1(8) = "восемь "
        numb1(9) = "девять "
        numb1(10) = "десять "
        numb1(11) = "одиннадцать "
        numb1(12) = "двенадцать "
        numb1(13) = "тринадцать "
        numb1(14) = "четырнадцать "
        numb1(15) = "пятнадцать "
        numb1(16) = "шестнадцать "
        numb1(17) = "семнадцать "
        numb1(18) = "восемнадцать "
        numb1(19) = "девятнадцать "
        numb2(0) = ""
        numb2(1) = ""
        numb2(2) = "двадцать "
        numb2(3) = "тридцать "
        numb2(4) = "сорок "
        numb2(5) = "пятьдесят "
        numb2(6) = "шестьдесят "
        numb2(7) = "семьдесят "
        numb2(8) = "восемьдесят "
        numb2(9) = "девяносто "
        numb3(0) = ""
        numb3(1) = "сто "
        numb3(2) = "двести "
        numb3(3) = "триста "
        numb3(4) = "четыреста "
        numb3(5) = "пятьсот "
        numb3(6) = "шестьсот "
        numb3(7) = "семьсот "
        numb3(8) = "восемьсот "
        numb3(9) = "девятьсот "
        Dim txt As String = ""
        If ss@ <> 0 Then
            n = MsgBox("Сумма выходит за границы формата", 16, "Пропись")
            Пропись = ""
            Exit Function
        End If
        For i% = 4 To 1 Step -1
            n = 0
            If triad(i%) > 0 Then
                n% = Int(triad(i%) / 100)
                txt$ = txt$ & numb3(n%)
                n% = Int((triad(i%) - n% * 100) / 10)
                txt$ = txt$ & numb2(n%)
                If n% < 2 Then
                    n% = triad(i%) - (Int(triad(i%) / 10) - n%) * 10
                Else
                    n% = triad(i%) - Int(triad(i%) / 10) * 10
                End If
                Select Case n%
                    Case 1
                        If i% = 2 Then txt$ = txt$ & "одна " Else txt$ = txt$ & "один "
                    Case 2
                        If i% = 2 Then txt$ = txt$ & "две " Else txt$ = txt$ & "два "
                    Case Else
                        txt$ = txt$ & numb1(n%)
                End Select
                Select Case i%
                    Case 2
                        If n% = 0 Or n% > 4 Then
                            txt$ = txt$ + "тысяч "
                        Else
                            If n% = 1 Then txt$ = txt$ + "тысяча " Else txt$ = txt$ + "тысячи "
                        End If
                    Case 3
                        If n% = 0 Or n% > 4 Then
                            txt$ = txt$ + "миллионов "
                        Else
                            If n% = 1 Then txt$ = txt$ + "миллион " Else txt$ = txt$ + "миллиона "
                        End If
                    Case 4
                        If n% = 0 Or n% > 4 Then
                            txt$ = txt$ + "миллиардов "
                        Else
                            If n% = 1 Then txt$ = txt$ + "миллиард " Else txt$ = txt$ + "миллиарда "
                        End If
                End Select
            End If
        Next i%
        If n% = 0 Or n% > 4 Then
            txt$ = txt$ + ""
        Else
            If n% = 1 Then txt$ = txt$ + "" Else txt$ = txt$ + ""
        End If
        txt$ = UCase$(Left$(txt$, 1)) & Mid$(txt$, 2)
        Пропись = txt$
        Return Пропись
    End Function

    Public Function ДобОконч(ByRef ДолжСОконч As String) As String

        Select Case ДолжСОконч
            Case "Директор"
                ДобОконч = ДолжСОконч + "а"
            Case "Повар"
                ДобОконч = ДолжСОконч + "а"
            Case "Шеф-повар"
                ДобОконч = ДолжСОконч + "а"
            Case "Инженер по транспорту"
                ДобОконч = "Инженерa по транспорту"
            Case "Кладовщик"
                ДобОконч = ДолжСОконч + "а"
            Case "Курьер"
                ДобОконч = ДолжСОконч + "а"
            Case "Оператор диспетчерской службы"
                ДобОконч = "Оператора диспетчерской службы"
            Case "Технолог"
                ДобОконч = ДолжСОконч + "а"
            Case "Уборщик помещений"
                ДобОконч = "Уборщика помещений"
            Case "Официант"
                ДобОконч = ДолжСОконч + "а"
            Case "Администратор"
                ДобОконч = ДолжСОконч + "а"
            Case "Администратор зала"
                ДобОконч = "Администратора зала"
            Case "Зам.заведующим производством"
                ДобОконч = "Заместителя заведующего производством"
            Case "Маркетолог"
                ДобОконч = ДолжСОконч + "а"
            Case "Агент по снабжению"
                ДобОконч = "Агента по снабжению"
            Case "Специалист по охране труда"
                ДобОконч = "Специалиста по охране труда"
            Case "Старший бармен"
                ДобОконч = "Старшего бармена"
            Case "Бармен"
                ДобОконч = "Бармена"
            Case "Кухонный рабочий"
                ДобОконч = "Кухонного рабочего"
            Case "Заместитель директора ЗАО «Акцент-Инвест»"
                ДобОконч = "Заместителя директора ЗАО «Акцент-Инвест»"
            Case "Рабочий"
                ДобОконч = "Рабочего"
            Case "Менеджер по персоналу"
                ДобОконч = "Менеджера по персоналу"
            Case "Заместитель заведующего проивоздством"
                ДобОконч = "Заместителя заведующего проивоздством"
        End Select
        If Not ДобОконч <> "" Then
            ДобОконч = окончание(ДолжСОконч, 1)
        End If
        Return ДобОконч
    End Function

    Public Function Склонение(ByVal ставка As String) As String
        Select Case ставка
            Case "1.0"
                Склонение = "ставку"
            Case "0.5"
                Склонение = "ставки"
            Case "0.25"
                Склонение = "ставки"
            Case "0.75"
                Склонение = "ставки"
        End Select
        Return Склонение
    End Function

    Public Function Склонение2(ByVal СрокКонтр As String) As String
        Select Case СрокКонтр
            Case "1"
                Склонение2 = "год"
            Case "2"
                Склонение2 = "года"
            Case "3"
                Склонение2 = "года"
            Case "4"
                Склонение2 = "года"
            Case "5"
                Склонение2 = "лет"
        End Select
        Return Склонение2
    End Function

    Public Function ЧислПроп(ByVal число As String) As String
        Select Case число
            Case "1"
                ЧислПроп = "один"
            Case "2"
                ЧислПроп = "два"
            Case "3"
                ЧислПроп = "три"
            Case "4"
                ЧислПроп = "четыре"
            Case "5"
                ЧислПроп = "пять"
        End Select
        Return ЧислПроп
    End Function
    Public Function разрядстрока(ByVal разр As Integer) As String
        Select Case разр
            Case 1
                разрядстрока = "1-го разряда"
            Case 2
                разрядстрока = "2-го разряда"
            Case 3
                разрядстрока = "3-го разряда"
            Case 4
                разрядстрока = "4-го разряда"
            Case 5
                разрядстрока = "5-го разряда"
            Case 6
                разрядстрока = "6-го разряда"
            Case 7
                разрядстрока = "7-го разряда"
        End Select
        Return разрядстрока
    End Function
    Public Function КалендарДней(ByVal число As Integer) As String
        Select Case число
            Case 1
                Return "календарный день"
            Case 2
                Return "календарных дня"
            Case 3
                Return "календарных дня"
            Case 4
                Return "календарных дня"
            Case 5
                Return "календарных дней"
            Case 6
                Return "календарных дней"
            Case 7
                Return "календарных дней"
            Case 8
                Return "календарных дней"
            Case 9
                Return "календарных дней"
            Case 10
                Return "календарных дней"
            Case 11
                Return "календарных дней"
            Case 12
                Return "календарных дней"
            Case 13
                Return "календарных дней"
            Case 14
                Return "календарных дней"
            Case 15
                Return "календарных дней"
            Case 16
                Return "календарных дней"
            Case 17
                Return "календарных дней"
            Case 18
                Return "календарных дней"
            Case 19
                Return "календарных дней"
            Case 20
                Return "календарных дней"
            Case 21
                Return "календарных день"
            Case 22
                Return "календарных дня"
            Case 23
                Return "календарных дня"
            Case 24
                Return "календарных дня"
            Case 25
                Return "календарных дней"
        End Select

        Return "календарных дней"
    End Function
    Public Function ДолжРодПадежФункц(ByVal Долж As String) As String
        Select Case Долж
            Case "Директор"
                ДолжРодПадежФункц = Долж + "у"
            Case "Повар"
                ДолжРодПадежФункц = Долж + "у"
            Case "Шеф-повар"
                ДолжРодПадежФункц = Долж + "у"
            Case "Инженер по транспорту"
                ДолжРодПадежФункц = "Инженеру по транспорту"
            Case "Кладовщик"
                ДолжРодПадежФункц = Долж + "у"
            Case "Курьер"
                ДолжРодПадежФункц = Долж + "у"
            Case "Оператор диспетчерской службы"
                ДолжРодПадежФункц = "Оператору диспетчерской службы"
            Case "Технолог"
                ДолжРодПадежФункц = Долж + "у"
            Case "Уборщик помещений"
                ДолжРодПадежФункц = "Уборщику помещений"
            Case "Официант"
                ДолжРодПадежФункц = Долж + "у"
            Case "Администратор"
                ДолжРодПадежФункц = Долж + "у"
            Case "Администратор зала"
                ДолжРодПадежФункц = "Администратору зала"
            Case "Зам.заведующим производством"
                ДолжРодПадежФункц = "Зам.заведующего производством"
            Case "Специалист по охране труда"
                ДолжРодПадежФункц = "Специалисту по охране труда"
            Case "Агент по снабжению"
                ДолжРодПадежФункц = "Агенту по снабжению"
            Case "Старший бармен"
                ДолжРодПадежФункц = "Старшему бармену"
            Case "Маркетолог"
                ДолжРодПадежФункц = "Маркетологу"
            Case "Кухонный рабочий"
                ДолжРодПадежФункц = "Кухонному рабочему"
            Case "Менеджер по персоналу"
                ДолжРодПадежФункц = "Менеджеру по персоналу"
            Case "Заместитель заведующего проивоздством"
                ДолжРодПадежФункц = "Заместителю заведующего проивоздством"
            Case "Мойщик посуды"
                ДолжРодПадежФункц = "Мойщику посуды"


        End Select
        If Not ДолжРодПадежФункц <> "" Then
            ДолжРодПадежФункц = окончание(Долж, 2)
        End If

        Return ДолжРодПадежФункц


    End Function
    Public Function ФИОКорРук(ByVal ФИОПол As String, ByVal рукИП As Boolean)

        Dim РуковИП As String
        If рукИП = True Or рукИП = "True" Then
            РуковИП = "ИП "
        Else
            РуковИП = ""
        End If



        Dim nm As String = ФИОПол
        Dim nm0 As Integer = Len(ФИОПол)
        Dim nm1 As String = Strings.Left(nm, InStr(nm, " "))
        Dim nm2 As Integer = Len(nm1)
        Dim nm3 As String = Strings.Right(nm, (nm0 - nm2))
        Dim nm31 As Integer = Len(nm3)
        Dim nm4 As String = Strings.UCase(Strings.Left(Strings.Left(nm3, InStr(nm3, " ")), 1))
        Dim nm41 As Integer = Len(Strings.Left(nm3, InStr(nm3, " ")))
        Dim nm5 As String = Strings.UCase(Strings.Left(Strings.Right(nm3, nm31 - nm41), 1))




        Dim ФИОКор As String = РуковИП & nm1 & "" & nm4 & "." & nm5 & "."
        Return ФИОКор
    End Function

    Public Function ОсновУвольн(ByVal Text As String) As String

        Select Case Text
            Case "По соглашению сторон"
                Return "по соглашению сторон в соответствии с п. 1 ч. 2 статьи 35 Трудового кодекса Республики Беларусь."

            Case "По истечению срока контракта"
                Return "в связи с истечением срока действия контракта в соответствии с п. 2 ч. 2 ст. 35 Трудового кодекса Республики Беларусь."
        End Select
        Return ""
    End Function

    Public Function ДолжТворПадеж(ByVal Долж As String) As String
        Select Case Долж
            Case "Директор"
                ДолжТворПадеж = Долж + "ом"
            Case "Повар"
                ДолжТворПадеж = Долж + "ом"
            Case "Шеф-повар"
                ДолжТворПадеж = Долж + "ом"
            Case "Инженер по транспорту"
                ДолжТворПадеж = "Инженером по транспорту"
            Case "Кладовщик"
                ДолжТворПадеж = Долж + "ом"
            Case "Курьер"
                ДолжТворПадеж = Долж + "ом"
            Case "Оператор диспетчерской службы"
                ДолжТворПадеж = "Оператором диспетчерской службы"
            Case "Технолог"
                ДолжТворПадеж = Долж + "ом"
            Case "Уборщик помещений"
                ДолжТворПадеж = "Уборщиком помещений"
            Case "Официант"
                ДолжТворПадеж = Долж + "ом"
            Case "Администратор"
                ДолжТворПадеж = Долж + "ом"
            Case "Администратор зала"
                ДолжТворПадеж = "Администратором зала"
            Case "Зам.заведующим производством"
                ДолжТворПадеж = "Зам.заведующим производством"
            Case "Специалист по охране труда"
                ДолжТворПадеж = "Специалистом по охране труда"
            Case "Агент по снабжению"
                ДолжТворПадеж = "Агентом по снабжению"
            Case "Старший бармен"
                ДолжТворПадеж = "Старшим барменом"
            Case "Маркетолог"
                ДолжТворПадеж = "Маркетологом"
            Case "Кухонный рабочий"
                ДолжТворПадеж = "Кухонным рабочим"
            Case "Зав.производством"
                ДолжТворПадеж = "Зав.производством"
            Case "Инженер"
                ДолжТворПадеж = "Инженером"
            Case "Менеджер по персоналу"
                ДолжТворПадеж = "Менеджером по персоналу"
            Case "Заместитель заведующего проивоздством"
                ДолжТворПадеж = "Заместителем заведующего проивоздством"
            Case "Мойщик посуды"
                ДолжТворПадеж = "Мойщиком посуды"


        End Select

        If Not ДолжТворПадеж <> "" Then
            ДолжТворПадеж = окончание(Долж, 3)
        End If
        Return ДолжТворПадеж


    End Function

    Public Function ФормСобствКор(ByVal Полн As String) As String
        'Dim StrSql3 As String = "SELECT ФормаСобств.Сокращенное FROM ФормаСобств WHERE ФормаСобств.ПолноеНазвание ='" & Полн & "'"
        Dim ds = dtformft.Select("ПолноеНазвание ='" & Полн & "'")
        Return ds(0).Item(2).ToString
    End Function
    Private Sub ProcEx() Handles Proc.Exited
        MsgBox("Процесс завершен")
    End Sub
    Public Sub ПечатьДоковКол(ByVal mass As String, ByVal d As Integer)

        Dim wdApp As New Microsoft.Office.Interop.Word.Application
        Dim wdDoc As Microsoft.Office.Interop.Word.Document
        wdApp.Visible = False
        wdDoc = wdApp.Documents.Open(FileName:=mass) 'заявление
        Try
            wdDoc.PrintOut(True,,,,,,, d)
            'Process.Start("rundll32", "shell32,Control_RunDLL main.cpl @2")
            'Proc = Process.GetProcessesByName("Calc")(0)
            'Proc.EnableRaisingEvents = True

        Catch ex As Exception
            KillProc()
        End Try

        'wdApp.Visible = True

        Try

        Catch ex As Exception

        End Try
        wdDoc.Close()
        wdApp.Quit()
    End Sub



    Public Sub ПечатьДоков(ByVal mass() As String)


        Dim wdApp As New Microsoft.Office.Interop.Word.Application
        Dim wdDoc As Microsoft.Office.Interop.Word.Document
        wdApp.Visible = False
        Dim i As Integer
        For i = 0 To mass.Length - 1
            wdDoc = wdApp.Documents.Open(FileName:=mass(i)) 'заявление
            Try
                wdDoc.PrintOut(True)
            Catch ex As Exception

            End Try
            wdDoc.Close()


        Next

        wdApp.Quit()

        'Dim wdDoc1 As Microsoft.Office.Interop.Word.Document = wdApp.Documents.Open(FileName:=СохрКонтр) 'контракт
        '    wdDoc1.PrintOut(True,,,,,,, 2)

        '    Dim wdDoc2 As Microsoft.Office.Interop.Word.Document = wdApp.Documents.Open(FileName:=СохрПрик) 'приказ
        '    wdDoc2.PrintOut(True)

        ''wdDoc.PrintOut(Range:=4, Pages:="2-4")  параметры для диапазона Range [url]


        'wdDoc1.Close()
        'wdDoc2.Close()


    End Sub
    Public Sub ПечатьДоков2(ByVal mass() As String, ParamArray mass2() As String)
        Dim wdApp As New Microsoft.Office.Interop.Word.Application
        Dim wdDoc As Microsoft.Office.Interop.Word.Document
        wdApp.Visible = False
        Dim i As Integer
        For i = 0 To mass.Length - 1
            If mass(i) = Nothing Then Continue For
            Try
                wdDoc = wdApp.Documents.Open(FileName:=mass(i)) 'заявление
                wdDoc.PrintOut(True)

            Catch ex As Exception

            End Try

            'Dim p = Process.  (mass(i))

            'Process.WaitForExit(p)

        Next
        'wdApp.Visible = True
        If mass2.Length > 1 Then
            wdDoc = Nothing
            For i = 0 To mass2.Length - 1
                If mass2(i) = Nothing Then Continue For
                Try
                    wdDoc = wdApp.Documents.Open(FileName:=mass2(i)) 'заявление
                    wdDoc.PrintOut(True)
                Catch ex As Exception

                End Try

            Next
        End If

        wdDoc.Close()
        wdApp.Quit()

        'Dim wdDoc1 As Microsoft.Office.Interop.Word.Document = wdApp.Documents.Open(FileName:=СохрКонтр) 'контракт
        '    wdDoc1.PrintOut(True,,,,,,, 2)

        '    Dim wdDoc2 As Microsoft.Office.Interop.Word.Document = wdApp.Documents.Open(FileName:=СохрПрик) 'приказ
        '    wdDoc2.PrintOut(True)

        ''wdDoc.PrintOut(Range:=4, Pages:="2-4")  параметры для диапазона Range [url]


        'wdDoc1.Close()
        'wdDoc2.Close()


    End Sub

    Public Sub ПечатьДоков3(ByVal mass() As String, ByVal mass2() As String, ByVal mass3() As String)
        Dim wdApp As New Microsoft.Office.Interop.Word.Application
        Dim wdDoc As Microsoft.Office.Interop.Word.Document
        wdApp.Visible = False
        Dim i As Integer
        For i = 0 To mass.Length - 1
            If mass(i) = Nothing Then Continue For
            Try
                wdDoc = wdApp.Documents.Open(FileName:=mass(i)) 'заявление
                wdDoc.PrintOut(True)
                wdDoc = Nothing
            Catch ex As Exception

            End Try

        Next

        wdApp.Quit()




        Dim wdApp1 As New Microsoft.Office.Interop.Word.Application
        Dim wdDoc1 As Microsoft.Office.Interop.Word.Document
        wdApp1.Visible = False


        For i = 0 To mass2.Length - 1
            If mass2(i) = Nothing Then Continue For
            Try
                wdDoc1 = wdApp1.Documents.Open(FileName:=mass2(i)) 'заявление
                wdDoc1.PrintOut(True)
                wdDoc1 = Nothing
            Catch ex As Exception

            End Try

        Next
        wdApp1.Quit()


        Dim wdApp2 As New Microsoft.Office.Interop.Word.Application
        Dim wdDoc2 As Microsoft.Office.Interop.Word.Document
        wdApp2.Visible = False
        For i = 0 To mass3.Length - 1
            If mass3(i) = Nothing Then Continue For
            Try
                wdDoc2 = wdApp2.Documents.Open(FileName:=mass3(i)) 'заявление
                wdDoc2.PrintOut(True)
                wdDoc2 = Nothing
            Catch ex As Exception

            End Try

        Next
        wdApp2.Quit()

    End Sub

    Public Function ДатаУведомл(ByVal срокконт As String, ByVal iTms As String) As String 'ДАТА УВЕДОМЛЕНИЯ О ПРОДЛЕНИИ КОНТРАКТА
        Dim iTm As Date = CDate(iTms)
        Dim Датаув As String
        Select Case срокконт
            Case "1"
                iTm = iTm.AddMonths(11) 'дата уведомления через 11 месяцев после даты контракта
                iTm = iTm.AddDays(-2)
                Датаув = iTm
                Return Датаув
            Case "2"

                iTm = iTm.AddMonths(23) 'дата уведомления через 11 месяцев после даты контракта
                iTm = iTm.AddDays(-2)
                Датаув = iTm
                Return Датаув

            Case "3"
                iTm = iTm.AddMonths(35) 'дата уведомления через 11 месяцев после даты контракта
                iTm = iTm.AddDays(-2)
                Датаув = iTm
                Return Датаув
            Case "4"
                iTm = iTm.AddMonths(47) 'дата уведомления через 11 месяцев после даты контракта
                iTm = iTm.AddDays(-2)
                Датаув = iTm
                Return Датаув
            Case "5"
                iTm = iTm.AddMonths(59) 'дата уведомления через 11 месяцев после даты контракта
                iTm = iTm.AddDays(-2)
                Датаув = iTm
                Return Датаув

        End Select
        Return Now
    End Function
    Public Function ЧислоПропис(ByVal чис As String)

        Dim ikstr, ipstr, ПолнСтрока As String
        Dim ik, ip As Integer
        Dim sf As Double
        чис = чис.Replace(".", ",")

        Dim b As Boolean = чис.Contains(",")

        If b Then

            sf = Nothing
            sf = CType(чис, Double)
            ik = Math.Floor(sf)
            ikstr = Пропись(ik) & "бел.руб, "

            sf = Math.Round(sf - ik, 2)
            ipstr = CType(sf, String)
            If ipstr.Length > 3 Then
                ipstr = Strings.Right(ipstr, 2)
            Else
                ipstr = Strings.Right(ipstr, 1)
            End If
            ip = CType(ipstr, Integer)
            ipstr = Strings.LCase(Пропись(ip) & " коп.")
            ПолнСтрока = ikstr & ipstr
        Else
            ik = CType(чис, Integer)
            ikstr = Пропись(ik) & "бел.руб 00 копеек"
            ПолнСтрока = ikstr
        End If


        Return ПолнСтрока



    End Function
    Public Function ЧислоПрописДляСправки(ByVal чис As String)

        Dim ikstr, ipstr, ПолнСтрока As String
        Dim ik, ip As Integer
        Dim sf As Double
        чис = чис.Replace(".", ",")

        Dim b As Boolean = чис.Contains(",")

        If b Then

            sf = Nothing
            sf = CType(чис, Double)
            ik = Math.Floor(sf)
            ikstr = Пропись(ik) & "бел.руб, "

            sf = Math.Round(sf - ik, 2)
            ipstr = CType(sf, String)
            If ipstr.Length > 3 Then
                ipstr = Strings.Right(ipstr, 2)
            Else
                ipstr = Strings.Right(ipstr, 1)
                ipstr = ipstr & "0"
            End If
            ip = CType(ipstr, Integer)
            If ip = 0 Then
                ipstr = "ноль коп."
                ПолнСтрока = ikstr & ipstr
            Else
                ipstr = Strings.LCase(Пропись(ip) & " коп.")
                ПолнСтрока = ikstr & ipstr
            End If
        Else
            ik = CType(чис, Integer)
            ikstr = Пропись(ik) & "бел.руб 00 копеек"
            ПолнСтрока = ikstr
        End If


        Return ПолнСтрока

    End Function

    Public Function Org(ByVal НазвОргП As String) As DataRow()
        'Dim strsql As String = "SELECT * FROM Клиент WHERE НазвОрг='" & НазвОргП & "'"
        'Dim df As DataTable = Selects(strsql)

        Dim df = dtClientAll.Select("НазвОрг='" & НазвОргП & "'")
        Return df
    End Function
    Public Function Sotrudnic(ByVal НазвОргП As Integer) As DataRow()
        'Dim strsql As String = "SELECT * FROM Сотрудники WHERE КодСотрудники=" & НазвОргП & ""
        'Dim df As DataTable = Selects(strsql)

        Dim df = dtSotrudnikiAll.Select("КодСотрудники=" & НазвОргП & "")

        Return df
    End Function

    Public Function Подоходный(ByVal число As String) As Object
        Dim ipstr, ПолнСтрока, fszn, ПолСтр2 As String
        Dim ik, gk As Integer
        число = число.Replace(".", ",")
        Dim b As Boolean = число.Contains(",")
        Dim c, h As Boolean
        Dim sf, bf, fz, bz, sf1 As Double
        Dim a, j As String

        sf1 = CType(число, Double)
        sf = Math.Round(sf1 * 13 / 100, 2)
        fz = Math.Round(sf1 * 1 / 100, 2)
        ipstr = CType(sf, String)
        fszn = CType(fz, String)
        h = fszn.Contains(",")
        If h = True Then
            gk = Math.Floor(fz)
            bz = Math.Round(fz - gk, 2)
            ПолСтр2 = CType(bz, String)
            If ПолСтр2.Length = 4 Then
                a = fszn
            Else
                a = fszn & "0"

            End If
        Else
            a = fszn

        End If
        c = ipstr.Contains(",")

        If c = True Then
            ik = Math.Floor(sf)
            bf = Math.Round(sf - ik, 2)
            ПолнСтрока = CType(bf, String)
            If ПолнСтрока.Length = 4 Then
                j = ipstr
            Else
                j = ipstr & "0"

            End If
        Else
            j = ipstr
        End If
        Return {j, a}
    End Function

    Public Function ВремяНач(ByVal f As String) As String
        Select Case f
            Case "8.30"
                Return "8 часов 30 минут"
            Case "9.00"
                Return "9 часов 00 минут"
            Case "10.00"
                Return "10 часов 00 минут"
            Case "10.30"
                Return "10 часов 30 минут"
            Case "11.00"
                Return "11 часов 00 минут"
            Case "12.00"
                Return "12 часов 00 минут"
            Case "13.00"
                Return "13 часов 00 минут"
            Case "14.00"
                Return "14 часов 00 минут"
            Case "15.00"
                Return "15 часов 00 минут"
            Case "16.00"
                Return "16 часов 00 минут"
            Case "16.30"
                Return "16 часов 30 минут"
            Case "17.00"
                Return "17 часов 00 минут"
            Case "17.30"
                Return "17 часов 30 минут"
            Case "18.00"
                Return "18 часов 00 минут"
            Case "18.30"
                Return "18 часов 30 минут"
            Case "19.00"
                Return "19 часов 00 минут"
            Case "19.30"
                Return "19 часов 30 минут"
            Case "20.00"
                Return "20 часов 00 минут"
            Case "20.30"
                Return "20 часов 30 минут"
            Case "21.00"
                Return "21 часов 00 минут"
            Case "21.30"
                Return "21 часов 30 минут"
            Case "22.00"
                Return "22 часов 00 минут"
            Case "22.30"
                Return "22 часов 30 минут"
            Case "23.00"
                Return "23 часов 00 минут"
            Case "23.30"
                Return "23 часов 30 минут"
            Case "00.00"
                Return "00 часов 00 минут"

        End Select
        Return "9 часов 00 минут"
    End Function
    Public Function ПровИндивидКонтр(ByVal d As String) As Boolean
        'Dim strValues As String() = New String() {КонтрПровИндивид} 'из массива в лист оф очень класная штука
        Dim strList As List(Of String) = КонтрПровИндивид.ToList()


        For i As Integer = 0 To strList.Count - 1
            If strList(i) = d Then
                Return True
            End If
        Next
        Return False
    End Function

End Module
