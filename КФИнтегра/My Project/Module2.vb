Module Module2a
    Dim n As Integer
    Function Пропись2(ByVal число As Integer) As String

        Static triad(4) As Integer, numb1(0 To 19) As String, numb2(0 To 9) As String, numb3(0 To 9) As String
        If число = 0 Then
            Пропись2 = ""
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
            n = MsgBox("Сумма выходит за границы формата", 16, "Пропись2")
            Пропись2 = ""
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
        Пропись2 = txt$
        Return txt$
    End Function


End Module
