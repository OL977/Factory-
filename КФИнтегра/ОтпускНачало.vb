Public Class ОтпускНачало
    Dim strsql, Разн, СохрЗаявл As String
    Dim Files2(), d8, d1 As String
    Dim dt As DataTable
    Dim Int As Integer = 1
    Dim idcn As Integer
    Dim hg As Integer
    Dim dssotr As DataTable
    Dim dsorg As DataTable
    Dim massFTP3 As New ArrayList()


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim d2, d As Date
        Dim s, а As String


        If MessageBox.Show("Сохранить данные?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Me.Close()
            Exit Sub
        End If

        If TextBox4.Text = "" Then
            TextBox4.Text = "0"
        End If


        If MaskedTextBox1.MaskCompleted = False And TextBox2.Text = "" And MaskedTextBox2.MaskCompleted = False And TextBox3.Text = "" And TextBox4.Text = "0" Then
            Me.Close()
            Exit Sub
        End If

        If MaskedTextBox1.MaskCompleted = True And TextBox2.Text = "" Or IsNumeric(TextBox2.Text) = False Then
            MessageBox.Show("Выберите продолжительность для первой части отпуска!", Рик)
            Exit Sub
        End If

        If MaskedTextBox2.MaskCompleted = True And TextBox3.Text = "" Then
            MessageBox.Show("Выберите продолжительность для второй части отпуска!", Рик)
            Exit Sub
        End If

        If TextBox5.Text = "" Or IsNumeric(TextBox5.Text) = False Then
            MessageBox.Show("Введите номер приказа!", Рик)
            Exit Sub
        End If

        If MaskedTextBox3.MaskCompleted = False Then
            MessageBox.Show("Выберите дату приказа!", Рик)
            Exit Sub
        End If





        Me.Cursor = Cursors.WaitCursor

        If MaskedTextBox1.MaskCompleted = True And TextBox2.Text <> "" Then
            d = MaskedTextBox1.Text
            d = d.AddDays(CType(TextBox2.Text, Integer) - 1)
            d1 = d.ToShortDateString
        Else
            d1 = ""
        End If


        If MaskedTextBox2.MaskCompleted = True And TextBox3.Text <> "" Then
            d2 = MaskedTextBox2.Text
            d2 = d2.AddDays(CType(TextBox3.Text, Integer) - 1)
            d8 = d2.ToShortDateString
        Else
            d8 = ""
        End If

        If MaskedTextBox1.MaskCompleted = True And TextBox2.Text <> "" And MaskedTextBox2.MaskCompleted = True And TextBox3.Text <> "" Then
            Int = 2
            s = CType(CType(TextBox2.Text, Integer) + CType(TextBox3.Text, Integer), String)
        ElseIf TextBox2.Text <> "" And TextBox3.Text = "" Then
            s = TextBox2.Text
        Else
            s = TextBox3.Text
        End If

        'If TextBox4.Text = "" Then 'если нет остатка за прошлый год
        '    а = CType(Отпуск.ДнОтпус - CType(s, Integer), String)
        '    Разн = ""
        'Else
        '    Try
        '        If CType(TextBox4.Text, Integer) > CType(s, Integer) Then
        '            Разн = CType(CType(TextBox4.Text, Integer) - CType(s, Integer), String)
        '            а = Отпуск.ДнОтпус
        '        Else
        '            а = Отпуск.ДнОтпус
        '            а = CType(CType(а, Integer) - CType(s, Integer) - CType(TextBox4.Text, Integer), String)
        '            Разн = ""

        '        End If
        '    Catch ex As Exception
        '        MessageBox.Show("Заполните период первого отрезка отпуска!" & vbCrLf & "И повторите ввод остатка неиспользованных дней отпуска ", Рик)
        '        Exit Sub
        '    End Try


        'End If


        Dim ИтогоВЧисло As String

        If TextBox4.Text = "0" Then
            ИтогоВЧисло = Отпуск.Grid3.CurrentRow.Cells("Итого").Value
            Разн = "0"
            Try
                а = CType(CType(ИтогоВЧисло, Integer) - (CType(TextBox2.Text, Integer) + CType(TextBox3.Text, Integer)), String)

            Catch ex As Exception
                а = CType(CType(ИтогоВЧисло, Integer) - CType(TextBox2.Text, Integer), String)

            End Try


            'If TextBox3.Text = "" Then
            '    а = CType(CType(ИтогоВЧисло, Integer) - CType(TextBox2.Text, Integer), String)
            'ElseIf TextBox3.Text <> "" And TextBox2.Text <> "" Then
            '    а = CType(CType(ИтогоВЧисло, Integer) - CType(TextBox2.Text, Integer) + CType(TextBox3.Text, Integer), String)
            'End If
        End If

        If Not TextBox4.Text = "0" And CType(TextBox4.Text, Integer) > 0 Then

            Разн = TextBox4.Text
            ИтогоВЧисло = CType(CType(TextBox4.Text, Integer) + CType(Отпуск.Grid3.CurrentRow.Cells("Положено дней отпуска").Value, Integer), String)

            Try
                а = CType(CType(ИтогоВЧисло, Integer) - (CType(TextBox2.Text, Integer) + CType(TextBox3.Text, Integer)), String)

            Catch ex As Exception
                а = CType(CType(ИтогоВЧисло, Integer) - CType(TextBox2.Text, Integer), String)

            End Try



        End If






        Dim msk1, msk2 As String
        If MaskedTextBox1.MaskCompleted = True Then
            msk1 = MaskedTextBox1.Text
        Else
            msk1 = ""
        End If

        If MaskedTextBox2.MaskCompleted = True Then
            msk2 = MaskedTextBox2.Text
        Else
            msk2 = ""
        End If

        Dim rt As Integer
        If TextBox3.Text <> "" Then
            rt = (CType(TextBox2.Text, Integer) + CType(TextBox3.Text, Integer))
        Else
            rt = CType(TextBox2.Text, Integer)
        End If

        Dim rf As Integer
        If Разн <> "" Then
            rf = (CType(Отпуск.ДнОтпус, Integer) + CType(Разн, Integer))
        Else
            rf = CType(Отпуск.ДнОтпус, Integer)
        End If




        If rt > rf Then
            MessageBox.Show("Фактический срок отпуска не может быть больше запланированного!", Рик)
            Me.Cursor = Cursors.Default
            Exit Sub
        End If


        strsql = "UPDATE ОтпускСотрудники SET ДатаНач1='" & msk1 & "', Продолж1='" & TextBox2.Text & "',
ДатаОконч1='" & d1 & "',ДатаНач2='" & msk2 & "', Продолж2='" & TextBox3.Text & "', ДатаОконч2='" & d8 & "',
Израсходовано='" & s & "', ОсталосьЭтотГод ='" & а & "' , ОсталосьПрошлГод= '" & Разн & "', Итого='" & ИтогоВЧисло & "'
WHERE Код=" & Отпуск.idgr3cod & ""
        Updates(strsql)

        Доки()

        Статистика1(TextBox1.Text, "Отправка сотрудника в отпуск", Отпуск.ComboBox2.Text)
        hg = 0
        Отпуск.grcellclick()
        Отпуск.grid3activ()
        Me.Cursor = Cursors.Default
        Me.Close()
    End Sub
    Private Sub Доки()



        Dim w1, w2 As String
        'strsql = ""
        'strsql = "SELECT ПродлКонтрС, ПродлКонтрПо FROM КарточкаСотрудника WHERE IDСотр= " & idcn & ""
        'Dim dsd As DataTable = Selects(strsql)

        Dim dsd = dtKartochkaSotrudnikaAll.Select("IDСотр= " & idcn & "")


        'If errds = 1 Then
        '    Dim strsql As String = "SELECT ДатаПриема FROM КарточкаСотрудника WHERE IDСотр= " & idcn & ""
        '    Dim dv As DataTable = Selects(strsql)
        '    w1 = dv.Rows(0).Item(0).ToString
        '    Dim strsql4 As String = "SELECT СрокОкончКонтр FROM ДогСотрудн WHERE IDСотр= " & idcn & ""
        '    Dim dv1 As DataTable = Selects(strsql4)
        '    w2 = dv1.Rows(0).Item(0).ToString
        'Else
        '    w1 = dsd.Rows(0).Item(0).ToString
        '    w2 = dsd.Rows(0).Item(1).ToString
        'End If

        If dsd(0).Item("ПродлКонтрС").ToString <> "" Then
            w1 = dsd(0).Item("ПродлКонтрС").ToString
            w2 = dsd(0).Item("ПродлКонтрПо").ToString
        Else
            'Dim strsql As String = "SELECT ДатаПриема FROM КарточкаСотрудника WHERE IDСотр= " & idcn & ""
            'Dim dv As DataTable = Selects(strsql)
            w1 = dsd(0).Item("ДатаПриема").ToString
            'Dim strsql4 As String = "SELECT СрокОкончКонтр FROM ДогСотрудн WHERE IDСотр= " & idcn & ""
            'Dim dv1 As DataTable = Selects(strsql4)

            Dim dv1 = dtDogovorSotrudnikAll.Select("IDСотр= " & idcn & "")
            w2 = dv1(0).Item("СрокОкончКонтр").ToString
        End If


        'Dim strsql1 As String = "SELECT * FROM Клиент WHERE НазвОрг= '" & Отпуск.ComboBox2.Text & "'"
        'Dim dsd2 As DataTable = Selects(strsql1)

        Dim dsd2 = dtClientAll.Select("НазвОрг= '" & Отпуск.ComboBox2.Text & "'")


        'Dim strsql2 As String = "SELECT Должность, Разряд FROM Штатное WHERE ИДСотр= " & idcn & ""
        'Dim dsd3 As DataTable = Selects(strsql2)
        Dim dsd3 = dtShtatnoeAll.Select("ИДСотр= " & idcn & "")


        'Dim strsql3 As String = "SELECT ФамилияДляЗаявления, ИмяДляЗаявления, ОтчествоДляЗаявления FROM Сотрудники WHERE КодСотрудники= " & idcn & ""
        'Dim dsd4 As DataTable = Selects(strsql3)

        Dim dsd4 = dtSotrudnikiAll.Select("КодСотрудники= " & idcn & "")


        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document
        oWord = CreateObject("Word.Application")
        oWord.Visible = False

        Начало("PrikazNaOtpusk.doc")
        oWordDoc = oWord.Documents.Add(firthtPath & "\PrikazNaOtpusk.doc")

        With oWordDoc.Bookmarks
            .Item("П1").Range.Text = MaskedTextBox3.Text
            .Item("П2").Range.Text = TextBox5.Text & " - отп"
            .Item("П3").Range.Text = Trim(InputName1(TextBox1.Text, "ОтпускТруд"))
            If Int = 1 Then
                .Item("П4").Range.Text = TextBox2.Text
                .Item("П5").Range.Text = MaskedTextBox1.Text
                .Item("П6").Range.Text = d1
            Else
                .Item("П4").Range.Text = TextBox3.Text
                .Item("П5").Range.Text = MaskedTextBox2.Text
                .Item("П6").Range.Text = d8
            End If
            .Item("П7").Range.Text = Strings.Left(w1, 10)
            .Item("П8").Range.Text = Strings.Left(w2, 10)
            .Item("П9").Range.Text = dsd4(0).Item("ФамилияДляЗаявления").ToString & " " & Strings.Left(dsd4(0).Item("ИмяДляЗаявления").ToString, 1) & "." & Strings.Left(dsd4(0).Item("ОтчествоДляЗаявления").ToString, 1) & "."

            If dsd2(0).Item(1).ToString = "Индивидуальный предприниматель" Then
                .Item("П10").Range.Text = dsd2(0).Item(18).ToString
            Else
                .Item("П10").Range.Text = dsd2(0).Item(18).ToString & " " & ФормСобствКор(dsd2(0).Item(1).ToString) & " """ & Отпуск.ComboBox2.Text & """ "

            End If

            If dsd2(0).Item(31) = True Then
                .Item("П11").Range.Text = ФИОКорРук(dsd2(0).Item(19).ToString, True)
            Else
                .Item("П11").Range.Text = ФИОКорРук(dsd2(0).Item(19).ToString, False)
            End If

            .Item("П12").Range.Text = ФИОКорРук(TextBox1.Text, False)

            If dsd3(0).Item("Разряд").ToString <> "" And Not dsd3(0).Item("Разряд").ToString = "-" Then

                If IsNumeric(dsd3(0).Item("Разряд").ToString) Then
                    .Item("П13").Range.Text = Strings.LCase(ДолжРодПадежФункц(dsd3(0).Item("Должность").ToString)) & " " & разрядстрока(CType(dsd3(0).Item("Разряд").ToString, Integer))
                End If

            Else
                .Item("П13").Range.Text = Strings.LCase(ДолжРодПадежФункц(dsd3(0).Item("Должность").ToString))
            End If


            .Item("П25").Range.Text = dsd2(0).Item(1).ToString

            If dsd2(0).Item(1).ToString = "Индивидуальный предприниматель" Then
                .Item("П26").Range.Text = dsd2(0).Item(0).ToString
            Else
                .Item("П26").Range.Text = " «" & dsd2(0).Item(0).ToString & "» "
            End If

            .Item("П27").Range.Text = dsd2(0).Item(4).ToString
            .Item("П28").Range.Text = dsd2(0).Item(2).ToString
            .Item("П29").Range.Text = dsd2(0).Item(14).ToString
            .Item("П30").Range.Text = dsd2(0).Item(12).ToString
            .Item("П31").Range.Text = dsd2(0).Item(11).ToString
            .Item("П33").Range.Text = dsd2(0).Item(8).ToString
            .Item("П34").Range.Text = dsd2(0).Item(6).ToString

        End With

        Dim Name As String = TextBox5.Text & "-отп " & ФИОКорРук(TextBox1.Text, False) & " с " & Me.MaskedTextBox1.Text & " по " & d1 & " (Приказ.ТрудОтпуск часть 1)" & ".doc"
        Dim СохрЗак As New List(Of String)
        СохрЗак.AddRange(New String() {Отпуск.ComboBox2.Text & "\Приказ\" & Now.Year & "\", Name})
        oWordDoc.SaveAs2(PathVremyanka & Name,,,,,, False)
        oWordDoc.Close(True)
        oWord.Quit(True)
        Конец(Отпуск.ComboBox2.Text & "\Приказ\" & Now.Year, Name, idcn, Отпуск.ComboBox2.Text, "\PrikazNaOtpusk.doc", "Приказ.ТрудОтпуск часть 1")
        massFTP3.Add(СохрЗак)







        'If Not IO.Directory.Exists(OnePath & Отпуск.ComboBox2.Text & "\Приказ\" & Now.Year & "\Отпуск") Then
        '    IO.Directory.CreateDirectory(OnePath & Отпуск.ComboBox2.Text & "\Приказ\" & Now.Year & "\Отпуск")
        'End If
        'Dim СохрПрик As String

        'If Int = 1 And Not Отпуск.ДнОтпус = CType(TextBox2.Text, Integer) Then
        '    oWordDoc.SaveAs2("C:\Users\Public\Documents\Рик\" & TextBox5.Text & "-отп " & ФИОКорРук(TextBox1.Text, False) & " с " & Me.MaskedTextBox1.Text & " по " & d1 & " (Приказ.ТрудОтпуск часть 1)" & ".doc",,,,,, False)
        '    Try
        '        IO.File.Copy("C:\Users\Public\Documents\Рик\" & TextBox5.Text & "-отп " & ФИОКорРук(TextBox1.Text, False) & " с " & Me.MaskedTextBox1.Text & " по " & d1 & " (Приказ.ТрудОтпуск часть 1)" & ".doc", OnePath & Отпуск.ComboBox2.Text & "\Приказ\" & Now.Year & "\Отпуск\" & TextBox5.Text & "-отп " & ФИОКорРук(TextBox1.Text, False) & " с " & Me.MaskedTextBox1.Text & " по " & d1 & " (Приказ.ТрудОтпуск часть 1)" & ".doc")
        '    Catch ex As Exception
        '        If MessageBox.Show("Приказ № " & TextBox5.Text & "-отп " & ФИОКорРук(TextBox1.Text, False) & " от " & Me.MaskedTextBox1.Text & " уже существует. Заменить старый документ новым?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then
        '            IO.File.Delete(OnePath & Отпуск.ComboBox2.Text & "\Приказ\" & Now.Year & "\Отпуск\" & TextBox5.Text & "-отп " & ФИОКорРук(TextBox1.Text, False) & " с " & Me.MaskedTextBox1.Text & " по " & d1 & " (Приказ.ТрудОтпуск часть 1)" & ".doc")
        '            IO.File.Copy("C:\Users\Public\Documents\Рик\" & TextBox5.Text & "-отп " & ФИОКорРук(TextBox1.Text, False) & " с " & Me.MaskedTextBox1.Text & " по " & d1 & " (Приказ.ТрудОтпуск часть 1)" & ".doc", OnePath & Отпуск.ComboBox2.Text & "\Приказ\" & Now.Year & "\Отпуск\" & TextBox5.Text & "-отп " & ФИОКорРук(TextBox1.Text, False) & " с " & Me.MaskedTextBox1.Text & " по " & d1 & " (Приказ.ТрудОтпуск часть 1)" & ".doc")
        '        End If
        '    End Try
        '    СохрПрик = OnePath & Отпуск.ComboBox2.Text & "\Приказ\" & Now.Year & "\Отпуск\" & TextBox5.Text & "-отп " & ФИОКорРук(TextBox1.Text, False) & " с " & Me.MaskedTextBox1.Text & " по " & d1 & " (Приказ.ТрудОтпуск часть 1)" & ".doc"

        'ElseIf Int = 1 And Отпуск.ДнОтпус = CType(TextBox2.Text, Integer) Then

        '    oWordDoc.SaveAs2("C:\Users\Public\Documents\Рик\" & TextBox5.Text & "-отп " & ФИОКорРук(TextBox1.Text, False) & " с " & Me.MaskedTextBox1.Text & " по " & d1 & " (Приказ.ТрудОтпуск)" & ".doc",,,,,, False)
        '    Try
        '        IO.File.Copy("C:\Users\Public\Documents\Рик\" & TextBox5.Text & "-отп " & ФИОКорРук(TextBox1.Text, False) & " с " & Me.MaskedTextBox1.Text & " по " & d1 & " (Приказ.ТрудОтпуск)" & ".doc", OnePath & Отпуск.ComboBox2.Text & "\Приказ\" & Now.Year & "\Отпуск\" & TextBox5.Text & "-отп " & ФИОКорРук(TextBox1.Text, False) & " с " & Me.MaskedTextBox1.Text & " по " & d1 & " (Приказ.ТрудОтпуск)" & ".doc")
        '    Catch ex As Exception
        '        If MessageBox.Show("Приказ № " & TextBox5.Text & "-отп " & ФИОКорРук(TextBox1.Text, False) & " от " & Me.MaskedTextBox1.Text & " уже существует. Заменить старый документ новым?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then
        '            IO.File.Delete(OnePath & Отпуск.ComboBox2.Text & "\Приказ\" & Now.Year & "\Отпуск\" & TextBox5.Text & "-отп " & ФИОКорРук(TextBox1.Text, False) & " с " & Me.MaskedTextBox1.Text & " по " & d1 & " (Приказ.ТрудОтпуск)" & ".doc")
        '            IO.File.Copy("C:\Users\Public\Documents\Рик\" & TextBox5.Text & "-отп " & ФИОКорРук(TextBox1.Text, False) & " с " & Me.MaskedTextBox1.Text & " по " & d1 & " (Приказ.ТрудОтпуск)" & ".doc", OnePath & Отпуск.ComboBox2.Text & "\Приказ\" & Now.Year & "\Отпуск\" & TextBox5.Text & "-отп " & ФИОКорРук(TextBox1.Text, False) & " с " & Me.MaskedTextBox1.Text & " по " & d1 & " (Приказ.ТрудОтпуск)" & ".doc")
        '        End If
        '    End Try

        '    СохрПрик = OnePath & Отпуск.ComboBox2.Text & "\Приказ\" & Now.Year & "\Отпуск\" & TextBox5.Text & "-отп " & ФИОКорРук(TextBox1.Text, False) & " с " & Me.MaskedTextBox1.Text & " по " & d1 & " (Приказ.ТрудОтпуск)" & ".doc"

        'Else
        '    oWordDoc.SaveAs2("C:\Users\Public\Documents\Рик\" & TextBox5.Text & "-отп " & ФИОКорРук(TextBox1.Text, False) & " с " & Me.MaskedTextBox2.Text & " по " & d8 & " (Приказ.ТрудОтпуск часть 2)" & ".doc",,,,,, False)
        '    Try
        '        IO.File.Copy("C:\Users\Public\Documents\Рик\" & TextBox5.Text & "-отп " & ФИОКорРук(TextBox1.Text, False) & " с " & Me.MaskedTextBox2.Text & " по " & d8 & " (Приказ.ТрудОтпуск часть 2)" & ".doc", OnePath & Отпуск.ComboBox2.Text & "\Приказ\" & Now.Year & "\Отпуск\" & TextBox5.Text & "-отп " & ФИОКорРук(TextBox1.Text, False) & " с " & Me.MaskedTextBox2.Text & " по " & d8 & " (Приказ.ТрудОтпуск часть 2)" & ".doc")
        '    Catch ex As Exception
        '        If MessageBox.Show("Приказ № " & TextBox5.Text & "-отп " & ФИОКорРук(TextBox1.Text, False) & " от " & Me.MaskedTextBox2.Text & " уже существует. Заменить старый документ новым?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then
        '            IO.File.Delete(OnePath & Отпуск.ComboBox2.Text & "\Приказ\" & Now.Year & "\Отпуск\" & TextBox5.Text & "-отп " & ФИОКорРук(TextBox1.Text, False) & " с " & Me.MaskedTextBox2.Text & " по " & d8 & " (Приказ.ТрудОтпуск часть 2)" & ".doc")
        '            IO.File.Copy("C:\Users\Public\Documents\Рик\" & TextBox5.Text & "-отп " & ФИОКорРук(TextBox1.Text, False) & " с " & Me.MaskedTextBox2.Text & " по " & d8 & " (Приказ.ТрудОтпуск часть 2)" & ".doc", OnePath & Отпуск.ComboBox2.Text & "\Приказ\" & Now.Year & "\Отпуск\" & TextBox5.Text & "-отп " & ФИОКорРук(TextBox1.Text, False) & " с " & Me.MaskedTextBox2.Text & " по " & d8 & " (Приказ.ТрудОтпуск часть 2)" & ".doc")
        '        End If
        '    End Try

        '    СохрПрик = OnePath & Отпуск.ComboBox2.Text & "\Приказ\" & Now.Year & "\Отпуск\" & TextBox5.Text & "-отп " & ФИОКорРук(TextBox1.Text, False) & " с " & Me.MaskedTextBox2.Text & " по " & d8 & " (Приказ.ТрудОтпуск часть 2)" & ".doc"
        'End If

        'oWordDoc.Close(True)
        'oWord.Quit(True)

        If MessageBox.Show("Заявление оформить?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
            hg = 1
            ОформлениеЗаявления()
        End If

        If hg = 0 Then
            If MessageBox.Show("Приказ оформлен! Распечатать? ", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.None) = DialogResult.OK Then
                ПечатьДоковFTP(massFTP3)
            End If
        Else
            If MessageBox.Show("Приказ и заявление оформлены! Распечатать? ", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.None) = DialogResult.OK Then
                ПечатьДоковFTP(massFTP3)

            End If
        End If



    End Sub
    Private Sub ОформлениеЗаявления()
        СборДаннОрганиз()
        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document
        oWord = CreateObject("Word.Application")
        oWord.Visible = False
        Dim n As String

        Начало("ZayavlenieTrudOtpusk.doc")
        oWordDoc = oWord.Documents.Add(firthtPath & "\ZayavlenieTrudOtpusk.doc")

        With oWordDoc.Bookmarks
            If dsorg.Rows(0).Item(18).ToString = "Индивидуальный предприниматель" Then
                .Item("ЗСО1").Range.Text = ДолжРодПадежФункц(dsorg.Rows(0).Item(18).ToString)
                .Item("ЗСО2").Range.Text = ФИОКорРук(dsorg.Rows(0).Item(30).ToString, False)
            Else
                .Item("ЗСО1").Range.Text = ДолжРодПадежФункц(dsorg.Rows(0).Item(18).ToString) & " " & ФормСобствКор(dsorg.Rows(0).Item(1).ToString) & " «" & Отпуск.ComboBox2.Text & "» "
                If dsorg.Rows(0).Item(31) = True Then
                    .Item("ЗСО2").Range.Text = ФИОКорРук(dsorg.Rows(0).Item(30).ToString, True)
                Else
                    .Item("ЗСО2").Range.Text = ФИОКорРук(dsorg.Rows(0).Item(30).ToString, False)
                End If
            End If

            If dssotr.Rows(0).Item(4).ToString = "" Or dssotr.Rows(0).Item(4).ToString = "-" Then
                .Item("ЗСО3").Range.Text = dssotr.Rows(0).Item(3).ToString
            Else
                .Item("ЗСО3").Range.Text = dssotr.Rows(0).Item(3).ToString & " " & разрядстрока(CType(dssotr.Rows(0).Item(4).ToString, Integer))
            End If
            .Item("ЗСО4").Range.Text = dssotr.Rows(0).Item(5).ToString & " " & dssotr.Rows(0).Item(6).ToString & " " & dssotr.Rows(0).Item(7).ToString

            If Отпуск.ДнОтпус = CType(TextBox2.Text, Integer) Then
                .Item("ЗСО5").Range.Text = "трудовой отпуск"
            Else
                .Item("ЗСО5").Range.Text = "часть трудового отпуска"
            End If

            If Int = 1 Then
                .Item("ЗСО6").Range.Text = TextBox2.Text
                .Item("ЗСО7").Range.Text = Пропись(CType(TextBox2.Text, Integer))
                .Item("ЗСО8").Range.Text = MaskedTextBox1.Text
                .Item("ЗСО9").Range.Text = d1
                .Item("ЗСО10").Range.Text = ДатаЗаявл(MaskedTextBox1.Text)
                n = dssotr.Rows(0).Item(0).ToString & " " & Strings.Left(dssotr.Rows(0).Item(1).ToString, 1) & "." & Strings.Left(dssotr.Rows(0).Item(2).ToString, 1) & ". c " & MaskedTextBox1.Text & " по " & d1 & " (Заявление.ТрудОтпуск)"
            Else
                .Item("ЗСО6").Range.Text = TextBox3.Text
                .Item("ЗСО7").Range.Text = Пропись(CType(TextBox3.Text, Integer))
                .Item("ЗСО8").Range.Text = MaskedTextBox2.Text
                .Item("ЗСО9").Range.Text = d8
                .Item("ЗСО10").Range.Text = ДатаЗаявл(MaskedTextBox2.Text)
                n = dssotr.Rows(0).Item(0).ToString & " " & Strings.Left(dssotr.Rows(0).Item(1).ToString, 1) & "." & Strings.Left(dssotr.Rows(0).Item(2).ToString, 1) & ". c " & MaskedTextBox2.Text & " по " & d8 & " (Заявление.ТрудОтпуск)"
            End If

            .Item("ЗСО11").Range.Text = Strings.Left(dssotr.Rows(0).Item(1).ToString, 1) & "." & Strings.Left(dssotr.Rows(0).Item(2).ToString, 1) & ". " & dssotr.Rows(0).Item(0).ToString
        End With

        Dim f As String = TextBox5.Text & " - отп "

        Dim Name As String = f & n & ".doc"
        Dim СохрЗак2 As New List(Of String)
        СохрЗак2.AddRange(New String() {Отпуск.ComboBox2.Text & "\Заявление\" & Now.Year & "\", Name})
        oWordDoc.SaveAs2(PathVremyanka & Name,,,,,, False)
        oWordDoc.Close(True)
        oWord.Quit(True)
        Конец(Отпуск.ComboBox2.Text & "\Заявление\" & Now.Year, Name, idcn, Отпуск.ComboBox2.Text, "\ZayavlenieTrudOtpusk.doc", "Заявление на труд.отпуск")
        massFTP3.Add(СохрЗак2)










        'oWordDoc.SaveAs2("C:\Users\Public\Documents\Рик\" & f & n & ".doc",,,,,, False)
        'СохрЗаявл = "C:\Users\Public\Documents\Рик\" & f & n & ".doc"
        'Try
        '    IO.File.Copy("C:\Users\Public\Documents\Рик\" & f & n & ".doc", OnePath & Отпуск.ComboBox2.Text & "\Приказ\" & Now.Year & "\Отпуск\" & f & n & ".doc")
        'Catch ex As Exception
        '    If MessageBox.Show("Заявление на труд.отпуск с сотрудником " & dssotr.Rows(0).Item(0).ToString & " существует. Заменить старый документ новым?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then
        '        Try
        '            IO.File.Delete(OnePath & Отпуск.ComboBox2.Text & "\Приказ\" & Now.Year & "\Отпуск\" & f & n & ".doc")
        '        Catch ex1 As Exception
        '            MessageBox.Show("Закройте файл!", Рик)
        '        End Try

        '        IO.File.Copy("C:\Users\Public\Documents\Рик\" & f & n & ".doc", OnePath & Отпуск.ComboBox2.Text & "\Приказ\" & Now.Year & "\Отпуск\" & f & n & ".doc")
        '    End If
        'End Try

        'oWordDoc.Close(True)
        'oWord.Quit(True)

    End Sub
    Private Function ДатаЗаявл(ByVal d As String) As String
        Dim den As Date
        den = CDate(d)
        den = den.AddDays(-4)
        Dim m As String = Format(den, "dddd")

        Select Case m
            Case "суббота"
                den = den.AddDays(-1)
            Case "воскресенье"
                den = den.AddDays(-2)
        End Select
        Dim g As String
        g = Strings.Left(den, 10)
        Return g
    End Function
    Private Sub СборДаннОрганиз()


        Dim dh = dtSotrudnikiAll.Select("ФИОСборное='" & TextBox1.Text & "' and НазвОрганиз='" & Отпуск.ComboBox2.Text & "'")

        'Dim strsql3 As String = "SELECT КодСотрудники FROM Сотрудники WHERE ФИОСборное='" & TextBox1.Text & "' and НазвОрганиз='" & Отпуск.ComboBox2.Text & "'"
        'Dim dh As DataTable = Selects(strsql3)
        Dim list As New Dictionary(Of String, Object)
        list.Add("@КодСотрудники", CType(dh(0).Item("КодСотрудники"), Integer))
        list.Add("@НазвОрг", Отпуск.ComboBox2.Text)

        dssotr = Selects(StrSql:= "SELECT Сотрудники.Фамилия, Сотрудники.Имя, Сотрудники.Отчество, Штатное.Должность, Штатное.Разряд,
Сотрудники.ФамилияДляЗаявления,Сотрудники.ИмяДляЗаявления,Сотрудники.ОтчествоДляЗаявления
FROM Сотрудники INNER JOIN Штатное ON Сотрудники.КодСотрудники = Штатное.ИДСотр
WHERE Сотрудники.КодСотрудники=@КодСотрудники", list)

        'dsorg = From x In dtClientAll.TableName Where x.Item("НазвОрг") = Отпуск.ComboBox2.Text Select x

        dsorg = Selects(StrSql:="SELECT * FROM Клиент WHERE НазвОрг=@НазвОрг", list)


    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub MaskedTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox2.Focus()
        End If
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            MaskedTextBox2.Focus()
        End If
    End Sub

    Private Sub MaskedTextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            TextBox3.Focus()
        End If
    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            TextBox4.Focus()
        End If
    End Sub
    Private Sub Приказы()
        Dim gth3 As String

        'Try

        '    Files2 = (IO.Directory.GetFiles(OnePath & Отпуск.ComboBox2.Text & "\Приказ\" & ComboBox1.Text & "\Отпуск\", "*.doc", IO.SearchOption.TopDirectoryOnly))
        '    For n As Integer = 0 To Files2.Length - 1
        '        gth3 = ""
        '        gth3 = IO.Path.GetFileName(Files2(n))
        '        Files2(n) = gth3
        '    Next

        Dim dt = listFluentFTP("/" & Отпуск.ComboBox2.Text & "/Приказ/" & ComboBox1.Text & "/")
        ComboBox2.Items.Clear()
        For Each v In dt
            ComboBox2.Items.Add(v.ToString)
        Next

        'Catch ex As Exception
        'MessageBox.Show("В " & ComboBox1.Text & " году нет приказов на отпуск!", Рик)
        'End Try
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Приказы()
    End Sub

    Private Sub ОтпускНачало_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim ut() As Object = {Now.Year - 2, Now.Year - 1, Now.Year}
        ComboBox1.Items.AddRange(ut)
        MaskedTextBox3.Text = Now.ToShortDateString
        ComboBox1.Text = Now.Year
        'Приказы()

        If ComboBox2.Items.Count = 0 Then
            Throw New System.Exception("Нет файлов в папке!")
        End If

        'ComboBox2.Text = ComboBox2.Items.last
        Dim list As New Dictionary(Of String, Object)
        list.Add("@НазвОрганиз", NameOrg)
        list.Add("@ФИО", TextBox1.Text)




        '        Dim df = Selects(StrSql:="SELECT Сотрудники.КодСотрудники 
        'FROM Сотрудники WHERE НазвОрганиз =@НазвОрганиз and ФИОСборное=@ФИОСборное", list)
        Dim df = dtSotrudnikiAll.Select("НазвОрганиз ='" & NameOrg & "' and ФИОСборное='" & TextBox1.Text & "'")

        Try
            idcn = df(0).Item("КодСотрудники")
        Catch ex As Exception
            MessageBox.Show("Данные по сотруднику не найдены! Возможно был удален из базы", Рик)

            Updates(stroka:="DELETE FROM ОтпускСотрудники WHERE ФИО='" & TextBox1.Text & "'", list, "ОтпускСотрудники")

            Отпуск.grcellclick()
            Отпуск.grid3activ()
            Me.Close()
        End Try








    End Sub

    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            MaskedTextBox3.Focus()

            Dim pl As String
            If TextBox5.Text <> "" Then
                Dim i As Integer = CInt(TextBox5.Text)
                Select Case i

                    Case < 10
                        pl = Str(i)
                        TextBox5.Text = "00" & i

                    Case 10 To 99
                        pl = Str(i)
                        TextBox5.Text = "0" & i
                End Select
            End If
        End If
    End Sub

    Private Sub MaskedTextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Button1.Focus()
        End If
    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            TextBox5.Focus()
        End If
    End Sub
End Class