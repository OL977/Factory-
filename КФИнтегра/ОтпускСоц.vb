Option Explicit On
Imports System.Data.OleDb
Imports System.Threading
Public Class ОтпускСоц1
    Dim idsotr As Integer
    Dim dssotr As DataTable
    Dim dsorg As DataRow()
    Dim hg As Integer = 0
    'Dim СохрЗак, СохрЗаявл As String
    Dim massFTP3 As New ArrayList()
    Dim massFTP As New ArrayList()
    Dim СохрЗак2 As New List(Of String)
    Dim СохрЗак As New List(Of String)

    Private Sub ОтпускСоц1_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Me.ComboBox1.AutoCompleteCustomSource.Clear()
        Me.ComboBox1.Items.Clear()
        For Each r As DataRow In СписокКлиентовОсновной.Rows
            Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox1.Items.Add(r(0).ToString)
        Next

        MaskedTextBox1.Text = Now.ToShortDateString
        MaskedTextBox2.Text = Now.ToShortDateString
        MaskedTextBox3.Text = Now.ToShortDateString

        Dim d() As String = {"без сохранения заработной платы", "с сохранением заработной платы"}
        ComboBox2.Items.AddRange(d)
        ComboBox2.Text = "без сохранения заработной платы"

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ComboBox19.Text = ""
        Com1sel()

    End Sub
    Private Sub ClAll()

        ComboBox2.Text = ""
        ComboBox3.Text = ""
        RichTextBox1.Text = ""
        TextBox57.Text = ""
        TextBox1.Text = ""
        ListBox1.Items.Clear()
        MaskedTextBox1.Text = Now.ToShortDateString
        MaskedTextBox2.Text = Now.ToShortDateString
        MaskedTextBox3.Text = Now.ToShortDateString

    End Sub
    Private Sub Com1sel()
        ClAll()

        Dim dsGeneral = dtSotrudnikiAll.Select("НазвОрганиз='" & ComboBox1.Text & "'")

        'Dim strsql As String = "SELECT ФИОСборное,КодСотрудники FROM Сотрудники WHERE НазвОрганиз='" & ComboBox1.Text & "' ORDER BY ФИОСборное "
        'Dim dsGeneral As DataTable = Selects(strsql)

        Me.ComboBox19.AutoCompleteCustomSource.Clear()
        Me.ComboBox19.Items.Clear()
        ComboBox26.Items.Clear()

        For Each r As DataRow In dsGeneral
            Me.ComboBox19.AutoCompleteCustomSource.Add(r.Item("ФИОСборное").ToString())
            Me.ComboBox19.Items.Add(r.Item("ФИОСборное").ToString)
            Me.ComboBox26.Items.Add(r.Item("КодСотрудники").ToString)
        Next

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
        '    MessageBox.Show("У данной организации нет папки приказ и отпуск!", Рик)
        '    Exit Sub
        'End Try

        Dim list = listFluentFTP(ComboBox1.Text & "\Приказ\")

        ComboBox3.Items.Clear()

        For Each f In list
            ComboBox3.Items.Add(f.ToString)
        Next


    End Sub

    Private Sub ComboBox19_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox19.SelectedIndexChanged

        Label96.Text = ComboBox26.Items.Item(ComboBox19.SelectedIndex)
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

        'Dim Files2(), gth3 As String

        'Try

        '    Files2 = (IO.Directory.GetFiles(OnePath & ComboBox1.Text & "\Приказ\" & ComboBox3.Text & "\Отпуск", "*.doc", IO.SearchOption.TopDirectoryOnly))
        '    For n As Integer = 0 To Files2.Length - 1
        '        gth3 = ""
        '        gth3 = IO.Path.GetFileName(Files2(n))
        '        Files2(n) = gth3
        '    Next
        '    ListBox1.Items.Clear()
        '    ListBox1.Items.AddRange(Files2)
        'Catch ex As Exception
        '    MessageBox.Show("В " & ComboBox3.Text & " году нет приказов на отпуск!", Рик)
        'End Try

        Dim list = listFluentFTP(ComboBox1.Text & "\Приказ\" & ComboBox3.Text & "\")
        ListBox1.Items.Clear()
        For Each f In list
            ListBox1.Items.Add(f.ToString)
        Next

    End Sub

    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick
        If ListBox1.SelectedIndex = -1 Then
            MessageBox.Show("Выберите документ для просмотра!", Рик, MessageBoxButtons.OK)
            Exit Sub
        End If
        ОткрытиеФайлаБезПути(ListBox1.SelectedItem)
    End Sub



    Private Sub TextBox57_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox57.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            MaskedTextBox1.Focus()
        End If
    End Sub

    Private Sub MaskedTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            ComboBox2.Focus()
        End If
    End Sub

    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            TextBox1.Focus()
        End If
    End Sub
    Private Sub расчет()


        Dim d As Date = CDate(MaskedTextBox2.Text)
        d = d.AddDays(CType(TextBox1.Text, Integer) - 1)
        MaskedTextBox3.Text = d.ToShortDateString
    End Sub
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            MaskedTextBox2.Focus()
            расчет()
        End If
    End Sub

    Private Sub MaskedTextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            расчет()
            Button1.Focus()
        End If
    End Sub

    Private Sub MaskedTextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Button1.Focus()
        End If
    End Sub
    Private Function пров()
        If ComboBox1.Text = "" Or ComboBox19.Text = "" Then
            MessageBox.Show("Выберите организацию или сотрудника!", Рик)
            Return 1
        End If

        If RichTextBox1.Text = "" Then
            MessageBox.Show("Выберите номер приказа!", Рик)
            Return 1
        End If

        If MaskedTextBox1.MaskCompleted = False Then
            MessageBox.Show("Введите правильно дату приказа!", Рик)
            Return 1
        End If

        If ComboBox2.Text = "" Then
            MessageBox.Show("Выберите условия оплаты!", Рик)
            Return 1
        End If

        If TextBox1.Text = "" Then
            MessageBox.Show("Выберите количество дней отпуска!", Рик)
            Return 1
        End If

        If MaskedTextBox2.MaskCompleted = False Then
            MessageBox.Show("Введите правильно дату начала отпуска!", Рик)
            Return 1
        End If

        Return 0
    End Function
    Private Sub СборДаннОрганиз()

        Dim list As New Dictionary(Of String, Object)
        list.Add("@КодСотрудники", idsotr)

        dssotr = Selects(StrSql:="SELECT Сотрудники.Фамилия, Сотрудники.Имя, Сотрудники.Отчество, Штатное.Должность, Штатное.Разряд,
Сотрудники.ФамилияДляЗаявления,Сотрудники.ИмяДляЗаявления,Сотрудники.ОтчествоДляЗаявления
FROM Сотрудники INNER JOIN Штатное ON Сотрудники.КодСотрудники = Штатное.ИДСотр
WHERE Сотрудники.КодСотрудники=@КодСотрудники", list)


        'Dim strsql1 As String = "SELECT * FROM Клиент WHERE НазвОрг='" & ComboBox1.Text & "'"
        dsorg = dtClientAll.Select("НазвОрг='" & ComboBox1.Text & "'")

    End Sub

    Private Sub Доки()
        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document
        oWord = CreateObject("Word.Application")
        oWord.Visible = False

        'Try
        '    IO.File.Copy(OnePath & "\ОБЩДОКИ\General\PrikazNaOtpuskSoc.doc", "C:\Users\Public\Documents\Рик\PrikazNaOtpuskSoc.doc")
        'Catch ex As Exception
        '    'If "Zayavlenie.doc" <> "" Then IO.File.Delete("C:\Users\Public\Documents\Рик\Zayavlenie.doc")
        '    If Not IO.Directory.Exists("c:\Users\Public\Documents\Рик") Then
        '        IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рик")
        '    End If
        '    IO.File.Copy(OnePath & "\ОБЩДОКИ\General\PrikazNaOtpuskSoc.doc", "C:\Users\Public\Documents\Рик\PrikazNaOtpuskSoc.doc")
        'End Try

        Начало("PrikazNaOtpuskSoc.doc")
        oWordDoc = oWord.Documents.Add(firthtPath & "\PrikazNaOtpuskSoc.doc")
        Dim фсотз As String
        Try
            фсотз = dssotr.Rows(0).Item(5).ToString & " " & dssotr.Rows(0).Item(6).ToString & " " & dssotr.Rows(0).Item(7).ToString
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Exit Sub
        End Try



        Dim f As String
        With oWordDoc.Bookmarks
            .Item("ПОс1").Range.Text = MaskedTextBox1.Text
            If TextBox57.Text <> "" Then
                .Item("ПОс2").Range.Text = RichTextBox1.Text & " - отп - " & TextBox57.Text
                f = RichTextBox1.Text & " - отп -" & TextBox57.Text
            Else
                .Item("ПОс2").Range.Text = RichTextBox1.Text & " - отп"
                f = RichTextBox1.Text & " - отп "
            End If
            .Item("ПОс3").Range.Text = ComboBox2.Text
            .Item("ПОс4").Range.Text = ComboBox2.Text
            .Item("ПОс5").Range.Text = InputName1(ComboBox19.Text, "ОтпускСоц")
            If dssotr.Rows(0).Item(4).ToString <> "" And Not dssotr.Rows(0).Item(4).ToString = "-" Then
                .Item("ПОс6").Range.Text = Strings.LCase(ДолжРодПадежФункц(dssotr.Rows(0).Item(3).ToString)) & " " & разрядстрока(CType(dssotr.Rows(0).Item(4).ToString, Integer))
            Else
                .Item("ПОс6").Range.Text = Strings.LCase(ДолжРодПадежФункц(dssotr.Rows(0).Item(3).ToString))
            End If
            .Item("ПОс7").Range.Text = TextBox1.Text
            .Item("ПОс8").Range.Text = КалендарДней(CType(TextBox1.Text, Integer))
            .Item("ПОс10").Range.Text = MaskedTextBox2.Text
            .Item("ПОс11").Range.Text = MaskedTextBox3.Text

            .Item("ПОс12").Range.Text = ФИОКорРук(фсотз, False)

            .Item("ПОс13").Range.Text = dsorg(0).Item(18).ToString
            If dsorg(0).Item(31) = True Then
                .Item("ПОс14").Range.Text = ФИОКорРук(dsorg(0).Item(19).ToString, True)
            Else
                .Item("ПОс14").Range.Text = ФИОКорРук(dsorg(0).Item(19).ToString, False)
            End If
            .Item("ПОс15").Range.Text = ФИОКорРук(ComboBox19.Text, False)
            .Item("ПОс16").Range.Text = dsorg(0).Item(1).ToString
            If dsorg(0).Item(1).ToString = "Индивидуальный предприниматель" Then
                .Item("ПОс17").Range.Text = dsorg(0).Item(0).ToString
            Else
                .Item("ПОс17").Range.Text = "«" & dsorg(0).Item(0).ToString & "»"
            End If

            .Item("ПОс18").Range.Text = dsorg(0).Item(4).ToString
            .Item("ПОс19").Range.Text = dsorg(0).Item(2).ToString
            .Item("ПОс20").Range.Text = dsorg(0).Item(14).ToString
            .Item("ПОс21").Range.Text = dsorg(0).Item(12).ToString
            .Item("ПОс22").Range.Text = dsorg(0).Item(11).ToString
            .Item("ПОс23").Range.Text = dsorg(0).Item(8).ToString
            .Item("ПОс24").Range.Text = dsorg(0).Item(6).ToString
        End With

        Dim Name As String = f & " " & dssotr.Rows(0).Item(0).ToString & " " & Strings.Left(dssotr.Rows(0).Item(1).ToString, 1) & "." & Strings.Left(dssotr.Rows(0).Item(2).ToString, 1) & ". c " & MaskedTextBox2.Text & " по " & MaskedTextBox3.Text & " (Приказ.СоцОтпуск)" & ".doc"
        СохрЗак.AddRange(New String() {ComboBox1.Text & "\Приказ\" & Now.Year & "\", Name})
        oWordDoc.SaveAs2(PathVremyanka & Name,,,,,, False)
        oWordDoc.Close(True)
        oWord.Quit(True)
        Конец(ComboBox1.Text & "\Приказ\" & Now.Year, Name, idsotr, ComboBox1.Text, "\PrikazNaOtpuskSoc.doc", "Приказ.СоцОтпуск")
        massFTP3.Add(СохрЗак)
        massFTP.Add(СохрЗак)




        'If Not IO.Directory.Exists(OnePath & ComboBox1.Text & "\Приказ\" & Now.Year & "\Отпуск") Then
        '    IO.Directory.CreateDirectory(OnePath & ComboBox1.Text & "\Приказ\" & Now.Year & "\Отпуск")
        'End If

        'oWordDoc.SaveAs2("C:\Users\Public\Documents\Рик\" & f & " " & dssotr.Rows(0).Item(0).ToString & " " & Strings.Left(dssotr.Rows(0).Item(1).ToString, 1) & "." & Strings.Left(dssotr.Rows(0).Item(2).ToString, 1) & ". c " & MaskedTextBox2.Text & " по " & MaskedTextBox3.Text & " (Приказ.СоцОтпуск)" & ".doc",,,,,, False)
        ''СохрЗак = "C:\Users\Public\Documents\Рик\" & f & " " & dssotr.Rows(0).Item(0).ToString & " " & Strings.Left(dssotr.Rows(0).Item(1).ToString, 1) & "." & Strings.Left(dssotr.Rows(0).Item(2).ToString, 1) & ". c " & MaskedTextBox2.Text & " по " & MaskedTextBox3.Text & " (Приказ.СоцОтпуск)" & ".doc"
        '''oWordDoc.SaveAs2("U: \Офис\Финансовый\6. Бух.услуги\Кадры\" & Клиент & "\Заявление\" & Год & "\" & Заявление(9) & " (заявление)" & ".doc",,,,,, False)
        'Try
        '    IO.File.Copy("C:\Users\Public\Documents\Рик\" & f & " " & dssotr.Rows(0).Item(0).ToString & " " & Strings.Left(dssotr.Rows(0).Item(1).ToString, 1) & "." & Strings.Left(dssotr.Rows(0).Item(2).ToString, 1) & ". c " & MaskedTextBox2.Text & " по " & MaskedTextBox3.Text & " (Приказ.СоцОтпуск)" & ".doc", OnePath & ComboBox1.Text & "\Приказ\" & Now.Year & "\Отпуск\" & f & " " & dssotr.Rows(0).Item(0).ToString & " " & Strings.Left(dssotr.Rows(0).Item(1).ToString, 1) & "." & Strings.Left(dssotr.Rows(0).Item(2).ToString, 1) & ". c " & MaskedTextBox2.Text & " по " & MaskedTextBox3.Text & " (Приказ.СоцОтпуск)" & ".doc")
        'Catch ex As Exception
        '    If MessageBox.Show("Приказ на соц.отпуск с сотрудником " & dssotr.Rows(0).Item(0).ToString & " существует. Заменить старый документ новым?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then
        '        Try
        '            IO.File.Delete(OnePath & ComboBox1.Text & "\Приказ\" & Now.Year & "\Отпуск\" & f & " " & dssotr.Rows(0).Item(0).ToString & " " & Strings.Left(dssotr.Rows(0).Item(1).ToString, 1) & "." & Strings.Left(dssotr.Rows(0).Item(2).ToString, 1) & ". c " & MaskedTextBox2.Text & " по " & MaskedTextBox3.Text & " (Приказ.СоцОтпуск)" & ".doc")
        '        Catch ex1 As Exception
        '            MessageBox.Show("Закройте файл!", Рик)
        '        End Try


        '        IO.File.Copy("C:\Users\Public\Documents\Рик\" & f & " " & dssotr.Rows(0).Item(0).ToString & " " & Strings.Left(dssotr.Rows(0).Item(1).ToString, 1) & "." & Strings.Left(dssotr.Rows(0).Item(2).ToString, 1) & ". c " & MaskedTextBox2.Text & " по " & MaskedTextBox3.Text & " (Приказ.СоцОтпуск)" & ".doc", OnePath & ComboBox1.Text & "\Приказ\" & Now.Year & "\Отпуск\" & f & " " & dssotr.Rows(0).Item(0).ToString & " " & Strings.Left(dssotr.Rows(0).Item(1).ToString, 1) & "." & Strings.Left(dssotr.Rows(0).Item(2).ToString, 1) & ". c " & MaskedTextBox2.Text & " по " & MaskedTextBox3.Text & " (Приказ.СоцОтпуск)" & ".doc")
        '    End If
        'End Try
        'СохрЗак = OnePath & ComboBox1.Text & "\Приказ\" & Now.Year & "\Отпуск\" & f & " " & dssotr.Rows(0).Item(0).ToString & " " & Strings.Left(dssotr.Rows(0).Item(1).ToString, 1) & "." & Strings.Left(dssotr.Rows(0).Item(2).ToString, 1) & ". c " & MaskedTextBox2.Text & " по " & MaskedTextBox3.Text & " (Приказ.СоцОтпуск)" & ".doc"

        'oWordDoc.Close(True)
        'oWord.Quit(True)

        If MessageBox.Show("Заявление оформить?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
            hg = 1
            ОформлениеЗаявления()
        End If

    End Sub
    Private Sub ОформлениеЗаявления()

        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document

        oWord = CreateObject("Word.Application")
        oWord.Visible = False

        Начало("ZayavlenieSocOtpusk.doc")
        oWordDoc = oWord.Documents.Add(firthtPath & "\ZayavlenieSocOtpusk.doc")

        With oWordDoc.Bookmarks
            If dsorg(0).Item(18).ToString = "Индивидуальный предприниматель" Then
                .Item("ЗСО1").Range.Text = ДолжРодПадежФункц(dsorg(0).Item(18).ToString)
                .Item("ЗСО2").Range.Text = ФИОКорРук(dsorg(0).Item(30).ToString, False)
            Else
                .Item("ЗСО1").Range.Text = ДолжРодПадежФункц(dsorg(0).Item(18).ToString) & " " & ФормСобствКор(dsorg(0).Item(1).ToString) & " «" & ComboBox1.Text & "» "
                If dsorg(0).Item(31) = True Then
                    .Item("ЗСО2").Range.Text = ФИОКорРук(dsorg(0).Item(30).ToString, True)
                Else
                    .Item("ЗСО2").Range.Text = ФИОКорРук(dsorg(0).Item(30).ToString, False)
                End If
            End If

            If dssotr.Rows(0).Item(4).ToString = "" Or dssotr.Rows(0).Item(4).ToString = "-" Then
                .Item("ЗСО3").Range.Text = dssotr.Rows(0).Item(3).ToString
            Else
                .Item("ЗСО3").Range.Text = dssotr.Rows(0).Item(3).ToString & " " & разрядстрока(CType(dssotr.Rows(0).Item(4).ToString, Integer))
            End If
            .Item("ЗСО4").Range.Text = dssotr.Rows(0).Item(5).ToString & " " & dssotr.Rows(0).Item(6).ToString & " " & dssotr.Rows(0).Item(7).ToString
            .Item("ЗСО5").Range.Text = ComboBox2.Text
            .Item("ЗСО6").Range.Text = TextBox1.Text
            .Item("ЗСО7").Range.Text = Пропись(CType(TextBox1.Text, Integer))
            .Item("ЗСО8").Range.Text = MaskedTextBox2.Text
            .Item("ЗСО9").Range.Text = MaskedTextBox3.Text
            .Item("ЗСО10").Range.Text = MaskedTextBox1.Text
            .Item("ЗСО11").Range.Text = Strings.Left(dssotr.Rows(0).Item(1).ToString, 1) & "." & Strings.Left(dssotr.Rows(0).Item(2).ToString, 1) & ". " & dssotr.Rows(0).Item(0).ToString
        End With
        Dim f As String
        If TextBox57.Text <> "" Then
            f = RichTextBox1.Text & " - отп -" & TextBox57.Text
        Else
            f = RichTextBox1.Text & " - отп "
        End If

        Dim Name As String = f & " " & dssotr.Rows(0).Item(0).ToString & " " & Strings.Left(dssotr.Rows(0).Item(1).ToString, 1) & "." & Strings.Left(dssotr.Rows(0).Item(2).ToString, 1) & ". c " & MaskedTextBox2.Text & " по " & MaskedTextBox3.Text & " (Заявление.СоцОтпуск)" & ".doc"

        СохрЗак2.AddRange(New String() {ComboBox1.Text & "\Заявление\" & Now.Year & "\", Name})
        oWordDoc.SaveAs2(PathVremyanka & Name,,,,,, False)
        oWordDoc.Close(True)
        oWord.Quit(True)
        Конец(ComboBox1.Text & "\Заявление\" & Now.Year, Name, idsotr, ComboBox1.Text, "\ZayavlenieSocOtpusk.doc", "Заявление.СоцОтпуск")
        massFTP3.Add(СохрЗак2)






        'oWordDoc.SaveAs2("C:\Users\Public\Documents\Рик\" & f & " " & dssotr.Rows(0).Item(0).ToString & " " & Strings.Left(dssotr.Rows(0).Item(1).ToString, 1) & "." & Strings.Left(dssotr.Rows(0).Item(2).ToString, 1) & ". c " & MaskedTextBox2.Text & " по " & MaskedTextBox3.Text & " (Заявление.СоцОтпуск)" & ".doc",,,,,, False)
        'СохрЗаявл = "C:\Users\Public\Documents\Рик\" & f & " " & dssotr.Rows(0).Item(0).ToString & " " & Strings.Left(dssotr.Rows(0).Item(1).ToString, 1) & "." & Strings.Left(dssotr.Rows(0).Item(2).ToString, 1) & ". c " & MaskedTextBox2.Text & " по " & MaskedTextBox3.Text & " (Заявление.СоцОтпуск)" & ".doc"
        ''oWordDoc.SaveAs2("U: \Офис\Финансовый\6. Бух.услуги\Кадры\" & Клиент & "\Заявление\" & Год & "\" & Заявление(9) & " (заявление)" & ".doc",,,,,, False)
        'Try
        '    IO.File.Copy("C:\Users\Public\Documents\Рик\" & f & " " & dssotr.Rows(0).Item(0).ToString & " " & Strings.Left(dssotr.Rows(0).Item(1).ToString, 1) & "." & Strings.Left(dssotr.Rows(0).Item(2).ToString, 1) & ". c " & MaskedTextBox2.Text & " по " & MaskedTextBox3.Text & " (Заявление.СоцОтпуск)" & ".doc", OnePath & ComboBox1.Text & "\Приказ\" & Now.Year & "\Отпуск\" & f & " " & dssotr.Rows(0).Item(0).ToString & " " & Strings.Left(dssotr.Rows(0).Item(1).ToString, 1) & "." & Strings.Left(dssotr.Rows(0).Item(2).ToString, 1) & ". c " & MaskedTextBox2.Text & " по " & MaskedTextBox3.Text & " (Заявление.СоцОтпуск)" & ".doc")
        'Catch ex As Exception
        '    If MessageBox.Show("Заявление на соц.отпуск с сотрудником " & dssotr.Rows(0).Item(0).ToString & " существует. Заменить старый документ новым?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then
        '        Try
        '            IO.File.Delete(OnePath & ComboBox1.Text & "\Приказ\" & Now.Year & "\Отпуск\" & f & " " & dssotr.Rows(0).Item(0).ToString & " " & Strings.Left(dssotr.Rows(0).Item(1).ToString, 1) & "." & Strings.Left(dssotr.Rows(0).Item(2).ToString, 1) & ". c " & MaskedTextBox2.Text & " по " & MaskedTextBox3.Text & " (Заявление.СоцОтпуск)" & ".doc")
        '        Catch ex1 As Exception
        '            MessageBox.Show("Закройте файл!", Рик)
        '        End Try


        '        IO.File.Copy("C:\Users\Public\Documents\Рик\" & f & " " & dssotr.Rows(0).Item(0).ToString & " " & Strings.Left(dssotr.Rows(0).Item(1).ToString, 1) & "." & Strings.Left(dssotr.Rows(0).Item(2).ToString, 1) & ". c " & MaskedTextBox2.Text & " по " & MaskedTextBox3.Text & " (Заявление.СоцОтпуск)" & ".doc", OnePath & ComboBox1.Text & "\Приказ\" & Now.Year & "\Отпуск\" & f & " " & dssotr.Rows(0).Item(0).ToString & " " & Strings.Left(dssotr.Rows(0).Item(1).ToString, 1) & "." & Strings.Left(dssotr.Rows(0).Item(2).ToString, 1) & ". c " & MaskedTextBox2.Text & " по " & MaskedTextBox3.Text & " (Заявление.СоцОтпуск)" & ".doc")
        '    End If
        'End Try

        'oWordDoc.Close(True)
        'oWord.Quit(True)


    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If пров() = 1 Then Exit Sub
        idsotr = CType(Label96.Text, Integer)
        Cursor = Cursors.WaitCursor

        СохрВБазу()



        СборДаннОрганиз()
        Доки()
        Статистика1(ComboBox19.Text, "Отправка сотрудника в соц.отпуск", ComboBox1.Text)
        If hg = 0 Then
            If MessageBox.Show("Приказ оформлен! Распечатать? ", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.None) = DialogResult.OK Then
                ПечатьДоковFTP(massFTP)
                'ПечатьДоковКол(СохрЗак, 2)
            End If
        Else
            If MessageBox.Show("Приказ и заявление оформлены! Распечатать? ", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.None) = DialogResult.OK Then
                ПечатьДоковFTP(massFTP3)
            End If
        End If
        hg = 0
        Com1sel()
        Cursor = Cursors.Default
    End Sub
    Private Sub СохрВБазу()
        Dim f As String
        If TextBox57.Text <> "" Then
            f = RichTextBox1.Text & " - отп - " & TextBox57.Text
        Else
            f = RichTextBox1.Text & " - отп "
        End If

        Dim idsot As Integer = CType(Label96.Text, Integer)

        Dim ds As DataRow()
        Dim list As New Dictionary(Of String, Object)
        Dim l As Object
        'If dtOtpuskSocAll.Rows.Count > 0 Then
        ds = dtOtpuskSocAll.Select("Организация='" & ComboBox1.Text & "' AND Сотрудник='" & ComboBox19.Text & "' AND Приказ='" & f & "'AND
ДатаПриказа='" & MaskedTextBox1.Text & "' AND ИДСотр=" & idsot & "")
        '        Dim strsql As String = "SELECT * FROM ОтпускСоц WHERE Организация='" & ComboBox1.Text & "', Сотрудник='" & ComboBox19.Text & "', Приказ='" & f & "',
        'ДатаПриказа= '" & MaskedTextBox1.Text & "', УсловияОплаты= '" & ComboBox2.Text & "', КоличествоДней='" & TextBox1.Text & "', ПериодС='" & MaskedTextBox2.Text & "',
        'ПериодПо='" & MaskedTextBox3.Text & "'"
        '        Dim ds As DataTable = Selects(strsql)

        'End If



        If ds.Length = 0 Then
            l = 0
            list.Add("@Код", l)
            Updates(stroka:="INSERT INTO ОтпускСоц(Организация,Сотрудник,Приказ,ДатаПриказа,УсловияОплаты,КоличествоДней,ПериодС,ПериодПо,ИДСотр)
VALUES('" & ComboBox1.Text & "','" & ComboBox19.Text & "','" & f & "','" & MaskedTextBox1.Text & "','" & ComboBox2.Text & "','" & TextBox1.Text & "','" & MaskedTextBox2.Text & "',
'" & MaskedTextBox3.Text & "'," & idsot & ")", list, "ОтпускСоц")

        Else
            l = ds(0).Item("Код")
            list.Add("@Код", l)
            Updates(stroka:="UPDATE ОтпускСоц SET Организация='" & ComboBox1.Text & "', Сотрудник='" & ComboBox19.Text & "', Приказ='" & f & "',
ДатаПриказа= '" & MaskedTextBox1.Text & "', УсловияОплаты= '" & ComboBox2.Text & "', КоличествоДней='" & TextBox1.Text & "', ПериодС='" & MaskedTextBox2.Text & "',
ПериодПо='" & MaskedTextBox3.Text & "' WHERE Код=@Код", list, "ОтпускСоц")

        End If
    End Sub
    Private Sub RichTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox57.Focus()

            Dim pl As String
            If RichTextBox1.Text <> "" Then
                Try
                    Dim i As Integer = CInt(RichTextBox1.Text)
                    Select Case i
                        Case < 10
                            pl = Str(i)
                            RichTextBox1.Text = "00" & i

                        Case 10 To 99
                            pl = Str(i)
                            RichTextBox1.Text = "0" & i
                    End Select
                Catch ex As Exception
                    RichTextBox1.Text = Replace(RichTextBox1.Text, "/", ".")
                    RichTextBox1.Text = Replace(RichTextBox1.Text, "\", ".")
                    RichTextBox1.Text = "б.н"
                End Try
            Else
                RichTextBox1.Text = "б.н"
            End If

        End If
    End Sub
End Class