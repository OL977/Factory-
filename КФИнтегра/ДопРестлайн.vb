Option Explicit On
Imports System.Data.OleDb


Public Class ДопРестлайн

    Public ds As DataTable
    Dim StrSql, ФИОРукРодПад, ФИОПолнРук, ФИОрКОР, ФИОСотрКор, ФормСобсКоротко, ДолжДирСОконч,
        ДолжОконСотр, Разряд, inp, ДатаКонтр, НомерКонтр, ФормСобПолн, ДолжРук, ОснДейств, ДатОкон, ФИОПолноеСотр As String
    Dim cl, IDСотр, errs As Integer

    Dim ТарифСт, ПроцОкл, РасчДолжОкл As Double





    Private Sub ДопРестлайн_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Me.MdiParent = MDIParent1
        'Me.WindowState = FormWindowState.Maximized
        'Соед(0)

        StrSql = "SELECT Сотрудники.ФИОСборное FROM Сотрудники WHERE НазвОрганиз='Рестлайн' ORDER BY ФИОСборное"
        ds = Selects(StrSql)



        ListBox1.Items.Clear()

        For Each r As DataRow In ds.Rows
            Me.ListBox1.Items.Add(r(0).ToString)
        Next

    End Sub

    Private Sub MaskedTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ListBox1.Focus()
        End If
    End Sub
    Private Sub Очист()


        ListBox1.Items.Clear()

        MaskedTextBox1.Text = ""

    End Sub
    Private Sub refreshes()
        Чист()

        'StrSql = "SELECT ФИОСборное FROM Сотрудники WHERE НазвОрганиз='" & ComboBox1.Text & "' ORDER BY ФИОСборное"
        'ds = Selects(StrSql)

        For Each r As DataRow In ds.Rows
            ListBox1.Items.Add(r.Item(0).ToString())
            'ListBox2.Items.Add(r.Item(0).ToString())
        Next
    End Sub
    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs)
        Очист()
        refreshes()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) 

        ListBox1.Items.Clear()
    End Sub
    Private Function ПровЗаполн()

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
        ds.Clear()
    End Sub

    Private Sub ДанСотр(ByVal IDсот As Integer)

        Чист() ' выбираем данные по сотруднику
        StrSql = "SELECT Штатное.ТарифнаяСтавка, Штатное.ПовышОклПроц, Штатное.РасчДолжностнОклад, ДогСотрудн.Контракт, ДогСотрудн.ДатаКонтракта, Сотрудники.ФИОСборное
FROM (Сотрудники INNER JOIN ДогСотрудн ON Сотрудники.КодСотрудники = ДогСотрудн.IDСотр) INNER JOIN Штатное ON Сотрудники.КодСотрудники = Штатное.ИДСотр
WHERE Сотрудники.КодСотрудники=" & IDСотр & ""
        ds = Selects(StrSql)
        Dim df As String = ds.Rows(0).Item(5).ToString

        ФИОСотрКор = ФИОКорРук(ds.Rows(0).Item(5).ToString, False)

        ДатаКонтр = Strings.Left(ds.Rows(0).Item(4), 10)
        НомерКонтр = ds.Rows(0).Item(3).ToString
        ФИОПолноеСотр = ds.Rows(0).Item(5).ToString
        ТарифСт = ds.Rows(0).Item(0)
        ПроцОкл = ds.Rows(0).Item(1)
        РасчДолжОкл = ds.Rows(0).Item(2)

    End Sub
    Private Sub доки()

        Try
            IO.Directory.Delete("c:\Users\Public\Documents\Рик", True)
        Catch ex As Exception

        End Try





        Dim СохрЗак(ListBox1.SelectedItems.Count - 1) As String

        For i = 0 To ListBox1.SelectedItems.Count - 1

            Чист()
            StrSql = "SELECT КодСотрудники FROM Сотрудники WHERE Сотрудники.НазвОрганиз='Рестлайн' AND Сотрудники.ФИОСборное='" & ListBox1.SelectedItems(i) & "'"
            ds = Selects(StrSql)
            IDСотр = Nothing
            IDСотр = ds.Rows(0).Item(0)

            ДанСотр(IDСотр)

            Чист()

            'Dim oWordPara As Microsoft.Office.Interop.Word.Paragraph

            'Dim Pr As Process = Process.Start("PrintOut")
            'Pr.WaitForExit()


            Try
                If IO.Directory.Exists("c:\Users\Public\Documents\Рик") Then
                    'IO.Directory.Delete("c:\Users\Public\Documents\Рик", True)
                    'IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рик")
                Else
                    IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рик")
                End If
            Catch ex As Exception

            End Try



            Try
                IO.File.Copy(OnePath & "\ОБЩДОКИ\Рестлайн\ДопСогл от 29.12.2018.docx", "C:\Users\Public\Documents\Рик\ДопСогл от 29.12.2018.docx")
            Catch ex As Exception
                If "ДопСогл от 29.12.2018.docx" <> "" Then IO.File.Delete("C:\Users\Public\Documents\Рик\ДопСогл от 29.12.2018.docx")
                IO.File.Copy(OnePath & "\ОБЩДОКИ\Рестлайн\ДопСогл от 29.12.2018.docx", "C:\Users\Public\Documents\Рик\ДопСогл от 29.12.2018.docx")
            End Try

            Dim oWord As Microsoft.Office.Interop.Word.Application
            Dim oWordDoc As Microsoft.Office.Interop.Word.Document
            oWord = CreateObject("Word.Application")
            oWord.Visible = False
            oWordDoc = oWord.Documents.Add("C:\Users\Public\Documents\Рик\ДопСогл от 29.12.2018.docx")

            With oWordDoc.Bookmarks
                .Item("ДопСогл1").Range.Text = ДатаКонтр
                .Item("ДопСогл2").Range.Text = НомерКонтр
                .Item("ДопСогл3").Range.Text = MaskedTextBox1.Text
                .Item("ДопСогл4").Range.Text = ФИОПолноеСотр
                .Item("ДопСогл5").Range.Text = ТарифСт
                .Item("ДопСогл6").Range.Text = ПроцОкл
                .Item("ДопСогл7").Range.Text = РасчДолжОкл & " (" & ЧислоПропис(РасчДолжОкл) & ") "
                .Item("ДопСогл8").Range.Text = ФИОСотрКор
                .Item("ДопСогл9").Range.Text = " (" & ЧислоПропис(ТарифСт) & ") "

            End With

            If Not IO.Directory.Exists(OnePath & "\Кадры\Рестлайн\DopSoglashenie\" & Year(Now)) Then
                IO.Directory.CreateDirectory(OnePath & "\Кадры\Рестлайн\DopSoglashenie\" & Year(Now))
            End If

            oWordDoc.SaveAs2("C:\Users\Public\Documents\Рик\DopSoglashenie " & ФИОСотрКор & " от " & MaskedTextBox1.Text & " (Изм.Оклад.Вып.Зп.)" & ".docx",,,,,, False)
            СохрЗак(i) = "C:\Users\Public\Documents\Рик\DopSoglashenie " & ФИОСотрКор & " от " & MaskedTextBox1.Text & " (Изм.Оклад.Вып.Зп.)" & ".docx"
            'oWordDoc.SaveAs2("U: \Офис\Финансовый\6. Бух.услуги\Кадры\" & Клиент & "\Заявление\" & Год & "\" & Заявление(9) & " (заявление)" & ".doc",,,,,, False)


            Try
                IO.File.Copy("C:\Users\Public\Documents\Рик\DopSoglashenie " & ФИОСотрКор & " от " & MaskedTextBox1.Text & " (Изм.Оклад.Вып.Зп.)" & ".docx", OnePath & "\Кадры\Рестлайн\DopSoglashenie\" & Year(Now) & "\DopSoglashenie " & ФИОСотрКор & " от " & MaskedTextBox1.Text & " (Изм.Оклад.Вып.Зп.)" & ".docx.")
            Catch ex As Exception
                If MessageBox.Show("DopSoglashenie с сотрудником " & ФИОСотрКор & " существует. Заменить старый документ новым?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then
                    Try
                        IO.File.Delete(OnePath & "\Кадры\Рестлайн\DopSoglashenie\" & Year(Now) & "\DopSoglashenie " & ФИОСотрКор & " от " & MaskedTextBox1.Text & " (Изм.Оклад.Вып.Зп.)" & ".docx.")
                    Catch ex1 As Exception
                        MessageBox.Show("Закройте файл!", Рик)
                    End Try
                    IO.File.Copy("C:\Users\Public\Documents\Рик\DopSoglashenie " & ФИОСотрКор & " от " & MaskedTextBox1.Text & " (Изм.Оклад.Вып.Зп.)" & ".docx", OnePath & "\Кадры\Рестлайн\DopSoglashenie\" & Year(Now) & "\DopSoglashenie " & ФИОСотрКор & " от " & MaskedTextBox1.Text & " (Изм.Оклад.Вып.Зп.)" & ".docx.")
                End If
            End Try


            oWordDoc.Close(True)
            oWord.Quit(True)

        Next

        ПечатьДоков(СохрЗак)



    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim sd As Integer = ПровЗаполн()
        If sd = 1 Then
            Exit Sub
        End If
        Me.Cursor = Cursors.WaitCursor

        доки()

        Me.Cursor = Cursors.Default
        MessageBox.Show("Документы в печати!", Рик)
        Me.Close()



    End Sub






End Class