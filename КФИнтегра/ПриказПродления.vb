Option Explicit On
Imports System.Data.OleDb

Public Class ПриказПродления
    Public Da As New OleDbDataAdapter 'Адаптер
    Public Ds As New DataSet 'Пустой набор записей
    Dim tbl As New DataTable
    Dim cb As OleDb.OleDbCommandBuilder

    Dim Год2, НомерПриказа, ФормаСобстПолн, ДолжРуков, ФИОРукРодПад, ОснованиеДейств, ФИОКор, УНП, КонтТелефон,
        ЮрАдрес, РасСчет, БИК, АдресБанка, ЭлАдрес, ФормаСобствКор, СборноеРеквПолн, Банк As String
    Dim IDСотр As Integer


    Private Sub ПриказПродления_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MdiParent = MDIParent1
        'Me.WindowState = FormWindowState.Maximized
        Год2 = Year(Now)

        Me.ComboBox1.Items.Clear()
        For Each r As DataRow In СписокКлиентовОсновной.Rows
            Me.ComboBox1.Items.Add(r(0).ToString)
        Next

        ComboBox2.Enabled = False


    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        Me.ComboBox19.Text = ""
        Me.ComboBox2.Text = ""
        Me.TextBox41.Text = ""
        Me.TextBox57.Text = ""
        Me.ComboBox3.Items.Clear()
        Me.ComboBox3.Text = ""
        Dim Нет As String = "Нет"



        Dim f = From x In dtSotrudnikiAll Where x.Item("НазвОрганиз") = ComboBox1.Text And x.Item("НаличеДогПодряда") = "Нет" Order By x.Item("ФИОСборное") Select x.Item("ФИОСборное")


        Me.ComboBox19.Items.Clear()
        Me.ComboBox19.AutoCompleteCustomSource.Clear()
        For Each r In f
            Me.ComboBox19.AutoCompleteCustomSource.Add(r.ToString())
            Me.ComboBox19.Items.Add(r.ToString)
        Next

        Dim ftp = listFluentFTP(ComboBox1.Text & "/Приказ")
        'Dim Folders() As String
        'Try
        '    Folders = IO.Directory.GetDirectories(OnePath & ComboBox1.Text & "\Приказ", "*", IO.SearchOption.TopDirectoryOnly)
        'Catch ex As Exception
        '    MessageBox.Show("Невозможно сформировать Год, т.к. не существует папка - Приказы!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
        '    Exit Sub
        'End Try

        'Dim gth4 As String
        'For n As Integer = 0 To Folders.Length - 1
        '    gth4 = ""
        '    gth4 = IO.Path.GetFileName(Folders(n))
        '    Folders(n) = gth4
        '    'TextBox44.Text &= gth + vbCrLf
        'Next


        For Each r In ftp
            ComboBox3.Items.Add(r.ToString)
        Next


    End Sub



    Private Sub TextBox41_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox41.KeyDown

        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox58.Focus()


            Dim pl As String
            If TextBox41.Text <> "" Then
                Dim i As Integer = CInt(TextBox41.Text)
                Select Case i

                    Case < 10
                        pl = Str(i)
                        TextBox41.Text="00" & i

                    Case 10 To 99
                        pl = Str(i)
                        TextBox41.Text="0" & i
                End Select
            End If

        End If
    End Sub

    Private Sub TextBox41_TextChanged(sender As Object, e As EventArgs) Handles TextBox41.TextChanged

    End Sub

    Private Sub ComboBox19_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox19.SelectedIndexChanged

    End Sub

    Private Sub ComboBox19_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox19.SelectedValueChanged
        If ComboBox19.Text <> "" Then ComboBox2.Enabled = True
    End Sub

    Private Sub Label65_Click(sender As Object, e As EventArgs) Handles Label65.Click

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub ComboBox3_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedValueChanged
        ComboBox2.Items.Clear()
        ComboBox2.Text = ""

        Dim ftp = listFluentFTP(ComboBox1.Text & "/Приказ/" & ComboBox3.Text)
        'Dim Files3()

        'Files3 = (IO.Directory.GetFiles(OnePath & ComboBox1.Text & "\Приказ\" & ComboBox3.Text, "*.doc", IO.SearchOption.TopDirectoryOnly))

        'Dim gth As String
        'For n As Integer = 0 To Files3.Length - 1
        '    gth = ""
        '    gth = IO.Path.GetFileName(Files3(n))
        '    Files3(n) = gth
        '    'TextBox44.Text &= gth + vbCrLf
        'Next

        For Each r In ftp
            ComboBox2.Items.Add(r.ToString)
        Next
    End Sub

    Private Sub TextBox58_TextChanged(sender As Object, e As EventArgs) Handles TextBox58.TextChanged

    End Sub

    Private Sub TextBox58_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox58.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox57.Focus()
        End If
    End Sub

    Private Sub TextBox57_TextChanged(sender As Object, e As EventArgs) Handles TextBox57.TextChanged

    End Sub

    Private Sub TextBox57_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox57.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.Button1.Focus()
        End If
    End Sub

    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox19.Focus()
        End If
    End Sub

    Private Sub ComboBox19_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox19.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox3.Focus()
        End If
    End Sub

    Private Sub ComboBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox2.Focus()
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub

    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox41.Focus()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If MessageBox.Show("Создать приказ о продлении контракта?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.Cancel Then Exit Sub
        If ComboBox1.Text = "" Then
            MessageBox.Show("Выберите организацию!", Рик)
            Exit Sub
        End If
        If ComboBox19.Text = "" Then
            MessageBox.Show("Выберите сотрудника!", Рик)
            Exit Sub
        End If
        If TextBox41.Text = "" Then
            MessageBox.Show("Напишите номер приказа!", Рик)
            Exit Sub
        End If


        Me.Cursor = Cursors.WaitCursor


        Dim st = From x In dtSotrudnikiAll Where x.Item("НазвОрганиз") = ComboBox1.Text And x.Item("ФИОСборное") = ComboBox19.Text Select x.Item("КодСотрудники")


        '        Dim StrSql As String = "Select КодСотрудники From Сотрудники
        'Where НазвОрганиз='" & ComboBox1.Text & "' AND ФИОСборное = '" & ComboBox19.Text & "'"

        '        Dim c As New OleDbCommand With {
        '                .Connection = conn,
        '                .CommandText = StrSql
        '            }
        '        Dim ds As New DataSet
        '        Dim da As New OleDbDataAdapter(c)
        '        da.Fill(ds, "Конт")
        IDСотр = st(0)

        If TextBox57.Text <> "" Then
            НомерПриказа = TextBox41.Text & "-" & TextBox58.Text & "-" & TextBox57.Text
        Else
            НомерПриказа = TextBox41.Text & "-" & TextBox58.Text
        End If

        Dim list As New Dictionary(Of String, Object)()
        '
        list.Add("@ID", IDСотр)
        list.Add("@НазвОрг", ComboBox1.Text)
        'list.Add("@Должность", Должность)


        Updates(stroka:="UPDATE КарточкаСотрудника
SET ПриказПродлКонтр='" & НомерПриказа & "'
Where IDСотр=@ID", list)

        '______________________  ________________________________ выборка по клиенту




        Dim ds8 = Selects(StrSql:="SELECT Клиент.ФормаСобств, Клиент.УНП, Клиент.ЮрАдрес, 
Клиент.Банк, Клиент.БИКБанка, Клиент.АдресБанка, Клиент.Отделение, Клиент.РасчСчетРубли, Клиент.ДолжнРуководителя, 
Клиент.ФИОРуководителя, Клиент.ОснованиеДейств, Клиент.ФИОРукРодПадеж, Клиент.КонтТелефон, Клиент.ЭлАдрес, Клиент.РукИП
From Клиент
Where Клиент.НазвОрг =@НазвОрг", list)


        Dim РуковИП As String
        If ds8.Rows(0).Item(14) = True Then
            РуковИП = "ИП "
        Else
            РуковИП = ""
        End If


        ФормаСобстПолн = ds8.Rows(0).Item(0).ToString
        ДолжРуков = ds8.Rows(0).Item(8).ToString
        ФИОРукРодПад = РуковИП & ds8.Rows(0).Item(11).ToString
        ОснованиеДейств = ds8.Rows(0).Item(10).ToString
        'МестоРаб = n & " " & w & " " & ComboBox18.Text

        'короткое фио клиента
        Dim nm As String = ds8.Rows(0).Item(9).ToString
        Dim nm0 As Integer = Len(ds8.Rows(0).Item(9).ToString)
        Dim nm1 As String = Strings.Left(nm, InStr(nm, " "))
        Dim nm2 As Integer = Len(nm1)
        Dim nm3 As String = Strings.Right(nm, (nm0 - nm2))
        Dim nm31 As Integer = Len(nm3)
        Dim nm4 As String = Strings.UCase(Strings.Left(Strings.Left(nm3, InStr(nm3, " ")), 1))
        Dim nm41 As Integer = Len(Strings.Left(nm3, InStr(nm3, " ")))
        Dim nm5 As String = Strings.UCase(Strings.Left(Strings.Right(nm3, nm31 - nm41), 1))


        ФИОКор = РуковИП & nm1 & " " & nm4 & "." & nm5 & "."
        УНП = ds8.Rows(0).Item(1).ToString
        КонтТелефон = ds8.Rows(0).Item(12).ToString
        ЮрАдрес = ds8.Rows(0).Item(2).ToString
        РасСчет = ds8.Rows(0).Item(7).ToString
        Банк = ds8.Rows(0).Item(3).ToString
        БИК = ds8.Rows(0).Item(4).ToString
        АдресБанка = ds8.Rows(0).Item(5).ToString
        ЭлАдрес = ds8.Rows(0).Item(13).ToString

        '______

        'сокращенное название и сборное клиента
        Dim ds9 = dtformft.Select("ПолноеНазвание = '" & ds8.Rows(0).Item(0).ToString & "'")

        ФормаСобствКор = ds9(0).Item(0).ToString
        СборноеРеквПолн = ФормаСобствКор & " """ & ComboBox1.Text & """ " & ds8.Rows(0).Item(2).ToString & " IBAN " _
        & ds8.Rows(0).Item(7).ToString & " в " & ds8.Rows(0).Item(3).ToString & " " _
        & ds8.Rows(0).Item(5).ToString & " " & ds8.Rows(0).Item(6).ToString & " БИК " _
        & ds8.Rows(0).Item(4).ToString & " УНП " & ds8.Rows(0).Item(1).ToString

        '__________________________  ____________________ выборка по сотруднику

        Dim ds4 = Selects(StrSql:="SELECT КарточкаСотрудника.ПродлКонтрС, КарточкаСотрудника.ПриказПродлКонтр, Штатное.Должность, Штатное.Разряд,
ДогСотрудн.Контракт, ДогСотрудн.ДатаКонтракта, КарточкаСотрудника.СрокПродлКонтракта,
КарточкаСотрудника.ПродлКонтрПо, КарточкаСотрудника.НомерУведомлПродКонтр,
КарточкаСотрудника.ДатаУведомлПродКонтр, Сотрудники.Фамилия, Сотрудники.Имя, Сотрудники.Отчество, КарточкаСотрудника.ОснованиеУвольн
FROM ((Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр) INNER JOIN ДогСотрудн ON Сотрудники.КодСотрудники = ДогСотрудн.IDСотр) INNER JOIN Штатное ON Сотрудники.КодСотрудники = Штатное.ИДСотр
Where Сотрудники.КодСотрудники =@ID", list)


        Dim Dat As Date
        Try
            If ds4.Rows(0).Item(13).ToString <> "" Then
                MessageBox.Show("Cотрудник уже уволен !", Рик, MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show("Данный сотрудник не нуждается в продлении контракта")
            Me.Cursor = Cursors.Default
            Exit Sub

        End Try


        Try
            Dat =ds4.Rows(0).Item(0) 'дата приказа за 15 дней до наступления начала контракта продления
        Catch ex As Exception
            MessageBox.Show("С cотрудником еще не продлен контракт !", Рик, MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            Me.Cursor = Cursors.Default
            Exit Sub
        End Try



        Dim Datco As String = Dat.AddDays(-15)

        Dim inp As String = InputBox("Введите ФИО сотрудника " & ComboBox19.Text & " в Творительном падеже 'Продлить контракт с Кем?, Чем?'", Рик)

        Do Until inp <> ""
            MessageBox.Show("Повторите ввод данных!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Error)
            inp = InputBox("Введите ФИО сотрудника " & ComboBox19.Text & " в Творительном падеже 'Продлить контракт с Кем?, Чем?'", Рик)
        Loop
        Me.Cursor = Cursors.WaitCursor
        Dim ДолТворПад As String = ДолжТворПадеж(ds4.Rows(0).Item(2).ToString())

        Dim Разряд As String ' если есть разряд то соединяем должность и разряд
        If ds4.Rows(0).Item(3).ToString() <> "" And Not ds4.Rows(0).Item(3).ToString() = "-" Then
            Разряд = разрядстрока(ds4.Rows(0).Item(3).ToString())
            ДолТворПад = ДолТворПад & " " & Разряд
        End If

        Dim КорИмя As String = Strings.Left(ds4.Rows(0).Item(11).ToString(), 1)
        Dim КорОтч As String = Strings.Left(ds4.Rows(0).Item(12).ToString(), 1)


        Dim ДатаУведомления As Date =ds4.Rows(0).Item(9)
        ДатаУведомления = ДатаУведомления.AddYears(-Int(ds4.Rows(0).Item(6)))

        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document
        'Dim oWordPara As Microsoft.Office.Interop.Word.Paragraph

        oWord = CreateObject("Word.Application")
        oWord.Visible = False

        'приказ на продление контракта
        'Try 'проверка если есть в С: папке файл Контакт его удаляем и создаем новый
        '    IO.File.Copy(OnePath & "\ОБЩДОКИ\General\PrikazNaProdlenie.doc", "C:\Users\Public\Documents\PrikazNaProdlenie.doc")
        'Catch ex As Exception
        '    If "PrikazNaProdlenie.doc" <> "" Then IO.File.Delete("C:\Users\Public\Documents\PrikazNaProdlenie.doc")
        '    IO.File.Copy(OnePath & "\ОБЩДОКИ\General\PrikazNaProdlenie.doc", "C:\Users\Public\Documents\PrikazNaProdlenie.doc")
        'End Try

        Начало("PrikazNaProdlenie.doc")
        oWordDoc = oWord.Documents.Add(firthtPath & "\PrikazNaProdlenie.doc")

        With oWordDoc.Bookmarks
            .Item("П1").Range.Text = Strings.Left(Datco, 10) & "г" ' срок продления контракта с минцс 15 дней
            .Item("П2").Range.Text = ds4.Rows(0).Item(1).ToString 'номер приказа
            .Item("П3").Range.Text = inp ' ФИО сотрудника в творит падеже
            .Item("П4").Range.Text = LCase(ДолТворПад) ' должность и разряд
            .Item("П5").Range.Text = ds4.Rows(0).Item(4).ToString() ' номер контракта
            .Item("П6").Range.Text = Strings.Left(ds4.Rows(0).Item(5).ToString(), 10) 'дата контракта
            .Item("П7").Range.Text = ds4.Rows(0).Item(6).ToString() ' срок продления контракта
            .Item("П8").Range.Text = Склонение2(ds4.Rows(0).Item(6).ToString()) ' срок продления - склонение времени
            .Item("П9").Range.Text = Strings.Left(ds4.Rows(0).Item(0).ToString(), 10) 'продление контракта с
            .Item("П10").Range.Text = Strings.Left(ds4.Rows(0).Item(7).ToString(), 10) 'продление контракта по
            .Item("П11").Range.Text = ds4.Rows(0).Item(8).ToString() 'номер уведомления
            .Item("П12").Range.Text = Strings.Left(ДатаУведомления.ToString(), 10) 'дата уведомления
            .Item("П13").Range.Text = ds4.Rows(0).Item(4).ToString()
            .Item("П14").Range.Text = Strings.Left(ds4.Rows(0).Item(5).ToString(), 10)
            .Item("П22").Range.Text = ds4.Rows(0).Item(10).ToString() & " " & КорИмя & ". " & КорОтч & "."
            .Item("П25").Range.Text = ФормаСобстПолн
            .Item("П26").Range.Text = ComboBox1.Text
            .Item("П27").Range.Text = ЮрАдрес
            .Item("П28").Range.Text = УНП
            .Item("П30").Range.Text = АдресБанка
            .Item("П31").Range.Text = БИК
            .Item("П33").Range.Text = ЭлАдрес
            .Item("П34").Range.Text = КонтТелефон
            .Item("П36").Range.Text = ДолжРуков & " " & ФормаСобствКор & " """ & ComboBox1.Text & " """
            .Item("П38").Range.Text = ФИОКор
            .Item("П15").Range.Text = "р/с " & РасСчет
        End With

        Dim fName As String = ds4.Rows(0).Item(1).ToString & " продление контракта " & ds4.Rows(0).Item(10).ToString() & " от " & Strings.Left(Datco, 10) & ".doc"
        oWordDoc.SaveAs2(PathVremyanka & fName,,,,,, False)
        Dim ДопСоглFTP As New List(Of String)
        ДопСоглFTP.AddRange(New String() {ComboBox1.Text & "\Приказ\" & Now.Year, fName})
        oWordDoc.Close()
        oWord.Quit()

        Конец(ComboBox1.Text & "\Приказ\" & Now.Year, fName, IDСотр, ComboBox1.Text, "\PrikazNaProdlenie.doc", "ПриказПродлениеКонтракта")

        dtKartochkaSotrudnika()
        If MessageBox.Show("Приказ о продлении контракта для сотрудника " & ComboBox19.Text & " создан!" & vbCrLf & "Распечатать приказ о продлении?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
            Dim massFTP As New ArrayList()
            massFTP.Add(ДопСоглFTP)
            ПечатьДоковFTP(massFTP)
        End If



        Me.Cursor = Cursors.Default

        Статистика1(ComboBox19.Text, "Оформление приказа продления контракта", ComboBox1.Text)

        ComboBox1.Text = ""
        ComboBox19.Items.Clear()
        ComboBox19.Text = ""
        ComboBox3.Items.Clear()
        ComboBox3.Text = ""
        ComboBox2.Items.Clear()
        ComboBox2.Text = ""
        TextBox41.Text = ""
        TextBox57.Text = ""


    End Sub


End Class