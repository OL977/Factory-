Option Explicit On
Imports System.Data.OleDb
Imports System.IO
Imports System.Net.FtpClient
Imports System.Data.Linq


Public Class Печать
    Dim Фамилия As String
    Dim FilesList() As String
    Dim listCombo2 As List(Of String)
    Dim listCombo3 As List(Of String)
    Dim listCombo4 As New List(Of String)
    Dim listCombo5 As List(Of String)
    Dim proc1 As Process

    Dim combx1 As String


    Private Sub Печать_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MdiParent = MDIParent1






        Dim Год As Date
        Год = Now
        Год = Format(Год, "dd MMMM yyyy")
        'If Me.Прием_Load = vbTrue Then Form1.Load = False


        Me.ComboBox1.Items.Clear()
        For Each r As DataRow In СписокКлиентовОсновной.Rows
            Me.ComboBox1.Items.Add(r(0).ToString)
        Next



        For i As Integer = 0 To 3
            Me.ComboBox3.Items.Add(Year(Now) - i)
        Next
        dtPutiDokumentov()
    End Sub
    Public Async Sub allAsync()
        Await Task.Run(Sub() allFiles())
    End Sub


    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        combx1 = ComboBox1.Text
        'allAsync()
        'Parallel.Invoke(Sub() all())


        Dim list = listFluentFTP(ComboBox1.Text & "/")

        ComboBox2.Items.Clear()

        For r As Integer = 0 To list.Count - 1
            Me.ComboBox2.AutoCompleteCustomSource.Add(list(r).ToString())
            Me.ComboBox2.Items.Add(list(r).ToString())
        Next
        ComboBox2.Text = ""
        ComboBox3.Text = ""




        'Dim dataView As New DataView(dtSotrudnikiAll) 'соритровка datatable
        'dataView.Sort = " НазвОрганиз DESC"
        'Dim dataTable As DataTable = dataView.ToTable()

        Dim var = From x In dtSotrudnikiAll.Rows Where x.Item("НазвОрганиз") = ComboBox1.Text Order By x.Item("ФИОСборное") Select x

        'Dim ds As DataTable = Selects(StrSql:="SELECT ФИОСборное FROM Сотрудники WHERE НазвОрганиз= '" & ComboBox1.Text & "' ORDER BY ФИОСборное")
        'Await Task.Run(Sub() all())
        Me.ComboBox4.AutoCompleteCustomSource.Clear()
        Me.ComboBox4.Items.Clear()
        Me.ComboBox5.Items.Clear()
        For Each r As DataRow In var
            Me.ComboBox4.AutoCompleteCustomSource.Add(r.Item("ФИОСборное").ToString())
            Me.ComboBox4.Items.Add(r.Item("ФИОСборное").ToString)
            Me.ComboBox5.Items.Add(r.Item("КодСотрудники").ToString)
        Next
        ComboBox4.Text = ""


        If ComboBox1.Text <> "" Then
            Label5.Text = """" & ComboBox1.Text & """"
            Label5.ForeColor = Color.Red
        Else
            Label5.Text = ""
            Label5.ForeColor = Color.Black
        End If

        ListBox2.Items.Clear()
        ListBox1.Items.Clear()



    End Sub
    Private Sub allFiles()
        'Dim sw As New Stopwatch 'вычисление выполнения метода
        'sw.Start()

        Готовый.Clear()
        Dim listPath = listFluentFTP(combx1 & "/")
        Dim listPath2 As New List(Of String)
        Dim listYear As New List(Of String)
        Dim Промеж As New List(Of String)

        For Each item In listPath
            listPath2 = listFluentFTP(combx1 & "/" & item)
        Next
        'For Each item In listPath2
        '    listYear = listFluentFTP(ComboBox1.Text & "/" & item)
        'Next
        For Each item1 In listPath
            For Each item2 In listPath2
                If item1.Contains("инструкции") Then
                    Промеж = listFluentFTP(combx1 & "/" & item1 & "/")
                    Готовый.AddRange(Промеж)
                Else
                    Промеж = listFluentFTP(combx1 & "/" & item1 & "/" & item2 & "/")
                    Готовый.AddRange(Промеж)
                End If

            Next
        Next
        'sw.Stop()
        'MessageBox.Show((sw.ElapsedMilliseconds / 100.0).ToString())
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        refreshList2()

    End Sub
    Private Sub refreshList2()

        Dim var() = dtSotrudnikiAll.Select("ФИОСборное='" & ComboBox4.Text & "'")
        Фамилия = var(0).Item("Фамилия")

        Label6.Text = ComboBox5.Items.Item(ComboBox4.SelectedIndex)

        'Dim var6 = From x In Готовый Where x.Contains(Фамилия) Select x
        Dim var3 As EnumerableRowCollection(Of DataRow), var4

        Try
            var3 = From y In dtPutiDokumentovAll.AsEnumerable() Where Not IsDBNull(y.Item("IDСотрудник"))

            var4 = From x In var3.AsEnumerable() Where x.Item("IDСотрудник") = CType(Label6.Text, Integer) Select x.Item("ИмяФайла")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


        ListBox2.Items.Clear()
        For Each r In var4
            ListBox2.Items.Add(r.ToString)
        Next


    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

        ListBox1.Items.Clear()
        ComboBox3.Text = String.Empty
        If ComboBox2.Text = "Должностные инструкции" Then

            listCombo2 = listFluentFTP(ComboBox1.Text & "/Должностные инструкции")

            For x As Integer = 0 To listCombo2.Count - 1
                'ListBox1.Items.Add(Replace(listCombo2(x).ToString, ComboBox1.Text & "/Должностные инструкции", ""))
                ListBox1.Items.Add(listCombo2(x).ToString)
            Next

        End If

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        RefreshList1()

    End Sub
    'Private Sub ДолжИнстр()

    '    'Dim Files2() As String
    '    'Try
    '    '    Files2 = IO.Directory.GetFiles(OnePath & ComboBox1.Text & "\" & ComboBox2.Text, "*.doc*", IO.SearchOption.TopDirectoryOnly)
    '    'Catch ex As Exception
    '    '    MessageBox.Show("Нет документов!", Рик)
    '    '    Exit Sub
    '    'End Try

    '    'Dim gth4 As String
    '    'For n As Integer = 0 To Files2.Length - 1
    '    '    gth4 = ""
    '    '    gth4 = IO.Path.GetFileName(Files2(n))
    '    '    Files2(n) = gth4
    '    '    'TextBox44.Text &= gth + vbCrLf
    '    'Next

    '    ''ListBox2.Items.Add(Files2)

    '    Dim list = listFTP()

    '    ListBox1.Items.Clear()

    '    For i = 0 To Files2.Length - 1 ' Распечатываем весь получившийся массив
    '        ListBox1.Items.Add(Files2(i))
    '    Next ' На ListBox2
    'End Sub
    Private Sub ШтЭксел()

        Try
            If listCombo2.Count > 0 Then
                listCombo2.Clear()
            End If
        Catch ex As Exception

        End Try

        Dim listCombo6 As New List(Of String)
        Try
            'Files2 = IO.Directory.GetFiles(OnePath & ComboBox1.Text & "\" & ComboBox2.Text & "\" & ComboBox3.Text, "*.xls*", IO.SearchOption.TopDirectoryOnly)
            'listCombo2 = listFTP(ComboBox2.Text)
            listCombo6 = listFluentFTP(ComboBox1.Text & "/" & ComboBox2.Text & "/" & ComboBox3.Text & "/")

            'Dim fs As String = listCombo2(0).ToString

            For Each item In listCombo6
                ListBox1.Items.Add(item)
            Next



            'For x As Integer = 0 To listCombo6.Count - 1
            '    Dim f As String = listCombo6(x).ToString
            '    listCombo4.Add(ComboBox1.Text & "/" & ComboBox2.Text & "/" & ComboBox3.Text & "/" & f)
            'Next



            'listCombo2 = listCombo2 & "/" & listCombo3
        Catch ex As Exception
            MessageBox.Show("В этому году нет такой папки!", Рик)
            Exit Sub
        End Try

        'ListBox1.Items.Clear()
        'For x As Integer = 0 To listCombo3.Count - 1
        '    ListBox1.Items.Add(listCombo3(x).ToString)
        'Next
    End Sub

    Private Sub RefreshList1()

        If ComboBox2.Text = "Штатное расписание" Then
            ШтЭксел()
            Exit Sub
        End If




        listCombo5 = listFluentFTP(ComboBox1.Text & "/" & ComboBox2.Text & "/" & ComboBox3.Text & "/")
        ListBox1.Items.Clear()
        For x As Integer = 0 To listCombo5.Count - 1
            'ListBox1.Items.Add(Replace(listCombo2(x).ToString, ComboBox1.Text & "/Должностные инструкции", ""))
            ListBox1.Items.Add(listCombo5(x).ToString)
        Next






        'Dim Files2() As String
        'Try
        '    Files2 = IO.Directory.GetFiles(OnePath & ComboBox1.Text & "\" & ComboBox2.Text & "\" & ComboBox3.Text, "*.doc*", IO.SearchOption.TopDirectoryOnly)
        'Catch ex As Exception
        '    MessageBox.Show("В этому году нет такой папки!", Рик)
        '    Exit Sub
        'End Try

        'Dim gth4 As String
        'For n As Integer = 0 To Files2.Length - 1
        '    gth4 = ""
        '    gth4 = IO.Path.GetFileName(Files2(n))
        '    Files2(n) = gth4
        '    'TextBox44.Text &= gth + vbCrLf
        'Next

        ''ListBox2.Items.Add(Files2)


        'ListBox1.Items.Clear()

        'For i = 0 To Files2.Length - 1 ' Распечатываем весь получившийся массив
        '    ListBox1.Items.Add(Files2(i)) ' На ListBox2
        'Next
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If ListBox2.SelectedIndex = -1 Then
            MessageBox.Show("Выберите документ для печати!", Рик, MessageBoxButtons.OK)
            Exit Sub
        End If




        If Not ListBox2.SelectedIndex = -1 Then
            Dim ft = From x In dtPutiDokumentovAll.Rows Where x.item("IDСотрудник") = CType(Label6.Text, Integer) And x.item("ИмяФайла") = ListBox2.SelectedItem
                     Select x.item("ИмяФайла")
            Dim ft1 = From x In dtPutiDokumentovAll.Rows Where x.item("IDСотрудник") = CType(Label6.Text, Integer) And x.item("ИмяФайла") = ListBox2.SelectedItem
                      Select x.item("Путь")

            Dim _string As String = ft1(0).ToString & ft(0).ToString
            ВыгрузкаФайловНаЛокалыныйКомп(_string, PathVremyanka & ft(0).ToString)

            If _string = "" Then Exit Sub

            Dim wdApp As New Microsoft.Office.Interop.Word.Application
            Dim wdDoc As Microsoft.Office.Interop.Word.Document
            wdApp.Visible = False
            wdDoc = wdApp.Documents.Open(FileName:=PathVremyanka & ft(0).ToString)
            Try
                wdDoc.PrintOut(True) 'печать
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
            wdDoc.Close()
            wdApp.Quit()

            Try
                ВременнаяПапкаУдалениеФайла(PathVremyanka & ft(0).ToString)
            Catch ex As Exception

            End Try


        End If





        'If Not ListBox2.SelectedIndex = -1 Then
        '    Dim ft = From x In dtPutiDokumentovAll.Rows Where x.item("IDСотрудник") = CType(Label6.Text, Integer) And x.item("ИмяФайла") = ListBox2.SelectedItem
        '             Select x.item("ИмяФайла")
        '    Dim ft1 = From x In dtPutiDokumentovAll.Rows Where x.item("IDСотрудник") = CType(Label6.Text, Integer) And x.item("ИмяФайла") = ListBox2.SelectedItem
        '              Select x.item("Путь")

        '    Dim _string As String = ft1(0).ToString & ft(0).ToString
        '    ВыгрузкаФайловНаЛокалыныйКомп(_string, PathVremyanka & ft(0).ToString)
        '    If _string = "" Then Exit Sub

        '    Dim wdApp As New Microsoft.Office.Interop.Word.Application
        '    Dim wdDoc As Microsoft.Office.Interop.Word.Document
        '    wdApp.Visible = False
        '    wdDoc = wdApp.Documents.Open(FileName:=PathVremyanka & ft(0).ToString)
        '    Try
        '        wdDoc.PrintOut(True) 'печать
        '    Catch ex As Exception
        '        MessageBox.Show(ex.Message)
        '    End Try
        '    wdDoc.Close()
        '    wdApp.Quit()

        'End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ComboBox4.Items.Clear()
        ComboBox2.Items.Clear()
        ComboBox4.Text = String.Empty
        ComboBox3.Text = String.Empty
        ComboBox2.Text = String.Empty
        ComboBox1.Text = String.Empty
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If ListBox1.SelectedIndex = -1 Then
            MessageBox.Show("Выберите документ для удаления!", Рик, MessageBoxButtons.OK)
            Exit Sub
        End If

        Dim inp As String
        If ПодтверждПароляУдаление = False Then
            inp = InputBox("Введите пароль", Рик, "Пароль")
            If inp = "" Then Exit Sub
            Dim fn As Integer
            Try
                fn = CType(inp, Integer)
            Catch ex As Exception
                Exit Sub
            End Try


            If Not fn = 6986577 Then
                Exit Sub
            Else
                ПодтверждПароляУдаление = True
                    If MessageBox.Show("Удалить выбранный файл?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.Cancel Then
                        Exit Sub
                    End If

                    Dim ft, ft1 As String
                    'Dim ft = From x In DirAll Where x.Key = ListBox1.SelectedItem Select x.Value
                    If ComboBox2.Text = "Должностные инструкции" Then

                        ft = "/" & ComboBox1.Text & "/" & ComboBox2.Text & "/"
                        ft1 = ComboBox1.Text & "/" & ComboBox2.Text & "/"
                    Else
                        ft = "/" & ComboBox1.Text & "/" & ComboBox2.Text & "/" & ComboBox3.Text & "/"
                        ft1 = ComboBox1.Text & "/" & ComboBox2.Text & "/" & ComboBox3.Text & "/"
                    End If


                    Dim lp = FTPString & ft1 & ListBox1.SelectedItem

                    Dim list As New Dictionary(Of String, Object)
                    list.Add("@ПолныйПуть", lp)

                    Updates(stroka:="DELETE FROM ПутиДокументов WHERE ПолныйПуть=@ПолныйПуть", list, "ПутиДокументов")

                    DeleteFluentFTP(ft & ListBox1.SelectedItem)
                End If

            Else

                If MessageBox.Show("Удалить выбранный файл?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.Cancel Then
                    Exit Sub
                End If

                Dim ft, ft1 As String
                'Dim ft = From x In DirAll Where x.Key = ListBox1.SelectedItem Select x.Value
                If ComboBox2.Text = "Должностные инструкции" Then

                    ft = "/" & ComboBox1.Text & "/" & ComboBox2.Text & "/"
                    ft1 = ComboBox1.Text & "/" & ComboBox2.Text & "/"
                Else
                    ft = "/" & ComboBox1.Text & "/" & ComboBox2.Text & "/" & ComboBox3.Text & "/"
                    ft1 = ComboBox1.Text & "/" & ComboBox2.Text & "/" & ComboBox3.Text & "/"
                End If


                Dim lp = FTPString & ft1 & ListBox1.SelectedItem

                Dim list As New Dictionary(Of String, Object)
                list.Add("@ПолныйПуть", lp)

                Updates(stroka:="DELETE FROM ПутиДокументов WHERE ПолныйПуть=@ПолныйПуть", list, "ПутиДокументов")

            DeleteFluentFTP(ft & ListBox1.SelectedItem)
        End If









        'Do Until inp <> ""
        '    MessageBox.Show("Повторите ввод данных!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    inp = InputBox("Введите ФИО сотрудника " & vbCrLf & ФИО & vbCrLf & " в Дательном падеже 'Предоставить отпуск Кому?'", Рик)
        'Loop

















        'Dim i As Integer
        'If MessageBox.Show("Удалить выбранные файлы?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.Cancel Then
        '    Exit Sub
        'End If
        'If Not ListBox1.SelectedIndex = -1 Then
        '    For i = 0 To ListBox1.SelectedItems.Count - 1
        '        IO.File.Delete(OnePath & ComboBox1.Text & "\" & ComboBox2.Text & "\" & ComboBox3.Text & "\" & ListBox1.SelectedItems(i))
        '        Статистика(OnePath & ComboBox1.Text & "\" & ComboBox2.Text & "\" & ComboBox3.Text & "\" & ListBox1.SelectedItems(i), "Удаление документов сотрудника", ComboBox1.Text)
        '    Next
        '    MessageBox.Show("Документы удалены!", Рик)
        RefreshList1()

        'End If

        'Dim ff As ListBox.SelectedIndexCollection = ListBox2.SelectedIndices

        'If Not ListBox2.SelectedIndex = -1 Then

        '    For Each p As Integer In ff
        '        IO.File.Delete(FilesList(p)) 'договор подряда
        '        Статистика((FilesList(p)), "Удаление документов сотрудника", ComboBox1.Text)
        '    Next

        '    MessageBox.Show("Документы удалены!", Рик)
        '    refreshList2()
        'End If


    End Sub



    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick

        If ListBox1.SelectedIndex = -1 Then
            MessageBox.Show("Выберите документ для просмотра!", Рик, MessageBoxButtons.OK)
            Exit Sub
        End If

        'Dim f = Process.GetProcessesByName("proc1").Any()




        ''Dim list = listFTP("Должностные инструкции")

        'If ComboBox3.Text = "" Then
        '    If Not ListBox1.SelectedIndex = -1 Then
        '        For i = 0 To ListBox1.SelectedItems.Count - 1
        '            For x As Integer = 0 To listCombo2.Count - 1
        '                If listCombo2(x).ToString.Contains(ListBox1.SelectedItems(i)) Then

        '                    ЗагрФайлаИзСервера(listCombo2(x).ToString, PathVremyanka & ListBox1.SelectedItems(i))
        '                    Dim proc As Process = Process.Start(PathVremyanka & ListBox1.SelectedItems(i))
        '                    proc.WaitForExit()
        '                    proc.Close()

        '                    ЗагрузкаФайловНаСервер(PathVremyanka & ListBox1.SelectedItems(i), FTPString & listCombo2(x).ToString)
        '                End If
        '            Next
        '        Next

        '    End If

        'Else

        '    If Not ListBox1.SelectedIndex = -1 Then
        '        'For i = 0 To ListBox1.SelectedItems.Count - 1
        '        '    For x As Integer = 0 To listCombo4.Count - 1
        '        '        If listCombo4(x).ToString.Contains(ListBox1.SelectedItems(i)) Then

        '        '            ЗагрФайлаИзСервера(listCombo4(x).ToString, PathVremyanka & ListBox1.SelectedItems(i))
        '        '            Dim proc As Process = Process.Start(PathVremyanka & ListBox1.SelectedItems(i))
        '        '            proc.WaitForExit()
        '        '            proc.Close()

        '        '            ЗагрузкаФайловНаСервер(PathVremyanka & ListBox1.SelectedItems(i), FTPString & listCombo4(x).ToString)
        '        '            Exit For
        '        '        End If
        '        '    Next
        '        'Next
        '        Dim str As String = FTPString & ComboBox1.Text & "/" & ComboBox2.Text & "/" & ComboBox3.Text & "/" & ListBox1.SelectedItem
        '        'ЗагрФайлаИзСервера(str, PathVremyanka & ListBox1.SelectedItem)


        Dim ft As String
        'Dim ft = From x In DirAll Where x.Key = ListBox1.SelectedItem Select x.Value
        If ComboBox2.Text = "Должностные инструкции" Then

            ft = "/" & ComboBox1.Text & "/" & ComboBox2.Text & "/"
        Else
            ft = "/" & ComboBox1.Text & "/" & ComboBox2.Text & "/" & ComboBox3.Text & "/"
        End If


        ВыгрузкаФайловНаЛокалыныйКомп(ft & ListBox1.SelectedItem, PathVremyanka & ListBox1.SelectedItem)
        proc1 = Process.Start(PathVremyanka & ListBox1.SelectedItem)
        proc1.WaitForExit()
        proc1.Close()
        ЗагрНаСерверИУдаление(PathVremyanka & ListBox1.SelectedItem, ft & ListBox1.SelectedItem, PathVremyanka & ListBox1.SelectedItem)


        '    End If

        'End If

    End Sub

    Private Sub ListBox2_DoubleClick(sender As Object, e As EventArgs) Handles ListBox2.DoubleClick

        If ListBox2.SelectedIndex = -1 Then
            MessageBox.Show("Выберите документ для просмотра!", Рик, MessageBoxButtons.OK)
            Exit Sub
        End If

        'Dim ft = From x In dtPutiDokumentovAll.Rows Where x.item("IDСотрудник") = CType(Label6.Text, Integer) And x.item("ИмяФайла") = ListBox2.SelectedItem
        '         Select x.item("ИмяФайла")
        'Dim ft1 = From x In dtPutiDokumentovAll.Rows Where x.item("IDСотрудник") = CType(Label6.Text, Integer) And x.item("ИмяФайла") = ListBox2.SelectedItem
        '          Select x.item("Путь")
        'Dim _string As String = ft1(0).ToString & ft(0).ToString
        'Dim fl = From v In DirAll Where v.Key = ListBox2.SelectedItem Select v.Value

        'Dim _string As String = Replace(fl(0).ToString, ListBox2.SelectedItem & ",", "")
        '_string = Replace(_string, "[", "")
        '_string = Replace(_string, "]", "")
        '_string = Strings.Right(_string, _string.Length - 1) Not IsDBNull(x.Item("IDСотрудник"))
        'If _string = "" Then Exit Sub  x1.Item("IDСотрудник") IsNot Nothing

        Dim var = From x In dtPutiDokumentovAll.AsEnumerable Where Not IsDBNull(x.Item("IDСотрудник")) Select x

        Dim var1 = From x1 In var.AsEnumerable Where x1.Item("IDСотрудник") = CType(Label6.Text, Integer) _
                                            And x1.Item("ИмяФайла") = ListBox2.SelectedItem
                   Select x1.Item("ПолныйПуть")



        Dim _string As String = var1.FirstOrDefault

        ВыгрузкаФайловНаЛокалыныйКомп(_string, PathVremyanka & ListBox2.SelectedItem)

        Dim proc As Process = Process.Start(PathVremyanka & ListBox2.SelectedItem)
        proc.WaitForExit()
        proc.Close()

        ЗагрНаСерверИУдаление(PathVremyanka & ListBox2.SelectedItem, _string, PathVremyanka & ListBox2.SelectedItem)

        'Dim var = From x In dtPutiDokumentovAll.Rows Where x.item("IDСотрудник") = CType(Label6.Text, Integer)
        '          Select x.item("ИмяФайла") Distinct 'Distinct 'отбор без дублекатов


        'Dim ff As ListBox.SelectedIndexCollection = ListBox2.SelectedIndices


        'If Not ListBox2.SelectedIndex = -1 Then

        '    For Each p As Integer In ff
        '        Process.Start(FilesList(p))
        '    Next

        'End If
    End Sub


End Class