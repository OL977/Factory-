Option Explicit On
Imports System.Data.OleDb
Imports System.IO
Imports System.Threading
'Imports Microsoft.Office.Interop.Excel

Public Class ШтатноеСписки
    Public Da As New OleDbDataAdapter 'Адаптер
    Public Dsd As New DataSet 'Пустой набор записей
    Dim tbl As New DataTable
    Dim cb As OleDb.OleDbCommandBuilder

    Dim Организ, Год, Фиорукор As String
    Dim Mas() As String
    Dim datgen As Date
    Dim Должность, Разряд, Ставка, ТарифСтавка As String
    Dim СтрокаФайла As Task(Of String)
    Dim strAdres As String
    Dim Поток, Поток23 As Thread
    Dim ПотокЭксел, ПотокСборка As Thread
    Private Delegate Sub WriteText() ' Делегат 
    Dim combx1 As String
    Dim xlapp As Microsoft.Office.Interop.Excel.Application
    Dim xlworkbook As Microsoft.Office.Interop.Excel.Workbook
    Dim xlworksheet As Microsoft.Office.Interop.Excel.Worksheet
    Dim misvalue As Object = Reflection.Missing.Value
    Dim ds22 As DataRow()
    Dim df As String
    Dim КолСотр As Integer
    Dim Ds As DataTable
    Dim collname As New ArrayList()
    Dim t As Boolean = False

    Private Sub WriteTextSub() ' Процедура делегата
        If Not (Me.InvokeRequired) Then   ' Если запрос из родного (для элемента) потока, то просто записываем текст
            Me.Text = ComboBox1.Text
        Else ' Если доступ из другого потока, то используем делегат
            Me.Invoke(New WriteText(AddressOf WriteTextSub))
        End If
    End Sub
    'Private Sub WriteTextSub() ' Процедура делегата
    '    If Not (ComboBox1.InvokeRequired) Then   ' Если запрос из родного (для элемента) потока, то просто записываем текст
    '        ComboBox1.Text = ComboBox1.Text
    '    Else ' Если доступ из другого потока, то используем делегат
    '        ComboBox1.Invoke(New WriteText(AddressOf WriteTextSub))
    '    End If
    'End Sub
    Private Sub ШтатноеСписки_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MdiParent = MDIParent1
        Me.WindowState = FormWindowState.Maximized

        Год = Year(Now)
        'If Me.Прием_Load = vbTrue Then Form1.Load = False


        Me.ComboBox1.Items.Clear()
        For Each r As DataRow In СписокКлиентовОсновной.Rows
            Me.ComboBox1.Items.Add(r(0).ToString)
        Next
        t = True
        DateTimePicker1.Enabled = False
        DateTimePicker1.Text = Date.Now
        TextBox1.Text = "001"

    End Sub
    Public Sub Допзагр()
        xlapp = New Microsoft.Office.Interop.Excel.Application
        'xlworkbook = xlapp.Workbooks.Add(misvalue)
        Начало("Shtatnoe.xlsx")
        xlworkbook = xlapp.Workbooks.Add(firthtPath & "\Shtatnoe.xlsx")
        xlworksheet = xlworkbook.Sheets("Лист1")

    End Sub
    Private Sub DatTab22()
        'df = ""
        'df = Format(DateTimePicker1.Value, "MM\/dd\/yyyy")

        Dim list As New Dictionary(Of String, Object)
        list.Add("@НазвОрганиз", Организ)
        list.Add("@ДатаУвольнения", DateTimePicker1.Value.ToShortDateString)

        Dim ds = Selects(StrSql:="SELECT COUNT (Сотрудники.Фамилия)
FROM Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр
WHERE Сотрудники.НазвОрганиз=@НазвОрганиз AND ((КарточкаСотрудника.ДатаУвольнения)>=@ДатаУвольнения Or (КарточкаСотрудника.ДатаУвольнения) Is Null) AND КарточкаСотрудника.ДатаПриема<=@ДатаУвольнения", list)

        КолСотр = Nothing
        КолСотр = ds.Rows(0).Item(0)
        ds.Clear()

    End Sub
    Private Function ПровПоДевФамилии(ByVal d As Integer) As String
        Dim ds = dtSotrudnikiAll.Select("КодСотрудники=" & d & "")

        'Dim strsql As String = "SELECT ФамилияСтар,Фамилия,ДатаИзменения FROM Сотрудники WHERE КодСотрудники=" & d & ""
        'Dim ds As DataTable = Selects(strsql)

        If ds(0).Item("ФамилияСтар").ToString <> "" Then
            Dim dm As Date = ds(0).Item("ДатаИзменения")
            If dm > DateTimePicker1.Value Then
                Return ds(0).Item("ФамилияСтар").ToString
            End If
        End If
        Return ds(0).Item("Фамилия").ToString
    End Function
    Private Sub DatTab23()
        Dim list As New Dictionary(Of String, Object)
        list.Add("@НазвОрганиз", Организ)

        Dim ds = Selects(StrSql:="SELECT DISTINCT Штатное.Отдел
FROM Сотрудники INNER JOIN Штатное ON Сотрудники.КодСотрудники = Штатное.ИДСотр
WHERE Сотрудники.НазвОрганиз=@НазвОрганиз", list)
        'Erase Mas
        collname.Clear()

        For Each r As DataRow In ds.Rows
            collname.Add(r.Item(0).ToString())
        Next

        Dim рук, менедж As String
        For ir As Integer = 0 To collname.Count - 1
            Select Case collname(ir)
                Case "Руководители"
                    рук = collname(ir)
                Case "Специалисты"
                    менедж = collname(ir)
            End Select
        Next
        If менедж <> "" Then
            collname.Remove("Специалисты")
            collname.Add(менедж)
        End If
        If рук <> "" Then
            collname.Remove("Руководители")
            collname.Add(рук)
        End If
        If менедж <> "" Or рук <> "" Then
            collname.Reverse()
        End If

        Dim dfg(collname.Count - 1) As String
        For uy As Integer = 0 To collname.Count - 1
            dfg(uy) = collname(uy).ToString

        Next
    End Sub
    Private Function ДлПроц(ByVal gf As String) As Integer
        If gf.Length > 2 Then
            Return 1
        Else
            Return 0
        End If
    End Function
    Private Sub Сборка()
        Me.Cursor = Cursors.WaitCursor
        Dim i, j As Integer 'сохранение в эксель
        'Dim xlapp As Microsoft.Office.Interop.Excel.Application
        'Dim xlworkbook As Microsoft.Office.Interop.Excel.Workbook
        'Dim xlworksheet As Microsoft.Office.Interop.Excel.Worksheet
        'Dim misvalue As Object = Reflection.Missing.Value
        'xlapp = New Microsoft.Office.Interop.Excel.Application
        ''xlworkbook = xlapp.Workbooks.Add(misvalue)
        'xlworkbook = xlapp.Workbooks.Add(OnePath & "\ОБЩДОКИ\General\Shtatnoe.xlsx")
        'xlworksheet = xlworkbook.Sheets("Лист1")

        Dim StrSql2 As String
        Dim fdr As Integer
        Dim rnd As Integer
        Dim allstr As Integer = 18
        Dim ФИО() As String
        i = 18
        Dim sumall, ставall As Double

        If ПотокСборка.IsAlive Then
            ПотокСборка.Join()
        End If

        If Поток23.IsAlive Then
            Поток23.Join()
        End If

        'Выборка Среднемесячной нормы продолжительности рабочего времени за выбранный год
        Dim _Year As String = CType(DateTimePicker1.Value.Year, String)
        Dim _Норма As String
        Using dbcx As New DbAllDataContext
            Dim var = (From x In dbcx.СНПРВ.AsEnumerable
                       Where x.Год = _Year
                       Select x).FirstOrDefault
            If var IsNot Nothing Then
                _Норма = Replace(var.Норма, ",", ".")
            Else
                _Норма = 167.3
            End If
        End Using








        'If ds22.Rows(0).Item(3) = True Then
        '    Фиорукор = "ИП " & ФИОКорРук(ds22.Rows(0).Item(2), False)
        'Else
        '    Фиорукор = ФИОКорРук(ds22.Rows(0).Item(2), False)
        'End If

        Try
            xlworksheet.Cells(2, 1) = ds22(0).Item("ФормаСобств").ToString & " """ & Организ & """ "
        Catch ex As Exception
            Dim ПотокЭксел As New Thread(AddressOf Допзагр)
            ПотокЭксел.IsBackground = True
            ПотокЭксел.Start()
            ПотокЭксел.Join()
        Finally
            xlworksheet.Cells(2, 1) = ds22(0).Item("ФормаСобств").ToString & " """ & Организ & """ "
        End Try


        xlworksheet.Cells(8, 8) = "_______________________" & Фиорукор

        If Not Организ = "Итал Гэлэри Плюс" Then
            xlworksheet.Cells(7, 8) = ds22(0).Item("ДолжнРуководителя").ToString & " " & ФормСобствКор(ds22(0).Item("ФормаСобств")) & " """ & Организ & """ "
        Else
            With xlworksheet.Range(Cell1:="B4", Cell2:="F7")
                .Merge()
                .HorizontalAlignment = -4152
                .VerticalAlignment = -4107
                .WrapText = True
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = Microsoft.Office.Interop.Excel.XlOrder.xlDownThenOver
            End With
            xlworksheet.Cells(4, 2) = ds22(0).Item("ДолжнРуководителя").ToString
            xlworksheet.Cells(7, 8) = ФормСобствКор(ds22(0).Item("ФормаСобств")) & " """ & Организ & """ "
            xlworksheet.Cells(7, 8).HorizontalAlignment = -4152
        End If

        xlworksheet.Cells(13, 9) = "Вводится в действие с " & Format(DateTimePicker1.Value, "dd.MM.yyyy") & "г."

        'заполняем в шапке последнюю ячейку (СНПРВ)
        xlworksheet.Cells(14, 10) = "Часовая тарифная ставка (Фонд раб.времени - " & _Норма & ")"

        xlworksheet.Cells(9, 8) = Format(DateTimePicker1.Value, "dd.MM.yyyy")
        Dim КолОтделов As Integer = collname.Count
        КолСотр = КолСотр + (КолОтделов * 2) + 18

        For rn As Integer = 0 To collname.Count - 1
            With xlworksheet
                .Cells(rnd + 17, 1) = collname(rn).ToString
                .Cells(rnd + 17, 1).font.size = 11
                .Cells(rnd + 17, 1).Font.Bold = True
            End With

            Dim list As New Dictionary(Of String, Object)
            list.Add("@Отдел", collname(rn))
            list.Add("@НазвОрганиз", Организ)
            list.Add("@ДатаУвольнения", DateTimePicker1.Value.ToShortDateString)
            list.Add("@ДатаПриема", DateTimePicker1.Value.ToShortDateString)


            Dim ds2 = Selects(StrSql:="Select Штатное.Должность, Штатное.Разряд, КарточкаСотрудника.Ставка, Штатное.ТарифнаяСтавка, Штатное.ПовышОклПроц,
Сотрудники.Фамилия, Сотрудники.Имя, Сотрудники.Отчество, КарточкаСотрудника.IDСотр
FROM(Сотрудники INNER JOIN КарточкаСотрудника On Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр)
INNER JOIN Штатное On Сотрудники.КодСотрудники = Штатное.ИДСотр
WHERE Штатное.Отдел =@Отдел And Сотрудники.НазвОрганиз =@НазвОрганиз
AND ((КарточкаСотрудника.ДатаУвольнения)>=@ДатаУвольнения  Or (КарточкаСотрудника.ДатаУвольнения) Is Null) AND КарточкаСотрудника.ДатаПриема<=@ДатаПриема", list)
            'Dim ds2 As DataTable = Selects(StrSql2)

            allstr = allstr + ds2.Rows.Count
            ReDim ФИО(ds2.Rows.Count)

            For il As Integer = 0 To ds2.Rows.Count - 1
                ds2.Rows(il).Item(6) = Strings.Left(ds2.Rows(il).Item(6).ToString, 1)
                ds2.Rows(il).Item(7) = Strings.Left(ds2.Rows(il).Item(7).ToString, 1)

                'ПровПоДевФамилии(ds2.Rows(il).Item(8))

                ФИО(il) = ПровПоДевФамилии(ds2.Rows(il).Item(8)) & " " & ds2.Rows(il).Item(6).ToString & "." & ds2.Rows(il).Item(7).ToString & "."

                ПровПоПереводу(ds2.Rows(il).Item(8))

                If Ставка <> "" Then
                    ds2.Rows(il).Item(2) = Ставка
                End If

                If Разряд = "NO" Then
                    ds2.Rows(il).Item(1) = ""
                ElseIf Разряд <> "" And Not Разряд = "-" Then
                    ds2.Rows(il).Item(1) = Разряд
                End If

                If Должность <> "" Then
                    ds2.Rows(il).Item(0) = Должность
                End If

                If ТарифСтавка <> "" Then
                    ds2.Rows(il).Item(3) = ТарифСтавка
                End If


                Dim СтавкаТар As String = ПоискИзмененияСтавки(ds2.Rows(il).Item(8)) 'проверка по изменению разряда
                If Not СтавкаТар = "0" Then
                    ds2.Rows(il).Item(3) = СтавкаТар
                End If




                If ds2.Rows(il).Item(1).ToString <> "" And Not ds2.Rows(il).Item(1).ToString = "-" Then
                    ds2.Rows(il).Item(0) = ds2.Rows(il).Item(0).ToString & " " & ds2.Rows(il).Item(1).ToString & " p"
                End If
            Next

            xlworksheet.Range(Cell1:="A18", Cell2:="J" & КолСотр).Cells.Borders.LineStyle = 1 'рисуем границы


            For ddh As Integer = 0 To ds2.Rows.Count - 1 'основа для вставки в эксель

                With xlworksheet
                    .Cells(ddh + i, 1) = ddh + 1
                    .Cells(ddh + i, 3) = ФИО(ddh).ToString
                    .Cells(ddh + i, 2) = ds2.Rows(ddh).Item(0).ToString

                    .Cells(ddh + i, 4).NumberFormat = "0.00"
                    .Cells(ddh + i, 4) = ds2.Rows(ddh).Item(2).ToString
                    .Cells(ddh + i, 5).NumberFormat = "0.00"
                    .Cells(ddh + i, 5) = ds2.Rows(ddh).Item(3).ToString
                    'If ДлПроц(ds2.Rows(ddh).Item(4).ToString) = 0 Then 'поверка на наличие десятичных знаков в процентах
                    .Cells(ddh + i, 6).NumberFormat = "0.00"
                    'End If
                    .Cells(ddh + i, 6) = ds2.Rows(ddh).Item(4).ToString
                    .Cells(ddh + i, 7).NumberFormat = "0.00"
                    .Cells(ddh + i, 7) = "=ROUND(RC[-2]*RC[-1]/100,2)"
                    .Cells(ddh + i, 8).NumberFormat = "0.00"
                    .Cells(ddh + i, 8) = "=ROUND(RC[-3]+RC[-1],2)"
                    .Cells(ddh + i, 9).NumberFormat = "0.00"
                    .Cells(ddh + i, 9) = "=ROUND(RC[-5]*RC[-1],2)"
                    .Cells(ddh + i, 10).NumberFormat = "0.00"
                    'расчитываем часовую тарифную ставку в зависимости от года формирования штатного расписания
                    .Cells(ddh + i, 10) = "=ROUND(RC[-2]/" & _Норма & ",2)"
                End With
            Next
            i = i + ds2.Rows.Count + 2
            fdr = ds2.Rows.Count

            If Not ds2.Rows.Count = 0 Then
                With xlworksheet
                    .Range(Cell1:="A" & (allstr), Cell2:="J" & (allstr + 1)).ClearContents()
                    .Range(Cell1:="A" & (allstr), Cell2:="B" & (allstr)).Merge()
                    .Range(Cell1:="A" & (allstr), Cell2:="B" & (allstr)).HorizontalAlignment = -4131
                    .Cells(allstr, 1) = "Итого:"
                    .Cells(allstr, 1).Font.Bold = True
                    .Cells(allstr, 1).font.size = 11
                    .Cells(allstr, 1).font.Name = "Times New Roman"
                    .Cells(allstr, 4).Font.Bold = True
                    .Cells(allstr, 4).font.size = 11
                    .Cells(allstr, 4).font.Name = "Times New Roman"
                    .Cells(allstr, 4).FormulaR1C1 = "=ROUND(SUM(R[-" & fdr & "]C:R[-1]C),2)"
                    .Cells(allstr, 9).Font.Bold = True
                    .Cells(allstr, 9).font.size = 11
                    .Cells(allstr, 9).font.Name = "Times New Roman"
                    .Cells(allstr, 9).FormulaR1C1 = "=ROUND(SUM(R[-" & fdr & "]C:R[-1]C),2)"
                    .Range(Cell1:="A" & (allstr + 1), Cell2:="J" & (allstr + 1)).Interior.Color = 13434828
                    .Range(Cell1:="A" & (allstr + 1), Cell2:="J" & (allstr + 1)).Merge()
                    .Range(Cell1:="A" & (allstr + 1), Cell2:="J" & (allstr + 1)).Font.Bold = True
                    .Range(Cell1:="A" & (allstr + 1), Cell2:="J" & (allstr + 1)).HorizontalAlignment = -4108
                End With

                sumall = sumall + (xlworksheet.Cells(allstr, 9).value)
                ставall = ставall + (xlworksheet.Cells(allstr, 4).value)
            End If


            rnd = rnd + ds2.Rows.Count + 2
            allstr = allstr + 2
        Next

        With xlworksheet
            .Range(Cell1:="A" & (КолСотр), Cell2:="B" & (КолСотр)).Merge()
            .Range(Cell1:="A" & (КолСотр), Cell2:="B" & (КолСотр)).HorizontalAlignment = -4131
            .Cells(allstr, 1) = "Всего:"
            .Cells(allstr, 1).Font.Bold = True
            .Cells(allstr, 1).font.size = 12
            .Cells(allstr, 1).font.Name = "Times New Roman"
            .Cells(allstr, 9).value = sumall
            .Cells(allstr, 9).NumberFormat = "0.00"
            .Cells(allstr, 9).font.size = 12
            .Cells(allstr, 9).Font.Bold = True
            .Cells(allstr, 4).NumberFormat = "0.00"
            .Cells(allstr, 4).value = ставall
            .Cells(allstr, 4).font.size = 12
            .Cells(allstr, 4).Font.Bold = True
            .Cells(6, 8) = "заработной платы " & sumall & " рублей"
            .Cells(4, 8) = ставall & " штатных единиц"
            .Cells(11, 7) = "№" & TextBox1.Text
            .Cells(КолСотр + 3, 2).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = 1
            .Cells(КолСотр + 4, 2) = "(подпись)"
            .Cells(КолСотр + 4, 2).HorizontalAlignment = -4108
            .Range(Cell1:="E" & (КолСотр + 3), Cell2:="F" & (КолСотр + 3)).Merge()
            .Cells(КолСотр + 3, 5) = Фиорукор
            .Cells(КолСотр + 3, 5).Font.Bold = True
            .Cells(КолСотр + 3, 5).HorizontalAlignment = -4108
            .Range(Cell1:="E" & (КолСотр + 4), Cell2:="F" & (КолСотр + 4)).Merge()
            With .Range(Cell1:="E" & (КолСотр + 4), Cell2:="F" & (КолСотр + 4)).Cells

                .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = 1

            End With
        End With


        'If Not IO.Directory.Exists(OnePath & Организ & "\Штатное расписание\" & Now.Year) Then
        '    IO.Directory.CreateDirectory(OnePath & Организ & "\Штатное расписание\" & Now.Year)
        'End If
        'Организ & "\Штатное расписание\" & Now.Year & "\Штатное расписание " &

        'Try
        Dim name = TextBox1.Text & " от " & Format(DateTimePicker1.Value, "dd.MM.yyyy") & ".xlsx"
        xlworksheet.SaveAs(PathVremyanka & name)

        xlworkbook.Close()
        xlapp.Quit()
        Parallel.Invoke(Sub() releaseobject(xlapp))
        Parallel.Invoke(Sub() releaseobject(xlworkbook))
        Parallel.Invoke(Sub() releaseobject(xlworksheet))




        'Catch ex As Exception
        '    'KillExcel()
        '    'If MessageBox.Show("Такой файл уже существует! Заменить старый файл новым?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
        '    IO.File.Delete(OnePath & Организ & "\Штатное расписание\" & Now.Year & "\Штатное расписание " & TextBox1.Text & " от " & Format(DateTimePicker1.Value, "dd.MM.yyyy") & ".xlsx")
        '    xlworksheet.SaveAs(OnePath & Организ & "\Штатное расписание\" & Now.Year & "\Штатное расписание " & TextBox1.Text & " от " & Format(DateTimePicker1.Value, "dd.MM.yyyy") & ".xlsx")
        '    'End If
        'End Try
        strAdres = FTPString & name

        'Return strAdres

        ЗагрНаСерверИУдаление(PathVremyanka & name, FTPString & name, PathVremyanka & name)


        Dim Res As DialogResult = MessageBox.Show("Экспорт завершен!. При нажатии Да будет открыт сгенерированный файл, при нажатии Нет будет предложено сохранить файл.", Рик, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
        If Res = DialogResult.Yes Then

            Dim str As String = PathVremyanka & name
            ВыгрузкаФайловНаЛокалыныйКомп(FTPString & name, str)
            Process.Start(str)
            Me.Cursor = Cursors.Default
            Exit Sub

        ElseIf Res = DialogResult.No Then

            'Dim времянач1 As String = времянач.Replace("/", ".")
            'Dim времякон1 As String = времякон.Replace("/", ".")

            Dim Filename As String = ""
            'SaveFileDialog1.FileName = "Принятые сотрудники предприятия_ " & ComboBox1.Text & "(" & " c " & времянач1 & " по " & времякон1 & ")"
            SaveFileDialog1.Filter = "xls files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
            SaveFileDialog1.FilterIndex = 1
            SaveFileDialog1.RestoreDirectory = True

            If SaveFileDialog1.ShowDialog = DialogResult.OK Then
                Filename = SaveFileDialog1.FileName
                xlworkbook.SaveAs(SaveFileDialog1.FileName)
                MessageBox.Show("Данные сохранены!", Рик)
            Else
                Me.Cursor = Cursors.Default
                Exit Sub
            End If


        ElseIf Res = DialogResult.Cancel Then
            MessageBox.Show("Сохранение результатов экспорта отменено!")
        End If

        Me.Cursor = Cursors.Default






















        Me.Cursor = Cursors.Default


    End Sub
    Private Sub Обнул()
        Ставка = ""
        Разряд = ""
        Должность = ""
        ТарифСтавка = ""
    End Sub
    Private Function ТарифСтавка1(ByVal a As String, ByVal f As String, ByVal d As String) As String

        Dim list As New Dictionary(Of String, Object)
        list.Add("@Клиент", ComboBox1.Text)
        list.Add("@Отделы", a)
        list.Add("@Должность", f)
        list.Add("@Разряд", d)


        Dim ds = Selects(StrSql:="SELECT ШтСвод.ТарифнаяСтавка 
FROM ШтОтделы INNER JOIN ШтСвод ON ШтОтделы.Код = ШтСвод.Отдел
WHERE ШтОтделы.Клиент=@Клиент AND ШтОтделы.Отделы=@Отделы AND ШтСвод.Должность=@Должность AND ШтСвод.Разряд=@Разряд", list)

        If errds = 0 Then
            Return ds.Rows(0).Item(0).ToString
        End If
        Return "0"
    End Function
    Private Function ПоискИзмененияСтавки(ByVal d As Integer) As String

        Dim ds = dtShtatnoeAll.Select("ИДСотр=" & d & "")

        'Dim strsql As String = "SELECT Отдел,Должность,Разряд FROM Штатное WHERE ИДСотр=" & d & ""
        'Dim ds As DataTable = Selects(strsql)
        Dim list As New Dictionary(Of String, Object)
        list.Add("@Клиент", ComboBox1.Text)
        list.Add("@Отделы", ds(0).Item("Отдел").ToString)
        list.Add("@Должность", ds(0).Item("Должность").ToString)
        list.Add("@Разряд", ds(0).Item("Разряд").ToString)


        Dim ds1 = Selects(StrSql:="SELECT ШтСвод.КодШтСвод, ТарифнаяСтавка
FROM ШтОтделы INNER JOIN (ШтСвод INNER JOIN ШтСводИзмСтавка ON ШтСвод.КодШтСвод = ШтСводИзмСтавка.IDКодШтСвод) ON ШтОтделы.Код = ШтСвод.Отдел
WHERE ШтОтделы.Клиент=@Клиент AND ШтОтделы.Отделы=@Отделы AND ШтСвод.Должность=@Должность AND ШтСвод.Разряд=@Разряд", list)

        If errds = 1 Then
            Return "0"
        End If

        Dim list1 As New Dictionary(Of String, Object)
        list1.Add("@IDКодШтСвод", ds1.Rows(0).Item(0))
        list1.Add("@Дата", DateTimePicker1.Value.ToShortDateString)


        'Dim dat As String = Format(DateTimePicker1.Value, "MM\/dd\/yyyy")

        Dim ds2 = Selects(StrSql:= "SELECT Ставка, Дата FROM ШтСводИзмСтавка
WHERE IDКодШтСвод=@IDКодШтСвод AND Дата <= @Дата ORDER BY Дата DESC", list1)


        If errds = 1 Then
            Return ds1.Rows(0).Item(1).ToString
        Else
            Return ds2.Rows(0).Item(0).ToString
        End If



    End Function
    Private Sub Поиск2(ByVal idc As Integer)

        Dim ds1 = From x In dtPerevodAll Where x.Item("IDСотр") = idc Select x
        'Dim strsql1 As String = "SELECT * FROM Перевод WHERE IDСотр=" & idc & ""
        'Dim ds1 As DataTable = Selects(strsql1)
        For i As Integer = 0 To ds1.Count - 1
            'Dim day As Integer = DateDiff(DateInterval.Day, ds1.Rows(i).Item(2), ds1.Rows(i).Item(5))
            'If day = 1 Then
            '    If datgen = ds1.Rows(i).Item(5) Then
            '        Ставка = ds1.Rows(i).Item(12).ToString
            '        Должность = ds1.Rows(i).Item(4).ToString
            '        If ds1.Rows(i).Item(8).ToString <> "" And Not ds1.Rows(i).Item(8).ToString = "-" Then
            '            Разряд = ds1.Rows(i).Item(8).ToString
            '        Else
            '            Разряд = "NO"
            '        End If
            '        ТарифСтавка = ТарифСтавка1(ds1.Rows(i).Item(17).ToString, ds1.Rows(i).Item(4).ToString, ds1.Rows(i).Item(8).ToString)
            '        Exit Sub
            '    End If

            'Else

            Ставка = ""
            Должность = ""
            Разряд = ""
            ТарифСтавка = ""

            If datgen >= ds1(i).Item(2) And datgen < ds1(i).Item(5) Then
                Ставка = ds1(i).Item(11).ToString
                Должность = ds1(i).Item(3).ToString
                If ds1(i).Item(7).ToString <> "" And Not ds1(i).Item(7).ToString = "-" Then
                    Разряд = ds1(i).Item(7).ToString
                Else
                    Разряд = "NO"
                End If

                ТарифСтавка = ТарифСтавка1(ds1(i).Item(16).ToString, ds1(i).Item(3).ToString, ds1(i).Item(7).ToString)
                Exit Sub
            End If
            'End If
        Next

    End Sub
    Private Sub ПровПоПереводу(ByVal idc As Integer)
        Обнул()
        Dim ds = dtPerevodAll.Select("IDСотр=" & idc & "")
        'Dim strsql As String = "SELECT * FROM Перевод WHERE IDСотр=" & idc & ""
        'Dim ds As DataTable = Selects(strsql)
        datgen = Nothing
        datgen = CDate(DateTimePicker1.Value.ToShortDateString)
        If ds.Length = 0 Then Exit Sub


        If datgen >= ds(0).Item(5) Then
            Exit Sub
        Else
            Поиск2(idc)
        End If

    End Sub
    Private Sub ExtrExcel()
        'Dim ds, ds22 As New DataTable
        'Process.GetProcessesByName("Excel.Application")(0).Kill()


        If ComboBox1.Text = "" Then
            MessageBox.Show("Выберите организацию!")
            Exit Sub
        End If

        If DateTimePicker1.Text = "" Then
            MessageBox.Show("Введите дату штатного расписания!")
            Exit Sub
        End If

        If TextBox1.Text = "" Then
            MessageBox.Show("Выберите номер штатного расписания!")
            Exit Sub
        End If
        If Grid1.Rows.Count < 1 Then
            MessageBox.Show("Нет данных для экспорта!")
            Exit Sub
        End If

        ПотокСборка = New Thread(AddressOf DatTab22)
        ПотокСборка.IsBackground = True
        ПотокСборка.Start()

        Me.Cursor = Cursors.WaitCursor

        If ПотокЭксел.IsAlive Then
            ПотокЭксел.Join()
        End If


        Сборка()





        'If MessageBox.Show("Открыть файл?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
        '    Process.Start("C:\Users\Public\Downloads\Таблица.xlsx")
        'End If
    End Sub
    Private Sub releaseobject(ByVal obj As Object)
        Try
            Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ПотокЭксел = New Thread(AddressOf Допзагр)
        ПотокЭксел.IsBackground = True
        ПотокЭксел.Start()
        ExtrExcel()
        Parallel.Invoke(Sub() ЛистБокс())
    End Sub



    ''' <summary>
    ''' возвращает двумерный массив представляющий данные DataGridView
    ''' </summary>
    ''' <param name="DGV">DataGridView</param>
    ''' <param name="allData"> копировать всю таблицу (TRUE) или выделенный диапазон (FALSE)</param>
    ''' <param name="withoutC0">без поля заголовков строк (TRUE)</param>
    ''' <param name="withoutRN">без последней (пустой) строки (TRUE) [использовать для варианта allData=TRUE]</param>
    Private Function getDataDGV(ByVal DGV As DataGridView, ByVal allData As Boolean, ByVal withoutC0 As Boolean, ByVal withoutRN As Boolean) As String(,)
        If allData Then DGV.SelectAll()
        If Grid1.GetCellCount(DataGridViewElementStates.Selected) > 0 Then
            Try
                Clipboard.SetDataObject(Grid1.GetClipboardContent(), False)
                Dim data_object As IDataObject = Clipboard.GetDataObject()
                Dim ss As String = data_object.GetData(DataFormats.Text)
                Dim rr() As String = ss.Split(vbCrLf)
                Dim tt() As String = rr(0).Split(vbTab)
                Dim n As Integer = rr.Length
                If withoutRN Then n -= 1
                Dim m As Integer = tt.Length
                Dim cc(n - 1, m - 1) As String
                Dim k As Integer = 0
                If withoutC0 Then
                    ReDim cc(n - 1, m - 2)
                    k = 1
                End If
                For i = 0 To n - 1
                    tt = rr(i).Split(vbTab)
                    For j = k To m - 1
                        cc(i, j - k) = tt(j)
                    Next
                Next
                Return cc
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End If
        Return Nothing
    End Function
    ''' <summary>
    ''' добавляем таблицу в конец Rtf документа
    ''' </summary>
    ''' <param name="cc">матрица данных</param>
    ''' <param name="colWidth">размеры столбцов таблицы</param>
    ''' <param name="sRtf">имеющийся в документе Rtf текст</param>
    ''' <param name="fnt">шрифт для таблицы</param>
    Private Function addTableToRtf(ByVal cc(,) As String, ByVal colWidth() As Integer, ByVal sRtf As String, ByVal fnt As Font) As String
        Dim j As Integer = sRtf.LastIndexOf("}"c)
        sRtf = sRtf.Remove(j, 1)
        Return sRtf & TableToRTF(colWidth, cc, fnt) & "}"
    End Function
    ''' <summary>
    ''' Представление данных таблицы в Rtf формате.
    ''' </summary>
    ''' <param name="cw">размеры столбцов таблицы</param>
    ''' <param name="cc">матрица данных</param>
    ''' <param name="fnt">шрифт</param>
    Private Function TableToRTF(ByVal cw() As Integer, ByVal cc(,) As String, ByVal fnt As Font) As String
        Const CrLf As String = " "
        Dim nCol As Integer = cc.GetLength(1)
        Dim nRow As Integer = cc.GetLength(0)
        Dim Kx As Double = TwipsPerPixelXY()
        Dim fnName As String = fnt.Name
        Dim fnSize As Single = fnt.Size
        Dim IncCellWidth, CellMargin As Integer
        CellMargin = CInt(10 * Kx)
        Dim Result As String = "{" & CrLf
        Result &= "{\fonttbl{\f0\fnil " & fnName & "}}"
        Result &= "\trowd\f0\fs" + "20"  'Font.Size
        Result &= "\brdrs \trgaph" & CellMargin.ToString & CrLf & "\trqc" & CrLf 'центрируем таблицу на листе
        IncCellWidth = 0
        For i As Integer = 0 To nCol - 1
            IncCellWidth = CInt(IncCellWidth + cw(i) * Kx)
            Result &= "\cellx" + IncCellWidth.ToString & CrLf
        Next
        Dim h(nCol - 1) As String
        For j = 0 To nCol - 1
            h(j) = cc(0, j)
        Next
        Result &= "\b \intbl" & CrLf
        For j As Integer = 0 To nCol - 1
            Result &= h(j) & "\cell" & CrLf
        Next
        Result &= "\b0 \row" & vbCrLf
        For i As Integer = 1 To nRow - 1
            Result &= "\intbl" & CrLf
            For j As Integer = 0 To nCol - 1
                Result &= cc(i, j) & "\cell" & CrLf
            Next
            Result &= "\row" & CrLf
        Next
        Result &= "}"
        Return Result
    End Function
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Integer, ByVal nIndex As Integer) As Integer
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Integer) As Integer
    Private Function TwipsPerPixelXY() As Double
        Const LOGPIXELSX As Integer = 88
        Const LOGPIXELSY As Integer = 90
        Dim hwnd As Integer = Me.Handle 'GetDesktopWindow
        Dim hdc As Integer = GetDC(hwnd)
        Dim px As Integer = GetDeviceCaps(hdc, LOGPIXELSX)
        Dim py As Integer = GetDeviceCaps(hdc, LOGPIXELSY)
        Return 1440 / px
    End Function
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click


        '    Dim dd(,) As String = getDataDGV(Grid1, True, True, True) 'выгрузка данных в ричбокс
        '    Dim wCol() As Integer = {Grid1.Columns(0).Width, Grid1.Columns(1).Width, Grid1.Columns(2).Width,
        '        Grid1.Columns(3).Width, Grid1.Columns(4).Width, Grid1.Columns(5).Width, Grid1.Columns(6).Width,
        '        Grid1.Columns(7).Width, Grid1.Columns(8).Width}
        '    RichTextBox1.AppendText("Пример вывода таблицы" & vbCrLf)
        'RichTextBox1.Rtf = addTableToRtf(dd, wCol, RichTextBox1.Rtf, RichTextBox1.Font)


        'Копирование содержимого без заголовков
        'Grid1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText

        Grid1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText

        'Выделение содержимого DGV
        Grid1.SelectAll()

        'Помещаем в буфер обмена выделенные ячейки
        Clipboard.SetDataObject(Grid1.GetClipboardContent())
        'ОткрЭкселИзГрида()
        Dim путь1
        Dim j As String = "C:\Users\Public\Documents\dgv.txt"
        If IO.File.Exists(j) = False Then
            IO.File.Copy(OnePath & "\ОБЩДОКИ\General\dgv.txt", j)
            путь1 = j
        Else
            путь1 = j
        End If


        'Dim путь = "C:\Users\" & My.Computer.Name & "\Documents\dgv.html"
        'Dim путь1 = "C:\Users\" & My.Computer.Name & "\Documents\dgv.txt"
        'Dim путь2 = "C:\Users\" & My.Computer.Name & "\Documents\dgv.rtf"
        'Dim путь3 = "C:\Users\" & My.Computer.Name & "\Documents\dgv.xlsx"

        'Записываем текст из буфера обмена в файл
        Using writer As New StreamWriter(путь1, False, System.Text.Encoding.Unicode)
            writer.Write(Clipboard.GetText())
            'writer.Encoding.("UTF8")
            'writer.Write(Clipboard.GetText(TextDataFormat.html))
            'writer.Write(Clipboard.GetText(TextDataFormat.Rtf))
        End Using

        'Process.Start(путь3, Chr(34) & путь1 & Chr(34))
        Process.Start("excel.exe", Chr(34) & путь1 & Chr(34))

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click


        Grid1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        Grid1.SelectAll()
        Clipboard.SetDataObject(Grid1.GetClipboardContent())


        Начало("dgv.html")

        Dim путь = firthtPath & "\dgv.html"
        'Dim j As String = "C:\Users\Public\Documents\dgv.html"
        'If IO.File.Exists(j) = False Then
        '    IO.File.Copy(OnePath & "\ОБЩДОКИ\General\dgv.html", j)
        '    путь = j
        'Else
        '    путь = j
        'End If


        Using writer As New StreamWriter(путь, False, System.Text.Encoding.Unicode)
            writer.Write(Clipboard.GetText(TextDataFormat.Html))
        End Using
        Process.Start(путь)
    End Sub

    Private Sub ЛистБокс()

        ListBox1.Items.Clear()


        Dim list = listFluentFTP(ComboBox1.Text & "/Штатное расписание/" & Now.Year)
        'Dim Files2() As String = Nothing
        'Try
        '    Files2 = IO.Directory.GetFiles(OnePath & ComboBox1.Text & "\Штатное расписание\" & Now.Year, "*.xls*", IO.SearchOption.TopDirectoryOnly)
        'Catch ex As Exception
        '    'MessageBox.Show("В этому году нет такой папки!", Рик)
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
        For Each x In list
            ListBox1.Items.Add(x.ToString) ' На ListBox2
        Next



    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Cursor = Cursors.WaitCursor
        If t = False Then
            SortGrid1()
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        'Me.Cursor = Cursors.WaitCursor
        'If t = False Then
        '    SortGrid1()
        'End If
        'Me.Cursor = Cursors.Default
    End Sub

    Private Sub SortGrid1()
        'Dim StrSql1 As String

        '        StrSql1 = "Select Сотрудники.НазвОрганиз, Штатное.Должность As [Должность], Сотрудники.ФИОСборное As [ФИО], КарточкаСотрудника.Ставка, Штатное.ТарифнаяСтавка As [Тарифная ставка(оклад), руб],
        'Штатное.ПовышОклПроц as [Повышение оклада, %], Штатное.ПовышОклРуб as [Повышение оклада, руб], Штатное.РасчДолжностнОклад as [Расчетный должностной оклад], Штатное.ФонОплатыТруда as [ФОТ],
        'КарточкаСотрудника.ДатаУвольнения as [Дата увольнения]
        'FROM (Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр) INNER JOIN Штатное ON Сотрудники.КодСотрудники = Штатное.ИДСотр
        'WHERE Сотрудники.НазвОрганиз='" & Организ & "' ORDER BY Сотрудники.ФИОСборное"
        'df = ""
        'df = Format(DateTimePicker1.Value, "MM\/dd\/yyyy")

        Dim list As New Dictionary(Of String, Object)
        list.Add("@НазвОрганиз", Организ)
        list.Add("@ДатаУвольнения", DateTimePicker1.Value.ToShortDateString)
        list.Add("@ДатаПриема", DateTimePicker1.Value.ToShortDateString)

        Dim tbl = Selects(StrSql:="SELECT Сотрудники.НазвОрганиз, Штатное.Должность as [Должность], Сотрудники.ФИОСборное as [ФИО], КарточкаСотрудника.Ставка, Штатное.ТарифнаяСтавка as [Тарифная ставка(оклад),руб],
        Штатное.ПовышОклПроц as [Повышение оклада, %], Штатное.ПовышОклРуб as [Повышение оклада, руб], Штатное.РасчДолжностнОклад as [Расчетный должностной оклад], Штатное.ФонОплатыТруда as [ФОТ],
        КарточкаСотрудника.ДатаУвольнения as [Дата увольнения], КарточкаСотрудника.IDСотр, Штатное.Разряд, Сотрудники.Имя, Сотрудники.Отчество
        FROM (Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр) INNER JOIN Штатное ON Сотрудники.КодСотрудники = Штатное.ИДСотр
        WHERE Сотрудники.НазвОрганиз =@НазвОрганиз And ((КарточкаСотрудника.ДатаУвольнения)>=@ДатаУвольнения Or (КарточкаСотрудника.ДатаУвольнения) Is Null) 
        AND КарточкаСотрудника.ДатаПриема<=@ДатаПриема ORDER BY Сотрудники.ФИОСборное", list)




        'ds.WriteXml("C:\Users\" & My.Computer.Name & "\Desktop\tabl.xml")'xml
        'ds.ReadXml("C:\Users\" & My.Computer.Name & "\Desktop\tabl.xml")


        For i As Integer = 0 To tbl.Rows.Count - 1

            tbl.Rows(i).Item(2) = ПровПоДевФамилии(tbl.Rows(i).Item(10)) & " " & tbl.Rows(i).Item(12).ToString & " " & tbl.Rows(i).Item(13).ToString


            ПровПоПереводу(tbl.Rows(i).Item(10))

            If Ставка <> "" Then
                tbl.Rows(i).Item(3) = Ставка
                'tbl.Rows(i).Item(4) = Ставка
            End If

            If Разряд = "NO" Then
                tbl.Rows(i).Item(11) = ""
            ElseIf Разряд <> "" And Not Разряд = "-" Then
                tbl.Rows(i).Item(11) = Разряд
            End If

            If Должность <> "" Then
                tbl.Rows(i).Item(1) = Должность & " " & Разряд

            End If

            If ТарифСтавка <> "" Then
                tbl.Rows(i).Item(4) = ТарифСтавка
            End If


            Dim СтавкаТар As String = ПоискИзмененияСтавки(tbl.Rows(i).Item(10)) 'проверка по изменению разряда
            If Not СтавкаТар = "0" Then
                tbl.Rows(i).Item(4) = СтавкаТар
            End If


        Next

        Grid1.DataSource = tbl
        GridView(Grid1)
        Dim strikethrough_style As New DataGridViewCellStyle

        strikethrough_style.Font = New Font("Times New Roman", 10, FontStyle.Regular)
        strikethrough_style.BackColor = Color.White
        For Each row As DataGridViewRow In Grid1.Rows
            For i = 0 To Grid1.Columns.Count - 1
                row.Cells(i).Style = strikethrough_style
            Next
        Next

        Grid1.Columns(0).Visible = False
        'Grid1.Columns(7).Visible = False
        'Grid1.Columns(4).Width = 60
        'Grid1.Columns(5).Width = 100
        Grid1.Columns(1).Width = 150
        Grid1.Columns(2).Width = 200
        Grid1.Columns(10).Visible = False
        Grid1.Columns(11).Visible = False
        Grid1.Columns(12).Visible = False
        Grid1.Columns(13).Visible = False
        'Grid1.Columns(3).Width = 30

        Grid1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Grid1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Grid1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft 'выравниваем текст в ячейках по центру
        Grid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    End Sub




    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        t = False
        Организ = ComboBox1.Text
        СозданиепапкиНаСервере(ComboBox1.Text & "/Штатное расписание/" & Now.Year)
        Parallel.Invoke(Sub() ЛистБокс())
        Me.Cursor = Cursors.WaitCursor
        SortGrid1()


        DateTimePicker1.Enabled = True

        ds22 = dtClientAll.Select("НазвОрг='" & Организ & "'")
        '        Dim StrSql5 As String = "Select Клиент.ФормаСобств, Клиент.ДолжнРуководителя, Клиент.ФИОРуководителя, Клиент.РукИП 
        'From Клиент WHERE НазвОрг='" & Организ & "'"


        If ds22(0).Item("РукИП") = "True" Then
            Фиорукор = "ИП " & ФИОКорРук(ds22(0).Item("ФИОРуководителя"), False)
        Else
            Фиорукор = ФИОКорРук(ds22(0).Item("ФИОРуководителя"), False)
        End If

        Поток23 = New Thread(AddressOf DatTab23)
        Поток23.IsBackground = True
        Поток23.Start()

        Me.Cursor = Cursors.Default

    End Sub

    Private Sub TextBox1_KeyDown_1(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown, DateTimePicker1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Button1.Focus()

            Dim pl As String
            If TextBox1.Text <> "" Then

                Dim i As Integer = CInt(TextBox1.Text)
                Select Case i
                    Case < 10
                        pl = Str(i)
                        TextBox1.Text = "00" & i

                    Case 10 To 99
                        pl = Str(i)
                        TextBox1.Text = "0" & i
                End Select


            End If
        End If
    End Sub

    Private Sub DateTimePicker1_CloseUp(sender As Object, e As EventArgs) Handles DateTimePicker1.CloseUp
        ПотокСборка = New Thread(AddressOf DatTab22)
        ПотокСборка.IsBackground = True
        ПотокСборка.Start()
    End Sub

    Private Sub Grid1_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles Grid1.ColumnHeaderMouseClick
        Grid1.Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
        Grid1.Columns(4).SortMode = DataGridViewColumnSortMode.NotSortable
        Grid1.Columns(5).SortMode = DataGridViewColumnSortMode.NotSortable
        Grid1.Columns(6).SortMode = DataGridViewColumnSortMode.NotSortable
        Grid1.Columns(7).SortMode = DataGridViewColumnSortMode.NotSortable
        Grid1.Columns(8).SortMode = DataGridViewColumnSortMode.NotSortable
    End Sub

    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick
        If Not ListBox1.SelectedIndex = -1 Then
            ВыгрузкаФайловНаЛокалыныйКомп(ComboBox1.Text & "\Штатное расписание\" & Now.Year & "\" & ListBox1.SelectedItems(0), PathVremyanka & ComboBox1.Text & "\Штатное расписание\" & Now.Year & "\" & ListBox1.SelectedItems(0))
            Process.Start(PathVremyanka & ComboBox1.Text & "\Штатное расписание\" & Now.Year & "\" & ListBox1.SelectedItems(0))
        End If
    End Sub
End Class