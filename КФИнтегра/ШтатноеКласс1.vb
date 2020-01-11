Option Explicit On
Imports System.Data.OleDb
Public Class ШтатноеКласс1
    Public Da As New OleDbDataAdapter 'Адаптер
    'Public Ds As New DataSet 'Пустой набор записей
    Dim tbl As New DataTable
    Dim ds As DataTable
    Dim cb As OleDb.OleDbCommandBuilder
    Dim Рик As String = "ООО РикКонсалтинг"
    Dim Год, Организ, Процент, ТарСтавка, thb0, thb, StrSql As String 'Разряд,Отдел, Должность,
    Dim s, s2, se, ip, mas, изменен, srt, КодDBC, ГлКод, КодДолжн As Integer
    Dim Отд, Дол, Раз, ТСтавка, ПовышПроц As String
    Public v As Boolean = False
    Public FT As Boolean = False
    Dim СумНов As String
    Dim fnm9 As Integer
    Dim btnclick As Integer
    Dim timtick As String
    Private Sub ШтатноеКласс1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MdiParent = MDIParent1
        WindowState = FormWindowState.Maximized


        Год = Year(Now)
        'If Me.Прием_Load = vbTrue Then Form1.Load = False

        'If Not ComboBox1.Items Is Nothing Or ComboBox1.Items.Count > 0 Then
        '    ComboBox1.Items.Clear()
        'End If

        For Each r As DataRow In СписокКлиентовОсновной.Rows
            ComboBox1.Items.Add(r(0).ToString)
        Next
        'MaskedTextBox1.Text = DateTime.Now.ToString("dd.MM.yyyy")


        For i As Integer = 0 To Grid1.Rows.Count - 1
            For y As Integer = 0 To Grid1.Columns.Count - 1
                Grid1.Item(y, i).Style.Font = New Font("times new roman", 11)
            Next
        Next
        FT = True
        DateTimePicker1.Value = Date.Now
        DateTimePicker1.Enabled = False
    End Sub

    Private Sub ВставкаВШтСводИзмСтавка(ByVal idsotr As Integer)

        Dim list As New Dictionary(Of String, Object)
        list.Add("@Отдел", idsotr)
        list.Add("@Должность", Trim(TextBox5.Text))
        list.Add("@Разряд", Trim(TextBox3.Text))
        list.Add("@ТарифнаяСтавка", Trim(TextBox1.Text))



        Dim ds = Selects(StrSql:="SELECT КодШтСвод FROM ШтСвод WHERE Отдел=@Отдел AND Должность=@Должность AND Разряд=@Разряд AND ТарифнаяСтавка=@ТарифнаяСтавка", list)


        Updates(stroka:="INSERT INTO ШтСводИзмСтавка(IDКодШтСвод,Дата,Ставка) 
VALUES(" & ds.Rows(0).Item(0) & ", '01.01.1990','" & Trim(TextBox1.Text) & "')", list, "ШтСводИзмСтавка")


        Updates(stroka:="INSERT INTO ШтСводИзмСтавка(IDКодШтСвод,Дата,Ставка) 
VALUES(" & ds.Rows(0).Item(0) & ", '" & DateTimePicker1.Value.ToShortDateString & "','" & Trim(TextBox1.Text) & "')")
        dtShtatnoeSvodnoeIzmenStavka()

    End Sub
    Private Sub Добавить()
        If MessageBox.Show("Добавить данные?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.Cancel Then
            Exit Sub
        End If

        Чист()
        Dim ds = Selects(StrSql:="SELECT ШтСвод.КодШтСвод FROM ШтОтделы INNER JOIN ШтСвод ON ШтОтделы.Код = ШтСвод.Отдел
WHERE ШтОтделы.Клиент='" & ComboBox1.Text & "' AND ШтОтделы.Отделы='" & Trim(TextBox4.Text) & "' 
AND ШтСвод.Должность ='" & Trim(TextBox5.Text) & "' AND ШтСвод.Разряд='" & Trim(TextBox3.Text) & "'")


        Try
            If IsNumeric(ds.Rows(0).Item(0)) = True Then
                MessageBox.Show("В организации " & ComboBox1.Text & "уже есть отдел " & Trim(TextBox4.Text) & " с должностью " & Trim(TextBox5.Text) & " и разрядом " & Trim(TextBox3.Text) & "." & vbCrLf & "Добавить такой же отдел с такой-же должностью и разрядом невозможно!", Рик)
                Exit Sub
            End If
        Catch ex As Exception

        End Try


        Dim ds1 = dtShtatnoeOtdelyAll.Select("Клиент='" & ComboBox1.Text & "' AND Отделы='" & Trim(TextBox4.Text) & "'")

        'Чист()
        'StrSql = "SELECT Код FROM ШтОтделы WHERE ШтОтделы.Клиент='" & ComboBox1.Text & "' AND ШтОтделы.Отделы='" & Trim(TextBox4.Text) & "'"
        'ds = Selects(StrSql)
        Try
            If IsNumeric(ds1(0).Item("Код")) = True Then

                Updates(stroka:="INSERT INTO ШтСвод(Отдел, Должность,Разряд,ТарифнаяСтавка,ПовышениеПроц,ДолжИнструкция)
VALUES(" & ds1(0).Item("Код") & ",'" & Trim(TextBox5.Text) & "','" & Trim(TextBox3.Text) & "','" & CType(Math.Round(CDbl(TextBox1.Text), 2), String) & "','" & Trim(TextBox2.Text) & "','False')")
                dtShtatnoeSvodnoe()

                ВставкаВШтСводИзмСтавка(ds1(0).Item("Код"))


                MessageBox.Show("Данные добавлены!", Рик)
            End If
        Catch ex As Exception

            Updates(stroka:="INSERT INTO ШтОтделы(Клиент, Отделы) VALUES('" & ComboBox1.Text & "','" & Trim(TextBox4.Text) & "')")
            dtShtatnoeOtdely()


            Dim ds2 = Selects(StrSql:="SELECT ШтОтделы.Код 
FROM ШтОтделы
WHERE ШтОтделы.Клиент='" & ComboBox1.Text & "' AND ШтОтделы.Отделы='" & Trim(TextBox4.Text) & "'")

            Dim idsotr As Integer = ds2.Rows(0).Item(0)

            Updates(stroka:="INSERT INTO ШтСвод(Отдел,Должность,Разряд,ТарифнаяСтавка,ПовышениеПроц,ДолжИнструкция)
VALUES(" & idsotr & ",'" & Trim(TextBox5.Text) & "','" & Trim(TextBox3.Text) & "','" & Trim(TextBox1.Text) & "','" & Trim(TextBox2.Text) & "','False')")
            dtShtatnoeSvodnoe()

            ВставкаВШтСводИзмСтавка(idsotr)


            MessageBox.Show("Данные добавлены!", Рик)
        End Try
        Dim gf As String = TextBox5.Text & ". Разряд" & TextBox3.Text & ".Отдел " & TextBox4.Text
        Статистика1("Должность " & gf, "Данные добавлены", ComboBox1.Text)

    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If TextBox4.Text = "" Or TextBox5.Text = "" Then
            MessageBox.Show("Поле отдел и должность не могут быть пустыми!")
            Exit Sub
        End If
        btnclick = 1
        TextBox1.Text = CType(Math.Round(CDbl(Replace(TextBox1.Text, ".", ",")), 2), String)
        Добавить()
        Очистка()
        ВыборСтавкиПоДате()
        'Refreshgrid()

    End Sub
    Private Function ПроверкаНаналичиеСуществДанных(ByVal дата As String, ByVal ставка As String) As Boolean
        Dim strsql As String = "SELECT " & ставка & " FROM ШтСвод WHERE КодШтСвод=" & КодДолжн & ""
        Dim ds As DataTable = Selects(strsql)

        If ds.Rows(0).Item(0).ToString = "" Then
            Return False
        Else
            Return True
        End If
    End Function
    Private Sub ИзменСтавкиПриказ()
        If MessageBox.Show("Изменить ставку с учетом даты" & vbCrLf & DateTimePicker1.Value & "?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Exit Sub
        End If





        'Dim дата As String = Replace(Strings.Left(CType(MaskedTextBox1.Text, String), 10), ".", "") 'перевод числа в строку

        'Dim ставка As String = "СТ" & дата
        'Dim strsql1 As String = "ALTER TABLE ШтСвод ADD COLUMN " & дата & " DATETIME, " & ставка & " TEXT(255)" 'добавление столбца в базу

        'Dim strsql As String = "Select * FROM ШтСвод"
        'Dim ds As DataTable = Selects(strsql)
        'Dim ПровДаты As Boolean = False

        'For Each r As DataColumn In ds.Columns
        '    If r.ColumnName = дата Then
        '        ПровДаты = True
        '    End If
        'Next


        'If ПровДаты = True Then
        '    If ПроверкаНаналичиеСуществДанных(дата, ставка) = True Then
        '        If MessageBox.Show("Заменить старые данные?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
        '            Exit Sub
        '        End If
        '    End If
        '    Обнданных(дата, ставка)
        'Else
        '    Updates(strsql1)
        '    Обнданных(дата, ставка)
        'End If



        'Dim d1 As String = Strings.Left(дата, 2) 'перевод числа в дату
        'Dim d2 As String = Strings.Left(дата, 4)
        'Dim d3 As String = Strings.Right(дата, 4)
        'd2 = Strings.Right(d2, 2) & "."
        'd1 = d1 & "."
        'дата = d1 & d2 & d3

    End Sub
    '    Private Sub Обнданных(ByVal дата As String, ByVal ставка As String)
    '        Dim strsql2 As String = "UPDATE ШтСвод SET " & дата & " = '" & MaskedTextBox1.Text & "', " & ставка & " = '" & Replace(TextBox1.Text, ".", ",") & "'
    'WHERE КодШтСвод=" & КодДолжн & ""
    '        Updates(strsql2)
    '    End Sub
    Private Sub Изменить()

        Dim list As New Dictionary(Of String, Object)
        list.Add("@КодШтСвод", КодДолжн)
        list.Add("@Код", ГлКод)

        Try

            Updates(stroka:="Update ШтСвод Set Должность='" & Trim(TextBox5.Text) & "', Разряд='" & Trim(TextBox3.Text) & "',ТарифнаяСтавка='" & Trim(TextBox1.Text) & "',
ПовышениеПроц='" & Trim(TextBox2.Text) & "'
WHERE ШтСвод.КодШтСвод =@КодШтСвод", list, "ШтСвод")

            Updates(stroka:="Update ШтОтделы Set Отделы='" & Trim(TextBox4.Text) & "' WHERE ШтОтделы.Код =@Код", list, "ШтОтделы")


        Catch ex As Exception
            MessageBox.Show("В базе нет данных относительно вашего запроса!", Рик)
        End Try
        Dim gf As String = TextBox4.Text & " " & TextBox5.Text & " " & TextBox3.Text

        Статистика1("Изменение " & gf, "Изменены данные в должности или отеделе или разряде", ComboBox1.Text)


        'ОбнВСтавкеИстория()

        'If errds = 1 Then
        '    If MessageBox.Show("Данный отдел и должность не существуют в базе. Добавить данные?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
        '        Чист()
        '        StrSql = "INSERT INTO ШтОтделы(Клиент, Отделы) VALUES('" & ComboBox1.Text & "','" & TextBox4.Text & "')"
        '        Updates(StrSql)

        '        Чист()
        '        StrSql = "SELECT ШтОтделы.Код FROM ШтОтделы WHERE ШтОтделы.Клиент='" & ComboBox1.Text & "' AND ШтОтделы.Отделы='" & TextBox4.Text & "'"
        '        ds = Selects(StrSql)
        '        Dim idsotr As Integer = ds.Rows(0).Item(0)

        '        StrSql = "INSERT INTO ШтСвод(Отдел, Должность,Разряд,ТарифнаяСтавка,ПовышениеПроц) VALUES(" & idsotr & ",'" & TextBox5.Text & "','" & TextBox3.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "')"
        '        Updates(StrSql)


        '        MessageBox.Show("Данные добавлены!", Рик)
        '    End If
        'Else



        'Чист()
        '    StrSql = "Update ШтОтделы Set ШтОтделы.Отделы='" & TextBox4.Text & "' WHERE ШтОтделы.Код=" & idsotr & ""
        '    Updates(StrSql)



        'End If



    End Sub
    Private Sub ОбнВСтавкеИстория()

        Dim strsql1 As String = "SELECT * FROM ШтСвод WHERE КодШтСвод=" & КодДолжн & ""
        Dim ds1 As DataTable = Selects(strsql1)

        'Dim strsql4 As String = "SELECT DISTINCT КодШтСвод FROM ШтСводИзмСтавка WHERE КодШтСвод=" & КодДолжн & "" 'собираем данные только по нключам,  сортируем
        'Dim ds4 As DataTable = Selects(strsql4)


        Dim strsql As String = "SELECT * FROM ШтСводИзмСтавка WHERE IDКодШтСвод=" & КодДолжн & ""
        Dim ds As DataTable = Selects(strsql)

        If errds = 1 Then
            Dim strsql2 As String = "INSERT INTO ШтСводИзмСтавка(IDКодШтСвод, Дата, Ставка)
VALUES(" & КодДолжн & ",'" & DateTimePicker1.Value.ToShortDateString & "','" & Replace(TextBox1.Text, ".", ",") & "')"
            Updates(strsql2) 'создаем новую строку со ставкой 

            Dim strsql3 As String = "INSERT INTO ШтСводИзмСтавка(IDКодШтСвод, Дата, Ставка)
VALUES(" & КодДолжн & ",'01.01.1990','" & ds1.Rows(0).Item(4).ToString & "')"
            Updates(strsql3) 'копируем старую ставку 

        Else
            Dim ik As Boolean = False
            For Each r As DataRow In ds.Rows
                If Strings.Left(r.ItemArray(2), 10) = Strings.Left(DateTimePicker1.Value, 10) Then
                    Dim strsql3 As String = "UPDATE ШтСводИзмСтавка SET Ставка='" & Replace(TextBox1.Text, ".", ",") & "', Дата= '" & DateTimePicker1.Value.ToShortDateString & "'
WHERE IDКодШтСвод =" & КодДолжн & ""
                    Updates(strsql3)
                    ik = True
                End If
            Next
            If ik = False Then
                Dim strsql2 As String = "INSERT INTO ШтСводИзмСтавка(IDКодШтСвод, Дата, Ставка)
VALUES(" & КодДолжн & ",'" & DateTimePicker1.Value.ToShortDateString & "','" & Replace(TextBox1.Text, ".", ",") & "')"
                Updates(strsql2)
            End If
        End If
        СумНов = Replace(TextBox1.Text, ".", ",")

        dtShtatnoeSvodnoeIzmenStavka()

    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If TextBox4.Text = "" Or TextBox5.Text = "" Then
            MessageBox.Show("Поле отдел и должность не могут быть пустыми!")
            Exit Sub
        End If

        If MessageBox.Show("Изменить данные?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor
        btnclick = 1
        TextBox1.Text = CType(Math.Round(CDbl(Replace(TextBox1.Text, ".", ",")), 2), String)
        ОбнВСтавкеИстория()


        If Date.Now.ToShortDateString <= DateTimePicker1.Value.ToShortDateString Then
            Изменить()
            ВыборСтавкиПоДате()
            'Refreshgrid()
        Else
            ВыборСтавкиПоДате()
        End If
        Me.Cursor = Cursors.Default
        MessageBox.Show("Данные изменены!", Рик)

    End Sub
    Private Function ВыборкаБезДублей(ByVal DateEx As String, ByVal com1 As String) As List(Of String)

        Dim objlist As New List(Of String)()

        Dim strsql As String = "SELECT DISTINCT ШтСводИзмСтавка.IDКодШтСвод
FROM ШтОтделы INNER JOIN (ШтСвод INNER JOIN ШтСводИзмСтавка ON ШтСвод.КодШтСвод = ШтСводИзмСтавка.IDКодШтСвод) ON ШтОтделы.Код = ШтСвод.Отдел
WHERE Дата <= '" & DateEx & "' AND  ШтОтделы.Клиент='" & com1 & "'"
        Dim ds As DataTable = Selects(strsql)

        If ds.Rows.Count = 0 Then
            objlist.Add("0")
            Return objlist
        End If


        For Each r As DataRow In ds.Rows 'собрали только коды для поиска
            objlist.Add(r.Item(0).ToString)
        Next
        Return objlist

    End Function
    Public Sub ВыборСтавкиПоДате()

        Dim list As New Dictionary(Of String, Object)
        list.Add("@Клиент", Организ)


        Организ = ComboBox1.Text
        Dim ds1 = Selects(StrSql:="SELECT ШтОтделы.Код, ШтОтделы.Клиент, ШтОтделы.Отделы, ШтСвод.Должность, ШтСвод.Разряд, 
ШтСвод.ТарифнаяСтавка as Ставка, ШтСвод.ПовышениеПроц as Процент, ШтСвод.КодШтСвод, ШтСвод.ДолжИнструкция as [Инструкц]
FROM ШтОтделы INNER JOIN ШтСвод ON ШтОтделы.Код = ШтСвод.Отдел
        WHERE ШтОтделы.Клиент =@Клиент", list)

        Dim DateEx As String = Replace(Format(DateTimePicker1.Value, "yyyy\/MM\/dd"), "/", "")


        Dim ds = Selects(StrSql:="SELECT ШтСводИзмСтавка.* 
FROM ШтОтделы INNER JOIN (ШтСвод INNER JOIN ШтСводИзмСтавка ON ШтСвод.КодШтСвод = ШтСводИзмСтавка.IDКодШтСвод) ON ШтОтделы.Код = ШтСвод.Отдел
WHERE Дата<='" & DateEx & "' AND ШтОтделы.Клиент=@Клиент", list)


        If ds.Rows.Count = 0 Then
            Очистка()
            'ВыборСтавкиПоДате()
            Refreshgrid()
            Exit Sub
        End If

        Dim lst2 As New List(Of String)()
        lst2 = ВыборкаБезДублей(DateEx, ComboBox1.Text)

        For ib As Integer = 0 To ds1.Rows.Count - 1 ' основной модуль по созданию таблицы
            If lst2.Contains(ds1.Rows(ib).Item(7).ToString) Then
                For i As Integer = 0 To ds.Rows.Count - 1
                    If ds.Rows(i).Item(1) = ds1.Rows(ib).Item(7) Then

                        Dim strsql3 As String = "SELECT * FROM ШтСводИзмСтавка
WHERE IDКодШтСвод=" & ds.Rows(i).Item(1) & " AND Дата <= '" & DateEx & "' ORDER BY Дата DESC"
                        Dim ds3 As DataTable = Selects(strsql3)
                        If errds = 1 Then
                            Dim strsql31 As String = "SELECT * FROM ШтСводИзмСтавка
WHERE IDКодШтСвод=" & ds.Rows(i).Item(1) & " AND Дата >= '" & DateEx & "' ORDER BY Дата"
                            Dim ds31 As DataTable = Selects(strsql31)
                            ds1.Rows(ib).Item(5) = ds31.Rows(0).Item(3).ToString
                        Else
                            ds1.Rows(ib).Item(5) = ds3.Rows(0).Item(3).ToString
                        End If

                        Exit For
                    End If
                Next
            End If
        Next

        Очистка()

        For x As Integer = 0 To ds1.Rows.Count - 1
            If ds1.Rows(x).Item(8).ToString = "True" Then
                ds1.Rows(x).Item(8) = "Есть"
            Else
                ds1.Rows(x).Item(8) = "Нет"
            End If
        Next



        ds1.DefaultView.Sort = "Отделы" & " DESC" 'сортировка столбца по возрастанию datatable

        Grid1.DataSource = ds1

        Grid1.Columns(1).Visible = False
        Grid1.Columns(7).Visible = False
        Grid1.Columns(0).Visible = False
        'Grid1.Columns(4).Width = 60
        'Grid1.Columns(5).Width = 100
        'Grid1.Columns(1).Width = 150
        Grid1.Columns(2).Width = 200
        Grid1.Columns(3).Width = 200
        Grid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect ' выделяет всю строку в grid1
        GridView(Grid1)
        'Grid1.Columns(6).Width = 60

        'Grid1.Rows(1).Cells(3).Selected = True
        'Grid1_CellClick(Grid1, New DataGridViewCellEventArgs(<b>3</b>, <b>1</b>))
        'Acti()
        cb = New OleDb.OleDbCommandBuilder(Da)
        s = Grid1.Rows.Count - 1
        изменен = 0
        'NumberAllRows()

    End Sub
    Public Sub ВыборСтавкиПоДате(ByVal com1 As String)


        Dim StrSql1 As String = "SELECT ШтОтделы.Код, ШтОтделы.Клиент, ШтОтделы.Отделы, ШтСвод.Должность, ШтСвод.Разряд, 
ШтСвод.ТарифнаяСтавка as Ставка, ШтСвод.ПовышениеПроц as Процент, ШтСвод.КодШтСвод, ШтСвод.ДолжИнструкция as [Инструкц]
FROM ШтОтделы INNER JOIN ШтСвод ON ШтОтделы.Код = ШтСвод.Отдел
        WHERE ШтОтделы.Клиент = '" & com1 & "'"
        Dim ds1 As DataTable = Selects(StrSql1)


        Dim DateEx As String = Replace(Format(DateTimePicker1.Value, "yyyy\/MM\/dd"), "/", "")


        Dim strsql As String = "SELECT ШтСводИзмСтавка.* 
FROM ШтОтделы INNER JOIN (ШтСвод INNER JOIN ШтСводИзмСтавка ON ШтСвод.КодШтСвод = ШтСводИзмСтавка.IDКодШтСвод) ON ШтОтделы.Код = ШтСвод.Отдел
WHERE Дата<='" & DateEx & "' AND ШтОтделы.Клиент='" & com1 & "'"
        Dim ds As DataTable = Selects(strsql)

        If ds.Rows.Count = 0 Then
            Очистка()
            'ВыборСтавкиПоДате()
            Refreshgrid()
            Exit Sub
        End If

        Dim lst2 As New List(Of String)()
        lst2 = ВыборкаБезДублей(DateEx, com1)

        For ib As Integer = 0 To ds1.Rows.Count - 1 ' основной модуль по созданию таблицы
            If lst2.Contains(ds1.Rows(ib).Item(7).ToString) Then
                For i As Integer = 0 To ds.Rows.Count - 1
                    If ds.Rows(i).Item(1) = ds1.Rows(ib).Item(7) Then

                        Dim strsql3 As String = "SELECT * FROM ШтСводИзмСтавка
WHERE IDКодШтСвод=" & ds.Rows(i).Item(1) & " AND Дата <= '" & DateEx & "' ORDER BY Дата DESC"
                        Dim ds3 As DataTable = Selects(strsql3)
                        If errds = 1 Then
                            Dim strsql31 As String = "SELECT * FROM ШтСводИзмСтавка
WHERE IDКодШтСвод=" & ds.Rows(i).Item(1) & " AND Дата >= '" & DateEx & "' ORDER BY Дата"
                            Dim ds31 As DataTable = Selects(strsql31)
                            ds1.Rows(ib).Item(5) = ds31.Rows(0).Item(3).ToString
                        Else
                            ds1.Rows(ib).Item(5) = ds3.Rows(0).Item(3).ToString
                        End If

                        Exit For
                    End If
                Next
            End If
        Next

        Очистка()

        For x As Integer = 0 To ds1.Rows.Count - 1
            If ds1.Rows(x).Item(8).ToString = "True" Then
                ds1.Rows(x).Item(8) = "Есть"
            Else
                ds1.Rows(x).Item(8) = "Нет"
            End If
        Next


        Grid1.DataSource = ds1
        Grid1.Columns(1).Visible = False
        Grid1.Columns(7).Visible = False
        Grid1.Columns(0).Visible = False
        'Grid1.Columns(4).Width = 60
        'Grid1.Columns(5).Width = 100
        'Grid1.Columns(1).Width = 150
        Grid1.Columns(2).Width = 200
        Grid1.Columns(3).Width = 200
        Grid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect ' выделяет всю строку в grid1

        'Grid1.Columns(6).Width = 60

        'Grid1.Rows(1).Cells(3).Selected = True
        'Grid1_CellClick(Grid1, New DataGridViewCellEventArgs(<b>3</b>, <b>1</b>))
        'Acti()
        cb = New OleDb.OleDbCommandBuilder(Da)
        s = Grid1.Rows.Count - 1
        изменен = 0
        'NumberAllRows()

    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If TextBox4.Text = "" Then
            MessageBox.Show("Поле отдел не может быть пустым!")
            Exit Sub
        End If
        btnclick = 1
        УдалитьОтдел()
        Очистка()
        ВыборСтавкиПоДате()
        'Refreshgrid()
    End Sub
    Private Sub УдалитьОтдел()
        If MessageBox.Show("Будет удален отдел и все должности!", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.Cancel Then
            Exit Sub
        End If
        Чист()
        StrSql = "DELETE FROM ШтОтделы WHERE Код =" & ГлКод & ""
        Updates(StrSql)
        Статистика1("Отдел " & TextBox4.Text, "Удаление отдела", ComboBox1.Text)
        MessageBox.Show("Данные удалены!", Рик)
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Очистка()
    End Sub
    Private Sub Очистка()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        КодДолжн = Nothing
        ГлКод = Nothing
    End Sub
    Private Sub Удалить()

        If MessageBox.Show("Удалить должность!", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.Cancel Then
            Exit Sub
        End If
        Dim list As New Dictionary(Of String, Object)
        list.Add("@Отдел", ГлКод)
        list.Add("@КодШтСвод", КодДолжн)
        list.Add("@Код", ГлКод)

        Dim ds = Selects(StrSql:="SELECT COUNT (Отдел) FROM ШтСвод WHERE ШтСвод.Отдел =" & ГлКод & "", list)

        If ds.Rows(0).Item(0) > 1 Then

            Updates(stroka:="DELETE FROM ШтСвод WHERE КодШтСвод =@КодШтСвод", list, "ШтСвод")
            MessageBox.Show("Данные удалены!", Рик)
        Else


            Updates(stroka:="DELETE FROM ШтОтделы WHERE Код =@Код", list, "ШтОтделы")


            MessageBox.Show("Данные удалены!", Рик)
        End If
        Статистика1("Должность " & TextBox5.Text, "Удаление должности", ComboBox1.Text)
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If TextBox4.Text = "" Or TextBox5.Text = "" Then
            MessageBox.Show("Поле отдел и должность не могут быть пустыми!")
            Exit Sub
        End If
        btnclick = 1
        Удалить()
        Очистка()
        ВыборСтавкиПоДате()
        'Refreshgrid()
    End Sub

    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Dim sd As String = Strings.UCase(Strings.Left(TextBox5.Text, 1))
            TextBox5.Text = sd & Strings.Right(TextBox5.Text, (TextBox5.TextLength - 1))

            Me.TextBox3.Focus()
        End If
    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox1.Focus()
        End If
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then

            If IsNumeric(TextBox1.Text) = True Then
                Dim f = CDbl(TextBox1.Text)
                f = Math.Round(f, 2)
                Button5.Enabled = True
                Button6.Enabled = True
                Button7.Enabled = True


            Else
                If e.KeyCode = Keys.Decimal Or e.KeyCode = Keys.Oem2 Or e.KeyCode = Keys.OemPeriod Then
                    Replace(TextBox1.Text, ".", ",")
                    Replace(TextBox1.Text, "/", ",")
                    Exit Sub
                End If

                MessageBox.Show("Введите числовое значение!", Рик)
                Button5.Enabled = False
                Button6.Enabled = False
                Button7.Enabled = False
                Exit Sub

            End If

        End If

    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Button5.Focus()
        End If
    End Sub

    Dim ОтдDBC, ДолDBC, РазDBC, ТСтавкаDBC, ПовышПроцDBC As String


    Private Sub Доки()

        Dim strsql As String
        strsql = "SELECT * FROM Клиент WHERE НазвОрг='" & ComboBox1.Text & "'"
        Dim ds As DataTable
        ds = Selects(strsql)


        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document
        'Dim oWordPara As Microsoft.Office.Interop.Word.Paragraph

        'KillProc()

        oWord = CreateObject("Word.Application")
        oWord.Visible = False

        ВыгрузкаФайловНаЛокалыныйКомп(FTPStringAllDOC & "Instrukciya.docx", firthtPath & "\Instrukciya.docx")
        oWordDoc = oWord.Documents.Add(firthtPath & "\Instrukciya.docx") 'из папки на компе выбираем нужный файл

        With oWordDoc.Bookmarks

            If ds.Rows(0).Item(1).ToString = "Индивидуальный предприниматель" Then

                .Item("Инстр1").Range.Text = ФормСобствКор(ds.Rows(0).Item(1).ToString) & " " & ComboBox1.Text
            Else
                .Item("Инстр1").Range.Text = ФормСобствКор(ds.Rows(0).Item(1).ToString) & " «" & ComboBox1.Text & "»"
            End If

            .Item("Инстр2").Range.Text = ДолжИнстр.Дат
            .Item("Инстр3").Range.Text = ДолжИнстр.Ном

            If ds.Rows(0).Item(31) = True Then
                .Item("Инстр4").Range.Text = ds.Rows(0).Item(18).ToString & " " & ФИОКорРук(ds.Rows(0).Item(19).ToString, True)
                .Item("Инстр5").Range.Text = ФИОКорРук(ds.Rows(0).Item(19).ToString, True)
                .Item("Инстр9").Range.Text = ФИОКорРук(ds.Rows(0).Item(19).ToString, True)
            Else
                If ds.Rows(0).Item(1).ToString = "Индивидуальный предприниматель" Then
                    .Item("Инстр4").Range.Text = ds.Rows(0).Item(18).ToString
                Else
                    .Item("Инстр4").Range.Text = ds.Rows(0).Item(18).ToString & " " & ФИОКорРук(ds.Rows(0).Item(19).ToString, False)
                End If

                .Item("Инстр5").Range.Text = ФИОКорРук(ds.Rows(0).Item(19).ToString, False)
                .Item("Инстр9").Range.Text = ФИОКорРук(ds.Rows(0).Item(19).ToString, False)
            End If

            .Item("Инстр6").Range.Text = ДолжИнстр.Дат
            .Item("Инстр7").Range.Text = ДолжИнстр.текст
            .Item("Инстр8").Range.Text = ds.Rows(0).Item(18).ToString
            .Item("Инстр10").Range.Text = Now.Year
            .Item("Инстр11").Range.Text = Now.Year

        End With



        Dim dirstring As String = "Должностные инструкции/" 'место сохранения файла
        dirstring = СозданиепапкиНаСервере(ComboBox1.Text & "/" & dirstring) 'полный путь на сервер(кроме имени и разрешения файла)

        Dim put, Name As String

        If TextBox3.Text = "" Or TextBox3.Text = "-" Then
            Name = ДолжИнстр.Ном & " " & Trim(TextBox4.Text) & " " & Trim(TextBox5.Text) & ".doc"
            put = PathVremyanka & Name 'место в корне программы

            oWordDoc.SaveAs2(put,,,,,, False)
            dirstring += Name

            oWordDoc.Close(True)
            oWord.Quit(True)
            MessageBox.Show("Инструкция добавлена!", Рик)

            ЗагрНаСерверИУдаление(put, dirstring, put) 'загружаем на сервис и чистим времянку

        Else

            Name = ДолжИнстр.Ном & " " & Trim(TextBox4.Text) & " " & Trim(TextBox5.Text) & " " & Trim(TextBox3.Text) & ".doc"
            put = PathVremyanka & Name
            Try
                oWordDoc.SaveAs2(put,,,,,, False)
            Catch ex As Exception
                MessageBox.Show("Не допустимое имя файла!", Рик)
                oWordDoc.Close(True)
                oWord.Quit(True)
                Exit Sub
                'put = Replace(put, "*", "")
                'oWordDoc.SaveAs2(put,,,,,, False)

            End Try

            dirstring += Name

            oWordDoc.Close(True)
            oWord.Quit(True)

            Try
                ЗагрНаСерверИУдаление(put, dirstring, put) 'загружаем на сервис и чистим времянку
                MessageBox.Show("Инструкция добавлена!", Рик)
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If


        Dim gf As String = "True"
        If TextBox3.Text = "" Or TextBox3.Text = "-" Then
            strsql = ""
            strsql = "UPDATE ШтСвод SET ДолжИнструкция='" & gf & "', НомерДолжИнстр='" & ДолжИнстр.Ном & " " & Trim(TextBox4.Text) & " " & Trim(TextBox5.Text) & "', ТекстИнструкции='" & ДолжИнстр.текст & "', ДатаИнструкции='" & ДолжИнстр.Дат & "'  WHERE КодШтСвод=" & КодДолжн & ""
        Else
            strsql = ""
            strsql = "UPDATE ШтСвод SET ДолжИнструкция='" & gf & "', НомерДолжИнстр='" & ДолжИнстр.Ном & " " & Trim(TextBox4.Text) & " " & Trim(TextBox5.Text) & " " & Trim(TextBox3.Text) & "', ТекстИнструкции='" & ДолжИнстр.текст & "', ДатаИнструкции='" & ДолжИнстр.Дат & "' WHERE КодШтСвод=" & КодДолжн & ""
        End If

        Updates(strsql)
        ВыборСтавкиПоДате()
        ВременнаяПапкаУдалениеФайла(firthtPath & "\Instrukciya.docx")
        'Refreshgrid()

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        If MessageBox.Show("Удалить инструкцию?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Exit Sub
        End If

        Dim dv As String = ""

        Dim ds As DataTable = Selects(StrSql:="SELECT НомерДолжИнстр FROM ШтСвод WHERE КодШтСвод =" & КодДолжн & "")

        Try
            УдалениеФайлаНаСервере(FTPString & ComboBox1.Text & "/Должностные инструкции/" & ds.Rows(0).Item(0).ToString & ".doc")

        Catch ex As Exception
            MessageBox.Show("Инструкция не найдена!", Рик)
            Exit Sub
        End Try
        MessageBox.Show("Инструкция удалена!", Рик)


        Dim gf As String = "False"
        обновasync(ComboBox1.Text)
        Updates(stroka:="UPDATE ШтСвод SET ДолжИнструкция='" & gf & "', НомерДолжИнстр='" & ds.Rows(0).Item(0).ToString & "' WHERE КодШтСвод =" & КодДолжн & "")
        Updates(stroka:="UPDATE ШтСвод SET ДолжИнструкция='False', НомерДолжИнстр='', ТекстИнструкции='', ДатаИнструкции='' WHERE КодШтСвод =" & КодДолжн & "")
        ВыборСтавкиПоДате(ComboBox1.Text)

        'Parallel.Invoke(Sub() ВыборСтавкиПоДате(f))
        'Refreshgrid()
    End Sub
    Private Async Sub обновasync(ByVal f As String)
        Await Task.Run(Sub() обнов(f))
    End Sub
    Private Sub обнов(ByVal f As String)
        Dim gf As String = "False"
        Статистика1("Должностная инструкция" & gf, "Удаление инструкции", f)
    End Sub
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click

        If КодДолжн = Nothing Then
            MessageBox.Show("Выберите должность для изменения!")
        End If
        Me.Cursor = Cursors.WaitCursor
        Dim tr As String = "True"
        Dim strsql As String = "SELECT * FROM ШтСвод WHERE КодШтСвод=" & КодДолжн & " AND ДолжИнструкция='" & tr & "'"
        Dim ds As DataTable = Selects(strsql)
        If errds = 1 Then
            MessageBox.Show("Для данной должности еще не сформирована должностная инструкция!", Рик)
            v = False
            ДолжИнстр.ShowDialog()
        Else
            ДолжИнстр.TextBox1.Text = Strings.Left(ds.Rows(0).Item(9).ToString, 3)
            ДолжИнстр.MaskedTextBox1.Text = ds.Rows(0).Item(11).ToString
            ДолжИнстр.RichTextBox1.Text = ds.Rows(0).Item(10).ToString
            ДолжИнстр.x = True
            v = False
            ДолжИнстр.ShowDialog()
        End If


        If v = False Then
            Me.Cursor = Cursors.Default
            Exit Sub

        End If
        If ДолжИнстр.текст = "" Or ДолжИнстр.Ном = "" Then
            Me.Cursor = Cursors.Default
            Exit Sub

        End If

        ДокиИзмен()
        Me.Cursor = Cursors.Default
        Статистика1("Должностная инструкция" & TextBox5.Text, "Изменение инструкции", ComboBox1.Text)




    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        If FT = False Then
            If ComboBox1.Text = "" Then
                MessageBox.Show("Выберите организацию!", Рик)
                Exit Sub
            End If
            ВыборСтавкиПоДате()
        End If
        FT = False

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        timtick = Timer1.Tag
    End Sub

    Private Sub ДокиИзмен()

        Dim ds = dtClientAll.Select("НазвОрг='" & ComboBox1.Text & "'")


        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document
        'Dim oWordPara As Microsoft.Office.Interop.Word.Paragraph

        'KillProc()

        oWord = CreateObject("Word.Application")
        oWord.Visible = False




        'Try
        '    IO.File.Copy(OnePath & "\ОБЩДОКИ\General\Instrukciya.docx", "C:\Users\Public\Documents\Рик\Instrukciya.docx")
        'Catch ex As Exception
        '    'If "Zayavlenie.doc" <> "" Then IO.File.Delete("C:\Users\Public\Documents\Рик\Zayavlenie.doc")
        '    If Not IO.Directory.Exists("c:\Users\Public\Documents\Рик") Then
        '        IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рик")
        '    End If
        '    IO.File.Copy(OnePath & "\ОБЩДОКИ\General\Instrukciya.docx", "C:\Users\Public\Documents\Рик\Instrukciya.docx")
        'End Try

        Начало("Instrukciya.docx")
        oWordDoc = oWord.Documents.Add(firthtPath & "\Instrukciya.docx")

        With oWordDoc.Bookmarks

            If ds(0).Item(1).ToString = "Индивидуальный предприниматель" Then

                .Item("Инстр1").Range.Text = ФормСобствКор(ds(0).Item(1).ToString) & " " & ComboBox1.Text
            Else
                .Item("Инстр1").Range.Text = ФормСобствКор(ds(0).Item(1).ToString) & " «" & ComboBox1.Text & "»"
            End If

            .Item("Инстр2").Range.Text = ДолжИнстр.Дат
            .Item("Инстр3").Range.Text = ДолжИнстр.Ном

            If ds(0).Item(31) = True Then
                .Item("Инстр4").Range.Text = ds(0).Item(18).ToString & " " & ФИОКорРук(ds(0).Item(19).ToString, True)
                .Item("Инстр5").Range.Text = ФИОКорРук(ds(0).Item(19).ToString, True)
                .Item("Инстр9").Range.Text = ФИОКорРук(ds(0).Item(19).ToString, True)
            Else
                If ds(0).Item(1).ToString = "Индивидуальный предприниматель" Then
                    .Item("Инстр4").Range.Text = ds(0).Item(18).ToString
                Else
                    .Item("Инстр4").Range.Text = ds(0).Item(18).ToString & " " & ФИОКорРук(ds(0).Item(19).ToString, False)
                End If

                .Item("Инстр5").Range.Text = ФИОКорРук(ds(0).Item(19).ToString, False)
                .Item("Инстр9").Range.Text = ФИОКорРук(ds(0).Item(19).ToString, False)
            End If

            .Item("Инстр6").Range.Text = ДолжИнстр.Дат
            .Item("Инстр7").Range.Text = ДолжИнстр.текст
            .Item("Инстр8").Range.Text = ds(0).Item(18).ToString
            .Item("Инстр10").Range.Text = Now.Year
            .Item("Инстр11").Range.Text = Now.Year

        End With



        Dim dirstring As String = "Должностные инструкции/" 'место сохранения файла
        dirstring = СозданиепапкиНаСервере(ComboBox1.Text & "/" & dirstring) 'полный путь на сервер(кроме имени и разрешения файла)

        Dim put, Name As String
        If TextBox3.Text = "" Or TextBox3.Text = "-" Then

            Name = ДолжИнстр.Ном & " " & Trim(TextBox4.Text) & " " & Trim(TextBox5.Text) & ".doc"
            put = PathVremyanka & Name 'место в корне программы

            oWordDoc.SaveAs2(put,,,,,, False)
            dirstring += Name

            oWordDoc.Close(True)
            oWord.Quit(True)

            ЗагрНаСерверИУдаление(put, dirstring, put)
            MessageBox.Show("Инструкция изменена!", Рик)

        Else

            Name = ДолжИнстр.Ном & " " & Trim(TextBox4.Text) & " " & Trim(TextBox5.Text) & " " & Trim(TextBox3.Text) & ".doc"
            put = PathVremyanka & Name 'место в корне программы

            oWordDoc.SaveAs2(put,,,,,, False)
            dirstring += Name

            oWordDoc.Close(True)
            oWord.Quit(True)

            ЗагрНаСерверИУдаление(put, dirstring, put)
            MessageBox.Show("Инструкция изменена!", Рик)



        End If


        If TextBox3.Text = "" Then
            StrSql = ""
            StrSql = "UPDATE ШтСвод SET ДолжИнструкция='True', НомерДолжИнстр='" & ДолжИнстр.Ном & " " & Trim(TextBox4.Text) & " " & Trim(TextBox5.Text) & "', ТекстИнструкции='" & ДолжИнстр.текст & "', ДатаИнструкции='" & ДолжИнстр.Дат & "'  WHERE КодШтСвод=" & КодДолжн & ""
        Else
            StrSql = ""
            StrSql = "UPDATE ШтСвод SET ДолжИнструкция='True', НомерДолжИнстр='" & ДолжИнстр.Ном & " " & Trim(TextBox4.Text) & " " & Trim(TextBox5.Text) & " " & Trim(TextBox3.Text) & "', ТекстИнструкции='" & ДолжИнстр.текст & "', ДатаИнструкции='" & ДолжИнстр.Дат & "' WHERE КодШтСвод=" & КодДолжн & ""
        End If

        Updates(StrSql)
        ВыборСтавкиПоДате()
        'Refreshgrid()






    End Sub
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Me.Cursor = Cursors.WaitCursor
        If ComboBox1.Text = "" Then
            MessageBox.Show("Выберите организацию!", Рик)
            Me.Cursor = Cursors.Default
            Exit Sub
        End If

        If TextBox4.Text = "" Or TextBox5.Text = "" Then
            MessageBox.Show("Выберите Отдел и Должность!", Рик)
            Me.Cursor = Cursors.Default
            Exit Sub
        End If

        v = False
        ДолжИнстр.ShowDialog()

        If ДолжИнстр.текст = "" Or ДолжИнстр.Ном = "" Then
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        If v = False Then
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        Доки()
        Me.Cursor = Cursors.Default




        'Dim SFD As New SaveFileDialog With {.Filter = "Файлы Word|*.doc*"}  вызов диалогового окна
        'If SFD.ShowDialog = Windows.Forms.DialogResult.OK Then
        '    MsgBox(SFD.FileName)
        'End If



        'Dim SFD As New OpenFileDialog With {.Filter = "Файлы Word|*.doc*"}
        'If SFD.ShowDialog = Windows.Forms.DialogResult.Cancel Then
        '    Exit Sub
        '    'MsgBox(SFD.FileName)
        'End If

        'If Not IO.Directory.Exists(OnePath & ComboBox1.Text & "\Должностные инструкции\") Then
        '    IO.Directory.CreateDirectory(OnePath & ComboBox1.Text & "\Должностные инструкции\")
        'End If

        'Try
        '    IO.File.Copy(SFD.FileName, OnePath & ComboBox1.Text & "\Должностные инструкции\" & Trim(TextBox4.Text) & " " & Trim(TextBox5.Text) & " " & Trim(TextBox3.Text) & ".doc")

        '    MessageBox.Show("Инструкция добавлена!", Рик)
        'Catch ex As Exception
        '    If MessageBox.Show("Инструкция " & Trim(TextBox4.Text) & " " & Trim(TextBox5.Text) & " " & Trim(TextBox3.Text) & " существует. Заменить старый документ новым?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then
        '        IO.File.Delete(OnePath & ComboBox1.Text & "\Должностные инструкции\" & Trim(TextBox4.Text) & " " & Trim(TextBox5.Text) & " " & Trim(TextBox5.Text) & ".doc")
        '        IO.File.Copy(SFD.FileName, OnePath & ComboBox1.Text & "\Должностные инструкции\" & Trim(TextBox4.Text) & " " & Trim(TextBox5.Text) & " " & Trim(TextBox5.Text) & ".doc")
        '        MessageBox.Show("Инструкция обновлена!", Рик)
        '    End If
        'End Try
        'Dim gf As Boolean = True
        'StrSql = ""
        'StrSql = "UPDATE ШтСвод SET ДолжИнструкция=" & gf & " WHERE КодШтСвод=" & КодДолжн & ""
        'Updates(StrSql)
        'Refreshgrid()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            srt = Grid1.CurrentRow.Cells("Код").Value

        Catch ex As Exception
            MessageBox.Show("Данные не изменены!" & vbCrLf & "Предварительно надо добавить данные, а после их менять!", Рик)
            Exit Sub

        End Try

        'Отд = Grid1.CurrentRow.Cells("Отделы").Value.ToString
        'Дол = Grid1.CurrentRow.Cells("Должность").Value.ToString
        'Раз = Grid1.CurrentRow.Cells("Разряд").Value.ToString
        'ТСтавка = Grid1.CurrentRow.Cells("ТарифнаяСтавка").Value.ToString
        'ПовышПроц = Grid1.CurrentRow.Cells("ПовышениеПроц").Value.ToString

        'WHERE Отдел=" & srt & " AND Должность='" & ДолDBC & "' AND Разряд='" & РазDBC & "' AND ТарифнаяСтавка='" & ТСтавкаDBC & "' AND ПовышениеПроц='" & ПовышПроцDBC & "' AND  КодШтСвод=" & КодDBC & ""
        Dim StrSql As String = "UPDATE ШтСвод SET ТарифнаяСтавка='" & ТСтавкаDBC & "', ПовышениеПроц='" & ПовышПроцDBC & "', Должность= '" & ДолDBC & "', Разряд='" & РазDBC & "'
        WHERE КодШтСвод=" & КодDBC & ""

        Dim c As New OleDbCommand
        c.Connection = conn
        c.CommandText = StrSql
        Try
            c.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show("Ошибка в изменении строки штатного расписания", Рик)
            Exit Sub
        End Try

        Dim StrSql2 As String = "UPDATE ШтОтделы INNER JOIN ШтСвод ON ШтОтделы.Код = ШтСвод.Отдел SET Отделы= '" & ОтдDBC & "'
            WHERE Код=" & srt & " AND КодШтСвод=" & КодDBC & ""

        ' 
        Dim c2 As New OleDbCommand
        c2.Connection = conn
        c2.CommandText = StrSql2
        Try
            c2.ExecuteNonQuery()

        Catch ex As Exception
            MessageBox.Show("Ошибка, строка 60!", Рик)
            Exit Sub
        End Try


        Dim StrSql3 As String = "UPDATE Штатное INNER JOIN Сотрудники ON Сотрудники.КодСотрудники = Штатное.ИДСотр SET ТарифнаяСтавка='" & ТСтавкаDBC & "', ПовышОклПроц='" & ПовышПроцDBC & "'
Where Сотрудники.НазвОрганиз = '" & ComboBox1.Text & "' And Штатное.Должность = '" & ДолDBC & "' AND Штатное.Отдел = '" & ОтдDBC & "'"

        Dim c3 As New OleDbCommand
        c3.Connection = conn
        c3.CommandText = StrSql3
        Try
            c3.ExecuteNonQuery()
            MessageBox.Show("Данные изменены!", Рик)
            ВыборСтавкиПоДате()
            'Refreshgrid()
        Catch ex As Exception
            Exit Sub
        End Try

    End Sub

    Dim mas2, mas3

    Private Sub Refreshgrid()
        Организ = ComboBox1.Text

        Dim StrSql1 As String
        tbl.Clear()

        StrSql1 = "SELECT ШтОтделы.Код, ШтОтделы.Клиент, ШтОтделы.Отделы, ШтСвод.Должность, ШтСвод.Разряд, 
ШтСвод.ТарифнаяСтавка as Ставка, ШтСвод.ПовышениеПроц as Процент, ШтСвод.КодШтСвод, ШтСвод.ДолжИнструкция as [Инструкц]
FROM ШтОтделы INNER JOIN ШтСвод ON ШтОтделы.Код = ШтСвод.Отдел
        WHERE ШтОтделы.Клиент = '" & Организ & "'"

        tbl = Selects(StrSql1)

        For x As Integer = 0 To tbl.Rows.Count - 1
            If tbl.Rows(x).Item(8).ToString = "True" Then
                tbl.Rows(x).Item(8) = "Есть"
            Else
                tbl.Rows(x).Item(8) = "Нет"
            End If
        Next

        Grid1.DataSource = tbl
        Grid1.Columns(1).Visible = False
        Grid1.Columns(7).Visible = False
        Grid1.Columns(0).Visible = False
        'Grid1.Columns(4).Width = 60
        'Grid1.Columns(5).Width = 100
        'Grid1.Columns(1).Width = 150
        Grid1.Columns(2).Width = 200
        Grid1.Columns(3).Width = 200
        Grid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect ' выделяет всю строку в grid1

        'Grid1.Columns(6).Width = 60

        'Grid1.Rows(1).Cells(3).Selected = True
        'Grid1_CellClick(Grid1, New DataGridViewCellEventArgs(<b>3</b>, <b>1</b>))
        'Acti()

        s = Grid1.Rows.Count - 1
        изменен = 0
        'NumberAllRows()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click 'удаление

        Dim bnbc As Integer = MsgBox("Удалить строку?", vbOKCancel, Рик)

        Select Case bnbc
            Case 2
                ВыборСтавкиПоДате()
                'Refreshgrid()
                Exit Sub
            Case 1

        End Select



        Dim k As Integer
        Dim n, m, y, t As String
        k = Grid1.CurrentRow.Cells("Код").Value
        n = Grid1.CurrentRow.Cells("Должность").Value
        m = Grid1.CurrentRow.Cells("Разряд").Value
        y = Grid1.CurrentRow.Cells("ТарифнаяСтавка").Value
        t = Grid1.CurrentRow.Cells("ПовышениеПроц").Value



        Dim StrSql1 As String = "DELETE FROM ШтСвод WHERE Отдел =" & k & " And Должность='" & n & "' and Разряд='" & m & "' and ТарифнаяСтавка='" & y & "' and ПовышениеПроц='" & t & "'"
        Dim c As New OleDbCommand
        c.Connection = conn
        c.CommandText = StrSql1
        c.ExecuteNonQuery()
        ВыборСтавкиПоДате()
        'Refreshgrid()
    End Sub
    Private Sub ЗагрПроцОклРазр()
        Чист()
        StrSql = ""
        StrSql = "SELECT ШтСвод.Разряд, ШтСвод.ТарифнаяСтавка, ШтСвод.ПовышениеПроц
FROM ШтОтделы INNER JOIN ШтСвод ON ШтОтделы.Код = ШтСвод.Отдел
WHERE ШтОтделы.Клиент='" & ComboBox1.Text & "' AND ШтОтделы.Отделы='" & TextBox4.Text & "' AND ШтСвод.Должность='" & TextBox5.Text & "'"
        ds = Selects(StrSql)

        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""

        TextBox1.Text = ds.Rows(0).Item(1).ToString
        TextBox2.Text = ds.Rows(0).Item(2).ToString
        TextBox3.Text = ds.Rows(0).Item(0).ToString

    End Sub


    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        DateTimePicker1.Enabled = True
        s = Grid1.Rows.Count - 1
        s2 = Grid1.Rows.Count - 2
        'ЗагрОтделов()

        ВыборСтавкиПоДате()

    End Sub

    Private Sub Чист()
        StrSql = ""
        Try
            ds.Clear()
        Catch ex As Exception

        End Try

    End Sub



    Function ПровЗапПолей() As Integer

        For ip As Integer = 0 To Grid1.Rows.Count - 2 'проверяем заполненность поля должность

            Dim sOtd1 As String = Grid1.Rows(ip).Cells(3).Value.ToString
            Dim fd As Integer
            If sOtd1 = "" Then
                fd = MsgBox("Заполните колонку Должность - строки " & ip + 1, vbOKCancel, Рик)
                Select Case fd
                    Case 1
                        Return 1
                    Case 2
                        'Refreshgrid()
                        Return 1
                End Select
            End If

            sOtd1 = ""
            sOtd1 = Grid1.Rows(ip).Cells(5).Value.ToString
            If sOtd1 = "" Then
                fd = MsgBox("Заполните столбец Тарифная ставка - строки " & ip + 1, vbOKCancel, Рик)
                Select Case fd
                    Case 1
                        Return 1
                    Case 2
                        'Refreshgrid()
                        Return 1
                End Select
            End If

            sOtd1 = ""
            sOtd1 = Grid1.Rows(ip).Cells(2).Value.ToString
            If sOtd1 = "" Then
                fd = MsgBox("Заполните столбец Отдел - строки " & ip + 1, vbOKCancel, Рик)
                Select Case fd
                    Case 1
                        Return 1
                    Case 2
                        'Refreshgrid()
                        Return 1

                End Select
            End If

        Next
        Return 2
    End Function

    Private Sub ВстИДНовОтд()

        Dim ses As Integer = se - s
        Dim StrSql2, StrSql4, StrSql5 As String
        Dim c1, c2, c3 As New OleDbCommand
        Dim ds1, ds8 As New DataSet
        Dim da1 As New OleDbDataAdapter(c1)
        Dim da2 As New OleDbDataAdapter(c2)
        Dim da3 As New OleDbDataAdapter(c3)
        Dim coli As Integer
        Dim i As Integer 'проверяем есть ли в базе уже такая должность и если есть присваиваем код соответсвующий должности
        Dim sOtd As String
        For i = 0 To ses - 1

            sOtd = ""
            StrSql2 = ""
            sOtd = Grid1.Rows(s + i).Cells(2).Value.ToString
            Try ' заполняем базу и возвращаем номер ИД дл
                StrSql2 = "Select Код FROM ШтОтделы WHERE Клиент = '" & Организ & "' AND Отделы='" & sOtd & "' "
                With c1
                    .Connection = conn
                    .CommandText = StrSql2
                End With

                da1.Fill(ds1, "f")
                Grid1.Rows(s + i).Cells(0).Value = ds1.Tables("f").Rows(0).Item(0).ToString
                StrSql5 = ""
                StrSql5 = "INSERT INTO ШтСвод(Отдел,Должность,Разряд,ТарифнаяСтавка,ПовышениеПроц) VALUES (" & Grid1.Rows(s + i).Cells(0).Value & ",'" & StrConv(Grid1.Rows(s + i).Cells(3).Value, VbStrConv.ProperCase) & "','" & Grid1.Rows(s + i).Cells(4).Value & "','" & Grid1.Rows(s + i).Cells(5).Value & "','" & Grid1.Rows(s + i).Cells(6).Value & "')" ' вставляем в базу должность, тар.ставку, повыш, и разряд
                c3.Connection = conn
                c3.CommandText = StrSql5
                c3.ExecuteNonQuery()


                Dim коднов As Integer = ds1.Tables("f").Rows(0).Item(0)
                coli = 1
                coli += i
                ds1.Clear()
            Catch ex As Exception ' вставка в базу нового отдела и выборка оттуда номера отдела
                StrSql4 = "INSERT INTO ШтОтделы(Отделы,Клиент) VALUES ('" & StrConv(Grid1.Rows(s + i).Cells(2).Value, VbStrConv.ProperCase) & "','" & Организ & "')" ' вставляем в базу должность, тар.ставку, повыш, и разряд
                c2.Connection = conn
                c2.CommandText = StrSql4
                c2.ExecuteNonQuery()

                Dim c7 As New OleDbCommand
                Dim ds7 As New DataSet
                Dim da7 As New OleDbDataAdapter(c7)
                Dim StrSql7 As String = "Select Код FROM ШтОтделы WHERE Клиент = '" & Организ & "'   AND Отделы='" & Grid1.Rows(s + i).Cells(2).Value & "' "
                With c7
                    .Connection = conn
                    .CommandText = StrSql7
                End With

                da7.Fill(ds7, "f") 'вставка номера отдела в таблицу
                Grid1.Rows(s + i).Cells(0).Value = ds7.Tables("f").Rows(0).Item(0).ToString
                StrSql5 = ""
                StrSql5 = "INSERT INTO ШтСвод(Отдел,Должность,Разряд,ТарифнаяСтавка,ПовышениеПроц) VALUES (" & Grid1.Rows(s + i).Cells(0).Value & ",'" & Grid1.Rows(s + i).Cells(3).Value & "','" & Grid1.Rows(s + i).Cells(4).Value & "','" & Grid1.Rows(s + i).Cells(5).Value & "','" & Grid1.Rows(s + i).Cells(6).Value.ToString & "')" ' вставляем в базу должность, тар.ставку, повыш, и разряд
                c3.Connection = conn
                c3.CommandText = StrSql5
                c3.ExecuteNonQuery()




            End Try
        Next
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        se = Grid1.Rows.Count - 1
        If изменен = 0 And s = se Then
            MsgBox("Нет изменений",, Рик)
            Exit Sub
        End If

        Dim bnbc As Integer = MsgBox("Сохранить данные?", vbOKCancel, Рик)
        Select Case bnbc
            Case 2

                Exit Sub
            Case 1

        End Select

        Dim fg As Integer = ПровЗапПолей()
        Select Case fg
            Case 1
                Exit Sub
        End Select
        If s < se Then ВстИДНовОтд()

        Dim MosiFF(Grid1.Columns.Count - 1, Grid1.Rows.Count - 1)
        Dim Str As String = ""

        For Row As Integer = 0 To Grid1.Rows.Count - 1
            For Col As Integer = 0 To Grid1.Columns.Count - 1

                MosiFF(Col, Row) = Grid1.Item(Col, Row).Value
                'Str &= MosiFF(Col, Row) & " "
            Next
            'Str &= vbCrLf
        Next


        Dim i As Integer = 0
        Dim StrSql As String 'сохранение в базу
        Dim vbn As Integer
        If s = se Then
            vbn = 1
        Else
            vbn = 2
        End If
        For i = 0 To Grid1.Rows.Count - 2 'LBound(MosiFF) To UBound(MosiFF)
            StrSql = "UPDATE ШтСвод  SET ТарифнаяСтавка= '" & MosiFF(5, i) & "',ПовышениеПроц='" & MosiFF(6, i) & "'
            WHERE ШтСвод.Отдел=" & MosiFF(0, i) & " AND ШтСвод.Должность='" & MosiFF(3, i) & "' AND ШтСвод.Разряд='" & MosiFF(4, i).ToString & "'"
            Dim c As New OleDbCommand
            c.Connection = conn
            c.CommandText = StrSql
            c.ExecuteNonQuery()

        Next

        MsgBox("Данные сохранены!",, Рик)
        'Refreshgrid()
        ВыборСтавкиПоДате()


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        'Da.UpdateCommand = cb.GetUpdateCommand() 'обновление одной таблицы
        'Da.Update(tbl)
        se = Grid1.Rows.Count - 1
        If изменен = 0 And s = se Then
            MsgBox("Нет изменений")

        End If



        'thb = Grid1.Rows(mas3).Cells(mas2).value.ToString 'проверяем изменения значения до и после редакции
        If s2 >= изменен Then
            Dim fg As Integer = ПровЗапПолей()
            Select Case fg
                Case 1
                    Exit Sub
            End Select

            If s < se Then ВстИДНовОтд()


            Dim MosiFF(Grid1.Columns.Count - 1, Grid1.Rows.Count - 1)


            For Row As Integer = 0 To Grid1.Rows.Count - 1
                For Col As Integer = 0 To Grid1.Columns.Count - 1
                    MosiFF(Col, Row) = Grid1.Item(Col, Row).Value
                    'Str &= MosiFF(Col, Row) & " "
                Next
                'Str &= vbCrLf
            Next



            'Dim bool As Boolean
            Dim i As Integer = 0
            Dim StrSql As String 'сохранение в базу

            For i = 0 To Grid1.Rows.Count - 2 'LBound(MosiFF) To UBound(MosiFF)
                StrSql = "UPDATE ШтСвод  SET ТарифнаяСтавка= '" & MosiFF(5, i) & "',ПовышениеПроц='" & MosiFF(6, i) & "'
            WHERE ШтСвод.Отдел=" & MosiFF(0, i) & " AND ШтСвод.Должность='" & MosiFF(3, i) & "' AND ШтСвод.Разряд='" & MosiFF(4, i).ToString & "'"
                Updates(StrSql)

            Next

        Else
            Dim fg2 As Integer = ПровЗапПолей()
            Select Case fg2
                Case 1
                    Exit Sub
            End Select

            Dim ses As Integer = se - s
            Dim StrSql2, StrSql4 As String
            'Dim c1, c2 As New OleDbCommand
            'Dim ds1 As New DataSet
            'Dim da1 As New OleDbDataAdapter(c1)

            Select Case ses
                Case > 0

                    Dim coli As Integer
                    Dim i As Integer 'проверяем есть ли в базе уже такая должность и если есть присваиваем код соответсвующий должности
                    For i = 0 To ses - 1
                        Dim sOtd As String = Grid1.Rows(s + i).Cells(2).Value.ToString
                        Try ' заполняем базу и возвращаем номер ИД дл
                            StrSql2 = "Select Код FROM ШтОтделы WHERE Клиент = '" & Организ & "'   AND Отделы='" & sOtd & "' "
                            Dim ds1 As DataTable = Selects(StrSql2)

                            Grid1.Rows(s + i).Cells(0).Value = ds1.Rows(0).Item(0).ToString

                            Dim коднов As Integer = ds1.Rows(0).Item(0)
                            coli = 1
                            coli += i


                            StrSql4 = "INSERT INTO ШтСвод(Отдел,Должность,Разряд,ТарифнаяСтавка,ПовышениеПроц)
VALUES (" & коднов & ",'" & Grid1.Rows(s + i).Cells(3).Value & "','" & Grid1.Rows(s + i).Cells(4).Value & "','" & Grid1.Rows(s + i).Cells(5).Value & "','" & Grid1.Rows(s + i).Cells(6).Value & "')" ' вставляем в базу должность, тар.ставку, повыш, и разряд
                            Updates(StrSql4)

                        Catch ex As Exception

                        End Try
                    Next
                    If coli = ses Then
                        Exit Sub
                    End If
            End Select

            For i = 0 To ses - 1

                Dim StrSql1 As String = "INSERT INTO ШтОтделы(Клиент,Отделы) VALUES ('" & Организ & "','" & Grid1.Rows(s + i).Cells(2).Value & "')"
                Updates(StrSql1)

                Dim StrSql5 As String = "SELECT Код FROM ШтОтделы WHERE Клиент = '" & Организ & "' AND Отделы='" & Grid1.Rows(s + i).Cells(2).Value & "' "
                Dim ds5 As DataTable = Selects(StrSql5)

                Grid1.Rows(s + i).Cells(0).Value = ds5.Rows(0).Item(0).ToString

                Dim длж As String = Grid1.Rows(s + i).Cells(3).Value.ToString
                Dim разр As String = Grid1.Rows(s + i).Cells(4).Value.ToString
                Dim тстав As String = Grid1.Rows(s + i).Cells(5).Value.ToString
                Dim проц As String = Grid1.Rows(s + i).Cells(6).Value.ToString

                Dim StrSql6 As String = "INSERT INTO ШтСвод(Отдел,Должность,Разряд,ТарифнаяСтавка,ПовышениеПроц) VALUES (" & ds5.Rows(0).Item(0) & ",'" & длж & "','" & разр & "','" & тстав & "','" & проц & "')" ' вставляем в базу должность, тар.ставку, повыш, и разряд
                Updates(StrSql6)


            Next
        End If



Конец:
        ВыборСтавкиПоДате()

    End Sub

    Private Sub Grid1_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles Grid1.CellBeginEdit
        Select Case e.ColumnIndex
            Case 0
                e.Cancel = True
        End Select


    End Sub



    Private Sub Grid1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellEndEdit

        ОтдDBC = Grid1.CurrentRow.Cells("Отделы").Value.ToString
        ДолDBC = Grid1.CurrentRow.Cells("Должность").Value.ToString
        РазDBC = Grid1.CurrentRow.Cells("Разряд").Value.ToString
        ТСтавкаDBC = Grid1.CurrentRow.Cells("ТарифнаяСтавка").Value.ToString
        ПовышПроцDBC = Grid1.CurrentRow.Cells("ПовышениеПроц").Value.ToString
        Try
            КодDBC = Grid1.CurrentRow.Cells("КодШтСвод").Value
        Catch ex As Exception

        End Try











    End Sub

    Private Sub Grid1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellValueChanged
        If (e.ColumnIndex = -1) Then Return



        изменен = Grid1.CurrentCellAddress.Y
        изменен += 1

    End Sub

    Private Sub Grid1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellDoubleClick



        Dim IDсвод As Integer = Grid1.CurrentRow.Cells("КодШтСвод").Value.ToString()
        mast2.Clear()
        mast2.AddRange({IDсвод, ComboBox1.Text})
        ШтИзмСтавкиВспл.ShowDialog()

    End Sub


    Private Sub Acti()
        Try
            TextBox4.Text = Grid1.CurrentRow.Cells("Отделы").Value.ToString()
            TextBox5.Text = Grid1.CurrentRow.Cells("Должность").Value.ToString
            TextBox3.Text = Grid1.CurrentRow.Cells("Разряд").Value.ToString
            TextBox1.Text = Grid1.CurrentRow.Cells("Ставка").Value.ToString
            TextBox2.Text = Grid1.CurrentRow.Cells("Процент").Value.ToString
        Catch ex As Exception
            MessageBox.Show("Кликните по полю таблицы!", Рик)
            Exit Sub
        End Try




        ГлКод = Nothing
        'ГлКод = ds.Rows(0).Item(1)
        ГлКод = Grid1.CurrentRow.Cells("Код").Value
        КодДолжн = Grid1.CurrentRow.Cells("КодШтСвод").Value

        If Grid1.CurrentRow.Cells("Инструкц").Value.ToString = "Есть" Then
            Button10.Enabled = False
            Button12.Enabled = True
            Button11.Enabled = True
        Else
            Button10.Enabled = True
            Button12.Enabled = False
            Button11.Enabled = False
        End If


    End Sub
    Private Sub Grid1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellClick

        Acti()

    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Dim sd As String = Strings.UCase(Strings.Left(TextBox5.Text, 1))
            Try
                TextBox5.Text = sd & Strings.Right(TextBox5.Text, (TextBox5.TextLength - 1))
            Catch ex As Exception

            End Try


            Me.TextBox5.Focus()
        End If
    End Sub



    Private Sub TextBox5_LostFocus(sender As Object, e As EventArgs) Handles TextBox5.LostFocus
        Dim sd As String = Strings.UCase(Strings.Left(TextBox5.Text, 1))
        Try
            TextBox5.Text = sd & Strings.Right(TextBox5.Text, (TextBox5.TextLength - 1))

        Catch ex As Exception

        End Try

    End Sub
    Private Sub TextBox4_LostFocus(sender As Object, e As EventArgs) Handles TextBox4.LostFocus
        Dim sd As String = Strings.UCase(Strings.Left(TextBox5.Text, 1))
        Try
            TextBox5.Text = sd & Strings.Right(TextBox5.Text, (TextBox5.TextLength - 1))
        Catch ex As Exception

        End Try


    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If btnclick = 1 Then Exit Sub 'если нажата кнопка не проверяем это поле
        If fnm9 = 1 Then
            fnm9 = 0
            Exit Sub
        End If
        If IsNumeric(TextBox1.Text) = False Then
            If TextBox1.Text.Contains(".") Then
                Replace(TextBox1.Text, ".", ",")
                Exit Sub
            End If
            If TextBox1.Text = "" Then
                Exit Sub
            End If
            fnm9 = 1
            MessageBox.Show("Введите числовое значение!", Рик)
            TextBox1.Text = ""
            Button5.Enabled = False
            Button6.Enabled = False
            Button7.Enabled = False
        Else
            Button5.Enabled = True
            Button6.Enabled = True
            Button7.Enabled = True

        End If
    End Sub

    'Private Sub Штатное_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
    '    Timer1.Tag = sender.name
    '    If e.KeyCode = Keys.Decimal Then
    '        MessageBox.Show("or")
    '    End If

    '    timtick = e.KeyValue.ToString
    'End Sub




End Class