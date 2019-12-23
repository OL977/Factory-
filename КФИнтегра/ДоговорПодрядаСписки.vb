Option Explicit On
Imports System.Data.OleDb
Imports System.IO
Imports System.Threading
Imports System.Data.Linq
Public Class ДоговорПодрядаСписки
    Dim IdСот As Integer
    Public clb2 As Boolean
    Private Delegate Sub comb19()
    Private Delegate Sub comb1()
    Dim strikethrough_style As New DataGridViewCellStyle
    Dim dsAll, dsAll1 As DataTable
    Public dsAct As DataTable
    Dim com1 As Boolean = False, com19 As Boolean = False
    Dim cobbx1 As String
    Dim dsId As DataTable
    Public Property dsart() As DataTable

    Private Sub Обход1()
        If ComboBox1.InvokeRequired Then
            Me.Invoke(New comb1(AddressOf Обход1))
        Else
            ComboBox1.AutoCompleteCustomSource.Clear()
            ComboBox1.Items.Clear()
            For Each r As DataRow In СписокКлиентовОсновной.Rows
                Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
                Me.ComboBox1.Items.Add(r(0).ToString)
            Next
        End If
    End Sub
    Private Sub Обход2()

        If ComboBox19.InvokeRequired Or ComboBox2.InvokeRequired Then
            Me.Invoke(New comb19(AddressOf Обход2))
        Else
            Dim StrSql1 As String = "SELECT DISTINCT Сотрудники.ФИОСборное, ДогПодряда.ID
FROM Сотрудники INNER JOIN ДогПодряда ON Сотрудники.КодСотрудники = ДогПодряда.ID ORDER BY ФИОСборное"
            Dim ds1 As DataTable = Selects(StrSql1)

            Me.ComboBox19.AutoCompleteCustomSource.Clear()
            Me.ComboBox19.Items.Clear()
            Me.ComboBox2.Items.Clear()

            For Each r As DataRow In ds1.Rows
                Me.ComboBox19.AutoCompleteCustomSource.Add(r.Item(0).ToString())
                Me.ComboBox19.Items.Add(r(0).ToString)
                Me.ComboBox2.Items.Add(r(1).ToString)
            Next
        End If
    End Sub
    Private Sub ДоговорПодрядаСписки_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.MdiParent = MDIParent1

        Dim fd As New Thread(AddressOf Обход1)
        fd.IsBackground = True
        fd.Start()

        Dim fd1 As New Thread(AddressOf Обход2)
        fd1.IsBackground = True
        fd1.Start()

        Dim f As New Thread(AddressOf AllGrid)
        f.IsBackground = True
        f.Start()
    End Sub
    Private Sub AllGrid()

        Dim StrSql2 As String = "SELECT Сотрудники.КодСотрудники, Сотрудники.НазвОрганиз as [Организация], Сотрудники.ФИОСборное as [ФИО],
ДогПодряда.НомерДогПодр as [Номер Договора Подряда],ДогПодряда.ДатаДогПодр as [Дата договора подряда], ДогПодряда.Должность,
ДогПодряда.ДатаНачала as [Дата начала ДогПодряда], ДогПодряда.ДатаОконч as [Дата оконч ДогПодряда], ДогПодряда.СтоимЧасаРуб as [руб], 
ДогПодряда.СтоимЧасаКоп as [коп], ДогПодряда.ОбъекОбщепита as [Объект], ДогПодрядаАкт.ПорНомерАкта as [№ акта],
ДогПодрядаАкт.ДатаАкта as [Дата акта], ДогПодрядаАкт.ЗаПериодС as [За период  с], ДогПодрядаАкт.ЗаПериодПо as [За период по],
ДогПодрядаАкт.СтоимРабот as [Стоимость работ], ДогПодряда.Примечание
FROM (Сотрудники INNER JOIN ДогПодряда ON Сотрудники.КодСотрудники = ДогПодряда.ID) INNER JOIN ДогПодрядаАкт ON ДогПодряда.Код = ДогПодрядаАкт.IDДогПодр
ORDER BY Сотрудники.НазвОрганиз"
        dsAll = Selects(StrSql2)

    End Sub
    Private Function Orgall() As DataTable
        Dim list As New Dictionary(Of String, Object)
        list.Add("@НазвОрганиз", ComboBox1.Text)

        Dim ds = Selects(StrSql:="SELECT Сотрудники.КодСотрудники, Сотрудники.НазвОрганиз as [Организация], Сотрудники.ФИОСборное as [ФИО],
ДогПодряда.НомерДогПодр as [Номер Договора Подряда],ДогПодряда.ДатаДогПодр as [Дата договора подряда], ДогПодряда.Должность,
ДогПодряда.ДатаНачала as [Дата начала ДогПодряда], ДогПодряда.ДатаОконч as [Дата оконч ДогПодряда], ДогПодряда.СтоимЧасаРуб as [руб], 
ДогПодряда.СтоимЧасаКоп as [коп], ДогПодряда.ОбъекОбщепита as [Объект], ДогПодрядаАкт.ПорНомерАкта as [№ акта],
ДогПодрядаАкт.ДатаАкта as [Дата акта], ДогПодрядаАкт.ЗаПериодС as [За период  с], ДогПодрядаАкт.ЗаПериодПо as [За период по],
ДогПодрядаАкт.СтоимРабот as [Стоимость работ], ДогПодряда.Примечание
FROM (Сотрудники INNER JOIN ДогПодряда ON Сотрудники.КодСотрудники = ДогПодряда.ID) INNER JOIN ДогПодрядаАкт ON ДогПодряда.Код = ДогПодрядаАкт.IDДогПодр
WHERE Сотрудники.НазвОрганиз=@НазвОрганиз ORDER BY Сотрудники.НазвОрганиз", list)

        Return ds
    End Function
    Private Sub Grid4(ByVal dt As DataTable)
        Dim list As New Dictionary(Of String, Object)
        list.Add("@НазвОрганиз", ComboBox1.Text)

        Dim ds = Selects(StrSql:="SELECT DISTINCT НомерДогПодр 
FROM Сотрудники INNER JOIN ДогПодряда ON Сотрудники.КодСотрудники = ДогПодряда.ID
WHERE Сотрудники.НазвОрганиз=@НазвОрганиз", list)


        Dim nList As New List(Of Integer)
        Dim mas As New List(Of Integer)
        Dim mas2 As New List(Of Integer)

        For i As Integer = 0 To ds.Rows.Count - 1

            For Each row As DataRow In dt.Rows
                If row.Item("Номер ДогПодр").ToString = ds.Rows(i).Item(0).ToString Then
                    nList.Add(row.Item(17))
                End If
            Next


            If nList.Count = 1 Then
                mas2.Add(nList(0))
            End If
            If nList.Count > 1 Then
                'nList.Reverse()
                mas2.Add(nList(0))
                nList.RemoveAt(0)
                For x As Integer = 0 To nList.Count - 1
                    mas.Add(nList.Item(x))
                Next
            End If
            nList.Clear()
        Next

        For i As Integer = 0 To mas.Count - 1
            For Each row As DataRow In dt.Rows
                If row.Item("Код") = mas.Item(i) Then
                    With row
                        .Item("ФИО") = Nothing
                        .Item("Дата") = Nothing
                        .Item("Номер ДогПодр") = Nothing
                        .Item("Дата начала") = Nothing
                        .Item("Дата оконч") = Nothing
                        .Item("Должность") = Nothing
                    End With

                End If
            Next
        Next


        Grid1.DataSource = dt
        For i As Integer = 0 To mas2.Count - 1
            For Each fow As DataGridViewRow In Grid1.Rows
                If fow.Cells("Код").Value = mas2.Item(i) Then
                    fow.DefaultCellStyle.BackColor = Color.LightGreen
                End If
            Next
        Next


        For i As Integer = 0 To mas.Count - 1
            For Each fow As DataGridViewRow In Grid1.Rows
                If fow.Cells("Код").Value = mas.Item(i) Then
                    fow.DefaultCellStyle.BackColor = Color.MistyRose
                End If
            Next
        Next





        Grid1.Columns(1).Visible = False
        'Grid1.Columns(17).Visible = False
        Grid1.Columns(2).Visible = True
        Grid1.Columns(0).Visible = False

        Grid1.Columns(17).Visible = True

        Grid1.Columns(10).Visible = False
        Grid1.Columns(8).Visible = False
        Grid1.Columns(9).Visible = False
        'Grid1.Columns(1).Width = 200
        'Grid1.Columns(2).Width = 200
        Grid1.Columns(5).Width = 150
        Grid1.Columns(11).Width = 50
        Grid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Grid1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

    End Sub



    Private Sub Отсортировка(ByVal grid1 As DataGridView)


        'Dim strsql As String = "SELECT DISTINCT НомерДогПодр FROM ДогПодряда WHERE ID=" & CType(Label96.Text, Integer) & ""
        'Dim ds As DataTable = Selects(strsql)
        Dim ds = From x In dtDogovorPadriadaAll Where x.Item("ID") = (CType(Label96.Text, Integer)) Select x.Item("НомерДогПодр") Distinct


        Dim mas2 As New List(Of Integer)


        Dim nList As New List(Of Integer)
        Dim mas As New List(Of Integer)

        For i As Integer = 0 To ds.Count - 1

            For Each row As DataGridViewRow In Me.Grid1.Rows
                If row.Cells("Номер ДогПодр").Value.ToString = ds(i) Then
                    nList.Add(row.Cells.Item(17).Value)
                End If
            Next



            If nList.Count = 1 Then
                mas2.Add(nList(0))
            End If


            If nList.Count > 1 Then
                'nList.Reverse()
                mas2.Add(nList(0))
                nList.RemoveAt(0)
                For x As Integer = 0 To nList.Count - 1
                    mas.Add(nList.Item(x))
                Next
            End If
            nList.Clear()
        Next

        For i As Integer = 0 To mas.Count - 1
            For Each row As DataGridViewRow In Me.Grid1.Rows
                If row.Cells("Код").Value = mas.Item(i) Then
                    With row
                        .Cells("Номер ДогПодр").Value = Nothing
                        .Cells("Дата").Value = Nothing
                        .Cells("Дата начала").Value = Nothing
                        .Cells("Дата оконч").Value = Nothing
                        .Cells("Должность").Value = Nothing
                    End With
                End If
            Next
        Next

        For i As Integer = 0 To mas2.Count - 1
            For Each fow As DataGridViewRow In grid1.Rows
                If fow.Cells("Код").Value = mas2.Item(i) Then
                    fow.DefaultCellStyle.BackColor = Color.LightGreen
                End If
            Next
        Next


        For i As Integer = 0 To mas.Count - 1
            For Each fow As DataGridViewRow In grid1.Rows
                If fow.Cells("Код").Value = mas.Item(i) Then
                    fow.DefaultCellStyle.BackColor = Color.MistyRose
                End If
            Next
        Next



    End Sub
    Private Sub grid3(ByVal ds2 As DataTable)


        Grid1.DataSource = ds2
        Отсортировка(Grid1)
        Grid1.Columns(2).Visible = False
        Grid1.Columns(1).Visible = False
        'Grid1.Columns(17).Visible = False

        Grid1.Columns(0).Visible = False
        'Grid1.Columns(1).Visible = False
        Grid1.Columns(10).Visible = False
        Grid1.Columns(8).Visible = False
        Grid1.Columns(9).Visible = False
        'Grid1.Columns(1).Width = 200
        'Grid1.Columns(2).Width = 200
        Grid1.Columns(5).Width = 150
        Grid1.Columns(11).Width = 50
        Grid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Grid1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Grid1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter





    End Sub
    Private Sub grid2(ByVal ds2 As DataTable)


        Grid1.DataSource = ds2
        If com1 = True Then
            Grid1.Columns(1).Visible = False

        Else
            Grid1.Columns(1).Visible = True
        End If





        Try
            'Grid1.DefaultCellStyle = strikethrough_style
            'Grid1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Grid1.Columns(0).Visible = False
            'Grid1.Columns(1).Visible = False
            Grid1.Columns(10).Visible = False
            Grid1.Columns(8).Visible = False
            Grid1.Columns(9).Visible = False
            'Grid1.Columns(1).Width = 200
            'Grid1.Columns(2).Width = 200
            Grid1.Columns(5).Width = 150
            Grid1.Columns(11).Width = 50
            Grid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            Grid1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            Grid1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        ComboBox19.Text = ""

        Dim list As New Dictionary(Of String, Object)
        list.Add("@НазвОрганиз", ComboBox1.SelectedItem)


        Dim ds2 = Selects(StrSql:="SELECT Сотрудники.КодСотрудники, Сотрудники.НазвОрганиз as [Организация], Сотрудники.ФИОСборное as [ФИО],
ДогПодряда.НомерДогПодр as [Номер ДогПодр],ДогПодряда.ДатаДогПодр as [Дата], ДогПодряда.Должность,
ДогПодряда.ДатаНачала as [Дата начала], ДогПодряда.ДатаОконч as [Дата оконч], ДогПодряда.СтоимЧасаРуб as [руб], 
ДогПодряда.СтоимЧасаКоп as [коп], ДогПодряда.ОбъекОбщепита as [Объект], ДогПодрядаАкт.ПорНомерАкта as [№ акта],
ДогПодрядаАкт.ДатаАкта as [Дата акта], ДогПодрядаАкт.ЗаПериодС as [За период с], ДогПодрядаАкт.ЗаПериодПо as [За период по],
ДогПодрядаАкт.СтоимРабот as [Стоимость работ], ДогПодряда.Примечание, ДогПодрядаАкт.Код
FROM (Сотрудники INNER JOIN ДогПодряда ON Сотрудники.КодСотрудники = ДогПодряда.ID) INNER JOIN ДогПодрядаАкт ON ДогПодряда.Код = ДогПодрядаАкт.IDДогПодр
        WHERE Сотрудники.НазвОрганиз =@НазвОрганиз ORDER BY Сотрудники.ФИОСборное, ДогПодрядаАкт.ПорНомерАкта", list)

        com1 = True
        Grid4(ds2)
    End Sub
    '    Private Sub All()
    '        Dim StrSql2 As String = "SELECT Сотрудники.КодСотрудники, Сотрудники.НазвОрганиз as [Организация], Сотрудники.ФИОСборное as [ФИО], ДогПодряда.НомерДогПодр as [Номер ДогПодр],
    'ДогПодряда.ДатаДогПодр as [Дата], ДогПодряда.Должность, ДогПодряда.ДатаНачала as [Дата начала], ДогПодряда.ДатаОконч as [Дата оконч], ДогПодряда.СтоимЧасаРуб as [руб], ДогПодряда.СтоимЧасаКоп as [коп],
    'ДогПодряда.ОбъекОбщепита as [Объект], ДогПодряда.Примечание
    'FROM Сотрудники INNER JOIN ДогПодряда ON Сотрудники.КодСотрудники = ДогПодряда.ID
    'ORDER BY Сотрудники.НазвОрганиз"
    '        Dim ds2 As DataTable = Selects(StrSql2)

    '        grid2(ds2)
    '    End Sub

    Private Sub ComboBox19_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox19.SelectedIndexChanged
        ComboBox1.Text = ""
        Label96.Text = ComboBox2.Items.Item(ComboBox19.SelectedIndex)

        Dim list As New Dictionary(Of String, Object)
        list.Add("@КодСотрудники", CType(Label96.Text, Integer))


        Dim ds2 = Selects(StrSql:="SELECT Сотрудники.КодСотрудники, Сотрудники.НазвОрганиз as [Организация], Сотрудники.ФИОСборное as [ФИО],
ДогПодряда.НомерДогПодр as [Номер ДогПодр],ДогПодряда.ДатаДогПодр as [Дата], ДогПодряда.Должность,
ДогПодряда.ДатаНачала as [Дата начала], ДогПодряда.ДатаОконч as [Дата оконч], ДогПодряда.СтоимЧасаРуб as [руб], 
ДогПодряда.СтоимЧасаКоп as [коп], ДогПодряда.ОбъекОбщепита as [Объект], ДогПодрядаАкт.ПорНомерАкта as [№ акта],
ДогПодрядаАкт.ДатаАкта as [Дата акта], ДогПодрядаАкт.ЗаПериодС as [За период с], ДогПодрядаАкт.ЗаПериодПо as [За период по],
ДогПодрядаАкт.СтоимРабот as [Стоимость работ], ДогПодряда.Примечание, ДогПодрядаАкт.Код
FROM (Сотрудники INNER JOIN ДогПодряда ON Сотрудники.КодСотрудники = ДогПодряда.ID) INNER JOIN ДогПодрядаАкт ON ДогПодряда.Код = ДогПодрядаАкт.IDДогПодр
             WHERE Сотрудники.КодСотрудники=@КодСотрудники ORDER BY Сотрудники.НазвОрганиз, ДогПодрядаАкт.ПорНомерАкта", list)

        com19 = True
        grid3(ds2)
    End Sub

    Private Sub Grid1_DoubleClick(sender As Object, e As EventArgs) Handles Grid1.DoubleClick


    End Sub
    Private Function dsActM(ByVal ind As Integer)

        Dim list As New Dictionary(Of String, Object)
        list.Add("@КодСотрудники", ind)


        Dim dsAct = Selects(StrSql:="Select Сотрудники.КодСотрудники, Сотрудники.ФИОСборное, ДогПодряда.НомерДогПодр as [Номер договора подряда],
ДогПодряда.ДатаДогПодр as [Дата договора подряда], ДогПодрядаАкт.ПорНомерАкта as [Номер акта], ДогПодрядаАкт.ДатаАкта as [Дата акта],
ДогПодрядаАкт.ЗаПериодС as [Период с], ДогПодрядаАкт.ЗаПериодПо as [Период по],
ДогПодрядаАкт.СтоимРабот as [Стоимость], Сотрудники.НазвОрганиз
FROM(Сотрудники INNER JOIN ДогПодряда On Сотрудники.КодСотрудники = ДогПодряда.ID) INNER JOIN ДогПодрядаАкт On ДогПодряда.Код = ДогПодрядаАкт.IDДогПодр
            WHERE Сотрудники.КодСотрудники = @КодСотрудники", list)

        If errds = 1 Then
            dsAct.Clear()
            Return 0
        Else
            dsart = dsAct
            Return 1
        End If

    End Function

    Private Sub Grid1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellDoubleClick
        IdСот = Nothing
        IdСот = Grid1.CurrentRow.Cells("КодСотрудники").Value
        Dim y As Integer = dsActM(IdСот)

        If y = 1 Then
            ДоговорПодрядаСпискиАкты.ShowDialog()
        Else
            MessageBox.Show("С этим сотрудником не оформлялся акт выполненных работ!", Рик)
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Grid1.Rows.Count = 0 Then
            Exit Sub
        End If
        Grid1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        Grid1.SelectAll()

        Clipboard.SetDataObject(Grid1.GetClipboardContent())

        Dim путь1

        'If IO.File.Exists("C:\Users\Public\Documents\dgv.txt") = False Then
        '    IO.File.Copy(OnePath & "\ОБЩДОКИ\General\dgv.txt", "C:\Users\Public\Documents\dgv.txt")
        '    путь1 = "C:\Users\Public\Documents\dgv.txt"
        'Else
        '    путь1 = "C:\Users\Public\Documents\dgv.txt"
        'End If
        Начало("dgv.txt")
        путь1 = firthtPath & "\dgv.txt"

        'Записываем текст из буфера обмена в файл
        Using writer As New StreamWriter(путь1, False, System.Text.Encoding.Unicode)
            writer.Write(Clipboard.GetText())
        End Using

        'Process.Start(путь3, Chr(34) & путь1 & Chr(34))
        Process.Start("excel.exe", Chr(34) & путь1 & Chr(34))
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Grid1.Rows.Count = 0 Then
            Exit Sub
        End If

        Grid1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        Grid1.SelectAll()
        Clipboard.SetDataObject(Grid1.GetClipboardContent())

        Dim путь

        'If IO.File.Exists("C:\Users\Public\Documents\dgv.html") = False Then
        '    IO.File.Copy(OnePath & "\ОБЩДОКИ\General\dgv.html", "C:\Users\Public\Documents\dgv.html")
        '    путь = "C:\Users\Public\Documents\dgv.html"
        'Else
        '    путь = "C:\Users\Public\Documents\dgv.html"
        'End If
        Начало("dgv.html")
        путь = firthtPath & "\dgv.html"


        Using writer As New StreamWriter(путь, False, System.Text.Encoding.Unicode)
            writer.Write(Clipboard.GetText(TextDataFormat.Html))
        End Using
        Process.Start(путь)
    End Sub
    Private Function ОтборID() As Tuple(Of String, String, String)() 'Dictionary(Of String, String) 'List(Of String)
        Dim list As New Dictionary(Of String, Object)
        list.Add("@НазвОрганиз", ComboBox1.Text)


        dsId = Selects(StrSql:="SELECT DISTINCT ДогПодряда.НомерДогПодр, Сотрудники.КодСотрудники,
Сотрудники.НазвОрганиз, Сотрудники.ФИОСборное, ДогПодряда.ДатаДогПодр, ДогПодряда.Должность,
ДогПодряда.ДатаНачала, ДогПодряда.ДатаОконч, ДогПодряда.СтоимЧасаРуб, ДогПодряда.СтоимЧасаКоп, ДогПодряда.ОбъекОбщепита
FROM Сотрудники INNER JOIN ДогПодряда ON Сотрудники.КодСотрудники = ДогПодряда.ID
WHERE Сотрудники.НазвОрганиз=@НазвОрганиз", list)


        'Dim fdic As New Dictionary(Of String, String)()

        Dim cor(dsId.Rows.Count - 1) As Tuple(Of String, String, String)

        '        '"SELECT Сотрудники.КодСотрудники, Сотрудники.НазвОрганиз as [Организация], Сотрудники.ФИОСборное as [ФИО],
        '        ДогПодряда.НомерДогПодр as [Номер Договора Подряда],ДогПодряда.ДатаДогПодр as [Дата договора подряда], ДогПодряда.Должность,
        'ДогПодряда.ДатаНачала as [Дата начала ДогПодряда], ДогПодряда.ДатаОконч as [Дата оконч ДогПодряда], ДогПодряда.СтоимЧасаРуб as [руб], 
        'ДогПодряда.СтоимЧасаКоп as [коп]



        'Dim lt As New List(Of String)(ds.Rows.Count - 1)

        For d As Integer = 0 To dsId.Rows.Count - 1
            'fdic.Add(ds.Rows(d).Item(1).ToString, ds.Rows(d).Item(0).ToString)
            'lt.Add(ds.Rows(d).Item(1).ToString)
            cor(d) = New Tuple(Of String, String, String)(dsId.Rows(d).Item(1).ToString, dsId.Rows(d).Item(0).ToString, dsId.Rows(d).Item(3).ToString) 'КодСотрудника,НомерДог,ФИО сотр
        Next

        Return cor

    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim lisALL As Tuple(Of String, String, String)() = ОтборID() 'отобрали ИД номера всех сотрудников

        Dim f As DataTable = Orgall() 'отобрали все договора у кого есть акты выполн работ

        'Dim lis As New List(Of String)(f.Rows.Count - 1) 'отобрали ИД у всех у кого есть акты
        Dim lis(f.Rows.Count - 1) As Tuple(Of String, String, String)

        For d As Integer = 0 To f.Rows.Count - 1
            lis(d) = New Tuple(Of String, String, String)(f.Rows(d).Item(0).ToString, f.Rows(d).Item(3).ToString, f.Rows(d).Item(2).ToString) 'КодСотрудника,НомерДог,ФИО сотр
        Next

        Dim lisRep As New List(Of Integer)()

        For x As Integer = 0 To lisALL.Count - 1
            For m As Integer = 0 To lis.Count - 1
                'If lisALL.Item(x).ToString = lis.Item(m).ToString Then
                '    lisRep.Add(x)
                'End If
                If lisALL(x).Item1 = lis(m).Item1 Then
                    If lisALL(x).Item2 = lis(m).Item2 Then
                        If lisALL(x).Item3 = lis(m).Item3 Then
                            lisRep.Add(x)
                        End If
                    End If
                End If
            Next
        Next

        dsId.Columns(0).ColumnName = "Номер Договора Подряда"
        dsId.Columns(0).SetOrdinal(1)

        dsId.Columns(3).ColumnName = "ФИО"
        dsId.Columns(3).SetOrdinal(0)

        dsId.Columns(4).ColumnName = "Дата договора подряда"
        dsId.Columns(4).SetOrdinal(2)

        dsId.Columns(6).ColumnName = "Дата начала ДогПодряда"
        dsId.Columns(6).SetOrdinal(3)

        dsId.Columns(7).ColumnName = "Дата оконч ДогПодряда"
        dsId.Columns(7).SetOrdinal(4)

        For x As Integer = 0 To lisALL.Count - 1

            If lisRep.Contains(x) Then
                Continue For
            Else
                f.ImportRow(dsId.Rows(x))
            End If
        Next
        f.DefaultView.Sort = "ФИО" & " ASC" 'сортировка столбца по возрастанию datatable
        'f.DefaultView.Sort = "ФИО" & " DESC" 'сортировка столбца по убыванию datatable 

        'Grid1.DataSource = f
        grid2(f)
    End Sub

    Private Sub Grid1_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles Grid1.CellMouseClick

        For Each column As DataGridViewColumn In Grid1.Columns
            column.SortMode = DataGridViewColumnSortMode.NotSortable
        Next



        If e.Button = MouseButtons.Right Then
            clb2 = False
            Прим = Nothing
            Прим = 3
            IdСот = Nothing
            IdСот = Grid1.CurrentRow.Cells("КодСотрудники").Value
            Примечание.TextBox2.Text = Grid1.CurrentRow.Cells("ФИО").Value.ToString
            Примечание.Label2.Text = IdСот
            Примечание.RichTextBox1.Text = Grid1.CurrentRow.Cells(11).Value.ToString
            Примечание.ShowDialog()
        End If
    End Sub

    Private Sub Grid1_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles Grid1.ColumnHeaderMouseClick

        For Each column As DataGridViewColumn In Grid1.Columns
            column.SortMode = DataGridViewColumnSortMode.NotSortable
        Next

    End Sub

    Private Sub Grid1_ColumnHeaderMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles Grid1.ColumnHeaderMouseDoubleClick
        For Each column As DataGridViewColumn In Grid1.Columns
            column.SortMode = DataGridViewColumnSortMode.NotSortable
        Next
    End Sub


End Class