Option Explicit On
Imports System.Data.OleDb
Public Class Иностранцы

    Private Sub Иностранцы_Load(sender As Object, e As EventArgs) Handles MyBase.Load



        Me.ComboBox1.AutoCompleteCustomSource.Clear()
        Me.ComboBox1.Items.Clear()
        For Each r As DataRow In СписокКлиентовОсновной.Rows
            Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox1.Items.Add(r(0).ToString)
        Next
        RunMoving2() 'обновляем сотрудников


        Dim ds1 = From x In dtSotrudnikiAll Where x.Item("Иностранец") = "True" Order By x.Item("ФИОСборное") Select (x.Item("ФИОСборное"), x.Item("КодСотрудники"))

        'Dim StrSql1 As String = "SELECT ФИОСборное,КодСотрудники FROM Сотрудники WHERE Иностранец = True ORDER BY ФИОСборное"
        'Dim ds1 As DataTable = Selects(StrSql1)
        Me.ComboBox19.AutoCompleteCustomSource.Clear()
        Me.ComboBox19.Items.Clear()
        Me.ComboBox2.Items.Clear()

        For Each r In ds1
            Me.ComboBox19.AutoCompleteCustomSource.Add(r.Item1.ToString)
            Me.ComboBox19.Items.Add(r.Item1.ToString)
            Me.ComboBox2.Items.Add(r.Item2.ToString)
        Next

        Dim ds2 = Selects(StrSql:="SELECT Сотрудники.КодСотрудники, Сотрудники.НазвОрганиз as [Организация], Сотрудники.ФИОСборное as [ФИО], КарточкаСотрудника.ДатаПриема as [Дата приема], КарточкаСотрудника.ДатаУвольнения as [Дата увольнения], ДогСотрудн.Контракт, Штатное.Отдел, Штатное.Должность, Штатное.Разряд
FROM((Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр) INNER JOIN ДогСотрудн On Сотрудники.КодСотрудники = ДогСотрудн.IDСотр) INNER JOIN Штатное On Сотрудники.КодСотрудники = Штатное.ИДСотр
WHERE Сотрудники.Иностранец = True ORDER BY Сотрудники.НазвОрганиз")

        If ds2.Rows.Count > 0 Then
            grid(ds2)
        End If


    End Sub
    Public Sub grid(ByVal ds2 As DataTable)
        If ds2.Rows.Count = 0 Then
            'MessageBox.Show("Нет данных!", Рик)
            Exit Sub
        End If
        Grid1.DataSource = ds2
        GridView(Grid1)
        Grid1.Columns(0).Visible = False
        Grid1.Columns(1).Width = 200
        Grid1.Columns(2).Width = 200

        'Dim strikethrough_style As New DataGridViewCellStyle

        'strikethrough_style.Font = New Font("Times New Roman", 11, FontStyle.Regular)


        'Try
        '    Grid1.DefaultCellStyle = strikethrough_style
        '    Grid1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        '    Grid1.Columns(0).Visible = False
        '    Grid1.Columns(1).Width = 200
        '    Grid1.Columns(2).Width = 200
        '    Grid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        'Catch ex As Exception

        'End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim list As New Dictionary(Of String, Object)
        list.Add("@НазвОрганиз", ComboBox1.Text)
        list.Add("@Иностранец", "True")

        Dim ds2 = Selects(StrSql:="SELECT Сотрудники.КодСотрудники, Сотрудники.НазвОрганиз as [Организация], Сотрудники.ФИОСборное as [ФИО], КарточкаСотрудника.ДатаПриема as [Дата приема],КарточкаСотрудника.ДатаУвольнения as [Дата увольнения], ДогСотрудн.Контракт, Штатное.Отдел, Штатное.Должность, Штатное.Разряд
FROM((Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр) INNER JOIN ДогСотрудн On Сотрудники.КодСотрудники = ДогСотрудн.IDСотр) INNER JOIN Штатное On Сотрудники.КодСотрудники = Штатное.ИДСотр
WHERE Сотрудники.НазвОрганиз=@НазвОрганиз AND Сотрудники.Иностранец=@Иностранец ORDER BY Сотрудники.НазвОрганиз", list)

        grid(ds2)
    End Sub

    Private Sub ComboBox19_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox19.SelectedIndexChanged
        Label96.Text = ComboBox2.Items.Item(ComboBox19.SelectedIndex)

        Dim list As New Dictionary(Of String, Object)
        list.Add("@КодСотрудники", CType(Label96.Text, Integer))
        ComboBox1.Text = ""

        Dim ds2 = Selects(StrSql:="SELECT Сотрудники.КодСотрудники, Сотрудники.НазвОрганиз as [Организация], Сотрудники.ФИОСборное as [ФИО], КарточкаСотрудника.ДатаПриема as [Дата приема],КарточкаСотрудника.ДатаУвольнения as [Дата увольнения], ДогСотрудн.Контракт, Штатное.Отдел, Штатное.Должность, Штатное.Разряд
FROM((Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр) INNER JOIN ДогСотрудн On Сотрудники.КодСотрудники = ДогСотрудн.IDСотр) INNER JOIN Штатное On Сотрудники.КодСотрудники = Штатное.ИДСотр
WHERE Сотрудники.КодСотрудники=@КодСотрудники ORDER BY Сотрудники.НазвОрганиз", list)

        grid(ds2)
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged



        Dim list As New Dictionary(Of String, Object)
        list.Add("@Иностранец", "True")



        If CheckBox1.Checked = True Then
            Dim ds2 = Selects(StrSql:="SELECT Сотрудники.КодСотрудники, Сотрудники.НазвОрганиз as [Организация], Сотрудники.ФИОСборное as [ФИО], КарточкаСотрудника.ДатаПриема as [Дата приема],КарточкаСотрудника.ДатаУвольнения as [Дата увольнения], ДогСотрудн.Контракт, Штатное.Отдел, Штатное.Должность, Штатное.Разряд
FROM((Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр) INNER JOIN ДогСотрудн On Сотрудники.КодСотрудники = ДогСотрудн.IDСотр) INNER JOIN Штатное On Сотрудники.КодСотрудники = Штатное.ИДСотр
WHERE Сотрудники.Иностранец = @Иностранец ORDER BY Сотрудники.НазвОрганиз", list)


            grid(ds2)
        Else
            If ComboBox1.Text = "" And ComboBox19.Text = "" Then
                Dim ds2 = Selects(StrSql:="SELECT Сотрудники.КодСотрудники, Сотрудники.НазвОрганиз as [Организация], Сотрудники.ФИОСборное as [ФИО], КарточкаСотрудника.ДатаПриема as [Дата приема],КарточкаСотрудника.ДатаУвольнения as [Дата увольнения], ДогСотрудн.Контракт, Штатное.Отдел, Штатное.Должность, Штатное.Разряд
FROM((Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр) INNER JOIN ДогСотрудн On Сотрудники.КодСотрудники = ДогСотрудн.IDСотр) INNER JOIN Штатное On Сотрудники.КодСотрудники = Штатное.ИДСотр
WHERE Сотрудники.Иностранец=@Иностранец ORDER BY Сотрудники.НазвОрганиз", list)


                grid(ds2)
            ElseIf ComboBox1.Text <> "" Then
                list.Add("@НазвОрганиз", ComboBox1.SelectedItem)
                Dim ds = Selects(StrSql:="SELECT Сотрудники.КодСотрудники, Сотрудники.НазвОрганиз as [Организация], Сотрудники.ФИОСборное as [ФИО], КарточкаСотрудника.ДатаПриема as [Дата приема],КарточкаСотрудника.ДатаУвольнения as [Дата увольнения], ДогСотрудн.Контракт, Штатное.Отдел, Штатное.Должность, Штатное.Разряд
FROM((Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр) INNER JOIN ДогСотрудн On Сотрудники.КодСотрудники = ДогСотрудн.IDСотр) INNER JOIN Штатное On Сотрудники.КодСотрудники = Штатное.ИДСотр
WHERE Сотрудники.НазвОрганиз=@НазвОрганиз AND Сотрудники.Иностранец=@Иностранец ORDER BY Сотрудники.НазвОрганиз", list)
                grid(ds)
            Else
                Label96.Text = ComboBox2.Items.Item(ComboBox19.SelectedIndex)
                list.Add("@КодСотрудники", CType(Label96.Text, Integer))
                Dim ds1 = Selects(StrSql:="SELECT Сотрудники.КодСотрудники, Сотрудники.НазвОрганиз as [Организация], Сотрудники.ФИОСборное as [ФИО], КарточкаСотрудника.ДатаПриема as [Дата приема],КарточкаСотрудника.ДатаУвольнения as [Дата увольнения], ДогСотрудн.Контракт, Штатное.Отдел, Штатное.Должность, Штатное.Разряд
FROM((Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр) INNER JOIN ДогСотрудн On Сотрудники.КодСотрудники = ДогСотрудн.IDСотр) INNER JOIN Штатное On Сотрудники.КодСотрудники = Штатное.ИДСотр
WHERE Сотрудники.КодСотрудники=@КодСотрудники ORDER BY Сотрудники.НазвОрганиз", list)
                grid(ds1)
            End If

        End If
    End Sub
End Class