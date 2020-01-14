Option Explicit On
Imports System.Data.OleDb
Public Class Отпуск1
    Public dt, dt2, dt3, dt4, dt5 As DataTable
    Dim StrSql As String
    Public indrow, idwor, idgr3cod, ДнОтпус As Integer
    Public ПерС As String

    Private Sub Отпуск1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MdiParent = MDIParent1
        Me.WindowState = FormWindowState.Maximized

        Me.ComboBox2.AutoCompleteCustomSource.Clear()
        Me.ComboBox2.Items.Clear()
        For Each r As DataRow In СписокКлиентовОсновной.Rows
            Me.ComboBox2.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox2.Items.Add(r(0).ToString)
        Next



    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ComboBox2.Text = "" Then
            MessageBox.Show("Выберите организацию!", Рик)
            Exit Sub
        End If
        ОтпускНовыйГрафик.ShowDialog()
        refreshus()
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub
    Public Sub ВыборСотр()
        Try
            dt2.Clear()
        Catch ex As Exception

        End Try
        ОтпускДобСотр.ComboBox1.Items.Clear()
        Dim no As String = "Нет"
        StrSql = ""
        StrSql = "SELECT Сотрудники.ФИОСборное
FROM Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр
WHERE Сотрудники.НазвОрганиз='" & ComboBox2.Text & "' AND Сотрудники.НаличеДогПодряда='" & no & "' AND КарточкаСотрудника.ДатаУвольнения is Null 
ORDER BY Сотрудники.ФИОСборное"
        dt2 = Selects(StrSql)
        ОтпускДобСотр.ComboBox1.Text = ""
        ОтпускДобСотр.ComboBox1.AutoCompleteCustomSource.Clear()

        For Each r As DataRow In dt2.Rows
            ОтпускДобСотр.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            ОтпускДобСотр.ComboBox1.Items.Add(r(0).ToString)
        Next

    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        refreshus()

    End Sub
    Public Sub ПерзагрGrid1()
        grcellclick()
        grid3activ()
        Grid2.Visible = True
    End Sub
    Private Sub Grid1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellClick
        'Grid1.ClearSelection()
        ПерзагрGrid1()
    End Sub
    Public Sub grid3activ()


        Dim list As New Dictionary(Of String, Object)
        list.Add("@IDОтпуск", indrow)

        Dim dt4 = Selects(StrSql:="SELECT ФИО, ДатаНач1 as [Дата начала первой части отпуска], Продолж1 as [Продолж отпуска],
ДатаОконч1 as [Дата оконч первой части отпуска], ДатаНач2 as [Дата начала второй части отпуска], Продолж2 as [Продол отпуска],
ДатаОконч2 as [Дата оконч второй части отпуска], КолДнейОтпуска as [Положено дней отпуска], Израсходовано as [Использ], ОсталосьЭтотГод as [Остаток за этот год],
ОсталосьПрошлГод as [Остаток прошл год], Код, Итого
FROM ОтпускСотрудники WHERE IDОтпуск=@IDОтпуск", list)

        Grid3.DataSource = dt4
        Grid3.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Grid3.Columns(11).Visible = False
        Grid3.Columns(0).Width = 300
        Grid3.SelectionMode = DataGridViewSelectionMode.FullRowSelect


    End Sub
    Public Sub grcellclick()

        indrow = 0

        Try
            indrow = Grid1.CurrentRow.Cells("Код").Value
        Catch ex As Exception
            MessageBox.Show("Выберите график?", Рик)
            Exit Sub
        End Try

        ПерС = Grid1.CurrentRow.Cells("НаГод").Value
        GroupBox2.Text = "Работа с графиком №" & Grid1.CurrentRow.Cells("Номер").Value & " за " & Grid1.CurrentRow.Cells("НаГод").Value
        Dim list As New Dictionary(Of String, Object)
        list.Add("@IDОтпуск", indrow)

        Dim dt3 = Selects(StrSql:="SELECT Отдел,Должность,ФИО,КолДнейОтпуска as [Кол-во дней отпуск],ДатаПриема as [Дата приема],
ПериодС as [Период с],ПериодПо as [Период по],Нарботано as [Наработ дней отпуска],Примечание,Код
FROM ОтпускСотрудники WHERE IDОтпуск=@IDОтпуск", list)
        Grid2.DataSource = dt3
        Grid2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Grid2.Columns(2).Width = 200
        Grid2.Columns(0).Width = 100
        Grid2.Columns(1).Width = 100
        Grid2.Columns(9).Visible = False
        Grid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Grid2.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If MessageBox.Show("Удалить график отпусков?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Exit Sub
        End If

        StrSql = "delete FROM Отпуск WHERE Код=" & indrow & ""
        Updates(StrSql)
        refreshus()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If MessageBox.Show("Удалить cотрудника?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Exit Sub
        End If
        StrSql = ""
        StrSql = "delete FROM ОтпускСотрудники WHERE Код=" & Grid2.CurrentRow.Cells("Код").Value & ""
        Updates(StrSql)
        StrSql = ""
        grcellclick()
        Grid2.Visible = True
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If MessageBox.Show("Удалить всех сотрудников?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Exit Sub
        End If
        StrSql = ""
        StrSql = "DELETE FROM ОтпускСотрудники WHERE IDОтпуск=" & indrow & ""
        Updates(StrSql)
        StrSql = ""
        grcellclick()
        Grid2.Visible = True
    End Sub

    Public Sub refreshus()

        'If dtOtpuskAll Is Nothing Or dtOtpuskAll.Rows.Count = 0 Or IsDBNull(dtOtpuskAll) Then Exit Sub
        'Dim dg = From x In dtOtpuskAll Order By x.Item("Год") Select x

        Dim list As New Dictionary(Of String, Object)
        list.Add("@Орг", ComboBox2.Text)

        dt = Selects(StrSql:="SELECT Код, НаГод, Номер, Составлен, Утвержден FROM Отпуск WHERE Орг=@Орг ORDER BY НаГод", list)
        'dt = Selects(StrSql)
        Grid1.DataSource = dt
        Grid1.Columns(0).Visible = False
        Grid2.Visible = False
        NameOrg = ComboBox2.Text
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If ComboBox2.Text = "" Then
            MessageBox.Show("Выберите организацию!", Рик)
            Exit Sub
        End If
        ОтпускДобСотр.Button1.Visible = True
        ОтпускДобСотр.Button3.Visible = False
        ВыборСотр()
        ОтпускДобСотр.ShowDialog()
    End Sub

    Private Sub Grid2_DoubleClick(sender As Object, e As EventArgs) Handles Grid2.DoubleClick

        ОтпускДобСотр.ComboBox1.Text = Grid2.CurrentRow.Cells("ФИО").Value
        idwor = Grid2.CurrentRow.Cells("Код").Value
        ОтпускДобСотр.com1select()

        ОтпускДобСотр.Button1.Visible = False
        ОтпускДобСотр.Button3.Visible = True
        ОтпускДобСотр.ShowDialog()
        ПерзагрGrid1()

    End Sub

    Private Sub Grid3_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid3.CellDoubleClick
        Grid3Celcl()
    End Sub
    Public Sub Grid3Celcl()
        Grid3.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        idgr3cod = Grid3.CurrentRow.Cells("Код").Value
        ОтпускНачало.TextBox1.Text = Grid3.CurrentRow.Cells("ФИО").Value

        Try
            ДнОтпус = Grid3.CurrentRow.Cells("Положено дней отпуска").Value
        Catch ex As Exception
            MessageBox.Show("Введите планируемое количество дней отпуска за год!", Рик)
            Exit Sub
        End Try

        StrSql = ""
        StrSql = "SELECT * FROM ОтпускСотрудники WHERE Код=" & idgr3cod & ""
        dt5 = Selects(StrSql)
        ОтпускНачало.MaskedTextBox1.Text = dt5.Rows(0).Item(11).ToString
        ОтпускНачало.TextBox2.Text = dt5.Rows(0).Item(12).ToString
        ОтпускНачало.MaskedTextBox2.Text = dt5.Rows(0).Item(14).ToString
        ОтпускНачало.TextBox3.Text = dt5.Rows(0).Item(15).ToString


        If dt5.Rows(0).Item(20).ToString = "" Then
            ОтпускНачало.TextBox4.Text = "0"
        Else
            ОтпускНачало.TextBox4.Text = dt5.Rows(0).Item(20).ToString
        End If
        ОтпускНачало.ShowDialog()
    End Sub

    Private Sub Grid2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid2.CellClick
        Grid2.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    End Sub
End Class