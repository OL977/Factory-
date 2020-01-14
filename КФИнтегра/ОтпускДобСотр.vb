Option Explicit On
Imports System.Data.OleDb
Public Class ОтпускДобСотр
    Dim strsql As String
    Dim dt, dt2 As DataTable
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Me.Close()


    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        com1select()



    End Sub

    Public Sub com1select()
        Dim strsql As String = "Select Штатное.Отдел as [Отдел], Штатное.Должность as [Должность],
Штатное.Разряд as [Разряд], Штатное.РасчДолжностнОклад as [РасчДолжнОклад], КарточкаСотрудника.ДатаПриема
FROM(Сотрудники INNER JOIN КарточкаСотрудника On Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр) INNER JOIN Штатное On Сотрудники.КодСотрудники = Штатное.ИДСотр
WHERE Сотрудники.НазвОрганиз  = '" & Отпуск1.ComboBox2.Text & "' And Сотрудники.ФИОСборное = '" & Me.ComboBox1.Text & "'"
        dt = Selects(strsql)
        Grid1.DataSource = dt
        Grid1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Grid1.Columns(4).Visible = False
    End Sub

    Private Sub ОтпускДобСотр_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        Me.Close()


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ComboBox1.Text = "" Then
            MessageBox.Show("Выберите сотрудника!", Рик)
            Exit Sub
        End If

        If TextBox1.Text = "" Then
            MessageBox.Show("Выберите срок положенного отпуска!", Рик)
            Exit Sub
        End If

        Me.Close()
        refreshin()
        Отпуск1.grcellclick()
        Отпуск1.ПерзагрGrid1()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
        Обнов()
        Отпуск1.grcellclick()
    End Sub
    Private Sub Обнов()
        strsql = "Select Штатное.Отдел as [Отдел], Штатное.Должность as [Должность],
Штатное.Разряд as [Разряд], Штатное.РасчДолжностнОклад as [РасчДолжнОклад], КарточкаСотрудника.ДатаПриема
FROM(Сотрудники INNER JOIN КарточкаСотрудника On Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр) INNER JOIN Штатное On Сотрудники.КодСотрудники = Штатное.ИДСотр
WHERE Сотрудники.НазвОрганиз  = '" & Отпуск1.ComboBox2.Text & "' And Сотрудники.ФИОСборное = '" & Me.ComboBox1.Text & "'"
        Dim tr As DataTable = Selects(strsql)

        strsql = ""
        strsql = "UPDATE ОтпускСотрудники SET КолДнейОтпуска='" & TextBox1.Text & "', ДатаПриема = '" & tr.Rows(0).Item(4).ToString & "',
ПериодС='" & Отпуск1.ПерС & "', ПериодПо=" & CType(Отпуск1.ПерС, Integer) + 1 & " , Итого = '" & TextBox1.Text & "', ОсталосьПрошлГод='0'
        WHERE ОтпускСотрудники.Код=" & Отпуск1.idwor & ""
        Selects(strsql)


    End Sub
    Private Sub refreshin()

        Dim Нараб As Date = CDate(dt.Rows(0).Item(4).ToString)
        Dim dn As Date = Now
        Dim dspan As TimeSpan = dn - Нараб
        Dim ДнОтпус As Integer
        Dim ЧислЦел As Integer = Math.Floor(dspan.TotalDays / 30)
        ДнОтпус = ЧислЦел * 2
        Dim lp As Integer = CType(Отпуск1.ПерС, Integer)
        lp = lp + 1

        Dim list As New Dictionary(Of String, Object)
        list.Add("@IDОтпуск", Отпуск1.indrow)


        Updates(stroka:="INSERT INTO ОтпускСотрудники(IDОтпуск,Отдел,Должность,ФИО,КолДнейОтпуска,ДатаПриема,ПериодС,ПериодПо,Нарботано,ОсталосьПрошлГод,Итого)
VALUES(" & Отпуск1.indrow & ", '" & dt.Rows(0).Item(0).ToString & "','" & dt.Rows(0).Item(1).ToString & "','" & ComboBox1.Text & "', '" & TextBox1.Text & "',
'" & dt.Rows(0).Item(4).ToString & "', '" & Отпуск1.ПерС & "', " & lp & ", " & ДнОтпус & ",'0','" & TextBox1.Text & "')", list, "dtOtpuskAll")



        Dim dt2 = Selects(StrSql:="SELECT Отдел,Должность,ФИО,КолДнейОтпуска as [Кол-во дней отпуск],ДатаПриема as [Дата приема],
ПериодС as [Период с],ПериодПо as [Период по], Нарботано as [Наработ дней отпуска], Примечание
FROM ОтпускСотрудники WHERE IDОтпуск=@IDОтпуск", list)
        Отпуск1.Grid2.DataSource = dt2
        Отпуск1.Grid2.Columns(2).Width = 200
        Отпуск1.Grid2.Columns(0).Width = 100
        Отпуск1.Grid2.Columns(1).Width = 100

        'Отпуск.Grid2.Columns(1).Visible = False
    End Sub

    Private Sub Grid1_Click(sender As Object, e As EventArgs) Handles Grid1.Click

    End Sub
End Class