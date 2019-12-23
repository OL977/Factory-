Option Explicit On
Imports System.ComponentModel
Imports System.Data.OleDb
Public Class ВсплывФормаПриЗагр

    Dim StrSql As String
    'Private Delegate Sub grbox()
    'Private Delegate Sub grbox2()
    Private Sub ВсплывФормаПриЗагр_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'ВсплывФорма()
    End Sub

    Public Sub ВсплывФорма()
        'Me.Hide()

        'Await Task.Delay(2000)
        gr1 = 0
        gr2 = 0
        'Соед(0)
        TextBox1.Text = Date.Now.ToShortDateString
        Dim нач = Now.Date
        Dim кон = Now.Date
        нач = нач.AddDays(5)
        кон = кон.AddDays(1)
        Dim нач1 As String = нач.AddDays(5).ToShortDateString
        Dim кон1 As String = кон.AddDays(1).ToShortDateString
        'Dim DateEx As String = Format(кон, "MM\/dd\/yyyy")
        'Dim DateEx2 As String = Format(нач, "MM\/dd\/yyyy")
        'Dim DateEx As String = Replace(Format(кон, "yyyy\/MM\/dd"), "/", "")
        'Dim DateEx2 As String = Replace(Format(нач, "yyyy\/MM\/dd"), "/", "")

        Dim list As New Dictionary(Of String, Object)
        list.Add("@ДатаУведомлПродКонтр", кон1)
        list.Add("@ДатаУведомлПрод", нач1)
        list.Add("@Now", Now.Date.ToShortDateString)

        Dim ds = Selects(StrSql:="SELECT Сотрудники.НазвОрганиз, Сотрудники.ФИОСборное, КарточкаСотрудника.ДатаУведомлПродКонтр, КарточкаСотрудника.СрокПродлКонтракта
FROM Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр
WHERE КарточкаСотрудника.ДатаУведомлПродКонтр=@ДатаУведомлПродКонтр ORDER BY Сотрудники.НазвОрганиз", list)

        If ds.Rows.Count > 0 Then
            Grid1.DataSource = ds
            'Grid1.Columns(0).Width = 220
            'Grid1.Columns(1).Width = 320
            GroupBox1.Text = "Остался 1 день до уведомления о продлении контракта!"

        Else
            'If GroupBox1.InvokeRequired Then
            '    Me.Invoke(New grbox(AddressOf ВсплывФорма))
            'Else
            GroupBox1.Visible = False
            'End If
            gr1 = 1
        End If

        RunMoving2()
        RunMoving6()

        'Dim ds2 = Selects(StrSql:="SELECT Сотрудники.НазвОрганиз, Сотрудники.ФИОСборное,
        'КарточкаСотрудника.ДатаУведомлПродКонтр, КарточкаСотрудника.СрокПродлКонтракта
        'FROM Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр
        'WHERE КарточкаСотрудника.ДатаУведомлПродКонтр=@ДатаУведомлПрод ORDER BY Сотрудники.НазвОрганиз", list)

        Dim ds2 = (From x In dtSotrudnikiAll.AsEnumerable()'linq datatble
                   Join y In dtKartochkaSotrudnikaAll.AsEnumerable() On x.Field(Of Integer)("КодСотрудники") Equals
                      y.Field(Of Integer)("IDСотр")
                   Where y.Field(Of Object)("ДатаУведомлПродКонтр") = нач1 And y.Field(Of Object)("ДатаУвольнения") Is Nothing
                   Order By x.Item("НазвОрганиз")
                   Select New With {.Организация = x.Field(Of String)("НазвОрганиз"), .ФИО = x.Field(Of String)("ФИОСборное"),
                       .Дата_Уведомления = y.Field(Of Object)("ДатаУведомлПродКонтр"), .СрокКонтракта = y.Field(Of Object)("СрокПродлКонтракта")}).ToList

        'Select (x.Field(Of String)("НазвОрганиз"), x.Field(Of String)("ФИОСборное"), y.Field(Of String)("ДатаУведомлПродКонтр"), y.Field(Of String)("СрокПродлКонтракта"))






        If ds2.Count > 0 Then
            Grid2.DataSource = ds2
            'Grid2.Columns(0).Width = 220
            'Grid2.Columns(1).Width = 320
            GroupBox2.Text = "Осталось 5 дней до уведомления о продлении контракта!"

        Else
            'If GroupBox2.InvokeRequired Then
            '    Me.Invoke(New grbox2(AddressOf ВсплывФорма))
            'Else
            GroupBox2.Visible = False
            'End If
            gr2 = 1
        End If


        If gr1 = 0 Or gr2 = 0 Then
            'Dim fmz As String = Replace(Format(Now, "yyyy\/MM\/dd"), "/", "")
            Dim ds3 = Selects(StrSql:="select Код FROM ШтатРаспис WHERE ДатаСформШт=@Now", list)

            If errds = 1 Then

                'Me.ShowDialog()
                Updates(stroka:="INSERT INTO ШтатРаспис(ДатаСформШт) VALUES('" & Now.ToShortDateString & "')", list, "ШтатРаспис")

            End If
            Me.ShowDialog()
        ElseIf gr1 = 1 And gr2 = 1 Then
            'Me.Close()

            'Me.Show()
        End If
        'Соед(1)
    End Sub

    Private Sub Grid2_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid2.CellDoubleClick

    End Sub

    Private Sub ВсплывФормаПриЗагр_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Me.Hide()








        'While (Grid1.Rows.Count > 0)
        '        For i As Integer = 0 To Grid1.Rows.Count - 1
        '            Grid1.Rows.Remove(Grid1.Rows(i))
        '        Next
        '    End While


        'While (Grid2.Rows.Count > 0)
        '    For id As Integer = 0 To Grid2.Rows.Count - 1
        '        Grid2.Rows.Remove(Grid2.Rows(id))
        '    Next
        'End While
    End Sub
End Class