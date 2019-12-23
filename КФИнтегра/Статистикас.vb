Option Explicit On
Imports System.Data.OleDb
Public Class Статистикаc
    Dim strsql As String
    Dim ds, ds1 As DataTable
    Dim времянач As String
    Private Sub Статистика_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'conn = New OleDbConnection
        'conn.ConnectionString = ConString
        'Try
        '    conn.Open()
        'Catch ex As Exception
        '    MessageBox.Show("Не подключен диск U")
        'End Try

        времянач = ""
        времянач = Format(Now.Date, "MM\/dd\/yyyy")

        Dim list As New Dictionary(Of String, Object)
        list.Add("@Дата", Now.Date)


        Dim ds2 = Selects(StrSql:="SELECT * FROM Статистика WHERE Дата =@Дата", list)
        'ds2 = Selects(strsql) 'Дата like #" & времянач & "#"
        Grid1.DataSource = ds2
        Grid1.Columns(0).Visible = False
        Try
            Grid1.Columns(1).Width = 80
            Grid1.Columns(2).Width = 80
        Catch ex As Exception

        End Try


        Dim ds1 = From x In dtStatistikaAll Order By x.Item("КемИзменено") Descending Select x.Item("КемИзменено") Distinct
        'Dim strsql1 As String = "SELECT DISTINCT КемИзменено FROM Статистика ORDER BY КемИзменено"
        'ds1 = Selects(strsql1)
        Me.ComboBox1.AutoCompleteCustomSource.Clear()
        Me.ComboBox1.Items.Clear()
        For Each r In ds1
            Me.ComboBox1.AutoCompleteCustomSource.Add(r.ToString())
            Me.ComboBox1.Items.Add(r.ToString)
        Next


    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        'времянач = Nothing
        'времянач = Format(DateTimePicker1.Value, "MM\/dd\/yyyy")

        Dim list As New Dictionary(Of String, Object)
        list.Add("@времянач", DateTimePicker1.Value.ToShortDateString)
        list.Add("@КемИзменено", ComboBox1.Text)

        Dim ds = Selects(StrSql:="SELECT * FROM Статистика WHERE Дата=@времянач AND КемИзменено=@КемИзменено", list)
        'ds = Selects(strsql) 'Дата like #" & времянач & "#"
        Грид(ds)
        CheckBox1.Checked = False
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        'времянач = Nothing
        'времянач = Format(DateTimePicker1.Value, "MM\/dd\/yyyy")
        'If ComboBox1.Text = "" Then
        '    Try
        '        ds.Clear()
        '        strsql = ""
        '    Catch ex As Exception

        '    End Try
        Dim list As New Dictionary(Of String, Object)
        list.Add("@времянач", DateTimePicker1.Value.ToShortDateString)
        list.Add("@КемИзменено", ComboBox1.Text)

        If ComboBox1.Text = "" Then


            Dim ds = Selects(StrSql:="SELECT * FROM Статистика WHERE Дата=@времянач", list)
            'ds = Selects(strsql) 'Дата like #" & времянач & "#"
            Грид(ds)
            CheckBox1.Checked = False
        Else

            Dim ds = Selects(StrSql:="SELECT * FROM Статистика WHERE Дата=@времянач AND КемИзменено=@КемИзменено", list)
            ds = Selects(strsql) 'Дата like #" & времянач & "#"
            Грид(ds)
            CheckBox1.Checked = False
        End If

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

        If CheckBox1.Checked = True Then

            Dim ds = From x In dtStatistikaAll Order By x.Item("Дата") Descending Select x
            'ds = Selects(strsql)
            If ds.Count = 0 Then
                Exit Sub
            End If
            Grid1.DataSource = ds.CopyToDataTable
            Grid1.Columns(0).Visible = False

            Grid1.Columns(1).Width = 80
            Grid1.Columns(2).Width = 80

        End If
    End Sub
    Private Sub Грид(ByVal ds As DataTable)
        Grid1.DataSource = ds
        If ds.Rows.Count = 0 Then
            Exit Sub
        End If
        Grid1.Columns(0).Visible = False

        Grid1.Columns(1).Width = 80
        Grid1.Columns(2).Width = 80


    End Sub

    Private Sub Статистикаc_Closed(sender As Object, e As EventArgs) Handles Me.Closed

    End Sub
End Class