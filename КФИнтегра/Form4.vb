Option Explicit On
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Threading
Imports System.Data.SqlClient
Public Class Контрагент

    Public Da As New OleDbDataAdapter 'Адаптер
    Public ds As New DataTable 'Пустой набор записей
    Dim tbl As New DataTable
    Dim cb As OleDb.OleDbCommandBuilder
    Dim МРаботы As String

    Dim ЮрЛицо, ФизЛиц, Клиент, CorName, CorOtch, ФАдрес, ФПочт, batclick, Организ As String
    Dim s, s2, Fh As Integer
    Dim sw = New Stopwatch() 'замер времнеи работы процедуры

    Private Sub ВстВБазуНовКонтр()
        Fh = Nothing
        ПровЗапПолей()
        If Fh = 1 Then Exit Sub

        Dim StrSql As String = "SELECT * FROM Клиент"

        Dim conn As New SqlConnection(ConString)
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        'Dim c As New OleDbCommand
        Dim c As New SqlCommand(StrSql, conn)

        Dim ds As New DataSet
        Dim da As New SqlDataAdapter(c)
        da.Fill(ds, "Контрагент")

        Dim cb As New SqlCommandBuilder(da)
        Dim dsNewRow As DataRow
        dsNewRow = ds.Tables("Контрагент").NewRow()

        dsNewRow.Item("НазвОрг") = Trim(TextBox1.Text)
        dsNewRow.Item("ФормаСобств") = Me.ComboBox2.Text
        dsNewRow.Item("УНП") = Trim(Me.TextBox3.Text)
        dsNewRow.Item("ФИОРуководителя") = Me.TextBox4.Text
        dsNewRow.Item("ДолжнРуководителя") = Me.RichTextBox1.Text
        dsNewRow.Item("ОснованиеДейств") = Me.TextBox6.Text
        dsNewRow.Item("ТелРуков") = Me.TextBox10.Text
        dsNewRow.Item("ФИОДопЛица") = Me.TextBox9.Text
        dsNewRow.Item("ДолжнДопЛица") = Me.TextBox8.Text
        dsNewRow.Item("ТелДопЛица") = Me.TextBox11.Text
        dsNewRow.Item("ЮрАдрес") = Me.TextBox7.Text
        dsNewRow.Item("ФактичАдрес") = ФАдрес
        dsNewRow.Item("ПочтАдрес") = ФПочт
        dsNewRow.Item("КонтТелефон") = Me.TextBox14.Text
        dsNewRow.Item("Факс") = Me.TextBox15.Text
        dsNewRow.Item("ЭлАдрес") = Me.TextBox16.Text
        dsNewRow.Item("ДругиеКонтакты") = Me.TextBox17.Text
        dsNewRow.Item("Банк") = Me.ComboBox3.Text
        dsNewRow.Item("БИКБанка") = Me.TextBox18.Text
        dsNewRow.Item("АдресБанка") = Me.ComboBox4.Text
        dsNewRow.Item("Отделение") = Me.ComboBox4.Text
        dsNewRow.Item("РасчСчетРубли") = Me.TextBox22.Text
        dsNewRow.Item("РасчСчетЕвро") = Me.TextBox23.Text
        dsNewRow.Item("РасчСчетДоллар") = Me.TextBox24.Text
        dsNewRow.Item("РасчСчетРоссРубли") = Me.TextBox25.Text
        dsNewRow.Item("Операционист") = Me.TextBox26.Text
        dsNewRow.Item("КонтТелОпер") = Me.TextBox27.Text
        dsNewRow.Item("ФИОРукРодПадеж") = Me.TextBox28.Text
        dsNewRow.Item("ФИОРукДатПадеж") = Me.TextBox2.Text

        If ЮрЛицо = 1 Then
            dsNewRow.Item("ЮрЛицо") = 1
            dsNewRow.Item("ФизЛицо") = 0
        Else
            dsNewRow.Item("ФизЛицо") = 1
            dsNewRow.Item("ЮрЛицо") = 0
        End If

        dsNewRow.Item("РасчСчетДоллар") = Me.TextBox24.Text
        dsNewRow.Item("РукИП") = CheckBox5.Checked

        ds.Tables("Контрагент").Rows.Add(dsNewRow)




        Try
            da.Update(ds, "Контрагент")
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Такой котрагент уже существует!" & vbCrLf & "Редактируйте его с выбором из уже сущствующих контрагентов", Рик, MessageBoxButtons.OK, MessageBoxIcon.Stop)
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            Exit Sub
            Fh = 1
        End Try



    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If sender Is Button5 Then
            ФизЛиц = 1
            Button4.BackColor = Color.FromArgb(255, 255, 192)
        End If
        ЮрЛицо = 0
        Button5.BackColor = Color.FromArgb(152, 251, 152)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If sender Is Button4 Then
            ЮрЛицо = 1
            Button5.BackColor = Color.FromArgb(255, 255, 192)
        End If
        ФизЛиц = 0
        Button4.BackColor = Color.FromArgb(152, 251, 152)
    End Sub
    Private Sub Refreshgrid()

        tbl.Clear()
        tbl = Selects(StrSql:="SELECT ID, НазвОрг, ТипОбъекта, НазОбъекта, АдресОбъекта FROM ОбъектОбщепита WHERE НазвОрг = '" & Клиент & "' ORDER BY  ТипОбъекта")
        'Dim c As New OleDbCommand
        'c.Connection = conn
        'c.CommandText = StrSql1
        ''Dim ds As New DataSet
        'Dim da As New OleDbDataAdapter(c)
        ''da.Fill(ds, "Сотрудники")
        'da.Fill(tbl)
        Grid1.DataSource = tbl
        Grid1.Columns(0).Visible = False
        Grid1.Columns(1).Visible = False
        'Grid1.Columns(2).Width = 50
        'Grid1.Columns(3).Width = 100
        'Grid1.Columns(0).Width = 60
        'Grid1.Columns(6).Width = 60

        cb = New OleDb.OleDbCommandBuilder(da)
        s = Grid1.Rows.Count - 2
        'NumberAllRows()
    End Sub

    Private Sub ДобОбОбщеп()

        Dim se As Integer = Grid1.Rows.Count - 1
        Dim fg As Integer = ПровЗапПолей()
        If Fh = 1 Then Exit Sub

        Dim MosiFF(Grid1.Columns.Count - 1, Grid1.Rows.Count - 1)
        Dim Str As String = ""

        For Row As Integer = 0 To Grid1.Rows.Count - 1
            For Col As Integer = 0 To Grid1.Columns.Count - 1
                MosiFF(Col, Row) = Grid1.Item(Col, Row).Value
                'Str &= MosiFF(Col, Row) & " "
            Next
            'Str &= vbCrLf
        Next

        Dim colic3 As Integer 'подбираем количество циклов для правильной вставки в таблицу
        Dim colCicl As Integer = Grid1.Rows.Count - 2

        Dim colCicl2 As Integer = Grid1.Rows.Count - 2 - s - 1

        If colCicl2 = 0 Then
            colic3 = 1
        Else
            colic3 = colCicl2
        End If

        colCicl = colCicl - (s2 - s) + colic3

        Dim i As Integer = 0
        Select Case CheckBox6.Checked
            Case False         'сохранение в базу
                For i = 0 To colCicl 'LBound(MosiFF) To UBound(MosiFF)

                    If IsDBNull(MosiFF(2, i)) Then
                        MosiFF(2, i) = ""
                    End If
                    If IsDBNull(MosiFF(3, i)) Then
                        MosiFF(3, i) = ""
                    End If
                    If IsDBNull(MosiFF(4, i)) Then
                        MosiFF(4, i) = ""
                    End If


                    Try

                        Updates(stroka:="UPDATE ОбъектОбщепита  SET ТипОбъекта='" & Trim(MosiFF(2, i)) & "',НазОбъекта='" & Trim(MosiFF(3, i)) & "',АдресОбъекта='" & Trim(MosiFF(4, i)) & "'
            WHERE НазвОрг='" & MosiFF(1, i) & "' And ID =  " & MosiFF(0, i) & "")

                        '            Dim StrSql As String = "UPDATE ОбъектОбщепита  SET ТипОбъекта='" & Trim(MosiFF(2, i)) & "',НазОбъекта='" & Trim(MosiFF(3, i)) & "',АдресОбъекта='" & Trim(MosiFF(4, i)) & "'
                        'WHERE НазвОрг='" & MosiFF(1, i) & "' And ID =  " & MosiFF(0, i) & ""
                        '            Dim c As New OleDbCommand
                        '            c.Connection = conn
                        '            c.CommandText = StrSql

                        '            c.ExecuteNonQuery()
                    Catch ex As Exception
                        Updates(stroka:="INSERT INTO ОбъектОбщепита(НазвОрг,ТипОбъекта,НазОбъекта,АдресОбъекта) VALUES('" & Клиент & "','" & Trim(MosiFF(2, i)) & "','" & Trim(MosiFF(3, i)) & "','" & Trim(MosiFF(4, i)) & "')")

                        'Dim StrSql5 As String = "INSERT INTO ОбъектОбщепита(НазвОрг,ТипОбъекта,НазОбъекта,АдресОбъекта) VALUES('" & Клиент & "','" & Trim(MosiFF(2, i)) & "','" & Trim(MosiFF(3, i)) & "','" & Trim(MosiFF(4, i)) & "')"
                        'Dim c25 As New OleDbCommand
                        'c25.Connection = conn
                        'c25.CommandText = StrSql5
                        'c25.ExecuteNonQuery()
                        ''Next
                    End Try

                Next
            Case True
                For i = 0 To Grid1.Rows.Count - 2
                    If IsDBNull(MosiFF(2, i)) Then
                        MosiFF(2, i) = ""
                    End If
                    If IsDBNull(MosiFF(3, i)) Then
                        MosiFF(3, i) = ""
                    End If
                    If IsDBNull(MosiFF(4, i)) Then
                        MosiFF(4, i) = ""
                    End If
                    Updates(stroka:="INSERT INTO ОбъектОбщепита(НазвОрг,ТипОбъекта,НазОбъекта,АдресОбъекта) VALUES('" & Клиент & "','" & Trim(MosiFF(2, i)) & "','" & Trim(MosiFF(3, i)) & "','" & Trim(MosiFF(4, i)) & "')")
                    'Dim StrSql2 As String = "INSERT INTO ОбъектОбщепита(НазвОрг,ТипОбъекта,НазОбъекта,АдресОбъекта) VALUES('" & Клиент & "','" & Trim(MosiFF(2, i)) & "','" & Trim(MosiFF(3, i)) & "','" & Trim(MosiFF(4, i)) & "')"
                    'Dim c2 As New OleDbCommand
                    'c2.Connection = conn
                    'c2.CommandText = StrSql2
                    'c2.ExecuteNonQuery()
                Next
        End Select




        s = Nothing
        s = Grid1.Rows.Count - 2


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        If MessageBox.Show("Изменить данные?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.No Then Exit Sub




        Dim se As Integer = Grid1.Rows.Count - 1
        Dim fg As Integer = ПровЗапПолей()
        Select Case fg
            Case 1
                Exit Sub
        End Select
        'If s < se Then ВстИДНовОтд()
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
        For i = 0 To Grid1.Rows.Count - 2 'LBound(MosiFF) To UBound(MosiFF)
            If IsDBNull(MosiFF(2, i)) Then
                MosiFF(2, i) = ""
            End If

            If IsDBNull(MosiFF(3, i)) Then
                MosiFF(3, i) = ""
            End If

            If IsDBNull(MosiFF(4, i)) Then
                MosiFF(4, i) = ""
            End If

            Dim dtg As DataTable = Selects(StrSql:="SELECT * FROM ОбъектОбщепита WHERE НазвОрг='" & MosiFF(1, i) & "' And ID =" & MosiFF(0, i) & "")
            If errds = 1 Then
                Updates(stroka:="INSERT INTO ОбъектОбщепита(НазвОрг,ТипОбъекта,НазОбъекта,АдресОбъекта)
VALUES('" & Клиент & "','" & Trim(MosiFF(2, i)) & "','" & Trim(MosiFF(3, i)) & "','" & Trim(MosiFF(4, i)) & "')")
            Else
                Updates(stroka:="UPDATE ОбъектОбщепита  SET ТипОбъекта='" & Trim(MosiFF(2, i)) & "',НазОбъекта='" & Trim(MosiFF(3, i)) & "',АдресОбъекта='" & Trim(MosiFF(4, i)) & "'
            WHERE НазвОрг='" & MosiFF(1, i) & "' And ID =" & MosiFF(0, i) & "")
            End If
        Next

        MsgBox("Данные сохранены!",, Рик)
        Refreshgrid()
    End Sub
    Function ПровЗапПолей() As Integer
        If CheckBox7.Checked = True Then Return 2
        Fh = 0
        For ip As Integer = 0 To Grid1.Rows.Count - 2 'проверяем заполненность поля адрес

            If Grid1.Rows(ip).Cells(4).Value.ToString = "" Then
                If MessageBox.Show("Предлагаем заполнить в таблице адрес объекта общепита - строки " & ip + 1, Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then Fh = 1
            End If
        Next
        Return 2
    End Function

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim bnbc As Integer = MsgBox("Удалить строку?", vbOKCancel, Рик)

        Select Case bnbc
            Case 2
                Refreshgrid()
                Exit Sub
            Case 1

        End Select




        Dim n As String
        Dim m As Integer
        Try
            n = Grid1.CurrentRow.Cells("НазвОрг").Value.ToString
        Catch ex As Exception
            MessageBox.Show("Нельзя удалять единственную строку!", Рик)
            Exit Sub
        End Try

        m = Grid1.CurrentRow.Cells("ID").Value

        Updates(stroka:="DELETE FROM ОбъектОбщепита WHERE НазвОрг ='" & n & "' And ID=" & m & "")
        Refreshgrid()

    End Sub

    Private Sub ДобавлНовКонтр()
        Dim StrSql As String = "SELECT * FROM Клиент"
        Dim conn As New SqlConnection(ConString)
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If
        Dim c As New SqlCommand(StrSql, conn)

        Dim ds As New DataSet
        Dim da As New SqlDataAdapter(c)
        da.Fill(ds, "Контрагент")

        Dim cb As New SqlCommandBuilder(da)
        Dim dsNewRow As DataRow
        dsNewRow = ds.Tables("Контрагент").NewRow()

        dsNewRow.Item("НазвОрг") = Trim(TextBox1.Text)
        dsNewRow.Item("ФормаСобств") = Me.ComboBox2.Text
        dsNewRow.Item("УНП") = Trim(Me.TextBox3.Text)
        dsNewRow.Item("ФИОРуководителя") = Me.TextBox4.Text
        dsNewRow.Item("ДолжнРуководителя") = Me.RichTextBox1.Text
        dsNewRow.Item("ОснованиеДейств") = Me.TextBox6.Text
        dsNewRow.Item("ТелРуков") = Me.TextBox10.Text
        dsNewRow.Item("ФИОДопЛица") = Me.TextBox9.Text
        dsNewRow.Item("ДолжнДопЛица") = Me.TextBox8.Text
        dsNewRow.Item("ТелДопЛица") = Me.TextBox11.Text
        dsNewRow.Item("ЮрАдрес") = Me.TextBox7.Text
        dsNewRow.Item("ФактичАдрес") = ФАдрес
        dsNewRow.Item("ПочтАдрес") = ФПочт
        dsNewRow.Item("КонтТелефон") = Me.TextBox14.Text
        dsNewRow.Item("Факс") = Me.TextBox15.Text
        dsNewRow.Item("ЭлАдрес") = Me.TextBox16.Text
        dsNewRow.Item("ДругиеКонтакты") = Me.TextBox17.Text
        dsNewRow.Item("Банк") = Me.ComboBox3.Text
        dsNewRow.Item("БИКБанка") = Me.TextBox18.Text
        dsNewRow.Item("АдресБанка") = Me.ComboBox4.Text
        dsNewRow.Item("Отделение") = Me.ComboBox4.Text
        dsNewRow.Item("РасчСчетРубли") = Me.TextBox22.Text
        dsNewRow.Item("РасчСчетЕвро") = Me.TextBox23.Text
        dsNewRow.Item("РасчСчетДоллар") = Me.TextBox24.Text
        dsNewRow.Item("РасчСчетРоссРубли") = Me.TextBox25.Text
        dsNewRow.Item("Операционист") = Me.TextBox26.Text
        dsNewRow.Item("КонтТелОпер") = Me.TextBox27.Text
        dsNewRow.Item("ФИОРукРодПадеж") = Me.TextBox28.Text
        dsNewRow.Item("ФИОРукДатПадеж") = Me.TextBox2.Text

        If ЮрЛицо = 1 Then
            dsNewRow.Item("ЮрЛицо") = 1
            dsNewRow.Item("ФизЛицо") = 0
        Else
            dsNewRow.Item("ФизЛицо") = 1
            dsNewRow.Item("ЮрЛицо") = 0
        End If

        dsNewRow.Item("РасчСчетДоллар") = Me.TextBox24.Text
        dsNewRow.Item("РукИП") = CheckBox5.Checked

        ds.Tables("Контрагент").Rows.Add(dsNewRow)

        da.Update(ds, "Контрагент")

        If conn.State = ConnectionState.Open Then
            conn.Close()
        End If

    End Sub
    Private Sub Сохранение()

        s2 = Grid1.Rows.Count - 2
        If CheckBox6.Checked = True And TextBox1.Text = "" Then
            MessageBox.Show("Заполните наименование предпряития нового контрагента", Рик, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        If CheckBox6.Checked = False And ComboBox1.Text = "" Then
            MessageBox.Show("Не выбран контрагент!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If



        If MessageBox.Show("Сохранить данные", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then Exit Sub



        'sw.Start() 'замер времнеи работы процедуры
        If Me.CheckBox1.Checked = False Then
            ФАдрес = Me.TextBox12.Text
        End If
        If Me.CheckBox2.Checked = False Then
            ФПочт = Me.TextBox12.Text
        End If

        If CheckBox6.Checked = True Then
            Клиент = Trim(TextBox1.Text)
            ВстВБазуНовКонтр()
            If Fh = 1 Then Exit Sub
            If CheckBox7.Checked = False Then
                ДобОбОбщеп()
                If Fh = 1 Then Exit Sub
            Else
                БезОбОбщ()
            End If

            refresh2()
            CheckBox6.Checked = False
            MessageBox.Show("Новый контрагент добавлен в базу." & vbCrLf & "Продолжайте работу с выбором контрагента в списке сверху.", Рик, MessageBoxButtons.OK, MessageBoxIcon.Asterisk)

            ОбновлениеСпискаНазванийОрганизаций()
        Else
            Клиент = ComboBox1.Text
            ВносИзмен()
            If Fh = 1 Then Exit Sub
            If CheckBox7.Checked = True Then
                БезОбОбщ()
            End If

            Обнов()
            refresh2()
            ОбновлениеСпискаНазванийОрганизаций()
            MessageBox.Show("Данные сохранены!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Asterisk)



            'Dim StrSql As String = "SELECT НазвОрг FROM Клиент ORDER BY НазвОрг" 'обновляем список клиентов
            'СписокКлиентовОсновной = Selects(StrSql)


        End If

        СозданиепапкиНаСервере(Trim(TextBox1.Text))

        'ДобОбОбщеп()
        'If Fh = 1 Then Exit Sub
        If CheckBox6.Checked = True Then
            Статистика1("Нет", "Создание новой организации", Trim(TextBox1.Text))
        Else
            Статистика1("Нет", "Изменение данных организации", ComboBox1.Text)
        End If

        RichTextBox1.Enabled = True
    End Sub

    Private Sub ОбновлениеСпискаНазванийОрганизаций()
        RunMoving1()
        'Dim results As IEnumerable(Of DataRow) = dtClientAll.AsEnumerable().GroupBy(Function(t) t("НазвОрг")).[Select](Function(g) g.First()) 'выборка из datatable LINQ

        If Not IsDBNull(СписокКлиентовОсновной) Then
            СписокКлиентовОсновной.Clear()
        End If

        For x As Integer = 0 To dtClientAll.Rows.Count - 1
            Dim row2 As DataRow = СписокКлиентовОсновной.NewRow
            row2("НазвОрг") = dtClientAll.Rows(x).Item("НазвОрг").ToString
            СписокКлиентовОсновной.Rows.Add(row2)
        Next

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Сохранение()
    End Sub
    Public Sub БезОбОбщ()
        Updates(stroka:="INSERT INTO ОбъектОбщепита(НазвОрг,АдресОбъекта) VALUES('" & Trim(TextBox1.Text) & "', '" & Trim(МРаботы) & "')")
    End Sub
    Private Sub Обнов()
        For Each TabPage As Control In TabControl1.Controls
            For Each TxtBox As Control In TabPage.Controls
                If TypeName(TxtBox) = "TextBox" Then
                    TxtBox.Text = ""
                End If
            Next

            'For Each TxtBox2 As Control In TabPage.Controls
            '    If TypeName(TxtBox2) = "ComboBox" Then
            '        TxtBox2.Text = ""
            '    End If
            'Next
        Next






        'sw.Start() 'замер времнеи работы процедуры
        Me.TextBox12.Enabled = True
        Me.TextBox13.Enabled = True
        Me.CheckBox1.Checked = False
        Me.CheckBox2.Checked = False


        Dim StrSql As String = "Select * From Клиент Where НазвОрг ='" & Клиент & "'"
        Dim conn As New SqlConnection(ConString)
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If
        Dim c As New SqlCommand(StrSql, conn)

        Dim ds As New DataSet
        Dim da As New SqlDataAdapter(c)
        da.Fill(ds, "КонтРед")
        With Me
            .TextBox1.Text = ds.Tables("КонтРед").Rows(0).Item(0).ToString
            .ComboBox2.Text = ds.Tables("КонтРед").Rows(0).Item(1).ToString
            .TextBox3.Text = ds.Tables("КонтРед").Rows(0).Item(2).ToString
            .TextBox7.Text = ds.Tables("КонтРед").Rows(0).Item(3).ToString
            .TextBox12.Text = ds.Tables("КонтРед").Rows(0).Item(4).ToString
            .TextBox13.Text = ds.Tables("КонтРед").Rows(0).Item(5).ToString
            .TextBox14.Text = ds.Tables("КонтРед").Rows(0).Item(6).ToString
            .TextBox15.Text = ds.Tables("КонтРед").Rows(0).Item(7).ToString
            .TextBox16.Text = ds.Tables("КонтРед").Rows(0).Item(8).ToString
            .TextBox17.Text = ds.Tables("КонтРед").Rows(0).Item(9).ToString
            .ComboBox3.Text = ds.Tables("КонтРед").Rows(0).Item(10).ToString
            .TextBox18.Text = ds.Tables("КонтРед").Rows(0).Item(11).ToString
            .ComboBox4.Text = ds.Tables("КонтРед").Rows(0).Item(12).ToString
            .ComboBox4.Text = ds.Tables("КонтРед").Rows(0).Item(13).ToString
            .TextBox22.Text = ds.Tables("КонтРед").Rows(0).Item(14).ToString
            .TextBox23.Text = ds.Tables("КонтРед").Rows(0).Item(15).ToString
            .TextBox24.Text = ds.Tables("КонтРед").Rows(0).Item(16).ToString
            .TextBox25.Text = ds.Tables("КонтРед").Rows(0).Item(17).ToString
            .RichTextBox1.Text = ds.Tables("КонтРед").Rows(0).Item(18).ToString
            .TextBox4.Text = ds.Tables("КонтРед").Rows(0).Item(19).ToString
            .TextBox6.Text = ds.Tables("КонтРед").Rows(0).Item(20).ToString
            .TextBox10.Text = ds.Tables("КонтРед").Rows(0).Item(21).ToString
            .TextBox8.Text = ds.Tables("КонтРед").Rows(0).Item(22).ToString
            .TextBox9.Text = ds.Tables("КонтРед").Rows(0).Item(23).ToString
            .TextBox11.Text = ds.Tables("КонтРед").Rows(0).Item(24).ToString
            .TextBox26.Text = ds.Tables("КонтРед").Rows(0).Item(25).ToString
            .TextBox27.Text = ds.Tables("КонтРед").Rows(0).Item(26).ToString
            .TextBox28.Text = ds.Tables("КонтРед").Rows(0).Item(29).ToString
            .TextBox2.Text = ds.Tables("КонтРед").Rows(0).Item(30).ToString
            .CheckBox5.Checked = ds.Tables("КонтРед").Rows(0).Item(31).ToString
            CorName = ds.Tables("КонтРед").Rows(0).Item(27)
            If CorName = 1 Then
                Button5.PerformClick()
            Else
                Button4.PerformClick()
            End If

        End With

        Refreshgrid()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        s = Grid1.Rows.Count - 2
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            Me.TextBox13.Enabled = False
            Me.TextBox13.Text = Me.TextBox7.Text
            ФПочт = Me.TextBox13.Text
        ElseIf CheckBox2.Checked = False Then
            Me.TextBox13.Enabled = True
            Me.TextBox13.Text = ""
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            Me.TextBox12.Enabled = False
            Me.TextBox12.Text = Me.TextBox7.Text
            ФАдрес = Me.TextBox12.Text

        ElseIf CheckBox1.Checked = False Then
            Me.TextBox12.Enabled = True
            Me.TextBox12.Text = ""
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If CheckBox6.Checked = True Then
            If ComboBox2.Text = "Индивидуальный предприниматель" Then
                CheckBox5.Checked = False
                CheckBox5.Enabled = False
                RichTextBox1.Text = ComboBox2.Text
                RichTextBox1.Enabled = False
                TextBox6.Text = "Свидетельства о государственной регистрации "
                TextBox4.Text = Trim(TextBox1.Text)
            Else
                CheckBox5.Checked = False
                CheckBox5.Enabled = True
                RichTextBox1.Text = ""
                RichTextBox1.Enabled = True
                TextBox6.Text = ""
                TextBox4.Text = ""
            End If
        End If
    End Sub

    Private Sub Контрагент_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MdiParent = MDIParent1
        Me.WindowState = FormWindowState.Maximized

        'Dim ds As DataTable = SelectsOleDb(StrSql:="SELECT * FROM ФормаСобств")
        'For Each r As DataRow In ds.Rows
        '    Updates(stroka:="INSERT INTO ФормаСобств(ПолноеНазвание,Сокращенное) VALUES('" & r.Item(1).ToString & "','" & r.Item(2).ToString & "')")
        'Next


        'conn = New OleDbConnection
        'conn.ConnectionString = ConString
        'Try
        '    conn.Open()
        'Catch ex As Exception
        '    MessageBox.Show("Не подключен диск U")
        'End Try
        refresh2()
        'TextBox1.Enabled = False

    End Sub
    Private Sub refresh2()


        Dim f As New Thread(Sub() COMxt(Me, "SELECT НазвОрг FROM Клиент ORDER BY НазвОрг", ComboBox1))
        f.IsBackground = True
        f.Start()

        Dim f1 As New Thread(Sub() COMxt(Me, "SELECT ПолноеНазвание FROM ФормаСобств ORDER BY ПолноеНазвание", ComboBox2))
        f1.IsBackground = True
        f1.Start()

        Dim f2 As New Thread(Sub() COMxt(Me, "SELECT DISTINCT КорНазБанк FROM БанкКор ORDER BY КорНазБанк", ComboBox3))
        f2.IsBackground = True
        f2.Start()



        If CheckBox6.Checked = True Then

            Dim dts9 As New DataTable ' - создание объекта таблица данных
            Dim dst9 As New DataSet

            Grid1.DataSource = dts9

            dts9.Columns.Add("ID")
            dts9.Columns.Add("Название Организации")
            dts9.Columns.Add("ТипОбъекта")
            dts9.Columns.Add("НазОбъекта")
            dts9.Columns.Add("АдресОбъекта")
            dts9.Rows.Add(9)
            dst9.Tables.Add(dts9)
            Grid1.Columns(0).Visible = False
            Grid1.Columns(1).Visible = False

        End If

        Button4.PerformClick()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)

        Сохранение()

        Me.Close()
        'Штатное.Show()

        ''sw.Stop()
        ''MessageBox.Show(sw.Elapsed.ToString())
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged



        For Each TabPage As Control In TabControl1.Controls
            For Each TxtBox As Control In TabPage.Controls
                If TypeName(TxtBox) = "TextBox" Then
                    TxtBox.Text = ""
                End If
            Next

            'For Each TxtBox2 As Control In TabPage.Controls
            '    If TypeName(TxtBox2) = "ComboBox" Then
            '        TxtBox2.Text = ""
            '    End If
            'Next
        Next






        'sw.Start() 'замер времнеи работы процедуры
        Me.TextBox12.Enabled = True
        Me.TextBox13.Enabled = True
        Me.CheckBox1.Checked = False
        Me.CheckBox2.Checked = False
        Клиент = Me.ComboBox1.Text

        Dim StrSql As String = "Select * From Клиент Where НазвОрг ='" & Клиент & "'"
        Dim conn As New SqlConnection(ConString)
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If
        Dim c As New SqlCommand(StrSql, conn)

        Dim ds As New DataSet
        Dim da As New SqlDataAdapter(c)
        da.Fill(ds, "КонтРед")
        With Me
            .TextBox1.Text = ds.Tables("КонтРед").Rows(0).Item(0).ToString
            .ComboBox2.Text = ds.Tables("КонтРед").Rows(0).Item(1).ToString
            .TextBox3.Text = ds.Tables("КонтРед").Rows(0).Item(2).ToString
            .TextBox7.Text = ds.Tables("КонтРед").Rows(0).Item(3).ToString
            .TextBox12.Text = ds.Tables("КонтРед").Rows(0).Item(4).ToString
            .TextBox13.Text = ds.Tables("КонтРед").Rows(0).Item(5).ToString
            .TextBox14.Text = ds.Tables("КонтРед").Rows(0).Item(6).ToString
            .TextBox15.Text = ds.Tables("КонтРед").Rows(0).Item(7).ToString
            .TextBox16.Text = ds.Tables("КонтРед").Rows(0).Item(8).ToString
            .TextBox17.Text = ds.Tables("КонтРед").Rows(0).Item(9).ToString
            .ComboBox3.Text = ds.Tables("КонтРед").Rows(0).Item(10).ToString
            .TextBox18.Text = ds.Tables("КонтРед").Rows(0).Item(11).ToString
            .ComboBox4.Text = ds.Tables("КонтРед").Rows(0).Item(12).ToString
            .ComboBox4.Text = ds.Tables("КонтРед").Rows(0).Item(13).ToString
            .TextBox22.Text = ds.Tables("КонтРед").Rows(0).Item(14).ToString
            .TextBox23.Text = ds.Tables("КонтРед").Rows(0).Item(15).ToString
            .TextBox24.Text = ds.Tables("КонтРед").Rows(0).Item(16).ToString
            .TextBox25.Text = ds.Tables("КонтРед").Rows(0).Item(17).ToString
            .RichTextBox1.Text = ds.Tables("КонтРед").Rows(0).Item(18).ToString
            .TextBox4.Text = ds.Tables("КонтРед").Rows(0).Item(19).ToString
            .TextBox6.Text = ds.Tables("КонтРед").Rows(0).Item(20).ToString
            .TextBox10.Text = ds.Tables("КонтРед").Rows(0).Item(21).ToString
            .TextBox8.Text = ds.Tables("КонтРед").Rows(0).Item(22).ToString
            .TextBox9.Text = ds.Tables("КонтРед").Rows(0).Item(23).ToString
            .TextBox11.Text = ds.Tables("КонтРед").Rows(0).Item(24).ToString
            .TextBox26.Text = ds.Tables("КонтРед").Rows(0).Item(25).ToString
            .TextBox27.Text = ds.Tables("КонтРед").Rows(0).Item(26).ToString
            .TextBox28.Text = ds.Tables("КонтРед").Rows(0).Item(29).ToString
            .TextBox2.Text = ds.Tables("КонтРед").Rows(0).Item(30).ToString
            .CheckBox5.Checked = ds.Tables("КонтРед").Rows(0).Item(31).ToString

            CorName = ds.Tables("КонтРед").Rows(0).Item(27)
            If CorName = 1 Then
                Button5.PerformClick()
            Else
                Button4.PerformClick()
            End If

        End With
        'sw.Stop()
        'MessageBox.Show(sw.Elapsed.ToString())
        Refreshgrid()
    End Sub

    Private Sub ВносИзмен()
        'sw.Start() 'замер времнеи работы процедуры

        Updates(stroka:="UPDATE Клиент SET Клиент.ФормаСобств = '" & ComboBox2.Text & "', Клиент.УНП = '" & TextBox3.Text & "',
Клиент.ФактичАдрес = '" & TextBox7.Text & "', Клиент.ЮрАдрес = '" & TextBox12.Text & "', Клиент.ПочтАдрес = '" & TextBox13.Text & "',
Клиент.КонтТелефон = '" & TextBox14.Text & "', Клиент.Факс = '" & TextBox15.Text & "', Клиент.ЭлАдрес = '" & TextBox16.Text & "',
Клиент.ДругиеКонтакты = '" & TextBox17.Text & "', Клиент.Банк = '" & ComboBox3.Text & "', Клиент.БИКБанка = '" & TextBox18.Text & "',
Клиент.АдресБанка = '" & ComboBox4.Text & "', Клиент.Отделение = '" & ComboBox4.Text & "', Клиент.РасчСчетРубли = '" & TextBox22.Text & "',
Клиент.РасчСчетЕвро = '" & TextBox23.Text & "', Клиент.РасчСчетДоллар = '" & TextBox24.Text & "', Клиент.РасчСчетРоссРубли = '" & TextBox25.Text & "',
Клиент.ДолжнРуководителя = '" & RichTextBox1.Text & "', Клиент.ФИОРуководителя = '" & TextBox4.Text & "', Клиент.ОснованиеДейств= '" & TextBox6.Text & "',
Клиент.ТелРуков = '" & TextBox10.Text & "', Клиент.ДолжнДопЛица = '" & TextBox8.Text & "', Клиент.ФИОДопЛица = '" & TextBox9.Text & "',
Клиент.ТелДопЛица = '" & TextBox11.Text & "', Клиент.Операционист = '" & TextBox26.Text & "', Клиент.КонтТелОпер = '" & TextBox27.Text & "',
Клиент.ФизЛицо = '" & ФизЛиц & "', Клиент.ЮрЛицо = '" & ЮрЛицо & "', ФИОРукРодПадеж = '" & Me.TextBox28.Text & "', ФИОРукДатПадеж = '" & Me.TextBox2.Text & "', РукИП = '" & CheckBox5.Checked & "'
WHERE НазвОрг = '" & Клиент & "'")
        ДобОбОбщеп()
    End Sub


    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox2.Focus()
        End If
    End Sub

    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox3.Focus()
        End If
    End Sub

    Private Sub TabPage3_KeyDown(sender As Object, e As KeyEventArgs) Handles TabPage3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox4.Focus()
        End If
    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox28.Focus()
        End If
    End Sub

    Private Sub TextBox28_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox28.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox2.Focus()
        End If
    End Sub

    Private Sub RichTextBox1_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True

            If Not RichTextBox1.Text = "" Then
                Dim sf As String = RichTextBox1.Text
                sf = Strings.Right(sf, Len(sf) - 1)
                sf = UCase(Strings.Left(RichTextBox1.Text, 1)) & sf
                RichTextBox1.Text = Trim(sf)
            End If






            Me.TextBox6.Focus()
        End If
    End Sub

    Private Sub TextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True

            If Not sender.Text = "" Then
                Dim sf As String = sender.Text
                sf = Strings.Right(sf, Len(sf) - 1)
                sf = UCase(Strings.Left(sender.Text, 1)) & sf
                sender.Text = sf
            End If






            Me.TextBox10.Focus()
        End If
    End Sub

    Private Sub TextBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox10.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox9.Focus()
        End If
    End Sub

    Private Sub TextBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox9.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox8.Focus()
        End If
    End Sub

    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox11.Focus()
        End If
    End Sub

    Private Sub TextBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox7.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox12.Focus()
        End If
    End Sub

    Private Sub TextBox11_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox11.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True

            TabControl1.SelectedTab = TabControl1.TabPages(1)
            Me.TextBox7.Focus()
        End If
    End Sub

    Private Sub TextBox12_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox12.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox13.Focus()
        End If
    End Sub

    Private Sub TextBox13_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox13.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox14.Focus()
        End If
    End Sub

    Private Sub TextBox14_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox14.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox16.Focus()
        End If
    End Sub

    Private Sub TextBox16_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox16.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox15.Focus()
        End If
    End Sub

    Private Sub TextBox15_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox15.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox17.Focus()
        End If
    End Sub

    Private Sub TextBox17_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox17.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.Grid1.Focus()
        End If
    End Sub

    Private Sub TextBox19_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox18.Focus()
        End If
    End Sub

    Private Sub TextBox18_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox18.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox4.Focus()
        End If
    End Sub
    Public Sub refr()
        Dim StrSql8 As String = "SELECT Наименование FROM Банк WHERE Наименование LIKE '%" & ComboBox3.Text & "%' ORDER BY Наименование "
        Dim ds As DataTable = Selects(StrSql8)
        Me.ComboBox4.Items.Clear()
        For Each r As DataRow In ds.Rows
            Me.ComboBox4.Items.Add(r(0).ToString)
        Next
    End Sub
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

        ComboBox4.Text = ""
        TextBox18.Text = ""

        Dim ds As DataTable = Selects(StrSql:="SELECT Наименование FROM Банк WHERE Наименование LIKE '%" & ComboBox3.Text & "%' ORDER BY Наименование ")

        Me.ComboBox4.Items.Clear()
        For Each r As DataRow In ds.Rows
            Me.ComboBox4.Items.Add(r(0).ToString)
        Next
    End Sub

    Private Sub TextBox21_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox4.Focus()
        End If
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        For Each TabPage As Control In TabControl1.Controls
            For Each TxtBox As Control In TabPage.Controls
                If TypeName(TxtBox) = "TextBox" Then
                    TxtBox.Text = ""
                End If
            Next

            For Each TxtBox2 As Control In TabPage.Controls
                If TypeName(TxtBox2) = "ComboBox" Then
                    TxtBox2.Text = ""
                End If
            Next
        Next

        Dim dts9 As New DataTable ' - создание объекта таблица данных
        Dim dst9 As New DataSet

        Grid1.DataSource = dts9

        dts9.Columns.Add("ID")
        dts9.Columns.Add("Название Организации")
        dts9.Columns.Add("ТипОбъекта")
        dts9.Columns.Add("НазОбъекта")
        dts9.Columns.Add("АдресОбъекта")
        dts9.Rows.Add(9)
        dst9.Tables.Add(dts9)

        Grid1.Columns(0).Visible = False
        Grid1.Columns(1).Visible = False

        Button4.PerformClick()

    End Sub


    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        Dim ds As DataTable = Selects(StrSql:="Select БИК From Банк Where Наименование='" & ComboBox4.Text & "'")
        Try
            TextBox18.Text = ds.Rows(0).Item(0).ToString
        Catch ex As Exception
            MessageBox.Show("В нашей базе нет данных по вшему запросу", Рик, MessageBoxButtons.OK)
        End Try

    End Sub

    Private Sub TextBox20_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox22.Focus()
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        DelOrg()
    End Sub
    Private Sub DelOrg()

        If ComboBox1.Text = "" Or CheckBox6.Checked = True Then
            MessageBox.Show("Выберите организацию для удаления!", Рик)
            Exit Sub
        End If

        If MessageBox.Show("Вы уверены что хотите удалить организацию " & ComboBox1.Text & " ?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.Cancel Then
            Exit Sub
        End If

        Dim strsql As String = "delete FROM Клиент WHERE НазвОрг='" & ComboBox1.Text & "'"
        Updates(strsql)
        Статистика1("Нет", "Удаление организации", ComboBox1.Text)
        refresh2()
        MessageBox.Show("Организация удалена!", Рик)




    End Sub
    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged

        If CheckBox6.Checked = True Then
            TextBox1.Enabled = True
            ComboBox1.Enabled = False

            For Each TabPage As Control In TabControl1.Controls
                For Each TxtBox As Control In TabPage.Controls
                    If TypeName(TxtBox) = "TextBox" Then
                        TxtBox.Text = ""
                    End If
                Next

                For Each TxtBox2 As Control In TabPage.Controls
                    If TypeName(TxtBox2) = "ComboBox" Then
                        TxtBox2.Text = ""
                    End If
                Next
            Next

            Dim dts9 As New DataTable ' - создание объекта таблица данных
            Dim dst9 As New DataSet

            Grid1.DataSource = dts9

            dts9.Columns.Add("ID")
            dts9.Columns.Add("Название Организации")
            dts9.Columns.Add("ТипОбъекта")
            dts9.Columns.Add("НазОбъекта")
            dts9.Columns.Add("АдресОбъекта")
            dts9.Rows.Add(9)
            dst9.Tables.Add(dts9)

            Grid1.Columns(0).Visible = False
            Grid1.Columns(1).Visible = False




        Else
            TextBox1.Enabled = False
            ComboBox1.Enabled = True
        End If


    End Sub



    Private Sub TextBox22_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox22.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox25.Focus()
        End If
    End Sub
    Function корназвОрг() As String
        Dim strsql As String
        strsql = "SELECT Сокращенное FROM ФормаСобств WHERE ПолноеНазвание='" & ComboBox2.Text & "'"
        Try
            ds.Clear()
        Catch ex As Exception

        End Try
        ds = Selects(strsql)

        Return ds.Rows(0).Item(0).ToString
    End Function
    Private Sub CheckBox7_CheckedChanged_1(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        If CheckBox7.Checked = True Then
            Grid1.Visible = False
            GroupBox1.Enabled = False

            If ComboBox2.Text = "Индивидуальный предприниматель" Then
                МРаботы = корназвОрг() & " " & Trim(TextBox1.Text)
            Else
                МРаботы = корназвОрг() & " """ & Trim(TextBox1.Text) & """ "
            End If
            'НовОргДопФорма.ShowDialog()
        Else
            Grid1.Visible = True
            GroupBox1.Enabled = True
            МРаботы = ""
        End If

    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        TextBox28.Text = TextBox4.Text
        TextBox2.Text = TextBox4.Text
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Банк.ShowDialog()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        ИзменБанк.ShowDialog()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        БанкОтдДобав.ShowDialog()
    End Sub

    Private Sub TextBox25_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox25.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox23.Focus()
        End If
    End Sub

    Private Sub TextBox23_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox23.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox24.Focus()
        End If
    End Sub

    Private Sub TextBox24_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox24.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox26.Focus()
        End If
    End Sub

    Private Sub TextBox26_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox26.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox27.Focus()
        End If
    End Sub

    Private Sub TextBox27_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox27.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.Button2.Focus()
        End If
    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox4.Focus()
        End If
    End Sub


    Private Sub RichTextBox1_LostFocus(sender As Object, e As EventArgs)
        If Not RichTextBox1.Text = "" Then
            Dim sf As String = RichTextBox1.Text
            sf = Strings.Right(sf, Len(sf) - 1)
            sf = UCase(Strings.Left(RichTextBox1.Text, 1)) & sf
            RichTextBox1.Text = sf
        End If
    End Sub
End Class