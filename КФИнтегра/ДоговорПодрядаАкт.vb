Option Explicit On
Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.Threading

Public Class ДоговорПодрядаАкт
    Dim КодАкт As Integer
    Dim ПослНомАкт As Integer
    Dim strsql, strsql1, ПослНомерДоговора As String
    Dim ds, ds1, ds2 As DataTable
    Dim file2() As String
    Dim FilesList() As String
    Dim СохрЗак As String
    Dim arrtbox As New Dictionary(Of String, String)
    Dim arrcombox As New Dictionary(Of String, String)
    Dim arrlabel As New Dictionary(Of String, String)
    Private Delegate Sub comb19()
    Private Delegate Sub listbx1()
    Private Sub arrbx()
        If arrtbox.Any Then
            arrtbox.Clear()
        End If

        If arrlabel.Any Then
            arrlabel.Clear()
        End If

        If arrcombox.Any Then
            arrcombox.Clear()
        End If

        Dim Ctrl As Control
        Dim Ctrl1 As Control
        Dim Ctrl2 As Control

        For Each Ctrl In Me.Controls 'перебираем текстбоксы вне tabcontrol и groupbox
            If TypeName(Ctrl) = "TextBox" Then
                arrtbox.Add(Ctrl.Name, Ctrl.Text)
                'Ctrl.Value = "бла-бла-бла"
            End If
        Next

        For Each Ctrl1 In Me.Controls 'перебираем combobox вне tabcontrol и groupbox
            If TypeName(Ctrl1) = "ComboBox" Then
                arrcombox.Add(Ctrl1.Name, Ctrl1.Text)
                'Ctrl.Value = "бла-бла-бла"
            End If
        Next

        For Each Ctrl2 In Me.Controls 'перебираем label вне tabcontrol и groupbox
            If TypeName(Ctrl2) = "label" Then
                arrlabel.Add(Ctrl2.Name, Ctrl2.Text)
                'Ctrl.Value = "бла-бла-бла"
            End If
        Next


        For Each gh In Me.Controls.OfType(Of GroupBox) 'перебираем combobox вне tabcontrol но в groupbox

            For Each tx In gh.Controls.OfType(Of ComboBox)
                arrcombox.Add(tx.Name, tx.Text)
            Next

            For Each ts In gh.Controls.OfType(Of Label)
                arrlabel.Add(ts.Name, ts.Text)
            Next
            For Each tf In gh.Controls.OfType(Of TextBox)
                arrtbox.Add(tf.Name, tf.Text)
            Next
        Next











    End Sub

    Private Sub ComboBox19_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox19.SelectedIndexChanged
        Label96.Text = ComboBox2.Items.Item(ComboBox19.SelectedIndex)
        arrbx()
        очист()

        ТолькоДоговора(CType(Label96.Text, Integer))
        Com1Sel9()

    End Sub

    Private Sub Com1Sel()
        Dim list As New Dictionary(Of String, Object)
        list.Add("@НазвОрганиз", ComboBox1.SelectedItem)

        If Not IsDBNull(ds) And ds IsNot Nothing Then 'проверка datatable
            ds.Clear()
        End If

        ds = Selects(StrSql:="SELECT DISTINCT Сотрудники.ФИОСборное, ДогПодряда.ID
FROM Сотрудники INNER JOIN ДогПодряда ON Сотрудники.КодСотрудники = ДогПодряда.ID
WHERE Сотрудники.НазвОрганиз=@НазвОрганиз ORDER BY Сотрудники.ФИОСборное", list)

        Label96.Text = "N"
        ComboBox2.Items.Clear()
        Me.ComboBox19.AutoCompleteCustomSource.Clear()
        Me.ComboBox19.Items.Clear()
        For Each r As DataRow In ds.Rows
            Me.ComboBox19.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox19.Items.Add(r(0).ToString)
            Me.ComboBox2.Items.Add(r(1).ToString)
        Next

    End Sub
    Private Sub refreshList2()
        Dim id As Integer = CType(arrlabel("Label96"), Integer)
        'Dim ds = dtSotrudnikiAll.Select("КодСотрудники=" & id & "")
        'Dim StrSql2 As String = "Select Фамилия From Сотрудники Where КодСотрудники=" & CType(arrlabel("Label96"), Integer) & ""
        'Dim ds As DataTable = Selects(StrSql2)

        Dim list = From x In (From c In dtPutiDokumentovAll.AsEnumerable Where Not IsDBNull(c.Item("IDСотрудник")) Select c) Where x.Item("IDСотрудник") = id And x.Item("ДокМесто").ToString.Contains("Акт договор подряда") Select x

        'From c In dtPutiDokumentovAll.AsEnumerable Where Not IsDBNull(c.Item("IDСотрудник")) Select y




        'FilesList = Nothing
        'file2 = Nothing
        'Dim gth4 As String
        'Try
        '    FilesList = IO.Directory.GetFiles(OnePath & arrcombox("ComboBox1"), "*" & ds(0).Item("Фамилия") & "*.doc*", IO.SearchOption.AllDirectories)
        '    file2 = IO.Directory.GetFiles(OnePath & arrcombox("ComboBox1"), "*" & ds(0).Item("Фамилия") & "*.doc*", IO.SearchOption.AllDirectories)

        '    For n As Integer = 0 To FilesList.Length - 1
        '        gth4 = ""
        '        gth4 = IO.Path.GetFileName(file2(n))
        '        file2(n) = gth4
        '        'TextBox44.Text &= gth + vbCrLf
        '    Next
        'Catch ex As Exception

        'End Try

        'If FilesList.Length = 0 Then Exit Sub

        'If ListBox1.InvokeRequired Then
        '    Me.Invoke(New listbx1(AddressOf refreshList2))
        'Else
        ListBox1.Items.Clear()
        ComboBox4.Items.Clear()

        For Each f In list ' Распечатываем весь получившийся массив
            ListBox1.Items.Add(f.Item("ИмяФайла").ToString) ' На ListBox2
            ComboBox4.Items.Add(f.Item("ПолныйПуть"))
        Next
        'End If




        'ListBox2.Items.Add(Files2)



    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            MaskedTextBox3.Text = MaskedTextBox2.Text
            MaskedTextBox6.Text = MaskedTextBox2.Text
        Else
            MaskedTextBox3.Text = Now.Date
            MaskedTextBox6.Text = Now.Date
        End If
    End Sub
    Private Sub очист()
        TextBox4.Text = ""
        TextBox3.Text = ""
        TextBox2.Text = ""
        TextBox1.Text = ""

        MaskedTextBox1.Text = ""
        MaskedTextBox2.Text = ""
        MaskedTextBox3.Text = ""
        MaskedTextBox6.Text = ""

        ListBox1.Items.Clear()
        Try
            Grid1.Rows.Clear()
        Catch ex As Exception
            For i As Integer = 0 To Grid1.ColumnCount - 1
                Grid1.Columns.RemoveAt(0)
            Next
        End Try

    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        очист()
        Com1Sel()
        ComboBox19.Text = String.Empty
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            MaskedTextBox1.Enabled = True
        Else
            MaskedTextBox1.Enabled = False
        End If
    End Sub

    Private Sub MaskedTextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            If CheckBox2.Checked = True Then
                MaskedTextBox3.Text = MaskedTextBox2.Text
                MaskedTextBox6.Text = MaskedTextBox2.Text
            End If


            Me.MaskedTextBox3.Focus()
            ВычДатВыплат(MaskedTextBox2.Text)
        End If
    End Sub

    Private Sub MaskedTextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.MaskedTextBox6.Focus()

        End If
    End Sub

    Private Sub MaskedTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.MaskedTextBox6.Focus()

        End If
    End Sub

    Private Sub MaskedTextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            ВычДатВыплат(CDate(MaskedTextBox6.Text))
            If TextBox1.Enabled = False Then
                TextBox3.Focus()
            Else
                TextBox1.Focus()
            End If

        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            TextBox2.Enabled = True
        Else
            TextBox2.Enabled = False
        End If
    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            TextBox1.Text = ""
            Dim f As Double
            f = Replace(TextBox3.Text, ".", ",")
            TextBox3.SelectionStart = TextBox3.Text.Length '14

            If СправкаПоЗарплате.bool(f) = True Then
                TextBox3.Text = f & ",00"
            Else
                TextBox3.Text = f
                If СправкаПоЗарплате.Count(f) = 1 Then
                    TextBox3.Text = f & "0"
                End If
            End If
            Button1.Focus()
            If TextBox2.Text <> "" Then
                txt1(TextBox3.Text)
            End If
        End If

    End Sub
    Private Sub txt1(ByVal txt As String)
        Dim f As Double = CDbl(txt)
        f = Math.Round(f / CDbl(TextBox2.Text), 2)
        f = Replace(f, ".", ",")
        If СправкаПоЗарплате.bool(f) = True Then
            TextBox1.Text = f & ",00"
        Else
            TextBox1.Text = f
            If СправкаПоЗарплате.Count(f) = 1 Then
                TextBox1.Text = f & "0"
            End If
        End If

    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown

        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Dim f As Double
            f = Replace(TextBox2.Text, ".", ",")
            TextBox2.SelectionStart = TextBox2.Text.Length '14

            If СправкаПоЗарплате.bool(f) = True Then
                TextBox2.Text = f & ",00"
            Else
                TextBox2.Text = f
                If СправкаПоЗарплате.Count(f) = 1 Then
                    TextBox2.Text = f & "0"
                End If
            End If
            If TextBox3.Text <> "" Then
                TextBox1.Text = ""
                Dim f1 As Double
                Dim f2 As Double = CDbl(TextBox3.Text)
                f1 = f2 / f
                f1 = Math.Round(f1, 2)
                If СправкаПоЗарплате.bool(f1) = True Then
                    TextBox1.Text = f1 & ",00"
                Else
                    TextBox1.Text = f1
                    If СправкаПоЗарплате.Count(f1) = 1 Then
                        TextBox1.Text = f1 & "0"
                    End If
                End If
            End If
            TextBox3.Focus()

        End If
    End Sub
    Private Sub grid1all()
        Dim ds4 As DataTable

        '        Dim strsql2 As String = "SELECT ДогПодряда.Код FROM ДогПодряда INNER JOIN ДогПодрядаАкт ON ДогПодряда.Код = ДогПодрядаАкт.IDДогПодр
        'WHERE ДогПодряда.ID=" & CType(arrlabel("Label96"), Integer) & ""
        '        Dim dcv As DataTable = Selects(strsql2)

        Dim list As New Dictionary(Of String, Object)
        list.Add("@ID", CType(arrlabel("Label96"), Integer))

        ds4 = Selects(StrSql:= "SELECT ДогПодрядаАкт.Код, ДогПодрядаАкт.IDДогПодр, ДогПодряда.НомерДогПодр as [Номер договора подряда], ДогПодрядаАкт.ПорНомерАкта as [Номер акта], ДогПодрядаАкт.ЗаПериодС as [C],
ДогПодрядаАкт.ЗаПериодПо as [ПО], ДогПодрядаАкт.ВремяРабот as [Время работ], ДогПодрядаАкт.СтоимЧаса as [Стоим часа], ДогПодрядаАкт.СтоимРабот as [Стоим работ],
ДогПодрядаАкт.ДатаАкта as [Дата акта], ДогПодрядаАкт.ДатаОплатыРабот as [Дата оплаты]
FROM ДогПодряда INNER JOIN ДогПодрядаАкт ON ДогПодряда.Код = ДогПодрядаАкт.IDДогПодр
WHERE ДогПодряда.ID=@ID", list)



        Grid1.DataSource = ds4
            Grid1.Columns(0).Visible = False
            Grid1.Columns(1).Visible = False
            Grid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            Grid1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            Grid1.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter



    End Sub
    Private Sub ТолькоДоговора(ByVal int As Integer)

        Dim ds1
        Using dbcx As New DbAllDataContext
            ds1 = (From x In dbcx.ДогПодряда.AsEnumerable
                   Where x.ID = int
                   Select x.НомерДогПодр Distinct).ToList
        End Using
        'Dim ds1 = From x In dtDogovorPadriadaAll Where x.Item("ID") = int Select x.Item("НомерДогПодр") Distinct




        '        Dim strsql1 As String = "SELECT DISTINCT ДогПодряда.НомерДогПодр 
        'FROM ДогПодряда WHERE ДогПодряда.ID=" & int & " ORDER BY НомерДогПодр"
        '        Dim ds1 As DataTable = Selects(strsql1)

        'ComboBox3.Items.Clear()
        'For Each r In ds1
        '    Me.ComboBox3.Items.Add(r.ToString)
        'Next

        ComboBox3.DataSource = ds1

    End Sub
    Private Sub Com1Sel9()

        'Dim rf As New Thread(AddressOf refreshList2) 'новый поток для refreshList2
        'rf.IsBackground = True




        Dim list As New Dictionary(Of String, Object)
        list.Add("@ID", CType(Label96.Text, Integer))

        ds1 = Selects(StrSql:="SELECT ДогПодряда.ID, ДогПодрядаАкт.ПорНомерАкта, ДогПодрядаАкт.ЗаПериодС,
ДогПодрядаАкт.ЗаПериодПо, ДогПодрядаАкт.ВремяРабот, ДогПодрядаАкт.СтоимЧаса, ДогПодрядаАкт.СтоимРабот, ДогПодрядаАкт.Код, ДогПодрядаАкт.ДатаАкта,
ДогПодрядаАкт.ДатаОплатыРабот 
FROM ДогПодряда INNER JOIN ДогПодрядаАкт ON ДогПодряда.Код = ДогПодрядаАкт.IDДогПодр
WHERE ДогПодряда.ID=@ID ORDER BY ДогПодрядаАкт.ПорНомерАкта DESC", list)

        If errds = 1 Then
            КодАкт = Nothing
            refreshList2()
            ПослНомАкт = 0
            Exit Sub
        Else
            ЗаполнДанн(ds1)
            refreshList2()
            grid1all()

        End If

    End Sub
    Private Function Проверка()
        If ComboBox1.Text = "" Or ComboBox19.Text = "" Then
            MessageBox.Show("Выберите организацию и сотрудника!", Рик)
            Return 1
        End If

        If TextBox4.Text = "" Then
            MessageBox.Show("Заполните номер акта!", Рик)
            Return 1
        End If

        If MaskedTextBox2.MaskCompleted = False Or MaskedTextBox3.MaskCompleted = False Then
            MessageBox.Show("Заполните раздел 'период'!", Рик)
            Return 1
        End If

        If MaskedTextBox6.MaskCompleted = False Then
            MessageBox.Show("Заполните дату акта!", Рик)
            Return 1
        End If

        If MaskedTextBox1.MaskCompleted = False Then
            MessageBox.Show("Выберите дату оплаты работ!", Рик)
            Return 1
        End If

        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Then
            MessageBox.Show("Заполните раздел 'Отработанное время и начисленная сумма'!", Рик)
            Return 1
        End If

        If ComboBox3.Text = "" Then
            MessageBox.Show("Выберите номер договора!", Рик)
            Return 1
        End If



        Return 0
    End Function
    Private Function ПровДубл()
        Dim list As New Dictionary(Of String, Object)
        list.Add("@ПорНомерАкта", CType(TextBox4.Text, Integer))
        list.Add("@НомерДогПодр", ComboBox3.Text)

        Dim df = Selects(StrSql:= "SELECT ДогПодрядаАкт.ПорНомерАкта
FROM ДогПодряда INNER JOIN ДогПодрядаАкт ON ДогПодряда.Код = ДогПодрядаАкт.IDДогПодр
WHERE ПорНомерАкта=@ПорНомерАкта AND ДогПодряда.НомерДогПодр=@НомерДогПодр", list)
        If errds = 0 Then
            If MessageBox.Show("Заменить старые данные акта №" & CType(TextBox4.Text, Integer) & " -новыми?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                Return 1
            End If
        End If
        Return 0
    End Function
    Private Sub НовСтрока(ByVal strsql2 As String, ByVal код As Integer)

        Dim c As New OleDbCommand
        c.Connection = conn
        c.CommandText = strsql2
        Dim ds As New DataSet
        Dim da As New OleDbDataAdapter(c)
        da.Fill(ds, "Сохранение")

        Dim cb As New OleDbCommandBuilder(da)
        Dim dsNewRow As DataRow

        dsNewRow = ds.Tables("Сохранение").NewRow()
        dsNewRow.Item("IDДогПодр") = код
        dsNewRow.Item("ПорНомерАкта") = TextBox4.Text
        dsNewRow.Item("ЗаПериодС") = MaskedTextBox2.Text
        dsNewRow.Item("ЗаПериодПо") = MaskedTextBox3.Text
        dsNewRow.Item("ВремяРабот") = TextBox1.Text
        dsNewRow.Item("СтоимЧаса") = TextBox2.Text
        dsNewRow.Item("СтоимРабот") = TextBox3.Text
        dsNewRow.Item("ДатаАкта") = MaskedTextBox6.Text
        dsNewRow.Item("ДатаОплатыРабот") = MaskedTextBox1.Text
        ds.Tables("Сохранение").Rows.Add(dsNewRow)

        'ds.Tables("Сохранение").Rows(0).Item(0) = a
        'ds.Tables("Сохранение").Rows(0).Item(1) = Me.TextBox1.Text
        da.Update(ds, "Сохранение")

        Updates(stroka:="INSERT INTO ")


    End Sub
    Private Sub СохрВБазу()
        Dim strsql As String


        '        strsql = "SELECT ДогПодрядаАкт.ПорНомерАкта
        'FROM ДогПодряда INNER JOIN ДогПодрядаАкт ON ДогПодряда.Код = ДогПодрядаАкт.IDДогПодр
        'WHERE ДогПодряда.ID=" & CType(Label96.Text, Integer) & ""
        '        ds3 = Selects(strsql)



        Dim ds3 = dtDogovorPadriadaAll.Select("ID=" & CType(Label96.Text, Integer) & " AND НомерДогПодр='" & ComboBox3.Text & "'")
        Dim код As Integer = ds3(0).Item("Код")


        'strsql = "INSERT INTO ДогПодрядаАкт(IDДогПодр, ПорНомерАкта, ЗаПериодС, ЗаПериодПо, ВремяРабот, СтоимЧаса, СтоимРабот)
        'VALUES(" & код & ",'" & TextBox4.Text & "','" & MaskedTextBox2.Text & "','" & MaskedTextBox3.Text & "',
        ''" & TextBox1.Text & "','" & TextBox2.Text & "', '" & TextBox3.Text & "')"
        strsql = "SELECT ДогПодрядаАкт.IDДогПодр, ДогПодрядаАкт.ПорНомерАкта, ДогПодрядаАкт.ЗаПериодС,
ДогПодрядаАкт.ЗаПериодПо, ДогПодрядаАкт.ВремяРабот, ДогПодрядаАкт.СтоимЧаса, ДогПодрядаАкт.СтоимРабот, ДогПодрядаАкт.ДатаАкта, ДогПодрядаАкт.ДатаОплатыРабот
FROM ДогПодрядаАкт"
        Dim list As New Dictionary(Of String, Object)
        list.Add("@IDДогПодр", код)
        list.Add("@ПорНомерАкта", TextBox4.Text)
        list.Add("@ЗаПериодС", MaskedTextBox2.Text)
        list.Add("@ЗаПериодПо", MaskedTextBox3.Text)
        list.Add("@ВремяРабот", TextBox1.Text)
        list.Add("@СтоимЧаса", TextBox2.Text)
        list.Add("@СтоимРабот", TextBox3.Text)
        list.Add("@ДатаАкта", MaskedTextBox6.Text)
        list.Add("@ДатаОплатыРабот", MaskedTextBox1.Text)


        Updates(stroka:="INSERT INTO ДогПодрядаАкт(IDДогПодр, ПорНомерАкта,
ЗаПериодС,ЗаПериодПо, ВремяРабот, СтоимЧаса, СтоимРабот, ДатаАкта, ДатаОплатыРабот)
VALUES(@IDДогПодр,@ПорНомерАкта,@ЗаПериодС,@ЗаПериодПо,@ВремяРабот,@СтоимЧаса,@СтоимРабот,@ДатаАкта,@ДатаОплатыРабот)", list, "ДогПодрядаАкт")

        'НовСтрока(strsql, код)



    End Sub
    Private Sub Обновление()
        Dim list As New Dictionary(Of String, Object)
        list.Add("@Код", КодАкт)

        Updates(stroka:="UPDATE ДогПодрядаАкт SET ПорНомерАкта=" & CType(TextBox4.Text, Integer) & ",ЗаПериодС='" & MaskedTextBox2.Text & "',ЗаПериодПо='" & MaskedTextBox3.Text & "',
ВремяРабот='" & TextBox1.Text & "',СтоимЧаса='" & TextBox2.Text & "', СтоимРабот='" & TextBox3.Text & "', ДатаАкта='" & MaskedTextBox6.Text & "', ДатаОплатыРабот='" & MaskedTextBox1.Text & "'
        WHERE ДогПодрядаАкт.Код=@Код", list, "ДогПодрядаАкт")

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Проверка() = 1 Then
            Exit Sub
        End If

        'Dim rf As New Thread(AddressOf refreshList2) 'новый поток для refreshList2
        'rf.IsBackground = True


        'refreshList2()

        'grid1all()


        'Dim gr As New Thread(AddressOf grid1all) 'новый поток для grid1all
        'gr.IsBackground = True

        If ПослНомАкт = CType(TextBox4.Text, Integer) And ПослНомерДоговора = ComboBox3.Text Then
            If ПровДубл() = 1 Then Exit Sub
            Обновление()
        Else
            СохрВБазу()
        End If
        'rf.Start()
        'gr.Start()

        If CheckBox4.Checked = False Then
            Доки()
            очист()
            Com1Sel9()
        Else
            очист()
            Com1Sel9()
            MessageBox.Show("Данные удачно внесены в базу!", Рик)
        End If
        dtPutiDokumentov()
    End Sub
    Private Function ДогПодНом() As DataRow()

        Dim df = dtDogovorPadriadaAll.Select("ID=" & CType(Label96.Text, Integer) & " AND НомерДогПодр='" & ComboBox3.Text & "'")
        'Dim strsql As String = "SELECT ДатаДогПодр FROM ДогПодряда WHERE ID=" & CType(Label96.Text, Integer) & " AND НомерДогПодр='" & ComboBox3.Text & "'"

        'Dim df As DataTable = Selects(strsql)
        Return df
    End Function
    Private Sub Доки()
        'Dim ДолжСОконч, СтавкаНов, СклонГод, СрКонтПроп, ПоСовмИлиОсн, ПоСовмПриказ As String
        'Dim oWord As Word.Application
        'Dim oDoc As Word.Document

        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document
        Me.Cursor = Cursors.WaitCursor

        oWord = CreateObject("Word.Application")
        oWord.Visible = False

        'Try
        '    IO.File.Copy(OnePath & "\ОБЩДОКИ\General\ActPodriada.doc", "C:\Users\Public\Documents\Рик\ActPodriada.doc")
        'Catch ex As Exception
        '    'If "Zayavlenie.doc" <> "" Then IO.File.Delete("C:\Users\Public\Documents\Рик\Zayavlenie.doc")
        '    If Not IO.Directory.Exists("c:\Users\Public\Documents\Рик") Then
        '        IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рик")
        '    End If
        '    IO.File.Copy(OnePath & "\ОБЩДОКИ\General\ActPodriada.doc", "C:\Users\Public\Documents\Рик\ActPodriada.doc")
        'End Try

        Начало("ActPodriada.doc")
        oWordDoc = oWord.Documents.Add(firthtPath & "\ActPodriada.doc")

        Dim mObj As Object = Подоходный(TextBox3.Text)
        Dim Организация = Org(ComboBox1.Text)
        Dim Сотрудник = Sotrudnic(CType(Label96.Text, Integer))
        Dim ДП As DataRow() = ДогПодНом()
        'MsgBox(ДолжСОконч)
        With oWordDoc.Bookmarks
            .Item("АктПодр1").Range.Text = TextBox4.Text
            .Item("АктПодр2").Range.Text = ДП(0).Item("ДатаДогПодр").ToString
            .Item("АктПодр3").Range.Text = ComboBox3.Text & " - " & Strings.Right(ДП(0).Item("ДатаДогПодр").ToString, 4)
            .Item("АктПодр4").Range.Text = MaskedTextBox2.Text
            .Item("АктПодр5").Range.Text = MaskedTextBox3.Text
            If Организация(0).Item(1).ToString = "Индивидуальный предприниматель" Then
                .Item("АктПодр6").Range.Text = Организация(0).Item(1).ToString & " " & Организация(0).Item(0).ToString
                .Item("АктПодр21").Range.Text = Организация(0).Item(1).ToString & " " & Организация(0).Item(0).ToString
            Else
                .Item("АктПодр6").Range.Text = Организация(0).Item(1).ToString & " «" & Организация(0).Item(0).ToString & "» "
                .Item("АктПодр21").Range.Text = Организация(0).Item(1).ToString & " «" & Организация(0).Item(0).ToString & "» "
            End If
            If Организация(0).Item(0).ToString = "Итал Гэлэри Плюс" Then
                .Item("АктПодр7").Range.Text = ДобОконч(Организация(0).Item(18).ToString) & " " & Организация(0).Item(29).ToString
                .Item("АктПодр8").Range.Text = ""
            ElseIf Организация(0).Item(1).ToString = "Индивидуальный предприниматель" Then
                .Item("АктПодр7").Range.Text = Организация(0).Item(29).ToString
                .Item("АктПодр8").Range.Text = "действующего на основании " & Организация(0).Item(20).ToString
            Else
                .Item("АктПодр7").Range.Text = ДобОконч(Организация(0).Item(18).ToString) & " " & Организация(0).Item(29).ToString
                .Item("АктПодр8").Range.Text = "действующего на основании " & Организация(0).Item(20).ToString
            End If
            .Item("АктПодр9").Range.Text = ComboBox19.Text
            .Item("АктПодр10").Range.Text = "N " & ComboBox3.Text & " - " & Strings.Right(ДП(0).Item("ДатаДогПодр").ToString, 4)
            .Item("АктПодр11").Range.Text = ДП(0).Item("ДатаДогПодр").ToString
            .Item("АктПодр12").Range.Text = ComboBox3.Text & " - " & Strings.Right(ДП(0).Item("ДатаДогПодр").ToString, 4) & " от " & ДП(0).Item("ДатаДогПодр").ToString
            .Item("АктПодр13").Range.Text = TextBox1.Text
            .Item("АктПодр14").Range.Text = TextBox2.Text
            .Item("АктПодр15").Range.Text = TextBox3.Text
            .Item("АктПодр16").Range.Text = TextBox3.Text
            .Item("АктПодр17").Range.Text = ЧислоПрописДляСправки(TextBox3.Text)
            .Item("АктПодр18").Range.Text = mObj(0) & " руб."
            .Item("АктПодр19").Range.Text = mObj(1) & " руб."
            .Item("АктПодр20").Range.Text = MaskedTextBox1.Text
            .Item("АктПодр22").Range.Text = Организация(0).Item(4).ToString
            .Item("АктПодр23").Range.Text = Организация(0).Item(2).ToString
            .Item("АктПодр24").Range.Text = Организация(0).Item(14).ToString
            .Item("АктПодр25").Range.Text = Организация(0).Item(12).ToString
            .Item("АктПодр26").Range.Text = Организация(0).Item(11).ToString
            If Организация(0).Item(31) = True And Not Организация(0).Item(1).ToString = "Индивидуальный предприниматель" Then
                .Item("АктПодр27").Range.Text = ФИОКорРук(Организация(0).Item(19).ToString, True)
            ElseIf Организация(0).Item(31) = True And Организация(0).Item(1).ToString = "Индивидуальный предприниматель" Then
                .Item("АктПодр27").Range.Text = ФИОКорРук(Организация(0).Item(19).ToString, True)
            ElseIf Организация(0).Item(31) = False And Организация(0).Item(1).ToString = "Индивидуальный предприниматель" Then
                .Item("АктПодр27").Range.Text = ФИОКорРук(Организация(0).Item(19).ToString, True)
            Else
                .Item("АктПодр27").Range.Text = ФИОКорРук(Организация(0).Item(19).ToString, False)
            End If

            .Item("АктПодр28").Range.Text = ComboBox19.Text
            .Item("АктПодр29").Range.Text = Сотрудник(0).Item(10).ToString & Сотрудник(0).Item(11).ToString
            .Item("АктПодр30").Range.Text = Сотрудник(0).Item(12).ToString
            .Item("АктПодр31").Range.Text = Сотрудник(0).Item(14).ToString
            .Item("АктПодр32").Range.Text = Сотрудник(0).Item(16).ToString
            .Item("АктПодр33").Range.Text = Сотрудник(0).Item(15).ToString
            .Item("АктПодр34").Range.Text = ФИОКорРук(ComboBox19.Text, False)
            .Item("АктПодр35").Range.Text = MaskedTextBox6.Text



        End With


        Dim Name As String = TextBox4.Text & " " & ФИОКорРук(ComboBox19.Text, False) & " от " & MaskedTextBox6.Text & " (Акт договор подряда № " & ComboBox3.Text & ").doc"
        Dim СохрЗак2 As New List(Of String)
        СохрЗак2.AddRange(New String() {ComboBox1.Text & "\Договор подряда\" & Now.Year & "\", Name})
        oWordDoc.SaveAs2(PathVremyanka & Name,,,,,, False)
        oWordDoc.Close(True)
        oWord.Quit(True)
        Конец(ComboBox1.Text & "\Договор подряда\" & Now.Year, Name, CType(Label96.Text, Integer), ComboBox1.Text, "\ActPodriada.doc", "Акт договор подряда")
        Dim massFTP3 As New ArrayList
        massFTP3.Add(СохрЗак2)
        Parallel.Invoke(Sub() RunMoving4())






        'If Not IO.Directory.Exists(OnePath & ComboBox1.Text & "\Договор подряда\" & Now.Year) Then
        '    IO.Directory.CreateDirectory(OnePath & ComboBox1.Text & "\Договор подряда\" & Now.Year)
        'End If

        'Dim d As String = "C:\Users\Public\Documents\Рик\" & TextBox4.Text & " " & ФИОКорРук(ComboBox19.Text, False) & " от " & MaskedTextBox6.Text & " (Акт договор подряда № " & ComboBox3.Text & ").doc"
        'Dim dnew As String = OnePath & ComboBox1.Text & "\Договор подряда\" & Now.Year & "\" & TextBox4.Text & " " & ФИОКорРук(ComboBox19.Text, False) & " от " & MaskedTextBox6.Text & " (Акт договор подряда № " & ComboBox3.Text & ").doc"

        'oWordDoc.SaveAs2(d,,,,,, False)

        ''oWordDoc.SaveAs2("U: \Офис\Финансовый\6. Бух.услуги\Кадры\" & Клиент & "\Заявление\" & Год & "\" & Заявление(9) & " (заявление)" & ".doc",,,,,, False)
        'Try
        '    IO.File.Copy(d, dnew)
        'Catch ex As Exception
        '    If MessageBox.Show("Акт договора подряда с сотрудником " & ФИОКорРук(ComboBox19.Text, False) & " существует. Заменить старый документ новым?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then
        '        Try
        '            IO.File.Delete(dnew)
        '        Catch ex1 As Exception
        '            MessageBox.Show("Закройте файл!", Рик)
        '        End Try


        '        IO.File.Copy(d, dnew)
        '    End If
        'End Try
        'СохрЗак = dnew

        'oWordDoc.Close(True)
        'oWord.Quit(True)

        'Dim mass() As String
        If MessageBox.Show("Акт договора подряда №" & ComboBox3.Text & " с сотрудником " & vbCrLf & ФИОКорРук(ComboBox19.Text, False) & " сформирован успешно!" & vbCrLf & "Распечатать документ!", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.None) = DialogResult.OK Then
            ПечатьДоковFTP(massFTP3, 2)
        End If
        Me.Cursor = Cursors.Default

    End Sub
    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            MaskedTextBox2.Focus()

            Dim pl As String
            If TextBox4.Text <> "" Then
                Dim i As Integer = CInt(TextBox4.Text)
                Select Case i

                    Case < 10
                        pl = Str(i)
                        TextBox4.Text = "00" & i

                    Case 10 To 99
                        pl = Str(i)
                        TextBox4.Text = "0" & i
                End Select
            End If
        End If
    End Sub

    Private Sub Grid1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellDoubleClick
        КодАкт = Grid1.CurrentRow.Cells("Код").Value
        Dim list As New Dictionary(Of String, Object)
        list.Add("@Код", КодАкт)

        Dim ds = Selects(StrSql:="SELECT ДогПодрядаАкт.*, ДогПодряда.НомерДогПодр
FROM ДогПодряда INNER JOIN ДогПодрядаАкт ON ДогПодряда.Код = ДогПодрядаАкт.IDДогПодр
WHERE ДогПодрядаАкт.Код=@Код", list)

        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        MaskedTextBox1.Text = ""
        MaskedTextBox2.Text = ""
        MaskedTextBox3.Text = ""
        MaskedTextBox6.Text = ""

        TextBox1.Text = ds.Rows(0).Item(5).ToString
        TextBox2.Text = ds.Rows(0).Item(6).ToString
        TextBox3.Text = ds.Rows(0).Item(7).ToString
        TextBox4.Text = ds.Rows(0).Item(2).ToString
        MaskedTextBox1.Text = ds.Rows(0).Item(9).ToString
        MaskedTextBox2.Text = ds.Rows(0).Item(3).ToString
        MaskedTextBox3.Text = ds.Rows(0).Item(4).ToString
        MaskedTextBox6.Text = ds.Rows(0).Item(8).ToString
        ПослНомАкт = Nothing
        ПослНомАкт = ds.Rows(0).Item(2)
        ПослНомерДоговора = ""
        ПослНомерДоговора = ds.Rows(0).Item(10).ToString
        ComboBox3.Text = ds.Rows(0).Item(10).ToString

    End Sub

    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick
        If ListBox1.SelectedIndex = -1 Then
            MessageBox.Show("Выберите документ для просмотра!", Рик, MessageBoxButtons.OK)
            Exit Sub
        End If

        Dim l As String = ComboBox4.Items.Item(ListBox1.SelectedIndex)

        ВыгрузкаФайловНаЛокалыныйКомп(l, PathVremyanka & ListBox1.SelectedItem)

        Dim proc As Process = Process.Start(PathVremyanka & ListBox1.SelectedItem)
        proc.WaitForExit()
        proc.Close()

        ЗагрНаСерверИУдаление(PathVremyanka & ListBox1.SelectedItem, l, ListBox1.SelectedItem)


        'Dim ff As ListBox.SelectedIndexCollection = ListBox1.SelectedIndices


        'If Not ListBox1.SelectedIndex = -1 Then

        '    For Each p As Integer In ff
        '        Process.Start(FilesList(p))
        '    Next

        'End If
    End Sub

    Private Sub ЗаполнДанн(ByVal ds As DataTable)

        КодАкт = Nothing
        КодАкт = ds.Rows(0).Item(7)
        MaskedTextBox2.Text = ds.Rows(0).Item(2).ToString
        If ds.Rows(0).Item(1).ToString.Length = 1 Then
            TextBox4.Text = "00" & ds.Rows(0).Item(1).ToString
        ElseIf ds.Rows(0).Item(1).ToString.Length = 2 Then
            TextBox4.Text = "0" & ds.Rows(0).Item(1).ToString
        Else
            TextBox4.Text = ds.Rows(0).Item(1).ToString
        End If

        MaskedTextBox3.Text = ds.Rows(0).Item(3).ToString
        TextBox1.Text = ds.Rows(0).Item(4).ToString
        TextBox2.Text = ds.Rows(0).Item(5).ToString
        TextBox3.Text = ds.Rows(0).Item(6).ToString
        MaskedTextBox6.Text = ds.Rows(0).Item(8).ToString
        MaskedTextBox1.Text = ds.Rows(0).Item(9).ToString
        ПослНомАкт = Nothing
        ПослНомАкт = ds.Rows(0).Item(1)


    End Sub
    Private Sub ListBox1_DrawItem(sender As Object, e As DrawItemEventArgs) Handles ListBox1.DrawItem

        If e.Index < 0 Then Exit Sub

        e.DrawBackground()

        Dim bcolor As Color = If(Equals(e.BackColor, ListBox1.BackColor),
                                  ListBox1.BackColor,
                                  Color.LightGreen)




        e.Graphics.FillRectangle(New SolidBrush(bcolor), e.Bounds)

        e.Graphics.DrawString(ListBox1.Items.Item(e.Index).ToString,
                              ListBox1.Font,
                              New SolidBrush(ListBox1.ForeColor),
                              e.Bounds,
                              New StringFormat With {.Alignment = StringAlignment.Near,
                                                     .LineAlignment = StringAlignment.Center,
                                                     .Trimming = StringTrimming.None,
                                                     .FormatFlags = StringFormatFlags.NoWrap})

    End Sub
    Private Sub ДоговорПодрядаАкт_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        dtPutiDokumentov() 'обновляем пути документов

        Me.ComboBox1.AutoCompleteCustomSource.Clear()
        Me.ComboBox1.Items.Clear()
        For Each r As DataRow In СписокКлиентовОсновной.Rows
            Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox1.Items.Add(r(0).ToString)
        Next

        MaskedTextBox1.Enabled = False
        MaskedTextBox2.Text = Now.Date
        MaskedTextBox3.Text = Now.Date
        MaskedTextBox6.Text = Now.Date
        ВычДатВыплат(MaskedTextBox6.Text)
        TextBox2.Enabled = False
    End Sub
    Private Sub ВычДатВыплат(ByVal dt As Date)
        Dim int As Integer = dt.Month
        Dim год As Integer = dt.Year

        If int = 12 Then
            Dim dp As Date = CDate("01." & "01." & CType(год + 1, String))
            MaskedTextBox1.Text = dp.AddDays(10)
        Else
            int = int + 1
            Dim st As String = CType(int, String)
            If int < 10 Then
                st = CType("0" & int, String)
            End If
            Dim dp As Date = CDate("01." & st & "." & CType(год, String))
            MaskedTextBox1.Text = dp.AddDays(9)
        End If


    End Sub

    Private Sub ДоговорПодрядаАкт_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        ComboBox1.Text = ""
        ComboBox1.Items.Clear()
        ComboBox19.Text = ""
        ComboBox19.Items.Clear()
        ComboBox3.Text = ""
        'ComboBox3.Items.Clear()
        ComboBox2.Text = ""
        ComboBox2.Items.Clear()
        TextBox4.Text = ""
        TextBox3.Text = ""
        TextBox2.Text = ""
        TextBox1.Text = ""
        MaskedTextBox6.Text = ""
        MaskedTextBox1.Text = ""
        MaskedTextBox2.Text = ""
        MaskedTextBox3.Text = ""
        ListBox1.Items.Clear()
        Grid1.DataSource = Nothing

    End Sub

    Private Sub TextBox2_LostFocus(sender As Object, e As EventArgs) Handles TextBox2.LostFocus

        'If TextBox2.Text <> "" And IsNumeric(TextBox2.Text) And TextBox3.Text <> "" And IsNumeric(TextBox3.Text) Then
        '    TextBox1.Text = ""
        '    Dim f As Double
        '    f = Replace(TextBox3.Text, ".", ",")
        '    TextBox3.SelectionStart = TextBox3.Text.Length '14

        '    If СправкаПоЗарплате.bool(f) = True Then
        '        TextBox3.Text = f & ",00"
        '    Else
        '        TextBox3.Text = f
        '        If СправкаПоЗарплате.Count(f) = 1 Then
        '            TextBox3.Text = f & "0"
        '        End If
        '    End If
        '    Button1.Focus()
        '    If TextBox2.Text <> "" Then
        '        txt1(TextBox3.Text)
        '    End If
        'End If

        If TextBox2.Text <> "" And IsNumeric(TextBox2.Text) And TextBox3.Text <> "" And IsNumeric(TextBox3.Text) Then
            Dim f As Double
            f = Replace(TextBox2.Text, ".", ",")
            TextBox2.SelectionStart = TextBox2.Text.Length '14

            If СправкаПоЗарплате.bool(f) = True Then
                TextBox2.Text = f & ",00"
            Else
                TextBox2.Text = f
                If СправкаПоЗарплате.Count(f) = 1 Then
                    TextBox2.Text = f & "0"
                End If
            End If
            If TextBox3.Text <> "" Then
                TextBox1.Text = ""
                Dim f1 As Double
                Dim f2 As Double = CDbl(TextBox3.Text)
                f1 = f2 / f
                f1 = Math.Round(f1, 2)
                If СправкаПоЗарплате.bool(f1) = True Then
                    TextBox1.Text = f1 & ",00"
                Else
                    TextBox1.Text = f1
                    If СправкаПоЗарплате.Count(f1) = 1 Then
                        TextBox1.Text = f1 & "0"
                    End If
                End If
            End If
        End If
        'TextBox3.Focus()










    End Sub

    Private Sub TextBox3_LostFocus(sender As Object, e As EventArgs) Handles TextBox3.LostFocus
        If TextBox2.Text <> "" And IsNumeric(TextBox2.Text) And TextBox3.Text <> "" And IsNumeric(TextBox3.Text) Then
            TextBox1.Text = ""
            Dim f As Double
            f = Replace(TextBox3.Text, ".", ",")
            TextBox3.SelectionStart = TextBox3.Text.Length '14

            If СправкаПоЗарплате.bool(f) = True Then
                TextBox3.Text = f & ",00"
            Else
                TextBox3.Text = f
                If СправкаПоЗарплате.Count(f) = 1 Then
                    TextBox3.Text = f & "0"
                End If
            End If
            Button1.Focus()
            If TextBox2.Text <> "" Then
                txt1(TextBox3.Text)
            End If
        End If
    End Sub

    Private Sub TextBox4_LostFocus(sender As Object, e As EventArgs) Handles TextBox4.LostFocus
        If TextBox4.Text <> "" And IsNumeric(TextBox4.Text) Then
            MaskedTextBox2.Focus()
            Dim pl As String
            If TextBox4.Text <> "" Then
                Dim i As Integer = CInt(TextBox4.Text)
                Select Case i

                    Case < 10
                        pl = Str(i)
                        TextBox4.Text = "00" & i

                    Case 10 To 99
                        pl = Str(i)
                        TextBox4.Text = "0" & i
                End Select
            End If
        End If

    End Sub
End Class