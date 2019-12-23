
Option Explicit On
Imports System.Data.OleDb



Public Class ШтатноеПослеИзменения

    Public Da As New OleDbDataAdapter 'Адаптер
        'Public Ds As New DataSet 'Пустой набор записей
        Dim tbl As New DataTable
        Dim ds As DataTable
        Dim cb As OleDb.OleDbCommandBuilder
        Dim Рик As String = "ООО РикКонсалтинг"
        Dim Год, Организ, Должность, Отдел, Разряд, Процент, ТарСтавка, thb0, thb, StrSql As String
        Dim s, s2, se, ip, mas, изменен, srt, КодDBC, ГлКод, КодДолжн As Integer
        Dim Отд, Дол, Раз, ТСтавка, ПовышПроц As String

        Private Sub TextBox4_SelectedIndexChanged(sender As Object, e As EventArgs)
            '        Чист()
            '        StrSql = "SELECT DISTINCT ШтСвод.Должность FROM ШтОтделы INNER JOIN ШтСвод ON ШтОтделы.Код = ШтСвод.Отдел
            'WHERE ШтОтделы.Клиент='" & ComboBox1.Text & "' AND ШтОтделы.Отделы='" & TextBox4.Text & "'"
            '        ds = Selects(StrSql)

            '        TextBox5.Text = ""
            '        Me.TextBox5.AutoCompleteCustomSource.Clear()
            '        Me.TextBox5.Items.Clear()
            '        For Each r As DataRow In ds.Rows
            '            Me.TextBox5.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            '            Me.TextBox5.Items.Add(r(0).ToString)
            '        Next

            '        Чист()
            '        StrSql = "SELECT Код FROM ШтОтделы WHERE ШтОтделы.Клиент='" & ComboBox1.Text & "' AND ШтОтделы.Отделы='" & TextBox4.Text & "'"
            '        ds = Selects(StrSql)
            '        ГлКод = Nothing
            '        ГлКод = ds.Rows(0).Item(0)

        End Sub

        Private Sub TextBox5_SelectedIndexChanged(sender As Object, e As EventArgs)
            'ЗагрПроцОклРазр()
        End Sub

        Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

        End Sub
        Private Sub Добавить()
            If MessageBox.Show("Добавить данные?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.Cancel Then
                Exit Sub
            End If

            Чист()
            StrSql = "SELECT ШтСвод.КодШтСвод FROM ШтОтделы INNER JOIN ШтСвод ON ШтОтделы.Код = ШтСвод.Отдел
WHERE ШтОтделы.Клиент='" & ComboBox1.Text & "' AND ШтОтделы.Отделы='" & Trim(TextBox4.Text) & "' AND ШтСвод.Должность ='" & Trim(TextBox5.Text) & "' AND ШтСвод.Разряд='" & TextBox3.Text & "'"
            ds = Selects(StrSql)

            Try
                If IsNumeric(ds.Rows(0).Item(0)) = True Then
                    MessageBox.Show("В организации " & ComboBox1.Text & "уже есть отдел " & Trim(TextBox4.Text) & " с должностью " & Trim(TextBox5.Text) & " и разрядом " & TextBox3.Text & "." & vbCrLf & "Добавить такой же отдел с такой-же должностью и разрядом невозможно!", Рик)
                    Exit Sub
                End If
            Catch ex As Exception

            End Try





            Чист()
            StrSql = "SELECT Код FROM ШтОтделы WHERE ШтОтделы.Клиент='" & ComboBox1.Text & "' AND ШтОтделы.Отделы='" & Trim(TextBox4.Text) & "'"
            ds = Selects(StrSql)
            Try
                If IsNumeric(ds.Rows(0).Item(0)) = True Then
                    StrSql = ""
                StrSql = "INSERT INTO ШтСвод(Отдел, Должность,Разряд,ТарифнаяСтавка,ПовышениеПроц,ТарСтПослеИспСрока,ПовПроцПослеИспСрока)
VALUES(" & ds.Rows(0).Item(0) & ",'" & TextBox5.Text & "','" & TextBox3.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox7.Text & "','" & TextBox6.Text & "')"
                Updates(StrSql)
                    MessageBox.Show("Данные добавлены!", Рик)
                End If
            Catch ex As Exception
                Чист()
                StrSql = "INSERT INTO ШтОтделы(Клиент, Отделы) VALUES('" & ComboBox1.Text & "','" & Trim(TextBox4.Text) & "')"
                Updates(StrSql)


                Чист()
                StrSql = "SELECT ШтОтделы.Код FROM ШтОтделы WHERE ШтОтделы.Клиент='" & ComboBox1.Text & "' AND ШтОтделы.Отделы='" & Trim(TextBox4.Text) & "'"
                ds = Selects(StrSql)
                Dim idsotr As Integer = ds.Rows(0).Item(0)

            StrSql = "INSERT INTO ШтСвод(Отдел, Должность,Разряд,ТарифнаяСтавка,ПовышениеПроц,ТарСтПослеИспСрока,ПовПроцПослеИспСрока)
VALUES(" & idsotr & ",'" & TextBox5.Text & "','" & TextBox3.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "' ,'" & TextBox7.Text & "','" & TextBox6.Text & "')"
            Updates(StrSql)


                MessageBox.Show("Данные добавлены!", Рик)
            End Try







        End Sub
        Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
            If TextBox4.Text = "" Or TextBox5.Text = "" Then
                MessageBox.Show("Поле отдел и должность не могут быть пустыми!")
                Exit Sub
            End If
            Добавить()
            Очистка()
            Refreshgrid()

        End Sub
        Private Sub Изменить()
            If MessageBox.Show("Изменить данные?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.Cancel Then
                Exit Sub
            End If

            'Чист()
            'StrSql = "SELECT ШтОтделы.Код FROM ШтОтделы WHERE ШтОтделы.Клиент='" & ComboBox1.Text & "' AND ШтОтделы.Отделы='" & TextBox4.Text & "'"
            'ds = Selects(StrSql)
            'Dim idsotr As Integer = ds.Rows(0).Item(0)
            Try
                Чист()
            StrSql = "Update ШтСвод Set Должность='" & TextBox5.Text & "', Разряд='" & TextBox3.Text & "',ТарифнаяСтавка='" & TextBox1.Text & "',
ПовышениеПроц='" & TextBox2.Text & "', ТарСтПослеИспСрока='" & TextBox7.Text & "', ПовПроцПослеИспСрока='" & TextBox6.Text & "'
WHERE ШтСвод.КодШтСвод =" & КодДолжн & ""
            Updates(StrSql)

                Чист()
                StrSql = "Update ШтОтделы Set Отделы='" & TextBox4.Text & "' WHERE ШтОтделы.Код =" & ГлКод & ""
                Updates(StrSql)
                MessageBox.Show("Данные изменены!", Рик)
            Catch ex As Exception
                MessageBox.Show("В базе нет данных относительно вашего запроса!", Рик)
            End Try


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
        Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
            If TextBox4.Text = "" Or TextBox5.Text = "" Then
                MessageBox.Show("Поле отдел и должность не могут быть пустыми!")
                Exit Sub
            End If


            Изменить()
            Очистка()
            Refreshgrid()
        End Sub
        Private Sub ПровЗаполн()

        End Sub
        Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
            If TextBox4.Text = "" Then
                MessageBox.Show("Поле отдел не может быть пустым!")
                Exit Sub
            End If
            УдалитьОтдел()
            Очистка()
            Refreshgrid()
        End Sub
        Private Sub УдалитьОтдел()
            If MessageBox.Show("Будет удален отдел и все должности!", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.Cancel Then
                Exit Sub
            End If
            Чист()
        StrSql = "DELETE FROM ШтОтделы WHERE Код =" & ГлКод & ""
        Updates(StrSql)
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
        TextBox6.Text = ""
        TextBox7.Text = ""
        КодДолжн = Nothing
        ГлКод = Nothing
    End Sub
    Private Sub Удалить()

            If MessageBox.Show("Удалить должность!", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.Cancel Then
                Exit Sub
            End If
            Чист()
            StrSql = "SELECT COUNT (Отдел) FROM ШтСвод WHERE ШтСвод.Отдел =" & ГлКод & ""
            ds = Selects(StrSql)
            If ds.Rows(0).Item(0) > 1 Then

                Чист()
            StrSql = "DELETE FROM ШтСвод WHERE КодШтСвод =" & КодДолжн & ""
            Updates(StrSql)

                MessageBox.Show("Данные удалены!", Рик)
            Else

                Чист()
            StrSql = "DELETE FROM ШтОтделы WHERE Код =" & ГлКод & ""
            Updates(StrSql)

                MessageBox.Show("Данные удалены!", Рик)
            End If

        End Sub

        Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

        End Sub

        Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
            If TextBox4.Text = "" Or TextBox5.Text = "" Then
                MessageBox.Show("Поле отдел и должность не могут быть пустыми!")
                Exit Sub
            End If
            Удалить()
            Очистка()
            Refreshgrid()
        End Sub

        Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
            If e.KeyCode = Keys.Enter Then
                e.SuppressKeyPress = True
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
            e.SuppressKeyPress = True
            Me.TextBox2.Focus()
        End If
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
            If e.KeyCode = Keys.Enter Then
                e.SuppressKeyPress = True
                Button5.Focus()
            End If
        End Sub

        Dim ОтдDBC, ДолDBC, РазDBC, ТСтавкаDBC, ПовышПроцDBC As String

    Dim mas2, mas3

    Private Sub ШтатноеПослеИзменения_Load(sender As Object, e As EventArgs) Handles MyBase.Load

            Me.MdiParent = MDIParent1
            Me.WindowState = FormWindowState.Maximized

        Год = Year(Now)
        'If Me.Прием_Load = vbTrue Then Form1.Load = False


        Me.ComboBox1.Items.Clear()
        For Each r As DataRow In СписокКлиентовОсновной.Rows
            Me.ComboBox1.Items.Add(r(0).ToString)
        Next
        Me.TextBox42.Text = DateTime.Now.ToString("dd.MM.yyyy")

            For i As Integer = 0 To Grid1.Rows.Count - 1
                For y As Integer = 0 To Grid1.Columns.Count - 1
                    Grid1.Item(y, i).Style.Font = New Font("times new roman", 11)
                Next
            Next



        End Sub
        'Private Sub NumberAllRows() ' Add row headers.'нумерация строк
        '    If Grid1.RowCount > 0 Then
        '        For i As Integer = 0 To Grid1.Rows.Count - 1
        '            Grid1.Rows(i).HeaderCell.Value = Val(i + 1).ToString
        '        Next i
        '    End If
        'End Sub
        Private Sub Refreshgrid()
            Организ = ComboBox1.Text
            Dim StrSql1 As String
            tbl.Clear()
            StrSql1 = "SELECT ШтОтделы.Код, ШтОтделы.Клиент, ШтОтделы.Отделы, ШтСвод.Должность, ШтСвод.Разряд, 
ШтСвод.ТарифнаяСтавка as Ставка, ШтСвод.ПовышениеПроц as Процент, ШтСвод.КодШтСвод,
ШтСвод.ТарСтПослеИспСрока as [Ставка после испыт_срока], ШтСвод.ПовПроцПослеИспСрока as [Процент после испыт_срока]
FROM ШтОтделы INNER JOIN ШтСвод ON ШтОтделы.Код = ШтСвод.Отдел
        WHERE ШтОтделы.Клиент = '" & Организ & "'"
            Dim c As New OleDbCommand
            c.Connection = conn
            c.CommandText = StrSql1
            'Dim ds As New DataSet
            Dim da As New OleDbDataAdapter(c)
            'da.Fill(ds, "Сотрудники")
            da.Fill(tbl)
            Grid1.DataSource = tbl
            Grid1.Columns(1).Visible = False
            Grid1.Columns(7).Visible = False
            Grid1.Columns(0).Visible = False
            'Grid1.Columns(4).Width = 60
            'Grid1.Columns(5).Width = 100
            Grid1.Columns(4).Width = 60

            Grid1.Columns(5).Width = 80
        Grid1.Columns(6).Width = 80
        Grid1.Columns(8).Width = 80
        Grid1.Columns(9).Width = 90

        'Grid1.Columns(6).Width = 60

        'Grid1.Rows(1).Cells(3).Selected = True
        'Grid1_CellClick(Grid1, New DataGridViewCellEventArgs(<b>3</b>, <b>1</b>))
        'Acti()
        cb = New OleDb.OleDbCommandBuilder(da)
            s = Grid1.Rows.Count - 1
            изменен = 0
            'NumberAllRows()
        End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged

            ЗагрОтделов()
            Refreshgrid()



            '        Dim StrSql As String
            '        StrSql = "SELECT ШтОтделы.Отделы
            'FROM Клиент INNER JOIN ШтОтделы ON Клиент.НазвОрг = ШтОтделы.Клиент
            'WHERE Клиент.НазвОрг='" & Организ & "'"
            '        Dim c1 As New OleDbCommand With {
            '                .Connection = conn,
            '                .CommandText = StrSql
            '            }
            '        Dim ds1 As New DataTable
            '        Dim da1 As New OleDbDataAdapter(c1)
            '        da1.Fill(ds1)
            '        Me.TextBox4.Items.Clear()

            '        For Each r As DataRow In ds1.Rows
            '            Me.TextBox4.Items.Add(r(0).ToString)
            '        Next

        End Sub

        Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
            s = Grid1.Rows.Count - 1
            s2 = Grid1.Rows.Count - 2

        End Sub

        Private Sub ЗагрОтделов()

            'StrSql = ""
            'StrSql = "SELECT DISTINCT ШтОтделы.Отделы From ШтОтделы WHERE ШтОтделы.Клиент='" & ComboBox1.Text & "'"
            'ds = Selects(StrSql)
            'TextBox4.Text = ""
            'Me.TextBox4.AutoCompleteCustomSource.Clear()
            'Me.TextBox4.Items.Clear()
            'For Each r As DataRow In ds.Rows
            '    Me.TextBox4.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            '    Me.TextBox4.Items.Add(r(0).ToString)
            'Next
            'TextBox5.Text = ""
            'Me.TextBox5.AutoCompleteCustomSource.Clear()
            'Me.TextBox5.Items.Clear()

            'TextBox1.Text = ""
            'TextBox2.Text = ""
            'TextBox3.Text = ""

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
                Dim Str As String = ""

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
                    Dim c As New OleDbCommand
                    c.Connection = conn
                    c.CommandText = StrSql
                    c.ExecuteNonQuery()

                Next

            Else
                Dim fg2 As Integer = ПровЗапПолей()
                Select Case fg2
                    Case 1
                        Exit Sub
                End Select

                Dim ses As Integer = se - s
                Dim StrSql2, StrSql4 As String
                Dim c1, c2 As New OleDbCommand
                Dim ds1 As New DataSet
                Dim da1 As New OleDbDataAdapter(c1)

                Select Case ses
                    Case > 0

                        Dim coli As Integer
                        Dim i As Integer 'проверяем есть ли в базе уже такая должность и если есть присваиваем код соответсвующий должности
                        For i = 0 To ses - 1
                            Dim sOtd As String = Grid1.Rows(s + i).Cells(2).Value.ToString
                            Try ' заполняем базу и возвращаем номер ИД дл
                                StrSql2 = "Select Код FROM ШтОтделы WHERE Клиент = '" & Организ & "'   AND Отделы='" & sOtd & "' "
                                With c1
                                    .Connection = conn
                                    .CommandText = StrSql2
                                End With

                                da1.Fill(ds1, "f")
                                Grid1.Rows(s + i).Cells(0).Value = ds1.Tables("f").Rows(0).Item(0).ToString

                                Dim коднов As Integer = ds1.Tables("f").Rows(0).Item(0)
                                coli = 1
                                coli += i


                                StrSql4 = "INSERT INTO ШтСвод(Отдел,Должность,Разряд,ТарифнаяСтавка,ПовышениеПроц)
VALUES (" & коднов & ",'" & Grid1.Rows(s + i).Cells(3).Value & "','" & Grid1.Rows(s + i).Cells(4).Value & "','" & Grid1.Rows(s + i).Cells(5).Value & "','" & Grid1.Rows(s + i).Cells(6).Value & "')" ' вставляем в базу должность, тар.ставку, повыш, и разряд

                                c2.Connection = conn
                                c2.CommandText = StrSql4
                                c2.ExecuteNonQuery()
                            Catch ex As Exception

                            End Try
                        Next
                        If coli = ses Then
                            Exit Sub
                        End If
                End Select

                For i = 0 To ses - 1

                    Dim StrSql1 As String = "INSERT INTO ШтОтделы(Клиент,Отделы) VALUES ('" & Организ & "','" & Grid1.Rows(s + i).Cells(2).Value & "')"
                    Dim c As New OleDbCommand
                    c.Connection = conn
                    c.CommandText = StrSql1
                    c.ExecuteNonQuery()

                    Dim StrSql5 As String = "SELECT Код FROM ШтОтделы WHERE Клиент = '" & Организ & "' AND Отделы='" & Grid1.Rows(s + i).Cells(2).Value & "' "
                    Dim c5 As New OleDbCommand With {
                .Connection = conn,
                .CommandText = StrSql5
            }
                    Dim ds5 As New DataSet
                    Dim da5 As New OleDbDataAdapter(c5)
                    da5.Fill(ds5, "f")
                    Grid1.Rows(s + i).Cells(0).Value = ds5.Tables("f").Rows(0).Item(0).ToString

                    Dim длж As String = Grid1.Rows(s + i).Cells(3).Value.ToString
                    Dim разр As String = Grid1.Rows(s + i).Cells(4).Value.ToString
                    Dim тстав As String = Grid1.Rows(s + i).Cells(5).Value.ToString
                    Dim проц As String = Grid1.Rows(s + i).Cells(6).Value.ToString

                    Dim StrSql6 As String = "INSERT INTO ШтСвод(Отдел,Должность,Разряд,ТарифнаяСтавка,ПовышениеПроц) VALUES (" & ds5.Tables("f").Rows(0).Item(0) & ",'" & длж & "','" & разр & "','" & тстав & "','" & проц & "')" ' вставляем в базу должность, тар.ставку, повыш, и разряд
                    Dim c6 As New OleDbCommand
                    c6.Connection = conn
                    c6.CommandText = StrSql6
                    c6.ExecuteNonQuery()

            Next
            End If



Конец:
            Refreshgrid()
            ''        Разряд = TextBox3.Text
            ''        Процент = TextBox2.Text
            ''        ТарСтавка = TextBox1.Text

            ''        Dim StrSql As String
            ''        StrSql = "SELECT ШтОтделы.Отделы
            ''FROM Клиент INNER JOIN ШтОтделы ON Клиент.НазвОрг = ШтОтделы.Клиент
            ''WHERE Клиент.НазвОрг='" & Организ & "'"
            ''        Dim c1 As New OleDbCommand With {
            ''                .Connection = conn,
            ''                .CommandText = StrSql
            ''            }
            ''        Dim ds1 As New DataSet
            ''        Dim da1 As New OleDbDataAdapter(c1)
            ''        da1.Fill(ds1, "Customers")
            ''        Dim foundRows() As Data.DataRow

            ''        'Dim foundRows() As Data.DataRow Образец выборки из ДатаСет
            ''        'foundRows = DataSet1.Tables("Customers").Select("CompanyName Like 'A%'")


            ''        foundRows = ds1.Tables("Customers").Select("Отделы Like '" & Отдел & "'")
            ''        If foundRows.Length = 1 Then

            ''            MsgBox("ok")
            ''        End If
        End Sub

        Private Sub Grid1_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles Grid1.CellBeginEdit
            Select Case e.ColumnIndex
                Case 0
                    e.Cancel = True
            End Select

            'ОтдDBC = Grid1.CurrentRow.Cells("Отделы").Value.ToString
            'ДолDBC = Grid1.CurrentRow.Cells("Должность").Value.ToString
            'РазDBC = Grid1.CurrentRow.Cells("Разряд").Value.ToString
            'ТСтавкаDBC = Grid1.CurrentRow.Cells("ТарифнаяСтавка").Value.ToString
            'ПовышПроцDBC = Grid1.CurrentRow.Cells("ПовышениеПроц").Value.ToString
            'КодDBC = Grid1.CurrentRow.Cells("КодШтСвод").Value

            'mas2 = Grid1.CurrentCellAddress.X
            'mas3 = Grid1.CurrentCellAddress.Y
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








            'srt = Grid1.CurrentRow.Cells("Код").Value
            'Отд = Grid1.CurrentRow.Cells("Отделы").Value.ToString
            'Дол = Grid1.CurrentRow.Cells("Должность").Value.ToString
            'Раз = Grid1.CurrentRow.Cells("Разряд").Value.ToString
            'ТСтавка = Grid1.CurrentRow.Cells("ТарифнаяСтавка").Value.ToString
            'ПовышПроц = Grid1.CurrentRow.Cells("ПовышениеПроц").Value.ToString

            'Dim StrSql As String = "UPDATE ШтСвод  SET ТарифнаяСтавка= '" & ТСтавка & "',ПовышениеПроц='" & ПовышПроц & "',Должность= '" & Дол & "',Разряд='" & Раз & "'
            '    WHERE ШтСвод.Код=" & srt & ""
            'Dim c As New OleDbCommand
            'c.Connection = conn
            'c.CommandText = StrSql
            'Try
            '    c.ExecuteNonQuery()
            'Catch ex As Exception
            '    Exit Sub
            'End Try

            'Dim StrSql2 As String = "UPDATE ШтОтделы  SET Отделы= '" & Отд & "'
            '    WHERE ШтСвод.Код=" & srt & ""
            'Dim c2 As New OleDbCommand
            'c2.Connection = conn
            'c2.CommandText = StrSql2
            'Try
            '    c2.ExecuteNonQuery()
            '    MessageBox.Show("Данные изменены!", Рик)
            'Catch ex As Exception
            '    Exit Sub
            'End Try

            '''MsgBox(Grid1.CurrentRow.Cells("Код").Value.ToString)
            '''thb0 = Grid1.CurrentRow.Cells("Код").ToString

        End Sub

        Private Sub Grid1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellValueChanged
            If (e.ColumnIndex = -1) Then Return



            изменен = Grid1.CurrentCellAddress.Y
            изменен += 1
            'If Grid1.CurrentRow.Cells("Код").Value.ToString <> "" Then
            '    MsgBox(Grid1.CurrentRow.Cells("Код").Value.ToString)
            'End If

        End Sub

        Private Sub Grid1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellDoubleClick

            'ОтдDBC = Grid1.CurrentRow.Cells("Отделы").Value.ToString
            'ДолDBC = Grid1.CurrentRow.Cells("Должность").Value.ToString
            'РазDBC = Grid1.CurrentRow.Cells("Разряд").Value.ToString
            'ТСтавкаDBC = Grid1.CurrentRow.Cells("ТарифнаяСтавка").Value.ToString
            'ПовышПроцDBC = Grid1.CurrentRow.Cells("ПовышениеПроц").Value.ToString

            'Dim s As String = sender.ToString
            'MsgBox(s)

            '        TextBox4.Text = Grid1.CurrentRow.Cells("Отделы").Value.ToString()
            '        TextBox5.Text = Grid1.CurrentRow.Cells("Должность").Value.ToString
            '        TextBox3.Text = Grid1.CurrentRow.Cells("Разряд").Value.ToString
            '        TextBox1.Text = Grid1.CurrentRow.Cells("ТарифнаяСтавка").Value.ToString
            '        TextBox2.Text = Grid1.CurrentRow.Cells("ПовышениеПроц").Value.ToString


            '        '        StrSql = "SELECT ШтОтделы.Код FROM ШтОтделы WHERE ШтОтделы.Клиент='" & ComboBox1.Text & "' AND ШтСвод.Должность='" & TextBox5.Text & "'
            '        'AND ШтСвод.Разряд ='" & TextBox3.Text & "'  AND ШтСвод.ТарифнаяСтавка='" & TextBox1.Text & "'
            '        'AND ШтСвод.ПовышениеПроц='" & TextBox2.Text & "' AND ШтОтделы.Отделы ='" & TextBox4.Text & "'"


            '        StrSql = "SELECT ШтСвод.КодШтСвод, ШтОтделы.Код FROM ШтОтделы INNER JOIN ШтСвод ON ШтОтделы.Код = ШтСвод.Отдел
            'WHERE ШтОтделы.Клиент='" & ComboBox1.Text & "' AND ШтСвод.Должность='" & TextBox5.Text & "'
            'AND ШтСвод.Разряд ='" & TextBox3.Text & "'  AND ШтСвод.ТарифнаяСтавка='" & TextBox1.Text & "'
            'AND ШтСвод.ПовышениеПроц='" & TextBox2.Text & "' AND ШтОтделы.Отделы ='" & TextBox4.Text & "'"

            '        ds = Selects(StrSql)
            '        ГлКод = Nothing
            '        ГлКод = ds.Rows(0).Item(1)
            '        КодДолжн = Nothing
            '        КодДолжн = ds.Rows(0).Item(0)


        End Sub

        Private Sub Grid1_KeyDown(sender As Object, e As KeyEventArgs) Handles Grid1.KeyDown
            'Try
            '    srt = Grid1.CurrentRow.Cells("Код").Value
            '    Отд = Grid1.CurrentRow.Cells("Отделы").Value.ToString
            '    Дол = Grid1.CurrentRow.Cells("Должность").Value.ToString
            '    Раз = Grid1.CurrentRow.Cells("Разряд").Value.ToString
            '    ТСтавка = Grid1.CurrentRow.Cells("ТарифнаяСтавка").Value.ToString
            '    ПовышПроц = Grid1.CurrentRow.Cells("ПовышениеПроц").Value.ToString
            'Catch ex As Exception
            '    Exit Sub
            'End Try



            'If e.KeyCode = Keys.Enter Then
            '    e.SuppressKeyPress = True


            '    Dim StrSql As String = "UPDATE ШтСвод  Set ТарифнаяСтавка= '" & ТСтавка & "',ПовышениеПроц='" & ПовышПроц & "',Должность= '" & Дол & "',Разряд='" & Раз & "'
            '    WHERE ШтСвод.Код=" & srt & ""
            '    Dim c As New OleDbCommand
            '    c.Connection = conn
            '    c.CommandText = StrSql
            '    Try
            '        c.ExecuteNonQuery()
            '    Catch ex As Exception
            '        Exit Sub
            '    End Try

            '    Dim StrSql2 As String = "UPDATE ШтОтделы  SET Отделы= '" & Отд & "'
            '    WHERE Код=" & srt & ""
            '    Dim c2 As New OleDbCommand
            '    c2.Connection = conn
            '    c2.CommandText = StrSql2
            '    Try
            '        c2.ExecuteNonQuery()
            '        MessageBox.Show("Данные изменены!", Рик)
            '    Catch ex As Exception
            '        Exit Sub
            '    End Try
            'End If
        End Sub
        Private Sub Acti()
        Try

            TextBox4.Text = Grid1.CurrentRow.Cells("Отделы").Value.ToString
            TextBox5.Text = Grid1.CurrentRow.Cells("Должность").Value.ToString
            TextBox3.Text = Grid1.CurrentRow.Cells("Разряд").Value.ToString
            TextBox1.Text = Grid1.CurrentRow.Cells("Ставка").Value.ToString
            TextBox2.Text = Grid1.CurrentRow.Cells("Процент").Value.ToString
            TextBox7.Text = Grid1.CurrentRow.Cells("Ставка после испыт_срока").Value.ToString
            TextBox6.Text = Grid1.CurrentRow.Cells("Процент после испыт_срока").Value.ToString
        Catch ex As Exception
            MessageBox.Show("Кликните по полю таблицы!", Рик)
                Exit Sub
            End Try



            StrSql = "SELECT ШтСвод.КодШтСвод, ШтОтделы.Код FROM ШтОтделы INNER JOIN ШтСвод ON ШтОтделы.Код = ШтСвод.Отдел
        WHERE ШтОтделы.Клиент='" & ComboBox1.Text & "' AND ШтСвод.Должность='" & TextBox5.Text & "'
        AND ШтСвод.Разряд ='" & TextBox3.Text & "'  AND ШтСвод.ТарифнаяСтавка='" & TextBox1.Text & "'
        AND ШтСвод.ПовышениеПроц='" & TextBox2.Text & "' AND ШтОтделы.Отделы ='" & TextBox4.Text & "'"

            ds = Selects(StrSql)
            ГлКод = Nothing
            ГлКод = ds.Rows(0).Item(1)
            КодДолжн = Nothing
            КодДолжн = ds.Rows(0).Item(0)
        End Sub
        Private Sub Grid1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellClick

            Acti()

        End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox5.Focus()
        End If
    End Sub
End Class
