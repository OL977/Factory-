Option Explicit On
Imports System.Data.OleDb
Public Class СправкаПоЗарплате
    Dim gd, подох, фсзн, прочие, итогоудерж, квыдаче As Double
    Dim цел, дроб As Integer
    Dim СохрЗак As String
    Dim МаксДох, СумВыч, ОдинРеб, ДваРеб As Double
    Dim r1() As Object
    Dim dx As DataTable
    Dim massFTP3 As New ArrayList()

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            If CheckBox3.Checked = True Then CheckBox3.Checked = False
            TextBox20.Enabled = True
            TextBox21.Enabled = True
            TextBox22.Enabled = True
            TextBox23.Enabled = True
            TextBox24.Enabled = True
            TextBox25.Enabled = True
            TextBox14.Enabled = False
            TextBox15.Enabled = False
            TextBox16.Enabled = False
            TextBox17.Enabled = False
            TextBox18.Enabled = False
            TextBox19.Enabled = False

        Else

            TextBox20.Enabled = False
            TextBox21.Enabled = False
            TextBox22.Enabled = False
            TextBox23.Enabled = False
            TextBox24.Enabled = False
            TextBox25.Enabled = False
            TextBox14.Enabled = True
            TextBox15.Enabled = True
            TextBox16.Enabled = True
            TextBox17.Enabled = True
            TextBox18.Enabled = True
            TextBox19.Enabled = True
        End If
    End Sub

    Private Sub TextBox26_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub СправкаПоЗарплате_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.ComboBox3.AutoCompleteCustomSource.Clear()
        Me.ComboBox3.Items.Clear()
        For i As Integer = 1 To 12
            Me.ComboBox3.AutoCompleteCustomSource.Add(MonthName(i).ToString())
            Me.ComboBox3.Items.Add(MonthName(i).ToString)
            Me.ComboBox5.Items.Add((i).ToString)
        Next
        Dim ut() As Object = {Now.Year - 2, Now.Year - 1, Now.Year}

        Me.ComboBox4.Items.Clear()
        Me.ComboBox9.Items.Clear()
        ComboBox4.Items.AddRange(ut)
        ComboBox9.Items.AddRange(ut)


        Me.ComboBox1.AutoCompleteCustomSource.Clear()
        Me.ComboBox1.Items.Clear()
        For Each r As DataRow In СписокКлиентовОсновной.Rows
            Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox1.Items.Add(r(0).ToString)
        Next

        MaskedTextBox1.Text = Now.ToShortDateString
        TextBox37.Text = "0,00"
        TextBox36.Text = "0,00"
        TextBox35.Text = "0,00"
        TextBox34.Text = "0,00"
        TextBox33.Text = "0,00"
        TextBox32.Text = "0,00"
        TextBox58.Text = "0"
        TextBox59.Text = "0"
        'CheckBox4.Checked = True
        Label20.Visible = False

        'Dim StrSql3 As String = "SELECT * FROM КонстантаПоВычетамЗп"
        'Dim ds3 As DataTable = Selects(StrSql3)

        'МаксДох = CDbl(ds3.Rows(0).Item(1).ToString)
        'СумВыч = CDbl(ds3.Rows(0).Item(2).ToString)
        'ОдинРеб = CDbl(ds3.Rows(0).Item(3).ToString)
        'ДваРеб = CDbl(ds3.Rows(0).Item(4).ToString)
        ComboBox10.Items.Clear()
        ComboBox7.Text = 0
        r1 = {3, 4, 5, 6}
        ComboBox10.Items.AddRange(r1)

    End Sub

    Private Sub TextBox26_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True

            TextBox6.Text = MonthName(Month(FormatDateTime(Now)))
        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

        Label10.Text = ComboBox5.Items.Item(ComboBox3.SelectedIndex)
        Dim df As Integer = CType(Label10.Text, Integer)

        TextBox6.Text = ComboBox3.SelectedItem

        If df - 1 = 0 Then
            df = 12
            TextBox7.Text = MonthName(df)

        Else
            TextBox7.Text = MonthName(df - 1)
            df = df - 1
        End If

        If df - 1 = 0 Then
            df = 12
            TextBox4.Text = MonthName(df)

        Else
            TextBox4.Text = MonthName(df - 1)
            df = df - 1
        End If

        If df - 1 = 0 Then
            df = 12
            TextBox5.Text = MonthName(df)

        Else
            TextBox5.Text = MonthName(df - 1)
            df = df - 1
        End If

        If df - 1 = 0 Then
            df = 12
            TextBox3.Text = MonthName(df)

        Else
            TextBox3.Text = MonthName(df - 1)
            df = df - 1
        End If

        If df - 1 = 0 Then
            df = 12
            TextBox2.Text = MonthName(df)

        Else
            TextBox2.Text = MonthName(df - 1)
            df = df - 1
        End If

        Dim mint As Integer = Now.Month
        If mint >= CType(Label10.Text, Integer) And CheckBox2.Checked = False Then
            TextBox12.Text = Now.Year
        Else
            Dim d As Date = Now
            d = d.AddYears(-1)
            TextBox12.Text = d.Year
        End If

        If CheckBox2.Checked = False And ComboBox4.Text = "" Then

            If CType(Label10.Text, Integer) - 1 = 0 Then
                TextBox13.Text = TextBox12.Text - 1
            Else
                TextBox13.Text = TextBox12.Text
            End If

            If CType(Label10.Text, Integer) - 2 = 0 Then
                TextBox10.Text = TextBox13.Text - 1
            Else
                TextBox10.Text = TextBox13.Text
            End If

            If CType(Label10.Text, Integer) - 3 = 0 Then
                TextBox11.Text = TextBox10.Text - 1
            Else
                TextBox11.Text = TextBox10.Text
            End If

            If CType(Label10.Text, Integer) - 4 = 0 Then
                TextBox9.Text = TextBox11.Text - 1
            Else
                TextBox9.Text = TextBox11.Text
            End If

            If CType(Label10.Text, Integer) - 5 = 0 Then
                TextBox8.Text = TextBox9.Text - 1
            Else
                TextBox8.Text = TextBox9.Text
            End If



        End If




        'TextBox5.Text = ComboBox3.SelectedItem
        'TextBox3.Text = ComboBox3.SelectedItem
        'TextBox2.Text = ComboBox3.SelectedItem
        'TextBox2.Text = Month(MonthName(ComboBox3.SelectedItem))

        'Dim Currentmonth As MonthCalendar
        'Currentmonth = Month(Now.Month)
        'Dim CurrentDate As Date
        'CurrentDate = DateAdd("m", M, #1/1/1990#) 'добавили M месяцев

        'MsgBox CurrentDate                      'показ получившейся даты
        'MsgBox DatePart("yyyy", CurrentDate)    'показ её года
        'MsgBox MonthName(Month(CurrentDate))    'показ названия её месяца (что мы и ищем)
        'MsgBox DatePart("d", CurrentDate) 







    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            ComboBox4.Enabled = True

        Else
            ComboBox4.Enabled = False
            ComboBox4.Text = ""
        End If
    End Sub

    'Private Sub TextBox14_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox14.KeyDown

    '    If e.KeyCode = Keys.Enter Then
    '        e.SuppressKeyPress = True
    '        If TextBox14.Text = "" Then Exit Sub
    '        TextBox14.Text = Replace(TextBox14.Text, ".", ",")
    '        'Dim dd As Double = CDbl(TextBox14.Text)
    '        'dd = dd * 13 / 100
    '        'If CheckBox1.Checked = False Then
    '        '    TextBox25.Text = Math.Round(dd, 2)

    '        'End If


    '        TextBox25.Text = Math.Round(CDbl(TextBox14.Text) * 13 / 100, 2)
    '        TextBox31.Text = Math.Round(CDbl(TextBox14.Text) * 1 / 100, 2)
    '        TextBox43.Text = Math.Round(CDbl(TextBox31.Text) + CDbl(TextBox37.Text) + CDbl(TextBox25.Text), 2)
    '        TextBox49.Text = Math.Round(CDbl(TextBox14.Text) - CDbl(TextBox43.Text), 2)
    '        Try
    '            TextBox50.Text = Math.Round(CDbl(TextBox14.Text) + CDbl(TextBox15.Text) + CDbl(TextBox16.Text) + CDbl(TextBox17.Text) + CDbl(TextBox18.Text) + CDbl(TextBox19.Text), 2)
    '        Catch ex As Exception

    '        End Try
    '        gd = Math.Round(CDbl(TextBox14.Text), 2)
    '        TextBox50.Text = gd
    '        подох = Math.Round(CDbl(TextBox25.Text), 2)
    '        TextBox51.Text = подох
    '        фсзн = Math.Round(CDbl(TextBox31.Text), 2)
    '        TextBox52.Text = фсзн
    '        прочие = Math.Round(CDbl(TextBox37.Text), 2)
    '        TextBox53.Text = прочие
    '        итогоудерж = Math.Round(CDbl(TextBox43.Text), 2)
    '        TextBox54.Text = итогоудерж
    '        квыдаче = Math.Round(CDbl(TextBox49.Text), 2)
    '        TextBox55.Text = квыдаче







    '        TextBox15.Focus()
    '    End If
    'End Sub

    Private Sub TextBox15_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox15.KeyDown
        Dim a, b, c, d, f As Double
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            If TextBox9.Text = "" Or ComboBox10.Text = "" Then
                MessageBox.Show("Выберите из списка последний отчетный месяц или период работы обязательно!", Рик)
                Exit Sub
            End If

            If TextBox15.Text = "" Then
                TextBox24.Text = ""
                TextBox30.Text = ""
                TextBox42.Text = ""
                TextBox48.Text = ""
                TextBox36.Text = "0,00"
                Exit Sub
            Else
                f = Replace(TextBox15.Text, ".", ",")
                Вычеты(TextBox9.Text)
                If errds = 1 Then Exit Sub
                a = Math.Round(ИзменБазыВычетов(f) * 13 / 100, 2) 'TextBox24.Text
                b = Math.Round(f * 1 / 100, 2) 'TextBox30.Text
                c = Math.Round(b + CDbl(TextBox36.Text) + a, 2) 'TextBox42.Text
                d = Math.Round(f - c, 2) 'TextBox48.Text

                If bool(a) = True Then
                    TextBox24.Text = a & ",00"
                Else
                    TextBox24.Text = a
                    If Count(a) = 1 Then
                        TextBox24.Text = a & "0"
                    End If
                End If

                If bool(b) = True Then
                    TextBox30.Text = b & ",00"
                Else
                    TextBox30.Text = b
                    If Count(b) = 1 Then
                        TextBox30.Text = b & "0"
                    End If
                End If

                If bool(c) = True Then
                    TextBox42.Text = c & ",00"
                Else
                    TextBox42.Text = c
                    If Count(c) = 1 Then
                        TextBox42.Text = c & "0"
                    End If
                End If

                If bool(d) = True Then
                    TextBox48.Text = d & ",00"
                Else
                    TextBox48.Text = d
                    If Count(d) = 1 Then
                        TextBox48.Text = d & "0"
                    End If
                End If
                If bool(f) = True Then
                    TextBox15.Text = f & ",00"
                Else
                    TextBox15.Text = f
                    If Count(f) = 1 Then
                        TextBox15.Text = f & "0"
                    End If
                End If
                пров()
            End If
            TextBox17.Focus()
        End If
    End Sub

    Private Sub TextBox17_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox17.KeyDown
        Dim a, b, c, d, f As Double
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True

            If TextBox11.Text = "" Or ComboBox10.Text = "" Then
                MessageBox.Show("Выберите из списка последний отчетный месяц или период работы обязательно!", Рик)
                Exit Sub
            End If

            If TextBox17.Text = "" Then

                TextBox23.Text = ""
                TextBox29.Text = ""
                TextBox35.Text = "0,00"
                TextBox41.Text = ""
                TextBox47.Text = ""
                Exit Sub
            Else
                f = Replace(TextBox17.Text, ".", ",") '17
                Вычеты(TextBox11.Text)
                If errds = 1 Then Exit Sub
                a = Math.Round(ИзменБазыВычетов(f) * 13 / 100, 2) 'TextBox23
                b = Math.Round(f * 1 / 100, 2) 'TextBox29
                c = Math.Round(b + CDbl(TextBox35.Text) + a, 2) 'TextBox41
                d = Math.Round(f - c, 2) 'TextBox47

                If bool(a) = True Then
                    TextBox23.Text = a & ",00"
                Else
                    TextBox23.Text = a
                    If Count(a) = 1 Then
                        TextBox23.Text = a & "0"
                    End If
                End If

                If bool(b) = True Then
                    TextBox29.Text = b & ",00"
                Else
                    TextBox29.Text = b
                    If Count(b) = 1 Then
                        TextBox29.Text = b & "0"
                    End If
                End If

                If bool(c) = True Then
                    TextBox41.Text = c & ",00"
                Else
                    TextBox41.Text = c
                    If Count(c) = 1 Then
                        TextBox41.Text = c & "0"
                    End If
                End If

                If bool(d) = True Then
                    TextBox47.Text = d & ",00"
                Else
                    TextBox47.Text = d
                    If Count(d) = 1 Then
                        TextBox47.Text = d & "0"
                    End If
                End If

                If bool(f) = True Then
                    TextBox17.Text = f & ",00"
                Else
                    TextBox17.Text = f
                    If Count(f) = 1 Then
                        TextBox17.Text = f & "0"
                    End If
                End If
                пров()
            End If
            TextBox16.Focus()
        End If
    End Sub

    Private Sub TextBox16_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox16.KeyDown
        Dim a, b, c, d, f As Double
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True

            If TextBox10.Text = "" Or ComboBox10.Text = "" Then
                MessageBox.Show("Выберите из списка последний отчетный месяц или период работы обязательно!", Рик)
                Exit Sub
            End If




            If TextBox16.Text = "" Then
                TextBox22.Text = ""
                TextBox28.Text = ""
                TextBox34.Text = "0,00"
                TextBox40.Text = ""
                TextBox46.Text = ""
                Exit Sub
            Else
                f = Replace(TextBox16.Text, ".", ",") '16
                Вычеты(TextBox10.Text)
                If errds = 1 Then Exit Sub
                a = Math.Round(ИзменБазыВычетов(f) * 13 / 100, 2) 'TextBox22
                b = Math.Round(f * 1 / 100, 2) 'TextBox28
                c = Math.Round(b + CDbl(TextBox34.Text) + a, 2) 'TextBox40
                d = Math.Round(f - c, 2) 'TextBox46

                If bool(a) = True Then
                    TextBox22.Text = a & ",00"
                Else
                    TextBox22.Text = a
                    If Count(a) = 1 Then
                        TextBox22.Text = a & "0"
                    End If
                End If

                If bool(b) = True Then
                    TextBox28.Text = b & ",00"
                Else
                    TextBox28.Text = b
                    If Count(b) = 1 Then
                        TextBox28.Text = b & "0"
                    End If
                End If

                If bool(c) = True Then
                    TextBox40.Text = c & ",00"
                Else
                    TextBox40.Text = c
                    If Count(c) = 1 Then
                        TextBox40.Text = c & "0"
                    End If
                End If

                If bool(d) = True Then
                    TextBox46.Text = d & ",00"
                Else
                    TextBox46.Text = d
                    If Count(d) = 1 Then
                        TextBox46.Text = d & "0"
                    End If
                End If

                If bool(f) = True Then
                    TextBox16.Text = f & ",00"
                Else
                    TextBox16.Text = f
                    If Count(f) = 1 Then
                        TextBox16.Text = f & "0"
                    End If
                End If
                пров()
            End If
            TextBox19.Focus()
        End If
    End Sub

    Private Sub TextBox19_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox19.KeyDown
        Dim a, b, c, d, f As Double
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True

            If TextBox13.Text = "" Or ComboBox10.Text = "" Then
                MessageBox.Show("Выберите из списка последний отчетный месяц или период работы обязательно!", Рик)
                Exit Sub
            End If
            If TextBox19.Text = "" Then
                TextBox21.Text = ""
                TextBox27.Text = ""
                TextBox33.Text = "0,00"
                TextBox39.Text = ""
                TextBox45.Text = ""
                Exit Sub
            Else
                f = Replace(TextBox19.Text, ".", ",") '19
                Вычеты(TextBox13.Text)
                If errds = 1 Then Exit Sub
                a = Math.Round(ИзменБазыВычетов(f) * 13 / 100, 2) 'TextBox21
                b = Math.Round(f * 1 / 100, 2) 'TextBox27
                c = Math.Round(b + CDbl(TextBox33.Text) + a, 2) 'TextBox39
                d = Math.Round(f - c, 2) 'TextBox45

                If bool(a) = True Then
                    TextBox21.Text = a & ",00"
                Else
                    TextBox21.Text = a
                    If Count(a) = 1 Then
                        TextBox21.Text = a & "0"
                    End If
                End If

                If bool(b) = True Then
                    TextBox27.Text = b & ",00"
                Else
                    TextBox27.Text = b
                    If Count(b) = 1 Then
                        TextBox27.Text = b & "0"
                    End If
                End If

                If bool(c) = True Then
                    TextBox39.Text = c & ",00"
                Else
                    TextBox39.Text = c
                    If Count(c) = 1 Then
                        TextBox39.Text = c & "0"
                    End If
                End If

                If bool(d) = True Then
                    TextBox45.Text = d & ",00"
                Else
                    TextBox45.Text = d
                    If Count(d) = 1 Then
                        TextBox45.Text = d & "0"
                    End If
                End If
                If bool(f) = True Then
                    TextBox19.Text = f & ",00"
                Else
                    TextBox19.Text = f
                    If Count(f) = 1 Then
                        TextBox19.Text = f & "0"
                    End If
                End If
                пров()
            End If
            TextBox18.Focus()
        End If
    End Sub


    Private Sub TextBox18_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox18.KeyDown
        Dim a, b, c, d, f As Double
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            If TextBox12.Text = "" Or ComboBox10.Text = "" Then
                MessageBox.Show("Выберите из списка последний отчетный месяц или период работы обязательно!", Рик)
                Exit Sub
            End If
            If TextBox18.Text = "" Then
                TextBox20.Text = ""
                TextBox26.Text = ""
                TextBox32.Text = "0,00"
                TextBox38.Text = ""
                TextBox44.Text = ""

                Exit Sub
            Else
                f = Replace(TextBox18.Text, ".", ",") '18
                Вычеты(TextBox12.Text)
                If errds = 1 Then Exit Sub
                a = Math.Round(ИзменБазыВычетов(f) * 13 / 100, 2) 'TextBox20
                b = Math.Round(f * 1 / 100, 2) 'TextBox26
                c = Math.Round(b + CDbl(TextBox32.Text) + a, 2) 'TextBox38
                d = Math.Round(f - c, 2) 'TextBox44

                If bool(a) = True Then
                    TextBox20.Text = a & ",00"
                Else
                    TextBox20.Text = a
                    If Count(a) = 1 Then
                        TextBox20.Text = a & "0"
                    End If
                End If

                If bool(b) = True Then
                    TextBox26.Text = b & ",00"
                Else
                    TextBox26.Text = b
                    If Count(b) = 1 Then
                        TextBox26.Text = b & "0"
                    End If
                End If

                If bool(c) = True Then
                    TextBox38.Text = c & ",00"
                Else
                    TextBox38.Text = c
                    If Count(c) = 1 Then
                        TextBox38.Text = c & "0"
                    End If
                End If

                If bool(d) = True Then
                    TextBox44.Text = d & ",00"
                Else
                    TextBox44.Text = d
                    If Count(d) = 1 Then
                        TextBox44.Text = d & "0"
                    End If
                End If
                If bool(f) = True Then
                    TextBox18.Text = f & ",00"
                Else
                    TextBox18.Text = f
                    If Count(f) = 1 Then
                        TextBox18.Text = f & "0"
                    End If
                End If
                пров()
                TextBox58.Focus()
            End If
        End If
    End Sub

    'Private Sub TextBox14_LostFocus(sender As Object, e As EventArgs) Handles TextBox14.LostFocus
    '    'If TextBox14.Text <> "" Then
    '    '    TextBox14.Text = Replace(TextBox14.Text, ".", ",")
    '    '    Dim dd As Double = CDbl(TextBox14.Text)
    '    '    dd = dd * 13 / 100
    '    '    If CheckBox1.Checked = False Then
    '    '        TextBox25.Text = Math.Round(dd, 2)
    '    '    End If
    '    'End If
    'End Sub

    'Private Sub TextBox15_LostFocus(sender As Object, e As EventArgs) Handles TextBox15.LostFocus
    '    If TextBox15.Text <> "" Then
    '        TextBox15.Text = Replace(TextBox15.Text, ".", ",")
    '        Dim dd As Double = CDbl(TextBox15.Text)
    '        dd = dd * 13 / 100
    '        If CheckBox1.Checked = False Then
    '            TextBox24.Text = Math.Round(dd, 2)
    '        End If
    '    End If
    'End Sub

    'Private Sub TextBox17_LostFocus(sender As Object, e As EventArgs) Handles TextBox17.LostFocus
    '    If TextBox17.Text <> "" Then
    '        TextBox17.Text = Replace(TextBox17.Text, ".", ",")
    '        Dim dd As Double = CDbl(TextBox17.Text)
    '        dd = dd * 13 / 100
    '        If CheckBox1.Checked = False Then
    '            TextBox23.Text = Math.Round(dd, 2)
    '        End If
    '    End If
    'End Sub

    Private Sub TextBox51_TextChanged(sender As Object, e As EventArgs) Handles TextBox51.TextChanged
        Try
            Dim a As Double
            Select Case CType(ComboBox10.Text, Integer)
                Case 3
                    a = Math.Round((CDbl(TextBox51.Text) / 3), 2)
                Case 4
                    a = Math.Round((CDbl(TextBox51.Text) / 4), 2)
                Case 5
                    a = Math.Round((CDbl(TextBox51.Text) / 5), 2)
                Case 6
                    a = Math.Round((CDbl(TextBox51.Text) / 6), 2)

            End Select

            If bool(a) = True Then
                TextBox56.Text = a & ",00"
            Else
                TextBox56.Text = a
                If Count(a) = 1 Then
                    TextBox56.Text = a & "0"
                End If
            End If
        Catch ex As Exception
            TextBox56.Text = ""
        End Try


    End Sub

    Private Sub TextBox52_TextChanged(sender As Object, e As EventArgs) Handles TextBox52.TextChanged
        Try
            Dim a As Double
            Select Case CType(ComboBox10.Text, Integer)
                Case 3
                    a = Math.Round((CDbl(TextBox52.Text) / 3), 2)
                Case 4
                    a = Math.Round((CDbl(TextBox52.Text) / 4), 2)
                Case 5
                    a = Math.Round((CDbl(TextBox52.Text) / 5), 2)
                Case 6
                    a = Math.Round((CDbl(TextBox52.Text) / 6), 2)

            End Select

            If bool(a) = True Then
                TextBox57.Text = a & ",00"
            Else
                TextBox57.Text = a
                If Count(a) = 1 Then
                    TextBox57.Text = a & "0"
                End If
            End If
        Catch ex As Exception
            TextBox57.Text = ""
        End Try

    End Sub

    Private Sub Вычеты(ByVal год As String)
        errds = 0
        Dim strsql As String = "SELECT * FROM КонстантаПоВычетамЗп WHERE год=" & CType(год, Integer) & ""
        Try
            dx.Clear()
            dx = Selects(strsql)
        Catch ex As Exception
            dx = Selects(strsql)
        End Try
        If errds = 1 Then
            MessageBox.Show("За " & год & " год не внесены предельные значения по вычетам!", Рик)
            Exit Sub
        End If
        МаксДох = Nothing
        СумВыч = Nothing
        ОдинРеб = Nothing
        ДваРеб = Nothing

        МаксДох = CDbl(dx.Rows(0).Item(1).ToString)
        СумВыч = CDbl(dx.Rows(0).Item(2).ToString)
        ОдинРеб = CDbl(dx.Rows(0).Item(3).ToString)
        ДваРеб = CDbl(dx.Rows(0).Item(4).ToString)



    End Sub


    Private Function ИзменБазыВычетов(ByVal f As Double) As Double

        If Label20.Visible = True Then
            Return f
            Exit Function
        End If

        If f <= МаксДох And ComboBox7.Text = "0" Then
            f = f - СумВыч
            Return f
        ElseIf f <= МаксДох And Not ComboBox7.Text = "0" Then
            f = f - СумВыч
            Select Case ComboBox7.Text
                Case "1"
                    f = f - ОдинРеб
                    Return f
                Case "2"
                    f = f - ДваРеб * 2
                    Return f
                Case "3"
                    f = f - ДваРеб * 3
                    Return f
                Case "4"
                    f = f - ДваРеб * 4
                    Return f
                Case "5"
                    f = f - ДваРеб * 5
                    Return f
                Case "6"
                    f = f - ДваРеб * 6
                    Return f
                Case "7"
                    f = f - ДваРеб * 7
                    Return f
                Case "8"
                    f = f - ДваРеб * 8
                    Return f
                Case "9"
                    f = f - ДваРеб * 9
                    Return f
                Case "10"
                    f = f - ДваРеб * 10
                    Return f
            End Select
        End If

        If f > МаксДох And Not ComboBox7.Text = "0" Then

            Select Case ComboBox7.Text
                Case "1"
                    f = f - ОдинРеб
                    Return f
                Case "2"
                    f = f - ДваРеб * 2
                    Return f
                Case "3"
                    f = f - ДваРеб * 3
                    Return f
                Case "4"
                    f = f - ДваРеб * 4
                    Return f
                Case "5"
                    f = f - ДваРеб * 5
                    Return f
                Case "6"
                    f = f - ДваРеб * 6
                    Return f
                Case "7"
                    f = f - ДваРеб * 7
                    Return f
                Case "8"
                    f = f - ДваРеб * 8
                    Return f
                Case "9"
                    f = f - ДваРеб * 9
                    Return f
                Case "10"
                    f = f - ДваРеб * 10
                    Return f
            End Select
        End If
        Return f
    End Function

    'Private Sub TextBox16_LostFocus(sender As Object, e As EventArgs) Handles TextBox16.LostFocus
    '    If TextBox16.Text <> "" Then
    '        TextBox16.Text = Replace(TextBox16.Text, ".", ",")
    '        Dim dd As Double = CDbl(TextBox16.Text)
    '        dd = dd * 13 / 100
    '        If CheckBox1.Checked = False Then
    '            TextBox22.Text = Math.Round(dd, 2)
    '        End If
    '    End If
    'End Sub

    'Private Sub TextBox19_LostFocus(sender As Object, e As EventArgs) Handles TextBox19.LostFocus
    '    If TextBox19.Text <> "" Then
    '        TextBox19.Text = Replace(TextBox19.Text, ".", ",")
    '        Dim dd As Double = CDbl(TextBox19.Text)
    '        dd = dd * 13 / 100
    '        If CheckBox1.Checked = False Then
    '            TextBox21.Text = Math.Round(dd, 2)
    '        End If
    '    End If
    'End Sub

    'Private Sub TextBox18_LostFocus(sender As Object, e As EventArgs) Handles TextBox18.LostFocus
    '    If TextBox18.Text <> "" Then
    '        TextBox18.Text = Replace(TextBox18.Text, ".", ",")
    '        Dim dd As Double = CDbl(TextBox18.Text)
    '        dd = dd * 13 / 100
    '        If CheckBox1.Checked = False Then
    '            TextBox20.Text = Math.Round(dd, 2)
    '        End If
    '    End If
    'End Sub

    Private Sub TextBox14_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox14.KeyDown
        Dim a, b, c, d, f As Double
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True

            If TextBox8.Text = "" Or ComboBox10.Text = "" Then
                MessageBox.Show("Выберите из списка последний отчетный месяц или период работы обязательно!", Рик)
                TextBox14.Text = ""
                Exit Sub
            End If




            If TextBox14.Text = "" Then
                TextBox25.Text = ""
                TextBox31.Text = ""
                TextBox37.Text = "0,00"
                TextBox43.Text = ""
                TextBox49.Text = ""
                Exit Sub
            Else
                f = Replace(TextBox14.Text, ".", ",") '14
                TextBox14.SelectionStart = TextBox14.Text.Length

                Вычеты(TextBox8.Text)
                If errds = 1 Then Exit Sub
                a = Math.Round(ИзменБазыВычетов(f) * 13 / 100, 2) 'txt25
                b = Math.Round(f * 1 / 100, 2) 'txt31
                c = Math.Round(b + CDbl(TextBox37.Text) + a, 2) 'txt43
                d = Math.Round(f - c, 2) 'txt49


                If bool(a) = True Then
                    TextBox25.Text = a & ",00"
                Else
                    TextBox25.Text = a
                    If Count(a) = 1 Then
                        TextBox25.Text = a & "0"
                    End If
                End If

                If bool(b) = True Then
                    TextBox31.Text = b & ",00"
                Else
                    TextBox31.Text = b
                    If Count(b) = 1 Then
                        TextBox31.Text = b & "0"
                    End If
                End If

                If bool(c) = True Then
                    TextBox43.Text = c & ",00"
                Else
                    TextBox43.Text = c
                    If Count(c) = 1 Then
                        TextBox43.Text = c & "0"
                    End If
                End If

                If bool(d) = True Then
                    TextBox49.Text = d & ",00"
                Else
                    TextBox49.Text = d
                    If Count(d) = 1 Then
                        TextBox49.Text = d & "0"
                    End If
                End If

                If bool(f) = True Then
                    TextBox14.Text = f & ",00"
                Else
                    TextBox14.Text = f
                    If Count(f) = 1 Then
                        TextBox14.Text = f & "0"
                    End If
                End If





                пров()
            End If
            TextBox15.Focus()
        End If
    End Sub
    Function bool(ByVal dm As Double)
        If dm = Int(dm) Then
            Return True
        End If
        Return False
    End Function
    Private Sub очистка()
        TextBox14.Text = ""
        TextBox15.Text = ""
        TextBox16.Text = ""
        TextBox17.Text = ""
        TextBox18.Text = ""
        TextBox19.Text = ""
        TextBox20.Text = ""
        TextBox21.Text = ""
        TextBox22.Text = ""
        TextBox23.Text = ""
        TextBox24.Text = ""
        TextBox25.Text = ""
        TextBox26.Text = ""
        TextBox27.Text = ""
        TextBox28.Text = ""
        TextBox29.Text = ""
        TextBox30.Text = ""
        TextBox31.Text = ""
        TextBox32.Text = "0,00"
        TextBox33.Text = "0,00"
        TextBox34.Text = "0,00"
        TextBox35.Text = "0,00"
        TextBox36.Text = "0,00"
        TextBox37.Text = "0,00"
        TextBox38.Text = ""
        TextBox39.Text = ""
        TextBox40.Text = ""
        TextBox41.Text = ""
        TextBox42.Text = ""
        TextBox43.Text = ""
        TextBox44.Text = ""
        TextBox45.Text = ""
        TextBox46.Text = ""
        TextBox47.Text = ""
        TextBox48.Text = ""
        TextBox49.Text = ""
        TextBox50.Text = ""
        TextBox51.Text = ""
        TextBox52.Text = ""
        TextBox53.Text = ""
        TextBox54.Text = ""
        TextBox55.Text = ""
        TextBox56.Text = ""
        TextBox57.Text = ""
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        очистка()
    End Sub


    Private Sub мес3()

        If TextBox18.Text <> "" And TextBox16.Text <> "" And TextBox19.Text <> "" Then
            Dim a As Double
            a = Math.Round(CDbl(TextBox18.Text) + CDbl(TextBox19.Text) + CDbl(TextBox16.Text), 2) 'TextBox50.Text
            If bool(a) = True Then
                TextBox50.Text = a & ",00"
            Else
                TextBox50.Text = a
                If Count(a) = 1 Then
                    TextBox50.Text = a & "0"
                End If
            End If
        Else
            TextBox50.Text = ""
        End If

        If TextBox20.Text <> "" And TextBox21.Text <> "" And TextBox22.Text <> "" Then
            Dim a As Double
            a = Math.Round(CDbl(TextBox20.Text) + CDbl(TextBox21.Text) + CDbl(TextBox22.Text), 2) 'TextBox51.Text
            If bool(a) = True Then
                TextBox51.Text = a & ",00"
            Else
                TextBox51.Text = a
                If Count(a) = 1 Then
                    TextBox51.Text = a & "0"
                End If
            End If

        Else
            TextBox51.Text = ""
        End If

        If TextBox26.Text <> "" And TextBox27.Text <> "" And TextBox28.Text <> "" Then
            Dim a As Double
            a = Math.Round(CDbl(TextBox26.Text) + CDbl(TextBox27.Text) + CDbl(TextBox28.Text), 2) 'TextBox52.Text
            If bool(a) = True Then
                TextBox52.Text = a & ",00"
            Else
                TextBox52.Text = a
                If Count(a) = 1 Then
                    TextBox52.Text = a & "0"
                End If
            End If
        Else
            TextBox52.Text = ""
        End If

        If TextBox32.Text <> "" And TextBox33.Text <> "" And TextBox34.Text <> "" Then
            Dim a As Double
            a = Math.Round(CDbl(TextBox32.Text) + CDbl(TextBox33.Text) + CDbl(TextBox34.Text), 2) 'TextBox53.Text

            If bool(a) = True Then
                TextBox53.Text = a & ",00"
            Else
                TextBox53.Text = a
                If Count(a) = 1 Then
                    TextBox53.Text = a & "0"
                End If
            End If

        Else
            TextBox53.Text = ""
        End If

        If TextBox38.Text <> "" And TextBox39.Text <> "" And TextBox40.Text <> "" Then
            Dim a As Double

            a = Math.Round(CDbl(TextBox38.Text) + CDbl(TextBox39.Text) + CDbl(TextBox40.Text), 2) 'TextBox54.Text
            If bool(a) = True Then
                TextBox54.Text = a & ",00"
            Else
                TextBox54.Text = a
                If Count(a) = 1 Then
                    TextBox54.Text = a & "0"
                End If
            End If
        Else
            TextBox54.Text = ""
        End If

        If TextBox44.Text <> "" And TextBox45.Text <> "" And TextBox46.Text <> "" Then
            Dim a As Double
            a = Math.Round(CDbl(TextBox44.Text) + CDbl(TextBox45.Text) + CDbl(TextBox46.Text), 2) 'TextBox55.Text
            If bool(a) = True Then
                TextBox55.Text = a & ",00"
            Else
                TextBox55.Text = a
                If Count(a) = 1 Then
                    TextBox55.Text = a & "0"
                End If
            End If
        Else
            TextBox55.Text = ""
        End If
    End Sub
    Private Sub мес4()
        If TextBox18.Text <> "" And TextBox16.Text <> "" And TextBox19.Text <> "" And TextBox17.Text <> "" Then
            Dim a As Double
            a = Math.Round(CDbl(TextBox18.Text) + CDbl(TextBox19.Text) + CDbl(TextBox16.Text) + CDbl(TextBox17.Text), 2) 'TextBox50.Text
            If bool(a) = True Then
                TextBox50.Text = a & ",00"
            Else
                TextBox50.Text = a
                If Count(a) = 1 Then
                    TextBox50.Text = a & "0"
                End If
            End If
        Else
            TextBox50.Text = ""
        End If

        If TextBox20.Text <> "" And TextBox21.Text <> "" And TextBox22.Text <> "" And TextBox23.Text <> "" Then
            Dim a As Double
            a = Math.Round(CDbl(TextBox20.Text) + CDbl(TextBox21.Text) + CDbl(TextBox22.Text) + CDbl(TextBox23.Text), 2) 'TextBox51.Text
            If bool(a) = True Then
                TextBox51.Text = a & ",00"
            Else
                TextBox51.Text = a
                If Count(a) = 1 Then
                    TextBox51.Text = a & "0"
                End If
            End If

        Else
            TextBox51.Text = ""
        End If

        If TextBox26.Text <> "" And TextBox27.Text <> "" And TextBox28.Text <> "" And TextBox29.Text <> "" Then
            Dim a As Double
            a = Math.Round(CDbl(TextBox26.Text) + CDbl(TextBox27.Text) + CDbl(TextBox28.Text) + CDbl(TextBox29.Text), 2) 'TextBox52.Text
            If bool(a) = True Then
                TextBox52.Text = a & ",00"
            Else
                TextBox52.Text = a
                If Count(a) = 1 Then
                    TextBox52.Text = a & "0"
                End If
            End If
        Else
            TextBox52.Text = ""
        End If

        If TextBox32.Text <> "" And TextBox33.Text <> "" And TextBox34.Text <> "" And TextBox35.Text <> "" Then
            Dim a As Double
            a = Math.Round(CDbl(TextBox32.Text) + CDbl(TextBox33.Text) + CDbl(TextBox34.Text) + CDbl(TextBox35.Text), 2) 'TextBox53.Text

            If bool(a) = True Then
                TextBox53.Text = a & ",00"
            Else
                TextBox53.Text = a
                If Count(a) = 1 Then
                    TextBox53.Text = a & "0"
                End If
            End If

        Else
            TextBox53.Text = ""
        End If

        If TextBox38.Text <> "" And TextBox39.Text <> "" And TextBox40.Text <> "" And TextBox41.Text <> "" Then
            Dim a As Double

            a = Math.Round(CDbl(TextBox38.Text) + CDbl(TextBox39.Text) + CDbl(TextBox40.Text) + CDbl(TextBox41.Text), 2) 'TextBox54.Text
            If bool(a) = True Then
                TextBox54.Text = a & ",00"
            Else
                TextBox54.Text = a
                If Count(a) = 1 Then
                    TextBox54.Text = a & "0"
                End If
            End If
        Else
            TextBox54.Text = ""
        End If

        If TextBox44.Text <> "" And TextBox45.Text <> "" And TextBox46.Text <> "" And TextBox47.Text <> "" Then
            Dim a As Double
            a = Math.Round(CDbl(TextBox44.Text) + CDbl(TextBox45.Text) + CDbl(TextBox46.Text) + CDbl(TextBox47.Text), 2) 'TextBox55.Text
            If bool(a) = True Then
                TextBox55.Text = a & ",00"
            Else
                TextBox55.Text = a
                If Count(a) = 1 Then
                    TextBox55.Text = a & "0"
                End If
            End If
        Else
            TextBox55.Text = ""
        End If
    End Sub
    Private Sub мес5()
        If TextBox18.Text <> "" And TextBox16.Text <> "" And TextBox19.Text <> "" And TextBox17.Text <> "" And TextBox15.Text <> "" Then
            Dim a As Double
            a = Math.Round(CDbl(TextBox18.Text) + CDbl(TextBox19.Text) + CDbl(TextBox16.Text) + CDbl(TextBox17.Text) + CDbl(TextBox15.Text), 2) 'TextBox50.Text
            If bool(a) = True Then
                TextBox50.Text = a & ",00"
            Else
                TextBox50.Text = a
                If Count(a) = 1 Then
                    TextBox50.Text = a & "0"
                End If
            End If
        Else
            TextBox50.Text = ""
        End If

        If TextBox20.Text <> "" And TextBox21.Text <> "" And TextBox22.Text <> "" And TextBox23.Text <> "" And TextBox24.Text <> "" Then
            Dim a As Double
            a = Math.Round(CDbl(TextBox20.Text) + CDbl(TextBox21.Text) + CDbl(TextBox22.Text) + CDbl(TextBox23.Text) + CDbl(TextBox24.Text), 2) 'TextBox51.Text
            If bool(a) = True Then
                TextBox51.Text = a & ",00"
            Else
                TextBox51.Text = a
                If Count(a) = 1 Then
                    TextBox51.Text = a & "0"
                End If
            End If

        Else
            TextBox51.Text = ""
        End If

        If TextBox26.Text <> "" And TextBox27.Text <> "" And TextBox28.Text <> "" And TextBox29.Text <> "" And TextBox30.Text <> "" Then
            Dim a As Double
            a = Math.Round(CDbl(TextBox26.Text) + CDbl(TextBox27.Text) + CDbl(TextBox28.Text) + CDbl(TextBox29.Text) + CDbl(TextBox30.Text), 2) 'TextBox52.Text
            If bool(a) = True Then
                TextBox52.Text = a & ",00"
            Else
                TextBox52.Text = a
                If Count(a) = 1 Then
                    TextBox52.Text = a & "0"
                End If
            End If
        Else
            TextBox52.Text = ""
        End If

        If TextBox32.Text <> "" And TextBox33.Text <> "" And TextBox34.Text <> "" And TextBox35.Text <> "" And TextBox36.Text <> "" Then
            Dim a As Double
            a = Math.Round(CDbl(TextBox32.Text) + CDbl(TextBox33.Text) + CDbl(TextBox34.Text) + CDbl(TextBox35.Text) + CDbl(TextBox36.Text), 2) 'TextBox53.Text

            If bool(a) = True Then
                TextBox53.Text = a & ",00"
            Else
                TextBox53.Text = a
                If Count(a) = 1 Then
                    TextBox53.Text = a & "0"
                End If
            End If

        Else
            TextBox53.Text = ""
        End If

        If TextBox38.Text <> "" And TextBox39.Text <> "" And TextBox40.Text <> "" And TextBox41.Text <> "" And TextBox42.Text <> "" Then
            Dim a As Double

            a = Math.Round(CDbl(TextBox38.Text) + CDbl(TextBox39.Text) + CDbl(TextBox40.Text) + CDbl(TextBox41.Text) + CDbl(TextBox42.Text), 2) 'TextBox54.Text
            If bool(a) = True Then
                TextBox54.Text = a & ",00"
            Else
                TextBox54.Text = a
                If Count(a) = 1 Then
                    TextBox54.Text = a & "0"
                End If
            End If
        Else
            TextBox54.Text = ""
        End If

        If TextBox44.Text <> "" And TextBox45.Text <> "" And TextBox46.Text <> "" And TextBox47.Text <> "" And TextBox48.Text <> "" Then
            Dim a As Double
            a = Math.Round(CDbl(TextBox44.Text) + CDbl(TextBox45.Text) + CDbl(TextBox46.Text) + CDbl(TextBox47.Text) + CDbl(TextBox48.Text), 2) 'TextBox55.Text
            If bool(a) = True Then
                TextBox55.Text = a & ",00"
            Else
                TextBox55.Text = a
                If Count(a) = 1 Then
                    TextBox55.Text = a & "0"
                End If
            End If
        Else
            TextBox55.Text = ""
        End If
    End Sub
    Private Sub пров()

        Select Case CType(ComboBox10.Text, Integer)
            Case 3
                мес3()
                Exit Sub
            Case 4
                мес4()
                Exit Sub
            Case 5
                мес5()
                Exit Sub
        End Select

        If TextBox18.Text <> "" And TextBox17.Text <> "" And TextBox16.Text <> "" And TextBox15.Text <> "" And TextBox14.Text <> "" And TextBox19.Text <> "" Then
            Dim a As Double
            a = Math.Round(CDbl(TextBox18.Text) + CDbl(TextBox19.Text) + CDbl(TextBox17.Text) + CDbl(TextBox16.Text) + CDbl(TextBox15.Text) + CDbl(TextBox14.Text), 2) 'TextBox50.Text
            If bool(a) = True Then
                TextBox50.Text = a & ",00"
            Else
                TextBox50.Text = a
                If Count(a) = 1 Then
                    TextBox50.Text = a & "0"
                End If
            End If
        Else
            TextBox50.Text = ""
        End If

        If TextBox20.Text <> "" And TextBox21.Text <> "" And TextBox22.Text <> "" And TextBox23.Text <> "" And TextBox24.Text <> "" And TextBox25.Text <> "" Then
            Dim a As Double
            a = Math.Round(CDbl(TextBox20.Text) + CDbl(TextBox21.Text) + CDbl(TextBox22.Text) + CDbl(TextBox23.Text) + CDbl(TextBox24.Text) + CDbl(TextBox25.Text), 2) 'TextBox51.Text
            If bool(a) = True Then
                TextBox51.Text = a & ",00"
            Else
                TextBox51.Text = a
                If Count(a) = 1 Then
                    TextBox51.Text = a & "0"
                End If
            End If

        Else
            TextBox51.Text = ""
        End If

        If TextBox26.Text <> "" And TextBox27.Text <> "" And TextBox28.Text <> "" And TextBox29.Text <> "" And TextBox30.Text <> "" And TextBox31.Text <> "" Then
            Dim a As Double
            a = Math.Round(CDbl(TextBox26.Text) + CDbl(TextBox27.Text) + CDbl(TextBox28.Text) + CDbl(TextBox29.Text) + CDbl(TextBox30.Text) + CDbl(TextBox31.Text), 2) 'TextBox52.Text
            If bool(a) = True Then
                TextBox52.Text = a & ",00"
            Else
                TextBox52.Text = a
                If Count(a) = 1 Then
                    TextBox52.Text = a & "0"
                End If
            End If
        Else
            TextBox52.Text = ""
        End If

        If TextBox32.Text <> "" And TextBox33.Text <> "" And TextBox34.Text <> "" And TextBox35.Text <> "" And TextBox36.Text <> "" And TextBox37.Text <> "" Then
            Dim a As Double
            a = Math.Round(CDbl(TextBox32.Text) + CDbl(TextBox33.Text) + CDbl(TextBox34.Text) + CDbl(TextBox35.Text) + CDbl(TextBox36.Text) + CDbl(TextBox37.Text), 2) 'TextBox53.Text

            If bool(a) = True Then
                TextBox53.Text = a & ",00"
            Else
                TextBox53.Text = a
                If Count(a) = 1 Then
                    TextBox53.Text = a & "0"
                End If
            End If

        Else
            TextBox53.Text = ""
        End If

        If TextBox38.Text <> "" And TextBox39.Text <> "" And TextBox40.Text <> "" And TextBox41.Text <> "" And TextBox42.Text <> "" And TextBox43.Text <> "" Then
            Dim a As Double

            a = Math.Round(CDbl(TextBox38.Text) + CDbl(TextBox39.Text) + CDbl(TextBox40.Text) + CDbl(TextBox41.Text) + CDbl(TextBox42.Text) + CDbl(TextBox43.Text), 2) 'TextBox54.Text
            If bool(a) = True Then
                TextBox54.Text = a & ",00"
            Else
                TextBox54.Text = a
                If Count(a) = 1 Then
                    TextBox54.Text = a & "0"
                End If
            End If
        Else
            TextBox54.Text = ""
        End If

        If TextBox44.Text <> "" And TextBox45.Text <> "" And TextBox46.Text <> "" And TextBox47.Text <> "" And TextBox48.Text <> "" And TextBox49.Text <> "" Then
            Dim a As Double
            a = Math.Round(CDbl(TextBox44.Text) + CDbl(TextBox45.Text) + CDbl(TextBox46.Text) + CDbl(TextBox47.Text) + CDbl(TextBox48.Text) + CDbl(TextBox49.Text), 2) 'TextBox55.Text
            If bool(a) = True Then
                TextBox55.Text = a & ",00"
            Else
                TextBox55.Text = a
                If Count(a) = 1 Then
                    TextBox55.Text = a & "0"
                End If
            End If
        Else
            TextBox55.Text = ""
        End If
    End Sub
    Public Sub фсзнгл(ByVal фсзнР As Double)
        фсзн = фсзнР + фсзн
        TextBox52.Text = фсзн
    End Sub
    Public Sub подохгл(ByVal фсзнР As Double)
        подох = фсзнР + подох
        TextBox51.Text = подох
    End Sub

    Public Sub gdгл(ByVal фсзнР As Double)
        gd = фсзнР + gd
        TextBox50.Text = gd

    End Sub

    Private Sub Полнрасч(ByVal d As Integer)
        Dim v, vu As Double
        Try


            If d = 25 Or d = 37 Then
                v = Math.Round(CDbl(TextBox25.Text) + CDbl(TextBox31.Text) + CDbl(TextBox37.Text), 2)
                If bool(v) = True Then
                    TextBox43.Text = v & ",00"
                Else
                    TextBox43.Text = v
                    If Count(v) = 1 Then
                        TextBox43.Text = v & "0"
                    End If
                End If

                vu = Math.Round(CDbl(TextBox14.Text) - CDbl(TextBox43.Text), 2)
                If bool(vu) = True Then
                    TextBox49.Text = vu & ",00"
                Else
                    TextBox49.Text = vu
                    If Count(vu) = 1 Then
                        TextBox49.Text = vu & "0"
                    End If
                End If
            ElseIf d = 24 Or d = 36 Then
                v = Math.Round(CDbl(TextBox24.Text) + CDbl(TextBox30.Text) + CDbl(TextBox36.Text), 2)
                If bool(v) = True Then
                    TextBox42.Text = v & ",00"
                Else
                    TextBox42.Text = v
                    If Count(v) = 1 Then
                        TextBox42.Text = v & "0"
                    End If
                End If

                vu = Math.Round(CDbl(TextBox15.Text) - CDbl(TextBox42.Text), 2)
                If bool(vu) = True Then
                    TextBox48.Text = vu & ",00"
                Else
                    TextBox48.Text = vu
                    If Count(vu) = 1 Then
                        TextBox48.Text = vu & "0"
                    End If
                End If
            ElseIf d = 23 Or d = 35 Then
                v = Math.Round(CDbl(TextBox23.Text) + CDbl(TextBox29.Text) + CDbl(TextBox35.Text), 2)
                If bool(v) = True Then
                    TextBox41.Text = v & ",00"
                Else
                    TextBox41.Text = v
                    If Count(v) = 1 Then
                        TextBox41.Text = v & "0"
                    End If
                End If

                vu = Math.Round(CDbl(TextBox17.Text) - CDbl(TextBox41.Text), 2)
                If bool(vu) = True Then
                    TextBox47.Text = vu & ",00"
                Else
                    TextBox47.Text = vu
                    If Count(vu) = 1 Then
                        TextBox47.Text = vu & "0"
                    End If
                End If
            ElseIf d = 22 Or d = 34 Then
                v = Math.Round(CDbl(TextBox22.Text) + CDbl(TextBox28.Text) + CDbl(TextBox34.Text), 2)
                If bool(v) = True Then
                    TextBox40.Text = v & ",00"
                Else
                    TextBox40.Text = v
                    If Count(v) = 1 Then
                        TextBox40.Text = v & "0"
                    End If
                End If

                vu = Math.Round(CDbl(TextBox16.Text) - CDbl(TextBox40.Text), 2)
                If bool(vu) = True Then
                    TextBox46.Text = vu & ",00"
                Else
                    TextBox46.Text = vu
                    If Count(vu) = 1 Then
                        TextBox46.Text = vu & "0"
                    End If
                End If
            ElseIf d = 21 Or d = 33 Then
                v = Math.Round(CDbl(TextBox21.Text) + CDbl(TextBox27.Text) + CDbl(TextBox33.Text), 2)
                If bool(v) = True Then
                    TextBox39.Text = v & ",00"
                Else
                    TextBox39.Text = v
                    If Count(v) = 1 Then
                        TextBox39.Text = v & "0"
                    End If
                End If

                vu = Math.Round(CDbl(TextBox19.Text) - CDbl(TextBox39.Text), 2)
                If bool(vu) = True Then
                    TextBox45.Text = vu & ",00"
                Else
                    TextBox45.Text = vu
                    If Count(vu) = 1 Then
                        TextBox45.Text = vu & "0"
                    End If
                End If
            Else
                v = Math.Round(CDbl(TextBox20.Text) + CDbl(TextBox26.Text) + CDbl(TextBox32.Text), 2)
                If bool(v) = True Then
                    TextBox38.Text = v & ",00"
                Else
                    TextBox38.Text = v
                    If Count(v) = 1 Then
                        TextBox38.Text = v & "0"
                    End If
                End If

                vu = Math.Round(CDbl(TextBox18.Text) - CDbl(TextBox38.Text), 2)
                If bool(vu) = True Then
                    TextBox44.Text = vu & ",00"
                Else
                    TextBox44.Text = vu
                    If Count(vu) = 1 Then
                        TextBox44.Text = vu & "0"
                    End If
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub
    Private Sub общсправа(ByVal d As Integer)

        Select Case d
            Case 25
                Полнрасч(25)
            Case 24
                Полнрасч(24)
            Case 23
                Полнрасч(23)
            Case 22
                Полнрасч(22)
            Case 21
                Полнрасч(21)
            Case 20
                Полнрасч(20)
            Case 37
                Полнрасч(37)
            Case 36
                Полнрасч(36)
            Case 35
                Полнрасч(35)
            Case 34
                Полнрасч(34)
            Case 33
                Полнрасч(33)
            Case 32
                Полнрасч(32)
        End Select

    End Sub
    Private Sub TextBox25_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox25.KeyDown

        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Dim a As Double
            a = Replace(TextBox25.Text, ".", ",")
            TextBox25.SelectionStart = TextBox25.Text.Length

            If bool(a) = True Then
                TextBox25.Text = a & ",00"
            Else
                TextBox25.Text = a
                If Count(a) = 1 Then
                    TextBox25.Text = a & "0"
                End If
            End If
            общсправа(25)
            пров()
            TextBox24.Focus()
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged

        If CheckBox3.Checked = True Then
            If CheckBox1.Checked = True Then CheckBox1.Checked = False
            TextBox37.Enabled = True
            TextBox35.Enabled = True
            TextBox34.Enabled = True
            TextBox33.Enabled = True
            TextBox32.Enabled = True
            TextBox36.Enabled = True
            TextBox14.Enabled = False
            TextBox15.Enabled = False
            TextBox16.Enabled = False
            TextBox17.Enabled = False
            TextBox18.Enabled = False
            TextBox19.Enabled = False


        Else

            TextBox37.Enabled = False
            TextBox35.Enabled = False
            TextBox34.Enabled = False
            TextBox33.Enabled = False
            TextBox32.Enabled = False
            TextBox36.Enabled = False
            TextBox14.Enabled = True

            TextBox15.Enabled = True
            TextBox16.Enabled = True
            TextBox17.Enabled = True
            TextBox18.Enabled = True
            TextBox19.Enabled = True
        End If
    End Sub

    Private Sub TextBox24_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox24.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Dim a As Double
            a = Replace(TextBox24.Text, ".", ",")
            TextBox24.SelectionStart = TextBox24.Text.Length
            If bool(a) = True Then
                TextBox24.Text = a & ",00"
            Else
                TextBox24.Text = a
                If Count(a) = 1 Then
                    TextBox24.Text = a & "0"
                End If
            End If
            общсправа(24)
            пров()
            TextBox23.Focus()
        End If
    End Sub

    Private Sub TextBox37_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox37.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Dim a As Double
            a = Replace(TextBox37.Text, ".", ",")
            TextBox37.SelectionStart = TextBox37.Text.Length
            If bool(a) = True Then
                TextBox37.Text = a & ",00"
            Else
                TextBox37.Text = a
                If Count(a) = 1 Then
                    TextBox37.Text = a & "0"
                End If
            End If
            общсправа(37)
            пров()
            TextBox36.Focus()
        End If
    End Sub

    Private Sub TextBox36_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox36.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Dim a As Double
            a = Replace(TextBox36.Text, ".", ",")
            TextBox36.SelectionStart = TextBox36.Text.Length
            If bool(a) = True Then
                TextBox36.Text = a & ",00"
            Else
                TextBox36.Text = a
                If Count(a) = 1 Then
                    TextBox36.Text = a & "0"
                End If
            End If
            общсправа(36)
            пров()
            TextBox35.Focus()
        End If
    End Sub

    Private Sub TextBox34_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox34.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Dim a As Double
            a = Replace(TextBox34.Text, ".", ",")
            TextBox34.SelectionStart = TextBox34.Text.Length
            If bool(a) = True Then
                TextBox34.Text = a & ",00"
            Else
                TextBox34.Text = a
                If Count(a) = 1 Then
                    TextBox34.Text = a & "0"
                End If
            End If
            общсправа(34)
            пров()
            TextBox33.Focus()
        End If
    End Sub


    Private Sub TextBox35_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox35.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Dim a As Double
            a = Replace(TextBox35.Text, ".", ",")
            TextBox35.SelectionStart = TextBox35.Text.Length
            If bool(a) = True Then
                TextBox35.Text = a & ",00"
            Else
                TextBox35.Text = a
                If Count(a) = 1 Then
                    TextBox35.Text = a & "0"
                End If
            End If
            общсправа(35)
            пров()
            TextBox34.Focus()
        End If
    End Sub

    Private Sub TextBox33_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox33.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Dim a As Double
            a = Replace(TextBox33.Text, ".", ",")
            TextBox33.SelectionStart = TextBox33.Text.Length
            If bool(a) = True Then
                TextBox33.Text = a & ",00"
            Else
                TextBox33.Text = a
                If Count(a) = 1 Then
                    TextBox33.Text = a & "0"
                End If
            End If
            общсправа(33)
            пров()
            TextBox32.Focus()
        End If
    End Sub

    Private Sub TextBox32_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox32.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Dim a As Double
            a = Replace(TextBox32.Text, ".", ",")
            TextBox32.SelectionStart = TextBox32.Text.Length
            If bool(a) = True Then
                TextBox32.Text = a & ",00"
            Else
                TextBox32.Text = a
                If Count(a) = 1 Then
                    TextBox32.Text = a & "0"
                End If
            End If
            общсправа(32)
            пров()
            Button1.Focus()
        End If
    End Sub
    Private Sub Com1sel()
        'Dim StrSql As String


        'StrSql = "SELECT ФИОСборное,КодСотрудники FROM Сотрудники WHERE НазвОрганиз='" & ComboBox1.Text & "' ORDER BY ФИОСборное "
        'Dim ds As DataTable = Selects(StrSql)
        Dim ds = From x In dtSotrudnikiAll Where x.Item("НазвОрганиз") = ComboBox1.Text And x.Item("НаличеДогПодряда") = "Нет" Order By x.Item("ФИОСборное") Select x

        Me.ComboBox2.AutoCompleteCustomSource.Clear()
        Me.ComboBox2.Items.Clear()
        ComboBox6.Items.Clear()

        For Each r In ds
            Me.ComboBox2.AutoCompleteCustomSource.Add(r.Item("ФИОСборное").ToString())
            Me.ComboBox2.Items.Add(r.Item("ФИОСборное").ToString())
            'Me.ComboBox19.Items.Add(r(1).ToString)
            Me.ComboBox6.Items.Add(r.Item("КодСотрудники").ToString())
        Next
        ComboBox2.Text = ""
        ComboBox9.Text = Now.Year


        Dim list = listFluentFTP(ComboBox1.Text & "\Справки\" & Now.Year)
        ComboBox8.Items.Clear()
        For Each r In list
            ComboBox8.Items.Add(r.ToString)
        Next


        очистка()
        TextBox1.Text = ""
        MaskedTextBox1.Text = Now.ToShortDateString
        ComboBox7.Text = "0"
        TextBox58.Text = "0, 0"
        TextBox59.Text = "0, 0"
        'CheckBox4.Checked = True
        CheckBox1.Checked = False
        CheckBox3.Checked = False
        ComboBox3.Text = ""
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Com1sel()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Label20.Visible = False
        Label17.Text = ComboBox6.Items.Item(ComboBox2.SelectedIndex)
        Dim strsql As String = "Select ПоСовмест FROM КарточкаСотрудника WHERE IDСотр=" & CType(Label17.Text, Integer) & ""
        Dim d As DataTable = Selects(strsql)
        Try
            If d.Rows(0).Item(0).ToString <> "" Then
                Label20.Visible = True
            End If
        Catch ex As Exception
            MessageBox.Show("С сотрудником заключен договор подряда!", Рик)
            Exit Sub
        End Try
        очистка()
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            MaskedTextBox1.Focus()

            Dim pl As String
            If TextBox1.Text <> "" Then
                Dim i As Integer = CInt(TextBox1.Text)
                Select Case i

                    Case < 10
                        pl = Str(i)
                        TextBox1.Text = "00" & i

                    Case 10 To 99
                        pl = Str(i)
                        TextBox1.Text = "0" & i
                End Select
            End If
        End If
    End Sub

    Private Sub MaskedTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            ComboBox3.Focus()


        End If
    End Sub

    Private Sub ComboBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            ComboBox10.Focus()
        End If
    End Sub
    Private Function Проверка() As Integer

        If ComboBox1.Text = "" Or ComboBox2.Text = "" Then
            MessageBox.Show("Выберите организацию или сотрудника!", Рик)
            Return 1
        End If

        If TextBox1.Text = "" Or MaskedTextBox1.MaskCompleted = False Then
            MessageBox.Show("Введите номер или дату документа!", Рик)
            Return 1
        End If

        If ComboBox3.Text = "" Then
            MessageBox.Show("Выберите последний отчетный месяц!", Рик)
            Return 1
        End If

        If TextBox58.Text = "" Or TextBox59.Text = "" Then
            MessageBox.Show("Заполните поля по исполнительным листам или кредиту!", Рик)
            Return 1
        End If


        Select Case CType(ComboBox10.Text, Integer)

            Case 3
                If TextBox16.Text = "" Or TextBox18.Text = "" Or TextBox19.Text = "" Then
                    MessageBox.Show("Заполните поле начислено", Рик)
                    Return 1
                End If
            Case 4
                If TextBox16.Text = "" Or TextBox17.Text = "" Or TextBox18.Text = "" Or TextBox19.Text = "" Then
                    MessageBox.Show("Заполните поле начислено", Рик)
                    Return 1
                End If
            Case 5
                If TextBox16.Text = "" Or TextBox17.Text = "" Or TextBox18.Text = "" Or TextBox19.Text = "" Or TextBox15.Text = "" Then
                    MessageBox.Show("Заполните поле начислено", Рик)
                    Return 1
                End If
            Case 6
                If TextBox14.Text = "" Or TextBox15.Text = "" Or TextBox16.Text = "" Or TextBox17.Text = "" Or TextBox18.Text = "" Or TextBox19.Text = "" Then
                    MessageBox.Show("Заполните поле начислено", Рик)
                    Return 1
                End If
        End Select

        Return 0
    End Function

    Private Function Contr(ByVal id As Integer) As String
        Dim strsql, strsql2 As String
        Dim ds, ds2 As DataTable
        Dim срокпр As Integer
        Dim датаоконч As String
        Dim стргод As String
        Dim об As String
        strsql = "Select * FROM ПродлКонтракта WHERE IDСотр=" & id & ""
        ds = Selects(strsql)
        strsql2 = "Select ДатаПриема, СрокКонтракта FROM КарточкаСотрудника WHERE IDСотр=" & id & ""
        ds2 = Selects(strsql2)

        Try
            If Not ds.Rows(0).Item(5).ToString <> "" Then
                ds.Rows(0).Item(5) = "0"
            End If
        Catch ex As Exception
            MessageBox.Show("Данный сотрудник не внесен в базу котрактов!" & vbCrLf & "Будет продолжено но некоторые данные необходимо будет внести вручную.", Рик)
            срокпр = CType(ds2.Rows(0).Item(1).ToString, Integer)
            Dim d As Date
            d = CDate(ds2.Rows(0).Item(0).ToString)
            d = d.AddYears(срокпр)
            d = d.AddDays(-1)
            датаоконч = d.ToShortDateString
            стргод = Склонение2(CType(срокпр, String))
            об = срокпр & " " & стргод & " c " & ds2.Rows(0).Item(0).ToString & "г. по " & датаоконч
            Return об
        End Try

        If Not ds.Rows(0).Item(9).ToString <> "" Then
            ds.Rows(0).Item(9) = "0"
        End If

        If Not ds.Rows(0).Item(13).ToString <> "" Then
            ds.Rows(0).Item(13) = "0"
        End If

        If Not ds.Rows(0).Item(17).ToString <> "" Then
            ds.Rows(0).Item(17) = "0"
        End If

        If Not ds.Rows(0).Item(21).ToString <> "" Then
            ds.Rows(0).Item(21) = "0"
        End If

        срокпр = CType(ds.Rows(0).Item(5).ToString, Integer) + CType(ds.Rows(0).Item(9).ToString, Integer) + CType(ds.Rows(0).Item(13).ToString, Integer) + CType(ds.Rows(0).Item(17).ToString, Integer) + CType(ds.Rows(0).Item(21).ToString, Integer)


        If ds.Rows(0).Item(8).ToString = "" Then
            датаоконч = ds.Rows(0).Item(4).ToString
        ElseIf ds.Rows(0).Item(12).ToString = "" Then
            датаоконч = ds.Rows(0).Item(8).ToString
        ElseIf ds.Rows(0).Item(16).ToString = "" Then
            датаоконч = ds.Rows(0).Item(12).ToString
        ElseIf ds.Rows(0).Item(20).ToString = "" Then
            датаоконч = ds.Rows(0).Item(16).ToString
        Else
            датаоконч = ds.Rows(0).Item(20).ToString
        End If

        стргод = Склонение2(CType(срокпр, String))
        об = срокпр & " " & стргод & " c " & ds.Rows(0).Item(3).ToString & "г. по " & датаоконч

        Return об
    End Function
    Private Sub TextBox58_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox58.KeyDown
        Dim a As Double
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            a = CDbl(TextBox58.Text)
            If bool(a) = True Then
                TextBox58.Text = a & ", 0"
            Else
                TextBox25.Text = a
                If Count(a) = 1 Then
                    TextBox58.Text = a & "0"
                End If
            End If
            TextBox59.Focus()
        End If
    End Sub

    Private Sub TextBox59_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox59.KeyDown
        Dim a As Double
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            a = CDbl(TextBox59.Text)
            If bool(a) = True Then
                TextBox59.Text = a & ", 0"
            Else
                TextBox25.Text = a
                If Count(a) = 1 Then
                    TextBox59.Text = a & "0"
                End If
            End If
            Button1.Focus()
        End If
    End Sub
    Private Sub Отбор(ByVal r As Integer)
        Select Case r
            Case 3
                GroupBox2.Visible = False
                GroupBox3.Visible = False
                GroupBox4.Visible = False
                очистка()
            Case 4
                GroupBox2.Visible = False
                GroupBox3.Visible = False
                GroupBox4.Visible = True
                очистка()
            Case 5
                GroupBox2.Visible = False
                GroupBox3.Visible = True
                GroupBox4.Visible = True
                очистка()
            Case 6
                GroupBox2.Visible = True
                GroupBox3.Visible = True
                GroupBox4.Visible = True
                очистка()
        End Select


    End Sub

    Private Sub ComboBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox10.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Select Case ComboBox10.SelectedItem
                Case 3
                    TextBox16.Focus()
                Case 4
                    TextBox17.Focus()
                Case 5
                    TextBox15.Focus()
                Case 6
                    TextBox14.Focus()
            End Select

        End If
    End Sub



    Private Sub ComboBox10_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox10.SelectedIndexChanged
        Отбор(CType(ComboBox10.SelectedItem, Integer))
    End Sub
    Private Sub доки3()

        Me.Cursor = Cursors.WaitCursor
        Dim strsql, strsql1, strsql2, strsql3 As String
        Dim ds1, ds3 As DataTable

        Dim ds = dtClientAll.Select("НазвОрг='" & ComboBox1.Text & "'")

        Dim list As New Dictionary(Of String, Object)
        list.Add("@КодСотрудники", CType(Label17.Text, Integer))



        'strsql = "SELECT * FROM Клиент WHERE НазвОрг='" & ComboBox1.Text & "'" 'Данные клиента
        'ds = Selects(strsql)

        ds1 = Selects(StrSql:="SELECT Сотрудники.ФИОСборное, Сотрудники.ПаспортСерия, Сотрудники.ПаспортНомер, Сотрудники.ПаспортКогдаВыдан,
Сотрудники.ПаспортКемВыдан, Сотрудники.ИДНомер, Сотрудники.Пол, КарточкаСотрудника.ДатаПриема, КарточкаСотрудника.СрокКонтракта, ДогСотрудн.СрокОкончКонтр,
Штатное.Должность, КарточкаСотрудника.ПродлКонтрС, КарточкаСотрудника.ПродлКонтрПо, КарточкаСотрудника.СрокПродлКонтракта, КарточкаСотрудника.АдресОбъектаОбщепита, Штатное.Разряд,
Сотрудники.ФамилияДляЗаявления, Сотрудники.ИмяДляЗаявления, Сотрудники.ОтчествоДляЗаявления
FROM ((Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр) INNER JOIN ДогСотрудн ON Сотрудники.КодСотрудники = ДогСотрудн.IDСотр) INNER JOIN Штатное ON Сотрудники.КодСотрудники = Штатное.ИДСотр
WHERE Сотрудники.КодСотрудники=@КодСотрудники", list)
        'ds1 = Selects(strsql1)

        Dim ds2 = dtObjectObshepitaAll.Select("НазвОрг='" & ds(0).Item(0).ToString & "' AND АдресОбъекта='" & ds1.Rows(0).Item(14).ToString & "'")


        Dim obj As String
        Try
            If Not ds2(0).Item("ТипОбъекта").ToString <> "" And Not ds2(0).Item("НазОбъекта").ToString <> "" Then
                obj = ds1.Rows(0).Item(14).ToString
            ElseIf ds2(0).Item("ТипОбъекта").ToString = "" Then
                obj = ds2(0).Item("НазОбъекта").ToString & " " & ds1.Rows(0).Item(14).ToString
            ElseIf ds2(0).Item("НазОбъекта").ToString = "" Then
                obj = ds2(0).Item("ТипОбъекта").ToString & " " & ds1.Rows(0).Item(14).ToString
            Else
                obj = ds2(0).Item("ТипОбъекта").ToString & " " & ds2(0).Item("НазОбъекта").ToString & " " & ds1.Rows(0).Item(14).ToString
            End If
        Catch ex As Exception
            MessageBox.Show("Не найден адрес объекта общепита, перепроверьте!", Рик)
            Me.Cursor = DefaultCursor
            Exit Sub
        End Try





        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document

        oWord = CreateObject("Word.Application")
        oWord.Visible = False



        Начало("Spravka3mes.doc")
        oWordDoc = oWord.Documents.Add(firthtPath & "\Spravka3mes.doc")

        With oWordDoc.Bookmarks
            .Item("Сп1").Range.Text = ds(0).Item(1).ToString
            If ds(0).Item(1).ToString = "Индивидуальный предприниматель" Then
                .Item("Сп2").Range.Text = ds(0).Item(0).ToString
                .Item("Сп18").Range.Text = "у " & ФормСобствКор(ds(0).Item(1).ToString) & " " & ds(0).Item(0).ToString & " "
            Else
                .Item("Сп2").Range.Text = "«" & ds(0).Item(0).ToString & "»"
                .Item("Сп18").Range.Text = "в " & ФормСобствКор(ds(0).Item(1).ToString) & " «" & ds(0).Item(0).ToString & "» "
            End If
            .Item("Сп3").Range.Text = ds(0).Item(4).ToString 'адрес
            .Item("Сп4").Range.Text = ds(0).Item(2).ToString 'унп
            .Item("Сп5").Range.Text = ds(0).Item(14).ToString 'рс
            .Item("Сп6").Range.Text = ds(0).Item(12).ToString 'банк
            .Item("Сп7").Range.Text = ds(0).Item(11).ToString 'бик
            .Item("Сп8").Range.Text = ds(0).Item(8).ToString 'мыло
            .Item("Сп9").Range.Text = ds(0).Item(6).ToString 'тел
            .Item("Сп10").Range.Text = TextBox1.Text & "-" & Now.Year
            .Item("Сп11").Range.Text = MaskedTextBox1.Text

            Dim inp As String = InputBox("Введите ФИО сотрудника " & vbCrLf & ComboBox2.Text & vbCrLf & " в  Дательном падеже 'Справка выдана Кому?'", Рик)

            Do Until inp <> ""
                MessageBox.Show("Повторите ввод данных!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Error)
                inp = InputBox("Введите ФИО сотрудника " & vbCrLf & ComboBox2.Text & vbCrLf & " в  Дательном падеже 'Справка выдана Кому?'", Рик)
            Loop

            .Item("Сп12").Range.Text = inp
            .Item("Сп13").Range.Text = ds1.Rows(0).Item(1).ToString & ds1.Rows(0).Item(2).ToString
            .Item("Сп14").Range.Text = ds1.Rows(0).Item(4).ToString
            .Item("Сп15").Range.Text = ds1.Rows(0).Item(3).ToString
            .Item("Сп16").Range.Text = ds1.Rows(0).Item(5).ToString 'ид паспорта
            If ds1.Rows(0).Item(6).ToString = "М" Then
                .Item("Сп17").Range.Text = "он"
            Else
                .Item("Сп17").Range.Text = "она"
            End If
            .Item("Сп19").Range.Text = Strings.Left(ds1.Rows(0).Item(7).ToString, 10)

            If ds1.Rows(0).Item(15).ToString <> "" Then
                If ds1.Rows(0).Item(15).ToString = "-" Then
                    .Item("Сп20").Range.Text = Strings.LCase(ds1.Rows(0).Item(10).ToString)
                Else
                    .Item("Сп20").Range.Text = Strings.LCase(ds1.Rows(0).Item(10).ToString) & " " & разрядстрока(CType(ds1.Rows(0).Item(15).ToString, Integer))
                End If

            Else
                .Item("Сп20").Range.Text = Strings.LCase(ds1.Rows(0).Item(10).ToString) 'должность
            End If

            If ФормСобствКор(ds(0).Item(1).ToString) & " " & ds(0).Item(0).ToString = obj Then
                .Item("Сп21").Range.Text = ""
            Else
                .Item("Сп21").Range.Text = obj
            End If

            .Item("Сп22").Range.Text = Contr(CType(Label17.Text, Integer))
            .Item("Сп29").Range.Text = TextBox4.Text
            .Item("Сп30").Range.Text = TextBox7.Text
            .Item("Сп31").Range.Text = TextBox6.Text

            .Item("Сп35").Range.Text = TextBox10.Text
            .Item("Сп36").Range.Text = TextBox13.Text
            .Item("Сп37").Range.Text = TextBox12.Text

            .Item("Сп41").Range.Text = TextBox16.Text
            .Item("Сп42").Range.Text = TextBox19.Text
            .Item("Сп43").Range.Text = TextBox18.Text

            .Item("Сп47").Range.Text = TextBox22.Text
            .Item("Сп48").Range.Text = TextBox21.Text
            .Item("Сп49").Range.Text = TextBox20.Text

            .Item("Сп53").Range.Text = TextBox28.Text
            .Item("Сп54").Range.Text = TextBox27.Text
            .Item("Сп55").Range.Text = TextBox26.Text

            .Item("Сп59").Range.Text = TextBox34.Text
            .Item("Сп60").Range.Text = TextBox33.Text
            .Item("Сп61").Range.Text = TextBox32.Text

            .Item("Сп65").Range.Text = TextBox40.Text
            .Item("Сп66").Range.Text = TextBox39.Text
            .Item("Сп67").Range.Text = TextBox38.Text

            .Item("Сп71").Range.Text = TextBox46.Text
            .Item("Сп72").Range.Text = TextBox45.Text
            .Item("Сп73").Range.Text = TextBox44.Text
            .Item("Сп74").Range.Text = TextBox50.Text
            .Item("Сп75").Range.Text = TextBox51.Text
            .Item("Сп76").Range.Text = TextBox52.Text
            .Item("Сп77").Range.Text = TextBox53.Text
            .Item("Сп78").Range.Text = TextBox54.Text
            .Item("Сп79").Range.Text = TextBox55.Text
            .Item("Сп80").Range.Text = TextBox56.Text
            .Item("Сп81").Range.Text = TextBox57.Text
            .Item("Сп82").Range.Text = TextBox58.Text
            .Item("Сп83").Range.Text = TextBox59.Text
            .Item("Сп84").Range.Text = ds1.Rows(0).Item(16).ToString & " " & Strings.Left(ds1.Rows(0).Item(17).ToString, 1) & "." & Strings.Left(ds1.Rows(0).Item(18).ToString, 1) & "."
            .Item("Сп85").Range.Text = TextBox55.Text
            .Item("Сп86").Range.Text = ЧислоПрописДляСправки(TextBox55.Text)
            If ds(0).Item(18).ToString = "Индивидуальный предприниматель" Then
                .Item("Сп88").Range.Text = ds(0).Item(18).ToString
                .Item("Сп89").Range.Text = ФИОКорРук(ds(0).Item(19).ToString, False)
            Else
                .Item("Сп88").Range.Text = ds(0).Item(18).ToString & " " & ФормСобствКор(ds(0).Item(1).ToString) & " «" & ComboBox1.Text & "» "
                If ds(0).Item(31) = True Then
                    .Item("Сп89").Range.Text = ФИОКорРук(ds(0).Item(19).ToString, True)
                Else
                    .Item("Сп89").Range.Text = ФИОКорРук(ds(0).Item(19).ToString, False)
                End If


            End If

        End With

        Dim Name As String = "Справка №" & TextBox1.Text & "-" & Now.Year & " от " & MaskedTextBox1.Text & " " & ФИОКорРук(ComboBox2.Text, False) & " 3 месяца" & ".doc"
        Dim СохрЗак2 As New List(Of String)
        СохрЗак2.AddRange(New String() {ComboBox1.Text & "\Справки\" & Now.Year & "\", Name})
        oWordDoc.SaveAs2(PathVremyanka & Name,,,,,, False)
        oWordDoc.Close(True)
        oWord.Quit(True)
        Конец(ComboBox1.Text & "\Справки\" & Now.Year, Name, CType(Label17.Text, Integer), ComboBox1.Text, "\Spravka3mes.doc", "Справка по зарплате 3 месяца")
        Dim massFTP As New ArrayList()
        massFTP.Add(СохрЗак2)








        'If Not IO.Directory.Exists(OnePath & ComboBox1.Text & "\Справки\" & Now.Year) Then
        '    IO.Directory.CreateDirectory(OnePath & ComboBox1.Text & "\Справки\" & Now.Year)
        'End If

        'oWordDoc.SaveAs2("C:\Users\Public\Documents\Рик\Справка №" & TextBox1.Text & "-" & Now.Year & " от " & MaskedTextBox1.Text & " " & ФИОКорРук(ComboBox2.Text, False) & " 3 месяца" & ".doc",,,,,, False)

        'Try
        '    IO.File.Copy("C:\Users\Public\Documents\Рик\Справка №" & TextBox1.Text & "-" & Now.Year & " от " & MaskedTextBox1.Text & " " & ФИОКорРук(ComboBox2.Text, False) & " 3 месяца" & ".doc", OnePath & ComboBox1.Text & "\Справки\" & Now.Year & "\Справка №" & TextBox1.Text & "-" & Now.Year & " от " & MaskedTextBox1.Text & " " & ФИОКорРук(ComboBox2.Text, False) & " 3 месяца" & ".doc")
        'Catch ex As Exception
        '    If MessageBox.Show("Справка №" & TextBox1.Text & "-" & Now.Year & " с сотрудником " & ФИОКорРук(ComboBox2.Text, False) & " существует. Заменить старый документ новым?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then
        '        Try
        '            IO.File.Delete(OnePath & ComboBox1.Text & "\Справки\" & Now.Year & "\Справка №" & TextBox1.Text & "-" & Now.Year & " от " & MaskedTextBox1.Text & " " & ФИОКорРук(ComboBox2.Text, False) & " 3 месяца" & ".doc")
        '        Catch ex1 As Exception
        '            MessageBox.Show("Закройте файл!", Рик)
        '        End Try

        '        IO.File.Copy("C:\Users\Public\Documents\Рик\Справка №" & TextBox1.Text & "-" & Now.Year & " от " & MaskedTextBox1.Text & " " & ФИОКорРук(ComboBox2.Text, False) & " 3 месяца" & ".doc", OnePath & ComboBox1.Text & "\Справки\" & Now.Year & "\Справка №" & TextBox1.Text & "-" & Now.Year & " от " & MaskedTextBox1.Text & " " & ФИОКорРук(ComboBox2.Text, False) & " 3 месяца" & ".doc")
        '    End If
        'End Try
        'СохрЗак = OnePath & ComboBox1.Text & "\Справки\" & Now.Year & "\Справка №" & TextBox1.Text & "-" & Now.Year & " от " & MaskedTextBox1.Text & " " & ФИОКорРук(ComboBox2.Text, False) & " 3 месяца" & ".doc"
        'oWordDoc.Close(True)
        'Dim mass() As String
        'СохрЗак = ""
        'mass = {СохрЗак}
        If MessageBox.Show("Справка оформлена!" & vbCrLf & "Распечатать?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            ПечатьДоковFTP(massFTP)
        End If
        Me.Cursor = DefaultCursor
        Me.Close()
    End Sub


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        СправкаЗпКонстанта.ShowDialog()

    End Sub
    Private Sub доки4()
        Me.Cursor = Cursors.WaitCursor

        Dim ds1 As DataTable


        Dim ds = dtClientAll.Select("НазвОрг='" & ComboBox1.Text & "'")

        Dim list As New Dictionary(Of String, Object)
        list.Add("@КодСотрудники", CType(Label17.Text, Integer))

        'strsql = "SELECT * FROM Клиент WHERE НазвОрг='" & ComboBox1.Text & "'" 'Данные клиента
        'ds = Selects(strsql)

        ds1 = Selects(StrSql:= "SELECT Сотрудники.ФИОСборное, Сотрудники.ПаспортСерия, Сотрудники.ПаспортНомер, Сотрудники.ПаспортКогдаВыдан,
Сотрудники.ПаспортКемВыдан, Сотрудники.ИДНомер, Сотрудники.Пол, КарточкаСотрудника.ДатаПриема, КарточкаСотрудника.СрокКонтракта, ДогСотрудн.СрокОкончКонтр,
Штатное.Должность, КарточкаСотрудника.ПродлКонтрС, КарточкаСотрудника.ПродлКонтрПо, КарточкаСотрудника.СрокПродлКонтракта, КарточкаСотрудника.АдресОбъектаОбщепита, Штатное.Разряд,
Сотрудники.ФамилияДляЗаявления, Сотрудники.ИмяДляЗаявления, Сотрудники.ОтчествоДляЗаявления
FROM ((Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр) INNER JOIN ДогСотрудн ON Сотрудники.КодСотрудники = ДогСотрудн.IDСотр) INNER JOIN Штатное ON Сотрудники.КодСотрудники = Штатное.ИДСотр
WHERE Сотрудники.КодСотрудники=@КодСотрудники", list)


        'Dim ds2 = Selects(StrSql:="SELECT ТипОбъекта, НазОбъекта FROM ОбъектОбщепита WHERE НазвОрг='" & ds(0).Item(0).ToString & "' AND АдресОбъекта='" & ds1.Rows(0).Item(14).ToString & "'"
        Dim ds2 = dtObjectObshepitaAll.Select("НазвОрг='" & ds(0).Item(0).ToString & "' AND АдресОбъекта='" & ds1.Rows(0).Item(14).ToString & "'")

        Dim obj As String
        Try
            If Not ds2(0).Item("ТипОбъекта").ToString <> "" And Not ds2(0).Item("НазОбъекта").ToString <> "" Then
                obj = ds1.Rows(0).Item(14).ToString
            ElseIf ds2(0).Item("ТипОбъекта").ToString = "" Then
                obj = ds2(0).Item("НазОбъекта").ToString & " " & ds1.Rows(0).Item(14).ToString
            ElseIf ds2(0).Item(1).ToString = "" Then
                obj = ds2(0).Item("ТипОбъекта").ToString & " " & ds1.Rows(0).Item(14).ToString
            Else
                obj = ds2(0).Item("ТипОбъекта").ToString & " " & ds2(0).Item("НазОбъекта").ToString & " " & ds1.Rows(0).Item(14).ToString
            End If
        Catch ex As Exception
            MessageBox.Show("Не найден адрес объекта общепита, перепроверьте!", Рик)
            Me.Cursor = DefaultCursor
            Exit Sub
        End Try





        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document

        oWord = CreateObject("Word.Application")
        oWord.Visible = False

        Начало("Spravka4mes.doc")
        oWordDoc = oWord.Documents.Add(firthtPath & "\Spravka4mes.doc")

        With oWordDoc.Bookmarks
            .Item("Сп1").Range.Text = ds(0).Item(1).ToString
            If ds(0).Item(1).ToString = "Индивидуальный предприниматель" Then
                .Item("Сп2").Range.Text = ds(0).Item(0).ToString
                .Item("Сп18").Range.Text = "у " & ФормСобствКор(ds(0).Item(1).ToString) & " " & ds(0).Item(0).ToString & " "
            Else
                .Item("Сп2").Range.Text = "«" & ds(0).Item(0).ToString & "»"
                .Item("Сп18").Range.Text = "в " & ФормСобствКор(ds(0).Item(1).ToString) & " «" & ds(0).Item(0).ToString & "» "
            End If
            .Item("Сп3").Range.Text = ds(0).Item(4).ToString 'адрес
            .Item("Сп4").Range.Text = ds(0).Item(2).ToString 'унп
            .Item("Сп5").Range.Text = ds(0).Item(14).ToString 'рс
            .Item("Сп6").Range.Text = ds(0).Item(12).ToString 'банк
            .Item("Сп7").Range.Text = ds(0).Item(11).ToString 'бик
            .Item("Сп8").Range.Text = ds(0).Item(8).ToString 'мыло
            .Item("Сп9").Range.Text = ds(0).Item(6).ToString 'тел
            .Item("Сп10").Range.Text = TextBox1.Text & "-" & Now.Year
            .Item("Сп11").Range.Text = MaskedTextBox1.Text

            Dim inp As String = InputBox("Введите ФИО сотрудника " & vbCrLf & ComboBox2.Text & vbCrLf & " в  Дательном падеже 'Справка выдана Кому?'", Рик)

            Do Until inp <> ""
                MessageBox.Show("Повторите ввод данных!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Error)
                inp = InputBox("Введите ФИО сотрудника " & vbCrLf & ComboBox2.Text & vbCrLf & " в  Дательном падеже 'Справка выдана Кому?'", Рик)
            Loop

            .Item("Сп12").Range.Text = inp
            .Item("Сп13").Range.Text = ds1.Rows(0).Item(1).ToString & ds1.Rows(0).Item(2).ToString
            .Item("Сп14").Range.Text = ds1.Rows(0).Item(4).ToString
            .Item("Сп15").Range.Text = ds1.Rows(0).Item(3).ToString
            .Item("Сп16").Range.Text = ds1.Rows(0).Item(5).ToString 'ид паспорта
            If ds1.Rows(0).Item(6).ToString = "М" Then
                .Item("Сп17").Range.Text = "он"
            Else
                .Item("Сп17").Range.Text = "она"
            End If
            .Item("Сп19").Range.Text = Strings.Left(ds1.Rows(0).Item(7).ToString, 10)

            If ds1.Rows(0).Item(15).ToString <> "" Then
                If ds1.Rows(0).Item(15).ToString = "-" Then
                    .Item("Сп20").Range.Text = Strings.LCase(ds1.Rows(0).Item(10).ToString)
                Else
                    .Item("Сп20").Range.Text = Strings.LCase(ds1.Rows(0).Item(10).ToString) & " " & разрядстрока(CType(ds1.Rows(0).Item(15).ToString, Integer))
                End If

            Else
                .Item("Сп20").Range.Text = Strings.LCase(ds1.Rows(0).Item(10).ToString) 'должность
            End If

            If ФормСобствКор(ds(0).Item(1).ToString) & " " & ds(0).Item(0).ToString = obj Then
                .Item("Сп21").Range.Text = ""
            Else
                .Item("Сп21").Range.Text = obj
            End If

            .Item("Сп22").Range.Text = Contr(CType(Label17.Text, Integer))

            '.Item("Сп26").Range.Text = TextBox2.Text
            '.Item("Сп27").Range.Text = TextBox3.Text
            .Item("Сп28").Range.Text = TextBox5.Text
            .Item("Сп29").Range.Text = TextBox4.Text
            .Item("Сп30").Range.Text = TextBox7.Text
            .Item("Сп31").Range.Text = TextBox6.Text
            '.Item("Сп32").Range.Text = TextBox8.Text
            '.Item("Сп33").Range.Text = TextBox9.Text
            .Item("Сп34").Range.Text = TextBox11.Text
            .Item("Сп35").Range.Text = TextBox10.Text
            .Item("Сп36").Range.Text = TextBox13.Text
            .Item("Сп37").Range.Text = TextBox12.Text
            '.Item("Сп38").Range.Text = TextBox14.Text
            '.Item("Сп39").Range.Text = TextBox15.Text
            .Item("Сп40").Range.Text = TextBox17.Text
            .Item("Сп41").Range.Text = TextBox16.Text
            .Item("Сп42").Range.Text = TextBox19.Text
            .Item("Сп43").Range.Text = TextBox18.Text
            '.Item("Сп44").Range.Text = TextBox25.Text
            '.Item("Сп45").Range.Text = TextBox24.Text
            .Item("Сп46").Range.Text = TextBox23.Text
            .Item("Сп47").Range.Text = TextBox22.Text
            .Item("Сп48").Range.Text = TextBox21.Text
            .Item("Сп49").Range.Text = TextBox20.Text
            '.Item("Сп50").Range.Text = TextBox31.Text
            '.Item("Сп51").Range.Text = TextBox30.Text
            .Item("Сп52").Range.Text = TextBox29.Text
            .Item("Сп53").Range.Text = TextBox28.Text
            .Item("Сп54").Range.Text = TextBox27.Text
            .Item("Сп55").Range.Text = TextBox26.Text
            '.Item("Сп56").Range.Text = TextBox37.Text
            '.Item("Сп57").Range.Text = TextBox36.Text
            .Item("Сп58").Range.Text = TextBox35.Text
            .Item("Сп59").Range.Text = TextBox34.Text
            .Item("Сп60").Range.Text = TextBox33.Text
            .Item("Сп61").Range.Text = TextBox32.Text
            '.Item("Сп62").Range.Text = TextBox43.Text
            '.Item("Сп63").Range.Text = TextBox42.Text
            .Item("Сп64").Range.Text = TextBox41.Text
            .Item("Сп65").Range.Text = TextBox40.Text
            .Item("Сп66").Range.Text = TextBox39.Text
            .Item("Сп67").Range.Text = TextBox38.Text
            '.Item("Сп68").Range.Text = TextBox49.Text
            '.Item("Сп69").Range.Text = TextBox48.Text
            .Item("Сп70").Range.Text = TextBox47.Text
            .Item("Сп71").Range.Text = TextBox46.Text
            .Item("Сп72").Range.Text = TextBox45.Text
            .Item("Сп73").Range.Text = TextBox44.Text
            .Item("Сп74").Range.Text = TextBox50.Text
            .Item("Сп75").Range.Text = TextBox51.Text
            .Item("Сп76").Range.Text = TextBox52.Text
            .Item("Сп77").Range.Text = TextBox53.Text
            .Item("Сп78").Range.Text = TextBox54.Text
            .Item("Сп79").Range.Text = TextBox55.Text
            .Item("Сп80").Range.Text = TextBox56.Text
            .Item("Сп81").Range.Text = TextBox57.Text
            .Item("Сп82").Range.Text = TextBox58.Text
            .Item("Сп83").Range.Text = TextBox59.Text
            .Item("Сп84").Range.Text = ds1.Rows(0).Item(16).ToString & " " & Strings.Left(ds1.Rows(0).Item(17).ToString, 1) & "." & Strings.Left(ds1.Rows(0).Item(18).ToString, 1) & "."
            .Item("Сп85").Range.Text = TextBox55.Text
            .Item("Сп86").Range.Text = ЧислоПрописДляСправки(TextBox55.Text)
            If ds(0).Item(18).ToString = "Индивидуальный предприниматель" Then
                .Item("Сп88").Range.Text = ds(0).Item(18).ToString
                .Item("Сп89").Range.Text = ФИОКорРук(ds(0).Item(19).ToString, False)
            Else
                .Item("Сп88").Range.Text = ds(0).Item(18).ToString & " " & ФормСобствКор(ds(0).Item(1).ToString) & " «" & ComboBox1.Text & "» "
                If ds(0).Item(31) = True Then
                    .Item("Сп89").Range.Text = ФИОКорРук(ds(0).Item(19).ToString, True)
                Else
                    .Item("Сп89").Range.Text = ФИОКорРук(ds(0).Item(19).ToString, False)
                End If


            End If
        End With


        Dim Name As String = "Справка №" & TextBox1.Text & "-" & Now.Year & " от " & MaskedTextBox1.Text & " " & ФИОКорРук(ComboBox2.Text, False) & " 4 месяца" & ".doc"
        Dim СохрЗак2 As New List(Of String)
        СохрЗак2.AddRange(New String() {ComboBox1.Text & "\Справки\" & Now.Year & "\", Name})
        oWordDoc.SaveAs2(PathVremyanka & Name,,,,,, False)
        oWordDoc.Close(True)
        oWord.Quit(True)
        Конец(ComboBox1.Text & "\Справки\" & Now.Year, Name, CType(Label17.Text, Integer), ComboBox1.Text, "\Spravka4mes.doc", "Справка по зарплате на 4 месяца")
        Dim massFTP As New ArrayList()
        massFTP.Add(СохрЗак2)



        If MessageBox.Show("Справка оформлена!" & vbCrLf & "Распечатать?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            ПечатьДоковFTP(massFTP)
        End If
        Статистика1(ComboBox2.Text, "Оформление справки по месту требования на 4 месяца", ComboBox1.Text)
        Me.Cursor = DefaultCursor

        Me.Close()
    End Sub
    Private Sub доки5()
        Me.Cursor = Cursors.WaitCursor

        Dim ds1 As DataTable

        'strsql = "SELECT * FROM Клиент WHERE НазвОрг='" & ComboBox1.Text & "'" 'Данные клиента
        'ds = Selects(strsql)

        Dim ds = dtClientAll.Select("НазвОрг='" & ComboBox1.Text & "'")
        Dim list As New Dictionary(Of String, Object)
        list.Add("@КодСотрудники", CType(Label17.Text, Integer))

        ds1 = Selects(StrSql:="SELECT Сотрудники.ФИОСборное, Сотрудники.ПаспортСерия, Сотрудники.ПаспортНомер, Сотрудники.ПаспортКогдаВыдан,
Сотрудники.ПаспортКемВыдан, Сотрудники.ИДНомер, Сотрудники.Пол, КарточкаСотрудника.ДатаПриема, КарточкаСотрудника.СрокКонтракта, ДогСотрудн.СрокОкончКонтр,
Штатное.Должность, КарточкаСотрудника.ПродлКонтрС, КарточкаСотрудника.ПродлКонтрПо, КарточкаСотрудника.СрокПродлКонтракта, КарточкаСотрудника.АдресОбъектаОбщепита, Штатное.Разряд,
Сотрудники.ФамилияДляЗаявления, Сотрудники.ИмяДляЗаявления, Сотрудники.ОтчествоДляЗаявления
FROM ((Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр) INNER JOIN ДогСотрудн ON Сотрудники.КодСотрудники = ДогСотрудн.IDСотр) INNER JOIN Штатное ON Сотрудники.КодСотрудники = Штатное.ИДСотр
WHERE Сотрудники.КодСотрудники=@КодСотрудники", list)

        Dim ds2 = dtObjectObshepitaAll.Select("НазвОрг='" & ds(0).Item(0).ToString & "' AND АдресОбъекта='" & ds1.Rows(0).Item(14).ToString & "'")
        'strsql2 = "SELECT ТипОбъекта, НазОбъекта FROM ОбъектОбщепита WHERE НазвОрг='" & ds(0).Item(0).ToString & "' AND АдресОбъекта='" & ds1.Rows(0).Item(14).ToString & "'"
        'ds2 = Selects(strsql2)
        Dim obj As String
        Try
            If Not ds2(0).Item("ТипОбъекта").ToString <> "" And Not ds2(0).Item("НазОбъекта").ToString <> "" Then
                obj = ds1.Rows(0).Item(14).ToString
            ElseIf ds2(0).Item("ТипОбъекта").ToString = "" Then
                obj = ds2(0).Item("НазОбъекта").ToString & " " & ds1.Rows(0).Item(14).ToString
            ElseIf ds2(0).Item("НазОбъекта").ToString = "" Then
                obj = ds2(0).Item("ТипОбъекта").ToString & " " & ds1.Rows(0).Item(14).ToString
            Else
                obj = ds2(0).Item("ТипОбъекта").ToString & " " & ds2(0).Item("НазОбъекта").ToString & " " & ds1.Rows(0).Item(14).ToString
            End If
        Catch ex As Exception
            MessageBox.Show("Не найден адрес объекта общепита, перепроверьте!", Рик)
            Me.Cursor = DefaultCursor
            Exit Sub
        End Try





        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document
        oWord = CreateObject("Word.Application")
        oWord.Visible = False
        Начало("Spravka5mes.doc")
        oWordDoc = oWord.Documents.Add(firthtPath & "\Spravka5mes.doc")

        With oWordDoc.Bookmarks
            .Item("Сп1").Range.Text = ds(0).Item(1).ToString
            If ds(0).Item(1).ToString = "Индивидуальный предприниматель" Then
                .Item("Сп2").Range.Text = ds(0).Item(0).ToString
                .Item("Сп18").Range.Text = "у " & ФормСобствКор(ds(0).Item(1).ToString) & " " & ds(0).Item(0).ToString & " "
            Else
                .Item("Сп2").Range.Text = "«" & ds(0).Item(0).ToString & "»"
                .Item("Сп18").Range.Text = "в " & ФормСобствКор(ds(0).Item(1).ToString) & " «" & ds(0).Item(0).ToString & "» "
            End If
            .Item("Сп3").Range.Text = ds(0).Item(4).ToString 'адрес
            .Item("Сп4").Range.Text = ds(0).Item(2).ToString 'унп
            .Item("Сп5").Range.Text = ds(0).Item(14).ToString 'рс
            .Item("Сп6").Range.Text = ds(0).Item(12).ToString 'банк
            .Item("Сп7").Range.Text = ds(0).Item(11).ToString 'бик
            .Item("Сп8").Range.Text = ds(0).Item(8).ToString 'мыло
            .Item("Сп9").Range.Text = ds(0).Item(6).ToString 'тел
            .Item("Сп10").Range.Text = TextBox1.Text & "-" & Now.Year
            .Item("Сп11").Range.Text = MaskedTextBox1.Text

            Dim inp As String = InputBox("Введите ФИО сотрудника " & vbCrLf & ComboBox2.Text & vbCrLf & " в  Дательном падеже 'Справка выдана Кому?'", Рик)

            Do Until inp <> ""
                MessageBox.Show("Повторите ввод данных!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Error)
                inp = InputBox("Введите ФИО сотрудника " & vbCrLf & ComboBox2.Text & vbCrLf & " в  Дательном падеже 'Справка выдана Кому?'", Рик)
            Loop

            .Item("Сп12").Range.Text = inp
            .Item("Сп13").Range.Text = ds1.Rows(0).Item(1).ToString & ds1.Rows(0).Item(2).ToString
            .Item("Сп14").Range.Text = ds1.Rows(0).Item(4).ToString
            .Item("Сп15").Range.Text = ds1.Rows(0).Item(3).ToString
            .Item("Сп16").Range.Text = ds1.Rows(0).Item(5).ToString 'ид паспорта
            If ds1.Rows(0).Item(6).ToString = "М" Then
                .Item("Сп17").Range.Text = "он"
            Else
                .Item("Сп17").Range.Text = "она"
            End If
            .Item("Сп19").Range.Text = Strings.Left(ds1.Rows(0).Item(7).ToString, 10)

            If ds1.Rows(0).Item(15).ToString <> "" Then
                If ds1.Rows(0).Item(15).ToString = "-" Then
                    .Item("Сп20").Range.Text = Strings.LCase(ds1.Rows(0).Item(10).ToString)
                Else
                    .Item("Сп20").Range.Text = Strings.LCase(ds1.Rows(0).Item(10).ToString) & " " & разрядстрока(CType(ds1.Rows(0).Item(15).ToString, Integer))
                End If

            Else
                .Item("Сп20").Range.Text = Strings.LCase(ds1.Rows(0).Item(10).ToString) 'должность
            End If

            If ФормСобствКор(ds(0).Item(1).ToString) & " " & ds(0).Item(0).ToString = obj Then
                .Item("Сп21").Range.Text = ""
            Else
                .Item("Сп21").Range.Text = obj
            End If

            .Item("Сп22").Range.Text = Contr(CType(Label17.Text, Integer))

            '.Item("Сп26").Range.Text = TextBox2.Text
            .Item("Сп27").Range.Text = TextBox3.Text
            .Item("Сп28").Range.Text = TextBox5.Text
            .Item("Сп29").Range.Text = TextBox4.Text
            .Item("Сп30").Range.Text = TextBox7.Text
            .Item("Сп31").Range.Text = TextBox6.Text
            '.Item("Сп32").Range.Text = TextBox8.Text
            .Item("Сп33").Range.Text = TextBox9.Text
            .Item("Сп34").Range.Text = TextBox11.Text
            .Item("Сп35").Range.Text = TextBox10.Text
            .Item("Сп36").Range.Text = TextBox13.Text
            .Item("Сп37").Range.Text = TextBox12.Text
            '.Item("Сп38").Range.Text = TextBox14.Text
            .Item("Сп39").Range.Text = TextBox15.Text
            .Item("Сп40").Range.Text = TextBox17.Text
            .Item("Сп41").Range.Text = TextBox16.Text
            .Item("Сп42").Range.Text = TextBox19.Text
            .Item("Сп43").Range.Text = TextBox18.Text
            '.Item("Сп44").Range.Text = TextBox25.Text
            .Item("Сп45").Range.Text = TextBox24.Text
            .Item("Сп46").Range.Text = TextBox23.Text
            .Item("Сп47").Range.Text = TextBox22.Text
            .Item("Сп48").Range.Text = TextBox21.Text
            .Item("Сп49").Range.Text = TextBox20.Text
            '.Item("Сп50").Range.Text = TextBox31.Text
            .Item("Сп51").Range.Text = TextBox30.Text
            .Item("Сп52").Range.Text = TextBox29.Text
            .Item("Сп53").Range.Text = TextBox28.Text
            .Item("Сп54").Range.Text = TextBox27.Text
            .Item("Сп55").Range.Text = TextBox26.Text
            '.Item("Сп56").Range.Text = TextBox37.Text
            .Item("Сп57").Range.Text = TextBox36.Text
            .Item("Сп58").Range.Text = TextBox35.Text
            .Item("Сп59").Range.Text = TextBox34.Text
            .Item("Сп60").Range.Text = TextBox33.Text
            .Item("Сп61").Range.Text = TextBox32.Text
            '.Item("Сп62").Range.Text = TextBox43.Text
            .Item("Сп63").Range.Text = TextBox42.Text
            .Item("Сп64").Range.Text = TextBox41.Text
            .Item("Сп65").Range.Text = TextBox40.Text
            .Item("Сп66").Range.Text = TextBox39.Text
            .Item("Сп67").Range.Text = TextBox38.Text
            '.Item("Сп68").Range.Text = TextBox49.Text
            .Item("Сп69").Range.Text = TextBox48.Text
            .Item("Сп70").Range.Text = TextBox47.Text
            .Item("Сп71").Range.Text = TextBox46.Text
            .Item("Сп72").Range.Text = TextBox45.Text
            .Item("Сп73").Range.Text = TextBox44.Text
            .Item("Сп74").Range.Text = TextBox50.Text
            .Item("Сп75").Range.Text = TextBox51.Text
            .Item("Сп76").Range.Text = TextBox52.Text
            .Item("Сп77").Range.Text = TextBox53.Text
            .Item("Сп78").Range.Text = TextBox54.Text
            .Item("Сп79").Range.Text = TextBox55.Text
            .Item("Сп80").Range.Text = TextBox56.Text
            .Item("Сп81").Range.Text = TextBox57.Text
            .Item("Сп82").Range.Text = TextBox58.Text
            .Item("Сп83").Range.Text = TextBox59.Text
            .Item("Сп84").Range.Text = ds1.Rows(0).Item(16).ToString & " " & Strings.Left(ds1.Rows(0).Item(17).ToString, 1) & "." & Strings.Left(ds1.Rows(0).Item(18).ToString, 1) & "."
            .Item("Сп85").Range.Text = TextBox55.Text
            .Item("Сп86").Range.Text = ЧислоПрописДляСправки(TextBox55.Text)
            If ds(0).Item(18).ToString = "Индивидуальный предприниматель" Then
                .Item("Сп88").Range.Text = ds(0).Item(18).ToString
                .Item("Сп89").Range.Text = ФИОКорРук(ds(0).Item(19).ToString, False)
            Else
                .Item("Сп88").Range.Text = ds(0).Item(18).ToString & " " & ФормСобствКор(ds(0).Item(1).ToString) & " «" & ComboBox1.Text & "» "
                If ds(0).Item(31) = True Then
                    .Item("Сп89").Range.Text = ФИОКорРук(ds(0).Item(19).ToString, True)
                Else
                    .Item("Сп89").Range.Text = ФИОКорРук(ds(0).Item(19).ToString, False)
                End If


            End If

        End With


        Dim Name As String = "Справка №" & TextBox1.Text & "-" & Now.Year & " от " & MaskedTextBox1.Text & " " & ФИОКорРук(ComboBox2.Text, False) & " 5 месяцев" & ".doc"
        Dim СохрЗак2 As New List(Of String)
        СохрЗак2.AddRange(New String() {ComboBox1.Text & "\Справки\" & Now.Year & "\", Name})
        oWordDoc.SaveAs2(PathVremyanka & Name,,,,,, False)
        oWordDoc.Close(True)
        oWord.Quit(True)
        Конец(ComboBox1.Text & "\Справки\" & Now.Year, Name, CType(Label17.Text, Integer), ComboBox1.Text, "\Spravka5mes.doc", "Справка по зарплате на 5 месяцев")
        Dim massFTP As New ArrayList()
        massFTP.Add(СохрЗак2)


        If MessageBox.Show("Справка оформлена!" & vbCrLf & "Распечатать?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            ПечатьДоковFTP(massFTP)
        End If
        Статистика1(ComboBox2.Text, "Оформление справки по месту требования на 5 месяцев", ComboBox1.Text)
        Me.Cursor = DefaultCursor

        Me.Close()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Проверка() = 1 Then Exit Sub

        Select Case CType(ComboBox10.Text, Integer)
            Case 3
                доки3()
                Статистика1(ComboBox2.Text, "Оформление справки по месту требования на 3 месяца", ComboBox1.Text)
                Exit Sub
            Case 4
                доки4()
                Статистика1(ComboBox2.Text, "Оформление справки по месту требования на 4 месяца", ComboBox1.Text)
                Exit Sub
            Case 5
                доки5()
                Статистика1(ComboBox2.Text, "Оформление справки по месту требования на 5 месяцев", ComboBox1.Text)
                Exit Sub
        End Select

        Me.Cursor = Cursors.WaitCursor
        Dim strsql, strsql1, strsql2, strsql3 As String
        Dim ds1, ds3 As DataTable

        'strsql = "SELECT * FROM Клиент WHERE НазвОрг='" & ComboBox1.Text & "'" 'Данные клиента
        'ds = Selects(strsql)

        Dim ds = dtClientAll.Select("НазвОрг='" & ComboBox1.Text & "'")

        Dim list As New Dictionary(Of String, Object)
        list.Add("@КодСотрудники", CType(Label17.Text, Integer))

        ds1 = Selects(StrSql:= "SELECT Сотрудники.ФИОСборное, Сотрудники.ПаспортСерия, Сотрудники.ПаспортНомер, Сотрудники.ПаспортКогдаВыдан,
Сотрудники.ПаспортКемВыдан, Сотрудники.ИДНомер, Сотрудники.Пол, КарточкаСотрудника.ДатаПриема, КарточкаСотрудника.СрокКонтракта, ДогСотрудн.СрокОкончКонтр,
Штатное.Должность, КарточкаСотрудника.ПродлКонтрС, КарточкаСотрудника.ПродлКонтрПо, КарточкаСотрудника.СрокПродлКонтракта, КарточкаСотрудника.АдресОбъектаОбщепита, Штатное.Разряд,
Сотрудники.ФамилияДляЗаявления, Сотрудники.ИмяДляЗаявления, Сотрудники.ОтчествоДляЗаявления
FROM ((Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр) INNER JOIN ДогСотрудн ON Сотрудники.КодСотрудники = ДогСотрудн.IDСотр) INNER JOIN Штатное ON Сотрудники.КодСотрудники = Штатное.ИДСотр
WHERE Сотрудники.КодСотрудники=@КодСотрудники", list)

        Dim ds2 = dtObjectObshepitaAll.Select("НазвОрг='" & ds(0).Item(0).ToString & "' AND АдресОбъекта='" & ds1.Rows(0).Item(14).ToString & "'")

        'strsql2 = "SELECT ТипОбъекта, НазОбъекта FROM ОбъектОбщепита WHERE НазвОрг='" & ds.Rows(0).Item(0).ToString & "' AND АдресОбъекта='" & ds1.Rows(0).Item(14).ToString & "'"
        'ds2 = Selects(strsql2)
        Dim obj As String
        Try
            If Not ds2(0).Item("ТипОбъекта").ToString <> "" And Not ds2(0).Item("НазОбъекта").ToString <> "" Then
                obj = ds1.Rows(0).Item(14).ToString
            ElseIf ds2(0).Item("ТипОбъекта").ToString = "" Then
                obj = ds2(0).Item("НазОбъекта").ToString & " " & ds1.Rows(0).Item(14).ToString
            ElseIf ds2(0).Item("НазОбъекта").ToString = "" Then
                obj = ds2(0).Item("ТипОбъекта").ToString & " " & ds1.Rows(0).Item(14).ToString
            Else
                obj = ds2(0).Item("ТипОбъекта").ToString & " " & ds2(0).Item("НазОбъекта").ToString & " " & ds1.Rows(0).Item(14).ToString
            End If
        Catch ex As Exception
            MessageBox.Show("Не найден адрес объекта общепита, перепроверьте!", Рик)
            Me.Cursor = DefaultCursor
            Exit Sub
        End Try

        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document
        oWord = CreateObject("Word.Application")
        oWord.Visible = False

        Начало("Spravka.doc")
        oWordDoc = oWord.Documents.Add(firthtPath & "\Spravka.doc")

        With oWordDoc.Bookmarks
            .Item("Сп1").Range.Text = ds(0).Item(1).ToString
            If ds(0).Item(1).ToString = "Индивидуальный предприниматель" Then
                .Item("Сп2").Range.Text = ds(0).Item(0).ToString
                .Item("Сп18").Range.Text = "у " & ФормСобствКор(ds(0).Item(1).ToString) & " " & ds(0).Item(0).ToString & " "
            Else
                .Item("Сп2").Range.Text = "«" & ds(0).Item(0).ToString & "»"
                .Item("Сп18").Range.Text = "в " & ФормСобствКор(ds(0).Item(1).ToString) & " «" & ds(0).Item(0).ToString & "» "
            End If
            .Item("Сп3").Range.Text = ds(0).Item(4).ToString 'адрес
            .Item("Сп4").Range.Text = ds(0).Item(2).ToString 'унп
            .Item("Сп5").Range.Text = ds(0).Item(14).ToString 'рс
            .Item("Сп6").Range.Text = ds(0).Item(12).ToString 'банк
            .Item("Сп7").Range.Text = ds(0).Item(11).ToString 'бик
            .Item("Сп8").Range.Text = ds(0).Item(8).ToString 'мыло
            .Item("Сп9").Range.Text = ds(0).Item(6).ToString 'тел
            .Item("Сп10").Range.Text = TextBox1.Text & "-" & Now.Year
            .Item("Сп11").Range.Text = MaskedTextBox1.Text

            Dim inp As String = InputBox("Введите ФИО сотрудника " & vbCrLf & ComboBox2.Text & vbCrLf & " в  Дательном падеже 'Справка выдана Кому?'", Рик)

            Do Until inp <> ""
                MessageBox.Show("Повторите ввод данных!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Error)
                inp = InputBox("Введите ФИО сотрудника " & vbCrLf & ComboBox2.Text & vbCrLf & " в  Дательном падеже 'Справка выдана Кому?'", Рик)
            Loop

            .Item("Сп12").Range.Text = inp
            .Item("Сп13").Range.Text = ds1.Rows(0).Item(1).ToString & ds1.Rows(0).Item(2).ToString
            .Item("Сп14").Range.Text = ds1.Rows(0).Item(4).ToString
            .Item("Сп15").Range.Text = ds1.Rows(0).Item(3).ToString
            .Item("Сп16").Range.Text = ds1.Rows(0).Item(5).ToString 'ид паспорта
            If ds1.Rows(0).Item(6).ToString = "М" Then
                .Item("Сп17").Range.Text = "он"
            Else
                .Item("Сп17").Range.Text = "она"
            End If
            .Item("Сп19").Range.Text = Strings.Left(ds1.Rows(0).Item(7).ToString, 10)

            If ds1.Rows(0).Item(15).ToString <> "" Then
                If ds1.Rows(0).Item(15).ToString = "-" Then
                    .Item("Сп20").Range.Text = Strings.LCase(ds1.Rows(0).Item(10).ToString)
                Else
                    .Item("Сп20").Range.Text = Strings.LCase(ds1.Rows(0).Item(10).ToString) & " " & разрядстрока(CType(ds1.Rows(0).Item(15).ToString, Integer))
                End If

            Else
                .Item("Сп20").Range.Text = Strings.LCase(ds1.Rows(0).Item(10).ToString) 'должность
            End If

            If ФормСобствКор(ds(0).Item(1).ToString) & " " & ds(0).Item(0).ToString = obj Then
                .Item("Сп21").Range.Text = ""
            Else
                .Item("Сп21").Range.Text = obj
            End If

            .Item("Сп22").Range.Text = Contr(CType(Label17.Text, Integer))

            .Item("Сп26").Range.Text = TextBox2.Text
            .Item("Сп27").Range.Text = TextBox3.Text
            .Item("Сп28").Range.Text = TextBox5.Text
            .Item("Сп29").Range.Text = TextBox4.Text
            .Item("Сп30").Range.Text = TextBox7.Text
            .Item("Сп31").Range.Text = TextBox6.Text
            .Item("Сп32").Range.Text = TextBox8.Text
            .Item("Сп33").Range.Text = TextBox9.Text
            .Item("Сп34").Range.Text = TextBox11.Text
            .Item("Сп35").Range.Text = TextBox10.Text
            .Item("Сп36").Range.Text = TextBox13.Text
            .Item("Сп37").Range.Text = TextBox12.Text
            .Item("Сп38").Range.Text = TextBox14.Text
            .Item("Сп39").Range.Text = TextBox15.Text
            .Item("Сп40").Range.Text = TextBox17.Text
            .Item("Сп41").Range.Text = TextBox16.Text
            .Item("Сп42").Range.Text = TextBox19.Text
            .Item("Сп43").Range.Text = TextBox18.Text
            .Item("Сп44").Range.Text = TextBox25.Text
            .Item("Сп45").Range.Text = TextBox24.Text
            .Item("Сп46").Range.Text = TextBox23.Text
            .Item("Сп47").Range.Text = TextBox22.Text
            .Item("Сп48").Range.Text = TextBox21.Text
            .Item("Сп49").Range.Text = TextBox20.Text
            .Item("Сп50").Range.Text = TextBox31.Text
            .Item("Сп51").Range.Text = TextBox30.Text
            .Item("Сп52").Range.Text = TextBox29.Text
            .Item("Сп53").Range.Text = TextBox28.Text
            .Item("Сп54").Range.Text = TextBox27.Text
            .Item("Сп55").Range.Text = TextBox26.Text
            .Item("Сп56").Range.Text = TextBox37.Text
            .Item("Сп57").Range.Text = TextBox36.Text
            .Item("Сп58").Range.Text = TextBox35.Text
            .Item("Сп59").Range.Text = TextBox34.Text
            .Item("Сп60").Range.Text = TextBox33.Text
            .Item("Сп61").Range.Text = TextBox32.Text
            .Item("Сп62").Range.Text = TextBox43.Text
            .Item("Сп63").Range.Text = TextBox42.Text
            .Item("Сп64").Range.Text = TextBox41.Text
            .Item("Сп65").Range.Text = TextBox40.Text
            .Item("Сп66").Range.Text = TextBox39.Text
            .Item("Сп67").Range.Text = TextBox38.Text
            .Item("Сп68").Range.Text = TextBox49.Text
            .Item("Сп69").Range.Text = TextBox48.Text
            .Item("Сп70").Range.Text = TextBox47.Text
            .Item("Сп71").Range.Text = TextBox46.Text
            .Item("Сп72").Range.Text = TextBox45.Text
            .Item("Сп73").Range.Text = TextBox44.Text
            .Item("Сп74").Range.Text = TextBox50.Text
            .Item("Сп75").Range.Text = TextBox51.Text
            .Item("Сп76").Range.Text = TextBox52.Text
            .Item("Сп77").Range.Text = TextBox53.Text
            .Item("Сп78").Range.Text = TextBox54.Text
            .Item("Сп79").Range.Text = TextBox55.Text
            .Item("Сп80").Range.Text = TextBox56.Text
            .Item("Сп81").Range.Text = TextBox57.Text
            .Item("Сп82").Range.Text = TextBox58.Text
            .Item("Сп83").Range.Text = TextBox59.Text
            .Item("Сп84").Range.Text = ds1.Rows(0).Item(16).ToString & " " & Strings.Left(ds1.Rows(0).Item(17).ToString, 1) & "." & Strings.Left(ds1.Rows(0).Item(18).ToString, 1) & "."
            .Item("Сп85").Range.Text = TextBox55.Text
            .Item("Сп86").Range.Text = ЧислоПрописДляСправки(TextBox55.Text)
            If ds(0).Item(18).ToString = "Индивидуальный предприниматель" Then
                .Item("Сп88").Range.Text = ds(0).Item(18).ToString
                .Item("Сп89").Range.Text = ФИОКорРук(ds(0).Item(19).ToString, False)
            Else
                .Item("Сп88").Range.Text = ds(0).Item(18).ToString & " " & ФормСобствКор(ds(0).Item(1).ToString) & " «" & ComboBox1.Text & "» "
                If ds(0).Item(31) = True Then
                    .Item("Сп89").Range.Text = ФИОКорРук(ds(0).Item(19).ToString, True)
                Else
                    .Item("Сп89").Range.Text = ФИОКорРук(ds(0).Item(19).ToString, False)
                End If


            End If

        End With

        Dim Name As String = "Справка №" & TextBox1.Text & "-" & Now.Year & " от " & MaskedTextBox1.Text & " " & ФИОКорРук(ComboBox2.Text, False) & ".doc"
        Dim СохрЗак2 As New List(Of String)
        СохрЗак2.AddRange(New String() {ComboBox1.Text & "\Справки\" & Now.Year & "\", Name})
        oWordDoc.SaveAs2(PathVremyanka & Name,,,,,, False)
        oWordDoc.Close(True)
        oWord.Quit(True)
        Конец(ComboBox1.Text & "\Справки\" & Now.Year, Name, CType(Label17.Text, Integer), ComboBox1.Text, "\Spravka.doc", "Справка по зарплате")
        massFTP3.Add(СохрЗак2)


        'If Not IO.Directory.Exists(OnePath & ComboBox1.Text & "\Справки\" & Now.Year) Then
        '    IO.Directory.CreateDirectory(OnePath & ComboBox1.Text & "\Справки\" & Now.Year)
        'End If

        'oWordDoc.SaveAs2("C:\Users\Public\Documents\Рик\Справка №" & TextBox1.Text & "-" & Now.Year & " от " & MaskedTextBox1.Text & " " & ФИОКорРук(ComboBox2.Text, False) & ".doc",,,,,, False)

        'Try
        '    IO.File.Copy("C:\Users\Public\Documents\Рик\Справка №" & TextBox1.Text & "-" & Now.Year & " от " & MaskedTextBox1.Text & " " & ФИОКорРук(ComboBox2.Text, False) & ".doc", OnePath & ComboBox1.Text & "\Справки\" & Now.Year & "\Справка №" & TextBox1.Text & "-" & Now.Year & " от " & MaskedTextBox1.Text & " " & ФИОКорРук(ComboBox2.Text, False) & ".doc")
        'Catch ex As Exception
        '    If MessageBox.Show("Справка №" & TextBox1.Text & "-" & Now.Year & " с сотрудником " & ФИОКорРук(ComboBox2.Text, False) & " существует. Заменить старый документ новым?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then
        '        Try
        '            IO.File.Delete(OnePath & ComboBox1.Text & "\Справки\" & Now.Year & "\Справка №" & TextBox1.Text & "-" & Now.Year & " от " & MaskedTextBox1.Text & " " & ФИОКорРук(ComboBox2.Text, False) & ".doc")
        '        Catch ex1 As Exception
        '            MessageBox.Show("Закройте файл!", Рик)
        '        End Try

        '        IO.File.Copy("C:\Users\Public\Documents\Рик\Справка №" & TextBox1.Text & "-" & Now.Year & " от " & MaskedTextBox1.Text & " " & ФИОКорРук(ComboBox2.Text, False) & ".doc", OnePath & ComboBox1.Text & "\Справки\" & Now.Year & "\Справка №" & TextBox1.Text & "-" & Now.Year & " от " & MaskedTextBox1.Text & " " & ФИОКорРук(ComboBox2.Text, False) & ".doc")
        '    End If
        'End Try
        'СохрЗак = OnePath & ComboBox1.Text & "\Справки\" & Now.Year & "\Справка №" & TextBox1.Text & "-" & Now.Year & " от " & MaskedTextBox1.Text & " " & ФИОКорРук(ComboBox2.Text, False) & ".doc"
        'oWordDoc.Close(True)
        'Dim mass() As String

        'mass = {СохрЗак}
        If MessageBox.Show("Справка оформлена!" & vbCrLf & "Распечатать?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            ПечатьДоковFTP(massFTP3)
        End If
        Статистика1(ComboBox2.Text, "Оформление справки по месту требования", ComboBox1.Text)
        Me.Cursor = DefaultCursor

        Me.Close()
    End Sub

    Function Count(ByVal Number As Double) As Integer
        Return Split(Number.ToString(), ",")(1).Length
    End Function

    Private Sub TextBox23_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox23.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Dim a As Double
            a = Replace(TextBox23.Text, ".", ",")
            TextBox23.SelectionStart = TextBox23.Text.Length
            If bool(a) = True Then
                TextBox23.Text = a & ",00"
            Else
                TextBox23.Text = a
                If Count(a) = 1 Then
                    TextBox23.Text = a & "0"
                End If
            End If
            общсправа(23)
            пров()
            TextBox22.Focus()
        End If
    End Sub

    Private Sub TextBox22_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox22.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Dim a As Double
            a = Replace(TextBox22.Text, ".", ",")
            TextBox22.SelectionStart = TextBox22.Text.Length
            If bool(a) = True Then
                TextBox22.Text = a & ",00"
            Else
                TextBox22.Text = a
                If Count(a) = 1 Then
                    TextBox22.Text = a & "0"
                End If
            End If
            общсправа(22)
            пров()
            TextBox21.Focus()
        End If
    End Sub

    Private Sub TextBox21_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox21.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Dim a As Double
            a = Replace(TextBox21.Text, ".", ",")
            TextBox21.SelectionStart = TextBox21.Text.Length
            If bool(a) = True Then
                TextBox21.Text = a & ",00"
            Else
                TextBox21.Text = a
                If Count(a) = 1 Then
                    TextBox21.Text = a & "0"
                End If
            End If
            общсправа(21)
            пров()
            TextBox20.Focus()
        End If
    End Sub

    Private Sub TextBox20_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox20.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Dim a As Double
            a = Replace(TextBox20.Text, ".", ",")
            TextBox20.SelectionStart = TextBox20.Text.Length
            If bool(a) = True Then
                TextBox20.Text = a & ",00"
            Else
                TextBox20.Text = a
                If Count(a) = 1 Then
                    TextBox20.Text = a & "0"
                End If
            End If
            общсправа(20)
            пров()
            CheckBox1.Focus()
        End If
    End Sub
End Class