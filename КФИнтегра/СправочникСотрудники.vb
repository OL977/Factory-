Public Class СправочникСотрудники
    Public BtnClick As String



    Private Sub TabControl1_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles TabControl1.DrawItem
        Dim g As Graphics
        Dim sText As String

        Dim iX As Integer
        Dim iY As Integer
        Dim sizeText As SizeF

        Dim ctlTab As TabControl

        ctlTab = CType(sender, TabControl)

        g = e.Graphics

        sText = ctlTab.TabPages(e.Index).Text
        sizeText = g.MeasureString(sText, ctlTab.Font)

        iX = e.Bounds.Left + 6
        iY = e.Bounds.Top + (e.Bounds.Height - sizeText.Height) / 2

        g.DrawString(sText, ctlTab.Font, Brushes.Black, iX, iY)
    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub СправочникСотрудники_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox3.Visible = False
        Label33.Visible = False

        Dim dbcx As New DbAllDataContext
        Dim ds = From x In dbcx.Клиент.AsEnumerable Order By x.НазвОрг Select x.НазвОрг
        ComboBox2.DataSource = ds.ToList
        If BtnClick = "Изменить" Then
            ComboBox3.Visible = True
            Label33.Visible = True
            com2()
        End If
        ToolTip1.SetToolTip(Me.Button5, "Добавить")
        ToolTip1.SetToolTip(Me.Button6, "Очистить")
        ToolTip1.SetToolTip(Me.Button2, "Изменить")
        ToolTip1.SetToolTip(Me.Button3, "Удалить")


        Grid1Load()



    End Sub
    Private Sub Grid1Load()
        Dim _bs As New BindingSource()
        Dim dt As New DataTable
        dt.Columns.Add("ФИО")
        dt.Columns.Add("Пол")
        dt.Columns.Add("Дата рождения")
        _bs.DataSource = dt
        Grid1.DataSource = _bs
        BindingNavigator1.BindingSource = _bs

        GridView(Grid1)
        'Grid1.Columns(2).DefaultCellStyle.Format = "d"  ' ToString("dd/MM/yyyy")


        'Grid1.Columns(0).Width = 250
    End Sub
    Private Sub com2()
        Dim dbcx As New DbAllDataContext
        If ComboBox3.Visible = True Then
            Dim ds = From x In dbcx.Сотрудники.AsEnumerable
                     Where x.НазвОрганиз = ComboBox2.Text
                     Order By x.ФИОСборное
                     Select x.ФИОСборное, x.КодСотрудники
            ComboBox3.DataSource = ds.ToList
            ComboBox3.DisplayMember = "ФИОСборное"
            ComboBox3.ValueMember = "КодСотрудники"

        End If

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        com2()

    End Sub

    Private Sub Btn6()
        For Each tc In TabControl1.Controls.OfType(Of TabPage)
            For Each tb In tc.Controls.OfType(Of TextBox)
                tb.Text = ""
            Next
            For Each rk In tc.Controls.OfType(Of RichTextBox)
                rk.Text = ""
            Next

            For Each mk In tc.Controls.OfType(Of MaskedTextBox)
                mk.Text = ""
            Next

            For Each gb In tc.Controls.OfType(Of GroupBox)
                For Each txt In gb.Controls.OfType(Of TextBox)
                    txt.Text = ""
                Next
            Next


        Next 'таб1
        CheckBox1.Checked = False
        Dim dt As New DataTable
        dt.Columns.Add("ФИО")
        dt.Columns.Add("Пол")
        dt.Columns.Add("Дата рождения")
        Grid1.DataSource = dt

    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Btn6()

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        sender.text = StrConv(sender.text, VbStrConv.ProperCase)
        sender.SelectionStart = sender.text.Length
        TextBox6.Text = TextBox1.Text
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        sender.text = StrConv(sender.text, VbStrConv.ProperCase)
        sender.SelectionStart = sender.text.Length
        TextBox5.Text = TextBox2.Text
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        sender.text = StrConv(sender.text, VbStrConv.ProperCase)
        sender.SelectionStart = sender.text.Length
        TextBox4.Text = TextBox3.Text
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        sender.text = StrConv(sender.text, VbStrConv.ProperCase)
        sender.SelectionStart = sender.text.Length
        TextBox9.Text = TextBox6.Text
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        sender.text = StrConv(sender.text, VbStrConv.ProperCase)
        sender.SelectionStart = sender.text.Length
        TextBox8.Text = TextBox5.Text
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        sender.text = StrConv(sender.text, VbStrConv.ProperCase)
        sender.SelectionStart = sender.text.Length
        TextBox7.Text = TextBox4.Text
    End Sub

    Private Sub TextBox21_TextChanged(sender As Object, e As EventArgs) Handles TextBox21.TextChanged
        TextBox20.Text = TextBox21.Text
    End Sub

    Private Sub TextBox13_TextChanged(sender As Object, e As EventArgs) Handles TextBox13.TextChanged
        TextBox13.Text = TextBox13.Text.ToUpper()
        TextBox13.Select(TextBox13.Text.Length, 0)
        If CheckBox1.Checked = False Then
            TextBox13.MaxLength = 2
            If TextBox13.Text.Length < 2 Then
                TextBox13.ForeColor = Color.Red
            Else
                TextBox13.ForeColor = Color.Green
            End If
        Else
            TextBox13.MaxLength = 25
            TextBox13.ForeColor = Color.Black
        End If

    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged
        TextBox10.Text = TextBox10.Text.ToUpper()
        TextBox10.Select(TextBox10.Text.Length, 0)
        If CheckBox1.Checked = False Then
            TextBox10.MaxLength = 7
            If TextBox10.Text.Length < 7 Then
                TextBox10.ForeColor = Color.Red
            Else
                TextBox10.ForeColor = Color.Green
            End If

        Else
            TextBox10.MaxLength = 25
            TextBox10.ForeColor = Color.Black
        End If


    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged
        TextBox11.Text = TextBox11.Text.ToUpper()
        TextBox11.Select(TextBox11.Text.Length, 0)


        If CheckBox1.Checked = False Then
            TextBox11.MaxLength = 14
            If TextBox11.Text.Length < 14 Then
                TextBox11.ForeColor = Color.Red
            Else
                TextBox11.ForeColor = Color.Green
            End If
        Else
            TextBox11.MaxLength = 25
            TextBox11.ForeColor = Color.Black
        End If
        TextBox45.Text = TextBox11.Text
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Dim _bs As New BindingSource()

        If Проверка() = 1 Then
            Exit Sub
        End If

        If MessageBox.Show("Добавить?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.Cancel Then
            Exit Sub
        End If

        Dim txt12 As String = RTrim(TextBox12.Text)
        txt12 = LTrim(txt12)
        txt12 = txt12.Replace("  ", " ")
        txt12 = Trim(txt12)
        txt12 = StrConv(txt12, VbStrConv.ProperCase)



        If Grid1.Rows.Count = 0 Then

            'Grid1.Columns.Add("_ФИО", "ФИО")
            'Grid1.Columns.Add("_Пол", "Пол")
            'Grid1.Columns.Add("_Дата рождения", "Дата рождения")

            Dim dt As New DataTable
            dt.Columns.Add("ФИО")
            dt.Columns.Add("Пол")
            dt.Columns.Add("Дата рождения")

            Dim row As DataRow = dt.NewRow
            row("ФИО") = txt12
            row("Пол") = ComboBox1.Text
            If MaskedTextBox3.MaskCompleted = False Then
                MessageBox.Show("Введите правильно дату!", Рик)
                Exit Sub
            Else
                'Grid1.Rows.Add(txt12, ComboBox1.Text, MaskedTextBox3.Text)
                row("Дата рождения") = MaskedTextBox3.Text
            End If
            dt.Rows.Add(row)
            '_bs.DataSource = dt

            Grid1.DataSource = dt
            'BindingNavigator1.BindingSource = _bs

            GridView(Grid1)
            Grid1.Columns(0).Width = 250
        Else
            'Dim _bs1 As New BindingSource()
            Dim f As New DataTable
            f.Columns.Add("ФИО")
            f.Columns.Add("Пол")
            f.Columns.Add("Дата рождения")

            For Each row1 As DataGridViewRow In Grid1.Rows
                Dim rw As DataRow = f.NewRow
                For Each cell As DataGridViewCell In row1.Cells
                    If cell.ColumnIndex = 2 Then
                        rw(cell.ColumnIndex) = Strings.Left(cell.Value, 10)
                    Else
                        rw(cell.ColumnIndex) = cell.Value
                    End If

                Next
                f.Rows.Add(rw)
            Next

            'f = DirectCast(Grid1.DataSource, DataTable)
            'f = _bs1.DataSource
            Dim row As DataRow = f.NewRow
            row("ФИО") = txt12
            row("Пол") = ComboBox1.Text
            If MaskedTextBox3.MaskCompleted = False Then
                MessageBox.Show("Введите правильно дату рождения!", Рик)
                Exit Sub
            Else
                row("Дата рождения") = MaskedTextBox3.Text
            End If
            f.Rows.Add(row)

            '_bs.DataSource = f

            Grid1.DataSource = f
            'BindingNavigator1.BindingSource = _bs

            GridView(Grid1)
            Grid1.Columns(0).Width = 250
        End If
        clear()
    End Sub
    Private Function Проверка()
        If ComboBox1.Text = "" Then
            MessageBox.Show("Выберите поле 'Пол'!", Рик)
            Return 1
        End If
        If MaskedTextBox3.Text = "" Then
            MessageBox.Show("Заполните поле 'Дата'!", Рик)
            Return 1
        End If
        If TextBox12.Text = "" Then
            MessageBox.Show("Заполните поле 'ФИО'!", Рик)
            Return 1
        End If
        Return 0
    End Function
    'Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged
    '    sender.text = StrConv(sender.text, VbStrConv.ProperCase)
    '    sender.SelectionStart = sender.text.Length
    'End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click


        If Grid1 IsNot Nothing And Grid1.CurrentRow IsNot Nothing Then
            If Grid1.CurrentRow.Index < 0 Then
                MessageBox.Show("Выберите объект для удаления!", Рик)
                Exit Sub
            End If
        Else
            Exit Sub
        End If

        If MessageBox.Show("Удалить?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.Cancel Then
            Exit Sub
        End If
        Grid1.Rows.RemoveAt(Grid1.CurrentRow.Index)
        clear()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Проверка() = 1 Then
            Exit Sub
        End If
        If Grid1.CurrentRow.Index < 0 Then
            MessageBox.Show("Выберите объект для изменения!", Рик)
            Exit Sub
        End If
        If MessageBox.Show("Изменить?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.Cancel Then
            Exit Sub
        End If


        Grid1.CurrentRow.Cells(0).Value = Trim(StrConv(TextBox12.Text, VbStrConv.ProperCase))
        Grid1.CurrentRow.Cells(1).Value = ComboBox1.Text
        Grid1.CurrentRow.Cells(2).Value = MaskedTextBox3.Text
        clear()
    End Sub

    Private Sub Grid1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellClick
        If Grid1.CurrentRow.Index < 0 Then
            Exit Sub
        End If
        TextBox12.Text = Grid1.CurrentRow.Cells(0).Value
        ComboBox1.Text = Grid1.CurrentRow.Cells(1).Value
        MaskedTextBox3.Text = Grid1.CurrentRow.Cells(2).Value

    End Sub
    Private Sub clear()
        TextBox12.Text = ""
        ComboBox1.Text = ""
        MaskedTextBox3.Text = ""
    End Sub
    Private Function ГлавнаяПроверка()
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Then
            MessageBox.Show("Введите ФИО сотрудника!", Рик)
            Return 1
        End If
        If TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox5.Text = "" Or TextBox7.Text = "" Or TextBox8.Text = "" Or TextBox9.Text = "" Then
            MessageBox.Show("Заполните раздел 'Склонение ФИО'!", Рик)
            Return 1
        End If


        If TextBox21.Text = "" Or TextBox20.Text = "" Then
            MessageBox.Show("Заполните поля 'Прописка' и 'Проживание'!", Рик)
            Return 1
        End If

        If CheckBox1.Checked = False Then
            If TextBox13.ForeColor = Color.Red Then
                MessageBox.Show("Заполните поле 'Серия паспорта'!", Рик)
                Return 1
            End If
            If TextBox10.ForeColor = Color.Red Then
                MessageBox.Show("Заполните поле 'Номер паспорта'!", Рик)
                Return 1
            End If
        End If

        If MaskedTextBox1.MaskCompleted = False Then
            MessageBox.Show("Заполните корректно раздел Паспорт поле 'Дата выдачи'!", Рик)
            Return 1
        End If

        If MaskedTextBox2.MaskCompleted = False Then
            MessageBox.Show("Заполните корректно раздел Паспорт поле 'Срок действия'!", Рик)
            Return 1
        End If

        If RichTextBox1.Text = "" Then
            MessageBox.Show("Заполните корректно раздел Паспорт поле 'Кем выдан'!", Рик)
            Return 1
        End If

        If TextBox11.Text = "" Or TextBox45.Text = "" Or TextBox11.ForeColor = Color.Red Then
            MessageBox.Show("Заполните корректно раздел Паспорт " & vbCrLf & "поле 'Идентификационный номер' или 'Страховой номер'!", Рик)
            Return 1
        End If



    End Function
    Private Sub InsertInfo()
        Dim idNew As Integer
        Using dbcx = New DbAllDataContext  'мой insert
            Dim f As New Сотрудники()
            With f
                .НазвОрганиз = ComboBox2.Text
                .Фамилия = Trim(TextBox1.Text)
                .Имя = Trim(TextBox2.Text)
                .Отчество = Trim(TextBox3.Text)
                .ДанныеИзСправочника = "True"
                .Пол = ComboBox28.Text
                .ФИОСборное = Trim(TextBox1.Text) & " " & Trim(TextBox2.Text) & " " & Trim(TextBox3.Text)
                .ФамилияРодПад = Trim(TextBox6.Text)
                .ИмяРодПад = Trim(TextBox5.Text)
                .ОтчествоРодПад = Trim(TextBox4.Text)
                .ФИОРодПод = Trim(TextBox6.Text) & " " & Trim(TextBox5.Text) & " " & Trim(TextBox4.Text)
                .ФамилияДляЗаявления = Trim(TextBox9.Text)
                .ИмяДляЗаявления = Trim(TextBox8.Text)
                .ОтчествоДляЗаявления = Trim(TextBox7.Text)
                .Гражданин = Trim(TextBox51.Text)
                .Регистрация = Trim(TextBox21.Text)
                .МестоПрожив = Trim(TextBox20.Text)
                .КонтТелГор = Trim(TextBox37.Text)
                .КонтТелефон = MaskedTextBox10.Text
                If CheckBox1.Checked = False Then
                    .Иностранец = "False"
                Else
                    .Иностранец = "True"
                End If
                .ПаспортСерия = Trim(TextBox13.Text)
                .ПаспортНомер = Trim(TextBox10.Text)
                .ПаспортКогдаВыдан = MaskedTextBox1.Text
                .ДоКакогоДейств = MaskedTextBox2.Text
                .ПаспортКемВыдан = RichTextBox1.Text
                .ИДНомер = Trim(TextBox11.Text)
                .СтраховойПолис = Trim(TextBox45.Text)
                .ДатаРожд = MaskedTextBox9.Text

            End With
            dbcx.Сотрудники.InsertOnSubmit(f)
            dbcx.SubmitChanges()
            idNew = f.КодСотрудники
        End Using

        Using dbcx = New DbAllDataContext 'мой insert
            Dim d As New СоставСемьи()
            With d
                .IDСотр = idNew
                .ФИО = Trim(TextBox24.Text)
                .МестоРаботы = Trim(TextBox23.Text)
                .Телефон = Trim(TextBox19.Text)
                If Grid1.Rows.Count > 0 Then
                    .КолДетей = Grid1.Rows.Count
                Else
                    .КолДетей = "Нет"
                End If
            End With
            dbcx.СоставСемьи.InsertOnSubmit(d)
            dbcx.SubmitChanges()
        End Using


        If Grid1.Rows.Count > 0 Then
            For x As Integer = 0 To Grid1.Rows.Count - 1
                Using dbcx = New DbAllDataContext 'мой insert
                    Dim v As New Дети()
                    With v
                        .IDСотр = idNew
                        .ФИО = Grid1.Rows(x).Cells(0).Value
                        .Пол = Grid1.Rows(x).Cells(1).Value
                        .ДатаРождения = CDate(Grid1.Rows(x).Cells(2).Value)
                    End With
                    dbcx.Дети.InsertOnSubmit(v)
                    dbcx.SubmitChanges()
                End Using
            Next
        End If
        Btn6()
        MessageBox.Show("Данные внесены!", Рик)
    End Sub
    Private Sub UpdateInfo()

        Using dbcx = New DbAllDataContext              'мой insert
            Dim var = (From x In dbcx.Сотрудники.AsEnumerable
                       Where x.КодСотрудники = ComboBox3.SelectedValue
                       Select x).Single

            If var IsNot Nothing Then
                With var
                    .НазвОрганиз = ComboBox2.Text
                    .Фамилия = Trim(TextBox1.Text)
                    .Имя = Trim(TextBox2.Text)
                    .Отчество = Trim(TextBox3.Text)
                    .ДанныеИзСправочника = "True"
                    .Пол = ComboBox28.Text
                    .ФИОСборное = Trim(TextBox1.Text) & " " & Trim(TextBox2.Text) & " " & Trim(TextBox3.Text)
                    .ФамилияРодПад = Trim(TextBox6.Text)
                    .ИмяРодПад = Trim(TextBox5.Text)
                    .ОтчествоРодПад = Trim(TextBox4.Text)
                    .ФИОРодПод = Trim(TextBox6.Text) & " " & Trim(TextBox5.Text) & " " & Trim(TextBox4.Text)
                    .ФамилияДляЗаявления = Trim(TextBox9.Text)
                    .ИмяДляЗаявления = Trim(TextBox8.Text)
                    .ОтчествоДляЗаявления = Trim(TextBox7.Text)
                    .Гражданин = Trim(TextBox51.Text)
                    .Регистрация = Trim(TextBox21.Text)
                    .МестоПрожив = Trim(TextBox20.Text)
                    .КонтТелГор = Trim(TextBox37.Text)
                    .КонтТелефон = MaskedTextBox10.Text
                    If CheckBox1.Checked = False Then
                        .Иностранец = "False"
                    Else
                        .Иностранец = "True"
                    End If
                    .ПаспортСерия = Trim(TextBox13.Text)
                    .ПаспортНомер = Trim(TextBox10.Text)
                    .ПаспортКогдаВыдан = MaskedTextBox1.Text
                    .ДоКакогоДейств = MaskedTextBox2.Text
                    .ПаспортКемВыдан = RichTextBox1.Text
                    .ИДНомер = Trim(TextBox11.Text)
                    .СтраховойПолис = Trim(TextBox45.Text)
                    .ДатаРожд = MaskedTextBox9.Text
                    dbcx.SubmitChanges()
                End With
            End If
        End Using

        Using dbcx = New DbAllDataContext 'мой update
            Dim var = (From x In dbcx.СоставСемьи.AsEnumerable
                       Where x.IDСотр = ComboBox3.SelectedValue
                       Select x).SingleOrDefault
            If var IsNot Nothing Then
                With var
                    .ФИО = Trim(TextBox24.Text)
                    .МестоРаботы = Trim(TextBox23.Text)
                    .Телефон = Trim(TextBox19.Text)
                    If Grid1.Rows.Count > 0 Then
                        .КолДетей = Grid1.Rows.Count
                    Else
                        .КолДетей = "Нет"
                    End If
                End With
                dbcx.SubmitChanges()
            Else
                Using dbx As New DbAllDataContext
                    Dim var2 As New СоставСемьи()
                    With var2
                        .IDСотр = ComboBox3.SelectedValue
                        .ФИО = Trim(TextBox24.Text)
                        .МестоРаботы = Trim(TextBox23.Text)
                        .Телефон = Trim(TextBox19.Text)
                        If Grid1.Rows.Count > 0 Then
                            .КолДетей = Grid1.Rows.Count
                        Else
                            .КолДетей = "Нет"
                        End If
                    End With

                    dbcx.СоставСемьи.InsertOnSubmit(var2)
                    dbcx.SubmitChanges()
                End Using
            End If
        End Using


        Using dbcx = New DbAllDataContext  'удаляем старые данные из таблицы дети
            Dim var = (From x In dbcx.Дети.AsEnumerable
                       Where x.IDСотр = ComboBox3.SelectedValue
                       Select x).ToList
            If var.Count > 0 Then
                For Each item In var
                    dbcx.Дети.DeleteOnSubmit(item)
                    dbcx.SubmitChanges()
                Next

            End If
            'If var Is Nothing Then

            'End If

        End Using





        If Grid1.Rows.Count > 0 Then 'вставляем данные в таблицу дети
            For x As Integer = 0 To Grid1.Rows.Count - 1
                Using dbcx = New DbAllDataContext 'мой insert
                    Dim v As New Дети()
                    With v
                        .IDСотр = ComboBox3.SelectedValue
                        .ФИО = Grid1.Rows(x).Cells(0).Value
                        .Пол = Grid1.Rows(x).Cells(1).Value
                        .ДатаРождения = CDate(Grid1.Rows(x).Cells(2).Value)
                    End With
                    dbcx.Дети.InsertOnSubmit(v)
                    dbcx.SubmitChanges()
                End Using
            Next
        End If
        Btn6()
        MessageBox.Show("Данные внесены!", Рик)







    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ГлавнаяПроверка() = 1 Then
            Exit Sub
        End If

        If ComboBox3.Visible = False Then
            InsertInfo()
        Else
            UpdateInfo()
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
            Me.TextBox3.Focus()
        End If
    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            ComboBox28.Focus()
        End If
    End Sub

    Private Sub ComboBox28_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox28.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox6.Focus()
        End If
    End Sub

    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox4.Focus()
        End If
    End Sub

    Private Sub TextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox5.Focus()
        End If
    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
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
            Me.TextBox7.Focus()
        End If
    End Sub

    Private Sub TextBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox7.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox21.Focus()
        End If
    End Sub

    Private Sub TextBox21_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox21.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox20.Focus()
        End If
    End Sub

    Private Sub TextBox20_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox20.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox37.Focus()
        End If
    End Sub

    Private Sub TextBox37_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox37.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            MaskedTextBox10.Focus()
        End If
    End Sub

    Private Sub MaskedTextBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox10.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox51.Focus()
        End If
    End Sub

    Private Sub TextBox51_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox51.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox13.Focus()
        End If
    End Sub

    Private Sub TextBox13_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox13.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox10.Focus()
        End If
    End Sub

    Private Sub TextBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox10.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            MaskedTextBox1.Focus()
        End If
    End Sub

    Private Sub MaskedTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            MaskedTextBox2.Focus()
        End If
    End Sub

    Private Sub MaskedTextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            RichTextBox1.Focus()
        End If
    End Sub

    Private Sub RichTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            TextBox11.Focus()
        End If
    End Sub

    Private Sub TextBox11_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox11.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            TextBox45.Focus()
        End If
    End Sub

    Private Sub TextBox45_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox45.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            MaskedTextBox9.Focus()
        End If
    End Sub

    Private Sub MaskedTextBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox9.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            TextBox24.Focus()
        End If
    End Sub

    Private Sub TextBox24_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox24.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            TextBox23.Focus()
        End If
    End Sub

    Private Sub TextBox23_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox23.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            TextBox19.Focus()
        End If
    End Sub

    Private Sub TextBox19_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox19.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            TextBox12.Focus()
        End If
    End Sub

    Private Sub TextBox12_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox12.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            MaskedTextBox3.Focus()
        End If
    End Sub

    Private Sub MaskedTextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            ComboBox1.Focus()
        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

        Dim fd = sender
        Dim fd1 = e

        If Not ComboBox3.ValueMember <> "" Then
            Exit Sub
        End If

        Dim dbc As New DbAllDataContext
        Dim var = (From x In dbc.Сотрудники.AsEnumerable
                   Where CType(x.КодСотрудники.ToString, Integer) = ComboBox3.SelectedValue
                   Select x).FirstOrDefault()




        TextBox1.Text = var.Фамилия
        TextBox2.Text = var.Имя
        TextBox3.Text = var.Отчество
        ComboBox28.Text = var.Пол

        TextBox4.Text = var.ОтчествоРодПад
        TextBox5.Text = var.ИмяРодПад
        TextBox6.Text = var.ФамилияРодПад

        TextBox7.Text = var.ОтчествоДляЗаявления
        TextBox8.Text = var.ИмяДляЗаявления
        TextBox9.Text = var.ФамилияДляЗаявления

        TextBox21.Text = var.Регистрация
        TextBox20.Text = var.МестоПрожив
        TextBox37.Text = var.КонтТелГор
        MaskedTextBox10.Text = var.КонтТелефон
        TextBox51.Text = var.Гражданин

        If var.Иностранец = True Then
            CheckBox1.Checked = True
        Else
            CheckBox1.Checked = False
        End If

        TextBox13.Text = var.ПаспортСерия
        TextBox10.Text = var.ПаспортНомер
        MaskedTextBox1.Text = var.ПаспортКогдаВыдан
        MaskedTextBox2.Text = var.ДоКакогоДейств
        RichTextBox1.Text = var.ПаспортКемВыдан
        TextBox11.Text = var.ИДНомер
        TextBox45.Text = var.СтраховойПолис
        MaskedTextBox9.Text = var.ДатаРожд



        Dim dbcx As New DbAllDataContext
        Dim var1 = (From x In dbc.СоставСемьи.AsEnumerable
                    Where x.IDСотр = ComboBox3.SelectedValue
                    Select x).FirstOrDefault()
        If var1 IsNot Nothing Then
            TextBox24.Text = var1.ФИО
            TextBox23.Text = var1.МестоРаботы
            TextBox19.Text = var1.Телефон
        End If



        Dim dbcx1 As New DbAllDataContext
        Dim var2 = From x In dbc.Дети.AsEnumerable
                   Where x.IDСотр = ComboBox3.SelectedValue
                   Select x.ФИО, x.Пол, x.ДатаРождения

        If var2 IsNot Nothing Then
            Grid1.DataSource = var2.ToList()
            GridView(Grid1)
        End If





    End Sub



    'Pri
    'Pvate Sub Grid1_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles Grid1.CellValidating
    '    'Dim sep As Char = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator
    '    If e.ColumnIndex = 2 And e.FormattedValue <> "" Then
    '        'If sep = "," Then
    '        '    'если в качестве десятичного разделителя использована точка (а десятичный разделитель - запятая), то меняем её на запятую
    '        '    sender.Tag = Format(Convert.ToSingle(Replace(e.FormattedValue, ".", sep)), "00.00.0000")
    '        'Else 'sep = "."
    '        '    'если в качестве десятичного разделителя использована запятая (а десятичный разделитель - точка), то меняем её на точку
    '        '    sender.Tag = Format(Convert.ToSingle(Replace(e.FormattedValue, ",", sep)), "00.00.0000")
    '        'End If
    '        Dim f As String = e.FormattedValue
    '        If f.Length > 10 Then
    '            f = Strings.Left(f.ToString, 10)
    '        End If

    '        f = Replace(f, "/", ".")
    '        f = Replace(f, ",", ".")
    '        Try
    '            Dim d As Date = CDate(f)
    '            If d > Now.Date Then
    '                MessageBox.Show("Введите корректную дату!", Рик)
    '                'Grid1.Rows(e.RowIndex).Cells(2).Value = ""

    '                f = "z"
    '            End If
    '        Catch ex As Exception
    '            MessageBox.Show("Введите корректную дату!", Рик)
    '            'Grid1.Rows(Grid1.CurrentRow.Index).Cells(2).Value = ""
    '            f = "z"
    '        End Try

    '        sender.Tag = f

    '    Else
    '        sender.Tag = e.FormattedValue
    '    End If


    '    If e.ColumnIndex = 1 And e.FormattedValue <> "" Then
    '        If e.FormattedValue = "м" Or e.FormattedValue = "ж" Or e.FormattedValue = "М" Or e.FormattedValue = "Ж" Then
    '            sender.Tag = e.FormattedValue
    '        Else
    '            MessageBox.Show("Введите корректно пол ребенка!", Рик)
    '            sender.Tag = ""
    '        End If
    '    End If

    '    If e.ColumnIndex = 0 And e.FormattedValue <> "" Then
    '        Dim h As String
    '        h = e.FormattedValue
    '        sender.Tag = Trim(StrConv(h, VbStrConv.ProperCase))
    '        'sender.SelectionStart = sender.text.Length

    '    End If





    'End Sub

    'Private Sub Grid1_CellValidated(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellValidated


    '    'столбец 2
    '    If sender.Tag = "z" And e.ColumnIndex = 2 Then
    '        sender.Tag = ""
    '        sender(e.ColumnIndex, e.RowIndex).Value = sender.Tag
    '    ElseIf sender.Tag <> "" Then
    '        sender(e.ColumnIndex, e.RowIndex).Value = sender.Tag
    '    End If

    '    'столбец 1
    '    If sender.Tag = "" And e.ColumnIndex = 1 Then
    '        sender.Tag = ""
    '        sender(e.ColumnIndex, e.RowIndex).Value = sender.Tag
    '    Else
    '        sender(e.ColumnIndex, e.RowIndex).Value = sender.Tag
    '    End If

    '    'столбец 0 
    '    If sender.Tag <> "" And e.ColumnIndex = 0 Then
    '        sender(e.ColumnIndex, e.RowIndex).Value = sender.Tag
    '    End If



    'End Sub
End Class