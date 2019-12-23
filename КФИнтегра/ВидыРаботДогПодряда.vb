Imports System.Threading
Imports System.ComponentModel

Public Class ВидыРаботДогПодряда
    Public Прием10 As Прием
    Dim er As Integer = 0
    Dim закрКрестик As Boolean = True

    Private Function Пров() As Integer
        Dim strsql As String = "SELECT СтоимЧасаРуб FROM ДогПодряда WHERE ID=" & CType(Прием.Label96.Text, Integer) & ""

        Dim ds As DataTable = Selects(strsql)
        If ds.Rows(0).Item(0).ToString <> "" Then
            Return 1
        Else
            Return 0
        End If


    End Function
    Private Sub ВидыРаботДогПодряда_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim ut() As Object = {"м2", "м3", "м.п."}
        КрестикНажатиеДогПодряда = False
        If ДогПодрВклЧекбокс5 = False Then
            Try
                ComboBox1.Items.Clear()
                ComboBox2.Items.Clear()
                ComboBox5.Items.Clear()
                ComboBox7.Items.Clear()
                ComboBox9.Items.Clear()
            Catch ex As Exception

            End Try
            ComboBox1.Items.AddRange(ut)
            ComboBox2.Items.AddRange(ut)
            ComboBox5.Items.AddRange(ut)
            ComboBox7.Items.AddRange(ut)
            ComboBox9.Items.AddRange(ut)
            GroupBox3.Enabled = False
            GroupBox4.Enabled = False
            GroupBox5.Enabled = False
            GroupBox6.Enabled = False
        Else
            'If Пров() = 1 Then
            'Try
            '    ComboBox1.Items.Clear()
            '    ComboBox2.Items.Clear()
            '    ComboBox5.Items.Clear()
            '    ComboBox7.Items.Clear()
            '    ComboBox9.Items.Clear()
            'Catch ex As Exception

            'End Try
            'ComboBox1.Items.AddRange(ut)
            'ComboBox2.Items.AddRange(ut)
            'ComboBox5.Items.AddRange(ut)
            'ComboBox7.Items.AddRange(ut)
            'ComboBox9.Items.AddRange(ut)
            'GroupBox3.Enabled = False
            'GroupBox4.Enabled = False
            'GroupBox5.Enabled = False
            'GroupBox6.Enabled = False
            'End If
        End If

        If Дпод1 = "выполненных работ" Then
            Try
                ComboBox1.Items.Clear()
                ComboBox2.Items.Clear()
                ComboBox5.Items.Clear()
                ComboBox7.Items.Clear()
                ComboBox9.Items.Clear()
            Catch ex As Exception

            End Try
            GroupBox3.Enabled = False
            GroupBox4.Enabled = False
            GroupBox5.Enabled = False
            GroupBox6.Enabled = False



            ComboBox1.Items.AddRange(ut)
            ComboBox2.Items.AddRange(ut)
            ComboBox5.Items.AddRange(ut)
            ComboBox7.Items.AddRange(ut)
            ComboBox9.Items.AddRange(ut)

            CheckBox2.Checked = False
            CheckBox1.Checked = False
            CheckBox3.Checked = False
            CheckBox4.Checked = False


            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""
            TextBox6.Text = ""
            TextBox7.Text = ""
            TextBox8.Text = ""

            ComboBox1.Text = ""
            ComboBox2.Text = ""
            ComboBox5.Text = ""
            ComboBox7.Text = ""
            ComboBox9.Text = ""
            ComboBox3.Text = String.Empty
            ComboBox4.Text = String.Empty
            ComboBox6.Text = ""
            ComboBox8.Text = ""
            ComboBox9.Text = ""

        End If

    End Sub
    Public Async Sub ОчисВидыРаботДогПодряда()
        'Await Task.Delay(20)
        'Me.CheckBox2.Checked = True
        'Me.CheckBox1.Checked = True
        'Me.CheckBox3.Checked = True
        'Me.CheckBox4.Checked = True
        'Me.CheckBox2.Checked = False
        'Me.CheckBox1.Checked = False
        'Me.CheckBox3.Checked = False
        'Me.CheckBox4.Checked = False
        'Me.Очистка(Me)
    End Sub
    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            GroupBox3.Enabled = True
            TextBox2.Text = "00"
        Else
            GroupBox3.Enabled = False
            TextBox1.Text = ""
            TextBox2.Text = ""
            ComboBox2.Text = ""
            ComboBox3.Text = ""
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            GroupBox4.Enabled = True
            TextBox3.Text = "00"
        Else
            GroupBox4.Enabled = False
            TextBox3.Text = ""
            TextBox4.Text = ""
            ComboBox4.Text = ""
            ComboBox5.Text = ""
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            GroupBox5.Enabled = True
            TextBox5.Text = "00"
        Else
            GroupBox5.Enabled = False
            TextBox5.Text = ""
            TextBox6.Text = ""
            ComboBox7.Text = ""
            ComboBox6.Text = ""
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            GroupBox6.Enabled = True
            TextBox7.Text = "00"
        Else
            GroupBox6.Enabled = False
            TextBox7.Text = ""
            TextBox8.Text = ""
            ComboBox8.Text = ""
            ComboBox9.Text = ""
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ComboBox1.Text = "" Then
            MessageBox.Show("Выберите еденицу измерения!", Рик)
            Exit Sub
        End If
        If RichTextBox1.Text = "" Then
            MessageBox.Show("Заполните поле выполняемых работ!", Рик)
            Exit Sub
        End If

        Dim inp As String = InputBox("Введите вид работы " & Trim(RichTextBox1.Text) & " в Именительном падеже", Рик)
        Do Until inp <> ""
            MessageBox.Show("Повторите ввод!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Error)
            inp = InputBox("Введите вид работы " & Trim(RichTextBox1.Text) & " в Именительном падеже", Рик)
        Loop

        Dim strsql As String = "INSERT INTO ДогПодОсобен(ЕденицаИзм,Текст,ТесктИменПад) VALUES('" & ComboBox1.Text & "','" & Trim(RichTextBox1.Text) & "',
'" & inp & "')"
        Updates(strsql)

        MessageBox.Show("Данные добавлены!", Рик)
        ComboBox1.Text = ""
        RichTextBox1.Text = ""
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim strsql As String = "SELECT * FROM ДогПодОсобен WHERE ЕденицаИзм='" & ComboBox2.Text & "' ORDER BY Текст"
        Dim ds As DataTable = Selects(strsql)

        ComboBox3.AutoCompleteCustomSource.Clear()
        ComboBox3.Items.Clear()
        For Each r As DataRow In ds.Rows
            ComboBox3.AutoCompleteCustomSource.Add(r.Item(2).ToString())
            ComboBox3.Items.Add(r(2).ToString)
        Next
        If ds.Rows.Count > 0 Then
            ComboBox3.Text = ds.Rows(0).Item(2).ToString
        Else
            ComboBox3.Text = ""
        End If
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        Dim strsql As String = "SELECT * FROM ДогПодОсобен WHERE ЕденицаИзм='" & ComboBox5.Text & "' ORDER BY Текст"
        Dim ds As DataTable = Selects(strsql)
        ComboBox4.AutoCompleteCustomSource.Clear()
        ComboBox4.Items.Clear()
        For Each r As DataRow In ds.Rows
            ComboBox4.AutoCompleteCustomSource.Add(r.Item(2).ToString())
            ComboBox4.Items.Add(r(2).ToString)
        Next
        If ds.Rows.Count > 0 Then
            ComboBox4.Text = ds.Rows(0).Item(2).ToString
        Else
            ComboBox4.Text = ""
        End If
    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        Dim strsql As String = "SELECT * FROM ДогПодОсобен WHERE ЕденицаИзм='" & ComboBox7.Text & "' ORDER BY Текст"
        Dim ds As DataTable = Selects(strsql)
        ComboBox6.AutoCompleteCustomSource.Clear()
        ComboBox6.Items.Clear()
        For Each r As DataRow In ds.Rows
            ComboBox6.AutoCompleteCustomSource.Add(r.Item(2).ToString())
            ComboBox6.Items.Add(r(2).ToString)
        Next
        If ds.Rows.Count > 0 Then
            ComboBox6.Text = ds.Rows(0).Item(2).ToString
        Else
            ComboBox6.Text = ""
        End If
    End Sub

    Private Sub ComboBox9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox9.SelectedIndexChanged
        Dim strsql As String = "SELECT * FROM ДогПодОсобен WHERE ЕденицаИзм='" & ComboBox9.Text & "' ORDER BY Текст"
        Dim ds As DataTable = Selects(strsql)
        ComboBox8.AutoCompleteCustomSource.Clear()
        ComboBox8.Items.Clear()
        For Each r As DataRow In ds.Rows
            ComboBox8.AutoCompleteCustomSource.Add(r.Item(2).ToString())
            ComboBox8.Items.Add(r(2).ToString)
        Next
        If ds.Rows.Count > 0 Then
            ComboBox8.Text = ds.Rows(0).Item(2).ToString
        Else
            ComboBox8.Text = ""
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
            ComboBox2.Focus()
        End If
    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox3.Focus()
        End If
    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            ComboBox5.Focus()
        End If
    End Sub

    Private Sub TextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox5.Focus()
        End If
    End Sub

    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            ComboBox7.Focus()
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
            ComboBox9.Focus()
        End If
    End Sub
    Private Sub ОпредИзменДействИлиНов()
        Dim пр As New Прием 'обращение к первой форме контролов
        If ДогПодномДогПодСтДог <> "" Then
            If пр.TextBox39.Text <> "" Then
                ДогПодномДогПодСтДог = ДогПодномДогПодСтДог & "." & пр.TextBox39.Text
            End If
        End If

        er = 0
        If ДогПодНомерСтар = ДогПодномДогПодСтДог Then
            If MessageBox.Show("Заменить действующие условия для договора " & ДогПодНомерСтар & " ?", Рик, MessageBoxButtons.YesNo) = DialogResult.No Then
                er = 1
                Exit Sub
            End If
        Else
            If MessageBox.Show("Создать новый договор подряда?", Рик, MessageBoxButtons.YesNo) = DialogResult.Yes Then
                ДогПодрНомНовы = 1
            Else
                er = 1
                Exit Sub
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If CheckBox2.Checked = False And CheckBox1.Checked = False And CheckBox3.Checked = False And CheckBox4.Checked = False Then
            If MessageBox.Show("Вы не выбрали данные для сохранения! Выйти?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                Me.Close()
            Else
                Exit Sub
            End If
        End If

        ОпредИзменДействИлиНов()
        ДогПодномДогПодСтДог = ""
        If er = 1 Then Exit Sub

        If CheckBox2.Checked = True And CheckBox1.Checked = False And CheckBox3.Checked = False And CheckBox4.Checked = False Then
            Дпод2 = "Стоимость 1" & ComboBox2.Text & " " & ComboBox3.Text & " – " & TextBox1.Text & "р. " & TextBox2.Text & "коп."
            ДогПодрВыпРаб = New List(Of String) From {ComboBox3.Text}
            ДогПодрВыпРабСтР = New List(Of String) From {TextBox1.Text}
            ДогПодрВыпРабСтК = New List(Of String) From {TextBox2.Text}
            ДогПодрВыпРабСтОб = New List(Of String) From {ComboBox2.Text}
        ElseIf CheckBox2.Checked = True And CheckBox1.Checked = True And CheckBox3.Checked = False And CheckBox4.Checked = False Then
            Дпод2 = "Стоимость 1" & ComboBox2.Text & " " & ComboBox3.Text & " – " & TextBox1.Text & "р. " & TextBox2.Text & "коп." &
                ", стоимость 1" & ComboBox5.Text & " " & ComboBox4.Text & " – " & TextBox4.Text & "р. " & TextBox3.Text & "коп."
            ДогПодрВыпРаб = New List(Of String) From {ComboBox3.Text, ComboBox4.Text}
            ДогПодрВыпРабСтР = New List(Of String) From {TextBox1.Text, TextBox4.Text}
            ДогПодрВыпРабСтК = New List(Of String) From {TextBox2.Text, TextBox3.Text}
            ДогПодрВыпРабСтОб = New List(Of String) From {ComboBox2.Text, ComboBox5.Text}
        ElseIf CheckBox2.Checked = True And CheckBox1.Checked = True And CheckBox3.Checked = True And CheckBox4.Checked = False Then
            Дпод2 = "Стоимость 1" & ComboBox2.Text & " " & ComboBox3.Text & " – " & TextBox1.Text & "р. " & TextBox2.Text & "коп." &
                ", стоимость 1" & ComboBox5.Text & " " & ComboBox4.Text & " – " & TextBox4.Text & "р. " & TextBox3.Text & "коп." &
                ", стоимость 1" & ComboBox7.Text & " " & ComboBox6.Text & " – " & TextBox6.Text & "р. " & TextBox5.Text & "коп."

            ДогПодрВыпРаб = New List(Of String) From {ComboBox3.Text, ComboBox4.Text, ComboBox6.Text}
            ДогПодрВыпРабСтР = New List(Of String) From {TextBox1.Text, TextBox4.Text, TextBox6.Text}
            ДогПодрВыпРабСтК = New List(Of String) From {TextBox2.Text, TextBox3.Text, TextBox5.Text}
            ДогПодрВыпРабСтОб = New List(Of String) From {ComboBox2.Text, ComboBox5.Text, ComboBox7.Text}
        Else
            Дпод2 = "Стоимость 1" & ComboBox2.Text & " " & ComboBox3.Text & " – " & TextBox1.Text & "р. " & TextBox2.Text & "коп." &
                ", стоимость 1" & ComboBox5.Text & " " & ComboBox4.Text & " – " & TextBox4.Text & "р. " & TextBox3.Text & "коп." &
                ", стоимость 1" & ComboBox7.Text & " " & ComboBox6.Text & " – " & TextBox6.Text & "р. " & TextBox5.Text & "коп." &
                ", стоимость 1" & ComboBox9.Text & " " & ComboBox8.Text & " – " & TextBox8.Text & "р. " & TextBox7.Text & "коп."
            ДогПодрВыпРаб = New List(Of String) From {ComboBox3.Text, ComboBox4.Text, ComboBox6.Text, ComboBox8.Text}
            ДогПодрВыпРабСтР = New List(Of String) From {TextBox1.Text, TextBox4.Text, TextBox6.Text, TextBox8.Text}
            ДогПодрВыпРабСтК = New List(Of String) From {TextBox2.Text, TextBox3.Text, TextBox5.Text, TextBox7.Text}
            ДогПодрВыпРабСтОб = New List(Of String) From {ComboBox2.Text, ComboBox5.Text, ComboBox7.Text, ComboBox9.Text}
        End If


        закрКрестик = False
        Очистка(Me)
        Me.Close()
        закрКрестик = True


    End Sub
    Public Sub Очистка(ByRef F As Form) 'очистка контролов
        For Each F_Control As Control In F.Controls
            Dim _control As Object = F.Controls(F_Control.Name)
            If TypeOf _control Is TextBox Then
                _control.Text = ""
                'ElseIf TypeOf _control Is ListBox Then
                '    _control.items.clear()
            ElseIf TypeOf _control Is ComboBox Then
                _control.selectedindex = -1
            ElseIf TypeOf _control Is RichTextBox Then
                _control.text = ""
            End If
        Next F_Control

    End Sub

    Private Sub ВидыРаботДогПодряда_Closing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.Closing 'закрытие на крестик

        If закрКрестик = True Then
            ДогПодномДогПодСтДог = ""
            ПрЗакрВидыРаб = Com19ForДогПодр
            КрестикНажатиеДогПодряда = True
        End If


        'If e.CloseReason = CloseReason.UserClosing Then
        '    ДогПодномДогПодСтДог = ""
        '    ПрЗакрВидыРаб = Com19ForДогПодр
        '    КрестикНажатиеДогПодряда = True
        'End If

    End Sub

    'Public Sub ВстДанных(ByVal ds As DataTable)

    '    Dim ut() As Object = {"м2", "м3", "м.п."}
    '    Try
    '        ComboBox1.Items.Clear()
    '        ComboBox2.Items.Clear()
    '        ComboBox5.Items.Clear()
    '        ComboBox7.Items.Clear()
    '        ComboBox9.Items.Clear()
    '    Catch ex As Exception

    '    End Try

    '    GroupBox3.Enabled = False
    '    GroupBox4.Enabled = False
    '    GroupBox5.Enabled = False
    '    GroupBox6.Enabled = False



    '    ComboBox1.Items.AddRange(ut)
    '    ComboBox2.Items.AddRange(ut)
    '    ComboBox5.Items.AddRange(ut)
    '    ComboBox7.Items.AddRange(ut)
    '    ComboBox9.Items.AddRange(ut)

    '    'Dim strsql As String = "SELECT * FROM ДогПодОсобен"
    '    'Dim ds1 As DataTable = Selects(strsql)
    '    For i As Integer = 0 To ds.Rows.Count - 1
    '        Select Case i
    '            Case 0
    '                CheckBox2.Checked = True
    '                ComboBox3.Text = ds.Rows(i).Item(11).ToString
    '                ComboBox2.Text = ds.Rows(i).Item(14).ToString
    '                TextBox1.Text = ds.Rows(i).Item(12).ToString
    '                TextBox2.Text = ds.Rows(i).Item(13).ToString
    '                GroupBox3.Enabled = True

    '            Case 1
    '                CheckBox1.Checked = True
    '                ComboBox4.Text = ds.Rows(i).Item(11).ToString
    '                ComboBox5.Text = ds.Rows(i).Item(14).ToString
    '                TextBox4.Text = ds.Rows(i).Item(12).ToString
    '                TextBox3.Text = ds.Rows(i).Item(13).ToString
    '                GroupBox4.Enabled = True

    '            Case 2
    '                CheckBox3.Checked = True
    '                ComboBox6.Text = ds.Rows(i).Item(11).ToString
    '                ComboBox7.Text = ds.Rows(i).Item(14).ToString
    '                TextBox6.Text = ds.Rows(i).Item(12).ToString
    '                TextBox5.Text = ds.Rows(i).Item(13).ToString
    '                GroupBox5.Enabled = True
    '            Case 3
    '                CheckBox4.Checked = True
    '                ComboBox8.Text = ds.Rows(i).Item(11).ToString
    '                ComboBox9.Text = ds.Rows(i).Item(14).ToString
    '                TextBox8.Text = ds.Rows(i).Item(12).ToString
    '                TextBox7.Text = ds.Rows(i).Item(13).ToString
    '                GroupBox6.Enabled = True
    '        End Select
    '    Next

    '    Прием.GroupBox28.Visible = False

    'End Sub


End Class