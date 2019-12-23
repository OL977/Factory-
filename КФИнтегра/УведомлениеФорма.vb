Public Class УведомлениеФорма
    Dim file2() As String
    Dim FilesList() As String
    Dim ds As DataTable
    Private Sub УведомлениеФорма_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim strsql, орг As String

        If proverka = 0 Then
            strsql = "SELECT Фамилия FROM Сотрудники WHERE НазвОрганиз='" & Уведомление.ComboBox1.Text & "' AND ФИОСборное='" & Уведомление.name2 & "'"
            орг = Уведомление.ComboBox1.Text
        Else
            strsql = "SELECT Фамилия FROM Сотрудники WHERE КодСотрудники=" & CType(Прием.Label96.Text, Integer) & ""
            орг = Прием.ComboBox1.Text
        End If

        ds = Selects(strsql)

        FilesList = IO.Directory.GetFiles(OnePath & орг, "*" & ds.Rows(0).Item(0).ToString & "*.doc", IO.SearchOption.AllDirectories)
        Dim gth4 As String

        file2 = IO.Directory.GetFiles(OnePath & орг, "*" & ds.Rows(0).Item(0).ToString & "*.doc", IO.SearchOption.AllDirectories)


        For n As Integer = 0 To FilesList.Length - 1
            gth4 = ""
            gth4 = IO.Path.GetFileName(file2(n))
            file2(n) = gth4
            'TextBox44.Text &= gth + vbCrLf
        Next

        'ListBox2.Items.Add(Files2)


        ListBox1.Items.Clear()

        For i = 0 To file2.Length - 1 ' Распечатываем весь получившийся массив
            ListBox1.Items.Add(file2(i)) ' На ListBox2
        Next





    End Sub

    Private Sub УведомлениеФорма_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        TextBox15.Text = ""
        TextBox16.Text = ""
        TextBox17.Text = ""
        TextBox18.Text = ""
        TextBox19.Text = ""
        TextBox20.Text = ""
        TextBox21.Text = ""

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub

    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick
        If ListBox1.SelectedIndex = -1 Then
            MessageBox.Show("Выберите документ для просмотра!", Рик, MessageBoxButtons.OK)
            Exit Sub
        End If

        If Not ListBox1.SelectedIndex = -1 Then

            Process.Start(FilesList(ListBox1.SelectedIndex))

        End If
    End Sub
End Class