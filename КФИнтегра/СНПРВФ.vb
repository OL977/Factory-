Public Class СНПРВФ
    Private Sub СНПРВФ_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Start()
        ComboBox1.Items.Clear()
        Dim f() As String = {Now.Year, Now.Year - 1, Now.Year - 2, Now.Year - 3}
        ComboBox1.Items.AddRange(f)


    End Sub
    Private Sub Start()
        Using dbcx As New DbAll1DataContext
            Dim var = (From x In dbcx.СНПРВ.AsEnumerable
                       Order By x.Год
                       Select x).ToList
            If var.Count > 0 Then
                ListBox1.Items.Clear()
                For Each x In var
                    ListBox1.Items.Add(x.Год & "г. - " & x.Норма)
                Next
            End If
        End Using
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or ComboBox1.Text = "" Then
            MessageBox.Show("Заполните поле дата и норма!", Рик)
            Exit Sub
        End If

        TextBox1.Text = Replace(TextBox1.Text, ".", ",")

        If Not IsNumeric(TextBox1.Text) Then
            MessageBox.Show("Введите корректные данные в поле" & vbCrLf & "'Среднемесячная норма, час'!", Рик)
            Exit Sub
        End If

        Dim n = Math.Round(CDbl(TextBox1.Text), 2)

        Dim m As String = CType(n, String)

        Using dbcx As New DbAll1DataContext
            Dim var = (From x In dbcx.СНПРВ.AsEnumerable
                       Where x.Год = ComboBox1.Text
                       Select x).FirstOrDefault
            If var IsNot Nothing Then
                var.Норма = m
                dbcx.SubmitChanges()
            Else
                Dim f As New СНПРВ()
                f.Год = ComboBox1.Text
                f.Норма = m
                dbcx.СНПРВ.InsertOnSubmit(f)
                dbcx.SubmitChanges()
            End If
        End Using





        'Dim list As New Dictionary(Of String, Object)
        'list.Add("@Год", ComboBox1.Text)
        'list.Add("@Норма", m)

        'Dim dt As DataTable = Selects(StrSql:="SELECT * FROM СНПРВ WHERE Год=@Год", list)

        'If dt.Rows.Count > 0 Then
        '    Updates(stroka:="UPDATE СНПРВ SET Норма='" & m & "' WHERE Код=" & dt(0).Item("Код") & "")
        'Else
        '    Updates(stroka:="INSERT INTO СНПРВ(Год,Норма) VALUES(@Год,@Норма)", list)
        'End If

        MessageBox.Show("Данные приняты!", Рик)



        Start()


        TextBox1.Text = ""
        ComboBox1.Text = ""
    End Sub
End Class