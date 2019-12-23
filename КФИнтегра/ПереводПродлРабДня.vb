Imports WindowsApp3.Перевод

Public Class ПереводПродлРабДня
    Private Sub ПереводПродлРабДня_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim fd As New List(Of String)() From {"2.15", "4.30", "9.00", "10.00", "11.00", "12.00"}
        ListBox1.Items.Clear()
        For i As Integer = 0 To fd.Count - 1
            ListBox1.Items.Add(fd(i))
        Next

    End Sub

    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick
        If Not ListBox1.SelectedIndex = -1 Then
            Перевод.ComboBox11.Text = ListBox1.SelectedItem.ToString
            РасчПер()
            Me.Close()
        End If
    End Sub

    Public Sub РасчПер()
        Dim часы As Decimal = Val(Перевод.ComboBox11.Text) 'расчет времени обеда и конца рабочего дня
        Dim ВрНач As Decimal = Val(Перевод.ComboBox10.Text)
        'часы = Math.Floor(часы)
        Select Case часы
            Case 9
                If Not (ВрНач = 8.3 Or ВрНач = 10.3) Then
                    Перевод.TextBox12.Text = Str(часы + ВрНач) & ".00"
                    Dim с As String = Str(4 + ВрНач)
                    Dim по As String = Str(4 + ВрНач + 1)
                    Перевод.TextBox11.Text = "с" & с & ".00 до" & по & ".00"
                Else
                    Перевод.TextBox12.Text = Str(часы + ВрНач) & "0"
                    Dim с As String = Str(4 + ВрНач)
                    Dim по As String = Str(4 + ВрНач + 1)
                    Перевод.TextBox11.Text = "с" & с & "0 до" & по & "0"
                End If

            Case 10
                If Not (ВрНач = 8.3 Or ВрНач = 10.3) Then
                    Перевод.TextBox12.Text = Str(часы + ВрНач) & ".00"
                    Dim с As String = Str(5 + ВрНач)
                    Dim по As String = Str(5 + ВрНач + 1)
                    Перевод.TextBox11.Text = "с" & с & ".00 до" & по & ".00"
                Else
                    Перевод.TextBox12.Text = Str(часы + ВрНач) & "0"
                    Dim с As String = Str(5 + ВрНач)
                    Dim по As String = Str(5 + ВрНач + 1)
                    Перевод.TextBox11.Text = "с" & с & "0 до" & по & "0"
                End If

            Case 11
                If Not (ВрНач = 8.3 Or ВрНач = 10.3) Then
                    Перевод.TextBox12.Text = Str(часы + ВрНач) & ".00"
                    Dim с As String = Str(5 + ВрНач)
                    Dim по As String = Str(5 + ВрНач + 1)
                    Перевод.TextBox11.Text = "с" & с & ".00 до" & по & ".00"
                Else
                    Перевод.TextBox12.Text = Str(часы + ВрНач) & "0"
                    Dim с As String = Str(5 + ВрНач)
                    Dim по As String = Str(5 + ВрНач + 1)
                    Перевод.TextBox11.Text = "с" & с & "0 до" & по & "0"
                End If
            Case 12
                If Not (ВрНач = 8.3 Or ВрНач = 10.3) Then
                    Перевод.TextBox12.Text = Str(часы + ВрНач) & ".00"
                    Dim с As String = Str(6 + ВрНач)
                    Dim по As String = Str(6 + ВрНач + 1)
                    Перевод.TextBox11.Text = "с" & с & ".00 до" & по & ".00"
                Else
                    Перевод.TextBox12.Text = Str(часы + ВрНач) & "0"
                    Dim с As String = Str(6 + ВрНач)
                    Dim по As String = Str(6 + ВрНач + 1)
                    Перевод.TextBox11.Text = "с" & с & "0 до" & по & "0"
                End If
            Case 4.3
                If Not (ВрНач = 8.3 Or ВрНач = 10.3) Then
                    Перевод.TextBox12.Text = Str(часы + ВрНач) & "0"
                    Dim с As String = Str(2 + ВрНач)
                    Dim по As String = Str(2 + ВрНач + 0.3)
                    Перевод.TextBox11.Text = "с" & с & ".00 до" & по & "0"
                Else
                    Select Case ВрНач
                        Case 8.3
                            Перевод.TextBox12.Text = "13.00"
                            Перевод.TextBox11.Text = "с 10.30 по 11.00"
                        Case 10.3
                            Перевод.TextBox12.Text = "15.00"
                            Перевод.TextBox11.Text = "с 12.30 по 13.00"
                    End Select

                End If
            Case 2.15
                If Not (ВрНач = 8.3 Or ВрНач = 10.3) Then
                    Перевод.TextBox12.Text = Str(часы + ВрНач)
                    Dim с As String = Str(1 + ВрНач)
                    Dim по As String = Str(1 + ВрНач + 0.15)
                    Перевод.TextBox11.Text = "с" & с & ".00 до " & по
                Else
                    Select Case ВрНач
                        Case 8.3
                            Перевод.TextBox12.Text = "10.45"
                            Перевод.TextBox11.Text = "с 9.30 по 9.45"
                        Case 10.3
                            Перевод.TextBox12.Text = "12.45"
                            Перевод.TextBox11.Text = "с 11.30 по 11.45"
                    End Select
                End If

        End Select
    End Sub



End Class