Option Explicit On
Imports System.ComponentModel
Imports System.Data.OleDb

Public Class ДолжИнстр
    Public Ном As String
    Public Дат As String
    Public текст As String
    Public x As Boolean = False, f As Boolean = False, v As Boolean = False
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
        TextBox1.Text = ""
        MaskedTextBox1.Text = Now.ToShortDateString
        RichTextBox1.Text = ""
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
            RichTextBox1.Focus()
        End If
    End Sub

    Private Sub RichTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Button1.Focus()
        End If
    End Sub

    Private Sub ДолжИнстр_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If x = False Then
            MaskedTextBox1.Text = Now.ToShortDateString
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then
            MessageBox.Show("Выберите номер инструкции!", Рик)
            Exit Sub
        End If
        If MaskedTextBox1.MaskCompleted = False Then
            MessageBox.Show("Выберите правильно дату!", Рик)
            Exit Sub
        End If
        If RichTextBox1.Text = "" Then
            MessageBox.Show("Введите текст в поле инструкции!", Рик)
            Exit Sub
        End If
        Ном = ""
        Дат = ""
        текст = ""
        Ном = TextBox1.Text
        Дат = MaskedTextBox1.Text
        текст = RichTextBox1.Text

        If MessageBox.Show("Данные приняты, продолжить?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.Cancel Then
            Exit Sub
        Else
            Me.Close()
        End If
        Штатное.v = True
        Прием.v = True
        TextBox1.Text = ""
        MaskedTextBox1.Text = Now.ToShortDateString
        RichTextBox1.Text = ""

        Статистика("Нет", "Создание должностной инструкции", Штатное.ComboBox1.Text)
    End Sub

    Private Sub ДолжИнстр_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        TextBox1.Text = ""
        MaskedTextBox1.Text = Now.ToShortDateString
        RichTextBox1.Text = ""
        If Me.UseWaitCursor Then
            Me.Cursor = Cursors.Default
        End If

    End Sub

    Private Sub TextBox1_LostFocus(sender As Object, e As EventArgs) Handles TextBox1.LostFocus

        Dim pl As String
        If TextBox1.Text <> "" And IsNumeric(TextBox1.Text) = True Then
            Dim i As Integer = CInt(TextBox1.Text)
            Select Case i

                Case < 10
                    pl = Str(i)
                    TextBox1.Text = "00" & i

                Case 10 To 99
                    pl = Str(i)
                    TextBox1.Text = "0" & i
            End Select
        Else
            MessageBox.Show("Введите числовое значение!", Рик)
        End If

    End Sub

    Private Sub ДолжИнстр_Closing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        'If (e.CloseReason = CloseReason.UserClosing) Then
        '    f = True
        'End If
    End Sub
End Class