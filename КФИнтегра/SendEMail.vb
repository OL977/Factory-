Imports EASendMail
Public Class SendEMail
    Private Sub SendEMail_Load(sender As Object, e As EventArgs) Handles MyBase.Load




    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim oMail As New SmtpMail("TryIt")
        Dim oSmtp As New SmtpClient()
        oMail.From = TextBox1.Text
        oMail.To = TextBox2.Text
        oMail.Subject = TextBox3.Text
        oMail.TextBody = TextBox4.Text
        Dim oServer As New SmtpServer("")


        If TextBox1.Text = "6289925@mail.ru" Then
            'oServer.Server = "smtp.mail.ru"
            'oServer.User = "6289925@mail.ru"
            'oServer.Password = "6807057a"
            'oServer.Port = 465
            'oServer.ConnectType = SmtpConnectType.ConnectSSLAuto 'если требуется обязательно ssl
        ElseIf TextBox1.Text = "1389925@gmail.com" Then
            'oServer.Server = "smtp.gmail.com"
            'oServer.User = "1389925@gmail.com"
            'oServer.Password = "oleg110403"
            'oServer.Port = 465
            'oServer.ConnectType = SmtpConnectType.ConnectSSLAuto 'если требуется обязательно ssl
        ElseIf TextBox1.Text = "os@2trans.by" Then
            'oServer.Server = "smtp.yandex.ru"
            'oServer.User = "os@2trans.by"
            'oServer.Password = "oleg61351127441"
            'oServer.Port = 465
            'oServer.ConnectType = SmtpConnectType.ConnectSSLAuto 'если требуется обязательно ssl
        Else

        End If

        'Dim oServer As New SmtpServer("smtp.emailarchitect.net")
        'oServer.User = "test@emailarchitect.net"
        'oServer.Password = "testpassword"
        Try
            oSmtp.SendMail(oServer, oMail)
            MessageBox.Show("Письмо отправлено!", Рик)
        Catch ex As Exception
            MessageBox.Show("Ошибка передачи " & ex.ToString)
        End Try

    End Sub
End Class