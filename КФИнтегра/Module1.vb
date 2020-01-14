Imports System.Data.OleDb
Imports MySql.Data.MySqlClient
Imports System.Management
Imports System.IO
Imports System.Security.Policy
Imports System.Data.SqlClient
Imports NLog

Module Module1
    Public conn As OleDbConnection
    Public Отдел, ИмяКомп, ЛемелТрОтп, ЛемелИспытСрок, Дпод1, Дпод2 As String
    Public Рик As String = "Кондитерская фабрика Интегра"
    Public conn2 As MySqlConnection
    Public errds, gr1, gr2, errupd, proverka As Integer, ДогПодрНомНовы As Integer = 0
    Public Прим As Integer
    Public FilesList27 As New List(Of String)
    Public ConStringOleDb As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=U:\Офис\Рикконсалтинг\Рик.accdb; Persist Security Info=False;"
    Public ConString1 As String = "Data Source=45.14.50.13\723\SQLEXPRESS,1433;Network Library=DBMSSOCN;Initial Catalog=Интегра;User ID=userintegra;Password=61kzHBRa4e6Mfl"
    Public ConString As String = "Data Source=45.14.50.142\2749\SQLEXPRESS,1433;Network Library=DBMSSOCN;Initial Catalog=Интегра;User ID=userintegra1;Password=61kzHBRa4e6Mfl"

    Public FTPString As String = "ftp://86.57.135.184:21/"
    Public HstName As String = "86.57.135.184"
    Public FTPUser As String = "user1"
    Public FTPPass As String = "Jd3Kds9"
    Public FTPStringAllDOC As String = FTPString & "ALLINALLDATABASE/"
    'Public ConString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\Users\OLEG\Desktop\Рик.accdb;Persist Security Info=False;"
    Public ДогПодрВыпРаб As List(Of String)
    Public ДогПодрВыпРабСтР As List(Of String)
    Public ДогПодрВыпРабСтК As List(Of String)
    Public ДогПодрВыпРабСтОб As List(Of String)
    Public ДогПодНомерСтар, ПрЗакрВидыРаб As String
    Public СписокКлиентовОсновной As DataTable
    Public mast2 As New ArrayList()
    Public NameOrg As String
    Public КрестикНажатиеДогПодряда As Boolean = False
    Public Delegate Sub coxt(ByVal form As Form, ByVal strsql As String, ByVal c As ComboBox)
    Public Delegate Sub coxt1(ByVal form As Form, ByVal strsql As String, ByVal c As ListBox)
    Public Delegate Sub coxt2()
    Public Delegate Sub coxt3()
    Public Delegate Sub coxt4()
    Public Delegate Sub coxt5()
    Public UserAdmin As String
    Public OnePath As String
    Public connДоработчик As SqlConnection
    Public Logger = LogManager.GetCurrentClassLogger()
    Public ПодтверждПароляУдаление As Boolean = False
    Public dbcx As DbAllDataContext

    Public Sub COMxt(ByVal form As Form, ByVal strsql As String, ByVal c As ComboBox)
        'Dim strsql As String = "SELECT DISTINCT Страна FROM Страна ORDER BY Страна"
        'Dim sty As String = form.Name
        'Dim c1 As ComboBox = form.Controls("ComboBox" & c)

        Dim ds As DataTable = Selects(strsql)
        If c.InvokeRequired Then
            form.Invoke(New coxt(AddressOf COMxt), form, strsql, c)
        Else
            c.AutoCompleteCustomSource.Clear()
            c.Items.Clear()
            For Each r As DataRow In ds.Rows
                c.AutoCompleteCustomSource.Add(r.Item(0).ToString())
                c.Items.Add(r(0).ToString)
            Next
        End If
    End Sub
    Public Sub Listxt(ByVal form As Form, ByVal strsql As String, ByVal c As ListBox)
        'Dim strsql As String = "SELECT DISTINCT Страна FROM Страна ORDER BY Страна"
        'Dim sty As String = form.Name
        'Dim c1 As ComboBox = form.Controls("ComboBox" & c)

        Dim ds As DataTable = Selects(strsql)
        If c.InvokeRequired Then
            form.Invoke(New coxt1(AddressOf Listxt), form, strsql, c)
        Else
            c.Items.Clear()
            For Each r As DataRow In ds.Rows
                c.Items.Add(r(0).ToString)
            Next
        End If
    End Sub
    Public Sub Pext(ByVal strsql As String, ByVal c As DataTable)
        c = Selects(strsql)

        'If c.InvokeRequired Then
        '    Form.Invoke(New coxt1(AddressOf Pext), Form, strsql, c)
        'Else
        '    c.Items.Clear()
        '    For Each r As DataRow In ds.Rows
        '        c.Items.Add(r(0).ToString)
        '    Next
        'End If
    End Sub

    'Public ConString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=U:\Офис\Рикконсалтинг\Рик.accdb; Persist Security Info=False;"


    'Dim conn As New MySqlConnection("Server=os2trahw.beget.tech;Database=os2trahw_rikacce;Uid=os2trahw_rikacce;Pwd=oleg1389925")
    'Dim cmd As New MySqlCommand
    'Dim reader As MySqlDataReader
    'Dim adapter As MySqlDataAdapter
    'Dim dt As New DataTable
    'Try
    '    conn.Open()
    '        cmd.Connection = conn
    '    cmd.CommandText = "SELECT ФИОСборное FROM Сотрудники WHERE НазвОрганиз='ППТорг'"
    '    'cmd.CommandText = "INSERT INTO `test_table` (`id`, `test_info`) VALUES (NULL, 'some text info for current id');"
    '    Try
    '            cmd.ExecuteNonQuery()
    '        Catch ex As Exception
    ''описание того, что программа должна делать в случае возникновения каких-либо непредвиденных обстоятельств
    'End Try
    '    'для получения данных из таблиц (запросы типа select) используется reader.


    '    'reader = cmd.ExecuteReader()
    '    'While reader.Read()
    '    '    'получаем и сообщаем пользователю значения первого столбца базы данных для всех выбранных запросом строк
    '    '    'MsgBox(reader.GetValue(0))
    '    'End While
    '    'Catch ex As Exception
    '    '    MessageBox.Show(ex.Message)
    '    '    'описание действий при проблемах с подключением к БД
    '    'End Try
    '    adapter = New MySqlDataAdapter(cmd)
    '    adapter.Fill(dt)

    '    For Each r As DataRow In dt.Rows
    'Me.ComboBox1.Items.Add(r(0).ToString)
    'Next

    Public Sub Соед(ByVal intd As Integer)

        Try
            If conn.State = ConnectionState.Open Then
                'conn = New OleDbConnection
                'conn.ConnectionString = ConString
                'Try
                '    conn.Open()
                'Catch ex As Exception
                '    MessageBox.Show("Не подключен диск U")
                'End Try
                conn.Close()
                If intd = 1 Then
                    conn.Dispose()
                End If
                Exit Sub
            End If
            If conn.State = ConnectionState.Closed Then
                conn = New OleDbConnection
                conn.ConnectionString = ConString
                Try
                    conn.Open()
                Catch ex As Exception
                    MessageBox.Show("Не подключен диск U")

                End Try
            End If

        Catch ex As Exception
            conn = New OleDbConnection
            conn.ConnectionString = ConString

            Try
                conn.Open()
            Catch ex1 As Exception
                'ConnVPN()
                Process.Start("U:")
            End Try





        End Try

    End Sub
    Public Sub ОткрытиеФайлаБезПути(ByVal d As String)
        Dim currentdirectory = OnePath
        Dim filename = d
        Dim path = System.IO.Directory.GetFiles(currentdirectory, "*" & filename, IO.SearchOption.AllDirectories)(0)
        Process.Start(path)
    End Sub

    Public Sub Updates(ByVal stroka As String)
        Dim conn4 As New SqlConnection(ConString)
        If conn4.State = ConnectionState.Closed Then
            conn4.Open()
        End If
        Dim c As New SqlCommand(stroka, conn4)
        Try
            c.ExecuteNonQuery()
            If conn4.State = ConnectionState.Open Then
                conn4.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            If conn4.State = ConnectionState.Open Then
                conn4.Close()
            End If
        End Try
    End Sub
    Public Sub Updates(ByVal stroka As String, ByVal Parram As Dictionary(Of String, Object))
        Dim conn4 As New SqlConnection(ConString)
        If conn4.State = ConnectionState.Closed Then
            conn4.Open()
        End If
        Dim c As New SqlCommand(stroka, conn4)

        c.Parameters.Clear()
        Dim list = New List(Of SqlParameter)()

        For Each r In Parram
            list.Add(New SqlParameter(r.Key, r.Value))
        Next

        c.Parameters.AddRange(list.ToArray())
        Try
            c.ExecuteNonQuery()
            If conn4.State = ConnectionState.Open Then
                conn4.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            If conn4.State = ConnectionState.Open Then
                conn4.Close()
            End If
        End Try

    End Sub
    Public Sub Updates(ByVal stroka As String, ByVal Parram As Dictionary(Of String, Object), ByVal d As String)
        Dim conn4 As New SqlConnection(ConString)
        If conn4.State = ConnectionState.Closed Then
            conn4.Open()
        End If
        Dim c As New SqlCommand(stroka, conn4)

        c.Parameters.Clear()
        Dim list = New List(Of SqlParameter)()

        For Each r In Parram
            list.Add(New SqlParameter(r.Key, r.Value))
        Next

        c.Parameters.AddRange(list.ToArray())
        Try
            c.ExecuteNonQuery()
            If conn4.State = ConnectionState.Open Then
                conn4.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            If conn4.State = ConnectionState.Open Then
                conn4.Close()
            End If
            Exit Sub
        End Try
        АвтоОбновлТаблиц(d)
    End Sub

    Public Function Updates(ByVal stroka As String, ByVal Parram As Dictionary(Of String, Object), ByVal d As String, ByVal f As Integer) As Integer 'возвращает id новой записи
        Dim id As Integer
        Dim conn4 As New SqlConnection(ConString)
        If conn4.State = ConnectionState.Closed Then
            conn4.Open()
        End If
        Dim c As New SqlCommand(stroka, conn4)

        c.Parameters.Clear()
        Dim list = New List(Of SqlParameter)()

        For Each r In Parram
            list.Add(New SqlParameter(r.Key, r.Value))
        Next

        c.Parameters.AddRange(list.ToArray())
        Try
            id = c.ExecuteScalar()
            If conn4.State = ConnectionState.Open Then
                conn4.Close()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            If conn4.State = ConnectionState.Open Then
                conn4.Close()
            End If
            Return 0
            Exit Function
        End Try

        АвтоОбновлТаблиц(d)
        Return id
    End Function

    Private Sub АвтоОбновлТаблиц(ByVal d As String)

        Select Case d
            Case "ФормаСобств"
                dtformftAsync()
            Case "Клиент"
                dtClient()
            Case "Сотрудники"
                dtSotrudniki()
            Case "Штатное"
                dtShtatnoe()
            Case "ПутиДокументов"
                dtPutiDokumentov()
            Case "ШтОтделы"
                dtShtatnoeOtdely()
            Case "КарточкаСотрудника"
                dtKartochkaSotrudnika()
            Case "ПродлКонтракта"
                dtProdlenieKontrakta()
            Case "ДогСотруд"
                dtDogovorSotrudnik()
            Case "ДогПодряда"
                dtDogovorPadriada()
            Case "ШтСводИзмСтавка"
                dtShtatnoeSvodnoeIzmenStavka()
            Case "ШтСвод"
                dtShtatnoeSvodnoe()
            Case "СоставСемьи"
                dtSostavSemyi()
            Case "Статистика"
                dtStatistika()
            Case "Перевод"
                dtPerevod()
            Case "ОбъектОбщепита"
                dtObjectObshepita()
            Case "Отпуск"
                dtOtpusk()
            Case "ОтпускСотрудники"
                dtOtpuskSotrudniki()
            Case "ОтпускСоц"
                dtOtpuskSoc()
            Case "ДогПодрядаАкт"
                dtDogPodryadaAkt()
            Case "Окончание"
                dtOkonchanie()
            Case "ДогПодОсобен"
                dtPodriadaOsoben()
            Case "ШтатРаспис"
                dtStatnoeRaspisanie()
            Case "ДогПодДолжн"
                dtDogPodrDoljnost()
            Case "ДогПодрядаАктИное"
                dtDogPodrAktInoe()
            Case "ДогПодрядаРаботыИное"
                dtDogPodrRabotyInoe()
        End Select




    End Sub





    'старый вариант подключения 
    'Public Sub Updates(ByVal stroka As String)
    '    Dim c As New OleDbCommand
    '    c.Connection = conn
    '    c.CommandText = stroka
    '    Try
    '        c.ExecuteNonQuery()
    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message)
    '    End Try
    'End Sub

    ''Public Sub UpdatesPar(ByVal stroka As String, ByVal t As String)
    ''    Dim c As New OleDbCommand
    ''    c.Connection = conn
    ''    c.CommandText = stroka
    ''    c.Parameters.Add(New OleDbParameter(t, ""))
    ''    Try
    ''        c.ExecuteNonQuery()
    ''    Catch ex As Exception
    ''        MessageBox.Show(ex.Message)
    ''    End Try
    ''End Sub

    'старый вариант подключения 
    Public Function SelectsOleDb(ByVal StrSql As String) As DataTable
        errds = 0
        Dim conn As New OleDbConnection(ConStringOleDb)
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If
        Dim c As New OleDbCommand(StrSql, conn)

        Dim dst As New DataTable
        Dim da As New OleDbDataAdapter(c)
        Try
            da.Fill(dst)
            Dim gf As Object = dst.Rows(0).Item(0)
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If



            Return dst

        Catch ex As Exception
            errds = 1
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            Return dst
        End Try

    End Function


    Public Function Selects(ByVal StrSql As String) As DataTable
        errds = 0
        Dim conn3 As New SqlConnection(ConString)
        If conn3.State = ConnectionState.Closed Then
            conn3.Open()
        End If

        Dim c As New SqlCommand(StrSql, conn3)



        Dim dst As New DataTable

        Dim da As New SqlDataAdapter(c)
        Try
            da.Fill(dst)
            Dim gf As Object = dst.Rows(0).Item(0)

            If conn3.State = ConnectionState.Open Then
                conn3.Close()
            End If
            Return dst
        Catch ex As Exception
            errds = 1
            If conn3.State = ConnectionState.Open Then
                conn3.Close()
            End If
            Return dst
        End Try

    End Function
    Public Function Selects(ByVal StrSql As String, ByVal Parram As List(Of Date)) As DataTable
        errds = 0
        Dim conn3 As New SqlConnection(ConString)
        If conn3.State = ConnectionState.Closed Then
            conn3.Open()
        End If



        Dim c As New SqlCommand(StrSql, conn3)
        c.Parameters.Clear()
        Dim list = New List(Of SqlParameter)()
        list.Add(New SqlParameter("@времянач", Parram.Item(0)))
        list.Add(New SqlParameter("@времякон", Parram.Item(1)))

        c.Parameters.AddRange(list.ToArray())

        'c.Parameters.AddRange(New SqlParameter() {New SqlParameter("@времянач", SqlDbType.DateTime) With {
        '  Key.Value = Parram(0)})
        'For Each r In Parram
        '    c.Parameters.Add(r.Value(Of Date))
        'Next

        '        cm.Parameters.AddRange(New SqlParameter() {New SqlParameter("@TaskID", SqlDbType.Int) With {
        '    Key.Value = DirectCast(dr("BSTaskID"), Integer)
        '}, New SqlParameter("@DateNow", SqlDbType.DateTime) With {
        '    Key.Value = DateTime.Now
        '}, New SqlParameter("@ExecDate", SqlDbType.DateTime) With {
        '    Key.Value = execDate
        '}, New SqlParameter("@UserId", SqlDbType.Int) With {
        '    Key.Value = DirectCast(dr("BSUserId"), Integer)
        '}, New SqlParameter("@TaskParameters", SqlDbType.VarChar) With {
        '    Key.Value = parametersXML
        '}})

        Dim dst As New DataTable

        Dim da As New SqlDataAdapter(c)
        Try
            da.Fill(dst)
            Dim gf As Object = dst.Rows(0).Item(0)

            If conn3.State = ConnectionState.Open Then
                conn3.Close()
            End If
            Return dst
        Catch ex As Exception
            errds = 1
            If conn3.State = ConnectionState.Open Then
                conn3.Close()
            End If
            Return dst
        End Try

    End Function

    Public Function Selects(ByVal StrSql As String, ByVal Parram As Dictionary(Of String, String)) As DataTable  'Parram As List(Of String)
        errds = 0
        Dim conn3 As New SqlConnection(ConString)
        If conn3.State = ConnectionState.Closed Then
            conn3.Open()
        End If
        Dim c As New SqlCommand(StrSql, conn3)
        c.Parameters.Clear()
        Dim list = New List(Of SqlParameter)()

        For Each r In Parram
            list.Add(New SqlParameter(r.Key, r.Value))
        Next

        c.Parameters.AddRange(list.ToArray())


        Dim dst As New DataTable

        Dim da As New SqlDataAdapter(c)
        Try
            da.Fill(dst)
            Dim gf As Object = dst.Rows(0).Item(0)

            If conn3.State = ConnectionState.Open Then
                conn3.Close()
            End If
            Return dst
        Catch ex As Exception
            errds = 1
            If conn3.State = ConnectionState.Open Then
                conn3.Close()
            End If
            Return dst
        End Try

    End Function

    Public Function Selects(ByVal StrSql As String, ByVal Parram As Dictionary(Of String, Object)) As DataTable  'Parram As List(Of String)
        errds = 0
        Dim conn3 As New SqlConnection(ConString)
        If conn3.State = ConnectionState.Closed Then
            conn3.Open()
        End If
        Dim c As New SqlCommand(StrSql, conn3)
        c.Parameters.Clear()
        Dim list = New List(Of SqlParameter)()

        For Each r In Parram
            list.Add(New SqlParameter(r.Key, r.Value))
        Next

        c.Parameters.AddRange(list.ToArray())


        Dim dst As New DataTable

        Dim da As New SqlDataAdapter(c)
        Try
            da.Fill(dst)
            Dim gf As Object = dst.Rows(0).Item(0)

            If conn3.State = ConnectionState.Open Then
                conn3.Close()
            End If
            Return dst
        Catch ex As Exception
            errds = 1
            If conn3.State = ConnectionState.Open Then
                conn3.Close()
            End If
            Return dst
        End Try

    End Function

    Public Sub Контракты(ByVal год As Integer, ByVal Папка As String, ByVal Организация As String)

        Dim l = listFluentFTP("/" & Организация & "/" & Папка & "/" & год & "/")




        'Dim Files3(), gth3 As String

        'Try

        '    Files3 = (IO.Directory.GetFiles(OnePath & Прием.ComboBox1.Text & "\Контракт\" & год, "*.doc", IO.SearchOption.TopDirectoryOnly))
        '    For n As Integer = 0 To Files3.Length - 1
        '        gth3 = ""
        '        gth3 = IO.Path.GetFileName(Files3(n))
        '        Files3(n) = gth3
        '    Next
        '    Прием.ComboBox2.Items.Clear()
        '    Прием.ComboBox2.Items.AddRange(Files3)
        'Catch ex As Exception
        '    MessageBox.Show("В " & год & " году нет контрактов!", Рик)
        'End Try




    End Sub
    Public Sub Справки(ByVal год As Integer)


        'Dim Files3(), gth3 As String

        'Try

        '    Files3 = (IO.Directory.GetFiles(OnePath & СправкаПоЗарплате.ComboBox1.Text & "\Справки\" & год, "*.doc", IO.SearchOption.TopDirectoryOnly))
        '    For n As Integer = 0 To Files3.Length - 1
        '        gth3 = ""
        '        gth3 = IO.Path.GetFileName(Files3(n))
        '        Files3(n) = gth3
        '    Next
        '    СправкаПоЗарплате.ComboBox8.Items.Clear()
        '    СправкаПоЗарплате.ComboBox8.Items.AddRange(Files3)
        '    СправкаПоЗарплате.ComboBox8.Text = Files3.Last
        'Catch ex As Exception

        'End Try




    End Sub




    Public Sub Приказы(ByVal год As Integer)
        Dim Files2(), gth3 As String

        Try

            Files2 = (IO.Directory.GetFiles(OnePath & Прием.ComboBox1.Text & "\Приказ\" & год, "*.doc", IO.SearchOption.TopDirectoryOnly))
            For n As Integer = 0 To Files2.Length - 1
                gth3 = ""
                gth3 = IO.Path.GetFileName(Files2(n))
                Files2(n) = gth3
            Next
            Прием.ComboBox17.Items.Clear()
            Прием.ComboBox17.Items.AddRange(Files2)
        Catch ex As Exception
            MessageBox.Show("В " & год & " году нет приказов!", Рик)
        End Try


        'Dim strValues As String() = New String() {"1", "2", "3", "4"} 'из массива в лист оф очень класная штука
        'Dim strList As List(Of String) = strValues.ToList()
        'strList.Remove("4")
        'strValues = strList.ToArray()

    End Sub

    Public Sub KillProc()
        Exit Sub

        Try
            If IO.Directory.Exists("c:\Users\Public\Documents\Рик") Then
                IO.Directory.Delete("c:\Users\Public\Documents\Рик", True)
                IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рик")
            Else
                IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рик")
            End If
        Catch ex As Exception

            For Each p As Process In Process.GetProcessesByName("winword")
                p.Kill()
                p.WaitForExit()
            Next

            Try
                If IO.Directory.Exists("c:\Users\Public\Documents\Рик") Then
                    IO.Directory.Delete("c:\Users\Public\Documents\Рик", True)

                    IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рик")
                Else
                    IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рик")
                End If
            Catch ex1 As Exception
                MessageBox.Show("Закройте файл Excel или Word" & vbCrLf & "И нажмите 'OK'!", Рик)
                If IO.Directory.Exists("c:\Users\Public\Documents\Рик") Then
                    IO.Directory.Delete("c:\Users\Public\Documents\Рик", True)

                    IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рик")
                Else
                    IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рик")
                End If
            End Try





        End Try
    End Sub
    Public Sub KillExcel()
        For Each p As Process In Process.GetProcessesByName("Excel")
            p.Kill()
            p.WaitForExit()
        Next
    End Sub
    Public Async Sub Статистика1(ByVal ФИО As String, ByVal Событие As String, ByVal Организ As String)
        Await Task.Run(Sub() Статистика2(ФИО, Событие, Организ))
    End Sub
    Private Sub Статистика2(ByVal ФИО As String, ByVal Событие As String, ByVal Организ As String)
        Dim f As String = Now.ToShortTimeString
        f = Strings.Right(f, 5)
        Dim Strsql As String = "INSERT INTO Статистика(Дата,Время,Организация,Событие,Сотрудник,КемИзменено) VALUES('" & Now.Date & "','" & f & "', '" & Организ & "',
'" & Событие & "','" & ФИО & "','" & ИмяКомп & "')"
        Updates(Strsql)
    End Sub

    Public Sub connectTo(ByVal name As String) 'подключение Wi Fi
        Dim p = "netsh.exe"
        Dim sInfo As New ProcessStartInfo(p, "wlan connect " & name)
        sInfo.CreateNoWindow = True
        sInfo.WindowStyle = ProcessWindowStyle.Hidden
        Process.Start(sInfo)
    End Sub

    Public Sub ConnVPN()
        Dim p As New Process

        p.StartInfo.UseShellExecute = False
        p.StartInfo.RedirectStandardOutput = True
        p.StartInfo.RedirectStandardError = True
        p.StartInfo.FileName = "rasdial.exe"
        p.StartInfo.Arguments = """VPN-РикКонсалтинг"""
        p.Start()

        'If Not p.WaitForExit(My.Settings.VpnTimeout) Then
        '    Throw New Exception(
        'String.Format("Connecting to ""{0}"" VPN failed after {1}ms", sVpn, My.Settings.VpnTimeout))
        'End If

        'If p.ExitCode <> 0 Then
        '    Throw New Exception(
        'String.Format("Failed connecting to ""{0}"" with exit code {1}. Errors {2}", sVpn, p.ExitCode, p.StandardOutput.ReadToEnd.Replace(vbCrLf, "")))
        'End If
    End Sub

    Public Function InputName1(ByVal ФИО As String, ByVal DOC As String) As String
        Dim inp As String
        Dim strsql As String = "SELECT " & DOC & " FROM InputName WHERE ФИООриганл='" & ФИО & "'"
        Dim ds As DataTable = Selects(strsql)

        If errds = 0 Then
            Return ds.Rows(0).Item(0).ToString
        Else

            inp = InputBox("Введите ФИО сотрудника " & vbCrLf & ФИО & vbCrLf & " в Дательном падеже 'Предоставить отпуск Кому?'", Рик)

            Do Until inp <> ""
                MessageBox.Show("Повторите ввод данных!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Error)
                inp = InputBox("Введите ФИО сотрудника " & vbCrLf & ФИО & vbCrLf & " в Дательном падеже 'Предоставить отпуск Кому?'", Рик)
            Loop

            Dim strsql2 As String = "INSERT INTO InputName(ФИООриганл, " & DOC & ") VALUES('" & ФИО & "','" & inp & "')"
            Updates(strsql2)

            Return inp
        End If

    End Function

    Public Function ДобРазрядности(ByVal d As String) As String
        Dim f As Double
        f = Replace(d, ".", ",")
        If СправкаПоЗарплате.bool(f) = True Then
            d = f & ",00"
            Return d
        Else
            d = f
            If СправкаПоЗарплате.Count(f) = 1 Then
                d = f & "0"
                Return d
            End If
        End If
        Return d
    End Function

    'Sub GetFindSub(ByVal d As String, ByVal f As String) 'поиск файлов на компьютере
    '    Try
    '        Application.DoEvents()
    '        Dim ListFiles As String() = IO.Directory.GetFiles(d, f)
    '        Dim ListFolders As String() = IO.Directory.GetDirectories(d)
    '        For Each item In ListFiles
    '            ListBox1.Items.Add(item)
    '        Next
    '        For Each item In ListFolders
    '            GetFindSub(item, f)
    '        Next
    '    Catch ex As Exception
    '    End Try
    'End Sub


    Public Sub GridView(ByVal d As DataGridView)
        d.EnableHeadersVisualStyles = False
        d.ColumnHeadersDefaultCellStyle.BackColor = Color.LightGreen
        d.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Raised
        d.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        d.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        d.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Public Sub GridViewRed(ByVal d As DataGridView)
        d.EnableHeadersVisualStyles = False
        d.ColumnHeadersDefaultCellStyle.BackColor = Color.Yellow
        d.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Raised
        d.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        d.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        d.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub


    Public Function Доработчик(ByVal strsql5 As String) As SqlDataAdapter
        connДоработчик = New SqlConnection(ConString)
        If connДоработчик.State = ConnectionState.Closed Then
            connДоработчик.Open()
        End If
        Dim c5 As New SqlCommand(strsql5, connДоработчик)
        Dim da5 As New SqlDataAdapter(c5)
        Return da5
    End Function


    'Public Sub ЗагрузканасерверFTP()
    '    Dim request As Net.FtpWebRequest = CType(Net.WebRequest.Create("ftp://86.57.135.184:21/" & st), Net.FtpWebRequest)
    '    request.Credentials = New Net.NetworkCredential("user1", "Jd3Kds9")
    '    request.Method = Net.WebRequestMethods.Ftp.MakeDirectory
    'End Sub


    ' Dim sw As New Stopwatch 'вычисление выполнения метода
    'sw.Start()
    'sw.Stop()
    '  MessageBox.Show((sw.ElapsedMilliseconds / 100.0).ToString())
    Public Sub GridView2(ByVal d As DataGridView)
        d.EnableHeadersVisualStyles = False
        d.ColumnHeadersDefaultCellStyle.BackColor = Color.LightGreen
        d.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Raised
        d.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        d.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        d.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
End Module
