Imports WinSCP
Imports System
Imports System.IO
Imports System.IO.Compression
Imports System.Text
'Imports System.Net
Imports FluentFTP

Imports System.Net
Imports System.Security.AccessControl

Module Ftp

    Public Function ВыборкаФайловПоОпределенномуПути(ByVal Путь As String)

        Dim request As FtpWebRequest = CType(WebRequest.Create(FTPString & Путь), FtpWebRequest)
        Dim List = New List(Of String)()
        request.Credentials = New NetworkCredential(FTPUser, FTPPass)
        request.Method = WebRequestMethods.Ftp.ListDirectory

        Dim response = CType(request.GetResponse(), FtpWebResponse)
        Dim reader As StreamReader = New StreamReader(response.GetResponseStream())
        Dim dir As New List(Of String)()
        Dim line2 As String = reader.ReadLine()
        While (Not String.IsNullOrEmpty(line2))

            dir.Add(line2)
            line2 = reader.ReadLine()
        End While
        response.Close()
        response.Dispose()

        Return dir

    End Function

    'Public Sub NetFtpClient()

    '    Dim SessionOptions As SessionOptions = New SessionOptions
    '    With SessionOptions
    '        .Protocol = Protocol.Ftp
    '        .HostName = HstName ' "ftp.example.com"ftp://86.57.135.184:21
    '        .UserName = FTPUser
    '        .Password = FTPPass
    '    End With

    '    Using session As Session = New Session()
    '        ' Connect
    '        session.Open(SessionOptions)
    '        'session.GetFiles("/Dokument2/*", firthtPath & "\*").Check()

    '        'Dim f = session.ListDirectory("/Кадры/*")
    '        Dim f = session.ListDirectory("/Кадры")
    '        session.Close()
    '    End Using



    '    '    Dim ftp As System.Net.FtpClient.FtpClient = New System.Net.FtpClient.FtpClient()
    '    'ftp.Host = HstName
    '    'ftp.Credentials = New NetworkCredential(FTPUser, FTPPass)
    '    'ftp.Connect()
    '    'Dim items() = ftp.GetNameListing("*/2019/*")
    '    'Dim items2 As List(Of String)


    '    ''For x As Integer = 0 To items.Length
    '    ''    items2.Add(ftp.GetNameListing("Кадры/*")

    '    ''Next




    '    'ftp.Disconnect()


    'End Sub




    Public Sub FluedЗагрузкаФайлов(ByVal ПутьИсходногоФайла As String, ByVal ПутьНазначенияИИмяФайла As String)
        Dim ftp As FluentFTP.FtpClient = New FluentFTP.FtpClient(HstName, FTPUser, FTPPass)
        'ftp.DataConnectionType = FtpDataConnectionType.AutoPassive
        ftp.Connect()

        Dim str As String = Replace(ПутьИсходногоФайла, FTPString, "/")
        'Dim items = ftp.GetNameListing("*")        '("Кадры")
        ftp.DownloadFile(ПутьНазначенияИИмяФайла, str, True)

        'Dim b As String = items(0).ToString
        'Dim g = ftp.GetObjectInfo("Кадры")

        ''Dim items2 As FtpListItem = ftp.DownloadFiles("/Кадры/Штатное расписание/2019/")

        'Dim f1 = ftp.GetWorkingDirectory()
        'ftp.SetWorkingDirectory("Кадры/Должностные инструкции")
        'Dim f5 = ftp.GetNameListing("*")
        'Dim f2 = ftp.GetNameListing()
        ftp.Disconnect()


        'Dim f10 As Integer = items.Length
        'FtpClient ftp = New FtpClient(txtUsername.Text, txtPassword.Text, txtFTPAddress.Text);
        'FtpListItem [] items = ftp.GetListing (); // здесь вы можете получить список с типом, именем, датой изменения и другими свойствами.
        'FtpFile File = New FtpFile(ftp, "8051812.xml");
        '// файл для получения file.Download ("c: \\ 8051812.xml");
        '// загрузка файла.Name = "8051814.xml"; 
        '// измените имя, чтобы получить новый файл. Загрузка ("c: \\ 8051814.xml"); ftp.Disconnect (); // закрываем
    End Sub



    Public Sub ЗагрузкаФайловНаСервер(ByVal ПутьИсходногоФайла As String, ByVal ПутьНазначенияИИмяФайла As String)

        If Not ПутьИсходногоФайла.Contains("ftp:") Then
            Dim f As String = ПутьНазначенияИИмяФайла
            Dim f2 As String = ПутьИсходногоФайла
            ПутьИсходногоФайла = f
            ПутьНазначенияИИмяФайла = f2
        End If


        Using ftp As FluentFTP.FtpClient = New FluentFTP.FtpClient(HstName, FTPUser, FTPPass)
            'ftp.DataConnectionType = FtpDataConnectionType.AutoActive
            ftp.Connect()
            Dim str As String = Replace(ПутьИсходногоФайла, FTPString, "/")
            Try
                ftp.UploadFile(ПутьНазначенияИИмяФайла, str, True)
                ftp.Disconnect()
            Catch ex As Exception
                If ex.Message.Contains("так как этот файл используется другим процессом.") = True Then
                    MessageBox.Show("Файл уже открыт в другой программе!", Рик)
                End If
                ftp.Disconnect()
            End Try

        End Using



    End Sub
    Public Async Sub _DeleteFluentFTP(ByVal ПутьУдаляемогоФайла As String)
        Await Task.Run(Sub() DeleteFluentFTP(ПутьУдаляемогоФайла))
    End Sub

    Public Sub DeleteFluentFTP(ByVal ПутьУдаляемогоФайла As String)
        Using ftp As FluentFTP.FtpClient = New FluentFTP.FtpClient(HstName, FTPUser, FTPPass)
            'ftp.DataConnectionType = FtpDataConnectionType.AutoPassive

            ftp.Connect()
            Dim str As String = Replace(ПутьУдаляемогоФайла, FTPString, "/")
            Try

                ftp.DeleteFile(str)
                ftp.Disconnect()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                ftp.Disconnect()


            End Try
        End Using

    End Sub
    Public Function ВыгрузкаФайловНаЛокалыныйКомп(ByVal ПутьИсходногоФайла As String, ByVal ПутьНазначенияИИмяФайла As String)



        Using ftp As FluentFTP.FtpClient = New FluentFTP.FtpClient(HstName, FTPUser, FTPPass)
            'ftp.DataConnectionType = FtpDataConnectionType.AutoPassive

            ftp.Connect()

            Dim str As String = Replace(ПутьИсходногоФайла, FTPString, "/")
            Try
                ftp.DownloadFile(ПутьНазначенияИИмяФайла, str, True)
                ftp.Disconnect()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                ftp.Disconnect()

                Return 1
            End Try
        End Using
        Return 0
    End Function

    'Public Sub ВыгрузкаФайлаССервера(ByVal ПутьИсходногоФайла As String, ByVal ПутьНазначенияИИмяФайла As String)
    '    Try
    '        My.Computer.Network.DownloadFile(ПутьИсходногоФайла, ПутьНазначенияИИмяФайла, FTPUser, FTPPass, True, 500, True) 'выгрузка с сервера
    '    Catch ex As IOException
    '        IO.File.Delete(ПутьНазначенияИИмяФайла)
    '        My.Computer.Network.DownloadFile(ПутьИсходногоФайла, ПутьНазначенияИИмяФайла, FTPUser, FTPPass) 'выгрузка с сервера
    '    End Try

    'End Sub

    Public Sub ЗагрВсехФайловсСервера()
        Dim SessionOptions As SessionOptions = New SessionOptions
        With SessionOptions
            .Protocol = Protocol.Ftp
            .HostName = "os2trahw.beget.tech" '"ftp.example.com"
            .UserName = "os2trahw_rikcon"
            .Password = "oleg110403"
        End With

        Using session As Session = New Session()
            ' Connect
            session.Open(SessionOptions)

            ' Download files
            session.GetFiles("/Dokument2/*", firthtPath & "\*").Check()

            'session.GetFiles("/directory/to/download/*", "C:\target\directory\*").Check()
            session.Close()
        End Using
        zipasync()

    End Sub
    Public Sub ЗагрФайлаИзСервера(ByVal ПутьОтСервера As String, ByVal ПутьСохранФайлаЛокально As String)
        Dim SessionOptions As SessionOptions = New SessionOptions
        With SessionOptions
            .Protocol = Protocol.Ftp
            .HostName = HstName ' "ftp.example.com"ftp://86.57.135.184:21
            .UserName = FTPUser
            .Password = FTPPass
        End With

        Using session As Session = New Session()
            ' Connect
            session.Open(SessionOptions)

            ' Download files
            Try
                session.GetFiles(ПутьОтСервера, ПутьСохранФайлаЛокально).Check()
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

            'session.GetFiles("/directory/to/download/*", "C:\target\directory\*").Check()
            session.Close()
        End Using
    End Sub
    Public Async Sub zipasync()
        Await Task.Run((Sub() zipa()))
    End Sub

    Public Sub zipa()
        Dim zipPath As String = firthtPath & "\General(English).zip"
        Dim extractPath As String = firthtPath & "/"

        Using archive As ZipArchive = ZipFile.OpenRead(zipPath)
            For Each entry As ZipArchiveEntry In archive.Entries
                Dim entryFullname = Path.Combine(extractPath, entry.FullName)
                Dim entryPath = Path.GetDirectoryName(entryFullname)
                If (Not (Directory.Exists(entryPath))) Then
                    Directory.CreateDirectory(entryPath)
                End If

                Dim entryFn = Path.GetFileName(entryFullname)
                If (Not String.IsNullOrEmpty(entryFn)) Then
                    entry.ExtractToFile(entryFullname, True)
                End If
            Next
        End Using
    End Sub

    Public Function СозданиепапкиНаСервере(ByVal st As String) As String
        Dim request As FtpWebRequest = CType(WebRequest.Create(FTPString & st), FtpWebRequest)
        request.Credentials = New NetworkCredential(FTPUser, FTPPass)
        request.Method = WebRequestMethods.Ftp.MakeDirectory
        Dim responce As FtpWebResponse = Nothing
        Try
            responce = CType(request.GetResponse(), FtpWebResponse)
            responce.Close()
        Catch ex As Exception

        End Try
        Return FTPString & st
    End Function
    Public Sub ВременнаяПапкаСоздание()

        If IO.Directory.Exists(firthtPath & "\Времянка") Then
            Try

                IO.Directory.Delete(firthtPath & "\", True)
                IO.Directory.CreateDirectory(firthtPath & "\Времянка")
            Catch ex As Exception
                IO.Directory.CreateDirectory(firthtPath & "\Времянка")
            End Try

        Else
            IO.Directory.CreateDirectory(firthtPath & "\Времянка")
        End If
    End Sub
    Public Sub ВременнаяПапкаУдалениеФайла(ByVal f As String)
        Try
            IO.File.Delete(f)
        Catch ex As Exception

        End Try

    End Sub
    Public Sub ЗагрузкаФайловНаСервер2(ByVal ПутьИсходногоФайла As String, ByVal ПутьНазначенияИИмяФайла As String)
        Dim wc As New WebClient
        wc.Credentials = New NetworkCredential(FTPUser, FTPPass)
        Try
            wc.UploadFile(New Uri(FTPString & ПутьНазначенияИИмяФайла), ПутьИсходногоФайла)
            wc.Dispose()
        Catch ex As Exception
            wc.Dispose()
            MessageBox.Show(ex.Message)

        End Try


    End Sub



    Public Async Sub ЗагрНаСерверИУдаление(ByVal ПутьИсходногоФайла As String, ByVal ПутьНазначенияИИмяФайла As String, ByVal ПутьИсходногоФайла2 As String)
        Await Task.Run(Sub() ЗагрузкаФайловНаСервер(ПутьИсходногоФайла, ПутьНазначенияИИмяФайла))
        'Await Task.Run(Sub() ЗагрузкаФайловНаСервер2(ПутьИсходногоФайла, ПутьНазначенияИИмяФайла))

        ВременнаяПапкаУдалениеФайла(ПутьИсходногоФайла2)
    End Sub
    Public Sub УдалениеФайлаНаСервере(ByVal f As String)
        Dim FTPRequest As FtpWebRequest = DirectCast(WebRequest.Create(f), FtpWebRequest)
        FTPRequest.Credentials = New NetworkCredential(FTPUser, FTPPass)
        FTPRequest.Method = WebRequestMethods.Ftp.DeleteFile
        Dim responce As FtpWebResponse = CType(FTPRequest.GetResponse(), FtpWebResponse)
        responce.Close()
        '    Dim ftpResp As FtpWebResponse = FTPRequest.GetResponse
    End Sub

    Public Function listFTP() As List(Of String)
        Dim requ As FtpWebRequest = Nothing
        Dim resp As FtpWebResponse = Nothing
        Dim reader As StreamReader = Nothing

        Dim list As New List(Of String)()

        Try

            requ = CType(WebRequest.Create(FTPString), FtpWebRequest)
            requ.Credentials = New NetworkCredential(FTPUser, FTPPass)
            requ.Method = WebRequestMethods.Ftp.ListDirectory
            resp = CType(requ.GetResponse(), FtpWebResponse)
            reader = New StreamReader(resp.GetResponseStream())


            While (reader.Peek() > -1)
                list.Add(reader.ReadLine())
            End While
            resp.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Return list

    End Function
    Public Function listFluentFTP(ByVal Директория As String) As List(Of String)
        Dim list As New List(Of String)()
        Dim ftp As FluentFTP.FtpClient = New FluentFTP.FtpClient(HstName, FTPUser, FTPPass)
        'ftp.DataConnectionType = FtpDataConnectionType.PASV
        ftp.Connect()
        Dim fg As String = Replace(Директория, FTPString, "/")

        For Each item As FtpListItem In ftp.GetListing(fg)

            If item.Type = FtpFileSystemObjectType.File Then
                list.Add(item.Name) 'client.UploadFile(@"C:\MyVideo.mp4", "/htdocs/MyVideo.mp4");
            ElseIf item.Type = FtpFileSystemObjectType.Directory Then
                list.Add(item.Name)
            End If

        Next




        ftp.Disconnect()
        Return list


        '// if this Is a file
        'If (item.Type == FtpFileSystemObjectType.File) Then{

        '	// get the file size
        '	Long Size = client.GetFileSize(item.FullName);

        '}

        '// get modified date/time of the file Or folder
        'DateTime time = client.GetModifiedTime(item.FullName);

        '// calculate a hash for the file on the server side (default algorithm)
        'FtpHash hash = client.GetHash(item.FullName);

    End Function
    Public Sub listFluentFTP2(ByVal Директория As String)
        Dim list As New List(Of String)()
        Dim ftp As FluentFTP.FtpClient = New FluentFTP.FtpClient(HstName, FTPUser, FTPPass)
        'ftp.DataConnectionType = FtpDataConnectionType.PASV
        ftp.Connect()
        Dim fg As String = Replace(Директория, FTPString, "/")

        'For Each item As FtpListItem In ftp.GetListing(fg)

        '    If item.Type = FtpFileSystemObjectType.File Then
        '        list.Add(item.Name) 'client.UploadFile(@"C:\MyVideo.mp4", "/htdocs/MyVideo.mp4");
        '    ElseIf item.Type = FtpFileSystemObjectType.Directory Then
        '        list.Add(item.Name)
        '    End If

        'Next

        Dim files As FtpListItem()
        files = ftp.GetListing(fg, FtpListOption.AllFiles)
        Dim lst = From x In files Select x.Name
        list.AddRange(lst.ToList)
        files.Distinct()
        list.Remove("ALLINALLDATABASE")

        Dim list2 As New List(Of String)()
        For Each item In list
            If Not item = "ALLINALLDATABASE" Then
                files = ftp.GetListing(item & "/", FtpListOption.AllFiles)
                Dim lst2 = From x In files Where Not x.Name = "ALLINALLDATABASE" Select x.Name
                list2.AddRange(lst2.ToList)
            End If
        Next
        list2 = list2.Distinct().ToList
        files.Distinct()

        Dim list3 As New List(Of String)()
        For Each item In list
            For Each item1 In list2
                If Not item1.Contains("инструкции") Then
                    files = ftp.GetListing(item & "/" & item1 & "/", FtpListOption.AllFiles)
                    Dim lst3 = From x In files Select x.Name
                    list3.AddRange(lst3.ToList)
                    'listYear.AddRange(listFluentFTP2(item & "/" & item1 & "/"))
                End If
            Next
        Next
        list3 = list3.Distinct().ToList
        files.Distinct()

        Dim fgt As Dictionary(Of String, String)
        fgt = New Dictionary(Of String, String)
        For Each item In list
            For Each item1 In list2
                For Each item2 In list3
                    files = ftp.GetListing(item & "/" & item1 & "/" & item2 & "/", FtpListOption.AllFiles)

                    Dim var = files.ToDictionary(Function(mc) mc.Name.ToString(), Function(mc) mc.FullName.ToString, StringComparer.OrdinalIgnoreCase)
                    fgt = var.ToDictionary(Function(mc1) mc1.Key, Function(mc1) mc1.ToString)
                    If fgt.Count = 0 Then Continue For
                    For Each f In fgt
                        DirAll.Add(f.Key, f.Value)
                    Next

                Next
            Next
        Next





        ftp.Disconnect()





    End Sub
    Public Function listFTP(ByVal Директория As String) As List(Of String)
        Dim requ As FtpWebRequest = Nothing
        Dim resp As FtpWebResponse = Nothing
        Dim reader As StreamReader = Nothing

        Dim list As New List(Of String)()

        Try

            requ = CType(WebRequest.Create(FTPString & Директория), FtpWebRequest)
            requ.Credentials = New NetworkCredential(FTPUser, FTPPass)
            requ.Method = WebRequestMethods.Ftp.ListDirectory
            resp = CType(requ.GetResponse(), FtpWebResponse)
            reader = New StreamReader(resp.GetResponseStream())


            While (reader.Peek() > -1)
                list.Add(reader.ReadLine())
            End While
            resp.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Return list






    End Function

    Public Sub ПечатьДоковFTP(ByVal mass As ArrayList)

        Dim wdApp As New Microsoft.Office.Interop.Word.Application
        Dim wdDoc As Microsoft.Office.Interop.Word.Document
        wdApp.Visible = False
        Dim print As New List(Of String)()

        For x As Integer = 0 To mass.Count - 1
            Dim t As Integer = ВыгрузкаФайловНаЛокалыныйКомп(mass.Item(x)(0).ToString & mass.Item(x)(1).ToString, PathVremyanka & mass.Item(x)(1).ToString)
            If t = 1 Then Continue For
            print.Add(PathVremyanka & mass.Item(x)(1).ToString)
        Next

        For v As Integer = 0 To print.Count - 1
            wdDoc = wdApp.Documents.Open(FileName:=print(v).ToString)
            Try
                wdDoc.PrintOut(True) 'печать
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
            wdDoc.Close()
        Next
        wdApp.Quit()

        For v As Integer = 0 To print.Count - 1
            Try
                IO.File.Delete(print(v))
            Catch ex As Exception

            End Try

        Next



    End Sub
    Public Sub ПечатьДоковFTP(ByVal mass As ArrayList, ByVal int As Integer)

        Dim wdApp As New Microsoft.Office.Interop.Word.Application
        Dim wdDoc As Microsoft.Office.Interop.Word.Document
        wdApp.Visible = False
        Dim print As New List(Of String)()

        For x As Integer = 0 To mass.Count - 1
            Dim t As Integer = ВыгрузкаФайловНаЛокалыныйКомп(mass.Item(x)(0).ToString & mass.Item(x)(1).ToString, PathVremyanka & mass.Item(x)(1).ToString)
            If t = 1 Then Continue For
            print.Add(PathVremyanka & mass.Item(x)(1).ToString)
        Next

        For v As Integer = 0 To print.Count - 1
            wdDoc = wdApp.Documents.Open(FileName:=print(v).ToString)
            Try
                wdDoc.PrintOut(True,,,,,,, int) 'печать
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
            wdDoc.Close()
        Next
        wdApp.Quit()

        For v As Integer = 0 To print.Count - 1
            Try
                IO.File.Delete(print(v))
            Catch ex As Exception

            End Try

        Next



    End Sub
    Public Sub DatabasSave()
        Dim f, p As String
        f = "FTPUser1"
        p = "x&AkgAe&YmSw"
        Dim connstr As String = "ftp://45.14.50.13/"
        Dim ПутьНазначенияИИмяФайла As String = "B:\БазыSQL (Интегра и Рикманс)\"
        Dim str = "/"


        Dim SessionOptions As SessionOptions = New SessionOptions
        With SessionOptions
            .Protocol = Protocol.Ftp
            .HostName = "45.14.50.13" ' "ftp.example.com"ftp://86.57.135.184:21
            .UserName = "FTPUser1"
            .Password = "x&AkgAe&YmSw"
        End With

        Using session As Session = New Session()
            ' Connect
            session.Open(SessionOptions)

            ' Download files
            Try
                'session.GetFiles(ПутьОтСервера, ПутьСохранФайлаЛокально).Check()
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

            'session.GetFiles("/directory/to/download/*", "C:\target\directory\*").Check()
            session.Close()
        End Using










        'My.Computer.Network.DownloadFile(ПутьИсходногоФайла, ПутьНазначенияИИмяФайла, FTPUser, FTPPass, True, 500, True) 'выгрузка с сервера


        Using ftp As FluentFTP.FtpClient = New FluentFTP.FtpClient(connstr, f, p)
            'ftp.DataConnectionType = FtpDataConnectionType.AutoPassive

            ftp.Connect()

            Try
                ftp.DownloadFile(ПутьНазначенияИИмяФайла, str, True)
                ftp.Disconnect()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                ftp.Disconnect()
            End Try
        End Using


    End Sub


End Module
