Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.Threading
Imports System.IO
Imports WinSCP
Imports System.Linq
Imports System.IO.Compression
Imports System.Data.Linq
Imports System.Data.Linq.Mapping
Imports NLog
'Imports Zidium
'Imports System.ComponentModel
Imports Zidium.Api
Imports System.ComponentModel
Imports Zidium

Public Class MDIParent1

    Private Sub ShowNewForm(ByVal sender As Object, ByVal e As EventArgs)
        ' Создать новый экземпляр дочерней формы.
        Dim ChildForm As New System.Windows.Forms.Form
        ' Сделать ее дочерней для данной формы MDI перед отображением.
        ChildForm.MdiParent = Me

        m_ChildFormNumber += 1
        ChildForm.Text = "Окно " & m_ChildFormNumber

        ChildForm.Show()
    End Sub

    Private Sub OpenFile(ByVal sender As Object, ByVal e As EventArgs)
        Dim OpenFileDialog As New OpenFileDialog
        OpenFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        OpenFileDialog.Filter = "Текстовые файлы (*.txt)|*.txt|Все файлы (*.*)|*.*"
        If (OpenFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
            Dim FileName As String = OpenFileDialog.FileName
            ' TODO: добавьте здесь код открытия файла.
        End If
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim SaveFileDialog As New SaveFileDialog
        SaveFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        SaveFileDialog.Filter = "Текстовые файлы (*.txt)|*.txt|Все файлы (*.*)|*.*"

        If (SaveFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
            Dim FileName As String = SaveFileDialog.FileName
            ' TODO: добавить код для сохранения содержимого формы в файл.
        End If
    End Sub


    Private Sub ExitToolsStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.Close()
    End Sub

    Private Sub CutToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Использовать My.Computer.Clipboard для помещения выбранного текста или изображений в буфер обмена
    End Sub

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Использовать My.Computer.Clipboard для помещения выбранного текста или изображений в буфер обмена
    End Sub

    Private Sub PasteToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        'Использовать My.Computer.Clipboard.GetText() или My.Computer.Clipboard.GetData для получения информации из буфера обмена.
    End Sub

    'Private Sub ToolBarToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
    '    Me.ToolStrip.Visible = Me.ToolBarToolStripMenuItem.Checked
    'End Sub

    'Private Sub StatusBarToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
    '    Dim StatusBarToolStripMenuItem As ch
    '    Me.StatusStrip.Visible = Me.StatusBarToolStripMenuItem.Checked
    'End Sub

    Private Sub CascadeToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub TileVerticalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.TileVertical)
    End Sub

    Private Sub TileHorizontalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub ArrangeIconsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.ArrangeIcons)
    End Sub

    Private Sub CloseAllToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Закрыть все дочерние формы указанного родителя.
        For Each ChildForm As Form In Me.MdiChildren
            ChildForm.Close()
        Next
    End Sub

    Private m_ChildFormNumber As Integer

    Private Sub ПриемНаРаботуToolStripMenuItem_Click(sender As Object, e As EventArgs)
        'Dim F2 As Прием
        'F2.MdiParent = Me

        'F2.Show()
        'F2.WindowState = FormWindowState.Maximized
        Прием.WindowState = FormWindowState.Maximized
        Прием.Show()
    End Sub

    Private Sub ОрганизацияToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОрганизацияToolStripMenuItem.Click
        Try
            ActiveMdiChild.Close()
        Catch ex As Exception

        End Try


        'Dim F2 As Контрагент
        'F2.MdiParent = Me
        'F2.WindowState = FormWindowState.Maximized
        Контрагент.WindowState = FormWindowState.Maximized
        AddHandler Контрагент.FormClosing, AddressOf запускзакрытия
        Контрагент.CheckBox6.Checked = True
        Контрагент.Show()
    End Sub

    Private Sub Preload() 'работа linq to sql

        'Await Task.Run((Sub() ЗагрВсехФайловсСервера()))

        Dim db As New DataContext(ConString)

        Dim ПутиДокументов As Table(Of ПутиДок)

        ПутиДокументов = db.GetTable(Of ПутиДок)()

        'Dim f = From x In ПутиДокументов Where x.IDСотрудник = 1015 Select x.ДокМесто

        'Dim Клиент As Table(Of Клиент)

        'Клиент = db.GetTable(Of Клиент)()

        Dim f1 = db.GetTable(Of Клиент)().GroupBy(Function(u) u.НазвОрг).Select(Function(c) c.First)
        Dim f2 = db.GetTable(Of ПутиДок)().GroupBy(Function(u1) u1.Код).Select(Function(c1) c1.First)

        For Each g1 In f2
            'MessageBox.Show(g1.НазвОрг & " / " & g1.ФормаСобств & " / " & g1.УНП) ' r.Путь, r.ИмяФайла
            'MessageBox.Show(g1.IDСотрудник) ' r.Путь, r.ИмяФайла
        Next


        'обновление данных
        Dim Докместо As ПутиДок = New ПутиДок()
        Докместо = db.GetTable(Of ПутиДок).FirstOrDefault
        Докместо.ДокМесто = "Прием-Дог Подряд2"
        db.SubmitChanges()



        'добавление данных
        Dim NewString As New ПутиДок With {
        .IDСотрудник = 1965,
        .ДокМесто = "Place1"
        }
        db.GetTable(Of ПутиДок)().InsertOnSubmit(NewString)
        db.SubmitChanges()

        'удаление данных
        Dim var = db.GetTable(Of Test).OrderByDescending(Function(c) c.Код).FirstOrDefault
        db.GetTable(Of Test)().DeleteOnSubmit(var)
        db.SubmitChanges()



        'sql запросы

        'IEnumerable<User> users = db.ExecuteQuery<User>("SELECT * FROM Users WHERE Age>{0}", 23);

        'db.ExecuteQuery<User>("SELECT * FROM Users WHERE Age>{0}", 23);
        'Dim k = db.ExecuteCommand("SELECT * FROM Банк")

        'Dim h = From d In db.ExecuteQuery("SELECT * FROM Банк",,) Select d
        'For Each g1 In h
        '    'MessageBox.Show(g1.НазвОрг & " / " & g1.ФормаСобств & " / " & g1.УНП) ' r.Путь, r.ИмяФайла
        '    MessageBox.Show(g1.ToString) ' r.Путь, r.ИмяФайла
        'Next



    End Sub

    Private Sub Загрузка()
        UserAdmin = (Environment.GetEnvironmentVariable("USERNAME"))
        'Dim path = System.IO.Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName) 'образец не удалять
        'Dim oath As String = Application.StartupPath.ToString() & "\ИнтеграКадры.exe" 'образец не удалять

        'УдалениеФайлаНаСервере("ftp://86.57.135.184:21/Кадры/Должностные инструкции/5265 Руководители Директор4 7.doc")

        'ВыборкаФайловПоОпределенномуПути()
        'FluedЗагрузкаФайлов()

        'DatabasSave()
        Me.Cursor = Cursors.WaitCursor
        'allFilesAsync()
        Me.Cursor = Cursors.Default
        firthtPath = My.Application.Info.DirectoryPath & "\General"

        PathVremyanka = firthtPath & "\Времянка\"

        Parallel.Invoke(Sub() ВременнаяПапкаСоздание()) 'создание временной папки

        'Parallel.Invoke(Sub() Preload()) 'загрузка файлов с сервера на диск


        'OnePath = "C:\Users\" & UserAdmin & "\Google Диск\КадрыIntegraAll\"
        'If Not IO.Directory.Exists(OnePath) Then
        '    MsgBox("Не установлен Google Диск!" & vbCrLf & "Обратитесь к администратору!")
        '    Me.Close()
        '    Exit Sub
        'End If


        'Preload()


        'ВсплывФормаПриЗагр.ВсплывФорма()
        'ВсплывФормаПриЗагр.ShowDialog()



        Прием.MdiParent = Me
        ИмяКомп = My.Computer.Name




        'Pext("SELECT НазвОрг FROM Клиент ORDER BY НазвОрг", СписокКлиентовОсновной)

        СписокКлиентовОсновной = Selects(StrSql:="SELECT НазвОрг FROM Клиент ORDER BY НазвОрг")

        'Dim list As New Dictionary(Of String, String)
        'list.Add("@унп", "690674582")
        'list.Add("@тел", "ЗАО МТБАНК")
        'СписокКлиентовОсновной = Selects(StrSql:="SELECT * FROM Клиент WHERE УНП=@унп and Банк=@тел", list)

        ALLALL()
        'Dim gb As New Thread(Sub()
        '                         gb.IsBackground = True
        '                         gb.Start()



        Прием.Show()
        ВсплывФормапередЗагрузкой()
    End Sub
    Public Sub ОбработкаОшибок(ByVal ex As Exception, ByVal str As String)
        ex.Data.Add(My.Computer.Name, Date.Now)
        Logger.Error(ex, str)
    End Sub

    Private Sub loggers()
        Try
            Dim var As Integer = CType("bvc", Integer) 'логи ошибок
        Catch ex As Exception

            'Dim client1 = Client.Instance
            'Dim component = client1.GetDefaultComponentControl()

            'Dim unitTest = component.GetOrCreateUnitTestControl("MyUnitTest")
            'unitTest.SendResult(UnitTestResult.Success)

            'component.SendMetric("HDD", 1024)

            'Dim Logger = LogManager.GetCurrentClassLogger()
            ex.Data.Add(My.Computer.Name, Date.Now)
            Logger.Error(ex, "Комментарий к ошибке")

        End Try
        Exit Sub
    End Sub
    Private Sub MDIParent1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Загрузка()
    End Sub
    Sub GetFindSub(ByVal d As String, ByVal f As String)
        Dim list As New List(Of String)
        Try
            Application.DoEvents()
            Dim ListFiles As String() = IO.Directory.GetFiles(d, f)
            Dim ListFolders As String() = IO.Directory.GetDirectories(d)
            For Each item In ListFiles
                list.Add(item)
            Next
            For Each item In ListFolders
                GetFindSub(item, f)
            Next
        Catch ex As Exception

        End Try
    End Sub


    Private Sub SaveToolStripButton_Click(sender As Object, e As EventArgs)
        'Dim sd As VariantType
        'sd = CType(ActiveMdiChild, Прием).Activate


    End Sub
    Public Sub запускзакрытия(sender As Object, e As EventArgs)
        'If CType(sender, Контрагент).TextBox1.Text <> "" Then 'рабочий вариант
        '    MsgBox("Сохранить изменения!")
        'End If
    End Sub

    Private Sub УвольнениеToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Увольнение.WindowState = FormWindowState.Maximized
        Увольнение.Show()

    End Sub

    Private Sub ШтатноеToolStripMenuItem_Click(sender As Object, e As EventArgs)

        ШтатноеКласс1.Show()
        ШтатноеКласс1.WindowState = FormWindowState.Maximized
    End Sub



    Private Sub ПриемНаРаботуToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        Прием.WindowState = FormWindowState.Maximized
        AddHandler Контрагент.FormClosing, AddressOf запускзакрытия

        Прием.Show()
    End Sub

    Private Sub ШтатноеToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        ШтатноеКласс1.Show()
        ШтатноеКласс1.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub УвольнениеToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles УвольнениеToolStripMenuItem1.Click
        Try
            ActiveMdiChild.Close()
        Catch ex As Exception

        End Try
        Увольнение.WindowState = FormWindowState.Maximized
        Увольнение.Show()

    End Sub

    Private Sub ИзменитьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ИзменитьToolStripMenuItem.Click

        Try
            ActiveMdiChild.Close()
        Catch ex As Exception

        End Try



        Контрагент.WindowState = FormWindowState.Maximized
        AddHandler Контрагент.FormClosing, AddressOf запускзакрытия
        Контрагент.CheckBox6.Checked = False
        Контрагент.TextBox1.Enabled = False
        Контрагент.Show()
    End Sub



    Private Sub ПриемНаРаботуToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles ПриемНаРаботуToolStripMenuItem.Click

        Dim f As Form
        f = ActiveMdiChild

        'Прием.WindowState = FormWindowState.Maximized
        'AddHandler Штатное.FormClosing, AddressOf запускзакрытия
        Прием.CheckBox7.Checked = False
        Прием.ComboBox1.Text = ""
        Прием.MdiParent = Me
        Прием.Show()
        ПровФормы(f, "Прием")
    End Sub
    Private Sub ПровФормы(ByVal f As Form, ByVal g As String)
        If f Is Nothing Then
            Exit Sub
        End If
        If Not f.Name = g Then
            Try
                f.Close()
            Catch ex As Exception

            End Try

        End If
    End Sub
    Private Sub ПриказToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПриказToolStripMenuItem.Click
        Try
            ActiveMdiChild.Close()
        Catch ex As Exception

        End Try
        ПриказПродления.Show()

    End Sub

    Private Sub ПродлениеКонтрактаToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПродлениеКонтрактаToolStripMenuItem.Click
        Try
            ActiveMdiChild.Close()
        Catch ex As Exception

        End Try
        Уведомление.Show()
        Уведомление.WindowState = FormWindowState.Maximized
    End Sub


    Private Sub АктToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Try
            ActiveMdiChild.Close()
        Catch ex As Exception

        End Try
        Прием.WindowState = FormWindowState.Maximized
        'AddHandler Штатное.FormClosing, AddressOf запускзакрытия
        ВыборОрганизации.WindowState = FormWindowState.Normal
        ВыборОрганизации.Show()

    End Sub


    Private Sub АктДоговораПодрядаToolStripMenuItem_Click_1(sender As Object, e As EventArgs)
        ДоговорПодрядаАкт.ShowDialog()
    End Sub

    Private Sub PrintToolStripButton_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ToolTip_Popup(sender As Object, e As PopupEventArgs) Handles ToolTip.Popup

    End Sub

    Private Sub УволенныеToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles УволенныеToolStripMenuItem.Click
        Try
            'Соед(1)
            ActiveMdiChild.Close()
        Catch ex As Exception

        End Try
        Уволенные.Show()
        'Уволенные.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub ПринятыеToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПринятыеToolStripMenuItem.Click
        Try
            ActiveMdiChild.Close()
        Catch ex As Exception

        End Try
        ПринятыеСписки.Show()
        ПринятыеСписки.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub ПечатьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПечатьToolStripMenuItem.Click
        Try
            ActiveMdiChild.Close()
        Catch ex As Exception

        End Try
        Печать.Show()
    End Sub

    Private Sub ШтатноеToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles ШтатноеToolStripMenuItem.Click
        Try
            ActiveMdiChild.Close()
        Catch ex As Exception

        End Try
        ШтатноеСписки.Show()
    End Sub

    Private Sub ОтчетыToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОтчетыToolStripMenuItem.Click

    End Sub

    Private Sub ДоговорПодрядаToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ДоговорПодрядаToolStripMenuItem1.Click
        Try
            ActiveMdiChild.Close()
        Catch ex As Exception

        End Try
        ДоговорПодрядаСписки.Show()

    End Sub

    Private Sub ПриказыToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ПринятыеToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        'Try
        '    ActiveMdiChild.Close()
        'Catch ex As Exception

        'End Try
    End Sub

    Private Sub УволенныеToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        'Try
        '    ActiveMdiChild.Close()
        'Catch ex As Exception

        'End Try
    End Sub

    Private Sub ПродлениеКонтрактаToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        'Try
        '    ActiveMdiChild.Close()
        'Catch ex As Exception

        'End Try
    End Sub

    Private Sub ДопПоЗарплатеToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ToolStrip_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs)

    End Sub



    Private Sub ПослеToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПослеToolStripMenuItem.Click

        Try
            ActiveMdiChild.Close()
        Catch ex As Exception

        End Try
        ШтатноеПослеИзменения.WindowState = FormWindowState.Maximized
        'AddHandler Штатное.FormClosing, AddressOf запускзакрытия
        ШтатноеПослеИзменения.Show()
    End Sub

    Private Sub КласическоеToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles КласическоеToolStripMenuItem.Click
        Try
            ActiveMdiChild.Close()
        Catch ex As Exception

        End Try
        ШтатноеКласс1.WindowState = FormWindowState.Maximized
        'AddHandler Штатное.FormClosing, AddressOf запускзакрытия
        ШтатноеКласс1.Show()
    End Sub

    Private Sub ПоискToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПоискToolStripMenuItem.Click

        ActiveMdiChild.Close()
        'Try
        'ActiveMdiChild.Close()
        'Catch ex As Exception

        'End Try

        Поиск.Show()
        Поиск.WindowState = FormWindowState.Maximized

    End Sub

    Private Sub ДопРестлайнToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ПереводToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПереводToolStripMenuItem.Click
        Перевод.ShowDialog()
    End Sub

    'Private Sub НеподписанныедокументыToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles НеподписанныедокументыToolStripMenuItem.Click
    '    Try
    '        ActiveMdiChild.Close()
    '    Catch ex As Exception

    '    End Try
    '    НеподпДокументы.Show()
    'End Sub

    Private Sub ОбУровнеЗарплатыToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОбУровнеЗарплатыToolStripMenuItem.Click
        СправкаПоЗарплате.ShowDialog()
    End Sub

    Private Sub ИностранцыToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ИностранцыToolStripMenuItem.Click
        Иностранцы.ShowDialog()
    End Sub

    Private Sub СтатистикаToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СтатистикаToolStripMenuItem.Click
        ПарольНаСтатистику.ShowDialog()

        'Статистикаc.ShowDialog()
    End Sub

    Private Sub ТекущийToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ТекущийToolStripMenuItem.Click
        Try
            ActiveMdiChild.Close()
        Catch ex As Exception

        End Try
        Отпуск.WindowState = FormWindowState.Maximized
        'AddHandler Штатное.FormClosing, AddressOf запускзакрытия
        Отпуск.Show()
    End Sub

    Private Sub ПоТребованиюToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПоТребованиюToolStripMenuItem.Click
        ОтпускСоц.ShowDialog()

    End Sub

    Private Sub ПробаToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Прием2.Show()
    End Sub

    Private Sub ОтправкаEmailToolStripMenuItem_Click(sender As Object, e As EventArgs)
        SendEMail.ShowDialog()
    End Sub

    Private Sub УведомлениеОПродленииКонтрактаToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles УведомлениеОПродленииКонтрактаToolStripMenuItem.Click
        'Dim df As Threading.Thread = New Threading.Thread(AddressOf ВсплывФормаПриЗагр.ВсплывФорма)
        'df.SetApartmentState(ApartmentState.STA)
        'df.Start()

        ВсплывФормаПриЗагр.ВсплывФорма()
    End Sub

    Private Sub ПоЧасамToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПоЧасамToolStripMenuItem.Click
        ДоговорПодрядаАкт.ShowDialog()
    End Sub

    Private Sub ТикToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Form3.ShowDialog()
    End Sub

    Private Sub УведомлениеОбИзмененииСроковВыплатыЗарплатыToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles УведомлениеОбИзмененииСроковВыплатыЗарплатыToolStripMenuItem.Click
        Try
            ActiveMdiChild.Close()
        Catch ex As Exception

        End Try
        ДопПоСрокамОплаты.Show()
    End Sub

    Private Sub УведомлениеОбИзмененииУслвоияТрудаРестлайнToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles УведомлениеОбИзмененииУслвоияТрудаРестлайнToolStripMenuItem.Click
        ДопРестлайн.ShowDialog(Me)
    End Sub

    Private Sub УведомлениеОбИзмененеииОкладаToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles УведомлениеОбИзмененеииОкладаToolStripMenuItem.Click
        Try
            ActiveMdiChild.Close()
        Catch ex As Exception

        End Try
        КонтрДопИзменСтавка.Show()
    End Sub

    Private Sub ОтпускToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОтпускToolStripMenuItem.Click
        Try
            ActiveMdiChild.Close()
        Catch ex As Exception

        End Try
        ОтпускСписки.Show()
    End Sub

    Private Sub СотрудникиToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СотрудникиToolStripMenuItem.Click
        Try
            ActiveMdiChild.Close()
        Catch ex As Exception

        End Try
        ОтчетСотрудники.Show()
    End Sub

    Private Sub MenuStrip_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles MenuStrip.ItemClicked

    End Sub

    Private Sub НеподписанныеDokumentyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles НеподписанныеDokumentyToolStripMenuItem.Click
        Try
            ActiveMdiChild.Close()
        Catch ex As Exception

        End Try
        НеподпДокументы.Show()
    End Sub

    Private Sub MDIParent1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Api.Client.Instance.EventManager.Flush()
        Api.Client.Instance.WebLogManager.Flush()
    End Sub

    Private Sub ИзменитьToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ИзменитьToolStripMenuItem1.Click
        ДоговорПодрядаАктИное.ShowDialog()
    End Sub

    Private Sub СоздатьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СоздатьToolStripMenuItem.Click
        ДогПодрядаАктИноеСоздать.ShowDialog()
    End Sub

    Private Sub ДобавитьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ДобавитьToolStripMenuItem.Click
        СправочникСотрудники.BtnClick = "Добавить"
        СправочникСотрудники.ShowDialog()
    End Sub

    Private Sub ИзменитьToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ИзменитьToolStripMenuItem2.Click
        СправочникСотрудники.BtnClick = "Изменить"
        СправочникСотрудники.ShowDialog()
    End Sub

    Private Sub СНПРВToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СНПРВToolStripMenuItem.Click
        СНПРВФ.ShowDialog()
    End Sub
End Class
