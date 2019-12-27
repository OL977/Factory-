Option Explicit On
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Linq
Imports System.Linq.Dynamic

Public Class Уведомление
    Public Da As New OleDbDataAdapter 'Адаптер
    Public Ds As New DataSet 'Пустой набор записей
    Dim tbl As New DataTable
    Dim cb As OleDb.OleDbCommandBuilder
    'PrintУвол, PrintПрик, PrintКонтр, PrintУвед
    Dim ПровДанн As Integer
    Dim sd As Boolean
    Dim prprov1 As Boolean = False
    Dim prprov2 As Boolean = False
    Dim УволFTP As New List(Of String)()
    Dim УведомлFTP As New List(Of String)()
    Dim ДопСоглFTP As New List(Of String)()

    Dim ДатаНач, ДатаОконч, Организ, idClient, год2, год, КорИмя, Коротч, ПровПродКонтр,
        ПродлКонтрС, СрокОкончКонтр, ПродлКонтрПо As String
    Public name2 As String
    Dim ФормаСобстПолн, ЭлАдрес, Банк, БИК, АдресБанка, РасСчет, ЮрАдрес, УНП, ДолжРуков, ФИОРукРодПад, ОснованиеДейств, МестоРаб, ФИОКор, ФормаСобствКор, СборноеРеквПолн, ДолжРуковРодПад, ДолжРуковВинПад, КонтТелефон As String


    Private Sub Уведомление_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MdiParent = MDIParent1
        Me.WindowState = FormWindowState.Maximized
        год2 = Year(Now)

        год = Year(Now)
        'If Me.Прием_Load = vbTrue Then Form1.Load = False

        Me.ComboBox1.Items.Clear()
        For Each r As DataRow In СписокКлиентовОсновной.Rows
            Me.ComboBox1.Items.Add(r(0).ToString)
        Next
        Me.TextBox1.Text = DateTime.Now.ToString("dd.MM.yyyy")


        DateTimePicker1.Format = DateTimePickerFormat.Short
        DateTimePicker2.Format = DateTimePickerFormat.Short

    End Sub



    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs)
        'Grid1.DataSource = Nothing
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        refreshgrid()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)
        PrintDialog1.ShowDialog()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        PrintPreviewDialog1.ShowDialog()
    End Sub

    Private Sub Grid1_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles Grid1.CellMouseClick
        If e.Button = MouseButtons.Right Then
            proverka = 0
            ВсплывФорма()
        ElseIf e.Button = MouseButtons.Left Then
            If IsDBNull(Grid1.CurrentRow.Cells(0).Value) Then Exit Sub
            Dim ind As String = Grid1.CurrentRow.Cells(0).Value
                Dim dt As New DataTable
                dt = TryCast(Grid1.DataSource, DataTable) 'перевод из Datatgridviw в datatable

                Dim ft = dt.Select("ID=" & ind & "")

                Dim f = From x In dt.Rows Where x.Item(0) = ind Select x  'Where x.index = ind
                'Dim ft = DirectCast(DataGrid.ItemContainerGenerator.ContainerFromIndex(ind), DataGridRow)

                TextBox3.Text = ft(0).Item(2).ToString
            MaskedTextBox1.Text = ft(0).Item("Дата приема сотрудника").ToString
            MaskedTextBox2.Text = ft(0).Item("Дата окончания контракта").ToString
            MaskedTextBox3.Text = ft(0).Item("Дата уведомления о продлении контракта").ToString
            TextBox7.Text = ft(0).Item("Номер уведомления").ToString
                TextBox8.Text = ft(0).Item("Период продления, год").ToString
            TextBox9.Text = Strings.Left(ft(0).Item("Продление контракта По").ToString, 10)
            TextBox10.Text = Strings.Left(ft(0).Item("Продление контракта с").ToString, 10)
            ComboBox2.Text = ft(0).Item("Продление").ToString
            Label12.Text = ind
        End If
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim mRow As Integer = 0
        Dim newpage As Boolean = True
        ' sets it to show '...' for long text
        Dim fmt As StringFormat = New StringFormat(StringFormatFlags.LineLimit)
        fmt.LineAlignment = StringAlignment.Center
        fmt.Trimming = StringTrimming.EllipsisCharacter
        Dim y As Int32 = e.MarginBounds.Top
        Dim rc As Rectangle
        Dim x As Int32
        Dim h As Int32 = 0
        Dim row As DataGridViewRow

        ' print the header text for a new page
        '   use a grey bg just like the control
        If newpage Then
            row = Grid1.Rows(mRow)
            x = e.MarginBounds.Left
            For Each cell As DataGridViewCell In row.Cells
                ' since we are printing the control's view,
                ' skip invidible columns
                If cell.Visible Then
                    rc = New Rectangle(x, y, cell.Size.Width, cell.Size.Height)

                    e.Graphics.FillRectangle(Brushes.LightGray, rc)
                    e.Graphics.DrawRectangle(Pens.Black, rc)

                    ' reused in the data pront - should be a function
                    Select Case Grid1.Columns(cell.ColumnIndex).DefaultCellStyle.Alignment
                        Case DataGridViewContentAlignment.BottomRight,
                         DataGridViewContentAlignment.MiddleRight
                            fmt.Alignment = StringAlignment.Far
                            rc.Offset(-1, 0)
                        Case DataGridViewContentAlignment.BottomCenter,
                        DataGridViewContentAlignment.MiddleCenter
                            fmt.Alignment = StringAlignment.Center
                        Case Else
                            fmt.Alignment = StringAlignment.Near
                            rc.Offset(2, 0)
                    End Select

                    e.Graphics.DrawString(Grid1.Columns(cell.ColumnIndex).HeaderText,
                                            Grid1.Font, Brushes.Black, rc, fmt)
                    x += rc.Width
                    h = Math.Max(h, rc.Height)
                End If
            Next
            y += h

        End If
        newpage = False

        ' now print the data for each row
        Dim thisNDX As Int32
        For thisNDX = mRow To Grid1.RowCount - 1
            ' no need to try to print the new row
            If Grid1.Rows(thisNDX).IsNewRow Then Exit For

            row = Grid1.Rows(thisNDX)
            x = e.MarginBounds.Left
            h = 0

            ' reset X for data
            x = e.MarginBounds.Left

            ' print the data
            For Each cell As DataGridViewCell In row.Cells
                If cell.Visible Then
                    rc = New Rectangle(x, y, cell.Size.Width, cell.Size.Height)

                    ' SAMPLE CODE: How To 
                    ' up a RowPrePaint rule
                    'If Convert.ToDecimal(row.Cells(5).Value) < 9.99 Then
                    '    Using br As New SolidBrush(Color.MistyRose)
                    '        e.Graphics.FillRectangle(br, rc)
                    '    End Using
                    'End If

                    e.Graphics.DrawRectangle(Pens.Black, rc)

                    Select Case Grid1.Columns(cell.ColumnIndex).DefaultCellStyle.Alignment
                        Case DataGridViewContentAlignment.BottomRight,
                         DataGridViewContentAlignment.MiddleRight
                            fmt.Alignment = StringAlignment.Far
                            rc.Offset(-1, 0)
                        Case DataGridViewContentAlignment.BottomCenter,
                        DataGridViewContentAlignment.MiddleCenter
                            fmt.Alignment = StringAlignment.Center
                        Case Else
                            fmt.Alignment = StringAlignment.Near
                            rc.Offset(2, 0)
                    End Select

                    e.Graphics.DrawString(cell.FormattedValue.ToString(),
                                      Grid1.Font, Brushes.Black, rc, fmt)

                    x += rc.Width
                    h = Math.Max(h, rc.Height)
                End If

            Next
            y += h
            ' next row to print
            mRow = thisNDX + 1

            If y + h > e.MarginBounds.Bottom Then
                e.HasMorePages = True
                ' mRow -= 1   causes last row to rePrint on next page
                newpage = True
                Return
            End If
        Next


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        PrintPreviewDialog1.Document = PrintDocument1
        PrintPreviewDialog1.ShowDialog()
    End Sub

    'Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage

    '    With Grid1
    '        Dim fmt As StringFormat = New StringFormat(StringFormatFlags.LineLimit)
    '        fmt.LineAlignment = StringAlignment.Center
    '        fmt.Trimming = StringTrimming.EllipsisCharacter
    '        Dim y As Single = e.MarginBounds.Top
    '        Do While mRow < .RowCount
    '            Dim row As DataGridViewRow = .Rows(mRow)
    '            Dim x As Single = e.MarginBounds.Left
    '            Dim h As Single = 0
    '            For Each cell As DataGridViewCell In row.Cells
    '                Dim rc As RectangleF = New RectangleF(x, y, cell.Size.Width, cell.Size.Height)
    '                e.Graphics.DrawRectangle(Pens.Black, rc.Left, rc.Top, rc.Width, rc.Height)
    '                If (newpage) Then
    '                    e.Graphics.DrawString(Grid1.Columns(cell.ColumnIndex).HeaderText, .Font, Brushes.Black, rc, fmt)
    '                Else
    '                    e.Graphics.DrawString(Grid1.Rows(cell.RowIndex).Cells(cell.ColumnIndex).FormattedValue.ToString(), .Font, Brushes.Black, rc, fmt)
    '                End If
    '                x += rc.Width
    '                h = Math.Max(h, rc.Height)
    '            Next
    '            newpage = False
    '            y += h
    '            mRow += 1
    '            If y + h > e.MarginBounds.Bottom Then
    '                e.HasMorePages = True
    '                mRow -= 1
    '                newpage = True
    '                Exit Sub
    '            End If
    '        Loop
    '        mRow = 0
    '    End With
    '    End
    'End Sub
    Private Sub СотрСпис()

        'Dim времянач As String = Format(DateTimePicker1.Value, "MM\/dd\/yyyy")
        'Dim времякон As String = Format(DateTimePicker2.Value, "MM\/dd\/yyyy")

        Dim времянач As String = Replace(Format(DateTimePicker1.Value, "yyyy\/MM\/dd"), "\", "")
        Dim времякон As String = Replace(Format(DateTimePicker2.Value, "yyyy\/MM\/dd"), "\", "")

        Dim ds = Selects(StrSql:="SELECT Сотрудники.ФИОСборное
From Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр
WHERE Сотрудники.НазвОрганиз='" & Организ & "' AND 
КарточкаСотрудника.ДатаУведомлПродКонтр Between '" & времянач & "' And '" & времякон & "' ORDER BY Сотрудники.ФИОСборное")


        'Me.ComboBox2.Items.Clear()
        'For Each r As DataRow In ds.Rows
        '    Me.ComboBox2.Items.Add(r(0).ToString)
        'Next



    End Sub




    Private Sub refreshgrid()
        Организ = ComboBox1.Text


        'Dim времянач As String = Format(DateTimePicker1.Value, "MM\/dd\/yyyy")
        'Dim времякон As String = Format(DateTimePicker2.Value, "MM\/dd\/yyyy")

        'Dim времянач As String = Replace(Format(DateTimePicker1.Value, "yyyy\/MM\/dd"), "/", "")
        'Dim времякон As String = Replace(Format(DateTimePicker2.Value, "yyyy\/MM\/dd"), "/", "")

        Dim времянач As Date = Format(DateTimePicker1.Value, "yyyy\/MM\/dd")
        Dim времякон As Date = Format(DateTimePicker2.Value, "yyyy\/MM\/dd")
        tbl.Clear()


        Dim list As New Dictionary(Of String, Object)()

        'Dim list As New List(Of Date)()
        list.Add("@времянач", DateTimePicker1.Value)
        list.Add("@времякон", DateTimePicker2.Value)
        list.Add("@Организ", Организ)


        'Dim sqlparams As New List(Of SqlParameter())
        'Dim s As New Dictionary(Of String, Object)
        's.Add("@времянач", времянач)
        's.Add("@времякон", времякон)
        'Dim fg As New Object()
        'fg.времянач
        'fg(1) = времякон

        tbl = Selects(StrSql:="SELECT Сотрудники.КодСотрудники as [ID], Сотрудники.НазвОрганиз as Наименование, 
Сотрудники.ФИОСборное as [ФИО Сотрудника], КарточкаСотрудника.ДатаПриема as [Дата приема сотрудника],
ДогСотрудн.СрокОкончКонтр as [Дата окончания контракта], КарточкаСотрудника.ДатаУведомлПродКонтр as [Дата уведомления о продлении контракта],
КарточкаСотрудника.НомерУведомлПродКонтр as [Номер уведомления],КарточкаСотрудника.СрокПродлКонтракта as [Период продления, год],
КарточкаСотрудника.НеПродлениеКонтр as [Продление], КарточкаСотрудника.ПродлКонтрС as [Продление контракта с],
КарточкаСотрудника.ПродлКонтрПо as [Продление контракта По]
        From (Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр) INNER JOIN ДогСотрудн ON Сотрудники.КодСотрудники = ДогСотрудн.IDСотр
        Where Сотрудники.НазвОрганиз =@Организ And КарточкаСотрудника.ДатаУведомлПродКонтр BETWEEN @времянач AND @времякон AND (КарточкаСотрудника.НеПродлениеКонтр is null OR КарточкаСотрудника.НеПродлениеКонтр='False')  ORDER BY Сотрудники.ФИОСборное", list)


        For Each r As DataRow In tbl.Rows
            If r.Item("Продление").ToString = "False" Then
                r.Item("Продление") = "Да"
            End If
        Next

        Grid1.DataSource = tbl
        Grid1.Columns(0).Visible = False
        'Grid1.Columns(2).Width = 320

        For r As Integer = 0 To Grid1.Rows.Count - 1
            Grid1.Rows(r).Cells(6).Value = ""
            Grid1.Rows(r).Cells(7).Value = ""
        Next
        GridView2(Grid1)

        'cb = New OleDb.OleDbCommandBuilder(da)

        'Grid1.DataSource = ds
        'Grid1.DataMember = "Сотрудники"
        'СотрСпис()

    End Sub

    Private Sub ДопПродлКонтр(ByVal Ds As DataTable, ByVal Mass As Integer, ByVal СклонВремя As String, ByVal ДатаОтвета As String)

        ДобОконч(ДолжРуков)

        Dim oWord2 As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc2 As Microsoft.Office.Interop.Word.Document
        oWord2 = CreateObject("Word.Application")
        oWord2.Visible = False

        Начало("DopSoglashenie.doc")
        oWordDoc2 = oWord2.Documents.Add(firthtPath & "\DopSoglashenie.doc")

        With oWordDoc2.Bookmarks
            .Item("ДопС1").Range.Text = Strings.Left(Ds.Rows(0).Item(7).ToString, 10) & "г" 'дата контракта
            .Item("ДопС2").Range.Text = Ds.Rows(0).Item(6).ToString 'номер контракта
            .Item("ДопС3").Range.Text = ДатаОтвета & "г"
            .Item("ДопС4").Range.Text = Ds.Rows(0).Item(13).ToString 'фио сборное
            .Item("ДопС7").Range.Text = Ds.Rows(0).Item(9).ToString ' срок продления контракта
            .Item("ДопС8").Range.Text = СклонВремя 'склонение времени
            .Item("ДопС9").Range.Text = Strings.Left(Ds.Rows(0).Item(10).ToString, 10) & "г" ' срок продления контракта с
            .Item("ДопС10").Range.Text = Strings.Left(Ds.Rows(0).Item(11).ToString, 10) & "г" ' срок продления контракта по
            .Item("ДопС11").Range.Text = Ds.Rows(0).Item(1).ToString & " " & КорИмя & " " & Коротч
            '.Item("ДопС12").Range.Text = КорИмя
            '.Item("ДопС13").Range.Text = Коротч
            .Item("ДопС14").Range.Text = ФормаСобстПолн
            .Item("ДопС15").Range.Text = ComboBox1.Text
            .Item("ДопС16").Range.Text = ДобОконч(ДолжРуков)
            .Item("ДопС17").Range.Text = ФИОРукРодПад
            .Item("ДопС18").Range.Text = ОснованиеДейств
            .Item("ДопС19").Range.Text = ФормаСобствКор & " """ & ComboBox1.Text & """ "
            .Item("ДопС20").Range.Text = ФИОКор
            .Item("ДопС21").Range.Text = ПровДанн

        End With
        Dim ВремяС As String = Strings.Right(Strings.Left(Ds.Rows(0).Item(10).ToString, 10), 4)

        Dim name As String = "ДопСогл " & ПровДанн & "_" & ВремяС & " " & Ds.Rows(0).Item(1).ToString & " " & КорИмя & " " & Коротч & "(Доп.Продл.Контр)" & ".doc"

        ДопСоглFTP.AddRange(New String() {ComboBox1.Text & "\Дополнительное солгашение\" & Now.Year, name})

        oWordDoc2.SaveAs2(PathVremyanka & name,,,,,, False)
        oWordDoc2.Close(True)
        oWord2.Quit(True)
        Конец(ComboBox1.Text & "\Дополнительное солгашение\" & Now.Year, name, Mass, ComboBox1.Text, "\DopSoglashenie.doc", "ДопCолгашПродлКонтракта")


    End Sub




    Private Sub УведПродлКонтр(ByVal ДолжРодПадеж As String, ByVal Ds As DataTable, ByVal dfc() As DataRow, ByVal Mass As Integer, ByVal ДатаУвед As String,
                               ByVal СрокПродлКонтрПроп As String, ByVal СклонВремя As String, ByVal ДатаОтвета As String)

        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document
        oWord = CreateObject("Word.Application")
        oWord.Visible = False


        Начало("Uvedomlenie.doc")
        oWordDoc = oWord.Documents.Add(firthtPath & "\Uvedomlenie.doc")

        With oWordDoc.Bookmarks
            .Item("Увед1").Range.Text = ДолжРодПадеж
            .Item("Увед2").Range.Text = Ds.Rows(0).Item(1).ToString
            .Item("Увед3").Range.Text = КорИмя
            .Item("Увед4").Range.Text = Коротч
            .Item("Увед5").Range.Text = ДатаУвед & "г." 'дата уведомления
            .Item("Увед6").Range.Text = Ds.Rows(0).Item(5).ToString 'номер уведомления
            .Item("Увед7").Range.Text = Ds.Rows(0).Item(6).ToString 'номер контракта
            .Item("Увед8").Range.Text = Strings.Left(Ds.Rows(0).Item(7).ToString, 10) & "г" 'дата контракта
            Dim fg As Date = CDate(dfc(0).Item("ПродлКонтрС").ToString) 'дата продления с
            fg = fg.AddYears(-(Ds.Rows(0).Item(9)))
            If dfc(0).Item("ПродлКонтрС") = Nothing Then
                .Item("Увед9").Range.Text = Strings.Left(Ds.Rows(0).Item(7).ToString, 10) & "г" 'дата контракта
            Else
                .Item("Увед9").Range.Text = fg & "г" 'старая дата с
            End If

            Dim fg2 As Date = CDate(dfc(0).Item("ПродлКонтрПо").ToString) 'старая дата по
            fg2 = fg2.AddYears(-(Ds.Rows(0).Item(9)))
            .Item("Увед10").Range.Text = fg2 & "г" 'дата окончания контракта
            .Item("Увед11").Range.Text = fg2 & "г" 'дата окончания контракта
            .Item("Увед12").Range.Text = Ds.Rows(0).Item(9).ToString ' срок продления контракта
            .Item("Увед13").Range.Text = СрокПродлКонтрПроп 'прописью
            .Item("Увед14").Range.Text = СклонВремя 'склонение времени
            .Item("Увед15").Range.Text = Strings.Left(Ds.Rows(0).Item(10).ToString, 10) & "г" ' срок продления контракта с
            .Item("Увед16").Range.Text = Strings.Left(Ds.Rows(0).Item(11).ToString, 10) & "г" ' срок продления контракта по
            .Item("Увед17").Range.Text = ДатаОтвета & "г"
            .Item("Увед18").Range.Text = Ds.Rows(0).Item(1).ToString
            .Item("Увед19").Range.Text = КорИмя
            .Item("Увед20").Range.Text = Коротч
            .Item("Увед21").Range.Text = ФормаСобстПолн
            .Item("Увед22").Range.Text = Me.ComboBox1.Text
            .Item("Увед23").Range.Text = ФормаСобствКор
            .Item("Увед24").Range.Text = Me.ComboBox1.Text
            .Item("Увед25").Range.Text = ДолжРуков
            .Item("Увед26").Range.Text = ФормаСобствКор
            .Item("Увед27").Range.Text = Me.ComboBox1.Text
            .Item("Увед28").Range.Text = ФИОКор
        End With
        Dim НомУвед As String = Ds.Rows(0).Item(5).ToString

        'If Not IO.Directory.Exists(OnePath & Me.ComboBox1.Text & "\Уведомление\" & Year(Now)) Then
        '    IO.Directory.CreateDirectory(OnePath & Me.ComboBox1.Text & "\Уведомление\" & Year(Now))
        'End If

        Dim ИмяФайла As String = НомУвед & " " & Ds.Rows(0).Item(1).ToString & " " & КорИмя & " " & Коротч & "(Увед.Продл.Контр)" & ".doc"
        УведомлFTP.AddRange(New String() {ComboBox1.Text & "\Уведомление\" & Now.Year, ИмяФайла})
        oWordDoc.SaveAs2(PathVremyanka & ИмяФайла,,,,,, False)
        oWordDoc.Close(True)
        oWord.Quit(True)
        Конец(ComboBox1.Text & "\Уведомление\" & Now.Year, ИмяФайла, Mass, ComboBox1.Text, "\Uvedomlenie.doc", "Увед.Продл.Контр")

    End Sub

    Private Sub УведУвольнение(ByVal ДолжРодПадеж As String, ByVal разрядпроп As String, ByVal inp As String,
                               ByVal Ds As DataTable, ByVal dfc() As DataRow, ByVal Mass As Integer)

        Dim oWord1 As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc1 As Microsoft.Office.Interop.Word.Document
        'Dim oWordPara As Microsoft.Office.Interop.Word.Paragraph

        oWord1 = CreateObject("Word.Application")
        oWord1.Visible = False


        Начало("UvedObUvolnenii.doc")

        oWordDoc1 = oWord1.Documents.Add(firthtPath & "\UvedObUvolnenii.doc")

        With oWordDoc1.Bookmarks
            .Item("УвеУвол1").Range.Text = ФормаСобстПолн & " """ & ComboBox1.Text & """ " 'формсобств полное и название
            .Item("УвеУвол2").Range.Text = ДолжРодПадеж & " " & разрядпроп
            .Item("УвеУвол3").Range.Text = inp & " " & КорИмя & Коротч
            .Item("УвеУвол4").Range.Text = Strings.Left(Ds.Rows(0).Item(4).ToString, 10) & "г." ' 
            .Item("УвеУвол5").Range.Text = Ds.Rows(0).Item(5).ToString 'номер уведомления
            .Item("УвеУвол6").Range.Text = Strings.Left(Ds.Rows(0).Item(7).ToString, 10) 'дата контракта я закончил на этом месте 16.12.18
            .Item("УвеУвол7").Range.Text = Ds.Rows(0).Item(6).ToString 'номер контракта
            .Item("УвеУвол8").Range.Text = Strings.Left(dfc(0).Item("ПродлКонтрС").ToString, 10) 'дата контракта c
            .Item("УвеУвол9").Range.Text = Strings.Left(dfc(0).Item("ПродлКонтрПо").ToString, 10) 'дата контракта по
            .Item("УвеУвол10").Range.Text = Strings.Left(dfc(0).Item("ПродлКонтрПо").ToString, 10)
            .Item("УвеУвол11").Range.Text = ДолжРуков 'должность руководителя
            .Item("УвеУвол12").Range.Text = ФормаСобствКор & " " & ComboBox1.Text
            .Item("УвеУвол13").Range.Text = ФИОКор ' 
            .Item("УвеУвол14").Range.Text = Ds.Rows(0).Item(1).ToString & " " & КорИмя & Коротч
            .Item("УвеУвол15").Range.Text = Strings.Left(Ds.Rows(0).Item(4).ToString, 10)
            .Item("УвеУвол16").Range.Text = Ds.Rows(0).Item(1).ToString & " " & КорИмя & Коротч
            .Item("УвеУвол17").Range.Text = Strings.Left(Ds.Rows(0).Item(4).ToString, 10)

        End With
        Dim Name1 As String = Ds.Rows(0).Item(5).ToString & " " & Ds.Rows(0).Item(1).ToString & " " & КорИмя & " " & Коротч & "(Увед.Увольнение)" & ".doc"
        УволFTP.AddRange(New String() {ComboBox1.Text & "\Уведомление-увольнение\" & Now.Year, Name1})
        oWordDoc1.SaveAs2(PathVremyanka & Name1,,,,,, False)
        oWordDoc1.Close(True)
        oWord1.Quit(True)
        Конец(ComboBox1.Text & "\Уведомление-увольнение\" & Now.Year, Name1, Mass, ComboBox1.Text, "\UvedObUvolnenii.doc", "Увед.Увольнение")

    End Sub


    Private Sub Доки(ByVal Mass As Integer)

        'Dim sw As New Stopwatch 'вычисление выполнения метода (время)
        'sw.Start()

        Dim list As New Dictionary(Of String, Object)()        '
        list.Add("@НазвОрганиз", ComboBox1.Text)
        list.Add("@ID", CType(Label12.Text, Integer))
        'list.Add("@dt3", lb)

        Me.Cursor = Cursors.WaitCursor

        'выборка по клиенту(реквититы)
        Dim ds8 As DataTable = Selects(StrSql:="Select Клиент.ФормаСобств, Клиент.УНП, Клиент.ЮрАдрес,
Клиент.Банк, Клиент.БИКБанка, Клиент.АдресБанка, Клиент.Отделение, Клиент.РасчСчетРубли, Клиент.ДолжнРуководителя,
Клиент.ФИОРуководителя, Клиент.ОснованиеДейств, Клиент.ФИОРукРодПадеж, Клиент.КонтТелефон, Клиент.ЭлАдрес, Клиент.РукИП From Клиент
Where Клиент.НазвОрг =@НазвОрганиз", list)

        Dim РуковИП As String
        If ds8.Rows(0).Item(14) = True Or ds8.Rows(0).Item(14) = "True" Then
            РуковИП = "ИП "
        Else
            РуковИП = ""
        End If

        ФормаСобстПолн = ds8.Rows(0).Item(0).ToString
        ДолжРуков = ds8.Rows(0).Item(8).ToString
        ФИОРукРодПад = РуковИП & ds8.Rows(0).Item(11).ToString
        ОснованиеДейств = ds8.Rows(0).Item(10).ToString
        'МестоРаб = n & " " & w & " " & ComboBox18.Text

        'короткое фио клиента
        Dim nm As String = ds8.Rows(0).Item(9).ToString
        Dim nm0 As Integer = Len(ds8.Rows(0).Item(9).ToString)
        Dim nm1 As String = Strings.Left(nm, InStr(nm, " "))
        Dim nm2 As Integer = Len(nm1)
        Dim nm3 As String = Strings.Right(nm, (nm0 - nm2))
        Dim nm31 As Integer = Len(nm3)
        Dim nm4 As String = Strings.UCase(Strings.Left(Strings.Left(nm3, InStr(nm3, " ")), 1))
        Dim nm41 As Integer = Len(Strings.Left(nm3, InStr(nm3, " ")))
        Dim nm5 As String = Strings.UCase(Strings.Left(Strings.Right(nm3, nm31 - nm41), 1))


        ФИОКор = РуковИП & nm1 & " " & nm4 & "." & nm5 & "."
        УНП = ds8.Rows(0).Item(1).ToString
        КонтТелефон = ds8.Rows(0).Item(12).ToString
        ЮрАдрес = ds8.Rows(0).Item(2).ToString
        РасСчет = ds8.Rows(0).Item(7).ToString
        Банк = ds8.Rows(0).Item(3).ToString
        БИК = ds8.Rows(0).Item(4).ToString
        АдресБанка = ds8.Rows(0).Item(5).ToString
        ЭлАдрес = ds8.Rows(0).Item(13).ToString

        '______

        'сокращенное название и сборное клиента
        'Dim StrSql9 As String = "Select Сокращенное From ФормаСобств Where ПолноеНазвание = '" & ds8.Rows(0).Item(0).ToString & "'"
        'Dim c9 As New OleDbCommand With {
        '        .Connection = conn,
        '        .CommandText = StrSql9
        '    }
        'Dim ds9 As New DataSet
        'Dim da9 As New OleDbDataAdapter(c9)
        'da9.Fill(ds9, "Ст")

        Dim ds9 As DataRow() = dtformft.Select("ПолноеНазвание = '" & ds8.Rows(0).Item(0).ToString & "'")

        ФормаСобствКор = ds9(0).Item("Сокращенное").ToString
        СборноеРеквПолн = ФормаСобствКор & " """ & ComboBox1.Text & """ " & ds8.Rows(0).Item(2).ToString & " IBAN " _
        & ds8.Rows(0).Item(7).ToString & " в " & ds8.Rows(0).Item(3).ToString & " " _
        & ds8.Rows(0).Item(5).ToString & " " & ds8.Rows(0).Item(6).ToString & " БИК " _
        & ds8.Rows(0).Item(4).ToString & " УНП " & ds8.Rows(0).Item(1).ToString

        'KillProc()

        Dim i As Integer



        Dim dfc As DataRow() = dtKartochkaSotrudnikaAll.Select("IDСотр= " & Mass & "")

        '            Dim StrSql2 As String = "SELECT ПродлКонтрС, ПродлКонтрПо FROM КарточкаСотрудника 
        'WHERE IDСотр= " & Mass(i) & ""
        'Dim dfc As DataTable = Selects(StrSql2)




        Dim ds As DataTable = Selects(StrSql:="SELECT Штатное.Должность, Сотрудники.Фамилия, Сотрудники.Имя, Сотрудники.Отчество, 
КарточкаСотрудника.ДатаУведомлПродКонтр, КарточкаСотрудника.НомерУведомлПродКонтр, ДогСотрудн.Контракт, 
ДогСотрудн.ДатаКонтракта, ДогСотрудн.СрокОкончКонтр, КарточкаСотрудника.СрокПродлКонтракта, КарточкаСотрудника.ПродлКонтрС, 
КарточкаСотрудника.ПродлКонтрПо, Сотрудники.НазвОрганиз, Сотрудники.ФИОСборное, КарточкаСотрудника.НеПродлениеКонтр, Штатное.Разряд
FROM((Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр)
INNER JOIN Штатное On Сотрудники.КодСотрудники = Штатное.ИДСотр) INNER JOIN ДогСотрудн On Сотрудники.КодСотрудники = ДогСотрудн.IDСотр
WHERE Сотрудники.НазвОрганиз =@НазвОрганиз AND Сотрудники.КодСотрудники =@ID", list)


        Dim Долж As String
        Dim iЛет, НомерУведо As Integer

        Try
            НомерУведо = CType(ds.Rows(0).Item(5).ToString, Integer)
            Долж = ds.Rows(0).Item(0).ToString
        Catch ex As Exception
            MessageBox.Show("Нет должности.Проверьте данные!", Рик)
            Exit Sub
        End Try


        Dim ДолжРодПадеж As String = ДолжРодПадежФункц(Долж)
        КорИмя = Mid(ds.Rows(0).Item(2).ToString, 1, 1) & "."
        Коротч = Mid(ds.Rows(0).Item(3).ToString, 1, 1) & "."
        Dim СрокПродлКонтр As String = ds.Rows(0).Item(9).ToString
        Dim СрокПродлКонтрПроп As String = ЧислПроп(СрокПродлКонтр)
        Dim СклонВремя As String = Склонение2(СрокПродлКонтр)
        Dim ДатаУвед As Date = ds.Rows(0).Item(4)

        If ds.Rows(0).Item(9).ToString <> "" Then
            iЛет = Val(ds.Rows(0).Item(9))
        End If
        ДатаУвед = ДатаУвед.AddYears(-iЛет)

        Dim ДатаОтвета As String = ДатаУвед.AddDays(15)
        Dim НепродлКон As Boolean = ds.Rows(0).Item(14)

        Dim разряд As Integer = Val(ds.Rows(0).Item(15))
        Dim разрядпроп As String = ""
        If разряд > 0 Then
            разрядпроп = разрядстрока(разряд)
        End If

        If ds.Rows(0).Item(14) = True Or ds.Rows(0).Item(14) = "True" Then 'оформление уведомления о не продлении (уведомление)

            Dim inp As String = InputBox("Введите фамилию сотрудника " & ds.Rows(0).Item(1).ToString & " в Дательном падеже 'Кому?, Чему?'", Рик)

            Do Until inp <> ""
                MessageBox.Show("Повторите ввод фамилии!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Error)
                inp = InputBox("Введите фамилию сотрудника " & ds.Rows(0).Item(1).ToString & " в Дательном падеже 'Кому?, Чему?'", Рик)

            Loop
            УведУвольнение(ДолжРодПадеж, разрядпроп, inp, ds, dfc, Mass)  'доки  УведУвольнение 
            prprov1 = True
        Else
            УведПродлКонтр(ДолжРодПадеж, ds, dfc, Mass, ДатаУвед, СрокПродлКонтрПроп, СклонВремя, ДатаОтвета)
            ДопПродлКонтр(ds, Mass, СклонВремя, ДатаОтвета)   'доки  ДопПродлКонтр
            prprov2 = True
        End If

        'Parallel.Invoke(Sub() УведПродлКонтр(ДолжРодПадеж, ds, dfc, Mass, ДатаУвед, СрокПродлКонтрПроп, СклонВремя, ДатаОтвета))  'доки  УведПродлКонтр


        Dim massFTP As New ArrayList()

        'sw.Stop()
        'MessageBox.Show((sw.ElapsedMilliseconds / 100.0).ToString())

        If MessageBox.Show("Документы оформлены! Распечатать?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            If prprov2 = True Then
                'ПечатьДоков3(PrintУвед, PrintКонтр, PrintКонтр)
                massFTP.Add(УведомлFTP)
                massFTP.Add(ДопСоглFTP)
                massFTP.Add(ДопСоглFTP)
                prprov2 = False
            End If

            If prprov1 = True Then
                'ПечатьДоков(PrintУвол)
                massFTP.Add(УволFTP)
                prprov1 = False
            End If

            ПечатьДоковFTP(massFTP)
        End If


        Me.Cursor = Cursors.Default
    End Sub

    Private Function ДатПродлКонтр(ByVal i As Integer) As Object
        Dim f As Integer

        Dim StrSql2 As String = "SELECT * FROM ПродлКонтракта WHERE IDСотр=" & i & ""
        Dim ds2 As DataTable = Selects(StrSql2)

        Dim StrSql As String = "SELECT СрокКонтракта, ПервоеПродлениеСрок, ВтороеПродлениеСрок, ТретьеПродлениеСрок, ЧетвертоеПродлениеСрок
FROM ПродлКонтракта WHERE IDСотр=" & i & ""
        Dim ds As DataTable = Selects(StrSql)

        If errds = 1 Then
            Dim strsql1 As String = "SELECT * FROM ДогСотрудн WHERE IDСотр=" & i & ""
            Dim ds1 As DataTable = Selects(strsql1)
            Return {ds1.Rows(0).Item(3).ToString, ds1.Rows(0).Item(4).ToString}
        Else
            If ds.Rows(0).Item(0).ToString <> "" Then
                f = 1
                If ds.Rows(0).Item(1).ToString <> "" Then
                    f = 2
                    If ds.Rows(0).Item(2).ToString <> "" Then
                        f = 3
                        If ds.Rows(0).Item(3).ToString <> "" Then
                            f = 4
                            If ds.Rows(0).Item(4).ToString <> "" Then
                                f = 5
                            End If
                        End If
                    End If
                End If
            End If

            Select Case f
                Case 1
                    Return {ds2.Rows(0).Item(3).ToString(), ds2.Rows(0).Item(4).ToString()}
                Case 2
                    Return {ds2.Rows(0).Item(7).ToString(), ds2.Rows(0).Item(8).ToString()}
                Case 3
                    Return {ds2.Rows(0).Item(11).ToString(), ds2.Rows(0).Item(12).ToString()}
                Case 4
                    Return {ds2.Rows(0).Item(15).ToString(), ds2.Rows(0).Item(16).ToString()}
                Case 5
                    Return {ds2.Rows(0).Item(19).ToString(), ds2.Rows(0).Item(20).ToString()}
            End Select
        End If
        Return {"ошибка"}
    End Function
    Private Function Проверка()
        If MaskedTextBox1.MaskCompleted = False Or MaskedTextBox2.MaskCompleted = False Or MaskedTextBox3.MaskCompleted = False Then
            MessageBox.Show("Выберите объект для изменения!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
            Return 1
        End If


        If TextBox7.Text = "" And TextBox8.Text = "" And ComboBox2.Text = "" Then ' нет действий
            MessageBox.Show("Нечего сохранять!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
            Return 1
        End If

        If ComboBox2.Text = "Да" And (TextBox8.Text = "" Or TextBox7.Text = "") Then ' нет действий
            MessageBox.Show("Заполните все поля!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
            Return 1
        End If
        If ComboBox2.Text = "Нет" And TextBox7.Text = "" Then ' нет действий
            MessageBox.Show("Заполните все поля!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
            Return 1
        End If

        If ComboBox2.Text = "Нет" And TextBox7.Text <> "" Then ' нет действий
            TextBox8.Text = ""
        End If



        If ComboBox1.Text = "" Then ' не  выбрана организация
            MessageBox.Show("Выберите организацию!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
            Return 1
        End If

        If CheckBox2.Checked = True Then 'предупреждение о не формировании документов
            If MessageBox.Show("Документы не будут сформированы!", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1) = DialogResult.Cancel Then
                Return 1
            End If
        End If

        Dim период, номерс As String
        'Dim Dfi As Integer = cty
        Dim j
        Try
            j = CDate(MaskedTextBox2.Text) - CDate(MaskedTextBox1.Text)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return 1
        End Try


        Dim f As Integer = Math.Round(CType(j.Days, Integer) / 363)

        Dim n As Integer = 5 - f - CType(TextBox8.Text, Integer)
        If n < 0 Then
            MessageBox.Show("Срок контракта не может быть более 5 лет!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
            Return 1
        End If

        If Now.Date < CDate(MaskedTextBox3.Text) Then
            MessageBox.Show("Продление контракта с сотрудником " & Grid1.CurrentRow.Cells(2).Value & " не наступило!", Рик)
            Return 1
        End If



        Return 0
    End Function
    Private Sub Новое()
        If Проверка() = 1 Then
            Exit Sub
        End If

        Dim r = From x In dtProdlenieKontraktaAll Where x.Item("IDСотр") = CType(Label12.Text, Integer) Select x

        If r.Count > 0 Then
            Dim fd1, fd2, fd3, fd4 As Integer
            Dim f As Integer
            Dim fd As Integer = CType(r(0).Item("СрокКонтракта"), Integer)
            If r(0).Item("ПервоеПродлениеСрок").ToString = "" Then
                f = fd
            ElseIf r(0).Item("ВтороеПродлениеСрок").ToString = "" Then
                fd1 = CType(r(0).Item("ПервоеПродлениеСрок"), Integer)
                f = fd1 + fd
            ElseIf r(0).Item("ТретьеПродлениеСрок").ToString = "" Then
                fd2 = CType(r(0).Item("ВтороеПродлениеСрок"), Integer)
                f = fd1 + fd + fd2
            ElseIf r(0).Item("ЧетвертоеПродлениеСрок").ToString = "" Then
                fd3 = CType(r(0).Item("ТретьеПродлениеСрок"), Integer)
                f = fd1 + fd + fd2 + fd3
            Else
                fd4 = CType(r(0).Item("ЧетвертоеПродлениеСрок"), Integer)
                f = fd1 + fd + fd2 + fd3 + fd4
            End If
            If f > 5 Then
                MessageBox.Show("Период работы с сотрудником " & r(0).Item("ФИО") & "не может превышать более 5 лет по одному контракту." & vbCrLf & "Необходимо оформить новый контракт или изменить срок продления контракта!", Рик)
                Exit Sub
            End If
        End If


        Dim УведомлПродл As Date = CDate(MaskedTextBox3.Text)
        УведомлПродл = УведомлПродл.AddYears(CType(TextBox8.Text, Integer))        'дата уведомления о продлении

        Dim dt1, dt2 As Date
        Dim ColProd As New List(Of String)

        If r(0).Item("ПервоеПродлениеС").ToString = "" Then

            dt2 = CDate(r(0).Item("ДатаОкончания"))
            ColProd.AddRange(New String() {"ПервоеПродлениеС", "ПервоеПродлениеПо", "ПервоеПродлениеСрок", "НомерУвед1"})
            ПровДанн = 1
        ElseIf r(0).Item("ВтороеПродлениеС").ToString = "" Then
            ПровДанн = 2
            dt2 = CDate(r(0).Item("ПервоеПродлениеПо"))
            ColProd.AddRange(New String() {"ВтороеПродлениеС", "ВтороеПродлениеПо", "ВтороеПродлениеСрок", "НомерУвед2"})

        ElseIf r(0).Item("ТретьеПродлениеС").ToString = "" Then
            ПровДанн = 3
            ColProd.AddRange(New String() {"ТретьеПродлениеС", "ТретьеПродлениеПо", "ТретьеПродлениеСрок", "НомерУвед3"})
            dt2 = CDate(r(0).Item("ВтороеПродлениеПо"))

        Else
            ПровДанн = 4
            ColProd.AddRange(New String() {"ЧетвертоеПродлениеС", "ЧетвертоеПродлениеПо", "ЧетвертоеПродлениеСрок", "НомерУвед4"})
            dt2 = CDate(r(0).Item("ТретьеПродлениеПо"))

        End If

        dt1 = dt2.AddDays(1) 'дата по какое действ контракт
        dt2 = dt2.AddYears(1) 'дата с какого действ контракт


        Dim m As String
        If ComboBox2.Text = "Да" Or ComboBox2.Text = "" Then
            m = "False"
        Else
            m = "True"
        End If

        Dim lb As String = Strings.Left((dt2.ToShortDateString), 10)
        Dim list As New Dictionary(Of String, Object)()        '
        list.Add("@УведомлПродл", УведомлПродл)
        list.Add("@dt1", dt1)
        list.Add("@dt2", dt2)
        list.Add("@ID", CType(Label12.Text, Integer))
        list.Add("@dt3", lb)
        list.Add("@dt4", m)

        If Not ComboBox2.Text = "Нет" Then
            Updates(stroka:="UPDATE КарточкаСотрудника
            SET КарточкаСотрудника.НомерУведомлПродКонтр = '" & TextBox7.Text & "',КарточкаСотрудника.СрокПродлКонтракта='" & TextBox8.Text & "',
            КарточкаСотрудника.НеПродлениеКонтр=@dt4, КарточкаСотрудника.ДатаУведомлПродКонтр=@УведомлПродл,
            КарточкаСотрудника.ПродлКонтрС=@dt1, КарточкаСотрудника.ПродлКонтрПо=@dt2
                                Where КарточкаСотрудника.IDСотр =@ID", list)

            Updates(stroka:="UPDATE ДогСотрудн SET СрокОкончКонтр=@dt3 Where IDСотр =@ID", list)


            Updates(stroka:="Update ПродлКонтракта
Set " & ColProd(0) & " = '" & dt1.ToShortDateString & "', " & ColProd(1) & " = '" & dt2.ToShortDateString & "', " & ColProd(2) & "= '" & TextBox8.Text & "', " & ColProd(3) & "= '" & TextBox7.Text & "'
Where IDСотр =@ID", list)

            Статистика1(TextBox3.Text, "Уведомление о продлении контракта", ComboBox1.Text)
        Else

            Updates(stroka:="UPDATE КарточкаСотрудника  Set КарточкаСотрудника.НомерУведомлПродКонтр = '" & TextBox7.Text & "',
КарточкаСотрудника.НеПродлениеКонтр=" & m & "
            Where КарточкаСотрудника.IDСотр=@ID", list)
        End If

        RunMoving6()

        If CheckBox2.Checked = False Then ' Документы продление контракта
            Доки(CType(Label12.Text, Integer))
        End If

        MessageBox.Show("Все данные изменены!", Рик)
        Parallel.Invoke(Sub() ClearAsync())
        refreshgrid()

        dtKartochkaSotrudnika()
        dtProdlenieKontrakta()
        dtDogovorSotrudnik()
    End Sub
    Private Sub ClearAsync()
        TextBox3.Text = ""
        TextBox7.Text = ""
        TextBox10.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        MaskedTextBox1.Text = ""
        MaskedTextBox2.Text = ""
        MaskedTextBox3.Text = ""
        ComboBox2.Text = ""
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click

        Новое()
        Exit Sub



        Dim dg, fm As DataTable
        If (Grid1.Rows.Count - 1) = -1 Then ' нет действий
            MessageBox.Show("Нечего сохранять!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)

            Exit Sub
        End If

        If ComboBox1.Text = "" Then ' не  выбрана организация
            MessageBox.Show("Выберите организацию!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
            Exit Sub
        End If

        If CheckBox2.Checked = True Then 'предупреждение о не формировании документов
            If MessageBox.Show("Документы не будут сформированы!", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1) = DialogResult.Cancel Then
                Exit Sub
            End If
        End If


        Dim MosiFF(Grid1.Columns.Count - 1, Grid1.Rows.Count - 1)
        Dim Str As String = ""

        For Row As Integer = 0 To Grid1.Rows.Count - 1
            For Col As Integer = 0 To Grid1.Columns.Count - 1
                MosiFF(Col, Row) = Grid1.Item(Col, Row).Value
                'Str &= MosiFF(Col, Row) & " "
            Next
            'Str &= vbCrLf
        Next

        If MessageBox.Show("Сохранить данные и сформировать пакет документов?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = DialogResult.Cancel Then

            Exit Sub

        End If
        Dim MassID(Grid1.Rows.Count - 1) As Integer
        'Dim bool As Boolean
        Dim i As Integer = 0
        Dim StrSql As String 'сохранение в базу

        For i = 0 To Grid1.Rows.Count - 1 'LBound(MosiFF) To UBound(MosiFF)

            Dim период, номерс As String
            номерс = MosiFF(6, i).ToString
            период = MosiFF(7, i).ToString
            Dim период2 As Integer = Val(период)

            If период2 > 5 Then
                MessageBox.Show("Срок контракта не может быть более 5 лет!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1)
                Exit Sub
            End If


            If период = "" And номерс = "" And MosiFF(8, i) = False Then ' если три столбца пустых
                '                StrSql = "UPDATE КарточкаСотрудника  SET КарточкаСотрудника.НомерУведомлПродКонтр= '" & номерс & "'
                ',КарточкаСотрудника.СрокПродлКонтракта='" & период & "', КарточкаСотрудника.НеПродлениеКонтр=" & MosiFF(8, i) & "
                '            Where КарточкаСотрудника.IDСотр = " & MosiFF(0, i) & ""
                '                Dim c As New OleDbCommand
                '                c.Connection = conn
                '                c.CommandText = StrSql
                '                c.ExecuteNonQuery()

                Continue For

            End If


            Dim ДатаСтрока As String = Format(Now, "yyyy") 'сравниваем года
            Dim ПровПродКонтр1 As String = Strings.Left(MosiFF(5, i), 10)
            ПровПродКонтр = Strings.Right(ПровПродКонтр1, 4)
            Dim Пр As Integer = Val(ПровПродКонтр)
            Dim Сег As Integer = Val(ДатаСтрока)

            If Пр > Сег Then 'проверка - можно ли продлевать контракт
                MessageBox.Show("Продление контракта с сотрудником " & Grid1.CurrentRow.Cells(2).Value & " не наступило!", Рик)
                Exit Sub

            End If


            If период <> "" And номерс <> "" And MosiFF(8, i) = True Then ' если три столбца заполнены значениями
                MsgBox("Сделайте правильный выбор!" & vbCrLf & "Продление контракта или Не продление продление для сотрудника " & vbCrLf & MosiFF(2, i),, Рик)
                Exit Sub
            End If


            errds = 0
            StrSql = ""
            StrSql = "SELECT СрокКонтракта, ПервоеПродлениеСрок, ВтороеПродлениеСрок, ТретьеПродлениеСрок, ЧетвертоеПродлениеСрок
FROM ПродлКонтракта WHERE IDСотр=" & MosiFF(0, i) & ""
            Dim datts As DataTable = Selects(StrSql)

            Dim elet(datts.Columns.Count - 1) As Integer
            If errds = 0 Then
                For gt As Integer = 0 To datts.Columns.Count - 1
                    If datts.Rows(0).Item(gt).ToString = "" Then
                        datts.Rows(0).Item(gt) = 0
                    End If
                    elet(gt) = CType(datts.Rows(0).Item(gt), Integer)
                Next
            End If
            Dim m As Integer = elet(0) + elet(1) + elet(2) + elet(3) + elet(4) + CType(MosiFF(7, i), Integer)
            If m > 5 Then
                Dim gv As Integer = elet(0) + elet(1) + elet(2) + elet(3) + elet(4)
                Dim gv1 As String

                If gv = 5 Then
                    gv1 = " лет"
                Else
                    gv1 = " года"
                End If
                MessageBox.Show("Период работы с сотрудником " & vbCrLf & MosiFF(2, i) & vbCrLf & "не может превышать более 5 лет по одному контракту." & vbCrLf & "Необходимо оформить новый контракт или изменить срок продления контракта!" _
                                & vbCrLf & "На данный момент сотрудник проработал " & elet(0) + elet(1) + elet(2) + elet(3) & " " & gv1, Рик)
                'If Grid1.Rows.Count - 1 = 0 Then
                '    refreshgrid()
                '    Me.Cursor = Cursors.Default
                '    Exit Sub
                'End If
                Continue For
            End If

            Select Case MosiFF(8, i)'выбор продление или непродление контракта
                Case False

                    If период = "" Or номерс = "" Then
                        MsgBox("Заполните срок продления контракта или номер уведомления " & MosiFF(2, i),, Рик)
                        Exit Sub
                    End If

                    Dim iDat, dateWith As Date

                    dateWith = MosiFF(5, i) 'дата уведомл о продлении контракта старая
                    iDat = MosiFF(5, i)

                    iDat = iDat.AddYears(MosiFF(7, i))
                    MosiFF(5, i) = iDat 'дата уведомления о продлении


                    Dim obj As Object = ДатПродлКонтр(MosiFF(0, i))

                    Dim DtFierst, DtFinish As Date

                    'DtFierst = CDate(obj(0).ToString)
                    DtFinish = CDate(obj(1).ToString)
                    'DtFierst = DtFierst.AddYears(CType(MosiFF(7, i), Integer))
                    DtFierst = DtFinish.AddDays(1)
                    DtFinish = DtFinish.AddYears(CType(MosiFF(7, i), Integer))
                    MosiFF(9, i) = DtFierst
                    MosiFF(10, i) = DtFinish
                    'dateWith = dateWith.AddMonths(1)
                    'dateWith = dateWith.AddDays(+2) 'дата продления с
                    'MosiFF(9, i) = dateWith

                    'iDat = iDat.AddMonths(1)
                    'iDat = iDat.AddDays(1)
                    'MosiFF(10, i) = iDat

                    StrSql = "UPDATE КарточкаСотрудника  SET КарточкаСотрудника.НомерУведомлПродКонтр = '" & MosiFF(6, i) & "',
КарточкаСотрудника.СрокПродлКонтракта='" & MosiFF(7, i) & "', КарточкаСотрудника.НеПродлениеКонтр=" & MosiFF(8, i) & ",
КарточкаСотрудника.ДатаУведомлПродКонтр='" & MosiFF(5, i) & "', КарточкаСотрудника.ПродлКонтрС='" & MosiFF(9, i) & "',
КарточкаСотрудника.ПродлКонтрПо='" & MosiFF(10, i) & "'
            Where КарточкаСотрудника.IDСотр = " & MosiFF(0, i) & ""
                    Dim c As New OleDbCommand
                    c.Connection = conn
                    c.CommandText = StrSql
                    c.ExecuteNonQuery()

                    'StrSql = ""
                    StrSql = "UPDATE ДогСотрудн SET СрокОкончКонтр='" & MosiFF(10, i) & "' Where IDСотр = " & MosiFF(0, i) & ""
                    Updates(StrSql)

                    MassID(i) = MosiFF(0, i)



                    errds = 0 'начало блока добавления в таблицу ПродлКонтракта
                    StrSql = ""
                    StrSql = "SELECT НомерУвед,НомерУвед1,НомерУвед2,НомерУвед3,НомерУвед4
FROM Сотрудники INNER JOIN ПродлКонтракта ON Сотрудники.КодСотрудники = ПродлКонтракта.IDСотр
Where Сотрудники.КодСотрудники = " & MosiFF(0, i) & ""
                    dg = Selects(StrSql)

                    If errds = 0 Then

                        If dg.Rows(0).Item(0).ToString <> "" Then
                            If dg.Rows(0).Item(1).ToString <> "" Then
                                If dg.Rows(0).Item(2).ToString <> "" Then
                                    If dg.Rows(0).Item(3).ToString <> "" Then
                                        ПровДанн = 4
                                    Else
                                        ПровДанн = 3
                                    End If
                                Else
                                    ПровДанн = 2
                                End If
                            Else
                                ПровДанн = 1
                            End If
                        Else
                            StrSql = "" 'Втсавляем данные впервый раз(если произошла ошибка)
                            StrSql = "SELECT КарточкаСотрудника.СрокКонтракта, ДогСотрудн.СрокОкончКонтр
FROM КарточкаСотрудника, ДогСотрудн
WHERE КарточкаСотрудника.IDСотр=" & MosiFF(0, i) & " AND ДогСотрудн.IDСотр=" & MosiFF(0, i) & ""
                            fm = Selects(StrSql)

                            StrSql = ""
                            StrSql = "INSERT INTO ПродлКонтракта(IDСотр,ФИО,ДатаПриема,ДатаОкончания,СрокКонтракта,НомерУвед)
VALUES(" & MosiFF(0, i) & ",'" & MosiFF(2, i) & "','" & MosiFF(3, i) & "','" & fm.Rows(0).Item(1).ToString & "','" & fm.Rows(0).Item(0).ToString & "','1' )"
                            ПровДанн = 1
                        End If
                    Else

                        StrSql = "" 'Втсавляем данные впервый раз(если нет ID)
                        StrSql = "SELECT КарточкаСотрудника.СрокКонтракта, ДогСотрудн.СрокОкончКонтр
                        From КарточкаСотрудника, ДогСотрудн
WHERE КарточкаСотрудника.IDСотр=" & MosiFF(0, i) & " AND ДогСотрудн.IDСотр=" & MosiFF(0, i) & ""
                        fm = Selects(StrSql)

                        StrSql = ""
                        StrSql = "INSERT INTO ПродлКонтракта(IDСотр,ФИО,ДатаПриема,ДатаОкончания,СрокКонтракта,НомерУвед)
VALUES(" & MosiFF(0, i) & ",'" & MosiFF(2, i) & "','" & MosiFF(3, i) & "','" & fm.Rows(0).Item(1).ToString & "','" & fm.Rows(0).Item(0).ToString & "','1' )"
                        Updates(StrSql)
                        ПровДанн = 1
                    End If


                    Select Case ПровДанн
                        Case 1
                            ДаннПродлКонт(1, MosiFF(9, i), MosiFF(10, i), MosiFF(7, i), MosiFF(0, i), MosiFF(6, i))
                        Case 2
                            ДаннПродлКонт(2, MosiFF(9, i), MosiFF(10, i), MosiFF(7, i), MosiFF(0, i), MosiFF(6, i))
                        Case 3
                            ДаннПродлКонт(3, MosiFF(9, i), MosiFF(10, i), MosiFF(7, i), MosiFF(0, i), MosiFF(6, i))
                        Case 4
                            ДаннПродлКонт(4, MosiFF(9, i), MosiFF(10, i), MosiFF(7, i), MosiFF(0, i), MosiFF(6, i))
                    End Select

                Case True
                    If номерс = "" Then
                        MsgBox("Заполните номер уведомления " & MosiFF(2, i),, Рик)
                        Exit Sub
                    End If



                    StrSql = "UPDATE КарточкаСотрудника  Set КарточкаСотрудника.НомерУведомлПродКонтр = '" & MosiFF(6, i) & "',
КарточкаСотрудника.НеПродлениеКонтр=" & MosiFF(8, i) & "
            Where КарточкаСотрудника.IDСотр = " & MosiFF(0, i) & ""
                    Dim c As New OleDbCommand
                    c.Connection = conn
                    c.CommandText = StrSql
                    c.ExecuteNonQuery()

                    MassID(i) = MosiFF(0, i)


            End Select
            Статистика1(MosiFF(2, i), "Уведомление о продлении контракта", ComboBox1.Text)

        Next






        If CheckBox2.Checked = False Then ' Документы продление контракта

            'Доки(MassID)
            '___________________________
        End If

        MessageBox.Show("Все данные изменены!", Рик)
        refreshgrid()

    End Sub
    Private Sub ДаннПродлКонт(ByVal числ As Integer, ByVal fr As String, ByVal До As String, ByVal Срок As String, ByVal id As Integer, ByVal НомУвед As String)
        Dim StrSql As String

        If числ = 1 Then
            StrSql = ""
            StrSql = "Update ПродлКонтракта Set ПервоеПродлениеС = '" & fr & "', ПервоеПродлениеПо = '" & До & "',
ПервоеПродлениеСрок= '" & Срок & "', НомерУвед1= '" & НомУвед & "'  Where IDСотр = " & id & ""
            Updates(StrSql)
            Exit Sub
        End If

        If числ = 2 Then
            StrSql = ""
            StrSql = "Update ПродлКонтракта Set ВтороеПродлениеС = '" & fr & "', ВтороеПродлениеПо = '" & До & "',
ВтороеПродлениеСрок= '" & Срок & "', НомерУвед2= '" & НомУвед & "'  Where IDСотр = " & id & ""
            Updates(StrSql)
            Exit Sub
        End If

        If числ = 3 Then
            StrSql = ""
            StrSql = "Update ПродлКонтракта Set ТретьеПродлениеС = '" & fr & "', ТретьеПродлениеПо = '" & До & "',
ТретьеПродлениеСрок= '" & Срок & "', НомерУвед3= '" & НомУвед & "'  Where IDСотр = " & id & ""
            Updates(StrSql)
            Exit Sub
        End If

        If числ = 4 Then
            StrSql = ""
            StrSql = "Update ПродлКонтракта Set ЧетвертоеПродлениеС = '" & fr & "', ЧетвертоеПродлениеПо = '" & До & "',
ЧетвертоеПродлениеСрок= '" & Срок & "', НомерУвед4= '" & НомУвед & "'  Where IDСотр = " & id & ""
            Updates(StrSql)
            Exit Sub
        End If


    End Sub
    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        'refreshgrid()
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        'refreshgrid()
    End Sub

    Private Sub Grid1_ColumnSortModeChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles Grid1.ColumnSortModeChanged

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs)

        'Select Case CheckBox1.Checked
        '    Case True
        '        ComboBox3.Enabled = False
        '    Case False
        '        ComboBox3.Enabled = True
        'End Select

    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs)
        refreshgrid()

    End Sub
    Private Sub ВсплывФорма()

        Dim hy As Integer = Grid1.CurrentRow.Cells("ID").Value
        name2 = Grid1.CurrentRow.Cells("ФИО Сотрудника").Value.ToString

        errds = 0
        Dim strsql As String
        strsql = "Select * From ПродлКонтракта WHERE IDСотр=" & hy & ""
        Dim dt As DataTable = Selects(strsql)
        If errds = 1 Then
            MessageBox.Show("Данные по сотруднику не внесены в таблицу!", Рик)

        ElseIf errds = 0 Then

            With УведомлениеФорма
                .TextBox1.Text = name2
                .TextBox2.Text = dt.Rows(0).Item(3).ToString
                .TextBox3.Text = dt.Rows(0).Item(4).ToString
                .TextBox4.Text = dt.Rows(0).Item(5).ToString
                .TextBox5.Text = dt.Rows(0).Item(9).ToString
                .TextBox6.Text = dt.Rows(0).Item(8).ToString
                .TextBox7.Text = dt.Rows(0).Item(7).ToString
                .TextBox8.Text = dt.Rows(0).Item(10).ToString
                .TextBox9.Text = dt.Rows(0).Item(14).ToString
                .TextBox10.Text = dt.Rows(0).Item(13).ToString
                .TextBox11.Text = dt.Rows(0).Item(12).ToString
                .TextBox12.Text = dt.Rows(0).Item(11).ToString
                .TextBox13.Text = dt.Rows(0).Item(18).ToString
                .TextBox14.Text = dt.Rows(0).Item(17).ToString
                .TextBox15.Text = dt.Rows(0).Item(16).ToString
                .TextBox16.Text = dt.Rows(0).Item(15).ToString
                .TextBox17.Text = dt.Rows(0).Item(22).ToString
                .TextBox18.Text = dt.Rows(0).Item(21).ToString
                .TextBox19.Text = dt.Rows(0).Item(20).ToString
                .TextBox20.Text = dt.Rows(0).Item(19).ToString
                .TextBox21.Text = dt.Rows(0).Item(6).ToString
            End With
            УведомлениеФорма.ShowDialog()
        End If
    End Sub



    Private Sub Grid1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellContentDoubleClick
        'TextBox2.Text = ""
        'If Grid1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value <> "" Then
        '    TextBox2.Text = Grid1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString()
        'End If
        'Dim index As Integer
        'index = e.RowIndex
        'Dim selectedRow As DataGridViewRow
        'selectedRow = Grid1.Rows(index)
        'TextBox2.Text = selectedRow.Cells(1).Value.ToString

        'ВсплывФорма()







    End Sub

    'Вставка из datagrid в эксель
    ''''''Private Sub tsbtnCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbtnCopy.Click
    ''''''    dgv01.SuspendLayout()
    ''''''    dgv01.RowHeadersVisible = False
    ''''''    If dgv01.SelectedRows.Count = 0 Then dgv01.SelectAll()
    ''''''    Clipboard.SetDataObject(dgv01.GetClipboardContent())
    ''''''    dgv01.ClearSelection()
    ''''''    dgv01.RowHeadersVisible = True
    ''''''    dgv01.ResumeLayout()
    ''''''End Sub
    '''
    Private Sub idСотрудника()
        Dim StrSql As String
        '        StrSql = "SELECT Сотрудники.КодСотрудники, Сотрудники.ФИОСборное, КарточкаСотрудника.ДатаУведомлПродКонтр, 
        'КарточкаСотрудника.ПродлКонтрС, ДогСотрудн.СрокОкончКонтр, КарточкаСотрудника.ПродлКонтрПо
        'FROM (Сотрудники INNER JOIN КарточкаСотрудника ON Сотрудники.КодСотрудники = КарточкаСотрудника.IDСотр) INNER JOIN ДогСотрудн ON Сотрудники.КодСотрудники = ДогСотрудн.IDСотр
        'WHERE ФИОСборное='" & ComboBox2.Text & "'"

        '        Dim c2 As New OleDbCommand
        '        c2.Connection = conn
        '        c2.CommandText = StrSql
        '        Dim ds2 As New DataSet
        '        Dim da2 As New OleDbDataAdapter(c2)
        '        da2.Fill(ds2, "ИД")
        '        idClient = ds2.Tables("ИД").Rows(0).Item(0)
        '        ПровПродКонтр = ds2.Tables("ИД").Rows(0).Item(2).ToString
        '        ПродлКонтрС = ds2.Tables("ИД").Rows(0).Item(3).ToString
        '        СрокОкончКонтр = ds2.Tables("ИД").Rows(0).Item(4).ToString
        '        ПродлКонтрПо = ds2.Tables("ИД").Rows(0).Item(5).ToString
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged

        refreshgrid()

    End Sub

    Private Sub Grid1_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles Grid1.ColumnHeaderMouseClick
        Grid1.Columns(6).SortMode = DataGridViewColumnSortMode.NotSortable
        Grid1.Columns(7).SortMode = DataGridViewColumnSortMode.NotSortable
        Grid1.Columns(8).SortMode = DataGridViewColumnSortMode.NotSortable
        Grid1.Columns(9).SortMode = DataGridViewColumnSortMode.NotSortable
        Grid1.Columns(10).SortMode = DataGridViewColumnSortMode.NotSortable
        Grid1.Columns(5).SortMode = DataGridViewColumnSortMode.NotSortable
    End Sub

    Private Sub ComboBox3_SelectedValueChanged(sender As Object, e As EventArgs)
        'Select Case ComboBox3.Text
        '    Case <> ""
        '        Me.CheckBox1.Enabled = False
        '    Case ""
        '        Me.CheckBox1.Enabled = True
        'End Select
    End Sub

    Private Sub Grid1_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles Grid1.CellBeginEdit
        Select Case e.ColumnIndex
            Case 0
                e.Cancel = True
            Case 1
                e.Cancel = True
            Case 2
                e.Cancel = True
            Case 3
                e.Cancel = True
            Case 4
                e.Cancel = True
            Case 5
                e.Cancel = True
            Case 9
                e.Cancel = True
            Case 10
                e.Cancel = True

        End Select

    End Sub

    Private Sub Grid1_ColumnHeaderMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles Grid1.ColumnHeaderMouseDoubleClick
        Grid1.Columns(6).SortMode = DataGridViewColumnSortMode.NotSortable
        Grid1.Columns(7).SortMode = DataGridViewColumnSortMode.NotSortable
        Grid1.Columns(8).SortMode = DataGridViewColumnSortMode.NotSortable
        Grid1.Columns(9).SortMode = DataGridViewColumnSortMode.NotSortable
        Grid1.Columns(10).SortMode = DataGridViewColumnSortMode.NotSortable
        Grid1.Columns(5).SortMode = DataGridViewColumnSortMode.NotSortable

    End Sub

    'Private Sub Grid1_ColumnDisplayIndexChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles Grid1.ColumnDisplayIndexChanged

    'End Sub





    'Private Sub Grid1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellClick
    '    TextBox2.Text = ""
    '    'If Grid1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value <> "" Then
    '    '    TextBox2.Text = Grid1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString()
    '    'End If
    '    Dim index As Integer
    '    index = e.RowIndex
    '    Dim selectedRow As DataGridViewRow
    '    selectedRow = Grid1.Rows(index)
    '    TextBox2.Text = selectedRow.Cells(1).Value.ToString
    'End Sub
End Class