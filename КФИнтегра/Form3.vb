
Imports System.Runtime.InteropServices

Public Class Form3
    <StructLayout(LayoutKind.Sequential)>
    Structure LASTINPUTINFO
        <MarshalAs(UnmanagedType.U4)>
        Public cbSize As Integer
        <MarshalAs(UnmanagedType.U4)>
        Public dwTime As Integer
    End Structure
    <DllImport("user32.dll")>
    Shared Function GetLastInputInfo(ByRef plii As LASTINPUTINFO) As Boolean
    End Function

    Dim idletime As Integer
    Dim lastInputInf As New LASTINPUTINFO()
    Public Function GetLastInputTime() As Integer
        idletime = 0
        lastInputInf.cbSize = Marshal.SizeOf(lastInputInf)
        lastInputInf.dwTime = 0

        If GetLastInputInfo(lastInputInf) Then
            idletime = Environment.TickCount - lastInputInf.dwTime
        End If

        If idletime > 0 Then
            Return idletime / 1000
        Else : Return 0
        End If
    End Function

    Private sumofidletime As TimeSpan = New TimeSpan(0)
    Private LastLastIdletime As Integer = 0
    'Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

    'End Sub





    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'ListBox1.Items.AddRange(IO.Directory.GetFileSystemEntries("C: \Users\OLEG\Desktop\РАБОЧАЯ\КЛИЕНТ\Walls"))
        ListBox1.Items.AddRange(IO.Directory.GetFiles("C: \Users\OLEG\Desktop\РАБОЧАЯ\КЛИЕНТ\Walls"))
    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")
        Timer1.Start()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '    SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        '    Dim ef As ExcelFile = New ExcelFile
        '    Dim ws As ExcelWorksheet = ef.Worksheets.Add("Hello World")
        '    ws.Cells(0, 0).Value = "English:"
        '    ws.Cells(0, 1).Value = "Hello"

        '    ws.Cells(1, 0).Value = "Russian:"
        '    ' Using UNICODE string.
        '    ws.Cells(1, 1).Value = New String(New Char() {ChrW(&H417), ChrW(&H434), ChrW(&H440), ChrW(&H430), ChrW(&H432), ChrW(&H441), ChrW(&H442), ChrW(&H432), ChrW(&H443), ChrW(&H439), ChrW(&H442), ChrW(&H435)})

        '    ws.Cells(2, 0).Value = "Chinese:"
        '    ' Using UNICODE string.
        '    ws.Cells(2, 1).Value = New String(New Char() {ChrW(&H4F60), ChrW(&H597D)})

        '    ws.Cells(4, 0).Value = "In order to see Russian and Chinese characters you need to have appropriate fonts on your PC."
        '    ws.Cells.GetSubrangeAbsolute(4, 0, 4, 7).Merged = True

        '    ef.Save("Hello World.xlsx")
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim it As Integer = GetLastInputTime()
        If LastLastIdletime > it Then
            Label1.Text = "Начало простоя!"
            sumofidletime = sumofidletime.Add(TimeSpan.FromSeconds(LastLastIdletime))
            Label2.Text = "Время простоя: " & sumofidletime.ToString
        Else
            Label1.Text = GetLastInputTime()
        End If
        LastLastIdletime = it
    End Sub
End Class