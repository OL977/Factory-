Option Explicit On
Imports System.Data.OleDb
Imports System.Data.SqlClient

Module ДолжОконч
    Public ДогПодномДогПод As Integer
    Public ДогПодномДогПодНомДог As String
    Public ДогПодномДогПодСтДог As String
    Public ДогПодрВклЧекбокс5 As Boolean
    Public Com19ForДогПодр As String
    'Public dуеrs As Task = New Task(AddressOf Поиск.Работники)
    'Public dуеrs1 As Task = New Task(AddressOf Поиск.Организ)
    Public Function окончание1(ByVal должность As String, ByVal номер As Integer) As String
        Dim пров As Integer
        Dim вопрос, вопрос2 As String
        'conn = New OleDbConnection
        'conn.ConnectionString = ConString
        'Try
        '    conn.Open()
        'Catch ex As Exception
        '    MessageBox.Show("Не подключен диск U")
        'End Try

        Select Case номер
            Case 1
                вопрос = "Кого"

            Case 2
                вопрос = "Кому"

            Case 3
                вопрос = "Кем"

        End Select

        Dim ds5 As DataTable = Selects(StrSql:="SELECT Должность FROM Окончание WHERE Должность='" & должность & "'")
        Dim проверка As String
        If errds = 1 Then
            проверка = ""
        Else
            проверка = ds5.Rows(0).Item(0).ToString
        End If




        Dim ds As DataTable = Selects(StrSql:="SELECT " & вопрос & " FROM Окончание WHERE Должность='" & должность & "'")

        Try
            окончание1 = ds.Rows(0).Item(0).ToString
            If окончание1 <> "" Then
                Return окончание1
            End If


        Catch ex As Exception
            пров = 1
        End Try

        Do
            окончание1 = InputBox("Введите должность " & должность & " в соотвествующем падеже!" & vbCrLf & "Вопрос - " & вопрос & "?", Рик, должность)
        Loop Until окончание1 <> ""

        Dim inr As Integer = окончание1.Length - 1
        окончание1 = StrConv(Strings.Left(окончание1, 1), VbStrConv.ProperCase) & Strings.Right(окончание1, inr)
        Dim conn As SqlConnection
        If проверка = "" Then
            Dim StrSql5 As String = "INSERT INTO Окончание(Должность, " & вопрос & ") VALUES('" & должность & "','" & окончание1 & "')"

            conn = New SqlConnection(ConString)
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            Dim c25 As New SqlCommand(StrSql5, conn)
            Try
                c25.ExecuteNonQuery()
                MessageBox.Show("Данные внесены в базу!", Рик)

            Catch ex As Exception
                MessageBox.Show("Что то пошло не так, добавьте должность через форму!", Рик)
            End Try
        Else
            Dim StrSql7 As String = "UPDATE Окончание SET " & вопрос & "='" & окончание1 & "' WHERE Должность='" & должность & "'"

            conn = New SqlConnection(ConString)
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            Dim c2 As New SqlCommand(StrSql7, conn)

            Try
                c2.ExecuteNonQuery()
                MessageBox.Show("Данные внесены в базу!", Рик)

            Catch ex As Exception
                MessageBox.Show("Что то пошло не так, добавьте должность через форму!", Рик)
            End Try

        End If
        If conn.State = ConnectionState.Open Then
            conn.Close()
        End If
        dtOkonchanie()


        Return окончание1


    End Function
End Module
