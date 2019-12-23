Imports System.Data.OleDb
Imports System.Threading
Public Class ШтИзмСтавкиВсплКласс
    Public d1 As Integer
    Public f1 As String
    Public ds As DataTable
    Sub New(ByVal d As Integer, ByVal f As String)
        d1 = d
        f1 = f
    End Sub

    Sub таблица()
        Dim strsql As String = "SELECT ШтСводИзмСтавка.Код, ШтОтделы.Клиент, ШтСвод.Должность, ШтСвод.Разряд,
ШтСводИзмСтавка.Дата, ШтСводИзмСтавка.Ставка 
FROM ШтОтделы INNER JOIN (ШтСвод INNER JOIN ШтСводИзмСтавка ON ШтСвод.КодШтСвод = ШтСводИзмСтавка.IDКодШтСвод) ON ШтОтделы.Код = ШтСвод.Отдел
WHERE ШтСводИзмСтавка.IDКодШтСвод=" & d1 & " ORDER BY ШтСводИзмСтавка.Дата"
        ds = Selects(strsql)
        If errds = 1 Then
            Dim strsql1 As String = "SELECT ШтСвод.КодШтСвод, ШтОтделы.Клиент, ШтСвод.Должность, ШтСвод.Разряд, ШтСвод.ТарифнаяСтавка
FROM ШтОтделы INNER JOIN ШтСвод ON ШтОтделы.Код = ШтСвод.Отдел
WHERE ШтСвод.КодШтСвод=" & d1 & ""
            ds = Selects(strsql1)
        End If

    End Sub


End Class
