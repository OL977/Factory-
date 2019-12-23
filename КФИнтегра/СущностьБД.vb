Imports System.Data.Linq.Mapping
Module СущностьБД

    <Table(Name:="ПутиДокументов")>
    Public Class ПутиДок
        <Column(IsPrimaryKey:=True, IsDbGenerated:=True)>        '<Column(AutoSync:=True)>
        Public Property Код() As Integer

        <Column>
        Public Property IDСотрудник() As Integer
        <Column>
        Public Property Путь() As String

        <Column>
        Public Property ИмяФайла() As String

        <Column>
        Public Property ДокМесто() As String

        <Column>
        Public Property Предприятие() As String

        <Column>
        Public Property ПолныйПуть() As String
    End Class
    <Table(Name:="Клиент")>
    Public Class Клиент
        <Column(IsPrimaryKey:=True, CanBeNull:=False)>
        Public Property НазвОрг() As String
        <Column(CanBeNull:=True)>
        Public Property ФормаСобств() As String
        <Column(CanBeNull:=True)>
        Public Property УНП() As String

        <Column(CanBeNull:=True)>
        Public Property ФактичАдрес() As String

        <Column(CanBeNull:=True)>
        Public Property ЮрАдрес() As String

        <Column(CanBeNull:=True)>
        Public Property ПочтАдрес() As String

        <Column(CanBeNull:=True)>
        Public Property КонтТелефон() As String

        <Column(CanBeNull:=True)>
        Public Property Факс() As String
        <Column(CanBeNull:=True)>
        Public Property ЭлАдрес() As String
        <Column(CanBeNull:=True)>
        Public Property ДругиеКонтакты() As String
        <Column(CanBeNull:=True)>
        Public Property Банк() As String
        <Column(CanBeNull:=True)>
        Public Property БИКБанка() As String
        <Column(CanBeNull:=True)>
        Public Property АдресБанка() As String
        <Column(CanBeNull:=True)>
        Public Property Отделение() As String
        <Column(CanBeNull:=True)>
        Public Property РасчСчетРубли() As String
        <Column(CanBeNull:=True)>
        Public Property РасчСчетЕвро() As String
        <Column(CanBeNull:=True)>
        Public Property РасчСчетДоллар() As String
        <Column(CanBeNull:=True)>
        Public Property РасчСчетРоссРубли() As String
        <Column(CanBeNull:=True)>
        Public Property ДолжнРуководителя() As String
        <Column(CanBeNull:=True)>
        Public Property ФИОРуководителя() As String
        <Column(CanBeNull:=True)>
        Public Property ОснованиеДейств() As String
        <Column(CanBeNull:=True)>
        Public Property ТелРуков() As String
        <Column(CanBeNull:=True)>
        Public Property ДолжнДопЛица() As String
        <Column(CanBeNull:=True)>
        Public Property ФИОДопЛица() As String
        <Column(CanBeNull:=True)>
        Public Property ТелДопЛица() As String
        <Column(CanBeNull:=True)>
        Public Property Операционист() As String
        <Column(CanBeNull:=True)>
        Public Property КонтТелОпер() As String
        <Column(CanBeNull:=True)>
        Public Property ФизЛицо() As Integer
        <Column(CanBeNull:=True)>
        Public Property ЮрЛицо() As Integer
        <Column(CanBeNull:=True)>
        Public Property ФИОРукРодПадеж() As String
        <Column(CanBeNull:=True)>
        Public Property ФИОРукДатПадеж() As String
        <Column(CanBeNull:=True)>
        Public Property РукИП() As String

    End Class

    <Table(Name:="InputName")>
    Public Class Test
        <Column(IsPrimaryKey:=True, CanBeNull:=False)>
        Public Property Код() As Integer
        <Column(CanBeNull:=True)>
        Public Property ФИООриганл() As String
        <Column(CanBeNull:=True)>
        Public Property ОтпускСоц() As String

        <Column(CanBeNull:=True)>
        Public Property ОтпускТруд() As String
    End Class
    <Table(Name:="Банк")>
    Public Class Банк
        <Column(CanBeNull:=True)>
        Public Property Код() As Double
        <Column(CanBeNull:=True)>
        Public Property БИК() As String
        <Column(CanBeNull:=True)>
        Public Property Наименование() As String

        <Column(IsPrimaryKey:=True, CanBeNull:=False)>
        Public Property КодГлав() As Integer
    End Class

    <Table(Name:="КарточкаСотрудника")> 'Карточка сотрудника
    Public Class КартСотр
        <Column(IsPrimaryKey:=True, CanBeNull:=False)>
        Public Property Код() As Integer
        <Column(CanBeNull:=True)>
        Public Property IDСотр() As Integer
        <Column>
        Public Property ДатаПриема() As Date

        <Column>
        Public Property ДатаУвольнения() As Date

        <Column(CanBeNull:=True)>
        Public Property СрокКонтракта() As String

        <Column(CanBeNull:=True)>
        Public Property ТипРаботы() As String

        <Column(CanBeNull:=True)>
        Public Property Ставка() As String
        <Column(CanBeNull:=True)>
        Public Property ВремяНачРаботы() As String
        <Column(CanBeNull:=True)>
        Public Property ПродолРабДня() As String
        <Column(CanBeNull:=True)>
        Public Property Обед() As String
        <Column(CanBeNull:=True)>
        Public Property ОкончРабДня() As String
        <Column(CanBeNull:=True)>
        Public Property Выходные() As String
        <Column(CanBeNull:=True)>
        Public Property ПриказОбУвольн() As String
        <Column>
        Public Property ДатаПриказаОбУвольн() As Date
        <Column(CanBeNull:=True)>
        Public Property ОснованиеУвольн() As String
        <Column>
        Public Property ДатаУведомлПродКонтр() As Date
        <Column(CanBeNull:=True)>
        Public Property НомерУведомлПродКонтр() As String
        <Column(CanBeNull:=True)>
        Public Property СрокПродлКонтракта() As String
        <Column>
        Public Property ПродлКонтрС() As Date
        <Column>
        Public Property ПродлКонтрПо() As Date
        <Column(CanBeNull:=True)>
        Public Property НеПродлениеКонтр() As String
        <Column(CanBeNull:=True)>
        Public Property СрокОтветаНаУведомл() As String
        <Column(CanBeNull:=True)>
        Public Property АдресОбъектаОбщепита() As String
        <Column(CanBeNull:=True)>
        Public Property ПриказПродлКонтр() As String
        <Column(CanBeNull:=True)>
        Public Property ДатаЗарплаты() As String
        <Column(CanBeNull:=True)>
        Public Property ДатаАванса() As String
        <Column(CanBeNull:=True)>
        Public Property ПоСовмест() As String
        <Column(CanBeNull:=True)>
        Public Property СуммирУчет() As String
        <Column(CanBeNull:=True)>
        Public Property НомУведИзмСрокЗарп() As Integer
        <Column(CanBeNull:=True)>
        Public Property ДатаСогласияНаИзмен() As String
        <Column(CanBeNull:=True)>
        Public Property ДатаВсуплСоглаш() As String
        <Column(CanBeNull:=True)>
        Public Property ДатаУведом() As String
        <Column(CanBeNull:=True)>
        Public Property ДатаПеревода() As String
        <Column(CanBeNull:=True)>
        Public Property ДатаЗаявленияПеревода() As String
        <Column(CanBeNull:=True)>
        Public Property ДатаПриказаПеревода() As String
        <Column(CanBeNull:=True)>
        Public Property НомерПриказаПеревода() As String
        <Column(CanBeNull:=True)>
        Public Property Примечание() As String
        <Column(CanBeNull:=True)>
        Public Property НаличиеИспытСрока() As String
        <Column(CanBeNull:=True)>
        Public Property ПериодОтпДляКонтр() As String
        <Column(CanBeNull:=True)>
        Public Property НомерУведИзмОклада() As Integer
        <Column(CanBeNull:=True)>
        Public Property ДатаУведомИзмОклада() As String

    End Class
    <Table(Name:="Сотрудники")> 'Сотрудники
    Public Class Сотрудники
        <Column(IsPrimaryKey:=True, CanBeNull:=False)>
        Public Property КодСотрудники() As Integer
        <Column(CanBeNull:=True)>
        Public Property НазвОрганиз() As String
        <Column(CanBeNull:=True)>
        Public Property Фамилия() As String

        <Column(CanBeNull:=True)>
        Public Property Имя() As String

        <Column(CanBeNull:=True)>
        Public Property Отчество() As String

        <Column(CanBeNull:=True)>
        Public Property ФИОСборное() As String

        <Column(CanBeNull:=True)>
        Public Property ФамилияРодПад() As String
        <Column(CanBeNull:=True)>
        Public Property ИмяРодПад() As String
        <Column(CanBeNull:=True)>
        Public Property ОтчествоРодПад() As String
        <Column(CanBeNull:=True)>
        Public Property ФИОРодПод() As String
        <Column(CanBeNull:=True)>
        Public Property ПаспортСерия() As String
        <Column(CanBeNull:=True)>
        Public Property ПаспортНомер() As String
        <Column(CanBeNull:=True)>
        Public Property ПаспортКогдаВыдан() As String
        <Column(CanBeNull:=True)>
        Public Property ДоКакогоДейств() As String
        <Column(CanBeNull:=True)>
        Public Property ПаспортКемВыдан() As String
        <Column(CanBeNull:=True)>
        Public Property ИДНомер() As String
        <Column(CanBeNull:=True)>
        Public Property Регистрация() As String
        <Column(CanBeNull:=True)>
        Public Property МестоПрожив() As String
        <Column(CanBeNull:=True)>
        Public Property КонтТелГор() As String
        <Column(CanBeNull:=True)>
        Public Property КонтТелефон() As String
        <Column(CanBeNull:=True)>
        Public Property ФамилияДляУвольнения() As String
        <Column(CanBeNull:=True)>
        Public Property СтраховойПолис() As String
        <Column(CanBeNull:=True)>
        Public Property НаличеДогПодряда() As String
        <Column(CanBeNull:=True)>
        Public Property ФамилияДляЗаявления() As String
        <Column(CanBeNull:=True)>
        Public Property ИмяДляЗаявления() As String
        <Column(CanBeNull:=True)>
        Public Property ОтчествоДляЗаявления() As String
        <Column(CanBeNull:=True)>
        Public Property Пол() As String
        <Column(CanBeNull:=True)>
        Public Property ФИОДатПадКому() As String
        <Column(CanBeNull:=True)>
        Public Property ДатаРожд() As String
        <Column(CanBeNull:=True)>
        Public Property Гражданин() As String
        <Column(CanBeNull:=True)>
        Public Property ПровДатыКонтр() As String
        <Column>
        Public Property Иностранец() As String
        <Column(CanBeNull:=True)>
        Public Property ФамилияСтар() As String
        <Column(CanBeNull:=True)>
        Public Property ИмяСтар() As String
        <Column(CanBeNull:=True)>
        Public Property ОтчествоСтар() As String
        <Column(CanBeNull:=True)>
        Public Property ФИОСборноеСтар() As String
        <Column(CanBeNull:=True)>
        Public Property ФамилияРодПадСтар() As String
        <Column(CanBeNull:=True)>
        Public Property ИмяРодПадСтар() As String
        <Column(CanBeNull:=True)>
        Public Property ОтчествоРодПадСтар() As String
        <Column(CanBeNull:=True)>
        Public Property ФИОРодПодСтар() As String
        <Column(CanBeNull:=True)>
        Public Property ДатаИзменения() As DateTime

    End Class

End Module
