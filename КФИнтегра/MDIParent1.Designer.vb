<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class MDIParent1
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub


    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MDIParent1))
        Me.MenuStrip = New System.Windows.Forms.MenuStrip()
        Me.ВводДанныхToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ОрганизацияToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ИзменитьToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.КадрыToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.КласическоеToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПослеToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ОтчетыToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПриемНаРаботуToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.УвольнениеToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПроодлениеКонтрактаToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПродлениеКонтрактаToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПриказToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ДополнительныеСоглашенияToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.УведомлениеОбИзмененеииОкладаToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.УведомлениеОбИзмененииСроковВыплатыЗарплатыToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.УведомлениеОбИзмененииУслвоияТрудаРестлайнToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПереводToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ГрафикОтпусковToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ТекущийToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПоТребованиюToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.СправкиToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ОбУровнеЗарплатыToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.АктыToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ДоговорПодрядаToolStripMenuItem2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПоЧасамToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ИноеToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.СоздатьToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ИзменитьToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.СпискиToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.УволенныеToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПринятыеToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ДоговорПодрядаToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ШтатноеToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.НеподписанныеDokumentyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ИностранцыToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ОтпускToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.СтатистикаToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.УведомлениеОПродленииКонтрактаToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ОтчетToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.СотрудникиToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПечатьToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПоискToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StatusStrip = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.MenuStrip.SuspendLayout()
        Me.StatusStrip.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip
        '
        Me.MenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ВводДанныхToolStripMenuItem, Me.ОтчетыToolStripMenuItem, Me.СпискиToolStripMenuItem, Me.ПечатьToolStripMenuItem, Me.ПоискToolStripMenuItem})
        Me.MenuStrip.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip.Name = "MenuStrip"
        Me.MenuStrip.Padding = New System.Windows.Forms.Padding(8, 3, 0, 3)
        Me.MenuStrip.Size = New System.Drawing.Size(1604, 25)
        Me.MenuStrip.TabIndex = 5
        Me.MenuStrip.Text = "MenuStrip"
        '
        'ВводДанныхToolStripMenuItem
        '
        Me.ВводДанныхToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ОрганизацияToolStripMenuItem, Me.ИзменитьToolStripMenuItem, Me.КадрыToolStripMenuItem})
        Me.ВводДанныхToolStripMenuItem.Name = "ВводДанныхToolStripMenuItem"
        Me.ВводДанныхToolStripMenuItem.Size = New System.Drawing.Size(91, 19)
        Me.ВводДанныхToolStripMenuItem.Text = "Организация"
        '
        'ОрганизацияToolStripMenuItem
        '
        Me.ОрганизацияToolStripMenuItem.Name = "ОрганизацияToolStripMenuItem"
        Me.ОрганизацияToolStripMenuItem.Size = New System.Drawing.Size(128, 22)
        Me.ОрганизацияToolStripMenuItem.Text = "Добавить"
        '
        'ИзменитьToolStripMenuItem
        '
        Me.ИзменитьToolStripMenuItem.Name = "ИзменитьToolStripMenuItem"
        Me.ИзменитьToolStripMenuItem.Size = New System.Drawing.Size(128, 22)
        Me.ИзменитьToolStripMenuItem.Text = "Изменить"
        '
        'КадрыToolStripMenuItem
        '
        Me.КадрыToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.КласическоеToolStripMenuItem, Me.ПослеToolStripMenuItem})
        Me.КадрыToolStripMenuItem.Name = "КадрыToolStripMenuItem"
        Me.КадрыToolStripMenuItem.Size = New System.Drawing.Size(128, 22)
        Me.КадрыToolStripMenuItem.Text = "Штатное"
        '
        'КласическоеToolStripMenuItem
        '
        Me.КласическоеToolStripMenuItem.Name = "КласическоеToolStripMenuItem"
        Me.КласическоеToolStripMenuItem.Size = New System.Drawing.Size(235, 22)
        Me.КласическоеToolStripMenuItem.Text = "Класическое"
        '
        'ПослеToolStripMenuItem
        '
        Me.ПослеToolStripMenuItem.Enabled = False
        Me.ПослеToolStripMenuItem.Name = "ПослеToolStripMenuItem"
        Me.ПослеToolStripMenuItem.Size = New System.Drawing.Size(235, 22)
        Me.ПослеToolStripMenuItem.Text = "Изменение после испытания"
        '
        'ОтчетыToolStripMenuItem
        '
        Me.ОтчетыToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ПриемНаРаботуToolStripMenuItem, Me.УвольнениеToolStripMenuItem1, Me.ПроодлениеКонтрактаToolStripMenuItem, Me.ПереводToolStripMenuItem, Me.ГрафикОтпусковToolStripMenuItem, Me.СправкиToolStripMenuItem, Me.АктыToolStripMenuItem})
        Me.ОтчетыToolStripMenuItem.Name = "ОтчетыToolStripMenuItem"
        Me.ОтчетыToolStripMenuItem.Size = New System.Drawing.Size(54, 19)
        Me.ОтчетыToolStripMenuItem.Text = "Кадры"
        '
        'ПриемНаРаботуToolStripMenuItem
        '
        Me.ПриемНаРаботуToolStripMenuItem.Name = "ПриемНаРаботуToolStripMenuItem"
        Me.ПриемНаРаботуToolStripMenuItem.Size = New System.Drawing.Size(274, 22)
        Me.ПриемНаРаботуToolStripMenuItem.Text = "Прием (Контракт, Договор подряда)"
        '
        'УвольнениеToolStripMenuItem1
        '
        Me.УвольнениеToolStripMenuItem1.Name = "УвольнениеToolStripMenuItem1"
        Me.УвольнениеToolStripMenuItem1.Size = New System.Drawing.Size(274, 22)
        Me.УвольнениеToolStripMenuItem1.Text = "Увольнение"
        '
        'ПроодлениеКонтрактаToolStripMenuItem
        '
        Me.ПроодлениеКонтрактаToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ПродлениеКонтрактаToolStripMenuItem, Me.ПриказToolStripMenuItem, Me.ДополнительныеСоглашенияToolStripMenuItem})
        Me.ПроодлениеКонтрактаToolStripMenuItem.Name = "ПроодлениеКонтрактаToolStripMenuItem"
        Me.ПроодлениеКонтрактаToolStripMenuItem.Size = New System.Drawing.Size(274, 22)
        Me.ПроодлениеКонтрактаToolStripMenuItem.Text = "Контракт"
        '
        'ПродлениеКонтрактаToolStripMenuItem
        '
        Me.ПродлениеКонтрактаToolStripMenuItem.Name = "ПродлениеКонтрактаToolStripMenuItem"
        Me.ПродлениеКонтрактаToolStripMenuItem.Size = New System.Drawing.Size(297, 22)
        Me.ПродлениеКонтрактаToolStripMenuItem.Text = "Продление или не продление контракта"
        '
        'ПриказToolStripMenuItem
        '
        Me.ПриказToolStripMenuItem.Name = "ПриказToolStripMenuItem"
        Me.ПриказToolStripMenuItem.Size = New System.Drawing.Size(297, 22)
        Me.ПриказToolStripMenuItem.Text = "Приказ продления контракта"
        '
        'ДополнительныеСоглашенияToolStripMenuItem
        '
        Me.ДополнительныеСоглашенияToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.УведомлениеОбИзмененеииОкладаToolStripMenuItem, Me.УведомлениеОбИзмененииСроковВыплатыЗарплатыToolStripMenuItem, Me.УведомлениеОбИзмененииУслвоияТрудаРестлайнToolStripMenuItem})
        Me.ДополнительныеСоглашенияToolStripMenuItem.Name = "ДополнительныеСоглашенияToolStripMenuItem"
        Me.ДополнительныеСоглашенияToolStripMenuItem.Size = New System.Drawing.Size(297, 22)
        Me.ДополнительныеСоглашенияToolStripMenuItem.Text = "Дополнительные соглашения"
        '
        'УведомлениеОбИзмененеииОкладаToolStripMenuItem
        '
        Me.УведомлениеОбИзмененеииОкладаToolStripMenuItem.Name = "УведомлениеОбИзмененеииОкладаToolStripMenuItem"
        Me.УведомлениеОбИзмененеииОкладаToolStripMenuItem.Size = New System.Drawing.Size(378, 22)
        Me.УведомлениеОбИзмененеииОкладаToolStripMenuItem.Text = "Уведомление(универсальное)"
        '
        'УведомлениеОбИзмененииСроковВыплатыЗарплатыToolStripMenuItem
        '
        Me.УведомлениеОбИзмененииСроковВыплатыЗарплатыToolStripMenuItem.Name = "УведомлениеОбИзмененииСроковВыплатыЗарплатыToolStripMenuItem"
        Me.УведомлениеОбИзмененииСроковВыплатыЗарплатыToolStripMenuItem.Size = New System.Drawing.Size(378, 22)
        Me.УведомлениеОбИзмененииСроковВыплатыЗарплатыToolStripMenuItem.Text = "Уведомление об изменении сроков выплаты зарплаты"
        '
        'УведомлениеОбИзмененииУслвоияТрудаРестлайнToolStripMenuItem
        '
        Me.УведомлениеОбИзмененииУслвоияТрудаРестлайнToolStripMenuItem.Name = "УведомлениеОбИзмененииУслвоияТрудаРестлайнToolStripMenuItem"
        Me.УведомлениеОбИзмененииУслвоияТрудаРестлайнToolStripMenuItem.Size = New System.Drawing.Size(378, 22)
        Me.УведомлениеОбИзмененииУслвоияТрудаРестлайнToolStripMenuItem.Text = "Уведомление об изменении услвоия труда(Рестлайн)"
        '
        'ПереводToolStripMenuItem
        '
        Me.ПереводToolStripMenuItem.Name = "ПереводToolStripMenuItem"
        Me.ПереводToolStripMenuItem.Size = New System.Drawing.Size(274, 22)
        Me.ПереводToolStripMenuItem.Text = "Перевод"
        '
        'ГрафикОтпусковToolStripMenuItem
        '
        Me.ГрафикОтпусковToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ТекущийToolStripMenuItem, Me.ПоТребованиюToolStripMenuItem})
        Me.ГрафикОтпусковToolStripMenuItem.Name = "ГрафикОтпусковToolStripMenuItem"
        Me.ГрафикОтпусковToolStripMenuItem.Size = New System.Drawing.Size(274, 22)
        Me.ГрафикОтпусковToolStripMenuItem.Text = "Отпуск"
        '
        'ТекущийToolStripMenuItem
        '
        Me.ТекущийToolStripMenuItem.Name = "ТекущийToolStripMenuItem"
        Me.ТекущийToolStripMenuItem.Size = New System.Drawing.Size(145, 22)
        Me.ТекущийToolStripMenuItem.Text = "Трудовой"
        '
        'ПоТребованиюToolStripMenuItem
        '
        Me.ПоТребованиюToolStripMenuItem.Name = "ПоТребованиюToolStripMenuItem"
        Me.ПоТребованиюToolStripMenuItem.Size = New System.Drawing.Size(145, 22)
        Me.ПоТребованиюToolStripMenuItem.Text = "Социальный"
        '
        'СправкиToolStripMenuItem
        '
        Me.СправкиToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ОбУровнеЗарплатыToolStripMenuItem})
        Me.СправкиToolStripMenuItem.Name = "СправкиToolStripMenuItem"
        Me.СправкиToolStripMenuItem.Size = New System.Drawing.Size(274, 22)
        Me.СправкиToolStripMenuItem.Text = "Справки"
        '
        'ОбУровнеЗарплатыToolStripMenuItem
        '
        Me.ОбУровнеЗарплатыToolStripMenuItem.Name = "ОбУровнеЗарплатыToolStripMenuItem"
        Me.ОбУровнеЗарплатыToolStripMenuItem.Size = New System.Drawing.Size(187, 22)
        Me.ОбУровнеЗарплатыToolStripMenuItem.Text = "Об уровне зарплаты"
        '
        'АктыToolStripMenuItem
        '
        Me.АктыToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ДоговорПодрядаToolStripMenuItem2})
        Me.АктыToolStripMenuItem.Name = "АктыToolStripMenuItem"
        Me.АктыToolStripMenuItem.Size = New System.Drawing.Size(274, 22)
        Me.АктыToolStripMenuItem.Text = "Акты"
        '
        'ДоговорПодрядаToolStripMenuItem2
        '
        Me.ДоговорПодрядаToolStripMenuItem2.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ПоЧасамToolStripMenuItem, Me.ИноеToolStripMenuItem})
        Me.ДоговорПодрядаToolStripMenuItem2.Name = "ДоговорПодрядаToolStripMenuItem2"
        Me.ДоговорПодрядаToolStripMenuItem2.Size = New System.Drawing.Size(180, 22)
        Me.ДоговорПодрядаToolStripMenuItem2.Text = "Договор подряда"
        '
        'ПоЧасамToolStripMenuItem
        '
        Me.ПоЧасамToolStripMenuItem.Name = "ПоЧасамToolStripMenuItem"
        Me.ПоЧасамToolStripMenuItem.Size = New System.Drawing.Size(180, 22)
        Me.ПоЧасамToolStripMenuItem.Text = "По часам"
        '
        'ИноеToolStripMenuItem
        '
        Me.ИноеToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.СоздатьToolStripMenuItem, Me.ИзменитьToolStripMenuItem1})
        Me.ИноеToolStripMenuItem.Name = "ИноеToolStripMenuItem"
        Me.ИноеToolStripMenuItem.Size = New System.Drawing.Size(180, 22)
        Me.ИноеToolStripMenuItem.Text = "Иное"
        '
        'СоздатьToolStripMenuItem
        '
        Me.СоздатьToolStripMenuItem.Name = "СоздатьToolStripMenuItem"
        Me.СоздатьToolStripMenuItem.Size = New System.Drawing.Size(180, 22)
        Me.СоздатьToolStripMenuItem.Text = "Создать"
        '
        'ИзменитьToolStripMenuItem1
        '
        Me.ИзменитьToolStripMenuItem1.Name = "ИзменитьToolStripMenuItem1"
        Me.ИзменитьToolStripMenuItem1.Size = New System.Drawing.Size(180, 22)
        Me.ИзменитьToolStripMenuItem1.Text = "Изменить"
        '
        'СпискиToolStripMenuItem
        '
        Me.СпискиToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.УволенныеToolStripMenuItem, Me.ПринятыеToolStripMenuItem, Me.ДоговорПодрядаToolStripMenuItem1, Me.ШтатноеToolStripMenuItem, Me.НеподписанныеDokumentyToolStripMenuItem, Me.ИностранцыToolStripMenuItem, Me.ОтпускToolStripMenuItem, Me.СтатистикаToolStripMenuItem, Me.УведомлениеОПродленииКонтрактаToolStripMenuItem, Me.ОтчетToolStripMenuItem})
        Me.СпискиToolStripMenuItem.Name = "СпискиToolStripMenuItem"
        Me.СпискиToolStripMenuItem.Size = New System.Drawing.Size(60, 19)
        Me.СпискиToolStripMenuItem.Text = "Отчеты"
        '
        'УволенныеToolStripMenuItem
        '
        Me.УволенныеToolStripMenuItem.Name = "УволенныеToolStripMenuItem"
        Me.УволенныеToolStripMenuItem.Size = New System.Drawing.Size(287, 22)
        Me.УволенныеToolStripMenuItem.Text = "Уволенные"
        '
        'ПринятыеToolStripMenuItem
        '
        Me.ПринятыеToolStripMenuItem.Name = "ПринятыеToolStripMenuItem"
        Me.ПринятыеToolStripMenuItem.Size = New System.Drawing.Size(287, 22)
        Me.ПринятыеToolStripMenuItem.Text = "Принятые"
        '
        'ДоговорПодрядаToolStripMenuItem1
        '
        Me.ДоговорПодрядаToolStripMenuItem1.Name = "ДоговорПодрядаToolStripMenuItem1"
        Me.ДоговорПодрядаToolStripMenuItem1.Size = New System.Drawing.Size(287, 22)
        Me.ДоговорПодрядаToolStripMenuItem1.Text = "Договор подряда"
        '
        'ШтатноеToolStripMenuItem
        '
        Me.ШтатноеToolStripMenuItem.Name = "ШтатноеToolStripMenuItem"
        Me.ШтатноеToolStripMenuItem.Size = New System.Drawing.Size(287, 22)
        Me.ШтатноеToolStripMenuItem.Text = "Штатное"
        '
        'НеподписанныеDokumentyToolStripMenuItem
        '
        Me.НеподписанныеDokumentyToolStripMenuItem.Name = "НеподписанныеDokumentyToolStripMenuItem"
        Me.НеподписанныеDokumentyToolStripMenuItem.Size = New System.Drawing.Size(287, 22)
        Me.НеподписанныеDokumentyToolStripMenuItem.Text = "Неподписанные документы"
        '
        'ИностранцыToolStripMenuItem
        '
        Me.ИностранцыToolStripMenuItem.Name = "ИностранцыToolStripMenuItem"
        Me.ИностранцыToolStripMenuItem.Size = New System.Drawing.Size(287, 22)
        Me.ИностранцыToolStripMenuItem.Text = "Иностранцы"
        '
        'ОтпускToolStripMenuItem
        '
        Me.ОтпускToolStripMenuItem.Name = "ОтпускToolStripMenuItem"
        Me.ОтпускToolStripMenuItem.Size = New System.Drawing.Size(287, 22)
        Me.ОтпускToolStripMenuItem.Text = "Отпуск"
        '
        'СтатистикаToolStripMenuItem
        '
        Me.СтатистикаToolStripMenuItem.Name = "СтатистикаToolStripMenuItem"
        Me.СтатистикаToolStripMenuItem.Size = New System.Drawing.Size(287, 22)
        Me.СтатистикаToolStripMenuItem.Text = "Статистика"
        '
        'УведомлениеОПродленииКонтрактаToolStripMenuItem
        '
        Me.УведомлениеОПродленииКонтрактаToolStripMenuItem.Name = "УведомлениеОПродленииКонтрактаToolStripMenuItem"
        Me.УведомлениеОПродленииКонтрактаToolStripMenuItem.Size = New System.Drawing.Size(287, 22)
        Me.УведомлениеОПродленииКонтрактаToolStripMenuItem.Text = "Уведомление о продлении контрактов"
        '
        'ОтчетToolStripMenuItem
        '
        Me.ОтчетToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.СотрудникиToolStripMenuItem})
        Me.ОтчетToolStripMenuItem.Name = "ОтчетToolStripMenuItem"
        Me.ОтчетToolStripMenuItem.Size = New System.Drawing.Size(287, 22)
        Me.ОтчетToolStripMenuItem.Text = "Ведомость"
        '
        'СотрудникиToolStripMenuItem
        '
        Me.СотрудникиToolStripMenuItem.Name = "СотрудникиToolStripMenuItem"
        Me.СотрудникиToolStripMenuItem.Size = New System.Drawing.Size(133, 22)
        Me.СотрудникиToolStripMenuItem.Text = "Сотрудник"
        '
        'ПечатьToolStripMenuItem
        '
        Me.ПечатьToolStripMenuItem.Name = "ПечатьToolStripMenuItem"
        Me.ПечатьToolStripMenuItem.Size = New System.Drawing.Size(82, 19)
        Me.ПечатьToolStripMenuItem.Text = "Документы"
        '
        'ПоискToolStripMenuItem
        '
        Me.ПоискToolStripMenuItem.Name = "ПоискToolStripMenuItem"
        Me.ПоискToolStripMenuItem.Size = New System.Drawing.Size(62, 19)
        Me.ПоискToolStripMenuItem.Text = "Данные"
        '
        'StatusStrip
        '
        Me.StatusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel})
        Me.StatusStrip.Location = New System.Drawing.Point(0, 859)
        Me.StatusStrip.Name = "StatusStrip"
        Me.StatusStrip.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip.Size = New System.Drawing.Size(1604, 22)
        Me.StatusStrip.TabIndex = 7
        Me.StatusStrip.Text = "StatusStrip"
        '
        'ToolStripStatusLabel
        '
        Me.ToolStripStatusLabel.Name = "ToolStripStatusLabel"
        Me.ToolStripStatusLabel.Size = New System.Drawing.Size(66, 17)
        Me.ToolStripStatusLabel.Text = "Состояние"
        '
        'ToolTip
        '
        '
        'MDIParent1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.ClientSize = New System.Drawing.Size(1604, 881)
        Me.Controls.Add(Me.MenuStrip)
        Me.Controls.Add(Me.StatusStrip)
        Me.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.MainMenuStrip = Me.MenuStrip
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "MDIParent1"
        Me.Text = "Кадры"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.MenuStrip.ResumeLayout(False)
        Me.MenuStrip.PerformLayout()
        Me.StatusStrip.ResumeLayout(False)
        Me.StatusStrip.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ToolTip As System.Windows.Forms.ToolTip
    Friend WithEvents ToolStripStatusLabel As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip As System.Windows.Forms.StatusStrip
    Friend WithEvents MenuStrip As System.Windows.Forms.MenuStrip
    Friend WithEvents ВводДанныхToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ОрганизацияToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ОтчетыToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ПроодлениеКонтрактаToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents КадрыToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents УвольнениеToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ИзменитьToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ПриемНаРаботуToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ПриказToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ПродлениеКонтрактаToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents СпискиToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents УволенныеToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ПринятыеToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ДоговорПодрядаToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ПечатьToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ШтатноеToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents КласическоеToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ПослеToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ПоискToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ПереводToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ГрафикОтпусковToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents НеподписанныеDokumentyToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents СправкиToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ОбУровнеЗарплатыToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ИностранцыToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents СтатистикаToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ТекущийToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ПоТребованиюToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents УведомлениеОПродленииКонтрактаToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents АктыToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ДоговорПодрядаToolStripMenuItem2 As ToolStripMenuItem
    Friend WithEvents ПоЧасамToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ИноеToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ДополнительныеСоглашенияToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents УведомлениеОбИзмененииСроковВыплатыЗарплатыToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents УведомлениеОбИзмененииУслвоияТрудаРестлайнToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents УведомлениеОбИзмененеииОкладаToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ОтпускToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ОтчетToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents СотрудникиToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents СоздатьToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ИзменитьToolStripMenuItem1 As ToolStripMenuItem
End Class
