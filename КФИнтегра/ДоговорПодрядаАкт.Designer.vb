<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ДоговорПодрядаАкт
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ДоговорПодрядаАкт))
        Me.GroupBox13 = New System.Windows.Forms.GroupBox()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.GroupBox12 = New System.Windows.Forms.GroupBox()
        Me.Label96 = New System.Windows.Forms.Label()
        Me.ComboBox19 = New System.Windows.Forms.ComboBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Grid1 = New System.Windows.Forms.DataGridView()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.MaskedTextBox6 = New System.Windows.Forms.MaskedTextBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.CheckBox3 = New System.Windows.Forms.CheckBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.MaskedTextBox1 = New System.Windows.Forms.MaskedTextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.MaskedTextBox2 = New System.Windows.Forms.MaskedTextBox()
        Me.MaskedTextBox3 = New System.Windows.Forms.MaskedTextBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.CheckBox4 = New System.Windows.Forms.CheckBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.ComboBox3 = New System.Windows.Forms.ComboBox()
        Me.ComboBox4 = New System.Windows.Forms.ComboBox()
        Me.GroupBox13.SuspendLayout()
        Me.GroupBox12.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.Grid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox13
        '
        Me.GroupBox13.Controls.Add(Me.ComboBox1)
        Me.GroupBox13.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.GroupBox13.Location = New System.Drawing.Point(296, 13)
        Me.GroupBox13.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox13.Name = "GroupBox13"
        Me.GroupBox13.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox13.Size = New System.Drawing.Size(427, 66)
        Me.GroupBox13.TabIndex = 64
        Me.GroupBox13.TabStop = False
        Me.GroupBox13.Text = "Организация"
        '
        'ComboBox1
        '
        Me.ComboBox1.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.ComboBox1.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource
        Me.ComboBox1.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(20, 23)
        Me.ComboBox1.Margin = New System.Windows.Forms.Padding(4)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(399, 25)
        Me.ComboBox1.TabIndex = 0
        '
        'GroupBox12
        '
        Me.GroupBox12.Controls.Add(Me.Label96)
        Me.GroupBox12.Controls.Add(Me.ComboBox19)
        Me.GroupBox12.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.GroupBox12.Location = New System.Drawing.Point(730, 13)
        Me.GroupBox12.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox12.Name = "GroupBox12"
        Me.GroupBox12.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox12.Size = New System.Drawing.Size(432, 66)
        Me.GroupBox12.TabIndex = 63
        Me.GroupBox12.TabStop = False
        Me.GroupBox12.Text = "Сотрудники"
        '
        'Label96
        '
        Me.Label96.AutoSize = True
        Me.Label96.Location = New System.Drawing.Point(276, -4)
        Me.Label96.Name = "Label96"
        Me.Label96.Size = New System.Drawing.Size(23, 17)
        Me.Label96.TabIndex = 65
        Me.Label96.Text = "№"
        '
        'ComboBox19
        '
        Me.ComboBox19.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.ComboBox19.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource
        Me.ComboBox19.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.ComboBox19.FormattingEnabled = True
        Me.ComboBox19.Location = New System.Drawing.Point(8, 23)
        Me.ComboBox19.Margin = New System.Windows.Forms.Padding(4)
        Me.ComboBox19.Name = "ComboBox19"
        Me.ComboBox19.Size = New System.Drawing.Size(416, 25)
        Me.ComboBox19.TabIndex = 58
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Grid1)
        Me.GroupBox1.Controls.Add(Me.ListBox1)
        Me.GroupBox1.Location = New System.Drawing.Point(478, 156)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(734, 370)
        Me.GroupBox1.TabIndex = 66
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Список актов"
        '
        'Grid1
        '
        Me.Grid1.AllowUserToAddRows = False
        Me.Grid1.AllowUserToDeleteRows = False
        Me.Grid1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.Grid1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.Grid1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Grid1.Location = New System.Drawing.Point(6, 177)
        Me.Grid1.Name = "Grid1"
        Me.Grid1.ReadOnly = True
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Grid1.RowsDefaultCellStyle = DataGridViewCellStyle1
        Me.Grid1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.Grid1.Size = New System.Drawing.Size(722, 181)
        Me.Grid1.TabIndex = 66
        '
        'ListBox1
        '
        Me.ListBox1.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.ItemHeight = 17
        Me.ListBox1.Location = New System.Drawing.Point(6, 24)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.ScrollAlwaysVisible = True
        Me.ListBox1.Size = New System.Drawing.Size(722, 140)
        Me.ListBox1.TabIndex = 65
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(273, 277)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(71, 17)
        Me.Label1.TabIndex = 67
        Me.Label1.Text = "Дата акта"
        '
        'MaskedTextBox6
        '
        Me.MaskedTextBox6.Location = New System.Drawing.Point(350, 274)
        Me.MaskedTextBox6.Mask = "00/00/0000"
        Me.MaskedTextBox6.Name = "MaskedTextBox6"
        Me.MaskedTextBox6.Size = New System.Drawing.Size(92, 25)
        Me.MaskedTextBox6.TabIndex = 68
        Me.MaskedTextBox6.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.MaskedTextBox6.ValidatingType = GetType(Date)
        '
        'TextBox1
        '
        Me.TextBox1.Enabled = False
        Me.TextBox1.Location = New System.Drawing.Point(263, 39)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(92, 25)
        Me.TextBox1.TabIndex = 69
        Me.TextBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(29, 39)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(176, 45)
        Me.Label2.TabIndex = 70
        Me.Label2.Text = "Время, затраченное на выполнение работы, часы"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(20, 94)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(176, 45)
        Me.Label3.TabIndex = 72
        Me.Label3.Text = "Стоимость одного часа, руб."
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(263, 94)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(92, 25)
        Me.TextBox2.TabIndex = 71
        Me.TextBox2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(20, 148)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(176, 45)
        Me.Label4.TabIndex = 74
        Me.Label4.Text = "Стоимость выполненных работ, руб."
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(263, 148)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(92, 25)
        Me.TextBox3.TabIndex = 73
        Me.TextBox3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.CheckBox3)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.TextBox3)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.TextBox2)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.TextBox1)
        Me.GroupBox2.Location = New System.Drawing.Point(87, 322)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(385, 204)
        Me.GroupBox2.TabIndex = 75
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Отработанное время и начисленная сумма"
        '
        'CheckBox3
        '
        Me.CheckBox3.AutoSize = True
        Me.CheckBox3.Location = New System.Drawing.Point(241, 99)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(15, 14)
        Me.CheckBox3.TabIndex = 85
        Me.CheckBox3.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Button1.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Button1.Location = New System.Drawing.Point(19, 13)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(116, 35)
        Me.Button1.TabIndex = 76
        Me.Button1.Text = "Оформить акт"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'MaskedTextBox1
        '
        Me.MaskedTextBox1.Location = New System.Drawing.Point(141, 274)
        Me.MaskedTextBox1.Mask = "00/00/0000"
        Me.MaskedTextBox1.Name = "MaskedTextBox1"
        Me.MaskedTextBox1.Size = New System.Drawing.Size(92, 25)
        Me.MaskedTextBox1.TabIndex = 78
        Me.MaskedTextBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.MaskedTextBox1.ValidatingType = GetType(Date)
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(7, 277)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(128, 17)
        Me.Label5.TabIndex = 77
        Me.Label5.Text = "Дата оплаты работ"
        '
        'ComboBox2
        '
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Location = New System.Drawing.Point(886, 133)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(63, 25)
        Me.ComboBox2.TabIndex = 79
        Me.ComboBox2.Visible = False
        '
        'MaskedTextBox2
        '
        Me.MaskedTextBox2.Location = New System.Drawing.Point(40, 32)
        Me.MaskedTextBox2.Mask = "00/00/0000"
        Me.MaskedTextBox2.Name = "MaskedTextBox2"
        Me.MaskedTextBox2.Size = New System.Drawing.Size(92, 25)
        Me.MaskedTextBox2.TabIndex = 80
        Me.MaskedTextBox2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.MaskedTextBox2.ValidatingType = GetType(Date)
        '
        'MaskedTextBox3
        '
        Me.MaskedTextBox3.Location = New System.Drawing.Point(249, 32)
        Me.MaskedTextBox3.Mask = "00/00/0000"
        Me.MaskedTextBox3.Name = "MaskedTextBox3"
        Me.MaskedTextBox3.Size = New System.Drawing.Size(92, 25)
        Me.MaskedTextBox3.TabIndex = 81
        Me.MaskedTextBox3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.MaskedTextBox3.ValidatingType = GetType(Date)
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Label7)
        Me.GroupBox3.Controls.Add(Me.Label6)
        Me.GroupBox3.Controls.Add(Me.MaskedTextBox3)
        Me.GroupBox3.Controls.Add(Me.MaskedTextBox2)
        Me.GroupBox3.Location = New System.Drawing.Point(101, 179)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(357, 80)
        Me.GroupBox3.TabIndex = 82
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Период"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(212, 35)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(30, 17)
        Me.Label7.TabIndex = 84
        Me.Label7.Text = "ПО"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(15, 35)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(18, 17)
        Me.Label6.TabIndex = 83
        Me.Label6.Text = "С"
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(239, 280)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(15, 14)
        Me.CheckBox1.TabIndex = 83
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Location = New System.Drawing.Point(101, 152)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(372, 21)
        Me.CheckBox2.TabIndex = 84
        Me.CheckBox2.Text = "Отметьте, если дата акта и период (с - по) совпадают!"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(30, 80)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(92, 22)
        Me.Label8.TabIndex = 86
        Me.Label8.Text = "Номер акта"
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(30, 105)
        Me.TextBox4.MaxLength = 3
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(92, 25)
        Me.TextBox4.TabIndex = 85
        Me.TextBox4.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'CheckBox4
        '
        Me.CheckBox4.Location = New System.Drawing.Point(151, 12)
        Me.CheckBox4.Name = "CheckBox4"
        Me.CheckBox4.Size = New System.Drawing.Size(122, 49)
        Me.CheckBox4.TabIndex = 87
        Me.CheckBox4.Text = "Не оформлять Документы"
        Me.CheckBox4.UseVisualStyleBackColor = True
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(162, 80)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(111, 22)
        Me.Label9.TabIndex = 89
        Me.Label9.Text = "Номер договора"
        '
        'ComboBox3
        '
        Me.ComboBox3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox3.FormattingEnabled = True
        Me.ComboBox3.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.ComboBox3.Location = New System.Drawing.Point(151, 105)
        Me.ComboBox3.Name = "ComboBox3"
        Me.ComboBox3.Size = New System.Drawing.Size(122, 25)
        Me.ComboBox3.TabIndex = 90
        '
        'ComboBox4
        '
        Me.ComboBox4.FormattingEnabled = True
        Me.ComboBox4.Location = New System.Drawing.Point(776, 133)
        Me.ComboBox4.Name = "ComboBox4"
        Me.ComboBox4.Size = New System.Drawing.Size(63, 25)
        Me.ComboBox4.TabIndex = 91
        Me.ComboBox4.Visible = False
        '
        'ДоговорПодрядаАкт
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1224, 556)
        Me.Controls.Add(Me.ComboBox4)
        Me.Controls.Add(Me.ComboBox3)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.CheckBox4)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.TextBox4)
        Me.Controls.Add(Me.CheckBox2)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.ComboBox2)
        Me.Controls.Add(Me.MaskedTextBox1)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.MaskedTextBox6)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBox13)
        Me.Controls.Add(Me.GroupBox12)
        Me.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ДоговорПодрядаАкт"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "АКТ Договора подряда"
        Me.GroupBox13.ResumeLayout(False)
        Me.GroupBox12.ResumeLayout(False)
        Me.GroupBox12.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.Grid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents GroupBox13 As GroupBox
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents GroupBox12 As GroupBox
    Friend WithEvents Label96 As Label
    Friend WithEvents ComboBox19 As ComboBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents Label1 As Label
    Friend WithEvents MaskedTextBox6 As MaskedTextBox
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents TextBox3 As TextBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents Button1 As Button
    Friend WithEvents MaskedTextBox1 As MaskedTextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents ComboBox2 As ComboBox
    Friend WithEvents MaskedTextBox2 As MaskedTextBox
    Friend WithEvents MaskedTextBox3 As MaskedTextBox
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents Label7 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents CheckBox2 As CheckBox
    Friend WithEvents CheckBox3 As CheckBox
    Friend WithEvents Label8 As Label
    Friend WithEvents TextBox4 As TextBox
    Friend WithEvents Grid1 As DataGridView
    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents CheckBox4 As CheckBox
    Friend WithEvents Label9 As Label
    Friend WithEvents ComboBox3 As ComboBox
    Friend WithEvents ComboBox4 As ComboBox
End Class
