﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ШтатноеСписки
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ШтатноеСписки))
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Grid1 = New System.Windows.Forms.DataGridView()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        CType(Me.Grid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ComboBox1
        '
        Me.ComboBox1.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(137, 34)
        Me.ComboBox1.Margin = New System.Windows.Forms.Padding(4)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(448, 25)
        Me.ComboBox1.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(23, 37)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(87, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Организация"
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Button1.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Button1.Location = New System.Drawing.Point(923, 13)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(132, 46)
        Me.Button1.TabIndex = 25
        Me.Button1.Text = "Сформировать расписание"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Grid1
        '
        Me.Grid1.AllowUserToAddRows = False
        Me.Grid1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Grid1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.Grid1.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
        Me.Grid1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Grid1.Location = New System.Drawing.Point(19, 66)
        Me.Grid1.Name = "Grid1"
        Me.Grid1.ReadOnly = True
        Me.Grid1.Size = New System.Drawing.Size(1036, 513)
        Me.Grid1.TabIndex = 26
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(592, 38)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(46, 19)
        Me.Label3.TabIndex = 30
        Me.Label3.Text = "Дата"
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimePicker1.Location = New System.Drawing.Point(644, 34)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(113, 25)
        Me.DateTimePicker1.TabIndex = 31
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(837, 34)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(67, 25)
        Me.TextBox1.TabIndex = 32
        Me.TextBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(780, 37)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(51, 19)
        Me.Label2.TabIndex = 33
        Me.Label2.Text = "Номер"
        '
        'ListBox1
        '
        Me.ListBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.ItemHeight = 17
        Me.ListBox1.Location = New System.Drawing.Point(1074, 66)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(330, 514)
        Me.ListBox1.TabIndex = 34
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Button2.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Button2.Location = New System.Drawing.Point(1074, 32)
        Me.Button2.Margin = New System.Windows.Forms.Padding(4)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(87, 25)
        Me.Button2.TabIndex = 35
        Me.Button2.Text = "EXCEL"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Button3.Enabled = False
        Me.Button3.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Button3.Location = New System.Drawing.Point(1169, 32)
        Me.Button3.Margin = New System.Windows.Forms.Padding(4)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(87, 25)
        Me.Button3.TabIndex = 36
        Me.Button3.Text = "HTML"
        Me.Button3.UseVisualStyleBackColor = False
        '
        'ШтатноеСписки
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1427, 607)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.DateTimePicker1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Grid1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ComboBox1)
        Me.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "ШтатноеСписки"
        Me.Text = "Штатное расписание"
        CType(Me.Grid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Button1 As Button
    Friend WithEvents Grid1 As DataGridView
    Friend WithEvents SaveFileDialog1 As SaveFileDialog
    Friend WithEvents Label3 As Label
    Friend WithEvents DateTimePicker1 As DateTimePicker
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
End Class
