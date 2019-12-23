<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ПриказПродления
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ПриказПродления))
        Me.GroupBox13 = New System.Windows.Forms.GroupBox()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.GroupBox12 = New System.Windows.Forms.GroupBox()
        Me.ComboBox19 = New System.Windows.Forms.ComboBox()
        Me.TextBox57 = New System.Windows.Forms.TextBox()
        Me.TextBox58 = New System.Windows.Forms.TextBox()
        Me.TextBox41 = New System.Windows.Forms.TextBox()
        Me.Label65 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ComboBox3 = New System.Windows.Forms.ComboBox()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.GroupBox13.SuspendLayout()
        Me.GroupBox12.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox13
        '
        Me.GroupBox13.BackColor = System.Drawing.Color.Silver
        Me.GroupBox13.Controls.Add(Me.ComboBox1)
        Me.GroupBox13.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.GroupBox13.Location = New System.Drawing.Point(7, 20)
        Me.GroupBox13.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox13.Name = "GroupBox13"
        Me.GroupBox13.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox13.Size = New System.Drawing.Size(427, 81)
        Me.GroupBox13.TabIndex = 64
        Me.GroupBox13.TabStop = False
        Me.GroupBox13.Text = "Организация"
        '
        'ComboBox1
        '
        Me.ComboBox1.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(8, 26)
        Me.ComboBox1.Margin = New System.Windows.Forms.Padding(4)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(411, 25)
        Me.ComboBox1.TabIndex = 0
        '
        'GroupBox12
        '
        Me.GroupBox12.BackColor = System.Drawing.Color.Silver
        Me.GroupBox12.Controls.Add(Me.ComboBox19)
        Me.GroupBox12.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.GroupBox12.Location = New System.Drawing.Point(441, 20)
        Me.GroupBox12.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox12.Name = "GroupBox12"
        Me.GroupBox12.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox12.Size = New System.Drawing.Size(379, 81)
        Me.GroupBox12.TabIndex = 63
        Me.GroupBox12.TabStop = False
        Me.GroupBox12.Text = "Сотрудник"
        '
        'ComboBox19
        '
        Me.ComboBox19.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.ComboBox19.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource
        Me.ComboBox19.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.ComboBox19.FormattingEnabled = True
        Me.ComboBox19.Location = New System.Drawing.Point(8, 26)
        Me.ComboBox19.Margin = New System.Windows.Forms.Padding(4)
        Me.ComboBox19.Name = "ComboBox19"
        Me.ComboBox19.Size = New System.Drawing.Size(353, 25)
        Me.ComboBox19.TabIndex = 58
        '
        'TextBox57
        '
        Me.TextBox57.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.TextBox57.Location = New System.Drawing.Point(197, 47)
        Me.TextBox57.Margin = New System.Windows.Forms.Padding(4)
        Me.TextBox57.MaxLength = 3
        Me.TextBox57.Name = "TextBox57"
        Me.TextBox57.Size = New System.Drawing.Size(30, 25)
        Me.TextBox57.TabIndex = 68
        Me.TextBox57.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox58
        '
        Me.TextBox58.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.TextBox58.Location = New System.Drawing.Point(163, 47)
        Me.TextBox58.Margin = New System.Windows.Forms.Padding(4)
        Me.TextBox58.MaxLength = 3
        Me.TextBox58.Name = "TextBox58"
        Me.TextBox58.Size = New System.Drawing.Size(26, 25)
        Me.TextBox58.TabIndex = 67
        Me.TextBox58.Text = "к"
        Me.TextBox58.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox41
        '
        Me.TextBox41.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.TextBox41.Location = New System.Drawing.Point(104, 47)
        Me.TextBox41.Margin = New System.Windows.Forms.Padding(4)
        Me.TextBox41.MaxLength = 3
        Me.TextBox41.Name = "TextBox41"
        Me.TextBox41.Size = New System.Drawing.Size(53, 25)
        Me.TextBox41.TabIndex = 65
        Me.TextBox41.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label65
        '
        Me.Label65.AutoSize = True
        Me.Label65.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label65.Location = New System.Drawing.Point(25, 50)
        Me.Label65.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label65.Name = "Label65"
        Me.Label65.Size = New System.Drawing.Size(71, 17)
        Me.Label65.TabIndex = 66
        Me.Label65.Text = "Приказ №"
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Silver
        Me.GroupBox1.Controls.Add(Me.TextBox57)
        Me.GroupBox1.Controls.Add(Me.TextBox58)
        Me.GroupBox1.Controls.Add(Me.TextBox41)
        Me.GroupBox1.Controls.Add(Me.Label65)
        Me.GroupBox1.Location = New System.Drawing.Point(567, 111)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(253, 81)
        Me.GroupBox1.TabIndex = 69
        Me.GroupBox1.TabStop = False
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Button1.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Button1.Location = New System.Drawing.Point(24, 13)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(129, 35)
        Me.Button1.TabIndex = 70
        Me.Button1.Text = "Выполнить"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.Snow
        Me.GroupBox2.Controls.Add(Me.GroupBox3)
        Me.GroupBox2.Controls.Add(Me.GroupBox1)
        Me.GroupBox2.Controls.Add(Me.GroupBox13)
        Me.GroupBox2.Controls.Add(Me.GroupBox12)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(24, 66)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(839, 204)
        Me.GroupBox2.TabIndex = 71
        Me.GroupBox2.TabStop = False
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.Color.Silver
        Me.GroupBox3.Controls.Add(Me.GroupBox4)
        Me.GroupBox3.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.GroupBox3.Location = New System.Drawing.Point(7, 111)
        Me.GroupBox3.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox3.Size = New System.Drawing.Size(553, 81)
        Me.GroupBox3.TabIndex = 65
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Последний номер приказа"
        '
        'GroupBox4
        '
        Me.GroupBox4.BackColor = System.Drawing.Color.Silver
        Me.GroupBox4.Controls.Add(Me.Label2)
        Me.GroupBox4.Controls.Add(Me.Label1)
        Me.GroupBox4.Controls.Add(Me.ComboBox3)
        Me.GroupBox4.Controls.Add(Me.ComboBox2)
        Me.GroupBox4.Location = New System.Drawing.Point(5, 22)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(541, 52)
        Me.GroupBox4.TabIndex = 70
        Me.GroupBox4.TabStop = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label2.Location = New System.Drawing.Point(274, 0)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(53, 17)
        Me.Label2.TabIndex = 69
        Me.Label2.Text = "Список"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label1.Location = New System.Drawing.Point(27, 0)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(31, 17)
        Me.Label1.TabIndex = 69
        Me.Label1.Text = "Год"
        '
        'ComboBox3
        '
        Me.ComboBox3.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.ComboBox3.FormattingEnabled = True
        Me.ComboBox3.Location = New System.Drawing.Point(3, 25)
        Me.ComboBox3.Margin = New System.Windows.Forms.Padding(4)
        Me.ComboBox3.Name = "ComboBox3"
        Me.ComboBox3.Size = New System.Drawing.Size(81, 25)
        Me.ComboBox3.TabIndex = 1
        '
        'ComboBox2
        '
        Me.ComboBox2.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Location = New System.Drawing.Point(92, 25)
        Me.ComboBox2.Margin = New System.Windows.Forms.Padding(4)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(448, 25)
        Me.ComboBox2.TabIndex = 0
        '
        'ПриказПродления
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(885, 297)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.Button1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "ПриказПродления"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Продление контракта"
        Me.GroupBox13.ResumeLayout(False)
        Me.GroupBox12.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox13 As GroupBox
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents GroupBox12 As GroupBox
    Friend WithEvents ComboBox19 As ComboBox
    Friend WithEvents TextBox57 As TextBox
    Friend WithEvents TextBox58 As TextBox
    Friend WithEvents TextBox41 As TextBox
    Friend WithEvents Label65 As Label
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents Button1 As Button
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents ComboBox2 As ComboBox
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents ComboBox3 As ComboBox
End Class
