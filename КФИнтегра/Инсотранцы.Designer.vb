<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Иностранцы
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Иностранцы))
        Me.GroupBox13 = New System.Windows.Forms.GroupBox()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.GroupBox12 = New System.Windows.Forms.GroupBox()
        Me.Label96 = New System.Windows.Forms.Label()
        Me.ComboBox19 = New System.Windows.Forms.ComboBox()
        Me.Grid1 = New System.Windows.Forms.DataGridView()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.GroupBox13.SuspendLayout()
        Me.GroupBox12.SuspendLayout()
        CType(Me.Grid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox13
        '
        Me.GroupBox13.Controls.Add(Me.ComboBox1)
        Me.GroupBox13.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.GroupBox13.Location = New System.Drawing.Point(13, 13)
        Me.GroupBox13.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox13.Name = "GroupBox13"
        Me.GroupBox13.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox13.Size = New System.Drawing.Size(402, 61)
        Me.GroupBox13.TabIndex = 64
        Me.GroupBox13.TabStop = False
        Me.GroupBox13.Text = "Организации"
        '
        'ComboBox1
        '
        Me.ComboBox1.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.ComboBox1.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource
        Me.ComboBox1.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(7, 22)
        Me.ComboBox1.Margin = New System.Windows.Forms.Padding(4)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(386, 25)
        Me.ComboBox1.TabIndex = 0
        '
        'GroupBox12
        '
        Me.GroupBox12.Controls.Add(Me.Label96)
        Me.GroupBox12.Controls.Add(Me.ComboBox19)
        Me.GroupBox12.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.GroupBox12.Location = New System.Drawing.Point(421, 13)
        Me.GroupBox12.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox12.Name = "GroupBox12"
        Me.GroupBox12.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox12.Size = New System.Drawing.Size(417, 61)
        Me.GroupBox12.TabIndex = 63
        Me.GroupBox12.TabStop = False
        Me.GroupBox12.Text = "Иностранцы"
        '
        'Label96
        '
        Me.Label96.AutoSize = True
        Me.Label96.Location = New System.Drawing.Point(245, -4)
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
        Me.ComboBox19.Location = New System.Drawing.Point(7, 22)
        Me.ComboBox19.Margin = New System.Windows.Forms.Padding(4)
        Me.ComboBox19.Name = "ComboBox19"
        Me.ComboBox19.Size = New System.Drawing.Size(403, 25)
        Me.ComboBox19.TabIndex = 58
        '
        'Grid1
        '
        Me.Grid1.AllowUserToAddRows = False
        Me.Grid1.AllowUserToDeleteRows = False
        Me.Grid1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Grid1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.Grid1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.Grid1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Grid1.Location = New System.Drawing.Point(12, 100)
        Me.Grid1.Name = "Grid1"
        Me.Grid1.ReadOnly = True
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Grid1.RowsDefaultCellStyle = DataGridViewCellStyle1
        Me.Grid1.Size = New System.Drawing.Size(1121, 478)
        Me.Grid1.TabIndex = 65
        '
        'ComboBox2
        '
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Location = New System.Drawing.Point(375, 9)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(31, 25)
        Me.ComboBox2.TabIndex = 66
        Me.ComboBox2.Visible = False
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(13, 73)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(51, 21)
        Me.CheckBox1.TabIndex = 67
        Me.CheckBox1.Text = "Все"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Иностранцы
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1143, 588)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.ComboBox2)
        Me.Controls.Add(Me.Grid1)
        Me.Controls.Add(Me.GroupBox13)
        Me.Controls.Add(Me.GroupBox12)
        Me.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Иностранцы"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Список иностранных сотрудников"
        Me.GroupBox13.ResumeLayout(False)
        Me.GroupBox12.ResumeLayout(False)
        Me.GroupBox12.PerformLayout()
        CType(Me.Grid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents GroupBox13 As GroupBox
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents GroupBox12 As GroupBox
    Friend WithEvents Label96 As Label
    Friend WithEvents ComboBox19 As ComboBox
    Friend WithEvents Grid1 As DataGridView
    Friend WithEvents ComboBox2 As ComboBox
    Friend WithEvents CheckBox1 As CheckBox
End Class
