<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ОтчетФорма
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
        Dim ReportDataSource1 As Microsoft.Reporting.WinForms.ReportDataSource = New Microsoft.Reporting.WinForms.ReportDataSource()
        Me.ReportViewer1 = New Microsoft.Reporting.WinForms.ReportViewer()
        Me.ИнтеграDataSet = New ИнтеграКадры.ИнтеграDataSet()
        Me.ДогПодОсобенBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.ДогПодОсобенTableAdapter = New ИнтеграКадры.ИнтеграDataSetTableAdapters.ДогПодОсобенTableAdapter()
        CType(Me.ИнтеграDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ДогПодОсобенBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ReportViewer1
        '
        Me.ReportViewer1.Dock = System.Windows.Forms.DockStyle.Fill
        ReportDataSource1.Name = "DataSet1"
        ReportDataSource1.Value = Me.ДогПодОсобенBindingSource
        Me.ReportViewer1.LocalReport.DataSources.Add(ReportDataSource1)
        Me.ReportViewer1.LocalReport.ReportEmbeddedResource = "ИнтеграКадры.Report2.rdlc"
        Me.ReportViewer1.Location = New System.Drawing.Point(0, 0)
        Me.ReportViewer1.Name = "ReportViewer1"
        Me.ReportViewer1.ServerReport.BearerToken = Nothing
        Me.ReportViewer1.Size = New System.Drawing.Size(800, 450)
        Me.ReportViewer1.TabIndex = 0
        '
        'ИнтеграDataSet
        '
        Me.ИнтеграDataSet.DataSetName = "ИнтеграDataSet"
        Me.ИнтеграDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'ДогПодОсобенBindingSource
        '
        Me.ДогПодОсобенBindingSource.DataMember = "ДогПодОсобен"
        Me.ДогПодОсобенBindingSource.DataSource = Me.ИнтеграDataSet
        '
        'ДогПодОсобенTableAdapter
        '
        Me.ДогПодОсобенTableAdapter.ClearBeforeFill = True
        '
        'ОтчетФорма
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.ReportViewer1)
        Me.Name = "ОтчетФорма"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ОтчетФорма"
        CType(Me.ИнтеграDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ДогПодОсобенBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents ReportViewer1 As Microsoft.Reporting.WinForms.ReportViewer
    Friend WithEvents ДогПодОсобенBindingSource As BindingSource
    Friend WithEvents ИнтеграDataSet As ИнтеграDataSet
    Friend WithEvents ДогПодОсобенTableAdapter As ИнтеграDataSetTableAdapters.ДогПодОсобенTableAdapter
End Class
