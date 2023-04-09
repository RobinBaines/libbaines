
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmHelpText
    Inherits frmStandard

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.TheDataSet = New TheDataSet()
        Me.M_form_helptextBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.M_form_helptextTableAdapter = New TheDataSetTableAdapters.m_form_helptextTableAdapter()
        Me.M_form_helptextDataGridView = New dgvEnter()
        CType(Me.TheDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.M_form_helptextBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.M_form_helptextDataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'HelpTextBox
        '
        Me.HelpTextBox.Location = New System.Drawing.Point(20, 58)
        '
        'TheDataSet
        '
        Me.TheDataSet.DataSetName = "TheDataSet"
        Me.TheDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'M_form_helptextBindingSource
        '
        Me.M_form_helptextBindingSource.DataMember = "m_form_helptext"
        Me.M_form_helptextBindingSource.DataSource = Me.TheDataSet
        '
        'M_form_helptextTableAdapter
        '
        Me.M_form_helptextTableAdapter.ClearBeforeFill = True
        '
        'M_form_helptextDataGridView
        '
        Me.M_form_helptextDataGridView.blnDirty = False
        Me.M_form_helptextDataGridView.blnMeIsSource = False
        Me.M_form_helptextDataGridView.blnMove = True
        Me.M_form_helptextDataGridView.blnRO = False
        Me.M_form_helptextDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.M_form_helptextDataGridView.Location = New System.Drawing.Point(20, 75)
        Me.M_form_helptextDataGridView.Name = "M_form_helptextDataGridView"
        Me.M_form_helptextDataGridView.RowEnterInterval = 0
        Me.M_form_helptextDataGridView.RowIndex = -1
        Me.M_form_helptextDataGridView.Size = New System.Drawing.Size(300, 220)
        Me.M_form_helptextDataGridView.ta = Nothing
        Me.M_form_helptextDataGridView.TabIndex = 236
        '
        'frmHelpText
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = False
        Me.ClientSize = New System.Drawing.Size(862, 511)
        Me.Controls.Add(Me.M_form_helptextDataGridView)
        Me.Location = New System.Drawing.Point(0, 0)
        Me.Name = "frmHelpText"
        Me.Text = "frmhelpText"
        Me.Controls.SetChildIndex(Me.M_form_helptextDataGridView, 0)
        Me.Controls.SetChildIndex(Me.HelpTextBox, 0)
        CType(Me.TheDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.M_form_helptextBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.M_form_helptextDataGridView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TheDataSet As TheDataSet
    Friend WithEvents M_form_helptextBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents M_form_helptextTableAdapter As TheDataSetTableAdapters.m_form_helptextTableAdapter
    Friend WithEvents M_form_helptextDataGridView As dgvEnter
End Class
