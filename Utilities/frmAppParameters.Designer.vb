Imports Utilities
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAppParameters
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
        Me.components = New System.ComponentModel.Container
        Me.TheDataSet = New TheDataSet
        Me.B_app_parameterBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.B_app_parameterTableAdapter = New TheDataSetTableAdapters.b_app_parameterTableAdapter
        Me.dgParent1 = New dgvEnter
        Me.B_app_colorBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.B_app_colorTableAdapter = New TheDataSetTableAdapters.b_app_colorTableAdapter
        Me.dgParent2 = New dgvEnter
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.TextBox1 = New System.Windows.Forms.TextBox
        CType(Me.TheDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.B_app_parameterBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgParent1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.B_app_colorBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgParent2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TheDataSet
        '
        Me.TheDataSet.DataSetName = "TheDataSet"
        Me.TheDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'B_app_parameterBindingSource
        '
        Me.B_app_parameterBindingSource.DataMember = "b_app_parameter"
        Me.B_app_parameterBindingSource.DataSource = Me.TheDataSet
        '
        'B_app_parameterTableAdapter
        '
        Me.B_app_parameterTableAdapter.ClearBeforeFill = True
        '
        'dgParent1
        '
        Me.dgParent1.blnDirty = False
        Me.dgParent1.Location = New System.Drawing.Point(12, 68)
        Me.dgParent1.Name = "dgParent1"
        Me.dgParent1.Size = New System.Drawing.Size(621, 291)
        Me.dgParent1.ta = Nothing
        Me.dgParent1.TabIndex = 1
        '
        'B_app_colorBindingSource
        '
        Me.B_app_colorBindingSource.DataMember = "b_app_color"
        Me.B_app_colorBindingSource.DataSource = Me.TheDataSet
        '
        'B_app_colorTableAdapter
        '
        Me.B_app_colorTableAdapter.ClearBeforeFill = True
        '
        'dgParent2
        '
        Me.dgParent2.blnDirty = False
        Me.dgParent2.Location = New System.Drawing.Point(12, 397)
        Me.dgParent2.Name = "dgParent2"
        Me.dgParent2.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgParent2.Size = New System.Drawing.Size(599, 291)
        Me.dgParent2.ta = Nothing
        Me.dgParent2.TabIndex = 2
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(906, 397)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(120, 95)
        Me.ListBox1.TabIndex = 4
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(1062, 397)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(100, 20)
        Me.TextBox1.TabIndex = 5
        '
        'frmAppParameters
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1311, 762)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.dgParent2)
        Me.Controls.Add(Me.dgParent1)
        Me.Name = "frmAppParameters"
        Me.Text = "frmAppParameters"
        Me.Controls.SetChildIndex(Me.dgParent1, 0)
        Me.Controls.SetChildIndex(Me.dgParent2, 0)
        Me.Controls.SetChildIndex(Me.ListBox1, 0)
        Me.Controls.SetChildIndex(Me.TextBox1, 0)
        CType(Me.TheDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.B_app_parameterBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgParent1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.B_app_colorBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgParent2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TheDataSet As TheDataSet
    Friend WithEvents B_app_parameterBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents B_app_parameterTableAdapter As TheDataSetTableAdapters.b_app_parameterTableAdapter
    Friend WithEvents dgParent1 As dgvEnter
    Friend WithEvents B_app_colorBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents B_app_colorTableAdapter As TheDataSetTableAdapters.b_app_colorTableAdapter
    Friend WithEvents dgParent2 As dgvEnter
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
End Class
