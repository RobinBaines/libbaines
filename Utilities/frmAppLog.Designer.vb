Imports Utilities
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAppLog
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
        Me.M_app_logBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.M_app_logTableAdapter = New TheDataSetTableAdapters.m_app_logTableAdapter
        Me.dgParent = New dgvEnter
        CType(Me.TheDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.M_app_logBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgParent, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TheDataSet
        '
        Me.TheDataSet.DataSetName = "TheDataSet"
        Me.TheDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'M_app_logBindingSource
        '
        Me.M_app_logBindingSource.DataMember = "m_app_log"
        Me.M_app_logBindingSource.DataSource = Me.TheDataSet
        '
        'M_app_logTableAdapter
        '
        Me.M_app_logTableAdapter.ClearBeforeFill = True
        '
        'dgParent
        '
        Me.dgParent.Location = New System.Drawing.Point(12, 58)
        Me.dgParent.Name = "dgParent"
        Me.dgParent.Size = New System.Drawing.Size(300, 220)
        Me.dgParent.TabIndex = 1
        '
        'frmAppLog
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(862, 472)
        Me.Controls.Add(Me.dgParent)
        Me.Name = "frmAppLog"
        Me.Text = "frmAppLog"
        Me.Controls.SetChildIndex(Me.dgParent, 0)
        CType(Me.TheDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.M_app_logBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgParent, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TheDataSet As TheDataSet
    Friend WithEvents M_app_logBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents M_app_logTableAdapter As TheDataSetTableAdapters.m_app_logTableAdapter
    Friend WithEvents dgParent As dgvEnter
End Class
