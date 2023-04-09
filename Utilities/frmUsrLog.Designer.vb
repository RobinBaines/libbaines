<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmUsrLog
    Inherits frmStandard

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.dgParent = New dgvEnter
        Me.TheDataSet = New TheDataSet
        Me.V_usr_logBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.V_usr_logTableAdapter = New TheDataSetTableAdapters.v_usr_logTableAdapter
        Me.dgChild1 = New dgvEnter
        Me.M_usr_logBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.M_usr_logTableAdapter = New TheDataSetTableAdapters.m_usr_logTableAdapter
        CType(Me.dgParent, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TheDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.V_usr_logBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgChild1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.M_usr_logBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgParent
        '
        Me.dgParent.blnDirty = False
        Me.dgParent.blnMeIsSource = False
        Me.dgParent.blnMove = True
        Me.dgParent.Location = New System.Drawing.Point(9, 52)
        Me.dgParent.Margin = New System.Windows.Forms.Padding(2)
        Me.dgParent.Name = "dgParent"
        Me.dgParent.RowTemplate.Height = 24
        Me.dgParent.Size = New System.Drawing.Size(421, 1000)
        Me.dgParent.ta = Nothing
        Me.dgParent.TabIndex = 1
        '
        'TheDataSet
        '
        Me.TheDataSet.DataSetName = "TheDataSet"
        Me.TheDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'V_usr_logBindingSource
        '
        Me.V_usr_logBindingSource.DataMember = "v_usr_log"
        Me.V_usr_logBindingSource.DataSource = Me.TheDataSet
        '
        'V_usr_logTableAdapter
        '
        Me.V_usr_logTableAdapter.ClearBeforeFill = True
        '
        'dgChild1
        '
        Me.dgChild1.blnDirty = False
        Me.dgChild1.blnMeIsSource = False
        Me.dgChild1.blnMove = True
        Me.dgChild1.Location = New System.Drawing.Point(434, 52)
        Me.dgChild1.Margin = New System.Windows.Forms.Padding(2)
        Me.dgChild1.Name = "dgChild1"
        Me.dgChild1.RowTemplate.Height = 24
        Me.dgChild1.Size = New System.Drawing.Size(397, 1000)
        Me.dgChild1.ta = Nothing
        Me.dgChild1.TabIndex = 2
        '
        'M_usr_logBindingSource
        '
        Me.M_usr_logBindingSource.DataMember = "m_usr_log"
        Me.M_usr_logBindingSource.DataSource = Me.TheDataSet
        '
        'M_usr_logTableAdapter
        '
        Me.M_usr_logTableAdapter.ClearBeforeFill = True
        '
        'frmUsrLog
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.ClientSize = New System.Drawing.Size(862, 759)
        Me.Controls.Add(Me.dgChild1)
        Me.Controls.Add(Me.dgParent)
        Me.Name = "frmUsrLog"
        Me.Controls.SetChildIndex(Me.dgParent, 0)
        Me.Controls.SetChildIndex(Me.dgChild1, 0)
        CType(Me.dgParent, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TheDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.V_usr_logBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgChild1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.M_usr_logBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TheDataSet As TheDataSet
    Friend WithEvents V_usr_logBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents V_usr_logTableAdapter As TheDataSetTableAdapters.v_usr_logTableAdapter
    Friend WithEvents dgParent As dgvEnter
    Friend WithEvents dgChild1 As dgvEnter
    Friend WithEvents M_usr_logBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents M_usr_logTableAdapter As TheDataSetTableAdapters.m_usr_logTableAdapter

End Class
