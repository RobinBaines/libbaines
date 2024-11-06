Imports Utilities

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmUsrLog2
    Inherits frmStandard

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.dgParent = New Utilities.dgvEnter()
        Me.TheDataSet = New TestApp.TheDataSet()
        Me.V_usr_logBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.V_usr_logTableAdapter = New TestApp.TheDataSetTableAdapters.v_usr_logTableAdapter()
        CType(Me.dgParent, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TheDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.V_usr_logBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'HelpTextBox
        '
        Me.HelpTextBox.Location = New System.Drawing.Point(20, 58)
        '
        'dgParent
        '
        Me.dgParent.blnDirty = False
        Me.dgParent.blnMeIsSource = False
        Me.dgParent.blnMove = True
        Me.dgParent.blnRO = False
        Me.dgParent.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgParent.Location = New System.Drawing.Point(12, 58)
        Me.dgParent.Name = "dgParent"
        Me.dgParent.RowEnterInterval = 0
        Me.dgParent.RowIndex = -1
        Me.dgParent.Size = New System.Drawing.Size(907, 387)
        Me.dgParent.ta = Nothing
        Me.dgParent.TabIndex = 237
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
        'frmUsrLog2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = False
        Me.ClientSize = New System.Drawing.Size(1149, 799)
        Me.Controls.Add(Me.dgParent)
        Me.Location = New System.Drawing.Point(0, 0)
        Me.Name = "frmUsrLog2"
        Me.Text = "frmUsrLog2"
        Me.Controls.SetChildIndex(Me.HelpTextBox, 0)
        Me.Controls.SetChildIndex(Me.dgParent, 0)
        CType(Me.dgParent, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TheDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.V_usr_logBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents dgParent As dgvEnter
    Friend WithEvents TheDataSet As TheDataSet
    Friend WithEvents V_usr_logBindingSource As BindingSource
    Friend WithEvents V_usr_logTableAdapter As TheDataSetTableAdapters.v_usr_logTableAdapter
End Class
