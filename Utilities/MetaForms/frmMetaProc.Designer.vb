<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMetaProc
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
        Me.MetaData = New MetaDataSet()
        Me.dgROUTINES = New dgvEnter()
        Me.tbROUTINES = New System.Windows.Forms.RichTextBox()
        Me.ROUTINESBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.dgReferenced = New dgvEnter()
        Me.dgReferencing = New dgvEnter()
        Me.TableAdapterManager = New MetaDataSetTableAdapters.TableAdapterManager()
        Me.ROUTINESTableAdapter = New MetaDataSetTableAdapters.ROUTINESTableAdapter()
        Me.Dm_sql_referenced_entitiesBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.Dm_sql_referenced_entitiesTableAdapter = New MetaDataSetTableAdapters.dm_sql_referenced_entitiesTableAdapter()
        Me.Dm_sql_referencing_entitiesBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.Dm_sql_referencing_entitiesTableAdapter = New MetaDataSetTableAdapters.dm_sql_referencing_entitiesTableAdapter()
        Me.bProcCopy = New System.Windows.Forms.Button()
        CType(Me.MetaData, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgROUTINES, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ROUTINESBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgReferenced, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgReferencing, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Dm_sql_referenced_entitiesBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Dm_sql_referencing_entitiesBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'HelpTextBox
        '
        Me.HelpTextBox.Location = New System.Drawing.Point(20, 58)
        '
        'MetaData
        '
        Me.MetaData.DataSetName = "MetaData"
        Me.MetaData.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'dgROUTINES
        '
        Me.dgROUTINES.blnDirty = False
        Me.dgROUTINES.blnMeIsSource = False
        Me.dgROUTINES.blnMove = True
        Me.dgROUTINES.blnRO = False
        Me.dgROUTINES.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgROUTINES.Location = New System.Drawing.Point(12, 57)
        Me.dgROUTINES.Name = "dgROUTINES"
        Me.dgROUTINES.RowEnterInterval = 0
        Me.dgROUTINES.RowIndex = -1
        Me.dgROUTINES.Size = New System.Drawing.Size(300, 562)
        Me.dgROUTINES.ta = Nothing
        Me.dgROUTINES.TabIndex = 109
        '
        'tbROUTINES
        '
        Me.tbROUTINES.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.ROUTINESBindingSource, "ROUTINE_DEFINITION", True))
        Me.tbROUTINES.Font = New System.Drawing.Font("Courier New", 9.75!)
        Me.tbROUTINES.Location = New System.Drawing.Point(330, 57)
        Me.tbROUTINES.Name = "tbROUTINES"
        Me.tbROUTINES.ReadOnly = True
        Me.tbROUTINES.Size = New System.Drawing.Size(800, 562)
        Me.tbROUTINES.TabIndex = 110
        Me.tbROUTINES.Text = ""
        Me.tbROUTINES.WordWrap = False
        '
        'ROUTINESBindingSource
        '
        Me.ROUTINESBindingSource.DataMember = "ROUTINES"
        Me.ROUTINESBindingSource.DataSource = Me.MetaData
        '
        'dgReferenced
        '
        Me.dgReferenced.blnDirty = False
        Me.dgReferenced.blnMeIsSource = False
        Me.dgReferenced.blnMove = True
        Me.dgReferenced.blnRO = False
        Me.dgReferenced.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgReferenced.Location = New System.Drawing.Point(12, 655)
        Me.dgReferenced.Name = "dgReferenced"
        Me.dgReferenced.RowEnterInterval = 0
        Me.dgReferenced.RowIndex = -1
        Me.dgReferenced.Size = New System.Drawing.Size(300, 105)
        Me.dgReferenced.ta = Nothing
        Me.dgReferenced.TabIndex = 111
        '
        'dgReferencing
        '
        Me.dgReferencing.blnDirty = False
        Me.dgReferencing.blnMeIsSource = False
        Me.dgReferencing.blnMove = True
        Me.dgReferencing.blnRO = False
        Me.dgReferencing.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgReferencing.Location = New System.Drawing.Point(406, 655)
        Me.dgReferencing.Name = "dgReferencing"
        Me.dgReferencing.RowEnterInterval = 0
        Me.dgReferencing.RowIndex = -1
        Me.dgReferencing.Size = New System.Drawing.Size(300, 105)
        Me.dgReferencing.ta = Nothing
        Me.dgReferencing.TabIndex = 112
        '
        'TableAdapterManager
        '
        Me.TableAdapterManager.BackupDataSetBeforeUpdate = False
        Me.TableAdapterManager.Connection = Nothing
        Me.TableAdapterManager.m_keywordsTableAdapter = Nothing
        Me.TableAdapterManager.m_sql_charactersTableAdapter = Nothing
        Me.TableAdapterManager.UpdateOrder = MetaDataSetTableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete
        '
        'ROUTINESTableAdapter
        '
        Me.ROUTINESTableAdapter.ClearBeforeFill = True
        '
        'Dm_sql_referenced_entitiesBindingSource
        '
        Me.Dm_sql_referenced_entitiesBindingSource.DataMember = "dm_sql_referenced_entities"
        Me.Dm_sql_referenced_entitiesBindingSource.DataSource = Me.MetaData
        '
        'Dm_sql_referenced_entitiesTableAdapter
        '
        Me.Dm_sql_referenced_entitiesTableAdapter.ClearBeforeFill = True
        '
        'Dm_sql_referencing_entitiesBindingSource
        '
        Me.Dm_sql_referencing_entitiesBindingSource.DataMember = "dm_sql_referencing_entities"
        Me.Dm_sql_referencing_entitiesBindingSource.DataSource = Me.MetaData
        '
        'Dm_sql_referencing_entitiesTableAdapter
        '
        Me.Dm_sql_referencing_entitiesTableAdapter.ClearBeforeFill = True
        '
        'bProcCopy
        '
        Me.bProcCopy.Location = New System.Drawing.Point(330, 32)
        Me.bProcCopy.Name = "bProcCopy"
        Me.bProcCopy.Size = New System.Drawing.Size(75, 23)
        Me.bProcCopy.TabIndex = 114
        Me.bProcCopy.Text = "Copy"
        Me.bProcCopy.UseVisualStyleBackColor = True
        '
        'frmMetaProc
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = False
        Me.ClientSize = New System.Drawing.Size(1605, 931)
        Me.Controls.Add(Me.bProcCopy)
        Me.Controls.Add(Me.dgReferencing)
        Me.Controls.Add(Me.dgReferenced)
        Me.Controls.Add(Me.tbROUTINES)
        Me.Controls.Add(Me.dgROUTINES)
        Me.Location = New System.Drawing.Point(0, 0)
        Me.Name = "frmMetaProc"
        Me.Text = "frmMetaView"
        Me.Controls.SetChildIndex(Me.HelpTextBox, 0)
        Me.Controls.SetChildIndex(Me.dgROUTINES, 0)
        Me.Controls.SetChildIndex(Me.tbROUTINES, 0)
        Me.Controls.SetChildIndex(Me.dgReferenced, 0)
        Me.Controls.SetChildIndex(Me.dgReferencing, 0)
        Me.Controls.SetChildIndex(Me.bProcCopy, 0)
        CType(Me.MetaData, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgROUTINES, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ROUTINESBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgReferenced, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgReferencing, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Dm_sql_referenced_entitiesBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Dm_sql_referencing_entitiesBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MetaData As MetaDataSet
    Friend WithEvents TableAdapterManager As MetaDataSetTableAdapters.TableAdapterManager
    Friend WithEvents ROUTINESBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents ROUTINESTableAdapter As MetaDataSetTableAdapters.ROUTINESTableAdapter
    Friend WithEvents dgROUTINES As dgvEnter
    Friend WithEvents tbROUTINES As System.Windows.Forms.RichTextBox
    Friend WithEvents Dm_sql_referenced_entitiesBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents Dm_sql_referenced_entitiesTableAdapter As MetaDataSetTableAdapters.dm_sql_referenced_entitiesTableAdapter
    Friend WithEvents dgReferenced As dgvEnter
    Friend WithEvents Dm_sql_referencing_entitiesBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents Dm_sql_referencing_entitiesTableAdapter As MetaDataSetTableAdapters.dm_sql_referencing_entitiesTableAdapter
    Friend WithEvents dgReferencing As dgvEnter
    Friend WithEvents bProcCopy As System.Windows.Forms.Button
End Class
