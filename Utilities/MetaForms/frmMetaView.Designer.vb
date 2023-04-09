<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMetaView
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
        Me.dgParent = New dgvEnter()
        Me.dgChild1 = New dgvEnter()
        Me.tbView = New System.Windows.Forms.RichTextBox()
        Me.v_all_viewsBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.MetaData = New MetaDataSet()
        Me.ROUTINESBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.dgReferenced = New dgvEnter()
        Me.dgReferencing = New dgvEnter()
        Me.TableAdapterManager = New MetaDataSetTableAdapters.TableAdapterManager()
        Me.V_all_views_columnBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.V_all_views_columnTableAdapter = New MetaDataSetTableAdapters.v_all_views_columnTableAdapter()
        Me.v_all_viewsTableAdapter = New MetaDataSetTableAdapters.v_all_viewsTableAdapter()
        Me.Dm_sql_referenced_entitiesBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.Dm_sql_referenced_entitiesTableAdapter = New MetaDataSetTableAdapters.dm_sql_referenced_entitiesTableAdapter()
        Me.Dm_sql_referencing_entitiesBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.Dm_sql_referencing_entitiesTableAdapter = New MetaDataSetTableAdapters.dm_sql_referencing_entitiesTableAdapter()
        Me.bViewCopy = New System.Windows.Forms.Button()
        CType(Me.dgParent, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgChild1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.v_all_viewsBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MetaData, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ROUTINESBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgReferenced, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgReferencing, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.V_all_views_columnBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Dm_sql_referenced_entitiesBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Dm_sql_referencing_entitiesBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgParent
        '
        Me.dgParent.blnDirty = False
        Me.dgParent.blnMeIsSource = False
        Me.dgParent.blnMove = True
        Me.dgParent.blnRO = False
        Me.dgParent.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgParent.Location = New System.Drawing.Point(8, 59)
        Me.dgParent.Name = "dgParent"
        Me.dgParent.RowEnterInterval = 0
        Me.dgParent.RowIndex = -1
        Me.dgParent.Size = New System.Drawing.Size(300, 692)
        Me.dgParent.ta = Nothing
        Me.dgParent.TabIndex = 100
        '
        'dgChild1
        '
        Me.dgChild1.blnDirty = False
        Me.dgChild1.blnMeIsSource = False
        Me.dgChild1.blnMove = True
        Me.dgChild1.blnRO = False
        Me.dgChild1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgChild1.Location = New System.Drawing.Point(330, 61)
        Me.dgChild1.Name = "dgChild1"
        Me.dgChild1.RowEnterInterval = 0
        Me.dgChild1.RowIndex = -1
        Me.dgChild1.Size = New System.Drawing.Size(439, 692)
        Me.dgChild1.ta = Nothing
        Me.dgChild1.TabIndex = 100
        '
        'tbView
        '
        Me.tbView.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.v_all_viewsBindingSource, "VIEW_DEFINITION", True))
        Me.tbView.Font = New System.Drawing.Font("Courier New", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbView.Location = New System.Drawing.Point(785, 61)
        Me.tbView.Name = "tbView"
        Me.tbView.ReadOnly = True
        Me.tbView.Size = New System.Drawing.Size(800, 692)
        Me.tbView.TabIndex = 109
        Me.tbView.Text = ""
        '
        'v_all_viewsBindingSource
        '
        Me.v_all_viewsBindingSource.DataMember = "v_all_views"
        Me.v_all_viewsBindingSource.DataSource = Me.MetaData
        '
        'MetaData
        '
        Me.MetaData.DataSetName = "MetaData"
        Me.MetaData.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
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
        Me.dgReferenced.Location = New System.Drawing.Point(12, 797)
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
        Me.dgReferencing.Location = New System.Drawing.Point(406, 797)
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
        'V_all_views_columnBindingSource
        '
        Me.V_all_views_columnBindingSource.DataMember = "v_all_views_column"
        Me.V_all_views_columnBindingSource.DataSource = Me.MetaData
        '
        'V_all_views_columnTableAdapter
        '
        Me.V_all_views_columnTableAdapter.ClearBeforeFill = True
        '
        'v_all_viewsTableAdapter
        '
        Me.v_all_viewsTableAdapter.ClearBeforeFill = True
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
        'bViewCopy
        '
        Me.bViewCopy.Location = New System.Drawing.Point(785, 36)
        Me.bViewCopy.Name = "bViewCopy"
        Me.bViewCopy.Size = New System.Drawing.Size(75, 23)
        Me.bViewCopy.TabIndex = 113
        Me.bViewCopy.Text = "Copy"
        Me.bViewCopy.UseVisualStyleBackColor = True
        '
        'frmMetaView
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = False
        Me.ClientSize = New System.Drawing.Size(1594, 833)
        Me.Controls.Add(Me.bViewCopy)
        Me.Controls.Add(Me.dgReferencing)
        Me.Controls.Add(Me.dgReferenced)
        Me.Controls.Add(Me.tbView)
        Me.Controls.Add(Me.dgChild1)
        Me.Controls.Add(Me.dgParent)
        Me.Location = New System.Drawing.Point(0, 0)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "frmMetaView"
        Me.Text = "frmMetaView"
        Me.Controls.SetChildIndex(Me.dgParent, 0)
        Me.Controls.SetChildIndex(Me.dgChild1, 0)
        Me.Controls.SetChildIndex(Me.tbView, 0)
        Me.Controls.SetChildIndex(Me.dgReferenced, 0)
        Me.Controls.SetChildIndex(Me.dgReferencing, 0)
        Me.Controls.SetChildIndex(Me.bViewCopy, 0)
        CType(Me.dgParent, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgChild1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.v_all_viewsBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MetaData, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ROUTINESBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgReferenced, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgReferencing, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.V_all_views_columnBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Dm_sql_referenced_entitiesBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Dm_sql_referencing_entitiesBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MetaData As MetaDataSet
    Friend WithEvents TableAdapterManager As MetaDataSetTableAdapters.TableAdapterManager
    Friend WithEvents dgParent As dgvEnter
    Friend WithEvents V_all_views_columnBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents V_all_views_columnTableAdapter As MetaDataSetTableAdapters.v_all_views_columnTableAdapter
    Friend WithEvents dgChild1 As dgvEnter
    Friend WithEvents tbView As System.Windows.Forms.RichTextBox
    Friend WithEvents v_all_viewsBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents v_all_viewsTableAdapter As MetaDataSetTableAdapters.v_all_viewsTableAdapter
    Friend WithEvents ROUTINESBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents Dm_sql_referenced_entitiesBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents Dm_sql_referenced_entitiesTableAdapter As MetaDataSetTableAdapters.dm_sql_referenced_entitiesTableAdapter
    Friend WithEvents dgReferenced As dgvEnter
    Friend WithEvents Dm_sql_referencing_entitiesBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents Dm_sql_referencing_entitiesTableAdapter As MetaDataSetTableAdapters.dm_sql_referencing_entitiesTableAdapter
    Friend WithEvents dgReferencing As dgvEnter
    Friend WithEvents bViewCopy As System.Windows.Forms.Button
End Class
