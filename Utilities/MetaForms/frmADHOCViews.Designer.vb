<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAdhocViews
    Inherits frmStandard



    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.MetaDataSet = New MetaDataSet()
        Me.V_adhoc_viewsBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.V_adhoc_viewsTableAdapter = New MetaDataSetTableAdapters.v_adhoc_viewsTableAdapter()
        Me.TableAdapterManager = New MetaDataSetTableAdapters.TableAdapterManager()
        Me.V_adhoc_view_columnsBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.V_adhoc_view_columnsTableAdapter = New MetaDataSetTableAdapters.v_adhoc_view_columnsTableAdapter()
        Me.lbColumns = New System.Windows.Forms.ListBox()
        Me.tbWhere = New System.Windows.Forms.TextBox()
        Me.btnShowData = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lbAdhocViews = New System.Windows.Forms.ListBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        CType(Me.MetaDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.V_adhoc_viewsBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.V_adhoc_view_columnsBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ADHOCDataSet
        '
        Me.MetaDataSet.DataSetName = "MetaDataSet"
        Me.MetaDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'V_adhoc_viewsBindingSource
        '
        Me.V_adhoc_viewsBindingSource.DataMember = "v_adhoc_views"
        Me.V_adhoc_viewsBindingSource.DataSource = Me.MetaDataSet
        '
        'V_adhoc_viewsTableAdapter
        '
        Me.V_adhoc_viewsTableAdapter.ClearBeforeFill = True
        '
        'TableAdapterManager
        '
        Me.TableAdapterManager.BackupDataSetBeforeUpdate = False
        Me.TableAdapterManager.Connection = Nothing
        Me.TableAdapterManager.UpdateOrder = MetaDataSetTableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete
        '
        'V_adhoc_view_columnsBindingSource
        '
        Me.V_adhoc_view_columnsBindingSource.DataMember = "v_adhoc_view_columns"
        Me.V_adhoc_view_columnsBindingSource.DataSource = Me.MetaDataSet
        '
        'V_adhoc_view_columnsTableAdapter
        '
        Me.V_adhoc_view_columnsTableAdapter.ClearBeforeFill = True
        '
        'lbColumns
        '
        Me.lbColumns.DataSource = Me.V_adhoc_view_columnsBindingSource
        Me.lbColumns.DisplayMember = "column_name"
        Me.lbColumns.FormattingEnabled = True
        Me.lbColumns.Location = New System.Drawing.Point(351, 60)
        Me.lbColumns.Name = "lbColumns"
        Me.lbColumns.Size = New System.Drawing.Size(293, 420)
        Me.lbColumns.TabIndex = 101
        Me.lbColumns.ValueMember = "column_name"
        '
        'tbWhere
        '
        Me.tbWhere.Location = New System.Drawing.Point(97, 501)
        Me.tbWhere.Name = "tbWhere"
        Me.tbWhere.Size = New System.Drawing.Size(856, 20)
        Me.tbWhere.TabIndex = 102
        '
        'btnShowData
        '
        Me.btnShowData.Location = New System.Drawing.Point(43, 527)
        Me.btnShowData.Name = "btnShowData"
        Me.btnShowData.Size = New System.Drawing.Size(130, 23)
        Me.btnShowData.TabIndex = 103
        Me.btnShowData.Text = "Show Data"
        Me.btnShowData.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(348, 41)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(135, 13)
        Me.Label1.TabIndex = 104
        Me.Label1.Text = "Fields of the selected view."
        '
        'lbAdhocViews
        '
        Me.lbAdhocViews.DataSource = Me.V_adhoc_viewsBindingSource
        Me.lbAdhocViews.DisplayMember = "TABLE_NAME"
        Me.lbAdhocViews.FormattingEnabled = True
        Me.lbAdhocViews.Location = New System.Drawing.Point(39, 60)
        Me.lbAdhocViews.Name = "lbAdhocViews"
        Me.lbAdhocViews.Size = New System.Drawing.Size(293, 420)
        Me.lbAdhocViews.TabIndex = 105
        Me.lbAdhocViews.ValueMember = "TABLE_NAME"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(36, 41)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(71, 13)
        Me.Label2.TabIndex = 106
        Me.Label2.Text = "Adhoc views."
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(40, 504)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(51, 13)
        Me.Label3.TabIndex = 107
        Me.Label3.Text = "WHERE "
        '
        'TextBox1
        '
        Me.TextBox1.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.V_adhoc_viewsBindingSource, "VIEW_DEFINITION", True))
        Me.TextBox1.Location = New System.Drawing.Point(662, 60)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.TextBox1.Size = New System.Drawing.Size(800, 420)
        Me.TextBox1.TabIndex = 108
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(659, 41)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(54, 13)
        Me.Label4.TabIndex = 109
        Me.Label4.Text = "Definition."
        '
        'frmAdhocViews
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = False
        Me.ClientSize = New System.Drawing.Size(1486, 796)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lbAdhocViews)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnShowData)
        Me.Controls.Add(Me.tbWhere)
        Me.Controls.Add(Me.lbColumns)
        Me.Location = New System.Drawing.Point(0, 0)
        Me.Name = "frmAdhocViews"
        Me.Text = "frmADHOCViews"
        Me.Controls.SetChildIndex(Me.lbColumns, 0)
        Me.Controls.SetChildIndex(Me.tbWhere, 0)
        Me.Controls.SetChildIndex(Me.btnShowData, 0)
        Me.Controls.SetChildIndex(Me.Label1, 0)
        Me.Controls.SetChildIndex(Me.lbAdhocViews, 0)
        Me.Controls.SetChildIndex(Me.Label2, 0)
        Me.Controls.SetChildIndex(Me.Label3, 0)
        Me.Controls.SetChildIndex(Me.TextBox1, 0)
        Me.Controls.SetChildIndex(Me.Label4, 0)
        CType(Me.MetaDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.V_adhoc_viewsBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.V_adhoc_view_columnsBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MetaDataSet As MetaDataSet
    Friend WithEvents V_adhoc_viewsBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents V_adhoc_viewsTableAdapter As MetaDataSetTableAdapters.v_adhoc_viewsTableAdapter
    Friend WithEvents TableAdapterManager As MetaDataSetTableAdapters.TableAdapterManager
    Friend WithEvents V_adhoc_view_columnsBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents V_adhoc_view_columnsTableAdapter As MetaDataSetTableAdapters.v_adhoc_view_columnsTableAdapter
    Friend WithEvents lbColumns As System.Windows.Forms.ListBox
    Friend WithEvents tbWhere As System.Windows.Forms.TextBox
    Friend WithEvents btnShowData As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lbAdhocViews As System.Windows.Forms.ListBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
End Class
