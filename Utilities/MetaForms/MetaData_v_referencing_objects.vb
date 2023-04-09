'------------------------------------------------
'Name: Module gen_MetaData_v_referencing_objects.vb.
'Function: 
'Copyright Robin Baines 2008. All rights reserved.
'Created 12-12-2015 0:00:00.
'Notes: 
'Modifications:dm_sql_referencing_entitiesTableAdapter
'------------------------------------------------
Imports Utilities
Imports System.Windows.Forms
Imports System.Drawing
Public Class MetaData_v_referencing_objects
Inherits dgColumns
Friend WithEvents dgreferencing_schema_name As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgreferencing_entity_name As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgreferencing_id As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgreferencing_class_desc As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgis_caller_dependent As System.Windows.Forms.DataGridViewCheckBoxColumn
    Public Sub New(ByVal strForm As String, ByVal _bs As BindingSource, ByVal _dg As dgvEnter, _
    ByVal _ta As MetaDataSetTableAdapters.dm_sql_referencing_entitiesTableAdapter, _
    ByVal _ds As DataSet, _
    ByVal _components As System.ComponentModel.Container, _
    ByVal _MainDefs As MainDefinitions, _
    ByVal blnRO As Boolean, _
    ByVal _Controls As Control.ControlCollection, ByVal _frmStandard As frmStandard, _
    ByVal blnFilters As Boolean)
        MyBase.New(strForm, "Dm_sql_referencing_entities", _bs, _dg, _ta, _ds, _MainDefs, blnRO, _
        "", "", _Controls, _frmStandard, blnFilters)
        _ta.Connection.ConnectionString = GetConnectionString()
    End Sub
    Public Overrides Sub Createcolumns()
        dgreferencing_schema_name = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgreferencing_entity_name = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgreferencing_id = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgreferencing_class_desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgis_caller_dependent = New System.Windows.Forms.DataGridViewCheckBoxColumn
    End Sub
    Public Overrides Sub Adjustcolumns(ByVal blnAdjustWidth As Boolean)
        MyBase.Adjustcolumns(blnAdjustWidth)
        Try
            DefineColumn(dgreferencing_schema_name, "referencing_schema_name", blnRO, ds.Dm_sql_referencing_entities.referencing_schema_nameColumn.MaxLength)
            DefineColumn(dgreferencing_entity_name, "referencing_entity_name", blnRO, ds.Dm_sql_referencing_entities.referencing_entity_nameColumn.MaxLength)
            DefineColumn(dgreferencing_id, "referencing_id", blnRO, ds.Dm_sql_referencing_entities.referencing_idColumn.MaxLength)
            DefineColumn(dgreferencing_class_desc, "referencing_class_desc", blnRO, ds.Dm_sql_referencing_entities.referencing_class_descColumn.MaxLength)
            DefineColumn(dgis_caller_dependent, "is_caller_dependent", blnRO, ds.Dm_sql_referencing_entities.is_caller_dependentColumn.MaxLength)
            PutColumnsInGrid()
            AdjustDataGridWidth(blnAdjustWidth)
            RefreshCombos()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
Public Overrides Sub RefreshCombos()
MyBase.RefreshCombos()
dg.CancelEdit()
iComboCount = 0
End sub
#Region "Filter"
Public Overrides Sub CreateFilterBoxes(ByVal _Controls As Control.ControlCollection)
MyBase.CreateFilterBoxes(_Controls)
CreateAFilterBox(tbreferencing_schema_nameFind, "referencing_schema_name", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbreferencing_entity_nameFind, "referencing_entity_name", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbreferencing_idFind, "referencing_id", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbreferencing_class_descFind, "referencing_class_desc", AddressOf tbFind_TextChanged, _Controls)
CreateACheckBox(cbis_caller_dependentFind, "is_caller_dependent", AddressOf cbFind_CheckChanged, _Controls)
End Sub
Friend WithEvents tbreferencing_schema_nameFind As System.Windows.Forms.TextBox
Friend WithEvents tbreferencing_entity_nameFind As System.Windows.Forms.TextBox
Friend WithEvents tbreferencing_idFind As System.Windows.Forms.TextBox
Friend WithEvents tbreferencing_class_descFind As System.Windows.Forms.TextBox
Friend WithEvents cbis_caller_dependentFind As System.Windows.Forms.CheckBox
Private Sub cbFind_CheckChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
MakeFilter(False)
End Sub
Private Sub tbFind_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
 MakeFilter(False)
End Sub
#End Region
End Class
