'------------------------------------------------
'Name: Module gen_MetaData_v_referenced_objects.vb.
'Function: 
'Copyright Robin Baines 2008. All rights reserved.
'Created 12-12-2015 0:00:00.
'Notes: 
'Modifications:dm_sql_referenced_entitiesTableAdapter
'------------------------------------------------
Imports Utilities
Imports System.Windows.Forms
Imports System.Drawing
Public Class MetaData_v_referenced_objects
    Inherits dgColumns
    Friend WithEvents dgreferenced_server_name As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgreferenced_schema_name As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgreferenced_entity_name As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgreferenced_id As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgreferenced_class As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgreferenced_class_desc As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgis_caller_dependent As System.Windows.Forms.DataGridViewCheckBoxColumn
Friend WithEvents dgis_ambiguous As System.Windows.Forms.DataGridViewCheckBoxColumn
Friend WithEvents dgis_selected As System.Windows.Forms.DataGridViewCheckBoxColumn
Friend WithEvents dgis_updated As System.Windows.Forms.DataGridViewCheckBoxColumn
Friend WithEvents dgis_select_all As System.Windows.Forms.DataGridViewCheckBoxColumn
Friend WithEvents dgis_all_columns_found As System.Windows.Forms.DataGridViewCheckBoxColumn
    Public Sub New(ByVal strForm As String, ByVal _bs As BindingSource, ByVal _dg As dgvEnter, _
    ByVal _ta As MetaDataSetTableAdapters.dm_sql_referenced_entitiesTableAdapter, _
    ByVal _ds As DataSet, _
    ByVal _components As System.ComponentModel.Container, _
    ByVal _MainDefs As MainDefinitions, _
    ByVal blnRO As Boolean, _
    ByVal _Controls As Control.ControlCollection, ByVal _frmStandard As frmStandard, _
    ByVal blnFilters As Boolean)
        MyBase.New(strForm, "Dm_sql_referenced_entities", _bs, _dg, _ta, _ds, _MainDefs, blnRO, _
        "", "", _Controls, _frmStandard, blnFilters)
        _ta.Connection.ConnectionString = GetConnectionString()
    End Sub
    Public Overrides Sub Createcolumns()
        dgreferenced_server_name = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgreferenced_schema_name = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgreferenced_entity_name = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgreferenced_id = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgreferenced_class = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgreferenced_class_desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgis_caller_dependent = New System.Windows.Forms.DataGridViewCheckBoxColumn
        dgis_ambiguous = New System.Windows.Forms.DataGridViewCheckBoxColumn
        dgis_selected = New System.Windows.Forms.DataGridViewCheckBoxColumn
        dgis_updated = New System.Windows.Forms.DataGridViewCheckBoxColumn
        dgis_select_all = New System.Windows.Forms.DataGridViewCheckBoxColumn
        dgis_all_columns_found = New System.Windows.Forms.DataGridViewCheckBoxColumn
    End Sub
    Public Overrides Sub Adjustcolumns(ByVal blnAdjustWidth As Boolean)
        MyBase.Adjustcolumns(blnAdjustWidth)
        Try
            DefineColumn(dgreferenced_server_name, "referenced_server_name", blnRO, ds.v_referenced_objects.referenced_server_nameColumn.MaxLength)
            DefineColumn(dgreferenced_schema_name, "referenced_schema_name", blnRO, ds.v_referenced_objects.referenced_schema_nameColumn.MaxLength)
            DefineColumn(dgreferenced_entity_name, "referenced_entity_name", blnRO, ds.Dm_sql_referenced_entities.referenced_entity_nameColumn.MaxLength)
            DefineColumn(dgreferenced_id, "referenced_id", blnRO, ds.Dm_sql_referenced_entities.referenced_idColumn.MaxLength)
            DefineColumn(dgreferenced_class, "referenced_class", blnRO, ds.Dm_sql_referenced_entities.referenced_classColumn.MaxLength)
            DefineColumn(dgreferenced_class_desc, "referenced_class_desc", blnRO, ds.Dm_sql_referenced_entities.referenced_class_descColumn.MaxLength)
            DefineColumn(dgis_caller_dependent, "is_caller_dependent", blnRO, ds.Dm_sql_referenced_entities.is_caller_dependentColumn.MaxLength)
            DefineColumn(dgis_ambiguous, "is_ambiguous", blnRO, ds.Dm_sql_referenced_entities.is_ambiguousColumn.MaxLength)
            DefineColumn(dgis_selected, "is_selected", blnRO, ds.Dm_sql_referenced_entities.is_selectedColumn.MaxLength)
            DefineColumn(dgis_updated, "is_updated", blnRO, ds.Dm_sql_referenced_entities.is_updatedColumn.MaxLength)
            DefineColumn(dgis_select_all, "is_select_all", blnRO, ds.Dm_sql_referenced_entities.is_select_allColumn.MaxLength)
            DefineColumn(dgis_all_columns_found, "is_all_columns_found", blnRO, ds.Dm_sql_referenced_entities.is_all_columns_foundColumn.MaxLength)
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
        CreateAFilterBox(tbreferenced_server_nameFind, "referenced_server_name", AddressOf tbFind_TextChanged, _Controls)
        CreateAFilterBox(tbreferenced_schema_nameFind, "referenced_schema_name", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbreferenced_entity_nameFind, "referenced_entity_name", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbreferenced_idFind, "referenced_id", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbreferenced_classFind, "referenced_class", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbreferenced_class_descFind, "referenced_class_desc", AddressOf tbFind_TextChanged, _Controls)
CreateACheckBox(cbis_caller_dependentFind, "is_caller_dependent", AddressOf cbFind_CheckChanged, _Controls)
CreateACheckBox(cbis_ambiguousFind, "is_ambiguous", AddressOf cbFind_CheckChanged, _Controls)
CreateACheckBox(cbis_selectedFind, "is_selected", AddressOf cbFind_CheckChanged, _Controls)
CreateACheckBox(cbis_updatedFind, "is_updated", AddressOf cbFind_CheckChanged, _Controls)
CreateACheckBox(cbis_select_allFind, "is_select_all", AddressOf cbFind_CheckChanged, _Controls)
CreateACheckBox(cbis_all_columns_foundFind, "is_all_columns_found", AddressOf cbFind_CheckChanged, _Controls)
    End Sub
    Friend WithEvents tbreferenced_server_nameFind As System.Windows.Forms.TextBox
    Friend WithEvents tbreferenced_schema_nameFind As System.Windows.Forms.TextBox
Friend WithEvents tbreferenced_entity_nameFind As System.Windows.Forms.TextBox
Friend WithEvents tbreferenced_idFind As System.Windows.Forms.TextBox
Friend WithEvents tbreferenced_classFind As System.Windows.Forms.TextBox
Friend WithEvents tbreferenced_class_descFind As System.Windows.Forms.TextBox
Friend WithEvents cbis_caller_dependentFind As System.Windows.Forms.CheckBox
Friend WithEvents cbis_ambiguousFind As System.Windows.Forms.CheckBox
Friend WithEvents cbis_selectedFind As System.Windows.Forms.CheckBox
Friend WithEvents cbis_updatedFind As System.Windows.Forms.CheckBox
Friend WithEvents cbis_select_allFind As System.Windows.Forms.CheckBox
Friend WithEvents cbis_all_columns_foundFind As System.Windows.Forms.CheckBox
Private Sub cbFind_CheckChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
MakeFilter(False)
End Sub
Private Sub tbFind_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
 MakeFilter(False)
End Sub
#End Region
End Class
