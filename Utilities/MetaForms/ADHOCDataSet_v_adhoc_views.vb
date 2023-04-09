'------------------------------------------------
'Name: Module gen_ADHOCDataSet_v_adhoc_views.vb.
'Function: 
'Copyright Robin Baines 2008. All rights reserved.
'Created 11-12-2015 0:00:00.
'Notes: 
'Modifications:VIEWSTableAdapter
'------------------------------------------------
Imports Utilities
Imports System.Windows.Forms
Imports System.Drawing
Public Class ADHOCDataSet_v_all_views
    Inherits dgColumns
    Friend WithEvents dgTABLE_CATALOG As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgTABLE_SCHEMA As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgTABLE_NAME As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgVIEW_DEFINITION As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgCHECK_OPTION As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgIS_UPDATABLE As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgCOMPLETE_NAME As System.Windows.Forms.DataGridViewTextBoxColumn
    Public Sub New(ByVal strForm As String, ByVal _bs As BindingSource, ByVal _dg As dgvEnter, _
    ByVal _ta As MetaDataSetTableAdapters.v_all_viewsTableAdapter, _
    ByVal _ds As DataSet, _
    ByVal _components As System.ComponentModel.Container, _
    ByVal _MainDefs As MainDefinitions, _
    ByVal blnRO As Boolean, _
    ByVal _Controls As Control.ControlCollection, ByVal _frmStandard As frmStandard, _
    ByVal blnFilters As Boolean)
        MyBase.New(strForm, "v_all_views", _bs, _dg, _ta, _ds, _MainDefs, blnRO, _
        "", "", _Controls, _frmStandard, blnFilters)
        _ta.Connection.ConnectionString = GetConnectionString()
    End Sub
    Public Overrides Sub Createcolumns()
        dgTABLE_CATALOG = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgTABLE_SCHEMA = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgTABLE_NAME = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgVIEW_DEFINITION = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgCHECK_OPTION = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgIS_UPDATABLE = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgCOMPLETE_NAME = New System.Windows.Forms.DataGridViewTextBoxColumn
    End Sub
    Public Overrides Sub Adjustcolumns(ByVal blnAdjustWidth As Boolean)
        MyBase.Adjustcolumns(blnAdjustWidth)
        Try
            DefineColumn(dgTABLE_CATALOG, "TABLE_CATALOG", blnRO, ds.v_all_views.TABLE_CATALOGColumn.MaxLength)
            DefineColumn(dgTABLE_SCHEMA, "TABLE_SCHEMA", blnRO, ds.v_all_views.TABLE_SCHEMAColumn.MaxLength)
            DefineColumn(dgTABLE_NAME, "TABLE_NAME", blnRO, ds.v_all_views.TABLE_NAMEColumn.MaxLength)
            DefineColumn(dgVIEW_DEFINITION, "VIEW_DEFINITION", blnRO, ds.v_all_views.VIEW_DEFINITIONColumn.MaxLength)
            DefineColumn(dgCHECK_OPTION, "CHECK_OPTION", blnRO, ds.v_all_views.CHECK_OPTIONColumn.MaxLength)
            DefineColumn(dgIS_UPDATABLE, "IS_UPDATABLE", blnRO, ds.v_all_views.IS_UPDATABLEColumn.MaxLength)
            DefineColumn(dgCOMPLETE_NAME, "COMPLETE_NAME", blnRO, ds.v_all_views.COMPLETE_NAMEColumn.MaxLength)

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
    End Sub
#Region "Filter"
    Public Overrides Sub CreateFilterBoxes(ByVal _Controls As Control.ControlCollection)
        MyBase.CreateFilterBoxes(_Controls)
        CreateAFilterBox(tbTABLE_CATALOGFind, "TABLE_CATALOG", AddressOf tbFind_TextChanged, _Controls)
        CreateAFilterBox(tbTABLE_SCHEMAFind, "TABLE_SCHEMA", AddressOf tbFind_TextChanged, _Controls)
        CreateAFilterBox(tbTABLE_NAMEFind, "TABLE_NAME", AddressOf tbFind_TextChanged, _Controls)
        CreateAFilterBox(tbVIEW_DEFINITIONFind, "VIEW_DEFINITION", AddressOf tbFind_TextChanged, _Controls)
        CreateAFilterBox(tbCHECK_OPTIONFind, "CHECK_OPTION", AddressOf tbFind_TextChanged, _Controls)
        CreateAFilterBox(tbIS_UPDATABLEFind, "IS_UPDATABLE", AddressOf tbFind_TextChanged, _Controls)
        CreateAFilterBox(tbCOMPLETE_NAMEFind, "COMPLETE_NAME", AddressOf tbFind_TextChanged, _Controls)
    End Sub
    Friend WithEvents tbTABLE_CATALOGFind As System.Windows.Forms.TextBox
    Friend WithEvents tbTABLE_SCHEMAFind As System.Windows.Forms.TextBox
    Friend WithEvents tbTABLE_NAMEFind As System.Windows.Forms.TextBox
    Friend WithEvents tbVIEW_DEFINITIONFind As System.Windows.Forms.TextBox
    Friend WithEvents tbCHECK_OPTIONFind As System.Windows.Forms.TextBox
    Friend WithEvents tbIS_UPDATABLEFind As System.Windows.Forms.TextBox
    Friend WithEvents tbCOMPLETE_NAMEFind As System.Windows.Forms.TextBox
    Private Sub cbFind_CheckChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MakeFilter(False)
    End Sub
    Private Sub tbFind_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MakeFilter(False)
    End Sub
#End Region
End Class
