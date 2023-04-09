'------------------------------------------------
'Name: Module gen_ADHOCDataSet_v_adhoc_view_columns.vb.
'Function: 
'Copyright Robin Baines 2008. All rights reserved.
'Created 11-12-2015 0:00:00.
'Notes: 
'Modifications:
'------------------------------------------------
Imports Utilities
Imports System.Windows.Forms
Imports System.Drawing
Public Class MetaData_v_all_views_column
    Inherits dgColumns
    Friend WithEvents dgtable_name As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcolumn_name As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgtable_schema As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgDATA_TYPE As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgCHARACTER_MAXIMUM_LENGTH As System.Windows.Forms.DataGridViewTextBoxColumn
    Public Sub New(ByVal strForm As String, ByVal _bs As BindingSource, ByVal _dg As dgvEnter, _
    ByVal _ta As MetaDataSetTableAdapters.v_all_views_columnTableAdapter, _
    ByVal _ds As DataSet, _
    ByVal _components As System.ComponentModel.Container, _
    ByVal _MainDefs As MainDefinitions, _
    ByVal blnRO As Boolean, _
    ByVal _Controls As Control.ControlCollection, ByVal _frmStandard As frmStandard, _
    ByVal blnFilters As Boolean)
        MyBase.New(strForm, "v_all_views_column", _bs, _dg, _ta, _ds, _MainDefs, blnRO, _
        "", "", _Controls, _frmStandard, blnFilters)
        _ta.Connection.ConnectionString = GetConnectionString()
    End Sub
    Public Overrides Sub Createcolumns()
        dgtable_name = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgtable_schema = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgcolumn_name = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgDATA_TYPE = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgCHARACTER_MAXIMUM_LENGTH = New System.Windows.Forms.DataGridViewTextBoxColumn
    End Sub
    Public Overrides Sub Adjustcolumns(ByVal blnAdjustWidth As Boolean)
        MyBase.Adjustcolumns(blnAdjustWidth)
        Try
            DefineColumn(dgtable_name, "table_name", blnRO, ds.v_all_views_column.table_nameColumn.MaxLength)
            DefineColumn(dgtable_schema, "table_schema", blnRO, ds.v_all_views_column.TABLE_SCHEMAColumn.MaxLength)
            DefineColumn(dgcolumn_name, "column_name", blnRO, ds.v_all_views_column.column_nameColumn.MaxLength)
            DefineColumn(dgDATA_TYPE, "DATA_TYPE", blnRO, ds.v_all_views_column.DATA_TYPEColumn.MaxLength)
            DefineColumn(dgCHARACTER_MAXIMUM_LENGTH, "CHARACTER_MAXIMUM_LENGTH", blnRO, ds.v_all_views_column.CHARACTER_MAXIMUM_LENGTHColumn.MaxLength)
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
        CreateAFilterBox(tbtable_nameFind, "table_name", AddressOf tbFind_TextChanged, _Controls)
        CreateAFilterBox(tbtable_schemaFind, "table_schema", AddressOf tbFind_TextChanged, _Controls)
        CreateAFilterBox(tbcolumn_nameFind, "column_name", AddressOf tbFind_TextChanged, _Controls)
        CreateAFilterBox(tbDATA_TYPEFind, "DATA_TYPE", AddressOf tbFind_TextChanged, _Controls)
        CreateAFilterBox(tbCHARACTER_MAXIMUM_LENGTHFind, "CHARACTER_MAXIMUM_LENGTH", AddressOf tbFind_TextChanged, _Controls)
    End Sub
    Friend WithEvents tbtable_nameFind As System.Windows.Forms.TextBox
    Friend WithEvents tbtable_schemaFind As System.Windows.Forms.TextBox
    Friend WithEvents tbcolumn_nameFind As System.Windows.Forms.TextBox
    Friend WithEvents tbDATA_TYPEFind As System.Windows.Forms.TextBox
    Friend WithEvents tbCHARACTER_MAXIMUM_LENGTHFind As System.Windows.Forms.TextBox
    Private Sub cbFind_CheckChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MakeFilter(False)
    End Sub
    Private Sub tbFind_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MakeFilter(False)
    End Sub
#End Region
End Class
