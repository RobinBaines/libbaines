'------------------------------------------------
'Name: Module gen_TheDataSet_m_form_tble_column__visibility.vb.
'Function: 
'Copyright Robin Baines 2008. All rights reserved.
'Created 7/8/2012 12:00:00 AM.
'Notes: 
'Modifications:dgform, dgtble and dgcolmn changed from blnRO to true and changed dgcolmn to Text box instead of Combo box.
'Added default_filter
'------------------------------------------------
Imports Utilities
Imports System.Windows.Forms
Imports System.Drawing
Public Class TheDataSet_m_form_tble_column__visibility
Inherits dgColumns
    Friend WithEvents dgform As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgtble As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgcolmn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgvisible As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents dgprnt As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents dgsequence As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgbold As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents dgdefault_filter As System.Windows.Forms.DataGridViewTextBoxColumn
    'Public bsm_form As BindingSource
    'Friend WithEvents m_formTableAdapter As TheDataSetTableAdapters.m_formTableAdapter
    'Public bsm_tble_column As BindingSource
    'Friend WithEvents m_tble_columnTableAdapter As TheDataSetTableAdapters.m_tble_columnTableAdapter
    Public Sub New(ByVal strForm As String, ByVal _bs As BindingSource, ByVal _dg As dgvEnter, _
    ByVal _ta As TheDataSetTableAdapters.m_form_tble_column__visibilityTableAdapter, _
    ByVal _ds As DataSet, _
    ByVal _components As System.ComponentModel.Container, _
    ByVal _MainDefs As MainDefinitions, _
    ByVal blnRO As Boolean, _
    ByVal _Controls As Control.ControlCollection, ByVal _frmStandard As frmStandard, _
    ByVal blnFilters As Boolean)
        MyBase.New(strForm, "m_form_tble_column__visibility", _bs, _dg, _ta, _ds, _MainDefs, blnRO, _
        "form", "tble", _Controls, _frmStandard, blnFilters)
        _ta.Connection.ConnectionString = GetConnectionString()
        'Me.bsm_form = New System.Windows.Forms.BindingSource(_components)
        'm_formTableAdapter = New TheDataSetTableAdapters.m_formTableAdapter
        'm_formTableAdapter.Connection.ConnectionString = GetConnectionString()
        'Me.bsm_form.DataMember = "m_form"
        'Me.bsm_form.DataSource = ds
        'Me.bsm_tble_column = New System.Windows.Forms.BindingSource(_components)
        'm_tble_columnTableAdapter = New TheDataSetTableAdapters.m_tble_columnTableAdapter
        'm_tble_columnTableAdapter.Connection.ConnectionString = GetConnectionString()
        'Me.bsm_tble_column.DataMember = "m_tble_column"
        'Me.bsm_tble_column.DataSource = ds
    End Sub
    Public Overrides Sub Createcolumns()
        dgform = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgtble = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgcolmn = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgvisible = New System.Windows.Forms.DataGridViewCheckBoxColumn
        dgprnt = New System.Windows.Forms.DataGridViewCheckBoxColumn
        dgsequence = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgbold = New System.Windows.Forms.DataGridViewCheckBoxColumn
        dgdefault_filter = New System.Windows.Forms.DataGridViewTextBoxColumn
    End Sub
    Public Overrides Sub Adjustcolumns(ByVal blnAdjustWidth As Boolean)
        MyBase.Adjustcolumns(blnAdjustWidth)
        '        DefineComboBoxColumn(dgform, MainDefs.strGetFormat("TYP_M_STRING"), True, "form", "", FieldWidths.GENWIDTH, True, True, "", bsm_form, "form", "form", Color.Lavender)
        'If blnRO = True Then dgform.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
        '        DefineComboBoxColumn(dgtble, MainDefs.strGetFormat("TYP_M_STRING"), True, "tble", "", FieldWidths.GENWIDTH, True, True, "", bsm_tble_column, "tble", "tble", Color.Lavender)
        '        If blnRO = True Then dgtble.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing


        DefineColumn(dgform, "form", True, ds.m_form_tble_column__visibility.formColumn.MaxLength)
        DefineColumn(dgtble, "tble", True, ds.m_form_tble_column__visibility.tbleColumn.MaxLength)

        'DefineComboBoxColumn(dgcolmn, MainDefs.strGetFormat("TYP_M_STRING"), True, "colmn", "", FieldWidths.GENWIDTH, blnRO, true, "",  bsm_tble_column, "colmn" ,"colmn", Color.Lavender)
        '        If blnRO = True Then dgcolmn.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
        DefineColumn(dgcolmn, "colmn", True, ds.m_form_tble_column__visibility.colmnColumn.MaxLength)
        DefineColumn(dgvisible, "visible", blnRO, ds.m_form_tble_column__visibility.visibleColumn.MaxLength)
        DefineColumn(dgprnt, "prnt", blnRO, ds.m_form_tble_column__visibility.prntColumn.MaxLength)
        DefineColumn(dgsequence, "sequence", blnRO, ds.m_form_tble_column__visibility.sequenceColumn.MaxLength)
        DefineColumn(dgbold, "bold", blnRO, ds.m_form_tble_column__visibility.boldColumn.MaxLength)
        DefineColumn(dgdefault_filter, "default_filter", blnRO, ds.m_form_tble_column__visibility.default_filterColumn.MaxLength)
        PutColumnsInGrid()
        AdjustDataGridWidth(blnAdjustWidth)
        RefreshCombos()
    End Sub
    Public Overrides Sub RefreshCombos()
        MyBase.RefreshCombos()
        'Me.m_formTableAdapter.Fill(Me.ds.m_form)
        '    Me.m_tble_columnTableAdapter.Fill(Me.ds.m_tble_column)
        dg.CancelEdit()
        iComboCount = 0
    End Sub
#Region "Editing"
Public Overrides Sub dg_UserDeletingRow(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs)
Try
Dim tadap As TheDataSetTableAdapters.m_form_tble_column__visibilityTableAdapter
tadap = CType(ta, TheDataSetTableAdapters.m_form_tble_column__visibilityTableAdapter)
tadap.Delete(e.Row.Cells(dg.Columns("form").Index).Value.ToString(),e.Row.Cells(dg.Columns("tble").Index).Value.ToString(),e.Row.Cells(dg.Columns("colmn").Index).Value.ToString())
MyBase.dg_UserDeletingRow(sender, e)
Catch ex As Exception
MsgBox("Delete failed. Most common cause is that record is in use in another table." + ex.message)
e.Cancel = True
End Try
End Sub
Private Sub dg_DefaultValuesNeeded(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dg.DefaultValuesNeeded
With e.Row
            '.Cells("visible").Value =  ((1))
            '.Cells("prnt").Value =  ((1))
            '.Cells("bold").Value =  ((0))
            '.Cells("default_filter").Value =  ('')
End With
End Sub
#End Region
#Region "Filter"
Public Overrides Sub CreateFilterBoxes(ByVal _Controls As Control.ControlCollection)
MyBase.CreateFilterBoxes(_Controls)
CreateAFilterBox(tbformFind, "form", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbtbleFind, "tble", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbcolmnFind, "colmn", AddressOf tbFind_TextChanged, _Controls)
CreateACheckBox(cbvisibleFind, "visible", AddressOf cbFind_CheckChanged, _Controls)
CreateACheckBox(cbprntFind, "prnt", AddressOf cbFind_CheckChanged, _Controls)
CreateAFilterBox(tbsequenceFind, "sequence", AddressOf tbFind_TextChanged, _Controls)
CreateACheckBox(cbboldFind, "bold", AddressOf cbFind_CheckChanged, _Controls)
CreateAFilterBox(tbdefault_filterFind, "default_filter", AddressOf tbFind_TextChanged, _Controls)
End Sub
Friend WithEvents tbformFind As System.Windows.Forms.TextBox
Friend WithEvents tbtbleFind As System.Windows.Forms.TextBox
Friend WithEvents tbcolmnFind As System.Windows.Forms.TextBox
Friend WithEvents cbvisibleFind As System.Windows.Forms.CheckBox
Friend WithEvents cbprntFind As System.Windows.Forms.CheckBox
Friend WithEvents tbsequenceFind As System.Windows.Forms.TextBox
Friend WithEvents cbboldFind As System.Windows.Forms.CheckBox
Friend WithEvents tbdefault_filterFind As System.Windows.Forms.TextBox
Private Sub cbFind_CheckChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
MakeFilter(False)
End Sub
Private Sub tbFind_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
 MakeFilter(False)
End Sub
#End Region
End Class
