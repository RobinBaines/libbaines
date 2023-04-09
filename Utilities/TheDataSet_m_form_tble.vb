'------------------------------------------------
'Name: Module gen_TheDataSet_m_form_tble.vb.
'Function: 
'Copyright Robin Baines 2008. All rights reserved.
'Created 7/8/2012 12:00:00 AM.
'Notes: 
'Modifications: combo -> textbox
'------------------------------------------------
Imports Utilities
Imports System.Windows.Forms
Imports System.Drawing
Public Class TheDataSet_m_form_tble
Inherits dgColumns
    Friend WithEvents dgform As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgtble As System.Windows.Forms.DataGridViewTextBoxColumn
    'Public bsm_form As BindingSource
    'Friend WithEvents m_formTableAdapter As TheDataSetTableAdapters.m_formTableAdapter
    'Public bsm_tble As BindingSource
    'Friend WithEvents m_tbleTableAdapter As TheDataSetTableAdapters.m_tbleTableAdapter
Public Sub New(ByVal strForm As String, ByVal _bs As BindingSource, ByVal _dg As dgvEnter, _
ByVal _ta As TheDataSetTableAdapters.m_form_tbleTableAdapter, _
ByVal _ds As DataSet, _
ByVal _components As System.ComponentModel.Container, _
ByVal _MainDefs As MainDefinitions, _
ByVal blnRO As Boolean, _
ByVal _Controls As Control.ControlCollection, ByVal _frmStandard As frmStandard, _
ByVal blnFilters As Boolean)
MyBase.New(strForm, "m_form_tble", _bs, _dg, _ta, _ds, _MainDefs, blnRO, _
"form","tble",_Controls, _frmStandard, blnFilters)
_ta.Connection.ConnectionString = GetConnectionString()
        'Me.bsm_form = New System.Windows.Forms.BindingSource(_components)
        'm_formTableAdapter = New TheDataSetTableAdapters.m_formTableAdapter
        'm_formTableAdapter.Connection.ConnectionString = GetConnectionString()
        'Me.bsm_form.DataMember = "m_form"
        'Me.bsm_form.DataSource = ds
        'Me.bsm_tble = New System.Windows.Forms.BindingSource(_components)
        'm_tbleTableAdapter = New TheDataSetTableAdapters.m_tbleTableAdapter
        'm_tbleTableAdapter.Connection.ConnectionString = GetConnectionString()
        'Me.bsm_tble.DataMember = "m_tble"
        'Me.bsm_tble.DataSource = ds
End Sub
Public Overrides Sub Createcolumns()
        dgform = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgtble = New System.Windows.Forms.DataGridViewTextBoxColumn
End Sub
Public Overrides Sub Adjustcolumns(ByVal blnAdjustWidth As Boolean)
 MyBase.Adjustcolumns(blnAdjustWidth)
        'DefineComboBoxColumn(dgform, MainDefs.strGetFormat("TYP_M_STRING"), True, "form", "", FieldWidths.GENWIDTH, blnRO, true, "",  bsm_form, "form" ,"form", Color.Lavender)
        'If blnRO = True Then dgform.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
        'DefineComboBoxColumn(dgtble, MainDefs.strGetFormat("TYP_M_STRING"), True, "tble", "", FieldWidths.GENWIDTH, blnRO, true, "",  bsm_tble, "tble" ,"tble", Color.Lavender)
        '        If blnRO = True Then dgtble.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing

        DefineColumn(dgform, "form", True, ds.m_form_tble.formColumn.MaxLength)
        DefineColumn(dgtble, "tble", True, ds.m_form_tble.tbleColumn.MaxLength)

PutColumnsInGrid()
AdjustDataGridWidth(blnAdjustWidth)
RefreshCombos()
End Sub
Public Overrides Sub RefreshCombos()
MyBase.RefreshCombos()
        'Me.m_formTableAdapter.Fill(Me.ds.m_form)
        '  Me.m_tbleTableAdapter.Fill(Me.ds.m_tble)
dg.CancelEdit()
        iComboCount = 0
End sub
#Region "Editing"
Public Overrides Sub dg_UserDeletingRow(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs)
Try
Dim tadap As TheDataSetTableAdapters.m_form_tbleTableAdapter
tadap = CType(ta, TheDataSetTableAdapters.m_form_tbleTableAdapter)
tadap.Delete(e.Row.Cells(dg.Columns("form").Index).Value.ToString(),e.Row.Cells(dg.Columns("tble").Index).Value.ToString())
MyBase.dg_UserDeletingRow(sender, e)
Catch ex As Exception
MsgBox("Delete failed. Most common cause is that record is in use in another table." + ex.message)
e.Cancel = True
End Try
End Sub
Private Sub dg_DefaultValuesNeeded(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dg.DefaultValuesNeeded
End Sub
#End Region
#Region "Filter"
Public Overrides Sub CreateFilterBoxes(ByVal _Controls As Control.ControlCollection)
MyBase.CreateFilterBoxes(_Controls)
CreateAFilterBox(tbformFind, "form", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbtbleFind, "tble", AddressOf tbFind_TextChanged, _Controls)
End Sub
Friend WithEvents tbformFind As System.Windows.Forms.TextBox
Friend WithEvents tbtbleFind As System.Windows.Forms.TextBox
Private Sub cbFind_CheckChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
MakeFilter(False)
End Sub
Private Sub tbFind_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
 MakeFilter(False)
End Sub
#End Region
End Class
