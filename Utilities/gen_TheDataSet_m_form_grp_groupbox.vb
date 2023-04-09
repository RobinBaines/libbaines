'------------------------------------------------
'Name: Module gen_TheDataSet_m_form_grp_groupbox.vb.
'Function: 
'Copyright Robin Baines 2008. All rights reserved.
'Created 7/16/2012 12:00:00 AM.
'Notes: 
'Modifications:
'------------------------------------------------
Imports Utilities
Imports System.Windows.Forms
Imports System.Drawing
Public Class TheDataSet_m_form_grp_groupbox
Inherits dgColumns
    Friend WithEvents dggrp As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgform As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dggroupbox As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgRO As System.Windows.Forms.DataGridViewCheckBoxColumn
    'Public bsm_form_grp As BindingSource
    'Friend WithEvents m_form_grpTableAdapter As TheDataSetTableAdapters.m_form_grpTableAdapter
Public Sub New(ByVal strForm As String, ByVal _bs As BindingSource, ByVal _dg As dgvEnter, _
ByVal _ta As TheDataSetTableAdapters.m_form_grp_groupboxTableAdapter, _
ByVal _ds As DataSet, _
ByVal _components As System.ComponentModel.Container, _
ByVal _MainDefs As MainDefinitions, _
ByVal blnRO As Boolean, _
ByVal _Controls As Control.ControlCollection, ByVal _frmStandard As frmStandard, _
ByVal blnFilters As Boolean)
MyBase.New(strForm, "m_form_grp_groupbox", _bs, _dg, _ta, _ds, _MainDefs, blnRO, _
"grp","form",_Controls, _frmStandard, blnFilters)
_ta.Connection.ConnectionString = GetConnectionString()
        'Me.bsm_form_grp = New System.Windows.Forms.BindingSource(_components)
        'm_form_grpTableAdapter = New TheDataSetTableAdapters.m_form_grpTableAdapter
        'm_form_grpTableAdapter.Connection.ConnectionString = GetConnectionString()
        'Me.bsm_form_grp.DataMember = "m_form_grp"
        'Me.bsm_form_grp.DataSource = ds
End Sub
Public Overrides Sub Createcolumns()
        dggrp = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgform = New System.Windows.Forms.DataGridViewTextBoxColumn
dggroupbox = New System.Windows.Forms.DataGridViewTextBoxColumn
dgRO = New System.Windows.Forms.DataGridViewCheckBoxColumn
End Sub
Public Overrides Sub Adjustcolumns(ByVal blnAdjustWidth As Boolean)
 MyBase.Adjustcolumns(blnAdjustWidth)
        'DefineComboBoxColumn(dggrp, MainDefs.strGetFormat("NVarChar"), True, "grp", "", FieldWidths.GENWIDTH, blnRO, true, "",  bsm_form_grp, "grp" ,"grp", Color.Lavender)
        'If blnRO = True Then dggrp.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
        'DefineComboBoxColumn(dgform, MainDefs.strGetFormat("NVarChar"), True, "form", "", FieldWidths.GENWIDTH, blnRO, true, "",  bsm_form_grp, "form" ,"form", Color.Lavender)
        '        If blnRO = True Then dgform.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
        DefineColumn(dggrp, "grp", blnRO, ds.m_form_grp_groupbox.grpColumn.MaxLength)
        DefineColumn(dgform, "form", blnRO, ds.m_form_grp_groupbox.formColumn.MaxLength)

DefineColumn(dggroupbox, "groupbox", blnRO, ds.m_form_grp_groupbox.groupboxColumn.MaxLength)
DefineColumn(dgRO, "RO", blnRO, ds.m_form_grp_groupbox.ROColumn.MaxLength)
PutColumnsInGrid()
AdjustDataGridWidth(blnAdjustWidth)
RefreshCombos()
End Sub
Public Overrides Sub RefreshCombos()
MyBase.RefreshCombos()
        'Me.m_form_grpTableAdapter.Fill(Me.ds.m_form_grp)
dg.CancelEdit()
        iComboCount = 0
End sub
#Region "Editing"
Public Overrides Sub dg_UserDeletingRow(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs)
Try
Dim tadap As TheDataSetTableAdapters.m_form_grp_groupboxTableAdapter
tadap = CType(ta, TheDataSetTableAdapters.m_form_grp_groupboxTableAdapter)
tadap.Delete(e.Row.Cells(dg.Columns("grp").Index).Value.ToString(),e.Row.Cells(dg.Columns("form").Index).Value.ToString(),e.Row.Cells(dg.Columns("groupbox").Index).Value.ToString())
MyBase.dg_UserDeletingRow(sender, e)
Catch ex As Exception
MsgBox("Delete failed. Most common cause is that record is in use in another table." + ex.message)
e.Cancel = True
End Try
End Sub
Private Sub dg_DefaultValuesNeeded(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dg.DefaultValuesNeeded
With e.Row
.Cells("RO").Value =  ((1))
End With
End Sub
#End Region
#Region "Filter"
Public Overrides Sub CreateFilterBoxes(ByVal _Controls As Control.ControlCollection)
MyBase.CreateFilterBoxes(_Controls)
CreateAFilterBox(tbgrpFind, "grp", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbformFind, "form", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbgroupboxFind, "groupbox", AddressOf tbFind_TextChanged, _Controls)
CreateACheckBox(cbROFind, "RO", AddressOf cbFind_CheckChanged, _Controls)
End Sub
Friend WithEvents tbgrpFind As System.Windows.Forms.TextBox
Friend WithEvents tbformFind As System.Windows.Forms.TextBox
Friend WithEvents tbgroupboxFind As System.Windows.Forms.TextBox
Friend WithEvents cbROFind As System.Windows.Forms.CheckBox
Private Sub cbFind_CheckChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
MakeFilter(False)
End Sub
Private Sub tbFind_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
 MakeFilter(False)
End Sub
#End Region
End Class
