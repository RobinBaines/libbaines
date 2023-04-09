'------------------------------------------------
'Name: Module gen_TheDataSet_m_form_grp.vb.
'Function: 
'Copyright Robin Baines 2008. All rights reserved.
'Created 7/8/2012 12:00:00 AM.
'Notes: 
'Modifications: combo to textbox.
'------------------------------------------------
Imports Utilities
Imports System.Windows.Forms
Imports System.Drawing
Public Class TheDataSet_m_form_grp
Inherits dgColumns
    Friend WithEvents dggrp As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgform As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgRO As System.Windows.Forms.DataGridViewCheckBoxColumn
    'Public bsm_grp As BindingSource
    'Friend WithEvents m_grpTableAdapter As TheDataSetTableAdapters.m_grpTableAdapter
    'Public bsm_form As BindingSource
    'Friend WithEvents m_formTableAdapter As TheDataSetTableAdapters.m_formTableAdapter
    Public Sub New(ByVal strForm As String, ByVal _bs As BindingSource, ByVal _dg As dgvEnter, _
    ByVal _ta As TheDataSetTableAdapters.m_form_grpTableAdapter, _
    ByVal _ds As DataSet, _
    ByVal _components As System.ComponentModel.Container, _
    ByVal _MainDefs As MainDefinitions, _
    ByVal blnRO As Boolean, _
    ByVal _Controls As Control.ControlCollection, ByVal _frmStandard As frmStandard, _
    ByVal blnFilters As Boolean, strTableName As String)
        MyBase.New(strForm, strTableName, _bs, _dg, _ta, _ds, _MainDefs, blnRO, _
        "grp", "form", _Controls, _frmStandard, blnFilters)
        _ta.Connection.ConnectionString = GetConnectionString()
        'Me.bsm_grp = New System.Windows.Forms.BindingSource(_components)
        'm_grpTableAdapter = New TheDataSetTableAdapters.m_grpTableAdapter
        'm_grpTableAdapter.Connection.ConnectionString = GetConnectionString()
        'Me.bsm_grp.DataMember = "m_grp"
        'Me.bsm_grp.DataSource = ds
        'Me.bsm_form = New System.Windows.Forms.BindingSource(_components)
        'm_formTableAdapter = New TheDataSetTableAdapters.m_formTableAdapter
        'm_formTableAdapter.Connection.ConnectionString = GetConnectionString()
        'Me.bsm_form.DataMember = "m_form"
        'Me.bsm_form.DataSource = ds
    End Sub
Public Overrides Sub Createcolumns()
        dggrp = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgform = New System.Windows.Forms.DataGridViewTextBoxColumn
dgRO = New System.Windows.Forms.DataGridViewCheckBoxColumn
End Sub
Public Overrides Sub Adjustcolumns(ByVal blnAdjustWidth As Boolean)
 MyBase.Adjustcolumns(blnAdjustWidth)
        'DefineComboBoxColumn(dggrp, MainDefs.strGetFormat("TYP_M_STRING"), True, "grp", "", FieldWidths.GENWIDTH, blnRO, true, "",  bsm_grp, "grp" ,"grp", Color.Lavender)
        '        If blnRO = True Then dggrp.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
        DefineColumn(dggrp, "grp", blnRO, ds.m_form_grp.grpColumn.MaxLength)
        DefineColumn(dgform, "form", blnRO, ds.m_form_grp.formColumn.MaxLength)
        'DefineComboBoxColumn(dgform, MainDefs.strGetFormat("TYP_M_STRING"), True, "form", "", FieldWidths.GENWIDTH, blnRO, true, "",  bsm_form, "form" ,"form", Color.Lavender)
        'If blnRO = True Then dgform.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
DefineColumn(dgRO, "RO", blnRO, ds.m_form_grp.ROColumn.MaxLength)
PutColumnsInGrid()
AdjustDataGridWidth(blnAdjustWidth)
RefreshCombos()
End Sub
Public Overrides Sub RefreshCombos()
MyBase.RefreshCombos()
        'Me.m_grpTableAdapter.Fill(Me.ds.m_grp)
        'Me.m_formTableAdapter.Fill(Me.ds.m_form)
dg.CancelEdit()
        iComboCount = 0
End sub
#Region "Editing"
Public Overrides Sub dg_UserDeletingRow(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs)
Try
Dim tadap As TheDataSetTableAdapters.m_form_grpTableAdapter
tadap = CType(ta, TheDataSetTableAdapters.m_form_grpTableAdapter)
tadap.Delete(e.Row.Cells(dg.Columns("grp").Index).Value.ToString(),e.Row.Cells(dg.Columns("form").Index).Value.ToString())
MyBase.dg_UserDeletingRow(sender, e)
Catch ex As Exception
MsgBox("Delete failed. Most common cause is that record is in use in another table." + ex.message)
e.Cancel = True
End Try
End Sub
    'Private Sub dg_DefaultValuesNeeded(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dg.DefaultValuesNeeded
    'With e.Row
    '.Cells("RO").Value =  ((1))
    'End With
    'End Sub
#End Region
#Region "Filter"
Public Overrides Sub CreateFilterBoxes(ByVal _Controls As Control.ControlCollection)
MyBase.CreateFilterBoxes(_Controls)
CreateAFilterBox(tbgrpFind, "grp", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbformFind, "form", AddressOf tbFind_TextChanged, _Controls)
CreateACheckBox(cbROFind, "RO", AddressOf cbFind_CheckChanged, _Controls)
End Sub
Friend WithEvents tbgrpFind As System.Windows.Forms.TextBox
Friend WithEvents tbformFind As System.Windows.Forms.TextBox
Friend WithEvents cbROFind As System.Windows.Forms.CheckBox
Private Sub cbFind_CheckChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
MakeFilter(False)
End Sub
Private Sub tbFind_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
 MakeFilter(False)
End Sub
#End Region
End Class
