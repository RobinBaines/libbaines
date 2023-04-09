'------------------------------------------------
'Name: Module gen_TheDataSet_m_form_grp_log.vb.
'Function: 
'Copyright Robin Baines 2008. All rights reserved.
'Created 29-12-2014 0:00:00.
'Notes: 
'Modifications: Edited the get default values.
'------------------------------------------------
Imports Utilities
Imports System.Windows.Forms
Imports System.Drawing
Public Class TheDataSet_m_form_grp_log
Inherits dgColumns
Friend WithEvents dgId As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dggrp As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgform As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgRO As System.Windows.Forms.DataGridViewCheckBoxColumn
Friend WithEvents dgcreatetime As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgusr As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgremark As System.Windows.Forms.DataGridViewTextBoxColumn
    Public Sub New(ByVal strForm As String, ByVal _bs As BindingSource, ByVal _dg As dgvEnter, _
    ByVal _ta As TheDataSetTableAdapters.m_form_grp_logTableAdapter, _
    ByVal _ds As DataSet, _
    ByVal _components As System.ComponentModel.Container, _
    ByVal _MainDefs As MainDefinitions, _
    ByVal blnRO As Boolean, _
    ByVal _Controls As Control.ControlCollection, ByVal _frmStandard As frmStandard, _
    ByVal blnFilters As Boolean, strTableName As String)
        MyBase.New(strForm, strTableName, _bs, _dg, _ta, _ds, _MainDefs, blnRO, _
        "Id", "", _Controls, _frmStandard, blnFilters)
        _ta.Connection.ConnectionString = GetConnectionString()
    End Sub
Public Overrides Sub Createcolumns()
dgId = New System.Windows.Forms.DataGridViewTextBoxColumn
dggrp = New System.Windows.Forms.DataGridViewTextBoxColumn
dgform = New System.Windows.Forms.DataGridViewTextBoxColumn
dgRO = New System.Windows.Forms.DataGridViewCheckBoxColumn
dgcreatetime = New System.Windows.Forms.DataGridViewTextBoxColumn
dgusr = New System.Windows.Forms.DataGridViewTextBoxColumn
dgremark = New System.Windows.Forms.DataGridViewTextBoxColumn
End Sub
Public Overrides Sub Adjustcolumns(ByVal blnAdjustWidth As Boolean)
 MyBase.Adjustcolumns(blnAdjustWidth)
 Try
DefineColumn(dgId, "Id", blnRO, ds.m_form_grp_log.IdColumn.MaxLength)
DefineColumn(dggrp, "grp", blnRO, ds.m_form_grp_log.grpColumn.MaxLength)
DefineColumn(dgform, "form", blnRO, ds.m_form_grp_log.formColumn.MaxLength)
DefineColumn(dgRO, "RO", blnRO, ds.m_form_grp_log.ROColumn.MaxLength)
DefineColumn(dgcreatetime, "createtime", blnRO, ds.m_form_grp_log.createtimeColumn.MaxLength)
DefineColumn(dgusr, "usr", blnRO, ds.m_form_grp_log.usrColumn.MaxLength)
DefineColumn(dgremark, "remark", blnRO, ds.m_form_grp_log.remarkColumn.MaxLength)
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
#Region "Editing"
Public Overrides Sub dg_UserDeletingRow(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs)
Try
Dim tadap As TheDataSetTableAdapters.m_form_grp_logTableAdapter
tadap = CType(ta, TheDataSetTableAdapters.m_form_grp_logTableAdapter)
tadap.Delete(e.Row.Cells(dg.Columns("Id").Index).Value.ToString())
MyBase.dg_UserDeletingRow(sender, e)
Catch ex As Exception
MsgBox("Delete failed. Most common cause is that record is in use in another table." + ex.message)
e.Cancel = True
End Try
End Sub
Private Sub dg_DefaultValuesNeeded(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dg.DefaultValuesNeeded
With e.Row
            '.Cells("createtime").Value =  (getdate())
            '.Cells("usr").Value =  (suser_name())
End With
End Sub
#End Region
#Region "Filter"
Public Overrides Sub CreateFilterBoxes(ByVal _Controls As Control.ControlCollection)
MyBase.CreateFilterBoxes(_Controls)
CreateAFilterBox(tbIdFind, "Id", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbgrpFind, "grp", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbformFind, "form", AddressOf tbFind_TextChanged, _Controls)
CreateACheckBox(cbROFind, "RO", AddressOf cbFind_CheckChanged, _Controls)
CreateAFilterBox(tbcreatetimeFind, "createtime", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbusrFind, "usr", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbremarkFind, "remark", AddressOf tbFind_TextChanged, _Controls)
End Sub
Friend WithEvents tbIdFind As System.Windows.Forms.TextBox
Friend WithEvents tbgrpFind As System.Windows.Forms.TextBox
Friend WithEvents tbformFind As System.Windows.Forms.TextBox
Friend WithEvents cbROFind As System.Windows.Forms.CheckBox
Friend WithEvents tbcreatetimeFind As System.Windows.Forms.TextBox
Friend WithEvents tbusrFind As System.Windows.Forms.TextBox
Friend WithEvents tbremarkFind As System.Windows.Forms.TextBox
Private Sub cbFind_CheckChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
MakeFilter(False)
End Sub
Private Sub tbFind_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
 MakeFilter(False)
End Sub
#End Region
End Class
