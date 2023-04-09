'------------------------------------------------
'Name: Module gen_TheDataSet_m_form.vb.
'Function: 
'Copyright Robin Baines 2008. All rights reserved.
'Created 7/8/2012 12:00:00 AM.
'Notes: 
'Modifications:
'------------------------------------------------
Imports Utilities
Imports System.Windows.Forms
Imports System.Drawing
Public Class TheDataSet_m_form
Inherits dgColumns
Friend WithEvents dgform As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgmenu As System.Windows.Forms.DataGridViewCheckBoxColumn
Friend WithEvents dgmenu_entry As System.Windows.Forms.DataGridViewCheckBoxColumn
Public Sub New(ByVal strForm As String, ByVal _bs As BindingSource, ByVal _dg As dgvEnter, _
ByVal _ta As TheDataSetTableAdapters.m_formTableAdapter, _
ByVal _ds As DataSet, _
ByVal _components As System.ComponentModel.Container, _
ByVal _MainDefs As MainDefinitions, _
ByVal blnRO As Boolean, _
ByVal _Controls As Control.ControlCollection, ByVal _frmStandard As frmStandard, _
ByVal blnFilters As Boolean)
MyBase.New(strForm, "m_form", _bs, _dg, _ta, _ds, _MainDefs, blnRO, _
"form","",_Controls, _frmStandard, blnFilters)
_ta.Connection.ConnectionString = GetConnectionString()
End Sub
Public Overrides Sub Createcolumns()
dgform = New System.Windows.Forms.DataGridViewTextBoxColumn
dgmenu = New System.Windows.Forms.DataGridViewCheckBoxColumn
dgmenu_entry = New System.Windows.Forms.DataGridViewCheckBoxColumn
End Sub
Public Overrides Sub Adjustcolumns(ByVal blnAdjustWidth As Boolean)
 MyBase.Adjustcolumns(blnAdjustWidth)
DefineColumn(dgform, "form", blnRO, ds.m_form.formColumn.MaxLength)
DefineColumn(dgmenu, "menu", blnRO, ds.m_form.menuColumn.MaxLength)
DefineColumn(dgmenu_entry, "menu_entry", blnRO, ds.m_form.menu_entryColumn.MaxLength)
PutColumnsInGrid()
AdjustDataGridWidth(blnAdjustWidth)
RefreshCombos()
End Sub
Public Overrides Sub RefreshCombos()
MyBase.RefreshCombos()
dg.CancelEdit()
iComboCount = 0
End sub
#Region "Editing"
Public Overrides Sub dg_UserDeletingRow(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs)
Try
Dim tadap As TheDataSetTableAdapters.m_formTableAdapter
tadap = CType(ta, TheDataSetTableAdapters.m_formTableAdapter)
tadap.Delete(e.Row.Cells(dg.Columns("form").Index).Value.ToString())
MyBase.dg_UserDeletingRow(sender, e)
Catch ex As Exception
MsgBox("Delete failed. Most common cause is that record is in use in another table." + ex.message)
e.Cancel = True
End Try
End Sub
Private Sub dg_DefaultValuesNeeded(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dg.DefaultValuesNeeded
With e.Row
.Cells("menu").Value =  ((0))
.Cells("menu_entry").Value =  ((0))
End With
End Sub
#End Region
#Region "Filter"
Public Overrides Sub CreateFilterBoxes(ByVal _Controls As Control.ControlCollection)
MyBase.CreateFilterBoxes(_Controls)
CreateAFilterBox(tbformFind, "form", AddressOf tbFind_TextChanged, _Controls)
CreateACheckBox(cbmenuFind, "menu", AddressOf cbFind_CheckChanged, _Controls)
CreateACheckBox(cbmenu_entryFind, "menu_entry", AddressOf cbFind_CheckChanged, _Controls)
End Sub
Friend WithEvents tbformFind As System.Windows.Forms.TextBox
Friend WithEvents cbmenuFind As System.Windows.Forms.CheckBox
Friend WithEvents cbmenu_entryFind As System.Windows.Forms.CheckBox
Private Sub cbFind_CheckChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
MakeFilter(False)
End Sub
Private Sub tbFind_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
 MakeFilter(False)
End Sub
#End Region
End Class
