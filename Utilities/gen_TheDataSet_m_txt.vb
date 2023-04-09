'------------------------------------------------
'Name: Module gen_TheDataSet_m_txt.vb.
'Function: 
'Copyright Robin Baines 2008. All rights reserved.
'Created 7/8/2012 12:00:00 AM.
'Notes: 
'Modifications:
'------------------------------------------------
Imports Utilities
Imports System.Windows.Forms
Imports System.Drawing
Public Class TheDataSet_m_txt
Inherits dgColumns
Friend WithEvents dgtxt As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgdescr As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgtyp As System.Windows.Forms.DataGridViewTextBoxColumn
Public Sub New(ByVal strForm As String, ByVal _bs As BindingSource, ByVal _dg As dgvEnter, _
ByVal _ta As TheDataSetTableAdapters.m_txtTableAdapter, _
ByVal _ds As DataSet, _
ByVal _components As System.ComponentModel.Container, _
ByVal _MainDefs As MainDefinitions, _
ByVal blnRO As Boolean, _
ByVal _Controls As Control.ControlCollection, ByVal _frmStandard As frmStandard, _
ByVal blnFilters As Boolean)
MyBase.New(strForm, "m_txt", _bs, _dg, _ta, _ds, _MainDefs, blnRO, _
"txt","",_Controls, _frmStandard, blnFilters)
_ta.Connection.ConnectionString = GetConnectionString()
End Sub
Public Overrides Sub Createcolumns()
dgtxt = New System.Windows.Forms.DataGridViewTextBoxColumn
dgdescr = New System.Windows.Forms.DataGridViewTextBoxColumn
dgtyp = New System.Windows.Forms.DataGridViewTextBoxColumn
End Sub
Public Overrides Sub Adjustcolumns(ByVal blnAdjustWidth As Boolean)
 MyBase.Adjustcolumns(blnAdjustWidth)
DefineColumn(dgtxt, "txt", blnRO, ds.m_txt.txtColumn.MaxLength)
DefineColumn(dgdescr, "descr", blnRO, ds.m_txt.descrColumn.MaxLength)
DefineColumn(dgtyp, "typ", blnRO, ds.m_txt.typColumn.MaxLength)
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
Dim tadap As TheDataSetTableAdapters.m_txtTableAdapter
tadap = CType(ta, TheDataSetTableAdapters.m_txtTableAdapter)
tadap.Delete(e.Row.Cells(dg.Columns("txt").Index).Value.ToString())
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
CreateAFilterBox(tbtxtFind, "txt", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbdescrFind, "descr", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbtypFind, "typ", AddressOf tbFind_TextChanged, _Controls)
End Sub
Friend WithEvents tbtxtFind As System.Windows.Forms.TextBox
Friend WithEvents tbdescrFind As System.Windows.Forms.TextBox
Friend WithEvents tbtypFind As System.Windows.Forms.TextBox
Private Sub cbFind_CheckChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
MakeFilter(False)
End Sub
Private Sub tbFind_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
 MakeFilter(False)
End Sub
#End Region
End Class
