'------------------------------------------------
'Name: Module gen_TheDataSet_v_txt_header.vb.
'Function: 
'Copyright Robin Baines 2008. All rights reserved.
'Created 7/8/2012 12:00:00 AM.
'Notes: 
'Modifications: Added delete
'------------------------------------------------
Imports Utilities
Imports System.Windows.Forms
Imports System.Drawing
Public Class TheDataSet_v_txt_header
Inherits dgColumns
Friend WithEvents dgtxt As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dglang As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgheader As System.Windows.Forms.DataGridViewTextBoxColumn
Public Sub New(ByVal strForm As String, ByVal _bs As BindingSource, ByVal _dg As dgvEnter, _
ByVal _ta As TheDataSetTableAdapters.v_txt_headerTableAdapter, _
ByVal _ds As DataSet, _
ByVal _components As System.ComponentModel.Container, _
ByVal _MainDefs As MainDefinitions, _
ByVal blnRO As Boolean, _
ByVal _Controls As Control.ControlCollection, ByVal _frmStandard As frmStandard, _
ByVal blnFilters As Boolean)
MyBase.New(strForm, "v_txt_header", _bs, _dg, _ta, _ds, _MainDefs, blnRO, _
"","",_Controls, _frmStandard, blnFilters)
_ta.Connection.ConnectionString = GetConnectionString()
End Sub
Public Overrides Sub Createcolumns()
dgtxt = New System.Windows.Forms.DataGridViewTextBoxColumn
dglang = New System.Windows.Forms.DataGridViewTextBoxColumn
dgheader = New System.Windows.Forms.DataGridViewTextBoxColumn
End Sub
Public Overrides Sub Adjustcolumns(ByVal blnAdjustWidth As Boolean)
 MyBase.Adjustcolumns(blnAdjustWidth)
DefineColumn(dgtxt, "txt", blnRO, ds.v_txt_header.txtColumn.MaxLength)
DefineColumn(dglang, "lang", blnRO, ds.v_txt_header.langColumn.MaxLength)
DefineColumn(dgheader, "header", blnRO, ds.v_txt_header.headerColumn.MaxLength)
PutColumnsInGrid()
AdjustDataGridWidth(blnAdjustWidth)
RefreshCombos()
End Sub
Public Overrides Sub RefreshCombos()
MyBase.RefreshCombos()
dg.CancelEdit()
iComboCount = 0
    End Sub
#Region "Editing"
    Public Overrides Sub dg_UserDeletingRow(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs)
        Try
            Dim tadap As TheDataSetTableAdapters.v_txt_headerTableAdapter
            tadap = CType(ta, TheDataSetTableAdapters.v_txt_headerTableAdapter)
            tadap.DeleteQuery(e.Row.Cells(dg.Columns("txt").Index).Value.ToString(), e.Row.Cells(dg.Columns("lang").Index).Value.ToString())
            e.Cancel = True 'Do this so the row stays.

            'Then refresh the data.
            Dim fManage As frmManage
            fManage = TryCast(frmParent, frmManage)

            If Not fManage Is Nothing Then
                fManage.Fills()
            End If
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
CreateAFilterBox(tblangFind, "lang", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbheaderFind, "header", AddressOf tbFind_TextChanged, _Controls)
End Sub
Friend WithEvents tbtxtFind As System.Windows.Forms.TextBox
Friend WithEvents tblangFind As System.Windows.Forms.TextBox
Friend WithEvents tbheaderFind As System.Windows.Forms.TextBox
Private Sub cbFind_CheckChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
MakeFilter(False)
End Sub
Private Sub tbFind_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
 MakeFilter(False)
End Sub
#End Region
End Class
