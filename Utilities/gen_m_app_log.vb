'------------------------------------------------
'Name: Module m_app_log.vb.
'Function: 
'Copyright Robin Baines 2008. All rights reserved.
'Created 4/13/2011 12:00:00 AM.
'Notes: 
'Modifications:
'------------------------------------------------
Imports Utilities
Imports System.Windows.Forms
Imports System.Drawing
Public Class m_app_log
Inherits dgColumns
Friend WithEvents dgId As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgapp As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgusr As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgerror As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgraised_in As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgpriority As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgcreateTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Public Sub New(ByVal strForm As String, ByVal _bs As BindingSource, ByVal _dg As dgvEnter, _
    ByVal _ta As TheDataSetTableAdapters.m_app_logTableAdapter, _
    ByVal _ds As DataSet, _
    ByVal _components As System.ComponentModel.Container, _
    ByVal _MainDefs As MainDefinitions, _
    ByVal blnRO As Boolean, _
    ByVal _Controls As Control.ControlCollection, ByVal _frmStandard As frmStandard, _blnFilters As Boolean)
        MyBase.New(strForm, "m_app_log", _bs, _dg, _ta, _ds, _MainDefs, blnRO, _
"Id", "", _Controls, _frmStandard, _blnFilters)
        _ta.Connection.ConnectionString = GetConnectionString()
    End Sub
Public Overrides Sub Createcolumns()
dgId = New System.Windows.Forms.DataGridViewTextBoxColumn
dgapp = New System.Windows.Forms.DataGridViewTextBoxColumn
dgusr = New System.Windows.Forms.DataGridViewTextBoxColumn
dgerror = New System.Windows.Forms.DataGridViewTextBoxColumn
dgraised_in = New System.Windows.Forms.DataGridViewTextBoxColumn
dgpriority = New System.Windows.Forms.DataGridViewTextBoxColumn
dgcreateTime = New System.Windows.Forms.DataGridViewTextBoxColumn
End Sub
Public Overrides Sub Adjustcolumns(ByVal blnAdjustWidth As Boolean)
 Dim TheDataSet As TheDataSet = New TheDataSet
DefineColumn(dgId, MainDefs.strGetFormat("Int"), True, "Id", "", FieldWidths.FLOATWIDTH, blnRO, true, "", false, ds.m_app_log.IdColumn.MaxLength)
DefineColumn(dgapp, MainDefs.strGetFormat("NVarChar"), True, "app", "", FieldWidths.GENWIDTH, blnRO, true, "", false, ds.m_app_log.appColumn.MaxLength)
DefineColumn(dgusr, MainDefs.strGetFormat("NVarChar"), True, "usr", "", FieldWidths.GENWIDTH, blnRO, true, "", false, ds.m_app_log.usrColumn.MaxLength)
DefineColumn(dgerror, MainDefs.strGetFormat("NVarCharMax"), True, "error", "", FieldWidths.GENWIDTH, blnRO, true, "", false, ds.m_app_log.errorColumn.MaxLength)
DefineColumn(dgraised_in, MainDefs.strGetFormat("NVarChar"), True, "raised_in", "", FieldWidths.GENWIDTH, blnRO, true, "", false, ds.m_app_log.raised_inColumn.MaxLength)
DefineColumn(dgpriority, MainDefs.strGetFormat("Int"), True, "priority", "", FieldWidths.FLOATWIDTH, blnRO, true, "", false, ds.m_app_log.priorityColumn.MaxLength)
DefineColumn(dgcreateTime, MainDefs.strGetFormat("DateTime"), True, "createTime", "", FieldWidths.GENWIDTH, blnRO, true, "", false, ds.m_app_log.createTimeColumn.MaxLength)
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
Dim tadap As TheDataSetTableAdapters.m_app_logTableAdapter
tadap = CType(ta, TheDataSetTableAdapters.m_app_logTableAdapter)
tadap.Delete(e.Row.Cells(dg.Columns("Id").Index).Value.ToString())
MyBase.dg_UserDeletingRow(sender, e)
Catch ex As Exception
MsgBox("Delete failed. Most common cause is that record is in use in another table." + ex.message)
e.Cancel = True
End Try
End Sub
Private Sub dg_DefaultValuesNeeded(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dg.DefaultValuesNeeded
With e.Row
.Cells("priority").Value =  ParseConstraint("((0))")
.Cells("createTime").Value =  ParseConstraint("(getdate())")
End With
End Sub
#End Region
#Region "Filter"
Public Overrides Sub CreateFilterBoxes(ByVal _Controls As Control.ControlCollection)
MyBase.CreateFilterBoxes(_Controls)
CreateAFilterBox(tbIdFind, "Id", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbappFind, "app", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbusrFind, "usr", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tberrorFind, "error", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbraised_inFind, "raised_in", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbpriorityFind, "priority", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbcreateTimeFind, "createTime", AddressOf tbFind_TextChanged, _Controls)
End Sub
Friend WithEvents tbIdFind As System.Windows.Forms.TextBox
Friend WithEvents tbappFind As System.Windows.Forms.TextBox
Friend WithEvents tbusrFind As System.Windows.Forms.TextBox
Friend WithEvents tberrorFind As System.Windows.Forms.TextBox
Friend WithEvents tbraised_inFind As System.Windows.Forms.TextBox
Friend WithEvents tbpriorityFind As System.Windows.Forms.TextBox
Friend WithEvents tbcreateTimeFind As System.Windows.Forms.TextBox
Private Sub cbFind_CheckChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
MakeFilter(False)
End Sub
Private Sub tbFind_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
 MakeFilter(False)
End Sub
#End Region
End Class
