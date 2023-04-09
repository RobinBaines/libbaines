'------------------------------------------------
'Name: Module gen_TheDataSet_m_usr_log.vb.
'Function: 
'Copyright Robin Baines 2008. All rights reserved.
'Created 7/8/2012 12:00:00 AM.
'Notes: 
'Modifications:
'------------------------------------------------
Imports Utilities
Imports System.Windows.Forms
Imports System.Drawing
Public Class TheDataSet_m_usr_log
Inherits dgColumns
Friend WithEvents dgId As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgapp As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dglogout As System.Windows.Forms.DataGridViewCheckBoxColumn
Friend WithEvents dgsql_usr As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgwindowsIdentity As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgusr As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgauthorized As System.Windows.Forms.DataGridViewCheckBoxColumn
Friend WithEvents dgcreateTime As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgminutes As System.Windows.Forms.DataGridViewTextBoxColumn
Public Sub New(ByVal strForm As String, ByVal _bs As BindingSource, ByVal _dg As dgvEnter, _
ByVal _ta As TheDataSetTableAdapters.m_usr_logTableAdapter, _
ByVal _ds As DataSet, _
ByVal _components As System.ComponentModel.Container, _
ByVal _MainDefs As MainDefinitions, _
ByVal blnRO As Boolean, _
ByVal _Controls As Control.ControlCollection, ByVal _frmStandard As frmStandard, _
ByVal blnFilters As Boolean)
MyBase.New(strForm, "m_usr_log", _bs, _dg, _ta, _ds, _MainDefs, blnRO, _
"Id","",_Controls, _frmStandard, blnFilters)
_ta.Connection.ConnectionString = GetConnectionString()
End Sub
Public Overrides Sub Createcolumns()
dgId = New System.Windows.Forms.DataGridViewTextBoxColumn
dgapp = New System.Windows.Forms.DataGridViewTextBoxColumn
dglogout = New System.Windows.Forms.DataGridViewCheckBoxColumn
dgsql_usr = New System.Windows.Forms.DataGridViewTextBoxColumn
dgwindowsIdentity = New System.Windows.Forms.DataGridViewTextBoxColumn
dgusr = New System.Windows.Forms.DataGridViewTextBoxColumn
dgauthorized = New System.Windows.Forms.DataGridViewCheckBoxColumn
dgcreateTime = New System.Windows.Forms.DataGridViewTextBoxColumn
dgminutes = New System.Windows.Forms.DataGridViewTextBoxColumn
End Sub
Public Overrides Sub Adjustcolumns(ByVal blnAdjustWidth As Boolean)
 MyBase.Adjustcolumns(blnAdjustWidth)
DefineColumn(dgId, "Id", blnRO, ds.m_usr_log.IdColumn.MaxLength)
DefineColumn(dgapp, "app", blnRO, ds.m_usr_log.appColumn.MaxLength)
DefineColumn(dglogout, "logout", blnRO, ds.m_usr_log.logoutColumn.MaxLength)
DefineColumn(dgsql_usr, "sql_usr", blnRO, ds.m_usr_log.sql_usrColumn.MaxLength)
DefineColumn(dgwindowsIdentity, "windowsIdentity", blnRO, ds.m_usr_log.windowsIdentityColumn.MaxLength)
DefineColumn(dgusr, "usr", blnRO, ds.m_usr_log.usrColumn.MaxLength)
DefineColumn(dgauthorized, "authorized", blnRO, ds.m_usr_log.authorizedColumn.MaxLength)
DefineColumn(dgcreateTime, "createTime", blnRO, ds.m_usr_log.createTimeColumn.MaxLength)
DefineColumn(dgminutes, "minutes", blnRO, ds.m_usr_log.minutesColumn.MaxLength)
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
Dim tadap As TheDataSetTableAdapters.m_usr_logTableAdapter
tadap = CType(ta, TheDataSetTableAdapters.m_usr_logTableAdapter)
tadap.Delete(e.Row.Cells(dg.Columns("Id").Index).Value.ToString())
MyBase.dg_UserDeletingRow(sender, e)
Catch ex As Exception
MsgBox("Delete failed. Most common cause is that record is in use in another table." + ex.message)
e.Cancel = True
End Try
End Sub
Private Sub dg_DefaultValuesNeeded(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dg.DefaultValuesNeeded
With e.Row
            '.Cells("logout").Value =  ((0))
            '.Cells("sql_usr").Value =  (left(suser_sname(),(128)))
            '.Cells("authorized").Value =  ((0))
            '.Cells("createTime").Value =  (getdate())
            '.Cells("minutes").Value =  ((0))
End With
End Sub
#End Region
#Region "Filter"
Public Overrides Sub CreateFilterBoxes(ByVal _Controls As Control.ControlCollection)
MyBase.CreateFilterBoxes(_Controls)
CreateAFilterBox(tbIdFind, "Id", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbappFind, "app", AddressOf tbFind_TextChanged, _Controls)
CreateACheckBox(cblogoutFind, "logout", AddressOf cbFind_CheckChanged, _Controls)
CreateAFilterBox(tbsql_usrFind, "sql_usr", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbwindowsIdentityFind, "windowsIdentity", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbusrFind, "usr", AddressOf tbFind_TextChanged, _Controls)
CreateACheckBox(cbauthorizedFind, "authorized", AddressOf cbFind_CheckChanged, _Controls)
CreateAFilterBox(tbcreateTimeFind, "createTime", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbminutesFind, "minutes", AddressOf tbFind_TextChanged, _Controls)
End Sub
Friend WithEvents tbIdFind As System.Windows.Forms.TextBox
Friend WithEvents tbappFind As System.Windows.Forms.TextBox
Friend WithEvents cblogoutFind As System.Windows.Forms.CheckBox
Friend WithEvents tbsql_usrFind As System.Windows.Forms.TextBox
Friend WithEvents tbwindowsIdentityFind As System.Windows.Forms.TextBox
Friend WithEvents tbusrFind As System.Windows.Forms.TextBox
Friend WithEvents cbauthorizedFind As System.Windows.Forms.CheckBox
Friend WithEvents tbcreateTimeFind As System.Windows.Forms.TextBox
Friend WithEvents tbminutesFind As System.Windows.Forms.TextBox
Private Sub cbFind_CheckChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
MakeFilter(False)
End Sub
Private Sub tbFind_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
 MakeFilter(False)
End Sub
#End Region
End Class
