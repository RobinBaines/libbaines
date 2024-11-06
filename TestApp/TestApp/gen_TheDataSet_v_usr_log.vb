'------------------------------------------------
'Name: Module gen_TheDataSet_v_usr_log.vb.
'Function: 
'Copyright Robin Baines 2008. All rights reserved.
'Created 4-11-2024 00:00:00.
'Notes: 
'Modifications:
'------------------------------------------------
Imports Utilities
Imports System.Windows.Forms
Imports System.Drawing
Public Class TheDataSet_v_usr_log
Inherits dgColumns
Friend WithEvents dgusr As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgapp As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgyear_minutes As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgyear_loggedin As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgyear_loggedin_unauthorized As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgyear_logout As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgyear_average As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgmonth_minutes As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgmonth_loggedin As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgmonth_loggedin_unauthorized As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgmonth_logout As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgmonth_average As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dglogged_in As System.Windows.Forms.DataGridViewCheckBoxColumn
Friend WithEvents dgdays_since_log As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgblocked As System.Windows.Forms.DataGridViewCheckBoxColumn
Friend WithEvents dgblocking As System.Windows.Forms.DataGridViewCheckBoxColumn
Friend WithEvents dgname As System.Windows.Forms.DataGridViewTextBoxColumn
Public Sub New(ByVal strForm As String, ByVal _bs As BindingSource, ByVal _dg As dgvEnter, _
ByVal _ta As TheDataSetTableAdapters.v_usr_logTableAdapter, _
ByVal _ds As DataSet, _
ByVal _components As System.ComponentModel.Container, _
ByVal _MainDefs As MainDefinitions, _
ByVal blnRO As Boolean, _
ByVal _Controls As Control.ControlCollection, ByVal _frmStandard As frmStandard, _
ByVal blnFilters As Boolean)
MyBase.New(strForm, "v_usr_log", _bs, _dg, _ta, _ds, _MainDefs, blnRO, _
"","",_Controls, _frmStandard, blnFilters)
_ta.Connection.ConnectionString = GetConnectionString()
End Sub
Public Overrides Sub Createcolumns()
dgusr = New System.Windows.Forms.DataGridViewTextBoxColumn
dgapp = New System.Windows.Forms.DataGridViewTextBoxColumn
dgyear_minutes = New System.Windows.Forms.DataGridViewTextBoxColumn
dgyear_loggedin = New System.Windows.Forms.DataGridViewTextBoxColumn
dgyear_loggedin_unauthorized = New System.Windows.Forms.DataGridViewTextBoxColumn
dgyear_logout = New System.Windows.Forms.DataGridViewTextBoxColumn
dgyear_average = New System.Windows.Forms.DataGridViewTextBoxColumn
dgmonth_minutes = New System.Windows.Forms.DataGridViewTextBoxColumn
dgmonth_loggedin = New System.Windows.Forms.DataGridViewTextBoxColumn
dgmonth_loggedin_unauthorized = New System.Windows.Forms.DataGridViewTextBoxColumn
dgmonth_logout = New System.Windows.Forms.DataGridViewTextBoxColumn
dgmonth_average = New System.Windows.Forms.DataGridViewTextBoxColumn
dglogged_in = New System.Windows.Forms.DataGridViewCheckBoxColumn
dgdays_since_log = New System.Windows.Forms.DataGridViewTextBoxColumn
dgblocked = New System.Windows.Forms.DataGridViewCheckBoxColumn
dgblocking = New System.Windows.Forms.DataGridViewCheckBoxColumn
dgname = New System.Windows.Forms.DataGridViewTextBoxColumn
End Sub
Public Overrides Sub Adjustcolumns(ByVal blnAdjustWidth As Boolean)
 MyBase.Adjustcolumns(blnAdjustWidth)
 Try
DefineColumn(dgusr, "usr", blnRO, ds.v_usr_log.usrColumn.MaxLength)
DefineColumn(dgapp, "app", blnRO, ds.v_usr_log.appColumn.MaxLength)
DefineColumn(dgyear_minutes, "year_minutes", blnRO, ds.v_usr_log.year_minutesColumn.MaxLength)
DefineColumn(dgyear_loggedin, "year_loggedin", blnRO, ds.v_usr_log.year_loggedinColumn.MaxLength)
DefineColumn(dgyear_loggedin_unauthorized, "year_loggedin_unauthorized", blnRO, ds.v_usr_log.year_loggedin_unauthorizedColumn.MaxLength)
DefineColumn(dgyear_logout, "year_logout", blnRO, ds.v_usr_log.year_logoutColumn.MaxLength)
DefineColumn(dgyear_average, "year_average", blnRO, ds.v_usr_log.year_averageColumn.MaxLength)
DefineColumn(dgmonth_minutes, "month_minutes", blnRO, ds.v_usr_log.month_minutesColumn.MaxLength)
DefineColumn(dgmonth_loggedin, "month_loggedin", blnRO, ds.v_usr_log.month_loggedinColumn.MaxLength)
DefineColumn(dgmonth_loggedin_unauthorized, "month_loggedin_unauthorized", blnRO, ds.v_usr_log.month_loggedin_unauthorizedColumn.MaxLength)
DefineColumn(dgmonth_logout, "month_logout", blnRO, ds.v_usr_log.month_logoutColumn.MaxLength)
DefineColumn(dgmonth_average, "month_average", blnRO, ds.v_usr_log.month_averageColumn.MaxLength)
DefineColumn(dglogged_in, "logged_in", blnRO, ds.v_usr_log.logged_inColumn.MaxLength)
DefineColumn(dgdays_since_log, "days_since_log", blnRO, ds.v_usr_log.days_since_logColumn.MaxLength)
DefineColumn(dgblocked, "blocked", blnRO, ds.v_usr_log.blockedColumn.MaxLength)
DefineColumn(dgblocking, "blocking", blnRO, ds.v_usr_log.blockingColumn.MaxLength)
DefineColumn(dgname, "name", blnRO, ds.v_usr_log.nameColumn.MaxLength)
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
#Region "Filter"
Public Overrides Sub CreateFilterBoxes(ByVal _Controls As Control.ControlCollection)
MyBase.CreateFilterBoxes(_Controls)
CreateAFilterBox(tbusrFind, "usr", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbappFind, "app", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbyear_minutesFind, "year_minutes", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbyear_loggedinFind, "year_loggedin", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbyear_loggedin_unauthorizedFind, "year_loggedin_unauthorized", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbyear_logoutFind, "year_logout", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbyear_averageFind, "year_average", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbmonth_minutesFind, "month_minutes", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbmonth_loggedinFind, "month_loggedin", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbmonth_loggedin_unauthorizedFind, "month_loggedin_unauthorized", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbmonth_logoutFind, "month_logout", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbmonth_averageFind, "month_average", AddressOf tbFind_TextChanged, _Controls)
CreateACheckBox(cblogged_inFind, "logged_in", AddressOf cbFind_CheckChanged, _Controls)
CreateAFilterBox(tbdays_since_logFind, "days_since_log", AddressOf tbFind_TextChanged, _Controls)
CreateACheckBox(cbblockedFind, "blocked", AddressOf cbFind_CheckChanged, _Controls)
CreateACheckBox(cbblockingFind, "blocking", AddressOf cbFind_CheckChanged, _Controls)
CreateAFilterBox(tbnameFind, "name", AddressOf tbFind_TextChanged, _Controls)
End Sub
Friend WithEvents tbusrFind As System.Windows.Forms.TextBox
Friend WithEvents tbappFind As System.Windows.Forms.TextBox
Friend WithEvents tbyear_minutesFind As System.Windows.Forms.TextBox
Friend WithEvents tbyear_loggedinFind As System.Windows.Forms.TextBox
Friend WithEvents tbyear_loggedin_unauthorizedFind As System.Windows.Forms.TextBox
Friend WithEvents tbyear_logoutFind As System.Windows.Forms.TextBox
Friend WithEvents tbyear_averageFind As System.Windows.Forms.TextBox
Friend WithEvents tbmonth_minutesFind As System.Windows.Forms.TextBox
Friend WithEvents tbmonth_loggedinFind As System.Windows.Forms.TextBox
Friend WithEvents tbmonth_loggedin_unauthorizedFind As System.Windows.Forms.TextBox
Friend WithEvents tbmonth_logoutFind As System.Windows.Forms.TextBox
Friend WithEvents tbmonth_averageFind As System.Windows.Forms.TextBox
Friend WithEvents cblogged_inFind As System.Windows.Forms.CheckBox
Friend WithEvents tbdays_since_logFind As System.Windows.Forms.TextBox
Friend WithEvents cbblockedFind As System.Windows.Forms.CheckBox
Friend WithEvents cbblockingFind As System.Windows.Forms.CheckBox
Friend WithEvents tbnameFind As System.Windows.Forms.TextBox
Private Sub cbFind_CheckChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
MakeFilter(False)
End Sub
Private Sub tbFind_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
 MakeFilter(False)
End Sub
#End Region
End Class
