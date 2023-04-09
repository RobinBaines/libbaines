'------------------------------------------------
'Name: Module TheDataSet_m_usr_change_log.vb.
'Function: 
'Copyright Robin Baines 2014. All rights reserved.
'Created 30-12-2014 0:00:00.
'Notes: 
'Modifications: default values commented out.
'------------------------------------------------
Imports Utilities
Imports System.Windows.Forms
Imports System.Drawing
Public Class TheDataSet_m_usr_change_log
Inherits dgColumns
Friend WithEvents dgId As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgusr As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dggrp As System.Windows.Forms.DataGridViewComboBoxColumn
Friend WithEvents dglang As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgblocking As System.Windows.Forms.DataGridViewCheckBoxColumn
Friend WithEvents dgname As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgemail As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgtelephone As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgcreatetime As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgusr_change As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgremark As System.Windows.Forms.DataGridViewTextBoxColumn
Public bsm_grp As BindingSource
Friend WithEvents m_grpTableAdapter As TheDataSetTableAdapters.m_grpTableAdapter
    Public Sub New(ByVal strForm As String, ByVal _bs As BindingSource, ByVal _dg As dgvEnter, _
    ByVal _ta As TheDataSetTableAdapters.m_usr_change_logTableAdapter, _
    ByVal _ds As DataSet, _
    ByVal _components As System.ComponentModel.Container, _
    ByVal _MainDefs As MainDefinitions, _
    ByVal blnRO As Boolean, _
    ByVal _Controls As Control.ControlCollection, ByVal _frmStandard As frmStandard, _
    ByVal blnFilters As Boolean, strTableName As String)
        MyBase.New(strForm, strTableName, _bs, _dg, _ta, _ds, _MainDefs, blnRO, _
        "Id", "", _Controls, _frmStandard, blnFilters)
        _ta.Connection.ConnectionString = GetConnectionString()
        Me.bsm_grp = New System.Windows.Forms.BindingSource(_components)
        m_grpTableAdapter = New TheDataSetTableAdapters.m_grpTableAdapter
        m_grpTableAdapter.Connection.ConnectionString = GetConnectionString()
        Me.bsm_grp.DataMember = "m_grp"
        Me.bsm_grp.DataSource = ds
    End Sub
Public Overrides Sub Createcolumns()
dgId = New System.Windows.Forms.DataGridViewTextBoxColumn
dgusr = New System.Windows.Forms.DataGridViewTextBoxColumn
dggrp = New System.Windows.Forms.DataGridViewComboBoxColumn
dglang = New System.Windows.Forms.DataGridViewTextBoxColumn
dgblocking = New System.Windows.Forms.DataGridViewCheckBoxColumn
dgname = New System.Windows.Forms.DataGridViewTextBoxColumn
dgemail = New System.Windows.Forms.DataGridViewTextBoxColumn
dgtelephone = New System.Windows.Forms.DataGridViewTextBoxColumn
dgcreatetime = New System.Windows.Forms.DataGridViewTextBoxColumn
dgusr_change = New System.Windows.Forms.DataGridViewTextBoxColumn
dgremark = New System.Windows.Forms.DataGridViewTextBoxColumn
End Sub
Public Overrides Sub Adjustcolumns(ByVal blnAdjustWidth As Boolean)
 MyBase.Adjustcolumns(blnAdjustWidth)
 Try
DefineColumn(dgId, "Id", blnRO, ds.m_usr_change_log.IdColumn.MaxLength)
DefineColumn(dgusr, "usr", blnRO, ds.m_usr_change_log.usrColumn.MaxLength)
DefineComboBoxColumn(dggrp, MainDefs.strGetFormat("NVarChar"), True, "grp", "", FieldWidths.GENWIDTH, blnRO, true, "",  bsm_grp, "grp" ,"grp", Color.Lavender)
If blnRO = True Then dggrp.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
DefineColumn(dglang, "lang", blnRO, ds.m_usr_change_log.langColumn.MaxLength)
DefineColumn(dgblocking, "blocking", blnRO, ds.m_usr_change_log.blockingColumn.MaxLength)
DefineColumn(dgname, "name", blnRO, ds.m_usr_change_log.nameColumn.MaxLength)
DefineColumn(dgemail, "email", blnRO, ds.m_usr_change_log.emailColumn.MaxLength)
DefineColumn(dgtelephone, "telephone", blnRO, ds.m_usr_change_log.telephoneColumn.MaxLength)
DefineColumn(dgcreatetime, "createtime", blnRO, ds.m_usr_change_log.createtimeColumn.MaxLength)
DefineColumn(dgusr_change, "usr_change", blnRO, ds.m_usr_change_log.usr_changeColumn.MaxLength)
DefineColumn(dgremark, "remark", blnRO, ds.m_usr_change_log.remarkColumn.MaxLength)
PutColumnsInGrid()
AdjustDataGridWidth(blnAdjustWidth)
RefreshCombos()
    Catch ex As Exception
MsgBox(ex.Message)
End Try
End Sub
Public Overrides Sub RefreshCombos()
MyBase.RefreshCombos()
Me.m_grpTableAdapter.Fill(Me.ds.m_grp)
dg.CancelEdit()
iComboCount = 1
End sub
#Region "Editing"
Public Overrides Sub dg_UserDeletingRow(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs)
Try
Dim tadap As TheDataSetTableAdapters.m_usr_change_logTableAdapter
tadap = CType(ta, TheDataSetTableAdapters.m_usr_change_logTableAdapter)
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
            '.Cells("usr_change").Value =  (suser_name())
End With
End Sub
#End Region
#Region "Filter"
Public Overrides Sub CreateFilterBoxes(ByVal _Controls As Control.ControlCollection)
MyBase.CreateFilterBoxes(_Controls)
CreateAFilterBox(tbIdFind, "Id", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbusrFind, "usr", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbgrpFind, "grp", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tblangFind, "lang", AddressOf tbFind_TextChanged, _Controls)
CreateACheckBox(cbblockingFind, "blocking", AddressOf cbFind_CheckChanged, _Controls)
CreateAFilterBox(tbnameFind, "name", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbemailFind, "email", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbtelephoneFind, "telephone", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbcreatetimeFind, "createtime", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbusr_changeFind, "usr_change", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbremarkFind, "remark", AddressOf tbFind_TextChanged, _Controls)
End Sub
Friend WithEvents tbIdFind As System.Windows.Forms.TextBox
Friend WithEvents tbusrFind As System.Windows.Forms.TextBox
Friend WithEvents tbgrpFind As System.Windows.Forms.TextBox
Friend WithEvents tblangFind As System.Windows.Forms.TextBox
Friend WithEvents cbblockingFind As System.Windows.Forms.CheckBox
Friend WithEvents tbnameFind As System.Windows.Forms.TextBox
Friend WithEvents tbemailFind As System.Windows.Forms.TextBox
Friend WithEvents tbtelephoneFind As System.Windows.Forms.TextBox
Friend WithEvents tbcreatetimeFind As System.Windows.Forms.TextBox
Friend WithEvents tbusr_changeFind As System.Windows.Forms.TextBox
Friend WithEvents tbremarkFind As System.Windows.Forms.TextBox
Private Sub cbFind_CheckChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
MakeFilter(False)
End Sub
Private Sub tbFind_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
 MakeFilter(False)
End Sub
#End Region
End Class
