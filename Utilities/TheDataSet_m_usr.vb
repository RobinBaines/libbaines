'------------------------------------------------
'Name: Module TheDataSet_m_usr.vb.
'Function: 
'Copyright Robin Baines 2008. All rights reserved.
'Created 30-12-2014 0:00:00.
'Notes: 
'Modifications: edited the defaults.
'------------------------------------------------
Imports Utilities
Imports System.Windows.Forms
Imports System.Drawing
Public Class TheDataSet_m_usr
Inherits dgColumns
Friend WithEvents dgusr As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dggrp As System.Windows.Forms.DataGridViewComboBoxColumn
Friend WithEvents dglang As System.Windows.Forms.DataGridViewComboBoxColumn
Friend WithEvents dgblocking As System.Windows.Forms.DataGridViewCheckBoxColumn
Friend WithEvents dgname As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgemail As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgtelephone As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgcreatetime As System.Windows.Forms.DataGridViewTextBoxColumn
Public bsm_grp As BindingSource
Friend WithEvents m_grpTableAdapter As TheDataSetTableAdapters.m_grpTableAdapter
Public bsm_lang As BindingSource
Friend WithEvents m_langTableAdapter As TheDataSetTableAdapters.m_langTableAdapter
Public Sub New(ByVal strForm As String, ByVal _bs As BindingSource, ByVal _dg As dgvEnter, _
ByVal _ta As TheDataSetTableAdapters.m_usrTableAdapter, _
ByVal _ds As DataSet, _
ByVal _components As System.ComponentModel.Container, _
ByVal _MainDefs As MainDefinitions, _
ByVal blnRO As Boolean, _
ByVal _Controls As Control.ControlCollection, ByVal _frmStandard As frmStandard, _
ByVal blnFilters As Boolean)
MyBase.New(strForm, "m_usr", _bs, _dg, _ta, _ds, _MainDefs, blnRO, _
"usr","",_Controls, _frmStandard, blnFilters)
_ta.Connection.ConnectionString = GetConnectionString()
Me.bsm_grp = New System.Windows.Forms.BindingSource(_components)
m_grpTableAdapter = New TheDataSetTableAdapters.m_grpTableAdapter
m_grpTableAdapter.Connection.ConnectionString = GetConnectionString()
Me.bsm_grp.DataMember = "m_grp"
Me.bsm_grp.DataSource = ds
Me.bsm_lang = New System.Windows.Forms.BindingSource(_components)
m_langTableAdapter = New TheDataSetTableAdapters.m_langTableAdapter
m_langTableAdapter.Connection.ConnectionString = GetConnectionString()
Me.bsm_lang.DataMember = "m_lang"
Me.bsm_lang.DataSource = ds
End Sub
Public Overrides Sub Createcolumns()
dgusr = New System.Windows.Forms.DataGridViewTextBoxColumn
dggrp = New System.Windows.Forms.DataGridViewComboBoxColumn
dglang = New System.Windows.Forms.DataGridViewComboBoxColumn
dgblocking = New System.Windows.Forms.DataGridViewCheckBoxColumn
dgname = New System.Windows.Forms.DataGridViewTextBoxColumn
dgemail = New System.Windows.Forms.DataGridViewTextBoxColumn
dgtelephone = New System.Windows.Forms.DataGridViewTextBoxColumn
dgcreatetime = New System.Windows.Forms.DataGridViewTextBoxColumn
End Sub
Public Overrides Sub Adjustcolumns(ByVal blnAdjustWidth As Boolean)
 MyBase.Adjustcolumns(blnAdjustWidth)
 Try
DefineColumn(dgusr, "usr", blnRO, ds.m_usr.usrColumn.MaxLength)
DefineComboBoxColumn(dggrp, MainDefs.strGetFormat("TYP_M_STRING"), True, "grp", "", FieldWidths.GENWIDTH, blnRO, true, "",  bsm_grp, "grp" ,"grp", Color.Lavender)
If blnRO = True Then dggrp.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
DefineComboBoxColumn(dglang, MainDefs.strGetFormat("TYP_M_LANG"), True, "lang", "", FieldWidths.GENWIDTH, blnRO, true, "",  bsm_lang, "lang" ,"lang", Color.Lavender)
If blnRO = True Then dglang.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
DefineColumn(dgblocking, "blocking", blnRO, ds.m_usr.blockingColumn.MaxLength)
DefineColumn(dgname, "name", blnRO, ds.m_usr.nameColumn.MaxLength)
DefineColumn(dgemail, "email", blnRO, ds.m_usr.emailColumn.MaxLength)
DefineColumn(dgtelephone, "telephone", blnRO, ds.m_usr.telephoneColumn.MaxLength)
DefineColumn(dgcreatetime, "createtime", blnRO, ds.m_usr.createtimeColumn.MaxLength)
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
Me.m_langTableAdapter.Fill(Me.ds.m_lang)
dg.CancelEdit()
iComboCount = 2
End sub
#Region "Editing"
Public Overrides Sub dg_UserDeletingRow(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs)
Try
Dim tadap As TheDataSetTableAdapters.m_usrTableAdapter
tadap = CType(ta, TheDataSetTableAdapters.m_usrTableAdapter)
tadap.Delete(e.Row.Cells(dg.Columns("usr").Index).Value.ToString())
MyBase.dg_UserDeletingRow(sender, e)
Catch ex As Exception
MsgBox("Delete failed. Most common cause is that record is in use in another table." + ex.message)
e.Cancel = True
End Try
End Sub
Private Sub dg_DefaultValuesNeeded(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dg.DefaultValuesNeeded
With e.Row
            '.Cells("grp").Value =  ('Commercial')
            '.Cells("blocking").Value =  ((1))
            '.Cells("name").Value =  ('')
            '.Cells("email").Value =  ('')
            '.Cells("telephone").Value =  ('')
            '.Cells("createtime").Value =  (getdate())
End With
End Sub
#End Region
#Region "Filter"
Public Overrides Sub CreateFilterBoxes(ByVal _Controls As Control.ControlCollection)
MyBase.CreateFilterBoxes(_Controls)
CreateAFilterBox(tbusrFind, "usr", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbgrpFind, "grp", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tblangFind, "lang", AddressOf tbFind_TextChanged, _Controls)
CreateACheckBox(cbblockingFind, "blocking", AddressOf cbFind_CheckChanged, _Controls)
CreateAFilterBox(tbnameFind, "name", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbemailFind, "email", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbtelephoneFind, "telephone", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbcreatetimeFind, "createtime", AddressOf tbFind_TextChanged, _Controls)
End Sub
Friend WithEvents tbusrFind As System.Windows.Forms.TextBox
Friend WithEvents tbgrpFind As System.Windows.Forms.TextBox
Friend WithEvents tblangFind As System.Windows.Forms.TextBox
Friend WithEvents cbblockingFind As System.Windows.Forms.CheckBox
Friend WithEvents tbnameFind As System.Windows.Forms.TextBox
Friend WithEvents tbemailFind As System.Windows.Forms.TextBox
Friend WithEvents tbtelephoneFind As System.Windows.Forms.TextBox
Friend WithEvents tbcreatetimeFind As System.Windows.Forms.TextBox
Private Sub cbFind_CheckChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
MakeFilter(False)
End Sub
Private Sub tbFind_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
 MakeFilter(False)
End Sub
#End Region
End Class
