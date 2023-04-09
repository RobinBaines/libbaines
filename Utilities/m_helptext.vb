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
Public Class m_helptext
    Inherits dgColumns
    Friend WithEvents dgForm As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgHelpText As System.Windows.Forms.DataGridViewTextBoxColumn
    Public Sub New(ByVal strForm As String, ByVal _bs As BindingSource, ByVal _dg As dgvEnter, _
  ByVal _ta As TheDataSetTableAdapters.m_form_helptextTableAdapter, _
  ByVal _ds As DataSet, _
  ByVal _components As System.ComponentModel.Container, _
  ByVal _MainDefs As MainDefinitions, _
  ByVal blnRO As Boolean, _
  ByVal _Controls As Control.ControlCollection, ByVal _frmStandard As frmStandard, _
  ByVal blnFilters As Boolean)
        MyBase.New(strForm, "m_helptext", _bs, _dg, _ta, _ds, _MainDefs, blnRO, _
        "form", "", _Controls, _frmStandard, blnFilters)
        _ta.Connection.ConnectionString = GetConnectionString()
    End Sub

    Public Overrides Sub Createcolumns()
        dgForm = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgHelpText = New System.Windows.Forms.DataGridViewTextBoxColumn
    End Sub

    Public Overrides Sub Adjustcolumns(ByVal blnAdjustWidth As Boolean)
        MyBase.Adjustcolumns(blnAdjustWidth)
        DefineColumn(dgForm, "form", True, ds.m_form_helptext.formColumn.MaxLength)
        DefineColumn(dgHelpText, "helptext", blnRO, ds.m_form_helptext.helptextColumn.MaxLength)
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
            Dim tadap As TheDataSetTableAdapters.m_form_helptextTableAdapter
            tadap = CType(ta, TheDataSetTableAdapters.m_form_helptextTableAdapter)
            tadap.Delete(e.Row.Cells(dg.Columns("form").Index).Value.ToString())
            MyBase.dg_UserDeletingRow(sender, e)
        Catch ex As Exception
            MsgBox("Delete failed. Most common cause is that record is in use in another table." + ex.Message)
            e.Cancel = True
        End Try
    End Sub

    Private Sub dg_DefaultValuesNeeded(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dg.DefaultValuesNeeded
        With e.Row
        End With
    End Sub

#End Region

#Region "Filter"
    Public Overrides Sub CreateFilterBoxes(ByVal _Controls As Control.ControlCollection)
        MyBase.CreateFilterBoxes(_Controls)
        CreateAFilterBox(tbFormFind, "form", AddressOf tbFind_TextChanged, _Controls)
        CreateAFilterBox(tbHelpText, "helptext", AddressOf tbFind_TextChanged, _Controls)
    End Sub

    Friend WithEvents tbFormFind As System.Windows.Forms.TextBox
    Friend WithEvents tbHelpText As System.Windows.Forms.TextBox
    
    Private Sub cbFind_CheckChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MakeFilter(False)
    End Sub

    Private Sub tbFind_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MakeFilter(False)
    End Sub

#End Region
End Class

