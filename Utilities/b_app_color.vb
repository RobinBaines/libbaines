'------------------------------------------------
'Name: Module b_app_color.vb.
'Function: The colours the user may select in the frmAppParameters.
'Copyright Robin Baines 2008. All rights reserved.
'Notes: 
'Modifications:
'------------------------------------------------
Imports Utilities
Imports System.Windows.Forms
Imports System.Drawing
Public Class b_app_color
    Inherits dgColumns
    Friend WithEvents dgParameter As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgValueString As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgColor As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgRemark As System.Windows.Forms.DataGridViewTextBoxColumn

    Public Sub New(ByVal strForm As String, ByVal _bs As BindingSource, ByVal _dg As dgvEnter, _
    ByVal _ta As TheDataSetTableAdapters.b_app_colorTableAdapter, _
    ByVal _ds As DataSet, _
    ByVal _components As System.ComponentModel.Container, _
    ByVal _MainDefs As MainDefinitions, _
    ByVal blnRO As Boolean, ByVal blnActiveFilters As Boolean, _
    ByVal _Controls As Control.ControlCollection, ByVal _frmStandard As frmStandard)
        MyBase.New(strForm, "b_app_color", _bs, _dg, _ta, _ds, _MainDefs, blnRO, _
        "Parameter", "", _Controls, _frmStandard)
        _ta.Connection.ConnectionString = GetConnectionString()
    End Sub

    Public Overrides Sub Createcolumns()
        dgParameter = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgValueString = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgColor = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgRemark = New System.Windows.Forms.DataGridViewTextBoxColumn
    End Sub

    Public Overrides Sub Adjustcolumns(ByVal blnAdjustWidth As Boolean)
        Dim TheDataSet As TheDataSet = New TheDataSet
        DefineColumn(dgParameter, MainDefs.strGetFormat("VarChar"), True, "Parameter", "", FieldWidths.GENWIDTH, True, True, "", False, TheDataSet.b_app_parameter.ParameterColumn.MaxLength)
        DefineColumn(dgValueString, MainDefs.strGetFormat("NVarCharMax"), True, "ValueString", "", FieldWidths.GENWIDTH, True, True, "", False, TheDataSet.b_app_parameter.ValueStringColumn.MaxLength)

        'Dim blnShowcolor As Boolean = False
        DefineColumn(dgColor, "", False, "Color", "", FieldWidths.GENWIDTH, True, True, "", False, TheDataSet.b_app_parameter.ValueStringColumn.MaxLength)

        DefineColumn(dgRemark, MainDefs.strGetFormat("NVarCharMax"), True, "Remark", "", FieldWidths.GENWIDTH, True, True, "", False, TheDataSet.b_app_parameter.RemarkColumn.MaxLength)

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
    Private Sub dg_CellFormatting(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dg.CellFormatting
        'Dim strT As String
        Try
            If e.ColumnIndex >= 0 Then
                If sender.Columns(e.ColumnIndex).Name.ToString() = "Color" Then
                    '           strT = sender.Rows(e.RowIndex).Cells("ValueString").value.ToString.ToLower
                    'e.CellStyle.BackColor = Color.FromName(sender.Rows(e.RowIndex).Cells("ValueString").value) 'sender.Rows(e.RowIndex).Cells("ValueString").value
                    e.CellStyle.BackColor = Color.FromName(dg.Rows(e.RowIndex).Cells("ValueString").Value)
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Overrides Sub dg_UserDeletingRow(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs)
        Try
            Dim tadap As TheDataSetTableAdapters.b_app_parameterTableAdapter
            tadap = CType(ta, TheDataSetTableAdapters.b_app_parameterTableAdapter)
            tadap.Delete(e.Row.Cells(dg.Columns("Parameter").Index).Value.ToString())
            MyBase.dg_UserDeletingRow(sender, e)
        Catch ex As Exception
            MsgBox("Delete failed. Most common cause is that record is in use in another table." + ex.Message)
            e.Cancel = True
        End Try
    End Sub

    Private Sub dg_DefaultValuesNeeded(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dg.DefaultValuesNeeded
        With e.Row
            .Cells("ValueString").Value = ParseConstraint("('')")
        End With
    End Sub

#End Region
#Region "Filter"
    Public Overrides Sub CreateFilterBoxes(ByVal _Controls As Control.ControlCollection)
        MyBase.CreateFilterBoxes(_Controls)
        CreateAFilterBox(tbParameterFind, "Parameter", AddressOf tbParameterFind_TextChanged, _Controls)
        CreateAFilterBox(tbValueStringFind, "ValueString", AddressOf tbValueStringFind_TextChanged, _Controls)
        CreateAFilterBox(tbRemarkFind, "Remark", AddressOf tbRemarkFind_TextChanged, _Controls)
    End Sub

    Friend WithEvents tbParameterFind As System.Windows.Forms.TextBox
    Friend WithEvents tbValueStringFind As System.Windows.Forms.TextBox
    Friend WithEvents tbRemarkFind As System.Windows.Forms.TextBox
    Private Sub tbParameterFind_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbParameterFind.TextChanged
        MakeFilter(False)
    End Sub

    Private Sub tbValueStringFind_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbValueStringFind.TextChanged
        MakeFilter(False)
    End Sub

    Private Sub tbRemarkFind_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbRemarkFind.TextChanged
        MakeFilter(False)
    End Sub
#End Region
End Class
