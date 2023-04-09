'------------------------------------------------
'Name: Module gen_TheDataSet_m_tble_column.vb.
'Function: 
'Copyright Robin Baines 2008. All rights reserved.
'Created 7/8/2012 12:00:00 AM.
'Notes: 
'Modifications:dgtble and dgcolmn changed from blnRO to true.
'------------------------------------------------
Imports Utilities
Imports System.Windows.Forms
Imports System.Drawing
Public Class TheDataSet_m_tble_column
Inherits dgColumns
    Friend WithEvents dgtble As System.Windows.Forms.DataGridViewTextBoxColumn
    Dim blnTbleVisible As Boolean = True
    Friend WithEvents dgcolmn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgwidth As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgformat As System.Windows.Forms.DataGridViewTextBoxColumn

    'Public bsm_tble As BindingSource
    'Friend WithEvents m_tbleTableAdapter As TheDataSetTableAdapters.m_tbleTableAdapter
    Public Sub New(ByVal strForm As String, ByVal _bs As BindingSource, ByVal _dg As dgvEnter, _
    ByVal _ta As TheDataSetTableAdapters.m_tble_columnTableAdapter, _
    ByVal _ds As DataSet, _
    ByVal _components As System.ComponentModel.Container, _
    ByVal _MainDefs As MainDefinitions, _
    ByVal blnRO As Boolean, _
    ByVal _Controls As Control.ControlCollection, ByVal _frmStandard As frmStandard, _
    ByVal blnFilters As Boolean)
        MyBase.New(strForm, "m_tble_column", _bs, _dg, _ta, _ds, _MainDefs, blnRO, _
        "tble", "colmn", _Controls, _frmStandard, blnFilters)
        _ta.Connection.ConnectionString = GetConnectionString()
        'Me.bsm_tble = New System.Windows.Forms.BindingSource(_components)
        'm_tbleTableAdapter = New TheDataSetTableAdapters.m_tbleTableAdapter
        'm_tbleTableAdapter.Connection.ConnectionString = GetConnectionString()
        'Me.bsm_tble.DataMember = "m_tble"
        'Me.bsm_tble.DataSource = ds
    End Sub
    Public Overrides Sub Createcolumns()
        dgtble = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgcolmn = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgwidth = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgformat = New System.Windows.Forms.DataGridViewTextBoxColumn

    End Sub
    Public Overrides Sub Adjustcolumns(ByVal blnAdjustWidth As Boolean)
        MyBase.Adjustcolumns(blnAdjustWidth)
        'DefineComboBoxColumn(dgtble, MainDefs.strGetFormat("TYP_M_STRING"), True, "tble", "", FieldWidths.GENWIDTH, True, True, "", bsm_tble, "tble", "tble", Color.Lavender)
        'If blnRO = True Then dgtble.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
        DefineColumn(dgtble, "tble", True, ds.m_tble_column.tbleColumn.MaxLength)
        blnTbleVisible = dgtble.Visible
        DefineColumn(dgcolmn, "colmn", True, ds.m_tble_column.colmnColumn.MaxLength)
        DefineColumn(dgwidth, "width", blnRO, ds.m_tble_column.widthColumn.MaxLength)
        DefineColumn(dgformat, "format", blnRO, ds.m_tble_column.formatColumn.MaxLength)
        PutColumnsInGrid()
        AdjustDataGridWidth(blnAdjustWidth)
        RefreshCombos()
    End Sub
    Public Overrides Sub RefreshCombos()
        MyBase.RefreshCombos()
        ' Me.m_tbleTableAdapter.Fill(Me.ds.m_tble)
        dg.CancelEdit()
        iComboCount = 0
    End Sub
#Region "Editing"
    Public Overrides Sub dg_UserDeletingRow(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs)
        Try
            Dim tadap As TheDataSetTableAdapters.m_tble_columnTableAdapter
            tadap = CType(ta, TheDataSetTableAdapters.m_tble_columnTableAdapter)
            tadap.Delete(e.Row.Cells(dg.Columns("tble").Index).Value.ToString(), e.Row.Cells(dg.Columns("colmn").Index).Value.ToString())
            MyBase.dg_UserDeletingRow(sender, e)
        Catch ex As Exception
            MsgBox("Delete failed. Most common cause is that record is in use in another table." + ex.message)
            e.Cancel = True
        End Try
    End Sub
    Private Sub dg_DefaultValuesNeeded(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dg.DefaultValuesNeeded
        With e.Row
            .Cells("width").Value = ((100))
        End With
    End Sub
#End Region
#Region "Filter"
    Public Overrides Sub CreateFilterBoxes(ByVal _Controls As Control.ControlCollection)
        MyBase.CreateFilterBoxes(_Controls)

        CreateAFilterBox(tbcolmnFind, "colmn", AddressOf tbFind_TextChanged, _Controls)
        CreateAFilterBox(tbwidthFind, "width", AddressOf tbFind_TextChanged, _Controls)
        CreateAFilterBox(tbformatFind, "format", AddressOf tbFind_TextChanged, _Controls)
        CreateAFilterBox(tbtbleFind, "tble", AddressOf tbFind_TextChanged, _Controls)
    End Sub
    Friend WithEvents tbtbleFind As System.Windows.Forms.TextBox
    Friend WithEvents tbcolmnFind As System.Windows.Forms.TextBox
    Friend WithEvents tbwidthFind As System.Windows.Forms.TextBox
    Friend WithEvents tbformatFind As System.Windows.Forms.TextBox
    Private Sub cbFind_CheckChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MakeFilter(False)
    End Sub
    Private Sub tbFind_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MakeFilter(False)
    End Sub
#End Region

    Private Sub dg_DataBindingComplete(ByVal sender As Object, ByVal e As DataGridViewBindingCompleteEventArgs) Handles dg.DataBindingComplete
        With dg
            Try
                ' MessageBox.Show(dgtble.Visible.ToString(), "dgtble")
                .Columns("tble").Visible = blnTbleVisible
            Catch ex As Exception
            End Try
        End With
    End Sub
    Private Sub dataGridView1_CellStateChanged(ByVal sender As Object, _
    ByVal e As DataGridViewCellStateChangedEventArgs) _
    Handles dg.CellStateChanged

        Dim state As DataGridViewElementStates = e.StateChanged
        Dim msg As String = String.Format( _
            "Row {0}, Column {1}, {2}, {3}", _
            e.Cell.RowIndex, dg.Columns(e.Cell.ColumnIndex).Name, e.StateChanged, dg.Columns(e.Cell.ColumnIndex).Visible)
        '    MessageBox.Show(msg, "Cell State Changed")

    End Sub
End Class
