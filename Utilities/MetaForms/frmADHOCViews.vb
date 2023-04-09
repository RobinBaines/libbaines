'------------------------------------------------
'Name: Module frmADHOCViews.vb
'Function: The Adhoc views form.
'Copyright Baines 2013. All rights reserved.
'Notes: 
'Modifications: 
'------------------------------------------------
Imports Utilities
Imports System
Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Threading
Imports ExcelInterface.XMLExcelInterface
Imports System.ComponentModel
Imports System.IO
Imports System.Collections
Imports System.Windows.Forms
Imports System.Drawing
Public Class frmAdhocViews

    Friend WithEvents dgParent As dgvEnter

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
            If Not ADHOCTable Is Nothing Then
                ADHOCTable.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

#Region "new"
    Public Sub New(ByVal tsb As ToolStripItem _
              , ByVal strSecurityName As String, ByVal _MainDefs As MainDefinitions)

        MyBase.New(tsb, strSecurityName, _MainDefs)
        InitializeComponent()
        Me.dgParent = New dgvEnter()
        '
        'dgParent
        '
        Me.dgParent.blnDirty = False
        Me.dgParent.blnMeIsSource = False
        Me.dgParent.blnMove = True
        Me.dgParent.blnRO = False
        Me.dgParent.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgParent.Location = New System.Drawing.Point(12, Me.btnShowData.Location.Y + btnShowData.Height + 50)
        Me.dgParent.Name = "dgParent"
        Me.dgParent.Size = New System.Drawing.Size(300, 220)
        Me.dgParent.ta = Nothing
        Me.dgParent.TabIndex = 100

        Me.Controls.Add(Me.dgParent)
        Me.SwitchOffPrintDetail()
        Me.SwitchOffPrint()
        Me.SwitchOffUpdate()

        V_adhoc_viewsTableAdapter.Connection.ConnectionString = MainDefs.GetConnectionString()
        V_adhoc_view_columnsTableAdapter.Connection.ConnectionString = MainDefs.GetConnectionString()
    End Sub

#End Region

#Region "Load"
    Protected Overrides Sub frmLoad(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles MyBase.Load
        MyBase.frmLoad(sender, e)
        Try
            blnAllowUpdate = True
            FillTableAdapter()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Protected Overrides Sub FillTableAdapter()
        MyBase.FillTableAdapter()
        Try
            If blnAllowUpdate = True Then
                Me.V_adhoc_viewsTableAdapter.Fill(Me.MetaDataSet.v_adhoc_views)
                blnAllowUpdate = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub
#End Region

#Region "Scroll"
    Protected Overrides Sub frm_Layout(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LayoutEventArgs)
        MyBase.frm_Layout(sender, e)
        If TestActiveMDIChild() = True And Not ADHOCTable Is Nothing Then
            ADHOCTable.SetHeight(Me.ClientRectangle.Height)
        End If
    End Sub
#End Region

    'switching views so load the columns of the new view and reset the where clause.
    Private Sub lbAdhocViews_SelectedValueChanged(sender As Object, e As EventArgs) Handles lbAdhocViews.SelectedValueChanged
        If Not lbAdhocViews.SelectedValue Is Nothing Then
            Me.V_adhoc_view_columnsTableAdapter.FillBy(Me.MetaDataSet.v_adhoc_view_columns, lbAdhocViews.SelectedValue)

            Me.tbWhere.Text = ""
        End If
    End Sub

    Private Function CheckWhere(str As String)
        'If str.ToUpper.Contains("SELECT") Then
        '    Return False
        'End If
        If str.ToUpper.Contains("DROP") Then
            Return False
        End If
        If str.ToUpper.Contains("TRUNCATE") Then
            Return False
        End If
        If str.ToUpper.Contains("TABLE") Then
            Return False
        End If
        If str.ToUpper.Contains("DELETE") Then
            Return False
        End If
        If str.ToUpper.Contains("INSERT") Then
            Return False
        End If
        If str.ToUpper.Contains("UPDATE") Then
            Return False
        End If
        'If str.ToUpper.Contains("FROM") Then
        '    Return False
        'End If
        If str.ToUpper.Contains("VALUES") Then
            Return False
        End If
        Return True
    End Function

    Dim ADHOCTable As ADHOCTable
    Private bindingSource1 As BindingSource
    Dim dataAdapter As SqlDataAdapter
    Dim table As DataTable

    'create the ADHOCTable using the selected view.
    Private Sub btnShowData_Click(sender As Object, e As EventArgs) Handles btnShowData.Click
        Me.Cursor = Cursors.WaitCursor
        If Not ADHOCTable Is Nothing Then

            ADHOCTable.Dispose()
            ADHOCTable = Nothing
            dgParent.Columns.Clear()
            table.Dispose()
            'commandBuilder.Dispose()
            dataAdapter.Dispose()
            bindingSource1.Dispose()

            Me.vGrids.Clear()
        End If

        bindingSource1 = New BindingSource()
        Me.dgParent.DataSource = Me.bindingSource1

        ADHOCTable = New ADHOCTable(Me.Name, bindingSource1, dgParent, Nothing, Nothing, Nothing, MainDefs, True, Me.Controls, Me, True, _
                                    lbAdhocViews.SelectedValue, Me.MetaDataSet.v_adhoc_view_columns)

        ADHOCTable.CreateFilterBoxes()
        ADHOCTable.Adjustcolumns(True)
        ADHOCTable.AdjustFilterBoxes()
        ADHOCTable.FilterBoxesShow(True)
        'bindingSource1 = New BindingSource()
        'Me.dgParent.DataSource = Me.bindingSource1

        SetBindingNavigatorSource(bindingSource1)
        ' Create a new data adapter based on the specified query.
        Dim strSQL = "SELECT * FROM ADHOC." + lbAdhocViews.SelectedValue

        Dim blnRun As Boolean = True
        If Me.tbWhere.Text.Length Then
            If CheckWhere(tbWhere.Text) Then
                strSQL = strSQL + String.Format(" WHERE {0} ", tbWhere.Text)
                dataAdapter = New SqlDataAdapter(strSQL, ConnectionString.ConnectionString)
            Else
                blnRun = False
                MsgBox(statics.get_txt_header("The Where clause is illegal and could be injecting code. The view will not execute.", "User advice in frmADHOCViews.", "User information"))
            End If

        Else
            dataAdapter = New SqlDataAdapter(strSQL, ConnectionString.ConnectionString)
        End If

        If blnRun Then
            Try

                'Populate a new data table and bind it to the BindingSource.
                table = New DataTable()
                table.Locale = System.Globalization.CultureInfo.InvariantCulture
                dataAdapter.SelectCommand.CommandTimeout = 1800  'seconds
                dataAdapter.Fill(table)
                bindingSource1.DataSource = table
            Catch ex As Exception
                Dim strT As String = statics.get_txt_header("The ADHOC view did not execute correctly. Exception message was ", "User advice in frmADHOCViews.", "User information")
                MsgBox(strT + ex.Message)
            End Try
        End If

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub lbColumns_MouseClick(sender As Object, e As MouseEventArgs) Handles lbColumns.MouseClick
        'Clipboard.SetText(lbColumns.SelectedValue)  '(lbColumns.Items(lbColumns.SelectedIndex).ToString())
        tbWhere.Text = tbWhere.Text + lbColumns.SelectedValue
    End Sub

End Class

Public Class ADHOCTable
    Inherits dgColumns
    Implements IDisposable

    Protected Columns As New Dictionary(Of String, DataGridViewTextBoxColumn)

    Dim v_adhoc_view_columns As MetaDataSet.v_adhoc_view_columnsDataTable
    Public Sub New(ByVal strForm As String, ByVal _bs As BindingSource, ByVal _dg As dgvEnter, _
            ByVal _ta As MetaDataSetTableAdapters.v_adhoc_view_columnsTableAdapter, _
            ByVal _ds As DataSet, _
            ByVal _components As System.ComponentModel.Container, _
            ByVal _MainDefs As MainDefinitions, _
            ByVal blnRO As Boolean, _
            ByVal _Controls As Control.ControlCollection, ByVal _frmStandard As frmStandard, _
            ByVal blnFilters As Boolean, _
            tablename As String, _
            _v_adhoc_view_columns As MetaDataSet.v_adhoc_view_columnsDataTable)

        MyBase.New(strForm, tablename, _bs, _dg, _ta, _ds, _MainDefs, blnRO, _
        "", "", _Controls, _frmStandard, blnFilters)
        If Not ta Is Nothing Then
            ta.Connection.ConnectionString = GetConnectionString()
        End If
        v_adhoc_view_columns = _v_adhoc_view_columns
    End Sub


    Public Overrides Sub Createcolumns()

    End Sub

    Public Overrides Sub Adjustcolumns(ByVal blnAdjustWidth As Boolean)
        MyBase.Adjustcolumns(blnAdjustWidth)
        Try
            For Each field As MetaDataSet.v_adhoc_view_columnsRow In v_adhoc_view_columns
                Dim dg As DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
                Columns.Add(field.column_name, dg)
                Dim field_length As Integer = 100
                If Not field.IsCHARACTER_MAXIMUM_LENGTHNull Then
                    field_length = field.CHARACTER_MAXIMUM_LENGTH
                End If

                DefineColumn(dg, field.column_name, True, field_length)
            Next
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
    End Sub

#Region "Filter"
    Public Overrides Sub CreateFilterBoxes(ByVal _Controls As Control.ControlCollection)
        MyBase.CreateFilterBoxes(_Controls)
        For Each field As MetaDataSet.v_adhoc_view_columnsRow In v_adhoc_view_columns

            If field.DATA_TYPE = "bit" Then
                Dim tb As CheckBox = New System.Windows.Forms.CheckBox
                CreateACheckBox(tb, field.column_name, AddressOf cbFind_CheckChanged, _Controls)
            Else
                Dim tb As TextBox  '= New System.Windows.Forms.TextBox
                CreateAFilterBox(tb, field.column_name, AddressOf tbFind_TextChanged, _Controls)
            End If
        Next
    End Sub

    Private Sub cbFind_CheckChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MakeFilter(False)
    End Sub

    Private Sub tbFind_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MakeFilter(False)
    End Sub
#End Region

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        MyBase.ADispose()
        If Not Me.disposedValue Then
            If disposing Then
                For Each tb As Object In Columns
                    tb.Value.Dispose()
                Next
                Columns.Clear()
            End If
        End If
        Me.disposedValue = True
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class

