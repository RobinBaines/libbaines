'------------------------------------------------
'Name: Module dgColumns.vb.
'Function: The placeholder for the columns and filter boxes of a DataGridView.
'Copyright Robin Baines 2006. All rights reserved.
'20091124 RPB added error handling to Handle_DataGridView_RowValidating
'for cases where data gets truncated.
'20091124 RPB created blnStringHasWildCards.
'20091124 RPB modified MakeFilter. If the string contains a Like wildcard then do not use Like.
'20091209 RPB modified DefineColumn when setting bold font to use dg default style and not the tb.
'20091124 RPB Modified MakeFilter: Do not use Like if string contains a wildcard character.
'20091209 RPB Modified MakeFilter: Allow filter on DateTime by converting to string.
'20100105 RPB Modified Handle_DataGridView_RowValidating 
'Dropped this check on null column for RAP.Additives.
'But check this in other apps.
'20100106 RPB added filter function HideFilter() for via.
'20100114 RPB modified MakeFilter(): To be independent of the regional settings the filter value has to be in the
'MS SQL format  mm/dd/yyyy
'20100122 RPB Modified blnStringHasWildCards: Do not check for * or _.
'20100122 RPB Modified MakeFilter: Only add * if there is no * in the string already.
'20100122 RPB modified DefineColumn by commenting out rightalignment on format as this is 
'done much better with AdjustHorizontalAlignment() when columns are put in the grid.
'20100122 RPB created AdjustHorizontalAlignment()
'20100122 RPB modified PutColumnsInGrid by making call to AdjustHorizontalAlignment()
'20100122 RPB created AdjustFilterTextBox to do alignment and RO for filter boxes.
'20100220 Validate returns true if insert took place.
'20100222 Added storing the first row displayed in grid to restore after refresh.
'20100222 RPB added frmParent to a new new and then use if update/insert fails to refresh the form.
'20110216 RPB added [] to Fieldname in context  Dim strDataPropertyName = "[" & GetBoundColumnName(tbEntry.Value) & "]"
'This is necessary because the LOP project has fieldnames which start with an integer.
'20120708 RPB removed blnUseGroupBox and old New().
'20120708 RPB solved problem with bad positioning of the GroupBox for filters after, decreasing size of form
'to less than the datagrid width then scrolling right in the grid and then increasinf size of form.
'20140419 RPB modified GotoRow() If iIndex > 0  to  If iIndex > -1 because it was missing the first row.
'20170302 Moved Controls.Add(Me.gbForFiltersGroupBox) to Adjustcolumns.
'2070203 Removed Overridable GetConnectionString as this is being called in the constructor and is not overidden.
'20180120 GetBoundColumnType: When called from Adhoc table the ValueType may be nothing.
'20181231 AdjustFilterTextBox: removed readonly True so that user can remove after double clicking.
'20200102 Dispose of the columns in ADispose.
'20200123 In some circumstances functions may be called when the form is being disposed. So check dgSortedListOfColumns before disposing the columns.
'20200222 So check gbForFiltersGroupBox before disposing the columns.
'------------------------------------------------
Imports Utilities
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Text
Public Class dgColumns

    Dim blnFilters As Boolean = False
    Protected MainDefs As MainDefinitions
    Public WithEvents dg As DataGridView
    Protected bs As BindingSource
    Friend WithEvents __ta As Object
    Friend WithEvents __ds As Object    'TheDataSet
    Protected FindTbs As New Dictionary(Of Control, String)

    '20200102 Not used
    'Protected ut As New Utilities

    Protected blnRO As Boolean

    '20100222
    Protected frmParent As frmStandard

    Dim strForm As String
    Dim strTable As String

    'Fields for restoring row after sort.
    Dim strSortFieldName As String 'If this is defined in the constructor then the row selected before the sort
    'will be selected afterwards.
    Dim strSortFieldName2 As String 'If this is defined in the constructor then the row selected before the sort
    'will be selected afterwards.
    Dim strPrj As String = ""   'Store the value of the strSortFieldName at the current row so that it can be found again after a sort.
    Dim strPrj2 As String = ""   'Store the value of the strSortFieldName at the current row so that it can be found again after a sort.

    Friend WithEvents gbForFilters As System.Windows.Forms.GroupBox = New System.Windows.Forms.GroupBox
    Friend WithEvents gbForFiltersGroupBox As System.Windows.Forms.GroupBox = New System.Windows.Forms.GroupBox
    Dim iSequence As Integer
    Dim dgSortedListOfColumns As SortedList(Of Integer, DataGridViewColumn)

    'Friend WithEvents dgId As System.Windows.Forms.DataGridViewTextBoxColumn
    Private _iComboCount As Integer

    'Use to store the first row displayed in grid and to restore after refresh.
    Dim iParentFirstDisplayedRowIndex As Integer

    'Set this flag to indicate that a refresh is taking place.
    Dim blnJustRefreshing As Boolean = False

    'Used to store position before reloading a table.
    Protected iIndex As Integer = -1
    Protected iCount As Integer
    Protected iColumnIndex As Integer

    Protected idgIndex As Integer
    Protected idgCount As Integer

    Protected Controls As Control.ControlCollection

    'This will be set to true if DefineColumn stored column data.
    Protected blnStoredColumnData As Boolean
    Dim dgInitialWidth As Integer
    Const iCONSTROWHEADERWIDTH = 21
    Dim Color_Unfocussed As Color = Color.DarkSlateGray

    '20120906 Associate a label with the grid.
    Dim dgLabel As Label = Nothing

    Const TEXTBOX_Y_IN_FILTER_GROUPBOX = 10
    Const TEXTBOX_HEIGHT_IN_FILTER_GROUPBOX = 20

    Public ReadOnly Property ReadOnlyBackGroundColor() As System.Drawing.Color
        Get
            Return statics.ReadOnlyBackGroundColor
        End Get
    End Property

    Public Sub New(ByVal _ParentForm As MainForm)
        MyBase.New()
    End Sub
    Protected Enum FieldWidths
        FLOATWIDTH = 65
        MONTHWIDTH = 100
        REMARKWIDTH = 120
        GENWIDTH = 100
        SMALLWIDTH = 35
        BOOLWIDTH = 40
    End Enum

#Region "ParseConstraint"
    'Called to get default values for fields so use Object to cope with integers and strings.
    'Alternative is to try and get the insert not to include fields which are still null
    'and let the default value in sql do it; but that is complicated on fields which a 
    'user may enter.
    Public Function ParseConstraint(ByVal strConstraint As String) As Object
        Dim strRet As Object = Nothing
        If strConstraint = "(N'INSTR')" Then
            strRet = "INSTR"
        End If
        If strConstraint = "('N')" Then
            strRet = "N"
        End If
        If strConstraint = "('INSTR')" Then
            strRet = "INSTR"
        End If
        If strConstraint = "((0))" Then
            strRet = 0
        End If
        If strConstraint = "((1))" Then
            strRet = 1
        End If
        If strConstraint = "('')" Then
            strRet = ""
        End If
        Return strRet
    End Function
#End Region

#Region "Connection"
    'Public Overridable Function GetConnectionString() As String
    '2070203 Removed Overridable GetConnectionString as this is being called in the constructor and is not overidden.
    Public Function GetConnectionString() As String
        Return Me.MainDefs.GetConnectionString()    '(My.Settings.QualityDatabase)
    End Function

    Public Function GetConnectionString(ByVal blnQuality As Boolean) As String
        Return Me.MainDefs.GetConnectionString(blnQuality)
    End Function
#End Region

#Region "Property"
    Public Property iComboCount() As Integer
        Get

            Return _iComboCount
        End Get
        Set(ByVal value As Integer)
            _iComboCount = value
        End Set
    End Property
    Public Property ta() As Object
        Get

            Return __ta
        End Get
        Set(ByVal value As Object)
            __ta = value
        End Set
    End Property
    Public Property ds() As Object ' TheDataSet
        Get

            Return __ds
        End Get
        Set(ByVal value As Object)  'TheDataSet)
            __ds = value
        End Set
    End Property
#End Region

#Region "New"
    'The filters are added to a groupbox which is also added to another groupbox so that the outer group box can be scrolled.
    Private Sub CreateFilterGroupBox()
        gbForFilters.Controls.Clear()
        gbForFiltersGroupBox.Controls.Clear()

        'This is the location in the outer group box. By using 0,0 the outer box is not visible.
        Me.gbForFilters.Location = New System.Drawing.Point(0, 0)
        Me.gbForFilters.Name = "gbForFilters"
        Me.gbForFilters.Size = New System.Drawing.Size(1532, statics.GroupBoxRelativeVerticalLocation)
        Me.gbForFilters.TabIndex = 28
        Me.gbForFilters.TabStop = False
        Me.gbForFiltersGroupBox.Controls.Add(gbForFilters)

        'Search on this below because the location of this gb is syncht with the dg.
        Me.gbForFiltersGroupBox.Location = New System.Drawing.Point(dg.Location.X, dg.Location.Y - statics.GroupBoxRelativeVerticalLocation)
        Me.gbForFiltersGroupBox.Name = statics.FilterGroupBoxName
        Me.gbForFiltersGroupBox.Size = New System.Drawing.Size(1517, statics.GroupBoxRelativeVerticalLocation)
        Me.gbForFiltersGroupBox.TabIndex = 29
        Me.gbForFiltersGroupBox.TabStop = False
        Me.gbForFiltersGroupBox.Controls.Add(gbForFilters)
    End Sub

    Private Sub Init(ByVal _bs As BindingSource, ByVal _dg As DataGridView, _
    ByVal _ta As Object, _
    ByVal _ds As DataSet, _
    ByVal _MainDefs As MainDefinitions _
        , ByVal _blnRO As Boolean)

        iSequence = 0
        MainDefs = _MainDefs
        __ds = _ds
        __ta = _ta
        dg = _dg
        bs = _bs
        blnRO = _blnRO
        dg.AutoGenerateColumns = False
        dg.DataSource = bs
        dg.AllowUserToAddRows = Not blnRO
        dg.AllowUserToDeleteRows = Not blnRO
        dg.ScrollBars = ScrollBars.Both
        dg.ReadOnly = blnRO
        dg.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        dg.RowHeadersWidth = iCONSTROWHEADERWIDTH
        iComboCount = 0

        blnStoredColumnData = False
        dgInitialWidth = 0

        CreateFilterGroupBox()
        gbForFiltersGroupBox.Visible = False

        dg.ForeColor = Color_Unfocussed

        '20100224 Put the name of the grid in the ToolTip top left.         '
        dg.TopLeftHeaderCell.ToolTipText = strTable
        Createcolumns()

        '20100327 Store the table adapater in the dgvEnter version of DataGridView. This is then 
        'used by the tsbUpdate button UpdateData() call in the frmStandard.
        Dim dgv As dgvEnter = TryCast(dg, dgvEnter)
        If Not dgv Is Nothing Then
            dgv.ta = ta
        End If
    End Sub

    Public Sub New(ByVal _strForm As String, _
        ByVal _strTable As String, _
        ByVal _bs As BindingSource, ByVal _dg As DataGridView, _
        ByVal _ta As Object, _
        ByVal _ds As DataSet, _
        ByVal _MainDefs As MainDefinitions _
       , ByVal _blnRO As Boolean _
       , ByVal _strSortFieldName As String, ByVal _strSortFieldName2 As String _
       , ByVal _Controls As Control.ControlCollection)

        '20120717 remove fast key prefix from security names.
        strForm = _strForm.Replace("&", "")
        strTable = _strTable
        strSortFieldName = _strSortFieldName
        strSortFieldName2 = _strSortFieldName2
        dgSortedListOfColumns = New SortedList(Of Integer, DataGridViewColumn)
        frmParent = Nothing
        Init(_bs, _dg, _ta, _ds, _MainDefs, _blnRO)
        Controls = _Controls

    End Sub

    '20100222 RPB added frmParent to this new new.
    Public Sub New(ByVal _strForm As String, _
        ByVal _strTable As String, _
        ByVal _bs As BindingSource, ByVal _dg As DataGridView, _
        ByVal _ta As Object, _
        ByVal _ds As DataSet, _
        ByVal _MainDefs As MainDefinitions _
       , ByVal _blnRO As Boolean _
       , ByVal _strSortFieldName As String, ByVal _strSortFieldName2 As String _
       , ByVal _Controls As Control.ControlCollection _
       , ByVal _frmParent As frmStandard)

        strForm = _strForm
        strTable = _strTable
        strSortFieldName = _strSortFieldName
        strSortFieldName2 = _strSortFieldName2
        dgSortedListOfColumns = New SortedList(Of Integer, DataGridViewColumn)
        frmParent = _frmParent
        Init(_bs, _dg, _ta, _ds, _MainDefs, _blnRO)
        Controls = _Controls

    End Sub

    Private Sub New201310(ByVal _strForm As String, _
        ByVal _strTable As String, _
        ByVal _bs As BindingSource, ByVal _dg As DataGridView, _
        ByVal _ta As Object, _
        ByVal _ds As DataSet, _
        ByVal _MainDefs As MainDefinitions _
       , ByVal _blnRO As Boolean _
       , ByVal _strSortFieldName As String, ByVal _strSortFieldName2 As String _
       , ByVal _Controls As Control.ControlCollection _
       , ByVal _frmParent As frmStandard, ByVal _blnFilters As Boolean)

        strForm = _strForm
        strTable = _strTable
        strSortFieldName = _strSortFieldName
        strSortFieldName2 = _strSortFieldName2
        blnFilters = _blnFilters
        dgSortedListOfColumns = New SortedList(Of Integer, DataGridViewColumn)
        frmParent = _frmParent
        Init(_bs, _dg, _ta, _ds, _MainDefs, _blnRO)
        Controls = _Controls

        If Not frmParent Is Nothing Then
            frmParent.RegisterGrid(Me)
        End If
    End Sub
    '20100222 RPB added frmParent to this new new.
    '20120210 RPB added blnFilters.
    Public Sub New(ByVal _strForm As String, _
        ByVal _strTable As String, _
        ByVal _bs As BindingSource, ByVal _dg As DataGridView, _
        ByVal _ta As Object, _
        ByVal _ds As DataSet, _
        ByVal _MainDefs As MainDefinitions _
       , ByVal _blnRO As Boolean _
       , ByVal _strSortFieldName As String, ByVal _strSortFieldName2 As String _
       , ByVal _Controls As Control.ControlCollection _
       , ByVal _frmParent As frmStandard, ByVal _blnFilters As Boolean)

        New201310(_strForm, _strTable, _bs, _dg, _ta, _ds, _MainDefs, _blnRO, _strSortFieldName, _strSortFieldName2, _Controls, _frmParent, _blnFilters)
    End Sub

    Public Overridable Sub ADispose()
        FindTbs.Clear()

        '20190506 Added these so that ADHOC views works correctly when switching to a new ADHOC view.
        If Not gbForFilters Is Nothing Then
            gbForFilters.Controls.Clear()
            gbForFilters.Dispose()
            gbForFilters = Nothing
        End If
        If Not gbForFiltersGroupBox Is Nothing Then
            gbForFiltersGroupBox.Controls.Clear()
            gbForFiltersGroupBox.Dispose()
            gbForFiltersGroupBox = Nothing
        End If

        '20200102 Dispose of the columns in ADispose.
        '20200123 In some circumstances this function may be called when the form is being disposed. So check dgSortedListOfColumns before disposing the columns.
        If Not dgSortedListOfColumns Is Nothing Then
            For Each col As DataGridViewColumn In dgSortedListOfColumns.Values
                col.Dispose()
                col = Nothing
            Next col
            dgSortedListOfColumns.Clear()
            dgSortedListOfColumns = Nothing
        End If
    End Sub

    Public Overridable Sub Createcolumns()
        ' dgId = New System.Windows.Forms.DataGridViewTextBoxColumn
    End Sub

    Public Overridable Sub CreateFilterBoxes()

        'Is not always called from old code.
        If blnFilters = True Then
            CreateFilterBoxes(dg.Controls)
        End If
    End Sub

    Public Overridable Sub CreateFilterBoxes(ByVal Controls As Control.ControlCollection)
        gbForFiltersGroupBox.Visible = True
    End Sub

    Public Overridable Sub FilterBoxesShow(ByVal blnVisible As Boolean)
        gbForFiltersGroupBox.Visible = blnVisible
    End Sub

    '20120910 switch all on and off
    Public Overridable Sub Visible(ByVal blnVisible As Boolean)
        'Is not always called from old code.
        dg.Visible = blnVisible
        If Not dgLabel Is Nothing Then
            dgLabel.Visible = blnVisible
        End If
        If blnFilters = True Then
            gbForFiltersGroupBox.Visible = blnVisible
        End If
    End Sub

    Public Sub SetdgLabel(lab As Label)
        dgLabel = lab

        'translate the label
        dgLabel.Text = statics.get_txt_header(dgLabel.Text, "datagridview  label", strForm)
        AdjustdgLabelPosition()
    End Sub

    Public Sub SetdgLabel(strT As String)
        dgLabel = New Label
        dgLabel.Text = strT
        dg.Parent.Controls.Add(dgLabel)
        dgLabel.Width = dg.Width
        'translate the label
        dgLabel.Text = statics.get_txt_header(dgLabel.Text, "datagridview  label", strForm)
        AdjustdgLabelPosition()
    End Sub

#End Region

#Region "Load"
    Public Overridable Sub Fill(ByVal table As Object)
        Me.StoreRowIndexWithFocus()
        ta.fill(table)
        Me.ResetFocusRow()
    End Sub

    Public Overridable Sub Adjustcolumns(ByVal blnAdjustWidth As Boolean)
        dgSortedListOfColumns.Clear()
        statics.strLastForm = ""
        statics.strLastTble = ""
        If Not Controls Is Nothing Then
            Controls.Add(Me.gbForFiltersGroupBox)
        End If

        'DefineColumn(dgId, "", True, "Id", "", 100, blnRO, True, MainDefs.DONOTPRINT, False)
    End Sub
    Public Overridable Sub RefreshCombos()

        StoreRowIndexWithFocus()
        iComboCount = 0
    End Sub

    Public Overridable Sub ResetCursor()
        ResetFocusRow()
    End Sub

    ' Store the row position in a grid before re-loading data. Use ResetFocusRow() to give the row the focus again.
    Public Overridable Sub StoreRowIndexWithFocus()

        '20100222 Store the display index
        blnJustRefreshing = True

        Try
            iParentFirstDisplayedRowIndex = dg.FirstDisplayedScrollingRowIndex
        Catch ex As Exception
        End Try

        'Store position in dg.
        If Not dg.CurrentRow Is Nothing Then
            Try
                '20140110 Sometimes currentrow is 0 while a row has been selected.
                'CurrentRow does not necessarily change with the selection of the row header.
                'SelectedRow does but it is not set if user selects a cell in a row.
                iIndex = dg.CurrentRow.Index
                If iIndex = 0 Then
                    If Not dg.SelectedRows.Count = 0 Then
                        iIndex = dg.SelectedRows(0).Index
                    Else
                        iIndex = -1
                    End If
                End If

                iColumnIndex = dg.CurrentCell.ColumnIndex
                iCount = dg.Rows.Count
            Catch ex As Exception
            End Try
        End If
        iComboCount = 0
    End Sub

    Public ReadOnly Property GetFocusRowIndex() As Integer
        Get
            Return iIndex
        End Get
    End Property

    '20131021 Used to set the current row if the row header is clicked.
    'dg.Rows(e.RowIndex).Selected = True is different. It just selects the row (or rows!) but does not set the current unless FullRowSelect is set.
    Protected Sub SetCurrentRow(RowIndex As Integer)
        For iC As Integer = 0 To dg.ColumnCount - 1
            Dim cCell As DataGridViewCell = dg.Rows(RowIndex).Cells(iC)
            If cCell.Visible Then
                If Not cCell Is Nothing Then
                    If cCell.Visible = True Then
                        dg.CurrentCell = cCell
                    End If
                End If
                Exit For
            End If
        Next
    End Sub

    ' See StoreRowIndexWithFocus.
    Public Overridable Sub ResetFocusRow()
        Dim cCell As DataGridViewCell = Nothing
        Try

            If iIndex <> -1 Then
                'MOD RPB 20080902 Remember that the number of records in refreshed data may have increased and
                'as they are added from the top the iIndex needs to be adjusted.
                ' iCount = dg.RowCount - iCount
                iCount = dg.Rows.Count - iCount
                'For Each cC As DataGridViewColumn In dg.Columns
                '    If cC.Visible = True And (iIndex + iCount) < dg.Rows.Count And (iIndex + iCount) > 0 Then
                '        cCell = dg.Rows(iIndex + iCount).Cells(cC.Index)
                '        Exit For
                '    End If
                'Next
                If (iIndex + iCount) < dg.Rows.Count And (iIndex + iCount) >= 0 Then
                    cCell = dg.Rows(iIndex + iCount).Cells(iColumnIndex)

                    If Not cCell Is Nothing Then
                        If cCell.Visible = True Then
                            dg.CurrentCell = cCell
                            If dg.SelectionMode = DataGridViewSelectionMode.CellSelect Then
                                dg.CurrentCell.Selected = True
                            Else
                                dg.CurrentRow.Selected = True
                            End If
                        End If
                    End If
                End If
                iIndex = -1
            End If

            '20101206 RPB put catch around the dg.CurrentCell = cCell as this fails if the cell is not visible.
        Catch ex As Exception
        End Try
        blnJustRefreshing = False
    End Sub

    '20140419 Modified  If iIndex > 0  to  If iIndex > -1 because it was missing the first row.
    Public Overridable Function GotoRow(iIndex As Integer, iColIndex As Integer) As Boolean
        Dim cCell As DataGridViewCell = Nothing
        Try
            If iIndex > -1 And iIndex < dg.RowCount And iColIndex > -1 And iColIndex < dg.ColumnCount Then
                cCell = dg.Rows(iIndex).Cells(iColIndex)
                If Not cCell Is Nothing Then
                    If cCell.Visible = True Then
                        dg.CurrentCell = cCell
                        dg.CurrentCell.Selected = True
                        Return True
                    Else
                        '20140424 If cell is not visible find the first visible one.
                        Dim strVisibleCol As String = GetFirstVisibleColumn()
                        dg.CurrentCell = dg.Rows(iIndex).Cells(strVisibleCol)

                        '20140807 GIS needs to goto the row even if column is not visible.
                        dg.CurrentCell.Selected = True
                        Return True
                    End If
                End If
            End If
        Catch ex As Exception
        End Try
        Return False
    End Function

    'strColumnName must be visible for this to work.
    Public Overridable Function FindAndGotoRow(strColumnName As String, strColumnValue As String) As Boolean
        Try
            Dim iColIndex = dg.Columns(strColumnName).Index
            If iColIndex > -1 Then
                Dim iIndex = bs.Find(strColumnName, strColumnValue)
                If iIndex >= 0 Then
                    If iIndex = dg.CurrentRow.Index Then
                        Return True
                    Else
                        Return GotoRow(iIndex, iColIndex)
                    End If
                End If
            End If
        Catch ex As Exception
        End Try
        Return False
    End Function

    Private Sub dg_CellFormatting(sender As Object, e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dg.CellFormatting

    End Sub

    '20100222 Restore the display index
    Private Sub dg_DataBindingComplete(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewBindingCompleteEventArgs) Handles dg.DataBindingComplete
        Try
            If blnJustRefreshing = True And iParentFirstDisplayedRowIndex >= 0 Then
                If iParentFirstDisplayedRowIndex < dg.RowCount Then
                    dg.FirstDisplayedScrollingRowIndex = iParentFirstDisplayedRowIndex
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

    '20120307 Add a width parameter. This value was set to 20.
    Private Function GetDataGridWidth(ByVal iwidth As Integer) As Integer
        Dim iW As Integer = iwidth ' 20 'for scroll width '* 3
        If dg.RowHeadersVisible = True Then
            iW += dg.RowHeadersWidth
        End If

        For Each c As Object In dg.Columns
            If c.visible = True Then
                iW = iW + c.width
            End If
        Next
        Return iW
    End Function
    Public Sub ResetFillout()

        '20100305 Reset the fill out.
        Dim iC As Integer = dg.ColumnCount - 1
        Do While iC >= 0
            If dg.Columns(iC).AutoSizeMode <> DataGridViewAutoSizeColumnMode.NotSet Then
                Try
                    dg.Columns(iC).AutoSizeMode = DataGridViewAutoSizeColumnMode.NotSet
                Catch ex As Exception

                End Try
            End If
            iC = iC - 1
        Loop
    End Sub
    Public Sub Fillout(ByVal _AutoSizeMode As DataGridViewAutoSizeColumnMode)

        '20081002 Fill out with the last visible column
        ResetFillout()
        Dim iC As Integer = dg.ColumnCount - 1
        If _AutoSizeMode <> DataGridViewAutoSizeColumnMode.NotSet Then
            iC = dg.ColumnCount - 1
            Do While iC >= 0
                If dg.Columns(iC).Visible = True Then
                    If dg.Columns(iC).AutoSizeMode <> _AutoSizeMode Then
                        Try
                            dg.Columns(iC).AutoSizeMode = _AutoSizeMode
                        Catch ex As Exception
                        End Try
                    End If
                    Exit Do
                End If
                iC = iC - 1
            Loop
        End If
    End Sub

    '20120307 added version to explicitly add in the width of a vertical scrollbar.
    Public Sub AdjustDataGridWidth(ByVal blnFillOut As Boolean, ByVal blnIncludeVerticalScrollBar As Boolean)

        If blnIncludeVerticalScrollBar = True Then
            dg.Width = GetDataGridWidth(Windows.Forms.SystemInformation.VerticalScrollBarWidth())
        Else
            dg.Width = GetDataGridWidth(0)
        End If

        dgInitialWidth = dg.Width
        If blnFillOut = True Then
            Fillout(DataGridViewAutoSizeColumnMode.Fill)
        End If

        'Adjust the position of the filter group boxes.
        If Not gbForFiltersGroupBox Is Nothing Then
            Me.gbForFiltersGroupBox.Location = New System.Drawing.Point(dg.Location.X, dg.Location.Y - statics.GroupBoxRelativeVerticalLocation)
        End If
    End Sub

    Public Sub AdjustDataGridWidth(ByVal blnFillOut As Boolean)

        dgInitialWidth = GetDataGridWidth(20)
        dg.Width = dgInitialWidth
        If blnFillOut = True Then
            Fillout(DataGridViewAutoSizeColumnMode.Fill)
        End If

        'Adjust the position of the filter group boxes.
        If Not gbForFiltersGroupBox Is Nothing Then
            Me.gbForFiltersGroupBox.Location = New System.Drawing.Point(dg.Location.X, dg.Location.Y - statics.GroupBoxRelativeVerticalLocation)
        End If
    End Sub
    'Adjust position of the grid.
    Public Sub AdjustPosition(ByVal p As Point)

        '20130127 test blnFilters instead of visibility because gb may be not visible because dgColumn is temporarily invisible.
        If blnFilters = True Then
            'If gbForFiltersGroupBox.Visible = True Then
            If Not gbForFiltersGroupBox Is Nothing Then
                Me.gbForFiltersGroupBox.Location = p
                dg.Location = New System.Drawing.Point(p.X, p.Y + statics.GroupBoxRelativeVerticalLocation)
            End If
        Else
            dg.Location = p
        End If
        AdjustdgLabelPosition()
    End Sub
    'Adjust horizontal position of the grid.
    Const dgLABELOFFSET = 2
    Private Sub AdjustdgLabelPosition()
        If Not dgLabel Is Nothing Then
            dgLabel.Location = New System.Drawing.Point(dg.Location.X, dg.Location.Y + dg.Height + dgLABELOFFSET)
        End If
    End Sub

    '20130923 Adjusted to allow for filters.
    Public Sub AdjustPosition(ByVal vLeft As Control)
        If Not gbForFiltersGroupBox Is Nothing Then
            Dim iY As Integer = vLeft.Location.Y
            If Me.blnFilters = True Then
                iY = iY + gbForFiltersGroupBox.Height - 2 * gbForFiltersGroupBox.Margin.Top 'statics.GroupBoxRelativeVerticalLocation
            End If
            dg.Location = New System.Drawing.Point(vLeft.Location.X + vLeft.Width + 10, iY)
            Me.gbForFiltersGroupBox.Location = New System.Drawing.Point(dg.Location.X, dg.Location.Y - statics.GroupBoxRelativeVerticalLocation)
        End If
        AdjustdgLabelPosition()
    End Sub

    Public Sub AdjustPosition(ByVal vLeft As dgColumns)
        dg.Location = New System.Drawing.Point(vLeft.dg.Location.X + vLeft.dg.Width + 10, dg.Location.Y)
        If Not gbForFiltersGroupBox Is Nothing Then
            Me.gbForFiltersGroupBox.Location = New System.Drawing.Point(dg.Location.X, dg.Location.Y - statics.GroupBoxRelativeVerticalLocation)
        End If
        AdjustdgLabelPosition()
    End Sub

    Public Sub AdjustPositionHV(ByVal vLeft As dgColumns)
        Dim iY As Integer = vLeft.dg.Location.Y
        If Not gbForFiltersGroupBox Is Nothing Then
            If vLeft.blnFilters = False And Me.blnFilters = True Then

                iY = iY + gbForFiltersGroupBox.Height - 2 * gbForFiltersGroupBox.Margin.Top 'statics.GroupBoxRelativeVerticalLocation
            End If
            dg.Location = New System.Drawing.Point(vLeft.dg.Location.X + vLeft.dg.Width + 10, iY)
            Me.gbForFiltersGroupBox.Location = New System.Drawing.Point(dg.Location.X, dg.Location.Y - statics.GroupBoxRelativeVerticalLocation)
        End If
        AdjustdgLabelPosition()
    End Sub

    'Adjust vertical position taking the heights of 2 as base.
    Public Sub AdjustPosition(ByVal vLeft1 As dgColumns, ByVal vLeft2 As dgColumns)
        If (vLeft1.dg.Location.X + vLeft1.dg.Width) > (vLeft2.dg.Location.X + vLeft2.dg.Width) Then
            AdjustPosition(vLeft1)
        Else
            AdjustPosition(vLeft2)
        End If
    End Sub

    'Adjust position of this dg relative to the one above.
    '20190101 With adjustable spacing.
    Public Sub AdjustVerticalPosition(ByVal vAbove As dgColumns, iOffset As Integer)
        Dim p As Point
        If Not vAbove.dgLabel Is Nothing Then
            p = vAbove.dgLabel.Location
            p.Y = p.Y + vAbove.dgLabel.Height + dgLABELOFFSET + iOffset
            p.X = dg.Location.X
        Else
            p = vAbove.dg.Location
            p.Y = p.Y + vAbove.dg.Height + 5
            p.X = dg.Location.X
        End If
        AdjustPosition(p)
        AdjustdgLabelPosition()
    End Sub

    'Adjust position of this dg relative to the one above.
    Public Sub AdjustVerticalPosition(ByVal vAbove As dgColumns)
        AdjustVerticalPosition(vAbove, 5)
    End Sub

    'Adjust vertical position taking the heights of 2 as base.
    '20190101 With adjustable spacing.
    Public Sub AdjustVerticalPosition(ByVal vAbove1 As dgColumns, ByVal vAbove2 As dgColumns, iOffset As Integer)
        If vAbove1.GetHeight() > vAbove2.GetHeight() Then
            AdjustVerticalPosition(vAbove1, iOffset)
        Else
            AdjustVerticalPosition(vAbove2, iOffset)
        End If
    End Sub

    'Adjust vertical position taking the heights of 2 as base.
    Public Sub AdjustVerticalPosition(ByVal vAbove1 As dgColumns, ByVal vAbove2 As dgColumns)
        AdjustVerticalPosition(vAbove1, vAbove2, 5)

    End Sub
    'Public Sub Init(ByVal bs As BindingSource, ByVal iRowHeadersWidth As Integer)
    'End Sub

    'Dim RowIndex As Integer
    Public Overridable Sub RowEnter(ByVal blnAllowUpdate As Boolean, ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        'If aTimer.Interval > 0 Then
        '    aTimer.Stop()
        '    aTimer.Start() 'if already running only resets after a stop
        '    RowIndex = e.RowIndex
        'End If
    End Sub
    'Public Overridable Sub TimerEventProcessor(ByVal myObject As Object, _
    '                                    ByVal myEventArgs As EventArgs) _
    '                                Handles aTimer.Tick
    '    aTimer.Stop()
    'End Sub


    Public Overridable Sub CellFormatting(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs)

    End Sub

    Public Overridable Sub dgResize(ByVal width As Integer, ByVal blnVScrollVisible As Boolean)

        'Idea is to call from containing form with the width of the form. dg sizes a bit less wide.
        Dim iW As Integer
        iW = width - dg.Location.X - 20
        If blnVScrollVisible = True Then
            iW = iW - 25
        End If

        'When form is maximised the filter boxes will not align with the grid if grid is scrolled 
        'unless this is called. Should only be called if data is loaded so switched off the MsgBox in next call.
        If iW < dgInitialWidth Or dgInitialWidth = 0 Then
            Fillout(DataGridViewAutoSizeColumnMode.NotSet)
            If dg.Width <> iW Then
                dg.Width = iW
                AdjustFilterBoxes(dg.HorizontalScrollingOffset)
            End If
        Else
            Fillout(DataGridViewAutoSizeColumnMode.Fill)
            If dg.Width <> dgInitialWidth Then
                dg.Width = dgInitialWidth
                'There is a problem here with scrolling. If a column has been set to Fill on maximisation
                'it still behaves as Fill even though it is NotSet when small again and scrolling.
                Fillout(DataGridViewAutoSizeColumnMode.Fill)
                AdjustFilterBoxes(dg.HorizontalScrollingOffset)
            End If
        End If
    End Sub
    Public Sub SwitchOffCursor()
        '20130212 changed ReadOnlyBackGroundColor to Color.Transparent 

        dg.DefaultCellStyle.SelectionForeColor = dg.DefaultCellStyle.ForeColor
        dg.DefaultCellStyle.SelectionBackColor = statics.GetFieldColor(True) 'read only color
    End Sub

#End Region

#Region "Sizing"
    'If datagrid is limited and fixed in number of rows then this sub will fix the height and disable all scrolling.
    Public Sub SetHeight()
        SetHeightRows(dg.Rows.Count)
    End Sub

    'adjust the height of the grid so that it reaches to the bottom of the form.
    Public Sub SetHeight(ByVal ClientRectangleHeight As Integer)

        Dim iHeight = ClientRectangleHeight - dg.Location.Y - SystemInformation.HorizontalScrollBarHeight
        If ClientRectangleHeight = dg.Parent.ClientRectangle.Height() Then
            Dim frmStandard As frmStandard
            frmStandard = TryCast(dg.Parent, frmStandard)
            If Not frmStandard Is Nothing Then
                frmStandard.HelpTextPosition()
                If frmStandard.HelpTextVertical Then
                    iHeight = iHeight - frmStandard.HelpTextHeight
                End If
            End If
        End If

        'make room for the label.
        If Not dgLabel Is Nothing Then
            dg.Height = iHeight - (dgLabel.Height - dgLABELOFFSET)
        Else
            dg.Height = iHeight
        End If
        AdjustdgLabelPosition()
    End Sub

    '20130202
    Public Sub SetHeightRows(iRows As Integer)
        dg.Height = dg.ColumnHeadersHeight + iRows * dg.RowTemplate.Height() + 2
        'dg.ScrollBars = ScrollBars.None
        AdjustdgLabelPosition()
    End Sub

    'depends on visibility of filterboxes and label.
    Public ReadOnly Property GetHeight() As Integer
        Get

            Dim iHeight = dg.Height
            If Not gbForFiltersGroupBox Is Nothing Then
                If gbForFiltersGroupBox.Visible Then
                    ' iHeight += gbForFiltersGroupBox.Height
                    iHeight = iHeight + (dg.Location.Y - gbForFiltersGroupBox.Location.Y)
                End If
            End If
            If Not dgLabel Is Nothing Then
                'iHeight = iHeight + dgLabel.Height
                iHeight = iHeight + (dgLabel.Location.Y - (dg.Location.Y + dg.Height))
            End If
            Return iHeight
        End Get
    End Property
#End Region

#Region "GridColumns"
    Private Function AddTbToArray(ByVal sequence As Integer, ByVal tb As DataGridViewColumn) As Integer

        'Make sure the index in the array is not already in use by a new column and then add the column 
        'with the unique value.
        Dim iI = 0
        Do While dgSortedListOfColumns.ContainsKey(sequence + iI) = True
            iI = iI + 1
        Loop
        dgSortedListOfColumns.Add(sequence + iI, tb)
        Return sequence + iI
    End Function

    ''' <summary>
    ''' Lookup the name of the first visible column in the datagrid.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetFirstVisibleColumn() As String
        Dim strRet As String = ""
        For Each col As DataGridViewColumn In dg.Columns
            If col.Visible = True Then
                strRet = col.DataPropertyName
                Exit For
            End If
        Next col
        Return strRet
    End Function

    ''' <summary>
    ''' 20100122 RPB created AdjustHorizontalAlignment()
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub AdjustHorizontalAlignment()
        Try
            If Not dg Is Nothing Then
                For Each col As DataGridViewColumn In dg.Columns

                    'ValueType is sometimes undefined on a String type.
                    If Not col.ValueType Is Nothing Then
                        Dim strT As String = col.ValueType.FullName()

                        'Also include boolean because filter box always starts next to previous
                        'column and not the next one and is nice to have filter and field above each other.
                        If strT.StartsWith("System.String") = True Or strT.StartsWith("System.Boolean") = True Or strT.StartsWith("System.DateTime") = True Then
                        Else
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' 20100122 RPB modified PutColumnsInGrid by making call to AdjustHorizontalAlignment()
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub PutColumnsInGrid()

        'Having stored the list of columns with sequence as sorting column get the list of columns and 
        'add to the dg in correct sequence.
        Try
            '20200123 In some circumstances this function may be called when the form is being disposed. So check dgSortedListOfColumns before disposing the columns.
            If Not dgSortedListOfColumns Is Nothing Then
                Dim theColumns As IList(Of DataGridViewColumn) = dgSortedListOfColumns.Values
                Dim blnV As Boolean
                For Each col As DataGridViewColumn In theColumns
                    blnV = col.Visible
                    dg.Columns.Add(col)
                    col.Visible = blnV
                Next col
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        '20100122 RPB created AdjustHorizontalAlignment()
        AdjustHorizontalAlignment()
    End Sub

    Public Sub PutOnlyVisibleColumnsInGrid()

        'Having stored the list of columns with sequence as sorting column get the list of columns and 
        'add to the dg in correct sequence.
        Try
            '20200123 In some circumstances this function may be called when the form is being disposed. So check dgSortedListOfColumns before disposing the columns.
            If Not dgSortedListOfColumns Is Nothing Then
                Dim theColumns As IList(Of DataGridViewColumn) = dgSortedListOfColumns.Values
                Dim blnV As Boolean
                For Each col As DataGridViewColumn In theColumns
                    If col.Visible = True Then
                        blnV = col.Visible
                        dg.Columns.Add(col)
                        col.Visible = blnV
                    End If
                Next col
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        '20100122 RPB created AdjustHorizontalAlignment()
        AdjustHorizontalAlignment()
    End Sub

    Public Sub SetColumnsReadOnly()

        'Having stored the list of columns with sequence as sorting column get the list of columns and 
        'add to the dg in correct sequence.
        Try
            '20200123 In some circumstances this function may be called when the form is being disposed. So check dgSortedListOfColumns before disposing the columns.
            If Not dgSortedListOfColumns Is Nothing Then
                Dim theColumns As IList(Of DataGridViewColumn) = dgSortedListOfColumns.Values
                For Each col As DataGridViewColumn In theColumns
                    col.ReadOnly = True
                    col.DefaultCellStyle.BackColor = statics.GetFieldColor(True) 'ReadOnlyBackGroundColor()
                Next col
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Construct a select statement from the visible fields. Can be used to construct a 'select * from xx where yy' based on filter settings.
    ''' Is useful if there is alot of data and the filter settings are made before the FillTableAdapter() call; see Pricing Tool for an example.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function strGetEfficientSelect()

        Dim strRet As String = "select "
        Try
            '20200123 In some circumstances this function may be called when the form is being disposed. So check dgSortedListOfColumns before disposing the columns.
            If Not dgSortedListOfColumns Is Nothing Then
                Dim theColumns As IList(Of DataGridViewColumn) = dgSortedListOfColumns.Values
                If theColumns.Count > 0 Then
                    Dim blnFirstField As Boolean = True
                    For Each col As DataGridViewColumn In theColumns
                        If col.Visible = True Then
                            If blnFirstField = False Then
                                strRet += ", "
                            Else
                                blnFirstField = False
                            End If
                            strRet += col.Name
                        End If
                    Next col
                    strRet += " from " + strTable
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return strRet
    End Function

    ''' <summary>
    ''' 20120212 Simplified and preferred public version.
    ''' </summary>
    Public Sub DefineColumn(ByVal tb As DataGridViewColumn, _
        ByVal strName As String, _
        ByVal blnRO As Boolean, _
        ByVal iMaxLength As Integer)

        DefineColumn(tb, "", _
            True, strName, "", _
            FieldWidths.GENWIDTH, _
            blnRO, _
            True, _
            "", _
            False)
        If iMaxLength > 0 Then  'non strings have no length.
            Dim tbb As DataGridViewTextBoxColumn
            tbb = TryCast(tb, DataGridViewTextBoxColumn)
            If Not tbb Is Nothing Then
                tbb.MaxInputLength = iMaxLength
            End If
        End If
    End Sub

    'This define colmun includes a max number of characters. Column stops accepting characters when reached.
    Public Sub DefineColumn(ByVal tb As DataGridViewColumn, ByVal strFormat As String, _
        ByVal blnBound As Boolean, ByVal strName As String, ByVal strHeader As String, _
        ByVal iwidth As Integer, _
        ByVal blnRO As Boolean, _
        ByVal blnVisible As Boolean, _
        ByVal strPrintFilter As String, _
        ByVal blnBold As Boolean, ByVal iMaxLength As Integer)

        DefineColumn(tb, strFormat, _
            blnBound, strName, strHeader, _
            iwidth, _
            blnRO, _
            blnVisible, _
            strPrintFilter, _
            blnBold)
        If iMaxLength > 0 Then  'non strings have no length.
            Dim tbb As DataGridViewTextBoxColumn
            tbb = TryCast(tb, DataGridViewTextBoxColumn)
            If Not tbb Is Nothing Then
                tbb.MaxInputLength = iMaxLength
            End If
        End If
    End Sub

    Public Sub DefineColumn(ByVal tb As DataGridViewColumn, ByVal strFormat As String, _
        ByVal blnBound As Boolean, ByVal strName As String, ByVal strHeader As String, _
        ByVal iwidth As Integer, _
        ByVal blnRO As Boolean, _
        ByVal blnVisible As Boolean, _
        ByVal strPrintFilter As String, _
        ByVal blnBold As Boolean)


        Dim sdcColor As System.Drawing.Color = statics.GetFieldColor(blnRO)

        If strForm <> "" And strTable <> "" Then

            'Assumes Textbox if dsBindingSource is not defined (Nothing) otherwise a combobox.
            SynchWithColumns(tb, strFormat, blnBound, strName, strHeader, iwidth, blnRO, blnVisible, strPrintFilter, Nothing, "", "", sdcColor, False)
        Else
            DefineColumn(tb, strFormat, blnBound, strName, strHeader, iwidth, _
                    blnRO, blnVisible, strPrintFilter, sdcColor, blnBold)
        End If
    End Sub

    Public Sub DefineColumn(ByVal tb As DataGridViewColumn, ByVal strFormat As String, _
           ByVal blnBound As Boolean, ByVal strName As String, ByVal strHeader As String, _
           ByVal iwidth As Integer, _
           ByVal blnRO As Boolean, _
           ByVal blnVisible As Boolean, _
           ByVal strPrintFilter As String, _
           ByVal blnBold As Boolean, ByVal sdcColor As System.Drawing.Color, ByVal iMaxLength As Integer)

        If strForm <> "" And strTable <> "" Then

            'Assumes Textbox if dsBindingSource is not defined (Nothing) otherwise a combobox.
            SynchWithColumns(tb, strFormat, blnBound, strName, strHeader, iwidth, blnRO, blnVisible, strPrintFilter, Nothing, "", "", sdcColor, False)
        Else
            DefineColumn(tb, strFormat, blnBound, strName, strHeader, iwidth, _
                    blnRO, blnVisible, strPrintFilter, sdcColor, blnBold)
        End If
        If iMaxLength > 0 Then  'non strings have no length.
            Dim tbb As DataGridViewTextBoxColumn
            tbb = TryCast(tb, DataGridViewTextBoxColumn)
            If Not tbb Is Nothing Then
                tbb.MaxInputLength = iMaxLength
            End If
        End If
    End Sub

    Public Sub DefineColumnNoSynch(ByVal tb As DataGridViewColumn, ByVal strFormat As String, _
    ByVal blnBound As Boolean, ByVal strName As String, ByVal strHeader As String, _
    ByVal iwidth As Integer, _
    ByVal blnRO As Boolean, ByVal blnVisible As Boolean, ByVal strPrintFilter As String, _
    ByVal sdcColor As System.Drawing.Color, ByVal blnBold As Boolean)

        DefineColumn(tb, strFormat, _
                 blnBound, strName, strHeader, _
                 iwidth, _
                 blnRO, blnVisible, strPrintFilter, _
                 sdcColor, blnBold)
        AddTbToArray(1, tb)
    End Sub

    ' 20100122 RPB modified DefineColumn by commenting out rightalignment on format as this is 
    ' done much better with AdjustHorizontalAlignment() when columns are put in the grid.
    Private Sub DefineColumn(ByVal tb As DataGridViewColumn, ByVal strFormat As String, _
        ByVal blnBound As Boolean, ByVal strName As String, ByVal strHeader As String, _
        ByVal iwidth As Integer, _
        ByVal blnRO As Boolean, ByVal blnVisible As Boolean, ByVal strPrintFilter As String, _
        ByVal sdcColor As System.Drawing.Color, ByVal blnBold As Boolean)

        Try
            If blnBound Then tb.DataPropertyName = strName
            Dim tbStyle As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
            If strFormat.Length <> 0 Then
                tbStyle.Format = strFormat
            End If

            'Mod RPB March 2007. Set to bold if necessary.
            '20091209 RPB modified DefineColumn when setting bold font to use dg default style and not the tb.
            If blnBold = True Then
                If tb.HasDefaultCellStyle() = True Then
                    tbStyle.Font = New Font(dg.DefaultCellStyle.Font, FontStyle.Bold)
                Else
                    tbStyle.Font = New Font("Arial", 11, FontStyle.Bold, GraphicsUnit.Pixel)
                End If
            End If
            tbStyle.BackColor = sdcColor

            tb.DefaultCellStyle = tbStyle
            If strHeader.Length <> 0 Then
                tb.HeaderText = strHeader
            Else
                tb.HeaderText = strName
            End If
            tb.HeaderCell.Style.Font = New Font("Arial", 11, FontStyle.Bold, GraphicsUnit.Pixel)

            tb.Name = strName
            tb.Width = iwidth
            tb.ReadOnly = blnRO

            tb.Visible = blnVisible
            tb.SortMode = DataGridViewColumnSortMode.Automatic
            tb.Tag = strPrintFilter

            '20100224
            tb.ToolTipText = strName
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Public Sub SynchWithColumns(ByVal tb As DataGridViewColumn, ByVal strFormat As String, _
            ByVal blnBound As Boolean, ByVal strColumn As String, ByVal strHeader As String, _
            ByVal iwidth As Integer, _
            ByVal blnRO As Boolean, ByVal blnVisible As Boolean, ByVal strPrintFilter As String, _
            ByVal dsBindingSource As BindingSource, _
            ByVal strMember As String, _
            ByVal strDisplayMember As String, _
            ByVal sdcColor As System.Drawing.Color, _
            ByVal blnBold As Boolean _
            )

        'Assume Textbox if dsBindingSource is not defined otherwise a combobox.
        'Lookup the column define data.
        'Use the frmColumns data table which is loaded when program starts and if column data is altered.
        Dim blnPrnt As Boolean = blnVisible

        'update from database if available otherwise store.
        statics.get_v_form_tble_column(strForm, strTable, strColumn, _
                    strHeader, strFormat, iwidth, blnVisible, blnPrnt, blnBold, iSequence)
        Dim strPrint = ""
        If blnPrnt = False Then
            strPrint = MainDefs.DONOTPRINT
        End If

        '20090417 Removed spaces from format string.
        If Not dsBindingSource Is Nothing Then
            DoDefineComboBoxColumn(tb, strFormat, blnBound, strColumn, strHeader, iwidth, blnRO, blnVisible, strPrint, dsBindingSource, strMember, strDisplayMember, sdcColor)
        Else
            DefineColumn(tb, strFormat, blnBound, strColumn, strHeader, iwidth, _
                    blnRO, blnVisible, strPrint, sdcColor, blnBold)
        End If
        AddTbToArray(iSequence, tb)
        iSequence = iSequence + 10
        ' End If
    End Sub

    Private Sub DoDefineComboBoxColumn(ByVal tb As DataGridViewComboBoxColumn, ByVal strFormat As String, _
            ByVal blnBound As Boolean, ByVal strName As String, ByVal strHeader As String, _
            ByVal iwidth As Integer, _
            ByVal blnRO As Boolean, ByVal blnVisible As Boolean, ByVal strPrintFilter As String, _
            ByVal dsBindingSource As BindingSource, ByVal strMember As String, _
            ByVal strDisplayMember As String, ByVal sdcColor As System.Drawing.Color)

        Dim tbStyle As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle

        If blnBound Then tb.DataPropertyName = strName
        If blnRO Then
            tbStyle.BackColor = statics.GetFieldColor(True) 'ReadOnlyBackGroundColor
        Else
            tbStyle.BackColor = sdcColor
        End If

        If strFormat.Length <> 0 Then
            '    Dim tbStyle As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
            tbStyle.Format = strFormat
            tbStyle.BackColor = sdcColor
            If strFormat.StartsWith("N") Then
                tbStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            End If
            tb.DefaultCellStyle = tbStyle
        Else
            tb.DefaultCellStyle.BackColor = sdcColor
        End If

        If strHeader.Length <> 0 Then
            tb.HeaderText = strHeader
        Else
            tb.HeaderText = strName
        End If

        tb.HeaderCell.Style.Font = New Font("Arial", 11, FontStyle.Bold, GraphicsUnit.Pixel)
        tb.HeaderCell.Style.BackColor = Color.AliceBlue
        tb.Name = strName
        tb.Width = iwidth
        tb.ReadOnly = blnRO
        tb.Visible = blnVisible
        tb.DataSource = dsBindingSource
        tb.DisplayMember = strDisplayMember
        tb.ValueMember = strMember
        tb.Tag = strPrintFilter
        tb.SortMode = DataGridViewColumnSortMode.Automatic

        '20110426 This solves problem of not being able to set the backcolor on combo in windows 7.
        tb.FlatStyle = FlatStyle.Flat
        '20100224
        tb.ToolTipText = strName
    End Sub

    Public Sub DefineComboBoxColumn(ByVal tb As DataGridViewComboBoxColumn, ByVal strFormat As String, _
        ByVal blnBound As Boolean, ByVal strName As String, ByVal strHeader As String, _
        ByVal iwidth As Integer, _
        ByVal blnRO As Boolean, ByVal blnVisible As Boolean, ByVal strPrintFilter As String, _
        ByVal dsBindingSource As BindingSource, ByVal strMember As String, _
        ByVal strDisplayMember As String, ByVal sdcColor As System.Drawing.Color)

        If strForm <> "" And strTable <> "" Then

            'Assumes Textbox if dsBindingSource is not defined (Nothing) otherwise a combobox.
            SynchWithColumns(tb, strFormat, blnBound, strName, strHeader, iwidth, blnRO, blnVisible, strPrintFilter, dsBindingSource, strMember, strDisplayMember, sdcColor, False)
        Else
            DoDefineComboBoxColumn(tb, strFormat, blnBound, strName, strHeader, iwidth, blnRO, blnVisible, strPrintFilter, dsBindingSource, strMember, strDisplayMember, sdcColor)
        End If
    End Sub

    Public Sub DefineComboBoxColumn(ByVal tb As DataGridViewComboBoxColumn, ByVal strFormat As String, _
        ByVal blnBound As Boolean, ByVal strName As String, ByVal strHeader As String, _
        ByVal iwidth As Integer, _
        ByVal blnRO As Boolean, ByVal blnVisible As Boolean, ByVal strPrintFilter As String, _
        ByVal dsBindingSource As BindingSource, ByVal strMember As String, ByVal strDisplayMember As String)

        '20090407
        Dim sdcColor As System.Drawing.Color = statics.GetFieldColor(blnRO)
        'If blnRO = True Then
        '    sdcColor = ReadOnlyBackGroundColor
        'End If
        'sdcColor = statics.GetFieldColor(blnRO)

        DefineComboBoxColumn(tb, strFormat, blnBound, strName, strHeader, iwidth, _
                blnRO, blnVisible, strPrintFilter, dsBindingSource, strMember, strDisplayMember, sdcColor)
    End Sub
#End Region

#Region "Filter"
    Public Sub BroadcastFilter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim frm As MainForm

        'Can fail if a form is being used as a dialog. See for example frmSAPCosts.
        Try
            frm = TryCast(Me.dg.Parent.Parent.Parent, MainForm)
            If Not frm Is Nothing Then
                frm.BroadcastFilter(sender, e)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Protected Overridable Function GetExcelFileName() As String
        Return Me.strTable
    End Function
    Protected Overridable Sub dg_CellDoubleClick(ByVal sender As System.Object, _
        ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dg.CellDoubleClick
        If e.ColumnIndex = -1 Then
            If e.RowIndex = -1 Then
                dg.ClearSelection()

                'upper left -> dump to Excel
                Dim frm As frmStandard
                frm = TryCast(Me.frmParent, frmStandard)
                If Not frm Is Nothing Then
                    Dim strFilename = GetExcelFileName()
                    frm.PrintExcel(strFilename, dg)
                End If
            End If
        Else
            If e.RowIndex >= 0 And e.ColumnIndex >= 0 Then

                'Do this after the above otherwise the datagrid may already be filtered before the ColumnIndex and RowIndex are
                'used.
                'RPB Aug 2008. Decided to use the MainDefs global value here to simplify online switching.
                If MainDefs.blnActiveFilters = True Then
                    BroadcastFilter(sender, e)
                    ColumnDoubleClick(sender, e, False)
                End If
            End If
        End If
    End Sub

    Public Sub FilterFromOtherForm(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        ColumnDoubleClick(sender, e)
    End Sub
    Public Function CheckTag(ByVal iColumn As DataGridViewColumn, ByVal strTagFilter As String) As Boolean

        'Return False if the character in the Tag of the Column is contained in the TagFilter string.
        'Otherwise return True.
        If IsNothing(iColumn.Tag) Then Return True
        If iColumn.Tag.ToString.Length = 0 Then
            Return True
        End If
        If strTagFilter.Contains(iColumn.Tag) Then Return False
        Return True
    End Function


    '20090907 RPB added bTabStop to CreateAFilterBox to allow user to set TabStop to false
    'to prevent Tab moving to this control.
    Protected Sub CreateAFilterBox(ByRef tb As TextBox, ByVal strField As String, _
    ByRef tb_TextChanged As EventHandler, ByVal Controls As Control.ControlCollection, _
    ByVal bTabStop As Boolean)

        tb = New System.Windows.Forms.TextBox
        tb.Name = strField
        AddHandler tb.TextChanged, tb_TextChanged
        FindTbs.Add(tb, strField)
        Me.gbForFilters.Controls.Add(tb)
        tb.TabStop = bTabStop
    End Sub

    Protected Sub CreateAFilterBox(ByRef tb As TextBox, ByVal strField As String, _
    ByRef tb_TextChanged As EventHandler, ByVal Controls As Control.ControlCollection)
        CreateAFilterBox(tb, strField, tb_TextChanged, Controls, True)
    End Sub

    '20090907 RPB added bTabStop to CreateACheckBox to allow user to set TabStop to false
    'to prevent Tab moving to this control.
    Public Sub CreateACheckBox(ByRef tb As CheckBox, ByVal strField As String, _
        ByRef tb_TextChanged As EventHandler, _
        ByVal Controls As Control.ControlCollection, _
        ByVal bTabStop As Boolean)

        tb = New System.Windows.Forms.CheckBox
        tb.ThreeState = True
        tb.CheckState = CheckState.Indeterminate
        tb.Name = "cb" & strField & "Find"

        'AddHandler tb.CheckStateChanged, tb_TextChanged
        AddHandler tb.CheckStateChanged, tb_TextChanged
        FindTbs.Add(tb, strField)
        Me.gbForFilters.Controls.Add(tb)
        tb.TabStop = bTabStop
    End Sub

    Public Sub CreateACheckBox(ByRef tb As CheckBox, ByVal strField As String, _
        ByRef tb_TextChanged As EventHandler, _
        ByVal Controls As Control.ControlCollection)
        CreateACheckBox(tb, strField, tb_TextChanged, Controls, True)
    End Sub

    '20110930 RPB added iTab to return the column position.
    Public Function GetLeftOfColumnInGrid(ByVal col As String, ByRef iColPosition As Integer) As Integer

        'Return the sum of the widths of the datagridview columns up to but not including the col parameter.
        Dim i As Integer
        Dim w As Integer
        Dim iRet As Integer = -1
        iColPosition = -1
        'RPB Feb 2007. Started adjusting the RowHeader column so needed to use the actual value here.
        If dg.RowHeadersVisible = True Then
            w = dg.RowHeadersWidth
        Else
            w = 0
        End If

        i = 0
        While i < dg.Columns.Count
            If dg.Columns(i).Name = col Then
                iRet = w
                Exit While
            End If
            If dg.Columns(i).Visible = True Then
                w = w + dg.Columns(i).Width
            End If
            i = i + 1
        End While
        iColPosition = i
        Return iRet
    End Function

    ' 20100122 RPB created AdjustFilterTextBox. 
    'Set Some filter boxes RO: Otherwise just allow user to do double click.
    'Looks obvious to call from CreateFilterBox but does not work there because DefineColumns gets
    'called afterwards and is needed to get the Datagrid column types.
    'Also set filter textbox alignment.
    Private Sub AdjustFilterTextBox(ByVal cTb As Control, ByVal strField As String)
        Dim tb As TextBox = TryCast(cTb, TextBox)
        If Not tb Is Nothing Then
            Dim strDataPropertyType As String = GetBoundColumnType(strField)

            If strDataPropertyType.StartsWith("System.DateTime") Then
                'tb.ReadOnly = True
                '20181231 removed readonly True so that user can remove after double clicking.
                tb.ReadOnly = False
            Else
                '20100517 Added this explicit Readonly to false because this was true if the datagridview was read only.
                tb.ReadOnly = False
            End If

            If strDataPropertyType.StartsWith("System.String") Then
                tb.TextAlign = HorizontalAlignment.Left
            Else
                tb.TextAlign = HorizontalAlignment.Right
            End If
        End If
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="iScrollposition"></param>
    ''' <returns>The width required for all the filters.</returns>
    ''' <remarks></remarks>
    Private Function AdjustFilterBoxesInGB(ByVal iScrollposition As Integer) As Integer
        Dim iTotalWidth = 0
        Try

            'Place the filter boxes above the columns in the datagridview.
            Dim tab As Integer = 1
            Dim iLeft As Integer
            For Each tbEntry As KeyValuePair(Of Control, String) In FindTbs

                'Get the left hand position for the filter. Is -1 if filter name <> a grid column name
                iLeft = GetLeftOfColumnInGrid(tbEntry.Value, tab)
                If iLeft <> -1 Then
                    tbEntry.Key.Visible = dg.Columns(tbEntry.Value).Visible
                    tbEntry.Key.Size = New System.Drawing.Size(dg.Columns(tbEntry.Value).Width, TEXTBOX_HEIGHT_IN_FILTER_GROUPBOX)
                    tbEntry.Key.Location = New System.Drawing.Point(iLeft, TEXTBOX_Y_IN_FILTER_GROUPBOX)
                    tbEntry.Key.TabIndex = tab

                    '20100122
                    AdjustFilterTextBox(tbEntry.Key, tbEntry.Value)
                Else
                    '20100521 If the filter box cant be coupled to a field in the dg then dont show.
                    tbEntry.Key.Visible = False
                End If
                If (iLeft + tbEntry.Key.Size.Width) > iTotalWidth Then
                    iTotalWidth = iLeft + tbEntry.Key.Size.Width
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Return iTotalWidth
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Dim iScrollPosition As Integer = 0
    Private Sub dg_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles dg.Scroll
        If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then

            '20211014 ADHOC view can cause problems when scrolling if gbForFilters has not been set when another view is selected.
            If Not gbForFilters Is Nothing Then
                'Adjust filter visibility and position.
                iScrollPosition = e.NewValue
                Dim p As Point = gbForFilters.Location
                p.X = -e.NewValue
                gbForFilters.Location = p
            End If


        End If
    End Sub

    Private Sub dg_ColumnWidthChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs) Handles dg.ColumnWidthChanged
        Try
            'Adjust filter visibility and position.
            Me.AdjustFilterBoxes(iScrollPosition)
        Catch ex As Exception
        End Try
    End Sub

    Public Sub AdjustColumnWidth()
        AdjustFilterBoxes(iScrollPosition)
    End Sub

    Public Sub AdjustFilterBoxes()
        AdjustFilterBoxes(0)

        '20100529 Set the default_filter value.
        'Do this here so it is only done during loading of the form and not when adjusting column position etc
        'when AdjustFilterBoxes(ByVal iScrollposition As Integer) is called.
        For Each tbEntry As KeyValuePair(Of Control, String) In FindTbs
            Dim strFilter As String = statics.strGetDefault_Filter(Me.strForm, Me.strTable, tbEntry.Value)
            If strFilter.Length > 0 Then
                Dim ch As CheckBox = TryCast(tbEntry.Key, CheckBox)
                If ch Is Nothing Then
                    ColumnDoubleClick(tbEntry.Value, strFilter)
                Else
                    If strFilter = "1" Then
                        ch.CheckState = CheckState.Checked
                        MakeFilter(True)
                    Else
                        If strFilter = "0" Then
                            ch.CheckState = CheckState.Unchecked
                            MakeFilter(True)
                        End If
                    End If
                End If
            End If
        Next
    End Sub

    Public Sub AdjustFilterBoxes(ByVal iHorizScrollposition As Integer)

        If Not gbForFiltersGroupBox Is Nothing And Not gbForFilters Is Nothing Then

            '20130309 The width of gbForFilters was set to the width of the datagridview and that was not correct.
            gbForFilters.Width = AdjustFilterBoxesInGB(iHorizScrollposition)
            Dim p As Point = gbForFilters.Location

            '20120708 RPB solves problem with bad positioning of the GroupBox for filters after, decreasing size of form
            'to less than the datagrid width then scrolling right in the grid and then increasinf size of form.
            p.X = p.X + iScrollPosition - iHorizScrollposition
            iScrollPosition = iHorizScrollposition
            gbForFilters.Location = p
            gbForFiltersGroupBox.Width = dg.Width
        End If
    End Sub

    ''' <summary>
    ''' Call to give the first visible filter textbox the focus. The user can immediately start typing.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub SelectFirstTb()

        'Make sure that other controls do not have the same tab values!
        'Give the first control the focus.
        For Each tbEntry As KeyValuePair(Of Control, String) In FindTbs
            If tbEntry.Key.Visible = True Then
                tbEntry.Key.Select()
                Exit For
            End If
        Next
    End Sub

    Public Sub SelectFirstTbUserOrder()

        'Make sure that other controls do not have the same tab values!
        'Give the first control the focus.
        Dim c As Control = Nothing
        Dim ti As Integer = Integer.MaxValue
        For Each tbEntry As KeyValuePair(Of Control, String) In FindTbs
            If tbEntry.Key.TabIndex < ti And tbEntry.Key.Visible = True Then
                ti = tbEntry.Key.TabIndex
                c = tbEntry.Key
            End If
        Next
        If Not c Is Nothing Then
            c.Select()
        End If
    End Sub

    ''' <summary>
    ''' 20100105 RPB added for via
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub HideFilter()
        '
        For Each tbEntry As KeyValuePair(Of Control, String) In FindTbs
            tbEntry.Key.Visible = False
        Next
    End Sub

    Public Sub ResetFilter()
        'Remove all strings from the filter textboxes. Is used when re-showing a form as a dialog.
        For Each tbEntry As KeyValuePair(Of Control, String) In FindTbs
            If tbEntry.Key.ToString.Contains("CheckBox") = True Then
                Dim tb As CheckBox
                tb = CType(tbEntry.Key, CheckBox)
                tb.CheckState = CheckState.Indeterminate
            Else
                tbEntry.Key.Text = ""
            End If
        Next

        '20150625 And then set to the preset values.
        AdjustFilterBoxes()
        'Is not necessary because textbox event handlers call MakeFilter()
        'bs.RemoveFilter()
    End Sub

    Public Function GetBoundColumnName(ByVal dgColumnName As String) As String

        'Return the name of the bound column.
        Dim strRet As String = ""
        Try
            strRet = dg.Columns(dgColumnName).DataPropertyName
        Catch ex As Exception
        End Try
        Return strRet
    End Function

    '20180120 GetBoundColumnType: When called from Adhoc table the ValueType may be nothing.
    Public Function GetBoundColumnType(ByVal dgColumnName As String) As String

        'Return the type of the bound column.
        'For example System.String or System.Int32.
        Dim strRet As String = ""
        Try
            Dim iType As System.Type
            iType = dg.Columns(dgColumnName).ValueType
            If Not iType Is Nothing Then
                strRet = iType.ToString()
            End If
        Catch ex As Exception
        End Try
        Return strRet
    End Function

    ''' <summary>
    ''' 20091124 RPB created blnStringHasWildCards.
    ''' 20100122 RPB Modified blnStringHasWildCards: Do not check for *.
    ''' Return true if the string contains a 'Like' wildcard.
    ''' </summary>
    ''' <param name="strText"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function blnStringHasWildCards(ByVal strText As String) As Boolean
        If strText.Contains("%") Then Return True
        If strText.Contains("[^]") Then Return True
        '20100122 Do not check for *
        '        If strText.Contains("*") Then Return True
        If strText.Contains("[") And strText.Contains("]") Then Return True
        Return False
    End Function

    ''' <summary>
    ''' 20091124 RPB Modified MakeFilter: Do not use Like if string contains a wildcard character.
    ''' 20091209 RPB Modified MakeFilter: Allow filter on DateTime by converting to string.
    ''' 20100122 RPB Modified MakeFilter: Only add * if there is no * in the string already.
    ''' </summary>
    ''' <param name="blnExact"></param>
    ''' <remarks></remarks>
    ''' 
    Public Sub MakeFilter(ByVal blnExact As Boolean)
        Dim strF As String = strMakeFilter(blnExact)
        Try
            strF = strF.Trim
            If strF.Trim.Length = 0 Then
                bs.RemoveFilter()
            Else
                bs.Filter = strF
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly)
        End Try

    End Sub


    ''' <summary>
    ''' 20111004 Return the filter as a string.
    ''' 20140807 Made function overriable. Used in GIS to set a flag if filters change.
    ''' </summary>
    ''' <param name="blnExact"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function strMakeFilter(ByVal blnExact As Boolean) As String

        ' Iterate through the dictionary of find list boxes to construct the filter and then set the filter.
        Dim strF As String
        strF = ""
        For Each tbEntry As KeyValuePair(Of Control, String) In FindTbs
            If tbEntry.Key.ToString.Contains("TextBox") = True Then
                If tbEntry.Key.Text.Trim.Length <> 0 Then

                    'Get the bound column name of the datagrid.
                    Dim strDataPropertyName = "[" & GetBoundColumnName(tbEntry.Value) & "]"
                    Dim strDataPropertyType As String = GetBoundColumnType(tbEntry.Value)

                    If strDataPropertyName <> "[]" Then
                        If strF.Length <> 0 Then
                            strF = strF & " and "
                        End If
                        If strDataPropertyType.StartsWith("System.String") Then
                            If blnExact = False Then

                                '20091124 RPB modified MakeFilter. If the string contains a Like wildcard then do not use Like.
                                If blnStringHasWildCards(tbEntry.Key.Text) = True Then
                                    strF = strF & " " & strDataPropertyName & " = '" & tbEntry.Key.Text.Replace("'", "''") & "' "
                                Else
                                    '20100122 Only add * if there is no * in the string already.
                                    If tbEntry.Key.Text.Contains("*") Then
                                        strF = strF & " " & strDataPropertyName & " Like '" & tbEntry.Key.Text.Replace("'", "''") & "' "
                                    Else

                                        strF = strF & " " & strDataPropertyName & " Like '" & tbEntry.Key.Text.Replace("'", "''") & "*' "
                                    End If
                                End If
                            Else
                                strF = strF & " " & strDataPropertyName & " = '" & tbEntry.Key.Text.Replace("'", "''") & "' "
                            End If
                        Else
                            If strDataPropertyType.StartsWith("System.DateTime") Then

                                '20100114 RPB modified MakeFilter(): To be independent of the regional settings the filter value has to be in the
                                'MS SQL format  mm/dd/yyyy. The previous version worked if English US was used
                                'but failed in for example French(French) regional settings.
                                'However depends on Text entry being a complete date so put try around and make 
                                'datetime and number columns RO so that they only work on double click.
                                Try
                                    Dim dd As Date = tbEntry.Key.Text
                                    tbEntry.Key.Text = dd.Date.ToString("d")
                                    dd = dd.AddDays(1)
                                    strF = strF & " " & strDataPropertyName & " < '" & dd.Date.ToString("d") & "' "
                                    dd = dd.AddDays(-2)
                                    strF = strF & " and " & strDataPropertyName & " > '" & dd.Date.ToString("d") & "' "
                                Catch ex As Exception
                                End Try
                            Else
                                strF = strF & " " & strDataPropertyName & " = " & tbEntry.Key.Text.Replace("'", "''")
                            End If

                        End If
                    End If
                    'strF = strF & " " & tbEntry.Value & " Like '" & tbEntry.Key.Text.Replace("'", "''") & "' "
                End If
            Else
                If tbEntry.Key.ToString.Contains("CheckBox") = True Then
                    Dim tb As CheckBox
                    tb = CType(tbEntry.Key, CheckBox)
                    Dim strDataPropertyName = "[" & GetBoundColumnName(tbEntry.Value) & "]"
                    If strDataPropertyName <> "[]" Then
                        If tb.CheckState = CheckState.Checked Then
                            If strF.Length <> 0 Then
                                strF = strF & " and "
                            End If
                            strF = strF & " " & strDataPropertyName & " = 1 "
                        End If
                        If tb.CheckState = CheckState.Unchecked Then
                            If strF.Length <> 0 Then
                                strF = strF & " and "
                            End If
                            strF = strF & " " & strDataPropertyName & " = 0 "
                        End If
                        If tb.CheckState = CheckState.Indeterminate Then
                        End If
                    End If
                End If
            End If
        Next
        Return strF
    End Function

    Private Function blnIsDatagridCombo(ByVal tbKey As Control) As Boolean

        '20100910 RPB simplified.
        Dim dgc As DataGridViewComboBoxColumn = TryCast(dg.Columns(tbKey.Name), DataGridViewComboBoxColumn)
        If Not dgc Is Nothing Then
            Return True
            'If dgc.CellType.Name.Contains("ComboBox") Then Return True
        End If
        Return False
    End Function

    '20100910 Added function to allow the filter on combo box to be supressed.
    Public Function ColumnDoubleClick( _
        ByVal sender As System.Object, _
        ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, ByVal blnFilterComboBox As Boolean) As Boolean

        Dim blnRet = False
        If e.RowIndex >= 0 And e.ColumnIndex >= 0 Then
            Dim dg As DataGridView
            dg = CType(sender, DataGridView)
            Dim strColumnName As String = dg.Columns(e.ColumnIndex).Name
            Dim strValue As String
            Try
                strValue = dg.Rows(e.RowIndex).Cells(e.ColumnIndex).Value()
                ColumnDoubleClick(strColumnName, strValue, blnFilterComboBox)
            Catch ex As Exception

            End Try
        End If
        Return blnRet
    End Function

    ''' <summary>
    ''' 20091005 Added true to ColumnDoubleClick(strColumnName, strValue, True) to allow 
    ''' filtering on comboboxes from other forms.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ColumnDoubleClick( _
        ByVal sender As System.Object, _
        ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) As Boolean

        Return ColumnDoubleClick(sender, e, True)
    End Function

    Public Sub ColumnDoubleClick(ByVal Fields As Dictionary(Of String, String))
        For Each kvp As KeyValuePair(Of String, String) In Fields
            ColumnDoubleClick(kvp.Key, kvp.Value)
        Next
    End Sub

    ''' <summary>
    ''' 20091005 Override so that Filtering from other form does allow ColumnDoubleClick
    ''' </summary>
    ''' <param name="strColumnName"></param>
    ''' <param name="strFilterText"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ColumnDoubleClick( _
        ByVal strColumnName As String, _
        ByVal strFilterText As String) As Boolean
        Return ColumnDoubleClick(strColumnName, strFilterText, False)
    End Function

    Public Function ColumnDoubleClickExact( _
    ByVal strColumnName As String, _
    ByVal strFilterText As String) As Boolean
        Return ColumnDoubleClick(strColumnName, strFilterText, False, True)
    End Function
    Public Function ColumnDoubleClick( _
       ByVal strColumnName As String, _
       ByVal strFilterText As String, ByVal blnAllowCombo As Boolean) As Boolean
        ColumnDoubleClick(strColumnName, strFilterText, blnAllowCombo, False)
    End Function

    Public Function ColumnDoubleClick( _
        ByVal strColumnName As String, _
        ByVal strFilterText As String, ByVal blnAllowCombo As Boolean, blnExact As Boolean) As Boolean

        'RPB Feb 2008. Lookup the column and place the text in it.
        Dim blnRet = False

        For Each tbEntry As KeyValuePair(Of Control, String) In FindTbs
            Dim strLB As String
            strLB = tbEntry.Key.Name
            If strLB.ToLower() = strColumnName.ToLower() Then

                '20090310 RPB Do not allow dbl click on comboboxes
                If blnAllowCombo = True Or blnIsDatagridCombo(tbEntry.Key) = False Then
                    tbEntry.Key.Text = strFilterText
                    MakeFilter(blnExact)
                    blnRet = True
                    Exit For
                End If
            End If
        Next
    End Function
#End Region

#Region "Editing"
    Protected Overridable Sub dg_CellValidated(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dg.CellValidated
    End Sub

    Private _blnRowAdded As Boolean
    Public Property blnRowAdded() As Boolean
        Get
            Return _blnRowAdded
        End Get
        Set(ByVal value As Boolean)
            _blnRowAdded = value
        End Set
    End Property

    Public Overridable Sub dg_RowValidated(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dg.RowValidated
        Try
            dg.Rows(e.RowIndex).ErrorText = Nothing
        Catch ex As Exception
        End Try

        'If there is a SortFieldName (preferably a field with unique value) then store the current value.
        'This is used to restore the row position after the sort.
        If Not strSortFieldName Is Nothing Then
            If strSortFieldName.Length > 0 Then
                Try
                    If Not dg.Rows(e.RowIndex).Cells(strSortFieldName) Is System.DBNull.Value Then

                        'Store the prj at the current row so that it can be found again after a sort.
                        If Not dg.Rows(e.RowIndex).Cells(strSortFieldName).Value Is System.DBNull.Value Then
                            strPrj = dg.Rows(e.RowIndex).Cells(strSortFieldName).Value
                        End If
                    End If
                Catch ex As Exception

                End Try
            End If
        End If
        If Not strSortFieldName2 Is Nothing Then
            If strSortFieldName2.Length > 0 Then
                Try
                    If Not dg.Rows(e.RowIndex).Cells(strSortFieldName2) Is System.DBNull.Value Then

                        'Store the prj at the current row so that it can be found again after a sort.
                        If Not dg.Rows(e.RowIndex).Cells(strSortFieldName2).Value Is System.DBNull.Value Then
                            strPrj2 = dg.Rows(e.RowIndex).Cells(strSortFieldName2).Value
                        End If
                    End If
                Catch ex As Exception

                End Try
            End If
        End If

    End Sub
    Private Sub dg_Sorted(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dg.Sorted

        'If there is a SortFieldName and the calue is defined restore the row.
        Try
            If strPrj.Length > 0 And strSortFieldName.Length > 0 And strSortFieldName2.Length = 0 Then
                Dim cCell As DataGridViewCell = Nothing
                For i As Integer = 0 To (dg.Rows.Count - 1)
                    If dg.Rows(i).Cells(strSortFieldName).Value.ToString() = strPrj Then
                        cCell = dg.Rows(i).Cells(dg.SortedColumn.Index)
                        If Not cCell Is Nothing And cCell.Visible = True Then
                            dg.CurrentCell = cCell
                            dg.CurrentRow.Selected = True
                        End If
                        Exit For
                    End If
                Next
            End If
            If strPrj.Length > 0 And strPrj2.Length > 0 And strSortFieldName.Length > 0 And strSortFieldName2.Length > 0 Then
                Dim cCell As DataGridViewCell = Nothing
                For i As Integer = 0 To (dg.Rows.Count - 1)
                    If dg.Rows(i).Cells(strSortFieldName).Value.ToString() = strPrj And dg.Rows(i).Cells(strSortFieldName2).Value.ToString() = strPrj2 Then
                        cCell = dg.Rows(i).Cells(dg.SortedColumn.Index)
                        If Not cCell Is Nothing And cCell.Visible = True Then
                            dg.CurrentCell = cCell
                            dg.CurrentRow.Selected = True
                        End If
                        Exit For
                    End If
                Next
            End If

        Catch ex As Exception
        End Try
    End Sub
    Private Sub dg_UserAddedRow(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dg.UserAddedRow
        'Is needed to prevent problems when deleting the last row in a child datagrid.
        blnRowAdded = True
    End Sub
    Public Overridable Sub dg_UserDeletingRow(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dg.UserDeletingRow
        'THIS IS TRICKY. IF Overrides Sub has Handles dg.UserDeletingRow 
        'APPENDED IT IS CALLED TWICE AND WILL DELETE 2 RECORDS!!! 
        blnRowAdded = False
    End Sub
    Private Sub dg_DataError(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dg.DataError

        'THIS IS NEEDED TO PREVENT ERROR MESSAGES IN EMPTY ROWS.
        'This is the standard error handler for data grid columns. Is triggered, for example, 
        'if an non number is entered in an N2 column. 
        'Have not discovered how to define the formatting more accurately. For example Brix 
        'should not only be N2 but should not be negative. This extra checking has been implemented in the 
        'CellValidating handler below. 

        'CHECK: Is needed here to show the user the error text for non valid fields after editing has started.
        'It could be better to let the database throw the exception directly but see comment on 
        'EndEdit in RowValidating.

        'Dim dg As DataGridView = CType(sender, DataGridView)
        'Dim rv As DataGridViewCell
        'rv = dg.CurrentCell
        'e.ThrowException = False
        ''Debug.Print("data error")
        'e.Cancel = True
        'dg.CurrentRow.ErrorText = "Data not saved." & vbCrLf & e.Exception.Message
    End Sub

    '20100220 Validate returns true if insert took place.
    Public Function Validate(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) As Boolean
        Dim blnRet As Boolean = False
        Try
            If blnValidating = False Then
                blnValidating = True
                blnUpdated = False
                blnInserted = False
                blnRet = Handle_DataGridView_RowValidating("dg_RowValidating", Me.__ta, sender, e, 0)
                blnValidating = False
            End If
        Catch ex As Exception
            e.Cancel = True
        End Try
        Return blnRet
    End Function

    '20100225 Provide flags to say what happended.
    Protected blnUpdated As Boolean = False
    Protected blnInserted As Boolean = False

    '20090319 RPB needed to prevent recursive call of this handler when processing a stored procedure in the update
    'command. Not sure why.
    Dim blnValidating = False
    Protected Overridable Sub dg_RowValidating(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dg.RowValidating
        Dim blnRet As Boolean

        Try
            'Find index of first visible column
            Dim id As Integer = -1
            For Each c As DataGridViewColumn In dg.Columns
                If c.Visible = True Then
                    id = c.Index
                    Exit For
                End If
            Next
            If id <> -1 And blnValidating = False Then
                blnValidating = True
                blnUpdated = False
                blnInserted = False
                blnRet = Handle_DataGridView_RowValidating("dg_RowValidating", Me.__ta, sender, e, id)
                blnValidating = False
                If blnRet = False Then

                End If
            End If
        Catch ex As Exception
            e.Cancel = True
            blnValidating = False
        Finally

        End Try

    End Sub
    Public Overridable Function Handle_DataGridView_RowValidating(ByVal EventName As String, ByRef TableAdapter As System.Object, ByVal sender As System.Object, _
    ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs, ByVal iCheckColumn As Integer) As Boolean
        'Return true if Update.

        Dim blnRet As Boolean
        '20100219 Changed blnRet = True to blnRet = false in Handle_DataGridView_RowValidating
        blnRet = False
        Dim dg As DataGridView = CType(sender, DataGridView)

        Dim dgv As dgvEnter = TryCast(sender, dgvEnter)
        If Not dgv Is Nothing Then
            dgv.blnDirty = False
        End If
        Try

            Dim bs As BindingSource = CType(dg.DataSource, BindingSource)
            If Not bs Is Nothing Then
                Dim aRow As System.Data.DataRow
                Dim sD As System.Data.DataRowView
                sD = bs.Current

                If Not sD Is Nothing Then
                    aRow = sD.Row
                    Select Case aRow.RowState
                        Case DataRowState.Added
                        Case DataRowState.Deleted
                        Case DataRowState.Detached

                            '20100105 RPB Modified Handle_DataGridView_RowValidating 
                            'Dropped this check on null column for RAP.Additives.
                            'But check this in other apps.
                            'If Not aRow.IsNull(iCheckColumn) Then
                            Try 'this can fail is another column in the row may not be null but is.
                                'just ignore.
                                bs.EndEdit()
                            Catch ex As Exception

                                'Only show error message is the error is occurring after a row has been added.
                                If blnRowAdded = True Then
                                    MsgBox(ex.Message)
                                    e.Cancel = True
                                End If
                            End Try
                            'End If
                        Case DataRowState.Modified
                        Case DataRowState.Unchanged
                            bs.EndEdit()
                    End Select
                    'Debug.Print("--->" & aRow.RowState.ToString())
                    Select Case aRow.RowState
                        Case DataRowState.Added
                            CallByName(TableAdapter, "Update", CallType.Method, aRow)
                            blnInserted = True
                            blnRet = True
                        Case DataRowState.Deleted
                        Case DataRowState.Detached
                        Case DataRowState.Modified
                            'Try
                            CallByName(TableAdapter, "Update", CallType.Method, aRow)
                            blnUpdated = True
                            '20100220 Only return true if insert took place.
                            'blnRet = True
                            '                Catch ex As Exception
                            '    'No e.Cancel = true so the user can move off the row.
                            '    MsgBox("EXCEPTION: " & ex.Message & " tt")
                            'End Try
                        Case DataRowState.Unchanged
                            '                    Case DataRowState.Unchanged
                    End Select
                End If
            End If
        Catch ex As Exception

            '20091124 RPB added error handling to Handle_DataGridView_RowValidating
            'for cases where data gets truncated.
            Dim exx As System.Data.SqlClient.SqlException
            exx = TryCast(ex, System.Data.SqlClient.SqlException)
            If Not exx Is Nothing Then
                If exx.ErrorCode = -2146232060 Then     'data would be truncated statement ended.
                    Dim errorMessages As New StringBuilder()
                    Dim i As Integer
                    For i = 0 To exx.Errors.Count - 1
                        errorMessages.Append("Index #" & i.ToString() & ControlChars.NewLine _
                            & "Message: " & exx.Errors(i).Message & ControlChars.NewLine _
                            & "LineNumber: " & exx.Errors(i).LineNumber & ControlChars.NewLine _
                            & "Source: " & exx.Errors(i).Source & ControlChars.NewLine _
                            & "Procedure: " & exx.Errors(i).Procedure & ControlChars.NewLine)
                    Next i

                    MsgBox("EXCEPTION: " & ex.Message)

                    'Bit too verbose!
                    'MsgBox(errorMessages)
                End If
            End If

            'RPB Jan 2008 Gives a problem when using proposed data on a ComboBox input field.
            'If ex.Message.IndexOf("no Proposed data") = 0 Then
            'RPB Aug 2009. Suppress the error text and just delte from the datagrid.
            'MsgBox("EXCEPTION: " & ex.Message & " DELETE the row you are entering to start again.")
            e.Cancel = True
            blnRet = False
            bs.Current.Delete()

            '20100222 Add this to force a refresh after deleting the row.
            'This is because Current.Delete deltes fro the form which the underlying table may still contain the record
            'for example after an Update.
            'However will only work with the new New above.
            If Not frmParent Is Nothing Then
                frmParent.RefreshTheForm()
            End If

            'bs.CancelEdit()    does not work after the EndEdit.
        Finally
            'Debug.Print("--->" & dr.RowState.ToString())
        End Try
        Return blnRet
    End Function

#End Region

#Region "Focus"
    'Idea is to highlight which datagrid has the focus for keybord users.
    Private Sub dg_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dg.Leave
        dg.ForeColor = Color_Unfocussed
    End Sub

    Private Sub dg_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dg.Enter
        dg.ForeColor = Color.Black
    End Sub

#End Region
End Class
